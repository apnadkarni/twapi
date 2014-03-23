/*
 * Copyright (c) 2014, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

/*
 * TwapiCStruct is a Tcl "type" that holds definition of a C structure.
 * string table. 
 * The Tcl_Obj.internalRep.twoPtrValue.ptr1 points to a TwapiCStructRep
 * which holds the decoded information.
 * and Tcl_Obj.internalRep.twoPtrValue.ptr2 not used.
 */
static void DupCStructType(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void FreeCStructType(Tcl_Obj *objP);
static void UpdateCStructTypeString(Tcl_Obj *objP);
static struct Tcl_ObjType gCStructType = {
    "TwapiCStruct",
    FreeCStructType,
    DupCStructType,
    UpdateCStructTypeString, /* Will panic. Not expected to be called as
                                we never change the string rep once created.
                                Could set this to NULL, but prefer to have
                                an explicit panic message */
    NULL, /* jenglish says keep this NULL */
};

typedef struct TwapiCStructField_s {
    Tcl_Obj *name;           /* Name of the field */
    Tcl_Obj *child;          /* Nested structure name or NULL */
    unsigned int count;         /* If >=1, field is an array of this size
                                   Note == 1 still means operand will be
                                   treated as a list and first element used.
                                   If 0, scalar */
    unsigned int offset;        /* Offset from beginning of structure */
    unsigned int size;          /* Size of the field. As int as we may allow
                                   nesting in the future */
    char type;                  /* The type of the field */
} TwapiCStructField;
#define CSTRUCT_REP(o_)    ((TwapiCStructRep *)((o_)->internalRep.twoPtrValue.ptr1))

typedef struct TwapiCStructRep_s {
    int nrefs;                  /* Reference count for this structure */
    int nfields;                /* Number of elements in fields[] */
    int size;                   /* Size of the defined structure */
    TwapiCStructField fields[1];
} TwapiCStructRep;

static CStructRepDecrRefs(TwapiCStructRep *csP)
{
    csP->nrefs -= 1;
    if (csP->nrefs <= 0) {
        int i;
        for (i = 0; i < csP->nrefs; ++i) {
            if (csP->fields[i].name)
                ObjDecrRefs(csP->fields[i].name);
            if (csP->fields[i].child)
                ObjDecrRefs(csP->fields[i].child);
        }
        TwapiFree(csP);
    }
}

static void DupCStructType(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    dstP->typePtr = srcP->typePtr;
    dstP->internalRep = srcP->internalRep;
    CSTRUCT_REP(dstP)->nrefs += 1;
}

static void FreeCStructType(Tcl_Obj *objP)
{
    TwapiCStructRep *csP = CSTRUCT_REP(objP);
    TWAPI_ASSERT(csP->nrefs > 0);
    CStructRepDecrRefs(csP);
    objP->internalRep.twoPtrValue.ptr1 = NULL;
    objP->internalRep.twoPtrValue.ptr2 = NULL;
    objP->typePtr = NULL;
}

static void UpdateCStructTypeString(Tcl_Obj *objP)
{
    Tcl_Panic("UpdateCStructTypeString called.");
}

/* These string are chosen to match the ones used in VARIANT. No
   particular reason, just consistency */
static const char *cstruct_types[] = {
    "bool", "i1", "ui1", "i2", "ui2", "i4", "ui4", "i8", "ui8", "r8", "lpstr", "lpwstr", "cbsize", "handle", NULL
};
enum cstruct_types_enum {
    CSTRUCT_BOOLEAN, CSTRUCT_CHAR, CSTRUCT_UCHAR, CSTRUCT_SHORT, CSTRUCT_USHORT, CSTRUCT_INT, CSTRUCT_UINT, CSTRUCT_INT64, CSTRUCT_UINT64,
    CSTRUCT_DOUBLE, CSTRUCT_STRING, CSTRUCT_WSTRING, CSTRUCT_CBSIZE, CSTRUCT_HANDLE
};
TCL_RESULT ObjCastToCStruct(Tcl_Interp *interp, Tcl_Obj *csObj)
{
    Tcl_Obj **fieldObjs;
    int       i, nfields;
    TwapiCStructRep *csP = NULL;
    unsigned int offset, struct_alignment;
    TCL_RESULT res;

    if (csObj->typePtr == &gCStructType)
        return TCL_OK;          /* Already the correct type */

    /* Make sure the string rep exists before we convert as we
       don't supply a string generation procedure
    */
    ObjToString(csObj);

    if ((res = ObjGetElements(interp, csObj, &nfields, &fieldObjs)) != TCL_OK)
        return res;
    if (nfields == 0)
        goto invalid_def;

    csP = TwapiAlloc(sizeof(*csP) + (nfields-1)*sizeof(csP->fields[0]));
    csP->nrefs = 1;


    /* Now parse all the elements */

    /* We will update csP->nfields as we loop so we can free the right number
       of fields if there is an error part way through */
    csP->nfields = 0;
    struct_alignment = 1;
    for (offset = 0, i = 0; i < nfields; ++i, csP->nfields = i) {
        Tcl_Obj **defObjs;      /* Field definition elements */
        int       ndefs;
        int       array_size = 0;
        int       deftype;
        int       elem_size;

        if (ObjGetElements(interp, fieldObjs[i], &ndefs, &defObjs) != TCL_OK ||
            ndefs < 2 || ndefs > 3 ||
            Tcl_GetIndexFromObj(interp, defObjs[1], cstruct_types, "type", TCL_EXACT, &deftype) != TCL_OK ||
            (ndefs == 3 &&
             (ObjToInt(interp, defObjs[2], &array_size) != TCL_OK ||
              array_size < 0))) {
            goto invalid_def;
        }

        switch (deftype) {
        case CSTRUCT_UCHAR: /* Fall thru */
        case CSTRUCT_CHAR: elem_size = sizeof(char); break;
        case CSTRUCT_USHORT: /* Fall thru */
        case CSTRUCT_SHORT: elem_size = sizeof(short); break;
        case CSTRUCT_BOOLEAN: /* Fall thru */
        case CSTRUCT_UINT: /* Fall thru */
        case CSTRUCT_INT: elem_size = sizeof(int); break;
        case CSTRUCT_UINT64: /* Fall thru */
        case CSTRUCT_INT64: elem_size = sizeof(__int64); break;
        case CSTRUCT_DOUBLE: elem_size = sizeof(double); break;
        case CSTRUCT_STRING: elem_size = sizeof(char*); break;
        case CSTRUCT_WSTRING: elem_size = sizeof(WCHAR *); break;
        case CSTRUCT_CBSIZE:
            /* NOTE : if there are any strings in the struct,
               this does not include storage for the string themselves */
            if (array_size)
                goto invalid_def; /* Cannot be an array */
            elem_size = sizeof(int);
            break;
        case CSTRUCT_HANDLE: elem_size = sizeof(HANDLE); break;
        }

        if (elem_size > struct_alignment)
            struct_alignment = elem_size;
        /* See if offset needs to be aligned */
        offset = (offset + elem_size - 1) & ~(elem_size - 1);
        csP->fields[i].name = NULL;
        csP->fields[i].child = NULL;
        csP->fields[i].offset = offset;
        csP->fields[i].count = array_size;
        csP->fields[i].type = deftype;
        csP->fields[i].size = elem_size;
        /* Array size of 0 means a scalar, so size 1 */
        if (array_size == 0)
            offset += elem_size; /* For next field */
        else
            offset += array_size*elem_size;
    }

    /* Whole structure has to be aligned */
    csP->size = (offset + struct_alignment - 1) & ~(struct_alignment - 1);

    /* OK, valid opaque rep. Convert the passed object's internal rep */
    if (csObj->typePtr && csObj->typePtr->freeIntRepProc) {
        csObj->typePtr->freeIntRepProc(csObj);
    }
    csObj->typePtr = &gCStructType;
    CSTRUCT_REP(csObj) = csP;
    csObj->internalRep.twoPtrValue.ptr2 = NULL;

    return TCL_OK;

invalid_def:
    TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS,
                        "Invalid CStruct definition");
error_return:
    if (csP)
        TwapiFree(csP);
    return TCL_ERROR;
}

/* Caller responsible for cleanup for memlifoP in all cases, success or error */
TCL_RESULT ParseCStruct (Tcl_Interp *interp, MemLifo *memlifoP,
                         Tcl_Obj *csvalObj, DWORD *sizeP, void **ppv)
{
    Tcl_Obj **objPP;
    int i, nobjs;
    TCL_RESULT res;
    TwapiCStructRep *csP = NULL;
    void *pv;

    if (ObjGetElements(interp, csvalObj, &nobjs, &objPP) != TCL_OK ||
        (nobjs != 0 && nobjs != 2))
        goto invalid_def;
        
    /* Empty string means struct pointer is treated as NULL */
    if (nobjs == 0) {
        *sizeP = 0;
        *ppv = NULL;
        return TCL_OK;
    }

    if (ObjCastToCStruct(interp, objPP[0]) != TCL_OK)
        return TCL_ERROR;

    csP = CSTRUCT_REP(objPP[0]);
    csP->nrefs += 1;            /* So it is not shimmered away underneath us */
    
    if (ObjGetElements(interp, objPP[1], &nobjs, &objPP) != TCL_OK ||
        nobjs != csP->nfields)  /* Not correct number of values */
        goto invalid_def;
    
    pv = MemLifoAlloc(memlifoP, csP->size, NULL);
    for (i = 0; i < nobjs; ++i) {
        int count = csP->fields[i].count;
        void *pv2 = ADDPTR(pv, csP->fields[i].offset, void*);
        Tcl_Obj **arrayObj;
        int       nelems;        /* # elements in array */
        void *s;
        int j, elem_size, len;
        TCL_RESULT (*fn)(Tcl_Interp *, Tcl_Obj *, void *);

        /* count > 0 => array and source obj is a list (even if count == 1) */
        if (count) {
            if (ObjGetElements(interp, objPP[i], &nelems, &arrayObj) != TCL_OK)
                goto error_return;
            if (count > nelems) {
                TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Too few elements in cstruct array field");
                goto error_return;
            }
        }
        elem_size = csP->fields[i].size;

        switch (csP->fields[i].type) {
        case CSTRUCT_BOOLEAN: fn = ObjToBoolean; break;
        case CSTRUCT_CHAR: fn = ObjToCHAR; break;
        case CSTRUCT_UCHAR: fn = ObjToUCHAR; break;
        case CSTRUCT_SHORT: fn = ObjToSHORT; break;
        case CSTRUCT_USHORT: fn = ObjToUSHORT; break;
        case CSTRUCT_INT: fn = ObjToLong; break;
        case CSTRUCT_UINT: fn = ObjToLong; break; // TBD - handles unsigned ?
        case CSTRUCT_INT64: fn = ObjToWideInt; break;
        case CSTRUCT_UINT64: fn = ObjToWideInt; break; // TBD-handles unsigned ?
        case CSTRUCT_DOUBLE: fn = ObjToDouble; break;
        case CSTRUCT_HANDLE: fn = ObjToHANDLE; break;
        case CSTRUCT_STRING:
            if (count) {
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    s = ObjToStringN(arrayObj[j], &len);
                    *(char **)pv2 = MemLifoCopy(memlifoP, s, len);
                }                
            } else {
                s = ObjToStringN(objPP[i], &len);
                *(char **)pv2 = MemLifoCopy(memlifoP, s, len);
            }
            continue;

        case CSTRUCT_WSTRING:
            if (count) {
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    s = ObjToUnicodeN(arrayObj[j], &len);
                    *(char **)pv2 = MemLifoCopy(memlifoP, s, len);
                }
            } else {
                s = ObjToUnicodeN(objPP[i], &len);
                *(char **)pv2 = MemLifoCopy(memlifoP, s, sizeof(WCHAR)*len);
            }
            continue;

        case CSTRUCT_CBSIZE:
            TWAPI_ASSERT(count == 0);
            fn = ObjToInt;
            break;

        default:
            TwapiReturnErrorEx(interp, TWAPI_BUG, Tcl_ObjPrintf("Unknown Cstruct type %d", csP->fields[i].type));
        }

        if (count) {
            for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                if (fn(interp, arrayObj[j], pv2) != TCL_OK)
                        goto error_return;
            }
        } else {
            if (fn(interp, objPP[i], pv2) != TCL_OK)
                goto error_return;
            if (csP->fields[i].type == CSTRUCT_CBSIZE) {
                int ssize = *(int *)pv2;
                if (ssize > csP->size || ssize < 0)
                    goto invalid_def;
                if (ssize == 0)
                    *(int *)pv2 = csP->size;
            }
        }
    }
    
    *ppv = pv;
    *sizeP = csP->size;
    CStructRepDecrRefs(csP);
    return TCL_OK;

invalid_def:
    TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS,
                        "Invalid CStruct value");
error_return:
    if (csP)
        CStructRepDecrRefs(csP);
    return TCL_ERROR;
}

TCL_RESULT ObjFromCStruct(Tcl_Interp *interp, void *pv, int nbytes, Tcl_Obj *csObj, Tcl_Obj **objPP)
{
    TwapiCStructRep *csP = NULL;
    Tcl_Obj *objs[32];          /* Assume no more than 32 fields in a struct */
    int i;

    if (ObjCastToCStruct(interp, csObj) != TCL_OK)
        return TCL_ERROR;

    csP = CSTRUCT_REP(objPP[0]);
    
    if (nbytes != 0 && nbytes != csP->size) 
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "Size mismatch with cstruct definition");

    if (csP->nfields > ARRAYSIZE(objs))
        return TwapiReturnErrorMsg(interp, TWAPI_INTERNAL_LIMIT, "Not enough space to decode all cstruct fields");

    TWAPI_ASSERT(csP->nrefs > 0);
    csP->nrefs += 1;            /* So it is not shimmered away underneath us */
    for (i = 0; i < csP->nfields; ++i) {
        int count = csP->fields[i].count;
        void *pv2 = ADDPTR(pv, csP->fields[i].offset, void*);
        Tcl_Obj *arrayObj;
        int       nelems;        /* # elements in array */
        void *s;
        int j, elem_size, len;
        TCL_RESULT (*fn)(Tcl_Interp *, Tcl_Obj *, void *);

        elem_size = csP->fields[i].size;

#define EXTRACT(type_, fn_)                                             \
        do {                                                            \
            if (count) {                                                \
                arrayObj = ObjNewList(count, NULL);                     \
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, sizeof(type_), void*)) { \
                    ObjAppendElement(NULL, arrayObj, (fn_)(*(type_ *)pv2)); \
                }                                                       \
                objs[i] = arrayObj;                                     \
            } else                                                      \
                objs[i] = (fn_)(*(type_ *)pv2);                         \
        } while (0)
        switch (csP->fields[i].type) {
        case CSTRUCT_BOOLEAN: EXTRACT(int, ObjFromBoolean); break;
        case CSTRUCT_CHAR: EXTRACT(char, ObjFromInt); break;
        case CSTRUCT_UCHAR: EXTRACT(unsigned char, ObjFromInt); break;
        case CSTRUCT_SHORT: EXTRACT(short, ObjFromInt); break;
        case CSTRUCT_USHORT: EXTRACT(unsigned short, ObjFromInt); break;
        case CSTRUCT_INT: EXTRACT(int, ObjFromInt); break;
        case CSTRUCT_UINT: EXTRACT(DWORD, ObjFromWideInt); break;
        case CSTRUCT_INT64: EXTRACT(__int64, ObjFromWideInt); break;
        case CSTRUCT_UINT64: EXTRACT(__int64, ObjFromWideInt); break; // TBD-handles unsigned ?
        case CSTRUCT_DOUBLE: EXTRACT(double, ObjFromDouble); break;
        case CSTRUCT_HANDLE: EXTRACT(HANDLE, ObjFromHANDLE); break;
        case CSTRUCT_STRING: EXTRACT(char*, ObjFromString); break;
        case CSTRUCT_WSTRING: EXTRACT(WCHAR*, ObjFromUnicode); break;
        case CSTRUCT_CBSIZE: EXTRACT(DWORD, ObjFromDWORD); break;
        }
    }
    
    *objPP = ObjNewList(csP->nfields, objs);
    if (csP)
        CStructRepDecrRefs(csP);
    return TCL_OK;
}



#if TWAPI_ENABLE_INSTRUMENTATION
TCL_RESULT TwapiCStructDefDump(Tcl_Interp *interp, Tcl_Obj *csObj)
{
    TwapiCStructRep *csP = NULL;
    Tcl_Obj *objP;
    int i;

    if (ObjCastToCStruct(interp, csObj) != TCL_OK)
            return TCL_ERROR;
        
    csP = CSTRUCT_REP(csObj);
    objP = Tcl_ObjPrintf("nrefs=%d, nfields=%d, size=%d",
                         csP->nrefs, csP->nfields, csP->size);
    for (i = 0; i < csP->nfields; ++i) {
        Tcl_AppendPrintfToObj(objP, "\n\tField %s: count=%d, offset=%d, size=%d, type=%d",
                              ObjToString(csP->fields[i].name),
                              csP->fields[i].count,
                              csP->fields[i].offset, 
                              csP->fields[i].size,
                              csP->fields[i].type);
    }
    ObjSetResult(interp, objP);
    return TCL_OK;
}
#endif
