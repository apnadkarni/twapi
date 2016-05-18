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

struct TwapiCStructRep_s;       /* Forward decl */
typedef struct TwapiCStructField_s {
    Tcl_Obj *name;           /* Name of the field */
    struct TwapiCStructRep_s *child; /* Nested structure name or NULL */
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
#define CSTRUCT_REP_SET(o_)    (o_)->internalRep.twoPtrValue.ptr1

typedef struct TwapiCStructRep_s {
    int nrefs;                  /* Reference count for this structure */
    int nfields;                /* Number of elements in fields[] */
    unsigned int size;                   /* Size of the defined structure */
    unsigned int alignment;              /* Alignment required (calculated based
                                   on contained fields) */
    TwapiCStructField fields[1];
} TwapiCStructRep;

static void CStructRepDecrRefs(TwapiCStructRep *csP)
{
    csP->nrefs -= 1;
    if (csP->nrefs <= 0) {
        int i;
        for (i = 0; i < csP->nrefs; ++i) {
            if (csP->fields[i].name)
                ObjDecrRefs(csP->fields[i].name);
            if (csP->fields[i].child)
                CStructRepDecrRefs(csP->fields[i].child);
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
    "bool", "i1", "ui1", "i2", "ui2", "i4", "ui4", "i8", "ui8", "r8", "lpstr", "lpwstr", "cbsize", "handle", "psid", "struct", NULL
};
enum cstruct_types_enum {
    CSTRUCT_BOOLEAN, CSTRUCT_CHAR, CSTRUCT_UCHAR, CSTRUCT_SHORT, CSTRUCT_USHORT, CSTRUCT_INT, CSTRUCT_UINT, CSTRUCT_INT64, CSTRUCT_UINT64,
    CSTRUCT_DOUBLE, CSTRUCT_STRING, CSTRUCT_WSTRING, CSTRUCT_CBSIZE, CSTRUCT_HANDLE, CSTRUCT_PSID, CSTRUCT_STRUCT
};

TWAPI_EXTERN TCL_RESULT ObjCastToCStruct(Tcl_Interp *interp, Tcl_Obj *csObj)
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
        DWORD       field_alignment;
        TwapiCStructRep  *child = NULL;

        /* Note for ease of cleanup in case of errors, we only set
           the csP fields for this field entry at bottom of loop when 
           sure that no errors are possible */

        if (ObjGetElements(interp, fieldObjs[i], &ndefs, &defObjs) != TCL_OK ||
            ndefs < 2 || ndefs > 4 ||
            Tcl_GetIndexFromObj(interp, defObjs[1], cstruct_types, "type", TCL_EXACT, &deftype) != TCL_OK ||
            (ndefs > 2 &&
             (ObjToInt(interp, defObjs[2], &array_size) != TCL_OK ||
              array_size < 0))) {
            goto invalid_def;
        }

        field_alignment = 0; /* Use elem_size for all except structs */


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
        case CSTRUCT_PSID: elem_size = sizeof(PSID); break;
        case CSTRUCT_STRUCT:
            if (ndefs < 4)
                goto invalid_def; /* Struct descriptor missing */
            if (ObjCastToCStruct(interp, defObjs[3]) != TCL_OK)
                goto error_return; /* Error message already in interp */
            TWAPI_ASSERT(defObjs[3]->typePtr == &gCStructType);
            child = CSTRUCT_REP(defObjs[3]);
            child->nrefs += 1;  /* Since we will link to it below */
            elem_size = child->size;
            field_alignment = child->alignment;
            break;
        default:
            goto invalid_def;   /* Should not happen... */
        }

        if (field_alignment == 0)
            field_alignment = elem_size;
        if (field_alignment > struct_alignment)
            struct_alignment = field_alignment;

        /* See if offset needs to be aligned */
        offset = (offset + field_alignment - 1) & ~(field_alignment - 1);
        csP->fields[i].name = defObjs[0];
        ObjIncrRefs(defObjs[0]);
        csP->fields[i].child = child;
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
    csP->alignment = struct_alignment;
    csP->size = (offset + struct_alignment - 1) & ~(struct_alignment - 1);

    /* OK, valid opaque rep. Convert the passed object's internal rep */
    if (csObj->typePtr && csObj->typePtr->freeIntRepProc) {
        csObj->typePtr->freeIntRepProc(csObj);
    }
    csObj->typePtr = &gCStructType;
    CSTRUCT_REP_SET(csObj) = csP;
    csObj->internalRep.twoPtrValue.ptr2 = NULL;

    return TCL_OK;

invalid_def:
    TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS,
                        "Invalid CStruct definition");
error_return:
    if (csP)
        CStructRepDecrRefs(csP);
    return TCL_ERROR;
}


/* Caller responsible for all MemLifo releasing */
static TCL_RESULT ParseCStructHelper (Tcl_Interp *interp, MemLifo *memlifoP,
                                      TwapiCStructRep *csP, 
                                      Tcl_Obj *valObj, DWORD flags,
                                      DWORD size, void *pv)
{
    Tcl_Obj **objPP;
    int i, nobjs;
    TCL_RESULT res;

    csP->nrefs += 1;            /* So it is not shimmered away underneath us */
    
    if (csP->size != size) {
        TwapiReturnError(interp, TWAPI_BUFFER_OVERRUN);
        goto error_return;
    }

    if (ObjGetElements(interp, valObj, &nobjs, &objPP) != TCL_OK ||
        nobjs != csP->nfields)  /* Not correct number of values */
        goto invalid_def;
    
    for (i = 0; i < nobjs; ++i) {
        int count = csP->fields[i].count;
        void *pv2 = ADDPTR(pv, csP->fields[i].offset, void*);
        Tcl_Obj **arrayObj;
        int       nelems;        /* # elements in array */
        void *s;
        int j, elem_size, len;
#if 0
        TCL_RESULT (*fn)(Tcl_Interp *, Tcl_Obj *, void *);
#else
        TCL_RESULT (*fn)();
#endif
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

        /* TBD - can combine the count == 0/count != 0 cases
           by setting arrayObjs to &objPP[i] and count = 1 for the former case
        */

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
        case CSTRUCT_PSID: 
            if (count) {
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    PSID sidP;
                    if (! ConvertStringSidToSidA(ObjToString(arrayObj[j]), &sidP))
                        goto invalid_def;
                    len = GetLengthSid(sidP);
                    *(PSID*)pv2 = MemLifoAlloc(memlifoP, len, NULL);
                    if (! CopySid(len, *(PSID*)pv2, sidP)) {
                        LocalFree(sidP);
                        goto invalid_def;
                    }
                    LocalFree(sidP);
                }                
            } else {
                PSID sidP;
                if (! ConvertStringSidToSidA(ObjToString(objPP[i]), &sidP))
                    goto invalid_def;
                len = GetLengthSid(sidP);
                *(PSID*)pv2 = MemLifoAlloc(memlifoP, len, NULL);
                if (! CopySid(len, *(PSID*)pv2, sidP)) {
                    LocalFree(sidP);
                    goto invalid_def;
                }
                LocalFree(sidP);
                
            }
            continue;

        case CSTRUCT_STRING:
            if (count) {
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    s = ObjToStringN(arrayObj[j], &len);
                    *(char **)pv2 = MemLifoCopy(memlifoP, s, len+1);
                }                
            } else {
                s = ObjToStringN(objPP[i], &len);
                *(char **)pv2 = MemLifoCopy(memlifoP, s, len+1);
            }
            continue;

        case CSTRUCT_WSTRING:
            if (count) {
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    s = ObjToWinCharsN(arrayObj[j], &len);
                    *(char **)pv2 = MemLifoCopy(memlifoP, s, sizeof(WCHAR)*(len+1));
                }
            } else {
                s = ObjToWinCharsN(objPP[i], &len);
                *(char **)pv2 = MemLifoCopy(memlifoP, s, sizeof(WCHAR)*(len+1));
            }
            continue;

        case CSTRUCT_CBSIZE:
            TWAPI_ASSERT(count == 0);
            fn = ObjToInt;
            break;

        case CSTRUCT_STRUCT:
            /*
             * objPP[i] is a nested struct value or array of them.
             * csP->fields[i].child points to its definition.
             */
            TWAPI_ASSERT(csP->fields[i].child);
            if (count) {
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    res = ParseCStructHelper(interp, memlifoP, csP->fields[i].child,
                                             arrayObj[j], flags, elem_size, pv2);
                    if (res != TCL_OK)
                        goto invalid_def;
                }                
            } else {
                res = ParseCStructHelper(interp, memlifoP, csP->fields[i].child,
                                         objPP[i], flags, elem_size, pv2);
                if (res != TCL_OK)
                    goto invalid_def;
            }
            continue;

        default:
            TwapiReturnErrorEx(interp, TWAPI_BUG, Tcl_ObjPrintf("Unknown Cstruct type %d", csP->fields[i].type));
            goto error_return;
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
                unsigned int ssize = *(int *)pv2;
                if (ssize > csP->size || ssize < 0)
                    goto invalid_def;
                if (ssize == 0)
                    *(int *)pv2 = csP->size;
            }
        }
    }
    
    CStructRepDecrRefs(csP);
    return TCL_OK;

invalid_def:
    TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS,
                        "Invalid CStruct value");
error_return:
    CStructRepDecrRefs(csP);
    return TCL_ERROR;
}

TCL_RESULT TwapiCStructParse (Tcl_Interp *interp, MemLifo *memlifoP,
                         Tcl_Obj *csvalObj, DWORD flags, DWORD *sizeP, void **ppv)
{
    Tcl_Obj **objPP;
    int nobjs;
    TCL_RESULT res;
    TwapiCStructRep *csP = NULL;
    void *pv;

    if (ObjGetElements(interp, csvalObj, &nobjs, &objPP) != TCL_OK ||
        (nobjs != 0 && nobjs != 2))
        goto invalid_def;
        
    /* Empty string means struct pointer is treated as NULL */
    if (nobjs == 0) {
        if (flags & CSTRUCT_ALLOW_NULL) {
            *sizeP = 0;
            *ppv = NULL;
            return TCL_OK;
        } else
            goto invalid_def;
    }

    if (ObjCastToCStruct(interp, objPP[0]) != TCL_OK)
        goto error_return;

    csP = CSTRUCT_REP(objPP[0]);
    csP->nrefs += 1;            /* So it is not shimmered away underneath us */
    /* REMEMBER to release csP from this point on */

    pv = MemLifoAlloc(memlifoP, csP->size, NULL);
    res = ParseCStructHelper(interp, memlifoP, csP, objPP[1], flags, csP->size, pv);
    if (res == TCL_OK) {
        *sizeP = csP->size;
        *ppv = pv;

        CStructRepDecrRefs(csP);
        return TCL_OK;
    }

invalid_def:
    TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS,
                        "Invalid CStruct value");
error_return:
    if (csP)
        CStructRepDecrRefs(csP);
    return TCL_ERROR;
}


static TCL_RESULT ObjFromCStructHelper(Tcl_Interp *interp, void *pv, unsigned int nbytes, TwapiCStructRep *csP, DWORD flags, Tcl_Obj **objPP)
{
    Tcl_Obj *objs[2*32]; /* Assume no more than 32 fields in a struct */
    unsigned int objindex, include_key;
    int i, max_fields;

    TWAPI_ASSERT(csP->nrefs > 0);
    csP->nrefs += 1;            /* So it is not shimmered away underneath us */

    if (nbytes != 0 && nbytes != csP->size) {
        TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "Size mismatch with cstruct definition");
        goto error_return;
    }

    if (flags & CSTRUCT_RETURN_DICT) {
        include_key = 1;
        max_fields = ARRAYSIZE(objs) / 2;
    } else {
        include_key = 0;
        max_fields = ARRAYSIZE(objs);
    }

    if (csP->nfields > max_fields) {
        TwapiReturnErrorMsg(interp, TWAPI_INTERNAL_LIMIT, "Not enough space to decode all cstruct fields");
        goto error_return;
    }

    for (i = 0, objindex = 0; i < csP->nfields; ++i) {
        int count = csP->fields[i].count;
        void *pv2 = ADDPTR(pv, csP->fields[i].offset, void*);
        Tcl_Obj *arrayObj;
        int j, elem_size;

        elem_size = csP->fields[i].size;

#define EXTRACT(type_, fn_)                                             \
        do {                                                            \
            if (include_key) objs[objindex++] = csP->fields[i].name;    \
            if (count) {                                                \
                arrayObj = ObjNewList(count, NULL);                     \
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, sizeof(type_), void*)) { \
                    ObjAppendElement(NULL, arrayObj, (fn_)(*(type_ *)pv2)); \
                }                                                       \
                objs[objindex++] = arrayObj;                            \
            } else {                                                    \
                objs[objindex++] = (fn_)(*(type_ *)pv2);                \
            }                                                           \
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
        case CSTRUCT_WSTRING: EXTRACT(WCHAR*, ObjFromWinChars); break;
        case CSTRUCT_CBSIZE: EXTRACT(DWORD, ObjFromDWORD); break;
        case CSTRUCT_PSID:
            if (include_key) objs[objindex++] = csP->fields[i].name;
            if (count) {
                arrayObj = ObjNewList(count, NULL);
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    Tcl_Obj *elemObj;

                    if (ObjFromSID(interp, *(PSID*)pv2, &elemObj) != TCL_OK) {
                        ObjDecrRefs(arrayObj);
                        goto error_return;
                    }
                    ObjAppendElement(NULL, arrayObj, elemObj);
                }
                objs[objindex++] = arrayObj;
            } else {
                if (ObjFromSID(interp, *(PSID*)pv2, &objs[objindex]) != TCL_OK)
                    goto error_return;
                objindex++;
            }
            break;

        case CSTRUCT_STRUCT:
            if (include_key) objs[objindex++] = csP->fields[i].name;
            if (count) {
                arrayObj = ObjNewList(count, NULL);
                for (j = 0; j < count; j++, pv2 = ADDPTR(pv2, elem_size, void*)) {
                    Tcl_Obj *elemObj;
                    if (ObjFromCStructHelper(interp, pv2, elem_size,
                                             csP->fields[i].child, flags, &elemObj)
                        != TCL_OK) {
                        ObjDecrRefs(arrayObj);
                        goto error_return;
                    }
                    ObjAppendElement(NULL, arrayObj, elemObj);
                }
                objs[objindex++] = arrayObj;
            } else {
                TWAPI_ASSERT(csP->fields[i].child);
                if (ObjFromCStructHelper(interp, pv2, elem_size,
                                         csP->fields[i].child, flags, &objs[objindex])
                    != TCL_OK)
                    goto error_return;
                objindex++;
            }
            break;
        }
    }
    
    *objPP = ObjNewList(objindex, objs);
    CStructRepDecrRefs(csP);
    return TCL_OK;

error_return:
    CStructRepDecrRefs(csP);
    return TCL_ERROR;

}

TCL_RESULT ObjFromCStruct(Tcl_Interp *interp, void *pv, int nbytes, Tcl_Obj *csObj, DWORD flags, Tcl_Obj **objPP)
{
    Tcl_Obj *objP;

    if (ObjCastToCStruct(interp, csObj) == TCL_OK &&
        ObjFromCStructHelper(interp, pv, nbytes, CSTRUCT_REP(csObj), flags, &objP) == TCL_OK) {
        if (objPP)
            *objPP = objP;
        else
            ObjSetResult(interp, objP);
        return TCL_OK;
    }
    return TCL_ERROR;
}


TCL_RESULT TwapiCStructSize(Tcl_Interp *interp, Tcl_Obj *csObj, int *szP)
{
    TwapiCStructRep *csP;
    TCL_RESULT res;
    res = ObjCastToCStruct(interp, csObj);
    if (res != TCL_OK)
        return res;
        
    csP = CSTRUCT_REP(csObj);
    *szP = csP->size;
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
