/*
 * Copyright (c) 2010-2014, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"
#include "tclTomMath.h"

/*
 * Struct for mapping VARTYPE values to strings.
 * We search linearly so order based on most likely types. - TBD make hash
 * Only the basic types are covered in this table. The code itself
 * handles the special/complex cases.
 */
struct vt_token_pair {
    VARTYPE vt;
    char   *tok;
};
static struct vt_token_pair vt_base_tokens[] = {
    {VT_BSTR, "bstr"},
    {VT_I4, "i4"},
    {VT_DISPATCH, "idispatch"},
    {VT_UNKNOWN, "iunknown"},
    {VT_BOOL, "bool"},
    {VT_VARIANT, "variant"},
    {VT_I2, "i2"},
    {VT_PTR, "ptr"},
    {VT_R4, "r4"},
    {VT_R8, "r8"},
    {VT_CY, "cy"},
    {VT_DATE, "date"},
    {VT_ERROR, "error"},
    {VT_UI1, "ui1"},
    {VT_DECIMAL, "decimal"},
    {VT_I1, "i1"},
    {VT_UI2, "ui2"},
    {VT_UI4, "ui4"},
    {VT_I8, "i8"},
    {VT_UI8, "ui8"},
    {VT_INT, "int"},
    {VT_UINT, "uint"},
    {VT_HRESULT, "hresult"},
    {VT_VOID, "void"},
    {VT_LPSTR, "lpstr"},
    {VT_LPWSTR, "lpwstr"},
    {VT_RECORD, "record"},
    {VT_USERDEFINED, "userdefined"},
    {VT_EMPTY, "empty"},
    {VT_NULL, "null"},
};

/* Support up to these many dimensions in a SAFEARRAY */
#define TWAPI_MAX_SAFEARRAY_DIMS 10

/*
 * Used for deciphering  unknown types when passing to COM. Note
 * any or all of these may be NULL.
 */
static struct TwapiTclTypeMap {
    char *typename;
    Tcl_ObjType *typeptr;
} gTclTypes[TWAPI_TCLTYPE_BOUND];

/*
 * TwapiOpaque is a Tcl "type" whose internal representation is stored
 * as the pointer value and an associated C pointer/handle type.
 * The Tcl_Obj.internalRep.twoPtrValue.ptr1 holds the C pointer value
 * andTcl_Obj.internalRep.twoPtrValue.ptr2 holds a Tcl_Obj describing
 * the type. This may be NULL if no type info needs to be associated
 * with the value
 */
static void DupOpaqueType(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void FreeOpaqueType(Tcl_Obj *objP);
static void UpdateOpaqueTypeString(Tcl_Obj *objP);
static struct Tcl_ObjType gOpaqueType = {
    "TwapiOpaque",
    FreeOpaqueType,
    DupOpaqueType,
    UpdateOpaqueTypeString,
    NULL,     /* jenglish says keep this NULL */
};


/*
 * TwapiVariant is a Tcl "type" whose internal representation preserves
 * the COM VARIANT type as far as possible.
 * The Tcl_Obj.internalRep.ptrAndLongRep.value holds the VT_* type tag.
 * and Tcl_Obj.internalRep.ptrAndLongRep.ptr holds a tag dependent value.
 */
static void DupVariantType(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void FreeVariantType(Tcl_Obj *objP);
static void UpdateVariantTypeString(Tcl_Obj *objP);
static struct Tcl_ObjType gVariantType = {
    "TwapiVariant",
    FreeVariantType,
    DupVariantType,
    UpdateVariantTypeString,
    NULL,     /* jenglish says keep this NULL */
};

static void TwapiInvalidVariantTypeMessage(Tcl_Interp *interp, VARTYPE vt)
{
    if (interp) {
        (void) ObjSetResult(interp,
                         Tcl_ObjPrintf("Invalid or unsupported VARTYPE (%d)",
                                       vt));
    }
}

static int LookupBaseVT(Tcl_Interp *interp, VARTYPE vt, const char **tokP)
{
    int i;
    for (i=0; i < ARRAYSIZE(vt_base_tokens); ++i) {
        if (vt_base_tokens[i].vt == vt) {
            if (tokP)
                *tokP = vt_base_tokens[i].tok;
            return TCL_OK;
        }
    }

    TwapiInvalidVariantTypeMessage(interp, vt);
    return TCL_ERROR;
}

static int LookupBaseVTToken(Tcl_Interp *interp, const char *tok, VARTYPE *vtP)
{
    int i;
    if (tok != NULL) {
        for (i=0; i < ARRAYSIZE(vt_base_tokens); ++i) {
            if (STREQ(vt_base_tokens[i].tok, tok)) {
                if (vtP)
                    *vtP = vt_base_tokens[i].vt;
                return TCL_OK;
            }
        }
    }
    if (interp) {
        Tcl_Obj *objP; 
        objP = STRING_LITERAL_OBJ("Invalid or unsupported VARTYPE token: ");
        Tcl_AppendToObj(objP, tok ? tok : "<null pointer>", -1);
        (void)ObjSetResult(interp, objP);
    }

    return TCL_ERROR;
}

static void UpdateOpaqueTypeString(Tcl_Obj *objP)
{
    Tcl_Obj *objs[2];
    Tcl_Obj *listObj;

    TWAPI_ASSERT(objP->bytes == NULL);
    TWAPI_ASSERT(objP->typePtr == &gOpaqueType);

    objs[0] = ObjFromDWORD_PTR(OPAQUE_REP_VALUE(objP));
    if (OPAQUE_REP_CTYPE(objP) == NULL)
        objs[1] = Tcl_NewObj();
    else
        objs[1] = OPAQUE_REP_CTYPE(objP);

    listObj = ObjNewList(2, objs);
    ObjToString(listObj);     /* Ensure string rep */

    /* We could just shift the bytes field from listObj to objP resetting
       the former to NULL. But I'm nervous about doing that behind Tcl's back */
    objP->length = listObj->length; /* Note does not include terminating \0 */
    objP->bytes = ckalloc(listObj->length + 1);
    CopyMemory(objP->bytes, listObj->bytes, listObj->length+1);
    ObjDecrRefs(listObj);
}

static void FreeOpaqueType(Tcl_Obj *objP)
{
    if (OPAQUE_REP_CTYPE(objP))
        ObjDecrRefs(OPAQUE_REP_CTYPE(objP));
    OPAQUE_REP_VALUE(objP) = NULL;
    OPAQUE_REP_CTYPE(objP) = NULL;
    objP->typePtr = NULL;
}

static void DupOpaqueType(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    dstP->typePtr = &gOpaqueType;
    OPAQUE_REP_VALUE(dstP) = OPAQUE_REP_VALUE(srcP);
    OPAQUE_REP_CTYPE(dstP) = OPAQUE_REP_CTYPE(srcP);
    if (OPAQUE_REP_CTYPE(dstP))
        ObjIncrRefs(OPAQUE_REP_CTYPE(dstP));
}

TCL_RESULT SetOpaqueFromAny(Tcl_Interp *interp, Tcl_Obj *objP)
{
    Tcl_Obj **objs;
    int       nobjs, val;
    void *pv;
    Tcl_Obj *ctype;
    char *s;

    if (objP->typePtr == &gOpaqueType)
        return TCL_OK;

    /* For backward compat with SWIG based script, we accept NULL
       as a valid pointer of any type and for convenience 0 as well */
    s = ObjToString(objP);
    if (STREQ(s, "NULL") ||
        (ObjToInt(NULL, objP, &val) == TCL_OK && val == 0)) {
        pv = NULL;
        ctype = NULL;
    } else {        
        DWORD_PTR dwp;

        /* Should be a two element list */
        if (ObjGetElements(NULL, objP, &nobjs, &objs) != TCL_OK ||
            nobjs != 2) {
            if (interp)
                Tcl_AppendResult(interp, "Invalid pointer or opaque value: '",
                                 s, "'.", NULL);
            return TCL_ERROR;
        }
        if (ObjToDWORD_PTR(NULL, objs[0], &dwp) != TCL_OK) {
            if (interp)
                Tcl_AppendResult(interp, "Invalid pointer or opaque value '",
                                 ObjToString(objs[0]), "'.", NULL);
            return TCL_ERROR;
        }
        pv = (void*) dwp;
        s = ObjToString(objs[1]);
        if (s[0] == 0)
            ctype = NULL;
        else {
            ctype = objs[1];
            ObjIncrRefs(ctype);
        }
    }

    /* OK, valid opaque rep. Convert the passed object's internal rep */
    if (objP->typePtr && objP->typePtr->freeIntRepProc) {
        objP->typePtr->freeIntRepProc(objP);
    }
    objP->typePtr = &gOpaqueType;
    OPAQUE_REP_VALUE(objP) = pv;
    OPAQUE_REP_CTYPE(objP) = ctype;
    return TCL_OK;
}

static void UpdateVariantTypeString(Tcl_Obj *objP)
{
    TWAPI_ASSERT(objP->bytes == NULL);
    TWAPI_ASSERT(objP->typePtr == &gVariantType);

    switch (VARIANT_REP_VT(objP)) {
    case VT_EMPTY:
    case VT_NULL:
        objP->length = 0; /* Note does not include terminating \0 */
        objP->bytes = ckalloc(1);
        objP->bytes[0] = 0;
        break;
    default:
        Tcl_Panic("Unexpected VT type (%d) in Tcl_Obj VARIANT", VARIANT_REP_VT(objP));
    }
}

static void FreeVariantType(Tcl_Obj *objP)
{
    /* Nothing to do since we never allocate any resources currently */
}

static void DupVariantType(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    dstP->typePtr = &gVariantType;
    dstP->internalRep = srcP->internalRep;
}


int TwapiInitTclTypes(void)
{
    int i;

    gTclTypes[TWAPI_TCLTYPE_NONE].typename = "none";
    gTclTypes[TWAPI_TCLTYPE_NONE].typeptr = NULL; /* No internal type set */
    gTclTypes[TWAPI_TCLTYPE_STRING].typename = "string";
    gTclTypes[TWAPI_TCLTYPE_BOOLEAN].typename = "boolean";
    gTclTypes[TWAPI_TCLTYPE_INT].typename = "int";
    gTclTypes[TWAPI_TCLTYPE_DOUBLE].typename = "double";
    gTclTypes[TWAPI_TCLTYPE_BYTEARRAY].typename = "bytearray";
    gTclTypes[TWAPI_TCLTYPE_LIST].typename = "list";
    gTclTypes[TWAPI_TCLTYPE_DICT].typename = "dict";
    gTclTypes[TWAPI_TCLTYPE_WIDEINT].typename = "wideInt";
    gTclTypes[TWAPI_TCLTYPE_BOOLEANSTRING].typename = "booleanString";

    for (i = 1; i < TWAPI_TCLTYPE_NATIVE_END; ++i) {
        gTclTypes[i].typeptr =
            Tcl_GetObjType(gTclTypes[i].typename); /* May be NULL */
    }

    /* "booleanString" type is not always registered (if ever). Get it
     *  by hook or by crook
     */
    if (gTclTypes[TWAPI_TCLTYPE_BOOLEANSTRING].typeptr == NULL) {
        Tcl_Obj *objP = STRING_LITERAL_OBJ("true");
        ObjToBoolean(NULL, objP, &i);
        /* This may still be NULL, but what can we do ? */
        gTclTypes[TWAPI_TCLTYPE_BOOLEANSTRING].typeptr = objP->typePtr;
    }    

    /* Add our TWAPI types */
    gTclTypes[TWAPI_TCLTYPE_OPAQUE].typename = gOpaqueType.name;
    gTclTypes[TWAPI_TCLTYPE_OPAQUE].typeptr = &gOpaqueType;
    gTclTypes[TWAPI_TCLTYPE_VARIANT].typename = gVariantType.name;
    gTclTypes[TWAPI_TCLTYPE_VARIANT].typeptr = &gVariantType;

    return TCL_OK;
}

int TwapiGetTclType(Tcl_Obj *objP)
{
    int i;
    
    for (i = 0; i < ARRAYSIZE(gTclTypes); ++i) {
        if (gTclTypes[i].typeptr == objP->typePtr)
            return i;
    }

    return -1;
}

int Twapi_GetTclTypeObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (objv[1]->typePtr) {
        if (objv[1]->typePtr == &gVariantType) {
            const char *typename;
            if (LookupBaseVT(NULL, (VARTYPE) VARIANT_REP_VT(objv[1]), &typename) == TCL_OK &&
                typename) {
                return ObjSetResult(interp, ObjFromString(typename));
            }
        }
        ObjSetResult(interp, ObjFromString(objv[1]->typePtr->name));
    } else {
        /* Leave result as empty string */
    }
    return TCL_OK;
}

int Twapi_InternalCastObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    Tcl_Obj *objP;
    Tcl_ObjType *typeP;
    const char *typename;
    int i;
    VARTYPE vt;

    if (objc != 3)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    typename = ObjToString(objv[1]);

    if (*typename == '\0') {
        /* No type, keep as is */
        return ObjSetResult(interp, objv[2]);
    }
        
    /*
     * We special case double, because SetAnyFromProc will optimize to lowest
     * compatible type, so for example casting 1 to double will result in an
     * int object. We want to force it to double.
     *
     * We special case "boolean" and "booleanString" because they will keep
     * numerics as numerics while we want to force to boolean Tcl_Obj type.
     * We do this even before GetObjType because "booleanString" is not
     * even a registered type in Tcl.
     *
     * We can't do anything about wideInt because Tcl_NewDoubleObj, Tcl_GetDoubleFromObj,
     * SetDoubleFromAny, will also return an int Tcl_Obj if the vallue fits in the 32 bits.
     */
    if (STREQ(typename, "double")) {
        double dval;
        if (ObjToDouble(interp, objv[2], &dval) == TCL_ERROR)
            return TCL_ERROR;
        return ObjSetResult(interp, ObjFromDouble(dval));
    }

    if (STREQ(typename, "boolean") || STREQ(typename, "booleanString")) {
        if (ObjToBoolean(interp, objv[2], &i) == TCL_ERROR)
            return TCL_ERROR;
        /* Directly calling Tcl_NewBooleanObj returns an int type object */
        objP = ObjFromString(i ? "true" : "false");
        ObjToBoolean(NULL, objP, &i);
        return ObjSetResult(interp, objP);
    }

    typeP = Tcl_GetObjType(typename);
    if (typeP) {
        if (objv[2]->typePtr == typeP) {
            /* If type is already correct, no need to do anything */
            objP = objv[2];
        } else {
            /*
             * Need to convert it. If not shared, do in place else allocate
             * new object
             */

            if (typeP->setFromAnyProc == NULL)
                goto error_handler;

            if (Tcl_IsShared(objv[2])) {
                objP = ObjDuplicate(objv[2]);
            } else {
                objP = objv[2];
            }
        

            if (Tcl_ConvertToType(interp, objP, typeP) == TCL_ERROR) {
                if (objP != objv[2]) {
                    ObjDecrRefs(objP);
                }
                return TCL_ERROR;
            }
        }

        return ObjSetResult(interp, objP);
    }

    /* Not a registered Tcl type. See if one of ours */
    if (LookupBaseVTToken(NULL, typename, &vt) == TCL_OK) {
        switch (vt) {
        case VT_EMPTY:
        case VT_NULL:
            if (ObjCharLength(objv[2]) != 0)
                goto error_handler;
            objP = ObjFromEmptyString();
            Tcl_InvalidateStringRep(objP);
            objP->typePtr = &gVariantType;
            VARIANT_REP_VT(objP) = vt;
            return ObjSetResult(interp, objP);
        }
    }

error_handler:
    Tcl_AppendResult(interp, "Cannot convert '", ObjToString(objv[2]), "' to type '", typename, "'", NULL);
    return TCL_ERROR;
}

/* Call to set static result */
void ObjSetStaticResult(Tcl_Interp *interp, CONST char s[])
{
    Tcl_SetResult(interp, (char *) s, TCL_STATIC);
}

TCL_RESULT ObjSetResult(Tcl_Interp *interp, Tcl_Obj *objP)
{
    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}

Tcl_Obj *ObjGetResult(Tcl_Interp *interp)
{
    return Tcl_GetObjResult(interp);
}

Tcl_Obj *ObjDuplicate(Tcl_Obj *objP)
{
    return Tcl_DuplicateObj(objP);
}

/*
 * Generic function for setting a Tcl result. Note following special cases
 * - for TRT_OBJV types, the objv[] objects are added to a list (and
 *   their ref counts are therby incremented)
 * - for TRT_VARIANT types, VariantClear is called. This allows
 *   callers to directly return after calling this function without
 *   having to clean up the variant before returning.
 * - Similarly, TRT_UNICODE_DYNAMIC, TRT_CHARS_DYNAMIC, TRT_OLESTR and TRT_PIDL
 *   have their memory freed
 *  The return value reflects whether the result was set from an error or not,
 *  NOT whether the result was successfully set (which it always is).
 **/
TCL_RESULT TwapiSetResult(Tcl_Interp *interp, TwapiResult *resultP)
{
    char *typenameP;
    Tcl_Obj *resultObj = NULL;

    switch (resultP->type) {
    case TRT_GETLASTERROR:      /* Error in GetLastError() */
        return TwapiReturnSystemError(interp);

    case TRT_BOOL:
        resultObj = ObjFromBoolean(resultP->value.bval);
        break;

    case TRT_EXCEPTION_ON_FALSE:
    case TRT_NONZERO_RESULT:
        /* If 0, generate exception */
        if (! resultP->value.ival)
            return TwapiReturnSystemError(interp);

        if (resultP->type == TRT_NONZERO_RESULT)
            resultObj = ObjFromLong(resultP->value.ival);
        /* else an empty result is returned */
        break;

    case TRT_EXCEPTION_ON_ERROR:
        /* If non-0, generate exception */
        if (resultP->value.ival) {
            return Twapi_AppendSystemError(interp, resultP->value.ival);
        }
        break;

    case TRT_EXCEPTION_ON_WNET_ERROR:
        /* If non-0, generate exception */
        if (resultP->value.ival) {
            return Twapi_AppendWNetError(interp, resultP->value.ival);
        }
        break;

    case TRT_EXCEPTION_ON_MINUSONE:
        /* If -1, generate exception */
        if (resultP->value.ival == -1)
            return TwapiReturnSystemError(interp);

        /* Other values are to be returned */
        resultObj = ObjFromLong(resultP->value.ival);
        break;

    case TRT_UNICODE_DYNAMIC:
    case TRT_UNICODE:
        if (resultP->value.unicode.str) {
            resultObj = ObjFromUnicodeN(resultP->value.unicode.str,
                                        resultP->value.unicode.len);
        }
        break;

    case TRT_CHARS_DYNAMIC:
    case TRT_CHARS:
        if (resultP->value.chars.str)
            resultObj = ObjFromStringN(resultP->value.chars.str,
                                         resultP->value.chars.len);
        break;

    case TRT_BINARY:
        resultObj = ObjFromByteArray(resultP->value.binary.p,
                                        resultP->value.binary.len);
        break;

    case TRT_OBJ:
        resultObj = resultP->value.obj;
        break;

    case TRT_OBJV:
        resultObj = ObjNewList(resultP->value.objv.nobj, resultP->value.objv.objPP);
        break;

    case TRT_RECT:
        resultObj = ObjFromRECT(&resultP->value.rect);
        break;

    case TRT_POINT:
        resultObj = ObjFromPOINT(&resultP->value.point);
        break;

    case TRT_VALID_HANDLE:
        if (resultP->value.hval == INVALID_HANDLE_VALUE) {
            return TwapiReturnSystemError(interp);
        }
        resultObj = ObjFromHANDLE(resultP->value.hval);
        break;

    case TRT_HWND:
        // Note unlike other handles, we do not return an error if NULL
        resultObj = ObjFromOpaque(resultP->value.hwin, "HWND");
        break;

    case TRT_HMODULE:
        // Note unlike other handles, we do not return an error if NULL
        resultObj = ObjFromOpaque(resultP->value.hmodule, "HMODULE");
        break;

    case TRT_HANDLE:
    case TRT_HGLOBAL:
    case TRT_HDC:
    case TRT_HDESK:
    case TRT_HMONITOR:
    case TRT_HWINSTA:
    case TRT_SC_HANDLE:
    case TRT_LSA_HANDLE:
    case TRT_HDEVINFO:
    case TRT_HRGN:
    case TRT_HKEY:
        if (resultP->value.hval == NULL) {
            return TwapiReturnSystemError(interp);
        }
        switch (resultP->type) {
        case TRT_HANDLE:
            typenameP = "HANDLE";
            break;
        case TRT_HGLOBAL:
            typenameP = "HGLOBAL";
            break;
        case TRT_HDC:
            typenameP = "HDC";
            break;
        case TRT_HDESK:
            typenameP = "HDESK";
            break;
        case TRT_HMONITOR:
            typenameP = "HMONITOR";
            break;
        case TRT_HWINSTA:
            typenameP = "HWINSTA";
            break;
        case TRT_SC_HANDLE:
            typenameP = "SC_HANDLE";
            break;
        case TRT_LSA_HANDLE:
            typenameP = "LSA_HANDLE";
            break;
        case TRT_HDEVINFO:
            typenameP = "HDEVINFO";
            break;
        case TRT_HRGN:
            typenameP = "HRGN";
            break;
        case TRT_HKEY:
            typenameP = "HKEY";
            break;
        default:
            ObjSetStaticResult(interp, "Internal error: TwapiSetResult - inconsistent nesting of case statements");
            return TCL_ERROR;
        }
        resultObj = ObjFromOpaque(resultP->value.hval, typenameP);
        break;

    case TRT_LONG:
        resultObj = ObjFromLong(resultP->value.ival);
        break;

    case TRT_DWORD:
        resultObj = ObjFromWideInt((Tcl_WideInt) resultP->value.uval);
        break;
        
    case TRT_WIDE:
        resultObj = ObjFromWideInt(resultP->value.wide);
        break;

    case TRT_DOUBLE:
        resultObj = Tcl_NewDoubleObj(resultP->value.dval);
        break;

    case TRT_FILETIME:
        resultObj = ObjFromFILETIME(&resultP->value.filetime);
        break;

    case TRT_SYSTEMTIME:
        resultObj = ObjFromSYSTEMTIME(&resultP->value.systime);
        break;

    case TRT_EMPTY:
        Tcl_ResetResult(interp);
        break;

    case TRT_UUID:
        resultObj = ObjFromUUID(&resultP->value.uuid);
        break;

    case TRT_GUID:
        resultObj = ObjFromGUID(&resultP->value.guid);
        break;

    case TRT_LUID:
        resultObj = ObjFromLUID(&resultP->value.luid);
        break;
        
    case TRT_DWORD_PTR:
        resultObj = ObjFromDWORD_PTR(resultP->value.dwp);
        break;

    case TRT_INTERFACE:
        resultObj = ObjFromOpaque(resultP->value.ifc.p, resultP->value.ifc.name);
        break;

    case TRT_VARIANT:
        resultObj = ObjFromVARIANT(&resultP->value.var, 0);
        break;

    case TRT_LPOLESTR:
        if (resultP->value.lpolestr) {
            resultObj = ObjFromUnicode(resultP->value.lpolestr);
        } else
            Tcl_ResetResult(interp);
        break;

    case TRT_PIDL:
        if (resultP->value.pidl) {
            resultObj = ObjFromPIDL(resultP->value.pidl);
        } else
            Tcl_ResetResult(interp);
        break;

    case TRT_NONNULL_PTR:
        if (resultP->value.ptr.p == NULL)
            return TwapiReturnSystemError(interp);
        /* FALLTHRU */
    case TRT_PTR:
        resultObj = ObjFromOpaque(resultP->value.ptr.p,
                                  resultP->value.ptr.name);
        break;

    case TRT_TCL_RESULT:
        /* interp result already stored. Status in ival */
        return resultP->value.ival;
        
    case TRT_NTSTATUS:
        if (resultP->value.ival != STATUS_SUCCESS)
            return Twapi_AppendSystemError(interp,
                                           TwapiNTSTATUSToError(resultP->value.ival));

        break;

    case TRT_BADFUNCTIONCODE:
        return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);

    case TRT_TWAPI_ERROR:
        if (resultP->value.ival != TWAPI_NO_ERROR)
            return TwapiReturnError(interp, resultP->value.ival);
        break;

    default:
        ObjSetStaticResult(interp, "Unknown TwapiResultType type code passed to TwapiSetResult");
        return TCL_ERROR;
    }

    TwapiClearResult(resultP);  /* Clear out resources */

    if (resultObj)
        ObjSetResult(interp, resultObj);


    return TCL_OK;
}

/* Frees allocated resources, sets resultP to type TRT_EMPTY */
void TwapiClearResult(TwapiResult *resultP)
{
    switch (resultP->type) {
    case TRT_UNICODE_DYNAMIC:
        if (resultP->value.unicode.str)
            TwapiFree(resultP->value.unicode.str);
        break;
    case TRT_CHARS_DYNAMIC:
        if (resultP->value.chars.str)
            TwapiFree(resultP->value.chars.str);
        break;
    case TRT_VARIANT:
        VariantClear(&resultP->value.var);
        break;
    case TRT_LPOLESTR:
        if (resultP->value.lpolestr)
            CoTaskMemFree(resultP->value.lpolestr);
        break;

    case TRT_PIDL:
        if (resultP->value.pidl)
            CoTaskMemFree(resultP->value.pidl);
        break;

    default:
        break;                  /* Nothing to clear */

    }

    resultP->type = TRT_EMPTY;
}


/* Appends the given strings objv[] to a result object, separated by
 * the passed string. The passed resultObj must not be a shared object!
 */
Tcl_Obj *TwapiAppendObjArray(Tcl_Obj *resultObj, int objc, Tcl_Obj **objv,
                         char *joiner)
{
    int i;
    int len;
    char *s;
    int joinlen = (int) strlen(joiner);
#if 0
    Not needed - Tcl_AppendToObj will do the panic for us below
    if (Tcl_IsShared(resultObj)) {
        panic("TwapiAppendObjArray called on shared object");
    }
#endif

    for (i = 0;  i < objc;  ++i) {
        s = ObjToStringN(objv[i], &len);
        if (i > 0) {
            Tcl_AppendToObj(resultObj, joiner, joinlen);
        }
        Tcl_AppendToObj(resultObj, s, len);
    }

    return resultObj;
}



LPWSTR ObjToLPWSTR_NULL_IF_EMPTY(Tcl_Obj *objP)
{
    if (objP) {
        int len;
        LPWSTR p = ObjToUnicodeN(objP, &len);
        if (len > 0)
            return p;
    }
    return NULL;
}

LPWSTR ObjToLPWSTR_WITH_NULL(Tcl_Obj *objP)
{
    if (objP) {
        LPWSTR s = ObjToUnicode(objP);
        if (lstrcmpW(s, NULL_TOKEN_L) == 0) {
            s = NULL;
        }
        return s;
    }
    return NULL;
}

// Return SysAllocStringLen-allocated string
int ObjToBSTR(Tcl_Interp *interp, Tcl_Obj *objP, BSTR *bstrP)
{
    int len;
    WCHAR *wcharP;

    if (objP == NULL) {
        wcharP = L"";
        len = 0;
    } else {
        wcharP = ObjToUnicodeN(objP, &len);
    }
    if (bstrP) {
        *bstrP = SysAllocStringLen(wcharP, len);
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, E_OUTOFMEMORY);
    }
}

Tcl_Obj *ObjFromBSTR (BSTR bstr)
{
    return bstr ?
        ObjFromUnicodeN(bstr, SysStringLen(bstr))
        : ObjFromEmptyString();
}

Tcl_Obj *ObjFromStringLimited(const char *strP, int max, int *remainP)
{
    const char *p, *endP;

    if (max < 0) {
        if (remainP)
            *remainP = 0;
        return ObjFromString(strP);
    }        

    p = strP;
    endP = max + strP;
    while (p < endP && *p)
        ++p;

    if (remainP) {
        if (p == endP)
            *remainP = 0;
        else
            *remainP = (int) (endP-p)-1; /* -1 to skip over \0 */
    }

    return ObjFromStringN(strP, (int) (p-strP));
}

Tcl_Obj *ObjFromUnicodeLimited(const WCHAR *strP, int max, int *remainP)
{
    const WCHAR *p, *endP;

    if (max < 0) {
        if (remainP)
            *remainP = 0;
        return ObjFromUnicode(strP);
    }        

    p = strP;
    endP = max + strP;
    while (p < endP && *p)
        ++p;

    if (remainP) {
        if (p == endP)
            *remainP = 0;
        else
            *remainP = (int) (endP-p)-1; /* -1 to skip over \0 */
    }

    return ObjFromUnicodeN(strP, (int) (p-strP));
}

/* Some Win32 APIs, ETW in particular return strings with a trailing space.
   Return a Tcl_Obj without this single trailing space if present */
Tcl_Obj *ObjFromUnicodeNoTrailingSpace(const WCHAR *strP)
{
    int len;

    len = lstrlenW(strP);
    if (len && strP[len-1] == L' ')
        --len;
    return ObjFromUnicodeN(strP, len);
}

/*
 * Gets an integer from an object within the specified range
 * Returns TCL_OK if integer within range [low,high], else error
 */
int ObjToRangedInt(Tcl_Interp *interp, Tcl_Obj *obj, int low, int high, int *iP)
{
    int i;
    if (ObjToInt(interp, obj, &i) != TCL_OK)
        return TCL_ERROR;

    if (i < low || i > high)
        return TwapiReturnError(interp, TWAPI_OUT_OF_RANGE);

    if (iP)
        *iP = i;

    return TCL_OK;
}

/*
 * Convert a system time structure to a list
 * Year month day hour min sec msecs
 */
Tcl_Obj *ObjFromSYSTEMTIME(const SYSTEMTIME *timeP)
{
    Tcl_Obj *objv[8];

    /* Fields are not in order they occur in SYSTEMTIME struct
       This is intentional for ease of formatting at script level */
    objv[0] = ObjFromInt(timeP->wYear);
    objv[1] = ObjFromInt(timeP->wMonth);
    objv[2] = ObjFromInt(timeP->wDay);
    objv[3] = ObjFromInt(timeP->wHour);
    objv[4] = ObjFromInt(timeP->wMinute);
    objv[5] = ObjFromInt(timeP->wSecond);
    objv[6] = ObjFromInt(timeP->wMilliseconds);
    objv[7] = ObjFromInt(timeP->wDayOfWeek);

    return ObjNewList(ARRAYSIZE(objv), objv);
}


/*
 * Convert a Tcl Obj to SYSTEMTIME
 */
TCL_RESULT ObjToSYSTEMTIME(Tcl_Interp *interp, Tcl_Obj *timeObj, LPSYSTEMTIME timeP)
{
    Tcl_Obj **objv;
    int       objc;
    FILETIME  ft;

    if (ObjGetElements(interp, timeObj, &objc, &objv) != TCL_OK)
        goto syntax_error;

    /*
     * List size - 
     * 0 - current date and time
     * 3-7 - year, month, day, ?hour, min, sec, ms?
     * Additional elements are ignored (since the reverse operation
     * returns day of week as 8th element)
     */
    if (objc == 0) {
        GetSystemTime(timeP);
        return TCL_OK;
    }
    if (objc < 3)
        goto syntax_error;

    if (ObjToWord(NULL, objv[0], &timeP->wYear) != TCL_OK ||
        ObjToWord(NULL, objv[1], &timeP->wMonth) != TCL_OK ||
        ObjToWord(NULL, objv[2], &timeP->wDay) != TCL_OK)
        goto syntax_error;
    timeP->wHour = timeP->wMinute = timeP->wSecond = timeP->wMilliseconds = 0;
    TWAPI_ASSERT(objc >= 3);
    switch (objc) {
    default: /* > 7 */
    case 7:
        if (ObjToWord(NULL, objv[6], &timeP->wMilliseconds) != TCL_OK)
            goto syntax_error;
    case 6:
        if (ObjToWord(NULL, objv[5], &timeP->wSecond) != TCL_OK)
            goto syntax_error;
    case 5:
        if (ObjToWord(NULL, objv[4], &timeP->wMinute) != TCL_OK)
            goto syntax_error;
    case 4:
        if (ObjToWord(NULL, objv[3], &timeP->wHour) != TCL_OK)
            goto syntax_error;
        break;
    case 3:
        break;
    }

    /* Validate field values */
    if (SystemTimeToFileTime(timeP, &ft))
        return TCL_OK;

syntax_error:
    ObjSetResult(interp, Tcl_ObjPrintf("Invalid time list '%s'", ObjToString(timeObj)));
    return TCL_ERROR;
}

Tcl_Obj *ObjFromFILETIME(FILETIME *ftimeP)
{
    LARGE_INTEGER large;
    large.LowPart = ftimeP->dwLowDateTime;
    large.HighPart = ftimeP->dwHighDateTime;
    return ObjFromLARGE_INTEGER(large);
}

int ObjToFILETIME(Tcl_Interp *interp, Tcl_Obj *obj, FILETIME *ftimeP)
{
    LARGE_INTEGER large;
    if (ObjToWideInt(interp, obj, &large.QuadPart) != TCL_OK)
        return TCL_ERROR;

    if (ftimeP) {
        ftimeP->dwLowDateTime = large.LowPart;
        ftimeP->dwHighDateTime = large.HighPart;
    }

    return TCL_OK;
}

Tcl_Obj *ObjFromCY(const CY *cyP)
{
    /* TBD - for now just return as 8 byte wide int */
    return ObjFromWideInt(*(Tcl_WideInt *)cyP);
}


int ObjToCY(Tcl_Interp *interp, Tcl_Obj *obj, CY *cyP)
{
    Tcl_WideInt wi;
    if (ObjToWideInt(interp, obj, &wi) != TCL_OK)
        return TCL_ERROR;

    if (cyP)
        *cyP = *(CY *)&wi;

    return TCL_OK;
}

Tcl_Obj *ObjFromDECIMAL(DECIMAL *decP)
{
    /* TBD - for now just return as string.
       Problem is how is this formatted (separators etc.) ?
    */
    Tcl_Obj *obj;
    BSTR bstr = NULL;
    if (VarBstrFromDec(decP, 0, 0, &bstr) != S_OK) {
        return ObjFromEmptyString();
    }

    obj = ObjFromBSTR(bstr);
    SysFreeString(bstr);
    return obj;
}


int ObjToDECIMAL(Tcl_Interp *interp, Tcl_Obj *obj, DECIMAL *decP)
{
    HRESULT hr;
    DECIMAL dec;

    if (decP == NULL)
        decP = &dec;
    hr = VarDecFromStr(ObjToUnicode(obj), 0, 0, decP);
    if (FAILED(hr)) {
        if (interp)
            Twapi_AppendSystemError(interp, hr);
        return TCL_ERROR;
    }
    return TCL_OK;
}

Tcl_Obj *ObjFromPIDL(LPCITEMIDLIST pidl)
{
    /* Scan until we find an item with length 0 */
    unsigned char *p = (char *) pidl;
    do {
        /* p[1,0] is length of this item */
        int len = 256*p[1] + p[0];
        if (len == 0)
            break;              /* 0 length -> end of list */

        p += len;
    } while (1);

    /* p points to terminating null field */
    return ObjFromByteArray((unsigned char *)pidl,
                               (int) (2 + p - (unsigned char *)pidl));

}

/* On success, returns TCL_OK and stores pointer to an ITEMIDLIST
   in *idlistP. The value stored may be NULL which is valid in many cases.
   On error, returns TCL_ERROR, with an error message in interp if it
   is not NULL.
   The ITEMIDLIST must be freed by caller by calling TwapiFreePIDL
*/
int ObjToPIDL(Tcl_Interp *interp, Tcl_Obj *objP, LPITEMIDLIST *idsPP)
{
    int      numbytes;
    LPITEMIDLIST idsP;

    idsP = (LPITEMIDLIST) ObjToByteArray(objP, &numbytes);
    if (numbytes < 2) {
        *idsPP = NULL;              /* Empty string */
        return TCL_OK;
    }
    else {
        /* Verify format. Passing bad PIDL's can crash the app */
        unsigned char *p = (unsigned char *) idsP;
        /*
         * At top of loop p points to length of next item
         * and p[0] and p[1] are guaranteed valid part of buffer.
         */
        while (p[0] || p[1]) {
            int itemlen = p[1]*256 + p[0]; /* Assumes little endian */

            /* Verify that item length fits within buffer
             * The 2 is for the trailing 2 null bytes
             */
            if (itemlen > (numbytes-2)) {
                ObjSetStaticResult(interp, "Invalid item id list format");
                return TCL_ERROR;
            }
            numbytes -= itemlen;
            p += itemlen;
        }
    }

    *idsPP = CoTaskMemAlloc(numbytes);
    if (*idsPP == NULL) {
        if (interp)
            ObjSetStaticResult(interp, "CoTaskMemAlloc failed in SHChangeNotify");
        return TCL_ERROR;
    }

    CopyMemory(*idsPP, idsP, numbytes);

    return TCL_OK;
}

void TwapiFreePIDL(LPITEMIDLIST idlistP)
{
    if (idlistP) {
        CoTaskMemFree(idlistP);
    }
}

Tcl_Obj *ObjFromGUID(GUID *guidP)
{
    wchar_t  str[40];
    Tcl_Obj *obj;


    if (guidP == NULL || StringFromGUID2(guidP, str, sizeof(str)/sizeof(str[0])) == 0)
        return ObjFromEmptyString("", 0);

    obj = ObjFromUnicode(str);
    return obj;
}

int ObjToGUID(Tcl_Interp *interp, Tcl_Obj *objP, GUID *guidP)
{
    HRESULT hr;
    WCHAR *wsP;
    if (objP) {
        wsP = ObjToUnicode(objP);

        /* Accept both GUID and UUID forms */
        if (*wsP == L'{') {
            /* GUID form */
            /* We *used* to use CLSIDFromString but it turns out that 
               accepts Prog IDs as valid GUIDs as well */
            if ((hr = IIDFromString(wsP, guidP)) != NOERROR) {
                Twapi_AppendSystemError(interp, hr);
                return TCL_ERROR;
            }
        } else {
            /* Might be UUID form */
            RPC_STATUS status = UuidFromStringW(wsP, guidP);
            if (status != RPC_S_OK) {
                Twapi_AppendSystemError(interp, status);
                return TCL_ERROR;
            }
        }
    } else
        ZeroMemory(guidP, sizeof(*guidP));
    return TCL_OK;
}

int ObjToGUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, GUID **guidPP)
{
    if (ObjCharLength(objP) == 0) {
        *guidPP = NULL;
        return TCL_OK;
    } else 
        return ObjToGUID(interp, objP, *guidPP);
}


Tcl_Obj *ObjFromUUID (UUID *uuidP)
{
    unsigned char *uuidStr;
    Tcl_Obj       *objP;
    if (UuidToStringA(uuidP, &uuidStr) != RPC_S_OK)
        return NULL;

    /* NOTE UUID and GUID have same binary format but are formatted
       differently based on the component. */
    objP = ObjFromString(uuidStr);
    RpcStringFree(&uuidStr);
    return objP;
}

int ObjToUUID(Tcl_Interp *interp, Tcl_Obj *objP, UUID *uuidP)
{
    /* NOTE UUID and GUID have same binary format but are formatted
       differently based on the component.  We accept both forms here */

    if (objP) {
        RPC_STATUS status = UuidFromStringA(ObjToString(objP), uuidP);
        if (status != RPC_S_OK) {
            /* Try as GUID form */
            return ObjToGUID(interp, objP, uuidP);
        }
    } else {
        ZeroMemory(uuidP, sizeof(*uuidP));
    }
    return TCL_OK;
}

int ObjToUUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, UUID **uuidPP)
{
    if (ObjCharLength(objP) == 0) {
        *uuidPP = NULL;
        return TCL_OK;
    } else 
        return ObjToUUID(interp, objP, *uuidPP);
}

Tcl_Obj *ObjFromLSA_UNICODE_STRING(const LSA_UNICODE_STRING *lsauniP)
{
    /* Note LSA_UNICODE_STRING Length field is in *bytes* NOT characters */
    return ObjFromUnicodeN(lsauniP->Buffer, lsauniP->Length / sizeof(WCHAR));
}

void ObjToLSA_UNICODE_STRING(Tcl_Obj *objP, LSA_UNICODE_STRING *lsauniP)
{
    int nchars;
    lsauniP->Buffer = ObjToUnicodeN(objP, &nchars);
    lsauniP->Length = (USHORT) (sizeof(WCHAR)*nchars); /* in bytes */
    lsauniP->MaximumLength = lsauniP->Length;
}


/* interp may be NULL */
TCL_RESULT ObjFromSID (Tcl_Interp *interp, SID *sidP, Tcl_Obj **objPP)
{
    char *strP;

    if (ConvertSidToStringSidA(sidP, &strP) == 0) {
        if (interp)
            TwapiReturnSystemError(interp);
        return TCL_ERROR;
    }

    *objPP = ObjFromString(strP);
    LocalFree(strP);
    return TCL_OK;
}

/* Like ObjFromSID but returns empty object on error */
Tcl_Obj *ObjFromSIDNoFail(SID *sidP)
{
    Tcl_Obj *objP;
    if (sidP == NULL || ObjFromSID(NULL, sidP, &objP) != TCL_OK)
        return ObjFromEmptyString();
    else
        return objP;
}


/*
 * Convert a Tcl list to a "MULTI_SZ" list of Unicode strings, terminated
 * with two nulls.  Pointer to dynamically alloced multi_sz is stored
 * in *multiszPtrPtr on success. If lifoP is NULL, the allocated memory
 * must be freed using TwapiFree. If lifoP is not NULL, it is allocated
 * from that pool and is to be freed as appropriate. Note in this latter
 * case, memory may or may not have been allocated from pool even for errors.
 */
TCL_RESULT ObjToMultiSzEx (
     Tcl_Interp *interp,
     Tcl_Obj    *listPtr,
     LPCWSTR     *multiszPtrPtr,
     MemLifo    *lifoP
    )
{
    int       i;
    int       len;
    Tcl_Obj  *objPtr;
    LPWSTR    buf;
    LPWSTR    dst;
    LPCWSTR   src;

    *multiszPtrPtr = NULL;
    for (i=0, len=0; ; ++i) {
        if (ObjListIndex(interp, listPtr, i, &objPtr) == TCL_ERROR)
            return TCL_ERROR;
        if (objPtr == NULL)
            break;              /* No more items */
        len += ObjCharLength(objPtr) + 1;
    }

    ++len;                      /* One extra null char at the end */
    if (lifoP)
        buf = MemLifoAlloc(lifoP, len*sizeof(*buf), NULL);
    else
        buf = TwapiAlloc(len*sizeof(*buf));

    for (i=0, dst=buf; ; ++i) {
        if (ObjListIndex(interp, listPtr, i, &objPtr) == TCL_ERROR) {
            if (lifoP == NULL)
                TwapiFree(buf);
            return TCL_ERROR;
        }
        if (objPtr == NULL)
            break;              /* No more items */
        src = ObjToUnicodeN(objPtr, &len);
        if (src) {
            ++len;               /* Include the terminating null */
            CopyMemory(dst, src, len*sizeof(*src));
            dst += len;
        }
    }

    /* Add the final terminating null */
    *dst = 0;

    *multiszPtrPtr = buf;
    return TCL_OK;
}

/* This wrapper needed for ObjToMultiSzEx so it can be used from GETVAR */
TCL_RESULT ObjToMultiSz (
     Tcl_Interp *interp,
     Tcl_Obj    *listPtr,
     LPCWSTR     *multiszPtrPtr
    )
{
    return ObjToMultiSzEx(interp, listPtr, multiszPtrPtr, NULL);
}

/*
 * Convert a "MULTI_SZ" list of Unicode strings, terminated with two nulls to
 * a Tcl list. For example, a list of three strings - "abc", "def" and
 * "hij" would look like 'abc\0def\0hij\0\0'. This function will create
 * a list Tcl_Obj and return it. Will return NULL on error.
 *
 * maxlen is provided because registry data can be badly formatted
 * by applications. So we optionally ensure we do not read beyond
 * maxlen characters. This also lets it be used from EvtFormatMessage
 * code where termination is determined by length, not a second \0.
 */
Tcl_Obj *ObjFromMultiSz(LPCWSTR lpcw, int maxlen)
{
    Tcl_Obj *listPtr = ObjNewList(0, NULL);
    LPCWSTR start = lpcw;

    if (lpcw == NULL || maxlen == 0)
        return listPtr;

    if (maxlen == -1)
        maxlen = INT_MAX;

    while ((lpcw - start) < maxlen && *lpcw) {
        LPCWSTR s;
        /* Locate end of this string */
        s = lpcw;
        while ((lpcw - start) < maxlen && *lpcw)
            ++lpcw;
        if (s == lpcw) {
            /* Zero-length string - end of multi-sz */
            break;
        }

        ObjAppendElement(NULL, listPtr, ObjFromUnicodeN(s, (int) (lpcw-s)));
        ++lpcw;            /* Point beyond this string, possibly beyond end */
    }

    return listPtr;
}

TCL_RESULT ObjToUCHAR(Tcl_Interp *interp, Tcl_Obj *obj, UCHAR *ucP)
{
    int lval;

    TWAPI_ASSERT(sizeof(UCHAR) == sizeof(unsigned char));
    if (ObjToRangedInt(interp, obj, 0, UCHAR_MAX, &lval) != TCL_OK)
        return TCL_ERROR;
    *ucP = (UCHAR) lval;
    return TCL_OK;
}

TCL_RESULT ObjToCHAR(Tcl_Interp *interp, Tcl_Obj *obj, CHAR *cP)
{
    int lval;

    TWAPI_ASSERT(sizeof(CHAR) == sizeof(char));
    if (ObjToRangedInt(interp, obj, SCHAR_MIN, SCHAR_MAX, &lval) != TCL_OK)
        return TCL_ERROR;
    *cP = (CHAR) lval;
    return TCL_OK;
}

TCL_RESULT ObjToUSHORT(Tcl_Interp *interp, Tcl_Obj *obj, WORD *wordP)
{
    int lval;

    TWAPI_ASSERT(sizeof(WORD) == sizeof(unsignd short));
    if (ObjToRangedInt(interp, obj, 0, USHRT_MAX, &lval) != TCL_OK)
        return TCL_ERROR;
    *wordP = (WORD) lval;
    return TCL_OK;
}

TCL_RESULT ObjToSHORT(Tcl_Interp *interp, Tcl_Obj *obj, SHORT *shortP)
{
    int lval;
    TWAPI_ASSERT(sizeof(WORD) == sizeof(short));
    if (ObjToRangedInt(interp, obj, SHRT_MIN, SHRT_MAX, &lval) != TCL_OK)
        return TCL_ERROR;
    *shortP = (SHORT) lval;
    return TCL_OK;
}

int ObjToRECT (Tcl_Interp *interp, Tcl_Obj *obj, RECT *rectP)
{
    Tcl_Obj **objv;
    int       objc;

    if (ObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 4) {
        ObjSetStaticResult(interp, "Invalid RECT format.");
        return TCL_ERROR;
    }
    if ((ObjToLong(interp, objv[0], &rectP->left) != TCL_OK) ||
        (ObjToLong(interp, objv[1], &rectP->top) != TCL_OK) ||
        (ObjToLong(interp, objv[2], &rectP->right) != TCL_OK) ||
        (ObjToLong(interp, objv[3], &rectP->bottom) != TCL_OK)) {
        return TCL_ERROR;
    }

    return TCL_OK;
}

/* *rectPP must be  VALID memory ! */
int ObjToRECT_NULL(Tcl_Interp *interp, Tcl_Obj *obj, RECT **rectPP)
{
    int len;
    if (ObjListLength(interp, obj, &len) != TCL_OK)
        return TCL_ERROR;
    if (len == 0) {
        *rectPP = NULL;
        return TCL_OK;
    } else
        return ObjToRECT(interp, obj, *rectPP);
}


/* Return a Tcl Obj from a RECT structure */
Tcl_Obj *ObjFromRECT(RECT *rectP)
{
    Tcl_Obj *objv[4];

    objv[0] = ObjFromLong(rectP->left);
    objv[1] = ObjFromLong(rectP->top);
    objv[2] = ObjFromLong(rectP->right);
    objv[3] = ObjFromLong(rectP->bottom);
    return ObjNewList(4, objv);
}

/* Return a Tcl Obj from a POINT structure */
Tcl_Obj *ObjFromPOINT(POINT *ptP)
{
    Tcl_Obj *objv[2];

    objv[0] = ObjFromLong(ptP->x);
    objv[1] = ObjFromLong(ptP->y);
    return ObjNewList(2, objv);
}

/* Convert a Tcl_Obj to a POINT */
int ObjToPOINT (Tcl_Interp *interp, Tcl_Obj *obj, POINT *ptP)
{
    Tcl_Obj **objv;
    int       objc;

    if (ObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 2) {
        ObjSetStaticResult(interp, "Invalid POINT format.");
        return TCL_ERROR;
    }
    if ((ObjToLong(interp, objv[0], &ptP->x) != TCL_OK) ||
        (ObjToLong(interp, objv[1], &ptP->y) != TCL_OK)) {
        return TCL_ERROR;
    }

    return TCL_OK;
}

/* Return a Tcl Obj from a POINT structure */
Tcl_Obj *ObjFromPOINTS(POINTS *ptP)
{
    Tcl_Obj *objv[2];

    objv[0] = ObjFromInt(ptP->x);
    objv[1] = ObjFromInt(ptP->y);

    return ObjNewList(2, objv);
}


Tcl_Obj *ObjFromLUID (const LUID *luidP)
{
    return Tcl_ObjPrintf("%.8x-%.8x", luidP->HighPart, luidP->LowPart);
}

/*
 * Convert a string LUID to a LUID structure. Returns luidP on success,
 * else NULL on failure (invalid string format). interp may be NULL
 */
int ObjToLUID(Tcl_Interp *interp, Tcl_Obj *objP, LUID *luidP)
{
    char *markerP;
    int   len;
    char *strP = ObjToStringN(objP, &len);

    /* Format must be "XXXXXXXX-XXXXXXXX" */
    if ((len == 17) && (strP[8] == '-')) {
        luidP->HighPart = strtoul(strP, &markerP, 16);
        if (markerP == (strP+8)) {
            luidP->LowPart = strtoul(&strP[9], &markerP, 16);
            if (markerP == (strP+17))
                return TCL_OK;
        }
    }
    if (interp) {
        ObjSetStaticResult(interp, "Invalid LUID format: ");
        Tcl_AppendResult(interp, strP, NULL);
    }
    return TCL_ERROR;
}

/* *luidP MUST POINT TO A LUID STRUCTURE. This function will write the
 *  LUID there.   However if the obj is empty, it will store NULL in *luidP
*/
int ObjToLUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, LUID **luidPP)
{
    if (ObjCharLength(objP) == 0) {
        *luidPP = NULL;
        return TCL_OK;
    } else
        return ObjToLUID(interp, objP, *luidPP);
}


Tcl_Obj *ObjFromRegValue(Tcl_Interp *interp, int regtype,
                         BYTE *bufP, int count)
{
    Tcl_Obj *objv[2];
    char *typestr = NULL;

    switch (regtype) {
    case REG_LINK:
        typestr = "link";
        // FALLTHRU
    case REG_SZ:
        if (typestr == NULL)
            typestr = "sz";
        // FALLTHRU
    case REG_EXPAND_SZ:
        if (typestr == NULL)
            typestr = "expand_sz";
        /*
         * As per MS docs, may not always be null terminated.
         * If it is, we need to strip off the null.
         */
        count /= 2;             /*  Assumed to be Unicode. */
        if (count && bufP[count-1] == 0)
            --count;        /* Do not include \0 */
        objv[1] = ObjFromUnicodeN((WCHAR *)bufP, count);
        break;
            
    case REG_DWORD_BIG_ENDIAN:
        /* Since we are returning *typed* values, do not byte swap */
        /* FALLTHRU */
    case REG_DWORD:
        if (count != 4)
            goto badformat;
        typestr = regtype == REG_DWORD ? "dword" : "dword_be";
        objv[1] = ObjFromLong(*(int *)bufP);
        break;

    case REG_QWORD:
        if (count != 8)
            goto badformat;
        typestr = regtype == REG_QWORD ? "qword" : "qword_be";
        objv[1] = ObjFromWideInt(*(Tcl_WideInt *)bufP);
        break;

    case REG_MULTI_SZ:
        typestr = "multi_sz";
        objv[1] = ObjFromMultiSz((LPCWSTR) bufP, count/2);
        break;

    case REG_BINARY:
        typestr = "binary";
        // FALLTHRU
    default:
        objv[1] = ObjFromByteArray(bufP, count);
        break;
    }

    if (typestr)
        objv[0] = ObjFromString(typestr);
    else
        objv[0] = ObjFromLong(regtype);
    return ObjNewList(2, objv);

badformat:
    ObjSetStaticResult(interp, "Badly formatted registry value");
    return NULL;
}


TCL_RESULT ObjToArgvA(Tcl_Interp *interp, Tcl_Obj *objP, char **argv, int argc, int *argcP)
{
    int       objc, i;
    Tcl_Obj **objv;

    if (ObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if ((objc+1) > argc) {
        return TwapiReturnErrorEx(interp,
                                  TWAPI_INTERNAL_LIMIT,
                                  Tcl_ObjPrintf("Number of strings (%d) in list exceeds size of argument array.", objc));
    }

    for (i = 0; i < objc; ++i)
        argv[i] = ObjToString(objv[i]);
    argv[i] = NULL;
    *argcP = objc;
    return TCL_OK;
}

Tcl_Obj *ObjFromArgvA(int argc, char **argv)
{
    int i;
    Tcl_Obj *objP = Tcl_NewListObj(0, NULL);
    if (argv) {
        for (i=0; i < argc; ++i) {
            if (argv[i] == NULL)
                break;
            ObjAppendElement(NULL, objP, ObjFromString(argv[i]));
        }
    }
    return objP;
}

int ObjToArgvW(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *argv, int argc, int *argcP)
{
    int       objc, i;
    Tcl_Obj **objv;

    if (ObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if ((objc+1) > argc) {
        return TwapiReturnErrorEx(interp,
                                  TWAPI_INTERNAL_LIMIT,
                                  Tcl_ObjPrintf("Number of strings (%d) in list exceeds size of argument array (%d).", objc, argc-1));
    }

    for (i = 0; i < objc; ++i)
        argv[i] = ObjToUnicode(objv[i]);
    argv[i] = NULL;
    *argcP = objc;
    return TCL_OK;
}


LPWSTR *ObjToMemLifoArgvW(TwapiInterpContext *ticP, Tcl_Obj *objP, int *argcP)
{
    int       j, objc;
    Tcl_Obj **objv;
    LPWSTR *argv;

    if (ObjGetElements(ticP->interp, objP, &objc, &objv) != TCL_OK)
        return NULL;

    argv = MemLifoAlloc(ticP->memlifoP, (objc+1) * sizeof(LPCWSTR), NULL);
    for (j = 0; j < objc; ++j) {
        WCHAR *s;
        int slen;
        s = ObjToUnicodeN(objv[j], &slen);
        argv[j] = MemLifoCopy(ticP->memlifoP, s, sizeof(WCHAR)*(slen+1));
    }
    argv[j] = NULL;
    *argcP = objc;
    return argv;
}


Tcl_Obj *ObjFromOpaque(void *pv, char *name)
{
    Tcl_Obj *objP;

    objP = Tcl_NewObj();
    Tcl_InvalidateStringRep(objP);
    OPAQUE_REP_VALUE(objP) = pv;
    if (name) {
        OPAQUE_REP_CTYPE(objP) = ObjFromString(name);
        ObjIncrRefs(OPAQUE_REP_CTYPE(objP));
    } else {
        OPAQUE_REP_CTYPE(objP) = NULL;
    }        
    objP->typePtr = &gOpaqueType;
    return objP;
}

TCL_RESULT ObjToOpaque(Tcl_Interp *interp, Tcl_Obj *objP, void **pvP, const char *name)
{
    char *s;

    /* Fast common case */
    if (objP->typePtr == &gOpaqueType && name == NULL) {
        *pvP = OPAQUE_REP_VALUE(objP);
        return TCL_OK;
    }

    if (objP->typePtr != &gOpaqueType) {
        if (SetOpaqueFromAny(interp, objP) != TCL_OK)
            return TCL_ERROR;
    }

    /* We need to check types only if both object type and caller specified
       type are not void */
    if (name && name[0] == 0)
        name = NULL;            /* Note we are not checking for "void*". Should we ? */
    if (name && OPAQUE_REP_CTYPE(objP)) {
        s = ObjToString(OPAQUE_REP_CTYPE(objP));
        if (!STREQ(name, s)) {
            if (interp) {
                Tcl_AppendResult(interp, "Unexpected type '", s, "', expected '",
                                 name, "'.", NULL);
            }
            return TCL_ERROR;
        }
    }

    *pvP = OPAQUE_REP_VALUE(objP);
    return TCL_OK;
}

TCL_RESULT ObjToLPVOID(Tcl_Interp *interp, Tcl_Obj *objP, HANDLE *pvP)
{
    return ObjToOpaque(interp, objP, pvP, NULL);
}

/* Converts a Tcl_Obj to a pointer of any of the specified types */
TCL_RESULT ObjToOpaqueMulti(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, int ntypes, char **types)
{
    int i;
    if (ntypes == 0 || types == NULL)
        return ObjToOpaque(interp, obj, pvP, NULL);

    for (i = 0; i < ntypes; ++i) {
        if (ObjToOpaque(interp, obj, pvP, types[i]) == TCL_OK) {
            Tcl_ResetResult(interp); /* Clean up errors from any prev type attempts */
            return TCL_OK;
        }
    }

    return TCL_ERROR;
}

TCL_RESULT ObjToVerifiedPointerOrNullTic(TwapiInterpContext *ticP, Tcl_Obj *objP, void **pvP, const char *name, void *verifier)
{
    void *pv;
    TCL_RESULT res;

    res = ObjToOpaque(ticP->interp, objP, &pv, name);
    if (res == TCL_OK && pv) {
        int error = TwapiVerifyPointerTic(ticP, pv, verifier);
        if (error != TWAPI_NO_ERROR)
            res = TwapiReturnError(ticP->interp, error);
    }
    if (res == TCL_OK)
        *pvP = pv;
    return res;
}

TCL_RESULT ObjToVerifiedPointerTic(TwapiInterpContext *ticP, Tcl_Obj *objP, void **pvP, const char *name, void *verifier)
{
    void *pv;
    TCL_RESULT res;

    res = ObjToVerifiedPointerOrNullTic(ticP, objP, &pv, name, verifier);
    if (res == TCL_OK) {
        if (pv == NULL)
            res = TwapiReturnError(ticP->interp, TWAPI_NULL_POINTER);
    }
    if (res == TCL_OK)
        *pvP = pv;
    return res;
}

TCL_RESULT ObjToVerifiedPointer(Tcl_Interp *interp, Tcl_Obj *objP, void **pvP, const char *name, void *verifier) 
{
    TwapiInterpContext *ticP = TwapiGetBaseContext(interp);
    TWAPI_ASSERT(ticP);
    return ObjToVerifiedPointerTic(ticP, objP, pvP, name, verifier);
}

TCL_RESULT ObjToVerifiedPointerOrNull(Tcl_Interp *interp, Tcl_Obj *objP, void **pvP, const char *name, void *verifier)
{
    TwapiInterpContext *ticP = TwapiGetBaseContext(interp);
    TWAPI_ASSERT(ticP);
    return ObjToVerifiedPointerOrNullTic(ticP, objP, pvP, name, verifier);
}

int ObjToIDispatch(Tcl_Interp *interp, Tcl_Obj *obj, IDispatch **dispP) 
{
    /*
     * Either IDispatchEx or IDispatch is acceptable. We try in that
     * order so error message, if any will refer to IDispatch.
     */
    if (ObjToOpaque(NULL, obj, dispP, "IDispatchEx") == TCL_OK)
        return TCL_OK;
    return ObjToOpaque(interp, obj, dispP, "IDispatch");
}

Tcl_Obj *ObjFromSYSTEM_POWER_STATUS(SYSTEM_POWER_STATUS *spsP)
{
    Tcl_Obj *objv[6];
    objv[0] = ObjFromInt(spsP->ACLineStatus);
    objv[1] = ObjFromInt(spsP->BatteryFlag);
    objv[2] = ObjFromInt(spsP->BatteryLifePercent);
    objv[3] = ObjFromInt(spsP->Reserved1);
    objv[4] = ObjFromDWORD(spsP->BatteryLifeTime);
    objv[5] = ObjFromDWORD(spsP->BatteryFullLifeTime);
    return ObjNewList(6, objv);
}


Tcl_Obj *TwapiUtf8ObjFromUnicode(CONST WCHAR *wsP, int nchars)
{
    Tcl_DString ds;
    Tcl_Obj *objP;
    int nbytes;

    /*
     * Note - not using Tcl_WinTCharToUtf because there is no way
     * of telling if the Tcl core was compiled with _UNICODE defined
     * or not
     */

    /* Note WideChar... does not like 0 length strings */
    if (wsP == NULL || nchars == 0)
        return ObjFromEmptyString();

    Tcl_DStringInit(&ds);

    nbytes = WideCharToMultiByte(
        CP_UTF8, /* CodePag */
        0,       /* dwFlags */
        wsP,     /* lpWideCharStr */
        nchars < 0 ? -1 : nchars, /* cchWideChar */
        NULL,    /* lpMultiByteStr */
        0,       /* cbMultiByte */
        NULL,    /* lpDefaultChar */
        NULL     /* lpUsedDefaultChar */
        );
    
    if (nbytes == 0) {
        Tcl_Panic("WideCharToMultiByte returned 0, with error code %d", GetLastError());
    }

    Tcl_DStringSetLength(&ds, nbytes);

    nbytes = WideCharToMultiByte(
        CP_UTF8, /* CodePag */
        0,       /* dwFlags */
        wsP,     /* lpWideCharStr */
        nchars < 0 ? -1 : nchars, /* cchWideChar */
        Tcl_DStringValue(&ds),    /* lpMultiByteStr */
        Tcl_DStringLength(&ds),   /* cbMultiByte */
        NULL,    /* lpDefaultChar */
        NULL     /* lpUsedDefaultChar */
        );
    
    if (nbytes == 0) {
        Tcl_Panic("WideCharToMultiByte returned 0, with error code %d", GetLastError());
    }

    /*
     * Note WideCharToMultiByte does not explicitly terminate with \0
     * if nchars was specifically specified
     */
    if (nchars < 0)
        --nbytes;                /* Exclude terminating \0 */
    objP = ObjFromStringN(Tcl_DStringValue(&ds), nbytes);

    Tcl_DStringFree(&ds);
    return objP;
}

Tcl_Obj *ObjFromTIME_ZONE_INFORMATION(const TIME_ZONE_INFORMATION *tzP)
{
    Tcl_Obj *objs[7];

    objs[0] = ObjFromLong(tzP->Bias);
    objs[1] = ObjFromUnicodeLimited(tzP->StandardName, ARRAYSIZE(tzP->StandardName), NULL);
    objs[2] = ObjFromSYSTEMTIME(&tzP->StandardDate);
    objs[3] = ObjFromLong(tzP->StandardBias);
    objs[4] = ObjFromUnicodeLimited(tzP->DaylightName, ARRAYSIZE(tzP->DaylightName), NULL);
    objs[5] = ObjFromSYSTEMTIME(&tzP->DaylightDate);
    objs[6] = ObjFromLong(tzP->DaylightBias);
    return ObjNewList(7, objs);
}

TCL_RESULT ObjToTIME_ZONE_INFORMATION(Tcl_Interp *interp,
                                      Tcl_Obj *tzObj,
                                      TIME_ZONE_INFORMATION *tzP)
{
    Tcl_Obj **objPP;
    int nobjs;
    Tcl_Obj *stdObj, *daylightObj;

    if (ObjGetElements(NULL, tzObj, &nobjs, &objPP) == TCL_OK &&
        TwapiGetArgs(interp, nobjs, objPP,
                     GETINT(tzP->Bias),
                     GETOBJ(stdObj), GETVAR(tzP->StandardDate, ObjToSYSTEMTIME), GETINT(tzP->StandardBias),
                     GETOBJ(daylightObj), GETVAR(tzP->DaylightDate, ObjToSYSTEMTIME), GETINT(tzP->DaylightBias),
                     ARGEND) == TCL_OK) {
        WCHAR *std_name, *daylight_name;
        int std_len, daylight_len;
        std_name = ObjToUnicodeN(stdObj, &std_len);
        daylight_name = ObjToUnicodeN(daylightObj, &daylight_len);
        /* Remember we need a terminating \0 */
        if (std_len < ARRAYSIZE(tzP->StandardName) &&
            daylight_len < ARRAYSIZE(tzP->DaylightName)) {
            wcscpy(tzP->StandardName, std_name);
            wcscpy(tzP->DaylightName, daylight_name);
            return TCL_OK;
        }
    }

    return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid time zone format");
}

static char hexmap[] = "0123456789abcdef";

Tcl_Obj *ObjFromUCHARHex(UCHAR val)
{
    char buf[4];

    // Significantly faster than Tcl_ObjPrintf

    buf[0] = '0';
    buf[1] = 'x';

    buf[2] = hexmap[(val >> 4) & 0xf];
    buf[3] = hexmap[val & 0xf];

    return ObjFromStringN(buf, 4);
}


Tcl_Obj *ObjFromUSHORTHex(USHORT val)
{
    char buf[6];

    // Significantly faster than Tcl_ObjPrintf

    buf[0] = '0';
    buf[1] = 'x';

    buf[2] = hexmap[(val >> 12) & 0xf];
    buf[3] = hexmap[(val >> 8) & 0xf];
    buf[4] = hexmap[(val >> 4) & 0xf];
    buf[5] = hexmap[val & 0xf];

    return ObjFromStringN(buf, 6);
}

Tcl_Obj *ObjFromULONGHex(ULONG val)
{
    char buf[10];

    // Significantly faster than Tcl_ObjPrintf("0x%8.8x", val);

    buf[0] = '0';
    buf[1] = 'x';

    buf[2] = hexmap[(val >> 28) & 0xf];
    buf[3] = hexmap[(val >> 24) & 0xf];
    buf[4] = hexmap[(val >> 20) & 0xf];
    buf[5] = hexmap[(val >> 16) & 0xf];
    buf[6] = hexmap[(val >> 12) & 0xf];
    buf[7] = hexmap[(val >> 8) & 0xf];
    buf[8] = hexmap[(val >> 4) & 0xf];
    buf[9] = hexmap[val & 0xf];

    return ObjFromStringN(buf, 10);
}

Tcl_Obj *ObjFromULONGLONGHex(ULONGLONG ull)
{
    Tcl_Obj *objP;

    /* Unfortunately, Tcl_Objprintf does not handle 64 bits currently */
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
    Tcl_Obj *wideObj;
    wideobj = ObjFromWideInt((Tcl_WideInt) ull);
    objP = Tcl_Format(NULL, "0x%16.16lx", 1, &wideobj);
    ObjDecrRefs(wideobj);
#else
    char buf[20];
    int i;

    //_snprintf(buf, sizeof(buf), "0x%16.16I64x", ull);
    buf[0] = '0';
    buf[1] = 'x';
    for (i = 0; i < 16; ++i) {
        buf[15-i+2] = hexmap[(ull >> (i*4)) & 0xf];
    }
#endif
    objP = ObjFromStringN(buf, 18);
    return objP;
}


Tcl_Obj *ObjFromULONGLONG(ULONGLONG ull)
{
    /*
     * Unsigned 64-bit ints with the high bit set will not fit in Tcl_WideInt.
     * We need to convert to a bignum from a hex string.
     */

    if (ull & 0x8000000000000000) {
        Tcl_Obj *objP;
        mp_int mpi;
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
        Tcl_Obj *mpobj;
        objP = ObjFromWideInt((Tcl_WideInt) ull);
        mpobj = Tcl_Format(NULL, "0x%lx", 1, &objP);
        ObjDecrRefs(objP);
        objP = mpobj;
#else
        char buf[40];
        _snprintf(buf, sizeof(buf), "%I64u", ull);
        objP = ObjFromString(buf);
#endif
        /* Force to bignum because COM interface sometimes needs to check type*/
        if (Tcl_GetBignumFromObj(NULL, objP, &mpi) == TCL_ERROR)
            return objP;
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
        Tcl_InvalidateStringRep(objP); /* So we get a decimal string rep */
#endif
        TclBN_mp_clear(&mpi);
        return objP;
    } else {
        return ObjFromWideInt((Tcl_WideInt) ull);
    }
}

/* Given a IP address as a DWORD, returns a Tcl string */
Tcl_Obj *IPAddrObjFromDWORD(DWORD addr)
{
    struct in_addr inaddr;
    inaddr.S_un.S_addr = addr;
    /* Note inet_ntoa is thread safe on Windows */
    return ObjFromString(inet_ntoa(inaddr));
}

/* Given a string, return the IP address */
int IPAddrObjToDWORD(Tcl_Interp *interp, Tcl_Obj *objP, DWORD *addrP)
{
    DWORD addr;
    char *p = ObjToString(objP);
    if ((addr = inet_addr(p)) == INADDR_NONE) {
        /* Bad format or 255.255.255.255 */
        if (! STREQ("255.255.255.255", p)) {
            if (interp) {
                Tcl_AppendResult(interp, "Invalid IP address format: ", p, NULL);
            }
            return TCL_ERROR;
        }
        /* Fine, addr contains 0xffffffff */
    }
    *addrP = addr;
    return TCL_OK;
}


/* Given a IP_ADDR_STRING list, return a Tcl_Obj */
Tcl_Obj *ObjFromIP_ADDR_STRING (
    Tcl_Interp *interp, const IP_ADDR_STRING *ipaddrstrP
)
{
    Tcl_Obj *resultObj = ObjNewList(0, NULL);
    while (ipaddrstrP) {
        Tcl_Obj *objv[3];

        if (ipaddrstrP->IpAddress.String[0]) {
            objv[0] = ObjFromString(ipaddrstrP->IpAddress.String);
            objv[1] = ObjFromString(ipaddrstrP->IpMask.String);
            objv[2] = ObjFromDWORD(ipaddrstrP->Context);
            ObjAppendElement(interp, resultObj,
                                     ObjNewList(3, objv));
        }

        ipaddrstrP = ipaddrstrP->Next;
    }

    return resultObj;
}


/* Note - port is not returned - only address. Can return NULL */
Tcl_Obj *ObjFromSOCKADDR_address(SOCKADDR *saP)
{
    char buf[50];
    DWORD bufsz = ARRAYSIZE(buf);
    
    if (WSAAddressToStringA(saP,
                            ((SOCKADDR_IN6 *)saP)->sin6_family == AF_INET6 ? sizeof(SOCKADDR_IN6) : sizeof(SOCKADDR_IN),
                            NULL,
                            buf,
                            &bufsz) == 0) {
        if (bufsz && buf[bufsz-1] == 0)
            --bufsz;        /* Terminating \0 */
        return ObjFromStringN(buf, bufsz);
    }
    /* Error already set */
    return NULL;
}

/* Can return NULL on error */
Tcl_Obj *ObjFromSOCKADDR(SOCKADDR *saP)
{
    short save_port;
    Tcl_Obj *objv[2];

    /* Stash port as 0 so does not show in address string */
    if (((SOCKADDR_IN6 *)saP)->sin6_family == AF_INET6) {
        save_port = ((SOCKADDR_IN6 *)saP)->sin6_port;
        ((SOCKADDR_IN6 *)saP)->sin6_port = 0;
    } else {
        save_port = ((SOCKADDR_IN *)saP)->sin_port;
        ((SOCKADDR_IN *)saP)->sin_port = 0;
    }
    
    objv[0] = ObjFromSOCKADDR_address(saP);
    if (objv[0] == NULL)
        return NULL;

    objv[1] = ObjFromInt((WORD)(ntohs(save_port)));

    if (((SOCKADDR_IN6 *)saP)->sin6_family == AF_INET6) {
        ((SOCKADDR_IN6 *)saP)->sin6_port = save_port;
    } else {
        ((SOCKADDR_IN *)saP)->sin_port = save_port;
    }

    return ObjNewList(2, objv);
}


/* Can return NULL */
Tcl_Obj *ObjFromIPv6Addr(const char *addrP, DWORD scope_id)
{
    SOCKADDR_IN6 si;

    si.sin6_family = AF_INET6;
    si.sin6_port = 0;
    si.sin6_flowinfo = 0;
    CopyMemory(si.sin6_addr.u.Byte, addrP, 16);
    si.sin6_scope_id = scope_id;
    return ObjFromSOCKADDR_address((SOCKADDR *)&si);
}


/* Convert a VT string rep to corresponding integer */
int ObjToVT(Tcl_Interp *interp, Tcl_Obj *obj, VARTYPE *vtP)
{
    int i;
    Tcl_Obj **objv;
    int       objc;
    VARTYPE   vt;

    /* The VT may be take one of the following forms:
     *    - integer
     *    - symbol
     *    - list {ptr VT}
     *    - list {userdefined VT}
     */
    if (ObjToInt(NULL, obj, &i) == TCL_OK) {
        *vtP = (VARTYPE) i;
        return TCL_OK;
    } else if (LookupBaseVTToken(interp, ObjToString(obj), vtP) == TCL_OK) {
        return TCL_OK;
    }

    /*
     * See if it's a list. Note interp contains an error msg at this point
     * that we preserve.
     */

    if (ObjGetElements(NULL, obj, &objc, &objv) != TCL_OK ||
        objc < 2) {
        return TCL_ERROR;
    }
    if (ObjToInt(NULL, objv[0], &i) == TCL_OK) {
        vt = (VARTYPE) i;
    } else if (LookupBaseVTToken(NULL, ObjToString(objv[0]), &vt) != TCL_OK) {
        return TCL_ERROR;
    }

    /* TBD - do we not need to look at the objv[1] at all ? */
    /* vt must be either pointer, array or UDT in the list case */
    if (vt == VT_PTR || vt == VT_SAFEARRAY || vt == VT_USERDEFINED) {
        *vtP = vt;
        Tcl_ResetResult(interp); // Get rid of old error message.
        return TCL_OK;
    }
    else
        return TCL_ERROR;
}

static TCL_RESULT ObjToSAFEARRAY(Tcl_Interp *interp, Tcl_Obj *valueObj, SAFEARRAY **saPP, VARTYPE *vtP)
{
    VARTYPE vt = *vtP;
    int i, ndim, cur_dim;
    Tcl_Obj *objP;
    void *valP;
    SAFEARRAY *saP = NULL;
    SAFEARRAYBOUND bounds[TWAPI_MAX_SAFEARRAY_DIMS];
    unsigned long indices[TWAPI_MAX_SAFEARRAY_DIMS];
    Tcl_Obj *objs[TWAPI_MAX_SAFEARRAY_DIMS];
    HRESULT hr;
    int tcltype;

    TWAPI_ASSERT(vt & VT_ARRAY);
    vt &= ~ VT_ARRAY;

    /* Change VARTYPE's that are not marshallable */
    if (vt == VT_INT || vt == VT_HRESULT)
        vt = VT_I4;
    else if (vt == VT_UINT)
        vt = VT_UI4;
    *vtP = vt | VT_ARRAY;

    tcltype = TwapiGetTclType(valueObj);

    if ((vt == VT_UI1 || vt == VT_VARIANT) &&
        tcltype == TWAPI_TCLTYPE_BYTEARRAY) {
        /* Special case byte array */
        valP = ObjToByteArray(valueObj, &i);
        saP = SafeArrayCreateVector(VT_UI1, 0, i);
        if (saP == NULL)
            return TwapiReturnErrorMsg(interp, TWAPI_SYSTEM_ERROR,
                                       "Allocation of UI1 SAFEARRAY failed.");
        SafeArrayLock(saP);
        CopyMemory(saP->pvData, valP, i);
        SafeArrayUnlock(saP);
        *saPP = saP;
        return TCL_OK;
    }

    /*
     * Except for the above case, a SAFEARRAY is a nested list in Tcl.
     * First figure out the number of dimensions based on nesting level.
     * valueObj is expected to be a list and we treat it as such even
     * if its current Tcl type is not list. For nested levels, we do
     * treat it as a list only if it is actually already typed as a list.
     */
    if (ObjListIndex(interp, valueObj, 0, &objP) != TCL_OK ||
        ObjListLength(interp, valueObj, &i) != TCL_OK)
        return TCL_ERROR;       /* Top level obj must be a list */

    bounds[0].lLbound = 0;
    bounds[0].cElements = i;
    ndim = 1;

    /* Note objP may be NULL */

    while (objP) {
        /* Note we check type before calling ListObjIndex else object
           will shimmer into a list even if it is not. */
        if (TwapiGetTclType(objP) != TWAPI_TCLTYPE_LIST)
            break;

        if (ndim >= ARRAYSIZE(bounds))
            return TwapiReturnError(interp, TWAPI_INTERNAL_LIMIT);


        if (ObjListLength(interp, objP, &i) != TCL_OK)
            return TCL_ERROR;   /* Huh? Type was list so why fail? */
        bounds[ndim].lLbound = 0;
        bounds[ndim].cElements = i;

        ++ndim; /* Additional level of nesting implies one more dimension */

        if (ObjListIndex(interp, objP, 0, &objP) != TCL_OK)
            return TCL_ERROR;   /* Huh? Type was list so why fail? */

        /* Note objP can be NULL (empty list) */
    }

    saP = SafeArrayCreate(vt, ndim, bounds);
    if (bounds[0].cElements == 0) {
        /* Empty array */
        *saPP = saP;
        return TCL_OK;
    }
        
    if (saP == NULL)
        return TwapiReturnErrorEx(interp, TWAPI_SYSTEM_ERROR,
                                  Tcl_ObjPrintf("Allocation of %d-dimensional SAFEARRAY of type %d failed.", ndim, vt));
    SafeArrayLock(saP);

    /*
     * We will start from index 0,0..0 and increment in turn carrying over
     * to next dimension when a dimension's element count is reached.
     * The objs[] array will keep track of the nested list corresponding
     * to each dimension.
     */
    cur_dim = 0;
    objs[0] = valueObj;
    indices[0] = 0;

    /* TBD - is this loop needed? Is it not just repeated at top of 
       while loop below ?*/
    for (i = 1; i < ndim; ++i) {
        indices[i] = 0;
        if (ObjListIndex(interp, objs[i-1], 0, &objs[i]) != TCL_OK)
            goto error_handler;
    }

    /*
     * So now indices[] is all 0 (ie each dimension index is 0) and
     * the objs[] array contains the first Tcl_Obj at each nested level
     * (which corresponds to a dimension). We will iterate through all
     * elements incrementing indices[]. 
     */

    
    while (1) {
        int ival;
        double dval;

        /*
         * At top of the loop, cur_dim is the innermost dimension that has
         * valid entry objs[] and indices[] entries. objs[cur_dim] is
         * is the element of that dimension being processed,
         * indices[cur_dim] is the index into objs[cur_dim] that is
         * the element to be processed in the next inner dimension.
         * We have to reset all inner dimension indices to 0 (and so
         * also the corresponding objs[]. As an example, on first entry
         * to the loop, cur_dim will be 0, indices[0] is also 0 and
         * the loop below initializes all elements to be the first element
         * of the dimension. On the other hand, when the outer loop is
         * iterating over the innermost dimension (dimension ndim-1),
         * and have not exhausted
         * the number of elements therein, we will not have to reset anything
         * at all and the loop below is not entered at all.
         */
        while (++cur_dim < ndim) {
            if (ObjListIndex(interp, objs[cur_dim-1], indices[cur_dim-1], &objs[cur_dim]) != TCL_OK)
                goto error_handler;
        }


        if ((hr = SafeArrayPtrOfIndex(saP, indices, &valP)) != S_OK)
            goto system_error_handler;

        if (ObjListIndex(interp, objs[ndim-1], indices[ndim-1], &objP) != TCL_OK)
            goto error_handler;

        if (objP == NULL) {
            TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS,
                                    "Too few elements in SAFEARRAY.");
            goto error_handler;
        }

        switch (vt) {
        case VT_I1:
        case VT_UI1:
        case VT_I2:
        case VT_UI2:
        case VT_I4:
        case VT_UI4:
        case VT_INT:
        case VT_UINT:
            if (ObjToInt(interp, objP, &ival) != TCL_OK)
                goto error_handler;
            switch (vt) {
            case VT_I1:
            case VT_UI1:
                *(char *)valP = ival;
                break;
            case VT_I2:
            case VT_UI2:
                *(short *)valP = ival;
                break;
            case VT_I4:
            case VT_UI4:
            case VT_INT:
            case VT_UINT:
                *(int *)valP = ival;
                break;
            }

            break;

        case VT_R4:
        case VT_R8:
        case VT_DATE:
            if (ObjToDouble(interp, objP, &dval) != TCL_OK)
                goto error_handler;
            if (vt == VT_R4)
                *(float *) valP = (float) dval;
            else
                *(double *) valP = dval;
            break;

        case VT_BSTR:
            if (ObjToBSTR(interp, objP, valP) != TCL_OK)
                goto error_handler;
            break;

        case VT_CY:
            if (ObjToCY(interp, objP, valP) != TCL_OK)
                goto error_handler;
            break;

        case VT_BOOL:
            if (ObjToBoolean(interp, objP, &ival) != TCL_OK)
                goto error_handler;
            *(VARIANT_BOOL *)valP = ival ? VARIANT_TRUE : VARIANT_FALSE;
            break;

        case VT_DISPATCH:
            /* AddRef as it will be Release'd when safearray is freed */
            if (ObjToIDispatch(interp, objP, valP) != TCL_OK)
                goto error_handler;
            if (*(IDispatch **)valP)
                (*(IDispatch **)valP)->lpVtbl->AddRef(*(IDispatch **)valP);
            break;

        case VT_UNKNOWN:
            /* AddRef as it will be Release'd when safearray is freed */
            if (ObjToIUnknown(interp, objP, valP) != TCL_OK)
                goto error_handler;
            if (*(IUnknown **)valP)
                (*(IUnknown **)valP)->lpVtbl->AddRef(*(IUnknown **)valP);
            break;

        case VT_VARIANT:
            if (ObjToVARIANT(interp, objP, valP, VT_VARIANT) != TCL_OK)
                goto error_handler;
            else {
                VARIANT *variantP = (VARIANT *) valP;
                /* Again, IDispatch and IVariant will be released when
                   safeaarray is destroyed so addref them since we are
                   holding on to them
                */
                switch (V_VT(variantP)) {
                case VT_DISPATCH:
                case VT_UNKNOWN:
                    if (V_UNKNOWN(variantP))
                        (V_UNKNOWN(variantP))->lpVtbl->AddRef(V_UNKNOWN(variantP));
                    break;
                }
            }
            break;

        case VT_DECIMAL:
            if (ObjToDECIMAL(interp, objP, valP) != TCL_OK)
                goto error_handler;
            break;

        case VT_I8:
        case VT_UI8:
            if (ObjToWideInt(interp, objP, valP) != TCL_OK)
                goto error_handler;
            break;

        default:
            /* Dunno how to handle these */
            TwapiReturnErrorEx(interp, TWAPI_UNSUPPORTED_TYPE,
                               Tcl_ObjPrintf("Unsupported SAFEARRAY type %d", vt));
            goto error_handler;
        }

        /*
         * Now increment indices[] to point to next entry. We increment
         * the index for each dimension and if it exceeds the number of
         * elements in that dimension, we reset its index to 0 and 
         * carry over and increment the previous dimension.
         */
        for (cur_dim = ndim-1; cur_dim >= 0; --cur_dim) {
            if (++indices[cur_dim] < bounds[cur_dim].cElements)
                break;          /* No overflow for this dimension */
            indices[cur_dim] = 0;
        }

        /*
         * The above loop terminates when
         *  cur_dim < 0 - implies even the outermost dimension elements have
         *          been iterated and no more elements need to be processed.
         *  cur_dim >= 0 - implies that dimension is not done with. indices[i]
         *          contains the index into elements in that dimension
         *          that we next iterate over. We also have to reset
         *          the innermost dimension indices accordingly. This 
         *          is done at the top of the loop.
         */

        if (cur_dim < 0)
            break;              /* All done, no more elements */

    }

    SafeArrayUnlock(saP);
    *saPP = saP;
    return TCL_OK;

system_error_handler: /* hr - Win32 error */
    Twapi_AppendSystemError(interp, hr);

error_handler:
    if (saP) {
        SafeArrayUnlock(saP);
        SafeArrayDestroy(saP);
    }
    return TCL_ERROR;
}



/*
 * Returns a Tcl_Obj that is a list representing the elements in the
 * dimension dim. indices[] is an array specifying the dimension to
 * retrieve. For example, consider a 4-dimensional safearray. When 
 * this function is called with dim=2 (meaning third dimension),
 * indices[0], indices[1] will contain valid indices for the dimensions
 * 0 and 1. The returned Tcl_Obj will be a list containing the elements
 * of dimension 2. Each such element will be a Tcl_Obj that is itself
 * a list containing the elements in dimension 3.
 *
 * Returns NULL on any errors.
 *
 * IMPORTANT: saP must be SafeArrayLock'ed on entry !
 */
static Tcl_Obj *ObjFromSAFEARRAYDimension(SAFEARRAY *saP, int dim,
                                          long indices[], int indices_size)
{
    unsigned long i;
    Tcl_Obj *objP;
    Tcl_Obj *resultObj = NULL;
    VARIANT *variantP;
    unsigned long upper, lower;
    VARTYPE vt;
    
    if (indices_size < saP-> cDims)
        return NULL;            /* Not supported as exceed max dimensions */

    if (indices_size <= dim)
        return NULL;            /* Should not happen really - internal error */

    if (SafeArrayGetLBound(saP, dim+1, &lower) != S_OK ||
        SafeArrayGetUBound(saP, dim+1, &upper) != S_OK)
        return NULL;

    if (SafeArrayGetVartype(saP, &vt) != S_OK)
        return NULL;

    /* Special case - One-dim array of UI1 is treated as binary data */
    if (vt == VT_UI1 && saP->cDims == 1 && saP->pvData) {
        TWAPI_ASSERT(dim == 0);
        return ObjFromByteArray(saP->pvData, saP->rgsabound[0].cElements);
    }

    resultObj = Tcl_NewListObj(0, NULL);

    /* Loop through all elements in this dimension. */
    if (dim < (saP->cDims-1)) {

        /* This is an intermediate dimension. Recurse */
        for (i = lower; i <= upper; ++i) {
            indices[dim] = i;
            objP = ObjFromSAFEARRAYDimension(saP, dim+1, indices, indices_size);
            if (objP == NULL)
                goto error_handler;
            else
                ObjAppendElement(NULL, resultObj, objP);
        }
    } else {

        /* This is the final dimension. Loop to collect elements */
        for (i = lower; i <= upper; ++i) {
            void *valP;
            indices[dim] = i;
            if (SafeArrayPtrOfIndex(saP, indices, &valP) != S_OK)
                goto error_handler;
            objP = NULL;
            switch (vt) {
            case VT_I2: objP = ObjFromInt(*(short *)valP); break;
            case VT_INT: /* FALLTHROUGH */
            case VT_I4: objP = ObjFromLong(*(long *)valP); break;
            case VT_R4: objP = Tcl_NewDoubleObj(*(float *)valP); break;
            case VT_R8: objP = Tcl_NewDoubleObj(*(double *)valP); break;
            case VT_CY: objP = ObjFromCY((CY *) valP); break;
            case VT_DATE: objP = Tcl_NewDoubleObj(*(double *)valP); break;
            case VT_BSTR:
                objP = ObjFromUnicodeN(*(BSTR *)valP, SysStringLen(*(BSTR *)valP));
                break;
            case VT_DISPATCH:
                /* AddRef as it will be Release'd when safearray is freed */
                if (*(IDispatch **)valP)
                    (*(IDispatch **)valP)->lpVtbl->AddRef(*(IDispatch **)valP);
                objP = ObjFromIDispatch(*(IDispatch **)valP);
                break;
            case VT_ERROR: objP = ObjFromInt(*(SCODE *)valP); break;
            case VT_BOOL: objP = ObjFromBoolean(*(VARIANT_BOOL *)valP); break;
            case VT_VARIANT:
                variantP = (VARIANT *) valP;
                /* Again, IDispatch and IVariant will be released when
                   safeaarray is destroyed so addref them since we are
                   holding on to them
                */
                switch (V_VT(variantP)) {
                case VT_DISPATCH:
                    if (V_DISPATCH(variantP))
                        (V_DISPATCH(variantP))->lpVtbl->AddRef(V_DISPATCH(variantP));
                    break;
                case VT_UNKNOWN:
                    if (V_UNKNOWN(variantP))
                        (V_UNKNOWN(variantP))->lpVtbl->AddRef(V_UNKNOWN(variantP));
                    break;
                }
                objP = ObjFromVARIANT(variantP, 0);
                break;
            case VT_DECIMAL: objP = ObjFromDECIMAL((DECIMAL *)valP); break;
            case VT_UNKNOWN:
                /* AddRef as it will be Release'd when safearray is freed */
                if (*(IUnknown **)valP)
                    (*(IUnknown **)valP)->lpVtbl->AddRef(*(IUnknown **)valP);
                objP = ObjFromIDispatch(*(IUnknown **)valP);
                break;

            case VT_I1: objP = ObjFromInt(*(char *)valP); break;
            case VT_UI1: objP = ObjFromInt(*(unsigned char *)valP); break;
            case VT_UI2: objP = ObjFromInt(*(unsigned short *)valP); break;
            case VT_UINT: /* FALLTHROUGH */
            case VT_UI4: objP = ObjFromDWORD(*(DWORD *)valP); break;
            case VT_I8: /* FALLTHRU */
            case VT_UI8: objP = ObjFromWideInt(*(__int64 *)valP); break;

                /* Dunno how to handle these */
            default:
                goto error_handler;
            }
            if (objP == NULL)
                goto error_handler;
            else
                ObjAppendElement(NULL, resultObj, objP);
        }
    }

    return resultObj;

error_handler:
    if (resultObj)
        ObjDecrRefs(resultObj);
    return NULL;
}


/*
 * Return a Tcl object that is a list
 * {VT_xxx dimensionlist valuelist}.
 * dimensionlist is a flat list of lowbound, upperbound pairs, one
 * for each dimension.
 * If VT_xxx is not recognized, valuelist is missing
 * If there is no vartype information, VT_XXX is also missing
 *
 * Returns NULL on error
 */
static Tcl_Obj *ObjFromSAFEARRAY(SAFEARRAY *saP, int value_only)
{
    Tcl_Obj *objv[2];           /* dimensions,  value */
    long     i;
    VARTYPE  vt;
    long indices[TWAPI_MAX_SAFEARRAY_DIMS];

    /* We require the safearray to have a type associated */
    if (saP == NULL || SafeArrayGetVartype(saP, &vt) != S_OK) {
        return NULL;
    }

    if (SafeArrayLock(saP) != S_OK)
        return NULL;
        
    objv[1] = ObjFromSAFEARRAYDimension(saP, 0, indices, ARRAYSIZE(indices));
    if (objv[1] == NULL || value_only) {
        SafeArrayUnlock(saP);
        return objv[1];          /* May be NULL */
    }

    /* List of dimensions */
    objv[0] = ObjNewList(0, NULL);
    for (i = 0; i < saP->cDims; ++i) {
        ObjAppendElement(NULL, objv[0], ObjFromLong(saP->rgsabound[i].lLbound));
        ObjAppendElement(NULL, objv[0], ObjFromLong(saP->rgsabound[i].cElements));
    }

    SafeArrayUnlock(saP);
    return ObjNewList(2, objv);
}

VARTYPE ObjTypeToVT(Tcl_Obj *objP)
{
    char *s;
    VARTYPE vt;
    Tcl_Obj *elemObj;

    switch (TwapiGetTclType(objP)) {
    case TWAPI_TCLTYPE_BOOLEAN: /* Fallthru */
    case TWAPI_TCLTYPE_BOOLEANSTRING:
        return VT_BOOL;
    case TWAPI_TCLTYPE_INT:
        return VT_I4;
    case TWAPI_TCLTYPE_WIDEINT:
        return VT_I8;
    case TWAPI_TCLTYPE_DOUBLE:
        return VT_R8;
    case TWAPI_TCLTYPE_BYTEARRAY:
        return VT_UI1 | VT_ARRAY;
    case TWAPI_TCLTYPE_LIST:
        /*
         * A list is usually a SAFEARRAY. However, it could be
         * an IDispatch or IUnknown in certain special cases.
         */
#if 0 // Commented out because ObjToOpaque (via ObjToIDispatch) will shimmer type
        // of first element even when it is not a Opaque pointer. We thus lose
        // the typing info needed below
        if (ObjListLength(NULL, objP, &i) == TCL_OK && i == 2) {
            /* Possibly IUnknown or IDispatch */
            if (ObjToIDispatch(NULL, objP, &pv) == TCL_OK)
                return VT_DISPATCH;
            else if (ObjToIUnknown(NULL, objP, &pv) == TCL_OK)
                return VT_UNKNOWN;
        }
#endif        

        /* We do not know the type of each SAFEARRAY element. Guess on element value */
        vt = VT_VARIANT;   /* In case we cannot tell */
        if (ObjListIndex(NULL, objP, 0, &elemObj) == TCL_OK && elemObj) {
            switch (TwapiGetTclType(elemObj)) {
            case TWAPI_TCLTYPE_BOOLEAN: /* Fallthru */
            case TWAPI_TCLTYPE_BOOLEANSTRING: vt = VT_BOOL; break;
            case TWAPI_TCLTYPE_INT:           vt = VT_I4; break;
            case TWAPI_TCLTYPE_WIDEINT:       vt = VT_I8; break;
            case TWAPI_TCLTYPE_DOUBLE:        vt = VT_R8; break;
            case TWAPI_TCLTYPE_STRING:        vt = VT_BSTR; break;
            }            
        }
        return vt | VT_ARRAY;

    case TWAPI_TCLTYPE_DICT:
        /* Something that is constructed like a dictionary cannot
           really be a numeric or boolean. Since there is no
           dictionary type in COM, pass as a string.
        */
        return VT_BSTR;

    case TWAPI_TCLTYPE_OPAQUE:
        TWAPI_ASSERT(objP->typePtr == &gOpaqueType);
        if (OPAQUE_REP_CTYPE(objP)) {
            s = ObjToString(OPAQUE_REP_CTYPE(objP));
            if (STREQ(s, "IDispatch"))
                return VT_DISPATCH;
            if (STREQ(s, "IUnknown"))
                return VT_UNKNOWN;
        }
        return VT_VARIANT;
    case TWAPI_TCLTYPE_VARIANT:
        TWAPI_ASSERT(objP->typePtr == &gVariantType);
        return (VARTYPE) VARIANT_REP_VT(objP);
    case TWAPI_TCLTYPE_STRING:
        /* In Tcl everything is a string, including numerics. However
           a command such as [set x 123] will result in x being set
           as a "pure" string, ie. TWAPI_TCLTYPE_NONE. So we mark
           it as a BSTR if it is an explicit string type. - TBD ?
        */
        return VT_BSTR;
    case TWAPI_TCLTYPE_NONE: /* Completely untyped */
    default:
        /* Caller has to use value-based heuristics to determine type */
        return VT_VARIANT;
    }
}


TCL_RESULT ObjToVARIANT(Tcl_Interp *interp, Tcl_Obj *objP, VARIANT *varP, VARTYPE vt)
{
    HRESULT hr;
    long lval;

    if (vt & VT_ARRAY) {
        if (ObjToSAFEARRAY(interp, objP, &varP->parray, &vt) != TCL_OK)
            return TCL_ERROR;
        V_VT(varP) = vt;
        return TCL_OK;
    }

    varP->vt = vt;
    switch (vt) {
    case VT_EMPTY:
    case VT_NULL:
        break;
    case VT_I2:   return ObjToSHORT(interp, objP, &V_I2(varP));
    case VT_UI2:  return ObjToUSHORT(interp, objP, &V_UI2(varP));
    case VT_I1:   return ObjToCHAR(interp, objP, &V_I1(varP));
    case VT_UI1:   return ObjToUCHAR(interp, objP, &V_UI1(varP));

    /* For compatibility reasons, allow interchangeable signed/unsigned ints */
    case VT_I4:
    case VT_UI4:
    case VT_INT:
    case VT_UINT:
    case VT_HRESULT:
        if (ObjToLong(interp, objP, &lval) != TCL_OK)
            return TCL_ERROR;
        switch (vt) {
        case VT_I4:
        case VT_INT:
        case VT_HRESULT:
            vt = VT_I4;        /* VT_INT and VT_HRESULT cannot be marshalled */
            V_I4(varP) = lval;
            break;
        case VT_UI4:
        case VT_UINT:
            vt = VT_UI4;        /* VT_UINT cannot be marshalled */
            V_UI4(varP) = lval;
            break;
        }

        varP->vt = vt;
        break;

    case VT_R4:
    case VT_R8:
        if (ObjToDouble(interp, objP, &varP->dblVal) != TCL_OK)
            return TCL_ERROR;
        varP->vt = VT_R8;
        if (vt == VT_R4) {
            hr = VariantChangeType(varP, varP, 0, VT_R4);
            if (FAILED(hr)) {
                Twapi_AppendSystemError(interp, hr);
                return TCL_ERROR;
            }
        }
        break;

    case VT_CY:
        if (ObjToCY(interp, objP, & V_CY(varP)) != TCL_OK)
            return TCL_ERROR;
        break;

    case VT_DATE:
        if (ObjToDouble(interp, objP, & V_DATE(varP)) != TCL_OK)
            return TCL_ERROR;
        break;

    case VT_VARIANT:
        /* Value is VARIANT so we don't really know the type.
         * Note VT_VARIANT is only valid in type descriptions and
         * is not valid for VARIANTARG (ie. for
         * the actual value). It has to be a concrete type.
         * We have been unable to guess the type based on the Tcl
         * internal type pointer so make a guess based on value.
         * Note we pass NULL for interp
         * when to GetLong and GetDouble as we don't want an
         * error message left in the interp.
         *
         * We only base our logic here on values. Any type information
         * from Tcl_Obj.typePtr should have been considered by the caller.
         */
        if (ObjToLong(NULL, objP, &varP->lVal) == TCL_OK) {
            varP->vt = VT_I4;
        } else if (ObjToDouble(NULL, objP, &varP->dblVal) == TCL_OK) {
            varP->vt = VT_R8;
        } else if (ObjToIDispatch(NULL, objP, &varP->pdispVal) == TCL_OK) {
            varP->vt = VT_DISPATCH;
        } else if (ObjToIUnknown(NULL, objP, &varP->punkVal) == TCL_OK) {
            varP->vt = VT_UNKNOWN;
#if 0
        } else if (ObjCharLength(objP) == 0) {
            varP->vt = VT_NULL;
#endif
        } else {
            /* Cannot guess type, just pass as a BSTR */
            if (ObjToBSTR(interp, objP, &varP->bstrVal) != TCL_OK)
                return TCL_ERROR;
            varP->vt = VT_BSTR;
        }

        break;

    case VT_BSTR:
        return ObjToBSTR(interp, objP, &varP->bstrVal);


    case VT_DISPATCH:
        return ObjToIDispatch(interp, objP, (void **)&varP->pdispVal);

    case VT_VOID: /* FALLTHRU */
    case VT_ERROR:
        /* Treat as optional argument */
        varP->vt = VT_ERROR;
        varP->scode = DISP_E_PARAMNOTFOUND;
        break;

    case VT_BOOL:
        if (ObjToBoolean(interp, objP, &varP->intVal) != TCL_OK)
            return TCL_ERROR;
        varP->boolVal = varP->intVal ? VARIANT_TRUE : VARIANT_FALSE;
        varP->vt = VT_BOOL;
        break;

    case VT_UNKNOWN:
        return ObjToIUnknown(interp, objP, (void **) &varP->punkVal);

    case VT_DECIMAL:
        return ObjToDECIMAL(interp, objP, & V_DECIMAL(varP));

    case VT_I8:
    case VT_UI8:
        varP->vt = VT_I8;
        return ObjToWideInt(interp, objP, &varP->llVal);

    default:
        ObjSetResult(interp,
                          Tcl_ObjPrintf("Invalid or unsupported VARTYPE (%d)",
                                        vt));
        return TCL_ERROR;
    }

    return TCL_OK;

}


/* 
 * If value_only is 0, returns a Tcl object that is a list {VT_xxx value}.
 * If VT_xxx is not known, value is missing (only the VT_xxx is
 * returned). 
 * If value_only is 1, returns only the value object and am empty object
 * if VT_xxx is not known.
 * Never returns NULL.
 */
Tcl_Obj *ObjFromVARIANT(VARIANT *varP, int value_only)
{
    Tcl_Obj *objv[2];
    Tcl_Obj *valObj[2];
    unsigned long ulval;
    VARIANT  empty;
    void    *recdataP;
    IDispatch *idispP;
    IUnknown  *iunkP;

    if (varP == NULL) {
        VariantInit(&empty);
        varP = &empty;
    }

    if (V_VT(varP) & VT_ARRAY) {
        if (V_VT(varP) & VT_BYREF)
            objv[1] = ObjFromSAFEARRAY(*(varP->pparray), value_only);
        else
            objv[1] = ObjFromSAFEARRAY(varP->parray, value_only);

        if (objv[1] && value_only)
            return objv[1];

        objv[0] = ObjFromInt(V_VT(varP) & ~ VT_BYREF);
        return ObjNewList(objv[1] ? 2 : 1, objv);

    }

    if ((V_VT(varP) == (VT_BYREF|VT_VARIANT)) && varP->pvarVal)
        return ObjFromVARIANT(varP->pvarVal, value_only);

    objv[0] = NULL;
    objv[1] = NULL;

    switch (V_VT(varP)) {
    case VT_EMPTY|VT_BYREF:
    case VT_EMPTY:
    case VT_NULL|VT_BYREF:
    case VT_NULL:
        break;

    case VT_I2|VT_BYREF:
    case VT_I2:
        objv[1] = ObjFromInt(V_VT(varP) == VT_I2 ? V_I2(varP) : * V_I2REF(varP));
        break;

    case VT_I4|VT_BYREF:
    case VT_I4:
        objv[1] = ObjFromInt(V_VT(varP) == VT_I4 ? V_I4(varP) : * V_I4REF(varP));
        break;

    case VT_R4|VT_BYREF:
    case VT_R4:
        objv[1] = Tcl_NewDoubleObj(V_VT(varP) == VT_R4 ? V_R4(varP) : * V_R4REF(varP));
        break;

    case VT_R8|VT_BYREF:
    case VT_R8:
        objv[1] = Tcl_NewDoubleObj(V_VT(varP) == VT_R8 ? V_R8(varP) : * V_R8REF(varP));
        break;

    case VT_CY|VT_BYREF:
    case VT_CY:
        objv[1] = ObjFromCY(
            V_VT(varP) == VT_CY ? & V_CY(varP) : V_CYREF(varP)
            );
        break;

    case VT_BSTR|VT_BYREF:
    case VT_BSTR:
        if (V_VT(varP) == VT_BSTR)
            objv[1] = ObjFromUnicodeN(V_BSTR(varP),
                                        SysStringLen(V_BSTR(varP)));
        else
            objv[1] = ObjFromUnicodeN(* V_BSTRREF(varP),
                                        SysStringLen(* V_BSTRREF(varP)));
        break;

    case VT_DISPATCH|VT_BYREF:
        /* If VT_BYREF is set, then a reference to an existing
         * IUnknown is being returned. In this case, at the script level
         * we should not Release it but there is no way for the script
         * to know that. We therefore do a AddRef on the pointer here
         * so it can be released later in the script (ie. the script
         * can treat VT_DISPATCH and VT_DISPATCH|VT_BYREF the same
         * TBD - revisit this as to whether Release is required
         */
        idispP = * (V_DISPATCHREF(varP));
        idispP->lpVtbl->AddRef(idispP);
        objv[1] = ObjFromIDispatch(idispP);
        break;

    case VT_DISPATCH:
        idispP = V_DISPATCH(varP);
        objv[1] = ObjFromIDispatch(idispP);
        break;

    case VT_ERROR|VT_BYREF:
    case VT_ERROR:
        objv[1] = ObjFromInt(V_VT(varP) == VT_ERROR ? V_ERROR(varP) : * V_ERRORREF(varP));
        break;

    case VT_BOOL|VT_BYREF:
    case VT_BOOL:
        objv[1] = ObjFromBoolean(V_VT(varP) == VT_BOOL ? V_BOOL(varP) : * V_BOOLREF(varP));
        break;

    case VT_DATE|VT_BYREF:
    case VT_DATE:
        objv[1] = Tcl_NewDoubleObj(V_VT(varP) == VT_DATE ? V_DATE(varP) : * V_DATEREF(varP));
        break;

    case VT_VARIANT|VT_BYREF:
        /* This is for the case where varP->pvarVal is NULL. The non-NULL
           case was already handled in an if stmt above */
        break;

    case VT_DECIMAL|VT_BYREF:
    case VT_DECIMAL:
        objv[1] = ObjFromDECIMAL(
            V_VT(varP) == VT_DECIMAL ? & V_DECIMAL(varP) : V_DECIMALREF(varP)
            );
        break;


    case VT_UNKNOWN|VT_BYREF:
        /* If VT_BYREF is set, then a reference to an existing
         * IUnknown is being returned. In this case, at the script level
         * we should not Release it but there is no way for the script
         * to know that. We therefore do a AddRef on the pointer here
         * so it can be released later
         * TBD - revisit this as to whether Release is required
         */
        iunkP = * (V_UNKNOWNREF(varP));
        iunkP->lpVtbl->AddRef(iunkP);
        objv[1] = ObjFromIUnknown(iunkP);
        break;

    case VT_UNKNOWN:
        iunkP = V_UNKNOWN(varP);
        objv[1] = ObjFromIUnknown(iunkP);
        break;

    case VT_I1|VT_BYREF:
    case VT_I1:
        objv[1] = ObjFromInt(V_VT(varP) == VT_I1 ? V_I1(varP) : * V_I1REF(varP));
        break;

    case VT_UI1|VT_BYREF:
    case VT_UI1:
        objv[1] = ObjFromInt(V_VT(varP) == VT_UI1 ? V_UI1(varP) : * V_UI1REF(varP));
        break;

    case VT_UI2|VT_BYREF:
    case VT_UI2:
        objv[1] = ObjFromInt(V_VT(varP) == VT_UI2 ? V_UI2(varP) : * V_UI2REF(varP));
        break;

    case VT_UI4|VT_BYREF:
    case VT_UI4:
        /* store as wide integer if it does not fit in signed 32 bits */
        ulval = V_VT(varP) == VT_UI4 ? V_UI4(varP) : * V_UI4REF(varP);
        objv[1] = ObjFromDWORD(ulval);
        break;

    case VT_I8|VT_BYREF:
    case VT_I8:
        objv[1] = ObjFromWideInt(V_VT(varP) == VT_I8 ? V_I8(varP) : * V_I8REF(varP));
        break;

    case VT_UI8|VT_BYREF:
    case VT_UI8:
        objv[1] = ObjFromWideInt(V_VT(varP) == VT_UI8 ? V_UI8(varP) : * V_UI8REF(varP));
        break;


    case VT_INT|VT_BYREF:
    case VT_INT:
        objv[1] = ObjFromInt(V_VT(varP) == VT_INT ? V_INT(varP) : * V_INTREF(varP));
        break;

    case VT_UINT|VT_BYREF:
    case VT_UINT:
        /* store as wide integer if it does not fit in signed 32 bits */
        ulval = V_VT(varP) == VT_UINT ? V_UINT(varP) : * V_UINTREF(varP);
        if (ulval & 0x80000000) {
            objv[1] = ObjFromWideInt(ulval);
        }
        else {
            objv[1] = ObjFromLong(ulval);
        }
        break;

    case VT_RECORD:
        recdataP = NULL;
        if (V_RECORDINFO(varP) &&
            V_RECORD(varP) &&
            V_RECORDINFO(varP)->lpVtbl->RecordCreateCopy(V_RECORDINFO(varP), V_RECORD(varP), &recdataP) == S_OK
            ) {
            /*
             * Construct return value as pair of IRecordInfo* void* (data)
             */
            valObj[0] = ObjFromOpaque(V_RECORDINFO(varP), "IRecordInfo");
            // TBD - we pass pointers to record instances as void* as per
            // the IRecordInfo interface. We should change this to be
            // more typesafe
            valObj[1] = ObjFromLPVOID(recdataP);
            objv[1] = ObjNewList(2, valObj);
        }
        break;

        /* Dunno how to handle these */
    case VT_RECORD|VT_BYREF:
    case VT_VARIANT: /* Note VT_VARIANT is illegal */
    default:
        break;
    }

    if (value_only)
        return objv[1] ? objv[1] : ObjFromEmptyString();
    else {
        objv[0] = ObjFromInt(V_VT(varP) & ~VT_BYREF);
        return ObjNewList(objv[1] ? 2 : 1, objv);
    }
}


/* Returned memory in *arrayPP has to be freed by caller */
int ObjToLSASTRINGARRAY(Tcl_Interp *interp, Tcl_Obj *obj, LSA_UNICODE_STRING **arrayP, ULONG *countP)
{
    Tcl_Obj **listobjv;
    int       i, nitems, sz;
    LSA_UNICODE_STRING *ustrP;
    WCHAR    *dstP;

    if (ObjGetElements(interp, obj, &nitems, &listobjv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    /* Figure out how much space we need */
    sz = nitems * sizeof(LSA_UNICODE_STRING);
    for (i = 0; i < nitems; ++i) {
        sz += sizeof(*dstP) * (ObjCharLength(listobjv[i]) + 1);
    }

    ustrP = TwapiAlloc(sz);

    /* Figure out where the string area starts and do the construction */
    dstP = (WCHAR *) ((nitems * sizeof(LSA_UNICODE_STRING)) + (char *)(ustrP));
    for (i = 0; i < nitems; ++i) {
        WCHAR *srcP;
        int    slen;
        srcP = ObjToUnicodeN(listobjv[i], &slen);
        CopyMemory(dstP, srcP, sizeof(WCHAR)*(slen+1));
        ustrP[i].Buffer = dstP;
        ustrP[i].Length = (USHORT) (sizeof(WCHAR) * slen); /* Num *bytes*, not WCHARs */
        ustrP[i].MaximumLength = ustrP[i].Length;
        dstP += slen+1;
    }

    *arrayP = ustrP;
    *countP = nitems;

    return TCL_OK;
}




/*
 * Returns a pointer to dynamic memory containing a SID corresponding
 * to the given string representation. Returns NULL on error, and
 * sets the windows error
 */
PSID TwapiGetSidFromStringRep(char *strP)
{
    DWORD   len;
    PSID    sidP;
    PSID    local_sidP;
    int error;

    local_sidP = NULL;
    sidP = NULL;

    if (ConvertStringSidToSidA(strP, &local_sidP) == 0)
        return NULL;

    /*
     * Have a valid SID
     * Copy it into dynamic memory after validating
     */
    len = GetLengthSid(local_sidP);
    sidP = TwapiAlloc(len);
    if (! CopySid(len, sidP, local_sidP)) {
        goto errorExit;
    }

    /* Free memory allocated by ConvertStringSidToSidA */
    LocalFree(local_sidP);
    return sidP;

 errorExit:
    error = GetLastError();

    if (local_sidP) {
        LocalFree(local_sidP);
    }

    if (sidP)
        TwapiFree(sidP);

    SetLastError(error);
    return NULL;
}

/* Tcl_Obj to SID - the object may hold the SID string rep, a binary
   or a list of ints. If the object is an empty string, *sidPP is
   stored as NULL. Else the SID is dynamically allocated and a pointer to it is
   stored in *sidPP. Caller must release it by calling TwapiFree
*/
int ObjToPSID(Tcl_Interp *interp, Tcl_Obj *obj, PSID *sidPP)
{
    char *s;
    DWORD   len;
    SID  *sidP;
    DWORD winerror;

    s = ObjToStringN(obj, &len);
    if (len == 0) {
        *sidPP = NULL;
        return TCL_OK;
    }

    *sidPP = TwapiGetSidFromStringRep(ObjToString(obj));
    if (*sidPP)
        return TCL_OK;

    winerror = GetLastError();

    /* Not a string rep. See if it is a binary of the right size */
    sidP = (SID *) ObjToByteArray(obj, &len);
    if (len >= sizeof(*sidP)) {
        /* Seems big enough, validate revision and size */
        if (IsValidSid(sidP) && GetLengthSid(sidP) == len) {
            *sidPP = TwapiAlloc(len);
            /* Note SID is a variable length struct so we cannot do this
                    *(SID *) (*sidPP) = *sidP;
               (from bitter experience!)
             */
            if (CopySid(len, *sidPP, sidP))
                return TCL_OK;
            winerror = GetLastError();
        }
    }
    return Twapi_AppendSystemError(interp, winerror);
}

/* Convert a ACE object to a Tcl list. interp may be NULL */
Tcl_Obj *ObjFromACE (Tcl_Interp *interp, void *aceP)
{
    Tcl_Obj    *resultObj = NULL;
    Tcl_Obj    *obj = NULL;
    ACE_HEADER *acehdrP = &((ACCESS_ALLOWED_ACE *) aceP)->Header;
    ACCESS_ALLOWED_OBJECT_ACE *objectAceP;
    SID        *sidP;

    if (aceP == NULL) {
        if (interp)
            ObjSetStaticResult(interp, "NULL ACE pointer");
        return NULL;
    }

    resultObj = ObjNewList(0, NULL);

    /* ACE type */
    ObjAppendElement(interp, resultObj,
                             ObjFromInt(acehdrP->AceType));

    /* ACE flags */
    ObjAppendElement(interp, resultObj,
                             ObjFromInt(acehdrP->AceFlags));

    /* Now for type specific fields */
    switch (acehdrP->AceType) {
    case ACCESS_ALLOWED_ACE_TYPE:
    case ACCESS_DENIED_ACE_TYPE:
    case SYSTEM_AUDIT_ACE_TYPE:
    case SYSTEM_MANDATORY_LABEL_ACE_TYPE:
        ObjAppendElement(interp, resultObj,
                                 ObjFromDWORD(((ACCESS_ALLOWED_ACE *)aceP)->Mask));

        /* and the SID */
        obj = NULL;                /* In case of errors */
        if (ObjFromSID(interp,
                         (SID *)&((ACCESS_ALLOWED_ACE *)aceP)->SidStart,
                         &obj)
            != TCL_OK) {
            goto error_return;
        }
        ObjAppendElement(interp, resultObj, obj);
        break;

    case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
    case ACCESS_DENIED_OBJECT_ACE_TYPE:
    case SYSTEM_AUDIT_OBJECT_ACE_TYPE:
        objectAceP = (ACCESS_ALLOWED_OBJECT_ACE *)aceP;
        ObjAppendElement(interp, resultObj,
                                 ObjFromDWORD(objectAceP->Mask));
        if (objectAceP->Flags & ACE_OBJECT_TYPE_PRESENT) {
            ObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->ObjectType));
            if (objectAceP->Flags & ACE_INHERITED_OBJECT_TYPE_PRESENT) {
                ObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->InheritedObjectType));
                sidP = (SID *) &objectAceP->SidStart;
            } else {
                ObjAppendElement(interp, resultObj, ObjFromEmptyString());
                sidP = (SID *) &objectAceP->InheritedObjectType;
            }
        } else if (objectAceP->Flags & ACE_INHERITED_OBJECT_TYPE_PRESENT) {
            ObjAppendElement(interp, resultObj, ObjFromEmptyString());
            ObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->ObjectType));
            sidP = (SID *) &objectAceP->InheritedObjectType;
        } else {
            ObjAppendElement(interp, resultObj, ObjFromEmptyString());
            ObjAppendElement(interp, resultObj, ObjFromEmptyString());
            sidP = (SID *) &objectAceP->ObjectType;
        }
        obj = NULL;                /* In case of errors */
        if (ObjFromSID(interp, sidP, &obj) != TCL_OK)
            goto error_return;
        ObjAppendElement(interp, resultObj, obj);
        
        break;

    default:
        /*
         * Return a binary rep of the whole dang thing.
         * There are no pointers in there, just values so this
         * should work, I think :)
         */
        obj = ObjFromByteArray((unsigned char *) aceP, acehdrP->AceSize);

        if (ObjAppendElement(interp, resultObj, obj) != TCL_OK)
            goto error_return;

        break;
    }


    return resultObj;

 error_return:
    Twapi_FreeNewTclObj(obj); /* OK if null */
    Twapi_FreeNewTclObj(resultObj); /* OK if null */
    return NULL;
}


int ObjToACE (Tcl_Interp *interp, Tcl_Obj *aceobj, void **acePP)
{
    Tcl_Obj **objv;
    int       objc;
    int       acetype;
    int       aceflags;
    int       acesz;
    SID      *sidP;
    unsigned char *bytes;
    int            bytecount;
    ACCESS_ALLOWED_ACE    *aceP;

    *acePP = NULL;

    if (ObjGetElements(interp, aceobj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc < 2)
        goto format_error;

    if ((ObjToInt(interp, objv[0], &acetype) != TCL_OK) ||
        (ObjToInt(interp, objv[1], &aceflags) != TCL_OK)) {
        return TCL_ERROR;
    }

    /* Max size of an SID */
    acesz = GetSidLengthRequired(SID_MAX_SUB_AUTHORITIES);

    /* Figure out how much space is required for the ACE based on type */
    switch (acetype) {
    case ACCESS_ALLOWED_ACE_TYPE:
    case ACCESS_DENIED_ACE_TYPE:
    case SYSTEM_AUDIT_ACE_TYPE:
    case SYSTEM_MANDATORY_LABEL_ACE_TYPE:
        if (objc != 4)
            goto format_error;
        acesz += sizeof(*aceP);
        aceP = (ACCESS_ALLOWED_ACE *) TwapiAlloc(acesz);
        aceP->Header.AceType = acetype;
        aceP->Header.AceFlags = aceflags;
        aceP->Header.AceSize  = acesz; /* TBD - this is a upper bound since we
                                          allocated max SID size. Is that OK?*/
        if (ObjToInt(interp, objv[2], &aceP->Mask) != TCL_OK)
            goto format_error;

        sidP = TwapiGetSidFromStringRep(ObjToString(objv[3]));
        if (sidP == NULL)
            goto system_error;

        if (! CopySid(aceP->Header.AceSize - sizeof(*aceP) + sizeof(aceP->SidStart),
                      &aceP->SidStart, sidP)) {
            TwapiFree(sidP);
            goto system_error;
        }

        TwapiFree(sidP);
        sidP = NULL;

        break;

    default:
        if (objc != 3)
            goto format_error;
        bytes = ObjToByteArray(objv[2], &bytecount);
        acesz += bytecount;
        aceP = (ACCESS_ALLOWED_ACE *) TwapiAlloc(acesz);
        CopyMemory(aceP, bytes, bytecount);
        break;
    }

    *acePP = aceP;
    return TCL_OK;

 format_error:
    if (interp)
        ObjSetStaticResult(interp, "Invalid ACE format.");
    return TCL_ERROR;

 system_error:
    return TwapiReturnSystemError(interp);
}

Tcl_Obj *ObjFromACL (
    Tcl_Interp *interp,
    ACL *aclP                   /* May be NULL */
)
{
    Tcl_Obj                 *objv[2] = { NULL, NULL} ;
    ACL_REVISION_INFORMATION acl_rev;
    ACL_SIZE_INFORMATION     acl_szinfo;
    DWORD                    i;

    if (aclP == NULL) {
        return STRING_LITERAL_OBJ("null");
    }

    if ((GetAclInformation(aclP, &acl_rev, sizeof(acl_rev),
                           AclRevisionInformation) == 0) ||
        GetAclInformation(aclP, &acl_szinfo, sizeof(acl_szinfo),
                          AclSizeInformation) == 0) {
        TwapiReturnSystemError(interp);
        return NULL;
    }

    objv[0] = ObjFromInt(acl_rev.AclRevision);
    objv[1] = ObjNewList(0, NULL);

    /* Loop and add the list of ACE's */
    for (i = 0; i < acl_szinfo.AceCount; ++i) {
        void    *aceP;
        Tcl_Obj *ace_obj;

        if (GetAce(aclP, i, &aceP) == 0) {
            TwapiReturnSystemError(interp);
            goto error_return;
        }
        ace_obj = ObjFromACE(interp, aceP);
        if (ace_obj == NULL)
            goto error_return;
        if (ObjAppendElement(interp, objv[1], ace_obj) != TCL_OK) {
            goto error_return;
        }
    }


    return ObjNewList(2, objv);

 error_return:
    Twapi_FreeNewTclObj(objv[0]); /* OK if null */
    Twapi_FreeNewTclObj(objv[1]); /* OK if null */
    return NULL;
}


/*
 * Returns a pointer to dynamic memory containing a ACL corresponding
 * to the given string representation. The string "null" is treated
 * as no acl and a NULL pointer is returned in *aclPP
 */
int ObjToPACL(Tcl_Interp *interp, Tcl_Obj *aclObj, ACL **aclPP)
{
    int       objc;
    Tcl_Obj **objv;
    Tcl_Obj **aceobjv;
    int       aceobjc;
    void    **acePP = NULL;
    int       i;
    int       aclsz;
    ACE_HEADER *acehdrP;
    int       aclrev;

    *aclPP = NULL;
    if (!lstrcmpA("null", ObjToString(aclObj)))
        return TCL_OK;

    if (ObjGetElements(interp, aclObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc != 2) {
        if (interp)
            ObjSetStaticResult(interp, "Invalid ACL format.");
        return TCL_ERROR;
    }

    /*
     * First figure out how much space we need to allocate. For this, we
     * first need to figure out space for the ACE's
     */
#if 0
    objv[0] is the ACL rev. We always recalculate it, ignore value passed in.
    if (ObjToInt(interp, objv[0], &aclrev) != TCL_OK)
        goto error_return;
#endif
    if (ObjGetElements(interp, objv[1], &aceobjc, &aceobjv) != TCL_OK)
        goto error_return;

    aclsz = sizeof(ACL);
    aclrev = ACL_REVISION;
    if (aceobjc) {
        acePP = TwapiAlloc(aceobjc*sizeof(*acePP));
        for (i = 0; i < aceobjc; ++i)
            acePP[i] = NULL;        /* Init for error return */

        for (i = 0; i < aceobjc; ++i) {
            if (ObjToACE(interp, aceobjv[i], &acePP[i]) != TCL_OK)
                goto error_return;
            acehdrP = (ACE_HEADER *)acePP[i];
            aclsz += acehdrP->AceSize;
            switch (acehdrP->AceType) {
            case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
            case ACCESS_DENIED_OBJECT_ACE_TYPE:
            case SYSTEM_AUDIT_OBJECT_ACE_TYPE:
            case SYSTEM_ALARM_OBJECT_ACE_TYPE:
            case ACCESS_ALLOWED_CALLBACK_OBJECT_ACE_TYPE:
            case ACCESS_DENIED_CALLBACK_OBJECT_ACE_TYPE:
            case SYSTEM_AUDIT_CALLBACK_OBJECT_ACE_TYPE:
            case SYSTEM_ALARM_CALLBACK_OBJECT_ACE_TYPE:
                /* Change rev if object ace's present */
                aclrev = ACL_REVISION_DS;
                break;
            default:
                break;
            }

        }
    }

    /*
     * OK, now allocate the ACL and add the ACE's to it
     * We currently use AddAce, not AddMandatoryAce even for integrity labels.
     * This seems to work and avoids AddMandatoryAce which is not present
     * on XP/2k3
     */
    *aclPP = TwapiAlloc(aclsz);
    InitializeAcl(*aclPP, aclsz, aclrev);
    for (i = 0; i < aceobjc; ++i) {
        acehdrP = (ACE_HEADER *)acePP[i];
        if (! AddAce(*aclPP, aclrev, MAXDWORD, acePP[i], acehdrP->AceSize)) {
            TwapiReturnSystemError(interp);
            goto error_return;
        }
    }

    if (! IsValidAcl(*aclPP)) {
        if (interp)
            ObjSetStaticResult(interp, "Internal error constructing ACL");
        goto error_return;
    }

    /* Free up temporary ACE storage */
    if (acePP) {
        for (i = 0; i < aceobjc; ++i)
            TwapiFree(acePP[i]);
        TwapiFree(acePP);
    }

    return TCL_OK;

 error_return:
    if (acePP) {
        for (i = 0; i < aceobjc; ++i)
            TwapiFree(acePP[i]);
        TwapiFree(acePP);
    }

    if (*aclPP) {
        TwapiFree(*aclPP);
        *aclPP = NULL;
    }

    return TCL_ERROR;
}

/* Create a list object from a security descriptor */
Tcl_Obj *ObjFromSECURITY_DESCRIPTOR(
    Tcl_Interp *interp,
    SECURITY_DESCRIPTOR *secdP
)
{
    SECURITY_DESCRIPTOR_CONTROL secd_control;
    SID      *sidP;
    ACL      *aclP;
    BOOL      aclpresent;
    Tcl_Obj  *objv[5] = { NULL, NULL, NULL, NULL, NULL} ;
    DWORD    rev;
    BOOL     defaulted;

    if (secdP == NULL) {
        return ObjNewList(0, NULL);
    }

    if (! GetSecurityDescriptorControl(secdP, &secd_control, &rev))
        goto system_error;

    if (rev != SECURITY_DESCRIPTOR_REVISION) {
        /* Dunno how to handle this */
        if (interp)
            ObjSetStaticResult(interp, "Unsupported SECURITY_DESCRIPTOR version");
        goto error_return;
    }

    /* Control bits */
    objv[0] = ObjFromInt(secd_control);

    /* Owner SID */
    if (! GetSecurityDescriptorOwner(secdP, &sidP, &defaulted))
        goto system_error;
    if (sidP == NULL)
        objv[1] = ObjFromEmptyString();
    else {
        if (ObjFromSID(interp, sidP, &objv[1]) != TCL_OK)
            goto error_return;
    }

    /* Group SID */
    if (! GetSecurityDescriptorGroup(secdP, &sidP, &defaulted))
        goto system_error;
    if (sidP == NULL)
        objv[2] = ObjFromEmptyString();
    else {
        if (ObjFromSID(interp, sidP, &objv[2]) != TCL_OK)
            goto error_return;
    }

    /* DACL */
    if (! GetSecurityDescriptorDacl(secdP, &aclpresent, &aclP, &defaulted))
        goto system_error;
    if (! aclpresent)
        aclP = NULL;
    objv[3] = ObjFromACL(interp, aclP);

    /* SACL */
    if (! GetSecurityDescriptorSacl(secdP, &aclpresent, &aclP, &defaulted))
        goto system_error;
    if (! aclpresent)
        aclP = NULL;
    objv[4] = ObjFromACL(interp, aclP);

    /* All done, phew ... */
    return ObjNewList(5, objv);

 system_error:
    TwapiReturnSystemError(interp);

 error_return:
    for (rev = 0; rev < sizeof(objv)/sizeof(objv[0]); ++rev) {
        Twapi_FreeNewTclObj(objv[rev]);
    }
    return NULL;
}


/*
 * Returns a pointer to dynamic memory containing a structure corresponding
 * to the given string representation. Note that the owner, group, sacl
 * and dacl fields of the descriptor point to dynamic memory as well!
 */
int ObjToPSECURITY_DESCRIPTOR(
    Tcl_Interp *interp,
    Tcl_Obj *secdObj,
    SECURITY_DESCRIPTOR **secdPP
)
{
    int       objc;
    Tcl_Obj **objv;
    int       temp;
    SECURITY_DESCRIPTOR_CONTROL      secd_control;
    SECURITY_DESCRIPTOR_CONTROL      secd_control_mask;
    SID      *owner_sidP;
    SID      *group_sidP;
    ACL      *daclP;
    ACL      *saclP;
    char     *s;
    int       slen;

    owner_sidP = group_sidP = NULL;
    *secdPP = NULL;

    if (ObjGetElements(interp, secdObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc == 0)
        return TCL_OK;          /* NULL security descriptor */

    if (objc != 5) {
        if (interp)
            ObjSetStaticResult(interp, "Invalid SECURITY_DESCRIPTOR format.");
        return TCL_ERROR;
    }


    *secdPP = TwapiAlloc (sizeof(SECURITY_DESCRIPTOR));
    if (! InitializeSecurityDescriptor(*secdPP, SECURITY_DESCRIPTOR_REVISION))
        goto system_error;

    /*
     * Set control field
     */
    if (ObjToInt(interp, objv[0], &temp) != TCL_OK)
        goto error_return;
    secd_control = (SECURITY_DESCRIPTOR_CONTROL) temp;
    if (secd_control != temp) {
        /* Truncation error */
        if (interp)
            ObjSetStaticResult(interp, "Invalid control flags for SECURITY_DESCRIPTOR");
        goto error_return;
    }

    /* Mask of control bits to be set through SetSecurityDescriptorControl*/
    /* Note you cannot set any other bits than these through the
       SetSecurityDescriptorControl */
    secd_control_mask =  (SE_DACL_AUTO_INHERIT_REQ | SE_DACL_AUTO_INHERITED |
                          SE_DACL_PROTECTED |
                          SE_SACL_AUTO_INHERIT_REQ | SE_SACL_AUTO_INHERITED |
                          SE_SACL_PROTECTED);

    if (! SetSecurityDescriptorControl(*secdPP, secd_control_mask, (SECURITY_DESCRIPTOR_CONTROL) (secd_control_mask & secd_control)))
        goto system_error;

    /*
     * Set Owner field if specified
     */
    s = ObjToStringN(objv[1], &slen);
    if (slen) {
        owner_sidP = TwapiGetSidFromStringRep(s);
        if (owner_sidP == NULL)
            goto system_error;
        if (! SetSecurityDescriptorOwner(*secdPP, owner_sidP,
                                         secd_control & SE_OWNER_DEFAULTED))
            goto system_error;
        /* Note the owner field in *secdPP now points directly to owner_sidP! */
    }

    /*
     * Set group field if specified
     */
    s = ObjToStringN(objv[2], &slen);
    if (slen) {
        group_sidP = TwapiGetSidFromStringRep(s);
        if (group_sidP == NULL)
            goto system_error;

        if (! SetSecurityDescriptorGroup(*secdPP, group_sidP,
                                         secd_control & SE_GROUP_DEFAULTED))
            goto system_error;
        /* Note the group field in *secdPP now points directly to group_sidP! */
    }

    /*
     * Set the DACL. Keyword "null" means no DACL (as opposed to an empty one)
     */
    if (ObjToPACL(interp, objv[3], &daclP) != TCL_OK)
        goto error_return;
    if (! SetSecurityDescriptorDacl(*secdPP, (daclP != NULL), daclP,
                                  (secd_control & SE_DACL_DEFAULTED)))
        goto system_error;
    /* Note the dacl field in *secdPP now points directly to daclP! */


    /*
     * Set the SACL. Keyword "null" means no SACL (as opposed to an empty one)
     */
    if (ObjToPACL(interp, objv[4], &saclP) != TCL_OK)
        goto error_return;
    if (! SetSecurityDescriptorSacl(*secdPP, (saclP != NULL), saclP,
                                  (secd_control & SE_SACL_DEFAULTED)))
        goto system_error;
    /* Note the sacl field in *secdPP now points directly to saclP! */
    return TCL_OK;

 system_error:
    TwapiReturnSystemError(interp);
    goto error_return;

 error_return:
    if (owner_sidP)
        TwapiFree(owner_sidP);
    if (group_sidP)
        TwapiFree(group_sidP);
    if (daclP)
        TwapiFree(daclP);
    if (saclP)
        TwapiFree(saclP);
    if (*secdPP) {
        TwapiFree(*secdPP);
        *secdPP = NULL;
    }
    return TCL_ERROR;
}


/* Free the security descriptor contents as if it was allocated through
 * ObjToPSECURITY_DESCRIPTOR
 */
void TwapiFreeSECURITY_DESCRIPTOR(SECURITY_DESCRIPTOR *secdP)
{
    SID      *sidP;
    ACL      *aclP;
    BOOL      aclpresent;
    BOOL      defaulted;

    if (secdP == NULL)
        return;

    if (!IsValidSecurityDescriptor(secdP)) {
        return;                 /* TBD - Should log an error here */
    }

    /* Owner SID */
    if (GetSecurityDescriptorOwner(secdP, &sidP, &defaulted) && sidP)
        TwapiFree(sidP);

    /* Group SID */
    if (GetSecurityDescriptorGroup(secdP, &sidP, &defaulted) && sidP)
        TwapiFree(sidP);

    /* DACL */
    if (GetSecurityDescriptorDacl(secdP, &aclpresent, &aclP, &defaulted)
        && aclpresent
        && aclP) {
        TwapiFree(aclP);
    }

    /* SACL */
    if (GetSecurityDescriptorSacl(secdP, &aclpresent, &aclP, &defaulted)
        && aclpresent
        && aclP) {

        TwapiFree(aclP);
    }

    TwapiFree(secdP);
}

/* Free the security descriptor contents as if it was allocated through
 * ObjToPSECURITY_ATTRIBUTES
 */
void TwapiFreeSECURITY_ATTRIBUTES(SECURITY_ATTRIBUTES *secattrP)
{
    if (secattrP == NULL)
        return;
    if (secattrP->lpSecurityDescriptor)
        TwapiFreeSECURITY_DESCRIPTOR(secattrP->lpSecurityDescriptor);

    TwapiFree(secattrP);
}


/*
 * Returns a pointer to dynamic memory containing a structure corresponding
 * to the given string representation.
 * The SECURITY_DESCRIPTOR field should be freed through
 * TwapiFreeSECURITY_DESCRIPTOR
 */
int ObjToPSECURITY_ATTRIBUTES(
    Tcl_Interp *interp,
    Tcl_Obj *secattrObj,
    SECURITY_ATTRIBUTES **secattrPP
)
{
    int       objc;
    Tcl_Obj **objv;
    int       inherit;


    *secattrPP = NULL;

    if (ObjGetElements(interp, secattrObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc == 0)
        return TCL_OK;          /* NULL security attributes */

    if (objc != 2) {
        if (interp)
            ObjSetStaticResult(interp, "Invalid SECURITY_ATTRIBUTES format.");
        return TCL_ERROR;
    }


    *secattrPP = TwapiAlloc (sizeof(**secattrPP));
    (*secattrPP)->nLength = sizeof(**secattrPP);

    if (ObjToInt(interp, objv[1], &inherit) == TCL_ERROR)
        goto error_return;
    (*secattrPP)->bInheritHandle = (inherit != 0);

    if (ObjToPSECURITY_DESCRIPTOR(interp, objv[0],
                                     &(SECURITY_DESCRIPTOR *)((*secattrPP)->lpSecurityDescriptor))
        == TCL_ERROR) {
        goto error_return;
    }

    return TCL_OK;

 error_return:
    if (*secattrPP) {
        TwapiFree(*secattrPP);
        *secattrPP = NULL;
    }
    return TCL_ERROR;
}

void ObjIncrRefs(Tcl_Obj *objP) 
{
    Tcl_IncrRefCount(objP);
}

void ObjDecrRefs(Tcl_Obj *objP) 
{
    Tcl_DecrRefCount(objP);
}

void ObjDecrArrayRefs(int objc, Tcl_Obj *objv[])
{
    int i;
    for (i = 0; i < objc; ++i)
        ObjDecrRefs(objv[i]);
}


Tcl_UniChar *ObjToUnicode(Tcl_Obj *objP)
{
    return Tcl_GetUnicode(objP);
}

Tcl_UniChar *ObjToUnicodeN(Tcl_Obj *objP, int *lenP)
{
    return Tcl_GetUnicodeFromObj(objP, lenP);
}

Tcl_Obj *ObjFromUnicodeN(const Tcl_UniChar *ws, int len)
{
    if (ws == NULL)
        return ObjFromEmptyString(); /* TBD - log ? */

#if USE_UNICODE_OBJ
    return Tcl_NewUnicodeObj(ws, len);
#else
    return TwapiUtf8ObjFromUnicode(ws, len);
#endif
}

Tcl_Obj *ObjFromUnicode(const Tcl_UniChar *ws)
{
    if (ws == NULL)
        return ObjFromEmptyString(); /* TBD - log ? */

#if USE_UNICODE_OBJ
    return Tcl_NewUnicodeObj(ws, -1);
#else
    return TwapiUtf8ObjFromUnicode(ws, -1);
#endif
}

char *ObjToString(Tcl_Obj *objP)
{
    return Tcl_GetString(objP);
}

char *ObjToStringN(Tcl_Obj *objP, int *lenP)
{
    return Tcl_GetStringFromObj(objP, lenP);
}

Tcl_Obj *ObjFromStringN(const char *s, int len)
{
    if (s == NULL)
        return ObjFromEmptyString(); /* TBD - log ? */
    return Tcl_NewStringObj(s, len);
}

Tcl_Obj *ObjFromString(const char *s)
{
    if (s == NULL)
        return ObjFromEmptyString(); /* TBD - log ? */
    return Tcl_NewStringObj(s, -1);
}

int ObjCharLength(Tcl_Obj *objP)
{
    return Tcl_GetCharLength(objP);
}

TCL_RESULT ObjToLong(Tcl_Interp *interp, Tcl_Obj *objP, long *lvalP)
{
    return Tcl_GetLongFromObj(interp, objP, lvalP);
}

TCL_RESULT ObjToBoolean(Tcl_Interp *interp, Tcl_Obj *objP, int *valP)
{
    return Tcl_GetBooleanFromObj(interp, objP, valP);
}

TCL_RESULT ObjToWideInt(Tcl_Interp *interp, Tcl_Obj *objP, Tcl_WideInt *wideP)
{
    return Tcl_GetWideIntFromObj(interp, objP, wideP);
}

TCL_RESULT ObjToDouble(Tcl_Interp *interp, Tcl_Obj *objP, double *dblP)
{
    return Tcl_GetDoubleFromObj(interp, objP, dblP);
}

Tcl_Obj *ObjNewList(int objc, Tcl_Obj * const objv[])
{
    return Tcl_NewListObj(objc, objv);
}

Tcl_Obj *ObjEmptyList(void)
{
    return Tcl_NewObj();
}

TCL_RESULT ObjListLength(Tcl_Interp *interp, Tcl_Obj *l, int *lenP)
{
    return Tcl_ListObjLength(interp, l, lenP);
}

TCL_RESULT ObjAppendElement(Tcl_Interp *interp, Tcl_Obj *l, Tcl_Obj *e)
{
    return Tcl_ListObjAppendElement(interp, l, e);
}

TCL_RESULT ObjListIndex(Tcl_Interp *interp, Tcl_Obj *l, int ix, Tcl_Obj **objPP)
{
    return Tcl_ListObjIndex(interp, l, ix, objPP);
}

TCL_RESULT ObjGetElements(Tcl_Interp *interp, Tcl_Obj *l, int *objcP, Tcl_Obj ***objvP)
{
    return Tcl_ListObjGetElements(interp, l, objcP, objvP);
}

TCL_RESULT ObjListReplace(Tcl_Interp *interp, Tcl_Obj *l, int first, int count, int objc, Tcl_Obj *const objv[])
{
    return Tcl_ListObjReplace(interp, l, first, count, objc, objv);
}

Tcl_Obj *ObjNewDict()
{
    return Tcl_NewDictObj();
}

TCL_RESULT ObjDictGet(Tcl_Interp *interp, Tcl_Obj *dictObj, Tcl_Obj *keyObj, Tcl_Obj **valueObjP)
{
    return Tcl_DictObjGet(interp, dictObj, keyObj, valueObjP);
}

TCL_RESULT ObjDictPut(Tcl_Interp *interp, Tcl_Obj *dictObj, Tcl_Obj *keyObj, Tcl_Obj *valueObj)
{
    return Tcl_DictObjPut(interp, dictObj, keyObj, valueObj);
}


Tcl_Obj *ObjFromLong(long val)
{
    return Tcl_NewLongObj(val);
}

Tcl_Obj *ObjFromWideInt(Tcl_WideInt val)
{
    return Tcl_NewWideIntObj(val);
}

Tcl_Obj *ObjFromDouble(double val)
{
    return Tcl_NewDoubleObj(val);
}

Tcl_Obj *ObjFromBoolean(int bval)
{
    return Tcl_NewBooleanObj(bval);
}

Tcl_Obj *ObjFromEmptyString()
{
    return Tcl_NewObj();
}

Tcl_Obj *ObjFromByteArray(const unsigned char *bytes, int len)
{
    /* Older versions (< 8.6b1) of Tcl do not allow bytes to be NULL */

    /* Assumes major version==8 check already made at init time */
    if (bytes != NULL ||
        (gTclVersion.minor == 6 && gTclVersion.reltype == TCL_FINAL_RELEASE) ||
        gTclVersion.minor > 6) {
        return Tcl_NewByteArrayObj(bytes, len);
    } else {
        Tcl_Obj *o;
        o =  Tcl_NewByteArrayObj("", 0);
        Tcl_SetByteArrayLength(o, len);
        return o;
    }
}

unsigned char *ObjToByteArray(Tcl_Obj *objP, int *lenP)
{
    return Tcl_GetByteArrayFromObj(objP, lenP);
}

/* RtlEncryptMemory (aka SystemFunction040)
 * The Feb 2003 SDK does not define this in the headers, nor does it
 * include it in the export libs. So we have to dynamically load it.
*/
typedef BOOLEAN (NTAPI *SystemFunction040_t)(PVOID, ULONG, ULONG);
MAKE_DYNLOAD_FUNC(SystemFunction040, advapi32, SystemFunction040_t)

/* Encrypts the Unicode string rep to a byte array. We choose Unicode
   as the base format because most API's need Unicode and makes it
   easier to SecureZeroMemory the buffer. 
*/
Tcl_Obj *ObjEncryptUnicode(Tcl_Interp *interp, WCHAR *uniP, int nchars)
{
    int sz, len, pad_len;
    Tcl_Obj *objP;
    NTSTATUS status;
    SystemFunction040_t fnP = Twapi_GetProc_SystemFunction040();
    char *p;

    if (fnP == NULL) {
        Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
        return NULL;
    }

    if (nchars < 0)
        nchars = lstrlenW(uniP);

    len = sizeof(WCHAR) * nchars;

    /*
     * Total length has to be multiple of encryption block size
     * so encryption involves padding. We will stick a byte
     * at the end to hold actual pad length
     */
#ifndef RTL_ENCRYPT_MEMORY_SIZE // Not defined in all SDK's
# define RTL_ENCRYPT_MEMORY_SIZE 8
#endif    
#define BLOCK_SIZE_MASK (RTL_ENCRYPT_MEMORY_SIZE-1)
    
    if (len & BLOCK_SIZE_MASK) {
        /* Not a multiple of RTL_ENCRYPT_MEMORY_SIZE */
        sz = (len + BLOCK_SIZE_MASK) & ~BLOCK_SIZE_MASK;
        pad_len = sz - len;
    } else {
        /* Exact size. But we need a byte for the pad length field
           so need to add an entire block for that */
        pad_len = RTL_ENCRYPT_MEMORY_SIZE;
        sz = len + pad_len;
    }
    TWAPI_ASSERT(pad_len > 0);
    TWAPI_ASSERT(pad_len <= RTL_ENCRYPT_MEMORY_SIZE);

    objP = ObjFromByteArray(NULL, sz);
    p = ObjToByteArray(objP, &sz);
    p[sz-1] = (unsigned char) pad_len;
    CopyMemory(p, uniP, sizeof(WCHAR)*nchars);
    
    /* RtlEncryptMemory */
    status = fnP(p, sz, 0);
    if (status != 0) {
        ObjDecrRefs(objP);
        Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status));
        return NULL;
    }
    return objP;
}


/* RtlDecryptMemory (aka SystemFunction041)
 * The Feb 2003 SDK does not define this in the headers, nor does it
 * include it in the export libs. So we have to dynamically load it.
*/
typedef BOOLEAN (NTAPI *SystemFunction041_t)(PVOID, ULONG, ULONG);
MAKE_DYNLOAD_FUNC(SystemFunction041, advapi32, SystemFunction041_t)

/*
 * Decrypts Unicode string in objP to memory. Must be freed via TwapiFree
 * Returns length of string (in unichars, not bytes) NOT including terminating
 * char in *ncharsP
 */
WCHAR * ObjDecryptUnicode(Tcl_Interp *interp,
                          Tcl_Obj *objP,
                          int *ncharsP /* May be NULL */
    )
{
    int len;
    char *enc, *dec;
    int pad_len;
    NTSTATUS status;
    SystemFunction041_t fnP = Twapi_GetProc_SystemFunction041();

    if (fnP == NULL) {
        Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
        return NULL;
    }
    
    enc = ObjToByteArray(objP, &len);
    if (len == 0 || (len & BLOCK_SIZE_MASK)) {
        if (interp)
            TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                               Tcl_ObjPrintf("Invalid length (%d) of encrypted object. Must be non-zero multiple of block size (%d).", len, RTL_ENCRYPT_MEMORY_SIZE));
        return NULL;
    }

    /* len is number of encrypted bytes. Number of decrypted bytes
       will be at most len. We will need another 2 bytes for a WCHAR \0
    */
    dec = TwapiAlloc(len+sizeof(WCHAR));
    CopyMemory(dec, enc, len);

    /* RtlDecryptMemory */
    status = fnP(dec, len, 0);
    if (status != 0) {
        if (interp)
            Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status));
        goto error_return;
    }

    pad_len = (unsigned char) dec[len-1];

    /* Note we always encode unicode chars so pad_len is always even */
    if (pad_len == 0 || (pad_len & 1) ||
        pad_len > RTL_ENCRYPT_MEMORY_SIZE ||
        pad_len > len) {
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid pad (%d) in decrypted object. Object corrupted or was not encrypted.", pad_len));
        goto error_return;
    }
    
    len -= pad_len;
    len /= sizeof(WCHAR);
    *(len + (WCHAR*) dec) = 0;             /* Terminate the Unicode string */
    if (ncharsP)
        *ncharsP = len;

    return (WCHAR*)dec;

error_return:
    if (dec)
        TwapiFree(dec);
    return NULL;
}


/* If decryption fails, assumes password in unencrypted form and returns it.
 * Returned buffer must be freed via TwapiFree
 */
WCHAR * ObjDecryptPassword(Tcl_Obj *objP, int *ncharsP)
{
    WCHAR *uniP;

    uniP =  ObjDecryptUnicode(NULL, objP, ncharsP);
    return uniP ? uniP : TwapiAllocWStringFromObj(objP, ncharsP);
}

void TwapiFreeDecryptedPassword(WCHAR *p, int len)
{
    if (p) {
        SecureZeroMemory(p, len * sizeof(WCHAR));
        TwapiFree(p);
    }
}

/*
 * TwapiEnum is a Tcl "type" that maps strings to their positions in a
 * string table. 
 * The Tcl_Obj.internalRep.ptrAndLongRep.value holds the position
 * and Tcl_Obj.internalRep.ptrAndLongRep.ptr holds the table as a Tcl_Obj
 * of type list.
 */
static void DupEnumType(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void FreeEnumType(Tcl_Obj *objP);
static void UpdateEnumTypeString(Tcl_Obj *objP);
static struct Tcl_ObjType gEnumType = {
    "TwapiEnum",
    FreeEnumType,
    DupEnumType,
    UpdateEnumTypeString,
    NULL,     /* jenglish says keep this NULL */
};



static void DupEnumType(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    dstP->typePtr = srcP->typePtr;
    dstP->internalRep = srcP->internalRep;
    ObjIncrRefs(dstP->internalRep.ptrAndLongRep.ptr);
}

static void FreeEnumType(Tcl_Obj *objP)
{
    ObjDecrRefs(objP->internalRep.ptrAndLongRep.ptr);
    objP->internalRep.ptrAndLongRep.ptr = NULL;
    objP->internalRep.ptrAndLongRep.value = -1;
    objP->typePtr = NULL;
}

static void UpdateEnumTypeString(Tcl_Obj *objP)
{
    Tcl_Obj *obj2P;
    char *p;
    int len;
    TCL_RESULT res;

    TWAPI_ASSERT(objP->bytes == NULL);
    TWAPI_ASSERT(objP->typePtr == &gEnumType);
    TWAPI_ASSERT(objP->internalRep.ptrAndLongRep.ptr);
    
    res = ObjListIndex(NULL, objP->internalRep.ptrAndLongRep.ptr, objP->internalRep.ptrAndLongRep.value, &obj2P);
    TWAPI_ASSERT(res == TCL_OK);
    TWAPI_ASSERT(obj2P);

    p = ObjToStringN(obj2P, &len);
    objP->bytes = ckalloc(len+1);
    objP->length = len;
    CopyMemory(objP->bytes, p, len+1);
}

TCL_RESULT ObjToEnum(Tcl_Interp *interp, Tcl_Obj *enumsObj, Tcl_Obj *nameObj,
                     int *valP)
{
    TCL_RESULT res;

    /* Reconstruct if not gEnumType or if it is one but for a different table */
    if (nameObj->typePtr != &gEnumType ||
        enumsObj != nameObj->internalRep.ptrAndLongRep.ptr) {
        Tcl_Obj **objs;
        int i, nobjs;
        char *nameP;

        if ((res = ObjGetElements(interp, enumsObj, &nobjs, &objs)) != TCL_OK)
            return res;

        nameP = ObjToString(nameObj);
        for (i = 0; i < nobjs; ++i) {
            if (STREQ(nameP, ObjToString(objs[i])))
                break;
        }
        if (i == nobjs) {
            if (interp)
                ObjSetResult(interp, Tcl_ObjPrintf("Invalid enum \"%s\"", nameP));
            return TCL_ERROR;
        }

        if (nameObj->typePtr && nameObj->typePtr->freeIntRepProc) {
            nameObj->typePtr->freeIntRepProc(nameObj);
        }

        nameObj->typePtr = &gEnumType;
        nameObj->internalRep.ptrAndLongRep.value = i;
        nameObj->internalRep.ptrAndLongRep.ptr = enumsObj;
        ObjIncrRefs(enumsObj);
    }

    *valP = nameObj->internalRep.ptrAndLongRep.value;
    return TCL_OK;
}
