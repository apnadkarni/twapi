/*
 * Copyright (c) 2010-2015, Ashok P. Nadkarni
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
    {VT_TWAPI_VARNAME, "outvar"} /* TWAPI-specific hack to indicate contents
                                    to be treated as a variable name. Used
                                    for tclcast */
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
            nobjs != 2)
            goto invalid_value;
        if (ObjToDWORD_PTR(NULL, objs[0], &dwp) != TCL_OK) {
            s = ObjToString(objs[0]);
            goto invalid_value;
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

invalid_value: /* s must point to value */
    if (interp)
        Tcl_AppendResult(interp, "Invalid pointer or opaque value '",
                         s, "'.", NULL);
    return TCL_ERROR;
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
    case VT_BSTR: /* Fall through */
    case VT_TWAPI_VARNAME: /* Fall through */
    default:
        /* String rep should already be there. This function should not have been called. */
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
    Tcl_Obj *objP = NULL;
    Tcl_ObjType *typeP;
    const char *typename;
    int i;
    VARTYPE vt;

    if (objc != 3)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    /* NOTE
     * ALWAYS DUP/CREATE NEW object and do not convert it in place.
     * This reduces the chances of it shimmering back to a different
     * type. An example test case is illustrated in test
     * case variant_param_passing-safearray-2.18. The safearray conversion
     * in that test converts "0" to variant ("") but that would
     * get shimmered back to an int in the safearray proc lindex call.
     * We want to prevent this.
     *
     * TBD - may be we can optimize this if the object is unshared ?
     */

    typename = ObjToString(objv[1]);

    if (*typename == '\0') {
        /* No type, make pure string */
        return ObjSetResult(interp, ObjFromString(ObjToString(objv[2])));
    }
        
    /*
     * We special case int because Tcl will convert 0x80000000 to wide int
     * but we want it to be passed as 32-bits in case of some COM calls.
     * which will barf for VT_I8
     *
     * We special case double, because SetAnyFromProc will optimize to lowest
     * compatible type, so for example casting 1 to double will result in an
     * int object. We want to force it to double.
     *
     * We special case "boolean" and "booleanString" because they will keep
     * numerics as numerics while we want to force to boolean Tcl_Obj type.
     * We do this even before GetObjType because "booleanString" is not
     * even a registered type in Tcl.
     *
     * We can't do anything about wideInt because Tcl_NewWideIntObj,
     * and SetWideIntFromAny, will also return an
     * int Tcl_Obj if the value fits in the 32 bits.
     */
    if (STREQ(typename, "int")) {
        int ival;
        if (ObjToInt(interp, objv[2], &ival) == TCL_ERROR)
            return TCL_ERROR;
        return ObjSetResult(interp, ObjFromInt(ival));
    }

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

    objP = NULL;
    typeP = Tcl_GetObjType(typename);
    if (typeP) {
        objP = ObjDuplicate(objv[2]);
        if (objP->typePtr != typeP) {
            if (typeP->setFromAnyProc == NULL)
                goto convert_error;
            if (Tcl_ConvertToType(interp, objP, typeP) == TCL_ERROR) {
                goto error_handler;
            }
        }
    } else {
        /* Not a registered Tcl type. See if one of ours */
        if (LookupBaseVTToken(NULL, typename, &vt) == TCL_OK) {
            switch (vt) {
            case VT_EMPTY:
            case VT_NULL:
                if (ObjCharLength(objv[2]) != 0)
                    goto convert_error;
                /* Fall thru */
            case VT_BSTR: /* Fall through */
            case VT_TWAPI_VARNAME:
                objP = ObjDuplicate(objv[2]);
                ObjToString(objP);  /* Make sure string rep exists */
                objP->typePtr = &gVariantType;
                VARIANT_REP_VT(objP) = vt;
            }
        }
    }

    if (objP)
        return ObjSetResult(interp, objP);

convert_error:
    Tcl_AppendResult(interp, "Cannot convert '", ObjToString(objv[2]), "' to type '", typename, "'", NULL);
error_handler:
    if (objP)
        ObjDecrRefs(objP);
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
    DWORD dw;

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
    case TRT_HMACHINE:
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
        case TRT_HMACHINE:
            typenameP = "HMACHINE";
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

    case TRT_CONFIGRET:
        if (resultP->value.ival != 0) {
            ObjSetResult(interp, Tcl_ObjPrintf("PnP Manager returned error code %d", resultP->value.ival));
            return TCL_ERROR;
        }
        break;

    case TRT_GETLASTERROR_SETUPAPI:
        dw = GetLastError();
        return Twapi_AppendSystemError(interp, HRESULT_FROM_SETUPAPI(dw));

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
    int      numbytes, len;
    LPITEMIDLIST idsP;

    idsP = (LPITEMIDLIST) ObjToByteArray(objP, &numbytes);
    len = numbytes;
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

    *idsPP = CoTaskMemAlloc(len);
    if (*idsPP == NULL) {
        if (interp)
            ObjSetStaticResult(interp, "CoTaskMemAlloc failed in SHChangeNotify");
        return TCL_ERROR;
    }

    CopyMemory(*idsPP, idsP, len);

    return TCL_OK;
}

void TwapiFreePIDL(LPITEMIDLIST idlistP)
{
    if (idlistP) {
        CoTaskMemFree(idlistP);
    }
}

Tcl_Obj *ObjFromGUID(const GUID *guidP)
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
    /* Permit objP to be NULL so we can use it with ARGVARWITHDEFAULT */
    if (objP == NULL || ObjCharLength(objP) == 0) {
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
 * with two nulls.  Pointer to memlifo alloced multi_sz is stored
 * in *multiszPtrPtr on success. Note that
 * memory may or may not have been allocated from pool even when error
 * is returned and must be freed.
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

    TWAPI_ASSERT(lifoP); /* Check since original API allowed lifoP == NULL */
    
    *multiszPtrPtr = NULL;
    for (i=0, len=0; ; ++i) {
        if (ObjListIndex(interp, listPtr, i, &objPtr) == TCL_ERROR)
            return TCL_ERROR;
        if (objPtr == NULL)
            break;              /* No more items */
        len += ObjCharLength(objPtr) + 1;
    }

    ++len;                      /* One extra null char at the end */
    buf = MemLifoAlloc(lifoP, len*sizeof(*buf), NULL);

    for (i=0, dst=buf; ; ++i) {
        if (ObjListIndex(interp, listPtr, i, &objPtr) == TCL_ERROR)
            return TCL_ERROR;
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

    TWAPI_ASSERT(sizeof(WORD) == sizeof(unsigned short));
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


Tcl_Obj *ObjFromEXPAND_SZW(WCHAR *ws)
{
    DWORD sz, required;
    WCHAR buf[MAX_PATH+1], *bufP;
    Tcl_Obj *objP;

    bufP = buf;
    sz = ARRAYSIZE(buf);
    required = ExpandEnvironmentStringsW(ws, bufP, sz);
    if (required > sz) {
        // Need a bigger buffer
        bufP = SWSPushFrame(required*sizeof(WCHAR), NULL);
        sz = required;
        required = ExpandEnvironmentStringsW(ws, bufP, sz);
    }
    if (required <= 0 || required > sz)
        objP = ObjFromUnicode(ws); /* Something went wrong, return as is */
    else {
        --required;         /* Terminating null */
        objP = ObjFromUnicodeN(bufP, required);
    }

    if (bufP != buf)
        SWSPopFrame();
    return objP;
}

Tcl_Obj *ObjFromRegValue(Tcl_Interp *interp, int regtype,
                         BYTE *bufP, int count)
{
    Tcl_Obj *objv[2];
    char *typestr = NULL;
    WCHAR *ws;

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
        ws = (WCHAR *)bufP;
        count /= 2;             /*  Assumed to be Unicode. */
        if (count && ws[count-1] == 0)
            --count;        /* Do not include \0 */
        objv[1] = ObjFromUnicodeN(ws, count);
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


Tcl_Obj *ObjFromRegValueCooked(Tcl_Interp *interp, int regtype,
                         BYTE *bufP, int count)
{
    WCHAR *ws;
    int terminated = 0;

    switch (regtype) {
    case REG_LINK:
    case REG_SZ:
    case REG_EXPAND_SZ:
        /*
         * As per MS docs, may not always be null terminated.
         * If it is, we need to strip off the null.
         */
        ws = (WCHAR *)bufP;
        count /= 2;             /*  Assumed to be Unicode. */
        if (count && ws[count-1] == 0) {
            terminated = 1;
            --count;        /* Do not include \0 */
        }
        if (regtype != REG_EXPAND_SZ || count == 0)
            return ObjFromUnicodeN(ws, count);
        else {
            /* Expand expects null terminated */
            if (terminated)
                return ObjFromEXPAND_SZW(ws);
            else {
                Tcl_Obj *objP;
                WCHAR *ws2 = SWSPushFrame(sizeof(WCHAR)*(count+1), NULL);
                memcpy(ws2, ws, sizeof(WCHAR)*count);
                ws2[count] = 0;
                objP = ObjFromEXPAND_SZW(ws2);
                SWSPopFrame();
                return objP;
            }
        }
        break;
            
    case REG_DWORD_BIG_ENDIAN:
        if (count != 4)
            goto badformat;
        return ObjFromLong(swap4(*(unsigned long *)bufP));

    case REG_DWORD:
        if (count != 4)
            goto badformat;
        return ObjFromLong(*(long *)bufP);

    case REG_QWORD:
        if (count != 8)
            goto badformat;
        return ObjFromWideInt(*(Tcl_WideInt *)bufP);

    case REG_MULTI_SZ:
        return ObjFromMultiSz((LPCWSTR) bufP, count/2);

    case REG_BINARY:  // FALLTHRU
    default:
        return ObjFromByteArray(bufP, count);
    }

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

/* If buf is NULL or buf_sz is 0, returns the number of bytes required
   for the result. Else stores the utf-8 in the buffer and returns number
   of bytes stored. 

   The function does not explicitly add a terminating null. If the length
   (nchars) of wsP is explicitly specified, a terminating null will be included
   only if the wsP[nchars] was a null. If nchars was passed as a negative
   number, the terminating null is implicitly included in the length and will
   be present in the returned buffer as well.

   On error returns -1. Detail about error should be retrieved with
   GetLastError()
*/
int TwapiUnicharToUtf8(CONST WCHAR *wsP, int nchars, char *buf, int buf_sz)
{
    int nbytes;
    
    if (buf_sz < 1)
        buf = NULL;
    else if (buf == NULL)
        buf_sz = 0;
    
    /* Note WideChar... does not like 0 length strings */
    if (wsP == NULL || nchars == 0)
        return 0;

    if (nchars < 0)
        nchars = -1;
    
    nbytes = WideCharToMultiByte(
        CP_UTF8, /* CodePag */
        0,       /* dwFlags */
        wsP,     /* lpWideCharStr */
        nchars,  /* cchWideChar */
        buf,     /* lpMultiByteStr */
        buf_sz,  /* cbMultiByte */
        NULL,    /* lpDefaultChar */
        NULL     /* lpUsedDefaultChar */
        );
    
    if (nbytes == 0)
        return -1; /* Caller should use GetLastError() */

    return nbytes;
}

/* The Unicode string should not contain embedded nulls */
Tcl_Obj *TwapiUtf8ObjFromUnicode(CONST WCHAR *wsP, int nchars)
{
    int nbytes;
    char *p;
    Tcl_Obj *objP;
    
    nbytes = TwapiUnicharToUtf8(wsP, nchars, NULL, 0);
    if (nbytes == -1)
        goto do_panic;
    
    p = ckalloc(nbytes+1); /* +1 for case where WCTMB doesn't terminate \0 */ 
    nbytes = TwapiUnicharToUtf8(wsP, nchars, p, nbytes);
    if (nbytes == -1)
        goto do_panic;

    if (nbytes == 0)
        p[0] = '\0';
    else if (p[nbytes-1] == '\0')
        nbytes -= 1; /* Don't include terminating null */
    else
        p[nbytes] = '\0'; /* Terminate string */
    
    objP = Tcl_NewObj();
    if (objP->bytes)
        Tcl_InvalidateStringRep(objP);
    objP->bytes = p;
    objP->length = nbytes; /* Note length does not include terminating \0 */

    return objP;

do_panic:
    Tcl_Panic("TwapiUnicharToUtf8 returned -1, with error code %d", GetLastError());
    return NULL; /* To keep compiler happy. Never reaches here */
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
    Tcl_Obj **objs;
    int nobjs;
    int i;
    Tcl_WideInt wide;

    /* Return should be purely based on current type ptr of Tcl_Obj,
       NOT heuristics so be careful not to shimmer BEFORE checking */

    switch (TwapiGetTclType(objP)) {
    case TWAPI_TCLTYPE_BOOLEAN: /* Fallthru */
    case TWAPI_TCLTYPE_BOOLEANSTRING:
        return VT_BOOL;
    case TWAPI_TCLTYPE_INT:
        return VT_I4;
    case TWAPI_TCLTYPE_WIDEINT:
        /* For compatibility with some COM types, we want to return VT_I4
           if value fits in 32-bits irrespective of sign */
        if (ObjToWideInt(NULL, objP, &wide) == TCL_OK &&
            (wide & 0xffffffff00000000) == 0)
            return VT_I4;
        else
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
        if (ObjGetElements(NULL, objP, &nobjs, &objs) != TCL_OK)
            return vt;          /* Should not really happen */
        if (nobjs) {
            /* Base our guess on the first type. Note we don't coerce it
               since we might want to pass "1" as a string, not int.
               If *remaining* elements do not match that type after coercion,
               we fall back to VT_VARIANT.
            */
            switch (TwapiGetTclType(objs[0])) {
            case TWAPI_TCLTYPE_BOOLEAN: /* Fallthru */
            case TWAPI_TCLTYPE_BOOLEANSTRING:
                vt = VT_BOOL;
                for (i = 1; i < nobjs; ++ i) {
                    int bval;
                    if (ObjToBoolean(NULL, objs[i], &bval) != TCL_OK) {
                        vt = VT_VARIANT;
                        break;
                    }
                }
                break;
            case TWAPI_TCLTYPE_INT:
                vt = VT_I4;
                for (i = 1; i < nobjs; ++ i) {
                    int ival;
                    if (ObjToLong(NULL, objs[i], &ival) != TCL_OK) {
                        vt = VT_VARIANT;
                        break;
                    }
                }
                break;
            case TWAPI_TCLTYPE_WIDEINT:
                vt = VT_I8;
                for (i = 1; i < nobjs; ++ i) {
                    Tcl_WideInt wide;
                    if (ObjToWideInt(NULL, objs[i], &wide) != TCL_OK) {
                        vt = VT_VARIANT;
                        break;
                    }
                }
                break;
            case TWAPI_TCLTYPE_DOUBLE:
                vt = VT_R8;
                for (i = 1; i < nobjs; ++ i) {
                    double dbl;
                    if (ObjToDouble(NULL, objs[i], &dbl) != TCL_OK) {
                        vt = VT_VARIANT;
                        break;
                    }
                }
                break;
            case TWAPI_TCLTYPE_STRING:
                vt = VT_BSTR;
                break;
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

/* On error leaves variant in VT_EMPTY state */
TCL_RESULT ObjToVARIANT(Tcl_Interp *interp, Tcl_Obj *objP, VARIANT *varP, VARTYPE vt)
{
    HRESULT hr;
    long lval;
    TCL_RESULT res;

    if (vt & VT_ARRAY) {
        if (ObjToSAFEARRAY(interp, objP, &varP->parray, &vt) != TCL_OK) {
            varP->vt = VT_EMPTY;
            return TCL_ERROR;
        }
        varP->vt = vt;
        return TCL_OK;
    }

    switch (vt) {
    case VT_EMPTY:
    case VT_NULL:
        res = TCL_OK;
        break;
    case VT_I2:   res = ObjToSHORT(interp, objP, &V_I2(varP)); break;
    case VT_UI2:  res = ObjToUSHORT(interp, objP, &V_UI2(varP)); break;
    case VT_I1:   res = ObjToCHAR(interp, objP, &V_I1(varP)); break;
    case VT_UI1:   res = ObjToUCHAR(interp, objP, &V_UI1(varP)); break;

    /* For compatibility reasons, allow interchangeable signed/unsigned ints */
    case VT_I4:
    case VT_UI4:
    case VT_INT:
    case VT_UINT:
    case VT_HRESULT:
        res = ObjToInt(interp, objP, &lval);
        if (res == TCL_OK) {
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
        }
        break;

    case VT_R4:
    case VT_R8:
        res = ObjToDouble(interp, objP, &varP->dblVal);
        if (res == TCL_OK) {
            varP->vt = VT_R8;       /* Needed for VariantChangeType */
            if (vt == VT_R4) {
                hr = VariantChangeType(varP, varP, 0, VT_R4);
                if (FAILED(hr))
                    res = Twapi_AppendSystemError(interp, hr);
            }
        }
        break;

    case VT_CY:
        res = ObjToCY(interp, objP, & V_CY(varP));
        break;

    case VT_DATE:
        res = ObjToDouble(interp, objP, & V_DATE(varP));
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
        res = TCL_OK;
        if (ObjToLong(NULL, objP, &varP->lVal) == TCL_OK) {
            vt = VT_I4;
        } else if (ObjToDouble(NULL, objP, &varP->dblVal) == TCL_OK) {
            vt = VT_R8;
        } else if (ObjToIDispatch(NULL, objP, &varP->pdispVal) == TCL_OK) {
            vt = VT_DISPATCH;
        } else if (ObjToIUnknown(NULL, objP, &varP->punkVal) == TCL_OK) {
            vt = VT_UNKNOWN;
#if 0
        } else if (ObjCharLength(objP) == 0) {
            vt = VT_NULL;
#endif
        } else {
            /* Cannot guess type, just pass as a BSTR */
            vt = VT_BSTR;
            res = ObjToBSTR(interp, objP, &varP->bstrVal);
        }
        break;

    case VT_BSTR:
        res = ObjToBSTR(interp, objP, &varP->bstrVal);
        break;

    case VT_DISPATCH:
        res = ObjToIDispatch(interp, objP, (void **)&varP->pdispVal);
        break;

    case VT_VOID: /* FALLTHRU */
    case VT_ERROR:
        /* Treat as optional argument */
        vt = VT_ERROR;
        varP->scode = DISP_E_PARAMNOTFOUND;
        res = TCL_OK;
        break;

    case VT_BOOL:
        res = ObjToBoolean(interp, objP, &varP->intVal);
        if (res == TCL_OK)
            varP->boolVal = varP->intVal ? VARIANT_TRUE : VARIANT_FALSE;
        break;

    case VT_UNKNOWN:
        res = ObjToIUnknown(interp, objP, (void **) &varP->punkVal);
        break;

    case VT_DECIMAL:
        res = ObjToDECIMAL(interp, objP, & V_DECIMAL(varP));
        break;

    case VT_I8:
    case VT_UI8:
        vt = VT_I8;
        res = ObjToWideInt(interp, objP, &varP->llVal);
        break;

    default:
        ObjSetResult(interp,
                          Tcl_ObjPrintf("Invalid or unsupported VARTYPE (%d)",
                                        vt));
        res = TCL_ERROR;
        break;
    }

    varP->vt = res == TCL_OK ? vt : VT_EMPTY;

    return res;
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
    case VT_VARIANT: /* Note VT_VARIANT is illegal in a concrete VARIANT */
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


/* Returned SWS memory has to be freed by caller even in case of errors */
int ObjToLSASTRINGARRAYSWS(Tcl_Interp *interp, Tcl_Obj *obj, LSA_UNICODE_STRING **arrayP, ULONG *countP)
{
    Tcl_Obj **listobjv;
    int       i, nitems, sz;
    LSA_UNICODE_STRING *ustrP;
    MemLifo *memlifoP = SWS();

    if (ObjGetElements(interp, obj, &nitems, &listobjv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    /* Allocate the array of structures */
    sz = nitems * sizeof(LSA_UNICODE_STRING);
    ustrP = MemLifoAlloc(memlifoP, sz, NULL);
    
    for (i = 0; i < nitems; ++i) {
        WCHAR *srcP;
        int    slen;
        srcP = ObjToUnicodeN(listobjv[i], &slen);
        if (slen >= (USHRT_MAX/sizeof(WCHAR))) {
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "LSA_UNICODE_STRING length must be less than 32767.");
        }
        slen *= sizeof(WCHAR); /* # bytes */
        ustrP[i].Buffer = MemLifoCopy(memlifoP, srcP, slen+sizeof(WCHAR));
        ustrP[i].Length = (USHORT) slen; /* Num *bytes*, not WCHARs */
        ustrP[i].MaximumLength = ustrP[i].Length;
    }

    *arrayP = ustrP;
    *countP = nitems;

    return TCL_OK;
}

/*
 * Returns a pointer to SWS memory containing a SID corresponding
 * to the given string representation. Returns NULL on error, and
 * sets the windows error. Caller responsible for SWS memory in
 * both success and failure cases.
 */
PSID TwapiSidFromStringSWS(char *strP)
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
    sidP = SWSAlloc(len, NULL);
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

    SetLastError(error);
    return NULL;
}

/* Tcl_Obj to SID - the object may hold the SID string rep, a binary
   or a list of ints. If the object is an empty string, error returned.
   Else the SID is allocated  on the SWS and a pointer to it is
   stored in *sidPP. Caller responsible for SWS memory on success
   and failure.
*/
TCL_RESULT ObjToPSIDNonNullSWS(Tcl_Interp *interp, Tcl_Obj *obj, PSID *sidPP)
{
    DWORD   len;
    SID  *sidP;
    DWORD winerror;

    *sidPP = TwapiSidFromStringSWS(ObjToString(obj));
    if (*sidPP)
        return TCL_OK;

    winerror = GetLastError();

    /* Not a string rep. See if it is a binary of the right size */
    sidP = (SID *) ObjToByteArray(obj, &len);
    if (len >= sizeof(*sidP)) {
        /* Seems big enough, validate revision and size */
        if (IsValidSid(sidP) && GetLengthSid(sidP) == len) {
            *sidPP = SWSAlloc(len, NULL);
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

/* Tcl_Obj to SID - the object may hold the SID string rep, a binary
   or a list of ints. If the object is an empty string, *sidPP is
   stored as NULL. Else the SID is allocated on the SWS and a pointer to it is
   stored in *sidPP. Caller responsible for all SWS management.
*/
TCL_RESULT ObjToPSIDSWS(Tcl_Interp *interp, Tcl_Obj *obj, PSID *sidPP)
{
    char *s;
    DWORD   len;

    s = ObjToStringN(obj, &len);
    if (len == 0) {
        *sidPP = NULL;
        return TCL_OK;
    }

    return ObjToPSIDNonNullSWS(interp, obj, sidPP);
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

/* Returns an allocated on SWS. Caller responsible for all SWS management */
TCL_RESULT ObjToACESWS (Tcl_Interp *interp, Tcl_Obj *aceobj, void **acePP)
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
        aceP = (ACCESS_ALLOWED_ACE *) SWSAlloc(acesz, NULL);
        aceP->Header.AceType = acetype;
        aceP->Header.AceFlags = aceflags;
        aceP->Header.AceSize  = acesz; /* TBD - this is a upper bound since we
                                          allocated max SID size. Is that OK?*/
        if (ObjToInt(interp, objv[2], &aceP->Mask) != TCL_OK)
            goto format_error;

        sidP = TwapiSidFromStringSWS(ObjToString(objv[3]));
        if (sidP == NULL)
            goto system_error;

        if (! CopySid(aceP->Header.AceSize - sizeof(*aceP) + sizeof(aceP->SidStart),
                      &aceP->SidStart, sidP)) {
            goto system_error;
        }

        sidP = NULL;
        break;

    default: /* TBD - what is the logic of this? What is objv[2] ?  Need tests*/
        /* TBD - ObjFromACE seems to handle many more types than we do here */
        if (objc != 3)
            goto format_error;
        bytes = ObjToByteArray(objv[2], &bytecount);
        acesz += bytecount;
        aceP = (ACCESS_ALLOWED_ACE *) SWSAlloc(acesz, NULL);
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
 * Returns a pointer to SWS memory containing a ACL corresponding
 * to the given string representation. The string "null" is treated
 * as no acl and a NULL pointer is returned in *aclPP
 * Caller responsible for all SWS memory management
 */
int ObjToPACLSWS(Tcl_Interp *interp, Tcl_Obj *aclObj, ACL **aclPP)
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
        acePP = SWSAlloc(aceobjc*sizeof(*acePP), NULL);
        for (i = 0; i < aceobjc; ++i) {
            if (ObjToACESWS(interp, aceobjv[i], &acePP[i]) != TCL_OK)
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
                /* TBD - do we need to add others for VISTA and later ? */
                break;
            }

        }
    }

    /*
     * OK, now allocate the ACL and add the ACE's to it
     * We currently use AddAce, not AddMandatoryAce even for integrity labels.
     * This seems to work and avoids AddMandatoryAce which is not present
     * on XP/2k3. TBD
     */
    *aclPP = SWSAlloc(aclsz, NULL);
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

    return TCL_OK;

 error_return:
    *aclPP = NULL;
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
 * Returns a pointer to SWS memory containing a structure corresponding
 * to the given string representation. Note that the owner, group, sacl
 * and dacl fields of the descriptor point to SWS memory as well!
 * Note some functions that use this, such as CoInitializeSecurity
 * require the security descriptor to be in absolute format so do not
 * change this function to return a self-relative descriptor
 * Caller is responsible for all SWS memory in both success and failure cases
 */
TCL_RESULT ObjToPSECURITY_DESCRIPTORSWS(
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

    *secdPP = SWSAlloc (sizeof(SECURITY_DESCRIPTOR), NULL);
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
        owner_sidP = TwapiSidFromStringSWS(s);
        if (owner_sidP == NULL)
            goto system_error;
        /* TBD - the owner field is allowed to be NULL. How do we set that ?*/
        if (! SetSecurityDescriptorOwner(*secdPP, owner_sidP,
                                         (secd_control & SE_OWNER_DEFAULTED) != 0))
            goto system_error;
        /* Note the owner field in *secdPP now points directly to owner_sidP! */
    }

    /*
     * Set group field if specified
     */
    s = ObjToStringN(objv[2], &slen);
    if (slen) {
        group_sidP = TwapiSidFromStringSWS(s);
        if (group_sidP == NULL)
            goto system_error;

        if (! SetSecurityDescriptorGroup(*secdPP, group_sidP,
                                         (secd_control & SE_GROUP_DEFAULTED) != 0))
            goto system_error;
        /* Note the group field in *secdPP now points directly to group_sidP! */
    }

    /*
     * Set the DACL if the control flags mark it as being present.
     */
    if (secd_control & SE_DACL_PRESENT) {
        /*
         * Keyword "null" means no DACL (as opposed to an empty one)
         * In that case daclP will be NULL. We don't need to call
         * SetSecurityDescriptorDacl in that case because it is already 
         * initialized to not have a DACL. However, if the secd_control
         * flags have SE_DACL_DEFAULTED is set, we make the call because
         * we're not sure how the system treates a NULL DACL with
         * that flag set.
         */
        if (ObjToPACLSWS(interp, objv[3], &daclP) != TCL_OK)
            goto error_return;
        if (daclP || (secd_control & SE_DACL_DEFAULTED)) {
            if (! SetSecurityDescriptorDacl(*secdPP, TRUE, daclP,
                                            (secd_control & SE_DACL_DEFAULTED) != 0))
                goto system_error;
        }
        /* Note the dacl field in *secdPP now points directly to daclP! */
    }

    /* Same as above but for SACLs */
    if (secd_control & SE_SACL_PRESENT) {
        if (ObjToPACLSWS(interp, objv[4], &saclP) != TCL_OK)
            goto error_return;
        if (saclP || (secd_control & SE_SACL_DEFAULTED)) {
            if (! SetSecurityDescriptorSacl(*secdPP, TRUE, saclP,
                                            (secd_control & SE_SACL_DEFAULTED) != 0))
                goto system_error;
        }
        /* Note the sacl field in *secdPP now points directly to saclP! */
    }
    
    return TCL_OK;

 system_error:
    TwapiReturnSystemError(interp);
    
 error_return:
    *secdPP = NULL;
    return TCL_ERROR;
}

/*
 * Returns a pointer to SWS memory containing a structure corresponding
 * to the given string representation.
 * Caller responsible for all SWS management.
 */
int ObjToPSECURITY_ATTRIBUTESSWS(
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

    *secattrPP = SWSAlloc (sizeof(**secattrPP), NULL);
    (*secattrPP)->nLength = sizeof(**secattrPP);

    if (ObjToInt(interp, objv[1], &inherit) == TCL_ERROR)
        goto error_return;
    (*secattrPP)->bInheritHandle = (inherit != 0);

    if (ObjToPSECURITY_DESCRIPTORSWS(interp, objv[0],
                                     &(SECURITY_DESCRIPTOR *)((*secattrPP)->lpSecurityDescriptor))
        == TCL_ERROR) {
        goto error_return;
    }

    return TCL_OK;

 error_return:
    *secattrPP = NULL;
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

    if (gBaseSettings.use_unicode_obj)
        return Tcl_NewUnicodeObj(ws, len);
    else
        return TwapiUtf8ObjFromUnicode(ws, len);
}

Tcl_Obj *ObjFromUnicode(const Tcl_UniChar *ws)
{
    if (ws == NULL)
        return ObjFromEmptyString(); /* TBD - log ? */

    if (gBaseSettings.use_unicode_obj)
        return Tcl_NewUnicodeObj(ws, -1);
    else
        return TwapiUtf8ObjFromUnicode(ws, -1);
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

TWAPI_EXTERN Tcl_Obj *ObjAllocateByteArray(int len, void **ppv)
{
    Tcl_Obj *objP;
    objP = ObjFromByteArray(NULL, len);
    if (ppv)
        *ppv = Tcl_GetByteArrayFromObj(objP, NULL);
    return objP;
}


/* Hex conversion copied from Tcl 86 since 85 does not have it. */
static const char HexDigits[16] = {
    '0', '1', '2', '3', '4', '5', '6', '7',
    '8', '9', 'a', 'b', 'c', 'd', 'e', 'f'
};
Tcl_Obj *ObjFromByteArrayHex(const unsigned char *bytes, int len)
{
    Tcl_Obj *resultObj = NULL;
    unsigned char *cursor = NULL;
    int offset = 0;

    resultObj = Tcl_NewObj();
    cursor = Tcl_SetByteArrayLength(resultObj, len * 2);
    for (offset = 0; offset < len; ++offset) {
	*cursor++ = HexDigits[((bytes[offset] >> 4) & 0x0f)];
	*cursor++ = HexDigits[(bytes[offset] & 0x0f)];
    }
    return resultObj;
}

unsigned char *ObjToByteArray(Tcl_Obj *objP, int *lenP)
{
    return Tcl_GetByteArrayFromObj(objP, lenP);
}

/*
 * RtlEncryptMemory (aka SystemFunction040) and
 * RtlDecryptMemory (aka SystemFunction041)
 * The Feb 2003 SDK does not define these in the headers, nor does it
 * include them in the export libs. So we have to dynamically load them
*/
typedef BOOLEAN (NTAPI *SystemFunction040_t)(PVOID, ULONG, ULONG);
MAKE_DYNLOAD_FUNC(SystemFunction040, advapi32, SystemFunction040_t)
typedef BOOLEAN (NTAPI *SystemFunction041_t)(PVOID, ULONG, ULONG);
MAKE_DYNLOAD_FUNC(SystemFunction041, advapi32, SystemFunction041_t)

/* Encrypt the given source bytes into the output buffer and length in *noutP.
 * If outP is NULL places required buffer size in *noutP.
 */
TCL_RESULT TwapiEncryptData(Tcl_Interp *interp, BYTE *inP, int nin, BYTE *outP, int *noutP)
{
    int sz, pad_len;
    NTSTATUS status;
    SystemFunction040_t fnP = Twapi_GetProc_SystemFunction040();

    if (fnP == NULL)
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);

    /*
     * Total length has to be multiple of encryption block size
     * so encryption involves padding. We will stick a byte
     * at the end to hold actual pad length
     */
#ifndef RTL_ENCRYPT_MEMORY_SIZE // Not defined in all SDK's
# define RTL_ENCRYPT_MEMORY_SIZE 8
#endif    
#define BLOCK_SIZE_MASK (RTL_ENCRYPT_MEMORY_SIZE-1)
    
    if (nin & BLOCK_SIZE_MASK) {
        /* Not a multiple of RTL_ENCRYPT_MEMORY_SIZE */
        sz = (nin + BLOCK_SIZE_MASK) & ~BLOCK_SIZE_MASK;
        pad_len = sz - nin;
    } else {
        /* Exact size. But we need a byte for the pad length field
           so need to add an entire block for that */
        pad_len = RTL_ENCRYPT_MEMORY_SIZE;
        sz = nin + pad_len;
    }
    TWAPI_ASSERT(pad_len > 0);
    TWAPI_ASSERT(pad_len <= RTL_ENCRYPT_MEMORY_SIZE);

    if (outP == NULL) {
        /* Caller just wants to know what size buffer is needed */
        *noutP = sz;
        return TCL_OK;
    }
            
    if (sz > *noutP)
        return TwapiReturnError(interp, TWAPI_BUFFER_OVERRUN);
    
    outP[sz-1] = (BYTE) pad_len;
    CopyMemory(outP, inP, nin);
    
    /* RtlEncryptMemory */
    status = fnP(outP, sz, 0);
    if (status != 0)
        return Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status));

    *noutP = sz;
    return TCL_OK;
}

/* Return a bytearray Tcl_Obj containing ciphertext of given source bytes */
Tcl_Obj *ObjEncryptBytes(Tcl_Interp *interp, void *pv, int nbytes)
{
    int nenc, nalloc;
    Tcl_Obj *objP = NULL;
    unsigned char *bytes = pv;
    WCHAR *wsP;

    /* Note data is always encrypted in Unicode form */
    
    nalloc = nbytes*sizeof(WCHAR);
    if (nalloc == 0)
        nalloc = sizeof(WCHAR); /* Can't alloc 0 */

    wsP = SWSPushFrame(nalloc, NULL);
    
    for (nenc = 0; nenc < nbytes; ++nenc)
        wsP[nenc] = bytes[nenc];
    objP = ObjEncryptUnicode(interp, wsP, nbytes);

    SWSPopFrame();
    return objP;
}

static TCL_RESULT TwapiDecryptBlockLengthError(Tcl_Interp *interp, int len)
{
    return TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                              Tcl_ObjPrintf("Invalid length (%d) of encrypted object. Must be non-zero multiple of block size (%d).", len, RTL_ENCRYPT_MEMORY_SIZE));
}

static TCL_RESULT TwapiDecryptPadLengthError(Tcl_Interp *interp, int pad_len)
{
    return TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                       Tcl_ObjPrintf("Invalid pad (%d) in decrypted object. Object corrupted or was not encrypted.", pad_len));
}

TCL_RESULT TwapiDecryptData(Tcl_Interp *interp, BYTE *encP, int nenc, BYTE *outP, int *noutP)
{
    int pad_len;
    NTSTATUS status;
    SystemFunction041_t fnP = Twapi_GetProc_SystemFunction041();

    if (fnP == NULL)
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    
    if (nenc == 0 || (nenc & BLOCK_SIZE_MASK))
        return TwapiDecryptPadLengthError(interp, nenc);

    /* Plaintext length is always less than ciphertext */
    if (outP == NULL) {
        *noutP = nenc; /* Max buffer size needed */
        return TCL_OK;
    }

    if (*noutP < nenc)
        return TwapiReturnError(interp, TWAPI_BUFFER_OVERRUN);

    CopyMemory(outP, encP, nenc);

    /* RtlDecryptMemory */
    status = fnP(outP, nenc, 0);
    if (status != 0) 
            return Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status));

    /* Last byte contains pad count */
    pad_len = outP[nenc-1];

    if (pad_len == 0 || pad_len > RTL_ENCRYPT_MEMORY_SIZE || pad_len > nenc)
        return TwapiDecryptPadLengthError(interp, pad_len);
    
    *noutP = nenc - pad_len;
    return TCL_OK;
}

/* Note: caller responsible for SWS maintenance */
BYTE *TwapiDecryptDataSWS(Tcl_Interp *interp, BYTE *encP, int nenc, int *noutP)
{
    TCL_RESULT res;
    int nout;
    BYTE *outP;
    
    res = TwapiDecryptData(interp, encP, nenc, NULL, &nout);
    if (res == TCL_OK) {
        outP = SWSAlloc(nout, NULL);
        res = TwapiDecryptData(interp, encP, nenc, outP, &nout);
        if (res == TCL_OK) {
            *noutP = nout;
            return outP;
        }
    }
    return NULL;
}

/* Encrypts the Unicode string rep to a byte array. We choose Unicode
   as the base format for strings because most API's need Unicode and makes it
   easier to SecureZeroMemory the buffer. 
*/
Tcl_Obj *ObjEncryptUnicode(Tcl_Interp *interp, WCHAR *uniP, int nchars)
{
    int nenc;
    Tcl_Obj *objP = NULL;
    TCL_RESULT res;
    
    if (nchars < 0)
        nchars = lstrlenW(uniP);

    res = TwapiEncryptData(interp, (BYTE *) uniP, nchars*sizeof(WCHAR), NULL, &nenc);
    if (res == TCL_OK) {
        unsigned char *encP;
        objP = ObjAllocateByteArray(nenc, &encP);
        res = TwapiEncryptData(interp, (BYTE *) uniP, nchars*sizeof(WCHAR), encP, &nenc);
        if (res == TCL_OK)
            Tcl_SetByteArrayLength(objP, nenc);
        else {
            ObjDecrRefs(objP);
            objP = NULL;
        }
    }
    return objP;
}

/*
 * Decrypts Unicode string in objP to SWS. Caller responsible for
 * SWS storage management and should also zero out the password
 * after use.
 * Returns length of string (in unichars, not bytes) NOT including terminating
 * char in *ncharsP
 */
WCHAR * ObjDecryptUnicodeSWS(Tcl_Interp *interp,
                          Tcl_Obj *objP,
                          int *ncharsP /* May be NULL */
    )
{
    int nenc, ndec;
    char *enc, *dec;
    TCL_RESULT res;
    
    enc = ObjToByteArray(objP, &nenc);
    
    res = TwapiDecryptData(interp, enc, nenc, NULL, &ndec);
    if (res != TCL_OK)
        return NULL;

    /* additional WCHAR is for terminating \0 */
    dec = SWSAlloc(ndec+sizeof(WCHAR), NULL);
    
    res = TwapiDecryptData(interp, enc, nenc, dec, &ndec);
    if (res != TCL_OK)
        return NULL;

    ndec /= sizeof(WCHAR);
    *(ndec + (WCHAR*) dec) = 0;             /* Terminate the Unicode string */
    if (ncharsP)
        *ncharsP = ndec;

    return (WCHAR*)dec;
}

/*
 * Decrypts encrypted objP to SWS in UTF8 encoding.
 * Caller responsible for SWS storage management and should also zero out 
 * the plaintext after use.
 * Returns length of string (in bytes) NOT including terminating
 * char in *ncharsP
 */
char *ObjDecryptUtf8SWS(Tcl_Interp *interp, Tcl_Obj *objP, int *nbytesP)
{
    WCHAR *uniP;
    char *utfP;
    int nuni, nutf;

    uniP = ObjDecryptUnicodeSWS(interp, objP, &nuni);
    if (uniP == NULL)
        return NULL;
    TWAPI_ASSERT(uniP[nuni] == 0); /* Should be terminated */

    /* REMEMBER we have to zero out uniP beyond this point */
    
    nutf = TwapiUnicharToUtf8(uniP, nuni, NULL, 0);
    if (nutf == -1)
        utfP = NULL;
    else {
        utfP = SWSAlloc(nutf+1, NULL); /* Additional byte for \0 */
        nutf = TwapiUnicharToUtf8(uniP, nuni, utfP, nutf);
        if (nutf == -1)
            utfP = NULL;
         else {
            utfP[nutf] = '\0';
            if (nbytesP)
                *nbytesP = nutf;
        }
    }
    if (utfP == NULL)
        TwapiReturnSystemError(interp);

    SecureZeroMemory(uniP, sizeof(WCHAR) * nuni);
    
    return utfP;
}

/*
 * Decrypts encrypted objP to SWS as a byte array.
 * Caller responsible for SWS storage management and should also zero out 
 * the plaintext after use.
 * Leaves nleading bytes space in front so decrypted data starts at
 * offset nleading bytes from returned pointer. *nbytesP is length of
 * decrypted data not including nleading.
 */
void *ObjDecryptBytesExSWS(Tcl_Interp *interp, Tcl_Obj *objP, int nleading, int *nbytesP)
{
    WCHAR *uniP, *from, *end;
    unsigned char *p, *to;
    int nuni, nalloc;

    uniP = ObjDecryptUnicodeSWS(interp, objP, &nuni);
    if (uniP == NULL)
        return NULL;

    /* REMEMBER we have to zero out uniP[] beyond this point */

    nalloc = nleading + nuni;
    if (nalloc == 0)
        nalloc = 1;
    p = SWSAlloc(nalloc, NULL);
    
    to = p + nleading;    /* Leave nleading bytes space in front */
    from = uniP;
    end = uniP + nuni;
    while (from < end)
        *to++ = (unsigned char) *from++;
    
    SecureZeroMemory(uniP, sizeof(WCHAR) * nuni);
    
    if (nbytesP)
        *nbytesP = nuni;
    return p;
}

/*
 * Like ObjDecryptUnicodeSWS but
 * if decryption fails, assumes password in unencrypted form and returns it.
 * Return data is on the SWS.
 * See ObjDecryptUnicodeSWS comments regarding SWS memory management.
 * ncharsP may be NULL
 */
WCHAR * ObjDecryptPasswordSWS(Tcl_Obj *objP, int *ncharsP)
{
    WCHAR *uniP, *toP;
    int nchars;

    toP =  ObjDecryptUnicodeSWS(NULL, objP, ncharsP);
    if (toP)
        return toP;
    /* Not encrypted, assume plaintext password */
    uniP = ObjToUnicodeN(objP, &nchars);
    toP = MemLifoCopy(SWS(), uniP, sizeof(WCHAR)*(nchars+1));
    if (ncharsP)
        *ncharsP = nchars;
    return toP;
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

void SecureZeroSEC_WINNT_AUTH_IDENTITY(PSEC_WINNT_AUTH_IDENTITY_W swaiP)
{
    int len;
    if (swaiP == NULL)
        return;
    if (swaiP->Password == NULL || swaiP->PasswordLength == 0)
        return;
    if (swaiP->Flags & SEC_WINNT_AUTH_IDENTITY_UNICODE)
        len = swaiP->PasswordLength * sizeof(WCHAR);
    else 
        len = swaiP->PasswordLength;
    SecureZeroMemory(swaiP->Password, len);
}


TCL_RESULT ParsePSEC_WINNT_AUTH_IDENTITY (
    TwapiInterpContext *ticP,
    Tcl_Obj *authObj,
    SEC_WINNT_AUTH_IDENTITY_W **swaiPP
    )
{
    Tcl_Obj *passwordObj;
    Tcl_Obj **objv;
    int objc;
    TCL_RESULT res;
    SEC_WINNT_AUTH_IDENTITY_W *swaiP;
    
    if ((res = ObjGetElements(ticP->interp, authObj, &objc, &objv)) != TCL_OK)
        return res;

    if (objc == 0) {
        *swaiPP = NULL;
        return TCL_OK;
    }

    swaiP = MemLifoAlloc(ticP->memlifoP, sizeof(*swaiP), NULL);
    swaiP->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETWSTRN(swaiP->User, swaiP->UserLength),
                         GETWSTRN(swaiP->Domain, swaiP->DomainLength),
                         GETOBJ(passwordObj),
                         ARGEND);
    if (res != TCL_OK)
        return res;

    /* The decrypted password will be on the SWS which should be the same
       as ticP->memlifoP
    */
    TWAPI_ASSERT(SWS() == ticP->memlifoP);
    swaiP->Password = ObjDecryptPasswordSWS(passwordObj, &swaiP->PasswordLength);

    *swaiPP = swaiP;
    return TCL_OK;
}


