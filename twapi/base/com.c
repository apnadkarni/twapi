/* 
 * Copyright (c) 2006-2009, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 *
 * Utility functions used by the COM module
 */

#include "twapi.h"

/*
 * Struct for mapping VARTYPE values to strings.
 * We search linearly so order based on most likely types.
 * Only the basic types are covered in this table. The code itself
 * handles the special/complex cases.
 */
struct vt_token_pair {
    VARTYPE vt;
    char   *tok;
};
static struct vt_token_pair vt_base_tokens[] = {
    {VT_BOOL, "bool"},
    {VT_I2, "i2"},
    {VT_I4, "i4"},
    {VT_PTR, "ptr"},
    {VT_R4, "r4"},
    {VT_R8, "r8"},
    {VT_CY, "cy"},
    {VT_DATE, "date"},
    {VT_BSTR, "bstr"},
    {VT_DISPATCH, "idispatch"},
    {VT_ERROR, "error"},
    {VT_VARIANT, "variant"},
    {VT_UNKNOWN, "iunknown"},
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
    {VT_USERDEFINED, "userdefined"}
};

/*
 * Event sink definitions
 */
HRESULT STDMETHODCALLTYPE Twapi_EventSink_QueryInterface(
    IDispatch *this,
    REFIID riid,
    void **ifcPP);
ULONG STDMETHODCALLTYPE Twapi_EventSink_AddRef(IDispatch *this);
ULONG STDMETHODCALLTYPE Twapi_EventSink_Release(IDispatch *this);
HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetTypeInfoCount(
    IDispatch *this,
    UINT *pctP);
HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetTypeInfo(
    IDispatch *this,
    UINT tinfo,
    LCID lcid,
    ITypeInfo **tiPP);
HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetIDsOfNames(
    IDispatch *this,
    REFIID   riid,
    LPOLESTR *namesP,
    UINT namesc,
    LCID lcid,
    DISPID *rgDispId);
HRESULT STDMETHODCALLTYPE Twapi_EventSink_Invoke(
    IDispatch *this,
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD flags,
    DISPPARAMS *dispparamsP,
    VARIANT *resultvarP,
    EXCEPINFO *excepP,
    UINT *argErrP);


/* Vtbl for Twapi_EventSink */
static struct IDispatchVtbl Twapi_EventSink_Vtbl = {
    Twapi_EventSink_QueryInterface,
    Twapi_EventSink_AddRef,
    Twapi_EventSink_Release,
    Twapi_EventSink_GetTypeInfoCount,
    Twapi_EventSink_GetTypeInfo,
    Twapi_EventSink_GetIDsOfNames,
    Twapi_EventSink_Invoke
};


/* TBD - does this (related methods) need to be made thread safe? Right now
 * we do not support multithreaded Tcl
 */
typedef struct Twapi_EventSink {
    interface IDispatch idispP; /* Must be first field */
    IID iid;                    /* IID for this event sink interface */
    int refc;                   /* Ref count */
    Tcl_Interp *interp;
#define  MAX_EVENTSINK_CMDARGS 16
    Tcl_Obj *cmd;               /* Stores the callback command arg list */

} Twapi_EventSink;

static void TwapiInvalidVariantTypeMessage(Tcl_Interp *interp, VARTYPE vt)
{
    char buf[80];
    if (interp) {
        StringCbPrintfA(buf, sizeof(buf),
                       "Invalid or unsupported VARTYPE (%d)",
                       vt);
        Tcl_SetResult(interp, buf, TCL_VOLATILE);
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
        Tcl_SetObjResult(interp, objP);
    }

    return TCL_ERROR;
}

static HRESULT STDMETHODCALLTYPE Twapi_EventSink_QueryInterface(
    IDispatch *this,
    REFIID riid,
    void **ifcPP)
{

    if (!IsEqualIID(riid, &((Twapi_EventSink *)this)->iid) &&
        !IsEqualIID(riid, &IID_IUnknown) &&
        !IsEqualIID(riid, &IID_IDispatch)) {
        /* Not a supported interface */
        *ifcPP = NULL;
        return E_NOINTERFACE;
    }

    this->lpVtbl->AddRef(this);
    *ifcPP = this;
    return S_OK;
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
    if (Tcl_GetIntFromObj(NULL, obj, &i) == TCL_OK) {
        *vtP = (VARTYPE) i;
        return TCL_OK;
    } else if (LookupBaseVTToken(interp, Tcl_GetString(obj), vtP) == TCL_OK) {
        return TCL_OK;
    }

    /*
     * See if it's a list. Note interp contains an error msg at this point
     */

    if (Tcl_ListObjGetElements(NULL, obj, &objc, &objv) != TCL_OK ||
        objc < 2) {
        return TCL_ERROR;
    }
    if (Tcl_GetIntFromObj(NULL, objv[0], &i) == TCL_OK) {
        vt = (VARTYPE) i;
    } else if (LookupBaseVTToken(NULL, Tcl_GetString(objv[0]), &vt) != TCL_OK) {
        return TCL_ERROR;
    }

    /* vt must be either pointer, array or UDT in the list case */
    if (vt == VT_PTR || vt == VT_SAFEARRAY || vt == VT_USERDEFINED) {
        *vtP = vt;
        Tcl_ResetResult(interp); // Get rid of old error message.
        return TCL_OK;
    }
    else
        return TCL_ERROR;
}

Tcl_Obj *ObjFromCONNECTDATA (const CONNECTDATA *cdP)
{
    Tcl_Obj *objv[2];
    objv[0] = ObjFromIUnknown(cdP->pUnk);
    objv[1] = Tcl_NewLongObj(cdP->dwCookie);
    return Tcl_NewListObj(2, objv);
}


/*
 * Return a Tcl object that is a list
 * {"safearray" dimensionlist VT_xxx valuelist}.
 * dimensionlist is a flat list of lowbound, upperbound pairs, one
 * for each dimension.
 * If VT_xxx is not recognized, valuelist is missing
 * If there is no vartype information, VT_XXX is also missing
 * Never returns NULL.
 */
static Tcl_Obj *ObjFromSAFEARRAY(SAFEARRAY *arrP)
{
    Tcl_Obj *objv[3];           /* "safearray|vt", dimensions,  value */
    int      objc;
    long     i;
    VARTYPE  vt;
    HRESULT  hr;
    long     num_elems;
    void     *valP;
#define GETVAL(index_, type_) (((type_ *)valP)[index_])

    /* We require the safearray to have a type associated */
    objc = 1;
    if (SafeArrayGetVartype(arrP, &vt) == S_OK) {
        objv[0] = Tcl_NewIntObj(vt|VT_ARRAY);
    } else {
        objv[0] = Tcl_NewIntObj(VT_ARRAY);
        goto alldone;
    }

    hr = SafeArrayAccessData(arrP, &valP);
    if (FAILED(hr))
        goto alldone;

    objv[1] = Tcl_NewListObj(0, NULL);
    num_elems = 1;
    for (i = 0; i < arrP->cDims; ++i) {
        Tcl_ListObjAppendElement(NULL, objv[1], Tcl_NewLongObj(arrP->rgsabound[i].lLbound));
        Tcl_ListObjAppendElement(NULL, objv[1], Tcl_NewLongObj(arrP->rgsabound[i].cElements));
        num_elems *= arrP->rgsabound[i].cElements;
    }

    /* TBD - it might be more efficient to allocate an array and then
       use Tcl_NewListObj to create from there instead of calling
       Tcl_ListObjAppend for each element
    */

    objv[2] = Tcl_NewListObj(0, NULL); /* Value object */
    objc = 3;

    switch (vt) {
    case VT_EMPTY:
    case VT_NULL:
        break;

    case VT_I2: {
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewIntObj(GETVAL(i,short)));
        }
        break;
    }

    case VT_INT: /* FALLTHROUGH */
    case VT_I4:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewLongObj(GETVAL(i,long)));
        }
        break;

    case VT_R4:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewDoubleObj(GETVAL(i,float)));
        }
        break;

    case VT_R8:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewDoubleObj(GETVAL(i,double)));
        }
        break;

    case VT_CY:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                ObjFromCY(&(((CY *)valP)[i]))
                );
        }
        break;

    case VT_DATE:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                Tcl_NewDoubleObj(GETVAL(i,double)));
        }
        break;

    case VT_BSTR:
        for (i = 0; i < num_elems; ++i) {
            BSTR bstr = GETVAL(i,BSTR);
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                Tcl_NewUnicodeObj(bstr, SysStringLen(bstr))
                );
        }
        break;

    case VT_DISPATCH:
        for (i = 0; i < num_elems; ++i) {
            IDispatch *idispP = GETVAL(i,IDispatch *);
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                ObjFromIDispatch(idispP)
                );
        }
        break;

    case VT_ERROR:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewIntObj(GETVAL(i,SCODE)));
        }
        break;

    case VT_BOOL:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                Tcl_NewBooleanObj(GETVAL(i,VARIANT_BOOL))
                );
        }
        break;

    case VT_VARIANT:
        for (i = 0; i < num_elems; ++i) {
            VARIANT *varP = &((( VARIANT *)valP)[i]);
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                ObjFromVARIANT(varP, 0));
        }
        break;

    case VT_DECIMAL:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                ObjFromDECIMAL(&((( DECIMAL *)valP)[i]))
                );
        }
        break;

    case VT_UNKNOWN:
        for (i = 0; i < num_elems; ++i) {
            IUnknown *idispP = GETVAL(i, IUnknown *);
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                ObjFromIUnknown(idispP));
        }
        break;

    case VT_I1:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewIntObj(GETVAL(i,char)));
        }
        break;

    case VT_UI1:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewIntObj(GETVAL(i,unsigned char)));
        }
        break;

    case VT_UI2:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewIntObj(GETVAL(i,unsigned short)));
        }
        break;

    case VT_UINT: /* FALLTHROUGH */
    case VT_UI4:
        for (i = 0; i < num_elems; ++i) {
            unsigned long ulval = GETVAL(i, unsigned long);
            /* store as wide integer if it does not fit in signed 32 bits */
            Tcl_ListObjAppendElement(
                NULL, objv[2],
                (ulval & 0x80000000) ? Tcl_NewWideIntObj(ulval) : Tcl_NewLongObj(ulval));
        }
        break;

    case VT_I8: /* FALLTHRU */
    case VT_UI8:
        for (i = 0; i < num_elems; ++i) {
            Tcl_ListObjAppendElement(NULL, objv[2],
                                     Tcl_NewWideIntObj(GETVAL(i,__int64)));
        }
        break;

        /* Dunno how to handle these */
    default:
        break;

    }

    SafeArrayUnaccessData(arrP);

 alldone:
    return Tcl_NewListObj(objc, objv);
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
            return ObjFromSAFEARRAY(*(varP->pparray));
        else
            return ObjFromSAFEARRAY(varP->parray);
    }
    if ((V_VT(varP) == (VT_BYREF|VT_VARIANT)) && varP->pvarVal)
        return ObjFromVARIANT(varP->pvarVal, 0);

    objv[0] = Tcl_NewIntObj(V_VT(varP) & ~VT_BYREF);
    objv[1] = NULL;

    switch (V_VT(varP)) {
    case VT_EMPTY|VT_BYREF:
    case VT_EMPTY:
    case VT_NULL|VT_BYREF:
    case VT_NULL:
        break;

    case VT_I2|VT_BYREF:
    case VT_I2:
        objv[1] = Tcl_NewIntObj(V_VT(varP) == VT_I2 ? V_I2(varP) : * V_I2REF(varP));
        break;

    case VT_I4|VT_BYREF:
    case VT_I4:
        objv[1] = Tcl_NewIntObj(V_VT(varP) == VT_I4 ? V_I4(varP) : * V_I4REF(varP));
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
            objv[1] = Tcl_NewUnicodeObj(V_BSTR(varP),
                                        SysStringLen(V_BSTR(varP)));
        else
            objv[1] = Tcl_NewUnicodeObj(* V_BSTRREF(varP),
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
        objv[1] = Tcl_NewIntObj(V_VT(varP) == VT_ERROR ? V_ERROR(varP) : * V_ERRORREF(varP));
        break;

    case VT_BOOL|VT_BYREF:
    case VT_BOOL:
        objv[1] = Tcl_NewBooleanObj(V_VT(varP) == VT_BOOL ? V_BOOL(varP) : * V_BOOLREF(varP));
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
        objv[1] = Tcl_NewIntObj(V_VT(varP) == VT_I1 ? V_I1(varP) : * V_I1REF(varP));
        break;

    case VT_UI1|VT_BYREF:
    case VT_UI1:
        objv[1] = Tcl_NewIntObj(V_VT(varP) == VT_UI1 ? V_UI1(varP) : * V_UI1REF(varP));
        break;

    case VT_UI2|VT_BYREF:
    case VT_UI2:
        objv[1] = Tcl_NewIntObj(V_VT(varP) == VT_UI2 ? V_UI2(varP) : * V_UI2REF(varP));
        break;

    case VT_UI4|VT_BYREF:
    case VT_UI4:
        /* store as wide integer if it does not fit in signed 32 bits */
        ulval = V_VT(varP) == VT_UI4 ? V_UI4(varP) : * V_UI4REF(varP);
        if (ulval & 0x80000000) {
            objv[1] = Tcl_NewWideIntObj(ulval);
        }
        else {
            objv[1] = Tcl_NewLongObj(ulval);
        }
        break;

    case VT_I8|VT_BYREF:
    case VT_I8:
        objv[1] = Tcl_NewWideIntObj(V_VT(varP) == VT_I8 ? V_I8(varP) : * V_I8REF(varP));
        break;

    case VT_UI8|VT_BYREF:
    case VT_UI8:
        objv[1] = Tcl_NewWideIntObj(V_VT(varP) == VT_UI8 ? V_UI8(varP) : * V_UI8REF(varP));
        break;


    case VT_INT|VT_BYREF:
    case VT_INT:
        objv[1] = Tcl_NewIntObj(V_VT(varP) == VT_INT ? V_INT(varP) : * V_INTREF(varP));
        break;

    case VT_UINT|VT_BYREF:
    case VT_UINT:
        /* store as wide integer if it does not fit in signed 32 bits */
        ulval = V_VT(varP) == VT_UINT ? V_UINT(varP) : * V_UINTREF(varP);
        if (ulval & 0x80000000) {
            objv[1] = Tcl_NewWideIntObj(ulval);
        }
        else {
            objv[1] = Tcl_NewLongObj(ulval);
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
            objv[1] = Tcl_NewListObj(2, valObj);
        }
        break;

        /* Dunno how to handle these */
    case VT_RECORD|VT_BYREF:
    case VT_VARIANT: /* Note VT_VARIANT is illegal */
    default:
        break;
    }

    if (value_only)
        return objv[1] ? objv[1] : Tcl_NewStringObj("", 0);
    else
        return Tcl_NewListObj(objv[1] ? 2 : 1, objv);
}


/* Returns a Tcl_Obj corresponding to a type descriptor. */
static Tcl_Obj *ObjFromTYPEDESC(Tcl_Interp *interp, TYPEDESC *tdP, ITypeInfo *tiP)
{
    Tcl_Obj *objv[3];
    int      objc;
    int      i;
#if 0
    ITypeInfo *utiP; /* For userdefined typeinfo */
#endif

    if (tdP == NULL) {
        if (interp) {
            Tcl_SetResult(interp, "Internal error: ObjFromTYPEDESC: NULL TYPEDESC pointer", TCL_STATIC);
        }
        return NULL;
    }

    switch(tdP->vt) {
        // VARIANT/VARIANTARG compatible types
    case VT_PTR:   /* FALL THROUGH */
    case VT_SAFEARRAY:
        objv[1] = ObjFromTYPEDESC(interp, tdP->lptdesc, tiP); /* Recurse */
        if (objv[1] == NULL)
            return NULL;
        objc = 2;
        break;

    case VT_CARRAY:
        objv[1] = ObjFromTYPEDESC(interp, &tdP->lpadesc->tdescElem, tiP);
        if (objv[1] == NULL)
            return NULL;
        objv[2] = Tcl_NewListObj(0, NULL); /* For dimension info */
        for (i = 0; i < tdP->lpadesc->cDims; ++i) {
            Tcl_ListObjAppendElement(interp, objv[2],
                                     Tcl_NewIntObj(tdP->lpadesc->rgbounds[i].lLbound));
            Tcl_ListObjAppendElement(interp, objv[2],
                                     Tcl_NewIntObj(tdP->lpadesc->rgbounds[i].cElements));
        }
        objc = 3;
        break;

    case VT_USERDEFINED:
        // Original code used to resolve the name. However, this was not
        // very useful by itself so we now leave it up to the Tcl level
        // to do so if required.
#if 1
        objv[1] = Tcl_NewIntObj(tdP->hreftype);
#else
        objv[1] = NULL;
        if (tiP->lpVtbl->GetRefTypeInfo(tiP, tdP->hreftype, &utiP) == S_OK) {
            BSTR bstr;
            if (utiP->lpVtbl->GetDocumentation(utiP, MEMBERID_NIL, &bstr, NULL, NULL, NULL) == S_OK) {
                objv[1] = Tcl_NewUnicodeObj(bstr, SysStringLen(bstr));
                SysFreeString(bstr);
            }
            utiP->lpVtbl->Release(utiP);
        }
        if (objv[1] == NULL) {
            /* Could not get name of custom type. */
            objv[1] = Tcl_NewStringObj(NULL, 0);
        }
#endif
        objc = 2;
        break;

    default:
        objc = 1;
        break;
    }

    objv[0] = Tcl_NewIntObj(tdP->vt);
    return Tcl_NewListObj(objc, objv);
}

static Tcl_Obj *ObjFromPARAMDESC(Tcl_Interp *interp, PARAMDESC *vdP, ITypeInfo *tiP)
{
    Tcl_Obj *objv[2];

    objv[0] = Tcl_NewIntObj(vdP->wParamFlags);
    objv[1] = NULL;
    if ((vdP->wParamFlags & PARAMFLAG_FOPT) &&
        (vdP->wParamFlags & PARAMFLAG_FHASDEFAULT) &&
        vdP->pparamdescex) {
        objv[1] = ObjFromVARIANT(&vdP->pparamdescex->varDefaultValue, 0);
    }

    return Tcl_NewListObj(objv[1] ? 2 : 1, objv);
}

static Tcl_Obj *ObjFromVARDESC(Tcl_Interp *interp, VARDESC *vdP, ITypeInfo *tiP)
{
    Tcl_Obj *resultObj = Tcl_NewListObj(0, 0);
    Tcl_Obj *obj;

    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, vdP, memid);
    if (vdP->lpstrSchema) {
        Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, resultObj, vdP, lpstrSchema);
    }
    else {
        Tcl_ListObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("lpstrSchema"));
        Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewStringObj("", 0));
    }

    switch (vdP->varkind) {
    case VAR_PERINSTANCE: /* FALLTHROUGH */
    case VAR_DISPATCH: /* FALLTHROUGH */
    case VAR_STATIC:
        Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, vdP, oInst);
        break;
    case VAR_CONST:
        Tcl_ListObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("lpvarValue"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromVARIANT(vdP->lpvarValue, 0));
        break;
    }

    Tcl_ListObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("elemdescVar.tdesc"));
    obj = ObjFromTYPEDESC(interp, &vdP->elemdescVar.tdesc, tiP);
    Tcl_ListObjAppendElement(interp, resultObj,
                             obj ? obj : Tcl_NewListObj(0, NULL));

    Tcl_ListObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("elemdescVar.paramdesc"));
    obj = ObjFromPARAMDESC(interp, &vdP->elemdescVar.paramdesc, tiP);
    Tcl_ListObjAppendElement(interp, resultObj,
                             obj ? obj : Tcl_NewListObj(0, NULL));

    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, vdP, varkind);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, vdP, wVarFlags);

    return resultObj;
}


static Tcl_Obj *ObjFromFUNCDESC(Tcl_Interp *interp, FUNCDESC *fdP, ITypeInfo *tiP)
{
    Tcl_Obj *resultObj = Tcl_NewListObj(0, 0);
    Tcl_Obj *obj;
    int      i;

    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, memid);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, funckind);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, invkind);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, callconv);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, cParams);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, cParamsOpt);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, oVft);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, fdP, wFuncFlags);
    Tcl_ListObjAppendElement(interp, resultObj,
                             STRING_LITERAL_OBJ("elemdescFunc.tdesc"));
    obj = ObjFromTYPEDESC(interp, &fdP->elemdescFunc.tdesc, tiP);
    Tcl_ListObjAppendElement(interp, resultObj,
                             obj ? obj : Tcl_NewListObj(0, NULL));

    Tcl_ListObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("elemdescFunc.paramdesc"));
    obj = ObjFromPARAMDESC(interp, &fdP->elemdescFunc.paramdesc, tiP);
    Tcl_ListObjAppendElement(interp, resultObj,
                             obj ? obj : Tcl_NewListObj(0, NULL));

    /* List of possible return codes */
    obj = Tcl_NewListObj(0, NULL);
    if (fdP->lprgscode) {
        for (i = 0; i < fdP->cScodes; ++i) {
            Tcl_ListObjAppendElement(interp, resultObj,
                                     Tcl_NewIntObj(fdP->lprgscode[i]));
        }
    }
    Tcl_ListObjAppendElement(interp, resultObj,
                             STRING_LITERAL_OBJ("lprgscode"));
    Tcl_ListObjAppendElement(interp, resultObj, obj);

    /* List of parameter descriptors */
    obj = Tcl_NewListObj(0, NULL);
    if (fdP->lprgelemdescParam) {
        for (i = 0; i < fdP->cParams; ++i) {
            Tcl_Obj *paramObj[2];
            paramObj[0] =
                ObjFromTYPEDESC(interp,
                                &(fdP->lprgelemdescParam[i].tdesc), tiP);
            if (paramObj[0] == NULL)
                paramObj[0] = Tcl_NewListObj(0, 0);
            paramObj[1] =
                ObjFromPARAMDESC(interp,
                                 &(fdP->lprgelemdescParam[i].paramdesc), tiP);
            if (paramObj[1] == NULL)
                paramObj[1] = Tcl_NewListObj(0, 0);
            Tcl_ListObjAppendElement(interp, obj, Tcl_NewListObj(2, paramObj));
        }
    }
    Tcl_ListObjAppendElement(interp, resultObj,
                             STRING_LITERAL_OBJ("lprgelemdescParam"));
    Tcl_ListObjAppendElement(interp, resultObj, obj);

    return resultObj;
}


static ULONG STDMETHODCALLTYPE Twapi_EventSink_AddRef(IDispatch *this)
{
    ((Twapi_EventSink *)this)->refc += 1;
    return ((Twapi_EventSink *)this)->refc;
}

static ULONG STDMETHODCALLTYPE Twapi_EventSink_Release(IDispatch *this)
{
    Twapi_EventSink *me = (Twapi_EventSink *) this;

    me->refc -= 1;
    if (((Twapi_EventSink *)this)->refc == 0) {
        if (me->interp)
            Tcl_Release(me->interp);
        if (me->cmd)
            Tcl_DecrRefCount(me->cmd);
        TwapiFree(this);
        return 0;
    } else
        return ((Twapi_EventSink *)this)->refc;
}

static HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetTypeInfoCount
(
    IDispatch *this,
    UINT *pctP
)
{
    /* We do not provide type information */
    if (pctP)
        *pctP = 0;
    return S_OK;
}


static HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetTypeInfo(
    IDispatch *this,
    UINT tinfo,
    LCID lcid,
    ITypeInfo **tiPP)
{
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetIDsOfNames(
    IDispatch *this,
    REFIID   riid,
    LPOLESTR *namesP,
    UINT namesc,
    LCID lcid,
    DISPID *rgDispId)
{
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Twapi_EventSink_Invoke(
    IDispatch *this,
    DISPID dispid,
    REFIID riid,
    LCID lcid,
    WORD flags,
    DISPPARAMS *dispparamsP,
    VARIANT *retvarP,
    EXCEPINFO *excepP,
    UINT *argErrP)
{
    Twapi_EventSink *me = (Twapi_EventSink *) this;
    int     i;
    HRESULT hr;
    Tcl_Obj **cmdobjv;
    Tcl_Obj **cmdprefixv;
    int     cmdobjc;
    Tcl_SavedResult savedresult;

    if (me == NULL)
        return E_POINTER;

    if (Tcl_ListObjGetElements(NULL, me->cmd, &cmdobjc, &cmdprefixv) != TCL_OK) {
        /* Internal error - should not happen. Should we log background error?*/
        return E_FAIL;
    }

    /* Note we  will tack on 3 additional fixed arguments plus dispparms */
    /* TBD - where is this freed ? */
    cmdobjv = TwapiAlloc((cmdobjc+4) * sizeof(*cmdobjv));
    
    for (i = 0; i < cmdobjc; ++i) {
        cmdobjv[i] = cmdprefixv[i];
    }

    cmdobjv[cmdobjc] = Tcl_NewLongObj(dispid);
    cmdobjv[cmdobjc+1] = Tcl_NewLongObj(lcid);
    cmdobjv[cmdobjc+2] = Tcl_NewIntObj(flags);
    cmdobjc += 3;

    /* Add the passed parameters */
    cmdobjv[cmdobjc] = Tcl_NewListObj(0, NULL);
    if (dispparamsP) {
        /* Note parameters are in reverse order */
        for (i = dispparamsP->cArgs - 1; i >= 0 ; --i) {
            Tcl_ListObjAppendElement(
                NULL,
                cmdobjv[cmdobjc],
                ObjFromVARIANT(&dispparamsP->rgvarg[i], 0)
                );
        }
    }
    ++cmdobjc;
                 

    for (i = 0; i < cmdobjc; ++i) {
        Tcl_IncrRefCount(cmdobjv[i]); /* Protect while we are using it.
                                         Required by Tcl_EvalObjv */
    }

    /*
     * Before eval'ing, addref ourselves so we don't get deleted in a
     * recursive callback
     */
    this->lpVtbl->AddRef(this);

    /* TBD - is this safe as we are being called from the message dispatch
       loop? Or should we queue to pending callback queue ? But in that
       case we cannot get results back as we can't block in this thread
       as the script invocation will also be in this thread. Also, is
       the Tcl_SaveResult/RestoreResult really necessary ?
       Note tclWinDde also evals in this fashion.
    */
    /* If hr is not TCL_OK, it is a HRESULT error code */
    Tcl_SaveResult(me->interp, &savedresult);
    hr = Tcl_EvalObjv(me->interp, cmdobjc, cmdobjv, TCL_EVAL_GLOBAL);
    if (hr != TCL_OK) {
        Tcl_BackgroundError(me->interp);
        if (excepP) {
            ZeroMemory(excepP, sizeof(*excepP));
            excepP->scode = hr;
        }
    } else {
        /* TBD - appropriately init retvarP from Tcl_GetObjResult keeping
         * in mind that the retvarP by be BYREF as well.
         */
        if (retvarP)
            VariantInit(retvarP);
        hr = S_OK;
    }
    Tcl_RestoreResult(me->interp, &savedresult);

    /* Free the objects we allocated */
    for (i = 0; i < cmdobjc; ++i) {
        Tcl_DecrRefCount(cmdobjv[i]);
    }

    /* Undo the AddRef we did before */
    this->lpVtbl->Release(this);
    /* this/me may be invalid at this point! Make sure we don't access them */
    this = NULL;
    me = NULL;

    return hr;
}

/*
 * Called from a script create a event sink. Returns the IDispatch interface
 * that can be used as an event sink.
 * 
 */
int Twapi_ComEventSinkObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Twapi_EventSink *sinkP;
    IID iid;
    HRESULT hr;

    if (objc != 3) {
        Tcl_WrongNumArgs(interp, 1, objv, "IID CMD");
        return TCL_ERROR;
    }
    
    hr = IIDFromString(Tcl_GetUnicode(objv[1]), &iid);
    if (FAILED(hr)) {
        Twapi_AppendSystemError(interp, hr);
        return TCL_ERROR;
    }

    /* This is the sink object. Memory is freed when the object is released */
    sinkP = TwapiAlloc(sizeof(*sinkP));

    /* Fill in the cmdargs slots from the arguments */
    sinkP->idispP.lpVtbl = &Twapi_EventSink_Vtbl;
    sinkP->iid = iid;
    Tcl_Preserve(interp);   /* TBD - inappropriate use of interp, use ticP */
    sinkP->interp = interp;
    sinkP->refc = 1;
    Tcl_IncrRefCount(objv[2]);
    sinkP->cmd = objv[2];

    Tcl_SetObjResult(interp, ObjFromIUnknown(sinkP));


    return TCL_OK;
}




int Twapi_IDispatch_InvokeObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    IDispatch *idispP;
    LCID       lcid;
    WORD       flags;
    int        i, j;
    DISPID     dispid;
    DISPPARAMS dispparams;
    Tcl_Obj  **params;
    int        nparams;
    DISPID     named_dispid;
    VARIANT   *dispargP = NULL;
    USHORT    *paramflagsP = NULL;
    int        nargalloc;
    VARTYPE    retvar_vt;
    EXCEPINFO  einfo;
    UINT       badarg_index;
    HRESULT    hr = S_OK;
    int        status = TCL_ERROR;
    Tcl_Obj  **protov;          // Prototype list
    int        protoc;          // Prototype element count


    /* objv[] contains
     *   0 - command name
     *   1 - IDispatch pointer
     *   2 - Prototype
     *   remaining - parameter values
     *
     * objv[2] is the prototype for the command which itself is a list
     *   0 - DISPID
     *   1 - LCID
     *   2 - flags
     *   3 - return type
     *   4 - (optional) list of parameters to dispatch function. This is a list
     *       of elements of the form {type paramflags ?defaultvalue?}. If this
     *       element is missing (as opposed to empty), it means no parameter
     *       info is available and we will go strictly by the arguments
     *       present.
     */

    if (objc < 3) {
        Tcl_WrongNumArgs(interp, 1, objv, "IDISPATCH PROTOTYPE ?ARG1 ARG2...?");
        return TCL_ERROR;
    }

    if (ObjToIDispatch(interp, objv[1], &idispP) != TCL_OK)
        return TCL_ERROR;

    if (Tcl_ListObjGetElements(interp, objv[2], &protoc, &protov) != TCL_OK)
        return TCL_ERROR;

    /* Extract prototype information */
#if 1
    if (TwapiGetArgs(interp, protoc, protov,
                     GETINT(dispid), GETINT(lcid),
                     GETWORD(flags), GETVAR(retvar_vt, ObjToVT),
                     ARGTERM) != TCL_OK) {
        Tcl_SetResult(interp, "Invalid IDispatch prototype - must contain DISPID LCID FLAGS RETTYPE ?PARAMTYPES?", TCL_STATIC);
        return TCL_ERROR;
    }
#else
    if (protoc < 5) {
        Tcl_SetResult(interp, "Invalid IDispatch prototype - must contain DISPID LCID FLAGS RETTYPE ?PARAMTYPES?", TCL_STATIC);
        return TCL_ERROR;
    }

    if (Tcl_GetLongFromObj(interp, protov[0], &dispid) != TCL_OK)
        return TCL_ERROR;

    /* Note We do not care about RIID in protov[1] */

    if (Tcl_GetLongFromObj(interp, protov[2], &lcid) != TCL_OK)
        return TCL_ERROR;
    
    if (Tcl_GetIntFromObj(interp, protov[3], &i) != TCL_OK)
        return TCL_ERROR;
    flags = (WORD) i;

    if (ObjToVT(interp, protov[4], &retvar_vt) != TCL_OK)
        return TCL_ERROR;
#endif

    if (protoc >= 5) {
        /* Extract the parameter information */
        if (Tcl_ListObjGetElements(interp, protov[4], &nparams, &params) != TCL_OK)
            return TCL_ERROR;
    } else {
        /* No parameter information available. Base count on number of
         * arguments provided
         */
        params = NULL;
        nparams = objc - 3;
    }

    /* Initialize parameter structures */
    dispparams.cArgs = nparams;
    /*
     * Allocate space for params and return value and initialize it.
     * First element in the array is the return value.
     * For each param, we allocate two structs to take care of the case
     * where the param is by reference.
     */
    nargalloc = (1+(2*nparams));
    dispargP = MemLifoPushFrame(&ticP->memlifo,
                                nargalloc * sizeof(*dispargP),
                                NULL);
    paramflagsP = MemLifoAlloc(&ticP->memlifo,
                               (nparams+1)*sizeof(*paramflagsP),
                               NULL);

    /* Init all so they're all valid in case we take an early error exit
     *   and have to clear them before return
     */
    for (i = 0; i < nargalloc; ++i)
        V_VT(&dispargP[i]) = VT_EMPTY;
    dispparams.rgvarg = nparams ? &dispargP[1] : NULL;

    /* Init param structures */
    named_dispid  = DISPID_PROPERTYPUT;
    if(flags & (DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF)) {
        if (nparams != 1) {
            /* TBD - not sure if this is the case. What about indexed props ? */
            Tcl_SetResult(interp, "Property put methods must have exactly one parameter", TCL_STATIC);
            goto vamoose;
        }
        dispparams.cNamedArgs = 1;
        dispparams.rgdispidNamedArgs = &named_dispid;
        retvar_vt = VT_VOID;    /* Property put never has a return value */
    } else {
        dispparams.cNamedArgs = 0;
        dispparams.rgdispidNamedArgs = NULL;
    }

    /*
     * Parameters go in reverse order in allocated array occupying
     * positions from 1 to nparams. Note pos 0 is return value so
     * params go from 1 to nparams in dispargP array. Elements nparams+j to
     * the end are used as targets for element j in case the parameter
     * is a by ref.
     *
     * The parameter values supplied by caller, if any, are in
     * objv[3] onward.
     */
    for (i=0, j=nparams; j; ++i, --j) {
        VariantInit(&dispargP[j]);
        VariantInit(&dispargP[j+nparams]); /* byref variant if needed */
        if (TwapiMakeVariantParam(interp,
                                  params ? params[i] : NULL,
                                  &dispargP[j],
                                  &dispargP[j+nparams],
                                  &paramflagsP[j],
                                  ((3+i) >= objc ? NULL : objv[3+i])
                ) != TCL_OK)
            goto vamoose;
    }
    
    /* Init exception structure for error handling */
    ZeroMemory(&einfo, sizeof(einfo));
    badarg_index = (UINT) -1;

    /*
     * Now, we're ready to invoke. The Invoke might call us back (futures?)
     * so do we need to do AddRef/Release pair? Well, right now we don't
     * access idisP again so we don't do that for now.
     */

    hr = idispP->lpVtbl->Invoke(idispP,
                                dispid,
                                &IID_NULL,
                                lcid,
                                flags,
                                &dispparams,
                                retvar_vt == VT_VOID ? NULL : &dispargP[0],
                                &einfo,
                                &badarg_index);
    idispP = NULL;    /* May have gone away - see comment above Invoke call */

    if (SUCCEEDED(hr)) {
        if (retvar_vt != VT_VOID) {
            Tcl_SetObjResult(interp, ObjFromVARIANT(&dispargP[0], 0));
        }

        /*
         * Store any out parameters in supplied variables. Note params
         * are stored at indices 1 to nparams. We store the output values
         * in the corresponding variables
         */
        for (i=3, j=nparams; (i < objc) && j; ++i, --j) {
            if (paramflagsP[j] & PARAMFLAG_FOUT) {
                if (Tcl_ObjSetVar2(interp, objv[i], NULL, ObjFromVARIANT(&dispargP[j], 0), TCL_LEAVE_ERR_MSG) == NULL)
                    goto vamoose;
            }
        }

        status = TCL_OK;
    } else {
        /* Failure, fill in exception and return error */
        Tcl_ResetResult(interp); /* Clear out any left-over from arg checking */
        if (hr == DISP_E_EXCEPTION) {
            Tcl_Obj *errorcode_extra; /* Extra argument for error code */
            Tcl_Obj *errorResultObj;

            if (einfo.pfnDeferredFillIn)
                einfo.pfnDeferredFillIn(&einfo);
            
            /* Create an extra argument for the error code */
            errorcode_extra = Tcl_NewListObj(0, NULL);
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     STRING_LITERAL_OBJ("bstrSource"));
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     ObjFromBSTR(einfo.bstrSource));

            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     STRING_LITERAL_OBJ("bstrDescription"));
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     ObjFromBSTR(einfo.bstrDescription));

            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     STRING_LITERAL_OBJ("bstrHelpFile"));
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     ObjFromBSTR(einfo.bstrHelpFile));
            
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     STRING_LITERAL_OBJ("dwHelpContext"));
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     Tcl_NewLongObj(einfo.dwHelpContext));

            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     STRING_LITERAL_OBJ("scode"));
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     Tcl_NewLongObj(einfo.scode));

            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     STRING_LITERAL_OBJ("wCode"));
            Tcl_ListObjAppendElement(NULL, errorcode_extra,
                                     Tcl_NewLongObj(einfo.wCode));

            Twapi_AppendSystemError2(interp, hr, errorcode_extra);
            if (einfo.bstrDescription) {
                errorResultObj = Tcl_DuplicateObj(Tcl_GetObjResult(interp));
                Tcl_AppendUnicodeToObj(errorResultObj, L" ", 1);
                Tcl_AppendUnicodeToObj(errorResultObj,
                                       einfo.bstrDescription,
                                       SysStringLen(einfo.bstrDescription));
                Tcl_SetObjResult(interp, errorResultObj);
            } else {
                /* No error description. Perhaps the scode field
                 * tells us something more.
                 */
                if (einfo.scode &&
                    (FACILITY_WIN32 == HRESULT_FACILITY(einfo.scode) ||
                     FACILITY_WINDOWS == HRESULT_FACILITY(einfo.scode) ||
                     FACILITY_DISPATCH == HRESULT_FACILITY(einfo.scode) ||
                     FACILITY_RPC == HRESULT_FACILITY(einfo.scode))) {
                    Tcl_Obj *scodeObj = Twapi_MapWindowsErrorToString(einfo.scode);
                    if (scodeObj) {
                        Tcl_IncrRefCount(scodeObj);
                        errorResultObj = Tcl_DuplicateObj(Tcl_GetObjResult(interp));
                        Tcl_AppendUnicodeToObj(errorResultObj, L" ", 1);
                        Tcl_AppendObjToObj(errorResultObj, scodeObj);
                        Tcl_SetObjResult(interp, errorResultObj);
                        Tcl_DecrRefCount(scodeObj);
                    }
                }
            }
            
            SysFreeString(einfo.bstrSource);
            SysFreeString(einfo.bstrDescription);
            SysFreeString(einfo.bstrHelpFile);
        } else {
            Twapi_AppendSystemError(interp, hr);
            if ((hr == DISP_E_PARAMNOTFOUND  || hr == DISP_E_TYPEMISMATCH) &&
                badarg_index != -1) {
                char buf[20];
                /* TBD - are the parameter indices backward (ie. from
                 * the Tcl perspective, numbered from the end. In that
                 * case, we should probably map badarg_index appropriately
                 */
                StringCbPrintfA(buf, sizeof(buf), "%d", badarg_index);
                Tcl_AppendResult(interp, " Offending parameter index ", buf, NULL);
            }
        }
    }

 vamoose:
    if (dispargP) {
        /* Release any allocations for return value */
        if (SUCCEEDED(hr)) {
            /* TBD - do we need to clear this for anything else?
             * In particular, make sure we DON"T for interface pointers
             * (VT_DISPATCH, VT_UNKNOWN)
             */
            if (V_VT(&dispargP[0]) == VT_BSTR)
                VariantClear(&dispargP[0]);     /* Will do a SysFreeString */
        }

        /* Release anything allocated for each param */
        for (i = 1; i < nargalloc; ++i)
            TwapiClearVariantParam(interp, &dispargP[i]);
    }
    if (dispargP || paramflagsP)
        MemLifoPopFrame(&ticP->memlifo);

    return status;
}

/*
 * Always returns TCL_ERROR
 */
int Twapi_AppendCOMError(Tcl_Interp *interp, HRESULT hr, ISupportErrorInfo *sei, REFIID iid)
{
    IErrorInfo *ei = NULL;
    if (sei && iid) {
        if (SUCCEEDED(sei->lpVtbl->InterfaceSupportsErrorInfo(sei, iid))) {
            GetErrorInfo(0, &ei);
        }
    }
    if (ei) {
        BSTR msg;
        ei->lpVtbl->GetDescription(ei, &msg);
        Twapi_AppendSystemError2(interp, hr,
                                 Tcl_NewUnicodeObj(msg,SysStringLen(msg)));

        SysFreeString(msg);
        ei->lpVtbl->Release(ei);
    } else {
        Twapi_AppendSystemError(interp, hr);
    }

    return TCL_ERROR;
}

static int TwapiGetIDsOfNamesHelper(
    TwapiInterpContext *ticP,
    void *ifcP,
    Tcl_Obj *namesObj,
    LCID lcid,
    int ifc_type /* 0 -> IDispatch, 1 -> ITypeInfo */
    )
{
    Tcl_Interp *interp = ticP->interp;
    DISPID *ids = NULL;
    HRESULT hr;
    int     status = TCL_ERROR;
    Tcl_Obj **items;
    int     i, nitems;
    LPWSTR *names = NULL;

    /* Convert the list object into an array of points to strings */
    if (Tcl_ListObjGetElements(interp, namesObj, &nitems, &items) == TCL_ERROR)
        return TCL_ERROR;

    names = MemLifoPushFrame(&ticP->memlifo, nitems*sizeof(*names), NULL);

    for (i = 0; i < nitems; i++)
        names[i] = Tcl_GetUnicode(items[i]);

    /* Allocate an array to hold returned ids */
    ids = MemLifoAlloc(&ticP->memlifo, nitems * sizeof(*ids), NULL);

    /* Map the names to ids */
    switch (ifc_type) {
    case 0:
        hr = ((IDispatch *)ifcP)->lpVtbl->GetIDsOfNames((IDispatch *)ifcP, &IID_NULL, names, nitems, lcid, ids);
        break;
    case 1:
        hr = ((ITypeInfo *)ifcP)->lpVtbl->GetIDsOfNames((ITypeInfo *)ifcP, names, nitems, ids);
        break;
    default:
        TwapiReturnTwapiError(interp, "TwapiGetIDsOfNamesHelper: unknown ifc_type", TWAPI_BUG);
        goto vamoose;
    }

    if (SUCCEEDED(hr)) {
        Tcl_Obj *resultObj;
        int      i;

        resultObj = Tcl_NewListObj(0, NULL);
        for (i = 0; i < nitems; ++i) {
            Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewUnicodeObj(names[i], -1));
            Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewIntObj(ids[i]));
        }
        Tcl_SetObjResult(interp, resultObj);
        status = TCL_OK;
    }
    else {
        Twapi_AppendSystemError(interp, hr);
        status = TCL_ERROR;
    }
vamoose:
    MemLifoPopFrame(&ticP->memlifo);

    return status;
}


int Twapi_ITypeInfo_GetTypeAttr(Tcl_Interp *interp, ITypeInfo *tiP)
{
    HRESULT   hr;
    TYPEATTR *taP;
    Tcl_Obj *objv[36];

    hr = tiP->lpVtbl->GetTypeAttr(tiP, &taP);
    if (FAILED(hr))
        return Twapi_AppendSystemError(interp, hr);

    objv[0] = STRING_LITERAL_OBJ("guid");
    objv[1] = ObjFromGUID(&taP->guid);
    objv[2] = STRING_LITERAL_OBJ("lcid");
    objv[3] = Tcl_NewLongObj(taP->lcid);
    objv[4] = STRING_LITERAL_OBJ("dwReserved");
    objv[5] = Tcl_NewLongObj(taP->dwReserved);
    objv[6] = STRING_LITERAL_OBJ("memidConstructor");
    objv[7] = Tcl_NewLongObj(taP->memidConstructor);
    objv[8] = STRING_LITERAL_OBJ("memidDestructor");
    objv[9] = Tcl_NewLongObj(taP->memidDestructor);
    objv[10] = STRING_LITERAL_OBJ("lpstrSchema");
    objv[11] = Tcl_NewUnicodeObj(taP->lpstrSchema ? taP->lpstrSchema : L"", -1);
    objv[12] = STRING_LITERAL_OBJ("cbSizeInstance");
    objv[13] = Tcl_NewLongObj(taP->cbSizeInstance);
    objv[14] = STRING_LITERAL_OBJ("typekind");
    objv[15] = Tcl_NewLongObj(taP->typekind);
    objv[16] = STRING_LITERAL_OBJ("cFuncs");
    objv[17] = Tcl_NewLongObj(taP->cFuncs);
    objv[18] = STRING_LITERAL_OBJ("cVars");
    objv[19] = Tcl_NewLongObj(taP->cVars);
    objv[20] = STRING_LITERAL_OBJ("cImplTypes");
    objv[21] = Tcl_NewLongObj(taP->cImplTypes);
    objv[22] = STRING_LITERAL_OBJ("cbSizeVft");
    objv[23] = Tcl_NewLongObj(taP->cbSizeVft);
    objv[24] = STRING_LITERAL_OBJ("cbAlignment");
    objv[25] = Tcl_NewLongObj(taP->cbAlignment);
    objv[26] = STRING_LITERAL_OBJ("wTypeFlags");
    objv[27] = Tcl_NewLongObj(taP->wTypeFlags);
    objv[28] = STRING_LITERAL_OBJ("wMajorVerNum");
    objv[29] = Tcl_NewLongObj(taP->wMajorVerNum);
    objv[30] = STRING_LITERAL_OBJ("wMinorVerNum");
    objv[31] = Tcl_NewLongObj(taP->wMinorVerNum);
    objv[32] = STRING_LITERAL_OBJ("tdescAlias");
    objv[33] = NULL;
    if (taP->typekind == TKIND_ALIAS) {
        objv[33] = ObjFromTYPEDESC(interp, &taP->tdescAlias, tiP);
    }
    if (objv[33] == NULL)
        objv[33] = Tcl_NewListObj(0, NULL);
    objv[34] = STRING_LITERAL_OBJ("idldescType");
    objv[35] = Tcl_NewIntObj(taP->idldescType.wIDLFlags);

    tiP->lpVtbl->ReleaseTypeAttr(tiP, taP);
    Tcl_SetObjResult(interp, Tcl_NewListObj(36, objv));
    return TCL_OK;
}


int Twapi_ITypeInfo_GetNames(
    Tcl_Interp *interp,
    ITypeInfo *tiP,
    MEMBERID memid)
{
    HRESULT hr;
    BSTR    names[64];
    int     name_count;

    ZeroMemory(names, sizeof(names));

    hr = tiP->lpVtbl->GetNames(tiP, memid, names, sizeof(names)/sizeof(names[0]), &name_count);

    if (SUCCEEDED(hr)) {
        Tcl_Obj *resultObj;
        int      i;

        resultObj = Tcl_NewListObj(0, NULL);
        for (i = 0; i < name_count; ++i) {
            Tcl_ListObjAppendElement(
                interp,
                resultObj,
                Tcl_NewUnicodeObj(names[i], SysStringLen(names[i])));
            SysFreeString(names[i]);
            names[i] = NULL;
        }
        Tcl_SetObjResult(interp, resultObj);
        return TCL_OK;
    }
    else {
        return Twapi_AppendSystemError(interp, hr);
    }
}

int Twapi_ITypeLib_GetLibAttr(Tcl_Interp *interp, ITypeLib *tlP)
{
    HRESULT   hr;
    TLIBATTR *attrP;

    hr = tlP->lpVtbl->GetLibAttr(tlP, &attrP);
    if (SUCCEEDED(hr)) {
        Tcl_Obj *objv[12];
        objv[0] = STRING_LITERAL_OBJ("guid");
        objv[1] = ObjFromGUID(&attrP->guid);
        objv[2] = STRING_LITERAL_OBJ("lcid");
        objv[3] = Tcl_NewLongObj(attrP->lcid);
        objv[4] = STRING_LITERAL_OBJ("syskind");
        objv[5] = Tcl_NewLongObj(attrP->syskind);
        objv[6] = STRING_LITERAL_OBJ("wMajorVerNum");
        objv[7] = Tcl_NewLongObj(attrP->wMajorVerNum);
        objv[8] = STRING_LITERAL_OBJ("wMinorVerNum");
        objv[9] = Tcl_NewLongObj(attrP->wMinorVerNum);
        objv[10] = STRING_LITERAL_OBJ("wLibFlags");
        objv[11] = Tcl_NewLongObj(attrP->wLibFlags);

        tlP->lpVtbl->ReleaseTLibAttr(tlP, attrP);

        Tcl_SetObjResult(interp, Tcl_NewListObj(12, objv));
        return TCL_OK;
    }
    else {
        return Twapi_AppendSystemError(interp, hr);
    }
}


/*
 * Converts a parameter definition in Tcl format into the corresponding
 * VARIANT to be passed to IDispatch::Invoke.
 * IMPORTANT: The created VARIANT must be cleared by calling
 * TwapiClearVariantParam
 * varP is the variant to construct, refvarP is the
 * variant to use if a level of indirection is needed.
 * Both must have been VariantInit'ed.
 * The function also fills in varP->pflags based on the parameter
 * flags field.
 * valueObj is either the name of the variable containing the value
 * to be passed (if the parameter is out or inout) or the actual
 * value itself.
 */
int TwapiMakeVariantParam(
    Tcl_Interp *interp,
    Tcl_Obj *paramDescriptorP,  /* May be NULL if param type is unknown */
    VARIANT *varP,
    VARIANT *refvarP,
    USHORT  *paramflagsP,
    Tcl_Obj *valueObj
    )
{
    Tcl_Obj   **paramv;
    int         paramc;
    VARTYPE     vt;
    HRESULT     hr;
    int         len;
    WCHAR      *wcharP;
    Tcl_Obj   **typev;
    int         typec;
    Tcl_Obj   **reftypev;
    int         reftypec;
    VARIANT    *targetP;         /* Where the actual value is stored */
    int         itemp;
    Tcl_Obj    *paramdefaultObj = NULL;
    Tcl_Obj **paramfields;
    int       paramfieldsc;
    int       status = TCL_ERROR;
    /*
     * paramDescriptorP is a list where the first element is the param type,
     * second element is one/two element list {flags, optional default value}
     * If this info is missing, we assume a BSTR IN parameter
     *
     */
    paramc = 0;
    paramv = NULL;
    typec = 0;
    if (paramDescriptorP) {
        if (Tcl_ListObjGetElements(interp, paramDescriptorP, &paramc, &paramv) != TCL_OK)
            return TCL_ERROR;

        if (paramc >= 1) {
            /* The type information is a list of one or two elements */
            if (Tcl_ListObjGetElements(interp, paramv[0], &typec, &typev) != TCL_OK)
                return TCL_ERROR;

            if (typec == 0 ||
                typec > 2 ||
                ObjToVT(interp, typev[0], &vt) != TCL_OK) {
                goto invalid_type;
            }
        }
    }

    /* Store the parameter flags and default value */
    *paramflagsP = PARAMFLAG_FIN; // In case no paramflags

    if (paramc == 0) {
        /* We did not have parameter info. Assume BSTR */
        vt = VT_BSTR;
    }
    else if (paramc > 1) {
        if (Tcl_ListObjGetElements(interp, paramv[1], &paramfieldsc, &paramfields) != TCL_OK)
            goto vamoose;

        /* First fields is the flags */
        if (paramfieldsc > 0) {
            if (Tcl_GetIntFromObj(NULL, paramfields[0], &itemp) == TCL_OK) {
                *paramflagsP = (USHORT) (itemp ? itemp : PARAMFLAG_FIN);
            } else {
                /* Not an int, see if it is a token */
                char *s = Tcl_GetStringFromObj(paramfields[0], &len);
                if (len == 0 || STREQ(s, "in"))
                    *paramflagsP = PARAMFLAG_FIN;
                else if (STREQ(s, "out"))
                    *paramflagsP = PARAMFLAG_FOUT;
                else if (STREQ(s, "inout"))
                    *paramflagsP = PARAMFLAG_FOUT | PARAMFLAG_FIN;
                else {
                    Tcl_SetResult(interp, "Unknown parameter modifiers", TCL_STATIC);
                    goto vamoose;
                }
            }

        }

        /* Second field, if present is the default. Set valueObj to this
         * in case it is NULL
         */
        if (paramfieldsc > 1) {
            /* Default is {vt value} list.
               We are only interested in the second object */
            if (Tcl_ListObjIndex(interp, paramfields[1], 1, &paramdefaultObj) != TCL_OK)
                goto vamoose;

            /* Note paramdefaultObj may be NULL even if return was TCL_OK */
            if (paramdefaultObj) {
                /* We hold on to this for a while, so make sure it does not
                   go away due to shimmering while the list holding it is
                   modified or changes type */
                Tcl_IncrRefCount(paramdefaultObj);
            }
        }
        /* paramdefaultObj points to actual value (or NULL) at this point */
    }



    /* Note vt is what is allowed in a typedesc, not what is allowed in a
     * VARIANT
     */
    if (vt == VT_SAFEARRAY)
        goto invalid_type;      /* Not yet supported */
    else if (vt == VT_PTR) {
        /* Next element of type list is referenced type */
        if (typec != 2)
            goto invalid_type;

        /* Get the referenced type information. Note that automation does
         * not allow more than one level of indirection so we don't need
         * keep recursing
         */
        if (Tcl_ListObjGetElements(interp, typev[1], &reftypec, &reftypev) != TCL_OK)
            goto vamoose;
        /*
         * Only support base types so there should be exactly one element
         * describing the referenced type.
         */
        /* TBD - add support for referenced arrays */
        if (reftypec != 1 ||
            ObjToVT(interp, reftypev[0], &vt) != TCL_OK) {
            goto invalid_type;
        }
        vt |= VT_BYREF;
        targetP = refvarP;
    } else if (vt == VT_VARIANT) {
#if 0
        //Based on Google groups discussions, we will treat
        //VT_VARIANT the same as VT_VARIANT|VT_BYREF. See for example
        // http://groups.google.co.in/group/borland.public.cppbuilder.activex/browse_thread/thread/ca1c49b278fe7f57/99729a102eedf216?lnk=st&q=VT_VARIANT+parameter&rnum=3#99729a102eedf216
        vt |= VT_BYREF;
        targetP = refvarP;
#else
        /* Types of VT_VARIANT (no BYREF) are passed as BSTR's
         * unless they "look' integer or floating point. The above Google'ed
         * method does not work with the Shell's InvokeVerb method
         */
        if (vt == VT_VARIANT) {
            long ldummy;
            double ddummy;
            if (valueObj &&
                Tcl_GetLongFromObj(NULL, valueObj, &ldummy) == TCL_OK) {
                vt = VT_I4;
            } else if (valueObj &&
                       Tcl_GetDoubleFromObj(NULL, valueObj, &ddummy) == TCL_OK) {
                vt = VT_R8;
            } else
                vt = VT_BSTR;
        }
        targetP = varP;
#endif
    } else {
        /* Either it is a supported base type or we will catch it as
         * unsupported in the switch below
         */
        /* Base types will have typec == 1. It might also be 0 if there
         * was no parameter info provided at all
         */
        if (typec > 1)
            goto invalid_type;

        targetP = varP;
    }

    /*
     * Parameters may be in, out, or inout

     * For in and inout parameters, we need to store the actual value
     * in *targetP. Note that in the case of a pointer type, targetP will
     * be the referenced variant. valueObj holds the value to be stored
     * in the "in" case. In the "inout" case, it holds the name of
     * the variable containing the value to be stored.
     *
     * For 'out" parameters, there is no value to store but we need
     * to construct the variant structure.
     */
    if (! (*paramflagsP & PARAMFLAG_FIN)) {
        /* PARAMFLAG_OUT only. Just init to zero and type convert below */
        V_VT(targetP) = VT_I4;
        V_I4(targetP) = 0;
    } else {
        /* IN or INOUT */
        if (*paramflagsP & PARAMFLAG_FOUT) {
            /* IN/OUT. valueObj is the name of a variable */
            if (valueObj)
                valueObj = Tcl_ObjGetVar2(interp, valueObj, NULL, 0);
        }
        /*
         * valueObj points to actual value. If this is NULL, then we mark
         * the parameter as missing if it is optional. If it is not optional
         * we use the default value
         */
        if (valueObj == NULL) {
            if (*paramflagsP & PARAMFLAG_FOPT) {
                targetP = varP; /* Reset to point to primary VARIANT */
                vt = VT_ERROR;  // Indicates optional param below
            } else {
                /* Not optional - Check if there is a default supplied */
                if (paramdefaultObj == NULL) {
                    /* TBD - should we assume optional argument anyway ? */
                    Tcl_SetResult(interp, "Missing value and no default for IDispatch invoke parameter", TCL_STATIC);
                    goto vamoose;
                }
                /* Default value to be used if value is not specified */
                valueObj = paramdefaultObj;
            }
        }

        /*
         * At this point, targetP points to the variant where the actual
         * value will be stored. This will be refvarP if VT_BYREF is set
         * and varP otherwise.
         *
         */
        switch (vt & ~VT_BYREF) {
        case VT_I2:
        case VT_I4:
        case VT_I1:
        case VT_UI2:
        case VT_UI4:
        case VT_UI1:
        case VT_INT:
        case VT_UINT:
        case VT_HRESULT:
            if (Tcl_GetLongFromObj(interp, valueObj, &targetP->lVal) != TCL_OK)
                goto vamoose;
            targetP->vt = VT_I4;
            break;

        case VT_R4:
        case VT_R8:
            if (Tcl_GetDoubleFromObj(interp, valueObj, &targetP->dblVal) != TCL_OK)
                goto vamoose;
            targetP->vt = VT_R8;
            break;

        case VT_CY:
            if (ObjToCY(interp, valueObj, & V_CY(targetP)) != TCL_OK)
                goto vamoose;
            targetP->vt = VT_CY;
            break;

        case VT_DATE:
            if (Tcl_GetDoubleFromObj(interp, valueObj, & V_DATE(targetP)) != TCL_OK)
                goto vamoose;
            targetP->vt = VT_DATE;
            break;

        case VT_VARIANT:
            /* Only valid if VT_BYREF was set */
            if (! (vt & VT_BYREF)) {
                /*
                 * Should not really happen since we would have set
                 * to VT_BYREF|VT_VARIANT above
                 */
                TwapiInvalidVariantTypeMessage(interp, vt);
                goto vamoose;
            } else {
                /* Value is VARIANT so we don't really know the type.
                 * Just make a best guess. Note we pass NULL for interp
                 * when to GetLong and GetDouble as we don't want an
                 * error message left in the interp.
                 */
                if (Tcl_GetLongFromObj(NULL, valueObj, &targetP->lVal) == TCL_OK) {
                    targetP->vt = VT_I4;
                } else if (Tcl_GetDoubleFromObj(NULL, valueObj, &targetP->dblVal) == TCL_OK) {
                    targetP->vt = VT_R8;
                } else {
                    /* Cannot guess type, just pass as a BSTR */
                    wcharP = Tcl_GetUnicodeFromObj(valueObj,&len);
                    targetP->bstrVal = SysAllocStringLen(wcharP, len);
                    if (targetP->bstrVal == NULL) {
                        Tcl_SetResult(interp, "Insufficient memory", TCL_STATIC);
                        goto vamoose;
                    }
                    targetP->vt = VT_BSTR;
                }
            }
            break;

        case VT_BSTR:
            wcharP = Tcl_GetUnicodeFromObj(valueObj,&len);
            targetP->bstrVal = SysAllocStringLen(wcharP, len);
            if (targetP->bstrVal == NULL) {
                Tcl_SetResult(interp, "Insufficient memory", TCL_STATIC);
                goto vamoose;
            }
            targetP->vt = VT_BSTR;
            break;

        case VT_DISPATCH:
            if (ObjToIDispatch(interp, valueObj, (void **)&targetP->pdispVal)
                != TCL_OK)
                goto vamoose;
            targetP->vt = VT_DISPATCH;
            break;

        case VT_VOID: /* FALLTHRU */
        case VT_ERROR:
            /* Treat as optional argument */
            targetP->vt = VT_ERROR;
            targetP->scode = DISP_E_PARAMNOTFOUND;
            break;

        case VT_BOOL:
            if (Tcl_GetBooleanFromObj(interp, valueObj, &targetP->intVal) != TCL_OK)
                goto vamoose;
            targetP->boolVal = targetP->intVal ? VARIANT_TRUE : VARIANT_FALSE;
            targetP->vt = VT_BOOL;
            break;

        case VT_UNKNOWN:
            if (ObjToIUnknown(interp, valueObj, (void **) &targetP->punkVal)
                != TCL_OK)
                goto vamoose;
            targetP->vt = VT_UNKNOWN;
            break;

        case VT_DECIMAL:
            if (ObjToDECIMAL(interp, valueObj, & V_DECIMAL(targetP)) != TCL_OK)
                goto vamoose;
            targetP->vt = VT_DECIMAL;
            break;

        case VT_I8:
        case VT_UI8:
            if (Tcl_GetWideIntFromObj(interp, valueObj, &targetP->llVal) != TCL_OK)
                goto vamoose;
            targetP->vt = VT_I8;
            break;

        default:
            TwapiInvalidVariantTypeMessage(interp, vt);
            goto vamoose;
        }

        /* Coerce type if necessary of the variant that contains the
         * actual value since the above code only stores the closest
         * related type. An exception is for a byref VT_VARIANT
         */
        if ((vt & ~VT_BYREF)  != targetP->vt &&
            vt != (VT_VARIANT | VT_BYREF)) {
            hr = VariantChangeType(targetP, targetP, 0, (VARTYPE)(vt & ~VT_BYREF));
            if (FAILED(hr)) {
                Twapi_AppendSystemError(interp, hr);
                goto vamoose;
            }
        }

    } /* End Handling of IN and INOUT params */


    /* If the parameter is byref, the above code would have set the referenced
     * variant (containing value). we need to store the appropriate pointer
     * and vt field in the first variant which would not be init'ed yet
     */
    if (vt & VT_BYREF) {
        if (vt == (VT_BYREF|VT_VARIANT)) {
            /* Variant refs are special cased since the target
             * vt is a base type and the primary vt should be VT_BYREF|VARIANT
             */
            V_VT(varP) = vt;
            V_VARIANTREF(varP) = targetP;
        } else {
            V_VT(varP) = V_VT(targetP) | VT_BYREF;
            switch (V_VT(targetP)) {
            case VT_I2:
                V_I2REF(varP) = &V_I2(targetP);
                break;
            case VT_I4:
                V_I4REF(varP) = &V_I4(targetP);
                break;
            case VT_I1:
                V_I1REF(varP) = &V_I1(targetP);
                break;
            case VT_UI2:
                V_UI2REF(varP) = &V_UI2(targetP);
                break;
            case VT_UI4:
                V_UI4REF(varP) = &V_UI4(targetP);
                break;
            case VT_UI1:
                V_UI1REF(varP) = &V_UI1(targetP);
                break;
            case VT_INT:
                V_INTREF(varP) = &V_INT(targetP);
                break;
            case VT_UINT:
                V_UINTREF(varP) = &V_UINT(targetP);
                break;
            case VT_R4:
                V_R4REF(varP) = &V_R4(targetP);
                break;
            case VT_R8:
                V_R8REF(varP) = &V_R8(targetP);
                break;
            case VT_CY:
                V_CYREF(varP) = &V_CY(targetP);
                break;
            case VT_DATE:
                V_DATEREF(varP) = &V_DATE(targetP);
                break;
            case VT_BSTR:
                V_BSTRREF(varP) = &V_BSTR(targetP);
                break;
            case VT_DISPATCH:
                V_DISPATCHREF(varP) = &V_DISPATCH(targetP);
                break;
            case VT_ERROR:
                V_ERRORREF(varP) = &V_ERROR(targetP);
                break;
            case VT_BOOL:
                V_BOOLREF(varP) = &V_BOOL(targetP);
                break;
            case VT_UNKNOWN:
                V_UNKNOWNREF(varP) = &V_UNKNOWN(targetP);
                break;
            case VT_DECIMAL:
                V_DECIMALREF(varP) = &V_DECIMAL(targetP);
                break;
            case VT_I8:
                V_I8REF(varP) = &V_I8(targetP);
                break;
            case VT_UI8:
                V_UI8REF(varP) = &V_UI8(targetP);
                break;
            default:
                Tcl_SetResult(interp, "Internal error while constructing referenced VARIANT parameter", TCL_STATIC);
                goto vamoose;
            }
        }
    }

    status = TCL_OK;

vamoose:
    if (paramdefaultObj)
        Tcl_DecrRefCount(paramdefaultObj);

    return status;

    invalid_type:
        Tcl_SetResult(interp, "Unsupported or invalid type information format in parameter", TCL_STATIC);
        goto vamoose;
}

/*
 * Clears resources for a VARIANT previously created
 * through TwapiMakeVariantParam
 */
void TwapiClearVariantParam(Tcl_Interp *interp, VARIANT *varP)
{
    // TBD - need to revisit this. Currently only VT_BSTR's need deallocs
    /* Note for VT_BYREF|VT_BSTR case, the BSTR is freed when the
     * referenced variant is freed
     */
    if (varP && V_VT(varP) == VT_BSTR)
        VariantClear(varP);     /* Will do a SysFreeString */
}

int Twapi_ITypeComp_Bind(Tcl_Interp *interp, ITypeComp *tcP, LPWSTR nameP, long hashval, unsigned short flags)
{
    ITypeInfo *tiP;
    DESCKIND   desckind;
    BINDPTR    bind;
    HRESULT    hr;
    Tcl_Obj   *objv[3];
    Tcl_Obj   *resultObj;
    int        status;

    hr = tcP->lpVtbl->Bind(tcP,
                           nameP,
                           hashval,
                           flags,
                           &tiP,
                           &desckind,
                           &bind);

    if (hr != S_OK)
        return Twapi_AppendSystemError(interp, hr);

    status = TCL_ERROR;
    switch (desckind) {
    case DESCKIND_NONE:
        resultObj = Tcl_NewListObj(0, NULL);
        status = TCL_OK;
        break;

    case DESCKIND_FUNCDESC:
        objv[0] = STRING_LITERAL_OBJ("funcdesc");
        objv[1] = ObjFromFUNCDESC(interp, bind.lpfuncdesc, tiP);
        objv[2] = ObjFromOpaque(tiP, "ITypeInfo");
        tiP->lpVtbl->ReleaseFuncDesc(tiP, bind.lpfuncdesc);
        resultObj = Tcl_NewListObj(3, objv);
        status = TCL_OK;
        break;

    case DESCKIND_VARDESC:
        objv[0] = STRING_LITERAL_OBJ("vardesc");
        objv[1] = ObjFromVARDESC(interp, bind.lpvardesc, tiP);
        objv[2] = ObjFromOpaque(tiP, "ITypeInfo");
        tiP->lpVtbl->ReleaseVarDesc(tiP, bind.lpvardesc);
        resultObj = Tcl_NewListObj(3, objv);
        status = TCL_OK;
        break;

    case DESCKIND_TYPECOMP:
        objv[0] = STRING_LITERAL_OBJ("typecomp");
        objv[1] = ObjFromOpaque(bind.lptcomp, "ITypeComp");
        resultObj = Tcl_NewListObj(2, objv);
        status = TCL_OK;
        break;

    case DESCKIND_IMPLICITAPPOBJ: /* FALLTHRU */
    default:
        resultObj = STRING_LITERAL_OBJ("Unsupported ITypeComp desckind value");
        break;
    }

    Tcl_SetObjResult(interp, resultObj);
    return status;
}


int Twapi_IRecordInfo_GetFieldNames(Tcl_Interp *interp, IRecordInfo *riP)
{
    ULONG i, nbstrs;
    BSTR  bstrs[50];
    Tcl_Obj *objv[50];
    HRESULT hr;

    ZeroMemory(bstrs, sizeof(bstrs));
    nbstrs = ARRAYSIZE(bstrs);
    hr = riP->lpVtbl->GetFieldNames(riP, &nbstrs, bstrs);
    if (hr != S_OK)
        return Twapi_AppendSystemError(interp, hr);

    /* TBD - revisit - does GetFieldNames return a count in nbstrs or
       should we check for NULL objv[i] values ? */
    for (i = 0; i < nbstrs; ++i) {
        objv[i] = ObjFromBSTR(bstrs[i]);
        SysFreeString(bstrs[i]);
    }

    Tcl_SetObjResult(interp, Tcl_NewListObj(nbstrs, objv));
    return TCL_OK;
}


int TwapiIEnumNextHelper(TwapiInterpContext *ticP,
                         void *com_interface,
                         unsigned long count,
                         int enum_type,
                         int flags
    )
{
    Tcl_Interp *interp = ticP->interp;
    Tcl_Obj *objv[2];           // {More, List_of_elements}
    union {
        void        *pv;
        CONNECTDATA *cdP;
        VARIANT     *varP;
        IConnectionPoint* *icpP;
    } u;
    union {
        void *com_interface;
        IEnumConnections  *IEnumConnections;
        IEnumVARIANT *IEnumVARIANT;
        IEnumConnectionPoints *IEnumConnectionPoints;
        
    } ifc;
    unsigned long ret_count;
    HRESULT  hr;
    unsigned long i;
    size_t elem_size;

    if (count == 0) {
        objv[0] = Tcl_NewBooleanObj(1);
        objv[1] = Tcl_NewListObj(0, NULL);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
        return TCL_OK;
    }

    ifc.com_interface = com_interface;

    switch (enum_type) {
    case 0:
        elem_size = sizeof(*(u.cdP));
        break;
    case 1:
        elem_size = sizeof(*(u.varP));
        break;
    case 2:
        elem_size = sizeof(*(u.icpP));
        break;
    default:
        Tcl_SetResult(interp, "Unknown enum_type passed to TwapiIEnumNextHelper", TCL_STATIC);
        return TCL_ERROR;
        
    }

    u.pv = MemLifoPushFrame(&ticP->memlifo, (DWORD) (count * elem_size), NULL);

    switch (enum_type) {
    case 0:
        hr = ifc.IEnumConnections->lpVtbl->Next(ifc.IEnumConnections,
                                                count, u.cdP, &ret_count);
        break;
    case 1:
        hr = ifc.IEnumVARIANT->lpVtbl->Next(ifc.IEnumVARIANT,
                                            count, u.varP, &ret_count);
        break;
    case 2:
        hr = ifc.IEnumConnectionPoints->lpVtbl->Next(ifc.IEnumConnectionPoints,
                                                     count, u.icpP, &ret_count);
        break;
    }

    /*
     * hr - S_OK ret_count elements returned, more to come
     *      S_FALSE returned elements returned, no more
     *      else error
     */
    if (hr != S_OK && hr != S_FALSE) {
        MemLifoPopFrame(&ticP->memlifo);
        return Twapi_AppendSystemError(interp, hr);
    }

    objv[0] = Tcl_NewBooleanObj(hr == S_OK); // More to come?
    objv[1] = Tcl_NewListObj(0, NULL);
    for (i = 0; i < ret_count; ++i) {
        switch (enum_type) {
        case 0:
            Tcl_ListObjAppendElement(interp, objv[1],
                                     ObjFromCONNECTDATA(&(u.cdP[i])));
            break;
        case 1:
            Tcl_ListObjAppendElement(interp, objv[1],
                                     ObjFromVARIANT(&(u.varP[i]), flags));
            if (u.varP[i].vt == VT_BSTR)
                VariantClear(&(u.varP[i])); // TBD - any other types need clearing?
            break;
        case 3:
            Tcl_ListObjAppendElement(interp, objv[1],
                                     ObjFromOpaque(u.icpP[i],
                                                   "IConnectionPoint"));
            break;
        }
    }
        
    MemLifoPopFrame(&ticP->memlifo);
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
    return TCL_OK;
}


/* Dispatcher for calling COM object methods */
int Twapi_CallCOMObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    union {
        IUnknown *unknown;
        IDispatch *dispatch;
        IDispatchEx *dispatchex;
        IMoniker *moniker;
        ITypeInfo *typeinfo;
        ITypeLib  *typelib;
        IRecordInfo *recordinfo;
        IEnumVARIANT *enumvariant;
        IConnectionPoint *connectionpoint;
        IEnumConnectionPoints *enumconnectionpoints;
        IConnectionPointContainer *connectionpointcontainer;
        IEnumConnections *enumconnections;
        IProvideClassInfo *provideclassinfo;
        IProvideClassInfo2 *provideclassinfo2;
        ITypeComp *typecomp;
        ITaskScheduler *taskscheduler;
        IEnumWorkItems *enumworkitems;
        IScheduledWorkItem *scheduledworkitem;
        ITask *task;
        ITaskTrigger *tasktrigger;
        IPersistFile *persistfile;
    } ifc;
    int func;
    HRESULT hr;
    TwapiResult result;
    DWORD dw1,dw2,dw3;
    HANDLE h;
    BSTR bstr1 = NULL;          /* Initialize for tracking frees! */
    BSTR bstr2 = NULL;
    BSTR bstr3 = NULL;
    int tcl_status;
    void *pv;
    void *pv2;
    GUID guid, guid2;
    TYPEKIND tk;
    LPWSTR s, s2;
    WORD w, w2;
    Tcl_Obj *objs[4];
    SYSTEMTIME systime, systime2;
    FUNCDESC *funcdesc;
    VARDESC  *vardesc;
    TASK_TRIGGER tasktrigger;
    char *cP;

    hr = S_OK;
    result.type = TRT_BADFUNCTIONCODE;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), ARGSKIP,
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    // ARGSKIP makes sure at least one more argument
    if (func < 10000) {
        /* Interface based calls. func codes are all below 10000 */
        if (ObjToLPVOID(interp, objv[2], &pv) != TCL_OK)
            return TCL_ERROR;
        if (pv == NULL) {
            Tcl_SetResult(interp, "NULL interface pointer.", TCL_STATIC);
            return TCL_ERROR;
        }
    }

    /* We want stronger type checking so we have to convert the interface
       pointer on a per interface type basis even though it is repetitive. */

    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    /* function codes are split into ranges assigned to interfaces */
    if (func < 100) {
        /* Specifically for IUnknown we allow any pointer to be converted
         * since all interfaces support the IUnknown interface. This does
         * mean though that even non-interface pointers will be accepted.
         * TBD - strengthen the syntactic check.
         */
        ifc.unknown = (IUnknown *) pv;

        switch (func) {
        case 1: // Release
            result.type = TRT_DWORD;
            result.value.ival = ifc.unknown->lpVtbl->Release(ifc.unknown);
            break;
        case 2:
            result.type = TRT_DWORD;
            result.value.ival = ifc.unknown->lpVtbl->AddRef(ifc.unknown);
            break;
        case 3:
            if (objc < 5)
                return TCL_ERROR;
            hr = CLSIDFromString(Tcl_GetUnicode(objv[3]), &guid);
            if (hr != S_OK)
                break;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = Tcl_GetString(objv[4]);
            hr = ifc.unknown->lpVtbl->QueryInterface(ifc.unknown, &guid,
                                                     &result.value.ifc.p);
            break;
        case 4:
            /* Note this is not a method but a function call */
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = OleRun(ifc.unknown);
            break;
        }
    } else if (func < 200) {
        /* IDispatch */
        /* We accept both IDispatch and IDispatchEx interfaces here */
        if (ObjToIDispatch(interp, objv[2], (void **)&ifc.dispatch) != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 101:
            result.type = TRT_DWORD;
            hr = ifc.dispatch->lpVtbl->GetTypeInfoCount(ifc.dispatch, &result.value.ival);
            break;
        case 102:
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETINT(dw1), GETINT(dw2),
                             ARGEND) != TCL_OK)
                goto ret_error;

            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeInfo";
            hr = ifc.dispatch->lpVtbl->GetTypeInfo(ifc.dispatch, dw1, dw2, (ITypeInfo **)&result.value.ifc.p);
            break;
        case 103: // GetIDsOfNames
            if (objc < 5)
                return TCL_ERROR;
            CHECK_INTEGER_OBJ(interp, dw1, objv[4]);
            return TwapiGetIDsOfNamesHelper(
                ticP, ifc.dispatch, objv[3],
                dw1,              /* LCID */
                0);             /* 0->IDispatch interface */
        }
    } else if (func < 300) {
        /* IDispatchEx */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.dispatchex, "IDispatchEx")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 201: // GetDispID
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETVAR(bstr1, ObjToBSTR), GETINT(dw1),
                             ARGEND) != TCL_OK)
                goto ret_error;
            result.type = TRT_DWORD;
            hr = ifc.dispatchex->lpVtbl->GetDispID(ifc.dispatchex, bstr1,
                                                   dw1, &result.value.ival);
            break;
        case 202: // GetMemberName
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            VariantInit(&result.value.var);
            hr = ifc.dispatchex->lpVtbl->GetMemberName(
                ifc.dispatchex, dw1, &result.value.var.bstrVal);
            if (hr == S_OK) {
                result.value.var.vt = VT_BSTR;
                result.type = TRT_VARIANT;
            }
            break;
        case 203: // GetMemberProperties
        case 204: // GetNextDispID
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETINT(dw1), GETINT(dw2),
                             ARGEND) != TCL_OK)
                goto ret_error;
            result.type = TRT_DWORD;
            if (func == 203)
                hr = ifc.dispatchex->lpVtbl->GetMemberProperties(
                    ifc.dispatchex, dw1, dw2, &result.value.ival);
            else
                hr = ifc.dispatchex->lpVtbl->GetNextDispID(
                    ifc.dispatchex, dw1, dw2, &result.value.ival);
            break;
        case 205: // GetNameSpaceParent
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IUnknown";
            hr = ifc.dispatchex->lpVtbl->GetNameSpaceParent(ifc.dispatchex,
                                                            (IUnknown **)&result.value.ifc.p);
            break;
        case 206: // DeleteMemberByName
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETVAR(bstr1, ObjToBSTR), GETINT(dw1),
                             ARGEND) != TCL_OK)
                goto ret_error;
            hr = S_OK;
            result.type = TRT_BOOL;
            result.value.bval = ifc.dispatchex->lpVtbl->DeleteMemberByName(
                ifc.dispatchex, bstr1, dw1) == S_OK ? 1 : 0;
            break;
        case 207: // DeleteMemberByDispID
            if (TwapiGetArgs(interp, objc-3, objv+3, GETINT(dw1), ARGEND)
                != TCL_OK)
                goto ret_error;
            hr = S_OK;
            result.type = TRT_BOOL;
            result.value.bval = ifc.dispatchex->lpVtbl->DeleteMemberByDispID(
                ifc.dispatchex, dw1) == S_OK ? 1 : 0;
            break;

        }
    } else if (func < 400) {
        /* ITypeInfo */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.typeinfo, "ITypeInfo")
            != TCL_OK)
            return TCL_ERROR;

        if (func == 399) {
            if (objc != 4)
                goto badargs;
            return TwapiGetIDsOfNamesHelper(
                ticP, ifc.typeinfo, objv[3],
                0,              /* Unused */
                1);             /* 1->ITypeInfo interface */
        }

        /* Every other method either has no params or one integer param */
        if (objc > 4)
            goto badargs;

        dw1 = 0;
        if (objc == 4)
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);

        switch (func) {
        case 301: //GetRefTypeOfImplType
            result.type = TRT_DWORD;
            hr = ifc.typeinfo->lpVtbl->GetRefTypeOfImplType(
                ifc.typeinfo, dw1, &result.value.ival);
            break;
        case 302: //GetRefTypeInfo
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeInfo";
            hr = ifc.typeinfo->lpVtbl->GetRefTypeInfo(
                ifc.typeinfo, dw1, (ITypeInfo **)&result.value.ifc.p);
            break;
        case 303: //GetTypeComp
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeComp";
            hr = ifc.typeinfo->lpVtbl->GetTypeComp(
                ifc.typeinfo, (ITypeComp **)&result.value.ifc.p);
            break;
        case 304: // GetContainingTypeLib
            hr = ifc.typeinfo->lpVtbl->GetContainingTypeLib(ifc.typeinfo,
                                                            (ITypeLib **) &pv,
                                                            &dw1);
            if (hr == S_OK) {
                result.type = TRT_OBJV;
                objs[0] = ObjFromOpaque(pv, "ITypeLib");
                objs[1] = Tcl_NewLongObj(dw1);
                result.value.objv.nobj = 2;
                result.value.objv.objPP = objs;
            }
            break;
        case 305: // GetDocumentation
            result.type = TRT_OBJV;
            hr = ifc.typeinfo->lpVtbl->GetDocumentation(
                ifc.typeinfo, dw1, &bstr1, &bstr2, &dw2, &bstr3);
            if (hr == S_OK) {
                objs[0] = ObjFromBSTR(bstr1);
                objs[1] = ObjFromBSTR(bstr2);
                objs[2] = Tcl_NewLongObj(dw2);
                objs[3] = ObjFromBSTR(bstr3);
                result.value.objv.nobj = 4;
                result.value.objv.objPP = objs;
            }
            break;
        case 306: // GetImplTypeFlags
            result.type = TRT_DWORD;
            hr = ifc.typeinfo->lpVtbl->GetImplTypeFlags(
                ifc.typeinfo, dw1, &result.value.ival);
            break;
        case 307: // GetRecordInfofromTypeInfo
            // Note this is NOT a method call but a function which
            // happens to take ITypeInfo * as its first parameters
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IRecordInfo";
            hr = GetRecordInfoFromTypeInfo(ifc.typeinfo,
                                           (IRecordInfo **)&result.value.ifc.p);
            break;
        case 308: // GetNames
            return Twapi_ITypeInfo_GetNames(interp, ifc.typeinfo, dw1);
        case 309:
            return Twapi_ITypeInfo_GetTypeAttr(interp, ifc.typeinfo);
        case 310: // GetFuncDesc
            hr = ifc.typeinfo->lpVtbl->GetFuncDesc(ifc.typeinfo, dw1, &funcdesc);
            if (hr != S_OK)
                break;
            result.type = TRT_OBJ;
            result.value.obj = ObjFromFUNCDESC(interp, funcdesc, ifc.typeinfo);
            ifc.typeinfo->lpVtbl->ReleaseFuncDesc(ifc.typeinfo,
                                                  funcdesc);
            if (result.value.obj == NULL)
                return TCL_ERROR;
            break;
        case 311: // GetVarDesc
            hr = ifc.typeinfo->lpVtbl->GetVarDesc(ifc.typeinfo, dw1, &vardesc);
            if (hr != S_OK)
                break;
            result.type = TRT_OBJ;
            result.value.obj = ObjFromVARDESC(interp, vardesc, ifc.typeinfo);
            ifc.typeinfo->lpVtbl->ReleaseVarDesc(ifc.typeinfo, vardesc);
            if (result.value.obj == NULL)
                goto ret_error;
            break;
        }
    } else if (func < 500) {
        /* ITypeLib */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.typelib, "ITypeLib")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 401: // GetDocumentation
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_OBJV;
            hr = ifc.typelib->lpVtbl->GetDocumentation(
                ifc.typelib, dw1, &bstr1, &bstr2, &dw2, &bstr3);
            if (hr == S_OK) {
                objs[0] = ObjFromBSTR(bstr1);
                objs[1] = ObjFromBSTR(bstr2);
                objs[2] = Tcl_NewLongObj(dw2);
                objs[3] = ObjFromBSTR(bstr3);
                result.value.objv.nobj = 4;
                result.value.objv.objPP = objs;
            }
            break;
        case 402: // GetTypeInfoCount
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            result.value.ival =
                ifc.typelib->lpVtbl->GetTypeInfoCount(ifc.typelib);
            break;
        case 403: // GetTypeInfoType
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_DWORD;
            hr = ifc.typelib->lpVtbl->GetTypeInfoType(ifc.typelib, dw1, &tk);
            if (hr == S_OK)
                result.value.ival = tk;
            break;
        case 404: // GetTypeInfo
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeInfo";
            hr = ifc.typelib->lpVtbl->GetTypeInfo(
                ifc.typelib, dw1, (ITypeInfo **)&result.value.ifc.p);
            break;
        case 405: // GetTypeInfoOfGuid
            if (objc != 4)
                goto badargs;
            hr = CLSIDFromString(Tcl_GetUnicode(objv[3]), &guid);
            if (hr == S_OK) {
                result.type = TRT_INTERFACE;
                result.value.ifc.name = "ITypeInfo";
                hr = ifc.typelib->lpVtbl->GetTypeInfoOfGuid(
                    ifc.typelib, &guid,
                    (ITypeInfo **)&result.value.ifc.p);
            }
            break;
        case 406: // GetLibAttr
            return Twapi_ITypeLib_GetLibAttr(interp, ifc.typelib);

        case 407: // RegisterTypeLib
            if (objc != 5)
                goto badargs;
            s = Tcl_GetUnicode(objv[4]);
            NULLIFY_EMPTY(s);
            result.type = TRT_EMPTY;
            hr = RegisterTypeLib(ifc.typelib, Tcl_GetUnicode(objv[3]), s);
            break;
        }
    } else if (func < 600) {
        /* IRecordInfo */
        /* TBD - for record data, we should create dummy type-safe pointers
           instead of passing around voids even though that is what
           the IRecordInfo interface does */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.recordinfo, "IRecordInfo")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 501: // GetField
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETVOIDP(pv), GETWSTR(s),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            VariantInit(&result.value.var);
            result.type = TRT_VARIANT;
            hr = ifc.recordinfo->lpVtbl->GetField(
                ifc.recordinfo, pv, s, &result.value.var);
            break;
        case 502: // GetGuid
            if (objc != 3)
                goto badargs;
            result.type = TRT_GUID;
            hr = ifc.recordinfo->lpVtbl->GetGuid(ifc.recordinfo, &result.value.guid);
            break;
        case 503: // GetName
            if (objc != 3)
                goto badargs;
            VariantInit(&result.value.var);
            hr = ifc.recordinfo->lpVtbl->GetName(ifc.recordinfo, &result.value.var.bstrVal);
            if (hr == S_OK) {
                result.type = VT_VARIANT;
                result.value.var.vt = VT_BSTR;
            }
            break;
        case 504: // GetSize
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.recordinfo->lpVtbl->GetSize(ifc.recordinfo, &result.value.ival);
            break;
        case 505: // GetTypeInfp
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeInfo";
            hr = ifc.recordinfo->lpVtbl->GetTypeInfo(
                ifc.recordinfo, (ITypeInfo **)&result.value.ifc.p);
            break;
        case 506: // IsMatchingType
            if (objc != 4)
                goto badargs;
            if (ObjToOpaque(interp, objv[3], &pv, "IRecordInfo") != TCL_OK)
                goto ret_error;
            result.type = TRT_DWORD;
            result.value.ival = ifc.recordinfo->lpVtbl->IsMatchingType(
                ifc.recordinfo, (IRecordInfo *) pv);
            break;
        case 507: // RecordClear
            if (objc != 4)
                goto badargs;
            if (ObjToLPVOID(interp, objv[3], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordClear(ifc.recordinfo, pv);
            break;
        case 508: // RecordCopy
            if (objc != 5)
                goto badargs;
            if (ObjToLPVOID(interp, objv[3], &pv) != TCL_OK &&
                ObjToLPVOID(interp, objv[4], &pv2) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordCopy(ifc.recordinfo, pv, pv2);
            break;
        case 509: // RecordCreate
            result.type = TRT_LPVOID;
            result.value.pv =
                ifc.recordinfo->lpVtbl->RecordCreate(ifc.recordinfo);
            break;
        case 510: // RecordCreateCopy
            if (objc != 4)
                goto badargs;
            if (ObjToLPVOID(interp, objv[3], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_LPVOID;
            hr = ifc.recordinfo->lpVtbl->RecordCreateCopy(ifc.recordinfo, pv,
                                                          &result.value.pv);
            break;
        case 511: // RecordDestroy
            if (objc != 4)
                goto badargs;
            if (ObjToLPVOID(interp, objv[3], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordDestroy(ifc.recordinfo, pv);
            break;
        case 512: // RecordInit
            if (objc != 4)
                goto badargs;
            if (ObjToLPVOID(interp, objv[3], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordInit(ifc.recordinfo, pv);
            break;
        case 513:
            if (objc != 3)
                goto badargs;
            return Twapi_IRecordInfo_GetFieldNames(interp, ifc.recordinfo);
        }
    } else if (func < 700) {
        /* IMoniker */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.moniker, "IMoniker")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 601: // GetDisplayName
            if (objc != 5)
                goto badargs;
            if (ObjToOpaque(interp, objv[3], &pv, "IBindCtx") != TCL_OK)
                goto ret_error;
            if (ObjToOpaque(interp, objv[4], &pv2, "IMoniker") != TCL_OK)
                goto ret_error;
            result.type = TRT_LPOLESTR;
            hr = ifc.moniker->lpVtbl->GetDisplayName(
                ifc.moniker, (IBindCtx *)pv, (IMoniker *)pv2,
                &result.value.lpolestr);
            break;
        }
    } else if (func < 800) {
        /* IEnumVARIANT */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.enumvariant, "IEnumVARIANT")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 701: // Clone
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumVARIANT";
            hr = ifc.enumvariant->lpVtbl->Clone(
                ifc.enumvariant, (IEnumVARIANT **)&result.value.ifc.p);
            break;
        case 702: // Reset
            if (objc != 3)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumvariant->lpVtbl->Reset(ifc.enumvariant);
            break;
        case 703: // Skip
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_EMPTY;
            hr = ifc.enumvariant->lpVtbl->Skip(ifc.enumvariant, dw1);
            break;
        case 704: // Next
            if (objc < 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            /* Let caller decide if he wants flat untagged variant list
               or a tagged variant value */
            dw2 = 0;            /* By default tagged value */
            if (objc > 4)
                CHECK_INTEGER_OBJ(interp, dw2, objv[4]);
            return TwapiIEnumNextHelper(ticP,ifc.enumvariant,dw1,1,dw2);
        }        
    } else if (func < 900) {
        /* IConnectionPoint */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.connectionpoint,
                        "IConnectionPoint")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 801: // Advise
            if (objc != 4)
                goto badargs;
            if (ObjToIUnknown(interp, objv[3], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_DWORD;
            hr = ifc.connectionpoint->lpVtbl->Advise(
                ifc.connectionpoint, (IUnknown *)pv, &result.value.ival);
            break;
            
        case 802: // EnumConnections
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnections";
            hr = ifc.connectionpoint->lpVtbl->EnumConnections(ifc.connectionpoint, (IEnumConnections **) &result.value.ifc.p);
            break;

        case 803: // GetConnectionInterface
            if (objc != 3)
                goto badargs;
            result.type = TRT_GUID;
            hr = ifc.connectionpoint->lpVtbl->GetConnectionInterface(ifc.connectionpoint, &result.value.guid);
            break;

        case 804: // GetConnectionPointContainer
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IConnectionPointContainer";
            hr = ifc.connectionpoint->lpVtbl->GetConnectionPointContainer(ifc.connectionpoint, (IConnectionPointContainer **)&result.value.ifc.p);
            break;

        case 805: // Unadvise
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_EMPTY;
            hr = ifc.connectionpoint->lpVtbl->Unadvise(ifc.connectionpoint, dw1);
            break;
        }
    } else if (func < 1000) {
        /* IConnectionPointContainer */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.connectionpointcontainer,
                        "IConnectionPointContainer") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 901: // EnumConnectionPoints
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnectionPoints";
            hr = ifc.connectionpointcontainer->lpVtbl->EnumConnectionPoints(
                ifc.connectionpointcontainer,
                (IEnumConnectionPoints **)&result.value.ifc.p);
            break;
        case 902: // FindConnectionPoint
            if (objc != 4)
                goto badargs;
            hr = CLSIDFromString(Tcl_GetUnicode(objv[3]), &guid);
            if (hr == S_OK) {
                result.type = TRT_INTERFACE;
                result.value.ifc.name = "IConnectionPoint";
                hr = ifc.connectionpointcontainer->lpVtbl->FindConnectionPoint(
                    ifc.connectionpointcontainer,
                    &guid,
                    (IConnectionPoint **)&result.value.ifc.p);
            }
            break;
        }
    } else if (func < 1100) {
        /* IEnumConnectionPoints */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.enumconnectionpoints,
                        "IEnumConnectionPoints") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 1001: // Clone
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnectionPoints";
            hr = ifc.enumconnectionpoints->lpVtbl->Clone(
                ifc.enumconnectionpoints,
                (IEnumConnectionPoints **) &result.value.ifc.p);
            break;
        case 1002: // Reset
            if (objc != 3)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumconnectionpoints->lpVtbl->Reset(
                ifc.enumconnectionpoints);
            break;
        case 1003: // Skip
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_EMPTY;
            hr = ifc.enumconnectionpoints->lpVtbl->Skip(
                ifc.enumconnectionpoints,   dw1);
            break;
        case 1004: // Next
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            return TwapiIEnumNextHelper(ticP,ifc.enumconnectionpoints,
                                        dw1, 2, 0);
        }        
    } else if (func < 1200) {
        /* IEnumConnections */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.enumconnections,
                        "IEnumConnections") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 1101: // Clone
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnections";
            hr = ifc.enumconnections->lpVtbl->Clone(
                ifc.enumconnections, (IEnumConnections **)&result.value.ifc.p);
            break;
        case 1102: // Reset
            if (objc != 3)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumconnections->lpVtbl->Reset(ifc.enumconnections);
            break;
        case 1103: // Skip
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_EMPTY;
            hr = ifc.enumconnections->lpVtbl->Skip(ifc.enumconnections, dw1);
            break;
        case 1104: // Next
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            return TwapiIEnumNextHelper(ticP, ifc.enumconnections,dw1,0,0);
        }        
    } else if (func == 1201) {
        /* IProvideClassInfo */
        /* We accept both IProvideClassInfo and IProvideClassInfo2 interfaces */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.provideclassinfo,
                        "IProvideClassInfo") != TCL_OK &&
            ObjToOpaque(interp, objv[2], (void **)&ifc.provideclassinfo,
                        "IProvideClassInfo2") != TCL_OK)
            return TCL_ERROR;

        result.type = TRT_INTERFACE;
        result.value.ifc.name = "ITypeInfo";
        hr = ifc.provideclassinfo->lpVtbl->GetClassInfo(
            ifc.provideclassinfo, (ITypeInfo **)&result.value.ifc.p);
    } else if (func == 1301) {
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.provideclassinfo2,
                        "IProvideClassInfo2")
            != TCL_OK)
            return TCL_ERROR;

        if (objc != 3)
            goto badargs;
        CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
        result.type = TRT_GUID;
        hr = ifc.provideclassinfo2->lpVtbl->GetGUID(ifc.provideclassinfo2,
                                                    dw1, &result.value.guid);
    } else if (func < 1500) {
        /* ITypeComp */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.typecomp, "ITypeComp")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 1401:
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETWSTR(s), GETINT(dw1), GETWORD(w),
                             ARGEND) != TCL_OK)
                goto ret_error;
            return Twapi_ITypeComp_Bind(interp, ifc.typecomp, s, dw1, w);
        }
    } else if (func < 5100) {
        /* ITaskScheduler */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.taskscheduler,
                        "ITaskScheduler") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5001: // Activate
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETWSTR(s), GETVAR(guid, ObjToGUID),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IUnknown";
            hr = ifc.taskscheduler->lpVtbl->Activate(
                ifc.taskscheduler,   s,   &guid,
                (IUnknown **) &result.value.ifc.p);
            break;
        case 5002: // AddWorkItem
            if (objc != 5)
                goto badargs;
            if (ObjToOpaque(interp, objv[4], &pv, "IScheduledWorkItem") != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.taskscheduler->lpVtbl->AddWorkItem(
                ifc.taskscheduler,
                Tcl_GetUnicode(objv[3]),
                (IScheduledWorkItem *) pv );
            break;
        case 5003: // Delete
            if (objc != 4)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.taskscheduler->lpVtbl->Delete(ifc.taskscheduler,
                                                   Tcl_GetUnicode(objv[3]));
            break;
        case 5004: // Enum
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumWorkItems";
            hr = ifc.taskscheduler->lpVtbl->Enum(
                ifc.taskscheduler,
                (IEnumWorkItems **) &result.value.ifc.p);
            break;
        case 5005: // IsOfType
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETWSTR(s), GETVAR(guid, ObjToGUID),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_DWORD;
            result.value.ival = ifc.taskscheduler->lpVtbl->IsOfType(
                ifc.taskscheduler, s, &guid);
            break;
        case 5006: // NewWorkItem
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETWSTR(s),
                             GETGUID(guid), GETGUID(guid2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IUnknown";
            hr = ifc.taskscheduler->lpVtbl->NewWorkItem(
                ifc.taskscheduler, s, &guid, &guid2,
                (IUnknown **) &result.value.ifc.p);
            break;
        case 5007: // SetTargetComputer
            if (objc != 4)
                goto badargs;
            s = ObjToLPWSTR_NULL_IF_EMPTY(objv[3]);
            hr = ifc.taskscheduler->lpVtbl->SetTargetComputer(
                ifc.taskscheduler, s);
            break;
        case 5008: // GetTargetComputer
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.taskscheduler->lpVtbl->GetTargetComputer(
                ifc.taskscheduler, &result.value.lpolestr);
            break;
        }        
    } else if (func < 5200) {
        /* IEnumWorkItems */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.enumworkitems,
                        "IEnumWorkItems") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5101: // Clone
            if (objc != 3)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumWorkItems";
            hr = ifc.enumworkitems->lpVtbl->Clone(
                ifc.enumworkitems, (IEnumWorkItems **)&result.value.ifc.p);
            break;
        case 5102: // Reset
            if (objc != 3)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumworkitems->lpVtbl->Reset(ifc.enumworkitems);
            break;
        case 5103: // Skip
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_EMPTY;
            hr = ifc.enumworkitems->lpVtbl->Skip(ifc.enumworkitems, dw1);
            break;
        case 5104: // Next
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            return Twapi_IEnumWorkItems_Next(interp,ifc.enumworkitems,dw1);
        }
    } else if (func < 5300) {
        /* IScheduledWorkItem */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.scheduledworkitem,
                        "IScheduledWorkItem") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5201: // CreateTrigger
            if (objc != 3)
                goto badargs;
            result.type = TRT_OBJV;
            hr = ifc.scheduledworkitem->lpVtbl->CreateTrigger(
                ifc.scheduledworkitem, &w, (ITaskTrigger **) &pv);
            if (hr != S_OK)
                break;
            objs[0] = Tcl_NewLongObj(w);
            objs[1] = ObjFromOpaque(pv, "ITaskTrigger");
            result.type = TRT_OBJV;
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            break;
        case 5202: // DeleteTrigger
            if (objc != 4)
                goto badargs;
            if (ObjToWord(interp, objv[3], &w) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->DeleteTrigger(
                ifc.scheduledworkitem, w);
            break;
        case 5203: // EditWorkItem
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETHANDLET(h, HWND), GETINT(dw1),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->EditWorkItem(
                ifc.scheduledworkitem, h, dw1);
            break;
        case 5204: // GetAccountInformation
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetAccountInformation(
                ifc.scheduledworkitem, &result.value.lpolestr);
            break;
        case 5205: // GetComment
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetComment(
                ifc.scheduledworkitem, &result.value.lpolestr);
            break;
        case 5206: // GetCreator
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetCreator(
                ifc.scheduledworkitem, &result.value.lpolestr);
            break;
        case 5207: // GetErrorRetryCount
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetErrorRetryCount(
                ifc.scheduledworkitem, &w);
            result.value.ival = w;
            break;
        case 5208: // GetErrorRetryInterval
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetErrorRetryInterval(
                ifc.scheduledworkitem, &w);
            result.value.ival = w;
            break;
        case 5209: // GetExitCode
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetExitCode(
                ifc.scheduledworkitem, &result.value.ival);
            break;
        case 5210: // GetFlags
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetFlags(
                ifc.scheduledworkitem, &result.value.ival);
            break;
        case 5211: // GetIdleWait
            if (objc != 3)
                goto badargs;
            hr = ifc.scheduledworkitem->lpVtbl->GetIdleWait(
                ifc.scheduledworkitem, &w, &w2);
            if (hr != S_OK)
                break;
            objs[0] = Tcl_NewLongObj(w);
            objs[1] = Tcl_NewLongObj(w2);
            result.type = TRT_OBJV;
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            break;
        case 5212: // GetMostRecentRunTime
            if (objc != 3)
                goto badargs;
            result.type = TRT_SYSTEMTIME;
            hr = ifc.scheduledworkitem->lpVtbl->GetMostRecentRunTime(
                ifc.scheduledworkitem, &result.value.systime);
            break;
        case 5213: // GetNextRunTime
            if (objc != 3)
                goto badargs;
            result.type = TRT_SYSTEMTIME;
            hr = ifc.scheduledworkitem->lpVtbl->GetNextRunTime(
                ifc.scheduledworkitem, &result.value.systime);
            break;
        case 5214: // GetStatus
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetStatus(
                ifc.scheduledworkitem, &result.value.ival);
            break;
        case 5215: // GetTrigger
            if (objc != 4)
                goto badargs;
            if (ObjToWord(interp, objv[3], &w) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITaskTrigger";
            hr = ifc.scheduledworkitem->lpVtbl->GetTrigger(
                ifc.scheduledworkitem, w, (ITaskTrigger **)&result.value.ifc.p);
            break;
        case 5216: // GetTriggerCount
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetTriggerCount(
                ifc.scheduledworkitem, &w);
            result.value.ival = w;
            break;
        case 5217: // GetTriggerString
            if (objc != 4)
                goto badargs;
            if (ObjToWord(interp, objv[3], &w) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetTriggerString(
                ifc.scheduledworkitem, w, &result.value.lpolestr);
            break;
        case 5218: // Run
            if (objc != 3)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->Run(ifc.scheduledworkitem);
            break;
        case 5219: // SetAccountInformation
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETWSTR(s),
                             ARGUSEDEFAULT,
                             GETNULLTOKEN(s2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetAccountInformation(
                ifc.scheduledworkitem, s, s2);
            break;
        case 5220: // SetComment
            if (objc != 4)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetComment(
                ifc.scheduledworkitem, Tcl_GetUnicode(objv[3]));
            break;
        case 5221: // SetCreator
            if (objc != 4)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetCreator(
                ifc.scheduledworkitem, Tcl_GetUnicode(objv[3]));
            break;
        case 5222: // SetErrorRetryCount
            if (objc != 4)
                goto badargs;
            if (ObjToWord(interp, objv[3], &w) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetErrorRetryCount(
                ifc.scheduledworkitem, w);
            break;
        case 5223: // SetErrorRetryInterval
            if (objc != 4)
                goto badargs;
            if (ObjToWord(interp, objv[3], &w) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetErrorRetryInterval(
                ifc.scheduledworkitem, w);
            break;
        case 5224: // SetFlags
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetFlags(
                ifc.scheduledworkitem, dw1);
            break;
        case 5225: // SetIdleWait
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETWORD(w), GETWORD(w2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetIdleWait(
                ifc.scheduledworkitem, w, w2);
            break;
        case 5226: // Terminate
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->Terminate(ifc.scheduledworkitem);
            break;
        case 5227: // SetWorkItemData
            if (objc != 4)
                goto badargs;
            pv = Tcl_GetByteArrayFromObj(objv[3], &dw1);
            if (dw1 > MAXWORD) {
                Tcl_SetResult(interp, "Binary data exceeds MAXWORD", TCL_STATIC);
                return TCL_ERROR;
            }
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetWorkItemData(
                ifc.scheduledworkitem, (WORD) dw1, pv);
            break;
        case 5228: // GetWorkItemData
            if (objc != 3)
                goto badargs;
            return Twapi_IScheduledWorkItem_GetWorkItemData(
                interp, ifc.scheduledworkitem);
        case 5229: // GetRunTimes
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETVAR(systime, ObjToSYSTEMTIME),
                             GETVAR(systime2, ObjToSYSTEMTIME),
                             GETWORD(w),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_IScheduledWorkItem_GetRunTimes(
                interp, ifc.scheduledworkitem, &systime, &systime2, w);
        }
    } else if (func < 5400) {
        /* ITask */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.task, "ITask") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5301: // GetApplicationName
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.task->lpVtbl->GetApplicationName(
                ifc.task, &result.value.lpolestr);
            break;
        case 5302: // GetMaxRunTime
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.task->lpVtbl->GetMaxRunTime(ifc.task, &result.value.ival);
            break;
        case 5303: // GetParameters
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.task->lpVtbl->GetParameters(
                ifc.task, &result.value.lpolestr);
            break;
        case 5304: // GetPriority
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.task->lpVtbl->GetPriority(ifc.task, &result.value.ival);
            break;
        case 5305: // GetTaskFlags
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.task->lpVtbl->GetTaskFlags(ifc.task, &result.value.ival);
            break;
        case 5306: // GetWorkingDirectory
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.task->lpVtbl->GetWorkingDirectory(
                ifc.task, &result.value.lpolestr);
            break;
        case 5307: // SetApplicationName
            if (objc != 4)
                goto badargs;
            hr = ifc.task->lpVtbl->SetApplicationName(
                ifc.task, Tcl_GetUnicode(objv[3]));
            break;
        case 5308: // SetParameters
            if (objc != 4)
                goto badargs;
            hr = ifc.task->lpVtbl->SetParameters(
                ifc.task, Tcl_GetUnicode(objv[3]));
            break;
        case 5309: // SetWorkingDirectory
            if (objc != 4)
                goto badargs;
            hr = ifc.task->lpVtbl->SetWorkingDirectory(
                ifc.task, Tcl_GetUnicode(objv[3]));
            break;
        case 5310: // SetMaxRunTime
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            hr = ifc.task->lpVtbl->SetMaxRunTime(ifc.task, dw1);
            break;
        case 5311: // SetPriority
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            hr = ifc.task->lpVtbl->SetPriority(ifc.task, dw1);
            break;
        case 5312: // SetTaskFlags
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            hr = ifc.task->lpVtbl->SetTaskFlags(ifc.task, dw1);
            break;
        }
    } else if (func < 5500) {
        /* ITaskTrigger */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.tasktrigger,
                        "ITaskTrigger") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5401: // GetTrigger
            if (objc != 3)
                goto badargs;
            hr = ifc.tasktrigger->lpVtbl->GetTrigger(ifc.tasktrigger,
                                                     &tasktrigger);
            if (hr != S_OK)
                break;
            result.type = TRT_OBJ;
            result.value.obj = ObjFromTASK_TRIGGER(&tasktrigger);
            break;
        case 5402: // GetTriggerString
            if (objc != 3)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.tasktrigger->lpVtbl->GetTriggerString(
                ifc.tasktrigger, &result.value.lpolestr);
            break;
        case 5403: // SetTrigger
            if (objc != 4)
                goto badargs;
            if (ObjToTASK_TRIGGER(interp, objv[3], &tasktrigger) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.tasktrigger->lpVtbl->SetTrigger(ifc.tasktrigger,
                                                     &tasktrigger);
            break;
        }
    } else if (func < 5600) {
        /* IPersistFile */
        if (ObjToOpaque(interp, objv[2], (void **)&ifc.persistfile,
                        "IPersistFile") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5501: // GetCurFile
            if (objc != 3)
                goto badargs;
            hr = ifc.persistfile->lpVtbl->GetCurFile(
                ifc.persistfile, &result.value.lpolestr);
            if (hr != S_OK && hr != S_FALSE)
                break;
            /* Note S_FALSE also is a success return */
            result.type = TRT_OBJV;
            objs[0] = Tcl_NewLongObj(hr);
            objs[1] = Tcl_NewUnicodeObj(result.value.lpolestr, -1);
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            break;
        case 5502: // IsDirty
            if (objc != 3)
                goto badargs;
            result.type = TRT_DWORD;
            result.value.ival = ifc.persistfile->lpVtbl->IsDirty(ifc.persistfile);
            break;
        case 5503: // Load
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETWSTR(s), GETINT(dw1),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            hr = ifc.persistfile->lpVtbl->Load(ifc.persistfile, s, dw1);
            break;
        case 5504: // Save
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETNULLIFEMPTY(s), GETINT(dw1),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.persistfile->lpVtbl->Save(ifc.persistfile, s, dw1);
            break;
        case 5505: // SaveCompleted
            if (objc != 4)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.persistfile->lpVtbl->SaveCompleted(
                ifc.persistfile, Tcl_GetUnicode(objv[3]));
            break;
        }
    } else {
        /* Commands that are not method calls on interfaces. These
           are here and not in twapi_calls.c because they make use
           of COM typedefs (historical, not sure that's still true) */
        switch (func) {
        case 10001: // CreateFileMoniker
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IMoniker";
            hr = CreateFileMoniker(Tcl_GetUnicode(objv[2]),
                                   (IMoniker **)&result.value.ifc.p);
            break;
        case 10002: // CreateBindCtx
            CHECK_INTEGER_OBJ(interp, dw1, objv[2]);
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IBindCtx";
            hr = CreateBindCtx(dw1, (IBindCtx **)&result.value.ifc.p);
            break;
        case 10003: // GetRecordInfoFromGuids
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETGUID(guid), GETINT(dw1), GETINT(dw2),
                             GETINT(dw3), GETGUID(guid2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IRecordInfo";
            hr = GetRecordInfoFromGuids(&guid, dw1, dw2, dw3, &guid2,
                                        (IRecordInfo **) &result.value.ifc.p);
            break;
        case 10004: // QueryPathOfRegTypeLib
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETGUID(guid), GETWORD(w), GETWORD(w2),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.value.var.bstrVal = NULL;
            hr = QueryPathOfRegTypeLib(&guid, w, w2, dw3, &result.value.var.bstrVal);
            if (hr == S_OK) {
                result.value.var.vt = VT_BSTR;
                result.type = TRT_VARIANT;
            }
            break;
        case 10005: // UnRegisterTypeLib
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETGUID(guid), GETWORD(w), GETWORD(w2),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = UnRegisterTypeLib(&guid, w, w2, dw3, SYS_WIN32);
            break;
        case 10006: // LoadRegTypeLib
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETGUID(guid), GETWORD(w), GETWORD(w2),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeLib";
            hr = LoadRegTypeLib(&guid, w, w2, dw3,
                                (ITypeLib **) &result.value.ifc.p);
            break;
        case 10007: // LoadTypeLibEx
            if (objc != 4)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeLib";
            hr = LoadTypeLibEx(Tcl_GetUnicode(objv[2]), dw1,
                                              (ITypeLib **)&result.value.ifc.p);
            break;
        case 10008: // CoGetObject
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), ARGSKIP, GETGUID(guid), GETASTR(cP),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (Tcl_ListObjLength(interp, objv[3], &dw1) == TCL_ERROR ||
                dw1 != 0) {
                Tcl_SetResult(interp, "Bind options are not supported for CoGetOjbect and must be specified as empty.", TCL_STATIC); // TBD
                goto ret_error;
            }
            result.type = TRT_INTERFACE;
            result.value.ifc.name = cP;
            hr = CoGetObject(s, NULL, &guid, &result.value.ifc.p);
            break;
        case 10009: // GetActiveObject
            if (objc != 3)
                goto badargs;
            if (ObjToGUID(interp, objv[2], &guid) != TCL_OK)
                goto ret_error;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IUnknown";
            hr = GetActiveObject(&guid, NULL, (IUnknown **)&result.value.ifc.p);
            break;
        case 10010: // CLSIDFromString
            if (objc != 3)
                goto badargs;
            if (ObjToGUID(interp, objv[2], &guid) != TCL_OK)
                goto ret_error;
            result.type = TRT_LPOLESTR;
            hr = ProgIDFromCLSID(&guid, &result.value.lpolestr);
            break;
        case 10011:  // CLSIDFromProgID
            if (objc != 3)
                goto badargs;
            result.type = TRT_GUID;
            hr = CLSIDFromProgID(Tcl_GetUnicode(objv[2]), &result.value.guid);
            break;
        case 10012: // CoCreateInstance
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETGUID(guid), ARGSKIP, GETINT(dw1),
                             GETGUID(guid2), GETASTR(cP),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (ObjToIUnknown(interp, objv[3], (void **)&ifc.unknown)
                != TCL_OK)
                goto ret_error;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = cP;
            hr = CoCreateInstance(&guid, ifc.unknown, dw1, &guid2,
                                  &result.value.ifc.p);
            break;
        case 10013:
            if (objc != 3)
                goto badargs;
            result.type = TRT_BOOL;
            result.value.bval = CLSIDFromString(Tcl_GetUnicode(objv[2]), &guid) == S_OK;
            break;
        }
    }

    if (hr != S_OK) {
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = hr;
    }

    tcl_status = TwapiSetResult(interp, &result);

vamoose:
    // Free bstr AFTER setting result as result.value.unicode may point to it */
    SysFreeString(bstr1);        /* OK if bstr is NULL */
    SysFreeString(bstr2);
    SysFreeString(bstr3);
    return tcl_status;

badargs:
    TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

ret_error:
    tcl_status = TCL_ERROR;
    goto vamoose;
}


#if 0
// TBD Only on Win2k3
EXCEPTION_ON_ERROR RegisterTypeLibForUser(
    ITypeLib *ptlib,
    LPWSTR    szFullPath,
    LPWSTR    szHelpDir);
EXCEPTION_ON_ERROR UnRegisterTypeLibForUser(
    UUID *INPUT,
    unsigned short  wVerMajor,
    unsigned short  wVerMinor,
    DWORD  lcid,
    int  syskind);
#endif

