/* 
 * Copyright (c) 2006-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 *
 * Utility functions used by the COM module
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

static int TwapiComInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP);
static int TwapiMakeVariantParam(
    Tcl_Interp *interp,
    Tcl_Obj *paramDescriptorP,
    VARIANT *varP,
    VARIANT *refvarP,
    USHORT  *paramflagsP,
    Tcl_Obj *valueObj
    );

static TwapiModuleDef gModuleDef = {
    MODULENAME,
    TwapiComInitCalls,
    NULL,
    0
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




/*
 *  TBD - does this (related methods) need to be made thread safe?
 */
typedef struct Twapi_EventSink {
    interface IDispatch idispP; /* Must be first field */
    IID iid;                    /* IID for this event sink interface */
    int refc;                   /* Ref count */
    Tcl_Interp *interp;
#define  MAX_EVENTSINK_CMDARGS 16
    Tcl_Obj *cmd;               /* Stores the callback command arg list */

} Twapi_EventSink;



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


Tcl_Obj *ObjFromCONNECTDATA (const CONNECTDATA *cdP)
{
    Tcl_Obj *objv[2];
    objv[0] = ObjFromIUnknown(cdP->pUnk);
    objv[1] = ObjFromLong(cdP->dwCookie);
    return ObjNewList(2, objv);
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
            ObjSetStaticResult(interp, "Internal error: ObjFromTYPEDESC: NULL TYPEDESC pointer");
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
        objv[2] = ObjNewList(0, NULL); /* For dimension info */
        for (i = 0; i < tdP->lpadesc->cDims; ++i) {
            ObjAppendElement(interp, objv[2],
                                     ObjFromInt(tdP->lpadesc->rgbounds[i].lLbound));
            ObjAppendElement(interp, objv[2],
                                     ObjFromInt(tdP->lpadesc->rgbounds[i].cElements));
        }
        objc = 3;
        break;

    case VT_USERDEFINED:
        // Original code used to resolve the name. However, this was not
        // very useful by itself so we now leave it up to the Tcl level
        // to do so if required.
#if 1
        objv[1] = ObjFromDWORD(tdP->hreftype);
#else
        objv[1] = NULL;
        if (tiP->lpVtbl->GetRefTypeInfo(tiP, tdP->hreftype, &utiP) == S_OK) {
            BSTR bstr;
            if (utiP->lpVtbl->GetDocumentation(utiP, MEMBERID_NIL, &bstr, NULL, NULL, NULL) == S_OK) {
                objv[1] = ObjFromUnicodeN(bstr, SysStringLen(bstr));
                SysFreeString(bstr);
            }
            utiP->lpVtbl->Release(utiP);
        }
        if (objv[1] == NULL) {
            /* Could not get name of custom type. */
            objv[1] = ObjFromEmptyString();
        }
#endif
        objc = 2;
        break;

    default:
        objc = 1;
        break;
    }

    objv[0] = ObjFromInt(tdP->vt);
    return ObjNewList(objc, objv);
}

static Tcl_Obj *ObjFromPARAMDESC(Tcl_Interp *interp, PARAMDESC *vdP, ITypeInfo *tiP)
{
    Tcl_Obj *objv[2];

    objv[0] = ObjFromInt(vdP->wParamFlags);
    objv[1] = NULL;
    if ((vdP->wParamFlags & PARAMFLAG_FOPT) &&
        (vdP->wParamFlags & PARAMFLAG_FHASDEFAULT) &&
        vdP->pparamdescex) {
        objv[1] = ObjFromVARIANT(&vdP->pparamdescex->varDefaultValue, 0);
    }

    return ObjNewList(objv[1] ? 2 : 1, objv);
}

static Tcl_Obj *ObjFromVARDESC(Tcl_Interp *interp, VARDESC *vdP, ITypeInfo *tiP)
{
    Tcl_Obj *resultObj = ObjNewList(0, 0);
    Tcl_Obj *obj;

    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, vdP, memid);
    if (vdP->lpstrSchema) {
        Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, resultObj, vdP, lpstrSchema);
    }
    else {
        ObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("lpstrSchema"));
        ObjAppendElement(interp, resultObj, ObjFromEmptyString());
    }

    switch (vdP->varkind) {
    case VAR_PERINSTANCE: /* FALLTHROUGH */
    case VAR_DISPATCH: /* FALLTHROUGH */
    case VAR_STATIC:
        Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, vdP, oInst);
        break;
    case VAR_CONST:
        ObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("lpvarValue"));
        ObjAppendElement(interp, resultObj, ObjFromVARIANT(vdP->lpvarValue, 0));
        break;
    }

    ObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("elemdescVar.tdesc"));
    obj = ObjFromTYPEDESC(interp, &vdP->elemdescVar.tdesc, tiP);
    ObjAppendElement(interp, resultObj,
                             obj ? obj : ObjNewList(0, NULL));

    ObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("elemdescVar.paramdesc"));
    obj = ObjFromPARAMDESC(interp, &vdP->elemdescVar.paramdesc, tiP);
    ObjAppendElement(interp, resultObj,
                             obj ? obj : ObjNewList(0, NULL));

    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, vdP, varkind);
    Twapi_APPEND_WORD_FIELD_TO_LIST(interp, resultObj, vdP, wVarFlags);

    return resultObj;
}


static Tcl_Obj *ObjFromFUNCDESC(Tcl_Interp *interp, FUNCDESC *fdP, ITypeInfo *tiP)
{
    Tcl_Obj *resultObj = ObjNewList(0, 0);
    Tcl_Obj *obj;
    int      i;

    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, memid);
    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, funckind);
    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, invkind);
    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, callconv);
    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, cParams);
    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, cParamsOpt);
    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, oVft);
    Twapi_APPEND_LONG_FIELD_TO_LIST(interp, resultObj, fdP, wFuncFlags);
    ObjAppendElement(interp, resultObj,
                             STRING_LITERAL_OBJ("elemdescFunc.tdesc"));
    obj = ObjFromTYPEDESC(interp, &fdP->elemdescFunc.tdesc, tiP);
    ObjAppendElement(interp, resultObj,
                             obj ? obj : ObjNewList(0, NULL));

    ObjAppendElement(interp, resultObj, STRING_LITERAL_OBJ("elemdescFunc.paramdesc"));
    obj = ObjFromPARAMDESC(interp, &fdP->elemdescFunc.paramdesc, tiP);
    ObjAppendElement(interp, resultObj,
                             obj ? obj : ObjNewList(0, NULL));

    /* List of possible return codes */
    obj = ObjNewList(0, NULL);
    if (fdP->lprgscode) {
        for (i = 0; i < fdP->cScodes; ++i) {
            ObjAppendElement(interp, resultObj,
                                     ObjFromDWORD(fdP->lprgscode[i]));
        }
    }
    ObjAppendElement(interp, resultObj,
                             STRING_LITERAL_OBJ("lprgscode"));
    ObjAppendElement(interp, resultObj, obj);

    /* List of parameter descriptors */
    obj = ObjNewList(0, NULL);
    if (fdP->lprgelemdescParam) {
        for (i = 0; i < fdP->cParams; ++i) {
            Tcl_Obj *paramObj[2];
            paramObj[0] =
                ObjFromTYPEDESC(interp,
                                &(fdP->lprgelemdescParam[i].tdesc), tiP);
            if (paramObj[0] == NULL)
                paramObj[0] = ObjNewList(0, 0);
            paramObj[1] =
                ObjFromPARAMDESC(interp,
                                 &(fdP->lprgelemdescParam[i].paramdesc), tiP);
            if (paramObj[1] == NULL)
                paramObj[1] = ObjNewList(0, 0);
            ObjAppendElement(interp, obj, ObjNewList(2, paramObj));
        }
    }
    ObjAppendElement(interp, resultObj,
                             STRING_LITERAL_OBJ("lprgelemdescParam"));
    ObjAppendElement(interp, resultObj, obj);

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
            ObjDecrRefs(me->cmd);
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
    Tcl_InterpState savedState;

    if (me == NULL)
        return E_POINTER;

    if (ObjGetElements(NULL, me->cmd, &cmdobjc, &cmdprefixv) != TCL_OK) {
        /* Internal error - should not happen. Should we log background error?*/
        return E_FAIL;
    }

    /* Note we  will tack on 3 additional fixed arguments plus dispparms */
    /* TBD - where is this freed ? */
    cmdobjv = TwapiAlloc((cmdobjc+4) * sizeof(*cmdobjv));
    
    for (i = 0; i < cmdobjc; ++i) {
        cmdobjv[i] = cmdprefixv[i];
    }

    cmdobjv[cmdobjc] = ObjFromLong(dispid);
    cmdobjv[cmdobjc+1] = ObjFromLong(lcid);
    cmdobjv[cmdobjc+2] = ObjFromInt(flags);
    cmdobjc += 3;

    /* Add the passed parameters */
    cmdobjv[cmdobjc] = ObjNewList(0, NULL);
    if (dispparamsP) {
        /* Note parameters are in reverse order */
        for (i = dispparamsP->cArgs - 1; i >= 0 ; --i) {
            ObjAppendElement(
                NULL,
                cmdobjv[cmdobjc],
                ObjFromVARIANT(&dispparamsP->rgvarg[i], 0)
                );
        }
    }
    ++cmdobjc;
                 

    for (i = 0; i < cmdobjc; ++i) {
        ObjIncrRefs(cmdobjv[i]); /* Protect while we are using it.
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
       the Tcl_SaveInterpState/RestoreInterpState really necessary ?
       Note tclWinDde also evals in this fashion.
    */
    /* If hr is not TCL_OK, it is a HRESULT error code */
    savedState = Tcl_SaveInterpState(me->interp, TCL_OK);
    Tcl_ResetResult (me->interp);
    hr = Tcl_EvalObjv(me->interp, cmdobjc, cmdobjv, TCL_EVAL_GLOBAL);
    if (hr != TCL_OK) {
        Tcl_BackgroundError(me->interp);
        if (excepP) {
            TwapiZeroMemory(excepP, sizeof(*excepP));
            excepP->scode = hr;
        }
    } else {
        /* TBD - appropriately init retvarP from ObjGetResult keeping
         * in mind that the retvarP by be BYREF as well.
         */
        if (retvarP)
            VariantInit(retvarP);
        hr = S_OK;
    }
    Tcl_RestoreInterpState(me->interp, savedState);

    /* Free the objects we allocated */
    for (i = 0; i < cmdobjc; ++i) {
        ObjDecrRefs(cmdobjv[i]);
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
    
    hr = IIDFromString(ObjToUnicode(objv[1]), &iid);
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
    ObjIncrRefs(objv[2]);
    sinkP->cmd = objv[2];

    ObjSetResult(interp, ObjFromIUnknown(sinkP));


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

    TWAPI_OBJ_LOG_IF(gModuleDef.log_flags, interp, ObjNewList(objc, objv));

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
     *   4 - (optional) list of parameter types to dispatch function. This is a list
     *       of elements of the form {type paramflags ?defaultvalue?}. If this
     *       element is missing (as opposed to empty), it means no parameter
     *       info is available and we will go strictly by the arguments
     *       present.
     *   5 - (optional) list of param names. Not used here. Only present
     *       if param types field is present
     */

    if (objc < 3) {
        Tcl_WrongNumArgs(interp, 1, objv, "IDISPATCH PROTOTYPE ?ARG1 ARG2...?");
        return TCL_ERROR;
    }

    if (ObjToIDispatch(interp, objv[1], &idispP) != TCL_OK)
        return TCL_ERROR;

    if (ObjGetElements(interp, objv[2], &protoc, &protov) != TCL_OK)
        return TCL_ERROR;

    /* Extract prototype information */
    if (TwapiGetArgs(interp, protoc, protov,
                     GETINT(dispid), GETINT(lcid),
                     GETWORD(flags), GETVAR(retvar_vt, ObjToVT),
                     ARGTERM) != TCL_OK) {
        ObjSetStaticResult(interp, "Invalid IDispatch prototype - must contain DISPID LCID FLAGS RETTYPE ?PARAMTYPES?");
        return TCL_ERROR;
    }

    if (protoc >= 5) {
        /* Extract the parameter information */
        if (ObjGetElements(interp, protov[4], &nparams, &params) != TCL_OK)
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
    if(flags & (DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF)) {
        dispparams.cNamedArgs = 1;
        named_dispid  = DISPID_PROPERTYPUT;
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
    TwapiZeroMemory(&einfo, sizeof(einfo));
    badarg_index = (UINT) -1;

    TWAPI_LOG_BLOCK(gModuleDef.log_flags) {
        Tcl_Obj *logObj = ObjEmptyList();
        ObjAppendElement(NULL, logObj, ObjFromString("Invoke parameters"));
        /* Note parameters start at index 1 */
        for (i=1; i <= nparams; ++i) {
            Tcl_Obj *pair[2];
            pair[0] = ObjFromInt(dispargP[i].vt);
            pair[1] = ObjFromVARIANT(&dispargP[i], 0);
            ObjAppendElement(NULL, logObj, ObjNewList(2, pair));
        }
        TWAPI_OBJ_LOG(interp, logObj);
    }

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
        /*
         * Some calls return the "return value" as output of the
         * last parameter which would be marked as OUT|RETVAL.
         * In that case use it as command return instead of
         * using the call return value
         */
        if (retvar_vt == VT_HRESULT &&
            nparams &&
            (paramflagsP[nparams] & (PARAMFLAG_FOUT | PARAMFLAG_FRETVAL)) == (PARAMFLAG_FOUT | PARAMFLAG_FRETVAL)) {
            /* Yep, one of those funky results */
            if (SUCCEEDED(V_I4(&dispargP[0]))) {
                /* Yes return status is success, pick up return value */
                ObjSetResult(interp, ObjFromVARIANT(&dispargP[nparams], 0));
            } else {
                /* Hmm, toplevel HRESULT is success, retval HRESULT is not
                   WTF does that mean ? Treat as error */
                /* Not sure if standard COM error handling should apply  so skip */
                ObjSetStaticResult(interp, "Unexpected COM result: Invoke returned success but retval param is error.");
                goto vamoose;
            }
        } else {
            if (retvar_vt != VT_VOID) {
                ObjSetResult(interp, ObjFromVARIANT(&dispargP[0], 0));
            }
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
            Tcl_Obj *errorcode_extra[12]; /* Extra argument for error code */
            Tcl_Obj *errorResultObj;

            if (einfo.pfnDeferredFillIn)
                einfo.pfnDeferredFillIn(&einfo);
            
            /* Create an extra argument for the error code */
            errorcode_extra[0] = STRING_LITERAL_OBJ("bstrSource");
            errorcode_extra[1] = ObjFromBSTR(einfo.bstrSource);
            errorcode_extra[2] = STRING_LITERAL_OBJ("bstrDescription");
            errorcode_extra[3] = ObjFromBSTR(einfo.bstrDescription);
            errorcode_extra[4] = STRING_LITERAL_OBJ("bstrHelpFile");
            errorcode_extra[5] = ObjFromBSTR(einfo.bstrHelpFile);
            errorcode_extra[6] = STRING_LITERAL_OBJ("dwHelpContext");
            errorcode_extra[7] = ObjFromLong(einfo.dwHelpContext);
            errorcode_extra[8] = STRING_LITERAL_OBJ("scode");
            errorcode_extra[9] = ObjFromLong(einfo.scode);
            errorcode_extra[10] = STRING_LITERAL_OBJ("wCode");
            errorcode_extra[11] = ObjFromLong(einfo.wCode);

            Twapi_AppendSystemErrorEx(interp, hr, ObjNewList(ARRAYSIZE(errorcode_extra), errorcode_extra));

            if (einfo.bstrDescription) {
                errorResultObj = ObjDuplicate(ObjGetResult(interp));
                Tcl_AppendUnicodeToObj(errorResultObj, L" ", 1);
                Tcl_AppendUnicodeToObj(errorResultObj,
                                       einfo.bstrDescription,
                                       SysStringLen(einfo.bstrDescription));
                ObjSetResult(interp, errorResultObj);
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
                        ObjIncrRefs(scodeObj);
                        errorResultObj = ObjDuplicate(ObjGetResult(interp));
                        Tcl_AppendUnicodeToObj(errorResultObj, L" ", 1);
                        Tcl_AppendObjToObj(errorResultObj, scodeObj);
                        ObjSetResult(interp, errorResultObj);
                        ObjDecrRefs(scodeObj);
                    }
                }
            }
            
            SysFreeString(einfo.bstrSource);
            SysFreeString(einfo.bstrDescription);
            SysFreeString(einfo.bstrHelpFile);
        } else {
            if ((hr == DISP_E_PARAMNOTFOUND  || hr == DISP_E_TYPEMISMATCH) &&
                badarg_index != -1) {
                /* Note parameter indices are backward (ie. from
                 * the Tcl perspective, numbered from the end) and 0-based,
                 * our error message parameter position is 1-based.
                 */
                ObjSetResult(interp,
                                 Tcl_ObjPrintf(
                                     "Parameter error. Offending parameter position %d.", nparams-badarg_index));
            }
            Twapi_AppendSystemError(interp, hr);
        }
    }

 vamoose:
    if (dispargP) {
        /*
         * We have to release VARIANT resources.
         * - if the VT_BYREF flag is set, do not do anything with
         *   the variant. The referenced variant will also be released if
         *   necessary in the loop and that is sufficient.
         * - VT_DISPATCH and VT_UNKNOWN - do not call VariantClear because
         *   that will decrement their ref count and potentially free
         *   them which we do not want as we are actually passing the
         *   objects up and are not actually done with them.
         * - VT_ARRAY and VT_BSTR - need to be released. Note for VT_ARRAY
         *   if it contains IDispatch or IUnknown, they would already
         *   have been AddRef'ed in the safearray extraction code and
         *   therefore releasing them is ok.
         * - VT_RECORD - TBD
         * - VT_* - need not clear, nothing to release
         */
        for (i = 0; i < nargalloc; ++i) {
            VARTYPE vt = V_VT(&dispargP[i]);
            if (vt == VT_BSTR ||
                ((vt & VT_ARRAY) && ! (vt & VT_BYREF)))
                VariantClear(&dispargP[i]);
        }
    }

    if (dispargP || paramflagsP)
        MemLifoPopFrame(&ticP->memlifo);

    return status;
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
    if (ObjGetElements(interp, namesObj, &nitems, &items) == TCL_ERROR)
        return TCL_ERROR;

    names = MemLifoPushFrame(&ticP->memlifo, nitems*sizeof(*names), NULL);

    for (i = 0; i < nitems; i++)
        names[i] = ObjToUnicode(items[i]);

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
        TwapiReturnErrorEx(interp, TWAPI_BUG,
                           Tcl_ObjPrintf("TwapiGetIDsOfNamesHelper: unknown ifc_type %d", ifc_type));
        goto vamoose;
    }

    if (SUCCEEDED(hr)) {
        Tcl_Obj *resultObj;
        int      i;

        resultObj = ObjNewList(0, NULL);
        for (i = 0; i < nitems; ++i) {
            ObjAppendElement(interp, resultObj, ObjFromUnicode(names[i]));
            ObjAppendElement(interp, resultObj, ObjFromLong(ids[i]));
        }
        ObjSetResult(interp, resultObj);
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
    objv[3] = ObjFromLong(taP->lcid);
    objv[4] = STRING_LITERAL_OBJ("dwReserved");
    objv[5] = ObjFromLong(taP->dwReserved);
    objv[6] = STRING_LITERAL_OBJ("memidConstructor");
    objv[7] = ObjFromLong(taP->memidConstructor);
    objv[8] = STRING_LITERAL_OBJ("memidDestructor");
    objv[9] = ObjFromLong(taP->memidDestructor);
    objv[10] = STRING_LITERAL_OBJ("lpstrSchema");
    objv[11] = ObjFromUnicode(taP->lpstrSchema ? taP->lpstrSchema : L"");
    objv[12] = STRING_LITERAL_OBJ("cbSizeInstance");
    objv[13] = ObjFromLong(taP->cbSizeInstance);
    objv[14] = STRING_LITERAL_OBJ("typekind");
    objv[15] = ObjFromLong(taP->typekind);
    objv[16] = STRING_LITERAL_OBJ("cFuncs");
    objv[17] = ObjFromLong(taP->cFuncs);
    objv[18] = STRING_LITERAL_OBJ("cVars");
    objv[19] = ObjFromLong(taP->cVars);
    objv[20] = STRING_LITERAL_OBJ("cImplTypes");
    objv[21] = ObjFromLong(taP->cImplTypes);
    objv[22] = STRING_LITERAL_OBJ("cbSizeVft");
    objv[23] = ObjFromLong(taP->cbSizeVft);
    objv[24] = STRING_LITERAL_OBJ("cbAlignment");
    objv[25] = ObjFromLong(taP->cbAlignment);
    objv[26] = STRING_LITERAL_OBJ("wTypeFlags");
    objv[27] = ObjFromLong(taP->wTypeFlags);
    objv[28] = STRING_LITERAL_OBJ("wMajorVerNum");
    objv[29] = ObjFromLong(taP->wMajorVerNum);
    objv[30] = STRING_LITERAL_OBJ("wMinorVerNum");
    objv[31] = ObjFromLong(taP->wMinorVerNum);
    objv[32] = STRING_LITERAL_OBJ("tdescAlias");
    objv[33] = NULL;
    if (taP->typekind == TKIND_ALIAS) {
        objv[33] = ObjFromTYPEDESC(interp, &taP->tdescAlias, tiP);
    }
    if (objv[33] == NULL)
        objv[33] = ObjNewList(0, NULL);
    objv[34] = STRING_LITERAL_OBJ("idldescType");
    objv[35] = ObjFromInt(taP->idldescType.wIDLFlags);

    tiP->lpVtbl->ReleaseTypeAttr(tiP, taP);
    ObjSetResult(interp, ObjNewList(36, objv));
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

    TwapiZeroMemory(names, sizeof(names));

    hr = tiP->lpVtbl->GetNames(tiP, memid, names, sizeof(names)/sizeof(names[0]), &name_count);

    if (SUCCEEDED(hr)) {
        Tcl_Obj *resultObj;
        int      i;

        resultObj = ObjNewList(0, NULL);
        for (i = 0; i < name_count; ++i) {
            ObjAppendElement(
                interp,
                resultObj,
                ObjFromUnicodeN(names[i], SysStringLen(names[i])));
            SysFreeString(names[i]);
            names[i] = NULL;
        }
        ObjSetResult(interp, resultObj);
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
        objv[3] = ObjFromLong(attrP->lcid);
        objv[4] = STRING_LITERAL_OBJ("syskind");
        objv[5] = ObjFromLong(attrP->syskind);
        objv[6] = STRING_LITERAL_OBJ("wMajorVerNum");
        objv[7] = ObjFromLong(attrP->wMajorVerNum);
        objv[8] = STRING_LITERAL_OBJ("wMinorVerNum");
        objv[9] = ObjFromLong(attrP->wMinorVerNum);
        objv[10] = STRING_LITERAL_OBJ("wLibFlags");
        objv[11] = ObjFromLong(attrP->wLibFlags);

        tlP->lpVtbl->ReleaseTLibAttr(tlP, attrP);

        ObjSetResult(interp, ObjNewList(12, objv));
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
    VARTYPE     vt, target_vt;
    int         len;
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
     * If this info is missing, we assume a BSTR IN parameter.
     * The third element, if present, is the parameter value passed
     * in as a named argument.
     *
     */
    paramc = 0;
    paramv = NULL;
    typec = 0;
    if (paramDescriptorP) {
        if (ObjGetElements(interp, paramDescriptorP, &paramc, &paramv) != TCL_OK)
            return TCL_ERROR;

        if (paramc >= 1) {
            /* The type information is a list of one or two elements */
            if (ObjGetElements(interp, paramv[0], &typec, &typev) != TCL_OK)
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
        vt = VT_BSTR; // TBD - maybe make this VT_VARIANT and guess
    }
    else if (paramc > 1) {
        /* If no value supplied as positional parameter, see if it is 
           supplied as named argument */
        if (valueObj == NULL && paramc > 2) {
            valueObj = paramv[2];
        }

        /* Get the flags and default value */
        if (ObjGetElements(interp, paramv[1], &paramfieldsc, &paramfields) != TCL_OK)
            goto vamoose;

        /* First field is the flags */
        if (paramfieldsc > 0) {
            if (ObjToInt(NULL, paramfields[0], &itemp) == TCL_OK) {
                *paramflagsP = (USHORT) (itemp ? itemp : PARAMFLAG_FIN);
            } else {
                /* Not an int, see if it is a token */
                char *s = ObjToStringN(paramfields[0], &len);
                if (len == 0 || STREQ(s, "in"))
                    *paramflagsP = PARAMFLAG_FIN;
                else if (STREQ(s, "out"))
                    *paramflagsP = PARAMFLAG_FOUT;
                else if (STREQ(s, "inout"))
                    *paramflagsP = PARAMFLAG_FOUT | PARAMFLAG_FIN;
                else {
                    ObjSetStaticResult(interp, "Unknown parameter modifiers");
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
            if (ObjListIndex(interp, paramfields[1], 1, &paramdefaultObj) != TCL_OK)
                goto vamoose;

            /* Note paramdefaultObj may be NULL even if return was TCL_OK */
            if (paramdefaultObj) {
                /* We hold on to this for a while, so make sure it does not
                   go away due to shimmering while the list holding it is
                   modified or changes type */
                ObjIncrRefs(paramdefaultObj);
            }
        }
        /* paramdefaultObj points to actual value (or NULL) at this point */
    }

    /* Note vt is what is allowed in a typedesc, not what is allowed in a
     * VARIANT
     */

    /* Note we have already checked previously that typec == 1 or 2 */
    if (typec == 2) {
        /* Only VT_PTR and VT_SAFEARRAY can have typec == 2 */

        /* Get the referenced type information. Note that automation does
         * not allow more than one level of indirection so we don't need
         * keep recursing for VT_PTR
         */
        if (ObjGetElements(interp, typev[1], &reftypec, &reftypev) != TCL_OK)
            goto invalid_type;
        if (reftypec == 0 || reftypec > 2 ||
            ObjToVT(interp, reftypev[0], &target_vt) != TCL_OK) {
            goto invalid_type;
        }

        if (vt == VT_PTR) {
            targetP = refvarP;
            /* What it points to must be a base type or a safearray */
            if (target_vt == VT_SAFEARRAY) {
                if (reftypec != 2)
                    goto invalid_type;
                /* Resolve the referenced safearray type */
                if (ObjGetElements(interp, reftypev[1], &reftypec, &reftypev) != TCL_OK)
                    goto invalid_type;
                /* This must be a base type */
                if (reftypec != 1 ||
                    ObjToVT(interp, reftypev[0], &target_vt) != TCL_OK) {
                    goto invalid_type;
                }
                target_vt = target_vt | VT_ARRAY;
            } else {
                /* Pointer to something other than safearray. */
                if (reftypec != 1)
                    goto invalid_type;
            }
            vt = target_vt | VT_BYREF;
        } else if (vt == VT_SAFEARRAY) {
            if (reftypec != 1)
                goto invalid_type; /* Safearrays only allow base types */
            targetP = varP;
            target_vt |= VT_ARRAY;
            vt = target_vt;
        } else
            goto invalid_type;  /* No other vt should have typec == 2 */
    } else {
        if (vt == VT_PTR || vt == VT_SAFEARRAY)
            goto invalid_type;
        targetP = varP;
        target_vt = vt;
    }

    /*
     * At this point,
     *  targetP points to the VARIANT where the param value will be stored
     *  target_vt is its associated type
     *  vt is the type of the "primary" VARIANT
     *     == target_vt or == (target_vt | VT_BYREF)
     */

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
        // TBD - Huh? What type conversion below ?
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
         * valueObj points to actual value. If this is NULL, parameter
         * better be optional. Note in versions before 3.2.2, we used
         * to set this to the default value specified in the prototype.
         * We now instead just set it to VT_ERROR since some methods
         * like Excel Range.AutoFilter differentiate behaviour on whether
         * the parameter was actually passed or defaulted.
         */
        if (valueObj == NULL) {
            /* No value or default supplied. Parameter better be optional */
            if (*paramflagsP & PARAMFLAG_FOPT) {
                targetP = varP; /* Reset to point to primary VARIANT */
                vt = VT_ERROR;  // Indicates optional param below
                target_vt = VT_ERROR;
            } else {
                ObjSetStaticResult(interp, "Missing value and no default for IDispatch invoke parameter");
                goto vamoose;
            }
        }

        /*
         * At this point, targetP points to the variant where the actual
         * value will be stored. This will be refvarP if VT_BYREF is set
         * and varP otherwise.
         */
        
        /* When vt is VARIANT we don't really know the type.
         * Note VT_VARIANT is only valid in type descriptions and
         * is not valid for VARIANTARG (ie. for
         * the actual value). It has to be a concrete type.
         *
         * To store it as an appropriate concrete VT type, we check to see
         * if it is internally a specific Tcl type. If so, we set vt
         * accordingly so it gets handled below. If not, vt will stay as
         * VT_VARIANT and we will make a best guess later.
         */
        if (target_vt == VT_VARIANT)
            target_vt = ObjTypeToVT(valueObj);
        
        TWAPI_ASSERT(valueObj != NULL || target_vt == VT_ERROR);

        if (ObjToVARIANT(interp, valueObj, targetP, target_vt) != TCL_OK)
            goto vamoose;
        target_vt = V_VT(targetP); /* Just to ensure consistency as it might have been changed */

        /*
         * If it is a IN or OUT param, no need to muck with ref counts.
         * If it is an INOUT though, we need to AddRef as the COM object
         * will call Release on it before overwriting the value and we
         * do not want our comobj target to be released.
         */
        if ((targetP->vt == VT_DISPATCH || targetP->vt == VT_UNKNOWN)
            &&
            (*paramflagsP & (PARAMFLAG_FIN | PARAMFLAG_FOUT)) == (PARAMFLAG_FIN | PARAMFLAG_FOUT)) {
            /* Both pdispVal and punkVal are really the same field and same
               vtbl layout so no need to distinguish  */
            if (targetP->punkVal != NULL)
                targetP->punkVal->lpVtbl->AddRef(targetP->punkVal);
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
        } else if (vt & VT_ARRAY) {
            V_VT(varP) = V_VT(targetP) | VT_BYREF;
            varP->pparray = &targetP->parray;
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
                ObjSetStaticResult(interp, "Internal error while constructing referenced VARIANT parameter");
                goto vamoose;
            }
        }
    }

    status = TCL_OK;

vamoose:
    if (paramdefaultObj)
        ObjDecrRefs(paramdefaultObj);

    return status;

invalid_type:
    ObjSetStaticResult(interp, "Unsupported or invalid type information format in parameter");
    goto vamoose;
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
        resultObj = ObjNewList(0, NULL);
        status = TCL_OK;
        break;

    case DESCKIND_FUNCDESC:
        objv[0] = STRING_LITERAL_OBJ("funcdesc");
        objv[1] = ObjFromFUNCDESC(interp, bind.lpfuncdesc, tiP);
        objv[2] = ObjFromOpaque(tiP, "ITypeInfo");
        tiP->lpVtbl->ReleaseFuncDesc(tiP, bind.lpfuncdesc);
        resultObj = ObjNewList(3, objv);
        status = TCL_OK;
        break;

    case DESCKIND_VARDESC:
        objv[0] = STRING_LITERAL_OBJ("vardesc");
        objv[1] = ObjFromVARDESC(interp, bind.lpvardesc, tiP);
        objv[2] = ObjFromOpaque(tiP, "ITypeInfo");
        tiP->lpVtbl->ReleaseVarDesc(tiP, bind.lpvardesc);
        resultObj = ObjNewList(3, objv);
        status = TCL_OK;
        break;

    case DESCKIND_TYPECOMP:
        objv[0] = STRING_LITERAL_OBJ("typecomp");
        objv[1] = ObjFromOpaque(bind.lptcomp, "ITypeComp");
        resultObj = ObjNewList(2, objv);
        status = TCL_OK;
        break;

    case DESCKIND_IMPLICITAPPOBJ: /* FALLTHRU */
    default:
        resultObj = STRING_LITERAL_OBJ("Unsupported ITypeComp desckind value");
        break;
    }

    ObjSetResult(interp, resultObj);
    return status;
}


int Twapi_IRecordInfo_GetFieldNames(Tcl_Interp *interp, IRecordInfo *riP)
{
    ULONG i, nbstrs;
    BSTR  bstrs[50];
    Tcl_Obj *objv[50];
    HRESULT hr;

    TwapiZeroMemory(bstrs, sizeof(bstrs));
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

    ObjSetResult(interp, ObjNewList(nbstrs, objv));
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
        objv[0] = ObjFromBoolean(1);
        objv[1] = ObjNewList(0, NULL);
        ObjSetResult(interp, ObjNewList(2, objv));
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
        ObjSetStaticResult(interp, "Unknown enum_type passed to TwapiIEnumNextHelper");
        return TCL_ERROR;
        
    }

    u.pv = MemLifoPushFrame(&ticP->memlifo, (DWORD) (count * elem_size), NULL);

    /*
     * Note, although these are output parameters, some COM objects expect
     * them to be initialized. For example, the COMAdminCollection
     * IEnumVARIANT interface (Bug 3185933)
     */
    switch (enum_type) {
    case 0:
        for (i = 0; i < count; ++i) {
            u.cdP[i].pUnk = NULL;
            u.cdP[i].dwCookie = 0;
        }
        hr = ifc.IEnumConnections->lpVtbl->Next(ifc.IEnumConnections,
                                                count, u.cdP, &ret_count);
        break;
    case 1:
        for (i = 0; i < count ; ++i) {
            VariantInit(&(u.varP[i]));
        }
        hr = ifc.IEnumVARIANT->lpVtbl->Next(ifc.IEnumVARIANT,
                                            count, u.varP, &ret_count);
        break;
    case 2:
        for (i = 0; i < count ; ++i) {
            u.icpP[i] = NULL;
        }
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

    objv[0] = ObjFromBoolean(hr == S_OK); // More to come?
    objv[1] = ObjNewList(0, NULL);
    for (i = 0; i < ret_count; ++i) {
        switch (enum_type) {
        case 0:
            ObjAppendElement(interp, objv[1],
                                     ObjFromCONNECTDATA(&(u.cdP[i])));
            break;
        case 1:
            ObjAppendElement(interp, objv[1],
                                     ObjFromVARIANT(&(u.varP[i]), flags));
            if (u.varP[i].vt == VT_BSTR)
                VariantClear(&(u.varP[i])); // TBD - any other types need clearing?
            break;
        case 3:
            ObjAppendElement(interp, objv[1],
                                     ObjFromOpaque(u.icpP[i],
                                                   "IConnectionPoint"));
            break;
        }
    }
        
    MemLifoPopFrame(&ticP->memlifo);
    ObjSetResult(interp, ObjNewList(2, objv));
    return TCL_OK;
}


/* Dispatcher for calling COM object methods */
int Twapi_CallCOMObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
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
        IPersistFile *persistfile;
    } ifc;
    HRESULT hr;
    TwapiResult result;
    DWORD dw1,dw2,dw3;
    BSTR bstr1 = NULL;          /* Initialize for tracking frees! */
    BSTR bstr2 = NULL;
    BSTR bstr3 = NULL;
    int tcl_status;
    void *pv;
    void *pv2;
    GUID guid, guid2;
    TYPEKIND tk;
    LPWSTR s;
    WORD w, w2;
    Tcl_Obj *objs[4];
    FUNCDESC *funcdesc;
    VARDESC  *vardesc;
    char *cP;
    int func = PtrToInt(clientdata);
    Tcl_Obj *sObj;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    --objc;
    ++objv;

    hr = S_OK;
    result.type = TRT_BADFUNCTIONCODE;

    if (func < 10000) {
        /* Interface based calls. func codes are all below 10000. Make sure
         * pointer is not null
         */
        if (ObjToLPVOID(interp, objv[0], &pv) != TCL_OK)
            return TCL_ERROR;
        if (pv == NULL) {
            ObjSetStaticResult(interp, "NULL interface pointer.");
            return TCL_ERROR;
        }
    }

    /* We want stronger type checking so we have to convert the interface
       pointer on a per interface type basis even though it is repetitive. */

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
            result.value.uval = ifc.unknown->lpVtbl->Release(ifc.unknown);
            break;
        case 2:
            result.type = TRT_DWORD;
            result.value.uval = ifc.unknown->lpVtbl->AddRef(ifc.unknown);
            break;
        case 3:
            if (objc < 3)
                goto badargs;
            hr = CLSIDFromString(ObjToUnicode(objv[1]), &guid);
            if (hr != S_OK)
                break;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = ObjToString(objv[2]);
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
        if (ObjToIDispatch(interp, objv[0], (void **)&ifc.dispatch) != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 101:
            result.type = TRT_DWORD;
            hr = ifc.dispatch->lpVtbl->GetTypeInfoCount(ifc.dispatch, &result.value.uval);
            break;
        case 102:
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETINT(dw1), GETINT(dw2),
                             ARGEND) != TCL_OK)
                goto ret_error;

            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeInfo";
            hr = ifc.dispatch->lpVtbl->GetTypeInfo(ifc.dispatch, dw1, dw2, (ITypeInfo **)&result.value.ifc.p);
            break;
        }
    } else if (func < 300) {
        /* IDispatchEx */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.dispatchex, "IDispatchEx")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 201: // GetDispID
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETVAR(bstr1, ObjToBSTR), GETINT(dw1),
                             ARGEND) != TCL_OK)
                goto ret_error;
            result.type = TRT_LONG;
            hr = ifc.dispatchex->lpVtbl->GetDispID(ifc.dispatchex, bstr1,
                                                   dw1, &result.value.ival);
            break;
        case 202: // GetMemberName
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
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
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETINT(dw1), GETINT(dw2),
                             ARGEND) != TCL_OK)
                goto ret_error;
            result.type = TRT_DWORD;
            if (func == 203)
                hr = ifc.dispatchex->lpVtbl->GetMemberProperties(
                    ifc.dispatchex, dw1, dw2, &result.value.uval);
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
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETVAR(bstr1, ObjToBSTR), GETINT(dw1),
                             ARGEND) != TCL_OK)
                goto ret_error;
            // hr = S_OK; -> Already set at top of function
            result.type = TRT_BOOL;
            result.value.bval = ifc.dispatchex->lpVtbl->DeleteMemberByName(
                ifc.dispatchex, bstr1, dw1) == S_OK ? 1 : 0;
            break;
        case 207: // DeleteMemberByDispID
            if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(dw1), ARGEND)
                != TCL_OK)
                goto ret_error;
            // hr = S_OK; -> Already set at top of function
            result.type = TRT_BOOL;
            result.value.bval = ifc.dispatchex->lpVtbl->DeleteMemberByDispID(
                ifc.dispatchex, dw1) == S_OK ? 1 : 0;
            break;

        }
    } else if (func < 400) {
        /* ITypeInfo */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.typeinfo, "ITypeInfo")
            != TCL_OK)
            return TCL_ERROR;

        /* Every other method either has no params or one integer param */
        if (objc > 2)
            goto badargs;

        dw1 = 0;
        if (objc == 2)
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);

        switch (func) {
        case 301: //GetRefTypeOfImplType
            result.type = TRT_DWORD;
            hr = ifc.typeinfo->lpVtbl->GetRefTypeOfImplType(
                ifc.typeinfo, dw1, &result.value.uval);
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
                objs[1] = ObjFromLong(dw1);
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
                objs[2] = ObjFromLong(dw2);
                objs[3] = ObjFromBSTR(bstr3);
                result.value.objv.nobj = 4;
                result.value.objv.objPP = objs;
            }
            break;
        case 306: // GetImplTypeFlags
            result.type = TRT_LONG;
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
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.typelib, "ITypeLib")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 401: // GetDocumentation
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_OBJV;
            hr = ifc.typelib->lpVtbl->GetDocumentation(
                ifc.typelib, dw1, &bstr1, &bstr2, &dw2, &bstr3);
            if (hr == S_OK) {
                objs[0] = ObjFromBSTR(bstr1);
                objs[1] = ObjFromBSTR(bstr2);
                objs[2] = ObjFromLong(dw2);
                objs[3] = ObjFromBSTR(bstr3);
                result.value.objv.nobj = 4;
                result.value.objv.objPP = objs;
            }
            break;
        case 402: // GetTypeInfoCount
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            result.value.uval =
                ifc.typelib->lpVtbl->GetTypeInfoCount(ifc.typelib);
            break;
        case 403: // GetTypeInfoType
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_LONG;
            hr = ifc.typelib->lpVtbl->GetTypeInfoType(ifc.typelib, dw1, &tk);
            if (hr == S_OK)
                result.value.ival = tk;
            break;
        case 404: // GetTypeInfo
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeInfo";
            hr = ifc.typelib->lpVtbl->GetTypeInfo(
                ifc.typelib, dw1, (ITypeInfo **)&result.value.ifc.p);
            break;
        case 405: // GetTypeInfoOfGuid
            if (objc != 2)
                goto badargs;
            hr = CLSIDFromString(ObjToUnicode(objv[1]), &guid);
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
            if (objc != 3)
                goto badargs;
            s = ObjToUnicode(objv[2]);
            NULLIFY_EMPTY(s);
            result.type = TRT_EMPTY;
            hr = RegisterTypeLib(ifc.typelib, ObjToUnicode(objv[1]), s);
            break;
        }
    } else if (func < 600) {
        /* IRecordInfo */
        /* TBD - for record data, we should create dummy type-safe pointers
           instead of passing around voids even though that is what
           the IRecordInfo interface does */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.recordinfo, "IRecordInfo")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 501: // GetField
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETVOIDP(pv), GETOBJ(sObj),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            VariantInit(&result.value.var);
            result.type = TRT_VARIANT;
            hr = ifc.recordinfo->lpVtbl->GetField(
                ifc.recordinfo, pv, ObjToUnicode(sObj), &result.value.var);
            break;
        case 502: // GetGuid
            if (objc != 1)
                goto badargs;
            result.type = TRT_GUID;
            hr = ifc.recordinfo->lpVtbl->GetGuid(ifc.recordinfo, &result.value.guid);
            break;
        case 503: // GetName
            if (objc != 1)
                goto badargs;
            VariantInit(&result.value.var);
            hr = ifc.recordinfo->lpVtbl->GetName(ifc.recordinfo, &result.value.var.bstrVal);
            if (hr == S_OK) {
                result.type = VT_VARIANT;
                result.value.var.vt = VT_BSTR;
            }
            break;
        case 504: // GetSize
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.recordinfo->lpVtbl->GetSize(ifc.recordinfo, &result.value.uval);
            break;
        case 505: // GetTypeInfp
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeInfo";
            hr = ifc.recordinfo->lpVtbl->GetTypeInfo(
                ifc.recordinfo, (ITypeInfo **)&result.value.ifc.p);
            break;
        case 506: // IsMatchingType
            if (objc != 2)
                goto badargs;
            if (ObjToOpaque(interp, objv[1], &pv, "IRecordInfo") != TCL_OK)
                goto ret_error;
            result.type = TRT_BOOL;
            result.value.bval = ifc.recordinfo->lpVtbl->IsMatchingType(
                ifc.recordinfo, (IRecordInfo *) pv);
            break;
        case 507: // RecordClear
            if (objc != 2)
                goto badargs;
            if (ObjToLPVOID(interp, objv[1], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordClear(ifc.recordinfo, pv);
            break;
        case 508: // RecordCopy
            if (objc != 3)
                goto badargs;
            if (ObjToLPVOID(interp, objv[1], &pv) != TCL_OK &&
                ObjToLPVOID(interp, objv[2], &pv2) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordCopy(ifc.recordinfo, pv, pv2);
            break;
        case 509: // RecordCreate
            result.type = TRT_PTR;
            result.value.ptr.p =
                ifc.recordinfo->lpVtbl->RecordCreate(ifc.recordinfo);
            result.value.ptr.name = "void*";
            break;
        case 510: // RecordCreateCopy
            if (objc != 2)
                goto badargs;
            if (ObjToLPVOID(interp, objv[1], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_PTR;
            result.value.ptr.name = "void*";
            hr = ifc.recordinfo->lpVtbl->RecordCreateCopy(ifc.recordinfo, pv,
                                                          &result.value.ptr.p);
            break;
        case 511: // RecordDestroy
            if (objc != 2)
                goto badargs;
            if (ObjToLPVOID(interp, objv[3], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordDestroy(ifc.recordinfo, pv);
            break;
        case 512: // RecordInit
            if (objc != 2)
                goto badargs;
            if (ObjToLPVOID(interp, objv[1], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_EMPTY;
            hr = ifc.recordinfo->lpVtbl->RecordInit(ifc.recordinfo, pv);
            break;
        case 513:
            if (objc != 1)
                goto badargs;
            return Twapi_IRecordInfo_GetFieldNames(interp, ifc.recordinfo);
        }
    } else if (func < 700) {
        /* IMoniker */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.moniker, "IMoniker")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 601: // GetDisplayName
            if (objc != 3)
                goto badargs;
            if (ObjToOpaque(interp, objv[1], &pv, "IBindCtx") != TCL_OK)
                goto ret_error;
            if (ObjToOpaque(interp, objv[2], &pv2, "IMoniker") != TCL_OK)
                goto ret_error;
            result.type = TRT_LPOLESTR;
            hr = ifc.moniker->lpVtbl->GetDisplayName(
                ifc.moniker, (IBindCtx *)pv, (IMoniker *)pv2,
                &result.value.lpolestr);
            break;
        }
    } else if (func < 800) {
        /* IEnumVARIANT */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.enumvariant, "IEnumVARIANT")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 701: // Clone
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumVARIANT";
            hr = ifc.enumvariant->lpVtbl->Clone(
                ifc.enumvariant, (IEnumVARIANT **)&result.value.ifc.p);
            break;
        case 702: // Reset
            if (objc != 1)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumvariant->lpVtbl->Reset(ifc.enumvariant);
            break;
        case 703: // Skip
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.enumvariant->lpVtbl->Skip(ifc.enumvariant, dw1);
            break;
        }        
    } else if (func < 900) {
        /* IConnectionPoint */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.connectionpoint,
                        "IConnectionPoint")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 801: // Advise
            if (objc != 2)
                goto badargs;
            if (ObjToIUnknown(interp, objv[1], &pv) != TCL_OK)
                goto ret_error;
            result.type = TRT_DWORD;
            hr = ifc.connectionpoint->lpVtbl->Advise(
                ifc.connectionpoint, (IUnknown *)pv, &result.value.uval);
            break;
            
        case 802: // EnumConnections
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnections";
            hr = ifc.connectionpoint->lpVtbl->EnumConnections(ifc.connectionpoint, (IEnumConnections **) &result.value.ifc.p);
            break;

        case 803: // GetConnectionInterface
            if (objc != 1)
                goto badargs;
            result.type = TRT_GUID;
            hr = ifc.connectionpoint->lpVtbl->GetConnectionInterface(ifc.connectionpoint, &result.value.guid);
            break;

        case 804: // GetConnectionPointContainer
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IConnectionPointContainer";
            hr = ifc.connectionpoint->lpVtbl->GetConnectionPointContainer(ifc.connectionpoint, (IConnectionPointContainer **)&result.value.ifc.p);
            break;

        case 805: // Unadvise
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.connectionpoint->lpVtbl->Unadvise(ifc.connectionpoint, dw1);
            break;
        }
    } else if (func < 1000) {
        /* IConnectionPointContainer */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.connectionpointcontainer,
                        "IConnectionPointContainer") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 901: // EnumConnectionPoints
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnectionPoints";
            hr = ifc.connectionpointcontainer->lpVtbl->EnumConnectionPoints(
                ifc.connectionpointcontainer,
                (IEnumConnectionPoints **)&result.value.ifc.p);
            break;
        case 902: // FindConnectionPoint
            if (objc != 2)
                goto badargs;
            hr = CLSIDFromString(ObjToUnicode(objv[1]), &guid);
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
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.enumconnectionpoints,
                        "IEnumConnectionPoints") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 1001: // Clone
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnectionPoints";
            hr = ifc.enumconnectionpoints->lpVtbl->Clone(
                ifc.enumconnectionpoints,
                (IEnumConnectionPoints **) &result.value.ifc.p);
            break;
        case 1002: // Reset
            if (objc != 1)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumconnectionpoints->lpVtbl->Reset(
                ifc.enumconnectionpoints);
            break;
        case 1003: // Skip
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.enumconnectionpoints->lpVtbl->Skip(
                ifc.enumconnectionpoints,   dw1);
            break;
        }        
    } else if (func < 1200) {
        /* IEnumConnections */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.enumconnections,
                        "IEnumConnections") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 1101: // Clone
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumConnections";
            hr = ifc.enumconnections->lpVtbl->Clone(
                ifc.enumconnections, (IEnumConnections **)&result.value.ifc.p);
            break;
        case 1102: // Reset
            if (objc != 1)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumconnections->lpVtbl->Reset(ifc.enumconnections);
            break;
        case 1103: // Skip
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.enumconnections->lpVtbl->Skip(ifc.enumconnections, dw1);
            break;
        }        
    } else if (func == 1201) {
        /* IProvideClassInfo */
        /* We accept both IProvideClassInfo and IProvideClassInfo2 interfaces */
        if (ObjToOpaque(NULL, objv[0], (void **)&ifc.provideclassinfo,
                        "IProvideClassInfo") != TCL_OK &&
            ObjToOpaque(interp, objv[0], (void **)&ifc.provideclassinfo,
                        "IProvideClassInfo2") != TCL_OK)
            return TCL_ERROR;

        result.type = TRT_INTERFACE;
        result.value.ifc.name = "ITypeInfo";
        hr = ifc.provideclassinfo->lpVtbl->GetClassInfo(
            ifc.provideclassinfo, (ITypeInfo **)&result.value.ifc.p);
    } else if (func == 1301) {
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.provideclassinfo2,
                        "IProvideClassInfo2")
            != TCL_OK)
            return TCL_ERROR;

        if (objc != 2)
            goto badargs;
        CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
        result.type = TRT_GUID;
        hr = ifc.provideclassinfo2->lpVtbl->GetGUID(ifc.provideclassinfo2,
                                                    dw1, &result.value.guid);
    } else if (func < 1500) {
        /* ITypeComp */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.typecomp, "ITypeComp")
            != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 1401:
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETOBJ(sObj), GETINT(dw1), GETWORD(w),
                             ARGEND) != TCL_OK)
                goto ret_error;
            return Twapi_ITypeComp_Bind(interp, ifc.typecomp,
                                        ObjToUnicode(sObj), dw1, w);
        }
    } else if (func < 5600) {
        /* IPersistFile */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.persistfile,
                        "IPersistFile") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5501: // GetCurFile
            if (objc != 1)
                goto badargs;
            hr = ifc.persistfile->lpVtbl->GetCurFile(
                ifc.persistfile, &result.value.lpolestr);
            if (hr != S_OK && hr != S_FALSE)
                break;
            /* Note S_FALSE also is a success return */
            result.type = TRT_OBJV;
            objs[0] = ObjFromLong(hr);
            objs[1] = ObjFromUnicode(result.value.lpolestr);
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            break;
        case 5502: // IsDirty
            if (objc != 1)
                goto badargs;
            result.type = TRT_BOOL;
            result.value.bval = ifc.persistfile->lpVtbl->IsDirty(ifc.persistfile) == S_OK;
            break;
        case 5503: // Load
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETOBJ(sObj), GETINT(dw1),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.persistfile->lpVtbl->Load(ifc.persistfile,
                                               ObjToUnicode(sObj), dw1);
            break;
        case 5504: // Save
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETOBJ(sObj), GETINT(dw1),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.persistfile->lpVtbl->Save(ifc.persistfile, ObjToLPWSTR_NULL_IF_EMPTY(sObj), dw1);
            break;
        case 5505: // SaveCompleted
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.persistfile->lpVtbl->SaveCompleted(
                ifc.persistfile, ObjToUnicode(objv[1]));
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
            hr = CreateFileMoniker(ObjToUnicode(objv[0]),
                                   (IMoniker **)&result.value.ifc.p);
            break;
        case 10002: // CreateBindCtx
            CHECK_INTEGER_OBJ(interp, dw1, objv[0]);
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IBindCtx";
            hr = CreateBindCtx(dw1, (IBindCtx **)&result.value.ifc.p);
            break;
        case 10003: // GetRecordInfoFromGuids
            if (TwapiGetArgs(interp, objc, objv,
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
            if (TwapiGetArgs(interp, objc, objv,
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
            if (TwapiGetArgs(interp, objc, objv,
                             GETGUID(guid), GETWORD(w), GETWORD(w2),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = UnRegisterTypeLib(&guid, w, w2, dw3, SYS_WIN32);
            break;
        case 10006: // LoadRegTypeLib
            if (TwapiGetArgs(interp, objc, objv,
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
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITypeLib";
            hr = LoadTypeLibEx(ObjToUnicode(objv[0]), dw1,
                                              (ITypeLib **)&result.value.ifc.p);
            break;
        case 10008: // CoGetObject
            if (TwapiGetArgs(interp, objc, objv,
                             GETOBJ(sObj), ARGSKIP, GETGUID(guid), GETASTR(cP),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (ObjListLength(interp, objv[1], &dw1) == TCL_ERROR ||
                dw1 != 0) {
                ObjSetStaticResult(interp, "Bind options are not supported for CoGetOjbect and must be specified as empty."); //TBD
                goto ret_error;
            }
            result.type = TRT_INTERFACE;
            result.value.ifc.name = cP;
            hr = CoGetObject(ObjToUnicode(sObj),
                             NULL, &guid, &result.value.ifc.p);
            break;
        case 10009: // GetActiveObject
            if (objc != 1)
                goto badargs;
            if (ObjToGUID(interp, objv[0], &guid) != TCL_OK)
                goto ret_error;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IUnknown";
            hr = GetActiveObject(&guid, NULL, (IUnknown **)&result.value.ifc.p);
            break;
        case 10010: // ProgIDFromCLSID
            if (objc != 1)
                goto badargs;
            if (ObjToGUID(interp, objv[0], &guid) != TCL_OK)
                goto ret_error;
            result.type = TRT_LPOLESTR;
            hr = ProgIDFromCLSID(&guid, &result.value.lpolestr);
            break;
        case 10011:  // CLSIDFromProgID
            if (objc != 1)
                goto badargs;
            result.type = TRT_GUID;
            hr = CLSIDFromProgID(ObjToUnicode(objv[0]), &result.value.guid);
            break;
        case 10012: // CoCreateInstance
            if (TwapiGetArgs(interp, objc, objv,
                             GETGUID(guid), ARGSKIP, GETINT(dw1),
                             GETGUID(guid2), GETASTR(cP),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (ObjToIUnknown(interp, objv[1], (void **)&ifc.unknown)
                != TCL_OK)
                goto ret_error;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = cP;
            hr = CoCreateInstance(&guid, ifc.unknown, dw1, &guid2,
                                  &result.value.ifc.p);
            break;
        }
    }

    if (hr != S_OK) {
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = hr;
    }

    /* Note when hr == 0, result.type can be BADFUNCTION code! */
    tcl_status = TwapiSetResult(interp, &result);

vamoose:
    // Free bstr AFTER setting result as result.value.unicode may point to it */
    SysFreeString(bstr1);        /* OK if bstr is NULL */
    SysFreeString(bstr2);
    SysFreeString(bstr3);
    return tcl_status;

badargs:
    TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

ret_error:
    tcl_status = TCL_ERROR;
    goto vamoose;
}


/* Dispatcher for calling COM object methods that require a TwapiInterpContext */
int Twapi_CallCOMTicObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    void *ifc;
    int func;
    HRESULT hr;
    TwapiResult result;
    DWORD dw1,dw2;

    hr = S_OK;
    result.type = TRT_BADFUNCTIONCODE;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), ARGSKIP,
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    // ARGSKIP makes sure at least one more argument

    /* We want stronger type checking so we have to convert the interface
       pointer on a per interface type basis even though it is repetitive. */

    switch (func) {
    case 1:   /* IDispatch.GetIDsOfNames */
        /* We accept both IDispatch and IDispatchEx interfaces here */
        if (ObjToIDispatch(interp, objv[2], &ifc) != TCL_OK)
            return TCL_ERROR;
        if (ifc == NULL)
            goto null_interface_error;
        if (objc != 5)
            goto badargs;
        CHECK_INTEGER_OBJ(interp, dw1, objv[4]);
        return TwapiGetIDsOfNamesHelper(
            ticP, ifc, objv[3],
            dw1,              /* LCID */
            0);             /* 0->IDispatch interface */
    case 2: /* ITypeInfo.GetIDsOfNames */
        if (ObjToOpaque(interp, objv[2], &ifc, "ITypeInfo") != TCL_OK)
            return TCL_ERROR;
        if (ifc == NULL)
            goto null_interface_error;
        if (objc != 4)
            goto badargs;
        return TwapiGetIDsOfNamesHelper(
            ticP, ifc, objv[3],
            0,              /* Unused */
            1);             /* 1->ITypeInfo interface */
    case 3:  /* IEnumVARIANT.Next */
        if (ObjToOpaque(interp, objv[2], &ifc, "IEnumVARIANT")
            != TCL_OK)
            return TCL_ERROR;
        if (ifc == NULL)
            goto null_interface_error;
        if (objc < 4)
            goto badargs;
        CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
        /* Let caller decide if he wants flat untagged variant list
           or a tagged variant value */
        dw2 = 0;            /* By default tagged value */
        if (objc > 4)
            CHECK_INTEGER_OBJ(interp, dw2, objv[4]);
        return TwapiIEnumNextHelper(ticP,ifc,dw1,1,dw2);

    case 4: /* IEnumConnectionPoints.Next */
        if (ObjToOpaque(interp, objv[2], &ifc, "IEnumConnectionPoints") != TCL_OK)
            return TCL_ERROR;
        if (ifc == NULL)
            goto null_interface_error;
        if (objc != 4)
            goto badargs;
        CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
        return TwapiIEnumNextHelper(ticP,ifc,
                                    dw1, 2, 0);
    case 5: /* IEnumConnections.Next */
        if (ObjToOpaque(interp, objv[2], &ifc, "IEnumConnections") != TCL_OK)
            return TCL_ERROR;
        if (ifc == NULL)
            goto null_interface_error;
        if (objc != 4)
            goto badargs;
        CHECK_INTEGER_OBJ(interp, dw1, objv[3]);
        return TwapiIEnumNextHelper(ticP, ifc, dw1, 0, 0);
    }

    if (hr != S_OK) {
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = hr;
    }

    /* Note when hr == 0, result.type can be BADFUNCTION code! */
    return TwapiSetResult(interp, &result);

badargs:
    return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

null_interface_error:
    ObjSetStaticResult(interp, "NULL interface pointer.");
    return TCL_ERROR;
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




static int TwapiComInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* TBD - break apart commands that do not require a ticP */

    static struct fncode_dispatch_s ComDispatch[] = {
        DEFINE_FNCODE_CMD(IUnknown_Release, 1),
        DEFINE_FNCODE_CMD(IUnknown_AddRef, 2),
        DEFINE_FNCODE_CMD(Twapi_IUnknown_QueryInterface, 3),
        DEFINE_FNCODE_CMD(OleRun, 4), // Note - function, NOT method

        DEFINE_FNCODE_CMD(IDispatch_GetTypeInfoCount, 101),
        DEFINE_FNCODE_CMD(IDispatch_GetTypeInfo, 102),

        DEFINE_FNCODE_CMD(IDispatchEx_GetDispID, 201),
        DEFINE_FNCODE_CMD(IDispatchEx_GetMemberName, 202),
        DEFINE_FNCODE_CMD(IDispatchEx_GetMemberProperties, 203),
        DEFINE_FNCODE_CMD(IDispatchEx_GetNextDispID, 204),
        DEFINE_FNCODE_CMD(IDispatchEx_GetNameSpaceParent, 205),
        DEFINE_FNCODE_CMD(IDispatchEx_DeleteMemberByName, 206),
        DEFINE_FNCODE_CMD(IDispatchEx_DeleteMemberByDispID, 207),

        DEFINE_FNCODE_CMD(ITypeInfo_GetRefTypeOfImplType, 301),
        DEFINE_FNCODE_CMD(ITypeInfo_GetRefTypeInfo, 302),
        DEFINE_FNCODE_CMD(ITypeInfo_GetTypeComp, 303),
        DEFINE_FNCODE_CMD(ITypeInfo_GetContainingTypeLib, 304),
        DEFINE_FNCODE_CMD(ITypeInfo_GetDocumentation, 305),
        DEFINE_FNCODE_CMD(ITypeInfo_GetImplTypeFlags, 306),
        DEFINE_FNCODE_CMD(GetRecordInfoFromTypeInfo, 307), // Note - function, not method
        DEFINE_FNCODE_CMD(ITypeInfo_GetNames, 308),
        DEFINE_FNCODE_CMD(ITypeInfo_GetTypeAttr, 309),
        DEFINE_FNCODE_CMD(ITypeInfo_GetFuncDesc, 310),
        DEFINE_FNCODE_CMD(ITypeInfo_GetVarDesc, 311),

        DEFINE_FNCODE_CMD(ITypeLib_GetDocumentation, 401),
        DEFINE_FNCODE_CMD(ITypeLib_GetTypeInfoCount, 402),
        DEFINE_FNCODE_CMD(ITypeLib_GetTypeInfoType, 403),
        DEFINE_FNCODE_CMD(ITypeLib_GetTypeInfo, 404),
        DEFINE_FNCODE_CMD(ITypeLib_GetTypeInfoOfGuid, 405),
        DEFINE_FNCODE_CMD(ITypeLib_GetLibAttr, 406),
        DEFINE_FNCODE_CMD(RegisterTypeLib, 407), // Function, not method

        DEFINE_FNCODE_CMD(IRecordInfo_GetField, 501),
        DEFINE_FNCODE_CMD(IRecordInfo_GetGuid, 502),
        DEFINE_FNCODE_CMD(IRecordInfo_GetName, 503),
        DEFINE_FNCODE_CMD(IRecordInfo_GetSize, 504),
        DEFINE_FNCODE_CMD(IRecordInfo_GetTypeInfo, 505),
        DEFINE_FNCODE_CMD(IRecordInfo_IsMatchingType, 506),
        DEFINE_FNCODE_CMD(IRecordInfo_RecordClear, 507),
        DEFINE_FNCODE_CMD(IRecordInfo_RecordCopy, 508),
        DEFINE_FNCODE_CMD(IRecordInfo_RecordCreate, 509),
        DEFINE_FNCODE_CMD(IRecordInfo_RecordCreateCopy, 510),
        DEFINE_FNCODE_CMD(IRecordInfo_RecordDestroy, 511),
        DEFINE_FNCODE_CMD(IRecordInfo_RecordInit, 512),
        DEFINE_FNCODE_CMD(IRecordInfo_GetFieldNames, 513),

        DEFINE_FNCODE_CMD(IMoniker_GetDisplayName,601),

        DEFINE_FNCODE_CMD(IEnumVARIANT_Clone, 701),
        DEFINE_FNCODE_CMD(IEnumVARIANT_Reset, 702),
        DEFINE_FNCODE_CMD(IEnumVARIANT_Skip, 703),

        DEFINE_FNCODE_CMD(IConnectionPoint_Advise, 801),
        DEFINE_FNCODE_CMD(IConnectionPoint_EnumConnections, 802),
        DEFINE_FNCODE_CMD(IConnectionPoint_GetConnectionInterface, 803),
        DEFINE_FNCODE_CMD(IConnectionPoint_GetConnectionPointContainer, 804),
        DEFINE_FNCODE_CMD(IConnectionPoint_Unadvise, 805),

        DEFINE_FNCODE_CMD(IConnectionPointContainer_EnumConnectionPoints, 901),
        DEFINE_FNCODE_CMD(IConnectionPointContainer_FindConnectionPoint, 902),

        DEFINE_FNCODE_CMD(IEnumConnectionPoints_Clone, 1001),
        DEFINE_FNCODE_CMD(IEnumConnectionPoints_Reset, 1002),
        DEFINE_FNCODE_CMD(IEnumConnectionPoints_Skip, 1003),

        DEFINE_FNCODE_CMD(IEnumConnections_Clone, 1101),
        DEFINE_FNCODE_CMD(IEnumConnections_Reset, 1102),
        DEFINE_FNCODE_CMD(IEnumConnections_Skip, 1103),

        DEFINE_FNCODE_CMD(IProvideClassInfo_GetClassInfo, 1201),

        DEFINE_FNCODE_CMD(IProvideClassInfo2_GetGUID, 1301),

        DEFINE_FNCODE_CMD(ITypeComp_Bind, 1401),


        DEFINE_FNCODE_CMD(IPersistFile_GetCurFile, 5501),
        DEFINE_FNCODE_CMD(IPersistFile_IsDirty, 5502),
        DEFINE_FNCODE_CMD(IPersistFile_Load, 5503),
        DEFINE_FNCODE_CMD(IPersistFile_Save, 5504),
        DEFINE_FNCODE_CMD(IPersistFile_SaveCompleted, 5505),

        DEFINE_FNCODE_CMD(CreateFileMoniker, 10001),
        DEFINE_FNCODE_CMD(CreateBindCtx, 10002),
        DEFINE_FNCODE_CMD(GetRecordInfoFromGuids, 10003),
        DEFINE_FNCODE_CMD(QueryPathOfRegTypeLib, 10004),
        DEFINE_FNCODE_CMD(UnRegisterTypeLib, 10005),
        DEFINE_FNCODE_CMD(LoadRegTypeLib, 10006),
        DEFINE_FNCODE_CMD(LoadTypeLibEx, 10007),
        DEFINE_FNCODE_CMD(Twapi_CoGetObject, 10008),
        DEFINE_FNCODE_CMD(GetActiveObject, 10009),
        DEFINE_FNCODE_CMD(ProgIDFromCLSID, 10010),
        DEFINE_FNCODE_CMD(CLSIDFromProgID, 10011),
        DEFINE_FNCODE_CMD(Twapi_CoCreateInstance, 10012),
        
    };

    static struct alias_dispatch_s ComAliasDispatch[] = {
        DEFINE_ALIAS_CMD(IDispatch_GetIDsOfNames, 1),
        DEFINE_ALIAS_CMD(ITypeInfo_GetIDsOfNames, 2),
        DEFINE_ALIAS_CMD(IEnumVARIANT_Next, 3),
        DEFINE_ALIAS_CMD(IEnumConnectionPoints_Next, 4),
        DEFINE_ALIAS_CMD(IEnumConnections_Next, 5),
    };

    Tcl_CreateObjCommand(interp, "twapi::ComTicCall",
                         Twapi_CallCOMTicObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::IDispatch_Invoke",
                         Twapi_IDispatch_InvokeObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::ComEventSink",
                         Twapi_ComEventSinkObjCmd, ticP, NULL);

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(ComDispatch), ComDispatch, Twapi_CallCOMObjCmd);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(ComAliasDispatch), ComAliasDispatch, "twapi::ComTicCall");

    return TCL_OK;
}


#ifndef TWAPI_SINGLE_MODULE
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_SINGLE_MODULE
__declspec(dllexport) 
#endif
int Twapi_com_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

