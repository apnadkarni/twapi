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

static ULONG STDMETHODCALLTYPE Twapi_EventSink_AddRef(IDispatch *this)
{
    ((Twapi_EventSink *)this)->refc += 1;
    return ((Twapi_EventSink *)this)->refc;
}

ULONG STDMETHODCALLTYPE Twapi_EventSink_Release(IDispatch *this)
{
    Twapi_EventSink *me = (Twapi_EventSink *) this;

    me->refc -= 1;
    if (((Twapi_EventSink *)this)->refc == 0) {
        if (me->interp)
            Tcl_Release(me->interp);
        if (me->cmd)
            Tcl_DecrRefCount(me->cmd);
        free(this);
        return 0;
    } else
        return ((Twapi_EventSink *)this)->refc;
}

HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetTypeInfoCount
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


HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetTypeInfo(
    IDispatch *this,
    UINT tinfo,
    LCID lcid,
    ITypeInfo **tiPP)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE Twapi_EventSink_GetIDsOfNames(
    IDispatch *this,
    REFIID   riid,
    LPOLESTR *namesP,
    UINT namesc,
    LCID lcid,
    DISPID *rgDispId)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE Twapi_EventSink_Invoke(
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
    cmdobjv = malloc((cmdobjc+4) * sizeof(*cmdobjv));
    if (cmdobjv == NULL) {
        return E_OUTOFMEMORY;
    }

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
        Tcl_IncrRefCount(cmdobjv[i]); /* Protect while we are using it */
    }

    /*
     * Before eval'ing, addref ourselves so we don't get deleted in a
     * recursive callback
     */
    this->lpVtbl->AddRef(this);

    /* If hr is not TCL_OK, it is a HRESULT error code */
    Tcl_SaveResult(me->interp, &savedresult);
    hr = Tcl_EvalObjv(me->interp, cmdobjc, cmdobjv, TCL_EVAL_GLOBAL);
    if (hr != TCL_OK) {
        if (excepP) {
            memset(excepP, 0, sizeof(*excepP));
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

    if (Twapi_malloc(interp, NULL, sizeof(*sinkP), (void **)&sinkP) != TCL_OK)
        return TCL_ERROR;

    /* Fill in the cmdargs slots from the arguments */
    sinkP->idispP.lpVtbl = &Twapi_EventSink_Vtbl;
    sinkP->iid = iid;
    Tcl_Preserve(interp);
    sinkP->interp = interp;
    sinkP->refc = 1;
    Tcl_IncrRefCount(objv[2]);
    sinkP->cmd = objv[2];

    Tcl_SetObjResult(interp, ObjFromIUnknown(sinkP));


    return TCL_OK;
}




int Twapi_IDispatch_InvokeObjCmd(
    ClientData dummy,		/* Not used. */
    Tcl_Interp *interp,		/* Current interpreter. */
    int objc,			/* Number of arguments. */
    Tcl_Obj *CONST objv[])	/* Argument objects. */
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
     *   1 - RIID - ignored, always set to NULL
     *   2 - LCID
     *   3 - flags
     *   4 - return type
     *   5 - (optional) list of parameters to dispatch function. This is a list
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
    if (protoc < 5) {
        Tcl_SetResult(interp, "Invalid IDispatch prototype - must contain DISPID RIID LCID FLAGS RETTYPE ?PARAMTYPES?", TCL_STATIC);
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

    if (protoc >= 6) {
        /* Extract the parameter information */
        if (Tcl_ListObjGetElements(interp, protov[5], &nparams, &params) != TCL_OK)
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
    dispargP = malloc(nargalloc * sizeof(*dispargP));
    paramflagsP = malloc((nparams+1)*sizeof(*paramflagsP));
    if (dispargP == NULL || paramflagsP == NULL) {
        Tcl_SetResult(interp, "Unable to allocate memory", TCL_STATIC);
        return TCL_ERROR;
    }
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
    memset(&einfo, 0, sizeof(einfo));
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
                errorResultObj = Tcl_GetObjResult(interp);
                Tcl_AppendUnicodeToObj(errorResultObj, L" ", 1);
                Tcl_AppendUnicodeToObj(errorResultObj,
                                       einfo.bstrDescription,
                                       SysStringLen(einfo.bstrDescription));
            } else {
                /* No error description. Perhaps the scode field
                 * tells us something more.
                 */
                if (einfo.scode &&
                    (FACILITY_WIN32 == HRESULT_FACILITY(einfo.scode) ||
                     FACILITY_WINDOWS == HRESULT_FACILITY(einfo.scode) ||
                     FACILITY_DISPATCH == HRESULT_FACILITY(einfo.scode) ||
                     FACILITY_RPC == HRESULT_FACILITY(einfo.scode))) {
                    WCHAR *scode_msg = Twapi_MapWindowsErrorToString(einfo.scode);
                    if (scode_msg) {
                        errorResultObj = Tcl_GetObjResult(interp);
                        Tcl_AppendUnicodeToObj(errorResultObj, L" ", 1);
                        Tcl_AppendUnicodeToObj(errorResultObj, scode_msg, -1);
                        free(scode_msg);
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
                _snprintf(buf, ARRAYSIZE(buf), "%d", badarg_index);
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

        free(dispargP);
    }
    if (paramflagsP)
        free(paramflagsP);

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
