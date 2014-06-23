/* 
 * Copyright (c) 2014, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 *
 * COM server implementation
 */

#include "twapi.h"
#include "twapi_com.h"

/*
 * TBD - IMPLEMENTING COM SERVER 
 HINTS
Implementing IDispatch can be easy or hard. (Assuming you cannot use ATL).

The easy way is to not support TypeInfo (return 0 from GetTypeInfoCount and E_NOTIMPL from GetTypeInfo. Nobody should call it.).

Then all you have to support is GetIDsOfNames and Invoke. It's just a big lookup table essentially.

For GetIDsOfNames, return DISP_E_UNKNOWNNAME if cNames != 1. You aren't going to support argument names. Then you just have to lookup rgszNames[0] in your mapping of names-to-ids.

Finally, implement Invoke. Ignore everything except pDispParams and pVarResult. Use VariantChangeType to coerce the parameters to the types you expect, and pass to your implementation. Set the return value and return. Done.

The hard way is to use ITypeInfo and all that. I've never done it and wouldn't. ATL makes it easy so just use ATL.

If picking the hard way, Good luck.
*/


/*
 * IDispatch server implementation
 */
HRESULT STDMETHODCALLTYPE Twapi_ComServer_QueryInterface(
    IDispatch *this,
    REFIID riid,
    void **ifcPP);
ULONG STDMETHODCALLTYPE Twapi_ComServer_AddRef(IDispatch *this);
ULONG STDMETHODCALLTYPE Twapi_ComServer_Release(IDispatch *this);
HRESULT STDMETHODCALLTYPE Twapi_ComServer_GetTypeInfoCount(
    IDispatch *this,
    UINT *pctP);
HRESULT STDMETHODCALLTYPE Twapi_ComServer_GetTypeInfo(
    IDispatch *this,
    UINT tinfo,
    LCID lcid,
    ITypeInfo **tiPP);
HRESULT STDMETHODCALLTYPE Twapi_ComServer_GetIDsOfNames(
    IDispatch *this,
    REFIID   riid,
    LPOLESTR *namesP,
    UINT namesc,
    LCID lcid,
    DISPID *rgDispId);
HRESULT STDMETHODCALLTYPE Twapi_ComServer_Invoke(
    IDispatch *this,
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD flags,
    DISPPARAMS *dispparamsP,
    VARIANT *resultvarP,
    EXCEPINFO *excepP,
    UINT *argErrP);


/* Vtbl for Twapi_ComServer */
static struct IDispatchVtbl Twapi_ComServer_Vtbl = {
    Twapi_ComServer_QueryInterface,
    Twapi_ComServer_AddRef,
    Twapi_ComServer_Release,
    Twapi_ComServer_GetTypeInfoCount,
    Twapi_ComServer_GetTypeInfo,
    Twapi_ComServer_GetIDsOfNames,
    Twapi_ComServer_Invoke
};

/*
 * TBD - does this (related methods) need to be made thread safe?
 */
typedef struct Twapi_ComServer {
    interface IDispatch idispP; /* Must be first field */
    IID iid;                    /* IID for this interface. TBD - needed ? We only implement IDispatch, not the vtable for any other IID */
    int refc;                   /* Ref count */
    TwapiInterpContext *ticP;   /* Interpreter and related context */
    Tcl_Obj *memids;            /* List mapping member names to integer ids */
    Tcl_Obj *cmd;               /* Stores the callback command prefix */
} Twapi_ComServer;


static Tcl_Obj *TwapiComServerMemIdToName(Twapi_ComServer *me, DISPID dispid)
{
    Tcl_Obj **objs;
    int nobjs;
    int i;

    if (ObjGetElements(NULL, me->memids, &nobjs, &objs) != TCL_OK ||
        (nobjs & 1) != 0)
        return NULL;            /* Should not happen */

    for (i = 0; i < nobjs-1; i += 2) {
        int memid;
        if (ObjToInt(NULL, objs[i], &memid) != TCL_OK)
            return NULL;        /* Should not happen */
        if (memid == dispid)
            return objs[i+1];
    }

    return NULL;
}

static HRESULT STDMETHODCALLTYPE Twapi_ComServer_QueryInterface(
    IDispatch *this,
    REFIID riid,
    void **ifcPP)
{
    /* TBD - Should not check against this->iid ? Because we only implement IUNknown and IDispatch, not the vtable for the iid (e.g. in case of a sink) */

    if (!IsEqualIID(riid, &((Twapi_ComServer *)this)->iid) &&
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


static ULONG STDMETHODCALLTYPE Twapi_ComServer_AddRef(IDispatch *this)
{
    ((Twapi_ComServer *)this)->refc += 1;
    return ((Twapi_ComServer *)this)->refc;
}

static ULONG STDMETHODCALLTYPE Twapi_ComServer_Release(IDispatch *this)
{
    Twapi_ComServer *me = (Twapi_ComServer *) this;

    me->refc -= 1;
    if (((Twapi_ComServer *)this)->refc == 0) {
        if (me->memids)
            ObjDecrRefs(me->memids);
        if (me->cmd)
            ObjDecrRefs(me->cmd);
        if (me->ticP)
            TwapiInterpContextUnref(me->ticP, 1);
        TwapiFree(this);
        return 0;
    } else
        return ((Twapi_ComServer *)this)->refc;
}

static HRESULT STDMETHODCALLTYPE Twapi_ComServer_GetTypeInfoCount
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


static HRESULT STDMETHODCALLTYPE Twapi_ComServer_GetTypeInfo(
    IDispatch *this,
    UINT tinfo,
    LCID lcid,
    ITypeInfo **tiPP)
{
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Twapi_ComServer_GetIDsOfNames(
    IDispatch *this,
    REFIID   riid,
    LPOLESTR *namesP,
    UINT namesc,
    LCID lcid,
    DISPID *rgDispId)
{
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE Twapi_ComServer_Invoke(
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
    Twapi_ComServer *me = (Twapi_ComServer *) this;
    int     i;
    HRESULT hr;
    Tcl_Obj **cmdobjv;
    Tcl_Obj **cmdprefixv;
    int     cmdobjc;
    Tcl_InterpState savedState;
    Tcl_Interp *interp;
    Tcl_Obj *memberNameObj;

    if (me == NULL || me->ticP == NULL || me->ticP->interp == NULL)
        return E_POINTER;

    if (me->ticP->thread != Tcl_GetCurrentThread())
        Tcl_Panic("Twapi_ComServer_Invoke called from non-interpreter thread");

    interp = me->ticP->interp;
    if (Tcl_InterpDeleted(interp))
        return E_POINTER;

    if (ObjGetElements(NULL, me->cmd, &cmdobjc, &cmdprefixv) != TCL_OK) {
        /* Internal error - should not happen. Should we log background error?*/
        return E_FAIL;
    }

    if (flags == DISPATCH_PROPERTYPUTREF) {
        /* TBD - better error message */
        return E_FAIL;
    }

    memberNameObj = TwapiComServerMemIdToName(me, dispid);
    if (memberNameObj == NULL) {
        /* Should not really happen. Log internal error ? */
        return E_FAIL;
    }

    /* Note we will tack on member name plus dispparms */
    i = cmdobjc + 1;
    if (dispparamsP)
        i += dispparamsP->cArgs;
    cmdobjv = MemLifoPushFrame(me->ticP->memlifoP, i * sizeof(*cmdobjv), NULL);
    
    for (i = 0; i < cmdobjc; ++i) {
        cmdobjv[i] = cmdprefixv[i];
        ObjIncrRefs(cmdobjv[i]);
    }

    ObjIncrRefs(memberNameObj);
    cmdobjv[cmdobjc] = memberNameObj;
    cmdobjc += 1;

    /* Add the passed parameters */
    if (dispparamsP) {
        /* Note parameters are in reverse order */
        for (i = dispparamsP->cArgs - 1; i >= 0 ; --i) {
            /* Verify that we can handle the parameter types */
            if (dispparamsP->rgvarg[i].vt & VT_BYREF) {
                /* TBD - fill in error, free objects? */
                return E_FAIL;
            }
            cmdobjv[cmdobjc] = ObjFromVARIANT(&dispparamsP->rgvarg[i], 0);
            ObjIncrRefs(cmdobjv[cmdobjc]);
            ++cmdobjc;
        }
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
    savedState = Tcl_SaveInterpState(interp, TCL_OK);
    Tcl_ResetResult (interp);
    hr = Tcl_EvalObjv(interp, cmdobjc, cmdobjv, TCL_EVAL_GLOBAL);
    /* If hr is not TCL_OK, it is a HRESULT error code */
    if (hr != TCL_OK) {
        Tcl_BackgroundError(interp);
        if (excepP) {
            TwapiZeroMemory(excepP, sizeof(*excepP));
            excepP->scode = hr; // TBD - what error to give ? Also set error string
        }
    } else {
        /* TBD - check if interp deleted ? */

        /* TBD - appropriately init retvarP from ObjGetResult keeping
         * in mind that the retvarP by be BYREF as well.
         */
        if (retvarP) {
            VARTYPE ret_vt;
            Tcl_Obj *retObj = ObjGetResult(interp);
            
            VariantInit(retvarP); /* TBD - should be VariantClear ? */
            ret_vt = ObjTypeToVT(retObj);

            if (ObjToVARIANT(interp, retObj, retvarP, ret_vt) != TCL_OK) {
                hr = E_FAIL;
                goto restore_and_return;
            }
            if (retvarP->vt == VT_DISPATCH || retvarP->vt == VT_UNKNOWN) {
                /* When handing out interfaces, must increment their refs */
                if (retvarP->punkVal != NULL)
                    retvarP->punkVal->lpVtbl->AddRef(retvarP->punkVal);
            }
        }
        hr = S_OK;
    }

restore_and_return:
    Tcl_RestoreInterpState(interp, savedState);

    for (i = 0; i < cmdobjc; ++i) {
        ObjDecrRefs(cmdobjv[i]);
    }

    MemLifoPopFrame(me->ticP->memlifoP);

    /* Undo the AddRef we did before */
    this->lpVtbl->Release(this);
    /* this/me may be invalid at this point! Make sure we don't access them */

    return hr;
}

/*
 * Called from a script create an automation object.
 * Returns the IDispatch interface.
 */
int Twapi_ComServerObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Twapi_ComServer *comserverP;
    IID iid;
    HRESULT hr;
    Tcl_Obj **memidObjs;
    int i, nmemids;

    TWAPI_ASSERT(ticP->interp == interp);

    if (objc != 4) {
        Tcl_WrongNumArgs(interp, 1, objv, "IID MEMIDMAP CMD");
        return TCL_ERROR;
    }
    
    hr = IIDFromString(ObjToUnicode(objv[1]), &iid);
    if (FAILED(hr))
        return Twapi_AppendSystemError(interp, hr);

    if (ObjGetElements(interp, objv[2], &nmemids, &memidObjs) != TCL_OK)
        return TCL_ERROR;
    if (nmemids & 1)
        goto invalid_memids;    /* Need even number of elements */
    for (i = 0; i < nmemids-1; i += 2) {
        int memid;
        if (ObjToInt(interp, memidObjs[i], &memid) != TCL_OK)
            return TCL_ERROR;
    }

    /* Memory is freed when the object is released */
    comserverP = TwapiAlloc(sizeof(*comserverP));

    comserverP->memids = objv[2];
    ObjIncrRefs(objv[2]);


    /* Fill in the cmdargs slots from the arguments */
    comserverP->idispP.lpVtbl = &Twapi_ComServer_Vtbl;
    comserverP->iid = iid;
    TwapiInterpContextRef(ticP, 1);
    comserverP->ticP = ticP;
    comserverP->refc = 1;
    ObjIncrRefs(objv[3]);
    comserverP->cmd = objv[3];

    ObjSetResult(interp, ObjFromIUnknown(comserverP));

    return TCL_OK;

invalid_memids:
    ObjSetStaticResult(interp, "Invalid memid map");
    return TCL_ERROR;
}

