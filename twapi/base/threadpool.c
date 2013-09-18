/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

/* Note no locking necessary as only accessed from interp thread */
#define TwapiThreadPoolRegistrationRef(p_, incr_)    \
    do {(p_)->nrefs += (incr_);} while (0)

static void TwapiThreadPoolRegistrationUnref(TwapiThreadPoolRegistration *tprP,
                                             int unrefs)
{
    /* Note no locking necessary as only accessed from interp thread */
    tprP->nrefs -= unrefs;
    if (tprP->nrefs <= 0)
        TwapiFree(tprP);
}

static int TwapiThreadPoolRegistrationCallback(TwapiCallback *cbP)
{
    TwapiThreadPoolRegistration *tprP;
    TwapiInterpContext *ticP = cbP->ticP;
    TwapiId id;
    
    TWAPI_ASSERT(ticP);

    TwapiClearResult(&cbP->response); /* Not really necessary but for consistency */
    if (ticP->interp == NULL ||
        Tcl_InterpDeleted(ticP->interp)) {
        return TCL_ERROR;
    }

    id = (TwapiId) cbP->clientdata;
    ZLIST_LOCATE(tprP, &ticP->threadpool_registrations, id, id);
    if (tprP == NULL) {
        return TCL_OK;                 /* Stale is not an error */
    }

    /* Signal the event. Result is ignored */
    tprP->signal_handler(ticP, id, tprP->handle, (DWORD) cbP->clientdata2);
    return TCL_OK;
}


/* Called from the thread pool when a handle is signalled */
static VOID CALLBACK TwapiThreadPoolRegistrationProc(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
    )
{
    TwapiThreadPoolRegistration *tprP =
        (TwapiThreadPoolRegistration *) lpParameter;
    TwapiCallback *cbP;

    /*
     * Note - tprP is guaranteed to not have disappeared as it is ref counted
     * and not unref'ed until the handle is unregistered from the thread pool
     */
    cbP = TwapiCallbackNew(tprP->ticP,
                           TwapiThreadPoolRegistrationCallback,
                           sizeof(*cbP));

    /*
     * Even though tprP is valid at this point, it may not be valid
     * when the the callback is invoked. So we do not pass tprP directly,
     * but instead pass its id so it will looked up in the call back.
     */
    cbP->clientdata = (DWORD_PTR) tprP->id;
    cbP->clientdata2 = (DWORD_PTR) TimerOrWaitFired;
    cbP->winerr = ERROR_SUCCESS;
    TwapiEnqueueCallback(tprP->ticP, cbP,
                         TWAPI_ENQUEUE_DIRECT,
                         0, /* No response wanted */
                         NULL);
    /* TBD - on error, do we send an error notification ? */
}


void TwapiThreadPoolRegistrationShutdown(TwapiThreadPoolRegistration *tprP)
{
    int unrefs = 0;
    if (tprP->tp_handle != INVALID_HANDLE_VALUE) {
        if (! UnregisterWaitEx(tprP->tp_handle,
                               INVALID_HANDLE_VALUE /* Wait for callbacks to finish */
                )) {
            /*
             * TBD - how does one handle this ? Is it unregistered, was
             * never registered? or what ? At least log it.
             */
        }
        ++unrefs;           /* Since no longer referenced from thread pool */
    }

    /*
     * NOTE: Depending on the type of handle, NULL may or may not be
     * a valid handle. We always use INVALID_HANDLE_VALUE as a invalid
     * indicator.
     */
    if (tprP->unregistration_handler && tprP->handle != INVALID_HANDLE_VALUE)
        tprP->unregistration_handler(tprP->ticP, tprP->id, tprP->handle);
    tprP->handle = INVALID_HANDLE_VALUE;

    tprP->tp_handle = INVALID_HANDLE_VALUE;

    if (tprP->ticP) {
        ZLIST_REMOVE(&tprP->ticP->threadpool_registrations, tprP);
        TwapiInterpContextUnref(tprP->ticP, 1); /*  May be gone! */
        tprP->ticP = NULL;
        ++unrefs;                         /* Not referenced from list */
    }

    if (unrefs)
        TwapiThreadPoolRegistrationUnref(tprP, unrefs); /* May be gone! */
}


TCL_RESULT TwapiThreadPoolRegister(
    TwapiInterpContext *ticP,
    HANDLE h,
    ULONG wait_ms,
    DWORD  flags,
    void (*signal_handler)(TwapiInterpContext *ticP, TwapiId, HANDLE, DWORD),
    void (*unregistration_handler)(TwapiInterpContext *ticP, TwapiId, HANDLE)
    )
{
    TwapiThreadPoolRegistration *tprP = TwapiAlloc(sizeof(*tprP));

    tprP->handle = h;
    tprP->id = TWAPI_NEWID(ticP);
    tprP->signal_handler = signal_handler;
    tprP->unregistration_handler = unregistration_handler;

    /* Only certain flags are obeyed. */
    flags &= WT_EXECUTEONLYONCE;

    flags |= WT_EXECUTEDEFAULT;

    /*
     * Note once registered with thread pool call back might run even
     * before the registration call returns so set everything up
     * before the call
     */
    ZLIST_PREPEND(&ticP->threadpool_registrations, tprP);
    tprP->ticP = ticP;
    TwapiInterpContextRef(ticP, 1);
    /* One ref for list linkage, one for handing off to thread pool */
    TwapiThreadPoolRegistrationRef(tprP, 2);
    if (RegisterWaitForSingleObject(&tprP->tp_handle,
                                    h,
                                    TwapiThreadPoolRegistrationProc,
                                    tprP,
                                    wait_ms,
                                    flags)) {
        
        return ObjSetResult(ticP->interp, ObjFromTwapiId(tprP->id));
    } else {
        tprP->tp_handle = INVALID_HANDLE_VALUE; /* Just to be sure */
        /* Back out the ref for thread pool since it failed */
        TwapiThreadPoolRegistrationUnref(tprP, 1);
        TwapiThreadPoolRegistrationShutdown(tprP);

        return TwapiReturnError(ticP->interp, TWAPI_REGISTER_WAIT_FAILED);
    }
}
    
                                       
void TwapiThreadPoolUnregister(
    TwapiInterpContext *ticP,
    TwapiId id
    )
{
    TwapiThreadPoolRegistration *tprP;
    
    ZLIST_LOCATE(tprP, &ticP->threadpool_registrations, id, id);
    if (tprP == NULL)
        return;                 /* Stale? */

    TWAPI_ASSERT(ticP == tprP->ticP);

    TwapiThreadPoolRegistrationShutdown(tprP);
}


/* The callback that invokes the user level script for async handle waits */
void TwapiCallRegisteredWaitScript(TwapiInterpContext *ticP, TwapiId id, HANDLE h, DWORD timeout)
{
    Tcl_Obj *objs[4];
    int i;

    objs[0] = ObjFromString(TWAPI_TCL_NAMESPACE "::_wait_handler");
    objs[1] = ObjFromTwapiId(id);
    objs[2] = ObjFromHANDLE(h);
    if (timeout) 
        objs[3] = STRING_LITERAL_OBJ("timeout");
    else
        objs[3] = STRING_LITERAL_OBJ("signalled");

    for (i = 0; i < ARRAYSIZE(objs); ++i) {
        Tcl_IncrRefCount(objs[i]);
    }
    Tcl_EvalObjv(ticP->interp, ARRAYSIZE(objs), objs, 
                 TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);

    for (i = 0; i < ARRAYSIZE(objs); ++i) {
        ObjDecrRefs(objs[i]);
    }
}

