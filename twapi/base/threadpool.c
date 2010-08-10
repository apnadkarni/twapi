/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */


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
    HANDLE h;
    
    TWAPI_ASSERT(ticP);

    TwapiClearResult(&cbP->response); /* Not really necessary but for consistency */
    if (ticP->interp == NULL ||
        Tcl_InterpDeleted(ticP->interp)) {
        return TCL_ERROR;
    }

    h = (HANDLE)cbP->clientdata;
    ZLIST_LOCATE(tprP, &ticP->threadpool_registrations, handle, h);
    if (tprP == NULL) {
        return TCL_OK;                 /* Stale but ok */
    }

    /* Signal the event. Result is ignored */
    tprP->signal_handler(ticP, h, (DWORD) cbP->clientdata2);
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

    cbP = TwapiCallbackNew(tprP->ticP,
                           TwapiThreadPoolRegistrationCallback,
                           sizeof(*cbP));

    /* Note we do not directly pass tprP. If we did would need to Ref it */
    cbP->clientdata = (DWORD_PTR) tprP->handle;
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
        tprP->unregistration_handler(tprP->ticP, tprP->handle);
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


WIN32_ERROR TwapiThreadPoolRegister(
    TwapiInterpContext *ticP,
    HANDLE h,
    ULONG wait_ms,
    DWORD  flags,
    void (*signal_handler)(TwapiInterpContext *ticP, HANDLE, DWORD),
    void (*unregistration_handler)(TwapiInterpContext *ticP, HANDLE)
    )
{
    TwapiThreadPoolRegistration *tprP = TwapiAlloc(sizeof(*tprP));

    tprP->handle = h;
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
        return ERROR_SUCCESS;
    } else {
        tprP->tp_handle = INVALID_HANDLE_VALUE; /* Just to be sure */
        /* Back out the ref for thread pool since it failed */
        TwapiThreadPoolRegistrationUnref(tprP, 1);

        TwapiThreadPoolRegistrationShutdown(tprP);
        return TWAPI_ERROR_TO_WIN32(TWAPI_REGISTER_WAIT_FAILED);
    }
}
    
                                       
void TwapiThreadPoolUnregister(
    TwapiInterpContext *ticP,
    HANDLE h
    )
{
    TwapiThreadPoolRegistration *tprP;
    
    ZLIST_LOCATE(tprP, &ticP->threadpool_registrations, handle, h);
    if (tprP == NULL)
        return;                 /* Stale? */

    TWAPI_ASSERT(ticP == tprP->ticP);

    TwapiThreadPoolRegistrationShutdown(tprP);
}


/* The callback that invokes the user level script for async handle waits */
void TwapiCallRegisteredWaitScript(TwapiInterpContext *ticP, HANDLE h, DWORD timeout)
{
    Tcl_Obj *objs[3];
    int i;

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_wait_handler", -1);
    objs[1] = ObjFromHANDLE(h);
    if (timeout) 
        objs[2] = STRING_LITERAL_OBJ("timeout");
    else
        objs[2] = STRING_LITERAL_OBJ("signalled");

    for (i = 0; i < ARRAYSIZE(objs); ++i) {
        Tcl_IncrRefCount(objs[i]);
    }
    Tcl_EvalObjv(ticP->interp, ARRAYSIZE(objs), objs, 
                 TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);

    for (i = 0; i < ARRAYSIZE(objs); ++i) {
        Tcl_DecrRefCount(objs[i]);
    }
}

