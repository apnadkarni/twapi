/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Contains the interface to Tcl async processing */

#include "twapi.h"

static int Twapi_TclEventProc(Tcl_Event *tclevP, int flags);


/* This routine is called from a notification thread. Which may or may not
 * be a Tcl interpreter thread. It arranges for a callback to be invoked
 * in the interpreter. pcbP must not be accessed on return, either
 * successful or error.
 *
 * Returns 0 on success, or a Windows error code.
 * ERROR_INSUFFICIENT_BUFFER means buf[] not big enough for return value.
 */
int TwapiEnqueueCallback(
    TwapiInterpContext *ticP,   /* Caller has to ensure this does not
                                   go away until this function returns */
    TwapiPendingCallback *pcbP, /* Must be NEWLY Initialized. Ref count
                                   expected to be 0! */
    int    timeout,             /* How long to wait for a result. Must be 0
                                  IF CALLING FROM A TCL THREAD ELSE
                                  DEADLOCK MAY OCCUR */
    WCHAR **resultPP            /* TwapiAlloc'ed, must be TwapiFree'ed
                                   by caller */
    )
{
    DWORD status = ERROR_SUCCESS;

    /*
     * WE MAY NOT BE IN A TCL THREAD. DO NOT CALL ANY INTERP FUNCTIONS
     */

    if (timeout) {
        /* We have to wait for a response */

        /* TBD - events are probably not the most efficient synch mechanism */
        pcbP->completion_event = CreateEvent(NULL,
                                             FALSE, // Auto-reset
                                             FALSE, // Initially nonsignaled
                                             NULL);
        if (pcbP->completion_event == NULL) {
            status = GetLastError();
            TwapiPendingCallbackDelete(pcbP);
            return status;
        }
            
    }
    
    status = TwapiInterpContextEnqueueCallback(ticP, pcbP, (timeout != 0));
    if (status != ERROR_SUCCESS) {
        if (pcbP->completion_event)
            CloseHandle(pcbP->completion_event);
        TwapiPendingCallbackDelete(pcbP);
        return ERROR_RESOURCE_NOT_PRESENT; /* For lack of anything better */
    }

    if (timeout == 0)
        return ERROR_SUCCESS;   /* No need to wait for result */

    /* Need to wait for the result */

    status = WaitForSingleObject(pcbP->completion_event, timeout);
    if (WAIT_OBJECT_0 != status) {
        /* If status is WAIT_FAILED, we need to get error code. Others
           are already error codes */
        if (status == WAIT_FAILED)
            status = GetLastError();
    } else {
        if (resultPP)
                *resultPP = pcbP->resultP; /* Caller owns resultP */
            else if (pcbP->resultP)
                TwapiFree(pcbP->resultP); /* Else we free it */

        pcbP->resultP = NULL;
        status = pcbP->status;
        TwapiPendingCallbackUnref(pcbP, 1);
    }

    return status;
}

    
/*
 * Called from Tcl loop to check for events. The function checks if
 * any events are pending and queues on the Tcl event queue
 */
int Twapi_TclAsyncProc(TwapiInterpContext *ticP,
                       Tcl_Interp *arbitrary_interp, /* May be NULL or other
                                                        than ticP->interp */
                       int code)
{
    /*
     * When called by Tcl, passed interp is not necessarily the interp we
     * want to target It may even be NULL. Also so as to not interfere
     * with any evals in progress, we always return 'code'. See
     * Tcl_AsyncMark manpage for details.
     */
    
    EnterCriticalSection(&ticP->pending_cs);

    /*
     * Loop and queue up all pending events. Note we do this even if
     * pending of events is suspended. Twapi_TclEventProc will deal
     * with that case when unqueueing the Tcl event.
     */
    while (ZLIST_COUNT(&ticP->pending)) {
        TwapiPendingCallback *pcbP;
        TwapiTclEvent *tteP;

        pcbP = ZLIST_HEAD(&ticP->pending);
        ZLIST_REMOVE(&ticP->pending, pcbP);

        /* Unlock before entering Tcl code */
        LeaveCriticalSection(&ticP->pending_cs);

        /*
         * The following two Ref/Unref cancel each other so
         * we do not do them.
         TwapiPendingCallbackUnref(pcbP,1) - we have removed from pending list
         TwapiPendingCallbackRef(pcbP,1) -  we are putting on Tcl event queue
         *
         * Also, ticP is still pointed to by pcbP so we do not Unref that
         */

        /* Note the CallbackEvent gets freed by the Tcl code and hence
           must be allocated using Tcl_Alloc, not malloc or TwapiAlloc */
        tteP = (TwapiTclEvent *) Tcl_Alloc(sizeof(*tteP));
        tteP->event.proc = Twapi_TclEventProc;
        tteP->pending_callback = pcbP;
        Tcl_QueueEvent((Tcl_Event *) tteP, TCL_QUEUE_TAIL);

        /* Lock again before checking if empty */
        EnterCriticalSection(&ticP->pending_cs);
    }

    LeaveCriticalSection(&ticP->pending_cs);
    return code;
}


/*
 * Invoked from the Tcl event loop to execute a registered callback script
 */
static int Twapi_TclEventProc(Tcl_Event *tclevP, int flags)
{
    TwapiTclEvent *tteP = (TwapiTclEvent *) tclevP;
    TwapiPendingCallback *pcbP;
    DWORD status;                 /* Win32 error */
    Tcl_Obj *objP;
    int len;
    WCHAR *wP;

    /* We only handle file-type events here. TBD - is this right? */
    if (!(flags & TCL_FILE_EVENTS)) return 0;

    pcbP = tteP->pending_callback;

    // TBD - review this comment
    // Check if the interpreter has been deleted
    // We know the interp structure is still valid because of the
    // Tcl_Preserve(interp) in the PendingCallback constructor. However
    // it might have been logically deleted

    /*
     * The interpreter may have been deleted, either logically or physically.
     * We can check for this without locking because both those can only
     * happen in the interp's Tcl thread which is where this function runs.
     *
     * pcbP and pcbP->ticP themselves are protected through their ref counts
     */
    if (pcbP->ticP->interp == NULL ||
        Tcl_InterpDeleted(pcbP->ticP->interp)) {
        status = ERROR_DS_NO_SUCH_OBJECT; /* Best match we can find */
    } else {
        status = pcbP->callback(pcbP, &objP);
        if (objP) {        
            /* In case callback has no reference to it, we need to delete
             * object when done. On the other hand, if there is some ref
             * to it elsewhere, the Decr below will delete it if we do not
             * Incr it. To take care of both cases, we need a Incr/Decr pair
             */
            Tcl_IncrRefCount(objP);
            if (status == ERROR_SUCCESS) {
                /* Do we need TclSave/RestoreResult ? */
                status = Tcl_EvalObjEx(pcbP->ticP->interp, objP,
                                       TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);
                status == TCL_OK ? ERROR_SUCCESS : E_FAIL;
                Tcl_DecrRefCount(objP); /* GONE */
                /* Note interp may be (logically) deleted, but we
                   should still be able to get the result */
                objP = Tcl_GetObjResult(pcbP->ticP->interp);
                Tcl_IncrRefCount(objP);
            }

            /* Get and pass on result if necessary */
            if (pcbP->completion_event) {
                wP = Tcl_GetUnicodeFromObj(objP, &len);
                pcbP->resultP = TwapiAllocString(wP, len);
            }
            Tcl_DecrRefCount(objP);

            Tcl_ResetResult(pcbP->ticP->interp);/* Don't leave crud from eval */
        }
    }

    pcbP->status = status;
    if (pcbP->completion_event)
        SetEvent(pcbP->completion_event);

    /* Unhook the ticP from pcbP */
    TwapiInterpContextUnref(pcbP->ticP, 1);
    pcbP->ticP = NULL;
    /* This Unref matches the Ref  from the enqueue of pending callback
     * when it was placed on the pending list and subsequently passed through
     * the Tcl event queue via Twapi_TclAsyncProc.
     */
    TwapiPendingCallbackUnref(pcbP, 1);

    /* Note tteP itself gets deleted by Tcl */

    return 1; /* So Tcl removes the event from the queue */
}


/* This routine is called the notification thread. Which may or may not
   be a Tcl interpreter thread */
TwapiPendingCallback *TwapiPendingCallbackNew(
    TwapiInterpContext *ticP,   /* May be NULL if not a interp thread */
    TwapiCallbackFn *callback,  /* Callback function */
    size_t sz                   /* Including TwapiPendingCallback header */
)
{
    TwapiPendingCallback *pcbP;

    if (sz < sizeof(TwapiPendingCallback)) {
        if (ticP && ticP->interp)
            TwapiReturnTwapiError(ticP->interp, NULL, TWAPI_BUG);
        return NULL;
    }
        
    pcbP = (TwapiPendingCallback *) TwapiAlloc(sz);
    if (pcbP == NULL) {
        if (ticP && ticP->interp)
            Twapi_AppendSystemError(ticP->interp, E_OUTOFMEMORY);
        return NULL;
    }

    pcbP->callback = callback;
    pcbP->nrefs = 0;
    pcbP->completion_event = NULL;
    ZLINK_INIT(pcbP);
    return pcbP;
}

void TwapiPendingCallbackDelete(TwapiPendingCallback *pcbP)
{
    if (pcbP) {
        if (pcbP->completion_event)
            CloseHandle(pcbP->completion_event);
    }
}

void TwapiPendingCallbackUnref(TwapiPendingCallback *pcbP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized pcbP */
    if (InterlockedExchangeAdd(&pcbP->nrefs, -decr) <= decr)
        TwapiPendingCallbackDelete(pcbP);
}

