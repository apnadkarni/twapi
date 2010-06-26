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
 * successful or error. Note however, that *responseP may in fact point
 * to pcbP. However, it must be treated as independent.
 *
 * Returns 0 on success, or a Windows error code.
 * ERROR_INSUFFICIENT_BUFFER means buf[] not big enough for return value.
 */
int TwapiEnqueueCallback(
    TwapiInterpContext *ticP,   /* Caller has to ensure this does not
                                   go away until this function returns */
    TwapiPendingCallback *pcbP, /* Must be NEWLY Initialized. Ref count
                                   expected to be 0! */
    int    enqueue_method, /* Use the Tcl_Async route to queue. If 0,
                              Tcl_QueueEvent is directly called */
    int    timeout,             /* How long to wait for a result. Must be 0
                                  IF CALLING FROM A TCL THREAD ELSE
                                  DEADLOCK MAY OCCUR */
    TwapiPendingCallback **responseP /* May or may not be same as pcbP.
                                      If non-NULL, caller must call
                                      TwapiPendingCallbackUnref on it.
                                     */
    )
{
    DWORD status = ERROR_SUCCESS;

    /*
     * WE MAY NOT BE IN A TCL THREAD. DO NOT CALL ANY INTERP FUNCTIONS
     */

    if (responseP)
        *responseP = NULL;

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
    
    if (enqueue_method == TWAPI_ENQUEUE_ASYNC) {
        /* Queue via the older Tcl_Async mechanism. */
        EnterCriticalSection(&ticP->pending_cs);

        if (ticP->pending_suspended) {
            LeaveCriticalSection(&ticP->pending_cs);
            if (pcbP->completion_event)
                CloseHandle(pcbP->completion_event);
            TwapiPendingCallbackDelete(pcbP);
            return ERROR_RESOURCE_NOT_PRESENT; /* For lack of anything better */
        }

        /* Place on the pending queue. The Ref ensures it does not get
         * deallocated while on the queue. The corresponding Unref will 
         * be done by the receiver. ALWAYS. Do NOT add a Unref here 
         *
         * In addition, if we are not done with the pcbP after queueing
         * as we need to await for a response, we have to add another Ref
         * to make sure it does not go away. In that case we Ref by 2.
         * The corresponding Unref will happen below after we get the response
         * or time out.
         */
        TwapiPendingCallbackRef(pcbP, (timeout ? 2 : 1));
        ZLIST_APPEND(&ticP->pending, pcbP); /* Enqueue */

        /* Also make sure the ticP itself does not go away */
        pcbP->ticP = ticP;
        TwapiInterpContextRef(ticP, 1);

        /* To avoid races, the AsyncMark should also happen in the crit sec */
        Tcl_AsyncMark(ticP->async_handler);
    
        LeaveCriticalSection(&ticP->pending_cs);
    } else {
        /* Queue directly to the Tcl event loop for the thread */
        /* Note the CallbackEvent gets freed by the Tcl code and hence
           must be allocated using Tcl_Alloc, not malloc or TwapiAlloc */
        TwapiTclEvent *tteP = (TwapiTclEvent *) Tcl_Alloc(sizeof(*tteP));
        tteP->event.proc = Twapi_TclEventProc;
        tteP->pending_callback = pcbP;
        /* For similar reasons to above, bump ref counts */
        TwapiPendingCallbackRef(pcbP, (timeout ? 2 : 1));
        pcbP->ticP = ticP;
        TwapiInterpContextRef(ticP, 1);

        Tcl_ThreadQueueEvent(ticP->thread, (Tcl_Event *) tteP, TCL_QUEUE_TAIL);
        
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
        status = pcbP->status;
        if (responseP)
                *responseP = pcbP; /* Caller owns */
        else 
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

    /* We only handle window and file-type events here. TBD - is this right? */
    if (!(flags & (TCL_WINDOW_EVENTS|TCL_FILE_EVENTS))) return 0;

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
        pcbP->status = ERROR_DS_NO_SUCH_OBJECT; /* Best match we can find */
    } else {
        if (pcbP->callback(pcbP) != TCL_OK) {
            /* TBD - log internal error */
        }
        /* Note even for errors, pcbP response is set */
    }

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

    pcbP->callback = callback;
    pcbP->nrefs = 0;
    ZLINK_INIT(pcbP);
    pcbP->status = ERROR_SUCCESS;
    pcbP->completion_event = NULL;
    pcbP->response.type = TRT_EMPTY;
    return pcbP;
}

void TwapiPendingCallbackDelete(TwapiPendingCallback *pcbP)
{
    if (pcbP) {
        if (pcbP->completion_event)
            CloseHandle(pcbP->completion_event);
    }
    TwapiClearResult(&pcbP->response);
}

void TwapiPendingCallbackUnref(TwapiPendingCallback *pcbP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized pcbP */
    if (InterlockedExchangeAdd(&pcbP->nrefs, -decr) <= decr)
        TwapiPendingCallbackDelete(pcbP);
}


