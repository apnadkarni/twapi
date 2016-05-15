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
 * in the interpreter. cbP must not be accessed on return, either
 * successful or error. Note however, that *responseP may in fact point
 * to cbP. However, it must be treated as independent.
 *
 * Returns 0 on success, or a Windows error code.
 * ERROR_INSUFFICIENT_BUFFER means buf[] not big enough for return value.
 */
TWAPI_EXTERN int TwapiEnqueueCallback(
    TwapiInterpContext *ticP,   /* Caller has to ensure this does not
                                   go away until this function returns */
    TwapiCallback *cbP, /* Must be NEWLY Initialized. Ref count
                                   expected to be 0! */
    int    enqueue_method, /* If TCL_ENQUEUE_ASYNC, use the Tcl_Async route to
                              queue. Else Tcl_QueueEvent is directly called */
    int    timeout,             /* How long to wait for a result. Must be 0
                                  IF CALLING FROM A TCL THREAD ELSE
                                  DEADLOCK MAY OCCUR */
    TwapiCallback **responseP /* May or may not be same as cbP.
                                 If non-NULL, caller must call
                                 TwapiCallbackUnref on it.
                                 TBD - what if it is a derived type ? */
    )
{
    DWORD winerr = ERROR_SUCCESS;

    /*
     * WE MAY NOT BE IN A TCL THREAD. DO NOT CALL ANY INTERP FUNCTIONS
     */

    if (responseP)
        *responseP = NULL;

    if (timeout) {
        /* We have to wait for a response */

        /* TBD - events are probably not the most efficient synch mechanism */
        cbP->completion_event = CreateEvent(NULL,
                                             FALSE, // Auto-reset
                                             FALSE, // Initially nonsignaled
                                             NULL);
        if (cbP->completion_event == NULL) {
            winerr = GetLastError();
            /* TBD - what if some callback resources have to be freed ? */
            TwapiCallbackDelete(cbP);
            return winerr;
        }
            
    }
    
    if (enqueue_method == TWAPI_ENQUEUE_ASYNC) {
        /* No longer support this method - deprecated in Tcl */
        return ERROR_NOT_SUPPORTED;
    } else {
        /* Queue directly to the Tcl event loop for the thread */
        /* Note the CallbackEvent gets freed by the Tcl code and hence
           must be allocated using ckalloc_Alloc only */
        TwapiTclEvent *tteP = (TwapiTclEvent *) ckalloc(sizeof(*tteP));
        tteP->event.proc = Twapi_TclEventProc;
        tteP->pending_callback = cbP;

        /* Place on the pending queue. The Ref ensures it does not get
         * deallocated while on the queue. The corresponding Unref will 
         * be done by the receiver. ALWAYS. Do NOT add a Unref here 
         *
         * In addition, if we are not done with the cbP after queueing
         * as we need to await for a response, we have to add another Ref
         * to make sure it does not go away. In that case we Ref by 2.
         * The corresponding Unref will happen below after we get the response
         * or time out.
         */
        TwapiCallbackRef(cbP, (timeout ? 2 : 1));
        cbP->ticP = ticP;
        TwapiInterpContextRef(ticP, 1);

        TwapiEnqueueTclEvent(ticP, &tteP->event);
    }

    if (timeout == 0)
        return ERROR_SUCCESS;   /* No need to wait for result */

    /* Need to wait for the result */

    winerr = WaitForSingleObject(cbP->completion_event, timeout);
    if (WAIT_OBJECT_0 != winerr) {
        /* If winerr is WAIT_FAILED, we need to get error code. Others
           are already error codes */
        if (winerr == WAIT_FAILED)
            winerr = GetLastError();
    } else {
        winerr = cbP->winerr;
        if (responseP)
                *responseP = cbP; /* Caller owns */
        else 
            TwapiCallbackUnref(cbP, 1);
    }

    return winerr;
}

/*
 * Invoked from the Tcl event loop to execute a registered callback script
 */
static int Twapi_TclEventProc(Tcl_Event *tclevP, int flags)
{
    TwapiTclEvent *tteP = (TwapiTclEvent *) tclevP;
    TwapiCallback *cbP;

    /* We only handle window and file-type events here. TBD - is this right? */
    if (!(flags & (TCL_WINDOW_EVENTS|TCL_FILE_EVENTS))) return 0;

    cbP = tteP->pending_callback;

    /*
     * The interpreter may have been deleted, either logically or physically.
     * The callbacks can can check for this without locking because both 
     * those can only
     * happen in the interp's Tcl thread which is where this function runs.
     *
     * cbP and cbP->ticP themselves are protected through their ref counts
     *
     * Note we expect the actual callback to check for interp deletion
     * because we do not know here what the appropriate response should
     * be in such a case.
     */

    if (cbP->callback(cbP) != TCL_OK) {
        cbP->winerr = ERROR_BAD_ARGUMENTS;
        TwapiClearResult(&cbP->response);
    }

    if (cbP->completion_event)
        SetEvent(cbP->completion_event);

    /* Unhook the ticP from cbP */
    TwapiInterpContextUnref(cbP->ticP, 1);
    cbP->ticP = NULL;
    /* This Unref matches the Ref  from the enqueue of pending callback
     * when it was placed on the pending list and subsequently passed through
     * the Tcl event queue via Twapi_TclAsyncProc.
     */
    TwapiCallbackUnref(cbP, 1);

    /* Note tteP itself gets deleted by Tcl */

    return 1; /* So Tcl removes the event from the queue */
}


/* This routine is called the notification thread. Which may or may not
   be a Tcl interpreter thread */
TWAPI_EXTERN TwapiCallback *TwapiCallbackNew(
    TwapiInterpContext *ticP,   /* May be NULL if not a interp thread */
    TwapiCallbackFn *callback,  /* Callback function */
    int sz                   /* Including TwapiCallback header */
)
{
    TwapiCallback *cbP;

    if (sz < sizeof(TwapiCallback)) {
        if (ticP && ticP->interp)
            TwapiReturnErrorEx(ticP->interp, TWAPI_BUG, Tcl_ObjPrintf("Requested Callback size too small (%d).", sz));
        return NULL;
    }
        
    cbP = (TwapiCallback *) TwapiAlloc(sz);

    cbP->callback = callback;
    cbP->nrefs = 0;
    ZLINK_INIT(cbP);
    cbP->winerr = ERROR_SUCCESS;
    cbP->completion_event = NULL;
    cbP->response.type = TRT_EMPTY;
    cbP->clientdata = 0;
    return cbP;
}

TWAPI_EXTERN void TwapiCallbackDelete(TwapiCallback *cbP)
{
    if (cbP) {
        if (cbP->completion_event)
            CloseHandle(cbP->completion_event);
    }
    TwapiClearResult(&cbP->response);
}

TWAPI_EXTERN void TwapiCallbackUnref(TwapiCallback *cbP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized cbP */
    if (InterlockedExchangeAdd(&cbP->nrefs, -decr) <= decr)
        TwapiCallbackDelete(cbP);
}


