/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

typedef struct _TwapiPipeContext {
    TwapiInterpContext *ticP;
    HANDLE  hpipe;   /* Handle to the pipe */
    HANDLE  hsync;   /* Event used for synchronous I/O, both read and write.
                        Needed because we cannot use the hevent field
                        for sync i/o as the thread pool will also be waiting
                        on it */
    struct _TwapiPipeIO {
        OVERLAPPED ovl;         /* Used in async i/o. Note the hEvent field
                                   is always set at init time */
        HANDLE hwait;       /* Handle returned by thread pool wait functions */
        HANDLE hevent;      /* Event for thread pool to wait on */
        union {
            char read_ahead[1];     /* Used to read-ahead a single byte */
            Tcl_Obj *objP;      /* Data to write out */
        } data;
        ULONG volatile state;
#define IOBUF_IDLE         0
#define IOBUF_IO_PENDING   1
#define IOBUF_IO_COMPLETED 2
#define IOBUF_IO_COMPLETED_WITH_ERROR 3
    } io[2];                   /* 0 -> read, 1 -> write */
#define READER 0
#define WRITER 1

        int    flags;
#define PIPE_F_WATCHREAD       1
#define PIPE_F_WATCHWRITE      2
#define PIPE_F_NONBLOCKING     4

    ZLINK_DECL(TwapiPipeContext);
    ULONG volatile nrefs;              /* Ref count */
    WIN32_ERROR winerr;
} TwapiPipeContext;


/*
 * Prototypes
 */
static TwapiPipeContext *TwapiPipeContextNew(void);
static void TwapiPipeContextDelete(TwapiPipeContext *ctxP);
#define TwapiPipeContextRef(p_, incr_) InterlockedExchangeAdd(&(p_)->nrefs, (incr_))
void TwapiPipeContextUnref(TwapiPipeContext *ctxP, int decr);
static int TwapiPipeShutdown(TwapiPipeContext *ctxP);
static DWORD TwapiPipeConnectClient(TwapiPipeContext *ctxP);
static DWORD TwapiQueuePipeCallback(TwapiPipeContext *ctxP, const char *s);


/*
 * Sets up overlapped reads on a pipe. The read state must be IDLE so
 * that the data buffer is not in use and no reads are pending.
 */
static WIN32_ERROR TwapiPipeEnableReadWatch(TwapiPipeContext *ctxP)
{
    DWORD winerr;

    if (ctxP->winerr != ERROR_SUCCESS ||
        ctxP->io[READER].state != IOBUF_IDLE ||
        ! (ctxP->flags & PIPE_F_WATCHREAD)) {
        return TWAPI_ERROR_TO_WIN32(TWAPI_INVALID_STATE_FOR_OP);
    }

    /*
     * We set up a single byte read so we're told when data is available.
     * We do not bother to special case when i/o completes immediately.
     * Let it also follow the async path for simplicity.
     */
    ZeroMemory(&ctxP->io[READER].ovl, sizeof(ctxP->io[READER].ovl));
    ctxP->io[READER].ovl.hEvent = ctxP->io[READER].hevent;
    if (! ReadFile(ctxP->hpipe,
                   &ctxP->io[READER].data.read_ahead,
                   1,
                   NULL,
                   &ctxP->io[READER].ovl)) {
        winerr = GetLastError();
        if (winerr != ERROR_IO_PENDING)
            return winerr;
    }

    /* Have the thread pool wait on it if not already doing so */
    if (ctxP->io[READER].hwait == INVALID_HANDLE_VALUE) {
        /* The thread pool will hold a ref to the context */
        TwapiPipeContextRef(ctxP, 1);
        if (! RegisterWaitForSingleObject(
                &ctxP->io[READER].hwait,
                ctxP->io[READER].hevent,
                TwapiPipeReadThreadPoolFn,
                ctxP,
                INFINITE,           /* No timeout */
                WT_EXECUTEDEFAULT
                )) {
            TwapiPipeContextUnref(ctxP, 1);
            ctxP->io[READER].hwait = INVALID_HANDLE_VALUE; /* Just in case... */
            /* Note the call does not set GetLastError. Make up our own */
            return TWAPI_ERROR_TO_WIN32(TWAPI_REGISTER_WAIT_FAILED);
        }

    }
    ctxP->io[READER].state = IOBUF_IO_PENDING;

    return ERROR_SUCCESS;
}

static WIN32_ERROR TwapiPipeDisableReadWatch(TwapiPipeContext *ctxP)
{
    DWORD winerr;

    /*
     * If state is PENDING we do not want to unregister the read since
     * then there will be an overlapping read outstanding with no one
     * waiting on the event. Caller is supposed to have checked for that
     */
    if (ctxP->io[READER].state == IOBUF_IO_PENDING ||
        (ctxP->flags & PIPE_F_WATCHREAD)) {
        return TWAPI_ERROR_TO_WIN32(TWAPI_INVALID_STATE_FOR_OP);
    }

    TBD

}

/* Called from Tcl event loop with a pipe event notification */
DWORD TwapiPipeCallbackFn(TwapiCallback *cbP)
{
    Tcl_Obj *objs[4];
    int nobjs;
    TwapiPipeContext *ctxP = (TwapiPipeContext *) cbP->clientdata;
    int flags;
    char *eventstr;

    /* If the interp is gone, close down the pipe */
    if (ctxP->ticP == NULL ||
        ctxP->ticP->interp == NULL ||
        Tcl_InterpDeleted(ctxP->ticP->interp)) {
        cbP->clientdata = NULL;
        TwapiPipeShutdown(ctxP);
        TwapiPipeContextUnref(ctxP);
        ctxP = NULL;            /* Ensure we do not access it */
        cbP->winerr = ERROR_INVALID_FUNCTION; // TBD
        cbP->response.type = TRT_EMPTY;
        return TCL_ERROR;
    }

    if (cbP->clientdata2 == READABLE) {
        flags = PIPE_F_WATCHREAD;
        eventstr = "read";
    } else {
        flags = PIPE_F_WATCHWRITE;
        eventstr = "write";
    }

    /*
     * Check if pipe is still being watched. We only queue a notification
     * if that is the case. Moreover, if it is not being watched, we need to
     * turn off notifications.
     */
    if (ctxP->flags & flags) {
        /* Still being watched */
        objs[0] = STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_pipe_handler");
        objs[1] = ObjFromHANDLE(ctxP->hpipe);
        if (cbP->winerr == ERROR_SUCCESS || cbP->winerr == ERROR_HANDLE_EOF) {
            objs[2] = Tcl_NewStringObj(eventstr, -1);
            nobjs = 3;
        } else {
            /* TBD - what do sockets pass to fileevent when there is an error ? */
            objs[2] = STRING_LITERAL_OBJ("error");
            objs[3] = Tcl_NewLongObj((long)cbP->winerr);
            nobjs = 4;
        }
        return TwapiEvalAndUpdateCallback(cbP, nobjs, objs, TRT_EMPTY);
    } else {
        /* Pipe not being watched. Turn off notification if not already done */
        if (cbP->clientdata2 == READABLE)
            TwapiPipeDisableReadWatch(ctxP);
        else
            TwapiPipeDisableWriteWatch(ctxP);

        cbP->winerr = ERROR_SUCCESS;
        cbP->response.type = TRT_EMPTY;
        return TCL_OK;
    }
}


static void TwapiPipeEnqueueCallback(
    TwapiPipeContext *ctxP,
    int  direction              /* READER or WRITER */
)
{

    struct _TwapiPipeIO *ioP = &ctxP->io[direction];
    TwapiCallback *cbP;
    LONG new_state;

    TWAPI_ASSERT(HasOverlappedIoCompleted(&ioP->ovl));
    TWAPI_ASSERT(direction == READER || direction == WRITER);

    new_state = (ioP->ovl.Internal == ERROR_SUCCESS) ?
        IOBUF_IO_COMPLETED : IOBUF_IO_COMPLETED_WITH_ERROR;

    /*
     * If current state is not PENDING, ignore. Note state has to be changed
     * here, and not in the Tcl thread callback since that thread may be
     * blocked waiting for state to change from pending to completed.
     */
    if (InterlockedCompareExchange(&ioP->state, new_state, IOBUF_IO_PENDING)
        != IOBUF_IO_PENDING)
        return;

    cbP = TwapiCallbackNew(ctxP->ticP, TwapiPipeCallbackFn, sizeof(*cbP));
    cbP->winerr = ioP->ovl.Internal;
    TwapiPipeContextRef(ctxP);
    cbP->clientdata = ctxP;
    cbP->clientdata2 = direction;
    TwapiEnqueueCallback(ctxP->ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);

    return;
}

/*
 * Called from thread pool when a overlapped read completes.
 * Note this has the typedef for WaitOrTimerCallback 
 */
static VOID CALLBACK TwapiPipeReadThreadPoolFn(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
)
{
    TWAPI_ASSERT(TimerOrWaitFired == FALSE);
    TwapiPipeEnqueueCallback((TwapiPipeContext*) lpParameter, READER);
}

/*
 * Called from thread pool when a overlapped write completes.
 * Note this has the typedef for WaitOrTimerCallback 
 */
static VOID CALLBACK TwapiPipeWriteThreadPoolFn(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
)
{
    TWAPI_ASSERT(TimerOrWaitFired == FALSE);
    TwapiPipeEnqueueCallback((TwapiPipeContext*) lpParameter, WRITER);
}


/*
 * Return the current number of available bytes for reading including the
 * read-ahead byte.
 */
WIN32_ERROR TwapiPipeReadCount(TwapiPipeContext *ctxP, DWORD *countP)
{
    DWORD count;

    if (ctxP->io[READER].state == IOBUF_IO_PENDING) {
        *countP = 0;               /* I/O pending so obviously no data */
        return ERROR_SUCCESS;
    }

    /*
     * Note in COMPLETED, COMPLETED_WITH_ERROR or IDLE state, the thread
     * pool will not run so we do not have to worry about consistency 
     * between the read ahead and what's in the pipe. Also, even if
     * the state is COMPLETED_WITH_ERROR, we still allow the reads.
     */
       
    if (!PeekNamedPipe(ctxP->hpipe, NULL, 0, NULL, &count, NULL))
        return GetLastError();

    if (ctxP->io[READER].state == IOBUF_IO_COMPLETED)
        ++count;              /* The read-ahead byte */

    *countP = count;

    return ERROR_SUCCESS;
}


/*
 * Read count bytes and return a Tcl_Obj. Caller must have ensured there
 * are that many bytes available else function will block waiting for data.
 */
static WIN32_ERROR TwapiPipeReadData(TwapiPipeContext *ctxP, DWORD count, Tcl_Obj **objPP)
{
    struct _TwapiPipeIO *ioP = &ctxP->io[READER];
    Tcl_Obj *objP = Tcl_NewByteArrayObj(NULL, count);
    char *p = Tcl_GetByteArrayFromObj(objP, NULL);
    DWORD winerr;

    TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING);
    TWAPI_ASSERT(count);

    if (ioP->state == IOBUF_IO_COMPLETED) {
        --count;
        *p++ = ioP->data.read_ahead[0];
        ioP->state = IOBUF_IDLE;
    }

    if (count > 0) {
        DWORD read_count;
        ZeroMemory(&ioP->ovl, sizeof(ioP->ovl));
        ioP->hEvent = ctxP->hsync; /* Event used for "synchronous" reads */
        if (! ReadFile(ctxP->hpipe, p, count, NULL, &ioP->ovl)) {
            winerr = GetLastError();
            if (winerr != ERROR_IO_PENDING) {
                Tcl_DecrRefCount(objP);
                return winerr;
            }
            /* Pending I/O. Should not happen, but no matter, fall thru */
        }
        if (!GetOverlappedResult(ctxP->hpipe, &ioP->ovl, &read_count, TRUE)) {
            winerr = GetLastError();
            Tcl_DecrRefCount(objP);
            return winerr;
        }
        TWAPI_ASSERT(read_count == count);
    }

    *objPP = objP;
    return ERROR_SUCCESS;
}


int Twapi_PipeServer(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR name;
    DWORD open_mode, pipe_mode, max_instances;
    DWORD inbuf_sz, outbuf_sz, timeout;
    SECURITY_ATTRIBUTES *secattrP;
    TwapiPipeContext *ctxP;
    DWORD winerr;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETWSTR(name), GETINT(open_mode), GETINT(pipe_mode),
                     GETINT(max_instances), GETINT(outbuf_sz),
                     GETINT(inbuf_sz), GETINT(timeout),
                     GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES), ARGEND)
        != TCL_OK)
        return TCL_ERROR;

    ctxP = TwapiPipeContextNew();

    ctxP->hpipe = CreateNamedPipeW(name, open_mode, pipe_mode, max_instances,
                                   outbuf_sz, inbuf_sz, timeout, secattrP);

    if (ctxP->hpipe != INVALID_HANDLE_VALUE) {
        /* Create event used for sync i/o. Note this is manual reset event */
        ctxP->hsync = CreateEvent(NULL, TRUE, FALSE, NULL);

        if (ctxP->hsync) {
            /* 
             * Create events to use for notification of completion. The
             * events must be auto-reset to prevent multiple callback
             * queueing on a single input notification. See MSDN docs for
             * RegisterWaitForSingleObject.  As a consequence, we must
             * make sure we never call GetOverlappedResult in blocking
             * mode.
             */
            ctxP->io[READER].hevent = CreateEvent(NULL, FALSE, FALSE, NULL);
            if (ctxP->io[READER].hevent) {
                ctxP->io[WRITER].hevent = CreateEvent(NULL, FALSE, FALSE, NULL);
                if (ctxP->io[WRITER].hevent) {
                    /* Add to list of pipes */
                    TwapiPipeContextRef(ctxP, 1);
                    ZLIST_PREPEND(&ticP->pipes, ctxP);
                    ctxP->ticP = ticP;
                    TwapiInterpContextRef(ticP, 1);
                
                    winerr = TwapiPipeConnectClient(ctxP);
                    if (winerr == ERROR_SUCCESS || winerr == ERROR_IO_PENDING) {
                        Tcl_Obj *objs[2];
                        if (winerr == ERROR_SUCCESS)
                            objs[0] = STRING_LITERAL_OBJ("connected");
                        else
                            objs[0] = STRING_LITERAL_OBJ("pending");
                        objs[1] = ObjFromHANDLE(ctxP->hpipe);
                        Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(2, objs));
                        return TCL_OK;
                    }
                } else
                    winerr = GetLastError();
            } else
                winerr = GetLastError();
        } else
            winerr = GetLastError();
    }

    TwapiPipeShutdown(ctxP);
    return Twapi_AppendSystemError(ticP->interp, winerr);
}

int Twapi_PipeClose(TwapiInterpContext *ticP, HANDLE hpipe)
{
    TwapiPipeContext *ctxP;

    TWAPI_ASSERT(ticP->thread == Tcl_GetCurrentThread());
    
    ZLIST_LOCATE(ctxP, &ticP->pipes, hpipe, hpipe);
    if (ctxP == NULL)
        return TwapiReturnTwapiError(ticP->interp, NULL, TWAPI_UNKNOWN_OBJECT);

    TWAPI_ASSERT(ticP == ctxP->ticP);

    TwapiPipeShutdown(ctxP);

    return TCL_OK;
}

int Twapi_PipeRead(TwapiInterpContext *ticP, HANDLE hpipe, DWORD count)
{
    TwapiPipeContext *ctxP;
    DWORD winerr;
    DWORD nread, avail;
    struct _TwapiPipeIO *ioP;
    Tcl_Obj *objP;

    TWAPI_ASSERT(ticP->thread == Tcl_GetCurrentThread());
    
    ZLIST_LOCATE(ctxP, &ticP->pipes, hpipe, hpipe);
    if (ctxP == NULL)
        return TwapiReturnTwapiError(ticP->interp, NULL, TWAPI_UNKNOWN_OBJECT);

    TWAPI_ASSERT(ticP == ctxP->ticP);

    ioP = &ctxP->io[READER];
    if (ctxP->winerr != ERROR_SUCCESS)
        goto error_return;

    if (count == 0)
        return TCL_OK; // TBD ? Will get treated as EOF

    if ((ctxP->winerr = TwapiPipeReadCount(ctxP, &avail)) != ERROR_SUCCESS) {
        goto error_return;
    }

    if (ctxP->flags & PIPE_F_NONBLOCKING) {
        /* Non-blocking channel */
        if (avail == 0) {
            Tcl_SetResult(interp, "EAGAIN", TCL_STATIC);
            return TCL_ERROR;   /* As expected by Tcl channel implementation */
        }
        count = avail;      /* Non-blocking case - return what we have */

        TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING); /* When avail > 0 */
    } else {
        /* Blocking channel. */

        /*
         * Loop waiting for I/O to complete. Note this combination of
         * non-blocking I/O with use of fileevent (async notifications)
         * is rare so no need for more sophisticated mechanisms. The
         * thread pool handler will set state to IOBUF_IO_COMPLETED.
         */
        while (ioP->state == IOBUF_IO_PENDING) {
            SleepEx(1, FALSE);
        }
        /* Error might have occured while we waited */
        if (ctxP->winerr != ERROR_SUCCESS)
            goto error_return;

        /* Now that the async thread pool is out of the way, we can go
           ahead and do the blocking read */
    }

    ctxP->winerr = TwapiPipeReadData(ctxP, count, &objP);

    if (ctxP->winerr == ERROR_SUCCESS) {
        Tcl_SetObjResult(ticP->interp, objP);

        /*
         * Need to fire another read event if more data is available.
         * PipeWatch may return an error but we do not report that right away.
         * It will be picked up on the next opeation.
         * TBD - for non-blocking, should we queue a callback for errors?
         */
        if (ctxP->flags & PIPE_F_WATCHREAD)
            ctxP->winerr = TwapiPipeEnableReadWatch(ctxP);
        return TCL_OK;
    }
    /* Fall thru for errors */

error_return:
    if (ctxP->winerr == ERROR_HANDLE_EOF) {
        Tcl_ResetResult(ticP->interp);
        return TCL_OK;      /* EOF -> empty return result. TBD what about EPIPE? */
    }
    return Twapi_AppendSystemError(ticP->interp, ctxP->winerr);
}

int Twapi_PipeWatch(TwapiInterpContext *ticP, HANDLE hpipe, int watch_read, int watch_write) 
{
    TwapiPipeContext *ctxP;
    DWORD winerr;
    int data_availability = 0;

    TWAPI_ASSERT(ticP->thread == Tcl_GetCurrentThread());
    TWAPI_ASSERT(direction == READER || direction == WRITER);
    
    ZLIST_LOCATE(ctxP, &ticP->pipes, hpipe, hpipe);
    if (ctxP == NULL)
        return TwapiReturnTwapiError(ticP->interp, NULL, TWAPI_UNKNOWN_OBJECT);

    TWAPI_ASSERT(ticP == ctxP->ticP);

    CHeck winerr;

    /* First do the read side stuff */
    if (watch_read) {
TBD        
    }
    if ((flags & 1) == 0) {
        /* Do not want read notifications. Unregister from thread pool */
        TBD - do not unregister if state is IO_PENDING - let callback do that
        if (ctxP->io[READER].hwait != INVALID_HANDLE_VALUE) {
            UnregisterWaitEx(ctxP->io[READER].hwait, INVALID_HANDLE_VALUE);
            ctxP->io[READER].hwait = INVALID_HANDLE_VALUE;
            TwapiPipeContextUnref(ctxP, 1);
        }
    } else {
        /* Want to be notified when data is available */
        if (ctxP->io[READER].hwait == INVALID_HANDLE_VALUE) {
            /*
             * Notifications are not set up so do so.
             * If there is already data available, in our input buffer,
             * we will return "data available" so script level code
             * will queue a notification. In this case, we do not
             * enqueue an OS I/O since we are only using a single
             * data buffer and that already has data. The kernel
             * level async I/O will be set up when the existing data
             * is read.
             */
            switch (ctxP->io[READER].state) {
            case IOBUF_IO_COMPLETED:
                /* We have data in the buffer. */
                data_availability |= 1;
                break;
            case IOBUF_IDLE:
                /* Setup a read. Note it may complete asynchronously */
                ZeroMemory(&ctxP->io[READER].ovl, sizeof(ctxP->io[READER].ovl));
                ctxP->io[READER].ovl.hEvent = ctxP->io[READER].hevent;
                if (ReadFile(ctxP->hpipe,
                             &ctxP->io[READER].data.read_ahead,
                             1,
                             NULL,
                             &ctxP->io[READER].ovl)) {
                    /* Have data already. */
                    data_availability |= 1;
                    break;
                }
                if ((winerr = GetLastError()) != ERROR_IO_PENDING) {
                    /* Genuine error */
                    return TwapiPipeMapError(ticP->interp, winerr);
                }
                             
                /* ERROR_IO_PENDING - FALLTHRU */
            case IOBUF_IO_PENDING:
                /* The thread pool will hold a ref to the context */
                TwapiPipeContextRef(ctxP, 1);
                if (! RegisterWaitForSingleObject(
                        &ctxP->io[READER].hwait,
                        ctxP->io[READER].hevent,
                        TwapiPipeReadThreadPoolFn,
                        ctxP,
                        INFINITE,           /* No timeout */
                        WT_EXECUTEDEFAULT
                        )) {
                    TwapiPipeContextUnref(ctxP, 1);
                    /* Note the call does not set GetLastError */
                    return TwapiReturnTwapiError(ticP->interp, "Could not register wait on pipe.", TWAPI_SYSTEM_ERROR);
                }
            }
        } else {
            /*
             * Wait is already registered. Nothing to do. If there
             * was data available, the callback would have, or will,
             * queue the notification.
             */
        }
    }


    /* Now ditto for the write side. */
    if (flags & 2) {
        /* Register notification if not already done so */
        if (ctxP->io[WRITER].hwait == INVALID_HANDLE_VALUE) {
            /*
             * Notifications are not set up so do so.
             */
            if (ctxP->io[WRITER].state == IOBUF_IO_PENDING) {
                TWAPI_ASSERT(ctxP->io[WRITER].data.objP);
                /* Ongoing operation, set up notification when it is complete */
                /* The thread pool will hold a ref to the context */
                TwapiPipeContextRef(ctxP, 1);
                if (! RegisterWaitForSingleObject(
                        &ctxP->io[WRITER].hwait,
                        ctxP->io[WRITER].hevent,
                        TwapiPipeWriteThreadPoolFn,
                        ctxP,
                        INFINITE,           /* No timeout */
                        WT_EXECUTEDEFAULT
                        )) {
                    TwapiPipeContextUnref(ctxP, 1);
                    /* Note the call does not set GetLastError */
                    return TwapiReturnTwapiError(ticP->interp, "Could not register wait on pipe.", TWAPI_SYSTEM_ERROR);

                }
            } else {
                if (ctxP->io[WRITER].data.objP) {
                    Tcl_DecrRefCount(ctxP->io[WRITER].data.objP);
                    ctxP->io[WRITER].data.objP = NULL;
                    ctxP->io[WRITER].state = IOBUF_IDLE;
                }
                data_availability |= 2; /* Writable */
            }
        } else {
            /*
             * Wait is already registered. Nothing to do. If there
             * was data available, the callback would have, or will,
             * queue the notification.
             */
        }
    } else {
        /* Do not want notifications. Unregister from thread pool */
        if (ctxP->io[WRITER].hwait != INVALID_HANDLE_VALUE) {
            UnregisterWaitEx(ctxP->io[WRITER].hwait, INVALID_HANDLE_VALUE);
            ctxP->io[WRITER].hwait = INVALID_HANDLE_VALUE;
            TwapiPipeContextUnref(ctxP, 1);
        }
    }

    if (data_availability) {
        Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);
        if (data_availability & (1 << READER))
            Tcl_ListObjAppendElement(ticP->interp, resultObj, STRING_LITERAL_OBJ("readable"));
        if (data_availability & (1 << WRITER))
            Tcl_ListObjAppendElement(ticP->interp, resultObj, STRING_LITERAL_OBJ("writable"));

        Tcl_SetObjResult(ticP->interp, resultObj);
    }
    return TCL_OK;
}


static DWORD TwapiPipeConnectClient(TwapiPipeContext *ctxP)
{
    DWORD winerr;

    /* The reader side i/o structures are is also used for connecting */

    /*
     * Wait for client connection. If there is already a client waiting,
     * to connect, we will get success right away. Else we wait on the
     * event and queue a callback later
     */
    if (ConnectNamedPipe(ctxP->hpipe, &ctxP->io[READER].ovl) ||
        (winerr = GetLastError()) == ERROR_PIPE_CONNECTED) {
        /* Already a client in waiting, queue the callback */
        return ERROR_SUCCESS;
    } else if (winerr == ERROR_IO_PENDING) {
        /* Wait asynchronously for a connection */
        TwapiPipeContextRef(ctxP, 1); /* Since we are passing to thread pool */
        if (RegisterWaitForSingleObject(
                &ctxP->io[READER].hwait,
                ctxP->io[READER].hevent,
                TwapiPipeConnectThreadPoolFn,
                ctxP,
                INFINITE,           /* No timeout */
                WT_EXECUTEDEFAULT
                )) {
            return ERROR_IO_PENDING;
        } else {
            winerr = GetLastError();
            TwapiPipeContextUnref(ctxP, 1); /* Undo above Ref */
        }
    }

    /* Either synchronous completion or error */
    return winerr;
}



/* Always returns non-NULL, or panics */
static TwapiPipeContext *TwapiPipeContextNew(void)
{
    int i;
    TwapiPipeContext *ctxP;

    ctxP = (TwapiPipeContext *) TwapiAlloc(sizeof(*ctxP));
    ctxP->ticP = NULL;
    ctxP->hpipe = INVALID_HANDLE_VALUE;
    ctxP->hsync = NULL;

    ctxP->io[READER].hwait = INVALID_HANDLE_VALUE;
    ctxP->io[READER].hevent = NULL;
    ctxP->io[READER].state = IOBUF_IDLE;

    ctxP->io[WRITER].hwait = INVALID_HANDLE_VALUE;
    ctxP->io[WRITER].hevent = NULL;
    ctxP->io[WRITER].state = IOBUF_IDLE;
    ctxP->io[WRITER].data.objP = NULL;

    ctxP->flags = 0;

    ctxP->winerr = ERROR_SUCCESS;
    ctxP->nrefs = 0;
    ZLINK_INIT(ctxP);
    return ctxP;
}

static void TwapiPipeContextDelete(TwapiPipeContext *ctxP)
{
    int i;

    TWAPI_ASSERT(ctxP->nrefs <= 0);
    TWAPI_ASSERT(ctxP->io[READER].hwait == NULL);
    TWAPI_ASSERT(ctxP->io[READER].hevent == NULL);
    TWAPI_ASSERT(ctxP->io[WRITER].hwait == NULL);
    TWAPI_ASSERT(ctxP->io[WRITER].hevent == NULL);
    TWAPI_ASSERT(ctxP->hsync == NULL);
    TWAPI_ASSERT(ctxP->hpipe == INVALID_HANDLE_VALUE);

    if (ctxP->io[WRITER].data.objP)
        Tcl_DecrRefCount(ctxP->io[WRITER].data.objP);

    if (ctxP->ticP)
        TwapiInterpContextUnref(ctxP->ticP, 1);

    TwapiFree(ctxP);
}

void TwapiPipeContextUnref(TwapiPipeContext *ctxP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized pcbP */
    if (InterlockedExchangeAdd(&ctxP->nrefs, -decr) <= decr)
        TwapiPipeContextDelete(ctxP);
}

/*
 * Initiates shut down of a pipe. It unregisters the context from ticP,
 * and  thread pool which means the ctxP may be deallocated before returning
 * unless caller holds some other ref it.
 */
static int TwapiPipeShutdown(TwapiPipeContext *ctxP)
{
    int unrefs = 0;             /* How many times we need to unref  */
    int i;

    /*
     * We need to do things in a specific order.
     *
     * First, unlink the pipe context and ticP, so no callbacks will access
     * the interp/tic.
     *
     * Note all unrefs for the dmc are done at the end.
     */
    if (ctxP->ticP) {
        ZLIST_REMOVE(&ctxP->ticP->pipes, ctxP);
        TwapiInterpContextUnref(ctxP->ticP, 1);
        ctxP->ticP = NULL;
        ++unrefs;
    }

    /*
     * Second, stop the thread pool for this pipe. We need to do that before
     * closing handles. Note the UnregisterWaitEx can result in thread pool
     * callbacks running while it is blocked. The callbacks might queue
     * additional events to the interp thread. That's ok because we 
     * ctxP->ticP above.
     */
    if (ctxP->io[READER].hwait != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(ctxP->io[READER].hwait,
                         INVALID_HANDLE_VALUE);
        ctxP->io[READER].hwait = INVALID_HANDLE_VALUE;
        ++unrefs;   /* Remove the ref coming from the thread pool */
    }
    if (ctxP->io[READER].hevent != NULL) {
        CloseHandle(ctxP->io[READER].hevent);
        ctxP->io[READER].hevent = NULL;
    }

    if (ctxP->io[WRITER].hwait != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(ctxP->io[WRITER].hwait,
                         INVALID_HANDLE_VALUE);
        ctxP->io[WRITER].hwait = INVALID_HANDLE_VALUE;
        ++unrefs;   /* Remove the ref coming from the thread pool */
    }
    if (ctxP->io[WRITER].hevent != NULL) {
        CloseHandle(ctxP->io[WRITER].hevent);
        ctxP->io[WRITER].hevent = NULL;
    }

    /* Third, now that handles are unregistered, close them. */
    if (ctxP->hpipe != INVALID_HANDLE_VALUE)
        CloseHandle(ctxP->hpipe);

    if (ctxP->hsync != NULL)
        CloseHandle(ctxP->hsync);

    if (unrefs)
        TwapiPipeContextUnref(ctxP, unrefs); /* May be GONE! */

    return TCL_OK;
}

static int TwapiPipeMapError(TwapiInterpContext *ticP, DWORD winerr)
{
    switch (winerr) {
    case ERROR_HANDLE_EOF:
        Tcl_SetResult(ticP->interp, "eof", TCL_STATIC);
        break;
    case ERROR_BROKEN_PIPE:
        Tcl_SetResult(ticP->interp, "epipe", TCL_STATIC);
        break;
    default:
        return Twapi_AppendSystemError(ticP->interp, winerr);
    }
    return TCL_OK;
}
