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
        HANDLE hwait;       /* Handle returned by thread pool wait functions. */
        HANDLE hevent;      /* Event for thread pool to wait on */
        union {
            char read_ahead[1];     /* Used to read-ahead a single byte */
            Tcl_Obj *objP;      /* Data to write out */
        } data;
        LONG volatile state;
#define IOBUF_IDLE         0    /* I/O buffer not in use, no pending ops */
#define IOBUF_IO_PENDING   1    /* Overlapped I/O has been queued. Does
                                   not mean thread pool is waiting on it. */
#define IOBUF_IO_COMPLETED 2    /* Overlapped I/O completed, buffer in use */
#define IOBUF_IO_COMPLETED_WITH_ERROR 3 /* Overlapped I/O completed with error */
    } io[2];                   /* 0 -> read, 1 -> write */
#define READER 0
#define WRITER 1

    int    flags;
#define PIPE_F_WATCHREAD       1
#define PIPE_F_WATCHWRITE      2
#define PIPE_F_NONBLOCKING     4
#define PIPE_F_CONNECTED       8

    ZLINK_DECL(TwapiPipeContext);
    ULONG volatile nrefs;              /* Ref count */
    WIN32_ERROR winerr;

#define SET_CONTEXT_WINERR(ctxP_, err_) \
    ((ctxP_)->winerr == ERROR_SUCCESS ? ((ctxP_)->winerr = (err_)) : (ctxP_)->winerr)
} TwapiPipeContext;

#define PIPE_READ_COMPLETE    0
#define PIPE_WRITE_COMPLETE   1
#define PIPE_CONNECT_COMPLETE 2

/*
 * Prototypes
 */
static TwapiPipeContext *TwapiPipeContextNew(void);
static void TwapiPipeContextDelete(TwapiPipeContext *ctxP);
#define TwapiPipeContextRef(p_, incr_) InterlockedExchangeAdd(&(p_)->nrefs, (incr_))
void TwapiPipeContextUnref(TwapiPipeContext *ctxP, int decr);
static int TwapiPipeShutdown(TwapiPipeContext *ctxP);
static DWORD TwapiPipeConnectClient(TwapiPipeContext *ctxP);
static DWORD TwapiPipeCallbackFn(TwapiCallback *cbP);


static void TwapiPipeEnqueueCallback(
    TwapiPipeContext *ctxP,
    int pipe_event)
{

    struct _TwapiPipeIO *ioP = &ctxP->io[pipe_event == PIPE_WRITE_COMPLETE ? WRITER : READER];
    TwapiCallback *cbP;

    TWAPI_ASSERT(HasOverlappedIoCompleted(&ioP->ovl));

    cbP = TwapiCallbackNew(ctxP->ticP, TwapiPipeCallbackFn, sizeof(*cbP));
    cbP->winerr = ioP->ovl.Internal;
    TwapiPipeContextRef(ctxP, 1);
    cbP->clientdata = (DWORD_PTR) ctxP;
    cbP->clientdata2 = pipe_event;
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
    TwapiPipeEnqueueCallback((TwapiPipeContext*) lpParameter, PIPE_READ_COMPLETE);
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
    TwapiPipeEnqueueCallback((TwapiPipeContext*) lpParameter, PIPE_WRITE_COMPLETE);
}

/*
 * Called from thread pool when a connect completes.
 * Note this has the typedef for WaitOrTimerCallback 
 */
static VOID CALLBACK TwapiPipeConnectThreadPoolFn(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
)
{
    TWAPI_ASSERT(TimerOrWaitFired == FALSE);
    TwapiPipeEnqueueCallback((TwapiPipeContext*) lpParameter, PIPE_CONNECT_COMPLETE);
}

/*
 * Sets up overlapped reads on a pipe. The read state must be IDLE so
 * that the data buffer is not in use and no reads are pending.
 */
static WIN32_ERROR TwapiPipeEnableReadWatch(TwapiPipeContext *ctxP)
{
    if (ctxP->winerr != ERROR_SUCCESS ||
        ctxP->io[READER].state != IOBUF_IDLE ||
        ! (ctxP->flags & PIPE_F_WATCHREAD)) {
        /* Note we do not mark it as a pipe error (ctxP->winerr) */
        return TWAPI_ERROR_TO_WIN32(TWAPI_BUG_INVALID_STATE_FOR_OP);
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
        WIN32_ERROR winerr = GetLastError();
        if (winerr != ERROR_IO_PENDING) {
            ctxP->winerr = winerr;
            return winerr;
        }
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
            ctxP->winerr = TWAPI_ERROR_TO_WIN32(TWAPI_REGISTER_WAIT_FAILED);
            return ctxP->winerr;
        }

    }
    ctxP->io[READER].state = IOBUF_IO_PENDING;

    return ERROR_SUCCESS;
}

static WIN32_ERROR TwapiPipeDisableWatch(TwapiPipeContext *ctxP, int direction)
{
    if (ctxP->io[direction].hwait != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(ctxP->io[direction].hwait, INVALID_HANDLE_VALUE);
        ctxP->io[direction].hwait = INVALID_HANDLE_VALUE;
        TwapiPipeContextUnref(ctxP, 1);
    }

    ctxP->flags &= ~ (direction == READER ? PIPE_F_WATCHREAD : PIPE_F_WATCHWRITE);

    /*
     * Note io state might still be PENDING with an outstanding overlapped
     * operation that has not yet completed.
     */

    return ERROR_SUCCESS;
}

/* Called from Tcl event loop with a pipe event notification */
static DWORD TwapiPipeCallbackFn(TwapiCallback *cbP)
{
    Tcl_Obj *objs[4];
    int nobjs;
    TwapiPipeContext *ctxP = (TwapiPipeContext *) cbP->clientdata;
    int flags;
    char *eventstr;
    int pipe_event;
    int direction;

    /* If the interp is gone, close down the pipe */
    if (ctxP->ticP == NULL ||
        ctxP->ticP->interp == NULL ||
        Tcl_InterpDeleted(ctxP->ticP->interp)) {
        cbP->clientdata = 0;
        TwapiPipeShutdown(ctxP);
        TwapiPipeContextUnref(ctxP, 1); /* Corresponds to coming off event q */
        ctxP = NULL;            /* Ensure we do not access it */
        cbP->winerr = ERROR_INVALID_FUNCTION; // TBD
        cbP->response.type = TRT_EMPTY;
        return TCL_ERROR;
    }

    /*
     * Unref corresponding to coming off event queue. Note ctxP still valid
     * as it is linked to ticP so ref count will be > 0
     */
    TwapiPipeContextUnref(ctxP, 1);

    pipe_event = cbP->clientdata2;
    switch (pipe_event) {
    case PIPE_READ_COMPLETE:
        flags = PIPE_F_WATCHREAD;
        eventstr = "read";
        direction = READER;
        break;
    case PIPE_WRITE_COMPLETE:
        flags = PIPE_F_WATCHWRITE;
        eventstr = "write";
        direction = WRITER;
        break;
    case PIPE_CONNECT_COMPLETE:
        flags = 0;
        if (cbP->winerr == ERROR_SUCCESS)
            ctxP->flags |= PIPE_F_CONNECTED;
        eventstr = "connect";
        direction = READER;     /* CONNECT also uses READER iobuf */
        break;
    }

    /*
     * We need to check for the state because while the callback was queued,
     * the application might have turned off file events, and made synchronous
     * reads/write which would unregister the thread pool and also potentially
     * transition the state. In this case we do not want to overwrite the
     * state. Note we do not go on to queue the notification either.
     */                                                  
    if (ctxP->io[direction].state != IOBUF_IO_PENDING) {
        cbP->winerr = ERROR_SUCCESS;
        cbP->response.type = TRT_EMPTY;
        return TCL_OK;
    }

    if (cbP->winerr == ERROR_SUCCESS)
        ctxP->io[direction].state = IOBUF_IO_COMPLETED;
    else {
        ctxP->io[direction].state = IOBUF_IO_COMPLETED_WITH_ERROR;
        SET_CONTEXT_WINERR(ctxP, cbP->winerr);
    }

    /*
     * Check if pipe is still being watched. We only queue a notification
     * if that is the case. Moreover, if it is not being watched, we need to
     * turn off notifications. Connects are also notified.
     */
    if (pipe_event == PIPE_CONNECT_COMPLETE || (ctxP->flags & flags)) {
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
        if (cbP->clientdata2 == PIPE_READ_COMPLETE)
            TwapiPipeDisableWatch(ctxP, READER);
        else
            TwapiPipeDisableWatch(ctxP, WRITER);

        cbP->winerr = ERROR_SUCCESS;
        cbP->response.type = TRT_EMPTY;
        return TCL_OK;
    }
}




/*
 * Return the current number of available bytes for reading including the
 * read-ahead byte.
 */
WIN32_ERROR TwapiPipeReadCount(TwapiPipeContext *ctxP, DWORD *countP)
{
    DWORD readahead_count;
    DWORD pipe_count;

    TWAPI_ASSERT(ctxP->winerr == ERROR_SUCCESS);

    if (ctxP->io[READER].state == IOBUF_IO_PENDING) {
        /*
         * Even when pending, the I/O could have completed. Checking for
         * this is important, not just for better latency but also because
         * the thread pool might have been unregistered if script is
         * not interested in file events any more
         */
        if (! HasOverlappedIoCompleted(&ctxP->io[READER].ovl)) {
            *countP = 0;               /* I/O pending so obviously no data */
            return ERROR_SUCCESS;
        }

        if (ctxP->io[READER].ovl.Internal != ERROR_SUCCESS) {
            ctxP->io[READER].state = IOBUF_IO_COMPLETED_WITH_ERROR;
            ctxP->winerr = ctxP->io[READER].ovl.Internal;
            return ctxP->winerr;
        }
        ctxP->io[READER].state = IOBUF_IO_COMPLETED;
    }

    /*
     * Note in COMPLETED, COMPLETED_WITH_ERROR or IDLE state, the thread
     * pool will not run so we do not have to worry about consistency 
     * between the read ahead and what's in the pipe. Also, even if
     * the state is COMPLETED_WITH_ERROR, we still allow the reads.
     */
       
    if (ctxP->io[READER].state == IOBUF_IO_COMPLETED)
        readahead_count = 1;              /* The read-ahead byte */
    else
        readahead_count = 0;

    if (!PeekNamedPipe(ctxP->hpipe, NULL, 0, NULL, &pipe_count, NULL)) {
        /* Error, but if we have a read ahead byte, return it */
        if (readahead_count == 0) {
            ctxP->winerr = GetLastError();
            return ctxP->winerr;
        }
        pipe_count = 0;
    }

    *countP = pipe_count + readahead_count;

    return ERROR_SUCCESS;
}


/*
 * Read count bytes and return a Tcl_Obj. Caller must have ensured there
 * are that many bytes available else function will block waiting for data
 * (which is ok for blocking reads but not for nonblocking reads)
 */
static WIN32_ERROR TwapiPipeReadData(TwapiPipeContext *ctxP, DWORD count, Tcl_Obj **objPP)
{
    struct _TwapiPipeIO *ioP = &ctxP->io[READER];
    Tcl_Obj *objP = Tcl_NewByteArrayObj(NULL, count);
    char *p = Tcl_GetByteArrayFromObj(objP, NULL);

    TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING);
    TWAPI_ASSERT(count);
    TWAPI_ASSERT(ctxP->winerr == ERROR_SUCCESS);

    if (ioP->state == IOBUF_IO_COMPLETED) {
        --count;
        *p++ = ioP->data.read_ahead[0];
        ioP->state = IOBUF_IDLE;
    }

    if (count > 0) {
        DWORD read_count;
        ZeroMemory(&ioP->ovl, sizeof(ioP->ovl));
        ioP->hevent = ctxP->hsync; /* Event used for "synchronous" reads */
        if (! ReadFile(ctxP->hpipe, p, count, NULL, &ioP->ovl)) {
            WIN32_ERROR winerr = GetLastError();
            if (winerr != ERROR_IO_PENDING) {
                Tcl_DecrRefCount(objP);
                ctxP->winerr = winerr;
                return winerr;
            }
            /*
             * Pending I/O - when we have blocking reads and not enough data
             * in pipe. Fall thru, will block in GetOverlappedResult
             */
        }
        if (!GetOverlappedResult(ctxP->hpipe, &ioP->ovl, &read_count, TRUE)) {
            ctxP->winerr = GetLastError();
            Tcl_DecrRefCount(objP);
            return ctxP->winerr;
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

    TwapiPipeShutdown(ctxP);    /* ctxP might be gone */
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
    DWORD avail;
    struct _TwapiPipeIO *ioP;
    Tcl_Obj *objP;

    TWAPI_ASSERT(ticP->thread == Tcl_GetCurrentThread());
    
    ZLIST_LOCATE(ctxP, &ticP->pipes, hpipe, hpipe);
    if (ctxP == NULL)
        return TwapiReturnTwapiError(ticP->interp, NULL, TWAPI_UNKNOWN_OBJECT);

    TWAPI_ASSERT(ticP == ctxP->ticP);

    if ((ctxP->flags & PIPE_F_CONNECTED) == 0)
        return Twapi_AppendSystemError(ticP->interp, ERROR_PIPE_NOT_CONNECTED);

    ioP = &ctxP->io[READER];
    if (ctxP->winerr != ERROR_SUCCESS)
        goto error_return;

    if (count == 0)
        return TCL_OK; // TBD ? Will get treated as EOF

    if ((ctxP->flags & PIPE_F_NONBLOCKING) == 0 &&
        ioP->state == IOBUF_IO_PENDING) {
        /*
         * If Blocking channel and state is PENDING, we are waiting for data.
         * Also, there may be a thread pool
         * thread waiting on it via an overlapped read. Unregister it.
         * Might not be the most efficient mechanism but the combination
         * of fileevent with blocking channels is unlikely anyways.
         */
        if (ioP->hwait != INVALID_HANDLE_VALUE) {
            UnregisterWaitEx(ctxP->io[READER].hwait, INVALID_HANDLE_VALUE);
            ctxP->io[READER].hwait = INVALID_HANDLE_VALUE;
            TwapiPipeContextUnref(ctxP, 1);
        }
    }

    if (TwapiPipeReadCount(ctxP, &avail) != ERROR_SUCCESS) {
        goto error_return;
    }

    /* If avail is non-0, state CANNOT be PENDING */
    TWAPI_ASSERT(avail == 0 || ioP->state != IOBUF_IO_PENDING);
    /* If state is PENDING, avail must be 0 */
    TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING || avail == 0);

    if (ctxP->flags & PIPE_F_NONBLOCKING) {
        /* Non blocking pipe */
        if (avail == 0) {
            /* No bytes are available */
            Tcl_SetResult(ticP->interp, "EAGAIN", TCL_STATIC);
            return TCL_ERROR;   /* As expected by Tcl channel implementation */
        }
        if (avail < count)
            count = avail;      /* Return what we have */
    } else {
        /*
         * Blocking pipe. If state is PENDING, we are waiting for i/o
         * completion. Block on it. Note because we unregistered the thread
         * pool above, and then TwapiPipeReadCount checked the i/o event state,
         * we do not have to worry either about race conditions or that the
         * completion event in the OVERLAPPED buffer being already reset.
         */
        if (ioP->state == IOBUF_IO_PENDING) {
            TWAPI_ASSERT(avail == 0);
            if (GetOverlappedResult(ctxP->hpipe, &ioP->ovl, &avail, TRUE) == 0) {
                /* Failure */
                ioP->state = IOBUF_IO_COMPLETED_WITH_ERROR;
                ctxP->winerr = GetLastError();
                goto error_return;
            }
            ioP->state = IOBUF_IO_COMPLETED;
        }
    }

    if (TwapiPipeReadData(ctxP, count, &objP) == ERROR_SUCCESS) {
        Tcl_SetObjResult(ticP->interp, objP);

        /*
         * Need to fire another read event if more data is available. This
         * is also required in case we unregistered the thread pool wait
         * in the blocking I/O case above.
         * PipeWatch may return an error but we do not report that right away.
         * It will be picked up on the next operation.
         * TBD - for non-blocking, should we queue a callback for errors?
         */
        if (ctxP->flags & PIPE_F_WATCHREAD)
            TwapiPipeEnableReadWatch(ctxP);
        /*
         * Note error, if any, from TwapiPipeEnableReadWatch ignored,
         * will be picked up on next call
         */
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

#if 0
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
#endif

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
        ctxP->flags |= PIPE_F_CONNECTED;
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
