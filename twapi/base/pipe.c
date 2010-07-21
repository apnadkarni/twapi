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
        int    state;
#define IOBUF_IDLE         0
#define IOBUF_IO_PENDING   1
#define IOBUF_IO_COMPLETED 2
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
static DWORD TwapiSetupPipeRead(TwapiPipeContext *ctxP);


static WIN32_ERROR TwapiPipeWatchReads(TwapiPipeContext *ctxP)
{
    DWORD winerr;

    TWAPI_ASSERT(ctxP->winerr == ERROR_SUCCESS);
    TWAPI_ASSERT(ctxP->io[READER].state == IOBUF_IDLE);

    if (! (ctxP->flags & PIPE_F_WATCHREAD))
        return ERROR_SUCCESS;

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
            /* Note the call does not set GetLastError. Make up our own */
            return 
        }

    }
    ctxP->io[READER].state = IOBUF_IO_PENDING;

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
        /* Create event used for sync i/o */
        ctxP->hsync = CreateEvent(NULL, TRUE, FALSE, NULL);

        if (ctxP->hsync) {
            /*
             * Create events to use for notification of completion. The events
             * must be auto-reset to prevent multiple callback queueing on a single 
             * input notification. See MSDN docs for RegisterWaitForSingleObject.
             * As a consequence, we must make sure we never call GetOverlappedResult
             * in blocking mode.
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

    TWAPI_ASSERT(ticP->thread == Tcl_GetCurrentThread());
    
    ZLIST_LOCATE(ctxP, &ticP->pipes, hpipe, hpipe);
    if (ctxP == NULL)
        return TwapiReturnTwapiError(ticP->interp, NULL, TWAPI_UNKNOWN_OBJECT);

    TWAPI_ASSERT(ticP == ctxP->ticP);

    ioP = &ctxP->io[READER];
    if (ioP->state == IOBUF_IO_PENDING) {
        TWAPI_ASSERT(ioP->hwait != INVALID_HANDLE_VALUE);
        /*
         * If in non-blocking mode, throw EAGAIN as expected by the Tcl 
         * channel implementation.
         */
        if (ctxP->flags & PIPE_F_NONBLOCKING) {
            Tcl_SetResult(interp, "EAGAIN", TCL_STATIC);
            return TCL_ERROR;
        }

        /*
         * Loop waiting for I/O to complete. Note this combination of
         * non-blocking I/O with use of fileevent (async notifications)
         * is rare so no need for more sophisticated mechanisms. The
         * thread pool handler will set state to IOBUF_IO_COMPLETED.
         */
        while (ioP->state == IOBUF_IO_PENDING) {
            SleepEx(1, FALSE);
        }
    }

    TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING);
    
    nread = 0;
    if (ioP->state == IOBUF_IO_COMPLETE) {
        TWAPI_ASSERT(HasOverlappedIoCompleted(&ioP->ovl));
        if (! SUCCEEDED(ioP->ovl.Internal)) {
            return Twapi_AppendSystemError(ticP->interp, ioP->ovl.Internal);
        }
        TWAPI_ASSERT(ioP->ovl.InternalHigh == 1); /* One-byte read ahead */
        nread = 1;
    }

    /*
     * We are either in state IDLE or IO_COMPLETED successfully. In either
     * case, see how much additional data is available in the pipe itself.
     * TBD - with overlapped i/o can we just do a read or will that only
     * complete when specified number of bytes is available ?
     */
    avail = 0;
    if (count != nread) {
        if (! PeekNamedPipe(ctxP->hpipe, NULL, 0, NULL, &avail, NULL)) {
            TBD - error;
        }
    }
    
    if ((count <= (avail+nread)) {
        
    }



}

int Twapi_PipeWatch(TwapiInterpContext *ticP, HANDLE hpipe, DWORD flags) 
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

    /* First do the read side stuff (bit 0) */
    if ((flags & 1) == 0) {
        /* Do not want read notifications. Unregister from thread pool */
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
