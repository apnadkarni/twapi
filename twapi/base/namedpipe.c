/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

typedef struct _NPipeChannel NPipeChannel;
ZLINK_CREATE_TYPEDEFS(NPipeChannel); 
ZLIST_CREATE_TYPEDEFS(NPipeChannel);

typedef struct _NPipeChannel {
    ZLINK_DECL(NPipeChannel); /* List of registered pipes. Access sync
                                   through global pipe channel lock */
    Tcl_ThreadId thread;    /* The thread owning the channel */
    Tcl_Channel channel;           /* The corresponding Tcl_Channel handle */
    HANDLE  hpipe;   /* Handle to the pipe */
    HANDLE  hsync;   /* Event used for synchronous I/O, both read and write.
                        Needed because we cannot use the hevent field
                        for sync i/o as the thread pool will also be waiting
                        on it */
    struct _NPipeIO {
        OVERLAPPED ovl;         /* Used in async i/o. Note the hEvent field
                                   is always set at init time */
        HANDLE hwait;       /* Handle returned by thread pool wait functions. */
        HANDLE hevent;      /* Event for thread pool to wait on */
        union {
            char read_ahead[1];     /* Used to read-ahead a single byte */
            struct {
                char *p;      /* Data to write out */
                DWORD sz;   /* Size of buffer */
                DWORD len;  /* Amount of data in buffer */
            } write_buf;
        } data;
        LONG volatile state;    /* State values IOBUF_*. Writing PENDING to
                                   this location or writing to this location
                                   while it contains PENDING must be done
                                   using Interlocked* Win32 API since in
                                   that state the worker threads also access
                                   the location. */
#define IOBUF_IDLE         0    /* I/O buffer not in use, no pending ops */
#define IOBUF_IO_PENDING   1    /* Overlapped I/O has been queued. Does
                                   not mean thread pool is waiting on it. */
#define IOBUF_IO_COMPLETED 2    /* Overlapped I/O completed, buffer in use */
#define IOBUF_IO_COMPLETED_WITH_ERROR 3 /* Overlapped I/O completed with error */
    } io[2];                   /* 0 -> read, 1 -> write */
#define READER 0
#define WRITER 1

    int    flags;
#define NPIPE_F_WATCHREAD       1 /* Generate event when data available */
#define NPIPE_F_WATCHWRITE      2 /* Generate event when output possible */
#define NPIPE_F_NONBLOCKING     4 /* Channel is async */
#define NPIPE_F_CONNECTED       8 /* Client has successfully connected */
#define NPIPE_F_EVENT_QUEUED   16 /* A TCL event has been queued */ 

    ULONG volatile nrefs;              /* Ref count */
    WIN32_ERROR winerr;

#define SET_NPIPE_ERROR(ctxP_, err_) \
    ((ctxP_)->winerr == ERROR_SUCCESS ? ((ctxP_)->winerr = (err_)) : (ctxP_)->winerr)
};

#define NPIPE_CONNECTED(pcP_) ((pcP_)->flags & NPIPE_F_CONNECTED)
/*
 * When should we notify for a read - data available in buffer or error/eof.
 * Errors other than eof are only notified in connecting stage.
 */
#define NPIPE_READ_NOTIFIABLE(pcP_) \
    (((pcP_)->io[READER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->winerr == ERROR_HANDLE_EOF) ||                            \
     ((pcP_)->winerr != ERROR_SUCCESS && !NPIPE_CONNECTED(pcP_)))

/*
 * When should we notify for a write - data i/o completed, or error/eof
 * Note we do not include IDLE state as then we would continually generate
 * (potentially) notifications.
 * Errors other than eof are only notified in connecting stage.
 */
#define NPIPE_WRITE_NOTIFIABLE(pcP_) \
    (((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->winerr == ERROR_HANDLE_EOF) ||                            \
     ((pcP_)->winerr != ERROR_SUCCESS && !NPIPE_CONNECTED(pcP_)))

/* Combination of above */
#define NPIPE_NOTIFIABLE(pcP_) \
    (((pcP_)->io[READER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->winerr == ERROR_HANDLE_EOF) ||                            \
     ((pcP_)->winerr != ERROR_SUCCESS && !NPIPE_CONNECTED(pcP_)))

/*
 * Pipe channels are maintained on a per thread basis. 
 */
typedef struct _NPipeTls {
    /*
     * Create list header definitions for list of pipes.
     */
    CRITICAL_SECTION lock;
    ZLIST_DECL(NPipeChannel) pipes;
    
} NPipeTls;

static int gNPipeTlsSlot;
/* Use only when sure that module and thread init has been done */
#define GET_NPIPE_TLS() ((NPipeTls *) TWAPI_TLS_SLOT(gNPipeTlsSlot))
#define SET_NPIPE_TLS(p_) do {TWAPI_TLS_SLOT(gNPipeTlsSlot) = (DWORD_PTR) (p_);} while (0)

typedef struct _NPipeEvent {
    Tcl_Event header;
    HANDLE hpipe;               /* Pipe to which this event relates */
} NPipeEvent;


static TwapiOneTimeInitState gNPipeModuleInitialized;

/* Prototypes */
static int NPipeEventProc(Tcl_Event *, int flags);

static Tcl_DriverCloseProc NPipeCloseProc;
static Tcl_DriverInputProc NPipeInputProc;
static Tcl_DriverOutputProc NPipeOutputProc;
static Tcl_DriverWatchProc NPipeWatchProc;
static Tcl_DriverGetHandleProc NPipeGetHandleProc;
static Tcl_DriverBlockModeProc NPipeBlockProc;
static Tcl_DriverThreadActionProc NPipeThreadActionProc;

static Tcl_ChannelType gNPipeChannelDispatch = {
    "namedpipe",
    (Tcl_ChannelTypeVersion)TCL_CHANNEL_VERSION_4,
    NPipeCloseProc,
    NPipeInputProc,
    NPipeOutputProc,
    NULL /* ChannelSeek */,
    NULL /* ChannelSetOption */,
    NULL /* ChannelGetOption */,
    NPipeWatchProc,
    NPipeGetHandleProc,
    NULL /* ChannelClose2 */,
    NPipeBlockProc,
    NULL /* ChannelFlush */,
    NULL /* ChannelHandler */,
    NULL /* ChannelWideSeek if VERSION_2, Truncate if VERSION_5 */,
    NPipeThreadActionProc,       /* Unused for VERSION_2 */
};

static NPipeChannel *NPipeChannelNew(void);
static void NPipeChannelDelete(NPipeChannel *pcP);
#define NPipeChannelRef(p_, incr_) InterlockedExchangeAdd(&(p_)->nrefs, (incr_))
void NPipeChannelUnref(NPipeChannel *pcP, int decr);
static NPipeTls *GetNPipeTls();
static void NPipeShutdown(NPipeChannel *pcP, int unrefs);

/*
 * Map Win32 errors to Tcl errno. Note it is important to use the same
 * mapping as Tcl does so we have to call into Tcl to set errno and then
 * retrieve the value.
 */
static int NPipeSetTclErrnoFromWin32Error(WIN32_ERROR winerr)
{
    TWAPI_TCL85_INT_PLAT_STUB(tclWinConvertError) (winerr);
    return Tcl_GetErrno();
}

static int NPipeModuleInit(TwapiInterpContext *ticP)
{
    gNPipeTlsSlot = TwapiAssignTlsSlot();
    if (gNPipeTlsSlot < 0) {
        if (ticP && ticP->interp) {
            Tcl_SetResult(ticP->interp, "Could not assign private TLS slot", TCL_STATIC);
        }
        return TCL_ERROR;
    }

    return TCL_OK;
}


void NPipeThreadFinalize(void)
{
    NPipeTls *tlsP = GET_NPIPE_TLS();
    if (tlsP == NULL)
        return;

    //TBD - release all channels, events etc.
    // TBD - Tcl_DeleteEventSource(NPipeSetupProc, NPipeCheckProc, NULL);
        
    return;
}

/*
 * Invoked before Tcl_DoOneEvent blocks waiting for an event.
 * See Notifier man page in Tcl docs for expected behaviour.
 */
void NPipeSetupProc(
    ClientData data,		/* Not used. */
    int flags)			/* Event flags as passed to Tcl_DoOneEvent. */
{
    NPipeChannel *pcP;
    NPipeTls *tlsP;

    if (!(flags & TCL_FILE_EVENTS)) {
	return;
    }

    tlsP = GET_NPIPE_TLS();

    /*
     * Loop and check if there are any ready pipes.
     * The list is only accessed from this thread so no need to lock.
     */

    for (pcP = ZLIST_HEAD(&tlsP->pipes) ; pcP ; pcP = ZLIST_NEXT(pcP)) {
        if (NPIPE_NOTIFIABLE(pcP)) {
            Tcl_Time blockTime = { 0, 0 };
            /* Set block time to 0 so event loop will call us right away */
	    Tcl_SetMaxBlockTime(&blockTime);
	    break;
	}
    }
}



/*
 * Invoked by Tcl_DoOneEvent to allow queueing of ready events.
 * See Notifier man page in Tcl docs for expected behaviour.
 */
static void NPipeCheckProc(
    ClientData data,		/* Not used. */
    int flags)			/* Event flags as passed to Tcl_DoOneEvent. */
{
    NPipeChannel *pcP;
    NPipeTls *tlsP;
    NPipeEvent *evP;

    if (!(flags & TCL_FILE_EVENTS)) {
	return;
    }

    /*
     * Note: can use GET_NPIPE_TLS here because, this routine would not
     * be called without NPipeSetupProc being called first which guarantees
     * the TLS is set up.
     */
    tlsP = GET_NPIPE_TLS();

    /*
     * Loop and check if there are any ready pipes and queue events for
     * them. Note we do not queue events if we have already queued
     * one that has not been processed yet (NPIPE_F_EVENT_QUEUED)
     */

    for (pcP = ZLIST_HEAD(&tlsP->pipes) ; pcP ; pcP = ZLIST_NEXT(pcP)) {
        if ((((pcP->flags & NPIPE_F_WATCHREAD) && NPIPE_READ_NOTIFIABLE(pcP)) ||
             ((pcP->flags & NPIPE_F_WATCHWRITE) && NPIPE_WRITE_NOTIFIABLE(pcP))) &&
            !(pcP->flags & NPIPE_F_EVENT_QUEUED)) {
            /* Move pcP to front so event receiver will find it quicker */
            ZLIST_MOVETOHEAD(&tlsP->pipes, pcP);
	    evP = (NPipeEvent *) ckalloc(sizeof(*evP));
	    evP->header.proc = NPipeEventProc;
	    evP->hpipe = pcP->hpipe;
            /*
             * Note we do not protect the pcP from disappearing while
             * the event is on the event queue. The receiver will look
             * for it based on the pipe handle, and discard the event
             * if someone has deallocated in the meanwhile
             */
	    Tcl_QueueEvent(&evP->header, TCL_QUEUE_TAIL);
	}
    }
}

static NPipeTls *GetNPipeTls()
{
    NPipeTls *tlsP = GET_NPIPE_TLS();
    if (tlsP == NULL)
        return tlsP;

    tlsP = TwapiAlloc(sizeof(*tlsP));
    /* Initialize the critical section used for access to channel list.
     *
     * TBD - what's an appropriate spin count? Default of 0 is not desirable
     * As per MSDN, Windows heap manager uses 4000 so we do too.
     */
    InitializeCriticalSectionAndSpinCount(&tlsP->lock, 4000);
    ZLIST_INIT(&tlsP->pipes);

    //TBD - init tlsP;

    /* Store allocated TLS back into TLS slot */
    SET_NPIPE_TLS(tlsP);

    /* TBD - though convenient, this is really not the right place for this. */
    Tcl_CreateEventSource(NPipeSetupProc, NPipeCheckProc, NULL);

    return tlsP;
}

/* Always returns non-NULL, or panics */
static NPipeChannel *NPipeChannelNew(void)
{
    NPipeChannel *pcP;

    pcP = (NPipeChannel *) TwapiAlloc(sizeof(*pcP));
    pcP->thread = Tcl_GetCurrentThread();
    pcP->channel = NULL;
    pcP->hpipe = INVALID_HANDLE_VALUE;
    pcP->hsync = NULL;

    pcP->io[READER].hwait = INVALID_HANDLE_VALUE;
    pcP->io[READER].hevent = NULL;
    pcP->io[READER].state = IOBUF_IDLE;

    pcP->io[WRITER].hwait = INVALID_HANDLE_VALUE;
    pcP->io[WRITER].hevent = NULL;
    pcP->io[WRITER].state = IOBUF_IDLE;
    pcP->io[WRITER].data.write_buf.p = NULL;
    pcP->io[WRITER].data.write_buf.sz = 0;
    pcP->io[WRITER].data.write_buf.len = 0;

    pcP->flags = 0;

    pcP->winerr = ERROR_SUCCESS;
    pcP->nrefs = 0;
    ZLINK_INIT(pcP);
    return pcP;
}

void NPipeChannelUnref(NPipeChannel *pcP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized struct */
    if (InterlockedExchangeAdd(&pcP->nrefs, -decr) <= decr)
        NPipeChannelDelete(pcP);
}

/*
 * Called from thread pool when a overlapped read or write completes.
 * Note this must match the prototype typedef for WaitOrTimerCallback 
 */
static void NPipeThreadPoolHandler(
    NPipeChannel *pcP,
    int direction               /* READER or WRITER */
)
{
    LONG state;

    TWAPI_ASSERT(TimerOrWaitFired == FALSE);
    TWAPI_ASSERT(HasOverlappedIoCompleted(&pcP->io[direction].ovl));

    state = pcP->io[direction].ovl.Internal == ERROR_SUCCESS
        ? IOBUF_IO_COMPLETED : IOBUF_IO_COMPLETED_WITH_ERROR;

    /*
     * We only change state if it was IO_PENDING else someone has already
     * taken over the buffer
     */
    state = InterlockedCompareExchange(&pcP->io[direction].state,
                                       state, IOBUF_IO_PENDING);
    if (state == IOBUF_IO_PENDING) {
        /* Since we changed state, wake up the thread */
        Tcl_ThreadAlert(pcP->thread);
    }
}

/*
 * Called from thread pool when a overlapped read completes.
 * Note this must match the prototype typedef for WaitOrTimerCallback 
 */
static VOID CALLBACK NPipeReadThreadPoolFn(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
)
{
    NPipeThreadPoolHandler((NPipeChannel *) lpParameter, READER);
}

/*
 * Called from thread pool when a overlapped write completes.
 * Note this must match the prototype typedef for WaitOrTimerCallback 
 */
static VOID CALLBACK NPipeWriteThreadPoolFn(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
)
{
    NPipeThreadPoolHandler((NPipeChannel *) lpParameter, WRITER);
}

/*
 * Called from Tcl_ServiceEvent to process a pipe related event we
 * previously queued. See Notifier man page in Tcl docs for expected behaviour.
 */
static int NPipeEventProc(
    Tcl_Event *tcl_evP,		/* Event to service. */
    int flags)			/* Flags that indicate what events to handle,
				 * such as TCL_FILE_EVENTS. */
{
    NPipeEvent *evP = (NPipeEvent *)tcl_evP;
    NPipeChannel *pcP;
    NPipeTls *tlsP;
    int event_mask = 0;

    if (!(flags & TCL_FILE_EVENTS)) {
	return 0;               /* 0 -> Event will stay on queue */
    }

    /*
     * Note: can use GET_NPIPE_TLS here because, this routine would not
     * be called without NPipeSetupProc being called first which guarantees
     * the TLS is set up.
     */
    tlsP = GET_NPIPE_TLS();

    ZLIST_LOCATE(pcP, &tlsP->pipes, hpipe, evP->hpipe);
    if (pcP == NULL)
        return 1;               /* Stale event, pcP is gone */

    /* Indicate no events on queue so new events will be enqueued */
    pcP->flags &= ~ NPIPE_F_EVENT_QUEUED;

    if (! (pcP->flags & (NPIPE_F_WATCHWRITE | NPIPE_F_WATCHREAD)))
        return 1;               /* Not watching any reads or writes */


    /* Update channel state error if necessary */
    if (pcP->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) {
        SET_NPIPE_ERROR(pcP, pcP->io[READER].ovl.Internal);
    }
    if (pcP->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) {
        SET_NPIPE_ERROR(pcP, pcP->io[WRITER].ovl.Internal);
    }

    /* Now set the direction bits that are notifiable */
    if ((pcP->flags & NPIPE_F_WATCHREAD) && NPIPE_READ_NOTIFIABLE(pcP))
        event_mask |= TCL_READABLE;
    if ((pcP->flags & NPIPE_F_WATCHWRITE) && NPIPE_WRITE_NOTIFIABLE(pcP)) {
        event_mask |= TCL_WRITABLE;
        /* A write complete may actually be a connect complete */
        if (! NPIPE_CONNECTED(pcP)) {
            if (pcP->winerr == ERROR_SUCCESS) {
                pcP->flags |= NPIPE_F_CONNECTED;
                /* Also mark as readable on connect complete */
                if (pcP->flags & NPIPE_F_WATCHREAD)
                    event_mask |= TCL_READABLE;
            }
        }
    }
    /*
     * Mark write direction as idle if write completed successfully.
     * Note this includes a connection completion which also uses the WRITER
     * io struct.
     */
    InterlockedCompareExchange(&pcP->io[WRITER].state,
                               IOBUF_IO_COMPLETED,
                               IOBUF_IDLE);

    if (event_mask)
        Tcl_NotifyChannel(pcP->channel, event_mask);
    /*
     * Do not do anything more here because the Tcl_NotifyChannel may invoke
     * a callback which calls back into us and changes state, for example
     * even deallocating pcP.
     */
    return 1;
}

/* Sets up overlapped reads on a pipe. */
static void NPipeWatchReads(NPipeChannel *pcP)
{
    TWAPI_ASSERT(NPIPE_CONNECTED(pcP));

    /*
     * If channel has errored , ignore orders.
     * If a previous I/O has completed, ignore as well since we cannot
     * use the input buffer. The read completion handler will set up
     * the watch later.
     */
    if (pcP->winerr != ERROR_SUCCESS ||
        pcP->io[READER].state == IOBUF_IO_COMPLETED ||
        pcP->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) {
        /* Note we do not mark it here as a pipe error (pcP->winerr) */
        return;
    }

    /*
     * If a read is not already pending, initiate it.
     * We set up a single byte read so we're told when data is available.
     * We do not bother to special case when i/o completes immediately.
     * Let it also follow the async path for simplicity.
     */
    if (pcP->io[READER].state != IOBUF_IO_PENDING) {
        ZeroMemory(&pcP->io[READER].ovl, sizeof(pcP->io[READER].ovl));
        pcP->io[READER].ovl.hEvent = pcP->io[READER].hevent;
        if (! ReadFile(pcP->hpipe,
                       &pcP->io[READER].data.read_ahead,
                       1,
                       NULL,
                       &pcP->io[READER].ovl)) {
            WIN32_ERROR winerr = GetLastError();
            if (winerr != ERROR_IO_PENDING) {
                pcP->winerr = winerr;
                return;
            }
        }
    }

    /* Have the thread pool wait on it if not already doing so */
    if (pcP->io[READER].hwait == INVALID_HANDLE_VALUE) {
        /* The thread pool will hold a ref to the context */
        NPipeChannelRef(pcP, 1);
        if (! RegisterWaitForSingleObject(
                &pcP->io[READER].hwait,
                pcP->io[READER].hevent,
                NPipeReadThreadPoolFn,
                pcP,
                INFINITE,           /* No timeout */
                WT_EXECUTEDEFAULT
                )) {
            NPipeChannelUnref(pcP, 1);
            pcP->io[READER].hwait = INVALID_HANDLE_VALUE; /* Just in case... */
            /* Note the call does not set GetLastError. Make up our own */
            pcP->winerr = TWAPI_ERROR_TO_WIN32(TWAPI_REGISTER_WAIT_FAILED);
            return;
        }

    }
    pcP->io[READER].state = IOBUF_IO_PENDING;
    return;
}

static void NPipeDisableWatch(NPipeChannel *pcP, int direction)
{
    if (pcP->io[direction].hwait != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(pcP->io[direction].hwait, INVALID_HANDLE_VALUE);
        pcP->io[direction].hwait = INVALID_HANDLE_VALUE;
        NPipeChannelUnref(pcP, 1);
    }

    /*
     * Note io state might still be PENDING with an outstanding overlapped
     * operation that has not yet completed. The read/write routines
     * handle that appropriately.
     */
}

/* Called from Tcl I/O to indicate an interest in TCL_READABLE/TCL_WRITABLE */
static void NPipeWatchProc(ClientData clientdata, int mask)
{
    NPipeChannel *pcP = (NPipeChannel *)clientdata;

    if (mask & TCL_READABLE) {
        /* Only take action if we are not already watching reads */
        if (! (pcP->flags & NPIPE_F_WATCHREAD)) {
            pcP->flags |= NPIPE_F_WATCHREAD;
            /*
             * Set up the read only if we are connected.
             * If not connected, the connection complete code
             * will set up the read events.
             */
            if (NPIPE_CONNECTED(pcP)) {
                /*
                 * May set pcP errors. Those are automatically handled when
                 * we check for watchable events at function exit.
                 */
                NPipeWatchReads(pcP);
            }
        }
    } else {
        /* Not interested in reads, turn them off */
        pcP->flags &= ~ NPIPE_F_WATCHREAD;
        NPipeDisableWatch(pcP, READER);
    }

    /* Now do the write side in more or less identical fashion */
    if (mask & TCL_WRITABLE) {
        if (! (pcP->flags & NPIPE_F_WATCHWRITE)) {
            pcP->flags |= NPIPE_F_WATCHWRITE;
            /*
             * Note we do not need to set up the thread pool callback here.
             * Non-blocking pipes, it will already be set up. For blocking case,
             * write notifications will be generated after every write anyways.
             */
        }
    } else {
        /* Not interested in writes. Turn them off if currently on */
        if (pcP->flags & NPIPE_F_WATCHWRITE) {
            pcP->flags &= ~ NPIPE_F_WATCHWRITE;
            /*
             * Note we disable watching only on blocking pipes since
             * non-blocking pipes require the thread pool to be running
             */
            if ((pcP->flags & NPIPE_F_NONBLOCKING) == 0)
                NPipeDisableWatch(pcP, WRITER);
        }
    }


    /* Finally, if any watchable events, trigger them */
    if (((mask & TCL_READABLE) && NPIPE_READ_NOTIFIABLE(pcP)) ||
        ((mask & TCL_WRITABLE) && NPIPE_WRITE_NOTIFIABLE(pcP))) {
        /* Have Tcl event loop call us immediately to generate notifications */
        Tcl_Time blockTime = { 0, 0 };
        Tcl_SetMaxBlockTime(&blockTime);
    }
}

static TCL_RESULT NPipeCloseProc(ClientData clientdata, Tcl_Interp *interp)
{
    NPipeChannel *pcP = (NPipeChannel *) clientdata;

    /*
     * We need to unref pcP corresponding to the ref when we called
     * Tcl_CreateChannel. So we pass 1 to NPipeShutdown to do it for us.
     */
    NPipeShutdown(pcP, 1);

    return TCL_OK;
}


static int NPipeGetHandleProc(
    ClientData clientdata,
    int direction,		/* Not used. */
    ClientData *handlePtr)	/* Where to store the handle. */
{
    NPipeChannel *pcP = (NPipeChannel *)clientdata;

    *handlePtr = (ClientData) pcP->hpipe;
    return TCL_OK;
}

/* Called from Tcl I/O to indicate an interest in TCL_READABLE/TCL_WRITABLE */
static int NPipeBlockProc(ClientData clientdata, int mode)
{
    NPipeChannel *pcP = (NPipeChannel *)clientdata;

    if (mode == TCL_MODE_NONBLOCKING)
        pcP->flags |= NPIPE_F_NONBLOCKING;
    else
        pcP->flags &= ~ NPIPE_F_NONBLOCKING;

    return 0;                   /* POSIX error, 0 -> success */
}


/*
 * Read up to buf_sz bytes. Caller must have ensured there is at least one
 * byte else function will block waiting for data
 * (which is ok for blocking reads but not for nonblocking reads).
 * Returns number of bytes read on success, (at least one byte). Returns
 * 0 on error and sets pcP->winerr to error.
 */
static DWORD NPipeReadData(NPipeChannel *pcP, char *bufP, DWORD buf_sz)
{
    struct _NPipeIO *ioP = &pcP->io[READER];
    DWORD read_ahead_count = 0;
    DWORD avail;

    /* Must not be a pending read or error*/
    TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING && ioP->state != IOBUF_IO_COMPLETED_WITH_ERROR);
    /* Nonblocking pipes must have at least one byte data in read ahead */
    TWAPI_ASSERT((pcP->flags & NPIPE_F_NONBLOCKING) == 0 || ioP->state == IOBUF_IO_COMPLETED);
    TWAPI_ASSERT(ctxP->winerr == ERROR_SUCCESS);
    TWAPI_ASSERT(buf_sz > 0);

    if (ioP->state == IOBUF_IO_COMPLETED) {
        --buf_sz;
        ++read_ahead_count;
        *bufP++ = ioP->data.read_ahead[0];
        ioP->state = IOBUF_IDLE;
    }

    if (buf_sz == 0)
        return read_ahead_count;      /* Only asked for a single byte */

    if (!PeekNamedPipe(pcP->hpipe, NULL, 0, NULL, &avail, NULL)) {
        pcP->winerr = GetLastError();
        return 0;
    }

    if (avail == 0)
        avail = 1;              /* Wait for at least one byte */

    if (avail < buf_sz)
        buf_sz = avail;         /* Read what's there */
    
    ZeroMemory(&ioP->ovl, sizeof(ioP->ovl));
    ioP->hevent = pcP->hsync; /* Event used for "synchronous" reads */
    if (! ReadFile(pcP->hpipe, bufP, buf_sz, NULL, &ioP->ovl)) {
        WIN32_ERROR winerr = GetLastError();
        if (winerr != ERROR_IO_PENDING) {
            ioP->hevent = NULL;
            pcP->winerr = winerr;
            return 0;
        }
        /*
         * Pending I/O - when we have blocking reads and not enough data
         * in pipe. Fall thru, will block in GetOverlappedResult
         */
    }
    if (!GetOverlappedResult(ctxP->hpipe, &ioP->ovl, &avail, TRUE)) {
        ioP->hevent = NULL;
        pcP->winerr = GetLastError();
        return 0;
    }

    ioP->hevent = NULL;

    return avail + read_ahead_count;
}

/* Called from Tcl I/O channel layer to read bytes from the pipe */
static int NPipeInputProc(
    ClientData clientdata,
    char *bufP,
    int buf_sz,
    int *errnoP)
{
    NPipeChannel *pcP = (NPipeChannel *)clientdata;
    struct _NPipeIO *ioP;
    Tcl_Obj *objP;
    int read_count;

    TWAPI_ASSERT(pcP->thread == Tcl_GetCurrentThread());

    if (! NPIPE_CONNECTED(pcP)) {
        /* Note we do not set pcP->winerr here */
        *errnoP = NPipeSetTclErrnoFromWin32Error(ERROR_PIPE_NOT_CONNECTED);
        return -1;
    }

    ioP = &pcP->io[READER];

    /* If thread pool thread has set an error, we need to copy it. */
    if (ioP->state == IOBUF_IO_COMPLETED_WITH_ERROR)
        SET_NPIPE_ERROR(pcP, ioP->ovl.Internal);
    
    if (pcP->winerr != ERROR_SUCCESS)
        goto error_return;

    if (pcP->flags & PIPE_F_NONBLOCKING) {
        /* Non-blocking, we should always be either COMPLETED or PENDING */
        TWAPI_ASSERT(ioP->state == IOBUF_IO_COMPLETED || ioP->state == IOBUF_IO_PENDING);

        if (ioP->state == IOBUF_IO_PENDING) {
            /*
             * Non-blocking channel and we are awaiting data. We need to
             * return an EAGAIN or EWOULDBLOCK to Tcl. ERROR_PIPE_BUSY
             * maps to EAGAIN in the Tcl code.
             */
            *errnoP = NPipeSetTclErrnoFromWin32Error(ERROR_PIPE_BUSY);
            return -1;
        }

    } else {

        /* Blocking I/O */
        if (ioP->state == IOBUF_IO_PENDING) {
            /*
             * Blocking channel. Need to block waiting for data.
             * but there is an outstanding I/O pending. May be a thread pool
             * thread waiting on it via an overlapped read. Unregister it.
             * Might not be the most efficient mechanism but the combination
             * of fileevent with blocking channels is unlikely anyways.
             */
            if (ioP->hwait != INVALID_HANDLE_VALUE) {
                UnregisterWaitEx(pcP->io[READER].hwait, INVALID_HANDLE_VALUE);
                NPipeChannelUnref(pcP, 1);
                ioP->hwait = INVALID_HANDLE_VALUE;
            }
            /* Note ioP->state might have changed while we unregistered */
            /* If state is still pending, wait for outstanding read */
            if (ioP->state == IOBUF_IO_PENDING) {
                if (GetOverlappedResult(pcP->hpipe, &ioP->ovl, &read_count, TRUE))
                    ioP->state = IOBUF_IO_COMPLETED;
                else
                    ioP->state = IOBUF_IO_COMPLETED_WITH_ERROR;
            }
            if (ioP->state == IOBUF_IO_COMPLETED_WITH_ERROR) {
                SET_NPIPE_ERROR(pcP, ioP->ovl.Internal);
                goto error_return;
            }
        }
    }

    TWAPI_ASSERT(pcP->state != IOBUF_IO_PENDING);

    read_count = NPipeReadData(pcP, bufP, buf_sz);
    if (read_count == 0)
        goto error_return;

    /*
     * Need to fire another read event if more data is available. This
     * is also required in case we unregistered the thread pool wait
     * in the blocking I/O case above.
     */
    if (pcP->flags & PIPE_F_WATCHREAD)
        NPipeWatchReads(pcP);
    /*
     * Note error, if any, from TwapiPipeEnableReadWatch ignored,
     * will be picked up on next call
     */
    
    return read_count;

error_return:
    if (pcP->winerr == ERROR_HANDLE_EOF) {
        return 0;               /* EOF */
    } else {
        *errnoP = NPipeSetTclErrnoFromWin32Error(pcP->winerr);
        return -1;
    }        
}

/*
 * Initiates shut down of a pipe. 
 * It does NOT unregister the channel from the interp.
 * It does NOT remove the channel from the thread's pipe list.
 * Must be called from the thread that "owns" the channel to prevent races.
 * Note it does do an unref if unregistering from the thread pools so
 * to be sure that pcP is not deallocated on return, the caller must
 * ensure there is some other ref outstanding on it..
 * Caller can pass unrefs parameter as additional unrefs to do on pcP.
 */
static void NPipeShutdown(NPipeChannel *pcP, int unrefs)
{
    NPipeTls *tlsP = GET_NPIPE_TLS();

    TWAPI_ASSERT(pcP->thread == Tcl_GetCurrentThread());

    /*
     * stop the thread pool for this pipe. We need to do that before
     * closing handles. Note the UnregisterWaitEx can result in thread pool
     * callbacks running while it is blocked. That's ok because the thread
     * pool only changes the io state.
     */
    if (pcP->io[READER].hwait != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(pcP->io[READER].hwait,
                         INVALID_HANDLE_VALUE);
        pcP->io[READER].hwait = INVALID_HANDLE_VALUE;
        ++unrefs;   /* Remove the ref coming from the thread pool */
    }
    if (pcP->io[READER].hevent != NULL) {
        CloseHandle(pcP->io[READER].hevent);
        pcP->io[READER].hevent = NULL;
    }

    if (pcP->io[WRITER].hwait != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(pcP->io[WRITER].hwait,
                         INVALID_HANDLE_VALUE);
        pcP->io[WRITER].hwait = INVALID_HANDLE_VALUE;
        ++unrefs;   /* Remove the ref coming from the thread pool */
    }
    if (pcP->io[WRITER].hevent != NULL) {
        CloseHandle(pcP->io[WRITER].hevent);
        pcP->io[WRITER].hevent = NULL;
    }

    /* Third, now that handles are unregistered, close them. */
    if (pcP->hpipe != INVALID_HANDLE_VALUE) {
        CloseHandle(pcP->hpipe);
        pcP->hpipe = INVALID_HANDLE_VALUE;
    }

    if (pcP->hsync != NULL) {
        CloseHandle(pcP->hsync);
        pcP->hsync = NULL;
    }

    if (unrefs)
        NPipeChannelUnref(pcP, unrefs); /* May be GONE! */
}


static WIN32_ERROR NPipeAccept(NPipeChannel *pcP)
{
    WIN32_ERROR winerr;

    /* The writer side i/o structure are is also used for connecting */

    /*
     * Wait for client connection. If there is already a client waiting,
     * to connect, we will get success right away. Else we wait on the
     * event and queue a callback later
     */
    ZeroMemory(&pcP->io[WRITER].ovl, sizeof(pcP->io[WRITER].ovl));
    pcP->io[WRITER].ovl.hEvent = pcP->io[WRITER].hevent;
    if (ConnectNamedPipe(pcP->hpipe, &pcP->io[WRITER].ovl) ||
        (winerr = GetLastError()) == ERROR_PIPE_CONNECTED) {
        /* Already a client in waiting, queue the callback */
        pcP->flags |= NPIPE_F_CONNECTED;
        return ERROR_SUCCESS;
    } else if (winerr == ERROR_IO_PENDING) {
        /* Wait asynchronously for a connection */
        NPipeChannelRef(pcP, 1); /* Since we are passing to thread pool */
        if (RegisterWaitForSingleObject(
                &pcP->io[WRITER].hwait,
                pcP->io[WRITER].hevent,
                NPipeReadThreadPoolFn,
                pcP,
                INFINITE,           /* No timeout */
                WT_EXECUTEDEFAULT
                )) {
            pcP->io[WRITER].state = IOBUF_IO_PENDING;
            return ERROR_IO_PENDING;
        } else {
            winerr = GetLastError();
            NPipeChannelUnref(pcP, 1); /* Undo above Ref */
        }
    }

    /* Either synchronous completion or error */
    return winerr;
}

int Twapi_NPipeServer(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR name;
    DWORD open_mode, pipe_mode, max_instances;
    DWORD inbuf_sz, outbuf_sz, timeout;
    SECURITY_ATTRIBUTES *secattrP;
    NPipeChannel *pcP;
    DWORD winerr;
    Tcl_Interp *interp = ticP->interp;

    if (! TwapiDoOneTimeInit(&gNPipeModuleInitialized, NPipeModuleInit, ticP))
        return TCL_ERROR;

    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(name), GETINT(open_mode), GETINT(pipe_mode),
                     GETINT(max_instances), GETINT(outbuf_sz),
                     GETINT(inbuf_sz), GETINT(timeout),
                     GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES), ARGEND)
        != TCL_OK)
        return TCL_ERROR;

    if (pipe_mode & 0x7) {
        /* Currently, must be byte mode pipe and must not have NOWAIT flag */
        Tcl_SetResult(interp,  "Pipe mode must be byte mode and not specify the NPIPE_NOWAIT flag.", TCL_STATIC);
        return Twapi_AppendSystemError(interp, TWAPI_INVALID_ARGS);
    }

    pcP = NPipeChannelNew();

    open_mode |= FILE_FLAG_OVERLAPPED;
    pcP->hpipe = CreateNamedPipeW(name, open_mode, pipe_mode, max_instances,
                                  outbuf_sz, inbuf_sz, timeout, secattrP);

    if (pcP->hpipe != INVALID_HANDLE_VALUE) {
        /* Create event used for sync i/o. Note this is manual reset event */
        pcP->hsync = CreateEvent(NULL, TRUE, FALSE, NULL);

        if (pcP->hsync) {
            /* 
             * Create events to use for notification of completion. The
             * events must be auto-reset to prevent multiple callback
             * queueing on a single input notification. See MSDN docs for
             * RegisterWaitForSingleObject.  As a consequence, we must
             * make sure we never call GetOverlappedResult in blocking
             * mode when using one of these events (unlike hsync)
             */
            pcP->io[READER].hevent = CreateEvent(NULL, FALSE, FALSE, NULL);
            if (pcP->io[READER].hevent) {
                pcP->io[WRITER].hevent = CreateEvent(NULL, FALSE, FALSE, NULL);
                if (pcP->io[WRITER].hevent) {
                    int channel_mask = 0;
                    char instance_name[30];
                    NPipeTls *tlsP;

                    StringCbPrintfA(instance_name, sizeof(instance_name),
                                    "np%u", TWAPI_NEWID(ticP));
                    if (open_mode & PIPE_ACCESS_INBOUND)
                        channel_mask |= TCL_READABLE;
                    if (open_mode & PIPE_ACCESS_OUTBOUND)
                        channel_mask |= TCL_WRITABLE;
                    NPipeChannelRef(pcP, 1); /* Adding to Tcl channels */
                    pcP->channel = Tcl_CreateChannel(&gNPipeChannelDispatch,
                                                     instance_name, pcP,
                                                     channel_mask);
                    Tcl_SetChannelOption(interp, pcP->channel, "-encoding", "binary");
                    Tcl_SetChannelOption(interp, pcP->channel, "-translation", "auto crlf");
                    Tcl_SetChannelOption(NULL, pcP->channel, "-eofchar", "");
                    Tcl_RegisterChannel(interp, pcP->channel);

                    /*
                     * Note: Use GetNPipeTls, not GET_NPIPE_TLS here as tls
                     * might not have been initialized
                     */
                    tlsP = GetNPipeTls();

                    /*
                     * Add to list of pipes. It will be removed from the
                     * list when the thread action is called.
                     */
                    NPipeChannelRef(pcP, 1);
                    ZLIST_PREPEND(&tlsP->pipes, pcP);

                    /* Set up the accept */
                    winerr = NPipeAccept(pcP);
                    if (winerr == ERROR_SUCCESS || winerr == ERROR_IO_PENDING) {
                        /*
                         * On success (ie. immediate conn complete),
                         * we ask the event loop to call us back right away
                         * so we can generate the appropriate event.
                         */
                        if (winerr == ERROR_SUCCESS) {
                            Tcl_Time block_time = { 0, 0 };
                            Tcl_SetMaxBlockTime(&block_time);
                        }

                        /* Return channel name */
                        Tcl_SetObjResult(ticP->interp,
                                         Tcl_NewStringObj(instance_name, -1));
                        return TCL_OK;
                    } else {
                        /* Genuine error. */
                        Tcl_UnregisterChannel(interp, pcP->channel);
                        /* We do not NPipeChannelUnref here. That will
                         * happen when the Unregister calls our NPipeCloseProc
                         */
                        pcP->channel = NULL;
                    }
                } else
                    winerr = GetLastError();
            } else
                winerr = GetLastError();
        } else
            winerr = GetLastError();
    }

    pcP->winerr = winerr;
    NPipeShutdown(pcP, 0);    /* pcP might be gone */
    return Twapi_AppendSystemError(ticP->interp, winerr);
}



