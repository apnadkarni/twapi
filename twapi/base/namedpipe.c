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
    Tcl_ThreadId thread;    /* The thread owning the channel. If non-NULL,
                               this structure MUST be linked on the PipeTls
                               of the corresponding thread. */
    Tcl_Channel channel;           /* The corresponding Tcl_Channel handle */
    HANDLE  hpipe;   /* Handle to the pipe */
    HANDLE  hsync;   /* Event used for synchronous I/O, both read and write.
                        Needed because we cannot use the hevent field
                        for sync i/o as the thread pool will also be waiting
                        on it */
    struct _NPipeIO {
        OVERLAPPED ovl;         /* Used in async i/o. Note the ovl.hEvent field
                                   is always set at init time */
        HANDLE hwait;       /* Handle returned by thread pool wait functions. */
        HANDLE hevent;      /* Event for thread pool to wait on */
        union {
            char read_ahead[1];     /* Used to read-ahead a single byte */
            struct {
                char *p;     /* Data to write out */
                int   sz;    /* Size of buffer (not length of data) */
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
#define NPIPE_F_EOF_NOTIFIED   32 /* Have already notified EOF */

    ULONG volatile nrefs;              /* Ref count */
    WIN32_ERROR winerr;

#define SET_NPIPE_ERROR(ctxP_, err_) \
    ((ctxP_)->winerr == ERROR_SUCCESS ? ((ctxP_)->winerr = (err_)) : (ctxP_)->winerr)
};

#define NPIPE_CONNECTED(pcP_) ((pcP_)->flags & NPIPE_F_CONNECTED)

/*
 * EOF errors - note 0xc000014b is defined as STATUS_PIPE_BROKEN in
 * ntstatus.h. However that duplicates some defs from winnt.h and
 * hence cannot be included without generating compiler warnings.
 */
#define NPIPE_EOF(pcP_) \
    ((pcP_)->winerr == ERROR_HANDLE_EOF ||      \
     (pcP)->winerr == ERROR_BROKEN_PIPE ||      \
     (pcP)->winerr == 0xc000014b)

#define NPIPE_EOF_NOTIFIABLE(pcP_) \
    (NPIPE_EOF(pcP_) && !((pcP_)->flags & NPIPE_F_EOF_NOTIFIED))

/*
 * When should we notify for a read/write - data i/o completed
 * Note we do not include IDLE state as then we would continually generate
 * (potentially) notifications.
 * Errors are only notified in connecting stage ( is that correct ?)
 */
#define NPIPE_READ_NOTIFIABLE(pcP_) \
    (((pcP_)->io[READER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->winerr != ERROR_SUCCESS && !NPIPE_CONNECTED(pcP_)))

#define NPIPE_WRITE_NOTIFIABLE(pcP_) \
    (((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->winerr != ERROR_SUCCESS && !NPIPE_CONNECTED(pcP_)))

/* Combination of above */
#define NPIPE_NOTIFIABLE(pcP_) \
    (((pcP_)->io[READER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     NPIPE_EOF_NOTIFIABLE(pcP_) ||                                                 \
     ((pcP_)->winerr != ERROR_SUCCESS && !NPIPE_CONNECTED(pcP_)))

/* Is an async i/o pending ? */
#define NPIPE_READ_PENDING(pcP_) ((pcP_)->io[READER].state == IOBUF_IO_PENDING)
#define NPIPE_WRITE_PENDING(pcP_) ((pcP_)->io[WRITER].state == IOBUF_IO_PENDING)


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
static void NPipeShutdown(Tcl_Interp *interp, NPipeChannel *pcP, int unrefs);

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
     * one that has not been processed yet (NPIPE_F_EVENT_QUEUED).
     * Note, we do not take shortcuts for the case where no watches
     * are set on the channel since NPipeEventProc also does appropriate
     * state changes in addition to notifying the channel subsystem
     * (see bug 3245925)
     * TBD - can we not move the EVENT_QUEUED check to NPipeSetupProc
     * and save an unnecessary callback into here ?
     */

    for (pcP = ZLIST_HEAD(&tlsP->pipes) ; pcP ; pcP = ZLIST_NEXT(pcP)) {
        if (NPIPE_NOTIFIABLE(pcP) &&
            !(pcP->flags & NPIPE_F_EVENT_QUEUED)) {
            /* Move pcP to front so event receiver will find it quicker */
            ZLIST_MOVETOHEAD(&tlsP->pipes, pcP);
	    evP = (NPipeEvent *) ckalloc(sizeof(*evP));
	    evP->header.proc = NPipeEventProc;
	    evP->hpipe = pcP->hpipe;
            /* Indicate event queued so no new events will be enqueued */
            pcP->flags |= NPIPE_F_EVENT_QUEUED;
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
    if (tlsP != NULL)
        return tlsP;

    tlsP = TwapiAlloc(sizeof(*tlsP));
    /* Initialize the critical section used for access to channel list.
     *
     * TBD - what's an appropriate spin count? Default of 0 is not desirable
     * As per MSDN, Windows heap manager uses 4000 so we do too.
     */
    InitializeCriticalSectionAndSpinCount(&tlsP->lock, 4000);
    ZLIST_INIT(&tlsP->pipes);

    //TBD - any other init tlsP?

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
    pcP->thread = NULL;
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

    pcP->flags = 0;

    pcP->winerr = ERROR_SUCCESS;
    pcP->nrefs = 0;
    ZLINK_INIT(pcP);
    return pcP;
}

static void NPipeChannelDelete(NPipeChannel *pcP)
{
    TWAPI_ASSERT(pcP->thread == NULL);
    TWAPI_ASSERT(pcP->nrefs <= 0);
    TWAPI_ASSERT(pcP->io[READER].hwait == INVALID_HANDLE_VALUE);
    TWAPI_ASSERT(pcP->io[READER].hevent == NULL);
    TWAPI_ASSERT(pcP->io[WRITER].hwait == INVALID_HANDLE_VALUE);
    TWAPI_ASSERT(pcP->io[WRITER].hevent == NULL);
    TWAPI_ASSERT(pcP->hsync == NULL);
    TWAPI_ASSERT(pcP->hpipe == INVALID_HANDLE_VALUE);

    if (pcP->io[WRITER].data.write_buf.p)
        TwapiFree(pcP->io[WRITER].data.write_buf.p);

    TwapiFree(pcP);
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

#if 0
WRONG - there is something to do -> change the io buf state on writes
    /*
     * If we are connected but not watching any reads or writes, 
     * nothing to do.
     */
    if (NPIPE_CONNECTED(pcP) &&
        ! (pcP->flags & (NPIPE_F_WATCHWRITE | NPIPE_F_WATCHREAD))) {
        return 1;               /* Not watching any reads or writes */
    }
#endif

    /* Update channel state error if necessary */
    if (pcP->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) {
        SET_NPIPE_ERROR(pcP, (WIN32_ERROR) pcP->io[READER].ovl.Internal);
    }
    if (pcP->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) {
        SET_NPIPE_ERROR(pcP, (WIN32_ERROR) pcP->io[WRITER].ovl.Internal);
    }

    /* Now set the direction bits that are notifiable */
    if ((pcP->flags & NPIPE_F_WATCHREAD) && NPIPE_READ_NOTIFIABLE(pcP)) {
        event_mask |= TCL_READABLE;
    }

    if (NPIPE_WRITE_NOTIFIABLE(pcP)) {
        /* This might be a connect complete */
        if (! NPIPE_CONNECTED(pcP)) {
            pcP->flags |= NPIPE_F_CONNECTED;
            /* Also mark as readable on connect complete */
            if (pcP->flags & NPIPE_F_WATCHREAD)
                event_mask |= TCL_READABLE;
        }
        if (pcP->flags & NPIPE_F_WATCHWRITE)
            event_mask |= TCL_WRITABLE;
    }

    /* On EOF, both read and write notification are set */
    if (NPIPE_EOF_NOTIFIABLE(pcP)) {
        if (pcP->flags & NPIPE_F_WATCHWRITE)
            event_mask |= TCL_WRITABLE;
        if (pcP->flags & NPIPE_F_WATCHREAD)
            event_mask |= TCL_READABLE;
        /* Make sure we do not keep generating EOF notifications */
        pcP->flags |= NPIPE_F_EOF_NOTIFIED;
    }

    /*
     * Mark write direction as idle if write completed successfully.
     * Note this includes a connection completion which also uses the WRITER
     * io struct.
     */
    InterlockedCompareExchange(&pcP->io[WRITER].state,
                               IOBUF_IDLE,
                               IOBUF_IO_COMPLETED);

    if (event_mask) {
        Tcl_NotifyChannel(pcP->channel, event_mask);
    }
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
        pcP->io[READER].state = IOBUF_IO_PENDING; /* Important to set this first since reader thread might run immediately */
        TwapiZeroMemory(&pcP->io[READER].ovl, sizeof(pcP->io[READER].ovl));
        pcP->io[READER].ovl.hEvent = pcP->io[READER].hevent;
        if (! ReadFile(pcP->hpipe,
                       &pcP->io[READER].data.read_ahead,
                       1,
                       NULL,
                       &pcP->io[READER].ovl)) {
            WIN32_ERROR winerr = GetLastError();
            if (winerr != ERROR_IO_PENDING) {
                pcP->io[READER].state = IOBUF_IDLE;
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
    if (mask & (TCL_READABLE|TCL_WRITABLE)) {
        if (NPIPE_EOF_NOTIFIABLE(pcP) ||
            ((mask & TCL_READABLE) && NPIPE_READ_NOTIFIABLE(pcP)) ||
            ((mask & TCL_WRITABLE) && NPIPE_WRITE_NOTIFIABLE(pcP))) {
            /* Have Tcl event loop call us immediately to generate notifications */
            Tcl_Time blockTime = { 0, 0 };
            Tcl_SetMaxBlockTime(&blockTime);
        }
    }
}

static TCL_RESULT NPipeCloseProc(ClientData clientdata, Tcl_Interp *interp)
{
    NPipeChannel *pcP = (NPipeChannel *) clientdata;

    TWAPI_ASSERT(pcP->thread == NULL);

    /*
     * We need to unref pcP corresponding to the ref when we called
     * Tcl_CreateChannel. So we pass 1 to NPipeShutdown to do it for us.
     */
    NPipeShutdown(interp, pcP, 1);

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
 * Read up to buf_sz bytes. Returns number of bytes read. For non-blocking
 * channels, if no data available, returns 0. For blocking channels, will
 * block until at least one byte is available. In case of errors, returns
 * 0. Caller can distinguish between "no data" for non-blocking channels
 * and error by checking pcP->winerr.
 */
static DWORD NPipeReadData(NPipeChannel *pcP, char *bufP, DWORD bufsz)
{
    struct _NPipeIO *ioP = &pcP->io[READER];
    DWORD overlap_count;      /* Count returned in async completion */
    DWORD sync_count;         /* Count returned in sync completion */
    DWORD num_to_read;        /* How much to try and read */
    DWORD nread = 0;          /* Actually read */

    /* Must not be a pending read or error*/
    TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING && ioP->state != IOBUF_IO_COMPLETED_WITH_ERROR);
    TWAPI_ASSERT(bufsz > 0);

    if (ioP->state == IOBUF_IO_COMPLETED) {
        --bufsz;
        ++nread;
        *bufP++ = ioP->data.read_ahead[0];
        ioP->state = IOBUF_IDLE;
    }

    if (bufsz == 0)
        return nread;      /* Only asked for a single byte */

    if (!PeekNamedPipe(pcP->hpipe, NULL, 0, NULL, &num_to_read, NULL)) {
        pcP->winerr = GetLastError();
        return 0;
    }

    if (num_to_read == 0) {
        if (pcP->flags & NPIPE_F_NONBLOCKING) {
            /* Non-blocking. Return what we have, if anything */
            return nread; /* May be 0 */
        }
        num_to_read = 1;        /* Blocking chan - Wait for at least one byte */
    }

    if (num_to_read > bufsz)
        num_to_read = bufsz;
    
    TwapiZeroMemory(&ioP->ovl, sizeof(ioP->ovl));
    ioP->ovl.hEvent = pcP->hsync; /* Event used for "synchronous" reads */
    if (ReadFile(pcP->hpipe, bufP, num_to_read, &sync_count, &ioP->ovl)) {
        /*
         * Synchronous completion. Note must not use GetOverlappedResult
         * (see http://support.microsoft.com/kb/156932)
         */
        nread += sync_count;
    } else {
        WIN32_ERROR winerr = GetLastError();
        if (winerr != ERROR_IO_PENDING) {
            ioP->ovl.hEvent = NULL;
            pcP->winerr = winerr;
            return 0;
        }
        /*
         * Pending I/O - when we have blocking reads and not enough data
         * in pipe. Fall thru, will block in GetOverlappedResult
         */
        if (!GetOverlappedResult(pcP->hpipe, &ioP->ovl, &overlap_count, TRUE)) {
            ioP->ovl.hEvent = NULL;
            pcP->winerr = GetLastError();
            return 0;
        }
        nread += overlap_count;
    }

    /* TBD - if more data is available in pipe, get it as well if there is room */

    ioP->ovl.hEvent = NULL;

    return nread;
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
    int nread;

    TWAPI_ASSERT(pcP->thread == Tcl_GetCurrentThread());

    if (! NPIPE_CONNECTED(pcP)) {
        /* Note we do not set pcP->winerr here */
        *errnoP = NPipeSetTclErrnoFromWin32Error(ERROR_PIPE_NOT_CONNECTED);
        return -1;
    }

    ioP = &pcP->io[READER];

    /* If thread pool thread has set an error, we need to copy it. */
    if (ioP->state == IOBUF_IO_COMPLETED_WITH_ERROR)
        SET_NPIPE_ERROR(pcP, (WIN32_ERROR) ioP->ovl.Internal);
    
    if (pcP->winerr != ERROR_SUCCESS)
        goto error_return;

    if (pcP->flags & NPIPE_F_NONBLOCKING) {
        /* Non-blocking, we should always be either COMPLETED or PENDING */
#if 0
        TBD This assert fails as state is IDLE. I think that is legit but verify
        TWAPI_ASSERT(ioP->state == IOBUF_IO_COMPLETED || ioP->state == IOBUF_IO_PENDING);
#endif
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
                if (GetOverlappedResult(pcP->hpipe, &ioP->ovl, &nread, TRUE))
                    ioP->state = IOBUF_IO_COMPLETED;
                else
                    ioP->state = IOBUF_IO_COMPLETED_WITH_ERROR;
            }
            if (ioP->state == IOBUF_IO_COMPLETED_WITH_ERROR) {
                SET_NPIPE_ERROR(pcP, (WIN32_ERROR) ioP->ovl.Internal);
                goto error_return;
            }
        }
    }

    TWAPI_ASSERT(ioP->state != IOBUF_IO_PENDING);

    nread = NPipeReadData(pcP, bufP, buf_sz);
    /* Note nread may be 0 for non-blocking pipes and no data */
    if (pcP->winerr != ERROR_SUCCESS)
        goto error_return;

    /*
     * Need to fire another read event if more data is available. This
     * is also required in case we unregistered the thread pool wait
     * in the blocking I/O case above.
     */
    if (pcP->flags & NPIPE_F_WATCHREAD)
        NPipeWatchReads(pcP);
    /*
     * Note error, if any, from TwapiPipeEnableReadWatch ignored,
     * will be picked up on next call
     */

    return nread;

error_return:
    if (NPIPE_EOF(pcP)) {
        return 0;               /* EOF */
    } else {
        *errnoP = NPipeSetTclErrnoFromWin32Error(pcP->winerr);
        return -1;
    }        
}


/* Called from Tcl I/O channel layer to write bytes from the pipe */
static int NPipeOutputProc(
    ClientData clientdata,
    const char *bufP,
    int count,
    int *errnoP)
{
    NPipeChannel *pcP = (NPipeChannel *)clientdata;
    struct _NPipeIO *ioP;

    TWAPI_ASSERT(pcP->thread == Tcl_GetCurrentThread());

    if (! NPIPE_CONNECTED(pcP)) {
        /* Note we do not set pcP->winerr here */
        *errnoP = NPipeSetTclErrnoFromWin32Error(ERROR_PIPE_NOT_CONNECTED);
        return -1;
    }

    ioP = &pcP->io[WRITER];

    /* If thread pool thread has set an error, we need to copy it. */
    if (ioP->state == IOBUF_IO_COMPLETED_WITH_ERROR)
        SET_NPIPE_ERROR(pcP, (WIN32_ERROR) ioP->ovl.Internal);
    
    if (pcP->winerr != ERROR_SUCCESS)
        goto error_return;
    
    if (pcP->flags & NPIPE_F_NONBLOCKING) {
        /* Non-blocking */
        if (ioP->state == IOBUF_IO_PENDING) {
            /*
             * Non-blocking channel and we are awaiting data. We need to
             * return an EAGAIN or EWOULDBLOCK to Tcl. ERROR_PIPE_BUSY
             * maps to EAGAIN in the Tcl code.
             */
            *errnoP = NPipeSetTclErrnoFromWin32Error(ERROR_PIPE_BUSY);
            return -1;
        }

        TwapiZeroMemory(&ioP->ovl, sizeof(ioP->ovl));
        ioP->ovl.hEvent = ioP->hevent;

        /*
         * Reallocate buffer if necessary. Note as as aside, that once
         * allocated, we do not deallocate except to grow the buffer.
         */
        if (ioP->data.write_buf.sz < count) {
            size_t actual_sz;
            if (ioP->data.write_buf.p != NULL)
                TwapiFree(ioP->data.write_buf.p);
            ioP->data.write_buf.p = TwapiAllocSize(count, &actual_sz);
            ioP->data.write_buf.sz = (DWORD) actual_sz;
        }
        CopyMemory(ioP->data.write_buf.p, bufP, count);
        ioP->state = IOBUF_IO_PENDING;
        if (WriteFile(pcP->hpipe, ioP->data.write_buf.p, count, NULL, &ioP->ovl)) {
            /*
             * Operation completed right away. Double check to make
             * sure, and if so change state. The thread pool callback
             * will run any way but will basically be a no-op since
             * we change state here. Note the thread pool might have
             * run BEFORE us as well (immediately after the WriteFile above)
             * hence the interlocked check that state is still PENDING.
             * We do this so caller will not needlessly get EAGAIN if he tries
             * writing right away. This is striclty for efficiency.
             */
            InterlockedCompareExchange(&ioP->state,
                                       IOBUF_IO_COMPLETED, IOBUF_IO_PENDING);
            /* Note state is changed to COMPLETED, not IDLE, so that
               file event notification will be generated if necessary */
        } else {
            /* WriteFile did not complete successfully right away */
            WIN32_ERROR winerr = GetLastError();
            if (winerr != ERROR_IO_PENDING) {
                /* Genuine error */
                pcP->winerr = winerr;
                goto error_return;
            }

            /* Have the thread pool wait on it if not already doing so */
            if (ioP->hwait == INVALID_HANDLE_VALUE) {
                /* The thread pool will hold a ref to the context */
                NPipeChannelRef(pcP, 1);
                if (! RegisterWaitForSingleObject(
                        &ioP->hwait,
                        ioP->hevent,
                        NPipeWriteThreadPoolFn,
                        pcP,
                        INFINITE,           /* No timeout */
                        WT_EXECUTEDEFAULT
                        )) {
                    NPipeChannelUnref(pcP, 1);
                    ioP->hwait = INVALID_HANDLE_VALUE; /* Just in case... */
                    /* Note the call does not set GetLastError. Make up our own */
                    pcP->winerr = TWAPI_ERROR_TO_WIN32(TWAPI_REGISTER_WAIT_FAILED);
                    goto error_return;
                }
            }
        }
    } else {
        /*
         * Blocking I/O
         * We do not use the buffer in ioP since we will block for completion
         * anyway and might as well save the copy. In addition, we do not
         * need to check for IO_PENDING state. It's OK if there was
         * a previous async write on this pipe that has not yet completed.
         * The kernel will queue this sync write behind it. We just
         * make sure we use a different OVERLAPPED struct and event.
         */
        OVERLAPPED ovl;
        TwapiZeroMemory(&ovl, sizeof(ovl));
        ovl.hEvent = pcP->hsync;
        if (! WriteFile(pcP->hpipe, bufP, count, NULL, &ovl)) {
            WIN32_ERROR winerr = GetLastError();
            if (winerr != ERROR_IO_PENDING) {
                /* Genuine error */
                pcP->winerr = winerr;
                goto error_return;
            }
            /* Wait for I/O to complete */
            if (!GetOverlappedResult(pcP->hpipe, &ovl, &count, TRUE)) {
                pcP->winerr = GetLastError();
                goto error_return;
            }
        }
        /*
         * Sync I/O completed. Change state only if previous state was
         * IDLE. If previous state was PENDING from a previous non-blocking
         * call, the async callback will change state. If it was
         * COMPLETED_WITH_ERROR we do not want to change state.
         */
        InterlockedCompareExchange(&ioP->state,
                                   IOBUF_IO_COMPLETED, IOBUF_IDLE);
    }

    /* As far as caller concerned, all bytes written in all non-error cases */
    return count;

error_return:
    *errnoP = NPipeSetTclErrnoFromWin32Error(pcP->winerr);
    return -1;
}

/* Called from Tcl I/O layer to add/remove a channel from a thread */
static void NPipeThreadActionProc(
    ClientData clientdata,
    int action)
{
    NPipeChannel *pcP = (NPipeChannel *) clientdata;
    NPipeTls *tlsP = GET_NPIPE_TLS();
    Tcl_ThreadId tid;
    
    tid = Tcl_GetCurrentThread();
    if (action == TCL_CHANNEL_THREAD_INSERT) {
        /* TBD - if this gets called on creation, remove the
           assign to pcP->thread in Tcl_PipeServer */
        TWAPI_ASSERT(pcP->thread == NULL || pcP->thread == tid);
        if (pcP->thread != tid) {
            pcP->thread = tid;
            NPipeChannelRef(pcP, 1);
            ZLIST_PREPEND(&tlsP->pipes, pcP);
        }
    } else {
        TWAPI_ASSERT(pcP->thread == tid);
        pcP->thread = NULL;
        ZLIST_REMOVE(&tlsP->pipes, pcP);
        /* The Tcl_CreateChannel ref is expected to remain else Tcl channel
         * will be holding an invalid pointer if the unref deallocates
         */
        TWAPI_ASSERT(pcP->nrefs > 1);
        NPipeChannelUnref(pcP, 1);
    }
}

/*
 * Initiates shut down of a pipe. 
 * It does NOT unregister the channel from the interp.
 * It does NOT remove the channel from the thread's pipe list.
 * Must be called from the thread that "owns" the channel to prevent races.
 * Note it DOES do an unref if unregistering from the thread pools so
 * to be sure that pcP is not deallocated on return, the caller must
 * ensure there is some other ref outstanding on it..
 * Caller can pass unrefs parameter as additional unrefs to do on pcP.
 */
static void NPipeShutdown(Tcl_Interp *interp, NPipeChannel *pcP, int unrefs)
{
    NPipeTls *tlsP = GET_NPIPE_TLS();

    TWAPI_ASSERT(pcP->thread == NULL || pcP->thread == Tcl_GetCurrentThread());

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

    if (pcP->io[WRITER].hwait != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(pcP->io[WRITER].hwait,
                         INVALID_HANDLE_VALUE);
        pcP->io[WRITER].hwait = INVALID_HANDLE_VALUE;
        ++unrefs;   /* Remove the ref coming from the thread pool */
    }

    /* Third, now that handles are unregistered, close them. */
    if (pcP->hpipe != INVALID_HANDLE_VALUE) {
        CloseHandle(pcP->hpipe);
        pcP->hpipe = INVALID_HANDLE_VALUE;
    }

    if (pcP->io[WRITER].hevent != NULL) {
        /*
         * If there is any pending I/O, we have to wait or it to finish
         * even though we closed the channel.
         */
        if (NPIPE_WRITE_PENDING(pcP) && pcP->io[WRITER].ovl.hEvent == pcP->io[WRITER].hevent) {
            if (WaitForSingleObject(pcP->io[WRITER].hevent, 1000) != WAIT_OBJECT_0) {
                if (interp) {
                    Twapi_AppendLog(interp, L"WaitForSingleObject did not return WAIT_OBJECT_0 while shutting named pipe (WRITER)");
                }
            }
        }
        CloseHandle(pcP->io[WRITER].hevent);
        pcP->io[WRITER].ovl.hEvent = NULL;
        pcP->io[WRITER].hevent = NULL;
    }

    if (pcP->io[READER].hevent != NULL) {
        /* See WRITER comments above */
        if (NPIPE_READ_PENDING(pcP) && pcP->io[READER].ovl.hEvent == pcP->io[READER].hevent) {
            if (WaitForSingleObject(pcP->io[READER].hevent, 1000) != WAIT_OBJECT_0) {
                if (interp) {
                    Twapi_AppendLog(interp, L"WaitForSingleObject did not return WAIT_OBJECT_0 while shutting named pipe (READER)");
                }
            }
        }
        CloseHandle(pcP->io[READER].hevent);
        pcP->io[READER].hevent = NULL;
        pcP->io[READER].ovl.hEvent = NULL;
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
    TwapiZeroMemory(&pcP->io[WRITER].ovl, sizeof(pcP->io[WRITER].ovl));
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
                NPipeWriteThreadPoolFn,
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

/*
 * For consistency with sockets, we configure the same options
 * as channel defaults.
 */
static void NPipeConfigureChannelDefaults(Tcl_Interp *interp, NPipeChannel *pcP)
{
    Tcl_SetChannelOption(interp, pcP->channel, "-translation", "auto crlf");
    Tcl_SetChannelOption(NULL, pcP->channel, "-eofchar", "");
}

int Twapi_NPipeServer(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR name;
    DWORD open_mode, pipe_mode, max_instances;
    DWORD inbuf_sz, outbuf_sz, timeout;
    SECURITY_ATTRIBUTES *secattrP = NULL;
    NPipeChannel *pcP;
    DWORD winerr = ERROR_SUCCESS;
    Tcl_Interp *interp = ticP->interp;
    NPipeTls *tlsP;

    if (! TwapiDoOneTimeInit(&gNPipeModuleInitialized, NPipeModuleInit, ticP))
        return TCL_ERROR;

    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(name), GETINT(open_mode), GETINT(pipe_mode),
                     GETINT(max_instances), GETINT(outbuf_sz),
                     GETINT(inbuf_sz), GETINT(timeout),
                     GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES), ARGEND)
        != TCL_OK)
        return TCL_ERROR;

    /*
     * Note: Use GetNPipeTls, not GET_NPIPE_TLS here as tls
     * might not have been initialized. Also do this as the first thing
     * as various callbacks when registering channels will call functions
     * which expect the tls to have been initialized.
     */
    tlsP = GetNPipeTls();

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
                    wsprintf(instance_name, "np%u", TWAPI_NEWID(ticP));
                    if (open_mode & PIPE_ACCESS_INBOUND)
                        channel_mask |= TCL_READABLE;
                    if (open_mode & PIPE_ACCESS_OUTBOUND)
                        channel_mask |= TCL_WRITABLE;
                    NPipeChannelRef(pcP, 1); /* Adding to Tcl channels */
                    pcP->channel = Tcl_CreateChannel(&gNPipeChannelDispatch,
                                                     instance_name, pcP,
                                                     channel_mask);
                    /*
                     * Note the CreateChannel will call back into our
                     * ThreadActionProc which would have added pcP to
                     * the thread tls
                     */

                    NPipeConfigureChannelDefaults(interp, pcP);
                    Tcl_RegisterChannel(interp, pcP->channel);

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
                }
            }
        }
    }

    /* Only init winerr if not already done */
    if (winerr == ERROR_SUCCESS)
        winerr = GetLastError();

    pcP->winerr = winerr;
    NPipeShutdown(ticP->interp, pcP, 0);    /* pcP might be gone */
    return Twapi_AppendSystemError(ticP->interp, winerr);
}


int Twapi_NPipeClient(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR name;
    DWORD desired_access, share_mode, creation_disposition;
    DWORD flags_attr;
    SECURITY_ATTRIBUTES *secattrP = NULL;
    NPipeChannel *pcP;
    DWORD winerr;
    Tcl_Interp *interp = ticP->interp;
    NPipeTls *tlsP;
    HANDLE hpipe;

    if (! TwapiDoOneTimeInit(&gNPipeModuleInitialized, NPipeModuleInit, ticP))
        return TCL_ERROR;

    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(name),
                     GETINT(desired_access),
                     GETINT(share_mode),
                     GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                     GETINT(creation_disposition),
                     GETINT(flags_attr),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    /*
     * Note: Use GetNPipeTls, not GET_NPIPE_TLS here as tls
     * might not have been initialized. Also do this as the first thing
     * as various callbacks when registering channels will call functions
     * which expect the tls to have been initialized.
     */
    tlsP = GetNPipeTls();

    flags_attr |= FILE_FLAG_OVERLAPPED;
    hpipe = CreateFileW(name, desired_access,
                        share_mode, secattrP,
                        creation_disposition,
                        flags_attr, NULL);
    if (hpipe == INVALID_HANDLE_VALUE)
        return TwapiReturnSystemError(interp);

    if (GetFileType(hpipe) != FILE_TYPE_PIPE) {
        CloseHandle(hpipe);
        return Twapi_GenerateWin32Error(interp, ERROR_INVALID_NAME,
                                        "Specified path is not a pipe.");
    }


    pcP = NPipeChannelNew();

    pcP->hpipe = hpipe;

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
                wsprintf(instance_name, "np%u", TWAPI_NEWID(ticP));
                if (desired_access & (GENERIC_READ |FILE_READ_DATA))
                    channel_mask |= TCL_READABLE;
                if (desired_access & (GENERIC_WRITE|FILE_WRITE_DATA))
                    channel_mask |= TCL_WRITABLE;
                NPipeChannelRef(pcP, 1); /* Adding to Tcl channels */
                pcP->channel = Tcl_CreateChannel(&gNPipeChannelDispatch,
                                                 instance_name, pcP,
                                                 channel_mask);
                /*
                 * Note the CreateChannel will call back into our
                 * ThreadActionProc which would have added pcP to
                 * the thread tls
                 */

                NPipeConfigureChannelDefaults(interp, pcP);

                Tcl_RegisterChannel(interp, pcP->channel);
                
                pcP->flags |= NPIPE_F_CONNECTED;
                /* Return channel name */
                Tcl_SetObjResult(ticP->interp,
                                 Tcl_NewStringObj(instance_name, -1));
                return TCL_OK;
            }
        }
    }

    winerr = GetLastError();
    pcP->winerr = winerr;
    NPipeShutdown(ticP->interp, pcP, 0);    /* pcP might be gone */
    return Twapi_AppendSystemError(ticP->interp, winerr);
}



