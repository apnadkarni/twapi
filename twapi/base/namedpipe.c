/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

typedef struct _PipeChannel PipeChannel;
ZLINK_CREATE_TYPEDEFS(PipeChannel); 
ZLIST_CREATE_TYPEDEFS(PipeChannel);

typedef struct _PipeChannel {
    ZLINK_DECL(PipeChannel); /* List of registered pipes. Access sync
                                   through global pipe channel lock */
    Tcl_ThreadId thread;    /* The thread owning the channel */
    Tcl_Channel channel;           /* The corresponding Tcl_Channel handle */
    HANDLE  hpipe;   /* Handle to the pipe */
    HANDLE  hsync;   /* Event used for synchronous I/O, both read and write.
                        Needed because we cannot use the hevent field
                        for sync i/o as the thread pool will also be waiting
                        on it */
    struct _PipeIO {
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
#define PIPE_F_WATCHREAD       1 /* Generate event when data available */
#define PIPE_F_WATCHWRITE      2 /* Generate event when output possible */
#define PIPE_F_NONBLOCKING     4 /* Channel is async */
#define PIPE_F_CONNECTED       8 /* Client has successfully connected */
#define PIPE_F_EVENT_QUEUED   16 /* A TCL event has been queued */ 

    ULONG volatile nrefs;              /* Ref count */
    WIN32_ERROR winerr;

#define SET_CONTEXT_WINERR(ctxP_, err_) \
    ((ctxP_)->winerr == ERROR_SUCCESS ? ((ctxP_)->winerr = (err_)) : (ctxP_)->winerr)
};

#define PIPE_CONNECTED(pcP_) ((pcP_)->flags & PIPE_F_CONNECTED)
/*
 * When should we notify for a read - data available in buffer or error/eof.
 * Errors other than eof are only notified in connecting stage.
 */
#define PIPE_READ_NOTIFIABLE(pcP_) \
    (((pcP_)->io[READER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->winerr == ERROR_HANDLE_EOF) ||                            \
     ((pcP_)->winerr != ERROR_SUCCESS && !PIPE_CONNECTED(pcP_)))

/*
 * When should we notify for a write - data i/o completed, or error/eof
 * Note we do not include IDLE state as then we would continually generate
 * (potentially) notifications.
 * Errors other than eof are only notified in connecting stage.
 */
#define PIPE_WRITE_NOTIFIABLE(pcP_) \
    (((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED) ||                \
     ((pcP_)->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) ||     \
     ((pcP_)->winerr == ERROR_HANDLE_EOF) ||                            \
     ((pcP_)->winerr != ERROR_SUCCESS && !PIPE_CONNECTED(pcP_)))

/*
 * Pipe channels are maintained on a per thread basis. 
 */
typedef struct _PipeTls {
    /*
     * Create list header definitions for list of pipes.
     */
    CRITICAL_SECTION lock;
    ZLIST_DECL(PipeChannel) pipes;
    
} PipeTls;

static int gPipeTlsSlot;
/* Use only when sure that module and thread init has been done */
#define GET_PIPE_TLS() ((PipeTls *) TWAPI_TLS_SLOT(gPipeTlsSlot))
#define SET_PIPE_TLS(p_) do {TWAPI_TLS_SLOT(gPipeTlsSlot) = (DWORD_PTR) (p_);} while (0)

typedef struct _PipeEvent {
    Tcl_Event header;
    HANDLE hpipe;               /* Pipe to which this event relates */
} PipeChannelEvent;


static TwapiOneTimeInitState gPipeModuleInitialized;

/* Prototypes */
static Tcl_DriverCloseProc PipeCloseProc;
static Tcl_DriverInputProc PipeInputProc;
static Tcl_DriverOutputProc PipeOutputProc;
static Tcl_DriverWatchProc PipeWatchProc;
static Tcl_DriverGetHandleProc PipeGetHandleProc;
static Tcl_DriverBlockModeProc PipeBlockProc;
static Tcl_DriverThreadActionProc PipeThreadActionProc;

static Tcl_ChannelType gPipeChannelDispatch = {
    "namedpipe",
    (Tcl_ChannelTypeVersion)TCL_CHANNEL_VERSION_2,
    PipeCloseProc,
    PipeInputProc,
    PipeOutputProc,
    NULL /* ChannelSeek */,
    NULL /* ChannelSetOption */,
    NULL /* ChannelGetOption */,
    PipeWatchProc,
    PipeGetHandleProc,
    NULL /* ChannelClose2 */,
    PipeBlockProc,
    NULL /* ChannelFlush */,
    NULL /* ChannelHandler */,
    NULL /* ChannelWideSeek if VERSION_2, Truncate if VERSION_5 */
    PipeThreadActionProc,       /* Unused for VERSION_2 */
    NULL,               /* TruncateProc for VERSION_5, unused for VERSION_2 */
};

static PipeChannel *PipeChannelNew(void);
static void PipeChannelDelete(PipeChannel *pcP);
#define PipeChannelRef(p_, incr_) InterlockedExchangeAdd(&(p_)->nrefs, (incr_))
void PipeChannelUnref(PipeChannel *pcP, int decr);

static int PipeModuleInit(TwapiInterpContext *ticP)
{
    gPipeTlsSlot = TwapiAssignTlsSlot();
    if (gPipeTlsSlot < 0) {
        if (ticP && ticP->interp) {
            Tcl_SetResult(ticP->interp, "Could not assign private TLS slot", TCL_STATIC);
        }
        return TCL_ERROR;
    }

    return TCL_OK;
}


static PipeTls *GetPipeTls()
{
    PipeTls *tlsP = GET_PIPE_TLS();
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
    SET_PIPE_TLS(tlsP);

    return tlsP;
}

void
PipeThreadFinalize(void)
{
    PipeTls *tlsP = GET_PIPE_TLS();
    if (tlsP == NULL)
        return;

    //TBD - release all channels, events etc.
    // TBD - Tcl_DeleteEventSource(PipeSetupProc, PipeCheckProc, NULL);
        
    return;
}

/*
 * Invoked before Tcl_DoOneEvent blocks waiting for an event.
 * See Notifier man page in Tcl docs for expected behaviour.
 */
void PipeSetupProc(
    ClientData data,		/* Not used. */
    int flags)			/* Event flags as passed to Tcl_DoOneEvent. */
{
    PipeChannel *pcP;
    PipeTls *tlsP;

    if (!(flags & TCL_FILE_EVENTS)) {
	return;
    }

    /*
     * Note cannot use GET_PIPE_TLS because no guarantee this is not the
     * first call to get the tls for this thread. TBD - check this
     */
    tlsP = GetPipeTls();

    /*
     * Loop and check if there are any ready pipes.
     * The list is only accessed from this thread so no need to lock.
     */

    for (pcP = ZLIST_HEAD(&tlsP->pipes) ; pcP ; pcP = ZLIST_NEXT(pcP)) {
        if (PIPE_HAS_EVENTS(pcP)) {
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
static void PipeCheckProc(
    ClientData data,		/* Not used. */
    int flags)			/* Event flags as passed to Tcl_DoOneEvent. */
{
    PipeChannel *pcP;
    PipeTls *tlsP;
    PipeEvent *evP;

    if (!(flags & TCL_FILE_EVENTS)) {
	return;
    }

    /*
     * Note: can use GET_PIPE_TLS here because, this routine would not
     * be called without PipeSetupProc being called first which guarantees
     * the TLS is set up.
     */
    tlsP = GET_PIPE_TLS();

    /*
     * Loop and check if there are any ready pipes and queue events for
     * them. Note we do not queue events if we have already queued
     * one that has not been processed yet (PIPE_F_EVENT_QUEUED)
     */

    for (pcP = ZLIST_HEAD(&tlsP->pipes) ; pcP ; pcP = ZLIST_NEXT(pcP)) {
        if ((((pcP->flags & PIPE_F_WATCHREAD) && PIPE_READ_NOTIFIABLE(pcP)) ||
             ((pcP->flags & PIPE_F_WATCHWRITE) && PIPE_WRITE_NOTIFIABLE(pcP))) &&
            !(pcP->flags & PIPE_F_EVENT_QUEUED)) {
            /* Move pcP to front so event receiver will find it quicker */
            ZLIST_MOVETOHEAD(&tlsP->pipes, pcP);
	    evP = (PipeEvent *) ckalloc(sizeof(SocketEvent));
	    evP->header.proc = PipeEventProc;
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

/*
 * Called from Tcl_ServiceEvent to process a pipe related event we
 * previously queued. See Notifier man page in Tcl docs for expected behaviour.
 */

static int PipeEventProc(
    Tcl_Event *evP,		/* Event to service. */
    int flags)			/* Flags that indicate what events to handle,
				 * such as TCL_FILE_EVENTS. */
{
    PipeChannel *pcP;
    PipeTls *tlsP;
    int event_mask = 0;

    if (!(flags & TCL_FILE_EVENTS)) {
	return 0;               /* 0 -> Event will stay on queue */
    }

    /*
     * Note: can use GET_PIPE_TLS here because, this routine would not
     * be called without PipeSetupProc being called first which guarantees
     * the TLS is set up.
     */
    tlsP = GET_PIPE_TLS();

    ZLIST_LOCATE(pcP, &tlsP->pipes, hpipe, evP->hpipe);
    if (pcP == NULL)
        return 1;               /* Stale event, pcP is gone */

    /* Indicate no events on queue so new events will be enqueued */
    pcP->flags &= ~ PIPE_F_EVENT_QUEUED;

    if (! (pcP->flags & (PIPE_F_WATCHWRITE | PIPE_F_WATCHREAD)))
        return 1;               /* Not watching any reads or writes */


    /* Update channel state error if necessary */
    if (pcP->io[READER].state == IOBUF_IO_COMPLETED_WITH_ERROR) {
        SET_CONTEXT_WINERR(pcP, pcP->io[READER].ovl.Internal);
    }
    if (pcP->io[WRITER].state == IOBUF_IO_COMPLETED_WITH_ERROR) {
        SET_CONTEXT_WINERR(pcP, pcP->io[WRITER].ovl.Internal);
    }

    /* Now set the direction bits that are notifiable */
    if ((pcP->flags & PIPE_F_WATCHREAD) && PIPE_READ_NOTIFIABLE(pcP))
        event_mask |= TCL_READABLE;
    if ((pcP->flags & PIPE_F_WATCHWRITE) && PIPE_WRITE_NOTIFIABLE(pcP)) {
        event_mask |= TCL_WRITABLE;
        /* A write complete may actually be a connect complete */
        if (! PIPE_CONNECTED(pcP)) {
            if (pcP->winerr == ERROR_SUCCESS) {
                pcP->flags |= PIPE_F_CONNECTED;
                /* Also mark as readable on connect complete */
                if (pcP->flags & PIPE_F_WATCHREAD)
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
                               IOBUF_IO_IDLE);

    if (event_mask)
        Tcl_NotifyChannel(pcP->channel, event_mask);
    /*
     * Do not do anything more here because the Tcl_NotifyChannel may invoke
     * a callback which calls back into us and changes state, for example
     * even deallocating pcP.
     */
    return 1;
}



#ifdef NOTYET

static void PipeChannelRegister(PipeChannel *pcP)
{
    PipeChannelRef(pcP, 1);
    EnterCriticalSection(&gPipeChannelLock);
    ZLIST_PREPEND(&gPipeChannels, pcP);
    LeaveCriticalSection(&gPipeChannelLock);
}

static void PipeChannelUnregister(PipeChannel *pcP)
{
    EnterCriticalSection(&gPipeChannelLock);
    ZLIST_REMOVE(&gPipeChannels, pcP);
    LeaveCriticalSection(&gPipeChannelLock);
    PipeChannelUnref(pcP, 1);
}

/* Always returns non-NULL, or panics */
static PipeChannel *PipeChannelNew(void)
{
    PipeChannel *pcP;

    pcP = (PipeChannel *) TwapiAlloc(sizeof(*pcP));
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


int Twapi_PipeServer(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR name;
    DWORD open_mode, pipe_mode, max_instances;
    DWORD inbuf_sz, outbuf_sz, timeout;
    SECURITY_ATTRIBUTES *secattrP;
    PipeChannel *pcP;
    DWORD winerr;
    Tcl_Interp *interp = ticP->interp;

    if (! TwapiDoOneTimeInit(&gPipeModuleInitialized, PipeChannelModuleInit, ticP))
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
        Tcl_SetResult(interp,  "Pipe mode must be byte mode and not specify the PIPE_NOWAIT flag.", TCL_STATIC);
        return Twapi_AppendSystemError(interp, TWAPI_INVALID_ARGS);
    }

    pcP = PipeChannelNew();

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
                    StringCbPrintfA(instance_name, sizeof(instance_name),
                                    "np#%u", InterlockedIncrement(&gPipeChannelId));
                    if (open_mode & PIPE_ACCESS_INBOUND)
                        channel_mask |= TCL_READABLE;
                    if (open_mode & PIPE_ACCESS_OUTBOUND)
                        channel_mask |= TCL_WRITABLE;
                    pcP->channel = Tcl_CreateChannel(&gPipeChannelDispatch,
                                                     instance_name, pcP,
                                                     channel_mask);
                    Tcl_SetChannelOption(interp, pcP->channel, "-encoding", "binary");
                    Tcl_SetChannelOption(interp, pcP->channel, "-translation", "binary");
                    Tcl_RegisterChannel(interp, pcP->channel);

                    /* Add to list of pipes */
                    PipeChannelRegister(pcP);
                    Tcl_SetObjResult(ticP->interp,
                                     Tcl_NewStringObj(instance_name, -1));
                    return TCL_OK;
                } else
                    winerr = GetLastError();
            } else
                winerr = GetLastError();
        } else
            winerr = GetLastError();
    }

    PipeShutdown(pcP);    /* pcP might be gone */
    return Twapi_AppendSystemError(ticP->interp, winerr);
}


void PipeChannelUnref(PipeChannel *pcP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized struct */
    if (InterlockedExchangeAdd(&pcP->nrefs, -decr) <= decr)
        PipeChannelDelete(pcP);
}


TCL_RESULT PipeChannelClose(ClientData clientdata, Tcl_Interp *interp)
{
    PipeChannel *pcP = (PipeChannel *) clientdata;
    Tcl_Channel channel = pcP->channel;

    EnterCriticalSection(&gPipeChannelLock);
    ZLIST_LOCATE(&gPipeChannels, pcP, channel, channel);
    if (pcP == NULL || clientdata != (ClientData) pcP) {
        /* TBD - log not found or mismatch */
        LeaveCriticalSection(&gPipeChannelLock);
        return TCL_OK;
    }
    pcP->channel = NULL;
    ZLIST_REMOVE(&gPipeChannels, pcP);
    LeaveCriticalSection(&gPipeChannelLock);

    PipeShutdown(pcP);

    return TCL_OK;
}
#endif // NOTYET
