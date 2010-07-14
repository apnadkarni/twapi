/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"


typedef struct _TwapiPipeBuffer {
    OVERLAPPED ovl;
    int        buf_sz;          /* Actual size of buf[] */
    char       buf[1];          /* Variable sized */
} TwapiPipeBuffer;

typedef struct _TwapiPipeContext {
    TwapiInterpContext *ticP;
    HANDLE  hpipe;   /* Handle to the pipe */
    struct {
        HANDLE hwait;       /* Handle returned by thread pool wait functions */
        HANDLE hevent;      /* Used for read i/o signaling */
        TwapiPipeBuffer *iobP; /* Used for overlapped i/o */
    } io[2];                   /* 0 -> read, 1 -> write */
#define READER 0
#define WRITER 1
    ZLINK_DECL(TwapiPipeContext);
    ULONG volatile nrefs;              /* Ref count */
    /* VARIABLE SIZE AREA FOLLOWS */
} TwapiPipeContext;


/*
 * Prototypes
 */
static TwapiPipeContext *TwapiPipeContextNew(TwapiInterpContext *ticP);
static void TwapiPipeContextDelete(TwapiPipeContext *ctxP);
#define TwapiPipeContextRef(p_, incr_) InterlockedExchangeAdd(&(p_)->nrefs, (incr_))
void TwapiPipeContextUnref(TwapiPipeContext *ctxP, int decr);
static int TwapiPipeShutdown(TwapiPipeContext *ctxP);

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

    ctxP = TwapiPipeContextNew(ticP);

    ctxP->hpipe = CreateNamedPipeW(name, open_mode, pipe_mode, max_instances,
                                   outbuf_sz, inbuf_sz, timeout, secattrP);

    if (ctxP->hpipe != INVALID_HANDLE_VALUE) {
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
                
                Tcl_SetObjResult(ticP->interp, ObjFromHANDLE(ctxP->hpipe));
                return TCL_OK;
            }
            winerr = GetLastError(); /* Get error before calling CloseHandle */
            CloseHandle(ctxP->io[READER].hevent);
        } else
            winerr = GetLastError();

        CloseHandle(ctxP->hpipe);
    }

    Twapi_AppendSystemError(ticP->interp, winerr);
    TwapiPipeContextDelete(ctxP);
    return TCL_ERROR;
}


int Twapi_PipeClose(TwapiInterpContext *ticP, HANDLE hpipe)
{
    TwapiPipeContext *ctxP;

    TWAPI_ASSERT(ticP->thread == Tcl_GetCurrentThread());
    
    ZLIST_LOCATE(ctxP, &ticP->pipes, hpipe, hpipe);
    if (ctxP == NULL)
        return TwapiReturnTwapiError(ticP->interp, NULL, TWAPI_INVALID_ARGS);

    TWAPI_ASSERT(ticP == ctxP->ticP);

    TwapiPipeShutdown(ctxP);

    return TCL_OK;
}



/* Always returns non-NULL, or panics */
static TwapiPipeContext *TwapiPipeContextNew(TwapiInterpContext *ticP)
{
    int i;
    TwapiPipeContext *ctxP;

    ctxP = (TwapiPipeContext *) TwapiAlloc(sizeof(*ctxP));
    ctxP->ticP = ticP;
    if (ticP)
        TwapiInterpContextRef(ticP, 1);
    ctxP->hpipe = INVALID_HANDLE_VALUE;

    for (i = 0; i < ARRAYSIZE(ctxP->io); ++i) {
        ctxP->io[i].hwait = INVALID_HANDLE_VALUE;
        ctxP->io[i].hevent = NULL;
        ctxP->io[i].iobP = NULL;
    }
    ctxP->nrefs = 0;
    ZLINK_INIT(ctxP);
    return ctxP;
}

static void TwapiPipeContextDelete(TwapiPipeContext *ctxP)
{
    int i;

    TWAPI_ASSERT(ctxP->nrefs <= 0);
    TWAPI_ASSERT(ctxP->io[READER].hpool == NULL);
    TWAPI_ASSERT(ctxP->io[READER].hevent == NULL);
    TWAPI_ASSERT(ctxP->io[WRITER].hpool == NULL);
    TWAPI_ASSERT(ctxP->io[WRITER].hevent == NULL);
    TWAPI_ASSERT(ctxP->hpipe == INVALID_HANDLE_VALUE);

    for (i = 0; i < ARRAYSIZE(ctxP->io); ++i) {
        if (ctxP->io[i].iobP)
            TwapiFree(ctxP->io[i].iobP);
    }

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
     * additional events to the interp thread. That's ok because we unlinked
     * ctxP->ticP above.
     */
    for (i = 0; i < ARRAYSIZE(ctxP->io); ++i) {
        if (ctxP->io[i].hwait != INVALID_HANDLE_VALUE) {
            UnregisterWaitEx(ctxP->io[i].hwait,
                             INVALID_HANDLE_VALUE);
            ctxP->io[i].hwait = INVALID_HANDLE_VALUE;
            ++unrefs;   /* Remove the ref coming from the thread pool */
        }
        if (ctxP->io[i].hevent != NULL) {
            CloseHandle(ctxP->io[i].hevent);
            ctxP->io[i].hevent = NULL;
        }

    }

    /* Third, now that handles are unregistered, close them. */
    if (ctxP->hpipe != INVALID_HANDLE_VALUE)
        CloseHandle(ctxP->hpipe);

    if (unrefs)
        TwapiPipeContextUnref(ctxP, unrefs); /* May be GONE! */

    return TCL_OK;
}
