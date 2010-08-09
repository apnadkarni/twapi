/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#if _WIN32_WINNT < 0x0500
#error _WIN32_WINNT too low
#endif

        
#define TWAPI_TCL_MAJOR 8
#ifdef _WIN64
#define TWAPI_MIN_TCL_MINOR 5
#else
#define TWAPI_MIN_TCL_MINOR 4
#endif

/*
 * Globals
 */
OSVERSIONINFO gTwapiOSVersionInfo;
GUID gTwapiNullGuid;             /* Initialized to all zeroes */
struct TwapiTclVersion gTclVersion;
int gTclIsThreaded;
static DWORD gTlsIndex;         /* As returned by TlsAlloc */
static ULONG volatile gTlsNextSlot;  /* Index into private slots in Tls area. */

/* List of allocated interpreter - used primarily for unnotified cleanup */
CRITICAL_SECTION gTwapiInterpContextsCS; /* To protect the same */
ZLIST_DECL(TwapiInterpContext) gTwapiInterpContexts;

/* Used to generate unique id's */
TwapiId volatile gIdGenerator;

/*
 * Whether the callback dll/libray has been initialized.
 * The value must be managed using the InterlockedCompareExchange functions to
 * ensure thread safety. The value returned by InterlockedCompareExhange
 * 0 -> first to call, do init,  1 -> init in progress by some other thread
 * 2 -> Init done
 */
static TwapiOneTimeInitState gTwapiInitialized;

static void Twapi_Cleanup(ClientData clientdata);
static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp);
static TwapiInterpContext *TwapiInterpContextNew(Tcl_Interp *interp);
static void TwapiInterpContextDelete(ticP, interp);
static int TwapiOneTimeInit(Tcl_Interp *interp);

/* Main entry point */
__declspec(dllexport) int Twapi_Init(Tcl_Interp *interp)
{
    static LONG twapi_initialized;
    TwapiInterpContext *ticP;
    /* Make sure Winsock is initialized */
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs - should this be the
       done for EVERY interp creation or move into one-time above ? TBD
     */
#ifdef USE_TCL_STUBS
    if (Tcl_InitStubs(interp, "8.4", 0) == NULL) {
        return TCL_ERROR;
    }
#endif


    /* Init unless already done. */
    if (! TwapiDoOneTimeInit(&gTwapiInitialized, TwapiOneTimeInit, interp))
        return TCL_ERROR;

    /* NOTE: no point setting Tcl_SetResult for errors as they are not
       looked at when DLL is being loaded */

    /*
     * Per interp initialization
     */

    /*
     * Create the name space and some variables. 
     * Needed for some scripts bound into the dll
     */
    Tcl_Eval(interp, "namespace eval " TWAPI_TCL_NAMESPACE " { variable settings ; set settings(log_limit) 100}");

    /* Allocate a context that will be passed around in all interpreters */
    ticP = TwapiInterpContextNew(interp);
    if (ticP == NULL)
        return TCL_ERROR;

    /* For all the commands we register with the Tcl interp, we add a single
     * ref for the context, not one per command. This is sufficient since
     * when the interp gets deleted, all the commands get deleted as well.
     * The corresponding Unref happens when the interp is deleted.
     *
     * In addition, we add one more ref because we will place it on the global
     * queue.
     */
    TwapiInterpContextRef(ticP, 1+1);
    EnterCriticalSection(&gTwapiInterpContextsCS);
    ZLIST_PREPEND(&gTwapiInterpContexts, ticP);
    LeaveCriticalSection(&gTwapiInterpContextsCS);


    Tcl_CallWhenDeleted(interp, Twapi_InterpCleanup, ticP);

    /* Do our own commands. */
    if (Twapi_InitCalls(interp, ticP) != TCL_OK) {
        return TCL_ERROR;
    }

    Tcl_CreateObjCommand(interp, "twapi::parseargs", Twapi_ParseargsObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::try", Twapi_TryObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::kl_get", Twapi_KlGetObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::twine", Twapi_TwineObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::recordarray", Twapi_RecordArrayObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::get_build_config",
                         Twapi_GetTwapiBuildInfo, ticP, NULL);
#ifndef TWAPI_NODESKTOP
    Tcl_CreateObjCommand(interp, "twapi::IDispatch_Invoke", Twapi_IDispatch_InvokeObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::ComEventSink", Twapi_ComEventSinkObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::SHChangeNotify", Twapi_SHChangeNotify,
                         ticP, NULL);
#endif

    return TCL_OK;
}


int
Twapi_GetTwapiBuildInfo(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    Tcl_Obj *objs[8];
    int i = 0;

    if (objc != 1)
        return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

#ifdef TWAPI_NODESKTOP
    objs[i++] = STRING_LITERAL_OBJ("nodesktop");
#endif
#ifdef TWAPI_NOSERVER
    objs[i++] = STRING_LITERAL_OBJ("noserver");
#endif
#ifdef TWAPI_LEAN
    objs[i++] = STRING_LITERAL_OBJ("lean");
#endif
#ifdef TWAPI_NOCALLBACKS
    objs[i++] = STRING_LITERAL_OBJ("nocallbacks");
#endif

    /* Tcl 8.4 has no indication of 32/64 builds at Tcl level so we have to. */
#ifdef _WIN64
    objs[i++] = STRING_LITERAL_OBJ("x64");
#else
    objs[i++] = STRING_LITERAL_OBJ("x86");
#endif    
    
    Tcl_SetObjResult(interp, Tcl_NewListObj(i, objs));
    return TCL_OK;
}

static void Twapi_Cleanup(ClientData clientdata)
{
    /* TBD - do we need to protect against more than one call ? */

    // Commented out CoUninitialize for the time being.
    // If there are event sinks in use, and the application exits
    // when the main window is closed, then Twapi_Cleanup gets
    // called BEFORE the window destroy. We then call CoUninitialize
    // and then if subsequently a COM call is made as part of the
    // Tk window destroy binding, we crash.
    // We want to do this last but there seems to be no way to
    // control order of initialization/finalization other than
    // the documented LIFO unloading of packages. That cannot
    // guarantee that Tk will not be loaded before us. So for now
    // we do not call this.
    // Note that Tk destroy binding runs as thread finalization
    // which happens AFTER process finalization (where we get called)
#if 0
    CoUninitialize();
#endif
    // TBD - clean up allocated interp context lists, threads etc.

    DeleteCriticalSection(&gTwapiInterpContextsCS);
    WSACleanup();
}


static TwapiInterpContext* TwapiInterpContextNew(Tcl_Interp *interp)
{
    DWORD winerr;
    TwapiInterpContext* ticP = TwapiAlloc(sizeof(*ticP));

    winerr = MemLifoInit(&ticP->memlifo, NULL, NULL, NULL, 8000,
                         MEMLIFO_F_PANIC_ON_FAIL);
    if (winerr != ERROR_SUCCESS) {
        Twapi_AppendSystemError(interp, winerr);
        return NULL;
    }

    ticP->nrefs = 0;
    ticP->interp = interp;
    ticP->thread = Tcl_GetCurrentThread();

    /* Initialize the critical section used for controlling
     * various attached lists
     *
     * TBD - what's an appropriate spin count? Default of 0 is not desirable
     * As per MSDN, Windows heap manager uses 4000 so we do too.
     */
    InitializeCriticalSectionAndSpinCount(&ticP->lock, 4000);

    ticP->pending_suspended = 0;
    ZLIST_INIT(&ticP->pending);
    ZLIST_INIT(&ticP->threadpool_registrations);
    ZLIST_INIT(&ticP->directory_monitors);

    /* Register a async callback with Tcl. */
    /* TBD - do we really need a separate Tcl_AsyncCreate call per interp?
     * or should it be per process ? Or per thread ? Do we need this at all?
     */
    ticP->async_handler = Tcl_AsyncCreate(Twapi_TclAsyncProc, ticP);

    ticP->notification_win = NULL; /* Created only on demand */
    ticP->clipboard_win = NULL;    /* Created only on demand */
    ticP->power_events_on = 0;
    ticP->console_ctrl_hooked = 0;
    ticP->device_notification_tid = 0;

    return ticP;
}

static void TwapiInterpContextDelete(TwapiInterpContext *ticP, Tcl_Interp *interp)
{
#ifdef TBD
    Enable under debug
    if (ticP->interp)
        Tcl_Panic("TwapiInterpContext deleted with active interp");
#endif

    /* TBD - does this need to be done only from the Tcl thread ? */
    if (ticP->async_handler)
        Tcl_AsyncDelete(ticP->async_handler);
    ticP->async_handler = 0;    /* Just in case */

    DeleteCriticalSection(&ticP->lock);

    /* TBD - should rest of this be in the Twapi_InterpCleanup instead ? */
    if (ticP->notification_win) {
        DestroyWindow(ticP->notification_win);
        ticP->notification_win = 0;
    }
    if (ticP->clipboard_win) {
        DestroyWindow(ticP->clipboard_win);
        ticP->clipboard_win = 0;
    }

    // TBD - what about pipes and directory_monitors ?

    MemLifoClose(&ticP->memlifo);

    // TBD - what about freeing the memory?
}

/* Decrement ref count and free if 0 */
void TwapiInterpContextUnref(TwapiInterpContext *ticP, int decr)
{
    if (InterlockedExchangeAdd(&ticP->nrefs, -decr) <= decr)
        TwapiInterpContextDelete(ticP);
}

static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp)
{
    TwapiInterpContext *tic2P;

    if (ticP->interp == NULL)
        return;


    EnterCriticalSection(&gTwapiInterpContextsCS);
    ZLIST_LOCATE(tic2P, &gTwapiInterpContexts, interp, ticP->interp);
    if (tic2P != ticP) {
        LeaveCriticalSection(&gTwapiInterpContextsCS);
        /* Either not found, or linked to a different interp. */
        Tcl_Panic("TWAPI interpreter context not found or attached to the wrong Tcl interpreter.");
    }
    ZLIST_REMOVE(&gTwapiInterpContexts, ticP);
    LeaveCriticalSection(&gTwapiInterpContextsCS);

    ticP->interp = NULL;        /* Must not access during cleanup */
    
    EnterCriticalSection(&ticP->lock);
    ticP->pending_suspended = 1;
    LeaveCriticalSection(&ticP->lock);

    if (ticP->console_ctrl_hooked)
        Twapi_StopConsoleEventNotifier(ticP);

    /* TBD - call other callback module clean up procedures */
    /* TBD - terminate device notification thread; */
    
    /* Unref for unlinking interp, +1 for removal from gTwapiInterpContexts */
    TwapiInterpContextUnref(ticP, 1+1);
}


/* One time (per process) initialization */
static int TwapiOneTimeInit(Tcl_Interp *interp)
{
    HRESULT hr;
    WSADATA ws_data;
    WORD    ws_ver = MAKEWORD(1,1);

    gTlsIndex = TlsAlloc();
    if (gTlsIndex == TLS_OUT_OF_INDEXES)
        return TCL_ERROR;       /* No point storing error message.
                                   Discarded anyways by Tcl */

    InitializeCriticalSection(&gTwapiInterpContextsCS);
    ZLIST_INIT(&gTwapiInterpContexts);

    if (Tcl_GetVar2Ex(interp, "tcl_platform", "threaded", TCL_GLOBAL_ONLY))
        gTclIsThreaded = 1;
    else
        gTclIsThreaded = 0;

    Tcl_GetVersion(&gTclVersion.major,
                   &gTclVersion.minor,
                   &gTclVersion.patchlevel,
                   &gTclVersion.reltype);
    if (gTclVersion.major ==  TWAPI_TCL_MAJOR &&
        gTclVersion.minor >= TWAPI_MIN_TCL_MINOR) {
        gTwapiOSVersionInfo.dwOSVersionInfoSize =
            sizeof(gTwapiOSVersionInfo);
        if (GetVersionEx(&gTwapiOSVersionInfo)) {
            /* Sockets */
            if (WSAStartup(ws_ver, &ws_data) == 0) {
                /* Single-threaded COM model - TBD */
                hr = CoInitializeEx(
                    NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
                if (hr == S_OK || hr == S_FALSE) {
                    /* All init worked. */
                    Tcl_CreateExitHandler(Twapi_Cleanup, NULL);
                    return TCL_OK;
                } else {
                    WSACleanup();
                }
            }
        }
    }

    return TCL_ERROR;
}

int TwapiAssignTlsSlot()
{
    DWORD slot;
    slot = InterlockedIncrement(&gTlsNextSlot);
    if (slot > TWAPI_TLS_SLOTS) {
        InterlockedDecrement(&gTlsNextSlot); /* So it does not grow unbounded */
        return -1;
    }
    return slot-1;
}

TwapiTls *TwapiGetTls()
{
    TwapiTls *tlsP;

    tlsP = (TwapiTls *) TlsGetValue(gTlsIndex);
    if (tlsP)
        return tlsP;

    /* TLS for this thread not initialized yet */
    tlsP = (TwapiTls *) TwapiAllocZero(sizeof(*tlsP));
    tlsP->thread = Tcl_GetCurrentThread();
    TlsSetValue(gTlsIndex, tlsP);
    return tlsP;
}


/* Note no locking necessary as only accessed from interp thread */
#define TwapiThreadPoolRegistrationRef(p_, incr_)    \
    do {(p_)->nrefs += (incr_);} while (0)

static void TwapiThreadPoolRegistrationUnref(TwapiThreadPoolRegistration *tprP,
                                             int unrefs)
{
    /* Note no locking necessary as only accessed from interp thread */
    tprP->nrefs -= unrefs;
    if (tprP->nrefs <= 0)
        TwapiFree(tprP);
}

static int TwapiThreadPoolRegistrationCallback(TwapiCallback *cbP)
{
    TwapiThreadPoolRegistration *tprP;
    TwapiInterpContext *ticP = cbP->ticP;
    HANDLE h;
    
    TWAPI_ASSERT(ticP);

    TwapiClearResult(&cbP->response); /* Not really necessary but for consistency */
    if (ticP->interp == NULL ||
        Tcl_InterpDeleted(ticP->interp)) {
        return TCL_ERROR;
    }

    h = (HANDLE)cbP->clientdata;
    ZLIST_LOCATE(tprP, &ticP->threadpool_registrations, handle, h);
    if (tprP == NULL) {
        return TCL_OK;                 /* Stale but ok */
    }

    /* Signal the event. Result is ignored */
    tprP->signal_handler(ticP, h, (DWORD) cbP->clientdata2);
    return TCL_OK;
}


/* Called from the thread pool when a handle is signalled */
static VOID CALLBACK TwapiThreadPoolRegistrationProc(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
    )
{
    TwapiThreadPoolRegistration *tprP =
        (TwapiThreadPoolRegistration *) lpParameter;
    TwapiCallback *cbP;

    cbP = TwapiCallbackNew(tprP->ticP,
                           TwapiThreadPoolRegistrationCallback,
                           sizeof(*cbP));

    /* Note we do not directly pass tprP. If we did would need to Ref it */
    cbP->clientdata = (DWORD_PTR) tprP->handle;
    cbP->clientdata2 = (DWORD_PTR) TimerOrWaitFired;
    cbP->winerr = ERROR_SUCCESS;
    TwapiEnqueueCallback(tprP->ticP, cbP,
                         TWAPI_ENQUEUE_DIRECT,
                         0, /* No response wanted */
                         NULL);
    /* TBD - on error, do we send an error notification ? */
}


void TwapiThreadPoolRegistrationShutdown(TwapiThreadPoolRegistration *tprP)
{
    int unrefs = 0;
    if (tprP->tp_handle != INVALID_HANDLE_VALUE) {
        if (! UnregisterWaitEx(tprP->tp_handle,
                               INVALID_HANDLE_VALUE /* Wait for callbacks to finish */
                )) {
            /*
             * TBD - how does one handle this ? Is it unregistered, was
             * never registered? or what ? At least log it.
             */
        }
        ++unrefs;           /* Since no longer referenced from thread pool */
    }

    /*
     * NOTE: Depending on the type of handle, NULL may or may not be
     * a valid handle. We always use INVALID_HANDLE_VALUE as a invalid
     * indicator.
     */
    if (tprP->unregistration_handler && tprP->handle != INVALID_HANDLE_VALUE)
        tprP->unregistration_handler(tprP->ticP, tprP->handle);
    tprP->handle = INVALID_HANDLE_VALUE;

    tprP->tp_handle = INVALID_HANDLE_VALUE;

    if (tprP->ticP) {
        ZLIST_REMOVE(&tprP->ticP->threadpool_registrations, tprP);
        TwapiInterpContextUnref(tprP->ticP, 1); /*  May be gone! */
        tprP->ticP = NULL;
        ++unrefs;                         /* Not referenced from list */
    }

    if (unrefs)
        TwapiThreadPoolRegistrationUnref(tprP, unrefs); /* May be gone! */
}


WIN32_ERROR TwapiThreadPoolRegister(
    TwapiInterpContext *ticP,
    HANDLE h,
    ULONG wait_ms,
    DWORD  flags,
    void (*signal_handler)(TwapiInterpContext *ticP, HANDLE, DWORD),
    void (*unregistration_handler)(TwapiInterpContext *ticP, HANDLE)
    )
{
    TwapiThreadPoolRegistration *tprP = TwapiAlloc(sizeof(*tprP));

    tprP->handle = h;
    tprP->signal_handler = signal_handler;
    tprP->unregistration_handler = unregistration_handler;

    /* Only certain flags are obeyed. */
    flags &= WT_EXECUTEONLYONCE;

    flags |= WT_EXECUTEDEFAULT;

    /*
     * Note once registered with thread pool call back might run even
     * before the registration call returns so set everything up
     * before the call
     */
    ZLIST_PREPEND(&ticP->threadpool_registrations, tprP);
    tprP->ticP = ticP;
    TwapiInterpContextRef(ticP, 1);
    /* One ref for list linkage, one for handing off to thread pool */
    TwapiThreadPoolRegistrationRef(tprP, 2);
    if (RegisterWaitForSingleObject(&tprP->tp_handle,
                                    h,
                                    TwapiThreadPoolRegistrationProc,
                                    tprP,
                                    wait_ms,
                                    flags)) {
        return ERROR_SUCCESS;
    } else {
        tprP->tp_handle = INVALID_HANDLE_VALUE; /* Just to be sure */
        /* Back out the ref for thread pool since it failed */
        TwapiThreadPoolRegistrationUnref(tprP, 1);

        TwapiThreadPoolRegistrationShutdown(tprP);
        return TWAPI_ERROR_TO_WIN32(TWAPI_REGISTER_WAIT_FAILED);
    }
}
    
                                       
void TwapiThreadPoolUnregister(
    TwapiInterpContext *ticP,
    HANDLE h
    )
{
    TwapiThreadPoolRegistration *tprP;
    
    ZLIST_LOCATE(tprP, &ticP->threadpool_registrations, handle, h);
    if (tprP == NULL)
        return;                 /* Stale? */

    TWAPI_ASSERT(ticP == tprP->ticP);

    TwapiThreadPoolRegistrationShutdown(tprP);
}


/* The callback that invokes the user level script for async handle waits */
void TwapiCallRegisteredWaitScript(TwapiInterpContext *ticP, HANDLE h, DWORD timeout)
{
    Tcl_Obj *objs[3];
    int i;

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_wait_handler", -1);
    objs[1] = ObjFromHANDLE(h);
    if (timeout) 
        objs[2] = STRING_LITERAL_OBJ("timeout");
    else
        objs[2] = STRING_LITERAL_OBJ("signalled");

    for (i = 0; i < ARRAYSIZE(objs); ++i) {
        Tcl_IncrRefCount(objs[i]);
    }
    Tcl_EvalObjv(ticP->interp, ARRAYSIZE(objs), objs, 
                 TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);

    for (i = 0; i < ARRAYSIZE(objs); ++i) {
        Tcl_DecrRefCount(objs[i]);
    }
}

