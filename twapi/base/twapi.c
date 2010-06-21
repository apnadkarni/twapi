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

/*
 * Globals
 */
OSVERSIONINFO TwapiOSVersionInfo;
GUID TwapiNullGuid;             /* Initialized to all zeroes */
struct TwapiTclVersion TclVersion;
int TclIsThreaded;


static void Twapi_Cleanup(ClientData clientdata);
static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp);
static TwapiInterpContext *TwapiInterpContextNew(Tcl_Interp *interp);
static void TwapiInterpContextDelete(ticP, interp);

/* Main entry point */
__declspec(dllexport) int Twapi_Init(Tcl_Interp *interp)
{
    int result;
    static LONG twapi_initialized;
    TwapiInterpContext *ticP;

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
#ifdef USE_TCL_STUBS
    if (Tcl_InitStubs(interp, "8.4", 0) == NULL) {
        return TCL_ERROR;
    }
#endif
    Tcl_Eval(interp, "namespace eval " TWAPI_TCL_NAMESPACE " { }");

    Tcl_GetVersion(&TclVersion.major,
                   &TclVersion.minor,
                   &TclVersion.patchlevel,
                   &TclVersion.reltype);

#ifdef _WIN64
    if (TclVersion.major < 8 ||
        (TclVersion.major == 8 && TclVersion.minor < 5)) {
        return TCL_ERROR;
    }
#else
    if (TclVersion.major < 8 ||
        (TclVersion.major == 8 && TclVersion.minor < 4)) {
        return TCL_ERROR;
    }
#endif

    /* Do the once per process stuff */

    /* Yes, there is a tiny race condition
     * here. Even using the Interlocked* functions. TBD
     */
    if (InterlockedIncrement(&twapi_initialized) == 1) {

        /* Single-threaded COM model */
        CoInitializeEx(NULL,
                       COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

        TwapiOSVersionInfo.dwOSVersionInfoSize = sizeof(TwapiOSVersionInfo);
        if (! GetVersionEx(&TwapiOSVersionInfo)) {
            Tcl_SetResult(interp, "Could not get OS version", TCL_STATIC);
            return TCL_ERROR;
        }

        if (Tcl_GetVar2Ex(interp, "tcl_platform", "threaded", TCL_GLOBAL_ONLY))
            TclIsThreaded = 1;
        else
            TclIsThreaded = 0;
    } else {
        // Not first caller. Already initialized
        // TBD - FIX - SHOULD ENSURE WE WAIT FOR FIRST INITIALIZER TO FINISH
        // INITIALIZING
    }

    /* Per interp initialization */

    /* Allocate a context that will be passed around in all interpreters */
    ticP = TwapiInterpContextNew(interp);
    if (ticP == NULL)
        return TCL_ERROR;

    /* For all the commands we register with the Tcl interp, we add a single
     * ref for the context, not one per command. This is sufficient since
     * when the interp gets deleted, all the commands get deleted as well.
     * The corresponding Unref happens when the interp is deleted.
     */
    TwapiInterpContextRef(ticP, 1);

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
    Tcl_CreateExitHandler(Twapi_Cleanup, NULL);

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
}


static TwapiInterpContext* TwapiInterpContextNew(Tcl_Interp *interp)
{
    TwapiInterpContext* ticP = TwapiAlloc(sizeof(*ticP));
    if (ticP == NULL)
        return NULL;

    ticP->nrefs = 0;
    ticP->interp = interp;


    /* Initialize the critical section used for controlling
     * the notification list
     */
    ZLIST_INIT(&ticP->pending);
    ticP->pending_suspended = 0;
    InitializeCriticalSection(&ticP->pending_cs);

    /* Register a async callback with Tcl. */
    /* TBD - do we really need a separate Tcl_AsyncCreate call per interp?
       or should it be per process ? Or per thread ?
    */
    ticP->async_handler = Tcl_AsyncCreate(Twapi_TclAsyncProc, ticP);

    ticP->notification_win = NULL; /* Created only on demand */

    return ticP;
}

/* Enqueues a callback within an interp context. Returns a Win32 error code */
DWORD TwapiInterpContextEnqueueCallback(TwapiInterpContext *ticP,
                                     TwapiPendingCallback *pcbP,
                                     int need_response)
{
    EnterCriticalSection(&ticP->pending_cs);

    if (ticP->pending_suspended) {
        LeaveCriticalSection(&ticP->pending_cs);
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
    TwapiPendingCallbackRef(pcbP, (need_response ? 2 : 1));
    ZLIST_APPEND(&ticP->pending, pcbP); /* Enqueue */

    /* Also make sure the ticP itself does not go away */
    pcbP->ticP = ticP;
    TwapiInterpContextRef(ticP, 1);

    /* To avoid races, note the AsyncMark should also happen in the crit sec */
    Tcl_AsyncMark(ticP->async_handler);
    
    LeaveCriticalSection(&ticP->pending_cs);

    return ERROR_SUCCESS;
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

    DeleteCriticalSection(&ticP->pending_cs);

    /* TBD - should this be in the Twapi_InterpCleanup instead ? */
    if (ticP->notification_win) {
        DestroyWindow(ticP->notification_win);
        ticP->notification_win = 0;
    }
    
}

/* Decrement ref count and free if 0 */
void TwapiInterpContextUnref(TwapiInterpContext *ticP, int decr)
{
    if (InterlockedExchangeAdd(&ticP->nrefs, -decr) <= decr)
        TwapiInterpContextDelete(ticP);
}

static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp)
{
    if (ticP->interp == NULL)
        return;

    ticP->interp = NULL;        /* Must not access during cleanup */
    
    EnterCriticalSection(&ticP->pending_cs);
    ticP->pending_suspended = 1;
    LeaveCriticalSection(&ticP->pending_cs);

    /* TBD - call other callback module clean up procedures */
    
    TwapiInterpContextUnref(ticP, 1);
}
