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
OSVERSIONINFO TwapiOSVersionInfo;
GUID TwapiNullGuid;             /* Initialized to all zeroes */
struct TwapiTclVersion TclVersion;
int TclIsThreaded;
/*
 * Whether the callback dll/libray has been initialized.
 * The value must be managed using the InterlockedCompareExchange functions to
 * ensure thread safety. The value returned by InterlockedCompareExhange
 * 0 -> first to call, do init,  1 -> init in progress by some other thread
 * 2 -> Init done
 */
static TwapiOneTimeInitState TwapiInitialized;


static void Twapi_Cleanup(ClientData clientdata);
static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp);
static TwapiInterpContext *TwapiInterpContextNew(Tcl_Interp *interp);
static void TwapiInterpContextDelete(ticP, interp);
static int TwapiOneTimeInit(Tcl_Interp *interp);

/* Main entry point */
__declspec(dllexport) int Twapi_Init(Tcl_Interp *interp)
{
    int result;
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
    if (! TwapiDoOneTimeInit(&TwapiInitialized, TwapiOneTimeInit, interp))
        return TCL_ERROR;

    /* NOTE: no point setting Tcl_SetResult for errors as they are not
       looked at when DLL is being loaded */

    /*
     * Per interp initialization
     */

    /* Create the name space. Needed for some scripts bound into the dll */
    Tcl_Eval(interp, "namespace eval " TWAPI_TCL_NAMESPACE " { }");

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

    WSACleanup();
}


static TwapiInterpContext* TwapiInterpContextNew(Tcl_Interp *interp)
{
    TwapiInterpContext* ticP = TwapiAlloc(sizeof(*ticP));
    if (ticP == NULL)
        return NULL;

    ticP->nrefs = 0;
    ticP->interp = interp;
    ticP->thread = Tcl_GetCurrentThread();

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

    DeleteCriticalSection(&ticP->pending_cs);

    /* TBD - should rest of this be in the Twapi_InterpCleanup instead ? */
    if (ticP->notification_win) {
        DestroyWindow(ticP->notification_win);
        ticP->notification_win = 0;
    }
    if (ticP->clipboard_win) {
        DestroyWindow(ticP->clipboard_win);
        ticP->clipboard_win = 0;
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

    if (ticP->console_ctrl_hooked)
        Twapi_StopConsoleEventNotifier(ticP);

    /* TBD - call other callback module clean up procedures */
    // TBD - terminate device notification thread;
    
    TwapiInterpContextUnref(ticP, 1);
}


/* One time (per process) initialization */
static int TwapiOneTimeInit(Tcl_Interp *interp)
{
    HRESULT hr;
    WSADATA ws_data;
    WORD    ws_ver = MAKEWORD(1,1);


    if (Tcl_GetVar2Ex(interp, "tcl_platform", "threaded", TCL_GLOBAL_ONLY))
        TclIsThreaded = 1;
    else
        TclIsThreaded = 0;

    /*
     * Deeply nested if's because it is easier to track what to undo
     * in case of errors
     */

    Tcl_GetVersion(&TclVersion.major,
                   &TclVersion.minor,
                   &TclVersion.patchlevel,
                   &TclVersion.reltype);
    if (TclVersion.major ==  TWAPI_TCL_MAJOR &&
        TclVersion.minor >= TWAPI_MIN_TCL_MINOR) {
        TwapiOSVersionInfo.dwOSVersionInfoSize =
            sizeof(TwapiOSVersionInfo);
        if (GetVersionEx(&TwapiOSVersionInfo)) {
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
