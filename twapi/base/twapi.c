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

/* Main entry point */
__declspec(dllexport) int Twapi_Init(Tcl_Interp *interp)
{
    int result;
    static LONG twapi_initialized;

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
#ifdef USE_TCL_STUBS
    if (Tcl_InitStubs(interp, (char*)"8.1", 0) == NULL) {
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

#ifdef STATIC_CALLBACK_BUILD
# ifndef TWAPI_NOCALLBACKS
    if (Twapicallback_Init(interp) != TCL_OK)
        return TCL_ERROR;
# endif
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

    /* Do our own commands. */
    if (Twapi_InitCalls(interp) != TCL_OK) {
        return TCL_ERROR;
    }
    Tcl_CreateObjCommand(interp, "twapi::parseargs", Twapi_ParseargsObjCmd,
                         NULL, NULL);
    Tcl_CreateObjCommand(interp, "twapi::try", Twapi_TryObjCmd,
                         NULL, NULL);
    Tcl_CreateObjCommand(interp, "twapi::kl_get", Twapi_KlGetObjCmd,
                         NULL, NULL);
    Tcl_CreateObjCommand(interp, "twapi::twine", Twapi_TwineObjCmd,
                         NULL, NULL);
    Tcl_CreateObjCommand(interp, "twapi::recordarray", Twapi_RecordArrayObjCmd,
                         NULL, NULL);
    Tcl_CreateObjCommand(interp, "twapi::get_build_config",
                         Twapi_GetTwapiBuildInfo, NULL, NULL);
#ifndef TWAPI_NODESKTOP
    Tcl_CreateObjCommand(interp, "twapi::IDispatch_Invoke", Twapi_IDispatch_InvokeObjCmd,
                         NULL, NULL);
    Tcl_CreateObjCommand(interp, "twapi::ComEventSink", Twapi_ComEventSinkObjCmd,
                         NULL, NULL);
    Tcl_CreateObjCommand(interp, "twapi::SHChangeNotify", Twapi_SHChangeNotify,
                         NULL, NULL);
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

