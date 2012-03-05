/* 
 * Copyright (c) 2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include <wbemidl.h>

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

TCL_RESULT Twapi_IMofCompiler_CompileFileOrBuffer(TwapiInterpContext *ticP, int type, int objc, Tcl_Obj *CONST objv[])
{
    IMofCompiler *ifc;
    WCHAR *server_namespace;
    WCHAR *user;
    WCHAR *password;
    WCHAR *authority;
    int    optflags, classflags, instflags;
    BYTE  *buf;
    int    buflen;
    HRESULT hr;
    WBEM_COMPILE_STATUS_INFO wcsi;

    if (type == 2) {
        if (TwapiGetArgs(ticP->interp, objc, objv,
                         GETPTR(ifc, IMofCompiler), ARGSKIP, ARGSKIP,
                         ARGUSEDEFAULT,
                         GETNULLIFEMPTY(server_namespace),
                         GETINT(optflags), GETINT(classflags),
                         GETINT(instflags), ARGEND) != TCL_OK)
            return TCL_ERROR;
    } else {
        if (TwapiGetArgs(ticP->interp, objc, objv,
                         GETPTR(ifc, IMofCompiler), ARGSKIP, ARGUSEDEFAULT,
                         GETNULLIFEMPTY(server_namespace), GETNULLIFEMPTY(user),
                         GETNULLIFEMPTY(authority), GETNULLIFEMPTY(password),
                         GETINT(optflags), GETINT(classflags),
                         GETINT(instflags), ARGEND) != TCL_OK)
            return TCL_ERROR;
    }            

    ZeroMemory(&wcsi, sizeof(wcsi));

    switch (type) {
    case 0:
        buf = Tcl_GetStringFromObj(objv[1], &buflen);
        hr = ifc->lpVtbl->CompileBuffer(ifc, buflen, buf, server_namespace,
                                         user, authority, password,
                                         optflags, classflags, instflags,
                                         &wcsi);
        break;
    case 1:
        hr = ifc->lpVtbl->CompileFile(ifc, Tcl_GetUnicode(objv[1]),
                                       server_namespace,
                                       user, authority, password,
                                       optflags, classflags, instflags,
                                       &wcsi);
        break;

    case 2:
        hr = ifc->lpVtbl->CreateBMOF(ifc, Tcl_GetUnicode(objv[1]),
                                     Tcl_GetUnicode(objv[2]),
                                     server_namespace,
                                     optflags, classflags, instflags,
                                     &wcsi);
        break;

    default:
        return TwapiReturnErrorEx(ticP->interp,
                                  TWAPI_BUG,
                                  Tcl_ObjPrintf("Invalid IMofCompiler function code %d", type));
    }

    switch (hr) {
    case 2: /* Warning if autorecover pragma is not present. Not an error */
    case WBEM_S_NO_ERROR:
        return TCL_OK;

    case WBEM_S_FALSE: /* Fall thru */
    default:
        Tcl_SetObjResult(ticP->interp,
                         Tcl_ObjPrintf("IMofCompiler error: phase: %d, object number: %d, first line: %d, last line: %d.",
                                       wcsi.lPhaseError,
                                       wcsi.ObjectNum,
                                       wcsi.FirstLine,
                                       wcsi.LastLine));
        return Twapi_AppendSystemError(ticP->interp, wcsi.hRes ? wcsi.hRes : hr);
    }
}                     

static int Twapi_WmiCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    switch (func) {
    case 1: // IMofCompiler_CompileBuffer
        return Twapi_IMofCompiler_CompileFileOrBuffer(ticP, 0, objc-2, objv+2);
    case 2: // IMofCompiler_CompileFile
        return Twapi_IMofCompiler_CompileFileOrBuffer(ticP, 1, objc-2, objv+2);
    case 3: // IMofCompiler_CreateBMOF
        return Twapi_IMofCompiler_CompileFileOrBuffer(ticP, 2, objc-2, objv+2);
    }

    return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);
}


static int Twapi_WmiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::WmiCall", Twapi_WmiCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::WmiCall", # code_); \
    } while (0);

    CALL_(IMofCompiler_CompileBuffer, 1); // Tcl
    CALL_(IMofCompiler_CompileFile, 2); // Tcl
    CALL_(IMofCompiler_CreateBMOF, 3); // Tcl

#undef CALL_

    return TCL_OK;
}


#ifndef TWAPI_SINGLE_MODULE
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_SINGLE_MODULE
__declspec(dllexport) 
#endif
int Twapi_Wmi_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, WLITERAL(MODULENAME), MODULE_HANDLE,
                            Twapi_WmiInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

