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

static TCL_RESULT Twapi_IMofCompiler_CompileFileOrBuffer(Tcl_Interp *interp, int type, int objc, Tcl_Obj *CONST objv[])
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
        if (TwapiGetArgs(interp, objc, objv,
                         GETPTR(ifc, IMofCompiler), ARGSKIP, ARGSKIP,
                         ARGUSEDEFAULT,
                         GETNULLIFEMPTY(server_namespace),
                         GETINT(optflags), GETINT(classflags),
                         GETINT(instflags), ARGEND) != TCL_OK)
            return TCL_ERROR;
    } else {
        if (TwapiGetArgs(interp, objc, objv,
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
        return TwapiReturnErrorEx(interp,
                                  TWAPI_BUG,
                                  Tcl_ObjPrintf("Invalid IMofCompiler function code %d", type));
    }

    switch (hr) {
    case 2: /* Warning if autorecover pragma is not present. Not an error */
    case WBEM_S_NO_ERROR:
        return TCL_OK;

    case WBEM_S_FALSE: /* Fall thru */
    default:
        Tcl_SetObjResult(interp,
                         Tcl_ObjPrintf("IMofCompiler error: phase: %d, object number: %d, first line: %d, last line: %d.",
                                       wcsi.lPhaseError,
                                       wcsi.ObjectNum,
                                       wcsi.FirstLine,
                                       wcsi.LastLine));
        return Twapi_AppendSystemError(interp, wcsi.hRes ? wcsi.hRes : hr);
    }
}                     

static int Twapi_WmiCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    --objc;
    ++objv;
    switch (PtrToInt(clientdata)) {
    case 1: // IMofCompiler_CompileBuffer
        return Twapi_IMofCompiler_CompileFileOrBuffer(interp, 0, objc, objv);
    case 2: // IMofCompiler_CompileFile
        return Twapi_IMofCompiler_CompileFileOrBuffer(interp, 1, objc, objv);
    case 3: // IMofCompiler_CreateBMOF
        return Twapi_IMofCompiler_CompileFileOrBuffer(interp, 2, objc, objv);
    }

    return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);
}


static int TwapiWmiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s WmiDispatch[] = {
        DEFINE_FNCODE_CMD(IMofCompiler_CompileBuffer, 1),
        DEFINE_FNCODE_CMD(IMofCompiler_CompileFile, 2),
        DEFINE_FNCODE_CMD(IMofCompiler_CreateBMOF, 3),
    };

    /* Create the underlying call dispatch commands */
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(WmiDispatch), WmiDispatch,
                          Twapi_WmiCallObjCmd);

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
int Twapi_wmi_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiWmiInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;

}

