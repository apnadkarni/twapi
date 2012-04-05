/* 
 * Copyright (c) 2004-2012 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to multimedia */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif


static int Twapi_MmCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR s;
    DWORD dw, dw2;
    HMODULE hmod;
    TwapiResult result;
    int func = (int) clientdata;

    --objc;
    ++objv;
    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        if (TwapiGetArgs(interp, objc, objv,
                         GETNULLIFEMPTY(s), GETHANDLET(hmod, HMODULE), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.ival = PlaySoundW(s, hmod, dw);
        break;
    case 2:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.bval = MessageBeep(dw);
        break;
    case 3:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = Beep(dw, dw2);
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiMmInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s MmDispatch[] = {
        DEFINE_FNCODE_CMD(PlaySound, 1),
        DEFINE_FNCODE_CMD(MessageBeep, 2),
        DEFINE_FNCODE_CMD(Beep, 3),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(MmDispatch), MmDispatch, Twapi_MmCallObjCmd);
    
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
int Twapi_multimedia_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiMmInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

