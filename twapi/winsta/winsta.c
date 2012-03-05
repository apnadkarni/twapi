/* 
 * Copyright (c) 2012 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/*
 * Define interface to Windows API related to window stations and desktops
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

/* Window station enumeration callback */
BOOL CALLBACK Twapi_EnumWindowStationsOrDesktopsCallback(LPCWSTR p_winsta, LPARAM p_ctx) {
    TwapiEnumCtx *p_enum_ctx =
        (TwapiEnumCtx *) p_ctx;

    Tcl_ListObjAppendElement(p_enum_ctx->interp,
                             p_enum_ctx->objP,
                             ObjFromUnicode(p_winsta));
    return 1;
}

/* Window station enumeration */
int Twapi_EnumWindowStations(Tcl_Interp *interp)
{
    TwapiEnumCtx enum_ctx;

    enum_ctx.interp = interp;
    enum_ctx.objP = Tcl_NewListObj(0, NULL);

    
    if (EnumWindowStationsW(Twapi_EnumWindowStationsOrDesktopsCallback, (LPARAM)&enum_ctx) == 0) {
        TwapiReturnSystemError(interp);
        Twapi_FreeNewTclObj(enum_ctx.objP);
        return TCL_ERROR;
    }

    Tcl_SetObjResult(interp, enum_ctx.objP);
    return TCL_OK;
}

/* Desktop enumeration */
int Twapi_EnumDesktops(Tcl_Interp *interp, HWINSTA hwinsta)
{
    TwapiEnumCtx enum_ctx;

    enum_ctx.interp = interp;
    enum_ctx.objP = Tcl_NewListObj(0, NULL);

    
    if (EnumDesktopsW(hwinsta, Twapi_EnumWindowStationsOrDesktopsCallback, (LPARAM)&enum_ctx) == 0) {
        TwapiReturnSystemError(interp);
        Twapi_FreeNewTclObj(enum_ctx.objP);
        return TCL_ERROR;
    }

    Tcl_SetObjResult(interp, enum_ctx.objP);
    return TCL_OK;
}

static int Twapi_WinstaCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s;
    DWORD dw, dw2, dw3;
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h;
    TwapiResult result;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_EnumWindowStations(interp);
    case 2:
        result.type = TRT_HWINSTA;
        result.value.hval = GetProcessWindowStation();
        break;
    case 3:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETINT(dw), GETINT(dw2),
                         GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HWINSTA;
        result.value.hval = CreateWindowStationW(s, dw, dw2, secattrP);
        TwapiFreeSECURITY_ATTRIBUTES(secattrP);
        break;
    case 4:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETINT(dw), GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDESK;
        result.value.hval = OpenDesktopW(s, dw, dw2, dw3);
        break;
    case 5:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.hval = GetThreadDesktop(dw);
        result.type = TRT_HDESK;
        break;
    case 6:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETINT(dw), GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDESK;
        result.value.hval = OpenInputDesktop(dw, dw2, dw3);
        break;
    case 7:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HWINSTA;
        result.value.hval = OpenWindowStationW(s, dw, dw2);
        break;
    case 8: // CreateDesktopW
        /* Note second, third args are ignored and are reserved as NULL */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), ARGSKIP, ARGSKIP, GETINT(dw),
                         GETINT(dw2), ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (ObjToPSECURITY_ATTRIBUTES(interp, objv[7], &secattrP) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDESK;
        result.value.hval = CreateDesktopW(s, NULL, NULL, dw, dw2, secattrP);
        if (secattrP)
            TwapiFreeSECURITY_ATTRIBUTES(secattrP);
        break;
    default:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 31:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseDesktop(h);
            break;
        case 32:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SwitchDesktop(h);
            break;
        case 33:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetThreadDesktop(h);
            break;
        case 34:
            return Twapi_EnumDesktops(interp, h);
        case 35:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetProcessWindowStation(h);
            break;
        case 36:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseWindowStation(h);
            break;
        }
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_WinstaInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::WinstaCall", Twapi_WinstaCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::WinstaCall", # code_); \
    } while (0);

    CALL_(EnumWindowStations, 1);
    CALL_(GetProcessWindowStation, 2);
    CALL_(CreateWindowStation, 3);
    CALL_(OpenDesktop, 4);
    CALL_(GetThreadDesktop, 5);
    CALL_(OpenInputDesktop, 6); // TBD - Tcl
    CALL_(OpenWindowStation, 7);
    CALL_(CreateDesktop, 8); // TBD - Tcl
    CALL_(CloseDesktop, 31);
    CALL_(SwitchDesktop, 32);
    CALL_(SetThreadDesktop, 33);
    CALL_(EnumDesktops, 34);
    CALL_(SetProcessWindowStation, 35);
    CALL_(CloseWindowStation, 36);

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
int Twapi_winsta_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, WLITERAL(MODULENAME), MODULE_HANDLE,
                            Twapi_WinstaInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

