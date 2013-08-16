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

    ObjAppendElement(p_enum_ctx->interp,
                             p_enum_ctx->objP,
                             ObjFromUnicode(p_winsta));
    return 1;
}

/* Window station enumeration */
int Twapi_EnumWindowStations(Tcl_Interp *interp)
{
    TwapiEnumCtx enum_ctx;

    enum_ctx.interp = interp;
    enum_ctx.objP = ObjEmptyList();

    
    if (EnumWindowStationsW(Twapi_EnumWindowStationsOrDesktopsCallback, (LPARAM)&enum_ctx) == 0) {
        TwapiReturnSystemError(interp);
        Twapi_FreeNewTclObj(enum_ctx.objP);
        return TCL_ERROR;
    }

    ObjSetResult(interp, enum_ctx.objP);
    return TCL_OK;
}

/* Desktop enumeration */
int Twapi_EnumDesktops(Tcl_Interp *interp, HWINSTA hwinsta)
{
    TwapiEnumCtx enum_ctx;

    enum_ctx.interp = interp;
    enum_ctx.objP = ObjEmptyList();

    
    if (EnumDesktopsW(hwinsta, Twapi_EnumWindowStationsOrDesktopsCallback, (LPARAM)&enum_ctx) == 0) {
        TwapiReturnSystemError(interp);
        Twapi_FreeNewTclObj(enum_ctx.objP);
        return TCL_ERROR;
    }

    ObjSetResult(interp, enum_ctx.objP);
    return TCL_OK;
}

static int Twapi_WinstaCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD dw, dw2, dw3;
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h;
    int func = PtrToInt(clientdata);
    TwapiResult result;

    --objc;
    ++objv;
    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_EnumWindowStations(interp);
    case 2:
        result.type = TRT_HWINSTA;
        result.value.hval = GetProcessWindowStation();
        break;
    case 3:
        if (TwapiGetArgs(interp, objc, objv, ARGSKIP,
                         GETINT(dw), GETINT(dw2),
                         GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HWINSTA;
        result.value.hval = CreateWindowStationW(ObjToUnicode(objv[0]), dw, dw2, secattrP);
        TwapiFreeSECURITY_ATTRIBUTES(secattrP);
        break;
    case 4:
        if (TwapiGetArgs(interp, objc, objv,
                         ARGSKIP, GETINT(dw), GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDESK;
        result.value.hval = OpenDesktopW(ObjToUnicode(objv[0]), dw, dw2, dw3);
        break;
    case 5:
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.hval = GetThreadDesktop(dw);
        result.type = TRT_HDESK;
        break;
    case 6:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDESK;
        result.value.hval = OpenInputDesktop(dw, dw2, dw3);
        break;
    case 7:
        if (TwapiGetArgs(interp, objc, objv,
                         ARGSKIP, GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HWINSTA;
        result.value.hval = OpenWindowStationW(ObjToUnicode(objv[0]), dw, dw2);
        break;
    case 8: // CreateDesktopW
        /* Note second, third args are ignored and are reserved as NULL */
        if (TwapiGetArgs(interp, objc, objv,
                         ARGSKIP, ARGUNUSED, ARGUNUSED, GETINT(dw),
                         GETINT(dw2),
                         GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDESK;
        result.value.hval = CreateDesktopW(ObjToUnicode(objv[0]), NULL, NULL, dw, dw2, secattrP);
        if (secattrP)
            TwapiFreeSECURITY_ATTRIBUTES(secattrP);
        break;
    default:
        if (TwapiGetArgs(interp, objc, objv,
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


static int TwapiWinstaInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s WinstaDispatch[] = {
        DEFINE_FNCODE_CMD(EnumWindowStations, 1),
        DEFINE_FNCODE_CMD(GetProcessWindowStation, 2),
        DEFINE_FNCODE_CMD(CreateWindowStation, 3),
        DEFINE_FNCODE_CMD(OpenDesktop, 4),
        DEFINE_FNCODE_CMD(GetThreadDesktop, 5),
        DEFINE_FNCODE_CMD(OpenInputDesktop, 6), // TBD - Tcl
        DEFINE_FNCODE_CMD(OpenWindowStation, 7),
        DEFINE_FNCODE_CMD(CreateDesktop, 8), // TBD - Tcl
        DEFINE_FNCODE_CMD(CloseDesktop, 31),
        DEFINE_FNCODE_CMD(SwitchDesktop, 32),
        DEFINE_FNCODE_CMD(SetThreadDesktop, 33),
        DEFINE_FNCODE_CMD(EnumDesktops, 34),
        DEFINE_FNCODE_CMD(SetProcessWindowStation, 35),
        DEFINE_FNCODE_CMD(CloseWindowStation, 36),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(WinstaDispatch), WinstaDispatch, Twapi_WinstaCallObjCmd);

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
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiWinstaInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

