/* 
 * Copyright (c) 2006-2009, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

Tcl_Obj *ObjFromDISPLAY_DEVICE(DISPLAY_DEVICEW *ddP)
{
    Tcl_Obj *objv[5];

    objv[0] = Tcl_NewUnicodeObj(ddP->DeviceName, -1);
    objv[1] = Tcl_NewUnicodeObj(ddP->DeviceString, -1);
    objv[2] = Tcl_NewIntObj(ddP->StateFlags);
    objv[3] = Tcl_NewUnicodeObj(ddP->DeviceID, -1);
    objv[4] = Tcl_NewUnicodeObj(ddP->DeviceKey, -1);

    return Tcl_NewListObj(5, objv);
}

Tcl_Obj *ObjFromMONITORINFOEX(MONITORINFO *miP)
{
    Tcl_Obj *objv[4];
    int      objc;

    objc = 3;
    objv[0] = ObjFromRECT(&miP->rcMonitor);
    objv[1] = ObjFromRECT(&miP->rcWork);
    objv[2] = Tcl_NewLongObj(miP->dwFlags);
    /* miP could be either a MONITORINFO or MONITORINFOEX or MONITORINFOEXW */
    if (miP->cbSize == sizeof(MONITORINFOEX)) {
        objv[3] = Tcl_NewStringObj(((MONITORINFOEX *)miP)->szDevice, -1);
        objc = 4;
    }
    else if (miP->cbSize == sizeof(MONITORINFOEXW)) {
        objv[3] = Tcl_NewUnicodeObj(((MONITORINFOEXW *)miP)->szDevice, -1);
        objc = 4;
    }

    return Tcl_NewListObj(objc, objv);
}

/* Window enumeration callback */
BOOL CALLBACK Twapi_EnumDisplayMonitorsCallback(
    HMONITOR hmon, 
    HDC      hdc,
    RECT    *rectP,
    LPARAM p_ctx
)
{
    Tcl_Obj *objv[3];
    struct Twapi_EnumCtx *p_enum_ctx =
        (struct Twapi_EnumCtx *) p_ctx;

    objv[0] = ObjFromOpaque(hmon, "HMONITOR");
    objv[1] = ObjFromOpaque(hdc, "HDC");
    objv[2] = ObjFromRECT(rectP);
    Tcl_ListObjAppendElement(p_enum_ctx->interp,
                             p_enum_ctx->win_listobj,
                             Tcl_NewListObj(3, objv));
    return 1;
}

/*
 * Enumerate monitors
 */
int Twapi_EnumDisplayMonitors(
    Tcl_Interp *interp,
    HDC         hdc,
    const RECT *rectP
)
{
    struct Twapi_EnumCtx enum_ctx;

    enum_ctx.interp = interp;
    enum_ctx.win_listobj = Tcl_NewListObj(0, NULL);
    
    if (EnumDisplayMonitors(hdc, rectP, Twapi_EnumDisplayMonitorsCallback, (LPARAM)&enum_ctx) == 0) {
        TwapiReturnSystemError(interp);
        Twapi_FreeNewTclObj(enum_ctx.win_listobj);
        return TCL_ERROR;
    }

    Tcl_SetObjResult(interp, enum_ctx.win_listobj);
    return TCL_OK;
}
