/* 
 * Copyright (c) 2003-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to UI functions */
/* TBD - IsHungAppWindow */

/*
TBD functions:
EnumThreadWindows
GetKeyboardState
TileWindows
SetKeyboardState
ToAscii
ToUnicode
VkKeyScan
*/



/* TBD - TMT_* defines for GetThemeColor propid's
  BORDERCOLOR FILLCOLOR TEXTCOLOR EDGELIGHTCOLOR EDGESHADOWCOLOR EDGEFILLCOLOR TRASPARENTCOLOR GRADIENTCOLOR1 GRADIENTCOLOR2 GRADIENTCOLOR3 GRADIENTCOLOR4 GRADIENTCOLOR5 SHADOWCOLOR GLOWCOLOR TEXTBORDERCOLOR TEXTSHADOWCOLOR GLYPHTEXTCOLOR FILLCOLORHINT BORDERCOLORHINT ACCENTCOLORHINT BLENDCOLOR */

#include "twapi.h"
#include "twapi_ui.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

/* Return Tcl List for a LOGFONT structure */
Tcl_Obj *ObjFromLOGFONTW(LOGFONTW *lfP)
{
    Tcl_Obj *objv[28];

    objv[0] = STRING_LITERAL_OBJ("lfHeight");
    objv[1] = ObjFromLong(lfP->lfHeight);
    objv[2] = STRING_LITERAL_OBJ("lfWidth");
    objv[3] = ObjFromLong(lfP->lfWidth);
    objv[4] = STRING_LITERAL_OBJ("lfEscapement");
    objv[5] = ObjFromLong(lfP->lfEscapement);
    objv[6] = STRING_LITERAL_OBJ("lfOrientation");
    objv[7] = ObjFromLong(lfP->lfOrientation);
    objv[8] = STRING_LITERAL_OBJ("lfWeight");
    objv[9] = ObjFromLong(lfP->lfWeight);
    objv[10] = STRING_LITERAL_OBJ("lfItalic");
    objv[11] = ObjFromLong(lfP->lfItalic);
    objv[12] = STRING_LITERAL_OBJ("lfUnderline");
    objv[13] = ObjFromLong(lfP->lfUnderline);
    objv[14] = STRING_LITERAL_OBJ("lfStrikeOut");
    objv[15] = ObjFromLong(lfP->lfStrikeOut);
    objv[16] = STRING_LITERAL_OBJ("lfCharSet");
    objv[17] = ObjFromLong(lfP->lfCharSet);
    objv[18] = STRING_LITERAL_OBJ("lfOutPrecision");
    objv[19] = ObjFromLong(lfP->lfOutPrecision);
    objv[20] = STRING_LITERAL_OBJ("lfClipPrecision");
    objv[21] = ObjFromLong(lfP->lfClipPrecision);
    objv[22] = STRING_LITERAL_OBJ("lfQuality");
    objv[23] = ObjFromLong(lfP->lfQuality);
    objv[24] = STRING_LITERAL_OBJ("lfPitchAndFamily");
    objv[25] = ObjFromLong(lfP->lfPitchAndFamily);
    objv[26] = STRING_LITERAL_OBJ("lfFaceName");
    objv[27] = ObjFromUnicode(lfP->lfFaceName);

    return ObjNewList(28, objv);
}


int ObjToFLASHWINFO (Tcl_Interp *interp, Tcl_Obj *obj, FLASHWINFO *fwP)
{
    Tcl_Obj **objv;
    int       objc;
    HWND      hwnd;

    if (ObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    fwP->cbSize = sizeof(*fwP);
    if (objc == 4 &&
        (ObjToHWND(interp, objv[0], &hwnd) == TCL_OK) &&
        (ObjToLong(interp, objv[1], &fwP->dwFlags) == TCL_OK) &&
        (ObjToLong(interp, objv[2], &fwP->uCount) == TCL_OK) &&
        (ObjToLong(interp, objv[3], &fwP->dwTimeout) == TCL_OK)) {
        fwP->hwnd = (HWND) hwnd;
        return TCL_OK;
    }

    ObjSetStaticResult(interp, "Invalid FLASHWINFO structure. Need 4 elements - window handle, flags, count and timeout.");
    return TCL_ERROR;
}

Tcl_Obj *ObjFromWINDOWPLACEMENT(WINDOWPLACEMENT *wpP)
{
    Tcl_Obj *objv[5];
    objv[0] = ObjFromLong(wpP->flags);
    objv[1] = ObjFromLong(wpP->showCmd);
    objv[2] = ObjFromPOINT(&wpP->ptMinPosition);
    objv[3] = ObjFromPOINT(&wpP->ptMaxPosition);
    objv[4] = ObjFromRECT(&wpP->rcNormalPosition);
    return ObjNewList(5, objv);
}

int ObjToWINDOWPLACEMENT(Tcl_Interp *interp, Tcl_Obj *objP, WINDOWPLACEMENT *wpP)
{
    Tcl_Obj **objv;
    int objc;

    if (ObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc != 5)
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Incorrect format of WINDOWPLACEMENT argument.");

    if (ObjToLong(interp, objv[0], &wpP->flags) != TCL_OK ||
        ObjToLong(interp, objv[1], &wpP->showCmd) != TCL_OK ||
        ObjToPOINT(interp, objv[2], &wpP->ptMinPosition) != TCL_OK ||
        ObjToPOINT(interp, objv[3], &wpP->ptMaxPosition) != TCL_OK ||
        ObjToRECT(interp, objv[4], &wpP->rcNormalPosition) != TCL_OK) {
        return TCL_ERROR;
    }

    wpP->length = sizeof(*wpP);
    return TCL_OK;
}

Tcl_Obj *ObjFromWINDOWINFO (WINDOWINFO *wiP)
{
    Tcl_Obj *objv[9];
    objv[0] = ObjFromRECT(&wiP->rcWindow);
    objv[1] = ObjFromRECT(&wiP->rcClient);
    objv[2] = ObjFromLong(wiP->dwStyle);
    objv[3] = ObjFromLong(wiP->dwExStyle);
    objv[4] = ObjFromLong(wiP->dwWindowStatus);
    objv[5] = ObjFromLong(wiP->cxWindowBorders);
    objv[6] = ObjFromLong(wiP->cyWindowBorders);
    objv[7] = ObjFromLong(wiP->atomWindowType);
    objv[8] = ObjFromLong(wiP->wCreatorVersion);

    return ObjNewList(9, objv);
}

/*
 * Enumerate child windows of windows with handle parent_handle
 */
int Twapi_EnumChildWindows(Tcl_Interp *interp, HWND parent_handle)
{
    TwapiEnumCtx enum_win_ctx;

    enum_win_ctx.interp = interp;
    enum_win_ctx.objP = ObjEmptyList();
    
    /* As per newer MSDN docs, EnumChildWindows return value is meaningless.
     * We used to check GetLastError ERROR_PROC_NOT_FOUND and
     * ERROR_MOD_NOT_FOUND but
     * on Win2008R2, it returns ERROR_FILE_NOT_FOUND if no child windows.
     * We now follow the docs and ignore return value.
     */
    EnumChildWindows(parent_handle, Twapi_EnumWindowsCallback, (LPARAM)&enum_win_ctx);
    ObjSetResult(interp, enum_win_ctx.objP);
    return TCL_OK;
}

/* Enumerate desktop windows
 */
int Twapi_EnumDesktopWindows(Tcl_Interp *interp, HDESK desk_handle)
{
    TwapiEnumCtx enum_win_ctx;

    enum_win_ctx.interp = interp;
    enum_win_ctx.objP = ObjEmptyList();
    
    if (EnumDesktopWindows(desk_handle, Twapi_EnumWindowsCallback, (LPARAM)&enum_win_ctx) == 0) {
        /* Maybe the error is just that there are no child windows */
        int winerror = GetLastError();
        if (winerror && winerror != ERROR_INVALID_HANDLE) {
            /* Genuine error */
            Twapi_AppendSystemError(interp, winerror);
            Twapi_FreeNewTclObj(enum_win_ctx.objP);
            return TCL_ERROR;
        }
    }

    ObjSetResult(interp, enum_win_ctx.objP);
    return TCL_OK;
}


int Twapi_GetGUIThreadInfo(Tcl_Interp *interp, DWORD idThread)
{
    Tcl_Obj *objv[18];
    GUITHREADINFO gti;

    gti.cbSize = sizeof(gti);
    if (GetGUIThreadInfo(idThread, &gti) == 0) {
        return TwapiReturnSystemError(interp);
    }

    objv[0] = STRING_LITERAL_OBJ("cbSize");
    objv[1] = ObjFromLong(gti.cbSize);
    objv[2] = STRING_LITERAL_OBJ("flags");
    objv[3] = ObjFromLong(gti.flags);
    objv[4] = STRING_LITERAL_OBJ("hwndActive");
    objv[5] = ObjFromHWND(gti.hwndActive);
    objv[6] = STRING_LITERAL_OBJ("hwndFocus");
    objv[7] = ObjFromHWND(gti.hwndFocus);
    objv[8] = STRING_LITERAL_OBJ("hwndCapture");
    objv[9] = ObjFromHWND(gti.hwndCapture);
    objv[10] = STRING_LITERAL_OBJ("hwndMenuOwner");
    objv[11] = ObjFromHWND(gti.hwndMenuOwner);
    objv[12] = STRING_LITERAL_OBJ("hwndMoveSize");
    objv[13] = ObjFromHWND(gti.hwndMoveSize);
    objv[14] = STRING_LITERAL_OBJ("hwndCaret");
    objv[15] = ObjFromHWND(gti.hwndCaret);
    objv[16] = STRING_LITERAL_OBJ("rcCaret");
    objv[17] = ObjFromRECT(&gti.rcCaret);

    ObjSetResult(interp, ObjNewList(18, objv));
    return TCL_OK;
}


int Twapi_GetCurrentThemeName(Tcl_Interp *interp)
{
    WCHAR filename[MAX_PATH];
    WCHAR color[256];
    WCHAR size[256];
    HRESULT status;
    Tcl_Obj *objv[3];

    status = GetCurrentThemeName(filename, ARRAYSIZE(filename),
                   color, ARRAYSIZE(color),
                   size, ARRAYSIZE(size));

    if (status != S_OK) {
        return Twapi_AppendSystemError(interp, status);
    }

    objv[0] = ObjFromUnicode(filename);
    objv[1] = ObjFromUnicode(color);
    objv[2] = ObjFromUnicode(size);
    ObjSetResult(interp, ObjNewList(3, objv));
    return TCL_OK;
}

int Twapi_GetThemeColor(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HTHEME hTheme;
    int iPartId;
    int iStateId;
    int iPropId;

    HRESULT status;
    COLORREF color;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hTheme), GETINT(iPartId), GETINT(iStateId),
                     GETINT(iPropId),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    status =  GetThemeColor(hTheme, iPartId, iStateId, iPropId, &color);
    if (status != S_OK)
        return Twapi_AppendSystemError(interp, status);

    ObjSetResult(interp,
                     Tcl_ObjPrintf("#%2.2x%2.2x%2.2x",
                                   GetRValue(color),
                                   GetGValue(color),
                                   GetBValue(color)));
    return TCL_OK;
}


int Twapi_GetThemeFont(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HTHEME hTheme;
    HANDLE hdc;
    int iPartId;
    int iStateId;
    int iPropId;
    LOGFONTW lf;
    HRESULT hr;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hTheme),  GETHANDLE(hdc),
                     GETINT(iPartId), GETINT(iStateId), GETINT(iPropId),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

#if VER_PRODUCTVERSION <= 3790
    /* NOTE GetThemeFont ExPECTS LOGFONTW although the documentation/header
     * mentions LOGFONT
     */
    hr =  GetThemeFont(hTheme, hdc, iPartId, iStateId, iPropId, (LOGFONT*)&lf);
#else
    hr =  GetThemeFont(hTheme, hdc, iPartId, iStateId, iPropId, &lf);
#endif
    if (hr != S_OK)
        return Twapi_AppendSystemError(interp, hr);

    ObjSetResult(interp, ObjFromLOGFONTW(&lf));
    return TCL_OK;
}

static int Twapi_UiCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD dw, dw2, dw3, dw4, dw5, dw6;
    HANDLE h;
    TwapiResult result;
    union {
        FLASHWINFO flashw;
        MONITORINFOEXW minfo;
        RECT rect;
        POINT  pt;
        HGDIOBJ hgdiobj;
        DISPLAY_DEVICEW display_device;
        LOGFONTW lf;
    } u;
    RECT *rectP;
    LPVOID pv;
    int func = PtrToInt(clientdata);

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 1000) {
        /* Functions taking no arguments */

        HWND (WINAPI *hfn)(void);
        DWORD (WINAPI *dfn)();
        if (objc != 0)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 1:
        case 2:
            result.type = (func == 1 ? GetCursorPos : GetCaretPos) (&result.value.point) ? TRT_POINT : TRT_GETLASTERROR;
            break;
        case 3: dfn = GetCaretBlinkTime; break;
        case 4: hfn = GetFocus; break;
        case 5: hfn = GetDesktopWindow; break;
        case 6: hfn = GetShellWindow; break;
        case 7: hfn = GetForegroundWindow; break;
        case 8: hfn = GetActiveWindow; break;
        case 9: dfn = IsThemeActive; break;
        case 10: dfn = IsAppThemed; break;
        case 11:
            return Twapi_GetCurrentThemeName(interp);
        }
        if (hfn) {
            result.type = TRT_HWND;
            result.value.hwin = hfn();
        } else if (dfn) {
            result.type = TRT_DWORD;
            result.value.uval = dfn();
        }
    } else if (func < 2000) {
        /* Exactly one arg */
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1001:
            if (ObjToPOINT(interp, objv[0], &u.pt) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HWND;
            result.value.hwin = WindowFromPoint(u.pt);
            break;
        case 1002:
            if (ObjToFLASHWINFO(interp, objv[0], &u.flashw) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = FlashWindowEx(&u.flashw);
            break;
        }
    } else if (func < 4000) {
        /*
         * At least one integer arg. Additional args ignored for some
         * so arg errors are not quite accurate but that's ok, the
         * Tcl wrappers catch them.
         */
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw), ARGUSEDEFAULT,
                         GETINT(dw2), GETINT(dw3), GETINT(dw4),
                         GETINT(dw5), GETINT(dw6),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 3001:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetCaretBlinkTime(dw);
            break;
        case 3002:
            return Twapi_GetGUIThreadInfo(interp, dw);
        case 3003:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetCaretPos(dw, dw2);
            break;
        case 3004:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetCursorPos(dw, dw2);
            break;
        case 3005: // CreateRectRgn
        case 3006: // CreateEllipticRgn
            result.type = TRT_HRGN;
            result.value.hval = (func == 3001 ? CreateRectRgn : CreateEllipticRgn)(dw, dw2, dw3, dw4);
            break;
        case 3007: // CreateRoundRectRgn
            result.type = TRT_HRGN;
            result.value.hval = CreateRoundRectRgn(dw, dw2, dw3, dw4, dw5, dw6);
            break;
        }
    } else if (func < 5000) {
        /* First arg is a handle */
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h),
                         ARGUSEDEFAULT, GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 4001:
            result.type = TRT_EMPTY;
            CloseThemeData(h);
            break;
        case 4002:
            return Twapi_EnumDesktopWindows(interp, h);
        case 4003:
            u.minfo.cbSize = sizeof(u.minfo);
            if (GetMonitorInfoW(h, (MONITORINFO *)&u.minfo)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromMONITORINFOEX((MONITORINFO *)&u.minfo);
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 4004: // GetThemeSysColor
            result.type = TRT_DWORD;
            result.value.uval = GetThemeSysColor(h, dw);
            break;
        case 4005: // GetThemeSysFont
            result.type = TRT_OBJ;
#if VER_PRODUCTVERSION <= 3790
            /* NOTE GetThemeSysFont ExPECTS LOGFONTW although the
             * docs/header mentions LOGFONT
             */
            result.value.ival =  GetThemeSysFont(h, dw, (LOGFONT*)&u.lf);
#else
            result.value.ival =  GetThemeSysFont(h, dw, &u.lf);
#endif
            if (result.value.ival == S_OK) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromLOGFONTW(&u.lf);
            } else
                result.type = TRT_EXCEPTION_ON_ERROR;
            break;
        case 4006:
            /* Find the required buffer size */
            dw = GetObjectW(h, 0, NULL);
            if (dw == 0) {
                result.type = TRT_GETLASTERROR; /* TBD - is GetLastError set? */
                break;
            }
            result.value.obj = ObjFromByteArray(NULL, dw); // Alloc storage
            pv = ObjToByteArray(result.value.obj, &dw); // and get ptr to it
            dw = GetObjectW(h, dw, pv);
            if (dw == 0)
                result.type = TRT_GETLASTERROR;
            else {
                Tcl_SetByteArrayLength(result.value.obj, dw);
                result.type = TRT_OBJ;
            }
            break;
        case 4007:
            result.type = TRT_LONG;
            result.value.ival = GetDeviceCaps(h, dw);
            break;
        }
    } else {
        /* Arbitrary args */
        switch (func) {
        case 10001:
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(u.pt, ObjToPOINT), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromPoint(u.pt, dw);
            break;
        case 10002:
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(u.rect, ObjToRECT), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromRect(&u.rect, dw);
            break;
        case 10003: 
            return Twapi_GetThemeColor(interp, objc, objv);
        case 10004: 
            return Twapi_GetThemeFont(interp, objc, objv);
        case 10005:
            if (TwapiGetArgs(interp, objc, objv, ARGSKIP,
                             GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            u.display_device.cb = sizeof(u.display_device);
            if (EnumDisplayDevicesW(ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
                                    dw, &u.display_device, dw2)) {
                result.value.obj = ObjFromDISPLAY_DEVICE(&u.display_device);
                result.type = TRT_OBJ;
            } else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = ERROR_INVALID_PARAMETER;
            }
            break;
        case 10006:
            if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h),
                             ARGSKIP, ARGEND) != TCL_OK)
                return TCL_ERROR;
            rectP = &u.rect;
            if (ObjToRECT_NULL(interp, objv[1], &rectP) != TCL_OK)
                return TCL_ERROR;
            return Twapi_EnumDisplayMonitors(interp, h, rectP);
        case 10007:
            CHECK_NARGS(interp, objc, 2);
            result.type = TRT_HWND;
            result.value.hwin = FindWindowW(
                ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
                ObjToLPWSTR_NULL_IF_EMPTY(objv[1]));
            break;
        case 10008: // AddFontResourceExW
        case 10009: // RemoveFontResourceExW
            if (TwapiGetArgs(interp, objc, objv, ARGSKIP,
                             ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.value.ival =
                (func == 10008 ? AddFontResourceExW : RemoveFontResourceExW) (ObjToUnicode(objv[0]), dw, NULL);
            result.type = TRT_LONG;
            break;
        }
    }
    return TwapiSetResult(interp, &result);
}

int Twapi_UiCallWObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HWND hwnd, hwnd2;
    TwapiResult result;
    DWORD dw, dw2, dw3, dw4, dw5;
    int func = PtrToInt(clientdata);
    union {
        WINDOWPLACEMENT winplace;
        WINDOWINFO wininfo;
        WCHAR buf[MAX_PATH+1];
        RECT   rect;
        HRGN   hrgn;
    } u;
    RECT *rectP;
    HANDLE h;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHWND(hwnd),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    objc -= 2;
    objv += 2;
    result.type = TRT_BADFUNCTIONCODE;
    if (func < 500) {
        BOOL (WINAPI *bfn)(HWND) = NULL;
        DWORD (WINAPI *dfn)(HWND) = NULL;

        if (objc != 0)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1: bfn = IsIconic; break;
        case 2: bfn = IsZoomed; break;
        case 3: bfn = IsWindowVisible; break;
        case 4: bfn = IsWindow; break;
        case 5: bfn = IsWindowUnicode; break;
        case 6: bfn = IsWindowEnabled; break;
        case 7: bfn = ArrangeIconicWindows; break;
        case 8: bfn = SetForegroundWindow; break;
        case 9: dfn = OpenIcon; break;
        case 10: dfn = CloseWindow; break;
        case 11: dfn = DestroyWindow; break;
        case 12: dfn = UpdateWindow; break;
        case 13: dfn = HideCaret; break;
        case 14: dfn = ShowCaret; break;
        case 15:
            result.type = TRT_HWND;
            result.value.hwin = GetParent(hwnd);
            break;
        case 16:
            if (GetWindowInfo(hwnd, &u.wininfo)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromWINDOWINFO(&u.wininfo);
                break;
            } else {
                result.type = TRT_GETLASTERROR;
            }
            break;
        case 17:
            result.type = GetClientRect(hwnd, &result.value.rect) ? TRT_RECT : TRT_GETLASTERROR;
            break;
        case 18:
            result.type = GetWindowRect(hwnd, &result.value.rect) ? TRT_RECT : TRT_GETLASTERROR;
            break;
        case 19:
            result.type = TRT_HDC;
            result.value.hval = GetDC(hwnd);
            break;
        case 20:
            result.type = TRT_HWND;
            result.value.hwin = SetFocus(hwnd);
            break;
        case 21:
            result.value.hwin = SetActiveWindow(hwnd);
            result.type = result.value.hwin ? TRT_HWND : TRT_GETLASTERROR;
            break;
        case 22:
            result.value.unicode.len = GetClassNameW(hwnd, u.buf, ARRAYSIZE(u.buf));
            result.value.unicode.str = u.buf;
            result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
            break;
        case 23:
            result.value.unicode.len = RealGetWindowClassW(hwnd, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
            result.value.unicode.str = u.buf;
            result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
            break;
        case 24:
            return Twapi_EnumChildWindows(interp, hwnd);
        case 25:
            SetLastError(0);            /* Make sure error is not set */
            result.type = TRT_UNICODE;
            result.value.unicode.len = GetWindowTextW(hwnd, u.buf, ARRAYSIZE(u.buf));
            result.value.unicode.str = u.buf;
            /* Distinguish between error and empty string when count is 0 */
            if (result.value.unicode.len == 0 && GetLastError()) {
                result.type = TRT_GETLASTERROR;
            }
            break;
        case 26:
            result.type = TRT_HDC;
            result.value.hval = GetWindowDC(hwnd);
            break;
        }
        if (bfn) {
            result.type = TRT_BOOL;
            result.value.bval = bfn(hwnd);
        } else if (dfn) {
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = dfn(hwnd);
        }
    } else if (func < 1000) {
        /* Exactly one additional int arg */
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 501:
            result.value.hwin = GetAncestor(hwnd, dw);
            result.type = TRT_HWND;
            break;
        case 502:
            result.value.hwin = GetWindow(hwnd, dw);
            result.type = TRT_HWND;
            break;
        case 503:
            result.value.bval = ShowWindow(hwnd, dw);
            result.type = TRT_BOOL;
            break;
        case 504:
            result.value.bval = ShowWindowAsync(hwnd, dw);
            result.type = TRT_BOOL;
            break;
        case 505:
            result.value.bval = EnableWindow(hwnd, dw);
            result.type = TRT_BOOL;
            break;
        case 506:
            result.value.ival = ShowOwnedPopups(hwnd, dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 507:
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromWindow(hwnd, dw);
            break;
        }        
    } else {
        /* At least one additional arg */
        if (objc == 0)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1001: // SetWindowText
            result.value.ival = SetWindowTextW(hwnd, ObjToUnicode(objv[0]));
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 1002: // IsChild
            if (ObjToHWND(interp, objv[0], &hwnd2) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = IsChild(hwnd, hwnd2);
            break;
        case 1003: // UNUSED
            break;
        case 1004: // InvalidateRect
            rectP = &u.rect;
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(rectP, ObjToRECT_NULL), GETINT(dw), 
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = InvalidateRect(hwnd, rectP, dw);
            break;
        case 1005: // SetWindowPos
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLET(hwnd2, HWND), GETINT(dw), GETINT(dw2),
                             GETINT(dw3), GETINT(dw4), GETINT(dw5),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetWindowPos(hwnd, hwnd2, dw, dw2, dw3, dw4, dw5);
            break;
        case 1006:
            if (TwapiGetArgs(interp, objc, objv,
                             GETINT(dw), GETINT(dw2), GETINT(dw3),
                             GETINT(dw4), GETINT(dw5),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = MoveWindow(hwnd, dw, dw2, dw3, dw4, dw5);
            break;
        case 1007:
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLET(h, HDC),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_LONG;
            result.value.ival = ReleaseDC(hwnd, h);
            break;
        case 1008:
            result.type = TRT_HANDLE;
            result.value.hval = OpenThemeData(hwnd, ObjToUnicode(objv[0]));
            break;
        case 1009: // SetWindowRgn
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLET(u.hrgn, HRGN), GETBOOL(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetWindowRgn(hwnd, u.hrgn, dw);
            break;
        case 1010: // GetWindowRgn
            if (TwapiGetArgs(interp, objc, objv, 
                             GETHANDLET(u.hrgn, HRGN),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_LONG;
            result.value.ival = GetWindowRgn(hwnd, u.hrgn);
            break;
        case 1011:
            if (TwapiGetArgs(interp, objc, objv,
                             GETINT(dw), GETINT(dw2), GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetLayeredWindowAttributes(hwnd, dw, (BYTE)dw2, dw3);
            break;
        }
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_UiCallWStructObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    DWORD dw;
    HWND hwnd;
    int func = PtrToInt(clientdata);
    MemLifo *memlifoP;
    MemLifoMarkHandle mark;
    TCL_RESULT res;
    void *pv;
    Tcl_Obj *objP;
    union {
        WINDOWPLACEMENT winplace;
        WINDOWINFO wininfo;
        WCHAR buf[MAX_PATH+1];
        RECT   rect;
        HRGN   hrgn;
    } u;

    /* At least two args - window and another arg */
    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHWND(hwnd), GETOBJ(objP),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    memlifoP = TwapiMemLifo();
    mark = MemLifoPushMark(memlifoP);
    objc -= 3;
    objv += 3;
    result.type = TRT_BADFUNCTIONCODE;

    res = TCL_OK;
    switch (func) {
    case 1: // SetWindowPlacement
        res = ParseCStruct(interp, memlifoP, objP, 0, &dw, &pv);
        if (res == TCL_OK) {
            if (! SetWindowPlacement(hwnd, pv))
                result.type = TRT_EMPTY;
            else
                result.type = TRT_GETLASTERROR;
        } else {
            result.value.ival = res;
            result.type = TRT_TCL_RESULT;
        }
        break;
    case 2: // GetWindowPlacement
        if (GetWindowPlacement(hwnd, &u.winplace)) {
            res = ObjFromCStruct(interp, &u.winplace, sizeof(u.winplace), objP, 0, &result.value.obj);
            if (res == TCL_OK)
                result.type = TRT_OBJ;
            else {
                result.value.ival = res;
                result.type = TRT_TCL_RESULT;
            }
        } else {
            result.type = TRT_GETLASTERROR;
        }
        break;
    }

    res = TwapiSetResult(interp, &result);
    MemLifoPopMark(mark);
    return res;
}

static int TwapiUiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s UiCallDispatch[] = {
        DEFINE_FNCODE_CMD(GetCursorPos, 1),
        DEFINE_FNCODE_CMD(GetCaretPos, 2),
        DEFINE_FNCODE_CMD(GetCaretBlinkTime, 3),
        DEFINE_FNCODE_CMD(GetFocus, 4),                  // TBD Tcl
        DEFINE_FNCODE_CMD(GetDesktopWindow, 5),
        DEFINE_FNCODE_CMD(GetShellWindow, 6),
        DEFINE_FNCODE_CMD(GetForegroundWindow, 7),
        DEFINE_FNCODE_CMD(GetActiveWindow, 8),
        DEFINE_FNCODE_CMD(IsThemeActive, 9),
        DEFINE_FNCODE_CMD(IsAppThemed, 10),
        DEFINE_FNCODE_CMD(GetCurrentThemeName, 11),

        DEFINE_FNCODE_CMD(WindowFromPoint, 1001),
        DEFINE_FNCODE_CMD(FlashWindowEx, 1002),
        DEFINE_FNCODE_CMD(SetCaretBlinkTime, 3001),
        DEFINE_FNCODE_CMD(GetGUIThreadInfo, 3002),
        DEFINE_FNCODE_CMD(SetCaretPos, 3003),
        DEFINE_FNCODE_CMD(SetCursorPos, 3004),
        DEFINE_FNCODE_CMD(CreateRectRgn, 3005),
        DEFINE_FNCODE_CMD(CreateEllipticRgn, 3006),
        DEFINE_FNCODE_CMD(CreateRoundedRectRgn, 3007),
        DEFINE_FNCODE_CMD(CloseThemeData, 4001),
        DEFINE_FNCODE_CMD(EnumDesktopWindows, 4002),
        DEFINE_FNCODE_CMD(GetMonitorInfo, 4003),
        DEFINE_FNCODE_CMD(GetThemeSysColor, 4004), // TBD - tcl
        DEFINE_FNCODE_CMD(GetThemeSysFont, 4005),  // TBD - tcl
        DEFINE_FNCODE_CMD(GetObject, 4006), // TBD - tcl
        DEFINE_FNCODE_CMD(GetDeviceCaps, 4007),
        DEFINE_FNCODE_CMD(MonitorFromPoint, 10001),
        DEFINE_FNCODE_CMD(MonitorFromRect, 10002),
        DEFINE_FNCODE_CMD(GetThemeColor, 10003),
        DEFINE_FNCODE_CMD(GetThemeFont, 10004),
        DEFINE_FNCODE_CMD(EnumDisplayDevices, 10005),
        DEFINE_FNCODE_CMD(EnumDisplayMonitors, 10006),
        DEFINE_FNCODE_CMD(FindWindow, 10007),
        DEFINE_FNCODE_CMD(AddFontResourceEx, 10008), // TBD - Tcl
        DEFINE_FNCODE_CMD(RemoveFontResourceEx, 10009), // TBD - Tcl
    };

    static struct fncode_dispatch_s UiCallWDispatch[] = {
        DEFINE_FNCODE_CMD(IsIconic, 1),
        DEFINE_FNCODE_CMD(IsZoomed, 2),
        DEFINE_FNCODE_CMD(IsWindowVisible, 3),
        DEFINE_FNCODE_CMD(IsWindow, 4),
        DEFINE_FNCODE_CMD(IsWindowUnicode, 5),
        DEFINE_FNCODE_CMD(IsWindowEnabled, 6),
        DEFINE_FNCODE_CMD(ArrangeIconicWindows, 7),
        DEFINE_FNCODE_CMD(SetForegroundWindow, 8),
        DEFINE_FNCODE_CMD(OpenIcon, 9),
        DEFINE_FNCODE_CMD(CloseWindow, 10),
        DEFINE_FNCODE_CMD(DestroyWindow, 11),
        DEFINE_FNCODE_CMD(UpdateWindow, 12),
        DEFINE_FNCODE_CMD(HideCaret, 13),
        DEFINE_FNCODE_CMD(ShowCaret, 14),
        DEFINE_FNCODE_CMD(GetParent, 15),
        DEFINE_FNCODE_CMD(GetWindowInfo, 16), // TBD - Tcl 
        DEFINE_FNCODE_CMD(GetClientRect, 17),
        DEFINE_FNCODE_CMD(GetWindowRect, 18),
        DEFINE_FNCODE_CMD(GetDC, 19),
        DEFINE_FNCODE_CMD(SetFocus, 20),
        DEFINE_FNCODE_CMD(SetActiveWindow, 21),
        DEFINE_FNCODE_CMD(GetClassName, 22),
        DEFINE_FNCODE_CMD(RealGetWindowClass, 23),
        DEFINE_FNCODE_CMD(EnumChildWindows, 24),
        DEFINE_FNCODE_CMD(GetWindowText, 25),
        DEFINE_FNCODE_CMD(GetWindowDC, 26),
        DEFINE_FNCODE_CMD(GetAncestor, 501),
        DEFINE_FNCODE_CMD(GetWindow, 502),
        DEFINE_FNCODE_CMD(ShowWindow, 503),
        DEFINE_FNCODE_CMD(ShowWindowAsync, 504),
        DEFINE_FNCODE_CMD(EnableWindow, 505),
        DEFINE_FNCODE_CMD(ShowOwnedPopups, 506),
        DEFINE_FNCODE_CMD(MonitorFromWindow, 507),
        DEFINE_FNCODE_CMD(SetWindowText, 1001),
        DEFINE_FNCODE_CMD(IsChild, 1002),
        DEFINE_FNCODE_CMD(InvalidateRect, 1004),     // TBD - Tcl 
        DEFINE_FNCODE_CMD(SetWindowPos, 1005),
        DEFINE_FNCODE_CMD(MoveWindow, 1006),
        DEFINE_FNCODE_CMD(ReleaseDC, 1007),
        DEFINE_FNCODE_CMD(OpenThemeData, 1008),
        DEFINE_FNCODE_CMD(SetWindowRgn, 1009), // TBD - Tcl 
        DEFINE_FNCODE_CMD(GetWindowRgn, 1010), // TBD - Tcl
        DEFINE_FNCODE_CMD(SetLayeredWindowAttributes, 1011),
    };

    /* Command using memlifo */
    static struct fncode_dispatch_s UiCallWStructDispatch[] = {
        DEFINE_FNCODE_CMD(SetWindowPlacement, 1), // TBD - Tcl 
        DEFINE_FNCODE_CMD(GetWindowPlacement, 2), // TBD - Tcl 
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(UiCallDispatch), UiCallDispatch, Twapi_UiCallObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(UiCallWDispatch), UiCallWDispatch, Twapi_UiCallWObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(UiCallWStructDispatch), UiCallWStructDispatch, Twapi_UiCallWStructObjCmd);

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
int Twapi_ui_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiUiInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

