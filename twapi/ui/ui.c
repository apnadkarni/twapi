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

#ifndef TWAPI_STATIC_BUILD
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

/* Return Tcl List for a LOGFONT structure */
Tcl_Obj *ObjFromLOGFONTW(LOGFONTW *lfP)
{
    Tcl_Obj *objv[28];

    objv[0] = STRING_LITERAL_OBJ("lfHeight");
    objv[1] = Tcl_NewLongObj(lfP->lfHeight);
    objv[2] = STRING_LITERAL_OBJ("lfWidth");
    objv[3] = Tcl_NewLongObj(lfP->lfWidth);
    objv[4] = STRING_LITERAL_OBJ("lfEscapement");
    objv[5] = Tcl_NewLongObj(lfP->lfEscapement);
    objv[6] = STRING_LITERAL_OBJ("lfOrientation");
    objv[7] = Tcl_NewLongObj(lfP->lfOrientation);
    objv[8] = STRING_LITERAL_OBJ("lfWeight");
    objv[9] = Tcl_NewLongObj(lfP->lfWeight);
    objv[10] = STRING_LITERAL_OBJ("lfItalic");
    objv[11] = Tcl_NewLongObj(lfP->lfItalic);
    objv[12] = STRING_LITERAL_OBJ("lfUnderline");
    objv[13] = Tcl_NewLongObj(lfP->lfUnderline);
    objv[14] = STRING_LITERAL_OBJ("lfStrikeOut");
    objv[15] = Tcl_NewLongObj(lfP->lfStrikeOut);
    objv[16] = STRING_LITERAL_OBJ("lfCharSet");
    objv[17] = Tcl_NewLongObj(lfP->lfCharSet);
    objv[18] = STRING_LITERAL_OBJ("lfOutPrecision");
    objv[19] = Tcl_NewLongObj(lfP->lfOutPrecision);
    objv[20] = STRING_LITERAL_OBJ("lfClipPrecision");
    objv[21] = Tcl_NewLongObj(lfP->lfClipPrecision);
    objv[22] = STRING_LITERAL_OBJ("lfQuality");
    objv[23] = Tcl_NewLongObj(lfP->lfQuality);
    objv[24] = STRING_LITERAL_OBJ("lfPitchAndFamily");
    objv[25] = Tcl_NewLongObj(lfP->lfPitchAndFamily);
    objv[26] = STRING_LITERAL_OBJ("lfFaceName");
    objv[27] = ObjFromUnicode(lfP->lfFaceName);

    return Tcl_NewListObj(28, objv);
}


int ObjToFLASHWINFO (Tcl_Interp *interp, Tcl_Obj *obj, FLASHWINFO *fwP)
{
    Tcl_Obj **objv;
    int       objc;
    HWND      hwnd;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    fwP->cbSize = sizeof(*fwP);
    if (objc == 4 &&
        (ObjToHWND(interp, objv[0], &hwnd) == TCL_OK) &&
        (Tcl_GetLongFromObj(interp, objv[1], &fwP->dwFlags) == TCL_OK) &&
        (Tcl_GetLongFromObj(interp, objv[2], &fwP->uCount) == TCL_OK) &&
        (Tcl_GetLongFromObj(interp, objv[3], &fwP->dwTimeout) == TCL_OK)) {
        fwP->hwnd = (HWND) hwnd;
        return TCL_OK;
    }

    Tcl_SetResult(interp, "Invalid FLASHWINFO structure. Need 4 elements - window handle, flags, count and timeout.", TCL_STATIC);
    return TCL_ERROR;
}

Tcl_Obj *ObjFromWINDOWPLACEMENT(WINDOWPLACEMENT *wpP)
{
    Tcl_Obj *objv[5];
    objv[0] = Tcl_NewLongObj(wpP->flags);
    objv[1] = Tcl_NewLongObj(wpP->showCmd);
    objv[2] = ObjFromPOINT(&wpP->ptMinPosition);
    objv[3] = ObjFromPOINT(&wpP->ptMaxPosition);
    objv[4] = ObjFromRECT(&wpP->rcNormalPosition);
    return Tcl_NewListObj(5, objv);
}

int ObjToWINDOWPLACEMENT(Tcl_Interp *interp, Tcl_Obj *objP, WINDOWPLACEMENT *wpP)
{
    Tcl_Obj **objv;
    int objc;

    if (Tcl_ListObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc != 5)
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Incorrect format of WINDOWPLACEMENT argument.");

    if (Tcl_GetLongFromObj(interp, objv[0], &wpP->flags) != TCL_OK ||
        Tcl_GetLongFromObj(interp, objv[1], &wpP->showCmd) != TCL_OK ||
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
    objv[2] = Tcl_NewLongObj(wiP->dwStyle);
    objv[3] = Tcl_NewLongObj(wiP->dwExStyle);
    objv[4] = Tcl_NewLongObj(wiP->dwWindowStatus);
    objv[5] = Tcl_NewLongObj(wiP->cxWindowBorders);
    objv[6] = Tcl_NewLongObj(wiP->cyWindowBorders);
    objv[7] = Tcl_NewLongObj(wiP->atomWindowType);
    objv[8] = Tcl_NewLongObj(wiP->wCreatorVersion);

    return Tcl_NewListObj(9, objv);
}

/* Window enumeration callback */
BOOL CALLBACK Twapi_EnumWindowsCallback(HWND hwnd, LPARAM p_ctx) {
    TwapiEnumCtx *p_enum_win_ctx =
        (TwapiEnumCtx *) p_ctx;

    Tcl_ListObjAppendElement(p_enum_win_ctx->interp,
                             p_enum_win_ctx->objP,
                             ObjFromHWND(hwnd));
    
    return 1;
}


/*
 * Enumerate toplevel windows
 */
int Twapi_EnumWindows(Tcl_Interp *interp)
{
    TwapiEnumCtx enum_win_ctx;

    enum_win_ctx.interp = interp;
    enum_win_ctx.objP = Tcl_NewListObj(0, NULL);

    
    if (EnumWindows(Twapi_EnumWindowsCallback, (LPARAM)&enum_win_ctx) == 0) {
        TwapiReturnSystemError(interp);
        Twapi_FreeNewTclObj(enum_win_ctx.objP);
        return TCL_ERROR;
    }

    Tcl_SetObjResult(interp, enum_win_ctx.objP);
    return TCL_OK;
}


/*
 * Enumerate child windows of windows with handle parent_handle
 */
int Twapi_EnumChildWindows(Tcl_Interp *interp, HWND parent_handle)
{
    TwapiEnumCtx enum_win_ctx;

    enum_win_ctx.interp = interp;
    enum_win_ctx.objP = Tcl_NewListObj(0, NULL);
    
    /* As per newer MSDN docs, EnumChildWindows return value is meaningless.
     * We used to check GetLastError ERROR_PROC_NOT_FOUND and
     * ERROR_MOD_NOT_FOUND but
     * on Win2008R2, it returns ERROR_FILE_NOT_FOUND if no child windows.
     * We now follow the docs and ignore return value.
     */
    EnumChildWindows(parent_handle, Twapi_EnumWindowsCallback, (LPARAM)&enum_win_ctx);
    Tcl_SetObjResult(interp, enum_win_ctx.objP);
    return TCL_OK;
}

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


/* Enumerate desktop windows
 */
int Twapi_EnumDesktopWindows(Tcl_Interp *interp, HDESK desk_handle)
{
    TwapiEnumCtx enum_win_ctx;

    enum_win_ctx.interp = interp;
    enum_win_ctx.objP = Tcl_NewListObj(0, NULL);
    
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

    Tcl_SetObjResult(interp, enum_win_ctx.objP);
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
    objv[1] = Tcl_NewLongObj(gti.cbSize);
    objv[2] = STRING_LITERAL_OBJ("flags");
    objv[3] = Tcl_NewLongObj(gti.flags);
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

    Tcl_SetObjResult(interp, Tcl_NewListObj(18, objv));
    return TCL_OK;
}


int TwapiGetThemeDefine(Tcl_Interp *interp, char *name)
{
    int val;

    if (name[0] == 0) {
        val = 0;
        goto success_return;
    }

#define cmp_and_return_(def) \
    do { \
        if ( (# def)[0] == name[0] && \
             (lstrcmpi(# def, name) == 0) ) { \
            val = def; \
            goto success_return; \
        } \
    } while (0)

    // TBD - make more efficient by using Tcl hash tables
    // NOTE: some values have been commented out as they are not in my header
    // files though they are listed in MSDN

    cmp_and_return_(BP_CHECKBOX);
    cmp_and_return_(CBS_CHECKEDDISABLED);
    cmp_and_return_(CBS_CHECKEDHOT);
    cmp_and_return_(CBS_CHECKEDNORMAL);
    cmp_and_return_(CBS_CHECKEDPRESSED);
    cmp_and_return_(CBS_MIXEDDISABLED);
    cmp_and_return_(CBS_MIXEDHOT);
    cmp_and_return_(CBS_MIXEDNORMAL);
    cmp_and_return_(CBS_MIXEDPRESSED);
    cmp_and_return_(CBS_UNCHECKEDDISABLED);
    cmp_and_return_(CBS_UNCHECKEDHOT);
    cmp_and_return_(CBS_UNCHECKEDNORMAL);
    cmp_and_return_(CBS_UNCHECKEDPRESSED);
    cmp_and_return_(BP_GROUPBOX);
    cmp_and_return_(GBS_DISABLED);
    cmp_and_return_(GBS_NORMAL);
    cmp_and_return_(BP_PUSHBUTTON);
    cmp_and_return_(PBS_DEFAULTED);
    cmp_and_return_(PBS_DISABLED);
    cmp_and_return_(PBS_HOT);
    cmp_and_return_(PBS_NORMAL);
    cmp_and_return_(PBS_PRESSED);
    cmp_and_return_(BP_RADIOBUTTON);
    cmp_and_return_(RBS_CHECKEDDISABLED);
    cmp_and_return_(RBS_CHECKEDHOT);
    cmp_and_return_(RBS_CHECKEDNORMAL);
    cmp_and_return_(RBS_CHECKEDPRESSED);
    cmp_and_return_(RBS_UNCHECKEDDISABLED);
    cmp_and_return_(RBS_UNCHECKEDHOT);
    cmp_and_return_(RBS_UNCHECKEDNORMAL);
    cmp_and_return_(RBS_UNCHECKEDPRESSED);
    cmp_and_return_(BP_USERBUTTON);
    cmp_and_return_(CLP_TIME);
    cmp_and_return_(CLS_NORMAL);
    cmp_and_return_(CP_DROPDOWNBUTTON);
    cmp_and_return_(CBXS_DISABLED);
    cmp_and_return_(CBXS_HOT);
    cmp_and_return_(CBXS_NORMAL);
    cmp_and_return_(CBXS_PRESSED);
    cmp_and_return_(EP_CARET);
    cmp_and_return_(EP_EDITTEXT);
    cmp_and_return_(ETS_ASSIST);
    cmp_and_return_(ETS_DISABLED);
    cmp_and_return_(ETS_FOCUSED);
    cmp_and_return_(ETS_HOT);
    cmp_and_return_(ETS_NORMAL);
    cmp_and_return_(ETS_READONLY);
    cmp_and_return_(ETS_SELECTED);
    cmp_and_return_(EBP_HEADERBACKGROUND);
    cmp_and_return_(EBP_HEADERCLOSE);
    cmp_and_return_(EBHC_HOT);
    cmp_and_return_(EBHC_NORMAL);
    cmp_and_return_(EBHC_PRESSED);
    cmp_and_return_(EBP_HEADERPIN);
    cmp_and_return_(EBHP_HOT);
    cmp_and_return_(EBHP_NORMAL);
    cmp_and_return_(EBHP_PRESSED);
    cmp_and_return_(EBHP_SELECTEDHOT);
    cmp_and_return_(EBHP_SELECTEDNORMAL);
    cmp_and_return_(EBHP_SELECTEDPRESSED);
    cmp_and_return_(EBP_IEBARMENU);
    cmp_and_return_(EBM_HOT);
    cmp_and_return_(EBM_NORMAL);
    cmp_and_return_(EBM_PRESSED);
    cmp_and_return_(EBP_NORMALGROUPBACKGROUND);
    cmp_and_return_(EBP_NORMALGROUPCOLLAPSE);
    cmp_and_return_(EBNGC_HOT);
    cmp_and_return_(EBNGC_NORMAL);
    cmp_and_return_(EBNGC_PRESSED);
    cmp_and_return_(EBP_NORMALGROUPEXPAND);
    cmp_and_return_(EBNGE_HOT);
    cmp_and_return_(EBNGE_NORMAL);
    cmp_and_return_(EBNGE_PRESSED);
    cmp_and_return_(EBP_NORMALGROUPHEAD);
    cmp_and_return_(EBP_SPECIALGROUPBACKGROUND);
    cmp_and_return_(EBP_SPECIALGROUPCOLLAPSE);
    cmp_and_return_(EBSGC_HOT);
    cmp_and_return_(EBSGC_NORMAL);
    cmp_and_return_(EBSGC_PRESSED);
    cmp_and_return_(EBP_SPECIALGROUPEXPAND);
    cmp_and_return_(EBSGE_HOT);
    cmp_and_return_(EBSGE_NORMAL);
    cmp_and_return_(EBSGE_PRESSED);
    cmp_and_return_(EBP_SPECIALGROUPHEAD);
    //cmp_and_return_(GP_BORDER);
    //cmp_and_return_(BSS_FLAT);
    //cmp_and_return_(BSS_RAISED);
    //cmp_and_return_(BSS_SUNKEN);
    //cmp_and_return_(GP_LINEHORZ);
    //cmp_and_return_(LHS_FLAT);
    //cmp_and_return_(LHS_RAISED);
    //cmp_and_return_(LHS_SUNKEN);
    //cmp_and_return_(GP_LINEVERT);
    //cmp_and_return_(LVS_FLAT);
    //cmp_and_return_(LVS_RAISED);
    //cmp_and_return_(LVS_SUNKEN);
    cmp_and_return_(HP_HEADERITEM);
    cmp_and_return_(HIS_HOT);
    cmp_and_return_(HIS_NORMAL);
    cmp_and_return_(HIS_PRESSED);
    cmp_and_return_(HP_HEADERITEMLEFT);
    cmp_and_return_(HILS_HOT);
    cmp_and_return_(HILS_NORMAL);
    cmp_and_return_(HILS_PRESSED);
    cmp_and_return_(HP_HEADERITEMRIGHT);
    cmp_and_return_(HIRS_HOT);
    cmp_and_return_(HIRS_NORMAL);
    cmp_and_return_(HIRS_PRESSED);
    cmp_and_return_(HP_HEADERSORTARROW);
    cmp_and_return_(HSAS_SORTEDDOWN);
    cmp_and_return_(HSAS_SORTEDUP);
    cmp_and_return_(LVP_EMPTYTEXT);
    cmp_and_return_(LVP_LISTDETAIL);
    cmp_and_return_(LVP_LISTGROUP);
    cmp_and_return_(LVP_LISTITEM);
    cmp_and_return_(LIS_DISABLED);
    cmp_and_return_(LIS_HOT);
    cmp_and_return_(LIS_NORMAL);
    cmp_and_return_(LIS_SELECTED);
    cmp_and_return_(LIS_SELECTEDNOTFOCUS);
    cmp_and_return_(LVP_LISTSORTEDDETAIL);
    cmp_and_return_(MP_MENUBARDROPDOWN);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_MENUBARITEM);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_CHEVRON);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_MENUDROPDOWN);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_MENUITEM);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MP_SEPARATOR);
    cmp_and_return_(MS_DEMOTED);
    cmp_and_return_(MS_NORMAL);
    cmp_and_return_(MS_SELECTED);
    cmp_and_return_(MDP_NEWAPPBUTTON);
    cmp_and_return_(MDS_CHECKED);
    cmp_and_return_(MDS_DISABLED);
    cmp_and_return_(MDS_HOT);
    cmp_and_return_(MDS_HOTCHECKED);
    cmp_and_return_(MDS_NORMAL);
    cmp_and_return_(MDS_PRESSED);
    cmp_and_return_(MDP_SEPERATOR);
    cmp_and_return_(PGRP_DOWN);
    cmp_and_return_(DNS_DISABLED);
    cmp_and_return_(DNS_HOT);
    cmp_and_return_(DNS_NORMAL);
    cmp_and_return_(DNS_PRESSED);
    cmp_and_return_(PGRP_DOWNHORZ);
    cmp_and_return_(DNHZS_DISABLED);
    cmp_and_return_(DNHZS_HOT);
    cmp_and_return_(DNHZS_NORMAL);
    cmp_and_return_(DNHZS_PRESSED);
    cmp_and_return_(PGRP_UP);
    cmp_and_return_(UPS_DISABLED);
    cmp_and_return_(UPS_HOT);
    cmp_and_return_(UPS_NORMAL);
    cmp_and_return_(UPS_PRESSED);
    cmp_and_return_(PGRP_UPHORZ);
    cmp_and_return_(UPHZS_DISABLED);
    cmp_and_return_(UPHZS_HOT);
    cmp_and_return_(UPHZS_NORMAL);
    cmp_and_return_(UPHZS_PRESSED);
    cmp_and_return_(PP_BAR);
    cmp_and_return_(PP_BARVERT);
    cmp_and_return_(PP_CHUNK);
    cmp_and_return_(PP_CHUNKVERT);
    cmp_and_return_(RP_BAND);
    cmp_and_return_(RP_CHEVRON);
    cmp_and_return_(CHEVS_HOT);
    cmp_and_return_(CHEVS_NORMAL);
    cmp_and_return_(CHEVS_PRESSED);
    cmp_and_return_(RP_CHEVRONVERT);
    cmp_and_return_(RP_GRIPPER);
    cmp_and_return_(RP_GRIPPERVERT);
    cmp_and_return_(SBP_ARROWBTN);
    cmp_and_return_(ABS_DOWNDISABLED);
    cmp_and_return_(ABS_DOWNHOT);
    cmp_and_return_(ABS_DOWNNORMAL);
    cmp_and_return_(ABS_DOWNPRESSED);
    cmp_and_return_(ABS_UPDISABLED);
    cmp_and_return_(ABS_UPHOT);
    cmp_and_return_(ABS_UPNORMAL);
    cmp_and_return_(ABS_UPPRESSED);
    cmp_and_return_(ABS_LEFTDISABLED);
    cmp_and_return_(ABS_LEFTHOT);
    cmp_and_return_(ABS_LEFTNORMAL);
    cmp_and_return_(ABS_LEFTPRESSED);
    cmp_and_return_(ABS_RIGHTDISABLED);
    cmp_and_return_(ABS_RIGHTHOT);
    cmp_and_return_(ABS_RIGHTNORMAL);
    cmp_and_return_(ABS_RIGHTPRESSED);
    cmp_and_return_(SBP_GRIPPERHORZ);
    cmp_and_return_(SBP_GRIPPERVERT);
    cmp_and_return_(SBP_LOWERTRACKHORZ);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_LOWERTRACKVERT);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_THUMBBTNHORZ);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_THUMBBTNVERT);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_UPPERTRACKHORZ);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_UPPERTRACKVERT);
    cmp_and_return_(SCRBS_DISABLED);
    cmp_and_return_(SCRBS_HOT);
    cmp_and_return_(SCRBS_NORMAL);
    cmp_and_return_(SCRBS_PRESSED);
    cmp_and_return_(SBP_SIZEBOX);
    cmp_and_return_(SZB_LEFTALIGN);
    cmp_and_return_(SZB_RIGHTALIGN);
    cmp_and_return_(SPNP_DOWN);
    cmp_and_return_(DNS_DISABLED);
    cmp_and_return_(DNS_HOT);
    cmp_and_return_(DNS_NORMAL);
    cmp_and_return_(DNS_PRESSED);
    cmp_and_return_(SPNP_DOWNHORZ);
    cmp_and_return_(DNHZS_DISABLED);
    cmp_and_return_(DNHZS_HOT);
    cmp_and_return_(DNHZS_NORMAL);
    cmp_and_return_(DNHZS_PRESSED);
    cmp_and_return_(SPNP_UP);
    cmp_and_return_(UPS_DISABLED);
    cmp_and_return_(UPS_HOT);
    cmp_and_return_(UPS_NORMAL);
    cmp_and_return_(UPS_PRESSED);
    cmp_and_return_(SPNP_UPHORZ);
    cmp_and_return_(UPHZS_DISABLED);
    cmp_and_return_(UPHZS_HOT);
    cmp_and_return_(UPHZS_NORMAL);
    cmp_and_return_(UPHZS_PRESSED);
    cmp_and_return_(SPP_LOGOFF);
    cmp_and_return_(SPP_LOGOFFBUTTONS);
    cmp_and_return_(SPLS_HOT);
    cmp_and_return_(SPLS_NORMAL);
    cmp_and_return_(SPLS_PRESSED);
    cmp_and_return_(SPP_MOREPROGRAMS);
    cmp_and_return_(SPP_MOREPROGRAMSARROW);
    cmp_and_return_(SPS_HOT);
    cmp_and_return_(SPS_NORMAL);
    cmp_and_return_(SPS_PRESSED);
    cmp_and_return_(SPP_PLACESLIST);
    cmp_and_return_(SPP_PLACESLISTSEPARATOR);
    cmp_and_return_(SPP_PREVIEW);
    cmp_and_return_(SPP_PROGLIST);
    cmp_and_return_(SPP_PROGLISTSEPARATOR);
    cmp_and_return_(SPP_USERPANE);
    cmp_and_return_(SPP_USERPICTURE);
    cmp_and_return_(SP_GRIPPER);
    cmp_and_return_(SP_PANE);
    cmp_and_return_(SP_GRIPPERPANE);
    cmp_and_return_(TABP_BODY);
    cmp_and_return_(TABP_PANE);
    cmp_and_return_(TABP_TABITEM);
    cmp_and_return_(TIS_DISABLED);
    cmp_and_return_(TIS_FOCUSED);
    cmp_and_return_(TIS_HOT);
    cmp_and_return_(TIS_NORMAL);
    cmp_and_return_(TIS_SELECTED);
    cmp_and_return_(TABP_TABITEMBOTHEDGE);
    cmp_and_return_(TIBES_DISABLED);
    cmp_and_return_(TIBES_FOCUSED);
    cmp_and_return_(TIBES_HOT);
    cmp_and_return_(TIBES_NORMAL);
    cmp_and_return_(TIBES_SELECTED);
    cmp_and_return_(TABP_TABITEMLEFTEDGE);
    cmp_and_return_(TILES_DISABLED);
    cmp_and_return_(TILES_FOCUSED);
    cmp_and_return_(TILES_HOT);
    cmp_and_return_(TILES_NORMAL);
    cmp_and_return_(TILES_SELECTED);
    cmp_and_return_(TABP_TABITEMRIGHTEDGE);
    cmp_and_return_(TIRES_DISABLED);
    cmp_and_return_(TIRES_FOCUSED);
    cmp_and_return_(TIRES_HOT);
    cmp_and_return_(TIRES_NORMAL);
    cmp_and_return_(TIRES_SELECTED);
    cmp_and_return_(TABP_TOPTABITEM);
    cmp_and_return_(TTIS_DISABLED);
    cmp_and_return_(TTIS_FOCUSED);
    cmp_and_return_(TTIS_HOT);
    cmp_and_return_(TTIS_NORMAL);
    cmp_and_return_(TTIS_SELECTED);
    cmp_and_return_(TABP_TOPTABITEMBOTHEDGE);
    cmp_and_return_(TTIBES_DISABLED);
    cmp_and_return_(TTIBES_FOCUSED);
    cmp_and_return_(TTIBES_HOT);
    cmp_and_return_(TTIBES_NORMAL);
    cmp_and_return_(TTIBES_SELECTED);
    cmp_and_return_(TABP_TOPTABITEMLEFTEDGE);
    cmp_and_return_(TTILES_DISABLED);
    cmp_and_return_(TTILES_FOCUSED);
    cmp_and_return_(TTILES_HOT);
    cmp_and_return_(TTILES_NORMAL);
    cmp_and_return_(TTILES_SELECTED);
    cmp_and_return_(TABP_TOPTABITEMRIGHTEDGE);
    cmp_and_return_(TTIRES_DISABLED);
    cmp_and_return_(TTIRES_FOCUSED);
    cmp_and_return_(TTIRES_HOT);
    cmp_and_return_(TTIRES_NORMAL);
    cmp_and_return_(TTIRES_SELECTED);
    cmp_and_return_(TDP_GROUPCOUNT);
    cmp_and_return_(TDP_FLASHBUTTON);
    cmp_and_return_(TDP_FLASHBUTTONGROUPMENU);
    cmp_and_return_(TBP_BACKGROUNDBOTTOM);
    cmp_and_return_(TBP_BACKGROUNDLEFT);
    cmp_and_return_(TBP_BACKGROUNDRIGHT);
    cmp_and_return_(TBP_BACKGROUNDTOP);
    cmp_and_return_(TBP_SIZINGBARBOTTOM);
    //cmp_and_return_(TBP_SIZINGBARBOTTOMLEFT);
    cmp_and_return_(TBP_SIZINGBARRIGHT);
    cmp_and_return_(TBP_SIZINGBARTOP);
    cmp_and_return_(TP_BUTTON);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_DROPDOWNBUTTON);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SPLITBUTTON);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SPLITBUTTONDROPDOWN);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SEPARATOR);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TP_SEPARATORVERT);
    cmp_and_return_(TS_CHECKED);
    cmp_and_return_(TS_DISABLED);
    cmp_and_return_(TS_HOT);
    cmp_and_return_(TS_HOTCHECKED);
    cmp_and_return_(TS_NORMAL);
    cmp_and_return_(TS_PRESSED);
    cmp_and_return_(TTP_BALLOON);
    cmp_and_return_(TTBS_LINK);
    cmp_and_return_(TTBS_NORMAL);
    cmp_and_return_(TTP_BALLOONTITLE);
    cmp_and_return_(TTBS_LINK);
    cmp_and_return_(TTBS_NORMAL);
    cmp_and_return_(TTP_CLOSE);
    cmp_and_return_(TTCS_HOT);
    cmp_and_return_(TTCS_NORMAL);
    cmp_and_return_(TTCS_PRESSED);
    cmp_and_return_(TTP_STANDARD);
    cmp_and_return_(TTSS_LINK);
    cmp_and_return_(TTSS_NORMAL);
    cmp_and_return_(TTP_STANDARDTITLE);
    cmp_and_return_(TTSS_LINK);
    cmp_and_return_(TTSS_NORMAL);
    cmp_and_return_(TKP_THUMB);
    cmp_and_return_(TUS_DISABLED);
    cmp_and_return_(TUS_FOCUSED);
    cmp_and_return_(TUS_HOT);
    cmp_and_return_(TUS_NORMAL);
    cmp_and_return_(TUS_PRESSED);
    cmp_and_return_(TKP_THUMBBOTTOM);
    cmp_and_return_(TUBS_DISABLED);
    cmp_and_return_(TUBS_FOCUSED);
    cmp_and_return_(TUBS_HOT);
    cmp_and_return_(TUBS_NORMAL);
    cmp_and_return_(TUBS_PRESSED);
    cmp_and_return_(TKP_THUMBLEFT);
    cmp_and_return_(TUVLS_DISABLED);
    cmp_and_return_(TUVLS_FOCUSED);
    cmp_and_return_(TUVLS_HOT);
    cmp_and_return_(TUVLS_NORMAL);
    cmp_and_return_(TUVLS_PRESSED);
    cmp_and_return_(TKP_THUMBRIGHT);
    cmp_and_return_(TUVRS_DISABLED);
    cmp_and_return_(TUVRS_FOCUSED);
    cmp_and_return_(TUVRS_HOT);
    cmp_and_return_(TUVRS_NORMAL);
    cmp_and_return_(TUVRS_PRESSED);
    cmp_and_return_(TKP_THUMBTOP);
    cmp_and_return_(TUTS_DISABLED);
    cmp_and_return_(TUTS_FOCUSED);
    cmp_and_return_(TUTS_HOT);
    cmp_and_return_(TUTS_NORMAL);
    cmp_and_return_(TUTS_PRESSED);
    cmp_and_return_(TKP_THUMBVERT);
    cmp_and_return_(TUVS_DISABLED);
    cmp_and_return_(TUVS_FOCUSED);
    cmp_and_return_(TUVS_HOT);
    cmp_and_return_(TUVS_NORMAL);
    cmp_and_return_(TUVS_PRESSED);
    cmp_and_return_(TKP_TICS);
    cmp_and_return_(TSS_NORMAL);
    cmp_and_return_(TKP_TICSVERT);
    cmp_and_return_(TSVS_NORMAL);
    cmp_and_return_(TKP_TRACK);
    cmp_and_return_(TRS_NORMAL);
    cmp_and_return_(TKP_TRACKVERT);
    cmp_and_return_(TRVS_NORMAL);
    cmp_and_return_(TNP_ANIMBACKGROUND);
    cmp_and_return_(TNP_BACKGROUND);
    cmp_and_return_(TVP_BRANCH);
    cmp_and_return_(TVP_GLYPH);
    cmp_and_return_(GLPS_CLOSED);
    cmp_and_return_(GLPS_OPENED);
    cmp_and_return_(TVP_TREEITEM);
    cmp_and_return_(TREIS_DISABLED);
    cmp_and_return_(TREIS_HOT);
    cmp_and_return_(TREIS_NORMAL);
    cmp_and_return_(TREIS_SELECTED);
    cmp_and_return_(TREIS_SELECTEDNOTFOCUS);
    cmp_and_return_(WP_CAPTION);
    cmp_and_return_(CS_ACTIVE);
    cmp_and_return_(CS_DISABLED);
    cmp_and_return_(CS_INACTIVE);
    cmp_and_return_(WP_CAPTIONSIZINGTEMPLATE);
    cmp_and_return_(WP_CLOSEBUTTON);
    cmp_and_return_(CBS_DISABLED);
    cmp_and_return_(CBS_HOT);
    cmp_and_return_(CBS_NORMAL);
    cmp_and_return_(CBS_PUSHED);
    cmp_and_return_(WP_DIALOG);
    cmp_and_return_(WP_FRAMEBOTTOM);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_FRAMEBOTTOMSIZINGTEMPLATE);
    cmp_and_return_(WP_FRAMELEFT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_FRAMELEFTSIZINGTEMPLATE);
    cmp_and_return_(WP_FRAMERIGHT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_FRAMERIGHTSIZINGTEMPLATE);
    cmp_and_return_(WP_HELPBUTTON);
    cmp_and_return_(HBS_DISABLED);
    cmp_and_return_(HBS_HOT);
    cmp_and_return_(HBS_NORMAL);
    cmp_and_return_(HBS_PUSHED);
    cmp_and_return_(WP_HORZSCROLL);
    cmp_and_return_(HSS_DISABLED);
    cmp_and_return_(HSS_HOT);
    cmp_and_return_(HSS_NORMAL);
    cmp_and_return_(HSS_PUSHED);
    cmp_and_return_(WP_HORZTHUMB);
    cmp_and_return_(HTS_DISABLED);
    cmp_and_return_(HTS_HOT);
    cmp_and_return_(HTS_NORMAL);
    cmp_and_return_(HTS_PUSHED);
    //cmp_and_return_(WP_MAX_BUTTON);
    cmp_and_return_(MAXBS_DISABLED);
    cmp_and_return_(MAXBS_HOT);
    cmp_and_return_(MAXBS_NORMAL);
    cmp_and_return_(MAXBS_PUSHED);
    cmp_and_return_(WP_MAXCAPTION);
    cmp_and_return_(MXCS_ACTIVE);
    cmp_and_return_(MXCS_DISABLED);
    cmp_and_return_(MXCS_INACTIVE);
    cmp_and_return_(WP_MDICLOSEBUTTON);
    cmp_and_return_(CBS_DISABLED);
    cmp_and_return_(CBS_HOT);
    cmp_and_return_(CBS_NORMAL);
    cmp_and_return_(CBS_PUSHED);
    cmp_and_return_(WP_MDIHELPBUTTON);
    cmp_and_return_(HBS_DISABLED);
    cmp_and_return_(HBS_HOT);
    cmp_and_return_(HBS_NORMAL);
    cmp_and_return_(HBS_PUSHED);
    cmp_and_return_(WP_MDIMINBUTTON);
    cmp_and_return_(MINBS_DISABLED);
    cmp_and_return_(MINBS_HOT);
    cmp_and_return_(MINBS_NORMAL);
    cmp_and_return_(MINBS_PUSHED);
    cmp_and_return_(WP_MDIRESTOREBUTTON);
    cmp_and_return_(RBS_DISABLED);
    cmp_and_return_(RBS_HOT);
    cmp_and_return_(RBS_NORMAL);
    cmp_and_return_(RBS_PUSHED);
    cmp_and_return_(WP_MDISYSBUTTON);
    cmp_and_return_(SBS_DISABLED);
    cmp_and_return_(SBS_HOT);
    cmp_and_return_(SBS_NORMAL);
    cmp_and_return_(SBS_PUSHED);
    cmp_and_return_(WP_MINBUTTON);
    cmp_and_return_(MINBS_DISABLED);
    cmp_and_return_(MINBS_HOT);
    cmp_and_return_(MINBS_NORMAL);
    cmp_and_return_(MINBS_PUSHED);
    cmp_and_return_(WP_MINCAPTION);
    cmp_and_return_(MNCS_ACTIVE);
    cmp_and_return_(MNCS_DISABLED);
    cmp_and_return_(MNCS_INACTIVE);
    cmp_and_return_(WP_RESTOREBUTTON);
    cmp_and_return_(RBS_DISABLED);
    cmp_and_return_(RBS_HOT);
    cmp_and_return_(RBS_NORMAL);
    cmp_and_return_(RBS_PUSHED);
    cmp_and_return_(WP_SMALLCAPTION);
    cmp_and_return_(CS_ACTIVE);
    cmp_and_return_(CS_DISABLED);
    cmp_and_return_(CS_INACTIVE);
    cmp_and_return_(WP_SMALLCAPTIONSIZINGTEMPLATE);
    cmp_and_return_(WP_SMALLCLOSEBUTTON);
    cmp_and_return_(CBS_DISABLED);
    cmp_and_return_(CBS_HOT);
    cmp_and_return_(CBS_NORMAL);
    cmp_and_return_(CBS_PUSHED);
    cmp_and_return_(WP_SMALLFRAMEBOTTOM);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_SMALLFRAMEBOTTOMSIZINGTEMPLATE);
    cmp_and_return_(WP_SMALLFRAMELEFT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_SMALLFRAMELEFTSIZINGTEMPLATE);
    cmp_and_return_(WP_SMALLFRAMERIGHT);
    cmp_and_return_(FS_ACTIVE);
    cmp_and_return_(FS_INACTIVE);
    cmp_and_return_(WP_SMALLFRAMERIGHTSIZINGTEMPLATE);
    //cmp_and_return_(WP_SMALLHELPBUTTON);
    cmp_and_return_(HBS_DISABLED);
    cmp_and_return_(HBS_HOT);
    cmp_and_return_(HBS_NORMAL);
    cmp_and_return_(HBS_PUSHED);
    //cmp_and_return_(WP_SMALLMAXBUTTON);
    cmp_and_return_(MAXBS_DISABLED);
    cmp_and_return_(MAXBS_HOT);
    cmp_and_return_(MAXBS_NORMAL);
    cmp_and_return_(MAXBS_PUSHED);
    cmp_and_return_(WP_SMALLMAXCAPTION);
    cmp_and_return_(MXCS_ACTIVE);
    cmp_and_return_(MXCS_DISABLED);
    cmp_and_return_(MXCS_INACTIVE);
    cmp_and_return_(WP_SMALLMINCAPTION);
    cmp_and_return_(MNCS_ACTIVE);
    cmp_and_return_(MNCS_DISABLED);
    cmp_and_return_(MNCS_INACTIVE);
    //cmp_and_return_(WP_SMALLRESTOREBUTTON);
    cmp_and_return_(RBS_DISABLED);
    cmp_and_return_(RBS_HOT);
    cmp_and_return_(RBS_NORMAL);
    cmp_and_return_(RBS_PUSHED);
    //cmp_and_return_(WP_SMALLSYSBUTTON);
    cmp_and_return_(SBS_DISABLED);
    cmp_and_return_(SBS_HOT);
    cmp_and_return_(SBS_NORMAL);
    cmp_and_return_(SBS_PUSHED);
    cmp_and_return_(WP_SYSBUTTON);
    cmp_and_return_(SBS_DISABLED);
    cmp_and_return_(SBS_HOT);
    cmp_and_return_(SBS_NORMAL);
    cmp_and_return_(SBS_PUSHED);
    cmp_and_return_(WP_VERTSCROLL);
    cmp_and_return_(VSS_DISABLED);
    cmp_and_return_(VSS_HOT);
    cmp_and_return_(VSS_NORMAL);
    cmp_and_return_(VSS_PUSHED);
    cmp_and_return_(WP_VERTTHUMB);
    cmp_and_return_(VTS_DISABLED);
    cmp_and_return_(VTS_HOT);
    cmp_and_return_(VTS_NORMAL);
    cmp_and_return_(VTS_PUSHED); 
    
    // GetThemeColor

    cmp_and_return_(TMT_BORDERCOLOR);
    cmp_and_return_(TMT_FILLCOLOR);
    cmp_and_return_(TMT_TEXTCOLOR);
    cmp_and_return_(TMT_EDGELIGHTCOLOR);
    cmp_and_return_(TMT_EDGEHIGHLIGHTCOLOR);
    cmp_and_return_(TMT_EDGESHADOWCOLOR);
    cmp_and_return_(TMT_EDGEDKSHADOWCOLOR);
    cmp_and_return_(TMT_EDGEFILLCOLOR);
    cmp_and_return_(TMT_TRANSPARENTCOLOR);
    cmp_and_return_(TMT_GRADIENTCOLOR1);
    cmp_and_return_(TMT_GRADIENTCOLOR2);
    cmp_and_return_(TMT_GRADIENTCOLOR3);
    cmp_and_return_(TMT_GRADIENTCOLOR4);
    cmp_and_return_(TMT_GRADIENTCOLOR5);
    cmp_and_return_(TMT_SHADOWCOLOR);
    cmp_and_return_(TMT_GLOWCOLOR);
    cmp_and_return_(TMT_TEXTBORDERCOLOR);
    cmp_and_return_(TMT_TEXTSHADOWCOLOR);
    cmp_and_return_(TMT_GLYPHTEXTCOLOR);
    cmp_and_return_(TMT_GLYPHTRANSPARENTCOLOR);
    cmp_and_return_(TMT_FILLCOLORHINT);
    cmp_and_return_(TMT_BORDERCOLORHINT);
    cmp_and_return_(TMT_ACCENTCOLORHINT);
    cmp_and_return_(TMT_BLENDCOLOR);

    // GetThemeFont
    cmp_and_return_(TMT_GLYPHTYPE);
    cmp_and_return_(TMT_FONT);

    Tcl_SetObjErrorCode(interp,
                        Twapi_MakeTwapiErrorCodeObj(TWAPI_INVALID_ARGS));
    Tcl_AppendResult(interp, "Invalid theme symbol '", name, "'", NULL);
    return TCL_ERROR;

 success_return:
    Tcl_SetObjResult(interp, Tcl_NewIntObj(val));
    return TCL_OK;

#undef cmp_and_return_
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
    Tcl_SetObjResult(interp, Tcl_NewListObj(3, objv));
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

    Tcl_SetObjResult(interp,
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

    Tcl_SetObjResult(interp, ObjFromLOGFONTW(&lf));
    return TCL_OK;
}

static int Twapi_MmCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s;
    DWORD dw, dw2;
    HMODULE hmod;
    TwapiResult result;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETHANDLET(hmod, HMODULE), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s);
        result.type = TRT_BOOL;
        result.value.ival = PlaySoundW(s, hmod, dw);
        break;
    case 2:
        result.type = TRT_BOOL;
        result.value.bval = MessageBeep(dw);
        break;
    case 3:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = Beep(dw, dw2);
        break;
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_UiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::UiCall", Twapi_UiCallObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::UiCallW", Twapi_UiCallWObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::UiCallWU", Twapi_UiCallWUObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::Ui" #call_, # code_); \
    } while (0);

    CALL_(GetDesktopWindow, Call, 5);
    CALL_(GetShellWindow, Call, 6);
    CALL_(GetForegroundWindow, Call, 7);
    CALL_(GetActiveWindow, Call, 8);
    CALL_(IsThemeActive, Call, 54);
    CALL_(IsAppThemed, Call, 55);
    CALL_(GetCurrentThemeName, Call, 56);
    CALL_(GetCursorPos, Call, 61);
    CALL_(GetCaretPos, Call, 62);
    CALL_(GetCaretBlinkTime, Call, 63);
    CALL_(EnumWindows, Call, 64);
    CALL_(GetFocus, Call, 81);                  /* TBD Tcl */

    CALL_(FlashWindowEx, Call, 1002);
    CALL_(WindowFromPoint, Call, 1009);
    CALL_(TwapiGetThemeDefine, Call, 1021);
    CALL_(MonitorFromPoint, Call, 10010);
    CALL_(MonitorFromRect, Call, 10011);
    CALL_(EnumDisplayDevices, Call, 10012);
    CALL_(GetThemeColor, Call, 10089);
    CALL_(GetThemeFont, Call, 10090);

    CALL_(SetCaretBlinkTime, CallU, 29);
    CALL_(GetGUIThreadInfo, CallU, 31);
    CALL_(SetCaretPos, CallU, 1004);
    CALL_(SetCursorPos, CallU, 1005);
    CALL_(CreateRectRgn, CallU, 3001);
    CALL_(CreateEllipticRgn, CallU, 3002);
    CALL_(CreateRoundedRectRgn, CallU, 3003);
    CALL_(CloseThemeData, CallH, 30);

    CALL_(EnumDesktopWindows, CallH, 35);
    CALL_(GetMonitorInfo, CallH, 58);

    CALL_(GetThemeSysColor, CallH, 1021); /* TBD - tcl wrapper */
    CALL_(GetThemeSysFont, CallH, 1022);  /* TBD - tcl wrapper */

    CALL_(EnumDisplayMonitors, CallH, 10005);

    // CallW - function(HWND)
    CALL_(IsIconic, CallW, 1);
    CALL_(IsZoomed, CallW, 2);
    CALL_(IsWindowVisible, CallW, 3);
    CALL_(IsWindow, CallW, 4);
    CALL_(IsWindowUnicode, CallW, 5);
    CALL_(IsWindowEnabled, CallW, 6);
    CALL_(ArrangeIconicWindows, CallW, 7);
    CALL_(SetForegroundWindow, CallW, 8);
    CALL_(OpenIcon, CallW, 9);
    CALL_(CloseWindow, CallW, 10);
    CALL_(DestroyWindow, CallW, 11);
    CALL_(UpdateWindow, CallW, 12);
    CALL_(HideCaret, CallW, 13);
    CALL_(ShowCaret, CallW, 14);
    CALL_(GetParent, CallW, 15);
    CALL_(GetClientRect, CallW, 17);
    CALL_(GetWindowRect, CallW, 18);
    CALL_(GetDC, CallW, 19);
    CALL_(SetFocus, CallW, 20);
    CALL_(SetActiveWindow, CallW, 21);
    CALL_(GetClassName, CallW, 22);
    CALL_(RealGetWindowClass, CallW, 23);
    CALL_(GetWindowThreadProcessId, CallW, 24);
    CALL_(GetWindowText, CallW, 25);
    CALL_(GetWindowDC, CallW, 26);
    CALL_(EnumChildWindows, CallW, 28);
    CALL_(GetWindowPlacement, CallW, 30); // TBD - Tcl wrapper
    CALL_(GetWindowInfo, CallW, 31); // TBD - Tcl wrapper

    CALL_(SetWindowText, CallW, 1001);
    CALL_(IsChild, CallW, 1002);
    CALL_(SetWindowPlacement, CallW, 1003); // TBD - Tcl wrapper
    CALL_(InvalidateRect, CallW, 1004);     // TBD - Tcl wrapper
    CALL_(SetWindowPos, CallW, 1005);
    CALL_(FindWindowEx, CallW, 1006);
    CALL_(ReleaseDC, CallW, 1007);
    CALL_(OpenThemeData, CallW, 1008);
    CALL_(SetWindowRgn, CallW, 1009); // TBD - Tcl wrapper
    CALL_(GetWindowRgn, CallW, 1010); // TBD - Tcl wrapper

    // CallWU - function(HWND, DWORD)
    CALL_(GetAncestor, CallWU, 1);
    CALL_(GetWindow, CallWU, 2);
    CALL_(ShowWindow, CallWU, 3);
    CALL_(ShowWindowAsync, CallWU, 4);
    CALL_(EnableWindow, CallWU, 5);
    CALL_(ShowOwnedPopups, CallWU, 6);
    CALL_(MonitorFromWindow, CallWU, 7);

    CALL_(SetLayeredWindowAttributes, CallWU, 10001);
    CALL_(MoveWindow, CallWU, 10002);

    CALL_(FindWindow, CallSSSD, 38);

#undef CALL_

    return TCL_OK;
}

#ifndef TWAPI_STATIC_BUILD
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_STATIC_BUILD
__declspec(dllexport) 
#endif
int Twapi_ui_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_UiInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

