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

static int Twapi_UiCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s, s2;
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

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 1000) {
        /* Functions taking no arguments */
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 1:
            result.type = GetCursorPos(&result.value.point) ? TRT_POINT : TRT_GETLASTERROR;
            break;
        case 2:
            result.type = GetCaretPos(&result.value.point) ? TRT_POINT : TRT_GETLASTERROR;
            break;
        case 3:
            result.type = TRT_DWORD;
            result.value.ival = GetCaretBlinkTime();
            break;
        case 4:
            return Twapi_EnumWindows(interp);
        case 5:
            result.type = TRT_HWND;
            result.value.hwin = GetDesktopWindow();
            break;
        case 6:
            result.type = TRT_HWND;
            result.value.hwin = GetShellWindow();
            break;
        case 7:
            result.type = TRT_HWND;
            result.value.hwin = GetForegroundWindow();
            break;
        case 8:
            result.type = TRT_HWND;
            result.value.hwin = GetActiveWindow();
            break;
        case 9:
            result.type = TRT_BOOL;
            result.value.bval = IsThemeActive();
            break;
        case 10:
            result.type = TRT_BOOL;
            result.value.bval = IsAppThemed();
            break;
        case 11:
            return Twapi_GetCurrentThemeName(interp);
        case 12:
            result.type = TRT_HWND;
            result.value.hwin = GetFocus();
            break;
        }
    } else if (func < 2000) {
        /* Exactly least one arg */
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1001:
            if (ObjToPOINT(interp, objv[2], &u.pt) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HWND;
            result.value.hwin = WindowFromPoint(u.pt);
            break;
        case 1002:
            if (ObjToFLASHWINFO(interp, objv[2], &u.flashw) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = FlashWindowEx(&u.flashw);
            break;
        case 1003:
            return TwapiGetThemeDefine(interp, Tcl_GetString(objv[2]));
        case 1004:              /* DeleteObject */
            if (ObjToOpaque(interp, objv[2], &u.hgdiobj, "HGDIOBJ") != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeleteObject(u.hgdiobj);
            break;
        }
    } else if (func < 4000) {
        /*
         * At least one integer arg. Additional args ignored for some
         * so arg errors are not quite accurate but that's ok, the
         * Tcl wrappers catch them.
         */
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), ARGUSEDEFAULT,
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
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(h),
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
            result.value.ival = GetThemeSysColor(h, dw);
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
        }
    } else {
        /* Arbitrary args */
        switch (func) {
        case 10001:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(u.pt, ObjToPOINT), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromPoint(u.pt, dw);
            break;
        case 10002:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(u.rect, ObjToRECT), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromRect(&u.rect, dw);
            break;
        case 10003: 
            return Twapi_GetThemeColor(interp, objc-2, objv+2);
        case 10004: 
            return Twapi_GetThemeFont(interp, objc-2, objv+2);
        case 10005:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            u.display_device.cb = sizeof(u.display_device);
            if (EnumDisplayDevicesW(s, dw, &u.display_device, dw2)) {
                result.value.obj = ObjFromDISPLAY_DEVICE(&u.display_device);
                result.type = TRT_OBJ;
            } else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = ERROR_INVALID_PARAMETER;
            }
            break;
        case 10006:
            if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(h),
                             ARGSKIP, ARGEND) != TCL_OK)
                return TCL_ERROR;
            rectP = &u.rect;
            if (ObjToRECT_NULL(interp, objv[3], &rectP) != TCL_OK)
                return TCL_ERROR;
            return Twapi_EnumDisplayMonitors(interp, h, rectP);
        case 10007:
            if (TwapiGetArgs(interp, objc-2, objv+2, 
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HWND;
            result.value.hwin = FindWindowW(s,s2);
            break;
        }
    }
    return TwapiSetResult(interp, &result);
}

int Twapi_UiCallWObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HWND hwnd, hwnd2;
    TwapiResult result;
    DWORD dw, dw2, dw3, dw4, dw5;
    Tcl_Obj *objs[2];
    int func;
    union {
        WINDOWPLACEMENT winplace;
        WINDOWINFO wininfo;
        WCHAR buf[MAX_PATH+1];
        RECT   rect;
        HRGN   hrgn;
    } u;
    RECT *rectP;
    LPWSTR s, s2;
    HANDLE h;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETHWND(hwnd),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 1000) {
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1:
            result.type = TRT_BOOL;
            result.value.bval = IsIconic(hwnd);
            break;
        case 2:
            result.type = TRT_BOOL;
            result.value.bval = IsZoomed(hwnd);
            break;
        case 3:
            result.type = TRT_BOOL;
            result.value.bval = IsWindowVisible(hwnd);
            break;
        case 4:
            result.type = TRT_BOOL;
            result.value.bval = IsWindow(hwnd);
            break;
        case 5:
            result.type = TRT_BOOL;
            result.value.bval = IsWindowUnicode(hwnd);
            break;
        case 6:
            result.type = TRT_BOOL;
            result.value.bval = IsWindowEnabled(hwnd);
            break;
        case 7:
            result.type = TRT_BOOL;
            result.value.bval = ArrangeIconicWindows(hwnd);
            break;
        case 8:
            result.type = TRT_BOOL;
            result.value.bval = SetForegroundWindow(hwnd);
            break;
        case 9:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = OpenIcon(hwnd);
            break;
        case 10:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseWindow(hwnd);
            break;
        case 11:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DestroyWindow(hwnd);
            break;
        case 12:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = UpdateWindow(hwnd);
            break;
        case 13:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = HideCaret(hwnd);
            break;
        case 14:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ShowCaret(hwnd);
            break;
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
            dw2 = GetWindowThreadProcessId(hwnd, &dw);
            if (dw2 == 0) {
                result.type = TRT_GETLASTERROR;
            } else {
                objs[0] = Tcl_NewLongObj(dw2);
                objs[1] = Tcl_NewLongObj(dw);
                result.value.objv.nobj = 2;
                result.value.objv.objPP = objs;
                result.type = TRT_OBJV;
            }
            break;
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
        case 27:
            if (GetWindowPlacement(hwnd, &u.winplace)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromWINDOWPLACEMENT(&u.winplace);
                break;
            } else {
                result.type = TRT_GETLASTERROR;
            }
            break;

        case 28:
            return Twapi_EnumChildWindows(interp, hwnd);
        }
    } else if (func < 500) {
        /* Exactly one additional int arg */
        if (TwapiGetArgs(interp, objc-3, objv+3, GETINT(dw), ARGEND) != TCL_OK)
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
        if (objc < 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1001: // SetWindowText
            result.value.ival = SetWindowTextW(hwnd, Tcl_GetUnicode(objv[3]));
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 1002: // IsChild
            if (ObjToHWND(interp, objv[3], &hwnd2) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = IsChild(hwnd, hwnd2);
            break;
        case 1003: //SetWindowPlacement
            if (ObjToWINDOWPLACEMENT(interp, objv[3], &u.winplace) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetWindowPlacement(hwnd, &u.winplace);
            break;
        case 1004: // InvalidateRect
            rectP = &u.rect;
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETVAR(rectP, ObjToRECT_NULL), GETINT(dw), 
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = InvalidateRect(hwnd, rectP, dw);
            break;
        case 1005: // SetWindowPos
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETHANDLET(hwnd2, HWND), GETINT(dw), GETINT(dw2),
                             GETINT(dw3), GETINT(dw4), GETINT(dw5),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetWindowPos(hwnd, hwnd2, dw, dw2, dw3, dw4, dw5);
            break;
        case 1006: // FindWindowEx
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETHANDLET(hwnd2, HWND),
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HWND;
            result.value.hval = FindWindowExW(hwnd, hwnd2, s, s2);
            break;
        case 1007:
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETHANDLET(h, HDC),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_DWORD;
            result.value.ival = ReleaseDC(hwnd, h);
            break;
        case 1008:
            result.type = TRT_HANDLE;
            result.value.hval = OpenThemeData(hwnd, Tcl_GetUnicode(objv[3]));
            break;
        case 1009: // SetWindowRgn
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETHANDLET(u.hrgn, HRGN), GETBOOL(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetWindowRgn(hwnd, u.hrgn, dw);
            break;
        case 1010: // GetWindowRgn
            if (TwapiGetArgs(interp, objc-3, objv+3, 
                             GETHANDLET(u.hrgn, HRGN),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_DWORD;
            result.value.ival = GetWindowRgn(hwnd, u.hrgn);
            break;
        case 1011:
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETINT(dw), GETINT(dw2), GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetLayeredWindowAttributes(hwnd, dw, (BYTE)dw2, dw3);
            break;
        case 1012:
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETINT(dw), GETINT(dw2), GETINT(dw3),
                             GETINT(dw4), GETINT(dw5),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = MoveWindow(hwnd, dw, dw2, dw3, dw4, dw5);
            break;
        }
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_UiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::UiCall", Twapi_UiCallObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::UiCallW", Twapi_UiCallWObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::Ui" #call_, # code_); \
    } while (0);

    CALL_(GetCursorPos, Call, 1);
    CALL_(GetCaretPos, Call, 2);
    CALL_(GetCaretBlinkTime, Call, 3);
    CALL_(EnumWindows, Call, 4);
    CALL_(GetDesktopWindow, Call, 5);
    CALL_(GetShellWindow, Call, 6);
    CALL_(GetForegroundWindow, Call, 7);
    CALL_(GetActiveWindow, Call, 8);
    CALL_(IsThemeActive, Call, 9);
    CALL_(IsAppThemed, Call, 10);
    CALL_(GetCurrentThemeName, Call, 11);
    CALL_(GetFocus, Call, 12);                  /* TBD Tcl */

    CALL_(WindowFromPoint, Call, 1001);
    CALL_(FlashWindowEx, Call, 1002);
    CALL_(TwapiGetThemeDefine, Call, 1003);
    CALL_(DeleteObject, Call, 1004);
    CALL_(SetCaretBlinkTime, Call, 3001);
    CALL_(GetGUIThreadInfo, Call, 3002);
    CALL_(SetCaretPos, Call, 3003);
    CALL_(SetCursorPos, Call, 3004);
    CALL_(CreateRectRgn, Call, 3005);
    CALL_(CreateEllipticRgn, Call, 3006);
    CALL_(CreateRoundedRectRgn, Call, 3007);
    CALL_(CloseThemeData, Call, 4001);
    CALL_(EnumDesktopWindows, Call, 4002);
    CALL_(GetMonitorInfo, Call, 4003);
    CALL_(GetThemeSysColor, Call, 4004); /* TBD - tcl wrapper */
    CALL_(GetThemeSysFont, Call, 4005);  /* TBD - tcl wrapper */
    CALL_(MonitorFromPoint, Call, 10001);
    CALL_(MonitorFromRect, Call, 10002);
    CALL_(GetThemeColor, Call, 10003);
    CALL_(GetThemeFont, Call, 10004);
    CALL_(EnumDisplayDevices, Call, 10005);
    CALL_(EnumDisplayMonitors, Call, 10006);
    CALL_(FindWindow, Call, 10007);

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
    CALL_(GetWindowInfo, CallW, 16); // TBD - Tcl wrapper
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
    CALL_(GetWindowPlacement, CallW, 27); // TBD - Tcl wrapper
    CALL_(EnumChildWindows, CallW, 28);
    CALL_(GetAncestor, CallW, 501);
    CALL_(GetWindow, CallW, 502);
    CALL_(ShowWindow, CallW, 503);
    CALL_(ShowWindowAsync, CallW, 504);
    CALL_(EnableWindow, CallW, 505);
    CALL_(ShowOwnedPopups, CallW, 506);
    CALL_(MonitorFromWindow, CallW, 507);
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
    CALL_(SetLayeredWindowAttributes, CallW, 1011);
    CALL_(MoveWindow, CallW, 1012);


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

