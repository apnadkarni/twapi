/* 
 * Copyright (c) 2003-2009, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to UI functions */
/* TBD - IsHungAppWindow */

#include "twapi.h"

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


int Twapi_BlockInput(Tcl_Interp *interp, BOOL block)
{
    BOOL result = BlockInput(block);
    if (result || (GetLastError() == 0)) {
        Tcl_SetObjResult(interp, Tcl_NewIntObj(result));
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
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

/*
 * Need this to be able to distinguish between return value of 0
 * and a failure case.
 */
DWORD Twapi_SetWindowLongPtr(HWND hWnd, int nIndex, LONG_PTR lValue, LONG_PTR *retP)
{
    SetLastError(0);            /* Reset error */
    *retP = SetWindowLongPtrW(hWnd, nIndex, lValue);
    if (*retP == 0) {
        /* Possible error */
        DWORD error = GetLastError();
        if (error) {
            return FALSE;
        }
    }
    return TRUE;
}

static void init_keyboard_input(INPUT *pin, WORD vkey, DWORD flags);

int Twapi_SendInput(TwapiInterpContext *ticP, Tcl_Obj *input_obj) {
    int num_inputs;
    struct tagINPUT   *input;
    int i, j;
    int result = TCL_ERROR;
    Tcl_Interp *interp = ticP->interp;
    
    if (Tcl_ListObjLength(interp, input_obj, &num_inputs) != TCL_OK) {
        return TCL_ERROR;
    }

    input = MemLifoPushFrame(&ticP->memlifo, num_inputs * sizeof(*input), NULL);
    /* Loop through each element, parsing it and storing its descriptor */
    for (i = 0; i < num_inputs; ++i) {
        Tcl_Obj *event_obj;
        Tcl_Obj *field_obj[5];
        LONG     value[5];
        char *options[] = {"key", "mouse", NULL};
        int   option;


        if (Tcl_ListObjIndex(interp, input_obj, i, &event_obj) != TCL_OK)
            goto done;

        if (event_obj == NULL)
            break;

        /* This element is itself a list, parse it to get input type etc. */
        if (Tcl_ListObjIndex(interp, event_obj, 0, &field_obj[0]) != TCL_OK)
            goto  done;
        
        if (field_obj[0] == NULL)
            break;

        /* Figure out the input type and parse remaining fields */
        if (Tcl_GetIndexFromObj(interp, field_obj[0], options,
                                "input event type", TCL_EXACT, &option) != TCL_OK)
            goto done;

        switch (option) {
        case 0:
            /* A single key stroke. Fields are:
             *  virtualkey(1-254), scancode (0-65535), flags
             * Extra arguments ignored
             */
            for (j = 1; j < 4; ++j) {
                if (Tcl_ListObjIndex(interp, event_obj, j, &field_obj[j]) != TCL_OK)
                    goto done;
                if (field_obj[j] == NULL) {
                    Tcl_SetResult(interp, "Missing field in event of type key", TCL_STATIC);
                    goto done;
                }
                if (Tcl_GetLongFromObj(interp, field_obj[j], &value[j]) != TCL_OK)
                    goto done;
            }

            /* OK, our three fields have been parsed.
             * Validate and add to input
             */
            if (value[1] < 0 || value[1] > 254) {
                Tcl_SetResult(interp, "Invalid value specified for virtual key code. Must be between 1 and 254", TCL_STATIC);
                goto done;
            }
            if (value[2] < 0 || value[2] > 65535) {
                Tcl_SetResult(interp, "Invalid value specified for scan code. Must be between 1 and 65535", TCL_STATIC);
                goto done;
            }
            init_keyboard_input(&input[i], (WORD) value[1], value[3]);
            input[i].ki.wScan   = (WORD) value[2];
            break;

        case 1:
            /* Mouse event
             *  xpos ypos mousedata flags
             * Extra arguments ignored
             */
            for (j = 1; j < 5; ++j) {
                if (Tcl_ListObjIndex(interp, event_obj, j, &field_obj[j]) != TCL_OK)
                    goto done;
                if (field_obj[j] == NULL) {
                    Tcl_SetResult(interp, "Missing field in event of type mouse", TCL_STATIC);
                    goto done;
                }
                if (Tcl_GetLongFromObj(interp, field_obj[j], &value[j]) != TCL_OK)
                    goto done;
            }
            
            input[i].type           = INPUT_MOUSE;
            input[i].mi.dx          = value[1];
            input[i].mi.dy          = value[2];
            input[i].mi.mouseData   = value[3];
            input[i].mi.dwFlags     = value[4];
            input[i].mi.time        = 0;
            input[i].mi.dwExtraInfo = 0;
            break;

         default:
            /* Shouldn't happen else Tcl_GetIndexFromObj would return error */
            Tcl_SetResult(interp, "Unknown field event type", TCL_STATIC);
            goto done;
        }

    }

    
    /* i is actual number of elements found */
    if (i != num_inputs) {
        Tcl_SetResult(interp, "Invalid or empty element specified in input event list", TCL_STATIC);
        goto done;
    }

    /* OK, we have everything in the input[] array. Send it along */
    if (i) {
        num_inputs = SendInput(i, input, sizeof(input[0]));
        if (num_inputs == 0) {
            j = GetLastError();
            Tcl_SetResult(interp, "Error sending input events: ", TCL_STATIC);
            Twapi_AppendSystemError(interp, j);
            goto done;
        }
    }    

    Tcl_SetObjResult(interp, Tcl_NewIntObj(num_inputs));
    result = TCL_OK;

 done:

    MemLifoPopFrame(&ticP->memlifo);

    return result;
}


int Twapi_SendUnicode(TwapiInterpContext *ticP, Tcl_Obj *input_obj) {
    int num_chars;
    struct tagINPUT   *input = NULL;
    int i, j;
    int result = TCL_ERROR;
    int max_input_records;
    int sent_inputs;
    
    num_chars = Tcl_GetCharLength(input_obj);

    /* Now loop through every character adding it to the input event array */
    /* Win2K and up, accepts unicode characters */

    /* NUmber of events is twice number of chars (keydown + keyup) */
    max_input_records = 2 * num_chars;
    input = MemLifoAlloc(&ticP->memlifo, max_input_records * sizeof(*input), NULL);
    for (i = 0, j = 0; i < num_chars; ++i) {
        WCHAR wch;
            
        wch = Tcl_GetUniChar(input_obj, i);
#ifndef KEYEVENTF_UNICODE
#define KEYEVENTF_UNICODE     0x0004
#endif
        init_keyboard_input(&input[j], 0, KEYEVENTF_UNICODE);
        input[j].ki.wScan = wch;
        ++j;
        init_keyboard_input(&input[j], 0, KEYEVENTF_UNICODE|KEYEVENTF_KEYUP);
        input[j].ki.wScan  = wch;
        ++j;
    }
    
    /* j is actual number of input events created */
    assert (j <= max_input_records);

    /* OK, we have everything in the input[] array. Send it along */
    if (j) {
        sent_inputs = SendInput(j, input, sizeof(input[0]));
        if (sent_inputs == 0) {
            i = GetLastError();
            Tcl_SetResult(ticP->interp, "Error sending input events: ", TCL_STATIC);
            Twapi_AppendSystemError(ticP->interp, i);
            goto done;
        }
        /* TBD - what if we send fewer than expected, should we retry ? */
    } else {
        sent_inputs = 0;
    }

    Tcl_SetObjResult(ticP->interp, Tcl_NewIntObj(sent_inputs));
    result = TCL_OK;

 done:
    MemLifoPopFrame(&ticP->memlifo);

    return result;
}

static void init_keyboard_input(INPUT *pin, WORD vkey, DWORD flags)
{
    pin->type       = INPUT_KEYBOARD;
    pin->ki.wVk     = vkey;
    pin->ki.wScan   = 0;
    pin->ki.dwFlags = flags;
    pin->ki.time    = 0;
    pin->ki.dwExtraInfo = 0;
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
