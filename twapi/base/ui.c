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
    objv[27] = Tcl_NewUnicodeObj(lfP->lfFaceName, -1);

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
        return TwapiReturnTwapiError(interp, "Incorrect format of WINDOWPLACEMENT argument.", TWAPI_INVALID_ARGS);

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
                             Tcl_NewUnicodeObj(p_winsta, -1));
    return 1;
}

#ifndef TWAPI_LEAN
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
#endif

#ifndef TWAPI_LEAN
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
#endif


#ifndef TWAPI_LEAN
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
#endif

#ifndef TWAPI_LEAN
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
#endif



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
