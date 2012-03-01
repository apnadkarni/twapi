/*
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_wm.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

int Twapi_EnumClipboardFormats(Tcl_Interp *interp);
int Twapi_ClipboardMonitorStart(TwapiInterpContext *ticP);
int Twapi_ClipboardMonitorStop(TwapiInterpContext *ticP);
static DWORD TwapiClipboardCallbackFn(TwapiCallback *pcbP);

int Twapi_EnumClipboardFormats(Tcl_Interp *interp)
{
    UINT clip_fmt;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    Tcl_SetObjResult(interp, resultObj);

    clip_fmt = 0;
    while (1) {
        clip_fmt = EnumClipboardFormats(clip_fmt);
        if (clip_fmt)
            Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewIntObj(clip_fmt));
        else {
            DWORD error = GetLastError();
            if (error != ERROR_SUCCESS) {
                return TwapiReturnSystemError(interp);
            }
            break;
        }
    }
    return TCL_OK;
}

/* Structure to hold callback data */
struct TwapiClipboardMonitorState {
    HWND next_win;              // Next window in clipboard chain
};

static LRESULT TwapiClipboardMonitorWinProc(
    TwapiInterpContext *ticP,
    LONG_PTR clientdataP,
    HWND hwnd,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
    struct TwapiClipboardMonitorState *clip_stateP = (struct TwapiClipboardMonitorState *) clientdataP;
    TwapiCallback *cbP;

    switch (uMsg)
    {
        case TWAPI_WM_HIDDEN_WINDOW_INIT:
            // Add the window to the clipboard viewer chain.
            clip_stateP->next_win = SetClipboardViewer(hwnd);
            break;

        case WM_CHANGECBCHAIN:

            // If the next window is closing, repair the chain.
            // Otherwise, pass the message to the next link.
            if ((HWND) wParam == clip_stateP->next_win)
                clip_stateP->next_win = (HWND) lParam;
            else if (clip_stateP->next_win != NULL)
                SendMessage(clip_stateP->next_win, uMsg, wParam, lParam);

            break;

        case WM_DESTROY:
            ChangeClipboardChain(hwnd, clip_stateP->next_win);
            TwapiFree(clip_stateP);
            SetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET, (LONG_PTR) 0);
            break;

        case WM_DRAWCLIPBOARD:  // clipboard contents changed.
            cbP = TwapiCallbackNew(ticP, TwapiClipboardCallbackFn,
                                   sizeof(*cbP));
            TwapiEnqueueCallback(ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);

            /* Pass the message to the next window in clipboard chain */
            if (clip_stateP->next_win != NULL)
                SendMessage(clip_stateP->next_win, uMsg, wParam, lParam);
            break;

        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return (LRESULT) NULL;
}


TCL_RESULT Twapi_ClipboardMonitorStart(TwapiInterpContext *ticP)
{
    struct TwapiClipboardMonitorState *clip_stateP;

    if (ticP->module.data.hwnd) {
        Tcl_SetResult(ticP->interp, "Clipboard monitoring is already in progress", TCL_STATIC);
        return TCL_ERROR;
    }

    clip_stateP = (struct TwapiClipboardMonitorState *)
        TwapiAlloc(sizeof *clip_stateP);
    clip_stateP->next_win = NULL;

    /*
     * Create a hidden Twapi window for clipboard monitoring.
     * All the hardwork is done inside the window procedure.
     */
    if (Twapi_CreateHiddenWindow(ticP, TwapiClipboardMonitorWinProc, (LONG_PTR) clip_stateP, &ticP->module.data.hwnd) != TCL_OK) {
        TwapiFree(clip_stateP);
        return TCL_ERROR;
    }

    return TCL_OK;
}

TCL_RESULT Twapi_ClipboardMonitorStop(TwapiInterpContext *ticP)
{
    if (ticP->module.data.hwnd) {
        DestroyWindow(ticP->module.data.hwnd);
        ticP->module.data.hwnd = NULL;
    }

    // Note all data will be reset in the
    // window message handler after unhooking
    // from clipboard chain

    return TCL_OK;
}

/* Called (indirectly) from the Tcl notifier loop with a new clipboard event.
 * Follows behaviour specified by TwapiCallbackFn typedef.
 */
static DWORD TwapiClipboardCallbackFn(TwapiCallback *cbP)
{
    Tcl_Obj *objP = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_clipboard_handler", -1);
    return TwapiEvalAndUpdateCallback(cbP, 1, &objP, TRT_EMPTY);
}

static TCL_RESULT Twapi_ClipboardCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    HWND   hwnd;
    DWORD dw;
    HANDLE h;
    TwapiResult result;
    WCHAR buf[MAX_PATH+1];

    /* Every command has at least one argument */
    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_EnumClipboardFormats(interp);
    case 2:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CloseClipboard();
        break;
    case 3:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EmptyClipboard();
        break;
    case 4:
        result.type = TRT_HWND;
        result.value.hwin = GetOpenClipboardWindow();
        break;
    case 5:
        result.type = TRT_HWND;
        result.value.hwin = GetClipboardOwner();
        break;
    case 6:
        return Twapi_ClipboardMonitorStart(ticP);
    case 7:
        return Twapi_ClipboardMonitorStop(ticP);
    case 8:
    case 9:
    case 10:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        switch (func) {
        case 8:
            result.type = TRT_BOOL;
            result.value.bval = IsClipboardFormatAvailable(dw);
            break;
        case 9:
            result.type = TRT_HANDLE;
            result.value.hval = GetClipboardData(dw);
            break;
        case 10:
            result.value.unicode.len = GetClipboardFormatNameW(dw, buf, sizeof(buf)/sizeof(buf[0]));
            result.value.unicode.str = buf;
            result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
            break;
        }
        break;
    case 11:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), GETHANDLE(h), ARGEND) != TCL_ERROR)
            return TCL_ERROR;
        result.type = TRT_HANDLE;
        result.value.hval = SetClipboardData(dw, h);
        break;
    case 12:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        result.type = TRT_NONZERO_RESULT;
        result.value.ival = RegisterClipboardFormatW(Tcl_GetUnicode(objv[2]));
        break;
    case 16:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHWND(hwnd), ARGEND) != TCL_ERROR)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = OpenClipboard(hwnd);
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_ClipboardInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::ClipCall", Twapi_ClipboardCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::ClipCall", # code_); \
    } while (0);

    CALL_(Twapi_EnumClipboardFormats, 1);
    CALL_(CloseClipboard, 2);
    CALL_(EmptyClipboard, 3);
    CALL_(GetOpenClipboardWindow, 4);
    CALL_(GetClipboardOwner, 5);
    CALL_(Twapi_ClipboardMonitorStart, 6);
    CALL_(Twapi_ClipboardMonitorStop, 7);
    CALL_(IsClipboardFormatAvailable, 8);
    CALL_(GetClipboardData, 9);
    CALL_(GetClipboardFormatName, 10);
    CALL_(SetClipboardData, 11);
    CALL_(RegisterClipboardFormat, 12);
    CALL_(OpenClipboard, 13);


#undef CALL_

    return TCL_OK;
}

void TwapiClipboardCleanup(TwapiInterpContext *ticP)
{
    if (ticP->module.data.hwnd) {
        DestroyWindow(ticP->module.data.hwnd);
        ticP->module.data.hwnd = NULL;
    }
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
int Twapi_clipboard_Init(Tcl_Interp *interp)
{
    TwapiInterpContext *ticP;

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    ticP = Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_ClipboardInitCalls, TwapiClipboardCleanup);
    if (ticP) {
        ticP->module.data.hwnd = NULL;
        return TCL_OK;
    } else
        return TCL_ERROR;
}

