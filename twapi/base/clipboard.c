/*
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_wm.h"

static DWORD TwapiClipboardCallbackFn(TwapiPendingCallback *pcbP);


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
    TwapiPendingCallback *pcbP;

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
            pcbP = TwapiPendingCallbackNew(ticP, TwapiClipboardCallbackFn,
                                           sizeof(*pcbP));
            TwapiEnqueueCallback(ticP, pcbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);

            /* Pass the message to the next window in clipboard chain */
            if (clip_stateP->next_win != NULL)
                SendMessage(clip_stateP->next_win, uMsg, wParam, lParam);
            break;

        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return (LRESULT) NULL;
}


int Twapi_MonitorClipboardStart(TwapiInterpContext *ticP)
{
    struct TwapiClipboardMonitorState *clip_stateP;

    clip_stateP = (struct TwapiClipboardMonitorState *)
        TwapiAlloc(sizeof *clip_stateP);
    clip_stateP->next_win = NULL;

    /*
     * Create a hidden Twapi window for clipboard monitoring.
     * All the hardwork is done inside the window procedure.
     */
    if (Twapi_CreateHiddenWindow(ticP, TwapiClipboardMonitorWinProc, (LONG_PTR) clip_stateP, &ticP->clipboard_win) != TCL_OK) {
        TwapiFree(clip_stateP);
        return TCL_ERROR;
    }

    return TCL_OK;
}

int Twapi_MonitorClipboardStop(TwapiInterpContext *ticP)
{
    if (ticP->clipboard_win) {
        DestroyWindow(ticP->clipboard_win);
        ticP->clipboard_win = NULL;
    }

    // Note all data will be reset in the
    // window message handler after unhooking
    // from clipboard chain

    return TCL_OK;
}

/* Called (indirectly) from the Tcl notifier loop with a new clipboard event.
 * Follows behaviour specified by TwapiCallbackFn typedef.
 */
static DWORD TwapiClipboardCallbackFn(TwapiPendingCallback *pcbP)
{
    Tcl_Obj *objP = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_clipboard_handler", -1);
    return TwapiEvalAndUpdateCallback(pcbP, 1, &objP, TRT_EMPTY);
}
