/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>
#include "twapi_wm.h"


/* Window proc for Twapi hidden windows */
LRESULT CALLBACK TwapiHiddenWindowProc(
    HWND hwnd,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam
)
{
    TwapiInterpContext *ticP;
    TwapiHiddenWindowCallbackProc *winproc;
    LRESULT lres;
    LONG_PTR clientdata;

    ticP = (TwapiInterpContext *) GetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CONTEXT_OFFSET);
    winproc = (TwapiHiddenWindowCallbackProc *) GetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET);
    clientdata = GetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET);
    if (winproc) {
        lres = (*winproc)(ticP, clientdata, hwnd, uMsg, wParam, lParam);
        if (uMsg != WM_DESTROY)
            return (LRESULT) lres;

        /* If this was a destroy message, release the interp context. Note
         * we do this after invoking the callback so context is valid
         * in the callback
         */

        if (ticP)
            TwapiInterpContextUnref(ticP, 1);

        /* Reset the window data, just in case */
        SetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CONTEXT_OFFSET, (LONG_PTR) 0);
        SetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET, (LONG_PTR) 0);
        SetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET, (LONG_PTR) 0);
        /* Fall thru to call DefWindowProc for WM_DESTROY */
    } 

    /* Call the default window proc. This will happen during window init
       and destroy
    */
    lres = DefWindowProc(hwnd, uMsg, wParam, lParam);


    return (LRESULT) lres;
}



/* Create a hidden with the specified window procedure. */
int Twapi_CreateHiddenWindow(
    TwapiInterpContext *ticP,         /* May be NULL */
    TwapiHiddenWindowCallbackProc *winProc,
    LONG_PTR clientdata,
    HWND *winP                  /* If not NULL, contains HWND on success */
    )
{
    static ATOM hidden_win_class;
    HWND win;

    /*
     * Register the class for receiving messages. Note we are using
     * Unicode versions of the call so all received windows text
     * message will be unicode
     */
    if (hidden_win_class == 0) {
        WNDCLASSEXW w;
        ZeroMemory(&w, sizeof(w));
        w.cbSize = sizeof(w);
        w.hInstance = 0;
        w.cbWndExtra = TWAPI_HIDDEN_WINDOW_DATA_SIZE;
        w.lpfnWndProc = TwapiHiddenWindowProc;
        w.lpszClassName = TWAPI_HIDDEN_WINDOW_CLASS_L;

        hidden_win_class = RegisterClassExW(&w);
        if (hidden_win_class == 0) {
            /* Note interp can be null */
            if (ticP && ticP->interp)
                Twapi_AppendSystemError(ticP->interp, GetLastError());
            return TCL_ERROR;
        }
    }

    /*
     * Class is registered, now create the window. Originally, we created
     * this as a message-only window (by specifying HWND_MESSAGE as the
     * parent window. However, message-only windows do not receive broadcast
     * messages, WM_DEVICECHANGE in particular. So we now just create it as
     * a toplevel window.
     */
    win = CreateWindowW(TWAPI_HIDDEN_WINDOW_CLASS_L, NULL,
                        0, 0, 0, 0, 0,
                        NULL, NULL, 0, NULL);

    if (win == NULL) {
        if (ticP && ticP->interp)
            return Twapi_AppendSystemError(ticP->interp, GetLastError());
    }

    /* Store pointers to the interpreter and the real window function. */
    if (ticP)
        TwapiInterpContextRef(ticP, 1); /* So ticP does not disappear */

    SetLastError(0);            /* Else no way to detect error */
    if ((SetWindowLongPtrW(win, TWAPI_HIDDEN_WINDOW_CONTEXT_OFFSET, (LONG_PTR) ticP) != 0 || GetLastError() == 0) &&
        (SetWindowLongPtrW(win, TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET, (LONG_PTR) winProc) != 0 || GetLastError() == 0) &&
        (SetWindowLongPtrW(win, TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET, clientdata) !=0 || GetLastError() == 0)) {

        if (winP)
            *winP = win;
        PostMessage(win, TWAPI_WM_HIDDEN_WINDOW_INIT, 0, 0);
        return TCL_OK;
    }

    /* One of the calls failed */
    if (ticP) {
        if (ticP->interp)
            Twapi_AppendSystemError(ticP->interp, GetLastError());
        TwapiInterpContextUnref(ticP, 1); /* Matches Tcl_Preserve above */
    }

    return TCL_ERROR;
}


/* Called (indirectly) from the Tcl notifier loop with a new 
   windows message callback that is to be handled by the generic
   script level windows message handler.
 * Follows behaviour specified by TwapiCallbackFn typedef.
 */
static int TwapiScriptWMCallbackFn(TwapiCallback *cbP)
{
    Tcl_Obj *objs[4];

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_script_wm_handler", -1);
    objs[1] = ObjFromTwapiId(cbP->receiver_id);
    objs[2] = ObjFromDWORD_PTR(cbP->clientdata);
    objs[3] = ObjFromDWORD_PTR(cbP->clientdata2);
    return TwapiEvalAndUpdateCallback(cbP, 4, objs, TRT_EMPTY);
}

static LRESULT TwapiNotificationWindowHandler(
    TwapiInterpContext *ticP,
    LONG_PTR clientdata,
    HWND hwnd,
    UINT msg,
    WPARAM wParam, 
    LPARAM lParam
    )
{
    if (msg == WM_HOTKEY ||
        (msg >= TWAPI_WM_SCRIPT_BASE && msg <= TWAPI_WM_SCRIPT_LAST)) {
        TwapiCallback *cbP;
        cbP = TwapiCallbackNew(ticP, TwapiScriptWMCallbackFn, sizeof(*cbP));
        cbP->receiver_id = msg;
        cbP->clientdata = wParam;
        cbP->clientdata2 = lParam;
        TwapiEnqueueCallback(ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);
        return (LRESULT) NULL;
    } else {
        if (msg == WM_POWERBROADCAST) {
            if (ticP->power_events_on)
                return TwapiPowerHandler(ticP, msg, wParam, lParam);
            /* else FALLTHRU */
        }
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
}
    

/*
 * To minimize number of windows, we have a "common" window that can be shared
 * by notifications. This returns that window, creating if necessary. Note
 * This window is limited in the sense that it is not possible to store
 * custom data in the window.
 */
HWND TwapiGetNotificationWindow(TwapiInterpContext *ticP)
{
    int status;

    if (ticP->notification_win)
        return ticP->notification_win;

    status =  Twapi_CreateHiddenWindow(
        ticP, TwapiNotificationWindowHandler, 0, &ticP->notification_win);
    return status == TCL_OK ? ticP->notification_win : NULL;
}
