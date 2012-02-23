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
        TwapiZeroMemory(&w, sizeof(w));
        w.cbSize = sizeof(w);
        w.hInstance = gTwapiModuleHandle;
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
    Tcl_Obj *objs[6];

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_script_wm_handler", -1);
    objs[1] = ObjFromTwapiId(cbP->receiver_id);
    objs[2] = ObjFromDWORD_PTR(cbP->clientdata);
    objs[3] = ObjFromDWORD_PTR(cbP->clientdata2);
    objs[4] = ObjFromPOINTS(&cbP->wm_state.message_pos);
    objs[5] = ObjFromDWORD(cbP->wm_state.ticks);
    return TwapiEvalAndUpdateCallback(cbP, 6, objs, TRT_EMPTY);
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
    static UINT wm_taskbar_restart;

    if (msg == WM_HOTKEY ||
        (msg >= TWAPI_WM_SCRIPT_BASE && msg <= TWAPI_WM_SCRIPT_LAST)) {
        TwapiCallback *cbP;
        DWORD pos;

        cbP = TwapiCallbackNew(ticP, TwapiScriptWMCallbackFn, sizeof(*cbP));
        cbP->receiver_id = msg;
        cbP->clientdata = wParam;
        cbP->clientdata2 = lParam;
        pos = GetMessagePos();
        cbP->wm_state.message_pos = MAKEPOINTS(pos);
        cbP->wm_state.ticks = GetTickCount();
        TwapiEnqueueCallback(ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);
        return (LRESULT) NULL;
    } else {
        switch (msg) {
        case WM_CREATE:
            /*
             * On creation, stash the message number windows sends when
             * the taskbar is (re) created
             * Not thread safe, but no matter
             */
            if (wm_taskbar_restart == 0)
                wm_taskbar_restart = RegisterWindowMessageW(L"TaskbarCreated");
            break;
        case WM_POWERBROADCAST:
            if (ticP->power_events_on)
                return TwapiPowerHandler(ticP, msg, wParam, lParam);
            break;
        default:
            if (msg == WM_HOTKEY ||
                (msg >= TWAPI_WM_SCRIPT_BASE && msg <= TWAPI_WM_SCRIPT_LAST) ||
                msg == wm_taskbar_restart) {
#ifdef  DO_NOT_USE_WINMSG_DIRECT_CALL
                TwapiCallback *cbP;
                DWORD pos;
                if (msg == wm_taskbar_restart)
                    msg = TWAPI_WM_TASKBAR_RESTART; /* Map dynamic number to our fixed number */

                cbP = TwapiCallbackNew(ticP, TwapiScriptWMCallbackFn, sizeof(*cbP));
                cbP->receiver_id = msg;
                cbP->clientdata = wParam;
                cbP->clientdata2 = lParam;
                pos = GetMessagePos();
                cbP->wm_state.message_pos = MAKEPOINTS(pos);
                cbP->wm_state.ticks = GetTickCount();
                TwapiEnqueueCallback(ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);
                return (LRESULT) NULL;
#else
                return TwapiEvalWinMessage(ticP, msg, wParam, lParam);
#endif
            }
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
HWND Twapi_GetNotificationWindow(TwapiInterpContext *ticP)
{
    int status;

    if (ticP->notification_win)
        return ticP->notification_win;

    status =  Twapi_CreateHiddenWindow(
        ticP, TwapiNotificationWindowHandler, 0, &ticP->notification_win);
    return status == TCL_OK ? ticP->notification_win : NULL;
}


LRESULT TwapiEvalWinMessage(TwapiInterpContext *ticP, UINT msg, WPARAM wParam, LPARAM lParam)
{
    int i;
    Tcl_Obj *objs[6];
    DWORD pos;
    POINTS pts;
    Tcl_Interp *interp;
    Tcl_InterpState interp_state;
    LRESULT lresult = 0;

    interp = ticP->interp;
    if (interp == NULL ||
        Tcl_InterpDeleted(interp)) {
        return 0;
    }

    pos = GetMessagePos();
    pts = MAKEPOINTS(pos);
    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_script_wm_handler", -1);
    objs[1] = ObjFromDWORD(msg);
    objs[2] = ObjFromDWORD_PTR(wParam);
    objs[3] = ObjFromDWORD_PTR(lParam);
    objs[4] = ObjFromPOINTS(&pts);
    objs[5] = ObjFromDWORD(GetTickCount());

    for (i=0; i < ARRAYSIZE(objs); ++i) {
        Tcl_IncrRefCount(objs[i]);
    }

    /* Preserve structures during eval */
    TwapiInterpContextRef(ticP, 1);
    Tcl_Preserve(interp);

    /* Save the interp state.
       TBD - really required ? At this point interp should be dormant, no ?
    */
    interp_state = Tcl_SaveInterpState(interp, TCL_OK);

    lresult = 0;
    if (Tcl_EvalObjv(interp, ARRAYSIZE(objs), objs, 
                     TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL) == TCL_OK) {
        /* Note if not integer result, lresult stays 0 */
        ObjToDWORD_PTR(interp, Tcl_GetObjResult(interp), &lresult);
    }

    /* Restore Interp state */
    Tcl_RestoreInterpState(interp, interp_state);

    Tcl_Release(interp);
    TwapiInterpContextUnref(ticP, 1);

    /* 
     * Undo the incrref above. This will delete the object unless
     * caller had done an incr-ref on it.
     */
    for (i=0; i < ARRAYSIZE(objs); ++i) {
        Tcl_IncrRefCount(objs[i]);
    }

    return lresult;
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
