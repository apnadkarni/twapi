/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>
#include "twapi_wm.h"

/* Name of hidden window class */
#define TWAPI_HIDDEN_WINDOW_CLASS_L L"TwapiHiddenWindow"

/*
 * Define offsets in window data
 */
#define TWAPI_HIDDEN_WINDOW_INTERP_OFFSET     0
#define TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET   (TWAPI_HIDDEN_WINDOW_INTERP_OFFSET + sizeof(LONG_PTR))
#define TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET (TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET + sizeof(LONG_PTR))
#define TWAPI_HIDDEN_WINDOW_DATA_SIZE       (TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET + sizeof(LONG_PTR))


/* Window proc for Twapi hidden windows */
LRESULT CALLBACK TwapiHiddenWindowProc(
    HWND hwnd,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam
)
{
    Tcl_Interp *interp;
    TwapiHiddenWindowCallbackProc winproc;
    LRESULT lres;
    LONG_PTR clientdata;

    interp = (Tcl_Interp *) GetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_INTERP_OFFSET);
    winproc = (TwapiHiddenWindowCallbackProc) GetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET);
    clientdata = GetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET);
    if (winproc) {
        lres = (*winproc)(interp, clientdata, hwnd, uMsg, wParam, lParam);
        /* If this was a destroy message, release the interpreter. Note
           we do this after invoking the callback so interp is valid
           in the callback
        */
        if (uMsg == WM_DESTROY) {
            if (interp)
                Tcl_Release(interp);
            /* Reset the window data, just in case */
            SetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_INTERP_OFFSET, (LONG_PTR) 0);
            SetWindowLongPtrW(hwnd, TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET, (LONG_PTR) 0);
        }
    } else {
        /* Call the default window proc. This will happen during window init
           and destroy
        */
        lres = DefWindowProc(hwnd, uMsg, wParam, lParam);
    }


    return (LRESULT) lres;
}



/* Create a hidden with the specified window procedure. */
int Twapi_CreateHiddenWindow(
    Tcl_Interp *interp,         /* May be NULL */
    TwapiHiddenWindowCallbackProc winProc,
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
        memset(&w, 0, sizeof(w));
        w.cbSize = sizeof(w);
        w.hInstance = 0;
        w.cbWndExtra = TWAPI_HIDDEN_WINDOW_DATA_SIZE;
        w.lpfnWndProc = TwapiHiddenWindowProc;
        w.lpszClassName = TWAPI_HIDDEN_WINDOW_CLASS_L;

        hidden_win_class = RegisterClassExW(&w);
        if (hidden_win_class == 0) {
            /* Note interp may be null */
            return Twapi_AppendSystemError(interp, GetLastError());
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
        /* Note interp may be null */
        return Twapi_AppendSystemError(interp, GetLastError());
    }

    /* Store pointers to the interpreter and the real window function. */
    if (interp)
        Tcl_Preserve(interp);
    SetLastError(0);            /* Else no way to detect error */
    if ((SetWindowLongPtrW(win, TWAPI_HIDDEN_WINDOW_INTERP_OFFSET, (LONG_PTR) interp) != 0 || GetLastError() == 0) &&
        (SetWindowLongPtrW(win, TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET, (LONG_PTR) winProc) != 0 || GetLastError() == 0) &&
        (SetWindowLongPtrW(win, TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET, clientdata) !=0 || GetLastError() == 0)) {
        /* All success. Note interp is under a Tcl_Preserve as we have kept a pointer to it */
        if (winP)
            *winP = win;
        PostMessage(win, TWAPI_WM_HIDDEN_WINDOW_INIT, 0, 0);
        return TCL_OK;
    }

    /* One of the calls failed */
    Twapi_AppendSystemError(interp, GetLastError()); /* interp can be NULL */
    if (interp)
        Tcl_Release(interp); /* Matches Tcl_Preserve above */

    return TCL_ERROR;
}
