/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

int Twapi_RegisterHotKey(TwapiInterpContext *ticP, int id, UINT modifiers, UINT vk)
{
    HWND hwnd;

    // Get the common notification window.
    hwnd = TwapiGetNotificationWindow(ticP);
    if (hwnd == NULL)
        return TCL_ERROR;

    if (RegisterHotKey(hwnd, id, modifiers, vk))
        return TCL_OK;
    else
        return TwapiReturnSystemError(ticP->interp);
}


int Twapi_UnregisterHotKey(TwapiInterpContext *ticP, int id)
{
    // Note since we are using the common window for notifications,
    // we do not destroy it. Just unregister the hot key.
    HWND hwnd;
    hwnd = TwapiGetNotificationWindow(ticP);
    
    if (UnregisterHotKey(hwnd, id))
        return TCL_OK;
    else
        return TwapiReturnSystemError(ticP->interp);
}


/* Called (indirectly) from the Tcl notifier loop with a new hotkey event.
 * Constructs the hotkey script to be invoked in the interpreter.
 * Follows behaviour specified by TwapiCallbackFn typedef.
 */
int TwapiHotkeyCallbackFn(TwapiCallback *cbP)
{
    Tcl_Obj *objs[3];

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_hotkey_handler", -1);
    objs[1] = ObjFromDWORD_PTR(cbP->clientdata);
    objs[2] = ObjFromDWORD_PTR(cbP->clientdata2);
    return TwapiEvalAndUpdateCallback(cbP, 3, objs, TRT_EMPTY);
}


/*
 * Called from the notification window message handler when a hotkey message
 * is received. Allocates and enqueues a callback.
 */
LRESULT TwapiHotkeyHandler(TwapiInterpContext *ticP, UINT msg, WPARAM id, LPARAM key)
{
    TwapiCallback *cbP;

    if (msg == WM_HOTKEY) {
        cbP = TwapiCallbackNew(ticP, TwapiHotkeyCallbackFn, sizeof(*cbP));
        cbP->clientdata = id;
        cbP->clientdata2 = key;
        TwapiEnqueueCallback(ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);
    }

    return (LRESULT) NULL;
}
