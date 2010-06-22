/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

typedef struct _TwapiHotkeyCallback {
    TwapiPendingCallback pcb;   /* Must be first field */
    WPARAM wparam;
    LPARAM lparam;
} TwapiHotkeyCallback;

int Twapi_RegisterHotKey(TwapiInterpContext *ticP, int id, UINT modifiers, UINT vk)
{
    HWND hwnd;

    ERROR_IF_UNTHREADED(ticP->interp);

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
DWORD TwapiHotkeyCallbackFn(TwapiPendingCallback *pcbP, Tcl_Obj **objPP)
{
    Tcl_Obj *objs[3];

    TwapiHotkeyCallback *hkcbP = (TwapiHotkeyCallback *) pcbP;

    if (objPP == NULL)
        return ERROR_SUCCESS;   /* Allowed to be NULL */

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_hotkey_handler", -1);
    objs[1] = ObjFromDWORD_PTR(hkcbP->wparam);
    objs[2] = ObjFromDWORD_PTR(hkcbP->lparam);
    *objPP = Tcl_NewListObj(3, objs);
    return ERROR_SUCCESS;
}


/*
 * Called from the notification window message handler when a hotkey message
 * is received. Allocates and enqueues a callback.
 */
LRESULT TwapiHotkeyHandler(TwapiInterpContext *ticP, WPARAM id, LPARAM key)
{
    TwapiHotkeyCallback *hkcbP;

    hkcbP = (TwapiHotkeyCallback *) TwapiPendingCallbackNew(
        ticP, TwapiHotkeyCallbackFn, sizeof(*hkcbP));

    hkcbP->wparam = id;
    hkcbP->lparam = key;
    TwapiEnqueueCallback(ticP, (TwapiPendingCallback*) hkcbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);
    return (LRESULT) NULL;
}
