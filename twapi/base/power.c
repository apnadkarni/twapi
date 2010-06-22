/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

typedef struct _TwapiPowerCallback {
    TwapiPendingCallback pcb;   /* Must be first field */
    WPARAM wparam;
    LPARAM lparam;
} TwapiPowerCallback;

int Twapi_PowerNotifyStart(TwapiInterpContext *ticP)
{
    HWND hwnd;

    // Get the common notification window.
    hwnd = TwapiGetNotificationWindow(ticP);
    if (hwnd == NULL)
        return TCL_ERROR;

    ticP->power_events_on = 1;
    return TCL_OK;
}


int Twapi_PowerNotifyStop(TwapiInterpContext *ticP)
{
    /*
     * Note since we are using the common window for notifications,
     * we do not destroy it. Just mark power events off.
     */
    ticP->power_events_on = 0;
    return TCL_OK;
}


/* Called (indirectly) from the Tcl notifier loop with a new power event.
 * Constructs the script to be invoked in the interpreter.
 * Follows behaviour specified by TwapiCallbackFn typedef.
 */
DWORD TwapiPowerCallbackFn(TwapiPendingCallback *pcbP, Tcl_Obj **objPP)
{
    Tcl_Obj *objs[3];

    TwapiPowerCallback *cbP = (TwapiPowerCallback *) pcbP;

    if (objPP == NULL)
        return ERROR_SUCCESS;   /* Allowed to be NULL */

    objs[0] = STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_power_handler");
    objs[1] = ObjFromDWORD_PTR(cbP->wparam);
    objs[2] = ObjFromDWORD_PTR(cbP->lparam);
    *objPP = Tcl_NewListObj(3, objs);
    return ERROR_SUCCESS;
}


/*
 * Called from the notification window message handler when a hotkey message
 * is received. Allocates and enqueues a callback.
 */
LRESULT TwapiPowerHandler(TwapiInterpContext *ticP, UINT msg, WPARAM wparam, LPARAM lparam)
{
    TwapiPowerCallback *cbP;

    cbP = (TwapiPowerCallback *) TwapiPendingCallbackNew(
        ticP, TwapiPowerCallbackFn, sizeof(*cbP));

    cbP->wparam = wparam;
    cbP->lparam = lparam;
    TwapiEnqueueCallback(ticP, (TwapiPendingCallback*) cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);

    /* For querysuspend, make sure we allow the suspend */
    return (LRESULT) (wparam == PBT_APMQUERYSUSPEND ? TRUE : 0);
}
