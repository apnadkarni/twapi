/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

int Twapi_PowerNotifyStart(TwapiInterpContext *ticP)
{
    HWND hwnd;

    // Get the common notification window.
    hwnd = Twapi_GetNotificationWindow(ticP);
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


/* TBD - move to scrpt level code like hotkeys ? */
/* Called (indirectly) from the Tcl notifier loop with a new power event.
 * Constructs the script to be invoked in the interpreter.
 * Follows behaviour specified by TwapiCallbackFn typedef.
 */
DWORD TwapiPowerCallbackFn(TwapiCallback *cbP)
{
    Tcl_Obj *objs[3];

    objs[0] = STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_power_handler");
    objs[1] = ObjFromDWORD_PTR(cbP->clientdata);
    objs[2] = ObjFromDWORD_PTR(cbP->clientdata2);
    /* TBD - do power events have a response ? */
    return TwapiEvalAndUpdateCallback(cbP, 3, objs, TRT_EMPTY);
}


/*
 * Called from the notification window message handler when a hotkey message
 * is received. Allocates and enqueues a callback.
 */
LRESULT TwapiPowerHandler(TwapiInterpContext *ticP, UINT msg, WPARAM wparam, LPARAM lparam)
{
    TwapiCallback *cbP;

    if (msg == WM_POWERBROADCAST) {
        cbP = TwapiCallbackNew(ticP, TwapiPowerCallbackFn, sizeof(*cbP));

        cbP->clientdata = wparam;
        cbP->clientdata2 = lparam;
        TwapiEnqueueCallback(ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);

        /* For querysuspend, make sure we allow the suspend - TBD */
        return (LRESULT) (wparam == PBT_APMQUERYSUSPEND ? TRUE : 0);
    }
    return (LRESULT) NULL;
}
