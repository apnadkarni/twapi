/*
 * Copyright (c) 2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

/* TBD - move to script level code */
int Twapi_RegisterHotKey(TwapiInterpContext *ticP, int id, UINT modifiers, UINT vk)
{
    HWND hwnd;

    // Get the common notification window.
    hwnd = Twapi_GetNotificationWindow(ticP);
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
    hwnd = Twapi_GetNotificationWindow(ticP);
    
    if (UnregisterHotKey(hwnd, id))
        return TCL_OK;
    else
        return TwapiReturnSystemError(ticP->interp);
}


