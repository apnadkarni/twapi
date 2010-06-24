/* 
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

typedef struct _TwapiConsoleCtrlCallback {
    TwapiPendingCallback pcb;   /* Must be first field */
    DWORD ctrl;
} TwapiConsoleCtrlCallback;

static TwapiInterpContext * volatile console_control_ticP;

static BOOL WINAPI TwapiConsoleCtrlHandler(DWORD ctrl_event);
static DWORD TwapiConsoleCtrlCallbackFn(TwapiPendingCallback *pcbP);


int Twapi_ReadConsole(Tcl_Interp *interp, HANDLE conh, unsigned int numchars)
{
    WCHAR  buf[256];
    WCHAR *bufP = buf;
    DWORD  len;
    int status = TCL_ERROR;

    if (numchars > ARRAYSIZE(buf)) {
        if (Twapi_malloc(interp, NULL, sizeof(WCHAR) * numchars, &bufP) != TCL_OK)
            return TCL_ERROR;
    }

    if (! ReadConsoleW(conh, bufP, numchars, &len, NULL)) {
        TwapiReturnSystemError(interp);
        goto vamoose;
    }

    Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(buf, len));
    status = TCL_OK;

vamoose:
    if (bufP != buf)
        free(bufP);
    return status;
}

int Twapi_StartConsoleEventNotifier(TwapiInterpContext *ticP)
{
    void *pv;

    ERROR_IF_UNTHREADED(ticP->interp);
    pv = InterlockedCompareExchangePointer(&console_control_ticP,
                                           ticP, NULL);
    if (pv) {
        Tcl_SetResult(ticP->interp, "Console control handler is already set.", TCL_STATIC);
        return TCL_ERROR;
    }

    if (SetConsoleCtrlHandler(TwapiConsoleCtrlHandler, TRUE)) {
        ticP->console_ctrl_hooked = 1;
        TwapiInterpContextRef(ticP, 1);
        return TCL_OK;
    }
    else {
        InterlockedExchangePointer(&console_control_ticP, NULL);
        return TwapiReturnSystemError(ticP->interp);
    }
}

    
int Twapi_StopConsoleEventNotifier(TwapiInterpContext *ticP)
{
    void *pv;
    pv = InterlockedCompareExchangePointer(&console_control_ticP,
                                           NULL, ticP);
    if (pv != (void*) ticP) {
        Tcl_SetResult(ticP->interp, "Console control handler not set by this interpreter.", TCL_STATIC);
        return TCL_ERROR;
    }
    SetConsoleCtrlHandler(TwapiConsoleCtrlHandler, FALSE);
    ticP->console_ctrl_hooked = 0;
    TwapiInterpContextUnref(ticP, 1);
    return TCL_OK;
}


/* Directly called by Windows in a separate thread */
static BOOL WINAPI TwapiConsoleCtrlHandler(DWORD ctrl)
{
    TwapiConsoleCtrlCallback *cbP;
    WCHAR *resultP;
    BOOL handled = FALSE;
    WCHAR *w;

    /* TBD - there is a race here. */
    if (console_control_ticP == NULL)
        return FALSE;

    cbP = (TwapiConsoleCtrlCallback *) TwapiPendingCallbackNew(
        console_control_ticP, TwapiConsoleCtrlCallbackFn, sizeof(*cbP));

    cbP->ctrl = ctrl;
    if (TwapiEnqueueCallback(console_control_ticP,
                             (TwapiPendingCallback*) cbP,
                             TWAPI_ENQUEUE_DIRECT,
                             100, /* Timeout (ms) */
                             (TwapiPendingCallback **)&cbP)
        == ERROR_SUCCESS) {

        if (cbP && cbP->pcb.response.type == TRT_BOOL)
            handled = cbP->pcb.response.value.bval;
    }

    if (cbP)
        TwapiPendingCallbackUnref((TwapiPendingCallback *)cbP, 1);

    return handled;
}


static DWORD TwapiConsoleCtrlCallbackFn(TwapiPendingCallback *pcbP)
{
    TwapiConsoleCtrlCallback *cbP = (TwapiConsoleCtrlCallback *)pcbP;
    char *event_str;
    Tcl_Obj *objs[3];

    switch (cbP->ctrl) {
    case CTRL_C_EVENT:
        event_str = "ctrl-c";
        break;
    case CTRL_BREAK_EVENT:
        event_str = "ctrl-break";
        break;
    case CTRL_CLOSE_EVENT:
        event_str = "close";
        break;
    case CTRL_LOGOFF_EVENT:
        event_str = "logoff";
        break;
    case CTRL_SHUTDOWN_EVENT:
        event_str = "shutdown";
        break;
    default:
        // Unknown event type
        return ERROR_INVALID_PARAMETER;
    }

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_console_ctrl_handler", -1);
    objs[1] = Tcl_NewStringObj(event_str, -1);
    return TwapiEvalAndUpdateCallback(pcbP, 2, objs, TRT_BOOL);
}
