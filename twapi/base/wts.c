/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

void Twapi_WTSFreeMemory(void *p)
{
    if (p == NULL)
        return;
    WTSFreeMemory(p);
}

int Twapi_WTSEnumerateProcesses(Tcl_Interp *interp, HANDLE wtsH)
{
    WTS_PROCESS_INFOW *processP = NULL;
    DWORD count;
    DWORD i;
    Tcl_Obj *records = NULL;
    Tcl_Obj *fields = NULL;
    Tcl_Obj *objv[4];

    /* Note wtsH == NULL means current server */
    if (! (BOOL) (WTSEnumerateProcessesW)(wtsH, 0, 1, &processP, &count)) {
        return Twapi_AppendSystemError(interp, GetLastError());
    }


    /* Now create the data records */
    records = Tcl_NewListObj(0, NULL);
    for (i = 0; i < count; ++i) {
        Tcl_Obj *sidObj;

        /* Create a list corresponding to the fields for the process entry */
        objv[0] = Tcl_NewLongObj(processP[i].SessionId);
        objv[1] = Tcl_NewLongObj(processP[i].ProcessId);
        objv[2] = ObjFromUnicode(processP[i].pProcessName);
        if (processP[i].pUserSid) {
            if (ObjFromSID(interp, processP[i].pUserSid, &sidObj) != TCL_OK) {
                Twapi_WTSFreeMemory(processP);
                Twapi_FreeNewTclObj(records);
                return TCL_ERROR;
            }
            objv[3] = sidObj;
        }
        else {
            /* NULL SID pointer. */
            objv[3] = STRING_LITERAL_OBJ("");
        }

        Tcl_ListObjAppendElement(interp, records, Tcl_NewLongObj(processP[i].ProcessId));
        Tcl_ListObjAppendElement(interp, records, Tcl_NewListObj(4, objv));
    }

    Twapi_WTSFreeMemory(processP);

    /* Make up the field name descriptor */
    objv[0] = STRING_LITERAL_OBJ("SessionId");
    objv[1] = STRING_LITERAL_OBJ("ProcessId");
    objv[2] = STRING_LITERAL_OBJ("pProcessName");
    objv[3] = STRING_LITERAL_OBJ("pUserSid");
    fields = Tcl_NewListObj(4, objv);

    /* Put field names and records to make up the recordarray */
    objv[0] = fields;
    objv[1] = records;
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
    return TCL_OK;
}


int Twapi_WTSEnumerateSessions(Tcl_Interp *interp, HANDLE wtsH)
{
    WTS_SESSION_INFOW *sessP = NULL;
    DWORD count;
    DWORD i;
    Tcl_Obj *records = NULL;
    Tcl_Obj *fields = NULL;
    Tcl_Obj *objv[3];

    /* Note wtsH == NULL means current server */
    if (! (BOOL) (WTSEnumerateSessionsW)(wtsH, 0, 1, &sessP, &count)) {
        return Twapi_AppendSystemError(interp, GetLastError());
    }

    records = Tcl_NewListObj(0, NULL);
    for (i = 0; i < count; ++i) {
        objv[0] = Tcl_NewLongObj(sessP[i].SessionId);
        objv[1] = ObjFromUnicode(sessP[i].pWinStationName);
        objv[2] = Tcl_NewLongObj(sessP[i].State);

        /* Attach the session id as key and record to the result */
        Tcl_ListObjAppendElement(interp, records, objv[0]);
        Tcl_ListObjAppendElement(interp, records, Tcl_NewListObj(3, objv));
    }

    Twapi_WTSFreeMemory(sessP);

    /* Make up the field names */
    objv[0] = STRING_LITERAL_OBJ("SessionId");
    objv[1] = STRING_LITERAL_OBJ("pWinStationName");
    objv[2] = STRING_LITERAL_OBJ("State");
    fields = Tcl_NewListObj(3, objv);

    objv[0] = fields;
    objv[1] = records;
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
    return TCL_OK;
}

int Twapi_WTSQuerySessionInformation(
    Tcl_Interp *interp,
    HANDLE wtsH,
    DWORD  sess_id,
    WTS_INFO_CLASS info_class
)
{
    LPWSTR bufP = NULL;
    DWORD  bytes;
    DWORD winerr;

    if (! (BOOL) WTSQuerySessionInformationW(wtsH, sess_id, info_class, &bufP, &bytes )) {
        winerr = GetLastError();
        goto handle_error;

    }

    /* Note bufP can be NULL even on success! */

    switch (info_class) {
    case WTSApplicationName:
    case WTSClientDirectory:
    case WTSClientName:
    case WTSDomainName:
    case WTSInitialProgram:
    case WTSOEMId:
    case WTSUserName:
    case WTSWinStationName:
    case WTSWorkingDirectory:
        if (bufP)
            Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(bufP, -1));
        break;

    default:
        winerr = ERROR_NOT_SUPPORTED;
        goto handle_error;
    }


    /* Note bufP can be NULL even on success! */
    if (bufP)
        Twapi_WTSFreeMemory(bufP);

    return TCL_OK;

 handle_error:
    if (bufP)
        Twapi_WTSFreeMemory(bufP);

    Tcl_SetResult(interp,
                  "Could not query terminal session information. ",
                  TCL_STATIC);

    return Twapi_AppendSystemError(interp, winerr);
}


// TBD - Needs XP - DWORD WTSGetActiveConsoleSessionId(void);


