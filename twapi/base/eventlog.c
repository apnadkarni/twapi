/* 
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

/* Wrapper around ReportEvent just to rearrange argument to match typemap
 * definitions
 */
int Twapi_ReportEvent(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE hEventLog;
    WORD   wType;
    WORD  wCategory;
    DWORD dwEventID;
    PSID  lpUserSid = NULL;
    int   datalen;
    void *data;
    int     argc;
    LPCWSTR argv[32];
    int   status;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hEventLog), GETWORD(wType), GETWORD(wCategory),
                     GETINT(dwEventID), GETVAR(lpUserSid, ObjToPSID),
                     GETWARGV(argv, ARRAYSIZE(argv), argc),
                     GETBIN(data, datalen),
                     ARGEND) != TCL_OK) {
        if (lpUserSid)
            TwapiFree(lpUserSid);
        return TCL_ERROR;
    }

    if (datalen == 0)
        data = NULL;

    status = ReportEventW(hEventLog, wType, wCategory, dwEventID, lpUserSid,
                          (WORD) argc, datalen, argv, data);
    if (lpUserSid)
        TwapiFree(lpUserSid);

    return status ? TCL_OK : TwapiReturnSystemError(interp);
}
    
BOOL Twapi_IsEventLogFull(HANDLE hEventLog, int *fullP)
{
    EVENTLOG_FULL_INFORMATION evlinfo;
    DWORD bytesneeded;
    if (GetEventLogInformation(hEventLog,
                               EVENTLOG_FULL_INFO, &evlinfo,
                               sizeof(evlinfo), &bytesneeded)) {
        *fullP = evlinfo.dwFull;
        return TRUE;
    }
    return FALSE;
}

int Twapi_ReadEventLog(
    Tcl_Interp *interp,
    HANDLE evlH,
    DWORD  flags,
    DWORD  offset
    )
{
    char   buf[2048];
    DWORD  buf_sz = sizeof(buf);
    char  *bufP;
    DWORD  num_read;
    EVENTLOGRECORD *evlP;
    int result = TCL_ERROR;
    Tcl_Obj *resultObj;
    
    bufP = buf;
    if (! ReadEventLogW(evlH, flags, offset,
                        bufP, buf_sz, &num_read, &buf_sz)) {
        /* If buffer too small, dynamically allocate, else return error */
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) {
            return TwapiReturnSystemError(interp);
        }
        bufP = TwapiAlloc(buf_sz);
        /* Retry */
        if (! ReadEventLogW(evlH, flags, offset,
                            bufP, buf_sz, &num_read, &buf_sz)) {
            return TwapiReturnSystemError(interp);
        }
    }

    /* Loop through all records, adding them to the record list */
    evlP = (EVENTLOGRECORD *) bufP;
    resultObj = Tcl_NewListObj(0, NULL);
    result = TCL_OK;
    while (num_read > 0) {
        Tcl_Obj *objv[14];
        PSID     sidP;
        int      strindex;
        WCHAR   *strP;
        int      len;

        strP = (WCHAR *) (1 + &(evlP->DataOffset));
        len = lstrlenW(strP);
        objv[0] = Tcl_NewUnicodeObj(strP, len); /* Source name */
        strP += len + 1;
        objv[1] = Tcl_NewUnicodeObj(strP, -1); /* Computer name */
        objv[2] = Tcl_NewIntObj(evlP->Reserved);
        objv[3] = Tcl_NewIntObj(evlP->RecordNumber);
        objv[4] = Tcl_NewIntObj(evlP->TimeGenerated);
        objv[5] = Tcl_NewIntObj(evlP->TimeWritten);
        objv[6] = Tcl_NewIntObj(evlP->EventID);
        objv[7] = Tcl_NewIntObj(evlP->EventType);
        objv[8] = Tcl_NewIntObj(evlP->EventCategory);
        objv[9] = Tcl_NewIntObj(evlP->ReservedFlags);
        objv[10] = Tcl_NewIntObj(evlP->ClosingRecordNumber);

        /* Collect all the strings together into a list */
        objv[11] = Tcl_NewListObj(0, NULL);
        for (strP = (WCHAR *)(evlP->StringOffset + (char *)evlP), strindex = 0;
             strindex < evlP->NumStrings;
             ++strindex) {
            len = lstrlenW(strP);
            Tcl_ListObjAppendElement(interp, objv[11],
                                     Tcl_NewUnicodeObj(strP, len));
            strP += len + 1;
        }

        /* Get the SID */
        sidP = (PSID) (evlP->UserSidOffset + (char *)evlP);
        if ((evlP->UserSidLength == 0) ||
            (ObjFromSID(interp, sidP, &objv[12]) != TCL_OK)) {
            objv[12] = Tcl_NewStringObj("", 0);
        }

        /* Get the binary data */
        objv[13] =
            Tcl_NewByteArrayObj(evlP->DataOffset + (unsigned char *) evlP,
                                evlP->DataLength);


        /* Now attach this record to event record list */
        Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewListObj(14, objv));

        /* Move onto next record */
        num_read -= evlP->Length;
        evlP = (EVENTLOGRECORD *) (evlP->Length + (char *)evlP);
    }

    if (result == TCL_OK) {
        Tcl_SetObjResult(interp, resultObj);
    }

    if (bufP != buf)
        TwapiFree(bufP);
    return result;
}



