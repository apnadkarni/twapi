/* 
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

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
    TwapiInterpContext *ticP,
    HANDLE evlH,
    DWORD  flags,
    DWORD  offset
    )
{
    DWORD  buf_sz;
    char  *bufP;
    DWORD  num_read;
    EVENTLOGRECORD *evlP;
    Tcl_Obj *resultObj;
    DWORD winerr = ERROR_SUCCESS;
    Tcl_Interp *interp = ticP->interp;
    
    /* Ask for 1000 bytes alloc, will get more if available */
    bufP = MemLifoPushFrame(&ticP->memlifo, 1000, &buf_sz);

    if (! ReadEventLogW(evlH, flags, offset,
                        bufP, buf_sz, &num_read, &buf_sz)) {
        /* If buffer too small, dynamically allocate, else return error */
        winerr = GetLastError();
        if (winerr != ERROR_INSUFFICIENT_BUFFER)
            goto vamoose;
        /*
         * Don't bother popping the memlifo frame, just alloc new.
         * We allocated max in current memlifo chunk above anyways. Also,
         * remember MemLifoResize will do unnecessary copy so we don't use it.
         */
        bufP = MemLifoAlloc(&ticP->memlifo, buf_sz, NULL);
        /* Retry */
        if (! ReadEventLogW(evlH, flags, offset,
                            bufP, buf_sz, &num_read, &buf_sz)) {
            winerr = GetLastError();
            goto vamoose;
        }
        winerr = 0;             /* Reset for vamoose */
    }

    /* Loop through all records, adding them to the record list */
    evlP = (EVENTLOGRECORD *) bufP;
    resultObj = Tcl_NewListObj(0, NULL);
    while (num_read > 0) {
        Tcl_Obj *objv[14];
        PSID     sidP;
        int      strindex;
        WCHAR   *strP;
        int      len;

        strP = (WCHAR *) (1 + &(evlP->DataOffset));
        len = lstrlenW(strP);
        objv[0] = ObjFromUnicodeN(strP, len); /* Source name */
        strP += len + 1;
        objv[1] = ObjFromUnicode(strP); /* Computer name */
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
                                     ObjFromUnicodeN(strP, len));
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

vamoose:
    MemLifoPopFrame(&ticP->memlifo);
    if (winerr == ERROR_SUCCESS) {
        Tcl_SetObjResult(interp, resultObj);
        return TCL_OK;
    } else {
        Twapi_AppendSystemError(interp, winerr);
        return TCL_ERROR;
    }
}


static int Twapi_EventlogCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    DWORD dw, dw2;
    LPWSTR s, s2;
    HANDLE h, h2;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 100) {
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), ARGUSEDEFAULT, GETHANDLE(h2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        
        if ((func == 1 && objc != 4) || objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 1:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = NotifyChangeEventLog(h, h2);
            break;
        case 2:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseEventLog(h);
            break;
        case 3:
            result.type = GetNumberOfEventLogRecords(h,
                                                     &result.value.ival)
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 4:
            result.type = GetOldestEventLogRecord(h,
                                                  &result.value.ival) 
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 5:
            result.type = Twapi_IsEventLogFull(h,
                                               &result.value.ival) 
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 6:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeregisterEventSource(h);
            break;
        }
    } else {
        /* Arbitrary args */
        switch (func) {
        case 1001:
            if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(h),
                             GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_ReadEventLog(ticP, h, dw, dw2);
        case 1002:
        case 1003:
            if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(h),
                             GETWSTR(s), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = (func == 1002 ? BackupEventLogW : ClearEventLogW)(h, s);
            break;
        case 1004:
        case 1005:
        case 1006:
            if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(h),
                             GETWSTR(s), GETWSTR(s2), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            result.value.hval = (func == 1004 ? OpenEventLogW :
                                 (func == 1006 ? OpenBackupEventLogW : RegisterEventSourceW))(s, s2);
            break;
        case 1007:
            return Twapi_ReportEvent(interp, objc-2, objv+2);
        }
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_EventlogInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::EventlogCall", Twapi_EventlogCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::Eventlog" #call_, # code_); \
    } while (0);

    CALL_(NotifyChangeEventLog, Call, 1);
    CALL_(CloseEventLog, Call, 2);
    CALL_(GetNumberOfEventLogRecords, Call, 3);
    CALL_(GetOldestEventLogRecord, Call, 4);
    CALL_(Twapi_IsEventLogFull, Call, 5);
    CALL_(DeregisterEventSource, CallH, 6);
    CALL_(ReadEventLog, Call, 1001);
    CALL_(BackupEventLog, Call, 1002);
    CALL_(ClearEventLog, Call, 1003);
    CALL_(OpenEventLog, Call, 1004);
    CALL_(OpenBackupEventLog, Call, 1005);
    CALL_(RegisterEventSource, Call, 1006);
    CALL_(ReportEvent, Call, 1007);


#undef CALL_

    return TCL_OK;
}


#ifndef TWAPI_SINGLE_MODULE
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_SINGLE_MODULE
__declspec(dllexport) 
#endif
int Twapi_eventlog_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }


    return Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_EventlogInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

