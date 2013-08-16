/* 
 * Copyright (c) 2004-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_eventlog.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

static TwapiOneTimeInitState TwapiEventlogOneTimeInitialized;

    
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

static int Twapi_ReadEventLogObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE evlH;
    DWORD  flags;
    DWORD  offset;
    DWORD  buf_sz;
    char  *bufP;
    DWORD  num_read;
    int    i;
    EVENTLOGRECORD *evlP;
    Tcl_Obj *resultObj;
    DWORD winerr = ERROR_SUCCESS;
    static const char *fieldnames[] = {
        "-source", "-system", "-reserved", "-recordnum", "-timegenerated",
        "-timewritten", "-eventid", "-level", "-category", "-reservedflags",
        "-closingrecnum", "-params", "-sid", "-data" };

    Tcl_Obj *fields[ARRAYSIZE(fieldnames)];

    if (TwapiGetArgs(interp, objc-1, objv+1, GETHANDLE(evlH),
                     GETINT(flags), GETINT(offset), ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* Ask for 1000 bytes alloc, will get more if available. TBD - instrument */
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

    /*
     * Loop through all records, adding them to the record list
     * We use cached field name objects for efficiency. Note these
     * need not/should not be explicitly freed.
     */
    for (i = 0; i < ARRAYSIZE(fields); ++i) {
        fields[i] = TwapiGetAtom(ticP, fieldnames[i]);
    }
    
    evlP = (EVENTLOGRECORD *) bufP;
    resultObj = ObjNewList(0, NULL);
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
        objv[2] = ObjFromDWORD(evlP->Reserved);
        objv[3] = ObjFromDWORD(evlP->RecordNumber);
        objv[4] = ObjFromDWORD(evlP->TimeGenerated);
        objv[5] = ObjFromDWORD(evlP->TimeWritten);
        objv[6] = ObjFromDWORD(evlP->EventID);
        objv[7] = ObjFromInt(evlP->EventType);
        objv[8] = ObjFromInt(evlP->EventCategory);
        objv[9] = ObjFromInt(evlP->ReservedFlags);
        objv[10] = ObjFromDWORD(evlP->ClosingRecordNumber);

        /* Collect all the strings together into a list */
        objv[11] = ObjNewList(0, NULL);
        for (strP = (WCHAR *)(evlP->StringOffset + (char *)evlP), strindex = 0;
             strindex < evlP->NumStrings;
             ++strindex) {
            len = lstrlenW(strP);
            ObjAppendElement(interp, objv[11],
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
            ObjFromByteArray(evlP->DataOffset + (unsigned char *) evlP,
                                evlP->DataLength);

        /* Now attach this record to event record list */
        TWAPI_ASSERT(ARRAYSIZE(objv) == ARRAYSIZE(fields));
        ObjAppendElement(interp, resultObj, TwapiTwineObjv(fields, objv, ARRAYSIZE(objv)));

        /* Move onto next record */
        num_read -= evlP->Length;
        evlP = (EVENTLOGRECORD *) (evlP->Length + (char *)evlP);
    }

vamoose:
    MemLifoPopFrame(&ticP->memlifo);
    if (winerr == ERROR_SUCCESS) {
        ObjSetResult(interp, resultObj);
        return TCL_OK;
    } else {
        Twapi_AppendSystemError(interp, winerr);
        return TCL_ERROR;
    }
}


static int Twapi_EventlogCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func = PtrToInt(clientdata);
    HANDLE h, h2;

    --objc;
    ++objv;
    result.type = TRT_BADFUNCTIONCODE;
    if (func < 100) {
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLE(h), ARGUSEDEFAULT, GETHANDLE(h2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        
        /* func 1 has 2 args, rest all have 1 arg */
        if (func == 2) {
            if (objc != 2)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = NotifyChangeEventLog(h, h2);
        } else {
            if (objc != 1)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            switch (func) {
            case 3:
                result.type = GetNumberOfEventLogRecords(h,
                                                         &result.value.uval)
                    ? TRT_DWORD : TRT_GETLASTERROR;
                break;
            case 4:
                result.type = GetOldestEventLogRecord(h,
                                                      &result.value.uval) 
                    ? TRT_DWORD : TRT_GETLASTERROR;
                break;
            case 5:
                result.type = Twapi_IsEventLogFull(h,
                                                   &result.value.ival) 
                    ? TRT_LONG : TRT_GETLASTERROR;
                break;
            }
        }
    } else {
        /* Exactly 2 args */
        CHECK_NARGS(interp, objc, 2);
        switch (func) {
        case 1002:
            if (ObjToLPVOID(interp, objv[0], &h) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = BackupEventLogW(h, ObjToUnicode(objv[1]));
            break;
        case 1003:
            if (ObjToLPVOID(interp, objv[0], &h) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ClearEventLogW(h, ObjToLPWSTR_NULL_IF_EMPTY(objv[1]));
            break;
        case 1004:
            result.type = TRT_HANDLE;
            result.value.hval = OpenBackupEventLogW(
                ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
                ObjToUnicode(objv[1]));
            break;
        }
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_EventlogInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s EvlogCallDispatch[] = {
        DEFINE_FNCODE_CMD(NotifyChangeEventLog, 2),
        DEFINE_FNCODE_CMD(GetNumberOfEventLogRecords, 3),
        DEFINE_FNCODE_CMD(GetOldestEventLogRecord, 4),
        DEFINE_FNCODE_CMD(Twapi_IsEventLogFull, 5),
        DEFINE_FNCODE_CMD(BackupEventLog, 1002),
        DEFINE_FNCODE_CMD(ClearEventLog, 1003),
        DEFINE_FNCODE_CMD(OpenBackupEventLog, 1004),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(EvlogCallDispatch), EvlogCallDispatch, Twapi_EventlogCallObjCmd);

    Tcl_CreateObjCommand(interp, "twapi::ReadEventLog", Twapi_ReadEventLogObjCmd, ticP, NULL);

    return Twapi_EvtInitCalls(interp, ticP);
}

static int TwapiEventlogOneTimeInit(Tcl_Interp *interp)
{
    TwapiInitEvtStubs(interp);
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
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        Twapi_EventlogInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    if (! TwapiDoOneTimeInit(&TwapiEventlogOneTimeInitialized,
                             TwapiEventlogOneTimeInit, interp))
        return TCL_ERROR;


    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

