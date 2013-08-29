/*
 * Copyright (c) 2003-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows NT services */
#include "twapi.h"
#include "twapi_service.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;
#endif

/* Map service state int to string */
static char *ServiceStateString(DWORD state)
{
    switch (state) {
    case 0: return "0";
    case 1: return "stopped";
    case 2: return "start_pending";
    case 3: return "stop_pending";
    case 4: return "running";
    case 5: return "continue_pending";
    case 6: return "pause_pending";
    case 7: return "paused";
    default: return "unknown";
    }
}

static char *ServiceTypeString(DWORD service_type)
{
    service_type &=  ~SERVICE_INTERACTIVE_PROCESS;
    switch (service_type) {
    case 0x1: return "kernel_driver";
    case 0x2: return "file_system_driver";
    case 0x4: return "adapter";
    case 0x8: return "recognizer_driver";
    case 0x10: return "win32_own_process";
    case 0x20: return "win32_share_process";
    default: return "unknown";
    }
}

static int Twapi_QueryServiceStatusEx(Tcl_Interp *interp, SC_HANDLE h,
                                      SC_STATUS_TYPE infolevel)
{
    SERVICE_STATUS_PROCESS ssp;
    DWORD numbytes;
    DWORD result;
    Tcl_Obj *rec[20];
    char *cP;

    if (infolevel != SC_STATUS_PROCESS_INFO) {
        return Twapi_AppendSystemError(interp, ERROR_INVALID_LEVEL);
    }

    result = QueryServiceStatusEx(h, infolevel, (LPBYTE) &ssp, sizeof(ssp), &numbytes);
    if (! result) {
        return TwapiReturnSystemError(interp);
    }

    /* NOTE: Unlike other functions, we do not use TwapiGetAtom here
       because we are returning only one record and to not want to
       clutter the atom table
    */
    rec[0] = STRING_LITERAL_OBJ("servicetype");
    cP = ServiceTypeString(ssp.dwServiceType);
    rec[1] = cP ? ObjFromString(cP) : ObjFromDWORD(ssp.dwServiceType);
    rec[2] = STRING_LITERAL_OBJ("state");
    cP = ServiceStateString(ssp.dwCurrentState);
    rec[3] = cP ? ObjFromString(cP) : ObjFromDWORD(ssp.dwCurrentState);
    rec[4] = STRING_LITERAL_OBJ("control_accepted");
    rec[5] = ObjFromDWORD(ssp.dwControlsAccepted);
    rec[6] = STRING_LITERAL_OBJ("exitcode");
    rec[7] = ObjFromDWORD(ssp.dwWin32ExitCode);
    rec[8] = STRING_LITERAL_OBJ("service_code");
    rec[9] = ObjFromDWORD(ssp.dwServiceSpecificExitCode);
    rec[10] = STRING_LITERAL_OBJ("checkpoint");
    rec[11] = ObjFromDWORD(ssp.dwCheckPoint);
    rec[12] = STRING_LITERAL_OBJ("wait_hint");
    rec[13] = ObjFromDWORD(ssp.dwWaitHint);
    rec[14] = STRING_LITERAL_OBJ("pid");
    rec[15] = ObjFromDWORD(ssp.dwProcessId);
    rec[16] = STRING_LITERAL_OBJ("serviceflags");
    rec[17] = ObjFromDWORD(ssp.dwServiceFlags);
    rec[18] = STRING_LITERAL_OBJ("interactive");
    rec[19] = ObjFromBoolean(ssp.dwServiceType & SERVICE_INTERACTIVE_PROCESS);

    ObjSetResult(interp, ObjNewList(ARRAYSIZE(rec), rec));
    return TCL_OK;
}


int Twapi_QueryServiceConfig(TwapiInterpContext *ticP, SC_HANDLE hService)
{
    QUERY_SERVICE_CONFIGW *qbuf;
    DWORD buf_sz;
    Tcl_Obj *objv[20];
    DWORD winerr;
    int   tcl_result = TCL_ERROR;
    const char *cP;

    /* Ask for 1000 bytes alloc, will get more if available */
    qbuf = (QUERY_SERVICE_CONFIGW *) MemLifoPushFrame(&ticP->memlifo,
                                                      1000, &buf_sz);

    if (! QueryServiceConfigW(hService, qbuf, buf_sz, &buf_sz)) {
        /* For any error other than size, return */
        winerr = GetLastError();
        if (winerr != ERROR_INSUFFICIENT_BUFFER)
            goto vamoose;

        /* Retry allocating specified size */
        /*
         * Don't bother popping the memlifo frame, just alloc new.
         * We allocated max in current memlifo chunk above anyways. Also,
         * remember MemLifoResize will do unnecessary copy so we don't use it.
         */
        qbuf = (QUERY_SERVICE_CONFIGW *) MemLifoAlloc(&ticP->memlifo, buf_sz, NULL);

        /* Get the configuration information.  */
        if (! QueryServiceConfigW(hService, qbuf, buf_sz, &buf_sz)) {
            winerr = GetLastError();
            goto vamoose;
        }
    }

    objv[0] = STRING_LITERAL_OBJ("-dependencies");
    objv[1] = ObjFromMultiSz_MAX(qbuf->lpDependencies);
    objv[2] = STRING_LITERAL_OBJ("-servicetype");
    cP = ServiceTypeString(qbuf->dwServiceType);
    objv[3] = cP ? Tcl_NewStringObj(cP, -1) : ObjFromDWORD(qbuf->dwServiceType);
    objv[4] = STRING_LITERAL_OBJ("-starttype");
    objv[5] = ObjFromDWORD(qbuf->dwStartType);
    objv[6] = STRING_LITERAL_OBJ("-errorcontrol");
    objv[7] = ObjFromDWORD(qbuf->dwErrorControl);
    objv[8] = STRING_LITERAL_OBJ("-tagid");
    objv[9] = ObjFromDWORD(qbuf->dwTagId);
    objv[10] = STRING_LITERAL_OBJ("-command");
    objv[11] = ObjFromUnicode(qbuf->lpBinaryPathName);
    objv[12] = STRING_LITERAL_OBJ("-loadordergroup");
    objv[13] = ObjFromUnicode(qbuf->lpLoadOrderGroup);
    objv[14] = STRING_LITERAL_OBJ("-account");
    objv[15] = ObjFromUnicode(qbuf->lpServiceStartName);
    objv[16] = STRING_LITERAL_OBJ("-displayname");
    objv[17] = ObjFromUnicode(qbuf->lpDisplayName);
    objv[18] = STRING_LITERAL_OBJ("-interactive");
    objv[19] = ObjFromBoolean(qbuf->dwServiceType & SERVICE_INTERACTIVE_PROCESS);
    ObjSetResult(ticP->interp, ObjNewList(20,objv));
    tcl_result = TCL_OK;

vamoose:
    MemLifoPopFrame(&ticP->memlifo);
    if (tcl_result != TCL_OK)
        Twapi_AppendSystemError(ticP->interp, winerr);

    return tcl_result;
}


int Twapi_QueryServiceConfig2(TwapiInterpContext *ticP, SC_HANDLE hService, DWORD level)
{
    LPSERVICE_DESCRIPTIONW bufP;
    DWORD buf_sz;
    DWORD winerr;
    int   tcl_result = TCL_ERROR;

    if (level != SERVICE_CONFIG_DESCRIPTION) 
        return TwapiReturnError(ticP->interp, TWAPI_INVALID_OPTION);

    /* Ask for 256 bytes alloc, will get more if available */
    bufP = (LPSERVICE_DESCRIPTIONW) MemLifoPushFrame(&ticP->memlifo,
                                                      256, &buf_sz);

    if (! QueryServiceConfig2W(hService, level, (LPBYTE) bufP, buf_sz, &buf_sz)) {
        /* For any error other than size, return */
        winerr = GetLastError();
        if (winerr != ERROR_INSUFFICIENT_BUFFER)
            goto vamoose;

        /* Retry allocating specified size */
        /*
         * Don't bother popping the memlifo frame, just alloc new.
         * We allocated max in current memlifo chunk above anyways. Also,
         * remember MemLifoResize will do unnecessary copy so we don't use it.
         */
        bufP = (LPSERVICE_DESCRIPTIONW) MemLifoAlloc(&ticP->memlifo, buf_sz, NULL);

        /* Get the configuration information.  */
        if (! QueryServiceConfig2W(hService, level, (LPBYTE) bufP, buf_sz, &buf_sz)) {
            winerr = GetLastError();
            goto vamoose;
        }
    }

    /* If NULL, we keep result as empty string. Not an error */
    if (bufP->lpDescription)
        ObjSetResult(ticP->interp, ObjFromUnicode(bufP->lpDescription));
    tcl_result = TCL_OK;

vamoose:
    MemLifoPopFrame(&ticP->memlifo);
    if (tcl_result != TCL_OK)
        Twapi_AppendSystemError(ticP->interp, winerr);

    return tcl_result;
}


//#define NOOP_BEYOND_VISTA // TBD - remove this functionality ?
#ifdef NOOP_BEYOND_VISTA
/*
 * Helper function to retrieve service lock status info
 * Returns a single dynamic block block after figuring out required size
 */
int  Twapi_QueryServiceLockStatus(
    TwapiInterpContext *ticP,
    SC_HANDLE hService
    )
{
    QUERY_SERVICE_LOCK_STATUSW *sbuf;
    DWORD buf_sz;
    Tcl_Obj *obj;

    /* TBD - Note that the this function is a no-op on vista and beyond */

    /* First find out how big a buffer we need */
    if (! QueryServiceLockStatusW(hService, NULL, 0, &buf_sz)) {
        /* For any error other than size, return */
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
            return TwapiReturnSystemError(ticP->interp);
    }

    /* Allocate it */
    sbuf = MemLifoPushFrame(&ticP->memlifo, buf_sz, NULL);
    /* Get the configuration information.  */
    if (! QueryServiceLockStatusW(hService, sbuf, buf_sz, &buf_sz)) {
        TwapiReturnSystemError(interp); // Store before call to free
        MemLifoPopFrame(&ticP->memlifo);
        return TCL_ERROR;
    }

    obj = ObjEmptyList();
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, fIsLocked);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, obj, sbuf, lpLockOwner);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, dwLockDuration);

    MemLifoPopFrame(&ticP->memlifo);

    ObjSetResult(interp, obj);

    return TCL_OK;
}
#endif

/*
 * Helper function to retrieve list of services and status
 */
int Twapi_EnumServicesStatusEx(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    Tcl_Interp *interp = ticP->interp;
    SC_HANDLE hService;
    SC_ENUM_TYPE infolevel;
    DWORD     dwServiceType;
    DWORD     dwServiceState;
    ENUM_SERVICE_STATUS_PROCESSW *sbuf;
    DWORD buf_sz;
    DWORD buf_needed;
    DWORD services_returned;
    DWORD resume_handle;
    BOOL  success;
    Tcl_Obj *resultObj, *groupnameObj;
    Tcl_Obj *rec[12];    /* Holds values for each status record */
    Tcl_Obj *keys[ARRAYSIZE(rec)];
    Tcl_Obj *states[8]; /* Holds shared objects for states 1-7, 0 unused */
    DWORD winerr;
    DWORD i;
    TCL_RESULT status = TCL_ERROR;
    Tcl_Obj *service_types[6];

    if (TwapiGetArgs(interp, objc, objv,
                     GETPTR(hService, SC_HANDLE), GETINT(infolevel),
                     GETINT(dwServiceType),
                     GETINT(dwServiceState), GETOBJ(groupnameObj),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (infolevel != SC_ENUM_PROCESS_INFO) {
        ObjSetStaticResult(interp, "Unsupported information level");
        return Twapi_AppendSystemError(interp, ERROR_INVALID_PARAMETER);
    }

    /* 32000 - Initial estimate based on my system - TBD */
    sbuf = MemLifoPushFrame(&ticP->memlifo, 32000, &buf_sz);

    /* NOTE - the returned Tcl_Obj from TwapiGetAtom must not be freed */

    /* Create the map of service types codes to strings */
    service_types[0] = TwapiGetAtom(ticP, ServiceTypeString(0x1));
    service_types[1] = TwapiGetAtom(ticP, ServiceTypeString(0x2));
    service_types[2] = TwapiGetAtom(ticP, ServiceTypeString(0x4));
    service_types[3] = TwapiGetAtom(ticP, ServiceTypeString(0x8));
    service_types[4] = TwapiGetAtom(ticP, ServiceTypeString(0x10));
    service_types[5] = TwapiGetAtom(ticP, ServiceTypeString(0x20));

    /* And the state symbols ... */
    for (i=0; i < ARRAYSIZE(states); ++i) {
        states[i] = TwapiGetAtom(ticP, ServiceStateString(i));
    }

    /* Note order of names should be same as order of values below */
    /* Note order of field names should be same as order of values above */
    keys[0] = TwapiGetAtom(ticP, "servicetype");
    keys[1] = TwapiGetAtom(ticP, "state");
    keys[2] = TwapiGetAtom(ticP, "control_accepted");
    keys[3] = TwapiGetAtom(ticP, "exitcode");
    keys[4] = TwapiGetAtom(ticP, "service_code");
    keys[5] = TwapiGetAtom(ticP, "checkpoint");
    keys[6] = TwapiGetAtom(ticP, "wait_hint");
    keys[7] = TwapiGetAtom(ticP, "pid");
    keys[8] = TwapiGetAtom(ticP, "serviceflags");
    keys[9] = TwapiGetAtom(ticP, "name");
    keys[10] = TwapiGetAtom(ticP, "displayname");
    keys[11] = TwapiGetAtom(ticP, "interactive");

    resultObj = Tcl_NewDictObj();
    resume_handle = 0;
    do {
        /* Note we don't actually make use of buf_needed, just reuse the
         * buffer we have
         */
        services_returned = 0;
        success = EnumServicesStatusExW(
            hService,
            infolevel,
            dwServiceType,
            dwServiceState,
            (LPBYTE) sbuf,
            buf_sz,
            &buf_needed,
            &services_returned,
            &resume_handle,
            ObjToLPWSTR_WITH_NULL(groupnameObj));
        if ((!success) && ((winerr = GetLastError()) != ERROR_MORE_DATA)) {
            Twapi_FreeNewTclObj(resultObj);
            Twapi_AppendSystemError(interp, winerr);
            goto pop_and_vamoose;
        }

        /* Tack on the services returned */
        for (i = 0; i < services_returned; ++i) {
            DWORD dw;

            /* Note order should be same as order of field names above */

            dw = sbuf[i].ServiceStatusProcess.dwServiceType  &  ~SERVICE_INTERACTIVE_PROCESS;
            switch (dw) {
            case 0x1:  rec[0] = service_types[0]; break;
            case 0x2:  rec[0] = service_types[1]; break;
            case 0x4:  rec[0] = service_types[2]; break;
            case 0x8:  rec[0] = service_types[3]; break;
            case 0x10:  rec[0] = service_types[4]; break;
            case 0x20:  rec[0] = service_types[5]; break;
            default: rec[0] = ObjFromDWORD(dw); break;
            }

            dw = sbuf[i].ServiceStatusProcess.dwCurrentState;
            if (dw < ARRAYSIZE(states))
                rec[1] = states[dw];
            else
                rec[1] = ObjFromDWORD(dw);
            rec[2] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwControlsAccepted);
            rec[3] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwWin32ExitCode);
            rec[4] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwServiceSpecificExitCode);
            rec[5] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwCheckPoint);
            rec[6] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwWaitHint);
            rec[7] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwProcessId);
            rec[8] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwServiceFlags);
            rec[9] = ObjFromUnicode(sbuf[i].lpServiceName); /* KEY for record array */
            rec[10] = ObjFromUnicode(sbuf[i].lpDisplayName);
            rec[11] = Tcl_NewBooleanObj(sbuf[i].ServiceStatusProcess.dwServiceType & SERVICE_INTERACTIVE_PROCESS);

            /* Note rec[9] object is also appended as the "key" for the
               but in lower case form */
            Tcl_DictObjPut(interp, resultObj, TwapiLowerCaseObj(rec[9]),
                           TwapiTwineObjv(keys, rec, ARRAYSIZE(rec)));
        }
        /* If !success -> ERROR_MORE_DATA so keep looping */
    } while (! success);

    ObjSetResult(interp, resultObj);
    status = TCL_OK;

pop_and_vamoose:

    MemLifoPopFrame(&ticP->memlifo);
    return status;
}

int Twapi_EnumDependentServices(
    TwapiInterpContext *ticP,
    SC_HANDLE hService,
    DWORD     dwServiceState
    )
{
    Tcl_Interp *interp = ticP->interp;
    ENUM_SERVICE_STATUSW *sbuf;
    DWORD buf_sz;
    DWORD services_returned;
    BOOL  success;
    Tcl_Obj *rec[10];    /* Holds values for each status record */
    Tcl_Obj *keys[ARRAYSIZE(rec)];
    Tcl_Obj *states[8];
    DWORD winerr;
    DWORD i;
    TCL_RESULT status = TCL_ERROR;
    Tcl_Obj *service_types[6];
    Tcl_Obj *resultObj;


    sbuf = MemLifoPushFrame(&ticP->memlifo, 4000, &buf_sz);

    /* NOTE - the returned Tcl_Obj from TwapiGetAtom must not be freed */

    /* Create the map of service types codes to strings */
    service_types[0] = TwapiGetAtom(ticP, ServiceTypeString(0x1));
    service_types[1] = TwapiGetAtom(ticP, ServiceTypeString(0x2));
    service_types[2] = TwapiGetAtom(ticP, ServiceTypeString(0x4));
    service_types[3] = TwapiGetAtom(ticP, ServiceTypeString(0x8));
    service_types[4] = TwapiGetAtom(ticP, ServiceTypeString(0x10));
    service_types[5] = TwapiGetAtom(ticP, ServiceTypeString(0x20));

    /* And the state symbols ... */
    for (i=0; i < ARRAYSIZE(states); ++i) {
        states[i] = TwapiGetAtom(ticP, ServiceStateString(i));
    }

    do {
        success = EnumDependentServicesW(hService,
                                         dwServiceState,
                                         sbuf,
                                         buf_sz,
                                         &buf_sz,
                                         &services_returned);
        if (success)
            break;
        winerr = GetLastError();
        if (winerr != ERROR_MORE_DATA)  {
            Twapi_AppendSystemError(interp, winerr);
            goto pop_and_vamoose;
        }

        /* Need a bigger buffer */
        MemLifoPopFrame(&ticP->memlifo);
        sbuf = MemLifoPushFrame(&ticP->memlifo, buf_sz, NULL);
    } while (1);

    /* Note order of names should be same as order of values below */
    keys[0] = TwapiGetAtom(ticP, "servicetype");
    keys[1] = TwapiGetAtom(ticP, "state");
    keys[2] = TwapiGetAtom(ticP, "control_accepted");
    keys[3] = TwapiGetAtom(ticP, "exitcode");
    keys[4] = TwapiGetAtom(ticP, "service_code");
    keys[5] = TwapiGetAtom(ticP, "checkpoint");
    keys[6] = TwapiGetAtom(ticP, "wait_hint");
    keys[7] = TwapiGetAtom(ticP, "name");
    keys[8] = TwapiGetAtom(ticP, "displayname");
    keys[9] = TwapiGetAtom(ticP, "interactive");

    resultObj = Tcl_NewDictObj();
    for (i = 0; i < services_returned; ++i) {
        DWORD dw;

        /* Note order should be same as order of field names below */
        dw = sbuf[i].ServiceStatus.dwServiceType  &  ~SERVICE_INTERACTIVE_PROCESS;
        switch (dw) {
        case 0x1:  rec[0] = service_types[0]; break;
        case 0x2:  rec[0] = service_types[1]; break;
        case 0x4:  rec[0] = service_types[2]; break;
        case 0x8:  rec[0] = service_types[3]; break;
        case 0x10:  rec[0] = service_types[4]; break;
        case 0x20:  rec[0] = service_types[5]; break;
        default: rec[0] = ObjFromDWORD(dw); break;
        }

        dw = sbuf[i].ServiceStatus.dwCurrentState;
        if (dw < ARRAYSIZE(states))
            rec[1] = states[dw];
        else
            rec[1] = ObjFromDWORD(dw);
        rec[2] = ObjFromDWORD(sbuf[i].ServiceStatus.dwControlsAccepted);
        rec[3] = ObjFromDWORD(sbuf[i].ServiceStatus.dwWin32ExitCode);
        rec[4] = ObjFromDWORD(sbuf[i].ServiceStatus.dwServiceSpecificExitCode);
        rec[5] = ObjFromDWORD(sbuf[i].ServiceStatus.dwCheckPoint);
        rec[6] = ObjFromDWORD(sbuf[i].ServiceStatus.dwWaitHint);
        rec[7] = ObjFromUnicode(sbuf[i].lpServiceName); /* KEY for record array */
        rec[8] = ObjFromUnicode(sbuf[i].lpDisplayName);
        rec[9] = ObjFromDWORD(sbuf[i].ServiceStatus.dwServiceType & SERVICE_INTERACTIVE_PROCESS);

        /* Note rec[7] object is also appended as the "key" but in lowercase*/
        Tcl_DictObjPut(interp, resultObj, TwapiLowerCaseObj(rec[7]), TwapiTwineObjv(keys, rec, ARRAYSIZE(rec)));
    }

    ObjSetResult(interp, resultObj);
    status = TCL_OK;

pop_and_vamoose:
    MemLifoPopFrame(&ticP->memlifo);
    return status;
}


int Twapi_ChangeServiceConfig(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[]) {
    SC_HANDLE h;
    DWORD service_type;
    DWORD start_type;
    DWORD error_control;
    DWORD tag_id;
    DWORD *tag_idP;
    LPWSTR dependencies = NULL;
    TCL_RESULT res;
    LPWSTR path, logrp, start_name, password, display_name;
    int password_len;
    Tcl_Interp *interp = ticP->interp;
    MemLifoMarkHandle mark;
    Tcl_Obj *tagObj, *depObj, *passwordObj;

    mark = MemLifoPushMark(&ticP->memlifo);

    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETHANDLET(h, SC_HANDLE),
                         GETINT(service_type), GETINT(start_type),
                         GETINT(error_control), GETTOKENNULL(path),
                         GETTOKENNULL(logrp), GETOBJ(tagObj), GETOBJ(depObj),
                         GETTOKENNULL(start_name),
                         GETOBJ(passwordObj), GETTOKENNULL(display_name),
                         ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    if (ObjToLong(interp, tagObj, &tag_id) == TCL_OK)
        tag_idP = &tag_id;
    else {
        /* An empty string means value is not to be changed. Else error */
        if (Tcl_GetCharLength(tagObj) != 0) {
            res = TCL_ERROR;
            goto vamoose;   /* interp already holds error fom ObjToLong */
        }
        Tcl_ResetResult(interp);
        tag_idP = NULL;         /* Tag is not to be changed */
    }

    dependencies = ObjToUnicode(depObj);
    if (lstrcmpW(dependencies, NULL_TOKEN_L) == 0) {
        dependencies = NULL;
    } else {
        res = ObjToMultiSzEx(interp, depObj, (LPCWSTR*) &dependencies, &ticP->memlifo);
        if (res != TCL_OK)
            goto vamoose;
    }
    
    password = ObjDecryptPassword(passwordObj, &password_len);
    if (ChangeServiceConfigW(h, service_type, start_type, error_control,
                             path, logrp, tag_idP, dependencies,
                             start_name,
                             lstrcmpW(password, NULL_TOKEN_L) ? password : NULL,
                             display_name)) {
        /* If caller wants new tag id returned (by not specifying
         * an empty tagid, return it, else return empty string.
         */
        if (tag_idP)
            ObjSetResult(interp, Tcl_NewLongObj(*tag_idP));
        res = TCL_OK;
    } else
        res = TwapiReturnSystemError(interp);

    TwapiFreeDecryptedPassword(password, password_len);

vamoose:    
    MemLifoPopMark(mark);
    return res;
}

int
Twapi_CreateService(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[]) {
    SC_HANDLE    scmH;
    SC_HANDLE    svcH;
    DWORD        access;
    DWORD        service_type;
    DWORD        start_type;
    DWORD        error_control;
    DWORD tag_id;
    DWORD *tag_idP;
    LPWSTR dependencies;
    int res, password_len;
    LPWSTR service_name, display_name, path, logrp;
    LPWSTR service_start_name, password;
    Tcl_Interp *interp = ticP->interp;
    MemLifoMarkHandle mark;
    Tcl_Obj *tagObj, *depObj, *passwordObj;

    dependencies = NULL;
    mark = MemLifoPushMark(&ticP->memlifo);

    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETHANDLET(scmH, SC_HANDLE),
                         GETSTRW(service_name), GETSTRW(display_name),
                         GETINT(access), GETINT(service_type),
                         GETINT(start_type), GETINT(error_control),
                         GETSTRW(path), GETSTRW(logrp),
                         GETOBJ(tagObj), GETOBJ(depObj),
                         GETEMPTYASNULL(service_start_name),
                         GETOBJ(passwordObj),
                         ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    if (ObjToLong(NULL, tagObj, &tag_id) == TCL_OK)
        tag_idP = &tag_id;
    else {
        /* An empty string means value is not to be changed. Else error */
        if (Tcl_GetCharLength(tagObj) != 0)
            return TCL_ERROR;   /* interp already holds error */
        tag_idP = NULL;         /* Tag is not to be changed */
    }

    dependencies = ObjToUnicode(depObj);
    if (lstrcmpW(dependencies, NULL_TOKEN_L) == 0) {
        dependencies = NULL;
    } else {
        res = ObjToMultiSzEx(interp, depObj, (LPCWSTR*) &dependencies, &ticP->memlifo);
        if (res == TCL_ERROR)
            goto vamoose;
    }

    password = ObjDecryptPassword(passwordObj, &password_len);

    svcH = CreateServiceW(
        scmH, service_name, display_name, access, service_type,
        start_type, error_control, path, logrp,
        tag_idP, dependencies, service_start_name,
        password[0] ? password : NULL);

    TwapiFreeDecryptedPassword(password, password_len);

    /* Check return handle validity */
    if (svcH) {
        ObjSetResult(interp, ObjFromOpaque(svcH, "SC_HANDLE"));
        res = TCL_OK;
    } else
        res = TwapiReturnSystemError(interp);

vamoose:
    MemLifoPopMark(mark);
    return res;
}

int Twapi_StartService(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]) 
{
    SC_HANDLE svcH;
    LPCWSTR args[64];
    Tcl_Obj **argObjs;
    int     i, nargs;
    
    if (objc != 2) {
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    }

    if (ObjToHANDLE(interp, objv[0], &svcH) != TCL_OK)
        return TCL_ERROR;

    if (ObjGetElements(interp, objv[1], &nargs, &argObjs) == TCL_ERROR)
        return TCL_ERROR;

    if (nargs > ARRAYSIZE(args)) {
        ObjSetStaticResult(interp, "Exceeded limit on number of service arguments.");
        return TCL_ERROR;
    }

    for (i = 0; i < nargs; i++) {
        args[i] = ObjToUnicode(argObjs[i]);
    }
    if (StartServiceW(svcH, nargs, args))
        return TCL_OK;
    else
        return TwapiReturnSystemError(interp);
}

#ifdef NOOP_BEYOND_VISTA
/*
 * SC_LOCK is a opaque pointer that is a handle to a lock on the SCM database
 * We treat it as a pointer except that on return value, if it is NULL,
 * but the error is something other than database already being locked,
 * we generate a TCL exception. For database locked error, we will
 * return the string "NULL"
 */

SC_LOCK LockServiceDatabase(SC_HANDLE   hSCManager);
EXCEPTION_ON_FALSE UnlockServiceDatabase(SC_LOCK ScLock);
int Twapi_QueryServiceLockStatus(
    Tcl_Interp *interp,
    SC_HANDLE hSCManager
    );
#endif // NOOP_BEYOND_VISTA


#ifndef TWAPI_SINGLE_MODULE
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif


static int Twapi_ServiceCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    union {
        SERVICE_STATUS svcstatus;
        WCHAR buf[MAX_PATH+1];
    } u;
    DWORD dw;
    LPWSTR s;
    HANDLE h;
    Tcl_Obj *sObj;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 100) {
        /* Expect a single handle argument */
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToHANDLE(interp, objv[2], &h) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 1:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeleteService(h);
            break;
        case 2:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseServiceHandle(h);
            break;
        case 3:
            return Twapi_QueryServiceConfig(ticP, h);
        }
    } else if (func < 200) {
        /* Expect a handle and a int */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 101:
            result.type = TRT_EXCEPTION_ON_FALSE;
            /* svcstatus is not returned because it is not always filled
               in and is not very useful even when it is */
            result.value.ival = ControlService(h, dw, &u.svcstatus);
            break;
        case 102:
            return Twapi_EnumDependentServices(ticP, h, dw);
        case 103:
            return Twapi_QueryServiceStatusEx(interp, h, dw);
        case 104:
            return Twapi_QueryServiceConfig2(ticP, h, dw);
        }
    } else if (func < 300) {
        /* Handle, string, int */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETOBJ(sObj), ARGUSEDEFAULT,
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        s = ObjToUnicode(sObj);
        switch (func) {
        case 201:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetServiceKeyNameW(h, s, u.buf, &result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 202:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetServiceDisplayNameW(h, s, u.buf, &result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 203:
            /* If access type not specified, use SERVICE_ALL_ACCESS */
            if (objc < 5)
                dw = SERVICE_ALL_ACCESS;
            result.type = TRT_SC_HANDLE;
            result.value.hval = OpenServiceW(h, s, dw);
            break;
        }
    } else {
        /* Free for all */
        switch (func) {
        case 10001:
            return Twapi_EnumServicesStatusEx(ticP, objc-2, objv+2);
        case 10002:
            return Twapi_ChangeServiceConfig(ticP, objc-2, objv+2);
        case 10003:
            return Twapi_CreateService(ticP, objc-2, objv+2);
        case 10004:
            return Twapi_StartService(interp, objc-2, objv+2);
        case 10005:
            return Twapi_SetServiceStatus(ticP, objc-2, objv+2);
        case 10006:
            return Twapi_BecomeAService(ticP, objc-2, objv+2);
        case 10007:
            CHECK_NARGS(interp, objc, 5);
            CHECK_INTEGER_OBJ(interp, dw, objv[4]);
            result.type = TRT_SC_HANDLE;
            result.value.hval = OpenSCManagerW(
                ObjToLPWSTR_NULL_IF_EMPTY(objv[2]),
                ObjToLPWSTR_NULL_IF_EMPTY(objv[3]),
                dw);
            break;
        }
    }
    
    return TwapiSetResult(interp, &result);
}

static int TwapiServiceInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct alias_dispatch_s ServiceDispatch[] = {
        DEFINE_ALIAS_CMD(DeleteService, 1),
        DEFINE_ALIAS_CMD(CloseServiceHandle, 2),
        DEFINE_ALIAS_CMD(QueryServiceConfig, 3),

        DEFINE_ALIAS_CMD(ControlService, 101),
        DEFINE_ALIAS_CMD(EnumDependentServices, 102),
        DEFINE_ALIAS_CMD(QueryServiceStatusEx, 103),
        DEFINE_ALIAS_CMD(QueryServiceConfig2, 104),

        DEFINE_ALIAS_CMD(GetServiceKeyName, 201),
        DEFINE_ALIAS_CMD(GetServiceDisplayName, 202),
        DEFINE_ALIAS_CMD(OpenService, 203),

        DEFINE_ALIAS_CMD(EnumServicesStatusEx, 10001),
        DEFINE_ALIAS_CMD(ChangeServiceConfig, 10002),
        DEFINE_ALIAS_CMD(CreateService, 10003),
        DEFINE_ALIAS_CMD(StartService, 10004),
        DEFINE_ALIAS_CMD(Twapi_SetServiceStatus, 10005),
        DEFINE_ALIAS_CMD(Twapi_BecomeAService, 10006),
        DEFINE_ALIAS_CMD(OpenSCManager, 10007),
    };

    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::SvcCall", Twapi_ServiceCallObjCmd, ticP, NULL);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(ServiceDispatch), ServiceDispatch, "twapi::SvcCall");

    return TCL_OK;
}



/* Main entry point */
#ifndef TWAPI_SINGLE_MODULE
__declspec(dllexport) 
#endif
int Twapi_service_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiServiceInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

