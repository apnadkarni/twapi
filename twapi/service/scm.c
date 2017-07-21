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

#ifndef MODULENAME
#define MODULENAME "twapi_service"
#endif

TWAPI_STATIC_INLINE TCL_RESULT ObjToSC_HANDLE(Tcl_Interp *interp, Tcl_Obj *objP, SC_HANDLE *schP) {
    HANDLE sch;
    TCL_RESULT res;
    /* Use of the intermediary is to keep gcc happy */
    res = ObjToHANDLE(interp, objP, &sch);
    if (res == TCL_OK)
        *schP = sch;
    return res;
}

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

static Tcl_Obj *ServiceStateAtom(TwapiInterpContext *ticP, DWORD state)
{
    return TwapiGetAtom(ticP, ServiceStateString(state));
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
    case 0x30: return "win32";
    case 0x50: return "user_own_process";
    case 0x60: return "user_share_process";
    default: return "unknown";
    }
}

static Tcl_Obj *ServiceTypeAtom(TwapiInterpContext *ticP, DWORD service_type)
{
    return TwapiGetAtom(ticP, ServiceTypeString(service_type));
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
    rec[1] = ObjFromString(ServiceTypeString(ssp.dwServiceType));
    rec[2] = STRING_LITERAL_OBJ("state");
    cP = ServiceStateString(ssp.dwCurrentState);
    rec[3] = cP ? ObjFromString(cP) : ObjFromDWORD(ssp.dwCurrentState);
    rec[4] = STRING_LITERAL_OBJ("controls_accepted");
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

    /* Ask for 1000 bytes alloc, will get more if available */
    qbuf = (QUERY_SERVICE_CONFIGW *) MemLifoPushFrame(ticP->memlifoP,
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
        qbuf = (QUERY_SERVICE_CONFIGW *) MemLifoAlloc(ticP->memlifoP, buf_sz, NULL);

        /* Get the configuration information.  */
        if (! QueryServiceConfigW(hService, qbuf, buf_sz, &buf_sz)) {
            winerr = GetLastError();
            goto vamoose;
        }
    }

    objv[0] = STRING_LITERAL_OBJ("-dependencies");
    objv[1] = ObjFromMultiSz_MAX(qbuf->lpDependencies);
    objv[2] = STRING_LITERAL_OBJ("-servicetype");
    objv[3] = ServiceTypeAtom(ticP, qbuf->dwServiceType);
    objv[4] = STRING_LITERAL_OBJ("-starttype");
    objv[5] = ObjFromDWORD(qbuf->dwStartType);
    objv[6] = STRING_LITERAL_OBJ("-errorcontrol");
    objv[7] = ObjFromDWORD(qbuf->dwErrorControl);
    objv[8] = STRING_LITERAL_OBJ("-tagid");
    objv[9] = ObjFromDWORD(qbuf->dwTagId);
    objv[10] = STRING_LITERAL_OBJ("-command");
    objv[11] = ObjFromWinChars(qbuf->lpBinaryPathName);
    objv[12] = STRING_LITERAL_OBJ("-loadordergroup");
    objv[13] = ObjFromWinChars(qbuf->lpLoadOrderGroup);
    objv[14] = STRING_LITERAL_OBJ("-account");
    objv[15] = ObjFromWinChars(qbuf->lpServiceStartName);
    objv[16] = STRING_LITERAL_OBJ("-displayname");
    objv[17] = ObjFromWinChars(qbuf->lpDisplayName);
    objv[18] = STRING_LITERAL_OBJ("-interactive");
    objv[19] = ObjFromBoolean(qbuf->dwServiceType & SERVICE_INTERACTIVE_PROCESS);
    ObjSetResult(ticP->interp, ObjNewList(20,objv));
    tcl_result = TCL_OK;

vamoose:
    MemLifoPopFrame(ticP->memlifoP);
    if (tcl_result != TCL_OK)
        Twapi_AppendSystemError(ticP->interp, winerr);

    return tcl_result;
}


Tcl_Obj *ObjFromSERVICE_FAILURE_ACTIONS(SERVICE_FAILURE_ACTIONSW *sfaP)
{
    Tcl_Obj *objs[4];
    DWORD i;

    objs[0] = ObjFromLong(sfaP->dwResetPeriod);
    objs[1] = ObjFromWinChars(sfaP->lpRebootMsg);
    objs[2] = ObjFromWinChars(sfaP->lpCommand);
    objs[3] = ObjNewList(sfaP->cActions, NULL);
    if (sfaP->lpsaActions) {
        for (i = 0; i < sfaP->cActions; ++i) {
            Tcl_Obj *fields[2];
            fields[0] = ObjFromInt(sfaP->lpsaActions[i].Type);
            fields[1] = ObjFromDWORD(sfaP->lpsaActions[i].Delay);
            ObjAppendElement(NULL, objs[3], ObjNewList(2, fields));
        }
    }

    return ObjNewList(4, objs);
}


int Twapi_QueryServiceConfig2(TwapiInterpContext *ticP, SC_HANDLE hService, DWORD level)
{
    void *bufP;
    DWORD buf_sz;
    TCL_RESULT res = TCL_ERROR;
    

    /* Max size of buffer required is 8K as per MSDN */
    buf_sz = 8 * 1024;
    bufP = MemLifoPushFrame(ticP->memlifoP, buf_sz, &buf_sz);
    if (QueryServiceConfig2W(hService, level, (LPBYTE) bufP, buf_sz, &buf_sz)) {
        switch (level) {
        case SERVICE_CONFIG_DESCRIPTION:
            /* If NULL, we keep result as empty string. Not an error */
            res = TCL_OK;
            if (((SERVICE_DESCRIPTIONW *)bufP)->lpDescription)
                ObjSetResult(ticP->interp, ObjFromWinChars(((SERVICE_DESCRIPTIONW *)bufP)->lpDescription));
            break;

        case SERVICE_CONFIG_FAILURE_ACTIONS:
            ObjSetResult(ticP->interp, ObjFromSERVICE_FAILURE_ACTIONS(bufP));
            res = TCL_OK;
            break;

        default:
            res = TwapiReturnError(ticP->interp, TWAPI_INVALID_OPTION);
            break;
        }
    } else {
        /* Failure */
        res = TwapiReturnSystemError(ticP->interp);
    }

    MemLifoPopFrame(ticP->memlifoP);
    return res;
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
    sbuf = MemLifoPushFrame(ticP->memlifoP, buf_sz, NULL);
    /* Get the configuration information.  */
    if (! QueryServiceLockStatusW(hService, sbuf, buf_sz, &buf_sz)) {
        TwapiReturnSystemError(interp); // Store before call to free
        MemLifoPopFrame(ticP->memlifoP);
        return TCL_ERROR;
    }

    obj = ObjEmptyList();
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, fIsLocked);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, obj, sbuf, lpLockOwner);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, dwLockDuration);

    MemLifoPopFrame(ticP->memlifoP);

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
    DWORD winerr;
    DWORD i;
    TCL_RESULT status = TCL_ERROR;

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
    sbuf = MemLifoPushFrame(ticP->memlifoP, 32000, &buf_sz);

    resultObj = ObjNewList(200, NULL);
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

            /* Note order should be same as order of field names above */

            rec[0] = ServiceTypeAtom(ticP, sbuf[i].ServiceStatusProcess.dwServiceType);

            rec[1] = ServiceStateAtom(ticP, sbuf[i].ServiceStatusProcess.dwCurrentState);
            rec[2] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwControlsAccepted);
            rec[3] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwWin32ExitCode);
            rec[4] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwServiceSpecificExitCode);
            rec[5] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwCheckPoint);
            rec[6] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwWaitHint);
            rec[7] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwProcessId);
            rec[8] = ObjFromDWORD(sbuf[i].ServiceStatusProcess.dwServiceFlags);
            rec[9] = ObjFromWinChars(sbuf[i].lpServiceName); /* KEY for record array */
            rec[10] = ObjFromWinChars(sbuf[i].lpDisplayName);
            rec[11] = Tcl_NewBooleanObj(sbuf[i].ServiceStatusProcess.dwServiceType & SERVICE_INTERACTIVE_PROCESS);

            ObjAppendElement(NULL, resultObj, ObjNewList(ARRAYSIZE(rec), rec));
        }
        /* If !success -> ERROR_MORE_DATA so keep looping */
    } while (! success);

    ObjSetResult(interp, resultObj);
    status = TCL_OK;

pop_and_vamoose:

    MemLifoPopFrame(ticP->memlifoP);
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
    DWORD winerr;
    DWORD i;
    TCL_RESULT status = TCL_ERROR;
    Tcl_Obj *resultObj;


    sbuf = MemLifoPushFrame(ticP->memlifoP, 4000, &buf_sz);

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
        MemLifoPopFrame(ticP->memlifoP);
        sbuf = MemLifoPushFrame(ticP->memlifoP, buf_sz, NULL);
    } while (1);

    resultObj = ObjNewList(10, NULL);
    for (i = 0; i < services_returned; ++i) {

        /* Note order should be same as order of field names below */

        rec[0] = ServiceTypeAtom(ticP, sbuf[i].ServiceStatus.dwServiceType);
        rec[1] = ServiceStateAtom(ticP, sbuf[i].ServiceStatus.dwCurrentState);
        rec[2] = ObjFromDWORD(sbuf[i].ServiceStatus.dwControlsAccepted);
        rec[3] = ObjFromDWORD(sbuf[i].ServiceStatus.dwWin32ExitCode);
        rec[4] = ObjFromDWORD(sbuf[i].ServiceStatus.dwServiceSpecificExitCode);
        rec[5] = ObjFromDWORD(sbuf[i].ServiceStatus.dwCheckPoint);
        rec[6] = ObjFromDWORD(sbuf[i].ServiceStatus.dwWaitHint);
        rec[7] = ObjFromWinChars(sbuf[i].lpServiceName); /* KEY for record array */
        rec[8] = ObjFromWinChars(sbuf[i].lpDisplayName);
        rec[9] = ObjFromDWORD(sbuf[i].ServiceStatus.dwServiceType & SERVICE_INTERACTIVE_PROCESS);

        ObjAppendElement(NULL, resultObj, ObjNewList(ARRAYSIZE(rec), rec));
    }

    ObjSetResult(interp, resultObj);
    status = TCL_OK;

pop_and_vamoose:
    MemLifoPopFrame(ticP->memlifoP);
    return status;
}


/* Returns SERVICE_FAILURE_ACTIONS structure in *sfaP
   using memory from ticP->memlifo. Caller responsible for storage
   in both success and error cases
*/
static TCL_RESULT ParseSERVICE_FAILURE_ACTIONS(
    TwapiInterpContext *ticP,
    Tcl_Obj *sfaObj,
    SERVICE_FAILURE_ACTIONSW *sfaP
    )
{
    Tcl_Interp *interp = ticP->interp;
    Tcl_Obj **objs;
    Tcl_Obj *actionsObj;
    int i, nobjs;
    TCL_RESULT res;

    res = ObjGetElements(interp, sfaObj, &nobjs, &objs);
    if (res != TCL_OK)
        return res;
    res = TwapiGetArgsEx(ticP, nobjs, objs,
                         GETINT(sfaP->dwResetPeriod),
                         GETTOKENNULL(sfaP->lpRebootMsg),
                         GETTOKENNULL(sfaP->lpCommand),
                         ARGUSEDEFAULT,
                         GETOBJ(actionsObj), ARGEND);
    if (res != TCL_OK)
        return res;

    if (actionsObj == NULL) {
        /* No actions specified, so will be left unchanged */
        sfaP->cActions = 0;
        sfaP->lpsaActions = NULL;
        return TCL_OK;
    }

    res = ObjGetElements(interp, actionsObj, &nobjs, &objs);
    if (res != TCL_OK)
        return res;

    sfaP->cActions = nobjs;
    sfaP->lpsaActions = MemLifoAlloc(ticP->memlifoP,
                                     (nobjs ? nobjs : 1) * sizeof(SC_ACTION),
                                     NULL);

    /* Special case - to delete sfaP->lpsaActions, set sfaP->cActions to 0
       and sfaP->lpsaActions to non-NULL
    */
    if (nobjs == 0)
        return TCL_OK;

    for (i = 0; i < nobjs; ++i) {
        Tcl_Obj **fields;
        int nfields;
        int sc_type;

        res = ObjGetElements(interp, objs[i], &nfields, &fields);
        if (res != TCL_OK)
            return res;
        if (nfields != 2)
            return TwapiReturnError(interp, TWAPI_INVALID_DATA);
        if (ObjToInt(interp, fields[0], &sc_type) != TCL_OK ||
            ObjToDWORD(interp, fields[1], &sfaP->lpsaActions[i].Delay) != TCL_OK)
            return TCL_ERROR;
        sfaP->lpsaActions[i].Type = (SC_ACTION_TYPE) sc_type;
    }        

    return TCL_OK;
}


static TCL_RESULT Twapi_ChangeServiceConfig2(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    TCL_RESULT res;
    MemLifoMarkHandle mark;
    DWORD info_level;
    Tcl_Obj *infoObj;
    union {
        SERVICE_DESCRIPTIONW desc;
        SERVICE_FAILURE_ACTIONSW failure_actions;
    } u;
    SC_HANDLE h;
    void *pv;

    mark = MemLifoPushMark(ticP->memlifoP);
    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETHANDLET(h, SC_HANDLE), GETINT(info_level),
                         GETOBJ(infoObj), ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    switch (info_level) {
    case SERVICE_CONFIG_DESCRIPTION:
        pv = &u.desc;
        u.desc.lpDescription = ObjToWinChars(infoObj);
        break;
    case SERVICE_CONFIG_FAILURE_ACTIONS:
        pv = &u.failure_actions;
        res = ParseSERVICE_FAILURE_ACTIONS(ticP, infoObj, &u.failure_actions);
        break;
    default:
        res = TwapiReturnError(ticP->interp, TWAPI_INVALID_OPTION);
        break;
    }

    if (res == TCL_OK) {
        if (!ChangeServiceConfig2W(h, info_level, pv))
            res = TwapiReturnSystemError(ticP->interp);
    }

vamoose:    
    MemLifoPopMark(mark);
    return res;
}

int Twapi_ChangeServiceConfig(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
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

    mark = MemLifoPushMark(ticP->memlifoP);

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

    if (ObjToDWORD(interp, tagObj, &tag_id) == TCL_OK)
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

    dependencies = ObjToWinChars(depObj);
    if (lstrcmpW(dependencies, NULL_TOKEN_L) == 0) {
        dependencies = NULL;
    } else {
        res = ObjToMultiSzEx(interp, depObj, (LPCWSTR*) &dependencies, ticP->memlifoP);
        if (res != TCL_OK)
            goto vamoose;
    }
    
    TWAPI_ASSERT(ticP->memlifoP == SWS());
    password = ObjDecryptPasswordSWS(passwordObj, &password_len);
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

    SecureZeroMemory(password, password_len);

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
    mark = MemLifoPushMark(ticP->memlifoP);

    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETHANDLET(scmH, SC_HANDLE),
                         GETWSTR(service_name), GETWSTR(display_name),
                         GETINT(access), GETINT(service_type),
                         GETINT(start_type), GETINT(error_control),
                         GETWSTR(path), GETWSTR(logrp),
                         GETOBJ(tagObj), GETOBJ(depObj),
                         GETEMPTYASNULL(service_start_name),
                         GETOBJ(passwordObj),
                         ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    if (ObjToDWORD(NULL, tagObj, &tag_id) == TCL_OK)
        tag_idP = &tag_id;
    else {
        /* An empty string means value is not to be changed. Else error */
        if (Tcl_GetCharLength(tagObj) != 0)
            return TCL_ERROR;   /* interp already holds error */
        tag_idP = NULL;         /* Tag is not to be changed */
    }

    dependencies = ObjToWinChars(depObj);
    if (lstrcmpW(dependencies, NULL_TOKEN_L) == 0) {
        dependencies = NULL;
    } else {
        res = ObjToMultiSzEx(interp, depObj, (LPCWSTR*) &dependencies, ticP->memlifoP);
        if (res == TCL_ERROR)
            goto vamoose;
    }

    TWAPI_ASSERT(ticP->memlifoP == SWS());
    password = ObjDecryptPasswordSWS(passwordObj, &password_len);

    svcH = CreateServiceW(
        scmH, service_name, display_name, access, service_type,
        start_type, error_control, path, logrp,
        tag_idP, dependencies, service_start_name,
        password[0] ? password : NULL);

    SecureZeroMemory(password, password_len);

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

    if (ObjToSC_HANDLE(interp, objv[0], &svcH) != TCL_OK)
        return TCL_ERROR;

    if (ObjGetElements(interp, objv[1], &nargs, &argObjs) == TCL_ERROR)
        return TCL_ERROR;

    if (nargs > ARRAYSIZE(args)) {
        ObjSetStaticResult(interp, "Exceeded limit on number of service arguments.");
        return TCL_ERROR;
    }

    for (i = 0; i < nargs; i++) {
        args[i] = ObjToWinChars(argObjs[i]);
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


static int Twapi_ServiceCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = clientdata;
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
        s = ObjToWinChars(sObj);
        switch (func) {
        case 201:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetServiceKeyNameW(h, s, u.buf, (DWORD *)&result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 202:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetServiceDisplayNameW(h, s, u.buf, (DWORD *)&result.value.unicode.len)) {
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
            CHECK_DWORD_OBJ(interp, dw, objv[4]);
            result.type = TRT_SC_HANDLE;
            result.value.hval = OpenSCManagerW(
                ObjToLPWSTR_NULL_IF_EMPTY(objv[2]),
                ObjToLPWSTR_NULL_IF_EMPTY(objv[3]),
                dw);
            break;
        case 10008:
            return Twapi_ChangeServiceConfig2(ticP, objc-2, objv+2);
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
        DEFINE_ALIAS_CMD(ChangeServiceConfig2, 10008),
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

