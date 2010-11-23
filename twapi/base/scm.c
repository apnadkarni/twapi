/*
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows NT services */
#include "twapi.h"

/* Map service state int to string */
static Tcl_Obj *ObjFromServiceState(DWORD state)
{
    switch (state) {
    case 1: return STRING_LITERAL_OBJ("stopped");
    case 2: return STRING_LITERAL_OBJ("start_pending");
    case 3: return STRING_LITERAL_OBJ("stop_pending");
    case 4: return STRING_LITERAL_OBJ("running");
    case 5: return STRING_LITERAL_OBJ("continue_pending");
    case 6: return STRING_LITERAL_OBJ("pause_pending");
    case 7: return STRING_LITERAL_OBJ("paused");
    }

    return Tcl_NewLongObj(state);
}

static Tcl_Obj *ObjFromServiceType(DWORD service_type)
{
    service_type &=  ~SERVICE_INTERACTIVE_PROCESS;
    switch (service_type) {
    case 0x1: return STRING_LITERAL_OBJ("kernel_driver");
    case 0x2: return STRING_LITERAL_OBJ("file_system_driver");
    case 0x4: return STRING_LITERAL_OBJ("adapter");
    case 0x8: return STRING_LITERAL_OBJ("recognizer_driver");
    case 0x10: return STRING_LITERAL_OBJ("win32_own_process");
    case 0x20: return STRING_LITERAL_OBJ("win32_share_process");
    }

    return Tcl_NewLongObj(service_type);
}

/*
 * We want to share service type values when returning status for
 * multiple services. We accomplish this by populating a service
 * type to Tcl_Obj map and reusing these via ObjFromServiceTypeObjs
 * until FreeServiceTypeObjs is called.
 *
 * objPP[] must be size 6
 */
static void PopulateServiceTypeObjs(Tcl_Obj **objPP) 
{
    int i;
    objPP[0] = ObjFromServiceType(0x1);
    objPP[1] = ObjFromServiceType(0x2);
    objPP[2] = ObjFromServiceType(0x4);
    objPP[3] = ObjFromServiceType(0x8);
    objPP[4] = ObjFromServiceType(0x10);
    objPP[5] = ObjFromServiceType(0x20);
    for (i = 0; i < 6; ++i) {
        Tcl_IncrRefCount(objPP[i]);
    }
}

/* See PopulateServiceTypeObjs */
static void FreeServiceTypeObjs(Tcl_Obj **objPP) 
{
    int i;
    for (i = 0; i < 6; ++i) {
        Tcl_DecrRefCount(objPP[i]);
    }
}

/* See PopulateServiceTypeObjs */
static Tcl_Obj *ObjFromServiceTypeObjs(DWORD service_type, Tcl_Obj **objPP)
{
    service_type &=  ~SERVICE_INTERACTIVE_PROCESS;
    switch (service_type) {
    case 0x1: return objPP[0];
    case 0x2: return objPP[1];
    case 0x4: return objPP[2];
    case 0x8: return objPP[3];
    case 0x10: return objPP[4];
    case 0x20: return objPP[5];
    }

    return Tcl_NewLongObj(service_type);
}


int Twapi_QueryServiceStatusEx(Tcl_Interp *interp, SC_HANDLE h,
        SC_STATUS_TYPE infolevel)
{
    SERVICE_STATUS_PROCESS ssp;
    DWORD numbytes;
    DWORD result;
    Tcl_Obj *rec[20];

    if (infolevel != SC_STATUS_PROCESS_INFO) {
        return Twapi_AppendSystemError(interp, ERROR_INVALID_LEVEL);
    }

    result = QueryServiceStatusEx(h, infolevel, (LPBYTE) &ssp, sizeof(ssp), &numbytes);
    if (! result) {
        return TwapiReturnSystemError(interp);
    }

    rec[0] = STRING_LITERAL_OBJ("servicetype");
    rec[1] = ObjFromServiceType(ssp.dwServiceType);
    rec[2] = STRING_LITERAL_OBJ("state");
    rec[3] = ObjFromServiceState(ssp.dwCurrentState);
    rec[4] = STRING_LITERAL_OBJ("dwControlsAccepted");
    rec[5] = Tcl_NewLongObj(ssp.dwControlsAccepted);
    rec[6] = STRING_LITERAL_OBJ("dwWin32ExitCode");
    rec[7] = Tcl_NewLongObj(ssp.dwWin32ExitCode);
    rec[8] = STRING_LITERAL_OBJ("dwServiceSpecificExitCode");
    rec[9] = Tcl_NewLongObj(ssp.dwServiceSpecificExitCode);
    rec[10] = STRING_LITERAL_OBJ("dwCheckPoint");
    rec[11] = Tcl_NewLongObj(ssp.dwCheckPoint);
    rec[12] = STRING_LITERAL_OBJ("dwWaitHint");
    rec[13] = Tcl_NewLongObj(ssp.dwWaitHint);
    rec[14] = STRING_LITERAL_OBJ("dwProcessId");
    rec[15] = Tcl_NewLongObj(ssp.dwProcessId);
    rec[16] = STRING_LITERAL_OBJ("dwServiceFlags");
    rec[17] = Tcl_NewLongObj(ssp.dwServiceFlags);
    rec[18] = STRING_LITERAL_OBJ("interactive");
    rec[19] = Tcl_NewLongObj(ssp.dwServiceType & SERVICE_INTERACTIVE_PROCESS);


    Tcl_SetObjResult(interp, Tcl_NewListObj(sizeof(rec)/sizeof(rec[0]), rec));
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

    objv[0] = STRING_LITERAL_OBJ("lpDependencies");
    objv[1] = ObjFromMultiSz_MAX(qbuf->lpDependencies);
    objv[2] = STRING_LITERAL_OBJ("servicetype");
    objv[3] = ObjFromServiceType(qbuf->dwServiceType);
    objv[4] = STRING_LITERAL_OBJ("dwStartType");
    objv[5] = Tcl_NewLongObj(qbuf->dwStartType);
    objv[6] = STRING_LITERAL_OBJ("dwErrorControl");
    objv[7] = Tcl_NewLongObj(qbuf->dwErrorControl);
    objv[8] = STRING_LITERAL_OBJ("dwTagId");
    objv[9] = Tcl_NewLongObj(qbuf->dwTagId);
    objv[10] = STRING_LITERAL_OBJ("lpBinaryPathName");
    objv[11] = ObjFromUnicode(qbuf->lpBinaryPathName);
    objv[12] = STRING_LITERAL_OBJ("lpLoadOrderGroup");
    objv[13] = ObjFromUnicode(qbuf->lpLoadOrderGroup);
    objv[14] = STRING_LITERAL_OBJ("lpServiceStartName");
    objv[15] = ObjFromUnicode(qbuf->lpServiceStartName);
    objv[16] = STRING_LITERAL_OBJ("lpDisplayName");
    objv[17] = ObjFromUnicode(qbuf->lpDisplayName);
    objv[18] = STRING_LITERAL_OBJ("interactive");
    objv[19] = Tcl_NewBooleanObj(qbuf->dwServiceType & SERVICE_INTERACTIVE_PROCESS);

    Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(20,objv));
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

    obj = Tcl_NewListObj(0, NULL);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, fIsLocked);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, obj, sbuf, lpLockOwner);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, dwLockDuration);

    MemLifoPopFrame(&ticP->memlifo);

    Tcl_SetObjResult(interp, obj);

    return TCL_OK;
}
#endif

/*
 * Helper function to retrieve list of services and status
 */
int Twapi_EnumServicesStatusEx(
    TwapiInterpContext *ticP,
    SC_HANDLE hService,
    SC_ENUM_TYPE infolevel,
    DWORD     dwServiceType,
    DWORD     dwServiceState,
    LPCWSTR   groupname
    )
{
    Tcl_Interp *interp = ticP->interp;
    ENUM_SERVICE_STATUS_PROCESSW *sbuf;
    DWORD buf_sz;
    DWORD buf_needed;
    DWORD services_returned;
    DWORD resume_handle;
    BOOL  success;
    Tcl_Obj *ra[2];     /* recordarray return value */
    Tcl_Obj *rec[12];    /* Holds values for each status record */
    Tcl_Obj *states[8]; /* Holds shared objects for states 1-7, 0 unused */
    DWORD winerr;
    DWORD i;
    TCL_RESULT status = TCL_ERROR;
    Tcl_Obj *service_typemap[6];

    if (infolevel != SC_ENUM_PROCESS_INFO) {
        Tcl_SetResult(interp, "Unsupported information level", TCL_STATIC);
        return Twapi_AppendSystemError(interp, ERROR_INVALID_PARAMETER);
    }

    /* 32000 - Initial estimate based on my system */
    sbuf = MemLifoPushFrame(&ticP->memlifo, 32000, &buf_sz);

    PopulateServiceTypeObjs(service_typemap);
    for (i=0; i < sizeof(states)/sizeof(states[0]); ++i) {
        states[i] = ObjFromServiceState(i);
        Tcl_IncrRefCount(states[i]);
    }

    ra[1] = Tcl_NewListObj(0, NULL);
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
            groupname);
        if ((!success) && ((winerr = GetLastError()) != ERROR_MORE_DATA)) {
            Twapi_FreeNewTclObj(ra[1]);
            Twapi_AppendSystemError(interp, winerr);
            goto pop_and_vamoose;
        }

        /* Tack on the services returned */
        for (i = 0; i < services_returned; ++i) {
            DWORD state;

            /* Note order should be same as order of field names below */
            rec[0] = ObjFromServiceTypeObjs(sbuf[i].ServiceStatusProcess.dwServiceType, service_typemap);
            state = sbuf[i].ServiceStatusProcess.dwCurrentState;
            if (state < (sizeof(states)/sizeof(states[0]))) {
                rec[1] = states[state];
            } else {
                rec[1] = Tcl_NewLongObj(state);
            }
            rec[2] = Tcl_NewLongObj(sbuf[i].ServiceStatusProcess.dwControlsAccepted);
            rec[3] = Tcl_NewLongObj(sbuf[i].ServiceStatusProcess.dwWin32ExitCode);
            rec[4] = Tcl_NewLongObj(sbuf[i].ServiceStatusProcess.dwServiceSpecificExitCode);
            rec[5] = Tcl_NewLongObj(sbuf[i].ServiceStatusProcess.dwCheckPoint);
            rec[6] = Tcl_NewLongObj(sbuf[i].ServiceStatusProcess.dwWaitHint);
            rec[7] = Tcl_NewLongObj(sbuf[i].ServiceStatusProcess.dwProcessId);
            rec[8] = Tcl_NewLongObj(sbuf[i].ServiceStatusProcess.dwServiceFlags);
            rec[9] = ObjFromUnicode(sbuf[i].lpServiceName); /* KEY for record array */
            rec[10] = ObjFromUnicode(sbuf[i].lpDisplayName);
            rec[11] = Tcl_NewBooleanObj(sbuf[i].ServiceStatusProcess.dwServiceType & SERVICE_INTERACTIVE_PROCESS);


            /* Note rec[9] object is also appended as the "key" for the "record" */
            Tcl_ListObjAppendElement(interp, ra[1], rec[9]);
            Tcl_ListObjAppendElement(interp, ra[1],
                                     Tcl_NewListObj(sizeof(rec)/sizeof(rec[0]), rec));
        }
        /* If !success -> ERROR_MORE_DATA so keep looping */
    } while (! success);


    /* Note order of names should be same as order of values above */
    /* Note order of field names should be same as order of values above */
    rec[0] = STRING_LITERAL_OBJ("servicetype");
    rec[1] = STRING_LITERAL_OBJ("state");
    rec[2] = STRING_LITERAL_OBJ("dwControlsAccepted");
    rec[3] = STRING_LITERAL_OBJ("dwWin32ExitCode");
    rec[4] = STRING_LITERAL_OBJ("dwServiceSpecificExitCode");
    rec[5] = STRING_LITERAL_OBJ("dwCheckPoint");
    rec[6] = STRING_LITERAL_OBJ("dwWaitHint");
    rec[7] = STRING_LITERAL_OBJ("dwProcessId");
    rec[8] = STRING_LITERAL_OBJ("dwServiceFlags");
    rec[9] = STRING_LITERAL_OBJ("lpServiceName");
    rec[10] = STRING_LITERAL_OBJ("lpDisplayName");
    rec[11] = STRING_LITERAL_OBJ("interactive");

    ra[0] = Tcl_NewListObj(sizeof(rec)/sizeof(rec[0]), rec);

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, ra));
    status = TCL_OK;

pop_and_vamoose:
    for (i=0; i < sizeof(states)/sizeof(states[0]); ++i) {
        Tcl_DecrRefCount(states[i]);
    }
    FreeServiceTypeObjs(service_typemap);

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
    Tcl_Obj *ra[2];     /* recordarray return value */
    Tcl_Obj *rec[10];    /* Holds values for each status record */
    Tcl_Obj *states[8]; /* Holds shared objects for states 1-7, 0 unused */
    DWORD winerr;
    DWORD i;
    TCL_RESULT status = TCL_ERROR;
    Tcl_Obj *service_typemap[6];

    sbuf = MemLifoPushFrame(&ticP->memlifo, 4000, &buf_sz);

    PopulateServiceTypeObjs(service_typemap);
    for (i=0; i < sizeof(states)/sizeof(states[0]); ++i) {
        states[i] = ObjFromServiceState(i);
        Tcl_IncrRefCount(states[i]);
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

    ra[1] = Tcl_NewListObj(0, NULL);
    /* Tack on the services returned */
    for (i = 0; i < services_returned; ++i) {
        DWORD state;

        /* Note order should be same as order of field names below */
        rec[0] = ObjFromServiceTypeObjs(sbuf[i].ServiceStatus.dwServiceType,
                                        service_typemap);
        state = sbuf[i].ServiceStatus.dwCurrentState;
        if (state < (sizeof(states)/sizeof(states[0]))) {
            rec[1] = states[state];
        } else {
            rec[1] = Tcl_NewLongObj(state);
        }
        rec[2] = Tcl_NewLongObj(sbuf[i].ServiceStatus.dwControlsAccepted);
        rec[3] = Tcl_NewLongObj(sbuf[i].ServiceStatus.dwWin32ExitCode);
        rec[4] = Tcl_NewLongObj(sbuf[i].ServiceStatus.dwServiceSpecificExitCode);
        rec[5] = Tcl_NewLongObj(sbuf[i].ServiceStatus.dwCheckPoint);
        rec[6] = Tcl_NewLongObj(sbuf[i].ServiceStatus.dwWaitHint);
        rec[7] = ObjFromUnicode(sbuf[i].lpServiceName); /* KEY for record array */
        rec[8] = ObjFromUnicode(sbuf[i].lpDisplayName);
        rec[9] = Tcl_NewLongObj(sbuf[i].ServiceStatus.dwServiceType & SERVICE_INTERACTIVE_PROCESS);

        /* Note rec[7] object is also appended as the "key" for the "record" */
        Tcl_ListObjAppendElement(interp, ra[1], rec[7]);
        Tcl_ListObjAppendElement(interp, ra[1],
                                 Tcl_NewListObj(sizeof(rec)/sizeof(rec[0]), rec));
    }


    /* Note order of field names should be same as order of values above */
    rec[0] = STRING_LITERAL_OBJ("servicetype");
    rec[1] = STRING_LITERAL_OBJ("state");
    rec[2] = STRING_LITERAL_OBJ("dwControlsAccepted");
    rec[3] = STRING_LITERAL_OBJ("dwWin32ExitCode");
    rec[4] = STRING_LITERAL_OBJ("dwServiceSpecificExitCode");
    rec[5] = STRING_LITERAL_OBJ("dwCheckPoint");
    rec[6] = STRING_LITERAL_OBJ("dwWaitHint");
    rec[7] = STRING_LITERAL_OBJ("lpServiceName");
    rec[8] = STRING_LITERAL_OBJ("lpDisplayName");
    rec[9] = STRING_LITERAL_OBJ("interactive");

    ra[0] = Tcl_NewListObj(sizeof(rec)/sizeof(rec[0]), rec);

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, ra));

    status = TCL_OK;

pop_and_vamoose:
    for (i=0; i < sizeof(states)/sizeof(states[0]); ++i) {
        Tcl_DecrRefCount(states[i]);
    }
    FreeServiceTypeObjs(service_typemap);

    MemLifoPopFrame(&ticP->memlifo);
    return status;
}


int Twapi_ChangeServiceConfig(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]) {
    SC_HANDLE h;
    DWORD service_type;
    DWORD start_type;
    DWORD error_control;
    DWORD tag_id;
    DWORD *tag_idP;
    LPWSTR dependencies = NULL;
    int result = 0;
    LPWSTR path, logrp, start_name, password, display_name;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLET(h, SC_HANDLE),
                     GETINT(service_type), GETINT(start_type),
                     GETINT(error_control), GETNULLTOKEN(path),
                     GETNULLTOKEN(logrp), ARGSKIP, ARGSKIP,
                     GETNULLTOKEN(start_name),
                     GETNULLTOKEN(password), GETNULLTOKEN(display_name),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (Tcl_GetLongFromObj(interp, objv[6], &tag_id) == TCL_OK)
        tag_idP = &tag_id;
    else {
        /* An empty string means value is not to be changed. Else error */
        if (Tcl_GetCharLength(objv[6]) != 0)
            return TCL_ERROR;   /* interp already holds error */
        tag_idP = NULL;         /* Tag is not to be changed */
    }

    dependencies = Tcl_GetUnicode(objv[7]);
    if (lstrcmpW(dependencies, NULL_TOKEN_L) == 0) {
        dependencies = NULL;
    } else {
        if (ObjToMultiSz(interp, objv[7], (LPCWSTR*) &dependencies) == TCL_ERROR)
            goto vamoose;
    }
    
    result = ChangeServiceConfigW(h, service_type, start_type, error_control,
                                  path, logrp, tag_idP, dependencies,
                                  start_name, password, display_name);
    if (result) {
        /* If caller wants new tag id returned (by not specifying
         * an empty tagid, return it, else return empty string.
         */
        if (tag_idP)
            Tcl_SetObjResult(interp, Tcl_NewLongObj(*tag_idP));
    } else
        TwapiReturnSystemError(interp);

vamoose:    
    if (dependencies)
        TwapiFree(dependencies);

    return result ? TCL_OK : TCL_ERROR;
}


int
Twapi_CreateService(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]) {
    SC_HANDLE    scmH;
    SC_HANDLE    svcH;
    DWORD        access;
    DWORD        service_type;
    DWORD        start_type;
    DWORD        error_control;
    DWORD tag_id;
    DWORD *tag_idP;
    LPWSTR dependencies = NULL;
    int tcl_result = TCL_ERROR;
    LPWSTR service_name, display_name, path, logrp;
    LPWSTR service_start_name, password;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLET(scmH, SC_HANDLE),
                     GETWSTR(service_name), GETWSTR(display_name),
                     GETINT(access), GETINT(service_type),
                     GETINT(start_type), GETINT(error_control),
                     GETWSTR(path), GETWSTR(logrp),
                     ARGSKIP, ARGSKIP,
                     GETNULLIFEMPTY(service_start_name),
                     GETNULLIFEMPTY(password),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;


    if (Tcl_GetLongFromObj(NULL, objv[9], &tag_id) == TCL_OK)
        tag_idP = &tag_id;
    else {
        /* An empty string means value is not to be changed. Else error */
        if (Tcl_GetCharLength(objv[9]) != 0)
            return TCL_ERROR;   /* interp already holds error */
        tag_idP = NULL;         /* Tag is not to be changed */
    }

    dependencies = Tcl_GetUnicode(objv[10]);
    if (lstrcmpW(dependencies, NULL_TOKEN_L) == 0) {
        dependencies = NULL;
    } else {
        if (ObjToMultiSz(interp, objv[10], (LPCWSTR*) &dependencies) == TCL_ERROR)
            goto vamoose;
    }

    svcH = CreateServiceW(
        scmH, service_name, display_name, access, service_type,
        start_type, error_control, path, logrp,
        tag_idP, dependencies, service_start_name, password);

    /* Check return handle validity */
    if (svcH) {
        Tcl_SetObjResult(interp, ObjFromOpaque(svcH, "SC_HANDLE"));
        tcl_result = TCL_OK;
    } else {
        TwapiReturnSystemError(interp);
    }

vamoose:
    if (dependencies)
        TwapiFree(dependencies);

    return tcl_result;
}

int Twapi_StartService(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]) 
{
    SC_HANDLE svcH;
    LPCWSTR args[64];
    Tcl_Obj **argObjs;
    int     i, nargs;
    
    if (objc != 2) {
        Tcl_SetResult(interp, "Invalid number of arguments.", TCL_STATIC);
        return TCL_ERROR;
    }

    if (ObjToHANDLE(interp, objv[0], &svcH) != TCL_OK)
        return TCL_ERROR;

    if (Tcl_ListObjGetElements(interp, objv[1], &nargs, &argObjs) == TCL_ERROR)
        return TCL_ERROR;

    if (nargs > ARRAYSIZE(args)) {
        Tcl_SetResult(interp, "Exceeded limit on number of service arguments.", TCL_STATIC);
        return TCL_ERROR;
    }

    for (i = 0; i < nargs; i++) {
        args[i] = Tcl_GetUnicode(argObjs[i]);
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


