/*
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows NT services */
#include "twapi.h"

#define RETURN_SS_FIELDS(objmaker, structp)             \
    do { \
        Tcl_Obj *objv[7];                                             \
        objv[0] = objmaker(dwServiceType, Tcl_NewLongObj, structp);           \
        objv[1] = objmaker(dwCurrentState, Tcl_NewLongObj, structp);          \
        objv[2] = objmaker(dwControlsAccepted, Tcl_NewLongObj, structp);      \
        objv[3] = objmaker(dwWin32ExitCode, Tcl_NewLongObj, structp);         \
        objv[4] = objmaker(dwServiceSpecificExitCode, Tcl_NewLongObj, structp); \
        objv[5] = objmaker(dwCheckPoint, Tcl_NewLongObj, structp);            \
        objv[6] = objmaker(dwWaitHint, Tcl_NewLongObj, structp);              \
        return Tcl_NewListObj(sizeof(objv)/sizeof(objv[0]), objv);   \
    } while (0)

static Tcl_Obj *NamesFromSERVICE_STATUS(void) 
{
    RETURN_SS_FIELDS(FIELD_NAME_OBJ, notused);
}

static Tcl_Obj *ObjFromSERVICE_STATUS(SERVICE_STATUS *ssP)
{
    RETURN_SS_FIELDS(FIELD_VALUE_OBJ, ssP);
}
#undef RETURN_SS_FIELDS

#define RETURN_SSP_FIELDS(objmaker, structp)             \
    do { \
        Tcl_Obj *objv[9];                                             \
        objv[0] = objmaker(dwServiceType, Tcl_NewLongObj, structp);           \
        objv[1] = objmaker(dwCurrentState, Tcl_NewLongObj, structp);          \
        objv[2] = objmaker(dwControlsAccepted, Tcl_NewLongObj, structp);      \
        objv[3] = objmaker(dwWin32ExitCode, Tcl_NewLongObj, structp);         \
        objv[4] = objmaker(dwServiceSpecificExitCode, Tcl_NewLongObj, structp); \
        objv[5] = objmaker(dwCheckPoint, Tcl_NewLongObj, structp);            \
        objv[6] = objmaker(dwWaitHint, Tcl_NewLongObj, structp);              \
        objv[7] = objmaker(dwProcessId, Tcl_NewLongObj, structp);             \
        objv[8] = objmaker(dwServiceFlags, Tcl_NewLongObj, structp);          \
        return Tcl_NewListObj(sizeof(objv)/sizeof(objv[0]), objv);   \
    } while (0)

static Tcl_Obj *NamesFromSERVICE_STATUS_PROCESS(void) 
{
    RETURN_SSP_FIELDS(FIELD_NAME_OBJ, notused);
}

static Tcl_Obj *ObjFromSERVICE_STATUS_PROCESS(SERVICE_STATUS_PROCESS *ssP)
{
    RETURN_SSP_FIELDS(FIELD_VALUE_OBJ, ssP);
}
#undef RETURN_SSP_FIELDS



int Twapi_QueryServiceStatusEx(Tcl_Interp *interp, SC_HANDLE h,
        SC_STATUS_TYPE infolevel)
{
    SERVICE_STATUS_PROCESS ssp;
    DWORD numbytes;
    DWORD result;
    Tcl_Obj *objv[2];

    if (infolevel != SC_STATUS_PROCESS_INFO) {
        return Twapi_AppendSystemError(interp, ERROR_INVALID_LEVEL);
    }

    result = QueryServiceStatusEx(h, infolevel, (LPBYTE) &ssp, sizeof(ssp), &numbytes);
    if (! result) {
        return TwapiReturnSystemError(interp);
    }

    objv[0] = NamesFromSERVICE_STATUS_PROCESS();
    objv[1] = ObjFromSERVICE_STATUS_PROCESS(&ssp);
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
    return TCL_OK;
}


#define RETURN_QSC_FIELDS(objmaker, structp) \
    do { \
        Tcl_Obj *objv[9]; \
        objv[0] = objmaker(lpDependencies,  ObjFromMultiSz_MAX, structp); \
        objv[1] = objmaker(dwServiceType, Tcl_NewLongObj, structp); \
        objv[2] = objmaker(dwStartType, Tcl_NewLongObj, structp); \
        objv[3] = objmaker(dwErrorControl, Tcl_NewLongObj, structp); \
        objv[4] = objmaker(dwTagId, Tcl_NewLongObj, structp); \
        objv[5] = objmaker(lpBinaryPathName,  ObjFromUnicode, structp); \
        objv[6] = objmaker(lpLoadOrderGroup,  ObjFromUnicode, structp); \
        objv[7] = objmaker(lpServiceStartName,  ObjFromUnicode, structp); \
        objv[8] = objmaker(lpDisplayName,  ObjFromUnicode, structp); \
        return Tcl_NewListObj(sizeof(objv)/sizeof(objv[0]), objv);   \
    } while (0)



static Tcl_Obj *NamesFromQUERY_SERVICE_CONFIGW(void)
{
    RETURN_QSC_FIELDS(FIELD_NAME_OBJ, notused);
}

static Tcl_Obj *ObjFromQUERY_SERVICE_CONFIGW(QUERY_SERVICE_CONFIGW *qP)
{
    RETURN_QSC_FIELDS(FIELD_VALUE_OBJ, qP);
}
#undef RETURN_QSC_FIELDS

int Twapi_QueryServiceConfig(TwapiInterpContext *ticP, SC_HANDLE hService)
{
    QUERY_SERVICE_CONFIGW *qbuf;
    DWORD buf_sz;
    Tcl_Obj *objv[2];
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

    objv[0] = NamesFromQUERY_SERVICE_CONFIGW();
    objv[1] = ObjFromQUERY_SERVICE_CONFIGW(qbuf);
    Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(2,objv));
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
    Tcl_Obj *objv[2];
    DWORD winerr;
    DWORD i;

    if (infolevel != SC_ENUM_PROCESS_INFO) {
        Tcl_SetResult(interp, "Unsupported information level", TCL_STATIC);
        return Twapi_AppendSystemError(interp, ERROR_INVALID_PARAMETER);
    }

    /* 32000 - Initial estimate based on my system */
    sbuf = MemLifoPushFrame(&ticP->memlifo, 32000, &buf_sz);
    resume_handle = 0;

    objv[1] = Tcl_NewListObj(0, NULL);
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
            Twapi_FreeNewTclObj(objv[1]);
            Twapi_AppendSystemError(interp, winerr);
            MemLifoPopFrame(&ticP->memlifo);
            return TCL_ERROR;
        }

        /* Tack on the services returned */
        for (i = 0; i < services_returned; ++i) {
            Tcl_Obj *objP;
            Tcl_Obj *keyP;

            objP = ObjFromSERVICE_STATUS_PROCESS(&sbuf[i].ServiceStatusProcess);

            /* Note order of values should be same as order of names below */
            keyP = ObjFromUnicode(sbuf[i].lpServiceName);
            Tcl_ListObjAppendElement(NULL, objP, keyP);
            Tcl_ListObjAppendElement(NULL, objP,
                                     ObjFromUnicode(sbuf[i].lpDisplayName));

            Tcl_ListObjAppendElement(interp, objv[1], keyP);
            Tcl_ListObjAppendElement(interp, objv[1], objP);
        }
        /* If !success -> ERROR_MORE_DATA so keep looping */
    } while (! success);

    MemLifoPopFrame(&ticP->memlifo);

    /* Note order of names should be same as order of values above */
    objv[0] = NamesFromSERVICE_STATUS_PROCESS();
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpServiceName"));
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpDisplayName"));

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));

    return TCL_OK;
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
    Tcl_Obj *objv[2];
    DWORD winerr;
    DWORD i;

    sbuf = MemLifoPushFrame(&ticP->memlifo, 4000, &buf_sz);
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
        MemLifoPopFrame(&ticP->memlifo);
        if (winerr != ERROR_MORE_DATA) 
            return Twapi_AppendSystemError(interp, winerr);

        /* Need a bigger buffer */
        sbuf = MemLifoPushFrame(&ticP->memlifo, buf_sz, NULL);
    } while (1);

    objv[1] = Tcl_NewListObj(0, NULL);
    /* Tack on the services returned */
    for (i = 0; i < services_returned; ++i) {
        Tcl_Obj *objP;
        Tcl_Obj *keyP;

        objP = ObjFromSERVICE_STATUS(&sbuf[i].ServiceStatus);
        keyP = ObjFromUnicode(sbuf[i].lpServiceName);

        /* Note order should be same as order of field names below */
        Tcl_ListObjAppendElement(NULL, objP, keyP);
        Tcl_ListObjAppendElement(NULL, objP, ObjFromUnicode(sbuf[i].lpDisplayName));

        /* Note keyP object is also appended as the "key" for the "record" */
        Tcl_ListObjAppendElement(interp, objv[1], keyP);
        Tcl_ListObjAppendElement(interp, objv[1], objP);
    }

    /* Note order of field names should be same as order of values above */
    objv[0] = NamesFromSERVICE_STATUS();
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpServiceName"));
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpDisplayName"));
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));

    MemLifoPopFrame(&ticP->memlifo);
    return TCL_OK;
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
                     GETINT(error_control), GETWSTR(path),
                     GETWSTR(logrp), ARGSKIP, ARGSKIP, GETWSTR(start_name),
                     GETWSTR(password), GETWSTR(display_name),
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


