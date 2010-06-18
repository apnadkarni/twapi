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


#ifdef OBSOLETE
static Tcl_Obj *NamesFromENUM_SERVICE_STATUSW(void)
{
    Tcl_Obj *objP;
    // Order MUST BE same as in ObjFromENUM_SERVICE_STATUSW
    objP = NamesFromSERVICE_STATUS();
    Tcl_ListObjAppendElement(NULL, objP, STRING_LITERAL_OBJ("lpServiceName"));
    Tcl_ListObjAppendElement(NULL, objP, STRING_LITERAL_OBJ("lpDisplayName"));
    return objP;
}

static Tcl_Obj *ObjFromENUM_SERVICE_STATUSW(ENUM_SERVICE_STATUSW *essP)
{
    Tcl_Obj *objP;
    // Order MUST BE same as in NamesFromENUM_SERVICE_STATUSW
    objP = ObjFromSERVICE_STATUS(&essP->ServiceStatus);
    Tcl_ListObjAppendElement(NULL, objP, TWAPI_NEW_UNICODE_OBJ(essP->lpServiceName));
    Tcl_ListObjAppendElement(NULL, objP, TWAPI_NEW_UNICODE_OBJ(essP->lpDisplayName));
    return objP;
}

static Tcl_Obj *NamesFromENUM_SERVICE_STATUS_PROCESSW(void)
{
    Tcl_Obj *objP;
    // Order MUST BE same as in ObjFromENUM_SERVICE_STATUS_PROCESSW
    objP = NamesFromSERVICE_STATUS_PROCESS();
    Tcl_ListObjAppendElement(NULL, objP, STRING_LITERAL_OBJ("lpServiceName"));
    Tcl_ListObjAppendElement(NULL, objP, STRING_LITERAL_OBJ("lpDisplayName"));
    return objP;
    
}
#endif

#if 0
static Tcl_Obj *ObjFromENUM_SERVICE_STATUS_PROCESSW(ENUM_SERVICE_STATUS_PROCESSW *essP)
{
    Tcl_Obj *objP;

    // Order MUST BE same as in NamesFromENUM_SERVICE_STATUS_PROCESSW
    objP = ObjFromSERVICE_STATUS_PROCESS(&essP->ServiceStatusProcess);
    Tcl_ListObjAppendElement(NULL, objP, TWAPI_NEW_UNICODE_OBJ(essP->lpServiceName));
    Tcl_ListObjAppendElement(NULL, objP, TWAPI_NEW_UNICODE_OBJ(essP->lpDisplayName));
    return objP;
}
#endif

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
        objv[5] =objmaker(lpBinaryPathName,  TWAPI_NEW_UNICODE_OBJ, structp); \
        objv[6] =objmaker(lpLoadOrderGroup,  TWAPI_NEW_UNICODE_OBJ, structp); \
        objv[7] =objmaker(lpServiceStartName,  TWAPI_NEW_UNICODE_OBJ, structp); \
        objv[8] =objmaker(lpDisplayName,  TWAPI_NEW_UNICODE_OBJ, structp); \
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

/*
 * Helper function to retrieve service config info
 * Returns a single malloc'ed block after figuring out required size
 */
int Twapi_QueryServiceConfig(Tcl_Interp *interp, SC_HANDLE hService)
{
    QUERY_SERVICE_CONFIGW *qbuf;
    DWORD buf_sz;
    Tcl_Obj *objv[2];

    /* First find out how big a buffer we need */
    if (! QueryServiceConfigW(hService, NULL, 0, &buf_sz)) {
        /* For any error other than size, return */
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
            return TwapiReturnSystemError(interp);
    }

    /* Allocate it */
    if (Twapi_malloc(interp, NULL, buf_sz, &qbuf) != TCL_OK)
        return TCL_ERROR;

    /* Get the configuration information.  */
    if (! QueryServiceConfigW(hService, qbuf, buf_sz, &buf_sz)) {
        TwapiReturnSystemError(interp); /* Store error before calling free */
        free(qbuf);
        return TCL_ERROR;
    }

    objv[0] = NamesFromQUERY_SERVICE_CONFIGW();
    objv[1] = ObjFromQUERY_SERVICE_CONFIGW(qbuf);
    free(qbuf);
    Tcl_SetObjResult(interp, Tcl_NewListObj(2,objv));
    return TCL_OK;
}

#define NOOP_BEYOND_VISTA // TBD - remove this functionality ?
#ifdef NOOP_BEYOND_VISTA
/*
 * Helper function to retrieve service lock status info
 * Returns a single malloc'ed block after figuring out required size
 */
int  Twapi_QueryServiceLockStatus(
    Tcl_Interp *interp,
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
            return TwapiReturnSystemError(interp);
    }

    /* Allocate it */
    if (Twapi_malloc(interp, NULL, buf_sz, &sbuf) != TCL_OK)
        return TCL_ERROR;
    if (sbuf) {
        /* Get the configuration information.  */
        if (! QueryServiceLockStatusW(hService, sbuf, buf_sz, &buf_sz)) {
            TwapiReturnSystemError(interp); // Store before call to free
            free(sbuf);
            return TCL_ERROR;
        }
    }

    obj = Tcl_NewListObj(0, NULL);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, fIsLocked);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, obj, sbuf, lpLockOwner);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, sbuf, dwLockDuration);

    Tcl_SetObjResult(interp, obj);
    return TCL_OK;
}
#endif

/*
 * Helper function to retrieve list of services and status
 */
int Twapi_EnumServicesStatusEx(
    Tcl_Interp *interp,
    SC_HANDLE hService,
    SC_ENUM_TYPE infolevel,
    DWORD     dwServiceType,
    DWORD     dwServiceState,
    LPCWSTR   groupname
    )
{
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


    buf_sz = 32000; /* Initial estimate based on my system */
    resume_handle = 0;
    if (Twapi_malloc(interp, NULL, buf_sz, &sbuf) != TCL_OK)
        return TCL_ERROR;

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
            free(sbuf);
            return TCL_ERROR;
        }

        /* Tack on the services returned */
        for (i = 0; i < services_returned; ++i) {
            Tcl_Obj *objP;
            Tcl_Obj *keyP;

            objP = ObjFromSERVICE_STATUS_PROCESS(&sbuf[i].ServiceStatusProcess);

            /* Note order of values should be same as order of names below */
            keyP = TWAPI_NEW_UNICODE_OBJ(sbuf[i].lpServiceName);
            Tcl_ListObjAppendElement(NULL, objP, keyP);
            Tcl_ListObjAppendElement(NULL, objP,
                                     TWAPI_NEW_UNICODE_OBJ(sbuf[i].lpDisplayName));

            Tcl_ListObjAppendElement(interp, objv[1], keyP);
            Tcl_ListObjAppendElement(interp, objv[1], objP);
        }
        /* If !success -> ERROR_MORE_DATA so keep looping */
    } while (! success);

    /* Note order of names should be same as order of values above */
    objv[0] = NamesFromSERVICE_STATUS_PROCESS();
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpServiceName"));
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpDisplayName"));

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));

    free(sbuf);
    return TCL_OK;
}

int Twapi_EnumDependentServices(
    Tcl_Interp *interp,
    SC_HANDLE hService,
    DWORD     dwServiceState
    )
{
    ENUM_SERVICE_STATUSW *sbuf;
    DWORD buf_sz;
    DWORD services_returned;
    BOOL  success;
    Tcl_Obj *objv[2];
    DWORD winerr;
    DWORD i;

    buf_sz = 4000;
    if (Twapi_malloc(interp, NULL, buf_sz, &sbuf) != TCL_OK)
        return TCL_ERROR;

    do {
        success = EnumDependentServicesW(hService,
                                         dwServiceState,
                                         sbuf,
                                         buf_sz,
                                         &buf_sz,
                                         &services_returned);
        if (success)
            break;
        if ((winerr = GetLastError()) != ERROR_MORE_DATA)
            break;

        /* Need a bugger buffer */
        free(sbuf);
        if (Twapi_malloc(interp, NULL, buf_sz, &sbuf) != TCL_OK)
            return TCL_ERROR;
    } while (1);

    objv[1] = Tcl_NewListObj(0, NULL);
    /* Tack on the services returned */
    for (i = 0; i < services_returned; ++i) {
        Tcl_Obj *objP;
        Tcl_Obj *keyP;

        objP = ObjFromSERVICE_STATUS(&sbuf[i].ServiceStatus);
        keyP = TWAPI_NEW_UNICODE_OBJ(sbuf[i].lpServiceName);

        /* Note order should be same as order of field names below */
        Tcl_ListObjAppendElement(NULL, objP, keyP);
        Tcl_ListObjAppendElement(NULL, objP, TWAPI_NEW_UNICODE_OBJ(sbuf[i].lpDisplayName));

        /* Note keyP object is also appended as the "key" for the "record" */
        Tcl_ListObjAppendElement(interp, objv[1], keyP);
        Tcl_ListObjAppendElement(interp, objv[1], objP);
    }

    /* Note order of field names should be same as order of values above */
    objv[0] = NamesFromSERVICE_STATUS();
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpServiceName"));
    Tcl_ListObjAppendElement(NULL, objv[0], STRING_LITERAL_OBJ("lpDisplayName"));
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));

    free(sbuf);
    return TCL_OK;
}

#ifdef OBSOLETE
int Twapi_ControlService(
    Tcl_Interp *interp,
    SC_HANDLE           hService,
    DWORD               dwControl
    )
{
    SERVICE_STATUS ss;

    /* TBD - we currently throw away contents of ss. Not clear it is useful */

    if (ControlService(hService, dwControl, &ss) == 0) {
        return TwapiReturnSystemError(interp);
    }

    return TCL_OK;
}
#endif

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
    if (wcscmp(dependencies, NULL_TOKEN_L) == 0) {
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
        free(dependencies);

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
                     GETWSTR(service_start_name), GETWSTR(password),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;


    if (Tcl_GetLongFromObj(interp, objv[9], &tag_id) == TCL_OK)
        tag_idP = &tag_id;
    else {
        /* An empty string means value is not to be changed. Else error */
        if (Tcl_GetCharLength(objv[9]) != 0)
            return TCL_ERROR;   /* interp already holds error */
        tag_idP = NULL;         /* Tag is not to be changed */
    }

    dependencies = Tcl_GetUnicode(objv[10]);
    if (wcscmp(dependencies, NULL_TOKEN_L) == 0) {
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
    }

vamoose:
    if (dependencies)
        free(dependencies);

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

//#define NOOP_BEYOND_VISTA // TBD - decide whether to keep these calls

#ifdef NOOP_BEYOND_VISTA
/*
 * SC_LOCK is a opaque pointer that is a handle to a lock on the SCM database
 * We treat it as a pointer except that on return value, if it is NULL,
 * but the error is something other than database already being locked,
 * we generate a TCL exception. For database locked error, we will
 * return the string "NULL"
 */
%apply SWIGTYPE * {SC_LOCK};
%typemap(out) SC_LOCK %{
    if ($1 == NULL) {
        DWORD error_code;
        error_code = GetLastError();
        if (error_code != ERROR_SERVICE_DATABASE_LOCKED) {
            Twapi_AppendSystemError(interp, error_code);
            SWIG_fail;
        }
    }
    Tcl_SetObjResult(interp,SWIG_NewPointerObj((void *) $1, $1_descriptor,0));
%}

SC_LOCK LockServiceDatabase(SC_HANDLE   hSCManager);
EXCEPTION_ON_FALSE UnlockServiceDatabase(SC_LOCK ScLock);
%apply int Tcl_Result {int Twapi_QueryServiceLockStatus};
%rename(QueryServiceLockStatus) Twapi_QueryServiceLockStatus;
int Twapi_QueryServiceLockStatus(
    Tcl_Interp *interp,
    SC_HANDLE hSCManager
    );
#endif // NOOP_BEYOND_VISTA


