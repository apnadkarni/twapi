/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

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
        if (! (BOOL) WTSQuerySessionInformationW(wtsH, sess_id, info_class, &bufP, &bytes )) {
            winerr = GetLastError();
            goto handle_error;

        }
        /* Note bufP can be NULL even on success! */
        if (bufP)
            Tcl_SetObjResult(interp, ObjFromUnicode(bufP));
        break;

    case WTSClientBuildNumber:
    case WTSClientHardwareId:
    case WTSConnectState:
    case WTSSessionId:
        if (! (BOOL) WTSQuerySessionInformationW(wtsH, sess_id, info_class, &bufP, &bytes )) {
            winerr = GetLastError();
            goto handle_error;
        }
        if (bufP)
            Tcl_SetObjResult(interp, Tcl_NewLongObj(*(long *)bufP));
        break;
        
    case WTSClientProductId:
    case WTSClientProtocolType:
        if (! (BOOL) WTSQuerySessionInformationW(wtsH, sess_id, info_class, &bufP, &bytes )) {
            winerr = GetLastError();
            goto handle_error;
        }
        if (bufP)
            Tcl_SetObjResult(interp, Tcl_NewLongObj(*(USHORT *)bufP));
        break;

    case WTSClientAddress: /* TBD */
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


static int Twapi_RDSCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s, s2;
    DWORD dw, dw2, dw3, dw4;
    HANDLE h;
    int i, i2;
    TwapiResult result;

    /* At least one arg for every command */
    if (objc < 3)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.type = TRT_HANDLE;
        result.value.hval = WTSOpenServerW(Tcl_GetUnicode(objv[2]));
        break;
    case 2:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETINT(dw),
                         GETWSTRN(s, i), GETWSTRN(s2, i2),
                         GETINT(dw2), GETINT(dw3), GETBOOL(dw4),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (WTSSendMessageW(h, dw, s, sizeof(WCHAR)*i,
                            s2, sizeof(WCHAR)*i2, dw2,
                            dw3, &result.value.uval, dw4))
            result.type = TRT_DWORD;
        else
            result.type = TRT_GETLASTERROR;    
        break;
    default:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(h),
                         ARGUSEDEFAULT, GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 101:
            return Twapi_WTSEnumerateSessions(interp, h);
        case 102:
            /* h == NULL -> current server and does not need to be closed */
            if (h)
                WTSCloseServer(h);
            result.type = TRT_EMPTY;
            break;
        case 103:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = WTSDisconnectSession(h, dw, dw2);
            break;
        case 104:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = WTSLogoffSession(h, dw, dw2);
            break;
        case 105:
            return Twapi_WTSQuerySessionInformation(interp, h, dw, dw2);
        }
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_RDSInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::RDSCall", Twapi_RDSCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::RDSCall", # code_); \
    } while (0);

    CALL_(WTSOpenServer, 1);
    CALL_(WTSSendMessage, 2);
    CALL_(WTSEnumerateSessions, 101);
    CALL_(WTSCloseServer, 102);
    CALL_(WTSDisconnectSession, 103);
    CALL_(WTSLogoffSession, 104);        /* TBD - tcl wrapper */
    CALL_(WTSQuerySessionInformation, 105); /* TBD - tcl wrapper */

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
int Twapi_rds_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, WLITERAL(MODULENAME), MODULE_HANDLE,
                            Twapi_RDSInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}


// TBD - Needs XP - DWORD WTSGetActiveConsoleSessionId(void);


