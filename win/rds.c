/*
 * Copyright (c) 2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

#ifndef MODULENAME
#define MODULENAME "twapi_rds"
#endif

int Twapi_WTSEnumerateSessions(Tcl_Interp *interp, HANDLE wtsH)
{
    WTS_SESSION_INFOW *sessP = NULL;
    DWORD count;
    DWORD i;
    Tcl_Obj **records;
    Tcl_Obj *fields = NULL;
    Tcl_Obj *objv[3];

    /* Note wtsH == NULL means current server */
    if (! (BOOL) (WTSEnumerateSessionsW)(wtsH, 0, 1, &sessP, &count)) {
        return TwapiReturnSystemError(interp);
    }

    records = SWSPushFrame(count * sizeof(Tcl_Obj*), NULL);
    for (i = 0; i < count; ++i) {
        objv[0] = ObjFromLong(sessP[i].SessionId);
        objv[1] = ObjFromWinChars(sessP[i].pWinStationName);
        objv[2] = ObjFromLong(sessP[i].State);

        records[i] = ObjNewList(3, objv);
    }

    Twapi_WTSFreeMemory(sessP);

    /* Make up the field names */
    objv[0] = STRING_LITERAL_OBJ("SessionId");
    objv[1] = STRING_LITERAL_OBJ("pWinStationName");
    objv[2] = STRING_LITERAL_OBJ("State");
    fields = ObjNewList(3, objv);

    objv[0] = fields;
    objv[1] = ObjNewList(count, records);
    ObjSetResult(interp, ObjNewList(2, objv));
    SWSPopFrame();
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
            ObjSetResult(interp, ObjFromWinChars(bufP));
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
            ObjSetResult(interp, ObjFromLong(*(long *)bufP));
        break;
        
    case WTSClientProductId:
    case WTSClientProtocolType:
        if (! (BOOL) WTSQuerySessionInformationW(wtsH, sess_id, info_class, &bufP, &bytes )) {
            winerr = GetLastError();
            goto handle_error;
        }
        if (bufP)
            ObjSetResult(interp, ObjFromLong(*(USHORT *)bufP));
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

    ObjSetStaticResult(interp, "Could not query TS information.");

    return Twapi_AppendSystemError(interp, winerr);
}


static int Twapi_RDSCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR s, s2;
    DWORD dw, dw2, dw3;
    BOOL bval;
    HANDLE h;
    TwapiResult result;
    int func = PtrToInt(clientdata);
    Tcl_Size len, len2;

    /* At least one arg for every command */
    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.type = TRT_HANDLE;
        result.value.hval = WTSOpenServerW(ObjToWinChars(objv[0]));
        break;
    case 2:
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLE(h), GETDWORD(dw),
                         ARGSKIP, ARGSKIP,
                         GETDWORD(dw2), GETDWORD(dw3), GETBOOL(bval),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        s = ObjToWinCharsN(objv[2], &len);
        s2 = ObjToWinCharsN(objv[3], &len2);
        CHECK_DWORD(interp, len*sizeof(WCHAR));
        CHECK_DWORD(interp, len2*sizeof(WCHAR));
        if (WTSSendMessageW(h, dw, s, (DWORD) (sizeof(WCHAR)*len),
                            s2, (DWORD) (sizeof(WCHAR)*len2), dw2,
                            dw3, &result.value.uval, bval))
            result.type = TRT_DWORD;
        else
            result.type = TRT_GETLASTERROR;
        break;
    default:
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h),
                         ARGUSEDEFAULT, GETDWORD(dw), GETDWORD(dw2),
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


static int TwapiRdsInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s RdsDispatch[] = {
        DEFINE_FNCODE_CMD(rds_open_server, 1), //WTSOpenServer
        DEFINE_FNCODE_CMD(WTSSendMessage, 2),
        DEFINE_FNCODE_CMD(WTSEnumerateSessions, 101),
        DEFINE_FNCODE_CMD(rds_close_server, 102), //WTSCloseServer
        DEFINE_FNCODE_CMD(WTSDisconnectSession, 103),
        DEFINE_FNCODE_CMD(WTSLogoffSession, 104), // TBD - tcl wrapper
        DEFINE_FNCODE_CMD(WTSQuerySessionInformation, 105), // TBD - tcl wrapper
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(RdsDispatch), RdsDispatch, Twapi_RDSCallObjCmd);

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
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiRdsInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}


// TBD - Needs XP - DWORD WTSGetActiveConsoleSessionId(void);


