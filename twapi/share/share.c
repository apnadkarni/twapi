/* 
 * Copyright (c) 2003-2009, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

Tcl_Obj *ObjFromCONNECTION_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromUSE_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromSHARE_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromFILE_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromSESSION_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromUSER_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromGROUP_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromLOCALGROUP_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromGROUP_USERS_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);

int Twapi_WNetGetUser(Tcl_Interp *interp, LPCWSTR  lpName);
int Twapi_NetScheduleJobEnum(Tcl_Interp *interp, LPCWSTR servername);
int Twapi_NetShareEnum(Tcl_Interp *interp, LPWSTR server_name);
int Twapi_NetUseGetInfo(Tcl_Interp *interp, LPWSTR UncServer, LPWSTR UseName, DWORD level);
int Twapi_NetShareCheck(Tcl_Interp *interp, LPWSTR server, LPWSTR device);
int Twapi_NetShareGetInfo(Tcl_Interp *interp, LPWSTR server,
                          LPWSTR netname, DWORD level);
int Twapi_NetShareSetInfo(Tcl_Interp *interp, LPWSTR server_name,
                          LPWSTR net_name, LPWSTR remark, DWORD  max_uses,
                          SECURITY_DESCRIPTOR *secd);
int Twapi_NetConnectionEnum(Tcl_Interp    *interp, LPWSTR server,
                            LPWSTR qualifier, DWORD level);
int Twapi_NetFileEnum(Tcl_Interp *interp, LPWSTR server, LPWSTR basepath,
                      LPWSTR user, DWORD level);
int Twapi_NetFileGetInfo(Tcl_Interp    *interp, LPWSTR server,
                         DWORD fileid, DWORD level);
int Twapi_NetSessionEnum(Tcl_Interp    *interp, LPWSTR server, LPWSTR client,
                         LPWSTR user, DWORD level);
int Twapi_NetSessionGetInfo(Tcl_Interp *interp, LPWSTR server,
                            LPWSTR client, LPWSTR user, DWORD level);
int Twapi_WNetUseConnection(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_NetShareAdd(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);


/* Maps parameter error codes returned by NetShare* functions to
 * descriptions
 */
char *TwapiLookupShareParamError(DWORD parm_err)
{
    switch (parm_err) {
    case SHARE_NETNAME_PARMNUM:
        return "Invalid network share name parameter";
    case SHARE_TYPE_PARMNUM:
        return "Invalid network share type parameter";
    case SHARE_REMARK_PARMNUM:
        return "Invalid network share remark parameter";
    case SHARE_PERMISSIONS_PARMNUM:
        return "Invalid network share permissions parameter";
    case SHARE_MAX_USES_PARMNUM:
        return "Invalid network share maximum connections parameter";
    case SHARE_CURRENT_USES_PARMNUM:
        return "Invalid network share current connections parameter";
    case SHARE_PATH_PARMNUM:
        return "Invalid network share path parameter";
    case SHARE_PASSWD_PARMNUM:
        return "Invalid network share password parameter";
    case SHARE_FILE_SD_PARMNUM:
        return "Invalid network share security descriptor parameter";
    }
    return "Invalid network share parameter";
}

Tcl_Obj *ObjFromCONNECTION_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    CONNECTION_INFO_1 *ciP;
    int         objc;
    Tcl_Obj    *objv[7];

    ciP = (CONNECTION_INFO_1 *) infoP; /* May actually be CONNECTION_INFO_0 */

    objc = 1;
    switch (level) {
    case 1:
        objc += 6;
        objv[6] = ObjFromWinChars(ciP->coni1_netname);
        objv[5] = ObjFromWinChars(ciP->coni1_username);
        objv[4] = ObjFromLong(ciP->coni1_time);
        objv[3] = ObjFromLong(ciP->coni1_num_users);
        objv[2] = ObjFromLong(ciP->coni1_num_opens);
        objv[1] = ObjFromLong(ciP->coni1_type);
        /* FALLTHRU */
    case 0:
        objv[0] = ObjFromLong(ciP->coni1_id);
        break;
    default:
        Twapi_WrongLevelError(interp, level);
        return NULL;
    }

    return ObjNewList(objc, objv);
}

/* Convert USE_INFO_* to list */
Tcl_Obj *ObjFromUSE_INFO(
    Tcl_Interp *interp,
    LPBYTE infoP,
    DWORD level
    )
{
    USE_INFO_2 *uiP;
    int         objc;
    Tcl_Obj    *objv[9];

    uiP = (USE_INFO_2 *) infoP; /* May acutally be any USE_INFO_* */

    objc = 2;
    switch (level) {
    case 2:
        objc += 2;
        objv[8] = ObjFromWinChars(uiP->ui2_domainname);
        objv[7] = ObjFromWinChars(uiP->ui2_username);
        // FALLTHRU
    case 1:
        objc += 5;
        /* Does this contain a valid value and need to be encrypted ? - TBD */
        objv[6] = ObjFromDWORD(uiP->ui2_usecount);
        objv[5] = ObjFromDWORD(uiP->ui2_refcount);
        objv[4] = ObjFromDWORD(uiP->ui2_asg_type);
        objv[3] = ObjFromDWORD(uiP->ui2_status);
        objv[2] = ObjFromWinChars(uiP->ui2_password);
        //FALLTHRU
    case 0:
        objv[1] = ObjFromWinChars(uiP->ui2_remote);
        objv[0] = ObjFromWinChars(uiP->ui2_local);
        break;

    default:
        Twapi_WrongLevelError(interp, level);
        return NULL;
    }

    return ObjNewList(objc, objv);
}

Tcl_Obj *ObjFromSHARE_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    SHARE_INFO_502 *siP;
    int         objc;
    Tcl_Obj    *objv[10];

    siP = (SHARE_INFO_502 *) infoP; /* May acutally be any SHARE_INFO_* */

    objc = 1;
    switch (level) {
    case 502:
        objc += 2;
        objv[9] = ObjFromSECURITY_DESCRIPTOR(interp,
                                            siP->shi502_security_descriptor);
        objv[8] = ObjFromDWORD(siP->shi502_reserved);
        // FALLTHRU
    case 2:
        objc += 5;
        objv[7] = ObjFromWinChars(siP->shi502_passwd ? siP->shi502_passwd : L"");
        objv[6] = ObjFromWinChars(siP->shi502_path);
        objv[5] = ObjFromLong(siP->shi502_current_uses);
        objv[4] = ObjFromLong(siP->shi502_max_uses);
        objv[3] = ObjFromDWORD(siP->shi502_permissions);
        // FALLTHRU
    case 1:
        objc += 2;
        objv[2] = ObjFromWinChars(siP->shi502_remark);
        objv[1] = ObjFromDWORD(siP->shi502_type);
        // FALLTHRU
    case 0:
        objv[0] = ObjFromWinChars(siP->shi502_netname);
        break;

    default:
        Twapi_WrongLevelError(interp, level);
        return NULL;
    }

    return ObjNewList(objc, objv);
}


Tcl_Obj *ObjFromFILE_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    int         objc;
    Tcl_Obj    *objv[5];
    FILE_INFO_3 *fiP;

    fiP = (FILE_INFO_3 *) infoP; /* May acutally be FILE_INFO_* */

    objc = 1;
    switch (level) {
    case 3:
        objc += 4;
        objv[4] = ObjFromWinChars(fiP->fi3_username);
        objv[3] = ObjFromWinChars(fiP->fi3_pathname);
        objv[2] = ObjFromLong(fiP->fi3_num_locks);
        objv[1] = ObjFromLong(fiP->fi3_permissions);
        /* FALLTHRU */
    case 2:
        objv[0] = ObjFromLong(fiP->fi3_id);
        break;
    default:
        Twapi_WrongLevelError(interp, level);
        return NULL;
    }

    return ObjNewList(objc, objv);
}


Tcl_Obj *ObjFromSESSION_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    int         objc;
    Tcl_Obj    *objv[8];

    /* Note in the code below, the structures can be superimposed on one
       another except level 10
    */

    objc = 0;
    switch (level) {
    case 502:
        objc += 1;
        objv[7] = ObjFromWinChars(((SESSION_INFO_502 *)infoP)->sesi502_transport);
        /* FALLTHRU */
    case 2:
        objc += 1;
        objv[6] = ObjFromWinChars(((SESSION_INFO_2 *)infoP)->sesi2_cltype_name);
        /* FALLTHRU */
    case 1:
        objc += 5;
        objv[5] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_user_flags);
        objv[4] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_idle_time);
        objv[3] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_time);
        objv[2] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_num_opens);
        objv[1] = ObjFromWinChars(((SESSION_INFO_1 *)infoP)->sesi1_username);
        /* FALLTHRU */
    case 0:
        objc += 1;
        objv[0] = ObjFromWinChars(((SESSION_INFO_0 *)infoP)->sesi0_cname);
        break;

    case 10:
        objc = 4;
        objv[0] = ObjFromWinChars(((SESSION_INFO_10 *)infoP)->sesi10_cname);
        objv[1] = ObjFromWinChars(((SESSION_INFO_10 *)infoP)->sesi10_username);
        objv[2] = ObjFromLong(((SESSION_INFO_10 *)infoP)->sesi10_time);
        objv[3] = ObjFromLong(((SESSION_INFO_10 *)infoP)->sesi10_idle_time);
        break;

    default:
        Twapi_WrongLevelError(interp, level);
        return NULL;
    }

    return ObjNewList(objc, objv);
}


/*
 * Convert NETRESOURCEW structure to Tcl list. Returns TCL_OK/TCL_ERROR
 * interp may be NULL
 */
static Tcl_Obj *ListObjFromNETRESOURCEW(
    Tcl_Interp      *interp,
    NETRESOURCEW    *nrP
    ) 
{
    Tcl_Obj *objs[8];

    objs[0] = ObjFromDWORD(nrP->dwScope);
    objs[1] = ObjFromDWORD(nrP->dwType);
    objs[2] = ObjFromDWORD(nrP->dwDisplayType);
    objs[3] = ObjFromDWORD(nrP->dwUsage);
    objs[4] = ObjFromWinChars(nrP->lpLocalName);
    objs[5] = ObjFromWinChars(nrP->lpRemoteName);
    objs[6] = ObjFromWinChars(nrP->lpComment);
    objs[7] = ObjFromWinChars(nrP->lpProvider);

    return ObjNewList(8, objs);
}    

/*
 * Convert REMOTE_NAME_INFOW structure to Tcl list. Returns TCL_OK/TCL_ERROR
 * interp may be NULL
 */
static Tcl_Obj *ListObjFromREMOTE_NAME_INFOW(
    Tcl_Interp         *interp,
    REMOTE_NAME_INFOW *rniP
    ) 
{
    Tcl_Obj *objv[3];

    objv[0] = ObjFromWinChars(rniP->lpUniversalName);
    objv[1] = ObjFromWinChars(rniP->lpConnectionName);
    objv[2] = ObjFromWinChars(rniP->lpRemainingPath);

    return ObjNewList(3, objv);
}

int Twapi_NetShareAdd(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR server_name;
    SECURITY_DESCRIPTOR *secdP;
    SHARE_INFO_502 share_info;
    DWORD          parm_err;
    NET_API_STATUS status;
    SWSMark mark;

    CHECK_NARGS(interp, objc, 7);

    /* Extract integer args before WSTRs to avoid shimmering issues */
    CHECK_INTEGER_OBJ(interp, share_info.shi502_type, objv[2]);
    CHECK_INTEGER_OBJ(interp, share_info.shi502_max_uses, objv[4]);

    mark = SWSPushMark();
    if (ObjToPSECURITY_DESCRIPTORSWS(interp, objv[6], &secdP) != TCL_OK) {
        SWSPopMark(mark);
        return TCL_ERROR;
    }
    server_name = ObjToLPWSTR_NULL_IF_EMPTY(objv[0]);
    share_info.shi502_netname = ObjToWinChars(objv[1]);
    share_info.shi502_remark  = ObjToWinChars(objv[3]);
    share_info.shi502_path    = ObjToWinChars(objv[5]);
    share_info.shi502_permissions = 0;
    share_info.shi502_current_uses = 0;
    share_info.shi502_passwd       = NULL;
    share_info.shi502_reserved     = 0;
    share_info.shi502_security_descriptor = secdP;
    
    status = NetShareAdd(server_name, 502, (unsigned char *)&share_info, &parm_err);
    SWSPopMark(mark);

    if (status == NERR_Success)
        return TCL_OK;
    
    Twapi_AppendSystemError(interp, status);
    if (status == ERROR_INVALID_PARAMETER)
        Tcl_AppendResult(interp,
                         TwapiLookupShareParamError(parm_err),
                         NULL);
    return TCL_ERROR;

}

int Twapi_NetShareCheck(
    Tcl_Interp *interp,
    LPWSTR server_name,
    LPWSTR device_name
)
{
    NET_API_STATUS status;
    DWORD          type;
    Tcl_Obj       *objv[2];

    status = NetShareCheck(server_name, device_name, &type);
    if (status == NERR_Success) {
        objv[0] = ObjFromLong(1);
        objv[1] = ObjFromDWORD(type);
        ObjSetResult(interp, ObjNewList(2, objv));
    }
    else if (status == NERR_DeviceNotShared) {
        objv[0] = ObjFromLong(0);
        objv[1] = ObjFromLong(0);
        ObjSetResult(interp, ObjNewList(2, objv));
    }
    else {
        return Twapi_AppendSystemError(interp, status);
    }

    return TCL_OK;
}


int Twapi_NetShareGetInfo(
    Tcl_Interp *interp,
    LPWSTR     servername,
    LPWSTR     netname,
    DWORD       level
)
{
    NET_API_STATUS status;
    LPBYTE         shareP;

    switch (level) {
    case 1:
    case 2:
    case 502:
        status = NetShareGetInfo(servername, netname, level, &shareP);
        if (status != NERR_Success) {
            ObjSetStaticResult(interp, "Could not retrieve share information: ");
            return Twapi_AppendSystemError(interp, status);
        }
        ObjSetResult(
            interp,
            ObjFromSHARE_INFO(interp, shareP, level)
            );
        NetApiBufferFree(shareP);
        break;

    default:
        ObjSetStaticResult(interp, "Invalid or unsupported share information level specified");
        return TCL_ERROR;
    }

    return TCL_OK;
}


int Twapi_NetShareSetInfo(
    Tcl_Interp *interp,
    LPWSTR server_name,
    LPWSTR net_name,
    LPWSTR remark,
    DWORD  max_uses,
    SECURITY_DESCRIPTOR *secd
)
{
    SHARE_INFO_502 *shareP;
    DWORD           parm_err;
    NET_API_STATUS  status;

    status = NetShareGetInfo(server_name, net_name, 502, (LPBYTE *) &shareP);
    if (status != NERR_Success) {
        ObjSetStaticResult(interp, "Could not retrieve share information: ");
        return Twapi_AppendSystemError(interp, status);
    }

    shareP->shi502_remark  = remark;
    shareP->shi502_max_uses = max_uses;
    shareP->shi502_reserved = 0;
    shareP->shi502_security_descriptor = secd;
    
    status = NetShareSetInfo(server_name, net_name, 502, (unsigned char *)shareP, &parm_err);
    if (status == NERR_Success)
        return TCL_OK;
    
    NetApiBufferFree(shareP);

    Twapi_AppendSystemError(interp, status);
    if (status == ERROR_INVALID_PARAMETER)
        Tcl_AppendResult(interp,
                         TwapiLookupShareParamError(parm_err),
                         NULL);
    return TCL_ERROR;
}


int Twapi_WNetUseConnection(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HWND          winH;
    LPWSTR  usernameP;
    int                   ignore_password;
    LPWSTR  passwordP, decrypted_password;
    int     password_len;
    DWORD   flags;
    WCHAR accessname[MAX_PATH];
    DWORD accessname_size;
    NETRESOURCEW netresource;
    DWORD       outflags;       /* Not really used */
    int         error;
    SWSMark mark = NULL;

    CHECK_NARGS(interp, objc, 9);
    /* TBD - what is the type of HWD ? SHould it not use ObjToHWND ? */
    if (ObjToHWND(interp, objv[0], &winH) != TCL_OK)
        return TCL_ERROR;
    CHECK_INTEGER_OBJ(interp, netresource.dwType, objv[1]);
    CHECK_INTEGER_OBJ(interp, ignore_password, objv[6]);
    CHECK_INTEGER_OBJ(interp, flags, objv[8]);
    netresource.lpLocalName = ObjToLPWSTR_NULL_IF_EMPTY(objv[2]);
    netresource.lpRemoteName = ObjToWinChars(objv[3]);
    netresource.lpProvider = ObjToLPWSTR_NULL_IF_EMPTY(objv[4]);
    usernameP = ObjToLPWSTR_NULL_IF_EMPTY(objv[5]);

    mark = SWSPushMark();
    decrypted_password = ObjDecryptPasswordSWS(objv[7], &password_len);

    if (ignore_password) {
        passwordP = L"";        /* Password is ignored */
    } else {
        passwordP = decrypted_password;
        NULLIFY_EMPTY(passwordP); /* If NULL, default password */
    }

    accessname_size = ARRAYSIZE(accessname);
    error = WNetUseConnectionW(winH, &netresource, passwordP, usernameP, flags,
                               accessname, &accessname_size, &outflags);

    SecureZeroMemory(decrypted_password, password_len);
    SWSPopMark(mark);

    if (error == NO_ERROR) {
        ObjSetResult(interp, ObjFromWinChars(accessname));
        return TCL_OK;
    }
    else {
        return Twapi_AppendWNetError(interp, error);
    }
}

static TCL_RESULT Twapi_WNetGetUniversalNameObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int result;
    DWORD error;
    DWORD buf_sz;
    void *buf;
    LPCWSTR      localpathP;

    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    localpathP = ObjToWinChars(objv[1]);

    buf = MemLifoPushFrame(ticP->memlifoP, MAX_PATH+1, &buf_sz);
    error = WNetGetUniversalNameW(localpathP, REMOTE_NAME_INFO_LEVEL,
                                  buf, &buf_sz);
    if (error = ERROR_MORE_DATA) {
        /* Retry with larger buffer */
        buf = MemLifoAlloc(ticP->memlifoP, buf_sz, NULL);
        error = WNetGetUniversalNameW(localpathP, REMOTE_NAME_INFO_LEVEL,
                                      buf, &buf_sz);
    }
    if (error == NO_ERROR) {
        ObjSetResult(ticP->interp,
                         ListObjFromREMOTE_NAME_INFOW(ticP->interp,
                                                      (REMOTE_NAME_INFOW *) buf));
        result = TCL_OK;
    } 
    else {
        Twapi_AppendWNetError(ticP->interp, error);
        result = TCL_ERROR;
    }

    MemLifoPopFrame(ticP->memlifoP);

    return result;
}

static TCL_RESULT Twapi_WNetGetResourceInformationObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    NETRESOURCEW in;
    NETRESOURCEW *outP;
    DWORD       outsz;
    int         error;
    LPWSTR      systempart;
    Tcl_Obj    *objs[2];
    MemLifoMarkHandle mark;
    TCL_RESULT res;

    mark = MemLifoPushMark(ticP->memlifoP);

    /* TBD - check if GETNULLIFEMPTY is ok */
    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                         GETEMPTYASNULL(in.lpRemoteName),
                         GETEMPTYASNULL(in.lpProvider),
                         GETINT(in.dwType),
                         ARGEND);
    if (res == TCL_OK) {
        outP = MemLifoAlloc(ticP->memlifoP, 4000, &outsz);
        error = WNetGetResourceInformationW(&in, outP, &outsz, &systempart);
        if (error == ERROR_MORE_DATA) {
            /* Retry with larger buffer */
            outP = MemLifoAlloc(ticP->memlifoP, outsz, NULL);
            error = WNetGetResourceInformationW(&in, outP, &outsz, &systempart);
        }
        if (error == ERROR_SUCCESS) {
            /* TBD - replace with ObjFromCStruct */
            objs[0] = ListObjFromNETRESOURCEW(ticP->interp, outP);
            objs[1] = ObjFromWinChars(systempart);
            ObjSetResult(ticP->interp, ObjNewList(2, objs));
            res = TCL_OK;
        } else
            res = Twapi_AppendWNetError(ticP->interp, error);
    }

    MemLifoPopMark(mark);
    return res;
}


static TCL_RESULT Twapi_WNetAddConnection3ObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TCL_RESULT res;
    DWORD       flags, wnet_status, sz;
    MemLifoMarkHandle mark;
    LPCWSTR     username;
    HWND hwnd;
    WCHAR         *decryptedP;
    int            decrypted_len;
    NETRESOURCEW *netresP;

    mark = MemLifoPushMark(ticP->memlifoP);

    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                         GETHWND(hwnd),
                         ARGSKIP,
                         ARGSKIP,
                         GETWSTR(username),
                         GETINT(flags),
                         ARGEND);
    if (res == TCL_OK) {
        res = TwapiCStructParse(interp, ticP->memlifoP, objv[2],
                                0, &sz, &netresP);
        if (res == TCL_OK) {
            if (sz != sizeof(NETRESOURCEW))
                res = TwapiReturnError(interp, TWAPI_INVALID_ARGS);
            else {
                TWAPI_ASSERT(ticP->memlifoP == SWS());
                decryptedP = ObjDecryptPasswordSWS(objv[3], &decrypted_len);
                wnet_status = WNetAddConnection3W(hwnd, netresP,
                                         decryptedP, username, flags);
                if (wnet_status != NO_ERROR)
                    res = Twapi_AppendWNetError(ticP->interp, wnet_status);

                SecureZeroMemory(decryptedP, decrypted_len);
            }
        }
    }

    MemLifoPopMark(mark);
    return res;
}


int Twapi_NetUseGetInfo(
    Tcl_Interp *interp,
    LPWSTR UncServerName,
    LPWSTR UseName,
    DWORD level
)
{
    NET_API_STATUS status;
    LPBYTE buf;
    Tcl_Obj *objP;

    status = NetUseGetInfo(UncServerName, UseName, level, &buf);
    if (status != NERR_Success) {
        return Twapi_AppendSystemError(interp, status);
    }

    objP = ObjFromUSE_INFO(interp, buf, level);
    if (objP)
        ObjSetResult(interp, objP);

    NetApiBufferFree(buf);
    return objP ? TCL_OK : TCL_ERROR;
}

int Twapi_WNetGetUser(
    Tcl_Interp *interp,
    LPCWSTR  lpName
)
{
    WCHAR buf[256];
    DWORD bufsz = ARRAYSIZE(buf);
    DWORD error;

    error = WNetGetUserW(lpName, buf, &bufsz);
    if (error != NO_ERROR)
        return Twapi_AppendWNetError(interp, error);

    ObjSetResult(interp, ObjFromWinChars(buf));
    return TCL_OK;
}

int Twapi_NetSessionGetInfo(
    Tcl_Interp    *interp,
    LPWSTR server,
    LPWSTR client,
    LPWSTR user,
    DWORD level
    )
{
    Tcl_Obj *resultObj;
    LPBYTE         bufP;
    NET_API_STATUS status;

    status = NetSessionGetInfo(server, client, user, level, &bufP);

    if (status != NERR_Success) {
        return Twapi_AppendSystemError(interp, status);
    }
    
    resultObj = ObjFromSESSION_INFO(interp, bufP, level);
    if (resultObj)
        ObjSetResult(interp, resultObj);

    NetApiBufferFree(bufP);
    return resultObj ? TCL_OK : TCL_ERROR;
}

int Twapi_NetFileGetInfo(
    Tcl_Interp    *interp,      /* Where result is returned */
    LPWSTR server,
    DWORD fileid,
    DWORD level
    )
{
    Tcl_Obj *resultObj;
    LPBYTE         bufP;
    NET_API_STATUS status;

    status = NetFileGetInfo(server, fileid, level, &bufP);

    if (status != NERR_Success) {
        return Twapi_AppendSystemError(interp, status);
    }
    
    resultObj = ObjFromFILE_INFO(interp, bufP, level);
    if (resultObj)
        ObjSetResult(interp, resultObj);

    NetApiBufferFree(bufP);
    return resultObj ? TCL_OK : TCL_ERROR;
}



static int Twapi_ShareCallNetEnumObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD i;
    LPBYTE     p;
    DWORD   dwresume;
    TwapiNetEnumContext netenum;
    int struct_size;
    Tcl_Obj *(*objfn)(Tcl_Interp *, LPBYTE, DWORD);
    Tcl_Obj *objs[4];
    Tcl_Obj *enumObj = NULL;
    Tcl_Obj *sObj;
    int func = PtrToInt(clientdata);

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    sObj = objv[1];
    objc -= 2;
    objv += 2;

    /* WARNING:
       Many of the cases in the switch below cannot be combined even
       though they look similar because of slight variations in the
       Win32 function prototypes they call like const / non-const,
       size of resume handle etc. */

    /* NOTE as always, WSTR parameters must be extracted AFTER 
       scalar parameters to prevent shimmering problems */

    netenum.netbufP = NULL;
    switch (func) {
    case 1: // NetUseEnum system level resumehandle
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (netenum.level) {
        case 0: struct_size = sizeof(USE_INFO_0); break;
        case 1: struct_size = sizeof(USE_INFO_1); break;
        case 2: struct_size = sizeof(USE_INFO_2); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromUSE_INFO;
        netenum.status = NetUseEnum (
            ObjToWinChars(sObj), netenum.level,
            &netenum.netbufP,
            MAX_PREFERRED_LENGTH,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = (DWORD_PTR) dwresume;
        break;

    case 2: // NetShareEnum system level resumehandle
        // Not shared with above code because first param has a const
        // qualifier in above cases which results in warnings if 
        // combined with this case.
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (netenum.level) {
        case 0: struct_size = sizeof(SHARE_INFO_0); break;
        case 1: struct_size = sizeof(SHARE_INFO_1); break;
        case 2: struct_size = sizeof(SHARE_INFO_2); break;
        case 502: struct_size = sizeof(SHARE_INFO_502); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromSHARE_INFO;
        netenum.status = NetShareEnum(ObjToWinChars(sObj), netenum.level,
                                      &netenum.netbufP,
                                      MAX_PREFERRED_LENGTH,
                                      &netenum.entriesread,
                                      &netenum.totalentries,
                                      &dwresume);
        netenum.hresume = (DWORD_PTR) dwresume;
        break;

    case 3:  // NetConnectionEnum server group level resumehandle
        // Not shared with other code because first param has a const
        // qualifier in above cases which results in warnings if 
        // combined with this case.
        if (TwapiGetArgs(interp, objc, objv, ARGSKIP,
                         GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (netenum.level) {
        case 0: struct_size = sizeof(CONNECTION_INFO_0); break;
        case 1: struct_size = sizeof(CONNECTION_INFO_1); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromCONNECTION_INFO;
        netenum.status = NetConnectionEnum (
            ObjToWinChars(sObj),
            ObjToWinChars(objv[0]),
            netenum.level,
            &netenum.netbufP,
            MAX_PREFERRED_LENGTH,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = (DWORD_PTR)dwresume;
        break;

    case 4:
        if (TwapiGetArgs(interp, objc, objv, ARGSKIP, ARGSKIP,
                         GETINT(netenum.level), GETDWORD_PTR(netenum.hresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (netenum.level) {
        case 2: struct_size = sizeof(FILE_INFO_2); break;
        case 3: struct_size = sizeof(FILE_INFO_3); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromFILE_INFO;
        netenum.status = NetFileEnum (
            ObjToWinChars(sObj),
            ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
            ObjToLPWSTR_NULL_IF_EMPTY(objv[1]),
            netenum.level, 
            &netenum.netbufP,
            MAX_PREFERRED_LENGTH,
            &netenum.entriesread,
            &netenum.totalentries,
            &netenum.hresume);
        break;

    case 5:
        if (TwapiGetArgs(interp, objc, objv, ARGSKIP, ARGSKIP,
                         GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        switch (netenum.level) {
        case 0: struct_size = sizeof(SESSION_INFO_0); break;
        case 1: struct_size = sizeof(SESSION_INFO_1); break;
        case 2: struct_size = sizeof(SESSION_INFO_2); break;
        case 10: struct_size = sizeof(SESSION_INFO_10); break;
        case 502: struct_size = sizeof(SESSION_INFO_502); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromSESSION_INFO;
        netenum.status = NetSessionEnum (
            ObjToWinChars(sObj),
            ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
            ObjToLPWSTR_NULL_IF_EMPTY(objv[1]),
            netenum.level, 
            &netenum.netbufP,
            MAX_PREFERRED_LENGTH,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = dwresume;
        break;
    }

    if (netenum.status != NERR_Success && netenum.status != ERROR_MORE_DATA) {
        Twapi_AppendSystemError(interp, netenum.status);
        goto error_return;
    }

    enumObj = ObjEmptyList();
    p = netenum.netbufP;
    for (i = 0; i < netenum.entriesread; ++i, p += struct_size) {
        Tcl_Obj *objP;
        objP = objfn(interp, p, netenum.level);
        if (objP == NULL)
            goto error_return;
        ObjAppendElement(interp, enumObj, objP);
    }

    objs[0] = ObjFromLong(netenum.status == ERROR_MORE_DATA);
    objs[1] = ObjFromDWORD_PTR(netenum.hresume);
    objs[2] = ObjFromLong(netenum.totalentries);
    objs[3] = enumObj;

    ObjSetResult(interp, ObjNewList(4, objs));

    if (netenum.netbufP)
        NetApiBufferFree((LPBYTE) netenum.netbufP);
    return TCL_OK;

error_return:
    if (netenum.netbufP)
        NetApiBufferFree((LPBYTE) netenum.netbufP);
    if (enumObj)
        ObjDecrRefs(enumObj);

    return TCL_ERROR;

invalid_level_error:
    TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid info level.");
    goto error_return;
}

static int Twapi_ShareCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD   dw, dw2, dw3, dw4;
    SECURITY_DESCRIPTOR *secdP;
    TwapiResult result;
    int func = PtrToInt(clientdata);
    LPWSTR s, s2;
    SWSMark mark = NULL;
    NETRESOURCEW *netresP;
    HANDLE h;
    TCL_RESULT res;
    Tcl_Obj *objP;
    NETINFOSTRUCT netinfo;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_WNetUseConnection(interp, objc, objv);
    case 2:
        return Twapi_NetShareAdd(interp, objc, objv);
    case 3:
        CHECK_NARGS(interp, objc, 1);
        return Twapi_WNetGetUser(interp, ObjToWinChars(objv[0]));
    case 4: // NetFileClose
        CHECK_NARGS(interp, objc, 2);
        CHECK_INTEGER_OBJ(interp, dw, objv[1]);
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetFileClose(ObjToLPWSTR_NULL_IF_EMPTY(objv[0]), dw);
        break;
    case 5:
    case 6:
        if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        s = ObjToWinChars(objv[0]);
        if (func == 5) {
            result.type = TRT_EXCEPTION_ON_WNET_ERROR;
            result.value.ival = WNetCancelConnection2W(s, dw, dw2);
        } else {
            NULLIFY_EMPTY(s);
            return Twapi_NetFileGetInfo(interp, s, dw, dw2);
        }
        break;
    case 7: // NetShareSetInfo
        CHECK_NARGS(interp, objc, 5);
        CHECK_INTEGER_OBJ(interp, dw, objv[3]);
        mark = SWSPushMark();
        res = ObjToPSECURITY_DESCRIPTORSWS(interp, objv[4], &secdP);
        /* Note secdP may be NULL even on success */
        if (res == TCL_OK) {
            res = Twapi_NetShareSetInfo(interp,
                                        ObjToWinChars(objv[0]),
                                        ObjToWinChars(objv[1]),
                                        ObjToWinChars(objv[2]),
                                        dw, secdP);
        }
        SWSPopMark(mark);
        return res;

    case 8:
    case 9:
    case 10:
    case 11:
        CHECK_NARGS_RANGE(interp, objc, 2, 3);
        if (objc == 2)
            dw = 0;
        else {
            CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        }
        s = ObjToLPWSTR_NULL_IF_EMPTY(objv[0]);
        s2 = ObjToWinChars(objv[1]);
        switch (func) {
        case 8:
            return Twapi_NetUseGetInfo(interp, s, s2, dw);
        case 9:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetShareDel(s,s2,dw);
            break;
        case 10:
            return Twapi_NetShareGetInfo(interp, s, s2, dw);
        case 11:
            return Twapi_NetShareCheck(interp, s, s2);
        }
        break;
    case 12:
        CHECK_NARGS(interp, objc, 3);
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetSessionDel(
            ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
            ObjToLPWSTR_NULL_IF_EMPTY(objv[1]),
            ObjToLPWSTR_NULL_IF_EMPTY(objv[2]));
        break;
    case 13:
        CHECK_NARGS(interp, objc, 4);
        CHECK_INTEGER_OBJ(interp, dw, objv[3]);
        return Twapi_NetSessionGetInfo(interp,
                                       ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
                                       ObjToWinChars(objv[1]),
                                       ObjToWinChars(objv[2]),
                                       dw);
    case 14: // WNetOpenEnum
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETINT(dw2), GETINT(dw3),
                         ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
        mark = SWSPushMark();
        res = TwapiCStructParse(interp, SWS(), objv[3], CSTRUCT_ALLOW_NULL, &dw4, &netresP);
        if (res == TCL_OK) {
            if (netresP && dw4 != sizeof(NETRESOURCEW))
                res = TwapiReturnError(interp, TWAPI_INVALID_ARGS);
            else {
                result.value.ival = WNetOpenEnumW(dw, dw2, dw3, netresP, &h);
                if (result.value.ival != NO_ERROR) {
                    /* Don't change res. Error will be set by TwapiSetResult */
                    result.type = TRT_EXCEPTION_ON_WNET_ERROR;
                } else {
                    result.value.hval = h;
                    result.type = TRT_HANDLE;
                }
            }
        }
        if (res != TCL_OK) {
            result.type = TRT_TCL_RESULT;
            result.value.ival = res;
        }
        break;

    case 15: // WNetCloseEnum
        CHECK_NARGS(interp, objc, 1);
        res = ObjToHANDLE(interp, objv[0], &h);
        if (res != TCL_OK)
            return res;
        result.value.ival = WNetCloseEnum(h);
        result.type = TRT_EXCEPTION_ON_WNET_ERROR;
        break;

    case 16: // WNetEnumResource
        res = TwapiGetArgs(interp, objc, objv, GETHANDLE(h), GETINT(dw),
                           ARGSKIP, ARGEND);
        if (res != TCL_OK)
            return res;

        mark = SWSPushMark();
        netresP = SWSAlloc(1024, &dw2); /* 16K per MSDN */
        result.value.ival = WNetEnumResourceW(h, &dw, netresP, &dw2);
        switch (result.value.ival) {
        case NO_ERROR:
            result.type = TRT_OBJ;
            result.value.obj = ObjNewList(dw, NULL);
            for (dw3 = 0; dw3 < dw; ++dw3) {
                res = ObjFromCStruct(interp,
                                     &netresP[dw3], sizeof(netresP[dw3]),
                                     objv[2], 0, &objP);
                if (res != TCL_OK) {
                    ObjDecrRefs(result.value.obj);
                    result.value.ival = res;
                    result.type = TRT_TCL_RESULT;
                }
                ObjAppendElement(NULL, result.value.obj, objP);
            }
            break;
        case ERROR_NO_MORE_ITEMS:
            result.type = TRT_EMPTY;
            break;
        default:
            result.type = TRT_EXCEPTION_ON_WNET_ERROR;
            break;
        }

        break;

    case 17: // WNetGetConnection
        CHECK_NARGS(interp, objc, 1);
        mark = SWSPushMark();
        s = SWSAlloc(sizeof(WCHAR)*MAX_PATH, &dw2);
        dw2 /= sizeof(WCHAR);
        result.value.ival = WNetGetConnectionW(ObjToWinChars(objv[0]), s, &dw2);
        if (result.value.ival == NO_ERROR) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromWinChars(s);
        } else
            result.type = TRT_EXCEPTION_ON_WNET_ERROR;

        break;

    case 18: // WNetGetProviderName
        CHECK_NARGS(interp, objc, 1);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        mark = SWSPushMark();
        dw2 = MAX_PATH;
        s = SWSAlloc(sizeof(WCHAR)*dw2, NULL);
        result.value.ival = WNetGetProviderNameW(dw, s, &dw2);
        if (result.value.ival == NO_ERROR) {
            result.type = TRT_UNICODE;
            result.value.unicode.str = s;
            result.value.unicode.len = -1;
        } else
            result.type = TRT_EXCEPTION_ON_WNET_ERROR;
        break;

    case 19: // WnetGetNetworkInformation
        CHECK_NARGS(interp, objc, 2);
        netinfo.cbStructure = sizeof(netinfo);
        result.value.ival = WNetGetNetworkInformationW(ObjToWinChars(objv[0]),
                                                       &netinfo);
        if (result.value.ival != NO_ERROR)
            result.type = TRT_EXCEPTION_ON_WNET_ERROR;
        else {
            result.value.ival = ObjFromCStruct(interp, &netinfo, sizeof(netinfo),
                                 objv[1], 0, &objP);
            if (result.value.ival != TCL_OK)
                result.type = TRT_TCL_RESULT;
            else {
                result.value.obj = objP;
                result.type = TRT_OBJ;
            }
        }
        break;
    }
    res = TwapiSetResult(interp, &result);
    /* Clear memlifo AFTER setting result */
    if (mark)
        SWSPopMark(mark);
    return res;
}

static int TwapiShareInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s ShareEnumDispatch[] = {
        DEFINE_FNCODE_CMD(NetUseEnum, 1),
        DEFINE_FNCODE_CMD(NetShareEnum, 2),
        DEFINE_FNCODE_CMD(NetConnectionEnum, 3),
        DEFINE_FNCODE_CMD(NetFileEnum, 4),
        DEFINE_FNCODE_CMD(NetSessionEnum, 5),
    };

    static struct fncode_dispatch_s ShareDispatch[] = {
        DEFINE_FNCODE_CMD(Twapi_WNetUseConnection, 1),
        DEFINE_FNCODE_CMD(NetShareAdd, 2),
        DEFINE_FNCODE_CMD(WNetGetUser, 3),
        DEFINE_FNCODE_CMD(NetFileClose, 4),
        DEFINE_FNCODE_CMD(WNetCancelConnection2, 5),
        DEFINE_FNCODE_CMD(NetFileGetInfo, 6),
        DEFINE_FNCODE_CMD(NetShareSetInfo, 7),
        DEFINE_FNCODE_CMD(NetUseGetInfo, 8),
        DEFINE_FNCODE_CMD(NetShareDel, 9),
        DEFINE_FNCODE_CMD(NetShareGetInfo, 10),
        DEFINE_FNCODE_CMD(Twapi_NetShareCheck, 11),
        DEFINE_FNCODE_CMD(NetSessionDel, 12),
        DEFINE_FNCODE_CMD(NetSessionGetInfo, 13),
        DEFINE_FNCODE_CMD(WNetOpenEnum, 14),
        DEFINE_FNCODE_CMD(WNetCloseEnum, 15),
        DEFINE_FNCODE_CMD(WNetEnumResource, 16),
        DEFINE_FNCODE_CMD(WNetGetConnection, 17),
        DEFINE_FNCODE_CMD(WNetGetProviderName, 18),
        DEFINE_FNCODE_CMD(WNetGetNetworkInformation, 19),
    };

    static struct tcl_dispatch_s ShareCmdDispatch[] = {
        DEFINE_TCL_CMD(Twapi_WNetGetResourceInformation, Twapi_WNetGetResourceInformationObjCmd),
        DEFINE_TCL_CMD(WNetGetUniversalName, Twapi_WNetGetUniversalNameObjCmd),
        DEFINE_TCL_CMD(WNetAddConnection3, Twapi_WNetAddConnection3ObjCmd),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(ShareEnumDispatch), ShareEnumDispatch, Twapi_ShareCallNetEnumObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(ShareDispatch), ShareDispatch, Twapi_ShareCallObjCmd);

    TwapiDefineTclCmds(interp, ARRAYSIZE(ShareCmdDispatch), ShareCmdDispatch, ticP);

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
int Twapi_share_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiShareInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

