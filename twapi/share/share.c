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

int Twapi_WNetGetUniversalName(TwapiInterpContext *ticP, LPCWSTR localpathP);
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
int Twapi_WNetGetResourceInformation(TwapiInterpContext *ticP,
                                     LPWSTR remoteName, LPWSTR provider,
                                     DWORD  resourcetype);
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
    Tcl_Obj    *objv[14];

    ciP = (CONNECTION_INFO_1 *) infoP; /* May actually be CONNECTION_INFO_0 */
    /* We build from back of array since we would like basic elements
       at front of Tcl list we build */
    objc = sizeof(objv)/sizeof(objv[0]);
    switch (level) {
    case 1:
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_type);
        objv[--objc] = STRING_LITERAL_OBJ("type");
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_num_opens);
        objv[--objc] = STRING_LITERAL_OBJ("num_opens");
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_time);
        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_num_users);
        objv[--objc] = STRING_LITERAL_OBJ("num_users");
        objv[--objc] = ObjFromUnicode(ciP->coni1_username);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = ObjFromUnicode(ciP->coni1_netname);
        objv[--objc] = STRING_LITERAL_OBJ("netname");
        /* FALLTHRU */
    case 0:
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_id);
        objv[--objc] = STRING_LITERAL_OBJ("id");
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
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
    Tcl_Obj    *objv[18];

    uiP = (USE_INFO_2 *) infoP; /* May acutally be any USE_INFO_* */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build (most accessed elements in front */
    objc = sizeof(objv)/sizeof(objv[0]);
#define ADD_LPWSTR_(fld) do {                                      \
        objv[--objc] = ObjFromUnicode(uiP->ui2_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[--objc] = Tcl_NewLongObj(uiP->ui2_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)

    switch (level) {
    case 2:
        ADD_LPWSTR_(username);
        ADD_LPWSTR_(domainname);
        // FALLTHRU
    case 1:
        ADD_LPWSTR_(password);
        ADD_DWORD_(status);
        ADD_DWORD_(asg_type);
        ADD_DWORD_(refcount);
        ADD_DWORD_(usecount);
        //FALLTHRU
    case 0:
        ADD_LPWSTR_(local);
        ADD_LPWSTR_(remote);
        break;

    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);

}

Tcl_Obj *ObjFromSHARE_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    SHARE_INFO_502 *siP;
    int         objc;
    Tcl_Obj    *objv[18];

    siP = (SHARE_INFO_502 *) infoP; /* May acutally be any SHARE_INFO_* */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build (most accessed elements in front */
    objc = sizeof(objv)/sizeof(objv[0]);
#define ADD_LPWSTR_(fld) do {                                          \
        objv[--objc] = ObjFromUnicode(siP->shi502_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[--objc] = Tcl_NewLongObj(siP->shi502_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)

    switch (level) {
    case 502:
        objv[--objc] = ObjFromSECURITY_DESCRIPTOR(interp,
                                                      siP->shi502_security_descriptor);
        objv[--objc] = STRING_LITERAL_OBJ("security_descriptor");
        // FALLTHRU
    case 2:
        ADD_LPWSTR_(passwd);
        ADD_LPWSTR_(path);
        ADD_DWORD_(current_uses);
        ADD_DWORD_(max_uses);
        ADD_DWORD_(permissions);
        // FALLTHRU
    case 1:
        ADD_LPWSTR_(remark);
        ADD_DWORD_(type);
        // FALLTHRU
    case 0:
        ADD_LPWSTR_(netname);
        break;

    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
}


Tcl_Obj *ObjFromFILE_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    int         objc;
    Tcl_Obj    *objv[10];
    FILE_INFO_3 *fiP;

    fiP = (FILE_INFO_3 *) infoP; /* May acutally be FILE_INFO_* */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build */
    objc = sizeof(objv)/sizeof(objv[0]);
    switch (level) {
    case 3:
        objv[--objc] = Tcl_NewLongObj(fiP->fi3_permissions);
        objv[--objc] = STRING_LITERAL_OBJ("permissions");
        objv[--objc] = Tcl_NewLongObj(fiP->fi3_num_locks);
        objv[--objc] = STRING_LITERAL_OBJ("num_locks");
        objv[--objc] = ObjFromUnicode(fiP->fi3_username);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = ObjFromUnicode(fiP->fi3_pathname);
        objv[--objc] = STRING_LITERAL_OBJ("pathname");
        /* FALLTHRU */
    case 2:
        objv[--objc] = Tcl_NewLongObj(fiP->fi3_id);
        objv[--objc] = STRING_LITERAL_OBJ("id");
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
}


Tcl_Obj *ObjFromSESSION_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    int         objc;
    Tcl_Obj    *objv[16];

    /* Note in the code below, the structures can be superimposed on one
       another except level 10
    */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build */
    objc = sizeof(objv)/sizeof(objv[0]);
    switch (level) {
    case 502:
        objv[--objc] = ObjFromUnicode(((SESSION_INFO_502 *)infoP)->sesi502_transport);
        objv[--objc] = STRING_LITERAL_OBJ("transport");
        /* FALLTHRU */
    case 2:
        objv[--objc] = ObjFromUnicode(((SESSION_INFO_2 *)infoP)->sesi2_cltype_name);
        objv[--objc] = STRING_LITERAL_OBJ("cltype_name");
        /* FALLTHRU */
    case 1:
        objv[--objc] = ObjFromUnicode(((SESSION_INFO_1 *)infoP)->sesi1_username);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_num_opens);
        objv[--objc] = STRING_LITERAL_OBJ("num_opens");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_time);

        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_idle_time);

        objv[--objc] = STRING_LITERAL_OBJ("idle_time");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_user_flags);

        objv[--objc] = STRING_LITERAL_OBJ("user_flags");
        /* FALLTHRU */
    case 0:
        objv[--objc] = ObjFromUnicode(((SESSION_INFO_0 *)infoP)->sesi0_cname);
        objv[--objc] = STRING_LITERAL_OBJ("cname");
        break;

    case 10:
        objv[--objc] = ObjFromUnicode(((SESSION_INFO_10 *)infoP)->sesi10_cname);
        objv[--objc] = STRING_LITERAL_OBJ("cname");
        objv[--objc] = ObjFromUnicode(((SESSION_INFO_10 *)infoP)->sesi10_username);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_10 *)infoP)->sesi10_time);

        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_10 *)infoP)->sesi10_idle_time);

        objv[--objc] = STRING_LITERAL_OBJ("idle_time");
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
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
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, nrP, dwScope);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, nrP, dwType);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, nrP, dwDisplayType);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, nrP, dwUsage);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, resultObj, nrP, lpLocalName);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, resultObj, nrP, lpRemoteName);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, resultObj, nrP, lpComment);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, resultObj, nrP, lpProvider);

    return resultObj;
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

    objv[0] = ObjFromUnicode(rniP->lpUniversalName);
    objv[1] = ObjFromUnicode(rniP->lpConnectionName);
    objv[2] = ObjFromUnicode(rniP->lpRemainingPath);

    return Tcl_NewListObj(3, objv);
}

int Twapi_NetShareAdd(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR server_name;
    SECURITY_DESCRIPTOR *secdP;
    SHARE_INFO_502 share_info;
    DWORD          parm_err;
    NET_API_STATUS status;

    if (TwapiGetArgs(interp, objc, objv,
                     GETNULLIFEMPTY(server_name),
                     GETWSTR(share_info.shi502_netname),
                     GETINT(share_info.shi502_type),
                     GETWSTR(share_info.shi502_remark),
                     GETINT(share_info.shi502_max_uses),
                     GETWSTR(share_info.shi502_path),
                     GETVAR(secdP, ObjToPSECURITY_DESCRIPTOR),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    share_info.shi502_permissions = 0;
    share_info.shi502_current_uses = 0;
    share_info.shi502_passwd       = NULL;
    share_info.shi502_reserved     = 0;
    share_info.shi502_security_descriptor = secdP;
    
    status = NetShareAdd(server_name, 502, (unsigned char *)&share_info, &parm_err);
    if (secdP) TwapiFreeSECURITY_DESCRIPTOR(secdP);

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
        objv[0] = Tcl_NewIntObj(1);
        objv[1] = Tcl_NewIntObj(type);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
    }
    else if (status == NERR_DeviceNotShared) {
        objv[0] = Tcl_NewIntObj(0);
        objv[1] = Tcl_NewIntObj(0);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
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
            Tcl_SetResult(interp,
                          "Could not retrieve share information: ", TCL_STATIC);
            return Twapi_AppendSystemError(interp, status);
        }
        Tcl_SetObjResult(
            interp,
            ObjFromSHARE_INFO(interp, shareP, level)
            );
        NetApiBufferFree(shareP);
        break;

    default:
        Tcl_SetResult(interp, "Invalid or unsupported share information level specified", TCL_STATIC);
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
        Tcl_SetResult(interp,
                      "Could not retrieve share information: ", TCL_STATIC);
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
    LPWSTR  passwordP;
    DWORD                 flags;

    WCHAR accessname[MAX_PATH];
    DWORD accessname_size;
    NETRESOURCEW netresource;
    DWORD       outflags;       /* Not really used */
    int         error;

    if (TwapiGetArgs(interp, objc, objv,
                     GETDWORD_PTR(winH), GETINT(netresource.dwType),
                     GETNULLIFEMPTY(netresource.lpLocalName),
                     GETWSTR(netresource.lpRemoteName),
                     GETNULLIFEMPTY(netresource.lpProvider),
                     GETNULLIFEMPTY(usernameP),
                     GETINT(ignore_password), GETNULLIFEMPTY(passwordP),
                     GETINT(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (ignore_password) {
        passwordP = L"";
    }

    accessname_size = ARRAYSIZE(accessname);
    error = WNetUseConnectionW(winH, &netresource, passwordP, usernameP, flags,
                               accessname, &accessname_size, &outflags);
    if (error == NO_ERROR) {
        Tcl_SetObjResult(interp, ObjFromUnicode(accessname));
        return TCL_OK;
    }
    else {
        return Twapi_AppendWNetError(interp, error);
    }
}

int Twapi_WNetGetUniversalName (
    TwapiInterpContext *ticP, 
    LPCWSTR      localpathP
)
{
    int result;
    DWORD error;
    DWORD buf_sz;
    void *buf;

    buf = MemLifoPushFrame(&ticP->memlifo, MAX_PATH+1, &buf_sz);
    error = WNetGetUniversalNameW(localpathP, REMOTE_NAME_INFO_LEVEL,
                                  buf, &buf_sz);
    if (error = ERROR_MORE_DATA) {
        /* Retry with larger buffer */
        buf = MemLifoAlloc(&ticP->memlifo, buf_sz, NULL);
        error = WNetGetUniversalNameW(localpathP, REMOTE_NAME_INFO_LEVEL,
                                      buf, &buf_sz);
    }
    if (error == NO_ERROR) {
        Tcl_SetObjResult(ticP->interp,
                         ListObjFromREMOTE_NAME_INFOW(ticP->interp,
                                                      (REMOTE_NAME_INFOW *) buf));
        result = TCL_OK;
    } 
    else {
        Twapi_AppendWNetError(ticP->interp, error);
        result = TCL_ERROR;
    }

    MemLifoPopFrame(&ticP->memlifo);

    return result;
}



int Twapi_WNetGetResourceInformation(
    TwapiInterpContext *ticP,
    LPWSTR remoteName,
    LPWSTR provider,
    DWORD   resourcetype
    )
{
    NETRESOURCEW in;
    NETRESOURCEW *outP;
    DWORD       outsz;
    int         error;
    LPWSTR      systempart;
    Tcl_Obj    *objv[2];

    in.dwType = resourcetype;
    in.lpRemoteName = remoteName;
    in.lpProvider = provider;
    in.dwType = resourcetype;

    outP = MemLifoPushFrame(&ticP->memlifo, 4000, &outsz);
    error = WNetGetResourceInformationW(&in, outP, &outsz, &systempart);
    if (error == ERROR_MORE_DATA) {
        /* Retry with larger buffer */
        outP = MemLifoAlloc(&ticP->memlifo, outsz, NULL);
        error = WNetGetResourceInformationW(&in, outP, &outsz, &systempart);
    }
    if (error == ERROR_SUCCESS) {
        objv[0] = ListObjFromNETRESOURCEW(ticP->interp, outP);
        objv[1] = ObjFromUnicode(systempart);
        Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(2, objv));
    } else
        Twapi_AppendWNetError(ticP->interp, error);

    MemLifoPopFrame(&ticP->memlifo);

    return error == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
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
        Tcl_SetObjResult(interp, objP);

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

    Tcl_SetObjResult(interp, ObjFromUnicode(buf));
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
        Tcl_SetObjResult(interp, resultObj);

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
        Tcl_SetObjResult(interp, resultObj);

    NetApiBufferFree(bufP);
    return resultObj ? TCL_OK : TCL_ERROR;
}



static int Twapi_ShareCallNetEnumObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    DWORD i;
    LPBYTE     p;
    LPWSTR s1, s2, s3;
    DWORD   dwresume;
    TwapiNetEnumContext netenum;
    int struct_size;
    Tcl_Obj *(*objfn)(Tcl_Interp *, LPBYTE, DWORD);
    Tcl_Obj *objs[4];
    Tcl_Obj *enumObj = NULL;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETNULLIFEMPTY(s1),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    /* WARNING:
       Many of the cases in the switch below cannot be combined even
       though they look similar because of slight variations in the
       Win32 function prototypes they call like const / non-const,
       size of resume handle etc. */

    netenum.netbufP = NULL;
    switch (func) {
    case 1: // NetUseEnum system level resumehandle
        if (TwapiGetArgs(interp, objc-3, objv+3,
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
            s1, netenum.level,
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
        if (TwapiGetArgs(interp, objc-3, objv+3,
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
        netenum.status = NetShareEnum(s1, netenum.level,
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
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (netenum.level) {
        case 0: struct_size = sizeof(CONNECTION_INFO_0); break;
        case 1: struct_size = sizeof(CONNECTION_INFO_1); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromCONNECTION_INFO;
        netenum.status = NetConnectionEnum (
            s1,
            Tcl_GetUnicode(objv[3]),
            netenum.level,
            &netenum.netbufP,
            MAX_PREFERRED_LENGTH,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = (DWORD_PTR)dwresume;
        break;

    case 4:
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETNULLIFEMPTY(s2), GETNULLIFEMPTY(s3),
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
            s1, s2, s3, netenum.level, 
            &netenum.netbufP,
            MAX_PREFERRED_LENGTH,
            &netenum.entriesread,
            &netenum.totalentries,
            &netenum.hresume);
        break;

    case 5:
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETNULLIFEMPTY(s2), GETNULLIFEMPTY(s3),
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
            s1, s2, s3, netenum.level, 
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

    enumObj = Tcl_NewListObj(0, NULL);
    p = netenum.netbufP;
    for (i = 0; i < netenum.entriesread; ++i, p += struct_size) {
        Tcl_Obj *objP;
        objP = objfn(interp, p, netenum.level);
        if (objP == NULL)
            goto error_return;
        Tcl_ListObjAppendElement(interp, enumObj, objP);
    }

    objs[0] = Tcl_NewIntObj(netenum.status == ERROR_MORE_DATA);
    objs[1] = ObjFromDWORD_PTR(netenum.hresume);
    objs[2] = Tcl_NewLongObj(netenum.totalentries);
    objs[3] = enumObj;

    Tcl_SetObjResult(interp, Tcl_NewListObj(4, objs));

    if (netenum.netbufP)
        NetApiBufferFree((LPBYTE) netenum.netbufP);
    return TCL_OK;

error_return:
    if (netenum.netbufP)
        NetApiBufferFree((LPBYTE) netenum.netbufP);
    if (enumObj)
        Tcl_DecrRefCount(enumObj);

    return TCL_ERROR;

invalid_level_error:
    TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid info level.");
    goto error_return;
}


static int Twapi_ShareCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s1, s2, s3;
    DWORD   dw, dw2;
    SECURITY_DESCRIPTOR *secdP;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETWSTR(s1), ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_WNetUseConnection(interp, objc-2, objv+2);
    case 2:
        return Twapi_NetShareAdd(interp, objc-2, objv+2);
    case 3:
        return Twapi_WNetGetUser(interp, s1);
    case 4: // NetFileClose
        if (TwapiGetArgs(interp, objc-3, objv+3, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s1);
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetFileClose(s1, dw);
        break;
    case 5:
    case 6:
        if (TwapiGetArgs(interp, objc-3, objv+3, GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (func == 5) {
            result.type = TRT_EXCEPTION_ON_WNET_ERROR;
            result.value.ival = WNetCancelConnection2W(s1, dw, dw2);
        } else {
            NULLIFY_EMPTY(s1);
            return Twapi_NetFileGetInfo(interp, s1, dw, dw2);
        }
        break;
    case 7: // NetShareSetInfo
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETWSTR(s3), GETINT(dw), ARGSKIP,
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (ObjToPSECURITY_DESCRIPTOR(interp, objv[6], &secdP) != TCL_OK)
            return TCL_ERROR;
        /* Note secdP may be NULL even on success */
        result.value.ival = Twapi_NetShareSetInfo(interp, s1, s2, s3, dw, secdP);
        if (secdP)
            TwapiFreeSECURITY_DESCRIPTOR(secdP);
        return result.value.ival;

    case 8:
    case 9:
    case 10:
    case 11:
    case 12:
        if (TwapiGetArgs(interp, objc-3, objv+3, GETWSTR(s2),
                         ARGUSEDEFAULT, GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s1);
        switch (func) {
        case 8:
            return Twapi_NetUseGetInfo(interp, s1, s2, dw);
        case 9:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetShareDel(s1,s2,dw);
            break;
        case 10:
            return Twapi_NetShareGetInfo(interp, s1,s2,dw);
        case 11:
            /* TBD - test when s1 is empty (NULL) */
            NULLIFY_EMPTY(s2);
            return Twapi_WNetGetResourceInformation(ticP, s1, s2, dw);
        case 12:
            return Twapi_NetShareCheck(interp, s1, s2);
        }
        break;
    case 13:
        if (TwapiGetArgs(interp, objc-3, objv+3, GETWSTR(s2),
                         GETWSTR(s3), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s1);
        return Twapi_NetSessionGetInfo(interp, s1,s2,s3,dw);
    case 14:
        if (TwapiGetArgs(interp, objc-3, objv+3, GETNULLIFEMPTY(s2),
                         GETNULLIFEMPTY(s3), ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s1);
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetSessionDel(s1,s2,s3);
        break;
    case 15:
        return Twapi_WNetGetUniversalName(ticP, s1);
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_ShareInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::ShareCall", Twapi_ShareCallObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::ShareCallNetEnum", Twapi_ShareCallNetEnumObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::Share" #call_, # code_); \
    } while (0);

    CALL_(Twapi_WNetUseConnection, Call, 1);
    CALL_(NetShareAdd, Call, 2);
    CALL_(WNetGetUser, Call, 3);
    CALL_(NetFileClose, Call, 4);
    CALL_(WNetCancelConnection2, Call, 5);
    CALL_(NetFileGetInfo, Call, 6);
    CALL_(NetShareSetInfo, Call, 7);
    CALL_(NetUseGetInfo, Call, 8);
    CALL_(NetShareDel, Call, 9);
    CALL_(NetShareGetInfo, Call, 10);
    CALL_(Twapi_WNetGetResourceInformation, Call, 11);
    CALL_(Twapi_NetShareCheck, Call, 12);
    CALL_(NetSessionGetInfo, Call, 13);
    CALL_(NetSessionDel, Call, 14);
    CALL_(WNetGetUniversalName, Call, 15);

    CALL_(NetUseEnum, CallNetEnum, 1);
    CALL_(NetShareEnum, CallNetEnum, 2);
    CALL_(NetConnectionEnum, CallNetEnum, 3);
    CALL_(NetFileEnum, CallNetEnum, 4);
    CALL_(NetSessionEnum, CallNetEnum, 5);



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
int Twapi_share_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, WLITERAL(MODULENAME), MODULE_HANDLE,
                            Twapi_ShareInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

