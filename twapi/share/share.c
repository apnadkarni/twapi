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
    Tcl_Obj    *objv[14];

    ciP = (CONNECTION_INFO_1 *) infoP; /* May actually be CONNECTION_INFO_0 */
    /* We build from back of array since we would like basic elements
       at front of Tcl list we build */
    objc = sizeof(objv)/sizeof(objv[0]);
    switch (level) {
    case 1:
        objv[--objc] = ObjFromLong(ciP->coni1_type);
        objv[--objc] = STRING_LITERAL_OBJ("type");
        objv[--objc] = ObjFromLong(ciP->coni1_num_opens);
        objv[--objc] = STRING_LITERAL_OBJ("num_opens");
        objv[--objc] = ObjFromLong(ciP->coni1_time);
        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = ObjFromLong(ciP->coni1_num_users);
        objv[--objc] = STRING_LITERAL_OBJ("num_users");
        objv[--objc] = ObjFromUnicode(ciP->coni1_username);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = ObjFromUnicode(ciP->coni1_netname);
        objv[--objc] = STRING_LITERAL_OBJ("netname");
        /* FALLTHRU */
    case 0:
        objv[--objc] = ObjFromLong(ciP->coni1_id);
        objv[--objc] = STRING_LITERAL_OBJ("id");
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

    return ObjNewList((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
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
        objv[--objc] = ObjFromLong(uiP->ui2_ ## fld); \
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

    return ObjNewList((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);

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
        objv[--objc] = ObjFromLong(siP->shi502_ ## fld); \
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

    return ObjNewList((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
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
        objv[--objc] = ObjFromLong(fiP->fi3_permissions);
        objv[--objc] = STRING_LITERAL_OBJ("permissions");
        objv[--objc] = ObjFromLong(fiP->fi3_num_locks);
        objv[--objc] = STRING_LITERAL_OBJ("num_locks");
        objv[--objc] = ObjFromUnicode(fiP->fi3_username);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = ObjFromUnicode(fiP->fi3_pathname);
        objv[--objc] = STRING_LITERAL_OBJ("pathname");
        /* FALLTHRU */
    case 2:
        objv[--objc] = ObjFromLong(fiP->fi3_id);
        objv[--objc] = STRING_LITERAL_OBJ("id");
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

    return ObjNewList((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
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
        objv[--objc] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_num_opens);
        objv[--objc] = STRING_LITERAL_OBJ("num_opens");
        objv[--objc] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_time);

        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_idle_time);

        objv[--objc] = STRING_LITERAL_OBJ("idle_time");
        objv[--objc] = ObjFromLong(((SESSION_INFO_1 *)infoP)->sesi1_user_flags);

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
        objv[--objc] = ObjFromLong(((SESSION_INFO_10 *)infoP)->sesi10_time);

        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = ObjFromLong(((SESSION_INFO_10 *)infoP)->sesi10_idle_time);

        objv[--objc] = STRING_LITERAL_OBJ("idle_time");
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", level));
        return NULL;
    }

    return ObjNewList((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
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
    Tcl_Obj *resultObj = ObjEmptyList();
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

    return ObjNewList(3, objv);
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
        objv[0] = ObjFromLong(1);
        objv[1] = ObjFromDWORD(type);
        TwapiSetObjResult(interp, ObjNewList(2, objv));
    }
    else if (status == NERR_DeviceNotShared) {
        objv[0] = ObjFromLong(0);
        objv[1] = ObjFromLong(0);
        TwapiSetObjResult(interp, ObjNewList(2, objv));
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
            TwapiSetStaticResult(interp, "Could not retrieve share information: ");
            return Twapi_AppendSystemError(interp, status);
        }
        TwapiSetObjResult(
            interp,
            ObjFromSHARE_INFO(interp, shareP, level)
            );
        NetApiBufferFree(shareP);
        break;

    default:
        TwapiSetStaticResult(interp, "Invalid or unsupported share information level specified");
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
        TwapiSetStaticResult(interp, "Could not retrieve share information: ");
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
        TwapiSetObjResult(interp, ObjFromUnicode(accessname));
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

    localpathP = ObjToUnicode(objv[1]);

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
        TwapiSetObjResult(ticP->interp,
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

static TCL_RESULT Twapi_WNetGetResourceInformationObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR remoteName;
    LPWSTR provider;
    DWORD   resourcetype;

    NETRESOURCEW in;
    NETRESOURCEW *outP;
    DWORD       outsz;
    int         error;
    LPWSTR      systempart;
    Tcl_Obj    *objs[2];

    /* TBD - check if GETNULLIFEMPTY is ok */
    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETNULLIFEMPTY(remoteName),
                     GETNULLIFEMPTY(provider),
                     GETINT(resourcetype),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

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
        objs[0] = ListObjFromNETRESOURCEW(ticP->interp, outP);
        objs[1] = ObjFromUnicode(systempart);
        TwapiSetObjResult(ticP->interp, ObjNewList(2, objs));
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
        TwapiSetObjResult(interp, objP);

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

    TwapiSetObjResult(interp, ObjFromUnicode(buf));
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
        TwapiSetObjResult(interp, resultObj);

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
        TwapiSetObjResult(interp, resultObj);

    NetApiBufferFree(bufP);
    return resultObj ? TCL_OK : TCL_ERROR;
}



static int Twapi_ShareCallNetEnumObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD i;
    LPBYTE     p;
    LPWSTR s1, s2, s3;
    DWORD   dwresume;
    TwapiNetEnumContext netenum;
    int struct_size;
    Tcl_Obj *(*objfn)(Tcl_Interp *, LPBYTE, DWORD);
    Tcl_Obj *objs[4];
    Tcl_Obj *enumObj = NULL;
    int func = (int) clientdata;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    s1 = ObjToLPWSTR_NULL_IF_EMPTY(objv[1]);

    objc -= 2;
    objv += 2;

    /* WARNING:
       Many of the cases in the switch below cannot be combined even
       though they look similar because of slight variations in the
       Win32 function prototypes they call like const / non-const,
       size of resume handle etc. */

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
        if (TwapiGetArgs(interp, objc, objv,
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
            ObjToUnicode(objv[0]),
            netenum.level,
            &netenum.netbufP,
            MAX_PREFERRED_LENGTH,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = (DWORD_PTR)dwresume;
        break;

    case 4:
        if (TwapiGetArgs(interp, objc, objv,
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
        if (TwapiGetArgs(interp, objc, objv,
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

    enumObj = ObjEmptyList();
    p = netenum.netbufP;
    for (i = 0; i < netenum.entriesread; ++i, p += struct_size) {
        Tcl_Obj *objP;
        objP = objfn(interp, p, netenum.level);
        if (objP == NULL)
            goto error_return;
        Tcl_ListObjAppendElement(interp, enumObj, objP);
    }

    objs[0] = ObjFromLong(netenum.status == ERROR_MORE_DATA);
    objs[1] = ObjFromDWORD_PTR(netenum.hresume);
    objs[2] = ObjFromLong(netenum.totalentries);
    objs[3] = enumObj;

    TwapiSetObjResult(interp, ObjNewList(4, objs));

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

static int Twapi_ShareCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR s1, s2, s3;
    DWORD   dw, dw2;
    SECURITY_DESCRIPTOR *secdP;
    TwapiResult result;
    int func = (int) clientdata;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETWSTR(s1), ARGTERM) != TCL_OK)
        return TCL_ERROR;

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_WNetUseConnection(interp, objc, objv);
    case 2:
        return Twapi_NetShareAdd(interp, objc, objv);
    case 3:
        return Twapi_WNetGetUser(interp, s1);
    case 4: // NetFileClose
        if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s1);
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetFileClose(s1, dw);
        break;
    case 5:
    case 6:
        if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(dw), GETINT(dw2),
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
        if (TwapiGetArgs(interp, objc-1, objv+1,
                         GETWSTR(s2), GETWSTR(s3), GETINT(dw), ARGSKIP,
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (ObjToPSECURITY_DESCRIPTOR(interp, objv[4], &secdP) != TCL_OK)
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
        if (TwapiGetArgs(interp, objc-1, objv+1, GETWSTR(s2),
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
            return Twapi_NetShareCheck(interp, s1, s2);
        }
        break;
    case 12:
        if (TwapiGetArgs(interp, objc-1, objv+1, GETNULLIFEMPTY(s2),
                         GETNULLIFEMPTY(s3), ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s1);
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetSessionDel(s1,s2,s3);
        break;
    case 13:
        if (TwapiGetArgs(interp, objc-1, objv+1, GETWSTR(s2),
                         GETWSTR(s3), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s1);
        return Twapi_NetSessionGetInfo(interp, s1,s2,s3,dw);
    }

    return TwapiSetResult(interp, &result);
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
    };

    static struct tcl_dispatch_s ShareCmdDispatch[] = {
        DEFINE_TCL_CMD(Twapi_WNetGetResourceInformation, Twapi_WNetGetResourceInformationObjCmd),
        DEFINE_TCL_CMD(WNetGetUniversalName, Twapi_WNetGetUniversalNameObjCmd),
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

