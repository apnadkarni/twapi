/* 
 * Copyright (c) 2003-2009, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

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

    objv[0] = Tcl_NewUnicodeObj(rniP->lpUniversalName, -1);
    objv[1] = Tcl_NewUnicodeObj(rniP->lpConnectionName, -1);
    objv[2] = Tcl_NewUnicodeObj(rniP->lpRemainingPath, -1);

    return Tcl_NewListObj(3, objv);
}


#ifdef OBSOLETE // Use MSTASK instead
/*
 * Convert AT_INFO to Tcl list
 */

static Tcl_Obj *ListObjFromAT_INFO (const AT_INFO *atP)
{
    Tcl_Obj *objv[5];
    objv[0] = TWAPI_NEW_DWORD_PTR_OBJ(atP->JobTime);
    objv[1] = Tcl_NewIntObj(atP->DaysOfMonth);
    objv[2] = Tcl_NewIntObj(atP->DaysOfWeek);
    objv[3] = Tcl_NewIntObj(atP->Flags);
    objv[4] = Tcl_NewUnicodeObj(atP->Command, -1);
    return Tcl_NewListObj(5, objv);
}

static Twapi_freeAT_INFO (AT_INFO *atP) 
{
    if (atP)
        free(atP);
}


/* interp may be NULL. Returned memory must be freed with Twapi_freeAT_INFO */
static AT_INFO *ListObjToAT_INFO (
    Tcl_Interp *interp,
    Tcl_Obj *listobj
)
{
    AT_INFO  *atP;
    int       objc;
    Tcl_Obj **objv;
    WCHAR    *cmdP;
    int       cmd_len;
    int       days;
    int       flags;
    
    if (Tcl_ListObjGetElements(interp, listobj, &objc, &objv) != TCL_OK)
        return NULL;

    if (objc != 5) {
        if (interp)
            Tcl_SetResult(interp, "AT_INFO list must have exactly 5 elements", TCL_STATIC);
        return NULL;
    }
        
    cmdP = Tcl_GetUnicodeFromObj(objv[4], &cmd_len);
    
    if (Twapi_malloc(interp, NULL, sizeof(AT_INFO) + sizeof(WCHAR)*(cmd_len+1), &atP) != TCL_OK)
        return NULL;

    if (TWAPI_DWORD_PTR_FROM_OBJ(interp, objv[0], &atP->JobTime) != TCL_OK ||
        Tcl_GetLongFromObj(interp, objv[1], &atP->DaysOfMonth) != TCL_OK ||
        Tcl_GetLongFromObj(interp, objv[2], &days) != TCL_OK ||
        Tcl_GetLongFromObj(interp, objv[3], &flags) != TCL_OK) {
        Twapi_freeAT_INFO(atP);
        return NULL;
    }
    atP->DaysOfWeek = days;
    atP->Flags = flags;

    wcscpy((wchar_t *)(sizeof(AT_INFO)+(char *)atP), cmdP);
    return atP;
}


static Tcl_Obj *ListObjFromAT_ENUM(const AT_ENUM *atP)
{
    Tcl_Obj *objv[6];
    objv[0] = Tcl_NewIntObj(atP->JobId);
    objv[1] = TWAPI_NEW_DWORD_PTR_OBJ(atP->JobTime);
    objv[2] = Tcl_NewIntObj(atP->DaysOfMonth);
    objv[3] = Tcl_NewIntObj(atP->DaysOfWeek);
    objv[4] = Tcl_NewIntObj(atP->Flags);
    objv[5] = Tcl_NewUnicodeObj(atP->Command, -1);
    return Tcl_NewListObj(6, objv);
}
#endif // OBSOLETE


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
        Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(accessname, -1));
        return TCL_OK;
    }
    else {
        return Twapi_AppendWNetError(interp, error);
    }
}

int Twapi_WNetGetUniversalName (
    Tcl_Interp *interp, 
    LPCWSTR      localpathP
)
{
    int result;
    DWORD error;
    DWORD buf_sz = 2 * MAX_PATH;
    void *buf;

    if (Twapi_malloc(interp, NULL, buf_sz, &buf) != TCL_OK)
        return TCL_ERROR;
    error = WNetGetUniversalNameW(localpathP, REMOTE_NAME_INFO_LEVEL,
                                  buf, &buf_sz);
    if (error == NO_ERROR) {
        Tcl_SetObjResult(interp,
                         ListObjFromREMOTE_NAME_INFOW(interp,
                                                      (REMOTE_NAME_INFOW *) buf));
        result = TCL_OK;
    } 
    else {
        Twapi_AppendWNetError(interp, error);
        result = TCL_ERROR;
    }

    free(buf);

    return result;
}



int Twapi_WNetGetResourceInformation(
    Tcl_Interp *interp,
    LPWSTR remoteName,
    LPWSTR provider,
    DWORD   resourcetype
    )
{
    NETRESOURCEW in;
    char        buf[4096];
    NETRESOURCEW *outP = (NETRESOURCEW *)buf;
    DWORD       outsz = sizeof(buf);
    int         error;
    LPWSTR      systempart;
    Tcl_Obj    *objv[2];

    in.dwType = resourcetype;
    in.lpRemoteName = remoteName;
    in.lpProvider = provider;
    in.dwType = resourcetype;

    error = WNetGetResourceInformationW(&in, outP, &outsz, &systempart);
    if (error != NO_ERROR) {
        if (error != ERROR_MORE_DATA)
            goto error_return;

        /* We need a bigger buffer */
        if (Twapi_malloc(interp, NULL, outsz, &outP) != TCL_OK)
            return TCL_ERROR;

        /* Retry with larger buffer */
        error = WNetGetResourceInformationW(&in, outP, &outsz, &systempart);
        if (error != NO_ERROR)
            goto error_return;
    }

    objv[0] = ListObjFromNETRESOURCEW(interp, outP);
    objv[1] = Tcl_NewUnicodeObj(systempart, -1);
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));

    if ((char *)outP != buf)
        free(outP);
    return TCL_OK;

 error_return:
    Twapi_AppendWNetError(interp, error);
    if ((char *)outP != buf)
        free(outP);
    return TCL_ERROR;
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

    Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(buf, -1));
    return TCL_OK;
}

int Twapi_NetGetDCName(Tcl_Interp *interp, LPCWSTR servername, LPCWSTR domainname)
{
    NET_API_STATUS status;
    LPBYTE         bufP;
    status = NetGetDCName(servername, domainname, &bufP);
    if (status != NERR_Success) {
        return Twapi_AppendSystemError(interp, status);
    }
    Tcl_SetObjResult(interp, Tcl_NewUnicodeObj((wchar_t *)bufP, -1));
    NetApiBufferFree(bufP);
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

