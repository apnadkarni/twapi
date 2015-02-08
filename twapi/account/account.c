/* 
 * Copyright (c) 2012 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to user accounts */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

int Twapi_NetUserEnum(Tcl_Interp *interp, LPWSTR server_name, DWORD filter);
int Twapi_NetGroupEnum(Tcl_Interp *interp, LPWSTR server_name);
int Twapi_NetLocalGroupEnum(Tcl_Interp *interp, LPWSTR server_name);
int Twapi_NetUserGetGroups(Tcl_Interp *interp, LPWSTR server, LPWSTR user);
int Twapi_NetUserGetLocalGroups(Tcl_Interp *interp, LPWSTR server,
                                LPWSTR user, DWORD flags);
int Twapi_NetLocalGroupGetMembers(Tcl_Interp *interp, LPWSTR server, LPWSTR group);
int Twapi_NetGroupGetUsers(Tcl_Interp *interp, LPCWSTR server, LPCWSTR group);
int Twapi_NetUserGetInfo(Tcl_Interp *interp, LPCWSTR server,
                         LPCWSTR user, DWORD level);
int Twapi_NetGroupGetInfo(Tcl_Interp *interp, LPCWSTR server,
                          LPCWSTR group, DWORD level);
int Twapi_NetLocalGroupGetInfo(Tcl_Interp *interp, LPCWSTR server,
                               LPCWSTR group, DWORD level);
int Twapi_NetUserSetInfoDWORD(int fun, LPCWSTR server, LPCWSTR user, DWORD dw);
int Twapi_NetUserSetInfoLPWSTR(int fun, LPCWSTR server, LPCWSTR user, LPWSTR s);

/*
 * Control how large buffers passed to NetEnum* functions be. This
 * is managed at script level through twapi::Twapi_SetNetEnumBufSize
 * and is used primarily for testing purposes to test cases where 
 * we want to check correct operation when all data does not get returned
 * in a single buffer.
 */
static int g_netenum_buf_size = MAX_PREFERRED_LENGTH;


/*
 * Convert local info structure to a list. Returns TCL_OK/TCL_ERROR
 * interp may be NULL
 */
Tcl_Obj *ObjFromLOCALGROUP_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    LOCALGROUP_INFO_1 *groupinfoP = (LOCALGROUP_INFO_1 *) infoP;
    Tcl_Obj    *objs[2];
    int nobjs = 1;

    switch (info_level) {
    case 1:
        ++nobjs;
        objs[1] = ObjFromUnicode(groupinfoP->lgrpi1_comment);
        /* FALL THRU */
    case 0:
        objs[0] = ObjFromUnicode(groupinfoP->lgrpi1_name);
        break;
    default:
        Twapi_WrongLevelError(interp, info_level);
        return NULL;
    }

    return ObjNewList(nobjs, objs);
}

/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromGROUP_USERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    GROUP_USERS_INFO_1 *groupinfoP = (GROUP_USERS_INFO_1 *) infoP;
    int         objc = 1;
    Tcl_Obj    *objv[2];

    switch (info_level) {
    case 1:
        ++objc;
        objv[1] = ObjFromLong(groupinfoP->grui1_attributes);
        /* FALLTHRU */
    case 0:
        objv[0] = ObjFromUnicode(groupinfoP->grui1_name);
        break;
    default:
        Twapi_WrongLevelError(interp, info_level);
        return NULL;
    }

    return ObjNewList(objc, objv);
}

/*
 * Convert user info structure to a list. Returns NULL on error;
 */
Tcl_Obj *ObjFromUSER_INFO(
    Tcl_Interp *interp,
    LPBYTE     infoP,
    DWORD        info_level
    )
{
    Tcl_Obj    *objs[29];
    int nobjs;
    /* Define our own because older SDK's do not have this definition */
    struct _TWAPI_USER_INFO_24 {
        BOOL   usri24_internet_identity;
        DWORD  usri24_flags;
        LPWSTR usri24_internet_provider_name;
        LPWSTR usri24_internet_principal_name;
        PSID   usri24_user_sid;
    } *usri24P;

    /*
     * Note userinfoP may not point to a USER_INFO_3 struct! It depends
     * on what is specified as info_level. However, the initial fields
     * of the USER_INFO_{0,1,2,3} are the same allowing us to
     * use the following switch statement.
     */

    nobjs = 1;                  /* name field always present */
    switch (info_level) {
    case 24:
        usri24P = (struct _TWAPI_USER_INFO_24 *)infoP;
        if (! usri24P->usri24_internet_identity)
            return ObjFromEmptyString();
        objs[0] = ObjFromDWORD(usri24P->usri24_flags);
        objs[1] = ObjFromUnicode(usri24P->usri24_internet_provider_name);
        objs[2] = ObjFromUnicode(usri24P->usri24_internet_principal_name);
        objs[3] = ObjFromSIDNoFail(usri24P->usri24_user_sid);
        nobjs = 4;
        break;

    case 3:
    case 4:
        nobjs += 5;
        /* NOTE even when fields names are the same, level 3 and
           level 4 HAVE DIFFERENT FIELD OFFSETS */
        if (info_level == 3) {
            objs[24] = ObjFromDWORD(((USER_INFO_3*)infoP)->usri3_user_id);
            objs[25] = ObjFromDWORD(((USER_INFO_3*)infoP)->usri3_primary_group_id);
            objs[26] = ObjFromUnicode(((USER_INFO_3*)infoP)->usri3_profile);
            objs[27] = ObjFromUnicode(((USER_INFO_3*)infoP)->usri3_home_dir_drive);
            objs[28] = ObjFromDWORD(((USER_INFO_3*)infoP)->usri3_password_expired);
        } else {
            objs[24] = ObjFromSIDNoFail(((USER_INFO_4*)infoP)->usri4_user_sid);
            objs[25] = ObjFromDWORD(((USER_INFO_4*)infoP)->usri4_primary_group_id);
            objs[26] = ObjFromUnicode(((USER_INFO_4*)infoP)->usri4_profile);
            objs[27] = ObjFromUnicode(((USER_INFO_4*)infoP)->usri4_home_dir_drive);
            objs[28] = ObjFromDWORD(((USER_INFO_4*)infoP)->usri4_password_expired);
        }
        /* FALL THROUGH */
    case 2:
        nobjs += 16;
        objs[8] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_auth_flags);
        objs[9] = ObjFromUnicode(((USER_INFO_2*)infoP)->usri2_full_name);
        objs[10] = ObjFromUnicode(((USER_INFO_2*)infoP)->usri2_usr_comment);
        objs[11] = ObjFromUnicode(((USER_INFO_2*)infoP)->usri2_parms);
        objs[12] = ObjFromUnicode(((USER_INFO_2*)infoP)->usri2_workstations);
        objs[13] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_last_logon);
        objs[14] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_last_logoff);
        objs[15] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_acct_expires);
        if (((USER_INFO_2*)infoP)->usri2_max_storage == (DWORD) -1)
            objs[16] = ObjFromLong(-1); /* Do not want UINT_MAX */
        else 
            objs[16] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_max_storage);
        objs[17] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_units_per_week);
        objs[18] = ObjFromByteArray(((USER_INFO_2*)infoP)->usri2_logon_hours,21);
        objs[19] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_bad_pw_count);
        objs[20] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_num_logons);
        objs[21] = ObjFromUnicode(((USER_INFO_2*)infoP)->usri2_logon_server);
        objs[22] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_country_code);
        objs[23] = ObjFromDWORD(((USER_INFO_2*)infoP)->usri2_code_page);
        /* FALL THROUGH */
    case 1:
        nobjs += 7;
        objs[1] = ObjFromUnicode(((USER_INFO_1*)infoP)->usri1_password ? ((USER_INFO_1*)infoP)->usri1_password : L"");
        objs[2] = ObjFromDWORD(((USER_INFO_1*)infoP)->usri1_password_age);
        objs[3] = ObjFromDWORD(((USER_INFO_1*)infoP)->usri1_priv);
        objs[4] = ObjFromUnicode(((USER_INFO_1*)infoP)->usri1_home_dir);
        objs[5] = ObjFromUnicode(((USER_INFO_1*)infoP)->usri1_comment);
        objs[6] = ObjFromDWORD(((USER_INFO_1*)infoP)->usri1_flags);
        objs[7] = ObjFromUnicode(((USER_INFO_1*)infoP)->usri1_script_path);
        /* FALL THROUGH */
    case 0:
        objs[0] = ObjFromUnicode(((USER_INFO_0*)infoP)->usri0_name);
        break;
    default:
        Twapi_WrongLevelError(interp, info_level);
        return NULL;
    }

    return ObjNewList(nobjs, objs);
}


/*
 * Convert group info structure to a list. Returns TCL_OK/TCL_ERROR
 * interp may be NULL
 */
Tcl_Obj *ObjFromGROUP_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    int         nobjs;
    Tcl_Obj    *objs[4];

    /*
     * Note groupinfoP may not point to a GROUP_INFO_3 struct! It depends
     * on what is specified as info_level. However, the initial fields
     * of the GROUP_INFO_{0,1,2,3} are the same allowing us to
     * use the following switch statement.
     */

    nobjs = 1;
    switch (info_level) {
    case 3: /* FALL THROUGH */
    case 2:
        nobjs += 2;
        if (info_level == 2) {
            objs[3] = ObjFromDWORD(((GROUP_INFO_2*)infoP)->grpi2_attributes);
            objs[2] = ObjFromDWORD(((GROUP_INFO_2*)infoP)->grpi2_group_id);
        } else {
            objs[3] = ObjFromDWORD(((GROUP_INFO_3*)infoP)->grpi3_attributes);
            objs[2] = ObjFromSIDNoFail(((GROUP_INFO_3*)infoP)->grpi3_group_sid);
        }
        /* FALL THROUGH */
    case 1:
        nobjs += 1;
        objs[1] = ObjFromUnicode(((GROUP_INFO_1*)infoP)->grpi1_comment);
        /* FALL THROUGH */
    case 0:
        objs[0] = ObjFromUnicode(((GROUP_INFO_0*)infoP)->grpi0_name);
        break;
    default:
        Twapi_WrongLevelError(interp, info_level);
        return NULL;
    }

    return ObjNewList(nobjs, objs);
}

/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromLOCALGROUP_USERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    Tcl_Obj    *objs[1];

    /* Even though only one field, for consistency with other
       structures, we have to return as a list */

    switch (info_level) {
    case 0:
        objs[0] = ObjFromUnicode(((LOCALGROUP_USERS_INFO_0 *)infoP)->lgrui0_name);
        break;
    default:
        Twapi_WrongLevelError(interp, info_level);
        return NULL;
    }

    return ObjNewList(1, objs);
}

/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromLOCALGROUP_MEMBERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    int         objc = 0;
    Tcl_Obj    *objv[3];

    switch (info_level) {
    case 0:
        objv[objc++] = ObjFromSIDNoFail(((LOCALGROUP_MEMBERS_INFO_0 *)infoP)->lgrmi0_sid);
        break;
    case 1:
    case 2:
        objv[objc++] = ObjFromSIDNoFail(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_sid);
        objv[objc++] = ObjFromLong(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_sidusage);
        if (info_level == 1) {
            objv[objc++] = ObjFromUnicode(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_name);
        } else {
            objv[objc++] = ObjFromUnicode(((LOCALGROUP_MEMBERS_INFO_2 *)infoP)->lgrmi2_domainandname);
        }
        break;
    case 3:
        objv[objc++] = ObjFromUnicode(((LOCALGROUP_MEMBERS_INFO_3 *)infoP)->lgrmi3_domainandname);
        break;
        
    default:
        Twapi_WrongLevelError(interp, info_level);
        return NULL;
    }

    TWAPI_ASSERT(objc <= ARRAYSIZE(objv));

    return ObjNewList(objc, objv);
}

int TwapiNetUserOrGroupGetInfoHelper(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     name,
    DWORD       level,
    DWORD       type /* 0 - user, 1 - global group, 2 - local group */
    )
{
    NET_API_STATUS status;
    LPBYTE         infoP;
    typedef NET_API_STATUS NET_API_FUNCTION NETGETINFO(LPCWSTR, LPCWSTR, DWORD, LPBYTE*);
    NETGETINFO    *get_fn;
    Tcl_Obj *      (*fmt_fn)(Tcl_Interp *, LPBYTE, DWORD);

    switch (level) {
    case 0:
    case 1:
    case 2:
    case 3:
    case 4:
    case 24:
        switch (type) {
        case 0:
            get_fn = NetUserGetInfo;
            fmt_fn = ObjFromUSER_INFO;
            break;
        case 1:
            get_fn = NetGroupGetInfo;
            fmt_fn = ObjFromGROUP_INFO;
            break;
        case 2:
            get_fn = NetLocalGroupGetInfo;
            fmt_fn = ObjFromLOCALGROUP_INFO;
            break;
        default:
            ObjSetStaticResult(interp, "Internal error: bad type passed to TwapiNetUserOrGroupGetInfoHelper");
            return TCL_ERROR;
        }
        status = (*get_fn)(servername, name, level, &infoP);
        if (status != NERR_Success) {
            ObjSetStaticResult(interp,
                                 "Could not retrieve global user or group information: ");
            return Twapi_AppendSystemError(interp, status);
        }
        ObjSetResult(interp, (*fmt_fn)(interp, infoP, level));
        NetApiBufferFree(infoP);
        break;

    default:
        ObjSetStaticResult(interp, "Invalid or unsupported user or group information level specified");
        return TCL_ERROR;
    }

    return TCL_OK;
}

/* Wrapper around NetUserGetInfo */
int Twapi_NetUserGetInfo(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     username,
    DWORD       level
    )
{
    return TwapiNetUserOrGroupGetInfoHelper(interp, servername, username, level, 0);
}


int Twapi_NetGroupGetInfo(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     groupname,
    DWORD       level
    )
{
    return TwapiNetUserOrGroupGetInfoHelper(interp, servername, groupname, level, 1);
}

int Twapi_NetLocalGroupGetInfo(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     groupname,
    DWORD       level
    )
{
    if (level != 1) {
        ObjSetStaticResult(interp, "Invalid or unsupported user or group information level specified");
        return TCL_ERROR;
    }
    return TwapiNetUserOrGroupGetInfoHelper(interp, servername, groupname, level, 2);
}

static int Twapi_NetUserAddObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD       priv, flags;
    NET_API_STATUS status;
    USER_INFO_1    userinfo;
    DWORD          error_parm;
    char          *error_field;
    WCHAR         *decryptedP;
    int            decrypted_len;

    CHECK_NARGS(interp, objc, 9);

    /* As always to avoid shimmering problems, extract integer object first */
    CHECK_INTEGER_OBJ(interp, priv, objv[4]);
    CHECK_INTEGER_OBJ(interp, flags, objv[7]);

    userinfo.usri1_name        = ObjToUnicode(objv[2]);

    /* Now get the decrypted password object */
    decryptedP = ObjDecryptPassword(objv[3], &decrypted_len);
    userinfo.usri1_password    = decryptedP;
    userinfo.usri1_password_age = 0;
    userinfo.usri1_priv        = priv;
    userinfo.usri1_home_dir    = ObjToLPWSTR_NULL_IF_EMPTY(objv[5]);
    userinfo.usri1_comment     = ObjToLPWSTR_NULL_IF_EMPTY(objv[6]);
    userinfo.usri1_flags       = flags;
    userinfo.usri1_script_path = ObjToLPWSTR_NULL_IF_EMPTY(objv[8]);

    status = NetUserAdd(ObjToLPWSTR_NULL_IF_EMPTY(objv[1]), 1, (LPBYTE) &userinfo, &error_parm);

    TwapiFreeDecryptedPassword(decryptedP, decrypted_len);

    if (status == NERR_Success)
        return TCL_OK;

    /* Indicate the parameter */
    switch (error_parm) {
    case 0: error_field = "user name "; break;
    case 1: error_field = "password ";  break;
    case 3: error_field = "privilege level ";  break;
    case 4: error_field = "home directory ";  break;
    case 5: error_field = "comment ";  break;
    case 6: error_field = "flags ";  break;
    case 7: error_field = "script path ";  break;
    default: error_field = ""; break;
    }
    ObjSetResult(interp, Tcl_ObjPrintf("Error adding user account: invalid %sfied.", error_field));
    return Twapi_AppendSystemError(interp, status);
}

int Twapi_NetUserSetInfoDWORD(
    int func,
    LPCWSTR servername,
    LPCWSTR username,
    DWORD   dw
    )
{
    union {
        USER_INFO_1005 u1005;
        USER_INFO_1008 u1008;
        USER_INFO_1010 u1010;
        USER_INFO_1017 u1017;
        USER_INFO_1024 u1024;
    } userinfo;

    switch (func) {
    case 1005:
        userinfo.u1005.usri1005_priv = dw;
        break;
    case 1008:
        userinfo.u1008.usri1008_flags = dw;
        break;
    case 1010:
        userinfo.u1010.usri1010_auth_flags = dw;
        break;
    case 1017:
        userinfo.u1017.usri1017_acct_expires = dw;
        break;
    case 1024:
        userinfo.u1024.usri1024_country_code = dw;
        break;
    default:
        return ERROR_INVALID_PARAMETER;
    }
    return NetUserSetInfo(servername, username, func, (LPBYTE) &userinfo, NULL);
}

int Twapi_NetUserSetInfoLPWSTR(
    int func,
    LPCWSTR servername,
    LPCWSTR username,
    LPWSTR value
    )
{
    union {
        USER_INFO_0 u0;
        USER_INFO_1003 u1003;
        USER_INFO_1006 u1006;
        USER_INFO_1007 u1007;
        USER_INFO_1009 u1009;
        USER_INFO_1011 u1011;
        USER_INFO_1052 u1052;
        USER_INFO_1053 u1053;
    } userinfo;

    switch (func) {
    case 0:
        userinfo.u0.usri0_name = value;
        break;
    case 1003:
        userinfo.u1003.usri1003_password = value;
        break;
    case 1006:
        userinfo.u1006.usri1006_home_dir = value;
        break;
    case 1007:
        userinfo.u1007.usri1007_comment = value;
        break;
    case 1009:
        userinfo.u1009.usri1009_script_path = value;
        break;
    case 1011:
        userinfo.u1011.usri1011_full_name = value;
        break;
    case 1052:
        userinfo.u1052.usri1052_profile = value;
        break;
    case 1053:
        userinfo.u1053.usri1053_home_dir_drive = value;
        break;
    default:
        return ERROR_INVALID_PARAMETER;
    }
    return NetUserSetInfo(servername, username, func, (LPBYTE) &userinfo, NULL);
}

int Twapi_NetUserSetInfoObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s1, s3;
    DWORD   dw;
    TwapiResult result;
    Tcl_Obj *s1Obj, *s2Obj;
    WCHAR *passwordP = NULL;
    int    password_len;

    /* Note: to prevent shimmering issues, we do not extract the internal
       string pointers s1 and s2 until integer args have been parsed */

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETOBJ(s1Obj), 
                     GETOBJ(s2Obj), ARGSKIP,
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;

    switch (func) {
    case 1005: // This block of function codes maps directly to
    case 1008: // the function codes accepted by Twapi_NetUserSetInfoDWORD.
    case 1010: // A bit klugy but the easiest way to keep backwards
    case 1017: // compatibility.
    case 1024: 
        CHECK_INTEGER_OBJ(interp, dw, objv[4]);
        s1 = ObjToUnicode(s1Obj);
        if (*s1 == 0)
            s1 = NULL;
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = Twapi_NetUserSetInfoDWORD(func, s1, ObjToUnicode(s2Obj), dw);
        break;
        
    case 0:    // See note above except that this maps to 
    case 1003: // Twapi_NetUserSetInfoLPWSTR instead of the
    case 1006: // Twapi_NetUserSetInfoDWORD functions
    case 1007:
    case 1009:
    case 1011:
    case 1052:
    case 1053:
        if (func == 1003) {
            s3 = ObjDecryptPassword(objv[4], &password_len);
        } else
            s3 = ObjToUnicode(objv[4]);
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = Twapi_NetUserSetInfoLPWSTR(func, s1,
                                                       ObjToUnicode(s2Obj),
                                                       s3);
        if (func == 1003)
            TwapiFreeDecryptedPassword(s3, password_len);
        break;
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_NetUserModalsGetObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TCL_RESULT res;
    MemLifo *memlifoP = ticP->memlifoP;
    MemLifoMarkHandle mark;
    DWORD level, sz;
    Tcl_Obj *structObj;
    LPWSTR server_name;
    LPBYTE pv;
    NET_API_STATUS status;

    mark = MemLifoPushMark(memlifoP);
    res = TwapiGetArgsEx(ticP, objc-1, objv+1, GETEMPTYASNULL(server_name),
                         GETINT(level), GETOBJ(structObj), ARGEND);
    if (res == TCL_OK) {
        switch (level) {
        case 0: sz = sizeof(USER_MODALS_INFO_0); break;
        case 1: sz = sizeof(USER_MODALS_INFO_1); break;
        case 2: sz = sizeof(USER_MODALS_INFO_2); break;
        case 3: sz = sizeof(USER_MODALS_INFO_3); break;
        default:
            res = TwapiReturnError(interp, TWAPI_INVALID_ARGS);
        }
        if (res == TCL_OK) {
            status = NetUserModalsGet(server_name, level, &pv);
            if (status == NERR_Success) {
                res = ObjFromCStruct(interp, pv, sz, structObj, CSTRUCT_RETURN_DICT, NULL);
                NetApiBufferFree(pv);
            } else
                res = Twapi_AppendSystemError(interp, status);
        }
    }

    MemLifoPopMark(mark);
    return res;
}

static int Twapi_NetUserModalsSetObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TCL_RESULT res;
    MemLifo *memlifoP = ticP->memlifoP;
    MemLifoMarkHandle mark;
    DWORD level, sz, actual_sz;
    Tcl_Obj *objP;
    LPWSTR server_name;
    void *pv;
    NET_API_STATUS status;

    mark = MemLifoPushMark(memlifoP);
    res = TwapiGetArgsEx(ticP, objc-1, objv+1, GETEMPTYASNULL(server_name),
                         GETINT(level), GETOBJ(objP), ARGEND);
    if (res == TCL_OK) {
        switch (level) {
        case 0: sz = sizeof(USER_MODALS_INFO_0); break;
        case 1: sz = sizeof(USER_MODALS_INFO_1); break;
        case 2: sz = sizeof(USER_MODALS_INFO_2); break;
        case 3: sz = sizeof(USER_MODALS_INFO_3); break;
        case 1001: sz = sizeof(USER_MODALS_INFO_1001); break;
        case 1002: sz = sizeof(USER_MODALS_INFO_1002); break;
        case 1003: sz = sizeof(USER_MODALS_INFO_1003); break;
        case 1004: sz = sizeof(USER_MODALS_INFO_1004); break;
        case 1005: sz = sizeof(USER_MODALS_INFO_1005); break;
        case 1006: sz = sizeof(USER_MODALS_INFO_1006); break;
        case 1007: sz = sizeof(USER_MODALS_INFO_1007); break;
        default:
            res = TwapiReturnError(interp, TWAPI_INVALID_ARGS);
        }
        if (res == TCL_OK) {
            res = TwapiCStructParse(interp, memlifoP, objP, 0, &actual_sz, &pv);
            if (res == TCL_OK) {
                if (sz != actual_sz)
                    res = TwapiReturnError(interp, TWAPI_INVALID_ARGS);
                else {
                    status = NetUserModalsSet(server_name, level, pv, NULL);
                    if (status != NERR_Success)
                        res = Twapi_AppendSystemError(interp, status);
                }
            }
        }
    }

    MemLifoPopMark(mark);
    return res;
}

static int Twapi_AcctCallNetEnumGetObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD i;
    LPBYTE     p;
    Tcl_Obj *s1Obj, *s2Obj;
    LPWSTR s1;
    DWORD   dw, dwresume;
    TwapiNetEnumContext netenum;
    int struct_size;
    Tcl_Obj *(*objfn)(Tcl_Interp *, LPBYTE, DWORD);
    Tcl_Obj *objs[4];
    Tcl_Obj *enumObj = NULL;
    int func = PtrToInt(clientdata);

    if (TwapiGetArgs(interp, objc-1, objv+1, ARGSKIP, ARGTERM) != TCL_OK)
        return TCL_ERROR;


    /* NOTE : to prevent shimmering issues all wide strings are
       extracted from s1Obj AFTER all other arguments have been extracted
    */

    s1Obj = objv[1];
    objc -= 2;
    objv += 2;
    netenum.netbufP = NULL;

    /* WARNING:
       Many of the cases in the switch below cannot be combined even
       though they look similar because of slight variations in the
       Win32 function prototypes they call like const / non-const,
       size of resume handle etc.
    */

    switch (func) {
    case 3: // NetGroupEnum system level resumehandle
    case 4: // NetLocalGroupEnum system level resumehandle
        // Not shared with case 1: because different const qualifier on function
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(netenum.level), GETDWORD_PTR(netenum.hresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        netenum.status =
            (func == 3 ? NetGroupEnum : NetLocalGroupEnum) (
                s1, netenum.level,
                &netenum.netbufP,
                g_netenum_buf_size,
                &netenum.entriesread,
                &netenum.totalentries,
                &netenum.hresume);
        if (func == 3) {
            objfn = ObjFromGROUP_INFO;
            switch (netenum.level) {
            case 0: struct_size = sizeof(GROUP_INFO_0); break;
            case 1: struct_size = sizeof(GROUP_INFO_1); break;
            case 2: struct_size = sizeof(GROUP_INFO_2); break;
            case 3: struct_size = sizeof(GROUP_INFO_3); break;
            default: goto invalid_level_error;
            }
        } else {
            objfn = ObjFromLOCALGROUP_INFO;
            switch (netenum.level) {
            case 0: struct_size = sizeof(LOCALGROUP_INFO_0); break;
            case 1: struct_size = sizeof(LOCALGROUP_INFO_1); break;
            default: goto invalid_level_error;
            }
        }
        break;

    case 5:
        // NetUserEnum system level filter resumehandle
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(netenum.level), GETINT(dw), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        switch (netenum.level) {
        case 0: struct_size = sizeof(USER_INFO_0); break;
        case 1: struct_size = sizeof(USER_INFO_1); break;
        case 2: struct_size = sizeof(USER_INFO_2); break;
        case 3: struct_size = sizeof(USER_INFO_3); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromUSER_INFO;
        netenum.status = NetUserEnum(s1, netenum.level, dw,
                                       &netenum.netbufP,
                                       g_netenum_buf_size,
                                       &netenum.entriesread,
                                       &netenum.totalentries,
                                       &dwresume);
        netenum.hresume = (DWORD_PTR) dwresume;
        break;

    case 6:  // NetUserGetGroups server user level
        if (TwapiGetArgs(interp, objc, objv,
                         GETOBJ(s2Obj), GETINT(netenum.level),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        switch (netenum.level) {
        case 0: struct_size = sizeof(GROUP_USERS_INFO_0); break;
        case 1: struct_size = sizeof(GROUP_USERS_INFO_1); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromGROUP_USERS_INFO;
        netenum.hresume = 0; /* Not used for these calls */
        netenum.status = NetUserGetGroups(
            s1, ObjToUnicode(s2Obj), netenum.level, &netenum.netbufP,
            g_netenum_buf_size, &netenum.entriesread, &netenum.totalentries);
        break;

    case 7:
        // NetUserGetLocalGroups server user level flags
        if (TwapiGetArgs(interp, objc, objv,
                         GETOBJ(s2Obj), GETINT(netenum.level), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        if (netenum.level != 0)
            goto invalid_level_error;
        struct_size = sizeof(LOCALGROUP_USERS_INFO_0);
        objfn = ObjFromLOCALGROUP_USERS_INFO;
        netenum.hresume = 0; /* Not used for these calls */
        netenum.status = NetUserGetLocalGroups (
            s1, ObjToUnicode(s2Obj),
            netenum.level, dw, &netenum.netbufP, g_netenum_buf_size,
            &netenum.entriesread, &netenum.totalentries);
        break;
    case 8:
    case 9:
        // NetLocalGroupGetMembers server group level resumehandle
        if (TwapiGetArgs(interp, objc, objv,
                         GETOBJ(s2Obj), GETINT(netenum.level), GETDWORD_PTR(netenum.hresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        netenum.status = (func == 8 ? NetLocalGroupGetMembers : NetGroupGetUsers) (
            s1, ObjToUnicode(s2Obj),
            netenum.level, &netenum.netbufP, g_netenum_buf_size,
            &netenum.entriesread, &netenum.totalentries, &netenum.hresume);
        if (func == 8) {
            objfn = ObjFromLOCALGROUP_MEMBERS_INFO;
            switch (netenum.level) {
            case 0: struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_0); break;
            case 1: struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_1); break;
            case 2: struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_2); break;
            case 3: struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_3); break;
            default: goto invalid_level_error;
            }
        } else {            
            objfn = ObjFromGROUP_USERS_INFO;
            switch (netenum.level) {
            case 0: struct_size = sizeof(GROUP_USERS_INFO_0); break;
            case 1: struct_size = sizeof(GROUP_USERS_INFO_1); break;
            default: goto invalid_level_error;
            }
        }
        break;
    }

    if (netenum.status != NERR_Success && netenum.status != ERROR_MORE_DATA) {
        Twapi_AppendSystemError(interp, netenum.status);
        goto error_return;
    }

    enumObj = ObjNewList(0, NULL);
    p = netenum.netbufP;
    for (i = 0; i < netenum.entriesread; ++i, p += struct_size) {
        Tcl_Obj *objP;
        objP = objfn(interp, p, netenum.level);
        if (objP == NULL)
            goto error_return;
        ObjAppendElement(interp, enumObj, objP);
    }

    objs[0] = ObjFromInt(netenum.status == ERROR_MORE_DATA);
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


static TCL_RESULT Twapi_NetLocalGroupMembersObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    int i;
    int func;
    int level;
    Tcl_Obj **accts;
    int naccts;
    LPCWSTR servername, groupname;
    DWORD winerr;
    union {
        LOCALGROUP_MEMBERS_INFO_0 *lgmi0P;
        LOCALGROUP_MEMBERS_INFO_3 *lgmi3P;
        void *bufP;
    } u;
    Tcl_Obj *acctsObj;
    TCL_RESULT res;
    MemLifoMarkHandle mark;

    mark = MemLifoPushMark(ticP->memlifoP);
    naccts = 0;
    res = TwapiGetArgsEx(ticP, objc-1, objv+1, GETINT(func),
                         GETEMPTYASNULL(servername),
                         GETWSTR(groupname), GETINT(level),
                         GETOBJ(acctsObj), ARGEND);
    if (res == TCL_OK)
        res = ObjGetElements(interp, acctsObj, &naccts, &accts);

    if (res != TCL_OK || naccts == 0)
        goto vamoose;

    if (level == 0) {
        u.lgmi0P = MemLifoAlloc(ticP->memlifoP, naccts * sizeof(LOCALGROUP_MEMBERS_INFO_0), NULL);
        for (i = 0; i < naccts; ++i) {
            /* For efficiency reasons we do not use ObjToPSID */
            if (ConvertStringSidToSidW(ObjToUnicode(accts[i]), &u.lgmi0P[0].lgrmi0_sid) == 0) {
                res = TwapiReturnSystemError(interp);
                naccts = i;     /* So right num buffers get freed */
                goto vamoose;
            }
        }
    } else if (level == 3) {
        u.lgmi3P = MemLifoAlloc(ticP->memlifoP, naccts * sizeof(LOCALGROUP_MEMBERS_INFO_3), NULL);
        for (i = 0; i < naccts; ++i) {
            u.lgmi3P[i].lgrmi3_domainandname = ObjToUnicode(accts[i]);
        }
    } else {
        res = TwapiReturnError(interp, TWAPI_INVALID_ARGS);
        goto vamoose;
    }
    
    if (func == 0) {
        winerr = NetLocalGroupAddMembers(servername, groupname, level, u.bufP, naccts);
    } else {
        winerr = NetLocalGroupDelMembers(servername, groupname, level, u.bufP, naccts);
    }
    res = winerr == ERROR_SUCCESS ? TCL_OK : Twapi_AppendSystemError(interp, winerr);

vamoose:
    /* At this point
     * res is TCL_RESULT with interp holding result
     * naccts should be number of valid pointers in u.lgmi0P
     */
    if (level == 0) {
        for (i = 0; i < naccts; ++i) {
            LocalFree(u.lgmi0P[i].lgrmi0_sid);
        }
    }

    MemLifoPopMark(mark);
    return res;
}



static int Twapi_AcctCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    Tcl_Obj *s1Obj, *s2Obj;
    LPWSTR s1, s2, s3;
    DWORD   dw;
    TwapiResult result;
    union {
        WCHAR buf[MAX_PATH+1];
        GROUP_INFO_1 gi1;
        LOCALGROUP_INFO_1 lgi1;
        LOCALGROUP_MEMBERS_INFO_3 lgmi3;
        TwapiNetEnumContext netenum;
        SECURITY_DESCRIPTOR *secdP;
        SECURITY_ATTRIBUTES *secattrP;
    } u;
    int func = PtrToInt(clientdata);

    /* NOTE: To prevent shimmering issues, we all WSTR args must be 
       extracted after other arguments */

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETOBJ(s1Obj),
                     GETOBJ(s2Obj), ARGTERM) != TCL_OK)
        return TCL_ERROR;

    objc -= 3;
    objv += 3;
    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 10:
    case 11:
    case 12:
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        s2 = ObjToUnicode(s2Obj);
        switch (func) {
        case 10:
            return Twapi_NetUserGetInfo(interp, s1, s2, dw);
        case 11:
            return Twapi_NetGroupGetInfo(interp, s1, s2, dw);
        case 12:
            return Twapi_NetLocalGroupGetInfo(interp, s1, s2, dw);
        }
        break;
    case 13:
    case 14:
    case 15:
        if (objc)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        s2 = ObjToUnicode(s2Obj);
        switch (func) {
        case 13:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetGroupDel(s1,s2);
            break;
        case 14:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetLocalGroupDel(s1,s2);
            break;
        case 15: // NetUserDel
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetUserDel(s1, s2);
            break;
        }
        break;
    default:
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        s1 = ObjToLPWSTR_NULL_IF_EMPTY(s1Obj);
        s2 = ObjToUnicode(s2Obj);
        s3 = ObjToUnicode(objv[0]);
        switch (func) {
        case 16:
            u.lgi1.lgrpi1_name = s2;
            NULLIFY_EMPTY(s3);
            u.lgi1.lgrpi1_comment = s3;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetLocalGroupAdd(s1, 1, (LPBYTE)&u.lgi1, NULL);
            break;
        case 17:
            u.gi1.grpi1_name = s2;
            NULLIFY_EMPTY(s3);
            u.gi1.grpi1_comment = s3;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetGroupAdd(s1, 1, (LPBYTE)&u.gi1, NULL);
            break;
        case 18:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetGroupDelUser(s1,s2,s3);
            break;
        case 19:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetGroupAddUser(s1,s2,s3);
            break;
        }
        break;
    }

    return TwapiSetResult(interp, &result);
}

/* Used for testing purposes */
static TCL_RESULT Twapi_SetNetEnumBufSizeObjCmd(
    Tcl_Interp *clientdata,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc > 1) {
        if (ObjToLong(interp, objv[1], &g_netenum_buf_size) != TCL_OK)
            return TCL_ERROR;
        if (g_netenum_buf_size < 0)
            g_netenum_buf_size = MAX_PREFERRED_LENGTH;
    }

    ObjSetResult(interp, Tcl_NewIntObj(g_netenum_buf_size));
    return TCL_OK;
}

static int TwapiAcctInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s AcctCallDispatch[] = {
        DEFINE_FNCODE_CMD(NetUserGetInfo, 10),
        DEFINE_FNCODE_CMD(NetGroupGetInfo, 11),
        DEFINE_FNCODE_CMD(NetLocalGroupGetInfo, 12),
        DEFINE_FNCODE_CMD(NetGroupDel, 13),
        DEFINE_FNCODE_CMD(NetLocalGroupDel, 14),
        DEFINE_FNCODE_CMD(NetUserDel, 15),
        DEFINE_FNCODE_CMD(NetLocalGroupAdd, 16),
        DEFINE_FNCODE_CMD(NetGroupAdd, 17),
        DEFINE_FNCODE_CMD(NetGroupDelUser, 18),
        DEFINE_FNCODE_CMD(NetGroupAddUser, 19),
    };

    static struct fncode_dispatch_s AcctCallNetEnumGetDispatch[] = {
        DEFINE_FNCODE_CMD(NetGroupEnum, 3),
        DEFINE_FNCODE_CMD(NetLocalGroupEnum, 4),
        DEFINE_FNCODE_CMD(NetUserEnum, 5),
        DEFINE_FNCODE_CMD(NetUserGetGroups, 6),
        DEFINE_FNCODE_CMD(NetUserGetLocalGroups, 7),
        DEFINE_FNCODE_CMD(NetLocalGroupGetMembers, 8),
        DEFINE_FNCODE_CMD(NetGroupGetUsers, 9),
    };

    struct tcl_dispatch_s TclDispatch[] = {
        DEFINE_TCL_CMD(NetUserModalsGet, Twapi_NetUserModalsGetObjCmd),
        DEFINE_TCL_CMD(NetUserModalsSet, Twapi_NetUserModalsSetObjCmd),
        DEFINE_TCL_CMD(Twapi_NetUserSetInfo, Twapi_NetUserSetInfoObjCmd),
        DEFINE_TCL_CMD(Twapi_NetLocalGroupMembers, Twapi_NetLocalGroupMembersObjCmd),
        DEFINE_TCL_CMD(NetUserAdd, Twapi_NetUserAddObjCmd),
        DEFINE_TCL_CMD(Twapi_SetNetEnumBufSize, Twapi_SetNetEnumBufSizeObjCmd), /* For testing purposes */
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(AcctCallDispatch), AcctCallDispatch, Twapi_AcctCallObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(AcctCallNetEnumGetDispatch), AcctCallNetEnumGetDispatch, Twapi_AcctCallNetEnumGetObjCmd);
    TwapiDefineTclCmds(interp, ARRAYSIZE(TclDispatch), TclDispatch, ticP);

    /* TBD - write Tcl commands to add / delete multiple members at a time */
    /* TBD - write Tcl commands to add / delete SIDs */

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
int Twapi_account_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiAcctInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

