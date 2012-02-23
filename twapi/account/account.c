/* 
 * Copyright (c) 2012 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to user accounts */

#include "twapi.h"

#ifndef TWAPI_STATIC_BUILD
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
int Twapi_NetUserAdd(Tcl_Interp *interp, LPCWSTR servername, LPWSTR name,
                     LPWSTR password, DWORD priv, LPWSTR home_dir,
                     LPWSTR comment, DWORD flags, LPWSTR script_path);

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
    int         objc = 0;
    Tcl_Obj    *objv[4];


#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = ObjFromUnicode(groupinfoP->lgrpi1_ ## fld); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->lgrpi1_ ## fld); \
    } while (0)

    switch (info_level) {
    case 1:
        ADD_LPWSTR_(comment);
        /* FALLTHRU */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", info_level));
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
}

/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromGROUP_USERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    GROUP_USERS_INFO_1 *groupinfoP = (GROUP_USERS_INFO_1 *) infoP;
    int         objc = 0;
    Tcl_Obj    *objv[4];

#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = ObjFromUnicode(groupinfoP->grui1_ ## fld); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->grui1_ ## fld); \
    } while (0)

    switch (info_level) {
    case 1:
        ADD_DWORD_(attributes);
        /* FALLTHRU */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", info_level));
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
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
    USER_INFO_3 *userinfoP = (USER_INFO_3 *) infoP;
    int         objc;
    Tcl_Obj    *objv[60];


    /*
     * Note userinfoP may not point to a USER_INFO_3 struct! It depends
     * on what is specified as info_level. However, the initial fields
     * of the USER_INFO_{0,1,2,3} are the same allowing us to
     * use the following switch statement.
     */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build (most accessed elements in front */
    objc = sizeof(objv)/sizeof(objv[0]);
#define ADD_LPWSTR_(fld) do {                                          \
        objv[--objc] = ObjFromUnicode(userinfoP->usri3_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[--objc] = Tcl_NewLongObj(userinfoP->usri3_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)

    switch (info_level) {
    case 3:
        ADD_DWORD_(primary_group_id);
        ADD_LPWSTR_(home_dir_drive);
        ADD_DWORD_(password_expired);
        ADD_LPWSTR_(profile);
        ADD_DWORD_(user_id);
        /* FALL THROUGH */
    case 2:
        ADD_DWORD_(auth_flags);
        ADD_LPWSTR_(usr_comment);
        ADD_LPWSTR_(parms);
        ADD_LPWSTR_(workstations);
        ADD_DWORD_(last_logon);
        ADD_DWORD_(last_logoff);
        ADD_DWORD_(acct_expires);
        ADD_DWORD_(max_storage);
        ADD_DWORD_(units_per_week);
        objv[--objc] = Tcl_NewByteArrayObj(userinfoP->usri3_logon_hours,21);
        objv[--objc] = STRING_LITERAL_OBJ("logon_hours");
        ADD_DWORD_(bad_pw_count);
        ADD_DWORD_(num_logons);
        ADD_LPWSTR_(logon_server);
        ADD_DWORD_(country_code);
        ADD_DWORD_(code_page);
        ADD_LPWSTR_(full_name);
        /* FALL THROUGH */
    case 1:
        ADD_LPWSTR_(script_path);
        ADD_LPWSTR_(password);
        ADD_DWORD_(password_age);
        ADD_DWORD_(priv);
        ADD_DWORD_(flags);
        ADD_LPWSTR_(home_dir);
        ADD_LPWSTR_(comment);
        /* FALL THROUGH */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", info_level));
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
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
    GROUP_INFO_3 *groupinfoP = (GROUP_INFO_3 *) infoP;
    int         objc = 0;
    Tcl_Obj    *objv[10];

    /*
     * Note groupinfoP may not point to a GROUP_INFO_3 struct! It depends
     * on what is specified as info_level. However, the initial fields
     * of the GROUP_INFO_{0,1,2,3} are the same allowing us to
     * use the following switch statement.
     */

#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = ObjFromUnicode(groupinfoP->grpi3_ ## fld); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->grpi3_ ## fld); \
    } while (0)

    switch (info_level) {
    case 3: /* FALL THROUGH */
    case 2:
        ADD_DWORD_(attributes);
        if (info_level == 2) {
            objv[objc++] = STRING_LITERAL_OBJ("group_id");
            objv[objc++] = Tcl_NewIntObj(((GROUP_INFO_2 *)groupinfoP)->grpi2_group_id);
        } else {
            objv[objc++] = STRING_LITERAL_OBJ("group_sid");
            objv[objc++] = ObjFromSIDNoFail(groupinfoP->grpi3_group_sid);
        }
        /* FALL THROUGH */
    case 1:
        ADD_LPWSTR_(comment);
        /* FALL THROUGH */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", info_level));
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
}

/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromLOCALGROUP_USERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    LOCALGROUP_USERS_INFO_0 *groupinfoP = (LOCALGROUP_USERS_INFO_0 *) infoP;
    int         objc = 0;
    Tcl_Obj    *objv[2];

#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = ObjFromUnicode(groupinfoP->lgrui0_ ## fld); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->lgrui0_ ## fld); \
    } while (0)

    switch (info_level) {
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", info_level));
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
}

/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromLOCALGROUP_MEMBERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    int         objc = 0;
    Tcl_Obj    *objv[6];

    switch (info_level) {
    case 0:
        objv[objc++] = STRING_LITERAL_OBJ("sid");
        objv[objc++] = ObjFromSIDNoFail(((LOCALGROUP_MEMBERS_INFO_0 *)infoP)->lgrmi0_sid);
        break;
    case 1:
    case 2:
        objv[objc++] = STRING_LITERAL_OBJ("sid");
        objv[objc++] = ObjFromSIDNoFail(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_sid);
        objv[objc++] = STRING_LITERAL_OBJ("sidusage");
        objv[objc++] = Tcl_NewLongObj(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_sidusage);
        if (info_level == 1) {
            objv[objc++] = STRING_LITERAL_OBJ("name");
            objv[objc++] = ObjFromUnicode(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_name);
        } else {
            objv[objc++] = STRING_LITERAL_OBJ("domainandname");
            objv[objc++] = ObjFromUnicode(((LOCALGROUP_MEMBERS_INFO_2 *)infoP)->lgrmi2_domainandname);
        }
        break;
    case 3:
        objv[objc++] = STRING_LITERAL_OBJ("domainandname");
        objv[objc++] = ObjFromUnicode(((LOCALGROUP_MEMBERS_INFO_3 *)infoP)->lgrmi3_domainandname);
        break;
        
    default:
        TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid info level %d.", info_level));
        return NULL;
    }

    return Tcl_NewListObj(objc, objv);
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
            Tcl_SetResult(interp, "Internal error: bad type passed to TwapiNetUserOrGroupGetInfoHelper", TCL_STATIC);
            return TCL_ERROR;
        }
        status = (*get_fn)(servername, name, level, &infoP);
        if (status != NERR_Success) {
            Tcl_SetResult(interp,
                          "Could not retrieve global user or group information: ", TCL_STATIC);
            return Twapi_AppendSystemError(interp, status);
        }
        Tcl_SetObjResult(interp, (*fmt_fn)(interp, infoP, level));
        NetApiBufferFree(infoP);
        break;

    default:
        Tcl_SetResult(interp, "Invalid or unsupported user or group information level specified", TCL_STATIC);
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
        Tcl_SetResult(interp, "Invalid or unsupported user or group information level specified", TCL_STATIC);
        return TCL_ERROR;
    }
    return TwapiNetUserOrGroupGetInfoHelper(interp, servername, groupname, level, 2);
}

#ifndef TWAPI_LEAN
int Twapi_NetUserAdd(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPWSTR     name,
    LPWSTR     password,
    DWORD       priv,
    LPWSTR     home_dir,
    LPWSTR     comment,
    DWORD       flags,
    LPWSTR     script_path
    )
{
    NET_API_STATUS status;
    USER_INFO_1    userinfo;
    DWORD          error_parm;
    char          *error_field;

    userinfo.usri1_name        = name;
    userinfo.usri1_password    = password;
    userinfo.usri1_password_age = 0;
    userinfo.usri1_priv        = priv;
    userinfo.usri1_home_dir    = home_dir;
    userinfo.usri1_comment     = comment;
    userinfo.usri1_flags       = flags;
    userinfo.usri1_script_path = script_path;

    status = NetUserAdd(servername, 1, (LPBYTE) &userinfo, &error_parm);
    if (status == NERR_Success)
        return TCL_OK;

    /* Indicate the parameter */
    Tcl_SetResult(interp, "Error adding user account: ", TCL_STATIC);
    error_field = NULL;
    switch (error_parm) {
    case 0: error_field = "user name"; break;
    case 1: error_field = "password";  break;
    case 3: error_field = "privilege level";  break;
    case 4: error_field = "home directory";  break;
    case 5: error_field = "comment";  break;
    case 6: error_field = "flags";  break;
    case 7: error_field = "script path";  break;
    }
    if (error_field) {
        Tcl_AppendResult(interp, " invalid ", error_field, " field. ", NULL);
    }
    return Twapi_AppendSystemError(interp, status);
}
#endif // TWAPI_LEAN



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
    LPWSTR s1, s2;
    DWORD   dw;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETNULLIFEMPTY(s1), 
                     GETWSTR(s2), ARGSKIP,
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
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = Twapi_NetUserSetInfoDWORD(func, s1, s2, dw);
        break;
        
    case 0:    // See note above except that this maps to 
    case 1003: // Twapi_NetUserSetInfoLPWSTR instead of the
    case 1006: // Twapi_NetUserSetInfoDWORD functions
    case 1007:
    case 1009:
    case 1011:
    case 1052:
    case 1053:
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = Twapi_NetUserSetInfoLPWSTR(func, s1, s2, Tcl_GetUnicode(objv[4]));
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_AcctCallNetEnumGetObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    DWORD i;
    LPBYTE     p;
    LPWSTR s1, s2;
    DWORD   dw, dwresume;
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
    case 3: // NetGroupEnum system level resumehandle
    case 4: // NetLocalGroupEnum system level resumehandle
        // Not shared with case 1: because different const qualifier on function
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(netenum.level), GETDWORD_PTR(netenum.hresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status =
            (func == 3 ? NetGroupEnum : NetLocalGroupEnum) (
                s1, netenum.level,
                &netenum.netbufP,
                MAX_PREFERRED_LENGTH,
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
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(netenum.level), GETINT(dw), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
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
                                       MAX_PREFERRED_LENGTH,
                                       &netenum.entriesread,
                                       &netenum.totalentries,
                                       &dwresume);
        netenum.hresume = (DWORD_PTR) dwresume;
        break;

    case 6:  // NetUserGetGroups server user level
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (netenum.level) {
        case 0: struct_size = sizeof(GROUP_USERS_INFO_0); break;
        case 1: struct_size = sizeof(GROUP_USERS_INFO_1); break;
        default: goto invalid_level_error;
        }
        objfn = ObjFromGROUP_USERS_INFO;
        netenum.hresume = 0; /* Not used for these calls */
        netenum.status = NetUserGetGroups(
            s1, s2, netenum.level, &netenum.netbufP,
            MAX_PREFERRED_LENGTH, &netenum.entriesread, &netenum.totalentries);
        break;

    case 7:
        // NetUserGetLocalGroups server user level flags
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (netenum.level != 0)
            goto invalid_level_error;
        struct_size = sizeof(LOCALGROUP_USERS_INFO_0);
        objfn = ObjFromLOCALGROUP_USERS_INFO;
        netenum.hresume = 0; /* Not used for these calls */
        netenum.status = NetUserGetLocalGroups (
            s1, s2, netenum.level, dw, &netenum.netbufP, MAX_PREFERRED_LENGTH,
            &netenum.entriesread, &netenum.totalentries);
        break;
    case 8:
    case 9:
        // NetLocalGroupGetMembers server group level resumehandle
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level), GETDWORD_PTR(netenum.hresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status = (func == 8 ? NetLocalGroupGetMembers : NetGroupGetUsers) (
            s1, s2, netenum.level, &netenum.netbufP, MAX_PREFERRED_LENGTH,
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


static int Twapi_AcctCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s1, s2, s3, s4, s5, s6;
    DWORD   dw, dw2;
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

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETNULLIFEMPTY(s1),
                     GETWSTR(s2), ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 9: // NetUserAdd
        if (TwapiGetArgs(interp, objc-4, objv+4,
                         GETWSTR(s3), GETINT(dw),
                         GETNULLIFEMPTY(s4), GETNULLIFEMPTY(s5),
                         GETINT(dw2), GETNULLIFEMPTY(s6),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_NetUserAdd(interp, s1, s2, s3, dw, s4, s5, dw2, s6);
    case 10:
    case 11:
    case 12:
        if (TwapiGetArgs(interp, objc-5, objv+5, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
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
    default:
        if (objc != 5)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        s3 = Tcl_GetUnicode(objv[4]);
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
            u.lgmi3.lgrmi3_domainandname = s3;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetLocalGroupDelMembers(s1, s2, 3, (LPBYTE) &u.lgmi3, 1);
            break;
        case 19:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetGroupAddUser(s1,s2,s3);
            break;
        case 20:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetGroupDelUser(s1,s2,s3);
            break;
        case 21:
            u.lgmi3.lgrmi3_domainandname = s3;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetLocalGroupAddMembers(s1, s2, 3, (LPBYTE) &u.lgmi3, 1);
            break;
        }
        break;
    }

    return TwapiSetResult(interp, &result);



}

static int Twapi_AcctInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::AcctCall", Twapi_AcctCallObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::AcctCallNetEnumGet", Twapi_AcctCallNetEnumGetObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::Twapi_NetUserSetInfo", Twapi_NetUserSetInfoObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::Acct" #call_, # code_); \
    } while (0);



    CALL_(NetUserAdd, Call, 9);
    CALL_(NetUserGetInfo, Call, 10);
    CALL_(NetGroupGetInfo, Call, 11);
    CALL_(NetLocalGroupGetInfo, Call, 12);
    CALL_(NetGroupDel, Call, 13);
    CALL_(NetLocalGroupDel, Call, 14);
    CALL_(NetUserDel, Call, 15);
    CALL_(NetLocalGroupAdd, Call, 16);
    CALL_(NetGroupAdd, Call, 17);
    CALL_(Twapi_NetLocalGroupDelMember, Call, 18);
    CALL_(NetGroupAddUser, Call, 19);
    CALL_(NetGroupDelUser, Call, 20);
    CALL_(Twapi_NetLocalGroupAddMember, Call, 21);




    CALL_(NetGroupEnum, CallNetEnumGet, 3);
    CALL_(NetLocalGroupEnum, CallNetEnumGet, 4);
    CALL_(NetUserEnum, CallNetEnumGet, 5);
    CALL_(NetUserGetGroups, CallNetEnumGet, 6);
    CALL_(NetUserGetLocalGroups, CallNetEnumGet, 7);
    CALL_(NetLocalGroupGetMembers, CallNetEnumGet, 8);
    CALL_(NetGroupGetUsers, CallNetEnumGet, 9);


#undef CALL_

    return TCL_OK;
}


#ifndef TWAPI_STATIC_BUILD
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_STATIC_BUILD
__declspec(dllexport) 
#endif
int Twapi_account_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_AcctInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

