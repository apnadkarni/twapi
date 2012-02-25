/* 
 * Copyright (c) 2012 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/*
 * Define interface to Windows API related to misc application stuff
 */

#include "twapi.h"

#ifndef TWAPI_STATIC_BUILD
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

BOOLEAN Twapi_Wow64DisableWow64FsRedirection(LPVOID *oldvalueP);
BOOLEAN Twapi_Wow64RevertWow64FsRedirection(LPVOID addr);
BOOLEAN Twapi_Wow64EnableWow64FsRedirection(BOOLEAN enable_redirection);
int Twapi_LoadUserProfile(Tcl_Interp *interp, HANDLE hToken, DWORD flags,
                          LPWSTR username, LPWSTR profilepath);
int Twapi_GetPrivateProfileSection(TwapiInterpContext *ticP,
                                   LPCWSTR app, LPCWSTR fn);
int Twapi_GetPrivateProfileSectionNames(TwapiInterpContext *,LPCWSTR filename);

static int TwapiGetProfileSectionHelper(
    TwapiInterpContext *ticP,
    LPCWSTR lpAppName, /* If NULL, section names are retrieved */
    LPCWSTR lpFileName /* If NULL, win.ini is used */
    )
{
    WCHAR *bufP;
    DWORD  bufsz;
    DWORD  numchars;

    bufP = MemLifoPushFrame(&ticP->memlifo, 1000, &bufsz);
    while (1) {
        DWORD bufchars = bufsz/sizeof(WCHAR);
        if (lpAppName) {
            if (lpFileName)
                numchars = GetPrivateProfileSectionW(lpAppName,
                                                     bufP, bufchars,
                                                     lpFileName);
            else
                numchars = GetProfileSectionW(lpAppName, bufP, bufchars);
        } else {
            /* Get section names. Note lpFileName can be NULL */
            numchars = GetPrivateProfileSectionNamesW(bufP, bufchars, lpFileName);
        }

        if (numchars >= (bufchars-2)) {
            /* Buffer not big enough */
            MemLifoPopFrame(&ticP->memlifo);
            bufsz = 2*bufsz;
            bufP = MemLifoPushFrame(&ticP->memlifo, bufsz, NULL);
        } else
            break;
    }

    if (numchars)
        Tcl_SetObjResult(ticP->interp, ObjFromMultiSz(bufP, numchars+1));

    MemLifoPopFrame(&ticP->memlifo);

    return TCL_OK;
}

int Twapi_GetPrivateProfileSection(
    TwapiInterpContext *ticP,
    LPCWSTR lpAppName,
    LPCWSTR lpFileName
    )
{
    return TwapiGetProfileSectionHelper(ticP, lpAppName, lpFileName);
}

int Twapi_GetPrivateProfileSectionNames(
    TwapiInterpContext *ticP,
    LPCWSTR lpFileName
    )
{
    return TwapiGetProfileSectionHelper(ticP, NULL, lpFileName);
}

int Twapi_LoadUserProfile(
    Tcl_Interp *interp,
    HANDLE  hToken,
    DWORD                 flags,
    LPWSTR username,
    LPWSTR profilepath
    )
{
    PROFILEINFOW profileinfo;

    TwapiZeroMemory(&profileinfo, sizeof(profileinfo));
    profileinfo.dwSize        = sizeof(profileinfo);
    profileinfo.lpUserName    = username;
    profileinfo.lpProfilePath = profilepath;

    if (LoadUserProfileW(hToken, &profileinfo) == 0) {
        return TwapiReturnSystemError(interp);
    }

    Tcl_SetObjResult(interp, ObjFromHANDLE(profileinfo.hProfile));
    return TCL_OK;
}


typedef BOOLEAN (WINAPI *Wow64EnableWow64FsRedirection_t)(BOOLEAN);
MAKE_DYNLOAD_FUNC(Wow64EnableWow64FsRedirection, kernel32, Wow64EnableWow64FsRedirection_t)
BOOLEAN Twapi_Wow64EnableWow64FsRedirection(BOOLEAN enable_redirection)
{
    Wow64EnableWow64FsRedirection_t Wow64EnableWow64FsRedirectionPtr = Twapi_GetProc_Wow64EnableWow64FsRedirection();
    if (Wow64EnableWow64FsRedirectionPtr == NULL) {
        SetLastError(ERROR_PROC_NOT_FOUND);
    } else {
        if ((*Wow64EnableWow64FsRedirectionPtr)(enable_redirection))
            return TRUE;
        /* Not clear if the function sets last error so do it ourselves */
        if (GetLastError() == 0)
            SetLastError(ERROR_INVALID_FUNCTION); // For lack of better
    }
    return FALSE;
}

typedef BOOL (WINAPI *Wow64DisableWow64FsRedirection_t)(LPVOID *);
MAKE_DYNLOAD_FUNC(Wow64DisableWow64FsRedirection, kernel32, Wow64DisableWow64FsRedirection_t)
BOOLEAN Twapi_Wow64DisableWow64FsRedirection(LPVOID *oldvalueP)
{
    Wow64DisableWow64FsRedirection_t Wow64DisableWow64FsRedirectionPtr = Twapi_GetProc_Wow64DisableWow64FsRedirection();
    if (Wow64DisableWow64FsRedirectionPtr == NULL) {
        SetLastError(ERROR_PROC_NOT_FOUND);
        return FALSE;
    } else {
        return (*Wow64DisableWow64FsRedirectionPtr)(oldvalueP);
    }
}

typedef BOOL (WINAPI *Wow64RevertWow64FsRedirection_t)(LPVOID);
MAKE_DYNLOAD_FUNC(Wow64RevertWow64FsRedirection, kernel32, Wow64RevertWow64FsRedirection_t)
BOOLEAN Twapi_Wow64RevertWow64FsRedirection(LPVOID addr)
{
    Wow64RevertWow64FsRedirection_t Wow64RevertWow64FsRedirectionPtr = Twapi_GetProc_Wow64RevertWow64FsRedirection();
    if (Wow64RevertWow64FsRedirectionPtr == NULL) {
        SetLastError(ERROR_PROC_NOT_FOUND);
        return FALSE;
    } else {
        return (*Wow64RevertWow64FsRedirectionPtr)(addr);
    }
}

int Twapi_CommandLineToArgv(Tcl_Interp *interp, LPCWSTR cmdlineP)
{
    LPWSTR *argv;
    int     argc;
    int     i;
    Tcl_Obj *resultObj;

    argv = CommandLineToArgvW(cmdlineP, &argc);
    if (argv == NULL) {
        return TwapiReturnSystemError(interp);
    }

    resultObj = Tcl_NewListObj(0, NULL);
    for (i= 0; i < argc; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromUnicode(argv[i]));
    }

    Tcl_SetObjResult(interp, resultObj);

    GlobalFree(argv);
    return TCL_OK;
}

static int Twapi_AppCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s, s2, s3, s4;
    DWORD dw, dw2;
    HANDLE h, h2;
    WCHAR buf[MAX_PATH+1];
    TwapiResult result;
    LPVOID pv;
    WCHAR *bufP;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.type =
            GetProfileType(&result.value.ival) ? TRT_DWORD : TRT_GETLASTERROR;
        break;
    case 2:
        result.type = Twapi_Wow64DisableWow64FsRedirection(&result.value.pv) ?
            TRT_LPVOID : TRT_GETLASTERROR;
        break;
    case 3:
        result.value.unicode.str = GetCommandLineW();
        result.value.unicode.len = -1;
        result.type = TRT_UNICODE;
        break;
    case 4:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETVOIDP(pv),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = Twapi_Wow64RevertWow64FsRedirection(pv);
        break;
    case 5: // WritePrivateProfileString
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETNULLTOKEN(s2), GETNULLTOKEN(s3),
                         GETWSTR(s4),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = WritePrivateProfileStringW(s, s2, s3, s4);
        break;
    case 6:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETNULLTOKEN(s), GETNULLTOKEN(s2), GETNULLTOKEN(s3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = WriteProfileStringW(s, s2, s3);
        break;
    case 7:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETINT(dw), GETWSTR(s),
                         GETNULLIFEMPTY(s2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_LoadUserProfile(interp, h, dw, s, s2);
    case 8:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETHANDLE(h2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = UnloadUserProfile(h, h2);
        break;
    case 9:
    case 10:
    case 11:
    case 12:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        s = Tcl_GetUnicode(objv[2]);
        switch (func) {
        case 9:
            CHECK_INTEGER_OBJ(interp, dw, objv[2]);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = Twapi_Wow64EnableWow64FsRedirection((BOOLEAN)dw);
            break;
        case 10:
            return Twapi_CommandLineToArgv(interp, s);
        case 11:
            NULLIFY_EMPTY(s);
            return Twapi_GetPrivateProfileSectionNames(ticP, s);
        case 12:
            bufP = buf;
            dw = ExpandEnvironmentStringsW(s, bufP, ARRAYSIZE(buf));
            if (dw > ARRAYSIZE(buf)) {
                // Need a bigger buffer
                bufP = TwapiAlloc(dw * sizeof(WCHAR));
                dw2 = dw;
                dw = ExpandEnvironmentStringsW(s, bufP, dw2);
                if (dw > dw2) {
                    // Should not happen since we gave what we were asked
                    TwapiFree(bufP);
                    return TCL_ERROR;
                }
            }
            if (dw == 0)
                result.type = TRT_GETLASTERROR;
            else {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromUnicodeN(bufP, dw-1);
            }
            if (bufP != buf)
                TwapiFree(bufP);
            break;

        }
        break;
    case 13:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETWSTR(s2), GETINT(dw), GETWSTR(s3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_DWORD;
        result.value.ival = GetPrivateProfileIntW(s, s2, dw, s3);
        break;
    case 14:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETWSTR(s2), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_DWORD;
        result.value.ival = GetProfileIntW(s, s2, dw);
        break;
    case 15:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETWSTR(s2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_GetPrivateProfileSection(ticP, s, s2);
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_AppInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::AppCall", Twapi_AppCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::AppCall", # code_); \
    } while (0);

    CALL_(GetProfileType, 1);
    CALL_(Wow64DisableWow64FsRedirection, 2);
    CALL_(GetCommandLineW, 3);
    CALL_(Wow64RevertWow64FsRedirection, 4);
    CALL_(WritePrivateProfileString, 5);
    CALL_(WriteProfileString, 6);
    CALL_(Twapi_LoadUserProfile, 7);
    CALL_(UnloadUserProfile, 8);
    CALL_(Wow64EnableWow64FsRedirection, 9);
    CALL_(CommandLineToArgv, 10);
    CALL_(GetPrivateProfileSectionNames, 11);
    CALL_(ExpandEnvironmentStrings, 12);
    CALL_(GetPrivateProfileInt, 13);
    CALL_(GetProfileInt, 14);
    CALL_(GetPrivateProfileSection, 15);

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
int Twapi_apputil_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_AppInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

