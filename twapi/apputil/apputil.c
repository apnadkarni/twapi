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

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

BOOLEAN Twapi_Wow64DisableWow64FsRedirection(LPVOID *oldvalueP);
BOOLEAN Twapi_Wow64RevertWow64FsRedirection(LPVOID addr);
BOOLEAN Twapi_Wow64EnableWow64FsRedirection(BOOLEAN enable_redirection);

static int TwapiGetProfileSectionHelper(
    Tcl_Interp *interp,
    LPCWSTR lpAppName, /* If NULL, section names are retrieved */
    LPCWSTR lpFileName /* If NULL, win.ini is used */
    )
{
    WCHAR *bufP;
    DWORD  bufsz;
    DWORD  numchars;

    bufsz = 4000;
    bufP = SWSPushFrame(bufsz, NULL);
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
            /* Buf not big enough. Note no need to SWSPopFrame, just SWSAlloc */
            bufsz = 2*bufsz;
            bufP = SWSAlloc(bufsz, NULL);
        } else
            break;
    }

    if (numchars)
        ObjSetResult(interp, ObjFromMultiSz(bufP, numchars+1));

    SWSPopFrame();
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

    resultObj = ObjNewList(0, NULL);
    for (i= 0; i < argc; ++i) {
        ObjAppendElement(interp, resultObj, ObjFromUnicode(argv[i]));
    }

    ObjSetResult(interp, resultObj);

    GlobalFree(argv);
    return TCL_OK;
}

static int Twapi_AppCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR s;
    DWORD dw;
    TwapiResult result;
    LPVOID pv;
    int func = PtrToInt(clientdata);

    result.type = TRT_BADFUNCTIONCODE;
    --objc;
    ++objv;
    switch (func) {
    case 2:
        CHECK_NARGS(interp, objc, 0);
        if (Twapi_Wow64DisableWow64FsRedirection(&pv))
            TwapiResult_SET_PTR(result, void*, pv);
        else
            result.type = TRT_GETLASTERROR;
        break;
    case 3:
        CHECK_NARGS(interp, objc, 0);
        result.value.unicode.str = GetCommandLineW();
        result.value.unicode.len = -1;
        result.type = TRT_UNICODE;
        break;
    case 4:
        CHECK_NARGS(interp, objc, 1);
        if (ObjToLPVOID(interp, objv[0], &pv) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = Twapi_Wow64RevertWow64FsRedirection(pv);
        break;
    case 5: // WritePrivateProfileString
        CHECK_NARGS(interp, objc, 4);
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = WritePrivateProfileStringW(
            ObjToUnicode(objv[0]), ObjToLPWSTR_WITH_NULL(objv[1]),
            ObjToLPWSTR_WITH_NULL(objv[2]), ObjToUnicode(objv[3]));
        break;
    case 6:
        CHECK_NARGS(interp, objc, 3);
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = WriteProfileStringW(
            ObjToLPWSTR_WITH_NULL(objv[0]),
            ObjToLPWSTR_WITH_NULL(objv[1]),
            ObjToLPWSTR_WITH_NULL(objv[2]));
        break;
    case 7:
        CHECK_NARGS(interp, objc, 4);
        /* As always, extract the int first to avoid shimmering issues */
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        result.type = TRT_LONG;
        result.value.ival = GetPrivateProfileIntW(
            ObjToLPWSTR_NULL_IF_EMPTY(objv[0]),
            ObjToUnicode(objv[1]), dw,
            ObjToUnicode(objv[3]));
        break;
    case 8:
        CHECK_NARGS(interp, objc, 3);
        /* As always, extract the int first to avoid shimmering issues */
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        result.type = TRT_LONG;
        result.value.ival = GetProfileIntW(
            ObjToUnicode(objv[0]), ObjToUnicode(objv[1]), dw);
        break;
    case 9:
        CHECK_NARGS(interp, objc, 1);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = Twapi_Wow64EnableWow64FsRedirection((BOOLEAN)dw);
        break;
    case 10:
        CHECK_NARGS(interp, objc, 1);
        return Twapi_CommandLineToArgv(interp, ObjToUnicode(objv[0]));
    case 11:
        CHECK_NARGS(interp, objc, 1);
        s = ObjToUnicode(objv[0]);
        NULLIFY_EMPTY(s);
        return TwapiGetProfileSectionHelper(interp, NULL, s);
    case 12: // GetPrivateProfileSection
        CHECK_NARGS(interp, objc, 2);
        return TwapiGetProfileSectionHelper(
            interp,
            ObjToUnicode(objv[0]),
            ObjToLPWSTR_NULL_IF_EMPTY(objv[1]));
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiApputilInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s AppCallDispatch[] = {
        DEFINE_FNCODE_CMD(Wow64DisableWow64FsRedirection, 2),
        DEFINE_FNCODE_CMD(GetCommandLineW, 3),
        DEFINE_FNCODE_CMD(Wow64RevertWow64FsRedirection, 4),
        DEFINE_FNCODE_CMD(WritePrivateProfileString, 5),
        DEFINE_FNCODE_CMD(WriteProfileString, 6),
        DEFINE_FNCODE_CMD(GetPrivateProfileInt, 7),
        DEFINE_FNCODE_CMD(GetProfileInt, 8),
        DEFINE_FNCODE_CMD(Wow64EnableWow64FsRedirection, 9),
        DEFINE_FNCODE_CMD(CommandLineToArgv, 10),
        DEFINE_FNCODE_CMD(GetPrivateProfileSectionNames, 11),
        DEFINE_FNCODE_CMD(GetPrivateProfileSection, 12),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(AppCallDispatch), AppCallDispatch, Twapi_AppCallObjCmd);

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
int Twapi_apputil_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiApputilInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

