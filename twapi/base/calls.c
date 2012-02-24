/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

static char *apiprocs = 
    "#\n"
    "# Definitions for Win32 API calls through a standard handler\n"
    "# Note most are defined at C level in twapi_calls.c and this file\n"
    "# only contains those that could not be defined there for whatever\n"
    "# reason (generally because argument order is changed, or defaults are\n"
    "# needed etc.)\n"
    "namespace eval twapi {}\n"
    "# Call - function(void)\n"
    "proc twapi::UuidCreateNil {} { return 00000000-0000-0000-0000-000000000000 } \n"
    "                                                                        \n"
    "# twapi::CallSSSD - function(LPWSTR_NULL_IF_EMPTY, LPWSTR, LPWSTR, DWORD) \n"
    "proc twapi::DefineDosDevice  {flags devname path} {\n"
    "    return [CallSSSD 42 {} $devname $path $flags]\n"
    "}\n"
    "proc twapi::GetPrivateProfileInt {app key ival file} {\n"
    "    return [CallSSSD 7 $app $key $file $ival]\n"
    "}\n"
    "proc twapi::GetProfileInt {app key ival} {\n"
    "    return [CallSSSD 8 {} $app $key $ival]\n"
    "}\n"
    "proc twapi::GetPrivateProfileSection {app file} {\n"
    "    return [CallSSSD 9 $file $app]\n"
    "}\n"
    "proc twapi::IsValidSid {sid} {\n"
    "    return [CallPSID 1 {} $sid]\n"
    "}\n"
    ;

TCL_RESULT TwapiGetArgs(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[],
                 char fmtch, ...)
{
    int        argno;
    va_list    ap;
    void      *p;
    Tcl_Obj   *objP = 0;
    char      *typeP;              /* Type of a pointer */
    int       *lenP;
    int        ival;
    Tcl_WideInt wival;
    DWORD_PTR  dwval;
    void      *ptrval;
    double     dblval;
    WCHAR     *uval;
    char      *sval;
    TwapiGetArgsFn converter_fn;
    int        len;
    int        use_default = 0;
    int        *iP;

    va_start(ap,fmtch);
    for (argno = -1; fmtch != ARGEND && fmtch != ARGTERM; fmtch = va_arg(ap, char)) {
        if (fmtch == ARGUSEDEFAULT) {
            use_default = 1;
            continue;
        }

        if (++argno >= objc) {
            /* No more Tcl_Obj's. See if we can use defaults, else break */
            if (! use_default) {
                TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
                goto argerror;
            }
            objP = NULL;
        } else {
            objP = objv[argno];
        }

        if (fmtch == ARGSKIP)
            continue;           /* Jump over objv[argno] */
            
        p = va_arg(ap, void *); /* May be NULL if caller wants type check
                                   but does not care for value */
        switch (fmtch) {
        case ARGBOOL:
            ival = 0; // Default
            if (objP && Tcl_GetBooleanFromObj(interp, objP, &ival) != TCL_OK)
                    goto argerror;
            if (p)
                *(int *)p = ival;
            break;

        case ARGBIN: // bytearray
            lenP = va_arg(ap, int *);
            if (p || lenP) {
                ptrval = NULL; // Default
                len = 0; // Default
                if (objP)
                    ptrval = Tcl_GetByteArrayFromObj(objP, &len);
            }
            if (p)
                *(unsigned char **)p = (unsigned char *)ptrval;
            if (lenP)
                *lenP = len;
            break;

        case ARGDOUBLE: // double
            dblval = 0.0; // Default
            if (objP && Tcl_GetDoubleFromObj(interp, objP, &dblval) != TCL_OK)
                goto argerror;
            if (p)
                *(double *)p = dblval;
            break;
        case ARGINT:  // int
            ival = 0; // Default
            if (objP && Tcl_GetIntFromObj(interp, objP, &ival) != TCL_OK)
                goto argerror;
            if (p)
                *(int *)p = ival;
            break;
        case ARGWIDE: // 64-bit int
            wival = 0;
            if (objP && Tcl_GetWideIntFromObj(interp, objP, &wival) != TCL_OK)
                goto argerror;
            if (p)
                *(Tcl_WideInt *)p = wival;
            break;
        case ARGOBJ: // Tcl object
            if (p)
                *(Tcl_Obj **)p = objP; // May be NULL (when use_default is 1)
            break;
        case ARGPTR:
            typeP = va_arg(ap, char *);
            ptrval = NULL;
            if (objP && ObjToOpaque(interp, objP, &ptrval, typeP) != TCL_OK)
                goto argerror;
            if (p)
                *(void **)p = ptrval;
            break;
        case ARGDWORD_PTR: // pointer-size int
            dwval = 0;
            if (objP && ObjToDWORD_PTR(interp, objP, &dwval) != TCL_OK)
                goto argerror;
            if (p)
                *(DWORD_PTR *)p = dwval;
            break;
        case ARGASTR: // char string
            if (p)
                *(char **)p = objP ? Tcl_GetString(objP) : "";
            break;
        case ARGASTRN: // char string and its length
            lenP = va_arg(ap, int *);
            sval = "";
            len = 0;
            if (objP)
                sval = Tcl_GetStringFromObj(objP, &len);
            if (p)
                *(char **)p = sval;
            if (lenP)
                *lenP = len;
            break;
        case ARGWSTR: // Unicode string
            if (p) {
                *(WCHAR **)p = objP ? Tcl_GetUnicode(objP) : L"" ;
            }
            break;
        case ARGNULLIFEMPTY:
            if (p)
                *(WCHAR **)p = ObjToLPWSTR_NULL_IF_EMPTY(objP); // NULL objP ok
            break;
        case ARGNULLTOKEN:
            if (p)
                *(WCHAR **)p = ObjToLPWSTR_WITH_NULL(objP);     // NULL objP ok
            break;
        case ARGWSTRN:
            /* We want string and its length */
            lenP = va_arg(ap, int *);
            uval = L""; // Defaults
            len = 0;
            if (objP)
                uval = Tcl_GetUnicodeFromObj(objP, &len);
            if (p)
                *(WCHAR **)p = uval;
            if (lenP)
                *lenP = len;
            break;
        case ARGWORD: // WORD - 16 bits
            ival = 0;
            if (objP && Tcl_GetIntFromObj(interp, objP, &ival) != TCL_OK)
                goto argerror;
            if (ival & ~0xffff) {
                TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                                   Tcl_ObjPrintf("Value %d does not fit in 16 bits.", ival));
                goto argerror;
            }
            if (p)
                *(short *)p = (short) ival;
            break;
            
        case ARGVAR: // Does not handle default.
            if (objP == NULL) {
                Tcl_SetResult(interp, "Default values cannot be used for ARGVAR types.", TCL_STATIC);
                goto argerror;
            }
            // FALLTHRU
        case ARGVARWITHDEFAULT: // Allows objP to be NULL. The converter_fn should also allow that
            converter_fn = va_arg(ap, TwapiGetArgsFn);
            if (p) {
                if (converter_fn(interp, objP, p) != TCL_OK)
                    goto argerror;
            }
            break;

        case ARGAARGV:
        case ARGWARGV:
            if (objP) {
                ival = va_arg(ap, int);
                iP = va_arg(ap, int *);
                if (iP == NULL)
                    iP = &ival;
                if (fmtch == ARGAARGV) {
                    if (ObjToArgvA(interp, objP, p, ival, iP) != TCL_OK)
                        goto argerror;
                } else {
                    if (ObjToArgvW(interp, objP, p, ival, iP) != TCL_OK)
                        goto argerror;
                }
            } else if (iP)
                *iP = 0;
            break;

        default:
            Tcl_SetResult(interp, "Unexpted format character passed to TwapiGetArgs.", TCL_STATIC);
            goto argerror;
        }

    }

    if (fmtch == ARGEND) {
        /* Should be end of arguments. For an exact match against number
           of supplied objects, argno will be objc-1 since it is incremented
           inside the loop.
        */
        if (argno < (objc-1)) {
            TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            goto argerror;
        }
    } else if (fmtch == ARGTERM) {
        /* Caller only wants partial parse, don't care to check more args */
    } else {
        /* Premature end of arguments */
        Tcl_SetResult(interp, "Insufficient number of arguments.", TCL_STATIC);
        goto argerror;
    }

    va_end(ap);
    return TCL_OK;

argerror:
    /* interp is already supposed to contain an error message */
    va_end(ap);
    return TCL_ERROR;
}


void Twapi_MakeCallAlias(Tcl_Interp *interp, char *fn, char *callcmd, char *code)
{
   /*
    * Why a single line function ?
    * Making this a function instead of directly calling Tcl_CreateAlias from
    * Twapi_InitCalls saves about 4K in code space. (Yes, every K is important,
    * users are already complaining wrt the DLL size
    */

    Tcl_CreateAlias(interp, fn, interp, callcmd, 1, &code);
}

int Twapi_InitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::Call", Twapi_CallObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallU", Twapi_CallUObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallS", Twapi_CallSObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallH", Twapi_CallHObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallSSSD", Twapi_CallSSSDObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallWU", Twapi_CallWUObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallPSID", Twapi_CallPSIDObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallCOM", Twapi_CallCOMObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::" #call_, # code_); \
    } while (0);

    /*
     * NOTE: Some MISSING CALL NUMBERS are defined in win32calls.tcl
     * or as procs in the apiprocs global.
     */
    CALL_(GetCurrentProcess, Call, 1);
    CALL_(GetUserDefaultLangID, Call, 23);
    CALL_(GetSystemDefaultLangID, Call, 24);
    CALL_(GetUserDefaultLCID, Call, 25);
    CALL_(GetSystemDefaultLCID, Call, 26);
    CALL_(GetUserDefaultUILanguage, Call, 27);
    CALL_(GetSystemDefaultUILanguage, Call, 28);
    CALL_(GetThreadLocale, Call, 29);
    CALL_(GetACP, Call, 30);
    CALL_(GetOEMCP, Call, 31);
    CALL_(GetComputerName, Call, 32);
    CALL_(GetVersionEx, Call, 33);
    CALL_(GetSystemInfo, Call, 34);
    CALL_(GlobalMemoryStatus, Call, 35);
    CALL_(Twapi_SystemProcessorTimes, Call, 36);
    CALL_(Twapi_SystemPagefileInformation, Call, 37);
    CALL_(GetProfileType, Call, 38);
    CALL_(GetTickCount, Call, 39);
    CALL_(GetSystemTimeAsFileTime, Call, 40);
    CALL_(Wow64DisableWow64FsRedirection, Call, 41);
    CALL_(GetSystemWow64Directory, Call, 42);
    CALL_(GetCommandLineW, Call, 47);
    CALL_(AllocateLocallyUniqueId, Call, 48);
    CALL_(LockWorkStation, Call, 49);
    CALL_(LsaEnumerateLogonSessions, Call, 50);
    CALL_(RevertToSelf, Call, 51);
    CALL_(Twapi_InitializeSecurityDescriptor, Call, 52);
    CALL_(GetSystemPowerStatus, Call, 68);
    CALL_(Twapi_PowerNotifyStart, Call, 71);
    CALL_(Twapi_PowerNotifyStop, Call, 72);
    CALL_(TwapiId, Call, 74);
    CALL_(DebugBreak, Call, 75);
    CALL_(GetPerformanceInformation, Call, 76);
    CALL_(Twapi_GetNotificationWindow, Call, 77);
    CALL_(GetSystemWindowsDirectory, Call, 78); /* TBD Tcl */
    CALL_(GetWindowsDirectory, Call, 79);       /* TBD Tcl */
    CALL_(GetSystemDirectory, Call, 80);        /* TBD Tcl */
    CALL_(GetDefaultPrinter, Call, 82);         /* TBD Tcl */
    CALL_(GetTimeZoneInformation, Call, 83);    /* TBD Tcl */

    CALL_(Twapi_AddressToPointer, Call, 1001);
    CALL_(VariantTimeToSystemTime, Call, 1003);
    CALL_(SystemTimeToVariantTime, Call, 1004);
    CALL_(canonicalize_guid, Call, 1005); // TBD Document
    CALL_(IsValidSecurityDescriptor, Call, 1010);
    CALL_(IsValidAcl, Call, 1014);
    CALL_(FileTimeToSystemTime, Call, 1016);
    CALL_(SystemTimeToFileTime, Call, 1017);
    CALL_(Wow64RevertWow64FsRedirection, Call, 1018);
    CALL_(Twapi_IsValidGUID, Call, 1019);
    CALL_(Twapi_UnregisterWaitOnHandle, Call, 1020);
    CALL_(free, Call, 1022);

    CALL_(DuplicateHandle, Call, 10008);
    CALL_(Tcl_GetChannelHandle, Call, 10009);
    CALL_(ConvertSecurityDescriptorToStringSecurityDescriptor, Call, 10017);
    CALL_(LsaQueryInformationPolicy, Call, 10018);
    CALL_(LsaGetLogonSessionData, Call, 10019);
    CALL_(CreateFile, Call, 10031);
    CALL_(SetNamedSecurityInfo, Call, 10032);
    CALL_(LookupPrivilegeName, Call, 10036);
    CALL_(SetSecurityInfo, Call, 10041);
    CALL_(WTSSendMessage, Call, 10044);
    CALL_(DuplicateTokenEx, Call, 10045);
    CALL_(Twapi_AdjustTokenPrivileges, Call, 10047);
    CALL_(Twapi_PrivilegeCheck, Call, 10048);
    CALL_(DsGetDcName, Call, 10058);
    CALL_(GetNumberFormat, Call, 10070);
    CALL_(GetCurrencyFormat, Call, 10071);
    CALL_(InitiateSystemShutdown, Call, 10072);
    CALL_(FormatMessageFromModule, Call, 10073);
    CALL_(FormatMessageFromString, Call, 10074);
    CALL_(WritePrivateProfileString, Call, 10075);
    CALL_(WriteProfileString, Call, 10076);
    CALL_(SystemParametersInfo, Call, 10077);
    CALL_(Twapi_LoadUserProfile, Call, 10078);
    CALL_(UnloadUserProfile, Call, 10079);
    CALL_(SetSuspendState, Call, 10080);
    CALL_(win32_error, Call, 10081);
    CALL_(CreateMutex, Call, 10097);
    CALL_(OpenMutex, Call, 10098);
    CALL_(OpenSemaphore, Call, 10099); /* TBD - Tcl wrapper */
    CALL_(CreateSemaphore, Call, 10100); /* TBD - Tcl wrapper */
    CALL_(Twapi_ReadMemoryInt, Call, 10101);
    CALL_(Twapi_ReadMemoryBinary, Call, 10102);
    CALL_(Twapi_ReadMemoryChars, Call, 10103);
    CALL_(Twapi_ReadMemoryUnicode, Call, 10104);
    CALL_(Twapi_ReadMemoryPointer, Call, 10105);
    CALL_(Twapi_ReadMemoryWide, Call, 10106);
    CALL_(malloc, Call, 10110);        /* TBD - document, change to memalloc */
    CALL_(Twapi_WriteMemoryInt, Call, 10111);
    CALL_(Twapi_WriteMemoryBinary, Call, 10112);
    CALL_(Twapi_WriteMemoryChars, Call, 10113);
    CALL_(Twapi_WriteMemoryUnicode, Call, 10114);
    CALL_(Twapi_WriteMemoryPointer, Call, 10115);
    CALL_(Twapi_WriteMemoryWide, Call, 10116);
    CALL_(Twapi_IsEqualPtr, Call, 10119);
    CALL_(Twapi_IsNullPtr, Call, 10120);
    CALL_(Twapi_IsPtr, Call, 10121);
    CALL_(CreateEvent, Call, 10122);
    CALL_(TzSpecificLocalTimeToSystemTime, Call, 10134); // Tcl
    CALL_(SystemTimeToTzSpecificLocalTime, Call, 10135); // Tcl
    CALL_(IsEqualGUID, Call, 10136); // Tcl
    CALL_(IMofCompiler_CompileBuffer, Call, 10137); // Tcl
    CALL_(IMofCompiler_CompileFile, Call, 10138); // Tcl
    CALL_(IMofCompiler_CreateBMOF, Call, 10139); // Tcl

    // CallU API
    CALL_(GetStdHandle, CallU, 4);
    CALL_(VerLanguageName, CallU, 7);
    CALL_(GetComputerNameEx, CallU, 15);
    CALL_(Wow64EnableWow64FsRedirection, CallU, 16);
    CALL_(GetSystemMetrics, CallU, 17);
    CALL_(Sleep, CallU, 18);
    CALL_(Twapi_EnumPrinters_Level4, CallU, 20);
    CALL_(UuidCreate, CallU, 21);
    CALL_(GetUserNameEx, CallU, 22);
    CALL_(ImpersonateSelf, CallU, 23);
    CALL_(Sleep, CallU, 33);
    CALL_(Twapi_MapWindowsErrorToString, CallU, 34);
    CALL_(Twapi_MemLifoInit, CallU, 37);
    CALL_(GlobalDeleteAtom, CallU, 38); // TBD - tcl interface

    CALL_(GetLocaleInfo, CallU, 1006);
    CALL_(ExitWindowsEx, CallU, 1007);

    CALL_(AttachThreadInput, CallU, 2004); /* Must stay in twapi_base as
                                              potentially used by many exts */

    CALL_(GlobalAlloc, CallU, 10001);
    CALL_(LHashValOfName, CallU, 10002);
    CALL_(SetStdHandle, CallU, 10004);

    // CallS - function(LPWSTR)
    CALL_(CommandLineToArgv, CallS, 8);
    CALL_(Twapi_AppendLog, CallS, 11);
    CALL_(WTSOpenServer, CallS, 12);
    CALL_(AbortSystemShutdown, CallS, 17);
    CALL_(GetPrivateProfileSectionNames, CallS, 18);
    CALL_(ExpandEnvironmentStrings, CallS, 22);
    CALL_(GlobalAddAtom, CallS, 23); // TBD - Tcl interface
    CALL_(is_valid_sid_syntax, CallS, 27); // TBD - Tcl interface

    CALL_(ConvertStringSecurityDescriptorToSecurityDescriptor, CallS, 501);
    CALL_(Twapi_LsaOpenPolicy, CallS, 502);
    CALL_(LoadLibraryEx, CallS, 504);

    CALL_(GetNamedSecurityInfo, CallS, 1004);
    CALL_(TranslateName, CallS, 1005);

    // CallH - function(HANDLE)
    CALL_(GetHandleInformation, CallH, 14);
    CALL_(FreeLibrary, CallH, 15);
    CALL_(GetDevicePowerState, CallH, 16); // TBD - which module ?
    CALL_(LsaClose, CallH, 26);
    CALL_(ImpersonateLoggedOnUser, CallH, 27);
    CALL_(ReleaseMutex, CallH, 42);
    CALL_(CloseHandle, CallH, 43);
    CALL_(CastToHANDLE, CallH, 44);
    CALL_(GlobalFree, CallH, 45);
    CALL_(GlobalUnlock, CallH, 46);
    CALL_(GlobalSize, CallH, 47);
    CALL_(GlobalLock, CallH, 48);
    CALL_(WTSEnumerateSessions, CallH, 49);
    CALL_(WTSEnumerateProcesses, CallH, 50);
    CALL_(WTSCloseServer, CallH, 51);
    CALL_(Twapi_MemLifoClose, CallH, 54);
    CALL_(Twapi_MemLifoPopFrame, CallH, 55);
    CALL_(GetObject, CallH, 59);
    CALL_(Twapi_MemLifoPushMark, CallH, 60);
    CALL_(Twapi_MemLifoPopMark, CallH, 61);
    CALL_(Twapi_MemLifoValidate, CallH, 62);
    CALL_(Twapi_MemLifoDump, CallH, 63);
    CALL_(ImpersonateNamedPipeClient, CallH, 64);
    CALL_(SetEvent, CallH, 66);
    CALL_(ResetEvent, CallH, 67);

    CALL_(ReleaseSemaphore, CallH, 1001);
    CALL_(OpenProcessToken, CallH, 1005);
    CALL_(GetTokenInformation, CallH, 1006);
    CALL_(Twapi_SetTokenVirtualizationEnabled, CallH, 1007);
    CALL_(Twapi_SetTokenMandatoryPolicy, CallH, 1008);
    CALL_(GetDeviceCaps, CallH, 1016);
    CALL_(WaitForSingleObject, CallH, 1017);
    CALL_(Twapi_MemLifoAlloc, CallH, 1018);
    CALL_(Twapi_MemLifoPushFrame, CallH, 1019);

    CALL_(WTSDisconnectSession, CallH, 2001);
    CALL_(WTSLogoffSession, CallH, 2003);        /* TBD - tcl wrapper */
    CALL_(WTSQuerySessionInformation, CallH, 2003); /* TBD - tcl wrapper */
    CALL_(GetSecurityInfo, CallH, 2004);
    CALL_(OpenThreadToken, CallH, 2005);
    CALL_(SetHandleInformation, CallH, 2007); /* TBD - Tcl wrapper */
    CALL_(Twapi_MemLifoExpandLast, CallH, 2008);
    CALL_(Twapi_MemLifoShrinkLast, CallH, 2009);
    CALL_(Twapi_MemLifoResizeLast, CallH, 2010);
    CALL_(Twapi_RegisterWaitOnHandle, CallH, 2011);

    CALL_(SetThreadToken, CallH, 10002);
    CALL_(Twapi_LsaEnumerateAccountsWithUserRight, CallH, 10003);
    CALL_(Twapi_SetTokenIntegrityLevel, CallH, 10004);

    CALL_(GetWindowLongPtr, CallWU, 8);

    CALL_(PostMessage, CallWU, 1001);
    CALL_(SendNotifyMessage, CallWU, 1002);
    CALL_(SendMessageTimeout, CallWU, 1003);

    CALL_(SetWindowLongPtr, CallWU, 10003);

    // CallSSSD - function(LPWSTR_NULL_IF_EMPTY, LPWSTR, LPWSTR, DWORD)
    CALL_(LookupAccountName, CallSSSD, 1);
    CALL_(LogonUser, CallSSSD, 5);
    CALL_(LookupPrivilegeDisplayName, CallSSSD, 13);
    CALL_(LookupPrivilegeValue, CallSSSD, 14);
    CALL_(NetGetDCName, CallSSSD, 34);
    CALL_(QueryDosDevice, CallSSSD, 41); /* TBD what module should this go to */

    // CallPSID - function(ANY, SID, ...)
    CALL_(LookupAccountSid, CallPSID, 2);

    CALL_(CheckTokenMembership, CallPSID, 1002);
    CALL_(Twapi_SetTokenPrimaryGroup, CallPSID, 1003);
    CALL_(Twapi_SetTokenOwner, CallPSID, 1004);
    CALL_(Twapi_LsaEnumerateAccountRights, CallPSID, 1005);
    CALL_(LsaRemoveAccountRights, CallPSID, 1006);
    CALL_(LsaAddAccountRights, CallPSID, 1007);


#undef CALL_

    // CallCOM
#define CALLCOM_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::CallCOM", # code_); \
    } while (0);
    CALLCOM_(IUnknown_Release, 1);
    CALLCOM_(IUnknown_AddRef, 2);
    CALLCOM_(Twapi_IUnknown_QueryInterface, 3);
    CALLCOM_(OleRun, 4);       /* Note - function, NOT method */

    CALLCOM_(IDispatch_GetTypeInfoCount, 101);
    CALLCOM_(IDispatch_GetTypeInfo, 102);
    CALLCOM_(IDispatch_GetIDsOfNames, 103);

    CALLCOM_(IDispatchEx_GetDispID, 201);
    CALLCOM_(IDispatchEx_GetMemberName, 202);
    CALLCOM_(IDispatchEx_GetMemberProperties, 203);
    CALLCOM_(IDispatchEx_GetNextDispID, 204);
    CALLCOM_(IDispatchEx_GetNameSpaceParent, 205);
    CALLCOM_(IDispatchEx_DeleteMemberByName, 206);
    CALLCOM_(IDispatchEx_DeleteMemberByDispID, 207);

    CALLCOM_(ITypeInfo_GetRefTypeOfImplType, 301);
    CALLCOM_(ITypeInfo_GetRefTypeInfo, 302);
    CALLCOM_(ITypeInfo_GetTypeComp, 303);
    CALLCOM_(ITypeInfo_GetContainingTypeLib, 304);
    CALLCOM_(ITypeInfo_GetDocumentation, 305);
    CALLCOM_(ITypeInfo_GetImplTypeFlags, 306);
    CALLCOM_(GetRecordInfoFromTypeInfo, 307); /* Note - function, not method */
    CALLCOM_(ITypeInfo_GetNames, 308);
    CALLCOM_(ITypeInfo_GetTypeAttr, 309);
    CALLCOM_(ITypeInfo_GetFuncDesc, 310);
    CALLCOM_(ITypeInfo_GetVarDesc, 311);
    CALLCOM_(ITypeInfo_GetIDsOfNames, 399);

    CALLCOM_(ITypeLib_GetDocumentation, 401);
    CALLCOM_(ITypeLib_GetTypeInfoCount, 402);
    CALLCOM_(ITypeLib_GetTypeInfoType, 403);
    CALLCOM_(ITypeLib_GetTypeInfo, 404);
    CALLCOM_(ITypeLib_GetTypeInfoOfGuid, 405);
    CALLCOM_(ITypeLib_GetLibAttr, 406);
    CALLCOM_(RegisterTypeLib, 407); /* Function, not method */

    CALLCOM_(IRecordInfo_GetField, 501);
    CALLCOM_(IRecordInfo_GetGuid, 502);
    CALLCOM_(IRecordInfo_GetName, 503);
    CALLCOM_(IRecordInfo_GetSize, 504);
    CALLCOM_(IRecordInfo_GetTypeInfo, 505);
    CALLCOM_(IRecordInfo_IsMatchingType, 506);
    CALLCOM_(IRecordInfo_RecordClear, 507);
    CALLCOM_(IRecordInfo_RecordCopy, 508);
    CALLCOM_(IRecordInfo_RecordCreate, 509);
    CALLCOM_(IRecordInfo_RecordCreateCopy, 510);
    CALLCOM_(IRecordInfo_RecordDestroy, 511);
    CALLCOM_(IRecordInfo_RecordInit, 512);
    CALLCOM_(IRecordInfo_GetFieldNames, 513);

    CALLCOM_(IMoniker_GetDisplayName,601);

    CALLCOM_(IEnumVARIANT_Clone, 701);
    CALLCOM_(IEnumVARIANT_Reset, 702);
    CALLCOM_(IEnumVARIANT_Skip, 703);
    CALLCOM_(IEnumVARIANT_Next, 704);

    CALLCOM_(IConnectionPoint_Advise, 801);
    CALLCOM_(IConnectionPoint_EnumConnections, 802);
    CALLCOM_(IConnectionPoint_GetConnectionInterface, 803);
    CALLCOM_(IConnectionPoint_GetConnectionPointContainer, 804);
    CALLCOM_(IConnectionPoint_Unadvise, 805);

    CALLCOM_(IConnectionPointContainer_EnumConnectionPoints, 901);
    CALLCOM_(IConnectionPointContainer_FindConnectionPoint, 902);

    CALLCOM_(IEnumConnectionPoints_Clone, 1001);
    CALLCOM_(IEnumConnectionPoints_Reset, 1002);
    CALLCOM_(IEnumConnectionPoints_Skip, 1003);
    CALLCOM_(IEnumConnectionPoints_Next, 1004);

    CALLCOM_(IEnumConnections_Clone, 1101);
    CALLCOM_(IEnumConnections_Reset, 1102);
    CALLCOM_(IEnumConnections_Skip, 1103);
    CALLCOM_(IEnumConnections_Next, 1104);

    CALLCOM_(IProvideClassInfo_GetClassInfo, 1201);

    CALLCOM_(IProvideClassInfo2_GetGUID, 1301);

    CALLCOM_(ITypeComp_Bind, 1401);


    CALLCOM_(IPersistFile_GetCurFile, 5501);
    CALLCOM_(IPersistFile_IsDirty, 5502);
    CALLCOM_(IPersistFile_Load, 5503);
    CALLCOM_(IPersistFile_Save, 5504);
    CALLCOM_(IPersistFile_SaveCompleted, 5505);

    CALLCOM_(CreateFileMoniker, 10001);
    CALLCOM_(CreateBindCtx, 10002);
    CALLCOM_(GetRecordInfoFromGuids, 10003);
    CALLCOM_(QueryPathOfRegTypeLib, 10004);
    CALLCOM_(UnRegisterTypeLib, 10005);
    CALLCOM_(LoadRegTypeLib, 10006);
    CALLCOM_(LoadTypeLibEx, 10007);
    CALLCOM_(Twapi_CoGetObject, 10008);
    CALLCOM_(GetActiveObject, 10009);
    CALLCOM_(ProgIDFromCLSID, 10010);
    CALLCOM_(CLSIDFromProgID, 10011);
    CALLCOM_(Twapi_CoCreateInstance, 10012);
#undef CALLCOM_

    return Tcl_Eval(interp, apiprocs);
}


int Twapi_CallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    union {
        WCHAR buf[MAX_PATH+1];
        LPCWSTR wargv[100];     /* FormatMessage accepts up to 99 params + 1 for NULL */
        double d;
        FILETIME   filetime;
        TIME_ZONE_INFORMATION tzinfo;
        LARGE_INTEGER largeint;
        TOKEN_PRIVILEGES *tokprivsP;
        MIB_TCPROW tcprow;
        struct sockaddr_in sinaddr;
        SYSTEM_POWER_STATUS power_status;
        TwapiId twapi_id;
        GUID guid;
    } u;
    DWORD_PTR dwp;
    SECURITY_DESCRIPTOR *secdP;
    DWORD dw, dw2, dw3, dw4;
    int i, i2;
    LPWSTR s, s2, s3, s4;
    unsigned char *cP;
    LUID luid;
    void *pv, *pv2;
    Tcl_Obj *objs[2];
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h, h2, h3;
    PSID osidP, gsidP;
    ACL *daclP, *saclP;
    GUID guid;
    GUID *guidP;
    SYSTEMTIME systime;
    TIME_ZONE_INFORMATION *tzinfoP;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 1000) {
        /* Functions taking no arguments */
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1:
            result.type = TRT_HANDLE;
            result.value.hval = GetCurrentProcess();
            break;
            // 2-22 UNUSED
        case 23:
            result.type = TRT_DWORD;
            result.value.ival = GetUserDefaultLangID();
            break;
        case 24:
            result.type = TRT_DWORD;
            result.value.ival = GetSystemDefaultLangID();
            break;
        case 25:
            result.type = TRT_DWORD;
            result.value.ival = GetUserDefaultLCID();
            break;
        case 26:
            result.type = TRT_DWORD;
            result.value.ival = GetSystemDefaultLCID();
            break;
        case 27:
            result.type = TRT_NONZERO_RESULT;
            result.value.ival = GetUserDefaultUILanguage();
            break;
        case 28:
            result.type = TRT_NONZERO_RESULT;
            result.value.ival = GetSystemDefaultUILanguage();
            break;
        case 29:
            result.type = TRT_DWORD;
            result.value.ival = GetThreadLocale();
            break;
        case 30:
            result.type = TRT_DWORD;
            result.value.ival = GetACP();
            break;
        case 31:
            result.type = TRT_DWORD;
            result.value.ival = GetOEMCP();
            break;
        case 32:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetComputerNameW(u.buf, &result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 33:
            return Twapi_GetVersionEx(interp);
        case 34:
            return Twapi_GetSystemInfo(interp);
        case 35:
            return Twapi_GlobalMemoryStatus(interp);
        case 36:
            return Twapi_SystemProcessorTimes(ticP);
        case 37:
            return Twapi_SystemPagefileInformation(ticP);
        case 38:
            result.type = GetProfileType(&result.value.ival) ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 39:
            result.type = TRT_DWORD;
            result.value.ival = GetTickCount();
            break;
        case 40:
            result.type = TRT_FILETIME;
            GetSystemTimeAsFileTime(&result.value.filetime);
            break;
        case 41:
            result.type = Twapi_Wow64DisableWow64FsRedirection(&result.value.pv) ?
                TRT_LPVOID : TRT_GETLASTERROR;
            break;
        case 42:
            return Twapi_GetSystemWow64Directory(interp);
            // 43-46 UNUSED
        case 47:
            result.value.unicode.str = GetCommandLineW();
            result.value.unicode.len = -1;
            result.type = TRT_UNICODE;
            break;
        case 48:
            result.type = AllocateLocallyUniqueId(&result.value.luid) ? TRT_LUID : TRT_GETLASTERROR;
            break;
        case 49:
            result.value.ival = LockWorkStation();
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 50:
            return Twapi_LsaEnumerateLogonSessions(interp);
        case 51:
            result.value.ival = RevertToSelf();
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 52:
            return Twapi_InitializeSecurityDescriptor(interp);
            // 53-67 UNUSED
        case 68:
            if (GetSystemPowerStatus(&u.power_status)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromSYSTEM_POWER_STATUS(&u.power_status);
            } else
                result.type = TRT_EXCEPTION_ON_FALSE;
            break;
            // 69-70 UNUSED
        case 71:
            return Twapi_PowerNotifyStart(ticP);
        case 72:
            return Twapi_PowerNotifyStop(ticP);
        case 73: // UNUSED
            break;
        case 74:
            result.type = TRT_WIDE;
            result.value.wide = TWAPI_NEWID(ticP);
            break;
        case 75:
            result.type = TRT_EMPTY;
            DebugBreak();
            break;
        case 76:
            return Twapi_GetPerformanceInformation(ticP->interp);
        case 77:
            result.type = TRT_HWND;
            result.value.hwin = Twapi_GetNotificationWindow(ticP);
            break;
        case 78:                /* GetSystemWindowsDirectory */
        case 79:                /* GetWindowsDirectory */
        case 80:                /* GetSystemDirectory */
            result.type = TRT_UNICODE;
            result.value.unicode.str = u.buf;
            result.value.unicode.len =
                (func == 78
                 ? GetSystemWindowsDirectoryW
                 : (func == 79 ? GetWindowsDirectoryW : GetSystemDirectoryW)
                    ) (u.buf, ARRAYSIZE(u.buf));
            if (result.value.unicode.len >= ARRAYSIZE(u.buf) ||
                result.value.unicode.len == 0) {
                result.type = TRT_GETLASTERROR;
            }
            break;
            // 81 - UNUSED
        case 82:
            result.value.unicode.len = ARRAYSIZE(u.buf);
            if (GetDefaultPrinterW(u.buf, &result.value.unicode.len)) {
                result.value.unicode.len -= 1; /* Discard \0 */
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else {
                result.type = TRT_GETLASTERROR;
            }
            break;
        case 83:
            dw = GetTimeZoneInformation(&u.tzinfo);
            switch (dw) {
            case TIME_ZONE_ID_UNKNOWN:
            case TIME_ZONE_ID_STANDARD:
            case TIME_ZONE_ID_DAYLIGHT:
                objs[0] = Tcl_NewLongObj(dw);
                objs[1] = ObjFromTIME_ZONE_INFORMATION(&u.tzinfo);
                result.type = TRT_OBJV;
                result.value.objv.objPP = objs;
                result.value.objv.nobj = 2;
                break;
            default:
                result.type = TRT_GETLASTERROR;
                break;
            }
            break;
        }

        return TwapiSetResult(interp, &result);
    }

    if (func < 2000) {
        /* We should have exactly one additional argument. */

        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        
        switch (func) {
        case 1001:
            if (ObjToDWORD_PTR(interp, objv[2], &dwp) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_LPVOID;
            result.value.pv = (void *) dwp;
            break;
            // 1002 UNUSED
        case 1003:
            if (Tcl_GetDoubleFromObj(interp, objv[2], &u.d) != TCL_OK)
                return TCL_ERROR;
            result.type = VariantTimeToSystemTime(u.d, &result.value.systime) ?
                TRT_SYSTEMTIME : TRT_GETLASTERROR;
            break;
        case 1004:
            if (ObjToSYSTEMTIME(interp, objv[2], &systime) != TCL_OK)
                return TCL_ERROR;
            result.type = SystemTimeToVariantTime(&systime, &result.value.dval) ?
                TRT_DOUBLE : TRT_GETLASTERROR;
            break;
        case 1005: // canonicalize_guid
            /* Turn a GUID into canonical form */
            if (ObjToGUID(interp, objv[2], &result.value.guid) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_GUID;
            break;
        case 1006://UNUSED
        case 1007://UNUSED
        case 1008://UNUSED
        case 1009://UNUSED
            break;
        case 1010:
            if (ObjToPSECURITY_DESCRIPTOR(interp, objv[2], &secdP) != TCL_OK)
                return TCL_ERROR;
            // Note secdP may be NULL
            result.type = TRT_BOOL;
            result.value.bval = secdP ? IsValidSecurityDescriptor(secdP) : 0;
            if (secdP)
                TwapiFreeSECURITY_DESCRIPTOR(secdP);
            break;
        case 1011: // UNUSED
        case 1012: // UNUSED
        case 1013: // UNUSED
        case 1014:
            if (ObjToPACL(interp, objv[2], &daclP) != TCL_OK)
                return TCL_ERROR;
            // Note aclP may me NULL even on TCL_OK
            result.type = TRT_BOOL;
            result.value.bval = daclP ? IsValidAcl(daclP) : 0;
            if (daclP)
                TwapiFree(daclP);
            break;
        case 1015: // UNUSED
            break;
        case 1016:
            if (ObjToFILETIME(interp, objv[2], &u.filetime) != TCL_OK)
                return TCL_ERROR;
            if (FileTimeToSystemTime(&u.filetime, &result.value.systime))
                result.type = TRT_SYSTEMTIME;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1017:
            if (ObjToSYSTEMTIME(interp, objv[2], &systime) != TCL_OK)
                return TCL_ERROR;
            if (SystemTimeToFileTime(&systime, &result.value.filetime))
                result.type = TRT_FILETIME;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1018:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVOIDP(pv),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = Twapi_Wow64RevertWow64FsRedirection(pv);
            break;
        case 1019: // Twapi_IsValidGUID
            result.type = TRT_BOOL;
            result.value.bval = (ObjToGUID(NULL, objv[2], &guid) == TCL_OK);
            break;
        case 1020:
            if (ObjToTwapiId(interp, objv[2], &u.twapi_id) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            TwapiThreadPoolUnregister(ticP, u.twapi_id);
            break;
            // 1021 - UNUSED
        case 1022: // free
            if (ObjToLPVOID(interp, objv[2], &pv) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            if (pv)
                TwapiFree(pv);
            break;
        }
    } else {
        /* Free-for-all - each func responsible for checking arguments */
        switch (func) {
            // 10001-10007 UNUSED
        case 10008:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETHANDLE(h2),
                             GETHANDLE(h3), GETINT(dw), GETBOOL(dw2),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (DuplicateHandle(h, h2, h3, &result.value.hval, dw, dw2, dw3))
                result.type = TRT_HANDLE;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 10009:
            return Twapi_TclGetChannelHandle(interp, objc-2, objv+2);
        // 10010-16 UNUSED

        case 10017:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(secdP, ObjToPSECURITY_DESCRIPTOR),
                             GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
                return TCL_ERROR;
    
            if (ConvertSecurityDescriptorToStringSecurityDescriptorW(
                    secdP, dw, dw2, &s, &dw3)) {
                /* Cannot use TRT_UNICODE since buffer has to be freed */
                result.type = TRT_OBJ;
                /* Do not use dw3 as length because it seems to be size
                   of buffer, not string length as it includes padded nulls */
                result.value.obj = ObjFromUnicode(s);
                LocalFree(s);
            } else
                result.type = TRT_GETLASTERROR;
            if (secdP)
                TwapiFreeSECURITY_DESCRIPTOR(secdP);
            break;
#ifndef TWAPI_LEAN
        case 10018:
            return Twapi_LsaQueryInformationPolicy(interp, objc-2, objv+2);
        case 10019:
            return Twapi_LsaGetLogonSessionData(interp, objc-2, objv+2);
            // 10020-10030 UNUSED
#endif        
        case 10031: // CreateFile
            secattrP = NULL;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETINT(dw), GETINT(dw2),
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETINT(dw3), GETINT(dw4), GETHANDLE(h),
                             ARGEND) == TCL_OK) {
                result.type = TRT_VALID_HANDLE;
                result.value.hval = CreateFileW(s, dw, dw2, secattrP, dw3, dw4, h);
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            TwapiFreeSECURITY_ATTRIBUTES(secattrP); // Even in case of error or NULL
            break;
#ifndef TWAPI_LEAN
        case 10032: // SetNamedSecurityInfo
            osidP = gsidP = NULL;
            daclP = saclP = NULL;
            /* Note even in case of errors, sids and acls might have been alloced */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETINT(dw), GETINT(dw2),
                             GETVAR(osidP, ObjToPSID),
                             GETVAR(gsidP, ObjToPSID),
                             GETVAR(daclP, ObjToPACL),
                             GETVAR(saclP, ObjToPACL),
                             ARGEND) == TCL_OK) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = SetNamedSecurityInfoW(
                    s, dw, dw2, osidP, gsidP, daclP, saclP);
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            if (osidP) TwapiFree(osidP);
            if (gsidP) TwapiFree(gsidP);
            if (daclP) TwapiFree(daclP);
            if (saclP) TwapiFree(saclP);
            break;
#endif
            // 10033-35 UNUSED
        case 10036: // LookupPrivilegeName
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETVAR(luid, ObjToLUID),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (LookupPrivilegeNameW(s, &luid,
                                     u.buf, &result.value.unicode.len)) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 10037: // UNUSED
        case 10038: // UNUSED
        case 10039: // UNUSED
        case 10040: // UNUSED
            break;

#ifndef TWAPI_LEAN
        case 10041:
            /* Init to NULL as they may be partially init'ed on error
               and have to be freed */
            osidP = gsidP = NULL;
            daclP = saclP = NULL;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw), GETINT(dw2),
                             GETVAR(osidP, ObjToPSID),
                             GETVAR(gsidP, ObjToPSID),
                             GETVAR(daclP, ObjToPACL),
                             GETVAR(saclP, ObjToPACL),
                             ARGEND) == TCL_OK) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = SetSecurityInfo(
                    h, dw, dw2, osidP, gsidP, daclP, saclP);
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            if (osidP) TwapiFree(osidP);
            if (gsidP) TwapiFree(gsidP);
            if (daclP) TwapiFree(daclP);
            if (saclP) TwapiFree(saclP);
            break;
#endif
        case 10042: // UNUSED
        case 10043: // UNUSED
            break;
        case 10044:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw),
                             GETWSTRN(s, i), GETWSTRN(s2, i2),
                             GETINT(dw2), GETINT(dw3), GETBOOL(dw4),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (WTSSendMessageW(h, dw, s, sizeof(WCHAR)*i,
                                s2, sizeof(WCHAR)*i2, dw2,
                                dw3, &result.value.ival, dw4))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;    
            break;

        case 10045: // DuplicateTokenEx
            secattrP = NULL;        /* Even on error, it might be filled */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw),
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETINT(dw2), GETINT(dw3),
                             ARGEND) == TCL_OK) {
                if (DuplicateTokenEx(h, dw, secattrP, dw2, dw3, &result.value.hval))
                    result.type = TRT_HANDLE;
                else
                    result.type = TRT_GETLASTERROR;
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            TwapiFreeSECURITY_ATTRIBUTES(secattrP); // Even in case of error or NULL
            break;
        // 10046 UNUSED
        case 10047: // AdjustTokenPrivileges
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETBOOL(dw),
                             GETVAR(u.tokprivsP, ObjToPTOKEN_PRIVILEGES),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_AdjustTokenPrivileges(
                ticP, h, dw, u.tokprivsP);
            TwapiFreeTOKEN_PRIVILEGES(u.tokprivsP);
            break;
        case 10048: // PrivilegeCheck
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h),
                             GETVAR(u.tokprivsP, ObjToPTOKEN_PRIVILEGES),
                             GETBOOL(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (Twapi_PrivilegeCheck(h, u.tokprivsP, dw, &result.value.ival))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            TwapiFreeTOKEN_PRIVILEGES(u.tokprivsP);
            break;

        case 10049: // UNUSED
        case 10050: // UNUSED
        case 10051: // UNUSED
        case 10052: // UNUSED
        case 10053: // UNUSED
            break;
        case 10058: // DsGetDcName
            guidP = &guid;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETVAR(guidP, ObjToUUID_NULL),
                             GETNULLIFEMPTY(s3), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_DsGetDcName(interp, s, s2, guidP, s3, dw);
        // 10059 - 10069 UNUSED
        case 10070:
            return Twapi_GetNumberFormat(ticP, objc-2, objv+2);
        case 10071:
            return Twapi_GetCurrencyFormat(ticP, objc-2, objv+2);
        case 10072:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETINT(dw), GETBOOL(dw2), GETBOOL(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = InitiateSystemShutdownW(s, s2, dw, dw2, dw3);
            break;
        case 10073: // Twapi_FormatMessageFromModule
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), GETHANDLE(h), GETINT(dw2),
                             GETINT(dw3),
                             GETWARGV(u.wargv, ARRAYSIZE(u.wargv), dw4),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            /* Only look at select bits from dwFlags as others are used when
               formatting from string */
            dw &= FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_FROM_HMODULE;
            dw |=  FORMAT_MESSAGE_ARGUMENT_ARRAY;
            return TwapiFormatMessageHelper(interp, dw, h, dw2, dw3, dw4, u.wargv);
        case 10074: // Twapi_FormatMessageFromString
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), GETWSTR(s),
                             GETWARGV(u.wargv, ARRAYSIZE(u.wargv), dw4),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            /* Only look at select bits from dwFlags as others are used when
               formatting from module */
            dw &= FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK | FORMAT_MESSAGE_FROM_STRING;
            dw |=  FORMAT_MESSAGE_ARGUMENT_ARRAY;
            return TwapiFormatMessageHelper(interp, dw, s, 0, 0, dw4, u.wargv);
        case 10075: // WritePrivateProfileString
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETNULLTOKEN(s2), GETNULLTOKEN(s3),
                             GETWSTR(s4),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = WritePrivateProfileStringW(s, s2, s3, s4);
            break;
        case 10076:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLTOKEN(s), GETNULLTOKEN(s2), GETNULLTOKEN(s3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = WriteProfileStringW(s, s2, s3);
            break;
        case 10077:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), GETINT(dw2), GETVOIDP(pv), GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SystemParametersInfoW(dw, dw2, pv, dw3);
            break;
#ifndef TWAPI_LEAN
        case 10078:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw), GETWSTR(s),
                             GETNULLIFEMPTY(s2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_LoadUserProfile(interp, h, dw, s, s2);
        case 10079:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETHANDLE(h2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = UnloadUserProfile(h, h2);
            break;
#endif
        case 10080:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETBOOL(dw), GETBOOL(dw2), GETBOOL(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetSuspendState((BOOLEAN) dw, (BOOLEAN) dw2, (BOOLEAN) dw3);
            break;
        case 10081:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), ARGUSEDEFAULT, GETASTR(cP),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            NULLIFY_EMPTY(cP);
            return Twapi_GenerateWin32Error(interp, dw, cP);
        // 10082-97: // UNUSED
        case 10097:
            secattrP = NULL;        /* Even on error, it might be filled */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETBOOL(dw), GETNULLIFEMPTY(s),
                             ARGUSEDEFAULT, GETINT(dw2),
                             ARGEND) == TCL_OK) {

                result.type = TRT_HANDLE;                        
                result.value.hval = CreateMutexW(secattrP, dw, s);
                if (result.value.hval) {
                    if (dw2 & 1) {
                        /* Caller also wants indicator of whether object
                           already existed */
                        objs[0] = ObjFromHANDLE(result.value.hval);
                        objs[1] = Tcl_NewBooleanObj(GetLastError() == ERROR_ALREADY_EXISTS);
                        result.value.objv.objPP = objs;
                        result.value.objv.nobj = 2;
                        result.type = TRT_OBJV;
                    }
                }
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            TwapiFreeSECURITY_ATTRIBUTES(secattrP); // Even in case of error or NULL
            break;
        case 10098: // OpenMutex
        case 10099: // OpenSemaphore
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), GETBOOL(dw2), GETWSTR(s),
                             ARGUSEDEFAULT, GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            result.value.hval = (func == 10098 ? OpenMutexW : OpenSemaphoreW)
                (dw, dw2, s);
            if (result.value.hval) {
                if (dw3 & 1) {
                    /* Caller also wants indicator of whether object
                       already existed */
                    objs[0] = ObjFromHANDLE(result.value.hval);
                    objs[1] = Tcl_NewBooleanObj(GetLastError() == ERROR_ALREADY_EXISTS);
                    result.value.objv.objPP = objs;
                    result.value.objv.nobj = 2;
                    result.type = TRT_OBJV;
                }
            }
            break;
        case 10100:
            secattrP = NULL;        /* Even on error, it might be filled */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETINT(dw), GETINT(dw2), GETNULLIFEMPTY(s),
                             ARGEND) == TCL_OK) {
                result.type = TRT_HANDLE;
                result.value.hval = CreateSemaphoreW(secattrP, dw, dw2, s);
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            TwapiFreeSECURITY_ATTRIBUTES(secattrP); // Even in case of error or NULL
            break;
        case 10101: // DWORD
        case 10102: // BINARY
        case 10103: // CHARS
        case 10104: // UNICODE
        case 10105: // ADDRESS_LITERAL/POINTER
        case 10106: // Wide
            // We are passing the func code as well, hence only skip one arg
            return TwapiReadMemory(interp, objc-1, objv+1);
            // 10107-10109 UNUSED
        case 10110:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), ARGUSEDEFAULT, GETASTR(cP),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.value.opaque.p = TwapiAlloc(dw);
            result.value.opaque.name = cP[0] ? cP : "void*";
            result.type = TRT_OPAQUE;
            break;

        case 10111: // DWORD
        case 10112: // BINARY
        case 10113: // CHARS
        case 10114: // UNICODE
        case 10115: // ADDRESS_LITERAL/POINTER
        case 10116: // Wide
            // We are passing the func code as well, hence only skip one arg
            return TwapiWriteMemory(interp, objc-1, objv+1);

        case 10117: // UNUSED
        case 10118: // UNUSED
            break;

        case 10119: // IsEqualPtr
            if (objc != 4)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            if (ObjToOpaque(interp, objv[2], &pv, NULL) != TCL_OK ||
                ObjToOpaque(interp, objv[3], &pv2, NULL) != TCL_OK) {
                return TCL_ERROR;
            }
            result.type = TRT_BOOL;
            result.value.bval = (pv == pv2);
            break;
        case 10120: // IsNullPtr
            if (objc < 3 || objc > 4)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            cP = NULL;
            if (objc == 4) {
                cP = Tcl_GetString(objv[3]);
                NULLIFY_EMPTY(cP);
            }
            if (ObjToOpaque(interp, objv[2], &pv, cP) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = (pv == NULL);
            break;
        case 10121: // IsPtr
            if (objc < 3 || objc > 4)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            cP = NULL;
            if (objc == 4) {
                cP = Tcl_GetString(objv[3]);
                NULLIFY_EMPTY(cP);
            }
            result.type = TRT_BOOL;
            result.value.bval = (ObjToOpaque(interp, objv[2], &pv, cP) == TCL_OK);
            break;
        case 10122:
            secattrP = NULL;        /* Even on error, it might be filled */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETBOOL(dw), GETBOOL(dw2),
                             GETNULLIFEMPTY(s),
                             ARGEND) == TCL_OK) {
                h = CreateEventW(secattrP, dw, dw2, s);
                if (h) {
                    objs[1] = Tcl_NewBooleanObj(GetLastError() == ERROR_ALREADY_EXISTS); /* Do this before any other call */
                    objs[0] = ObjFromHANDLE(h);
                    result.type = TRT_OBJV;
                    result.value.objv.objPP = objs;
                    result.value.objv.nobj = 2;
                } else {
                    result.type = TRT_GETLASTERROR;
                }
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            TwapiFreeSECURITY_ATTRIBUTES(secattrP); // Even in case of error or NULL
            break;
            // 10123 - 10131 UNUSED
            // 10132 UNUSED
            // 10133 UNUSED
        case 10134: // TzLocalSpecificTimeToSystemTime
        case 10135: // SystemTimeToTzSpecificLocalTime
            if (objc == 3) {
                tzinfoP = NULL;
            } else if (objc == 4) {
                if (ObjToTIME_ZONE_INFORMATION(ticP->interp, objv[3], &u.tzinfo) != TCL_OK) {
                    return TCL_ERROR;
                }
                tzinfoP = &u.tzinfo;
            } else {
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            }
            if (ObjToSYSTEMTIME(interp, objv[2], &systime) != TCL_OK)
                return TCL_ERROR;
            if ((func == 10134 ? TzSpecificLocalTimeToSystemTime : SystemTimeToTzSpecificLocalTime) (tzinfoP, &systime, &result.value.systime))
                result.type = TRT_SYSTEMTIME;
            else
                result.type = TRT_GETLASTERROR;
            break;

        case 10136: // IsEqualGuid
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETGUID(guid), GETGUID(u.guid), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = IsEqualGUID(&guid, &u.guid);
            break;

        case 10137: // IMofCompiler_CompileBuffer
            return Twapi_IMofCompiler_CompileFileOrBuffer(ticP, 0, objc-2, objv+2);

        case 10138: // IMofCompiler_CompileFile
            return Twapi_IMofCompiler_CompileFileOrBuffer(ticP, 1, objc-2, objv+2);

        case 10139: // IMofCompiler_CreateBMOF
            return Twapi_IMofCompiler_CompileFileOrBuffer(ticP, 2, objc-2, objv+2);
        }
    }

    return TwapiSetResult(interp, &result);
}

int Twapi_CallUObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD dw, dw2, dw3;
    TwapiResult result;
    union {
        WCHAR buf[MAX_PATH+1];
        DWORD_PTR dwp;
        RPC_STATUS rpc_status;
        char *utf8;
        LPWSTR str;
        HANDLE h;
        SECURITY_ATTRIBUTES *secattrP;
        MemLifo *lifoP;
    } u;
    int func;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(func), GETINT(dw), ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 1000) {
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
            // 1-3 UNUSED
        case 4:
            result.value.hval = GetStdHandle(dw);
            if (result.value.hval == INVALID_HANDLE_VALUE)
                result.type = TRT_GETLASTERROR;
            else if (result.value.hval == NULL) {
                result.value.ival = ERROR_FILE_NOT_FOUND;
                result.type = TRT_EXCEPTION_ON_ERROR;
            } else
                result.type = TRT_HANDLE;
            break;
        case 5: // UNUSED
        case 6: // UNUSED
            break;
        case 7:
            result.value.unicode.len = VerLanguageNameW(dw, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
            result.value.unicode.str = u.buf;
            result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
            break;
            // 8-14 UNUSED
        case 15:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetComputerNameExW(dw, u.buf, &result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 16:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = Twapi_Wow64EnableWow64FsRedirection((BOOLEAN)dw);
        case 17:
            result.type = TRT_DWORD;
            result.value.ival = GetSystemMetrics(dw);
            break;
        case 18:
            result.type = TRT_EMPTY;
            Sleep(dw);
            break;
            // 19 - UNUSED
        case 20:
            return Twapi_EnumPrinters_Level4(interp, dw);
        case 21:
            u.rpc_status = UuidCreate(&result.value.uuid);
            /* dw is boolean indicating whether to allow strictly local uuids */
            if ((u.rpc_status == RPC_S_UUID_LOCAL_ONLY) && dw) {
                /* If caller does not mind a local only uuid, don't return error */
                u.rpc_status = RPC_S_OK;
            }
            result.type = u.rpc_status == RPC_S_OK ? TRT_UUID : TRT_GETLASTERROR;
            break;
        case 22:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetUserNameExW(dw, u.buf, &result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 23:
            result.value.ival = ImpersonateSelf(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
            // 24-32 UNUSED
        case 33:
            Sleep(dw);
            result.type = TRT_EMPTY;
            break;
        case 34:
            result.value.obj = Twapi_MapWindowsErrorToString(dw);
            result.type = TRT_OBJ;
            break;
        // 35 - 36 UNUSED
        case 37:
            u.lifoP = TwapiAlloc(sizeof(MemLifo));
            result.value.ival = MemLifoInit(u.lifoP, NULL, NULL, NULL, dw, 0);
            if (result.value.ival == ERROR_SUCCESS) {
                result.type = TRT_OPAQUE;
                result.value.opaque.p = u.lifoP;
                result.value.opaque.name = "MemLifo*";
            } else
                result.type = TRT_EXCEPTION_ON_ERROR;
            break;
        case 38:
            SetLastError(0);    /* As per MSDN */
            GlobalDeleteAtom((WORD)dw);
            result.value.ival = GetLastError();
            result.type = TRT_EXCEPTION_ON_ERROR;
            break;
        }
    } else if (func < 2000) {

        /* Check we have exactly one more integer argument */
        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        CHECK_INTEGER_OBJ(interp, dw2, objv[3]);
        switch (func) {
        case 1006: //GetLocaleInfo
            result.value.unicode.len = GetLocaleInfoW(dw, dw2, u.buf,
                                                      ARRAYSIZE(u.buf));
            if (result.value.unicode.len == 0) {
                result.type = TRT_GETLASTERROR;
            } else {
                if (dw2 & LOCALE_RETURN_NUMBER) {
                    // u.buf actually contains a number
                    result.value.ival = *(int *)u.buf;
                    result.type = TRT_DWORD;
                } else {
                    result.value.unicode.len -= 1;
                    result.value.unicode.str = u.buf;
                    result.type = TRT_UNICODE;
                }
            }
            break;

        case 1007:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ExitWindowsEx(dw, dw2);
            break;
        }
    } else if (func < 3000) {
        /* Check we have exactly two more integer arguments */
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
            // 2001-2003 UNUSED
        case 2004:
            result.type = TRT_BOOL;
            result.value.bval = AttachThreadInput(dw, dw2, dw3);
            break;
        }
    } else {
        /* Exactly one additional argument */
        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 10001:
            if (ObjToDWORD_PTR(interp, objv[3], &u.dwp) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HGLOBAL;
            result.value.hval = GlobalAlloc(dw, u.dwp);
            break;
        case 10002:
            result.type = TRT_DWORD;
            result.value.ival = LHashValOfName(dw, Tcl_GetUnicode(objv[3]));
            break;
        // 10003 UNUSED
        case 10004:
            if (ObjToHANDLE(interp, objv[3], &u.h) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetStdHandle(dw, u.h);
            break;
        }
    }

    return TwapiSetResult(interp, &result);
}

int Twapi_CallSObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR arg;
    TwapiResult result;
    union {
        WCHAR buf[MAX_PATH+1];
        LARGE_INTEGER largeint;
        SOCKADDR_STORAGE ss;
        PSID sidP;
    } u;
    int func;                   /* What function to call */
    DWORD dw, dw2, dw3;
    SECURITY_DESCRIPTOR *secdP;
    LSA_OBJECT_ATTRIBUTES lsa_oattr;
    LSA_UNICODE_STRING lsa_ustr; /* Used with lsa_oattr so not in union */
    WCHAR *bufP;

    result.type = TRT_BADFUNCTIONCODE;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETWSTR(arg),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    /* One argument commands - assign codes 1-999 */
    if (func < 500) {
        switch (func) {
            // 1-7 UNUSED
        case 8:
            return Twapi_CommandLineToArgv(interp, arg);
        //9- 10 - UNUSED
        case 11:
            return Twapi_AppendLog(interp, arg);
        case 12:
            result.type = TRT_HANDLE;
            result.value.hval = WTSOpenServerW(arg);
            break;
        // 13-16 UNUSED
        case 17:
            result.type = TRT_EXCEPTION_ON_FALSE;
            NULLIFY_EMPTY(arg);
            result.value.ival = AbortSystemShutdownW(arg);
        case 18:
            NULLIFY_EMPTY(arg);
            return Twapi_GetPrivateProfileSectionNames(ticP, arg);
            // 19-21 UNUSED
        case 22:
            bufP = u.buf;
            dw = ExpandEnvironmentStringsW(arg, bufP, ARRAYSIZE(u.buf));
            if (dw > ARRAYSIZE(u.buf)) {
                // Need a bigger buffer
                bufP = TwapiAlloc(dw * sizeof(WCHAR));
                dw2 = dw;
                dw = ExpandEnvironmentStringsW(arg, bufP, dw2);
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
            if (bufP != u.buf)
                TwapiFree(bufP);
            break;
        case 23: // GlobalAddAtom
            result.value.ival = GlobalAddAtomW(arg);
            result.type = result.value.ival ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        // 24-26 UNUSED
        case 27: // IsValidSidSyntax
            u.sidP = NULL;
            result.type = TRT_BOOL;
            result.value.bval = ConvertStringSidToSidW(arg, &u.sidP);
            if (u.sidP)
                LocalFree(u.sidP);
            break;
        }
    } else if (func < 1000) {
        /* One additional integer argument */

        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[3]);
        switch (func) {
        case 501:
            if (ConvertStringSecurityDescriptorToSecurityDescriptorW(
                    arg, dw, &secdP, NULL)) {
                result.value.obj = ObjFromSECURITY_DESCRIPTOR(interp, secdP);
                if (secdP)
                    LocalFree(secdP);
                if (result.value.obj)
                    result.type = TRT_OBJ;
                else
                    return TCL_ERROR;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 502: // LsaOpenPolicy
            ObjToLSA_UNICODE_STRING(objv[2], &lsa_ustr);
            TwapiZeroMemory(&lsa_oattr, sizeof(lsa_oattr));
            dw2 = LsaOpenPolicy(&lsa_ustr, &lsa_oattr, dw, &result.value.hval);
            if (dw2 == STATUS_SUCCESS) {
                result.type = TRT_LSA_HANDLE;
            } else {
                result.type = TRT_NTSTATUS;
                result.value.ival = dw2;
            }
            break;
        // 503 UNUSED
        case 504: // LoadLibrary
            result.type = TRT_HANDLE;
            result.value.hval = LoadLibraryExW(arg, NULL, dw);
            break;
        }
    } else if (func < 2000) {

        /* Commands with exactly two additional integer argument */
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        
        switch (func) {
        case 1004:
            return Twapi_GetNamedSecurityInfo(interp, arg, dw, dw2);
        case 1005:
            bufP = u.buf;
            dw3 = ARRAYSIZE(u.buf);
            if (! TranslateNameW(arg, dw, dw2, bufP, &dw3)) {
                result.value.ival = GetLastError();
                if (result.value.ival != ERROR_INSUFFICIENT_BUFFER) {
                    result.type = TRT_EXCEPTION_ON_ERROR;
                    result.value.ival = GetLastError();
                    MemLifoPopFrame(&ticP->memlifo);
                    break;
                }
                /* Retry with larger buffer */
                bufP = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR)*dw3,
                                        &dw3);
                dw3 /= sizeof(WCHAR);
                if (! TranslateNameW(arg, dw, dw2, bufP, &dw3)) {
                    result.type = TRT_EXCEPTION_ON_ERROR;
                    result.value.ival = GetLastError();
                    MemLifoPopFrame(&ticP->memlifo);
                    break;
                }
            }

            result.value.unicode.str = bufP;
            result.value.unicode.len = dw3 - 1 ;
            result.type = TRT_UNICODE;
            if (bufP != u.buf)
                MemLifoPopFrame(&ticP->memlifo);
            break;
        }
    }

    return TwapiSetResult(interp, &result);
}


int Twapi_CallHObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE h, h2;
    DWORD dw, dw2;
    TwapiResult result;
    union {
        COORD coord;
        WCHAR buf[MAX_PATH+1];
        TWAPI_TOKEN_MANDATORY_POLICY ttmp;
        TWAPI_TOKEN_MANDATORY_LABEL ttml;
        LSA_UNICODE_STRING lsa_ustr;
        SECURITY_ATTRIBUTES *secattrP;
        MemLifo *lifoP;
    } u;
    int func;
    int i;
    void *pv;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETHANDLE(h),
                     ARGTERM) != TCL_OK) {
        return TCL_ERROR;
    }

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 1000) {
        /* Command with a single handle argument */
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 14:
            result.type = GetHandleInformation(h, &result.value.ival)
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 15:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FreeLibrary(h);
            break;
        case 16:
            result.type = GetDevicePowerState(h, &result.value.bval)
                ? TRT_BOOL : TRT_GETLASTERROR;
            break;
            // 17-25 UNUSED
        case 26:
            result.value.ival = LsaClose(h);
            result.type = TRT_DWORD;
            break;
        case 27:
            result.value.ival = ImpersonateLoggedOnUser(h);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        // 28-41 UNUSED
#ifndef TWAPI_LEAN
        case 42:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ReleaseMutex(h);
            break;
#endif
        case 43:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseHandle(h);
            break;
        case 44:
            result.type = TRT_HANDLE;
            result.value.hval = h;
            break;
        case 45:
            result.type = TRT_EXCEPTION_ON_ERROR;
            /* GlobalFree will return a HANDLE on failure. */
            result.value.ival = GlobalFree(h) ? GetLastError() : 0;
            break;
        case 46:
            result.type = TRT_EXCEPTION_ON_ERROR;
            /* GlobalUnlock is an error if it returns 0 AND GetLastError is non-0 */
            result.value.ival = GlobalUnlock(h) ? 0 : GetLastError();
            break;
        case 47:
            result.type = TRT_DWORD_PTR;
            result.value.dwp = GlobalSize(h);
            break;
        case 48:
            result.type = TRT_NONNULL_LPVOID;
            result.value.pv = GlobalLock(h);
            break;
        case 49:
            return Twapi_WTSEnumerateSessions(interp, h);
        case 50:
            return Twapi_WTSEnumerateProcesses(interp, h);
        case 51:
            /* h == NULL -> current server and does not need to be closed */
            if (h)
                WTSCloseServer(h);
            result.type = TRT_EMPTY;
            break;
            // 52-53 UNUSED
        case 54:
            MemLifoClose(h);
            TwapiFree(h);
            result.type = TRT_EMPTY;
            break;
        case 55:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = MemLifoPopFrame((MemLifo *)h);
            break;
        // 56-58 UNUSED
        case 59:
            /* Find the required buffer size */
            dw = GetObject(h, 0, NULL);
            if (dw == 0) {
                result.type = TRT_GETLASTERROR; /* TBD - is GetLastError set? */
                break;
            }
            result.value.obj = Tcl_NewByteArrayObj(NULL, dw); // Alloc storage
            pv = Tcl_GetByteArrayFromObj(result.value.obj, &dw); // and get ptr to it
            dw = GetObject(h, dw, pv);
            if (dw == 0)
                result.type = TRT_GETLASTERROR;
            else {
                Tcl_SetByteArrayLength(result.value.obj, dw);
                result.type = TRT_OBJ;
            }
            break;
        case 60:
            result.type = TRT_OPAQUE;
            result.value.opaque.p = MemLifoPushMark(h);
            result.value.opaque.name = "MemLifoMark*";
            break;
        case 61:
            result.type = TRT_DWORD;
            result.value.ival = MemLifoPopMark(h);
            break;
        case 62:
            result.type = TRT_DWORD;
            result.value.ival = MemLifoValidate(h);
            break;
        case 63:
            return Twapi_MemLifoDump(ticP, h);
        case 64:
            /* This could be in the namedpip package, but we leave it
               here for now.
            */
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ImpersonateNamedPipeClient(h);
            break;
        // 65 - UNUSED
        case 66:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetEvent(h);
            break;
        case 67:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ResetEvent(h);
            break;
        }
    } else if (func < 2000) {

        // A single additional DWORD arg is present
        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[3]);

        switch (func) {
#ifndef TWAPI_LEAN
        case 1001:
            result.type = ReleaseSemaphore(h, dw, &result.value.ival) ?
                TRT_DWORD : TRT_GETLASTERROR;
            break;
#endif
            // 1002-1004 UNUSED
        case 1005:
            result.type = OpenProcessToken(h, dw, &result.value.hval) ?
                TRT_HANDLE : TRT_GETLASTERROR;
            break;
        case 1006:
            return Twapi_GetTokenInformation(interp, h, dw);
#ifndef TWAPI_LEAN
        case 1007:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetTokenInformation(h,
                                                    TwapiTokenVirtualizationEnabled,
                                                    &dw, sizeof(dw));
            break;
        case 1008:
            u.ttmp.Policy = dw;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetTokenInformation(h,
                                                    TwapiTokenMandatoryPolicy,
                                                    &u.ttmp, sizeof(u.ttmp));
            break;
#endif // TWAPI_LEAN
        // 1009-1015 UNUSED
        case 1016:
            result.type = TRT_DWORD;
            result.value.ival = GetDeviceCaps(h, dw);
            break;
        case 1017:
            result.value.ival = WaitForSingleObject(h, dw);
            if (result.value.ival == (DWORD) -1) {
                result.type = TRT_GETLASTERROR;
            } else {
                result.type = TRT_DWORD;
            }
            break;
        case 1018:
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoAlloc(h, dw, NULL);
            break;
        case 1019:
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoPushFrame(h, dw, NULL);
            break;
        }
    } else if (func < 3000) {

        // Two additional DWORD args present
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 2001:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = WTSDisconnectSession(h, dw, dw2);
            break;
        case 2002:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = WTSLogoffSession(h, dw, dw2);
            break;
        case 2003:
            return Twapi_WTSQuerySessionInformation(interp, h, dw, dw2);
#ifndef TWAPI_LEAN
        case 2004:
            return Twapi_GetSecurityInfo(interp, h, dw, dw2);
#endif            
        case 2005:
            result.type = OpenThreadToken(h, dw, dw2, &result.value.hval) ?
                TRT_HANDLE : TRT_GETLASTERROR;
            break;
        case 2006: // UNUSED
            break;
        case 2007:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetHandleInformation(h, dw, dw2);
            break;
        case 2008: 
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoExpandLast(h, dw, dw2);
            break;
        case 2009: 
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoShrinkLast(h, dw, dw2);
            break;
        case 2010: 
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoResizeLast(h, dw, dw2);
            break;
        case 2011:
            return TwapiThreadPoolRegister(
                ticP, h, dw, dw2, TwapiCallRegisteredWaitScript, NULL);
        }
    } else {
        /* Arbitrary additional arguments */

        switch (func) {
#ifndef TWAPI_LEAN
        case 10002: // SetThreadToken
            if (objc != 4)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            
            if (ObjToHANDLE(interp, objv[3], &h2) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetThreadToken(h, h2);
            break;
        case 10003:
            if (objc != 4)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            ObjToLSA_UNICODE_STRING(objv[3], &u.lsa_ustr);
            return Twapi_LsaEnumerateAccountsWithUserRight(interp, h,
                                                           &u.lsa_ustr);
#endif
        case 10004:
            if (objc != 4)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            if (ObjToSID_AND_ATTRIBUTES(interp, objv[3], &u.ttml.Label) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetTokenInformation(h,
                                                    TwapiTokenIntegrityLevel,
                                                    &u.ttml, sizeof(u.ttml));
            if (u.ttml.Label.Sid)
                TwapiFree(u.ttml.Label.Sid);
            break;
        }
    }
    return TwapiSetResult(interp, &result);
}


/* Call S1 S2 S3 DWORD
 * Note - s1 will be passed on as NULL if empty string (LPWSTR_NULL_IF_EMPTY
 * semantics). S2,S3,DWORD default to "", "", 0 respectively if not passed in.
*/
int Twapi_CallSSSDObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s1, s2, s3;
    DWORD   dw, dw2;
    TwapiResult result;
    WCHAR buf[MAX_PATH+1];

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETNULLIFEMPTY(s1), ARGUSEDEFAULT,
                     GETWSTR(s2), GETWSTR(s3), GETINT(dw),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_LookupAccountName(interp, s1, s2);
    // 2-4 - UNUSED
    case 5:
        if (objc != 7)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw2, objv[6]);
        EMPTIFY_NULL(s1);      /* Username must not be NULL - pass as "" */
        NULLIFY_EMPTY(s2);      /* Domain - NULL if empty string */
        if (LogonUserW(s1, s2, s3, dw,dw2, &result.value.hval))
            result.type = TRT_HANDLE;
        else
            result.type = TRT_GETLASTERROR;
        break;
        // 6 UNUSED
    case 7:
        result.type = TRT_DWORD;
        EMPTIFY_NULL(s1);      /* Undo s1 LPWSTR_NULL_IF_EMPTY semantics */
        result.value.ival = GetPrivateProfileIntW(s1, s2, dw, s3);
        break;
    case 8:
        result.type = TRT_DWORD;
        result.value.ival = GetProfileIntW(s2, s3, dw);
        break;
    case 9:
        // Note s2,s1, not s1,s2
        return Twapi_GetPrivateProfileSection(ticP, s2, s1);
    case 13:
        result.value.unicode.len = ARRAYSIZE(buf);
        if (LookupPrivilegeDisplayNameW(s1,s2,buf,&result.value.unicode.len,&dw)) {
            result.value.unicode.str = buf;
            result.type = TRT_UNICODE;
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 14:
        if (LookupPrivilegeValueW(s1,s2,&result.value.luid))
            result.type = TRT_LUID;
        else
            result.type = TRT_GETLASTERROR;
        break;
    // 15-33 UNUSED
    case 34:
        NULLIFY_EMPTY(s2);
        return Twapi_NetGetDCName(interp, s1,s2);
    // 35-40: UNUSED
    case 41:
        return Twapi_QueryDosDevice(ticP, s1);
    case 42:
        result.type = TRT_EXCEPTION_ON_FALSE;
        // Note we use s2, s3 here as s1 has LPWSTR_NULL_IF_EMPTY semantics
        NULLIFY_EMPTY(s3);
        result.value.ival = DefineDosDeviceW(dw, s2, s3);
        break;
    }

    return TwapiSetResult(interp, &result);
}


int Twapi_CallWUObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HWND hwnd;
    TwapiResult result;
    int func;
    DWORD dw, dw2, dw3;
    DWORD_PTR dwp, dwp2;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETHANDLET(hwnd, HWND), GETINT(dw),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 1000) {
        switch (func) {
        case 8:
            SetLastError(0);    /* Avoid spurious errors when checking GetLastError */
            result.value.dwp = GetWindowLongPtrW(hwnd, dw);
            if (result.value.dwp || GetLastError() == 0)
                result.type = TRT_DWORD_PTR;
            else
                result.type = TRT_GETLASTERROR;
            break;
        }
    } else if (func < 2000) {
        // HWIN UINT WPARAM LPARAM ?ARGS?
        if (TwapiGetArgs(interp, objc-4, objv+4,
                         GETDWORD_PTR(dwp), GETDWORD_PTR(dwp2),
                         ARGUSEDEFAULT,
                         GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 1001:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = PostMessageW(hwnd, dw, dwp, dwp2);
            break;
        case 1002:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SendNotifyMessageW(hwnd, dw, dwp, dwp2);
            break;
        case 1003:
            if (SendMessageTimeoutW(hwnd, dw, dwp, dwp2, dw2, dw3, &result.value.dwp))
                result.type = TRT_DWORD_PTR;
            else {
                /* On some systems, GetLastError() returns 0 on timeout */
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = GetLastError();
                if (result.value.ival == 0)
                    result.value.ival = ERROR_TIMEOUT;
            }
        }
    } else {
        /* Aribtrary *additional* arguments */
        switch (func) {
        case 10003:
            if (objc != 5)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            if (ObjToDWORD_PTR(interp, objv[4], &dwp) != TCL_OK)
                return TCL_ERROR;
            result.type = Twapi_SetWindowLongPtr(hwnd, dw, dwp, &result.value.dwp)
                ? TRT_DWORD_PTR : TRT_GETLASTERROR;
            break;
        }            
    }

    return TwapiSetResult(interp, &result);
}


/*
 * PSID - based calls. These are handled separately because we want to
 * ensure we do PSID conversion via ObjToPSID before any object
 * shimmering takes place. Otherwise we could have used one of the
 * string based call dispatchers Also, there is other special
 * processing required such as freeing memory.
*/
int Twapi_CallPSIDObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    PSID sidP = NULL;
    int func;
    LPCWSTR s;
    HANDLE h;
    union {
        TOKEN_PRIMARY_GROUP tpg;
        TOKEN_OWNER towner;
        DWORD dw;
    } u;
    TwapiResult result;
    LSA_UNICODE_STRING *lsa_strings;
    ULONG  lsa_count;

    func = 0;
    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func),
                     ARGSKIP,
                     GETVAR(sidP, ObjToPSID),
                     ARGTERM) != TCL_OK) {
        if (sidP)
            TwapiFree(sidP);         /* Might be alloc'ed even on fail */
        /* Special case - IsValidSid check. */
        if (func == 1 &&
            objc == 4) {
            /* Problem was with SID. IsValidSid return 0 with TCL_OK */
            Tcl_SetObjResult(interp, Tcl_NewIntObj(0));
            return TCL_OK;
        } else
            return TCL_ERROR;
    }
    /* sidP may legitimately be NULL, else it points to a Twapialloc'ed block */

    result.type = TRT_EXCEPTION_ON_FALSE; /* Likely result type */
    if (func < 1000) {
        switch (func) {
        case 1:
            /* The ObjToPSID above would have already checked SID validity */
            result.type = TRT_BOOL;
            result.value.bval = 1;
            break;
        case 2:
            s = ObjToLPWSTR_NULL_IF_EMPTY(objv[2]);
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_LookupAccountSid(interp, s, sidP);
            break;
        }
    } else if (func < 2000) {
        /* Codes expecting HANDLE as first parameter */
        if (ObjToHANDLE(interp, objv[2], &h) != TCL_OK) {
            result.type = TRT_TCL_RESULT;
            result.value.ival = TCL_ERROR;
        } else {
            switch (func) {
            case 1002:
                result.type = CheckTokenMembership(h, sidP, &result.value.bval)
                    ? TRT_BOOL : TRT_GETLASTERROR;
                break;
#ifndef TWAPI_LEAN
            case 1003:
                u.tpg.PrimaryGroup = sidP;
                result.type = TRT_EXCEPTION_ON_FALSE;
                result.value.ival = SetTokenInformation(h, TokenPrimaryGroup, &u.tpg, sizeof(u.tpg));
                break;
            case 1004:
                u.towner.Owner = sidP;
                result.type = TRT_EXCEPTION_ON_FALSE;
                result.value.ival = SetTokenInformation(h, TokenOwner,
                                                        &u.towner,
                                                        sizeof(u.towner));
                break;
            case 1005:
                result.type = TRT_TCL_RESULT;
                result.value.ival = Twapi_LsaEnumerateAccountRights(interp,
                                                                    h, sidP);
                break;
            case 1006:
                if (objc != 6)
                    return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
                /* FALLTHRU */
            case 1007:
                if (objc != 6 && objc != 5)
                    return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

                result.value.ival = ObjToLSASTRINGARRAY(interp, objv[func == 1006 ? 5 : 4], &lsa_strings, &lsa_count);
                if (result.value.ival != TCL_OK) {
                    result.type = TRT_TCL_RESULT;
                    break;
                }
                result.type = TRT_NTSTATUS;
                if (func == 1006) {
                    CHECK_INTEGER_OBJ(interp, u.dw, objv[4]);
                    result.value.ival = LsaRemoveAccountRights(
                        h, sidP, (BOOLEAN) (u.dw ? 1 : 0),
                                                               lsa_strings, lsa_count);
                } else {
                    result.value.ival = LsaAddAccountRights(h, sidP,
                                                            lsa_strings, lsa_count);
                }
                TwapiFree(lsa_strings);
                break;
#endif
            }
        }
    }


    if (sidP)
        TwapiFree(sidP);

    return TwapiSetResult(interp, &result);
}


