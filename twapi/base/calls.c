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
    "proc twapi::NetUserGetInfo {system account level} {\n"
    "    return [twapi::CallSSSD 10 $system $account {} $level]\n"
    "}\n"
    "proc twapi::NetGroupGetInfo {system account level} {\n"
    "    return [twapi::CallSSSD 11 $system $account {} $level]\n"
    "}\n"
    "proc twapi::NetLocalGroupGetInfo {system account level} {\n"
    "    return [twapi::CallSSSD 12 $system $account {} $level]\n"
    "}\n"
    "proc twapi::OpenSCManager {{machine {}} {database {}} {access 0x000F003F}} { \n"
    "    # 0x000F003f -> SC_MANAGER_ALL_ACCESS\n"
    "    return [CallSSSD 23 $machine $database {} $access]\n"
    "}\n"
    "proc twapi::NetUseGetInfo {server share level} {\n"
    "    return [CallSSSD 24 $server $share {} $level]\n"
    "}\n"
    "proc twapi::NetShareDel {server share flags} {                          \n"
    "    return [CallSSSD 25 $server $share {} $flags]\n"
    "}\n"
    "proc twapi::NetShareGetInfo {server share level} {\n"
    "    return [CallSSSD 27 $server $share {} $level]\n"
    "}\n"
    "proc twapi::Twapi_WNetGetResourceInformation {remotename provider restype} { \n"
    "    # Note first two args are interchanged\n"
    "    return [CallSSSD 36 $provider $remotename {} $restype]\n"
    "}\n"
    "proc twapi::Twapi_WriteUrlShortcut {path url flags} {\n"
    "    return [CallSSSD 37 $path $url {} $flags]\n"
    "}\n"
    "proc twapi::DefineDosDevice  {flags devname path} {\n"
    "    return [CallSSSD 42 {} $devname $path $flags]\n"
    "}\n"
    "proc twapi::SetVolumeMountPoint  {vol name} {\n"
    "    return [CallSSSD 43 {} $vol $name]\n"
    "}\n"
    "proc twapi::CreateScalableFontResource  {hidden fontres fontfile curpath} {\n"
    "    return [CallSSSD 47 $fontres $fontfile $curpath $hidden]\n"
    "}\n"
    "proc twapi::RemoveFontResourceEx  {filename flags} {\n"
    "    return [CallSSSD 48 {} $filename {} $flags]\n"
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
    "proc twapi::Twapi_NetUserSetInfoDWORD {func server name level} {\n"
    "    return [CallSSSD $func $server $name {} $level];\n"
    "}\n"
    "proc twapi::Twapi_NetUserSetInfoLPWSTR {func server name value} {\n"
    "    return [CallSSSD $func $server $name $value 0];\n"
    "}\n"
    "proc twapi::IsValidSid {sid} {\n"
    "    return [CallPSID 1 {} $sid]\n"
    "}\n"
    ;

DWORD twapi_netenum_bufsize = MAX_PREFERRED_LENGTH; /* Defined by LANMAN */

int TwapiGetArgs(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[],
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
                TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
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
                TwapiReturnTwapiError(interp, "Value does not fit in 16 bits.", TWAPI_INVALID_ARGS);
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
            TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
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


static TwapiMakeCallAlias(Tcl_Interp *interp, char *fn, char *callcmd, char *code)
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
    Tcl_CreateObjCommand(interp, "twapi::CallHSU", Twapi_CallHSUObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallSSSD", Twapi_CallSSSDObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallW", Twapi_CallWObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallWU", Twapi_CallWUObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallPSID", Twapi_CallPSIDObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallNetEnum", Twapi_CallNetEnumObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallPdh", Twapi_CallPdhObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CallCOM", Twapi_CallCOMObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        TwapiMakeCallAlias(interp, "twapi::" #fn_, "twapi::" #call_, # code_); \
    } while (0);

    /*
     * NOTE: Some MISSING CALL NUMBERS are defined in win32calls.tcl
     * or as procs in the apiprocs global.
     */
    CALL_(GetCurrentProcess, Call, 1);
    CALL_(CloseClipboard, Call, 2);
    CALL_(EmptyClipboard, Call, 3);
    CALL_(GetOpenClipboardWindow, Call, 4);
    CALL_(GetDesktopWindow, Call, 5);
    CALL_(GetShellWindow, Call, 6);
    CALL_(GetForegroundWindow, Call, 7);
    CALL_(GetActiveWindow, Call, 8);
    CALL_(GetClipboardOwner, Call, 9);
    CALL_(Twapi_EnumClipboardFormats,Call, 10);
    CALL_(AllocConsole, Call, 11);
    CALL_(FreeConsole, Call, 12);
    CALL_(GetConsoleCP, Call, 13);
    CALL_(GetConsoleOutputCP, Call, 14);
    CALL_(GetNumberOfConsoleMouseButtons, Call, 15);
    CALL_(GetConsoleTitle, Call, 16);
    CALL_(GetConsoleWindow, Call, 17);
    CALL_(GetLogicalDrives, Call, 18);
    CALL_(GetNetworkParams, Call, 19);
    CALL_(GetAdaptersInfo, Call, 20);
    CALL_(GetInterfaceInfo, Call, 21);
    CALL_(GetNumberOfInterfaces, Call, 22);
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
    CALL_(EnumProcesses, Call, 43);
    CALL_(EnumDeviceDrivers, Call, 44);
    CALL_(GetCurrentThreadId, Call, 45);
    CALL_(GetCurrentThread, Call, 46);
    CALL_(GetCommandLineW, Call, 47);
    CALL_(AllocateLocallyUniqueId, Call, 48);
    CALL_(LockWorkStation, Call, 49);
    CALL_(LsaEnumerateLogonSessions, Call, 50);
    CALL_(RevertToSelf, Call, 51);
    CALL_(Twapi_InitializeSecurityDescriptor, Call, 52);
    CALL_(EnumerateSecurityPackages, Call, 53);
    CALL_(IsThemeActive, Call, 54);
    CALL_(IsAppThemed, Call, 55);
    CALL_(GetCurrentThemeName, Call, 56);
    CALL_(Twapi_GetShellVersion, Call, 57);
    CALL_(GetDoubleClickTime, Call, 58);
    CALL_(EnumWindowStations, Call, 59);
    CALL_(GetProcessWindowStation, Call, 60);
    CALL_(GetCursorPos, Call, 61);
    CALL_(GetCaretPos, Call, 62);
    CALL_(GetCaretBlinkTime, Call, 63);
    CALL_(EnumWindows, Call, 64);
    CALL_(Twapi_StopConsoleEventNotifier, Call, 65);
    CALL_(FindFirstVolume, Call, 66);
    CALL_(GetLastInputInfo, Call, 67);
    CALL_(GetSystemPowerStatus, Call, 68);
    CALL_(Twapi_ClipboardMonitorStart, Call, 69);
    CALL_(Twapi_ClipboardMonitorStop, Call, 70);
    CALL_(Twapi_PowerNotifyStart, Call, 71);
    CALL_(Twapi_PowerNotifyStop, Call, 72);
    CALL_(Twapi_StartConsoleEventNotifier, Call, 73);
    CALL_(TwapiId, Call, 74);
    CALL_(DebugBreak, Call, 75);

    CALL_(Twapi_AddressToPointer, Call, 1001);
    CALL_(FlashWindowEx, Call, 1002);
    CALL_(VariantTimeToSystemTime, Call, 1003);
    CALL_(SystemTimeToVariantTime, Call, 1004);
    CALL_(GetDeviceDriverBaseName, Call, 1005);
    CALL_(GetDeviceDriverFileName, Call, 1006);
    CALL_(GetBestInterface, Call, 1007);
    CALL_(QuerySecurityContextToken, Call, 1008);
    CALL_(WindowFromPoint, Call, 1009);
    CALL_(IsValidSecurityDescriptor, Call, 1010);
    CALL_(FreeCredentialsHandle, Call, 1011);
    CALL_(DeleteSecurityContext, Call, 1012);
    CALL_(ImpersonateSecurityContext, Call, 1013);
    CALL_(IsValidAcl, Call, 1014);
    CALL_(SetTcpEntry, Call, 1015);
    CALL_(FileTimeToSystemTime, Call, 1016);
    CALL_(SystemTimeToFileTime, Call, 1017);
    CALL_(Wow64RevertWow64FsRedirection, Call, 1018);
    CALL_(Twapi_IsValidGUID, Call, 1019);
    CALL_(Twapi_UnregisterWaitOnHandle, Call, 1020);
    CALL_(TwapiGetThemeDefine, Call, 1021);
    CALL_(free, Call, 1022);

    CALL_(Twapi_FormatExtendedTcpTable, Call, 10000);
    CALL_(Twapi_FormatExtendedUdpTable, Call, 10001);
    CALL_(GetExtendedTcpTable, Call, 10002);
    CALL_(GetExtendedUdpTable, Call, 10003);
    CALL_(Twapi_ResolveAddressAsync, Call, 10004);
    CALL_(Twapi_ResolveHostnameAsync, Call, 10005);
    CALL_(getaddrinfo, Call, 10006);
    CALL_(getnameinfo, Call, 10007);
    CALL_(DuplicateHandle, Call, 10008);
    CALL_(Tcl_GetChannelHandle, Call, 10009);
    CALL_(MonitorFromPoint, Call, 10010);
    CALL_(MonitorFromRect, Call, 10011);
    CALL_(EnumDisplayDevices, Call, 10012);
    CALL_(ReportEvent, Call, 10013);
    CALL_(Twapi_RegisterDeviceNotification, Call, 10014);
    CALL_(CreateProcess, Call, 10015);
    CALL_(CreateProcessAsUser, Call, 10016);
    CALL_(ConvertSecurityDescriptorToStringSecurityDescriptor, Call, 10017);
    CALL_(LsaQueryInformationPolicy, Call, 10018);
    CALL_(LsaGetLogonSessionData, Call, 10019);
    CALL_(AcquireCredentialsHandle, Call, 10020);
    CALL_(InitializeSecurityContext, Call, 10021);
    CALL_(AcceptSecurityContext, Call, 10022);
    CALL_(QueryContextAttributes, Call, 10023);
    CALL_(MakeSignature, Call, 10024);
    CALL_(EncryptMessage, Call, 10025);
    CALL_(VerifySignature, Call, 10026);
    CALL_(DecryptMessage, Call, 10027);
    CALL_(CryptAcquireContext, Call, 10028);
    CALL_(CryptReleaseContext, Call, 10029);
    CALL_(CryptGenRandom, Call, 10030);
    CALL_(CreateFile, Call, 10031);
    CALL_(SetNamedSecurityInfo, Call, 10032);
    CALL_(CreateWindowStation, Call, 10033);
    CALL_(OpenDesktop, Call, 10034);
    CALL_(Twapi_RegisterDirectoryMonitor, Call, 10035);
    CALL_(LookupPrivilegeName, Call, 10036);
    CALL_(PlaySound, Call, 10037);
    CALL_(SetConsoleWindowInfo, Call, 10038);
    CALL_(FillConsoleOutputAttribute, Call, 10039);
    CALL_(EnumServicesStatusEx, Call, 10040);
    CALL_(SetSecurityInfo, Call, 10041);
    CALL_(ScrollConsoleScreenBuffer, Call, 10042);
    CALL_(WriteConsoleOutputCharacter, Call, 10043);
    CALL_(WTSSendMessage, Call, 10044);
    CALL_(DuplicateTokenEx, Call, 10045);
    CALL_(ReadProcessMemory, Call, 10046);
    CALL_(Twapi_AdjustTokenPrivileges, Call, 10047);
    CALL_(Twapi_PrivilegeCheck, Call, 10048);
    CALL_(ChangeServiceConfig, Call, 10049);
    CALL_(CreateService, Call, 10050);
    CALL_(StartService, Call, 10051);
    CALL_(SetConsoleCursorPosition, Call, 10052);
    CALL_(SetConsoleScreenBufferSize, Call, 10053);
    CALL_(GetModuleFileNameEx, Call, 10054);
    CALL_(GetModuleBaseName, Call, 10055);
    CALL_(GetModuleInformation, Call, 10056);
    CALL_(Twapi_VerQueryValue_STRING, Call, 10057);
    CALL_(DsGetDcName, Call, 10058);
    CALL_(GetBestRoute, Call, 10059);
    CALL_(SetupDiCreateDeviceInfoListEx, Call, 10060);
    CALL_(SetupDiGetClassDevsEx, Call, 10061);
    CALL_(SetupDiEnumDeviceInfo, Call, 10062);
    CALL_(SetupDiGetDeviceRegistryProperty, Call, 10063);
    CALL_(SetupDiEnumDeviceInterfaces, Call, 10064);
    CALL_(SetupDiGetDeviceInterfaceDetail, Call, 10065);
    CALL_(SetupDiClassNameFromGuidEx, Call, 10066);
    CALL_(SetupDiGetDeviceInstanceId, Call, 10067);
    CALL_(SetupDiClassGuidsFromNameEx, Call, 10068);
    CALL_(DeviceIoControl, Call, 10069);
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
    CALL_(Twapi_SetServiceStatus, Call, 10082);
    CALL_(Twapi_BecomeAService, Call, 10083);
    CALL_(Twapi_WNetUseConnection, Call, 10084);
    CALL_(NetShareAdd, Call, 10085);
    CALL_(SHGetFolderPath, Call, 10086);
    CALL_(SHGetSpecialFolderPath, Call, 10087);
    CALL_(SHGetPathFromIDList, Call, 10088); // TBD - Tcl wrapper
    CALL_(GetThemeColor, Call, 10089);
    CALL_(GetThemeFont, Call, 10090);
    CALL_(Twapi_WriteShortcut, Call, 10091);
    CALL_(Twapi_ReadShortcut, Call, 10092);
    CALL_(Twapi_InvokeUrlShortcut, Call, 10093);
    CALL_(SHInvokePrinterCommand, Call, 10094); // TBD - Tcl wrapper
    CALL_(Twapi_SHFileOperation, Call, 10095); // TBD - some more wrappers
    CALL_(Twapi_ShellExecuteEx, Call, 10096);
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
    CALL_(Twapi_VerQueryValue_FIXEDFILEINFO, Call, 10107);
    CALL_(Twapi_VerQueryValue_TRANSLATIONS, Call, 10108);
    CALL_(Twapi_FreeFileVersionInfo, Call, 10109);
    CALL_(malloc, Call, 10110);        /* TBD - document, change to memalloc */
    CALL_(Twapi_WriteMemoryInt, Call, 10111);
    CALL_(Twapi_WriteMemoryBinary, Call, 10112);
    CALL_(Twapi_WriteMemoryChars, Call, 10113);
    CALL_(Twapi_WriteMemoryUnicode, Call, 10114);
    CALL_(Twapi_WriteMemoryPointer, Call, 10115);
    CALL_(Twapi_WriteMemoryWide, Call, 10116);
    CALL_(Twapi_NPipeServer, Call, 10117);
    CALL_(Twapi_NPipeClient, Call, 10118);
    CALL_(Twapi_IsEqualPtr, Call, 10119);
    CALL_(Twapi_IsNullPtr, Call, 10120);
    CALL_(Twapi_IsPtr, Call, 10121);
    CALL_(CreateEvent, Call, 10122);
    CALL_(NotifyChangeEventLog, Call, 10123);
    CALL_(UpdateResource, Call, 10124);
    CALL_(FindResourceEx, Call, 10125);
    CALL_(Twapi_LoadResource, Call, 10126);
    CALL_(Twapi_EnumResourceNames, Call, 10127);
    CALL_(Twapi_EnumResourceLanguages, Call, 10128);
    CALL_(Twapi_SplitStringResource, Call, 10129);
    CALL_(Twapi_GetProcessList, Call, 10130);

    // CallU API
    CALL_(IsClipboardFormatAvailable, CallU, 1);
    CALL_(GetClipboardData, CallU, 2);
    CALL_(GetClipboardFormatName, CallU, 3);
    CALL_(GetStdHandle, CallU, 4);
    CALL_(SetConsoleCP, CallU, 5);
    CALL_(SetConsoleOutputCP, CallU, 6);
    CALL_(VerLanguageName, CallU, 7);
    CALL_(GetPerAdapterInfo, CallU, 8);
    CALL_(GetIfEntry, CallU, 9);
    CALL_(GetIfTable, CallU, 10);
    CALL_(GetIpAddrTable, CallU, 11);
    CALL_(GetIpNetTable, CallU, 12);
    CALL_(GetIpForwardTable, CallU, 13);
    CALL_(FlushIpNetTable, CallU, 14);
    CALL_(GetComputerNameEx, CallU, 15);
    CALL_(Wow64EnableWow64FsRedirection, CallU, 16);
    CALL_(GetSystemMetrics, CallU, 17);
    CALL_(Sleep, CallU, 18);
    CALL_(SetThreadExecutionState, CallU, 19);
    CALL_(Twapi_EnumPrinters_Level4, CallU, 20);
    CALL_(UuidCreate, CallU, 21);
    CALL_(GetUserNameEx, CallU, 22);
    CALL_(ImpersonateSelf, CallU, 23);
    CALL_(GetAsyncKeyState, CallU, 24);
    CALL_(GetKeyState, CallU, 25);
    CALL_(GetThreadDesktop, CallU, 26);
    CALL_(BlockInput, CallU, 27);
    CALL_(UnregisterHotKey, CallU, 28);
    CALL_(SetCaretBlinkTime, CallU, 29);
    CALL_(MessageBeep, CallU, 30);
    CALL_(GetGUIThreadInfo, CallU, 31);
    CALL_(Twapi_UnregisterDeviceNotification, CallU, 32);
    CALL_(Sleep, CallU, 33);
    CALL_(Twapi_MapWindowsErrorToString, CallU, 34);
    CALL_(ProcessIdToSessionId, CallU, 35);
    CALL_(Twapi_MemLifoInit, CallU, 37);
    CALL_(GlobalDeleteAtom, CallU, 38); // TBD - tcl interface

    CALL_(Beep, CallU, 1002);
    CALL_(MapVirtualKey, CallU, 1003);
    CALL_(SetCaretPos, CallU, 1004);
    CALL_(SetCursorPos, CallU, 1005);
    CALL_(GetLocaleInfo, CallU, 1006);
    CALL_(ExitWindowsEx, CallU, 1007);
    CALL_(GenerateConsoleCtrlEvent, CallU, 1008);
    CALL_(AllocateAndGetTcpExTableFromStack, CallU, 1009);
    CALL_(AllocateAndGetUdpExTableFromStack, CallU, 1010);

    CALL_(OpenProcess, CallU, 2001);
    CALL_(OpenThread, CallU, 2002);
    CALL_(OpenInputDesktop, CallU, 2003);
    CALL_(AttachThreadInput, CallU, 2004);
    CALL_(RegisterHotKey, CallU, 2005);

    CALL_(GlobalAlloc, CallU, 10001);
    CALL_(LHashValOfName, CallU, 10002);
    CALL_(SetClipboardData, CallU, 10003);
    CALL_(SetStdHandle, CallU, 10004);
    CALL_(CreateConsoleScreenBuffer, CallU, 10006);

    // CallS - function(LPWSTR)
    CALL_(RegisterClipboardFormat, CallS, 1);
    CALL_(SetConsoleTitle, CallS, 2);
    CALL_(GetDriveType, CallS, 3);
    CALL_(DeleteVolumeMountPoint, CallS, 4);
    CALL_(GetVolumeNameForVolumeMountPoint, CallS, 5);
    CALL_(GetVolumePathName, CallS, 6);
    CALL_(GetAdapterIndex, CallS, 7);
    CALL_(CommandLineToArgv, CallS, 8);
    CALL_(WNetGetUniversalName, CallS, 9);
    CALL_(WNetGetUser, CallS, 10);
    CALL_(Twapi_AppendLog, CallS, 11);
    CALL_(WTSOpenServer, CallS, 12);
    CALL_(Twapi_ReadUrlShortcut, CallS, 13);
    CALL_(GetVolumeInformation, CallS, 14);
    CALL_(FindFirstVolumeMountPoint, CallS, 15);
    CALL_(GetDiskFreeSpaceEx, CallS, 16);
#ifndef TWAPI_LEAN
    CALL_(AbortSystemShutdown, CallS, 17);
#endif
    CALL_(GetPrivateProfileSectionNames, CallS, 18);
    CALL_(SendInput, CallS, 19);
    CALL_(Twapi_SendUnicode, CallS, 20);
    CALL_(Twapi_GetFileVersionInfo, CallS, 21);
    CALL_(ExpandEnvironmentStrings, CallS, 22);
    CALL_(GlobalAddAtom, CallS, 23); // TBD - Tcl interface

    CALL_(ConvertStringSecurityDescriptorToSecurityDescriptor, CallS, 501);
    CALL_(Twapi_LsaOpenPolicy, CallS, 502);
    CALL_(NetFileClose, CallS, 503);
    CALL_(LoadLibraryEx, CallS, 504);
    CALL_(AddFontResourceEx, CallS, 505);
    CALL_(BeginUpdateResource, CallS, 506);

    CALL_(OpenWindowStation, CallS, 1001);
    CALL_(WNetCancelConnection2, CallS, 1002);
    CALL_(NetFileGetInfo, CallS, 1003);
    CALL_(GetNamedSecurityInfo, CallS, 1004);
    CALL_(TranslateName, CallS, 1005);


    // CallH - function(HANDLE)
    CALL_(FlushConsoleInputBuffer, CallH, 1);
    CALL_(GetConsoleMode, CallH, 2);
    CALL_(GetConsoleScreenBufferInfo, CallH, 3);
    CALL_(GetLargestConsoleWindowSize, CallH, 4);
    CALL_(GetNumberOfConsoleInputEvents, CallH, 5);
    CALL_(SetConsoleActiveScreenBuffer, CallH, 6);
    CALL_(FindVolumeClose, CallH, 7);
    CALL_(FindVolumeMountPointClose, CallH, 8);
    CALL_(DeregisterEventSource, CallH, 9);
    CALL_(CloseEventLog, CallH, 10);
    CALL_(GetNumberOfEventLogRecords, CallH, 11);
    CALL_(GetOldestEventLogRecord, CallH, 12);
    CALL_(Twapi_IsEventLogFull, CallH, 13);
    CALL_(GetHandleInformation, CallH, 14);
    CALL_(FreeLibrary, CallH, 15);
    CALL_(GetDevicePowerState, CallH, 16);
    CALL_(EnumProcessModules, CallH, 17);
    CALL_(IsWow64Process, CallH, 18);
    CALL_(ResumeThread, CallH, 19);
    CALL_(SuspendThread, CallH, 20);
    CALL_(GetPriorityClass, CallH, 21);
    CALL_(Twapi_NtQueryInformationProcessBasicInformation, CallH, 22);
    CALL_(Twapi_NtQueryInformationThreadBasicInformation, CallH, 23);
    CALL_(GetThreadPriority, CallH, 24);
    CALL_(GetExitCodeProcess, CallH, 25);
    CALL_(LsaClose, CallH, 26);
    CALL_(ImpersonateLoggedOnUser, CallH, 27);
    CALL_(DeleteService, CallH, 28);
    CALL_(CloseServiceHandle, CallH, 29);
    CALL_(CloseThemeData, CallH, 30);
    CALL_(CloseDesktop, CallH, 31);
    CALL_(SwitchDesktop, CallH, 32);
    CALL_(SetThreadDesktop, CallH, 33);
    CALL_(EnumDesktops, CallH, 34);
    CALL_(EnumDesktopWindows, CallH, 35);
    CALL_(CloseWindowStation, CallH, 36);
    CALL_(SetProcessWindowStation, CallH, 37);
    CALL_(FindNextVolume, CallH, 38);
    CALL_(FindNextVolumeMountPoint, CallH, 39);
    CALL_(GetFileType, CallH, 40); /* TBD - TCL wrapper */
    CALL_(QueryServiceConfig, CallH, 41);
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
    CALL_(GetFileTime, CallH, 52);
    CALL_(Twapi_UnregisterDirectoryMonitor, CallH, 53);
    CALL_(Twapi_MemLifoClose, CallH, 54);
    CALL_(Twapi_MemLifoPopFrame, CallH, 55);
    CALL_(Twapi_Free_SEC_WINNT_AUTH_IDENTITY, CallH, 56);
    CALL_(SetupDiDestroyDeviceInfoList, CallH, 57);
    CALL_(GetMonitorInfo, CallH, 58);
    CALL_(GetObject, CallH, 59);
    CALL_(Twapi_MemLifoPushMark, CallH, 60);
    CALL_(Twapi_MemLifoPopMark, CallH, 61);
    CALL_(Twapi_MemLifoValidate, CallH, 62);
    CALL_(Twapi_MemLifoDump, CallH, 63);
    CALL_(ImpersonateNamedPipeClient, CallH, 64);
    CALL_(SetEvent, CallH, 66);
    CALL_(ResetEvent, CallH, 67);
    CALL_(Twapi_EnumResourceTypes, CallH, 68);

    CALL_(ReleaseSemaphore, CallH, 1001);
    CALL_(ControlService, CallH, 1002);
    CALL_(EnumDependentServices, CallH, 1003);
    CALL_(QueryServiceStatusEx, CallH, 1004);
    CALL_(OpenProcessToken, CallH, 1005);
    CALL_(GetTokenInformation, CallH, 1006);
    CALL_(Twapi_SetTokenVirtualizationEnabled, CallH, 1007);
    CALL_(Twapi_SetTokenMandatoryPolicy, CallH, 1008);
    CALL_(TerminateProcess, CallH, 1009);
    CALL_(WaitForInputIdle, CallH, 1010);
    CALL_(SetPriorityClass, CallH, 1011);
    CALL_(SetThreadPriority, CallH, 1012);
    CALL_(SetConsoleMode, CallH, 1013);
    CALL_(SetConsoleTextAttribute, CallH, 1014);
    CALL_(ReadConsole, CallH, 1015);
    CALL_(GetDeviceCaps, CallH, 1016);
    CALL_(WaitForSingleObject, CallH, 1017);
    CALL_(Twapi_MemLifoAlloc, CallH, 1018);
    CALL_(Twapi_MemLifoPushFrame, CallH, 1019);
    CALL_(EndUpdateResource, CallH, 1020);

    CALL_(WTSDisconnectSession, CallH, 2001);
    CALL_(WTSLogoffSession, CallH, 2003);        /* TBD - tcl wrapper */
    CALL_(WTSQuerySessionInformation, CallH, 2003); /* TBD - tcl wrapper */
    CALL_(GetSecurityInfo, CallH, 2004);
    CALL_(OpenThreadToken, CallH, 2005);
    CALL_(ReadEventLog, CallH, 2006);
    CALL_(SetHandleInformation, CallH, 2007);
    CALL_(Twapi_MemLifoExpandLast, CallH, 2008);
    CALL_(Twapi_MemLifoShrinkLast, CallH, 2009);
    CALL_(Twapi_MemLifoResizeLast, CallH, 2010);
    CALL_(Twapi_RegisterWaitOnHandle, CallH, 2011);

    CALL_(SetFileTime, CallH, 10001);
    CALL_(SetThreadToken, CallH, 10002);
    CALL_(Twapi_LsaEnumerateAccountsWithUserRight, CallH, 10003);
    CALL_(Twapi_SetTokenIntegrityLevel, CallH, 10004);
    CALL_(EnumDisplayMonitors, CallH, 10005);

    // CallHSU - function(HANDLE, LPCWSTR, DWORD)
    CALL_(BackupEventLog, CallHSU, 1);
    CALL_(ClearEventLog, CallHSU, 2);
    CALL_(GetServiceKeyName, CallHSU, 3);
    CALL_(GetServiceDisplayName, CallHSU, 4);
    CALL_(WriteConsole, CallHSU, 5);
    CALL_(FillConsoleOutputCharacter, CallHSU, 6);
    CALL_(OpenService, CallHSU, 7);

    // CallW - function(HWND)
    CALL_(IsIconic, CallW, 1);
    CALL_(IsZoomed, CallW, 2);
    CALL_(IsWindowVisible, CallW, 3);
    CALL_(IsWindow, CallW, 4);
    CALL_(IsWindowUnicode, CallW, 5);
    CALL_(IsWindowEnabled, CallW, 6);
    CALL_(ArrangeIconicWindows, CallW, 7);
    CALL_(SetForegroundWindow, CallW, 8);
    CALL_(OpenIcon, CallW, 9);
    CALL_(CloseWindow, CallW, 10);
    CALL_(DestroyWindow, CallW, 11);
    CALL_(UpdateWindow, CallW, 12);
    CALL_(HideCaret, CallW, 13);
    CALL_(ShowCaret, CallW, 14);
    CALL_(GetParent, CallW, 15);
    CALL_(OpenClipboard, CallW, 16);
    CALL_(GetClientRect, CallW, 17);
    CALL_(GetWindowRect, CallW, 18);
    CALL_(GetDC, CallW, 19);
    CALL_(SetFocus, CallW, 20);
    CALL_(SetActiveWindow, CallW, 21);
    CALL_(GetClassName, CallW, 22);
    CALL_(RealGetWindowClass, CallW, 23);
    CALL_(GetWindowThreadProcessId, CallW, 24);
    CALL_(GetWindowText, CallW, 25);
    CALL_(GetWindowDC, CallW, 26);
    CALL_(EnumChildWindows, CallW, 28);
    CALL_(GetWindowPlacement, CallW, 30); // TBD - Tcl wrapper
    CALL_(GetWindowInfo, CallW, 31); // TBD - Tcl wrapper

    CALL_(SetWindowText, CallW, 1001);
    CALL_(IsChild, CallW, 1002);
    CALL_(SetWindowPlacement, CallW, 1003); // TBD - Tcl wrapper
    CALL_(InvalidateRect, CallW, 1004);     // TBD - Tcl wrapper
    CALL_(SetWindowPos, CallW, 1005);
    CALL_(FindWindowEx, CallW, 1006);
    CALL_(ReleaseDC, CallW, 1007);
    CALL_(OpenThemeData, CallW, 1008);

    // CallWU - function(HWND, DWORD)
    CALL_(GetAncestor, CallWU, 1);
    CALL_(GetWindow, CallWU, 2);
    CALL_(ShowWindow, CallWU, 3);
    CALL_(ShowWindowAsync, CallWU, 4);
    CALL_(EnableWindow, CallWU, 5);
    CALL_(ShowOwnedPopups, CallWU, 6);
    CALL_(MonitorFromWindow, CallWU, 7);
    CALL_(GetWindowLongPtr, CallWU, 8);
    CALL_(SHGetSpecialFolderLocation, CallWU, 9);

    CALL_(PostMessage, CallWU, 1001);
    CALL_(SendNotifyMessage, CallWU, 1002);
    CALL_(SendMessageTimeout, CallWU, 1003);

    CALL_(SetLayeredWindowAttributes, CallWU, 10001);
    CALL_(MoveWindow, CallWU, 10002);
    CALL_(SetWindowLongPtr, CallWU, 10003);
    CALL_(SHObjectProperties, CallWU, 10004);

    // CallSSSD - function(LPWSTR_NULL_IF_EMPTY, LPWSTR, LPWSTR, DWORD)
    CALL_(LookupAccountName, CallSSSD, 1);
    CALL_(NetUserAdd, CallSSSD, 2);
    CALL_(NetShareSetInfo, CallSSSD, 3);
    CALL_(Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY, CallSSSD, 4);
    CALL_(LogonUser, CallSSSD, 5);
    CALL_(CreateDesktop, CallSSSD, 6);
    CALL_(LookupPrivilegeDisplayName, CallSSSD, 13);
    CALL_(LookupPrivilegeValue, CallSSSD, 14);
    CALL_(NetGroupAdd, CallSSSD, 15);
    CALL_(NetLocalGroupAdd, CallSSSD, 16);
    CALL_(NetGroupDel, CallSSSD, 17);
    CALL_(NetLocalGroupDel, CallSSSD, 18);
    CALL_(NetGroupAddUser, CallSSSD, 19);
    CALL_(NetGroupDelUser, CallSSSD, 20);
    CALL_(Twapi_NetLocalGroupAddMember, CallSSSD, 21);
    CALL_(Twapi_NetLocalGroupDelMember, CallSSSD, 22);
    CALL_(Twapi_NetShareCheck, CallSSSD, 26);
    CALL_(NetUserDel, CallSSSD, 29);
    CALL_(NetSessionGetInfo, CallSSSD, 32);
    CALL_(NetSessionDel, CallSSSD, 33);
    CALL_(NetGetDCName, CallSSSD, 34);
    CALL_(FindWindow, CallSSSD, 38);
    CALL_(MoveFileEx, CallSSSD, 39);
    CALL_(SetVolumeLabel, CallSSSD, 40);
    CALL_(QueryDosDevice, CallSSSD, 41);
    CALL_(RegisterEventSource, CallSSSD, 44);
    CALL_(OpenEventLog, CallSSSD, 45);
    CALL_(OpenBackupEventLog, CallSSSD, 45);

    // CallPSID - function(ANY, SID, ...)
    CALL_(LookupAccountSid, CallPSID, 2);

    CALL_(CheckTokenMembership, CallPSID, 1002);
    CALL_(Twapi_SetTokenPrimaryGroup, CallPSID, 1003);
    CALL_(Twapi_SetTokenOwner, CallPSID, 1004);
    CALL_(Twapi_LsaEnumerateAccountRights, CallPSID, 1005);
    CALL_(LsaRemoveAccountRights, CallPSID, 1006);
    CALL_(LsaAddAccountRights, CallPSID, 1007);

    // CallNetEnum - function(STRING, ....)
    CALL_(NetUseEnum, CallNetEnum, 1);
    CALL_(NetUserEnum, CallNetEnum, 2);
    CALL_(NetGroupEnum, CallNetEnum, 3);
    CALL_(NetLocalGroupEnum, CallNetEnum, 4);
    CALL_(NetShareEnum, CallNetEnum, 5);
    CALL_(NetUserGetGroups, CallNetEnum, 6);
    CALL_(NetUserGetLocalGroups, CallNetEnum, 7);
    CALL_(NetLocalGroupGetMembers, CallNetEnum, 8);
    CALL_(NetGroupGetUsers, CallNetEnum, 9);
    CALL_(NetConnectionEnum, CallNetEnum, 10);
    CALL_(NetFileEnum, CallNetEnum, 11);
    CALL_(NetSessionEnum, CallNetEnum, 12);

    CALL_(PdhGetDllVersion, CallPdh, 1);
    CALL_(PdhBrowseCounters, CallPdh, 2);
    CALL_(PdhSetDefaultRealTimeDataSource, CallPdh, 101);
    CALL_(PdhConnectMachine, CallPdh, 102);
    CALL_(PdhValidatePath, CallPdh, 103);
    CALL_(PdhParseCounterPath, CallPdh, 201);
    CALL_(PdhLookupPerfNameByIndex, CallPdh, 202);
    CALL_(PdhOpenQuery, CallPdh, 203);
    CALL_(PdhRemoveCounter, CallPdh, 301);
    CALL_(PdhCollectQueryData, CallPdh, 302);
    CALL_(PdhCloseQuery, CallPdh, 303);
    CALL_(PdhGetFormattedCounterValue, CallPdh, 1001);
    CALL_(PdhAddCounter, CallPdh, 1002);
    CALL_(PdhMakeCounterPath, CallPdh, 1003);
    CALL_(PdhEnumObjectItems, CallPdh, 1004);
    CALL_(PdhEnumObjects, CallPdh, 1005);

#undef CALL_

    // CallCOM
#define CALLCOM_(fn_, code_)                                         \
    do {                                                                \
        TwapiMakeCallAlias(interp, "twapi::" #fn_, "twapi::CallCOM", # code_); \
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

    CALLCOM_(ITaskScheduler_Activate, 5001);
    CALLCOM_(ITaskScheduler_AddWorkItem, 5002);
    CALLCOM_(ITaskScheduler_Delete, 5003);
    CALLCOM_(ITaskScheduler_Enum, 5004);
    CALLCOM_(ITaskScheduler_IsOfType, 5005);
    CALLCOM_(ITaskScheduler_NewWorkItem, 5006);
    CALLCOM_(ITaskScheduler_SetTargetComputer, 5007);
    CALLCOM_(ITaskScheduler_GetTargetComputer, 5008);

    CALLCOM_(IEnumWorkItems_Clone, 5101);
    CALLCOM_(IEnumWorkItems_Reset, 5102);
    CALLCOM_(IEnumWorkItems_Skip, 5103);
    CALLCOM_(IEnumWorkItems_Next, 5104);

    CALLCOM_(IScheduledWorkItem_CreateTrigger, 5201);
    CALLCOM_(IScheduledWorkItem_DeleteTrigger, 5202);
    CALLCOM_(IScheduledWorkItem_EditWorkItem, 5203);
    CALLCOM_(IScheduledWorkItem_GetAccountInformation, 5204);
    CALLCOM_(IScheduledWorkItem_GetComment, 5205);
    CALLCOM_(IScheduledWorkItem_GetCreator, 5206);
    CALLCOM_(IScheduledWorkItem_GetErrorRetryCount, 5207);
    CALLCOM_(IScheduledWorkItem_GetErrorRetryInterval, 5208);
    CALLCOM_(IScheduledWorkItem_GetExitCode, 5209);
    CALLCOM_(IScheduledWorkItem_GetFlags, 5210);
    CALLCOM_(IScheduledWorkItem_GetIdleWait, 5211);
    CALLCOM_(IScheduledWorkItem_GetMostRecentRunTime, 5212);
    CALLCOM_(IScheduledWorkItem_GetNextRunTime, 5213);
    CALLCOM_(IScheduledWorkItem_GetStatus, 5214);
    CALLCOM_(IScheduledWorkItem_GetTrigger, 5215);
    CALLCOM_(IScheduledWorkItem_GetTriggerCount, 5216);
    CALLCOM_(IScheduledWorkItem_GetTriggerString, 5217);
    CALLCOM_(IScheduledWorkItem_Run, 5218);
    CALLCOM_(IScheduledWorkItem_SetAccountInformation, 5219);
    CALLCOM_(IScheduledWorkItem_SetComment, 5220);
    CALLCOM_(IScheduledWorkItem_SetCreator, 5221);
    CALLCOM_(IScheduledWorkItem_SetErrorRetryCount, 5222);
    CALLCOM_(IScheduledWorkItem_SetErrorRetryInterval, 5223);
    CALLCOM_(IScheduledWorkItem_SetFlags, 5224);
    CALLCOM_(IScheduledWorkItem_SetIdleWait, 5225);
    CALLCOM_(IScheduledWorkItem_Terminate, 5226);
    CALLCOM_(IScheduledWorkItem_SetWorkItemData, 5227);
    CALLCOM_(IScheduledWorkItem_GetWorkItemData, 5228);
    CALLCOM_(IScheduledWorkItem_GetRunTimes, 5229);

    CALLCOM_(ITask_GetApplicationName, 5301);
    CALLCOM_(ITask_GetMaxRunTime, 5302);
    CALLCOM_(ITask_GetParameters, 5303);
    CALLCOM_(ITask_GetPriority, 5304);
    CALLCOM_(ITask_GetTaskFlags, 5305);
    CALLCOM_(ITask_GetWorkingDirectory, 5306);
    CALLCOM_(ITask_SetApplicationName, 5307);
    CALLCOM_(ITask_SetParameters, 5308);
    CALLCOM_(ITask_SetWorkingDirectory, 5309);
    CALLCOM_(ITask_SetMaxRunTime, 5310);
    CALLCOM_(ITask_SetPriority, 5311);
    CALLCOM_(ITask_SetTaskFlags, 5312);

    CALLCOM_(ITaskTrigger_GetTrigger, 5401);
    CALLCOM_(ITaskTrigger_GetTriggerString, 5402);
    CALLCOM_(ITaskTrigger_SetTrigger, 5403);

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
    CALLCOM_(Twapi_ValidIID, 10013);

    return Tcl_Eval(interp, apiprocs);
}


int Twapi_CallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    union {
        WCHAR buf[MAX_PATH+1];
        LPCWSTR wargv[100];     /* FormatMessage accepts up to 99 params + 1 for NULL */
        FLASHWINFO flashw;
        LASTINPUTINFO lastin;
        double d;
        FILETIME   filetime;
        SYSTEMTIME systime;
        POINT  pt;
        LARGE_INTEGER largeint;
        RECT rect;
        SMALL_RECT srect[2];
        TOKEN_PRIVILEGES *tokprivsP;
        MODULEINFO moduleinfo;
        struct {
            SP_DEVINFO_DATA sp_devinfo_data;
            SP_DEVINFO_DATA *sp_devinfo_dataP;
            SP_DEVICE_INTERFACE_DATA sp_device_interface_data;
        } dev;
        DISPLAY_DEVICEW display_device;
        MIB_TCPROW tcprow;
        struct sockaddr_in sinaddr;
        SYSTEM_POWER_STATUS power_status;
        TwapiId twapi_id;
    } u;
    DWORD_PTR dwp;
    COORD coord;
    CHAR_INFO chinfo;
    SECURITY_DESCRIPTOR *secdP;
    DWORD dw, dw2, dw3, dw4;
    int i, i2;
    LPWSTR s, s2, s3, s4;
    unsigned char *cP;
    LUID luid;
    LUID *luidP;
    void *pv, *pv2, *pv3;
    SecHandle sech, sech2, *sech2P;
    Tcl_Obj *objs[2];
    SecBufferDesc sbd, *sbdP;
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h, h2, h3;
    HWND   hwnd;
    PSID osidP, gsidP;
    ACL *daclP, *saclP;
    WORD w;
    GUID guid;
    GUID *guidP;
    LPITEMIDLIST idlP;

    if (objc < 2)
        return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 1000) {
        /* Functions taking no arguments */
        if (objc != 2)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1:
            result.type = TRT_HANDLE;
            result.value.hval = GetCurrentProcess();
            break;
        case 2:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseClipboard();
            break;
        case 3:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EmptyClipboard();
            break;
        case 4:
            result.type = TRT_HWND;
            result.value.hwin = GetOpenClipboardWindow();
            break;
        case 5:
            result.type = TRT_HWND;
            result.value.hwin = GetDesktopWindow();
            break;
        case 6:
            result.type = TRT_HWND;
            result.value.hwin = GetShellWindow();
            break;
        case 7:
            result.type = TRT_HWND;
            result.value.hwin = GetForegroundWindow();
            break;
        case 8:
            result.type = TRT_HWND;
            result.value.hwin = GetActiveWindow();
            break;
        case 9:
            result.type = TRT_HWND;
            result.value.hwin = GetClipboardOwner();
            break;
        case 10:
            return Twapi_EnumClipboardFormats(interp);
        case 11:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = AllocConsole();
            break;
        case 12:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FreeConsole();
            break;
        case 13:
            result.type = TRT_DWORD;
            result.value.ival = GetConsoleCP();
            break;
        case 14:
            result.type = TRT_DWORD;
            result.value.ival = GetConsoleOutputCP();
            break;
        case 15:
            result.type = GetNumberOfConsoleMouseButtons(&result.value.ival) ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 16:
            /* Note : GetLastError == 0 means title is empty string */
            if ((result.value.unicode.len = GetConsoleTitleW(u.buf, sizeof(u.buf)/sizeof(u.buf[0]))) != 0 || GetLastError() == 0) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 17:
            result.value.hwin = GetConsoleWindow();
            result.type = result.value.hwin ? TRT_HWND : TRT_GETLASTERROR;
            break;
        case 18:
            result.value.ival = GetLogicalDrives();
            result.type = TRT_DWORD;
            break;
        case 19:
            return Twapi_GetNetworkParams(ticP);
        case 20:
            return Twapi_GetAdaptersInfo(ticP);
        case 21:
            return Twapi_GetInterfaceInfo(ticP);
        case 22:
            result.type = GetNumberOfInterfaces(&result.value.ival) ? TRT_GETLASTERROR : TRT_DWORD;
            break;
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
#ifndef TWAPI_LEAN
        case 43:
            return Twapi_EnumProcesses(ticP);
        case 44:
            return Twapi_EnumDeviceDrivers(ticP);
        case 45:
            result.type = TRT_DWORD;
            result.value.ival = GetCurrentThreadId();
            break;
        case 46:
            result.type = TRT_HANDLE;
            result.value.hval = GetCurrentThread();
            break;
#endif
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
#ifndef TWAPI_LEAN
        case 50:
            return Twapi_LsaEnumerateLogonSessions(interp);
        case 51:
            result.value.ival = RevertToSelf();
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 52:
            return Twapi_InitializeSecurityDescriptor(interp);
#endif
        case 53: 
            return Twapi_EnumerateSecurityPackages(interp);
        case 54:
            result.type = TRT_BOOL;
            result.value.bval = Twapi_IsThemeActive();
            break;
        case 55:
            result.type = TRT_BOOL;
            result.value.bval = Twapi_IsAppThemed();
            break;
        case 56:
            return Twapi_GetCurrentThemeName(interp);
        case 57:
            return Twapi_GetShellVersion(interp);
        case 58:
            result.type = TRT_DWORD;
            result.value.ival = GetDoubleClickTime();
            break;
        case 59:
            return Twapi_EnumWindowStations(interp);
        case 60:
            result.type = TRT_HWINSTA;
            result.value.hval = GetProcessWindowStation();
            break;
        case 61:
            result.type = GetCursorPos(&result.value.point) ? TRT_POINT : TRT_GETLASTERROR;
            break;
        case 62:
            result.type = GetCaretPos(&result.value.point) ? TRT_POINT : TRT_GETLASTERROR;
            break;
        case 63:
            result.type = TRT_DWORD;
            result.value.ival = GetCaretBlinkTime();
            break;
        case 64:
            return Twapi_EnumWindows(interp);
        case 65:
            return Twapi_StopConsoleEventNotifier(ticP);
        case 66:
            return TwapiFirstVolume(interp, NULL); /* FindFirstVolume */
        case 67:
            u.lastin.cbSize = sizeof(u.lastin);
            if (GetLastInputInfo(&u.lastin)) {
                result.type = TRT_DWORD;
                result.value.ival = u.lastin.dwTime;
            } else {
                result.type = TRT_GETLASTERROR;
            }
            break;
        case 68:
            if (GetSystemPowerStatus(&u.power_status)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromSYSTEM_POWER_STATUS(&u.power_status);
            } else
                result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 69:
            return Twapi_ClipboardMonitorStart(ticP);
        case 70:
            return Twapi_ClipboardMonitorStop(ticP);
        case 71:
            return Twapi_PowerNotifyStart(ticP);
        case 72:
            return Twapi_PowerNotifyStop(ticP);
        case 73:
            return Twapi_StartConsoleEventNotifier(ticP);
        case 74:
            result.type = TRT_WIDE;
            result.value.wide = TWAPI_NEWID(ticP);
            break;
        case 75:
            result.type = TRT_EMPTY;
            DebugBreak();
            break;
        }

        return TwapiSetResult(interp, &result);
    }

    if (func < 2000) {
        /* We should have exactly one additional argument. */

        if (objc != 3)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
        
        switch (func) {
        case 1001:
            if (ObjToDWORD_PTR(interp, objv[2], &dwp) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_LPVOID;
            result.value.pv = (void *) dwp;
            break;
        case 1002:
            if (ObjToFLASHWINFO(interp, objv[2], &u.flashw) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = FlashWindowEx(&u.flashw);
            break;
        case 1003:
            if (Tcl_GetDoubleFromObj(interp, objv[2], &u.d) != TCL_OK)
                return TCL_ERROR;
            result.type = VariantTimeToSystemTime(u.d, &result.value.systime) ?
                TRT_SYSTEMTIME : TRT_GETLASTERROR;
            break;
        case 1004:
            if (ObjToSYSTEMTIME(interp, objv[2], &u.systime) != TCL_OK)
                return TCL_ERROR;
            result.type = SystemTimeToVariantTime(&u.systime, &result.value.dval) ?
                TRT_DOUBLE : TRT_GETLASTERROR;
            break;
#ifndef TWAPI_LEAN
        case 1005:
        case 1006:
            if (ObjToDWORD_PTR(interp, objv[2], &dwp) != TCL_OK)
                return TCL_ERROR;
            if ((func == 1005 ?
                 GetDeviceDriverBaseNameW
                 : GetDeviceDriverFileNameW) (
                     (LPVOID) dwp,
                     u.buf,
                     ARRAYSIZE(u.buf))) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;
#endif
        case 1007:
            if (IPAddrObjToDWORD(interp, objv[2], &dw) == TCL_ERROR)
                result.type = TRT_TCL_RESULT;
            else {
                result.value.ival = GetBestInterface(dw, &dw2);
                if (result.value.ival)
                    result.type = TRT_EXCEPTION_ON_ERROR;
                else {
                    result.value.ival = dw2;
                    result.type = TRT_DWORD;
                }
            }
            break;

        case 1008:
            if (ObjToSecHandle(interp, objv[2], &sech) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            dw = QuerySecurityContextToken(&sech, &result.value.hval);
            if (dw) {
                result.value.ival =  dw;
                result.type = TRT_EXCEPTION_ON_ERROR;
            } else {
                result.type = TRT_HANDLE;
            }
            break;
        case 1009:
            if (ObjToPOINT(interp, objv[2], &u.pt) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HWND;
            result.value.hwin = WindowFromPoint(u.pt);
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
        case 1011: // FreeCredentialsHandle
            if (ObjToSecHandle(interp, objv[2], &sech) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = FreeCredentialsHandle(&sech);
            break;
        case 1012: // DeleteSecurityContext
            if (ObjToSecHandle(interp, objv[2], &sech) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = DeleteSecurityContext(&sech);
            break;
        case 1013: // ImpersonateSecurityContext
            if (ObjToSecHandle(interp, objv[2], &sech) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = ImpersonateSecurityContext(&sech);
            break;
        case 1014:
            if (ObjToPACL(interp, objv[2], &daclP) != TCL_OK)
                return TCL_ERROR;
            // Note aclP may me NULL even on TCL_OK
            result.type = TRT_BOOL;
            result.value.bval = daclP ? IsValidAcl(daclP) : 0;
            if (daclP)
                TwapiFree(daclP);
            break;
        case 1015:
            if (ObjToMIB_TCPROW(interp, objv[2], &u.tcprow) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = SetTcpEntry(&u.tcprow);
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
            if (ObjToSYSTEMTIME(interp, objv[2], &u.systime) != TCL_OK)
                return TCL_ERROR;
            if (SystemTimeToFileTime(&u.systime, &result.value.filetime))
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
        case 1021:
            return TwapiGetThemeDefine(interp, Tcl_GetString(objv[2]));
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
        case 10000: // Twapi_FormatExtendedTcpTable
        case 10001: // Twapi_FormatExtendedUdpTable
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVOIDP(pv), GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return (func == 10000 ? Twapi_FormatExtendedTcpTable : Twapi_FormatExtendedUdpTable)
                (interp, pv, dw, dw2);
        case 10002: // GetExtendedTcpTable
        case 10003: // GetExtendedUdpTable
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVOIDP(pv), GETINT(dw), GETBOOL(dw2),
                             GETINT(dw3), GETINT(dw4),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return (func == 10002 ? Twapi_GetExtendedTcpTable : Twapi_GetExtendedUdpTable)
                (interp, pv, dw, dw2, dw3, dw4);

        case 10004: // ResolveAddressAsync
        case 10005: // ResolveHostnameAsync
            return (func == 10004 ? Twapi_ResolveAddressAsync : Twapi_ResolveHostnameAsync)
                (ticP, objc-2, objv+2);

        case 10006:
            return Twapi_GetAddrInfo(interp, objc-2, objv+2);
        case 10007:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(u.sinaddr, ObjToSOCKADDR_IN), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_GetNameInfo(interp, &u.sinaddr, dw);

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
        case 10010:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(u.pt, ObjToPOINT), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromPoint(u.pt, dw);
            break;

        case 10011:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(u.rect, ObjToRECT), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromRect(&u.rect, dw);
            break;

        case 10012:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            u.display_device.cb = sizeof(u.display_device);
            if (EnumDisplayDevicesW(s, dw, &u.display_device, dw2)) {
                result.value.obj = ObjFromDISPLAY_DEVICE(&u.display_device);
                result.type = TRT_OBJ;
            } else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = ERROR_INVALID_PARAMETER;
            }
            break;
                      
        case 10013:
            return Twapi_ReportEvent(interp, objc-2, objv+2);
        case 10014:
            return Twapi_RegisterDeviceNotification(ticP, objc-2, objv+2);
        case 10015: // CreateProcess
        case 10016: // CreateProcessAsUser
            return TwapiCreateProcessHelper(interp, func==10016, objc-2, objv+2);
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
                result.value.obj = Tcl_NewUnicodeObj(s, -1);
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
        case 10020:
            luidP = &luid;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETWSTR(s2), GETINT(dw),
                             GETVAR(luidP, ObjToLUID_NULL),
                             GETVOIDP(pv), ARGEND) != TCL_OK)
                return TCL_ERROR;
            NULLIFY_EMPTY(s);
            result.value.ival = AcquireCredentialsHandleW(
                s, s2,
                dw, luidP, pv, NULL, NULL, &sech, &u.largeint);
            if (result.value.ival) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                break;
            }
            objs[0] = ObjFromSecHandle(&sech);
            objs[1] = Tcl_NewWideIntObj(u.largeint.QuadPart);
            result.type = TRT_OBJV;
            result.value.objv.objPP = objs;
            result.value.objv.nobj = 2;
            break;
        case 10021:
            sech2P = &sech2;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sech2P, ObjToSecHandle_NULL),
                             GETWSTR(s),
                             GETINT(dw),
                             GETINT(dw2),
                             GETINT(dw3),
                             GETVAR(sbd, ObjToSecBufferDescRO),
                             GETINT(dw4),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            sbdP = sbd.cBuffers ? &sbd : NULL;
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_InitializeSecurityContext(
                interp, &sech, sech2P, s,
                dw, dw2, dw3, sbdP, dw4);
            TwapiFreeSecBufferDesc(sbdP);
            break;

        case 10022: // AcceptSecurityContext
            sech2P = &sech2;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sech2P, ObjToSecHandle_NULL),
                             GETVAR(sbd, ObjToSecBufferDescRO),
                             GETINT(dw),
                             GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            sbdP = sbd.cBuffers ? &sbd : NULL;
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_AcceptSecurityContext(
                interp, &sech, sech2P, sbdP, dw, dw2);
            TwapiFreeSecBufferDesc(sbdP);
            break;
        
        case 10023: // QueryContextAttributes
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(sech, ObjToSecHandle),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_QueryContextAttributes(interp, &sech, dw);
        case 10024: // MakeSignature
        case 10025: // EncryptMessage
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(sech, ObjToSecHandle),
                             GETINT(dw),
                             GETBIN(cP, dw2),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return (func == 10024 ? Twapi_MakeSignature : Twapi_EncryptMessage) (
                ticP, &sech, dw, dw2, cP, dw3);

        case 10026: // VerifySignature
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sbd, ObjToSecBufferDescRO),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            sbdP = sbd.cBuffers ? &sbd : NULL;
            dw2 = VerifySignature(&sech, sbdP, dw, &result.value.ival);
            TwapiFreeSecBufferDesc(sbdP);
            if (dw2 == 0)
                result.type = TRT_DWORD;
            else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw2;
            }
            break;

        case 10027: // DecryptMessage
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sbd, ObjToSecBufferDescRW),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            dw2 = DecryptMessage(&sech, &sbd, dw, &result.value.ival);
            if (dw2 == 0) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromSecBufferDesc(&sbd);
            } else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw2;
            }
            TwapiFreeSecBufferDesc(&sbd);
            break;

        case 10028: // CryptAcquireContext
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2), GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (CryptAcquireContextW(&result.value.dwp, s, s2, dw, dw2))
                result.type = TRT_DWORD_PTR;
            else
                result.type = TRT_GETLASTERROR;
            break;

        case 10029:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETDWORD_PTR(dwp), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CryptReleaseContext(dwp, dw);
            break;

        case 10030:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETDWORD_PTR(dwp), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_CryptGenRandom(interp, dwp, dw);
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
        case 10033:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETINT(dw), GETINT(dw2),
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HWINSTA;
            result.value.hval = CreateWindowStationW(s, dw, dw2, secattrP);
            TwapiFreeSECURITY_ATTRIBUTES(secattrP);
            break;
        case 10034:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETINT(dw), GETINT(dw2), GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HDESK;
            result.value.hval = OpenDesktopW(s, dw, dw2, dw3);
            break;
        case 10035: // RegisterDirChangeNotifier
            return Twapi_RegisterDirectoryMonitor(ticP, objc-2, objv+2);
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
        case 10037: // PlaySound
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETDWORD_PTR(dwp), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            NULLIFY_EMPTY(s);
            result.type = TRT_BOOL;
            result.value.ival = PlaySoundW(s, (HMODULE)dwp, dw);
            break;
        case 10038: // SetConsoleWindowInfo
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETBOOL(dw), GETVAR(u.srect[0], ObjToSMALL_RECT),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleWindowInfo(h, dw, &u.srect[0]);
            break;
        case 10039: // FillConsoleOutputAttribute
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWORD(w), GETINT(dw),
                             GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            if (FillConsoleOutputAttribute(h, w, dw, coord, &result.value.ival))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;

        case 10040:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETPTR(h, SC_HANDLE), GETINT(dw), GETINT(dw2),
                             GETINT(dw3), GETNULLTOKEN(s),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_EnumServicesStatusEx(ticP, h, dw, dw2, dw3, s);
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
        case 10042: // ScrollConsoleScreenBuffer
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h),
                             GETVAR(u.srect[0], ObjToSMALL_RECT),
                             GETVAR(u.srect[1], ObjToSMALL_RECT),
                             GETVAR(coord, ObjToCOORD),
                             GETVAR(chinfo, ObjToCHAR_INFO),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ScrollConsoleScreenBufferW(
                h, &u.srect[0], &u.srect[1], coord, &chinfo);
            break;
        case 10043:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWSTRN(s, dw),
                             GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (WriteConsoleOutputCharacterW(h, s, dw, coord, &result.value.ival))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;    
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

        case 10046: // ReadProcessMemory
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETDWORD_PTR(dwp), GETVOIDP(pv),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type =
                ReadProcessMemory(h, (void *)dwp, pv, dw, &result.value.dwp)
                ? TRT_DWORD_PTR : TRT_GETLASTERROR;
            break;
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
        case 10049:
            return Twapi_ChangeServiceConfig(interp, objc-2, objv+2);
        case 10050:
            return Twapi_CreateService(interp, objc-2, objv+2);
        case 10051:
            return Twapi_StartService(interp, objc-2, objv+2);
        case 10052: // SetConsoleCursorPosition
        case 10053: // SetConsoleScreenBufferSize
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival =
                (func == 10052 ? SetConsoleCursorPosition : SetConsoleScreenBufferSize)
                (h, coord);
            break;
        case 10054: // GetModuleFileName
        case 10055: // GetModuleBaseName
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETDWORD_PTR(dwp),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if ((func == 10054 ?
                 GetModuleFileNameExW
                 : GetModuleBaseNameW)
                (h, (HMODULE) dwp, u.buf, ARRAYSIZE(u.buf))) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 10056: // GetModuleInformation
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETDWORD_PTR(dwp),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (GetModuleInformation(h, (HMODULE) dwp,
                                     &u.moduleinfo, sizeof(u.moduleinfo))) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromMODULEINFO(&u.moduleinfo);
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 10057: // VerQueryValue_STRING
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETPTR(pv, TWAPI_FILEVERINFO),
                             GETWSTR(s), GETWSTR(s2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_VerQueryValue_STRING(interp, pv, s, s2);
        case 10058: // DsGetDcName
            guidP = &guid;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETVAR(guidP, ObjToUUID_NULL),
                             GETNULLIFEMPTY(s3), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_DsGetDcName(interp, s, s2, guidP, s3, dw);
        case 10059:
            return Twapi_GetBestRoute(ticP, objc-2, objv+2);

#ifndef TWAPI_LEAN
        case 10060: // SetupDiCreateDeviceInfoListExW
            guidP = &guid;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(guidP, ObjToGUID_NULL),
                             GETHWND(hwnd),
                             GETNULLIFEMPTY(s),
                             GETVOIDP(pv),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HDEVINFO;
            result.value.hval = SetupDiCreateDeviceInfoListExW(guidP, hwnd, s, pv);
            break;

        case 10061:
            guidP = &guid;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(guidP, ObjToGUID_NULL),
                             GETNULLIFEMPTY(s),
                             GETHWND(hwnd),
                             GETINT(dw),
                             GETHANDLET(h, HDEVINFO),
                             GETNULLIFEMPTY(s2),
                             ARGUSEDEFAULT,
                             GETVOIDP(pv),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HDEVINFO;
            result.value.hval = SetupDiGetClassDevsExW(guidP, s, hwnd, dw,
                                                       h, s2, pv);
            break;
        case 10062: // SetupDiEnumDeviceInfo
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLET(h, HDEVINFO),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            u.dev.sp_devinfo_data.cbSize = sizeof(u.dev.sp_devinfo_data);
            if (SetupDiEnumDeviceInfo(h, dw, &u.dev.sp_devinfo_data)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromSP_DEVINFO_DATA(&u.dev.sp_devinfo_data);
            } else
                result.type = TRT_GETLASTERROR;
            break;

        case 10063: // Twapi_SetupDiGetDeviceRegistryProperty
            return Twapi_SetupDiGetDeviceRegistryProperty(ticP, objc-2, objv+2);
        case 10064: // SetupDiEnumDeviceInterfaces
            u.dev.sp_devinfo_dataP = & u.dev.sp_devinfo_data;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLET(h, HDEVINFO),
                             GETVAR(u.dev.sp_devinfo_dataP, ObjToSP_DEVINFO_DATA_NULL),
                             GETGUID(guid),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            u.dev.sp_device_interface_data.cbSize = sizeof(u.dev.sp_device_interface_data);
            if (SetupDiEnumDeviceInterfaces(
                    h, u.dev.sp_devinfo_dataP,  &guid,
                    dw, &u.dev.sp_device_interface_data)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromSP_DEVICE_INTERFACE_DATA(&u.dev.sp_device_interface_data);
            } else
                result.type = TRT_GETLASTERROR;
            break;

        case 10065:
            return Twapi_SetupDiGetDeviceInterfaceDetail(ticP, objc-2, objv+2);
#endif
        case 10066:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETGUID(guid),
                             ARGUSEDEFAULT,
                             GETNULLIFEMPTY(s),
                             GETVOIDP(pv),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (SetupDiClassNameFromGuidExW(&guid, u.buf, ARRAYSIZE(u.buf),
                                            NULL, s, pv)) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 10067:
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLET(h, HDEVINFO),
                             GETVAR(u.dev.sp_devinfo_data, ObjToSP_DEVINFO_DATA),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (SetupDiGetDeviceInstanceIdW(h, &u.dev.sp_devinfo_data,
                                            u.buf, ARRAYSIZE(u.buf), NULL)) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 10068:
            return Twapi_SetupDiClassGuidsFromNameEx(ticP, objc-2, objv+2);
        case 10069:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h),   GETINT(dw),
                             GETVOIDP(pv),   GETINT(dw2),
                             GETVOIDP(pv2),  GETINT(dw3),
                             ARGUSEDEFAULT,
                             GETVOIDP(pv3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (DeviceIoControl(h, dw, pv, dw2, pv2, dw3,
                                &result.value.ival, pv3))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;
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
        case 10082:
            return Twapi_SetServiceStatus(ticP, objc-2, objv+2);
        case 10083:
            return Twapi_BecomeAService(ticP, objc-2, objv+2);
        case 10084:
            return Twapi_WNetUseConnection(interp, objc-2, objv+2);
        case 10085:
            return Twapi_NetShareAdd(interp, objc-2, objv+2);
        case 10086: // SHGetFolderPath - TBD Tcl wrapper
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHWND(hwnd), GETINT(dw),
                             GETHANDLE(h), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            dw = Twapi_SHGetFolderPath(hwnd, dw, h, dw2, u.buf);
            if (dw == 0) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw;
            }
            break;
        case 10087:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHWND(hwnd), GETINT(dw),
                             GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (SHGetSpecialFolderPathW(hwnd, u.buf, dw, dw2)) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 10088: // SHGetPathFromIDList
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(idlP, ObjToPIDL),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (SHGetPathFromIDListW(idlP, u.buf)) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else {
                /* Need to get error before we call the pidl free */
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = GetLastError();
            }
            TwapiFreePIDL(idlP); /* OK if NULL */
            break;
        case 10089: 
            return Twapi_GetThemeColor(interp, objc-2, objv+2);
        case 10090: 
            return Twapi_GetThemeFont(interp, objc-2, objv+2);
        case 10091:
            return Twapi_WriteShortcut(interp, objc-2, objv+2);
        case 10092:
            return Twapi_ReadShortcut(interp, objc-2, objv+2);
        case 10093:
            return Twapi_InvokeUrlShortcut(interp, objc-2, objv+2);
        case 10094: // SHInvokePrinterCommand
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(hwnd), GETINT(dw),
                             GETWSTR(s), GETWSTR(s2), GETBOOL(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SHInvokePrinterCommandW(hwnd, dw, s, s2, dw2);
            break;
        case 10095:
            return Twapi_SHFileOperation(interp, objc-2, objv+2);
        case 10096:
            return Twapi_ShellExecuteEx(interp, objc-2, objv+2);
        case 10097:
            secattrP = NULL;        /* Even on error, it might be filled */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETBOOL(dw), GETNULLIFEMPTY(s),
                             ARGEND) == TCL_OK) {
                result.type = TRT_HANDLE;
                result.value.hval = CreateMutexW(secattrP, dw, s);
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
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            result.value.hval = (func == 10098 ? OpenMutexW : OpenSemaphoreW)
                (dw, dw2, s);
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
        case 10107:
        case 10108:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETPTR(pv, TWAPI_FILEVERINFO),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return (func == 10107 
                    ? Twapi_VerQueryValue_FIXEDFILEINFO
                    : Twapi_VerQueryValue_TRANSLATIONS)(interp, pv);
        case 10109:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETPTR(pv, TWAPI_FILEVERINFO),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            Twapi_FreeFileVersionInfo(pv);
            break;

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

        case 10117:
            return Twapi_NPipeServer(ticP, objc-2, objv+2);
        case 10118:
            return Twapi_NPipeClient(ticP, objc-2, objv+2);
        case 10119: // IsEqualPtr
            if (objc != 4)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
            if (ObjToOpaque(interp, objv[2], &pv, NULL) != TCL_OK ||
                ObjToOpaque(interp, objv[3], &pv2, NULL) != TCL_OK) {
                return TCL_ERROR;
            }
            result.type = TRT_BOOL;
            result.value.bval = (pv == pv2);
            break;
        case 10120: // IsNullPtr
            if (objc < 3 || objc > 4)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
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
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
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
        case 10123:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETHANDLE(h2),
                             ARGEND) == TCL_OK) {
                result.type = TRT_EXCEPTION_ON_FALSE;
                result.value.ival = NotifyChangeEventLog(h, h2);
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            break;
        case 10124:
            return Twapi_UpdateResource(interp, objc-2, objv+2);
        case 10125:
            return Twapi_FindResourceEx(interp, objc-2, objv+2);
        case 10126:
            return Twapi_LoadResource(interp, objc-2, objv+2);
        case 10127:
            return Twapi_EnumResourceNames(interp, objc-2, objv+2);
        case 10128:
            return Twapi_EnumResourceLanguages(interp, objc-2, objv+2);
        case 10129:
            return Twapi_SplitStringResource(interp, objc-2, objv+2);
#ifndef TWAPI_LEAN
        case 10130:
            return Twapi_GetProcessList(ticP, objc-2, objv+2);
#endif

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
        RPC_STATUS rpc_status;
        DWORD_PTR dwp;
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
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1:
            result.type = TRT_BOOL;
            result.value.bval = IsClipboardFormatAvailable(dw);
            break;
        case 2:
            result.type = TRT_HANDLE;
            result.value.hval = GetClipboardData(dw);
            break;
        case 3:
            result.value.unicode.len = GetClipboardFormatNameW(dw, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
            result.value.unicode.str = u.buf;
            result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
            break;
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
        case 5:
            result.value.ival = SetConsoleCP(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 6:
            result.value.ival = SetConsoleOutputCP(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 7:
            result.value.unicode.len = VerLanguageNameW(dw, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
            result.value.unicode.str = u.buf;
            result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
            break;
        case 8:
            return Twapi_GetPerAdapterInfo(ticP, dw);
        case 9:
            return Twapi_GetIfEntry(interp, dw);
        case 10:
            return Twapi_GetIfTable(ticP, dw);
        case 11:
            return Twapi_GetIpAddrTable(ticP, dw);
        case 12:
            return Twapi_GetIpNetTable(ticP, dw);
        case 13:
            return Twapi_GetIpForwardTable(ticP, dw);
        case 14:
            result.value.ival = FlushIpNetTable(dw);
            result.type = TRT_EXCEPTION_ON_ERROR;
            break;
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
            break;
        case 17:
            result.type = TRT_DWORD;
            result.value.ival = GetSystemMetrics(dw);
            break;
        case 18:
            result.type = TRT_EMPTY;
            Sleep(dw);
            break;
        case 19:
            result.type = TRT_DWORD;
            result.value.ival = SetThreadExecutionState(dw);
            break;
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
        case 24:
            result.type = TRT_DWORD;
            result.value.ival = GetAsyncKeyState(dw);
            break;
        case 25:
            result.type = TRT_DWORD;
            result.value.ival = GetKeyState(dw);
            break;
        case 26:
            result.value.hval = GetThreadDesktop(dw);
            result.type = TRT_HDESK;
            break;
        case 27:
            return Twapi_BlockInput(interp, dw);
        case 28:
            return Twapi_UnregisterHotKey(ticP, dw);
        case 29:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetCaretBlinkTime(dw);
            break;
        case 30:
            result.type = TRT_BOOL;
            result.value.bval = MessageBeep(dw);
            break;
        case 31:
            return Twapi_GetGUIThreadInfo(interp, dw);
        case 32:
            return Twapi_UnregisterDeviceNotification(ticP, dw);
        case 33:
            Sleep(dw);
            result.type = TRT_EMPTY;
            break;
        case 34:
            result.value.obj = Twapi_MapWindowsErrorToString(dw);
            result.type = TRT_OBJ;
            break;
        case 35:
            result.type = ProcessIdToSessionId(dw, &result.value.ival) ? TRT_DWORD 
                : TRT_GETLASTERROR;
            break;
//        case 36: // UNUSED
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
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

        CHECK_INTEGER_OBJ(interp, dw2, objv[3]);
        switch (func) {
//      case 1001: UNUSED            
        case 1002:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = Beep(dw, dw2);
            break;
        case 1003:
            result.type = TRT_DWORD;
            result.value.ival = MapVirtualKey(dw, dw2);
            break;
        case 1004:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetCaretPos(dw, dw2);
            break;
        case 1005:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetCursorPos(dw, dw2);
            break;
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
        case 1008:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = GenerateConsoleCtrlEvent(dw, dw2);
            break;
        case 1009:
            return Twapi_AllocateAndGetTcpExTableFromStack(ticP, dw, dw2);
        case 1010:
            return Twapi_AllocateAndGetUdpExTableFromStack(ticP, dw, dw2);
        }
    } else if (func < 3000) {
        /* Check we have exactly two more integer arguments */
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 2001:
            result.type = TRT_HANDLE;
            result.value.hval = OpenProcess(dw, dw2, dw3);
            break;
        case 2002:
            result.type = TRT_HANDLE;
            result.value.hval = OpenThread(dw, dw2, dw3);
            break;
        case 2003:
            result.type = TRT_HDESK;
            result.value.hval = OpenInputDesktop(dw, dw2, dw3);
            break;
        case 2004:
            result.type = TRT_BOOL;
            result.value.bval = AttachThreadInput(dw, dw2, dw3);
            break;
        case 2005:
            return Twapi_RegisterHotKey(ticP, dw, dw2, dw3);
        }
    } else {
        /* Any number (> 0) of additional arguments */
        if (objc < 4)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

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
        case 10003:
            if (ObjToHANDLE(interp, objv[3], &u.h) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            result.value.hval = SetClipboardData(dw, u.h);
            break;
        case 10004:
            if (ObjToHANDLE(interp, objv[3], &u.h) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetStdHandle(dw, u.h);
            break;
        case 10005: // UNUSED
        case 10006: // CreateConsoleScreenBuffer
            if (TwapiGetArgs(interp, objc-3, objv+3,
                             GETINT(dw2),
                             GETVAR(u.secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            result.value.hval = CreateConsoleScreenBuffer(dw, dw2, u.secattrP, dw3, NULL);
            TwapiFreeSECURITY_ATTRIBUTES(u.secattrP);
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
        case 1:
            result.type = TRT_NONZERO_RESULT;
            result.value.ival = RegisterClipboardFormatW(arg);
            break;
        case 2:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleTitleW(arg);
            break;
        case 3:
            result.type = TRT_DWORD;
            result.value.ival = GetDriveTypeW(arg);
            break;
        case 4:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeleteVolumeMountPointW(arg);
            break;
        case 5:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = GetVolumeNameForVolumeMountPointW(
                arg, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
            break;
        case 6:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = GetVolumePathNameW(
                arg, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
            break;
        case 7:
            result.type = GetAdapterIndex((LPWSTR)arg, &result.value.ival)
                ? TRT_GETLASTERROR
                : TRT_DWORD;
            break;
        case 8:
            return Twapi_CommandLineToArgv(interp, arg);
        case 9:
            return Twapi_WNetGetUniversalName(ticP, arg);
        case 10:
            return Twapi_WNetGetUser(interp, arg);
        case 11:
            return Twapi_AppendLog(interp, arg);
        case 12:
            result.type = TRT_HANDLE;
            result.value.hval = WTSOpenServerW(arg);
            break;
        case 13:
            return Twapi_ReadUrlShortcut(interp, arg);
        case 14:
            NULLIFY_EMPTY(arg);
            return Twapi_GetVolumeInformation(interp, arg);
        case 15:
            return TwapiFirstVolume(interp, arg); // FindFirstVolumeMountPoint
        case 16:
            return Twapi_GetDiskFreeSpaceEx(interp, arg);
        case 17:
            result.type = TRT_EXCEPTION_ON_FALSE;
            NULLIFY_EMPTY(arg);
            result.value.ival = AbortSystemShutdownW(arg);
        case 18:
            NULLIFY_EMPTY(arg);
            return Twapi_GetPrivateProfileSectionNames(ticP, arg);
        case 19:
            return Twapi_SendInput(ticP, objv[2]);
        case 20:
            return Twapi_SendUnicode(ticP, objv[2]);
        case 21:
            result.value.opaque.p = Twapi_GetFileVersionInfo(arg);
            if (result.value.opaque.p) {
                result.type = TRT_OPAQUE;
                result.value.opaque.name = "TWAPI_FILEVERINFO";
            } else
                result.type = TRT_GETLASTERROR;
            break;
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
                result.value.obj = Tcl_NewUnicodeObj(bufP, dw-1);
            }
            if (bufP != u.buf)
                TwapiFree(bufP);
            break;
        case 23: // GlobalAddAtom
            result.value.ival = GlobalAddAtomW(arg);
            result.type = result.value.ival ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        }
    } else if (func < 1000) {
        /* One additional integer argument */

        if (objc != 4)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
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
#ifndef TWAPI_LEAN
        case 502: // LsaOpenPolicy
            ObjToLSA_UNICODE_STRING(objv[2], &lsa_ustr);
            ZeroMemory(&lsa_oattr, sizeof(lsa_oattr));
            dw2 = LsaOpenPolicy(&lsa_ustr, &lsa_oattr, dw, &result.value.hval);
            if (dw2 == STATUS_SUCCESS) {
                result.type = TRT_LSA_HANDLE;
            } else {
                result.type = TRT_NTSTATUS;
                result.value.ival = dw2;
            }
            break;
#endif
        case 503: // NetFileClose
            NULLIFY_EMPTY(arg);
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = NetFileClose(arg, dw);
            break;
        case 504: // LoadLibrary
            result.type = TRT_HANDLE;
            result.value.hval = LoadLibraryExW(arg, NULL, dw);
            break;
        case 505: // AddFontResourceEx
            /* TBD - Tcl wrapper for AddFontResourceEx ? */
            result.type = TRT_DWORD;
            result.value.ival = AddFontResourceExW(arg, dw, NULL);
            break;
        case 506: // BeginUpdateResource
            result.type = TRT_HANDLE;
            result.value.hval = BeginUpdateResourceW(arg, dw);
            break;
        }
    } else if (func < 2000) {

        /* Commands with exactly two additional integer argument */
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        
        switch (func) {
        case 1001:
            result.type = TRT_HWINSTA;
            result.value.hval = OpenWindowStationW(arg, dw, dw2);
            break;
        case 1002:
            result.type = TRT_EXCEPTION_ON_WNET_ERROR;
            result.value.ival = WNetCancelConnection2W(arg, dw, dw2);
            break;
        case 1003:
            NULLIFY_EMPTY(arg);
            return Twapi_NetFileGetInfo(interp, arg, dw, dw2);
        case 1004:
            return Twapi_GetNamedSecurityInfo(interp, arg, dw, dw2);
        case 1005:
            bufP = u.buf;
            dw3 = ARRAYSIZE(u.buf);
            if (! TranslateNameW(arg, dw, dw2, bufP, &dw3)) {
                result.value.ival = GetLastError();
                if (result.value.ival != 0) {
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
        CONSOLE_SCREEN_BUFFER_INFO csbi;
        COORD coord;
        WCHAR buf[MAX_PATH+1];
        TWAPI_TOKEN_MANDATORY_POLICY ttmp;
        TWAPI_TOKEN_MANDATORY_LABEL ttml;
        MODULEINFO moduleinfo;
        SERVICE_STATUS svcstatus;
        LSA_UNICODE_STRING lsa_ustr;
        SECURITY_ATTRIBUTES *secattrP;
        MONITORINFOEXW minfo;
        RECT rect;
        MemLifo *lifoP;
    } u;
    int func;
    int i;
    FILETIME ft[3];
    FILETIME *ftP[3];
    Tcl_Obj *objs[3];
    void *pv;
    RECT *rectP;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETHANDLE(h),
                     ARGTERM) != TCL_OK) {
        return TCL_ERROR;
    }

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 1000) {
        /* Command with a single handle argument */
        if (objc != 3)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FlushConsoleInputBuffer(h);
            break;
        case 2:
            result.type = GetConsoleMode(h, &result.value.ival)
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 3:
            if (GetConsoleScreenBufferInfo(h, &u.csbi) == 0)
                result.type = TRT_GETLASTERROR;
            else {
                Tcl_SetObjResult(interp, ObjFromCONSOLE_SCREEN_BUFFER_INFO(interp, &u.csbi));
                return TCL_OK;
            }
            break;
        case 4:
            u.coord = GetLargestConsoleWindowSize(h);
            Tcl_SetObjResult(interp, ObjFromCOORD(interp, &u.coord));
            return TCL_OK;
        case 5:
            result.type = GetNumberOfConsoleInputEvents(h, &result.value.ival) ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 6:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleActiveScreenBuffer(h);
            break;
        case 7:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FindVolumeClose(h);
            break;
        case 8:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FindVolumeMountPointClose(h);
            break;
        case 9:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeregisterEventSource(h);
            break;
        case 10:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseEventLog(h);
            break;
        case 11:
            result.type = GetNumberOfEventLogRecords(h,
                                                     &result.value.ival)
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 12:
            result.type = GetOldestEventLogRecord(h,
                                                  &result.value.ival) 
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 13:
            result.type = Twapi_IsEventLogFull(h,
                                               &result.value.ival) 
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
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

#ifndef TWAPI_LEAN
        case 17:
            return Twapi_EnumProcessModules(ticP, h);
        case 18:
            result.type = Twapi_IsWow64Process(h, &result.value.bval)
                ? TRT_BOOL : TRT_GETLASTERROR;
            break;
        case 19:
            result.type = TRT_EXCEPTION_ON_MINUSONE;
            result.value.ival = ResumeThread(h);
            break;
        case 20:
            result.type = TRT_EXCEPTION_ON_MINUSONE;
            result.value.ival = SuspendThread(h);
            break;
        case 21:
            result.type = TRT_NONZERO_RESULT;
            result.value.ival = GetPriorityClass(h);
            break;
        case 22:
            return Twapi_NtQueryInformationProcessBasicInformation(interp, h);
        case 23:
            return Twapi_NtQueryInformationThreadBasicInformation(interp, h);
        case 24:
            result.value.ival = GetThreadPriority(h);
            result.type = result.value.ival == THREAD_PRIORITY_ERROR_RETURN
                ? TRT_GETLASTERROR : TRT_DWORD;
            break;
#endif // TWAPI_LEAN
        
        case 25:
            result.type = GetExitCodeProcess(h, &result.value.ival)
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 26:
            result.value.ival = LsaClose(h);
            result.type = TRT_DWORD;
            break;
        case 27:
            result.value.ival = ImpersonateLoggedOnUser(h);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 28:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeleteService(h);
            break;
        case 29:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseServiceHandle(h);
            break;
        case 30:
            result.type = TRT_EMPTY;
            Twapi_CloseThemeData(h);
            break;
        case 31:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseDesktop(h);
            break;
        case 32:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SwitchDesktop(h);
            break;
        case 33:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetThreadDesktop(h);
            break;
        case 34:
            return Twapi_EnumDesktops(interp, h);
        case 35:
            return Twapi_EnumDesktopWindows(interp, h);
        case 36:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseWindowStation(h);
            break;
        case 37:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetProcessWindowStation(h);
            break;
        case 38:
            return TwapiNextVolume(interp, 0, h);
        case 39:
            return TwapiNextVolume(interp, 1, h);
            break;
        case 40:
            return Twapi_GetFileType(interp, h);
        case 41:
            return Twapi_QueryServiceConfig(ticP, h);
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
        case 52:
            if (GetFileTime(h, &ft[0], &ft[1], &ft[2])) {
                objs[0] = ObjFromFILETIME(&ft[0]);
                objs[1] = ObjFromFILETIME(&ft[1]);
                objs[2] = ObjFromFILETIME(&ft[2]);
                result.type = TRT_OBJV;
                result.value.objv.objPP = objs;
                result.value.objv.nobj = 3;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 53: 
            return Twapi_UnregisterDirectoryMonitor(ticP, h);
        case 54:
            MemLifoClose(h);
            TwapiFree(h);
            result.type = TRT_EMPTY;
            break;
        case 55:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = MemLifoPopFrame((MemLifo *)h);
            break;
        case 56:
            result.type = TRT_EMPTY;
            Twapi_Free_SEC_WINNT_AUTH_IDENTITY(h);
            break;
        case 57:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetupDiDestroyDeviceInfoList(h);
            break;
        case 58:
            u.minfo.cbSize = sizeof(u.minfo);
            if (GetMonitorInfoW(h, (MONITORINFO *)&u.minfo)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromMONITORINFOEX((MONITORINFO *)&u.minfo);
            } else
                result.type = TRT_GETLASTERROR;
            break;
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
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ImpersonateNamedPipeClient(h);
            break;
//        case 65: UNUSED
        case 66:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetEvent(h);
            break;
        case 67:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ResetEvent(h);
            break;
        case 68:
            return Twapi_EnumResourceTypes(interp, h);
        }
    } else if (func < 2000) {

        // A single additional DWORD arg is present
        if (objc != 4)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[3]);

        switch (func) {
#ifndef TWAPI_LEAN
        case 1001:
            result.type = ReleaseSemaphore(h, dw, &result.value.ival) ?
                TRT_DWORD : TRT_GETLASTERROR;
            break;
#endif
        case 1002:
            result.type = TRT_EXCEPTION_ON_FALSE;
            /* svcstatus is not returned because it is not always filled
               in and is not very useful even when it is */
            result.value.ival = ControlService(h, dw, &u.svcstatus);
            break;
        case 1003:
            return Twapi_EnumDependentServices(ticP, h, dw);
        case 1004:
            return Twapi_QueryServiceStatusEx(interp, h, dw);
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

        case 1009:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = TerminateProcess(h, dw);
            break;
        case 1010:
            return Twapi_WaitForInputIdle(interp, h, dw);
        case 1011:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetPriorityClass(h, dw);
            break;
        case 1012:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetThreadPriority(h, dw);
            break;
        case 1013:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleMode(h, dw);
            break;
        case 1014:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleTextAttribute(h, (WORD) dw);
            break;
        case 1015:
            return Twapi_ReadConsole(ticP, h, dw);
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
        case 1020: // EndUpdateResource
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EndUpdateResourceW(h, dw);
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
        case 2006:
            return Twapi_ReadEventLog(ticP, h, dw, dw2);
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
        case 10001:
            if (objc != 6)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
            for (i = 0; i < 3; ++i) {
                if (Tcl_GetCharLength(objv[3+i]) == 0)
                    ftP[i] = NULL;
                else {
                    if (ObjToFILETIME(interp, objv[3+i], &ft[i]) != TCL_OK)
                        return TCL_ERROR;
                    ftP[i] = &ft[i];
                }
            }
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetFileTime(h, ftP[0], ftP[1], ftP[2]);
            break;

#ifndef TWAPI_LEAN
        case 10002: // SetThreadToken
            if (objc != 4)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
            
            if (ObjToHANDLE(interp, objv[3], &h2) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetThreadToken(h, h2);
            break;
        case 10003:
            if (objc != 4)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
            ObjToLSA_UNICODE_STRING(objv[3], &u.lsa_ustr);
            return Twapi_LsaEnumerateAccountsWithUserRight(interp, h,
                                                           &u.lsa_ustr);
#endif
#ifndef TWAPI_LEAN
        case 10004:
            if (objc != 4)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
            if (ObjToSID_AND_ATTRIBUTES(interp, objv[3], &u.ttml.Label) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetTokenInformation(h,
                                                    TwapiTokenIntegrityLevel,
                                                    &u.ttml, sizeof(u.ttml));
            if (u.ttml.Label.Sid)
                TwapiFree(u.ttml.Label.Sid);
            break;
#endif
        case 10005:
            if (objc != 4)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
            rectP = &u.rect;
            if (ObjToRECT_NULL(interp, objv[3], &rectP) != TCL_OK)
                return TCL_ERROR;
            return Twapi_EnumDisplayMonitors(interp, h, rectP);
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
                     GETINT(func), GETNULLIFEMPTY(s1), ARGUSEDEFAULT,
                     GETWSTR(s2), GETWSTR(s3), GETINT(dw),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;

    if (func == 0 || (func > 1000 && func < 2000)) {
        switch (func) {
        case 1005: // This block of function codes maps directly to
        case 1008: // the function codes accepted by Twapi_NetUserSetInfoDWORD.
        case 1010: // A bit klugy but the easiest way to keep backwards
        case 1017: // compatibility.
        case 1024: 
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
            result.value.ival = Twapi_NetUserSetInfoLPWSTR(func, s1, s2, s3);
            break;

        }
    }

    if (result.type != TRT_BADFUNCTIONCODE)
        return TwapiSetResult(interp, &result);

    switch (func) {
        // NOTE case 0: is defined above with the Twapi_NetUserSetInfoLPWSTR 
        // section.
    case 1:
        return Twapi_LookupAccountName(interp, s1, s2);
#ifndef TWAPI_LEAN
    case 2: // NetUserAdd
        if (TwapiGetArgs(interp, objc-6, objv+6,
                         GETNULLIFEMPTY(s4), GETNULLIFEMPTY(s5),
                         GETINT(dw2), GETNULLIFEMPTY(s6),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_NetUserAdd(interp, s1, s2, s3, dw, s4, s5, dw2, s6);
#endif

    case 3: // NetShareSetInfo
        if (objc != 7)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

        if (ObjToPSECURITY_DESCRIPTOR(interp, objv[6], &u.secdP) != TCL_OK)
            return TCL_ERROR;
        /* u.secdP may be NULL */
        result.value.ival = Twapi_NetShareSetInfo(interp, s1, s2, s3, dw, u.secdP);
        if (u.secdP)
            TwapiFreeSECURITY_DESCRIPTOR(u.secdP);
        return result.value.ival;

    case 4: 
        EMPTIFY_NULL(s1);
        result.type = TRT_SEC_WINNT_AUTH_IDENTITY;
        result.value.hval = Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY(s1, s2, s3);
        break;

#ifndef TWAPI_LEAN
    case 5:
        if (objc != 7)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw2, objv[6]);
        EMPTIFY_NULL(s1);      /* Username must not be NULL - pass as "" */
        NULLIFY_EMPTY(s2);      /* Domain - NULL if empty string */
        if (LogonUserW(s1, s2, s3, dw,dw2, &result.value.hval))
            result.type = TRT_HANDLE;
        else
            result.type = TRT_GETLASTERROR;
        break;
    case 6: // CreateDesktopW
        if (objc != 8)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw2, objv[6]);
        if (ObjToPSECURITY_ATTRIBUTES(interp, objv[6], &u.secattrP) != TCL_OK)
            return TCL_ERROR;
        /* Note s1, s2 are ignored and are reserved as NULL */
        result.type = TRT_HDESK;
        result.value.hval = CreateDesktopW(s1, NULL, NULL, dw, dw2, u.secattrP);
        if (u.secattrP)
            TwapiFreeSECURITY_ATTRIBUTES(u.secattrP);
        break;
#endif
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

#ifndef TWAPI_LEAN

    case 10:
        return Twapi_NetUserGetInfo(interp, s1, s2, dw);
    case 11:
        return Twapi_NetGroupGetInfo(interp, s1, s2, dw);
    case 12:
        return Twapi_NetLocalGroupGetInfo(interp, s1, s2, dw);

#endif // TWAPI_LEAN

    case 13:
        result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
        if (LookupPrivilegeDisplayNameW(s1,s2,u.buf,&result.value.unicode.len,&dw)) {
            result.value.unicode.str = u.buf;
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
#ifndef TWAPI_LEAN
    case 15:
        u.gi1.grpi1_name = s2;
        NULLIFY_EMPTY(s3);
        u.gi1.grpi1_comment = s3;
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetGroupAdd(s1, 1, (LPBYTE)&u.gi1, NULL);
        break;
    case 16:
        u.lgi1.lgrpi1_name = s2;
        NULLIFY_EMPTY(s3);
        u.lgi1.lgrpi1_comment = s3;
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetLocalGroupAdd(s1, 1, (LPBYTE)&u.lgi1, NULL);
        break;
    case 17:
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetGroupDel(s1,s2);
        break;
    case 18:
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetLocalGroupDel(s1,s2);
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
    case 22:
        u.lgmi3.lgrmi3_domainandname = s3;
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetLocalGroupDelMembers(s1, s2, 3, (LPBYTE) &u.lgmi3, 1);
        break;
#endif // TWAPI_LEAN
    case 23:
        result.type = TRT_SC_HANDLE;
        NULLIFY_EMPTY(s2);
        result.value.hval = OpenSCManagerW(s1, s2, dw);
        break;
    case 24:
        return Twapi_NetUseGetInfo(interp, s1, s2, dw);
    case 25:
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetShareDel(s1,s2,dw);
        break;
    case 26:
        return Twapi_NetShareCheck(interp, s1, s2);
    case 27:
        return Twapi_NetShareGetInfo(interp, s1,s2,dw);
//  case 28: UNUSED
    case 29: // NetUserDel
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = NetUserDel(s1, s2);
        break;
#ifndef TWAPI_LEAN
//  case 30: UNUSED
//  case 31: UNUSED
#endif
    case 32:
        return Twapi_NetSessionGetInfo(interp, s1,s2,s3,dw);

    case 33:
        result.type = TRT_EXCEPTION_ON_ERROR;
        NULLIFY_EMPTY(s2);
        NULLIFY_EMPTY(s3);
        result.value.ival = NetSessionDel(s1,s2,s3);
        break;
    case 34:
        NULLIFY_EMPTY(s2);
        return Twapi_NetGetDCName(interp, s1,s2);
//  case 35: UNUSED
    case 36:
        /* Note first param is s2, second is s1 */
        return Twapi_WNetGetResourceInformation(ticP, s2,s1,dw);
    case 37:
        if (s1 == NULL) s1 = L"";
        return Twapi_WriteUrlShortcut(interp, s1, s2, dw);
    case 38:
        NULLIFY_EMPTY(s2);
        result.type = TRT_HWND;
        result.value.hwin = FindWindowW(s1,s2);
        break;
    case 39:
        NULLIFY_EMPTY(s2);
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = MoveFileExW(s1,s2,dw);
        break;
    case 40:
        NULLIFY_EMPTY(s2);
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SetVolumeLabelW(s1,s2);
        break;
    case 41:
        return Twapi_QueryDosDevice(ticP, s1);
    case 42:
        result.type = TRT_EXCEPTION_ON_FALSE;
        // Note we use s2, s3 here as s1 has LPWSTR_NULL_IF_EMPTY semantics
        result.value.ival = DefineDosDeviceW(dw, s2, s3);
        break;
    case 43:
        result.type = TRT_EXCEPTION_ON_FALSE;
        // Note we use s2, s3 here as s1 has LPWSTR_NULL_IF_EMPTY semantics
        result.value.ival = SetVolumeMountPointW(s2, s3);
        break;
    case 44:
        result.type = TRT_HANDLE;
        result.value.hval = RegisterEventSourceW(s1, s2);
        break;
    case 45:
        result.type = TRT_HANDLE;
        result.value.hval = OpenEventLogW(s1, s2);
        break;
    case 46:
        result.type = TRT_HANDLE;
        result.value.hval = OpenBackupEventLogW(s1, s2);
        break;

#ifndef TWAPI_LEAN
    case 47:
        /* TBD - is there a Tcl wrapper for CreateScalableFontResource ? */
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CreateScalableFontResourceW(dw, s1,s2,s3);
        break;
    case 48:
        /* TBD - Tcl wrapper for RemoveFontResourceEx ? */
        result.type = TRT_BOOL;
        result.value.bval = RemoveFontResourceExW(s2, dw, NULL);
        break;
#endif

    }

    return TwapiSetResult(interp, &result);
}

int Twapi_CallWObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HWND hwnd, hwnd2;
    TwapiResult result;
    DWORD dw, dw2, dw3, dw4, dw5;
    Tcl_Obj *objs[2];
    int func;
    union {
        WINDOWPLACEMENT winplace;
        WINDOWINFO wininfo;
        WCHAR buf[MAX_PATH+1];
        RECT   rect;
    } u;
    RECT *rectP;
    LPWSTR s, s2;
    HANDLE h;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETHWND(hwnd),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.type = TRT_BOOL;
        result.value.bval = IsIconic(hwnd);
        break;
    case 2:
        result.type = TRT_BOOL;
        result.value.bval = IsZoomed(hwnd);
        break;
    case 3:
        result.type = TRT_BOOL;
        result.value.bval = IsWindowVisible(hwnd);
        break;
    case 4:
        result.type = TRT_BOOL;
        result.value.bval = IsWindow(hwnd);
        break;
    case 5:
        result.type = TRT_BOOL;
        result.value.bval = IsWindowUnicode(hwnd);
        break;
    case 6:
        result.type = TRT_BOOL;
        result.value.bval = IsWindowEnabled(hwnd);
        break;
    case 7:
        result.type = TRT_BOOL;
        result.value.bval = ArrangeIconicWindows(hwnd);
        break;
    case 8:
        result.type = TRT_BOOL;
        result.value.bval = SetForegroundWindow(hwnd);
        break;
    case 9:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = OpenIcon(hwnd);
        break;
    case 10:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CloseWindow(hwnd);
        break;
    case 11:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = DestroyWindow(hwnd);
        break;
    case 12:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = UpdateWindow(hwnd);
        break;
    case 13:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = HideCaret(hwnd);
        break;
    case 14:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = ShowCaret(hwnd);
        break;
    case 15:
        result.type = TRT_HWND;
        result.value.hwin = GetParent(hwnd);
        break;
    case 16:
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = OpenClipboard(hwnd);
        break;
    case 17:
        result.type = GetClientRect(hwnd, &result.value.rect) ? TRT_RECT : TRT_GETLASTERROR;
        break;
    case 18:
        result.type = GetWindowRect(hwnd, &result.value.rect) ? TRT_RECT : TRT_GETLASTERROR;
        break;
    case 19:
        result.type = TRT_HDC;
        result.value.hval = GetDC(hwnd);
        break;
    case 20:
        result.type = TRT_HWND;
        result.value.hwin = SetFocus(hwnd);
        break;
    case 21:
        result.value.hwin = SetActiveWindow(hwnd);
        result.type = result.value.hwin ? TRT_HWND : TRT_GETLASTERROR;
        break;
    case 22:
        result.value.unicode.len = GetClassNameW(hwnd, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
        result.value.unicode.str = u.buf;
        result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
        break;
    case 23:
        result.value.unicode.len = RealGetWindowClassW(hwnd, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
        result.value.unicode.str = u.buf;
        result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
        break;
    case 24:
        dw2 = GetWindowThreadProcessId(hwnd, &dw);
        if (dw2 == 0) {
            result.type = TRT_GETLASTERROR;
        } else {
            objs[0] = Tcl_NewLongObj(dw2);
            objs[1] = Tcl_NewLongObj(dw);
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            result.type = TRT_OBJV;
        }
        break;
    case 25:
        SetLastError(0);            /* Make sure error is not set */
        result.type = TRT_UNICODE;
        result.value.unicode.len = GetWindowTextW(hwnd, u.buf, sizeof(u.buf)/sizeof(u.buf[0]));
        result.value.unicode.str = u.buf;
        /* Distinguish between error and empty string when count is 0 */
        if (result.value.unicode.len == 0 && GetLastError()) {
            result.type = TRT_GETLASTERROR;
        }
        break;
    case 26:
        result.type = TRT_HDC;
        result.value.hval = GetWindowDC(hwnd);
        break;
//  case 27: NOT USED
    case 28:
        return Twapi_EnumChildWindows(interp, hwnd);

//  case 29: NOT USED

    case 30:
        if (GetWindowPlacement(hwnd, &u.winplace)) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromWINDOWPLACEMENT(&u.winplace);
            break;
        } else {
            result.type = TRT_GETLASTERROR;
        }
        break;
    case 31:
        if (GetWindowInfo(hwnd, &u.wininfo)) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromWINDOWINFO(&u.wininfo);
            break;
        } else {
            result.type = TRT_GETLASTERROR;
        }
        break;
    }

    if (result.type != TRT_BADFUNCTIONCODE)
        return TwapiSetResult(interp, &result);

    /* At least one additional arg */
    if (objc < 4)
        return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

    switch (func) {
    case 1001: // SetWindowText
        result.value.ival = SetWindowTextW(hwnd, Tcl_GetUnicode(objv[3]));
        result.type = TRT_EXCEPTION_ON_FALSE;
        break;
    case 1002: // IsChild
        if (ObjToHWND(interp, objv[3], &hwnd2) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.bval = IsChild(hwnd, hwnd2);
        break;
    case 1003: //SetWindowPlacement
        if (ObjToWINDOWPLACEMENT(interp, objv[3], &u.winplace) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SetWindowPlacement(hwnd, &u.winplace);
        break;
    case 1004: // InvalidateRect
        rectP = &u.rect;
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETVAR(rectP, ObjToRECT_NULL), GETINT(dw), 
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = InvalidateRect(hwnd, rectP, dw);
        break;
    case 1005: // SetWindowPos
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETHANDLET(hwnd2, HWND), GETINT(dw), GETINT(dw2),
                         GETINT(dw3), GETINT(dw4), GETINT(dw5),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SetWindowPos(hwnd, hwnd2, dw, dw2, dw3, dw4, dw5);
        break;
    case 1006: // FindWindowEx
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETHANDLET(hwnd2, HWND),
                         GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HWND;
        result.value.hval = FindWindowExW(hwnd, hwnd2, s, s2);
        break;
    case 1007:
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETHANDLET(h, HDC),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_DWORD;
        result.value.ival = ReleaseDC(hwnd, h);
        break;
    case 1008:
        result.type = TRT_HANDLE;
        result.value.hval = Twapi_OpenThemeData(hwnd, Tcl_GetUnicode(objv[3]));
        break;
    }

    return TwapiSetResult(interp, &result);
}

int Twapi_CallWUObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HWND hwnd;
    TwapiResult result;
    int func;
    DWORD dw, dw2, dw3, dw4, dw5;
    DWORD_PTR dwp, dwp2;
    LPWSTR s, s2;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETHANDLET(hwnd, HWND), GETINT(dw),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 1000) {
        switch (func) {
        case 1:
            result.value.hwin = GetAncestor(hwnd, dw);
            result.type = TRT_HWND;
            break;
        case 2:
            result.value.hwin = GetWindow(hwnd, dw);
            result.type = TRT_HWND;
            break;
        case 3:
            result.value.bval = ShowWindow(hwnd, dw);
            result.type = TRT_BOOL;
            break;
        case 4:
            result.value.bval = ShowWindowAsync(hwnd, dw);
            result.type = TRT_BOOL;
            break;
        case 5:
            result.value.bval = EnableWindow(hwnd, dw);
            result.type = TRT_BOOL;
            break;
        case 6:
            result.value.ival = ShowOwnedPopups(hwnd, dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 7:
            result.type = TRT_HMONITOR;
            result.value.hval = MonitorFromWindow(hwnd, dw);
            break;
        case 8:
            SetLastError(0);    /* Avoid spurious errors when checking GetLastError */
            result.value.dwp = GetWindowLongPtrW(hwnd, dw);
            if (result.value.dwp || GetLastError() == 0)
                result.type = TRT_DWORD_PTR;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 9:
            dw = SHGetSpecialFolderLocation(hwnd, dw, &result.value.pidl);
            if (dw == 0)
                result.type = TRT_PIDL;
            else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw;
            }
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
        case 10001:
        if (TwapiGetArgs(interp, objc-4, objv+4,
                         GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SetLayeredWindowAttributes(hwnd, dw, (BYTE)dw2, dw3);
        break;
        case 10002:
        if (TwapiGetArgs(interp, objc-4, objv+4,
                         GETINT(dw2), GETINT(dw3),
                         GETINT(dw4), GETINT(dw5),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = MoveWindow(hwnd, dw, dw2, dw3, dw4, dw5);
            break;
        case 10003:
            if (objc != 5)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
            if (ObjToDWORD_PTR(interp, objv[4], &dwp) != TCL_OK)
                return TCL_ERROR;
            result.type = Twapi_SetWindowLongPtr(hwnd, dw, dwp, &result.value.dwp)
                ? TRT_DWORD_PTR : TRT_GETLASTERROR;
            break;
        case 10004:
            if (TwapiGetArgs(interp, objc-4, objv+4,
                             GETWSTR(s), GETNULLIFEMPTY(s2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = Twapi_SHObjectProperties(hwnd, dw, s, s2);
            break;
        }            
    }

    return TwapiSetResult(interp, &result);
}


int Twapi_CallHSUObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE h;
    LPCWSTR s;
    DWORD dw;
    TwapiResult result;
    int func;
    union {
        WCHAR buf[MAX_PATH+1];
        COORD coord;
    } u;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETHANDLE(h), GETWSTR(s),
                     ARGUSEDEFAULT,
                     GETINT(dw),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_EXCEPTION_ON_FALSE; /* Likely result type */
    switch (func) {
    case 1:
        result.value.ival = BackupEventLogW(h, s);
        break;
    case 2:
        NULLIFY_EMPTY(s);
        result.value.ival = ClearEventLogW(h, s);
        break;
    case 3:
        result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
        if (GetServiceKeyNameW(h, s, u.buf, &result.value.unicode.len)) {
            result.value.unicode.str = u.buf;
            result.type = TRT_UNICODE;
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 4:
        result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
        if (GetServiceDisplayNameW(h, s, u.buf, &result.value.unicode.len)) {
            result.value.unicode.str = u.buf;
            result.type = TRT_UNICODE;
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 5: // WriteConsole
        if (WriteConsoleW(h, s, dw, &result.value.ival, NULL))
            result.type = TRT_DWORD;
        else
            result.type = TRT_GETLASTERROR;
        break;
    case 6:
        if (objc != 6)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
        if (ObjToCOORD(interp, objv[5], &u.coord) != TCL_OK)
            return TCL_ERROR;
        if (FillConsoleOutputCharacterW(h, s[0], dw, u.coord, &result.value.ival))
            result.type = TRT_DWORD;
        else
            result.type = TRT_GETLASTERROR;
        break;

    case 7:
        /* If access type not specified, use SERVICE_ALL_ACCESS */
        if (objc < 5)
            dw = SERVICE_ALL_ACCESS;
        result.type = TRT_SC_HANDLE;
        result.value.hval = OpenServiceW(h, s, dw);
        break;

    default:
        return TwapiReturnTwapiError(interp, NULL, TWAPI_INVALID_FUNCTION_CODE);
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

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func),
                     ARGSKIP,
                     GETVAR(sidP, ObjToPSID),
                     ARGTERM) != TCL_OK) {
        if (sidP)
            TwapiFree(sidP);         /* Might be alloc'ed even on fail */
        return TCL_ERROR;
    }
    /* sidP may legitimately be NULL, else it points to a Twapialloc'ed block */

    result.type = TRT_EXCEPTION_ON_FALSE; /* Likely result type */
    if (func < 1000) {
        switch (func) {
        case 1:
            result.type = TRT_BOOL;
            result.value.bval = IsValidSid(sidP);
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
                    return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
                /* FALLTHRU */
            case 1007:
                if (objc != 6 && objc != 5)
                    return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

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

/* Call Net*Enum style API */
int Twapi_CallNetEnumObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s1, s2, s3;
    DWORD   dw, dwresume;
    TwapiResult result;
    TwapiNetEnumContext netenum;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETNULLIFEMPTY(s1),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    /* WARNING:
       Many of the cases in the switch below cannot be combined even
       though they look similar because of slight variations in the
       Win32 function prototypes they call like const / non-const,
       size of resume handle etc. */

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1: // NetUseEnum system level resumehandle
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status = NetUseEnum (
            s1, netenum.level,
            &netenum.netbufP,
            twapi_netenum_bufsize,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = (DWORD_PTR) dwresume;
        netenum.tag = TWAPI_NETENUM_USEINFO;
        return TwapiReturnNetEnum(interp,&netenum);

    case 2:
        // NetUserEnum system level filter resumehandle
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(netenum.level), GETINT(dw), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status = NetUserEnum(s1, netenum.level, dw,
                                       &netenum.netbufP,
                                       twapi_netenum_bufsize,
                                       &netenum.entriesread,
                                       &netenum.totalentries,
                                       &dwresume);
        netenum.hresume = (DWORD_PTR) dwresume;
        netenum.tag = TWAPI_NETENUM_USERINFO;
        return TwapiReturnNetEnum(interp,&netenum);

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
                twapi_netenum_bufsize,
                &netenum.entriesread,
                &netenum.totalentries,
                &netenum.hresume);
        netenum.tag = (func == 3 ? TWAPI_NETENUM_GROUPINFO : TWAPI_NETENUM_LOCALGROUPINFO);
        return TwapiReturnNetEnum(interp,&netenum);

    case 5: // NetShareEnum system level resumehandle
        // Not shared with above code because first param has a const
        // qualifier in above cases which results in warnings if 
        // combined with this case.
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status = NetShareEnum(s1, netenum.level,
                                      &netenum.netbufP,
                                      twapi_netenum_bufsize,
                                      &netenum.entriesread,
                                      &netenum.totalentries,
                                      &dwresume);
        netenum.hresume = (DWORD_PTR) dwresume;
        netenum.tag = TWAPI_NETENUM_SHAREINFO;
        return TwapiReturnNetEnum(interp,&netenum);

    case 6:
        // NetUserGetGroups server user level
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.hresume = 0; /* Not used for these calls */
        netenum.status = NetUserGetGroups(
            s1, s2, netenum.level, &netenum.netbufP,
            twapi_netenum_bufsize, &netenum.entriesread, &netenum.totalentries);
        netenum.tag = TWAPI_NETENUM_GROUPUSERSINFO;
        return TwapiReturnNetEnum(interp,&netenum);

    case 7:
        // NetUserGetLocalGroups server user level flags
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.hresume = 0; /* Not used for these calls */
        netenum.status = NetUserGetLocalGroups (
            s1, s2, netenum.level, dw, &netenum.netbufP, twapi_netenum_bufsize,
            &netenum.entriesread, &netenum.totalentries);
        netenum.tag = TWAPI_NETENUM_LOCALGROUPUSERSINFO;
        return TwapiReturnNetEnum(interp, &netenum);

    case 8:
    case 9:
        // NetLocalGroupGetMembers server group level resumehandle
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level), GETDWORD_PTR(netenum.hresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status = (func == 8 ? NetLocalGroupGetMembers : NetGroupGetUsers) (
            s1, s2, netenum.level, &netenum.netbufP, twapi_netenum_bufsize,
            &netenum.entriesread, &netenum.totalentries, &netenum.hresume);
        netenum.tag = (func == 8 ?
                       TWAPI_NETENUM_LOCALGROUPMEMBERSINFO : TWAPI_NETENUM_GROUPUSERSINFO);
        return TwapiReturnNetEnum(interp,&netenum);

    case 10:  // NetConnectionEnum server group level resumehandle
        // Not shared with other code because first param has a const
        // qualifier in above cases which results in warnings if 
        // combined with this case.
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETWSTR(s2), GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        netenum.status = NetConnectionEnum (
            s1,
            Tcl_GetUnicode(objv[3]),
            netenum.level,
            &netenum.netbufP,
            twapi_netenum_bufsize,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = (DWORD_PTR)dwresume;
        netenum.tag = TWAPI_NETENUM_CONNECTIONINFO;
        return TwapiReturnNetEnum(interp,&netenum);

    case 11:
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETNULLIFEMPTY(s2), GETNULLIFEMPTY(s3),
                         GETINT(netenum.level), GETDWORD_PTR(netenum.hresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status = NetFileEnum (
            s1, s2, s3, netenum.level, 
            &netenum.netbufP,
            twapi_netenum_bufsize,
            &netenum.entriesread,
            &netenum.totalentries,
            &netenum.hresume);
        netenum.tag = TWAPI_NETENUM_FILEINFO;
        return TwapiReturnNetEnum(interp,&netenum);

    case 12:
        if (TwapiGetArgs(interp, objc-3, objv+3,
                         GETNULLIFEMPTY(s2), GETNULLIFEMPTY(s3),
                         GETINT(netenum.level), GETINT(dwresume),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        netenum.status = NetSessionEnum (
            s1, s2, s3, netenum.level, 
            &netenum.netbufP,
            twapi_netenum_bufsize,
            &netenum.entriesread,
            &netenum.totalentries,
            &dwresume);
        netenum.hresume = dwresume;
        netenum.tag = TWAPI_NETENUM_SESSIONINFO;
        return TwapiReturnNetEnum(interp,&netenum);

    }

    return TwapiSetResult(interp, &result);

}


/* Call PDH API. This is special-cased because we have to do a restore
   locale after every PDH call on some platforms */
int Twapi_CallPdhObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s, s2, s3;
    DWORD   dw, dw2;
    HANDLE h;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 100) {
        /* No arguments */
        if (objc != 2)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 1:
            dw = PdhGetDllVersion(&result.value.ival);
            if (dw == 0)
                result.type = TRT_DWORD;
            else {
                result.value.ival = dw;
                result.type = TRT_EXCEPTION_ON_ERROR;
            }
            break;
        case 2:
            return Twapi_PdhBrowseCounters(interp);
        }
    } else if (func < 200) {
        /* Single argument */
        if (objc != 3)
            return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 101:
            if (Tcl_GetLongFromObj(interp, objv[2], &dw) != TCL_OK)
                return TwapiReturnTwapiError(interp, NULL, TWAPI_INVALID_ARGS);
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhSetDefaultRealTimeDataSource(dw);
            break;
        case 102:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhConnectMachineW(ObjToLPWSTR_NULL_IF_EMPTY(objv[2]));
            break;
        case 103:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhValidatePathW(Tcl_GetUnicode(objv[2]));
            break;
        }
    } else if (func < 300) {
        /* Single string with integer arg */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 201: 
            return Twapi_PdhParseCounterPath(ticP, s, dw);
        case 202: 
            return Twapi_PdhLookupPerfNameByIndex(interp, s, dw);
        case 203:
            NULLIFY_EMPTY(s);
            dw = PdhOpenQueryW(s, dw, &result.value.hval);
            if (dw == 0)
                result.type = TRT_HANDLE;
            else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw;
            }
            break;
        }
    } else if (func < 400) {
        /* Single handle */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 301:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhRemoveCounter(h);
            break;
        case 302:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhCollectQueryData(h);
            break;
        case 303:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhCloseQuery(h);
            break;
        }
    } else {
        /* Free for all */
        switch (func) {
        case 1001:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhGetFormattedCounterValue(interp, h, dw);
        case 1002:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWSTR(s), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            dw = PdhAddCounterW(h, s, dw, &result.value.hval);
            if (dw == 0)
                result.type = TRT_HANDLE;
            else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw;
            }
            break;
        case 1003:
            return Twapi_PdhMakeCounterPath(ticP, objc-2, objv+2);
        case 1004:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETWSTR(s3), GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhEnumObjectItems(ticP, s, s2, s3, dw, dw2);
        case 1005:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETINT(dw), GETBOOL(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhEnumObjects(ticP, s, s2, dw, dw2);
        }
    }

    /* Set Tcl status before restoring locale as the latter might change
       value of GetLastError() */
    dw = TwapiSetResult(interp, &result);

    TwapiPdhRestoreLocale();

    return dw;
}
