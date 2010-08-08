/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>



typedef NTSTATUS (WINAPI *NtQuerySystemInformation_t)(int, PVOID, ULONG, PULONG);
MAKE_DYNLOAD_FUNC(NtQuerySystemInformation, ntdll, NtQuerySystemInformation_t)

int TwapiFormatMessageHelper(
    Tcl_Interp *interp,
    DWORD   dwFlags,
    LPCVOID lpSource,
    DWORD   dwMessageId,
    DWORD   dwLanguageId,
    int     argc,
    LPCWSTR *argv
)
{
    int result;
    WCHAR *msgP;

    result = TCL_ERROR;
    dwFlags |= FORMAT_MESSAGE_ALLOCATE_BUFFER;
#if 1
    /* For security reasons, MSDN recommends not to use FormatMessageW
       with arbitrary codes unless IGNORE_INSERTS is also used, in which
       case arguments are ignored and we do not need the __try. Note __try
       does not protect against malicious buffer overflows.

       There is also another problem in that we are passing all strings
       but the format specifiers in the message may expect (for example)
       integer values. We have not way of doing the right thing without
       building a FormatMessage parser ourselves.

       As a side-benefit, not using __try reduces CRT dependency.

       Despite all the above, we still allow inserts with the expectation
       that the application will not call this function without proper
       care. In other words, we are no better, no worse than FormatMessage
       itself. Disabling inserts disables too much functionality, eg.
       parsing event log messages.
    */
    __try {
        if (FormatMessageW(dwFlags, lpSource, dwMessageId, dwLanguageId, (LPWSTR) &msgP, argc, (va_list *)argv)) {
            Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(msgP, -1));
            LocalFree(msgP);
            result = TCL_OK;
        } else {
            TwapiReturnSystemError(interp);
        }
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        DWORD code;
        char  buf[80];
        switch (code = GetExceptionCode()) {
        case EXCEPTION_ACCESS_VIOLATION:
            Tcl_SetErrno(EFAULT);
            Tcl_PosixError(interp);
            Tcl_SetResult(interp, "Access violation in FormatMessage. Most likely, number of supplied arguments do not match those in format string", TCL_STATIC);
            break;
        default:
            StringCbPrintfA(buf, sizeof(buf),
                            "Exception %x raised by FormatMessage", code);
            Tcl_SetResult(interp, buf, TCL_VOLATILE);
            break;
        }
    }
#else
    dwFlags |= FORMAT_MESSAGE_IGNORE_INSERTS;
    if (FormatMessageW(dwFlags, lpSource, dwMessageId, dwLanguageId,
                       (LPWSTR) &msgP,
                       argc,
                       (va_list *)argv /* Actually ignored (IGNORE_INSERTS) */
            )) {
        Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(msgP, -1));
        LocalFree(msgP);
        result = TCL_OK;
    } else {
        TwapiReturnSystemError(interp);
    }
#endif /* OBSOLETE */

    return result;
}


#ifndef TWAPI_LEAN
int Twapi_GetSystemInfo(Tcl_Interp *interp)
{
    SYSTEM_INFO sysinfo;
    Tcl_Obj *objv[10];

    GetSystemInfo(&sysinfo);
    objv[0] = Tcl_NewIntObj((unsigned) sysinfo.wProcessorArchitecture);
    objv[1] = Tcl_NewLongObj(sysinfo.dwPageSize);
    objv[2] = ObjFromDWORD_PTR((DWORD_PTR) sysinfo.lpMinimumApplicationAddress);
    objv[3] = ObjFromDWORD_PTR((DWORD_PTR) sysinfo.lpMaximumApplicationAddress);
    objv[4] = ObjFromDWORD_PTR(sysinfo.dwActiveProcessorMask);
    objv[5] = Tcl_NewLongObj(sysinfo.dwNumberOfProcessors);
    objv[6] = Tcl_NewLongObj(sysinfo.dwProcessorType);
    objv[7] = Tcl_NewLongObj(sysinfo.dwAllocationGranularity);
    objv[8] = Tcl_NewIntObj((unsigned)sysinfo.wProcessorLevel);
    objv[9] = Tcl_NewIntObj((unsigned)sysinfo.wProcessorRevision);

    Tcl_SetObjResult(interp, Tcl_NewListObj(10, objv));
    return TCL_OK;
}
#endif // TWAPI_LEAN

int Twapi_GetVersionEx(Tcl_Interp *interp)
{
    OSVERSIONINFOEXW vi;
    Tcl_Obj *objP;

    vi.dwOSVersionInfoSize = sizeof(vi);
    if (GetVersionExW((OSVERSIONINFOW *)&vi) == 0) {
        return TwapiReturnSystemError(interp);
    }

    objP = Tcl_NewListObj(0, NULL);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwOSVersionInfoSize);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwMajorVersion);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwMinorVersion);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwBuildNumber);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwPlatformId);
    //TCHAR szCSDVersion[128];
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi,  wServicePackMajor);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi,  wServicePackMinor);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi,  wSuiteMask);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi,  wProductType);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi,  wReserved);

    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}

#ifndef TWAPI_LEAN
int Twapi_GlobalMemoryStatus(Tcl_Interp *interp)
{
    MEMORYSTATUSEX memstatex;
    MEMORYSTATUS memstat;
    Tcl_Obj *objv[4];

    memstatex.dwLength = sizeof(memstatex);
    if (GlobalMemoryStatusEx(&memstatex)) {
        objv[0] = Tcl_NewWideIntObj(memstatex.ullTotalPhys);
        objv[1] = Tcl_NewWideIntObj(memstatex.ullAvailPhys);
        objv[2] = Tcl_NewWideIntObj(memstatex.ullTotalPageFile);
        objv[3] = Tcl_NewWideIntObj(memstatex.ullAvailPageFile);
        Tcl_SetObjResult(interp, Tcl_NewListObj(4, objv));
        return TCL_OK;
    }

    memstat.dwLength = sizeof(memstat);
    GlobalMemoryStatus(&memstat);
    /* We create wide ints so Tcl does not treat 2GB+ as negative */
    objv[0] = Tcl_NewWideIntObj(memstat.dwTotalPhys);
    objv[1] = Tcl_NewWideIntObj(memstat.dwAvailPhys);
    objv[2] = Tcl_NewWideIntObj(memstat.dwTotalPageFile);
    objv[3] = Tcl_NewWideIntObj(memstat.dwAvailPageFile);
    Tcl_SetObjResult(interp, Tcl_NewListObj(4, objv));
    return TCL_OK;
}
#endif // TWAPI_LEAN


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
        if (lpAppName) {
            if (lpFileName)
                numchars = GetPrivateProfileSectionW(lpAppName,
                                                     bufP, bufsz,
                                                     lpFileName);
            else
                numchars = GetProfileSectionW(lpAppName, bufP, bufsz);
        } else {
            /* Get section names. Note lpFileName can be NULL */
            numchars = GetPrivateProfileSectionNamesW(bufP, bufsz, lpFileName);
        }

        if (numchars >= (bufsz-2)) {
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

#ifndef TWAPI_LEAN
int Twapi_SystemProcessorTimes(TwapiInterpContext *ticP)
{
    SYSTEM_INFO sysinfo;
    void  *bufP;
    int    bufsz;
    ULONG  dummy;
    NTSTATUS status;
    DWORD    i;
    NtQuerySystemInformation_t NtQuerySystemInformationPtr = Twapi_GetProc_NtQuerySystemInformation();
    Tcl_Obj *resultObj;
    Tcl_Interp *interp = ticP->interp;

    if (NtQuerySystemInformationPtr == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    /* Need to know number of elements returned */
    GetSystemInfo(&sysinfo);

    bufsz = sizeof(SYSTEM_PROCESSOR_TIMES) * sysinfo.dwNumberOfProcessors;
    bufP = MemLifoPushFrame(&ticP->memlifo, bufsz, &bufsz);
    status = (*NtQuerySystemInformationPtr)(8, bufP, bufsz, &dummy);

    if (status == 0) {
        resultObj = Tcl_NewListObj(0, NULL);
        for (i = 0; i < sysinfo.dwNumberOfProcessors; ++i) {
            Tcl_Obj *obj = Tcl_NewListObj(0, NULL);
            SYSTEM_PROCESSOR_TIMES *timesP = i+((SYSTEM_PROCESSOR_TIMES *)bufP);

            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, IdleTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, KernelTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, UserTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, DpcTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, InterruptTime);
            Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, timesP, InterruptCount);
            Tcl_ListObjAppendElement(interp, resultObj, obj);
        }

        Tcl_SetObjResult(interp, resultObj);
    }

    MemLifoPopFrame(&ticP->memlifo);
    return status ?
        Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status))
        : TCL_OK;
}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
/*
 *  Wrapper around NtQuerySystemInformation to get swapfile information
 *
 */
int Twapi_SystemPagefileInformation(TwapiInterpContext *ticP)
{
    struct _SYSTEM_PAGEFILE_INFORMATION *pagefileP;
    void  *bufP;
    ULONG  bufsz;          /* Number of bytes allocated */
    ULONG  dummy;
    NTSTATUS status;
    NtQuerySystemInformation_t NtQuerySystemInformationPtr = Twapi_GetProc_NtQuerySystemInformation();
    Tcl_Obj *resultObj;
    Tcl_Interp *interp = ticP->interp;

    if (NtQuerySystemInformationPtr == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    bufsz = 2000;
    bufP = MemLifoPushFrame(&ticP->memlifo, 2000, &bufsz);
    do {
        /*
         * Note for information class 18, the last parameter which
         * corresponds to number of bytes needed is not actually filled
         * in by the system so we ignore it and just double alloc size
         */
        status = (*NtQuerySystemInformationPtr)(18, bufP, bufsz, &dummy);
        if (status != STATUS_INFO_LENGTH_MISMATCH || bufsz >= 32000)
            break;
        bufsz = 2* bufsz;       /* For next iteration if needed */
        MemLifoPopFrame(&ticP->memlifo);
        bufP = MemLifoPushFrame(&ticP->memlifo, bufsz, NULL);
    } while (1);

    if (status) {
        MemLifoPopFrame(&ticP->memlifo);
        return Twapi_AppendSystemError(interp,
                                       TwapiNTSTATUSToError(status));
    }

    /* OK, now we got the info. Loop through to extract information
     * from the list. See Nebett's Window NT/2000 Native API
     * Reference for details
     */
    resultObj = Tcl_NewListObj(0, NULL);
    pagefileP = bufP;
    while (1) {
        Tcl_Obj *pagefileObj;

        pagefileObj = Tcl_NewListObj(0, NULL);
        Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, pagefileObj, pagefileP, CurrentSize);
        Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, pagefileObj, pagefileP, TotalUsed);
        Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, pagefileObj, pagefileP, PeakUsed);
        Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp, pagefileObj, pagefileP, FileName);

        Tcl_ListObjAppendElement(interp, resultObj, pagefileObj);

        /* Point to the next entry */
        if (pagefileP->NextEntryDelta == 0)
            break;              /* This was the last one */
        pagefileP = (struct _SYSTEM_PAGEFILE_INFORMATION *) (pagefileP->NextEntryDelta + (char *) pagefileP);
    }

    Tcl_SetObjResult(interp, resultObj);

    MemLifoPopFrame(&ticP->memlifo);

    return TCL_OK;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int Twapi_LoadUserProfile(
    Tcl_Interp *interp,
    HANDLE  hToken,
    DWORD                 flags,
    LPWSTR username,
    LPWSTR profilepath
    )
{
    PROFILEINFOW profileinfo;

    ZeroMemory(&profileinfo, sizeof(profileinfo));
    profileinfo.dwSize        = sizeof(profileinfo);
    profileinfo.lpUserName    = username;
    profileinfo.lpProfilePath = profilepath;

    if (LoadUserProfileW(hToken, &profileinfo) == 0) {
        return TwapiReturnSystemError(interp);
    }

    Tcl_SetObjResult(interp, ObjFromHANDLE(profileinfo.hProfile));
    return TCL_OK;
}
#endif // TWAPI_LEAN

typedef UINT (WINAPI *GetSystemWow64DirectoryW_t)(LPWSTR, UINT);
MAKE_DYNLOAD_FUNC(GetSystemWow64DirectoryW, kernel32, GetSystemWow64DirectoryW_t)
int Twapi_GetSystemWow64Directory(Tcl_Interp *interp)
{
    GetSystemWow64DirectoryW_t GetSystemWow64DirectoryPtr = Twapi_GetProc_GetSystemWow64DirectoryW();
    WCHAR path[MAX_PATH+1];
    UINT len;
    if (GetSystemWow64DirectoryPtr == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    len = (UINT) (*GetSystemWow64DirectoryPtr)(path, sizeof(path)/sizeof(path[0]));
    if (len == 0) {
        return TwapiReturnSystemError(interp);
    }
    if (len >= (sizeof(path)/sizeof(path[0]))) {
        return Twapi_AppendSystemError(interp, ERROR_INSUFFICIENT_BUFFER);
    }

    Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(path, len));
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
int Twapi_TclGetChannelHandle(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    char *chan_name;
    int mode, direction;
    ClientData h;
    Tcl_Channel chan;

    if (TwapiGetArgs(interp, objc, objv,
                     GETASTR(chan_name), GETINT(direction),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    chan = Tcl_GetChannel(interp, chan_name, &mode);
    if (chan == NULL) {
        Tcl_SetResult(interp, "Unknown channel", TCL_STATIC);
        return TCL_ERROR;
    }
    
    direction = direction ? TCL_WRITABLE : TCL_READABLE;
    
    if (Tcl_GetChannelHandle(chan, direction, &h) == TCL_ERROR) {
        Tcl_SetResult(interp, "Error getting channel handle", TCL_STATIC);
        return TCL_ERROR;
    }

    Tcl_SetObjResult(interp, ObjFromHANDLE(h));
    return TCL_OK;
}

