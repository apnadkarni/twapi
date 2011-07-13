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
    WCHAR *msgP;

    /* For security reasons, MSDN recommends not to use FormatMessageW
       with arbitrary codes unless IGNORE_INSERTS is also used, in which
       case arguments are ignored. As Richter suggested in his book,
       TWAPI used to use __try to protect against this but note _try
       does not protect against malicious buffer overflows.

       There is also another problem in that we are passing all strings
       but the format specifiers in the message may expect (for example)
       integer values. We have no way of doing the right thing without
       building a FormatMessage parser ourselves. That is what we do
       at the script level and force IGNORE_INSERTS here.

       As a side-benefit, not using __try reduces CRT dependency.
    */

    dwFlags |= FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS;
    if (FormatMessageW(dwFlags, lpSource, dwMessageId, dwLanguageId, (LPWSTR) &msgP, argc, (va_list *)argv)) {
        Tcl_SetObjResult(interp, ObjFromUnicode(msgP));
        LocalFree(msgP);
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
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

int Twapi_GlobalMemoryStatus(Tcl_Interp *interp)
{
    MEMORYSTATUSEX memstatex;
    Tcl_Obj *objv[14];

    memstatex.dwLength = sizeof(memstatex);
    if (GlobalMemoryStatusEx(&memstatex)) {
        objv[0] = STRING_LITERAL_OBJ("dwMemoryLoad");
        objv[1] = Tcl_NewIntObj(memstatex.dwMemoryLoad);
        objv[2] = STRING_LITERAL_OBJ("ullTotalPhys");
        objv[3] = ObjFromULONGLONG(memstatex.ullTotalPhys);
        objv[4] = STRING_LITERAL_OBJ("ullAvailPhys");
        objv[5] = ObjFromULONGLONG(memstatex.ullAvailPhys);
        objv[6] = STRING_LITERAL_OBJ("ullTotalPageFile");
        objv[7] = ObjFromULONGLONG(memstatex.ullTotalPageFile);
        objv[8] = STRING_LITERAL_OBJ("ullAvailPageFile");
        objv[9] = ObjFromULONGLONG(memstatex.ullAvailPageFile);
        objv[10] = STRING_LITERAL_OBJ("ullTotalVirtual");
        objv[11] = ObjFromULONGLONG(memstatex.ullTotalVirtual);
        objv[12] = STRING_LITERAL_OBJ("ullAvailVirtual");
        objv[13] = ObjFromULONGLONG(memstatex.ullAvailVirtual);
        
        Tcl_SetObjResult(interp, Tcl_NewListObj(14, objv));
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
}

int Twapi_GetPerformanceInformation(Tcl_Interp *interp)
{
    PERFORMANCE_INFORMATION perf;
    Tcl_Obj *objv[26];

    perf.cb = sizeof(perf);
    if (GetPerformanceInfo(&perf, sizeof(perf))) {
        objv[0] = STRING_LITERAL_OBJ("CommitTotal");
        objv[1] = ObjFromSIZE_T(perf.CommitTotal);
        objv[2] = STRING_LITERAL_OBJ("CommitLimit");
        objv[3] = ObjFromSIZE_T(perf.CommitLimit);
        objv[4] = STRING_LITERAL_OBJ("CommitPeak");
        objv[5] = ObjFromSIZE_T(perf.CommitPeak);
        objv[6] = STRING_LITERAL_OBJ("PhysicalTotal");
        objv[7] = ObjFromSIZE_T(perf.PhysicalTotal);
        objv[8] = STRING_LITERAL_OBJ("PhysicalAvailable");
        objv[9] = ObjFromSIZE_T(perf.PhysicalAvailable);
        objv[10] = STRING_LITERAL_OBJ("SystemCache");
        objv[11] = ObjFromSIZE_T(perf.SystemCache);
        objv[12] = STRING_LITERAL_OBJ("KernelTotal");
        objv[13] = ObjFromSIZE_T(perf.KernelTotal);
        objv[14] = STRING_LITERAL_OBJ("KernelPaged");
        objv[15] = ObjFromSIZE_T(perf.KernelPaged);
        objv[16] = STRING_LITERAL_OBJ("KernelNonpaged");
        objv[17] = ObjFromSIZE_T(perf.KernelNonpaged);
        objv[18] = STRING_LITERAL_OBJ("PageSize");
        objv[19] = ObjFromSIZE_T(perf.PageSize);
        objv[20] = STRING_LITERAL_OBJ("HandleCount");
        objv[21] = Tcl_NewIntObj(perf.HandleCount);
        objv[22] = STRING_LITERAL_OBJ("ProcessCount");
        objv[23] = Tcl_NewIntObj(perf.ProcessCount);
        objv[24] = STRING_LITERAL_OBJ("ThreadCount");
        objv[25] = Tcl_NewIntObj(perf.ThreadCount);

        Tcl_SetObjResult(interp, Tcl_NewListObj(26, objv));
        return TCL_OK;
    } else
        return TwapiReturnSystemError(interp);
}


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
#ifdef BADCODE
    /* On Vista and later, the system does not like it if we supply a buffer
       larger than required. So do not update bufsz to actual allocated
       size, just use exact size as below
    */
    bufP = MemLifoPushFrame(&ticP->memlifo, bufsz, &bufsz);
#else
    bufP = MemLifoPushFrame(&ticP->memlifo, bufsz, NULL);
#endif
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

