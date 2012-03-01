/*
 * Copyright (c) 2010-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

int Twapi_GetSystemWow64Directory(Tcl_Interp *interp);
int Twapi_GetSystemInfo(Tcl_Interp *interp);
TCL_RESULT Twapi_GlobalMemoryStatus(Tcl_Interp *interp);
TCL_RESULT Twapi_GetPerformanceInformation(Tcl_Interp *interp);
int Twapi_SystemProcessorTimes(TwapiInterpContext *ticP);
int Twapi_SystemPagefileInformation(TwapiInterpContext *ticP);
int Twapi_GetVersionEx(Tcl_Interp *interp);

static MAKE_DYNLOAD_FUNC(NtQuerySystemInformation, ntdll, NtQuerySystemInformation_t)

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


/* NOTE - DESPITE WHAT THE SDK DOCS AND HEADERS SAY, THE CALLING CONVENTION
   FOR THIS IS WINAPI. See MS ack in http://www.ureader.com/msg/147433.aspx
*/
static BOOL WINAPI TwapiEnumPageFilesProc(
    Tcl_Obj *objP,
    PENUM_PAGE_FILE_INFORMATION pfiP,
    LPCWSTR fnP
    )
{
    Tcl_Obj *objs[4];

    objs[0] = ObjFromSIZE_T(pfiP->TotalSize);
    objs[1] = ObjFromSIZE_T(pfiP->TotalInUse);
    objs[2] = ObjFromSIZE_T(pfiP->PeakUsage);
    objs[3] = ObjFromUnicode(fnP);
    Tcl_ListObjAppendElement(NULL, objP, Tcl_NewListObj(4, objs));

    return TRUE;
}


/*
 *  Wrapper around NtQuerySystemInformation to get swapfile information
 *
 */
int Twapi_SystemPagefileInformation(TwapiInterpContext *ticP)
{
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    /* MS HEADER BUG WORKAROUND: see http://www.ureader.com/msg/147433.aspx
       Despite what the SDK headers say, the callback function should
       follow WINAPI calling convention. So we have to implement the
       callback as WINAPI otherwise the stack gets screwed up. Then
       we need to CAST it so it matches the header.
    */

    if (EnumPageFilesW((PENUM_PAGE_FILE_CALLBACKW) TwapiEnumPageFilesProc, resultObj)) {
        Tcl_SetObjResult(ticP->interp, resultObj);
        return TCL_OK;
    } else {
        DWORD winerr = GetLastError();
        Tcl_DecrRefCount(resultObj);
        return Twapi_AppendSystemError(ticP->interp, winerr);
    }
}


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

    Tcl_SetObjResult(interp, ObjFromUnicodeN(path, len));
    return TCL_OK;
}

static int Twapi_OsCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s, s2;
    DWORD dw, dw2, dw3;
    TwapiResult result;
    union {
        WCHAR buf[MAX_PATH+1];
        TIME_ZONE_INFORMATION tzinfo;
    } u;
    Tcl_Obj *objs[2];
    SYSTEMTIME systime;
    TIME_ZONE_INFORMATION *tzinfoP;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 100) {
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 33:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetComputerNameW(u.buf, &result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 34:
            return Twapi_GetSystemInfo(interp);
        case 35:
            return Twapi_GlobalMemoryStatus(interp);
        case 36:
            return Twapi_SystemProcessorTimes(ticP);
        case 37:
            return Twapi_SystemPagefileInformation(ticP);
        case 38:
            return Twapi_GetSystemWow64Directory(interp);
        case 39:
            result.type = TRT_DWORD;
            result.value.ival = GetTickCount();
            break;
        case 40:
            return Twapi_GetPerformanceInformation(ticP->interp);
        case 41:                /* GetSystemWindowsDirectory */
        case 42:                /* GetWindowsDirectory */
        case 43:                /* GetSystemDirectory */
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
        case 44:
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
    } else if (func < 200) {
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        switch (func) {
        case 201:
            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (GetComputerNameExW(dw, u.buf, &result.value.unicode.len)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 202:
            result.type = TRT_DWORD;
            result.value.ival = GetSystemMetrics(dw);
            break;
        case 203:
            result.type = TRT_EMPTY;
            Sleep(dw);
            break;
        }
    } else {
        switch (func) {
        case 1001:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ExitWindowsEx(dw, dw2);
            break;
        case 1002:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = AbortSystemShutdownW(s);
        case 1003:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETINT(dw), GETBOOL(dw2), GETBOOL(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = InitiateSystemShutdownW(s, s2, dw, dw2, dw3);
            break;
        case 1004:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETBOOL(dw), GETBOOL(dw2), GETBOOL(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetSuspendState((BOOLEAN) dw, (BOOLEAN) dw2, (BOOLEAN) dw3);
            break;
        case 1005: // TzLocalSpecificTimeToSystemTime
        case 1006: // SystemTimeToTzSpecificLocalTime
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
        }
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_OsInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::OsCall", Twapi_OsCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::OsCall", # code_); \
    } while (0);

    CALL_(GetComputerName, 33);
    CALL_(GetSystemInfo, 34);
    CALL_(GlobalMemoryStatus, 35);
    CALL_(Twapi_SystemProcessorTimes, 36);
    CALL_(Twapi_SystemPagefileInformation, 37);
    CALL_(GetSystemWow64Directory, 38);
    CALL_(GetTickCount, 39);
    CALL_(GetPerformanceInformation, 40);
    CALL_(GetSystemWindowsDirectory, 41); /* TBD Tcl */
    CALL_(GetWindowsDirectory, 42);       /* TBD Tcl */
    CALL_(GetSystemDirectory, 43);        /* TBD Tcl */
    CALL_(GetTimeZoneInformation, 44);    /* TBD Tcl */
    CALL_(GetComputerNameEx, 201);
    CALL_(GetSystemMetrics, 202);
    CALL_(Sleep, 203);
    CALL_(ExitWindowsEx, 1001);
    CALL_(AbortSystemShutdown, 1002);
    CALL_(InitiateSystemShutdown, 1003);
    CALL_(SetSuspendState, 1004);
    CALL_(TzSpecificLocalTimeToSystemTime, 1005); // Tcl
    CALL_(SystemTimeToTzSpecificLocalTime, 1006); // Tcl

#undef CALL_

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
int Twapi_os_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_OsInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

