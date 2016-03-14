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
int Twapi_SystemPagefileInformation(Tcl_Interp *interp);

static MAKE_DYNLOAD_FUNC(NtQuerySystemInformation, ntdll, NtQuerySystemInformation_t)
static MAKE_DYNLOAD_FUNC(GetProductInfo, kernel32, FARPROC)
static MAKE_DYNLOAD_FUNC(GetLogicalProcessorInformation, kernel32, FARPROC)
static MAKE_DYNLOAD_FUNC(GetLogicalProcessorInformationEx, kernel32, FARPROC)


/*
 * Different SDK's have different versions of the following structures.
 * We define our own equivalents. Currently these correspond to the 8.1 SDK
 */
typedef struct _TWAPI_PROCESSOR_NUMBER {
    WORD   Group;
    BYTE  Number;
    BYTE  Reserved;
} TWAPI_PROCESSOR_NUMBER;

//
// Structure to represent a group-specific affinity, such as that of a
// thread.  Specifies the group number and the affinity within that group.
//

typedef struct _TWAPI_GROUP_AFFINITY {
    KAFFINITY Mask;
    WORD   Group;
    WORD   Reserved[3];
} TWAPI_GROUP_AFFINITY;

typedef enum _TWAPI_LOGICAL_PROCESSOR_RELATIONSHIP {
    TwapiRelationProcessorCore,
    TwapiRelationNumaNode,
    TwapiRelationCache,
    TwapiRelationProcessorPackage,
    TwapiRelationGroup,
    TwapiRelationAll = 0xffff
} TWAPI_LOGICAL_PROCESSOR_RELATIONSHIP;

#ifndef LTP_PC_SMT
#define LTP_PC_SMT 0x1
#endif

typedef enum _TWAPI_PROCESSOR_CACHE_TYPE {
    TwapiCacheUnified,
    TwapiCacheInstruction,
    TwapiCacheData,
    TwapiCacheTrace
} TWAPI_PROCESSOR_CACHE_TYPE;

#define TWAPI_CACHE_FULLY_ASSOCIATIVE 0xFF

typedef struct _TWAPI_CACHE_DESCRIPTOR {
    BYTE   Level;
    BYTE   Associativity;
    WORD   LineSize;
    DWORD  Size;
    TWAPI_PROCESSOR_CACHE_TYPE Type;
} TWAPI_CACHE_DESCRIPTOR;

typedef struct _TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION {
    ULONG_PTR   ProcessorMask;
    TWAPI_LOGICAL_PROCESSOR_RELATIONSHIP Relationship;
    union {
        struct {
            BYTE  Flags;
        } ProcessorCore;
        struct {
            DWORD NodeNumber;
        } NumaNode;
        TWAPI_CACHE_DESCRIPTOR Cache;
        ULONGLONG  Reserved[2];
    } DUMMYUNIONNAME;
} TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION;

typedef struct _TWAPI_PROCESSOR_RELATIONSHIP {
    BYTE  Flags;
    BYTE  Reserved[21];
    WORD   GroupCount;
    TWAPI_GROUP_AFFINITY GroupMask[ANYSIZE_ARRAY];
} TWAPI_PROCESSOR_RELATIONSHIP;

typedef struct _TWAPI_NUMA_NODE_RELATIONSHIP {
    DWORD NodeNumber;
    BYTE  Reserved[20];
    TWAPI_GROUP_AFFINITY GroupMask;
} TWAPI_NUMA_NODE_RELATIONSHIP;

typedef struct _TWAPI_CACHE_RELATIONSHIP {
    BYTE  Level;
    BYTE  Associativity;
    WORD   LineSize;
    DWORD CacheSize;
    TWAPI_PROCESSOR_CACHE_TYPE Type;
    BYTE  Reserved[20];
    TWAPI_GROUP_AFFINITY GroupMask;
} TWAPI_CACHE_RELATIONSHIP;

typedef struct _TWAPI_PROCESSOR_GROUP_INFO {
    BYTE  MaximumProcessorCount;
    BYTE  ActiveProcessorCount;
    BYTE  Reserved[38];
    KAFFINITY ActiveProcessorMask;
} TWAPI_PROCESSOR_GROUP_INFO;

typedef struct _TWAPI_GROUP_RELATIONSHIP {
    WORD   MaximumGroupCount;
    WORD   ActiveGroupCount;
    BYTE  Reserved[20];
    TWAPI_PROCESSOR_GROUP_INFO GroupInfo[ANYSIZE_ARRAY];
} TWAPI_GROUP_RELATIONSHIP;

typedef struct _TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX {
    TWAPI_LOGICAL_PROCESSOR_RELATIONSHIP Relationship;
    DWORD Size;
    union {
        TWAPI_PROCESSOR_RELATIONSHIP Processor;
        TWAPI_NUMA_NODE_RELATIONSHIP NumaNode;
        TWAPI_CACHE_RELATIONSHIP Cache;
        TWAPI_GROUP_RELATIONSHIP Group;
    } DUMMYUNIONNAME;
} TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX;

static TWAPI_INLINE Tcl_Obj *ObjFromKAFFINITY(KAFFINITY kaffinity) {
    return ObjFromULONG_PTR(kaffinity);
}

static Tcl_Obj *ObjFromGROUP_AFFINITY(TWAPI_GROUP_AFFINITY *gaP)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromKAFFINITY(gaP->Mask);
    objs[1] = ObjFromLong(gaP->Group);
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromSYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX(
    TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX *slpiexP
    )
{
    Tcl_Obj *objs[2];
    Tcl_Obj *elems[6];
    Tcl_Obj *elems2[3];
    int i, nelems;

    objs[0] = ObjFromLong(slpiexP->Relationship);
    switch (slpiexP->Relationship) {
    case TwapiRelationCache:
        elems[0] = ObjFromLong(slpiexP->Cache.Level);
        elems[1] = ObjFromLong(slpiexP->Cache.Associativity);
        elems[2] = ObjFromLong(slpiexP->Cache.LineSize);
        elems[3] = ObjFromDWORD(slpiexP->Cache.CacheSize);
        elems[4] = ObjFromLong(slpiexP->Cache.Type);
        elems[5] = ObjFromGROUP_AFFINITY(&slpiexP->Cache.GroupMask);
        nelems = 6;
        break;
    case TwapiRelationGroup:
        elems[0] = ObjFromLong(slpiexP->Group.MaximumGroupCount);
        elems[1] = ObjNewList(slpiexP->Group.ActiveGroupCount, NULL);
        for (i = 0; i < slpiexP->Group.ActiveGroupCount; ++i) {
            elems2[0] = ObjFromLong(slpiexP->Group.GroupInfo[i].MaximumProcessorCount);
            elems2[1] = ObjFromLong(slpiexP->Group.GroupInfo[i].ActiveProcessorCount);      
            elems2[2] = ObjFromKAFFINITY(slpiexP->Group.GroupInfo[i].ActiveProcessorMask);
            ObjAppendElement(NULL, elems[1], ObjNewList(3, elems2));
        }
        nelems = 2;
        break;
    case TwapiRelationNumaNode:
        elems[0] = ObjFromLong(slpiexP->NumaNode.NodeNumber);
        elems[1] = ObjFromGROUP_AFFINITY(&slpiexP->NumaNode.GroupMask);
        nelems = 2;
        break;
    case TwapiRelationProcessorCore:
    case TwapiRelationProcessorPackage:
        elems[0] = ObjFromLong(slpiexP->Processor.Flags);
        elems[1] = ObjNewList(slpiexP->Processor.GroupCount, NULL);
        for (i = 0; i < slpiexP->Processor.GroupCount; ++i)
            ObjAppendElement(NULL, elems[1], ObjFromGROUP_AFFINITY(&slpiexP->Processor.GroupMask[i]));
        nelems = 2;
        break;
    }

    objs[1] = ObjNewList(nelems, elems);
    return ObjNewList(2, objs);

}

static Tcl_Obj *ObjFromSYSTEM_LOGICAL_PROCESSOR_INFORMATION(
    TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION *slpiP
    )
{
    Tcl_Obj *objs[3];
    Tcl_Obj *elems[5];

    objs[0] = ObjFromULONG_PTR(slpiP->ProcessorMask);
    objs[1] = ObjFromLong(slpiP->Relationship);
    
    switch (slpiP->Relationship) {
    case TwapiRelationCache:
        elems[0] = ObjFromLong(slpiP->Cache.Level);
        elems[1] = ObjFromLong(slpiP->Cache.Associativity);
        elems[2] = ObjFromLong(slpiP->Cache.LineSize);
        elems[3] = ObjFromDWORD(slpiP->Cache.Size);
        elems[4] = ObjFromLong(slpiP->Cache.Type);
        objs[2] = ObjNewList(5, elems);
        break;
    case TwapiRelationNumaNode:
        objs[2] = ObjFromLong(slpiP->NumaNode.NodeNumber);
        break;
    case TwapiRelationProcessorCore:
        objs[2] = ObjFromLong(slpiP->ProcessorCore.Flags);
        break;
    case TwapiRelationProcessorPackage:
        objs[2] = ObjFromEmptyString();
        break;
    }

    return ObjNewList(3, objs);
}


TCL_RESULT Twapi_GetLogicalProcessorInformationEx(
    Tcl_Interp *interp,
    TWAPI_LOGICAL_PROCESSOR_RELATIONSHIP rel
    )
{
    FARPROC fn;
    DWORD sz = 0, winerr;
    TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX *slpiexP;

    fn = Twapi_GetProc_GetLogicalProcessorInformationEx();
    if (fn == NULL)
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);

    slpiexP = SWSPushFrame(1000, &sz);
    winerr = ERROR_SUCCESS;
    if (fn(rel, slpiexP, &sz) == 0) {
        winerr = GetLastError();
        if (winerr == ERROR_INSUFFICIENT_BUFFER) {
            winerr = ERROR_SUCCESS;
            slpiexP = SWSAlloc(sz, NULL);
            if (fn(rel, slpiexP, &sz) == 0)
                winerr = GetLastError();
        }
    }
    if (winerr == ERROR_SUCCESS) {
        Tcl_Obj *objP = ObjNewList(0, NULL);
        DWORD used;
        /* NOTE - TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX
           is actually a variable size structure even if
           the C struct is not defined as such. So we cannot
           just advance the pointer. We have to add the explicit
           size of the current structure */
        used = 0;
        while (used < sz) {
            ObjAppendElement(NULL, objP, ObjFromSYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX(slpiexP));
            used += slpiexP->Size;
            slpiexP = ADDPTR(slpiexP, slpiexP->Size, TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX*);
        }
        ObjSetResult(interp, objP);
    }
    SWSPopFrame();
    return winerr == ERROR_SUCCESS ? TCL_OK : Twapi_AppendSystemError(interp, winerr);
}

TCL_RESULT Twapi_GetLogicalProcessorInformation(Tcl_Interp *interp)
{
    FARPROC fn;
    DWORD sz = 0, winerr;
    TWAPI_SYSTEM_LOGICAL_PROCESSOR_INFORMATION *slpiP;

    fn = Twapi_GetProc_GetLogicalProcessorInformation();
    if (fn == NULL)
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);

    slpiP = SWSPushFrame(1000, &sz);
    winerr = ERROR_SUCCESS;
    if (fn(slpiP, &sz) == 0) {
        winerr = GetLastError();
        if (winerr == ERROR_INSUFFICIENT_BUFFER) {
            winerr = ERROR_SUCCESS;
            slpiP = SWSAlloc(sz, NULL);
            if (fn(slpiP, &sz) == 0)
                winerr = GetLastError();
        }
    }
    if (winerr == ERROR_SUCCESS) {
        Tcl_Obj *objP = ObjNewList(0, NULL);
        int used;
        used = 0;
        while ((used + sizeof(*slpiP)) < sz) {
            ObjAppendElement(NULL, objP, ObjFromSYSTEM_LOGICAL_PROCESSOR_INFORMATION(slpiP));
            used += sizeof(*slpiP);
            ++slpiP;
        }
        ObjSetResult(interp, objP);
    }
    SWSPopFrame();
    return winerr == ERROR_SUCCESS ? TCL_OK : Twapi_AppendSystemError(interp, winerr);
}

int Twapi_GetSystemInfo(Tcl_Interp *interp)
{
    SYSTEM_INFO sysinfo;
    Tcl_Obj *objv[10];

    GetSystemInfo(&sysinfo);
    objv[0] = ObjFromInt((unsigned) sysinfo.wProcessorArchitecture);
    objv[1] = ObjFromDWORD(sysinfo.dwPageSize);
    objv[2] = ObjFromDWORD_PTR((DWORD_PTR) sysinfo.lpMinimumApplicationAddress);
    objv[3] = ObjFromDWORD_PTR((DWORD_PTR) sysinfo.lpMaximumApplicationAddress);
    objv[4] = ObjFromDWORD_PTR(sysinfo.dwActiveProcessorMask);
    objv[5] = ObjFromLong(sysinfo.dwNumberOfProcessors);
    objv[6] = ObjFromLong(sysinfo.dwProcessorType);
    objv[7] = ObjFromLong(sysinfo.dwAllocationGranularity);
    objv[8] = ObjFromInt((unsigned)sysinfo.wProcessorLevel);
    objv[9] = ObjFromInt((unsigned)sysinfo.wProcessorRevision);

    ObjSetResult(interp, ObjNewList(10, objv));
    return TCL_OK;
}

int Twapi_GlobalMemoryStatus(Tcl_Interp *interp)
{
    MEMORYSTATUSEX memstatex;
    Tcl_Obj *objv[14];

    memstatex.dwLength = sizeof(memstatex);
    if (GlobalMemoryStatusEx(&memstatex)) {
        objv[0] = STRING_LITERAL_OBJ("dwMemoryLoad");
        objv[1] = ObjFromDWORD(memstatex.dwMemoryLoad);
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
        
        ObjSetResult(interp, ObjNewList(14, objv));
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
        objv[21] = ObjFromDWORD(perf.HandleCount);
        objv[22] = STRING_LITERAL_OBJ("ProcessCount");
        objv[23] = ObjFromDWORD(perf.ProcessCount);
        objv[24] = STRING_LITERAL_OBJ("ThreadCount");
        objv[25] = ObjFromDWORD(perf.ThreadCount);

        ObjSetResult(interp, ObjNewList(26, objv));
        return TCL_OK;
    } else
        return TwapiReturnSystemError(interp);
}

static TCL_RESULT Twapi_SystemProcessorTimesObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SYSTEM_INFO sysinfo;
    void  *bufP;
    int    bufsz;
    ULONG  dummy;
    NTSTATUS status;
    DWORD    i;
    NtQuerySystemInformation_t NtQuerySystemInformationPtr = Twapi_GetProc_NtQuerySystemInformation();
    Tcl_Obj *resultObj;

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
    bufP = MemLifoPushFrame(ticP->memlifoP, bufsz, &bufsz);
#else
    bufP = MemLifoPushFrame(ticP->memlifoP, bufsz, NULL);
#endif
    status = (*NtQuerySystemInformationPtr)(8, bufP, bufsz, &dummy);

    if (status == 0) {
        resultObj = ObjEmptyList();
        for (i = 0; i < sysinfo.dwNumberOfProcessors; ++i) {
            Tcl_Obj *obj = ObjEmptyList();
            SYSTEM_PROCESSOR_TIMES *timesP = i+((SYSTEM_PROCESSOR_TIMES *)bufP);

            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, IdleTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, KernelTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, UserTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, DpcTime);
            Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, obj, timesP, InterruptTime);
            Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, obj, timesP, InterruptCount);
            ObjAppendElement(interp, resultObj, obj);
        }

        ObjSetResult(interp, resultObj);
    }

    MemLifoPopFrame(ticP->memlifoP);
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
    ObjAppendElement(NULL, objP, ObjNewList(4, objs));

    return TRUE;
}


/*
 *  Wrapper around NtQuerySystemInformation to get swapfile information
 *
 */
int Twapi_SystemPagefileInformation(Tcl_Interp *interp)
{
    Tcl_Obj *resultObj = ObjEmptyList();

    /* MS HEADER BUG WORKAROUND: see http://www.ureader.com/msg/147433.aspx
       Despite what the SDK headers say, the callback function should
       follow WINAPI calling convention. So we have to implement the
       callback as WINAPI otherwise the stack gets screwed up. Then
       we need to CAST it so it matches the header.
    */

    if (EnumPageFilesW((PENUM_PAGE_FILE_CALLBACKW) TwapiEnumPageFilesProc, resultObj)) {
        ObjSetResult(interp, resultObj);
        return TCL_OK;
    } else {
        DWORD winerr = GetLastError();
        ObjDecrRefs(resultObj);
        return Twapi_AppendSystemError(interp, winerr);
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

    ObjSetResult(interp, ObjFromUnicodeN(path, len));
    return TCL_OK;
}

static int Twapi_OsCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD dw, dw2, dw3;
    TwapiResult result;
    union {
        WCHAR buf[MAX_PATH+1];
        TIME_ZONE_INFORMATION tzinfo;
    } u;
    Tcl_Obj *objs[2];
    SYSTEMTIME systime;
    TIME_ZONE_INFORMATION *tzinfoP;
    int func = PtrToInt(clientdata);
    Tcl_Obj *sObj, *s2Obj;
    FARPROC fn;

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 100) {
        if (objc != 0)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 32:
            return Twapi_GetLogicalProcessorInformation(interp);
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
            return Twapi_GetPerformanceInformation(interp);
        case 37:
            return Twapi_SystemPagefileInformation(interp);
        case 38:
            return Twapi_GetSystemWow64Directory(interp);
        case 39:
            dw = GetTimeZoneInformation(&u.tzinfo);
            switch (dw) {
            case TIME_ZONE_ID_UNKNOWN:
            case TIME_ZONE_ID_STANDARD:
            case TIME_ZONE_ID_DAYLIGHT:
                objs[0] = ObjFromLong(dw);
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
    } else if (func < 300) {
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
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
            result.type = TRT_LONG;
            result.value.ival = GetSystemMetrics(dw);
            break;
        case 203:
            result.type = TRT_EMPTY;
            Sleep(dw);
            break;
        case 204:
            return Twapi_GetLogicalProcessorInformationEx(interp, dw);
        }
    } else {
        switch (func) {
        case 1001:
            if (TwapiGetArgs(interp, objc, objv,
                             GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ExitWindowsEx(dw, dw2);
            break;
        case 1002:
            if (TwapiGetArgs(interp, objc, objv,
                             GETOBJ(sObj), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = AbortSystemShutdownW(ObjToLPWSTR_NULL_IF_EMPTY(sObj));
        case 1003:
            if (TwapiGetArgs(interp, objc, objv,
                             GETOBJ(sObj), GETOBJ(s2Obj),
                             GETINT(dw), GETBOOL(dw2), GETBOOL(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = InitiateSystemShutdownW(ObjToLPWSTR_NULL_IF_EMPTY(sObj), ObjToLPWSTR_NULL_IF_EMPTY(s2Obj), dw, dw2, dw3);
            break;
        case 1004:
            if (TwapiGetArgs(interp, objc, objv,
                             GETBOOL(dw), GETBOOL(dw2), GETBOOL(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetSuspendState((BOOLEAN) dw, (BOOLEAN) dw2, (BOOLEAN) dw3);
            break;
        case 1005: // TzLocalSpecificTimeToSystemTime
        case 1006: // SystemTimeToTzSpecificLocalTime
            if (objc == 1) {
                tzinfoP = NULL;
            } else if (objc == 2) {
                if (ObjToTIME_ZONE_INFORMATION(interp, objv[1], &u.tzinfo) != TCL_OK) {
                    return TCL_ERROR;
                }
                tzinfoP = &u.tzinfo;
            } else {
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            }
            if (ObjToSYSTEMTIME(interp, objv[0], &systime) != TCL_OK)
                return TCL_ERROR;
            if ((func == 10134 ? TzSpecificLocalTimeToSystemTime : SystemTimeToTzSpecificLocalTime) (tzinfoP, &systime, &result.value.systime))
                result.type = TRT_SYSTEMTIME;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1007: // GetProductInfo
            fn = Twapi_GetProc_GetProductInfo();
            if (fn == NULL) {
                return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
            }
            result.value.ival = 0;
            result.type = TRT_DWORD;
            (*fn)(6,0,0,0, &result.value.ival);
            break;
        }
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiOsInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s OsDispatch[] = {
        DEFINE_FNCODE_CMD(GetLogicalProcessorInformation, 32),
        DEFINE_FNCODE_CMD(GetComputerName, 33),
        DEFINE_FNCODE_CMD(GetSystemInfo, 34),
        DEFINE_FNCODE_CMD(GlobalMemoryStatus, 35),
        DEFINE_FNCODE_CMD(GetPerformanceInformation, 36),
        DEFINE_FNCODE_CMD(Twapi_SystemPagefileInformation, 37),
        DEFINE_FNCODE_CMD(GetSystemWow64Directory, 38),
        DEFINE_FNCODE_CMD(GetTimeZoneInformation, 39),    /* TBD Tcl */
        DEFINE_FNCODE_CMD(GetComputerNameEx, 201),
        DEFINE_FNCODE_CMD(GetSystemMetrics, 202),
        DEFINE_FNCODE_CMD(Sleep, 203),
        DEFINE_FNCODE_CMD(GetLogicalProcessorInformationEx, 204),
        DEFINE_FNCODE_CMD(ExitWindowsEx, 1001),
        DEFINE_FNCODE_CMD(AbortSystemShutdown, 1002),
        DEFINE_FNCODE_CMD(InitiateSystemShutdown, 1003),
        DEFINE_FNCODE_CMD(SetSuspendState, 1004),
        DEFINE_FNCODE_CMD(TzSpecificLocalTimeToSystemTime, 1005), // Tcl TBD
        DEFINE_FNCODE_CMD(SystemTimeToTzSpecificLocalTime, 1006), // Tcl TBD
        DEFINE_FNCODE_CMD(GetProductInfo, 1007),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(OsDispatch), OsDispatch, Twapi_OsCallObjCmd);
    Tcl_CreateObjCommand(interp, "twapi::Twapi_SystemProcessorTimes", Twapi_SystemProcessorTimesObjCmd, ticP, NULL);

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
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiOsInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

