/* 
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to process information */

#include "twapi.h"

#ifndef TWAPI_STATIC_BUILD
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

typedef NTSTATUS (WINAPI *NtQueryInformationProcess_t)(HANDLE, int, PVOID, ULONG, PULONG);
static MAKE_DYNLOAD_FUNC(NtQueryInformationProcess, ntdll, NtQueryInformationProcess_t)
typedef NTSTATUS (WINAPI *NtQueryInformationThread_t)(HANDLE, int, PVOID, ULONG, PULONG);
static MAKE_DYNLOAD_FUNC(NtQueryInformationThread, ntdll, NtQueryInformationThread_t)
static MAKE_DYNLOAD_FUNC(IsWow64Process, kernel32, FARPROC)

static MAKE_DYNLOAD_FUNC(NtQuerySystemInformation, ntdll, NtQuerySystemInformation_t)

/* Processes and threads */
int Twapi_GetProcessList(TwapiInterpContext *, int objc, Tcl_Obj * CONST objv[]);
int Twapi_EnumProcesses (TwapiInterpContext *ticP);
int Twapi_EnumDeviceDrivers(TwapiInterpContext *ticP);
int Twapi_EnumProcessModules(TwapiInterpContext *ticP, HANDLE phandle);
int TwapiCreateProcessHelper(Tcl_Interp *interp, int func, int objc, Tcl_Obj * CONST objv[]);
int Twapi_NtQueryInformationProcessBasicInformation(Tcl_Interp *interp,
                                                    HANDLE processH);
int Twapi_NtQueryInformationThreadBasicInformation(Tcl_Interp *interp,
                                                   HANDLE threadH);
/* Wrapper around NtQuerySystemInformation to process list */
int Twapi_GetProcessList(
    TwapiInterpContext *ticP,
    int  objc,
    Tcl_Obj *CONST objv[])
{
    struct _SYSTEM_PROCESSES *processP;
    Tcl_Interp *interp = ticP->interp;
    ULONG_PTR pid;
    int      first_iteration;
    void  *bufP;
    ULONG  bufsz;          /* Number of bytes allocated */
    ULONG  dummy;
    NTSTATUS status;
    NtQuerySystemInformation_t NtQuerySystemInformationPtr = Twapi_GetProc_NtQuerySystemInformation();
    Tcl_Obj *resultObj;
    Tcl_Obj *process[30];       /* Actually need only 28 */
    Tcl_Obj *process_fields[30];
    Tcl_Obj *thread[15];        /* Actually need 24 */
    Tcl_Obj *thread_fields[15];
    int      pi;
    int      ti;
    Tcl_Obj *process_fieldObj = NULL;
    Tcl_Obj *thread_fieldObj = NULL;
    Tcl_Obj *field_and_list[2];
    int      flags;
#define TWAPI_F_GETPROCESSLIST_STATE  1
#define TWAPI_F_GETPROCESSLIST_NAME   2
#define TWAPI_F_GETPROCESSLIST_PERF   4
#define TWAPI_F_GETPROCESSLIST_VM     8
#define TWAPI_F_GETPROCESSLIST_IO     16
#define TWAPI_F_GETPROCESSLIST_THREAD 32
#define TWAPI_F_GETPROCESSLIST_THREAD_PERF 64
#define TWAPI_F_GETPROCESSLIST_THREAD_STATE 128
    

    if (NtQuerySystemInformationPtr == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    if (TwapiGetArgs(interp, objc, objv,
                     GETDWORD_PTR(pid), GETINT(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;


    /* We do not bother with MemLifo* because these are large allocations */
    /* TBD - should we use a separate heap for this to avoid fragmentation ? */
    bufsz = 50000;              /* Initial guess based on my system */
    bufP = NULL;
    do {
        if (bufP)
            TwapiFree(bufP);         /* Previous buffer too small */
        bufP = TwapiAlloc(bufsz);

        /* Note for information class 5, the last parameter which
         * corresponds to number of bytes needed is not actually filled
         * in by the system so we ignore it and just double alloc size
         */
        status = (*NtQuerySystemInformationPtr)(5, bufP, bufsz, &dummy);
        bufsz = 2* bufsz;       /* For next iteration if needed */
    } while (status == STATUS_INFO_LENGTH_MISMATCH);

    if (status) {
        TwapiFree(bufP);
        return Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status));
    }    

    /* OK, now we got the info. Loop through to extract information
     * from the process list. See Nebett's Window NT/2000 Native API
     * Reference for details
     */
    resultObj = Tcl_NewListObj(0, NULL);
    processP = bufP;
    first_iteration = 1;
    while (1) {
        ULONG_PTR    this_pid;

        this_pid = processP->ProcessId;

        /* Only include this process if we want all or pid matches */
        if (pid == (ULONG_PTR) (LONG_PTR) -1 || pid == this_pid) {
            /* List contains PID, Process info list pairs (flat list) */
            Tcl_ListObjAppendElement(interp, resultObj,
                                     ObjFromULONG_PTR(processP->ProcessId));

            if (flags) {
                if (first_iteration)
                    process_fields[0] = STRING_LITERAL_OBJ("ProcessId");
                process[0] = ObjFromULONG_PTR(processP->ProcessId);

                pi = 1;
                if (flags & TWAPI_F_GETPROCESSLIST_STATE) {
                    if (first_iteration) {
                        process_fields[1] = STRING_LITERAL_OBJ("InheritedFromProcessId");
                        process_fields[2] = STRING_LITERAL_OBJ("SessionId");
                        process_fields[3] = STRING_LITERAL_OBJ("BasePriority");
                    }
                    process[1] = ObjFromULONG_PTR(processP->InheritedFromProcessId);
                    process[2] = Tcl_NewLongObj(processP->SessionId);
                    process[3] = Tcl_NewLongObj(processP->BasePriority);

                    pi = 4;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_NAME) {
                    if (first_iteration)
                        process_fields[pi] = STRING_LITERAL_OBJ("ProcessName");
                    if (processP->ProcessId)
                        process[pi] = ObjFromLSA_UNICODE_STRING(&processP->ProcessName);
                    else
                        process[pi] = STRING_LITERAL_OBJ("System Idle Process");
                    pi += 1;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_PERF) {
                    if (first_iteration) {
                        process_fields[pi] = STRING_LITERAL_OBJ("HandleCount");
                        process_fields[pi+1] = STRING_LITERAL_OBJ("ThreadCount");
                        process_fields[pi+2] = STRING_LITERAL_OBJ("CreateTime");
                        process_fields[pi+3] = STRING_LITERAL_OBJ("UserTime");
                        process_fields[pi+4] = STRING_LITERAL_OBJ("KernelTime");
                    }
                    process[pi] = Tcl_NewLongObj(processP->HandleCount);
                    process[pi+1] = Tcl_NewLongObj(processP->ThreadCount);
                    process[pi+2] = ObjFromLARGE_INTEGER(processP->CreateTime);
                    process[pi+3] = ObjFromLARGE_INTEGER(processP->UserTime);
                    process[pi+4] = ObjFromLARGE_INTEGER(processP->KernelTime);

                    pi += 5;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_VM) {
                    if (first_iteration) {
                        process_fields[pi] = STRING_LITERAL_OBJ("VmCounters.PeakVirtualSize");
                        process_fields[pi+1] = STRING_LITERAL_OBJ("VmCounters.VirtualSize");
                        process_fields[pi+2] = STRING_LITERAL_OBJ("VmCounters.PageFaultCount");
                        process_fields[pi+3] = STRING_LITERAL_OBJ("VmCounters.PeakWorkingSetSize");
                        process_fields[pi+4] = STRING_LITERAL_OBJ("VmCounters.WorkingSetSize");
                        process_fields[pi+5] = STRING_LITERAL_OBJ("VmCounters.QuotaPeakPagedPoolUsage");
                        process_fields[pi+6] = STRING_LITERAL_OBJ("VmCounters.QuotaPagedPoolUsage");
                        process_fields[pi+7] = STRING_LITERAL_OBJ("VmCounters.QuotaPeakNonPagedPoolUsage");
                        process_fields[pi+8] = STRING_LITERAL_OBJ("VmCounters.QuotaNonPagedPoolUsage");
                        process_fields[pi+9] = STRING_LITERAL_OBJ("VmCounters.PagefileUsage");
                        process_fields[pi+10] = STRING_LITERAL_OBJ("VmCounters.PeakPagefileUsage");
                    }

                    process[pi] = ObjFromSIZE_T(processP->VmCounters.PeakVirtualSize);
                    process[pi+1] = ObjFromSIZE_T(processP->VmCounters.VirtualSize);
                    process[pi+2] = Tcl_NewLongObj(processP->VmCounters.PageFaultCount);
                    process[pi+3] = ObjFromSIZE_T(processP->VmCounters.PeakWorkingSetSize);
                    process[pi+4] = ObjFromSIZE_T(processP->VmCounters.WorkingSetSize);
                    process[pi+5] = ObjFromSIZE_T(processP->VmCounters.QuotaPeakPagedPoolUsage);
                    process[pi+6] = ObjFromSIZE_T(processP->VmCounters.QuotaPagedPoolUsage);
                    process[pi+7] = ObjFromSIZE_T(processP->VmCounters.QuotaPeakNonPagedPoolUsage);
                    process[pi+8] = ObjFromSIZE_T(processP->VmCounters.QuotaNonPagedPoolUsage);
                    process[pi+9] = ObjFromSIZE_T(processP->VmCounters.PagefileUsage);
                    process[pi+10] = ObjFromSIZE_T(processP->VmCounters.PeakPagefileUsage);

                    pi += 11;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_IO) {
                    if (first_iteration) {
                        process_fields[pi] = STRING_LITERAL_OBJ("IoCounters.ReadOperationCount");
                        process_fields[pi+1] = STRING_LITERAL_OBJ("IoCounters.WriteOperationCount");
                        process_fields[pi+2] = STRING_LITERAL_OBJ("IoCounters.OtherOperationCount");
                        process_fields[pi+3] = STRING_LITERAL_OBJ("IoCounters.ReadTransferCount");
                        process_fields[pi+4] = STRING_LITERAL_OBJ("IoCounters.WriteTransferCount");
                        process_fields[pi+5] = STRING_LITERAL_OBJ("IoCounters.OtherTransferCount");
                    }

                    process[pi] = ObjFromULONGLONG(processP->IoCounters.ReadOperationCount);
                    process[pi+1] = ObjFromULONGLONG(processP->IoCounters.WriteOperationCount);
                    process[pi+2] = ObjFromULONGLONG(processP->IoCounters.OtherOperationCount);
                    process[pi+3] = ObjFromULONGLONG(processP->IoCounters.ReadTransferCount);
                    process[pi+4] = ObjFromULONGLONG(processP->IoCounters.WriteTransferCount);
                    process[pi+5] = ObjFromULONGLONG(processP->IoCounters.OtherTransferCount);

                    pi += 6;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_THREAD) {
                    SYSTEM_THREADS *threadP;
                    DWORD           i;
                    Tcl_Obj        *threadlistObj;

                    threadlistObj = Tcl_NewListObj(0, NULL);

                    threadP = &processP->Threads[0];
                    for (i=0; i < processP->ThreadCount; ++i, ++threadP) {
                        Tcl_ListObjAppendElement(interp,
                                                 threadlistObj,
                                                 ObjFromDWORD_PTR(threadP->ClientId.UniqueThread));

                        if (first_iteration) {
                            thread_fields[0] = STRING_LITERAL_OBJ("ClientId.UniqueProcess");
                            thread_fields[1] = STRING_LITERAL_OBJ("ClientId.UniqueThread");
                        }
                        thread[0] = ObjFromDWORD_PTR(threadP->ClientId.UniqueProcess);
                        thread[1] = ObjFromDWORD_PTR(threadP->ClientId.UniqueThread);
                        
                        ti = 2;
                        if (flags & TWAPI_F_GETPROCESSLIST_THREAD_STATE) {
                            if (first_iteration) {
                                thread_fields[ti] = STRING_LITERAL_OBJ("BasePriority");
                                thread_fields[ti+1] = STRING_LITERAL_OBJ("Priority");
                                thread_fields[ti+2] = STRING_LITERAL_OBJ("StartAddress");
                                thread_fields[ti+3] = STRING_LITERAL_OBJ("State");
                                thread_fields[ti+4] = STRING_LITERAL_OBJ("WaitReason");
                            }

                            thread[ti] = Tcl_NewLongObj(threadP->BasePriority);
                            thread[ti+1] = Tcl_NewLongObj(threadP->Priority);
                            thread[ti+2] = ObjFromDWORD_PTR(threadP->StartAddress);
                            thread[ti+3] = Tcl_NewLongObj(threadP->State);
                            thread[ti+4] = Tcl_NewLongObj(threadP->WaitReason);

                            ti += 5;
                        }
                        if (flags & TWAPI_F_GETPROCESSLIST_THREAD_PERF) {
                            if (first_iteration) {
                                thread_fields[ti] = STRING_LITERAL_OBJ("WaitTime");
                                thread_fields[ti+1] = STRING_LITERAL_OBJ("ContextSwitchCount");
                                thread_fields[ti+2] = STRING_LITERAL_OBJ("CreateTime");
                                thread_fields[ti+3] = STRING_LITERAL_OBJ("UserTime");
                                thread_fields[ti+4] = STRING_LITERAL_OBJ("KernelTime");
                            }

                            thread[ti] = Tcl_NewLongObj(threadP->WaitTime);
                            thread[ti+1] = Tcl_NewLongObj(threadP->ContextSwitchCount);
                            thread[ti+2] = ObjFromLARGE_INTEGER(threadP->CreateTime);
                            thread[ti+3] = ObjFromLARGE_INTEGER(threadP->UserTime);
                            thread[ti+4] = ObjFromLARGE_INTEGER(threadP->KernelTime);

                            ti += 5;
                        }

                        if (thread_fieldObj == NULL)
                            thread_fieldObj = Tcl_NewListObj(ti, thread_fields);
                        Tcl_ListObjAppendElement(interp, threadlistObj, Tcl_NewListObj(ti, thread));
                    }

                    process_fields[pi] = STRING_LITERAL_OBJ("Threads");
                    field_and_list[0] = thread_fieldObj;
                    field_and_list[1] = threadlistObj;
                    process[pi] = Tcl_NewListObj(2, field_and_list);
                    pi += 1;
                }

                if (first_iteration) {
                    process_fieldObj = Tcl_NewListObj(pi, process_fields);
                }
                Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewListObj(pi, process));
            }

            first_iteration = 0;

            /* If PID was specified and we found it, all done */
            if (pid != (ULONG_PTR) (LONG_PTR) -1)
                break;
        }

        /* Point to the next process struct */
        if (processP->NextEntryDelta == 0)
            break;              /* This was the last one */
        processP = (struct _SYSTEM_PROCESSES *) (processP->NextEntryDelta + (char *) processP);
    }

    if (flags && process_fieldObj) {
        field_and_list[0] = process_fieldObj;
        field_and_list[1] = resultObj;
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, field_and_list));
    } else
        Tcl_SetObjResult(interp, resultObj);

    TwapiFree(bufP);
    return TCL_OK;
}

/*
 * Helper to enumerate processes, or modules for a
 * process with a given pid
 * type = 0 for process, 1 for modules, 2 for drivers
 */
static int Twapi_EnumProcessesModules (TwapiInterpContext *ticP, int type, HANDLE phandle)
{
    void  *bufP;
    DWORD  buf_filled;
    DWORD  buf_size;
    int    result;
    BOOL   status;

    /*
     * EnumProcesses/EnumProcessModules do not return error if the 
     * buffer is too small
     * so we keep looping until the "required bytes" or "returned bytes"
     * is less than what
     * we supplied at which point we know the buffer was large enough
     */
    
    bufP = MemLifoPushFrame(&ticP->memlifo, 2000, &buf_size);
    result = TCL_ERROR;
    do {
        switch (type) {
        case 0:
            /* Looking for processes */
            status = EnumProcesses(bufP, buf_size, &buf_filled);
            break;

        case 1:
            /* Looking for modules for a process */
            status = EnumProcessModules(phandle, bufP, buf_size, &buf_filled);
            break;

        case 2:
            /* Looking for drivers */
            status = EnumDeviceDrivers(bufP, buf_size, &buf_filled);
            break;
        }

        if (! status) {
            TwapiReturnSystemError(ticP->interp);
            break;
        }

        /* Even if no errors, we may not have all data if provided buffer
           was not large enough */
        if (buf_filled < buf_size) {
            /* We have all the entries */
            result = TCL_OK;
            break;
        }

        /* Loop with bigger buffer */
        buf_size *= 2;
        MemLifoPopFrame(&ticP->memlifo);
        bufP = MemLifoPushFrame(&ticP->memlifo, buf_size, NULL);
    }  while (1);


    if (result == TCL_OK) {
        Tcl_Obj **objvP;
        int i, num;

        if (type == 0) {
            /* PID's - DWORDS */
            num = buf_filled/sizeof(DWORD);
            objvP = MemLifoAlloc(&ticP->memlifo, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = Tcl_NewLongObj(((DWORD *)bufP)[i]);
            }
        } else if (type == 1) {
            /* Module handles */
            num = buf_filled/sizeof(HMODULE);
            objvP = MemLifoAlloc(&ticP->memlifo, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = ObjFromOpaque(((HMODULE *)bufP)[i], "HMODULE");
            }
        } else {
            /* device handles - pointers, again not real handles */
            num = buf_filled/sizeof(LPVOID);
            objvP = MemLifoAlloc(&ticP->memlifo, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = ObjFromOpaque(((HMODULE *)bufP)[i], "HMODULE");
            }
        }

        Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(num, objvP));
    }

    MemLifoPopFrame(&ticP->memlifo);

    return result;
}

int Twapi_EnumProcesses (TwapiInterpContext *ticP) 
{
    return Twapi_EnumProcessesModules(ticP, 0, NULL);
}

int Twapi_EnumProcessModules(TwapiInterpContext *ticP, HANDLE phandle) 
{
    return Twapi_EnumProcessesModules(ticP, 1, phandle);
}

int Twapi_EnumDeviceDrivers(TwapiInterpContext *ticP)
{
    return Twapi_EnumProcessesModules(ticP, 2, NULL);
}

int Twapi_WaitForInputIdle(
    Tcl_Interp *interp,
    HANDLE hProcess,
    DWORD dwMilliseconds
)
{
    DWORD error;
    switch (WaitForInputIdle(hProcess, dwMilliseconds)) {
    case 0:
        Tcl_SetObjResult(interp, Tcl_NewIntObj(1));
        break;
    case WAIT_TIMEOUT:
        Tcl_SetObjResult(interp, Tcl_NewIntObj(0));
        break;
    default:
        error = GetLastError();
        if (error)
            return TwapiReturnSystemError(interp);
        /* No error - probably a console process. Treat as always ready */
        Tcl_SetObjResult(interp, Tcl_NewIntObj(1));
        break;
    }
    return TCL_OK;
}

int ListObjToSTARTUPINFO(Tcl_Interp *interp, Tcl_Obj *siObj, STARTUPINFOW *siP)
{
    int  objc;
    long flags;
    long lval;
    Tcl_Obj **objvP;

    TwapiZeroMemory(siP, sizeof(*siP));

    siP->cb = sizeof(*siP);

    if (Tcl_ListObjGetElements(interp, siObj, &objc, &objvP) != TCL_OK)
        return TCL_ERROR;

    if (objc != 12) {
        Tcl_SetResult(interp, "Invalid number of list elements for STARTUPINFO structure", TCL_STATIC);
        return TCL_ERROR;
    }

    if (Tcl_GetLongFromObj(interp, objvP[9], &flags) != TCL_OK)
        return TCL_ERROR;
    siP->dwFlags = (DWORD) flags;

    siP->lpDesktop = Tcl_GetUnicode(objvP[0]);
    if (!lstrcmpW(siP->lpDesktop, NULL_TOKEN_L))
        siP->lpDesktop = NULL;

    siP->lpTitle = Tcl_GetUnicode(objvP[1]);

    if (flags & STARTF_USEPOSITION) {
        if (Tcl_GetLongFromObj(interp, objvP[2], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->dwX = (DWORD) lval;
        if (Tcl_GetLongFromObj(interp, objvP[3], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->dwY = (DWORD) lval;
    }

    if (flags & STARTF_USESIZE) {
        if (Tcl_GetLongFromObj(interp, objvP[4], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->dwXSize = (DWORD) lval;
        if (Tcl_GetLongFromObj(interp, objvP[5], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->dwYSize = (DWORD) lval;
    }

    if (flags & STARTF_USECOUNTCHARS) {
        if (Tcl_GetLongFromObj(interp, objvP[6], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->dwXCountChars = (DWORD) lval;
        if (Tcl_GetLongFromObj(interp, objvP[7], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->dwYCountChars = (DWORD) lval;
    }
    
    if (flags & STARTF_USEFILLATTRIBUTE) {
        if (Tcl_GetLongFromObj(interp, objvP[8], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->dwFillAttribute = (DWORD) lval;
    }

    if (flags & STARTF_USESHOWWINDOW) {
        if (Tcl_GetLongFromObj(interp, objvP[10], &lval) != TCL_OK)
            return TCL_ERROR;
        siP->wShowWindow = (WORD) lval;
    }
    
    if (flags & STARTF_USESTDHANDLES) {
        if (Tcl_ListObjGetElements(interp, objvP[11], &objc, &objvP) != TCL_OK)
            return TCL_ERROR;

        if (objc != 3) {
            Tcl_SetResult(interp, "Invalid number of standard handles in STARTUPINFO structure", TCL_STATIC);
            return TCL_ERROR;
        }

        if (ObjToHANDLE(interp, objvP[0], &siP->hStdInput) != TCL_OK ||
            ObjToHANDLE(interp, objvP[0], &siP->hStdOutput) != TCL_OK ||
            ObjToHANDLE(interp, objvP[0], &siP->hStdError) != TCL_OK)
            return TCL_ERROR;
    }
    
    return TCL_OK;
}



Tcl_Obj *ObjFromMODULEINFO(LPMODULEINFO miP)
{
    Tcl_Obj *objv[6];

    objv[0] = STRING_LITERAL_OBJ("lpBaseOfDll");
    objv[1] = ObjFromDWORD_PTR(miP->lpBaseOfDll);
    objv[2] = STRING_LITERAL_OBJ("SizeOfImage");
    objv[3] = Tcl_NewLongObj(miP->SizeOfImage);
    objv[4] = STRING_LITERAL_OBJ("EntryPoint");
    objv[5] = ObjFromDWORD_PTR(miP->EntryPoint);
    return Tcl_NewListObj(6, objv);
}

#ifndef TWAPI_LEAN
Tcl_Obj *ListObjFromPROCESS_BASIC_INFORMATION(
    Tcl_Interp *interp, const PROCESS_BASIC_INFORMATION *pbiP
)
{
    Tcl_Obj *objv[6];

    objv[0] = Tcl_NewLongObj(pbiP->ExitStatus);
    objv[1] = ObjFromULONG_PTR((ULONG_PTR)pbiP->PebBaseAddress);
    objv[2] = ObjFromULONG_PTR(pbiP->AffinityMask);
    objv[3] = Tcl_NewLongObj(pbiP->BasePriority);
    objv[4] = ObjFromULONG_PTR((ULONG_PTR) pbiP->UniqueProcessId);
    objv[5] = ObjFromULONG_PTR((ULONG_PTR)pbiP->InheritedFromUniqueProcessID);

    return Tcl_NewListObj(6, objv);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
static int TwapiNtQueryInformationProcess(Tcl_Interp *interp, HANDLE processH, int info_type, void *buf, ULONG buf_sz)
{
    ULONG    buf_retsz;
    NTSTATUS status;
    NtQueryInformationProcess_t NtQueryInformationProcessPtr = Twapi_GetProc_NtQueryInformationProcess();

    if (NtQueryInformationProcessPtr == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    status = (*NtQueryInformationProcessPtr)(processH,
                                             info_type,// ProcessBasicInformation
                                             buf,
                                             buf_sz,
                                             &buf_retsz);
    if (! NT_SUCCESS(status)) {
        Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status));
        return TCL_ERROR;
    }
    
    return TCL_OK;
}
#endif

#ifndef TWAPI_LEAN
int Twapi_NtQueryInformationProcessBasicInformation(Tcl_Interp *interp, HANDLE processH)
{
    PROCESS_BASIC_INFORMATION pbi;

    if (TwapiNtQueryInformationProcess(interp, processH, 0,
                                        &pbi, sizeof(pbi)) != TCL_OK)
        return TCL_ERROR;

    Tcl_SetObjResult(interp,
                     ListObjFromPROCESS_BASIC_INFORMATION(interp, &pbi));
    
    return TCL_OK;
}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
Tcl_Obj *ListObjFromTHREAD_BASIC_INFORMATION(
    Tcl_Interp *interp, const THREAD_BASIC_INFORMATION *tbiP
)
{
    Tcl_Obj *objv[6];
    Tcl_Obj *ids[2];

    ids[0] = Tcl_NewLongObj((long) tbiP->ClientId.UniqueProcess);
    ids[1] = Tcl_NewLongObj((long) tbiP->ClientId.UniqueThread);

    objv[0] = Tcl_NewLongObj(tbiP->ExitStatus);
    objv[1] = ObjFromDWORD_PTR((DWORD_PTR) tbiP->TebBaseAddress);
    objv[2] = Tcl_NewListObj(2, ids);
    objv[3] = ObjFromULONG_PTR(tbiP->AffinityMask);
    objv[4] = Tcl_NewLongObj(tbiP->Priority);
    objv[5] = Tcl_NewLongObj(tbiP->BasePriority);


    return Tcl_NewListObj(6, objv);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
static int TwapiNtQueryInformationThread(Tcl_Interp *interp, HANDLE threadH, int info_type, void *buf, ULONG buf_sz)
{
    ULONG    buf_retsz;
    NTSTATUS status;
    NtQueryInformationThread_t NtQueryInformationThreadPtr = Twapi_GetProc_NtQueryInformationThread();

    if (NtQueryInformationThreadPtr == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    status = (*NtQueryInformationThreadPtr)(threadH,
                                            info_type,// ThreadBasicInformation
                                            buf,
                                            buf_sz,
                                            &buf_retsz);
    if (! NT_SUCCESS(status)) {
        return Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(status));
    }
    
    return TCL_OK;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int Twapi_NtQueryInformationThreadBasicInformation(Tcl_Interp *interp, HANDLE threadH)
{
    THREAD_BASIC_INFORMATION tbi;

    if (TwapiNtQueryInformationThread(interp, threadH, 0,
                                      &tbi, sizeof(tbi)) != TCL_OK)
        return TCL_ERROR;

    Tcl_SetObjResult(interp,
                     ListObjFromTHREAD_BASIC_INFORMATION(interp, &tbi));
    
    return TCL_OK;
}
#endif // TWAPI_LEAN


BOOL Twapi_IsWow64Process(HANDLE h, BOOL *is_wow64P)
{
    FARPROC IsWow64ProcessPtr = Twapi_GetProc_IsWow64Process();

    if (IsWow64ProcessPtr == NULL) {
        /* If IsWow64Process not available, could not be running on 64 bits */
        *is_wow64P = 0;
        return TRUE;
    }
    
    if (IsWow64ProcessPtr(h, is_wow64P)) {
        return TRUE;
    }
    
    return FALSE;               /* Some error */
}

#ifdef FRAGILE
/* Emulates the rundll32.exe interface */
int Twapi_Rundll(Tcl_Interp *interp, LPCSTR dll, LPCSTR function, HWND hwnd, LPCWSTR cmdline, int cmdshow)
{
    HINSTANCE hInst;

    hInst = LoadLibraryA(dll);
    if (hInst) {
        FARPROC fn;
        fn = GetProcAddress(hInst, function);
        if (fn) {
            Tcl_SetObjResult(
                interp,
                Tcl_NewIntObj((*fn)(hwnd, hInst, cmdline, cmdshow))
                );
            return TCL_OK;
        }
    }
    
    return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
}
#endif

Tcl_Obj *ObjFromPROCESS_INFORMATION(PROCESS_INFORMATION *piP)
{
    Tcl_Obj *objv[4];
    objv[0] = ObjFromHANDLE(piP->hProcess);
    objv[1] = ObjFromHANDLE(piP->hThread);
    objv[2] = Tcl_NewLongObj(piP->dwProcessId);
    objv[3] = Tcl_NewLongObj(piP->dwThreadId);
    return Tcl_NewListObj(4, objv);
}

int TwapiCreateProcessHelper(Tcl_Interp *interp, int asuser, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE tokH;
    SECURITY_ATTRIBUTES *pattrP = NULL;
    SECURITY_ATTRIBUTES *tattrP = NULL;
    int inherit;
    DWORD flags;
    STARTUPINFOW startinfo;
    LPWSTR envP = NULL;
    BOOL status = 0;
    PROCESS_INFORMATION pi;
    LPWSTR appname, cmdline, curdir;

    if (asuser != 0 && asuser != 1)
        return TwapiReturnErrorEx(interp, TWAPI_BUG,
                                  Tcl_ObjPrintf("Invalid asuser value %d.",
                                                asuser));
    if (objc != (9+asuser))
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (asuser) {
        if (ObjToHANDLE(interp, objv[0], &tokH) != TCL_OK)
            return TCL_ERROR;
    }
        
    pattrP = NULL;              /* May be alloc'ed even on error */
    tattrP = NULL;              /* May be alloc'ed even on error */
    envP = NULL;
    if (TwapiGetArgs(interp, objc-asuser, objv+asuser,
                     GETNULLIFEMPTY(appname),   GETNULLIFEMPTY(cmdline),
                     GETVAR(pattrP, ObjToPSECURITY_ATTRIBUTES),
                     GETVAR(tattrP, ObjToPSECURITY_ATTRIBUTES),
                     GETINT(inherit), GETINT(flags),
                     ARGSKIP,   GETNULLIFEMPTY(curdir),
                     GETVAR(startinfo, ListObjToSTARTUPINFO),
                     ARGEND) != TCL_OK)
        goto vamoose;           /* So any allocs can be freed */

    envP = ObjToLPWSTR_WITH_NULL(objv[asuser+6]);
    if (envP) {
        if (Twapi_ConvertTclListToMultiSz(interp, objv[asuser+6], (LPWSTR*) &envP) == TCL_ERROR) {
            goto vamoose;       /* So any allocs can be freed */
        }
        // Note envP is dynamically allocated
    }
    
    if (ListObjToSTARTUPINFO(interp, objv[asuser+8], &startinfo) != TCL_OK)
        goto vamoose;

    if (! asuser)
        status = CreateProcessW(
            appname,
            cmdline,
            pattrP, // lpProcessAttributes
            tattrP, // lpThreadAttributes
            inherit, // bInheritHandles
            flags, // dwCreationFlags
            envP, // lpEnvironment
            curdir,
            &startinfo, // lpStartupInfo
            &pi //PROCESS_INFORMATION * (output)
            );
    else
        status = CreateProcessAsUserW(
            tokH, // User token
            appname,
            cmdline,
            pattrP, // lpProcessAttributes
            tattrP, // lpThreadAttributes
            inherit, // bInheritHandles
            flags, // dwCreationFlags
            envP, // lpEnvironment
            curdir,
            &startinfo, // lpStartupInfo
            &pi //PROCESS_INFORMATION * (output)
            );

    if (status) {
        Tcl_SetObjResult(interp, ObjFromPROCESS_INFORMATION(&pi));
    } else {
        TwapiReturnSystemError(interp);
        // Do not return just yet as we have to free buffers
    }        
vamoose:
    if (pattrP)
        TwapiFreeSECURITY_ATTRIBUTES(pattrP);
    if (tattrP)
        TwapiFreeSECURITY_ATTRIBUTES(tattrP);
    if (envP)
        TwapiFree(envP);

    return status ? TCL_OK : TCL_ERROR;
}



#if 0 /* TBD - XP only */
EXCEPTION_ON_FALSE GetProcessMemoryInfo(
    HANDLE Process,
    PPROCESS_MEMORY_COUNTERS ppsmemCounters,
    DWORD cb
    );
#endif

static int Twapi_ProcessCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s;
    DWORD dw, dw2, dw3;
    union {
        DWORD_PTR dwp;
        WCHAR buf[MAX_PATH+1];
        MODULEINFO moduleinfo;
    } u;
    HANDLE h;
    HMODULE hmod;
    LPVOID pv;
    TwapiResult result;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_EnumProcesses(ticP);
    case 2:
        return Twapi_EnumDeviceDrivers(ticP);
    case 3:
        result.type = TRT_DWORD;
        result.value.ival = GetCurrentThreadId();
        break;
    case 4:
        result.type = TRT_HANDLE;
        result.value.hval = GetCurrentThread();
        break;
    case 5: // CreateProcess
    case 6: // CreateProcessAsUser
        return TwapiCreateProcessHelper(interp, func==6, objc-2, objv+2);
    case 7: // ReadProcessMemory
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETDWORD_PTR(u.dwp), GETVOIDP(pv),
                         GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type =
            ReadProcessMemory(h, (void *)u.dwp, pv, dw, &result.value.dwp)
            ? TRT_DWORD_PTR : TRT_GETLASTERROR;
        break;
    case 8: // GetModuleFileName
    case 9: // GetModuleBaseName
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETHANDLET(hmod, HMODULE),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if ((func == 8 ?
             GetModuleFileNameExW
             : GetModuleBaseNameW)
            (h, hmod, u.buf, ARRAYSIZE(u.buf))) {
            result.type = TRT_UNICODE;
            result.value.unicode.str = u.buf;
            result.value.unicode.len = -1;
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 10: // GetModuleInformation
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETHANDLET(hmod, HMODULE),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (GetModuleInformation(h, hmod,
                                 &u.moduleinfo, sizeof(u.moduleinfo))) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromMODULEINFO(&u.moduleinfo);
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 11:
        return Twapi_GetProcessList(ticP, objc-2, objv+2);
    case 12:
        break; // UNUSED
    case 13:
    case 14:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 13:
            result.type = TRT_DWORD;
            result.value.ival = SetThreadExecutionState(dw);
            break;
        case 14:
            result.type = ProcessIdToSessionId(dw, &result.value.ival) ? TRT_DWORD 
                : TRT_GETLASTERROR;
            break;
        }
        break;
    case 15:
    case 16:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETINT(dw), GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HANDLE;
        result.value.hval = (func == 15 ? OpenProcess : OpenThread)(dw, dw2, dw3);
        break;
    case 17:
    case 18:
    case 19:
    case 20:
    case 21:
    case 22:
    case 23:
    case 24:
    case 25:
    case 26:
    case 27:
    case 28:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(h), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
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
        case 25:
            result.type = GetExitCodeProcess(h, &result.value.ival)
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 26:
            result.value.unicode.len = GetProcessImageFileNameW(h, u.buf, ARRAYSIZE(u.buf));
            if (result.value.unicode.len) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
            } else {
                result.type = TRT_GETLASTERROR;
            }
            break;
        case 27:                /* GetDeviceDriverBaseNameW */
        case 28:                /* GetDeviceDriverFileNameW */
            if ((func == 27 ?
                 GetDeviceDriverBaseNameW
                 : GetDeviceDriverFileNameW) (
                     (LPVOID) h,
                     u.buf,
                     ARRAYSIZE(u.buf))) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        }
        break;
    case 29:
    case 30:
    case 31:
    case 32:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 29:
            return Twapi_WaitForInputIdle(interp, h, dw);
        case 30:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetPriorityClass(h, dw);
            break;
        case 31:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetThreadPriority(h, dw);
            break;
        case 32:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = TerminateProcess(h, dw);
            break;
        }
        break;
    case 33: // GetModuleHandleEx
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETINT(dw), ARGSKIP,
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        if (dw & GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS) {
            /* Argument is address in a module */
            if (ObjToDWORD_PTR(interp, objv[3], &u.dwp) != TCL_OK)
                return TCL_ERROR;
        } else {
            s = Tcl_GetUnicode(objv[3]);
            NULLIFY_EMPTY(s);
            u.dwp = (DWORD_PTR) s;
        }
        if (GetModuleHandleExW(dw, (LPCWSTR) u.dwp, &result.value.hmodule))
            result.type = TRT_HMODULE;
        else
            result.type = TRT_GETLASTERROR;
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_ProcessInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::ProcessCall", Twapi_ProcessCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::ProcessCall", # code_); \
    } while (0);

    CALL_(EnumProcesses, 1);
    CALL_(EnumDeviceDrivers, 2);
    CALL_(GetCurrentThreadId, 3);
    CALL_(GetCurrentThread, 4);
    CALL_(CreateProcess, 5);
    CALL_(CreateProcessAsUser, 6);
    CALL_(ReadProcessMemory, 7);
    CALL_(GetModuleFileNameEx, 8);
    CALL_(GetModuleBaseName, 9);
    CALL_(GetModuleInformation, 10);
    CALL_(Twapi_GetProcessList, 11);
    CALL_(SetThreadExecutionState, 13);
    CALL_(ProcessIdToSessionId, 14);
    CALL_(OpenProcess, 15);
    CALL_(OpenThread, 16);
    CALL_(EnumProcessModules, 17);
    CALL_(IsWow64Process, 18);
    CALL_(ResumeThread, 19);
    CALL_(SuspendThread, 20);
    CALL_(GetPriorityClass, 21);
    CALL_(Twapi_NtQueryInformationProcessBasicInformation, 22);
    CALL_(Twapi_NtQueryInformationThreadBasicInformation, 23);
    CALL_(GetThreadPriority, 24);
    CALL_(GetExitCodeProcess, 25);
    CALL_(GetProcessImageFileName, 26); /* TBD - Tcl wrapper */
    CALL_(GetDeviceDriverBaseName, 27);
    CALL_(GetDeviceDriverFileName, 28);

                                                   
    CALL_(WaitForInputIdle, 29);
    CALL_(SetPriorityClass, 30);
    CALL_(SetThreadPriority, 31);
    CALL_(TerminateProcess, 32);

    CALL_(GetModuleHandleEx, 33);

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
int Twapi_process_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_ProcessInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

