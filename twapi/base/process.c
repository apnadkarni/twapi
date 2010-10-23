/* 
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to process information */

#include "twapi.h"

#ifndef TWAPI_LEAN
typedef NTSTATUS (WINAPI *NtQueryInformationProcess_t)(HANDLE, int, PVOID, ULONG, PULONG);
MAKE_DYNLOAD_FUNC(NtQueryInformationProcess, ntdll, NtQueryInformationProcess_t)
typedef NTSTATUS (WINAPI *NtQueryInformationThread_t)(HANDLE, int, PVOID, ULONG, PULONG);
MAKE_DYNLOAD_FUNC(NtQueryInformationThread, ntdll, NtQueryInformationThread_t)
MAKE_DYNLOAD_FUNC(IsWow64Process, kernel32, FARPROC)
#endif

#ifndef TWAPI_LEAN
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
#endif

#ifndef TWAPI_LEAN

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
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
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
            /* Module handles - pointers but returned as integer as they are
               not really valid typed handles */
            num = buf_filled/sizeof(HMODULE);
            objvP = MemLifoAlloc(&ticP->memlifo, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = ObjFromDWORD_PTR(((HMODULE *)bufP)[i]);
            }
        } else {
            /* device handles - pointers, again not real handles */
            num = buf_filled/sizeof(LPVOID);
            objvP = MemLifoAlloc(&ticP->memlifo, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = ObjFromDWORD_PTR(((LPVOID *)bufP)[i]);
            }
        }

        Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(num, objvP));
    }

    MemLifoPopFrame(&ticP->memlifo);

    return result;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int Twapi_EnumProcesses (TwapiInterpContext *ticP) 
{
    return Twapi_EnumProcessesModules(ticP, 0, NULL);
}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
int Twapi_EnumProcessModules(TwapiInterpContext *ticP, HANDLE phandle) 
{
    return Twapi_EnumProcessesModules(ticP, 1, phandle);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int Twapi_EnumDeviceDrivers(TwapiInterpContext *ticP)
{
    return Twapi_EnumProcessesModules(ticP, 2, NULL);
}
#endif // TWAPI_LEAN

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

    ZeroMemory(siP, sizeof(*siP));

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
        Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewUnicodeObj(argv[i], -1));
    }

    Tcl_SetObjResult(interp, resultObj);

    GlobalFree(argv);
    return TCL_OK;
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
        return TwapiReturnTwapiError(interp, "Invalid asuser value", TWAPI_BUG);

    if (objc != (9+asuser))
        return TwapiReturnTwapiError(interp, NULL, TWAPI_BAD_ARG_COUNT);

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



#if 0 /* TBD - XP or later only */
DWORD GetProcessImageFileNameW (
    HANDLE  hProcess,
    LPWSTR  lpFilename,
    DWORD   nSize
);
#endif
#if 0 /* TBD - XP only */
EXCEPTION_ON_FALSE GetProcessMemoryInfo(
    HANDLE Process,
    PPROCESS_MEMORY_COUNTERS ppsmemCounters,
    DWORD cb
    );
#endif

#if 0 /* TBD - XP or later only EVEN THOUGH MSDN Docs CLAIN Win2k ALSO */
EXCEPTION_ON_FALSE GetPerformanceInfo (
    PPERFORMACE_INFORMATION pPerformanceInformation,
    DWORD cb
);
#endif

