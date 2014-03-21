/* 
 * Copyright (c) 2003-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to process information */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

typedef NTSTATUS (WINAPI *NtQueryInformationProcess_t)(HANDLE, int, PVOID, ULONG, PULONG);
static MAKE_DYNLOAD_FUNC(NtQueryInformationProcess, ntdll, NtQueryInformationProcess_t)
typedef NTSTATUS (WINAPI *NtQueryInformationThread_t)(HANDLE, int, PVOID, ULONG, PULONG);
static MAKE_DYNLOAD_FUNC(NtQueryInformationThread, ntdll, NtQueryInformationThread_t)
static MAKE_DYNLOAD_FUNC(IsWow64Process, kernel32, FARPROC)

static MAKE_DYNLOAD_FUNC(NtQuerySystemInformation, ntdll, NtQuerySystemInformation_t)

/* Processes and threads */
int Twapi_NtQueryInformationProcessBasicInformation(Tcl_Interp *interp,
                                                    HANDLE processH);
int Twapi_NtQueryInformationThreadBasicInformation(Tcl_Interp *interp,
                                                   HANDLE threadH);
/* Wrapper around NtQuerySystemInformation to process list */
static TCL_RESULT Twapi_GetProcessList(
    Tcl_Interp *interp,
    int  objc,
    Tcl_Obj *CONST objv[])
{
    struct _SYSTEM_PROCESSES *processP;
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
    resultObj = ObjEmptyList();
    processP = bufP;
    first_iteration = 1;
    while (1) {
        ULONG_PTR    this_pid;

        this_pid = processP->ProcessId;

        /* NOTE: FIELD NAMES HERE MATCH THOSE AT TCL LEVEL. IF YOU
           CHANGE ONE, NEED TO CHANGE THE OTHER! */

        /* Only include this process if we want all or pid matches */
        if (pid == (ULONG_PTR) (LONG_PTR) -1 || pid == this_pid) {
            /* List contains PID, Process info list pairs (flat list) */
            if (!flags) {
                ObjAppendElement(interp, resultObj,
                                 ObjFromULONG_PTR(processP->ProcessId));
            } else {
                if (first_iteration)
                    process_fields[0] = STRING_LITERAL_OBJ("-pid");
                process[0] = ObjFromULONG_PTR(processP->ProcessId);

                pi = 1;
                if (flags & TWAPI_F_GETPROCESSLIST_STATE) {
                    if (first_iteration) {
                        process_fields[1] = STRING_LITERAL_OBJ("-parent");
                        process_fields[2] = STRING_LITERAL_OBJ("-tssession");
                        process_fields[3] = STRING_LITERAL_OBJ("-basepriority");
                    }
                    process[1] = ObjFromULONG_PTR(processP->InheritedFromProcessId);
                    process[2] = ObjFromLong(processP->SessionId);
                    process[3] = ObjFromLong(processP->BasePriority);

                    pi = 4;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_NAME) {
                    if (first_iteration)
                        process_fields[pi] = STRING_LITERAL_OBJ("-name");
                    if (processP->ProcessId)
                        process[pi] = ObjFromLSA_UNICODE_STRING(&processP->ProcessName);
                    else
                        process[pi] = STRING_LITERAL_OBJ("System Idle Process");
                    pi += 1;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_PERF) {
                    if (first_iteration) {
                        process_fields[pi] = STRING_LITERAL_OBJ("-handlecount");
                        process_fields[pi+1] = STRING_LITERAL_OBJ("-threadcount");
                        process_fields[pi+2] = STRING_LITERAL_OBJ("-createtime");
                        process_fields[pi+3] = STRING_LITERAL_OBJ("-usertime");
                        process_fields[pi+4] = STRING_LITERAL_OBJ("-privilegedtime");
                    }
                    process[pi] = ObjFromLong(processP->HandleCount);
                    process[pi+1] = ObjFromLong(processP->ThreadCount);
                    process[pi+2] = ObjFromLARGE_INTEGER(processP->CreateTime);
                    process[pi+3] = ObjFromLARGE_INTEGER(processP->UserTime);
                    process[pi+4] = ObjFromLARGE_INTEGER(processP->KernelTime);

                    pi += 5;
                }

                if (flags & TWAPI_F_GETPROCESSLIST_VM) {
                    if (first_iteration) {
                        process_fields[pi] = STRING_LITERAL_OBJ("-virtualbytespeak");
                        process_fields[pi+1] = STRING_LITERAL_OBJ("-virtualbytes");
                        process_fields[pi+2] = STRING_LITERAL_OBJ("-pagefaults");
                        process_fields[pi+3] = STRING_LITERAL_OBJ("-workingsetpeak");
                        process_fields[pi+4] = STRING_LITERAL_OBJ("-workingset");
                        process_fields[pi+5] = STRING_LITERAL_OBJ("-poolpagedbytespeak");
                        process_fields[pi+6] = STRING_LITERAL_OBJ("-poolpagedbytes");
                        process_fields[pi+7] = STRING_LITERAL_OBJ("-poolnonpagedbytespeak");
                        process_fields[pi+8] = STRING_LITERAL_OBJ("-poolnonpagedbytes");
                        process_fields[pi+9] = STRING_LITERAL_OBJ("-pagefilebytes");
                        process_fields[pi+10] = STRING_LITERAL_OBJ("-pagefilebytespeak");
                    }

                    process[pi] = ObjFromSIZE_T(processP->VmCounters.PeakVirtualSize);
                    process[pi+1] = ObjFromSIZE_T(processP->VmCounters.VirtualSize);
                    process[pi+2] = ObjFromLong(processP->VmCounters.PageFaultCount);
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
                        process_fields[pi] = STRING_LITERAL_OBJ("-ioreadops");
                        process_fields[pi+1] = STRING_LITERAL_OBJ("-iowriteops");
                        process_fields[pi+2] = STRING_LITERAL_OBJ("-iootherops");
                        process_fields[pi+3] = STRING_LITERAL_OBJ("-ioreadbytes");
                        process_fields[pi+4] = STRING_LITERAL_OBJ("-iowritebytes");
                        process_fields[pi+5] = STRING_LITERAL_OBJ("-iootherbytes");
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

                    threadlistObj = ObjEmptyList();

                    threadP = &processP->Threads[0];
                    for (i=0; i < processP->ThreadCount; ++i, ++threadP) {
                        if (first_iteration) {
                            thread_fields[0] = STRING_LITERAL_OBJ("-pid");
                            thread_fields[1] = STRING_LITERAL_OBJ("-tid");
                        }
                        thread[0] = ObjFromDWORD_PTR(threadP->ClientId.UniqueProcess);
                        thread[1] = ObjFromDWORD_PTR(threadP->ClientId.UniqueThread);
                        
                        ti = 2;
                        if (flags & TWAPI_F_GETPROCESSLIST_THREAD_STATE) {
                            if (first_iteration) {
                                thread_fields[ti] = STRING_LITERAL_OBJ("-basepriority");
                                thread_fields[ti+1] = STRING_LITERAL_OBJ("-priority");
                                thread_fields[ti+2] = STRING_LITERAL_OBJ("-startaddress");
                                thread_fields[ti+3] = STRING_LITERAL_OBJ("-state");
                                thread_fields[ti+4] = STRING_LITERAL_OBJ("-waitreason");
                            }

                            thread[ti] = ObjFromLong(threadP->BasePriority);
                            thread[ti+1] = ObjFromLong(threadP->Priority);
                            thread[ti+2] = ObjFromDWORD_PTR(threadP->StartAddress);
                            thread[ti+3] = ObjFromLong(threadP->State);
                            thread[ti+4] = ObjFromLong(threadP->WaitReason);

                            ti += 5;
                        }
                        if (flags & TWAPI_F_GETPROCESSLIST_THREAD_PERF) {
                            if (first_iteration) {
                                thread_fields[ti] = STRING_LITERAL_OBJ("-waittime");
                                thread_fields[ti+1] = STRING_LITERAL_OBJ("-contextswitches");
                                thread_fields[ti+2] = STRING_LITERAL_OBJ("-createtime");
                                thread_fields[ti+3] = STRING_LITERAL_OBJ("-usertime");
                                thread_fields[ti+4] = STRING_LITERAL_OBJ("-privilegedtime");
                            }

                            thread[ti] = ObjFromLong(threadP->WaitTime);
                            thread[ti+1] = ObjFromLong(threadP->ContextSwitchCount);
                            thread[ti+2] = ObjFromLARGE_INTEGER(threadP->CreateTime);
                            thread[ti+3] = ObjFromLARGE_INTEGER(threadP->UserTime);
                            thread[ti+4] = ObjFromLARGE_INTEGER(threadP->KernelTime);

                            ti += 5;
                        }

                        if (thread_fieldObj == NULL)
                            thread_fieldObj = ObjNewList(ti, thread_fields);
                        ObjAppendElement(interp, threadlistObj, ObjNewList(ti, thread));
                    }

                    process_fields[pi] = STRING_LITERAL_OBJ("Threads");
                    field_and_list[0] = thread_fieldObj;
                    field_and_list[1] = threadlistObj;
                    process[pi] = ObjNewList(2, field_and_list);
                    pi += 1;
                }

                if (first_iteration) {
                    process_fieldObj = ObjNewList(pi, process_fields);
                }
                ObjAppendElement(interp, resultObj, ObjNewList(pi, process));
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
        ObjSetResult(interp, ObjNewList(2, field_and_list));
    } else
        ObjSetResult(interp, resultObj);

    TwapiFree(bufP);
    return TCL_OK;
}

/*
 * Helper to enumerate processes, or modules for a
 * process with a given pid
 * type = 0 for process, 1 for modules, 2 for drivers
 */
static int Twapi_EnumProcessesModulesObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int  objc,
    Tcl_Obj *CONST objv[])
{
    void  *bufP;
    DWORD  buf_filled;
    DWORD  buf_size;
    int    result;
    BOOL   status;
    int type;
    HANDLE phandle;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(type), ARGUSEDEFAULT, GETHANDLE(phandle),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    /*
     * EnumProcesses/EnumProcessModules do not return error if the 
     * buffer is too small
     * so we keep looping until the "required bytes" or "returned bytes"
     * is less than what
     * we supplied at which point we know the buffer was large enough
     */
    
    /* TBD - instrument buffer size */
    bufP = MemLifoPushFrame(ticP->memlifoP, 2000, &buf_size);
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
        MemLifoPopFrame(ticP->memlifoP);
        bufP = MemLifoPushFrame(ticP->memlifoP, buf_size, NULL);
    }  while (1);


    if (result == TCL_OK) {
        Tcl_Obj **objvP;
        int i, num;

        if (type == 0) {
            /* PID's - DWORDS */
            num = buf_filled/sizeof(DWORD);
            objvP = MemLifoAlloc(ticP->memlifoP, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = ObjFromLong(((DWORD *)bufP)[i]);
            }
        } else if (type == 1) {
            /* Module handles */
            num = buf_filled/sizeof(HMODULE);
            objvP = MemLifoAlloc(ticP->memlifoP, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = ObjFromOpaque(((HMODULE *)bufP)[i], "HMODULE");
            }
        } else {
            /* device handles - pointers, again not real handles */
            num = buf_filled/sizeof(LPVOID);
            objvP = MemLifoAlloc(ticP->memlifoP, num * sizeof(objvP[0]), NULL);
            for (i = 0; i < num; ++i) {
                objvP[i] = ObjFromOpaque(((HMODULE *)bufP)[i], "HMODULE");
            }
        }

        ObjSetResult(ticP->interp, ObjNewList(num, objvP));
    }

    MemLifoPopFrame(ticP->memlifoP);

    return result;
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
        ObjSetResult(interp, ObjFromInt(1));
        break;
    case WAIT_TIMEOUT:
        ObjSetResult(interp, ObjFromInt(0));
        break;
    default:
        error = GetLastError();
        if (error)
            return TwapiReturnSystemError(interp);
        /* No error - probably a console process. Treat as always ready */
        ObjSetResult(interp, ObjFromInt(1));
        break;
    }
    return TCL_OK;
}

/* Note this allocates memory from ticP->memlifo and expects caller to
   clean up */
static int ListObjToSTARTUPINFO(TwapiInterpContext *ticP, Tcl_Obj *siObj, STARTUPINFOW *siP)
{
    int  objc;
    Tcl_Obj **objvP;
    Tcl_Obj *stdhObj;
    Tcl_Interp *interp = ticP->interp;

    TwapiZeroMemory(siP, sizeof(*siP));

    siP->cb = sizeof(*siP);

    if (ObjGetElements(interp, siObj, &objc, &objvP) != TCL_OK ||
        TwapiGetArgsEx(ticP, objc, objvP,
                       GETTOKENNULL(siP->lpDesktop),
                       GETWSTR(siP->lpTitle),
                       GETINT(siP->dwX), GETINT(siP->dwY),
                       GETINT(siP->dwXSize), GETINT(siP->dwYSize),
                       GETINT(siP->dwXCountChars), GETINT(siP->dwYCountChars),
                       GETINT(siP->dwFillAttribute),
                       GETINT(siP->dwFlags),
                       GETWORD(siP->wShowWindow),
                       GETOBJ(stdhObj), ARGEND) != TCL_OK) {
        ObjSetStaticResult(interp, "Invalid STARTUPINFO list.");
        return TCL_ERROR;
    }

    if (siP->dwFlags & STARTF_USESTDHANDLES) {
        if (ObjGetElements(interp, stdhObj, &objc, &objvP) != TCL_OK)
            return TCL_ERROR;

        if (objc != 3) {
            ObjSetStaticResult(interp, "STARTUPINFO structure handles must be 3.");
            return TCL_ERROR;
        }

        if (ObjToHANDLE(interp, objvP[0], &siP->hStdInput) != TCL_OK ||
            ObjToHANDLE(interp, objvP[1], &siP->hStdOutput) != TCL_OK ||
            ObjToHANDLE(interp, objvP[2], &siP->hStdError) != TCL_OK)
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
    objv[3] = ObjFromLong(miP->SizeOfImage);
    objv[4] = STRING_LITERAL_OBJ("EntryPoint");
    objv[5] = ObjFromDWORD_PTR(miP->EntryPoint);
    return ObjNewList(6, objv);
}

#ifndef TWAPI_LEAN
Tcl_Obj *ListObjFromPROCESS_BASIC_INFORMATION(
    Tcl_Interp *interp, const PROCESS_BASIC_INFORMATION *pbiP
)
{
    Tcl_Obj *objv[6];

    objv[0] = ObjFromLong(pbiP->ExitStatus);
    objv[1] = ObjFromULONG_PTR((ULONG_PTR)pbiP->PebBaseAddress);
    objv[2] = ObjFromULONG_PTR(pbiP->AffinityMask);
    objv[3] = ObjFromLong(pbiP->BasePriority);
    objv[4] = ObjFromULONG_PTR((ULONG_PTR) pbiP->UniqueProcessId);
    objv[5] = ObjFromULONG_PTR((ULONG_PTR)pbiP->InheritedFromUniqueProcessID);

    return ObjNewList(6, objv);
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

    ObjSetResult(interp,
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

    ids[0] = ObjFromLong((long) tbiP->ClientId.UniqueProcess);
    ids[1] = ObjFromLong((long) tbiP->ClientId.UniqueThread);

    objv[0] = ObjFromLong(tbiP->ExitStatus);
    objv[1] = ObjFromDWORD_PTR((DWORD_PTR) tbiP->TebBaseAddress);
    objv[2] = ObjNewList(2, ids);
    objv[3] = ObjFromULONG_PTR(tbiP->AffinityMask);
    objv[4] = ObjFromLong(tbiP->Priority);
    objv[5] = ObjFromLong(tbiP->BasePriority);


    return ObjNewList(6, objv);
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

    ObjSetResult(interp,
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
            ObjSetResult(
                interp,
                ObjFromInt((*fn)(hwnd, hInst, cmdline, cmdshow))
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
    objv[2] = ObjFromLong(piP->dwProcessId);
    objv[3] = ObjFromLong(piP->dwThreadId);
    return ObjNewList(4, objv);
}

static int TwapiCreateProcessHelper(TwapiInterpContext *ticP, int asuser, int objc, Tcl_Obj *CONST objv[])
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
    Tcl_Interp *interp = ticP->interp;
    MemLifoMarkHandle mark;
    Tcl_Obj *startinfoObj, *envObj;

    asuser = (asuser != 0);     /* Make 0 or 1 */

    if (objc != (1+asuser+9))
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (asuser) {
        if (ObjToHANDLE(interp, objv[0], &tokH) != TCL_OK)
            return TCL_ERROR;
    }
        
    mark = MemLifoPushMark(ticP->memlifoP);

    pattrP = NULL;              /* May be alloc'ed even on error */
    tattrP = NULL;              /* May be alloc'ed even on error */
    envP = NULL;
    if (TwapiGetArgsEx(ticP, objc-asuser-1, objv+asuser+1,
                       GETEMPTYASNULL(appname),   GETEMPTYASNULL(cmdline),
                       GETVAR(pattrP, ObjToPSECURITY_ATTRIBUTES),
                       GETVAR(tattrP, ObjToPSECURITY_ATTRIBUTES),
                       GETINT(inherit), GETINT(flags),
                       GETOBJ(envObj),   GETEMPTYASNULL(curdir),
                       GETOBJ(startinfoObj), ARGEND) != TCL_OK)
        goto vamoose;           /* So any allocs/marks can be freed */

    envP = ObjToLPWSTR_WITH_NULL(envObj);
    if (envP) {
        if (ObjToMultiSzEx(interp, envObj, (LPWSTR*) &envP, ticP->memlifoP) == TCL_ERROR)
            goto vamoose;
        /* Note envP is allocated from ticP->memlifo */
    }
    
    if (ListObjToSTARTUPINFO(ticP, startinfoObj, &startinfo) != TCL_OK)
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
        ObjSetResult(interp, ObjFromPROCESS_INFORMATION(&pi));
    } else {
        TwapiReturnSystemError(interp);
        // Do not return just yet as we have to free buffers
    }        
vamoose:
    if (pattrP)
        TwapiFreeSECURITY_ATTRIBUTES(pattrP);
    if (tattrP)
        TwapiFreeSECURITY_ATTRIBUTES(tattrP);

    MemLifoPopMark(mark);
    return status ? TCL_OK : TCL_ERROR;
}


static TCL_RESULT Twapi_CreateProcessObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int  objc,
    Tcl_Obj *CONST objv[])
{
    return TwapiCreateProcessHelper(ticP, 0, objc, objv);
}

static TCL_RESULT Twapi_CreateProcessAsUserObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int  objc,
    Tcl_Obj *CONST objv[])
{
    return TwapiCreateProcessHelper(ticP, 1, objc, objv);
}


static int Twapi_ProcessCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD dw, dw2, dw3;
    union {
    DWORD_PTR dwp;
        WCHAR buf[MAX_PATH+1];
        MODULEINFO moduleinfo;
        PROCESS_MEMORY_COUNTERS_EX pmce;
    } u;
    HANDLE h;
    HMODULE hmod;
    LPVOID pv;
    TwapiResult result;
    int func = PtrToInt(clientdata);
    Tcl_Obj *objs[10];
    SIZE_T sz, sz2;

    --objc;
    ++objv;
    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 6:
        result.type = TRT_HANDLE;
        result.value.hval = GetCurrentThread();
        break;
    case 7: // ReadProcessMemory
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLE(h), /* Process handle */
                         GETDWORD_PTR(u.dwp), /*  address in target process */
                         GETINT(dw),          /* Num bytes to read */
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.obj = ObjFromByteArray(NULL, dw);
        pv = ObjToByteArray(result.value.obj, &dw);
        if (ReadProcessMemory(h, (void *)u.dwp, pv, dw, &sz)) {
            Tcl_SetByteArrayLength(result.value.obj, sz);
            result.type = TRT_OBJ;
        } else {
            ObjDecrRefs(result.value.obj);
            result.type = TRT_GETLASTERROR;
        }
        break;
    case 8: // GetModuleFileName
    case 9: // GetModuleBaseName
        if (TwapiGetArgs(interp, objc, objv,
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
        if (TwapiGetArgs(interp, objc, objv,
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
        return Twapi_GetProcessList(interp, objc, objv);
    case 12:
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h), GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        sz = dw == (DWORD)-1 ? (SIZE_T) -1 : dw;
        sz2 = dw2 == (DWORD)-1 ? (SIZE_T) -1 : dw2;
        result.type = SetProcessWorkingSetSize(h, sz, sz2) ? TRT_EMPTY : TRT_GETLASTERROR;
        break;
            
    case 13:
    case 14:
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 13:
            result.type = TRT_DWORD;
            result.value.uval = SetThreadExecutionState(dw);
            break;
        case 14:
            result.type = ProcessIdToSessionId(dw, &result.value.uval) ? TRT_DWORD 
                : TRT_GETLASTERROR;
            break;
        }
        break;
    case 15:
    case 16:
        if (TwapiGetArgs(interp, objc, objv,
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
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 17:
            dw = sizeof(u.pmce);
            u.pmce.cb = dw;
            if (! GetProcessMemoryInfo(h, (PROCESS_MEMORY_COUNTERS*) &u.pmce, sizeof(u.pmce))) {
                dw = sizeof(PROCESS_MEMORY_COUNTERS);
                TWAPI_ASSERT(sizeof(u.pmce) < dw);
                u.pmce.cb = dw;
                if (! GetProcessMemoryInfo(h, (PROCESS_MEMORY_COUNTERS*) &u.pmce, sizeof(u.pmce))) {
                    result.type = TRT_GETLASTERROR;
                    break;
                }
            }
            objs[0] = ObjFromDWORD(u.pmce.PageFaultCount);
            objs[1] = ObjFromSIZE_T(u.pmce.PeakWorkingSetSize);
            objs[2] = ObjFromSIZE_T(u.pmce.WorkingSetSize);
            objs[3] = ObjFromSIZE_T(u.pmce.QuotaPeakPagedPoolUsage);
            objs[4] = ObjFromSIZE_T(u.pmce.QuotaPagedPoolUsage);
            objs[5] = ObjFromSIZE_T(u.pmce.QuotaPeakNonPagedPoolUsage);
            objs[6] = ObjFromSIZE_T(u.pmce.QuotaNonPagedPoolUsage);
            objs[7] = ObjFromSIZE_T(u.pmce.PagefileUsage);
            objs[8] = ObjFromSIZE_T(u.pmce.PeakPagefileUsage);
            if (dw == sizeof(u.pmce))
                objs[9] = ObjFromSIZE_T(u.pmce.PrivateUsage);
            else
                objs[9] = ObjFromLong(0);
            result.value.objv.objPP = objs;
            result.value.objv.nobj = 10;
            result.type = TRT_OBJV;
            break;
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
                ? TRT_GETLASTERROR : TRT_LONG;
            break;
        case 25:
            result.type = GetExitCodeProcess(h, &result.value.uval)
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
        if (TwapiGetArgs(interp, objc, objv,
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
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), ARGSKIP,
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        if (dw & GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS) {
            /* Argument is address in a module */
            if (ObjToDWORD_PTR(interp, objv[1], &u.dwp) != TCL_OK)
                return TCL_ERROR;
        } else {
            u.dwp = (DWORD_PTR) ObjToLPWSTR_NULL_IF_EMPTY(objv[1]);
        }
        if (GetModuleHandleExW(dw, (LPCWSTR) u.dwp, &result.value.hmodule))
            result.type = TRT_HMODULE;
        else
            result.type = TRT_GETLASTERROR;
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiProcessInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s ProcessDispatch[] = {
        DEFINE_FNCODE_CMD(GetCurrentThread, 6),
        DEFINE_FNCODE_CMD(ReadProcessMemory, 7),
        DEFINE_FNCODE_CMD(GetModuleFileNameEx, 8),
        DEFINE_FNCODE_CMD(GetModuleBaseName, 9),
        DEFINE_FNCODE_CMD(GetModuleInformation, 10),
        DEFINE_FNCODE_CMD(Twapi_GetProcessList, 11),
        DEFINE_FNCODE_CMD(SetProcessWorkingSetSize, 12),
        DEFINE_FNCODE_CMD(SetThreadExecutionState, 13),
        DEFINE_FNCODE_CMD(ProcessIdToSessionId, 14),
        DEFINE_FNCODE_CMD(OpenProcess, 15),
        DEFINE_FNCODE_CMD(OpenThread, 16),
        DEFINE_FNCODE_CMD(GetProcessMemoryInfo, 17), // TBD - Tcl
        DEFINE_FNCODE_CMD(IsWow64Process, 18),
        DEFINE_FNCODE_CMD(ResumeThread, 19),
        DEFINE_FNCODE_CMD(SuspendThread, 20),
        DEFINE_FNCODE_CMD(GetPriorityClass, 21),
        DEFINE_FNCODE_CMD(Twapi_NtQueryInformationProcessBasicInformation, 22),
        DEFINE_FNCODE_CMD(Twapi_NtQueryInformationThreadBasicInformation, 23),
        DEFINE_FNCODE_CMD(GetThreadPriority, 24),
        DEFINE_FNCODE_CMD(GetExitCodeProcess, 25),
        DEFINE_FNCODE_CMD(GetProcessImageFileName, 26), /* TBD - Tcl wrapper */
        DEFINE_FNCODE_CMD(GetDeviceDriverBaseName, 27),
        DEFINE_FNCODE_CMD(GetDeviceDriverFileName, 28),
        DEFINE_FNCODE_CMD(WaitForInputIdle, 29),
        DEFINE_FNCODE_CMD(SetPriorityClass, 30),
        DEFINE_FNCODE_CMD(SetThreadPriority, 31),
        DEFINE_FNCODE_CMD(TerminateProcess, 32),
        DEFINE_FNCODE_CMD(GetModuleHandleEx, 33),
    };

    static struct alias_dispatch_s EnumDispatch[] = {
        DEFINE_ALIAS_CMD(EnumProcesses, 0),
        DEFINE_ALIAS_CMD(EnumProcessModules, 1),
        DEFINE_ALIAS_CMD(EnumDeviceDrivers, 2),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(ProcessDispatch), ProcessDispatch, Twapi_ProcessCallObjCmd);

    Tcl_CreateObjCommand(interp, "twapi::EnumProcessHelper", Twapi_EnumProcessesModulesObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CreateProcess", Twapi_CreateProcessObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CreateProcessAsUser", Twapi_CreateProcessAsUserObjCmd, ticP, NULL);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(EnumDispatch), EnumDispatch, "twapi::EnumProcessHelper");

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
int Twapi_process_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiProcessInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}
