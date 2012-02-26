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
    CALL_(GetVersionEx, Call, 2);
    CALL_(GetSystemTimeAsFileTime, Call, 40);
    CALL_(AllocateLocallyUniqueId, Call, 48);
    CALL_(LockWorkStation, Call, 49);
    CALL_(RevertToSelf, Call, 51); /* Leave here as it might be
                                      used from multiple extensions */
    CALL_(GetSystemPowerStatus, Call, 68);
    CALL_(TwapiId, Call, 74);
    CALL_(DebugBreak, Call, 75);
    CALL_(Twapi_GetNotificationWindow, Call, 77);
    CALL_(GetDefaultPrinter, Call, 82);         /* TBD Tcl */

    CALL_(Twapi_AddressToPointer, Call, 1001);
    CALL_(IsValidSid, Call, 1002);
    CALL_(VariantTimeToSystemTime, Call, 1003);
    CALL_(SystemTimeToVariantTime, Call, 1004);
    CALL_(canonicalize_guid, Call, 1005); // TBD Document
    CALL_(FileTimeToSystemTime, Call, 1016);
    CALL_(SystemTimeToFileTime, Call, 1017);
    CALL_(Twapi_IsValidGUID, Call, 1019);
    CALL_(Twapi_UnregisterWaitOnHandle, Call, 1020);
    CALL_(free, Call, 1022);

    CALL_(SystemParametersInfo, Call, 10001);
    CALL_(LookupAccountSid, Call, 10002);
    CALL_(DuplicateHandle, Call, 10008);
    CALL_(Tcl_GetChannelHandle, Call, 10009);
    CALL_(CreateFile, Call, 10031);
    CALL_(DsGetDcName, Call, 10058);
    CALL_(FormatMessageFromModule, Call, 10073);
    CALL_(FormatMessageFromString, Call, 10074);
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
    CALL_(IsEqualGUID, Call, 10136); // Tcl
    CALL_(IMofCompiler_CompileBuffer, Call, 10137); // Tcl
    CALL_(IMofCompiler_CompileFile, Call, 10138); // Tcl
    CALL_(IMofCompiler_CreateBMOF, Call, 10139); // Tcl

    // CallU API
    CALL_(GetStdHandle, CallU, 4);
    CALL_(Twapi_EnumPrinters_Level4, CallU, 20);
    CALL_(UuidCreate, CallU, 21);
    CALL_(GetUserNameEx, CallU, 22);
    CALL_(Twapi_MapWindowsErrorToString, CallU, 34);
    CALL_(Twapi_MemLifoInit, CallU, 37);
    CALL_(GlobalDeleteAtom, CallU, 38); // TBD - tcl interface

    CALL_(AttachThreadInput, CallU, 2004); /* Must stay in twapi_base as
                                              potentially used by many exts */

    CALL_(GlobalAlloc, CallU, 10001);
    CALL_(LHashValOfName, CallU, 10002);
    CALL_(SetStdHandle, CallU, 10004);

    // CallS - function(LPWSTR)
    CALL_(Twapi_AppendLog, CallS, 11);
    CALL_(GlobalAddAtom, CallS, 23); // TBD - Tcl interface
    CALL_(is_valid_sid_syntax, CallS, 27); // TBD - Tcl interface

    CALL_(LoadLibraryEx, CallS, 504);
    CALL_(TranslateName, CallS, 1005);

    // CallH - function(HANDLE)
    CALL_(GetHandleInformation, CallH, 14);
    CALL_(FreeLibrary, CallH, 15);
    CALL_(GetDevicePowerState, CallH, 16); // TBD - which module ?
    CALL_(ReleaseMutex, CallH, 42);
    CALL_(CloseHandle, CallH, 43);
    CALL_(CastToHANDLE, CallH, 44);
    CALL_(GlobalFree, CallH, 45);
    CALL_(GlobalUnlock, CallH, 46);
    CALL_(GlobalSize, CallH, 47);
    CALL_(GlobalLock, CallH, 48);
    CALL_(WTSEnumerateProcesses, CallH, 50); // Stays in base module as commonly useful
    CALL_(Twapi_MemLifoClose, CallH, 54);
    CALL_(Twapi_MemLifoPopFrame, CallH, 55);
    CALL_(Twapi_MemLifoPushMark, CallH, 60);
    CALL_(Twapi_MemLifoPopMark, CallH, 61);
    CALL_(Twapi_MemLifoValidate, CallH, 62);
    CALL_(Twapi_MemLifoDump, CallH, 63);
    CALL_(SetEvent, CallH, 66);
    CALL_(ResetEvent, CallH, 67);

    CALL_(ReleaseSemaphore, CallH, 1001);
    CALL_(WaitForSingleObject, CallH, 1017);
    CALL_(Twapi_MemLifoAlloc, CallH, 1018);
    CALL_(Twapi_MemLifoPushFrame, CallH, 1019);

    CALL_(SetHandleInformation, CallH, 2007); /* TBD - Tcl wrapper */
    CALL_(Twapi_MemLifoExpandLast, CallH, 2008);
    CALL_(Twapi_MemLifoShrinkLast, CallH, 2009);
    CALL_(Twapi_MemLifoResizeLast, CallH, 2010);
    CALL_(Twapi_RegisterWaitOnHandle, CallH, 2011);

    CALL_(GetWindowLongPtr, CallWU, 8);

    CALL_(PostMessage, CallWU, 1001);
    CALL_(SendNotifyMessage, CallWU, 1002);
    CALL_(SendMessageTimeout, CallWU, 1003);

    CALL_(SetWindowLongPtr, CallWU, 10003);

    // CallSSSD - function(LPWSTR_NULL_IF_EMPTY, LPWSTR, LPWSTR, DWORD)
    CALL_(LookupAccountName, CallSSSD, 1);
    CALL_(LogonUser, CallSSSD, 5);
    CALL_(NetGetDCName, CallSSSD, 34);

#undef CALL_

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
        SID *sidP;
    } u;
    DWORD_PTR dwp;
    DWORD dw, dw2, dw3, dw4;
    LPWSTR s, s2, s3;
    unsigned char *cP;
    void *pv, *pv2;
    Tcl_Obj *objs[2];
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h, h2, h3;
    GUID guid;
    GUID *guidP;
    SYSTEMTIME systime;

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
        case 2:
            return Twapi_GetVersionEx(interp);
        // 3-39 UNUSED
        case 40:
            result.type = TRT_FILETIME;
            GetSystemTimeAsFileTime(&result.value.filetime);
            break;
            // 41-47 UNUSED
        case 48:
            result.type = AllocateLocallyUniqueId(&result.value.luid) ? TRT_LUID : TRT_GETLASTERROR;
            break;
        case 49:
            result.value.ival = LockWorkStation();
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
            // 50
        case 51:
            result.value.ival = RevertToSelf();
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
            // 52-67 UNUSED
        case 68:
            if (GetSystemPowerStatus(&u.power_status)) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromSYSTEM_POWER_STATUS(&u.power_status);
            } else
                result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        // 69-73 UNUSED
        case 74:
            result.type = TRT_WIDE;
            result.value.wide = TWAPI_NEWID(ticP);
            break;
        case 75:
            result.type = TRT_EMPTY;
            DebugBreak();
            break;
            // 76 UNUSED
        case 77:
            result.type = TRT_HWND;
            result.value.hwin = Twapi_GetNotificationWindow(ticP);
            break;
            // 78-81 - UNUSED
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
        case 1002:
            u.sidP = NULL;
            result.type = TRT_BOOL;
            result.value.bval = (ObjToPSID(interp, objv[2], &u.sidP) == TCL_OK);
            if (u.sidP)
                TwapiFree(u.sidP);
            break;
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
            // 1006 - 1015 UNUSED
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
            // 1018
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
        case 10001:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw), GETINT(dw2), GETVOIDP(pv), GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SystemParametersInfoW(dw, dw2, pv, dw3);
            break;
        case 10002:
            u.sidP = NULL;
            result.type = TRT_TCL_RESULT;
            result.value.ival = TwapiGetArgs(interp, objc-2, objv+2,
                                             GETNULLIFEMPTY(s),
                                             GETVAR(u.sidP, ObjToPSID),
                                             ARGEND);
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_LookupAccountSid(interp, s, u.sidP);
            if (u.sidP)
                TwapiFree(u.sidP);
            break;
            
            // 10003-10007 UNUSED
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
        // 10010-30 UNUSED

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
        // 10032-57 UNUSED
        case 10058: // DsGetDcName
            guidP = &guid;
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETVAR(guidP, ObjToUUID_NULL),
                             GETNULLIFEMPTY(s3), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_DsGetDcName(interp, s, s2, guidP, s3, dw);
        // 10059 - 10072 UNUSED
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
            // 10075 - 10080 UNUSED
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
            // 10123 - 10135 UNUSED
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
            // 5-19 UNUSED
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
            // 23-33 UNUSED
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
    WCHAR *bufP;

    result.type = TRT_BADFUNCTIONCODE;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETWSTR(arg),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    /* One argument commands - assign codes 1-999 */
    if (func < 500) {
        switch (func) {
            // 1-10 UNUSED
        //9- 10 - UNUSED
        case 11:
            return Twapi_AppendLog(interp, arg);
        // 12-22 UNUSED
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
    HANDLE h;
    DWORD dw, dw2;
    TwapiResult result;
    int func;

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
        // 17-41 UNUSED

        case 42:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ReleaseMutex(h);
            break;

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
        case 50:
            return Twapi_WTSEnumerateProcesses(interp, h);
            // 51-53 UNUSED
        case 54:
            MemLifoClose(h);
            TwapiFree(h);
            result.type = TRT_EMPTY;
            break;
        case 55:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = MemLifoPopFrame((MemLifo *)h);
            break;
        // 56-59 UNUSED
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
        // 64-65 - UNUSED
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
        // 1002-1016 UNUSED
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
    // 6-33 UNUSED
    case 34:
        NULLIFY_EMPTY(s2);
        return Twapi_NetGetDCName(interp, s1,s2);
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

int Twapi_LookupAccountSid (
    Tcl_Interp *interp,
    LPCWSTR     lpSystemName,
    PSID        sidP
)
{
    WCHAR       *domainP;
    DWORD        domain_buf_size;
    SID_NAME_USE account_type;
    DWORD        error;
    int          result;
    Tcl_Obj     *objs[3];
    LPWSTR       nameP;
    DWORD        name_buf_size;
    int          i;

    for (i=0; i < (sizeof(objs)/sizeof(objs[0])); ++i)
        objs[i] = NULL;

    result = TCL_ERROR;

    domainP         = NULL;
    domain_buf_size = 0;
    nameP           = NULL;
    name_buf_size   = 0;
    error           = 0;
    if (LookupAccountSidW(lpSystemName, sidP, NULL, &name_buf_size,
                          NULL, &domain_buf_size, &account_type) == 0) {
        error = GetLastError();
    }

    if (error && (error != ERROR_INSUFFICIENT_BUFFER)) {
        Tcl_SetResult(interp, "Error looking up account SID: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, error);
        goto done;
    }

    /* Allocate required space */
    domainP = TwapiAlloc(domain_buf_size * sizeof(*domainP));
    nameP = TwapiAlloc(name_buf_size * sizeof(*nameP));

    if (LookupAccountSidW(lpSystemName, sidP, nameP, &name_buf_size,
                          domainP, &domain_buf_size, &account_type) == 0) {
        Tcl_SetResult(interp, "Error looking up account SID: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, GetLastError());
        goto done;
    }

    /*
     * Got everything we need, now format it
     * {NAME DOMAIN ACCOUNT}
     */
    objs[0] = ObjFromUnicode(nameP);   /* Will exit on alloc fail */
    objs[1] = ObjFromUnicode(domainP); /* Will exit on alloc fail */
    objs[2] = Tcl_NewIntObj(account_type);
    Tcl_SetObjResult(interp, Tcl_NewListObj(3, objs));
    result = TCL_OK;

 done:
    if (domainP)
        TwapiFree(domainP);
    if (nameP)
        TwapiFree(nameP);

    return result;
}


int Twapi_LookupAccountName (
    Tcl_Interp *interp,
    LPCWSTR     lpSystemName,
    LPCWSTR     lpAccountName
)
{
    PSID         sidP;
    DWORD        sid_buf_size;
    WCHAR       *domainP;
    DWORD        domain_buf_size;
    SID_NAME_USE account_type;
    DWORD        error;
    int          result;
    Tcl_Obj     *objs[3];
    int          i;

    /*
     * Special case check for empty string - else LookupAccountName
     * returns the same error as for insufficient buffer .
     */
    if (*lpAccountName == 0) {
        return Twapi_GenerateWin32Error(interp, ERROR_INVALID_PARAMETER, "Empty string passed for account name.");
    }

    for (i=0; i < (sizeof(objs)/sizeof(objs[0])); ++i)
        objs[i] = NULL;
    result = TCL_ERROR;


    domain_buf_size = 0;
    sid_buf_size    = 0;
    error           = 0;
    if (LookupAccountNameW(lpSystemName, lpAccountName, NULL, &sid_buf_size,
                          NULL, &domain_buf_size, &account_type) == 0) {
        error = GetLastError();
    }

    if (error && (error != ERROR_INSUFFICIENT_BUFFER)) {
        Tcl_SetResult(interp, "Error looking up account name: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, error);
        return TCL_ERROR;
    }

    /* Allocate required space */
    domainP = TwapiAlloc(domain_buf_size * sizeof(*domainP));
    sidP = TwapiAlloc(sid_buf_size);

    if (LookupAccountNameW(lpSystemName, lpAccountName, sidP, &sid_buf_size,
                          domainP, &domain_buf_size, &account_type) == 0) {
        Tcl_SetResult(interp, "Error looking up account name: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, GetLastError());
        goto done;
    }

    /*
     * There is a bug in LookupAccountName (see KB 185246) where
     * if the user name happens to be the machine name, the returned SID
     * is for the machine, not the user. As suggested there, we look
     * for this case by checking the account type returned and if we have hit
     * this case, recurse using a user name of "\\domain\\username"
     */
    if (account_type == SidTypeDomain) {
        /* Redo the operation */
        WCHAR *new_accountP;
        size_t len = 0;
        size_t sysnamelen, accnamelen;
        TWAPI_ASSERT(lpSystemName);
        TWAPI_ASSERT(lpAccountName);
        sysnamelen = lstrlenW(lpSystemName);
        accnamelen = lstrlenW(lpAccountName);
        len = sysnamelen + 1 + accnamelen + 1;
        new_accountP = TwapiAlloc(len * sizeof(*new_accountP));
        CopyMemory(new_accountP, lpSystemName, sizeof(*new_accountP)*sysnamelen);
        new_accountP[sysnamelen] = L'\\';
        CopyMemory(new_accountP+sysnamelen+1, lpAccountName, sizeof(*new_accountP)*accnamelen);
        new_accountP[sysnamelen+1+accnamelen] = 0;

        /* Recurse */
        result = Twapi_LookupAccountName(interp, lpSystemName, new_accountP);
        TwapiFree(new_accountP);
        goto done;
    }


    /*
     * Got everything we need, now format it
     * {SID DOMAIN ACCOUNT}
     */
    result = ObjFromSID(interp, sidP, &objs[0]);
    if (result != TCL_OK)
        goto done;
    objs[1] = ObjFromUnicode(domainP); /* Will exit on alloc fail */
    objs[2] = Tcl_NewIntObj(account_type);
    Tcl_SetObjResult(interp, Tcl_NewListObj(3, objs));
    result = TCL_OK;

 done:
    if (domainP)
        TwapiFree(domainP);
    if (sidP)
        TwapiFree(sidP);

    return result;
}

int Twapi_NetGetDCName(Tcl_Interp *interp, LPCWSTR servername, LPCWSTR domainname)
{
    NET_API_STATUS status;
    LPBYTE         bufP;
    status = NetGetDCName(servername, domainname, &bufP);
    if (status != NERR_Success) {
        return Twapi_AppendSystemError(interp, status);
    }
    Tcl_SetObjResult(interp, ObjFromUnicode((wchar_t *)bufP));
    NetApiBufferFree(bufP);
    return TCL_OK;
}
