/*
 * Copyright (c) 2010-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

static TCL_RESULT Twapi_LsaQueryInformationPolicy (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
    );

/* TBD - remove this call */
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
                    ptrval = ObjToByteArray(objP, &len);
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
            if (objP && ObjToInt(interp, objP, &ival) != TCL_OK)
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
                *(char **)p = objP ? ObjToString(objP) : "";
            break;
        case ARGASTRN: // char string and its length
            lenP = va_arg(ap, int *);
            sval = "";
            len = 0;
            if (objP)
                sval = ObjToStringN(objP, &len);
            if (p)
                *(char **)p = sval;
            if (lenP)
                *lenP = len;
            break;
        case ARGWSTR: // Unicode string
            if (p) {
                *(WCHAR **)p = objP ? ObjToUnicode(objP) : L"" ;
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
                uval = ObjToUnicodeN(objP, &len);
            if (p)
                *(WCHAR **)p = uval;
            if (lenP)
                *lenP = len;
            break;
        case ARGWORD: // WORD - 16 bits
            ival = 0;
            if (objP && ObjToInt(interp, objP, &ival) != TCL_OK)
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
                TwapiSetStaticResult(interp, "Default values invalid used for ARGVAR types.");
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
            TwapiSetStaticResult(interp, "TwapiGetArgs: unexpected format character.");
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
        TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        goto argerror;
    }

    va_end(ap);
    return TCL_OK;

argerror:
    /* interp is already supposed to contain an error message */
    va_end(ap);
    return TCL_ERROR;
}


static TCL_RESULT Twapi_CallNoargsObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;

   union {
        WCHAR buf[MAX_PATH+1];
        SYSTEM_POWER_STATUS power_status;
    } u;
   int func = PtrToInt(clientdata);
    
    if (objc != 1)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.type = TRT_HANDLE;
        result.value.hval = GetCurrentProcess();
        break;
    case 2:
        return Twapi_GetVersionEx(interp);
    case 3: // UuidCreateNil
        TwapiSetObjResult(interp, STRING_LITERAL_OBJ("00000000-0000-0000-0000-000000000000"));
        return TCL_OK;
    case 4: // Twapi_GetInstallDir
        result.value.obj = TwapiGetInstallDir(interp, NULL);
        if (result.value.obj == NULL)
            return TCL_ERROR; /* interp error result already set */
        result.type = TRT_OBJ;
        break;
    case 5:
        return Twapi_EnumWindows(interp);
    case 6:                /* GetSystemWindowsDirectory */
    case 7:                /* GetWindowsDirectory */
    case 8:                /* GetSystemDirectory */
        result.type = TRT_UNICODE;
        result.value.unicode.str = u.buf;
        result.value.unicode.len =
            (func == 6
             ? GetSystemWindowsDirectoryW
             : (func == 7 ? GetWindowsDirectoryW : GetSystemDirectoryW)
                ) (u.buf, ARRAYSIZE(u.buf));
        if (result.value.unicode.len >= ARRAYSIZE(u.buf) ||
            result.value.unicode.len == 0) {
            result.type = TRT_GETLASTERROR;
        }
        break;
    case 9:
        result.type = TRT_DWORD;
        result.value.uval = GetCurrentThreadId();
        break;
    case 10:
        result.type = TRT_DWORD;
        result.value.uval = GetTickCount();
        break;
    case 11:
        result.type = TRT_FILETIME;
        GetSystemTimeAsFileTime(&result.value.filetime);
        break;
    case 12:
        result.type = AllocateLocallyUniqueId(&result.value.luid) ? TRT_LUID : TRT_GETLASTERROR;
        break;
    case 13:
        result.value.ival = LockWorkStation();
        result.type = TRT_EXCEPTION_ON_FALSE;
        break;
    case 14:
        result.value.ival = RevertToSelf();
        result.type = TRT_EXCEPTION_ON_FALSE;
        break;
    case 15:
        if (GetSystemPowerStatus(&u.power_status)) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromSYSTEM_POWER_STATUS(&u.power_status);
        } else
            result.type = TRT_EXCEPTION_ON_FALSE;
        break;
    case 16:
        result.type = TRT_EMPTY;
        DebugBreak();
        break;
    case 17:
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

static int Twapi_CallIntArgObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func = PtrToInt(clientdata);
    DWORD dw;
    TwapiResult result;
    union {
        RPC_STATUS rpc_status;
        MemLifo *lifoP;
        WCHAR buf[MAX_PATH+1];
    } u;

    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, dw, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.value.hval = GetStdHandle(dw);
        if (result.value.hval == INVALID_HANDLE_VALUE)
            result.type = TRT_GETLASTERROR;
        else if (result.value.hval == NULL) {
            result.value.ival = ERROR_FILE_NOT_FOUND;
            result.type = TRT_EXCEPTION_ON_ERROR;
        } else
            result.type = TRT_HANDLE;
        break;
    case 2:
        u.rpc_status = UuidCreate(&result.value.uuid);
        /* dw boolean indicating whether to allow strictly local uuids */
        if ((u.rpc_status == RPC_S_UUID_LOCAL_ONLY) && dw) {
            /* If caller does not mind a local only uuid, don't return error */
            u.rpc_status = RPC_S_OK;
        }
        result.type = u.rpc_status == RPC_S_OK ? TRT_UUID : TRT_GETLASTERROR;
        break;
    case 3:
        result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
        if (GetUserNameExW(dw, u.buf, &result.value.unicode.len)) {
            result.value.unicode.str = u.buf;
            result.type = TRT_UNICODE;
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 4:
        result.value.obj = Twapi_MapWindowsErrorToString(dw);
        result.type = TRT_OBJ;
        break;
    case 5:
        u.lifoP = TwapiAlloc(sizeof(MemLifo));
        result.value.ival = MemLifoInit(u.lifoP, NULL, NULL, NULL, dw, 0);
        if (result.value.ival == ERROR_SUCCESS) {
            result.type = TRT_NONNULL;
            result.value.nonnull.p = u.lifoP;
            result.value.nonnull.name = "MemLifo*";
        } else
            result.type = TRT_EXCEPTION_ON_ERROR;
        break;
    case 6:
        SetLastError(0);    /* As per MSDN */
        GlobalDeleteAtom((WORD)dw);
        result.value.ival = GetLastError();
        result.type = TRT_EXCEPTION_ON_ERROR;
        break;
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_CallOneArgObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func = PtrToInt(clientdata);
    TwapiResult result;
    union {
        WCHAR buf[MAX_PATH+1];
        LPCWSTR wargv[100];     /* FormatMessage accepts up to 99 params + 1 for NULL */
        double d;
        FILETIME   filetime;
        TwapiId twapi_id;
        GUID guid;
        SID *sidP;
        MemLifo *lifoP;
        struct {
            HWND hwnd;
            HWND hwnd2;
        };
        LSA_OBJECT_ATTRIBUTES lsa_oattr;
    } u;
    DWORD dw, dw2;
    DWORD_PTR dwp;
    LPWSTR s;
    void *pv;
    Tcl_Obj *objs[2];
    GUID guid;
    WCHAR *bufP;
    SYSTEMTIME systime;

    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    --objc;
    ++objv;;
    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1008:
        if (ObjToDWORD_PTR(interp, objv[0], &dwp) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_LPVOID;
        result.value.pv = (void *) dwp;
        break;
    case 1009:
        u.sidP = NULL;
        result.type = TRT_BOOL;
        result.value.bval = (ObjToPSID(interp, objv[0], &u.sidP) == TCL_OK);
        if (u.sidP)
            TwapiFree(u.sidP);
        break;
    case 1010:
        if (Tcl_GetDoubleFromObj(interp, objv[0], &u.d) != TCL_OK)
            return TCL_ERROR;
        result.type = VariantTimeToSystemTime(u.d, &result.value.systime) ?
            TRT_SYSTEMTIME : TRT_GETLASTERROR;
        break;
    case 1011:
        if (ObjToSYSTEMTIME(interp, objv[0], &systime) != TCL_OK)
            return TCL_ERROR;
        result.type = SystemTimeToVariantTime(&systime, &result.value.dval) ?
            TRT_DOUBLE : TRT_GETLASTERROR;
        break;
    case 1012: // canonicalize_guid
        /* Turn a GUID into canonical form */
        if (ObjToGUID(interp, objv[0], &result.value.guid) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_GUID;
        break;
    case 1013:
        return Twapi_AppendLog(interp, ObjToUnicode(objv[0]));
    case 1014: // GlobalAddAtom
        result.value.ival = GlobalAddAtomW(ObjToUnicode(objv[0]));
        result.type = result.value.ival ? TRT_LONG : TRT_GETLASTERROR;
        break;
    case 1015:
        u.sidP = NULL;
        result.type = TRT_BOOL;
        result.value.bval = ConvertStringSidToSidW(ObjToUnicode(objv[0]),
                                                   &u.sidP);
        if (u.sidP)
            LocalFree(u.sidP);
        break;
    case 1016:
        if (ObjToFILETIME(interp, objv[0], &u.filetime) != TCL_OK)
            return TCL_ERROR;
        if (FileTimeToSystemTime(&u.filetime, &result.value.systime))
            result.type = TRT_SYSTEMTIME;
        else
            result.type = TRT_GETLASTERROR;
        break;
    case 1017:
        if (ObjToSYSTEMTIME(interp, objv[0], &systime) != TCL_OK)
            return TCL_ERROR;
        if (SystemTimeToFileTime(&systime, &result.value.filetime))
            result.type = TRT_FILETIME;
        else
            result.type = TRT_GETLASTERROR;
        break;
    case 1018: /* In twapi_base because needed in multiple modules */
        if (ObjToHWND(interp, objv[0], &u.hwnd) != TCL_OK)
            return TCL_ERROR;
        dw2 = GetWindowThreadProcessId(u.hwnd, &dw);
        if (dw2 == 0) {
            result.type = TRT_GETLASTERROR;
        } else {
            objs[0] = ObjFromDWORD(dw2);
            objs[1] = ObjFromDWORD(dw);
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            result.type = TRT_OBJV;
        }
        break;
    case 1019: // Twapi_IsValidGUID
        result.type = TRT_BOOL;
        result.value.bval = (ObjToGUID(NULL, objv[0], &guid) == TCL_OK);
        break;
    case 1020:
        bufP = u.buf;
        s = ObjToUnicode(objv[0]);
        dw = ExpandEnvironmentStringsW(s, bufP, ARRAYSIZE(u.buf));
        if (dw > ARRAYSIZE(u.buf)) {
            // Need a bigger buffer
            bufP = TwapiAlloc(dw * sizeof(WCHAR));
            dw2 = dw;
            dw = ExpandEnvironmentStringsW(s, bufP, dw2);
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
    case 1021: // free
        if (ObjToLPVOID(interp, objv[0], &pv) != TCL_OK ||
            TwapiUnregisterPointer(interp, pv, TwapiAlloc) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EMPTY;
        if (pv)
            TwapiFree(pv);
        break;
    case 1022: // registered_pointer
        if (ObjToLPVOID(interp, objv[0], &pv) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.bval = TwapiVerifyPointer(interp, pv, NULL);
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_CallArgsObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func = PtrToInt(clientdata);
    union {
        WCHAR buf[MAX_PATH+1];
        LPCWSTR wargv[100];     /* FormatMessage accepts up to 99 params + 1 for NULL */
        GUID guid;
        SID *sidP;
        struct {
            HWND hwnd;
            HWND hwnd2;
        };
        LSA_OBJECT_ATTRIBUTES lsa_oattr;
    } u;
    DWORD dw, dw2, dw3, dw4;
    DWORD_PTR dwp, dwp2;
    LPWSTR s, s2, s3;
    unsigned char *cP;
    void *pv, *pv2;
    Tcl_Obj *objs[2];
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h, h2, h3;
    GUID guid;
    GUID *guidP;
    LSA_UNICODE_STRING lsa_ustr; /* Used with lsa_oattr so not in union */
    TwapiResult result;

    --objc;
    ++objv;
    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 10001:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETINT(dw2), GETVOIDP(pv), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SystemParametersInfoW(dw, dw2, pv, dw3);
        break;
    case 10002:
        u.sidP = NULL;
        result.type = TRT_TCL_RESULT;
        result.value.ival = TwapiGetArgs(interp, objc, objv,
                                         GETNULLIFEMPTY(s),
                                         GETVAR(u.sidP, ObjToPSID),
                                         ARGEND);
        result.type = TRT_TCL_RESULT;
        result.value.ival = Twapi_LookupAccountSid(interp, s, u.sidP);
        if (u.sidP)
            TwapiFree(u.sidP);
        break;
    case 10003:
    case 10004:
        if (TwapiGetArgs(interp, objc, objv,
                         GETNULLIFEMPTY(s), GETWSTR(s2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (func == 10003)
            return Twapi_LookupAccountName(interp, s, s2);
        else {
            NULLIFY_EMPTY(s2);
            return Twapi_NetGetDCName(interp, s, s2);
        }
        break;
    case 10005:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.bval = AttachThreadInput(dw, dw2, dw3);
        break;
    case 10006:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETVAR(dwp, ObjToDWORD_PTR),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HGLOBAL;
        result.value.hval = GlobalAlloc(dw, dwp);
        break;
    case 10007:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETWSTR(s),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_DWORD;
        result.value.uval = LHashValOfName(dw, s);
        break;
    case 10008:
        if (TwapiGetArgs(interp, objc, objv,
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
        return Twapi_TclGetChannelHandle(interp, objc, objv);
    case 10010:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETHANDLE(h),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SetStdHandle(dw, h);
        break;
    case 10011:
        if (TwapiGetArgs(interp, objc, objv,
                         GETWSTR(s), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HANDLE;
        result.value.hval = LoadLibraryExW(s, NULL, dw);
        break;
    case 10012: // CreateFile
        secattrP = NULL;
        if (TwapiGetArgs(interp, objc, objv,
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
    case 10013:
        if (TwapiGetArgs(interp, objc, objv,
                         GETASTR(cP), ARGUSEDEFAULT,
                         GETINT(dw), GETHANDLE(h), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (h == NULL)
            h = gTwapiModuleHandle;
        return Twapi_SourceResource(interp, h, cP, dw);
    case 10014: // FindWindowEx in twapi_base because of common use
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(u.hwnd, HWND),
                         GETHANDLET(u.hwnd2, HWND),
                         GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HWND;
        result.value.hval = FindWindowExW(u.hwnd, u.hwnd2, s, s2);
        break;
    case 10015:
        return Twapi_LsaQueryInformationPolicy(interp, objc, objv);
    case 10016: // LsaOpenPolicy
        if (TwapiGetArgs(interp, objc, objv, ARGSKIP, GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        ObjToLSA_UNICODE_STRING(objv[0], &lsa_ustr);
        TwapiZeroMemory(&u.lsa_oattr, sizeof(u.lsa_oattr));
        dw2 = LsaOpenPolicy(&lsa_ustr, &u.lsa_oattr, dw, &result.value.hval);
        if (dw2 == STATUS_SUCCESS) {
            result.type = TRT_LSA_HANDLE;
        } else {
            result.type = TRT_NTSTATUS;
            result.value.ival = dw2;
        }
        break;
            
    case 10017:
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(u.hwnd, HWND), GETINT(dw),
                         ARGTERM) != TCL_OK)
            return TCL_ERROR;

        SetLastError(0);    /* Avoid spurious errors when checking GetLastError */
        result.value.dwp = GetWindowLongPtrW(u.hwnd, dw);
        if (result.value.dwp || GetLastError() == 0)
            result.type = TRT_DWORD_PTR;
        else
            result.type = TRT_GETLASTERROR;
        break;
    case 10018:
    case 10019:
    case 10020:
        // HWIN UINT WPARAM LPARAM ?ARGS?
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(u.hwnd, HWND), GETINT(dw),
                         GETDWORD_PTR(dwp), GETDWORD_PTR(dwp2),
                         ARGUSEDEFAULT,
                         GETINT(dw2), GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 10018:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = PostMessageW(u.hwnd, dw, dwp, dwp2);
            break;
        case 10019:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SendNotifyMessageW(u.hwnd, dw, dwp, dwp2);
            break;
        case 10020:
            if (SendMessageTimeoutW(u.hwnd, dw, dwp, dwp2, dw2, dw3, &result.value.dwp))
                result.type = TRT_DWORD_PTR;
            else {
                /* On some systems, GetLastError() returns 0 on timeout */
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = GetLastError();
                if (result.value.ival == 0)
                    result.value.ival = ERROR_TIMEOUT;
            }
            break;
        }
        break;
    case 10021:
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(u.hwnd, HWND), GETINT(dw),
                         GETDWORD_PTR(dwp),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = Twapi_SetWindowLongPtr(u.hwnd, dw, dwp, &result.value.dwp)
            ? TRT_DWORD_PTR : TRT_GETLASTERROR;
        break;
    case 10022: // DsGetDcName
        guidP = &guid;
        if (TwapiGetArgs(interp, objc, objv,
                         GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                         GETVAR(guidP, ObjToUUID_NULL),
                         GETNULLIFEMPTY(s3), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_DsGetDcName(interp, s, s2, guidP, s3, dw);
    case 10023: // Twapi_FormatMessageFromModule
        if (TwapiGetArgs(interp, objc, objv,
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
    case 10024: // Twapi_FormatMessageFromString
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETWSTR(s),
                         GETWARGV(u.wargv, ARRAYSIZE(u.wargv), dw4),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        /* Only look at select bits from dwFlags as others are used when
           formatting from module */
        dw &= FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK | FORMAT_MESSAGE_FROM_STRING;
        dw |=  FORMAT_MESSAGE_ARGUMENT_ARRAY;
        return TwapiFormatMessageHelper(interp, dw, s, 0, 0, dw4, u.wargv);
    case 10025:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), ARGUSEDEFAULT, GETASTR(cP),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(cP);
        return Twapi_GenerateWin32Error(interp, dw, cP);
    case 10026:
        secattrP = NULL;        /* Even on error, it might be filled */
        if (TwapiGetArgs(interp, objc, objv,
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
                    objs[1] = ObjFromBoolean(GetLastError() == ERROR_ALREADY_EXISTS);
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
    case 10027: // OpenMutex
    case 10028: // OpenSemaphore
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETBOOL(dw2), GETWSTR(s),
                         ARGUSEDEFAULT, GETINT(dw3),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HANDLE;
        result.value.hval = (func == 10027 ? OpenMutexW : OpenSemaphoreW)
            (dw, dw2, s);
        if (result.value.hval) {
            if (dw3 & 1) {
                /* Caller also wants indicator of whether object
                   already existed */
                objs[0] = ObjFromHANDLE(result.value.hval);
                objs[1] = ObjFromBoolean(GetLastError() == ERROR_ALREADY_EXISTS);
                result.value.objv.objPP = objs;
                result.value.objv.nobj = 2;
                result.type = TRT_OBJV;
            }
        }
        break;
    case 10029:
        secattrP = NULL;        /* Even on error, it might be filled */
        if (TwapiGetArgs(interp, objc, objv,
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
    case 10030: // malloc
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), ARGUSEDEFAULT, GETASTR(cP),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.nonnull.p = TwapiAlloc(dw);
        if (! TwapiRegisterPointer(interp, result.value.nonnull.p, TwapiAlloc)) {
            Tcl_Panic("Failed to register pointer");
        }
        result.value.nonnull.name = cP[0] ? cP : "void*";
        result.type = TRT_NONNULL;
        break;
    case 10031: // IsEqualPtr
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToOpaque(interp, objv[0], &pv, NULL) != TCL_OK ||
            ObjToOpaque(interp, objv[1], &pv2, NULL) != TCL_OK) {
            return TCL_ERROR;
        }
        result.type = TRT_BOOL;
        result.value.bval = (pv == pv2);
        break;
    case 10032: // IsNullPtr
        if (objc == 0 || objc > 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        cP = NULL;
        if (objc == 2) {
            cP = ObjToString(objv[1]);
            NULLIFY_EMPTY(cP);
        }
        if (ObjToOpaque(interp, objv[0], &pv, cP) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.bval = (pv == NULL);
        break;
    case 10033: // IsPtr
        if (objc == 0 || objc > 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        cP = NULL;
        if (objc == 2) {
            cP = ObjToString(objv[1]);
            NULLIFY_EMPTY(cP);
        }
        result.type = TRT_BOOL;
        result.value.bval = (ObjToOpaque(interp, objv[0], &pv, cP) == TCL_OK);
        break;
    case 10034:
        secattrP = NULL;        /* Even on error, it might be filled */
        if (TwapiGetArgs(interp, objc, objv,
                         GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                         GETBOOL(dw), GETBOOL(dw2),
                         GETNULLIFEMPTY(s),
                         ARGEND) == TCL_OK) {
            h = CreateEventW(secattrP, dw, dw2, s);
            if (h) {
                objs[1] = ObjFromBoolean(GetLastError() == ERROR_ALREADY_EXISTS); /* Do this before any other call */
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
    case 10035: // IsEqualGuid
        if (TwapiGetArgs(interp, objc, objv,
                         GETGUID(guid), GETGUID(u.guid), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.bval = IsEqualGUID(&guid, &u.guid);
        break;
    case 10036: // OpenEventLog
    case 10037: // RegisterEventSource
        if (TwapiGetArgs(interp, objc, objv,
                         GETNULLIFEMPTY(s), GETWSTR(s2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HANDLE;
        result.value.hval = (func == 10036 ?
                             OpenEventLogW : RegisterEventSourceW)(s, s2);
        break;
    }

    return TwapiSetResult(interp, &result);
}

/* This was the original dispatcher. Most stuff has now been broken out
   from here. This now handles commands that need a TwapiInterpContext
*/
static int Twapi_CallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    WCHAR buf[MAX_PATH+1];
    TwapiId twapi_id;
    DWORD dw, dw2, dw3;
    LPWSTR s;
    HANDLE h;
    WCHAR *bufP;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    objc -= 2;
    objv += 2;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.type = TRT_OBJ;
        result.value.obj = Twapi_GetAtomStats(ticP);
        break;
    case 2:
        result.type = TRT_WIDE;
        result.value.wide = TWAPI_NEWID(ticP);
        break;
    case 3:
        result.type = TRT_HWND;
        result.value.hwin = Twapi_GetNotificationWindow(ticP);
        break;
    case 4:
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToTwapiId(interp, objv[0], &twapi_id) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EMPTY;
        TwapiThreadPoolUnregister(ticP, twapi_id);
        break;
    case 5: // atomize
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        result.type = TRT_OBJ;
        result.value.obj = TwapiGetAtom(ticP, ObjToString(objv[0]));
        break;
    case 6:
        if (TwapiGetArgs(interp, objc, objv,
                         GETWSTR(s), GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        bufP = buf;
        dw3 = ARRAYSIZE(buf);
        if (! TranslateNameW(s, dw, dw2, bufP, &dw3)) {
            result.value.ival = GetLastError();
            if (result.value.ival != ERROR_INSUFFICIENT_BUFFER) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = GetLastError();
                break;
            }
            /* Retry with larger buffer */
            bufP = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR)*dw3,
                                    &dw3);
            dw3 /= sizeof(WCHAR);
            if (! TranslateNameW(s, dw, dw2, bufP, &dw3)) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = GetLastError();
                MemLifoPopFrame(&ticP->memlifo);
                break;
            }
        }
        result.value.obj = ObjFromUnicodeN(bufP, dw3-1);
        result.type = TRT_OBJ;
        if (bufP != buf)
            MemLifoPopFrame(&ticP->memlifo);
        break;
    case 7:
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h),
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return TwapiThreadPoolRegister(
            ticP, h, dw, dw2, TwapiCallRegisteredWaitScript, NULL);
    }

    return TwapiSetResult(interp, &result);
}

static int Twapi_CallHObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE h;
    DWORD dw, dw2;
    TwapiResult result;
    int func = PtrToInt(clientdata);

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHANDLE(h),
                     ARGTERM) != TCL_OK) {
        return TCL_ERROR;
    }

    --objc;
    ++objv;
    result.type = TRT_BADFUNCTIONCODE;
    if (func < 1000) {
        /* Command with a single handle argument */
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1:
            return Twapi_WTSEnumerateProcesses(interp, h);
        case 2:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ReleaseMutex(h);
            break;
        case 3:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseHandle(h);
            break;
        case 4:
            result.type = TRT_HANDLE;
            result.value.hval = h;
            break;
        case 5:
            result.type = TRT_EXCEPTION_ON_ERROR;
            /* GlobalFree will return a HANDLE on failure. */
            result.value.ival = GlobalFree(h) ? GetLastError() : 0;
            break;
        case 6:
            result.type = TRT_EXCEPTION_ON_ERROR;
            /* GlobalUnlock is an error if it returns 0 AND GetLastError is non-0 */
            result.value.ival = GlobalUnlock(h) ? 0 : GetLastError();
            break;
        case 7:
            result.type = TRT_DWORD_PTR;
            result.value.dwp = GlobalSize(h);
            break;
        case 8:
            result.type = TRT_NONNULL_LPVOID;
            result.value.pv = GlobalLock(h);
            break;
        case 9:
            MemLifoClose(h);
            TwapiFree(h);
            result.type = TRT_EMPTY;
            break;
        case 10:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = MemLifoPopFrame((MemLifo *)h);
            break;
        case 11:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetEvent(h);
            break;
        case 12:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ResetEvent(h);
            break;
        case 13:
            result.value.uval = LsaClose(h);
            result.type = TRT_DWORD; /* Not TRT_NTSTATUS because do not
                                        want error on invalid handle */
            break;
        case 14:
            result.type = GetHandleInformation(h, &result.value.ival)
                ? TRT_LONG : TRT_GETLASTERROR;
            break;
        case 15:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FreeLibrary(h);
            break;
        case 16:
            result.type = GetDevicePowerState(h, &result.value.bval)
                ? TRT_BOOL : TRT_GETLASTERROR;
            break;
        case 17:
            result.type = TRT_NONNULL;
            result.value.nonnull.p = MemLifoPushMark(h);
            result.value.nonnull.name = "MemLifoMark*";
            break;
        case 18:
            result.type = TRT_LONG;
            result.value.ival = MemLifoPopMark(h);
            break;
        case 19:
            result.type = TRT_LONG;
            result.value.ival = MemLifoValidate(h);
            break;
        case 20:
            return Twapi_MemLifoDump(interp, h);
        case 21:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CloseEventLog(h);
            break;
        case 22:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeregisterEventSource(h);
            break;
        case 23:              /* DeleteObject */
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeleteObject(h);
            break;

        }
    } else if (func < 2000) {

        // A single additional DWORD arg is present
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[1]);

        switch (func) {
        case 1001:
            result.type = ReleaseSemaphore(h, dw, &result.value.uval) ?
                TRT_LONG : TRT_GETLASTERROR;
            break;
        case 1002:
            result.value.ival = WaitForSingleObject(h, dw);
            if (result.value.ival == (DWORD) -1) {
                result.type = TRT_GETLASTERROR;
            } else {
                result.type = TRT_DWORD;
            }
            break;
        case 1003:
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoAlloc(h, dw, NULL);
            break;
        case 1004:
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoPushFrame(h, dw, NULL);
            break;
        }
    } else if (func < 3000) {

        // Two additional DWORD args present
        if (TwapiGetArgs(interp, objc-1, objv+1,
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 2001:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetHandleInformation(h, dw, dw2);
            break;
        case 2002: 
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoExpandLast(h, dw, dw2);
            break;
        case 2003: 
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoShrinkLast(h, dw, dw2);
            break;
        case 2004: 
            result.type = TRT_LPVOID;
            result.value.pv = MemLifoResizeLast(h, dw, dw2);
            break;
        }
    }
    return TwapiSetResult(interp, &result);
}


static int Twapi_ReportEventObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE hEventLog;
    WORD   wType;
    WORD  wCategory;
    DWORD dwEventID;
    PSID  lpUserSid = NULL;
    int   datalen;
    void *data;
    int     argc;
    LPCWSTR argv[32];
    int   status;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHANDLE(hEventLog), GETWORD(wType), GETWORD(wCategory),
                     GETINT(dwEventID), GETVAR(lpUserSid, ObjToPSID),
                     GETWARGV(argv, ARRAYSIZE(argv), argc),
                     GETBIN(data, datalen),
                     ARGEND) != TCL_OK) {
        if (lpUserSid)
            TwapiFree(lpUserSid);
        return TCL_ERROR;
    }

    if (datalen == 0)
        data = NULL;

    status = ReportEventW(hEventLog, wType, wCategory, dwEventID, lpUserSid,
                          (WORD) argc, datalen, argv, data);
    if (lpUserSid)
        TwapiFree(lpUserSid);

    return status ? TCL_OK : TwapiReturnSystemError(interp);
}

static TCL_RESULT Twapi_ReadMemoryObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    char *p;
    WCHAR *wp;
    int modifier;
    int type;
    int offset;
    int len;
    Tcl_Obj *objP;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(type), GETVOIDP(p), GETINT(offset),
                     ARGUSEDEFAULT, GETINT(len), GETINT(modifier),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    p += offset;
    switch (type) {
    case 0:                     /* Int */
        objP = ObjFromLong(*(int UNALIGNED *) p);
        break;

    case 1:                     /* binary */
        objP = ObjFromByteArray(p, len);
        break;

    case 2:                     /* chars */
        /* When len > 0, modifier != 0 means also check null terminator */
        if (len > 0 && modifier) {
            for (modifier = 0; modifier < len && p[modifier]; ++modifier)
                ;
            len = modifier;
        }
        objP = ObjFromStringN(p, len);
        break;
        
    case 3:                     /* wide chars */
        if (len > 0)
            len /= sizeof(WCHAR); /* Convert to num chars */
        wp = (WCHAR *) p;

        /* When len > 0, modifier != 0 means also check null terminator */
        if (len > 0 && modifier) {
            for (modifier = 0; modifier < len && wp[modifier]; ++modifier)
                ;
            len = modifier;
        }
        objP = ObjFromUnicodeN(wp, len);
        break;
        
    case 4:                     /* pointer */
        objP = ObjFromLPVOID(*(void* *) p);
        break;

    case 5:                     /* int64 */
        objP = ObjFromWideInt(*(Tcl_WideInt UNALIGNED *)p);
        break;

    default:
        TwapiSetStaticResult(interp, "Unknown type.");
        return TCL_ERROR;
    }

    return TwapiSetObjResult(interp, objP);
}


static TCL_RESULT Twapi_WriteMemoryObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    char *bufP;
    DWORD_PTR offset, buf_size;
    int func;
    int sz;
    int val;
    char *cp;
    WCHAR *wp;
    void  *pv;
    Tcl_WideInt wide;
    Tcl_Obj *dataObj;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func), GETVOIDP(bufP), GETDWORD_PTR(offset),
                     GETDWORD_PTR(buf_size), ARGSKIP,
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    dataObj = objv[5];

    /* Note the ARGSKIP ensures the objv[] in that position exists */
    switch (func) {
    case 0: // Int
        if ((offset + sizeof(int)) > buf_size)
            goto overrun;
        if (ObjToInt(interp, dataObj, &val) != TCL_OK)
            return TCL_ERROR;
        *(int UNALIGNED *)(offset + bufP) = val;
        break;
    case 1: // Binary
        cp = ObjToByteArray(dataObj, &sz);
        if ((offset + sz) > buf_size)
            goto overrun;
        CopyMemory(offset + bufP, cp, sz);
        break;
    case 2: // Chars
        cp = ObjToStringN(dataObj, &sz);
        /* Note we also include the terminating null */
        if ((offset + sz + 1) > buf_size)
            goto overrun;
        CopyMemory(offset + bufP, cp, sz+1);
        break;
    case 3: // Unicode
        wp = ObjToUnicodeN(dataObj, &sz);
        /* Note we also include the terminating null */
        if ((offset + (sizeof(WCHAR)*(sz + 1))) > buf_size)
            goto overrun;
        CopyMemory(offset + bufP, wp, sizeof(WCHAR)*(sz+1));
        break;
    case 4: // Pointer
        if ((offset + sizeof(void*)) > buf_size)
            goto overrun;
        if (ObjToLPVOID(interp, dataObj, &pv) != TCL_OK)
            return TCL_ERROR;
        CopyMemory(offset + bufP, &pv, sizeof(void*));
        break;
    case 5:
        if ((offset + sizeof(Tcl_WideInt)) > buf_size)
            goto overrun;
        if (Tcl_GetWideIntFromObj(interp, dataObj, &wide) != TCL_OK)
            return TCL_ERROR;
        *(Tcl_WideInt UNALIGNED *)(offset + bufP) = wide;
        break;
    default:
        TwapiSetStaticResult(interp, "Unknown type.");
        return TCL_ERROR;
    }        

    return TCL_OK;

overrun:
    return TwapiReturnError(interp, TWAPI_BUFFER_OVERRUN);
}

void TwapiDefineFncodeCmds(Tcl_Interp *interp, int count,
                                        struct fncode_dispatch_s *tabP, TwapiTclObjCmd *cmdfn)
{
    Tcl_DString ds;
    
    Tcl_DStringInit(&ds);
    Tcl_DStringAppend(&ds, "twapi::", ARRAYSIZE("twapi::")-1);

    while (count--) {
        Tcl_DStringSetLength(&ds, ARRAYSIZE("twapi::")-1);
        Tcl_DStringAppend(&ds, tabP->command_name, -1);
        Tcl_CreateObjCommand(interp, Tcl_DStringValue(&ds), cmdfn, IntToPtr(tabP->fncode), NULL);
        ++tabP;
    }
}


void TwapiDefineTclCmds(Tcl_Interp *interp, int count,
                        struct tcl_dispatch_s *tabP,
                        ClientData clientdata
    )
{
    Tcl_DString ds;
    
    Tcl_DStringInit(&ds);
    Tcl_DStringAppend(&ds, "twapi::", ARRAYSIZE("twapi::")-1);

    while (count--) {
        Tcl_DStringSetLength(&ds, ARRAYSIZE("twapi::")-1);
        Tcl_DStringAppend(&ds, tabP->command_name, -1);
        Tcl_CreateObjCommand(interp, Tcl_DStringValue(&ds), tabP->command_ptr, clientdata, NULL);
        ++tabP;
    }
}

void TwapiDefineAliasCmds(Tcl_Interp *interp, int count, struct alias_dispatch_s *tabP, const char *aliascmd)
{
    Tcl_DString ds;
    
    Tcl_DStringInit(&ds);
    Tcl_DStringAppend(&ds, "twapi::", ARRAYSIZE("twapi::")-1);

    while (count--) {
        Tcl_DStringSetLength(&ds, ARRAYSIZE("twapi::")-1);
        Tcl_DStringAppend(&ds, tabP->command_name, -1);
        Tcl_CreateAlias(interp, Tcl_DStringValue(&ds), interp, aliascmd, 1, &tabP->fncode);
        ++tabP;
    }
}

int Twapi_InitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s CallHDispatch[] = {
        DEFINE_FNCODE_CMD(WTSEnumerateProcesses, 1), // Kepp in base as commonly useful
        DEFINE_FNCODE_CMD(ReleaseMutex, 2),
        DEFINE_FNCODE_CMD(CloseHandle, 3),
        DEFINE_FNCODE_CMD(CastToHANDLE, 4),
        DEFINE_FNCODE_CMD(GlobalFree, 5),
        DEFINE_FNCODE_CMD(GlobalUnlock, 6),
        DEFINE_FNCODE_CMD(GlobalSize, 7),
        DEFINE_FNCODE_CMD(GlobalLock, 8),
        DEFINE_FNCODE_CMD(Twapi_MemLifoClose, 9),
        DEFINE_FNCODE_CMD(Twapi_MemLifoPopFrame, 10),
        DEFINE_FNCODE_CMD(SetEvent, 11),
        DEFINE_FNCODE_CMD(ResetEvent, 12),
        DEFINE_FNCODE_CMD(LsaClose, 13),
        DEFINE_FNCODE_CMD(GetHandleInformation, 14),
        DEFINE_FNCODE_CMD(FreeLibrary, 15),
        DEFINE_FNCODE_CMD(GetDevicePowerState, 16), // TBD - which module ?
        DEFINE_FNCODE_CMD(Twapi_MemLifoPushMark, 17),
        DEFINE_FNCODE_CMD(Twapi_MemLifoPopMark, 18),
        DEFINE_FNCODE_CMD(Twapi_MemLifoValidate, 19),
        DEFINE_FNCODE_CMD(Twapi_MemLifoDump, 20),
        DEFINE_FNCODE_CMD(CloseEventLog, 21),
        DEFINE_FNCODE_CMD(DeregisterEventSource, 22),
        DEFINE_FNCODE_CMD(DeleteObject, 23),

        DEFINE_FNCODE_CMD(ReleaseSemaphore, 1001),
        DEFINE_FNCODE_CMD(WaitForSingleObject, 1002),
        DEFINE_FNCODE_CMD(Twapi_MemLifoAlloc, 1003),
        DEFINE_FNCODE_CMD(Twapi_MemLifoPushFrame, 1004),

        DEFINE_FNCODE_CMD(SetHandleInformation, 2001), /* TBD - Tcl wrapper */
        DEFINE_FNCODE_CMD(Twapi_MemLifoExpandLast, 2002),
        DEFINE_FNCODE_CMD(Twapi_MemLifoShrinkLast, 2003),
        DEFINE_FNCODE_CMD(Twapi_MemLifoResizeLast, 2004),
    };

    static struct fncode_dispatch_s CallNoargsDispatch[] = {
        DEFINE_FNCODE_CMD(GetCurrentProcess, 1),
        DEFINE_FNCODE_CMD(GetVersionEx, 2),
        DEFINE_FNCODE_CMD(UuidCreateNil, 3),
        DEFINE_FNCODE_CMD(Twapi_GetInstallDir, 4),
        DEFINE_FNCODE_CMD(EnumWindows, 5),
        DEFINE_FNCODE_CMD(GetSystemWindowsDirectory, 6), /* TBD Tcl */
        DEFINE_FNCODE_CMD(GetWindowsDirectory, 7),       /* TBD Tcl */
        DEFINE_FNCODE_CMD(GetSystemDirectory, 8),        /* TBD Tcl */
        DEFINE_FNCODE_CMD(GetCurrentThreadId, 9),
        DEFINE_FNCODE_CMD(GetTickCount, 10),
        DEFINE_FNCODE_CMD(GetSystemTimeAsFileTime, 11),
        DEFINE_FNCODE_CMD(AllocateLocallyUniqueId, 12),
        DEFINE_FNCODE_CMD(LockWorkStation, 13),
        DEFINE_FNCODE_CMD(RevertToSelf, 14), /* Left in base module as it might be
                                       used from multiple extensions */
        DEFINE_FNCODE_CMD(GetSystemPowerStatus, 15),
        DEFINE_FNCODE_CMD(DebugBreak, 16),
        DEFINE_FNCODE_CMD(GetDefaultPrinter, 17),         /* TBD Tcl */
    };

    static struct fncode_dispatch_s CallIntArgDispatch[] = {
        DEFINE_FNCODE_CMD(GetStdHandle, 1),
        DEFINE_FNCODE_CMD(UuidCreate, 2),
        DEFINE_FNCODE_CMD(GetUserNameEx, 3),
        DEFINE_FNCODE_CMD(Twapi_MapWindowsErrorToString, 4),
        DEFINE_FNCODE_CMD(Twapi_MemLifoInit, 5),
        DEFINE_FNCODE_CMD(GlobalDeleteAtom, 6), // TBD - tcl interface
    };

    static struct fncode_dispatch_s CallOneArgDispatch[] = {
        DEFINE_FNCODE_CMD(Twapi_AddressToPointer, 1008),
        DEFINE_FNCODE_CMD(IsValidSid, 1009),
        DEFINE_FNCODE_CMD(VariantTimeToSystemTime, 1010),
        DEFINE_FNCODE_CMD(SystemTimeToVariantTime, 1011),
        DEFINE_FNCODE_CMD(canonicalize_guid, 1012), // TBD Document
        DEFINE_FNCODE_CMD(Twapi_AppendLog, 1013),
        DEFINE_FNCODE_CMD(GlobalAddAtom, 1014), // TBD - Tcl interface
        DEFINE_FNCODE_CMD(is_valid_sid_syntax, 1015),
        DEFINE_FNCODE_CMD(FileTimeToSystemTime, 1016),
        DEFINE_FNCODE_CMD(SystemTimeToFileTime, 1017),
        DEFINE_FNCODE_CMD(GetWindowThreadProcessId, 1018),
        DEFINE_FNCODE_CMD(Twapi_IsValidGUID, 1019),
        DEFINE_FNCODE_CMD(ExpandEnvironmentStrings, 1020),
        DEFINE_FNCODE_CMD(free, 1021),
        DEFINE_FNCODE_CMD(registered_pointer, 1022),
    };

    static struct fncode_dispatch_s CallArgsDispatch[] = {
        DEFINE_FNCODE_CMD(SystemParametersInfo, 10001),
        DEFINE_FNCODE_CMD(LookupAccountSid, 10002),
        DEFINE_FNCODE_CMD(LookupAccountName, 10003),
        DEFINE_FNCODE_CMD(NetGetDCName, 10004),
        DEFINE_FNCODE_CMD(AttachThreadInput, 10005),
        DEFINE_FNCODE_CMD(GlobalAlloc, 10006),
        DEFINE_FNCODE_CMD(LHashValOfName, 10007),
        DEFINE_FNCODE_CMD(DuplicateHandle, 10008),
        DEFINE_FNCODE_CMD(Tcl_GetChannelHandle, 10009),
        DEFINE_FNCODE_CMD(SetStdHandle, 10010),
        DEFINE_FNCODE_CMD(LoadLibraryEx, 10011),
        DEFINE_FNCODE_CMD(CreateFile, 10012),
        DEFINE_FNCODE_CMD(Twapi_SourceResource, 10013),
        DEFINE_FNCODE_CMD(FindWindowEx, 10014),
        DEFINE_FNCODE_CMD(LsaQueryInformationPolicy, 10015),
        DEFINE_FNCODE_CMD(Twapi_LsaOpenPolicy, 10016),
        DEFINE_FNCODE_CMD(GetWindowLongPtr, 10017),
        DEFINE_FNCODE_CMD(PostMessage, 10018),
        DEFINE_FNCODE_CMD(SendNotifyMessage, 10019),
        DEFINE_FNCODE_CMD(SendMessageTimeout, 10020),
        DEFINE_FNCODE_CMD(SetWindowLongPtr, 10021),
        DEFINE_FNCODE_CMD(DsGetDcName, 10022),
        DEFINE_FNCODE_CMD(FormatMessageFromModule, 10023),
        DEFINE_FNCODE_CMD(FormatMessageFromString, 10024),
        DEFINE_FNCODE_CMD(win32_error, 10025),
        DEFINE_FNCODE_CMD(CreateMutex, 10026),
        DEFINE_FNCODE_CMD(OpenMutex, 10027),
        DEFINE_FNCODE_CMD(OpenSemaphore, 10028), /* TBD - Tcl wrapper */
        DEFINE_FNCODE_CMD(CreateSemaphore, 10029), /* TBD - Tcl wrapper */
        DEFINE_FNCODE_CMD(malloc, 10030),        /* TBD - document, change to memalloc */
        DEFINE_FNCODE_CMD(Twapi_IsEqualPtr, 10031),
        DEFINE_FNCODE_CMD(Twapi_IsNullPtr, 10032),
        DEFINE_FNCODE_CMD(Twapi_IsPtr, 10033),
        DEFINE_FNCODE_CMD(CreateEvent, 10034),
        DEFINE_FNCODE_CMD(IsEqualGUID, 10035), // Tcl
        DEFINE_FNCODE_CMD(OpenEventLog, 10036),
        DEFINE_FNCODE_CMD(RegisterEventSource, 10037),
    };

    static struct alias_dispatch_s AliasDispatch[] = {
        DEFINE_ALIAS_CMD(Twapi_GetAtomStats, 1),
        DEFINE_ALIAS_CMD(TwapiId, 2),
        DEFINE_ALIAS_CMD(Twapi_GetNotificationWindow, 3),

        DEFINE_ALIAS_CMD(Twapi_UnregisterWaitOnHandle, 4),
        DEFINE_ALIAS_CMD(atomize, 5),

        DEFINE_ALIAS_CMD(TranslateName, 6),
        DEFINE_ALIAS_CMD(Twapi_RegisterWaitOnHandle, 7),
    };

    struct tcl_dispatch_s TclDispatch[] = {
        DEFINE_TCL_CMD(Call, Twapi_CallObjCmd),
        DEFINE_TCL_CMD(parseargs, Twapi_ParseargsObjCmd),
        DEFINE_TCL_CMD(trap, Twapi_TrapObjCmd),
        DEFINE_TCL_CMD(kl_get, Twapi_KlGetObjCmd),
        DEFINE_TCL_CMD(twine, Twapi_TwineObjCmd),
        DEFINE_TCL_CMD(recordarray, Twapi_RecordArrayObjCmd),
        DEFINE_TCL_CMD(GetTwapiBuildInfo, Twapi_GetTwapiBuildInfo),
        DEFINE_TCL_CMD(Twapi_ReadMemory, Twapi_ReadMemoryObjCmd),
        DEFINE_TCL_CMD(Twapi_WriteMemory, Twapi_WriteMemoryObjCmd),
        DEFINE_TCL_CMD(tclcast, Twapi_InternalCastObjCmd),
        DEFINE_TCL_CMD(tcltype, Twapi_GetTclTypeObjCmd),
        DEFINE_TCL_CMD(Twapi_EnumPrinters_Level4, Twapi_EnumPrintersLevel4ObjCmd),
        DEFINE_TCL_CMD(ReportEvent, Twapi_ReportEventObjCmd),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CallHDispatch), CallHDispatch, Twapi_CallHObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CallNoargsDispatch), CallNoargsDispatch, Twapi_CallNoargsObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CallIntArgDispatch), CallIntArgDispatch, Twapi_CallIntArgObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CallOneArgDispatch), CallOneArgDispatch, Twapi_CallOneArgObjCmd);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CallArgsDispatch), CallArgsDispatch, Twapi_CallArgsObjCmd);

    TwapiDefineTclCmds(interp, ARRAYSIZE(TclDispatch), TclDispatch, ticP);

    TwapiDefineAliasCmds(interp, ARRAYSIZE(AliasDispatch), AliasDispatch, "twapi::Call");

    return TCL_OK;
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
        TwapiSetStaticResult(interp, "Unknown channel");
        return TCL_ERROR;
    }
    
    direction = direction ? TCL_WRITABLE : TCL_READABLE;
    
    if (Tcl_GetChannelHandle(chan, direction, &h) == TCL_ERROR) {
        TwapiSetStaticResult(interp, "Error getting channel handle");
        return TCL_ERROR;
    }

    return TwapiSetObjResult(interp, ObjFromHANDLE(h));
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
        TwapiSetStaticResult(interp, "Error looking up account SID: ");
        Twapi_AppendSystemError(interp, error);
        goto done;
    }

    /* Allocate required space */
    domainP = TwapiAlloc(domain_buf_size * sizeof(*domainP));
    nameP = TwapiAlloc(name_buf_size * sizeof(*nameP));

    if (LookupAccountSidW(lpSystemName, sidP, nameP, &name_buf_size,
                          domainP, &domain_buf_size, &account_type) == 0) {
        TwapiSetStaticResult(interp, "Error looking up account SID: ");
        Twapi_AppendSystemError(interp, GetLastError());
        goto done;
    }

    /*
     * Got everything we need, now format it
     * {NAME DOMAIN ACCOUNT}
     */
    objs[0] = ObjFromUnicode(nameP);   /* Will exit on alloc fail */
    objs[1] = ObjFromUnicode(domainP); /* Will exit on alloc fail */
    objs[2] = ObjFromInt(account_type);
    // TBD - What about freeing nameP, domainP ?
    return TwapiSetObjResult(interp, ObjNewList(3, objs));

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
        TwapiSetStaticResult(interp, "Error looking up account name: ");
        Twapi_AppendSystemError(interp, error);
        return TCL_ERROR;
    }

    /* Allocate required space */
    domainP = TwapiAlloc(domain_buf_size * sizeof(*domainP));
    sidP = TwapiAlloc(sid_buf_size);

    if (LookupAccountNameW(lpSystemName, lpAccountName, sidP, &sid_buf_size,
                          domainP, &domain_buf_size, &account_type) == 0) {
        TwapiSetStaticResult(interp, "Error looking up account name: ");
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
    objs[2] = ObjFromInt(account_type);
    // TBD - what about freeing up sidP and domainP
    return TwapiSetObjResult(interp, ObjNewList(3, objs));

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
    TwapiSetObjResult(interp, ObjFromUnicode((wchar_t *)bufP));
    NetApiBufferFree(bufP);
    return TCL_OK;
}

/* Window enumeration callback. Note this is called from other modules also */
BOOL CALLBACK Twapi_EnumWindowsCallback(HWND hwnd, LPARAM p_ctx) {
    TwapiEnumCtx *p_enum_win_ctx =
        (TwapiEnumCtx *) p_ctx;

    ObjAppendElement(p_enum_win_ctx->interp,
                             p_enum_win_ctx->objP,
                             ObjFromHWND(hwnd));
    
    return 1;
}

/* Enumerate toplevel windows. In twapi_base because commonly needed */
int Twapi_EnumWindows(Tcl_Interp *interp)
{
    TwapiEnumCtx enum_win_ctx;

    enum_win_ctx.interp = interp;
    enum_win_ctx.objP = ObjNewList(0, NULL);
    
    if (EnumWindows(Twapi_EnumWindowsCallback, (LPARAM)&enum_win_ctx) == 0) {
        TwapiReturnSystemError(interp);
        Twapi_FreeNewTclObj(enum_win_ctx.objP);
        return TCL_ERROR;
    }

    return TwapiSetObjResult(interp, enum_win_ctx.objP);
}

TCL_RESULT Twapi_LsaQueryInformationPolicy (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    void    *buf;
    NTSTATUS ntstatus;
    int      retval;
    Tcl_Obj  *objs[5];
    LSA_HANDLE lsaH;
    int        infoclass;

    if (objc != 2 ||
        ObjToOpaque(interp, objv[0], (void **) &lsaH, "LSA_HANDLE") != TCL_OK ||
        ObjToLong(interp, objv[1], &infoclass) != TCL_OK) {
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
    }    

    ntstatus = LsaQueryInformationPolicy(lsaH, infoclass, &buf);
    if (ntstatus != STATUS_SUCCESS) {
        return Twapi_AppendSystemError(interp,
                                       TwapiNTSTATUSToError(ntstatus));
    }

    retval = TCL_OK;
    switch (infoclass) {
    case PolicyAccountDomainInformation:
        objs[0] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_ACCOUNT_DOMAIN_INFO *) buf)->DomainName)
            );
        objs[1] = ObjFromSIDNoFail(((POLICY_ACCOUNT_DOMAIN_INFO *) buf)->DomainSid);
        TwapiSetObjResult(interp, ObjNewList(2, objs));
        break;

    case PolicyDnsDomainInformation:
        objs[0] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_DNS_DOMAIN_INFO *) buf)->Name)
            );
        objs[1] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_DNS_DOMAIN_INFO *) buf)->DnsDomainName)
            );
        objs[2] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_DNS_DOMAIN_INFO *) buf)->DnsForestName)
            );
        objs[3] = ObjFromUUID(
            (UUID *) &(((POLICY_DNS_DOMAIN_INFO *) buf)->DomainGuid)
            );
        objs[4] = ObjFromSIDNoFail(((POLICY_DNS_DOMAIN_INFO *) buf)->Sid);
        TwapiSetObjResult(interp, ObjNewList(5, objs));

        break;

    default:
        TwapiSetStaticResult(interp, "Invalid/unsupported information class");
        retval = TCL_ERROR;
    }

    LsaFreeMemory(buf);

    return retval;
}

