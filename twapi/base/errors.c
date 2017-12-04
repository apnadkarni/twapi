/* 
 * Copyright (c) 2006-2009, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Mapping between TWAPI errors and messages */

#include "twapi.h"

struct TWAPI_ERROR_MAP {
    int   code;
    char *msg;
};

static struct TWAPI_ERROR_MAP error_map[] = {
    {TWAPI_NO_ERROR, "No error."},
    {TWAPI_INVALID_ARGS, "Invalid or badly formatted arguments specified."},
    {TWAPI_BUFFER_OVERRUN, "Attempt to write past the end of memory buffer."},
    {TWAPI_EXTRA_ARGS, "Extra arguments specified."},
    {TWAPI_BAD_ARG_COUNT, "Incorrect number of arguments."},
    {TWAPI_INTERNAL_LIMIT, "Internal limit exceeded."},
    {TWAPI_INVALID_OPTION, "Invalid option specified."},
    {TWAPI_INVALID_FUNCTION_CODE, "Invalid or unknown function code passed to Win32 API dispatcher."},
    {TWAPI_BUG, "An internal error has been detected."},
    {TWAPI_UNKNOWN_OBJECT, "Specified resource or object not found."},
    {TWAPI_SYSTEM_ERROR, "System error."},
    {TWAPI_REGISTER_WAIT_FAILED, "Could not register thread pool wait."},
    {TWAPI_BUG_INVALID_STATE_FOR_OP, "Internal error: requested operation is not valid for current state."},
    {TWAPI_OUT_OF_RANGE, "Value is out of range."},
    {TWAPI_UNSUPPORTED_TYPE, "Unsupported type, format or function."},
    {TWAPI_REGISTERED_POINTER_EXISTS, "Attempt to register duplicate pointer."},
    {TWAPI_REGISTERED_POINTER_TAG_MISMATCH, "Type of pointer does not match registered type."},
    {TWAPI_REGISTERED_POINTER_NOTFOUND, "Pointer is not registered. Probably invalid or already freed."},
    {TWAPI_NULL_POINTER, "Pointer is NULL."},
    {TWAPI_REGISTERED_POINTER_IS_NOT_COUNTED, "Pointer is registered as an uncounted pointer."},
    {TWAPI_INVALID_COMMAND_SCOPE, "Command cannot be called in this scope."},
    {TWAPI_SCRIPT_ERROR, "Script error."},
    {TWAPI_INVALID_DATA, "Invalid data."},
    {TWAPI_INVALID_PTR, "Invalid pointer."},
};
#define TWAPI_ERROR_MAP_SIZE (sizeof(error_map)/sizeof(error_map[0]))

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
        ObjSetResult(interp, ObjFromWinChars(msgP));
        LocalFree(msgP);
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
}

static Tcl_Obj *Twapi_FormatMsgFromModule(DWORD error, HANDLE hModule)
{
    int   length;
    DWORD flags;
    WCHAR *wMsgPtr = NULL;
    char  *msgPtr;
    Tcl_Obj *objP;

    if (hModule) {
        flags = FORMAT_MESSAGE_FROM_HMODULE;
    } else {
        flags = FORMAT_MESSAGE_FROM_SYSTEM;
    }
    flags |= FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_IGNORE_INSERTS;

// TBD - value of 0 -> specific lang search path else only specified user
// default language. Try both on French system and see what happens
//#define TWAPI_ERROR_LANGID MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
#define TWAPI_ERROR_LANGID 0

    length = FormatMessageW(flags, hModule, error,
                            TWAPI_ERROR_LANGID,
                            (WCHAR *) &wMsgPtr,
                            0, NULL);
    if (length > 0) {
        /* Strip trailing CR LF if any */
        if (wMsgPtr[length-1] == L'\n')
            --length;
        if (length > 0) {
            if (wMsgPtr[length-1] == L'\r')
                --length;
        }
        objP = ObjFromWinCharsN(wMsgPtr, length);
        LocalFree(wMsgPtr);
        return objP;
    }

    /* Try the ascii version. TBD - is this really meaningful if above failed ? */

    length = FormatMessageA(flags, hModule, error,
                            TWAPI_ERROR_LANGID,
                            (char *) &msgPtr,
                            0, NULL);

    if (length > 0) {
        /* Strip trailing CR LF if any */
        if (msgPtr[length-1] == '\n')
            --length;
        if (length > 0) {
            if (msgPtr[length-1] == L'\r')
                --length;
        }
        objP = ObjFromStringN(msgPtr, length);
        LocalFree(msgPtr);
        return objP;
    }

    return NULL;
}

/* Returns ptr to static string or NULL */
static const char *TwapiMapErrorCode(int error)
{
    char *msg = NULL;

    /* Optimization - the entry in table will likely be indexed by the error */
    if (error >= 0 &&
        error < TWAPI_ERROR_MAP_SIZE &&
        error_map[error].code == error) {
        msg = error_map[error].msg;
    } else {
        /* Loop to look for error */
        int i;
	for (i = 0 ; i < TWAPI_ERROR_MAP_SIZE ; ++i) {
	    if (error_map[i].code == error) {
                msg = error_map[i].msg;
		break;
	    }
	}
    }

    return msg;
}

Tcl_Obj *TwapiGetErrorMsg(int error)
{
    const char *msg = TwapiMapErrorCode(error);
    if (msg)
        return ObjFromString(msg);
    else
        return Tcl_ObjPrintf("Twapi error %d.", error);
}

/* Returns a Tcl errorCode object from a TWAPI error */
Tcl_Obj *Twapi_MakeTwapiErrorCodeObj(int error) 
{
    Tcl_Obj *objv[3];

    objv[0] = STRING_LITERAL_OBJ("TWAPI");
    objv[1] = ObjFromInt(error);
    objv[2] = TwapiGetErrorMsg(error);
    return ObjNewList(3, objv);
}

int TwapiReturnError(Tcl_Interp *interp, int code)
{
    if (interp) {
        Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(code));
        (void) ObjSetResult(interp, TwapiGetErrorMsg(code));
    }
    /* TBD - what if code is TWAPI_NO_ERROR ? */
    return TCL_ERROR;           /* Always, so caller can just return */
}

int TwapiReturnErrorMsg(Tcl_Interp *interp, int code, char *msg)
{
    if (msg)
        return TwapiReturnErrorEx(interp, code, ObjFromString(msg));
    else
        return TwapiReturnError(interp, code);
}

/* Sets the interpreter result to a TWAPI error with addition message */
int TwapiReturnErrorEx(Tcl_Interp *interp, int code, Tcl_Obj *objP)
{
    if (interp) {
        Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(code));
        if (objP) {
            const char *msgP;
            if (Tcl_IsShared(objP))
                objP = ObjDuplicate(objP); /* since we are modifying it */
            msgP = TwapiMapErrorCode(code);
            if (msgP)
                Tcl_AppendStringsToObj(objP, " TWAPI error: ", msgP, NULL);
            else
                Tcl_AppendPrintfToObj(objP, " TWAPI error code: %d", code);
            (void) ObjSetResult(interp, objP);
        } else {
            TwapiReturnError(interp, code);
        }
    }
    /* TBD - what if code is TWAPI_NO_ERROR ? */
    return TCL_ERROR; /* Always, so caller can just return */
}

int TwapiReturnSystemError(Tcl_Interp *interp)
{
    if (interp) {
        DWORD error = GetLastError();
        if (error == 0) error = ERROR_INVALID_PARAMETER;
        Tcl_ResetResult(interp);
        return Twapi_AppendSystemError(interp, error);
    } else
        return TCL_ERROR;
}

/*
 * Return a Tcl_Obj string corresponding to a Windows error
 * NEVER RETURNS NULL.
 */
Tcl_Obj *Twapi_MapWindowsErrorToString(DWORD error)
{
    static HMODULE hNetmsg;
    static HMODULE hPdh;
    static HMODULE hNtdll;
    static HMODULE hWmi;
    Tcl_Obj *objP;

    /* TBD -
       - loop instead of separately trying each module
       - construct full path to modules
       - write test cases for all modules
    */

    /* First try mapping as a system error */
    objP  = Twapi_FormatMsgFromModule(error, NULL);
    if (objP)
        return objP;

    /* TBD - do we need to FreeLibrary after a LoadLibraryExW ? */
    /* Next try as a netmsg error - it's not clear this is really
       required or whether the system formatting call also looks
       up netmsg.dll */
    if ((error >= NERR_BASE) && (error <= MAX_NERR)) {
        if (hNetmsg == NULL)
            hNetmsg = LoadLibraryExW(L"netmsg.dll", NULL,
                                     LOAD_LIBRARY_AS_DATAFILE);
        if (hNetmsg)
            objP = Twapi_FormatMsgFromModule(error, hNetmsg);
        if (objP)
            return objP;
    }

    /* Try as NTSTATUS code */
    if (hNtdll == NULL)
        hNtdll = GetModuleHandleW(L"ntdll.dll");
    if (hNtdll)
        objP = Twapi_FormatMsgFromModule(error, hNtdll);
    if (objP)
        return objP;

    /* How about WMI ? */
    if (hWmi == NULL)
        hWmi = LoadLibraryExW(L"wmiutils.dll", NULL,
                              LOAD_LIBRARY_AS_DATAFILE);
    if (hWmi)
        objP = Twapi_FormatMsgFromModule(error, hWmi);
    if (objP)
        return objP;


    /* Still no joy, try the PDH */
    if (hPdh == NULL)
        hPdh = LoadLibraryExW(L"pdh.dll", NULL,
                              LOAD_LIBRARY_AS_DATAFILE);
    if (hPdh)
        objP = Twapi_FormatMsgFromModule(error, hPdh);
    if (objP)
        return objP;

    /* Perhaps a TWAPI error ? */
    if (IS_TWAPI_WIN32_ERROR(error))
        return TwapiGetErrorMsg(TWAPI_WIN32_ERROR_TO_CODE(error));

    /* Just print out error code */
    if (error == ERROR_CALL_NOT_IMPLEMENTED) {
        return STRING_LITERAL_OBJ("Function not supported under this Windows version");
    } else {
        return Tcl_ObjPrintf("Windows error: %ld", error);
    }
}

/* Returns a Tcl errorCode object */
Tcl_Obj *Twapi_MakeWindowsErrorCodeObj(DWORD error, Tcl_Obj *extra) 
{
    Tcl_Obj *objv[4];

    objv[0] = STRING_LITERAL_OBJ(TWAPI_WIN32_ERRORCODE_TOKEN);
    objv[1] = ObjFromLong(error);
    objv[2] = Twapi_MapWindowsErrorToString(error);
    objv[3] = extra;
    return ObjNewList(objv[3] ? 4 : 3, objv);
}

/*
 * Format system error codes as the interpreter results. Also sets
 * the global errorCode Tcl variable.
 * Always returns TCL_ERROR so caller can just do a
 *  return Twapi_AppendSystemErrorEx(interp, error)
 * to return after an error.
 */
int Twapi_AppendSystemErrorEx(
    Tcl_Interp *interp,	/* Current interpreter. If NULL, this is a NO-OP
                           except for return value */
    DWORD error,	/* Result code from error. */
    Tcl_Obj *extra      /* Additional argument to be tacked on to
                           errorCode object, may be NULL */
    )
{
    Tcl_Obj *msgObj;
    Tcl_Obj *errorCodeObj;

    if (interp == NULL)
        return TCL_ERROR;

    errorCodeObj = Twapi_MakeWindowsErrorCodeObj(error, extra);

    /* Third element of error code is also the message */
    if (ObjListIndex(NULL, errorCodeObj, 2, &msgObj) == TCL_OK &&
        msgObj != NULL) {
        Tcl_Obj *resultObj = ObjDuplicate(ObjGetResult(interp));
        if (ObjCharLength(resultObj)) {
#if TCL_UTF_MAX <= 4
            Tcl_AppendUnicodeToObj(resultObj, L" ", 1);
#else
            /* Tcl_UniChar is int. So cannot use AppendUnicode. Have 
               to force a shimmer to string */
            Tcl_AppendToObj(resultObj, " ", 1);
#endif
        }
        Tcl_AppendObjToObj(resultObj, msgObj);
        (void) ObjSetResult(interp, resultObj);
    }
    Tcl_SetObjErrorCode(interp, errorCodeObj);

    return TCL_ERROR;           /* Always return TCL_ERROR */
}


int Twapi_AppendSystemError(
    Tcl_Interp *interp,	/* Current interpreter. If NULL, this is a NO-OP
                           except for return value */
    DWORD error	/* Result code from error. */
    )
{
    return Twapi_AppendSystemErrorEx(interp, error, NULL);
}


/*
 * Called to generate error message when a WNet* function fails.
 * Should be called RIGHT AFTER a WNet call fails
 * Always returns TCL_ERROR
 */
int Twapi_AppendWNetError(
    Tcl_Interp *interp,		/* Current interpreter. */
    DWORD error                 /* Win32 error */
    )
{
    WCHAR provider[256];
    WCHAR errorbuf[1024];
    DWORD wneterror;
    DWORD wnetcode;

    if (error == ERROR_EXTENDED_ERROR) {
        /* Get the WNet error BEFORE we do ANYTHING else */
        wneterror = WNetGetLastErrorW(&wnetcode,
                                      errorbuf, ARRAYSIZE(errorbuf),
                                      provider, ARRAYSIZE(provider));
    }

    /* First write the win32 error message */
    Twapi_AppendSystemError(interp, error);

    /* If we had a extended error and we got it successfully, 
     * append the WNet message 
     */
    if (error == ERROR_EXTENDED_ERROR && wneterror == NO_ERROR) {
        Tcl_Obj *resultObj = ObjDuplicate(ObjGetResult(interp));
        Tcl_DString ds;

        Tcl_DStringInit(&ds);
        Tcl_DStringAppend(&ds, (char *) L" ", sizeof (WCHAR));
        Tcl_DStringAppend(&ds, (char *) provider, lstrlenW(provider) * sizeof (WCHAR));
        Tcl_DStringAppend(&ds, (char *) L": ", 2 * sizeof (WCHAR));
        Tcl_DStringAppend(&ds, (char *) errorbuf, lstrlenW(errorbuf) * sizeof (WCHAR));
        Tcl_DStringAppend(&ds, (char *) L"\0", sizeof (WCHAR));
        Tcl_AppendObjToObj(resultObj, ObjFromWinChars((WCHAR *) Tcl_DStringValue(&ds)));
        Tcl_DStringFree(&ds);
        (void) ObjSetResult(interp, resultObj);
    }

    return TCL_ERROR;
}

int Twapi_GenerateWin32Error(
    Tcl_Interp *interp,
    DWORD error,
    char *msg
    )
{
    if (interp) {
        if (msg)
            Tcl_SetResult(interp, msg, TCL_VOLATILE);
        Twapi_AppendSystemError(interp, error);
    }
    return TCL_ERROR;           /* Always. We want to generate exception */
}


/* Map a NTSTATUS to a Windows error */
typedef ULONG (WINAPI *RtlNtStatusToDosError_t)(NTSTATUS);
MAKE_DYNLOAD_FUNC(RtlNtStatusToDosError, ntdll, RtlNtStatusToDosError_t)
DWORD TwapiNTSTATUSToError(NTSTATUS status)
{
    RtlNtStatusToDosError_t RtlNtStatusToDosErrorPtr = Twapi_GetProc_RtlNtStatusToDosError();
    if (RtlNtStatusToDosErrorPtr) {
        return (DWORD) (*RtlNtStatusToDosErrorPtr)(status);
    }
    else {
        return LsaNtStatusToWinError(status);
    }
}

/*
 * Always returns TCL_ERROR
 */
int Twapi_AppendCOMError(Tcl_Interp *interp, HRESULT hr, ISupportErrorInfo *sei, REFIID iid)
{
    IErrorInfo *ei = NULL;
    if (sei && iid) {
        if (SUCCEEDED(sei->lpVtbl->InterfaceSupportsErrorInfo(sei, iid))) {
            GetErrorInfo(0, &ei);
        }
    }
    if (ei) {
        BSTR msg;
        ei->lpVtbl->GetDescription(ei, &msg);
        Twapi_AppendSystemErrorEx(interp, hr,
                                 ObjFromWinCharsN(msg,SysStringLen(msg)));

        SysFreeString(msg);
        ei->lpVtbl->Release(ei);
    } else {
        Twapi_AppendSystemError(interp, hr);
    }

    return TCL_ERROR;
}

TCL_RESULT Twapi_WrongLevelError(Tcl_Interp *interp, int level)
{
    return TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                              Tcl_ObjPrintf("Invalid info level %d.", level));
}


/*
 * Write a message to the event log. To be used only for errors that
 * cannot be raised through the Tcl interpreter, for example, service
 * startup errors
 */
void TwapiWriteEventLogError(const char *msg)
{
    HANDLE hevl;
    hevl = RegisterEventSourceA(NULL, TWAPI_TCL_NAMESPACE);
    if (hevl) {
        (void) ReportEventA(hevl, EVENTLOG_ERROR_TYPE, 0, 0, NULL, 1, 0, &msg, NULL);
        DeregisterEventSource(hevl);
    }
}

TWAPI_EXTERN void TwapiUnreachablePanic()
{
    Tcl_Panic("Unreachable code executed.");
}
