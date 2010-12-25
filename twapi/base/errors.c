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
};
#define TWAPI_ERROR_MAP_SIZE (sizeof(error_map)/sizeof(error_map[0]))

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
        objP = Tcl_NewUnicodeObj(wMsgPtr, length);
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
        objP = Tcl_NewStringObj(msgPtr, length);
        LocalFree(msgPtr);
        return objP;
    }

    return NULL;
}


Tcl_Obj *TwapiGetErrorMsg(int error)
{
    char *msg = NULL;
    char  buf[128];

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

    if (msg == NULL) {
        StringCbPrintfA(buf, sizeof(buf), "Twapi error %d", error);
        msg = buf;
    }
        
    return Tcl_NewStringObj(msg, -1);
}

/* Returns a Tcl errorCode object from a TWAPI error */
Tcl_Obj *Twapi_MakeTwapiErrorCodeObj(int error) 
{
    Tcl_Obj *objv[3];

    objv[0] = Tcl_NewStringObj("TWAPI", 5);
    objv[1] = Tcl_NewIntObj(error);
    objv[2] = TwapiGetErrorMsg(error);
    return Tcl_NewListObj(3, objv);
}

/* Sets the interpreter result to a TWAPI error */
int TwapiReturnTwapiError(Tcl_Interp *interp, char *msg, int code)
{
    if (interp) {
        Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(code));
        if (msg)
            Tcl_SetObjResult(interp, Tcl_NewStringObj(msg, -1));
        else
            Tcl_SetObjResult(interp, TwapiGetErrorMsg(code));
    }
    return TCL_ERROR;           /* Always, so caller can just return */
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
    char msgbuf[24 + TCL_INTEGER_SPACE];
    Tcl_Obj *objP;

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
        hNtdll = GetModuleHandle("ntdll.dll");
    if (hNtdll)
        objP = Twapi_FormatMsgFromModule(error, hNtdll);
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
        StringCbPrintfA(msgbuf, sizeof(msgbuf), "Windows error: %ld", error);
        return Tcl_NewStringObj(msgbuf, -1);
    }
}

/* Returns a Tcl errorCode object */
Tcl_Obj *Twapi_MakeWindowsErrorCodeObj(DWORD error, Tcl_Obj *extra) 
{
    Tcl_Obj *objv[4];

    objv[0] = STRING_LITERAL_OBJ(TWAPI_WIN32_ERRORCODE_TOKEN);
    objv[1] = Tcl_NewLongObj(error);
    objv[2] = Twapi_MapWindowsErrorToString(error);
    objv[3] = extra;
    return Tcl_NewListObj(objv[3] ? 4 : 3, objv);
}

/*
 * Format system error codes as the interpreter results. Also sets
 * the global errorCode Tcl variable.
 * Always returns TCL_ERROR so caller can just do a
 *  return Twapi_AppendSystemError2(interp, error)
 * to return after an error.
 */
int Twapi_AppendSystemError2(
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
    if (Tcl_ListObjIndex(NULL, errorCodeObj, 2, &msgObj) == TCL_OK &&
        msgObj != NULL) {
        Tcl_Obj *resultObj = Tcl_DuplicateObj(Tcl_GetObjResult(interp));
        if (Tcl_GetCharLength(resultObj)) {
            Tcl_AppendUnicodeToObj(resultObj, L" ", 1);
        }
        Tcl_AppendObjToObj(resultObj, msgObj);
        Tcl_SetObjResult(interp, resultObj);
    }
    Tcl_SetObjErrorCode(interp, errorCodeObj);

    return TCL_ERROR;           /* Always return TCL_ERROR */
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
        Tcl_Obj *resultObj = Tcl_DuplicateObj(Tcl_GetObjResult(interp));
        Tcl_AppendUnicodeToObj(resultObj, L" ", 1);
        Tcl_AppendUnicodeToObj(resultObj, provider, -1);
        Tcl_AppendUnicodeToObj(resultObj, L": ", 2);
        Tcl_AppendUnicodeToObj(resultObj, errorbuf, -1);
        Tcl_SetObjResult(interp, resultObj);
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
