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
};
#define TWAPI_ERROR_MAP_SIZE (sizeof(error_map)/sizeof(error_map[0]))

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
        wsprintfA(buf, "Twapi error %d", error);
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
 * Return a Unicode string corresponding to a WIndows error
 */
LPWSTR Twapi_MapWindowsErrorToString(DWORD error)
{
    WCHAR *wMsgPtr;
    static HMODULE hNetmsg;
    static HMODULE hPdh;
    WCHAR wMsgBuf[24 + TCL_INTEGER_SPACE];

    /* First try mapping as a system error */
    wMsgPtr = Twapi_FormatMsgFromModule(error, NULL);
    if (wMsgPtr)
        return wMsgPtr;

    /* TBD - do we need to FreeLibrary after a LoadLibraryExW ? */
    /* Next try as a netmsg error - it's not clear this is really
       required or whether the system formatting call also looks
       up netmsg.dll */
    if ((error >= NERR_BASE) && (error <= MAX_NERR)) {
        if (hNetmsg == NULL)
            hNetmsg = LoadLibraryExW(L"netmsg.dll", NULL,
                                     LOAD_LIBRARY_AS_DATAFILE);
        if (hNetmsg)
            wMsgPtr = Twapi_FormatMsgFromModule(error, hNetmsg);
        if (wMsgPtr)
            return wMsgPtr;
    }

    /* Still no joy, try the PDH */
    if (hPdh == NULL)
        hPdh = LoadLibraryExW(L"pdh.dll", NULL,
                              LOAD_LIBRARY_AS_DATAFILE);
    if (hPdh)
        wMsgPtr = Twapi_FormatMsgFromModule(error, hPdh);
    if (wMsgPtr)
        return wMsgPtr;
    
    /* Just print out error code */
    if (error == ERROR_CALL_NOT_IMPLEMENTED) {
        wMsgPtr = TwapiAllocWString(L"function not supported under this Windows version", -1);
    } else {
        wsprintfW(wMsgBuf, L"Windows error: %ld", error);
        wMsgPtr = TwapiAllocWString(wMsgBuf, -1);
    }

    return wMsgPtr;
}

/* Returns a Tcl errorCode object */
Tcl_Obj *Twapi_MakeWindowsErrorCodeObj(DWORD error, Tcl_Obj *extra) 
{
    Tcl_Obj *objv[4];
    LPWSTR   wmsgP;
    char     buf[40];

    objv[0] = STRING_LITERAL_OBJ(TWAPI_WIN32_ERRORCODE_TOKEN);
    objv[1] = Tcl_NewLongObj(error);
    wmsgP = Twapi_MapWindowsErrorToString(error);
    if (wmsgP) {
        int len;

        len = (int) lstrlenW(wmsgP);

        objv[2] = Tcl_NewUnicodeObj(wmsgP, len);
        TwapiFree(wmsgP);
    } else {
        wsprintf(buf, "Windows error: %ld", error);
        objv[2] = Tcl_NewStringObj(buf, -1);
    }
    objv[3] = extra;
    return Tcl_NewListObj(objv[3] ? 4 : 3, objv);
}

/*
 * Format system error codes as the interpreter results. Also sets
 * the global errorCode Tcl variable.
 * Always returns TCL_ERROR so caller can just do a
 *  return Twapi_AppendSystemError(interp, error)
 * to return after an error.
 */
int Twapi_AppendSystemError2(
    Tcl_Interp *interp,	/* Current interpreter. If NULL, this is a NO-OP */
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
        Tcl_Obj *resultObj = Tcl_GetObjResult(interp);
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
 * Returns in *ppMsg, a TwapiAlloc'ed unicode string corresponding to the given
 * error code by looking up the appropriate dll or system if the module
 * handle is NULL.
 */
LPWSTR Twapi_FormatMsgFromModule(DWORD error, HANDLE hModule)
{
    int   length;
    DWORD flags;
    WCHAR *wMsgPtr = NULL;
    char  *msgPtr;

    if (hModule) {
        flags = FORMAT_MESSAGE_FROM_HMODULE;
    } else {
        flags = FORMAT_MESSAGE_FROM_SYSTEM;
    }
    flags |= FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_IGNORE_INSERTS;

    length = FormatMessageW(flags, hModule, error,
                            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                            (WCHAR *) &wMsgPtr,
                            0, NULL);
    if (length > 0) {
        /* Need a TwapiAlloc'ed buffer and not a LocalAlloc'ed one */
        WCHAR *temp = wMsgPtr;
        wMsgPtr = TwapiAllocWString(wMsgPtr, length);
        LocalFree(temp);
        goto done;
    }

    length = FormatMessageA(flags, hModule, error,
                            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                            (char *) &msgPtr,
                            0, NULL);
    if (length > 0) {
        wMsgPtr = (WCHAR *) TwapiAlloc((length + 1) * sizeof(WCHAR));
        if (wMsgPtr) {
            wMsgPtr[0] = 0;
            MultiByteToWideChar(CP_ACP, 0, msgPtr, length + 1, wMsgPtr,
                                length + 1);
            LocalFree(msgPtr);
        }
    }

done:
    if (wMsgPtr) {
        /* length does not include terminating \0 */
	/* Trim the trailing CR/LF from the system message. */
        if (length > 0) {
            if (wMsgPtr[length-1] == L'\n')
                --length;
        }
        if (length > 0) {
            if (wMsgPtr[length-1] == L'\r')
                --length;
        }
        wMsgPtr[length] = 0;
    }

    return wMsgPtr;
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
        Tcl_Obj *resultObj = Tcl_GetObjResult(interp);
        Tcl_AppendUnicodeToObj(resultObj, L" ", 1);
        Tcl_AppendUnicodeToObj(resultObj, provider, -1);
        Tcl_AppendUnicodeToObj(resultObj, L": ", 2);
        Tcl_AppendUnicodeToObj(resultObj, errorbuf, -1);
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
