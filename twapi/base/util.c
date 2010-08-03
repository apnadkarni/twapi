/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>


/* Define glob matching functions that fit lstrcmp prototype - return
   0 if match, 1 if no match */
int WINAPI TwapiGlobCmp (const char *s, const char *pat)
{
    return ! Tcl_StringCaseMatch(s, pat, 0);
}
int WINAPI TwapiGlobCmpCase (const char *s, const char *pat)
{
    return ! Tcl_StringCaseMatch(s, pat, 1);
}


/* Return the DLL version of the given dll. The version is returned as
 * 0 if the DLL does not support the given version
 */
void TwapiGetDllVersion(char *dll, DLLVERSIONINFO *verP)
{
    FARPROC fn;
    HINSTANCE h;

    verP->cbSize = sizeof(DLLVERSIONINFO);
    verP->dwMajorVersion = 0;
    verP->dwMinorVersion = 0;
    verP->dwBuildNumber = 0;
    verP->dwPlatformID = DLLVER_PLATFORM_NT;

    h = LoadLibraryA(dll);
    if (h == NULL)
        return;

    fn = (FARPROC) GetProcAddress(h, "DllGetVersion");
    if (fn)
        (void) (*fn)(verP);

    FreeLibrary(h);
}


void DebugOutput(char *s) {
    Tcl_Channel chan = Tcl_GetStdChannel(TCL_STDERR);
    Tcl_WriteChars(chan, s, -1);
    Tcl_Flush(chan);
}


int TwapiReadMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int restype;
    char *p;
    int offset;
    int len;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(restype), GETVOIDP(p), GETINT(offset),
                     ARGUSEDEFAULT, GETINT(len),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    p += offset;
    // Note: restype matches func dispatch code in Twapi_CallObjCmd
    switch (restype) {
    case 10101:
        result.type = TRT_DWORD;
        result.value.ival = *(int UNALIGNED *) p;
        break;

    case 10102:
        result.type = TRT_BINARY;
        result.value.binary.p = p;
        result.value.binary.len = len;
        break;

    case 10103:
        result.type = TRT_CHARS;
        result.value.chars.str = p;
        result.value.chars.len = len;
        break;
        
    case 10104:
        result.type = TRT_UNICODE;
        result.value.unicode.str = (WCHAR *) p;
        result.value.unicode.len =  (len == -1) ? -1 : (len / sizeof(WCHAR));
        break;
        
    case 10105:
        result.type = TRT_LPVOID;
        result.value.pv =  *(void* *) p;
        break;

    case 10106:
        result.type = TRT_WIDE;
        result.value.wide = *(Tcl_WideInt UNALIGNED *)p;
        break;

    default:
        Tcl_SetResult(interp, "Unknown result type passed to TwapiReadMemory", TCL_STATIC);
        return TCL_ERROR;
    }

    return TwapiSetResult(interp, &result);
}


int TwapiWriteMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
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

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(func), GETVOIDP(bufP), GETDWORD_PTR(offset),
                     GETDWORD_PTR(buf_size), ARGSKIP,
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* Note the ARGSKIP ensures the objv[] in that position exists */
    switch (func) {
    case 10111: // Int
        if ((offset + sizeof(int)) > buf_size)
            goto overrun;
        if (Tcl_GetIntFromObj(interp, objv[4], &val) != TCL_OK)
            return TCL_ERROR;
        *(int UNALIGNED *)(offset + bufP) = val;
        break;
    case 10112: // Binary
        cp = Tcl_GetByteArrayFromObj(objv[4], &sz);
        if ((offset + sz) > buf_size)
            goto overrun;
        CopyMemory(offset + bufP, cp, sz);
        break;
    case 10113: // Chars
        cp = Tcl_GetStringFromObj(objv[4], &sz);
        /* Note we also include the terminating null */
        if ((offset + sz + 1) > buf_size)
            goto overrun;
        CopyMemory(offset + bufP, cp, sz+1);
        break;
    case 10114: // Unicode
        wp = Tcl_GetUnicodeFromObj(objv[4], &sz);
        /* Note we also include the terminating null */
        if ((offset + (sizeof(WCHAR)*(sz + 1))) > buf_size)
            goto overrun;
        CopyMemory(offset + bufP, wp, sizeof(WCHAR)*(sz+1));
        break;
    case 10115: // Pointer
        if ((offset + sizeof(void*)) > buf_size)
            goto overrun;
        if (ObjToLPVOID(interp, objv[4], &pv) != TCL_OK)
            return TCL_ERROR;
        CopyMemory(offset + bufP, &pv, sizeof(void*));
        break;
    case 10116:
        if ((offset + sizeof(Tcl_WideInt)) > buf_size)
            goto overrun;
        if (Tcl_GetWideIntFromObj(interp, objv[4], &wide) != TCL_OK)
            return TCL_ERROR;
        *(Tcl_WideInt UNALIGNED *)(offset + bufP) = wide;
        break;
    }        

    return TCL_OK;

overrun:
    return TwapiReturnTwapiError(interp, "Buffer too small.",
                                 TWAPI_BUFFER_OVERRUN);
}


WCHAR *TwapiAllocWString(WCHAR *src, int len)
{
    WCHAR *dst;
    if (len < 0) {
        len = lstrlenW(src);
    }
    dst = TwapiAlloc(sizeof(WCHAR) * (len+1));
    CopyMemory(dst, src, sizeof(WCHAR)*len);
    dst[len] = 0; /* Source string may not have been terminated after len chars */
    return dst;
}

WCHAR *TwapiAllocWStringFromObj(Tcl_Obj *objP, int *lenP) {
    WCHAR *wP;
    int len;
    if (lenP == NULL)
        lenP = &len;
    wP = Tcl_GetUnicodeFromObj(objP, lenP);
    return TwapiAllocWString(wP, *lenP);
}

char *TwapiAllocAString(char *src, int len)
{
    char *dst;
    if (len < 0) {
        len = lstrlenA(src);
    }
    dst = TwapiAlloc(len+1);
    CopyMemory(dst, src, len);
    dst[len] = 0; /* Source string may not have been terminated after len chars */
    return dst;
}

char *TwapiAllocAStringFromObj(Tcl_Obj *objP, int *lenP) {
    char *cP;
    int len;
    if (lenP == NULL)
        lenP = &len;
    cP = Tcl_GetStringFromObj(objP, lenP);
    return TwapiAllocAString(cP, *lenP);
}

void *TwapiAlloc(size_t sz)
{
    void *p = HeapAlloc(GetProcessHeap(), 0, sz); /* TBD */
    if (p == NULL)
        Tcl_Panic("Could not allocate %d bytes.", sz);
    return p;
}

void *TwapiAllocSize(size_t sz, size_t *actual_sizeP)
{
    void *p = TwapiAlloc(sz);
    if (actual_sizeP)
        *actual_sizeP = HeapSize(GetProcessHeap(), 0, p);
    return p;
}

void *TwapiReallocTry(void *p, size_t sz)
{
    return HeapReAlloc(GetProcessHeap(), 0, p, sz);
}

void *TwapiAllocZero(size_t sz)
{
    void *p = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sz); /* TBD */
    if (p == NULL)
        Tcl_Panic("Could not allocate %d bytes.", sz);
    return p;
}


int TwapiDoOneTimeInit(TwapiOneTimeInitState *stateP, TwapiOneTimeInitFn *fn, ClientData clientdata)
{
    TwapiOneTimeInitState prev_state;

    /* Init unless already done. */
    switch (InterlockedCompareExchange(stateP,
                                       TWAPI_INITSTATE_IN_PROGRESS,
                                       TWAPI_INITSTATE_NOT_DONE)) {
    case TWAPI_INITSTATE_DONE:
        return 3;               /* Already done */

    case TWAPI_INITSTATE_IN_PROGRESS:
        while (1) {
            prev_state = InterlockedCompareExchange(stateP,
                                                    TWAPI_INITSTATE_IN_PROGRESS,
                                                    TWAPI_INITSTATE_IN_PROGRESS);
            if (prev_state == TWAPI_INITSTATE_DONE)
                return 2;       /* Done after waiting */

            if (prev_state != TWAPI_INITSTATE_IN_PROGRESS)
                return 0; /* Error but do not know what - someone else was
                             initializing */

            /*
             * Someone is initializing, wait in a spin
             * Note the Sleep() will yield to other threads, including
             * the one doing the init, so this is not a hard loop
             */
            Sleep(1);
        }
        break;

    case TWAPI_INITSTATE_NOT_DONE:
        /* We need to do the init */
        if (fn(clientdata) != TCL_OK) {
            InterlockedExchange(stateP, TWAPI_INITSTATE_ERROR);
            return 0;
        }
        InterlockedExchange(stateP, TWAPI_INITSTATE_DONE);
        return 1;               /* We init'ed successfully */

    case TWAPI_INITSTATE_ERROR:
        /* State was already in error. No way to recover safely 
           See comments about not calling Tcl_SetResult to set error */
        return 0;
    }
    
    return 0;
}

/* 
 * Invokes a Tcl command from objv[] and updates the passed 
 * TwapiPendingCallback structure result. Note the following:
 *  - the objv[] array generally must not be accessed again by caller
 *    unless they previously do a Tcl_IncrRef on each element of the array
 *  - the pcbP->status and pcbP->response fields are updated.
 *  - the pcbP->response field type will not always match response_type
 *    unless the function returns TCL_OK
 *
 * Return values are as follows:
 * TCL_ERROR - error - you called me with bad arguments, dummy.
 *   cbP is unchanged.
 *   Do what you will or let the framework return an error.
 * TCL_OK - either everything went fine, or there was an error in script
 *   evaluation. Distinguish
 *   by checking the cbP->status field. Either way cpP->response has
 *   been set appropriately.
 */
int TwapiEvalAndUpdateCallback(TwapiCallback *cbP, int objc, Tcl_Obj *objv[], TwapiResultType response_type)
{
    int i;
    Tcl_Obj *objP;
    int tcl_status;
    TwapiResult *responseP = &cbP->response;

    /*
     * Before we eval, make sure stuff does not disappear in the eval. We
     * do this even if the interp is deleted since we have to call decrref
     * on exit
     */
    for (i = 0; i < objc; ++i) {
        Tcl_IncrRefCount(objv[i]);
    }

    if (cbP->ticP->interp == NULL ||
        Tcl_InterpDeleted(cbP->ticP->interp)) {
        cbP->winerr = ERROR_DS_NO_SUCH_OBJECT; /* Best match we can find */
        cbP->response.type = TRT_EMPTY;
        tcl_status = TCL_OK;         /* See function comments */
        goto vamoose;
    }

    /* Preserve structures during eval */
    TwapiInterpContextRef(cbP->ticP, 1);
    Tcl_Preserve(cbP->ticP->interp);

    tcl_status = Tcl_EvalObjv(cbP->ticP->interp, objc, objv, 
                               TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);
    if (tcl_status != TCL_OK) {
        /* TBD - see if errorcode is TWAPI_WIN32 and return that if so */
        cbP->winerr = ERROR_FUNCTION_FAILED;
        response_type = TRT_CHARS_DYNAMIC;
        tcl_status = TCL_OK; /* See function comments */
    } else {
        cbP->winerr = ERROR_SUCCESS;
    }
    
    objP = Tcl_GetObjResult(cbP->ticP->interp); /* OK even if interp
                                                   logically deleted */
     /* 
      * If there are any errors in conversion, we will return TCL_ERROR
      * as caller should know the type of result to expect
      */
    switch (response_type) {
    case TRT_CHARS_DYNAMIC:
        responseP->value.chars.str =
            TwapiAllocAStringFromObj(
                objP, &responseP->value.chars.len);
        break;
    case TRT_UNICODE_DYNAMIC:
        responseP->value.unicode.str =
            TwapiAllocWStringFromObj(
                objP, &responseP->value.unicode.len);
        break;
    case TRT_DWORD:
        tcl_status = Tcl_GetLongFromObj(cbP->ticP->interp, objP,
                                        &responseP->value.ival);
        /* Errors will be handled below */
        break;
    case TRT_BOOL:
        tcl_status = Tcl_GetBooleanFromObj(cbP->ticP->interp, objP,
                                           &responseP->value.bval);
        /* Errors will be handled below */
        break;
    case TRT_EMPTY:
        break;                  /* Empty string */
    default:
        tcl_status = TCL_ERROR;
        cbP->winerr = ERROR_INVALID_PARAMETER;
        break;
    }

    if (tcl_status == TCL_OK)
        cbP->response.type = response_type;
    else
        cbP->response.type = TRT_EMPTY;

    Tcl_ResetResult(cbP->ticP->interp);/* Don't leave crud from eval */
    Tcl_Release(cbP->ticP->interp);
    TwapiInterpContextUnref(cbP->ticP, 1);

vamoose:
    /* 
     * Undo the incrref above. This will delete the object unless
     * caller had done an incr-ref on it.
     */
    for (i = 0; i < objc; ++i) {
        Tcl_DecrRefCount(objv[i]);
    }

    return tcl_status;
}

void TwapiEnqueueTclEvent(TwapiInterpContext *ticP, Tcl_Event *evP)
{
    Tcl_ThreadQueueEvent(ticP->thread, evP, TCL_QUEUE_TAIL);
    Tcl_ThreadAlert(ticP->thread); /* Wake up the thread */
}


int Twapi_AppendLog(TwapiInterpContext *ticP, WCHAR *msg)
{
    Tcl_Obj *var;
    int limit, len;
    Tcl_Interp *interp = ticP->interp;
    Tcl_Obj *msgObj;

    /* Check if the log variable exists. If not logging is disabled */
    var = Tcl_GetVar2Ex(interp, TWAPI_SETTINGS_VAR, "log_limit", 0);
    if (var == NULL)
        return TCL_OK;          /* Logging not enabled */


    if (Tcl_GetIntFromObj(interp, var, &limit) != TCL_OK ||
        limit == 0) {
        Tcl_ResetResult(interp); /* GetInt might have stuck error message */
        return TCL_OK;          /* Logging not enabled */
    }

    msgObj = Tcl_NewUnicodeObj(msg, -1);
    var = Tcl_GetVar2Ex(interp, TWAPI_LOG_VAR, NULL, 0);
    if (var) {
        if (Tcl_ListObjLength(interp, var, &len) != TCL_OK) {
            /* Not a list. Some error, blow it all away. */
            var = Tcl_NewObj();
            Tcl_SetVar2Ex(interp, TWAPI_LOG_VAR, NULL, var, 0);
            len = 0;
        } else {
            if (Tcl_IsShared(var)) {
                var = Tcl_DuplicateObj(var);
                Tcl_SetVar2Ex(interp, TWAPI_LOG_VAR, NULL, var, 0);
            }
        }
        TWAPI_ASSERT(! Tcl_IsShared(var));
        if (len >= limit) {
            /* Remove elements from front of list */
            Tcl_ListObjReplace(interp, var, 0, len-limit+1, 0, NULL);
        }
        Tcl_ListObjAppendElement(interp, var, msgObj);
    } else {
        /* log variable is currently not set */
        Tcl_SetVar2Ex(interp, TWAPI_LOG_VAR, NULL, Tcl_NewListObj(1, &msgObj), 0);
    }
    return TCL_OK;
}
