/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>


/* Define glob matching functions that fit strcmp prototype - return
   0 if match, 1 if no match */
int TwapiGlobCmp (const char *s, const char *pat)
{
    return ! Tcl_StringCaseMatch(s, pat, 0);
}
int TwapiGlobCmpCase (const char *s, const char *pat)
{
    return ! Tcl_StringCaseMatch(s, pat, 1);
}

/*
 * Allocates memory. On failure, sets interp result to out of memory error.
 * ALso sets Windows last error code appropriately
 * Returns TCL_OK/TCL_ERROR
 */
int Twapi_malloc(
    Tcl_Interp *interp,
    char *msg,                  /* Addition error msg. May be NULL */
    size_t size,
    void **pp)
{
    char buf[80];

    /* This *must* use malloc, nothing else like Tcl_Alloc because caller
       expects to free()
    */
    *pp = malloc(size);
    if (*pp)
        return TCL_OK;

    /* Failed to allocate memory */
    if (interp) {
        _snprintf(buf, ARRAYSIZE(buf), "Failed to allocate %d bytes. ", size);
        /* Note it's OK if msg is NULL */
        Tcl_AppendResult(interp, buf, msg, NULL);
        /*
         * Set the error code. This might be overkill if we are already in
         * low memory conditions but the assumption is the failure is because
         * size is too big, not because memory is really finished
         */
        Twapi_AppendSystemError(interp, ERROR_OUTOFMEMORY);
    }
    SetLastError(E_OUTOFMEMORY);
    return TCL_ERROR;
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
        result.type = TRT_ADDRESS_LITERAL;
        result.value.pval =  *(void* *) p;
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
        memmove(offset + bufP, cp, sz);
        break;
    case 10113: // Chars
        cp = Tcl_GetStringFromObj(objv[4], &sz);
        /* Note we also include the terminating null */
        if ((offset + sz + 1) > buf_size)
            goto overrun;
        memmove(offset + bufP, cp, sz+1);
        break;
    case 10114: // Unicode
        wp = Tcl_GetUnicodeFromObj(objv[4], &sz);
        /* Note we also include the terminating null */
        if ((offset + (sizeof(WCHAR)*(sz + 1))) > buf_size)
            goto overrun;
        memmove(offset + bufP, wp, sizeof(WCHAR)*(sz+1));
        break;
    case 10115: // Pointer
        if ((offset + sizeof(void*)) > buf_size)
            goto overrun;
        if (ObjToLPVOID(interp, objv[4], &pv) != TCL_OK)
            return TCL_ERROR;
        memmove(offset + bufP, &pv, sizeof(void*));
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
    MoveMemory(dst, src, sizeof(WCHAR)*len);
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
    MoveMemory(dst, src, len);
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
    void *p = malloc(sz);
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
 *  - assumes cbP->ticP->interp is valid
 *  - the pcbP->response field type will not always match response_type
 *    unless the function returns TCL_OK
 *
 * Return values are as follows:
 * TCL_ERROR - error - you called me with bad arguments, dummy.
 *   cbP is unchanged except its status is set to ERROR_INVALID_ARGS
 *   and cbP->response is set to TRT_EMPTY.
 *   Do what you will.
 * TCL_OK - either everything went fine, or there was an error. Distinguish
 *   by checking the cbP->status field. Either way cpP->response has
 *   been set appropriately.
 */
int TwapiEvalAndUpdateCallback(TwapiPendingCallback *cbP, int objc, Tcl_Obj *objv[], TwapiResultType response_type)
{
    int i;
    Tcl_Obj *objP;
    int tcl_status;
    TwapiResult *responseP = &cbP->response;

    if (cbP->ticP->interp == NULL ||
        Tcl_InterpDeleted(cbP->ticP->interp)) {
        cbP->status = ERROR_DS_NO_SUCH_OBJECT; /* Best match we can find */
        cbP->response.type = TRT_EMPTY;
        return TCL_OK;          /* See function comments */
    }

    /* Before we eval, make sure stuff does not disappear in the eval */
    for (i = 0; i < objc; ++i) {
        Tcl_IncrRefCount(objv[i]);
    }
    TwapiInterpContextRef(cbP->ticP, 1);
    Tcl_Preserve(cbP->ticP->interp);

    tcl_status = Tcl_EvalObjv(cbP->ticP->interp, objc, objv, 
                               TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);
    if (tcl_status != TCL_OK) {
        /* TBD - see if errorcode is TWAPI_WIN32 and return that if so */
        cbP->status = E_FAIL;
        response_type = TRT_CHARS_DYNAMIC;
        tcl_status = TCL_OK; /* See function comments */
    } else {
        cbP->status = ERROR_SUCCESS;
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
        cbP->status = ERROR_INVALID_PARAMETER;
        break;
    }

    if (tcl_status == TCL_OK)
        cbP->response.type = response_type;
    else
        cbP->response.type = TRT_EMPTY;

    Tcl_ResetResult(cbP->ticP->interp);/* Don't leave crud from eval */
    Tcl_Release(cbP->ticP->interp);
    TwapiInterpContextUnref(cbP->ticP, 1);

    /* 
     * Undo the incrref above. This will delete the object unless
     * caller had done an incr-ref on it.
     */
    for (i = 0; i < objc; ++i) {
        Tcl_DecrRefCount(objv[i]);
    }

    return tcl_status;
}
