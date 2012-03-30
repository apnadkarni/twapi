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

/* Return a Tcl_Obj that is a lower case version of passed object */
Tcl_Obj *TwapiLowerCaseObj(Tcl_Obj *objP)
{
    Tcl_Obj *resultObj;
    char *cP;
    int len;

    cP = Tcl_GetStringFromObj(objP, &len);
    resultObj = Tcl_NewStringObj(cP, len);
    len = Tcl_UtfToLower(Tcl_GetString(resultObj));
    Tcl_SetObjLength(resultObj, len);
    return resultObj;
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


void TwapiDebugOutput(char *s) {
    Tcl_Channel chan = Tcl_GetStdChannel(TCL_STDERR);
    Tcl_Obj *objP;
    /* We want a terminating newline, but it has to go in a single write
     * so we cannot do two Tcl_WriteObj (output from multiple threads
     * gets jumbled - actual experience. So need a temp buffer.
     */
    objP = Tcl_ObjPrintf("%s\n", s);
    Tcl_IncrRefCount(objP);
    Tcl_WriteObj(chan, objP);
    Tcl_Flush(chan);
    Tcl_DecrRefCount(objP);
}


int TwapiReadMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    char *p;
    WCHAR *wp;
    int modifier;
    int restype;
    int offset;
    int len;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(restype), GETVOIDP(p), GETINT(offset),
                     ARGUSEDEFAULT, GETINT(len), GETINT(modifier),
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
        /* When len > 0, modifier != 0 means also check null terminator */
        if (len > 0 && modifier) {
            for (modifier = 0; modifier < len && p[modifier]; ++modifier)
                ;
            len = modifier;
        }
        result.value.chars.str = p;
        result.value.chars.len = len;
        result.type = TRT_CHARS;
        break;
        
    case 10104:
        if (len > 0)
            len /= sizeof(WCHAR); /* Convert to num chars */
        wp = (WCHAR *) p;

        /* When len > 0, modifier != 0 means also check null terminator */
        if (len > 0 && modifier) {
            for (modifier = 0; modifier < len && wp[modifier]; ++modifier)
                ;
            len = modifier;
        }
        result.type = TRT_UNICODE;
        result.value.unicode.str = wp;
        result.value.unicode.len =  len;
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
    return TwapiReturnError(interp, TWAPI_BUFFER_OVERRUN);
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

void TwapiFree(void *p)
{
    HeapFree(GetProcessHeap(), 0, p);
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

    InitTwapiResult(responseP);

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
    case TRT_LONG:
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
        TwapiClearResult(&cbP->response);

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


int Twapi_AppendLog(Tcl_Interp *interp, WCHAR *msg)
{
    Tcl_Obj *var;
    int limit, len;
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

    msgObj = ObjFromUnicode(msg);
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

/* Sets interp result as handle if not null else GetLastError value */
TCL_RESULT TwapiReturnNonnullHandle(Tcl_Interp *interp, HANDLE h, char *typestr)
{
    if (h == NULL)
        return TwapiReturnSystemError(interp);

    if (typestr == NULL)
        typestr = "HANDLE";

    Tcl_SetObjResult(interp, ObjFromOpaque(h, typestr));
    return TCL_OK;
}


#ifdef TWAPI_REPLACE_CRT
/* 
 * strtoul.c --
 *
 *	Source code for the "strtoul" library procedure.
 *
 * Copyright (c) 1988 The Regents of the University of California.
 * Copyright (c) 1994 Sun Microsystems, Inc.
 *
 * See the file "license.terms" for information on usage and redistribution of
 * this file, and for a DISCLAIMER OF ALL WARRANTIES.
 *
 */

/*
 * The table below is used to convert from ASCII digits to a numerical
 * equivalent. It maps from '0' through 'z' to integers (100 for non-digit
 * characters).
 */

static char cvtIn[] = {
    0, 1, 2, 3, 4, 5, 6, 7, 8, 9,		/* '0' - '9' */
    100, 100, 100, 100, 100, 100, 100,		/* punctuation */
    10, 11, 12, 13, 14, 15, 16, 17, 18, 19,	/* 'A' - 'Z' */
    20, 21, 22, 23, 24, 25, 26, 27, 28, 29,
    30, 31, 32, 33, 34, 35,
    100, 100, 100, 100, 100, 100,		/* punctuation */
    10, 11, 12, 13, 14, 15, 16, 17, 18, 19,	/* 'a' - 'z' */
    20, 21, 22, 23, 24, 25, 26, 27, 28, 29,
    30, 31, 32, 33, 34, 35};

/*
 *----------------------------------------------------------------------
 *
 * strtoul --
 *
 *	Convert an ASCII string into an integer.
 *
 * Results:
 *	The return value is the integer equivalent of string. If endPtr is
 *	non-NULL, then *endPtr is filled in with the character after the last
 *	one that was part of the integer. If string doesn't contain a valid
 *	integer value, then zero is returned and *endPtr is set to string.
 *
 * Side effects:
 *	None.
 *
 *----------------------------------------------------------------------
 */

unsigned long int
strtoul(
    CONST char *string,		/* String of ASCII digits, possibly preceded
				 * by white space. For bases greater than 10,
				 * either lower- or upper-case digits may be
				 * used. */
    char **endPtr,		/* Where to store address of terminating
				 * character, or NULL. */
    int base)			/* Base for conversion.  Must be less than 37.
				 * If 0, then the base is chosen from the
				 * leading characters of string: "0x" means
				 * hex, "0" means octal, anything else means
				 * decimal. */
{
    register CONST char *p;
    register unsigned long int result = 0;
    register unsigned digit;
    int anyDigits = 0;
    int overflow=0;

    /*
     * Skip any leading blanks.
     */

    p = string;
    while (((unsigned char) *p) == ' ') {
	p += 1;
    }
    if (*p == '+') {
        p += 1;
    }

    /*
     * If no base was provided, pick one from the leading characters of the
     * string.
     */
    
    if (base == 0) {
	if (*p == '0') {
	    p += 1;
	    if ((*p == 'x') || (*p == 'X')) {
		p += 1;
		base = 16;
	    } else {
		/*
		 * Must set anyDigits here, otherwise "0" produces a "no
		 * digits" error.
		 */

		anyDigits = 1;
		base = 8;
	    }
	} else {
	    base = 10;
	}
    } else if (base == 16) {
	/*
	 * Skip a leading "0x" from hex numbers.
	 */

	if ((p[0] == '0') && ((p[1] == 'x') || (p[1] == 'X'))) {
	    p += 2;
	}
    }

    /*
     * Sorry this code is so messy, but speed seems important. Do different
     * things for base 8, 10, 16, and other.
     */

    if (base == 8) {
	unsigned long maxres = ULONG_MAX >> 3;

	for ( ; ; p += 1) {
	    digit = *p - '0';
	    if (digit > 7) {
		break;
	    }
	    if (result > maxres) { overflow = 1; }
	    result = (result << 3);
	    if (digit > (ULONG_MAX - result)) { overflow = 1; }
	    result += digit;
	    anyDigits = 1;
	}
    } else if (base == 10) {
	unsigned long maxres = ULONG_MAX / 10;

	for ( ; ; p += 1) {
	    digit = *p - '0';
	    if (digit > 9) {
		break;
	    }
	    if (result > maxres) { overflow = 1; }
	    result *= 10;
	    if (digit > (ULONG_MAX - result)) { overflow = 1; }
	    result += digit;
	    anyDigits = 1;
	}
    } else if (base == 16) {
	unsigned long maxres = ULONG_MAX >> 4;

	for ( ; ; p += 1) {
	    digit = *p - '0';
	    if (digit > ('z' - '0')) {
		break;
	    }
	    digit = cvtIn[digit];
	    if (digit > 15) {
		break;
	    }
	    if (result > maxres) { overflow = 1; }
	    result = (result << 4);
	    if (digit > (ULONG_MAX - result)) { overflow = 1; }
	    result += digit;
	    anyDigits = 1;
	}
    } else if (base >= 2 && base <= 36) {
	unsigned long maxres = ULONG_MAX / base;

	for ( ; ; p += 1) {
	    digit = *p - '0';
	    if (digit > ('z' - '0')) {
		break;
	    }
	    digit = cvtIn[digit];
	    if (digit >= ( (unsigned) base )) {
		break;
	    }
	    if (result > maxres) { overflow = 1; }
	    result *= base;
	    if (digit > (ULONG_MAX - result)) { overflow = 1; }
	    result += digit;
	    anyDigits = 1;
	}
    }

    /*
     * See if there were any digits at all.
     */

    if (!anyDigits) {
	p = string;
    }

    if (endPtr != 0) {
	/* unsafe, but required by the strtoul prototype */
	*endPtr = (char *) p;
    }

    if (overflow) {
	// NO CRTL errno = ERANGE;
	return ULONG_MAX;
    } 

    return result;
}

#endif


#if 0 // MS C does not like intrinsics redefined
#ifdef TWAPI_REPLACE_CRT
void *memset(
   void* dest, 
   int c, 
   size_t count 
)
{
    __stosb((unsigned char *)dest, (unsigned char *) c, count);
    return dest;
}


#endif
#endif
