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


