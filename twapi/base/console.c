/* 
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

int Twapi_ReadConsole(Tcl_Interp *interp, HANDLE conh, unsigned int numchars)
{
    WCHAR  buf[256];
    WCHAR *bufP = buf;
    DWORD  len;
    int status = TCL_ERROR;

    if (numchars > ARRAYSIZE(buf)) {
        if (Twapi_malloc(interp, NULL, sizeof(WCHAR) * numchars, &bufP) != TCL_OK)
            return TCL_ERROR;
    }

    if (! ReadConsoleW(conh, bufP, numchars, &len, NULL)) {
        TwapiReturnSystemError(interp);
        goto vamoose;
    }

    Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(buf, len));
    status = TCL_OK;

vamoose:
    if (bufP != buf)
        free(bufP);
    return status;
}
