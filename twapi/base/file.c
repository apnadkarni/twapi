/*
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"


int Twapi_GetFileType(Tcl_Interp *interp, HANDLE h)
{
    DWORD file_type = GetFileType(h);
    if (file_type == FILE_TYPE_UNKNOWN) {
        /* Is it really an error ? */
        DWORD winerr = GetLastError();
        if (winerr != NO_ERROR) {
            /* Yes it is */
            return Twapi_AppendSystemError(interp, winerr);
        }
    }
    Tcl_SetObjResult(interp, Tcl_NewLongObj(file_type));
    return TCL_OK;
}



