/*
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

int Twapi_EnumClipboardFormats(Tcl_Interp *interp)
{
    UINT clip_fmt;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    Tcl_SetObjResult(interp, resultObj);

    clip_fmt = 0;
    while (1) {
        clip_fmt = EnumClipboardFormats(clip_fmt);
        if (clip_fmt)
            Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewIntObj(clip_fmt));
        else {
            DWORD error = GetLastError();
            if (error != ERROR_SUCCESS) {
                return TwapiReturnSystemError(interp);
            }
            break;
        }
    }
    return TCL_OK;
}


