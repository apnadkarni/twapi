/* 
 * Copyright (c) 2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include <wbemidl.h>

TCL_RESULT Twapi_IMofCompiler_CompileFileOrBuffer(TwapiInterpContext *ticP, int type, int objc, Tcl_Obj *CONST objv[])
{
    IMofCompiler *ifc;
    WCHAR *server_namespace;
    WCHAR *user;
    WCHAR *password;
    WCHAR *authority;
    int    optflags, classflags, instflags;
    BYTE  *buf;
    int    buflen;
    HRESULT hr;
    WBEM_COMPILE_STATUS_INFO wcsi;

    if (type == 2) {
        if (TwapiGetArgs(ticP->interp, objc, objv,
                         GETPTR(ifc, IMofCompiler), ARGSKIP, ARGSKIP,
                         ARGUSEDEFAULT,
                         GETNULLIFEMPTY(server_namespace),
                         GETINT(optflags), GETINT(classflags),
                         GETINT(instflags), ARGEND) != TCL_OK)
            return TCL_ERROR;
    } else {
        if (TwapiGetArgs(ticP->interp, objc, objv,
                         GETPTR(ifc, IMofCompiler), ARGSKIP, ARGUSEDEFAULT,
                         GETNULLIFEMPTY(server_namespace), GETNULLIFEMPTY(user),
                         GETNULLIFEMPTY(authority), GETNULLIFEMPTY(password),
                         GETINT(optflags), GETINT(classflags),
                         GETINT(instflags), ARGEND) != TCL_OK)
            return TCL_ERROR;
    }            

    ZeroMemory(&wcsi, sizeof(wcsi));

    switch (type) {
    case 0:
        buf = Tcl_GetStringFromObj(objv[1], &buflen);
        hr = ifc->lpVtbl->CompileBuffer(ifc, buflen, buf, server_namespace,
                                         user, authority, password,
                                         optflags, classflags, instflags,
                                         &wcsi);
        break;
    case 1:
        hr = ifc->lpVtbl->CompileFile(ifc, Tcl_GetUnicode(objv[1]),
                                       server_namespace,
                                       user, authority, password,
                                       optflags, classflags, instflags,
                                       &wcsi);
        break;

    case 2:
        hr = ifc->lpVtbl->CreateBMOF(ifc, Tcl_GetUnicode(objv[1]),
                                     Tcl_GetUnicode(objv[2]),
                                     server_namespace,
                                     optflags, classflags, instflags,
                                     &wcsi);
        break;

    default:
        return TwapiReturnTwapiError(ticP->interp, "Invalid IMofCompiler function code", TWAPI_BUG);
    }

    switch (hr) {
    case 2: /* Warning if autorecover pragma is not present. Not an error */
    case WBEM_S_NO_ERROR:
        return TCL_OK;

    WBEM_S_FALSE: /* Fall thru */
    default:
        Tcl_SetObjResult(ticP->interp,
                         Tcl_ObjPrintf("IMofCompiler error: phase: %d, object number: %d, first line: %d, last line: %d.",
                                       wcsi.lPhaseError,
                                       wcsi.ObjectNum,
                                       wcsi.FirstLine,
                                       wcsi.LastLine));
        return Twapi_AppendSystemError(ticP->interp, wcsi.hRes ? wcsi.hRes : hr);
    }
}                     

