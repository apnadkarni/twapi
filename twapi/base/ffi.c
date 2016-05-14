/*
 * Copyright (c) 2010-2014, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

TCL_RESULT Twapi_FfiLoadObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    Tcl_Obj *objP;
    TwapiTls *tlsP;
    DWORD dw;

    CHECK_NARGS(interp, objc, 2);
    tlsP = Twapi_GetTls();
    --objc;
    ++objv;
    if (ObjDictGet(interp, tlsP->ffiObj, objv[0], &objP) != TCL_OK)
        return TCL_ERROR;
    if (objP == NULL) {
        /* Entry does not exist. Check if the DLL handle is there */
        Tcl_Obj **dll_and_func;
        Tcl_Obj *dllObj;
        HMODULE  hdll = NULL;
        FARPROC fn;
        if (ObjGetElements(NULL, objv[0], &dw, &dll_and_func) != TCL_OK ||
            dw != 2) {
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid DLL and function specification");
        }
        if (ObjDictGet(interp, tlsP->ffiObj, dll_and_func[0], &dllObj) != TCL_OK)
            return TCL_ERROR;
        if (dllObj == NULL) {
            hdll = LoadLibraryW(ObjToWinChars(dll_and_func[0]));
            ObjDictPut(interp, tlsP->ffiObj, dll_and_func[0], ObjFromHMODULE(hdll)); /* May be NULL */
        } else {
            if (ObjToHMODULE(interp, dllObj, &hdll) != TCL_OK)
                return TCL_ERROR;
        }
        if (hdll) {
            /* Have DLL, get the function address */
            fn = GetProcAddress(hdll, ObjToString(dll_and_func[1]));
        } else
            fn = 0;
        objP = ObjFromFARPROC(fn);
        ObjDictPut(NULL, tlsP->ffiObj, objv[0], objP);
    }
    ObjSetResult(interp, objP);
    return TCL_OK;
}

TCL_RESULT Twapi_Ffi0ObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    FARPROC fn;
    Tcl_WideInt ret;
    DWORD winerr;
    Tcl_Obj *objs[2];
    
    CHECK_NARGS(interp, objc, 2);

    /* ffi0 PROCADDR */
    if (ObjToFARPROC(interp, objv[1], &fn) != TCL_OK)
        return TCL_ERROR;

    TWAPI_ASSERT(fn);
    ret = (Tcl_WideInt) fn();
    winerr = GetLastError();
    objs[0] = ObjFromWideInt(ret);
    objs[1] = ObjFromDWORD(winerr);
    ObjSetResult(interp, ObjNewList(2, objs));
    return TCL_OK;
}

TCL_RESULT Twapi_FfiHObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    FARPROC fn;
    Tcl_WideInt ret;
    DWORD winerr;
    HANDLE h;
    Tcl_Obj *objs[2];
    
    CHECK_NARGS(interp, objc, 3);

    /* ffi0 PROCADDR RETURNTYPE */
    if (ObjToFARPROC(interp, objv[1], &fn) != TCL_OK ||
        ObjToHANDLE(interp, objv[2], &h) != TCL_OK)
        return TCL_ERROR;

    TWAPI_ASSERT(fn);
    ret = (Tcl_WideInt) fn(h);
    winerr = GetLastError();
    objs[0] = ObjFromWideInt(ret);
    objs[1] = ObjFromDWORD(winerr);
    ObjSetResult(interp, ObjNewList(2, objs));
    return TCL_OK;
}
