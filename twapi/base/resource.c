/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

/* Note: Param 'interp' is there to match prototype expected by TwapiGetArgs */
TCL_RESULT ObjToResourceIntOrString(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *wsP)
{
    int i;

    /*
     * Resource type and name can be integers or strings. If integer,
     * they must pass the IS_INTRESOURCE test as not all integers are valid.
     */
    if (Tcl_GetLongFromObj(NULL, objP, &i) == TCL_OK &&
        IS_INTRESOURCE(i))
        *wsP = MAKEINTRESOURCEW(i);
    else
        *wsP = Tcl_GetUnicode(objP);

    return TCL_OK;
}

Tcl_Obj *ObjFromResourceIntOrString(LPCWSTR s)
{
    if (IS_INTRESOURCE(s))
        return Tcl_NewLongObj((long) (LONG_PTR) s); /* Double cast to avoid warning */
    else
        return ObjFromUnicode(s);
}


TCL_RESULT Twapi_UpdateResource(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE h;
    WORD langid;
    void *resP;
    DWORD reslen;
    LPCWSTR restype;
    LPCWSTR resname;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(h),
                     GETVAR(restype, ObjToResourceIntOrString),
                     GETVAR(resname, ObjToResourceIntOrString),
                     GETWORD(langid),
                     ARGUSEDEFAULT,
                     GETBIN(resP, reslen),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }
    /* Note resP / reslen might be NULL/0 -> delete resource */

    if (UpdateResourceW(h, restype, resname, langid, resP, reslen))
        return TCL_OK;
    else
        return TwapiReturnSystemError(interp);
}

TCL_RESULT Twapi_FindResourceEx(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE h;
    WORD langid;
    LPCWSTR restype;
    LPCWSTR resname;
    HRSRC hrsrc;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(h),
                     GETVAR(restype, ObjToResourceIntOrString),
                     GETVAR(resname, ObjToResourceIntOrString),
                     GETWORD(langid),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    hrsrc = FindResourceExW(h, restype, resname, langid);
    if (hrsrc) {
        Tcl_SetObjResult(interp, ObjFromOpaque(hrsrc, "HRSRC"));
        return TCL_OK;
    } else
        return TwapiReturnSystemError(interp);
}

TCL_RESULT Twapi_LoadResource(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE hmodule;
    HRSRC hrsrc;
    HGLOBAL hglob;
    DWORD ressize;
    void *resP;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hmodule), GETHANDLET(hrsrc, HRSRC),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }


    ressize = SizeofResource(hmodule, hrsrc);
    if (ressize) {
        hglob = LoadResource(hmodule, hrsrc);
        if (hglob) {
            resP = LockResource(hglob);
            if (resP) {
                Tcl_SetObjResult(interp, Tcl_NewByteArrayObj(resP, ressize));
                return TCL_OK;
            }
        }
    }

    return TwapiReturnSystemError(interp);
}


static BOOL CALLBACK EnumResourceNamesProc(
    HMODULE hModule,
    LPCWSTR lpszType,
    LPWSTR lpszName,
    LONG_PTR lParam
)
{
    TwapiEnumCtx *ctxP = (TwapiEnumCtx *) lParam;

    Tcl_ListObjAppendElement(ctxP->interp, ctxP->objP,
                             ObjFromResourceIntOrString(lpszName));

    return TRUE;

}

TCL_RESULT Twapi_EnumResourceNames(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE hmodule;
    LPCWSTR restype;
    TwapiEnumCtx ctx;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hmodule),
                     GETVAR(restype, ObjToResourceIntOrString),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    ctx.interp = interp;
    ctx.objP = Tcl_NewListObj(0, NULL);
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceNamesW(hmodule, restype, EnumResourceNamesProc, (LONG_PTR) &ctx)) {
        Tcl_SetObjResult(interp, ctx.objP);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_OK;
    } else {
        TwapiReturnSystemError(interp);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_ERROR;
    }
}

static BOOL CALLBACK EnumResourceTypesProc(
    HMODULE hModule,
    LPCWSTR lpszType,
    LONG_PTR lParam
)
{
    TwapiEnumCtx *ctxP = (TwapiEnumCtx *) lParam;

    Tcl_ListObjAppendElement(ctxP->interp, ctxP->objP,
                             ObjFromResourceIntOrString(lpszType));
    return TRUE;
}

TCL_RESULT Twapi_EnumResourceTypes(
    Tcl_Interp *interp,
    HMODULE hmodule)
{
    TwapiEnumCtx ctx;

    ctx.interp = interp;
    ctx.objP = Tcl_NewListObj(0, NULL);
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceTypesW(hmodule, EnumResourceTypesProc, (LONG_PTR) &ctx)) {
        Tcl_SetObjResult(interp, ctx.objP);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_OK;
    } else {
        TwapiReturnSystemError(interp);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_ERROR;
    }
}

static BOOL CALLBACK EnumResourceLanguagesProc(
    HMODULE hModule,
    LPCWSTR lpszType,
    LPCWSTR lpszName,
    WORD langid,
    LONG_PTR lParam
)
{
    TwapiEnumCtx *ctxP = (TwapiEnumCtx *) lParam;

    Tcl_ListObjAppendElement(ctxP->interp, ctxP->objP, Tcl_NewIntObj(langid));
    return TRUE;
}

TCL_RESULT Twapi_EnumResourceLanguages(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE hmodule;
    LPCWSTR restype;
    LPCWSTR resname;
    TwapiEnumCtx ctx;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hmodule),
                     GETVAR(restype, ObjToResourceIntOrString),
                     GETVAR(resname, ObjToResourceIntOrString),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    ctx.interp = interp;
    ctx.objP = Tcl_NewListObj(0, NULL);
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceLanguagesW(hmodule, restype, resname,
                               EnumResourceLanguagesProc, (LONG_PTR) &ctx)) {
        Tcl_SetObjResult(interp, ctx.objP);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_OK;
    } else {
        TwapiReturnSystemError(interp);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_ERROR;
    }
}

/* Splits the 16-string block of a STRINGTABLE resource into a list of strings */
TCL_RESULT Twapi_SplitStringResource(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    WCHAR *wP;
    int len;
    Tcl_Obj *objP = Tcl_NewListObj(0, NULL);
    
    if (TwapiGetArgs(interp, objc, objv,
                     GETBIN(wP, len),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }
    
    while (len >= 2) {
        WORD slen = *wP++;
        if (slen > (sizeof(WCHAR)*len)) {
            /* Should not happen */
            break;
        }
        Tcl_ListObjAppendElement(interp, objP,
                                 ObjFromUnicodeN(wP, slen));
        wP += slen;
        len -= sizeof(WCHAR)*(1+slen);
    }

    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}

TCL_RESULT Twapi_LoadImage(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE h;
    LPCWSTR resname;
    DWORD image_type, cx, cy, flags;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(h),
                     GETVAR(resname, ObjToResourceIntOrString),
                     GETINT(image_type),
                     ARGUSEDEFAULT, GETINT(cx), GETINT(cy),
                     GETINT(flags),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    result.type = TRT_HANDLE;
    result.value.hval = LoadImageW(h, resname, image_type, cx, cy, flags);
    return TwapiSetResult(interp, &result);
}
