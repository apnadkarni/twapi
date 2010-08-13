/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

static TCL_RESULT ObjToResourceIntOrString(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *wsP)
{
    int i;

    /* Resource type and name can be integers or strings */
    if (Tcl_GetLongFromObj(NULL, objP, &i) == TCL_OK)
        *wsP = MAKEINTRESOURCEW(i);
    else
        *wsP = Tcl_GetUnicode(objP);

    return TCL_OK;
}

static Tcl_Obj *ObjFromResourceIntOrString(LPCWSTR s)
{
    if (IS_INTRESOURCE(s))
        return Tcl_NewLongObj((long) (LONG_PTR) s); /* Double cast to avoid warning */
    else
        return Tcl_NewUnicodeObj(s, -1);
}


TCL_RESULT Twapi_UpdateResource(
    TwapiInterpContext *ticP,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE h;
    WORD langid;
    void *resP;
    DWORD reslen;
    LPCWSTR restype;
    LPCWSTR resname;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETHANDLE(h),
                     GETVAR(restype, ObjToResourceIntOrString),
                     GETVAR(resname, ObjToResourceIntOrString),
                     GETWORD(langid), GETBIN(resP, reslen),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    if (UpdateResourceW(h, restype, resname, langid, resP, reslen))
        return TCL_OK;
    else
        return TwapiReturnSystemError(ticP->interp);
}

TCL_RESULT Twapi_FindResourceEx(
    TwapiInterpContext *ticP,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE h;
    WORD langid;
    LPCWSTR restype;
    LPCWSTR resname;
    HRSRC hrsrc;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETHANDLE(h),
                     GETVAR(restype, ObjToResourceIntOrString),
                     GETVAR(resname, ObjToResourceIntOrString),
                     GETWORD(langid),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    hrsrc = FindResourceExW(h, restype, resname, langid);
    if (hrsrc) {
        Tcl_SetObjResult(ticP->interp, ObjFromOpaque(hrsrc, "HRSRC"));
        return TCL_OK;
    } else
        return TwapiReturnSystemError(ticP->interp);
}

TCL_RESULT Twapi_LoadResource(
    TwapiInterpContext *ticP,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE hmodule;
    HRSRC hrsrc;
    HGLOBAL hglob;
    DWORD ressize;
    void *resP;

    if (TwapiGetArgs(ticP->interp, objc, objv,
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
                Tcl_SetObjResult(ticP->interp, Tcl_NewByteArrayObj(resP, ressize));
                return TCL_OK;
            }
        }
    }

    return TwapiReturnSystemError(ticP->interp);
}


static BOOL CALLBACK EnumResourceNamesProc(
    HMODULE hModule,
    LPCWSTR lpszType,
    LPWSTR lpszName,
    LONG_PTR lParam
)
{
    TwapiEnumCtx *ctxP = (TwapiEnumCtx *) lParam;
    Tcl_Obj *objs[3];

    objs[0] = ObjFromOpaque(hModule, "HMODULE");
    objs[1] = ObjFromResourceIntOrString(lpszType);
    objs[2] = ObjFromResourceIntOrString(lpszName);

    Tcl_ListObjAppendElement(ctxP->interp, ctxP->objP,
                             Tcl_NewListObj(3, objs));

    return TRUE;

}

TCL_RESULT Twapi_EnumResourceNames(
    TwapiInterpContext *ticP,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE hmodule;
    LPCWSTR restype;
    TwapiEnumCtx ctx;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETHANDLE(hmodule),
                     GETVAR(restype, ObjToResourceIntOrString),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    ctx.interp = ticP->interp;
    ctx.objP = Tcl_NewListObj(0, NULL);
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceNamesW(hmodule, restype, EnumResourceNamesProc, (LONG_PTR) &ctx)) {
        Tcl_SetObjResult(ticP->interp, ctx.objP);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_OK;
    } else {
        TwapiReturnSystemError(ticP->interp);
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
    Tcl_Obj *objs[2];

    objs[0] = ObjFromOpaque(hModule, "HMODULE");
    objs[1] = ObjFromResourceIntOrString(lpszType);

    Tcl_ListObjAppendElement(ctxP->interp, ctxP->objP,
                             Tcl_NewListObj(2, objs));

    return TRUE;
}

TCL_RESULT Twapi_EnumResourceTypes(TwapiInterpContext *ticP, HMODULE hmodule)
{
    TwapiEnumCtx ctx;

    ctx.interp = ticP->interp;
    ctx.objP = Tcl_NewListObj(0, NULL);
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceTypesW(hmodule, EnumResourceTypesProc, (LONG_PTR) &ctx)) {
        Tcl_SetObjResult(ticP->interp, ctx.objP);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_OK;
    } else {
        TwapiReturnSystemError(ticP->interp);
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
    Tcl_Obj *objs[4];

    objs[0] = ObjFromOpaque(hModule, "HMODULE");
    objs[1] = ObjFromResourceIntOrString(lpszType);
    objs[2] = ObjFromResourceIntOrString(lpszName);
    objs[3] = Tcl_NewIntObj(langid);

    Tcl_ListObjAppendElement(ctxP->interp, ctxP->objP,
                             Tcl_NewListObj(4, objs));

    return TRUE;

}

TCL_RESULT Twapi_EnumResourceLanguages(
    TwapiInterpContext *ticP,
    int objc,
    Tcl_Obj *CONST objv[])
{
    HANDLE hmodule;
    LPCWSTR restype;
    LPCWSTR resname;
    TwapiEnumCtx ctx;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETHANDLE(hmodule),
                     GETVAR(restype, ObjToResourceIntOrString),
                     GETVAR(resname, ObjToResourceIntOrString),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    ctx.interp = ticP->interp;
    ctx.objP = Tcl_NewListObj(0, NULL);
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceLanguagesW(hmodule, restype, resname,
                               EnumResourceLanguagesProc, (LONG_PTR) &ctx)) {
        Tcl_SetObjResult(ticP->interp, ctx.objP);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_OK;
    } else {
        TwapiReturnSystemError(ticP->interp);
        Tcl_DecrRefCount(ctx.objP);
        return TCL_ERROR;
    }
}
