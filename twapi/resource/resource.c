/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

/* Resource manipulation */
void Twapi_FreeFileVersionInfo(TWAPI_FILEVERINFO * verP);
int Twapi_LoadImage(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_UpdateResource(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_FindResourceEx(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_LoadResource(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_EnumResourceNames(Tcl_Interp *,int objc,Tcl_Obj *CONST objv[]);
int Twapi_EnumResourceLanguages(Tcl_Interp *,int objc,Tcl_Obj *CONST objv[]);
int Twapi_EnumResourceTypes(Tcl_Interp *, HMODULE hmodule);
TCL_RESULT Twapi_SplitStringResource(Tcl_Interp *interp, int objc,
                                     Tcl_Obj *CONST objv[]);
Tcl_Obj *ObjFromResourceIntOrString(LPCWSTR s);
TCL_RESULT ObjToResourceIntOrString(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *wsP);


/* Note: Param 'interp' is there to match prototype expected by TwapiGetArgs */
TCL_RESULT ObjToResourceIntOrString(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *wsP)
{
    int i;

    /*
     * Resource type and name can be integers or strings. If integer,
     * they must pass the IS_INTRESOURCE test as not all integers are valid.
     */
    if (ObjToLong(NULL, objP, &i) == TCL_OK &&
        IS_INTRESOURCE(i))
        *wsP = MAKEINTRESOURCEW(i);
    else
        *wsP = ObjToUnicode(objP);

    return TCL_OK;
}

Tcl_Obj *ObjFromResourceIntOrString(LPCWSTR s)
{
    if (IS_INTRESOURCE(s))
        return Tcl_NewLongObj((long) (LONG_PTR) s); /* Double cast to avoid warning */
    else
        return ObjFromUnicode(s);
}


TWAPI_FILEVERINFO* Twapi_GetFileVersionInfo(LPWSTR path)
{
    DWORD  sz;
    DWORD  unused;
    TWAPI_FILEVERINFO* fverinfoP;

    sz = GetFileVersionInfoSizeW(path, &unused);
    if (sz == 0)
        return NULL;

    fverinfoP = TwapiAlloc(sz);

    if (! GetFileVersionInfoW(path, 0, sz, fverinfoP)) {
        DWORD winerr = GetLastError();
        TwapiFree(fverinfoP);
        SetLastError(winerr);
        return NULL;
    }

    return fverinfoP;
}

void Twapi_FreeFileVersionInfo(TWAPI_FILEVERINFO* fverinfoP)
{
    if (fverinfoP)
        TwapiFree(fverinfoP);
}

int Twapi_VerQueryValue_FIXEDFILEINFO(
    Tcl_Interp *interp,
    TWAPI_FILEVERINFO* verP
    )
{
    Tcl_Obj           *resultObj;
    VS_FIXEDFILEINFO *ffiP;
    UINT              ffi_sz;

    if (! VerQueryValue(verP, "\\", (LPVOID) &ffiP, &ffi_sz)) {
        /* Return empty list, not error */
        return TCL_OK;
    }

    resultObj = ObjEmptyList();
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwSignature);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwStrucVersion);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileVersionMS);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileVersionLS);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwProductVersionMS);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwProductVersionLS);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileFlagsMask);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileFlags);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileOS);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileType);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileSubtype);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileDateMS);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, ffiP, dwFileDateLS);

    TwapiSetObjResult(interp, resultObj);
    return TCL_OK;
}

int Twapi_VerQueryValue_STRING(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    TWAPI_FILEVERINFO* verP;
    LPCSTR lang_and_cp;
    LPSTR name;
    Tcl_Obj *objP;
    WCHAR *valueP;
    UINT   len;

    if (TwapiGetArgs(interp, objc, objv, GETPTR(verP, TWAPI_FILEVERINFO),
                     GETASTR(lang_and_cp), GETASTR(name), ARGEND) != TCL_OK)
        return TCL_ERROR;

    objP = Tcl_ObjPrintf("\\StringFileInfo\\%s\\%s", lang_and_cp, name);
    if ((! VerQueryValueW(verP, ObjToUnicode(objP), (LPVOID) &valueP, &len)) ||
        len == 0) {
        /* Return empty string, not error */
        Tcl_DecrRefCount(objP);
        return TCL_OK;
    }
    Tcl_DecrRefCount(objP);

    /* Note valueP does not have to be freed, points into verP */
    TwapiSetObjResult(interp, ObjFromUnicode(valueP));
    return TCL_OK;
}

int Twapi_VerQueryValue_TRANSLATIONS(
    Tcl_Interp *interp,
    TWAPI_FILEVERINFO* verP
    )
{
    Tcl_Obj  *resultObj;
    DWORD    *bufP;
    DWORD    *dwP;
    UINT      len;

    if (! VerQueryValue(verP, "\\VarFileInfo\\Translation", (LPVOID) &bufP, &len)) {
        /* Return empty list, not error */
        return TCL_OK;
    }

    resultObj = ObjEmptyList();
    for (dwP = bufP; ((char *)dwP) <= (len - sizeof(*dwP) + (char *)bufP) ; ++dwP) {
        WORD *wP = (WORD *) dwP;
        /* Print as langid,codepage */
        ObjAppendElement(interp, resultObj,
                                 Tcl_ObjPrintf("%04x%04x", wP[0], wP[1])
            );
    }

    /* Note bufP does not have to be freed, points into verP */

    TwapiSetObjResult(interp, resultObj);
    return TCL_OK;
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
        TwapiSetObjResult(interp, ObjFromOpaque(hrsrc, "HRSRC"));
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
                TwapiSetObjResult(interp, ObjFromByteArray(resP, ressize));
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

    ObjAppendElement(ctxP->interp, ctxP->objP,
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
    ctx.objP = ObjEmptyList();
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceNamesW(hmodule, restype, EnumResourceNamesProc, (LONG_PTR) &ctx)) {
        TwapiSetObjResult(interp, ctx.objP);
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

    ObjAppendElement(ctxP->interp, ctxP->objP,
                             ObjFromResourceIntOrString(lpszType));
    return TRUE;
}

TCL_RESULT Twapi_EnumResourceTypes(
    Tcl_Interp *interp,
    HMODULE hmodule)
{
    TwapiEnumCtx ctx;

    ctx.interp = interp;
    ctx.objP = ObjEmptyList();
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceTypesW(hmodule, EnumResourceTypesProc, (LONG_PTR) &ctx)) {
        TwapiSetObjResult(interp, ctx.objP);
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

    ObjAppendElement(ctxP->interp, ctxP->objP, ObjFromDWORD(langid));
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
    ctx.objP = ObjEmptyList();
    Tcl_IncrRefCount(ctx.objP);  /* Protect in callback, just in case */
    if (EnumResourceLanguagesW(hmodule, restype, resname,
                               EnumResourceLanguagesProc, (LONG_PTR) &ctx)) {
        TwapiSetObjResult(interp, ctx.objP);
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
    Tcl_Obj *objP = ObjEmptyList();
    
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
        ObjAppendElement(interp, objP,
                                 ObjFromUnicodeN(wP, slen));
        wP += slen;
        len -= sizeof(WCHAR)*(1+slen);
    }

    TwapiSetObjResult(interp, objP);
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

static int Twapi_ResourceCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR s, s2, s3;
    DWORD dw;
    HANDLE h;
    TwapiResult result;
    void *pv;
    int func = PtrToInt(clientdata);

    /* Every command has at least one argument */
    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:              /* DestroyIcon */
        if (ObjToOpaque(interp, objv[0], &pv, "HICON") != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = DestroyIcon(pv);
        break;
    case 2:              /* DestroyCursor */
        if (ObjToOpaque(interp, objv[0], &pv, "HCURSOR") != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = DestroyCursor(pv);
        break;
    case 3: // VerQueryValue_STRING
        return Twapi_VerQueryValue_STRING(interp, objc, objv);
    case 4: // FALLTHRU
    case 5: // FALLTHRU
    case 6:
        if (ObjToOpaque(interp, objv[0], &pv, "TWAPI_FILEVERINFO") != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 4:
            return Twapi_VerQueryValue_FIXEDFILEINFO(interp, pv);
        case 5:
            return Twapi_VerQueryValue_TRANSLATIONS(interp, pv);
        case 6:
            Twapi_FreeFileVersionInfo(pv);
            return TCL_OK;
        }
        break;
    case 7:
        return Twapi_UpdateResource(interp, objc, objv);
    case 8:
        return Twapi_FindResourceEx(interp, objc, objv);
    case 9:
        return Twapi_LoadResource(interp, objc, objv);
    case 10:
        return Twapi_EnumResourceNames(interp, objc, objv);
    case 11:
        return Twapi_EnumResourceLanguages(interp, objc, objv);
    case 12:
        return Twapi_SplitStringResource(interp, objc, objv);
    case 13: // LoadImage
        return Twapi_LoadImage(interp, objc, objv);
    case 14:
        result.type = TRT_NONNULL;
        result.value.nonnull.name = "TWAPI_FILEVERINFO";
        result.value.nonnull.p = Twapi_GetFileVersionInfo(ObjToUnicode(objv[0]));
        break;
    case 15: // AddFontResourceEx
    case 16: // BeginUpdateResource
        if (TwapiGetArgs(interp, objc, objv, GETWSTR(s), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (func == 15) {
            result.type = TRT_LONG;
            result.value.ival = AddFontResourceExW(s, dw, NULL);
        } else {
            result.type = TRT_HANDLE;
            result.value.hval = BeginUpdateResourceW(s, dw);
        }
        break;
    case 17:
        if (ObjToHANDLE(interp, objv[0], &h) != TCL_OK)
            return TCL_ERROR;
        return Twapi_EnumResourceTypes(interp, h);
    case 18: // EndUpdateResource
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h), GETBOOL(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EndUpdateResourceW(h, dw);
        break;
    case 19:
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw), GETWSTR(s),
                         GETWSTR(s2), GETWSTR(s3), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CreateScalableFontResourceW(dw, s, s2, s3);
        break;
    case 20:
        if (TwapiGetArgs(interp, objc, objv, GETWSTR(s), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        result.value.bval = RemoveFontResourceExW(s, dw, NULL);
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiResourceInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s ResDispatch[] = {
        DEFINE_FNCODE_CMD(DestroyIcon, 1),
        DEFINE_FNCODE_CMD(DestroyCursor, 2),
        DEFINE_FNCODE_CMD(Twapi_VerQueryValue_STRING, 3),
        DEFINE_FNCODE_CMD(Twapi_VerQueryValue_FIXEDFILEINFO, 4),
        DEFINE_FNCODE_CMD(Twapi_VerQueryValue_TRANSLATIONS, 5),
        DEFINE_FNCODE_CMD(Twapi_FreeFileVersionInfo, 6),
        DEFINE_FNCODE_CMD(UpdateResource, 7),
        DEFINE_FNCODE_CMD(FindResourceEx, 8),
        DEFINE_FNCODE_CMD(Twapi_LoadResource, 9),
        DEFINE_FNCODE_CMD(Twapi_EnumResourceNames, 10),
        DEFINE_FNCODE_CMD(Twapi_EnumResourceLanguages, 11),
        DEFINE_FNCODE_CMD(Twapi_SplitStringResource, 12),
        DEFINE_FNCODE_CMD(LoadImage, 13),
        DEFINE_FNCODE_CMD(Twapi_GetFileVersionInfo, 14),
        DEFINE_FNCODE_CMD(AddFontResourceEx, 15), // TBD - Tcl wrapper
        DEFINE_FNCODE_CMD(BeginUpdateResource, 16),
        DEFINE_FNCODE_CMD(Twapi_EnumResourceTypes, 17),
        DEFINE_FNCODE_CMD(EndUpdateResource, 18),
        DEFINE_FNCODE_CMD(CreateScalableFontResource, 19), // TBD Tcl
        DEFINE_FNCODE_CMD(RemoveFontResourceEx, 20), // TBD Tcl 
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(ResDispatch), ResDispatch, Twapi_ResourceCallObjCmd);
    return TCL_OK;
}


#ifndef TWAPI_SINGLE_MODULE
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_SINGLE_MODULE
__declspec(dllexport) 
#endif
int Twapi_resource_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiResourceInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

