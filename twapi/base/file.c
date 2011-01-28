/*
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"


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

    resultObj = Tcl_NewListObj(0, NULL);
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

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}

int Twapi_VerQueryValue_STRING(
    Tcl_Interp *interp,
    TWAPI_FILEVERINFO* verP,
    LPCSTR lang_and_cp,
    LPSTR name
    )
{
    Tcl_Obj *objP;
    WCHAR *valueP;
    UINT   len;

    objP = Tcl_ObjPrintf("\\StringFileInfo\\%s\\%s", lang_and_cp, name);
    if ((! VerQueryValueW(verP, Tcl_GetUnicode(objP), (LPVOID) &valueP, &len)) ||
        len == 0) {
        /* Return empty string, not error */
        Tcl_DecrRefCount(objP);
        return TCL_OK;
    }
    Tcl_DecrRefCount(objP);

    /* Note valueP does not have to be freed, points into verP */
    Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(valueP, -1));
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
        //Tcl_SetResult(interp, "VerQueryValue returned 0", TCL_STATIC);
        return TCL_OK;
    }

    resultObj = Tcl_NewListObj(0, NULL);
    for (dwP = bufP; ((char *)dwP) <= (len - sizeof(*dwP) + (char *)bufP) ; ++dwP) {
        WORD *wP = (WORD *) dwP;
        /* Print as langid,codepage */
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_ObjPrintf("%04x%04x", wP[0], wP[1])
            );
    }

    /* Note bufP does not have to be freed, points into verP */

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}

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



