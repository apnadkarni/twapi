/*
 * Copyright (c) 2020, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include <shlwapi.h>

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

#ifndef MODULENAME
#define MODULENAME "twapi_registry"
#endif

#ifndef TWAPI_SINGLE_MODULE
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Window key length limits */
#define TWAPI_KEY_NAME_NCHARS 255
#define TWAPI_VALUE_NAME_NCHARS 16383

/*
 * Define API not in XP
 */
MAKE_DYNLOAD_FUNC(RegCopyTreeW, advapi32, FARPROC)
MAKE_DYNLOAD_FUNC2(RegDeleteKeyValueW, advapi32)
MAKE_DYNLOAD_FUNC2(RegDeleteKeyValueW, kernel32)
MAKE_DYNLOAD_FUNC(RegDeleteKeyExW, advapi32, FARPROC)
MAKE_DYNLOAD_FUNC(RegDeleteTreeW, advapi32, FARPROC)
MAKE_DYNLOAD_FUNC(RegDisableReflectionKey, advapi32, FARPROC)
MAKE_DYNLOAD_FUNC(RegEnableReflectionKey, advapi32, FARPROC)
MAKE_DYNLOAD_FUNC(RegGetValueW, advapi32, FARPROC)
MAKE_DYNLOAD_FUNC(RegSetKeyValueW, advapi32, FARPROC)
MAKE_DYNLOAD_FUNC(SHCopyKeyW, shlwapi, FARPROC)
MAKE_DYNLOAD_FUNC(SHDeleteKeyW, shlwapi, FARPROC)
MAKE_DYNLOAD_FUNC(SHDeleteValueW, shlwapi, FARPROC)
MAKE_DYNLOAD_FUNC(SHRegGetValueW, shlwapi, FARPROC)

static int TwapiRegEnumKeyEx(Tcl_Interp *interp, HKEY hkey, DWORD flags)
{
    Tcl_Obj *resultObj = NULL;
    FILETIME file_time;
    DWORD    nch_subkey, dwIndex;
    WCHAR    subkey[TWAPI_KEY_NAME_NCHARS];
    LONG     status;

    resultObj = ObjNewList(0, NULL);
    dwIndex = 0;
    while (1) {
        Tcl_Obj *objs[2];
        nch_subkey = ARRAYSIZE(subkey);
        status     = RegEnumKeyExW(hkey, dwIndex, subkey, &nch_subkey, NULL,
                                   NULL, NULL, &file_time);
        /* Should errors be ignored? TBD */
        if (status != ERROR_SUCCESS)
            break;
        objs[0] = ObjFromTclUniCharN(subkey, nch_subkey);
        if (flags & 1) {
            objs[1] = ObjFromFILETIME(&file_time);
            ObjAppendElement(interp, resultObj, ObjNewList(2, objs));
        } else
            ObjAppendElement(interp, resultObj, objs[0]);

        ++dwIndex;
    }
    if (status == ERROR_NO_MORE_ITEMS) {
        ObjSetResult(interp, resultObj);
        return TCL_OK;
    } else {
        if (resultObj)
            ObjDecrRefs(resultObj);
        return Twapi_AppendSystemError(interp, status);
    }
}

static int TwapiRegEnumValue(Tcl_Interp *interp, HKEY hkey, DWORD flags)
{
    Tcl_Obj *resultObj = NULL;
    FILETIME file_time;
    LPWSTR   value_name;
    DWORD    nch_value_name;
    DWORD    dwIndex;
    LONG  status;
    SWSMark  mark;

    mark = SWSPushMark();
    value_name          = SWSAlloc(sizeof(WCHAR) * TWAPI_VALUE_NAME_NCHARS, NULL);
    resultObj = ObjNewList(0, NULL);
    if (flags & 1) {
        /* Caller wants data as well. */
        DWORD nb_value_data = 256;
        LPBYTE value_data;
        DWORD  value_type;
        int    max_loop; /* Safety measure if buf size keeps changing */
        value_data = SWSAlloc(nb_value_data, &nb_value_data);
        dwIndex    = 0;
        max_loop   = 10; /* Retries for a particular key. Else error */
        while (--max_loop >= 0) {
            DWORD original_nb_value_data;
            original_nb_value_data = nb_value_data;
            nch_value_name         = TWAPI_VALUE_NAME_NCHARS;
            status         = RegEnumValueW(hkey,
                                   dwIndex,
                                   value_name,
                                   &nch_value_name,
                                   NULL,
                                   &value_type,
                                   value_data,
                                   &nb_value_data);
            if (status == ERROR_SUCCESS) {
                Tcl_Obj *objs[2];
                objs[1] = (flags & 2 ? ObjFromRegValueCooked : ObjFromRegValue)(
                    interp, value_type, value_data, nb_value_data);
                /* Bad values are skipped - TBD */
                if (objs[1] != NULL) {
                    /* Workaround for HKEY_PERFORMANCE_DATA bug - return
                     * length is off by 1.
                     */
                    if (hkey == HKEY_PERFORMANCE_DATA) {
                        value_name[TWAPI_VALUE_NAME_NCHARS - 1] = 0; /* Ensure null termination */
                        objs[0] = ObjFromTclUniChar(value_name);
                    } else {
                        objs[0] = ObjFromTclUniCharN(value_name, nch_value_name);
                    }
                    ObjAppendElement(interp, resultObj, objs[0]);
                    ObjAppendElement(interp, resultObj, objs[1]);
                }
                ++dwIndex;
                max_loop = 10; /* Reset safety check */
            }
            else if (status == ERROR_MORE_DATA) {
                /*
                 * Workaround for HKEY_PERFORMANCE_DATA bug - does not
                 * update nb_value_data. In that case, double the current
                 * size
                 */
                if (nb_value_data == original_nb_value_data)
                    nb_value_data *= 2; // TBD - check for overflow(!)
                value_data = SWSAlloc(nb_value_data, NULL);
                /* Do not increment dwIndex and retry for same */
            } else {
                /* ERROR_NO_MORE_ITEMS or some other error */
                break;
            }
        }
    } else {
        /* Only value names asked for */
        dwIndex = 0;
        while (1) {
            Tcl_Obj *objP;
            nch_value_name = TWAPI_VALUE_NAME_NCHARS;
            status         = RegEnumValueW(hkey,
                                   dwIndex,
                                   value_name,
                                   &nch_value_name,
                                   NULL,
                                   NULL,
                                   NULL,
                                   NULL);
            /* Since we are passing NULL as data buffer, the call will
               return ERROR_MORE_DATA even on success. Note that since
               we are passing max size name buffer, this error will
               not be for the value_name buffer */
            if (status != ERROR_SUCCESS && status != ERROR_MORE_DATA)
                break;
            /* Workaround for HKEY_PERFORMANCE_DATA bug - length is off by 1. */
            if (hkey == HKEY_PERFORMANCE_DATA) {
                value_name[TWAPI_VALUE_NAME_NCHARS - 1] = 0; /* Ensure null termination */
                objP = ObjFromTclUniChar(value_name);
            } else {
                objP = ObjFromTclUniCharN(value_name, nch_value_name);
            }
            ObjAppendElement(interp, resultObj, objP);
            ++dwIndex;
        }
    }
    SWSPopMark(mark);
    if (status == ERROR_NO_MORE_ITEMS) {
        ObjSetResult(interp, resultObj);
        return TCL_OK;
    } else {
        if (resultObj)
            ObjDecrRefs(resultObj);
        return Twapi_AppendSystemError(interp, status);
    }
}

static int
TwapiRegGetValue(Tcl_Interp *interp,
                 HKEY        hkey,
                 LPCWSTR     subkey,
                 LPCWSTR     value_name,
                 DWORD       flags,
                 BOOL        cooked
                 )
{
    Tcl_Obj *resultObj = NULL;
    LONG     status; /* Win32 code */
    FILETIME file_time;
    SWSMark  mark;
    LPBYTE   value_data;
    DWORD    value_type, nb_value_data;
    int      max_loop; /* Safety measure if buf size keeps changing */
    FARPROC func = Twapi_GetProc_RegGetValueW();

    mark = SWSPushMark();
    resultObj = ObjNewList(0, NULL);

    nb_value_data = 256;
    value_data = SWSAlloc(nb_value_data, &nb_value_data);
    max_loop   = 10; /* Retries for a particular key. Else error */
    while (--max_loop >= 0) {
        DWORD original_nb_value_data = nb_value_data;
        if (func) {
            flags &= 0x00030000; /* RRF_SUBKEY_WOW64{32,64}KEY */
            flags |= 0x1000ffff; /* RRF_NOEXPAND, RRF_RT_ANY */
            status = func(hkey,
                          subkey,
                          value_name,
                          flags,
                          &value_type,
                          value_data,
                          &nb_value_data);
        } else {
            /* Note if WOW64 bits were set in flags, func would not be NULL as 
             * RegGetValueW would be present and we would not come here */
            func = Twapi_GetProc_SHRegGetValueW();
            if (func)
                status = func(hkey,
                                        subkey,
                                        value_name,
                                        0x1000ffff, //SRRF_RT_ANY | SRRF_NOEXPAND
                                        &value_type,
                                        value_data,
                                        &nb_value_data);
            else
                status = ERROR_PROC_NOT_FOUND;
        }
        if (status == ERROR_SUCCESS) {
            if (cooked) {
                resultObj
                    = ObjFromRegValueCooked(interp, value_type, value_data, nb_value_data);
            } else {
                resultObj = ObjFromRegValue(
                    interp, value_type, value_data, nb_value_data);
            }
            break;
        } else if (status == ERROR_MORE_DATA) {
            /*
             * Workaround for HKEY_PERFORMANCE_DATA bug - does not
             * update nb_value_data. In that case, double the current
             * size
             */
            if (nb_value_data == original_nb_value_data)
                nb_value_data *= 2; // TBD - check for overflow(!)
            value_data = SWSAlloc(nb_value_data, NULL);
            /* Do not increment dwIndex and retry for same */
        } else {
            break;
        }
    }
    SWSPopMark(mark);
    if (status == ERROR_SUCCESS) {
        /* The Win32 call may have succeeded but some Tcl call might fail */
        if (resultObj == NULL)
            return TCL_ERROR; /* interp should already have error message */
        ObjSetResult(interp, resultObj);
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, status);
    }
}

static int
TwapiRegQueryValueEx(Tcl_Interp *interp,
                 HKEY        hkey,
                 LPCWSTR     value_name,
                 BOOL        cooked)
{
    Tcl_Obj *resultObj = NULL;
    LONG     status;
    SWSMark  mark;
    LPBYTE   value_data;
    DWORD    value_type, nb_value_data;
    int      max_loop; /* Safety measure if buf size keeps changing */

    mark = SWSPushMark();
    resultObj = ObjNewList(0, NULL);

    nb_value_data = 256;
    value_data = SWSAlloc(nb_value_data, &nb_value_data);
    max_loop   = 10; /* Retries for a particular key. Else error */
    while (--max_loop >= 0) {
        DWORD original_nb_value_data = nb_value_data;
        status         = RegQueryValueExW(hkey,
                              value_name,
                              NULL,
                              &value_type,
                              value_data,
                              &nb_value_data);
        if (status == ERROR_SUCCESS) {
            if (cooked) {
                resultObj
                    = ObjFromRegValueCooked(interp, value_type, value_data, nb_value_data);
            } else {
                resultObj = ObjFromRegValue(
                    interp, value_type, value_data, nb_value_data);
            }
            break;
        } else if (status == ERROR_MORE_DATA) {
            /*
             * Workaround for HKEY_PERFORMANCE_DATA bug - does not
             * update nb_value_data. In that case, double the current
             * size
             */
            if (nb_value_data == original_nb_value_data)
                nb_value_data *= 2; // TBD - check for overflow(!)
            value_data = SWSAlloc(nb_value_data, NULL);
            /* Do not increment dwIndex and retry for same */
        } else {
            /* ERROR_NO_MORE_ITEMS or some other error */
            break;
        }
    }
    SWSPopMark(mark);
    if (status == ERROR_SUCCESS) {
        /* The Win32 call may have succeeded but some Tcl call might fail */
        if (resultObj == NULL)
            return TCL_ERROR;
        ObjSetResult(interp, resultObj);
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, status);
    }
}

static int Twapi_RegCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HKEY                 hkey, hkey2;
    HANDLE               h;
    DWORD                dw, dw2, dw3, dw4;
    SECURITY_ATTRIBUTES *secattrP;
    SECURITY_DESCRIPTOR *secdP;
    SWSMark              mark = NULL;
    Tcl_Obj *            subkeyObj;
    Tcl_Obj *            nameObj;
    Tcl_Obj *            objP;
    Tcl_Obj *            obj2P;
    Tcl_Obj *            objs[2];
    TwapiResult          result;
    int                  func_code = PtrToInt(clientdata);

    --objc;
    ++objv;

    /* Assume error on system call */
    result.type = TRT_EXCEPTION_ON_ERROR;
    switch (func_code) {
    case 1: // RegOpenKeyEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj),
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival
            = RegOpenKeyExW(hkey, ObjToWinChars(subkeyObj), dw, dw2, &hkey2);
        if (result.value.ival == ERROR_SUCCESS) {
            result.type = TRT_HKEY;
            result.value.hval = hkey2;
        }
        break;

    case 2: // RegCreateKeyEx
        secattrP = NULL;
        mark = SWSPushMark();
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey),
                         GETOBJ(subkeyObj), GETINT(dw), ARGSKIP, GETINT(dw2),
                         GETINT(dw3),
                         GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTESSWS),
                         ARGEND) != TCL_OK) {
            SWSPopMark(mark);
            return TCL_ERROR;
        }
        result.value.ival = RegCreateKeyExW(hkey,
                                            ObjToWinChars(subkeyObj),
                                            dw,
                                            NULL,
                                            dw2,
                                            dw3,
                                            secattrP,
                                            &hkey2,
                                            &dw4);
        if (result.value.ival == ERROR_SUCCESS) {
            objs[0] = ObjFromOpaque(hkey2, "HKEY");
            objs[1] = ObjFromDWORD(dw4);
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            result.type = TRT_OBJV;
        }
        SWSPopMark(mark);
        break;

    case 3: // RegDeleteKeyValue
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey),
                         GETOBJ(subkeyObj), GETOBJ(objP),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            FARPROC func = Twapi_GetProc_RegDeleteKeyValueW_advapi32();
            if (func == NULL)
                func = Twapi_GetProc_RegDeleteKeyValueW_kernel32(); {
                if (func == NULL)
                    func = Twapi_GetProc_SHDeleteValueW();
            }
            if (func == NULL)
                result.value.ival = ERROR_PROC_NOT_FOUND;
            else {
                result.value.ival
                    = func(hkey, ObjToWinChars(subkeyObj), ObjToWinChars(objP));
                if (result.value.ival == ERROR_SUCCESS)
                    result.type = TRT_EMPTY;
            }
        }
        break;

    case 4: // RegDeleteKeyEx
        if (TwapiGetArgs(interp,
                         objc,
                         objv,
                         GETHKEY(hkey),
                         GETOBJ(subkeyObj),
                         ARGUSEDEFAULT,
                         GETINT(dw),
                         ARGEND)
            != TCL_OK)
            return TCL_ERROR;
        else {
            FARPROC func = Twapi_GetProc_RegDeleteKeyExW();
            if (func) {
                result.value.ival = func(hkey, ObjToWinChars(subkeyObj), dw, 0);
            }
            else {
                /* If the Ex call is not supported, the samDesired param
                * does not matter. Use legacy api
                */
                result.value.ival = RegDeleteKey(hkey, ObjToWinChars(subkeyObj));
            }
        }
        if (result.value.ival == ERROR_SUCCESS || result.value.ival == ERROR_FILE_NOT_FOUND)
            result.type = TRT_EMPTY;
        break;

    case 5: // RegDeleteValue
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(objP),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegDeleteValueW(hkey, ObjToWinChars(objP));
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 6: // RegDeleteTree
    /*
     * Ideally we want to use RegDeleteTree. That does not exist
     * on older systems so we use SHDeleteKey. That is not included
     * on VC++ 6 SP5 so we have to dynamically load everything.
     */
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            FARPROC func = Twapi_GetProc_RegDeleteTreeW();
            if (func == NULL)
                func = Twapi_GetProc_SHDeleteKeyW();
            if (func) {
                result.value.ival = func(hkey, ObjToWinChars(subkeyObj));
                if (result.value.ival == ERROR_SUCCESS)
                    result.type = TRT_EMPTY;
            }
            else {
                result.value.ival = ERROR_PROC_NOT_FOUND;
            }
        }
        break;

    case 7: // RegEnumKeyEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), ARGUSEDEFAULT, GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_TCL_RESULT;
        result.value.ival = TwapiRegEnumKeyEx(interp, hkey, dw);
        break;

    case 8: // RegEnumValue
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_TCL_RESULT;
        result.value.ival = TwapiRegEnumValue(interp, hkey, dw);
        break;

    case 9: // RegOpenCurrentUser
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegOpenCurrentUser(dw, &hkey);
        if (result.value.ival == ERROR_SUCCESS) {
            result.type = TRT_HKEY;
            result.value.hval = hkey;
        }
        break;

    case 10: // RegDisablePredefinedCache
        CHECK_NARGS(interp, objc, 0);
        result.value.ival = RegDisablePredefinedCache();
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 11: // RegGetKeySecurity
        if (TwapiGetArgs(
                interp, objc, objv, GETHKEY(hkey), GETINT(dw), ARGEND)
            != TCL_OK)
            return TCL_ERROR;
        mark              = SWSPushMark();
        secdP             = SWSAlloc(256, &dw2);
        result.value.ival = RegGetKeySecurity(hkey, dw, secdP, &dw2);
        if (result.value.ival == ERROR_INSUFFICIENT_BUFFER) {
            secdP = SWSAlloc(dw2, NULL);
            result.value.ival = RegGetKeySecurity(hkey, dw, secdP, &dw2);
        }
        if (result.value.ival == ERROR_SUCCESS) {
            result.value.obj = ObjFromSECURITY_DESCRIPTOR(interp, secdP);
            if (result.value.obj)
                result.type = TRT_OBJ;
            else {
                result.value.ival = TCL_ERROR;
                result.type       = TRT_TCL_RESULT;
            }
        }
        SWSPopMark(mark);
        break;

    case 12: // RegQueryValueEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey),
                         GETOBJ(objP),
                         GETBOOL(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_TCL_RESULT;
        result.value.ival
            = TwapiRegQueryValueEx(interp, hkey, ObjToWinChars(objP), dw);
        break;

    case 13: // RegCopyTree
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj), 
                         GETHKEY(hkey2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            FARPROC func = Twapi_GetProc_RegCopyTreeW();
            if (func) {
                Tcl_UniChar *subkey = ObjToWinCharsN(subkeyObj, &dw);
                if (dw == 0)
                    subkey = NULL;
                result.value.ival = func(hkey, subkey, hkey2);
                if (result.value.ival == ERROR_SUCCESS)
                    result.type = TRT_EMPTY;
            } else {
                result.value.ival = ERROR_PROC_NOT_FOUND;
            }
        }
        break;

    case 14: // RegOpenUserClassesRoot
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLE(h), GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegOpenUserClassesRoot(h, dw, dw2, &hkey);
        if (result.value.ival == ERROR_SUCCESS) {
            result.value.hkey = hkey;
            result.type       = TRT_HKEY;
        }
        break;

    case 15: // RegOverridePredefKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETHKEY(hkey2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegOverridePredefKey(hkey, hkey2);
        if (result.value.ival)
            result.type = TRT_EMPTY;
        break;

    case 16:  // RegSaveKeyEx
        secattrP = NULL;
        mark = SWSPushMark();
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey),
                         GETOBJ(objP),
                         GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTESSWS),
                         GETINT(dw),
                         ARGEND) != TCL_OK) {
            SWSPopMark(mark);
            return TCL_ERROR;
        }
        result.value.ival
            = RegSaveKeyExW(hkey, ObjToWinChars(objP), secattrP, dw);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        SWSPopMark(mark);
        break;

    case 17: // RegLoadKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj),
                         GETOBJ(objP), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival
            = RegLoadKeyW(hkey, ObjToWinChars(subkeyObj), ObjToWinChars(objP));
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 18: // RegUnLoadKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegUnLoadKeyW(hkey, ObjToWinChars(subkeyObj));
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 19: // RegReplaceKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj),
                         GETOBJ(objP), GETOBJ(obj2P),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegReplaceKeyW(hkey,
                                           ObjToWinChars(subkeyObj),
                                           ObjToWinChars(objP),
                                           ObjToWinChars(obj2P));
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 20: // RegRestoreKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey),
                         GETOBJ(objP), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegRestoreKeyW(hkey, ObjToWinChars(objP), dw);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 21: // RegSetValueEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(nameObj),
                         GETOBJ(objP), GETOBJ(obj2P),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            TwapiRegValue regval;
            mark              = SWSPushMark();
            result.value.ival = ObjToRegValueSWS(interp, objP, obj2P, &regval);
            if (result.value.ival != TCL_OK) {
                result.type = TRT_TCL_RESULT;
            } else {
                result.value.ival = RegSetValueExW(hkey,
                                                   ObjToWinChars(nameObj),
                                                   0,
                                                   regval.type,
                                                   regval.bytes,
                                                   regval.size);
                if (result.value.ival == TCL_OK)
                    result.type = TRT_EMPTY;
            }
            SWSPopMark(mark);
        }
        break;

    case 22:
        if (TwapiGetArgs(interp,
                         objc,
                         objv,
                         GETHKEY(hkey),
                         GETINT(dw),
                         GETINT(dw2),
                         ARGUSEDEFAULT,
                         GETHANDLE(h),
                         GETINT(dw3),
                         ARGEND)
            != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegNotifyChangeKeyValue(hkey, dw, dw2, h, dw3);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 23:
        secdP = NULL;
        mark = SWSPushMark();
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey),
                         GETINT(dw),
                         GETVAR(secdP, ObjToPSECURITY_DESCRIPTORSWS),
                         ARGEND) != TCL_OK) {
            SWSPopMark(mark);
            return TCL_ERROR;
        }
        result.value.ival =
            RegSetKeySecurity(hkey, dw, secdP);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        SWSPopMark(mark);
        break;

    case 24: // RegConnectRegistry
        if (TwapiGetArgs(interp, objc, objv, 
                         GETOBJ(objP), GETHKEY(hkey), ARGEND) != TCL_OK) {
            return TCL_ERROR;
        }
        result.value.ival = RegConnectRegistryW(ObjToWinChars(objP), hkey, &hkey2);
        if (result.value.ival == ERROR_SUCCESS) {
            result.type = TRT_HKEY;
            result.value.hval = hkey2;
        }
        break;

    case 25: // RegGetValue
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj), GETOBJ(nameObj),
                         GETINT(dw), GETBOOL(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_TCL_RESULT;
        result.value.ival
            = TwapiRegGetValue(interp, hkey, ObjToWinChars(subkeyObj), ObjToWinChars(nameObj), dw, dw2);
        break;

    case 26: // RegSetKeyValue
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj), GETOBJ(nameObj),
                         GETOBJ(objP), GETOBJ(obj2P),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            FARPROC func = Twapi_GetProc_RegSetKeyValueW();
            if (func == NULL) {
                result.value.ival = ERROR_PROC_NOT_FOUND;
            } else {
                TwapiRegValue regval;
                mark              = SWSPushMark();
                result.value.ival = ObjToRegValueSWS(interp, objP, obj2P, &regval);
                if (result.value.ival != TCL_OK) {
                    result.type = TRT_TCL_RESULT;
                } else {
                    result.value.ival = func(hkey,
                                             ObjToWinChars(subkeyObj),
                                             ObjToWinChars(nameObj),
                                             regval.type,
                                             regval.bytes,
                                             regval.size);
                    if (result.value.ival == TCL_OK)
                        result.type = TRT_EMPTY;
                }
                SWSPopMark(mark);
            }
        }
        break;

    case 27: // SHCopyKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), GETOBJ(subkeyObj), 
                         GETHKEY(hkey2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            FARPROC func = Twapi_GetProc_SHCopyKeyW();
            if (func) {
                Tcl_UniChar *subkey = ObjToWinCharsN(subkeyObj, &dw);
                if (dw == 0)
                    subkey = NULL;
                result.value.ival = func(hkey, subkey, hkey2, 0);
                if (result.value.ival == ERROR_SUCCESS)
                    result.type = TRT_EMPTY;
            } else
                result.value.ival = ERROR_PROC_NOT_FOUND;
        }
        break;

    case 28: // RegFlushKey
    case 29: // RegCloseKey
    case 30: // RegDisableReflectionKey
    case 31: // RegEnableReflectionKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETHKEY(hkey), ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            FARPROC func;
            switch (func_code) {
            case 28:
                func = (FARPROC) RegFlushKey;
                break;
            case 29:
                func = (FARPROC) RegCloseKey;
                break;
            case 30:
                func = Twapi_GetProc_RegDisableReflectionKey();
                break;
            case 31:
                func = Twapi_GetProc_RegEnableReflectionKey();
                break;
            }
            if (func) {
                result.value.ival = func(hkey);
                if (result.value.ival == ERROR_SUCCESS)
                    result.type = TRT_EMPTY;
            }
            else {
                result.value.ival = ERROR_PROC_NOT_FOUND;
            }
        }
        break;

    default:
        result.type = TRT_BADFUNCTIONCODE;
    }

    return TwapiSetResult(interp, &result);
}

static int TwapiRegInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s RegDispatch[] = {
        DEFINE_FNCODE_CMD(RegOpenKeyEx, 1),
        DEFINE_FNCODE_CMD(RegCreateKeyEx, 2),
        DEFINE_FNCODE_CMD(RegDeleteKeyValue, 3),
        DEFINE_FNCODE_CMD(RegDeleteKeyEx, 4),
        DEFINE_FNCODE_CMD(RegDeleteValue, 5),
        DEFINE_FNCODE_CMD(RegDeleteTree, 6),
        DEFINE_FNCODE_CMD(reg_key_prune, 6),
        DEFINE_FNCODE_CMD(RegEnumKeyEx, 7),
        DEFINE_FNCODE_CMD(reg_keys, 7),
        DEFINE_FNCODE_CMD(RegEnumValue, 8),
        DEFINE_FNCODE_CMD(RegOpenCurrentUser, 9),
        DEFINE_FNCODE_CMD(RegDisablePredefinedCache, 10),
        DEFINE_FNCODE_CMD(reg_disable_current_user_cache, 10),
        DEFINE_FNCODE_CMD(RegGetKeySecurity, 11),
        DEFINE_FNCODE_CMD(RegQueryValueEx, 12),
        DEFINE_FNCODE_CMD(RegCopyTree, 13),
        DEFINE_FNCODE_CMD(RegOpenUserClassesRoot, 14),
        DEFINE_FNCODE_CMD(RegOverridePredefKey, 15),
        DEFINE_FNCODE_CMD(reg_key_override, 15),
        DEFINE_FNCODE_CMD(RegSaveKeyEx, 16),
        DEFINE_FNCODE_CMD(RegLoadKey, 17),
        DEFINE_FNCODE_CMD(RegUnLoadKey, 18),
        DEFINE_FNCODE_CMD(RegReplaceKey, 19),
        DEFINE_FNCODE_CMD(reg_key_replace, 19), // TBD doc and test
        DEFINE_FNCODE_CMD(RegRestoreKey, 20),
        DEFINE_FNCODE_CMD(RegSetValueEx, 21),
        DEFINE_FNCODE_CMD(RegNotifyChangeKeyValue, 22),
        DEFINE_FNCODE_CMD(RegSetKeySecurity, 23),
        DEFINE_FNCODE_CMD(RegConnectRegistry, 24),
        DEFINE_FNCODE_CMD(reg_connect, 24),
        DEFINE_FNCODE_CMD(RegGetValue, 25),
        DEFINE_FNCODE_CMD(RegSetKeyValue, 26),
        DEFINE_FNCODE_CMD(SHCopyKey, 27),
        DEFINE_FNCODE_CMD(RegFlushKey, 28),
        DEFINE_FNCODE_CMD(reg_key_flush, 28),
        DEFINE_FNCODE_CMD(RegCloseKey, 29),
        DEFINE_FNCODE_CMD(reg_key_close, 29),
        DEFINE_FNCODE_CMD(RegDisableReflectionKey, 30),
        DEFINE_FNCODE_CMD(reg_key_disable_reflection, 30), // TBD doc and test
        DEFINE_FNCODE_CMD(RegEnableReflectionKey, 31),
        DEFINE_FNCODE_CMD(reg_key_enable_reflection, 31), // TBD doc and test
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(RegDispatch), RegDispatch, Twapi_RegCallObjCmd);

    return TCL_OK;
}

/* Main entry point */
#ifndef TWAPI_SINGLE_MODULE
__declspec(dllexport) 
#endif
int Twapi_registry_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiRegInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

