/*
 * Copyright (c) 2020, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

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

static int TwapiRegEnumKeyEx(Tcl_Interp *interp, HKEY hkey)
{
    Tcl_Obj *resultObj = NULL;
    FILETIME file_time;
    LPWSTR subkey;
    DWORD capacity_subkey;
    DWORD nch_subkey, dwIndex;
    LONG status;
    SWSMark mark;

    mark = SWSPushMark();
    resultObj = ObjNewList(0, NULL);
    dwIndex = 0;
    capacity_subkey = 256;
    subkey = SWSAlloc(sizeof(WCHAR) * capacity_subkey, &capacity_subkey);
    while (1) {
        nch_subkey = capacity_subkey;
        status     = RegEnumKeyExW(hkey, dwIndex, subkey, &nch_subkey, NULL,
                                   NULL, NULL, &file_time);
        if (status == ERROR_SUCCESS) {
            Tcl_Obj *objs[2];
            objs[0] = ObjFromTclUniCharN(subkey, nch_subkey);
            objs[1] = ObjFromFILETIME(&file_time);
            ObjAppendElement(interp, resultObj, ObjNewList(2, objs));
            ++dwIndex; /* Get next key */
        } else if (status == ERROR_MORE_DATA) {
            /* Need bigger buffer for this key */
            capacity_subkey *= 2;
            subkey = SWSAlloc(sizeof(WCHAR) * capacity_subkey, &capacity_subkey);
            /* Do not increment dwIndex */
        } else {
            /* ERROR_NO_MORE_ITEMS (done) or some other error */
            break;
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

static int TwapiRegEnumValue(Tcl_Interp *interp, HKEY hkey, DWORD flags)
{
    Tcl_Obj *resultObj = NULL;
    FILETIME file_time;
    LPWSTR   value_name;
    DWORD    capacity_value_name;
    DWORD    nch_value_name;
    DWORD    dwIndex;
    LONG  status;
    SWSMark  mark;

    mark = SWSPushMark();
    capacity_value_name = 32768; /* Max as per registry limits */
    value_name          = SWSAlloc(sizeof(WCHAR) * capacity_value_name, NULL);

    resultObj = ObjNewList(0, NULL);

    if (flags & 1) {
        /* Caller wants data as well. */
        LPBYTE value_data;
        DWORD capacity_value_data = 256;
        DWORD value_type, nb_value_data;
        int    max_loop; /* Safety measure if buf size keeps changing */
        value_data = SWSAlloc(capacity_value_data, &capacity_value_data);
        dwIndex    = 0;
        max_loop   = 5; /* Retries for a particular key. Else error */
        while (max_loop > 0) {
            nch_value_name = capacity_value_name;
            nb_value_data  = capacity_value_data;
            status         = RegEnumValueW(hkey,
                                   dwIndex,
                                   value_name,
                                   &nch_value_name,
                                   NULL,
                                   &value_type,
                                   value_data,
                                   &nb_value_data);
            if (status == ERROR_SUCCESS) {
                Tcl_Obj *objs[3];
                objs[0] = ObjFromTclUniCharN(value_name, nch_value_name);
                objs[1] = ObjFromDWORD(value_type);
                objs[2] = ObjFromRegValue(
                    interp, value_type, value_data, nb_value_data);
                ObjAppendElement(interp, resultObj, ObjNewList(3, objs));
                ++dwIndex;
                max_loop = 5; /* Reset safety check */
            }
            else if (status == ERROR_MORE_DATA) {
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
            nch_value_name = capacity_value_name;
            status         = RegEnumValueW(hkey,
                                   dwIndex,
                                   value_name,
                                   &nch_value_name,
                                   NULL,
                                   NULL,
                                   NULL,
                                   NULL);
            if (status != ERROR_SUCCESS)
                break;
            ObjAppendElement(interp, resultObj,
                             ObjFromTclUniCharN(value_name, nch_value_name));
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

static int Twapi_RegCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR subkey, value_name;
    HKEY   hkey, hkey2;
    TwapiResult result;
    DWORD dw, dw2, dw3, dw4;
    SECURITY_ATTRIBUTES *secattrP;
    SWSMark mark = NULL;
    Tcl_Obj *objs[2];
    int func = PtrToInt(clientdata);

    /* Every command has at least two arguments */
    if (objc < 3)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    --objc;
    ++objv;

    /* Assume error on system call */
    result.type = TRT_EXCEPTION_ON_ERROR;
    switch (func) {
    case 1: // RegOpenKeyEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hkey, HKEY), GETWSTR(subkey),
                         GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegOpenKeyExW(hkey, subkey, dw, dw2, &hkey2);
        if (result.value.ival == ERROR_SUCCESS) {
            result.type = TRT_HKEY;
            result.value.hval = hkey2;
        }
        break;

    case 2: // RegCreateKeyEx
        secattrP = NULL;
        mark = SWSPushMark();
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hkey, HKEY),
                         GETWSTR(subkey), GETINT(dw), ARGSKIP, GETINT(dw2),
                         GETINT(dw3),
                         GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTESSWS),
                         ARGEND) != TCL_OK) {
            SWSPopMark(mark);
            return TCL_ERROR;
        }
        result.value.ival = RegCreateKeyExW(hkey, subkey, dw, NULL,
                                            dw2, dw3, secattrP, &hkey2, &dw4);
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
                         GETHANDLET(hkey, HKEY),
                         GETWSTR(subkey), GETWSTR(value_name),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegDeleteKeyValueW(hkey, subkey, value_name);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 4: // RegDeleteKeyEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hkey, HKEY), GETWSTR(subkey), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegDeleteKeyExW(hkey, subkey, dw, 0);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;
    case 5: // RegDeleteValue
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hkey, HKEY), GETWSTR(value_name),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegDeleteValueW(hkey, value_name);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 6: // RegDeleteTree
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hkey, HKEY), GETWSTR(subkey),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegDeleteTreeW(hkey, subkey);
        if (result.value.ival == ERROR_SUCCESS)
            result.type = TRT_EMPTY;
        break;

    case 7: // RegEnumKeyEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hkey, HKEY),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_TCL_RESULT;
        result.value.ival = TwapiRegEnumKeyEx(interp, hkey);
        break;

    case 8: // RegEnumValue
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hkey, HKEY), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_TCL_RESULT;
        result.value.ival = TwapiRegEnumValue(interp, hkey, dw);
        break;

    case 9:
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = RegOpenCurrentUser(dw, &hkey);
        if (result.value.ival == ERROR_SUCCESS) {
            result.type = TRT_HKEY;
            result.value.hval = hkey;
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
        DEFINE_FNCODE_CMD(RegEnumKeyEx, 7),
        DEFINE_FNCODE_CMD(RegEnumValue, 8),
        DEFINE_FNCODE_CMD(RegOpenCurrentUser, 9),
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

