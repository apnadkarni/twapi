/* 
 * Copyright (c) 2003-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include <pdhmsg.h>
#include <pdh.h>         /* Include AFTER lm.h due to HLOG def conflict */

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

/* Define interfaces to the PDH performance monitoring library */

/* PDH */
void TwapiPdhRestoreLocale(void);
int Twapi_PdhParseCounterPath(TwapiInterpContext *, LPCWSTR buf, DWORD dwFlags);
int Twapi_PdhGetFormattedCounterValue(Tcl_Interp *, HANDLE hCtr, DWORD fmt);
int Twapi_PdhLookupPerfNameByIndex(Tcl_Interp *,  LPCWSTR machine, DWORD ctr);
int Twapi_PdhMakeCounterPath (TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[]);
int Twapi_PdhBrowseCounters(Tcl_Interp *interp);
int Twapi_PdhEnumObjects(TwapiInterpContext *ticP,
                         LPCWSTR source, LPCWSTR machine,
                         DWORD  dwDetailLevel, BOOL bRefresh);
int Twapi_PdhEnumObjectItems(TwapiInterpContext *,
                             LPCWSTR source, LPCWSTR machine,
                              LPCWSTR objname, DWORD detail, DWORD dwFlags);

#if (TCL_MAJOR_VERSION == 8 && TCL_MINOR_VERSION < 5)
/* Reset locale back to C if necessary. */
void TwapiPdhRestoreLocale()
{
    /*
     * PDH calls do a setlocale which causes Tcl's expression parsing to
     * fail if the decimal separator in the locale is not ".". To get
     * around this, we reset the locale after every PDH call. Tcl 8.5 and
     * later or OS XP or later apparently do not need this.
     */
    if (gTclVersion.major > 8 ||
        (gTclVersion.major == 8 && gTclVersion.minor > 4))
        return;

    /* Note test below assumes twapi does not load at all if major < 5 */
    if (gTwapiOSVersionInfo.dwMajorVersion > 5 ||
        gTwapiOSVersionInfo.dwMinorVersion != 0)
        return;

    setlocale(LC_ALL, "C");
}
#endif

int Twapi_PdhLookupPerfNameByIndex(
    Tcl_Interp *interp,
    LPCWSTR szMachineName,
    DWORD ctr_index)
{
    WCHAR buf[1024]; /* PDH_MAX_COUNTER_NAME on newer SDK's */
    DWORD bufsz;
    PDH_STATUS status;

    bufsz=ARRAYSIZE(buf);
    status = PdhLookupPerfNameByIndexW(szMachineName, ctr_index, buf, &bufsz);
    if (status == ERROR_SUCCESS) {
        Tcl_SetObjResult(interp, ObjFromUnicode(buf));
        return TCL_OK;
    }
    else {
        return Twapi_AppendSystemError(interp, status);
    }
}

int Twapi_PdhEnumObjects(
    TwapiInterpContext *ticP,
    LPCWSTR szDataSource,
    LPCWSTR szMachineName,
    DWORD   dwDetailLevel,
    BOOL    bRefresh
)
{
    PDH_STATUS status;
    LPWSTR     buf = NULL;
    DWORD      buf_sz;

    /*
     * NOTE : The first call MUST pass 0 as buffer size else call
     * the function does not correctly return required size if buffer
     * is not big enough.
     */
    buf_sz = 0;
    status = PdhEnumObjectsW(szDataSource, szMachineName, NULL, &buf_sz,
                            dwDetailLevel, bRefresh);

    if ((status != ERROR_SUCCESS) && status != PDH_MORE_DATA)
        return Twapi_AppendSystemError(ticP->interp, status);

    /*
     * Call again with a buffer. Note this time bRefresh is passed
     * as 0, as we want to get the data we already asked about else
     * the buffer size may be invalid
    */
    buf = MemLifoPushFrame(&ticP->memlifo, (1+buf_sz)*sizeof(WCHAR), NULL);
    status = PdhEnumObjectsW(szDataSource, szMachineName, buf, &buf_sz,
                            dwDetailLevel, 0);
    if (status == ERROR_SUCCESS)
        Tcl_SetObjResult(ticP->interp, ObjFromMultiSz(buf, buf_sz));
    else
        Twapi_AppendSystemError(ticP->interp, status);

    MemLifoPopFrame(&ticP->memlifo);

    return status == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
}


int Twapi_PdhEnumObjectItems(TwapiInterpContext *ticP,
                             LPCWSTR szDataSource, LPCWSTR szMachineName,
                             LPCWSTR szObjectName, DWORD dwDetailLevel,
                             DWORD dwFlags)
{
    PDH_STATUS pdh_status;
    WCHAR *counter_buf;
    DWORD  counter_buf_size;
    WCHAR *instance_buf;
    DWORD  instance_buf_size;
    Tcl_Obj *objs[2];           /* 0 - counter, 1 - instance */
    
    counter_buf = NULL;
    instance_buf = NULL;
    objs[0] = NULL;
    objs[1] = NULL;

    /*
     * First make a call to figure out how much space is required.
     * Note: do not preallocate as then the call does not correctly return
     * the required size
     */
    counter_buf_size  = 0;
    instance_buf_size = 0;
    pdh_status = PdhEnumObjectItemsW(
        szDataSource ,szMachineName, szObjectName, NULL, &counter_buf_size,
        NULL, &instance_buf_size, dwDetailLevel, dwFlags);

    if ((pdh_status != ERROR_SUCCESS) && (pdh_status != PDH_MORE_DATA)) {
        return Twapi_AppendSystemError(ticP->interp, pdh_status);
    }

    counter_buf = MemLifoPushFrame(&ticP->memlifo, counter_buf_size*sizeof(*counter_buf), NULL);
    /* Note instance_buf_size may be 0 if no instances - see SDK */
    if (instance_buf_size)
        instance_buf = MemLifoAlloc(&ticP->memlifo,
                                    instance_buf_size*sizeof(*instance_buf),
                                    NULL);

    pdh_status = PdhEnumObjectItemsW(
        szDataSource ,szMachineName, szObjectName,
        counter_buf, &counter_buf_size,
        instance_buf, &instance_buf_size,
        dwDetailLevel, dwFlags);
        
    if (pdh_status == ERROR_SUCCESS) {
        /*
         * Format and return as a list of two lists if both counters and
         * instances are present, else a list of one list in only counters
         * are present.
         */ 
        objs[0] = ObjFromMultiSz(counter_buf, counter_buf_size);
        if (instance_buf_size)
            objs[1] = ObjFromMultiSz(instance_buf, instance_buf_size);
        Tcl_SetObjResult(ticP->interp,
                         Tcl_NewListObj((instance_buf_size ? 2 : 1), objs));
    }

    MemLifoPopFrame(&ticP->memlifo);
    return pdh_status == ERROR_SUCCESS ?
        TCL_OK : Twapi_AppendSystemError(ticP->interp, pdh_status);
}



int Twapi_PdhMakeCounterPath (TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR szMachineName;
    LPWSTR szObjectName;
    LPWSTR szInstanceName;
    LPWSTR szParentInstance;
    DWORD   dwInstanceIndex;
    LPWSTR szCounterName;
    DWORD   dwFlags;

    PDH_COUNTER_PATH_ELEMENTS_W pdh_elements;
    PDH_STATUS  pdh_status;
    WCHAR      *path_buf; 
    DWORD       path_buf_size;
    int         result;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETNULLIFEMPTY(szMachineName), GETWSTR(szObjectName),
                     GETNULLIFEMPTY(szInstanceName),
                     GETNULLIFEMPTY(szParentInstance),
                     GETINT(dwInstanceIndex), GETWSTR(szCounterName),
                     GETINT(dwFlags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;


    pdh_elements.szMachineName    = szMachineName;
    pdh_elements.szObjectName     = szObjectName;
    pdh_elements.szInstanceName   = szInstanceName;
    pdh_elements.szParentInstance = szParentInstance;
    pdh_elements.dwInstanceIndex  = dwInstanceIndex;
    pdh_elements.szCounterName    = szCounterName;
    
    path_buf_size = 0;
    pdh_status = PdhMakeCounterPathW(&pdh_elements, NULL,
                                     &path_buf_size, dwFlags);
    if ((pdh_status != ERROR_SUCCESS) && (pdh_status != PDH_MORE_DATA)) {
        return Twapi_AppendSystemError(ticP->interp, pdh_status);
    }
    
    path_buf = MemLifoPushFrame(&ticP->memlifo,
                                path_buf_size*sizeof(*path_buf), NULL);
    pdh_status = PdhMakeCounterPathW(&pdh_elements, path_buf,
                                     &path_buf_size, dwFlags);
    if (pdh_status != ERROR_SUCCESS) {
        Twapi_AppendSystemError(ticP->interp, pdh_status);
        result = TCL_ERROR;
    }
    else {
        Tcl_SetObjResult(ticP->interp, ObjFromUnicode(path_buf));
        result = TCL_OK;
    }
    MemLifoPopFrame(&ticP->memlifo);

    return result;

}


int Twapi_PdhParseCounterPath(
    TwapiInterpContext *ticP,
    LPCWSTR szFullPathBuffer,
    DWORD   dwFlags
)
{
    PDH_STATUS           pdh_status;
    PDH_COUNTER_PATH_ELEMENTS_W *pdh_elems = NULL;
    DWORD                buf_size;
    int                  result;
    
    buf_size = 0;
    pdh_status = PdhParseCounterPathW(szFullPathBuffer, NULL,
                                     &buf_size, dwFlags);
    if ((pdh_status != ERROR_SUCCESS) && (pdh_status != PDH_MORE_DATA)) {
        Twapi_AppendSystemError(ticP->interp, pdh_status);
        result = TCL_ERROR;
    }
    else {
        pdh_elems = MemLifoPushFrame(&ticP->memlifo, buf_size, NULL);
        pdh_status = PdhParseCounterPathW(szFullPathBuffer, pdh_elems,
                                          &buf_size, dwFlags);
    
        if (pdh_status != ERROR_SUCCESS) {
            Twapi_AppendSystemError(ticP->interp, pdh_status);
            result = TCL_ERROR;
        } else {
            Tcl_Obj *objs[12];
            objs[0] = STRING_LITERAL_OBJ("szMachineName");
            objs[1] = ObjFromUnicode(pdh_elems->szMachineName);
            objs[2] = STRING_LITERAL_OBJ("szObjectName");
            objs[3] = ObjFromUnicode(pdh_elems->szObjectName);
            objs[4] = STRING_LITERAL_OBJ("szInstanceName");
            objs[5] = ObjFromUnicode(pdh_elems->szInstanceName);
            objs[6] = STRING_LITERAL_OBJ("szParentInstance");
            objs[7] = ObjFromUnicode(pdh_elems->szParentInstance);
            objs[8] = STRING_LITERAL_OBJ("dwInstanceIndex");
            objs[9] = Tcl_NewLongObj(pdh_elems->dwInstanceIndex);
            objs[10] = STRING_LITERAL_OBJ("szCounterName");
            objs[11] = ObjFromUnicode(pdh_elems->szCounterName);
            Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(12, objs));
            result = TCL_OK;
        }
        MemLifoPopFrame(&ticP->memlifo);
    }

    return result;
}


int Twapi_PdhGetFormattedCounterValue(
    Tcl_Interp *interp,
    HANDLE hCounter,
    DWORD dwFormat
)
{
    PDH_STATUS           pdh_status;
    DWORD                value_type;
    DWORD                counter_type;
    PDH_FMT_COUNTERVALUE counter_value;
    Tcl_Obj             *objs[2]; /* Holds type and value, result list */
    
    pdh_status = PdhGetFormattedCounterValue(hCounter, dwFormat,
                                             &counter_type, &counter_value);

    if ((pdh_status != ERROR_SUCCESS)
        || (counter_value.CStatus != ERROR_SUCCESS)) {
        Tcl_SetObjResult(interp,
                         Tcl_ObjPrintf("Error (0x%x/0x%x) retrieving counter value: ", pdh_status, counter_value.CStatus));
        return Twapi_AppendSystemError(interp,
                                       (counter_value.CStatus != ERROR_SUCCESS ?
                                        counter_value.CStatus : pdh_status));
    }

    objs[0] = NULL;
    objs[1] = NULL;

    value_type = dwFormat & (PDH_FMT_LARGE | PDH_FMT_DOUBLE | PDH_FMT_LONG);
    switch (value_type) {
    case PDH_FMT_LONG:
        objs[0] = Tcl_NewLongObj(counter_value.longValue);
        break;

    case PDH_FMT_LARGE:
        objs[0] = Tcl_NewWideIntObj(counter_value.largeValue);
        break;

    case PDH_FMT_DOUBLE:
        objs[0] = Tcl_NewDoubleObj(counter_value.doubleValue);
        break;

    default:
        Tcl_SetObjResult(interp,
                         Tcl_ObjPrintf("Invalid PDH counter format value 0x%x",
                                       dwFormat));
        return  TCL_ERROR;
    }
    
    objs[1] = Tcl_ObjPrintf("0x%x", counter_type);

    /* Create the result list consisting of type and value */
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objs));
    return TCL_OK;
}


PDH_STATUS __stdcall Twapi_CounterPathCallback(DWORD_PTR dwArg) 
{
    PDH_BROWSE_DLG_CONFIG_W *pbrowse_dlg = (PDH_BROWSE_DLG_CONFIG_W *) dwArg;

    if (pbrowse_dlg->CallBackStatus == PDH_MORE_DATA) {
        WCHAR *buf;
        DWORD new_size;

        new_size = 2 * pbrowse_dlg->cchReturnPathLength;
        buf = TwapiReallocTry(pbrowse_dlg->szReturnPathBuffer, new_size*sizeof(WCHAR));
        if (buf) {
            pbrowse_dlg->szReturnPathBuffer = buf;
            pbrowse_dlg->cchReturnPathLength = new_size;
            pbrowse_dlg->CallBackStatus = PDH_RETRY;
            return PDH_RETRY;
        } else {
            pbrowse_dlg->CallBackStatus = PDH_MEMORY_ALLOCATION_FAILURE;
            return PDH_MEMORY_ALLOCATION_FAILURE;
        }
    }

    /* Just return the status we got. Nothing to do */
    return pbrowse_dlg->CallBackStatus;
}

int Twapi_PdhBrowseCounters(Tcl_Interp *interp)
{
    PDH_BROWSE_DLG_CONFIG_W browse_dlg;
    PDH_STATUS pdh_status;

    TwapiZeroMemory(&browse_dlg, sizeof(browse_dlg));
    browse_dlg.bIncludeInstanceIndex = 1;
    browse_dlg.bSingleCounterPerDialog = 1;
    browse_dlg.cchReturnPathLength = 1000;
    /* Note we cannot use memlifo here as a callback is involved */
    browse_dlg.szReturnPathBuffer =
        TwapiAlloc(browse_dlg.cchReturnPathLength * sizeof(WCHAR));
    browse_dlg.pCallBack = Twapi_CounterPathCallback;
    browse_dlg.dwCallBackArg = (DWORD_PTR) &browse_dlg;
    browse_dlg.dwDefaultDetailLevel= PERF_DETAIL_WIZARD;
  
    pdh_status = PdhBrowseCountersW(&browse_dlg);

    if ( pdh_status != ERROR_SUCCESS) {
        TwapiFree(browse_dlg.szReturnPathBuffer);
        return Twapi_AppendSystemError(interp, pdh_status);
    }

    Tcl_SetObjResult(interp, ObjFromMultiSz(browse_dlg.szReturnPathBuffer,
                                            browse_dlg.cchReturnPathLength));
    TwapiFree(browse_dlg.szReturnPathBuffer);
    return TCL_OK;
}

#if 0
PdhExpandCounterPath has a bug on Win2K. So we do not wrap it; TBD
#endif


/* Call PDH API. This is special-cased because we have to do a restore
   locale after every PDH call on some platforms */
int Twapi_CallPdhObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s, s2, s3;
    DWORD   dw, dw2;
    HANDLE h;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 100) {
        /* No arguments */
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 1:
            dw = PdhGetDllVersion(&result.value.uval);
            if (dw == 0)
                result.type = TRT_DWORD;
            else {
                result.value.ival = dw;
                result.type = TRT_EXCEPTION_ON_ERROR;
            }
            break;
        case 2:
            return Twapi_PdhBrowseCounters(interp);
        }
    } else if (func < 200) {
        /* Single argument */
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 101:
            if (Tcl_GetLongFromObj(interp, objv[2], &dw) != TCL_OK)
                return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhSetDefaultRealTimeDataSource(dw);
            break;
        case 102:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhConnectMachineW(ObjToLPWSTR_NULL_IF_EMPTY(objv[2]));
            break;
        case 103:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhValidatePathW(Tcl_GetUnicode(objv[2]));
            break;
        }
    } else if (func < 300) {
        /* Single string with integer arg */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 201: 
            return Twapi_PdhParseCounterPath(ticP, s, dw);
        case 202: 
            NULLIFY_EMPTY(s);
            return Twapi_PdhLookupPerfNameByIndex(interp, s, dw);
        case 203:
            NULLIFY_EMPTY(s);
            dw = PdhOpenQueryW(s, dw, &result.value.hval);
            if (dw == 0)
                result.type = TRT_HANDLE;
            else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw;
            }
            break;
        }
    } else if (func < 400) {
        /* Single handle */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 301:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhRemoveCounter(h);
            break;
        case 302:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhCollectQueryData(h);
            break;
        case 303:
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = PdhCloseQuery(h);
            break;
        }
    } else {
        /* Free for all */
        switch (func) {
        case 1001:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhGetFormattedCounterValue(interp, h, dw);
        case 1002:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWSTR(s), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            dw = PdhAddCounterW(h, s, dw, &result.value.hval);
            if (dw == 0)
                result.type = TRT_HANDLE;
            else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw;
            }
            break;
        case 1003:
            return Twapi_PdhMakeCounterPath(ticP, objc-2, objv+2);
        case 1004:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETWSTR(s3), GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhEnumObjectItems(ticP, s, s2, s3, dw, dw2);
        case 1005:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETNULLIFEMPTY(s), GETNULLIFEMPTY(s2),
                             GETINT(dw), GETBOOL(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhEnumObjects(ticP, s, s2, dw, dw2);
        }
    }

    /* Set Tcl status before restoring locale as the latter might change
       value of GetLastError() */
    dw = TwapiSetResult(interp, &result);

#if (TCL_MAJOR_VERSION == 8 && TCL_MINOR_VERSION < 5)
    TwapiPdhRestoreLocale();
#endif

    return dw;
}

static int Twapi_PdhInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::CallPdh", Twapi_CallPdhObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::" #call_, # code_); \
    } while (0);

    CALL_(PdhGetDllVersion, CallPdh, 1);
    CALL_(PdhBrowseCounters, CallPdh, 2);
    CALL_(PdhSetDefaultRealTimeDataSource, CallPdh, 101);
    CALL_(PdhConnectMachine, CallPdh, 102);
    CALL_(PdhValidatePath, CallPdh, 103);
    CALL_(PdhParseCounterPath, CallPdh, 201);
    CALL_(PdhLookupPerfNameByIndex, CallPdh, 202);
    CALL_(PdhOpenQuery, CallPdh, 203);
    CALL_(PdhRemoveCounter, CallPdh, 301);
    CALL_(PdhCollectQueryData, CallPdh, 302);
    CALL_(PdhCloseQuery, CallPdh, 303);
    CALL_(PdhGetFormattedCounterValue, CallPdh, 1001);
    CALL_(PdhAddCounter, CallPdh, 1002);
    CALL_(PdhMakeCounterPath, CallPdh, 1003);
    CALL_(PdhEnumObjectItems, CallPdh, 1004);
    CALL_(PdhEnumObjects, CallPdh, 1005);

#undef CALL_

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
int Twapi_pdh_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, WLITERAL(MODULENAME), MODULE_HANDLE,
                            Twapi_PdhInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}


                           
