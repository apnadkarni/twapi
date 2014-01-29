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
static void TwapiPdhRestoreLocale()
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
        ObjSetResult(interp, ObjFromUnicode(buf));
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
    buf = MemLifoPushFrame(ticP->memlifoP, (1+buf_sz)*sizeof(WCHAR), NULL);
    status = PdhEnumObjectsW(szDataSource, szMachineName, buf, &buf_sz,
                            dwDetailLevel, 0);
    if (status == ERROR_SUCCESS)
        ObjSetResult(ticP->interp, ObjFromMultiSz(buf, buf_sz));
    else
        Twapi_AppendSystemError(ticP->interp, status);

    MemLifoPopFrame(ticP->memlifoP);

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

    counter_buf = MemLifoPushFrame(ticP->memlifoP, counter_buf_size*sizeof(*counter_buf), NULL);
    /* Note instance_buf_size may be 0 if no instances - see SDK */
    if (instance_buf_size)
        instance_buf = MemLifoAlloc(ticP->memlifoP,
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
        ObjSetResult(ticP->interp,
                         ObjNewList((instance_buf_size ? 2 : 1), objs));
    }

    MemLifoPopFrame(ticP->memlifoP);
    return pdh_status == ERROR_SUCCESS ?
        TCL_OK : Twapi_AppendSystemError(ticP->interp, pdh_status);
}



int Twapi_PdhMakeCounterPath (TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    DWORD   dwFlags;
    PDH_COUNTER_PATH_ELEMENTS_W pdh_elements;
    PDH_STATUS  pdh_status;
    WCHAR      *path_buf; 
    DWORD       path_buf_size;
    TCL_RESULT  result;
    MemLifoMarkHandle mark;

    mark = MemLifoPushMark(ticP->memlifoP);

    result = TwapiGetArgsEx(ticP, objc, objv,
                            GETEMPTYASNULL(pdh_elements.szMachineName),
                            GETWSTR(pdh_elements.szObjectName),
                            GETEMPTYASNULL(pdh_elements.szInstanceName),
                            GETEMPTYASNULL(pdh_elements.szParentInstance),
                            GETINT(pdh_elements.dwInstanceIndex),
                            GETWSTR(pdh_elements.szCounterName),
                            GETINT(dwFlags),
                            ARGEND);
    if (result == TCL_OK) {
        path_buf_size = 0;
        pdh_status = PdhMakeCounterPathW(&pdh_elements, NULL,
                                         &path_buf_size, dwFlags);
        if ((pdh_status != ERROR_SUCCESS) && (pdh_status != PDH_MORE_DATA)) {
            result = Twapi_AppendSystemError(ticP->interp, pdh_status);
        } else {
            path_buf = MemLifoAlloc(ticP->memlifoP,
                                    path_buf_size*sizeof(*path_buf), NULL);
            pdh_status = PdhMakeCounterPathW(&pdh_elements, path_buf,
                                             &path_buf_size, dwFlags);
            if (pdh_status != ERROR_SUCCESS)
                result = Twapi_AppendSystemError(ticP->interp, pdh_status);
            else
                ObjSetResult(ticP->interp, ObjFromUnicode(path_buf));
        }
    }

    MemLifoPopMark(mark);
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
        pdh_elems = MemLifoPushFrame(ticP->memlifoP, buf_size, NULL);
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
            objs[9] = ObjFromLong(pdh_elems->dwInstanceIndex);
            objs[10] = STRING_LITERAL_OBJ("szCounterName");
            objs[11] = ObjFromUnicode(pdh_elems->szCounterName);
            ObjSetResult(ticP->interp, ObjNewList(12, objs));
            result = TCL_OK;
        }
        MemLifoPopFrame(ticP->memlifoP);
    }

    return result;
}

static Tcl_Obj *ObjFromPDH_FMT_COUNTERVALUE(Tcl_Interp *interp, PDH_FMT_COUNTERVALUE *valP, DWORD dwFormat)
{
    if (valP->CStatus == ERROR_SUCCESS) {
        switch (dwFormat & (PDH_FMT_LARGE | PDH_FMT_DOUBLE | PDH_FMT_LONG)) {
        case PDH_FMT_LONG:   return ObjFromLong(valP->longValue);
        case PDH_FMT_LARGE:  return ObjFromWideInt(valP->largeValue);
        case PDH_FMT_DOUBLE: return Tcl_NewDoubleObj(valP->doubleValue);
        default:
            ObjSetResult(interp,
                         Tcl_ObjPrintf("Invalid PDH counter format value 0x%x",
                                       dwFormat));
            return NULL;
        }
    }
    Twapi_AppendSystemErrorEx(interp, valP->CStatus, Tcl_ObjPrintf("Counter value status failure %d", valP->CStatus));
    return NULL;
}

static Tcl_Obj *ObjFromPDH_COUNTER_VALUE_ITEM(Tcl_Interp *interp, PDH_FMT_COUNTERVALUE_ITEM_W *itemP, DWORD dwFormat)
{
    Tcl_Obj *objs[2];
    objs[1] = ObjFromPDH_FMT_COUNTERVALUE(interp, &itemP->FmtValue, dwFormat);
    if (objs[1] == NULL)
        return NULL;            /* interp already holds error */
        
    objs[0] = ObjFromUnicode(itemP->szName);
    return ObjNewList(ARRAYSIZE(objs), objs);
}

TCL_RESULT Twapi_PdhGetFormattedCounterArray(
    Tcl_Interp *interp,
    HANDLE hCounter,
    DWORD dwFormat
    )
{
    PDH_STATUS pdh_status;
    DWORD sz, nitems;
    PDH_FMT_COUNTERVALUE_ITEM_W *itemP;
    TCL_RESULT res;
    int i;
    MemLifo *memlifoP;
    MemLifoMarkHandle mark;

    memlifoP = TwapiMemLifo();
    mark = MemLifoPushMark(memlifoP);
    
    /* Always get required size first since docs say do not rely on
     * returned size for allocation if it was not passed in as 0.
     */
    sz = 0;
    pdh_status = PdhGetFormattedCounterArrayW(hCounter, dwFormat, &sz, &nitems, NULL);
    /* Number of items might change so try a few times in a loop */
    for (i = 0; i < 10 && pdh_status == PDH_MORE_DATA; ++i) {
        itemP = MemLifoPushFrame(memlifoP, sz, NULL);
        pdh_status = PdhGetFormattedCounterArrayW(hCounter, dwFormat, &sz, &nitems, itemP);
    }
    if (pdh_status != ERROR_SUCCESS)
        res = Twapi_AppendSystemError(interp, pdh_status);
    else {
        Tcl_Obj **itemObjs = MemLifoAlloc(memlifoP, nitems * sizeof(*itemObjs), NULL);
        res = TCL_OK;
        for (i = 0; i < nitems; ++i) {
            itemObjs[i] = ObjFromPDH_COUNTER_VALUE_ITEM(interp, &itemP[i], dwFormat);
            if (itemObjs[i] == NULL) {
                ObjDecrArrayRefs(i, itemObjs);
                res = TCL_ERROR;
                break;
            }
        }
        if (res == TCL_OK)
            ObjSetResult(interp, ObjNewList(nitems, itemObjs));
    }
    MemLifoPopMark(mark);
    return res;
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
        ObjSetResult(interp,
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
        objs[0] = ObjFromLong(counter_value.longValue);
        break;

    case PDH_FMT_LARGE:
        objs[0] = ObjFromWideInt(counter_value.largeValue);
        break;

    case PDH_FMT_DOUBLE:
        objs[0] = Tcl_NewDoubleObj(counter_value.doubleValue);
        break;

    default:
        ObjSetResult(interp,
                         Tcl_ObjPrintf("Invalid PDH counter format value 0x%x",
                                       dwFormat));
        return  TCL_ERROR;
    }
    
    objs[1] = Tcl_ObjPrintf("0x%x", counter_type);

    /* Create the result list consisting of type and value */
    ObjSetResult(interp, ObjNewList(2, objs));
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




    /* 
     * This function is not documented because of the following bugs
     * See http://helgeklein.com/blog/2009/10/found-my-first-bug-in-windows-7-apis-pdhbrowsecounters-requires-elevation/
     *
     * PDH.DLL on Vista & Win 7 is rife with bugs, this is just one of them
     *
     * You dont need elevation, you can just set this flag to false
     * and it should work: bSingleCounterPerAdd = FALSE
     * 
     * Also, the flag bWildCardInstances must ALWAYS be set to TRUE,
     * otherwise it causes crashes in many cases. The first one would
     * be nice to be able to use (unelevated), but you will just have to
     * ignore any secondary counters that are selected if a user
     * selects more than one.
     * 
     * Other bugs: - You need to manually set the title, it no longer
     * offers a default (NULL pointer).  - You cant set an initial
     * performance counter to be highlighted.  - Sometimes the
     * Instances wont show up on the bottom, depending on where you
     * click on the top portion of the dialog box. This causes an
     * invalid Counter path to be returned.  - PDH.DLL and 40 other
     * DLLs are loaded, and STAY loaded once PdhBrowseCounters is
     * called. Manually unloading the DLLs often causes crashes. This
     * may or may not have something to do with .NET 4.0 being
     * present.
     * 
     * Other general PDH bugs: - PdhExpandWildCardPath will not recognize
     * any new Processes (and probably other instances for other object
     * types), unless PDH.DLL is unloaded and reloaded. This also poses a
     * problem if you use PdhBrowseCounters, as it increases the PDH.DLL
     * load-count by 2. Forcing an un-load may be dangerous, but seems the
     * only option.
     * 
     * My system: Win 7 x64, .Net 4.0, all updates as of 6/1/2011
     */

    TwapiZeroMemory(&browse_dlg, sizeof(browse_dlg));
    browse_dlg.bIncludeInstanceIndex = 1;
    browse_dlg.bSingleCounterPerDialog = 1;
    browse_dlg.bSingleCounterPerAdd = 1; /* See bug list above */
    browse_dlg.bWildCardInstances = 1;   /* See bug list above */
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

    ObjSetResult(interp, ObjFromMultiSz(browse_dlg.szReturnPathBuffer,
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
    LPWSTR s;
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
        CHECK_NARGS(interp, objc, 3);

        switch (func) {
        case 101:
            if (ObjToLong(interp, objv[2], &dw) != TCL_OK)
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
            result.value.ival = PdhValidatePathW(ObjToUnicode(objv[2]));
            break;
        }
    } else if (func < 300) {
        /* Single string with integer arg. */
        CHECK_NARGS(interp, objc, 4);
        /* To prevent shimmering issues, get int arg first */
        CHECK_INTEGER_OBJ(interp, dw, objv[3]);
        s = ObjToUnicode(objv[2]);
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
        case 1006:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return (func == 1001 ? Twapi_PdhGetFormattedCounterValue : Twapi_PdhGetFormattedCounterArray)(interp, h, dw);
        case 1002:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), ARGSKIP, GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            dw = PdhAddCounterW(h, ObjToUnicode(objv[3]), dw, &result.value.hval);
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
                             ARGSKIP, ARGSKIP, ARGSKIP,
                             GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhEnumObjectItems(
                ticP,
                ObjToLPWSTR_NULL_IF_EMPTY(objv[2]),
                ObjToLPWSTR_NULL_IF_EMPTY(objv[3]),
                ObjToUnicode(objv[4]),
                dw, dw2);
        case 1005:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             ARGSKIP, ARGSKIP,
                             GETINT(dw), GETBOOL(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_PdhEnumObjects(
                ticP,
                ObjToLPWSTR_NULL_IF_EMPTY(objv[2]),
                ObjToLPWSTR_NULL_IF_EMPTY(objv[3]),
                dw, dw2);
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

static int TwapiPdhInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct alias_dispatch_s PdhDispatch[] = {
        DEFINE_ALIAS_CMD(PdhGetDllVersion, 1),
        DEFINE_ALIAS_CMD(PdhBrowseCounters, 2),
        DEFINE_ALIAS_CMD(PdhSetDefaultRealTimeDataSource, 101),
        DEFINE_ALIAS_CMD(PdhConnectMachine, 102),
        DEFINE_ALIAS_CMD(PdhValidatePath, 103),
        DEFINE_ALIAS_CMD(PdhParseCounterPath, 201),
        DEFINE_ALIAS_CMD(PdhLookupPerfNameByIndex, 202),
        DEFINE_ALIAS_CMD(PdhOpenQuery, 203),
        DEFINE_ALIAS_CMD(PdhRemoveCounter, 301),
        DEFINE_ALIAS_CMD(PdhCollectQueryData, 302),
        DEFINE_ALIAS_CMD(PdhCloseQuery, 303),
        DEFINE_ALIAS_CMD(PdhGetFormattedCounterValue, 1001),
        DEFINE_ALIAS_CMD(PdhAddCounter, 1002),
        DEFINE_ALIAS_CMD(PdhMakeCounterPath, 1003),
        DEFINE_ALIAS_CMD(PdhEnumObjectItems, 1004),
        DEFINE_ALIAS_CMD(PdhEnumObjects, 1005),
        DEFINE_ALIAS_CMD(PdhGetFormattedCounterArray, 1006),
    };

    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::CallPdh", Twapi_CallPdhObjCmd, ticP, NULL);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(PdhDispatch), PdhDispatch, "twapi::CallPdh");

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
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiPdhInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}


                           
