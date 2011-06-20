/* 
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

/* Define interfaces to the PDH performance monitoring library */

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
        Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(buf, -1));
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
        Tcl_SetObjResult(ticP->interp, Tcl_NewUnicodeObj(path_buf, -1));
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
            objs[1] = Tcl_NewUnicodeObj(pdh_elems->szMachineName, -1);
            objs[2] = STRING_LITERAL_OBJ("szObjectName");
            objs[3] = Tcl_NewUnicodeObj(pdh_elems->szObjectName, -1);
            objs[4] = STRING_LITERAL_OBJ("szInstanceName");
            objs[5] = Tcl_NewUnicodeObj(pdh_elems->szInstanceName, -1);
            objs[6] = STRING_LITERAL_OBJ("szParentInstance");
            objs[7] = Tcl_NewUnicodeObj(pdh_elems->szParentInstance, -1);
            objs[8] = STRING_LITERAL_OBJ("dwInstanceIndex");
            objs[9] = Tcl_NewLongObj(pdh_elems->dwInstanceIndex);
            objs[10] = STRING_LITERAL_OBJ("szCounterName");
            objs[11] = Tcl_NewUnicodeObj(pdh_elems->szCounterName, -1);
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
PdhExpandCounterPath has a bug on Win2K. So we do not wrap it;
#endif


                           
