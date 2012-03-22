/* 
 * Copyright (c) 2004-2012 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to process information */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

HRESULT Twapi_SHGetFolderPath(HWND hwndOwner, int nFolder, HANDLE hToken,
                          DWORD flags, WCHAR *pathbuf);
BOOL Twapi_SHObjectProperties(HWND hwnd, DWORD dwType,
                              LPCWSTR szObject, LPCWSTR szPage);

int Twapi_GetShellVersion(Tcl_Interp *interp);
int Twapi_ShellExecuteEx(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ReadShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_WriteShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ReadUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath);
int Twapi_WriteUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath, LPCWSTR url, DWORD flags);
int Twapi_InvokeUrlShortcut(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_SHFileOperation(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);

static DLLVERSIONINFO *TwapiShellVersion()
{
    static DLLVERSIONINFO shellver;
    static int initialized = 0;
    
    if (! initialized) {
        TwapiGetDllVersion("shell32.dll", &shellver);
        initialized = 1;
    }
    return &shellver;
}

int Twapi_GetShellVersion(Tcl_Interp *interp)
{
    DLLVERSIONINFO *ver = TwapiShellVersion();
    Tcl_SetObjResult(interp,
                     Tcl_ObjPrintf("%u.%u.%u",
                                   ver->dwMajorVersion,
                                   ver->dwMinorVersion,
                                   ver->dwBuildNumber));
    return TCL_OK;
}


/* Even though ShGetFolderPath exists on Win2K or later, the VC 6 does
   not like the format of the SDK shell32.lib so we have to stick
   with the VC6 shell32, which does not have this function. So we are
   forced to dynamically load it.
*/
typedef HRESULT (WINAPI *SHGetFolderPathW_t)(HWND, int, HANDLE, DWORD, LPWSTR);
MAKE_DYNLOAD_FUNC(SHGetFolderPathW, shell32, SHGetFolderPathW_t)
HRESULT Twapi_SHGetFolderPath(
    HWND hwndOwner,
    int nFolder,
    HANDLE hToken,
    DWORD flags,
    WCHAR *pathbuf              /* Must be MAX_PATH */
)
{
    SHGetFolderPathW_t SHGetFolderPathPtr = Twapi_GetProc_SHGetFolderPathW();

    if (SHGetFolderPathPtr == NULL) {
        SetLastError(ERROR_PROC_NOT_FOUND);
        return ERROR_PROC_NOT_FOUND;
    }

    return (*SHGetFolderPathPtr)(hwndOwner, nFolder, hToken, flags, pathbuf);
}

typedef BOOL (WINAPI *SHObjectProperties_t)(HWND, DWORD, PCWSTR, PCWSTR);
MAKE_DYNLOAD_FUNC(SHObjectProperties, shell32, SHObjectProperties_t)
MAKE_DYNLOAD_FUNC_ORDINAL(178, shell32)
BOOL Twapi_SHObjectProperties(
    HWND hwnd,
    DWORD dwType,
    LPCWSTR szObject,
    LPCWSTR szPage
)
{
    static SHObjectProperties_t fnSHObjectProperties;
    static int initialized = 0;

    if (! initialized) {
        fnSHObjectProperties = Twapi_GetProc_SHObjectProperties();
        if (fnSHObjectProperties == NULL) {
            /*
             * Could not get function by name. On Win 2K, function is
             * available but not by name. Try getting by ordinal
             * after making sure shell version is 5.0
             */
            DLLVERSIONINFO *ver = TwapiShellVersion();
            if (ver->dwMajorVersion == 5 && ver->dwMinorVersion == 0) {
                fnSHObjectProperties = (SHObjectProperties_t) Twapi_GetProc_shell32_178();
            }
        }
        initialized = 1;
    }

    if (fnSHObjectProperties == NULL) {
        SetLastError(ERROR_PROC_NOT_FOUND);
        return FALSE;
    }

    return (*fnSHObjectProperties)(hwnd, dwType, szObject, szPage);
}




// Create a shell link
int Twapi_WriteShortcut (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR linkPath;
    LPCWSTR objPath;
    LPITEMIDLIST itemIds = NULL;
    LPCWSTR commandArgs;
    LPCWSTR desc;
    WORD    hotkey;
    LPCWSTR iconPath;
    int     iconIndex;
    LPCWSTR relativePath;
    int     showCommand;
    LPCWSTR workingDirectory;

    HRESULT hres; 
    IShellLinkW* psl = NULL; 
    IPersistFile* ppf = NULL;
 
    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(linkPath), GETNULLIFEMPTY(objPath),
                     GETVAR(itemIds, ObjToPIDL),
                     GETNULLIFEMPTY(commandArgs),
                     GETNULLIFEMPTY(desc),
                     GETWORD(hotkey),
                     GETNULLIFEMPTY(iconPath),
                     GETINT(iconIndex),
                     GETNULLIFEMPTY(relativePath),
                     GETINT(showCommand),
                     GETNULLIFEMPTY(workingDirectory),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (objPath == NULL && itemIds == NULL)
        return ERROR_INVALID_PARAMETER;

    // Get a pointer to the IShellLink interface. 
    hres = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, 
                            &IID_IShellLinkW, (LPVOID*)&psl); 
    if (FAILED(hres)) 
        return hres;

    if (objPath)
        hres = psl->lpVtbl->SetPath(psl,objPath); 
    if (FAILED(hres))
        goto vamoose;
    if (itemIds)
        hres = psl->lpVtbl->SetIDList(psl, itemIds);
    if (FAILED(hres))
        goto vamoose;
    if (commandArgs)
        hres = psl->lpVtbl->SetArguments(psl, commandArgs);
    if (FAILED(hres))
        goto vamoose;
    if (desc)
        hres = psl->lpVtbl->SetDescription(psl, desc); 
    if (FAILED(hres))
        goto vamoose;
    if (hotkey)
        hres = psl->lpVtbl->SetHotkey(psl, hotkey);
    if (FAILED(hres))
        goto vamoose;
    if (iconPath)
        hres = psl->lpVtbl->SetIconLocation(psl, iconPath, iconIndex);
    if (FAILED(hres))
        goto vamoose;
    if (relativePath)
        hres = psl->lpVtbl->SetRelativePath(psl, relativePath, 0);
    if (FAILED(hres))
        goto vamoose;
    if (showCommand >= 0)
        hres = psl->lpVtbl->SetShowCmd(psl, showCommand);
    if (FAILED(hres))
        goto vamoose;
    if (workingDirectory)
        hres = psl->lpVtbl->SetWorkingDirectory(psl, workingDirectory);
    if (FAILED(hres))
        goto vamoose;

    hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    if (FAILED(hres))
        goto vamoose;
 
    /* Save the link  */
    hres = ppf->lpVtbl->Save(ppf, linkPath, TRUE); 
    ppf->lpVtbl->Release(ppf); 
    
 vamoose:
    if (psl)
        psl->lpVtbl->Release(psl); 
    TwapiFreePIDL(itemIds);     /* OK if NULL */

    if (hres != S_OK) {
        Twapi_AppendSystemError(interp, hres);
        return TCL_ERROR;
    } else
        return TCL_OK;
}

/* Read a shortcut */
int Twapi_ReadShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR linkPath;
    int pathFlags;
    HWND hwnd;
    DWORD resolve_flags;

    HRESULT hres; 
    IShellLinkW *psl = NULL; 
    IPersistFile *ppf = NULL;
    Tcl_Obj *resultObj = NULL;
#if (INFOTIPSIZE > MAX_PATH)
    WCHAR buf[INFOTIPSIZE+1];
#else
    WCHAR buf[MAX_PATH+1];
#endif
    WORD  wordval;
    int   intval;
    LPITEMIDLIST pidl;
    int   retval = TCL_ERROR;
 
    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(linkPath), GETINT(pathFlags),
                     GETDWORD_PTR(hwnd), GETINT(resolve_flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* Get a pointer to the IShellLink interface. */
    hres = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, 
                            &IID_IShellLinkW, (LPVOID*)&psl); 
    if (FAILED(hres))
        goto fail;

    /* Load the resource through the IPersist interface */
    hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    if (FAILED(hres))
        goto fail;
    
    hres = ppf->lpVtbl->Load(ppf, linkPath, STGM_READ);
    if (FAILED(hres))
        goto fail;

    /* Resolve the link */
    hres = psl->lpVtbl->Resolve(psl, hwnd, resolve_flags);
#if 0    /* Ignore resolve errors */
    if (FAILED(hres))
        goto fail;
#endif

    resultObj = Tcl_NewListObj(0, NULL);

    /*
     * Get each field. Note that inability to get a field is not treated
     * as an error. We just go on and try to get the next one
     */
    hres = psl->lpVtbl->GetArguments(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-args"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromUnicode(buf));
    }

    hres = psl->lpVtbl->GetDescription(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-desc"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromUnicode(buf));
    }

    hres = psl->lpVtbl->GetHotkey(psl, &wordval);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-hotkey"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewIntObj(wordval));
    }

    hres = psl->lpVtbl->GetIconLocation(psl,
                                        buf, sizeof(buf)/sizeof(buf[0]),
                                        &intval);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-iconindex"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewIntObj(intval));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-iconpath"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromUnicode(buf));
    }

    hres = psl->lpVtbl->GetIDList(psl, &pidl);
    if (hres == NOERROR) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-idl"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromPIDL(pidl));
        CoTaskMemFree(pidl);
    }

    hres = psl->lpVtbl->GetPath(psl, buf, sizeof(buf)/sizeof(buf[0]),
                                NULL, pathFlags);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-path"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromUnicode(buf));
    }

    hres = psl->lpVtbl->GetShowCmd(psl, &intval);
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-showcmd"));
        Tcl_ListObjAppendElement(interp, resultObj, 
                                 Tcl_NewIntObj(intval));
    }

    hres = psl->lpVtbl->GetWorkingDirectory(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-workdir"));
        Tcl_ListObjAppendElement(interp, resultObj, ObjFromUnicode(buf));
    }
    
    Tcl_SetObjResult(interp, resultObj);
    retval = TCL_OK;

 vamoose:
    if (psl)
        psl->lpVtbl->Release(psl); 
    if (ppf)
        ppf->lpVtbl->Release(ppf); 

    return retval;

 fail:
    if (resultObj)
        Twapi_FreeNewTclObj(resultObj);
    resultObj = NULL;
    Twapi_AppendSystemError(interp, hres);
    goto vamoose;
}


// Create a url link
int Twapi_WriteUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath, LPCWSTR url, DWORD flags)
{ 
    HRESULT hres; 
    IUniformResourceLocatorW *psl = NULL; 
    IPersistFile* ppf = NULL;
 
    // TBD - can this be done at script level since we have IPersistFile
    // interface ?
    // Get a pointer to the IShellLink interface. 
    hres = CoCreateInstance(&CLSID_InternetShortcut, NULL,
                            CLSCTX_INPROC_SERVER, 
                            &IID_IUniformResourceLocatorW, (LPVOID*)&psl); 

    if (FAILED(hres)) {
        /* No interface and hence no interface specific error, just return as standard error */
        return Twapi_AppendSystemError(interp, hres);
    }

    hres = psl->lpVtbl->SetURL(psl, url, flags);
    if (FAILED(hres)) {
        TWAPI_STORE_COM_ERROR(interp, hres, psl, &IID_IUniformResourceLocatorW);
    } else {
        hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf);
        if (FAILED(hres)) {
            /* No-op - this is a standard error so we do not get try
               getting ISupportErrorInfo */
            Twapi_AppendSystemError(interp, hres);
        } else {
            /* Save the link  */
            hres = ppf->lpVtbl->Save(ppf, linkPath, TRUE); 
            if (FAILED(hres)) {
                TWAPI_STORE_COM_ERROR(interp, hres, ppf, &IID_IPersistFile);
            }
        }
    }

    if (ppf)
        ppf->lpVtbl->Release(ppf); 
    if (psl)
        psl->lpVtbl->Release(psl); 

    return FAILED(hres) ? TCL_ERROR : TCL_OK;
}

/* Read a URL shortcut */
int Twapi_ReadUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath)
{
    HRESULT hres; 
    IUniformResourceLocatorW *psl = NULL; 
    IPersistFile *ppf = NULL;
    LPWSTR url;
    int   retval = TCL_ERROR;
 
    /* Get a pointer to the IShellLink interface. */
    hres = CoCreateInstance(&CLSID_InternetShortcut, NULL,
                            CLSCTX_INPROC_SERVER, 
                            &IID_IUniformResourceLocatorW, (LPVOID*)&psl); 
    if (FAILED(hres))
        goto fail;

    /* Load the resource through the IPersist interface */
    hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    if (FAILED(hres))
        goto fail;
    
    hres = ppf->lpVtbl->Load(ppf, linkPath, STGM_READ);
    if (FAILED(hres))
        goto fail;

    hres = psl->lpVtbl->GetURL(psl, &url);

    if (FAILED(hres))
        goto fail;

    Tcl_SetObjResult(interp, ObjFromUnicode(url));
    CoTaskMemFree(url);

    retval = TCL_OK;

 vamoose:
    if (psl)
        psl->lpVtbl->Release(psl); 
    if (ppf)
        ppf->lpVtbl->Release(ppf); 

    return retval;

 fail:
    Twapi_AppendSystemError(interp, hres);
    goto vamoose;
}


/* Invoke a URL shortcut */
int Twapi_InvokeUrlShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR linkPath;
    LPCWSTR verb;
    DWORD flags;
    HWND hwnd;

    HRESULT hres; 
    IUniformResourceLocatorW *psl = NULL; 
    IPersistFile *ppf = NULL;
    URLINVOKECOMMANDINFOW urlcmd;
 
    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(linkPath), GETWSTR(verb), GETINT(flags),
                     GETHANDLE(hwnd),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* Get a pointer to the IShellLink interface. */
    hres = CoCreateInstance(&CLSID_InternetShortcut, NULL,
                            CLSCTX_INPROC_SERVER, 
                            &IID_IUniformResourceLocatorW, (LPVOID*)&psl); 
    if (SUCCEEDED(hres)) {
        hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf); 
    }
    if (FAILED(hres)) {
        /* This CoCreateInstance or QueryInterface error so we do not get
           use TWAPI_STORE_COM_ERROR */
        Twapi_AppendSystemError(interp, hres);
    } else {
        hres = ppf->lpVtbl->Load(ppf, linkPath, STGM_READ);
        if (FAILED(hres)) {
            TWAPI_STORE_COM_ERROR(interp, hres, ppf, &IID_IPersistFile);
        } else {
            urlcmd.dwcbSize = sizeof(urlcmd);
            urlcmd.dwFlags = flags;
            urlcmd.hwndParent = hwnd;
            urlcmd.pcszVerb = verb;

            hres = psl->lpVtbl->InvokeCommand(psl, &urlcmd);
            if (FAILED(hres)) {
                TWAPI_STORE_COM_ERROR(interp, hres, psl, &IID_IUniformResourceLocatorW);
            }
        }
    }

    if (ppf)
        ppf->lpVtbl->Release(ppf); 
    if (psl)
        psl->lpVtbl->Release(psl); 

    return FAILED(hres) ? TCL_ERROR : TCL_OK;
}




int Twapi_SHFileOperation (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SHFILEOPSTRUCTW sfop;
    Tcl_Obj *objs[2];
    int      tcl_status = TCL_ERROR;

    sfop.pFrom = NULL;          /* To track necessary deallocs */
    sfop.pTo   = NULL;
    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(sfop.hwnd), GETINT(sfop.wFunc),
                     GETVAR(sfop.pFrom, ObjToMultiSz),
                     GETVAR(sfop.pTo, ObjToMultiSz),
                     GETWORD(sfop.fFlags), GETWSTR(sfop.lpszProgressTitle),
                     ARGEND) != TCL_OK)
        goto vamoose;

    sfop.hNameMappings = NULL;

    if (SHFileOperationW(&sfop) != 0) {
        // Note GetLastError() is not set by the call
        Tcl_SetResult(interp, "SHFileOperation failed", TCL_STATIC);
        goto vamoose;
    }

    objs[0] = Tcl_NewBooleanObj(sfop.fAnyOperationsAborted);
    objs[1] = Tcl_NewListObj(0, NULL);

    if (sfop.hNameMappings) {
        int i;
        SHNAMEMAPPINGW *mapP = *(SHNAMEMAPPINGW **)(((char *)sfop.hNameMappings) + 4);
        for (i = 0; i < *(int *) (sfop.hNameMappings); ++i) {
            Tcl_ListObjAppendElement(interp, objs[1],
                                     ObjFromUnicodeN(
                                         mapP[i].pszOldPath,
                                         mapP[i].cchOldPath));
            Tcl_ListObjAppendElement(interp, objs[1],
                                     ObjFromUnicodeN(
                                         mapP[i].pszNewPath,
                                         mapP[i].cchNewPath));
        }

        SHFreeNameMappings(sfop.hNameMappings);
    }

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objs));
    tcl_status = TCL_OK;

vamoose:
    if (sfop.pFrom)
        TwapiFree((void*)sfop.pFrom);
    if (sfop.pTo)
        TwapiFree((void*)sfop.pTo);

    return tcl_status;
}

int Twapi_ShellExecuteEx(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPCWSTR lpClass;
    HKEY hkeyClass;
    DWORD dwHotKey;
    HANDLE hIconOrMonitor;

    SHELLEXECUTEINFOW sei;
    
    TwapiZeroMemory(&sei, sizeof(sei)); /* Also sets sei.lpIDList = NULL -
                                     Need to track if it needs freeing */
    sei.cbSize = sizeof(sei);

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(sei.fMask), GETHANDLE(sei.hwnd),
                     GETNULLIFEMPTY(sei.lpVerb),
                     GETNULLIFEMPTY(sei.lpFile),
                     GETNULLIFEMPTY(sei.lpParameters),
                     GETNULLIFEMPTY(sei.lpDirectory),
                     GETINT(sei.nShow),
                     GETVAR(sei.lpIDList, ObjToPIDL),
                     GETNULLIFEMPTY(lpClass),
                     GETHANDLE(hkeyClass),
                     GETINT(dwHotKey),
                     GETHANDLE(hIconOrMonitor),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (sei.fMask & SEE_MASK_CLASSNAME)
        sei.lpClass = lpClass;
    if (sei.fMask & SEE_MASK_CLASSKEY)
        sei.hkeyClass = hkeyClass;
    if (sei.fMask & SEE_MASK_HOTKEY)
        sei.dwHotKey = dwHotKey;
    if (sei.fMask & SEE_MASK_ICON)
        sei.hIcon = hIconOrMonitor;
    if (sei.fMask & SEE_MASK_HMONITOR)
        sei.hMonitor = hIconOrMonitor;

    if (ShellExecuteExW(&sei) == 0) {
        DWORD winerr = GetLastError();
        /* Note: double cast (int)(ULONG_PTR) is to prevent 64-bit warnings */
        Tcl_Obj *objP = Tcl_ObjPrintf("ShellExecute specific error: %d.", (int) (ULONG_PTR) sei.hInstApp);
        TwapiFreePIDL(sei.lpIDList);     /* OK if NULL */
        return Twapi_AppendSystemErrorEx(interp, winerr, objP);
    }

    TwapiFreePIDL(sei.lpIDList);     /* OK if NULL */

    /* Success, see if any fields to be returned */
    if (sei.fMask & SEE_MASK_NOCLOSEPROCESS) {
        Tcl_SetObjResult(interp, ObjFromHANDLE(sei.hProcess));
    }
    return TCL_OK;
}


int Twapi_SHChangeNotify(
    ClientData notused,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    LONG event_id;
    UINT flags;
    LPVOID dwItem1 = NULL;
    LPVOID dwItem2 = NULL;
    LPITEMIDLIST idl1P = NULL;
    LPITEMIDLIST idl2P = NULL;
    int status = TCL_ERROR;


    if (objc < 3) {
        goto wrong_nargs_error;
    }

    if (Tcl_GetLongFromObj(interp, objv[1], &event_id) != TCL_OK ||
        Tcl_GetIntFromObj(interp, objv[2], &flags) != TCL_OK) {
        goto vamoose;
    }
        
    switch (flags & SHCNF_TYPE) {
    case SHCNF_DWORD:
    case SHCNF_PATHW:
    case SHCNF_PRINTERW:
    case SHCNF_IDLIST:
        /* Valid but no special treatment */
        break;
    case SHCNF_PATHA:
        /* Always pass as Unicode */
        flags = (flags & ~SHCNF_TYPE) | SHCNF_PATHW;
        break;
    case SHCNF_PRINTERA:
        /* Always pass as Unicode */
        flags = (flags & ~SHCNF_TYPE) | SHCNF_PRINTERW;
        break;
        
    default:
        goto invalid_flags_error;
    }


    switch (event_id) {
    case SHCNE_ASSOCCHANGED:
        /* Both dwItem1 and dwItem2 must be NULL (already set) */
        if (! (flags & SHCNF_IDLIST)) {
            /* SDK says this should be set */
            goto invalid_flags_error;
        }
        break;

    case SHCNE_ATTRIBUTES:
    case SHCNE_CREATE:
    case SHCNE_DELETE:
    case SHCNE_DRIVEADD:
    case SHCNE_DRIVEADDGUI:
    case SHCNE_DRIVEREMOVED:
    case SHCNE_FREESPACE:
    case SHCNE_MEDIAINSERTED:
    case SHCNE_MEDIAREMOVED:
    case SHCNE_MKDIR:
    case SHCNE_RMDIR:
    case SHCNE_NETSHARE:
    case SHCNE_NETUNSHARE:
    case SHCNE_UPDATEDIR:
    case SHCNE_UPDATEITEM:
    case SHCNE_SERVERDISCONNECT:
        /* For the above, only dwItem1 used, dwItem2 remains 0 */

        if (objc < 4)
            goto wrong_nargs_error;

        switch (flags & SHCNF_TYPE) {
        case SHCNF_IDLIST:
            if (ObjToPIDL(interp, objv[3], &idl1P) != TCL_OK)
                goto vamoose;
            dwItem1 = idl1P;
            break;
        case SHCNF_PATHW:
            dwItem1 = Tcl_GetUnicode(objv[3]);
            break;
        default:
            goto invalid_flags_error;
        }

        break;

    case SHCNE_RENAMEITEM:
    case SHCNE_RENAMEFOLDER:
        /* Both dwItem1 and dwItem2 used */
        if (objc < 5)
            goto wrong_nargs_error;

        switch (flags & SHCNF_TYPE) {
        case SHCNF_IDLIST:
            if (ObjToPIDL(interp, objv[3], &idl1P) != TCL_OK ||
                ObjToPIDL(interp, objv[4], &idl2P) != TCL_OK)
                goto vamoose;
            dwItem1 = idl1P;
            dwItem2 = idl2P;
            break;
        case SHCNF_PATHW:
            dwItem1 = Tcl_GetUnicode(objv[3]);
            dwItem2 = Tcl_GetUnicode(objv[4]);
            break;
        default:
            goto invalid_flags_error;
        }

        break;

    case SHCNE_UPDATEIMAGE:
        /* dwItem1 not used, dwItem2 is a DWORD */
        if (objc < 5)
            goto wrong_nargs_error;
        if (Tcl_GetLongFromObj(interp, objv[4], (long *)&dwItem2) != TCL_OK)
            goto vamoose;
        break;
        
    case SHCNE_ALLEVENTS:
    case SHCNE_DISKEVENTS:
    case SHCNE_GLOBALEVENTS:
    case SHCNE_INTERRUPT:
        /*
         * SDK docs do not really say what parameters are valid and
         * how to interpret them. So for the below cases we make
         * a best guess based on the parameter and flags supplied by the
         * caller. If number of parameters is less than 4, then assume
         * dwItem1&2 are both unused etc.
         */
        if (objc >= 4) {
            switch (flags & SHCNF_TYPE) {
            case SHCNF_DWORD:
                if (Tcl_GetLongFromObj(interp, objv[3], (long *)&dwItem1) != TCL_OK)
                    goto vamoose;
                if (objc > 4 &&
                    (Tcl_GetLongFromObj(interp, objv[4],
                                        (long *)&dwItem2) != TCL_OK))
                    goto vamoose;
                break;
            case SHCNF_IDLIST:
                if (ObjToPIDL(interp, objv[3], &idl1P) != TCL_OK)
                    goto vamoose;
                dwItem1 = idl1P;
                if (objc > 4) {
                    if (ObjToPIDL(interp, objv[4], &idl2P) != TCL_OK)
                        goto vamoose;
                    dwItem2 = idl2P;
                }
                break;
            case SHCNF_PATHW:
                dwItem1 = Tcl_GetUnicode(objv[3]);
                if (objc > 4)
                    dwItem2 = Tcl_GetUnicode(objv[4]);
                break;
            default:
                goto invalid_flags_error;
            }
        }

        break;

    default:
        Tcl_SetResult(interp, "Unknown or unsupported SHChangeNotify event type", TCL_STATIC);
        goto vamoose;

    }

    /* Note SHChangeNotify has no error return */
    SHChangeNotify(event_id, flags, dwItem1, dwItem2);
    status =  TCL_OK;

vamoose:
    if (idl1P)
        TwapiFreePIDL(idl1P);
    if (idl2P)
        TwapiFreePIDL(idl2P);

    return status;

invalid_flags_error:
    Tcl_SetResult(interp,
                  "Unknown or unsupported SHChangeNotify flags type",
                  TCL_STATIC);
    goto vamoose;

wrong_nargs_error:
    Tcl_WrongNumArgs(interp, 1, objv, "wEventID uFlags dwItem1 ?dwItem2?");
    goto vamoose;
}

static int Twapi_ShellCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    HWND   hwnd;
    LPWSTR s, s2;
    DWORD dw, dw2;
    union {
        NOTIFYICONDATAW *niP;
        WCHAR buf[MAX_PATH+1];
    } u;
    HANDLE h;
    TwapiResult result;
    LPITEMIDLIST idlP;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        return Twapi_WriteShortcut(interp, objc-2, objv+2);
    case 2:
        return Twapi_ReadShortcut(interp, objc-2, objv+2);
    case 3:
        return Twapi_InvokeUrlShortcut(interp, objc-2, objv+2);
    case 4: // SHInvokePrinterCommand
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(hwnd), GETINT(dw),
                         GETWSTR(s), GETWSTR(s2), GETBOOL(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SHInvokePrinterCommandW(hwnd, dw, s, s2, dw2);
        break;
    case 5:
        return Twapi_GetShellVersion(interp);
    case 6: // SHGetFolderPath - TBD Tcl wrapper
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHWND(hwnd), GETINT(dw),
                         GETHANDLE(h), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        dw = Twapi_SHGetFolderPath(hwnd, dw, h, dw2, u.buf);
        if (dw == 0) {
            result.type = TRT_UNICODE;
            result.value.unicode.str = u.buf;
            result.value.unicode.len = -1;
        } else {
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = dw;
        }
        break;
    case 7:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHWND(hwnd), GETINT(dw),
                         GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (SHGetSpecialFolderPathW(hwnd, u.buf, dw, dw2)) {
            result.type = TRT_UNICODE;
            result.value.unicode.str = u.buf;
            result.value.unicode.len = -1;
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 8: // SHGetPathFromIDList
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETVAR(idlP, ObjToPIDL),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (SHGetPathFromIDListW(idlP, u.buf)) {
            result.type = TRT_UNICODE;
            result.value.unicode.str = u.buf;
            result.value.unicode.len = -1;
        } else {
            /* Need to get error before we call the pidl free */
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = GetLastError();
        }
        TwapiFreePIDL(idlP); /* OK if NULL */
        break;
    case 9:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLET(hwnd, HWND), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        dw = SHGetSpecialFolderLocation(hwnd, dw, &result.value.pidl);
        if (dw == 0)
            result.type = TRT_PIDL;
        else {
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = dw;
        }
        break;

    case 10: // Shell_NotifyIcon
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETINT(dw), GETBIN(u.niP, dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (dw2 != sizeof(NOTIFYICONDATAW) || dw2 != u.niP->cbSize) {
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Inconsistent size of NOTIFYICONDATAW structure.");
        }
        result.type = TRT_EMPTY;
        if (Shell_NotifyIconW(dw, u.niP) == FALSE) {
            result.type = TRT_GETLASTERROR;
        }
        break;
    case 11:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        return Twapi_ReadUrlShortcut(interp, Tcl_GetUnicode(objv[2]));
    case 12:
        return Twapi_SHFileOperation(interp, objc-2, objv+2);
    case 13:
        return Twapi_ShellExecuteEx(interp, objc-2, objv+2);
    case 14:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLET(hwnd, HWND), GETINT(dw),
                         GETWSTR(s), GETNULLIFEMPTY(s2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = Twapi_SHObjectProperties(hwnd, dw, s, s2);
        break;
    case 15:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETWSTR(s2), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_WriteUrlShortcut(interp, s, s2, dw);
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiShellInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::ShellCall", Twapi_ShellCallObjCmd, ticP, NULL);

    /* TBD - is there a tcl wrapper for SHChangeNotify ? */
    /* This is a separate command as opposed to using the standard dispatch
       for historical reasons. Should probably change it some time */
    Tcl_CreateObjCommand(interp, "twapi::SHChangeNotify", Twapi_SHChangeNotify,
                         ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::ShellCall", # code_); \
    } while (0);

    CALL_(Twapi_WriteShortcut, 1);
    CALL_(Twapi_ReadShortcut, 2);
    CALL_(Twapi_InvokeUrlShortcut, 3);
    CALL_(SHInvokePrinterCommand, 4); // TBD - Tcl wrapper
    CALL_(Twapi_GetShellVersion, 5);
    CALL_(SHGetFolderPath, 6);
    CALL_(SHGetSpecialFolderPath, 7);
    CALL_(SHGetPathFromIDList, 8); // TBD - Tcl wrapper
    CALL_(SHGetSpecialFolderLocation, 9);
    CALL_(Shell_NotifyIcon, 10);
    CALL_(Twapi_ReadUrlShortcut, 11);
    CALL_(Twapi_SHFileOperation, 12); // TBD - some more wrappers
    CALL_(Twapi_ShellExecuteEx, 13);
    CALL_(SHObjectProperties, 14);
    CALL_(Twapi_WriteUrlShortcut, 15);


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
int Twapi_shell_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiShellInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

