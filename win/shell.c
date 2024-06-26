/*
 * Copyright (c) 2004-2024 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* TBD - move theme functions to UI module ? */
/* TBD - SHLoadIndirectString */

#include "twapi.h"
#include <initguid.h> /* GUIDs in all included files below this will be instantiated */

/* Note: some versions of mingw define IID_IShellLinkDataList and some don't.
   Unlike MS VC++ which does not complain if duplicate definitions match,
   gcc is not happy and I cannot get gcc's selectany attribute for
   ignoring duplicate definitions to work. So just we use our own variable
   with content that is identical to the original.
*/
DEFINE_GUID(TWAPI_IID_IShellLinkDataList,     0x45e2b4ae, 0xb1c3, 0x11d0, 0xb9, 0x2f, 0x0, 0xa0, 0xc9, 0x3, 0x12, 0xe1);

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

#ifndef MODULENAME
#define MODULENAME "twapi_shell"
#endif

static HRESULT Twapi_SHGetFolderPath(HWND hwndOwner, int nFolder, HANDLE hToken,
                          DWORD flags, WCHAR *pathbuf);
static BOOL Twapi_SHObjectProperties(HWND hwnd, DWORD dwType,
                              LPCWSTR szObject, LPCWSTR szPage);

int Twapi_GetShellVersion(Tcl_Interp *interp);
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
    ObjSetResult(interp,
                     Tcl_ObjPrintf("%lu.%lu.%lu",
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
static HRESULT Twapi_SHGetFolderPath(
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
static BOOL Twapi_SHObjectProperties(
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
static TCL_RESULT Twapi_WriteShortcutObjCmd(
    ClientData clientdata,
    Tcl_Interp *interp,
    int  objc,
    Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = clientdata;
    LPWSTR linkPath;
    LPWSTR objPath;
    LPITEMIDLIST itemIds;
    LPWSTR commandArgs;
    LPWSTR desc;
    WORD    hotkey;
    LPWSTR iconPath;
    int     iconIndex;
    LPWSTR relativePath;
    int     showCommand;
    LPWSTR workingDirectory;
    DWORD runas;

    HRESULT hres;
    IShellLinkW* psl;
    IPersistFile* ppf;
    IShellLinkDataList* psldl;
    MemLifoMarkHandle mark;
    TCL_RESULT res;

    psl = NULL;
    psldl = NULL;
    itemIds = NULL;
    ppf = NULL;
    mark = MemLifoPushMark(ticP->memlifoP);

    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                         GETWSTR(linkPath), GETEMPTYASNULL(objPath),
                         GETVAR(itemIds, ObjToPIDL),
                         GETEMPTYASNULL(commandArgs),
                         GETEMPTYASNULL(desc),
                         GETWORD(hotkey),
                         GETEMPTYASNULL(iconPath),
                         GETINT(iconIndex),
                         GETEMPTYASNULL(relativePath),
                         GETINT(showCommand),
                         GETEMPTYASNULL(workingDirectory),
                         ARGUSEDEFAULT,
                         GETDWORD(runas),
                         ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    if (objPath == NULL && itemIds == NULL) {
        hres = ERROR_INVALID_PARAMETER;
        goto hres_vamoose;
    }

    // Get a pointer to the IShellLink interface.
    hres = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
                            &IID_IShellLinkW, (LPVOID*)&psl);
    if (FAILED(hres))
        goto hres_vamoose;

    if (objPath)
        hres = psl->lpVtbl->SetPath(psl,objPath);
    if (FAILED(hres))
        goto hres_vamoose;
    if (itemIds)
        hres = psl->lpVtbl->SetIDList(psl, itemIds);
    if (FAILED(hres))
        goto hres_vamoose;
    if (commandArgs)
        hres = psl->lpVtbl->SetArguments(psl, commandArgs);
    if (FAILED(hres))
        goto hres_vamoose;
    if (desc)
        hres = psl->lpVtbl->SetDescription(psl, desc);
    if (FAILED(hres))
        goto hres_vamoose;
    if (hotkey)
        hres = psl->lpVtbl->SetHotkey(psl, hotkey);
    if (FAILED(hres))
        goto hres_vamoose;
    if (iconPath)
        hres = psl->lpVtbl->SetIconLocation(psl, iconPath, iconIndex);
    if (FAILED(hres))
        goto hres_vamoose;
    if (relativePath)
        hres = psl->lpVtbl->SetRelativePath(psl, relativePath, 0);
    if (FAILED(hres))
        goto hres_vamoose;
    if (showCommand >= 0)
        hres = psl->lpVtbl->SetShowCmd(psl, showCommand);
    if (FAILED(hres))
        goto hres_vamoose;
    if (workingDirectory)
        hres = psl->lpVtbl->SetWorkingDirectory(psl, workingDirectory);
    if (FAILED(hres))
        goto hres_vamoose;

    if (runas) {
        hres = psl->lpVtbl->QueryInterface(psl, &TWAPI_IID_IShellLinkDataList,
                                           (LPVOID*)&psldl);
        if (FAILED(hres))
            goto hres_vamoose;

        hres = psldl->lpVtbl->GetFlags(psldl, &runas);
        if (FAILED(hres))
            goto hres_vamoose;

        hres = psldl->lpVtbl->SetFlags(psldl, runas | SLDF_RUNAS_USER);
        if (FAILED(hres))
            goto hres_vamoose;
    }

    hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile, (LPVOID*)&ppf);
    if (FAILED(hres))
        goto hres_vamoose;

    /* Save the link  */
    hres = ppf->lpVtbl->Save(ppf, linkPath, TRUE);
    ppf->lpVtbl->Release(ppf);

 hres_vamoose:
    if (psl)
        psl->lpVtbl->Release(psl);
    if (psldl)
        psldl->lpVtbl->Release(psldl);
    TwapiFreePIDL(itemIds);     /* OK if NULL */

    if (hres != S_OK) {
        Twapi_AppendSystemError(interp, hres);
        res = TCL_ERROR;
    } else
        res = TCL_OK;

vamoose:
    MemLifoPopMark(mark);
    return res;
}

/* Read a shortcut */
int Twapi_ReadShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
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
    IShellLinkDataList* psldl = NULL;
    DWORD runas;
    Tcl_Obj *linkObj;

    if (TwapiGetArgs(interp, objc, objv,
                     GETOBJ(linkObj), GETINT(pathFlags),
                     GETHWND(hwnd), GETDWORD(resolve_flags),
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

    hres = ppf->lpVtbl->Load(ppf, ObjToWinChars(linkObj), STGM_READ);
    if (FAILED(hres))
        goto fail;

    /* Resolve the link */
    hres = psl->lpVtbl->Resolve(psl, hwnd, resolve_flags);
#if 0    /* Ignore resolve errors */
    if (FAILED(hres))
        goto fail;
#endif

    resultObj = ObjEmptyList();

    /*
     * Get each field. Note that inability to get a field is not treated
     * as an error. We just go on and try to get the next one
     */
    hres = psl->lpVtbl->GetArguments(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-args"));
        ObjAppendElement(interp, resultObj, ObjFromWinChars(buf));
    }

    hres = psl->lpVtbl->GetDescription(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-desc"));
        ObjAppendElement(interp, resultObj, ObjFromWinChars(buf));
    }

    hres = psl->lpVtbl->GetHotkey(psl, &wordval);
    if (SUCCEEDED(hres)) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-hotkey"));
        ObjAppendElement(interp, resultObj,
                                 ObjFromLong(wordval));
    }

    hres = psl->lpVtbl->GetIconLocation(psl,
                                        buf, sizeof(buf)/sizeof(buf[0]),
                                        &intval);
    if (SUCCEEDED(hres)) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-iconindex"));
        ObjAppendElement(interp, resultObj,
                                 ObjFromLong(intval));
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-iconpath"));
        ObjAppendElement(interp, resultObj, ObjFromWinChars(buf));
    }

    hres = psl->lpVtbl->GetIDList(psl, &pidl);
    if (hres == NOERROR) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-idl"));
        ObjAppendElement(interp, resultObj, ObjFromPIDL(pidl));
        CoTaskMemFree(pidl);
    }

    hres = psl->lpVtbl->GetPath(psl, buf, sizeof(buf)/sizeof(buf[0]),
                                NULL, pathFlags);
    if (SUCCEEDED(hres)) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-path"));
        ObjAppendElement(interp, resultObj, ObjFromWinChars(buf));
    }

    hres = psl->lpVtbl->GetShowCmd(psl, &intval);
    if (SUCCEEDED(hres)) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-showcmd"));
        ObjAppendElement(interp, resultObj,
                                 ObjFromLong(intval));
    }

    hres = psl->lpVtbl->GetWorkingDirectory(psl, buf, sizeof(buf)/sizeof(buf[0]));
    if (SUCCEEDED(hres)) {
        ObjAppendElement(interp, resultObj,
                                 STRING_LITERAL_OBJ("-workdir"));
        ObjAppendElement(interp, resultObj, ObjFromWinChars(buf));
    }

    hres = psl->lpVtbl->QueryInterface(psl, &TWAPI_IID_IShellLinkDataList,
                                       (LPVOID*)&psldl);
    if (SUCCEEDED(hres)) {
        hres = psldl->lpVtbl->GetFlags(psldl, &runas);
        if (SUCCEEDED(hres)) {
            ObjAppendElement(interp, resultObj,
                             STRING_LITERAL_OBJ("-runas"));
            ObjAppendElement(interp, resultObj, ObjFromInt(runas & SLDF_RUNAS_USER ? 1 : 0));
        }
    }

    ObjSetResult(interp, resultObj);
    retval = TCL_OK;

 vamoose:
    if (psl)
        psl->lpVtbl->Release(psl);
    if (psldl)
        psldl->lpVtbl->Release(psldl);
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

    ObjSetResult(interp, ObjFromWinChars(url));
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
    DWORD flags;
    HWND hwnd;

    HRESULT hres;
    IUniformResourceLocatorW *psl = NULL;
    IPersistFile *ppf = NULL;
    URLINVOKECOMMANDINFOW urlcmd;
    Tcl_Obj *linkObj, *verbObj;

    if (TwapiGetArgs(interp, objc, objv,
                     GETOBJ(linkObj), GETOBJ(verbObj), GETDWORD(flags),
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
        hres = ppf->lpVtbl->Load(ppf, ObjToWinChars(linkObj), STGM_READ);
        if (FAILED(hres)) {
            TWAPI_STORE_COM_ERROR(interp, hres, ppf, &IID_IPersistFile);
        } else {
            urlcmd.dwcbSize = sizeof(urlcmd);
            urlcmd.dwFlags = flags;
            urlcmd.hwndParent = hwnd;
            urlcmd.pcszVerb = ObjToWinChars(verbObj);

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
    Tcl_Obj  *titleObj, *fromObj, *toObj;
    SWSMark  mark = NULL;
    TCL_RESULT res;
    FILEOP_FLAGS flags;

    mark = SWSPushMark();
    res = TCL_ERROR;
    if (TwapiGetArgs(interp, objc, objv,
                       GETHANDLE(sfop.hwnd), GETUINT(sfop.wFunc),
                       GETOBJ(fromObj), GETOBJ(toObj),
                       GETWORD(sfop.fFlags), GETOBJ(titleObj),
                       ARGEND) != TCL_OK
        ||
        ObjToMultiSzSWS(interp, fromObj, &sfop.pFrom, NULL) != TCL_OK
        ||
        ObjToMultiSzSWS(interp, toObj, &sfop.pTo, NULL) != TCL_OK
        )
        goto vamoose;

    sfop.lpszProgressTitle = ObjToWinChars(titleObj);
    sfop.hNameMappings     = NULL;
    flags                  = sfop.fFlags;

    if (SHFileOperationW(&sfop) != 0) {
        // Note GetLastError() is not set by the call
        ObjSetStaticResult(interp, "SHFileOperation failed");
        goto vamoose;
    }

    objs[0] = ObjFromBoolean(sfop.fAnyOperationsAborted);
    objs[1] = ObjEmptyList();

    if ((flags & FOF_WANTMAPPINGHANDLE) && sfop.hNameMappings) {
        UINT i;
        struct handlemappings {
            UINT uNumberOfMappings;
            LPSHNAMEMAPPINGW nameMapP;
        }; /* See https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shfileopstructa */
        struct handlemappings *mapP  = sfop.hNameMappings;
        UINT                   n     = mapP->uNumberOfMappings;
        LPSHNAMEMAPPINGW       elemP = mapP->nameMapP;
        for (i = 0; i < n; ++i, ++elemP) {
            ObjAppendElement(interp, objs[1],
                                     ObjFromWinCharsN(
                                         elemP->pszOldPath,
                                         elemP->cchOldPath));
            ObjAppendElement(interp, objs[1],
                                     ObjFromWinCharsN(
                                         elemP->pszNewPath,
                                         elemP->cchNewPath));
        }

        SHFreeNameMappings(sfop.hNameMappings);
    }

    ObjSetResult(interp, ObjNewList(2, objs));
    res = TCL_OK;

vamoose:
    if (mark)
        SWSPopMark(mark);

    return res;
}

static TCL_RESULT Twapi_ShellExecuteExObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = clientdata;
    LPCWSTR lpClass;
    HKEY hkeyClass;
    DWORD dwHotKey;
    HANDLE hIconOrMonitor;
    TCL_RESULT res;
    SHELLEXECUTEINFOW sei;
    MemLifoMarkHandle mark;

    mark = MemLifoPushMark(ticP->memlifoP);

    TwapiZeroMemory(&sei, sizeof(sei)); /* Also sets sei.lpIDList = NULL -
                                     Need to track if it needs freeing */
    sei.cbSize = sizeof(sei);

    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                         GETDWORD(sei.fMask), GETHANDLE(sei.hwnd),
                         GETEMPTYASNULL(sei.lpVerb),
                         GETEMPTYASNULL(sei.lpFile),
                         GETEMPTYASNULL(sei.lpParameters),
                         GETEMPTYASNULL(sei.lpDirectory),
                         GETINT(sei.nShow),
                         GETVAR(sei.lpIDList, ObjToPIDL),
                         GETEMPTYASNULL(lpClass),
                         GETHANDLE(hkeyClass),
                         GETDWORD(dwHotKey),
                         GETHANDLE(hIconOrMonitor),
                         ARGEND);
    if (res == TCL_OK) {
        if (sei.fMask & SEE_MASK_CLASSNAME)
            sei.lpClass = lpClass;
        if (sei.fMask & SEE_MASK_CLASSKEY)
            sei.hkeyClass = hkeyClass;
        if (sei.fMask & SEE_MASK_HOTKEY)
            sei.dwHotKey = dwHotKey;
#ifndef SEE_MASK_ICON
# define SEE_MASK_ICON 0x00000010
#endif
        if (sei.fMask & SEE_MASK_ICON)
            sei.hIcon = hIconOrMonitor;
        if (sei.fMask & SEE_MASK_HMONITOR)
            sei.hMonitor = hIconOrMonitor;

        if (ShellExecuteExW(&sei)) {
            if (sei.fMask & SEE_MASK_NOCLOSEPROCESS) {
                ObjSetResult(interp, ObjFromHANDLE(sei.hProcess));
            }
        } else {
            /* Note: error code in se.hInstApp is ignored as per
               http://blogs.msdn.com/b/oldnewthing/archive/2012/10/18/10360604.aspx
            */
            res = TwapiReturnSystemError(interp);
        }
        TwapiFreePIDL(sei.lpIDList);     /* OK if NULL */
    }

    MemLifoPopMark(mark);
    return res;
}


int Twapi_SHChangeNotify(
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

    if (TwapiGetArgs(interp, objc, objv, GETLONG(event_id), GETUINT(flags),
                     ARGTERM) != TCL_OK)
        return TCL_ERROR;

    switch (flags & SHCNF_TYPE) {
    case SHCNF_DWORD:
    case SHCNF_PATHW:
    case SHCNF_PRINTERW:
    case SHCNF_IDLIST:
        /* Valid but no special treatment */
        break;
    case SHCNF_PATHA:
        /* Always pass as WCHAR */
        flags = (flags & ~SHCNF_TYPE) | SHCNF_PATHW;
        break;
    case SHCNF_PRINTERA:
        /* Always pass as WCHAR */
        flags = (flags & ~SHCNF_TYPE) | SHCNF_PRINTERW;
        break;

    default:
        goto invalid_flags_error;
    }


    switch (event_id) {
    case SHCNE_ASSOCCHANGED:
        /* Both dwItem1 and dwItem2 must be NULL (already set) */
        if ((flags & SHCNF_TYPE) != SHCNF_IDLIST) {
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

        if (objc < 3)
            goto wrong_nargs_error;

        switch (flags & SHCNF_TYPE) {
        case SHCNF_IDLIST:
            if (ObjToPIDL(interp, objv[2], &idl1P) != TCL_OK)
                goto vamoose;
            dwItem1 = idl1P;
            break;
        case SHCNF_PATHW:
            dwItem1 = ObjToWinChars(objv[2]);
            break;
        default:
            goto invalid_flags_error;
        }

        break;

    case SHCNE_RENAMEITEM:
    case SHCNE_RENAMEFOLDER:
        /* Both dwItem1 and dwItem2 used */
        if (objc < 4)
            goto wrong_nargs_error;

        switch (flags & SHCNF_TYPE) {
        case SHCNF_IDLIST:
            if (ObjToPIDL(interp, objv[2], &idl1P) != TCL_OK ||
                ObjToPIDL(interp, objv[3], &idl2P) != TCL_OK)
                goto vamoose;
            dwItem1 = idl1P;
            dwItem2 = idl2P;
            break;
        case SHCNF_PATHW:
            dwItem1 = ObjToWinChars(objv[2]);
            dwItem2 = ObjToWinChars(objv[3]);
            break;
        default:
            goto invalid_flags_error;
        }

        break;

    case SHCNE_UPDATEIMAGE:
        /* dwItem1 not used, dwItem2 is a DWORD */
        if (objc < 4)
            goto wrong_nargs_error;
        if (ObjToLong(interp, objv[3], (long *)&dwItem2) != TCL_OK)
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
        if (objc >= 3) {
            switch (flags & SHCNF_TYPE) {
            case SHCNF_DWORD:
                if (ObjToLong(interp, objv[2], (long *)&dwItem1) != TCL_OK)
                    goto vamoose;
                if (objc > 3 &&
                    (ObjToLong(interp, objv[3],
                                        (long *)&dwItem2) != TCL_OK))
                    goto vamoose;
                break;
            case SHCNF_IDLIST:
                if (ObjToPIDL(interp, objv[2], &idl1P) != TCL_OK)
                    goto vamoose;
                dwItem1 = idl1P;
                if (objc > 3) {
                    if (ObjToPIDL(interp, objv[3], &idl2P) != TCL_OK)
                        goto vamoose;
                    dwItem2 = idl2P;
                }
                break;
            case SHCNF_PATHW:
                dwItem1 = ObjToWinChars(objv[2]);
                if (objc > 3)
                    dwItem2 = ObjToWinChars(objv[3]);
                break;
            default:
                goto invalid_flags_error;
            }
        }

        break;

    default:
        ObjSetStaticResult(interp, "Unknown SHChangeNotify event type");
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
    ObjSetStaticResult(interp, "SHChangeNotify flags unknown or invalid for event type.");
    goto vamoose;

wrong_nargs_error:
    Tcl_WrongNumArgs(interp, 1, objv, "wEventID uFlags dwItem1 ?dwItem2?");
    goto vamoose;
}

static int Twapi_ShellCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HWND   hwnd;
    DWORD dw, dw2;
    BOOL  bval;
    union {
        NOTIFYICONDATAW *niP;
        WCHAR buf[MAX_PATH+1];
    } u;
    HANDLE h;
    TwapiResult result;
    LPITEMIDLIST idlP;
    Tcl_Obj *sObj, *s2Obj;
    int func = PtrToInt(clientdata);

    --objc;
    ++objv;
    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 2:
        return Twapi_ReadShortcut(interp, objc, objv);
    case 3:
        return Twapi_InvokeUrlShortcut(interp, objc, objv);
    case 4: // SHInvokePrinterCommand
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLE(hwnd), GETDWORD(dw),
                         GETOBJ(sObj), GETOBJ(s2Obj), GETBOOL(bval),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SHInvokePrinterCommandW(
            hwnd, dw, ObjToWinChars(sObj), ObjToWinChars(s2Obj), bval);
        break;
    case 5:
        return Twapi_GetShellVersion(interp);
    case 6: // SHGetFolderPath - TBD Tcl wrapper
        if (TwapiGetArgs(interp, objc, objv,
                         GETHWND(hwnd), GETDWORD(dw),
                         GETHANDLE(h), GETDWORD(dw2),
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
        if (TwapiGetArgs(interp, objc, objv,
                         GETHWND(hwnd), GETDWORD(dw),
                         GETDWORD(dw2),
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
        if (TwapiGetArgs(interp, objc, objv,
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
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hwnd, HWND), GETDWORD(dw),
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
        CHECK_NARGS(interp, objc, 2);
        CHECK_DWORD_OBJ(interp, dw, objv[0]);
        CHECK_RESULT(ObjToByteArrayDW(interp, objv[1], &dw2, (unsigned char **)&u.niP));
        if (dw2 != u.niP->cbSize) {
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Inconsistent size of NOTIFYICONDATAW structure.");
        }
        result.type = TRT_EMPTY;
        if (Shell_NotifyIconW(dw, u.niP) == FALSE) {
            result.type = TRT_GETLASTERROR;
        }
        break;
    case 11:
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        return Twapi_ReadUrlShortcut(interp, ObjToWinChars(objv[0]));
    case 12:
        return Twapi_SHFileOperation(interp, objc, objv);
    case 13:
        return Twapi_SHChangeNotify(interp, objc, objv);
    case 14:
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(hwnd, HWND), GETDWORD(dw),
                         GETOBJ(sObj), GETOBJ(s2Obj),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival =
            Twapi_SHObjectProperties(hwnd, dw,
                                     ObjToWinChars(sObj),
                                     ObjToLPWSTR_NULL_IF_EMPTY(s2Obj));
        break;
    case 15:
        if (TwapiGetArgs(interp, objc, objv,
                         GETOBJ(sObj), GETOBJ(s2Obj), GETDWORD(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_WriteUrlShortcut(interp,
                                      ObjToWinChars(sObj),
                                      ObjToWinChars(s2Obj), dw);
    }

    return TwapiSetResult(interp, &result);
}

static int TwapiShellInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s ShellDispatch[] = {
        DEFINE_FNCODE_CMD(Twapi_ReadShortcut, 2),
        DEFINE_FNCODE_CMD(Twapi_InvokeUrlShortcut, 3),
        DEFINE_FNCODE_CMD(SHInvokePrinterCommand, 4), // Deprecated in lieu of ShellExecute
        DEFINE_FNCODE_CMD(Twapi_GetShellVersion, 5),
        DEFINE_FNCODE_CMD(SHGetFolderPath, 6),
        DEFINE_FNCODE_CMD(SHGetSpecialFolderPath, 7),
        DEFINE_FNCODE_CMD(SHGetPathFromIDList, 8), // TBD - Tcl
        DEFINE_FNCODE_CMD(SHGetSpecialFolderLocation, 9),
        DEFINE_FNCODE_CMD(Shell_NotifyIcon, 10),
        DEFINE_FNCODE_CMD(Twapi_ReadUrlShortcut, 11),
        DEFINE_FNCODE_CMD(Twapi_SHFileOperation, 12), // TBD - some more wrapper
        DEFINE_FNCODE_CMD(SHChangeNotify, 13),
        DEFINE_FNCODE_CMD(SHObjectProperties, 14),
        DEFINE_FNCODE_CMD(Twapi_WriteUrlShortcut, 15),

    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(ShellDispatch), ShellDispatch, Twapi_ShellCallObjCmd);
    Tcl_CreateObjCommand(interp, "twapi::Twapi_WriteShortcut", Twapi_WriteShortcutObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::Twapi_ShellExecuteEx", Twapi_ShellExecuteExObjCmd, ticP, NULL);

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

