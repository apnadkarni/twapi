/* 
 * Copyright (c) 2003-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_storage.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

/* File and disk related */

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
    TwapiSetObjResult(interp, ObjFromLong(file_type));
    return TCL_OK;
}

int TwapiFirstVolume(
    Tcl_Interp *interp,
    LPCWSTR volpath            /* If NULL, Volume, else MountPoint */
)
{
    HANDLE h;
    WCHAR   buf[MAX_PATH+1];
    Tcl_Obj *objv[2];

    if (volpath)
        h = FindFirstVolumeMountPointW(volpath, buf, sizeof(buf)/sizeof(buf[0]));
    else
        h = FindFirstVolumeW(buf, sizeof(buf)/sizeof(buf[0]));

    if (h == INVALID_HANDLE_VALUE)
        return TwapiReturnSystemError(interp);

    objv[0] = ObjFromHANDLE(h);
    objv[1] = ObjFromUnicode(buf);

    TwapiSetObjResult(interp, ObjNewList(2, objv));
    return TCL_OK;
}

int TwapiNextVolume(Tcl_Interp *interp, int treat_as_mountpoint, HANDLE hFindVolume)
{
    BOOL    found;
    WCHAR   buf[MAX_PATH+1];

    found = (treat_as_mountpoint ?
             FindNextVolumeMountPointW : FindNextVolumeW)
        (hFindVolume, buf, sizeof(buf)/sizeof(buf[0]));

    if (found) {
        Tcl_Obj *objv[2];
        objv[0] = ObjFromLong(1);
        objv[1] = ObjFromUnicode(buf);
        TwapiSetObjResult(interp, ObjNewList(2, objv));
        return TCL_OK;
    } else {
        DWORD lasterr = GetLastError();
        buf[0] = 0;
        if (lasterr == ERROR_NO_MORE_FILES) {
            /* Not an error, signal no more volumes */
            TwapiSetObjResult(interp, ObjFromLong(0));
            return TCL_OK;
        }
        else
            return Twapi_AppendSystemError(interp, lasterr);
    }
}

int Twapi_GetVolumeInformation(Tcl_Interp *interp, LPCWSTR path)
{
    WCHAR volname[256], fsname[256];
    DWORD serial_no, max_component_len, sysflags;
    Tcl_Obj *objv[5];

    if (GetVolumeInformationW(
            path,
            volname, sizeof(volname)/sizeof(volname[0]),
            &serial_no,
            &max_component_len,
            &sysflags,
            fsname, sizeof(fsname)/sizeof(fsname[0])) == 0) {
        return TwapiReturnSystemError(interp);
    }

    objv[0] = ObjFromUnicode(volname);
    objv[1] = ObjFromLong(serial_no);
    objv[2] = ObjFromLong(max_component_len);
    objv[3] = ObjFromLong(sysflags);
    objv[4] = ObjFromUnicode(fsname);
    TwapiSetObjResult(interp, ObjNewList(5, objv));

    return TCL_OK;
}


int Twapi_QueryDosDevice(Tcl_Interp *interp, LPCWSTR lpDeviceName)
{
    WCHAR *pathP;
    DWORD  path_cch;
    DWORD result;

    path_cch = MAX_PATH;
    pathP = TwapiAlloc(sizeof(WCHAR)*path_cch); // TBD - instrument
    while (1) {
        // TBD - Tcl interface for when lpDeviceName is NULL ?

        result = QueryDosDeviceW(lpDeviceName, pathP, path_cch);
        if (result > 0) {
            /* On Win2K and NT 4, result is truncated without a terminating
               null if buffer is too small. In this case keep going */
            if (pathP[result-1] == 0 && pathP[result-2] == 0) {
                break;          /* Fine, done */
            }
            /* Else we will allocate more memory and retry */
        } else if (result > path_cch) {
            /* Should never happen, but ... */
            Tcl_Panic("QueryDosDeviceW overran buffer.");
        } else if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) {
            break;
        }

        path_cch = 2 * path_cch; /* Double buffer size */
        if (path_cch >= 128000) {
            /* Buffer is as big as we want it to be */
            result = 0;
            SetLastError(ERROR_INSUFFICIENT_BUFFER);
            break;
        }
        TwapiFree(pathP);
        pathP = TwapiAlloc(sizeof(WCHAR)*path_cch);
    }

    if (result) {
        TwapiSetObjResult(interp, ObjFromMultiSz(pathP, result));
        result = TCL_OK;
    } else {
        result = TwapiReturnSystemError(interp);
    }

    TwapiFree(pathP);
    return result;
}

int Twapi_GetDiskFreeSpaceEx(Tcl_Interp *interp, LPCWSTR dir)
{
    Tcl_Obj *objv[3];
    ULARGE_INTEGER free_avail;
    ULARGE_INTEGER total_bytes;
    ULARGE_INTEGER free_total;

    if (! GetDiskFreeSpaceExW(dir, &free_avail, &total_bytes, &free_total)) {
        return TwapiReturnSystemError(interp);
    }

    objv[0] = ObjFromWideInt(free_avail.QuadPart);
    objv[1] = ObjFromWideInt(total_bytes.QuadPart);
    objv[2] = ObjFromWideInt(free_total.QuadPart);
    TwapiSetObjResult(interp, ObjNewList(3, objv));
    return TCL_OK;
}

static int Twapi_StorageCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR s, s2;
    DWORD dw;
    HANDLE h;
    TwapiResult result;
    WCHAR buf[MAX_PATH+1];
    LARGE_INTEGER largeint;
    FILETIME ft[3];
    FILETIME *ftP[3];
    Tcl_Obj *objs[3];
    int i;
    int func = PtrToInt(clientdata);

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.value.uval = GetLogicalDrives();
        result.type = TRT_DWORD;
        break;
    case 2:
        return TwapiFirstVolume(interp, NULL); /* FindFirstVolume */
    case 3: // RegisterDirChangeNotifier
    case 4:
    case 5:
    case 6:
    case 7:
    case 8:
    case 9:
    case 10:
    case 11:
        if (TwapiGetArgs(interp, objc, objv, GETWSTR(s), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 3:
            return Twapi_QueryDosDevice(interp, s);
        case 4:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = DeleteVolumeMountPointW(s);
            break;
        case 5: // GetVolumeNameForVolumeMountPointW
        case 6: // GetVolumePathName
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival =
                (func == 5
                 ? GetVolumeNameForVolumeMountPointW
                 : GetVolumePathNameW)
                (s, buf, ARRAYSIZE(buf));
            if (result.value.ival) {
                result.value.unicode.str = buf;
                result.value.unicode.len = -1;
                result.type = TRT_UNICODE;
            }
            break;
        case 7:
            result.type = TRT_DWORD;
            result.value.uval = GetDriveTypeW(s);
            break;
        case 8:
            NULLIFY_EMPTY(s);
            return Twapi_GetVolumeInformation(interp, s);
        case 9:
            return TwapiFirstVolume(interp, s);
        case 10:
            return Twapi_GetDiskFreeSpaceEx(interp, s);
        case 11: // GetCompressedFileSize
            largeint.u.LowPart = GetCompressedFileSizeW(s, &largeint.u.HighPart);
            if (largeint.u.LowPart == INVALID_FILE_SIZE) {
                /* Does not mean failure, have to actually check */
                dw = GetLastError();
                if (dw != NO_ERROR) {
                    result.value.ival = dw;
                    result.type = TRT_EXCEPTION_ON_ERROR;
                    break;
                }
            }
            result.value.wide = largeint.QuadPart;
            result.type = TRT_WIDE;
            break;
        }
        break;
    case 12:
    case 13:
    case 14:
    case 15:
    case 16:
    case 17:
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(h),ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 12:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FindVolumeClose(h);
        break;
        case 13:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FindVolumeMountPointClose(h);
            break;
        case 14:
            return TwapiNextVolume(interp, 0, h);
        case 15:
            return TwapiNextVolume(interp, 1, h);
            break;
        case 16:
            return Twapi_GetFileType(interp, h);
        case 17:
            if (GetFileTime(h, &ft[0], &ft[1], &ft[2])) {
                objs[0] = ObjFromFILETIME(&ft[0]);
                objs[1] = ObjFromFILETIME(&ft[1]);
                objs[2] = ObjFromFILETIME(&ft[2]);
                result.type = TRT_OBJV;
                result.value.objv.objPP = objs;
                result.value.objv.nobj = 3;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        }
        break;
    case 18:
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETWSTR(s), GETNULLIFEMPTY(s2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = DefineDosDeviceW(dw, s, s2);
        break;
    case 19:
    case 20:
    case 21:
        if (TwapiGetArgs(interp, objc, objv,
                         GETWSTR(s), GETWSTR(s2),
                         ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 19:
            NULLIFY_EMPTY(s2);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = MoveFileExW(s,s2,dw);
            break;
        case 20:
            NULLIFY_EMPTY(s);
            NULLIFY_EMPTY(s2);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetVolumeLabelW(s,s2);
            break;
        case 21:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetVolumeMountPointW(s, s2);
            break;
        }
        break;
    case 22:
        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToHANDLE(interp, objv[0], &h) != TCL_OK)
            return TCL_ERROR;
        for (i = 0; i < 3; ++i) {
            if (Tcl_GetCharLength(objv[1+i]) == 0)
                ftP[i] = NULL;
            else {
                if (ObjToFILETIME(interp, objv[1+i], &ft[i]) != TCL_OK)
                    return TCL_ERROR;
                ftP[i] = &ft[i];
            }
        }
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SetFileTime(h, ftP[0], ftP[1], ftP[2]);
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiStorageInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s StorDispatch[] = {
        DEFINE_FNCODE_CMD(GetLogicalDrives, 1),
        DEFINE_FNCODE_CMD(FindFirstVolume, 2),
        DEFINE_FNCODE_CMD(QueryDosDevice, 3),
        DEFINE_FNCODE_CMD(DeleteVolumeMountPoint, 4),
        DEFINE_FNCODE_CMD(GetVolumeNameForVolumeMountPoint, 5),
        DEFINE_FNCODE_CMD(GetVolumePathName, 6),
        DEFINE_FNCODE_CMD(GetDriveType, 7),
        DEFINE_FNCODE_CMD(GetVolumeInformation, 8),
        DEFINE_FNCODE_CMD(FindFirstVolumeMountPoint, 9),
        DEFINE_FNCODE_CMD(GetDiskFreeSpaceEx, 10),
        DEFINE_FNCODE_CMD(GetCompressedFileSize, 11), // TBD - Tcl
        DEFINE_FNCODE_CMD(FindVolumeClose, 12),
        DEFINE_FNCODE_CMD(FindVolumeMountPointClose, 13),
        DEFINE_FNCODE_CMD(FindNextVolume, 14),
        DEFINE_FNCODE_CMD(FindNextVolumeMountPoint, 15),
        DEFINE_FNCODE_CMD(GetFileType, 16), // TBD - TCL 
        DEFINE_FNCODE_CMD(GetFileTime, 17),
        DEFINE_FNCODE_CMD(DefineDosDevice, 18),
        DEFINE_FNCODE_CMD(MoveFileEx, 19),
        DEFINE_FNCODE_CMD(SetVolumeLabel, 20),
        DEFINE_FNCODE_CMD(SetVolumeMountPoint, 21),
        DEFINE_FNCODE_CMD(SetFileTime, 22),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(StorDispatch), StorDispatch, Twapi_StorageCallObjCmd);
    Tcl_CreateObjCommand(interp, "twapi::Twapi_RegisterDirectoryMonitor", Twapi_RegisterDirectoryMonitorObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::Twapi_UnregisterDirectoryMonitor", Twapi_UnregisterDirectoryMonitorObjCmd, ticP, NULL);

    return TCL_OK;
}

static void TwapiStorageCleanup(TwapiInterpContext *ticP)
{
    if (ticP->module.data.pval) {
        TwapiStorageInterpContext *sicP =
            (TwapiStorageInterpContext *) ticP->module.data.pval;

        /* Shutdown all directory monitors */
        while (ZLIST_COUNT(&sicP->directory_monitors) != 0) {
            TwapiShutdownDirectoryMonitor(ZLIST_HEAD(&sicP->directory_monitors));
        }

        TwapiFree(ticP->module.data.pval);
        ticP->module.data.pval = NULL;
    }
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
int Twapi_storage_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiStorageInitCalls,
        TwapiStorageCleanup,
    };
    TwapiInterpContext *ticP;
    TwapiStorageInterpContext *sicP;
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    /* Cannot use DEFAULT_TIC because we have a cleanup routine and
     * use the ticP->module.data area
     */
    ticP = TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, NEW_TIC);

    if (ticP == NULL)
        return TCL_ERROR;

    sicP  = TwapiAlloc(sizeof(TwapiStorageInterpContext));
    ZLIST_INIT(&sicP->directory_monitors);
    ticP->module.data.pval = sicP;
    return TCL_OK;
}

