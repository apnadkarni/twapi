/* 
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

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

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
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
        objv[0] = Tcl_NewIntObj(1);
        objv[1] = ObjFromUnicode(buf);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
        return TCL_OK;
    } else {
        DWORD lasterr = GetLastError();
        buf[0] = 0;
        if (lasterr == ERROR_NO_MORE_FILES) {
            /* Not an error, signal no more volumes */
            Tcl_SetObjResult(interp, Tcl_NewIntObj(0));
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
    objv[1] = Tcl_NewLongObj(serial_no);
    objv[2] = Tcl_NewLongObj(max_component_len);
    objv[3] = Tcl_NewLongObj(sysflags);
    objv[4] = ObjFromUnicode(fsname);
    Tcl_SetObjResult(interp, Tcl_NewListObj(5, objv));

    return TCL_OK;
}

int Twapi_QueryDosDevice(TwapiInterpContext *ticP, LPCWSTR lpDeviceName)
{
    WCHAR *pathP;
    DWORD  path_cch;
    DWORD result;
    DWORD buf_sz;

    pathP = MemLifoPushFrame(&ticP->memlifo, 1000, &buf_sz);
    path_cch = buf_sz/sizeof(*pathP);
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
        MemLifoPopFrame(&ticP->memlifo);
        pathP = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR)*path_cch, NULL);
    }

    if (result) {
        Tcl_SetObjResult(ticP->interp, ObjFromMultiSz(pathP, result));
        result = TCL_OK;
    } else {
        result = TwapiReturnSystemError(ticP->interp);
    }

    MemLifoPopFrame(&ticP->memlifo);

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

    objv[0] = Tcl_NewWideIntObj(free_avail.QuadPart);
    objv[1] = Tcl_NewWideIntObj(total_bytes.QuadPart);
    objv[2] = Tcl_NewWideIntObj(free_total.QuadPart);
    Tcl_SetObjResult(interp, Tcl_NewListObj(3, objv));
    return TCL_OK;
}

