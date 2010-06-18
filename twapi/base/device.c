/*
 * Copyright (c) 2007-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

Tcl_Obj *ObjFromSP_DEVINFO_DATA(SP_DEVINFO_DATA *sddP)
{
    Tcl_Obj *objv[3];
    objv[0] = ObjFromGUID(&sddP->ClassGuid);
    objv[1] = Tcl_NewLongObj(sddP->DevInst);
    objv[3] = ObjFromDWORD_PTR(sddP->Reserved);
    return Tcl_NewListObj(3, objv);
}

/* objP may be NULL */
int ObjToSP_DEVINFO_DATA(Tcl_Interp *interp, Tcl_Obj *objP, SP_DEVINFO_DATA *sddP)
{
    if (objP) {
        /* Initialize based on passed param */
        Tcl_Obj **objs;
        int  nobjs;
        if (Tcl_ListObjGetElements(interp, objP, &nobjs, &objs) != TCL_OK ||
            TwapiGetArgs(interp, nobjs, objs,
                         ARGUSEDEFAULT,
                         GETVARWITHDEFAULT(sddP->ClassGuid, ObjToGUID),
                         GETINT(sddP->DevInst),
                         GETDWORD_PTR(sddP->Reserved)) != TCL_OK) {
            return TCL_ERROR;
        }
    } else
        ZeroMemory(sddP, sizeof(*sddP));

    sddP->cbSize = sizeof(*sddP);
    return TCL_OK;
}

/* sddPP MUST POINT TO VALID MEMORY */
int ObjToSP_DEVINFO_DATA_NULL(Tcl_Interp *interp, Tcl_Obj *objP, SP_DEVINFO_DATA **sddPP)
{
    int n;

    if (objP && Tcl_ListObjLength(interp, objP, &n) == TCL_OK && n != 0)
        return ObjToSP_DEVINFO_DATA(interp, objP, *sddPP);

    *sddPP = NULL;
    return TCL_OK;
}

Tcl_Obj *ObjFromSP_DEVICE_INTERFACE_DATA(SP_DEVICE_INTERFACE_DATA *sdidP)
{
    Tcl_Obj *objv[3];
    objv[0] = ObjFromGUID(&sdidP->InterfaceClassGuid);
    objv[1] = Tcl_NewLongObj(sdidP->Flags);
    objv[3] = ObjFromDWORD_PTR(sdidP->Reserved);
    return Tcl_NewListObj(3, objv);
}

/* objP may be NULL */
int ObjToSP_DEVICE_INTERFACE_DATA(Tcl_Interp *interp, Tcl_Obj *objP, SP_DEVICE_INTERFACE_DATA *sdiP)
{
    if (objP) {
        /* Initialize based on passed param */
        Tcl_Obj **objs;
        int  nobjs;
        if (Tcl_ListObjGetElements(interp, objP, &nobjs, &objs) != TCL_OK ||
            TwapiGetArgs(interp, nobjs, objs,
                         ARGUSEDEFAULT,
                         GETVARWITHDEFAULT(sdiP->InterfaceClassGuid, ObjToGUID),
                         GETINT(sdiP->Flags),
                         GETDWORD_PTR(sdiP->Reserved)) != TCL_OK) {
            return TCL_ERROR;
        }
    } else
        ZeroMemory(sdiP, sizeof(*sdiP));

    sdiP->cbSize = sizeof(*sdiP);
    return TCL_OK;
}

int Twapi_SetupDiGetDeviceRegistryProperty(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HDEVINFO hdi;
    SP_DEVINFO_DATA sdd;
    DWORD regprop;
    DWORD regtype;
    BYTE  buf[MAX_PATH+1];
    BYTE *bufP;
    DWORD buf_sz;
    int tcl_status = TCL_ERROR;
    Tcl_Obj *objP;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLET(hdi, HDEVINFO),
                     GETVAR(sdd, ObjToSP_DEVINFO_DATA),
                     GETINT(regprop),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* We will try to first retrieve using a stack buffer.
       If that fails, retry with malloc*/
    bufP = buf;
    buf_sz = sizeof(buf);
    if (! SetupDiGetDeviceRegistryPropertyW(
            hdi, &sdd, regprop, &regtype, bufP, buf_sz, &buf_sz)) {
        /* Unsuccessful call. See if we need a larger buffer */
        DWORD winerr = GetLastError();
        if (winerr != ERROR_INSUFFICIENT_BUFFER)
            return Twapi_AppendSystemError(interp, winerr);

        /* Try again with larger buffer - required size was
           returned in buf_sz */
        if (Twapi_malloc(interp, NULL, buf_sz, &bufP) != TCL_OK)
            return TCL_ERROR;
        /* Retry the call */
        if (! SetupDiGetDeviceRegistryPropertyW(
                hdi, &sdd, regprop, &regtype, bufP, buf_sz, &buf_sz)) {
            TwapiReturnSystemError(interp); /* Still failed */
            goto vamoose;
        }
    }

    /* Success. regprop contains the registry property type */
    objP = ObjFromRegValue(interp, regtype, bufP, buf_sz);
    if (objP) {
        Tcl_SetObjResult(interp, objP);
        tcl_status = TCL_OK;
    }

vamoose:
    if (bufP != buf)
        free(buf);
    return tcl_status;
}

int Twapi_SetupDiGetDeviceInterfaceDetail(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HDEVINFO hdi;
    SP_DEVICE_INTERFACE_DATA  sdid;
    SP_DEVINFO_DATA sdd;
    struct {
        SP_DEVICE_INTERFACE_DETAIL_DATA_W  sdidd;
        WCHAR extra[MAX_PATH+1];
    } buf;
    SP_DEVICE_INTERFACE_DETAIL_DATA_W *sdiddP;
    DWORD buf_sz;
    Tcl_Obj *objs[2];
    int success;
    DWORD winerr;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLET(hdi, HDEVINFO),
                     GETVAR(sdid, ObjToSP_DEVICE_INTERFACE_DATA),
                     GETVAR(sdd, ObjToSP_DEVINFO_DATA),
                     ARGEND) != TCL_OK)
            return TCL_ERROR;

    buf_sz = sizeof(buf);
    sdiddP = &buf.sdidd;
    while (sdiddP) {
        sdiddP->cbSize = sizeof(*sdiddP); /* NOT size of entire buffer */
        success = SetupDiGetDeviceInterfaceDetailW(
            hdi, &sdid, sdiddP, buf_sz, &buf_sz, &sdd);
        if (success || (winerr = GetLastError()) != ERROR_INSUFFICIENT_BUFFER)
            break;
        /* Retry with larger buffer size as returned in call */
        if (sdiddP != &buf.sdidd)
            free(sdiddP);
        sdiddP = (SP_DEVICE_INTERFACE_DETAIL_DATA_W *) malloc(buf_sz);
    }

    if (success) {
        objs[0] = Tcl_NewUnicodeObj(sdiddP->DevicePath, -1);
        objs[1] = ObjFromSP_DEVINFO_DATA(&sdd);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objs));
    } else
        Twapi_AppendSystemError(interp, winerr);

    if (sdiddP != &buf.sdidd)
        free(sdiddP);

    return success ? TCL_OK : TCL_ERROR;
}

int Twapi_SetupDiClassGuidsFromNameEx(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR class_name;
    GUID guids[32];
    LPWSTR system_name;
    GUID *guidP;
    DWORD allocated;
    DWORD needed;
    int success;
    void  *reserved;
    DWORD  i;

    if (TwapiGetArgs(interp, objc, objv,
                     GETWSTR(class_name),
                     ARGUSEDEFAULT,
                     GETNULLIFEMPTY(system_name),
                     GETVOIDP(reserved),
                     ARGEND) != TCL_OK)
            return TCL_ERROR;

    allocated = ARRAYSIZE(guids);
    guidP = guids;
    success = 0;
    while (guidP) {
        if (! SetupDiClassGuidsFromNameExW(class_name, guidP, allocated,
                                           &needed, system_name, reserved))
            break;
        if (needed <= allocated) {
            success = 1;
            break;
        }
        /* Retry with larger buffer size as returned in call */
        if (guidP != guids)
            free(guidP);
        allocated = needed;
        guidP = (GUID *) malloc(sizeof(GUID*) * allocated);
    }

    if (success) {
        Tcl_Obj *objP = Tcl_NewListObj(0, NULL);
        /* Note - use 'needed', not 'allocated' in loop! */
        for (i = 0; i < needed; ++i) {
            Tcl_ListObjAppendElement(interp, objP, ObjFromGUID(&guidP[i]));
        }
        Tcl_SetObjResult(interp, objP);
    } else
        Twapi_AppendSystemError(interp, GetLastError());

    if (guidP != guids)
        free(guidP);

    return success ? TCL_OK : TCL_ERROR;
}


