/*
 * Copyright (c) 2007-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_wm.h"
#include <setupapi.h>
#if !defined(TWAPI_REPLACE_CRT) && !defined(TWAPI_MINIMIZE_CRT)
# include <process.h>
#endif
#include <dbt.h>
#include <initguid.h>           // Instantiate ioevent.h guids
#include <ioevent.h>            // Custom GUID definitions
#include <cfgmgr32.h>           /* PnP CM_* calls */

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

static GUID gNullGuid;                 /* Init'ed to 0 */

typedef struct _TwapiDeviceNotificationContext TwapiDeviceNotificationContext;
ZLINK_CREATE_TYPEDEFS(TwapiDeviceNotificationContext); 

/*
 * Struct used to pass around device notification context. Note it is a
 * part of the fundamental design that an allocated context is in use
 * in only one thread at a time. Hence there is no synchronization between
 * threads. However, because it is passed to window procs, reentrancy
 * is a possibility and care must be taken to accordingly.
 */
typedef struct _TwapiDeviceNotificationContext {
    TwapiInterpContext *ticP;
    TwapiId   id;
    DWORD devtype;
    union {
        GUID  guid;                 /* GUID for notifications */
        HANDLE hdev;                /* Device handle for which notifications
                                   are to be received */
    } device;                          /* Based on devtype */
    int    nrefs;               /* Ref count - not interlocked because
                                   only accessed from one thread at a time */

    /* Remaining fields are only used in the notification thread itself */
    ZLINK_DECL(TwapiDeviceNotificationContext); /* Links all registrations */
    HWND  hwnd;                           /* Window used for notifications */
    HANDLE hnotification;                 /* Handle for notification delivery */
} TwapiDeviceNotificationContext;

/*
 * Create list header definitions for list of TwapiDeviceNotificationContext
 * This list is only accessed from the single device notification thread
 * and hence requires no locking.
 */
ZLIST_CREATE_TYPEDEFS(TwapiDeviceNotificationContext);
ZLIST_DECL(TwapiDeviceNotificationContext) gTwapiDeviceNotificationRegistry;

/*
 * Device notification events delivered via the Tcl event queue.
 *
 * NOTE THIS IS A VARIABLE SIZE STRUCTURE AS DEVICE_DEPENDENT DATA IS
 * TACKED ON AT THE END.
 */
typedef struct _TwapiDeviceNotificationCallback {
    TwapiCallback  cb;   /* Must be first field */
    union {
        struct {
            WPARAM wparam;              /* wparam from the WM_DEVICECHANGE message */
            DEV_BROADCAST_HDR dev_bcast_hdr; /* VARIABLE SIZED. MUST BE LAST */
        } device;
        char msg[1];            /* Actually variable size */
    } data;
} TwapiDeviceNotificationCallback;

static TwapiOneTimeInitState TwapiDeviceModuleInitialized;
static DWORD TwapiDeviceNotificationTid;

int ObjToSP_DEVINFO_DATA(Tcl_Interp *, Tcl_Obj *objP, SP_DEVINFO_DATA *sddP);
int ObjToSP_DEVINFO_DATA_NULL(Tcl_Interp *interp, Tcl_Obj *objP,
                              SP_DEVINFO_DATA **sddPP);
Tcl_Obj *ObjFromSP_DEVINFO_DATA(SP_DEVINFO_DATA *sddP);
int ObjToSP_DEVICE_INTERFACE_DATA(Tcl_Interp *interp, Tcl_Obj *objP,
                                  SP_DEVICE_INTERFACE_DATA *sdidP);
Tcl_Obj *ObjFromSP_DEVICE_INTERFACE_DATA(SP_DEVICE_INTERFACE_DATA *sdidP);
TwapiDeviceNotificationContext *TwapiDeviceNotificationContextNew(
    TwapiInterpContext *ticP);
void TwapiDeviceNotificationContextDelete(TwapiDeviceNotificationContext *dncP);
void TwapiDeviceNotificationContextUnref(TwapiDeviceNotificationContext *, int);
/* Note no locking necessary as device notification contexts are only accessed
   from one thread at a time */
#define TwapiDeviceNotificationContextRef(p_, incr_) \
    do {(p_)->nrefs += (incr_);} while (0)
static void TwapiReportDeviceNotificationError(
    TwapiDeviceNotificationContext *dncP, char *msg, DWORD winerr);
static Tcl_Obj *ObjFromCustomDeviceNotification(PDEV_BROADCAST_HDR  dbhP);
static Tcl_Obj *ObjFromDevtype(DWORD devtype);
static TwapiDeviceNotificationContext *TwapiFindDeviceNotificationById(TwapiId id);
static TwapiDeviceNotificationContext *TwapiFindDeviceNotificationByHwnd(HWND);
int Twapi_SetupDiGetDeviceRegistryProperty(TwapiInterpContext *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_SetupDiGetDeviceInterfaceDetail(TwapiInterpContext *, int objc,
                                          Tcl_Obj *CONST objv[]);
int Twapi_SetupDiClassGuidsFromNameEx(TwapiInterpContext *, int objc,
                                      Tcl_Obj *CONST objv[]);
int Twapi_RegisterDeviceNotification(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[]);
int Twapi_UnregisterDeviceNotification(TwapiInterpContext *ticP, TwapiId id);


Tcl_Obj *ObjFromSP_DEVINFO_DATA(SP_DEVINFO_DATA *sddP)
{
    Tcl_Obj *objv[3];
    objv[0] = ObjFromGUID(&sddP->ClassGuid);
    objv[1] = ObjFromLong(sddP->DevInst);
    objv[2] = ObjFromDWORD_PTR(sddP->Reserved);
    return ObjNewList(3, objv);
}

/* objP may be NULL */
int ObjToSP_DEVINFO_DATA(Tcl_Interp *interp, Tcl_Obj *objP, SP_DEVINFO_DATA *sddP)
{
    if (objP) {
        /* Initialize based on passed param */
        Tcl_Obj **objs;
        int  nobjs;
        if (ObjGetElements(interp, objP, &nobjs, &objs) != TCL_OK ||
            TwapiGetArgs(interp, nobjs, objs,
                         ARGUSEDEFAULT,
                         GETVARWITHDEFAULT(sddP->ClassGuid, ObjToGUID),
                         GETINT(sddP->DevInst),
                         GETDWORD_PTR(sddP->Reserved), ARGEND) != TCL_OK) {
            return TCL_ERROR;
        }
    } else
        TwapiZeroMemory(sddP, sizeof(*sddP));

    sddP->cbSize = sizeof(*sddP);
    return TCL_OK;
}

/* sddPP MUST POINT TO VALID MEMORY */
int ObjToSP_DEVINFO_DATA_NULL(Tcl_Interp *interp, Tcl_Obj *objP, SP_DEVINFO_DATA **sddPP)
{
    int n;

    if (objP && ObjListLength(interp, objP, &n) == TCL_OK && n != 0)
        return ObjToSP_DEVINFO_DATA(interp, objP, *sddPP);

    *sddPP = NULL;
    return TCL_OK;
}

Tcl_Obj *ObjFromSP_DEVICE_INTERFACE_DATA(SP_DEVICE_INTERFACE_DATA *sdidP)
{
    Tcl_Obj *objv[3];
    objv[0] = ObjFromGUID(&sdidP->InterfaceClassGuid);
    objv[1] = ObjFromLong(sdidP->Flags);
    objv[2] = ObjFromDWORD_PTR(sdidP->Reserved);
    return ObjNewList(3, objv);
}

/* objP may be NULL */
int ObjToSP_DEVICE_INTERFACE_DATA(Tcl_Interp *interp, Tcl_Obj *objP, SP_DEVICE_INTERFACE_DATA *sdiP)
{
    if (objP) {
        /* Initialize based on passed param */
        Tcl_Obj **objs;
        int  nobjs;
        if (ObjGetElements(interp, objP, &nobjs, &objs) != TCL_OK ||
            TwapiGetArgs(interp, nobjs, objs,
                         ARGUSEDEFAULT,
                         GETVARWITHDEFAULT(sdiP->InterfaceClassGuid, ObjToGUID),
                         GETINT(sdiP->Flags),
                         GETDWORD_PTR(sdiP->Reserved), ARGEND) != TCL_OK) {
            return TCL_ERROR;
        }
    } else
        TwapiZeroMemory(sdiP, sizeof(*sdiP));

    sdiP->cbSize = sizeof(*sdiP);
    return TCL_OK;
}

int Twapi_SetupDiGetDeviceRegistryProperty(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    HDEVINFO hdi;
    SP_DEVINFO_DATA sdd;
    DWORD regprop;
    DWORD regtype;
    BYTE *bufP;
    DWORD buf_sz;
    int tcl_status = TCL_ERROR;
    Tcl_Obj *objP;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETHANDLET(hdi, HDEVINFO),
                     GETVAR(sdd, ObjToSP_DEVINFO_DATA),
                     GETINT(regprop),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    bufP = MemLifoPushFrame(ticP->memlifoP, 256, &buf_sz);
    if (! SetupDiGetDeviceRegistryPropertyW(
            hdi, &sdd, regprop, &regtype, bufP, buf_sz, &buf_sz)) {
        /* Unsuccessful call. See if we need a larger buffer */
        DWORD winerr = GetLastError();
        if (winerr != ERROR_INSUFFICIENT_BUFFER) {
            Twapi_AppendSystemError(ticP->interp, HRESULT_FROM_SETUPAPI(winerr));
            goto vamoose;
        }

        /* Try again with larger buffer, don't realloc as that will
           unnecessarily copy */
        bufP = MemLifoAlloc(ticP->memlifoP, buf_sz, NULL);
        /* Retry the call */
        if (! SetupDiGetDeviceRegistryPropertyW(
                hdi, &sdd, regprop, &regtype, bufP, buf_sz, &buf_sz)) {
            TwapiReturnSystemError(ticP->interp); /* Still failed */
            goto vamoose;
        }
    }

    /* Success. regprop contains the registry property type */
    objP = ObjFromRegValueCooked(ticP->interp, regtype, bufP, buf_sz);
    if (objP) {
        ObjSetResult(ticP->interp, objP);
        tcl_status = TCL_OK;
    }

vamoose:
    MemLifoPopFrame(ticP->memlifoP);
    return tcl_status;
}

int Twapi_SetupDiGetDeviceInterfaceDetail(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    HDEVINFO hdi;
    SP_DEVICE_INTERFACE_DATA  sdid;
    SP_DEVINFO_DATA sdd;
    SP_DEVICE_INTERFACE_DETAIL_DATA_W *sdiddP;
    DWORD buf_sz;
    Tcl_Obj *objs[2];
    int success;
    DWORD winerr;
    int   i;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETHANDLET(hdi, HDEVINFO),
                     GETVAR(sdid, ObjToSP_DEVICE_INTERFACE_DATA),
                     GETVAR(sdd, ObjToSP_DEVINFO_DATA),
                     ARGEND) != TCL_OK)
            return TCL_ERROR;

    sdiddP = MemLifoPushFrame(ticP->memlifoP, sizeof(*sdiddP)+MAX_PATH, &buf_sz);
    /* To be safe against bugs, ours or the driver's limit to 5 attempts */
    for (i = 0; i < 5; ++i) {
        sdiddP->cbSize = sizeof(*sdiddP); /* NOT size of entire buffer */
        success = SetupDiGetDeviceInterfaceDetailW(
            hdi, &sdid, sdiddP, buf_sz, &buf_sz, &sdd);
        if (success || (winerr = GetLastError()) != ERROR_INSUFFICIENT_BUFFER)
            break;
        /* Retry with larger buffer size as returned in call */
        sdiddP = MemLifoAlloc(ticP->memlifoP, buf_sz, NULL);
    }

    if (success) {
        objs[0] = ObjFromUnicode(sdiddP->DevicePath);
        objs[1] = ObjFromSP_DEVINFO_DATA(&sdd);
        ObjSetResult(ticP->interp, ObjNewList(2, objs));
    } else
        Twapi_AppendSystemError(ticP->interp, HRESULT_FROM_SETUPAPI(winerr));

    MemLifoPopFrame(ticP->memlifoP);

    return success ? TCL_OK : TCL_ERROR;
}

int Twapi_SetupDiClassGuidsFromNameEx(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    LPWSTR class_name;
    LPWSTR system_name;
    GUID *guidP;
    DWORD allocated;
    DWORD needed;
    int success;
    void  *reserved;
    DWORD  i;
    DWORD buf_sz;
    MemLifoMarkHandle mark;

    mark = MemLifoPushMark(ticP->memlifoP);

    success = 0;
    if (TwapiGetArgsEx(ticP, objc, objv,
                       GETWSTR(class_name),
                       ARGUSEDEFAULT,
                       GETEMPTYASNULL(system_name),
                       GETVOIDP(reserved),
                       ARGEND) == TCL_OK) {

        guidP = MemLifoAlloc(ticP->memlifoP, 10 * sizeof(*guidP), &buf_sz);
        allocated = buf_sz / sizeof(*guidP);

        /*
         * We keep looping on success returns because a success return does
         * not mean all entries were retrieved.
         */
        while (SetupDiClassGuidsFromNameExW(class_name, guidP, allocated,
                                            &needed, system_name, reserved)) {
            if (needed <= allocated) {
                /* All retrieved */
                success = 1;
                break;
            }
            /* Retry with larger buffer size as returned in call */
            allocated = needed;
            MemLifoAlloc(ticP->memlifoP, allocated * sizeof(*guidP), NULL);
        }

        if (success) {
            Tcl_Obj *objP = ObjNewList(0, NULL);
            /* Note - use 'needed', not 'allocated' in loop! */
            for (i = 0; i < needed; ++i) {
                ObjAppendElement(ticP->interp, objP, ObjFromGUID(&guidP[i]));
            }
            ObjSetResult(ticP->interp, objP);
        } else {
            i = GetLastError();
            Twapi_AppendSystemError(ticP->interp, HRESULT_FROM_SETUPAPI(i));
        }
    }

    MemLifoPopMark(mark);
    return success ? TCL_OK : TCL_ERROR;
}

static int TwapiDeviceNotificationCallbackFn(TwapiCallback *p)
{
    TwapiDeviceNotificationCallback *cbP = (TwapiDeviceNotificationCallback *)p;
    PDEV_BROADCAST_HDR dbhP;
    char *notification_str = NULL;
    TwapiResultType response_type = TRT_EMPTY;

    Tcl_Obj *objs[10];
    int nobjs;

    if (cbP->cb.ticP->interp == NULL ||
        Tcl_InterpDeleted(cbP->cb.ticP->interp)) {
        cbP->cb.winerr = ERROR_INVALID_STATE; /* Best match we can find */
        cbP->cb.response.type = TRT_EMPTY;
        return TCL_ERROR;
    }

    /* Deal with the error notification case first. */
    if (cbP->cb.winerr != ERROR_SUCCESS) {
        objs[0] = ObjFromString(TWAPI_TCL_NAMESPACE "::_device_notification_handler");
        objs[1] = ObjFromTwapiId(cbP->cb.receiver_id);
        objs[2] = STRING_LITERAL_OBJ("error");
        objs[3] = ObjFromLong(cbP->cb.winerr);
        return TwapiEvalAndUpdateCallback(&cbP->cb, 4, objs, TRT_EMPTY);
    }

    dbhP = &cbP->data.device.dev_bcast_hdr;

    /* Note objs[0..2] are common and will be filled at end */
    nobjs = 3;
    switch (cbP->data.device.wparam) {
    case DBT_CONFIGCHANGECANCELED:
        notification_str = "configchangecanceled";
        break;

    case DBT_CONFIGCHANGED:
        notification_str = "configchanged";
        break;

    case DBT_CUSTOMEVENT:
        notification_str = "customevent";
        objs[nobjs++] = ObjFromDevtype(dbhP->dbch_devicetype);
        objs[nobjs++] = ObjFromCustomDeviceNotification(dbhP);
        break;

    case DBT_DEVICEARRIVAL: /* Fall thru */
    case DBT_DEVICEQUERYREMOVE: /* Fall thru */
    case DBT_DEVICEQUERYREMOVEFAILED: /* Fall thru */
    case DBT_DEVICEREMOVECOMPLETE:
    case DBT_DEVICEREMOVEPENDING:
    case DBT_DEVICETYPESPECIFIC:
        switch (cbP->data.device.wparam) {
        case DBT_DEVICEARRIVAL:
            notification_str = "devicearrival";
            break;
        case DBT_DEVICEQUERYREMOVE:
            notification_str = "devicequeryremove";
            response_type = TRT_BOOL;
            break;
        case DBT_DEVICEQUERYREMOVEFAILED: /* Fall thru */
            notification_str = "devicequeryremovefailed";
            break;
        case DBT_DEVICEREMOVECOMPLETE:
            notification_str = "deviceremovecomplete";
            break;
        case DBT_DEVICEREMOVEPENDING: 
            notification_str = "deviceremovepending";
            break;
        case DBT_DEVICETYPESPECIFIC: 
            notification_str = "devicetypespecific";
            break;
       }

        objs[nobjs++] = ObjFromDevtype(dbhP->dbch_devicetype);
        switch( dbhP->dbch_devicetype ) {
        case DBT_DEVTYP_DEVICEINTERFACE:
            /* First tack on the GUID, then the name */
            objs[nobjs++] = ObjFromGUID(&((PDEV_BROADCAST_DEVICEINTERFACE_W)dbhP)->dbcc_classguid);
            objs[nobjs++] = ObjFromUnicode(((PDEV_BROADCAST_DEVICEINTERFACE_W)dbhP)->dbcc_name);
            break;

        case DBT_DEVTYP_HANDLE:
            objs[nobjs++] = ObjFromHANDLE(((PDEV_BROADCAST_HANDLE)dbhP)->dbch_handle);
            objs[nobjs++] = ObjFromOpaque(((PDEV_BROADCAST_HANDLE)dbhP)->dbch_hdevnotify, "HDEVNOTIFY");

            /*
             * No additional arguments are passed since there is no other
             * useful information in the structure unless the event is
             * DBT_CUSTOMEVENT which is not handled here.
             */
            break;

        case DBT_DEVTYP_OEM:
            objs[nobjs++] = ObjFromLong(((PDEV_BROADCAST_OEM)dbhP)->dbco_identifier);
            objs[nobjs++] = ObjFromLong(((PDEV_BROADCAST_OEM)dbhP)->dbco_suppfunc);
            break;

        case DBT_DEVTYP_PORT:
            objs[nobjs++] = ObjFromUnicode(((PDEV_BROADCAST_PORT_W)dbhP)->dbcp_name);
            break;

        case DBT_DEVTYP_VOLUME:
            objs[nobjs++] = ObjFromLong(((PDEV_BROADCAST_VOLUME)dbhP)->dbcv_unitmask);
            objs[nobjs++] = ObjFromLong(((PDEV_BROADCAST_VOLUME)dbhP)->dbcv_flags);
            break;
        }
        
        break;

    case DBT_DEVNODES_CHANGED:
        notification_str = "devnodes_changed";
        break;

    case DBT_QUERYCHANGECONFIG:
        notification_str = "querychangeconfig";
        response_type = TRT_BOOL;     /* Force using response from script */
        break;

    case DBT_USERDEFINED:
        notification_str = "userdefined";
        objs[nobjs++] = ObjFromString(((struct _DEV_BROADCAST_USERDEFINED *)dbhP)->dbud_szName);
        break;

    default:
        break;            // No idea what this is, just ignore
    }

    /* Be paranoid in case we add more objects later and forget to grow array */
    if (nobjs > ARRAYSIZE(objs))
        Tcl_Panic("Internal error: exceeded bounds (%d) of device notification array", nobjs);

    objs[0] = STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_device_notification_handler");
    objs[1] = ObjFromTwapiId(cbP->cb.receiver_id);
    objs[2] = ObjFromString(notification_str);
    if (response_type == TRT_EMPTY) {
        /*
         * Return true, even on errors as we do not want to block a
         * notification. Likely will be ignored anyway, but just in case
         */
        if (TwapiEvalAndUpdateCallback(&cbP->cb, nobjs, objs, response_type) != TCL_OK) {
            /* TBD - log background error ? */
            TwapiClearResult(&cbP->cb.response);
        }
        cbP->cb.response.type = TRT_BOOL;
        cbP->cb.response.value.bval = 1;
        return TCL_OK;
    } else 
        return TwapiEvalAndUpdateCallback(&cbP->cb, nobjs, objs, response_type);
}

/* 
 * Cleans up a device notification window, including unregistering the
 * notification, removing from registered queue, deref'ing objects etc.
 */
static LRESULT TwapiDeviceNotificationWinCleanup(HWND hwnd)
{
    TwapiDeviceNotificationContext *dncP;

    /* Find the window in the registration chain */
    dncP = TwapiFindDeviceNotificationByHwnd(hwnd);
    if (dncP == NULL)
        return 0;               /* Can this happen? */

    /* Note single thread access so no locking anywhere here */

    if (dncP->hnotification)
        UnregisterDeviceNotification(dncP->hnotification);
    dncP->hnotification = NULL; /* Else we will repeat when actually freeing */

    /*
     * Unlink from the registered notificaitons chain and window before
     * freeing it up.
     */
    ZLIST_REMOVE(&gTwapiDeviceNotificationRegistry, dncP);

    /*
     * Window data points to dncP via TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET.
     * The hidden window andler will remove it so do the corresponding deref.
     * In addition, we removed from the notification registry list above so
     * deref once more for a total of 2 derefs.
     *
     * Note unlinking from ticP etc. will happen inside Unref if required.
     */
    TwapiDeviceNotificationContextUnref(dncP, 2);
    dncP = NULL;                /* Make sure we do not access it */

    return 0;
}


static LRESULT TwapiDeviceNotificationWinProc(
    TwapiInterpContext *notused, /* NULL */
    LONG_PTR clientdata,
    HWND hwnd,
    UINT msg,
    WPARAM wparam,
    LPARAM lParam)
{
    TwapiDeviceNotificationContext *dncP =
        (TwapiDeviceNotificationContext *) clientdata;
    PDEV_BROADCAST_HDR dbhP = NULL;
    TwapiDeviceNotificationCallback *cbP;
    int    need_response = 0;
    LRESULT lres = TRUE;

    if (msg == WM_DESTROY) {
        return TwapiDeviceNotificationWinCleanup(hwnd);
    } else if (msg != WM_DEVICECHANGE) {
        return 0;
    }

    /* msg == WM_DEVICECHANGE at this point */

    switch (wparam) {
    case DBT_CONFIGCHANGECANCELED:
        break;

    case DBT_CONFIGCHANGED:
        break;

    case DBT_CUSTOMEVENT:
        dbhP = (PDEV_BROADCAST_HDR) lParam; /* May still be NULL! */
        break;

    case DBT_DEVICEARRIVAL: /* Fall thru */
    case DBT_DEVICEQUERYREMOVE: /* Fall thru */
    case DBT_DEVICEQUERYREMOVEFAILED: /* Fall thru */
    case DBT_DEVICEREMOVECOMPLETE:
    case DBT_DEVICEREMOVEPENDING:
    case DBT_DEVICETYPESPECIFIC:
        // Some messages, like DBT_DEVTYP_VOLUME are always broadcast to
        // all windows, even when a specify device interface was registered.
        // Filter these out by checking device type.
        // TBD - do we really want to ? Or should we provide that as an
        // option?
        dbhP = (PDEV_BROADCAST_HDR) lParam;
        if (dbhP == NULL || dbhP->dbch_devicetype != dncP->devtype)
            return TRUE;
        
        if (dbhP->dbch_devicetype == DBT_DEVTYP_DEVICEINTERFACE &&
            (! IsEqualGUID(&dncP->device.guid, &gNullGuid)) &&
            (! IsEqualGUID(&dncP->device.guid,
                           & (((PDEV_BROADCAST_DEVICEINTERFACE_W)dbhP)->dbcc_classguid)))) {
            // We want a specific GUID and this one does not
            // match what we are looking for
            // TBD - not sure this additional check is really needed.
            // The window should not receive any other guids I think
            return TRUE;
        }

        switch (wparam) {
        case DBT_DEVICEARRIVAL:
            break;
        case DBT_DEVICEQUERYREMOVE:
            need_response = 1;     // Force using response from script
            break;
        case DBT_DEVICEQUERYREMOVEFAILED: /* Fall thru */
            break;
        case DBT_DEVICEREMOVECOMPLETE:
            break;
        case DBT_DEVICEREMOVEPENDING: 
            break;
        case DBT_DEVICETYPESPECIFIC: 
            break;
        }
        break;

    case DBT_DEVNODES_CHANGED:
        break;

    case DBT_QUERYCHANGECONFIG:
        need_response = TRUE;     // Force using response from script
        break;

    case DBT_USERDEFINED:
        dbhP = (PDEV_BROADCAST_HDR) lParam;
        break;

    default:
        return TRUE;            // No idea what this is, just ignore
    }

    /* Build the device notification Tcl event */
    if (dbhP) {
        /* Need to tack on additional data */
        cbP = (TwapiDeviceNotificationCallback *)
            TwapiCallbackNew(dncP->ticP,
                                    TwapiDeviceNotificationCallbackFn,
                                    sizeof(*cbP) + dbhP->dbch_size - sizeof(cbP->data.device.dev_bcast_hdr));
        CopyMemory(&cbP->data.device.dev_bcast_hdr, dbhP, dbhP->dbch_size);
    } else {
        /* No additional data */
        cbP = (TwapiDeviceNotificationCallback *)
            TwapiCallbackNew(dncP->ticP,
                                    TwapiDeviceNotificationCallbackFn,
                                    sizeof(*cbP));
    }
    cbP->cb.receiver_id = dncP->id;
    cbP->data.device.wparam = wparam;
    cbP->cb.winerr = ERROR_SUCCESS;
    if (need_response) {
        if (TwapiEnqueueCallback(dncP->ticP,
                                 &cbP->cb,
                                 TWAPI_ENQUEUE_DIRECT,
                                 30*1000, /* TBD - Timeout (ms) */
                                 (TwapiCallback **)&cbP)
            == ERROR_SUCCESS) {
            if (cbP && cbP->cb.response.type == TRT_BOOL)
                lres = cbP->cb.response.value.bval;
        }
        if (cbP)
            TwapiCallbackUnref((TwapiCallback *)cbP, 1);
    } else {
        TwapiEnqueueCallback(dncP->ticP, &cbP->cb,
                             TWAPI_ENQUEUE_DIRECT,
                             0, /* No response wanted */
                             NULL);
        /* TBD - on error, do we send an error notification ? */
    }
    return lres;
}

/*
 * Creates a window for registering device notifications. Because callers
 * might issue various combinations of options, we do not share a window
 * as considerable bookkeeping would be required to pass on a single
 * notification to multiple sinks. 
 * 
 * On success stores the window handle in dncP->hwnd and also adds it
 * to list of device notification registrations.
 * Returns TCL_OK or TCL_ERROR.
 */
static int TwapiCreateDeviceNotificationWindow(TwapiDeviceNotificationContext *dncP)
{
    DWORD notify_flags;

    if (Twapi_CreateHiddenWindow(NULL, TwapiDeviceNotificationWinProc,
                                 (LONG_PTR) dncP, &dncP->hwnd) != TCL_OK) {
        return TCL_ERROR;
    }
    TwapiDeviceNotificationContextRef(dncP, 1); /* Window data points to it */

    /* Port and volume notifications are always sent by the system without
       registering. Only register for device interfaces and handles */
    
    if (dncP->devtype == DBT_DEVTYP_HANDLE) {
        DEV_BROADCAST_HANDLE h_filter;

        TwapiZeroMemory(&h_filter, sizeof(h_filter));
        h_filter.dbch_size = sizeof(h_filter);
        h_filter.dbch_devicetype = DBT_DEVTYP_HANDLE;
        notify_flags = DEVICE_NOTIFY_WINDOW_HANDLE;
        h_filter.dbch_handle = dncP->device.hdev;
        dncP->hnotification =
            RegisterDeviceNotificationW(dncP->hwnd, &h_filter, notify_flags);
        if (dncP->hnotification == NULL) {
            TwapiReportDeviceNotificationError(dncP, "RegisterDeviceNotificationW failed.", GetLastError());
            DestroyWindow(dncP->hwnd);
            return TCL_ERROR;
        }
    }
    else if (dncP->devtype == DBT_DEVTYP_DEVICEINTERFACE) {
        DEV_BROADCAST_DEVICEINTERFACE di_filter;

        TwapiZeroMemory(&di_filter, sizeof(di_filter));
        di_filter.dbcc_size = sizeof(di_filter);
        di_filter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
        notify_flags = DEVICE_NOTIFY_WINDOW_HANDLE;
        if (! IsEqualGUID(&dncP->device.guid, &gNullGuid)) {
            // Specific GUID requested
            di_filter.dbcc_classguid = dncP->device.guid;
        }
        else {
            notify_flags |= 0x4; // DEVICE_NOTIFY_ALL_INTERFACE_CLASSES
        }
        dncP->hnotification =
            RegisterDeviceNotificationW(dncP->hwnd, &di_filter, notify_flags);
        if (dncP->hnotification == NULL) {
            TwapiReportDeviceNotificationError(dncP, "RegisterDeviceNotificationW failed.", GetLastError());
            return TCL_ERROR;
        }
    }

    /* Add to our list of registrations */
    TwapiDeviceNotificationContextRef(dncP, 1);
    ZLIST_APPEND(&gTwapiDeviceNotificationRegistry, dncP);
    return TCL_OK;

}


/* There is one thread used for device notifications for all threads/interps */
static unsigned __stdcall TwapiDeviceNotificationThread(HANDLE sig)
{
    MSG msg;
    TwapiDeviceNotificationContext *dncP;

    /* Keeps track of all device registrations */
    ZLIST_INIT(&gTwapiDeviceNotificationRegistry);      /* TBD - free this registry */

    /*
     * Initialize the message queue so the initiating interp can post
     * messages to us. (See SDK docs)
     */
    PeekMessage(&msg, NULL, WM_USER, WM_USER, PM_NOREMOVE);
    
    /* Now inform the creator that it's free to send us messages */
    SetEvent(sig);

    /* Now start processing messages */
    while (GetMessage(&msg, NULL, 0, 0)) { 
        /*
         * Note messages posted to the thread are not (and cannot be) 
         * dispatched through DispatchMessage. These are processed
         * by us directly.
         */
        if (msg.hwnd == NULL) {
            /* Message posted to our thread */
            switch (msg.message) {
            case WM_QUIT:
                /* Note - we do not expect to receive this until DLL is 
                   being unloaded or process exiting */
                /* TBD - what do we need to free ? */
                TwapiDeviceNotificationTid = 0;
                return 0;
            case TWAPI_WM_ADD_DEVICE_NOTIFICATION:
                dncP = (TwapiDeviceNotificationContext*) msg.wParam;
                TwapiCreateDeviceNotificationWindow(dncP);
                /*
                 * Window holds a ref. We unref here to match the ref in
                 * the initiating interp. All management of the dncP
                 * is now in the window proc.
                 */
                TwapiDeviceNotificationContextUnref(dncP, 1);
                break;

            case TWAPI_WM_REMOVE_DEVICE_NOTIFICATION:
                /*
                 * Find the appropriate window and destroy it. Once again,
                 * note no need for locking as all access through single
                 * thread.
                 */
                dncP = TwapiFindDeviceNotificationById((int)msg.wParam);
                if (dncP && dncP->hwnd)
                    DestroyWindow(dncP->hwnd);
                break;
            default:
                break;          /* Ignore */
            }
        } else {
            /* Most likely device notification messages */
            TranslateMessage(&msg);
            DispatchMessage(&msg); 
        }
    } // End of PeekMessage while loop.


    return 0;
}


static int TwapiDeviceModuleInit(Tcl_Interp *interp)
{
    /*
     * We have to create the device notification thread. Moreover, we
     * will have to wait for it to run and start processing the message loop 
     * else our first request might be lost.
     */
    HANDLE threadH;
    HANDLE sig;
    int status = TCL_ERROR;
    DWORD winerr;

    
    sig = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (sig) {
        /* TBD - when does the thread get asked to exit? */
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
        threadH = CreateThread(NULL, 0, TwapiDeviceNotificationThread, sig, 0,
                               &TwapiDeviceNotificationTid);
#else
        threadH = (HANDLE)  _beginthreadex(NULL, 0,
                                           TwapiDeviceNotificationThread,
                                           sig, 0, &TwapiDeviceNotificationTid);
#endif
        if (threadH) {
            CloseHandle(threadH);
            /* Wait for the thread to get running and sit in its message loop */
            if (WaitForSingleObject(sig, 5000) == WAIT_OBJECT_0)
                status = TCL_OK;
            else
                winerr = GetLastError();
        }
        CloseHandle(sig);
    }


    if (status != TCL_OK)
        Twapi_AppendSystemError(interp, winerr);

    return status;
}


int Twapi_RegisterDeviceNotification(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    TwapiDeviceNotificationContext *dncP;
    GUID *guidP;
    TwapiId id;

    /* Device notification threaded guaranteed running */
    dncP = TwapiDeviceNotificationContextNew(ticP); /* Always non-NULL return */

    guidP = &dncP->device.guid;
    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETINT(dncP->devtype),
                     GETVAR(guidP, ObjToGUID_NULL),
                     GETHANDLE(dncP->device.hdev),
                     ARGEND) == TCL_ERROR) {
        TwapiDeviceNotificationContextDelete(dncP);
        return TCL_ERROR;
    }

     /* We need separate variable id because we do not want to access
      * dncP->id after queueing
      */
    id = TWAPI_NEWID(ticP);
    dncP->id = id;

    /* guidP == NULL -> empty string - use a NULL guid */
    if (guidP == NULL)
        dncP->device.guid = gNullGuid;

    TwapiDeviceNotificationContextRef(dncP, 1);
    if (PostThreadMessageW(TwapiDeviceNotificationTid, TWAPI_WM_ADD_DEVICE_NOTIFICATION, (WPARAM) dncP, 0)) {
        ObjSetResult(ticP->interp, ObjFromTwapiId(id));
        return TCL_OK;
    } else {
        DWORD winerr = GetLastError();
        TwapiDeviceNotificationContextUnref(dncP, 1);
        return Twapi_AppendSystemError(ticP->interp, winerr);
    }
}

int Twapi_UnregisterDeviceNotification(TwapiInterpContext *ticP, TwapiId id)
{
    if (PostThreadMessageW(TwapiDeviceNotificationTid,
                           TWAPI_WM_REMOVE_DEVICE_NOTIFICATION,
                           (WPARAM) id,
                           0))
        return TCL_OK;
    else {
        DWORD winerr = GetLastError();
        return Twapi_AppendSystemError(ticP->interp, winerr);
    }
}

TwapiDeviceNotificationContext *TwapiDeviceNotificationContextNew(
    TwapiInterpContext *ticP)
{
    TwapiDeviceNotificationContext *dncP;

    dncP = (TwapiDeviceNotificationContext *) TwapiAlloc(sizeof(*dncP));
    dncP->ticP = ticP;
    if (ticP)
        TwapiInterpContextRef(ticP, 1);
    dncP->id = 0;
    dncP->devtype = DBT_DEVTYP_HANDLE;
    dncP->device.hdev = NULL;
    dncP->nrefs = 0;
     
    ZLINK_INIT(dncP);
    dncP->hwnd = NULL;
    dncP->hnotification = NULL;

    return dncP;
}



/* Delete a notification context. Internal resources freed up but it
 * must NOT be on the notification registry list or attached to a window.
 */
void TwapiDeviceNotificationContextDelete(TwapiDeviceNotificationContext *dncP)
{
    /* No locking needed as dncP is only accessed from one thread at a time */
    if (dncP->hnotification)
        UnregisterDeviceNotification(dncP->hnotification);
    dncP->hnotification = NULL;
    if (dncP->ticP) {
        TwapiInterpContextUnref(dncP->ticP, 1);
        dncP->ticP = NULL;
    }
}

void TwapiDeviceNotificationContextUnref(TwapiDeviceNotificationContext *dncP, int decr)
{
    /* No locks necessary as dncP is only manipulated from a single
     * thread at atime once it is allocated
     */
    if (dncP->nrefs > decr) {
        dncP->nrefs -= decr;
    } else {
        dncP->nrefs = 0;
        TwapiDeviceNotificationContextDelete(dncP);
    }
}

static void TwapiReportDeviceNotificationError(
    TwapiDeviceNotificationContext *dncP,
    char *msg,
    DWORD winerr
    )
{
//TBD
}


static Tcl_Obj *ObjFromDevtype(DWORD devtype)
{
    const char *cP;

    switch (devtype) {
    case DBT_DEVTYP_DEVICEINTERFACE:
        cP = "deviceinterface";
        break;
    case DBT_DEVTYP_HANDLE:
        cP = "handle";
        break;
    case DBT_DEVTYP_VOLUME:
        cP = "volume";
        break;
    case DBT_DEVTYP_OEM:
        cP = "oem";
        break;
    case DBT_DEVTYP_PORT:
        cP = "port";
        break;
    default:
        cP = "unknown";
        break;
    }
    return ObjFromString(cP);
}


static Tcl_Obj *ObjFromCustomDeviceNotification(PDEV_BROADCAST_HDR  dbhP)
{
    char *custom_str;
    PDEV_BROADCAST_HANDLE dhP = (PDEV_BROADCAST_HANDLE)dbhP;
    Tcl_Obj *objs[3];

    /*
     * Do not know how to parse anything else other than HANDLE. Also
     * verify minimum size.
     */
    if (dbhP->dbch_devicetype != DBT_DEVTYP_HANDLE ||
        dbhP->dbch_size < offsetof(DEV_BROADCAST_HANDLE, dbch_nameoffset))
        return ObjFromEmptyString();
    
    if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_CHANGE))
        custom_str = "io_volume_change";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_DISMOUNT))
        custom_str = "io_volume_dismount";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_DISMOUNT_FAILED))
        custom_str = "io_volume_dismount_failed";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_MOUNT))
        custom_str = "io_volume_mount";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_LOCK))
        custom_str = "io_volume_lock";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_LOCK_FAILED))
        custom_str = "io_volume_lock_failed";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_UNLOCK))
        custom_str = "io_volume_unlock";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_NAME_CHANGE))
        custom_str = "io_volume_name_change";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_PHYSICAL_CONFIGURATION_CHANGE))
        custom_str = "io_volume_physical_configuration_change";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_VOLUME_DEVICE_INTERFACE))
        custom_str = "io_volume_device_interface";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_MEDIA_ARRIVAL))
        custom_str = "io_media_arrival";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_MEDIA_REMOVAL))
        custom_str = "io_media_removal";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_DEVICE_BECOMING_READY))
        custom_str = "io_device_becoming_ready";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_DEVICE_EXTERNAL_REQUEST))
        custom_str = "io_device_external_request";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_MEDIA_EJECT_REQUEST))
        custom_str = "io_media_eject_request";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_DRIVE_REQUIRES_CLEANING))
        custom_str = "io_drive_requires_cleaning";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_TAPE_ERASE))
        custom_str = "io_tape_erase";
    else if (IsEqualGUID(&dhP->dbch_eventguid, &GUID_IO_DISK_LAYOUT_CHANGE))
        custom_str = "io_disk_layout_change";
    else
        custom_str = "unknown_guid";

    objs[0] = ObjFromString(custom_str);
    objs[1] = ObjFromGUID(&dhP->dbch_eventguid);
    if (dhP->dbch_size > sizeof(DEV_BROADCAST_HANDLE) &&
        dhP->dbch_nameoffset > 0) {
        objs[2] = ObjFromString(dhP->dbch_nameoffset + (char*)dhP);
    } else
        objs[2] = ObjFromEmptyString();
    
    return ObjNewList(3, objs);
}


/* Find the window handle corresponding to a device notification id */
static TwapiDeviceNotificationContext *TwapiFindDeviceNotificationById(DWORD_PTR id)
{
    TwapiDeviceNotificationContext *dncP;

    /* No need for locking, only accessed from a single thread */
    ZLIST_LOCATE(dncP, &gTwapiDeviceNotificationRegistry, id, id);
    return dncP;
}

/* Find the window handle corresponding to a device notification id */
static TwapiDeviceNotificationContext *TwapiFindDeviceNotificationByHwnd(HWND hwnd)
{
    TwapiDeviceNotificationContext *dncP;

    /* No need for locking, only accessed from a single thread */
    ZLIST_LOCATE(dncP, &gTwapiDeviceNotificationRegistry, hwnd, hwnd);
    return dncP;
}

static int Twapi_DeviceIoControlObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE   hdev;
    Tcl_Obj *inputObj;
    DWORD    ctrl;
    void    *inP, *outP;
    DWORD    nin, nout;
    MemLifo *memlifoP = ticP->memlifoP;
    MemLifoMarkHandle mark;
    TCL_RESULT res;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHANDLE(hdev), GETINT(ctrl),
                     ARGUSEDEFAULT, GETOBJ(inputObj),
                     GETINT(nout),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    mark = MemLifoPushMark(memlifoP);
    /* NOTE mark HAS TO BE POPPED ON ALL EXITS */

    if (inputObj) {
        res = TwapiCStructParse(interp, memlifoP, inputObj, CSTRUCT_ALLOW_NULL, &nin, &inP);
    } else {
        inP = NULL;
        nin = 0; 
        res = TCL_OK;
    }

    if (res == TCL_OK) {
        if (nout)
            outP = MemLifoAlloc(memlifoP, nout, &nout);
        else
            outP = NULL;

        /* We currently to do not bother with failing due to buffer being
           too small. Caller should handle that case
           (ERROR_INSUFFICIENT_BUFFER or ERROR_MORE_DATA) */
        if (DeviceIoControl(hdev, ctrl, inP, nin, outP, nout, &nout, NULL)) {
            if (outP && nout) {
                ObjSetResult(interp, ObjFromByteArray(outP, nout));
            }
            res = TCL_OK;
        } else
            res = TwapiReturnSystemError(interp);
    }

    MemLifoPopMark(mark);
    return res;
}

static int Twapi_DeviceCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    GUID guid;
    GUID *guidP;
    HWND   hwnd;
    LPVOID pv;
    Tcl_Obj *sObj, *s2Obj;
    DWORD dw, dw2;
    CONFIGRET cret;
    union {
        WCHAR buf[MAX_PATH+1];
        struct {
            SP_DEVINFO_DATA sp_devinfo_data;
            SP_DEVINFO_DATA *sp_devinfo_dataP;
            SP_DEVICE_INTERFACE_DATA sp_device_interface_data;
        } dev;
    } u;
    HANDLE h;
    TwapiResult result;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 56: // CM_Reenumerate_DevNode_Ex
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), GETINT(dw2),
                         ARGUSEDEFAULT, GETHANDLET(h, HMACHINE),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_CONFIGRET;
        result.value.ival = CM_Reenumerate_DevNode_Ex(dw, dw2, h);
        break;
        
    case 57: // CM_Locate_DevNode_Ex
        if (TwapiGetArgs(interp, objc-2, objv+2, ARGSKIP, GETINT(dw),
                         ARGUSEDEFAULT,
                         GETHANDLET(h, HMACHINE),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        cret = CM_Locate_DevNode_ExW(&result.value.ival,
                                        ObjToUnicode(objv[2]), dw, h);
        if (cret == CR_SUCCESS)
            result.type = TRT_DWORD;
        else {
            result.type = TRT_CONFIGRET;
            result.value.ival = cret;
        }
        break;
        
    case 58: // CM_Disconnect_Machine
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLET(h, HMACHINE),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_CONFIGRET;
        result.value.ival = CM_Disconnect_Machine(h);
        break;

    case 59: // CM_Connect_Machine
        CHECK_NARGS(interp, objc, 3);
        cret = CM_Connect_MachineW(ObjToLPWSTR_NULL_IF_EMPTY(objv[2]), &result.value.hval);
        
        if (cret == CR_SUCCESS)
            result.type = TRT_HMACHINE;
        else {
            result.type = TRT_CONFIGRET;
            result.value.ival = cret;
        }
        break;

    case 60: // SetupDiCreateDeviceInfoListExW
        guidP = &guid;
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETVAR(guidP, ObjToGUID_NULL),
                         GETHWND(hwnd),
                         GETOBJ(sObj),
                         GETVOIDP(pv),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDEVINFO;
        result.value.hval = SetupDiCreateDeviceInfoListExW(
            guidP, hwnd, ObjToLPWSTR_NULL_IF_EMPTY(sObj), pv);
        break;

    case 61:
        guidP = &guid;
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETVAR(guidP, ObjToGUID_NULL),
                         GETOBJ(sObj),
                         GETHWND(hwnd),
                         GETINT(dw),
                         GETHANDLET(h, HDEVINFO),
                         GETOBJ(s2Obj),
                         ARGUSEDEFAULT,
                         GETVOIDP(pv),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_HDEVINFO;
        result.value.hval = SetupDiGetClassDevsExW(
            guidP, ObjToLPWSTR_NULL_IF_EMPTY(sObj), hwnd, dw, h,
            ObjToLPWSTR_NULL_IF_EMPTY(s2Obj), pv);
        break;
    case 62: // SetupDiEnumDeviceInfo
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLET(h, HDEVINFO),
                         GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        u.dev.sp_devinfo_data.cbSize = sizeof(u.dev.sp_devinfo_data);
        if (SetupDiEnumDeviceInfo(h, dw, &u.dev.sp_devinfo_data)) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromSP_DEVINFO_DATA(&u.dev.sp_devinfo_data);
        } else
            result.type = TRT_GETLASTERROR_SETUPAPI;
        break;

    case 63: // Twapi_SetupDiGetDeviceRegistryProperty
        return Twapi_SetupDiGetDeviceRegistryProperty(ticP, objc-2, objv+2);
    case 64: // SetupDiEnumDeviceInterfaces
        u.dev.sp_devinfo_dataP = & u.dev.sp_devinfo_data;
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLET(h, HDEVINFO),
                         GETVAR(u.dev.sp_devinfo_dataP, ObjToSP_DEVINFO_DATA_NULL),
                         GETGUID(guid),
                         GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        
        u.dev.sp_device_interface_data.cbSize = sizeof(u.dev.sp_device_interface_data);
        if (SetupDiEnumDeviceInterfaces(
                h, u.dev.sp_devinfo_dataP,  &guid,
                dw, &u.dev.sp_device_interface_data)) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromSP_DEVICE_INTERFACE_DATA(&u.dev.sp_device_interface_data);
        } else
            result.type = TRT_GETLASTERROR_SETUPAPI;
        break;

    case 65:
        return Twapi_SetupDiGetDeviceInterfaceDetail(ticP, objc-2, objv+2);
    case 66:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETGUID(guid),
                         ARGUSEDEFAULT,
                         GETOBJ(sObj),
                         GETVOIDP(pv),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (SetupDiClassNameFromGuidExW(
                &guid, u.buf, ARRAYSIZE(u.buf),
                NULL, ObjToLPWSTR_NULL_IF_EMPTY(sObj), pv)) {
            result.type = TRT_UNICODE;
            result.value.unicode.str = u.buf;
            result.value.unicode.len = -1;
        } else
            result.type = TRT_GETLASTERROR_SETUPAPI;
        break;
    case 67:
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLET(h, HDEVINFO),
                         GETVAR(u.dev.sp_devinfo_data, ObjToSP_DEVINFO_DATA),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (SetupDiGetDeviceInstanceIdW(h, &u.dev.sp_devinfo_data,
                                        u.buf, ARRAYSIZE(u.buf), NULL)) {
            result.type = TRT_UNICODE;
            result.value.unicode.str = u.buf;
            result.value.unicode.len = -1;
        } else
            result.type = TRT_GETLASTERROR_SETUPAPI;
        break;
    case 68:
        return Twapi_SetupDiClassGuidsFromNameEx(ticP, objc-2, objv+2);
    case 69: // SetupDiOpenDeviceInfo
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLET(h, HDEVINFO), GETOBJ(sObj),
                         GETHWND(hwnd), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        u.dev.sp_devinfo_data.cbSize = sizeof(u.dev.sp_devinfo_data);
        if (SetupDiOpenDeviceInfoW(h, ObjToUnicode(sObj), hwnd, dw,
                                   &u.dev.sp_devinfo_data)) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromSP_DEVINFO_DATA(&u.dev.sp_devinfo_data);
        } else
            result.type = TRT_GETLASTERROR_SETUPAPI;
        break;
    case 70:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToHANDLE(interp, objv[2], &h) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = SetupDiDestroyDeviceInfoList(h);
        break;
    case 71:
        return Twapi_RegisterDeviceNotification(ticP, objc-2, objv+2);
    case 72:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        return Twapi_UnregisterDeviceNotification(ticP, dw);
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiDeviceInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Dispatch function codes as below for historical reasons.
       Nothing magic but change Twapi_DeviceCallObjCmd accordingly 
    */
    static struct alias_dispatch_s DeviceAliasDispatch[] = {
        DEFINE_ALIAS_CMD(CM_Reenumerate_DevNode_Ex, 56),
        DEFINE_ALIAS_CMD(CM_Locate_DevNode_Ex, 57),
        DEFINE_ALIAS_CMD(CM_Disconnect_Machine, 58),
        DEFINE_ALIAS_CMD(CM_Connect_Machine, 59),
        DEFINE_ALIAS_CMD(SetupDiCreateDeviceInfoListEx, 60),
        DEFINE_ALIAS_CMD(SetupDiGetClassDevsEx, 61),
        DEFINE_ALIAS_CMD(SetupDiEnumDeviceInfo, 62),
        DEFINE_ALIAS_CMD(Twapi_SetupDiGetDeviceRegistryProperty, 63),
        DEFINE_ALIAS_CMD(SetupDiEnumDeviceInterfaces, 64),
        DEFINE_ALIAS_CMD(SetupDiGetDeviceInterfaceDetail, 65),
        DEFINE_ALIAS_CMD(device_setup_class_guid_to_name, 66), //SetupDiClassNameFromGuidEx
        DEFINE_ALIAS_CMD(device_element_instance_id, 67), //SetupDiGetDeviceInstanceId
        DEFINE_ALIAS_CMD(SetupDiClassGuidsFromNameEx, 68),
        DEFINE_ALIAS_CMD(SetupDiOpenDeviceInfo, 69),
        DEFINE_ALIAS_CMD(devinfoset_close, 70), //SetupDiDestroyDeviceInfoList
        DEFINE_ALIAS_CMD(Twapi_RegisterDeviceNotification, 71),
        DEFINE_ALIAS_CMD(Twapi_UnregisterDeviceNotification, 72),
    };

    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::DeviceCall", Twapi_DeviceCallObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::DeviceIoControl", Twapi_DeviceIoControlObjCmd, ticP, NULL);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(DeviceAliasDispatch), DeviceAliasDispatch, "twapi::DeviceCall");

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
int Twapi_device_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiDeviceInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    ERROR_IF_UNTHREADED(interp);
    
    if (! TwapiDoOneTimeInit(&TwapiDeviceModuleInitialized,
                             TwapiDeviceModuleInit, interp))
        return TCL_ERROR;

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

