/*
 * Copyright (c) 2007-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_wm.h"
#include <process.h>            /* TBD - remove when crt is removed */
#include <dbt.h>
#include <initguid.h>           // Instantiate ioevent.h guids
#include <ioevent.h>            // Custom GUID definitions


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
    int   id;
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
ZLIST_DECL(TwapiDeviceNotificationContext) TwapiDeviceNotificationRegistry;

/*
 * Device notification events delivered via the Tcl event queue.
 *
 * NOTE THIS IS A VARIABLE SIZE STRUCTURE AS DEVICE_DEPENDENT DATA IS
 * TACKED ON AT THE END.
 */
typedef struct _TwapiDeviceNotificationEvent {
    TwapiPendingCallback  cbP;  /* Must be first field */
    int id;                     /* Id of notification context */
    DWORD winerr;               /* 0 or Win32 error */
    union {
        struct {
            WPARAM wparam;              /* wparam from the WM_DEVICECHANGE message */
            DEV_BROADCAST_HDR dev_bcast_hdr; /* VARIABLE SIZED. MUST BE LAST */
        } device;
        char msg[1];            /* Actually variable size */
    } data;
} TwapiDeviceNotificationEvent;

static TwapiOneTimeInitState TwapiDeviceNotificationInitialized;
static DWORD TwapiDeviceNotificationTid;

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
static const char *TwapiDecodeCustomDeviceNotification(PDEV_BROADCAST_HDR);
static const char *TwapiDevtypeToString(DWORD devtype);




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
    TwapiDeviceNotificationEvent *dneP;
    int    need_response = 0;

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
        dneP = (TwapiDeviceNotificationEvent *)
            TwapiPendingCallbackNew(dncP->ticP,
                                    TwapiDeviceNotificationCallbackFn,
                                    sizeof(*dneP) + dbhP->dbch_size - sizeof(dneP->dev_bcast_hdr));
        MoveMemory(&dneP->dev_bcast_hdr, dbhP, dbhP->dbch_size);
    } else {
        /* No additional data */
        dneP = (TwapiDeviceNotificationEvent *)
            TwapiPendingCallbackNew(dncP->ticP,
                                    TwapiDeviceNotificationCallbackFn,
                                    sizeof(*dneP));
        dneP = (TwapiDeviceNotificationEvent *) TwapiAlloc(sizeof(*dneP));
    }
    dneP->id = dncP->id;
    dneP->wparam = wparam;
    dneP->winerr = ERROR_SUCCESS;
    if (need_response) {
        TBD;
    } else {
        /* TBD - on error, do we send an error notification ? */
        TwapiEnqueuCallback(dncP->ticP, &dneP->cpB, TWAPI_ENQUEUE_DIRECT,
                            0, NULL);
        
        return TRUE;
    }
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

        ZeroMemory(&h_filter, sizeof(h_filter));
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

        ZeroMemory(&di_filter, sizeof(di_filter));
        di_filter.dbcc_size = sizeof(di_filter);
        di_filter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
        notify_flags = DEVICE_NOTIFY_WINDOW_HANDLE;
        if (! IsEqualGUID(&dncP->device.guid, &TwapiNullGuid)) {
            // Specific GUID requested
            di_filter.dbcc_classguid = dncP->device.guid;
        }
        else {
            // XP and later allow notification of ALL interfaces
            if (TwapiOSVersionInfo.dwMajorVersion == 5 &&
                TwapiOSVersionInfo.dwMinorVersion == 0) {
                TwapiReportDeviceNotificationError(dncP, "Device interface must be specified on Windows 2000.", ERROR_INVALID_FUNCTION);
                return FALSE;
            }
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
    ZLIST_APPEND(&TwapiDeviceNotificationRegistry, dncP);
    return TCL_OK;

}


/* There is one thread used for device notifications for all threads/interps */
static unsigned __stdcall TwapiDeviceNotificationThread(HANDLE sig)
{
    MSG msg;
    TwapiDeviceNotificationContext *dncP;

    /* Keeps track of all device registrations */
    ZLIST_INIT(&TwapiDeviceNotificationRegistry);      /* TBD - free this registry */

    /*
     * Initialize the message queue so the initiating interp can post
     * messages to us. (See SDK docs)
     */
    PeekMessage(&msg, NULL, WM_USER, WM_USER, PM_NOREMOVE);
    
    /* Now inform the creator that it's free to send us messages */
    SetEvent(sig);

    /* Now start processing messages */
    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) { 
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
                return msg.wParam;
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
                //TBD;
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


static int TwapiDeviceNotificationModuleInit(TwapiInterpContext *ticP)
{
    /*
     * We have to create the device notification thread. Moreover, we
     * will have to wait for it to run and start processing the message loop 
     * else our first request might be lost.
     */
    HANDLE threadH;
    HANDLE sig;
    int status = TCL_ERROR;

    
    sig = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (sig) {
        /* TBD - when does the thread get asked to exit? */
        threadH = (HANDLE)  _beginthreadex(NULL, 0,
                                           TwapiDeviceNotificationThread,
                                           sig, 0, &TwapiDeviceNotificationTid);
        if (threadH) {
            CloseHandle(threadH);
            /* Wait for the thread to get running and sit in its message loop */
            if (WaitForSingleObject(sig, 2000) == WAIT_OBJECT_0)
                status = TCL_OK;
        }
    }

    if (status != TCL_OK)
        TwapiReturnSystemError(ticP->interp);

    if (sig)
        CloseHandle(sig);
    return status;

}

int Twapi_RegisterDeviceNotification(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    TwapiDeviceNotificationContext *dncP;
    DWORD winerr;

    ERROR_IF_UNTHREADED(ticP->interp);
    
    if (! TwapiDoOneTimeInit(&TwapiDeviceNotificationInitialized,
                             TwapiDeviceNotificationModuleInit, NULL))
        return TCL_ERROR;

    /* Device notification threaded guaranteed running */
    

    dncP = TwapiDeviceNotificationContextNew(ticP); /* Always non-NULL return */

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETINT(dncP->id), GETINT(dncP->devtype),
                     GETGUID(dncP->device.guid), GETHANDLE(dncP->device.hdev),
                     ARGEND) == TCL_ERROR) {
        TwapiDeviceNotificationContextDelete(dncP);
        return TCL_ERROR;
    }


    TwapiDeviceNotificationContextRef(dncP, 1);
    if (PostThreadMessageW(TwapiDeviceNotificationTid, TWAPI_WM_ADD_DEVICE_NOTIFICATION, (WPARAM) dncP, NULL))
        return TCL_OK;
    else {
        DWORD winerr = GetLastError();
        TwapiDeviceNotificationContextUnref(dncP, 1);
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
}



void TwapiDeviceNotificationContextDelete(TwapiDeviceNotificationContext *dncP)
{
    /* No locking needed as dncP is only accessed from one thread at a time */
    if (dncP->ticP) {
        dncP->ticP = NULL;
        TwapiInterpContextUnref(dncP->ticP, 1);
    }
    if (dncP->hwnd) {
        DestroyWindow(dncP->hwnd);
        dncP->hwnd = NULL;
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


static const char *TwapiDevtypeToString(DWORD devtype)
{
    switch (devtype) {
    case DBT_DEVTYP_DEVICEINTERFACE:
        return "deviceinterface";
    case DBT_DEVTYP_HANDLE:
        return "handle";
    case DBT_DEVTYP_VOLUME:
        return "volume";
    case DBT_DEVTYP_OEM:
        return "oem";
    case DBT_DEVTYP_PORT:
        return "port";
    default:
        return "unknown";
    }
}


static const char *TwapiDecodeCustomDeviceNotification(PDEV_BROADCAST_HDR  dbhP)
{
    PDEV_BROADCAST_HANDLE dhP = (PDEV_BROADCAST_HANDLE)dbhP;

    /* Do not know how to parse anything else other than HANDLE */
    if (dbhP->dbch_devicetype != DBT_DEVTYP_HANDLE)
        return NULL;
    if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_CHANGE))
        return "io_volume_change";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_DISMOUNT))
        return "io_volume_dismount";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_DISMOUNT_FAILED))
        return "io_volume_dismount_failed";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_MOUNT))
        return "io_volume_mount";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_LOCK))
        return "io_volume_lock";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_LOCK_FAILED))
        return "io_volume_lock_failed";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_UNLOCK))
        return "io_volume_unlock";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_NAME_CHANGE))
        return "io_volume_name_change";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_PHYSICAL_CONFIGURATION_CHANGE))
        return "io_volume_physical_configuration_change";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_VOLUME_DEVICE_INTERFACE))
        return "io_volume_device_interface";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_MEDIA_ARRIVAL))
        return "io_media_arrival";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_MEDIA_REMOVAL))
        return "io_media_removal";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_DEVICE_BECOMING_READY))
        return "io_device_becoming_ready";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_DEVICE_EXTERNAL_REQUEST))
        return "io_device_external_request";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_MEDIA_EJECT_REQUEST))
        return "io_media_eject_request";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_DRIVE_REQUIRES_CLEANING))
        return "io_drive_requires_cleaning";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_TAPE_ERASE))
        return "io_tape_erase";
    else if (IsEqualGUID(dhP->dbch_eventguid, GUID_IO_DISK_LAYOUT_CHANGE))
        return "io_disk_layout_change";
}
