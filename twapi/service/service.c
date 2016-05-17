/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#if !defined(TWAPI_REPLACE_CRT) && !defined(TWAPI_MINIMIZE_CRT)
# include <process.h>
#endif
#include <dbt.h>                /* For DBT_* decls */

#include "twapi_service.h"

/* There can be only one....*/
#if __GNUC__
static void * volatile gServiceInterpContextP;
#else
static TwapiInterpContext * volatile gServiceInterpContextP;
#endif

/* Structure to keep track of a particular service state */
typedef struct _TwapiServiceContext {
    SERVICE_STATUS_HANDLE service_status_handle; /* Must NEVER be closed! */
    DWORD   controls_accepted;   /* Controls accepted by the service */
    WCHAR   name[1];               /* Variable size, name of service */
} TwapiServiceContext;

/* 
 * Size of above variable struct to hold a name of length namelen_. Note
 * name[1] and the required null terminator "cancel" each other when
 * calculating the length.
 */
#define SIZE_TwapiServiceContext(namelen_) \
    (sizeof(TwapiServiceContext) + (sizeof(WCHAR)*namelen_))

/*
 * Service control events delivered via the Tcl event queue.
 *
 * NOTE THIS IS A VARIABLE SIZE STRUCTURE AS DEVICE_DEPENDENT DATA IS
 * TACKED ON AT THE END.
 */
typedef struct _TwapiServiceControlCallback {
    TwapiCallback  cb;   /* Must be first field */
    int   service_index;        /* Index into global service table */
    DWORD ctrl;                 /* The service control received */
    DWORD event;                /* Control-dependent code */
    DWORD additional_info;      /* Additional control-dependent data */
} TwapiServiceControlCallback;


/*
 * Global data for services. These is intialized by the first call by an
 * interp to become a service (only one such call is allowed per process).
 * Thereafter they are not modified until process exit.
 * Each individual element of gServiceContexts[] is only modified by
 * the specific thread running that service.
 * Hence no locks are required for this data.
 */
static TwapiServiceContext **gServiceContexts;
static int gNumServiceContexts;
DWORD  gServiceType;        /* own process or shared and whether interactive */
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
/* Using CreateThread */
DWORD  gServiceMasterThreadId;   /* Thread of service main program */
#else
/* Using _beginthreadex */
unsigned int gServiceMasterThreadId;   /* Thread of service main program */
#endif
HANDLE gServiceMasterThreadHandle;


static void TwapiFreeServiceContexts(void);
static BOOL ConsoleCtrlHandler( DWORD ctrl);
static unsigned int WINAPI TwapiServiceMasterThread(LPVOID unused);
static int TwapiFindServiceIndex(const WCHAR *nameP);

int Twapi_SetServiceStatus(
    TwapiInterpContext *ticP,
    int objc,
    Tcl_Obj *CONST objv[])
{
    SERVICE_STATUS ss;
    int service_index;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     ARGSKIP, GETINT(ss.dwCurrentState),
                     GETINT(ss.dwWin32ExitCode),
                     GETINT(ss.dwServiceSpecificExitCode),
                     GETINT(ss.dwCheckPoint),
                     GETINT(ss.dwWaitHint), GETINT(ss.dwControlsAccepted),
                     ARGEND
            )
        != TCL_OK)
        return TCL_ERROR;

    service_index = TwapiFindServiceIndex(ObjToWinChars(objv[0]));
    if (service_index < 0)
        return Twapi_AppendSystemError(ticP->interp, ERROR_INVALID_NAME);

    ss.dwServiceType  = gServiceType;

    if (SetServiceStatus(gServiceContexts[service_index]->service_status_handle, &ss))
        return TCL_OK;
    else
        return TwapiReturnSystemError(ticP->interp);
}

int Twapi_BecomeAService(
    TwapiInterpContext *ticP,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Interp *interp;
    DWORD       service_type;
    int i;

    interp = ticP->interp;
    ERROR_IF_UNTHREADED(interp);

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (ObjToDWORD(interp, objv[0], &service_type) != TCL_OK)
        return TCL_ERROR;

    if (InterlockedCompareExchangePointer(&gServiceInterpContextP, ticP, NULL)
        != NULL) {
        ObjSetStaticResult(interp, "Twapi_BecomeAService multiple invocations.");
        return TCL_ERROR;
    }
    TwapiInterpContextRef(ticP, 1);

    /* Remaining arguments are service definitions. Allocate and parse */
    gNumServiceContexts = objc-1;
    gServiceContexts = TwapiAllocZero(gNumServiceContexts*sizeof(TwapiServiceContext));
    for (i = 0; i < gNumServiceContexts; ++i) {
        Tcl_Obj **objs;
        int       n;
        int       ctrls;
        WCHAR    *nameP;
        if (ObjGetElements(interp, objv[i+1], &n, &objs) != TCL_OK ||
            n != 2 ||
            ObjToInt(interp, objs[1], &ctrls) != TCL_OK) {
            ObjSetStaticResult(interp, "Invalid service specification.");
            goto error_handler;
        }
        /* Allocate a single block for the context and name */
        nameP = ObjToWinCharsN(objs[0], &n);
        gServiceContexts[i] = TwapiAlloc(SIZE_TwapiServiceContext(n));
        CopyMemory(gServiceContexts[i]->name, nameP, sizeof(WCHAR)*(n+1));
        gServiceContexts[i]->controls_accepted = ctrls;
        gServiceContexts[i]->service_status_handle = NULL;
    }

    /*
     * Set the console control handler to ignore signals else we will
     * exit when the user logs off. This happens only if some other
     * DLL or runtime has replaced the default handler.
     * See http://support.microsoft.com/kb/149901
     * Note: originally we were doing this in the Tcl code.
     */
    if (! SetConsoleCtrlHandler( (PHANDLER_ROUTINE) ConsoleCtrlHandler, TRUE )) {
        ObjSetStaticResult(interp, "Console control handler install failed.");
        return TCL_ERROR;
    }

    gServiceType = service_type & (SERVICE_WIN32 | SERVICE_INTERACTIVE_PROCESS);

    /* Now start the thread that will actually talk to the control manager */
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
    gServiceMasterThreadHandle = CreateThread(NULL, 0, TwapiServiceMasterThread,
                                        NULL, 0, &gServiceMasterThreadId);
#else
    gServiceMasterThreadHandle = (HANDLE)  _beginthreadex(NULL, 0,
                                                          TwapiServiceMasterThread, NULL,
                                                          0, &gServiceMasterThreadId);
#endif
    if (gServiceMasterThreadHandle == NULL) {
        Twapi_AppendSystemError(interp, GetLastError());
        return TCL_ERROR;
    }

    return TCL_OK;

error_handler:
    TwapiFreeServiceContexts();
    if (InterlockedCompareExchangePointer(&gServiceInterpContextP, NULL, ticP) == ticP)
        TwapiInterpContextUnref(ticP, 1);
    return TCL_ERROR;
}

static int TwapiServiceControlCallbackFn(TwapiCallback *p)
{
    TwapiServiceControlCallback *cbP = (TwapiServiceControlCallback *)p;
    char *ctrl_str = NULL;
    char *event_str = NULL;

    if (cbP->cb.ticP->interp == NULL ||
        Tcl_InterpDeleted(cbP->cb.ticP->interp)) {
        cbP->cb.winerr = ERROR_INVALID_STATE; /* Best match we can find */
        cbP->cb.response.type = TRT_EMPTY;
        return TCL_ERROR;
    }

    if (cbP->service_index >= gNumServiceContexts) {
        cbP->cb.winerr = ERROR_INVALID_PARAMETER;
        return TCL_ERROR;
    }
    switch (cbP->ctrl) {
    case 0:
        ctrl_str = "start";
        break;
    case SERVICE_CONTROL_STOP:
        ctrl_str = "stop";
        break;
    case SERVICE_CONTROL_PAUSE:
        ctrl_str = "pause";
        break;
    case SERVICE_CONTROL_CONTINUE:
        ctrl_str = "continue";
        break;
    case SERVICE_CONTROL_INTERROGATE:
        ctrl_str = "interrogate";
        break;
    case SERVICE_CONTROL_SHUTDOWN:
        ctrl_str = "shutdown";
        break;
    case SERVICE_CONTROL_PARAMCHANGE:
        ctrl_str = "paramchange";
        break;
    case SERVICE_CONTROL_NETBINDADD:
        ctrl_str = "netbindadd";
        break;
    case SERVICE_CONTROL_NETBINDREMOVE:
        ctrl_str = "netbindremove";
        break;
    case SERVICE_CONTROL_NETBINDENABLE:
        ctrl_str = "netbindenable";
        break;
    case SERVICE_CONTROL_NETBINDDISABLE:
        ctrl_str = "netbinddisable";
        break;
    case SERVICE_CONTROL_HARDWAREPROFILECHANGE:
        ctrl_str = "hardwareprofilechange";
        switch (cbP->event) {
        case DBT_CONFIGCHANGECANCELED:
            event_str = "configchangecanceled";
            break;
        case DBT_CONFIGCHANGED:
            event_str = "configchanged";
            break;
        case DBT_QUERYCHANGECONFIG:
            event_str = "querychangeconfig";
            break;
        }
        break;

    case SERVICE_CONTROL_POWEREVENT:
        ctrl_str = "powerevent";
        switch (cbP->event) {
        case PBT_APMSUSPEND:
            event_str = "apmsuspend";
            break;
        case PBT_APMSTANDBY:
            event_str = "apmstandby";
            break;
        case PBT_APMRESUMECRITICAL:
            event_str = "apmresumecritical";
            break;
        case PBT_APMRESUMESUSPEND:
            event_str = "apmresumesuspend";
            break;
        case PBT_APMRESUMESTANDBY:
            event_str = "apmresumestandby";
            break;
        case PBT_APMBATTERYLOW:
            event_str = "apmbatterylow";
            break;
        case PBT_APMPOWERSTATUSCHANGE:
            event_str = "apmpowerstatuschange";
            break;
        case PBT_APMOEMEVENT:
            event_str = "apmoemevent";
            break;
        case PBT_APMRESUMEAUTOMATIC:
            event_str = "apmresumeautomatic";
            break;
        case PBT_APMQUERYSUSPEND:
            event_str = "apmquerysuspend";
            break;
        case PBT_APMQUERYSTANDBY:
            event_str = "apmquerystandby";
            break;
        case PBT_APMQUERYSUSPENDFAILED:
            event_str = "apmquerysuspendfailed";
            break;
        case PBT_APMQUERYSTANDBYFAILED:
            event_str = "apmquerystandbyfailed";
            break;
        default:
            break;
        }

        break;

    case SERVICE_CONTROL_SESSIONCHANGE:
        ctrl_str = "sessionchange";
        /* SDK with _WINNT_VER == win2k does not have defines for these */
        switch (cbP->event) {
        case 1:
            event_str = "console_connect";
            break;
        case 2:
            event_str = "console_disconnect";
            break;
        case 3:
            event_str = "remote_connect";
            break;
        case 4:
            event_str = "remote_disconnect";
            break;
        case 5:
            event_str = "session_login";
            break;
        case 6:
            event_str = "session_logoff";
            break;
        case 7:
            event_str = "session_lock";
            break;
        case 8:
            event_str = "session_unlock";
            break;
        case 9:
            event_str = "session_remote_control";
            break;
        }
        break;

    default:
        if (cbP->ctrl >= 128 && cbP->ctrl <= 255) {
            // User defined
            ctrl_str = "userdefined";
            break;
        }
    }

    if (ctrl_str) {
        Tcl_Obj *objs[7];
        int nobjs = 0;
        objs[0] = STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_service_handler");
        objs[1] = ObjFromWinChars(gServiceContexts[cbP->service_index]->name);
        objs[2] = ObjFromOpaque(gServiceContexts[cbP->service_index]->service_status_handle, "SERVICE_STATUS_HANDLE");
        objs[3] = Tcl_NewStringObj(ctrl_str, -1);
        objs[4] = Tcl_NewLongObj(cbP->ctrl);
        nobjs = 5;
        if (event_str) {
            objs[nobjs++] = Tcl_NewStringObj(event_str, -1);
            if (cbP->ctrl == SERVICE_CONTROL_SESSIONCHANGE)
                objs[nobjs++] = Tcl_NewLongObj(cbP->additional_info);
        }
        return TwapiEvalAndUpdateCallback(&cbP->cb, nobjs, objs, TRT_LONG);
    } else {
        /* Unknown event - ignore TBD */
        cbP->cb.winerr = ERROR_INVALID_FUNCTION;
        return TCL_ERROR;
    }
}

/*
 * This is called by the SCM in the same thread as ServiceMasterThread
 * We need to pass on the event to the Tcl thread and return the response.
 */
static DWORD WINAPI TwapiServiceControlHandler(DWORD ctrl, DWORD event, PVOID event_dataP, PVOID context) {
    TwapiServiceControlCallback *cbP;
    int service_index = PtrToInt(context);
    int need_response=0;
    DWORD additional_info = 0;
    DWORD status;

    if (service_index >= gNumServiceContexts) {
        TwapiWriteEventLogError("TwapiServiceControlHandler: service_index is greater than number of service contexts. Ignoring.");
        return NO_ERROR; /* Should not happen? */
    }

    switch (ctrl) {
    case SERVICE_CONTROL_STOP:
    case SERVICE_CONTROL_PAUSE:
    case SERVICE_CONTROL_CONTINUE:
    case SERVICE_CONTROL_INTERROGATE:
    case SERVICE_CONTROL_SHUTDOWN:
    case SERVICE_CONTROL_PARAMCHANGE:
    case SERVICE_CONTROL_NETBINDADD:
    case SERVICE_CONTROL_NETBINDREMOVE:
    case SERVICE_CONTROL_NETBINDENABLE:
    case SERVICE_CONTROL_NETBINDDISABLE:
        need_response = 0;
        break;
    case SERVICE_CONTROL_HARDWAREPROFILECHANGE:
        switch (event) {
        case DBT_CONFIGCHANGECANCELED:
        case DBT_CONFIGCHANGED:
            need_response = 0;
            break;
        case DBT_QUERYCHANGECONFIG:
            need_response = 1;
            break;
        default:
            return NO_ERROR;
        }
        break;
    case SERVICE_CONTROL_POWEREVENT:
        switch (event) {
        case PBT_APMSUSPEND:
        case PBT_APMSTANDBY:
        case PBT_APMRESUMECRITICAL:
        case PBT_APMRESUMESUSPEND:
        case PBT_APMRESUMESTANDBY:
        case PBT_APMBATTERYLOW:
        case PBT_APMPOWERSTATUSCHANGE:
        case PBT_APMOEMEVENT:
        case PBT_APMRESUMEAUTOMATIC:
            need_response = 0;
            break;

        case PBT_APMQUERYSUSPEND:
        case PBT_APMQUERYSTANDBY:
        case PBT_APMQUERYSUSPENDFAILED:
        case PBT_APMQUERYSTANDBYFAILED:
            need_response = 1;
            break;
            
        default:
            /* We do not handles these */
            return NO_ERROR;
        }

        break;

    case SERVICE_CONTROL_SESSIONCHANGE:
        /* 
         * WTSSESSION_NOTIFICATION is not defined in the header because
         * it requires building against XP, Win2K is not enough. So
         * So we pick out the words explicitly. Second DWORD of the struct
         * is session id (WTSSESSION_NOTIFICATION.dwSessionId)
         */
        additional_info = ((DWORD *) event_dataP)[1];
        break;

    case SERVICE_CONTROL_DEVICEEVENT:
        return NO_ERROR;        /* Should not get these, but just in case */
    default:
        if (ctrl >= 128 && ctrl <= 255)
            break;
        return ERROR_CALL_NOT_IMPLEMENTED;
    }

    cbP = (TwapiServiceControlCallback *)
        TwapiCallbackNew(gServiceInterpContextP,
                                TwapiServiceControlCallbackFn,
                                sizeof(*cbP));

    /* Note unless the interp thread successfully responds with
       an status code, we assume status is NO_ERROR so as to not block
       notifications that need a response */
    status = NO_ERROR;
    cbP->service_index = service_index;
    cbP->ctrl = ctrl;
    cbP->event = event;
    cbP->additional_info = additional_info;
    if (need_response) {
        if (TwapiEnqueueCallback(gServiceInterpContextP,
                                 &cbP->cb,
                                 TWAPI_ENQUEUE_DIRECT,
                                 30*1000, /* TBD - Timeout (ms) */
                                 (TwapiCallback **)&cbP)
            == ERROR_SUCCESS) {
            if (cbP
                && (cbP->cb.response.type == TRT_DWORD ||
                    cbP->cb.response.type == TRT_LONG))
                status = cbP->cb.response.value.ival;
        }
        if (cbP)
            TwapiCallbackUnref(&cbP->cb, 1);
    } else {
        /* TBD - in call below, on error, do we send an error notification ? */
        TwapiEnqueueCallback(gServiceInterpContextP, &cbP->cb,
                             TWAPI_ENQUEUE_DIRECT,
                             0, /* No response wanted */
                             NULL);
    }
    return status;
}

/*
 * Called by the SCM to start a particular service. It does the intitialization
 * and exits since the actual service work is done by the interp thread.
 */
static void WINAPI TwapiServiceThread(DWORD argc, WCHAR **argv)
{
    int service_index;

    /* If argc or argv are 0, assume we are starting the first service */
    if (argc == 0 || argv == 0 || argv[0] == 0)
        service_index = 0;
    else {
        /* argv[0] is name of the service. Match it to our table */
        service_index = TwapiFindServiceIndex(argv[0]);
    }

    if (service_index == -1 || service_index >= gNumServiceContexts)
        return;           // Did not match any service? Should not happen

    gServiceContexts[service_index]->service_status_handle =
        RegisterServiceCtrlHandlerExW(argv[0],
                                     TwapiServiceControlHandler,
                                     IntToPtr(service_index));
    if (gServiceContexts[service_index]->service_status_handle == 0) {
        TwapiWriteEventLogError("TwapiServiceThread: RegisterServiceCtrlHandlerExW returned NULL handle. Service will not be started.");
    } else {
        SERVICE_STATUS ss;

        ss.dwServiceType  = gServiceType;
        ss.dwCurrentState = SERVICE_START_PENDING;
        ss.dwControlsAccepted = gServiceContexts[service_index]->controls_accepted;
        ss.dwWin32ExitCode = NO_ERROR;
        ss.dwServiceSpecificExitCode = NO_ERROR;
        ss.dwCheckPoint = 1;    // Must not be 0 for START_PENDING! See Richter
        ss.dwWaitHint = 5000;

        if (!SetServiceStatus(gServiceContexts[service_index]->service_status_handle,
                              &ss)) {
            TwapiWriteEventLogError("TwapiServiceThread: SetServiceStatus failed. Service will not be started.");
        } else {
            TwapiServiceControlCallback *cbP;

            cbP = (TwapiServiceControlCallback *)
                TwapiCallbackNew(gServiceInterpContextP,
                                 TwapiServiceControlCallbackFn,
                                 sizeof(*cbP));

            cbP->service_index = service_index;
            cbP->ctrl = 0;      /* 0 -> Start signal */
            cbP->event = 0;
            cbP->additional_info = 0;

            /* TBD - in call below, on error, do we send an error notification ? */
            TwapiEnqueueCallback(gServiceInterpContextP, &cbP->cb,
                                 TWAPI_ENQUEUE_DIRECT,
                                 0, /* No response wanted */
                                 NULL);
        }
    }

    /*
     * Nothing else to do. The SCM will call TwapiServiceControlHandler
     * in the service master thread. That will then pass on controls
     * to the interp thread.
     */

    return;
}

/*
 * This is the master thread for services. In a normal Windows service
 * application, this would be the main thread. Here, it gets started by
 * some interpreter thread. It calls the SCM which calls back into the
 * handlers.
 */
static unsigned int WINAPI TwapiServiceMasterThread(LPVOID unused)
{
    SERVICE_TABLE_ENTRYW *steP;
    int i;

    /*
     * Construct the service descriptors to pass to the SCM. Note
     * the strings within the descriptors point to dynamically allocated
     * memory that will not be deallocated until program termination. So
     * it's safe to point to it.
     */
    steP = TwapiAlloc((1+gNumServiceContexts) * sizeof(steP[0]));    
    for (i = 0; i < gNumServiceContexts; ++i) {
        steP[i].lpServiceName = gServiceContexts[i]->name;
        steP[i].lpServiceProc = TwapiServiceThread;
    }
    steP[i].lpServiceName = NULL;
    steP[i].lpServiceProc = NULL;

    if (! StartServiceCtrlDispatcherW(steP)) {
        // TBD - signal an error but note it might be service shutdown
        // by the time StartServiceCtrlDispatcher returns
    }

    return 0;
}


/* Free resources associated with the global service table */
static void TwapiFreeServiceContexts()
{
    int i;
    if (gServiceContexts) {
        for (i = 0; i < gNumServiceContexts; ++i) {
            // TBD - what to do about the handles. For now we assume they are dead
            TwapiFree(gServiceContexts[i]);
        }
        TwapiFree(gServiceContexts);
        gServiceContexts = NULL;
    }
}

/* Signal handler - Used to "disable" console signals */
static BOOL ConsoleCtrlHandler( DWORD ctrl) 
{ 
    return TRUE;
}

// Given a service name, return its index in the services table or -1
static int TwapiFindServiceIndex(const WCHAR *nameP)
{
    int i;

    if (gServiceContexts) {
        for (i = 0; i < gNumServiceContexts; ++i) {
            if (lstrcmpiW(gServiceContexts[i]->name, nameP) == 0) {
                return i;
            }
        }
    }
    return -1;
}
