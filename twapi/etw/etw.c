/* 
 * Copyright (c) 2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/*
 * TBD - replace use of string object cache with TwapiGetAtom
 * TBD - replace CALL_ with DEFINE_ALIAS_CMD oR DEFINE_FNCODE_CMD
 * TBD - replace dict ops with list ops if possible
 */

#include "twapi.h"
#include <evntrace.h>

/* Max length of file and session name fields in a trace (documented in SDK) */
#define MAX_TRACE_NAME_CHARS (1024+1)

#define ObjFromTRACEHANDLE(val_) ObjFromWideInt(val_)
#define ObjToTRACEHANDLE ObjToWideInt

/*
 * The Microsoft docs are not clear about invalid trace session handles.
 * The sample code shows comparisons against INVALID_HANDLE_VALUE. On the
 * other hand, the StartTrace documentation warns that invalid session
 * handles are 0, NOT INVALID_HANDLE_VALUE. For more confusion, also see -
 * http://msdn.microsoft.com/en-us/library/windows/desktop/aa364089(v=vs.85).aspx
 * For now, treat both as invalid. TBD
 */
#define INVALID_SESSIONTRACE_HANDLE(ht_) \
    ((ht_) == 0 || (HANDLE) (ht_) == INVALID_HANDLE_VALUE)

/* For efficiency reasons, when constructing many "records" each of which is
 * a keyed list or dictionary, we do not want to recreate the
 * key Tcl_Obj for every record. So we create them once and reuse them.
*/
struct TwapiObjKeyCache {
    const char *keystring;
    Tcl_Obj *keyObj;
};


/* IMPORTANT : 
 * Do not change order without changing ObjToPEVENT_TRACE_PROPERTIES()
 * and ObjFromEVENT_TRACE_PROPERTIES
 */
static const char * g_event_trace_fields[] = {
    "-logfile",
    "-sessionname",
    "-sessionguid",
    "-buffersize",
    "-minbuffers",
    "-maxbuffers",
    "-maximumfilesize",
    "-logfilemode",
    "-flushtimer",
    "-enableflags",
    "-clockresolution",
    "-agelimit",
    "-numbuffers",
    "-freebuffers",
    "-eventslost",
    "-bufferswritten",
    "-logbufferslost",
    "-realtimebufferslost",
    "-loggerthread",
    NULL,
};


/* IMPORTANT : 
 * Do not change order without changing  ObjFromTRACE_LOGFILE_HEADER
 */
static const char * g_trace_logfile_header_fields[] = {
    "-buffersize",
    "-majorversion",
    "-minorversion",
    "-subversion",
    "-subminorversion",
    "-providerversion",
    "-numberofprocessors",
    "-endtime",
    "-timerresolution",
    "-maximumfilesize",
    "-logfilemode",
    "-bufferswritten",
    "-pointersize",
    "-eventslost",
    "-cpuspeedinmhz",
    "-loggername",
    "-logfilename",
    "-timezone",
    "-boottime",
    "-perffreq",
    "-starttime",
    "-bufferslost",
    NULL,
};


/* Event Trace Consumer Support */
struct TwapiETWContext {
    Tcl_Interp *interp;

    Tcl_Obj *buffer_cmdObj;     /* Callback for buffers */

    /*
     * When inside a ProcessTrace,
     * EXACTLY ONE of eventsObj and errorObj must be non-NULL. Moreover,
     * when non-NULL, an Tcl_IncrRefCount must have been done on it
     * with a corresponding Tcl_DecrRefCount when setting to NULL
      */
    Tcl_Obj *eventsObj;
    Tcl_Obj *errorObj;

    TRACEHANDLE traceH;

    int   buffer_cmdlen;        /* length of buffer_cmdObj */
    ULONG pointer_size;
    ULONG timer_resolution;
    ULONG user_mode;
} gETWContext;                  /* IMPORTANT : Sync access via gETWCS */

/* IMPORTANT : 
 * Used for events recieved in TwapiETWEventCallback.
 * Do not change order without changing the code there.
 */
/* IMPORTANT : Sync access via gETWCS */
struct TwapiObjKeyCache gETWEventKeys[] = {
    {"-eventtype"},
    {"-level"},
    {"-version"},
    {"-threadid"},
    {"-processid"},
    {"-timestamp"},
    {"-guid"},
    {"-kerneltime"},
    {"-usertime"},
    {"-processortime"},
    {"-instanceid"},
    {"-parentinstanceid"},
    {"-parentguid"},
    {"-mofdata"},
};

/* IMPORTANT : 
 * Used for events recieved in TwapiETWBufferCallback.
 * Do not change order without changing the code there.
 */
/* IMPORTANT : Sync access via gETWCS */
struct TwapiObjKeyCache gETWBufferKeys[] = {
    {"-logfilename"},
    {"-loggername"},
    {"-currenttime"},
    {"-buffersread"},
    {"-logfilemode"},
    {"-buffersize"},
    {"-filled"},
    {"-eventslost"},
    {"-iskerneltrace"},


    {"-hdr_cpuspeedinmhz"},
    {"-hdr_buffersize"},
    {"-hdr_majorversion"},
    {"-hdr_minorversion"},
    {"-hdr_subversion"},
    {"-hdr_subminorversion"},
    {"-hdr_providerversion"},
    {"-hdr_numberofprocessors"},
    {"-hdr_endtime"},
    {"-hdr_timerresolution"},
    {"-hdr_maximumfilesize"},
    {"-hdr_logfilemode"},
    {"-hdr_bufferswritten"},
    {"-hdr_pointersize"},
    {"-hdr_eventslost"},
    {"-hdr_bufferslost"},
    {"-hdr_timezone"},
    {"-hdr_boottime"},
    {"-hdr_perffreq"},
    {"-hdr_starttime"},
};


CRITICAL_SECTION gETWCS; /* Access to gETWContext, gETWEventKeys, gETWBufferKeys */



/*
 * Event Trace Provider Support
 *
 * Currently only support *one* provider, shared among all interps and
 * modules. The extra code for locking/bookkeeping etc. for multiple 
 * providers not likely to be used.
 */
GUID   gETWProviderGuid;        /* GUID for our ETW Provider */
TRACEHANDLE gETWProviderRegistrationHandle; /* HANDLE for registered instance of provider */
GUID   gETWProviderEventClassGuid; /* GUID for the event class we use */
HANDLE gETWProviderEventClassRegistrationHandle; /* Handle returned by ETW for above */
ULONG  gETWProviderTraceEnableFlags;             /* Flags set by ETW controller */
ULONG  gETWProviderTraceEnableLevel;             /* Level set by ETW controller */
/* Session our provider is attached to */
TRACEHANDLE gETWProviderSessionHandle = (TRACEHANDLE) INVALID_HANDLE_VALUE;


/*
 * Whether the callback dll/libray has been initialized.
 * The value must be managed using the InterlockedCompareExchange functions to
 * ensure thread safety. The value returned by InterlockedCompareExhange
 * 0 -> first to call, do init,  1 -> init in progress by some other thread
 * 2 -> Init done
 */
static TwapiOneTimeInitState gETWInitialized;

#ifndef TWAPI_SINGLE_MODULE
HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

static GUID gNullGuid;             /* Initialized to all zeroes */

/* Prototypes */
TCL_RESULT Twapi_RegisterTraceGuids(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_UnregisterTraceGuids(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_TraceEvent(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_OpenTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_CloseTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_EnableTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_ControlTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_StartTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_ProcessTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
TCL_RESULT Twapi_ParseEventMofData(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);

/*
 * Functions
 */
#ifdef NOTUSEDYET
/* IMPORTANT : 
 * Do not change order without changing  ObjFromTRACE_LOGFILE_HEADER
 */
static const char * g_trace_logfile_header_fields[] = {
    "-buffersize",
    "-majorversion",
    "-minorversion",
    "-subversion",
    "-subminorversion",
    "-providerversion",
    "-numberofprocessors",
    "-endtime",
    "-timerresolution",
    "-maximumfilesize",
    "-logfilemode",
    "-bufferswritten",
    "-pointersize",
    "-eventslost",
    "-cpuspeedinmhz",
    "-loggername",
    "-logfilename",
    "-timezone",
    "-boottime",
    "-perffreq",
    "-starttime",
    "-bufferslost",
    NULL,
};
static Tcl_Obj *ObjFromTRACE_LOGFILE_HEADER(TRACE_LOGFILE_HEADER *tlhP)
{
    int i;
    Tcl_Obj *objs[44];
    TRACE_LOGFILE_HEADER *adjustedP;

    for (i = 0; g_trace_logfile_header_fields[i]; ++i) {
        objs[2*i] = ObjFromString(g_trace_logfile_header_fields[i]);
    }
    TWAPI_ASSERT((2*i) == ARRAYSIZE(objs));

    objs[1] = ObjFromULONG(tlhP->BufferSize);
    objs[3] = ObjFromInt(tlhP->VersionDetail.MajorVersion);
    objs[5] = ObjFromInt(tlhP->VersionDetail.MinorVersion);
    objs[7] = ObjFromInt(tlhP->VersionDetail.SubVersion);
    objs[9] = ObjFromInt(tlhP->VersionDetail.SubMinorVersion);
    objs[11] = ObjFromULONG(tlhP->ProviderVersion);
    objs[13] = ObjFromULONG(tlhP->NumberOfProcessors);
    objs[15] = ObjFromLARGE_INTEGER(tlhP->EndTime);
    objs[17] = ObjFromULONG(tlhP->TimerResolution);
    objs[19] = ObjFromULONG(tlhP->MaximumFileSize);
    objs[21] = ObjFromULONG(tlhP->LogFileMode);
    objs[23] = ObjFromULONG(tlhP->BuffersWritten);
    objs[25] = ObjFromULONG(tlhP->PointerSize);
    objs[27] = ObjFromULONG(tlhP->EventsLost);
    objs[29] = ObjFromULONG(tlhP->CpuSpeedInMHz);

    /* Actual position of remaining fields may not match.
     * if the file came from a different architecture
     * so we have to adjust pointer accordingly. See 
     * "Implementing an Event Callback Function" in the SDK docs
     */

    if (tlhP->PointerSize != sizeof(PVOID)) {
        adjustedP = (PTRACE_LOGFILE_HEADER)
            ((PUCHAR)tlhP + 2 * (tlhP->PointerSize - sizeof(PVOID)));
    } else
        adjustedP = tlhP;

    /* Now continue with remaining fields */
    
    /*
     * LoggerName and LogFileName fields are not to be used. Rather they
     * refer to null terminated fields right after this struct.
     * HOWEVER, it is not clear what happens when the struct is 
     * embedded inside another so for now we just punt and return
     * empty strings;
    */
#if 1
    objs[31] = ObjFromEmptyString();
    objs[33] = objs[31];    /* OK, Tcl_NewListObj will just incr ref twice */
#else
    ws = (WCHAR *)(sizeof(*adjustedP)+(char *)adjustedP);
    objs[31] = ObjFromUnicode(ws);
    ObjToUnicodeN(objs[31], &i); /* Get length of LoggerName... */
    ws += (i+1);                         /* so we can point beyond it... */
    objs[33] = ObjFromUnicode(ws);       /* to the LogFileName */
#endif

    objs[35] = ObjFromTIME_ZONE_INFORMATION(&adjustedP->TimeZone);
    
    objs[37] = ObjFromLARGE_INTEGER(adjustedP->BootTime);
    objs[39] = ObjFromLARGE_INTEGER(adjustedP->PerfFreq);
    objs[41] = ObjFromLARGE_INTEGER(adjustedP->StartTime);
    objs[43] = ObjFromULONG(adjustedP->BuffersLost);

    return ObjNewList(ARRAYSIZE(objs), objs);
}
#endif


TCL_RESULT ObjToPEVENT_TRACE_PROPERTIES(
    Tcl_Interp *interp,
    Tcl_Obj *objP,
    EVENT_TRACE_PROPERTIES **etPP)
{
    int i, buf_sz;
    Tcl_Obj **objv;
    int       objc;
    int       logfile_name_i;
    int       session_name_i;
    WCHAR    *logfile_name;
    WCHAR    *session_name;
    EVENT_TRACE_PROPERTIES *etP;
    Tcl_WideInt wide;
    int field;

    if (ObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc & 1) {
        TwapiSetStaticResult(interp, "Invalid EVENT_TRACE_PROPERTIES format");
        return TCL_ERROR;
    }

    /* First loop and find out required buffers size */
    session_name_i = -1;
    logfile_name_i = -1;
    for (i = 0 ; i < objc ; i += 2) {
        if (Tcl_GetIndexFromObj(interp, objv[i], g_event_trace_fields, "event trace field", TCL_EXACT, &field) != TCL_OK)
            return TCL_ERROR;
        switch (field) {
        case 0: // -logfile
            logfile_name_i = i; /* Duplicates result in last value being used */
            break;
        case 1: // -sessionname
            session_name_i = i;
            break;
        }
    }
    
    if (session_name_i >= 0) {
        session_name = ObjToUnicodeN(objv[session_name_i+1], &session_name_i);
    } else {
        session_name_i = 0;
    }

    if (logfile_name_i >= 0) {
        logfile_name = ObjToUnicodeN(objv[logfile_name_i+1], &logfile_name_i);
    } else {
        logfile_name_i = 0;
    }

    /* Note: session_name_i/logfile_name_i now contain lengths of names */

    buf_sz = sizeof(*etP);
    /* Note 0 length names means we do not specify it or know its length */
    if (session_name_i > 0) {
        buf_sz += sizeof(WCHAR) * (session_name_i+1);
    } else {
        /* Leave max space in case kernel wants to fill it in */
        buf_sz += sizeof(WCHAR) * (MAX_TRACE_NAME_CHARS+1);
    }
    if (logfile_name_i > 0) {
        /* Logfile name may be appended with numeric strings (PID/sequence) */
        buf_sz += sizeof(WCHAR) * (logfile_name_i+1 + 20);
    } else {
        /* Leave max space in case kernel wants to fill it in on a query */
        buf_sz += sizeof(WCHAR) * (MAX_TRACE_NAME_CHARS + 1);
    }

    etP = TwapiAlloc(buf_sz);
    ZeroMemory(etP, buf_sz);
    etP->Wnode.Flags = WNODE_FLAG_TRACED_GUID;
    etP->Wnode.BufferSize = buf_sz;

    /* Copy the session name and log file name to end of struct.
     * Note that the session name MUST come first
     */
    etP->LoggerNameOffset = sizeof(*etP);
    if (session_name_i > 0) {
        /*
         * We do not need to actually copy this here for StartTrace as it does
         * itself. But our own Twapi_ControlTrace picks up the name from
         * this structure so we need to copy it.
         */
        CopyMemory(etP->LoggerNameOffset + (char *)etP,
                   session_name, sizeof(WCHAR) * (session_name_i + 1));
        etP->LogFileNameOffset = sizeof(*etP) + (sizeof(WCHAR) * (session_name_i + 1));
    } else {
        /* We left max space for unknown session name */
        etP->LogFileNameOffset = sizeof(*etP) + (sizeof(WCHAR) * (MAX_TRACE_NAME_CHARS+1));
    }
        
    if (logfile_name_i > 0) {
        CopyMemory(etP->LogFileNameOffset + (char *) etP,
                   logfile_name, sizeof(WCHAR) * (logfile_name_i + 1));
    } else {
#if 0
        Do not set to 0 as ControlTrace Query may need to store filename
        etP->LogFileNameOffset = 0;
#endif
    }

    /* Now deal with the rest of the fields, after setting some defaults */
    etP->Wnode.ClientContext = 2; /* Default - system timer resolution is cheape
st */
    etP->LogFileMode = EVENT_TRACE_USE_PAGED_MEMORY;

    for (i = 0 ; i < objc ; i += 2) {
        ULONG *ulP;

        if (Tcl_GetIndexFromObj(interp, objv[i], g_event_trace_fields, "event trace field", TCL_EXACT, &field) == TCL_ERROR)
            return TCL_ERROR;
        ulP = NULL;
        switch (field) {
        case 0: // -logfile
        case 1: // -sessionname
            /* Already handled above */
            break;
        case 2: //-sessionguid
            if (ObjToGUID(interp, objv[i+1], &etP->Wnode.Guid) != TCL_OK)
                goto error_handler;
            break;
        case 3: // -buffersize
            ulP = &etP->BufferSize;
            break;
        case 4: // -minbuffers
            ulP = &etP->MinimumBuffers;
            break;
        case 5: // -maxbuffers
            ulP = &etP->MaximumBuffers;
            break;
        case 6: // -maximumfilesize
            ulP = &etP->MaximumFileSize;
            break;
        case 7: // -logfilemode
            ulP = &etP->LogFileMode;
            break;
        case 8: // -flushtimer
            ulP = &etP->FlushTimer;
            break;
        case 9: // -enableflags
            ulP = &etP->EnableFlags;
            break;
        case 10: // -clockresolution
            ulP = &etP->Wnode.ClientContext;
            break;
        case 11: // -agelimit
            ulP = &etP->AgeLimit;
            break;
        case 12: // -numbuffers
            ulP = &etP->NumberOfBuffers;
            break;
        case 13: // -freebuffers
            ulP = &etP->FreeBuffers;
            break;
        case 14: // -eventslost
            ulP = &etP->EventsLost;
            break;
        case 15: // -bufferswritten
            ulP = &etP->BuffersWritten;
            break;
        case 16: // -logbufferslost
            ulP = &etP->LogBuffersLost;
            break;
        case 17: // -realtimebufferslost
            ulP = &etP->RealTimeBuffersLost;
            break;
        case 18: // -loggerthread
            if (ObjToWideInt(interp, objv[i+1], &wide) != TCL_OK)
                goto error_handler;
            etP->LoggerThreadId = (HANDLE) wide;
            break;
        default:
            TwapiSetStaticResult(interp, "Internal error: Unexpected field index.");
            goto error_handler;
        }
        
        if (ulP == NULL) {
            /* Nothing to do, value already set in the switch */
        } else {
            if (ObjToLong(interp, objv[i+1], ulP) != TCL_OK)
                goto error_handler;
        }
    }
    
    *etPP = etP;
    return TCL_OK;

error_handler: /* interp must already contain error msg */
    if (etP)
        TwapiFree(etP);
    return TCL_ERROR;
}


static Tcl_Obj *ObjFromEVENT_TRACE_PROPERTIES(EVENT_TRACE_PROPERTIES *etP)
{
    int i;
    Tcl_Obj *objs[38];

    for (i = 0; g_event_trace_fields[i]; ++i) {
        objs[2*i] = ObjFromString(g_event_trace_fields[i]);
    }
    TWAPI_ASSERT((2*i) == ARRAYSIZE(objs));

    if (etP->LogFileNameOffset)
        objs[1] = ObjFromUnicode((WCHAR *)(etP->LogFileNameOffset + (char *) etP));
    else
        objs[1] = ObjFromEmptyString();

    if (etP->LoggerNameOffset)
        objs[3] = ObjFromUnicode((WCHAR *)(etP->LoggerNameOffset + (char *) etP));
    else
        objs[3] = ObjFromEmptyString();

    objs[5] = ObjFromGUID(&etP->Wnode.Guid);
    objs[7] = ObjFromLong(etP->BufferSize);
    objs[9] = ObjFromLong(etP->MinimumBuffers);
    objs[11] = ObjFromLong(etP->MaximumBuffers);
    objs[13] = ObjFromLong(etP->MaximumFileSize);
    objs[15] = ObjFromLong(etP->LogFileMode);
    objs[17] = ObjFromLong(etP->FlushTimer);
    objs[19] = ObjFromLong(etP->EnableFlags);
    objs[21] = ObjFromLong(etP->Wnode.ClientContext);
    objs[23] = ObjFromLong(etP->AgeLimit);
    objs[25] = ObjFromLong(etP->NumberOfBuffers);
    objs[27] = ObjFromLong(etP->FreeBuffers);
    objs[29] = ObjFromLong(etP->EventsLost);
    objs[31] = ObjFromLong(etP->BuffersWritten);
    objs[33] = ObjFromLong(etP->LogBuffersLost);
    objs[35] = ObjFromLong(etP->RealTimeBuffersLost);
    objs[37] = ObjFromWideInt((Tcl_WideInt) etP->LoggerThreadId);

    return ObjNewList(ARRAYSIZE(objs), objs);
}


static ULONG WINAPI TwapiETWProviderControlCallback(
    WMIDPREQUESTCODE request,
    PVOID contextP,             /* TBD - can we use this ? */
    ULONG* reserved, 
    PVOID headerP
    )
{
    ULONG rc = ERROR_SUCCESS;
    TRACEHANDLE session;
    ULONG enable_level;
    ULONG enable_flags;

    /* Mostly cloned from the SDK docs */

    switch (request) {
    case WMI_ENABLE_EVENTS:
        SetLastError(0);
        session = GetTraceLoggerHandle(headerP);
        if (INVALID_SESSIONTRACE_HANDLE(session)) {
            /* Bad handle, ignore, but return the error code */
            rc = GetLastError();
            break;
        }

        /* If we are already logging to a session we will ignore nonmatching */
        if (! INVALID_SESSIONTRACE_HANDLE(gETWProviderSessionHandle) &&
            session != gETWProviderSessionHandle) {
            rc = ERROR_INVALID_PARAMETER;
            break;
        }

        SetLastError(0);
        enable_level = GetTraceEnableLevel(session); 
        if (enable_level == 0) {
            /* *Possible* error */
            rc = GetLastError();
            if (rc)
                break; /* Yep, real error */
        }

        SetLastError(0);
        enable_flags = GetTraceEnableFlags(session);
        if (enable_flags == 0) {
            /* *Possible* error */
            rc = GetLastError();
            if (rc)
                break;
        }

        /* OK, all calls succeeded. Set up the global values */
        gETWProviderSessionHandle = session;
        gETWProviderTraceEnableLevel = enable_level;
        gETWProviderTraceEnableFlags = enable_flags;
        break;
 
    case WMI_DISABLE_EVENTS:  //Disable Provider.
        /* Don't we need to check session handle ? Sample MSDN does not */
        gETWProviderSessionHandle = (TRACEHANDLE) INVALID_HANDLE_VALUE;
        break;

    default:
        rc = ERROR_INVALID_PARAMETER;
        break;
    }

    return rc;
}


TCL_RESULT Twapi_RegisterTraceGuids(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD rc;
    GUID provider_guid, event_class_guid;
    TRACE_GUID_REGISTRATION event_class_reg;
    
    if (TwapiGetArgs(interp, objc-1, objv+1, GETUUID(provider_guid),
                     GETUUID(event_class_guid), ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (IsEqualGUID(&provider_guid, &gNullGuid) ||
        IsEqualGUID(&event_class_guid, &gNullGuid)) {
        TwapiSetStaticResult(interp, "NULL provider GUID specified.");
        return TCL_ERROR;
    }

    /* We should not already have registered a different provider */
    if (! IsEqualGUID(&gETWProviderGuid, &gNullGuid)) {
        if (IsEqualGUID(&gETWProviderGuid, &provider_guid)) {
            TwapiSetObjResult(interp, ObjFromTRACEHANDLE(gETWProviderRegistrationHandle));

            return TCL_OK;      /* Same GUID - ok */
        }
        else {
            TwapiSetStaticResult(interp, "ETW Provider GUID already registered");
            return TCL_ERROR;
        }
    }

    ZeroMemory(&event_class_reg, sizeof(event_class_reg));
    event_class_reg.Guid = &event_class_guid;
    rc = RegisterTraceGuids(
        (WMIDPREQUEST)TwapiETWProviderControlCallback,
        NULL,                   // No context TBD - can we use this ?
        &provider_guid,         // GUID that identifies the provider
        1,                      /* Number of event class GUIDs */
        &event_class_reg,      /* Event class GUID array */
        NULL,                          // Not used
        NULL,                          // Not used
        &gETWProviderRegistrationHandle // Used for UnregisterTraceGuids
        );

    if (rc == ERROR_SUCCESS) {
        gETWProviderGuid = provider_guid;
        gETWProviderEventClassGuid = event_class_guid;
        gETWProviderEventClassRegistrationHandle = event_class_reg.RegHandle;
        TwapiSetObjResult(interp, ObjFromTRACEHANDLE(gETWProviderRegistrationHandle));
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, rc);
    }
}

TCL_RESULT Twapi_UnregisterTraceGuids(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TRACEHANDLE traceH;
    DWORD rc;
    
    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (ObjToTRACEHANDLE(interp, objv[1], &traceH) != TCL_OK)
        return TCL_ERROR;

    if (traceH != gETWProviderRegistrationHandle) {
        TwapiSetStaticResult(interp, "Unknown ETW provider registration handle");
        return TCL_ERROR;
    }

    rc = UnregisterTraceGuids(gETWProviderRegistrationHandle);
    if (rc == ERROR_SUCCESS) {
        gETWProviderRegistrationHandle = 0;
        gETWProviderGuid = gNullGuid;
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, rc);
    }
}

TCL_RESULT Twapi_TraceEvent(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int i;
    struct {
        EVENT_TRACE_HEADER eth;
        MOF_FIELD mof[16];
    } event;
    int    type;
    int    level;
    TRACEHANDLE htrace;
    ULONG rc;

    /* Args: provider handle, type, level, ?binary strings...? */

    /* If no tracing is on, just ignore, do not raise error */
    if (INVALID_SESSIONTRACE_HANDLE(gETWProviderSessionHandle))
        return TCL_OK;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVAR(htrace, ObjToTRACEHANDLE),
                     GETINT(type), GETINT(level), ARGTERM) != TCL_OK)
        return TCL_ERROR;

    objc -= 4;
    objv += 4;

    /* We will only log up to 16 strings. Additional will be silently ignored */
    if (objc > 16)
        objc = 16;

    ZeroMemory(&event, sizeof(EVENT_TRACE_HEADER) + objc*sizeof(MOF_FIELD));
    event.eth.Size = (USHORT) (sizeof(EVENT_TRACE_HEADER) + objc*sizeof(MOF_FIELD));
    event.eth.Flags = WNODE_FLAG_TRACED_GUID |
        WNODE_FLAG_USE_GUID_PTR | WNODE_FLAG_USE_MOF_PTR;
    event.eth.GuidPtr = (ULONGLONG) &gETWProviderEventClassGuid;
    event.eth.Class.Type = type;
    event.eth.Class.Level = level;

    for (i = 0; i < objc; ++i) {
        event.mof[i].DataPtr = (ULONG64) ObjToByteArray(objv[i], &event.mof[i].Length);
    }

    rc = TraceEvent(gETWProviderSessionHandle, &event.eth);
    return rc == ERROR_SUCCESS ?
        TCL_OK : Twapi_AppendSystemError(interp, rc);
}


TCL_RESULT Twapi_StartTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    EVENT_TRACE_PROPERTIES *etP;
    TRACEHANDLE htrace;
    Tcl_Obj *objs[2];
    
    if (objc != 3)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (ObjToPEVENT_TRACE_PROPERTIES(interp, objv[2], &etP) != TCL_OK)
        return TCL_ERROR;

    /* If no log file specified, set logfilenameoffset to 0 
       since ObjToPEVENT_TRACE_PROPERTIES does not do that as it
       is also used by ControlTrace */
    if (*(WCHAR *) (etP->LogFileNameOffset + (char *)etP) == 0)
        etP->LogFileNameOffset = 0;

    if (StartTraceW(&htrace,
                    ObjToUnicode(objv[1]),
                    etP) != ERROR_SUCCESS) {
        Twapi_AppendSystemError(interp, GetLastError());
        TwapiFree(etP);
        return TCL_ERROR;
    }

    objs[0] = ObjFromTRACEHANDLE(htrace);
    objs[1] = ObjFromEVENT_TRACE_PROPERTIES(etP);
    TwapiFree(etP);
    TwapiSetObjResult(interp, ObjNewList(2, objs));
    return TCL_OK;
}

TCL_RESULT Twapi_ControlTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TRACEHANDLE htrace;
    EVENT_TRACE_PROPERTIES *etP;
    WCHAR *session_name;
    ULONG code;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(code), GETWIDE(htrace), ARGSKIP,
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (ObjToPEVENT_TRACE_PROPERTIES(interp, objv[3], &etP) != TCL_OK)
        return TCL_ERROR;

    if (etP->LoggerNameOffset &&
        (*(WCHAR *) (etP->LoggerNameOffset + (char *)etP) != 0))
        session_name = (WCHAR *)(etP->LoggerNameOffset + (char *)etP);
    else
        session_name = NULL;

    if (ControlTraceW(htrace, session_name, etP, code) != ERROR_SUCCESS) {
        Twapi_AppendSystemError(interp, GetLastError());
        code = TCL_ERROR;
    } else {
        TwapiSetObjResult(interp, ObjFromEVENT_TRACE_PROPERTIES(etP));
        code = TCL_OK;
    }

    TwapiFree(etP);
    return code;
}

TCL_RESULT Twapi_EnableTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    GUID guid;
    TRACEHANDLE htrace;
    ULONG enable, flags, level;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(enable), GETINT(flags), GETINT(level),
                     GETUUID(guid), GETWIDE(htrace),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    if (EnableTrace(enable, flags, level, &guid, htrace) != ERROR_SUCCESS)
        return TwapiReturnSystemError(interp);

    return TCL_OK;
}


void WINAPI TwapiETWEventCallback(
  PEVENT_TRACE evP
)
{
    Tcl_Interp *interp;
    Tcl_Obj *evObj;

    /* Called back from Win32 ProcessTrace call. Assumed that gETWContext is locked */
    TWAPI_ASSERT(gETWContext.interp != NULL);
    interp = gETWContext.interp;

    if (gETWContext.errorObj)   /* If some previous error occurred, return */
        return;

    /* IMPORTANT: the order is tied to order of gETWEventKeys[] ! */
    evObj = Tcl_NewDictObj();
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[0].keyObj, ObjFromInt(evP->Header.Class.Type));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[1].keyObj, ObjFromInt(evP->Header.Class.Level));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[2].keyObj, ObjFromInt(evP->Header.Class.Version));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[3].keyObj, ObjFromULONG(evP->Header.ThreadId));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[4].keyObj, ObjFromULONG(evP->Header.ProcessId));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[5].keyObj, ObjFromLARGE_INTEGER(evP->Header.TimeStamp));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[6].keyObj, ObjFromGUID(&evP->Header.Guid));
    /*
     * Note - for user mode sessions, KernelTime/UserTime are not valid
     * and the ProcessorTime member has to be used instead. However,
     * we do not know the type of session at this point so we leave it
     * to the app to figure out what to use
     */
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[7].keyObj, ObjFromULONG(evP->Header.KernelTime));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[8].keyObj, ObjFromULONG(evP->Header.UserTime));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[9].keyObj, ObjFromWideInt(evP->Header.ProcessorTime));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[10].keyObj, ObjFromULONG(evP->InstanceId));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[11].keyObj, ObjFromULONG(evP->ParentInstanceId));
    Tcl_DictObjPut(interp, evObj, gETWEventKeys[12].keyObj, ObjFromGUID(&evP->ParentGuid));
    if (evP->MofData && evP->MofLength)
        Tcl_DictObjPut(interp, evObj, gETWEventKeys[13].keyObj, ObjFromByteArray(evP->MofData, evP->MofLength));
    else
        Tcl_DictObjPut(interp, evObj, gETWEventKeys[13].keyObj, ObjFromEmptyString());
    
    ObjAppendElement(interp, gETWContext.eventsObj, evObj);


    return;
}

ULONG WINAPI TwapiETWBufferCallback(
  PEVENT_TRACE_LOGFILEW etlP
)
{
    TRACE_LOGFILE_HEADER *tlhP;
    TRACE_LOGFILE_HEADER *adjustedP;
    Tcl_Obj *objP;
    Tcl_Obj *bufObj;
    Tcl_Obj *args[2];
    Tcl_Interp *interp;
    int code;

    /* Called back from Win32 ProcessTrace call. Assumed that gETWContext is locked */
    TWAPI_ASSERT(gETWContext.interp != NULL);
    interp = gETWContext.interp;

    if (Tcl_InterpDeleted(interp))
        return FALSE;

    if (gETWContext.errorObj)   /* If some previous error occurred, return */
        return FALSE;

    if (gETWContext.buffer_cmdObj == NULL)
        return TRUE;            /* No buffer callback specified */

    TWAPI_ASSERT(gETWContext.eventsObj); /* since errorObj not NULL at this point */

    /*
     * Construct a command to call with the event. gETWContext.buffer_cmdObj
     * could be a shared object, either initially itself or result in a shared 
     * object in the callback. So we need to check for that and Dup it
     * if necessary
     */
    if (Tcl_IsShared(gETWContext.buffer_cmdObj)) {
        objP = Tcl_DuplicateObj(gETWContext.buffer_cmdObj);
        Tcl_IncrRefCount(objP);
    } else
        objP = gETWContext.buffer_cmdObj;

    /* Build up the arguments */
    bufObj = Tcl_NewDictObj();
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[0].keyObj, ObjFromUnicode(etlP->LogFileName ? etlP->LogFileName : L""));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[1].keyObj, ObjFromUnicode(etlP->LoggerName ? etlP->LoggerName : L""));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[2].keyObj, ObjFromULONGLONG(etlP->CurrentTime));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[3].keyObj, ObjFromLong(etlP->BuffersRead));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[4].keyObj, ObjFromLong(etlP->LogFileMode));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[5].keyObj, ObjFromLong(etlP->BufferSize));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[6].keyObj, ObjFromLong(etlP->Filled));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[7].keyObj, ObjFromLong(etlP->EventsLost));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[8].keyObj, ObjFromLong(etlP->IsKernelTrace));


    /* Add the fields from the logfile header */

    tlhP = &etlP->LogfileHeader;
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[9].keyObj, ObjFromULONG(tlhP->CpuSpeedInMHz));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[10].keyObj, ObjFromULONG(tlhP->BufferSize));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[11].keyObj, ObjFromInt(tlhP->VersionDetail.MajorVersion));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[12].keyObj, ObjFromInt(tlhP->VersionDetail.MinorVersion));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[13].keyObj, ObjFromInt(tlhP->VersionDetail.SubVersion));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[14].keyObj, ObjFromInt(tlhP->VersionDetail.SubMinorVersion));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[15].keyObj, ObjFromULONG(tlhP->ProviderVersion));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[16].keyObj, ObjFromULONG(tlhP->NumberOfProcessors));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[17].keyObj, ObjFromLARGE_INTEGER(tlhP->EndTime));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[18].keyObj, ObjFromULONG(tlhP->TimerResolution));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[19].keyObj, ObjFromULONG(tlhP->MaximumFileSize));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[20].keyObj, ObjFromULONG(tlhP->LogFileMode));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[21].keyObj, ObjFromULONG(tlhP->BuffersWritten));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[22].keyObj, ObjFromULONG(tlhP->PointerSize));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[23].keyObj, ObjFromULONG(tlhP->EventsLost));

    /* Actual position of remaining fields may not match.
     * if the file came from a different architecture
     * so we have to adjust pointer accordingly. See 
     * "Implementing an Event Callback Function" in the SDK docs
     */

    if (tlhP->PointerSize != sizeof(PVOID)) {
        adjustedP = (PTRACE_LOGFILE_HEADER)
            ((PUCHAR)tlhP + 2 * (tlhP->PointerSize - sizeof(PVOID)));
    } else
        adjustedP = tlhP;

    /* Now continue with remaining fields */
    
    /* NOTE: 
     * LoggerName and LogFileName fields are not to be used. Rather they
     * refer to null terminated fields right after this struct.
     * HOWEVER, it is not clear what happens when the struct is 
     * embedded inside another so for now we just do not include them
     * in the returned dictionary.
    */

    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[24].keyObj, ObjFromULONG(adjustedP->BuffersLost));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[25].keyObj, ObjFromTIME_ZONE_INFORMATION(&adjustedP->TimeZone));    
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[26].keyObj, ObjFromLARGE_INTEGER(adjustedP->BootTime));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[27].keyObj, ObjFromLARGE_INTEGER(adjustedP->PerfFreq));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[28].keyObj, ObjFromLARGE_INTEGER(adjustedP->StartTime));

    /*
     * Note: Do not need to Tcl_IncrRefCount bufObj[] because we are adding
     * it to the command list
     */
    args[0] = bufObj;
    args[1] = gETWContext.eventsObj;
    Tcl_ListObjReplace(interp, objP, gETWContext.buffer_cmdlen, ARRAYSIZE(args), ARRAYSIZE(args), args);
    Tcl_DecrRefCount(gETWContext.eventsObj); /* AFTER we place on list */

    gETWContext.eventsObj = ObjNewList(0, NULL);/* For next set of events */
    Tcl_IncrRefCount(gETWContext.eventsObj);

    code = Tcl_EvalObjEx(gETWContext.interp, objP, TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);

    /* Get rid of the command obj if we created it */
    if (objP != gETWContext.buffer_cmdObj)
        Tcl_DecrRefCount(objP);

    switch (code) {
    case TCL_BREAK:
        /* Any other value - not an error, but stop processing */
        return FALSE;
    case TCL_ERROR:
        gETWContext.errorObj = Tcl_GetReturnOptions(gETWContext.interp,
                                                    code);
        Tcl_IncrRefCount(gETWContext.errorObj);
        Tcl_DecrRefCount(gETWContext.eventsObj);
        gETWContext.eventsObj = NULL;
        return FALSE;
    default:
        /* Any other value - proceed as normal */
        return TRUE;
    }
}



TCL_RESULT Twapi_OpenTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TRACEHANDLE htrace;
    EVENT_TRACE_LOGFILEW etl;
    int real_time;
    WCHAR *s;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETWSTR(s), GETINT(real_time),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    ZeroMemory(&etl, sizeof(etl));
    if (real_time) {
        etl.LoggerName = s;
        etl.LogFileMode = EVENT_TRACE_REAL_TIME_MODE;
    } else 
        etl.LogFileName = s;

    etl.BufferCallback = TwapiETWBufferCallback;
    etl.EventCallback = TwapiETWEventCallback;

    htrace = OpenTraceW(&etl);
    if (INVALID_SESSIONTRACE_HANDLE(htrace))
        return TwapiReturnSystemError(interp);

    TwapiSetObjResult(interp, ObjFromTRACEHANDLE(htrace));
    return TCL_OK;
}

TCL_RESULT Twapi_CloseTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TRACEHANDLE htrace;
    ULONG err;

    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (ObjToTRACEHANDLE(interp, objv[1], &htrace) != TCL_OK)
        return TCL_ERROR;

    err = CloseTrace(htrace);
    if (err == ERROR_SUCCESS)
        return TCL_OK;
    else
        return Twapi_AppendSystemError(interp, err);
}


TCL_RESULT Twapi_ProcessTrace(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int i;
    FILETIME start, end, *startP, *endP;
    struct TwapiETWContext etwc;
    int buffer_cmdlen;
    int code;
    DWORD winerr;
    Tcl_Obj **htraceObjs;
    TRACEHANDLE htraces[8];
    int       ntraces;

    if (objc != 5)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (ObjGetElements(interp, objv[1], &ntraces, &htraceObjs) != TCL_OK)
        return TCL_ERROR;

    for (i = 0; i < ntraces; ++i) {
        if (ObjToTRACEHANDLE(interp, htraceObjs[i], &htraces[i]) != TCL_OK)
            return TCL_ERROR;
    }

    /* Verify command prefix is a list. It's OK to be empty, we
     * effectively drain the buffer without doing anything
     */
    buffer_cmdlen = 0;
    if (Tcl_ListObjLength(interp, objv[2], &buffer_cmdlen) != TCL_OK)
        return TCL_ERROR;

    if (Tcl_GetCharLength(objv[3]) == 0)
        startP = NULL;
    else if (ObjToFILETIME(interp, objv[3], &start) != TCL_OK)
            return TCL_ERROR;
    else
        startP = &start;
    
    if (Tcl_GetCharLength(objv[4]) == 0)
        endP = NULL;
    else if (ObjToFILETIME(interp, objv[4], &end) != TCL_OK)
            return TCL_ERROR;
    else
        endP = &end;
    
    EnterCriticalSection(&gETWCS);
    
    TWAPI_ASSERT(gETWContext.interp == NULL);

    gETWContext.traceH = htraces[0];
    gETWContext.buffer_cmdlen = buffer_cmdlen;
    gETWContext.buffer_cmdObj = buffer_cmdlen ? objv[1] : NULL;
    gETWContext.eventsObj = ObjNewList(0, NULL);
    Tcl_IncrRefCount(gETWContext.eventsObj);
    gETWContext.errorObj = NULL;
    gETWContext.interp = interp;

    /*
     * Initialize the dictionary objects we will use as keys. These are
     * tied to interp so have to redo on every call since the interp
     * may not be the same.
     */
    for (i = 0; i < ARRAYSIZE(gETWEventKeys); ++i) {
        gETWEventKeys[i].keyObj = ObjFromString(gETWEventKeys[i].keystring);
        Tcl_IncrRefCount(gETWEventKeys[i].keyObj);
    }
    for (i = 0; i < ARRAYSIZE(gETWBufferKeys); ++i) {
        gETWBufferKeys[i].keyObj = ObjFromString(gETWBufferKeys[i].keystring);
        Tcl_IncrRefCount(gETWBufferKeys[i].keyObj);
    }

    winerr = ProcessTrace(htraces, ntraces, startP, endP);

    /* Copy and reset context before unlocking */
    etwc = gETWContext;
    gETWContext.buffer_cmdObj = NULL;
    gETWContext.eventsObj = NULL;
    gETWContext.errorObj = NULL;
    gETWContext.interp = NULL;

    LeaveCriticalSection(&gETWCS);

    /* Deallocated the cached key objects */
    for (i = 0; i < ARRAYSIZE(gETWEventKeys); ++i) {
        Tcl_DecrRefCount(gETWEventKeys[i].keyObj);
        gETWEventKeys[i].keyObj = NULL; /* Just to catch bad access */
    }
    for (i = 0; i < ARRAYSIZE(gETWBufferKeys); ++i) {
        Tcl_DecrRefCount(gETWBufferKeys[i].keyObj);
        gETWBufferKeys[i].keyObj = NULL; /* Just to catch bad access */
    }

    if (etwc.eventsObj)
        Tcl_DecrRefCount(etwc.eventsObj);

    /* If the processing was successful, errorObj is NULL */
    if (etwc.errorObj) {
        code = Tcl_SetReturnOptions(interp, etwc.errorObj);
        Tcl_DecrRefCount(etwc.errorObj); /* Match one in the event callback */
        return code;
    }

    Tcl_ResetResult(interp); /* For any holdover from callbacks */

    /* A winerr of ERROR_CANCELLED means the callback returned TCL_BREAK
     * to terminate the processing. That is not treated as an error
     */
    if (winerr && winerr != ERROR_CANCELLED)
        return Twapi_AppendSystemError(interp, winerr);

    return TCL_OK;
}

TCL_RESULT Twapi_ParseEventMofData(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int       i;
    ULONG eaten;
    Tcl_Obj **types;            /* Field types */
    int       ntypes;           /* Number of fields/types */
    char     *bytesP;
    int       nbytes;
    ULONG     remain;
    Tcl_Obj  *resultObj = NULL;
    WCHAR     wc;
    GUID      guid;
    int       pointer_size;     /* Of target system, NOT us */
    union {
        SID sid;                /* For alignment */
        char buf[SECURITY_MAX_SID_SIZE];
    } u;
    short     port;

    if (objc != 4)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    
    bytesP = ObjToByteArray(objv[1], &nbytes);

    /* The field descriptor is a list of alternating field types and names */
    if (ObjGetElements(interp, objv[2], &ntypes, &types) != TCL_OK ||
        ObjToInt(interp, objv[3], &pointer_size) != TCL_OK)
        return TCL_ERROR;
        
    if (ntypes & 1) {
        return TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Field descriptor argument has odd number of elements (%d).", ntypes));
    }

    if (pointer_size != 4 && pointer_size != 8) {
        return TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS,
                           Tcl_ObjPrintf("Invalid pointer size parameter (%d), must be 4 or 8", pointer_size));
    }
    
    resultObj = Tcl_NewDictObj();

    for (i = 0, remain = nbytes; i < ntypes && remain > 0; i += 2, bytesP += eaten, remain -= eaten) {
        int typeenum;
        Tcl_Obj *objP;

        /* Index i is field name, i+1 is field type */

        if (ObjToInt(interp, types[i+1], &typeenum) != TCL_OK) {
            goto error_handler;
        }

        if (typeenum == 0) {
            /*
             * The type is valid but we don't support it and have no idea
             * how long it is so we cannot loop further.
             */
            break;
        }

        /* IMPORTANT: switch values are based on script _etw_typeenum defs */
        switch (typeenum) {
        case 1: // string / stringnullterminated
            /* We cannot rely on event being formatted correctly with a \0
               do not directly call Tcl_NewStringObj */
            objP = ObjFromStringLimited(bytesP, remain, &eaten);
            /* eaten is num remaining. Prime for loop iteration to
               be num used */
            eaten = remain - eaten;
            break;

        case 2: // wstring / wstringnullterminated
            /* Data may not be aligned ! Copy to align if necessary? TBD */
            /* We cannot rely on event being formatted correctly with a \0
               do not directly call Tcl_NewStringObj */
            objP = ObjFromUnicodeLimited((WCHAR *)bytesP, remain/sizeof(WCHAR), &eaten);
            /* eaten is num WCHARS remaining. Prime for loop iteration to
               be num used */
            eaten = remain - (sizeof(WCHAR)*eaten);
            break;

        case 3: // stringcounted
        case 4: // stringreversecounted
            if (remain < 2)
                goto done;      /* Data truncation */
            eaten = *(unsigned short UNALIGNED *)bytesP;
            if (typeenum == 4)  /* Need to swap bytes */
                eaten = ((eaten & 0xff) << 8) | (eaten >> 8);
            if (remain < (2+eaten)) {
                /* truncated */
                eaten = remain-2;
            }
            objP = ObjFromStringN(bytesP+2, eaten);
            eaten += 2;         /* include the length field */
            break;

        case 5: // wstringcounted
        case 6: // wstringreversecounted
            if (remain < 2)
                goto done;      /* Data truncation */
            eaten = *(unsigned short UNALIGNED *)bytesP;
            if (typeenum == 6)  /* Need to swap bytes */
                eaten = ((eaten & 0xff) << 8) | (eaten >> 8);
            /* Note eaten is num *characters*, not bytes at this point */
            /* TBD - there is much confusion as to whether the this
               is actually true. The latest SDK docs (Windows 7) states
               this is the case and its code example also treats it
               this way. On the other hand, older SDK's treat this
               as number of *bytes* in their sample code as does the
               well know LogParser utility. On XP there does not seem
               to be a MoF that actually uses this so maybe it is moot.
               For now leave as is 
            */

            if ((remain-2) < (sizeof(WCHAR)*eaten)) {
                /* truncated */
                objP = ObjFromUnicodeN((WCHAR *)(bytesP+2), (remain-2)/sizeof(WCHAR));
                eaten = remain; /* We used up all */
            } else {
                objP = ObjFromUnicodeN((WCHAR *)(bytesP+2), eaten);
                eaten = (sizeof(WCHAR)*eaten) + 2; /* include length field */
            }
            break;

        case 7: // boolean
            if (remain < sizeof(int))
                goto done;      /* Data truncation */
            objP = ObjFromBoolean((*(int UNALIGNED *)bytesP) ? 1 : 0);
            eaten = sizeof(int);
            break;

        case 8: // sint8
        case 9: // uint8
            objP = ObjFromInt(typeenum == 8 ? *(signed char *)bytesP : *(unsigned char *)bytesP);
            eaten = sizeof(char);
            break;
            
        case 10: // csint8
        case 11: // cuint8
            /* Return as an ascii char */
            objP = ObjFromStringN(bytesP, 1);
            eaten = sizeof(char);
            break;

        case 12: // sint16
        case 13: // uint16
            if (remain < sizeof(short))
                goto done;      /* Data truncation */
            objP = ObjFromInt(typeenum == 12 ?
                                 *(signed short UNALIGNED *)bytesP :
                                 *(unsigned short UNALIGNED *)bytesP);
            eaten = sizeof(short);
            break;

        case 14: // sint32
            if (remain < sizeof(int))
                goto done;      /* Data truncation */
            objP = ObjFromInt(*(signed int UNALIGNED *)bytesP);
            eaten = sizeof(int);
            break;

        case 15: // uint32
            if (remain < sizeof(int))
                goto done;      /* Data truncation */
            objP = ObjFromDWORD(*(unsigned int UNALIGNED *)bytesP);
            eaten = sizeof(int);
            break;

        case 16: // sint64
        case 17: // uint64
            if (remain < sizeof(__int64))
                goto done;      /* Data truncation */
            if (typeenum == 17) {
                objP = ObjFromULONGLONG(*(unsigned __int64 UNALIGNED *)bytesP);
            } else {
                objP = ObjFromWideInt(*(__int64 UNALIGNED *)bytesP);
            }
            eaten = sizeof(__int64);
            break;

        case 18: // xsint16
        case 19: // xuint16
            if (remain < sizeof(short))
                goto done;      /* Data truncation */
            objP = Tcl_ObjPrintf("0x%x", *(unsigned short UNALIGNED *)bytesP);
            eaten = sizeof(short);
            break;

        case 20: // xsint32
        case 21: // xuint32
            if (remain < sizeof(int))
                goto done;      /* Data truncation */
            objP = ObjFromULONGHex((ULONG) *(int UNALIGNED *)bytesP);
            eaten = sizeof(int);
            break;

        case 22: //xsint64
        case 23: //xuint64
            if (remain < sizeof(__int64))
                goto done;      /* Data truncation */
            objP = ObjFromULONGLONGHex(*(unsigned __int64 UNALIGNED *)bytesP);
            eaten = sizeof(__int64);
            break;
            
        case 24: // real32
            if (remain < 4)
                goto done;      /* Data truncation */
            objP = Tcl_NewDoubleObj((double)(*(float UNALIGNED *)bytesP));
            eaten = 4;
            break;

        case 25: // real64
            if (remain < 8)
                goto done;      /* Data truncation */
            objP = Tcl_NewDoubleObj((*(double UNALIGNED *)bytesP));
            eaten = 8;
            break;

        case 26: // object
            /* We should never see this except for a bad MoF definition where
             * the object does not have an associated qualifier. We can
             * only punt 
             */
            goto done;

        case 27: // char16
            if (remain < sizeof(WCHAR))
                goto done;      /* Data truncation */
            wc = *(WCHAR UNALIGNED *) bytesP;
            objP = ObjFromUnicodeN(&wc, 1);
            eaten = sizeof(WCHAR);
            break;

        case 28: // uint8guid
        case 29: // objectguid
            if (remain < sizeof(GUID))
                goto done;      /* Data truncation */
            CopyMemory(&guid, bytesP, sizeof(GUID)); /* For alignment reasons */
            objP = ObjFromGUID(&guid);
            eaten = sizeof(GUID);
            break;

        case 30: // objectipaddr, objectipaddrv4, uint32ipaddr
            if (remain < sizeof(DWORD))
                goto done;      /* Data truncation */
            objP = IPAddrObjFromDWORD(*(DWORD UNALIGNED *)bytesP);
            eaten = sizeof(DWORD);
            break;
            
        case 31: // objectipaddrv6
            if (remain < sizeof(IN6_ADDR))
                goto done;      /* Data truncation */
            objP = ObjFromIPv6Addr(bytesP, 0);
            eaten = sizeof(IN6_ADDR);
            break;

        case 32: // objectvariant
            eaten = *(ULONG UNALIGNED *)bytesP;
            if (remain < (eaten + sizeof(ULONG)))
                goto done;      /* Data truncation */
            objP = ObjFromByteArray(sizeof(ULONG)+bytesP, eaten);
            eaten += sizeof(ULONG);
            break;

        case 33: // objectsid
            if (remain < sizeof(ULONG))
                goto done;      /* Data truncation */
            if (*(ULONG UNALIGNED *)bytesP == 0) {
                /* Empty SID */
                eaten = sizeof(ULONG);
                objP = ObjFromEmptyString();
                break;
            }

            /* From MSDN -
             * A property with the Sid extension is actually a 
             * TOKEN_USER structure followed by the SID. The size
             * of the TOKEN_USER structure differs depending on 
             * whether the events were generated on a 32-bit or 
             * 64-bit architecture. Also the structure is aligned
             * on an 8-byte boundary, so its size is 8 bytes on a
             * 32-bit computer and 16 bytes on a 64-bit computer.
             * Doubling the pointer size handles both cases.
             *
             * Note that the ULONG * check above is part of this
             * structure so don't double count that and move bytesP
             * by another 4 bytes.
             */
            eaten = pointer_size * 2; /* Skip TOKEN_USER */
            remain -= eaten;
            bytesP += eaten;
            if (remain < sizeof(SID))
                goto done;      /* Truncation. Note sizeof(sid) is MIN size */
            MoveMemory(u.buf, bytesP,
                       remain > sizeof(u.buf) ? sizeof(u.buf) : remain);
            /* Sanity check before passing to SID converter */
            if (u.sid.Revision != SID_REVISION ||
                u.sid.SubAuthorityCount > SID_MAX_SUB_AUTHORITIES)
                goto done;

            eaten = TWAPI_SID_LENGTH(&u.sid);
            if (eaten > sizeof(u.buf) || eaten > remain)
                goto done;      /* Bad SID length */

            objP = ObjFromSIDNoFail(&u.sid);

            /* We have adjusted remain and bytesP for TOKEN_USER. eaten
             * contains actual SID length only without TOKEN_USER. SO
             * we are all set up for top of loop. Note as an aside
             * that not all bytes that were copied to u.buf[]
             * were necessarily used up
             */
            break;

        case 34: // uint64wmitime - TBD
        case 35: // objectwmitime - TBD
            if (remain < sizeof(ULONGLONG))
                goto done;      /* Data truncation */
            objP = ObjFromULONGLONG(*(ULONGLONG UNALIGNED *)bytesP);
            eaten = sizeof(ULONGLONG);
            break;

        case 38: // uint16port
        case 39: // objectport
            if (remain < sizeof(short))
                goto done;      /* Data truncation */
            port = *(short UNALIGNED *)bytesP;
            objP = ObjFromInt((WORD)port);
            eaten = sizeof(short);
            break;

        case 40: // datetime - TBD
            goto done;          /* TBD */
            break;
            
        case 41: // stringnotcounted
            objP = ObjFromStringN(bytesP, remain);
            eaten = remain;
            break;

        case 42: // wstringnotcounted
            objP = ObjFromUnicodeN((WCHAR *)bytesP, remain/sizeof(WCHAR));
            eaten = remain;
            break;

        case 43: // pointer
            if (pointer_size == 8)
                objP = ObjFromULONGLONGHex(*(ULONGLONG UNALIGNED *)bytesP);
            else
                objP = ObjFromULONGHex(*(ULONG UNALIGNED *)bytesP);
            eaten = pointer_size;
            break;

        default:
            TwapiSetObjResult(interp, Tcl_ObjPrintf("Internal error: unknown mof typeenum %d", typeenum));
            goto error_handler;
        }

        /* Some calls may result in objP being NULL */
        if (objP == NULL)
            objP = ObjFromEmptyString();
        Tcl_DictObjPut(interp, resultObj, types[i], objP);
    }
    
    done:
    TwapiSetObjResult(interp, resultObj);
    return TCL_OK;

error_handler:
    /* Tcl interp should already hold error code and result */

    if (resultObj)
        Tcl_DecrRefCount(resultObj);
    return TCL_ERROR;
}


static int Twapi_ETWCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    Tcl_Obj *objP;
    int func = PtrToInt(clientdata);

    objP = NULL;
    switch (func) {
    case 1: // etw_provider_enable_flags
        objP = ObjFromLong(gETWProviderTraceEnableFlags);
        break;
    case 2: // etw_provider_enable_level
        objP = ObjFromLong(gETWProviderTraceEnableLevel);
        break;
    case 3: // etw_provider_enabled
        objP = ObjFromBoolean((HANDLE)gETWProviderSessionHandle != INVALID_HANDLE_VALUE);
        break;
    default:
        return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);
    }

    TwapiSetObjResult(interp, objP);
    return TCL_OK;
}


static int TwapiETWInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    struct tcl_dispatch_s EtwDispatch[] = {
        DEFINE_TCL_CMD(StartTrace, Twapi_StartTrace),
        DEFINE_TCL_CMD(ControlTrace, Twapi_ControlTrace),
        DEFINE_TCL_CMD(etw_enable_trace, Twapi_EnableTrace), //EnableTrace
        DEFINE_TCL_CMD(OpenTrace, Twapi_OpenTrace),
        DEFINE_TCL_CMD(CloseTrace, Twapi_CloseTrace),
        DEFINE_TCL_CMD(ProcessTrace, Twapi_ProcessTrace),
        DEFINE_TCL_CMD(RegisterTraceGuids, Twapi_RegisterTraceGuids),
        DEFINE_TCL_CMD(etw_unregister_provider, Twapi_UnregisterTraceGuids), //UnregisterTraceGuids
        DEFINE_TCL_CMD(TraceEvent, Twapi_TraceEvent),
        DEFINE_TCL_CMD(Twapi, Twapi_ParseEventMofData),
    };

    struct fncode_dispatch_s EtwCallDispatch[] = {
        DEFINE_FNCODE_CMD(etw_provider_enable_flags, 1),     /* TBD docs */
        DEFINE_FNCODE_CMD(etw_provider_enable_level, 2),     /* TBD docs */
        DEFINE_FNCODE_CMD(etw_provider_enabled, 3),          /* TBD docs */
    };

    TwapiDefineTclCmds(interp, ARRAYSIZE(EtwDispatch), EtwDispatch, NULL);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(EtwCallDispatch), EtwCallDispatch, Twapi_ETWCallObjCmd);


    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::ETWCall", Twapi_ETWCallObjCmd, ticP, NULL);

    return TCL_OK;
}

static int ETWModuleOneTimeInit(Tcl_Interp *interp)
{
    InitializeCriticalSection(&gETWCS);
    return TCL_OK;
}

/* Called when interp is deleted */
static void TwapiETWCleanup(TwapiInterpContext *ticP)
{
    /* TBD - should we unregister providers or close sessions ? */
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
int Twapi_etw_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiETWInitCalls,
        TwapiETWCleanup
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    /* Init unless already done. */
    if (! TwapiDoOneTimeInit(&gETWInitialized, ETWModuleOneTimeInit, interp))
        return TCL_ERROR;

    /* NEW_TIC since we have a cleanup routine */
    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, NEW_TIC) ? TCL_OK : TCL_ERROR;
}

