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
 * TBD - TraceMessage
 */

#include "twapi.h"
#include <evntrace.h>

#include <ntverp.h>             /* Needed for VER_PRODUCTBUILD SDK version */

#if (VER_PRODUCTBUILD < 7600) || (_WIN32_WINNT <= 0x600)
# define RUNTIME_TDH_LOAD 1
#else
# include <tdh.h>
#pragma comment(lib, "delayimp.lib") /* Prevents TDH from loading unless necessary */
#pragma comment(lib, "tdh.lib")	 /* New TDH library for Vista and beyond */
#endif

#ifdef RUNTIME_TDH_LOAD

#define EVENT_HEADER_FLAG_EXTENDED_INFO         0x0001
#define EVENT_HEADER_FLAG_PRIVATE_SESSION       0x0002
#define EVENT_HEADER_FLAG_STRING_ONLY           0x0004
#define EVENT_HEADER_FLAG_TRACE_MESSAGE         0x0008
#define EVENT_HEADER_FLAG_NO_CPUTIME            0x0010
#define EVENT_HEADER_FLAG_32_BIT_HEADER         0x0020
#define EVENT_HEADER_FLAG_64_BIT_HEADER         0x0040
#define EVENT_HEADER_FLAG_CLASSIC_HEADER        0x0100

#define EVENT_HEADER_EXT_TYPE_RELATED_ACTIVITYID   0x0001
#define EVENT_HEADER_EXT_TYPE_SID                  0x0002
#define EVENT_HEADER_EXT_TYPE_TS_ID                0x0003
#define EVENT_HEADER_EXT_TYPE_INSTANCE_INFO        0x0004
#define EVENT_HEADER_EXT_TYPE_STACK_TRACE32        0x0005
#define EVENT_HEADER_EXT_TYPE_STACK_TRACE64        0x0006
#define EVENT_HEADER_EXT_TYPE_MAX                  0x0007

typedef struct _EVENT_DESCRIPTOR {

    USHORT      Id;
    UCHAR       Version;
    UCHAR       Channel;
    UCHAR       Level;
    UCHAR       Opcode;
    USHORT      Task;
    ULONGLONG   Keyword;

} EVENT_DESCRIPTOR, *PEVENT_DESCRIPTOR;

typedef const EVENT_DESCRIPTOR *PCEVENT_DESCRIPTOR;

typedef struct _EVENT_HEADER {
    USHORT Size;
    USHORT HeaderType;
    USHORT Flags;
    USHORT EventProperty;
    ULONG ThreadId;
    ULONG ProcessId;
    LARGE_INTEGER TimeStamp;
    GUID ProviderId;
    EVENT_DESCRIPTOR EventDescriptor;
    union {
        struct {
            ULONG KernelTime;
            ULONG UserTime;
        };
        ULONG64 ProcessorTime;
    };
    GUID ActivityId;

} EVENT_HEADER, *PEVENT_HEADER;

typedef struct _ETW_BUFFER_CONTEXT {
    UCHAR ProcessorNumber;
    UCHAR Alignment;
    USHORT LoggerId;
} ETW_BUFFER_CONTEXT, *PETW_BUFFER_CONTEXT;

typedef struct _EVENT_EXTENDED_ITEM_TS_ID {
  ULONG SessionId;
} EVENT_EXTENDED_ITEM_TS_ID, *PEVENT_EXTENDED_ITEM_TS_ID;

typedef struct _EVENT_EXTENDED_ITEM_RELATED_ACTIVITYID {
  GUID RelatedActivityId;
} EVENT_EXTENDED_ITEM_RELATED_ACTIVITYID, *PEVENT_EXTENDED_ITEM_RELATED_ACTIVITYID;

typedef struct _EVENT_EXTENDED_ITEM_INSTANCE {
  ULONG InstanceId;
  ULONG ParentInstanceId;
  GUID  ParentGuid;
} EVENT_EXTENDED_ITEM_INSTANCE, *PEVENT_EXTENDED_ITEM_INSTANCE;

typedef struct _EVENT_HEADER_EXTENDED_DATA_ITEM {
    USHORT Reserved1;
    USHORT ExtType;
    struct {
        USHORT Linkage :  1;
        USHORT Reserved2 : 15;
    };
    USHORT DataSize;
    ULONGLONG  DataPtr;

} EVENT_HEADER_EXTENDED_DATA_ITEM, *PEVENT_HEADER_EXTENDED_DATA_ITEM;

typedef struct _EVENT_RECORD {

    EVENT_HEADER EventHeader;
    ETW_BUFFER_CONTEXT BufferContext;
    USHORT ExtendedDataCount;
    USHORT UserDataLength;
    PEVENT_HEADER_EXTENDED_DATA_ITEM ExtendedData;
    PVOID UserData;
    PVOID UserContext;
} EVENT_RECORD, *PEVENT_RECORD;

#define EVENT_ENABLE_PROPERTY_SID                   0x00000001
#define EVENT_ENABLE_PROPERTY_TS_ID                 0x00000002
#define EVENT_ENABLE_PROPERTY_STACK_TRACE           0x00000004

#define PROCESS_TRACE_MODE_REAL_TIME                0x00000100
#define PROCESS_TRACE_MODE_RAW_TIMESTAMP            0x00001000
#define PROCESS_TRACE_MODE_EVENT_RECORD             0x10000000

typedef enum _TDH_CONTEXT_TYPE { 
    TDH_CONTEXT_WPP_TMFFILE = 0,
    TDH_CONTEXT_WPP_TMFSEARCHPATH = 1,
    TDH_CONTEXT_WPP_GMT = 2,
    TDH_CONTEXT_POINTERSIZE = 3,
    TDH_CONTEXT_PDB_PATH = 4,
    TDH_CONTEXT_MAXIMUM = 5
} TDH_CONTEXT_TYPE;

typedef struct _TDH_CONTEXT {
    ULONGLONG ParameterValue;
    TDH_CONTEXT_TYPE ParameterType;
    ULONG ParameterSize;
} TDH_CONTEXT;

typedef enum _DECODING_SOURCE { 
  DecodingSourceXMLFile  = 0,
  DecodingSourceWbem     = 1,
  DecodingSourceWPP      = 2
} DECODING_SOURCE;

typedef enum _TEMPLATE_FLAGS
{
    TEMPLATE_EVENT_DATA = 1,
    TEMPLATE_USER_DATA = 2
} TEMPLATE_FLAGS;

typedef enum _PROPERTY_FLAGS
{
   PropertyStruct        = 0x1,      // Type is struct.
   PropertyParamLength   = 0x2,      // Length field is index of param with length.
   PropertyParamCount    = 0x4,      // Count file is index of param with count.
   PropertyWBEMXmlFragment = 0x8,    // WBEM extension flag for property.
   PropertyParamFixedLength = 0x10   // Length of the parameter is fixed.
} PROPERTY_FLAGS;

typedef struct _EVENT_PROPERTY_INFO {
    PROPERTY_FLAGS Flags;
    ULONG NameOffset;
    union {
        struct _nonStructType {
            USHORT InType;
            USHORT OutType;
            ULONG MapNameOffset;
        } nonStructType;
        struct _structType {
            USHORT StructStartIndex;
            USHORT NumOfStructMembers;
            ULONG padding;
        } structType;
    };
    union {
        USHORT count;
        USHORT countPropertyIndex;
    };
    union {
        USHORT length;
        USHORT lengthPropertyIndex;
    };
    ULONG Reserved;
} EVENT_PROPERTY_INFO;
typedef EVENT_PROPERTY_INFO *PEVENT_PROPERTY_INFO;

typedef struct _TRACE_EVENT_INFO {
  GUID ProviderGuid;
  GUID EventGuid;
  EVENT_DESCRIPTOR EventDescriptor;
  DECODING_SOURCE DecodingSource;
  ULONG ProviderNameOffset;
  ULONG LevelNameOffset;
  ULONG ChannelNameOffset;
  ULONG KeywordsNameOffset;
  ULONG TaskNameOffset;
  ULONG OpcodeNameOffset;
  ULONG EventMessageOffset;
  ULONG ProviderMessageOffset;
  ULONG BinaryXMLOffset;
  ULONG BinaryXMLSize;
  ULONG ActivityIDNameOffset;
  ULONG RelatedActivityIDNameOffset;
  ULONG PropertyCount;
  ULONG TopLevelPropertyCount;
  TEMPLATE_FLAGS Flags;
  EVENT_PROPERTY_INFO EventPropertyInfoArray[ANYSIZE_ARRAY];
} TRACE_EVENT_INFO;

typedef ULONG __stdcall TdhGetEventInformation_T(PEVENT_RECORD, ULONG, TDH_CONTEXT *, TRACE_EVENT_INFO *, ULONG *);
static struct {
    TdhGetEventInformation_T  *_TdhGetEventInformation;
} gTdhStubs;

#define TdhGetEventInformation gTdhStubs._TdhGetEventInformation

int gTdhStatus;                 /* 0 - init, 1 - available, -1 - unavailable  */
HANDLE gTdhDllHandle;

#endif

/* Max length of file and session name fields in a trace (documented in SDK) */
#define MAX_TRACE_NAME_CHARS (1024+1)

#define ObjFromTRACEHANDLE(val_) ObjFromWideInt(val_)
#define ObjToTRACEHANDLE ObjToWideInt

static TRACEHANDLE gInvalidTraceHandle;
#define INVALID_SESSIONTRACE_HANDLE(ht_) ((ht_) == gInvalidTraceHandle)

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
    "-maxfilesize",
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
    "-maxfilesize",
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
    TwapiInterpContext *ticP;

    /* If a callback is supplied, buffer_cmdObj holds it and buffer_cmdlen
       is non-0. If no callback is supplied, we collect the buffer descriptor
       and events list in buffer_listObj and return it to caller.
    */
    union {
        Tcl_Obj *cmdObj;     /* Callback for buffers */
        Tcl_Obj *listObj;    /* List of buffer descriptor, event list pairs */
    } buffer;

    /*
     * when non-NULL, an ObjIncrRefs must have been done on eventsObj
     * and errorObj with a corresponding ObjDecrRefs when setting to NULL
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
    {"-hdr_maxfilesize"},
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
    "-maxfilesize",
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
        ObjSetStaticResult(interp, "Invalid EVENT_TRACE_PROPERTIES format");
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

    etP = TwapiAllocZero(buf_sz);
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
            ObjSetStaticResult(interp, "Internal error: Unexpected field index.");
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
        ObjSetStaticResult(interp, "NULL provider GUID specified.");
        return TCL_ERROR;
    }

    /* We should not already have registered a different provider */
    if (! IsEqualGUID(&gETWProviderGuid, &gNullGuid)) {
        if (IsEqualGUID(&gETWProviderGuid, &provider_guid)) {
            ObjSetResult(interp, ObjFromTRACEHANDLE(gETWProviderRegistrationHandle));

            return TCL_OK;      /* Same GUID - ok */
        }
        else {
            ObjSetStaticResult(interp, "ETW Provider GUID already registered");
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
        ObjSetResult(interp, ObjFromTRACEHANDLE(gETWProviderRegistrationHandle));
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
        ObjSetStaticResult(interp, "Unknown ETW provider registration handle");
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
    ObjSetResult(interp, ObjNewList(2, objs));
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
        ObjSetResult(interp, ObjFromEVENT_TRACE_PROPERTIES(etP));
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
    TWAPI_ASSERT(gETWContext.ticP != NULL);
    TWAPI_ASSERT(gETWContext.ticP->interp != NULL);
    interp = gETWContext.ticP->interp;

    if (gETWContext.errorObj)   /* If some previous error occurred, return */
        return;

    /* TBD - create as a list - faster than as a dict */
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
}

/* Uses memlifo frame. Caller responsible for cleanup */
static DWORD TwapiTdhGetEventInformation(TwapiInterpContext *ticP, PEVENT_RECORD recP, TRACE_EVENT_INFO **teiPP)
{
    DWORD sz, winerr;
    TRACE_EVENT_INFO *teiP;

    /* TBD - instrument how much to try for initially */
    teiP = MemLifoAlloc(&ticP->memlifo, 1000, &sz);

    /* TBD - initialize TDH_CONTEXT param to include pointer size indicator */
    winerr = TdhGetEventInformation(recP, 0, NULL, teiP, &sz);
    if (winerr == ERROR_INSUFFICIENT_BUFFER) {
        teiP = MemLifoAlloc(&ticP->memlifo, sz, &sz);
        winerr = TdhGetEventInformation(recP, 0, NULL, teiP, &sz);
    }
    
    /* We may have over allocated so shrink down before returning */
    if (winerr == ERROR_SUCCESS) {
        *teiPP = MemLifoResizeLast(&ticP->memlifo, sz, 1);
    }

    return winerr;
}

static Tcl_Obj *TwapiTEIUnicodeObj(TRACE_EVENT_INFO *teiP, int offset)
{
    if (offset == 0)
        return ObjFromEmptyString();
    else
        return ObjFromUnicode((WCHAR*) (offset + (char*)teiP));
}

static Tcl_Obj *ObjFromEVENT_DESCRIPTOR(EVENT_DESCRIPTOR *evdP)
{
    Tcl_Obj *objs[7];

    objs[0] = ObjFromLong(evdP->Id);
    objs[1] = ObjFromLong(evdP->Version);
    objs[2] = ObjFromLong(evdP->Channel);
    objs[3] = ObjFromLong(evdP->Level);
    objs[4] = ObjFromLong(evdP->Opcode);
    objs[5] = ObjFromLong(evdP->Task);
    objs[6] = ObjFromULONGLONG(evdP->Keyword);

    return ObjNewList(ARRAYSIZE(objs), objs);
}

static Tcl_Obj *ObjFromEVENT_HEADER(EVENT_HEADER *evhP)
{
    Tcl_Obj *objs[11];

    objs[0] = ObjFromLong(evhP->Flags);
    objs[1] = ObjFromLong(evhP->EventProperty);
    objs[2] = ObjFromLong(evhP->ThreadId);
    objs[3] = ObjFromLong(evhP->ProcessId);
    objs[4] = ObjFromLARGE_INTEGER(evhP->TimeStamp);
    objs[5] = ObjFromGUID(&evhP->ProviderId);
    objs[6] = ObjFromEVENT_DESCRIPTOR(&evhP->EventDescriptor);

    /*
     * Note - for user mode sessions, KernelTime/UserTime are not valid
     * and the ProcessorTime member has to be used instead. However,
     * we do not know the type of session at this point so we leave it
     * to the app to figure out what to use
     */
    objs[7] = ObjFromULONG(evhP->KernelTime);
    objs[8] = ObjFromULONG(evhP->UserTime);
    objs[9] = ObjFromULONGLONG(evhP->ProcessorTime);

    objs[10] = ObjFromGUID(&evhP->ActivityId);
    
    return ObjNewList(ARRAYSIZE(objs), objs);
}

static Tcl_Obj *ObjFromTRACE_EVENT_INFO(TRACE_EVENT_INFO *teiP)
{
    Tcl_Obj *objs[16];

    objs[0] = ObjFromGUID(&teiP->ProviderGuid);
    objs[1] = ObjFromGUID(&teiP->EventGuid);
    /* TBD - does this duplicate EventDescriptor struct in event header ? */
    objs[2] = ObjFromEVENT_DESCRIPTOR(&teiP->EventDescriptor);

    objs[3] = ObjFromLong(teiP->DecodingSource);
    objs[4] = TwapiTEIUnicodeObj(teiP, teiP->ProviderNameOffset);
    objs[5] = TwapiTEIUnicodeObj(teiP, teiP->LevelNameOffset);
    objs[6] = TwapiTEIUnicodeObj(teiP, teiP->ChannelNameOffset);
    objs[7] = TwapiTEIUnicodeObj(teiP, teiP->KeywordsNameOffset);
    objs[8] = TwapiTEIUnicodeObj(teiP, teiP->TaskNameOffset);
    objs[9] = TwapiTEIUnicodeObj(teiP, teiP->OpcodeNameOffset);
    objs[10] = TwapiTEIUnicodeObj(teiP, teiP->EventMessageOffset);
    objs[11] = TwapiTEIUnicodeObj(teiP, teiP->ProviderMessageOffset);
    objs[12] = TwapiTEIUnicodeObj(teiP, teiP->ActivityIDNameOffset);
    objs[13] = TwapiTEIUnicodeObj(teiP, teiP->RelatedActivityIDNameOffset);
    objs[14] = ObjFromLong(teiP->PropertyCount);
    objs[15] = ObjFromLong(teiP->TopLevelPropertyCount);

    return ObjNewList(ARRAYSIZE(objs), objs);
}


static VOID WINAPI TwapiETWEventRecordCallback(PEVENT_RECORD recP)
{
    TwapiInterpContext *ticP;
    int i;
    Tcl_Obj *recObjs[4];
    Tcl_Obj *objs[20];
    TRACE_EVENT_INFO *teiP;
    DWORD winerr;
    MemLifoMarkHandle mark;

    /* Called back from Win32 ProcessTrace call. Assumed that gETWContext is locked */
    TWAPI_ASSERT(gETWContext.ticP != NULL);
    TWAPI_ASSERT(gETWContext.ticP->interp != NULL);
    ticP = gETWContext.ticP;

    if (gETWContext.errorObj)   /* If some previous error occurred, return */
        return;

    if ((recP->EventHeader.Flags & EVENT_HEADER_FLAG_TRACE_MESSAGE) != 0)
        return; // Ignore WPP events. - TBD
    
    mark = MemLifoPushMark(&ticP->memlifo);

    winerr = TwapiTdhGetEventInformation(ticP, recP, &teiP);
    if (winerr != ERROR_SUCCESS) {
        gETWContext.errorObj = Twapi_MakeWindowsErrorCodeObj(winerr, NULL);
        ObjIncrRefs(gETWContext.errorObj);
        MemLifoPopMark(mark);
        return;
    }

    recObjs[0] = ObjFromEVENT_HEADER(&recP->EventHeader);

    /* Buffer Context */
    objs[0] = ObjFromLong(recP->BufferContext.ProcessorNumber);
    objs[1] = ObjFromLong(recP->BufferContext.Alignment);
    objs[2] = ObjFromLong(recP->BufferContext.LoggerId);
    recObjs[1] = ObjNewList(3, objs);
    
    /* Extended Data */
    recObjs[2] = ObjNewList(0, NULL);
    for (i = 0; i < recP->ExtendedDataCount; ++i) {
        EVENT_HEADER_EXTENDED_DATA_ITEM *ehdrP = &recP->ExtendedData[i];
        if (ehdrP->DataPtr == 0)
            continue;
        switch (ehdrP->ExtType) {
        case EVENT_HEADER_EXT_TYPE_RELATED_ACTIVITYID:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("-relatedactivity"));
            ObjAppendElement(NULL, recObjs[2], ObjFromGUID(&((PEVENT_EXTENDED_ITEM_RELATED_ACTIVITYID)ehdrP->DataPtr)->RelatedActivityId));
            break;
        case EVENT_HEADER_EXT_TYPE_SID:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("-sid"));
            ObjAppendElement(NULL, recObjs[2], ObjFromSIDNoFail((SID *)(ehdrP->DataPtr)));
            break;
        case EVENT_HEADER_EXT_TYPE_TS_ID:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("-tssession"));
            ObjAppendElement(NULL, recObjs[2], ObjFromULONG(((PEVENT_EXTENDED_ITEM_TS_ID)ehdrP->DataPtr)->SessionId));
            break;
        case EVENT_HEADER_EXT_TYPE_INSTANCE_INFO:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("-iteminstance"));
            objs[0] = ObjFromULONG(((PEVENT_EXTENDED_ITEM_INSTANCE)ehdrP->DataPtr)->InstanceId);
            objs[1] = ObjFromULONG(((PEVENT_EXTENDED_ITEM_INSTANCE)ehdrP->DataPtr)->ParentInstanceId);
            objs[2] = ObjFromGUID(&((PEVENT_EXTENDED_ITEM_INSTANCE)ehdrP->DataPtr)->ParentGuid);
            ObjAppendElement(NULL, recObjs[2], ObjNewList(3, objs));
            break;

        case EVENT_HEADER_EXT_TYPE_STACK_TRACE32:
        case EVENT_HEADER_EXT_TYPE_STACK_TRACE64:
        default:
            break;              /* Skip/ignore */
        }
    }
    
    recObjs[3] = ObjFromTRACE_EVENT_INFO(teiP);
    ObjAppendElement(ticP->interp, gETWContext.eventsObj, ObjNewList(ARRAYSIZE(recObjs), recObjs));

    MemLifoPopMark(mark);
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
    TWAPI_ASSERT(gETWContext.ticP != NULL);
    TWAPI_ASSERT(gETWContext.ticP->interp != NULL);
    interp = gETWContext.ticP->interp;

    if (Tcl_InterpDeleted(interp))
        return FALSE;

    if (gETWContext.errorObj)   /* If some previous error occurred, return */
        return FALSE;

    TWAPI_ASSERT(gETWContext.eventsObj); /* since errorObj not NULL at this point */

    if (gETWContext.buffer_cmdlen == 0) {
        /* We are simply collecting events without invoking callback */
        TWAPI_ASSERT(gETWContext.buffer.listObj != NULL);
    } else {
        /*
         * Construct a command to call with the event. 
         * gETWContext.buffer_cmdObj could be a shared object, either
         * initially itself or result in a shared object in the callback.
         * So we need to check for that and Dup it if necessary
         */

        if (Tcl_IsShared(gETWContext.buffer.cmdObj)) {
            objP = ObjDuplicate(gETWContext.buffer.cmdObj);
            ObjIncrRefs(objP);
        } else
            objP = gETWContext.buffer.cmdObj;
    }

    /* TBD - create as a list - faster than as a dict */
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
     * in the returned dictionary. - TBD
    */

    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[24].keyObj, ObjFromULONG(adjustedP->BuffersLost));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[25].keyObj, ObjFromTIME_ZONE_INFORMATION(&adjustedP->TimeZone));    
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[26].keyObj, ObjFromLARGE_INTEGER(adjustedP->BootTime));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[27].keyObj, ObjFromLARGE_INTEGER(adjustedP->PerfFreq));
    Tcl_DictObjPut(interp, bufObj, gETWBufferKeys[28].keyObj, ObjFromLARGE_INTEGER(adjustedP->StartTime));

    /*
     * Note: Do not need to ObjIncrRefs bufObj[] because we are adding
     * it to the command list
     */
    if (gETWContext.buffer_cmdlen) {
        args[0] = bufObj;
        args[1] = gETWContext.eventsObj;
        Tcl_ListObjReplace(interp, objP, gETWContext.buffer_cmdlen, ARRAYSIZE(args), ARRAYSIZE(args), args);
        code = Tcl_EvalObjEx(gETWContext.ticP->interp, objP, TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);

        /* Get rid of the command obj if we created it */
        if (objP != gETWContext.buffer.cmdObj)
            ObjDecrRefs(objP);
    } else {
        /* No callback. Just collect */
        ObjAppendElement(NULL, gETWContext.buffer.listObj, bufObj);
        ObjAppendElement(NULL, gETWContext.buffer.listObj, gETWContext.eventsObj);
        code = TCL_OK;
    }

    ObjDecrRefs(gETWContext.eventsObj);
    gETWContext.eventsObj = ObjNewList(0, NULL);/* For next set of events */
    ObjIncrRefs(gETWContext.eventsObj);

    switch (code) {
    case TCL_BREAK:
        /* Any other value - not an error, but stop processing */
        return FALSE;
    case TCL_ERROR:
        gETWContext.errorObj = Tcl_GetReturnOptions(interp, code);
        ObjIncrRefs(gETWContext.errorObj);
        ObjDecrRefs(gETWContext.eventsObj);
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
                     ARGSKIP, GETINT(real_time),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    s = ObjToUnicode(objv[1]);
    ZeroMemory(&etl, sizeof(etl));
    etl.BufferCallback = TwapiETWBufferCallback;
    /*
     * To support older compilers/SDK's, we use
     *   EventCallback instead of EventRecordCallback
     *   LogFileMode instead of ProcessTraceMode
     *   EVENT_TRACE_REAL_TIME_MODE instead of PROCESS_TRACE_MODE_REAL_TIME
     * They are either the same field in the union, or have the same #define value
     */
    if (gTdhStatus > 0) {
        etl.LogFileMode = PROCESS_TRACE_MODE_EVENT_RECORD; /* etl.ProcessTraceMode = */
        etl.EventCallback = (PEVENT_CALLBACK) TwapiETWEventRecordCallback;   /* etl.EventRecordCallback = */
    } else
        etl.EventCallback = TwapiETWEventCallback;

    if (real_time) {
        etl.LoggerName = s;
        etl.LogFileMode |= EVENT_TRACE_REAL_TIME_MODE; /* or etl.ProcessTraceMode |= PROCESS_TRACE_MODE_REAL_TIME */
    } else
        etl.LogFileName = s;


    htrace = OpenTraceW(&etl);
    if (INVALID_SESSIONTRACE_HANDLE(htrace))
        return TwapiReturnSystemError(interp);

    ObjSetResult(interp, ObjFromTRACEHANDLE(htrace));
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


TCL_RESULT Twapi_ProcessTrace(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
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

    /* Verify callback command prefix is a list. If empty, data
     * is returned instead.
     */
    if (ObjListLength(interp, objv[2], &buffer_cmdlen) != TCL_OK)
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
    
    if (gETWContext.ticP != NULL) {
        LeaveCriticalSection(&gETWCS);
        ObjSetStaticResult(interp, "Recursive call to ProcessTrace");
        return TCL_ERROR;
    }

    gETWContext.traceH = htraces[0];
    gETWContext.buffer_cmdlen = buffer_cmdlen;
    if (buffer_cmdlen)
        gETWContext.buffer.cmdObj = objv[2];
    else
        gETWContext.buffer.listObj = ObjNewList(0, NULL);
    gETWContext.eventsObj = ObjNewList(0, NULL);
    ObjIncrRefs(gETWContext.eventsObj);
    gETWContext.errorObj = NULL;
    gETWContext.ticP = ticP;

    /* TBD - make these per interpreter keys */
    /*
     * Initialize the dictionary objects we will use as keys. These are
     * tied to interp so have to redo on every call since the interp
     * may not be the same.
     */
    for (i = 0; i < ARRAYSIZE(gETWEventKeys); ++i) {
        gETWEventKeys[i].keyObj = ObjFromString(gETWEventKeys[i].keystring);
        ObjIncrRefs(gETWEventKeys[i].keyObj);
    }
    for (i = 0; i < ARRAYSIZE(gETWBufferKeys); ++i) {
        gETWBufferKeys[i].keyObj = ObjFromString(gETWBufferKeys[i].keystring);
        ObjIncrRefs(gETWBufferKeys[i].keyObj);
    }

    winerr = ProcessTrace(htraces, ntraces, startP, endP);

    /* Copy and reset context before unlocking */
    etwc = gETWContext;
    gETWContext.buffer.cmdObj = NULL;
    gETWContext.eventsObj = NULL;
    gETWContext.errorObj = NULL;
    gETWContext.ticP = NULL;

    LeaveCriticalSection(&gETWCS);

    /* Deallocated the cached key objects */
    for (i = 0; i < ARRAYSIZE(gETWEventKeys); ++i) {
        ObjDecrRefs(gETWEventKeys[i].keyObj);
        gETWEventKeys[i].keyObj = NULL; /* Just to catch bad access */
    }
    for (i = 0; i < ARRAYSIZE(gETWBufferKeys); ++i) {
        ObjDecrRefs(gETWBufferKeys[i].keyObj);
        gETWBufferKeys[i].keyObj = NULL; /* Just to catch bad access */
    }

    if (etwc.eventsObj)
        ObjDecrRefs(etwc.eventsObj);

    /* If the processing was successful, errorObj is NULL */
    if (etwc.errorObj) {
        code = Tcl_SetReturnOptions(interp, etwc.errorObj);
        ObjDecrRefs(etwc.errorObj); /* Match one in the event callback */
        if (etwc.buffer_cmdlen == 0)
            ObjDecrRefs(etwc.buffer.listObj);
        return code;
    }

    /* A winerr of ERROR_CANCELLED means the callback returned TCL_BREAK
     * to terminate the processing. That is not treated as an error
     */
    if (winerr && winerr != ERROR_CANCELLED)
        return Twapi_AppendSystemError(interp, winerr);

    if (etwc.buffer_cmdlen == 0) {
        /* No callback so return collected events */
        ObjSetResult(interp, etwc.buffer.listObj);
    } else
        Tcl_ResetResult(interp); /* For any holdover from callbacks */

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

    /* TBD - objv[1] and other objv[] may be samve and shimmer.
       Fix by using TwapiGetArgsEx
    */
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

        /* IMPORTANT: switch values based on _etw_decipher_mof_event_field_type proc */
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
            ObjSetResult(interp, Tcl_ObjPrintf("Internal error: unknown mof typeenum %d", typeenum));
            goto error_handler;
        }

        /* Some calls may result in objP being NULL */
        if (objP == NULL)
            objP = ObjFromEmptyString();
        Tcl_DictObjPut(interp, resultObj, types[i], objP);
    }
    
    done:
    ObjSetResult(interp, resultObj);
    return TCL_OK;

error_handler:
    /* Tcl interp should already hold error code and result */

    if (resultObj)
        ObjDecrRefs(resultObj);
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

    ObjSetResult(interp, objP);
    return TCL_OK;
}


static int TwapiETWInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    struct tcl_dispatch_s EtwDispatch[] = {
        DEFINE_TCL_CMD(StartTrace, Twapi_StartTrace),
        DEFINE_TCL_CMD(ControlTrace, Twapi_ControlTrace),
        DEFINE_TCL_CMD(EnableTrace, Twapi_EnableTrace),
        DEFINE_TCL_CMD(OpenTrace, Twapi_OpenTrace),
        DEFINE_TCL_CMD(CloseTrace, Twapi_CloseTrace),
        DEFINE_TCL_CMD(ProcessTrace, Twapi_ProcessTrace),
        DEFINE_TCL_CMD(RegisterTraceGuids, Twapi_RegisterTraceGuids),
        DEFINE_TCL_CMD(etw_unregister_provider, Twapi_UnregisterTraceGuids), //UnregisterTraceGuids
        DEFINE_TCL_CMD(TraceEvent, Twapi_TraceEvent),
        DEFINE_TCL_CMD(Twapi_ParseEventMofData, Twapi_ParseEventMofData),
    };

    struct fncode_dispatch_s EtwCallDispatch[] = {
        DEFINE_FNCODE_CMD(etw_provider_enable_flags, 1),     /* TBD docs */
        DEFINE_FNCODE_CMD(etw_provider_enable_level, 2),     /* TBD docs */
        DEFINE_FNCODE_CMD(etw_provider_enabled, 3),          /* TBD docs */
    };

    TwapiDefineTclCmds(interp, ARRAYSIZE(EtwDispatch), EtwDispatch, ticP);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(EtwCallDispatch), EtwCallDispatch, Twapi_ETWCallObjCmd);


    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::ETWCall", Twapi_ETWCallObjCmd, ticP, NULL);

    return TCL_OK;
}

void TwapiInitTdhStubs(Tcl_Interp *interp)
{
#ifdef RUNTIME_TDH_LOAD
    UINT len;
    WCHAR path[MAX_PATH+1];

    if (gTdhStatus)
        return;                 /* Already called */

    /* Assume failure */
    gTdhStatus = -1;

    len = GetSystemDirectoryW(path, ARRAYSIZE(path));
    if (len == 0)
        return; /* Error ignored. Functions will simply not be available */

#define TDHDLLNAME L"\\tdh.dll"
    if (len > (ARRAYSIZE(path) - ARRAYSIZE(TDHDLLNAME)))
        return;

    lstrcpynW(&path[len], TDHDLLNAME, ARRAYSIZE(TDHDLLNAME));

    /* Initialize the Evt stubs */
    gTdhDllHandle = LoadLibraryW(path);
    if (gTdhDllHandle == NULL)
        return;                 /* DLL not available */
    
#define INIT_TDH_STUB(fn) \
    do { \
      if (((gTdhStubs._ ## fn) = (void *)GetProcAddress(gTdhDllHandle, #fn)) == NULL) \
            return;                                                     \
    } while (0)

    INIT_TDH_STUB(TdhGetEventInformation);

#endif

    gTdhStatus = 1;
}

static int ETWModuleOneTimeInit(Tcl_Interp *interp)
{
    TwapiInitTdhStubs(interp);

    /* Depends on OS - see documentation of OpenTrace in SDK */
    if (sizeof(void*) == 8) {
        gInvalidTraceHandle = 0xFFFFFFFFFFFFFFFF;
    } else {
        /* 32-bit */
        if (TwapiMinOSVersion(6, 0))
            gInvalidTraceHandle = 0x00000000FFFFFFFF;
        else
            gInvalidTraceHandle = 0xFFFFFFFFFFFFFFFF;
    }

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

