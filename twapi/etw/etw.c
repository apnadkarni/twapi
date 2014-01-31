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
#define INITGUID // To get EventTraceGuid defined
# include <tdh.h>
#pragma comment(lib, "delayimp.lib") /* Prevents TDH from loading unless necessary */
#pragma comment(lib, "tdh.lib")	 /* New TDH library for Vista and beyond */
#endif

#ifdef RUNTIME_TDH_LOAD

#define INITGUID // To get EventTraceGuid defined
#include <guiddef.h>
DEFINE_GUID ( /* 68fdd900-4a3e-11d1-84f4-0000f80464e3 */
    EventTraceGuid,
    0x68fdd900,
    0x4a3e,
    0x11d1,
    0x84, 0xf4, 0x00, 0x00, 0xf8, 0x04, 0x64, 0xe3
  );


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

typedef struct _PROPERTY_DATA_DESCRIPTOR {
    ULONGLONG PropertyName;
    ULONG ArrayIndex;
    ULONG Reserved;
} PROPERTY_DATA_DESCRIPTOR;
typedef PROPERTY_DATA_DESCRIPTOR *PPROPERTY_DATA_DESCRIPTOR;

typedef struct _EVENT_MAP_ENTRY {
    ULONG OutputOffset;
    union {
        ULONG Value;
        ULONG InputOffset;
    };
} EVENT_MAP_ENTRY;
typedef EVENT_MAP_ENTRY *PEVENT_MAP_ENTRY;

typedef enum _MAP_FLAGS {
    EVENTMAP_INFO_FLAG_MANIFEST_VALUEMAP   = 0x1,
    EVENTMAP_INFO_FLAG_MANIFEST_BITMAP     = 0x2,
    EVENTMAP_INFO_FLAG_MANIFEST_PATTERNMAP = 0x4,
    EVENTMAP_INFO_FLAG_WBEM_VALUEMAP       = 0x8,
    EVENTMAP_INFO_FLAG_WBEM_BITMAP         = 0x10,
    EVENTMAP_INFO_FLAG_WBEM_FLAG           = 0x20,
    EVENTMAP_INFO_FLAG_WBEM_NO_MAP         = 0x40
} MAP_FLAGS;

typedef enum _MAP_VALUETYPE {
    EVENTMAP_ENTRY_VALUETYPE_ULONG,
    EVENTMAP_ENTRY_VALUETYPE_STRING
}  MAP_VALUETYPE;

typedef struct _EVENT_MAP_INFO {
    ULONG NameOffset;
    MAP_FLAGS Flag;
    ULONG EntryCount;
    union {
        MAP_VALUETYPE MapEntryValueType;
        ULONG FormatStringOffset;
    };
    EVENT_MAP_ENTRY MapEntryArray[ANYSIZE_ARRAY];
} EVENT_MAP_INFO;
typedef EVENT_MAP_INFO *PEVENT_MAP_INFO;

enum _TDH_IN_TYPE {
    TDH_INTYPE_NULL,
    TDH_INTYPE_UNICODESTRING,
    TDH_INTYPE_ANSISTRING,
    TDH_INTYPE_INT8,
    TDH_INTYPE_UINT8,
    TDH_INTYPE_INT16,
    TDH_INTYPE_UINT16,
    TDH_INTYPE_INT32,
    TDH_INTYPE_UINT32,
    TDH_INTYPE_INT64,
    TDH_INTYPE_UINT64,
    TDH_INTYPE_FLOAT,
    TDH_INTYPE_DOUBLE,
    TDH_INTYPE_BOOLEAN,
    TDH_INTYPE_BINARY,
    TDH_INTYPE_GUID,
    TDH_INTYPE_POINTER,
    TDH_INTYPE_FILETIME,
    TDH_INTYPE_SYSTEMTIME,
    TDH_INTYPE_SID,
    TDH_INTYPE_HEXINT32,
    TDH_INTYPE_HEXINT64,                    // End of winmeta intypes.
    TDH_INTYPE_COUNTEDSTRING = 300,         // Start of TDH intypes for WBEM.
    TDH_INTYPE_COUNTEDANSISTRING,
    TDH_INTYPE_REVERSEDCOUNTEDSTRING,
    TDH_INTYPE_REVERSEDCOUNTEDANSISTRING,
    TDH_INTYPE_NONNULLTERMINATEDSTRING,
    TDH_INTYPE_NONNULLTERMINATEDANSISTRING,
    TDH_INTYPE_UNICODECHAR,
    TDH_INTYPE_ANSICHAR,
    TDH_INTYPE_SIZET,
    TDH_INTYPE_HEXDUMP,
    TDH_INTYPE_WBEMSID
};

enum _TDH_OUT_TYPE {
    TDH_OUTTYPE_NULL,
    TDH_OUTTYPE_STRING,
    TDH_OUTTYPE_DATETIME,
    TDH_OUTTYPE_BYTE,
    TDH_OUTTYPE_UNSIGNEDBYTE,
    TDH_OUTTYPE_SHORT,
    TDH_OUTTYPE_UNSIGNEDSHORT,
    TDH_OUTTYPE_INT,
    TDH_OUTTYPE_UNSIGNEDINT,
    TDH_OUTTYPE_LONG,
    TDH_OUTTYPE_UNSIGNEDLONG,
    TDH_OUTTYPE_FLOAT,
    TDH_OUTTYPE_DOUBLE,
    TDH_OUTTYPE_BOOLEAN,
    TDH_OUTTYPE_GUID,
    TDH_OUTTYPE_HEXBINARY,
    TDH_OUTTYPE_HEXINT8,
    TDH_OUTTYPE_HEXINT16,
    TDH_OUTTYPE_HEXINT32,
    TDH_OUTTYPE_HEXINT64,
    TDH_OUTTYPE_PID,
    TDH_OUTTYPE_TID,
    TDH_OUTTYPE_PORT,
    TDH_OUTTYPE_IPV4,
    TDH_OUTTYPE_IPV6,
    TDH_OUTTYPE_SOCKETADDRESS,
    TDH_OUTTYPE_CIMDATETIME,
    TDH_OUTTYPE_ETWTIME,
    TDH_OUTTYPE_XML,
    TDH_OUTTYPE_ERRORCODE,
    TDH_OUTTYPE_WIN32ERROR,
    TDH_OUTTYPE_NTSTATUS,
    TDH_OUTTYPE_HRESULT,             // End of winmeta outtypes.
    TDH_OUTTYPE_CULTURE_INSENSITIVE_DATETIME, //Culture neutral datetime string.
    TDH_OUTTYPE_REDUCEDSTRING = 300, // Start of TDH outtypes for WBEM.
    TDH_OUTTYPE_NOPRINT
};


typedef ULONG __stdcall TdhGetEventInformation_T(PEVENT_RECORD, ULONG, TDH_CONTEXT *, TRACE_EVENT_INFO *, ULONG *);
typedef ULONG __stdcall TdhGetProperty_T(PEVENT_RECORD pEvent, ULONG TdhContextCount, TDH_CONTEXT *pTdhContext, ULONG PropertyDataCount, PPROPERTY_DATA_DESCRIPTOR pPropertyData, ULONG BufferSize, PBYTE pBuffer);
typedef ULONG __stdcall TdhGetPropertySize_T(PEVENT_RECORD pEvent, ULONG TdhContextCount, TDH_CONTEXT *pTdhContext, ULONG PropertyDataCount, PPROPERTY_DATA_DESCRIPTOR pPropertyData, ULONG *pPropertySize);
typedef ULONG __stdcall TdhGetEventMapInformation_T(PEVENT_RECORD pEvent, LPWSTR pMapName, PEVENT_MAP_INFO pBuffer, ULONG *pBufferSize);
static struct {
    TdhGetEventInformation_T  *_TdhGetEventInformation;
    TdhGetProperty_T *_TdhGetProperty;
    TdhGetPropertySize_T *_TdhGetPropertySize;
    TdhGetEventMapInformation_T *_TdhGetEventMapInformation;
} gTdhStubs;

#define TdhGetEventInformation gTdhStubs._TdhGetEventInformation
#define TdhGetProperty gTdhStubs._TdhGetProperty
#define TdhGetPropertySize gTdhStubs._TdhGetPropertySize
#define TdhGetEventMapInformation gTdhStubs._TdhGetEventMapInformation

int gTdhStatus;                 /* 0 - init, 1 - available, -1 - unavailable  */
HANDLE gTdhDllHandle;

#endif

#ifndef EVENT_HEADER_FLAG_PROCESSOR_INDEX
# define EVENT_HEADER_FLAG_PROCESSOR_INDEX 0x0200 /* Win 8, not in older SDKs */
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
     * with a corresponding ObjDecrRefs when setting to NULL
      */
    Tcl_Obj *eventsObj;
    TCL_RESULT status;

    TRACEHANDLE traceH;

    int   buffer_cmdlen;        /* length of buffer_cmdObj */
    ULONG pointer_size;         /* Used if event itself does not specify */
    ULONG timer_resolution;
    ULONG user_mode;
} gETWContext;                  /* IMPORTANT : Sync access via gETWCS */


/* IMPORTANT : Sync access via gETWCS */


CRITICAL_SECTION gETWCS; /* Access to gETWContext */

/* Used for testing old MOF based APIs on newer Windows OS'es */
static int gForceMofAPI = 0;

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

static int TwapiCalcPointerSize(EVENT_RECORD *evrP)
{
    if (evrP->EventHeader.Flags & EVENT_HEADER_FLAG_32_BIT_HEADER)
        return 4;
    else if (evrP->EventHeader.Flags & EVENT_HEADER_FLAG_64_BIT_HEADER)
        return 8;
    else
        return gETWContext.pointer_size;
}

static Tcl_Obj *ObjFromTRACE_LOGFILE_HEADER(TRACE_LOGFILE_HEADER *tlhP)
{
    Tcl_Obj *objs[21];
    TRACE_LOGFILE_HEADER *adjustedP;

    objs[0] = ObjFromULONG(tlhP->BufferSize);
    objs[1] = ObjFromInt(tlhP->VersionDetail.MajorVersion);
    objs[2] = ObjFromInt(tlhP->VersionDetail.MinorVersion);
    objs[3] = ObjFromInt(tlhP->VersionDetail.SubVersion);
    objs[4] = ObjFromInt(tlhP->VersionDetail.SubMinorVersion);
    objs[5] = ObjFromULONG(tlhP->ProviderVersion);
    objs[6] = ObjFromULONG(tlhP->NumberOfProcessors);
    objs[7] = ObjFromLARGE_INTEGER(tlhP->EndTime);
    objs[8] = ObjFromULONG(tlhP->TimerResolution);
    objs[9] = ObjFromULONG(tlhP->MaximumFileSize);
    objs[10] = ObjFromULONG(tlhP->LogFileMode);
    objs[11] = ObjFromULONG(tlhP->BuffersWritten);
    objs[12] = ObjFromULONG(tlhP->PointerSize);
    objs[13] = ObjFromULONG(tlhP->EventsLost);
    objs[14] = ObjFromULONG(tlhP->CpuSpeedInMHz);

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
    
    /* LoggerName and LogFileName fields are not to be used as per doc. */

    objs[15] = ObjFromTIME_ZONE_INFORMATION(&adjustedP->TimeZone);
    
    objs[16] = ObjFromLARGE_INTEGER(adjustedP->BootTime);
    objs[17] = ObjFromLARGE_INTEGER(adjustedP->PerfFreq);
    objs[18] = ObjFromLARGE_INTEGER(adjustedP->StartTime);
    objs[19] = ObjFromULONG(adjustedP->ReservedFlags);
    objs[20] = ObjFromULONG(adjustedP->BuffersLost);

    return ObjNewList(ARRAYSIZE(objs), objs);
}

    
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
    /* IMPORTANT : 
     * Do not change order without changing switch statement below!
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
    Tcl_Obj *objs[19];

    if (etP->LogFileNameOffset)
        objs[0] = ObjFromUnicode((WCHAR *)(etP->LogFileNameOffset + (char *) etP));
    else
        objs[0] = ObjFromEmptyString();

    if (etP->LoggerNameOffset)
        objs[1] = ObjFromUnicode((WCHAR *)(etP->LoggerNameOffset + (char *) etP));
    else
        objs[1] = ObjFromEmptyString();

    objs[2] = ObjFromGUID(&etP->Wnode.Guid);
    objs[3] = ObjFromLong(etP->BufferSize);
    objs[4] = ObjFromLong(etP->MinimumBuffers);
    objs[5] = ObjFromLong(etP->MaximumBuffers);
    objs[6] = ObjFromLong(etP->MaximumFileSize);
    objs[7] = ObjFromLong(etP->LogFileMode);
    objs[8] = ObjFromLong(etP->FlushTimer);
    objs[9] = ObjFromLong(etP->EnableFlags);
    objs[10] = ObjFromLong(etP->Wnode.ClientContext);
    objs[11] = ObjFromLong(etP->AgeLimit);
    objs[12] = ObjFromLong(etP->NumberOfBuffers);
    objs[13] = ObjFromLong(etP->FreeBuffers);
    objs[14] = ObjFromLong(etP->EventsLost);
    objs[15] = ObjFromLong(etP->BuffersWritten);
    objs[16] = ObjFromLong(etP->LogBuffersLost);
    objs[17] = ObjFromLong(etP->RealTimeBuffersLost);
    objs[18] = ObjFromWideInt((Tcl_WideInt) etP->LoggerThreadId);

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
    TCL_RESULT res;
    
    if (objc != 3)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (ObjToPEVENT_TRACE_PROPERTIES(interp, objv[2], &etP) != TCL_OK)
        return TCL_ERROR;

    /* Note etP has to be freed */

    /* If no log file specified, set logfilenameoffset to 0 
       since ObjToPEVENT_TRACE_PROPERTIES does not do that as it
       is also used by ControlTrace */
    if (*(WCHAR *) (etP->LogFileNameOffset + (char *)etP) == 0)
        etP->LogFileNameOffset = 0;

    if (StartTraceW(&htrace,
                    ObjToUnicode(objv[1]),
                    etP) == ERROR_SUCCESS) {
        ObjSetResult(interp, ObjFromTRACEHANDLE(htrace));
        res = TCL_OK;
    } else
        res = Twapi_AppendSystemError(interp, GetLastError());

    TwapiFree(etP);
    return res;
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


static Tcl_Obj *ObjFromEVENT_TRACE_HEADER(EVENT_TRACE_HEADER *ethP)
{
    Tcl_Obj *objs[10];

    objs[0] = ObjFromInt(ethP->Class.Type);
    objs[1] = ObjFromInt(ethP->Class.Level);
    objs[2] = ObjFromInt(ethP->Class.Version);
    objs[3] = ObjFromULONG(ethP->ThreadId);
    objs[4] = ObjFromULONG(ethP->ProcessId);
    objs[5] = ObjFromLARGE_INTEGER(ethP->TimeStamp);
    objs[6] = ObjFromGUID(&ethP->Guid);
    /*
     * Note - for user mode sessions, KernelTime/UserTime are not valid
     * and the ProcessorTime member has to be used instead. However,
     * we do not know the type of session at this point so we leave it
     * to the app to figure out what to use
     */
    objs[7] = ObjFromULONG(ethP->KernelTime);
    objs[8] = ObjFromULONG(ethP->UserTime);
    objs[9] = ObjFromWideInt(ethP->ProcessorTime);
    return ObjNewList(ARRAYSIZE(objs), objs);
}


void WINAPI TwapiETWEventCallback(
  PEVENT_TRACE evP
)
{
    Tcl_Obj *objs[5];

    /* Called back from Win32 ProcessTrace call. Assumed that gETWContext is locked */

    if (gETWContext.status != TCL_OK)   /* If some previous error occurred, return */
        return;

    if (evP->Header.Class.Type == EVENT_TRACE_TYPE_INFO &&
        IsEqualGUID(&evP->Header.Guid, &EventTraceGuid)) {
        /* If further events do not indicate pointer size, we will use this size*/
        gETWContext.pointer_size = ((TRACE_LOGFILE_HEADER *) evP->MofData)->PointerSize;
    }


    /* TBD - create as a list - faster than as a dict */
    objs[0] = ObjFromEVENT_TRACE_HEADER(&evP->Header);
    objs[1] = ObjFromULONG(evP->InstanceId);
    objs[2] = ObjFromULONG(evP->ParentInstanceId);
    objs[3] = ObjFromGUID(&evP->ParentGuid);
    if (evP->MofData && evP->MofLength)
        objs[4] = ObjFromByteArray(evP->MofData, evP->MofLength);
    else
        objs[4] = ObjFromEmptyString();
    
    ObjAppendElement(NULL, gETWContext.eventsObj, ObjNewList(ARRAYSIZE(objs), objs));
}

/* Used in constructing a Tcl_Obj for TRACE_EVENT_INFO when a name is 
   missing */
static Tcl_Obj *TwapiTEIUnicodeObj(TRACE_EVENT_INFO *teiP, int offset, ULONG numeric_val)
{
    if (offset == 0)
        return ObjFromULONG(numeric_val); /* Name missing, return the numeric */
    else
        return ObjFromUnicodeNoTrailingSpace((WCHAR*) (offset + (char*)teiP));
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
    Tcl_Obj *objs[9];

    objs[0] = ObjFromLong(evhP->Flags);
    objs[1] = ObjFromLong(evhP->EventProperty);
    objs[2] = ObjFromLong(evhP->ThreadId);
    objs[3] = ObjFromLong(evhP->ProcessId);
    objs[4] = ObjFromLARGE_INTEGER(evhP->TimeStamp);
    // Already returned from TRACE_EVENT_INFO so leave out here 
    // ObjFromEVENT_DESCRIPTOR(&evhP->EventDescriptor);
    // ObjFromGUID(&evhP->ProviderId);

    /*
     * Note - for user mode sessions, KernelTime/UserTime are not valid
     * and the ProcessorTime member has to be used instead. However,
     * we do not know the type of session at this point so we leave it
     * to the app to figure out what to use
     */
    objs[5] = ObjFromULONG(evhP->KernelTime);
    objs[6] = ObjFromULONG(evhP->UserTime);
    objs[7] = ObjFromULONGLONG(evhP->ProcessorTime);

    objs[8] = ObjFromGUID(&evhP->ActivityId);
    
    return ObjNewList(ARRAYSIZE(objs), objs);
}


static TCL_RESULT TwapiTdhPropertyArraySize(TwapiInterpContext *ticP,
                                            EVENT_RECORD *evrP,
                                            TRACE_EVENT_INFO *teiP,
                                            USHORT prop_index, USHORT *countP)
{
    EVENT_PROPERTY_INFO *epiP;
    DWORD winerr;
    USHORT ref_index;
    union {
        USHORT ushort_val;
        ULONG ulong_val;
    } ref_value;
    ULONG ref_value_size;
    PROPERTY_DATA_DESCRIPTOR pdd;
    TDH_CONTEXT tdhctx;

    epiP = &teiP->EventPropertyInfoArray[prop_index];
    if ((epiP->Flags & PropertyParamCount) == 0) {
        *countP = epiP->count; /* Size of array is directly specified */
        return TCL_OK;
    }

    tdhctx.ParameterValue = TwapiCalcPointerSize(evrP);
    tdhctx.ParameterType = TDH_CONTEXT_POINTERSIZE;
    tdhctx.ParameterSize = 0;   /* Reserved value */

    /* Size of array is indirectly specified through some other property */
    TWAPI_ASSERT(epiP->NameOffset != 0);

    ref_index = epiP->countPropertyIndex;
    pdd.PropertyName = (ULONGLONG)(teiP->EventPropertyInfoArray[ref_index].NameOffset + (char*) teiP);
    pdd.ArrayIndex = ULONG_MAX; /* Since index property is not an array */
    pdd.Reserved = 0;
    winerr = TdhGetPropertySize(evrP, 1, &tdhctx, 1, &pdd, &ref_value_size);
    if (winerr == ERROR_SUCCESS) {
        if (ref_value_size != 2 && ref_value_size != 4)
            return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_DATA, "Indirect property index size is not 2 or 4.");
        winerr = TdhGetProperty(evrP, 1, &tdhctx, 1, &pdd, ref_value_size, (PBYTE)&ref_value);
    }
    if (winerr != ERROR_SUCCESS)
        return Twapi_AppendSystemError(ticP->interp, winerr);

    if (ref_value_size == 2)
        *countP = ref_value.ushort_val;
    else {
        if (ref_value.ulong_val > teiP->PropertyCount)
            return TwapiReturnErrorEx(ticP->interp, TWAPI_INVALID_DATA,
                                      Tcl_ObjPrintf("Property index %d out of bounds.", ref_value.ulong_val));

        *countP = (USHORT) ref_value.ulong_val;
    }

    return TCL_OK;
}


static Tcl_Obj *TwapiMapTDHProperty(EVENT_MAP_INFO *emiP, ULONG val)
{
    ULONG i, bitmask;
    Tcl_Obj *objP;

    if (emiP->Flag == EVENTMAP_INFO_FLAG_MANIFEST_PATTERNMAP ||
        emiP->MapEntryValueType != EVENTMAP_ENTRY_VALUETYPE_ULONG) {
        /* SHould not happen since this routine is called only for
         * numerics and I assume these should not occur for numerics.
         * But still, to be safe, check and punt.
         */
        return ObjFromDWORD(val);
    }

    /*
     * We can map the following combinations based on emiP->Flag
     * (exactly these bits are set and no others)
     * EVENTMAP_INFO_FLAG_MANIFEST_VALUEMAP
     *   - integer to string value using a mapping array (manifest based)
     * EVENTMAP_INFO_FLAG_WBEM_VALUEMAP
     *   - integer to string value using a mapping array (MOF based)
     * EVENTMAP_INFO_FLAG_WBEM_VALUEMAP | EVENTMAP_INFO_FLAG_WBEM_NO_MAP
     *   - integer to string value using 0-based indexing (MOF based)
     * EVENTMAP_INFO_FLAG_WBEM_VALUEMAP | EVENTMAP_INFO_FLAG_WBEM_FLAG
     *   - bit flags to list of string values using a mapping array (MOF based)
     * EVENTMAP_INFO_FLAG_WBEM_VALUEMAP | EVENTMAP_INFO_FLAG_WBEM_FLAG | EVENTMAP_INFO_FLAG_WBEM_NO_MAP
     *   - not clear this is valid and makes sense 
     * EVENTMAP_INFO_FLAG_MANIFEST_BITMAP
     *   - bit map to list of strings using a mapping array (manifest based)
     * EVENTMAP_INFO_FLAG_WBEM_BITMAP
     *   - bit map to list of strings using a mapping array (MOF based)
     * EVENTMAP_INFO_FLAG_WBEM_BITMAP | EVENTMAP_INFO_FLAG_WBEM_NO_MAP
     *   - maps 0-based bit positions to list of string values (MOF based)
     */
     
    switch (emiP->Flag) {
    case EVENTMAP_INFO_FLAG_MANIFEST_VALUEMAP:
    case EVENTMAP_INFO_FLAG_WBEM_VALUEMAP:
        for (i = 0; i < emiP->EntryCount; ++i) {
            if (emiP->MapEntryArray[i].Value == val)
                return ObjFromUnicode((WCHAR *) (emiP->MapEntryArray[i].OutputOffset + (char*)emiP));
        }
        break;
    case EVENTMAP_INFO_FLAG_WBEM_VALUEMAP | EVENTMAP_INFO_FLAG_WBEM_NO_MAP:
        if (val < emiP->EntryCount)
            return ObjFromUnicode((WCHAR *) (emiP->MapEntryArray[val].OutputOffset + (char*)emiP));
        break;
    case EVENTMAP_INFO_FLAG_WBEM_VALUEMAP | EVENTMAP_INFO_FLAG_WBEM_FLAG:
        objP = ObjNewList(0, NULL);
        /* We ignore bits in val that are set but don't have a matching entry */
        for (i = 0; i < emiP->EntryCount; ++i) {
            if ((emiP->MapEntryArray[i].Value & val) == emiP->MapEntryArray[i].Value)
                ObjAppendElement(NULL, objP, ObjFromUnicode((WCHAR *) (emiP->MapEntryArray[i].OutputOffset + (char*)emiP)));
        }
        return objP;


    case EVENTMAP_INFO_FLAG_MANIFEST_BITMAP:
        objP = ObjNewList(0, NULL);
        for (i = 0; val != 0 && i < emiP->EntryCount; ++i) {
            /* Value field is the bit mask. Match with any bit means match */
            bitmask = emiP->MapEntryArray[i].Value;
            if (bitmask & val) {
                ObjAppendElement(NULL, objP, ObjFromUnicode((WCHAR *) (emiP->MapEntryArray[i].OutputOffset + (char*)emiP)));
                val &= ~ bitmask;
            }
        }
        return objP;

    case EVENTMAP_INFO_FLAG_WBEM_BITMAP:
        objP = ObjNewList(0, NULL);
        for (i = 0; val != 0 && i < emiP->EntryCount; ++i) {
            if (emiP->MapEntryArray[i].Value > 31)
                continue;
            /* Value field is the bit position to check */
            bitmask = (1 << emiP->MapEntryArray[i].Value);
            if (bitmask & val) {
                ObjAppendElement(NULL, objP, ObjFromUnicode((WCHAR *) (emiP->MapEntryArray[i].OutputOffset + (char*)emiP)));
                /* Reset the bit both for efficiency as well as so we only
                   include one element even in the MapEntryArray[] contains
                   duplicate values */
                val &= ~bitmask;
            }
        }
        return objP;

    case EVENTMAP_INFO_FLAG_WBEM_BITMAP | EVENTMAP_INFO_FLAG_WBEM_NO_MAP:
        objP = ObjNewList(0, NULL);
        for (i = 0; val != 0 && i < emiP->EntryCount && i < 32; ++i) {
            bitmask = (1 << i);
            if (val & bitmask) {
                ObjAppendElement(NULL, objP, ObjFromUnicode((WCHAR *) (emiP->MapEntryArray[i].OutputOffset + (char*)emiP)));
                val &= ~bitmask;
            }
        }
        return objP;

    case EVENTMAP_INFO_FLAG_WBEM_VALUEMAP | EVENTMAP_INFO_FLAG_WBEM_FLAG | EVENTMAP_INFO_FLAG_WBEM_NO_MAP:
        break;                  /* TBD */
    }

    /* No map value, return as integer */
    return ObjFromDWORD(val);
}


/* Given the raw bytes for a TDH property, return the Tcl_Obj */
static TCL_RESULT TwapiTdhPropertyValue(
    TwapiInterpContext *ticP,
    EVENT_RECORD *evrP,
    EVENT_PROPERTY_INFO *epiP,
    void *bytesP,               /* Raw bytes */
    ULONG prop_size,            /* Number of bytes */
    EVENT_MAP_INFO *emiP,       /* Value map, may be NULL */
    Tcl_Obj **valueObjP
    )
{
    union {
        __int64 i64;
        unsigned __int64 ui64;
        double dbl;
        struct {
            void *s;
            ULONG   len;
        } string;
        GUID guid;
        FILETIME ftime;
        SYSTEMTIME stime;
        char *bin;
    } u;
    ULONG remain = prop_size;
    DWORD dw;
    Tcl_Interp *interp = ticP->interp;


#define EXTRACT(var_, type_) \
    do { \
        if (prop_size < sizeof(type_)) goto size_error;        \
        var_ = *(type_ UNALIGNED *) bytesP; \
    } while (0)
    switch (epiP->nonStructType.InType) {
    case TDH_INTYPE_NULL:
        *valueObjP = ObjFromEmptyString(); /* TBD - use a NULL obj type*/
        return TCL_OK;
    case TDH_INTYPE_UNICODESTRING:
        u.string.s = bytesP;
        u.string.len = -1;
        break;
    case TDH_INTYPE_ANSISTRING:
        u.string.s = bytesP;
        u.string.len = -1;
        break;

    case TDH_INTYPE_INT8:  EXTRACT(u.i64, char); break;
    case TDH_INTYPE_UINT8: EXTRACT(u.i64, unsigned char); break;
    case TDH_INTYPE_INT16: EXTRACT(u.i64, short); break;
    case TDH_INTYPE_UINT16: EXTRACT(u.i64, unsigned short); break;
    case TDH_INTYPE_INT32:  EXTRACT(u.i64, int); break;
    case TDH_INTYPE_UINT32: EXTRACT(u.i64, unsigned int); break;
    case TDH_INTYPE_INT64:  EXTRACT(u.i64, __int64); break;
    case TDH_INTYPE_UINT64: EXTRACT(u.i64, unsigned __int64); break;
    case TDH_INTYPE_FLOAT:  EXTRACT(u.dbl, float); break;
    case TDH_INTYPE_DOUBLE: EXTRACT(u.dbl, double); break;

    case TDH_INTYPE_BOOLEAN:
        EXTRACT(u.i64, int);
        *valueObjP = ObjFromBoolean(u.i64 != 0);
        return TCL_OK;

    case TDH_INTYPE_BINARY:
        u.bin = bytesP;
        break;
    case TDH_INTYPE_GUID:
        if (prop_size < sizeof(GUID))
            goto size_error;
        *valueObjP = ObjFromGUID(bytesP);
        return TCL_OK;
    case TDH_INTYPE_POINTER:
    case TDH_INTYPE_HEXINT32:
    case TDH_INTYPE_HEXINT64:
        if (prop_size == 4)
            *valueObjP = ObjFromULONGHex(*(unsigned int UNALIGNED *)bytesP);
        else if (prop_size == 8)
            *valueObjP = ObjFromULONGLONGHex(*(unsigned __int64 UNALIGNED *)bytesP);
        else
            goto size_error;
        return TCL_OK;
    case TDH_INTYPE_FILETIME:
        EXTRACT(u.ftime, FILETIME);
        *valueObjP = ObjFromFILETIME(&u.ftime);
        return TCL_OK;
    case TDH_INTYPE_SYSTEMTIME:
        EXTRACT(u.stime, SYSTEMTIME);
        *valueObjP = ObjFromSYSTEMTIME(&u.stime);
        return TCL_OK;

    case TDH_INTYPE_COUNTEDSTRING:
    case TDH_INTYPE_REVERSEDCOUNTEDSTRING:
    case TDH_INTYPE_COUNTEDANSISTRING:
    case TDH_INTYPE_REVERSEDCOUNTEDANSISTRING:
        EXTRACT(u.string.len, unsigned short);
        if (epiP->nonStructType.InType == TDH_INTYPE_REVERSEDCOUNTEDSTRING ||
            epiP->nonStructType.InType == TDH_INTYPE_REVERSEDCOUNTEDANSISTRING)
            u.string.len = ((0xff & u.string.len) << 8) | (u.string.len >> 8);
        remain -= sizeof(unsigned short);
        u.string.s = (WCHAR *)(sizeof(unsigned short) + (char *)bytesP);
        break;
    case TDH_INTYPE_NONNULLTERMINATEDSTRING:
        u.string.s = bytesP;
        u.string.len = prop_size/sizeof(WCHAR);
        break;
    case TDH_INTYPE_UNICODECHAR:
    case TDH_INTYPE_ANSICHAR:
        u.string.s = bytesP;
        u.string.len = 1;
        break;
    case TDH_INTYPE_NONNULLTERMINATEDANSISTRING:
        u.string.s = bytesP;
        u.string.len = prop_size;
        break;

    case TDH_INTYPE_SIZET:
        if (prop_size == 4)
            *valueObjP = ObjFromULONG(*(unsigned int UNALIGNED *)bytesP);
        else if (prop_size == 8)
            *valueObjP = ObjFromULONGLONG(*(unsigned __int64 UNALIGNED *)bytesP);
        else
            goto size_error;
        return TCL_OK;

    case TDH_INTYPE_HEXDUMP:
        EXTRACT(dw, DWORD);
        remain -= sizeof(DWORD);
        *valueObjP = ObjFromByteArray(sizeof(DWORD)+(char*)bytesP,
                                      remain < dw ? remain : dw);
        return TCL_OK;

    case TDH_INTYPE_WBEMSID:
        /* TOKEN_USER structure followed by SID. Sizeof TOKEN_USER
           depends on 32/64 bittedness of event stream. */
        if (evrP->EventHeader.Flags & EVENT_HEADER_FLAG_32_BIT_HEADER)
            dw = 4;
        else if (evrP->EventHeader.Flags & EVENT_HEADER_FLAG_64_BIT_HEADER)
            dw = 8;
        else
            dw = gETWContext.pointer_size;
        dw *= 2; /* sizeof(TOKEN_USER) == 16 on 64bit arch, 8 on 32bit */
        if (prop_size < (dw+sizeof(SID)))
            goto size_error;
        bytesP = dw + (char*) bytesP;
        prop_size -= dw;
        /* FALLTHROUGH */

    case TDH_INTYPE_SID:
        if (TwapiValidateSID(interp, bytesP, prop_size) != TCL_OK)
            return TCL_ERROR;
        return ObjFromSID(interp, bytesP, valueObjP);


    default:
        return TwapiReturnErrorEx(interp, TWAPI_UNSUPPORTED_TYPE,
                                  Tcl_ObjPrintf("Unsupported TDH type %d.", epiP->nonStructType.InType));
    }

    
    /* Now format based on output type */
    switch (epiP->nonStructType.InType) {
    case TDH_INTYPE_UNICODESTRING:
    case TDH_INTYPE_COUNTEDSTRING:
    case TDH_INTYPE_REVERSEDCOUNTEDSTRING:
    case TDH_INTYPE_NONNULLTERMINATEDSTRING:
    case TDH_INTYPE_UNICODECHAR:
        if (u.string.len == -1) {
            *valueObjP = ObjFromUnicodeLimited((WCHAR*)u.string.s,
                                               remain/sizeof(WCHAR), NULL);
        } else {
            remain /= sizeof(WCHAR);
            if (remain < u.string.len)
                u.string.len = remain;
            *valueObjP = ObjFromUnicodeN((WCHAR*)u.string.s, u.string.len);
        }
        break;

    case TDH_INTYPE_ANSISTRING:
    case TDH_INTYPE_COUNTEDANSISTRING:
    case TDH_INTYPE_REVERSEDCOUNTEDANSISTRING:
    case TDH_INTYPE_ANSICHAR:
    case TDH_INTYPE_NONNULLTERMINATEDANSISTRING:
        /* TBD do we need to check that string really is ASCII ? */
        if (u.string.len == -1)
            *valueObjP = ObjFromStringLimited((char*)u.string.s, remain, NULL);
        else
            *valueObjP = ObjFromStringN((char*)u.string.s,
                                        remain < u.string.len ? remain : u.string.len);
        break;

    case TDH_INTYPE_INT8:
    case TDH_INTYPE_UINT8:
    case TDH_INTYPE_INT16:
    case TDH_INTYPE_UINT16:
    case TDH_INTYPE_INT32:
    case TDH_INTYPE_UINT32:
    case TDH_INTYPE_INT64:
    case TDH_INTYPE_UINT64:
        switch (epiP->nonStructType.OutType) {
        case TDH_OUTTYPE_HEXINT8:
            *valueObjP = ObjFromUCHARHex((unsigned char)u.i64);
            break;
        case TDH_OUTTYPE_HEXINT16:
            *valueObjP = ObjFromUSHORTHex((USHORT)u.i64);
            break;
        case TDH_OUTTYPE_HRESULT:
        case TDH_OUTTYPE_WIN32ERROR:
        case TDH_OUTTYPE_NTSTATUS:
        case TDH_OUTTYPE_HEXINT32:
            *valueObjP = ObjFromULONGHex((ULONG)u.i64);
            break;
        case TDH_OUTTYPE_HEXINT64:
            *valueObjP = ObjFromULONGLONGHex(u.i64);
            break;
        case TDH_OUTTYPE_IPV4:
            *valueObjP = IPAddrObjFromDWORD((DWORD) u.i64);
            break;
        case TDH_OUTTYPE_PORT:
            *valueObjP = ObjFromLong((USHORT)ntohs((USHORT)u.i64));
            break;
        default:
            if (emiP && u.i64 < ULONG_MAX && u.i64 >= 0) {
                *valueObjP = TwapiMapTDHProperty(emiP, (ULONG) u.i64); 
            } else {
                if (epiP->nonStructType.InType == TDH_INTYPE_UINT64)
                    *valueObjP = ObjFromULONGLONG(u.ui64);
                else
                    *valueObjP = ObjFromWideInt(u.i64);
            }
            break;
        }
        break;

    case TDH_INTYPE_FLOAT:
    case TDH_INTYPE_DOUBLE:
        *valueObjP = ObjFromDouble(u.dbl);
        break;

    case TDH_INTYPE_BINARY:
        if (epiP->nonStructType.OutType == TDH_OUTTYPE_IPV6) {
            if (remain != 16)
                goto size_error;
            *valueObjP = ObjFromIPv6Addr(u.bin, 0);
            if (*valueObjP == NULL)
                return TwapiReturnSystemError(interp);
        } else {
            *valueObjP = ObjFromByteArray(u.bin, remain);
        }
        break;

    default:
        return TwapiReturnErrorEx(interp, TWAPI_UNSUPPORTED_TYPE,
                                  Tcl_ObjPrintf("Unsupported TDH type %d.", epiP->nonStructType.InType));
    }

    return TCL_OK;

size_error:
    return TwapiReturnErrorEx(interp, TWAPI_INVALID_DATA,
                              Tcl_ObjPrintf("TDH property field of type %d trucated or wrong size", epiP->nonStructType.InType));
}


/* Uses ticP->memlifo, caller responsible for memory management always */
static TCL_RESULT TwapiDecodeEVENT_PROPERTY_INFO(
    TwapiInterpContext *ticP,
    EVENT_RECORD *evrP,
    TRACE_EVENT_INFO *teiP,
    USHORT prop_index,
    LPWSTR struct_name,         /* If non-NULL, property is actually member
                                   of a struct of this name */
    USHORT struct_index,        /* Index of owning struct property when
                                   prop_index is a struct member
                                 */
    Tcl_Obj **propnameObjP,
    Tcl_Obj **propvalObjP
    )
{
    EVENT_PROPERTY_INFO *epiP;
    Tcl_Obj **valueObjs;
    USHORT nvalues, array_index;
    TCL_RESULT res;
    DWORD winerr;
    Tcl_Interp *interp = ticP->interp;
    MemLifo *memlifoP = ticP->memlifoP;
    void *pv;
    TDH_CONTEXT tdhctx;
    PROPERTY_DATA_DESCRIPTOR pdd[2];
    int pdd_count;
    ULONG prop_size;
    ULONGLONG prop_name;

    epiP = &teiP->EventPropertyInfoArray[prop_index];

    if (epiP->NameOffset == 0) {
        /* Should not happen. */
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "NameOffset field is 0 for property in event record.");
    }

    tdhctx.ParameterValue = TwapiCalcPointerSize(evrP);
    tdhctx.ParameterType = TDH_CONTEXT_POINTERSIZE;
    tdhctx.ParameterSize = 0;   /* Reserved value */

    res = TwapiTdhPropertyArraySize(ticP, evrP, teiP, prop_index, &nvalues);
    if (res != TCL_OK)
        return res;
    
    prop_name = (ULONGLONG)(epiP->NameOffset + (char*) teiP);

    /* Special case arrays of UNICHAR and ANSICHAR. These are actually strings*/
    if ((epiP->Flags & PropertyStruct) == 0 &&
        epiP->nonStructType.OutType == TDH_OUTTYPE_STRING &&
        (epiP->nonStructType.InType == TDH_INTYPE_UNICODECHAR ||
         epiP->nonStructType.InType == TDH_INTYPE_ANSICHAR)) {
        pdd[0].PropertyName = prop_name;
        pdd[0].ArrayIndex = ULONG_MAX; /* We want size of whole array */
        pdd[0].Reserved = 0;
        winerr = TdhGetPropertySize(evrP, 1, &tdhctx, 1, pdd, &prop_size);
        if (winerr == ERROR_SUCCESS) {
            /* Do we need to check for presence of map info here ? */
            pv = MemLifoPushFrame(memlifoP, prop_size, NULL);
            winerr = TdhGetProperty(evrP, 1, &tdhctx, 1, pdd, prop_size, pv);
            if (winerr == ERROR_SUCCESS) {
                *propvalObjP  = ObjNewList(1, NULL);
                *propnameObjP = ObjFromUnicode((WCHAR *)(epiP->NameOffset + (char*)teiP));
                if (epiP->nonStructType.InType == TDH_INTYPE_UNICODECHAR)
                    ObjAppendElement(NULL, *propvalObjP,
                                     ObjFromUnicodeLimited(pv, prop_size/sizeof(WCHAR), NULL));
                else
                    ObjAppendElement(NULL, *propvalObjP,
                                     ObjFromStringLimited(pv, prop_size, NULL));
            }
            MemLifoPopFrame(memlifoP);
        }
        return winerr == ERROR_SUCCESS ? TCL_OK : Twapi_AppendSystemError(interp, winerr);
    }

    valueObjs = nvalues ?
        MemLifoAlloc(memlifoP, nvalues * sizeof(*valueObjs), NULL)
        : NULL;

    for (array_index = 0; array_index < nvalues; ++array_index) {
        Tcl_Obj *valueObj = NULL;
        if (epiP->Flags & PropertyStruct) {
            /* Property is a struct */
            USHORT member_index     = epiP->structType.StructStartIndex;
            ULONG member_index_bound = member_index + epiP->structType.NumOfStructMembers;
            valueObj = ObjNewList(2 * epiP->structType.NumOfStructMembers, NULL);
            if (member_index_bound > teiP->TopLevelPropertyCount) {
                res = TwapiReturnErrorEx(interp,
                                         TWAPI_INVALID_DATA,
                                         Tcl_ObjPrintf("Property index %d out of bounds.", member_index_bound));
            } else {
                while (member_index < member_index_bound) {
                    Tcl_Obj *membernameObj, *membervalObj;
                    res = TwapiDecodeEVENT_PROPERTY_INFO(ticP, evrP, teiP, member_index, (LPWSTR)(epiP->NameOffset + (char *) teiP), array_index, &membernameObj, &membervalObj);
                    if (res != TCL_OK)
                        break;
                    ObjAppendElement(NULL, valueObj, membernameObj);
                    ObjAppendElement(NULL, valueObj, membervalObj);
                    ++member_index;
                }
            }
        } else {
            /* Property is a scalar */
            if (struct_name) {
                pdd_count = 2;
                pdd[0].PropertyName = (ULONGLONG)struct_name;
                pdd[0].ArrayIndex = struct_index;
                pdd[0].Reserved = 0;
                pdd[1].PropertyName = prop_name;
                pdd[1].ArrayIndex = array_index; /* TBD - not clear what this should be */
                pdd[1].Reserved = 0;

            } else {
                /* Top level scalar, not part of a struct */
                pdd_count = 1;
                pdd[0].PropertyName = prop_name;
                pdd[0].ArrayIndex = array_index;
                pdd[0].Reserved = 0;

                /* TBD - sample in MSDN docs (not sdk sample) says tdh
                   cannot handle IPv6 data and skips event. Check on this */
            }            
            /* TBD - see GetPropertyLength in SDK doc article
               "Using TdhFormatProperty to Consume Event Data". More involved
               than just calling TdhGetPropertySize. Do we need to copy
               that code ? */
            winerr = TdhGetPropertySize(evrP, 1, &tdhctx, pdd_count, pdd, &prop_size);
            if (winerr == ERROR_SUCCESS) {
                ULONG map_size;
                EVENT_MAP_INFO *mapP;

                /* Since we might be looping, alloc and release memory in 
                   every iteration.
                   Not necessary for correctness since caller will pop memlifo
                   frame anyway so in error case, we don't bother to pop
                   the frame */
                pv = MemLifoPushFrame(memlifoP, prop_size, NULL);
                if (epiP->nonStructType.MapNameOffset == 0)
                    mapP = NULL;
                else {
                    map_size = 0;
                    winerr = TdhGetEventMapInformation(evrP, (LPWSTR) (epiP->nonStructType.MapNameOffset + (char *)teiP), NULL, &map_size);
                    if (winerr == ERROR_INSUFFICIENT_BUFFER) {
                        mapP = MemLifoAlloc(memlifoP, map_size, NULL);
                        winerr = TdhGetEventMapInformation(evrP, (LPWSTR) (epiP->nonStructType.MapNameOffset + (char *)teiP), mapP, &map_size);
                        
                    }
                }
                if (winerr == ERROR_SUCCESS) {
                    winerr = TdhGetProperty(evrP, 1, &tdhctx, pdd_count, pdd, prop_size, pv);
                    if (winerr == ERROR_SUCCESS)
                        res = TwapiTdhPropertyValue(ticP, evrP, epiP, pv, prop_size, mapP, &valueObj);
                }
                MemLifoPopFrame(memlifoP);
            }
            if (winerr != ERROR_SUCCESS)
                res = Twapi_AppendSystemError(interp, winerr);                
        }

        if (res != TCL_OK) {
            if (valueObj)
                ObjDecrRefs(valueObj);
            ObjDecrArrayRefs(array_index, valueObjs);
            return res;
        }
        valueObjs[array_index] = valueObj;
    }

    *propvalObjP  = ObjNewList(nvalues, valueObjs);
    *propnameObjP = ObjFromUnicode((WCHAR *)(epiP->NameOffset + (char*)teiP));
    return TCL_OK;
}


/* Uses memlifo frame. Caller responsible for cleanup */
static TCL_RESULT TwapiTdhGetEventInformation(TwapiInterpContext *ticP, EVENT_RECORD *evrP, Tcl_Obj **teiObjP)
{
    DWORD sz, winerr;
    Tcl_Obj *objs[15];
    TCL_RESULT status;
    TRACE_EVENT_INFO *teiP;
    EVENT_DESCRIPTOR *edP;
    TDH_CONTEXT tdhctx;
    int classic;
    Tcl_Obj *emptyObj;

    /* TBD - instrument how much to try for initially */
    teiP = MemLifoAlloc(ticP->memlifoP, 1000, &sz);

    tdhctx.ParameterValue = TwapiCalcPointerSize(evrP);
    tdhctx.ParameterType = TDH_CONTEXT_POINTERSIZE;
    tdhctx.ParameterSize = 0;   /* Reserved value */

    winerr = TdhGetEventInformation(evrP, 1, &tdhctx, teiP, &sz);
    if (winerr == ERROR_INSUFFICIENT_BUFFER) {
        teiP = MemLifoAlloc(ticP->memlifoP, sz, NULL);
        winerr = TdhGetEventInformation(evrP, 1, &tdhctx, teiP, &sz);
    }
    
    if (winerr != ERROR_SUCCESS)
        return Twapi_AppendSystemError(ticP->interp, winerr);

    edP = &teiP->EventDescriptor;

    switch (teiP->DecodingSource) {
    case DecodingSourceXMLFile: classic = 0; break;
    case DecodingSourceWbem:    classic = 1; break;
    default:
        return TwapiReturnErrorEx(ticP->interp, TWAPI_UNSUPPORTED_TYPE,
                                  Tcl_ObjPrintf("Unsupported ETW decoding source (%d)", teiP->DecodingSource));
    }


    emptyObj = ObjFromEmptyString();
    ObjIncrRefs(emptyObj);  /* Since we DecrRefs it for error handling */

#define OFFSET_TO_OBJ(field_) (teiP->field_ ? ObjFromUnicodeNoTrailingSpace((LPWSTR)(teiP->field_ + (char*)teiP)) : emptyObj)

    objs[0] = ObjFromGUID(&teiP->ProviderGuid);
    objs[1] = classic ? ObjFromGUID(&teiP->EventGuid) : emptyObj;
    objs[2] = ObjFromEVENT_DESCRIPTOR(&teiP->EventDescriptor);
    objs[3] = ObjFromLong(teiP->DecodingSource);
    objs[4] = OFFSET_TO_OBJ(ProviderNameOffset);
    objs[5] = TwapiTEIUnicodeObj(teiP, teiP->LevelNameOffset, edP->Level);
    objs[6] = TwapiTEIUnicodeObj(teiP, teiP->ChannelNameOffset, edP->Channel);
    if (teiP->KeywordsNameOffset)
        objs[7] = ObjFromMultiSz((LPWSTR) (teiP->KeywordsNameOffset + (char*)teiP), -1);
    else
        objs[7] = emptyObj;
    objs[8] = TwapiTEIUnicodeObj(teiP, teiP->TaskNameOffset, edP->Task);
    objs[9] = TwapiTEIUnicodeObj(teiP, teiP->OpcodeNameOffset, edP->Opcode);
    objs[10] = OFFSET_TO_OBJ(EventMessageOffset);
    objs[11] = OFFSET_TO_OBJ(ProviderMessageOffset);
    if (classic) {
        objs[12] = OFFSET_TO_OBJ(ActivityIDNameOffset);
        objs[13] = OFFSET_TO_OBJ(RelatedActivityIDNameOffset);
    } else {
        objs[12] = emptyObj;
        objs[13] = emptyObj;
    }

    objs[14] = ObjNewList(2 * teiP->TopLevelPropertyCount, NULL);
    if (evrP->EventHeader.Flags & EVENT_HEADER_FLAG_STRING_ONLY) {
        ObjAppendElement(NULL, objs[14], STRING_LITERAL_OBJ("stringdata"));
        ObjAppendElement(NULL, objs[14],
                         ObjFromUnicodeLimited(evrP->UserData,
                                               evrP->UserDataLength/sizeof(WCHAR), NULL));
    } else {
        USHORT i;

        for (i = 0; i < teiP->TopLevelPropertyCount; ++i) {
            Tcl_Obj *propnameObj, *propvalObj;
            status = TwapiDecodeEVENT_PROPERTY_INFO(ticP, evrP, teiP, i, NULL, 0, &propnameObj, &propvalObj);
            if (status != TCL_OK) {
                /* Cannot use ObjDecrArrayRefs here to free objs[] because
                   it contains multiple occurences of emptyObj. Explicitly
                   build a list and free it */
                ObjDecrRefs(ObjNewList(ARRAYSIZE(objs), objs));
                if (emptyObj)
                    ObjDecrRefs(emptyObj); /* For the additional IncrRefs above */
                return TCL_ERROR;
            }
            ObjAppendElement(NULL, objs[14], propnameObj);
            ObjAppendElement(NULL, objs[14], propvalObj);
        }
    }

    *teiObjP = ObjNewList(ARRAYSIZE(objs), objs);
    ObjDecrRefs(emptyObj);

    return TCL_OK;
}

static VOID WINAPI TwapiETWEventRecordCallback(PEVENT_RECORD evrP)
{
    int i;
    Tcl_Obj *recObjs[4];
    Tcl_Obj *objs[3];
    MemLifoMarkHandle mark;
    TwapiInterpContext *ticP;

    /* Called back from Win32 ProcessTrace call. Assumed that gETWContext is locked */
    TWAPI_ASSERT(gETWContext.ticP != NULL);
    TWAPI_ASSERT(gETWContext.ticP->interp != NULL);

    if (gETWContext.status != TCL_OK) /* If some previous error occurred, return */
        return;

    if ((evrP->EventHeader.Flags & EVENT_HEADER_FLAG_TRACE_MESSAGE) != 0)
        return; // Ignore WPP events. - TBD
    
    if (evrP->EventHeader.EventDescriptor.Opcode == EVENT_TRACE_TYPE_INFO &&
        IsEqualGUID(&evrP->EventHeader.ProviderId, &EventTraceGuid)) {
        /*
         * This event is generated per log file. If an individual event do not
         * indicate pointer size, we will use this size.
         */
        gETWContext.pointer_size = ((TRACE_LOGFILE_HEADER *) evrP->UserData)->PointerSize;
    }

    ticP = gETWContext.ticP;
    mark = MemLifoPushMark(ticP->memlifoP);

    recObjs[0] = ObjFromEVENT_HEADER(&evrP->EventHeader);

    /* Buffer Context */
    if (evrP->EventHeader.Flags & EVENT_HEADER_FLAG_PROCESSOR_INDEX) {
        /* Win 8 defines as a USHORT (field ProcessorIndex in new sdk) */
        objs[0] = ObjFromLong(*(USHORT *)&evrP->BufferContext.ProcessorNumber);
    } else {
        objs[0] = ObjFromLong(evrP->BufferContext.ProcessorNumber);
    }
    objs[1] = ObjFromLong(evrP->BufferContext.LoggerId);
    recObjs[1] = ObjNewList(2, objs);
    
    /* Extended Data */
    /* TBD - do we need to check the evrP->EventHeader.Flags EVENT_HEADER_FLAG_EXTENDED_INFO bit ? */
    recObjs[2] = ObjNewList(0, NULL);
    for (i = 0; i < evrP->ExtendedDataCount; ++i) {
        EVENT_HEADER_EXTENDED_DATA_ITEM *ehdrP = &evrP->ExtendedData[i];
        if (ehdrP->DataPtr == 0)
            continue;
        switch (ehdrP->ExtType) {
        case EVENT_HEADER_EXT_TYPE_RELATED_ACTIVITYID:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("relatedactivity"));
            ObjAppendElement(NULL, recObjs[2], ObjFromGUID(&((PEVENT_EXTENDED_ITEM_RELATED_ACTIVITYID)ehdrP->DataPtr)->RelatedActivityId));
            break;
        case EVENT_HEADER_EXT_TYPE_SID:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("sid"));
            ObjAppendElement(NULL, recObjs[2], ObjFromSIDNoFail((SID *)(ehdrP->DataPtr)));
            break;
        case EVENT_HEADER_EXT_TYPE_TS_ID:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("tssession"));
            ObjAppendElement(NULL, recObjs[2], ObjFromULONG(((PEVENT_EXTENDED_ITEM_TS_ID)ehdrP->DataPtr)->SessionId));
            break;
        case EVENT_HEADER_EXT_TYPE_INSTANCE_INFO:
            ObjAppendElement(NULL, recObjs[2], STRING_LITERAL_OBJ("iteminstance"));
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
    
    gETWContext.status = TwapiTdhGetEventInformation(ticP, evrP, &recObjs[3]);
    if (gETWContext.status == TCL_OK)
        ObjAppendElement(ticP->interp, gETWContext.eventsObj, ObjNewList(ARRAYSIZE(recObjs), recObjs));
    else
        ObjDecrArrayRefs(3, objs);

    MemLifoPopMark(mark);
    return;
}


ULONG WINAPI TwapiETWBufferCallback(
  PEVENT_TRACE_LOGFILEW etlP
)
{
    Tcl_Obj *evalObj;
    Tcl_Obj *bufObj;
    Tcl_Obj *args[2];
    Tcl_Interp *interp;
    int code;
    Tcl_Obj *objs[8];

    /* Called back from Win32 ProcessTrace call. Assumed that gETWContext is locked */
    TWAPI_ASSERT(gETWContext.ticP != NULL);
    TWAPI_ASSERT(gETWContext.ticP->interp != NULL);
    interp = gETWContext.ticP->interp;

    if (Tcl_InterpDeleted(interp))
        return FALSE;

    if (gETWContext.status != TCL_OK) /* If some previous error occurred, return */
        return FALSE;

    TWAPI_ASSERT(gETWContext.eventsObj);

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
            evalObj = ObjDuplicate(gETWContext.buffer.cmdObj);
            ObjIncrRefs(evalObj);
        } else
            evalObj = gETWContext.buffer.cmdObj;
    }

    objs[0] = etlP->LogFileName ? ObjFromUnicode(etlP->LogFileName) : ObjFromEmptyString();
    objs[1] = etlP->LoggerName ? ObjFromUnicode(etlP->LoggerName) : ObjFromEmptyString();
    objs[2] = ObjFromULONGLONG(etlP->CurrentTime);
    objs[3] = ObjFromLong(etlP->BuffersRead);
    //  ObjFromLong(etlP->LogFileMode) - docs say do not use
    objs[4] = ObjFromTRACE_LOGFILE_HEADER(&etlP->LogfileHeader);
    objs[5] = ObjFromLong(etlP->BufferSize);
    objs[6] = ObjFromLong(etlP->Filled);
    // Docs say unused -  ObjFromLong(etlP->EventsLost));
    objs[7] = ObjFromLong(etlP->IsKernelTrace);

    bufObj = ObjNewList(ARRAYSIZE(objs), objs);

    if (gETWContext.buffer_cmdlen) {
        args[0] = bufObj;
        args[1] = gETWContext.eventsObj;
        Tcl_ListObjReplace(interp, evalObj, gETWContext.buffer_cmdlen, ARRAYSIZE(args), ARRAYSIZE(args), args);
        code = Tcl_EvalObjEx(gETWContext.ticP->interp, evalObj, TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);

        /* Get rid of the command obj if we created it */
        if (evalObj != gETWContext.buffer.cmdObj)
            ObjDecrRefs(evalObj);
    } else {
        /* No callback. Just collect */
        ObjAppendElement(NULL, gETWContext.buffer.listObj, bufObj);
        ObjAppendElement(NULL, gETWContext.buffer.listObj, gETWContext.eventsObj);
        code = TCL_OK;
    }

    /* Note bufObj is ref'ed only in one of the lists above. Do not Decr it */
    /* eventObjs needs a DecrRefs to match the one when it was created */
    ObjDecrRefs(gETWContext.eventsObj);
    gETWContext.eventsObj = ObjNewList(0, NULL);/* For next set of events */
    ObjIncrRefs(gETWContext.eventsObj);

    switch (code) {
    case TCL_BREAK:
        /* Any other value - not an error, but stop processing */
        return FALSE;
    case TCL_ERROR:
        gETWContext.status = TCL_ERROR;
        ObjDecrRefs(gETWContext.eventsObj);
        gETWContext.eventsObj = NULL;
        return FALSE;
    case TCL_OK:
    default:        /* Any other value - proceed as normal - TBD */
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
     * The gForceMofAPI setting is primarily to test old APIs on newer OS'es
     */
    if (gTdhStatus > 0 && ! gForceMofAPI) {
        etl.LogFileMode = PROCESS_TRACE_MODE_EVENT_RECORD; /* Actually etl.ProcessTraceMode */
        etl.EventCallback = (PEVENT_CALLBACK) TwapiETWEventRecordCallback;   /* Actually etl.EventRecordCallback */
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
    gETWContext.status = TCL_OK;
    gETWContext.ticP = ticP;
    gETWContext.pointer_size = sizeof(void*); /* Default unless otherwise indicated */


    winerr = ProcessTrace(htraces, ntraces, startP, endP);

    /* Copy and reset context before unlocking */
    etwc = gETWContext;
    gETWContext.buffer.cmdObj = NULL;
    gETWContext.eventsObj = NULL;
    gETWContext.status = TCL_OK;
    gETWContext.ticP = NULL;

    LeaveCriticalSection(&gETWCS);

    if (etwc.eventsObj)
        ObjDecrRefs(etwc.eventsObj);

    if (etwc.status != TCL_OK) {
        if (etwc.buffer_cmdlen == 0)
            ObjDecrRefs(etwc.buffer.listObj);
        /* interp should already have the error */
        return etwc.status;
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
    case 4:
        CHECK_NARGS_RANGE(interp, objc, 1, 2);
        if (objc == 2) {
            CHECK_INTEGER_OBJ(interp, gForceMofAPI, objv[1]);
        }
        objP = ObjFromLong(gForceMofAPI);
        break;
    default:
        return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);
    }

    if (objP)
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
        DEFINE_FNCODE_CMD(etw_force_mof, 4),
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
    INIT_TDH_STUB(TdhGetProperty);
    INIT_TDH_STUB(TdhGetPropertySize);
    INIT_TDH_STUB(TdhGetEventMapInformation);

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

