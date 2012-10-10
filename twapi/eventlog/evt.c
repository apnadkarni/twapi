/* 
 * Copyright (c) 2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/*
 * Vista+ eventlog support (Evt* functions). These are all loaded dynamically
 * as they are not available on XP/2k3
 */

#include "twapi.h"
#include "twapi_eventlog.h"

#include <ntverp.h>             /* Needed for VER_PRODUCTBUILD SDK version */

#if VER_PRODUCTBUILD < 7600
# define RUNTIME_EVT_LOAD 1
#else
# include <winevt.h>
#pragma comment(lib, "delayimp.lib") /* Prevents winevt from loading unless necessary */
#pragma comment(lib, "wevtapi.lib")	 /* New Windows Events logging library for Vista and beyond */
#endif

#ifdef RUNTIME_EVT_LOAD

/* When building with pre-Vista compilers and SDKs we need to dynamically
   load winevt. 
*/

typedef DWORD (WINAPI *EVT_SUBSCRIBE_CALLBACK)(
    int Action,
    PVOID UserContext,
    HANDLE Event );

typedef enum _EVT_VARIANT_TYPE
{
    EvtVarTypeNull        = 0,
    EvtVarTypeString      = 1,
    EvtVarTypeAnsiString  = 2,
    EvtVarTypeSByte       = 3,
    EvtVarTypeByte        = 4,
    EvtVarTypeInt16       = 5,
    EvtVarTypeUInt16      = 6,
    EvtVarTypeInt32       = 7,
    EvtVarTypeUInt32      = 8,
    EvtVarTypeInt64       = 9,
    EvtVarTypeUInt64      = 10,
    EvtVarTypeSingle      = 11,
    EvtVarTypeDouble      = 12,
    EvtVarTypeBoolean     = 13,
    EvtVarTypeBinary      = 14,
    EvtVarTypeGuid        = 15,
    EvtVarTypeSizeT       = 16,
    EvtVarTypeFileTime    = 17,
    EvtVarTypeSysTime     = 18,
    EvtVarTypeSid         = 19,
    EvtVarTypeHexInt32    = 20,
    EvtVarTypeHexInt64    = 21,

    // these types used internally
    EvtVarTypeEvtHandle   = 32,
    EvtVarTypeEvtXml      = 35

} EVT_VARIANT_TYPE;


typedef struct _EVT_RPC_LOGIN {
  LPWSTR Server;
  LPWSTR User;
  LPWSTR Domain;
  LPWSTR Password;
  DWORD  Flags;
}EVT_RPC_LOGIN;

#ifndef EVT_VARIANT_TYPE_MASK
#define EVT_VARIANT_TYPE_MASK 0x7f
#define EVT_VARIANT_TYPE_ARRAY 128
#endif

typedef struct _EVT_VARIANT
{
    union
    {
        BOOL        BooleanVal;
        INT8        SByteVal;
        INT16       Int16Val;
        INT32       Int32Val;
        INT64       Int64Val;
        UINT8       ByteVal;
        UINT16      UInt16Val;
        UINT32      UInt32Val;
        UINT64      UInt64Val;
        float       SingleVal;
        double      DoubleVal;
        ULONGLONG   FileTimeVal;
        SYSTEMTIME* SysTimeVal;
        GUID*       GuidVal;
        LPCWSTR     StringVal;
        LPCSTR      AnsiStringVal;
        PBYTE       BinaryVal;
        PSID        SidVal;
        size_t      SizeTVal;

        // array fields
        BOOL*       BooleanArr;
        INT8*       SByteArr;
        INT16*      Int16Arr;
        INT32*      Int32Arr;
        INT64*      Int64Arr;
        UINT8*      ByteArr;
        UINT16*     UInt16Arr;
        UINT32*     UInt32Arr;
        UINT64*     UInt64Arr;
        float*      SingleArr;
        double*     DoubleArr;
        FILETIME*   FileTimeArr;
        SYSTEMTIME* SysTimeArr;
        GUID*       GuidArr;
        LPWSTR*     StringArr;
        LPSTR*      AnsiStringArr;
        PSID*       SidArr;
        size_t*     SizeTArr;

        // internal fields
        HANDLE      EvtHandleVal;
        LPCWSTR     XmlVal;
        LPCWSTR*    XmlValArr;
    };

    DWORD Count;   // number of elements (not length) in bytes.
    DWORD Type;

} EVT_VARIANT;

typedef enum _EVT_CHANNEL_CONFIG_PROPERTY_ID {
  EvtChannelConfigEnabled                 = 0,
  EvtChannelConfigIsolation               = 1,
  EvtChannelConfigType                    = 2,
  EvtChannelConfigOwningPublisher         = 3,
  EvtChannelConfigClassicEventlog         = 4,
  EvtChannelConfigAccess                  = 5,
  EvtChannelLoggingConfigRetention        = 6,
  EvtChannelLoggingConfigAutoBackup       = 7,
  EvtChannelLoggingConfigMaxSize          = 8,
  EvtChannelLoggingConfigLogFilePath      = 9,
  EvtChannelPublishingConfigLevel         = 10,
  EvtChannelPublishingConfigKeywords      = 11,
  EvtChannelPublishingConfigControlGuid   = 12,
  EvtChannelPublishingConfigBufferSize    = 13,
  EvtChannelPublishingConfigMinBuffers    = 14,
  EvtChannelPublishingConfigMaxBuffers    = 15,
  EvtChannelPublishingConfigLatency       = 16,
  EvtChannelPublishingConfigClockType     = 17,
  EvtChannelPublishingConfigSidType       = 18,
  EvtChannelPublisherList                 = 18,
  EvtChannelPublishingConfigFileMax       = 19,
  EvtChannelConfigPropertyIdEND           = 20 
} EVT_CHANNEL_CONFIG_PROPERTY_ID;

static struct {
    HANDLE (WINAPI *_EvtOpenSession)(int, PVOID, DWORD, DWORD);
    BOOL (WINAPI *_EvtClose)(HANDLE);
    BOOL (WINAPI *_EvtCancel)(HANDLE);
    DWORD (WINAPI *_EvtGetExtendedStatus)(DWORD, LPWSTR, PDWORD);
    HANDLE (WINAPI *_EvtQuery)(HANDLE, LPCWSTR, LPCWSTR, DWORD);
    BOOL (WINAPI *_EvtNext)(HANDLE, DWORD, HANDLE *, DWORD, DWORD, DWORD *);
    BOOL (WINAPI *_EvtSeek)(HANDLE, LONGLONG, HANDLE, DWORD, DWORD);
    HANDLE (WINAPI *_EvtSubscribe)(HANDLE,HANDLE,LPCWSTR,LPCWSTR,HANDLE,PVOID,EVT_SUBSCRIBE_CALLBACK, DWORD);
    HANDLE (WINAPI *_EvtCreateRenderContext)(DWORD, LPCWSTR*, DWORD);
    BOOL (WINAPI *_EvtRender)(HANDLE,HANDLE,DWORD,DWORD,PVOID, PDWORD, PDWORD);
    BOOL (WINAPI *_EvtFormatMessage)(HANDLE,HANDLE,DWORD,DWORD, EVT_VARIANT*,DWORD,DWORD, LPWSTR, PDWORD);
    HANDLE (WINAPI *_EvtOpenLog)(HANDLE,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtGetLogInfo)(HANDLE,int,DWORD,EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtClearLog)(HANDLE,LPCWSTR,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtExportLog)(HANDLE,LPCWSTR,LPCWSTR,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtArchiveExportedLog)(HANDLE,LPCWSTR,LCID,DWORD);
    HANDLE (WINAPI *_EvtOpenChannelEnum)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtNextChannelPath)(HANDLE,DWORD, LPWSTR, PDWORD);
    HANDLE (WINAPI *_EvtOpenChannelConfig)(HANDLE,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtSaveChannelConfig)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtSetChannelConfigProperty)(HANDLE, int, DWORD, EVT_VARIANT *);
    BOOL (WINAPI *_EvtGetChannelConfigProperty)(HANDLE,int,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    HANDLE (WINAPI *_EvtOpenPublisherEnum)(HANDLE,DWORD);
    HANDLE (WINAPI *_EvtOpenEventMetadataEnum)(HANDLE,DWORD);
    HANDLE (WINAPI *_EvtNextEventMetadata)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtNextPublisherId)(HANDLE,DWORD,LPWSTR,PDWORD);
    BOOL (WINAPI *_EvtGetPublisherMetadataProperty)(HANDLE,int,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtGetEventMetadataProperty)(HANDLE,int,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    HANDLE (WINAPI *_EvtOpenPublisherMetadata)(HANDLE,LPCWSTR,LPCWSTR,LCID,DWORD);
    BOOL (WINAPI *_EvtGetObjectArraySize)(HANDLE, PDWORD);
    BOOL (WINAPI *_EvtGetObjectArrayProperty)(HANDLE,DWORD,DWORD,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtGetQueryInfo)(HANDLE,int,DWORD,EVT_VARIANT *,PDWORD);
    BOOL (WINAPI *_EvtGetEventInfo)(HANDLE,int,DWORD,EVT_VARIANT *, PDWORD);
    HANDLE (WINAPI *_EvtCreateBookmark)(LPCWSTR);
    BOOL (WINAPI *_EvtUpdateBookmark)(HANDLE,HANDLE);
} gEvtStubs;

#define EvtOpenSession gEvtStubs._EvtOpenSession
#define EvtClose gEvtStubs._EvtClose
#define EvtCancel gEvtStubs._EvtCancel
#define EvtGetExtendedStatus gEvtStubs._EvtGetExtendedStatus
#define EvtQuery gEvtStubs._EvtQuery
#define EvtNext gEvtStubs._EvtNext
#define EvtSeek gEvtStubs._EvtSeek
#define EvtSubscribe gEvtStubs._EvtSubscribe
#define EvtCreateRenderContext gEvtStubs._EvtCreateRenderContext
#define EvtRender gEvtStubs._EvtRender
#define EvtFormatMessage gEvtStubs._EvtFormatMessage
#define EvtOpenLog gEvtStubs._EvtOpenLog
#define EvtGetLogInfo gEvtStubs._EvtGetLogInfo
#define EvtClearLog gEvtStubs._EvtClearLog
#define EvtExportLog gEvtStubs._EvtExportLog
#define EvtArchiveExportedLog gEvtStubs._EvtArchiveExportedLog
#define EvtOpenChannelEnum gEvtStubs._EvtOpenChannelEnum
#define EvtNextChannelPath gEvtStubs._EvtNextChannelPath
#define EvtOpenChannelConfig gEvtStubs._EvtOpenChannelConfig
#define EvtSaveChannelConfig gEvtStubs._EvtSaveChannelConfig
#define EvtSetChannelConfigProperty gEvtStubs._EvtSetChannelConfigProperty
#define EvtGetChannelConfigProperty gEvtStubs._EvtGetChannelConfigProperty
#define EvtOpenPublisherEnum gEvtStubs._EvtOpenPublisherEnum
#define EvtNextPublisherId gEvtStubs._EvtNextPublisherId
#define EvtOpenPublisherMetadata gEvtStubs._EvtOpenPublisherMetadata
#define EvtGetPublisherMetadataProperty gEvtStubs._EvtGetPublisherMetadataProperty
#define EvtOpenEventMetadataEnum gEvtStubs._EvtOpenEventMetadataEnum
#define EvtNextEventMetadata gEvtStubs._EvtNextEventMetadata
#define EvtGetEventMetadataProperty gEvtStubs._EvtGetEventMetadataProperty
#define EvtGetObjectArraySize gEvtStubs._EvtGetObjectArraySize
#define EvtGetObjectArrayProperty gEvtStubs._EvtGetObjectArrayProperty
#define EvtGetQueryInfo gEvtStubs._EvtGetQueryInfo
#define EvtCreateBookmark gEvtStubs._EvtCreateBookmark
#define EvtUpdateBookmark gEvtStubs._EvtUpdateBookmark
#define EvtGetEventInfo gEvtStubs._EvtGetEventInfo


#endif /* RUNTIME_EVT_LOAD */

typedef HANDLE EVT_HANDLE;
int gEvtStatus;                 /* 0 - init, 1 - available, -1 - unavailable  */
EVT_HANDLE gEvtDllHandle;

/* Used as a typedef for returning allocated memory to script level */
#define TWAPI_EVT_RENDER_VALUES_TYPESTR "EVT_RENDER_VALUES *"

#define ObjFromEVT_HANDLE(h_) ObjFromOpaque((h_), "EVT_HANDLE")
#define GETEVTH(h_) GETHANDLET((h_), EVT_HANDLE)

/*
 * Used to hold renderered values to be passed to script level
 * Actual buffer follows the header.
 */
typedef union _TwapiEVT_RENDER_VALUES_HEADER {
    void *align;            /* Align following buffer to quadword */
    struct {
        DWORD sz;           /* Size of following buffer */
        DWORD used;         /* Bytes used in following buffer */
        DWORD count;        /* Count of EVT_VARIANT values in buffer */
    } header;
} TwapiEVT_RENDER_VALUES_HEADER;
#define ERVHP_BUFFER(ervhp_) ((EVT_VARIANT *)(sizeof(*ervhp_) + (char *) (ervhp_)))

/* Always returns TCL_ERROR after storing extended error info */
static TCL_RESULT Twapi_AppendEvtExtendedStatus(Tcl_Interp *interp)
{
    DWORD sz, used;
    LPWSTR bufP;

    if (EvtGetExtendedStatus(0, NULL, &sz) != FALSE)
        return TCL_ERROR;       /* No additional info available */

    
    bufP = TwapiAlloc(sizeof(WCHAR) * sz);
    if (EvtGetExtendedStatus(sz, bufP, &used) != FALSE && used != 0) {
        /* TBD - verify this works (is bufP null terminated ?) */
        Tcl_AppendResult(interp, " ", bufP, NULL);
    }
    TwapiFree(bufP);

    return TCL_ERROR;           /* Always returns TCL_ERROR */
}


static Tcl_Obj *ObjFromEVT_VARIANT(TwapiInterpContext *ticP, EVT_VARIANT *varP,
                                   int flags) /* flags & 1 => returned tagged value */
{
    int i;
    Tcl_Obj *objP;
    Tcl_Obj **objPP;
    Tcl_Obj *retObjs[2];
    int count;

    objP = NULL;
    switch (varP->Type) {
    case EvtVarTypeNull:
        break;
    case EvtVarTypeString:
    case EvtVarTypeEvtXml:
        if (varP->StringVal)
            objP = ObjFromUnicode(varP->StringVal);
        break;
    case EvtVarTypeAnsiString:
        if (varP->AnsiStringVal)
            objP = ObjFromString(varP->AnsiStringVal);
        break;
    case EvtVarTypeSByte:
        objP = ObjFromInt(varP->SByteVal);
        break;
    case EvtVarTypeByte:
        objP = ObjFromInt(varP->ByteVal);
        break;
    case EvtVarTypeInt16:
        objP = ObjFromInt(varP->Int16Val);
        break;
    case EvtVarTypeUInt16:
        objP = ObjFromInt(varP->UInt16Val);
        break;
    case EvtVarTypeInt32:
        objP = ObjFromInt(varP->Int32Val);
        break;
    case EvtVarTypeUInt32:
        objP = ObjFromDWORD(varP->UInt32Val);
        break;
    case EvtVarTypeInt64:
        objP = ObjFromWideInt(varP->Int64Val);
        break;
    case EvtVarTypeUInt64:
        objP = ObjFromULONGLONG(varP->UInt64Val);
        break;
    case EvtVarTypeSingle:
        objP =  Tcl_NewDoubleObj(varP->SingleVal);
        break;
    case EvtVarTypeDouble:
        objP =  Tcl_NewDoubleObj(varP->DoubleVal);
        break;
    case EvtVarTypeBoolean:
        objP = ObjFromBoolean(varP->BooleanVal != 0);
        break;
    case EvtVarTypeBinary:      /* TBD - do not know how to interpret this  */
        break;
    case EvtVarTypeGuid:
        objP = ObjFromGUID(varP->GuidVal); /* OK if NULL */
        break;
    case EvtVarTypeSizeT:
        objP = ObjFromSIZE_T(varP->SizeTVal);
        break;
    case EvtVarTypeFileTime:
        objP = ObjFromULONGLONG(varP->FileTimeVal);
        break;
    case EvtVarTypeSysTime:
        if (varP->SysTimeVal)
            objP = ObjFromSYSTEMTIME(varP->SysTimeVal);
        break;
    case EvtVarTypeSid:
        objP = ObjFromSIDNoFail(varP->SidVal);
        break;
    case EvtVarTypeEvtHandle:
        objP = ObjFromEVT_HANDLE(varP->EvtHandleVal);
        break;
    case EvtVarTypeHexInt32:      /* TBD - do not know how to interpret this  */
        break;
    case EvtVarTypeHexInt64:      /* TBD - do not know how to interpret this  */
        break;
    default:
        /* Check if an array. */
        if ((varP->Type & EVT_VARIANT_TYPE_ARRAY) == 0)
            break;

        /* Check count and non-null pointer. Union so check any field */
        count = varP->Count;
        if (count == 0 || varP->BooleanArr == NULL)
            break;

        objPP = MemLifoPushFrame(&ticP->memlifo,
                                 count * sizeof(objPP[0]), NULL);

        switch (varP->Type & EVT_VARIANT_TYPE_MASK) {
        case EvtVarTypeString:
        case EvtVarTypeEvtXml:
            for (i = 0; i < count; ++i) {
                objPP[i] = varP->StringArr[i] ?
                    ObjFromUnicode(varP->StringArr[i])
                    : ObjFromEmptyString();
            }
            break;
        case EvtVarTypeAnsiString:
            for (i = 0; i < count; ++i) {
                objPP[i] = varP->AnsiStringArr[i] ?
                    ObjFromString(varP->AnsiStringArr[i])
                    : ObjFromEmptyString();
            }
            break;
        case EvtVarTypeSByte:
            /* TBD - should this be a byte array ? */
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromInt(varP->SByteArr[i]);
            }
            break;
        case EvtVarTypeByte:
            /* TBD - should this be a byte array ? */
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromInt(varP->ByteArr[i]);
            }
            break;
        case EvtVarTypeInt16:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromInt(varP->Int16Arr[i]);
            }
            break;
        case EvtVarTypeUInt16:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromInt(varP->UInt16Arr[i]);
            }
            break;
        case EvtVarTypeInt32:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromInt(varP->Int32Arr[i]);
            }
            break;
        case EvtVarTypeUInt32:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromDWORD(varP->UInt32Arr[i]);
            }
            break;
        case EvtVarTypeInt64:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromWideInt(varP->Int64Arr[i]);
            }
            break;
        case EvtVarTypeUInt64:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromULONGLONG(varP->UInt64Arr[i]);
            }
            break;
        case EvtVarTypeSingle:
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewDoubleObj(varP->SingleArr[i]);
            }
            break;
        case EvtVarTypeDouble:
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewDoubleObj(varP->DoubleArr[i]);
            }
            break;
        case EvtVarTypeBoolean:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromBoolean(varP->BooleanArr[i] != 0);
            }
            break;
        case EvtVarTypeGuid:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromGUID(&varP->GuidArr[i]);
            }
            break;
        case EvtVarTypeSizeT:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromSIZE_T(varP->SizeTArr[i]);
            }
            break;
        case EvtVarTypeFileTime:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromFILETIME(&varP->FileTimeArr[i]);
            }
            break;
        case EvtVarTypeSysTime:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromSYSTEMTIME(&varP->SysTimeArr[i]);
            }
            break;
        case EvtVarTypeSid:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromSIDNoFail(varP->SidArr[i]);
            }
            break;
        case EvtVarTypeNull:
        case EvtVarTypeBinary: /* TBD */
        case EvtVarTypeHexInt32: /* TBD */
        case EvtVarTypeHexInt64: /* TBD */
        default:
            /* Stuff that we do not handle or is unknown. Create an
             * array with that many elements
             */
            TWAPI_ASSERT(count != 0);
            objPP[0] = ObjFromEmptyString();
            for (i = 1; i < count; ++i) {
                objPP[i] = objPP[0];
            }
            break;
        }
        TWAPI_ASSERT(count > 0 && objPP);
        objP = ObjNewList(count, objPP);
        MemLifoPopFrame(&ticP->memlifo);
        break;
    }

    if (objP == NULL) {
        // TBD - debug output
        objP = ObjFromEmptyString();
    }

    if (flags) {
        retObjs[0] = ObjFromInt(varP->Type);
        retObjs[1] = objP;
        return ObjNewList(2, retObjs);
    } else {
        return objP;
    }
}

static Tcl_Obj *ObjFromEVT_VARIANT_ARRAY(TwapiInterpContext *ticP, EVT_VARIANT *varP, int count)
{
    int i;
    Tcl_Obj **objPP;
    Tcl_Obj *objP;

    objPP = MemLifoPushFrame(&ticP->memlifo, count * sizeof (objPP[0]), NULL);

    for (i = 0; i < count; ++i) {
        objPP[i] = ObjFromEVT_VARIANT(ticP, &varP[i], 0);
    }

    objP = ObjNewList(count, objPP);
    MemLifoPopFrame(&ticP->memlifo);

    return objP;
}


/* Should be called only once at init time */
void TwapiInitEvtStubs(Tcl_Interp *interp)
{
#ifdef RUNTIME_EVT_LOAD

    UINT len;
    WCHAR path[MAX_PATH+1];

    if (gEvtStatus)
        return;                 /* Already called */

    /* Assume failure */
    gEvtStatus = -1;

    len = GetSystemDirectoryW(path, ARRAYSIZE(path));
    if (len == 0)
        return; /* Error ignored. Functions will simply not be available */

#define WEVTDLLNAME L"\\wevtapi.dll"
    if (len > (ARRAYSIZE(path) - ARRAYSIZE(WEVTDLLNAME)))
        return;

    lstrcpynW(&path[len], WEVTDLLNAME, ARRAYSIZE(WEVTDLLNAME));

    /* Initialize the Evt stubs */
    gEvtDllHandle = LoadLibraryW(path);
    if (gEvtDllHandle == NULL)
        return;                 /* DLL not available */
    
#define INIT_EVT_STUB(fn) \
    do { \
      if (((gEvtStubs._ ## fn) = (void *)GetProcAddress(gEvtDllHandle, #fn)) == NULL) \
            return;                                                     \
    } while (0)

    INIT_EVT_STUB(EvtOpenSession);
    INIT_EVT_STUB(EvtClose);
    INIT_EVT_STUB(EvtCancel);
    INIT_EVT_STUB(EvtGetExtendedStatus);
    INIT_EVT_STUB(EvtQuery);
    INIT_EVT_STUB(EvtNext);
    INIT_EVT_STUB(EvtSeek);
    INIT_EVT_STUB(EvtSubscribe);
    INIT_EVT_STUB(EvtCreateRenderContext);
    INIT_EVT_STUB(EvtRender);
    INIT_EVT_STUB(EvtFormatMessage);
    INIT_EVT_STUB(EvtOpenLog);
    INIT_EVT_STUB(EvtGetLogInfo);
    INIT_EVT_STUB(EvtClearLog);
    INIT_EVT_STUB(EvtExportLog);
    INIT_EVT_STUB(EvtArchiveExportedLog);
    INIT_EVT_STUB(EvtOpenChannelEnum);
    INIT_EVT_STUB(EvtNextChannelPath);
    INIT_EVT_STUB(EvtOpenChannelConfig);
    INIT_EVT_STUB(EvtSaveChannelConfig);
    INIT_EVT_STUB(EvtSetChannelConfigProperty);
    INIT_EVT_STUB(EvtGetChannelConfigProperty);
    INIT_EVT_STUB(EvtOpenPublisherEnum);
    INIT_EVT_STUB(EvtNextPublisherId);
    INIT_EVT_STUB(EvtOpenPublisherMetadata);
    INIT_EVT_STUB(EvtGetPublisherMetadataProperty);
    INIT_EVT_STUB(EvtOpenEventMetadataEnum);
    INIT_EVT_STUB(EvtNextEventMetadata);
    INIT_EVT_STUB(EvtGetEventMetadataProperty);
    INIT_EVT_STUB(EvtGetObjectArraySize);
    INIT_EVT_STUB(EvtGetObjectArrayProperty);
    INIT_EVT_STUB(EvtGetQueryInfo);
    INIT_EVT_STUB(EvtCreateBookmark);
    INIT_EVT_STUB(EvtUpdateBookmark);
    INIT_EVT_STUB(EvtGetEventInfo);
#undef INIT_EVT_STUB

#endif // RUNTIME_EVT_LOAD

    /* Success */
    gEvtStatus = 1;
}

/* IMPORTANT:
 * If a valid buffer is passed in (as the 4th arg) caller must not
 * access it again irrespective of successful or error return unless
 * in the former case the same buffer is returned explicitly
 */
static TCL_RESULT Twapi_EvtRenderValuesObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE hevt, hevt2;
    DWORD status;
    void *bufP;
    TwapiEVT_RENDER_VALUES_HEADER *ervhP;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETEVTH(hevt),
                     GETHANDLET(hevt2, EVT_HANDLE),
                     GETPTR(ervhP, TwapiEVT_RENDER_VALUES_HEADER*),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* 4th arg is supposed to describe a previously returned buffer
       that we can reuse. It may also be NULL
    */
    if (ervhP) {
        status = TwapiVerifyPointer(interp, ervhP, Twapi_EvtRenderValuesObjCmd);
        if (status != TWAPI_NO_ERROR)
            return TwapiReturnError(interp, status);
    } else {
        /* Need to allocate buffer */

        /* TBD - instrument reallocation needs */
        ervhP = TwapiAllocRegisteredPointer(interp, sizeof(*ervhP) + 4000, Twapi_EvtRenderValuesObjCmd);
        ervhP->header.sz = 4000;
    }

    bufP = ERVHP_BUFFER(ervhP);

    /* We used to convert using ObjFromEVT_VARIANT but that does
       not work well with opaque values so we preserve as a
       binary blob to be passed around. Note we cannot use
       a Tcl_ByteArray either because the embedded pointers will
       be invalid when the byte array is copied around.
    */

    status = ERROR_SUCCESS;
    if (EvtRender(hevt, hevt2,
                  0,    /* EvtRenderEventValues -> 0 */
                  ervhP->header.sz, bufP,
                  &ervhP->header.used, &ervhP->header.count) == FALSE) {
        status = GetLastError();
        if (status == ERROR_INSUFFICIENT_BUFFER) {
            DWORD new_sz = ervhP->header.used;
            TwapiFreeRegisteredPointer(interp, ervhP, Twapi_EvtRenderValuesObjCmd);
            ervhP = TwapiAllocRegisteredPointer(interp, sizeof(*ervhP) + new_sz, Twapi_EvtRenderValuesObjCmd);
            ervhP->header.sz = new_sz;
            bufP = ERVHP_BUFFER(ervhP);
            status = ERROR_SUCCESS;
            if (EvtRender(hevt, hevt2, 0, ervhP->header.sz,
                          bufP, &ervhP->header.used, &ervhP->header.count) == FALSE) {
                status = GetLastError();
            }
        }
    }

    if (status != ERROR_SUCCESS) {
        TwapiFreeRegisteredPointer(interp, ervhP, Twapi_EvtRenderValuesObjCmd);
        return Twapi_AppendSystemError(interp, status);
    }


    TwapiSetObjResult(interp, ObjFromOpaque(ervhP, "TwapiEVT_RENDER_VALUES_HEADER*"));
    return TCL_OK;
}

/* EvtRender for Unicode return types */
static TCL_RESULT Twapi_EvtRenderUnicodeObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp,  int objc, Tcl_Obj *CONST objv[])
{
    HANDLE hevt, hevt2;
    DWORD flags, sz, count, status;
    void *bufP;
    Tcl_Obj *objP;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETEVTH(hevt),
                     GETHANDLET(hevt2, EVT_HANDLE), GETINT(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* 1 -> EvtRenderEventXml */
    /* 2 -> EvtRenderBookmark */
    if (flags != 1 && flags != 2)
        return TwapiReturnError(interp,TWAPI_INVALID_ARGS);

    /* TBD - instrument reallocation needs */
    sz = 256;
    bufP = MemLifoPushFrame(&ticP->memlifo, sz, &sz);
    status = ERROR_SUCCESS;
    if (EvtRender(hevt, hevt2, flags, sz, bufP, &sz, &count) == FALSE) {
        status = GetLastError();
        if (status == ERROR_INSUFFICIENT_BUFFER) {
            /* Note no need to MemlifoPopFrame before allocating more */
            bufP = MemLifoAlloc(&ticP->memlifo, sz, &sz);
            if (EvtRender(hevt, hevt2, flags, sz, bufP, &sz, &count) == FALSE)
                status = GetLastError();
            else
                status = ERROR_SUCCESS;
        }
    }

    if (status != ERROR_SUCCESS) {
        MemLifoPopFrame(&ticP->memlifo);
        return Twapi_AppendSystemError(interp, status);
    }

    /* Unicode string. Should we use sz/2 instead of -1 ? TBD */
    objP = ObjFromUnicode(bufP);
    MemLifoPopFrame(&ticP->memlifo);
    TwapiSetObjResult(ticP->interp, objP);
    return TCL_OK;
}

static TCL_RESULT Twapi_ExtractEVT_VARIANT_ARRAYObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    EVT_VARIANT *varP;
    int dw;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETHANDLE(varP), GETINT(dw),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
        
    Tcl_SetObjResult(interp, ObjFromEVT_VARIANT_ARRAY(ticP, varP, dw));
    return TCL_OK;
}

static TCL_RESULT Twapi_ExtractEVT_RENDER_VALUESObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiEVT_RENDER_VALUES_HEADER *ervhP;


    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVERIFIEDPTR(ervhP, TwapiEVT_RENDER_VALUES_HEADER*, Twapi_EvtRenderValuesObjCmd),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    Tcl_SetObjResult(interp, ObjFromEVT_VARIANT_ARRAY(ticP, ERVHP_BUFFER(ervhP), ervhP->header.count));
    return TCL_OK;
}

static TCL_RESULT Twapi_EvtNextObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    EVT_HANDLE hevt;
    EVT_HANDLE *hevtP;
    DWORD i, count, timeout, dw;
    Tcl_Obj **objPP;
    int result;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETEVTH(hevt), GETINT(count),
                     GETINT(timeout), GETINT(dw), ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (count > 1024) // TBD
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
    hevtP = MemLifoPushFrame(&ticP->memlifo, count*sizeof(*hevtP), NULL);
    if (EvtNext(hevt, count, hevtP, timeout, dw, &count) != FALSE) {
        if (count) {
            objPP = MemLifoAlloc(&ticP->memlifo, count*sizeof(*objPP), NULL);
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromEVT_HANDLE(hevtP[i]);
            }
            TwapiSetObjResult(interp, ObjNewList(count, objPP));
        }
        result = TCL_OK;
    } else {
        dw = GetLastError();
        if (dw == ERROR_NO_MORE_ITEMS || dw == ERROR_TIMEOUT)
            result = TCL_OK;
        else {
            Twapi_AppendSystemError(interp, dw);
            result = TCL_ERROR;
        }
    }
    MemLifoPopFrame(&ticP->memlifo);
    return result;
}

static TCL_RESULT Twapi_EvtCreateRenderContextObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    EVT_HANDLE hevt;
    int count;
    LPCWSTR *xpathsP = NULL;
    int flags;
    DWORD ret = TCL_ERROR;

    if (TwapiGetArgs(interp, objc-1, objv+1, ARGSKIP, GETINT(flags), ARGEND) != TCL_OK ||
        Tcl_ListObjLength(interp, objv[1], &count) != TCL_OK)
        return TCL_ERROR;

    if (count == 0) {
        xpathsP = NULL;
    } else {
        /* Note ObjToArgvW needs an extra entry for terminating NULL */
        xpathsP = MemLifoPushFrame(&ticP->memlifo, (count+1) * sizeof(xpathsP[0]), NULL);
        if (ObjToArgvW(interp, objv[1], xpathsP, count+1, &count) != TCL_OK)
            goto vamoose;
    }
    
    hevt = EvtCreateRenderContext(count, xpathsP, flags);
    if (hevt == NULL) {
        TwapiReturnSystemError(interp);
        goto vamoose;
    }
        
    TwapiSetObjResult(interp, ObjFromEVT_HANDLE(hevt));
    ret = TCL_OK;

vamoose:
    if (xpathsP)
        MemLifoPopFrame(&ticP->memlifo);
    return ret;
}

static TCL_RESULT Twapi_EvtFormatMessageObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    EVT_HANDLE hpub, hev;
    DWORD msgid, flags;
    EVT_VARIANT *valuesP;
    int nvalues;
    WCHAR buf[500];             /* TBD - instrument */
    int buf_sz, used;
    WCHAR *bufP;
    DWORD winerr;
    TwapiEVT_RENDER_VALUES_HEADER *ervhP;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHANDLET(hpub, EVT_HANDLE),
                     GETHANDLET(hev, EVT_HANDLE),
                     GETINT(msgid),
                     GETPTR(ervhP, TwapiEVT_RENDER_VALUES_HEADER*),
                     GETINT(flags),
                     ARGUSEDEFAULT, ARGSKIP, ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (ervhP) {
        winerr = TwapiVerifyPointer(interp, ervhP, Twapi_EvtRenderValuesObjCmd);
        if (winerr != TWAPI_NO_ERROR)
            return TwapiReturnError(interp, winerr);
        nvalues = ervhP->header.count;
        valuesP = ERVHP_BUFFER(ervhP);
    } else {
        nvalues = 0;
        valuesP = NULL;
    }

    /* TBD - instrument buffer size */
    bufP = buf;
    buf_sz = ARRAYSIZE(buf);
    winerr = ERROR_SUCCESS;
    /* Note buffer sizes are in WCHARs, not bytes */
    if (EvtFormatMessage(hpub, hev, msgid, nvalues, valuesP, flags, buf_sz, bufP, &used) == FALSE) {
        winerr = GetLastError();
        if (winerr == ERROR_INSUFFICIENT_BUFFER) {
            buf_sz = used;
            bufP = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR)*buf_sz, NULL);
            if (EvtFormatMessage(hpub, hev, msgid, nvalues, valuesP, flags, buf_sz, bufP, &used) == FALSE) {
                winerr = GetLastError();
            } else {
                winerr = ERROR_SUCCESS;
            }
        }
    }        

    /* For some error codes, the buffer is actually filled with 
       as much of the message as can be resolved.
    */
    switch (winerr) {
    case 15029: // ERROR_EVT_UNRESOLVED_VALUE_INSERT
    case 15030: // ERROR_EVT_UNRESOLVED_PARAMTER_INSERT
    case 15031: // ERROR_EVT_MAX_INSERTS_REACHED
        /* Sanity check */
        if (used && used <= buf_sz) {
            /* TBD - debug log */
            bufP[used-1] = 0; /* Ensure null termination */
            winerr = ERROR_SUCCESS; /* Treat as success case */
        }
    }

    if (winerr == ERROR_SUCCESS) {
        /* See comments in GetMessageString function at
           http://msdn.microsoft.com/en-us/windows/dd996923%28v=vs.85%29
           If flags == EvtFormatMessageKeyword,  the buffer may contain
           multiple concatenated null terminated keywords. */
        if (flags == 5 /* EvtFormatMessageKeyword */ ) {
            TwapiSetObjResult(interp, ObjFromMultiSz(bufP, used));
        } else {
            /* For other cases, like xml, used may be more than last char
               so depend on null termination */
            TwapiSetObjResult(interp, ObjFromUnicode(bufP));
        }
    } else {
        if (objc == 7) {
            /* Caller has specified errors should be ignored and
               the specified string returned instead
               TBD - log ?
            */
            winerr = ERROR_SUCCESS;
            TwapiSetObjResult(interp, objv[6]);
        } else
            Twapi_AppendSystemError(interp, winerr);
    }
    if (bufP != buf)
        MemLifoPopFrame(&ticP->memlifo);

    return winerr == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
}

static TCL_RESULT Twapi_EvtFormatMessage2ObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    EVT_HANDLE hpub, hev;
    DWORD msgid, flags;
    EVT_VARIANT *valuesP;
    int nvalues;
    WCHAR buf[500];             /* TBD - instrument */
    int buf_sz, used;
    WCHAR *bufP;
    DWORD winerr;
    TwapiEVT_RENDER_VALUES_HEADER *ervhP;
    Tcl_Obj *objP;
    TCL_RESULT status;

    /* objv[6], if specified, is the name of the variable to store
       message. If unspecified, message is returned in interp result. */
    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHANDLET(hpub, EVT_HANDLE),
                     GETHANDLET(hev, EVT_HANDLE),
                     GETINT(msgid),
                     GETPTR(ervhP, TwapiEVT_RENDER_VALUES_HEADER*),
                     GETINT(flags),
                     ARGUSEDEFAULT, ARGSKIP, ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (ervhP) {
        winerr = TwapiVerifyPointer(interp, ervhP, Twapi_EvtRenderValuesObjCmd);
        if (winerr != TWAPI_NO_ERROR)
            return TwapiReturnError(interp, winerr);
        nvalues = ervhP->header.count;
        valuesP = ERVHP_BUFFER(ervhP);
    } else {
        nvalues = 0;
        valuesP = NULL;
    }

    /* TBD - instrument buffer size */
    bufP = buf;
    buf_sz = ARRAYSIZE(buf);
    winerr = ERROR_SUCCESS;
    /* Note buffer sizes are in WCHARs, not bytes */
    if (EvtFormatMessage(hpub, hev, msgid, nvalues, valuesP, flags, buf_sz, bufP, &used) == FALSE) {
        winerr = GetLastError();
        if (winerr == ERROR_INSUFFICIENT_BUFFER) {
            buf_sz = used;
            bufP = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR)*buf_sz, NULL);
            if (EvtFormatMessage(hpub, hev, msgid, nvalues, valuesP, flags, buf_sz, bufP, &used) == FALSE) {
                winerr = GetLastError();
            } else {
                winerr = ERROR_SUCCESS;
            }
        }
    }        

    /* For some error codes, the buffer is actually filled with 
       as much of the message as can be resolved.
    */
    switch (winerr) {
    case 15029: // ERROR_EVT_UNRESOLVED_VALUE_INSERT
    case 15030: // ERROR_EVT_UNRESOLVED_PARAMTER_INSERT
    case 15031: // ERROR_EVT_MAX_INSERTS_REACHED
        /* Sanity check */
        if (used && used <= buf_sz) {
            /* TBD - debug log */
            bufP[used-1] = 0; /* Ensure null termination */
            winerr = ERROR_SUCCESS; /* Treat as success case */
        }
    }
    objP = NULL;
    if (winerr == ERROR_SUCCESS) {
        /* See comments in GetMessageString function at
           http://msdn.microsoft.com/en-us/windows/dd996923%28v=vs.85%29
           If flags == EvtFormatMessageKeyword,  the buffer may contain
           multiple concatenated null terminated keywords. */
        status = TCL_OK;
        if (flags == 5 /* EvtFormatMessageKeyword */ ) {
            objP = ObjFromMultiSz(bufP, used);
        } else {
            /* For other cases, like xml, used may be more than last char
               so depend on null termination, not used count.
               TBD - for performance reasons, verify this and may be
               make exception for xml only
            */
            objP = ObjFromUnicode(bufP);
        }
    } else {
        if (objc == 7) {
            objP = Twapi_MapWindowsErrorToString(winerr);
            status = TCL_OK;
        } else {
            Twapi_AppendSystemError(interp, winerr);
            status = TCL_ERROR;
        }
    }

    if (status == TCL_OK) {
        if (objc == 7) {
            TWAPI_ASSERT(objP != NULL);
            /* Set the value of the variable to the message or the error string */
            if (Tcl_ObjSetVar2(interp, objv[6], NULL, objP, TCL_LEAVE_ERR_MSG) == NULL) {
                Twapi_FreeNewTclObj(objP);
                status = TCL_ERROR; /* Invalid variable */
            }
            else {
                TwapiSetObjResult(interp, ObjFromInt(winerr == ERROR_SUCCESS));
            }
        } else
            TwapiSetObjResult(interp, objP);
    }

    if (bufP != buf)
        MemLifoPopFrame(&ticP->memlifo);

    return status;
}


static TCL_RESULT Twapi_EvtGetEVT_VARIANTObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    EVT_HANDLE hevt;
    int sz;
    EVT_VARIANT *varP;
    int func;
    DWORD dw, dw2, dw3;
    DWORD status;
    BOOL (WINAPI *fn3args)(HANDLE,int,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *fn2args)(HANDLE,int,DWORD,EVT_VARIANT *, PDWORD);

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(func), GETEVTH(hevt),
                     GETINT(dw), ARGUSEDEFAULT, GETINT(dw2),
                     GETINT(dw3), ARGEND) != TCL_OK)
        return TCL_ERROR;

    varP = MemLifoPushFrame(&ticP->memlifo, sizeof(EVT_VARIANT), &sz);
    while (1) {
        switch (func) {
        case 2:
        case 3:
        case 4:
            switch (func) {
            case 2: fn3args = EvtGetChannelConfigProperty; break;
            case 3: fn3args = EvtGetPublisherMetadataProperty; break;
            case 4: fn3args = EvtGetEventMetadataProperty; break;
            }
            status = fn3args(hevt, dw, dw2, sz, varP, &sz);
            break;
        case 5:
            status = EvtGetObjectArrayProperty(hevt, dw, dw2, dw3, sz, varP, &sz);
            break;
        case 6:
        case 7:
        case 8:
            switch (func) {
            case 6: fn2args = EvtGetQueryInfo; break;
            case 7: fn2args = EvtGetEventInfo; break;
            case 8: fn2args = EvtGetLogInfo; break;
            }                
            status = fn2args(hevt, dw, sz, varP, &sz);
            break;

        default:
            MemLifoPopFrame(&ticP->memlifo);
            return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);
        }        
        if (status != FALSE || GetLastError() != ERROR_INSUFFICIENT_BUFFER)
            break;
        /* Loop to retry larger buffer. No need to free previous alloc first */
        varP = MemLifoAlloc(&ticP->memlifo, sz, NULL);
    }

    if (status == FALSE) {
        TwapiReturnSystemError(interp);
        Twapi_AppendEvtExtendedStatus(interp);
    } else {
        TwapiSetObjResult(interp, ObjFromEVT_VARIANT(ticP, varP, 0));
    }
    MemLifoPopFrame(&ticP->memlifo);

    return status == FALSE ? TCL_ERROR : TCL_OK;
}


static TCL_RESULT Twapi_EvtOpenSessionObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int login_class;
    DWORD timeout, flags;
    Tcl_Obj **loginObjs;
    int nobjs;
    EVT_RPC_LOGIN erl;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(login_class),
                     ARGSKIP, ARGUSEDEFAULT, GETINT(timeout), GETINT(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (login_class != 1) {
        /* Only EvtRpcLogin (1) supported */
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid login class");
    }
    if (ObjGetElements(interp, objv[2], &nobjs, &loginObjs) != TCL_OK)
        return TCL_ERROR;

    if (TwapiGetArgs(interp, nobjs, loginObjs, GETWSTR(erl.Server),
                     GETNULLIFEMPTY(erl.User), GETNULLIFEMPTY(erl.Domain),
                     GETNULLIFEMPTY(erl.Password), GETINT(erl.Flags),
                     ARGEND) != TCL_OK) {
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid EVT_RPC_LOGIN structure");
    }

    return TwapiReturnNonnullHandle(interp,
                                    EvtOpenSession(login_class, &erl,
                                                   timeout, flags),
                                    "EVT_HANDLE");
}



int Twapi_EvtCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    DWORD dw, dw2;
    LPWSTR s, s2, s3;
    EVT_HANDLE hevt, hevt2;
    Tcl_WideInt wide;
    int func = PtrToInt(clientdata);
    HANDLE h;
    WCHAR buf[MAX_PATH+1];
    EVT_VARIANT var;
    GUID guid;
    
    if (gEvtStatus != 1)
        return Twapi_AppendSystemError(interp, ERROR_CALL_NOT_IMPLEMENTED);

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
    case 2:
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt), GETWSTR(s), GETNULLIFEMPTY(s2), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (func == 1) {
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EvtClearLog(hevt, s, s2, dw);
        } else {
            result.type = TRT_NONNULL;
            result.value.nonnull.name = "EVT_HANDLE";
            result.value.nonnull.p = EvtQuery(hevt, s, s2, dw);
        }
        break;
    case 3: // EvtSeek
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETWIDE(wide), GETHANDLET(hevt2, EVT_HANDLE),
                         GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtSeek(hevt, wide, hevt2, dw, dw2);
        break;
    case 4: // EvtOpenLog
    case 5: // EvtOpenChannelConfig
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETWSTR(s), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_NONNULL;
        result.value.nonnull.name = "EVT_HANDLE";
        result.value.nonnull.p = (func == 4 ? EvtOpenLog : EvtOpenChannelConfig) (hevt, s, dw);
        break;
    case 6: // EvtArchiveExportedLog
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETWSTR(s), GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtArchiveExportedLog(hevt, s, dw, dw2);
        break;
    case 7: // EvtSubscribe
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETHANDLE(h), GETWSTR(s), GETNULLIFEMPTY(s2),
                         GETEVTH(hevt2), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_NONNULL;
        result.value.nonnull.name = "EVT_HANDLE";
        result.value.nonnull.p = EvtSubscribe(hevt, h, s, s2, hevt2, NULL, NULL, dw);
        break;
    case 8: // EvtExportLog
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETWSTR(s), GETWSTR(s2), GETWSTR(s3), GETINT(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtExportLog(hevt, s, s2, s3, dw);
        break;

    case 9: // EvtSetChannelConfigProperty
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETINT(dw), GETINT(dw2), ARGSKIP,
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        var.Type = EvtVarTypeNull;
        var.Count = 0;
        /* dw is property id */
        switch (dw) {
        case EvtChannelConfigEnabled:
        case EvtChannelConfigClassicEventlog:
        case EvtChannelLoggingConfigRetention:
        case EvtChannelLoggingConfigAutoBackup:
            if (ObjToBoolean(interp, objv[3], &var.BooleanVal) != TCL_OK)
                return TCL_ERROR;
            var.Type = EvtVarTypeBoolean;
            break;
        case EvtChannelConfigIsolation:
        case EvtChannelConfigType:
        case EvtChannelPublishingConfigLevel:
        case EvtChannelPublishingConfigClockType:
        case EvtChannelPublishingConfigFileMax:
            if (ObjToLong(interp, objv[3], &var.UInt32Val) != TCL_OK)
                return TCL_ERROR;
            var.Type = EvtVarTypeUInt32;
            break;
        case EvtChannelConfigOwningPublisher:
        case EvtChannelConfigAccess:
        case EvtChannelLoggingConfigLogFilePath:
        case EvtChannelPublisherList:
            var.StringVal = ObjToUnicode(objv[3]);
            var.Type = EvtVarTypeString;
            break;

        case EvtChannelLoggingConfigMaxSize:
        case EvtChannelPublishingConfigKeywords:
            if (ObjToWideInt(interp, objv[3], &var.UInt64Val) != TCL_OK)
                return TCL_ERROR;
            var.Type = EvtVarTypeUInt64;
            break;
        case EvtChannelPublishingConfigControlGuid:
            if (ObjToGUID(interp, objv[3], &guid) != TCL_OK)
                return TCL_ERROR;
            var.Type = EvtVarTypeGuid;
            break;
        default:
            /* Note following properties cannot be set 
               case EvtChannelPublishingConfigBufferSize:
               case EvtChannelPublishingConfigMinBuffers:
               case EvtChannelPublishingConfigMaxBuffers:
               case EvtChannelPublishingConfigLatency:
               case EvtChannelPublishingConfigSidType:
            */
            return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
        }
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtSetChannelConfigProperty(hevt, dw, dw2, &var);
        break;

    case 10: // EvtOpenPublisherMetadata
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETWSTR(s), GETNULLIFEMPTY(s2), GETINT(dw),
                         GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_NONNULL;
        result.value.nonnull.name = "EVT_HANDLE";
        result.value.nonnull.p = EvtOpenPublisherMetadata(hevt, s, s2, dw, dw2);
        break;

    case 11: // evt_create_bookmark
        if (TwapiGetArgs(interp, objc, objv, ARGUSEDEFAULT,
                         GETWSTR(s), ARGEND) != TCL_OK)
            return TCL_ERROR;
        NULLIFY_EMPTY(s);
        result.type = TRT_NONNULL;
        result.value.nonnull.name = "EVT_HANDLE";
        result.value.nonnull.p = EvtCreateBookmark(s);
        break;

    case 12: // evt_update_bookmark
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt), GETEVTH(hevt2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtUpdateBookmark(hevt, hevt2);
        break;

    case 13: // evt_free
        if (TwapiGetArgs(interp, objc, objv, GETVERIFIEDVOIDP(h, NULL), ARGEND) != TCL_OK)
            return TCL_ERROR;
        TWAPI_ASSERT(h);
        TwapiFreeRegisteredPointer(interp, h, NULL);
        result.type = TRT_EMPTY;
        break;

    default:
        /* Params - HANDLE followed by optional DWORD */
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 101:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EvtClose(hevt);
            break;
        case 102:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EvtCancel(hevt);
            break;
        case 103:
            result.type = TRT_NONNULL;
            result.value.nonnull.name = "EVT_HANDLE";
            result.value.nonnull.p =EvtOpenChannelEnum(hevt, dw);
            break;
        case 104: // EvtNextChannelPath
        case 109: // EvtNextPublisherId
            /* Note channel/publisher is max 255 chars so no need to check
               for ERROR_INSUFFICIENT_BUFFER */
            if ((func == 104 ? EvtNextChannelPath : EvtNextPublisherId)(hevt, ARRAYSIZE(buf), buf, &dw) != FALSE) {
                TwapiSetObjResult(interp, ObjFromUnicodeN(buf, dw-1));
                return TCL_OK;
            }
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = GetLastError();
            if (result.value.ival == ERROR_NO_MORE_ITEMS)
                return TCL_OK;
            break;
        case 105:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EvtSaveChannelConfig(hevt, dw);
            break;
        case 106:
            result.type = TRT_NONNULL;
            result.value.nonnull.name = "EVT_HANDLE";
            result.value.nonnull.p = EvtOpenPublisherEnum(hevt, dw);
            break;
        case 107:
            result.type = TRT_NONNULL;
            result.value.nonnull.name = "EVT_HANDLE";
            result.value.nonnull.p = EvtOpenEventMetadataEnum(hevt, dw);
            break;
        case 108:
            result.type = TRT_NONNULL;
            result.value.nonnull.name = "EVT_HANDLE";
            result.value.nonnull.p = EvtNextEventMetadata(hevt, dw);
            if (result.value.nonnull.p == NULL &&
                GetLastError() == ERROR_NO_MORE_ITEMS)
                return TCL_OK;
            break;
        case 110:
            result.type =
                EvtGetObjectArraySize(hevt, &result.value.ival) ?
                TRT_DWORD : TRT_GETLASTERROR;
            break;
        }            
    }

    dw = TwapiSetResult(interp, &result);
    if (dw == TCL_OK)
        return TCL_OK;

    return Twapi_AppendEvtExtendedStatus(interp);
}

int Twapi_EvtInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct tcl_dispatch_s EvtTclDispatch[] = {
        DEFINE_TCL_CMD(GetEVT_VARIANT, Twapi_EvtGetEVT_VARIANTObjCmd),
        DEFINE_TCL_CMD(Twapi_EvtRenderValues, Twapi_EvtRenderValuesObjCmd),
        DEFINE_TCL_CMD(Twapi_EvtRenderUnicode, Twapi_EvtRenderUnicodeObjCmd),
        DEFINE_TCL_CMD(EvtNext, Twapi_EvtNextObjCmd),
        DEFINE_TCL_CMD(EvtCreateRenderContext, Twapi_EvtCreateRenderContextObjCmd),
        DEFINE_TCL_CMD(EvtFormatMessage, Twapi_EvtFormatMessageObjCmd),
        DEFINE_TCL_CMD(EvtOpenSession, Twapi_EvtOpenSessionObjCmd),
        DEFINE_TCL_CMD(Twapi_ExtractEVT_VARIANT_ARRAY, Twapi_ExtractEVT_VARIANT_ARRAYObjCmd),
        DEFINE_TCL_CMD(Twapi_ExtractEVT_RENDER_VALUES, Twapi_ExtractEVT_RENDER_VALUESObjCmd),
        DEFINE_TCL_CMD(EvtFormatMessage2, Twapi_EvtFormatMessage2ObjCmd),
    };

    static struct alias_dispatch_s EvtVariantGetDispatch[] = {
        DEFINE_ALIAS_CMD(EvtGetChannelConfigProperty, 2),
        DEFINE_ALIAS_CMD(EvtGetPublisherMetadataProperty, 3),
        DEFINE_ALIAS_CMD(EvtGetEventMetadataProperty, 4),
        DEFINE_ALIAS_CMD(EvtGetObjectArrayProperty, 5),
        DEFINE_ALIAS_CMD(EvtGetQueryInfo, 6),
        DEFINE_ALIAS_CMD(EvtGetEventInfo, 7),
        DEFINE_ALIAS_CMD(EvtGetLogInfo, 8),
    };

    static struct fncode_dispatch_s EvtFnDispatch[] = {
        DEFINE_FNCODE_CMD(EvtClearLog, 1),
        DEFINE_FNCODE_CMD(EvtQuery, 2),
        DEFINE_FNCODE_CMD(EvtSeek, 3),
        DEFINE_FNCODE_CMD(EvtOpenLog, 4),
        DEFINE_FNCODE_CMD(EvtOpenChannelConfig, 5),
        DEFINE_FNCODE_CMD(EvtArchiveExportedLog, 6),
        DEFINE_FNCODE_CMD(EvtSubscribe, 7),
        DEFINE_FNCODE_CMD(EvtExportLog, 8),
        DEFINE_FNCODE_CMD(EvtSetChannelConfigProperty, 9),
        DEFINE_FNCODE_CMD(EvtOpenPublisherMetadata, 10),
        DEFINE_FNCODE_CMD(evt_create_bookmark, 11),  // docs
        DEFINE_FNCODE_CMD(evt_update_bookmark, 12), // docs
        DEFINE_FNCODE_CMD(evt_free, 13), // docs
        DEFINE_FNCODE_CMD(evt_close, 101), // docs
        DEFINE_FNCODE_CMD(evt_cancel, 102), // docs
        DEFINE_FNCODE_CMD(EvtOpenChannelEnum, 103),
        DEFINE_FNCODE_CMD(EvtNextChannelPath, 104),
        DEFINE_FNCODE_CMD(EvtSaveChannelConfig, 105),
        DEFINE_FNCODE_CMD(EvtOpenPublisherEnum, 106),
        DEFINE_FNCODE_CMD(EvtOpenEventMetadataEnum, 107),
        DEFINE_FNCODE_CMD(EvtNextEventMetadata, 108),
        DEFINE_FNCODE_CMD(EvtNextPublisherId, 109),
        DEFINE_FNCODE_CMD(EvtGetObjectArraySize, 110),
    };

#if !defined(RUNTIME_EVT_LOAD)
    /* We are delay loading a dynamically linked winevt.dll. On XP/2k3
       this does not exist and must not be invoked, else crash will
       result
    */
    {
        OSVERSIONINFO osver;
        osver.dwOSVersionInfoSize = sizeof(osver);
        if (! GetVersionEx(&osver))
            return TwapiReturnSystemError(interp);
        if (osver.dwMajorVersion < 6) {
            /* Do not initialize evt_* commands */
            /* Not TCL_ERROR since other eventlog commands will still work */
            return TCL_OK;
        }
    }
#endif

    TwapiDefineTclCmds(interp, ARRAYSIZE(EvtTclDispatch), EvtTclDispatch, ticP);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(EvtFnDispatch), EvtFnDispatch, Twapi_EvtCallObjCmd);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(EvtVariantGetDispatch), EvtVariantGetDispatch, "twapi::GetEVT_VARIANT");
    return TCL_OK;
}
