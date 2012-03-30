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
    BOOL (WINAPI *_EvtNextPublisherId)(HANDLE,DWORD,LPWSTR,PDWORD);
    HANDLE (WINAPI *_EvtOpenPublisherMetadata)(HANDLE,LPCWSTR,LPCWSTR,LCID,DWORD);
    BOOL (WINAPI *_EvtGetPublisherMetadataProperty)(HANDLE,int,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    HANDLE (WINAPI *_EvtOpenEventMetadataEnum)(HANDLE,DWORD);
    HANDLE (WINAPI *_EvtNextEventMetadata)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtGetEventMetadataProperty)(HANDLE,int,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtGetObjectArraySize)(HANDLE, PDWORD);
    BOOL (WINAPI *_EvtGetObjectArrayProperty)(HANDLE,DWORD,DWORD,DWORD,DWORD,EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtGetQueryInfo)(HANDLE,int,DWORD,EVT_VARIANT *,PDWORD);
    HANDLE (WINAPI *_EvtCreateBookmark)(LPCWSTR);
    BOOL (WINAPI *_EvtUpdateBookmark)(HANDLE,HANDLE);
    BOOL (WINAPI *_EvtGetEventInfo)(HANDLE,int,DWORD,EVT_VARIANT *, PDWORD);
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

typedef HANDLE EVT_HANDLE;

int gEvtStatus;                 /* 0 - init, 1 - available, -1 - unavailable  */
EVT_HANDLE gEvtDllHandle;

/* Used as a typedef for returning allocated memory to script level */
#define TWAPI_EVT_RENDER_VALUES_TYPESTR "EVT_RENDER_VALUES *"

#define TWAPI_EVT_HANDLE_TYPESTR "EVT_HANDLE"
#define ObjFromEVT_HANDLE(h_) ObjFromOpaque((h_), TWAPI_EVT_HANDLE_TYPESTR)

static int ObjToEVT_VARIANT_ARRAY(
    Tcl_Interp *interp,
    Tcl_Obj *objP,
    void **bufPP,   /* Where to store pointer to array */
    int *szP,                   /* Where to store the size in bytes of buffer containing array */
    int *usedP,                 /* Where to store # bytes used */
    int *countP                 /* Where to store number of elements of array */
    )
{
    Tcl_Obj **objs;
    int nobjs;
    int sz, used, count;
    EVT_VARIANT *evaP;    

    if (Tcl_ListObjGetElements(interp, objP, &nobjs, &objs) != TCL_OK)
        return TCL_ERROR;
    if (nobjs == 0) {
        evaP = NULL;
        sz = 0;
        used = 0;
        count = 0;
    } else {
        /* If not empty, must be a list of two elements - size and
           ptr (which may be 0 and NULL) */
        if (nobjs != 4)
            return TwapiReturnError(interp, TWAPI_INVALID_ARGS);

        if (Tcl_GetIntFromObj(interp, objs[3], &count) != TCL_OK ||
            Tcl_GetIntFromObj(interp, objs[2], &used) != TCL_OK ||
            Tcl_GetIntFromObj(interp, objs[1], &sz) != TCL_OK ||
            ObjToOpaque(interp, objs[0], &evaP, TWAPI_EVT_RENDER_VALUES_TYPESTR) != TCL_OK)
            return TCL_ERROR;
    }

    if (bufPP)
        *bufPP = evaP;
    if (szP)
        *szP = sz;
    if (usedP)
        *usedP = used;
    if (countP)
        *countP = count;

    return TCL_OK;
}

Tcl_Obj *ObjFromEVT_VARIANT(TwapiInterpContext *ticP, EVT_VARIANT *varP)
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
        if (varP->StringVal)
            objP = ObjFromUnicode(varP->StringVal);
        break;
    case EvtVarTypeAnsiString:
        if (varP->AnsiStringVal)
            objP = Tcl_NewStringObj(varP->AnsiStringVal, -1);
        break;
    case EvtVarTypeSByte:
        objP = Tcl_NewIntObj(varP->SByteVal);
        break;
    case EvtVarTypeByte:
        objP = Tcl_NewIntObj(varP->ByteVal);
        break;
    case EvtVarTypeInt16:
        objP = Tcl_NewIntObj(varP->Int16Val);
        break;
    case EvtVarTypeUInt16:
        objP = Tcl_NewIntObj(varP->UInt16Val);
        break;
    case EvtVarTypeInt32:
        objP = Tcl_NewIntObj(varP->Int32Val);
        break;
    case EvtVarTypeUInt32:
        objP = ObjFromDWORD(varP->UInt32Val);
        break;
    case EvtVarTypeInt64:
        objP = Tcl_NewWideIntObj(varP->Int64Val);
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
        objP = Tcl_NewBooleanObj(varP->BooleanVal != 0);
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
            for (i = 0; i < count; ++i) {
                objPP[i] = varP->StringArr[i] ?
                    ObjFromUnicode(varP->StringArr[i])
                    : Tcl_NewObj();
            }
            break;
        case EvtVarTypeAnsiString:
            for (i = 0; i < count; ++i) {
                objPP[i] = varP->AnsiStringArr[i] ?
                    Tcl_NewStringObj(varP->AnsiStringArr[i], -1)
                    : Tcl_NewObj();
            }
            break;
        case EvtVarTypeSByte:
            /* TBD - should this be a byte array ? */
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewIntObj(varP->SByteArr[i]);
            }
            break;
        case EvtVarTypeByte:
            /* TBD - should this be a byte array ? */
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewIntObj(varP->ByteArr[i]);
            }
            break;
        case EvtVarTypeInt16:
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewIntObj(varP->Int16Arr[i]);
            }
            break;
        case EvtVarTypeUInt16:
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewIntObj(varP->UInt16Arr[i]);
            }
            break;
        case EvtVarTypeInt32:
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewIntObj(varP->Int32Arr[i]);
            }
            break;
        case EvtVarTypeUInt32:
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromDWORD(varP->UInt32Arr[i]);
            }
            break;
        case EvtVarTypeInt64:
            for (i = 0; i < count; ++i) {
                objPP[i] = Tcl_NewWideIntObj(varP->Int64Arr[i]);
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
                objPP[i] = Tcl_NewBooleanObj(varP->BooleanArr[i] != 0);
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
            objPP[0] = Tcl_NewObj();
            for (i = 1; i < count; ++i) {
                objPP[i] = objPP[0];
            }
            break;
        }
        TWAPI_ASSERT(count > 0 && objPP);
        objP = Tcl_NewListObj(count, objPP);
        MemLifoPopFrame(&ticP->memlifo);
        break;
    }

    if (objP == NULL) 
        objP = Tcl_NewObj();

    retObjs[0] = Tcl_NewIntObj(varP->Type);
    retObjs[1] = objP;
    return Tcl_NewListObj(2, retObjs);
}

/* Should be called only once at init time */
void TwapiInitEvtStubs(Tcl_Interp *interp)
{
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
        if (((* (FARPROC*) gEvtStubs._ ## fn) = GetProcAddress(gEvtDllHandle, #fn)) == NULL) \
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

    /* Success */
    gEvtStatus = 1;
}

/* IMPORTANT:
 * If a valid buffer is passed in (as the 4th arg) caller must not
 * access it again irrespective of successful or error return unless
 * in the former case the same buffer is returned explicitly
 */
static TCL_RESULT Twapi_EvtRenderValues(TwapiInterpContext *ticP, int objc,
                                        Tcl_Obj *CONST objv[])
{
    HANDLE hevt, hevt2;
    DWORD flags, sz, count, used, status;
    void *bufP;
    Tcl_Obj *objs[4];
    Tcl_Interp *interp = ticP->interp;

    if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt),
                     GETHANDLE(hevt2), GETINT(flags), ARGSKIP, ARGEND) != TCL_OK)
        return TCL_ERROR;

    bufP = NULL;
    /* 4th arg is supposed to describe a previously returned buffer
       that we can not reuse. It may also be NULL
    */
    if (ObjToEVT_VARIANT_ARRAY(interp, objv[3], &bufP, &sz, NULL, NULL) != TCL_OK)
        return TCL_ERROR;

    /* Allocate buffer if we were not passed one */
    if (bufP == NULL) {
        /* TBD - instrument reallocation needs */
        sz = 4000;
        bufP = TwapiAlloc(sz);
    }

    /* We used to convert using ObjFromEVT_VARIANT but that does
       not work well with opaque values so we preserve as a
       binary blob to be passed around. Note we cannot use
       a Tcl_ByteArray either because the embedded pointers will
       be invalid when the byte array is copied around.
    */

    /* EvtRenderEventValues -> 0 */
    status = ERROR_SUCCESS;
    if (EvtRender(hevt, hevt2, 0, sz, bufP, &used, &count) == FALSE) {
        TwapiFree(bufP);
        status = GetLastError();
        if (status == ERROR_INSUFFICIENT_BUFFER) {
            status = ERROR_SUCCESS;
            sz = used;
            bufP = TwapiAlloc(sz);
            if (EvtRender(hevt, hevt2, 0, sz,
                          bufP, &used, &count) == FALSE) {
                status = GetLastError();
            }
        }
    }

    if (status != ERROR_SUCCESS) {
        TwapiFree(bufP);
        return Twapi_AppendSystemError(interp, status);
    }

    objs[0] = ObjFromOpaque(bufP, TWAPI_EVT_RENDER_VALUES_TYPESTR);
    objs[1] = Tcl_NewIntObj(sz);
    objs[2] = Tcl_NewIntObj(used);
    objs[3] = Tcl_NewIntObj(count);
    Tcl_SetObjResult(interp, Tcl_NewListObj(4, objs));
    return TCL_OK;
}

/* EvtRender for Unicode return types */
static TCL_RESULT Twapi_EvtRenderUnicode(TwapiInterpContext *ticP, int objc,
                                  Tcl_Obj *CONST objv[])
{
    HANDLE hevt, hevt2;
    DWORD flags, sz, count, status;
    void *bufP;
    Tcl_Obj *objP;
    Tcl_Interp *interp = ticP->interp;

    if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt),
                     GETHANDLE(hevt2), GETINT(flags),
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
    objP = Tcl_NewUnicodeObj(bufP, -1);
    MemLifoPopFrame(&ticP->memlifo);
    Tcl_SetObjResult(ticP->interp, objP);
    return TCL_OK;
}

static TCL_RESULT Twapi_EvtNext(TwapiInterpContext *ticP, int objc,
                                Tcl_Obj *CONST objv[])
{
    EVT_HANDLE hevt;
    EVT_HANDLE *hevtP;
    DWORD dw, dw2, dw3;
     Tcl_Obj **objPP;
    Tcl_Interp *interp = ticP->interp;

    if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt), GETINT(dw),
                     GETINT(dw2), GETINT(dw3)) != TCL_OK)
        return TCL_ERROR;

    if (dw > 1024)
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
    hevtP = MemLifoPushFrame(&ticP->memlifo, dw*sizeof(*hevtP), NULL);
    if (EvtNext(hevt, dw, hevtP, dw2, dw3, &dw2) != FALSE) {
        if (dw2) {
            objPP = MemLifoAlloc(&ticP->memlifo, dw2*sizeof(*objPP), NULL);
            for (dw = 0; dw < dw2; ++dw) {
                objPP[dw] = ObjFromHANDLE(hevtP[dw]);
            }
            Tcl_SetObjResult(interp, Tcl_NewListObj(dw2, objPP));
        }
        dw = TCL_OK;
    } else {
        dw = GetLastError();
        if (dw == ERROR_NO_MORE_ITEMS)
            dw = TCL_OK;
        else {
            Twapi_AppendSystemError(interp, dw);
            dw = TCL_ERROR;
        }
    }
    MemLifoPopFrame(&ticP->memlifo);
    return dw;
}

static TCL_RESULT Twapi_EvtCreateRenderContext(TwapiInterpContext *ticP, int objc,
                                Tcl_Obj *CONST objv[])
{
    EVT_HANDLE hevt;
    int count;
    LPCWSTR xpaths[10];
    LPCWSTR *xpathsP;
    int flags;
    Tcl_Interp *interp = ticP->interp;
    DWORD ret = TCL_ERROR;

    if (TwapiGetArgs(interp, objc, objv, ARGSKIP, GETINT(flags)) != TCL_OK ||
        Tcl_ListObjLength(interp, objv[0], &count) != TCL_OK)
        return TCL_ERROR;

    if (count == 0) {
        xpathsP = NULL;
    } else {
        if (count <= ARRAYSIZE(xpaths)) {
            xpathsP = xpaths;
        } else {
            xpathsP = MemLifoPushFrame(&ticP->memlifo, count * sizeof(xpathsP[0]), NULL);
        }
        
        if (ObjToArgvW(interp, objv[0], xpathsP, count, &count) != TCL_OK)
            goto vamoose;
    }
    
    hevt = EvtCreateRenderContext(count, xpathsP, flags);
    if (hevt == NULL) {
        TwapiReturnSystemError(interp);
        goto vamoose;
    }
        
    Tcl_SetObjResult(interp, ObjFromEVT_HANDLE(hevt));
    ret = TCL_OK;

vamoose:
    if (xpathsP && xpathsP != xpaths)
        MemLifoPopFrame(&ticP->memlifo);
    return ret;
}

static TCL_RESULT Twapi_EvtFormatMessage(TwapiInterpContext *ticP, int objc,
                                Tcl_Obj *CONST objv[])
{
    Tcl_Interp *interp = ticP->interp;
    EVT_HANDLE hpub, hev;
    DWORD msgid, flags;
    EVT_VARIANT *valuesP;
    int nvalues;
    WCHAR buf[500];
    int used;
    WCHAR *bufP;
    DWORD winerr;

    if (TwapiGetArgs(interp, objc, objv,
                     GETHANDLE(hpub), GETHANDLE(hev),
                     GETINT(msgid), ARGSKIP, GETINT(flags), ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (ObjToEVT_VARIANT_ARRAY(interp, objv[3], &valuesP, NULL, NULL, &nvalues) != TCL_OK)
        return TCL_ERROR;

    /* TBD - instrument buffer size */
    bufP = buf;
    winerr = ERROR_SUCCESS;
    /* Note buffer sizes are in WCHARs, not bytes */
    if (EvtFormatMessage(hpub, hev, msgid, nvalues, valuesP, flags, ARRAYSIZE(buf), bufP, &used) == FALSE) {
        winerr = GetLastError();
        if (winerr == ERROR_INSUFFICIENT_BUFFER) {
            used *= sizeof(WCHAR);
            bufP = MemLifoPushFrame(&ticP->memlifo, used, NULL);
            if (EvtFormatMessage(hpub, hev, msgid, nvalues, valuesP, flags, used, bufP, &used) == FALSE) {
                winerr = GetLastError();
            } else {
                winerr = ERROR_SUCCESS;
            }
        }
    }        

    if (winerr == ERROR_SUCCESS) {
        /* TBD - see comments in GetMessageString function at
           http://msdn.microsoft.com/en-us/windows/dd996923%28v=vs.85%29
           - the buffer may contain multiple concatenated null terminated
           strings. For now pass the whole chunk as is. Later may
           split up either in here or in the script */
        Tcl_SetObjResult(interp, Tcl_NewUnicodeObj(bufP, used));
    } else {
        Twapi_AppendSystemError(interp, winerr);
    }
    if (bufP != buf)
        MemLifoPopFrame(&ticP->memlifo);

    return winerr == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
}

static TCL_RESULT Twapi_EvtGetLogInfo(TwapiInterpContext *ticP, int objc,
                                Tcl_Obj *CONST objv[])
{
    Tcl_Interp *interp = ticP->interp;
    EVT_HANDLE hevt;
    int propid, sz;
    EVT_VARIANT var;

    if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt), GETINT(propid),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (EvtGetLogInfo(hevt, propid, sizeof(var), &var, &sz) == FALSE)
        return TwapiReturnSystemError(interp);

    Tcl_SetObjResult(interp, ObjFromEVT_VARIANT(ticP, &var));
    return TCL_OK;
}

static TCL_RESULT Twapi_EvtOpenSession(TwapiInterpContext *ticP, int objc,
                                       Tcl_Obj *CONST objv[])
{
    Tcl_Interp *interp = ticP->interp;
    int login_class;
    DWORD timeout, flags;
    Tcl_Obj **loginObjs;
    int nobjs;
    EVT_RPC_LOGIN erl;

    if (TwapiGetArgs(interp, objc, objv, GETINT(login_class),
                     ARGSKIP, ARGUSEDEFAULT, GETINT(timeout), GETINT(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (login_class != 1) {
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid login class");
    }
    if (Tcl_ListObjGetElements(interp, objv[1], &nobjs, &loginObjs) != TCL_OK)
        return TCL_ERROR;

    if (nobjs != 5 || Tcl_GetIntFromObj(interp, loginObjs[4], &erl.Flags) != TCL_OK) {
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid EVT_RPC_LOGIN structure");
    }

    erl.Server = Tcl_GetUnicode(loginObjs[0]);
    erl.User = Tcl_GetUnicode(loginObjs[1]);
    erl.Domain = Tcl_GetUnicode(loginObjs[2]);
    erl.Password = Tcl_GetUnicode(loginObjs[3]);

    return TwapiReturnNonnullHandle(interp,
        EvtOpenSession(login_class, &erl, timeout, flags),
        TWAPI_EVT_HANDLE_TYPESTR);
}


int Twapi_EvtCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    DWORD dw, dw2;
    LPWSTR s, s2;
    HANDLE hevt, hevt2;
    Tcl_WideInt wide;
    
    if (gEvtStatus != 1)
        return Twapi_AppendSystemError(interp, ERROR_CALL_NOT_IMPLEMENTED);

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(func), ARGSKIP, ARGTERM) != TCL_OK)
        return TCL_ERROR;

    objc -= 2;
    objv += 2;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
    case 2:
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt), GETWSTR(s), GETNULLIFEMPTY(s2), GETINT(dw)) != TCL_OK)
            return TCL_ERROR;
        if (func == 1) {
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EvtClearLog(hevt, s, s2, dw);
        } else {
            result.type = TRT_NONNULL;
            result.value.nonnull.name = TWAPI_EVT_HANDLE_TYPESTR;
            result.value.nonnull.p = EvtQuery(hevt, s, s2, dw);
        }
        break;
    case 3: // EvtOpenSession
        return Twapi_EvtOpenSession(ticP, objc, objv);

    case 4: // EvtGetExtendedStatus
        break; // TBD - and does this even need a Tcl level access ?
    case 5: // EvtSubscribe - TBD
        break;
    case 6: // EvtNext
        return Twapi_EvtNext(ticP, objc, objv);

    case 7: // EvtSeek
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt),
                         GETWIDE(wide), GETHANDLE(hevt2), GETINT(dw),
                         GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (EvtSeek(hevt, wide, hevt2, dw, dw2) != FALSE)
            return TCL_OK;
        result.type = TRT_GETLASTERROR;
        break;
            
    case 8:
        return Twapi_EvtRenderUnicode(ticP, objc, objv);
    case 9:
        return Twapi_EvtRenderValues(ticP, objc, objv);

    case 10:
        return Twapi_EvtCreateRenderContext(ticP, objc, objv);

    case 11:
        return Twapi_EvtFormatMessage(ticP, objc, objv);

    case 12:
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt),
                         GETWSTR(s), GETINT(dw)) != TCL_OK)
            return TCL_ERROR;

        result.type = TRT_NONNULL;
        result.value.nonnull.name = TWAPI_EVT_HANDLE_TYPESTR;
        result.value.nonnull.p = EvtOpenLog(hevt, s, dw);
        break;
    case 13:
        return Twapi_EvtGetLogInfo(ticP, objc, objv);

    default:
        /* Params - HANDLE followed by optional DWORD */
        if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt),
                         ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 101:
            if (EvtClose(hevt) != FALSE)
                return TCL_OK;
            result.type = TRT_GETLASTERROR;
            break;
        case 102:
            if (EvtCancel(hevt) != FALSE)
                return TCL_OK;
            result.type = TRT_GETLASTERROR;
            break;
        }            
    }

    return TwapiSetResult(interp, &result);
}
