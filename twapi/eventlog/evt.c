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

typedef DWORD (WINAPI *TWAPI_EVT_SUBSCRIBE_CALLBACK)(
    int Action,
    PVOID UserContext,
    HANDLE Event );

typedef enum _TWAPI_EVT_VARIANT_TYPE
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

} TWAPI_EVT_VARIANT_TYPE;


#ifndef EVT_VARIANT_TYPE_MASK
#define EVT_VARIANT_TYPE_MASK 0x7f
#define EVT_VARIANT_TYPE_ARRAY 128
#endif

typedef struct _TWAPI_EVT_VARIANT
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

} TWAPI_EVT_VARIANT;


static struct {
    HANDLE (WINAPI *_EvtOpenSession)(int, PVOID, DWORD, DWORD);
    BOOL (WINAPI *_EvtClose)(HANDLE);
    BOOL (WINAPI *_EvtCancel)(HANDLE);
    DWORD (WINAPI *_EvtGetExtendedStatus)(DWORD, LPWSTR, PDWORD);
    HANDLE (WINAPI *_EvtQuery)(HANDLE, LPCWSTR, LPCWSTR, DWORD);
    BOOL (WINAPI *_EvtNext)(HANDLE, DWORD, HANDLE *, DWORD, DWORD, DWORD *);
    BOOL (WINAPI *_EvtSeek)(HANDLE, LONGLONG, HANDLE, DWORD, DWORD);
    HANDLE (WINAPI *_EvtSubscribe)(HANDLE,HANDLE,LPCWSTR,LPCWSTR,HANDLE,PVOID,TWAPI_EVT_SUBSCRIBE_CALLBACK, DWORD);
    HANDLE (WINAPI *_EvtCreateRenderContext)(DWORD, LPCWSTR*, DWORD);
    BOOL (WINAPI *_EvtRender)(HANDLE,HANDLE,DWORD,DWORD,PVOID, PDWORD, PDWORD);
    BOOL (WINAPI *_EvtFormatMessage)(HANDLE,HANDLE,DWORD,DWORD, TWAPI_EVT_VARIANT*,DWORD,DWORD, LPWSTR, PDWORD);
    HANDLE (WINAPI *_EvtOpenLog)(HANDLE,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtGetLogInfo)(HANDLE,int,DWORD,TWAPI_EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtClearLog)(HANDLE,LPCWSTR,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtExportLog)(HANDLE,LPCWSTR,LPCWSTR,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtArchiveExportedLog)(HANDLE,LPCWSTR,LCID,DWORD);
    HANDLE (WINAPI *_EvtOpenChannelEnum)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtNextChannelPath)(HANDLE,DWORD, LPWSTR, PDWORD);
    HANDLE (WINAPI *_EvtOpenChannelConfig)(HANDLE,LPCWSTR,DWORD);
    BOOL (WINAPI *_EvtSaveChannelConfig)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtSetChannelConfigProperty)(HANDLE, int, DWORD, TWAPI_EVT_VARIANT *);
    BOOL (WINAPI *_EvtGetChannelConfigProperty)(HANDLE,int,DWORD,DWORD,TWAPI_EVT_VARIANT *, PDWORD);
    HANDLE (WINAPI *_EvtOpenPublisherEnum)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtNextPublisherId)(HANDLE,DWORD,LPWSTR,PDWORD);
    HANDLE (WINAPI *_EvtOpenPublisherMetadata)(HANDLE,LPCWSTR,LPCWSTR,LCID,DWORD);
    BOOL (WINAPI *_EvtGetPublisherMetadataProperty)(HANDLE,int,DWORD,DWORD,TWAPI_EVT_VARIANT *, PDWORD);
    HANDLE (WINAPI *_EvtOpenEventMetadataEnum)(HANDLE,DWORD);
    HANDLE (WINAPI *_EvtNextEventMetadata)(HANDLE,DWORD);
    BOOL (WINAPI *_EvtGetEventMetadataProperty)(HANDLE,int,DWORD,DWORD,TWAPI_EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtGetObjectArraySize)(HANDLE, PDWORD);
    BOOL (WINAPI *_EvtGetObjectArrayProperty)(HANDLE,DWORD,DWORD,DWORD,DWORD,TWAPI_EVT_VARIANT *, PDWORD);
    BOOL (WINAPI *_EvtGetQueryInfo)(HANDLE,int,DWORD,TWAPI_EVT_VARIANT *,PDWORD);
    HANDLE (WINAPI *_EvtCreateBookmark)(LPCWSTR);
    BOOL (WINAPI *_EvtUpdateBookmark)(HANDLE,HANDLE);
    BOOL (WINAPI *_EvtGetEventInfo)(HANDLE,int,DWORD,TWAPI_EVT_VARIANT *, PDWORD);
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

int gEvtStatus;                 /* 0 - init, 1 - available, -1 - unavailable  */
HANDLE gEvtDllHandle;

Tcl_Obj *ObjFromEVT_VARIANT(TwapiInterpContext *ticP, TWAPI_EVT_VARIANT *varP)
{
    int i;
    Tcl_Obj *objP;
    Tcl_Obj *objPP;
    void *pv;

    if (varP->TYPE & EVT_VARIANT_TYPE_ARRAY) {
        count = varP->Count;
        pv = varP->ByteArr;     /* Does not really matter which field we pick */
        objPP = MemLifoPushFrame(ticP, count * sizeof(objPP[0]), NULL);
    } else {
        count = 1;
        pv = &varP->SByteVal;   /* Again no matte which union field we point to */
        objPP = &objP;
    }
        
#define PVELEM(p_, i_, t_) (*(i_ + (t_ *)p_))

    switch (varP->Type & EVT_VARIANT_TYPE_MASK) {
    case EvtVarTypeEvtXml:
    case EvtVarTypeString:
        for (i = 0; i < count; ++i) {
            objPP[i] = ObjFromUnicode(PVELEM(pv, i, LPCWSTR), -1);
        }
        break;
    case EvtVarTypeAnsiString:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewStringObj(PVELEM(pv, i, LPCSTR), -1);
        }
        break;
    case EvtVarTypeSByte:
        /* TBD - should this be a byte array ? */
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewIntObj(PVELEM(pv, i, signed char));
        }
        break;
    case EvtVarTypeByte:
        /* TBD - should this be a byte array ? */
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewIntObj(PVELEM(pv, i, unsigned char));
        }
        break;
    case EvtVarTypeInt16:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewIntObj(PVELEM(pv, i, signed short));
        }
        break;
    case EvtVarTypeUInt16:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewIntObj(PVELEM(pv, i, unsigned short));
        }
        break;
    case EvtVarTypeInt32:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewIntObj(PVELEM(pv, i, int));
        }
        break;
    case EvtVarTypeUInt32:
        for (i = 0; i < count; ++i) {
            objPP[i] = ObjFromDWORD(PVELEM(pv, i, unsigned int));
        }
        break;
    case EvtVarTypeInt64:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewWideIntObj(PVELEM(pv, i, Tcl_WideInt));
        }
        break;
    case EvtVarTypeUInt64:
        for (i = 0; i < count; ++i) {
            objPP[i] = ObjFromULONGLONG(PVELEM(pv, i, ULONGLONG));
        }
        break;
    case EvtVarTypeSingle:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewDoubleObj(PVELEM(pv,i,float));
        }
        break;
    case EvtVarTypeDouble:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewDoubleObj(PVELEM(pv,i,double));
        }
        break;
    case EvtVarTypeBoolean:
        for (i = 0; i < count; ++i) {
            objPP[i] = Tcl_NewBooleanObj(PVELEM(pv,i,BOOL));
        }
        break;
    case EvtVarTypeGuid:
        /* The way guid fields are defined, the standard way using
           PVELEM will not work for non-arrays. Do explicitly
        */
        if (varP->Type & EVT_VARIANT_TYPE_ARRAY) {
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromGUID(&varP->GuidArr[i]);
            }
        } else {
            objPP[0] = ObjFromGUID(varP->GuidVal);
        }
        break;
    case EvtVarTypeSizeT:
        for (i = 0; i < count; ++i) {
            objPP[i] = ObjFromSIZE_T(PVELEM(pv,i,size_t));
        }
        break;
    case EvtVarTypeFileTime:
        /* The way file time fields are defined, the standard way using
           PVELEM will not work for non-arrays. Do explicitly
        */
        if (varP->Type & EVT_VARIANT_TYPE_ARRAY) {
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromFILETIME(&varP->FileTimeArr[i]);
            }
        } else {
            objPP[0] = ObjFromULONGLONG(varP->FileTimeVal);
        }
        break;
    case EvtVarTypeSysTime:
        /* The way file time fields are defined, the standard way using
           PVELEM will not work for non-arrays. Do explicitly
        */
        if (varP->Type & EVT_VARIANT_TYPE_ARRAY) {
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromSYSTEMTIME(&varP->SysTimeArr[i]);
            }
        } else {
            objPP[0] = ObjFromSYSTEMTIME(varP->SysTimeVal);
        }
        break;
    case EvtVarTypeSid:
        for (i = 0; i < count; ++i) {
            objPP[i] = ObjFromSIDNoFail(PVELEM(pv,i,PSID));
        }
        break;
    case EvtVarTypeEvtHandle:
        for (i = 0; i < count; ++i) {
            objPP[i] = ObjFromOpaque(PVELEM(pv,i,void*), "EVT_HANDLE");
        }
        break;
    case EvtVarTypeHexInt32: /* TBD - what field to use ? */
    case EvtVarTypeHexInt64: /* TBD - what field to use ? */
    case EvtVarTypeBinary:
        /* No DOCS as to what this is. Type is PBYTE but length ? TBD */
        /* FALLTHRU */
    case EvtVarTypeNull:
    default:        
        /* Current contract always calls for valid Tcl_Obj to be returned
           so cannot raise error 
        */
        objPP[0] = Tcl_NewObj();
        for (i = 1; i < count; ++i) {
            objPP[i] = objPP[0];
        }
        break;
    }

    if (objPP != &objP) {
        /* Array of objects */
        objP = Tcl_NewListObj(count, objPP);
        MemLifoPopFrame(&ticP->memlifo);
    }
    return objP;
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

static TCL_RESULT Twapi_EvtRender(TwapiInterpContext *ticP, int objc,
                                  Tcl_Obj *CONST objv[])
{
    HANDLE hevt, hevt2;
    DWORD flags, sz, count, status;
    void *bufP;
    Tcl_Obj *objP;
    Tcl_Obj *objPP;

    if (TwapiGetArgs(interp, objc, objv, GETHANDLE(hevt),
                     GETHANDLE(hevt2), GETINT(flags), GETINT(sz),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    bufP = MemLifoPushFrame(&ticP->memlifo, sz, &sz);
    status = ERROR_SUCCESS;
    if (EvtRender(hevt, hevt2, flags, sz, bufP, &sz, &count) == FALSE) {
        status = GetLastError();
        if (status == ERROR_INSUFFICIENT_BUFFER) {
            /* TBD - instrument reallocation needs */
            bufP = Memlifoalloc(&ticP->memlifo, sz);
            if (EvtRender(hevt, hevt2, flags, sz, bufP, &sz, &count) == FALSE)
                status = GetLastError();
            else
                status = ERROR_SUCCESS;
        }
    }

    if (status != ERROR_SUCCESS) {
        Twapi_AppendSystemError(interp, status);
    } else {
        /* Result may be string or array of EVT_VARIANT */
        objP = NULL;
        switch (flags) {
        case 0: /* EvtRenderArrayValues - Array of variant structs */
            objPP = MemlifoAlloc(&ticP->memlifo, count * sizeof(objPP[0]));
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromEVT_VARIANT(i + (TWAPI_EVT_VARIANT *)bufP);
            }
            objP = Tcl_NewListObj(count, objPP);
            break;
        case 1: /* EvtRenderEventXml */
            /* FALLTHRU */
        case 2: /* EvtRenderBookmark */
            /* Unicode string. Should we use sz/2 instead of -1 ? TBD */
            objP = Tcl_NewUnicodeObj(bufP, -1);
            break;
        default: /* unknown */
            objP = Tcl_NewByteArrayObj(bufP, sz);
            break;
        }
        if (objP) {
            status = TCL_OK;
            Tcl_SetObjResult(ticP->interp, objP);
        }
    }
        
    MemLifoPopFrame(&ticP->memlifo);
    return status == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
}





int Twapi_EvtCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    DWORD dw, dw2, dw3;
    LPWSTR s, s2;
    HANDLE hevt, hevt2, hevt3;
    HANDLE *hevtP;
    Tcl_Obj *objPP;
    Tcl_WideInt wide;
    void *bufP;
    
    if (gEvtStatus != 1)
        return Twapi_AppendSystemError(interp, ERROR_CALL_NOT_IMPLEMENTED);

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(func), ARGSKIP, ARGTERM) != TCL_OK)
        return TCL_ERROR;

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 0:
    case 0:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(hevt), GETWSTR(s), GETWSTR(s2), GETINT(dw)) != TCL_OK)
            return TCL_ERROR;
        if (func == TBD) {
            result.type = TRT_EMPTY;
            if (! EvtClearLog(hevt, s, s2, dw))
                result.type = TRT_GETLASTERROR;
        } else {
            result.type = TRT_OPAQUE;
            if ((result.value.opaque.p = EvtQuery(hevt, s, s2, dw)) == NULL)
                result.type = TRT_GETLASTERROR;
        }
        break;
    case 0: // EvtOpenSession
        break;  /* TBD */
    case 0: // EvtGetExtendedStatus
        break; // TBD - and does this even need a Tcl level access ?
    case 0: // EvtSubscribe
        break;
    case 0: // EvtNext
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(hevt), GETINT(dw),
                         GETINT(dw2), GETINT(dw3)) != TCL_OK)
            return TCL_ERROR;
        if (dw > 1024)
            return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
        hevtP = MemLifoPushFrame(&ticP->memlifo, dw*sizeof(*hevtP), NULL);
        if (EvtNext(hevt, dw, hevtP, dw2, dw3, &dw2) != FALSE) {
            if (dw2) {
                objPP = MemLifoAlloc(&ticP->memlifo, dw2*sizeof(*objPP));
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

    case 0: // EvtSeek
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(hevt),
                         GETWIDE(wide), GETHANDLE(hevt2), GETINT(dw),
                         GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (EvtSeek(hevt, wide, hevt2, dw, dw2) != FALSE)
            return TCL_OK;
        result.type = TRT_GETLASTERROR;
        break;
            
    case 0: // EvtRender
        return Twapi_EvtRender(ticP, objc-2, objv+2);

    default:
        /* Params - HANDLE followed by optional DWORD */
        if (TwapiGetArgs(interp, objc-2, objv+2, GETHANDLE(hevt),
                         ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 0:
            if (EvtClose(hevt) != FALSE)
                return TCL_OK;
            result.type = TRT_GETLASTERROR;
            break;
        case 0:
            if (EvtCancel(hevt) != FALSE)
                return TCL_OK;
            result.type = TRT_GETLASTERROR;
            break;
        }            
    }

    return TwapiSetResult(interp, &result);
}
