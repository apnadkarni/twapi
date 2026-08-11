/*
 * Copyright (c) 2012-2024, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/*
 * Vista+ eventlog support (Evt* functions). These are all loaded dynamically
 * as they are not available on XP/2k3
 */

#include "twapi.h"
#include "twapi_events.h"

#include <ntverp.h>             /* Needed for VER_PRODUCTBUILD SDK version */

# include <winevt.h>
#ifdef _MSC_VER
#pragma comment(lib, "delayimp.lib") /* Prevents winevt from loading unless necessary */
#endif

#ifndef TWAPI_SINGLE_MODULE
HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

#ifndef MODULENAME
#define MODULENAME "twapi_evt"
#endif

static REGHANDLE gEvtRegHandle = 0;
static TwapiOneTimeInitState gEvtInitialized;

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

    
    bufP = SWSPushFrame(sizeof(WCHAR) * sz, NULL);
    if (EvtGetExtendedStatus(sz, bufP, &used) != FALSE && used != 0) {
        /* TBD - verify this works (is bufP null terminated ?) */
        Tcl_AppendResult(interp, " ", bufP, NULL);
    }
    SWSPopFrame();

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
            objP = ObjFromWinChars(varP->StringVal);
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

        objPP = MemLifoPushFrame(ticP->memlifoP,
                                 count * sizeof(objPP[0]), NULL);

        switch (varP->Type & EVT_VARIANT_TYPE_MASK) {
        case EvtVarTypeString:
        case EvtVarTypeEvtXml:
            for (i = 0; i < count; ++i) {
                objPP[i] = varP->StringArr[i] ?
                    ObjFromWinChars(varP->StringArr[i])
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
        MemLifoPopFrame(ticP->memlifoP);
        break;
    }

    if (objP == NULL) {
        // TBD - debug output
        objP = ObjFromEmptyString();
    }

    if (flags & 1) {
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

    objPP = MemLifoPushFrame(ticP->memlifoP, count * sizeof (objPP[0]), NULL);

    for (i = 0; i < count; ++i) {
        objPP[i] = ObjFromEVT_VARIANT(ticP, &varP[i], 0);
    }

    objP = ObjNewList(count, objPP);
    MemLifoPopFrame(ticP->memlifoP);

    return objP;
}

/* IMPORTANT:
 * If a valid buffer is passed in (as the 4th arg) caller must not
 * access it again irrespective of successful or error return unless
 * in the former case the same buffer is returned explicitly
 */
static TCL_RESULT Twapi_EvtRenderValuesObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE hevt, hevt2;
    DWORD status;
    void *bufP;
    TwapiEVT_RENDER_VALUES_HEADER *ervhP;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETEVTH(hevt),
                     GETHANDLET(hevt2, EVT_HANDLE),
                     GETVERIFIEDORNULL(ervhP, TwapiEVT_RENDER_VALUES_HEADER*, Twapi_EvtRenderValuesObjCmd),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* 4th arg is supposed to describe a previously returned buffer
       that we can reuse. It may also be NULL
    */
    if (ervhP == NULL) {
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


    ObjSetResult(interp, ObjFromOpaque(ervhP, "TwapiEVT_RENDER_VALUES_HEADER*"));
    return TCL_OK;
}

/* EvtRender for Unicode return types */
static TCL_RESULT Twapi_EvtRenderUnicodeObjCmd(ClientData clientdata, Tcl_Interp *interp,  int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    HANDLE hevt, hevt2;
    DWORD flags, sz, count, status;
    void *bufP;
    Tcl_Obj *objP;
    MemLifoSize len;
    if (TwapiGetArgs(interp, objc-1, objv+1, GETEVTH(hevt),
                     GETHANDLET(hevt2, EVT_HANDLE), GETDWORD(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* 1 -> EvtRenderEventXml */
    /* 2 -> EvtRenderBookmark */
    if (flags != 1 && flags != 2)
        return TwapiReturnError(interp,TWAPI_INVALID_ARGS);

    /* TBD - instrument reallocation needs */
    sz = 256;
    bufP = MemLifoPushFrame(ticP->memlifoP, sz, &len);
    sz = len > ULONG_MAX ? ULONG_MAX : (ULONG) len;
    status = ERROR_SUCCESS;
    if (EvtRender(hevt, hevt2, flags, sz, bufP, &sz, &count) == FALSE) {
        status = GetLastError();
        if (status == ERROR_INSUFFICIENT_BUFFER) {
            /* Note no need to MemlifoPopFrame before allocating more */
            bufP = MemLifoAlloc(ticP->memlifoP, sz, &len);
            sz = len > ULONG_MAX ? ULONG_MAX : (ULONG) len;
            if (EvtRender(hevt, hevt2, flags, sz, bufP, &sz, &count) == FALSE)
                status = GetLastError();
            else
                status = ERROR_SUCCESS;
        }
    }

    if (status != ERROR_SUCCESS) {
        MemLifoPopFrame(ticP->memlifoP);
        return Twapi_AppendSystemError(interp, status);
    }

    /* Unicode string. Should we use sz/2 instead of -1 ? TBD */
    objP = ObjFromWinChars(bufP);
    MemLifoPopFrame(ticP->memlifoP);
    ObjSetResult(ticP->interp, objP);
    return TCL_OK;
}

static TCL_RESULT Twapi_ExtractEVT_VARIANT_ARRAYObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    EVT_VARIANT *varP;
    int count;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETHANDLE(varP), GETINT(count),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
        
    ObjSetResult(interp, ObjFromEVT_VARIANT_ARRAY(ticP, varP, count));
    return TCL_OK;
}

static TCL_RESULT Twapi_ExtractEVT_RENDER_VALUESObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    TwapiEVT_RENDER_VALUES_HEADER *ervhP;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVERIFIEDPTR(ervhP, TwapiEVT_RENDER_VALUES_HEADER*, Twapi_EvtRenderValuesObjCmd),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    ObjSetResult(interp, ObjFromEVT_VARIANT_ARRAY(ticP, ERVHP_BUFFER(ervhP), ervhP->header.count));
    return TCL_OK;
}

static TCL_RESULT Twapi_EvtNextObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    EVT_HANDLE hevt;
    EVT_HANDLE *hevtP;
    DWORD i, count, timeout, dw;
    Tcl_Obj *objP;
    Tcl_Obj **objPP;
    int result;

    /* ARGSKIP is status. Only filled on error */
    if (TwapiGetArgs(interp, objc-1, objv+1, GETEVTH(hevt), GETDWORD(count),
                     GETDWORD(timeout), GETDWORD(dw), ARGUSEDEFAULT, ARGSKIP,
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (count > 1024) // TBD - why ?
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
    hevtP = MemLifoPushFrame(ticP->memlifoP, count*sizeof(*hevtP), NULL);
    if (EvtNext(hevt, count, hevtP, timeout, dw, &count) != FALSE) {
        if (count) {
            objPP = MemLifoAlloc(ticP->memlifoP, count*sizeof(*objPP), NULL);
            for (i = 0; i < count; ++i) {
                objPP[i] = ObjFromEVT_HANDLE(hevtP[i]);
            }
            ObjSetResult(interp, ObjNewList(count, objPP));
        }
        result = TCL_OK;
        dw = 0;
    } else {
        count = 0;
        result = TCL_ERROR;
        dw = GetLastError();
    }


    if (count == 0) {
        /* If empty list being returned, indicate status */
        if (objc == 6) {
            /*
             * Caller has supplied a variable to hold status. In this case
             * we do not generate an error even if there is one. Set
             * the variable to the status. The command return will stay empty.
             */
            objP = ObjFromDWORD(dw);
            if (Tcl_ObjSetVar2(interp, objv[5], NULL, objP, TCL_LEAVE_ERR_MSG) == NULL) {
                Twapi_FreeNewTclObj(objP);
                result = TCL_ERROR; /* Invalid variable */
            } else
                result = TCL_OK;
        } else {
            /* No status var supplied. In this case, eof is not an error */
            if (dw == 0 || dw == ERROR_NO_MORE_ITEMS || dw == ERROR_TIMEOUT)
                result = TCL_OK;
            else {
                result = Twapi_AppendSystemError(interp, dw);
            }
        }
    }

    MemLifoPopFrame(ticP->memlifoP);
    return result;
}

static TCL_RESULT Twapi_EvtCreateRenderContextObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    EVT_HANDLE hevt;
    Tcl_Size count;
    LPCWSTR *xpathsP = NULL;
    int flags;
    const char *msg = NULL;
    int ret;

    ret = TwapiGetArgs(
        interp, objc - 1, objv + 1, ARGSKIP, GETINT(flags), ARGEND);
    if (ret != TCL_OK ||
        ObjListLength(interp, objv[1], &count) != TCL_OK)
        return TCL_ERROR;

    if (flags == EvtRenderContextValues) {
        if (count == 0) {
            msg = "At least one XPATH expression must be specified.";
            ret = TCL_ERROR;
        }
        else {
            /* Note ObjToArgvW needs an extra entry for terminating NULL */
            xpathsP = MemLifoPushFrame(
                ticP->memlifoP, (count + 1) * sizeof(xpathsP[0]), NULL);
            ret = ObjToArgvW(interp, objv[1], xpathsP, count + 1, &count);
            if (ret == TCL_OK) {
                ret = DWORD_LIMIT_CHECK(interp, count);
            }
        }
    } else {
        if (count != 0) {
            msg = "No XPATH expressions must be specified.";
            ret = TCL_ERROR;
        }
    }
    if (ret == TCL_OK) {
        hevt = EvtCreateRenderContext((DWORD)count, xpathsP, flags);
        if (hevt == NULL) {
            ret = TwapiReturnSystemError(interp);
        } else {
            ObjSetResult(interp, ObjFromEVT_HANDLE(hevt));
            ret = TCL_OK;
        }
    }
    else {
	/* If msg is NULL, interp result is already set. */
        if (msg != NULL) {
            TwapiReturnErrorEx(
                interp, TWAPI_INVALID_DATA, Tcl_NewStringObj(msg, -1));
        }
    }
    if (xpathsP)
        MemLifoPopFrame(ticP->memlifoP);
    return ret;
}


static TCL_RESULT Twapi_EvtFormatMessageObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    EVT_HANDLE hpub, hev;
    DWORD msgid, flags, used, buf_sz;
    EVT_VARIANT *valuesP;
    int nvalues;
    WCHAR buf[500];             /* TBD - instrument */
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
                     GETDWORD(msgid),
                     GETVERIFIEDORNULL(ervhP, TwapiEVT_RENDER_VALUES_HEADER*, Twapi_EvtRenderValuesObjCmd),
                     GETDWORD(flags),
                     ARGUSEDEFAULT, ARGSKIP, ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (ervhP) {
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
            bufP = MemLifoPushFrame(ticP->memlifoP, sizeof(WCHAR)*buf_sz, NULL);
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
            objP = ObjFromWinChars(bufP);
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
                ObjSetResult(interp, ObjFromInt(winerr == ERROR_SUCCESS));
            }
        } else
            ObjSetResult(interp, objP);
    }

    if (bufP != buf)
        MemLifoPopFrame(ticP->memlifoP);

    return status;
}


static TCL_RESULT Twapi_EvtGetEVT_VARIANTObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    EVT_HANDLE hevt;
    EVT_VARIANT *varP;
    int func;
    DWORD sz, dw, dw2, dw3;
    DWORD status;
    MemLifoSize len;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(func), GETEVTH(hevt),
                     GETDWORD(dw), ARGUSEDEFAULT, GETDWORD(dw2),
                     GETDWORD(dw3), ARGEND) != TCL_OK)
        return TCL_ERROR;

    varP = MemLifoPushFrame(ticP->memlifoP, sizeof(EVT_VARIANT), &len);
    sz = len > ULONG_MAX ? ULONG_MAX : (ULONG) len;
    while (1) {
        switch (func) {
        case 2:
            status = EvtGetChannelConfigProperty(hevt, dw, dw2, sz, varP, &sz);
            break;
        case 3:
            status = EvtGetPublisherMetadataProperty(hevt, dw, dw2, sz, varP, &sz);
            break;
        case 4:
            status = EvtGetEventMetadataProperty(hevt, dw, dw2, sz, varP, &sz);
            break;
        case 5:
            status = EvtGetObjectArrayProperty(hevt, dw, dw2, dw3, sz, varP, &sz);
            break;
        case 6:
            status = EvtGetQueryInfo(hevt, dw, sz, varP, &sz);
            break;
        case 7:
            status = EvtGetEventInfo(hevt, dw, sz, varP, &sz);
            break;
        case 8:
            status = EvtGetLogInfo(hevt, dw, sz, varP, &sz);
            break;
        default:
            MemLifoPopFrame(ticP->memlifoP);
            return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);
        }
        if (status != FALSE || GetLastError() != ERROR_INSUFFICIENT_BUFFER)
            break;
        /* Loop to retry larger buffer. No need to free previous alloc first */
        varP = MemLifoAlloc(ticP->memlifoP, sz, NULL);
    }

    if (status == FALSE) {
        TwapiReturnSystemError(interp);
        Twapi_AppendEvtExtendedStatus(interp);
    } else {
        ObjSetResult(interp, ObjFromEVT_VARIANT(ticP, varP, 0));
    }
    MemLifoPopFrame(ticP->memlifoP);

    return status == FALSE ? TCL_ERROR : TCL_OK;
}


static TCL_RESULT Twapi_EvtOpenSessionObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    int login_class;
    DWORD timeout, flags;
    Tcl_Obj **loginObjs;
    Tcl_Size nobjs;
    EVT_RPC_LOGIN erl;
    TCL_RESULT res;
    WCHAR *passwordP;
    Tcl_Size password_len;
    MemLifoMarkHandle mark = NULL;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(login_class),
                     ARGSKIP, ARGUSEDEFAULT, GETDWORD(timeout), GETDWORD(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    if (login_class != 1) {
        /* Only EvtRpcLogin (1) supported */
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid login class");
    }
    if (ObjGetElements(interp, objv[2], &nobjs, &loginObjs) != TCL_OK)
        return TCL_ERROR;

    if (nobjs != 5 ||
        ObjToDWORD(interp, loginObjs[4], &erl.Flags) != TCL_OK) {
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid EVT_RPC_LOGIN structure");

    }
    erl.Server = ObjToWinChars(loginObjs[0]);
    erl.User = ObjToLPWSTR_NULL_IF_EMPTY(loginObjs[1]);
    erl.Domain = ObjToLPWSTR_NULL_IF_EMPTY(loginObjs[2]);
    mark = MemLifoPushMark(ticP->memlifoP);
    passwordP = ObjDecryptPasswordSWS(loginObjs[3], &password_len);
    erl.Password = passwordP;
    NULLIFY_EMPTY(erl.Password);
    res = TwapiReturnNonnullHandle(interp,
                                   EvtOpenSession(login_class, &erl,
                                                  timeout, flags),
                                   "EVT_HANDLE");
    SecureZeroMemory(passwordP, sizeof(WCHAR) * password_len);
    MemLifoPopMark(mark);
    return res;
}

int
Twapi_EvtCallObjCmd(ClientData clientdata,
                    Tcl_Interp *interp,
                    int objc,
                    Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    DWORD dw, dw2;
    Tcl_Obj *sObj, *s2Obj, *s3Obj;
    LPWSTR s, s2;
    EVT_HANDLE hevt, hevt2, *hevtP;
    Tcl_WideInt wide;
    int func = PtrToInt(clientdata);
    HANDLE h;
    union {
        WCHAR buf[MAX_PATH+1];
        EVT_HANDLE hevts[100];
    } u;
    EVT_VARIANT var;
    int i;

    --objc;
    ++objv;

    /* NOTE AS ALWAYS, TO AVOID SHIMMERING ISSUES, WSTR ARGS ARE ALWAYS
       EXTRACTED AFTER SCALAR ARGS */

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
    case 2:
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt), GETOBJ(sObj), GETOBJ(s2Obj), GETDWORD(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        s = ObjToWinChars(sObj);
        s2 = ObjToLPWSTR_NULL_IF_EMPTY(s2Obj);
        if (func == 1) {
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EvtClearLog(hevt, s, s2, dw);
        } else {
            TwapiResult_SET_NONNULL_PTR(
                result, EVT_HANDLE, EvtQuery(hevt, s, s2, dw));
        }
        break;
    case 3: // EvtSeek
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETWIDE(wide), GETHANDLET(hevt2, EVT_HANDLE),
                         GETDWORD(dw), GETDWORD(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtSeek(hevt, wide, hevt2, dw, dw2);
        break;
    case 4: // EvtOpenLog
    case 5: // EvtOpenChannelConfig
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETOBJ(sObj), GETDWORD(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        hevt2 = (func == 4 ? EvtOpenLog : EvtOpenChannelConfig) (hevt, ObjToWinChars(sObj), dw);
        TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, hevt2);
        break;
    case 6: // EvtArchiveExportedLog
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETOBJ(sObj), GETDWORD(dw), GETDWORD(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtArchiveExportedLog(hevt, ObjToWinChars(sObj), dw, dw2);
        break;
    case 7: // EvtSubscribe
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETHANDLE(h), GETOBJ(sObj), GETOBJ(s2Obj),
                         GETEVTH(hevt2), GETDWORD(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, EvtSubscribe(hevt, h, ObjToWinChars(sObj), ObjToWinChars(s2Obj), hevt2, NULL, NULL, dw));
        break;
    case 8: // EvtExportLog
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETOBJ(sObj), GETOBJ(s2Obj), GETOBJ(s3Obj), GETDWORD(dw),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtExportLog(hevt, ObjToWinChars(sObj), ObjToWinChars(s2Obj), ObjToWinChars(s3Obj), dw);
        break;

    case 9: // EvtSetChannelConfigProperty
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETDWORD(dw), GETDWORD(dw2), ARGSKIP,
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        var.Type = EvtVarTypeNull;
        var.Count = 0;
        /* dw is property id */
        switch (dw) {
        case EvtChannelConfigEnabled:
        case EvtChannelLoggingConfigRetention:
        case EvtChannelLoggingConfigAutoBackup:
            if (ObjToBoolean(interp, objv[3], &var.BooleanVal) != TCL_OK)
                return TCL_ERROR;
            var.Type = EvtVarTypeBoolean;
            break;
        case EvtChannelConfigIsolation:
        case EvtChannelPublishingConfigLevel:
        case EvtChannelPublishingConfigFileMax:
            if (ObjToDWORD(interp, objv[3], (DWORD *)&var.UInt32Val) != TCL_OK)
                return TCL_ERROR;
            var.Type = EvtVarTypeUInt32;
            break;
        case EvtChannelConfigAccess:
        case EvtChannelLoggingConfigLogFilePath:
            var.StringVal = ObjToWinChars(objv[3]);
            var.Type = EvtVarTypeString;
            break;

        case EvtChannelLoggingConfigMaxSize:
        case EvtChannelPublishingConfigKeywords:
            TWAPI_ASSERT(sizeof(Tcl_WideInt) == sizeof(UINT64));
            if (ObjToWideInt(interp, objv[3], (Tcl_WideInt *)&var.UInt64Val) != TCL_OK)
                return TCL_ERROR;
            var.Type = EvtVarTypeUInt64;
            break;
        default:
            /* Note following properties cannot be set
               case EvtChannelConfigType:
               case EvtChannelPublishingConfigBufferSize:
               case EvtChannelPublishingConfigMinBuffers:
               case EvtChannelPublishingConfigMaxBuffers:
               case EvtChannelPublishingConfigLatency:
               case EvtChannelPublishingConfigSidType:
               case EvtChannelConfigOwningPublisher:
               case EvtChannelConfigClassicEventlog:
               case EvtChannelPublishingConfigControlGuid:
               case EvtChannelPublishingConfigClockType:
               case EvtChannelPublisherList:
            */
            return TwapiReturnErrorEx(
                interp,
                TWAPI_INVALID_ARGS,
                Tcl_NewStringObj("Attempt to modify a read-only channel property.",
                                 -1));
        }
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = EvtSetChannelConfigProperty(hevt, dw, dw2, &var);
        break;

    case 10: // EvtOpenPublisherMetadata
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         GETOBJ(sObj), GETOBJ(s2Obj), GETDWORD(dw),
                         GETDWORD(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        s2 = ObjToWinChars(s2Obj);
        NULLIFY_EMPTY(s2);
        TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, EvtOpenPublisherMetadata(hevt, ObjToWinChars(sObj), s2, dw, dw2));
        break;

    case 11: // evt_create_bookmark
        if (TwapiGetArgs(interp, objc, objv, ARGUSEDEFAULT,
                         GETOBJ(sObj), ARGEND) != TCL_OK)
            return TCL_ERROR;
        s = ObjToLPWSTR_NULL_IF_EMPTY(sObj);
        TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, EvtCreateBookmark(s));
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

    case 14:
        /* First verify syntax of handles */
        if (objc >= ARRAYSIZE(u.hevts))
            hevtP = SWSPushFrame(objc * sizeof(*hevtP), NULL);
        else
            hevtP = u.hevts;
        for (i = 0; i < objc; ++i) {
            if (ObjToOpaque(interp, objv[i], &hevtP[i], "EVT_HANDLE") != TCL_OK) {
                if (objc >= ARRAYSIZE(u.hevts))
                    SWSPopFrame();
                return TCL_ERROR;
            }
        }
        /* Now do the actual close. On errors, we continue to close remaining */
        result.type = TRT_EMPTY;
        for (i = 0; i < objc; ++i) {
            if (! EvtClose(hevtP[i])) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = GetLastError();
            }
        }
        if (objc >= ARRAYSIZE(u.hevts))
            SWSPopFrame();
        break;

    default:
        /* Params - HANDLE followed by optional DWORD */
        if (TwapiGetArgs(interp, objc, objv, GETEVTH(hevt),
                         ARGUSEDEFAULT, GETDWORD(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 102:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = EvtCancel(hevt);
            break;
        case 103:
            TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, EvtOpenChannelEnum(hevt, dw));
            break;
        case 104: // EvtNextChannelPath
        case 109: // EvtNextPublisherId
            /* Note channel/publisher is max 255 chars so no need to check
               for ERROR_INSUFFICIENT_BUFFER */
            if ((func == 104 ? EvtNextChannelPath : EvtNextPublisherId)(hevt, ARRAYSIZE(u.buf), u.buf, &dw) != FALSE) {
                ObjSetResult(interp, ObjFromWinCharsN(u.buf, dw-1));
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
            TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, EvtOpenPublisherEnum(hevt, dw));
            break;
        case 107:
            TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, EvtOpenEventMetadataEnum(hevt, dw));
            break;
        case 108:
            TwapiResult_SET_NONNULL_PTR(result, EVT_HANDLE, EvtNextEventMetadata(hevt, dw));
            if (result.value.ptr.p == NULL &&
                GetLastError() == ERROR_NO_MORE_ITEMS)
                return TCL_OK;
            break;
        case 110:
            result.type =
                EvtGetObjectArraySize(hevt, &result.value.uval) ?
                TRT_DWORD : TRT_GETLASTERROR;
            break;
        }
    }

    dw = TwapiSetResult(interp, &result);
    if (dw == TCL_OK)
        return TCL_OK;

    return Twapi_AppendEvtExtendedStatus(interp);
}

static int
TwapiParseSeverity(Tcl_Interp *interp, Tcl_Obj *obj, BYTE *levelPtr)
{
    static const char *const names[] = {
        "critical", "error", "warning", "informational", "verbose", NULL
    };
    int level;

    if (Tcl_GetIntFromObj(NULL, obj, &level) == TCL_OK) {
	/* Our manifest only defines events for levels 1-5 */
        if (level < 1 || level > 5) {
            return TwapiReturnErrorMsg(
                interp,
                TWAPI_INVALID_DATA,
                "Numeric severity levels must be in range 1-5");
        }
        *levelPtr = (BYTE)level;
        return TCL_OK;
    }

    if (Tcl_GetIndexFromObj(interp, obj, names, "severity level", 0, &level) != TCL_OK) {
        return TCL_ERROR;
    }
    *levelPtr = (BYTE)(level + 1);
    return TCL_OK;
}

int
Twapi_EvtLogObjCmd(ClientData clientData,
                   Tcl_Interp *interp,
                   int objc,
                   Tcl_Obj *const objv[])
{
    BYTE        level;
    const char *opts[] = {
        "-application", "-activityid", "-relatedactivityid", NULL
    };
    enum opt { OPT_APP, OPT_ACTIVITY, OPT_RELATED_ACTIVITY };
    int index;
    Tcl_Obj *appObj = NULL;
    GUID activity_guid, *activity_guidP = NULL;
    GUID related_activity_guid, *related_activity_guidP = NULL;

    if (objc < 3) {
        Tcl_WrongNumArgs(interp, 1, objv, "level message");
        return TCL_ERROR;
    }

    for (int i = 3; i < objc; ++i) {
        if (Tcl_GetIndexFromObj(interp, objv[i], opts, "option", 0, &index)
            != TCL_OK) {
            return TCL_ERROR;
        }
        if (++i == objc) {
            return TwapiReturnMissingOptValueError(interp, objv[i - 1]);
        }
        switch (index) {
        case OPT_APP:
            appObj = objv[i];
            break;
        case OPT_ACTIVITY:
	    if (ObjToGUID(interp, objv[i], &activity_guid) != TCL_OK) {
                return TCL_ERROR;
            }
            activity_guidP = &activity_guid;
            break;
        case OPT_RELATED_ACTIVITY:
            if (ObjToGUID(interp, objv[i], &related_activity_guid) != TCL_OK) {
                return TCL_ERROR;
            }
            related_activity_guidP = &related_activity_guid;
            break;
        }
    }

    if (TwapiParseSeverity(interp, objv[1], &level) != TCL_OK) {
        return TCL_ERROR;
    }

    if (gEvtRegHandle == 0) {
	/* Initialization of gEvtRegHandle must have failed */
        return TwapiReturnErrorMsg(interp,
                                   TWAPI_INIT_FAILURE,
                                   "TWAPI EVT Provider handle is NULL.");
    }

    EVENT_DATA_DESCRIPTOR   data[3];
    const EVENT_DESCRIPTOR *evdP;
    switch (level) {
    case 1: evdP = &TWAPI_EVT_EVENT_CRITICAL; break;
    case 2: evdP = &TWAPI_EVT_EVENT_ERROR; break;
    case 3: evdP = &TWAPI_EVT_EVENT_WARNING; break;
    case 4: evdP = &TWAPI_EVT_EVENT_INFORMATIONAL; break;
    case 5: evdP = &TWAPI_EVT_EVENT_VERBOSE; break;
    default:
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
    }

    if (!EventEnabled(gEvtRegHandle, evdP)) {
        /* not enabled so no point writing it - will be discarded anyways */
        return TCL_OK;
    }

    Tcl_Size utf_len;
    const char *utfP;

    Tcl_DString ds;
    LPCWSTR msgP;
    ULONG msg_num_bytes;
    Tcl_DStringInit(&ds);
    utfP    = Tcl_GetStringFromObj(objv[2], &utf_len);
    msgP    = (LPCWSTR) Tcl_UtfToWCharDString(utfP, utf_len, &ds);
    msg_num_bytes = (ULONG) Tcl_DStringLength(&ds) + sizeof(WCHAR); /* + L'\0' */

    Tcl_DString dsApp;
    LPCWSTR app_nameP;
    ULONG app_name_num_bytes;
    Tcl_DStringInit(&dsApp);
    if (appObj) {
        utfP = Tcl_GetStringFromObj(appObj, &utf_len);
        app_nameP = (LPCWSTR) Tcl_UtfToWCharDString(utfP, utf_len, &dsApp);
        app_name_num_bytes = (ULONG) Tcl_DStringLength(&dsApp) + sizeof(WCHAR);
    }
    else {
#define DEFAULT_APP_NAME L"Tcl Application"
        app_nameP          = DEFAULT_APP_NAME;
        app_name_num_bytes = (ULONG) sizeof(DEFAULT_APP_NAME); /* inc. \0 */
#undef DEFAULT_APP_NAME
    }

    EventDataDescCreate(&data[0], app_nameP, (ULONG)app_name_num_bytes);
    LPCWSTR exe_path = TwapiWinPathGet(&gExePath);
    EventDataDescCreate(&data[1],
                        exe_path ? exe_path : L"",
                        (ULONG)(((exe_path ? gExePathLen : 0) + 1)
                            * sizeof(WCHAR)));
    EventDataDescCreate(&data[2], msgP, (ULONG)msg_num_bytes);
    ULONG status = EventWriteEx(gEvtRegHandle,
                                evdP,
                                0ULL, /* Filter */
                                0UL,  /* Flags */
                                activity_guidP,
                                related_activity_guidP,
                                3,
                                data);
    Tcl_DStringFree(&dsApp);
    Tcl_DStringFree(&ds);

    if (status != ERROR_SUCCESS) {
    	return Twapi_AppendSystemError(interp, status);
    }
    return TCL_OK;
}


int TwapiEvtInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct tcl_dispatch_s EvtTclDispatch[] = {
        DEFINE_TCL_CMD(GetEVT_VARIANT, Twapi_EvtGetEVT_VARIANTObjCmd),
        DEFINE_TCL_CMD(Twapi_EvtRenderValues, Twapi_EvtRenderValuesObjCmd),
        DEFINE_TCL_CMD(Twapi_EvtRenderUnicode, Twapi_EvtRenderUnicodeObjCmd),
        DEFINE_TCL_CMD(EvtNext, Twapi_EvtNextObjCmd),
        DEFINE_TCL_CMD(EvtCreateRenderContext, Twapi_EvtCreateRenderContextObjCmd),
        DEFINE_TCL_CMD(EvtFormatMessage, Twapi_EvtFormatMessageObjCmd),
        DEFINE_TCL_CMD(EvtOpenSession, Twapi_EvtOpenSessionObjCmd),
        DEFINE_TCL_CMD(evt_log, Twapi_EvtLogObjCmd),
        DEFINE_TCL_CMD(Twapi_ExtractEVT_RENDER_VALUES, Twapi_ExtractEVT_RENDER_VALUESObjCmd),
        DEFINE_TCL_CMD(Twapi_ExtractEVT_RENDER_VALUES, Twapi_ExtractEVT_RENDER_VALUESObjCmd),
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
        DEFINE_FNCODE_CMD(evt_create_bookmark, 11),
        DEFINE_FNCODE_CMD(evt_update_bookmark, 12),
        DEFINE_FNCODE_CMD(evt_free, 13), // TBD docs
        DEFINE_FNCODE_CMD(evt_close, 14),
        DEFINE_FNCODE_CMD(evt_cancel, 102), // TBD docs
        DEFINE_FNCODE_CMD(EvtOpenChannelEnum, 103),
        DEFINE_FNCODE_CMD(EvtNextChannelPath, 104),
        DEFINE_FNCODE_CMD(evt_channel_config_save, 105),
        DEFINE_FNCODE_CMD(EvtOpenPublisherEnum, 106),
        DEFINE_FNCODE_CMD(EvtOpenEventMetadataEnum, 107),
        DEFINE_FNCODE_CMD(EvtNextEventMetadata, 108),
        DEFINE_FNCODE_CMD(EvtNextPublisherId, 109),
        DEFINE_FNCODE_CMD(EvtGetObjectArraySize, 110),
    };

    TwapiDefineTclCmds(interp, ARRAYSIZE(EvtTclDispatch), EvtTclDispatch, ticP);
    TwapiDefineFncodeCmds(interp, ARRAYSIZE(EvtFnDispatch), EvtFnDispatch, Twapi_EvtCallObjCmd);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(EvtVariantGetDispatch), EvtVariantGetDispatch, "twapi::GetEVT_VARIANT");

    return TCL_OK;
}

static int EvtModuleOneTimeInit(void *arg)
{
    /* Ignore return value so failure does not prevent the whole module from loading. */
    (void) EventRegister(
        &TWAPI_EVT_PROVIDER,
        NULL,                 /* EnableCallback: none               */
        NULL,                 /* CallbackContext                    */
        &gEvtRegHandle);

    return TCL_OK;
}

/* Called when interp is deleted */
static void TwapiEvtCleanup(TwapiInterpContext *ticP)
{
    if (gEvtRegHandle) {
        EventUnregister(gEvtRegHandle);
        gEvtRegHandle = 0;
    }
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
int Twapi_evt_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiEvtInitCalls,
        TwapiEvtCleanup
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    /* Init unless already done. */
    if (! TwapiDoOneTimeInit(&gEvtInitialized, EvtModuleOneTimeInit, interp))
        return TCL_ERROR;

    /* NEW_TIC since we have a cleanup routine */
    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, NEW_TIC) ? TCL_OK : TCL_ERROR;
}

