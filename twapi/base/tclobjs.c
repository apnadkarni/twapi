/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "tclTomMath.h"

/*
 * Used for deciphering  unknown types when passing to COM. Note
 * any or all of these may be NULL.
 */
static struct TwapiTclTypeMap {
    char *typename;
    Tcl_ObjType *typeptr;
} gTclTypes[TWAPI_TCLTYPE_BOUND];


int TwapiInitTclTypes(void)
{
    int i;

    gTclTypes[TWAPI_TCLTYPE_NONE].typename = "none";
    gTclTypes[TWAPI_TCLTYPE_NONE].typeptr = NULL; /* No internal type set */
    gTclTypes[TWAPI_TCLTYPE_STRING].typename = "string";
    gTclTypes[TWAPI_TCLTYPE_BOOLEAN].typename = "boolean";
    gTclTypes[TWAPI_TCLTYPE_INT].typename = "int";
    gTclTypes[TWAPI_TCLTYPE_DOUBLE].typename = "double";
    gTclTypes[TWAPI_TCLTYPE_BYTEARRAY].typename = "bytearray";
    gTclTypes[TWAPI_TCLTYPE_LIST].typename = "list";
    gTclTypes[TWAPI_TCLTYPE_DICT].typename = "dict";
    gTclTypes[TWAPI_TCLTYPE_WIDEINT].typename = "wideInt";
    gTclTypes[TWAPI_TCLTYPE_BOOLEANSTRING].typename = "booleanString";

    for (i = 1; i < ARRAYSIZE(gTclTypes); ++i) {
        gTclTypes[i].typeptr =
            Tcl_GetObjType(gTclTypes[i].typename); /* May be NULL */
    }

    /* "booleanString" type is not always registered (if ever). Get it
     *  by hook or by crook
     */
    if (gTclTypes[TWAPI_TCLTYPE_BOOLEANSTRING].typeptr == NULL) {
        Tcl_Obj *objP = Tcl_NewStringObj("true", -1);
        Tcl_GetBooleanFromObj(NULL, objP, &i);
        /* This may still be NULL, but what can we do ? */
        gTclTypes[TWAPI_TCLTYPE_BOOLEANSTRING].typeptr = objP->typePtr;
    }    

    return TCL_OK;
}

int TwapiGetTclType(Tcl_Obj *objP)
{
    int i;
    
    for (i = 0; i < ARRAYSIZE(gTclTypes); ++i) {
        if (gTclTypes[i].typeptr == objP->typePtr)
            return i;
    }

    return -1;
}

int Twapi_GetTclTypeObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    if (objv[1]->typePtr != NULL) {
        Tcl_SetObjResult(interp, Tcl_NewStringObj(objv[1]->typePtr->name, -1));
    } else {
        /* Leave result as empty string */
    }
    return TCL_OK;
}

int Twapi_InternalCastObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    Tcl_Obj *objP;
    Tcl_ObjType *typeP;
    const char *typename;
    int i;

    if (objc != 3)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    typename = Tcl_GetString(objv[1]);

    if (*typename == '\0') {
        /* No type, keep as is */
        Tcl_SetObjResult(interp, objv[2]);
        return TCL_OK;
    }
        
    /*
     * We special case "boolean" and "booleanString" because they will keep
     * numerics as numerics while we want to force to boolean Tcl_Obj type.
     * We do this even before GetObjType because "booleanString" is not
     * even a registered type in Tcl.
     */
    if (STREQ(typename, "boolean") || STREQ(typename, "booleanString")) {
        if (Tcl_GetBooleanFromObj(interp, objv[2], &i) == TCL_ERROR)
            return TCL_ERROR;
        /* Directly calling Tcl_NewBooleanObj returns an int type object */
        objP = Tcl_NewStringObj(i ? "true" : "false", -1);
        Tcl_GetBooleanFromObj(NULL, objP, &i);
        Tcl_SetObjResult(interp, objP);
        return TCL_OK;
    }

    typeP = Tcl_GetObjType(typename);
    if (typeP == NULL) {
        Tcl_AppendResult(interp, "Invalid or unknown Tcl type '", typename, "'", NULL);
        return TCL_ERROR;
    }
        
    if (objv[2]->typePtr == typeP) {
        /* If type is already correct, no need to do anything */
        objP = objv[2];
    } else {
        /*
         * Need to convert it. If not shared, do in place else allocate
         * new object
         */

        if (Tcl_IsShared(objv[2])) {
            objP = Tcl_DuplicateObj(objv[2]);
        } else {
            objP = objv[2];
        }
        
        if (Tcl_ConvertToType(interp, objP, typeP) == TCL_ERROR) {
            if (objP != objv[2]) {
                Tcl_DecrRefCount(objP);
            }
            return TCL_ERROR;
        }
    }

    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}


/*
 * Generic function for setting a Tcl result. Note following special cases
 * - for TRT_OBJV types, the objv[] objects are added to a list (and
 *   their ref counts are therby incremented)
 * - for TRT_VARIANT types, VariantClear is called. This allows
 *   callers to directly return after calling this function without
 *   having to clean up the variant before returning.
 * - Similarly, TRT_UNICODE_DYNAMIC, TRT_CHARS_DYNAMIC, TRT_OLESTR and TRT_PIDL
 *   have their memory freed
 **/
TCL_RESULT TwapiSetResult(Tcl_Interp *interp, TwapiResult *resultP)
{
    char *typenameP;
    Tcl_Obj *resultObj = NULL;

    switch (resultP->type) {
    case TRT_GETLASTERROR:      /* Error in GetLastError() */
        return TwapiReturnSystemError(interp);

    case TRT_BOOL:
        resultObj = Tcl_NewBooleanObj(resultP->value.bval);
        break;

    case TRT_EXCEPTION_ON_FALSE:
    case TRT_NONZERO_RESULT:
        /* If 0, generate exception */
        if (! resultP->value.ival)
            return TwapiReturnSystemError(interp);

        if (resultP->type == TRT_NONZERO_RESULT)
            resultObj = Tcl_NewLongObj(resultP->value.ival);
        /* else an empty result is returned */
        break;

    case TRT_EXCEPTION_ON_ERROR:
        /* If non-0, generate exception */
        if (resultP->value.ival) {
            return Twapi_AppendSystemError(interp, resultP->value.ival);
        }
        break;

    case TRT_EXCEPTION_ON_WNET_ERROR:
        /* If non-0, generate exception */
        if (resultP->value.ival) {
            return Twapi_AppendWNetError(interp, resultP->value.ival);
        }
        break;

    case TRT_EXCEPTION_ON_MINUSONE:
        /* If -1, generate exception */
        if (resultP->value.ival == -1)
            return TwapiReturnSystemError(interp);

        /* Other values are to be returned */
        resultObj = Tcl_NewLongObj(resultP->value.ival);
        break;

    case TRT_UNICODE_DYNAMIC:
    case TRT_UNICODE:
        if (resultP->value.unicode.str) {
            resultObj = ObjFromUnicodeN(resultP->value.unicode.str,
                                        resultP->value.unicode.len);
        }
        break;

    case TRT_CHARS_DYNAMIC:
    case TRT_CHARS:
        if (resultP->value.chars.str)
            resultObj = Tcl_NewStringObj(resultP->value.chars.str,
                                         resultP->value.chars.len);
        break;

    case TRT_BINARY:
        resultObj = Tcl_NewByteArrayObj(resultP->value.binary.p,
                                        resultP->value.binary.len);
        break;

    case TRT_OBJ:
        resultObj = resultP->value.obj;
        break;

    case TRT_OBJV:
        resultObj = Tcl_NewListObj(resultP->value.objv.nobj, resultP->value.objv.objPP);
        break;

    case TRT_RECT:
        resultObj = ObjFromRECT(&resultP->value.rect);
        break;

    case TRT_POINT:
        resultObj = ObjFromPOINT(&resultP->value.point);
        break;

    case TRT_VALID_HANDLE:
        if (resultP->value.hval == INVALID_HANDLE_VALUE) {
            return TwapiReturnSystemError(interp);
        }
        resultObj = ObjFromHANDLE(resultP->value.hval);
        break;

    case TRT_HWND:
        // Note unlike other handles, we do not return an error if NULL
        resultObj = ObjFromOpaque(resultP->value.hwin, "HWND");
        break;

    case TRT_HMODULE:
        // Note unlike other handles, we do not return an error if NULL
        resultObj = ObjFromOpaque(resultP->value.hmodule, "HMODULE");
        break;

    case TRT_HANDLE:
    case TRT_HGLOBAL:
    case TRT_HDC:
    case TRT_HDESK:
    case TRT_HMONITOR:
    case TRT_HWINSTA:
    case TRT_SC_HANDLE:
    case TRT_LSA_HANDLE:
    case TRT_SEC_WINNT_AUTH_IDENTITY:
    case TRT_HDEVINFO:
    case TRT_HRGN:
    case TRT_NONNULL_LPVOID:
        if (resultP->value.hval == NULL) {
            return TwapiReturnSystemError(interp);
        }
        switch (resultP->type) {
        case TRT_HANDLE:
            typenameP = "HANDLE";
            break;
        case TRT_HGLOBAL:
            typenameP = "HGLOBAL";
            break;
        case TRT_HDC:
            typenameP = "HDC";
            break;
        case TRT_HDESK:
            typenameP = "HDESK";
            break;
        case TRT_HMONITOR:
            typenameP = "HMONITOR";
            break;
        case TRT_HWINSTA:
            typenameP = "HWINSTA";
            break;
        case TRT_SC_HANDLE:
            typenameP = "SC_HANDLE";
            break;
        case TRT_LSA_HANDLE:
            typenameP = "LSA_HANDLE";
            break;
        case TRT_SEC_WINNT_AUTH_IDENTITY:
            typenameP = "SEC_WINNT_AUTH_IDENTITY_W*";
            break;
        case TRT_HDEVINFO:
            typenameP = "HDEVINFO";
            break;
        case TRT_NONNULL_LPVOID:
            typenameP = "void*";
            break;
        case TRT_HRGN:
            typenameP = "HRGN";
            break;
        default:
            Tcl_SetResult(interp, "Internal error: TwapiSetResult - inconsistent nesting of case statements", TCL_STATIC);
            return TCL_ERROR;
        }
        resultObj = ObjFromOpaque(resultP->value.hval, typenameP);
        break;

    case TRT_LPVOID:
        resultObj = ObjFromOpaque(resultP->value.hval, "void*");
        break;

    case TRT_DWORD:
        resultObj = Tcl_NewLongObj(resultP->value.ival);
        break;

    case TRT_WIDE:
        resultObj = Tcl_NewWideIntObj(resultP->value.wide);
        break;

    case TRT_DOUBLE:
        resultObj = Tcl_NewDoubleObj(resultP->value.dval);
        break;

    case TRT_FILETIME:
        resultObj = ObjFromFILETIME(&resultP->value.filetime);
        break;

    case TRT_SYSTEMTIME:
        resultObj = ObjFromSYSTEMTIME(&resultP->value.systime);
        break;

    case TRT_EMPTY:
        Tcl_ResetResult(interp);
        break;

    case TRT_UUID:
        resultObj = ObjFromUUID(&resultP->value.uuid);
        break;

    case TRT_GUID:
        resultObj = ObjFromGUID(&resultP->value.guid);
        break;

    case TRT_LUID:
        resultObj = ObjFromLUID(&resultP->value.luid);
        break;
        
    case TRT_DWORD_PTR:
        resultObj = ObjFromDWORD_PTR(resultP->value.dwp);
        break;

    case TRT_INTERFACE:
        resultObj = ObjFromOpaque(resultP->value.ifc.p, resultP->value.ifc.name);
        break;

    case TRT_VARIANT:
        resultObj = ObjFromVARIANT(&resultP->value.var, 0);
        break;

    case TRT_LPOLESTR:
        if (resultP->value.lpolestr) {
            resultObj = ObjFromUnicode(resultP->value.lpolestr);
        } else
            Tcl_ResetResult(interp);
        break;

    case TRT_PIDL:
        if (resultP->value.pidl) {
            resultObj = ObjFromPIDL(resultP->value.pidl);
        } else
            Tcl_ResetResult(interp);
        break;

    case TRT_OPAQUE:
        resultObj = ObjFromOpaque(resultP->value.opaque.p,
                                  resultP->value.opaque.name);
        break;

    case TRT_TCL_RESULT:
        /* interp result already stored. Status in ival */
        return resultP->value.ival;
        
    case TRT_NTSTATUS:
        if (resultP->value.ival != STATUS_SUCCESS)
            return Twapi_AppendSystemError(interp,
                                           TwapiNTSTATUSToError(resultP->value.ival));

        break;

    case TRT_BADFUNCTIONCODE:
        return TwapiReturnError(interp, TWAPI_INVALID_FUNCTION_CODE);

    case TRT_TWAPI_ERROR:
        return TwapiReturnError(interp, resultP->value.ival);

    default:
        Tcl_SetResult(interp, "Unknown TwapiResultType type code passed to TwapiSetResult", TCL_STATIC);
        return TCL_ERROR;
    }

    TwapiClearResult(resultP);  /* Clear out resources */

    if (resultObj)
        Tcl_SetObjResult(interp, resultObj);


    return TCL_OK;
}

/* Frees allocated resources, sets resultP to type TRT_EMPTY */
void TwapiClearResult(TwapiResult *resultP)
{
    switch (resultP->type) {
    case TRT_UNICODE_DYNAMIC:
        if (resultP->value.unicode.str)
            TwapiFree(resultP->value.unicode.str);
        break;
    case TRT_CHARS_DYNAMIC:
        if (resultP->value.chars.str)
            TwapiFree(resultP->value.chars.str);
        break;
    case TRT_VARIANT:
        VariantClear(&resultP->value.var);
        break;
    case TRT_LPOLESTR:
        if (resultP->value.lpolestr)
            CoTaskMemFree(resultP->value.lpolestr);
        break;

    case TRT_PIDL:
        if (resultP->value.pidl)
            CoTaskMemFree(resultP->value.pidl);
        break;

    default:
        break;                  /* Nothing to clear */

    }

    resultP->type = TRT_EMPTY;
}


/* Appends the given strings objv[] to a result object, separated by
 * the passed string. The passed resultObj must not be a shared object!
 */
Tcl_Obj *TwapiAppendObjArray(Tcl_Obj *resultObj, int objc, Tcl_Obj **objv,
                         char *joiner)
{
    int i;
    int len;
    char *s;
    int joinlen = (int) strlen(joiner);
#if 0
    Not needed - Tcl_AppendToObj will do the panic for us below
    if (Tcl_IsShared(resultObj)) {
        panic("TwapiAppendObjArray called on shared object");
    }
#endif

    for (i = 0;  i < objc;  ++i) {
        s = Tcl_GetStringFromObj(objv[i], &len);
        if (i > 0) {
            Tcl_AppendToObj(resultObj, joiner, joinlen);
        }
        Tcl_AppendToObj(resultObj, s, len);
    }

    return resultObj;
}



LPWSTR ObjToLPWSTR_NULL_IF_EMPTY(Tcl_Obj *objP)
{
    if (objP) {
        int len;
        LPWSTR p = Tcl_GetUnicodeFromObj(objP, &len);
        if (len > 0)
            return p;
    }
    return NULL;
}

LPWSTR ObjToLPWSTR_WITH_NULL(Tcl_Obj *objP)
{
    if (objP) {
        LPWSTR s = Tcl_GetUnicode(objP);
        if (lstrcmpW(s, NULL_TOKEN_L) == 0) {
            s = NULL;
        }
        return s;
    }
    return NULL;
}

// Return SysAllocStringLen-allocated string
int ObjToBSTR(Tcl_Interp *interp, Tcl_Obj *objP, BSTR *bstrP)
{
    int len;
    WCHAR *wcharP;

    if (objP == NULL) {
        wcharP = L"";
        len = 0;
    } else {
        wcharP = Tcl_GetUnicodeFromObj(objP, &len);
    }
    if (bstrP) {
        *bstrP = SysAllocStringLen(wcharP, len);
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, E_OUTOFMEMORY);
    }
}

Tcl_Obj *ObjFromBSTR (BSTR bstr)
{
    return bstr ? ObjFromUnicodeN(bstr, SysStringLen(bstr)) : Tcl_NewStringObj("", 0);
}

Tcl_Obj *ObjFromStringLimited(const char *strP, int max, int *remainP)
{
    int len;

    if (max < 0) {
        if (remainP)
            *remainP = 0;
        return Tcl_NewStringObj(strP, -1);
    }        

    for (len = 0; len < max && strP[len]; ++len)
        ;
        
    /* We have to be careful about setting *remainP since the loop may
       terminate because of max exceeded or because \0 encountered.
    */
    if (remainP) {
        if ((len+1) <= max)
            *remainP = (max - (len+1)); /* \0 case */
        else
            *remainP = 0;       /* max reached case */
    }        

    return Tcl_NewStringObj(strP, len);
}

Tcl_Obj *ObjFromUnicodeLimited(const WCHAR *strP, int max, int *remainP)
{
    int len;

    if (max < 0) {
        if (remainP)
            *remainP = 0;
        return ObjFromUnicode(strP);
    }        

    for (len = 0; len < max && strP[len]; ++len)
        ;
        
    /* We have to be careful about setting *remainP since the loop may
       terminate because of max exceeded or because \0 encountered.
    */
    if (remainP) {
        if ((len+1) <= max)
            *remainP = (max - (len+1)); /* \0 case */
        else
            *remainP = 0;       /* max reached case */
    }        

    return ObjFromUnicodeN(strP, len);
}



/*
 * Gets an integer from an object within the specified range
 * Returns TCL_OK if integer within range [low,high], else error
 */
int ObjToRangedInt(Tcl_Interp *interp, Tcl_Obj *obj, int low, int high, int *iP)
{
    int i;
    if (Tcl_GetIntFromObj(interp, obj, &i) != TCL_OK)
        return TCL_ERROR;

    if (i < low || i > high) {
        if (interp) {
            Tcl_SetObjResult(interp,
                             Tcl_ObjPrintf("Integer '%d' not within range %d-%d", i, low, high));
        }
        return TCL_ERROR;
    }

    if (iP)
        *iP = i;

    return TCL_OK;
}

/*
 * Convert a system time structure to a list
 * Year month day hour min sec msecs
 */
Tcl_Obj *ObjFromSYSTEMTIME(LPSYSTEMTIME timeP)
{
    Tcl_Obj *objv[7];

    objv[0] = Tcl_NewIntObj(timeP->wYear);
    objv[1] = Tcl_NewIntObj(timeP->wMonth);
    objv[2] = Tcl_NewIntObj(timeP->wDay);
    objv[3] = Tcl_NewIntObj(timeP->wHour);
    objv[4] = Tcl_NewIntObj(timeP->wMinute);
    objv[5] = Tcl_NewIntObj(timeP->wSecond);
    objv[6] = Tcl_NewIntObj(timeP->wMilliseconds);

    return Tcl_NewListObj(7, objv);
}


/*
 * Convert a Tcl Obj to SYSTEMTIME
 */
int ObjToSYSTEMTIME(Tcl_Interp *interp, Tcl_Obj *timeObj, LPSYSTEMTIME timeP)
{
    Tcl_Obj **objv;
    int       objc;
    int       itemp;
    int       hindex;

    if (Tcl_ListObjGetElements(interp, timeObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    /*
     * There may be variable number of arguments.
     * 0 - current date and time
     * 1 - hour, min,sec,ms=0, current date
     * 2 - hour, min, sec,ms=0, current date
     * 3 - hour, min, sec, ms=0, current date
     * 4 - hour, min, sec, ms, current date
     * 5 - day, hour, min, sec, ms, current month and year
     * 6 - month, day, hour, min, sec, ms, current year
     * 7 - year, month, day, hour, min, sec, ms
     * Extra args ignored
     */
    if (objc < 7)
        GetSystemTime(timeP);
    timeP->wMinute = 0;
    timeP->wSecond = 0;
    timeP->wMilliseconds = 0;

    hindex = 0;               /* Slot where the Hindex is to be found */
    switch (objc) {
    case 0:
        break;              /* Just use current time */
    default:
        /* FALLTHRU (extra args ignored) */
    case 7:
        if (ObjToRangedInt(interp, objv[hindex], 0, 30827, &itemp) != TCL_OK)
            return TCL_ERROR;
        timeP->wYear = (WORD) itemp;
        ++hindex;                 /* Bump up remaining slots */
        /* FALLTHRU */
    case 6:
        if (ObjToRangedInt(interp, objv[hindex], 1, 12, &itemp) != TCL_OK)
            return TCL_ERROR;
        timeP->wMonth = (WORD) itemp;
        ++hindex;                 /* Bump up remaining slots */
        /* FALLTHRU */
    case 5:
        /* Should do more validation for day range */
        if (ObjToRangedInt(interp, objv[hindex], 0, 31, &itemp) != TCL_OK)
            return TCL_ERROR;
        timeP->wDay = (WORD) itemp;
        ++hindex;                 /* Bump up remaining slots */
        /* FALLTHRU */
    case 4:
        /* Max 60 to allow for leap seconds */
        if (ObjToRangedInt(interp, objv[hindex+3], 0, 999, &itemp) != TCL_OK)
            return TCL_ERROR;
        timeP->wMilliseconds = (WORD) itemp;
        /* FALLTHRU */
    case 3:
        /* Max 60 to allow for leap seconds */
        if (ObjToRangedInt(interp, objv[hindex+2], 0, 60, &itemp) != TCL_OK)
            return TCL_ERROR;
        timeP->wSecond = (WORD) itemp;
        /* FALLTHRU */
    case 2:
        if (ObjToRangedInt(interp, objv[hindex+1], 0, 59, &itemp) != TCL_OK)
            return TCL_ERROR;
        timeP->wMinute = (WORD) itemp;
        /* FALLTHRU */
    case 1:
        if (ObjToRangedInt(interp, objv[hindex], 0, 23, &itemp) != TCL_OK)
            return TCL_ERROR;
        timeP->wHour = (WORD) itemp;
        break;
    }
    return TCL_OK;
}


Tcl_Obj *ObjFromFILETIME(FILETIME *ftimeP)
{
    LARGE_INTEGER large;
    large.LowPart = ftimeP->dwLowDateTime;
    large.HighPart = ftimeP->dwHighDateTime;
    return ObjFromLARGE_INTEGER(large);
}

int ObjToFILETIME(Tcl_Interp *interp, Tcl_Obj *obj, FILETIME *ftimeP)
{
    LARGE_INTEGER large;
    if (Tcl_GetWideIntFromObj(interp, obj, &large.QuadPart) != TCL_OK)
        return TCL_ERROR;

    if (ftimeP) {
        ftimeP->dwLowDateTime = large.LowPart;
        ftimeP->dwHighDateTime = large.HighPart;
    }

    return TCL_OK;
}

Tcl_Obj *ObjFromCY(const CY *cyP)
{
    /* TBD - for now just return as 8 byte wide int */
    return Tcl_NewWideIntObj(*(Tcl_WideInt *)cyP);
}


int ObjToCY(Tcl_Interp *interp, Tcl_Obj *obj, CY *cyP)
{
    Tcl_WideInt wi;
    if (Tcl_GetWideIntFromObj(interp, obj, &wi) != TCL_OK)
        return TCL_ERROR;

    if (cyP)
        *cyP = *(CY *)&wi;

    return TCL_OK;
}

Tcl_Obj *ObjFromDECIMAL(DECIMAL *decP)
{
    /* TBD - for now just return as string.
       Problem is how is this formatted (separators etc.) ?
    */
    Tcl_Obj *obj;
    BSTR bstr = NULL;
    if (VarBstrFromDec(decP, 0, 0, &bstr) != S_OK) {
        return Tcl_NewStringObj("", 0);
    }

    obj = ObjFromBSTR(bstr);
    SysFreeString(bstr);
    return obj;
}


int ObjToDECIMAL(Tcl_Interp *interp, Tcl_Obj *obj, DECIMAL *decP)
{
    HRESULT hr;
    DECIMAL dec;

    if (decP == NULL)
        decP = &dec;
    hr = VarDecFromStr(Tcl_GetUnicodeFromObj(obj, NULL), 0, 0, decP);
    if (FAILED(hr)) {
        if (interp)
            Twapi_AppendSystemError(interp, hr);
        return TCL_ERROR;
    }
    return TCL_OK;
}

Tcl_Obj *ObjFromPIDL(LPCITEMIDLIST pidl)
{
    /* Scan until we find an item with length 0 */
    unsigned char *p = (char *) pidl;
    do {
        /* p[1,0] is length of this item */
        int len = 256*p[1] + p[0];
        if (len == 0)
            break;              /* 0 length -> end of list */

        p += len;
    } while (1);

    /* p points to terminating null field */
    return Tcl_NewByteArrayObj((unsigned char *)pidl,
                               (int) (2 + p - (unsigned char *)pidl));

}

/* On success, returns TCL_OK and stores pointer to an ITEMIDLIST
   in *idlistP. The value stored may be NULL which is valid in many cases.
   On error, returns TCL_ERROR, with an error message in interp if it
   is not NULL.
   The ITEMIDLIST must be freed by caller by calling TwapiFreePIDL
*/
int ObjToPIDL(Tcl_Interp *interp, Tcl_Obj *objP, LPITEMIDLIST *idsPP)
{
    int      numbytes;
    LPITEMIDLIST idsP;

    idsP = (LPITEMIDLIST) Tcl_GetByteArrayFromObj(objP, &numbytes);
    if (numbytes < 2) {
        *idsPP = NULL;              /* Empty string */
        return TCL_OK;
    }
    else {
        /* Verify format. Passing bad PIDL's can crash the app */
        unsigned char *p = (unsigned char *) idsP;
        /*
         * At top of loop p points to length of next item
         * and p[0] and p[1] are guaranteed valid part of buffer.
         */
        while (p[0] || p[1]) {
            int itemlen = p[1]*256 + p[0]; /* Assumes little endian */

            /* Verify that item length fits within buffer
             * The 2 is for the trailing 2 null bytes
             */
            if (itemlen > (numbytes-2)) {
                Tcl_SetResult(interp,
                              "Invalid item id list format",
                              TCL_STATIC);
                return TCL_ERROR;
            }
            numbytes -= itemlen;
            p += itemlen;
        }
    }

    *idsPP = CoTaskMemAlloc(numbytes);
    if (*idsPP == NULL) {
        if (interp)
            Tcl_SetResult(interp,
                          "CoTaskMemAlloc failed in SHChangeNotify",
                          TCL_STATIC);
        return TCL_ERROR;
    }

    CopyMemory(*idsPP, idsP, numbytes);

    return TCL_OK;
}

void TwapiFreePIDL(LPITEMIDLIST idlistP)
{
    if (idlistP) {
        CoTaskMemFree(idlistP);
    }
}

Tcl_Obj *ObjFromGUID(GUID *guidP)
{
    wchar_t  str[40];
    Tcl_Obj *obj;


    if (guidP == NULL || StringFromGUID2(guidP, str, sizeof(str)/sizeof(str[0])) == 0)
        return Tcl_NewStringObj("", 0);

    obj = ObjFromUnicode(str);
    return obj;
}

int ObjToGUID(Tcl_Interp *interp, Tcl_Obj *objP, GUID *guidP)
{
    HRESULT hr;
    WCHAR *wsP;
    if (objP) {
        wsP = Tcl_GetUnicode(objP);

        /* Accept both GUID and UUID forms */
        if (*wsP == L'{') {
            /* GUID form */
            /* We *used* to use CLSIDFromString but it turns out that 
               accepts Prog IDs as valid GUIDs as well */
            if ((hr = IIDFromString(wsP, guidP)) != NOERROR) {
                Twapi_AppendSystemError(interp, hr);
                return TCL_ERROR;
            }
        } else {
            /* Might be UUID form */
            RPC_STATUS status = UuidFromStringW(wsP, guidP);
            if (status != RPC_S_OK) {
                Twapi_AppendSystemError(interp, status);
                return TCL_ERROR;
            }
        }
    } else
        ZeroMemory(guidP, sizeof(*guidP));
    return TCL_OK;
}

int ObjToGUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, GUID **guidPP)
{
    if (Tcl_GetCharLength(objP) == 0) {
        *guidPP = NULL;
        return TCL_OK;
    } else 
        return ObjToGUID(interp, objP, *guidPP);
}


Tcl_Obj *ObjFromUUID (UUID *uuidP)
{
    unsigned char *uuidStr;
    Tcl_Obj       *objP;
    if (UuidToStringA(uuidP, &uuidStr) != RPC_S_OK)
        return NULL;

    /* NOTE UUID and GUID have same binary format but are formatted
       differently based on the component. */
    objP = Tcl_NewStringObj(uuidStr, -1);
    RpcStringFree(&uuidStr);
    return objP;
}

int ObjToUUID(Tcl_Interp *interp, Tcl_Obj *objP, UUID *uuidP)
{
    /* NOTE UUID and GUID have same binary format but are formatted
       differently based on the component.  We accept both forms here */

    if (objP) {
        RPC_STATUS status = UuidFromStringA(Tcl_GetString(objP), uuidP);
        if (status != RPC_S_OK) {
            /* Try as GUID form */
            return ObjToGUID(interp, objP, uuidP);
        }
    } else {
        ZeroMemory(uuidP, sizeof(*uuidP));
    }
    return TCL_OK;
}

int ObjToUUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, UUID **uuidPP)
{
    if (Tcl_GetCharLength(objP) == 0) {
        *uuidPP = NULL;
        return TCL_OK;
    } else 
        return ObjToUUID(interp, objP, *uuidPP);
}

Tcl_Obj *ObjFromLSA_UNICODE_STRING(const LSA_UNICODE_STRING *lsauniP)
{
    /* Note LSA_UNICODE_STRING Length field is in *bytes* NOT characters */
    return ObjFromUnicodeN(lsauniP->Buffer, lsauniP->Length / sizeof(WCHAR));
}

void ObjToLSA_UNICODE_STRING(Tcl_Obj *objP, LSA_UNICODE_STRING *lsauniP)
{
    int nchars;
    lsauniP->Buffer = Tcl_GetUnicodeFromObj(objP, &nchars);
    lsauniP->Length = (USHORT) (sizeof(WCHAR)*nchars); /* in bytes */
    lsauniP->MaximumLength = lsauniP->Length;
}


/* interp may be NULL */
int ObjFromSID (Tcl_Interp *interp, SID *sidP, Tcl_Obj **objPP)
{
    char *strP;

    if (ConvertSidToStringSidA(sidP, &strP) == 0) {
        if (interp)
            TwapiReturnSystemError(interp);
        return TCL_ERROR;
    }

    *objPP = Tcl_NewStringObj(strP, -1);
    LocalFree(strP);
    return TCL_OK;
}

/* Like ObjFromSID but returns empty object on error */
Tcl_Obj *ObjFromSIDNoFail(SID *sidP)
{
    Tcl_Obj *objP;
    return (ObjFromSID(NULL, sidP, &objP) == TCL_OK ? objP : Tcl_NewStringObj("", 0));
}


/*
 * Convert a Tcl list to a "MULTI_SZ" list of Unicode strings, terminated
 * with two nulls. Returns TCL_OK on success with a dynamically alloced multi_sz
 * string in *multiszPtrPtr. Returns TCL_ERROR on failure
 */
int ObjToMultiSz (
     Tcl_Interp *interp,
     Tcl_Obj    *listPtr,
     LPCWSTR     *multiszPtrPtr
    )
{
    int       i;
    int       len;
    Tcl_Obj  *objPtr;
    LPWSTR    buf;
    LPWSTR    dst;
    LPCWSTR   src;

    *multiszPtrPtr = NULL;
    for (i=0, len=0; ; ++i) {
        if (Tcl_ListObjIndex(interp, listPtr, i, &objPtr) == TCL_ERROR)
            return TCL_ERROR;
        if (objPtr == NULL)
            break;              /* No more items */
        len += Tcl_GetCharLength(objPtr) + 1;
    }

    ++len;                      /* One extra null char at the end */
    buf = TwapiAlloc(len*sizeof(*buf));

    for (i=0, dst=buf; ; ++i) {
        if (Tcl_ListObjIndex(interp, listPtr, i, &objPtr) == TCL_ERROR) {
            TwapiFree(buf);
            return TCL_ERROR;
        }
        if (objPtr == NULL)
            break;              /* No more items */
        src = Tcl_GetUnicodeFromObj(objPtr, &len);
        if (src) {
            ++len;               /* Include the terminating null */
            CopyMemory(dst, src, len*sizeof(*src));
            dst += len;
        }
    }

    /* Add the final terminating null */
    *dst = 0;

    *multiszPtrPtr = buf;
    return TCL_OK;
}


/*
 * Convert a "MULTI_SZ" list of Unicode strings, terminated with two nulls to
 * a Tcl list. For example, a list of three strings - "abc", "def" and
 * "hij" would look like 'abc\0def\0hij\0\0'. This function will create
 * a list Tcl_Obj and return it. Will return NULL on error.
 *
 * maxlen is provided because registry data can be badly formatted
 * by applications. So we optionally ensure we do not read beyond
 * maxlen characters.
 */
Tcl_Obj *ObjFromMultiSz(LPCWSTR lpcw, int maxlen)
{
    Tcl_Obj *listPtr = Tcl_NewListObj(0, NULL);
    LPCWSTR start = lpcw;

    if (lpcw == NULL || maxlen == 0)
        return listPtr;

    if (maxlen == -1)
        maxlen = INT_MAX;

    while ((lpcw - start) < maxlen && *lpcw) {
        LPCWSTR s;
        /* Locate end of this string */
        s = lpcw;
        while ((lpcw - start) < maxlen && *lpcw)
            ++lpcw;
        if (s == lpcw) {
            /* Zero-length string - end of multi-sz */
            break;
        }

        Tcl_ListObjAppendElement(NULL, listPtr, ObjFromUnicodeN(s, (int) (lpcw-s)));
        ++lpcw;            /* Point beyond this string, possibly beyond end */
    }

    return listPtr;
}

int ObjToWord(Tcl_Interp *interp, Tcl_Obj *obj, WORD *wordP)
{
    long lval;
    if (Tcl_GetLongFromObj(interp, obj, &lval) != TCL_OK)
        return TCL_ERROR;
    if (lval & 0xffff0000) {
        if (interp)
            Tcl_SetResult(interp, "Integer value must be less than 65536", TCL_STATIC);
        return TCL_ERROR;
    }
    *wordP = (WORD) lval;
    return TCL_OK;
}

int ObjToRECT (Tcl_Interp *interp, Tcl_Obj *obj, RECT *rectP)
{
    Tcl_Obj **objv;
    int       objc;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 4) {
        Tcl_SetResult(interp, "Need to specify exactly 4 integers for a RECT structure", TCL_STATIC);
        return TCL_ERROR;
    }
    if ((Tcl_GetLongFromObj(interp, objv[0], &rectP->left) != TCL_OK) ||
        (Tcl_GetLongFromObj(interp, objv[1], &rectP->top) != TCL_OK) ||
        (Tcl_GetLongFromObj(interp, objv[2], &rectP->right) != TCL_OK) ||
        (Tcl_GetLongFromObj(interp, objv[3], &rectP->bottom) != TCL_OK)) {
        return TCL_ERROR;
    }

    return TCL_OK;
}

/* *rectPP must be  VALID memory ! */
int ObjToRECT_NULL(Tcl_Interp *interp, Tcl_Obj *obj, RECT **rectPP)
{
    int len;
    if (Tcl_ListObjLength(interp, obj, &len) != TCL_OK)
        return TCL_ERROR;
    if (len == 0) {
        *rectPP = NULL;
        return TCL_OK;
    } else
        return ObjToRECT(interp, obj, *rectPP);
}


/* Return a Tcl Obj from a RECT structure */
Tcl_Obj *ObjFromRECT(RECT *rectP)
{
    Tcl_Obj *objv[4];

    objv[0] = Tcl_NewIntObj(rectP->left);
    objv[1] = Tcl_NewIntObj(rectP->top);
    objv[2] = Tcl_NewIntObj(rectP->right);
    objv[3] = Tcl_NewIntObj(rectP->bottom);
    return Tcl_NewListObj(4, objv);
}

/* Return a Tcl Obj from a POINT structure */
Tcl_Obj *ObjFromPOINT(POINT *ptP)
{
    Tcl_Obj *objv[2];

    objv[0] = Tcl_NewIntObj(ptP->x);
    objv[1] = Tcl_NewIntObj(ptP->y);
    return Tcl_NewListObj(2, objv);
}

/* Convert a Tcl_Obj to a POINT */
int ObjToPOINT (Tcl_Interp *interp, Tcl_Obj *obj, POINT *ptP)
{
    Tcl_Obj **objv;
    int       objc;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 2) {
        Tcl_SetResult(interp, "Need to specify exactly 2 integers for a POINT structure", TCL_STATIC);
        return TCL_ERROR;
    }
    if ((Tcl_GetLongFromObj(interp, objv[0], &ptP->x) != TCL_OK) ||
        (Tcl_GetLongFromObj(interp, objv[1], &ptP->y) != TCL_OK)) {
        return TCL_ERROR;
    }

    return TCL_OK;
}

/* Return a Tcl Obj from a POINT structure */
Tcl_Obj *ObjFromPOINTS(POINTS *ptP)
{
    Tcl_Obj *objv[2];

    objv[0] = Tcl_NewIntObj((int) ptP->x);
    objv[1] = Tcl_NewIntObj((int) ptP->y);

    return Tcl_NewListObj(2, objv);
}


Tcl_Obj *ObjFromLUID (const LUID *luidP)
{
    return Tcl_ObjPrintf("%.8x-%.8x", luidP->HighPart, luidP->LowPart);
}

/*
 * Convert a string LUID to a LUID structure. Returns luidP on success,
 * else NULL on failure (invalid string format). interp may be NULL
 */
int ObjToLUID(Tcl_Interp *interp, Tcl_Obj *objP, LUID *luidP)
{
    char *markerP;
    int   len;
    char *strP = Tcl_GetStringFromObj(objP, &len);

    /* Format must be "XXXXXXXX-XXXXXXXX" */
    if ((len == 17) && (strP[8] == '-')) {
        luidP->HighPart = strtoul(strP, &markerP, 16);
        if (markerP == (strP+8)) {
            luidP->LowPart = strtoul(&strP[9], &markerP, 16);
            if (markerP == (strP+17))
                return TCL_OK;
        }
    }
    if (interp) {
        Tcl_SetResult(interp, "Invalid LUID format: ", TCL_STATIC);
        Tcl_AppendResult(interp, strP, NULL);
    }
    return TCL_ERROR;
}

/* *luidP MUST POINT TO A LUID STRUCTURE. This function will write the
 *  LUID there.   However if the obj is empty, it will store NULL in *luidP
*/
int ObjToLUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, LUID **luidPP)
{
    if (Tcl_GetCharLength(objP) == 0) {
        *luidPP = NULL;
        return TCL_OK;
    } else
        return ObjToLUID(interp, objP, *luidPP);
}


Tcl_Obj *ObjFromRegValue(Tcl_Interp *interp, int regtype,
                         BYTE *bufP, int count)
{
    Tcl_Obj *objv[2];
    char *typestr = NULL;

    switch (regtype) {
    case REG_LINK:
        typestr = "link";
        // FALLTHRU
    case REG_SZ:
        if (typestr == NULL)
            typestr = "sz";
        // FALLTHRU
    case REG_EXPAND_SZ:
        if (typestr == NULL)
            typestr = "expand_sz";
        /*
         * As per MS docs, may not always be null terminated.
         * If it is, we need to strip off the null.
         */
        count /= 2;             /*  Assumed to be Unicode. */
        if (count && bufP[count-1] == 0)
            --count;        /* Do not include \0 */
        objv[1] = ObjFromUnicodeN((WCHAR *)bufP, count);
        break;
            
    case REG_DWORD_BIG_ENDIAN:
        /* Since we are returning *typed* values, do not byte swap */
        /* FALLTHRU */
    case REG_DWORD:
        if (count != 4)
            goto badformat;
        typestr = regtype == REG_DWORD ? "dword" : "dword_be";
        objv[1] = Tcl_NewLongObj(*(int *)bufP);
        break;

    case REG_QWORD:
        if (count != 8)
            goto badformat;
        typestr = regtype == REG_QWORD ? "qword" : "qword_be";
        objv[1] = Tcl_NewWideIntObj(*(Tcl_WideInt *)bufP);
        break;

    case REG_MULTI_SZ:
        typestr = "multi_sz";
        objv[1] = ObjFromMultiSz((LPCWSTR) bufP, count/2);
        break;

    case REG_BINARY:
        typestr = "binary";
        // FALLTHRU
    default:
        objv[1] = Tcl_NewByteArrayObj(bufP, count);
        break;
    }

    if (typestr)
        objv[0] = Tcl_NewStringObj(typestr, -1);
    else
        objv[0] = Tcl_NewLongObj(regtype);
    return Tcl_NewListObj(2, objv);

badformat:
    Tcl_SetResult(interp, "Badly formatted registry value", TCL_STATIC);
    return NULL;
}


int ObjToArgvA(Tcl_Interp *interp, Tcl_Obj *objP, char **argv, int argc, int *argcP)
{
    int       objc, i;
    Tcl_Obj **objv;

    if (Tcl_ListObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if ((objc+1) > argc) {
        return TwapiReturnErrorEx(interp,
                                  TWAPI_INTERNAL_LIMIT,
                                  Tcl_ObjPrintf("Number of strings (%d) in list exceeds size of argument array.", objc));
    }

    for (i = 0; i < objc; ++i)
        argv[i] = Tcl_GetString(objv[i]);
    argv[i] = NULL;
    *argcP = objc;
    return TCL_OK;
}

int ObjToArgvW(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *argv, int argc, int *argcP)
{
    int       objc, i;
    Tcl_Obj **objv;

    if (Tcl_ListObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if ((objc+1) > argc) {
        return TwapiReturnErrorEx(interp,
                                  TWAPI_INTERNAL_LIMIT,
                                  Tcl_ObjPrintf("Number of strings (%d) in list exceeds size of argument array.", objc));
    }

    for (i = 0; i < objc; ++i)
        argv[i] = Tcl_GetUnicode(objv[i]);
    argv[i] = NULL;
    *argcP = objc;
    return TCL_OK;
}

Tcl_Obj *ObjFromOpaque(void *pv, char *name)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromDWORD_PTR(pv);
    objs[1] = Tcl_NewStringObj(name ? name : "void*", -1);
    return Tcl_NewListObj(2, objs);
}

int ObjToOpaque(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, char *name)
{
    Tcl_Obj **objsP;
    int       nobj, val;
    DWORD_PTR dwp;

    if (Tcl_ListObjGetElements(interp, obj, &nobj, &objsP) != TCL_OK)
        return TCL_ERROR;
    if (nobj != 2) {
        /* For backward compat with SWIG based script, we accept NULL
           as a valid pointer of any type and for convenience 0 as well */
        if (nobj == 1 &&
            (lstrcmpA(Tcl_GetString(obj), "NULL") == 0 ||
             (Tcl_GetIntFromObj(interp, obj, &val) == TCL_OK && val == 0))) {
            *pvP = 0;
            return TCL_OK;
        }

        if (interp) {
            Tcl_ResetResult(interp);
            Tcl_AppendResult(interp, "Invalid pointer or opaque value: '",
                             Tcl_GetString(obj), "'.", NULL);
        }
        return TCL_ERROR;
    }

    /* If a type name is specified, see that it matches. Else any type ok */
    if (name) {
        char *s = Tcl_GetString(objsP[1]);
        if (! STREQ(s, name)) {
            if (interp) {
                Tcl_AppendResult(interp, "Unexpected type '", s, "', expected '",
                                 name, "'.", NULL);
                return TCL_ERROR;
            }
        }
    }
    
    if (ObjToDWORD_PTR(NULL, objsP[0], &dwp) != TCL_OK) {
        if (interp)
            Tcl_AppendResult(interp, "Invalid pointer or opaque value '",
                             Tcl_GetString(objsP[0]), "'.", NULL);
        return TCL_ERROR;
    }
    *pvP = (void*) dwp;
    return TCL_OK;
}

/* Converts a Tcl_Obj to a pointer of any of the specified types */
int ObjToOpaqueMulti(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, int ntypes, char **types)
{
    int i;
    if (ntypes == 0 || types == NULL)
        return ObjToOpaque(interp, obj, pvP, NULL);

    for (i = 0; i < ntypes; ++i) {
        if (ObjToOpaque(interp, obj, pvP, types[i]) == TCL_OK) {
            Tcl_ResetResult(interp); /* Clean up errors from any prev type attempts */
            return TCL_OK;
        }
    }

    return TCL_ERROR;
}

int ObjToIDispatch(Tcl_Interp *interp, Tcl_Obj *obj, IDispatch **dispP) 
{
    /*
     * Either IDispatchEx or IDispatch is acceptable. We try in that
     * order so error message, if any will refer to IDispatch.
     */
    if (ObjToOpaque(NULL, obj, dispP, "IDispatchEx") == TCL_OK)
        return TCL_OK;
    return ObjToOpaque(interp, obj, dispP, "IDispatch");
}

Tcl_Obj *ObjFromSYSTEM_POWER_STATUS(SYSTEM_POWER_STATUS *spsP)
{
    Tcl_Obj *objv[6];
    objv[0] = Tcl_NewIntObj(spsP->ACLineStatus);
    objv[1] = Tcl_NewIntObj(spsP->BatteryFlag);
    objv[2] = Tcl_NewIntObj(spsP->BatteryLifePercent);
    objv[3] = Tcl_NewIntObj(spsP->Reserved1);
    objv[4] = Tcl_NewIntObj(spsP->BatteryLifeTime);
    objv[5] = Tcl_NewIntObj(spsP->BatteryFullLifeTime);
    return Tcl_NewListObj(6, objv);
}


Tcl_Obj *TwapiUtf8ObjFromUnicode(CONST WCHAR *wsP, int nchars)
{
    Tcl_DString ds;
    Tcl_Obj *objP;
    int nbytes;

    /*
     * Note - not using Tcl_WinTCharToUtf because there is no way
     * of telling if the Tcl core was compiled with _UNICODE defined
     * or not
     */

    /* Note WideChar... does not like 0 length strings */
    if (wsP == NULL || nchars == 0)
        return Tcl_NewObj();

    Tcl_DStringInit(&ds);

    nbytes = WideCharToMultiByte(
        CP_UTF8, /* CodePag */
        0,       /* dwFlags */
        wsP,     /* lpWideCharStr */
        nchars < 0 ? -1 : nchars, /* cchWideChar */
        NULL,    /* lpMultiByteStr */
        0,       /* cbMultiByte */
        NULL,    /* lpDefaultChar */
        NULL     /* lpUsedDefaultChar */
        );
    
    if (nbytes == 0) {
        Tcl_Panic("WideCharToMultiByte returned 0, with error code %d", GetLastError());
    }

    Tcl_DStringSetLength(&ds, nbytes);

    nbytes = WideCharToMultiByte(
        CP_UTF8, /* CodePag */
        0,       /* dwFlags */
        wsP,     /* lpWideCharStr */
        nchars < 0 ? -1 : nchars, /* cchWideChar */
        Tcl_DStringValue(&ds),    /* lpMultiByteStr */
        Tcl_DStringLength(&ds),   /* cbMultiByte */
        NULL,    /* lpDefaultChar */
        NULL     /* lpUsedDefaultChar */
        );
    
    if (nbytes == 0) {
        Tcl_Panic("WideCharToMultiByte returned 0, with error code %d", GetLastError());
    }

    /*
     * Note WideCharToMultiByte does not explicitly terminate with \0
     * if nchars was specifically specified
     */
    if (nchars < 0)
        --nbytes;                /* Exclude terminating \0 */
    objP = Tcl_NewStringObj(Tcl_DStringValue(&ds), nbytes);

    Tcl_DStringFree(&ds);
    return objP;
}

Tcl_Obj *ObjFromTIME_ZONE_INFORMATION(const TIME_ZONE_INFORMATION *tzP)
{
    /*
     * We mostly just pass this around, so just keep as binary structure
     */
    return Tcl_NewByteArrayObj((unsigned char *)tzP, sizeof(*tzP));
}

TCL_RESULT ObjToTIME_ZONE_INFORMATION(Tcl_Interp *interp,
                                      Tcl_Obj *tzObj,
                                      TIME_ZONE_INFORMATION *tzP)
{
    unsigned char *p;
    int len;

    p = Tcl_GetByteArrayFromObj(tzObj, &len);
    if (len != sizeof(*tzP)) {
        return TwapiReturnErrorEx(interp,
                                  TWAPI_INVALID_ARGS,
                                  Tcl_ObjPrintf("Invalid TIME_ZONE_INFORMATION size %d", len));
    }

    *tzP = *((TIME_ZONE_INFORMATION *)p);
    return TCL_OK;
}

Tcl_Obj *ObjFromULONGHex(ULONG val)
{
    return Tcl_ObjPrintf("0x%8.8x", val);
}

Tcl_Obj *ObjFromULONGLONGHex(ULONGLONG ull)
{
    Tcl_Obj *objP;
    /* Unfortunately, Tcl_Objprintf does not handle 64 bits currently */
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
    Tcl_Obj *wideObj;
    wideobj = Tcl_NewWideIntObj((Tcl_WideInt) ull);
    objP = Tcl_Format(NULL, "0x%16.16lx", 1, &wideobj);
    Tcl_DecrRefCount(wideobj);
#else
    char buf[40];
    _snprintf(buf, sizeof(buf), "0x%16.16I64x", ull);
    objP = Tcl_NewStringObj(buf, -1);
#endif
    return objP;
}


Tcl_Obj *ObjFromULONGLONG(ULONGLONG ull)
{
    /*
     * Unsigned 64-bit ints with the high bit set will not fit in Tcl_WideInt.
     * We need to convert to a bignum from a hex string.
     */

    if (ull & 0x8000000000000000) {
        Tcl_Obj *objP;
        mp_int mpi;
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
        Tcl_Obj *mpobj;
        objP = Tcl_NewWideIntObj((Tcl_WideInt) ull);
        mpobj = Tcl_Format(NULL, "0x%lx", 1, &objP);
        Tcl_DecrRefCount(objP);
        objP = mpobj;
#else
        char buf[40];
        _snprintf(buf, sizeof(buf), "%I64u", ull);
        objP = Tcl_NewStringObj(buf, -1);
#endif
        /* Force to bignum because COM interface sometimes needs to check type*/
        if (Tcl_GetBignumFromObj(NULL, objP, &mpi) == TCL_ERROR)
            return objP;
#if defined(TWAPI_REPLACE_CRT) || defined(TWAPI_MINIMIZE_CRT)
        Tcl_InvalidateStringRep(objP); /* So we get a decimal string rep */
#endif
        TclBN_mp_clear(&mpi);
        return objP;
    } else {
        return Tcl_NewWideIntObj((Tcl_WideInt) ull);
    }
}

/* Given a IP address as a DWORD, returns a Tcl string */
Tcl_Obj *IPAddrObjFromDWORD(DWORD addr)
{
    struct in_addr inaddr;
    inaddr.S_un.S_addr = addr;
    return Tcl_NewStringObj(inet_ntoa(inaddr), -1);
}

/* Given a string, return the IP address */
int IPAddrObjToDWORD(Tcl_Interp *interp, Tcl_Obj *objP, DWORD *addrP)
{
    DWORD addr;
    char *p = Tcl_GetString(objP);
    if ((addr = inet_addr(p)) == INADDR_NONE) {
        /* Bad format or 255.255.255.255 */
        if (! STREQ("255.255.255.255", p)) {
            if (interp) {
                Tcl_AppendResult(interp, "Invalid IP address format: ", p, NULL);
            }
            return TCL_ERROR;
        }
        /* Fine, addr contains 0xffffffff */
    }
    *addrP = addr;
    return TCL_OK;
}


/* Given a IP_ADDR_STRING list, return a Tcl_Obj */
Tcl_Obj *ObjFromIP_ADDR_STRING (
    Tcl_Interp *interp, const IP_ADDR_STRING *ipaddrstrP
)
{
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);
    while (ipaddrstrP) {
        Tcl_Obj *objv[3];

        if (ipaddrstrP->IpAddress.String[0]) {
            objv[0] = Tcl_NewStringObj(ipaddrstrP->IpAddress.String, -1);
            objv[1] = Tcl_NewStringObj(ipaddrstrP->IpMask.String, -1);
            objv[2] = Tcl_NewIntObj(ipaddrstrP->Context);
            Tcl_ListObjAppendElement(interp, resultObj,
                                     Tcl_NewListObj(3, objv));
        }

        ipaddrstrP = ipaddrstrP->Next;
    }

    return resultObj;
}


/* Note - port is not returned - only address */
Tcl_Obj *ObjFromSOCKADDR_address(SOCKADDR *saP)
{
    char buf[50];
    DWORD bufsz = ARRAYSIZE(buf);
    
    if (WSAAddressToStringA(saP,
                            ((SOCKADDR_IN6 *)saP)->sin6_family == AF_INET6 ? sizeof(SOCKADDR_IN6) : sizeof(SOCKADDR_IN),
                            NULL,
                            buf,
                            &bufsz) == 0) {
        if (bufsz && buf[bufsz-1] == 0)
            --bufsz;        /* Terminating \0 */
        return Tcl_NewStringObj(buf, bufsz);
    }
    /* Error already set */
    return NULL;
}

/* Can return NULL on error */
Tcl_Obj *ObjFromSOCKADDR(SOCKADDR *saP)
{
    short save_port;
    Tcl_Obj *objv[2];

    /* Stash port as 0 so does not show in address string */
    if (((SOCKADDR_IN6 *)saP)->sin6_family == AF_INET6) {
        save_port = ((SOCKADDR_IN6 *)saP)->sin6_port;
        ((SOCKADDR_IN6 *)saP)->sin6_port = 0;
    } else {
        save_port = ((SOCKADDR_IN *)saP)->sin_port;
        ((SOCKADDR_IN *)saP)->sin_port = 0;
    }
    
    objv[0] = ObjFromSOCKADDR_address(saP);
    if (objv[0] == NULL)
        return NULL;

    objv[1] = Tcl_NewIntObj((WORD)(ntohs(save_port)));

    if (((SOCKADDR_IN6 *)saP)->sin6_family == AF_INET6) {
        ((SOCKADDR_IN6 *)saP)->sin6_port = save_port;
    } else {
        ((SOCKADDR_IN *)saP)->sin_port = save_port;
    }

    return Tcl_NewListObj(2, objv);
}


Tcl_Obj *ObjFromIPv6Addr(const char *addrP, DWORD scope_id)
{
    SOCKADDR_IN6 si;

    si.sin6_family = AF_INET6;
    si.sin6_port = 0;
    si.sin6_flowinfo = 0;
    CopyMemory(si.sin6_addr.u.Byte, addrP, 16);
    si.sin6_scope_id = scope_id;
    return ObjFromSOCKADDR_address((SOCKADDR *)&si);
}
