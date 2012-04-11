/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "tclTomMath.h"

/*
 * Struct for mapping VARTYPE values to strings.
 * We search linearly so order based on most likely types. - TBD make hash
 * Only the basic types are covered in this table. The code itself
 * handles the special/complex cases.
 */
struct vt_token_pair {
    VARTYPE vt;
    char   *tok;
};
static struct vt_token_pair vt_base_tokens[] = {
    {VT_BOOL, "bool"},
    {VT_I2, "i2"},
    {VT_I4, "i4"},
    {VT_PTR, "ptr"},
    {VT_R4, "r4"},
    {VT_R8, "r8"},
    {VT_CY, "cy"},
    {VT_DATE, "date"},
    {VT_BSTR, "bstr"},
    {VT_DISPATCH, "idispatch"},
    {VT_ERROR, "error"},
    {VT_VARIANT, "variant"},
    {VT_UNKNOWN, "iunknown"},
    {VT_UI1, "ui1"},
    {VT_DECIMAL, "decimal"},
    {VT_I1, "i1"},
    {VT_UI2, "ui2"},
    {VT_UI4, "ui4"},
    {VT_I8, "i8"},
    {VT_UI8, "ui8"},
    {VT_INT, "int"},
    {VT_UINT, "uint"},
    {VT_HRESULT, "hresult"},
    {VT_VOID, "void"},
    {VT_LPSTR, "lpstr"},
    {VT_LPWSTR, "lpwstr"},
    {VT_RECORD, "record"},
    {VT_USERDEFINED, "userdefined"}
};

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
        Tcl_Obj *objP = STRING_LITERAL_OBJ("true");
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
        TwapiSetObjResult(interp, ObjFromString(objv[1]->typePtr->name));
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

    typename = ObjToString(objv[1]);

    if (*typename == '\0') {
        /* No type, keep as is */
        return TwapiSetObjResult(interp, objv[2]);
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
        objP = ObjFromString(i ? "true" : "false");
        Tcl_GetBooleanFromObj(NULL, objP, &i);
        return TwapiSetObjResult(interp, objP);
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

    return TwapiSetObjResult(interp, objP);
}

/* Call to set static result */
void TwapiSetStaticResult(Tcl_Interp *interp, CONST char s[])
{
    Tcl_SetResult(interp, (char *) s, TCL_STATIC);
}

TCL_RESULT TwapiSetObjResult(Tcl_Interp *interp, Tcl_Obj *objP)
{
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
        resultObj = ObjFromBoolean(resultP->value.bval);
        break;

    case TRT_EXCEPTION_ON_FALSE:
    case TRT_NONZERO_RESULT:
        /* If 0, generate exception */
        if (! resultP->value.ival)
            return TwapiReturnSystemError(interp);

        if (resultP->type == TRT_NONZERO_RESULT)
            resultObj = ObjFromLong(resultP->value.ival);
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
        resultObj = ObjFromLong(resultP->value.ival);
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
            resultObj = ObjFromStringN(resultP->value.chars.str,
                                         resultP->value.chars.len);
        break;

    case TRT_BINARY:
        resultObj = ObjFromByteArray(resultP->value.binary.p,
                                        resultP->value.binary.len);
        break;

    case TRT_OBJ:
        resultObj = resultP->value.obj;
        break;

    case TRT_OBJV:
        resultObj = ObjNewList(resultP->value.objv.nobj, resultP->value.objv.objPP);
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
            TwapiSetStaticResult(interp, "Internal error: TwapiSetResult - inconsistent nesting of case statements");
            return TCL_ERROR;
        }
        resultObj = ObjFromOpaque(resultP->value.hval, typenameP);
        break;

    case TRT_LPVOID:
        resultObj = ObjFromOpaque(resultP->value.hval, "void*");
        break;

    case TRT_LONG:
        resultObj = ObjFromLong(resultP->value.ival);
        break;

    case TRT_DWORD:
        resultObj = ObjFromWideInt((Tcl_WideInt) resultP->value.uval);
        break;
        
    case TRT_WIDE:
        resultObj = ObjFromWideInt(resultP->value.wide);
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

    case TRT_NONNULL:
        if (resultP->value.nonnull.p == NULL)
            return TwapiReturnSystemError(interp);
        resultObj = ObjFromOpaque(resultP->value.nonnull.p,
                                  resultP->value.nonnull.name);
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
        TwapiSetStaticResult(interp, "Unknown TwapiResultType type code passed to TwapiSetResult");
        return TCL_ERROR;
    }

    TwapiClearResult(resultP);  /* Clear out resources */

    if (resultObj)
        TwapiSetObjResult(interp, resultObj);


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
        s = ObjToStringN(objv[i], &len);
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
        LPWSTR p = ObjToUnicodeN(objP, &len);
        if (len > 0)
            return p;
    }
    return NULL;
}

LPWSTR ObjToLPWSTR_WITH_NULL(Tcl_Obj *objP)
{
    if (objP) {
        LPWSTR s = ObjToUnicode(objP);
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
        wcharP = ObjToUnicodeN(objP, &len);
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
    return bstr ?
        ObjFromUnicodeN(bstr, SysStringLen(bstr))
        : ObjFromEmptyString();
}

Tcl_Obj *ObjFromStringLimited(const char *strP, int max, int *remainP)
{
    int len;

    if (max < 0) {
        if (remainP)
            *remainP = 0;
        return ObjFromString(strP);
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

    return ObjFromStringN(strP, len);
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
    if (ObjToInt(interp, obj, &i) != TCL_OK)
        return TCL_ERROR;

    if (i < low || i > high) {
        if (interp) {
            TwapiSetObjResult(interp,
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

    objv[0] = ObjFromInt(timeP->wYear);
    objv[1] = ObjFromInt(timeP->wMonth);
    objv[2] = ObjFromInt(timeP->wDay);
    objv[3] = ObjFromInt(timeP->wHour);
    objv[4] = ObjFromInt(timeP->wMinute);
    objv[5] = ObjFromInt(timeP->wSecond);
    objv[6] = ObjFromInt(timeP->wMilliseconds);

    return ObjNewList(7, objv);
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

    if (ObjGetElements(interp, timeObj, &objc, &objv) != TCL_OK)
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
    return ObjFromWideInt(*(Tcl_WideInt *)cyP);
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
        return ObjFromEmptyString();
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
    hr = VarDecFromStr(ObjToUnicode(obj), 0, 0, decP);
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
    return ObjFromByteArray((unsigned char *)pidl,
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

    idsP = (LPITEMIDLIST) ObjToByteArray(objP, &numbytes);
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
                TwapiSetStaticResult(interp, "Invalid item id list format");
                return TCL_ERROR;
            }
            numbytes -= itemlen;
            p += itemlen;
        }
    }

    *idsPP = CoTaskMemAlloc(numbytes);
    if (*idsPP == NULL) {
        if (interp)
            TwapiSetStaticResult(interp, "CoTaskMemAlloc failed in SHChangeNotify");
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
        return ObjFromEmptyString("", 0);

    obj = ObjFromUnicode(str);
    return obj;
}

int ObjToGUID(Tcl_Interp *interp, Tcl_Obj *objP, GUID *guidP)
{
    HRESULT hr;
    WCHAR *wsP;
    if (objP) {
        wsP = ObjToUnicode(objP);

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
    objP = ObjFromString(uuidStr);
    RpcStringFree(&uuidStr);
    return objP;
}

int ObjToUUID(Tcl_Interp *interp, Tcl_Obj *objP, UUID *uuidP)
{
    /* NOTE UUID and GUID have same binary format but are formatted
       differently based on the component.  We accept both forms here */

    if (objP) {
        RPC_STATUS status = UuidFromStringA(ObjToString(objP), uuidP);
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
    lsauniP->Buffer = ObjToUnicodeN(objP, &nchars);
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

    *objPP = ObjFromString(strP);
    LocalFree(strP);
    return TCL_OK;
}

/* Like ObjFromSID but returns empty object on error */
Tcl_Obj *ObjFromSIDNoFail(SID *sidP)
{
    Tcl_Obj *objP;
    if (sidP == NULL || ObjFromSID(NULL, sidP, &objP) != TCL_OK)
        return ObjFromEmptyString();
    else
        return objP;
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
        src = ObjToUnicodeN(objPtr, &len);
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
    Tcl_Obj *listPtr = ObjNewList(0, NULL);
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

        ObjAppendElement(NULL, listPtr, ObjFromUnicodeN(s, (int) (lpcw-s)));
        ++lpcw;            /* Point beyond this string, possibly beyond end */
    }

    return listPtr;
}

int ObjToWord(Tcl_Interp *interp, Tcl_Obj *obj, WORD *wordP)
{
    long lval;
    if (ObjToLong(interp, obj, &lval) != TCL_OK)
        return TCL_ERROR;
    if (lval & 0xffff0000) {
        if (interp)
            TwapiSetStaticResult(interp, "Integer value must be less than 65536");
        return TCL_ERROR;
    }
    *wordP = (WORD) lval;
    return TCL_OK;
}

int ObjToRECT (Tcl_Interp *interp, Tcl_Obj *obj, RECT *rectP)
{
    Tcl_Obj **objv;
    int       objc;

    if (ObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 4) {
        TwapiSetStaticResult(interp, "Invalid RECT format.");
        return TCL_ERROR;
    }
    if ((ObjToLong(interp, objv[0], &rectP->left) != TCL_OK) ||
        (ObjToLong(interp, objv[1], &rectP->top) != TCL_OK) ||
        (ObjToLong(interp, objv[2], &rectP->right) != TCL_OK) ||
        (ObjToLong(interp, objv[3], &rectP->bottom) != TCL_OK)) {
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

    objv[0] = ObjFromLong(rectP->left);
    objv[1] = ObjFromLong(rectP->top);
    objv[2] = ObjFromLong(rectP->right);
    objv[3] = ObjFromLong(rectP->bottom);
    return ObjNewList(4, objv);
}

/* Return a Tcl Obj from a POINT structure */
Tcl_Obj *ObjFromPOINT(POINT *ptP)
{
    Tcl_Obj *objv[2];

    objv[0] = ObjFromLong(ptP->x);
    objv[1] = ObjFromLong(ptP->y);
    return ObjNewList(2, objv);
}

/* Convert a Tcl_Obj to a POINT */
int ObjToPOINT (Tcl_Interp *interp, Tcl_Obj *obj, POINT *ptP)
{
    Tcl_Obj **objv;
    int       objc;

    if (ObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 2) {
        TwapiSetStaticResult(interp, "Invalid POINT format.");
        return TCL_ERROR;
    }
    if ((ObjToLong(interp, objv[0], &ptP->x) != TCL_OK) ||
        (ObjToLong(interp, objv[1], &ptP->y) != TCL_OK)) {
        return TCL_ERROR;
    }

    return TCL_OK;
}

/* Return a Tcl Obj from a POINT structure */
Tcl_Obj *ObjFromPOINTS(POINTS *ptP)
{
    Tcl_Obj *objv[2];

    objv[0] = ObjFromInt(ptP->x);
    objv[1] = ObjFromInt(ptP->y);

    return ObjNewList(2, objv);
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
    char *strP = ObjToStringN(objP, &len);

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
        TwapiSetStaticResult(interp, "Invalid LUID format: ");
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
        objv[1] = ObjFromLong(*(int *)bufP);
        break;

    case REG_QWORD:
        if (count != 8)
            goto badformat;
        typestr = regtype == REG_QWORD ? "qword" : "qword_be";
        objv[1] = ObjFromWideInt(*(Tcl_WideInt *)bufP);
        break;

    case REG_MULTI_SZ:
        typestr = "multi_sz";
        objv[1] = ObjFromMultiSz((LPCWSTR) bufP, count/2);
        break;

    case REG_BINARY:
        typestr = "binary";
        // FALLTHRU
    default:
        objv[1] = ObjFromByteArray(bufP, count);
        break;
    }

    if (typestr)
        objv[0] = ObjFromString(typestr);
    else
        objv[0] = ObjFromLong(regtype);
    return ObjNewList(2, objv);

badformat:
    TwapiSetStaticResult(interp, "Badly formatted registry value");
    return NULL;
}


int ObjToArgvA(Tcl_Interp *interp, Tcl_Obj *objP, char **argv, int argc, int *argcP)
{
    int       objc, i;
    Tcl_Obj **objv;

    if (ObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if ((objc+1) > argc) {
        return TwapiReturnErrorEx(interp,
                                  TWAPI_INTERNAL_LIMIT,
                                  Tcl_ObjPrintf("Number of strings (%d) in list exceeds size of argument array.", objc));
    }

    for (i = 0; i < objc; ++i)
        argv[i] = ObjToString(objv[i]);
    argv[i] = NULL;
    *argcP = objc;
    return TCL_OK;
}

int ObjToArgvW(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *argv, int argc, int *argcP)
{
    int       objc, i;
    Tcl_Obj **objv;

    if (ObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if ((objc+1) > argc) {
        return TwapiReturnErrorEx(interp,
                                  TWAPI_INTERNAL_LIMIT,
                                  Tcl_ObjPrintf("Number of strings (%d) in list exceeds size of argument array.", objc));
    }

    for (i = 0; i < objc; ++i)
        argv[i] = ObjToUnicode(objv[i]);
    argv[i] = NULL;
    *argcP = objc;
    return TCL_OK;
}

Tcl_Obj *ObjFromOpaque(void *pv, char *name)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromDWORD_PTR(pv);
    objs[1] = ObjFromString(name ? name : "void*");
    return ObjNewList(2, objs);
}

int ObjToOpaque(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, char *name)
{
    Tcl_Obj **objsP;
    int       nobj, val;
    DWORD_PTR dwp;

    if (ObjGetElements(interp, obj, &nobj, &objsP) != TCL_OK)
        return TCL_ERROR;
    if (nobj != 2) {
        /* For backward compat with SWIG based script, we accept NULL
           as a valid pointer of any type and for convenience 0 as well */
        if (nobj == 1 &&
            (lstrcmpA(ObjToString(obj), "NULL") == 0 ||
             (ObjToInt(interp, obj, &val) == TCL_OK && val == 0))) {
            *pvP = 0;
            return TCL_OK;
        }

        if (interp) {
            Tcl_ResetResult(interp);
            Tcl_AppendResult(interp, "Invalid pointer or opaque value: '",
                             ObjToString(obj), "'.", NULL);
        }
        return TCL_ERROR;
    }

    /* If a type name is specified, see that it matches. Else any type ok */
    if (name) {
        char *s = ObjToString(objsP[1]);
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
                             ObjToString(objsP[0]), "'.", NULL);
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
    objv[0] = ObjFromInt(spsP->ACLineStatus);
    objv[1] = ObjFromInt(spsP->BatteryFlag);
    objv[2] = ObjFromInt(spsP->BatteryLifePercent);
    objv[3] = ObjFromInt(spsP->Reserved1);
    objv[4] = ObjFromDWORD(spsP->BatteryLifeTime);
    objv[5] = ObjFromDWORD(spsP->BatteryFullLifeTime);
    return ObjNewList(6, objv);
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
        return ObjFromEmptyString();

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
    objP = ObjFromStringN(Tcl_DStringValue(&ds), nbytes);

    Tcl_DStringFree(&ds);
    return objP;
}

Tcl_Obj *ObjFromTIME_ZONE_INFORMATION(const TIME_ZONE_INFORMATION *tzP)
{
    /*
     * We mostly just pass this around, so just keep as binary structure
     */
    return ObjFromByteArray((unsigned char *)tzP, sizeof(*tzP));
}

TCL_RESULT ObjToTIME_ZONE_INFORMATION(Tcl_Interp *interp,
                                      Tcl_Obj *tzObj,
                                      TIME_ZONE_INFORMATION *tzP)
{
    unsigned char *p;
    int len;

    p = ObjToByteArray(tzObj, &len);
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
    wideobj = ObjFromWideInt((Tcl_WideInt) ull);
    objP = Tcl_Format(NULL, "0x%16.16lx", 1, &wideobj);
    Tcl_DecrRefCount(wideobj);
#else
    char buf[40];
    _snprintf(buf, sizeof(buf), "0x%16.16I64x", ull);
    objP = ObjFromString(buf);
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
        objP = ObjFromWideInt((Tcl_WideInt) ull);
        mpobj = Tcl_Format(NULL, "0x%lx", 1, &objP);
        Tcl_DecrRefCount(objP);
        objP = mpobj;
#else
        char buf[40];
        _snprintf(buf, sizeof(buf), "%I64u", ull);
        objP = ObjFromString(buf);
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
        return ObjFromWideInt((Tcl_WideInt) ull);
    }
}

/* Given a IP address as a DWORD, returns a Tcl string */
Tcl_Obj *IPAddrObjFromDWORD(DWORD addr)
{
    struct in_addr inaddr;
    inaddr.S_un.S_addr = addr;
    return ObjFromString(inet_ntoa(inaddr));
}

/* Given a string, return the IP address */
int IPAddrObjToDWORD(Tcl_Interp *interp, Tcl_Obj *objP, DWORD *addrP)
{
    DWORD addr;
    char *p = ObjToString(objP);
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
    Tcl_Obj *resultObj = ObjNewList(0, NULL);
    while (ipaddrstrP) {
        Tcl_Obj *objv[3];

        if (ipaddrstrP->IpAddress.String[0]) {
            objv[0] = ObjFromString(ipaddrstrP->IpAddress.String);
            objv[1] = ObjFromString(ipaddrstrP->IpMask.String);
            objv[2] = ObjFromDWORD(ipaddrstrP->Context);
            ObjAppendElement(interp, resultObj,
                                     ObjNewList(3, objv));
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
        return ObjFromStringN(buf, bufsz);
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

    objv[1] = ObjFromInt((WORD)(ntohs(save_port)));

    if (((SOCKADDR_IN6 *)saP)->sin6_family == AF_INET6) {
        ((SOCKADDR_IN6 *)saP)->sin6_port = save_port;
    } else {
        ((SOCKADDR_IN *)saP)->sin_port = save_port;
    }

    return ObjNewList(2, objv);
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

static void TwapiInvalidVariantTypeMessage(Tcl_Interp *interp, VARTYPE vt)
{
    if (interp) {
        (void) TwapiSetObjResult(interp,
                         Tcl_ObjPrintf("Invalid or unsupported VARTYPE (%d)",
                                       vt));
    }
}

static int LookupBaseVT(Tcl_Interp *interp, VARTYPE vt, const char **tokP)
{
    int i;
    for (i=0; i < ARRAYSIZE(vt_base_tokens); ++i) {
        if (vt_base_tokens[i].vt == vt) {
            if (tokP)
                *tokP = vt_base_tokens[i].tok;
            return TCL_OK;
        }
    }

    TwapiInvalidVariantTypeMessage(interp, vt);
    return TCL_ERROR;
}

static int LookupBaseVTToken(Tcl_Interp *interp, const char *tok, VARTYPE *vtP)
{
    int i;
    if (tok != NULL) {
        for (i=0; i < ARRAYSIZE(vt_base_tokens); ++i) {
            if (STREQ(vt_base_tokens[i].tok, tok)) {
                if (vtP)
                    *vtP = vt_base_tokens[i].vt;
                return TCL_OK;
            }
        }
    }
    if (interp) {
        Tcl_Obj *objP; 
        objP = STRING_LITERAL_OBJ("Invalid or unsupported VARTYPE token: ");
        Tcl_AppendToObj(objP, tok ? tok : "<null pointer>", -1);
        (void)TwapiSetObjResult(interp, objP);
    }

    return TCL_ERROR;
}

/* Convert a VT string rep to corresponding integer */
int ObjToVT(Tcl_Interp *interp, Tcl_Obj *obj, VARTYPE *vtP)
{
    int i;
    Tcl_Obj **objv;
    int       objc;
    VARTYPE   vt;

    /* The VT may be take one of the following forms:
     *    - integer
     *    - symbol
     *    - list {ptr VT}
     *    - list {userdefined VT}
     */
    if (ObjToInt(NULL, obj, &i) == TCL_OK) {
        *vtP = (VARTYPE) i;
        return TCL_OK;
    } else if (LookupBaseVTToken(interp, ObjToString(obj), vtP) == TCL_OK) {
        return TCL_OK;
    }

    /*
     * See if it's a list. Note interp contains an error msg at this point
     */

    if (ObjGetElements(NULL, obj, &objc, &objv) != TCL_OK ||
        objc < 2) {
        return TCL_ERROR;
    }
    if (ObjToInt(NULL, objv[0], &i) == TCL_OK) {
        vt = (VARTYPE) i;
    } else if (LookupBaseVTToken(NULL, ObjToString(objv[0]), &vt) != TCL_OK) {
        return TCL_ERROR;
    }

    /* vt must be either pointer, array or UDT in the list case */
    if (vt == VT_PTR || vt == VT_SAFEARRAY || vt == VT_USERDEFINED) {
        *vtP = vt;
        Tcl_ResetResult(interp); // Get rid of old error message.
        return TCL_OK;
    }
    else
        return TCL_ERROR;
}

/*
 * Return a Tcl object that is a list
 * {"safearray" dimensionlist VT_xxx valuelist}.
 * dimensionlist is a flat list of lowbound, upperbound pairs, one
 * for each dimension.
 * If VT_xxx is not recognized, valuelist is missing
 * If there is no vartype information, VT_XXX is also missing
 * Never returns NULL.
 */
static Tcl_Obj *ObjFromSAFEARRAY(SAFEARRAY *arrP)
{
    Tcl_Obj *objv[3];           /* "safearray|vt", dimensions,  value */
    int      objc;
    long     i;
    VARTYPE  vt;
    HRESULT  hr;
    long     num_elems;
    void     *valP;
#define GETVAL(index_, type_) (((type_ *)valP)[index_])

    /* We require the safearray to have a type associated */
    objc = 1;
    if (SafeArrayGetVartype(arrP, &vt) == S_OK) {
        objv[0] = ObjFromInt(vt|VT_ARRAY);
    } else {
        objv[0] = ObjFromInt(VT_ARRAY);
        goto alldone;
    }

    hr = SafeArrayAccessData(arrP, &valP);
    if (FAILED(hr))
        goto alldone;

    objv[1] = ObjNewList(0, NULL);
    num_elems = 1;
    for (i = 0; i < arrP->cDims; ++i) {
        ObjAppendElement(NULL, objv[1], ObjFromLong(arrP->rgsabound[i].lLbound));
        ObjAppendElement(NULL, objv[1], ObjFromLong(arrP->rgsabound[i].cElements));
        num_elems *= arrP->rgsabound[i].cElements;
    }

    /* TBD - it might be more efficient to allocate an array and then
       use Tcl_NewListObj to create from there instead of calling
       Tcl_ListObjAppend for each element
    */

    objv[2] = ObjNewList(0, NULL); /* Value object */
    objc = 3;

    switch (vt) {
    case VT_EMPTY:
    case VT_NULL:
        break;

    case VT_I2: {
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2], ObjFromInt(GETVAL(i,short)));
        }
        break;
    }

    case VT_INT: /* FALLTHROUGH */
    case VT_I4:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2], ObjFromLong(GETVAL(i,long)));
        }
        break;

    case VT_R4:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2],
                                     Tcl_NewDoubleObj(GETVAL(i,float)));
        }
        break;

    case VT_R8:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2],
                                     Tcl_NewDoubleObj(GETVAL(i,double)));
        }
        break;

    case VT_CY:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2], ObjFromCY(&(((CY *)valP)[i])) );
        }
        break;

    case VT_DATE:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(
                NULL, objv[2],
                Tcl_NewDoubleObj(GETVAL(i,double)));
        }
        break;

    case VT_BSTR:
        for (i = 0; i < num_elems; ++i) {
            BSTR bstr = GETVAL(i,BSTR);
            ObjAppendElement(
                NULL, objv[2],
                ObjFromUnicodeN(bstr, SysStringLen(bstr))
                );
        }
        break;

    case VT_DISPATCH:
        for (i = 0; i < num_elems; ++i) {
            IDispatch *idispP = GETVAL(i,IDispatch *);
            ObjAppendElement(
                NULL, objv[2],
                ObjFromIDispatch(idispP)
                );
        }
        break;

    case VT_ERROR:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2],
                                     ObjFromInt(GETVAL(i,SCODE)));
        }
        break;

    case VT_BOOL:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(
                NULL, objv[2],
                ObjFromBoolean(GETVAL(i,VARIANT_BOOL))
                );
        }
        break;

    case VT_VARIANT:
        for (i = 0; i < num_elems; ++i) {
            VARIANT *varP = &((( VARIANT *)valP)[i]);
            ObjAppendElement(
                NULL, objv[2],
                ObjFromVARIANT(varP, 0));
        }
        break;

    case VT_DECIMAL:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(
                NULL, objv[2],
                ObjFromDECIMAL(&((( DECIMAL *)valP)[i]))
                );
        }
        break;

    case VT_UNKNOWN:
        for (i = 0; i < num_elems; ++i) {
            IUnknown *idispP = GETVAL(i, IUnknown *);
            ObjAppendElement(
                NULL, objv[2],
                ObjFromIUnknown(idispP));
        }
        break;

    case VT_I1:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2],
                                     ObjFromInt(GETVAL(i,char)));
        }
        break;

    case VT_UI1:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2],
                                     ObjFromInt(GETVAL(i,unsigned char)));
        }
        break;

    case VT_UI2:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2],
                                     ObjFromInt(GETVAL(i,unsigned short)));
        }
        break;

    case VT_UINT: /* FALLTHROUGH */
    case VT_UI4:
        for (i = 0; i < num_elems; ++i) {
            unsigned long ulval = GETVAL(i, unsigned long);
            /* store as wide integer if it does not fit in signed 32 bits */
            ObjAppendElement(NULL, objv[2], ObjFromDWORD(ulval));
        }
        break;

    case VT_I8: /* FALLTHRU */
    case VT_UI8:
        for (i = 0; i < num_elems; ++i) {
            ObjAppendElement(NULL, objv[2],
                                     ObjFromWideInt(GETVAL(i,__int64)));
        }
        break;

        /* Dunno how to handle these */
    default:
        break;

    }

    SafeArrayUnaccessData(arrP);

 alldone:
    return ObjNewList(objc, objv);
}

/* 
 * If value_only is 0, returns a Tcl object that is a list {VT_xxx value}.
 * If VT_xxx is not known, value is missing (only the VT_xxx is
 * returned). 
 * If value_only is 1, returns only the value object and am empty object
 * if VT_xxx is not known.
 * Never returns NULL.
 */
Tcl_Obj *ObjFromVARIANT(VARIANT *varP, int value_only)
{
    Tcl_Obj *objv[2];
    Tcl_Obj *valObj[2];
    unsigned long ulval;
    VARIANT  empty;
    void    *recdataP;
    IDispatch *idispP;
    IUnknown  *iunkP;

    if (varP == NULL) {
        VariantInit(&empty);
        varP = &empty;
    }

    if (V_VT(varP) & VT_ARRAY) {
        if (V_VT(varP) & VT_BYREF)
            return ObjFromSAFEARRAY(*(varP->pparray));
        else
            return ObjFromSAFEARRAY(varP->parray);
    }
    if ((V_VT(varP) == (VT_BYREF|VT_VARIANT)) && varP->pvarVal)
        return ObjFromVARIANT(varP->pvarVal, 0);

    objv[0] = ObjFromInt(V_VT(varP) & ~VT_BYREF);
    objv[1] = NULL;

    switch (V_VT(varP)) {
    case VT_EMPTY|VT_BYREF:
    case VT_EMPTY:
    case VT_NULL|VT_BYREF:
    case VT_NULL:
        break;

    case VT_I2|VT_BYREF:
    case VT_I2:
        objv[1] = ObjFromInt(V_VT(varP) == VT_I2 ? V_I2(varP) : * V_I2REF(varP));
        break;

    case VT_I4|VT_BYREF:
    case VT_I4:
        objv[1] = ObjFromInt(V_VT(varP) == VT_I4 ? V_I4(varP) : * V_I4REF(varP));
        break;

    case VT_R4|VT_BYREF:
    case VT_R4:
        objv[1] = Tcl_NewDoubleObj(V_VT(varP) == VT_R4 ? V_R4(varP) : * V_R4REF(varP));
        break;

    case VT_R8|VT_BYREF:
    case VT_R8:
        objv[1] = Tcl_NewDoubleObj(V_VT(varP) == VT_R8 ? V_R8(varP) : * V_R8REF(varP));
        break;

    case VT_CY|VT_BYREF:
    case VT_CY:
        objv[1] = ObjFromCY(
            V_VT(varP) == VT_CY ? & V_CY(varP) : V_CYREF(varP)
            );
        break;

    case VT_BSTR|VT_BYREF:
    case VT_BSTR:
        if (V_VT(varP) == VT_BSTR)
            objv[1] = ObjFromUnicodeN(V_BSTR(varP),
                                        SysStringLen(V_BSTR(varP)));
        else
            objv[1] = ObjFromUnicodeN(* V_BSTRREF(varP),
                                        SysStringLen(* V_BSTRREF(varP)));
        break;

    case VT_DISPATCH|VT_BYREF:
        /* If VT_BYREF is set, then a reference to an existing
         * IUnknown is being returned. In this case, at the script level
         * we should not Release it but there is no way for the script
         * to know that. We therefore do a AddRef on the pointer here
         * so it can be released later in the script (ie. the script
         * can treat VT_DISPATCH and VT_DISPATCH|VT_BYREF the same
         * TBD - revisit this as to whether Release is required
         */
        idispP = * (V_DISPATCHREF(varP));
        idispP->lpVtbl->AddRef(idispP);
        objv[1] = ObjFromIDispatch(idispP);
        break;

    case VT_DISPATCH:
        idispP = V_DISPATCH(varP);
        objv[1] = ObjFromIDispatch(idispP);
        break;

    case VT_ERROR|VT_BYREF:
    case VT_ERROR:
        objv[1] = ObjFromInt(V_VT(varP) == VT_ERROR ? V_ERROR(varP) : * V_ERRORREF(varP));
        break;

    case VT_BOOL|VT_BYREF:
    case VT_BOOL:
        objv[1] = ObjFromBoolean(V_VT(varP) == VT_BOOL ? V_BOOL(varP) : * V_BOOLREF(varP));
        break;

    case VT_DATE|VT_BYREF:
    case VT_DATE:
        objv[1] = Tcl_NewDoubleObj(V_VT(varP) == VT_DATE ? V_DATE(varP) : * V_DATEREF(varP));
        break;

    case VT_VARIANT|VT_BYREF:
        /* This is for the case where varP->pvarVal is NULL. The non-NULL
           case was already handled in an if stmt above */
        break;

    case VT_DECIMAL|VT_BYREF:
    case VT_DECIMAL:
        objv[1] = ObjFromDECIMAL(
            V_VT(varP) == VT_DECIMAL ? & V_DECIMAL(varP) : V_DECIMALREF(varP)
            );
        break;


    case VT_UNKNOWN|VT_BYREF:
        /* If VT_BYREF is set, then a reference to an existing
         * IUnknown is being returned. In this case, at the script level
         * we should not Release it but there is no way for the script
         * to know that. We therefore do a AddRef on the pointer here
         * so it can be released later
         * TBD - revisit this as to whether Release is required
         */
        iunkP = * (V_UNKNOWNREF(varP));
        iunkP->lpVtbl->AddRef(iunkP);
        objv[1] = ObjFromIUnknown(iunkP);
        break;

    case VT_UNKNOWN:
        iunkP = V_UNKNOWN(varP);
        objv[1] = ObjFromIUnknown(iunkP);
        break;

    case VT_I1|VT_BYREF:
    case VT_I1:
        objv[1] = ObjFromInt(V_VT(varP) == VT_I1 ? V_I1(varP) : * V_I1REF(varP));
        break;

    case VT_UI1|VT_BYREF:
    case VT_UI1:
        objv[1] = ObjFromInt(V_VT(varP) == VT_UI1 ? V_UI1(varP) : * V_UI1REF(varP));
        break;

    case VT_UI2|VT_BYREF:
    case VT_UI2:
        objv[1] = ObjFromInt(V_VT(varP) == VT_UI2 ? V_UI2(varP) : * V_UI2REF(varP));
        break;

    case VT_UI4|VT_BYREF:
    case VT_UI4:
        /* store as wide integer if it does not fit in signed 32 bits */
        ulval = V_VT(varP) == VT_UI4 ? V_UI4(varP) : * V_UI4REF(varP);
        objv[1] = ObjFromDWORD(ulval);
        break;

    case VT_I8|VT_BYREF:
    case VT_I8:
        objv[1] = ObjFromWideInt(V_VT(varP) == VT_I8 ? V_I8(varP) : * V_I8REF(varP));
        break;

    case VT_UI8|VT_BYREF:
    case VT_UI8:
        objv[1] = ObjFromWideInt(V_VT(varP) == VT_UI8 ? V_UI8(varP) : * V_UI8REF(varP));
        break;


    case VT_INT|VT_BYREF:
    case VT_INT:
        objv[1] = ObjFromInt(V_VT(varP) == VT_INT ? V_INT(varP) : * V_INTREF(varP));
        break;

    case VT_UINT|VT_BYREF:
    case VT_UINT:
        /* store as wide integer if it does not fit in signed 32 bits */
        ulval = V_VT(varP) == VT_UINT ? V_UINT(varP) : * V_UINTREF(varP);
        if (ulval & 0x80000000) {
            objv[1] = ObjFromWideInt(ulval);
        }
        else {
            objv[1] = ObjFromLong(ulval);
        }
        break;

    case VT_RECORD:
        recdataP = NULL;
        if (V_RECORDINFO(varP) &&
            V_RECORD(varP) &&
            V_RECORDINFO(varP)->lpVtbl->RecordCreateCopy(V_RECORDINFO(varP), V_RECORD(varP), &recdataP) == S_OK
            ) {
            /*
             * Construct return value as pair of IRecordInfo* void* (data)
             */
            valObj[0] = ObjFromOpaque(V_RECORDINFO(varP), "IRecordInfo");
            // TBD - we pass pointers to record instances as void* as per
            // the IRecordInfo interface. We should change this to be
            // more typesafe
            valObj[1] = ObjFromLPVOID(recdataP);
            objv[1] = ObjNewList(2, valObj);
        }
        break;

        /* Dunno how to handle these */
    case VT_RECORD|VT_BYREF:
    case VT_VARIANT: /* Note VT_VARIANT is illegal */
    default:
        break;
    }

    if (value_only)
        return objv[1] ? objv[1] : ObjFromEmptyString();
    else
        return ObjNewList(objv[1] ? 2 : 1, objv);
}

/* Returned memory in *arrayPP has to be freed by caller */
int ObjToLSASTRINGARRAY(Tcl_Interp *interp, Tcl_Obj *obj, LSA_UNICODE_STRING **arrayP, ULONG *countP)
{
    Tcl_Obj **listobjv;
    int       i, nitems, sz;
    LSA_UNICODE_STRING *ustrP;
    WCHAR    *dstP;

    if (ObjGetElements(interp, obj, &nitems, &listobjv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    /* Figure out how much space we need */
    sz = nitems * sizeof(LSA_UNICODE_STRING);
    for (i = 0; i < nitems; ++i) {
        sz += sizeof(*dstP) * (Tcl_GetCharLength(listobjv[i]) + 1);
    }

    ustrP = TwapiAlloc(sz);

    /* Figure out where the string area starts and do the construction */
    dstP = (WCHAR *) ((nitems * sizeof(LSA_UNICODE_STRING)) + (char *)(ustrP));
    for (i = 0; i < nitems; ++i) {
        WCHAR *srcP;
        int    slen;
        srcP = ObjToUnicodeN(listobjv[i], &slen);
        CopyMemory(dstP, srcP, sizeof(WCHAR)*(slen+1));
        ustrP[i].Buffer = dstP;
        ustrP[i].Length = (USHORT) (sizeof(WCHAR) * slen); /* Num *bytes*, not WCHARs */
        ustrP[i].MaximumLength = ustrP[i].Length;
        dstP += slen+1;
    }

    *arrayP = ustrP;
    *countP = nitems;

    return TCL_OK;
}




/*
 * Returns a pointer to dynamic memory containing a SID corresponding
 * to the given string representation. Returns NULL on error, and
 * sets the windows error
 */
PSID TwapiGetSidFromStringRep(char *strP)
{
    DWORD   len;
    PSID    sidP;
    PSID    local_sidP;
    int error;

    local_sidP = NULL;
    sidP = NULL;

    if (ConvertStringSidToSidA(strP, &local_sidP) == 0)
        return NULL;

    /*
     * Have a valid SID
     * Copy it into dynamic memory after validating
     */
    len = GetLengthSid(local_sidP);
    sidP = TwapiAlloc(len);
    if (! CopySid(len, sidP, local_sidP)) {
        goto errorExit;
    }

    /* Free memory allocated by ConvertStringSidToSidA */
    LocalFree(local_sidP);
    return sidP;

 errorExit:
    error = GetLastError();

    if (local_sidP) {
        LocalFree(local_sidP);
    }

    if (sidP)
        TwapiFree(sidP);

    SetLastError(error);
    return NULL;
}

/* Tcl_Obj to SID - the object may hold the SID string rep, a binary
   or a list of ints. If the object is an empty string, *sidPP is
   stored as NULL. Else the SID is dynamically allocated and a pointer to it is
   stored in *sidPP. Caller must release it by calling TwapiFree
*/
int ObjToPSID(Tcl_Interp *interp, Tcl_Obj *obj, PSID *sidPP)
{
    char *s;
    DWORD   len;
    SID  *sidP;
    DWORD winerror;

    s = ObjToStringN(obj, &len);
    if (len == 0) {
        *sidPP = NULL;
        return TCL_OK;
    }

    *sidPP = TwapiGetSidFromStringRep(ObjToString(obj));
    if (*sidPP)
        return TCL_OK;

    winerror = GetLastError();

    /* Not a string rep. See if it is a binary of the right size */
    sidP = (SID *) ObjToByteArray(obj, &len);
    if (len >= sizeof(*sidP)) {
        /* Seems big enough, validate revision and size */
        if (IsValidSid(sidP) && GetLengthSid(sidP) == len) {
            *sidPP = TwapiAlloc(len);
            /* Note SID is a variable length struct so we cannot do this
                    *(SID *) (*sidPP) = *sidP;
               (from bitter experience!)
             */
            if (CopySid(len, *sidPP, sidP))
                return TCL_OK;
            winerror = GetLastError();
        }
    }
    return Twapi_AppendSystemError(interp, winerror);
}

/* Convert a ACE object to a Tcl list. interp may be NULL */
Tcl_Obj *ObjFromACE (Tcl_Interp *interp, void *aceP)
{
    Tcl_Obj    *resultObj = NULL;
    Tcl_Obj    *obj = NULL;
    ACE_HEADER *acehdrP = &((ACCESS_ALLOWED_ACE *) aceP)->Header;
    ACCESS_ALLOWED_OBJECT_ACE *objectAceP;
    SID        *sidP;

    if (aceP == NULL) {
        if (interp)
            TwapiSetStaticResult(interp, "NULL ACE pointer");
        return NULL;
    }

    resultObj = ObjNewList(0, NULL);

    /* ACE type */
    ObjAppendElement(interp, resultObj,
                             ObjFromInt(acehdrP->AceType));

    /* ACE flags */
    ObjAppendElement(interp, resultObj,
                             ObjFromInt(acehdrP->AceFlags));

    /* Now for type specific fields */
    switch (acehdrP->AceType) {
    case ACCESS_ALLOWED_ACE_TYPE:
    case ACCESS_DENIED_ACE_TYPE:
    case SYSTEM_AUDIT_ACE_TYPE:
    case SYSTEM_MANDATORY_LABEL_ACE_TYPE:
        ObjAppendElement(interp, resultObj,
                                 ObjFromDWORD(((ACCESS_ALLOWED_ACE *)aceP)->Mask));

        /* and the SID */
        obj = NULL;                /* In case of errors */
        if (ObjFromSID(interp,
                         (SID *)&((ACCESS_ALLOWED_ACE *)aceP)->SidStart,
                         &obj)
            != TCL_OK) {
            goto error_return;
        }
        ObjAppendElement(interp, resultObj, obj);
        break;

    case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
    case ACCESS_DENIED_OBJECT_ACE_TYPE:
    case SYSTEM_AUDIT_OBJECT_ACE_TYPE:
        objectAceP = (ACCESS_ALLOWED_OBJECT_ACE *)aceP;
        ObjAppendElement(interp, resultObj,
                                 ObjFromDWORD(objectAceP->Mask));
        if (objectAceP->Flags & ACE_OBJECT_TYPE_PRESENT) {
            ObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->ObjectType));
            if (objectAceP->Flags & ACE_INHERITED_OBJECT_TYPE_PRESENT) {
                ObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->InheritedObjectType));
                sidP = (SID *) &objectAceP->SidStart;
            } else {
                ObjAppendElement(interp, resultObj, ObjFromEmptyString());
                sidP = (SID *) &objectAceP->InheritedObjectType;
            }
        } else if (objectAceP->Flags & ACE_INHERITED_OBJECT_TYPE_PRESENT) {
            ObjAppendElement(interp, resultObj, ObjFromEmptyString());
            ObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->ObjectType));
            sidP = (SID *) &objectAceP->InheritedObjectType;
        } else {
            ObjAppendElement(interp, resultObj, ObjFromEmptyString());
            ObjAppendElement(interp, resultObj, ObjFromEmptyString());
            sidP = (SID *) &objectAceP->ObjectType;
        }
        obj = NULL;                /* In case of errors */
        if (ObjFromSID(interp, sidP, &obj) != TCL_OK)
            goto error_return;
        ObjAppendElement(interp, resultObj, obj);
        
        break;

    default:
        /*
         * Return a binary rep of the whole dang thing.
         * There are no pointers in there, just values so this
         * should work, I think :)
         */
        obj = ObjFromByteArray((unsigned char *) aceP, acehdrP->AceSize);

        if (ObjAppendElement(interp, resultObj, obj) != TCL_OK)
            goto error_return;

        break;
    }


    return resultObj;

 error_return:
    Twapi_FreeNewTclObj(obj); /* OK if null */
    Twapi_FreeNewTclObj(resultObj); /* OK if null */
    return NULL;
}


int ObjToACE (Tcl_Interp *interp, Tcl_Obj *aceobj, void **acePP)
{
    Tcl_Obj **objv;
    int       objc;
    int       acetype;
    int       aceflags;
    int       acesz;
    SID      *sidP;
    unsigned char *bytes;
    int            bytecount;
    ACCESS_ALLOWED_ACE    *aceP;

    *acePP = NULL;

    if (ObjGetElements(interp, aceobj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc < 2)
        goto format_error;

    if ((ObjToInt(interp, objv[0], &acetype) != TCL_OK) ||
        (ObjToInt(interp, objv[1], &aceflags) != TCL_OK)) {
        return TCL_ERROR;
    }

    /* Max size of an SID */
    acesz = GetSidLengthRequired(SID_MAX_SUB_AUTHORITIES);

    /* Figure out how much space is required for the ACE based on type */
    switch (acetype) {
    case ACCESS_ALLOWED_ACE_TYPE:
    case ACCESS_DENIED_ACE_TYPE:
    case SYSTEM_AUDIT_ACE_TYPE:
    case SYSTEM_MANDATORY_LABEL_ACE_TYPE:
        if (objc != 4)
            goto format_error;
        acesz += sizeof(*aceP);
        aceP = (ACCESS_ALLOWED_ACE *) TwapiAlloc(acesz);
        aceP->Header.AceType = acetype;
        aceP->Header.AceFlags = aceflags;
        aceP->Header.AceSize  = acesz; /* TBD - this is a upper bound since we
                                          allocated max SID size. Is that OK?*/
        if (ObjToInt(interp, objv[2], &aceP->Mask) != TCL_OK)
            goto format_error;

        sidP = TwapiGetSidFromStringRep(ObjToString(objv[3]));
        if (sidP == NULL)
            goto system_error;

        if (! CopySid(aceP->Header.AceSize - sizeof(*aceP) + sizeof(aceP->SidStart),
                      &aceP->SidStart, sidP)) {
            TwapiFree(sidP);
            goto system_error;
        }

        TwapiFree(sidP);
        sidP = NULL;

        break;

    default:
        if (objc != 3)
            goto format_error;
        bytes = ObjToByteArray(objv[2], &bytecount);
        acesz += bytecount;
        aceP = (ACCESS_ALLOWED_ACE *) TwapiAlloc(acesz);
        CopyMemory(aceP, bytes, bytecount);
        break;
    }

    *acePP = aceP;
    return TCL_OK;

 format_error:
    if (interp)
        TwapiSetStaticResult(interp, "Invalid ACE format.");
    return TCL_ERROR;

 system_error:
    return TwapiReturnSystemError(interp);
}

Tcl_Obj *ObjFromACL (
    Tcl_Interp *interp,
    ACL *aclP                   /* May be NULL */
)
{
    Tcl_Obj                 *objv[2] = { NULL, NULL} ;
    ACL_REVISION_INFORMATION acl_rev;
    ACL_SIZE_INFORMATION     acl_szinfo;
    DWORD                    i;

    if (aclP == NULL) {
        return STRING_LITERAL_OBJ("null");
    }

    if ((GetAclInformation(aclP, &acl_rev, sizeof(acl_rev),
                           AclRevisionInformation) == 0) ||
        GetAclInformation(aclP, &acl_szinfo, sizeof(acl_szinfo),
                          AclSizeInformation) == 0) {
        TwapiReturnSystemError(interp);
        return NULL;
    }

    objv[0] = ObjFromInt(acl_rev.AclRevision);
    objv[1] = ObjNewList(0, NULL);

    /* Loop and add the list of ACE's */
    for (i = 0; i < acl_szinfo.AceCount; ++i) {
        void    *aceP;
        Tcl_Obj *ace_obj;

        if (GetAce(aclP, i, &aceP) == 0) {
            TwapiReturnSystemError(interp);
            goto error_return;
        }
        ace_obj = ObjFromACE(interp, aceP);
        if (ace_obj == NULL)
            goto error_return;
        if (ObjAppendElement(interp, objv[1], ace_obj) != TCL_OK) {
            goto error_return;
        }
    }


    return ObjNewList(2, objv);

 error_return:
    Twapi_FreeNewTclObj(objv[0]); /* OK if null */
    Twapi_FreeNewTclObj(objv[1]); /* OK if null */
    return NULL;
}


/*
 * Returns a pointer to dynamic memory containing a ACL corresponding
 * to the given string representation. The string "null" is treated
 * as no acl and a NULL pointer is returned in *aclPP
 */
int ObjToPACL(Tcl_Interp *interp, Tcl_Obj *aclObj, ACL **aclPP)
{
    int       objc;
    Tcl_Obj **objv;
    Tcl_Obj **aceobjv;
    int       aceobjc;
    void    **acePP = NULL;
    int       i;
    int       aclsz;
    ACE_HEADER *acehdrP;
    int       aclrev;

    *aclPP = NULL;
    if (!lstrcmpA("null", ObjToString(aclObj)))
        return TCL_OK;

    if (ObjGetElements(interp, aclObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc != 2) {
        if (interp)
            TwapiSetStaticResult(interp, "Invalid ACL format.");
        return TCL_ERROR;
    }

    /*
     * First figure out how much space we need to allocate. For this, we
     * first need to figure out space for the ACE's
     */
#if 0
    objv[0] is the ACL rev. We always recalculate it, ignore value passed in.
    if (ObjToInt(interp, objv[0], &aclrev) != TCL_OK)
        goto error_return;
#endif
    if (ObjGetElements(interp, objv[1], &aceobjc, &aceobjv) != TCL_OK)
        goto error_return;

    aclsz = sizeof(ACL);
    aclrev = ACL_REVISION;
    if (aceobjc) {
        acePP = TwapiAlloc(aceobjc*sizeof(*acePP));
        for (i = 0; i < aceobjc; ++i)
            acePP[i] = NULL;        /* Init for error return */

        for (i = 0; i < aceobjc; ++i) {
            if (ObjToACE(interp, aceobjv[i], &acePP[i]) != TCL_OK)
                goto error_return;
            acehdrP = (ACE_HEADER *)acePP[i];
            aclsz += acehdrP->AceSize;
            switch (acehdrP->AceType) {
            case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
            case ACCESS_DENIED_OBJECT_ACE_TYPE:
            case SYSTEM_AUDIT_OBJECT_ACE_TYPE:
            case SYSTEM_ALARM_OBJECT_ACE_TYPE:
            case ACCESS_ALLOWED_CALLBACK_OBJECT_ACE_TYPE:
            case ACCESS_DENIED_CALLBACK_OBJECT_ACE_TYPE:
            case SYSTEM_AUDIT_CALLBACK_OBJECT_ACE_TYPE:
            case SYSTEM_ALARM_CALLBACK_OBJECT_ACE_TYPE:
                /* Change rev if object ace's present */
                aclrev = ACL_REVISION_DS;
                break;
            default:
                break;
            }

        }
    }

    /*
     * OK, now allocate the ACL and add the ACE's to it
     * We currently use AddAce, not AddMandatoryAce even for integrity labels.
     * This seems to work and avoids AddMandatoryAce which is not present
     * on XP/2k3
     */
    *aclPP = TwapiAlloc(aclsz);
    InitializeAcl(*aclPP, aclsz, aclrev);
    for (i = 0; i < aceobjc; ++i) {
        acehdrP = (ACE_HEADER *)acePP[i];
        if (! AddAce(*aclPP, aclrev, MAXDWORD, acePP[i], acehdrP->AceSize)) {
            TwapiReturnSystemError(interp);
            goto error_return;
        }
    }

    if (! IsValidAcl(*aclPP)) {
        if (interp)
            TwapiSetStaticResult(interp, "Internal error constructing ACL");
        goto error_return;
    }

    /* Free up temporary ACE storage */
    if (acePP) {
        for (i = 0; i < aceobjc; ++i)
            TwapiFree(acePP[i]);
        TwapiFree(acePP);
    }

    return TCL_OK;

 error_return:
    if (acePP) {
        for (i = 0; i < aceobjc; ++i)
            TwapiFree(acePP[i]);
        TwapiFree(acePP);
    }

    if (*aclPP) {
        TwapiFree(*aclPP);
        *aclPP = NULL;
    }

    return TCL_ERROR;
}

/* Create a list object from a security descriptor */
Tcl_Obj *ObjFromSECURITY_DESCRIPTOR(
    Tcl_Interp *interp,
    SECURITY_DESCRIPTOR *secdP
)
{
    SECURITY_DESCRIPTOR_CONTROL secd_control;
    SID      *sidP;
    ACL      *aclP;
    BOOL      aclpresent;
    Tcl_Obj  *objv[5] = { NULL, NULL, NULL, NULL, NULL} ;
    DWORD    rev;
    BOOL     defaulted;

    if (secdP == NULL) {
        return ObjNewList(0, NULL);
    }

    if (! GetSecurityDescriptorControl(secdP, &secd_control, &rev))
        goto system_error;

    if (rev != SECURITY_DESCRIPTOR_REVISION) {
        /* Dunno how to handle this */
        if (interp)
            TwapiSetStaticResult(interp, "Unsupported SECURITY_DESCRIPTOR version");
        goto error_return;
    }

    /* Control bits */
    objv[0] = ObjFromInt(secd_control);

    /* Owner SID */
    if (! GetSecurityDescriptorOwner(secdP, &sidP, &defaulted))
        goto system_error;
    if (sidP == NULL)
        objv[1] = ObjFromEmptyString();
    else {
        if (ObjFromSID(interp, sidP, &objv[1]) != TCL_OK)
            goto error_return;
    }

    /* Group SID */
    if (! GetSecurityDescriptorGroup(secdP, &sidP, &defaulted))
        goto system_error;
    if (sidP == NULL)
        objv[2] = ObjFromEmptyString();
    else {
        if (ObjFromSID(interp, sidP, &objv[2]) != TCL_OK)
            goto error_return;
    }

    /* DACL */
    if (! GetSecurityDescriptorDacl(secdP, &aclpresent, &aclP, &defaulted))
        goto system_error;
    if (! aclpresent)
        aclP = NULL;
    objv[3] = ObjFromACL(interp, aclP);

    /* SACL */
    if (! GetSecurityDescriptorSacl(secdP, &aclpresent, &aclP, &defaulted))
        goto system_error;
    if (! aclpresent)
        aclP = NULL;
    objv[4] = ObjFromACL(interp, aclP);

    /* All done, phew ... */
    return ObjNewList(5, objv);

 system_error:
    TwapiReturnSystemError(interp);

 error_return:
    for (rev = 0; rev < sizeof(objv)/sizeof(objv[0]); ++rev) {
        Twapi_FreeNewTclObj(objv[rev]);
    }
    return NULL;
}


/*
 * Returns a pointer to dynamic memory containing a structure corresponding
 * to the given string representation. Note that the owner, group, sacl
 * and dacl fields of the descriptor point to dynamic memory as well!
 */
int ObjToPSECURITY_DESCRIPTOR(
    Tcl_Interp *interp,
    Tcl_Obj *secdObj,
    SECURITY_DESCRIPTOR **secdPP
)
{
    int       objc;
    Tcl_Obj **objv;
    int       temp;
    SECURITY_DESCRIPTOR_CONTROL      secd_control;
    SECURITY_DESCRIPTOR_CONTROL      secd_control_mask;
    SID      *owner_sidP;
    SID      *group_sidP;
    ACL      *daclP;
    ACL      *saclP;
    char     *s;
    int       slen;

    owner_sidP = group_sidP = NULL;
    *secdPP = NULL;

    if (ObjGetElements(interp, secdObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc == 0)
        return TCL_OK;          /* NULL security descriptor */

    if (objc != 5) {
        if (interp)
            TwapiSetStaticResult(interp, "Invalid SECURITY_DESCRIPTOR format.");
        return TCL_ERROR;
    }


    *secdPP = TwapiAlloc (sizeof(SECURITY_DESCRIPTOR));
    if (! InitializeSecurityDescriptor(*secdPP, SECURITY_DESCRIPTOR_REVISION))
        goto system_error;

    /*
     * Set control field
     */
    if (ObjToInt(interp, objv[0], &temp) != TCL_OK)
        goto error_return;
    secd_control = (SECURITY_DESCRIPTOR_CONTROL) temp;
    if (secd_control != temp) {
        /* Truncation error */
        if (interp)
            TwapiSetStaticResult(interp, "Invalid control flags for SECURITY_DESCRIPTOR");
        goto error_return;
    }

    /* Mask of control bits to be set through SetSecurityDescriptorControl*/
    /* Note you cannot set any other bits than these through the
       SetSecurityDescriptorControl */
    secd_control_mask =  (SE_DACL_AUTO_INHERIT_REQ | SE_DACL_AUTO_INHERITED |
                          SE_DACL_PROTECTED |
                          SE_SACL_AUTO_INHERIT_REQ | SE_SACL_AUTO_INHERITED |
                          SE_SACL_PROTECTED);

    if (! SetSecurityDescriptorControl(*secdPP, secd_control_mask, (SECURITY_DESCRIPTOR_CONTROL) (secd_control_mask & secd_control)))
        goto system_error;

    /*
     * Set Owner field if specified
     */
    s = ObjToStringN(objv[1], &slen);
    if (slen) {
        owner_sidP = TwapiGetSidFromStringRep(s);
        if (owner_sidP == NULL)
            goto system_error;
        if (! SetSecurityDescriptorOwner(*secdPP, owner_sidP,
                                         secd_control & SE_OWNER_DEFAULTED))
            goto system_error;
        /* Note the owner field in *secdPP now points directly to owner_sidP! */
    }

    /*
     * Set group field if specified
     */
    s = ObjToStringN(objv[2], &slen);
    if (slen) {
        group_sidP = TwapiGetSidFromStringRep(s);
        if (group_sidP == NULL)
            goto system_error;

        if (! SetSecurityDescriptorGroup(*secdPP, group_sidP,
                                         secd_control & SE_GROUP_DEFAULTED))
            goto system_error;
        /* Note the group field in *secdPP now points directly to group_sidP! */
    }

    /*
     * Set the DACL. Keyword "null" means no DACL (as opposed to an empty one)
     */
    if (ObjToPACL(interp, objv[3], &daclP) != TCL_OK)
        goto error_return;
    if (! SetSecurityDescriptorDacl(*secdPP, (daclP != NULL), daclP,
                                  (secd_control & SE_DACL_DEFAULTED)))
        goto system_error;
    /* Note the dacl field in *secdPP now points directly to daclP! */


    /*
     * Set the SACL. Keyword "null" means no SACL (as opposed to an empty one)
     */
    if (ObjToPACL(interp, objv[4], &saclP) != TCL_OK)
        goto error_return;
    if (! SetSecurityDescriptorSacl(*secdPP, (saclP != NULL), saclP,
                                  (secd_control & SE_SACL_DEFAULTED)))
        goto system_error;
    /* Note the sacl field in *secdPP now points directly to saclP! */
    return TCL_OK;

 system_error:
    TwapiReturnSystemError(interp);
    goto error_return;

 error_return:
    if (owner_sidP)
        TwapiFree(owner_sidP);
    if (group_sidP)
        TwapiFree(group_sidP);
    if (daclP)
        TwapiFree(daclP);
    if (saclP)
        TwapiFree(saclP);
    if (*secdPP) {
        TwapiFree(*secdPP);
        *secdPP = NULL;
    }
    return TCL_ERROR;
}


/* Free the security descriptor contents as if it was allocated through
 * ObjToPSECURITY_DESCRIPTOR
 */
void TwapiFreeSECURITY_DESCRIPTOR(SECURITY_DESCRIPTOR *secdP)
{
    SID      *sidP;
    ACL      *aclP;
    BOOL      aclpresent;
    BOOL      defaulted;

    if (secdP == NULL)
        return;

    if (!IsValidSecurityDescriptor(secdP)) {
        return;                 /* TBD - Should log an error here */
    }

    /* Owner SID */
    if (GetSecurityDescriptorOwner(secdP, &sidP, &defaulted) && sidP)
        TwapiFree(sidP);

    /* Group SID */
    if (GetSecurityDescriptorGroup(secdP, &sidP, &defaulted) && sidP)
        TwapiFree(sidP);

    /* DACL */
    if (GetSecurityDescriptorDacl(secdP, &aclpresent, &aclP, &defaulted)
        && aclpresent
        && aclP) {
        TwapiFree(aclP);
    }

    /* SACL */
    if (GetSecurityDescriptorSacl(secdP, &aclpresent, &aclP, &defaulted)
        && aclpresent
        && aclP) {

        TwapiFree(aclP);
    }

    TwapiFree(secdP);
}

/* Free the security descriptor contents as if it was allocated through
 * ObjToPSECURITY_ATTRIBUTES
 */
void TwapiFreeSECURITY_ATTRIBUTES(SECURITY_ATTRIBUTES *secattrP)
{
    if (secattrP == NULL)
        return;
    if (secattrP->lpSecurityDescriptor)
        TwapiFreeSECURITY_DESCRIPTOR(secattrP->lpSecurityDescriptor);

    TwapiFree(secattrP);
}


/*
 * Returns a pointer to dynamic memory containing a structure corresponding
 * to the given string representation.
 * The SECURITY_DESCRIPTOR field should be freed through
 * TwapiFreeSECURITY_DESCRIPTOR
 */
int ObjToPSECURITY_ATTRIBUTES(
    Tcl_Interp *interp,
    Tcl_Obj *secattrObj,
    SECURITY_ATTRIBUTES **secattrPP
)
{
    int       objc;
    Tcl_Obj **objv;
    int       inherit;


    *secattrPP = NULL;

    if (ObjGetElements(interp, secattrObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc == 0)
        return TCL_OK;          /* NULL security attributes */

    if (objc != 2) {
        if (interp)
            TwapiSetStaticResult(interp, "Invalid SECURITY_ATTRIBUTES format.");
        return TCL_ERROR;
    }


    *secattrPP = TwapiAlloc (sizeof(**secattrPP));
    (*secattrPP)->nLength = sizeof(**secattrPP);

    if (ObjToInt(interp, objv[1], &inherit) == TCL_ERROR)
        goto error_return;
    (*secattrPP)->bInheritHandle = (inherit != 0);

    if (ObjToPSECURITY_DESCRIPTOR(interp, objv[0],
                                     &(SECURITY_DESCRIPTOR *)((*secattrPP)->lpSecurityDescriptor))
        == TCL_ERROR) {
        goto error_return;
    }

    return TCL_OK;

 error_return:
    if (*secattrPP) {
        TwapiFree(*secattrPP);
        *secattrPP = NULL;
    }
    return TCL_ERROR;
}

Tcl_UniChar *ObjToUnicode(Tcl_Obj *objP)
{
    return Tcl_GetUnicode(objP);
}

Tcl_UniChar *ObjToUnicodeN(Tcl_Obj *objP, int *lenP)
{
    return Tcl_GetUnicodeFromObj(objP, lenP);
}

Tcl_Obj *ObjFromUnicodeN(const Tcl_UniChar *ws, int len)
{
#if USE_UNICODE_OBJ
    return Tcl_NewUnicodeObj(ws, len);
#else
    return TwapiUtf8ObjFromUnicode(ws, len);
#endif
}

Tcl_Obj *ObjFromUnicode(const Tcl_UniChar *ws)
{
#if USE_UNICODE_OBJ
    return Tcl_NewUnicodeObj(ws, -1);
#else
    return TwapiUtf8ObjFromUnicode(ws, -1);
#endif
}

char *ObjToString(Tcl_Obj *objP)
{
    return Tcl_GetString(objP);
}

char *ObjToStringN(Tcl_Obj *objP, int *lenP)
{
    return Tcl_GetStringFromObj(objP, lenP);
}

Tcl_Obj *ObjFromStringN(const char *s, int len)
{
    return Tcl_NewStringObj(s, len);
}

Tcl_Obj *ObjFromString(const char *s)
{
    return Tcl_NewStringObj(s, -1);
}

TCL_RESULT ObjToLong(Tcl_Interp *interp, Tcl_Obj *objP, long *lvalP)
{
    return Tcl_GetLongFromObj(interp, objP, lvalP);
}

TCL_RESULT ObjToBoolean(Tcl_Interp *interp, Tcl_Obj *objP, int *valP)
{
    return Tcl_GetBooleanFromObj(interp, objP, valP);
}

TCL_RESULT ObjToWideInt(Tcl_Interp *interp, Tcl_Obj *objP, Tcl_WideInt *wideP)
{
    return Tcl_GetWideIntFromObj(interp, objP, wideP);
}

Tcl_Obj *ObjNewList(int objc, Tcl_Obj * const objv[])
{
    return Tcl_NewListObj(objc, objv);
}

Tcl_Obj *ObjEmptyList(void)
{
    return Tcl_NewObj();
}

TCL_RESULT ObjAppendElement(Tcl_Interp *interp, Tcl_Obj *l, Tcl_Obj *e)
{
    return Tcl_ListObjAppendElement(interp, l, e);
}

TCL_RESULT ObjGetElements(Tcl_Interp *interp, Tcl_Obj *l, int *objcP, Tcl_Obj ***objvP)
{
    return Tcl_ListObjGetElements(interp, l, objcP, objvP);
}

Tcl_Obj *ObjFromLong(long val)
{
    return Tcl_NewLongObj(val);
}

Tcl_Obj *ObjFromWideInt(Tcl_WideInt val)
{
    return Tcl_NewWideIntObj(val);
}

Tcl_Obj *ObjFromBoolean(int bval)
{
    return Tcl_NewBooleanObj(bval);
}

Tcl_Obj *ObjFromEmptyString()
{
    return Tcl_NewObj();
}

Tcl_Obj *ObjFromByteArray(const unsigned char *bytes, int len)
{
    return Tcl_NewByteArrayObj(bytes, len);
}

unsigned char *ObjToByteArray(Tcl_Obj *objP, int *lenP)
{
    return Tcl_GetByteArrayFromObj(objP, lenP);
}
