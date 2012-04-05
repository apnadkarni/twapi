/*
 * Copyright (c) 2010-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include <locale.h>

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

int Twapi_GetNumberFormat(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    DWORD opts;
    DWORD loc;
    DWORD flags;
    LPCWSTR number_string;
    UINT      ndigits; 
    UINT      leading_zero; 
    UINT      grouping; 
    LPWSTR    decimal_sep; 
    LPWSTR    thousand_sep; 
    UINT      negative_order;

    NUMBERFMTW numfmt;
    NUMBERFMTW *fmtP;
    int        numchars;
    WCHAR     *buf;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETINT(opts), GETINT(loc), GETINT(flags),
                     GETWSTR(number_string), GETINT(ndigits),
                     GETINT(leading_zero),
                     GETINT(grouping), GETWSTR(decimal_sep),
                     GETWSTR(thousand_sep), GETINT(negative_order),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;


#define TWAPI_GETNUMBERFORMAT_USELOCALE 1 // Defined in newer SDK's

    if (opts & TWAPI_GETNUMBERFORMAT_USELOCALE) {
        fmtP = NULL;
    } else {
        numfmt.NumDigits = ndigits;
        numfmt.LeadingZero = leading_zero;
        numfmt.Grouping = grouping;
        numfmt.lpDecimalSep = decimal_sep;
        numfmt.lpThousandSep = thousand_sep;
        numfmt.NegativeOrder = negative_order;
        fmtP = &numfmt;
    }

    numchars = GetNumberFormatW(loc, flags, number_string, fmtP, NULL, 0);
    if (numchars == 0) {
        return TwapiReturnSystemError(ticP->interp);
    }

    buf = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR)*(numchars+1), NULL);

    numchars = GetNumberFormatW(loc, flags, number_string, fmtP, buf, numchars);
    if (numchars == 0)
        TwapiReturnSystemError(ticP->interp);
    else
        Tcl_SetObjResult(ticP->interp, ObjFromUnicodeN(buf, numchars-1));

    MemLifoPopFrame(&ticP->memlifo);

    return numchars ? TCL_OK : TCL_ERROR;
}


int Twapi_GetCurrencyFormat(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    DWORD opts;
    DWORD loc;
    DWORD flags;              // options
    LPCWSTR number_string;            // input number string
    UINT      ndigits;
    UINT      leading_zero; 
    UINT      grouping; 
    LPWSTR    decimal_sep; 
    LPWSTR    thousand_sep; 
    UINT      negative_order;
    UINT      positive_order;
    LPWSTR    currency_sym;

    CURRENCYFMTW fmt;
    CURRENCYFMTW *fmtP;
    int        numchars;
    WCHAR     *buf;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETINT(opts), GETINT(loc), GETINT(flags),
                     GETWSTR(number_string), GETINT(ndigits),
                     GETINT(leading_zero),
                     GETINT(grouping), GETWSTR(decimal_sep),
                     GETWSTR(thousand_sep), GETINT(negative_order),
                     GETINT(positive_order), GETWSTR(currency_sym),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;


#define TWAPI_GETCURRENCYFORMAT_USELOCALE 1

    if (opts & TWAPI_GETCURRENCYFORMAT_USELOCALE) {
        fmtP = NULL;
    } else {
        fmt.NumDigits = ndigits;
        fmt.LeadingZero = leading_zero;
        fmt.Grouping = grouping;
        fmt.lpDecimalSep = decimal_sep;
        fmt.lpThousandSep = thousand_sep;
        fmt.NegativeOrder = negative_order;
        fmt.PositiveOrder = positive_order;
        fmt.lpCurrencySymbol = currency_sym;
        fmtP = &fmt;
    }

    numchars = GetCurrencyFormatW(loc, flags, number_string, fmtP, NULL, 0);
    if (numchars == 0) {
        return TwapiReturnSystemError(ticP->interp);
    }
    buf = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR)*(numchars+1), NULL);

    numchars = GetCurrencyFormatW(loc, flags, number_string, fmtP, buf, numchars);
    if (numchars == 0)
        TwapiReturnSystemError(ticP->interp);
    else
        Tcl_SetObjResult(ticP->interp, ObjFromUnicodeN(buf, numchars-1));

    MemLifoPopFrame(&ticP->memlifo);

    return numchars ? TCL_OK : TCL_ERROR;
}

static int Twapi_NlsCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    DWORD dw, dw2;
    TwapiResult result;
    WCHAR buf[MAX_PATH+1];

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 23:
        result.type = TRT_DWORD;
        result.value.uval = GetUserDefaultLangID();
        break;
    case 24:
        result.type = TRT_DWORD;
        result.value.uval = GetSystemDefaultLangID();
        break;
    case 25:
        result.type = TRT_DWORD;
        result.value.uval = GetUserDefaultLCID();
        break;
    case 26:
        result.type = TRT_DWORD;
        result.value.uval = GetSystemDefaultLCID();
        break;
    case 27:
        result.type = TRT_NONZERO_RESULT;
        result.value.uval = GetUserDefaultUILanguage();
        break;
    case 28:
        result.type = TRT_NONZERO_RESULT;
        result.value.uval = GetSystemDefaultUILanguage();
        break;
    case 29:
        result.type = TRT_DWORD;
        result.value.uval = GetThreadLocale();
        break;
    case 30:
        result.type = TRT_DWORD;
        result.value.uval = GetACP();
        break;
    case 31:
        result.type = TRT_DWORD;
        result.value.uval = GetOEMCP();
        break;
    case 32:
        return Twapi_GetNumberFormat(ticP, objc-2, objv+2);
    case 33:
        return Twapi_GetCurrencyFormat(ticP, objc-2, objv+2);
    case 34:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        result.value.unicode.len = VerLanguageNameW(dw, buf, sizeof(buf)/sizeof(buf[0]));
        result.value.unicode.str = buf;
        result.type = result.value.unicode.len ? TRT_UNICODE : TRT_GETLASTERROR;
        break;
    case 35: //GetLocaleInfo
        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        CHECK_INTEGER_OBJ(interp, dw2, objv[3]);
        result.value.unicode.len = GetLocaleInfoW(dw, dw2, buf,
                                                  ARRAYSIZE(buf));
        if (result.value.unicode.len == 0) {
            result.type = TRT_GETLASTERROR;
        } else {
            if (dw2 & LOCALE_RETURN_NUMBER) {
                // buf actually contains a number
                result.value.ival = *(int *)buf;
                result.type = TRT_LONG;
            } else {
                result.value.unicode.len -= 1;
                result.value.unicode.str = buf;
                result.type = TRT_UNICODE;
            }
        }
        break;
    }

    return TwapiSetResult(interp, &result);
}


static int TwapiNlsInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct alias_dispatch_s NlsAliasDispatch[] = {
        DEFINE_ALIAS_CMD(GetUserDefaultLangID, 23),
        DEFINE_ALIAS_CMD(GetSystemDefaultLangID, 24),
        DEFINE_ALIAS_CMD(GetUserDefaultLCID, 25),
        DEFINE_ALIAS_CMD(GetSystemDefaultLCID, 26),
        DEFINE_ALIAS_CMD(GetUserDefaultUILanguage, 27),
        DEFINE_ALIAS_CMD(GetSystemDefaultUILanguage, 28),
        DEFINE_ALIAS_CMD(GetThreadLocale, 29),
        DEFINE_ALIAS_CMD(GetACP, 30),
        DEFINE_ALIAS_CMD(GetOEMCP, 31),
        DEFINE_ALIAS_CMD(GetNumberFormat, 32),
        DEFINE_ALIAS_CMD(GetCurrencyFormat, 33),
        DEFINE_ALIAS_CMD(VerLanguageName, 34),
        DEFINE_ALIAS_CMD(GetLocaleInfo, 35),
    };

    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::NlsCall", Twapi_NlsCallObjCmd, ticP, NULL);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(NlsAliasDispatch), NlsAliasDispatch, "twapi::NlsCall");

    return TCL_OK;
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
int Twapi_nls_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiNlsInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

