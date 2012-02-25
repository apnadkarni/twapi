/*
 * Copyright (c) 2010-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include <locale.h>

#ifndef TWAPI_STATIC_BUILD
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
        result.value.ival = GetUserDefaultLangID();
        break;
    case 24:
        result.type = TRT_DWORD;
        result.value.ival = GetSystemDefaultLangID();
        break;
    case 25:
        result.type = TRT_DWORD;
        result.value.ival = GetUserDefaultLCID();
        break;
    case 26:
        result.type = TRT_DWORD;
        result.value.ival = GetSystemDefaultLCID();
        break;
    case 27:
        result.type = TRT_NONZERO_RESULT;
        result.value.ival = GetUserDefaultUILanguage();
        break;
    case 28:
        result.type = TRT_NONZERO_RESULT;
        result.value.ival = GetSystemDefaultUILanguage();
        break;
    case 29:
        result.type = TRT_DWORD;
        result.value.ival = GetThreadLocale();
        break;
    case 30:
        result.type = TRT_DWORD;
        result.value.ival = GetACP();
        break;
    case 31:
        result.type = TRT_DWORD;
        result.value.ival = GetOEMCP();
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
                result.type = TRT_DWORD;
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


static int Twapi_NlsInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::NlsCall", Twapi_NlsCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::NlsCall", # code_); \
    } while (0);

    CALL_(GetUserDefaultLangID, 23);
    CALL_(GetSystemDefaultLangID, 24);
    CALL_(GetUserDefaultLCID, 25);
    CALL_(GetSystemDefaultLCID, 26);
    CALL_(GetUserDefaultUILanguage, 27);
    CALL_(GetSystemDefaultUILanguage, 28);
    CALL_(GetThreadLocale, 29);
    CALL_(GetACP, 30);
    CALL_(GetOEMCP, 31);
    CALL_(GetNumberFormat, 32);
    CALL_(GetCurrencyFormat, 33);
    CALL_(VerLanguageName, 34);
    CALL_(GetLocaleInfo, 35);

#undef CALL_

    return TCL_OK;
}


#ifndef TWAPI_STATIC_BUILD
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_STATIC_BUILD
__declspec(dllexport) 
#endif
int Twapi_nls_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, MODULENAME, MODULE_HANDLE,
                            Twapi_NlsInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

