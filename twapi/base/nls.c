/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

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
