/* 
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

int Twapi_EnumPrinters_Level4(
    Tcl_Interp *interp,
    DWORD flags         // printer object types
)
{
    void   *buf = NULL;
    DWORD   sz = 1000;
    DWORD   needed_sz;
    DWORD   num_printers;
    DWORD   i;
    Tcl_Obj *resultObj;
    PRINTER_INFO_4W *printerInfoP;

    buf = Tcl_Alloc(sz);

    if (EnumPrintersW(flags, NULL, 4, buf, sz, &needed_sz, &num_printers) == FALSE) {
        DWORD err = GetLastError();
        if (err != ERROR_INSUFFICIENT_BUFFER) {
            Tcl_Free(buf);
            return Twapi_AppendSystemError(interp, err);
        }
    }

    if (needed_sz > sz) {
        Tcl_Free(buf);
        buf = Tcl_Alloc(needed_sz);
        if (EnumPrintersW(flags, NULL, 4, buf, needed_sz, &needed_sz, &num_printers) == FALSE) {
            TwapiReturnSystemError(interp); /* Store before calling free */
            Tcl_Free(buf);
            return TCL_ERROR;
        }
    }

    printerInfoP = (PRINTER_INFO_4W *)buf;
    resultObj = Tcl_NewListObj(0, NULL);
    for (i = 0; i < num_printers; ++i, ++printerInfoP) {
        Tcl_Obj *printerObj = Tcl_NewListObj(0, NULL);
        Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, printerObj, printerInfoP, pPrinterName);
        Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, printerObj, printerInfoP, pServerName);
        Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, printerObj, printerInfoP, Attributes);
        Tcl_ListObjAppendElement(interp, resultObj, printerObj);
    }

    Tcl_SetObjResult(interp, resultObj);
    Tcl_Free(buf);

    return TCL_OK;
}


