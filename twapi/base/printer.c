/* 
 * Copyright (c) 2004-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

int Twapi_EnumPrintersLevel4ObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    void   *buf = NULL;
    DWORD   sz = 1000;
    DWORD   needed_sz;
    DWORD   num_printers;
    DWORD   i;
    Tcl_Obj *resultObj;
    PRINTER_INFO_4W *printerInfoP;
    int flags;

    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, flags, objv[1]);

    buf = MemLifoPushFrame(&ticP->memlifo, sz, &sz);

    if (EnumPrintersW(flags, NULL, 4, buf, sz, &needed_sz, &num_printers) == FALSE) {
        DWORD err = GetLastError();
        if (err != ERROR_INSUFFICIENT_BUFFER) {
            MemLifoPopFrame(&ticP->memlifo);
            return Twapi_AppendSystemError(interp, err);
        }
    }

    if (needed_sz > sz) {
        /* Note - No need to free previous alloc. Will all get freed together */
        buf = MemLifoAlloc(&ticP->memlifo, needed_sz, &sz); 
        if (EnumPrintersW(flags, NULL, 4, buf, sz, &needed_sz, &num_printers) == FALSE) {
            TwapiReturnSystemError(interp); /* Store before calling free */
            MemLifoPopFrame(&ticP->memlifo);
            return TCL_ERROR;
        }
    }

    printerInfoP = (PRINTER_INFO_4W *)buf;
    resultObj = ObjNewList(0, NULL);
    for (i = 0; i < num_printers; ++i, ++printerInfoP) {
        Tcl_Obj *printerObj = ObjNewList(0, NULL);
        Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, printerObj, printerInfoP, pPrinterName);
        Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp, printerObj, printerInfoP, pServerName);
        Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, printerObj, printerInfoP, Attributes);
        ObjAppendElement(interp, resultObj, printerObj);
    }

    TwapiSetObjResult(interp, resultObj);
    MemLifoPopFrame(&ticP->memlifo);

    return TCL_OK;
}


