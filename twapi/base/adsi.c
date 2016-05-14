/*
 * Copyright (c) 2006-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

int Twapi_DsGetDcName(
    Tcl_Interp *interp,
    LPCWSTR     systemnameP,
    LPCWSTR     domainnameP,
    UUID       *uuidP,
    LPCWSTR     sitenameP,
    ULONG       flags
    )
{
    DWORD   status;
    PDOMAIN_CONTROLLER_INFOW dcP;

    status = DsGetDcNameW(systemnameP, domainnameP, uuidP, sitenameP, flags, &dcP);
    if (status != ERROR_SUCCESS) {
        return Twapi_AppendSystemError(interp, status);
    }

    if (dcP) {
        Tcl_Obj *objv[18];
        objv[0] = STRING_LITERAL_OBJ("DomainControllerName");
        objv[1] = ObjFromWinChars(dcP->DomainControllerName);
        objv[2] = STRING_LITERAL_OBJ("DomainControllerAddress");
        objv[3] = ObjFromWinChars(dcP->DomainControllerAddress);
        objv[4] = STRING_LITERAL_OBJ("DomainControllerAddressType");
        objv[5] = ObjFromLong(dcP->DomainControllerAddressType);
        objv[6] = STRING_LITERAL_OBJ("DomainGuid");
        objv[7] = ObjFromUUID(&dcP->DomainGuid);
        objv[8] = STRING_LITERAL_OBJ("DomainName");
        objv[9] = ObjFromWinChars(dcP->DomainName);
        objv[10] = STRING_LITERAL_OBJ("DnsForestName");
        objv[11] = ObjFromWinChars(dcP->DnsForestName);
        objv[12] = STRING_LITERAL_OBJ("Flags");
        objv[13] = ObjFromLong(dcP->Flags);
        objv[14] = STRING_LITERAL_OBJ("DcSiteName");
        objv[15] = ObjFromWinChars(dcP->DcSiteName);
        objv[16] = STRING_LITERAL_OBJ("ClientSiteName");
        objv[17] = ObjFromWinChars(dcP->ClientSiteName);

        ObjSetResult(interp, ObjNewList(18, objv));

        NetApiBufferFree(dcP);
    }

    return TCL_OK;
}

