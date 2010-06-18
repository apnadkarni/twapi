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
    GUID       *guidP,
    LPCWSTR     sitenameP,
    ULONG       flags
    )
{
    DWORD   status;
    PDOMAIN_CONTROLLER_INFOW dcP;

    status = DsGetDcNameW(systemnameP, domainnameP, guidP, sitenameP, flags, &dcP);
    if (status != ERROR_SUCCESS) {
        return Twapi_AppendSystemError(interp, status);
    }

    if (dcP) {
        Tcl_Obj *objv[18];
        objv[0] = STRING_LITERAL_OBJ("DomainControllerName");
        objv[1] = Tcl_NewUnicodeObj(dcP->DomainControllerName, -1);
        objv[2] = STRING_LITERAL_OBJ("DomainControllerAddress");
        objv[3] = Tcl_NewUnicodeObj(dcP->DomainControllerAddress, -1);
        objv[4] = STRING_LITERAL_OBJ("DomainControllerAddressType");
        objv[5] = Tcl_NewLongObj(dcP->DomainControllerAddressType);
        objv[6] = STRING_LITERAL_OBJ("DomainGuid");
        objv[7] = ObjFromGUID(&dcP->DomainGuid);
        objv[8] = STRING_LITERAL_OBJ("DomainName");
        objv[9] = Tcl_NewUnicodeObj(dcP->DomainName, -1);
        objv[10] = STRING_LITERAL_OBJ("DnsForestName");
        objv[11] = Tcl_NewUnicodeObj(dcP->DnsForestName, -1);
        objv[12] = STRING_LITERAL_OBJ("Flags");
        objv[13] = Tcl_NewLongObj(dcP->Flags);
        objv[14] = STRING_LITERAL_OBJ("DcSiteName");
        objv[15] = Tcl_NewUnicodeObj(dcP->DcSiteName, -1);
        objv[16] = STRING_LITERAL_OBJ("ClientSiteName");
        objv[17] = Tcl_NewUnicodeObj(dcP->ClientSiteName, -1);

        Tcl_SetObjResult(interp, Tcl_NewListObj(18, objv));

        NetApiBufferFree(dcP);
    }

    return TCL_OK;
}

