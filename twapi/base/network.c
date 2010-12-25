/*
 * Copyright (c) 2004-2010 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

typedef struct {
    PWCHAR pModuleName;
    PWCHAR pModulePath;
} TCPIP_OWNER_MODULE_BASIC_INFO;

/* Undocumented functions */
typedef DWORD (WINAPI *GetOwnerModuleFromTcpEntry_t)(PVOID, int, PVOID, PDWORD);
MAKE_DYNLOAD_FUNC(GetOwnerModuleFromTcpEntry, iphlpapi, GetOwnerModuleFromTcpEntry_t)
typedef DWORD (WINAPI *GetOwnerModuleFromUdpEntry_t)(PVOID, int, PVOID, PDWORD);
MAKE_DYNLOAD_FUNC(GetOwnerModuleFromUdpEntry, iphlpapi, GetOwnerModuleFromUdpEntry_t)

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

int ObjToSOCKADDR_IN(Tcl_Interp *interp, Tcl_Obj *objP, struct sockaddr_in *sinP)
{
    Tcl_Obj **objv;
    int       objc;

    if (Tcl_ListObjGetElements(interp, objP, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    sinP->sin_family = AF_INET;
    sinP->sin_addr.s_addr = 0;
    sinP->sin_port = 0;

    if (objc > 0) {
        sinP->sin_addr.s_addr = inet_addr(Tcl_GetString(objv[0]));
    }

    if (objc > 1) {
        if (ObjToWord(interp, objv[1], &sinP->sin_port) != TCL_OK)
            return TCL_ERROR;

        sinP->sin_port = htons(sinP->sin_port);
    }

    return TCL_OK;
}

/*
 * Given a IP_ADDR_STRING list, return a Tcl_Obj containing only
 * the IP Address components
 */
static Tcl_Obj *ObjFromIP_ADDR_STRINGAddress (
    Tcl_Interp *interp, const IP_ADDR_STRING *ipaddrstrP
)
{
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);
    while (ipaddrstrP) {
        if (ipaddrstrP->IpAddress.String[0])
            Tcl_ListObjAppendElement(interp, resultObj,
                                     Tcl_NewStringObj(ipaddrstrP->IpAddress.String, -1));
        ipaddrstrP = ipaddrstrP->Next;
    }

    return resultObj;
}

Tcl_Obj *ObjFromMIB_IPADDRROW(Tcl_Interp *interp, const MIB_IPADDRROW *iparP)
{
    Tcl_Obj *objv[5];

    objv[0] = IPAddrObjFromDWORD(iparP->dwAddr);
    objv[1] = Tcl_NewIntObj(iparP->dwIndex);
    objv[2] = IPAddrObjFromDWORD(iparP->dwMask);
    objv[3] = IPAddrObjFromDWORD(iparP->dwBCastAddr);
    objv[4] = Tcl_NewIntObj(iparP->dwReasmSize);
    return Tcl_NewListObj(5, objv);
}

Tcl_Obj *ObjFromMIB_IPADDRTABLE(Tcl_Interp *interp, MIB_IPADDRTABLE *ipatP)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < ipatP->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_IPADDRROW(interp,
                                                          &ipatP->table[i])
            );
    }
    return resultObj;
}

Tcl_Obj *ObjFromMIB_IFROW(Tcl_Interp *interp, const MIB_IFROW *ifrP)
{
    Tcl_Obj *objv[22];
    int len;
#if 0
    This field does not seem to contain a consistent format
    objv[0] = Tcl_NewUnicodeObj(ifrP->wszName, -1);
#else
    objv[0] = Tcl_NewStringObj("", 0);
#endif
    objv[1] = Tcl_NewIntObj(ifrP->dwIndex);
    objv[2] = Tcl_NewIntObj(ifrP->dwType);
    objv[3] = Tcl_NewIntObj(ifrP->dwMtu);
    objv[4] = Tcl_NewIntObj(ifrP->dwSpeed);
    objv[5] = Tcl_NewByteArrayObj(ifrP->bPhysAddr,ifrP->dwPhysAddrLen);
    objv[6] = Tcl_NewIntObj(ifrP->dwAdminStatus);
    objv[7] = Tcl_NewIntObj(ifrP->dwOperStatus);
    objv[8] = Tcl_NewIntObj(ifrP->dwLastChange);
    objv[9] = Tcl_NewWideIntObj(ifrP->dwInOctets);
    objv[10] = Tcl_NewWideIntObj(ifrP->dwInUcastPkts);
    objv[11] = Tcl_NewWideIntObj(ifrP->dwInNUcastPkts);
    objv[12] = Tcl_NewWideIntObj(ifrP->dwInDiscards);
    objv[13] = Tcl_NewWideIntObj(ifrP->dwInErrors);
    objv[14] = Tcl_NewIntObj(ifrP->dwInUnknownProtos);
    objv[15] = Tcl_NewWideIntObj(ifrP->dwOutOctets);
    objv[16] = Tcl_NewWideIntObj(ifrP->dwOutUcastPkts);
    objv[17] = Tcl_NewWideIntObj(ifrP->dwOutNUcastPkts);
    objv[18] = Tcl_NewWideIntObj(ifrP->dwOutDiscards);
    objv[19] = Tcl_NewWideIntObj(ifrP->dwOutErrors);
    objv[20] = Tcl_NewIntObj(ifrP->dwOutQLen);
    len =  ifrP->dwDescrLen;
    if (ifrP->bDescr[len-1] == 0)
        --len; /* Sometimes, not always, there is a terminating null */
    objv[21] = Tcl_NewStringObj(ifrP->bDescr, len);

    return Tcl_NewListObj(sizeof(objv)/sizeof(objv[0]), objv);
}

Tcl_Obj *ObjFromMIB_IFTABLE(Tcl_Interp *interp, MIB_IFTABLE *iftP)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < iftP->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_IFROW(interp,
                                                      &iftP->table[i])
            );
    }
    return resultObj;
}


Tcl_Obj *ObjFromMIB_IPNETROW(Tcl_Interp *interp, const MIB_IPNETROW *netrP)
{
    Tcl_Obj *objv[4];

    objv[0] = Tcl_NewIntObj(netrP->dwIndex);
    objv[1] = Tcl_NewByteArrayObj(netrP->bPhysAddr, netrP->dwPhysAddrLen);
    objv[2] = IPAddrObjFromDWORD(netrP->dwAddr);
    objv[3] = Tcl_NewIntObj(netrP->dwType);
    return Tcl_NewListObj(4, objv);
}

Tcl_Obj *ObjFromMIB_IPNETTABLE(Tcl_Interp *interp, MIB_IPNETTABLE *nettP)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < nettP->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_IPNETROW(interp,
                                                         &nettP->table[i])
            );
    }
    return resultObj;
}

Tcl_Obj *ObjFromMIB_IPFORWARDROW(Tcl_Interp *interp, const MIB_IPFORWARDROW *ipfrP)
{
    Tcl_Obj *objv[14];

    objv[0] = IPAddrObjFromDWORD(ipfrP->dwForwardDest);
    objv[1] = IPAddrObjFromDWORD(ipfrP->dwForwardMask);
    objv[2] = Tcl_NewIntObj(ipfrP->dwForwardPolicy);
    objv[3] = IPAddrObjFromDWORD(ipfrP->dwForwardNextHop);
    objv[4] = Tcl_NewIntObj(ipfrP->dwForwardIfIndex);
    objv[5] = Tcl_NewIntObj(ipfrP->dwForwardType);
    objv[6] = Tcl_NewIntObj(ipfrP->dwForwardProto);
    objv[7] = Tcl_NewIntObj(ipfrP->dwForwardAge);
    objv[8] = Tcl_NewIntObj(ipfrP->dwForwardNextHopAS);
    objv[9] = Tcl_NewIntObj(ipfrP->dwForwardMetric1);
    objv[10] = Tcl_NewIntObj(ipfrP->dwForwardMetric2);
    objv[11] = Tcl_NewIntObj(ipfrP->dwForwardMetric3);
    objv[12] = Tcl_NewIntObj(ipfrP->dwForwardMetric4);
    objv[13] = Tcl_NewIntObj(ipfrP->dwForwardMetric5);

    return Tcl_NewListObj(14, objv);
}

Tcl_Obj *ObjFromMIB_IPFORWARDTABLE(Tcl_Interp *interp, MIB_IPFORWARDTABLE *fwdP)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < fwdP->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_IPFORWARDROW(interp,
                                                             &fwdP->table[i])
            );
    }
    return resultObj;
}

Tcl_Obj *ObjFromIP_ADAPTER_INDEX_MAP(Tcl_Interp *interp, IP_ADAPTER_INDEX_MAP *iaimP)
{
    Tcl_Obj *objv[2];
    objv[0] = Tcl_NewIntObj(iaimP->Index);
    objv[1] = Tcl_NewUnicodeObj(iaimP->Name, -1);
    return Tcl_NewListObj(2, objv);
}

Tcl_Obj *ObjFromIP_INTERFACE_INFO(Tcl_Interp *interp, IP_INTERFACE_INFO *iiP)
{
    int i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < iiP->NumAdapters; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromIP_ADAPTER_INDEX_MAP(interp,
                                                                 &iiP->Adapter[i])
            );
    }
    return resultObj;
}


Tcl_Obj *ObjFromMIB_TCPROW(Tcl_Interp *interp, const MIB_TCPROW *row, int size)
{
    Tcl_Obj *obj[9];
    char     buf[sizeof(TCPIP_OWNER_MODULE_BASIC_INFO) + 4*MAX_PATH];
    DWORD    buf_sz;
    GetOwnerModuleFromTcpEntry_t  fn;

    obj[0] = Tcl_NewIntObj(row->dwState);
    obj[1] = IPAddrObjFromDWORD(row->dwLocalAddr);
    obj[2] = Tcl_NewIntObj(ntohs((WORD)row->dwLocalPort));
    obj[3] = IPAddrObjFromDWORD(row->dwRemoteAddr);
    obj[4] = Tcl_NewIntObj(ntohs((WORD)row->dwRemotePort));

    if (size < sizeof(MIB_TCPROW_OWNER_PID))
        return Tcl_NewListObj(5, obj);

    obj[5] = Tcl_NewIntObj(((MIB_TCPROW_OWNER_PID *)row)->dwOwningPid);

    if (size < sizeof(MIB_TCPROW_OWNER_MODULE))
        return Tcl_NewListObj(6, obj);

    obj[6] = Tcl_NewWideIntObj(((MIB_TCPROW_OWNER_MODULE *)row)->liCreateTimestamp.QuadPart);

    fn = Twapi_GetProc_GetOwnerModuleFromTcpEntry();
    buf_sz = sizeof(buf);
    if (fn && (*fn)((MIB_TCPROW_OWNER_MODULE *)row,
                    0, // TCPIP_OWNER_MODULE_BASIC_INFO enum
                    buf, &buf_sz) == NO_ERROR) {
        TCPIP_OWNER_MODULE_BASIC_INFO *modP;
        modP = (TCPIP_OWNER_MODULE_BASIC_INFO *) buf;
        obj[7] = Tcl_NewUnicodeObj(modP->pModuleName, -1);
        obj[8] = Tcl_NewUnicodeObj(modP->pModulePath, -1);
    } else {
        obj[7] = Tcl_NewStringObj("", -1);
        obj[8] = Tcl_NewStringObj("", -1);
    }

    return Tcl_NewListObj(9, obj);
}

int ObjToMIB_TCPROW(Tcl_Interp *interp, Tcl_Obj *listObj,
                    MIB_TCPROW *row)
{
    int  objc;
    Tcl_Obj **objv;

    if (Tcl_ListObjGetElements(interp, listObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc != 5) {
        if (interp)
            Tcl_AppendResult(interp, "Invalid TCP connection format: ",
                             Tcl_GetString(listObj),
                             NULL);
        return TCL_ERROR;
    }

    if ((Tcl_GetIntFromObj(interp, objv[0], &row->dwState) != TCL_OK) ||
        (IPAddrObjToDWORD(interp, objv[1], &row->dwLocalAddr) != TCL_OK) ||
        (Tcl_GetIntFromObj(interp, objv[2], &row->dwLocalPort) != TCL_OK) ||
        (IPAddrObjToDWORD(interp, objv[3], &row->dwRemoteAddr) != TCL_OK) ||
        (Tcl_GetIntFromObj(interp, objv[4], &row->dwRemotePort) != TCL_OK)) {
        /* interp already has error */
        return TCL_ERROR;
    }

    /* COnvert ports to network format */
    row->dwLocalPort = htons((short)row->dwLocalPort);
    row->dwRemotePort = htons((short)row->dwRemotePort);

    return TCL_OK;
}


Tcl_Obj *ObjFromMIB_UDPROW(Tcl_Interp *interp, MIB_UDPROW *row, int size)
{
    Tcl_Obj *obj[6];
    char     buf[sizeof(TCPIP_OWNER_MODULE_BASIC_INFO) + 4*MAX_PATH];
    DWORD    buf_sz;
    GetOwnerModuleFromUdpEntry_t fn;

    obj[0] = IPAddrObjFromDWORD(row->dwLocalAddr);
    obj[1] = Tcl_NewIntObj(ntohs((WORD)row->dwLocalPort));

    if (size < sizeof(MIB_UDPROW_OWNER_PID))
        return Tcl_NewListObj(2, obj);

    obj[2] = Tcl_NewIntObj(((MIB_UDPROW_OWNER_PID *)row)->dwOwningPid);

    if (size < sizeof(MIB_UDPROW_OWNER_MODULE))
        return Tcl_NewListObj(3, obj);

    obj[3] = Tcl_NewWideIntObj(((MIB_UDPROW_OWNER_MODULE *)row)->liCreateTimestamp.QuadPart);

    fn = Twapi_GetProc_GetOwnerModuleFromUdpEntry();
    buf_sz = sizeof(buf);
    if (fn && (*fn)((MIB_UDPROW_OWNER_MODULE *)row,
                    0, // TCPIP_OWNER_MODULE_BASIC_INFO enum
                    buf, &buf_sz) == NO_ERROR) {
        TCPIP_OWNER_MODULE_BASIC_INFO *modP;
        modP = (TCPIP_OWNER_MODULE_BASIC_INFO *) buf;
        obj[4] = Tcl_NewUnicodeObj(modP->pModuleName, -1);
        obj[5] = Tcl_NewUnicodeObj(modP->pModulePath, -1);
    } else {
        obj[4] = Tcl_NewStringObj("", -1);
        obj[5] = Tcl_NewStringObj("", -1);
    }

    return Tcl_NewListObj(6, obj);
}


Tcl_Obj *ObjFromMIB_TCPTABLE(Tcl_Interp *interp, MIB_TCPTABLE *tab)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < tab->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_TCPROW(interp, &(tab->table[i]), sizeof(MIB_TCPROW)));
    }

    return resultObj;
}

Tcl_Obj *ObjFromMIB_TCPTABLE_OWNER_PID(Tcl_Interp *interp, MIB_TCPTABLE_OWNER_PID *tab)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < tab->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_TCPROW(interp,  (MIB_TCPROW *) &(tab->table[i]), sizeof(MIB_TCPROW_OWNER_PID)));
    }

    return resultObj;
}

Tcl_Obj *ObjFromMIB_TCPTABLE_OWNER_MODULE(Tcl_Interp *interp, MIB_TCPTABLE_OWNER_MODULE *tab)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < tab->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_TCPROW(interp,  (MIB_TCPROW *) &(tab->table[i]), sizeof(MIB_TCPROW_OWNER_MODULE)));
    }

    return resultObj;
}


Tcl_Obj *ObjFromMIB_UDPTABLE(Tcl_Interp *interp, MIB_UDPTABLE *tab)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < tab->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_UDPROW(interp, &(tab->table[i]), sizeof(MIB_UDPROW)));
    }

    return resultObj;
}

Tcl_Obj *ObjFromMIB_UDPTABLE_OWNER_PID(Tcl_Interp *interp, MIB_UDPTABLE_OWNER_PID *tab)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < tab->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_UDPROW(interp, (MIB_UDPROW *) &(tab->table[i]), sizeof(MIB_UDPROW_OWNER_PID)));
    }

    return resultObj;
}


Tcl_Obj *ObjFromMIB_UDPTABLE_OWNER_MODULE(Tcl_Interp *interp, MIB_UDPTABLE_OWNER_MODULE *tab)
{
    DWORD i;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    for (i=0; i < tab->dwNumEntries; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromMIB_UDPROW(interp, (MIB_UDPROW *) &(tab->table[i]), sizeof(MIB_UDPTABLE_OWNER_MODULE)));
    }

    return resultObj;
}

Tcl_Obj *ObjFromTcpExTable(Tcl_Interp *interp, void *buf)
{
    /* It so happens the structure matches MIB_TCPTABLE_OWNER_PID */
    return ObjFromMIB_TCPTABLE_OWNER_PID(interp, (MIB_TCPTABLE_OWNER_PID *)buf);
}



Tcl_Obj *ObjFromUdpExTable(Tcl_Interp *interp, void *buf)
{
    /* It so happens the structure matches MIB_UDPTABLE_OWNER_PID */
    return ObjFromMIB_UDPTABLE_OWNER_PID(interp, (MIB_UDPTABLE_OWNER_PID *)buf);
}


int Twapi_FormatExtendedTcpTable(
    Tcl_Interp *interp,
    void *buf,
    int family,
    int table_class
    )
{
    Tcl_Obj *obj;

    if (family != AF_INET)
        return Twapi_AppendSystemError(interp, ERROR_INVALID_PARAMETER);

    // The constants below are TCP_TABLE_CLASS enumerations. These
    // is not defined in the SDK we are using so use integer constants
    switch (table_class) {
    case 0: // TCP_TABLE_BASIC_LISTENER
    case 1: // TCP_TABLE_BASIC_CONNECTIONS
    case 2: // TCP_TABLE_BASIC_ALL
        obj = ObjFromMIB_TCPTABLE(interp, (MIB_TCPTABLE *)buf);
        break;

    case 3: // TCP_TABLE_OWNER_PID_LISTENER
    case 4: // TCP_TABLE_OWNER_PID_CONNECTIONS
    case 5: // TCP_TABLE_OWNER_PID_ALL
        obj = ObjFromMIB_TCPTABLE_OWNER_PID(interp, (MIB_TCPTABLE_OWNER_PID *)buf);
        break;

    case 6: // TCP_TABLE_OWNER_MODULE_LISTENER
    case 7: // TCP_TABLE_OWNER_MODULE_CONNECTIONS
    case 8: // TCP_TABLE_OWNER_MODULE_ALL
        obj = ObjFromMIB_TCPTABLE_OWNER_MODULE(interp, (MIB_TCPTABLE_OWNER_MODULE *)buf);
        break;

    default:
        return Twapi_AppendSystemError(interp, ERROR_INVALID_PARAMETER);
    }

    Tcl_SetObjResult(interp, obj);
    return TCL_OK;
}

int Twapi_FormatExtendedUdpTable(
    Tcl_Interp *interp,
    void *buf,
    int family,
    int table_class
    )
{
    Tcl_Obj *obj;

    if (family != AF_INET)
        return Twapi_AppendSystemError(interp, ERROR_INVALID_PARAMETER);

    // The constants below are TCP_TABLE_CLASS enumerations. These
    // is not defined in the SDK we are using so use integer constants
    switch (table_class) {
    case 0: // UDP_TABLE_BASIC
        obj = ObjFromMIB_UDPTABLE(interp, (MIB_UDPTABLE *) buf);
        break;

    case 1: // UDP_TABLE_OWNER_PID
        obj = ObjFromMIB_UDPTABLE_OWNER_PID(interp, (MIB_UDPTABLE_OWNER_PID *) buf);
        break;

    case 2: // UDP_TABLE_OWNER_MODULE
        obj = ObjFromMIB_UDPTABLE_OWNER_MODULE(interp, (MIB_UDPTABLE_OWNER_MODULE *) buf);
        break;

    default:
        return Twapi_AppendSystemError(interp, ERROR_INVALID_PARAMETER);
    }

    Tcl_SetObjResult(interp, obj);
    return TCL_OK;
}



Tcl_Obj *ObjFromIP_ADAPTER_INFO(Tcl_Interp *interp, IP_ADAPTER_INFO *ainfoP)
{
    Tcl_Obj        *objv[14];

    objv[0] = Tcl_NewStringObj(ainfoP->AdapterName, -1);
    objv[1] = Tcl_NewStringObj(ainfoP->Description, -1);
    objv[2] = Tcl_NewByteArrayObj(ainfoP->Address, ainfoP->AddressLength);
    objv[3] = Tcl_NewIntObj(ainfoP->Index);
    objv[4] = Tcl_NewIntObj(ainfoP->Type);
    objv[5] = Tcl_NewIntObj(ainfoP->DhcpEnabled);
    objv[6] = ObjFromIP_ADDR_STRING(interp, &ainfoP->IpAddressList);
    objv[7] = Tcl_NewStringObj(ainfoP->GatewayList.IpAddress.String, -1);
    objv[8] = Tcl_NewStringObj(ainfoP->DhcpServer.IpAddress.String, -1);
    objv[9] = Tcl_NewIntObj(ainfoP->HaveWins);
    objv[10] = Tcl_NewStringObj(ainfoP->PrimaryWinsServer.IpAddress.String, -1);
    objv[11] = Tcl_NewStringObj(ainfoP->SecondaryWinsServer.IpAddress.String, -1);
    objv[12] = Tcl_NewWideIntObj(ainfoP->LeaseObtained);
    objv[13] = Tcl_NewWideIntObj(ainfoP->LeaseExpires);

    /* Attach to list of adapter data */
    return Tcl_NewListObj(14, objv);

}

Tcl_Obj *ObjFromIP_ADAPTER_INFO_table(Tcl_Interp *interp, IP_ADAPTER_INFO *ainfoP)
{
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);

    while (ainfoP) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromIP_ADAPTER_INFO(interp, ainfoP));
        ainfoP = ainfoP->Next;
    }

    return resultObj;
}


/* Helper function - common to all table retrieval functions */
static int TwapiIpConfigTableHelper(TwapiInterpContext *ticP, DWORD (FAR WINAPI *fn)(), Tcl_Obj *(*objbuilder)(Tcl_Interp *, ...), BOOL sortable, BOOL sort)
{
    int error;
    void *bufP;
    ULONG bufsz;
    int  tries;

    if (fn == NULL) {
        return Twapi_AppendSystemError(ticP->interp, ERROR_PROC_NOT_FOUND);
    }

    /*
     * Keep looping as long as we are told we need a bigger buffer.
     * For robustness, we set a limit on number of tries. Note required
     * size can keep changing so we try multiple times.
     */
    for (bufsz = 4000, tries=0; tries < 10 ; ++tries) {
        bufP = MemLifoPushFrame(&ticP->memlifo, bufsz, &bufsz);
        if (sortable)
            error = (*fn)(bufP, &bufsz, sort);
        else
            error = (*fn)(bufP, &bufsz);
        if (error != ERROR_INSUFFICIENT_BUFFER &&
            error != ERROR_BUFFER_OVERFLOW) {
            /* Either success or error unrelated to buffer size */
            break;
        }
        
        /* bufsz contains requried size as returned by the functions */

        MemLifoPopFrame(&ticP->memlifo);
    }

    if (error == NO_ERROR) {
        Tcl_SetObjResult(ticP->interp, (*objbuilder)(ticP->interp, bufP));
    } else {
        Twapi_AppendSystemError(ticP->interp, error);
    }

    MemLifoPopFrame(&ticP->memlifo);

    return error == NO_ERROR ? TCL_OK : TCL_ERROR;
}


int Twapi_GetNetworkParams(TwapiInterpContext *ticP)
{
    FIXED_INFO *netinfoP;
    ULONG netinfo_size;
    DWORD error;
    Tcl_Obj *objv[8];

    netinfoP = MemLifoPushFrame(&ticP->memlifo, sizeof(*netinfoP), &netinfo_size);
    error = GetNetworkParams(netinfoP, &netinfo_size);
    if (error == ERROR_BUFFER_OVERFLOW) {
        /* Allocate a bigger buffer of the required size. */
        MemLifoPopFrame(&ticP->memlifo);
        netinfoP = MemLifoPushFrame(&ticP->memlifo, netinfo_size, NULL);
        error = GetNetworkParams(netinfoP, &netinfo_size);
    }

    if (error == ERROR_SUCCESS) {
        objv[0] = Tcl_NewStringObj(netinfoP->HostName, -1);
        objv[1] = Tcl_NewStringObj(netinfoP->DomainName, -1);
        objv[2] = ObjFromIP_ADDR_STRINGAddress(ticP->interp,
                                               &netinfoP->DnsServerList);
        objv[3] = Tcl_NewIntObj(netinfoP->NodeType);
        objv[4] = Tcl_NewStringObj(netinfoP->ScopeId, -1);
        objv[5] = Tcl_NewIntObj(netinfoP->EnableRouting);
        objv[6] = Tcl_NewIntObj(netinfoP->EnableProxy);
        objv[7] = Tcl_NewIntObj(netinfoP->EnableDns);
        Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(8, objv));
    } else {
        Twapi_AppendSystemError(ticP->interp, error);
    }

    MemLifoPopFrame(&ticP->memlifo);

    return error == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
}


int Twapi_GetAdaptersInfo(TwapiInterpContext *ticP)
{
    return TwapiIpConfigTableHelper(
        ticP,
        GetAdaptersInfo,
        ObjFromIP_ADAPTER_INFO_table,
        0,
        0
        );
}


int Twapi_GetPerAdapterInfo(TwapiInterpContext *ticP, int adapter_index)
{
    IP_PER_ADAPTER_INFO *ainfoP;
    ULONG                ainfo_size;
    DWORD                error;
    Tcl_Obj             *objv[3];

    /* Make first allocation assuming two ip addresses */
    ainfoP = MemLifoPushFrame(&ticP->memlifo,
                              sizeof(*ainfoP)+2*sizeof(IP_ADDR_STRING),
                              &ainfo_size);
    error = GetPerAdapterInfo(adapter_index, ainfoP, &ainfo_size);
    if (error == ERROR_BUFFER_OVERFLOW) {
        /* Retry with indicated size */
        MemLifoPopFrame(&ticP->memlifo);
        ainfoP = MemLifoPushFrame(&ticP->memlifo, ainfo_size, NULL);
        error = GetPerAdapterInfo(adapter_index, ainfoP, &ainfo_size);
    }

    if (error == ERROR_SUCCESS) {
        objv[0] = Tcl_NewIntObj(ainfoP->AutoconfigEnabled);
        objv[1] = Tcl_NewIntObj(ainfoP->AutoconfigActive);
        objv[2] = ObjFromIP_ADDR_STRINGAddress(ticP->interp, &ainfoP->DnsServerList);
        Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(3, objv));
    } else
        Twapi_AppendSystemError(ticP->interp, error);

    MemLifoPopFrame(&ticP->memlifo);

    return error == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
}




int Twapi_GetInterfaceInfo(TwapiInterpContext *ticP)
{
    return TwapiIpConfigTableHelper(
        ticP,
        GetInterfaceInfo,
        ObjFromIP_INTERFACE_INFO,
        0,
        0
        );
}

int Twapi_GetIfEntry(Tcl_Interp *interp, int if_index)
{
    int error;
    MIB_IFROW ifr;

    ifr.dwIndex = if_index;
    error = GetIfEntry(&ifr);
    if (error) {
        return Twapi_AppendSystemError(interp, error);
    }
    Tcl_SetObjResult(interp, ObjFromMIB_IFROW(interp, &ifr));
    return TCL_OK;
}

int Twapi_GetIfTable(TwapiInterpContext *ticP, int sort)
{
    return TwapiIpConfigTableHelper(
        ticP,
        GetIfTable,
        ObjFromMIB_IFTABLE,
        1,
        sort
        );
}

int Twapi_GetIpAddrTable(TwapiInterpContext *ticP, int sort)
{
    return TwapiIpConfigTableHelper(
        ticP,
        GetIpAddrTable,
        ObjFromMIB_IPADDRTABLE,
        1,
        sort
        );
}


int Twapi_GetIpNetTable(TwapiInterpContext *ticP, int sort)
{
    return TwapiIpConfigTableHelper(
        ticP,
        GetIpNetTable,
        ObjFromMIB_IPNETTABLE,
        1,
        sort
        );
}

int Twapi_GetIpForwardTable(TwapiInterpContext *ticP, int sort)
{
    return TwapiIpConfigTableHelper(
        ticP,
        GetIpForwardTable,
        ObjFromMIB_IPFORWARDTABLE,
        1,
        sort
        );
}

typedef DWORD (WINAPI *GetExtendedTcpTable_t)(PVOID, PDWORD, BOOL, ULONG, ULONG, ULONG);
MAKE_DYNLOAD_FUNC(GetExtendedTcpTable, iphlpapi, GetExtendedTcpTable_t)
int Twapi_GetExtendedTcpTable(
    Tcl_Interp *interp,
    void *buf,
    DWORD buf_sz,
    BOOL sorted,
    ULONG family,
    int   table_class
    )
{
    int error;
    GetExtendedTcpTable_t fn = Twapi_GetProc_GetExtendedTcpTable();
    if (fn == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    error = (*fn)(buf, &buf_sz, sorted, family, table_class, 0);
    if (error == NO_ERROR || error == ERROR_INSUFFICIENT_BUFFER) {
        Tcl_SetObjResult(interp, Tcl_NewIntObj(buf_sz));
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, error);
    }
}


typedef DWORD (WINAPI *GetExtendedUdpTable_t)(PVOID, PDWORD, BOOL, ULONG, ULONG, ULONG);
MAKE_DYNLOAD_FUNC(GetExtendedUdpTable, iphlpapi, GetExtendedUdpTable_t)
int Twapi_GetExtendedUdpTable(
    Tcl_Interp *interp,
    void *buf,
    DWORD buf_sz,
    BOOL sorted,
    ULONG family,
    int   table_class
    )
{
    int error;
    GetExtendedUdpTable_t fn = Twapi_GetProc_GetExtendedUdpTable();
    if (fn == NULL) {
        return Twapi_AppendSystemError(interp, ERROR_PROC_NOT_FOUND);
    }

    error = (*fn)(buf, &buf_sz, sorted, family, table_class, 0);
    if (error == NO_ERROR || error == ERROR_INSUFFICIENT_BUFFER) {
        Tcl_SetObjResult(interp, Tcl_NewIntObj(buf_sz));
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(interp, error);
    }
}


typedef DWORD (WINAPI *AllocateAndGetTcpExTableFromStack_t)(PVOID *,BOOL,HANDLE,DWORD, DWORD);
MAKE_DYNLOAD_FUNC(AllocateAndGetTcpExTableFromStack, iphlpapi, AllocateAndGetTcpExTableFromStack_t)
int Twapi_AllocateAndGetTcpExTableFromStack(
    TwapiInterpContext *ticP,
    BOOL sorted,
    DWORD flags
)
{
    int error;
    AllocateAndGetTcpExTableFromStack_t fn = Twapi_GetProc_AllocateAndGetTcpExTableFromStack();

    if (fn) {
        void *buf = NULL;

        /* 2 -> AF_INET (IP v4) */
        error = (*fn)(&buf, sorted, GetProcessHeap(), flags, 2);
        if (error)
            return Twapi_AppendSystemError(ticP->interp, error);

        Tcl_SetObjResult(ticP->interp, ObjFromTcpExTable(ticP->interp, buf));
        HeapFree(GetProcessHeap(), 0, buf);
        return TCL_OK;
    } else {
        DWORD sz;
        MIB_TCPTABLE *tab = NULL;
        int i;

        /*
         * First get the required  buffer size.
         * Do this in a loop since size might change with an upper limit
         * on number of iterations.
         * We do not use MemLifo because allocations are likely quite large
         * so little benefit.
         */
        for (tab = NULL, sz = 0, i = 0; i < 10; ++i) {
            error = GetTcpTable(tab, &sz, sorted);
            if (error != ERROR_INSUFFICIENT_BUFFER)
                break;
            /* Retry with larger buffer */
            if (tab)
                TwapiFree(tab);
            tab = (MIB_TCPTABLE *) TwapiAlloc(sz);
        }
        
        if (error == ERROR_SUCCESS)
            Tcl_SetObjResult(ticP->interp, ObjFromMIB_TCPTABLE(ticP->interp, tab));
        else
            Twapi_AppendSystemError(ticP->interp, error);
        if (tab)
            TwapiFree(tab);
        return error == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
    }
}


typedef DWORD (WINAPI *AllocateAndGetUdpExTableFromStack_t)(PVOID *,BOOL,HANDLE,DWORD, DWORD);
MAKE_DYNLOAD_FUNC(AllocateAndGetUdpExTableFromStack, iphlpapi, AllocateAndGetUdpExTableFromStack_t)
int Twapi_AllocateAndGetUdpExTableFromStack(
    TwapiInterpContext *ticP,
    BOOL sorted,
    DWORD flags
)
{
    int error;
    AllocateAndGetUdpExTableFromStack_t fn = Twapi_GetProc_AllocateAndGetUdpExTableFromStack();

    if (fn) {
        void *buf = NULL;

        /* 2 -> AF_INET (IP v4) */
        error = (*fn)(&buf, sorted, GetProcessHeap(), flags, 2);
        if (error)
            return Twapi_AppendSystemError(ticP->interp, error);

        Tcl_SetObjResult(ticP->interp, ObjFromUdpExTable(ticP->interp, buf));
        HeapFree(GetProcessHeap(), 0, buf);
        return TCL_OK;
    } else {
        DWORD sz;
        MIB_UDPTABLE *tab = NULL;
        int i;

        /*
         * First get the required  buffer size.
         * Do this in a loop since size might change with an upper limit
         * on number of iterations.
         * We do not use MemLifo because allocations are likely quite large
         * so little benefit.
         */
        for (tab = NULL, sz = 0, i = 0; i < 10; ++i) {
            error = GetUdpTable(tab, &sz, sorted);
            if (error != ERROR_INSUFFICIENT_BUFFER)
                break;
            /* Retry with larger buffer */
            if (tab)
                TwapiFree(tab);
            tab = (MIB_UDPTABLE *) TwapiAlloc(sz);
        }
        
        if (error == ERROR_SUCCESS)
            Tcl_SetObjResult(ticP->interp, ObjFromMIB_UDPTABLE(ticP->interp, tab));
        else
            Twapi_AppendSystemError(ticP->interp, error);
        if (tab)
            TwapiFree(tab);
        return error == ERROR_SUCCESS ? TCL_OK : TCL_ERROR;
    }
}


int Twapi_GetNameInfo(
    Tcl_Interp *interp,
    const struct sockaddr_in* saP,
    int flags
    )
{
    int status;
    char hostname[NI_MAXHOST];
    char portname[NI_MAXSERV];
    Tcl_Obj *objv[2];

    status = getnameinfo((struct sockaddr *) saP, sizeof(*saP),
                         hostname, sizeof(hostname)/sizeof(hostname[0]),
                         portname, sizeof(portname)/sizeof(portname[0]),
                         flags);
    if (status != 0) {
        return Twapi_AppendSystemError(interp, status);
    }

    objv[0] = Tcl_NewStringObj(hostname, -1);
    objv[1] = Tcl_NewStringObj(portname, -1);

    Tcl_SetObjResult(interp, Tcl_NewListObj(2,objv));
    return TCL_OK;
}

int Twapi_GetAddrInfo(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    const char *hostname;
    const char *svcname;
    int status;
    struct addrinfo hints;
    struct addrinfo *addrP;
    struct addrinfo *saved_addrP;
    Tcl_Obj *resultObj;

    ZeroMemory(&hints, sizeof(hints));
    hints.ai_family = PF_INET;
    if (TwapiGetArgs(interp, objc, objv,
                     GETASTR(hostname), GETASTR(svcname),
                     GETINT(hints.ai_protocol),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    status = getaddrinfo(hostname, svcname, &hints, &addrP);
    if (status != 0) {
        return Twapi_AppendSystemError(interp, status);
    }

    resultObj = Tcl_NewListObj(0, NULL);
    saved_addrP = addrP;
    while (addrP) {
        Tcl_Obj *objv[2];
        struct sockaddr_in *saddrP = (struct sockaddr_in *)addrP->ai_addr;

        if (addrP->ai_family != PF_INET ||
            addrP->ai_addrlen != sizeof(struct sockaddr_in) ||
            saddrP->sin_family != AF_INET) {
            /* Not IP V4 */
            continue;
        }
        objv[0] = Tcl_NewStringObj(inet_ntoa(saddrP->sin_addr), -1);
        objv[1] = Tcl_NewIntObj((unsigned short) ntohs(saddrP->sin_port));
        Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewListObj(2, objv));

        addrP = addrP->ai_next;
    }
    if (saved_addrP)
        freeaddrinfo(saved_addrP);

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}

int Twapi_GetBestRoute(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    MIB_IPFORWARDROW route;
    int error;
    DWORD dest, src;

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETVAR(dest, IPAddrObjToDWORD),
                     GETVAR(src, IPAddrObjToDWORD),
                     ARGEND) != TCL_OK) {
        return TCL_ERROR;
    }

    error = GetBestRoute(dest, src, &route);
    if (error == NO_ERROR) {
        Tcl_SetObjResult(ticP->interp, ObjFromMIB_IPFORWARDROW(ticP->interp, &route));
        return TCL_OK;
    } else {
        return Twapi_AppendSystemError(ticP->interp, error);
    }
}
