/*
 * Copyright (c) 2003-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to Windows API related to security and access control functions */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

/*
TBD - get_wellknown_sid - returns well known sid's such as administrators group
TBD - the SID data type handling needs to be cleaned up
 */

int ObjToSID_AND_ATTRIBUTES(Tcl_Interp *interp, Tcl_Obj *obj, SID_AND_ATTRIBUTES *sidattrP);
Tcl_Obj *ObjFromSID_AND_ATTRIBUTES (Tcl_Interp *, const SID_AND_ATTRIBUTES *);
Tcl_Obj *ObjFromLUID_AND_ATTRIBUTES (Tcl_Interp *, const LUID_AND_ATTRIBUTES *);
int ObjToLUID_AND_ATTRIBUTES (Tcl_Interp *interp, Tcl_Obj *listobj,
                              LUID_AND_ATTRIBUTES *luidattrP);
int ObjToPTOKEN_PRIVILEGES(Tcl_Interp *interp,
                          Tcl_Obj *tokprivObj, TOKEN_PRIVILEGES **tokprivPP);
Tcl_Obj *ObjFromTOKEN_PRIVILEGES(Tcl_Interp *interp,
                                 const TOKEN_PRIVILEGES *tokprivP);
void TwapiFreeTOKEN_PRIVILEGES (TOKEN_PRIVILEGES *tokPrivP);


/*
 * Allocate and return memory for a TOKEN_PRIVILEGES structure big enough
 * to hold num_privs privileges
 */
static int TwapiAllocateTOKEN_PRIVILEGES (Tcl_Interp *interp, int num_privs, TOKEN_PRIVILEGES **tpPP)
{
    *tpPP = TwapiAlloc(sizeof(TOKEN_PRIVILEGES)
                       + (num_privs*sizeof((*tpPP)->Privileges[0]))
                       - sizeof((*tpPP)->Privileges));

    (*tpPP)->PrivilegeCount = num_privs;
    return TCL_OK;
}


void TwapiFreeTOKEN_PRIVILEGES (TOKEN_PRIVILEGES *tokPrivP)
{
    if (tokPrivP)
        TwapiFree(tokPrivP);
}

/* interp may be NULL */
Tcl_Obj *ObjFromSID_AND_ATTRIBUTES (
    Tcl_Interp *interp, const SID_AND_ATTRIBUTES *sidattrP
)
{
    Tcl_Obj *objv[2];

    if (ObjFromSID(interp, sidattrP->Sid, &objv[0]) != TCL_OK) {
        return NULL;
    }

    objv[1] = Tcl_NewLongObj(sidattrP->Attributes);
    return Tcl_NewListObj(2, objv);
}

Tcl_Obj *ObjFromSID_AND_ATTRIBUTES_Array (
    Tcl_Interp *interp, const SID_AND_ATTRIBUTES *sidattrP, int count
)
{
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);
    Tcl_Obj *obj;
    int i;
    for (i = 0; i < count; ++i, ++sidattrP) {
        obj = ObjFromSID_AND_ATTRIBUTES(interp, sidattrP);
        if (obj)
            Tcl_ListObjAppendElement(interp, resultObj, obj);
        else {
            Tcl_DecrRefCount(resultObj);
            return NULL;
        }
    }

    return resultObj;
}


/* Note sidattrP->Sid is dynamically allocated and must be freed by calling TwapiFree(). */
int ObjToSID_AND_ATTRIBUTES(Tcl_Interp *interp, Tcl_Obj *obj, SID_AND_ATTRIBUTES *sidattrP)
{
    Tcl_Obj **objv;
    int       objc;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) == TCL_OK &&
        objc == 2 &&
        Tcl_GetLongFromObj(interp, objv[1], &sidattrP->Attributes) == TCL_OK &&
        ObjToPSID(interp, objv[0], &sidattrP->Sid) == TCL_OK) {

        return TCL_OK;
    } else
        return TCL_ERROR;
}


/* interp may be NULL */
Tcl_Obj *ObjFromLUID_AND_ATTRIBUTES (
    Tcl_Interp *interp, const LUID_AND_ATTRIBUTES *luidattrP
)
{
    Tcl_Obj *objv[2];

    objv[0] = ObjFromLUID(&luidattrP->Luid);
    /* Create as wide int to try and preserve sign */
    objv[1] = Tcl_NewWideIntObj((ULONG_PTR)luidattrP->Attributes);
    return Tcl_NewListObj(2, objv);
}

/* interp may be NULL */
int ObjToLUID_AND_ATTRIBUTES (
    Tcl_Interp *interp, Tcl_Obj *listobj, LUID_AND_ATTRIBUTES *luidattrP
)
{
    int       objc;
    Tcl_Obj **objv;

    if (Tcl_ListObjGetElements(interp, listobj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if ((objc != 2) ||
        (Tcl_GetLongFromObj(interp,objv[1],&luidattrP->Attributes) != TCL_OK) ||
        (ObjToLUID(interp, objv[0], &luidattrP->Luid) != TCL_OK)) {
        if (interp) {
            Tcl_AppendResult(interp,
                             "\nInvalid LUID_AND_ATTRIBUTES: ",
                             Tcl_GetString(listobj),
                             NULL);
        }
        return TCL_ERROR;
    }

    return TCL_OK;
}



/* interp may be NULL */
static Tcl_Obj *ObjFromLUID_AND_ATTRIBUTES_Array(
    Tcl_Interp *interp, const LUID_AND_ATTRIBUTES *luidattrP, int count
)
{
    int      i;
    Tcl_Obj *obj;
    Tcl_Obj *resultObj = Tcl_NewListObj(0, NULL);
    for (i = 0; i < count; ++i) {
        obj = ObjFromLUID_AND_ATTRIBUTES(interp, luidattrP + i);
        if (obj)
            Tcl_ListObjAppendElement(interp, resultObj, obj);
        else {
            Twapi_FreeNewTclObj(resultObj);
            return NULL;
        }
    }

    return resultObj;
}


/* interp may be NULL */
Tcl_Obj *ObjFromTOKEN_PRIVILEGES(
    Tcl_Interp *interp, const TOKEN_PRIVILEGES *tokprivP
)
{
    return ObjFromLUID_AND_ATTRIBUTES_Array(interp,
                                                &tokprivP->Privileges[0],
                                                tokprivP->PrivilegeCount);
}


/* Returned memory should be freed with TwapiFreeTOKEN_PRIVILEGES.
   interp may be NULL */
int ObjToPTOKEN_PRIVILEGES(
    Tcl_Interp *interp, Tcl_Obj *tokprivObj, TOKEN_PRIVILEGES **tokprivPP
)
{
    int       objc;
    Tcl_Obj **objv;

    if (Tcl_ListObjGetElements(interp, tokprivObj, &objc, &objv) != TCL_OK ||
        TwapiAllocateTOKEN_PRIVILEGES(interp, objc, tokprivPP) != TCL_OK)
        return TCL_ERROR;

    (*tokprivPP)->PrivilegeCount = objc;
    while (objc--) {
        if (ObjToLUID_AND_ATTRIBUTES(interp, objv[objc],
                                         &(*tokprivPP)->Privileges[objc])
            == TCL_ERROR) {

            TwapiFreeTOKEN_PRIVILEGES(*tokprivPP);
            *tokprivPP = NULL;
            return TCL_ERROR;
        }
    }

    return TCL_OK;
}



/* Generic routine to retrieve token information - returns a dynamic alloc block */
static void *AllocateAndGetTokenInformation(HANDLE tokenH,
                                            TOKEN_INFORMATION_CLASS class)
{
    void *infoP = NULL;
    int   info_buf_sz = 0;
    DWORD error = 0;

    /* Keep looping since required buffer size may keep changing */
    do {
        if (GetTokenInformation(tokenH, (TOKEN_INFORMATION_CLASS) class,
                                infoP, info_buf_sz, &info_buf_sz)) {
            /* Got what we wanted */
            return infoP;
        }
        if ((error = GetLastError()) != ERROR_INSUFFICIENT_BUFFER) {
            break;
        }
        /* Need a bigger buffer */
        if (infoP)
            TwapiFree(infoP);
        infoP = TwapiAlloc(info_buf_sz);
    } while (infoP);

    /*
     * Some error occurred. Either no memory (infoP==NULL) or something
     * else. Variable error is already set above
     */
    if (infoP)
        TwapiFree(infoP);
    SetLastError(error);
    return NULL;
}


DWORD Twapi_PrivilegeCheck(
    HANDLE            tokenH,
    const TOKEN_PRIVILEGES *tokprivP,
    int               all_required,
    int              *resultP
    )
{
    PRIVILEGE_SET *privSet;
    int num_privs = tokprivP->PrivilegeCount;
    BOOL have_priv;
    BOOL success;
    int  sz;

    sz = sizeof(PRIVILEGE_SET)
        + (num_privs*sizeof(privSet->Privilege[0]))
        - sizeof(privSet->Privilege);
    privSet = TwapiAlloc(sz);

    privSet->PrivilegeCount = num_privs;
    privSet->Control = all_required ? PRIVILEGE_SET_ALL_NECESSARY : 0;
    while (num_privs--) {
        privSet->Privilege[num_privs] = tokprivP->Privileges[num_privs];
    }

    success = PrivilegeCheck(tokenH, privSet, &have_priv);
    if (success)
        *resultP = have_priv;

    TwapiFree(privSet);

    return success;
}


int Twapi_GetTokenInformation(
    Tcl_Interp *interp,
    HANDLE      tokenH,
    TOKEN_INFORMATION_CLASS token_class
    )
{
    void *infoP;
    Tcl_Obj *resultObj = NULL;
    Tcl_Obj *obj = NULL;
    DWORD    sz;
    int   result = TCL_ERROR;
    union {
        DWORD dw;
        TOKEN_TYPE type;
        SECURITY_IMPERSONATION_LEVEL sil;
        TWAPI_TOKEN_ELEVATION elevation;
        TWAPI_TOKEN_ELEVATION_TYPE elevation_type;
        TWAPI_TOKEN_LINKED_TOKEN linked_token;
    } value;

    switch (token_class) {
    case TwapiTokenElevation:
    case TwapiTokenElevationType:
    case TwapiTokenMandatoryPolicy:
    case TwapiTokenHasRestrictions:
    case TwapiTokenLinkedToken:
    case TokenImpersonationLevel:
    case TokenSandBoxInert:
    case TokenSessionId:
    case TokenType:
    case TwapiTokenUIAccess:
    case TwapiTokenVirtualizationAllowed:
    case TwapiTokenVirtualizationEnabled:
        
        sz = sizeof(value);
        if (GetTokenInformation(tokenH, token_class, &value, sz, &sz)) {
            switch (token_class) {
            case TwapiTokenLinkedToken:
                resultObj = ObjFromHANDLE(value.linked_token.LinkedToken);
                result = TCL_OK;
                break;

            case TwapiTokenMandatoryPolicy: /* FALLTHRU */
            case TwapiTokenElevation:
                resultObj = Tcl_NewIntObj(value.elevation.TokenIsElevated);
                result = TCL_OK;
                break;

            case TwapiTokenElevationType:
                resultObj = Tcl_NewIntObj(value.elevation_type);
                result = TCL_OK;
                break;

            case TokenImpersonationLevel:
                resultObj = Tcl_NewIntObj(value.sil);
                result = TCL_OK;
                break;

            case TwapiTokenHasRestrictions:
            case TwapiTokenUIAccess:
            case TwapiTokenVirtualizationAllowed:
            case TwapiTokenVirtualizationEnabled:
            case TokenSandBoxInert:
            case TokenSessionId:
                resultObj = Tcl_NewIntObj(value.dw);
                result = TCL_OK;
                break;
                
            case TokenType:
                resultObj = Tcl_NewIntObj(value.type);
                result = TCL_OK;
                break;

            }

            /*
             * At this point, if result is TCL_OK, resultObj should contain
             * the result to be returned. Else interp->result should
             * contain the error message. In this
             * case, resultObj will be freed if non-NULL
             */
            if (result == TCL_OK)
                Tcl_SetObjResult(interp, resultObj);
            else if (resultObj)
                Twapi_FreeNewTclObj(resultObj);

            return result;
        }
        infoP = NULL;           /* To hit error handler below */
        break;

    default:
        /* Other classes are variable size */
        infoP = AllocateAndGetTokenInformation(tokenH, token_class);
        break;
    }

    if (infoP == NULL) {
        Tcl_SetResult(interp, "Error getting security token information: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, GetLastError());
        return TCL_ERROR;
    }

    switch (token_class) {
    case TokenOwner:
        result = ObjFromSID(interp,
                              ((TOKEN_OWNER *)infoP)->Owner,
                              &resultObj);
        break;

    case TokenPrimaryGroup:
        result = ObjFromSID(interp,
                              ((TOKEN_PRIMARY_GROUP *)infoP)->PrimaryGroup,
                              &resultObj);
        break;

    case TwapiTokenIntegrityLevel: /* FALLTHRU */
    case TokenUser:
        resultObj=ObjFromSID_AND_ATTRIBUTES(interp, &((TOKEN_USER *) infoP)->User);
        if (resultObj)
            result = TCL_OK;
        break;

    case TokenRestrictedSids: /* Fall through */
    case TokenGroups:
        resultObj = ObjFromSID_AND_ATTRIBUTES_Array(
            interp,
            &((TOKEN_GROUPS *)infoP)->Groups[0],
            ((TOKEN_GROUPS *)infoP)->GroupCount);
        if (resultObj)
            result = TCL_OK;
        break;

    case TokenPrivileges:
        resultObj = ObjFromTOKEN_PRIVILEGES(interp,
                                                (TOKEN_PRIVILEGES *)infoP);
        if (resultObj)
            result = TCL_OK;

        break;

    case TokenGroupsAndPrivileges:
        resultObj = Tcl_NewListObj(0, NULL);
        obj = ObjFromSID_AND_ATTRIBUTES_Array(
            interp,
            ((TOKEN_GROUPS_AND_PRIVILEGES *)infoP)->Sids,
            ((TOKEN_GROUPS_AND_PRIVILEGES *)infoP)->SidCount);
        if (obj) {
            Tcl_ListObjAppendElement(interp, resultObj, obj);
            obj = ObjFromSID_AND_ATTRIBUTES_Array(
                interp,
                ((TOKEN_GROUPS_AND_PRIVILEGES *)infoP)->RestrictedSids,
                ((TOKEN_GROUPS_AND_PRIVILEGES *)infoP)->RestrictedSidCount);
            if (obj) {
                Tcl_ListObjAppendElement(interp, resultObj, obj);
                obj = ObjFromLUID_AND_ATTRIBUTES_Array(
                    interp, 
                    ((TOKEN_GROUPS_AND_PRIVILEGES *)infoP)->Privileges,
                    ((TOKEN_GROUPS_AND_PRIVILEGES *)infoP)->PrivilegeCount);
                if (obj) {
                    Tcl_ListObjAppendElement(interp, resultObj, obj);
                    obj = ObjFromLUID(&((TOKEN_GROUPS_AND_PRIVILEGES *)infoP)->AuthenticationId);
                    if (obj) {
                        Tcl_ListObjAppendElement(interp, resultObj, obj);
                        result = TCL_OK;
                    }
                }
            }
        }

        break;

    case TokenSource:
        resultObj = Tcl_NewListObj(0, NULL);
        obj = Tcl_NewStringObj(((TOKEN_SOURCE *)infoP)->SourceName, 8);
        Tcl_ListObjAppendElement(interp, resultObj, obj);
        obj =
            ObjFromLUID(&((TOKEN_SOURCE *)infoP)->SourceIdentifier);
        if (obj) {
            Tcl_ListObjAppendElement(interp, resultObj, obj);
            result = TCL_OK;
        } else {
            Tcl_SetResult(interp, "Could not convert token source to LUID", TCL_STATIC);
        }
        break;

    case TokenStatistics:
        resultObj = Tcl_NewListObj(0, NULL);
        obj = ObjFromLUID(&((TOKEN_STATISTICS *)infoP)->TokenId);
        if (obj == NULL)
            break;
        Tcl_ListObjAppendElement(interp, resultObj, obj);
        obj = ObjFromLUID(&((TOKEN_STATISTICS *)infoP)->AuthenticationId);
        if (obj == NULL)
            break;
        Tcl_ListObjAppendElement(interp, resultObj, obj);
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewWideIntObj(((TOKEN_STATISTICS *)infoP)->ExpirationTime.QuadPart));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(((TOKEN_STATISTICS *)infoP)->TokenType));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(((TOKEN_STATISTICS *)infoP)->ImpersonationLevel));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(((TOKEN_STATISTICS *)infoP)->DynamicCharged));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(((TOKEN_STATISTICS *)infoP)->DynamicAvailable));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(((TOKEN_STATISTICS *)infoP)->GroupCount));
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(((TOKEN_STATISTICS *)infoP)->PrivilegeCount));
        obj = ObjFromLUID(&((TOKEN_STATISTICS *)infoP)->ModifiedId);
        if (obj == NULL)
            break;
        Tcl_ListObjAppendElement(interp, resultObj, obj);
        result = TCL_OK;
        break;

#ifdef TBD
        Only implemented in W2K3
    case TokenOrigin:
        resultObj = ObjFromLUID(
            &((TOKEN_ORIGIN *)infoP)->OriginatingLogonSession );
        result = TCL_OK;
        break;
#endif

    case TokenSessionReference:
    case TokenDefaultDacl:
        Tcl_SetResult(interp, "Unsupported token information type", TCL_STATIC);
        break;

    default:
        Tcl_SetResult(interp, "Unknown token information type", TCL_STATIC);
        break;
    }

    /*
     * At this point,
     *  If result is TCL_OK, resultObj should contain the result to be returned
     *  Otherwise, interp->result should contain the error message. In this
     *  case, resultObj will be freed if non-NULL
     */
    if (result == TCL_OK)
        Tcl_SetObjResult(interp, resultObj);
    else if (resultObj)
        Twapi_FreeNewTclObj(resultObj);

    if (infoP)
        TwapiFree(infoP);

    return result;
}

int Twapi_AdjustTokenPrivileges(
    TwapiInterpContext *ticP,
    HANDLE tokenH,
    BOOL   disableAll,
    TOKEN_PRIVILEGES *tokprivP
)
{
    BOOL ret;
    Tcl_Obj *objP;
    Tcl_Interp *interp = ticP->interp;
    DWORD buf_size = 128;
    void *bufP;
    WIN32_ERROR winerr;
    TCL_RESULT tcl_result = TCL_ERROR;

    bufP = MemLifoPushFrame(&ticP->memlifo, buf_size, &buf_size);
    ret = AdjustTokenPrivileges(tokenH, disableAll, tokprivP,
                                buf_size,
                                (PTOKEN_PRIVILEGES) bufP, &buf_size);
    if (!ret) {
        winerr = GetLastError();
        if (winerr == ERROR_INSUFFICIENT_BUFFER) {
            /* Retry with larger buffer */
            bufP = MemLifoAlloc(&ticP->memlifo, buf_size, NULL);
            ret = AdjustTokenPrivileges(tokenH, disableAll, tokprivP,
                                        buf_size,
                                        (PTOKEN_PRIVILEGES) bufP, &buf_size);
        }
        if (!ret) {
            /* No joy.*/
            winerr = GetLastError();
            MemLifoPopFrame(&ticP->memlifo);
            return Twapi_AppendSystemError(interp, winerr);
        }
    }

    /*
     * Even if ret indicates success, we still need to check GetLastError()
     * (see SDK).
     */
    if (GetLastError() == ERROR_NOT_ALL_ASSIGNED) {
        /* Revert to previous privs. */
        AdjustTokenPrivileges(tokenH, disableAll,
                              (PTOKEN_PRIVILEGES) bufP,
                              0, NULL, NULL);
        Twapi_AppendSystemError(interp, ERROR_NOT_ALL_ASSIGNED);
    } else {
        objP = ObjFromTOKEN_PRIVILEGES(interp, (PTOKEN_PRIVILEGES) bufP);
        if (objP) {
            Tcl_SetObjResult(interp, objP);
            tcl_result = TCL_OK;
        }
        else {
            /* interp->result should already contain the error */
        }
    }
    MemLifoPopFrame(&ticP->memlifo);
    return tcl_result;
}


int Twapi_InitializeSecurityDescriptor(Tcl_Interp *interp)
{
    SECURITY_DESCRIPTOR secd;
    Tcl_Obj *resultObj;

    if (! InitializeSecurityDescriptor(&secd, SECURITY_DESCRIPTOR_REVISION)) {
        return TwapiReturnSystemError(interp);
    }

    resultObj = ObjFromSECURITY_DESCRIPTOR(interp, &secd);
    if (resultObj == NULL)
        return TCL_ERROR;

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}

#ifndef TWAPI_LEAN
int Twapi_GetNamedSecurityInfo (
    Tcl_Interp *interp,
    LPWSTR name,
    int type,
    int wanted_fields
)
{
    DWORD error;
    PSID ownerP;
    PSID groupP;
    PACL daclP;
    PACL saclP;
    PSECURITY_DESCRIPTOR secdP;
    Tcl_Obj *resultObj;

    error = GetNamedSecurityInfoW(name, type, wanted_fields,
                                  &ownerP, &groupP, &daclP, &saclP, &secdP);
    if (error) {
        TwapiReturnSystemError(interp);
        return TCL_ERROR;
    }

    resultObj = ObjFromSECURITY_DESCRIPTOR(interp, secdP);
    LocalFree(secdP);
    if (resultObj)
        Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;

}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int Twapi_GetSecurityInfo (
    Tcl_Interp *interp,
    HANDLE h,
    int type,
    int wanted_fields
)
{
    DWORD error;
    PSID ownerP;
    PSID groupP;
    PACL daclP;
    PACL saclP;
    PSECURITY_DESCRIPTOR secdP;
    Tcl_Obj *resultObj;

    error = GetSecurityInfo(h, type, wanted_fields,
                            &ownerP, &groupP, &daclP, &saclP, &secdP);
    if (error) {
        TwapiReturnSystemError(interp);
        return TCL_ERROR;
    }

    resultObj = ObjFromSECURITY_DESCRIPTOR(interp, secdP);
    LocalFree(secdP);
    if (resultObj)
        Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
// TBD - should this be in account.c ?
int Twapi_LsaEnumerateAccountRights(
    Tcl_Interp *interp,
    LSA_HANDLE policyH,
    PSID sidP
)
{
    LSA_UNICODE_STRING *rightsP;
    NTSTATUS ntstatus;
    ULONG i, count;
    Tcl_Obj *resultObj;
    int result = TCL_ERROR;

    ntstatus = LsaEnumerateAccountRights(policyH, sidP, &rightsP, &count);
    if (ntstatus != STATUS_SUCCESS) {
        Tcl_SetResult(interp, "Could not enumerate account rights: ",
                      TCL_STATIC);
        return Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(ntstatus));
    }

    resultObj = Tcl_NewListObj(0, NULL);
    for (i = 0; i < count; ++i) {
        result = Tcl_ListObjAppendElement(
            interp,
            resultObj,
            ObjFromLSA_UNICODE_STRING(&rightsP[i])
            );
        if (result != TCL_OK)
            break;              /* Break loop on error */
    }
    LsaFreeMemory(rightsP);

    if (result == TCL_OK)
        Tcl_SetObjResult(interp, resultObj);

    return result;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
// TBD - should this be in account.c ?
int Twapi_LsaEnumerateAccountsWithUserRight(
    Tcl_Interp *interp,
    LSA_HANDLE policyH,
    LSA_UNICODE_STRING *rightP
)
{
    LSA_ENUMERATION_INFORMATION *buf;
    ULONG i, count;
    NTSTATUS ntstatus;
    Tcl_Obj *resultObj;
    int result = TCL_ERROR;

    ntstatus = LsaEnumerateAccountsWithUserRight(policyH, rightP,
                                                 (void **) &buf, &count);
    if (ntstatus != STATUS_SUCCESS) {
        Tcl_SetResult(interp, "Could not enumerate accounts with specified privileges: ",
                      TCL_STATIC);
        return Twapi_AppendSystemError(interp, TwapiNTSTATUSToError(ntstatus));
    }

    resultObj = Tcl_NewListObj(0, NULL);
    for (i = 0; i < count; ++i) {
        Tcl_Obj *objP;
        result = ObjFromSID(interp, buf[i].Sid, &objP);
        if (result != TCL_OK)
            break;              /* Break loop on error */

        result = Tcl_ListObjAppendElement(interp, resultObj, objP);
        if (result != TCL_OK)
            break;              /* Break loop on error */
    }
    LsaFreeMemory(buf);

    if (result == TCL_OK)
        Tcl_SetObjResult(interp, resultObj);

    return result;


}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
/* Note LsaEnumerateLogonSessions is present on Win2k too although not documented*/
int Twapi_LsaEnumerateLogonSessions(Tcl_Interp *interp)
{
    ULONG i,count;
    PLUID luids;
    NTSTATUS status;
    Tcl_Obj *resultObj;

    status = LsaEnumerateLogonSessions(&count, &luids);
    if (status != STATUS_SUCCESS) {
        return Twapi_AppendSystemError(interp,
                                       TwapiNTSTATUSToError(status));
    }

    resultObj = Tcl_NewListObj(0, NULL);
    for (i = 0; i < count; ++i) {
        Tcl_ListObjAppendElement(interp, resultObj,
                                 ObjFromLUID(&luids[i]));
    }

    Tcl_SetObjResult(interp, resultObj);

    LsaFreeReturnBuffer(luids);

    return TCL_OK;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int Twapi_LsaGetLogonSessionData(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    NTSTATUS status;
    PSECURITY_LOGON_SESSION_DATA sessionP;
    Tcl_Obj *resultObj;
    LUID    luid;

    if (objc != 1 ||
        ObjToLUID(interp, objv[0], &luid) != TCL_OK)
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);

    status = LsaGetLogonSessionData(&luid, &sessionP);
    if (status != STATUS_SUCCESS) {
        return Twapi_AppendSystemError(interp,
                                       TwapiNTSTATUSToError(status));
    }

    if (sessionP == NULL) {
        // MSDN says this may happen for the LocalSystem session. Although
        // experimentation, at least on XP SP2 indicates otherwise.
        /* Caller has to figure out that no data was available. We  don't
           want to return an error since it is a valid logon session */
        return TCL_OK;
    }

    resultObj = Tcl_NewListObj(0, NULL);

    Twapi_APPEND_LUID_FIELD_TO_LIST(interp, resultObj, sessionP, LogonId);
    Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp, resultObj, sessionP, UserName);
    Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp, resultObj, sessionP, LogonDomain);
    Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp, resultObj, sessionP, AuthenticationPackage);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, sessionP, LogonType);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, resultObj, sessionP, Session);
    Twapi_APPEND_PSID_FIELD_TO_LIST(interp, resultObj, sessionP, Sid);
    Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp, resultObj, sessionP, LogonTime);
    /* Some fields are not present on Windows 2000 */
    if (sessionP->Size > offsetof(struct _SECURITY_LOGON_SESSION_DATA, LogonServer)) {
        Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp, resultObj, sessionP, LogonServer);
    }
    if (sessionP->Size > offsetof(struct _SECURITY_LOGON_SESSION_DATA, DnsDomainName)) {
        Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp, resultObj, sessionP, DnsDomainName);
    }
    if (sessionP->Size > offsetof(struct _SECURITY_LOGON_SESSION_DATA, Upn)) {
        Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp, resultObj, sessionP, Upn);
    }

    Tcl_SetObjResult(interp, resultObj);

    LsaFreeReturnBuffer(sessionP);
    return TCL_OK;
}
#endif // TWAPI_LEAN


#if 0
/* Not explicitly visible in Win2K and XP SP0 so we use CheckTokenMembership
   directly for which this is a wrapper */
BOOL IsUserAnAdmin();
#endif


static int Twapi_SecurityCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int func;
    LPWSTR s, s2, s3;
    DWORD dw, dw2, dw3;
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h, h2;
    SECURITY_DESCRIPTOR *secdP;
    LUID luid;
    ACL *daclP, *saclP;
    PSID osidP, gsidP;
    union {
        COORD coord;
        WCHAR buf[MAX_PATH+1];
        TWAPI_TOKEN_MANDATORY_POLICY ttmp;
        TWAPI_TOKEN_MANDATORY_LABEL ttml;
        TOKEN_PRIMARY_GROUP tpg;
        TOKEN_OWNER towner;
        LSA_UNICODE_STRING lsa_ustr;
        TOKEN_PRIVILEGES *tokprivsP;
    } u;
    ULONG  lsa_count;
    LSA_UNICODE_STRING *lsa_strings;
    TwapiResult result;


    /* These will be freed at end of routine if not NULL! */
    daclP = saclP = NULL;
    osidP = gsidP = NULL;

    /* secdP is allocated multiple ways so no common free */
   secdP = NULL;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 100) {
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 1:
            return Twapi_LsaEnumerateLogonSessions(interp);
        case 2:
            return Twapi_InitializeSecurityDescriptor(interp);
        }
    } else if (func < 200) {
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 101:
            if (ObjToPSECURITY_DESCRIPTOR(interp, objv[2], &secdP) != TCL_OK)
                return TCL_ERROR;
            // Note secdP may be NULL
            result.type = TRT_BOOL;
            result.value.bval = secdP ? IsValidSecurityDescriptor(secdP) : 0;
            if (secdP)
                TwapiFreeSECURITY_DESCRIPTOR(secdP);
            break;
        case 102:
            if (ObjToPACL(interp, objv[2], &daclP) != TCL_OK)
                return TCL_ERROR;
            // Note aclP may me NULL even on TCL_OK
            result.type = TRT_BOOL;
            result.value.bval = daclP ? IsValidAcl(daclP) : 0;
            break;
        case 103:
            CHECK_INTEGER_OBJ(interp, dw, objv[2]);
            result.value.ival = ImpersonateSelf(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 104:
            if (ObjToHANDLE(interp, objv[2], &h) != TCL_OK)
                return TCL_ERROR;
            result.value.ival = ImpersonateLoggedOnUser(h);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        }
    } else if (func < 500) {
        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        s = Tcl_GetUnicode(objv[2]);
        s2 = Tcl_GetUnicode(objv[3]);
        switch (func) {
        case 401:
            result.value.unicode.len = ARRAYSIZE(u.buf);
            if (LookupPrivilegeDisplayNameW(s,s2,u.buf,&result.value.unicode.len,&dw)) {
                result.value.unicode.str = u.buf;
                result.type = TRT_UNICODE;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 402:
            if (LookupPrivilegeValueW(s,s2,&result.value.luid))
                result.type = TRT_LUID;
            else
                result.type = TRT_GETLASTERROR;
            break;
        }
    } else if (func < 1000) {
        /* Args - string, dw, optional dw2 */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETWSTR(s), GETINT(dw), ARGUSEDEFAULT,
                         GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 501:
            if (ConvertStringSecurityDescriptorToSecurityDescriptorW(
                    s, dw, &secdP, NULL)) {
                result.value.obj = ObjFromSECURITY_DESCRIPTOR(interp, secdP);
                if (secdP)
                    LocalFree(secdP);
                if (result.value.obj)
                    result.type = TRT_OBJ;
                else
                    return TCL_ERROR;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 502:
            return Twapi_GetNamedSecurityInfo(interp, s, dw, dw2);
        }
    } else if (func < 2000) {
        /* Args - handle, int, optional int */
        if (TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETINT(dw), ARGUSEDEFAULT,
                         GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 1003:
            return Twapi_GetSecurityInfo(interp, h, dw, dw2);
        case 1004:
            result.type = OpenThreadToken(h, dw, dw2, &result.value.hval) ?
                TRT_HANDLE : TRT_GETLASTERROR;
            break;
        case 1005:
            result.type = OpenProcessToken(h, dw, &result.value.hval) ?
                TRT_HANDLE : TRT_GETLASTERROR;
            break;
        case 1006:
            return Twapi_GetTokenInformation(interp, h, dw);

        case 1007:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetTokenInformation(h,
                                                    TwapiTokenVirtualizationEnabled,
                                                    &dw, sizeof(dw));
            break;
        case 1008:
            u.ttmp.Policy = dw;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetTokenInformation(h,
                                                    TwapiTokenMandatoryPolicy,
                                                    &u.ttmp, sizeof(u.ttmp));
            break;
        }
    } else if (func < 3000) {
        if (objc != 4)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToHANDLE(interp, objv[2], &h) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 3002: // SetThreadToken
            if (ObjToHANDLE(interp, objv[3], &h2) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetThreadToken(h, h2);
            break;
        case 3003:
            ObjToLSA_UNICODE_STRING(objv[3], &u.lsa_ustr);
            return Twapi_LsaEnumerateAccountsWithUserRight(interp, h,
                                                           &u.lsa_ustr);
        case 3004:
            if (ObjToSID_AND_ATTRIBUTES(interp, objv[3], &u.ttml.Label) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetTokenInformation(h,
                                                    TwapiTokenIntegrityLevel,
                                                    &u.ttml, sizeof(u.ttml));
            if (u.ttml.Label.Sid)
                TwapiFree(u.ttml.Label.Sid);
            break;
        }
    } else if (func < 5000) {
        /* Note: pointers may have to be freed even when arg parsing
         * fails so always fall through to end of function even on errors!
         */
        result.type = TRT_TCL_RESULT;
        result.value.ival =
            TwapiGetArgs(interp, objc-2, objv+2,
                         GETHANDLE(h), GETVAR(osidP, ObjToPSID),
                         ARGEND);
        if (result.value.ival == TCL_OK) {
            switch (func) {
            case 4001:
                u.tpg.PrimaryGroup = osidP;
                result.type = TRT_EXCEPTION_ON_FALSE;
                result.value.ival = SetTokenInformation(h, TokenPrimaryGroup, &u.tpg, sizeof(u.tpg));
                break;
            case 4002:
                result.type = CheckTokenMembership(h, osidP, &result.value.bval)
                    ? TRT_BOOL : TRT_GETLASTERROR;
                break;
            case 4003:
                u.towner.Owner = osidP;
                result.type = TRT_EXCEPTION_ON_FALSE;
                result.value.ival = SetTokenInformation(h, TokenOwner,
                                                        &u.towner,
                                                        sizeof(u.towner));
                break;
            case 4004:
                result.type = TRT_TCL_RESULT;
                result.value.ival = Twapi_LsaEnumerateAccountRights(interp,
                                                                    h, osidP);
                break;
            }
        }
    } else {
        /* Arbitrary args */
        switch (func) {
        case 10014:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETNULLIFEMPTY(s2), GETWSTR(s3),
                             GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (LogonUserW(s, s2, s3, dw,dw2, &result.value.hval))
                result.type = TRT_HANDLE;
            else
                result.type = TRT_GETLASTERROR;
            break;
            
        case 10015:
            result.value.ival =
                TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETVAR(osidP, ObjToPSID),
                             ARGSKIP, ARGEND);
            result.type = TRT_TCL_RESULT;
            if (result.value.ival != TCL_OK)
                break;
            result.value.ival = ObjToLSASTRINGARRAY(interp, objv[4],
                                                    &lsa_strings, &lsa_count);
            if (result.value.ival != TCL_OK)
                break;

            result.type = TRT_NTSTATUS;
            result.value.ival = LsaAddAccountRights(h, osidP,
                                                    lsa_strings, lsa_count);
            TwapiFree(lsa_strings);
            break;

        case 10016:
            result.value.ival =
                TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETVAR(osidP, ObjToPSID),
                             GETINT(dw), ARGSKIP, ARGEND);
            result.type = TRT_TCL_RESULT;
            if (result.value.ival != TCL_OK)
                break;
            result.value.ival = ObjToLSASTRINGARRAY(interp, objv[5],
                                                    &lsa_strings, &lsa_count);
            if (result.value.ival != TCL_OK)
                break;

            result.type = TRT_NTSTATUS;
            result.value.ival = LsaRemoveAccountRights(
                h, osidP, (BOOLEAN) (dw ? 1 : 0), lsa_strings, lsa_count);
            TwapiFree(lsa_strings);
            break;

        case 10017:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETVAR(secdP, ObjToPSECURITY_DESCRIPTOR),
                             GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK) {
                if (secdP)
                    TwapiFreeSECURITY_DESCRIPTOR(secdP);
                return TCL_ERROR;
            }    
            if (ConvertSecurityDescriptorToStringSecurityDescriptorW(
                    secdP, dw, dw2, &s, &dw3)) {
                /* Cannot use TRT_UNICODE since buffer has to be freed */
                result.type = TRT_OBJ;
                /* Do not use dw3 as length because it seems to be size
                   of buffer, not string length as it includes padded nulls */
                result.value.obj = ObjFromUnicode(s);
                LocalFree(s);
            } else
                result.type = TRT_GETLASTERROR;
            if (secdP)
                TwapiFreeSECURITY_DESCRIPTOR(secdP);
            break;
        case 10018: // UNUSED
            break;
        case 10019:
            return Twapi_LsaGetLogonSessionData(interp, objc-2, objv+2);
        case 10020: // SetNamedSecurityInfo
            /* Note even in case of errors, sids and acls might have been alloced */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETINT(dw), GETINT(dw2),
                             GETVAR(osidP, ObjToPSID),
                             GETVAR(gsidP, ObjToPSID),
                             GETVAR(daclP, ObjToPACL),
                             GETVAR(saclP, ObjToPACL),
                             ARGEND) == TCL_OK) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = SetNamedSecurityInfoW(
                    s, dw, dw2, osidP, gsidP, daclP, saclP);
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            break;
        case 10021: // LookupPrivilegeName
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETWSTR(s), GETVAR(luid, ObjToLUID),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            result.value.unicode.len = sizeof(u.buf)/sizeof(u.buf[0]);
            if (LookupPrivilegeNameW(s, &luid,
                                     u.buf, &result.value.unicode.len)) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
                result.value.unicode.len = -1;
            } else
                result.type = TRT_GETLASTERROR;
            break;

        case 10022:
            /* Init to NULL as they may be partially init'ed on error
               and have to be freed */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw), GETINT(dw2),
                             GETVAR(osidP, ObjToPSID),
                             GETVAR(gsidP, ObjToPSID),
                             GETVAR(daclP, ObjToPACL),
                             GETVAR(saclP, ObjToPACL),
                             ARGEND) == TCL_OK) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = SetSecurityInfo(
                    h, dw, dw2, osidP, gsidP, daclP, saclP);
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            break;

        case 10023: // DuplicateTokenEx
            secattrP = NULL;        /* Even on error, it might be filled */
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw),
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETINT(dw2), GETINT(dw3),
                             ARGEND) == TCL_OK) {
                if (DuplicateTokenEx(h, dw, secattrP, dw2, dw3, &result.value.hval))
                    result.type = TRT_HANDLE;
                else
                    result.type = TRT_GETLASTERROR;
            } else {
                result.type = TRT_TCL_RESULT;
                result.value.ival = TCL_ERROR;
            }
            TwapiFreeSECURITY_ATTRIBUTES(secattrP); // Even in case of error or NULL
            break;

        case 10024: // AdjustTokenPrivileges
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETBOOL(dw),
                             GETVAR(u.tokprivsP, ObjToPTOKEN_PRIVILEGES),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_AdjustTokenPrivileges(
                ticP, h, dw, u.tokprivsP);
            TwapiFreeTOKEN_PRIVILEGES(u.tokprivsP);
            break;
        case 10025: // PrivilegeCheck
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h),
                             GETVAR(u.tokprivsP, ObjToPTOKEN_PRIVILEGES),
                             GETBOOL(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (Twapi_PrivilegeCheck(h, u.tokprivsP, dw, &result.value.ival))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            TwapiFreeTOKEN_PRIVILEGES(u.tokprivsP);
            break;
        }
    }

    if (daclP) TwapiFree(daclP);
    if (saclP) TwapiFree(saclP);
    if (osidP) TwapiFree(osidP);
    if (gsidP) TwapiFree(gsidP);

    return TwapiSetResult(interp, &result);
}


static int Twapi_SecurityInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::SecCall", Twapi_SecurityCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::SecCall", # code_); \
    } while (0);

    CALL_(LsaEnumerateLogonSessions, 1);
    CALL_(Twapi_InitializeSecurityDescriptor, 2);
    CALL_(IsValidSecurityDescriptor, 101);
    CALL_(IsValidAcl, 102);
    CALL_(ImpersonateSelf, 103);
    CALL_(ImpersonateLoggedOnUser, 104);
    CALL_(LookupPrivilegeDisplayName, 401);
    CALL_(LookupPrivilegeValue, 402);
    CALL_(ConvertStringSecurityDescriptorToSecurityDescriptor, 501);
    CALL_(GetNamedSecurityInfo, 502);
    CALL_(GetSecurityInfo, 1003);
    CALL_(OpenThreadToken, 1004);
    CALL_(OpenProcessToken, 1005);
    CALL_(GetTokenInformation, 1006);
    CALL_(Twapi_SetTokenVirtualizationEnabled, 1007);
    CALL_(Twapi_SetTokenMandatoryPolicy, 1008);
    CALL_(SetThreadToken, 3002);
    CALL_(Twapi_LsaEnumerateAccountsWithUserRight, 3003);
    CALL_(Twapi_SetTokenIntegrityLevel, 3004);
    CALL_(Twapi_SetTokenPrimaryGroup, 4001);
    CALL_(CheckTokenMembership, 4002);
    CALL_(Twapi_SetTokenOwner, 4003);
    CALL_(Twapi_LsaEnumerateAccountRights, 4004);
    CALL_(LogonUser, 10014);
    CALL_(LsaAddAccountRights, 10015);
    CALL_(LsaRemoveAccountRights, 10016);
    CALL_(ConvertSecurityDescriptorToStringSecurityDescriptor, 10017);
    CALL_(LsaGetLogonSessionData, 10019);
    CALL_(SetNamedSecurityInfo, 10020);
    CALL_(LookupPrivilegeName, 10021);
    CALL_(SetSecurityInfo, 10022);
    CALL_(DuplicateTokenEx, 10023);
    CALL_(Twapi_AdjustTokenPrivileges, 10024);
    CALL_(Twapi_PrivilegeCheck, 10025);

#undef CALL_

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
int Twapi_security_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return Twapi_ModuleInit(interp, WLITERAL(MODULENAME), MODULE_HANDLE,
                            Twapi_SecurityInitCalls, NULL) ? TCL_OK : TCL_ERROR;
}

