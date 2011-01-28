/*
 * Copyright (c) 2003-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to Windows API related to security and access control functions */

#include "twapi.h"

/*
TBD - get_wellknown_sid - returns well known sid's such as administrators group
TBD - the SID data type handling needs to be cleaned up
 */

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


/* Returned memory in *arrayPP has to be freed by caller */
int ObjToLSASTRINGARRAY(Tcl_Interp *interp, Tcl_Obj *obj, LSA_UNICODE_STRING **arrayP, ULONG *countP)
{
    Tcl_Obj **listobjv;
    int       i, nitems, sz;
    LSA_UNICODE_STRING *ustrP;
    WCHAR    *dstP;

    if (Tcl_ListObjGetElements(interp, obj, &nitems, &listobjv) == TCL_ERROR) {
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
        srcP = Tcl_GetUnicodeFromObj(listobjv[i], &slen);
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
static PSID TwapiGetSidFromStringRep(char *strP)
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

    s = Tcl_GetStringFromObj(obj, &len);
    if (len == 0) {
        *sidPP = NULL;
        return TCL_OK;
    }

    *sidPP = TwapiGetSidFromStringRep(Tcl_GetString(obj));
    if (*sidPP)
        return TCL_OK;

    winerror = GetLastError();

    /* Not a string rep. See if it is a binary of the right size */
    sidP = (SID *) Tcl_GetByteArrayFromObj(obj, &len);
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


/* Like ObjFromSID but returns empty object on error */
Tcl_Obj *ObjFromSIDNoFail(SID *sidP)
{
    Tcl_Obj *objP;
    return (ObjFromSID(NULL, sidP, &objP) == TCL_OK ? objP : Tcl_NewStringObj("", 0));
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
            Tcl_SetResult(interp, "NULL ACE pointer", TCL_STATIC);
        return NULL;
    }

    resultObj = Tcl_NewListObj(0, NULL);
    if (resultObj == NULL)
        goto allocation_error_return;

    /* ACE type */
    Tcl_ListObjAppendElement(interp, resultObj,
                             Tcl_NewIntObj(acehdrP->AceType));

    /* ACE flags */
    Tcl_ListObjAppendElement(interp, resultObj,
                             Tcl_NewIntObj(acehdrP->AceFlags));

    /* Now for type specific fields */
    switch (acehdrP->AceType) {
    case ACCESS_ALLOWED_ACE_TYPE:
    case ACCESS_DENIED_ACE_TYPE:
    case SYSTEM_AUDIT_ACE_TYPE:
    case SYSTEM_MANDATORY_LABEL_ACE_TYPE:
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(((ACCESS_ALLOWED_ACE *)aceP)->Mask));

        /* and the SID */
        obj = NULL;                /* In case of errors */
        if (ObjFromSID(interp,
                         (SID *)&((ACCESS_ALLOWED_ACE *)aceP)->SidStart,
                         &obj)
            != TCL_OK) {
            goto error_return;
        }
        Tcl_ListObjAppendElement(interp, resultObj, obj);
        break;

    case ACCESS_ALLOWED_OBJECT_ACE_TYPE:
    case ACCESS_DENIED_OBJECT_ACE_TYPE:
    case SYSTEM_AUDIT_OBJECT_ACE_TYPE:
        objectAceP = (ACCESS_ALLOWED_OBJECT_ACE *)aceP;
        Tcl_ListObjAppendElement(interp, resultObj,
                                 Tcl_NewIntObj(objectAceP->Mask));
        if (objectAceP->Flags & ACE_OBJECT_TYPE_PRESENT) {
            Tcl_ListObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->ObjectType));
            if (objectAceP->Flags & ACE_INHERITED_OBJECT_TYPE_PRESENT) {
                Tcl_ListObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->InheritedObjectType));
                sidP = (SID *) &objectAceP->SidStart;
            } else {
                Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewObj());
                sidP = (SID *) &objectAceP->InheritedObjectType;
            }
        } else if (objectAceP->Flags & ACE_INHERITED_OBJECT_TYPE_PRESENT) {
            Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewObj());
            Tcl_ListObjAppendElement(interp, resultObj, ObjFromGUID(&objectAceP->ObjectType));
            sidP = (SID *) &objectAceP->InheritedObjectType;
        } else {
            Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewObj());
            Tcl_ListObjAppendElement(interp, resultObj, Tcl_NewObj());
            sidP = (SID *) &objectAceP->ObjectType;
        }
        obj = NULL;                /* In case of errors */
        if (ObjFromSID(interp, sidP, &obj) != TCL_OK)
            goto error_return;
        Tcl_ListObjAppendElement(interp, resultObj, obj);
        
        break;

    default:
        /*
         * Return a binary rep of the whole dang thing.
         * There are no pointers in there, just values so this
         * should work, I think :)
         */
        obj = Tcl_NewByteArrayObj((unsigned char *) aceP, acehdrP->AceSize);
        if (obj == NULL)
            goto allocation_error_return;

        if (Tcl_ListObjAppendElement(interp, resultObj, obj) != TCL_OK)
            goto error_return;

        break;
    }


    return resultObj;


 allocation_error_return:
    if (interp) {
        Tcl_SetResult(interp, "Could not allocate Tcl object", TCL_STATIC);
    }

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

    if (Tcl_ListObjGetElements(interp, aceobj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc < 2)
        goto format_error;

    if ((Tcl_GetIntFromObj(interp, objv[0], &acetype) != TCL_OK) ||
        (Tcl_GetIntFromObj(interp, objv[1], &aceflags) != TCL_OK)) {
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
        if (Tcl_GetIntFromObj(interp, objv[2], &aceP->Mask) != TCL_OK)
            goto format_error;

        sidP = TwapiGetSidFromStringRep(Tcl_GetString(objv[3]));
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
        bytes = Tcl_GetByteArrayFromObj(objv[2], &bytecount);
        acesz += bytecount;
        aceP = (ACCESS_ALLOWED_ACE *) TwapiAlloc(acesz);
        CopyMemory(aceP, bytes, bytecount);
        break;
    }

    *acePP = aceP;
    return TCL_OK;

 format_error:
    if (interp)
        Tcl_SetResult(interp, "Invalid ACE format.", TCL_STATIC);
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
    Tcl_Obj                 *resultObj;

    if (aclP == NULL) {
        return Tcl_NewStringObj("null", -1);
    }

    if ((GetAclInformation(aclP, &acl_rev, sizeof(acl_rev),
                           AclRevisionInformation) == 0) ||
        GetAclInformation(aclP, &acl_szinfo, sizeof(acl_szinfo),
                          AclSizeInformation) == 0) {
        TwapiReturnSystemError(interp);
        return NULL;
    }

    objv[0] = Tcl_NewIntObj(acl_rev.AclRevision);
    objv[1] = Tcl_NewListObj(0, NULL);
    if (objv[0] == NULL || objv[1] == NULL) {
        goto allocation_error_return;
    }

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
        if (Tcl_ListObjAppendElement(interp, objv[1], ace_obj) != TCL_OK) {
            goto error_return;
        }
    }


    resultObj = Tcl_NewListObj(2, objv);
    if (resultObj == NULL)
        goto allocation_error_return;

    return resultObj;

 allocation_error_return:
    if (interp) {
        Tcl_SetResult(interp, "Could not allocate Tcl object", TCL_STATIC);
    }

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
    if (!lstrcmpA("null", Tcl_GetString(aclObj)))
        return TCL_OK;

    if (Tcl_ListObjGetElements(interp, aclObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc != 2) {
        if (interp)
            Tcl_SetResult(interp,
                          "Invalid ACL format. Should be 'null' or have exactly two elements",
                          TCL_STATIC);
        return TCL_ERROR;
    }

    /*
     * First figure out how much space we need to allocate. For this, we
     * first need to figure out space for the ACE's
     */
#if 0
    objv[0] is the ACL rev. We always recalculate it, ignore value passed in.
    if (Tcl_GetIntFromObj(interp, objv[0], &aclrev) != TCL_OK)
        goto error_return;
#endif
    if (Tcl_ListObjGetElements(interp, objv[1], &aceobjc, &aceobjv) != TCL_OK)
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
            Tcl_SetResult(interp,
                          "Internal error constructing ACL",
                          TCL_STATIC);
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
        return Tcl_NewListObj(0, NULL);
    }

    if (! GetSecurityDescriptorControl(secdP, &secd_control, &rev))
        goto system_error;

    if (rev != SECURITY_DESCRIPTOR_REVISION) {
        /* Dunno how to handle this */
        if (interp)
            Tcl_SetResult(interp, "Unsupported SECURITY_DESCRIPTOR version", TCL_STATIC);
        goto error_return;
    }

    /* Control bits */
    objv[0] = Tcl_NewIntObj(secd_control);

    /* Owner SID */
    if (! GetSecurityDescriptorOwner(secdP, &sidP, &defaulted))
        goto system_error;
    if (sidP == NULL)
        objv[1] = Tcl_NewStringObj("", -1);
    else {
        if (ObjFromSID(interp, sidP, &objv[1]) != TCL_OK)
            goto error_return;
    }

    /* Group SID */
    if (! GetSecurityDescriptorGroup(secdP, &sidP, &defaulted))
        goto system_error;
    if (sidP == NULL)
        objv[2] = Tcl_NewStringObj("", -1);
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
    return Tcl_NewListObj(5, objv);

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

    if (Tcl_ListObjGetElements(interp, secdObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc == 0)
        return TCL_OK;          /* NULL security descriptor */

    if (objc != 5) {
        if (interp)
            Tcl_SetResult(interp,
                          "Invalid SECURITY_DESCRIPTOR format. Should have 0 or five elements",
                          TCL_STATIC);
        return TCL_ERROR;
    }


    *secdPP = TwapiAlloc (sizeof(SECURITY_DESCRIPTOR));
    if (! InitializeSecurityDescriptor(*secdPP, SECURITY_DESCRIPTOR_REVISION))
        goto system_error;

    /*
     * Set control field
     */
    if (Tcl_GetIntFromObj(interp, objv[0], &temp) != TCL_OK)
        goto error_return;
    secd_control = (SECURITY_DESCRIPTOR_CONTROL) temp;
    if (secd_control != temp) {
        /* Truncation error */
        if (interp)
            Tcl_SetResult(interp, "Invalid control flags for SECURITY_DESCRIPTOR", TCL_STATIC);
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
    s = Tcl_GetStringFromObj(objv[1], &slen);
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
    s = Tcl_GetStringFromObj(objv[2], &slen);
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

    if (Tcl_ListObjGetElements(interp, secattrObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc == 0)
        return TCL_OK;          /* NULL security attributes */

    if (objc != 2) {
        if (interp)
            Tcl_SetResult(interp,
                          "Invalid SECURITY_ATTRIBUTES format. Should have 0 or 2 elements",
                          TCL_STATIC);
        return TCL_ERROR;
    }


    *secattrPP = TwapiAlloc (sizeof(**secattrPP));
    (*secattrPP)->nLength = sizeof(**secattrPP);

    if (Tcl_GetIntFromObj(interp, objv[1], &inherit) == TCL_ERROR)
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


int Twapi_LookupAccountSid (
    Tcl_Interp *interp,
    LPCWSTR     lpSystemName,
    PSID        sidP
)
{
    WCHAR       *domainP;
    DWORD        domain_buf_size;
    SID_NAME_USE account_type;
    DWORD        error;
    int          result;
    Tcl_Obj     *objs[3];
    LPWSTR       nameP;
    DWORD        name_buf_size;
    int          i;

    for (i=0; i < (sizeof(objs)/sizeof(objs[0])); ++i)
        objs[i] = NULL;

    result = TCL_ERROR;

    domainP         = NULL;
    domain_buf_size = 0;
    nameP           = NULL;
    name_buf_size   = 0;
    error           = 0;
    if (LookupAccountSidW(lpSystemName, sidP, NULL, &name_buf_size,
                          NULL, &domain_buf_size, &account_type) == 0) {
        error = GetLastError();
    }

    if (error && (error != ERROR_INSUFFICIENT_BUFFER)) {
        Tcl_SetResult(interp, "Error looking up account SID: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, error);
        goto done;
    }

    /* Allocate required space */
    domainP = TwapiAlloc(domain_buf_size * sizeof(*domainP));
    nameP = TwapiAlloc(name_buf_size * sizeof(*nameP));

    if (LookupAccountSidW(lpSystemName, sidP, nameP, &name_buf_size,
                          domainP, &domain_buf_size, &account_type) == 0) {
        Tcl_SetResult(interp, "Error looking up account SID: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, GetLastError());
        goto done;
    }

    /*
     * Got everything we need, now format it
     * {NAME DOMAIN ACCOUNT}
     */
    objs[0] = Tcl_NewUnicodeObj(nameP, -1);   /* Will exit on alloc fail */
    objs[1] = Tcl_NewUnicodeObj(domainP, -1); /* Will exit on alloc fail */
    objs[2] = Tcl_NewIntObj(account_type);
    Tcl_SetObjResult(interp, Tcl_NewListObj(3, objs));
    result = TCL_OK;

 done:
    if (domainP)
        TwapiFree(domainP);
    if (nameP)
        TwapiFree(nameP);

    return result;
}


int Twapi_LookupAccountName (
    Tcl_Interp *interp,
    LPCWSTR     lpSystemName,
    LPCWSTR     lpAccountName
)
{
    PSID         sidP;
    DWORD        sid_buf_size;
    WCHAR       *domainP;
    DWORD        domain_buf_size;
    SID_NAME_USE account_type;
    DWORD        error;
    int          result;
    Tcl_Obj     *objs[3];
    int          i;

    /*
     * Special case check for empty string - else LookupAccountName
     * returns the same error as for insufficient buffer .
     */
    if (*lpAccountName == 0) {
        return Twapi_GenerateWin32Error(interp, ERROR_INVALID_PARAMETER, "Empty string passed for account name.");
    }

    for (i=0; i < (sizeof(objs)/sizeof(objs[0])); ++i)
        objs[i] = NULL;
    result = TCL_ERROR;


    domain_buf_size = 0;
    sid_buf_size    = 0;
    error           = 0;
    if (LookupAccountNameW(lpSystemName, lpAccountName, NULL, &sid_buf_size,
                          NULL, &domain_buf_size, &account_type) == 0) {
        error = GetLastError();
    }

    if (error && (error != ERROR_INSUFFICIENT_BUFFER)) {
        Tcl_SetResult(interp, "Error looking up account name: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, error);
        return TCL_ERROR;
    }

    /* Allocate required space */
    domainP = TwapiAlloc(domain_buf_size * sizeof(*domainP));
    sidP = TwapiAlloc(sid_buf_size);

    if (LookupAccountNameW(lpSystemName, lpAccountName, sidP, &sid_buf_size,
                          domainP, &domain_buf_size, &account_type) == 0) {
        Tcl_SetResult(interp, "Error looking up account name: ", TCL_STATIC);
        Twapi_AppendSystemError(interp, GetLastError());
        goto done;
    }

    /*
     * There is a bug in LookupAccountName (see KB 185246) where
     * if the user name happens to be the machine name, the returned SID
     * is for the machine, not the user. As suggested there, we look
     * for this case by checking the account type returned and if we have hit
     * this case, recurse using a user name of "\\domain\\username"
     */
    if (account_type == SidTypeDomain) {
        /* Redo the operation */
        WCHAR *new_accountP;
        size_t len = 0;
        size_t sysnamelen, accnamelen;
        TWAPI_ASSERT(lpSystemName);
        TWAPI_ASSERT(lpAccountName);
        sysnamelen = lstrlenW(lpSystemName);
        accnamelen = lstrlenW(lpAccountName);
        len = sysnamelen + 1 + accnamelen + 1;
        new_accountP = TwapiAlloc(len * sizeof(*new_accountP));
        CopyMemory(new_accountP, lpSystemName, sizeof(*new_accountP)*sysnamelen);
        new_accountP[sysnamelen] = L'\\';
        CopyMemory(new_accountP+sysnamelen+1, lpAccountName, sizeof(*new_accountP)*accnamelen);
        new_accountP[sysnamelen+1+accnamelen] = 0;

        /* Recurse */
        result = Twapi_LookupAccountName(interp, lpSystemName, new_accountP);
        TwapiFree(new_accountP);
        goto done;
    }


    /*
     * Got everything we need, now format it
     * {SID DOMAIN ACCOUNT}
     */
    result = ObjFromSID(interp, sidP, &objs[0]);
    if (result != TCL_OK)
        goto done;
    objs[1] = Tcl_NewUnicodeObj(domainP, -1); /* Will exit on alloc fail */
    objs[2] = Tcl_NewIntObj(account_type);
    Tcl_SetObjResult(interp, Tcl_NewListObj(3, objs));
    result = TCL_OK;

 done:
    if (domainP)
        TwapiFree(domainP);
    if (sidP)
        TwapiFree(sidP);

    return result;
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
    DWORD    i,sz;
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
        resultObj = Tcl_NewListObj(0, NULL);
        for (i = 0; i < ((TOKEN_GROUPS *)infoP)->GroupCount; ++i) {
            obj = ObjFromSID_AND_ATTRIBUTES(
                interp,
                &((TOKEN_GROUPS *)infoP)->Groups[i]
                );
            if (obj)
                Tcl_ListObjAppendElement(interp, resultObj, obj);
            else
                break;
        }
        if (i == ((TOKEN_GROUPS *)infoP)->GroupCount)
            result = TCL_OK;
        break;

    case TokenPrivileges:
        resultObj = ObjFromTOKEN_PRIVILEGES(interp,
                                                (TOKEN_PRIVILEGES *)infoP);
        if (resultObj)
            result = TCL_OK;

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
    case TokenGroupsAndPrivileges:
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
        MemLifoPopFrame(&ticP->memlifo);
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

Tcl_Obj *ObjFromCONNECTION_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    CONNECTION_INFO_1 *ciP;
    int         objc;
    Tcl_Obj    *objv[14];

    ciP = (CONNECTION_INFO_1 *) infoP; /* May actually be CONNECTION_INFO_0 */
    /* We build from back of array since we would like basic elements
       at front of Tcl list we build */
    objc = sizeof(objv)/sizeof(objv[0]);
    switch (level) {
    case 1:
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_type);
        objv[--objc] = STRING_LITERAL_OBJ("type");
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_num_opens);
        objv[--objc] = STRING_LITERAL_OBJ("num_opens");
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_time);
        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_num_users);
        objv[--objc] = STRING_LITERAL_OBJ("num_users");
        objv[--objc] = Tcl_NewUnicodeObj(ciP->coni1_username, -1);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = Tcl_NewUnicodeObj(ciP->coni1_netname, -1);
        objv[--objc] = STRING_LITERAL_OBJ("netname");
        /* FALLTHRU */
    case 0:
        objv[--objc] = Tcl_NewLongObj(ciP->coni1_id);
        objv[--objc] = STRING_LITERAL_OBJ("id");
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
}
        
/*
 * Convert USE_INFO_* to list
 */
Tcl_Obj *ObjFromUSE_INFO(
    Tcl_Interp *interp,
    LPBYTE infoP,
    DWORD level
    )
{
    USE_INFO_2 *uiP;
    int         objc;
    Tcl_Obj    *objv[18];

    uiP = (USE_INFO_2 *) infoP; /* May acutally be any USE_INFO_* */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build (most accessed elements in front */
    objc = sizeof(objv)/sizeof(objv[0]);
#define ADD_LPWSTR_(fld) do {                                      \
        objv[--objc] = Tcl_NewUnicodeObj(uiP->ui2_ ## fld, -1); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[--objc] = Tcl_NewLongObj(uiP->ui2_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)

    switch (level) {
    case 2:
        ADD_LPWSTR_(username);
        ADD_LPWSTR_(domainname);
        // FALLTHRU
    case 1:
        ADD_LPWSTR_(password);
        ADD_DWORD_(status);
        ADD_DWORD_(asg_type);
        ADD_DWORD_(refcount);
        ADD_DWORD_(usecount);
        //FALLTHRU
    case 0:
        ADD_LPWSTR_(local);
        ADD_LPWSTR_(remote);
        break;

    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);

}

Tcl_Obj *ObjFromSHARE_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    SHARE_INFO_502 *siP;
    int         objc;
    Tcl_Obj    *objv[18];

    siP = (SHARE_INFO_502 *) infoP; /* May acutally be any SHARE_INFO_* */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build (most accessed elements in front */
    objc = sizeof(objv)/sizeof(objv[0]);
#define ADD_LPWSTR_(fld) do {                                          \
        objv[--objc] = Tcl_NewUnicodeObj(siP->shi502_ ## fld, -1); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[--objc] = Tcl_NewLongObj(siP->shi502_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)

    switch (level) {
    case 502:
        objv[--objc] = ObjFromSECURITY_DESCRIPTOR(interp,
                                                      siP->shi502_security_descriptor);
        objv[--objc] = STRING_LITERAL_OBJ("security_descriptor");
        // FALLTHRU
    case 2:
        ADD_LPWSTR_(passwd);
        ADD_LPWSTR_(path);
        ADD_DWORD_(current_uses);
        ADD_DWORD_(max_uses);
        ADD_DWORD_(permissions);
        // FALLTHRU
    case 1:
        ADD_LPWSTR_(remark);
        ADD_DWORD_(type);
        // FALLTHRU
    case 0:
        ADD_LPWSTR_(netname);
        break;

    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
}


Tcl_Obj *ObjFromFILE_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    int         objc;
    Tcl_Obj    *objv[10];
    FILE_INFO_3 *fiP;

    fiP = (FILE_INFO_3 *) infoP; /* May acutally be FILE_INFO_* */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build */
    objc = sizeof(objv)/sizeof(objv[0]);
    switch (level) {
    case 3:
        objv[--objc] = Tcl_NewLongObj(fiP->fi3_permissions);
        objv[--objc] = STRING_LITERAL_OBJ("permissions");
        objv[--objc] = Tcl_NewLongObj(fiP->fi3_num_locks);
        objv[--objc] = STRING_LITERAL_OBJ("num_locks");
        objv[--objc] = Tcl_NewUnicodeObj(fiP->fi3_username, -1);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = Tcl_NewUnicodeObj(fiP->fi3_pathname, -1);
        objv[--objc] = STRING_LITERAL_OBJ("pathname");
        /* FALLTHRU */
    case 2:
        objv[--objc] = Tcl_NewLongObj(fiP->fi3_id);
        objv[--objc] = STRING_LITERAL_OBJ("id");
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
}

Tcl_Obj *ObjFromSESSION_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       level
    ) 
{
    int         objc;
    Tcl_Obj    *objv[16];

    /* Note in the code below, the structures can be superimposed on one
       another except level 10
    */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build */
    objc = sizeof(objv)/sizeof(objv[0]);
    switch (level) {
    case 502:
        objv[--objc] = Tcl_NewUnicodeObj(((SESSION_INFO_502 *)infoP)->sesi502_transport, -1);
        objv[--objc] = STRING_LITERAL_OBJ("transport");
        /* FALLTHRU */
    case 2:
        objv[--objc] = Tcl_NewUnicodeObj(((SESSION_INFO_2 *)infoP)->sesi2_cltype_name, -1);
        objv[--objc] = STRING_LITERAL_OBJ("cltype_name");
        /* FALLTHRU */
    case 1:
        objv[--objc] = Tcl_NewUnicodeObj(((SESSION_INFO_1 *)infoP)->sesi1_username, -1);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_num_opens);
        objv[--objc] = STRING_LITERAL_OBJ("num_opens");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_time);

        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_idle_time);

        objv[--objc] = STRING_LITERAL_OBJ("idle_time");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_1 *)infoP)->sesi1_user_flags);

        objv[--objc] = STRING_LITERAL_OBJ("user_flags");
        /* FALLTHRU */
    case 0:
        objv[--objc] = Tcl_NewUnicodeObj(((SESSION_INFO_0 *)infoP)->sesi0_cname, -1);
        objv[--objc] = STRING_LITERAL_OBJ("cname");
        break;

    case 10:
        objv[--objc] = Tcl_NewUnicodeObj(((SESSION_INFO_10 *)infoP)->sesi10_cname, -1);
        objv[--objc] = STRING_LITERAL_OBJ("cname");
        objv[--objc] = Tcl_NewUnicodeObj(((SESSION_INFO_10 *)infoP)->sesi10_username, -1);
        objv[--objc] = STRING_LITERAL_OBJ("username");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_10 *)infoP)->sesi10_time);

        objv[--objc] = STRING_LITERAL_OBJ("time");
        objv[--objc] = Tcl_NewLongObj(((SESSION_INFO_10 *)infoP)->sesi10_idle_time);

        objv[--objc] = STRING_LITERAL_OBJ("idle_time");
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
}


#ifndef TWAPI_LEAN
/*
 * Convert user info structure to a list. Returns NULL on error;
 */
Tcl_Obj *ObjFromUSER_INFO(
    Tcl_Interp *interp,
    LPBYTE     infoP,
    DWORD        info_level
    )
{
    USER_INFO_3 *userinfoP = (USER_INFO_3 *) infoP;
    int         objc;
    Tcl_Obj    *objv[60];


    /*
     * Note userinfoP may not point to a USER_INFO_3 struct! It depends
     * on what is specified as info_level. However, the initial fields
     * of the USER_INFO_{0,1,2,3} are the same allowing us to
     * use the following switch statement.
     */

    /* We build from back of array since we would like basic elements
       at front of Tcl list we build (most accessed elements in front */
    objc = sizeof(objv)/sizeof(objv[0]);
#define ADD_LPWSTR_(fld) do {                                          \
        objv[--objc] = Tcl_NewUnicodeObj(userinfoP->usri3_ ## fld, -1); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[--objc] = Tcl_NewLongObj(userinfoP->usri3_ ## fld); \
        objv[--objc] = STRING_LITERAL_OBJ(# fld);                      \
    } while (0)

    switch (info_level) {
    case 3:
        ADD_DWORD_(primary_group_id);
        ADD_LPWSTR_(home_dir_drive);
        ADD_DWORD_(password_expired);
        ADD_LPWSTR_(profile);
        ADD_DWORD_(user_id);
        /* FALL THROUGH */
    case 2:
        ADD_DWORD_(auth_flags);
        ADD_LPWSTR_(usr_comment);
        ADD_LPWSTR_(parms);
        ADD_LPWSTR_(workstations);
        ADD_DWORD_(last_logon);
        ADD_DWORD_(last_logoff);
        ADD_DWORD_(acct_expires);
        ADD_DWORD_(max_storage);
        ADD_DWORD_(units_per_week);
        objv[--objc] = Tcl_NewByteArrayObj(userinfoP->usri3_logon_hours,21);
        objv[--objc] = STRING_LITERAL_OBJ("logon_hours");
        ADD_DWORD_(bad_pw_count);
        ADD_DWORD_(num_logons);
        ADD_LPWSTR_(logon_server);
        ADD_DWORD_(country_code);
        ADD_DWORD_(code_page);
        ADD_LPWSTR_(full_name);
        /* FALL THROUGH */
    case 1:
        ADD_LPWSTR_(script_path);
        ADD_LPWSTR_(password);
        ADD_DWORD_(password_age);
        ADD_DWORD_(priv);
        ADD_DWORD_(flags);
        ADD_LPWSTR_(home_dir);
        ADD_LPWSTR_(comment);
        /* FALL THROUGH */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj((sizeof(objv)/sizeof(objv[0])-objc), &objv[objc]);
}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
/*
 * Convert group info structure to a list. Returns TCL_OK/TCL_ERROR
 * interp may be NULL
 */
Tcl_Obj *ObjFromGROUP_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    GROUP_INFO_3 *groupinfoP = (GROUP_INFO_3 *) infoP;
    int         objc = 0;
    Tcl_Obj    *objv[10];

    /*
     * Note groupinfoP may not point to a GROUP_INFO_3 struct! It depends
     * on what is specified as info_level. However, the initial fields
     * of the GROUP_INFO_{0,1,2,3} are the same allowing us to
     * use the following switch statement.
     */

#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewUnicodeObj(groupinfoP->grpi3_ ## fld, -1); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->grpi3_ ## fld); \
    } while (0)

    switch (info_level) {
    case 3: /* FALL THROUGH */
    case 2:
        ADD_DWORD_(attributes);
        if (info_level == 2) {
            objv[objc++] = STRING_LITERAL_OBJ("group_id");
            objv[objc++] = Tcl_NewIntObj(((GROUP_INFO_2 *)groupinfoP)->grpi2_group_id);
        } else {
            objv[objc++] = STRING_LITERAL_OBJ("group_sid");
            objv[objc++] = ObjFromSIDNoFail(groupinfoP->grpi3_group_sid);
        }
        /* FALL THROUGH */
    case 1:
        ADD_LPWSTR_(comment);
        /* FALL THROUGH */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
/*
 * Convert local info structure to a list. Returns TCL_OK/TCL_ERROR
 * interp may be NULL
 */
Tcl_Obj *ObjFromLOCALGROUP_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    LOCALGROUP_INFO_1 *groupinfoP = (LOCALGROUP_INFO_1 *) infoP;
    int         objc = 0;
    Tcl_Obj    *objv[4];


#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewUnicodeObj(groupinfoP->lgrpi1_ ## fld, -1); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->lgrpi1_ ## fld); \
    } while (0)

    switch (info_level) {
    case 1:
        ADD_LPWSTR_(comment);
        /* FALLTHRU */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromGROUP_USERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    GROUP_USERS_INFO_1 *groupinfoP = (GROUP_USERS_INFO_1 *) infoP;
    int         objc = 0;
    Tcl_Obj    *objv[4];

#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewUnicodeObj(groupinfoP->grui1_ ## fld, -1); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->grui1_ ## fld); \
    } while (0)

    switch (info_level) {
    case 1:
        ADD_DWORD_(attributes);
        /* FALLTHRU */
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromLOCALGROUP_USERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    LOCALGROUP_USERS_INFO_0 *groupinfoP = (LOCALGROUP_USERS_INFO_0 *) infoP;
    int         objc = 0;
    Tcl_Obj    *objv[2];

#define ADD_LPWSTR_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewUnicodeObj(groupinfoP->lgrui0_ ## fld, -1); \
    } while (0)
#define ADD_DWORD_(fld) do {                                          \
        objv[objc++] = STRING_LITERAL_OBJ(# fld);                      \
        objv[objc++] = Tcl_NewLongObj(groupinfoP->lgrui0_ ## fld); \
    } while (0)

    switch (info_level) {
    case 0:
        ADD_LPWSTR_(name);
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

#undef ADD_DWORD_
#undef ADD_LPWSTR_

    return Tcl_NewListObj(objc, objv);
}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
/* Returns TCL_OK/TCL_ERROR. interp may be NULL */
Tcl_Obj *ObjFromLOCALGROUP_MEMBERS_INFO(
    Tcl_Interp *interp,
    LPBYTE      infoP,
    DWORD       info_level
    )
{
    int         objc = 0;
    Tcl_Obj    *objv[6];

    switch (info_level) {
    case 0:
        objv[objc++] = STRING_LITERAL_OBJ("sid");
        objv[objc++] = ObjFromSIDNoFail(((LOCALGROUP_MEMBERS_INFO_0 *)infoP)->lgrmi0_sid);
        break;
    case 1:
    case 2:
        objv[objc++] = STRING_LITERAL_OBJ("sid");
        objv[objc++] = ObjFromSIDNoFail(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_sid);
        objv[objc++] = STRING_LITERAL_OBJ("sidusage");
        objv[objc++] = Tcl_NewLongObj(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_sidusage);
        if (info_level == 1) {
            objv[objc++] = STRING_LITERAL_OBJ("name");
            objv[objc++] = Tcl_NewUnicodeObj(((LOCALGROUP_MEMBERS_INFO_1 *)infoP)->lgrmi1_name, -1);
        } else {
            objv[objc++] = STRING_LITERAL_OBJ("domainandname");
            objv[objc++] = Tcl_NewUnicodeObj(((LOCALGROUP_MEMBERS_INFO_2 *)infoP)->lgrmi2_domainandname, -1);
        }
        break;
    case 3:
        objv[objc++] = STRING_LITERAL_OBJ("domainandname");
        objv[objc++] = Tcl_NewUnicodeObj(((LOCALGROUP_MEMBERS_INFO_3 *)infoP)->lgrmi3_domainandname, -1);
        break;
        
    default:
        TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
        return NULL;
    }

    return Tcl_NewListObj(objc, objv);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN

int TwapiReturnNetEnum(
    Tcl_Interp    *interp,      /* Where result is returned */
    TwapiNetEnumContext *necP
)
{
    Tcl_Obj *enumObj = NULL;
    DWORD      i;
    LPBYTE     p;
    size_t   struct_size;
    Tcl_Obj *objv[4];
    Tcl_Obj *(*objfn)(Tcl_Interp *, LPBYTE, DWORD);

    if (necP->status != NERR_Success && necP->status != ERROR_MORE_DATA)
        return Twapi_AppendSystemError(interp, necP->status);


    enumObj = Tcl_NewListObj(0, NULL);
    switch (necP->tag) {
    case TWAPI_NETENUM_USERINFO:
        objfn = ObjFromUSER_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(USER_INFO_0);
            break;
        case 1:
            struct_size = sizeof(USER_INFO_1);
            break;
        case 2:
            struct_size = sizeof(USER_INFO_2);
            break;
        case 3:
            struct_size = sizeof(USER_INFO_3);
            break;
        default:
            goto invalid_level_error;
        }
        break;

    case TWAPI_NETENUM_GROUPINFO:
        objfn = ObjFromGROUP_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(GROUP_INFO_0);
            break;
        case 1:
            struct_size = sizeof(GROUP_INFO_1);
            break;
        case 2:
            struct_size = sizeof(GROUP_INFO_2);
            break;
        case 3:
            struct_size = sizeof(GROUP_INFO_3);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_LOCALGROUPINFO:
        objfn = ObjFromLOCALGROUP_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(LOCALGROUP_INFO_0);
            break;
        case 1:
            struct_size = sizeof(LOCALGROUP_INFO_1);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_GROUPUSERSINFO:
        objfn = ObjFromGROUP_USERS_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(GROUP_USERS_INFO_0);
            break;
        case 1:
            struct_size = sizeof(GROUP_USERS_INFO_1);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_LOCALGROUPUSERSINFO:
        objfn = ObjFromLOCALGROUP_USERS_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(LOCALGROUP_USERS_INFO_0);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_LOCALGROUPMEMBERSINFO:
        objfn = ObjFromLOCALGROUP_MEMBERS_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_0);
            break;
        case 1:
            struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_1);
            break;
        case 2:
            struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_2);
            break;
        case 3:
            struct_size = sizeof(LOCALGROUP_MEMBERS_INFO_3);
            break;
        default:
            goto invalid_level_error;
        }
        break;

    case TWAPI_NETENUM_SESSIONINFO:
        objfn = ObjFromSESSION_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(SESSION_INFO_0);
            break;
        case 1:
            struct_size = sizeof(SESSION_INFO_1);
            break;
        case 2:
            struct_size = sizeof(SESSION_INFO_2);
            break;
        case 10:
            struct_size = sizeof(SESSION_INFO_10);
            break;
        case 502:
            struct_size = sizeof(SESSION_INFO_502);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_FILEINFO:
        objfn = ObjFromFILE_INFO;
        switch (necP->level) {
        case 2:
            struct_size = sizeof(FILE_INFO_2);
            break;
        case 3:
            struct_size = sizeof(FILE_INFO_3);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_CONNECTIONINFO:
        objfn = ObjFromCONNECTION_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(CONNECTION_INFO_0);
            break;
        case 1:
            struct_size = sizeof(CONNECTION_INFO_1);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_SHAREINFO:
        objfn = ObjFromSHARE_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(SHARE_INFO_0);
            break;
        case 1:
            struct_size = sizeof(SHARE_INFO_1);
            break;
        case 2:
            struct_size = sizeof(SHARE_INFO_2);
            break;
        case 502:
            struct_size = sizeof(SHARE_INFO_502);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    case TWAPI_NETENUM_USEINFO:
        objfn = ObjFromUSE_INFO;
        switch (necP->level) {
        case 0:
            struct_size = sizeof(USE_INFO_0);
            break;
        case 1:
            struct_size = sizeof(USE_INFO_1);
            break;
        case 2:
            struct_size = sizeof(USE_INFO_2);
            break;
        default:
            goto invalid_level_error;
        }
        break;
    default:
        TwapiReturnTwapiError(interp, "Invalid enum type", TWAPI_INVALID_ARGS);
        goto error_return;
    }

    p = necP->netbufP;
    for (i = 0; i < necP->entriesread; ++i, p += struct_size) {
        Tcl_Obj *objP;
        objP = objfn(interp, p, necP->level);
        if (objP == NULL)
            goto error_return;
        Tcl_ListObjAppendElement(interp, enumObj, objP);
    }

    objv[0] = necP->status == ERROR_MORE_DATA ? Tcl_NewIntObj(1) : Tcl_NewIntObj(0);
    objv[1] = ObjFromDWORD_PTR(necP->hresume);
    objv[2] = Tcl_NewLongObj(necP->totalentries);
    objv[3] = enumObj;

    Tcl_SetObjResult(interp, Tcl_NewListObj(4, objv));

    if (necP->netbufP)
        NetApiBufferFree((LPBYTE) necP->netbufP);
    return TCL_OK;

error_return:
    if (necP->netbufP)
        NetApiBufferFree((LPBYTE) necP->netbufP);
    Twapi_FreeNewTclObj(enumObj);

    return TCL_ERROR;

invalid_level_error:
    TwapiReturnTwapiError(interp, "Invalid info level", TWAPI_INVALID_ARGS);
    goto error_return;

}
#endif // TWAPI_LEAN


#ifndef TWAPI_LEAN
int Twapi_NetUserAdd(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPWSTR     name,
    LPWSTR     password,
    DWORD       priv,
    LPWSTR     home_dir,
    LPWSTR     comment,
    DWORD       flags,
    LPWSTR     script_path
    )
{
    NET_API_STATUS status;
    USER_INFO_1    userinfo;
    DWORD          error_parm;
    char          *error_field;

    userinfo.usri1_name        = name;
    userinfo.usri1_password    = password;
    userinfo.usri1_password_age = 0;
    userinfo.usri1_priv        = priv;
    userinfo.usri1_home_dir    = home_dir;
    userinfo.usri1_comment     = comment;
    userinfo.usri1_flags       = flags;
    userinfo.usri1_script_path = script_path;

    status = NetUserAdd(servername, 1, (LPBYTE) &userinfo, &error_parm);
    if (status == NERR_Success)
        return TCL_OK;

    /* Indicate the parameter */
    Tcl_SetResult(interp, "Error adding user account: ", TCL_STATIC);
    error_field = NULL;
    switch (error_parm) {
    case 0: error_field = "user name"; break;
    case 1: error_field = "password";  break;
    case 3: error_field = "privilege level";  break;
    case 4: error_field = "home directory";  break;
    case 5: error_field = "comment";  break;
    case 6: error_field = "flags";  break;
    case 7: error_field = "script path";  break;
    }
    if (error_field) {
        Tcl_AppendResult(interp, " invalid ", error_field, " field. ", NULL);
    }
    return Twapi_AppendSystemError(interp, status);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int TwapiNetUserOrGroupGetInfoHelper(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     name,
    DWORD       level,
    DWORD       type /* 0 - user, 1 - global group, 2 - local group */
    )
{
    NET_API_STATUS status;
    LPBYTE         infoP;
    typedef NET_API_STATUS NET_API_FUNCTION NETGETINFO(LPCWSTR, LPCWSTR, DWORD, LPBYTE*);
    NETGETINFO    *get_fn;
    Tcl_Obj *      (*fmt_fn)(Tcl_Interp *, LPBYTE, DWORD);


    switch (level) {
    case 0:
    case 1:
    case 2:
    case 3:
        switch (type) {
        case 0:
            get_fn = NetUserGetInfo;
            fmt_fn = ObjFromUSER_INFO;
            break;
        case 1:
            get_fn = NetGroupGetInfo;
            fmt_fn = ObjFromGROUP_INFO;
            break;
        case 2:
            get_fn = NetLocalGroupGetInfo;
            fmt_fn = ObjFromLOCALGROUP_INFO;
            break;
        default:
            Tcl_SetResult(interp, "Internal error: bad type passed to TwapiNetUserOrGroupGetInfoHelper", TCL_STATIC);
            return TCL_ERROR;
        }
        status = (*get_fn)(servername, name, level, &infoP);
        if (status != NERR_Success) {
            Tcl_SetResult(interp,
                          "Could not retrieve global user or group information: ", TCL_STATIC);
            return Twapi_AppendSystemError(interp, status);
        }
        Tcl_SetObjResult(interp, (*fmt_fn)(interp, infoP, level));
        NetApiBufferFree(infoP);
        break;

    default:
        Tcl_SetResult(interp, "Invalid or unsupported user or group information level specified", TCL_STATIC);
        return TCL_ERROR;
    }

    return TCL_OK;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
/* Wrapper around NetUserGetInfo */
int Twapi_NetUserGetInfo(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     username,
    DWORD       level
    )
{
    return TwapiNetUserOrGroupGetInfoHelper(interp, servername, username, level, 0);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN

int Twapi_NetUserSetInfoDWORD(
    int func,
    LPCWSTR servername,
    LPCWSTR username,
    DWORD   dw
    )
{
    union {
        USER_INFO_1005 u1005;
        USER_INFO_1008 u1008;
        USER_INFO_1010 u1010;
        USER_INFO_1017 u1017;
        USER_INFO_1024 u1024;
    } userinfo;

    switch (func) {
    case 1005:
        userinfo.u1005.usri1005_priv = dw;
        break;
    case 1008:
        userinfo.u1008.usri1008_flags = dw;
        break;
    case 1010:
        userinfo.u1010.usri1010_auth_flags = dw;
        break;
    case 1017:
        userinfo.u1017.usri1017_acct_expires = dw;
        break;
    case 1024:
        userinfo.u1024.usri1024_country_code = dw;
        break;
    default:
        return ERROR_INVALID_PARAMETER;
    }
    return NetUserSetInfo(servername, username, func, (LPBYTE) &userinfo, NULL);
}


int Twapi_NetUserSetInfoLPWSTR(
    int func,
    LPCWSTR servername,
    LPCWSTR username,
    LPWSTR value
    )
{
    union {
        USER_INFO_0 u0;
        USER_INFO_1003 u1003;
        USER_INFO_1006 u1006;
        USER_INFO_1007 u1007;
        USER_INFO_1009 u1009;
        USER_INFO_1011 u1011;
        USER_INFO_1052 u1052;
        USER_INFO_1053 u1053;
    } userinfo;

    switch (func) {
    case 0:
        userinfo.u0.usri0_name = value;
        break;
    case 1003:
        userinfo.u1003.usri1003_password = value;
        break;
    case 1006:
        userinfo.u1006.usri1006_home_dir = value;
        break;
    case 1007:
        userinfo.u1007.usri1007_comment = value;
        break;
    case 1009:
        userinfo.u1009.usri1009_script_path = value;
        break;
    case 1011:
        userinfo.u1011.usri1011_full_name = value;
        break;
    case 1052:
        userinfo.u1052.usri1052_profile = value;
        break;
    case 1053:
        userinfo.u1053.usri1053_home_dir_drive = value;
        break;
    default:
        return ERROR_INVALID_PARAMETER;
    }
    return NetUserSetInfo(servername, username, func, (LPBYTE) &userinfo, NULL);
}
#endif // TWAPI_LEAN



#ifndef TWAPI_LEAN
int Twapi_NetGroupGetInfo(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     groupname,
    DWORD       level
    )
{
    return TwapiNetUserOrGroupGetInfoHelper(interp, servername, groupname, level, 1);
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
int Twapi_NetLocalGroupGetInfo(
    Tcl_Interp *interp,
    LPCWSTR     servername,
    LPCWSTR     groupname,
    DWORD       level
    )
{
    if (level != 1) {
        Tcl_SetResult(interp, "Invalid or unsupported user or group information level specified", TCL_STATIC);
        return TCL_ERROR;
    }
    return TwapiNetUserOrGroupGetInfoHelper(interp, servername, groupname, level, 2);
}
#endif // TWAPI_LEAN


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
LSA_HANDLE Twapi_LsaOpenPolicy(
    PLSA_UNICODE_STRING systemP,
    PLSA_OBJECT_ATTRIBUTES objattrP,
    DWORD access
)
{
    LSA_HANDLE h;
    NTSTATUS ntstatus;

    if (systemP->Length == 0)
        systemP = NULL;

    ntstatus = LsaOpenPolicy(systemP, objattrP, access, &h);
    if (ntstatus != STATUS_SUCCESS) {
        SetLastError(TwapiNTSTATUSToError(ntstatus));
        h = NULL;
    }

    return h;
}
#endif // TWAPI_LEAN

#ifndef TWAPI_LEAN
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
        return TwapiReturnTwapiError(interp, NULL, TWAPI_INVALID_ARGS);

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

#ifndef TWAPI_LEAN
int Twapi_LsaQueryInformationPolicy (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    void    *buf;
    NTSTATUS ntstatus;
    int      retval;
    Tcl_Obj  *objs[5];
    LSA_HANDLE lsaH;
    int        infoclass;

    if (objc != 2 ||
        ObjToOpaque(interp, objv[0], (void **) &lsaH, "LSA_HANDLE") != TCL_OK ||
        Tcl_GetLongFromObj(interp, objv[1], &infoclass) != TCL_OK) {
        return TwapiReturnTwapiError(interp, NULL, TWAPI_INVALID_ARGS);
    }    

    ntstatus = LsaQueryInformationPolicy(lsaH, infoclass, &buf);
    if (ntstatus != STATUS_SUCCESS) {
        return Twapi_AppendSystemError(interp,
                                       TwapiNTSTATUSToError(ntstatus));
    }

    retval = TCL_OK;
    switch (infoclass) {
    case PolicyAccountDomainInformation:
        objs[0] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_ACCOUNT_DOMAIN_INFO *) buf)->DomainName)
            );
        objs[1] = ObjFromSIDNoFail(((POLICY_ACCOUNT_DOMAIN_INFO *) buf)->DomainSid);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objs));
        break;

    case PolicyDnsDomainInformation:
        objs[0] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_DNS_DOMAIN_INFO *) buf)->Name)
            );
        objs[1] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_DNS_DOMAIN_INFO *) buf)->DnsDomainName)
            );
        objs[2] = ObjFromLSA_UNICODE_STRING(
            &(((POLICY_DNS_DOMAIN_INFO *) buf)->DnsForestName)
            );
        objs[3] = ObjFromUUID(
            (UUID *) &(((POLICY_DNS_DOMAIN_INFO *) buf)->DomainGuid)
            );
        objs[4] = ObjFromSIDNoFail(((POLICY_DNS_DOMAIN_INFO *) buf)->Sid);
        Tcl_SetObjResult(interp, Tcl_NewListObj(5, objs));

        break;

    default:
        Tcl_SetResult(interp, "Invalid or unsupported information class passed to Twapi_LsaQueryInformationPolicy", TCL_STATIC);
        retval = TCL_ERROR;
    }

    LsaFreeMemory(buf);

    return retval;
}
#endif // TWAPI_LEAN

#if 0
/* Not explicitly visible in Win2K and XP SP0 so we use CheckTokenMembership
   directly for which this is a wrapper */
BOOL IsUserAnAdmin();
#endif


