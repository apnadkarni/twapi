/* 
 * Copyright (c) 2007-2016 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to CryptoAPI */
/*
 * TBD - GetCryptProvFromCert, CryptVerifyCertificateSignature(Ex),
 * TBD - CryptRetrieveObjectByUrl, Crypt(Ex|Im)portKey
 */


#include "twapi.h"
#include "twapi_crypto.h"
#include <mscat.h>

#ifndef TWAPI_SINGLE_MODULE
HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

static Tcl_Obj *ObjFromCERT_EXTENSIONS(int nexts, CERT_EXTENSION *extP);
static TCL_RESULT ParseCERT_EXTENSIONS(
    TwapiInterpContext *ticP,
    Tcl_Obj *extsObj,
    DWORD *nextsP,
    CERT_EXTENSION **extsPP
    );
static TCL_RESULT TwapiCryptDecodeObject(
    Tcl_Interp *interp,
    void *poid, /* Either a Tcl_Obj or a #define X509* int value */
    void *penc,
    DWORD nenc,
    Tcl_Obj **objPP);
static BOOL WINAPI TwapiCertFreeCertificateChain(
  PCCERT_CHAIN_CONTEXT chainP
    );

/* 
 * Macro to define functions for ref counting different type of pointers.
 * We need to protect against invalid access since the underlying API's
 * themselves do not seem to check for validity of the handles/pointers
 * (double frees for example will crash)
 */
#define DEFINE_COUNTED_PTR_FUNCS(ptrtype_, tag_) \
void TwapiRegister ## ptrtype_ (Tcl_Interp *interp, ptrtype_ ptr)       \
{                                                                       \
    if (TwapiRegisterCountedPointer(interp, (void*)ptr, tag_) != TCL_OK) \
        Tcl_Panic("Failed to register " #ptrtype_ ": %s", Tcl_GetStringResult(interp)); \
}                                                                       \
                                                                        \
void TwapiRegister ## ptrtype_ ## Tic (TwapiInterpContext *ticP, ptrtype_ ptr) \
{                                                                       \
    if (TwapiRegisterCountedPointerTic(ticP, (void*)ptr, tag_) != TCL_OK) \
        Tcl_Panic("Failed to register " #ptrtype_ ": %s", Tcl_GetStringResult(ticP->interp)); \
}                                                                       \
                                                                        \
TCL_RESULT TwapiUnregister ## ptrtype_ (Tcl_Interp *interp, ptrtype_ ptr) \
{                                                                       \
    return TwapiUnregisterPointer(interp, (void*) ptr, tag_);            \
}                                                                       \
                                                                        \
TCL_RESULT TwapiUnregister ## ptrtype_ ## Tic(TwapiInterpContext *ticP, ptrtype_ ptr) \
{                                                                       \
    return TwapiUnregisterPointerTic(ticP, (void*)ptr, tag_);           \
}

DEFINE_COUNTED_PTR_FUNCS(PCCERT_CONTEXT, CertFreeCertificateContext)
DEFINE_COUNTED_PTR_FUNCS(HCERTSTORE, CertCloseStore)
DEFINE_COUNTED_PTR_FUNCS(HCRYPTMSG, CryptMsgClose)
DEFINE_COUNTED_PTR_FUNCS(PCCERT_CHAIN_CONTEXT, CertFreeCertificateChain)
DEFINE_COUNTED_PTR_FUNCS(PCCRL_CONTEXT, CertFreeCRLContext)
DEFINE_COUNTED_PTR_FUNCS(PCCTL_CONTEXT, CertFreeCTLContext)
DEFINE_COUNTED_PTR_FUNCS(HCATADMIN, CryptCATAdminReleaseContext)
DEFINE_COUNTED_PTR_FUNCS(HCATINFO, CryptCATAdminReleaseCatalogContext)

/* The following types do not seem to need protection against invalid
   or double freeing since the API appears to check for this. However,
   not sure if all CSP's will behave that way so protect them as well.
*/
DEFINE_COUNTED_PTR_FUNCS(HCRYPTKEY, CryptDestroyKey)
DEFINE_COUNTED_PTR_FUNCS(HCRYPTPROV, CryptReleaseContext)

/* This function exists only to make the return value compatible with
   other free functions */
static BOOL WINAPI TwapiCertFreeCertificateChain(
  PCCERT_CHAIN_CONTEXT chainP
)
{
    CertFreeCertificateChain(chainP);
    return 1;
}

/* This function exists only to make the return value compatible with
   other free functions */
static BOOL WINAPI TwapiCryptCATAdminReleaseContext(HCATADMIN h)
{
    return CryptCATAdminReleaseContext(h, 0);
}


/* RtlGenRandom in base provides this but this is much faster if already have
   a HCRYPTPROV 
*/
int Twapi_CryptGenRandom(Tcl_Interp *interp, HCRYPTPROV provH, DWORD len)
{
    BYTE buf[256];

    if (len > sizeof(buf)) {
        Tcl_SetObjErrorCode(interp,
                            Twapi_MakeTwapiErrorCodeObj(TWAPI_INTERNAL_LIMIT));
        ObjSetStaticResult(interp, "Too many random bytes requested.");
        return TCL_ERROR;
    }

    if (CryptGenRandom(provH, len, buf)) {
        ObjSetResult(interp, ObjFromByteArray(buf, len));
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
}

static Tcl_Obj *ObjFromCRYPT_BLOB(const CRYPT_DATA_BLOB *blobP)
{
    if (blobP && blobP->cbData && blobP->pbData)
        return ObjFromByteArray(blobP->pbData, blobP->cbData);
    else
        return ObjFromEmptyString();
}

static Tcl_Obj *ObjFromCRYPT_BIT_BLOB(CRYPT_BIT_BLOB *blobP)
{
    Tcl_Obj *objs[2];
    if (blobP->cbData && blobP->pbData) {
        objs[0] = ObjFromByteArray(blobP->pbData, blobP->cbData);
        objs[1] = ObjFromDWORD(blobP->cUnusedBits);
    } else {
        objs[0] = ObjFromEmptyString();
        objs[1] = ObjFromDWORD(0);
    }
    return ObjNewList(2, objs);
}

Tcl_Obj *ObjFromCERT_NAME_BLOB(CERT_NAME_BLOB *blobP, DWORD flags)
{
    int len;
    WCHAR *wP;
    Tcl_Obj *objP;
    WCHAR buf[200];

    len = CertNameToStrW(X509_ASN_ENCODING, blobP, flags, NULL, 0);
    if (len == 0)
        return ObjFromEmptyString();
    if (len > ARRAYSIZE(buf))
        wP = TwapiAlloc(len*sizeof(WCHAR));
    else
        wP = buf;
    len = CertNameToStrW(X509_ASN_ENCODING, blobP, flags, wP, len) - 1;
    objP = ObjFromUnicodeN(wP, len);
    if (wP != buf)
        TwapiFree(wP);
    return objP;
}

static Tcl_Obj *ObjFromCERT_NAME_VALUE(CERT_NAME_VALUE *valP)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromDWORD(valP->dwValueType);
    objs[1] = ObjFromCRYPT_BLOB(&valP->Value);
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromCERT_NAME_VALUE_Unicode(CERT_NAME_VALUE *valP)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromDWORD(valP->dwValueType);
    objs[1] = ObjFromUnicodeN((WCHAR *)valP->Value.pbData, valP->Value.cbData/sizeof(WCHAR));
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromCERT_ALT_NAME_ENTRY(CERT_ALT_NAME_ENTRY *caneP)
{
    Tcl_Obj *objs[2];
    int nobjs;

    nobjs = 2;
    switch (caneP->dwAltNameChoice) {
    case CERT_ALT_NAME_RFC822_NAME: /* FALLTHRU */
    case CERT_ALT_NAME_DNS_NAME:    /* FALLTHRU */
    case CERT_ALT_NAME_URL:
        objs[1] = ObjFromUnicode(caneP->pwszURL);
        break;
    case CERT_ALT_NAME_OTHER_NAME:
        objs[0] = ObjFromString(caneP->pOtherName->pszObjId);
        objs[1] = ObjFromByteArray(caneP->pOtherName->Value.pbData,
                                   caneP->pOtherName->Value.cbData);
        objs[1] = ObjNewList(2, objs);
        break;
    case CERT_ALT_NAME_DIRECTORY_NAME: /* FALLTHRU */
    case CERT_ALT_NAME_IP_ADDRESS:
        objs[1] = ObjFromByteArray(caneP->IPAddress.pbData,
                                   caneP->IPAddress.cbData);
        break;
    case CERT_ALT_NAME_REGISTERED_ID:
        objs[1] = ObjFromString(caneP->pszRegisteredID);
        break;
    default:
        nobjs = 1;              /* Only report type */
    }

    objs[0] = ObjFromDWORD(caneP->dwAltNameChoice);
    return ObjNewList(nobjs, objs);
}

static Tcl_Obj *ObjFromCERT_ALT_NAME_INFO(CERT_ALT_NAME_INFO *infP)
{
    Tcl_Obj *objP;
    DWORD i;

    objP = ObjNewList(infP->cAltEntry, NULL);
    for (i = 0; i < infP->cAltEntry; ++i) {
        ObjAppendElement(NULL, objP,
                         ObjFromCERT_ALT_NAME_ENTRY(&infP->rgAltEntry[i]));
    }
    return objP;
}

static Tcl_Obj *ObjFromCRYPT_KEY_PROV_INFO(CRYPT_KEY_PROV_INFO *infP)
{
    Tcl_Obj *objs[6];
    DWORD i;
    
    objs[0] = ObjFromUnicode(infP->pwszContainerName);
    objs[1] = ObjFromUnicode(infP->pwszProvName);
    objs[2] = ObjFromDWORD(infP->dwProvType);
    objs[3] = ObjFromDWORD(infP->dwFlags);
    objs[4] = ObjNewList(infP->cProvParam, NULL);
    objs[5] = ObjFromDWORD(infP->dwKeySpec);
    
    if (infP->rgProvParam) {
        for (i=0; i < infP->cProvParam; ++i) {
            /* TBD - for now just return raw bytes. */
            Tcl_Obj *parObjs[3];
            parObjs[0] = ObjFromDWORD(infP->rgProvParam[i].dwParam);
            parObjs[1] = ObjFromByteArray(infP->rgProvParam[i].pbData,
                                         infP->rgProvParam[i].cbData);
            parObjs[2] = ObjFromDWORD(infP->rgProvParam[i].dwFlags);
            ObjAppendElement(NULL, objs[4], ObjNewList(3, parObjs));
        }
    }

    return ObjNewList(6, objs);
}

static Tcl_Obj *ObjFromCRYPT_ALGORITHM_IDENTIFIER(CRYPT_ALGORITHM_IDENTIFIER *algP)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromString(algP->pszObjId);
    objs[1] = ObjFromCRYPT_BLOB(&algP->Parameters);
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromCERT_PUBLIC_KEY_INFO(CERT_PUBLIC_KEY_INFO *cpiP)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromCRYPT_ALGORITHM_IDENTIFIER(&cpiP->Algorithm);
    objs[1] = ObjFromCRYPT_BIT_BLOB(&cpiP->PublicKey);
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromCERT_POLICY_CONSTRAINTS_INFO(CERT_POLICY_CONSTRAINTS_INFO  *cpciP)
{
    Tcl_Obj *objs[4];
    objs[0] = ObjFromBoolean(cpciP->fRequireExplicitPolicy);
    objs[1] = ObjFromDWORD(cpciP->dwRequireExplicitPolicySkipCerts);
    objs[2] = ObjFromBoolean(cpciP->fInhibitPolicyMapping);
    objs[3] = ObjFromDWORD(cpciP->dwInhibitPolicyMappingSkipCerts);
    return ObjNewList(4, objs);
}

static Tcl_Obj *ObjFromCRYPT_OID_INFO(PCCRYPT_OID_INFO coiP)
{
    Tcl_Obj *objs[5];

    objs[0] = ObjFromString(coiP->pszOID);
    objs[1] = ObjFromUnicode(coiP->pwszName);
    objs[2] = ObjFromDWORD(coiP->dwGroupId);
    objs[3] = ObjFromDWORD(coiP->dwValue);
    objs[4] = ObjFromCRYPT_BLOB(&coiP->ExtraInfo);

    return ObjNewList(5, objs);
}

static Tcl_Obj *ObjFromCRYPT_ATTRIBUTE(CRYPT_ATTRIBUTE *caP)
{
    Tcl_Obj *objs[2];
    DWORD i;

    objs[0] = ObjFromString(caP->pszObjId);
    objs[1] = ObjNewList(caP->cValue, NULL);
    for (i = 0; i < caP->cValue; ++i) {
        Tcl_Obj *attrObj;
        /* Try to decode and if cannot do so, return plain byte array */
        if (TwapiCryptDecodeObject(NULL, objs[0],
                                   caP->rgValue[i].pbData, caP->rgValue[i].cbData,
                                   &attrObj) == TCL_OK)
            ObjAppendElement(NULL, objs[1], attrObj);
        else
            ObjAppendElement(NULL, objs[1], ObjFromCRYPT_BLOB(&caP->rgValue[i]));
    }
    return ObjNewList(2, objs);
}


static Tcl_Obj *ObjFromCERT_REQUEST_INFO(CERT_REQUEST_INFO *criP)
{
    Tcl_Obj *objs[4];
    DWORD i;

    objs[0] = ObjFromDWORD(criP->dwVersion);
    objs[1] = ObjFromCRYPT_BLOB(&criP->Subject);
    objs[2] = ObjFromCERT_PUBLIC_KEY_INFO(&criP->SubjectPublicKeyInfo);
    objs[3] = ObjNewList(criP->cAttribute, NULL);
    for (i = 0; i < criP->cAttribute; ++i) {
        ObjAppendElement(NULL, objs[3], ObjFromCRYPT_ATTRIBUTE(&criP->rgAttribute[i]));
    }
    return ObjNewList(4, objs);
}


static Tcl_Obj *ObjFromCERT_TRUST_STATUS(const CERT_TRUST_STATUS *ctsP)
{
    Tcl_Obj *objs[2];
    objs[0] = ObjFromDWORD(ctsP->dwErrorStatus);
    objs[1] = ObjFromDWORD(ctsP->dwInfoStatus);
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromCERT_CHAIN_ELEMENT(Tcl_Interp *interp, CERT_CHAIN_ELEMENT *cceP)
{
    Tcl_Obj *objs[2];
    PCCERT_CONTEXT certP;

    certP = CertDuplicateCertificateContext(cceP->pCertContext);
    TwapiRegisterPCCERT_CONTEXT(interp, certP);
    objs[0] = ObjFromOpaque((void*) certP, "PCCERT_CONTEXT");
    objs[1] = ObjFromCERT_TRUST_STATUS(&cceP->TrustStatus);
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromCERT_SIMPLE_CHAIN(Tcl_Interp *interp,
                                         CERT_SIMPLE_CHAIN *cscP)
{
    Tcl_Obj *objs[2];
    DWORD dw;
    objs[0] = ObjFromCERT_TRUST_STATUS(&cscP->TrustStatus);
    objs[1] = ObjNewList(0, NULL);
    for (dw = 0; dw < cscP->cElement; ++dw)
        ObjAppendElement(interp, objs[1], ObjFromCERT_CHAIN_ELEMENT(interp, cscP->rgpElement[dw]));
    return ObjNewList(2, objs);
}


/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseCRYPT_BLOB(
    TwapiInterpContext *ticP,
    Tcl_Obj *objP,
    CRYPT_DATA_BLOB *blobP
    )
{
    void *pv;
    int   len;
    pv = ObjToByteArray(objP, &len);
    if (len)
        blobP->pbData = MemLifoCopy(ticP->memlifoP, pv, len);
    else
        blobP->pbData = NULL;
    blobP->cbData = len;
    return TCL_OK;
}

/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseCRYPT_BIT_BLOB(
    TwapiInterpContext *ticP,
    Tcl_Obj *pkObj,
    CRYPT_BIT_BLOB *blobP
    )
{
    Tcl_Obj **objs;
    int       nobjs;
    Tcl_Interp *interp = ticP->interp;

    if (ObjGetElements(NULL, pkObj, &nobjs, &objs) == TCL_OK) {
        if (nobjs) {
            if (TwapiGetArgsEx(ticP, nobjs, objs, GETBA(blobP->pbData, blobP->cbData),
                               GETINT(blobP->cUnusedBits), ARGEND) == TCL_OK &&
                blobP->cUnusedBits <= 7)
                return TCL_OK;
        } else {
            blobP->pbData = NULL;
            blobP->cbData = 0;
            return TCL_OK;
        }
    }
    ObjSetStaticResult(interp, "Invalid CRYPT_BIT_BLOB structure");
    return TCL_ERROR;
}

/* Returns CERT_ALT_NAME_ENTRY structure in *caneP
   using memory from ticP->memlifo. Caller responsible for storage
   in both success and error cases
*/
static TCL_RESULT ParseCERT_ALT_NAME_ENTRY(
    TwapiInterpContext *ticP,
    Tcl_Obj *nameObj,
    CERT_ALT_NAME_ENTRY *caneP
    )
{
    Tcl_Obj **objs;
    int nobjs;
    DWORD name_type;
    int n;
    Tcl_Obj **otherObjs;
    void *pv;
    
    if (ObjGetElements(NULL, nameObj, &nobjs, &objs) != TCL_OK ||
        nobjs != 2 ||
        ObjToDWORD(NULL, objs[0], &name_type) != TCL_OK) {
        goto format_error;
    }

    switch (name_type) {
    case CERT_ALT_NAME_RFC822_NAME: /* FALLTHROUGH */
    case CERT_ALT_NAME_DNS_NAME: /* FALLTHROUGH */
    case CERT_ALT_NAME_URL:
        pv = ObjToUnicodeN(objs[1], &n);
        caneP->pwszRfc822Name = MemLifoCopy(ticP->memlifoP, pv, sizeof(WCHAR) * (n+1));
        break;
    case CERT_ALT_NAME_REGISTERED_ID:
        pv = ObjToStringN(objs[1], &n);
        caneP->pszRegisteredID = MemLifoCopy(ticP->memlifoP, pv, n+1);
        break;
    case CERT_ALT_NAME_OTHER_NAME:
        caneP->pOtherName = MemLifoAlloc(ticP->memlifoP, sizeof(CERT_OTHER_NAME), NULL);
        if (ObjGetElements(NULL, objs[1], &n, &otherObjs) != TCL_OK ||
            n != 2)
            goto format_error;
        pv = ObjToStringN(otherObjs[0], &n);
        caneP->pOtherName->pszObjId = MemLifoCopy(ticP->memlifoP, pv, n+1);
        pv = ObjToByteArray(otherObjs[1], &n);
        caneP->pOtherName->Value.pbData = MemLifoCopy(ticP->memlifoP, pv, n);
        caneP->pOtherName->Value.cbData = n;
        break;
    case CERT_ALT_NAME_DIRECTORY_NAME: /* FALLTHRU */
    case CERT_ALT_NAME_IP_ADDRESS: /* FALLTHRU */
        pv = ObjToByteArray(objs[1], &n);
        caneP->IPAddress.pbData = MemLifoCopy(ticP->memlifoP, pv, n);
        caneP->IPAddress.cbData = n;
        break;
        
    default:
        goto format_error;
    }

    caneP->dwAltNameChoice = name_type;
    return TCL_OK;

format_error:
    ObjSetResult(ticP->interp,
                 Tcl_ObjPrintf("Invalid or unsupported name format \"%s\"", ObjToString(nameObj)));
    return TCL_ERROR;
}

static TCL_RESULT ParseCERT_NAME_VALUE(
    TwapiInterpContext *ticP,
    Tcl_Obj *namevalObj,
    CERT_NAME_VALUE *cnvP
    )
{
    Tcl_Obj **objs;
    int nobjs;

    if (ObjGetElements(NULL, namevalObj, &nobjs, &objs) == TCL_OK &&
        nobjs == 2 &&
        ObjToDWORD(NULL, objs[0], &cnvP->dwValueType) == TCL_OK &&
        ParseCRYPT_BLOB(ticP, objs[1], &cnvP->Value) == TCL_OK) {
        return TCL_OK;
    } else
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid CERT_NAME_VALUE");
}

static TCL_RESULT ParseCERT_NAME_VALUE_Unicode(
    TwapiInterpContext *ticP,
    Tcl_Obj *namevalObj,
    CERT_NAME_VALUE *cnvP
    )
{
    int nchars;

    if (TwapiGetArgsExObj(ticP, namevalObj, GETINT(cnvP->dwValueType),
                          GETWSTRN(cnvP->Value.pbData, nchars), ARGEND)
        ==  TCL_OK) {
        cnvP->Value.cbData = nchars * sizeof(WCHAR);
        return TCL_OK;
    } else
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid CERT_NAME_VALUE");
}




/* Returns CERT_ALT_NAME_INFO structure in *caniP
   using memory from ticP->memlifo. Caller responsible for storage
   in both success and error cases
*/
static TCL_RESULT ParseCERT_ALT_NAME_INFO(
    TwapiInterpContext *ticP,
    Tcl_Obj *altnameObj,
    CERT_ALT_NAME_INFO *caniP
    )
{
    Tcl_Obj **nameObjs;
    int       nnames, i;
    TCL_RESULT res;
    CERT_ALT_NAME_ENTRY *entriesP;

    if ((res = ObjGetElements(ticP->interp, altnameObj, &nnames, &nameObjs))
        != TCL_OK)
        return res;

    if (nnames == 0) {
        caniP->cAltEntry = 0;
        caniP->rgAltEntry = NULL;
        return TCL_OK;
    }

    entriesP = MemLifoAlloc(ticP->memlifoP, nnames * sizeof(*entriesP), NULL);

    for (i = 0; i < nnames; ++i) {
        res = ParseCERT_ALT_NAME_ENTRY(ticP, nameObjs[i], &entriesP[i]);
        if (res != TCL_OK)
            return res;
    }

    caniP->cAltEntry = nnames;
    caniP->rgAltEntry = entriesP;
    return TCL_OK;
}


/* Returns CERT_ENHKEY_USAGE structure in *cekuP
   using memory from ticP->memlifo. Caller responsible for storage
   in both success and error cases
*/
static TCL_RESULT ParseCERT_ENHKEY_USAGE(
    TwapiInterpContext *ticP,
    Tcl_Obj *cekuObj,
    CERT_ENHKEY_USAGE *cekuP
    )
{
    Tcl_Obj **objs;
    int       nobjs;
    int       i, n;
    char     *p;
    TCL_RESULT res;

    if ((res = ObjGetElements(ticP->interp, cekuObj, &nobjs, &objs)) != TCL_OK)
        return res;

    cekuP->cUsageIdentifier = nobjs;
    cekuP->rgpszUsageIdentifier = MemLifoAlloc(ticP->memlifoP,
                                               nobjs * sizeof(cekuP->rgpszUsageIdentifier[0]),
                                               NULL);
    for (i = 0; i < nobjs; ++i) {
        p = ObjToStringN(objs[i], &n);
        cekuP->rgpszUsageIdentifier[i] = MemLifoCopy(ticP->memlifoP, p, n+1);
    }

    return TCL_OK;
}

/* Returns CERT_CHAIN_PARA structure using memory from ticP->memlifo.
 * Caller responsible for storage in both success and error cases
 */
static TCL_RESULT ParseCERT_CHAIN_PARA(
    TwapiInterpContext *ticP,
    Tcl_Obj *paramObj,
    CERT_CHAIN_PARA *paramP
    )
{
    Tcl_Obj **objs;
    int       i, n;
    Tcl_Interp *interp = ticP->interp;

    /*
     * CERT_CHAIN_PARA is a list, currently containing exactly one element -
     * the RequestedUsage field. This is a list of two elements, a DWORD
     * indicating the boolean operation and an array of CERT_ENHKEY_USAGE
     */
    if (ObjGetElements(NULL, paramObj, &n, &objs) != TCL_OK ||
        n != 1 ||
        ObjGetElements(NULL, objs[0], &n, &objs) != TCL_OK ||
        n != 2 ||
        ObjToInt(NULL, objs[0], &i) != TCL_OK)
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid CERT_CHAIN_PARA");

    ZeroMemory(paramP, sizeof(paramP));

    paramP->cbSize = sizeof(*paramP);
    paramP->RequestedUsage.dwType = i;
    if (ParseCERT_ENHKEY_USAGE(ticP, objs[1], &paramP->RequestedUsage.Usage)
        != TCL_OK)
        return TCL_ERROR;

    return TCL_OK;
}

/* Returns CERT_CHAIN_POLICY_PARA structure using memory from ticP->memlifo.
 * Caller responsible for releasing storage in both success and error cases
 */
static TCL_RESULT ParseCERT_CHAIN_POLICY_PARA(
    TwapiInterpContext *ticP,
    Tcl_Obj *paramObj,
    CERT_CHAIN_POLICY_PARA **policy_paramPP
    )
{
    int       flags;
    CERT_CHAIN_POLICY_PARA *policy_paramP;
    
    if (ObjToInt(NULL, paramObj, &flags) != TCL_OK)
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid CERT_CHAIN_POLICY_PARA (expected integer value)");

    policy_paramP = MemLifoAlloc(ticP->memlifoP, sizeof(*policy_paramP), NULL);
    ZeroMemory(policy_paramP, sizeof(*policy_paramP));
    policy_paramP->cbSize = sizeof(*policy_paramP);
    policy_paramP->dwFlags = flags;
    policy_paramP->pvExtraPolicyPara = NULL;
    *policy_paramPP = policy_paramP;
    return TCL_OK;
}

/* Returns CERT_CHAIN_POLICY_PARA structure using memory from ticP->memlifo.
 * Caller responsible for releasing storage in both success and error cases
 */
static TCL_RESULT ParseCERT_CHAIN_POLICY_PARA_SSL(
    TwapiInterpContext *ticP,
    Tcl_Obj *paramObj,
    CERT_CHAIN_POLICY_PARA **policy_paramPP
    )
{
    Tcl_Obj **objs;
    int       flags, n;
    CERT_CHAIN_POLICY_PARA *policy_paramP;
    
    /*
     * CERT_CHAIN_POLICY_PARA is a list, currently containing one or two
     * elements - a flags field, and optionally an
     * SSL_EXTRA_CERT_CHAIN_POLICY_PARA field.
     */
    if (ObjGetElements(NULL, paramObj, &n, &objs) != TCL_OK ||
        (n != 1 && n != 2) ||
        ObjToInt(NULL, objs[0], &flags) != TCL_OK)
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid CERT_CHAIN_POLICY_PARA");

    policy_paramP = MemLifoAlloc(ticP->memlifoP, sizeof(*policy_paramP), NULL);
    ZeroMemory(policy_paramP, sizeof(*policy_paramP));
    policy_paramP->cbSize = sizeof(*policy_paramP);
    policy_paramP->dwFlags = flags;

    if (n == 1)
        policy_paramP->pvExtraPolicyPara = NULL;
    else {
        SSL_EXTRA_CERT_CHAIN_POLICY_PARA *sslP;
        /* Parse the SSL_EXTRA_CERT_CHAIN_POLICY_PARA */
        sslP = MemLifoAlloc(ticP->memlifoP, sizeof(*sslP), NULL);
        sslP->cbSize = sizeof(*sslP);
        if (TwapiGetArgsExObj(ticP, objs[1],
                           GETINT(sslP->dwAuthType),
                           GETINT(sslP->fdwChecks),
                           GETWSTR(sslP->pwszServerName),
                           ARGEND) != TCL_OK)
            goto error_return;
        /* Although docs say pwszServerName is ignored for AUTHTYPE_CLIENT,
           it appears not to be so. It seems to do a name check if the
           field is not NULL */
        if (sslP->dwAuthType == AUTHTYPE_CLIENT && sslP->pwszServerName[0] == 0)
            sslP->pwszServerName = NULL;
        policy_paramP->pvExtraPolicyPara = sslP;
    }

    *policy_paramPP = policy_paramP;
    return TCL_OK;

error_return:
    return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid SSL_EXTRA_CERT_CHAIN_POLICY_PARA");

}

/* Fill in CRYPT_ALGORITHM_IDENTIFIER structure in *algidP
   using memory from ticP->memlifo. Caller responsible for storage
   in both success and error cases
*/
static TCL_RESULT ParseCRYPT_ALGORITHM_IDENTIFIER(
    TwapiInterpContext *ticP,
    Tcl_Obj *algObj,
    CRYPT_ALGORITHM_IDENTIFIER *algidP
    )
{
    TCL_RESULT res;
    Tcl_Obj **objs;
    int       n, nobjs;
    char     *p;

    if ((res = ObjGetElements(ticP->interp, algObj, &nobjs, &objs)) != TCL_OK)
        return res;

    if (nobjs != 1 && nobjs != 2) {
        ObjSetStaticResult(ticP->interp, "Invalid algorithm identifier format or unsupported parameters");
        return TCL_ERROR;
    }

    p = ObjToStringN(objs[0], &n);
    algidP->pszObjId = MemLifoCopy(ticP->memlifoP, p, n+1);
    if (nobjs == 1) {
        algidP->Parameters.cbData = 0;
        algidP->Parameters.pbData = 0;
    } else {
        p = ObjToByteArray(objs[1], &n);
        algidP->Parameters.pbData = MemLifoCopy(ticP->memlifoP, p, n);
        algidP->Parameters.cbData = n;
    }
    return TCL_OK;
}


/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseCERT_PUBLIC_KEY_INFO(
    TwapiInterpContext *ticP,
    Tcl_Obj *pkObj,
    CERT_PUBLIC_KEY_INFO *pkP
    )
{
    Tcl_Obj **objs;
    int       nobjs;

    if (ObjGetElements(NULL, pkObj, &nobjs, &objs) != TCL_OK || nobjs != 2) {
        ObjSetStaticResult(ticP->interp,
                             "Invalid CERT_PUBLIC_KEY_INFO structure");
        return TCL_ERROR;
    }
    
    if (ParseCRYPT_ALGORITHM_IDENTIFIER(ticP, objs[0], &pkP->Algorithm) != TCL_OK ||
        ParseCRYPT_BIT_BLOB(ticP, objs[1], &pkP->PublicKey) != TCL_OK)
        return TCL_ERROR;
    
    return TCL_OK;
}

/* Returns CRYPT_KEY_PROV_INFO structure using memory from ticP->memlifo.
 * Caller responsible for releasing storage in both success and error cases
 * *kiPP is returned as NULL in case of empty list structure.
 */
static TCL_RESULT ParseCRYPT_KEY_PROV_INFO(
    TwapiInterpContext *ticP,
    Tcl_Obj *kiObj,
    CRYPT_KEY_PROV_INFO **kiPP
    )
{
    Tcl_Obj **objs;
    int i, nobjs;
    CRYPT_KEY_PROV_INFO *kiP;
    Tcl_Obj *provparaObj;

    if (ObjGetElements(NULL, kiObj, &nobjs, &objs) != TCL_OK)
        goto error_return;
    
    if (nobjs == 0) {
        *kiPP = NULL;
        return TCL_OK;
    }

    kiP = MemLifoAlloc(ticP->memlifoP, sizeof(*kiP), NULL);
    if (TwapiGetArgsEx(ticP, nobjs, objs,
                       GETWSTR(kiP->pwszContainerName),
                       GETWSTR(kiP->pwszProvName),
                       GETINT(kiP->dwProvType),
                       GETINT(kiP->dwFlags),
                       GETOBJ(provparaObj),
                       GETINT(kiP->dwKeySpec),
                       ARGEND) != TCL_OK)
        goto error_return;

    if (ObjGetElements(NULL, provparaObj, &nobjs, &objs) != TCL_OK)
        goto error_return;
    if (nobjs == 0) {
        kiP->rgProvParam = NULL;
        kiP->cProvParam = 0;
    } else {
        kiP->cProvParam = nobjs;
        kiP->rgProvParam = MemLifoAlloc(ticP->memlifoP,
                                        kiP->cProvParam * sizeof(*kiP->rgProvParam),
                                        NULL);
        for (i = 0; i < nobjs; ++i) {
            CRYPT_KEY_PROV_PARAM *parP = &kiP->rgProvParam[i];
            if (TwapiGetArgsExObj(ticP, objs[i],
                               GETINT(parP->dwParam),
                               GETBA(parP->pbData, parP->cbData),
                               GETINT(parP->dwFlags),
                               ARGEND) != TCL_OK)
                goto error_return;
        }
    }
    *kiPP = kiP;
    return TCL_OK;

error_return:
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid key provider info structure.");

}


static Tcl_Obj *ObjFromCERT_POLICY_INFO(CERT_POLICY_INFO *cpiP)
{
    DWORD i;
    Tcl_Obj *objs[2];

    objs[0] = ObjFromString(cpiP->pszPolicyIdentifier);
    objs[1] = ObjNewList(cpiP->cPolicyQualifier, NULL);
    for (i = 0; i < cpiP->cPolicyQualifier; ++i) {
        Tcl_Obj *qualObjs[2];
        CERT_POLICY_QUALIFIER_INFO *qualP = &cpiP->rgPolicyQualifier[i];
        qualObjs[0] = ObjFromString(qualP->pszPolicyQualifierId);
        qualObjs[1] = ObjFromCRYPT_BLOB(&qualP->Qualifier);
        // TBD - should we decode qualifers defined in RFC 5280 ?
        ObjAppendElement(NULL, objs[1], ObjNewList(2, qualObjs));
    }
    return ObjNewList(2, objs);
}

static Tcl_Obj *ObjFromCERT_POLICIES_INFO(CERT_POLICIES_INFO *cpiP)
{
    DWORD i;
    Tcl_Obj *objP = ObjNewList(cpiP->cPolicyInfo, NULL);
    for (i = 0; i < cpiP->cPolicyInfo; ++i) {
        ObjAppendElement(NULL, objP,
                         ObjFromCERT_POLICY_INFO(&cpiP->rgPolicyInfo[i]));
    }
    return objP;
}

static Tcl_Obj *ObjFromCERT_POLICY_MAPPINGS_INFO(CERT_POLICY_MAPPINGS_INFO *cpmiP) {
    DWORD i;
    Tcl_Obj *objP = ObjNewList(cpmiP->cPolicyMapping, NULL);
    for (i = 0; i < cpmiP->cPolicyMapping; ++i) {
        Tcl_Obj *objs[2];
        objs[0] = ObjFromString(cpmiP->rgPolicyMapping[i].pszIssuerDomainPolicy);
        objs[1] = ObjFromString(cpmiP->rgPolicyMapping[i].pszSubjectDomainPolicy);
        ObjAppendElement(NULL, objP, ObjNewList(2, objs));
    }
    return objP;
}

static Tcl_Obj *ObjFromCRL_DIST_POINTS_INFO(CRL_DIST_POINTS_INFO *cdpiP)
{
    DWORD i;
    int j;
    Tcl_Obj *objP = ObjNewList(cdpiP->cDistPoint, NULL);
    for (i = 0; i < cdpiP->cDistPoint; ++i) {
        Tcl_Obj *objs[3];
        CRL_DIST_POINT *cdpP = &cdpiP->rgDistPoint[i];
        objs[0] = ObjFromInt(cdpP->DistPointName.dwDistPointNameChoice);
        j = 1;
        switch (cdpP->DistPointName.dwDistPointNameChoice) {
        case CRL_DIST_POINT_FULL_NAME:
            objs[1] = ObjFromCERT_ALT_NAME_INFO(&cdpP->DistPointName.FullName);
            ++j;
            break;
        case CRL_DIST_POINT_NO_NAME:
        case CRL_DIST_POINT_ISSUER_RDN_NAME:
        default:
            break;
        }
        objs[0] = ObjNewList(j, objs);
        objs[1] = ObjFromCRYPT_BIT_BLOB(&cdpP->ReasonFlags);
        objs[2] = ObjFromCERT_ALT_NAME_INFO(&cdpP->CRLIssuer);
        ObjAppendElement(NULL, objP, ObjNewList(3, objs));
    }
    return objP;
}

static Tcl_Obj *ObjFromCERT_AUTHORITY_INFO_ACCESS(CERT_AUTHORITY_INFO_ACCESS *caiP)
{
    DWORD i;
    Tcl_Obj *objP = ObjNewList(caiP->cAccDescr, NULL);
    for (i = 0; i < caiP->cAccDescr; ++i) {
        Tcl_Obj *objs[2];
        objs[0] = ObjFromString(caiP->rgAccDescr[i].pszAccessMethod);
        objs[1] = ObjFromCERT_ALT_NAME_ENTRY(&caiP->rgAccDescr[i].AccessLocation);
        ObjAppendElement(NULL, objP, ObjNewList(2, objs));
    }
    return objP;
}


static TCL_RESULT TwapiCryptDecodeObject(
    Tcl_Interp *interp,
    void *poid, /* Either a Tcl_Obj or a #define X509* int value */
    void *penc,
    DWORD nenc,
    Tcl_Obj **objPP)
{
    Tcl_Obj *objP;
    Tcl_Obj *objs[3];
    union {
        void *pv;
        CERT_ENHKEY_USAGE *enhkeyP;
        CERT_AUTHORITY_KEY_ID2_INFO *akeyidP;
        CERT_BASIC_CONSTRAINTS2_INFO *basicP;
        CERT_EXTENSIONS *cextsP;
    } u;
    DWORD n;
    LPCSTR oid;
    DWORD_PTR dwoid;
    Tcl_Obj * (*fnP)(void *);

    /*
     * poid may be a Tcl_Obj or a dword corresponding to a Win32 #define
     * This is how the CryptDecodeObjEx API works
     */
    if ((DWORD_PTR) poid <= 65535) {
        dwoid = (DWORD_PTR) poid;
        oid = poid;
    } else {
        /* It's a Tcl_Obj */
        Tcl_Obj *oidObj = poid;
        if (ObjToDWORD(NULL, oidObj, &n) == TCL_OK && n < 65536) {
            dwoid = (DWORD_PTR) n;
            oid = (LPSTR) (DWORD_PTR) n;
        } else {
            oid = ObjToString(oidObj);
            if (STREQ(oid, szOID_ENHANCED_KEY_USAGE))
                dwoid = (DWORD_PTR) X509_ENHANCED_KEY_USAGE;
            else if (STREQ(oid, szOID_KEY_USAGE))
                dwoid = (DWORD_PTR) X509_KEY_USAGE;
            else if (STREQ(oid, szOID_SUBJECT_ALT_NAME2) ||
                     STREQ(oid, szOID_ISSUER_ALT_NAME2) ||
                     STREQ(oid, szOID_SUBJECT_ALT_NAME) ||
                     STREQ(oid, szOID_ISSUER_ALT_NAME))
                dwoid = (DWORD_PTR) X509_ALTERNATE_NAME;
            else if (STREQ(oid, szOID_BASIC_CONSTRAINTS2))
                dwoid = (DWORD_PTR) X509_BASIC_CONSTRAINTS2;
            else if (STREQ(oid, szOID_AUTHORITY_KEY_IDENTIFIER2))
                dwoid = (DWORD_PTR) X509_AUTHORITY_KEY_ID2;
            else if (STREQ(oid, szOID_CERT_POLICIES))
                dwoid = (DWORD_PTR) X509_CERT_POLICIES;
            else if (STREQ(oid, szOID_POLICY_CONSTRAINTS))
                dwoid = (DWORD_PTR) X509_POLICY_CONSTRAINTS;
            else if (STREQ(oid, szOID_POLICY_MAPPINGS))
                dwoid = (DWORD_PTR) X509_POLICY_MAPPINGS;
            else if (STREQ(oid, szOID_CRL_DIST_POINTS))
                dwoid = (DWORD_PTR) X509_CRL_DIST_POINTS;
            else if (STREQ(oid, szOID_AUTHORITY_INFO_ACCESS))
                dwoid = (DWORD_PTR) X509_AUTHORITY_INFO_ACCESS;
            else if (STREQ(oid, szOID_SUBJECT_INFO_ACCESS))
                dwoid = (DWORD_PTR) X509_AUTHORITY_INFO_ACCESS; // same as AUTHORITY
            else if (STREQ(oid, szOID_SUBJECT_KEY_IDENTIFIER))
                dwoid = 65535-1;
            else if (STREQ(oid, szOID_CERT_EXTENSIONS) ||
                     STREQ(oid, szOID_RSA_certExtensions))
                dwoid = (DWORD_PTR) X509_EXTENSIONS;
            else
                dwoid = 65535;      /* Will return as a byte array */
        }
    }

    if (! CryptDecodeObjectEx(
            X509_ASN_ENCODING|PKCS_7_ASN_ENCODING,
            oid, penc, nenc,
            CRYPT_DECODE_ALLOC_FLAG | CRYPT_DECODE_NOCOPY_FLAG | CRYPT_DECODE_SHARE_OID_STRING_FLAG,
            NULL,
            &u.pv,
            &n))
        return TwapiReturnSystemError(interp);
    
    objP = NULL;
    switch (dwoid) {
    case (DWORD_PTR) X509_KEY_USAGE:
        fnP = ObjFromCRYPT_BIT_BLOB;
        break;
    case (DWORD_PTR) X509_ENHANCED_KEY_USAGE:
        objP = ObjFromArgvA(u.enhkeyP->cUsageIdentifier,
                            u.enhkeyP->rgpszUsageIdentifier);
        break;
    case (DWORD_PTR) X509_ALTERNATE_NAME:
        fnP = ObjFromCERT_ALT_NAME_INFO;
        break;
    case (DWORD_PTR) X509_BASIC_CONSTRAINTS2:
        objs[0] = ObjFromBoolean(u.basicP->fCA);
        objs[1] = ObjFromBoolean(u.basicP->fPathLenConstraint);
        objs[2] = ObjFromDWORD(u.basicP->dwPathLenConstraint);
        objP = ObjNewList(3, objs);
        break;
    case (DWORD_PTR) X509_AUTHORITY_KEY_ID2:
        objs[0] = ObjFromCRYPT_BLOB(&u.akeyidP->KeyId);
        objs[1] = ObjFromCERT_ALT_NAME_INFO(&u.akeyidP->AuthorityCertIssuer);
        objs[2] = ObjFromCRYPT_BLOB(&u.akeyidP->AuthorityCertSerialNumber);
        objP = ObjNewList(3, objs);
        break;
    case X509_ALGORITHM_IDENTIFIER:
        fnP = ObjFromCRYPT_ALGORITHM_IDENTIFIER;
        break;
    case X509_CERT_REQUEST_TO_BE_SIGNED:
        fnP = ObjFromCERT_REQUEST_INFO;
        break;
    case X509_CERT_POLICIES:
        fnP = ObjFromCERT_POLICIES_INFO;
        break;
    case X509_POLICY_CONSTRAINTS:
        fnP = ObjFromCERT_POLICY_CONSTRAINTS_INFO;
        break;
    case X509_POLICY_MAPPINGS:
        fnP = ObjFromCERT_POLICY_MAPPINGS_INFO;
        break;
    case X509_EXTENSIONS:
        objP = ObjFromCERT_EXTENSIONS(u.cextsP->cExtension, u.cextsP->rgExtension);
        break;
    case X509_CRL_DIST_POINTS:
        fnP = ObjFromCRL_DIST_POINTS_INFO;
        break;
    case X509_AUTHORITY_INFO_ACCESS:
        fnP = ObjFromCERT_AUTHORITY_INFO_ACCESS;
        break;
    case X509_UNICODE_ANY_STRING:
        fnP = ObjFromCERT_NAME_VALUE_Unicode;
        break;
    case 65535-1: // szOID_SUBJECT_KEY_IDENTIFIER
        fnP = ObjFromCRYPT_BLOB;
        break;
    default:
        objP = ObjFromByteArray(u.pv, n);
        break;
    }

    if (objP == NULL)
        objP = fnP(u.pv);

    LocalFree(u.pv);
    *objPP = objP;
    return TCL_OK;
}
    
/*
 * Note: Allocates memory for blobP from ticP lifo. Note structure internal
 * pointers may point to Tcl_Obj areas within valObj so
 *  TREAT RETURNED STRUCTURES AS VOLATILE.
 *
 * We use MemLifo instead of letting CryptEncodeObjectEx do its own
 * memory allocation because it greatly simplifies freeing memory in
 * caller when multiple allocations are made.
 */
static TCL_RESULT TwapiCryptEncodeObject(
    TwapiInterpContext *ticP,
    void *poid, /* Either a Tcl_Obj or a #define X509* int value */
    Tcl_Obj *valObj,
    CRYPT_OBJID_BLOB *blobP)
{
    DWORD     dw;
    TCL_RESULT res;
    void     *penc;
    int       nenc;
    union {
        CRYPT_DATA_BLOB blob;
        CERT_ALT_NAME_INFO cani;
        CERT_ENHKEY_USAGE  ceku;
        CRYPT_BIT_BLOB     bitblob;
        CERT_BASIC_CONSTRAINTS2_INFO basic;
        CERT_AUTHORITY_KEY_ID2_INFO auth_key_id;
        CRYPT_ALGORITHM_IDENTIFIER algid;
        CERT_EXTENSIONS cexts;
        CERT_NAME_VALUE cnv;
    } u;
    Tcl_Interp *interp = ticP->interp;
    Tcl_Obj **objs;
    int       nobjs;
    LPCSTR oid;
    DWORD_PTR dwoid;

    /*
     * poid may be a Tcl_Obj or a dword corresponding to a Win32 #define
     * This is how the CryptEncodeObjEx API works
     */
    if ((DWORD_PTR) poid <= 65535) {
        dwoid = (DWORD_PTR) poid;
        oid = poid;
    } else {
        /* It's a Tcl_Obj */
        Tcl_Obj *oidObj = poid;
        if (ObjToDWORD(NULL, oidObj, &dw) == TCL_OK && dw < 65536) {
            dwoid = (DWORD_PTR) dw;
            oid = (LPSTR) (DWORD_PTR) dw;
        } else {
            oid = ObjToString(oidObj);
            if (STREQ(oid, szOID_ENHANCED_KEY_USAGE))
                dwoid = (DWORD_PTR) X509_ENHANCED_KEY_USAGE;
            else if (STREQ(oid, szOID_KEY_USAGE))
                dwoid = (DWORD_PTR) X509_KEY_USAGE;
            else if (STREQ(oid, szOID_SUBJECT_ALT_NAME2) ||
                     STREQ(oid, szOID_ISSUER_ALT_NAME2) ||
                     STREQ(oid, szOID_SUBJECT_ALT_NAME) ||
                     STREQ(oid, szOID_ISSUER_ALT_NAME))
                dwoid = (DWORD_PTR) X509_ALTERNATE_NAME;
            else if (STREQ(oid, szOID_BASIC_CONSTRAINTS2))
                dwoid = (DWORD_PTR) X509_BASIC_CONSTRAINTS2;
            else if (STREQ(oid, szOID_AUTHORITY_KEY_IDENTIFIER2))
                dwoid = (DWORD_PTR) X509_AUTHORITY_KEY_ID2;
            else if (STREQ(oid, szOID_SUBJECT_KEY_IDENTIFIER))
                dwoid = 65535-1;
            else if (STREQ(oid, szOID_CERT_EXTENSIONS) ||
                     STREQ(oid, szOID_RSA_certExtensions))
                dwoid = (DWORD_PTR) X509_EXTENSIONS;
            else
                dwoid = 65535;      /* Will return as a byte array */
        }
    }

    switch (dwoid) {
    case (DWORD_PTR) X509_KEY_USAGE:
        if ((res = ParseCRYPT_BIT_BLOB(ticP, valObj, &u.bitblob)) != TCL_OK)
            return res;
        break;
    case (DWORD_PTR) X509_ENHANCED_KEY_USAGE:
        if ((res = ParseCERT_ENHKEY_USAGE(ticP, valObj, &u.ceku)) != TCL_OK)
            return res;
        break;
    case (DWORD_PTR) X509_ALTERNATE_NAME:
        if ((res = ParseCERT_ALT_NAME_INFO(ticP, valObj, &u.cani)) != TCL_OK)
            return res;
        break;
    case (DWORD_PTR) X509_BASIC_CONSTRAINTS2:
        if (ObjGetElements(NULL, valObj, &nobjs, &objs) != TCL_OK ||
            nobjs != 3 ||
            ObjToBoolean(NULL, objs[0], &u.basic.fCA) != TCL_OK ||
            ObjToBoolean(NULL, objs[1], &u.basic.fPathLenConstraint) != TCL_OK ||
            ObjToBoolean(NULL, objs[2], &u.basic.dwPathLenConstraint) != TCL_OK) {
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid basic constraints.");
        }
        break;
    case (DWORD_PTR) X509_AUTHORITY_KEY_ID2:
        if (ObjGetElements(NULL, valObj, &nobjs, &objs) != TCL_OK ||
            nobjs != 3 ||
            ParseCRYPT_BLOB(ticP, objs[0], &u.auth_key_id.KeyId) != TCL_OK ||
            ParseCERT_ALT_NAME_INFO(ticP, objs[1], &u.auth_key_id.AuthorityCertIssuer) != TCL_OK ||
            ParseCRYPT_BLOB(ticP, objs[2], &u.auth_key_id.AuthorityCertSerialNumber) != TCL_OK) {
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid authority key id.");
        }
        break;
    case X509_ALGORITHM_IDENTIFIER:
        res = ParseCRYPT_ALGORITHM_IDENTIFIER(ticP, valObj, &u.algid);
        if (res != TCL_OK)
            return res;
        break;
    case X509_EXTENSIONS:
        res = ParseCERT_EXTENSIONS(ticP, valObj, &u.cexts.cExtension, &u.cexts.rgExtension);
        break;
    case (DWORD_PTR) X509_UNICODE_ANY_STRING:
        if ((res = ParseCERT_NAME_VALUE_Unicode(ticP, valObj, &u.cnv)) != TCL_OK)
            return res;
        break;

    case 65535-1: // szOID_SUBJECT_KEY_IDENTIFIER
        res = ParseCRYPT_BLOB(ticP, valObj, &u.blob);
        if (res != TCL_OK)
            return res;
        break;
        
    default:
        return TwapiReturnErrorMsg(interp, TWAPI_UNSUPPORTED_TYPE, "Unsupported OID");
    }

    /* Assume 1000 bytes enough but get as much as we can */
    penc = MemLifoAlloc(ticP->memlifoP, 1000, &nenc);
    if (CryptEncodeObjectEx(PKCS_7_ASN_ENCODING|X509_ASN_ENCODING,
                            oid, &u, 
                            0, NULL, penc, &nenc) == 0) {
        if (GetLastError() != ERROR_MORE_DATA)
            return TwapiReturnSystemError(interp);
        /* Retry with specified buffer size */
        penc = MemLifoAlloc(ticP->memlifoP, nenc, &nenc);
        if (CryptEncodeObjectEx(PKCS_7_ASN_ENCODING|X509_ASN_ENCODING,
                                oid, &u, 0, NULL, penc, &nenc) == 0)
            return TwapiReturnSystemError(interp);
    }
    
    blobP->cbData = nenc;
    blobP->pbData = penc;

    /* Note caller has to MemLifoPop* to release lifo memory */
    return TCL_OK;
}


static Tcl_Obj *ObjFromCERT_EXTENSION(CERT_EXTENSION *extP)
{
    Tcl_Obj *objs[3];
    Tcl_Obj *extObj;

    objs[0] = ObjFromString(extP->pszObjId);
    objs[1] = ObjFromInt(extP->fCritical);
    extObj = ObjFromString(extP->pszObjId);
    if (TwapiCryptDecodeObject(NULL, extObj,
                               extP->Value.pbData,
                               extP->Value.cbData, &objs[2]) != TCL_OK) {
        /* TBD - this is not quite correct. How can caller distinguish error?*/
        objs[2] = ObjFromByteArray(extP->Value.pbData, extP->Value.cbData);
    }
    ObjDecrRefs(extObj);
    return ObjNewList(3, objs);
}

static Tcl_Obj *ObjFromCERT_EXTENSIONS(int nexts, CERT_EXTENSION *extP)
{
    int i;
    Tcl_Obj *objP = ObjNewList(nexts, NULL);
    for (i = 0; i < nexts; ++i) {
        ObjAppendElement(NULL, objP, ObjFromCERT_EXTENSION(i + extP));
    }
    return objP;
}

/* Returns pointer to a CERT_EXTENSIONS_IDENTIFIER structure in *extsPP
   using memory from ticP->memlifo. Caller responsible for storage in both
   success and error cases.
   Can return NULL in *extsPP if extObj is empty list.
*/
static TCL_RESULT ParseCERT_EXTENSION(
    TwapiInterpContext *ticP,
    Tcl_Obj *extObj,
    CERT_EXTENSION *extP
    )
{
    Tcl_Obj **objs;
    int       nobjs, n;
    TCL_RESULT res;
    void      *pv;
    BOOL       bval;

    if ((res = ObjGetElements(ticP->interp, extObj, &nobjs, &objs)) != TCL_OK)
        return res;
    if (nobjs != 2 && nobjs != 3) {
        ObjSetStaticResult(ticP->interp, "Certificate extension format invalid or not implemented");
        return TCL_ERROR;
    }
    if ((res = ObjToBoolean(ticP->interp, objs[1], &bval)) != TCL_OK)
        return res;

    pv = ObjToStringN(objs[0], &n);
    extP->pszObjId = MemLifoCopy(ticP->memlifoP, pv, n+1);
    extP->fCritical = (BOOL) bval;
    if (nobjs == 3) {
        res = TwapiCryptEncodeObject(ticP,
                                     objs[0], objs[2],
                                     &extP->Value);
        if (res != TCL_OK)
            return res;
    } else {
        extP->Value.cbData = 0;
        extP->Value.pbData = NULL;
    }

    return TCL_OK;
}

/* Returns pointer to a CERT_EXTENSIONS_IDENTIFIER structure in *extsPP
   using memory from ticP->memlifo. Caller responsible for storage in both
   success and error cases.
   Can return NULL in *extsPP if extObj is empty list.
*/
static TCL_RESULT ParseCERT_EXTENSIONS(
    TwapiInterpContext *ticP,
    Tcl_Obj *extsObj,
    DWORD *nextsP,
    CERT_EXTENSION **extsPP
    )
{
    CERT_EXTENSION *extsP;
    Tcl_Obj **objs;
    int       i, nobjs;
    TCL_RESULT res;
    Tcl_Interp *interp = ticP->interp;

    if ((res = ObjGetElements(interp, extsObj, &nobjs, &objs)) != TCL_OK)
        return res;

    if (nobjs == 0) {
        *extsPP = NULL;
        *nextsP = 0;
        return TCL_OK;
    }

    extsP = MemLifoAlloc(ticP->memlifoP, nobjs * sizeof(CERT_EXTENSION), NULL);
    for (i = 0; i < nobjs; ++i) {
        if ((res = ParseCERT_EXTENSION(ticP, objs[i], &extsP[i])) != TCL_OK)
            return res;
    }
    *extsPP = extsP;
    *nextsP = nobjs;
    return TCL_OK;
}

/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseCERT_INFO(
    TwapiInterpContext *ticP,
    Tcl_Obj *ciObj,
    CERT_INFO *ciP             /* Will contain garbage in case of errors */
    )
{
    Tcl_Interp *interp = ticP->interp;
    Tcl_Obj *algObj, *pubkeyObj, *issuerIdObj, *subjectIdObj, *extsObj;

    if (TwapiGetArgsExObj(ticP, ciObj,
                       GETINT(ciP->dwVersion),
                       GETBA(ciP->SerialNumber.pbData, ciP->SerialNumber.cbData),
                       GETOBJ(algObj),
                       GETBA(ciP->Issuer.pbData, ciP->Issuer.cbData),
                       GETVAR(ciP->NotBefore, ObjToFILETIME),
                       GETVAR(ciP->NotAfter, ObjToFILETIME),
                       GETBA(ciP->Subject.pbData, ciP->Subject.cbData),
                       GETOBJ(pubkeyObj),
                       GETOBJ(issuerIdObj),
                       GETOBJ(subjectIdObj),
                       GETOBJ(extsObj), ARGEND) != TCL_OK) {
        Tcl_AppendResult(interp, "Invalid CERT_INFO structure", NULL);
        return TCL_ERROR;
    }

    if (ParseCRYPT_ALGORITHM_IDENTIFIER(ticP, algObj, &ciP->SignatureAlgorithm) != TCL_OK ||
        ParseCERT_PUBLIC_KEY_INFO(ticP, pubkeyObj, &ciP->SubjectPublicKeyInfo) != TCL_OK ||
        ParseCRYPT_BIT_BLOB(ticP, issuerIdObj, &ciP->IssuerUniqueId) != TCL_OK ||
        ParseCRYPT_BIT_BLOB(ticP, subjectIdObj, &ciP->SubjectUniqueId) != TCL_OK ||
        ParseCERT_EXTENSIONS(ticP, extsObj, &ciP->cExtension, &ciP->rgExtension) != TCL_OK)
        return TCL_ERROR;
        
    return TCL_OK;
}

/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseCRYPT_ATTRIBUTE(
    TwapiInterpContext *ticP,
    Tcl_Obj *attrObj,
    CRYPT_ATTRIBUTE *attrP /* Will contain garbage in case of errors */
    )
{
    Tcl_Obj **objs;
    int       n, nobjs;
    Tcl_Interp *interp = ticP->interp;
    void *pv;

    if (ObjGetElements(NULL, attrObj, &nobjs, &objs) == TCL_OK &&
        nobjs == 2) {
        Tcl_Obj **valObjs;
        pv = ObjToStringN(objs[0], &n);
        attrP->pszObjId = MemLifoCopy(ticP->memlifoP, pv, n+1);
        if (ObjGetElements(NULL, objs[1], &nobjs, &valObjs) == TCL_OK) {
            attrP->cValue = nobjs;
            attrP->rgValue = MemLifoAlloc(ticP->memlifoP, nobjs*sizeof(*(attrP->rgValue)), NULL);
            for (n = 0; n < nobjs; ++n) {
                if (TwapiCryptEncodeObject(ticP, objs[0], valObjs[n], &attrP->rgValue[n]) != TCL_OK)
                    goto error_return;
            }
            return TCL_OK;
        }
    }

error_return:
    Tcl_AppendResult(interp, "Invalid CRYPT_ATTRIBUTE structure", NULL);
    return TCL_ERROR;
}

/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseCERT_REQUEST_INFO(
    TwapiInterpContext *ticP,
    Tcl_Obj *criObj,
    CERT_REQUEST_INFO *criP /* Will contain garbage in case of errors */
    )
{
    Tcl_Obj **objs;
    int       nobjs;
    Tcl_Interp *interp = ticP->interp;
    Tcl_Obj *pubkeyObj, *attrObj;

    if (TwapiGetArgsExObj(ticP, criObj,
                          GETINT(criP->dwVersion),
                          GETBA(criP->Subject.pbData, criP->Subject.cbData),
                          GETOBJ(pubkeyObj), GETOBJ(attrObj),
                          ARGEND) == TCL_OK &&
        ParseCERT_PUBLIC_KEY_INFO(ticP, pubkeyObj, &criP->SubjectPublicKeyInfo) == TCL_OK &&
        ObjGetElements(NULL, attrObj, &nobjs, &objs) == TCL_OK) {
        if (nobjs == 0) {
            criP->cAttribute = 0;
            criP->rgAttribute = NULL;
        } else {
            int i;
            criP->cAttribute = nobjs;
            criP->rgAttribute = MemLifoAlloc(ticP->memlifoP, nobjs * sizeof(*(criP->rgAttribute)), NULL);
            for (i = 0; i < nobjs; ++i) {
                if (ParseCRYPT_ATTRIBUTE(ticP, objs[i], &criP->rgAttribute[i]) != TCL_OK)
                    goto error_return;
            }
        }
        return TCL_OK;
    }


error_return:
    Tcl_AppendResult(interp, "Invalid CERT_REQUEST_INFO structure", NULL);
    return TCL_ERROR;
}

/* 
 * Parses a non-empty Tcl_Obj into a SYSTEMTIME structure *timeP 
 * and stores timeP in *timePP. If the Tcl_Obj is empty (meaning use default)
 * stores NULL in *timePP (and still return TCL_OK)
 */
static TCL_RESULT ParseSYSTEMTIME(
    Tcl_Interp *interp,
    Tcl_Obj *timeObj,
    SYSTEMTIME *timeP,
    SYSTEMTIME **timePP
    )
{
    Tcl_Obj **objs;
    int       nobjs;
    TCL_RESULT res;

    if ((res = ObjGetElements(interp, timeObj, &nobjs, &objs)) != TCL_OK)
        return res;
    if (nobjs == 0)
        *timePP = NULL;
    else {
        if ((res = ObjToSYSTEMTIME(interp, timeObj, timeP)) != TCL_OK)
            return res;
        *timePP = timeP;
    }
    return TCL_OK;
}


static int Twapi_CertCreateSelfSignCertificate(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    void *pv;
    HCRYPTPROV hprov;
    DWORD flags;
    int status;
    CERT_NAME_BLOB name_blob;
    CRYPT_KEY_PROV_INFO *kiP;
    CRYPT_ALGORITHM_IDENTIFIER algid, *algidP;
    int       nobjs;
    SYSTEMTIME start, end, *startP, *endP;
    PCERT_CONTEXT certP;
    CERT_EXTENSIONS exts;
    Tcl_Obj *algidObj, *startObj, *endObj, *extsObj, *provinfoObj;
    MemLifoMarkHandle mark;

    mark = MemLifoPushMark(ticP->memlifoP);

    if ((status = TwapiGetArgsEx(ticP, objc-1, objv+1,
                                 GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                                 GETBA(name_blob.pbData, name_blob.cbData),
                                 GETINT(flags),
                                 GETOBJ(provinfoObj),
                                 GETOBJ(algidObj),
                                 GETOBJ(startObj),
                                 GETOBJ(endObj),
                                 GETOBJ(extsObj),
                                 ARGEND)) != TCL_OK)
        goto vamoose;
    
    hprov = (HCRYPTPROV) pv;
 
    status = ParseCRYPT_KEY_PROV_INFO(ticP, provinfoObj, &kiP);
    if (status != TCL_OK)
        goto vamoose;

    /* Parse CRYPT_ALGORITHM_IDENTIFIER */
    if ((status = ObjListLength(interp, algidObj, &nobjs)) != TCL_OK)
        goto vamoose;
    if (nobjs == 0)
        algidP = NULL;
    else {
        algidP = &algid;
        status = ParseCRYPT_ALGORITHM_IDENTIFIER(ticP, algidObj, algidP);
        if (status != TCL_OK)
            goto vamoose;
    }

    if ((status = ParseSYSTEMTIME(interp, startObj, &start, &startP)) != TCL_OK)
        goto vamoose;
    if ((status = ParseSYSTEMTIME(interp, endObj, &end, &endP)) != TCL_OK)
        goto vamoose;

    if ((status = ParseCERT_EXTENSIONS(ticP, extsObj,
                                       &exts.cExtension, &exts.rgExtension))
        != TCL_OK)
        goto vamoose;

    certP = (PCERT_CONTEXT) CertCreateSelfSignCertificate(
        hprov, &name_blob, flags, kiP, algidP, startP, endP,
        exts.rgExtension ? &exts : NULL);

    if (certP) {
        TwapiRegisterPCCERT_CONTEXTTic(ticP, certP);
        ObjSetResult(interp, ObjFromOpaque(certP, "PCCERT_CONTEXT"));
        status = TCL_OK;
    } else {
        status = TwapiReturnSystemError(interp);
    }

vamoose:
    MemLifoPopMark(mark);
    return status;
}

static int Twapi_CertGetCertificateContextProperty(Tcl_Interp *interp, PCCERT_CONTEXT certP, DWORD prop_id, int cooked)
{
    DWORD n = 0;
    TwapiResult result;
    void *pv;
    CERT_KEY_CONTEXT ckctx;
    char *s;
    DWORD_PTR dwp;
    TCL_RESULT res;

    result.type = TRT_BADFUNCTIONCODE;
    if (cooked) {
        switch (prop_id) {
        case CERT_ACCESS_STATE_PROP_ID:
        case CERT_KEY_SPEC_PROP_ID:
            result.type = TRT_DWORD; 
            n = sizeof(result.value.ival);
            result.type = CertGetCertificateContextProperty(certP, prop_id, &result.value.uval, &n) ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case CERT_DATE_STAMP_PROP_ID:
            n = sizeof(result.value.filetime);
            result.type = CertGetCertificateContextProperty(certP, prop_id,
                                                            &result.value.filetime, &n)
                ? TRT_FILETIME : TRT_GETLASTERROR;
            break;
        case CERT_ARCHIVED_PROP_ID:
            result.type = TRT_BOOL;
            if (! CertGetCertificateContextProperty(certP, prop_id, NULL, &n)) {
                if ((result.value.ival = GetLastError()) == CRYPT_E_NOT_FOUND)
                    result.value.bval = 0;
                else
                    result.type = TRT_EXCEPTION_ON_ERROR;
            } else
                result.value.bval = 1;
            break;

        case CERT_ENHKEY_USAGE_PROP_ID:
            if (! CertGetCertificateContextProperty(certP, prop_id, NULL, &n))
                return TwapiReturnSystemError(interp);
            pv = TwapiAlloc(n);
            if (! CertGetCertificateContextProperty(certP, prop_id, pv, &n)) {
                TwapiFree(pv);
                return TwapiReturnSystemError(interp);
            }        
            res = TwapiCryptDecodeObject(interp, (void*)X509_ENHANCED_KEY_USAGE, pv, n, &result.value.obj);
            TwapiFree(pv);
            if (res != TCL_OK)
                return res;
            result.type = TRT_OBJ;
            break;

        case CERT_KEY_PROV_INFO_PROP_ID:
            if (! CertGetCertificateContextProperty(certP, prop_id, NULL, &n))
                return TwapiReturnSystemError(interp);
            pv = TwapiAlloc(n);
            if (! CertGetCertificateContextProperty(certP, prop_id, pv, &n)) {
                TwapiFree(pv);
                return TwapiReturnSystemError(interp);
            }        
            result.value.obj = ObjFromCRYPT_KEY_PROV_INFO(pv);
            TwapiFree(pv);
            result.type = TRT_OBJ;
            break;

        case CERT_KEY_CONTEXT_PROP_ID:
            n = ckctx.cbSize = sizeof(ckctx);
            if (CertGetCertificateContextProperty(certP, prop_id, &ckctx, &n)) {
                result.value.obj = ObjNewList(0, NULL);
                if (ckctx.dwKeySpec == AT_KEYEXCHANGE ||
                    ckctx.dwKeySpec == AT_SIGNATURE) {
                    TwapiRegisterHCRYPTPROV(interp, ckctx.hCryptProv);
                    s = "HCRYPTPROV";
                } else {
                    /* TBD - do we need to register this pointer as well? */
                    s = "NCRYPT_KEY_HANDLE";
                }
                ObjAppendElement(NULL, result.value.obj, ObjFromOpaque((void*)ckctx.hCryptProv, s));
                ObjAppendElement(NULL, result.value.obj, ObjFromDWORD(ckctx.dwKeySpec));
            } else
                result.type = TRT_GETLASTERROR;
            break;
        
        case CERT_KEY_PROV_HANDLE_PROP_ID:
            n = sizeof(dwp);
            if (CertGetCertificateContextProperty(certP, prop_id, &dwp, &n)) {
                TwapiRegisterHCRYPTPROV(interp, dwp);
                TwapiResult_SET_PTR(result, HCRYPTPROV, (void*)dwp);
            } else
                result.type = TRT_GETLASTERROR;
            break;

#ifndef CERT_REQUEST_ORIGINATOR_PROP_ID
# define CERT_REQUEST_ORIGINATOR_PROP_ID 71
#endif
        case CERT_REQUEST_ORIGINATOR_PROP_ID:
        case CERT_AUTO_ENROLL_PROP_ID:
        case CERT_EXTENDED_ERROR_INFO_PROP_ID:
        case CERT_FRIENDLY_NAME_PROP_ID:
        case CERT_PVK_FILE_PROP_ID:
        case CERT_DESCRIPTION_PROP_ID:
            if (! CertGetCertificateContextProperty(certP, prop_id, NULL, &n))
                return TwapiReturnSystemError(interp);
            result.value.unicode.str = TwapiAlloc(n);
            if (CertGetCertificateContextProperty(certP, prop_id,
                                                  result.value.unicode.str, &n)) {
                result.value.unicode.len = -1;
                result.type = TRT_UNICODE_DYNAMIC; /* Will also free memory */
            } else {
                TwapiReturnSystemError(interp);
                TwapiFree(result.value.unicode.str);
                return TCL_ERROR;
            }
            break;
        }
    } 

    if (result.type == TRT_BADFUNCTIONCODE) {
        /* Either raw format wanted or binary data */

        /*        
         * The following are handled via defaults for now
         *  CERT_HASH_PROP_ID:
         *  CERT_ISSUER_PUBLIC_KEY_MD5_HASH_PROP_ID:
         *  CERT_ISSUER_SERIAL_NUMBER_MD5_HASH_PROP_ID:
         *  CERT_ARCHIVED_KEY_HASH_PROP_ID:
         *  CERT_KEY_IDENTIFIER_PROP_ID:
         *  CERT_MD5_HASH_PROP_ID
         *  CERT_RENEWAL_PROP_ID
         *  CERT_SHA1_HASH_PROP_ID
         *  CERT_SIGNATURE_HASH_PROP_ID
         *  CERT_SUBJECT_PUBLIC_KEY_MD5_HASH_PROP_ID
         */

        if (! CertGetCertificateContextProperty(certP, prop_id, NULL, &n))
            return TwapiReturnSystemError(interp);
        result.type = TRT_OBJ;
        result.value.obj = ObjFromByteArray(NULL, n);
        if (! CertGetCertificateContextProperty(
                certP, prop_id,
                ObjToByteArray(result.value.obj, &n),
                &n)) {
            TwapiReturnSystemError(interp);
            ObjDecrRefs(result.value.obj);
            return TCL_ERROR;
        }
        Tcl_SetByteArrayLength(result.value.obj, n);
    }

    return TwapiSetResult(interp, &result);
}

static TCL_RESULT TwapiCertGetNameString(
    Tcl_Interp *interp,
    PCCERT_CONTEXT certP,
    DWORD type,
    DWORD flags,
    Tcl_Obj *owhat)
{
    void *pv;
    DWORD dw, nchars;
    WCHAR buf[1024];

    switch (type) {
    case CERT_NAME_EMAIL_TYPE: // 1
    case CERT_NAME_SIMPLE_DISPLAY_TYPE: // 4
    case CERT_NAME_FRIENDLY_DISPLAY_TYPE: // 5
    case CERT_NAME_DNS_TYPE: // 6
    case CERT_NAME_URL_TYPE: // 7
    case CERT_NAME_UPN_TYPE: // 8
        pv = NULL;
        break;
    case CERT_NAME_RDN_TYPE: // 2
        if (ObjToInt(interp, owhat, &dw) != TCL_OK)
            return TCL_ERROR;
        pv = &dw;
        break;
    case CERT_NAME_ATTR_TYPE: // 3
        pv = ObjToString(owhat);
        break;
    default:
        ObjSetResult(interp, Tcl_ObjPrintf("CertGetNameString: unknown type %d", type));
        return TCL_ERROR;
    }

    // 1 -> CERT_NAME_ISSUER_FLAG 
    // 0x00010000 -> CERT_NAME_DISABLE_IE4_UTF8_FLAG 
    // are supported.
    // 2 -> CERT_NAME_SEARCH_ALL_NAMES_FLAG
    // 0x00200000 -> CERT_NAME_STR_ENABLE_PUNYCODE_FLAG 
    // are post Win8 AND they will change output encoding/format
    // Only support what we know
    if (flags & ~(0x00010001)) {
        ObjSetResult(interp, Tcl_ObjPrintf("CertGetNameString: unsupported flags %d", flags));
        return TCL_ERROR;
    }

    nchars = CertGetNameStringW(certP, type, flags, pv, buf, ARRAYSIZE(buf));
    /* Note nchars includes terminating NULL */
    if (nchars > 1) {
        if (nchars < ARRAYSIZE(buf)) {
            ObjSetResult(interp, ObjFromUnicodeN(buf, nchars-1));
        } else {
            /* Buffer might have been truncated. Explicitly get buffer size */
            WCHAR *bufP;
            nchars = CertGetNameStringW(certP, type, flags, pv, NULL, 0);
            bufP = TwapiAlloc(nchars*sizeof(WCHAR));
            nchars = CertGetNameStringW(certP, type, flags, pv, bufP, nchars);
            ObjSetResult(interp, ObjFromUnicodeN(bufP, nchars-1));
            TwapiFree(bufP);
        }
    }
    return TCL_OK;
}

static TCL_RESULT Twapi_CryptSetProvParam(Tcl_Interp *interp,
                                          HCRYPTPROV hprov, DWORD param,
                                          DWORD flags, Tcl_Obj *objP)
{
    TCL_RESULT res;
    void *pv;
    HWND hwnd;
    SECURITY_DESCRIPTOR *secdP;

    switch (param) {
    case PP_CLIENT_HWND:
        if ((res = ObjToHWND(interp, objP, &hwnd)) != TCL_OK)
            return res;
        pv = &hwnd;
        break;
    case PP_DELETEKEY:
        pv = NULL;
        break;
    case PP_KEYEXCHANGE_PIN: /* FALLTHRU */
    case PP_SIGNATURE_PIN:
        pv = ObjToString(objP);
        break;
    case PP_KEYSET_SEC_DESCR:
        if ((res = ObjToPSECURITY_DESCRIPTOR(interp, objP, &secdP)) != TCL_OK)
            return res;
        /* TBD - check what happens with NULL secdP (which is valid) */
        pv = secdP;
        break;
#ifdef PP_PIN_PROMPT_STRING
    case PP_PIN_PROMPT_STRING:
#else
    case 44:
#endif
        /* FALLTHRU */
    case PP_UI_PROMPT:
        pv = ObjToUnicode(objP);
        break;
    default:
        return TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS, Tcl_ObjPrintf("Provider parameter %d not implemented", param));
    }

    if (CryptSetProvParam(hprov, param, pv, flags)) {
        res = TCL_OK;
    } else {
        res = TwapiReturnSystemError(interp);
    }

    TwapiFreeSECURITY_DESCRIPTOR(secdP); /* OK if NULL */
    
    return res;
}


static TCL_RESULT Twapi_CryptGetProvParam(Tcl_Interp *interp,
                                          HCRYPTPROV hprov,
                                          DWORD param, DWORD flags)
{
    Tcl_Obj *objP;
    DWORD n;
    void *pv;

    n = 0;
    /* Special case PP_ENUMCONTAINERS because of how the iteration
       works. We return ALL containers as opposed to one at a time */
    if (param == PP_ENUMCONTAINERS) {
        if (! CryptGetProvParam(hprov, param, NULL, &n, CRYPT_FIRST))
            return TwapiReturnSystemError(interp);
        /* n is now the max size buffer. Subsequent calls will not change that value */
        pv = TwapiAlloc(n * sizeof(char));
        objP = Tcl_NewListObj(0, NULL);
        flags = CRYPT_FIRST;
        while (CryptGetProvParam(hprov, param, pv, &n, flags)) {
            ObjAppendElement(NULL, objP, ObjFromString(pv));
            flags = CRYPT_NEXT;
        }
        n = GetLastError();
        TwapiFree(pv);
        if (n != ERROR_NO_MORE_ITEMS) {
            ObjDecrRefs(objP);
            return Twapi_AppendSystemError(interp, n);
        }
        ObjSetResult(interp, objP);
        return TCL_OK;
    }
    
    if (! CryptGetProvParam(hprov, param, NULL, &n, flags))
        return TwapiReturnSystemError(interp);
    
    if (param == PP_KEYSET_SEC_DESCR) {
        objP = NULL;
        pv = TwapiAlloc(n);
    } else {
        objP = ObjFromByteArray(NULL, n);
        pv = ObjToByteArray(objP, &n);
    }

    if (! CryptGetProvParam(hprov, param, pv, &n, flags)) {
        if (objP)
            ObjDecrRefs(objP);
        TwapiReturnSystemError(interp);
        return TCL_ERROR;
    }

    if (param == PP_KEYSET_SEC_DESCR) {
        if (n == 0)
            objP = ObjFromEmptyString();
        else
            objP = ObjFromSECURITY_DESCRIPTOR(interp, pv);
        TwapiFree(pv);
        if (objP == NULL)
            return TCL_ERROR;   /* interp already contains error */
    } else
        Tcl_SetByteArrayLength(objP, n);

    ObjSetResult(interp, objP);
    return TCL_OK;
}

static TCL_RESULT Twapi_CertOpenStore(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD store_provider, enc_type, flags;
    void *pv = NULL;
    HCERTSTORE hstore;
    HANDLE h;
    TCL_RESULT res;
    CRYPT_DATA_BLOB blob;

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(store_provider), GETINT(enc_type), ARGUNUSED,
                     GETINT(flags), ARGSKIP, ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    /* Using literals because the #defines are cast as LPCSTR */
    switch (store_provider) {
    case 2: // CERT_STORE_PROV_MEMORY
    case 11: // CERT_STORE_PROV_COLLECTION
        break;

    case 3: // CERT_STORE_PROV_FILE
        if ((res = ObjToOpaque(interp, objv[4], &h, "HANDLE")) != TCL_OK)
            return res;
        pv = &h;
        break;

    case 4: // CERT_STORE_PROV_REG
        /* Docs imply pv itself is the handle unlike the FILE case above */
        if ((res = ObjToOpaque(interp, objv[4], &pv, "HANDLE")) != TCL_OK)
            return res;
        break;

    case 8: // CERT_STORE_PROV_FILENAME_W
    case 14: // CERT_STORE_PROV_PHYSICAL_W
    case 10: // CERT_STORE_PROV_SYSTEM_W
    case 13: // CERT_STORE_PROV_SYSTEM_REGISTRY_W
        pv = ObjToUnicode(objv[4]);
        break;

    case 5: // CERT_STORE_PROV_PKCS7
    case 6: // CERT_STORE_PROV_SERIALIZED
        blob.pbData = ObjToByteArray(objv[4], &blob.cbData);
        pv = &blob;
        break;

    case 15: // CERT_STORE_PROV_SMART_CARD
    case 16: // CERT_STORE_PROV_LDAP
    case 1: // CERT_STORE_PROV_MSG
    default:
        ObjSetResult(interp,
                         Tcl_ObjPrintf("Invalid or unsupported store provider \"%d\"", store_provider));
        return TCL_ERROR;
    }

    hstore = CertOpenStore(IntToPtr(store_provider), enc_type, 0, flags, pv);
    if (hstore) {
        /* CertCloseStore does not check pointer validity! So do ourselves*/
        TwapiRegisterHCERTSTORE(interp, hstore);
        ObjSetResult(interp, ObjFromOpaque(hstore, "HCERTSTORE"));
        return TCL_OK;
    } else {
        if (flags & CERT_STORE_DELETE_FLAG) {
            /* Return value can mean success as well */
            if (GetLastError() == 0)
                return TCL_OK;
        }
        return TwapiReturnSystemError(interp);
    }
}

static TCL_RESULT Twapi_PFXExportCertStoreExObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HCERTSTORE hstore;
    LPWSTR password = NULL;
    int password_len;
    Tcl_Obj *objP, *passObj;
    CRYPT_DATA_BLOB blob;
    BOOL status;
    int flags;
    TCL_RESULT res;
    
    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVERIFIEDPTR(hstore, HCERTSTORE, CertCloseStore),
                     GETOBJ(passObj), ARGUNUSED,
                     GETINT(flags), ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    password = ObjDecryptUnicode(interp, passObj, &password_len);
    if (password == NULL)
        return TCL_ERROR;

    res = TCL_OK;

    blob.cbData = 0;
    blob.pbData = NULL;

    status = PFXExportCertStoreEx(hstore, &blob, password, NULL, flags);
    if (!status) {
        res = TwapiReturnSystemError(interp);
        goto vamoose;
    }

    if (blob.cbData == 0)
        goto vamoose;

    objP = ObjFromByteArray(NULL, blob.cbData);
    blob.pbData = ObjToByteArray(objP, &blob.cbData);
    status = PFXExportCertStoreEx(hstore, &blob, password, NULL, flags);
    if (! status) {
        res = TwapiReturnSystemError(interp);
        ObjDecrRefs(objP);
        goto vamoose;
    }
    ObjSetResult(interp, objP);

vamoose:
    if (password) {
        SecureZeroMemory(password, sizeof(WCHAR) * password_len);
        TwapiFree(password);
    }
    return res;
}

static TCL_RESULT Twapi_PFXImportCertStoreObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HCERTSTORE hstore;
    LPWSTR password = NULL;
    int password_len;
    Tcl_Obj *passObj;
    CRYPT_DATA_BLOB blob;
    int flags;
    TCL_RESULT res;
    
    if (TwapiGetArgsEx(ticP, objc-1, objv+1,
                       GETBA(blob.pbData, blob.cbData),
                       GETOBJ(passObj),
                       GETINT(flags), ARGEND) != TCL_OK)
        return TCL_ERROR;
    
    password = ObjDecryptUnicode(interp, passObj, &password_len);
    if (password == NULL)
        return TCL_ERROR;

    res = TCL_OK;

    hstore = PFXImportCertStore(&blob, password, flags);
    if (hstore == NULL) {
        res = TwapiReturnSystemError(interp);
        goto vamoose;
    }

    TwapiRegisterHCERTSTORETic(ticP, hstore);
    ObjSetResult(interp, ObjFromOpaque(hstore, "HCERTSTORE"));

vamoose:
    if (password) {
        SecureZeroMemory(password, sizeof(WCHAR) * password_len);
        TwapiFree(password);
    }
    return res;
}


static TCL_RESULT Twapi_CertFindCertificateInStoreObjCmd(
    TwapiInterpContext *ticP, Tcl_Interp *interp, int objc,
    Tcl_Obj *CONST objv[])
{
    HCERTSTORE hstore;
    PCCERT_CONTEXT certP, cert2P;
    DWORD         enctype, flags, findtype, dw;
    Tcl_Obj      *findObj;
    void         *pv;
    CERT_BLOB     blob;
    CERT_INFO     cinfo;
    TCL_RESULT    res;
    CERT_PUBLIC_KEY_INFO pki;
    MemLifoMarkHandle mark = NULL;

    certP = NULL;
    res = TwapiGetArgs(interp, objc-1, objv+1,
                       GETVERIFIEDPTR(hstore, HCERTSTORE, CertCloseStore),
                       GETINT(enctype), GETINT(flags), GETINT(findtype),
                       GETOBJ(findObj), GETVERIFIEDORNULL(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                       ARGEND);
    if (res != TCL_OK) {
        /* We have guaranteed caller certP will be freed even on error */
        if (certP && TwapiUnregisterPCCERT_CONTEXTTic(ticP, certP) == TCL_OK)
            CertFreeCertificateContext(certP);
        return res;
    }    

    switch (findtype) {
    case CERT_FIND_ANY:
        pv = NULL;
        break;
    case CERT_FIND_EXISTING:
        res = ObjToVerifiedPointerTic(ticP, findObj, (void **)&cert2P, "PCCERT_CONTEXT", CertFreeCertificateContext);
        if (res == TCL_OK)
            pv = (void *)cert2P;
        break;
    case CERT_FIND_SUBJECT_CERT:
        pv = &cinfo;
        mark = MemLifoPushMark(ticP->memlifoP);
        res = ParseCERT_INFO(ticP, findObj, &cinfo);
        break;
    case CERT_FIND_KEY_IDENTIFIER: /* FALLTHRU */
    case CERT_FIND_MD5_HASH:    /* FALLTHRU */
    case CERT_FIND_PUBKEY_MD5_HASH:    /* FALLTHRU */
    case CERT_FIND_SHA1_HASH:   /* FALLTHRU */
    case CERT_FIND_SIGNATURE_HASH: /* FALLTHRU */
    case CERT_FIND_ISSUER_NAME: /* FALLTHRU */
    case CERT_FIND_SUBJECT_NAME:
        blob.pbData = ObjToByteArray(findObj, &blob.cbData);
        pv = &blob;
        break;
    case CERT_FIND_ISSUER_STR_W: /* FALLTHRU */
    case CERT_FIND_SUBJECT_STR_W:
        pv = ObjToUnicode(findObj);
        break;
    case CERT_FIND_PROPERTY: /* FALLTHRU */
    case CERT_FIND_KEY_SPEC:
        res = ObjToDWORD(interp, findObj, &dw);
        pv = &dw;
        break;
    case CERT_FIND_PUBLIC_KEY:
        pv = &pki;
        mark = MemLifoPushMark(ticP->memlifoP);
        res = ParseCERT_PUBLIC_KEY_INFO(ticP, findObj, &pki);
        break;
    default:
        res = TwapiReturnError(interp, TWAPI_UNSUPPORTED_TYPE);
        break;
    }

    /*
     * CertFindCertificateInStore ALWAYS releases certP (even in error case)
     * Caller expects that to happen in all cases so if we are not
     * calling CertFindCertificateInStore because of previous errors,
     * do so ourselves
     */
    if (certP) {
        /* Do not change res unless it is an error */
        if (TwapiUnregisterPCCERT_CONTEXTTic(ticP, certP) != TCL_OK)
            res = TCL_ERROR;
    }
    if (res != TCL_OK)
        CertFreeCertificateContext(certP);
    else {
        certP = CertFindCertificateInStore(
            hstore,
            X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
            0,
            findtype,
            pv,
            certP);
        if (certP) {
            TwapiRegisterPCCERT_CONTEXTTic(ticP, certP);
            ObjSetResult(interp, ObjFromOpaque((void*)certP, "PCCERT_CONTEXT"));
        } else {
            /* EOF is not an error */
            if (GetLastError() != CRYPT_E_NOT_FOUND)
                res = TwapiReturnSystemError(interp);
        }
    }

    if (mark)
        MemLifoPopMark(mark);

    return res;
}

static TCL_RESULT Twapi_CryptSignAndEncodeCertObjCmd(
    TwapiInterpContext *ticP, Tcl_Interp *interp, int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Obj *algidObj, *certinfoObj, *encodedObj;
    TCL_RESULT res;
    CRYPT_ALGORITHM_IDENTIFIER algid;
    DWORD keyspec, enctype;
    union {
        CERT_INFO ci;
        CERT_REQUEST_INFO cri;
    } u;
    HCRYPTPROV hprov;
    MemLifoMarkHandle mark;
    DWORD nbytes;
    DWORD structtype;

    mark = MemLifoPushMark(ticP->memlifoP);
    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                         GETVERIFIEDPTR(hprov, HCRYPTPROV, CryptReleaseContext),
                         GETINT(keyspec), GETINT(enctype),
                         GETINT(structtype),
                         GETOBJ(certinfoObj), GETOBJ(algidObj),
                         ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    switch (structtype) {
    case (DWORD)(DWORD_PTR) X509_CERT_TO_BE_SIGNED:
        res = ParseCERT_INFO(ticP, certinfoObj, &u.ci);
        break;
    case (DWORD)(DWORD_PTR) X509_CERT_REQUEST_TO_BE_SIGNED:
        res = ParseCERT_REQUEST_INFO(ticP, certinfoObj, &u.cri);
        break;
    default:
        res = TwapiReturnError(interp, TWAPI_UNSUPPORTED_TYPE);
        break;
    }
    if (res != TCL_OK)
        goto vamoose;

    res = ParseCRYPT_ALGORITHM_IDENTIFIER(ticP, algidObj, &algid);
    if (res != TCL_OK)
        goto vamoose;

    if (! CryptSignAndEncodeCertificate(hprov, keyspec, enctype,
                                        (LPCSTR) (DWORD_PTR) structtype, &u,
                                        &algid, NULL, NULL, &nbytes))
        res = TwapiReturnSystemError(ticP->interp);
    else {
        encodedObj = ObjFromByteArray(NULL, nbytes);
        if (CryptSignAndEncodeCertificate(hprov, keyspec, enctype,
                                          (LPCSTR) (DWORD_PTR) structtype, &u,
                                          &algid, NULL,
                                          ObjToByteArray(encodedObj, NULL),
                                          &nbytes)) {
            Tcl_SetByteArrayLength(encodedObj, nbytes);
            ObjSetResult(ticP->interp, encodedObj);
        } else
            res = TwapiReturnSystemError(ticP->interp);
    }

vamoose:                       
    MemLifoPopMark(mark);
    return res;
}

BOOL WINAPI TwapiCertEnumSystemStoreCallback(
    const void *storeP,
    DWORD flags,
    PCERT_SYSTEM_STORE_INFO storeinfoP,
    void *reserved,
    Tcl_Obj *listObj
)
{
    Tcl_Obj *objs[2];

    if (flags & CERT_SYSTEM_STORE_RELOCATE_FLAG)
        return FALSE;     /* We do not know how to handle this currently */

    objs[0] = ObjFromUnicode(storeP);
    objs[1] = ObjFromDWORD(flags);
    ObjAppendElement(NULL, listObj, ObjNewList(2, objs));
    return TRUE;          /* Continue iteration */
}

BOOL WINAPI TwapiCertEnumSystemStoreLocationCallback(
    const void *locationP,
    DWORD flags,
    void *reserved,
    Tcl_Obj *listObj
)
{
    Tcl_Obj *objs[2];

    if (flags & CERT_SYSTEM_STORE_RELOCATE_FLAG)
        return FALSE;     /* We do not know how to handle this currently */

    objs[0] = ObjFromUnicode(locationP);
    objs[1] = ObjFromDWORD(flags);
    ObjAppendElement(NULL, listObj, ObjNewList(2, objs));
    return TRUE;          /* Continue iteration */
}

BOOL WINAPI TwapiCertEnumPhysicalStoreCallback(
    const void *system_storeP,
    DWORD flags,
    LPCWSTR store_nameP,
    PCERT_PHYSICAL_STORE_INFO storeinfoP,
    void *reserved,
    Tcl_Obj *listObj
)
{
    Tcl_Obj *objs[4];
    Tcl_Obj *infoObjs[6];

    if (flags & CERT_SYSTEM_STORE_RELOCATE_FLAG)
        return FALSE;     /* We do not know how to handle this currently */
    
    objs[0] = ObjFromUnicode(system_storeP);
    objs[1] = ObjFromDWORD(flags);
    objs[2] = ObjFromUnicode(store_nameP);

    infoObjs[0] = ObjFromString(storeinfoP->pszOpenStoreProvider);
    infoObjs[1] = ObjFromDWORD(storeinfoP->dwOpenEncodingType);
    infoObjs[2] = ObjFromDWORD(storeinfoP->dwOpenFlags);
    infoObjs[3] = ObjFromByteArray(storeinfoP->OpenParameters.pbData,
                                  storeinfoP->OpenParameters.cbData);
    infoObjs[4] = ObjFromDWORD(storeinfoP->dwFlags);
    infoObjs[5] = ObjFromDWORD(storeinfoP->dwPriority);
    objs[3] = ObjNewList(6, infoObjs);

    ObjAppendElement(NULL, listObj, ObjNewList(4, objs));
    return TRUE;          /* Continue iteration */
}

static BOOL WINAPI TwapiCryptEnumOIDInfoCB(PCCRYPT_OID_INFO coiP, void *pv)
{
    ObjAppendElement(NULL, (Tcl_Obj *)pv, ObjFromCRYPT_OID_INFO(coiP));
    return 1;
}

static TCL_RESULT Twapi_CertGetCertificateChainObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HCERTCHAINENGINE hce;
    PCCERT_CONTEXT certP;
    Tcl_Obj *paramObj, *ftObj;
    FILETIME  ft, *ftP;
    HCERTSTORE hstore;
    DWORD flags;
    MemLifoMarkHandle mark;
    TCL_RESULT res;
    CERT_CHAIN_PARA chain_params;
    PCCERT_CHAIN_CONTEXT chainP;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETHANDLET(hce, HCERTCHAINENGINE),
                     GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                     GETOBJ(ftObj),
                     GETVERIFIEDORNULL(hstore, HCERTSTORE, CertCloseStore),
                     GETOBJ(paramObj), GETINT(flags),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;
                     
    if (ObjToFILETIME(NULL, ftObj, &ft) == TCL_OK)
        ftP = &ft;
    else {
        if (ObjCharLength(ftObj) != 0)
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid time format");
        ftP = NULL;
    }

    mark = MemLifoPushMark(ticP->memlifoP);
    res = ParseCERT_CHAIN_PARA(ticP, paramObj, &chain_params);
    if (res == TCL_OK) {
        if (CertGetCertificateChain(hce, certP, ftP, hstore, &chain_params, flags, NULL, &chainP)) {
            TwapiRegisterPCCERT_CHAIN_CONTEXTTic(ticP, chainP);
            ObjSetResult(ticP->interp, ObjFromOpaque((void*)chainP, "PCCERT_CHAIN_CONTEXT"));
        } else
            res = TwapiReturnSystemError(ticP->interp);
    }
    
    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT  Twapi_HashPublicKeyInfoObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TCL_RESULT res;
    CERT_PUBLIC_KEY_INFO ckpi;
    Tcl_Obj *objP;
    DWORD    len;
    MemLifoMarkHandle mark;

    CHECK_NARGS(interp, objc, 2);

    mark = MemLifoPushMark(ticP->memlifoP);
    if ((res = ParseCERT_PUBLIC_KEY_INFO(ticP, objv[1], &ckpi)) == TCL_OK) {
        objP = ObjFromByteArray(NULL, 20);
        if (CryptHashPublicKeyInfo(0, CALG_SHA1, 0,
                                   X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
                                   &ckpi, 
                                   ObjToByteArray(objP, &len),
                                   &len)) {
            Tcl_SetByteArrayLength(objP, len);
            ObjSetResult(interp, objP);
            res = TCL_OK;
        } else {
            res = TwapiReturnSystemError(interp);
        }
    }

    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CryptFormatObjectObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD encoding, flags;
    void *encP, *bufP;
    int  enclen, buflen;
    TCL_RESULT res;
    MemLifoMarkHandle mark;
    char *oid;

    mark = MemLifoPushMark(ticP->memlifoP);
    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                         GETINT(encoding), ARGSKIP, GETINT(flags), ARGSKIP,
                         GETASTR(oid), GETBA(encP, enclen), ARGEND);
    if (res == TCL_OK) {
        /* First try a buffer size guess */
        bufP = MemLifoAlloc(ticP->memlifoP, enclen, &buflen);
        if (CryptFormatObject(encoding, 0, flags, NULL, oid, encP, enclen,
                              bufP, &buflen))
            ObjSetResult(interp, ObjFromUnicodeN(bufP, buflen/sizeof(WCHAR)));
        else
            res = TwapiReturnSystemError(interp);
    }

    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CryptEncodeObjectExObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    Tcl_Obj *typeObj, *valObj;
    TCL_RESULT res;
    CRYPT_OBJID_BLOB blob;
    MemLifoMarkHandle mark;

    mark = MemLifoPushMark(ticP->memlifoP);
    res = TwapiGetArgsEx(ticP, objc-1, objv+1, GETOBJ(typeObj),
                         GETOBJ(valObj), ARGEND);
    if (res == TCL_OK) {
        res = TwapiCryptEncodeObject(ticP, typeObj, valObj, &blob);
        if (res == TCL_OK)
            ObjSetResult(interp, ObjFromCRYPT_BLOB(&blob));
    }

    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CryptDecodeObjectExObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    Tcl_Obj *typeObj;
    void *encP;
    int  enc_len;
    Tcl_Obj *objP;
    TCL_RESULT res;
    MemLifoMarkHandle mark;

    mark = MemLifoPushMark(ticP->memlifoP);
    res = TwapiGetArgsEx(ticP, objc-1, objv+1, GETOBJ(typeObj),
                         GETBA(encP, enc_len), ARGEND);
    if (res == TCL_OK) {
        res = TwapiCryptDecodeObject(interp, typeObj, encP, enc_len, &objP);
        if (res == TCL_OK)
            ObjSetResult(interp, objP);
    }

    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CertVerifyChainPolicySSLObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    PCERT_CHAIN_POLICY_PARA policy_paramP;
    CERT_CHAIN_POLICY_STATUS policy_status;
    PCCERT_CHAIN_CONTEXT chainP;
    Tcl_Obj *paramObj;
    TCL_RESULT res;
    MemLifoMarkHandle mark;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVERIFIEDPTR(chainP, PCCERT_CHAIN_CONTEXT, CertFreeCertificateChain),
                     GETOBJ(paramObj), ARGEND) != TCL_OK)
        return TCL_ERROR;

    mark = MemLifoPushMark(ticP->memlifoP);
    res = ParseCERT_CHAIN_POLICY_PARA_SSL(ticP, paramObj, &policy_paramP);
    if (res == TCL_OK) {
        ZeroMemory(&policy_status, sizeof(policy_status));
        policy_status.cbSize = sizeof(policy_status);
        if (CertVerifyCertificateChainPolicy(CERT_CHAIN_POLICY_SSL,
                                             chainP,
                                             policy_paramP,
                                             &policy_status)) {
            ObjSetResult(interp, ObjFromDWORD(policy_status.dwError));
        } else
            res = TwapiReturnSystemError(interp);
    }

    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CertVerifyChainPolicyObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    PCERT_CHAIN_POLICY_PARA policy_paramP;
    CERT_CHAIN_POLICY_STATUS policy_status;
    TCL_RESULT (*param_parser)(TwapiInterpContext*, Tcl_Obj*, PCERT_CHAIN_POLICY_PARA*);
    PCCERT_CHAIN_CONTEXT chainP;
    Tcl_Obj *paramObj;
    TCL_RESULT res;
    DWORD policy;
    MemLifoMarkHandle mark;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(policy),
                     GETVERIFIEDPTR(chainP, PCCERT_CHAIN_CONTEXT, CertFreeCertificateChain),
                     GETOBJ(paramObj), ARGEND) != TCL_OK)
        return TCL_ERROR;
    switch (policy) {
    case 4: // CERT_CHAIN_POLICY_SSL
        param_parser = ParseCERT_CHAIN_POLICY_PARA_SSL;
        break;
    case 1: // CERT_CHAIN_POLICY_BASE
    case 5: // CERT_CHAIN_POLICY_BASIC_CONSTRAINTS
    case 6: // CERT_CHAIN_POLICY_NT_AUTH
    case 7: // CERT_CHAIN_POLICY_MICROSOFT_ROOT
    case 8: // CERT_CHAIN_POLICY_EV
        param_parser = ParseCERT_CHAIN_POLICY_PARA;
        break;
    default:
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid certificate policy identifier.");
    } 
    mark = MemLifoPushMark(ticP->memlifoP);
    res = (*param_parser)(ticP, paramObj, &policy_paramP);
    if (res == TCL_OK) {
        ZeroMemory(&policy_status, sizeof(policy_status));
        policy_status.cbSize = sizeof(policy_status);
        if (CertVerifyCertificateChainPolicy((LPCSTR) (DWORD_PTR)policy,
                                             chainP,
                                             policy_paramP,
                                             &policy_status)) {
            ObjSetResult(interp, ObjFromDWORD(policy_status.dwError));
        } else
            res = TwapiReturnSystemError(interp);
    }

    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CertChainContextsObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    PCCERT_CHAIN_CONTEXT chainP;
    Tcl_Obj *objs[2];
    DWORD dw;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVERIFIEDPTR(chainP, PCCERT_CHAIN_CONTEXT, CertFreeCertificateChain),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    objs[0] = ObjFromCERT_TRUST_STATUS(&chainP->TrustStatus);
    objs[1] = ObjNewList(0, NULL);
    for (dw = 0; dw < chainP->cChain; ++dw) {
        ObjAppendElement(NULL, objs[1], ObjFromCERT_SIMPLE_CHAIN(interp, chainP->rgpChain[dw]));
    }
    ObjSetResult(interp, ObjNewList(2, objs));
    return TCL_OK;
}

static int Twapi_CryptFindOIDInfoObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD keytype, group;
    Tcl_Obj *keyObj;
    void *pv;
    ALG_ID algids[2];
    Tcl_Obj **objs;
    int nobjs;
    PCCRYPT_OID_INFO coiP;
#if 0
#define CRYPT_OID_INFO_OID_KEY           1
#define CRYPT_OID_INFO_NAME_KEY          2
#define CRYPT_OID_INFO_ALGID_KEY         3
#define CRYPT_OID_INFO_SIGN_KEY          4

#endif
    
    if (TwapiGetArgs(interp, objc-1, objv+1, GETINT(keytype), GETOBJ(keyObj),
                     GETINT(group), ARGEND) != TCL_OK)
        return TCL_ERROR;

    switch (keytype & ~CRYPT_OID_INFO_OID_KEY_FLAGS_MASK) {
    case CRYPT_OID_INFO_OID_KEY:
        pv = ObjToString(keyObj);
        break;
    case CRYPT_OID_INFO_NAME_KEY:
        pv = ObjToUnicode(keyObj);
        break;
    case CRYPT_OID_INFO_ALGID_KEY:
        if (ObjToDWORD(interp, keyObj, &algids[0]) != TCL_OK)
            return TCL_ERROR;
        pv = &algids[0];
        break;
    case CRYPT_OID_INFO_SIGN_KEY:
        if (ObjGetElements(NULL, keyObj, &nobjs, &objs) != TCL_OK ||
            nobjs != 2 ||
            ObjToDWORD(NULL, objs[0], &algids[0]) != TCL_OK ||
            ObjToDWORD(NULL, objs[1], &algids[1]) != TCL_OK) {
            TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid CRYPT_OID_INFO_SIGN_KEY format");
            return TCL_ERROR;
        }
        pv = algids;
        break;
    }

    /* NOTE: coiP must NOT be freed */
    coiP = CryptFindOIDInfo(keytype, pv, group);
    if (coiP)
        ObjSetResult(interp, ObjFromCRYPT_OID_INFO(coiP));
    /* Else empty result */

    return TCL_OK;
}

static int Twapi_CertSetCertificateContextPropertyObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD prop_id, flags;
    Tcl_Obj *valObj;
    void *pv;
    CRYPT_DATA_BLOB blob;
    TCL_RESULT res;
    MemLifoMarkHandle mark;
    CRYPT_KEY_PROV_INFO *kiP;
    PCCERT_CONTEXT certP;

    mark = MemLifoPushMark(ticP->memlifoP);
    
    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         GETINT(prop_id), GETINT(flags), ARGUSEDEFAULT,
                         GETOBJ(valObj), ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    if (valObj == NULL)
        pv = NULL;              /* Property to be deleted */
    else {
        switch (prop_id) {
        case CERT_KEY_PROV_INFO_PROP_ID:
            res = ParseCRYPT_KEY_PROV_INFO(ticP, valObj, &kiP);
            pv = kiP;
            break;

        case CERT_ENHKEY_USAGE_PROP_ID:
        case CERT_DESCRIPTION_PROP_ID:
        case CERT_ARCHIVED_PROP_ID:
        case CERT_FRIENDLY_NAME_PROP_ID:
            pv = &blob;
            res = ParseCRYPT_BLOB(ticP, valObj, &blob);
            break;
        default:
            res = TwapiReturnError(interp, TWAPI_UNSUPPORTED_TYPE);
            break;
        }
    }
    if (res != TCL_OK)
        goto vamoose;

    if (CertSetCertificateContextProperty(certP, prop_id, flags, pv))
        res = TCL_OK;
    else
        res = TwapiReturnSystemError(interp);

vamoose:
    MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CryptQueryObjectObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD operand_type, enc_type;
    Tcl_Obj *objP;
    void *operand;
    DWORD expected_content_type, expected_format_type, flags;
    int types_only;
    HCERTSTORE hstore = NULL, *hstoreP;
    HCRYPTMSG hmsg = NULL, *hmsgP;
    void *pv = NULL;
    void **pvP;
    CERT_BLOB blob;
    MemLifoMarkHandle mark = NULL;
    TCL_RESULT res;

    res = TwapiGetArgs(interp, objc-1, objv+1, GETINT(operand_type),
                       GETOBJ(objP), GETINT(expected_content_type),
                       GETINT(expected_format_type), GETINT(flags),
                       ARGUSEDEFAULT, GETBOOL(types_only),
                       ARGEND);
    if (res != TCL_OK)
        return res;
    if (types_only) {
        hstoreP = NULL;
        hmsgP   = NULL;
        pvP     = NULL;
    } else {
        hstoreP = &hstore;
        hmsgP   = &hmsg;
        pvP     = &pv;
    }
    res = TCL_OK;
    if (operand_type == CERT_QUERY_OBJECT_FILE) {
        operand = ObjToUnicode(objP);
    } else {
        mark = MemLifoPushMark(ticP->memlifoP);
        operand = &blob;
        res = ParseCRYPT_BLOB(ticP, objP, operand);
    }
    if (res == TCL_OK) {
        if (CryptQueryObject(operand_type, operand, 
                             expected_content_type, expected_format_type,
                             flags, &enc_type, &expected_content_type,
                             &expected_format_type, hstoreP, hmsgP,
                             pvP) == 0)
            res = TwapiReturnSystemError(interp);
        else {
            Tcl_Obj *objs[12];
            int nobjs;
            objs[0] = STRING_LITERAL_OBJ("encoding");
            objs[1] = ObjFromDWORD(enc_type);
            objs[2] = STRING_LITERAL_OBJ("formattype");
            objs[3] = ObjFromDWORD(expected_format_type);
            objs[4] = STRING_LITERAL_OBJ("contenttype");
            objs[5] = ObjFromDWORD(expected_content_type);
            nobjs = 6;
            if (hstore) {
                TwapiRegisterHCERTSTORETic(ticP, hstore);
                objs[nobjs++] = STRING_LITERAL_OBJ("certstore");
                objs[nobjs++] = ObjFromOpaque(hstore, "HCERTSTORE");
            }
            if (hmsg) {
                TwapiRegisterHCRYPTMSGTic(ticP, hmsg);
                objs[nobjs++] = STRING_LITERAL_OBJ("cryptmsg");
                objs[nobjs++] = ObjFromOpaque(hstore, "HCRYPTMSG");
            }
            if (pvP) {
                switch (expected_content_type) {
                case CERT_QUERY_CONTENT_CERT:
                case CERT_QUERY_CONTENT_SERIALIZED_CERT:
                    TwapiRegisterPCCERT_CONTEXTTic(ticP, (PCCERT_CONTEXT)pv);
                    objs[nobjs++] = STRING_LITERAL_OBJ("cert_context");
                    objs[nobjs++] = ObjFromOpaque(pv, "PCCERT_CONTEXT");
                    break;
                case CERT_QUERY_CONTENT_CRL:
                case CERT_QUERY_CONTENT_SERIALIZED_CRL:
                    TwapiRegisterPCCRL_CONTEXTTic(ticP, (PCCRL_CONTEXT)pv);
                    objs[nobjs++] = STRING_LITERAL_OBJ("crl_context");
                    objs[nobjs++] = ObjFromOpaque(pv, "PCCRL_CONTEXT");
                    break;
                case CERT_QUERY_CONTENT_CTL:
                case CERT_QUERY_CONTENT_SERIALIZED_CTL:
                    TwapiRegisterPCCTL_CONTEXTTic(ticP, (PCCTL_CONTEXT)pv);
                    objs[nobjs++] = STRING_LITERAL_OBJ("ctl_context");
                    objs[nobjs++] = ObjFromOpaque(pv, "PCCTL_CONTEXT");
                    break;
                }
            }
                
            ObjSetResult(interp, ObjNewList(nobjs, objs));
            res = TCL_OK;
        }
    } 
    if (mark)
        MemLifoPopMark(mark);
    return res;
}

static TCL_RESULT Twapi_CryptGetKeyParam(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD winerr, param, flags;
    int nbytes;
    void *p;
    Tcl_Obj *objP = NULL;
    TCL_RESULT res;
    HCRYPTKEY hkey;
    DWORD dw;
    
    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVERIFIEDPTR(hkey, HCRYPTKEY, CryptDestroyKey),
                     GETINT(param), ARGUSEDEFAULT,
                     GETINT(flags), ARGEND) != TCL_OK)
        return TCL_ERROR;

    switch (param) {
    case KP_ALGID:   case KP_BLOCKLEN: case KP_KEYLEN: case KP_PERMISSIONS:
    case KP_P:       case KP_Q:        case KP_G:      case KP_EFFECTIVE_KEYLEN:
    case KP_PADDING: case KP_MODE:     case KP_MODE_BITS:
        p = &dw;
        nbytes = sizeof(dw);
        break;
    case KP_VERIFY_PARAMS:
        /* Special case, no data is actually returned */
        p = NULL;
        nbytes = 0;
        break;
    default:
        nbytes = 0;
        CryptGetKeyParam(hkey, param, NULL, &nbytes, flags);
        winerr = GetLastError();
        if (winerr != ERROR_MORE_DATA)
            return Twapi_AppendSystemError(interp, winerr);
        /* Create a bytearray in two steps because Tcl 8.5 is not
           documented as accepting a NULL in Tcl_NewByteArrayObj */
        objP = Tcl_NewObj();
        Tcl_SetByteArrayObj(objP, NULL, nbytes);
        p = Tcl_GetByteArrayFromObj(objP, &nbytes);
        break;
    }
    if (!CryptGetKeyParam(hkey, param, p, &nbytes, flags)) {
        if (objP)
            ObjDecrRefs(objP);
        res = TwapiReturnSystemError(interp);
    } else {
        switch (param) {
        case KP_ALGID:   case KP_BLOCKLEN: case KP_KEYLEN: case KP_PERMISSIONS:
        case KP_P:       case KP_Q:     case KP_G:  case KP_EFFECTIVE_KEYLEN:
        case KP_PADDING: case KP_MODE:  case KP_MODE_BITS:
            res = ObjSetResult(interp, ObjFromDWORD(dw));
            break;
        case KP_VERIFY_PARAMS:
            /* Special case, no data is actually returned */
            res = TCL_OK;
            break;
        default:
            TWAPI_ASSERT(objP != NULL);
            res = ObjSetResult(interp, objP);
            break;
        }
    }

    return res;
}

static TCL_RESULT TwapiCloseContext(Tcl_Interp *interp, Tcl_Obj *objP,
                                    const char *typeptr,
                                    TCL_RESULT (*unregfn)(Tcl_Interp *, HANDLE),
                                    BOOL (WINAPI *freefn)(HANDLE))
{
    HANDLE h;
    if (ObjToOpaque(interp, objP, &h, typeptr) != TCL_OK ||
        (*unregfn)(interp, h) != TCL_OK)
        return TCL_ERROR;

    if ((*freefn)(h) == FALSE) {
        DWORD dw = GetLastError();
        if (dw != CRYPT_E_PENDING_CLOSE)
            return TwapiReturnSystemError(interp);
    }
    return TCL_OK;
}

static TCL_RESULT Twapi_CryptCATAdminEnumCatalogFromHash(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TCL_RESULT res;
    HCATADMIN hca;
    HCATINFO  hci, prev_hci;
    BYTE *pb;
    DWORD cb, flags;
    MemLifoMarkHandle mark = NULL;
    
    mark = MemLifoPushMark(ticP->memlifoP);
    res = TwapiGetArgsEx(ticP, objc-1, objv+1,
                       GETVERIFIEDPTR(hca, HCATADMIN, CryptCATAdminReleaseContext),
                       GETBA(pb, cb), GETINT(flags),
                       ARGUSEDEFAULT,
                       GETVERIFIEDORNULL(prev_hci, HCATINFO, CryptCATAdminReleaseCatalogContext),
                       ARGEND);
    if (res == TCL_OK) {
        /* The previous HCATINFO will be freed by this call so unregister it */
        if (prev_hci)
            TwapiUnregisterHCATINFO(interp, prev_hci);
        hci = CryptCATAdminEnumCatalogFromHash(hca, pb, cb, flags, &prev_hci);
        if (hci) {
            TwapiRegisterHCATINFO(interp, hci);
            res = ObjSetResult(interp, ObjFromOpaque(hci, "HCATINFO"));
        } else {
            DWORD winerr = GetLastError();
            if (winerr != ERROR_NOT_FOUND)
                res = TwapiReturnSystemError(interp);
            /* else just return empty string */
        }
    }
    
vamoose:
    if (mark)
        MemLifoPopMark(mark);
    return res;
}
        
static TCL_RESULT Twapi_CryptoCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    DWORD dw, dw2, dw3;
    DWORD_PTR dwp;
    LPVOID pv;
    LPWSTR s1;
    LPSTR  cP;
    struct _CRYPTOAPI_BLOB blob, blob2;
    PCCERT_CONTEXT certP, cert2P;
    void *bufP;
    DWORD buf_sz;
    Tcl_Obj *s1Obj, *s2Obj;
    BOOL bval;
    int func = PtrToInt(clientdata);
    Tcl_Obj *objs[11];
    TCL_RESULT res;
    CERT_INFO *ciP;
    HCERTSTORE hstore;
    HANDLE h, h2;
    union {
        GUID guid;
        WCHAR uni[MAX_PATH+1];
        char ansi[MAX_PATH+1];
        CATALOG_INFO catinfo;
    } buf;

    --objc;
    ++objv;

    TWAPI_ASSERT(sizeof(HCRYPTPROV) <= sizeof(pv));
    TWAPI_ASSERT(sizeof(HCRYPTKEY) <= sizeof(pv));
    TWAPI_ASSERT(sizeof(dwp) <= sizeof(void*));

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 10000: // CryptAcquireContext
         if (TwapiGetArgs(interp, objc, objv,
                         GETOBJ(s1Obj), GETOBJ(s2Obj), GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (CryptAcquireContextW(&dwp,
                                 ObjToLPWSTR_NULL_IF_EMPTY(s1Obj),
                                 ObjToLPWSTR_NULL_IF_EMPTY(s2Obj),
                                 dw, dw2)) {
            if (dw2 & CRYPT_DELETEKEYSET)
                result.type = TRT_EMPTY;
            else {
                TwapiRegisterHCRYPTPROV(interp, dwp);
                TwapiResult_SET_PTR(result, HCRYPTPROV, (void*)dwp);
            }
        } else {
            result.type = TRT_GETLASTERROR;
        }
        break;

    case 10001: // CryptReleaseContext
        /* We can use GETHANDLET instead of GETVERIFIEDPTR here because
           anyways it is followed by TwapiUnregisterHCRYPTPROV which
           does the verification anyways */
        if (TwapiGetArgs(interp, objc, objv,
                         GETPTR(pv, HCRYPTPROV),
                         ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK
            || TwapiUnregisterHCRYPTPROV(interp, (HCRYPTPROV)pv) != TCL_OK)
            return TCL_ERROR;
        result.value.ival = CryptReleaseContext((HCRYPTPROV)pv, dw);
        result.type = TRT_EXCEPTION_ON_FALSE;
        break;

    case 10002: // CryptGetProvParam
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_CryptGetProvParam(interp, (HCRYPTPROV) pv, dw, dw2);

    case 10003: // CertOpenSystemStore
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        pv = CertOpenSystemStoreW(0, ObjToUnicode(objv[0]));
        if (pv) {
            /* CertCloseStore does not check ponter validity! So do ourselves*/
            TwapiRegisterHCERTSTORE(interp, pv);
            TwapiResult_SET_NONNULL_PTR(result, HCERTSTORE, pv);
        } else {
            return TwapiReturnSystemError(interp);
        }
        break;

    case 10014: // CertFreeCertificateContext
        CHECK_NARGS(interp, objc, 1);
        return TwapiCloseContext(interp, objv[0], "PCCERT_CONTEXT",
                                 TwapiUnregisterPCCERT_CONTEXT, CertFreeCertificateContext);
        
    case 10004:
    case 10015:
    case 10035:
    case 10036:
    case 10041:
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 10004: // CertDeleteCertificateFromStore
            /* Unregister previous context since the next call will free it,
               EVEN ON FAILURES */
            if (TwapiUnregisterPCCERT_CONTEXT(interp, certP) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CertDeleteCertificateFromStore(certP);
            break;
        case 10015: // Twapi_CertGetEncoded
            if (certP->pbCertEncoded && certP->cbCertEncoded) {
                objs[0] = ObjFromDWORD(certP->dwCertEncodingType);
                objs[1] = ObjFromByteArray(certP->pbCertEncoded, certP->cbCertEncoded);
                result.value.objv.objPP = objs;
                result.value.objv.nobj = 2;
                result.type = TRT_OBJV;
            }
            /* else empty result - TBD */
            break;
        case 10035: //Twapi_CertGetInfo
            ciP = certP->pCertInfo;
            if (ciP) {
                objs[0] = ObjFromInt(ciP->dwVersion);
                objs[1] = ObjFromCRYPT_BLOB(&ciP->SerialNumber);
                objs[2] = ObjFromCRYPT_ALGORITHM_IDENTIFIER(&ciP->SignatureAlgorithm);
                objs[3] = ObjFromCERT_NAME_BLOB(&ciP->Issuer, CERT_X500_NAME_STR);
                objs[4] = ObjFromFILETIME(&ciP->NotBefore);
                objs[5] = ObjFromFILETIME(&ciP->NotAfter);
                objs[6] = ObjFromCERT_NAME_BLOB(&ciP->Subject, CERT_X500_NAME_STR);
                objs[7] = ObjFromCERT_PUBLIC_KEY_INFO(&ciP->SubjectPublicKeyInfo);
                objs[8] = ObjFromCRYPT_BIT_BLOB(&ciP->IssuerUniqueId);
                objs[9] = ObjFromCRYPT_BIT_BLOB(&ciP->SubjectUniqueId);
                objs[10] = ObjFromCERT_EXTENSIONS(ciP->cExtension, ciP->rgExtension);
                result.value.objv.nobj = 11;
            } else
                result.value.objv.nobj = 0;

            result.value.objv.objPP = objs;
            result.type = TRT_OBJV;

            break;
        case 10036: //Twapi_CertGetExtensions
            ciP = certP->pCertInfo;
            if (ciP) {
                result.value.obj = ObjFromCERT_EXTENSIONS(ciP->cExtension, ciP->rgExtension);
                result.type = TRT_OBJ;
            } else
                result.type = TRT_EMPTY;
            break;
        case 10041: //CertDuplicateCertificateContext
            certP = CertDuplicateCertificateContext(certP);
            TwapiRegisterPCCERT_CONTEXT(interp, certP);
            TwapiResult_SET_NONNULL_PTR(result, PCCERT_CONTEXT, (void*)certP);
            break;
        }
        break;

    case 10005: // CertCreateCertificateContext
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
        cP = ObjToByteArray(objv[1], &dw2);
        certP = CertCreateCertificateContext(dw, cP, dw2);
        if (certP) {
            TwapiRegisterPCCERT_CONTEXT(interp, certP);
            TwapiResult_SET_NONNULL_PTR(result, PCCERT_CONTEXT, (void*)certP);
        } else
            result.type = TRT_GETLASTERROR;
        break;

    case 10006: // CertEnumCertificatesInStore
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCERTSTORE, CertCloseStore),
                         GETPTR(certP, PCCERT_CONTEXT), ARGEND) != TCL_OK)
            return TCL_ERROR;
        /* Unregister previous context since the next call will free it */
        if (certP &&
            TwapiUnregisterPCCERT_CONTEXT(interp, certP) != TCL_OK)
            return TCL_ERROR;
        certP = CertEnumCertificatesInStore(pv, certP);
        if (certP) {
            TwapiRegisterPCCERT_CONTEXT(interp, certP);
            TwapiResult_SET_NONNULL_PTR(result, PCCERT_CONTEXT, (void*)certP);
        } else {
            result.value.ival = GetLastError();
            if (result.value.ival == CRYPT_E_NOT_FOUND ||
                result.value.ival == ERROR_NO_MORE_FILES)
                result.type = TRT_EMPTY;
            else
                result.type = TRT_GETLASTERROR;
        }
        break;
    case 10007: // CertEnumCertificateContextProperties
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_DWORD;
        result.value.ival = CertEnumCertificateContextProperties(certP, dw);
        break;

    case 10008: // CertGetCertificateContextProperty
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         GETINT(dw), ARGUSEDEFAULT, GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_CertGetCertificateContextProperty(interp, certP, dw, dw2);

    case 10009: // CryptDestroyKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETPTR(pv, HCRYPTKEY),
                         ARGEND) != TCL_OK
            || TwapiUnregisterHCRYPTKEY(interp, (HCRYPTKEY) pv) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CryptDestroyKey((HCRYPTKEY) pv);
        break;
            
    case 10010: // CryptGenKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         GETINT(dw), GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (CryptGenKey((HCRYPTPROV) pv, dw, dw2, &dwp)) {
            TwapiRegisterHCRYPTKEY(interp, dwp);
            TwapiResult_SET_PTR(result, HCRYPTKEY, (void*)dwp);
        } else
            result.type = TRT_GETLASTERROR;
        break;

    case 10011: // CertStrToName
        if (TwapiGetArgs(interp, objc, objv, GETOBJ(s1Obj), ARGUSEDEFAULT,
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_GETLASTERROR;
        dw2 = 0;
        s1 = ObjToUnicode(s1Obj); /* Do AFTER extracting other args above */
        if (CertStrToNameW(X509_ASN_ENCODING, s1,
                           dw, NULL, NULL, &dw2, NULL)) {
            result.value.obj = ObjFromByteArray(NULL, dw2);
            if (CertStrToNameW(X509_ASN_ENCODING, s1, dw, NULL,
                               ObjToByteArray(result.value.obj, &dw2),
                               &dw2, NULL)) {
                Tcl_SetByteArrayLength(result.value.obj, dw2);
                result.type = TRT_OBJ;
            } else {
                ObjDecrRefs(result.value.obj);
            }
        }
        break;

    case 10012: // CertNameToStr
        if (TwapiGetArgs(interp, objc, objv, ARGSKIP, GETINT(dw), ARGEND)
            != TCL_OK)
            return TCL_ERROR;
        blob.pbData = ObjToByteArray(objv[0], &blob.cbData);
        result.type = TRT_OBJ;
        result.value.obj = ObjFromCERT_NAME_BLOB(&blob, dw);
        break;

    case 10013: // CertGetNameString
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         GETINT(dw), GETINT(dw2), ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
            
        return TwapiCertGetNameString(interp, certP, dw, dw2, objv[3]);

            
    case 10016: // CertUnregisterSystemStore
        /* This command is there to primarily clean up mistakes in testing */
        if (TwapiGetArgs(interp, objc, objv,
                         GETOBJ(s1Obj), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CertUnregisterSystemStore(ObjToUnicode(s1Obj), dw);
        break;
    case 10017: // CertCloseStore
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(pv, HCERTSTORE), ARGUSEDEFAULT,
                         GETINT(dw), ARGEND) != TCL_OK ||
            TwapiUnregisterHCERTSTORE(interp, pv) != TCL_OK)
            return TCL_ERROR;

        result.type = TRT_BOOL;
        result.value.bval = CertCloseStore(pv, dw);
        if (result.value.bval == FALSE) {
            if (GetLastError() != CRYPT_E_PENDING_CLOSE)
                result.type = TRT_GETLASTERROR;
        }
        break;

    case 10018: // CryptGetUserKey
    case 10034: // CryptGenRandom
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 10018: // CryptGetUserKey
            if (CryptGetUserKey((HCRYPTPROV) pv, dw, &dwp)) {
                TwapiRegisterHCRYPTKEY(interp, dwp);
                TwapiResult_SET_PTR(result, HCRYPTKEY, (void*)dwp);
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 10034:
            return Twapi_CryptGenRandom(interp, (HCRYPTPROV) pv, dw);
        }
        break;

    case 10019: // CryptSetProvParam
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         GETINT(dw), GETINT(dw2), ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_CryptSetProvParam(interp, (HCRYPTPROV) pv, dw, dw2, objv[3]);

    case 10020: // CertOpenStore
        return Twapi_CertOpenStore(interp, objc, objv);

    case 10021: // CryptEnumOIDInfo
        CHECK_NARGS(interp, objc, 1);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        result.value.obj = ObjNewList(0, NULL);
        CryptEnumOIDInfo(dw, 0, result.value.obj, TwapiCryptEnumOIDInfoCB);
        result.type = TRT_OBJ;
        break;

    case 10022: // CertAddCertificateContextToStore
        /* TBD - add option to not require added certificate context
           to be returned */
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCERTSTORE, CertCloseStore),
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (!CertAddCertificateContextToStore(pv, certP, dw, &certP))
            result.type = TRT_GETLASTERROR;
        else {
            TwapiRegisterPCCERT_CONTEXT(interp, certP);
            TwapiResult_SET_NONNULL_PTR(result, PCCERT_CONTEXT, (void*)certP);
        }
        break;

    case 10023:  // CryptExportPublicKeyInfoEx
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         GETINT(dw), // keyspec
                         GETINT(dw2), // enctype
                         GETASTR(cP), // publickeyobjid
                         GETINT(dw3), // flags
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        buf_sz = 0;
        if (!CryptExportPublicKeyInfoEx((HCRYPTPROV)pv, dw, dw2, cP, dw3, NULL, NULL, &buf_sz)) {
            result.type = TRT_GETLASTERROR;
            break;
        }
        bufP = TwapiAlloc(buf_sz);
        if (CryptExportPublicKeyInfoEx((HCRYPTPROV)pv, dw, dw2, cP, dw3, NULL, bufP, &buf_sz)) {
            result.type = TRT_OBJ;
            result.value.obj = ObjFromCERT_PUBLIC_KEY_INFO(bufP);
        } else {
            result.type = TRT_TCL_RESULT;
            result.value.ival = TwapiReturnSystemError(interp);
        }
        TwapiFree(bufP);
        break;

    case 10024:
        CHECK_NARGS(interp, objc, 2);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        if (dw & CERT_SYSTEM_STORE_RELOCATE_FLAG)
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "RELOCATE flag not supported.");
        result.value.obj = ObjNewList(0, NULL);
        if (CertEnumSystemStore(dw, ObjToLPWSTR_NULL_IF_EMPTY(objv[1]),
                                result.value.obj,
                                TwapiCertEnumSystemStoreCallback)) {
            result.type = TRT_OBJ;
        } else {
            TwapiReturnSystemError(interp);
            ObjDecrRefs(result.value.obj);
            return TCL_ERROR;
        }
        break;

    case 10025:
        CHECK_NARGS(interp, objc, 2);
        CHECK_INTEGER_OBJ(interp, dw, objv[1]);
        if (dw & CERT_SYSTEM_STORE_RELOCATE_FLAG)
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "RELOCATE flag not supported.");
        result.value.obj = ObjNewList(0, NULL);
        if (CertEnumPhysicalStore(ObjToUnicode(objv[0]), dw,
                                result.value.obj,
                                TwapiCertEnumPhysicalStoreCallback)) {
            result.type = TRT_OBJ;
        } else {
            TwapiReturnSystemError(interp);
            ObjDecrRefs(result.value.obj);
            return TCL_ERROR;
        }
        break;

    case 10026:
        CHECK_NARGS(interp, objc, 1);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        if (dw & CERT_SYSTEM_STORE_RELOCATE_FLAG)
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "RELOCATE flag not supported.");
        result.value.obj = ObjNewList(0, NULL);
        if (CertEnumSystemStoreLocation(dw, result.value.obj,
                                        TwapiCertEnumSystemStoreLocationCallback)) {
            result.type = TRT_OBJ;
        } else {
            TwapiReturnSystemError(interp);
            ObjDecrRefs(result.value.obj);
            return TCL_ERROR;
        }
        break;

    case 10027:
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         GETINT(dw), ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
        /* We only allow the following flags */
        if (dw & ~(CRYPT_ACQUIRE_CACHE_FLAG|CRYPT_ACQUIRE_COMPARE_KEY_FLAG|CRYPT_ACQUIRE_SILENT_FLAG|CRYPT_ACQUIRE_USE_PROV_INFO_FLAG)) {
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid flags");
        }
        if (CryptAcquireCertificatePrivateKey(certP,dw,NULL,&dwp,&dw2,&bval)) {
            TwapiRegisterHCRYPTPROV(interp, dwp);
            objs[0] = ObjFromOpaque((void*)dwp, "HCRYPTPROV");
            objs[1] = ObjFromLong(dw2);
            objs[2] = ObjFromBoolean(bval);
            result.value.objv.objPP = objs;
            result.value.objv.nobj = 3;
            result.type = TRT_OBJV;
        } else
            result.type = TRT_GETLASTERROR;
        break;

    case 10028: // CertGetEnhancedKeyUsage
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        dw2 = 0;
        result.type = TRT_GETLASTERROR;
        if (CertGetEnhancedKeyUsage(certP, dw, NULL, &dw2)) {
            pv = TwapiAlloc(dw2);
            if (CertGetEnhancedKeyUsage(certP, dw, pv, &dw2)) {
                CERT_ENHKEY_USAGE *ceuP = pv;
                result.type = TRT_OBJ;
                if (ceuP->cUsageIdentifier)
                    result.value.obj = ObjFromArgvA(ceuP->cUsageIdentifier,
                                                    ceuP->rgpszUsageIdentifier);
                else {
                    if (GetLastError() == CRYPT_E_NOT_FOUND) {
                        /* Extension not present -> all uses valid */
                        result.value.obj = STRING_LITERAL_OBJ("*");
                    } else /* No valid uses */
                        result.type = TRT_EMPTY;
                }
            }
            TwapiFree(pv);
        } else {
            if (GetLastError() == CRYPT_E_NOT_FOUND) {
                /* Extension/Property not present -> all uses valid */
                result.type = TRT_OBJ;
                result.value.obj = STRING_LITERAL_OBJ("*");
            }
        }
        break;

    case 10029: // Twapi_CertStoreCommit
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCERTSTORE, CertCloseStore),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (CertControlStore(pv, dw ? CERT_STORE_CTRL_COMMIT_FORCE_FLAG : 0,
                               CERT_STORE_CTRL_COMMIT, NULL))
            result.type = TRT_EMPTY;
        else
            result.type = TRT_GETLASTERROR;
        break;

    case 10030: // Twapi_CertGetIntendedKeyUsage
        if (TwapiGetArgs(interp, objc, objv, GETINT(dw),
                         GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        else {
            BYTE bytes[2];

            /* We do not currently support more than 2 bytes at Tcl level */
            dw = CertGetIntendedKeyUsage(dw, certP->pCertInfo, bytes, ARRAYSIZE(bytes));
            if (dw) {
                result.value.binary.p = bytes;
                result.value.binary.len = ARRAYSIZE(bytes);
                result.type = TRT_BINARY;
            } else {
                if (GetLastError() == 0)
                    result.type = TRT_EMPTY;
                else
                    result.type = TRT_GETLASTERROR;
            }
        }
        break;

    case 10031: // CertGetIssuerCertificateFromStore
        cert2P = NULL;
        res = TwapiGetArgs(interp, objc, objv,
                           GETVERIFIEDPTR(pv, HCERTSTORE, CertCloseStore),
                           GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                           GETVERIFIEDORNULL(cert2P, PCCERT_CONTEXT, CertFreeCertificateContext),
                           GETINT(dw), ARGEND);
        
        if (cert2P) {
            /* CertGetIssuerCertificateFromStore frees cert2P */
            if (TwapiUnregisterPCCERT_CONTEXT(interp, cert2P) != TCL_OK)
                return TCL_ERROR; /* Bad pointer, don't do anything more */
        }
        if (res != TCL_OK) {
            if (cert2P)
                CertFreeCertificateContext(cert2P); /* That's what we have
                                                       guaranteed caller */
            return res;
        }
        cert2P = CertGetIssuerCertificateFromStore(pv, certP, cert2P, &dw);
        if (cert2P) {
            TwapiRegisterPCCERT_CONTEXT(interp, cert2P);
            objs[0] = ObjFromOpaque((void*)cert2P, "PCCERT_CONTEXT");
            objs[1] = ObjFromDWORD(dw);
            result.type= TRT_OBJV;
            result.value.objv.objPP = objs;
            result.value.objv.nobj = 2;
        } else {
            result.type =
                GetLastError() == CRYPT_E_NOT_FOUND ?
                TRT_EMPTY : TRT_GETLASTERROR;
        }
        break;

    case 10032: // CertFreeCertificateChain
        CHECK_NARGS(interp, objc, 1);
        return TwapiCloseContext(interp, objv[0], "PCCERT_CHAIN_CONTEXT",
                                 TwapiUnregisterPCCERT_CHAIN_CONTEXT, TwapiCertFreeCertificateChain);
        
    case 10033: // CertFindExtension
        res = TwapiGetArgs(interp, objc, objv,
                           GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                           ARGSKIP, ARGEND);
        if (res != TCL_OK)
            return res;
        result.type = TRT_EMPTY; /* Assume extension does not exist */
        ciP = certP->pCertInfo;
        if (ciP && ciP->cExtension && ciP->rgExtension) {
            CERT_EXTENSION *extP =
                CertFindExtension(ObjToString(objv[1]),
                                  ciP->cExtension, ciP->rgExtension);
            if (extP) {
                result.value.obj = ObjFromCERT_EXTENSION(extP);
                result.type = TRT_OBJ;
            }
        }
        break;
    case 10037: // CryptFindCertificateKeyProvInfo
        res = TwapiGetArgs(interp, objc, objv,
                           GETVERIFIEDPTR(certP, PCCERT_CONTEXT, CertFreeCertificateContext),
                           GETINT(dw), ARGEND);
        if (res != TCL_OK)
            return res;
        result.type = TRT_BOOL;
        result.value.bval = CryptFindCertificateKeyProvInfo(certP, dw, NULL);
        break;
    case 10038: // CertAddEncodedCertificateToStore
        /* TBD - chromium source code comments suggest cert store handle
           can be NULL. Can we make use of this to implement a command to
           decode a cert without needing a temp store */
        /* TBD - add option to not require added certificate context
           to be returned */
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCERTSTORE, CertCloseStore),
                         GETINT(dw), ARGSKIP, GETINT(dw2), ARGEND) != TCL_OK)
            return TCL_ERROR;
        cP = ObjToByteArray(objv[2], &dw3);
        if (CertAddEncodedCertificateToStore(pv, dw, cP, dw3, dw2, &certP)) {
            TwapiRegisterPCCERT_CONTEXT(interp, certP);
            TwapiResult_SET_NONNULL_PTR(result, PCCERT_CONTEXT, (void*)certP);
        } else
            result.type = TRT_GETLASTERROR;
        break;
    case 10039: // CertOIDToAlgId
        CHECK_NARGS(interp, objc, 1);
        result.type = TRT_DWORD;
        result.value.uval = CertOIDToAlgId(ObjToString(objv[0]));
        break;
    case 10040: // CertAlgIdToOID
        CHECK_NARGS(interp, objc, 1);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        result.type = TRT_CHARS;
        result.value.chars.str = (char *) CertAlgIdToOID(dw);
        result.value.chars.len = -1;
        break;

        // case 10041 is above

    case 10042: // CertDuplicateStore
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCERTSTORE, CertCloseStore),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;

        pv = CertDuplicateStore(pv);
        TWAPI_ASSERT(pv);
        TwapiRegisterHCERTSTORE(interp, pv);
        result.value.obj = ObjFromOpaque(pv, "HCERTSTORE");
        result.type = TRT_OBJ;
        break;

    case 10043: // PFXIsPFXBlob
        CHECK_NARGS(interp, objc, 1);
        result.type = TRT_BOOL;
        blob.pbData = ObjToByteArray(objv[0], &blob.cbData);
        result.value.bval = PFXIsPFXBlob(&blob);
        break;

    case 10044: // PFXVerifyPassword
        CHECK_NARGS(interp, objc, 2);
        pv = ObjDecryptUnicode(interp, objv[1], &dw);
        if (pv == NULL)
            return TCL_ERROR;
        result.type = TRT_BOOL;
        blob.pbData = ObjToByteArray(objv[0], &blob.cbData);
        result.value.bval = PFXVerifyPassword(&blob, pv, 0);
        SecureZeroMemory(pv, sizeof(WCHAR) * dw);
        TwapiFree(pv);
        break;

    case 10045: // Twapi_CertStoreSerialize
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(hstore, HCERTSTORE, CertCloseStore),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        blob.pbData = NULL;
        blob.cbData = 0;
        if (! CertSaveStore(hstore, PKCS_7_ASN_ENCODING|X509_ASN_ENCODING,
                            dw, CERT_STORE_SAVE_TO_MEMORY, &blob, 0))
            return TwapiReturnSystemError(interp);
        result.value.obj = ObjFromByteArray(NULL, blob.cbData);
        blob.pbData = ObjToByteArray(result.value.obj, &blob.cbData);
        if (! CertSaveStore(hstore, PKCS_7_ASN_ENCODING | X509_ASN_ENCODING,
                            dw, CERT_STORE_SAVE_TO_MEMORY, &blob, 0)) {
            TwapiReturnSystemError(interp);
            ObjDecrRefs(result.value.obj);
            return TCL_ERROR;
        }
        Tcl_SetByteArrayLength(result.value.obj, blob.cbData);
        result.type = TRT_OBJ;
        break;

    case 10046: // CryptStringToBinary
        if (TwapiGetArgs(interp, objc, objv, GETOBJ(s1Obj),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        s1 = ObjToUnicodeN(s1Obj, &dw2);
        if (!CryptStringToBinaryW(s1, dw2, dw, NULL, &dw3, NULL, NULL))
            return TwapiReturnSystemError(interp);
        result.value.obj = ObjFromByteArray(NULL, dw3);
        pv = ObjToByteArray(result.value.obj, &dw3);
        if (!CryptStringToBinaryW(s1, dw2, dw, pv, &dw3, NULL, NULL)) {
            TwapiReturnSystemError(interp);
            ObjDecrRefs(result.value.obj);
        }
        Tcl_SetByteArrayLength(result.value.obj, dw3);
        result.type = TRT_OBJ;
        break;
    case 10047: // CryptBinaryToString
        if (TwapiGetArgs(interp, objc, objv, GETOBJ(s1Obj),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        pv = ObjToByteArray(s1Obj, &dw2);
        if (!CryptBinaryToStringW(pv, dw2, dw, NULL, &dw3))
            return TwapiReturnSystemError(interp);
        result.value.unicode.str = TwapiAlloc(sizeof(WCHAR)*dw3);
        if (!CryptBinaryToStringW(pv, dw2, dw, result.value.unicode.str, &dw3)) {
            TwapiReturnSystemError(interp);
            TwapiFree(result.value.unicode.str);
            return TCL_ERROR;
        }        
        result.value.unicode.len = dw3;
        result.type = TRT_UNICODE_DYNAMIC;
        break;
    case 10048: // CryptFindLocalizedName
        CHECK_NARGS(interp, objc, 1);
        result.value.unicode.str = (WCHAR *)CryptFindLocalizedName(ObjToUnicode(objv[0]));
        /* Note returned string is STATIC RESOURCE and must NOT be dealloced*/
        if (result.value.unicode.str) {
            result.value.unicode.len = -1;
            result.type = TRT_UNICODE;
        } else {
            result.value.obj = objv[0];
            result.type = TRT_OBJ;
        }
        break;
    case 10049: // CertCompareCertificateName
        // TBD - is this any different than doing a string compare in script?
        CHECK_NARGS(interp, objc, 3);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);
        blob.pbData = ObjToByteArray(objv[1], &blob.cbData);
        blob2.pbData = ObjToByteArray(objv[2], &blob2.cbData);
        result.type = TRT_BOOL;
        result.value.bval = CertCompareCertificateName(dw, &blob, &blob2);
        break;
    case 10050: // CryptEnumProviderTypes
    case 10051: // CryptEnumProvider
        CHECK_NARGS(interp, objc, 1);
        CHECK_INTEGER_OBJ(interp, dw, objv[0]);

        /* XP SP3 has a bug where the Unicode version of the
           CryptEnumProviderTypes always returns ERROR_MORE_DATA - see
           http://support.microsoft.com/kb/959160
           We special case this to use the ANSI version.
        */
        if (func == 10050 && ! TwapiMinOSVersion(6, 0)) {
            dw2 = sizeof(buf.ansi);   /* sizeof, NOT ARRAYSIZE in all cases */
            if (CryptEnumProviderTypesA(dw, NULL, 0, &dw3, buf.ansi, &dw2)) {
                objs[0] = ObjFromDWORD(dw3);
                objs[1] = ObjFromStringLimited(buf.ansi, dw2, NULL);
                result.type= TRT_OBJV;
                result.value.objv.objPP = objs;
                result.value.objv.nobj = 2;
                break;
            }
        } else {
            dw2 = sizeof(buf.uni);   /* sizeof, NOT ARRAYSIZE in all cases */
            if ((func == 10050 ? CryptEnumProviderTypesW : CryptEnumProvidersW)(dw, NULL, 0, &dw3, buf.uni, &dw2)) {
                objs[0] = ObjFromDWORD(dw3);
                objs[1] = ObjFromUnicodeLimited(buf.uni, dw2/sizeof(WCHAR), NULL);
                result.type= TRT_OBJV;
                result.value.objv.objPP = objs;
                result.value.objv.nobj = 2;
                break;
            }
        }
        /* Error handling */
        result.value.ival = GetLastError();
        if (result.value.ival == ERROR_NO_MORE_ITEMS)
            result.type = TRT_EMPTY;
        else
            result.type = TRT_EXCEPTION_ON_ERROR;
        break;

    case 10052: // CryptMsgClose
        CHECK_NARGS(interp, objc, 1);
        return TwapiCloseContext(interp, objv[0], "HCRYPTMSG",
                                 TwapiUnregisterHCRYPTMSG, CryptMsgClose);
    case 10053: // CertFreeCRLContext
        CHECK_NARGS(interp, objc, 1);
        return TwapiCloseContext(interp, objv[0], "PCCRL_CONTEXT",
                                 TwapiUnregisterPCCRL_CONTEXT, CertFreeCRLContext);
    case 10054: // CertFreeCTLContext
        CHECK_NARGS(interp, objc, 1);
        return TwapiCloseContext(interp, objv[0], "PCCTL_CONTEXT",
                                 TwapiUnregisterPCCTL_CONTEXT, CertFreeCTLContext);
        
    case 10055: // CryptCATAdminCalcHashFromFileHandle
        CHECK_NARGS(interp, objc, 1);
        if (ObjToHANDLE(interp, objv[0], &h) != TCL_OK)
            return TCL_ERROR;
        dw = sizeof(buf.ansi);
        if (CryptCATAdminCalcHashFromFileHandle(h, &dw, buf.ansi, 0) == FALSE)
            result.type = TRT_GETLASTERROR;
        else 
            return ObjSetResult(interp, ObjFromByteArray(buf.ansi, dw));
         break;
    case 10056: // CryptCATAdminAcquireContext 
        CHECK_NARGS(interp, objc, 1);
        if (ObjToGUID(interp, objv[0], &buf.guid) != TCL_OK)
            return TCL_ERROR;
        if (CryptCATAdminAcquireContext(&h, &buf.guid, 0) == FALSE)
            result.type = TRT_GETLASTERROR;
        else {
            TwapiRegisterHCATADMIN(interp, h);
            TwapiResult_SET_NONNULL_PTR(result, HCATADMIN, h);
        }
        break;

    case 10057: // CryptCATAdminReleaseContext
        CHECK_NARGS(interp, objc, 1);
        return TwapiCloseContext(interp, objv[0], "HCATADMIN",
                                 TwapiUnregisterHCATADMIN, TwapiCryptCATAdminReleaseContext);
        CHECK_NARGS(interp, objc, 1);
        if (ObjToOpaque(interp, objv[0], &h, "HCATADMIN") != TCL_OK)
            return TCL_ERROR;
        if (CryptCATAdminReleaseContext(h, 0) == FALSE)
            result.type = TRT_GETLASTERROR;
        else
            result.type = TRT_EMPTY;
        break;

    case 10058: // CryptCATAdminReleaseCatalogContext
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(h, HCATADMIN, CryptCATAdminReleaseContext),
                         GETVERIFIEDPTR(h2, HCATINFO, CryptCATAdminReleaseCatalogContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        TwapiUnregisterHCATINFO(interp, h2);
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CryptCATAdminReleaseCatalogContext(h, h2, dw);
        break;

    case 10059: // CryptCATCatalogInfoFromContext
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(h, HCATINFO, CryptCATAdminReleaseCatalogContext),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        buf.catinfo.cbStruct = sizeof(buf.catinfo);
        if (CryptCATCatalogInfoFromContext(h, &buf.catinfo, 0) == FALSE)
            result.type = TRT_GETLASTERROR;
        else {
            result.type = TRT_UNICODE;
            result.value.unicode.str = buf.catinfo.wszCatalogFile;
            result.value.unicode.len = -1;
        }
        break;
        
#ifdef TBD
    case TBD: // CertCreateContext
        if (TwapiGetArgs(interp, objc, objv,
                         GETINT(dw), GETINT(dw2), GETOBJ(s1Obj),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (dw != CERT_STORE_CERTIFICATE_CONTEXT) {
            return TwapiReturnError(interp, TWAPI_UNSUPPORTED_TYPE);
        }
        pv = ObjToByteArray(s1Obj, &dw3);
        certP = CertCreateContext(dw, dw2, pv, dw3, 0, NULL);
        if (certP == NULL)
            return TwapiReturnSystemError(interp);
        TwapiRegisterPCCERT_CONTEXT(interp, certP);
        TwapiResult_SET_NONNULL_PTR(result, PCCERT_CONTEXT, (void*)certP);
        break;
#endif

    }

    return TwapiSetResult(interp, &result);
}

/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseCRYPTPROTECT_PROMPTSTRUCT(TwapiInterpContext *ticP, Tcl_Obj *promptObj, CRYPTPROTECT_PROMPTSTRUCT *promptP)
{
    Tcl_Obj **objs;
    int nobjs;
    TCL_RESULT res;

    res = ObjGetElements(NULL, promptObj, &nobjs, &objs);
    if (res == TCL_OK) {
        promptP->cbSize = sizeof(*promptP);
        if (nobjs == 0) {
            promptP->dwPromptFlags = 0;
            promptP->hwndApp = NULL;
            promptP->szPrompt = NULL;
        } else {
            res = TwapiGetArgsEx(ticP, nobjs, objs,
                                 GETINT(promptP->dwPromptFlags),
                                 GETHWND(promptP->hwndApp),
                                 GETWSTR(promptP->szPrompt), ARGEND);
            }
    }
    return res;
}

static int Twapi_CryptProtectObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD flags;
    CRYPT_DATA_BLOB inblob, outblob;
    TCL_RESULT res;
    Tcl_Obj *inObj, *promptObj;
    LPCWSTR description;
    CRYPTPROTECT_PROMPTSTRUCT prompt;
    MemLifoMarkHandle mark;
    
    mark = MemLifoPushMark(ticP->memlifoP);
    
    /* We do not want to make a copy of the input data for performance
       reasons so we do not extract it directly using TwapiGetArgsEx. Other
       parameters are extracted using that function so we are safe
       in directly accessing the input data blob without fear of
       the Tcl_Obj shimmering underneath us
    */
    res = TwapiGetArgsEx(ticP, objc-1, objv+1, GETOBJ(inObj),
                         GETEMPTYASNULL(description), ARGSKIP,
                         ARGSKIP, GETOBJ(promptObj),
                         GETINT(flags), ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    res = ParseCRYPTPROTECT_PROMPTSTRUCT(ticP, promptObj, &prompt);
    if (res != TCL_OK)
        goto vamoose;

    inblob.pbData = ObjToByteArray(inObj, &inblob.cbData);
    outblob.pbData = NULL;
    outblob.cbData = 0;
    if (CryptProtectData(
            &inblob,
            description,        /* May be NULL */
            NULL,               /* Entropy, not supported */
            NULL,               /* Reserved - should be NULL */
            prompt.szPrompt ? &prompt : NULL,
            flags,
            &outblob)) {
        if (outblob.pbData) {
            ObjSetResult(interp, ObjFromByteArray(outblob.pbData, outblob.cbData));
            LocalFree(outblob.pbData);
        } /* else empty result. Should not happen ? */
    }
    else
        res = TwapiReturnSystemError(interp);

vamoose:
    MemLifoPopMark(mark);
    return res;
}


static int Twapi_CryptUnprotectObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    DWORD flags;
    CRYPT_DATA_BLOB inblob, outblob;
    Tcl_Obj *inObj, *promptObj;
    LPWSTR description;
    CRYPTPROTECT_PROMPTSTRUCT prompt;
    MemLifoMarkHandle mark;
    TCL_RESULT res;
    
    mark = MemLifoPushMark(ticP->memlifoP);
    
    /* We do not want to make a copy of the input data for performance
       reasons so we do not extract it directly using TwapiGetArgsEx. Other
       parameters are extracted using that function so we are safe
       in directly accessing the input data blob without fear of
       the Tcl_Obj shimmering underneath us
    */
    res = TwapiGetArgsEx(ticP, objc-1, objv+1, GETOBJ(inObj),
                         ARGSKIP, ARGSKIP, GETOBJ(promptObj),
                         GETINT(flags), ARGEND);
    if (res != TCL_OK)
        goto vamoose;

    res = ParseCRYPTPROTECT_PROMPTSTRUCT(ticP, promptObj, &prompt);
    if (res != TCL_OK)
        goto vamoose;

    inblob.pbData = ObjToByteArray(inObj, &inblob.cbData);
    outblob.pbData = NULL;
    outblob.cbData = 0;
    description = NULL;
    if (CryptUnprotectData(
            &inblob,
            &description,
            NULL,
            NULL,
            prompt.szPrompt ? &prompt : NULL,
            flags,
            &outblob)) {
        Tcl_Obj *objs[2];
        objs[0] = ObjFromByteArray(outblob.pbData, outblob.cbData);
        objs[1] = description ? ObjFromUnicode(description) : ObjFromEmptyString();
        ObjSetResult(interp, ObjNewList(2, objs));
        if (description)
            LocalFree(description);
        if (outblob.pbData)
            LocalFree(outblob.pbData);
    }
    else
        res = TwapiReturnSystemError(interp);

vamoose:
    MemLifoPopMark(mark);
    return res;
}

/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseWINTRUST_DATA(TwapiInterpContext *ticP, Tcl_Obj *objP, TWAPI_WINTRUST_DATA *wtdP)
{
    Tcl_Interp *interp = ticP->interp;
    TCL_RESULT ret;
    Tcl_Obj *trustObj;
    WINTRUST_FILE_INFO *wfiP;
    WINTRUST_CATALOG_INFO *wciP;

    ZeroMemory(wtdP, sizeof(*wtdP));
    wtdP->cbStruct = sizeof(*wtdP);
    ret = TwapiGetArgsExObj(ticP, objP,
                            ARGSKIP, // pPolicyCallbackData
                            ARGSKIP, // pSIPClientData
                            GETINT(wtdP->dwUIChoice),
                            GETINT(wtdP->fdwRevocationChecks),
                            GETINT(wtdP->dwUnionChoice), GETOBJ(trustObj),
                            GETINT(wtdP->dwStateAction),
                            GETHANDLET(wtdP->hWVTStateData, WVTStateData),
                            ARGSKIP, // GETWSTR(wtdP->pwszURLReference),
                            GETINT(wtdP->dwProvFlags),
                            GETINT(wtdP->dwUIContext),
                         // pSignatureSettings not present until Win8
                            ARGEND);
    if (ret != TCL_OK)
        return ret;
    switch (wtdP->dwUnionChoice) {
    case WTD_CHOICE_FILE:
        wfiP = MemLifoZeroes(ticP->memlifoP, sizeof(*wfiP));
        wtdP->pFile = wfiP;
        wfiP->cbStruct = sizeof(*wfiP);
        ret = TwapiGetArgsExObj(ticP, trustObj, GETWSTR(wfiP->pcwszFilePath),
                                ARGUSEDEFAULT,
                                GETHANDLE(wfiP->hFile),
                                GETVARWITHDEFAULT(wfiP->pgKnownSubject, ObjToGUID_NULL),
                                ARGEND);
        break;
    case WTD_CHOICE_CATALOG:
        wciP = MemLifoZeroes(ticP->memlifoP, sizeof(*wciP));
        wtdP->pCatalog = wciP;
        wciP->cbStruct = sizeof(*wciP);
        ret = TwapiGetArgsExObj(ticP, trustObj,
                                GETINT(wciP->dwCatalogVersion),
                                GETWSTR(wciP->pcwszCatalogFilePath),
                                GETWSTR(wciP->pcwszMemberTag),
                                GETWSTR(wciP->pcwszMemberFilePath),
                                GETHANDLE(wciP->hMemberFile),
                                GETBA(wciP->pbCalculatedFileHash, wciP->cbCalculatedFileHash),
                                // pcCatalogContext not yet implemented
                                // hCatAdmin needs Win8 or later
                                ARGEND);
        
        if (ret == TCL_OK)
            wciP->pcCatalogContext = NULL;
        
        break;
    default:
        ret = TwapiReturnErrorMsg(ticP->interp, TWAPI_UNSUPPORTED_TYPE, "Unsupported Wintrust type");
        break;
    }

    return ret;
}

static int Twapi_WinVerifyTrustObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TWAPI_WINTRUST_DATA wtd;
    Tcl_Obj *wtdObj;
    HWND hwnd;
    GUID guid;
    TCL_RESULT ret;
    LONG status;
    MemLifoMarkHandle mark;
    
    ret = TwapiGetArgs(interp, objc-1, objv+1, GETHWND(hwnd), GETGUID(guid),
                       GETOBJ(wtdObj), ARGEND);
    if (ret != TCL_OK)
        return ret;
    mark = MemLifoPushMark(ticP->memlifoP);
    ret = ParseWINTRUST_DATA(ticP, wtdObj, &wtd);
    if (ret == TCL_OK) {
        status = WinVerifyTrust(hwnd, &guid, (WINTRUST_DATA *)&wtd);

        wtd.dwStateAction = WTD_STATEACTION_CLOSE;
        wtd.dwUIChoice = WTD_UI_NONE;
        WinVerifyTrust((HWND) INVALID_HANDLE_VALUE, &guid, (WINTRUST_DATA *)&wtd);
        ObjSetResult(interp, ObjFromInt(status));
    }
    MemLifoPopMark(mark);
    return ret;
}

static int TwapiCryptoInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s CryptoDispatch[] = {
        DEFINE_FNCODE_CMD(CryptAcquireContext, 10000),
        DEFINE_FNCODE_CMD(CryptReleaseContext, 10001),
        DEFINE_FNCODE_CMD(CryptGetProvParam, 10002),
        DEFINE_FNCODE_CMD(CertOpenSystemStore, 10003),
        DEFINE_FNCODE_CMD(cert_delete_from_store, 10004),
        DEFINE_FNCODE_CMD(CertCreateCertificateContext, 10005),
        DEFINE_FNCODE_CMD(CertEnumCertificatesInStore, 10006),
        DEFINE_FNCODE_CMD(CertEnumCertificateContextProperties, 10007),
        DEFINE_FNCODE_CMD(CertGetCertificateContextProperty, 10008),
        DEFINE_FNCODE_CMD(crypt_key_free, 10009),
        DEFINE_FNCODE_CMD(CryptGenKey, 10010),
        DEFINE_FNCODE_CMD(CertStrToName, 10011),
        DEFINE_FNCODE_CMD(CertNameToStr, 10012),
        DEFINE_FNCODE_CMD(CertGetNameString, 10013),
        DEFINE_FNCODE_CMD(cert_release, 10014),
        DEFINE_FNCODE_CMD(Twapi_CertGetEncoded, 10015),
        DEFINE_FNCODE_CMD(CertUnregisterSystemStore, 10016),
        DEFINE_FNCODE_CMD(CertCloseStore, 10017),
        DEFINE_FNCODE_CMD(CryptGetUserKey, 10018),
        DEFINE_FNCODE_CMD(CryptSetProvParam, 10019),
        DEFINE_FNCODE_CMD(CertOpenStore, 10020),
        DEFINE_FNCODE_CMD(CryptEnumOIDInfo, 10021),
        DEFINE_FNCODE_CMD(CertAddCertificateContextToStore, 10022),
        DEFINE_FNCODE_CMD(CryptExportPublicKeyInfoEx, 10023),
        DEFINE_FNCODE_CMD(CertEnumSystemStore, 10024),
        DEFINE_FNCODE_CMD(CertEnumPhysicalStore, 10025),
        DEFINE_FNCODE_CMD(CertEnumSystemStoreLocation, 10026),
        DEFINE_FNCODE_CMD(CryptAcquireCertificatePrivateKey, 10027), // Tcl
        DEFINE_FNCODE_CMD(CertGetEnhancedKeyUsage, 10028),
        DEFINE_FNCODE_CMD(Twapi_CertStoreCommit, 10029),
        DEFINE_FNCODE_CMD(Twapi_CertGetIntendedKeyUsage, 10030),
        DEFINE_FNCODE_CMD(CertGetIssuerCertificateFromStore, 10031),
        DEFINE_FNCODE_CMD(CertFreeCertificateChain, 10032),
        DEFINE_FNCODE_CMD(CertFindExtension, 10033),
        DEFINE_FNCODE_CMD(CryptGenRandom, 10034),
        DEFINE_FNCODE_CMD(Twapi_CertGetInfo, 10035),
        DEFINE_FNCODE_CMD(Twapi_CertGetExtensions, 10036),
        DEFINE_FNCODE_CMD(CryptFindCertificateKeyProvInfo, 10037),
        DEFINE_FNCODE_CMD(CertAddEncodedCertificateToStore, 10038),
        DEFINE_FNCODE_CMD(CertOIDToAlgId, 10039),
        DEFINE_FNCODE_CMD(CertAlgIdToOID, 10040),
        DEFINE_FNCODE_CMD(cert_duplicate, 10041),
        DEFINE_FNCODE_CMD(cert_store_duplicate, 10042), // TBD - document
        DEFINE_FNCODE_CMD(PFXIsPFXBlob, 10043), //TBD - document
        DEFINE_FNCODE_CMD(PFXVerifyPassword, 10044), // TBD - document
        DEFINE_FNCODE_CMD(Twapi_CertStoreSerialize, 10045),
        DEFINE_FNCODE_CMD(CryptStringToBinary, 10046), // Tcl TBD
        DEFINE_FNCODE_CMD(CryptBinaryToString, 10047), // Tcl TBD
        DEFINE_FNCODE_CMD(crypt_localize_string,10048), // TBD - document
        DEFINE_FNCODE_CMD(CertCompareCertificateName, 10049), // TBD Tcl
        DEFINE_FNCODE_CMD(CryptEnumProviderTypes, 10050),
        DEFINE_FNCODE_CMD(CryptEnumProviders, 10051),
        DEFINE_FNCODE_CMD(CryptMsgClose, 10052),
        DEFINE_FNCODE_CMD(CertFreeCRLContext, 10053),
        DEFINE_FNCODE_CMD(CertFreeCTLContext, 10054),
        DEFINE_FNCODE_CMD(CryptCATAdminCalcHashFromFileHandle, 10055),
        // TBD DEFINE_FNCODE_CMD(CertCreateContext, TBD),
        DEFINE_FNCODE_CMD(CryptCATAdminAcquireContext, 10056),
        DEFINE_FNCODE_CMD(CryptCATAdminReleaseContext, 10057),
        DEFINE_FNCODE_CMD(CryptCATAdminReleaseCatalogContext, 10058),
        DEFINE_FNCODE_CMD(CryptCATCatalogInfoFromContext, 10059),
    };

    static struct tcl_dispatch_s TclDispatch[] = {
        DEFINE_TCL_CMD(CertCreateSelfSignCertificate, Twapi_CertCreateSelfSignCertificate),
        DEFINE_TCL_CMD(CryptSignAndEncodeCertificate, Twapi_CryptSignAndEncodeCertObjCmd),
        DEFINE_TCL_CMD(CertFindCertificateInStore, Twapi_CertFindCertificateInStoreObjCmd),
        DEFINE_TCL_CMD(CertGetCertificateChain, Twapi_CertGetCertificateChainObjCmd),
        DEFINE_TCL_CMD(Twapi_CertVerifyChainPolicy, Twapi_CertVerifyChainPolicyObjCmd),
        DEFINE_TCL_CMD(Twapi_CertVerifyChainPolicySSL, Twapi_CertVerifyChainPolicySSLObjCmd),
        DEFINE_TCL_CMD(Twapi_HashPublicKeyInfo, Twapi_HashPublicKeyInfoObjCmd),
        DEFINE_TCL_CMD(CryptFindOIDInfo, Twapi_CryptFindOIDInfoObjCmd),
        DEFINE_TCL_CMD(CryptDecodeObjectEx, Twapi_CryptDecodeObjectExObjCmd), // Tcl
        DEFINE_TCL_CMD(CryptEncodeObjectEx, Twapi_CryptEncodeObjectExObjCmd),
        DEFINE_TCL_CMD(CryptFormatObject, Twapi_CryptFormatObjectObjCmd), // Tcl
        DEFINE_TCL_CMD(CertSetCertificateContextProperty, Twapi_CertSetCertificateContextPropertyObjCmd),
        DEFINE_TCL_CMD(Twapi_CertChainContexts, Twapi_CertChainContextsObjCmd),
        DEFINE_TCL_CMD(CryptProtectData, Twapi_CryptProtectObjCmd),
        DEFINE_TCL_CMD(CryptUnprotectData, Twapi_CryptUnprotectObjCmd),
        DEFINE_TCL_CMD(PFXExportCertStoreEx, Twapi_PFXExportCertStoreExObjCmd),
        DEFINE_TCL_CMD(PFXImportCertStore, Twapi_PFXImportCertStoreObjCmd),
        DEFINE_TCL_CMD(CryptQueryObject, Twapi_CryptQueryObjectObjCmd),
        DEFINE_TCL_CMD(WinVerifyTrust, Twapi_WinVerifyTrustObjCmd),
        DEFINE_TCL_CMD(CryptCATAdminEnumCatalogFromHash, Twapi_CryptCATAdminEnumCatalogFromHash),
        DEFINE_TCL_CMD(CryptGetKeyParam, Twapi_CryptGetKeyParam), // TBD - Tcl
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CryptoDispatch), CryptoDispatch, Twapi_CryptoCallObjCmd);
    TwapiDefineTclCmds(interp, ARRAYSIZE(TclDispatch), TclDispatch, ticP);

    return TwapiSspiInitCalls(interp, ticP);
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
int Twapi_crypto_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiCryptoInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

