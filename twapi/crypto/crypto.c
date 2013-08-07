/* 
 * Copyright (c) 2007-2009 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to CryptoAPI */

#include "twapi.h"
#include "twapi_crypto.h"

#ifndef TWAPI_SINGLE_MODULE
HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif


static int Twapi_CertCreateSelfSignCertificate(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
static TCL_RESULT TwapiCryptEncodeObject(Tcl_Interp *interp, MemLifo *lifoP,
                                         Tcl_Obj *oidObj, Tcl_Obj *valObj,
                                         CRYPT_OBJID_BLOB *blobP);
static TwapiCertGetNameString(
    Tcl_Interp *interp,
    PCCERT_CONTEXT certP,
    DWORD type,
    DWORD flags,
    Tcl_Obj *owhat);


#ifdef NOTNEEDED
/* RtlGenRandom in base provides this */
int Twapi_CryptGenRandom(Tcl_Interp *interp, HCRYPTPROV provH, DWORD len)
{
    BYTE buf[256];

    if (len > sizeof(buf)) {
        Tcl_SetObjErrorCode(interp,
                            Twapi_MakeTwapiErrorCodeObj(TWAPI_INTERNAL_LIMIT));
        Tcl_SetResult(interp, "Too many random bytes requested.", TCL_STATIC);
        return TCL_ERROR;
    }

    if (CryptGenRandom(provH, len, buf)) {
        TwapiSetObjResult(interp, ObjFromByteArray(buf, len));
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
}
#endif

/* Note: Allocates memory for blobP from lifoP. Note structure internal
   pointers may point to Tcl_Obj areas within valObj so
   TREAT RETURNED STRUCTURES AS VOLATILE.
*/
static TCL_RESULT TwapiCryptEncodeObject(Tcl_Interp *interp, MemLifo *lifoP,
                                  Tcl_Obj *oidObj, Tcl_Obj *valObj,
                                  CRYPT_OBJID_BLOB *blobP)
{
    LPCSTR    soid;
    DWORD     dw;
    Tcl_Obj **objs;
    int       nobjs;
    int       status;
    void     *penc;
    int       nenc;
    union {
        void *pv;
        CERT_ALT_NAME_ENTRY  *altnameP;
    } p;

    /* Note: X509_ALTERNATE_NAME etc. are integer values cast as LPSTR in
       headers. Hence all the casting around soid. Ugh and Yuck */

    /* The oidobj may be specified as either a string or an integer */
    if (ObjToDWORD(NULL, oidObj, &dw) == TCL_OK && dw < 65536) {
        soid = (LPSTR) (DWORD_PTR) dw;
    } else {
        soid = ObjToString(oidObj);
        if (STREQ(soid, szOID_SUBJECT_ALT_NAME) ||
            STREQ(soid, szOID_ISSUER_ALT_NAME)) {
            soid = X509_ALTERNATE_NAME; /* soid NOW A DWORD!!! */
        } else {
            Tcl_SetObjResult(interp, Tcl_ObjPrintf("Unsupported OID \"%s\"",soid));
            return TCL_ERROR;
        }
    }

    switch ((DWORD_PTR)soid) {
    case (DWORD_PTR) X509_ALTERNATE_NAME:
        p.altnameP = MemLifoAlloc(lifoP, sizeof(*p.altnameP), NULL);
        if ((status = ObjGetElements(interp, valObj, &nobjs, &objs)) != TCL_OK)
            return status;
        if (nobjs != 2 ||
            ObjToDWORD(NULL, objs[0], &p.altnameP->dwAltNameChoice) != TCL_OK)
            goto invalid_name_error;
        switch (p.altnameP->dwAltNameChoice) {
        case CERT_ALT_NAME_RFC822_NAME: /* FALLTHROUGH */
        case CERT_ALT_NAME_DNS_NAME: /* FALLTHROUGH */
        case CERT_ALT_NAME_URL:
            p.altnameP->pwszRfc822Name = ObjToUnicode(objs[1]);
            break;
        case CERT_ALT_NAME_REGISTERED_ID:
            p.altnameP->pszRegisteredID = ObjToString(objs[1]);
            break;
        case CERT_ALT_NAME_OTHER_NAME: /* FALLTHRU */
        case CERT_ALT_NAME_DIRECTORY_NAME: /* FALLTHRU */
        case CERT_ALT_NAME_IP_ADDRESS: /* FALLTHRU */
        default:
            goto invalid_name_error;
        }
        break;

    default:
        Tcl_SetObjResult(interp, Tcl_ObjPrintf("Unsupported OID constant \"%d\"", (DWORD_PTR) soid));
        return TCL_ERROR;
    }

    /* Assume 256 bytes enough but get as much as we can */
    penc = MemLifoAlloc(lifoP, 256, &nenc);
    if (CryptEncodeObjectEx(PKCS_7_ASN_ENCODING|X509_ASN_ENCODING,
                            soid, /* Yuck */
                            p.pv, 0, NULL, penc, &nenc) == 0) {
        if (GetLastError() != ERROR_MORE_DATA)
            return TwapiReturnSystemError(interp);
        /* Retry with specified buffer size */
        penc = MemLifoAlloc(lifoP, nenc, &nenc);
        if (CryptEncodeObjectEx(PKCS_7_ASN_ENCODING|X509_ASN_ENCODING,
                                soid, /* Yuck */
                                p.pv, 0, NULL, penc, &nenc) == 0)
            return TwapiReturnSystemError(interp);
    }
    
    blobP->cbData = nenc;
    blobP->pbData = penc;

    /* Note caller has to MemLifoPopFrame to release lifo memory */
    return TCL_OK;

invalid_name_error:
    Tcl_SetObjResult(interp,
                     Tcl_ObjPrintf("Invalid or unsupported name format \"%s\"", ObjToString(valObj)));
    return TCL_ERROR;
}

static int Twapi_CertCreateSelfSignCertificate(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    void *pv;
    HCRYPTPROV hprov;
    DWORD flags;
    int status;
    CERT_NAME_BLOB name_blob;
    CRYPT_KEY_PROV_INFO ki, *kiP;
    CRYPT_ALGORITHM_IDENTIFIER algid, *algidP;
    Tcl_Obj **objs;
    int       nobjs;
    SYSTEMTIME start, end, *startP, *endP;
    PCERT_CONTEXT certP;
    CERT_EXTENSIONS exts, *extsP;

    if ((status = TwapiGetArgs(interp, objc-1, objv+1,
                               GETPTR(pv, HCRYPTPROV),
                               GETBIN(name_blob.pbData, name_blob.cbData),
                               GETINT(flags),
                               ARGSKIP, // CRYPT_KEY_PROV_INFO
                               ARGSKIP, // CRYPT_ALGORITHM_IDENTIFIER
                               ARGSKIP, // STARTTIME
                               ARGSKIP, // ENDTIME
                               ARGSKIP, // EXTENSIONS
                               ARGEND)) != TCL_OK)
        return status;
    
    if (pv && (status = TwapiVerifyPointer(interp, pv, CryptReleaseContext)) != TCL_OK)
        return status;
    hprov = (HCRYPTPROV) pv;

    /* Parse CRYPT_KEY_PROV_INFO */
    if ((status = ObjGetElements(interp, objv[4], &nobjs, &objs)) != TCL_OK)
        return status;
    if (nobjs == 0)
        kiP = NULL;
    else {
        if (TwapiGetArgs(interp, nobjs, objs,
                         GETWSTR(ki.pwszContainerName),
                         GETWSTR(ki.pwszProvName),
                         GETINT(ki.dwProvType),
                         GETINT(ki.dwFlags),
                         GETINT(ki.cProvParam),
                         ARGSKIP,
                         GETINT(ki.dwKeySpec),
                         ARGEND) != TCL_OK
            ||
            ki.cProvParam != 0) {
            Tcl_SetResult(interp, "Invalid or unimplemented provider parameters", TCL_STATIC);
            return TCL_ERROR;
        }
        ki.rgProvParam = NULL;
        kiP = &ki;
    }

    /* Parse CRYPT_ALGORITHM_IDENTIFIER */
    if ((status = ObjGetElements(interp, objv[5], &nobjs, &objs)) != TCL_OK)
        return status;
    if (nobjs == 0)
        algidP = NULL;
    else {
        if (nobjs > 2) {
            Tcl_SetResult(interp, "Invalid algorithm identifier format", TCL_STATIC);
            return TCL_ERROR;
        } else if (nobjs == 2) {
            Tcl_SetResult(interp, "Algorithm identifier formats with parameters are not implemented", TCL_STATIC);
            return TCL_ERROR;
        }
        algid.pszObjId = ObjToString(objs[0]);
        algid.Parameters.cbData = 0;
        algid.Parameters.pbData = 0;
        algidP = &algid;
    }

    if ((status = ObjGetElements(interp, objv[6], &nobjs, &objs)) != TCL_OK)
        return status;
    if (nobjs == 0)
        startP = NULL;
    else {
        if ((status = ObjToSYSTEMTIME(interp, objv[6], &start)) != TCL_OK)
            return status;
        startP = &start;
    }

    if ((status = ObjGetElements(interp, objv[7], &nobjs, &objs)) != TCL_OK)
        return status;
    if (nobjs == 0)
        endP = NULL;
    else {
        if ((status = ObjToSYSTEMTIME(interp, objv[7], &end)) != TCL_OK)
            return status;
        endP = &end;
    }

    if ((status = ObjGetElements(interp, objv[8], &nobjs, &objs)) != TCL_OK)
        return status;
    if (nobjs == 0)
        extsP = NULL;
    else {
        DWORD i;

        exts.rgExtension = MemLifoPushFrame(
            &ticP->memlifo, nobjs * sizeof(CERT_EXTENSION), NULL);
        exts.cExtension = nobjs;

        for (i = 0; i < exts.cExtension; ++i) {
            Tcl_Obj **extobjs;
            int       nextobjs;
            int       bval;
            PCERT_EXTENSION extP = &exts.rgExtension[i];

            status = ObjGetElements(interp, objs[i], &nextobjs, &extobjs);
            if (status == TCL_OK) {
                if (nextobjs == 2 || nextobjs == 3) {
                    status = ObjToBoolean(interp, extobjs[1], &bval);
                    if (status == TCL_OK) {
                        extP->pszObjId = ObjToString(extobjs[0]);
                        extP->fCritical = (BOOL) bval;
                        if (nextobjs == 3) {
                            status = TwapiCryptEncodeObject(
                                interp, &ticP->memlifo,
                                extobjs[0], extobjs[2],
                                &extP->Value);
                        } else {
                            extP->Value.cbData = 0;
                            extP->Value.pbData = NULL;
                        }
                    }
                } else {
                    Tcl_SetResult(interp, "Certificate extension format invalid or not implemented", TCL_STATIC);
                    status = TCL_ERROR;
                }
            }

            if (status != TCL_OK) {
                MemLifoPopFrame(&ticP->memlifo);
                return status;
            }
        }
    }

    certP = (PCERT_CONTEXT) CertCreateSelfSignCertificate(hprov, &name_blob, flags,
                                          kiP, algidP, startP, endP, extsP);
    if (extsP && extsP->rgExtension) {
        MemLifoPopFrame(&ticP->memlifo);
    }

    if (certP) {
        if (TwapiRegisterPointer(interp, certP, CertFreeCertificateContext) != TCL_OK)
            Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
        Tcl_SetObjResult(interp, ObjFromOpaque(certP, "CERT_CONTEXT*"));
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
}

static int Twapi_CertGetCertificateContextProperty(Tcl_Interp *interp, PCCERT_CONTEXT certP, DWORD prop_id)
{
    BOOL status;
    DWORD n = 0;
    Tcl_Obj *objP;

    if (! CertGetCertificateContextProperty(certP, prop_id, NULL, &n))
        return TwapiReturnSystemError(interp);

    objP = ObjFromByteArray(NULL, n);
    status = CertGetCertificateContextProperty(certP, prop_id,
                                               ObjToByteArray(objP, &n),
                                               &n);
    if (!status) {
        TwapiReturnSystemError(interp);
        Tcl_DecrRefCount(objP);
        return TCL_ERROR;
    }

    Tcl_SetByteArrayLength(objP, n);
    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}

static TCL_RESULT Twapi_SetCertContextKeyProvInfo(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    PCCERT_CONTEXT certP;
    CRYPT_KEY_PROV_INFO ckpi;
    Tcl_Obj **objs;
    int       nobjs;
    TCL_RESULT status;

    /* Note - objc/objv have initial command name arg removed by caller */
    if ((status = TwapiGetArgs(interp, objc, objv,
                               GETVERIFIEDPTR(certP, CERT_CONTEXT*, CertFreeCertificateContext),
                               ARGSKIP, ARGEND)) != TCL_OK)
        return status;

    if ((status = ObjGetElements(interp, objv[1], &nobjs, &objs)) != TCL_OK)
        return status;

    if ((status = TwapiGetArgs(interp, nobjs, objs,
                               GETWSTR(ckpi.pwszContainerName),
                               GETWSTR(ckpi.pwszProvName),
                               GETINT(ckpi.dwProvType),
                               GETINT(ckpi.dwFlags),
                               ARGSKIP, // cProvParam+rgProvParam
                               GETINT(ckpi.dwKeySpec),
                               ARGEND)) != TCL_OK)
        return status;

    ckpi.cProvParam = 0;
    ckpi.rgProvParam = NULL;

    if (CertSetCertificateContextProperty(certP, CERT_KEY_PROV_INFO_PROP_ID,
                                          0, &ckpi))
        return TCL_OK;
    else
        return TwapiReturnSystemError(interp);
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
        Tcl_SetObjResult(interp, Tcl_ObjPrintf("CertGetNameString: unknown type %d", type));
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
        Tcl_SetObjResult(interp, Tcl_ObjPrintf("CertGetNameString: unsupported flags %d", flags));
        return TCL_ERROR;
    }

    nchars = CertGetNameStringW(certP, type, flags, pv, buf, ARRAYSIZE(buf));
    /* Note nchars includes terminating NULL */
    if (nchars > 1) {
        if (nchars < ARRAYSIZE(buf)) {
            Tcl_SetObjResult(interp, ObjFromUnicodeN(buf, nchars-1));
        } else {
            /* Buffer might have been truncated. Explicitly get buffer size */
            WCHAR *bufP;
            nchars = CertGetNameStringW(certP, type, flags, pv, NULL, 0);
            bufP = TwapiAlloc(nchars*sizeof(WCHAR));
            nchars = CertGetNameStringW(certP, type, flags, pv, bufP, nchars);
            Tcl_SetObjResult(interp, ObjFromUnicodeN(bufP, nchars-1));
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
            Tcl_DecrRefCount(objP);
            return Twapi_AppendSystemError(interp, n);
        }
        Tcl_SetObjResult(interp, objP);
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
            Tcl_DecrRefCount(objP);
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

    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}

static int Twapi_CryptoCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    DWORD dw, dw2;
    DWORD_PTR dwp;
    LPVOID pv;
    LPWSTR s1;
    HANDLE h;
    int func = PtrToInt(clientdata);
    struct _CRYPTOAPI_BLOB blob;
    PCCERT_CONTEXT certP;

    --objc;
    ++objv;

    TWAPI_ASSERT(sizeof(HCRYPTPROV) <= sizeof(pv));
    TWAPI_ASSERT(sizeof(HCRYPTKEY) <= sizeof(pv));
    TWAPI_ASSERT(sizeof(dwp) <= sizeof(void*));

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 10001: // CryptReleaseContext
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK)
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

    case 10003: // cert_open_system_store
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        h = CertOpenSystemStoreW(0, ObjToUnicode(objv[0]));
        /* CertCloseStore does not check ponter validity! So do ourselves*/
        if (TwapiRegisterPointer(interp, h, CertCloseStore) != TCL_OK)
            Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
        TwapiResult_SET_NONNULL_PTR(result, HCERTSTORE, h);
        break;

    case 10004: // CertDeleteCertificateFromStore
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, CERT_CONTEXT*, CertFreeCertificateContext), ARGEND) != TCL_OK)
            return TCL_ERROR;
        /* Unregister previous context since the next call will free it,
           EVEN ON FAILURES */
        if (TwapiUnregisterPointer(interp, certP, CertFreeCertificateContext) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CertDeleteCertificateFromStore(certP);
        break;

    case 10005: // Twapi_SetCertContextKeyProvInfo
        return Twapi_SetCertContextKeyProvInfo(interp, objc, objv);

    case 10006: // CertEnumCertificatesInStore
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(h, HCERTSTORE, CertCloseStore),
                         GETPTR(certP, CERT_CONTEXT*), ARGEND) != TCL_OK)
            return TCL_ERROR;
        /* Unregister previous context since the next call will free it */
        if (certP &&
            TwapiUnregisterPointer(interp, certP, CertFreeCertificateContext) != TCL_OK)
            return TCL_ERROR;
        certP = CertEnumCertificatesInStore(h, certP);
        if (certP) {
            if (TwapiRegisterPointer(interp, certP, CertFreeCertificateContext) != TCL_OK)
                Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
            TwapiResult_SET_NONNULL_PTR(result, CERT_CONTEXT*, (void*)certP);
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
                         GETVERIFIEDPTR(certP, CERT_CONTEXT*, CertFreeCertificateContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_DWORD;
        result.value.ival = CertEnumCertificateContextProperties(certP, dw);
        break;

    case 10008: // CertGetCertificateContextProperty
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, CERT_CONTEXT*, CertFreeCertificateContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_CertGetCertificateContextProperty(interp, certP, dw);

    case 10009: // CryptDestroyKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTKEY, CryptDestroyKey),
                         ARGEND) != TCL_OK)
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
            if (TwapiRegisterPointer(interp, (void*)dwp, CryptDestroyKey) != TCL_OK)
                Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
            TwapiResult_SET_PTR(result, HCRYPTKEY, (void*)dwp);
        } else
            result.type = TRT_GETLASTERROR;
        break;

    case 10011: // CertStrToName
        if (TwapiGetArgs(interp, objc, objv, GETWSTR(s1), ARGUSEDEFAULT,
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_GETLASTERROR;
        dw2 = 0;
        if (CertStrToNameW(X509_ASN_ENCODING, s1, dw, NULL, NULL, &dw2, NULL)) {
            result.value.obj = ObjFromByteArray(NULL, dw2);
            if (CertStrToNameW(X509_ASN_ENCODING, s1, dw, NULL,
                               ObjToByteArray(result.value.obj, &dw2),
                               &dw2, NULL)) {
                Tcl_SetByteArrayLength(result.value.obj, dw2);
                result.type = TRT_OBJ;
            } else {
                Tcl_DecrRefCount(result.value.obj);
            }
        }
        break;

    case 10012: // CertNameToStr
        if (TwapiGetArgs(interp, objc, objv, ARGSKIP, GETINT(dw), ARGEND)
            != TCL_OK)
            return TCL_ERROR;
        blob.pbData = ObjToByteArray(objv[0], &blob.cbData);
        dw2 = CertNameToStrW(X509_ASN_ENCODING, &blob, dw, NULL, 0);
        result.value.unicode.str = TwapiAlloc(dw2*sizeof(WCHAR));
        result.value.unicode.len = CertNameToStrW(X509_ASN_ENCODING, &blob, dw, result.value.unicode.str, dw2) - 1;
        result.type = TRT_UNICODE_DYNAMIC;
        break;

    case 10013: // CertGetNameString
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(certP, CERT_CONTEXT*, CertFreeCertificateContext),
                         GETINT(dw), GETINT(dw2), ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
            
        return TwapiCertGetNameString(interp, certP, dw, dw2, objv[3]);

    case 10014: // CertFreeCertificateContext
        /* TBD -
           CertDuplicateCertificateContext will return the same pointer!
           However, our registration will barf when trying to release
           it the second time. Perhaps if the Cert API deals with bad
           pointer values, do not register it ourselves. Or do not
           implement the CertDuplicateCertificateContext call */
        if (TwapiGetArgs(interp, objc, objv,
                         GETPTR(certP, CERT_CONTEXT*), ARGEND) != TCL_OK ||
            TwapiUnregisterPointer(interp, certP, CertFreeCertificateContext) != TCL_OK)
            return TCL_ERROR;
        TWAPI_ASSERT(certP);
        result.type = TRT_EMPTY;
        CertFreeCertificateContext(certP);
        break;

    case 10015: // TwapiFindCertBySubjectName
        /* Supports tiny subset of CertFindCertificateInStore */
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(h, HCERTSTORE, CertCloseStore),
                         GETWSTR(s1), GETPTR(certP, CERT_CONTEXT*),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        /* Unregister previous context since the next call will free it */
        if (certP &&
            TwapiUnregisterPointer(interp, certP, CertFreeCertificateContext) != TCL_OK)
            return TCL_ERROR;
        certP = CertFindCertificateInStore(
            h,
            X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
            0,
            CERT_FIND_SUBJECT_STR_W,
            s1,
            certP);
        if (certP) {
            if (TwapiRegisterPointer(interp, certP, CertFreeCertificateContext) != TCL_OK)
                Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
            TwapiResult_SET_NONNULL_PTR(result, CERT_CONTEXT*, (void*)certP);
        } else {
            result.type = GetLastError() == CRYPT_E_NOT_FOUND ? TRT_EMPTY : TRT_GETLASTERROR;
        }
        break;
            
    case 10016: // CertUnregisterSystemStore
        /* This command is there to primarily clean up mistakes in testing */
        if (TwapiGetArgs(interp, objc, objv,
                         GETWSTR(s1), GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_EXCEPTION_ON_FALSE;
        result.value.ival = CertUnregisterSystemStore(s1, dw);
        break;
    case 10017:
        if (TwapiGetArgs(interp, objc, objv,
                         GETHANDLET(h, HCERTSTORE), ARGUSEDEFAULT,
                         GETINT(dw), ARGEND) != TCL_OK ||
            TwapiUnregisterPointer(interp, h, CertCloseStore) != TCL_OK)
            return TCL_ERROR;

        result.type = TRT_BOOL;
        result.value.bval = CertCloseStore(h, dw);
        if (result.value.bval == FALSE) {
            if (GetLastError() != CRYPT_E_PENDING_CLOSE)
                result.type = TRT_GETLASTERROR;
        }
        break;

    case 10018: // CryptGetUserKey
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        if (CryptGetUserKey((HCRYPTPROV) pv, dw, &dwp)) {
            if (TwapiRegisterPointer(interp, (void*)dwp, CryptDestroyKey) != TCL_OK)
                Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
            TwapiResult_SET_PTR(result, HCRYPTKEY, (void*)dwp);
        } else
            result.type = TRT_GETLASTERROR;
        break;

    case 10019: // CryptSetProvParam
        if (TwapiGetArgs(interp, objc, objv,
                         GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext),
                         GETINT(dw), GETINT(dw2), ARGSKIP, ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_CryptSetProvParam(interp, (HCRYPTPROV) pv, dw, dw2, objv[3]);

    }

    return TwapiSetResult(interp, &result);
}


static int TwapiCryptoInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s CryptoDispatch[] = {
        DEFINE_FNCODE_CMD(crypt_release_context, 10001), // Doc TBD
        DEFINE_FNCODE_CMD(CryptGetProvParam, 10002),
        DEFINE_FNCODE_CMD(cert_open_system_store, 10003), // Doc TBD
        DEFINE_FNCODE_CMD(cert_delete_from_store, 10004), // Doc TBD
        DEFINE_FNCODE_CMD(Twapi_SetCertContextKeyProvInfo, 10005),
        DEFINE_FNCODE_CMD(CertEnumCertificatesInStore, 10006),
        DEFINE_FNCODE_CMD(CertEnumCertificateContextProperties, 10007),
        DEFINE_FNCODE_CMD(CertGetCertificateContextProperty, 10008),
        DEFINE_FNCODE_CMD(crypt_destroy_key, 10009), // Doc TBD
        DEFINE_FNCODE_CMD(CryptGenKey, 10010),
        DEFINE_FNCODE_CMD(CertStrToName, 10011),
        DEFINE_FNCODE_CMD(CertNameToStr, 10012),
        DEFINE_FNCODE_CMD(CertGetNameString, 10013),
        DEFINE_FNCODE_CMD(cert_free, 10014), //CertFreeCertificateContext - doc
        DEFINE_FNCODE_CMD(TwapiFindCertBySubjectName, 10015),
        DEFINE_FNCODE_CMD(CertUnregisterSystemStore, 10016),
        DEFINE_FNCODE_CMD(CertCloseStore, 10017),
        DEFINE_FNCODE_CMD(CryptGetUserKey, 10018),
        DEFINE_FNCODE_CMD(CryptGetProvParam, 10019),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CryptoDispatch), CryptoDispatch, Twapi_CryptoCallObjCmd);
    Tcl_CreateObjCommand(interp, "twapi::CertCreateSelfSignCertificate", Twapi_CertCreateSelfSignCertificate, ticP, NULL);

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

