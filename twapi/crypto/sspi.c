/* 
 * Copyright (c) 2007-2009 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to CryptoAPI */

#include "twapi.h"
#include "twapi_crypto.h"

static Tcl_Obj *ObjFromSecHandle(SecHandle *shP);
static int ObjToSecHandle(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle *shP);
static int ObjToSecHandle_NULL(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle **shPP);
static Tcl_Obj *ObjFromSecPkgInfo(SecPkgInfoW *spiP);
static void TwapiFreeSecBufferDesc(SecBufferDesc *sbdP);
static int ObjToSecBufferDesc(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP, int readonly);
static int ObjToSecBufferDescRO(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP);
static int ObjToSecBufferDescRW(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP);
static Tcl_Obj *ObjFromSecBufferDesc(SecBufferDesc *sbdP);

Tcl_Obj *ObjFromSecHandle(SecHandle *shP)
{
    Tcl_Obj *objv[2];
    objv[0] = ObjFromULONG_PTR(shP->dwLower);
    objv[1] = ObjFromULONG_PTR(shP->dwUpper);
    return ObjNewList(2, objv);
}

int ObjToSecHandle(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle *shP)
{
    int       objc;
    Tcl_Obj **objv;

    if (ObjGetElements(interp, obj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if (objc != 2 ||
        ObjToULONG_PTR(interp, objv[0], &shP->dwLower) != TCL_OK ||
        ObjToULONG_PTR(interp, objv[1], &shP->dwUpper) != TCL_OK) {
        ObjSetStaticResult(interp, "Invalid security handle format");
        return TCL_ERROR;
    }
    return TCL_OK;
}

int ObjToSecHandle_NULL(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle **shPP)
{
    int n;
    if (ObjListLength(interp, obj, &n) != TCL_OK)
        return TCL_ERROR;
    if (n == 0) {
        *shPP = NULL;
        return TCL_OK;
    } else
        return ObjToSecHandle(interp, obj, *shPP);
}


Tcl_Obj *ObjFromSecPkgInfo(SecPkgInfoW *spiP)
{
    Tcl_Obj *obj = ObjNewList(0, NULL);

    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, obj, spiP, fCapabilities);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, obj, spiP, wVersion);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, obj, spiP, wRPCID);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, obj, spiP, cbMaxToken);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(NULL, obj, spiP, Name);
    Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(NULL, obj, spiP, Comment);

    return obj;
}

void TwapiFreeSecBufferDesc(SecBufferDesc *sbdP)
{
    ULONG i;
    if (sbdP == NULL || sbdP->pBuffers == NULL)
        return;
    for (i=0; i < sbdP->cBuffers; ++i) {
        if (sbdP->pBuffers[i].pvBuffer) {
            TwapiFree(sbdP->pBuffers[i].pvBuffer);
            sbdP->pBuffers[i].pvBuffer = NULL;
        }
    }
    TwapiFree(sbdP->pBuffers);
    return;
}


/* Returned buffer must be freed using TwapiFreeSecBufferDesc */
int ObjToSecBufferDesc(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP, int readonly)
{
    Tcl_Obj **objv;
    int      objc;
    int      i;

    if (ObjGetElements(interp, obj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    sbdP->ulVersion = SECBUFFER_VERSION;
    sbdP->cBuffers = 0;         /* We will incr as we go along so we know
                                   how many to free in case of errors */

    sbdP->pBuffers = TwapiAlloc(objc*sizeof(SecBuffer));
    
    /* Each element of the list is a SecBuffer consisting of a pair
     * containing the integer type and the data itself
     */
    for (i=0; i < objc; ++i) {
        Tcl_Obj **bufobjv;
        int       bufobjc;
        int       buftype;
        int       datalen;
        char     *dataP;
        if (ObjGetElements(interp, objv[i], &bufobjc, &bufobjv) != TCL_OK)
            return TCL_ERROR;
        if (bufobjc != 2 ||
            Tcl_GetIntFromObj(interp, bufobjv[0], &buftype) != TCL_OK) {
            ObjSetStaticResult(interp, "Invalid SecBuffer format");
            goto handle_error;
        }
        dataP = ObjToByteArray(bufobjv[1], &datalen);
        sbdP->pBuffers[i].pvBuffer = TwapiAlloc(datalen);
        sbdP->cBuffers++;
        sbdP->pBuffers[i].cbBuffer = datalen;
        if (readonly)
            buftype |= SECBUFFER_READONLY;
        sbdP->pBuffers[i].BufferType = buftype;
        CopyMemory(sbdP->pBuffers[i].pvBuffer, dataP, datalen);
    }

    return TCL_OK;

handle_error:
    /* Free any existing buffers */
    TwapiFreeSecBufferDesc(sbdP);
    return TCL_ERROR;
}

/* Returned buffer must be freed using TwapiFreeSecBufferDesc */
int ObjToSecBufferDescRO(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP)
{
    return ObjToSecBufferDesc(interp, obj, sbdP, 1);
}

/* Returned buffer must be freed using TwapiFreeSecBufferDesc */
int ObjToSecBufferDescRW(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP)
{
    return ObjToSecBufferDesc(interp, obj, sbdP, 0);
}



Tcl_Obj *ObjFromSecBufferDesc(SecBufferDesc *sbdP) 
{
    Tcl_Obj *resultObj;
    DWORD i;

    resultObj = ObjNewList(0, NULL);
    if (sbdP->ulVersion != SECBUFFER_VERSION)
        return resultObj;

    for (i = 0; i < sbdP->cBuffers; ++i) {
        Tcl_Obj *bufobj[2];
        bufobj[0] = ObjFromDWORD(sbdP->pBuffers[i].BufferType);
        bufobj[1] = ObjFromByteArray(sbdP->pBuffers[i].pvBuffer,
                                        sbdP->pBuffers[i].cbBuffer);
        ObjAppendElement(NULL, resultObj, ObjNewList(2, bufobj));
    }
    return resultObj;
}

int Twapi_EnumerateSecurityPackages(Tcl_Interp *interp)
{
    ULONG i, npkgs;
    SecPkgInfoW *spiP;
    SECURITY_STATUS status;
    Tcl_Obj *obj;

    status = EnumerateSecurityPackagesW(&npkgs, &spiP);
    if (status != SEC_E_OK)
        return Twapi_AppendSystemError(interp, status);

    obj = ObjNewList(0, NULL);
    for (i = 0; i < npkgs; ++i) {
        ObjAppendElement(interp, obj, ObjFromSecPkgInfo(&spiP[i]));
    }

    FreeContextBuffer(spiP);

    ObjSetResult(interp, obj);
    return TCL_OK;
}

int Twapi_InitializeSecurityContext(
    Tcl_Interp *interp,
    SecHandle *credentialP,
    SecHandle *contextP,
    LPWSTR     targetP,
    ULONG      contextreq,
    ULONG      reserved1,
    ULONG      targetdatarep,
    SecBufferDesc *sbd_inP,
    ULONG     reserved2)
{
    SecBuffer     sb_out;
    SecBufferDesc sbd_out;
    SECURITY_STATUS status;
    CtxtHandle    new_context;
    ULONG         new_context_attr;
    Tcl_Obj      *objv[5];
    TimeStamp     expiration;

    /* We will ask the function to allocate buffer for us */
    sb_out.BufferType = SECBUFFER_TOKEN;
    sb_out.cbBuffer   = 0;
    sb_out.pvBuffer   = NULL;

    sbd_out.cBuffers  = 1;
    sbd_out.pBuffers  = &sb_out;
    sbd_out.ulVersion = SECBUFFER_VERSION;

    status = InitializeSecurityContextW(
        credentialP,
        contextP,
        targetP,
        contextreq | ISC_REQ_ALLOCATE_MEMORY,
        reserved1,
        targetdatarep,
        sbd_inP,
        reserved2,
        &new_context,
        &sbd_out,
        &new_context_attr,
        &expiration);

    switch (status) {
    case SEC_E_OK:
        objv[0] = STRING_LITERAL_OBJ("ok");
        break;
    case SEC_I_CONTINUE_NEEDED:
        objv[0] = STRING_LITERAL_OBJ("continue");
        break;
    case SEC_I_COMPLETE_NEEDED:
        objv[0] = STRING_LITERAL_OBJ("complete");
        break;
    case SEC_I_COMPLETE_AND_CONTINUE:
        objv[0] = STRING_LITERAL_OBJ("complete_and_continue");
        break;
    default:
        return Twapi_AppendSystemError(interp, status);
    }

    objv[1] = ObjFromSecHandle(&new_context);
    objv[2] = ObjFromSecBufferDesc(&sbd_out);
    objv[3] = ObjFromLong(new_context_attr);
    objv[4] = ObjFromWideInt(expiration.QuadPart);

    ObjSetResult(interp, ObjNewList(5, objv));

    if (sb_out.pvBuffer)
        FreeContextBuffer(sb_out.pvBuffer);

    return TCL_OK;
}


int Twapi_AcceptSecurityContext(
    Tcl_Interp *interp,
    SecHandle *credentialP,
    SecHandle *contextP,
    SecBufferDesc *sbd_inP,
    ULONG      contextreq,
    ULONG      targetdatarep)
{
    SecBuffer     sb_out;
    SecBufferDesc sbd_out;
    SECURITY_STATUS status;
    CtxtHandle    new_context;
    ULONG         new_context_attr;
    Tcl_Obj      *objv[5];
    TimeStamp     expiration;

    /* We will ask the function to allocate buffer for us */
    sb_out.BufferType = SECBUFFER_TOKEN;
    sb_out.cbBuffer   = 0;
    sb_out.pvBuffer   = NULL;

    sbd_out.cBuffers  = 1;
    sbd_out.pBuffers  = &sb_out;
    sbd_out.ulVersion = SECBUFFER_VERSION;

    /* TBD - MSDN says expiration pointer should be NULL until
       last call in negotiation sequence. Does it really need
       to be NULL or can we expect caller just ignore the result?
       We assume the latter rfor now.
    */
    status = AcceptSecurityContext(
        credentialP,
        contextP,
        sbd_inP,
        contextreq | ASC_REQ_ALLOCATE_MEMORY,
        targetdatarep,
        &new_context,
        &sbd_out,
        &new_context_attr,
        &expiration);

    switch (status) {
    case SEC_E_OK:
        objv[0] = STRING_LITERAL_OBJ("ok");
        break;
    case SEC_I_CONTINUE_NEEDED:
        objv[0] = STRING_LITERAL_OBJ("continue");
        break;
    case SEC_I_COMPLETE_NEEDED:
        objv[0] = STRING_LITERAL_OBJ("complete");
        break;
    case SEC_I_COMPLETE_AND_CONTINUE:
        objv[0] = STRING_LITERAL_OBJ("complete_and_continue");
        break;
    case SEC_E_INCOMPLETE_MESSAGE:
        objv[0] = STRING_LITERAL_OBJ("incomplete_message");
        break;
    default:
        return Twapi_AppendSystemError(interp, status);
    }

    objv[1] = ObjFromSecHandle(&new_context);
    objv[2] = ObjFromSecBufferDesc(&sbd_out);
    objv[3] = ObjFromLong(new_context_attr);
    objv[4] = ObjFromWideInt(expiration.QuadPart);

    ObjSetResult(interp,
                     ObjNewList(5, objv));


    if (sb_out.pvBuffer)
        FreeContextBuffer(sb_out.pvBuffer);

    return TCL_OK;
}

/* Note caller has to clean up ticP->memlifo irrespective of success/error */
static TCL_RESULT ParseSEC_WINNT_AUTH_IDENTITY (
    TwapiInterpContext *ticP,
    Tcl_Obj *authObj,
    SEC_WINNT_AUTH_IDENTITY_W **swaiPP
    )
{
    Tcl_Obj *passwordObj;
    LPWSTR    password;
    Tcl_Obj **objv;
    int objc;
    TCL_RESULT res;
    SEC_WINNT_AUTH_IDENTITY_W *swaiP;
    
    if ((res = ObjGetElements(ticP->interp, authObj, &objc, &objv)) != TCL_OK)
        return res;

    if (objc == 0) {
        *swaiPP = NULL;
        return TCL_OK;
    }

    swaiP = MemLifoAlloc(&ticP->memlifo, sizeof(*swaiP), NULL);
    swaiP->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETSTRWN(swaiP->User, swaiP->UserLength),
                         GETSTRWN(swaiP->Domain, swaiP->DomainLength),
                         GETOBJ(passwordObj),
                         ARGEND);
    if (res != TCL_OK)
        return res;

    password = ObjDecryptPassword(passwordObj, &swaiP->PasswordLength);
    swaiP->Password = MemLifoCopy(&ticP->memlifo, password, sizeof(WCHAR)*(swaiP->PasswordLength+1));
    TwapiFreeDecryptedPassword(password, swaiP->PasswordLength);

    *swaiPP = swaiP;
    return TCL_OK;
}


static TCL_RESULT ParseSCHANNEL_CRED (
    TwapiInterpContext *ticP,
    Tcl_Obj *authObj,
    SCHANNEL_CRED **credPP
    )
{
    Tcl_Obj **objv;
    int objc;
    TCL_RESULT res;
    Tcl_Obj *certsObj;
    SCHANNEL_CRED *credP;
    
    if ((res = ObjGetElements(ticP->interp, authObj, &objc, &objv)) != TCL_OK)
        return res;

    if (objc == 0) {
        *credPP = NULL;
        return TCL_OK;
    }

    credP = MemLifoAlloc(&ticP->memlifo, sizeof(*credP), NULL);
    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETINT(credP->dwVersion),
                         GETOBJ(certsObj),
                         GETVERIFIEDORNULL(credP->hRootStore, HCERTSTORE, CertCloseStore),
                         ARGUSEDEFAULT,
                         ARGUNUSED, /* aphMappers */
                         ARGUNUSED, /* palgSupportedAlgs */
                         GETINT(credP->grbitEnabledProtocols),
                         GETINT(credP->dwMinimumCipherStrength),
                         GETINT(credP->dwMaximumCipherStrength),
                         GETINT(credP->dwSessionLifespan),
                         GETINT(credP->dwFlags),
                         GETINT(credP->dwCredFormat),
                         ARGEND);
    if (res != TCL_OK)
        return res;

    TWAPI_ASSERT(SCHANNEL_CRED_VERSION == 3);
    if (credP->dwVersion != SCHANNEL_CRED_VERSION) {
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid SCHANNEL_CRED_VERSION");
    }

    res = ObjGetElements(ticP->interp, certsObj, &objc, &objv);
    credP->cCreds = objc;
    credP->paCred = MemLifoAlloc(&ticP->memlifo, objc * sizeof(*credP->paCred), NULL);
    while (objc--) {
        res = ObjToVerifiedPointer(ticP->interp, objv[objc],
                                   (void **)&credP->paCred[objc],
                                   "CERT_CONTEXT*",
                                   CertFreeCertificateContext);
        /* TBD - should we dup the cert context ? Will that even help ?
           What if the app frees the cert_context ? Dos AcquireCredentials
           dup the context ? */
        if (res != TCL_OK)
            return res;
    }

    credP->cMappers = 0;
    credP->aphMappers = NULL;
    credP->cSupportedAlgs = 0;
    credP->palgSupportedAlgs = NULL;

    *credPP = credP;
    return TCL_OK;
}




int Twapi_QueryContextAttributes(
    Tcl_Interp *interp,
    SecHandle *ctxP,
    ULONG attr
)
{
    void *buf;
    union {
        SecPkgContext_AuthorityW authority;
        SecPkgContext_Flags      flags;
        SecPkgContext_Lifespan   lifespan;
        SecPkgContext_NamesW     names;
        SecPkgContext_Sizes      sizes;
        SecPkgContext_StreamSizes    streamsizes;
        SecPkgContext_NativeNamesW   nativenames;
        SecPkgContext_PasswordExpiry passwordexpiry;
    } param;
    SECURITY_STATUS ss;
    Tcl_Obj *obj;
    Tcl_Obj *objv[5];

    buf = NULL;
    obj = NULL;
    switch (attr) {
    case SECPKG_ATTR_AUTHORITY:
    case SECPKG_ATTR_FLAGS:
    case SECPKG_ATTR_LIFESPAN:
    case SECPKG_ATTR_SIZES:
    case SECPKG_ATTR_STREAM_SIZES:
    case SECPKG_ATTR_NAMES:
    case SECPKG_ATTR_NATIVE_NAMES:
    case SECPKG_ATTR_PASSWORD_EXPIRY:
        ss = QueryContextAttributesW(ctxP, attr, &param);
        if (ss == SEC_E_OK) {
            switch (attr) {
            case SECPKG_ATTR_AUTHORITY:
                buf = param.authority.sAuthorityName; /* Freed later */
                if (buf)
                    obj = ObjFromUnicode(buf);
                break;
            case SECPKG_ATTR_FLAGS:
                obj = ObjFromLong(param.flags.Flags);
                break;
            case SECPKG_ATTR_SIZES:
                objv[0] = ObjFromLong(param.sizes.cbMaxToken);
                objv[1] = ObjFromLong(param.sizes.cbMaxSignature);
                objv[2] = ObjFromLong(param.sizes.cbBlockSize);
                objv[3] = ObjFromLong(param.sizes.cbSecurityTrailer);
                obj = ObjNewList(4, objv);
                break;
            case SECPKG_ATTR_STREAM_SIZES:
                objv[0] = ObjFromLong(param.streamsizes.cbHeader);
                objv[1] = ObjFromLong(param.streamsizes.cbTrailer);
                objv[2] = ObjFromLong(param.streamsizes.cbMaximumMessage);
                objv[3] = ObjFromLong(param.streamsizes.cBuffers);
                objv[4] = ObjFromLong(param.streamsizes.cbBlockSize);
                obj = ObjNewList(5, objv);
                break;
            case SECPKG_ATTR_LIFESPAN:
                objv[0] = ObjFromWideInt(param.lifespan.tsStart.QuadPart);
                objv[1] = ObjFromWideInt(param.lifespan.tsExpiry.QuadPart);
                obj = ObjNewList(2, objv);
                break;
            case SECPKG_ATTR_NAMES:
                buf = param.names.sUserName; /* Freed later */
                if (buf)
                    obj = ObjFromUnicode(buf);
                break;
            case SECPKG_ATTR_NATIVE_NAMES:
                objv[0] = ObjFromUnicode(param.nativenames.sClientName ? param.nativenames.sClientName : L"");
                objv[1] = ObjFromUnicode(param.nativenames.sServerName ? param.nativenames.sServerName : L"");
                obj = ObjNewList(2, objv);
                if (param.nativenames.sClientName)
                    FreeContextBuffer(param.nativenames.sClientName);
                if (param.nativenames.sServerName)
                    FreeContextBuffer(param.nativenames.sServerName);
                break;
            case SECPKG_ATTR_PASSWORD_EXPIRY:
                obj = ObjFromWideInt(param.passwordexpiry.tsPasswordExpires.QuadPart);
                break;
            }
        }
        break;
        
    default:
        ObjSetStaticResult(interp, "Unsupported QuerySecurityContext attribute id");
    }

    if (buf)
        FreeContextBuffer(buf);

    if (ss)
        return Twapi_AppendSystemError(interp, ss);

    if (obj)
        ObjSetResult(interp, obj);

    return TCL_OK;
}

int Twapi_MakeSignature(
    TwapiInterpContext *ticP,
    SecHandle *ctxP,
    ULONG qop,
    int datalen,
    void *dataP, /* Points into Tcl_Obj, must NOT be modified ! */
    ULONG seqnum)
{
    SECURITY_STATUS ss;
    SecPkgContext_Sizes spc_sizes;
    void *sigP;
    SecBuffer sbufs[2];
    SecBufferDesc sbd;

    ss = QueryContextAttributesW(ctxP, SECPKG_ATTR_SIZES, &spc_sizes);
    if (ss != SEC_E_OK)
        return Twapi_AppendSystemError(ticP->interp, ss);

    /* TBD - change to directly use ByteArray without memlifo allocs */

    sigP = MemLifoPushFrame(&ticP->memlifo, spc_sizes.cbMaxSignature, NULL);
    
    sbufs[0].BufferType = SECBUFFER_TOKEN;
    sbufs[0].cbBuffer   = spc_sizes.cbMaxSignature;
    sbufs[0].pvBuffer   = sigP;
    sbufs[1].BufferType = SECBUFFER_DATA | SECBUFFER_READONLY;
    sbufs[1].cbBuffer   = datalen;
    sbufs[1].pvBuffer   = dataP;

    sbd.cBuffers = 2;
    sbd.pBuffers = sbufs;
    sbd.ulVersion = SECBUFFER_VERSION;

    ss = MakeSignature(ctxP, qop, &sbd, seqnum);
    if (ss != SEC_E_OK) {
        Twapi_AppendSystemError(ticP->interp, ss);
    } else {
        Tcl_Obj *objv[2];
        objv[0] = ObjFromByteArray(sbufs[0].pvBuffer, sbufs[0].cbBuffer);
        objv[1] = ObjFromByteArray(sbufs[1].pvBuffer, sbufs[1].cbBuffer);
        ObjSetResult(ticP->interp, ObjNewList(2, objv));
    }

    MemLifoPopFrame(&ticP->memlifo);

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}


int Twapi_EncryptMessage(
    TwapiInterpContext *ticP,
    SecHandle *ctxP,
    ULONG qop,
    int   datalen,
    void *dataP, /* Must not be modified, may point into Tcl_Obj owned space */
    ULONG seqnum
    )
{
    SECURITY_STATUS ss;
    SecPkgContext_Sizes spc_sizes;
    void *padP;
    void *trailerP;
    void *edataP;
    SecBuffer sbufs[3];
    SecBufferDesc sbd;

    ss = QueryContextAttributesW(ctxP, SECPKG_ATTR_SIZES, &spc_sizes);
    if (ss != SEC_E_OK)
        return Twapi_AppendSystemError(ticP->interp, ss);

    ss = SEC_E_INSUFFICIENT_MEMORY; /* Assumed error */

    /* TBD - change to directly use ByteArray without memlifo allocs */

    trailerP = MemLifoPushFrame(&ticP->memlifo,
                                spc_sizes.cbSecurityTrailer, NULL);
    padP = MemLifoAlloc(&ticP->memlifo, spc_sizes.cbBlockSize, NULL);
    edataP = MemLifoCopy(&ticP->memlifo, dataP, datalen);
    
    sbufs[0].BufferType = SECBUFFER_TOKEN;
    sbufs[0].cbBuffer   = spc_sizes.cbSecurityTrailer;
    sbufs[0].pvBuffer   = trailerP;
    sbufs[1].BufferType = SECBUFFER_DATA;
    sbufs[1].cbBuffer   = datalen;
    sbufs[1].pvBuffer   = edataP;
    sbufs[2].BufferType = SECBUFFER_PADDING;
    sbufs[2].cbBuffer   = spc_sizes.cbBlockSize;
    sbufs[2].pvBuffer   = padP;

    sbd.cBuffers = 3;
    sbd.pBuffers = sbufs;
    sbd.ulVersion = SECBUFFER_VERSION;

    ss = EncryptMessage(ctxP, qop, &sbd, seqnum);
    if (ss != SEC_E_OK) {
        Twapi_AppendSystemError(ticP->interp, ss);
    } else {
        Tcl_Obj *objv[3];
        objv[0] = ObjFromByteArray(sbufs[0].pvBuffer, sbufs[0].cbBuffer);
        objv[1] = ObjFromByteArray(sbufs[1].pvBuffer, sbufs[1].cbBuffer);
        objv[2] = ObjFromByteArray(sbufs[2].pvBuffer, sbufs[2].cbBuffer);
        ObjSetResult(ticP->interp, ObjNewList(3, objv));
    }

    MemLifoPopFrame(&ticP->memlifo);

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}

static int Twapi_SignEncryptObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SecHandle sech;
    int func;
    DWORD dw, dw2, dw3;
    unsigned char *cP;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETINT(func),
                     GETVAR(sech, ObjToSecHandle),
                     GETINT(dw),
                     ARGSKIP,
                     GETINT(dw3),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (func != 1 && func != 2)
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);

    cP = ObjToByteArray(objv[4], &dw2);

    return (func == 1 ? Twapi_MakeSignature : Twapi_EncryptMessage) (
        ticP, &sech, dw, dw2, cP, dw3);
}

static int Twapi_AcquireCredentialsHandleObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    LUID luid, *luidP;
    LPWSTR principalP;
    LPWSTR packageP;
    DWORD cred_use;
    Tcl_Obj *authObj;
    SECURITY_STATUS status;
    LARGE_INTEGER timestamp;
    SecHandle credH; 
    Tcl_Obj *objs[2];
    MemLifoMarkHandle mark;
    TCL_RESULT res = TCL_ERROR;
    int is_unisp;
    void *pv;

    pv = NULL;
    mark = MemLifoPushMark(&ticP->memlifo);
    luidP = &luid;
    if (TwapiGetArgsEx(ticP, objc-1, objv+1,
                       GETEMPTYASNULL(principalP), GETSTRW(packageP),
                       GETINT(cred_use), GETVAR(luidP, ObjToLUID_NULL),
                       GETOBJ(authObj), ARGEND) != TCL_OK)
        goto vamoose;

    if (WSTREQ(packageP, UNISP_NAME_W) == 0) {
        if (ParseSCHANNEL_CRED(ticP, authObj, &(SCHANNEL_CRED *)pv) != TCL_OK)
            goto vamoose;
        is_unisp = 1;
    } else if (WSTREQ(packageP, WDIGEST_SP_NAME_W) ||
              WSTREQ(packageP, NTLMSP_NAME) ||
              WSTREQ(packageP, NEGOSSP_NAME_W) ||
              WSTREQ(packageP, MICROSOFT_KERBEROS_NAME_W)) {
        if (ParseSEC_WINNT_AUTH_IDENTITY(ticP, authObj, &(SEC_WINNT_AUTH_IDENTITY_W *)pv) != TCL_OK)
            goto vamoose;
        is_unisp = 0;
    } else {
        return TwapiReturnErrorMsg(interp, TWAPI_UNSUPPORTED_TYPE, "Unsupported SSPI package");
    }

    status = AcquireCredentialsHandleW(principalP, packageP, cred_use,
                                       luidP, pv, NULL, NULL, &credH, &timestamp);
    if (status != SEC_E_OK) {
        Twapi_AppendSystemError(ticP->interp, status);
        goto vamoose;
    }
    objs[0] = ObjFromSecHandle(&credH);
    objs[1] = ObjFromWideInt(timestamp.QuadPart);
    ObjSetResult(ticP->interp, ObjNewList(2, objs));

    res = TCL_OK;

vamoose:
    if (pv && !is_unisp) {
        SEC_WINNT_AUTH_IDENTITY_W *swaiP = pv;
        if (swaiP->Password && swaiP->PasswordLength)
            SecureZeroMemory(swaiP->Password, sizeof(WCHAR)*(swaiP->PasswordLength));
    }

    MemLifoPopMark(mark);
    return res;
}

static int Twapi_SspiCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    DWORD dw, dw2, dw3, dw4;
    Tcl_Obj *s1Obj;
    SecHandle sech, sech2, *sech2P;
    SecBufferDesc sbd, *sbdP;
    int func = PtrToInt(clientdata);

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;
    if (func < 100) {
        /* Functions taking no arguments */
        if (objc > 0)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 1:
            return Twapi_EnumerateSecurityPackages(interp);
        }
    } else if (func < 200) {
        /* Single arg is a sechandle */
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToSecHandle(interp, objv[0], &sech) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 101:
            result.type = TRT_HANDLE;
            dw = QuerySecurityContextToken(&sech, &result.value.hval);
            if (dw) {
                result.value.ival =  dw;
                result.type = TRT_EXCEPTION_ON_ERROR;
            } else {
                result.type = TRT_HANDLE;
            }
            break;
        case 102: // FreeCredentialsHandle
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = FreeCredentialsHandle(&sech);
            break;
        case 103: // DeleteSecurityContext
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = DeleteSecurityContext(&sech);
            break;
        case 104: // ImpersonateSecurityContext
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = ImpersonateSecurityContext(&sech);
            break;
        }
    } else {
        /* Free-for-all - each func responsible for checking arguments */
        switch (func) {
        case 10021:
            sech2P = &sech2;
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sech2P, ObjToSecHandle_NULL),
                             GETOBJ(s1Obj),
                             GETINT(dw),
                             GETINT(dw2),
                             GETINT(dw3),
                             GETVAR(sbd, ObjToSecBufferDescRO),
                             GETINT(dw4),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            sbdP = sbd.cBuffers ? &sbd : NULL;
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_InitializeSecurityContext(
                interp, &sech, sech2P, ObjToUnicode(s1Obj),
                dw, dw2, dw3, sbdP, dw4);
            TwapiFreeSecBufferDesc(sbdP);
            break;

        case 10022: // AcceptSecurityContext
            sech2P = &sech2;
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sech2P, ObjToSecHandle_NULL),
                             GETVAR(sbd, ObjToSecBufferDescRO),
                             GETINT(dw),
                             GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            sbdP = sbd.cBuffers ? &sbd : NULL;
            result.type = TRT_TCL_RESULT;
            result.value.ival = Twapi_AcceptSecurityContext(
                interp, &sech, sech2P, sbdP, dw, dw2);
            TwapiFreeSecBufferDesc(sbdP);
            break;
        
        case 10023: // QueryContextAttributes
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(sech, ObjToSecHandle),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_QueryContextAttributes(interp, &sech, dw);

        case 10024: // VerifySignature
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sbd, ObjToSecBufferDescRO),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            sbdP = sbd.cBuffers ? &sbd : NULL;
            dw2 = VerifySignature(&sech, sbdP, dw, &result.value.uval);
            TwapiFreeSecBufferDesc(sbdP);
            if (dw2 == 0)
                result.type = TRT_DWORD;
            else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw2;
            }
            break;

        case 10025: // DecryptMessage
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sbd, ObjToSecBufferDescRW),
                             GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            dw2 = DecryptMessage(&sech, &sbd, dw, &result.value.ival);
            if (dw2 == 0) {
                result.type = TRT_OBJ;
                result.value.obj = ObjFromSecBufferDesc(&sbd);
            } else {
                result.type = TRT_EXCEPTION_ON_ERROR;
                result.value.ival = dw2;
            }
            TwapiFreeSecBufferDesc(&sbd);
            break;


        }
    }

    return TwapiSetResult(interp, &result);
}


int TwapiSspiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct alias_dispatch_s SignEncryptAliasDispatch[] = {
        DEFINE_ALIAS_CMD(MakeSignature, 1),
        DEFINE_ALIAS_CMD(EncryptMessage, 2)
    };

    static struct fncode_dispatch_s SspiDispatch[] = {
        DEFINE_FNCODE_CMD(EnumerateSecurityPackages, 1),
        DEFINE_FNCODE_CMD(QuerySecurityContextToken, 101),
        DEFINE_FNCODE_CMD(FreeCredentialsHandle, 102),
        DEFINE_FNCODE_CMD(DeleteSecurityContext, 103),
        DEFINE_FNCODE_CMD(ImpersonateSecurityContext, 104),
        DEFINE_FNCODE_CMD(InitializeSecurityContext, 10021),
        DEFINE_FNCODE_CMD(AcceptSecurityContext, 10022),
        DEFINE_FNCODE_CMD(QueryContextAttributes, 10023),
        DEFINE_FNCODE_CMD(VerifySignature, 10024),
        DEFINE_FNCODE_CMD(DecryptMessage, 10025),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(SspiDispatch), SspiDispatch, Twapi_SspiCallObjCmd);
    Tcl_CreateObjCommand(interp, "twapi::CallSignEncrypt", Twapi_SignEncryptObjCmd, ticP, NULL);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(SignEncryptAliasDispatch), SignEncryptAliasDispatch, "twapi::CallSignEncrypt");
    Tcl_CreateObjCommand(interp, "twapi::AcquireCredentialsHandle", Twapi_AcquireCredentialsHandleObjCmd, ticP, NULL);

    return TCL_OK;
}



