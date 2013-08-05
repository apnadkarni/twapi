/* 
 * Copyright (c) 2007-2009 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to CryptoAPI */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif


Tcl_Obj *ObjFromSecHandle(SecHandle *shP);
int ObjToSecHandle(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle *shP);
int ObjToSecHandle_NULL(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle **shPP);
Tcl_Obj *ObjFromSecPkgInfo(SecPkgInfoW *spiP);
void TwapiFreeSecBufferDesc(SecBufferDesc *sbdP);
int ObjToSecBufferDesc(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP, int readonly);
int ObjToSecBufferDescRO(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP);
int ObjToSecBufferDescRW(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP);
Tcl_Obj *ObjFromSecBufferDesc(SecBufferDesc *sbdP);

static int Twapi_SignEncryptObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_EnumerateSecurityPackages(Tcl_Interp *interp);
int Twapi_InitializeSecurityContext(
    Tcl_Interp *interp,
    SecHandle *credentialP,
    SecHandle *contextP,
    LPWSTR     targetP,
    ULONG      contextreq,
    ULONG      reserved1,
    ULONG      targetdatarep,
    SecBufferDesc *sbd_inP,
    ULONG     reserved2);
int Twapi_AcceptSecurityContext(Tcl_Interp *interp, SecHandle *credentialP,
                                SecHandle *contextP, SecBufferDesc *sbd_inP,
                                ULONG contextreq, ULONG targetdatarep);
int Twapi_QueryContextAttributes(Tcl_Interp *interp, SecHandle *INPUT,
                                 ULONG attr);
SEC_WINNT_AUTH_IDENTITY_W *Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY (
    LPCWSTR user, LPCWSTR domain, LPCWSTR password);
void Twapi_Free_SEC_WINNT_AUTH_IDENTITY (SEC_WINNT_AUTH_IDENTITY_W *swaiP);
int Twapi_MakeSignature(TwapiInterpContext *ticP, SecHandle *INPUT,
                        ULONG qop, int BINLEN, void *BINDATA, ULONG seqnum);
int Twapi_EncryptMessage(TwapiInterpContext *ticP, SecHandle *INPUT,
                        ULONG qop, int BINLEN, void *BINDATA, ULONG seqnum);
int Twapi_CryptGenRandom(Tcl_Interp *interp, HCRYPTPROV hProv, DWORD dwLen);

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
        Tcl_SetResult(interp, "Invalid security handle format", TCL_STATIC);
        return TCL_ERROR;
    }
    return TCL_OK;
}

int ObjToSecHandle_NULL(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle **shPP)
{
    int n;
    if (Tcl_ListObjLength(interp, obj, &n) != TCL_OK)
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
            Tcl_SetResult(interp, "Invalid SecBuffer format", TCL_STATIC);
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

    TwapiSetObjResult(interp, obj);
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

    TwapiSetObjResult(interp, ObjNewList(5, objv));

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

    TwapiSetObjResult(interp,
                     ObjNewList(5, objv));


    if (sb_out.pvBuffer)
        FreeContextBuffer(sb_out.pvBuffer);

    return TCL_OK;
}

SEC_WINNT_AUTH_IDENTITY_W *Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY (
    LPCWSTR    user,
    LPCWSTR    domain,
    LPCWSTR    password
    )
{
    int userlen, domainlen, passwordlen;
    SEC_WINNT_AUTH_IDENTITY_W *swaiP;

    userlen    = lstrlenW(user);
    domainlen  = lstrlenW(domain);
    passwordlen = lstrlenW(password);

    swaiP = TwapiAlloc(sizeof(*swaiP)+sizeof(WCHAR)*(userlen+domainlen+passwordlen+3));

    swaiP->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
    swaiP->User  = (LPWSTR) (sizeof(*swaiP)+(char *)swaiP);
    swaiP->UserLength = (unsigned short) userlen;
    swaiP->Domain = swaiP->UserLength + 1 + swaiP->User;
    swaiP->DomainLength = (unsigned short) domainlen;
    swaiP->Password = swaiP->DomainLength + 1 + swaiP->Domain;
    swaiP->PasswordLength = (unsigned short) passwordlen;

    CopyMemory(swaiP->User, user, sizeof(WCHAR)*(userlen+1));
    CopyMemory(swaiP->Domain, domain, sizeof(WCHAR)*(domainlen+1));
    CopyMemory(swaiP->Password, password, sizeof(WCHAR)*(passwordlen+1));

    return swaiP;
}

void Twapi_Free_SEC_WINNT_AUTH_IDENTITY (SEC_WINNT_AUTH_IDENTITY_W *swaiP)
{
    if (swaiP)
        TwapiFree(swaiP);
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
        Tcl_SetResult(interp, "Unsupported QuerySecurityContext attribute id", TCL_STATIC);
    }

    if (buf)
        FreeContextBuffer(buf);

    if (ss)
        return Twapi_AppendSystemError(interp, ss);

    if (obj)
        TwapiSetObjResult(interp, obj);

    return TCL_OK;
}

int Twapi_MakeSignature(
    TwapiInterpContext *ticP,
    SecHandle *ctxP,
    ULONG qop,
    int datalen,
    void *dataP,
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
        TwapiSetObjResult(ticP->interp, ObjNewList(2, objv));
    }

    MemLifoPopFrame(&ticP->memlifo);

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}


int Twapi_EncryptMessage(
    TwapiInterpContext *ticP,
    SecHandle *ctxP,
    ULONG qop,
    int   datalen,
    void *dataP,
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
    edataP = MemLifoAlloc(&ticP->memlifo, datalen, NULL);
    CopyMemory(edataP, dataP, datalen);
    
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
        TwapiSetObjResult(ticP->interp, ObjNewList(3, objv));
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
                     GETBIN(cP, dw2),
                     GETINT(dw3),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (func != 1 && func != 2)
        return TwapiReturnError(interp, TWAPI_INVALID_ARGS);

    return (func == 1 ? Twapi_MakeSignature : Twapi_EncryptMessage) (
        ticP, &sech, dw, dw2, cP, dw3);
}

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
 pointers may point to Tcl_Obj areas within valObj */
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

static TwapiCertGetNameString(
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


static int Twapi_CryptoCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    DWORD dw, dw2, dw3, dw4;
    DWORD_PTR dwp;
    LPVOID pv;
    LPWSTR s1, s2, s3;
    HANDLE h;
    SecHandle sech, sech2, *sech2P;
    SecBufferDesc sbd, *sbdP;
    LUID luid, *luidP;
    LARGE_INTEGER largeint;
    Tcl_Obj *objs[2];
    int func = PtrToInt(clientdata);
    PCCERT_CONTEXT certP;
    struct _CRYPTOAPI_BLOB blob;

    --objc;
    ++objv;

    TWAPI_ASSERT(sizeof(HCRYPTPROV) <= sizeof(pv));
    TWAPI_ASSERT(sizeof(HCRYPTKEY) <= sizeof(pv));
    TWAPI_ASSERT(sizeof(dwp) <= sizeof(void*));

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
    } else if (func < 300) {
        /* Single arg of any type */
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        switch (func) {
        case 201:
            h = CertOpenSystemStoreW(0, ObjToUnicode(objv[0]));
            /* CertCloseStore does not check ponter validity! So do ourselves*/
            if (TwapiRegisterPointer(interp, h, CertCloseStore) != TCL_OK)
                Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
            TwapiResult_SET_NONNULL_PTR(result, HCERTSTORE, h);
            break;
        }
    } else {
        /* Free-for-all - each func responsible for checking arguments */
        switch (func) {
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
            
        case 10016:
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
        case 10018:
            if (TwapiGetArgs(interp, objc, objv,
                             GETWSTR(s1), ARGUSEDEFAULT,
                             GETWSTR(s2), GETWSTR(s3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            ;
            TwapiResult_SET_NONNULL_PTR(result, SEC_WINNT_AUTH_IDENTITY_W*, Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY(s1, s2, s3));
            break;
        case 10019:
            if (objc != 1)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            if (ObjToHANDLE(interp, objv[0], &h) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            Twapi_Free_SEC_WINNT_AUTH_IDENTITY(h);
            break;
        case 10020:
            luidP = &luid;
            if (TwapiGetArgs(interp, objc, objv,
                             GETNULLIFEMPTY(s1), GETWSTR(s2), GETINT(dw),
                             GETVAR(luidP, ObjToLUID_NULL),
                             GETVOIDP(pv), ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.value.ival = AcquireCredentialsHandleW(
                s1, s2,
                dw, luidP, pv, NULL, NULL, &sech, &largeint);
            if (result.value.ival) {
                result.type = TRT_EXCEPTION_ON_ERROR;
                break;
            }
            objs[0] = ObjFromSecHandle(&sech);
            objs[1] = ObjFromWideInt(largeint.QuadPart);
            result.type = TRT_OBJV;
            result.value.objv.objPP = objs;
            result.value.objv.nobj = 2;
            break;
        case 10021:
            sech2P = &sech2;
            if (TwapiGetArgs(interp, objc, objv,
                             GETVAR(sech, ObjToSecHandle),
                             GETVAR(sech2P, ObjToSecHandle_NULL),
                             GETWSTR(s1),
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
                interp, &sech, sech2P, s1,
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

        case 10024: //CryptReleaseContext
            if (TwapiGetArgs(interp, objc, objv,
                             GETVERIFIEDPTR(pv, HCRYPTPROV, CryptReleaseContext), GETINT(dw),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = CryptReleaseContext((HCRYPTPROV)pv, dw);
            break;

        case 10025: // CryptAcquireContext
            if (TwapiGetArgs(interp, objc, objv,
                             GETNULLIFEMPTY(s1), GETNULLIFEMPTY(s2), GETINT(dw), GETINT(dw2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (CryptAcquireContextW(&dwp, s1, s2, dw, dw2)) {
                if (dw2 & CRYPT_DELETEKEYSET)
                    result.type = TRT_EMPTY;
                else {
                    if (TwapiRegisterPointer(interp, (void*)dwp, CryptReleaseContext) != TCL_OK)
                        Tcl_Panic("Failed to register pointer: %s", Tcl_GetStringResult(interp));
                    TwapiResult_SET_PTR(result, HCRYPTPROV, (void*)dwp);
                }
            } else
                result.type = TRT_GETLASTERROR;
            break;

        case 10026: // VerifySignature
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

        case 10027: // DecryptMessage
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


static int TwapiCryptoInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct alias_dispatch_s SignEncryptAliasDispatch[] = {
        DEFINE_ALIAS_CMD(MakeSignature, 1),
        DEFINE_ALIAS_CMD(EncryptMessage, 2)
    };

    static struct fncode_dispatch_s CryptoDispatch[] = {
        DEFINE_FNCODE_CMD(EnumerateSecurityPackages, 1),
        DEFINE_FNCODE_CMD(QuerySecurityContextToken, 101),
        DEFINE_FNCODE_CMD(FreeCredentialsHandle, 102),
        DEFINE_FNCODE_CMD(DeleteSecurityContext, 103),
        DEFINE_FNCODE_CMD(ImpersonateSecurityContext, 104),
        DEFINE_FNCODE_CMD(cert_open_system_store, 201), // Doc TBD
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
        DEFINE_FNCODE_CMD(Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY, 10018),
        DEFINE_FNCODE_CMD(Twapi_Free_SEC_WINNT_AUTH_IDENTITY, 10019),
        DEFINE_FNCODE_CMD(AcquireCredentialsHandle, 10020),
        DEFINE_FNCODE_CMD(InitializeSecurityContext, 10021),
        DEFINE_FNCODE_CMD(AcceptSecurityContext, 10022),
        DEFINE_FNCODE_CMD(QueryContextAttributes, 10023),
        DEFINE_FNCODE_CMD(CryptReleaseContext, 10024),
        DEFINE_FNCODE_CMD(CryptAcquireContext, 10025),
        DEFINE_FNCODE_CMD(VerifySignature, 10026),
        DEFINE_FNCODE_CMD(DecryptMessage, 10027),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(CryptoDispatch), CryptoDispatch, Twapi_CryptoCallObjCmd);
    Tcl_CreateObjCommand(interp, "twapi::CallSignEncrypt", Twapi_SignEncryptObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::CertCreateSelfSignCertificate", Twapi_CertCreateSelfSignCertificate, ticP, NULL);
    TwapiDefineAliasCmds(interp, ARRAYSIZE(SignEncryptAliasDispatch), SignEncryptAliasDispatch, "twapi::CallSignEncrypt");

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

