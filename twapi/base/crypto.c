/* 
 * Copyright (c) 2007-2009 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to CryptoAPI */

#include "twapi.h"

Tcl_Obj *ObjFromSecHandle(SecHandle *shP)
{
    Tcl_Obj *objv[2];
    objv[0] = ObjFromULONG_PTR(shP->dwLower);
    objv[1] = ObjFromULONG_PTR(shP->dwUpper);
    return Tcl_NewListObj(2, objv);
}

int ObjToSecHandle(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle *shP)
{
    int       objc;
    Tcl_Obj **objv;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) != TCL_OK)
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
    Tcl_Obj *obj = Tcl_NewListObj(0, NULL);

    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, obj, spiP, fCapabilities);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, obj, spiP, wVersion);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, obj, spiP, wRPCID);
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
            free(sbdP->pBuffers[i].pvBuffer);
            sbdP->pBuffers[i].pvBuffer = NULL;
        }
    }
    free(sbdP->pBuffers);
    return;
}


/* Returned buffer must be freed using TwapiFreeSecBufferDesc */
int ObjToSecBufferDesc(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP, int readonly)
{
    Tcl_Obj **objv;
    int      objc;
    int      i;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    sbdP->ulVersion = SECBUFFER_VERSION;
    sbdP->cBuffers = 0;         /* We will incr as we go along so we know
                                   how many to free in case of errors */

    if (Twapi_malloc(interp, NULL, objc*sizeof(SecBuffer), &sbdP->pBuffers) != TCL_OK)
        return TCL_ERROR;
    
    /* Each element of the list is a SecBuffer consisting of a pair
     * containing the integer type and the data itself
     */
    for (i=0; i < objc; ++i) {
        Tcl_Obj **bufobjv;
        int       bufobjc;
        int       buftype;
        int       datalen;
        char     *dataP;
        if (Tcl_ListObjGetElements(interp, objv[i], &bufobjc, &bufobjv) != TCL_OK)
            return TCL_ERROR;
        if (bufobjc != 2 ||
            Tcl_GetIntFromObj(interp, bufobjv[0], &buftype) != TCL_OK) {
            Tcl_SetResult(interp, "Invalid SecBuffer format", TCL_STATIC);
            goto handle_error;
        }
        dataP = Tcl_GetByteArrayFromObj(bufobjv[1], &datalen);
        if (Twapi_malloc(interp, NULL, datalen, &sbdP->pBuffers[i].pvBuffer) != TCL_OK)
            goto handle_error;
        sbdP->cBuffers++;
        sbdP->pBuffers[i].cbBuffer = datalen;
        if (readonly)
            buftype |= SECBUFFER_READONLY;
        sbdP->pBuffers[i].BufferType = buftype;
        memmove(sbdP->pBuffers[i].pvBuffer, dataP, datalen);
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

    resultObj = Tcl_NewListObj(0, NULL);
    if (sbdP->ulVersion != SECBUFFER_VERSION)
        return resultObj;

    for (i = 0; i < sbdP->cBuffers; ++i) {
        Tcl_Obj *bufobj[2];
        bufobj[0] = Tcl_NewIntObj(sbdP->pBuffers[i].BufferType);
        bufobj[1] = Tcl_NewByteArrayObj(sbdP->pBuffers[i].pvBuffer,
                                        sbdP->pBuffers[i].cbBuffer);
        Tcl_ListObjAppendElement(NULL, resultObj, Tcl_NewListObj(2, bufobj));
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

    obj = Tcl_NewListObj(0, NULL);
    for (i = 0; i < npkgs; ++i) {
        Tcl_ListObjAppendElement(interp, obj, ObjFromSecPkgInfo(&spiP[i]));
    }

    FreeContextBuffer(spiP);

    Tcl_SetObjResult(interp, obj);
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
    TimeStamp    *expirationP;

    if (gTwapiOSVersionInfo.dwMajorVersion == 5 &&
        gTwapiOSVersionInfo.dwMinorVersion >= 1) {
        /* XP and above */
        expirationP = &expiration;
    } else {
        expirationP = NULL;
    }

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
        expirationP);

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
    objv[3] = Tcl_NewLongObj(new_context_attr);
    if (expirationP)
        objv[4] = Tcl_NewWideIntObj(expirationP->QuadPart);

    Tcl_SetObjResult(interp,
                     Tcl_NewListObj(expirationP ? 5 : 4, objv));

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
    TimeStamp    *expirationP;

    if (gTwapiOSVersionInfo.dwMajorVersion == 5 &&
        gTwapiOSVersionInfo.dwMinorVersion >= 1) {
        /* XP and above */
        expirationP = &expiration;
    } else {
        expirationP = NULL;
    }

    /* We will ask the function to allocate buffer for us */
    sb_out.BufferType = SECBUFFER_TOKEN;
    sb_out.cbBuffer   = 0;
    sb_out.pvBuffer   = NULL;

    sbd_out.cBuffers  = 1;
    sbd_out.pBuffers  = &sb_out;
    sbd_out.ulVersion = SECBUFFER_VERSION;

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
    objv[3] = Tcl_NewLongObj(new_context_attr);
    if (expirationP)
        objv[4] = Tcl_NewWideIntObj(expirationP->QuadPart);

    Tcl_SetObjResult(interp,
                     Tcl_NewListObj(expirationP ? 5 : 4, objv));


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

    if (Twapi_malloc(NULL,
                     NULL,
                     sizeof(*swaiP)+sizeof(WCHAR)*(userlen+domainlen+passwordlen+3),
                     &swaiP) != TCL_OK)
        return NULL;

    swaiP->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
    swaiP->User  = (LPWSTR) (sizeof(*swaiP)+(char *)swaiP);
    swaiP->UserLength = (unsigned short) userlen;
    swaiP->Domain = swaiP->UserLength + 1 + swaiP->User;
    swaiP->DomainLength = (unsigned short) domainlen;
    swaiP->Password = swaiP->DomainLength + 1 + swaiP->Domain;
    swaiP->PasswordLength = (unsigned short) passwordlen;

    memmove(swaiP->User, user, sizeof(WCHAR)*(userlen+1));
    memmove(swaiP->Domain, domain, sizeof(WCHAR)*(domainlen+1));
    memmove(swaiP->Password, password, sizeof(WCHAR)*(passwordlen+1));

    return swaiP;
}

void Twapi_Free_SEC_WINNT_AUTH_IDENTITY (SEC_WINNT_AUTH_IDENTITY_W *swaiP)
{
    if (swaiP)
        free(swaiP);
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
                    obj = Tcl_NewUnicodeObj(buf, -1);
                break;
            case SECPKG_ATTR_FLAGS:
                obj = Tcl_NewLongObj(param.flags.Flags);
                break;
            case SECPKG_ATTR_SIZES:
                objv[0] = Tcl_NewLongObj(param.sizes.cbMaxToken);
                objv[1] = Tcl_NewLongObj(param.sizes.cbMaxSignature);
                objv[2] = Tcl_NewLongObj(param.sizes.cbBlockSize);
                objv[3] = Tcl_NewLongObj(param.sizes.cbSecurityTrailer);
                obj = Tcl_NewListObj(4, objv);
                break;
            case SECPKG_ATTR_STREAM_SIZES:
                objv[0] = Tcl_NewLongObj(param.streamsizes.cbHeader);
                objv[1] = Tcl_NewLongObj(param.streamsizes.cbTrailer);
                objv[2] = Tcl_NewLongObj(param.streamsizes.cbMaximumMessage);
                objv[3] = Tcl_NewLongObj(param.streamsizes.cBuffers);
                objv[4] = Tcl_NewLongObj(param.streamsizes.cbBlockSize);
                obj = Tcl_NewListObj(5, objv);
                break;
            case SECPKG_ATTR_LIFESPAN:
                objv[0] = Tcl_NewWideIntObj(param.lifespan.tsStart.QuadPart);
                objv[1] = Tcl_NewWideIntObj(param.lifespan.tsExpiry.QuadPart);
                obj = Tcl_NewListObj(2, objv);
                break;
            case SECPKG_ATTR_NAMES:
                buf = param.names.sUserName; /* Freed later */
                if (buf)
                    obj = Tcl_NewUnicodeObj(buf, -1);
                break;
            case SECPKG_ATTR_NATIVE_NAMES:
                objv[0] = Tcl_NewUnicodeObj(param.nativenames.sClientName ? param.nativenames.sClientName : L"", -1);
                objv[1] = Tcl_NewUnicodeObj(param.nativenames.sServerName ? param.nativenames.sServerName : L"", -1);
                obj = Tcl_NewListObj(2, objv);
                if (param.nativenames.sClientName)
                    FreeContextBuffer(param.nativenames.sClientName);
                if (param.nativenames.sServerName)
                    FreeContextBuffer(param.nativenames.sServerName);
                break;
            case SECPKG_ATTR_PASSWORD_EXPIRY:
                obj = Tcl_NewWideIntObj(param.passwordexpiry.tsPasswordExpires.QuadPart);
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
        Tcl_SetObjResult(interp, obj);

    return TCL_OK;
}

int Twapi_MakeSignature(
    Tcl_Interp *interp,
    SecHandle *ctxP,
    ULONG qop,
    int   datalen,
    void *dataP,
    ULONG seqnum
    )
{
    SECURITY_STATUS ss;
    SecPkgContext_Sizes spc_sizes;
    char sigbuf[32];
    void *sigP;
    SecBuffer sbufs[2];
    SecBufferDesc sbd;

    ss = QueryContextAttributesW(ctxP, SECPKG_ATTR_SIZES, &spc_sizes);
    if (ss != SEC_E_OK)
        return Twapi_AppendSystemError(interp, ss);

    if (spc_sizes.cbMaxSignature > sizeof(sigbuf)) {
        if (Twapi_malloc(interp, NULL, spc_sizes.cbMaxSignature, &sigP) != TCL_OK)
            return TCL_ERROR;
    } else
        sigP = &sigbuf;
    
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
        Twapi_AppendSystemError(interp, ss);
    } else {
        Tcl_Obj *objv[2];
        objv[0] = Tcl_NewByteArrayObj(sbufs[0].pvBuffer, sbufs[0].cbBuffer);
        objv[1] = Tcl_NewByteArrayObj(sbufs[1].pvBuffer, sbufs[1].cbBuffer);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
    }

    if (sigP != sigbuf)
        free(sigP);

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}


int Twapi_EncryptMessage(
    Tcl_Interp *interp,
    SecHandle *ctxP,
    ULONG qop,
    int   datalen,
    void *dataP,
    ULONG seqnum
    )
{
    SECURITY_STATUS ss;
    SecPkgContext_Sizes spc_sizes;
    char edatabuf[128];
    char padbuf[32];
    char trailerbuf[32];
    void *padP;
    void *trailerP;
    void *edataP;
    SecBuffer sbufs[3];
    SecBufferDesc sbd;

    ss = QueryContextAttributesW(ctxP, SECPKG_ATTR_SIZES, &spc_sizes);
    if (ss != SEC_E_OK)
        return Twapi_AppendSystemError(interp, ss);

    ss = SEC_E_INSUFFICIENT_MEMORY; /* Assumed error */

    if (spc_sizes.cbSecurityTrailer > sizeof(trailerbuf)) {
        if (Twapi_malloc(interp, NULL, spc_sizes.cbSecurityTrailer, &trailerP) != TCL_OK)
            goto vamoose;
    } else
        trailerP = &trailerbuf;

    if (spc_sizes.cbBlockSize > sizeof(padbuf)) {
        if (Twapi_malloc(interp, NULL, spc_sizes.cbBlockSize, &padP) != TCL_OK)
            goto vamoose;
    } else
        padP = &padbuf;

    if (datalen > sizeof(edatabuf)) {
        if (Twapi_malloc(interp, NULL, datalen, &edataP) != TCL_OK)
            goto vamoose;
    } else
        edataP = &edatabuf;
    memmove(edataP, dataP, datalen);
    
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
        Twapi_AppendSystemError(interp, ss);
    } else {
        Tcl_Obj *objv[3];
        objv[0] = Tcl_NewByteArrayObj(sbufs[0].pvBuffer, sbufs[0].cbBuffer);
        objv[1] = Tcl_NewByteArrayObj(sbufs[1].pvBuffer, sbufs[1].cbBuffer);
        objv[2] = Tcl_NewByteArrayObj(sbufs[2].pvBuffer, sbufs[2].cbBuffer);
        Tcl_SetObjResult(interp, Tcl_NewListObj(3, objv));
    }

vamoose:
    if (padP != padbuf)
        free(padP);
    if (edataP != edatabuf)
        free(edataP);
    if (trailerP != trailerbuf)
        free(trailerP);

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}

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
        Tcl_SetObjResult(interp, Tcl_NewByteArrayObj(buf, len));
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
}
