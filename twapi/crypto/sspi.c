/* 
 * Copyright (c) 2007-2009 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Interface to CryptoAPI */

/*
Table showing SECBUFFER usage for different functions/providers in different sample code
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|Source        |InitializeContext|AcceptContext|EncryptMesssage    |DecryptMessage     |MakeSignature |VerifySignature              |
|              |                 |             |                   |                   |              |                             |
|              |                 |             |                   |                   |              |                             |
|              |                 |             |                   |                   |              |                             |
|              |                 |             |                   |                   |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|Richter       |in=TOKEN         |in=TOKEN     |TOKEN, DATA,       |TOKEN, DATA,       |TOKEN, DATA   |TOKEN, DATA                  |
|              |out=TOKEN        |out=TOKEN    |PADDING            |PADDING            |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|Richter (ssl) |(in NULL first   |in=TOKEN,    |STREAM_HEADER,     |DATA, EMPTY,       |              |                             |
|              |call)            |EMPTY        |DATA,              |EMPTY, EMPTY       |              |                             |
|              |in=TOKEN,EMPTY   |out=TOKEN    |STREAM_TRAILER     |                   |              |                             |
|              |out=TOKEN        |             |                   |                   |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|WebClient     |(in NULL first   |             |STREAM_HEADER,     |DATA, EMPTY,       |              |                             |
|              |call)            |             |DATA,              |EMPTY, EMPTY       |              |                             |
|              |in=TOKEN,        |             |STREAM_TRAILER,    |                   |              |                             |
|              |EMPTY out=TOKEN  |             |EMPTY              |                   |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|WebServer     |                 |in=TOKEN,    |STREAM_HEADER,     |DATA, EMPTY,       |              |                             |
|              |                 |EMPTY        |DATA,              |EMPTY, EMPTY       |              |                             |
|              |                 |out=TOKEN    |STREAM_TRAILER,    |                   |              |                             |
|              |                 |             |EMPTY              |                   |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|curl          |first call       |             |same as            |Same as            |              |                             |
|              |in=NULL,         |             |WebClient/Webserver|WEbClient/Webserver|              |                             |
|              |out=EMPTY Further|             |                   |                   |              |                             |
|              |calls in=TOKEN,  |             |                   |                   |              |                             |
|              |EMPTY out=TOKEN, |             |                   |                   |              |                             |
|              |ALERT            |             |                   |                   |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|coast         |first call       |             | same as           |same as            |              |                             |
|              |in=NULL,         |             |WebCleint/Webserver|Webclient/Webserver|              |                             |
|              |out=TOKEN Further|             |                   |                   |              |                             |
|              |in=TOKEN, EMPTY  |             |                   |                   |              |                             |
|              |out=TOKEN        |             |                   |                   |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+
|codeproject   |                 |             |                   |                   |              |                             |
+--------------+-----------------+-------------+-------------------+-------------------+--------------+-----------------------------+

*/


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

/* Returns a Tcl_Obj from the SECBUFFER_EXTRA buffer in a SecBufferDesc.
   If no such buffer or sbdP is NULL, returns an empty Tcl_Obj.
   First buffer is expected to contain the input data.
*/
static Tcl_Obj *ObjFromSECBUFFER_EXTRA(SecBufferDesc *sbdP) 
{
    int i;
    if (sbdP) {
        SecBuffer *sb0P = &sbdP->pBuffers[0]; /* Should contain input data */
        /*
         * Go backward because EXTRA buffer likely to be at end. First
         * buffer is the input so we do not check that. 
         */
        for (i = sbdP->cBuffers - 1; i > 0; --i) {
            SecBuffer *sbP = &sbdP->pBuffers[i];
            /* MSDN docs say "pvBuffer" does not contain a copy of the
               data. What they probably mean is that it points into
               the original data, not to a copy of it. But the sample
               code in MSDN and curl etc. does not make use of this field
               so neither do we, just to be safe. */
            if (sbP->BufferType == SECBUFFER_EXTRA && sbP->cbBuffer) {
                TWAPI_ASSERT(sb0P->cbBuffer > sbP->cbBuffer);
                return ObjFromByteArray(((char*) sb0P->pvBuffer) + (sb0P->cbBuffer - sbP->cbBuffer),
                                        sbP->cbBuffer);
            }
        }
    }
    return ObjFromEmptyString();
}

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
    
    /* Each element of the list is a SecBuffer consisting of the
     * integer type and optionally the data itself
     */
    for (i=0; i < objc; ++i) {
        Tcl_Obj **bufobjv;
        int       bufobjc;
        int       buftype;
        int       datalen;
        char     *dataP;
        if (ObjGetElements(interp, objv[i], &bufobjc, &bufobjv) != TCL_OK)
            return TCL_ERROR;
        if ((bufobjc != 1 && bufobjc != 2) ||
            Tcl_GetIntFromObj(interp, bufobjv[0], &buftype) != TCL_OK) {
            ObjSetStaticResult(interp, "Invalid SecBuffer format");
            goto handle_error;
        }
        if (bufobjc == 2) {
            dataP = ObjToByteArray(bufobjv[1], &datalen);
            sbdP->pBuffers[i].pvBuffer = TwapiAlloc(datalen);
            sbdP->pBuffers[i].cbBuffer = datalen;
            CopyMemory(sbdP->pBuffers[i].pvBuffer, dataP, datalen);
        } else {
            sbdP->pBuffers[i].pvBuffer = NULL;
            sbdP->pBuffers[i].cbBuffer = 0;
        }
        sbdP->cBuffers++;
        if (readonly)
            buftype |= SECBUFFER_READONLY;
        sbdP->pBuffers[i].BufferType = buftype;
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

static TCL_RESULT Twapi_InitializeSecurityContextObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc, Tcl_Obj *CONST objv[])
{
    SecHandle credential, context, *contextP;
    LPWSTR     targetP;
    ULONG      contextreq, reserved1, targetdatarep, reserved2;
    SecBufferDesc sbd_in, *sbd_inP;
    SecBuffer     sb_out;
    SecBufferDesc sbd_out;
    SECURITY_STATUS status;
    CtxtHandle    new_context = {0, 0};
    ULONG         new_context_attr;
    Tcl_Obj      *objs[6];
    TimeStamp     expiration;

    /* NOTE - on XP for schannel at least, the input buffer must not 
       be read-only. Discovered in testing so made sbd_in read-write
       Later Win versions do not seem to mind */
    contextP = &context;
    if (TwapiGetArgsEx(ticP, objc-1, objv+1,
                       GETVAR(credential, ObjToSecHandle),
                       GETVAR(contextP, ObjToSecHandle_NULL),
                       GETEMPTYASNULL(targetP),
                       GETINT(contextreq),
                       GETINT(reserved1),
                       GETINT(targetdatarep),
                       GETVAR(sbd_in, ObjToSecBufferDescRW),
                       GETINT(reserved2),
                       ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (sbd_in.cBuffers == 0)
        sbd_inP = NULL;
    else {
        if ((sbd_in.pBuffers[0].BufferType & ~SECBUFFER_ATTRMASK) != SECBUFFER_TOKEN)
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "First buffer type must be SECBUFFER_TOKEN");
        sbd_inP = &sbd_in;
    }

    /*
     * We will ask the function to allocate buffer for us
     * Note all the providers we support take a single output buffer.
     */
    sb_out.BufferType = SECBUFFER_TOKEN;
    sb_out.cbBuffer   = 0;
    sb_out.pvBuffer   = NULL;

    sbd_out.cBuffers  = 1;
    sbd_out.pBuffers  = &sb_out;
    sbd_out.ulVersion = SECBUFFER_VERSION;

    status = InitializeSecurityContextW(
        &credential,
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
        objs[0] = STRING_LITERAL_OBJ("ok");
        break;
    case SEC_I_CONTINUE_NEEDED:
        objs[0] = STRING_LITERAL_OBJ("continue");
        break;
    case SEC_I_COMPLETE_NEEDED:
        objs[0] = STRING_LITERAL_OBJ("complete");
        break;
    case SEC_I_COMPLETE_AND_CONTINUE:
        objs[0] = STRING_LITERAL_OBJ("complete_and_continue");
        break;
    case SEC_I_INCOMPLETE_CREDENTIALS:
        objs[0] = STRING_LITERAL_OBJ("incomplete_credentials");
        break;
    case SEC_E_INCOMPLETE_MESSAGE:
        objs[0] = STRING_LITERAL_OBJ("incomplete_message");
        break;
    case SEC_I_CONTEXT_EXPIRED:
        objs[0] = STRING_LITERAL_OBJ("expired");
        break;
    default:
        TwapiFreeSecBufferDesc(sbd_inP);
        Twapi_AppendSystemError(interp, status);
        return TCL_ERROR;
    }

    /* TBD - check if below is valid for ALL status values */
    if (contextP) {
        objs[1] = ObjFromSecHandle(contextP);
        objs[3] = ObjFromLong(new_context_attr);
    } else {
        /* For some codes like INCOMPLETE_MESSAGE on the first call,
           new_context is not set by the AcceptSecurityContext call
        */
        if (new_context.dwLower || new_context.dwUpper) {
            /* Yes, a security handle was returned */
            objs[1] = ObjFromSecHandle(&new_context);
            objs[3] = ObjFromLong(new_context_attr);
        } else {
            objs[1] = ObjFromEmptyString();
            objs[3] = ObjFromLong(0);
        }
    }

    objs[2] = ObjFromSecBufferDesc(&sbd_out);
    objs[4] = ObjFromWideInt(expiration.QuadPart);

    /* Check if there was any unprocessed left over data that 
       has to be passed back to the caller. In case of INCOMPLETE_MESSAGE
       this is the original data passed in
    */
    if (status == SEC_E_INCOMPLETE_MESSAGE) {
        if (sbd_inP)
            objs[5] = ObjFromByteArray(sbd_inP->pBuffers[0].pvBuffer,
                                       sbd_inP->pBuffers[0].cbBuffer);
        else
            objs[5] = ObjFromEmptyString();
    } else
        objs[5] = ObjFromSECBUFFER_EXTRA(sbd_inP);

    ObjSetResult(interp, ObjNewList(6, objs));

    /* Note sb_out NOT allocated by us so DON'T call TwapiFreeSecBufferDesc */
    if (sb_out.pvBuffer)
        FreeContextBuffer(sb_out.pvBuffer);

    TwapiFreeSecBufferDesc(sbd_inP);
    return TCL_OK;
}

static int Twapi_AcceptSecurityContextObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SecHandle credential, context, *contextP;
    SecBufferDesc sbd_in, *sbd_inP;
    ULONG      contextreq, targetdatarep;
    SecBuffer     sb_out;
    SecBufferDesc sbd_out;
    SECURITY_STATUS status;
    CtxtHandle    new_context = {0, 0};
    ULONG         new_context_attr;
    Tcl_Obj      *objs[6];
    TimeStamp     expiration;

    /* NOTE - on XP for schannel at least, the input buffer must not 
       be read-only. Discovered in testing so made sbd_in read-write
       Later Win versions do not seem to mind */
    contextP = &context;
    if (TwapiGetArgsEx(ticP, objc-1, objv+1,
                       GETVAR(credential, ObjToSecHandle),
                       GETVAR(contextP, ObjToSecHandle_NULL),
                       GETVAR(sbd_in, ObjToSecBufferDescRW),
                       GETINT(contextreq),
                       GETINT(targetdatarep),
                       ARGEND) != TCL_OK)
        return TCL_ERROR;

    if (sbd_in.cBuffers == 0)
        sbd_inP = NULL;
    else {
        if ((sbd_in.pBuffers[0].BufferType & ~SECBUFFER_ATTRMASK) != SECBUFFER_TOKEN)
            return TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "First buffer type must be SECBUFFER_TOKEN");
        sbd_inP = &sbd_in;
    }

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
       We assume the latter for now.
    */
    status = AcceptSecurityContext(
        &credential,
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
        objs[0] = STRING_LITERAL_OBJ("ok");
        break;
    case SEC_I_CONTINUE_NEEDED:
        objs[0] = STRING_LITERAL_OBJ("continue");
        break;
    case SEC_I_COMPLETE_NEEDED:
        objs[0] = STRING_LITERAL_OBJ("complete");
        break;
    case SEC_I_COMPLETE_AND_CONTINUE:
        objs[0] = STRING_LITERAL_OBJ("complete_and_continue");
        break;
    case SEC_E_INCOMPLETE_MESSAGE:
        objs[0] = STRING_LITERAL_OBJ("incomplete_message");
        break;
    case SEC_I_CONTEXT_EXPIRED:
        objs[0] = STRING_LITERAL_OBJ("expired");
        break;
    default:
        Twapi_AppendSystemError(interp, status);
        TwapiFreeSecBufferDesc(sbd_inP);
        return TCL_ERROR;
    }

    /* TBD - check if below is valid for all status */
    if (contextP) {
        objs[1] = ObjFromSecHandle(contextP);
        objs[3] = ObjFromLong(new_context_attr);
    } else {
        /* For some codes like INCOMPLETE_MESSAGE on the first call,
           new_context is not set by the AcceptSecurityContext call
        */
        if (new_context.dwLower || new_context.dwUpper) {
            /* Yes, a security handle was returned */
            objs[1] = ObjFromSecHandle(&new_context);
            objs[3] = ObjFromLong(new_context_attr);
        } else {
            objs[1] = ObjFromEmptyString();
            objs[3] = ObjFromLong(0);
        }
    }
    objs[2] = ObjFromSecBufferDesc(&sbd_out);
    objs[4] = ObjFromWideInt(expiration.QuadPart);

    /* Check if there was any unprocessed left over data that 
       has to be passed back to the caller. In case of INCOMPLETE_MESSAGE
       this is the original data passed in
    */
    if (status == SEC_E_INCOMPLETE_MESSAGE) {
        if (sbd_inP)
            objs[5] = ObjFromByteArray(sbd_inP->pBuffers[0].pvBuffer,
                                       sbd_inP->pBuffers[0].cbBuffer);
        else
            objs[5] = ObjFromEmptyString();
    } else
        objs[5] = ObjFromSECBUFFER_EXTRA(sbd_inP);

    ObjSetResult(interp, ObjNewList(6, objs));

    /* Note sb_out NOT allocated by us so DON'T call TwapiFreeSecBufferDesc */
    if (sb_out.pvBuffer)
        FreeContextBuffer(sb_out.pvBuffer);

    TwapiFreeSecBufferDesc(sbd_inP);
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

    swaiP = MemLifoAlloc(ticP->memlifoP, sizeof(*swaiP), NULL);
    swaiP->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
    res = TwapiGetArgsEx(ticP, objc, objv,
                         GETWSTRN(swaiP->User, swaiP->UserLength),
                         GETWSTRN(swaiP->Domain, swaiP->DomainLength),
                         GETOBJ(passwordObj),
                         ARGEND);
    if (res != TCL_OK)
        return res;

    password = ObjDecryptPassword(passwordObj, &swaiP->PasswordLength);
    swaiP->Password = MemLifoCopy(ticP->memlifoP, password, sizeof(WCHAR)*(swaiP->PasswordLength+1));
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

    credP = MemLifoAlloc(ticP->memlifoP, sizeof(*credP), NULL);
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

    TWAPI_ASSERT(SCHANNEL_CRED_VERSION == 4);
    if (credP->dwVersion != SCHANNEL_CRED_VERSION) {
        return TwapiReturnErrorMsg(ticP->interp, TWAPI_INVALID_ARGS, "Invalid SCHANNEL_CRED_VERSION");
    }

    res = ObjGetElements(ticP->interp, certsObj, &objc, &objv);
    credP->cCreds = objc;
    credP->paCred = MemLifoAlloc(ticP->memlifoP, objc * sizeof(*credP->paCred), NULL);
    while (objc--) {
        res = ObjToVerifiedPointer(ticP->interp, objv[objc],
                                   (void **)&credP->paCred[objc],
                                   "CERT_CONTEXT*",
                                   CertFreeCertificateContext);
        /* TBD - should we dup the cert context ? Will that even help ?
           What if the app frees the cert_context ? Dos AcquireCredentials
           dup the context ? Apparently so - see comment in
           http://www.coastrd.com/c-schannel-smtp */

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
        SecPkgContext_PackageInfoW pkginfo;
        SecPkgContext_IssuerListInfoEx issuerlist;
        CERT_CONTEXT *certP;
    } param;
    SECURITY_STATUS ss;
    Tcl_Obj *obj;
    Tcl_Obj *objs[6];
    DWORD i;

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
    case SECPKG_ATTR_PACKAGE_INFO:
    case SECPKG_ATTR_LOCAL_CERT_CONTEXT:
    case SECPKG_ATTR_REMOTE_CERT_CONTEXT:
    case SECPKG_ATTR_ISSUER_LIST_EX:
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
                objs[0] = ObjFromLong(param.sizes.cbMaxToken);
                objs[1] = ObjFromLong(param.sizes.cbMaxSignature);
                objs[2] = ObjFromLong(param.sizes.cbBlockSize);
                objs[3] = ObjFromLong(param.sizes.cbSecurityTrailer);
                obj = ObjNewList(4, objs);
                break;
            case SECPKG_ATTR_STREAM_SIZES:
                objs[0] = ObjFromLong(param.streamsizes.cbHeader);
                objs[1] = ObjFromLong(param.streamsizes.cbTrailer);
                objs[2] = ObjFromLong(param.streamsizes.cbMaximumMessage);
                objs[3] = ObjFromLong(param.streamsizes.cBuffers);
                objs[4] = ObjFromLong(param.streamsizes.cbBlockSize);
                obj = ObjNewList(5, objs);
                break;
            case SECPKG_ATTR_LIFESPAN:
                objs[0] = ObjFromWideInt(param.lifespan.tsStart.QuadPart);
                objs[1] = ObjFromWideInt(param.lifespan.tsExpiry.QuadPart);
                obj = ObjNewList(2, objs);
                break;
            case SECPKG_ATTR_NAMES:
                buf = param.names.sUserName; /* Freed later */
                if (buf)
                    obj = ObjFromUnicode(buf);
                break;
            case SECPKG_ATTR_NATIVE_NAMES:
                objs[0] = ObjFromUnicode(param.nativenames.sClientName ? param.nativenames.sClientName : L"");
                objs[1] = ObjFromUnicode(param.nativenames.sServerName ? param.nativenames.sServerName : L"");
                obj = ObjNewList(2, objs);
                if (param.nativenames.sClientName)
                    FreeContextBuffer(param.nativenames.sClientName);
                if (param.nativenames.sServerName)
                    FreeContextBuffer(param.nativenames.sServerName);
                break;
            case SECPKG_ATTR_PASSWORD_EXPIRY:
                obj = ObjFromWideInt(param.passwordexpiry.tsPasswordExpires.QuadPart);
                break;
            case SECPKG_ATTR_PACKAGE_INFO:
                buf = param.pkginfo.PackageInfo; /* So it can be freed */
                if (param.pkginfo.PackageInfo) {
                    objs[0] = ObjFromULONG(param.pkginfo.PackageInfo->fCapabilities);
                    objs[1] = ObjFromInt(param.pkginfo.PackageInfo->wVersion);
                    objs[2] = ObjFromInt(param.pkginfo.PackageInfo->wRPCID);
                    objs[3] = ObjFromULONG(param.pkginfo.PackageInfo->cbMaxToken);
                    objs[4] = ObjFromUnicode(param.pkginfo.PackageInfo->Name);
                    objs[5] = ObjFromUnicode(param.pkginfo.PackageInfo->Comment);
                    obj = ObjNewList(6, objs);
                }
                break;
            case SECPKG_ATTR_LOCAL_CERT_CONTEXT: /* FALLTHRU */
            case SECPKG_ATTR_REMOTE_CERT_CONTEXT:
                TwapiRegisterCertPointer(interp, param.certP);
                obj = ObjFromOpaque(param.certP, "CERT_CONTEXT*");
                break;
            case SECPKG_ATTR_ISSUER_LIST_EX:
                obj = ObjNewList(param.issuerlist.cIssuers, NULL);
                for (i = 0; i < param.issuerlist.cIssuers; ++i) {
                    ObjAppendElement(interp, obj, ObjFromCERT_NAME_BLOB(&param.issuerlist.aIssuers[0], 0));
                }
                buf = param.issuerlist.aIssuers; /* So it is freed below */
                break;
            }
            
        }
        break;
        
    default:
        return TwapiReturnError(interp, TWAPI_UNSUPPORTED_TYPE);
    }

    if (buf)
        FreeContextBuffer(buf);

    if (ss)
        return Twapi_AppendSystemError(interp, ss);

    if (obj)
        ObjSetResult(interp, obj);

    return TCL_OK;
}

static TCL_RESULT Twapi_MakeSignatureObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SecHandle sech;
    ULONG qop;
    ULONG seqnum;

    SECURITY_STATUS ss;
    SecPkgContext_Sizes spc_sizes;
    SecBuffer sbufs[2];
    SecBufferDesc sbd;
    Tcl_Obj *objs[ARRAYSIZE(sbufs)];
    Tcl_Obj *dataObj;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVAR(sech, ObjToSecHandle),
                     GETINT(qop),
                     GETOBJ(dataObj),
                     GETINT(seqnum),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;


    ss = QueryContextAttributesW(&sech, SECPKG_ATTR_SIZES, &spc_sizes);
    if (ss != SEC_E_OK)
        return Twapi_AppendSystemError(interp, ss);

    objs[0] = ObjFromByteArray(NULL, spc_sizes.cbMaxSignature);
    sbufs[0].BufferType = SECBUFFER_TOKEN;
    sbufs[0].pvBuffer   = ObjToByteArray(objs[0], &sbufs[0].cbBuffer);

    objs[1] = ObjDuplicate(dataObj);
    TWAPI_ASSERT(! Tcl_IsShared(objs[1]));
    sbufs[1].BufferType = SECBUFFER_DATA | SECBUFFER_READONLY;
    sbufs[1].pvBuffer   = ObjToByteArray(objs[1], &sbufs[1].cbBuffer);
    
    sbd.cBuffers = 2;
    sbd.pBuffers = sbufs;
    sbd.ulVersion = SECBUFFER_VERSION;

    ss = MakeSignature(&sech, qop, &sbd, seqnum);
    if (ss != SEC_E_OK) {
        int i;
        Twapi_AppendSystemError(interp, ss);
        for (i=0; i < ARRAYSIZE(objs); ++i)
            ObjDecrRefs(objs[i]);
    } else {
        int i;
        for (i=0; i < ARRAYSIZE(objs); ++i)
            Tcl_SetByteArrayLength(objs[i], sbufs[i].cbBuffer);
        ObjSetResult(interp, ObjNewList(ARRAYSIZE(objs), objs));
    }

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}


static TCL_RESULT Twapi_EncryptMessageObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SecHandle sech;
    ULONG qop;
    ULONG seqnum;
    SECURITY_STATUS ss;
    SecPkgContext_Sizes spc_sizes;
    SecBuffer sbufs[3];
    SecBufferDesc sbd;
    Tcl_Obj *dataObj;
    Tcl_Obj *objs[ARRAYSIZE(sbufs)];

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVAR(sech, ObjToSecHandle),
                     GETINT(qop),
                     GETOBJ(dataObj),
                     GETINT(seqnum),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    ss = QueryContextAttributesW(&sech, SECPKG_ATTR_SIZES, &spc_sizes);
    if (ss != SEC_E_OK)
        return Twapi_AppendSystemError(ticP->interp, ss);

    objs[0] = ObjFromByteArray(NULL, spc_sizes.cbSecurityTrailer);
    sbufs[0].BufferType = SECBUFFER_TOKEN;
    sbufs[0].pvBuffer   = ObjToByteArray(objs[0], &sbufs[0].cbBuffer);

    objs[1] = ObjDuplicate(dataObj);
    TWAPI_ASSERT(! Tcl_IsShared(objs[1]));
    sbufs[1].BufferType = SECBUFFER_DATA;
    sbufs[1].pvBuffer   = ObjToByteArray(objs[1], &sbufs[1].cbBuffer);
    
    objs[2] = ObjFromByteArray(NULL, spc_sizes.cbBlockSize);
    sbufs[2].BufferType = SECBUFFER_PADDING;
    sbufs[2].pvBuffer   = ObjToByteArray(objs[2], &sbufs[2].cbBuffer);

    sbd.cBuffers = 3;
    sbd.pBuffers = sbufs;
    sbd.ulVersion = SECBUFFER_VERSION;

    ss = EncryptMessage(&sech, qop, &sbd, seqnum);
    if (ss != SEC_E_OK) {
        int i;
        Twapi_AppendSystemError(ticP->interp, ss);
        for (i=0; i < ARRAYSIZE(objs); ++i)
            ObjDecrRefs(objs[i]);
    } else {
        int i;
        for (i=0; i < ARRAYSIZE(objs); ++i)
            Tcl_SetByteArrayLength(objs[i], sbufs[i].cbBuffer);
        ObjSetResult(ticP->interp, ObjNewList(ARRAYSIZE(objs), objs));
    }

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}


static TCL_RESULT Twapi_EncryptStreamObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SecHandle sech;
    ULONG qop;
    SECURITY_STATUS ss;
    SecPkgContext_StreamSizes sizes;
    SecBuffer sbufs[4];
    SecBufferDesc sbd;
    Tcl_Obj *dataObj;
    Tcl_Obj *objs[2];           /* 0 encrypted data, 1 leftover data */
    char *dataP, *encP;
    DWORD   datalen;

    if (TwapiGetArgs(interp, objc-1, objv+1,
                     GETVAR(sech, ObjToSecHandle),
                     GETINT(qop),
                     GETOBJ(dataObj),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    ss = QueryContextAttributesW(&sech, SECPKG_ATTR_STREAM_SIZES, &sizes);
    if (ss != SEC_E_OK)
        return Twapi_AppendSystemError(interp, ss);

    dataP = ObjToByteArray(dataObj, &datalen);
    if (datalen > sizes.cbMaximumMessage) {
        objs[1] = ObjFromByteArray(dataP+sizes.cbMaximumMessage,
                                       datalen-sizes.cbMaximumMessage);
        datalen = sizes.cbMaximumMessage;
    } else
        objs[1] = ObjFromEmptyString();

    objs[0] = ObjFromByteArray(NULL, sizes.cbHeader + datalen + sizes.cbTrailer);
    encP = ObjToByteArray(objs[0], NULL);
    CopyMemory(encP + sizes.cbHeader, dataP, datalen);

    sbufs[0].BufferType = SECBUFFER_STREAM_HEADER;
    sbufs[0].pvBuffer   = encP;
    sbufs[0].cbBuffer   = sizes.cbHeader;

    sbufs[1].BufferType = SECBUFFER_DATA;
    sbufs[1].pvBuffer   = encP + sizes.cbHeader;
    sbufs[1].cbBuffer   = datalen;

    sbufs[2].BufferType = SECBUFFER_STREAM_TRAILER;
    sbufs[2].pvBuffer   = encP + sizes.cbHeader + datalen;
    sbufs[2].cbBuffer   = sizes.cbTrailer;

    sbufs[3].BufferType = SECBUFFER_EMPTY;
    sbufs[3].pvBuffer   = NULL;
    sbufs[3].cbBuffer   = 0;

    sbd.cBuffers = 4;
    sbd.pBuffers = sbufs;
    sbd.ulVersion = SECBUFFER_VERSION;

    ss = EncryptMessage(&sech, qop, &sbd, 0);
    if (ss != SEC_E_OK) {
        int i;
        Twapi_AppendSystemError(interp, ss);
        for (i=0; i < ARRAYSIZE(objs); ++i)
            ObjDecrRefs(objs[i]);
    } else {
        /* TBD - ensure that only length of trailer has changed. Else
           we have to copy each fragment separately into a new buffer */
        Tcl_SetByteArrayLength(objs[0], sbufs[0].cbBuffer + sbufs[1].cbBuffer + sbufs[2].cbBuffer);
        ObjSetResult(interp, ObjNewList(ARRAYSIZE(objs), objs));
    }

    return ss == SEC_E_OK ? TCL_OK : TCL_ERROR;
}


static TCL_RESULT Twapi_DecryptStreamObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    SecHandle sech;
    SECURITY_STATUS ss;
    SecBuffer sbufs[4];
    SecBufferDesc sbd;
    Tcl_Obj *objs[3];           /* 0 status, 1 decrypted data, 2 extra data */
    char *encP, *p;
    int  i, enclen;
    TCL_RESULT res;

    CHECK_NARGS_RANGE(interp, objc, 3, INT_MAX);
    if (ObjToSecHandle(interp, objv[1], &sech) != TCL_OK)
        return TCL_ERROR;

    sbufs[0].BufferType = SECBUFFER_DATA;
    sbufs[0].cbBuffer   = 0;
    for (i = 2; i < objc; ++i) {
        ObjToByteArray(objv[i], &enclen);
        sbufs[0].cbBuffer += enclen;
    }
    sbufs[0].pvBuffer = MemLifoPushFrame(ticP->memlifoP,
                                         sbufs[0].cbBuffer, NULL);
    p = sbufs[0].pvBuffer;
    for (i = 2; i < objc; ++i) {
        encP = ObjToByteArray(objv[i], &enclen);
        CopyMemory(p, encP, enclen);
        p += enclen;
    }
    TWAPI_ASSERT(p == (sbufs[0].cbBuffer + (char*) sbufs[0].pvBuffer));

    /* Note DecryptMessage decrypts in place */

    for (i = 1; i < ARRAYSIZE(sbufs); ++i)
        sbufs[i].BufferType = SECBUFFER_EMPTY;

    sbd.cBuffers = 4;
    sbd.pBuffers = sbufs;
    sbd.ulVersion = SECBUFFER_VERSION;

    /* DecryptMessage will not preserve values in sbufs[0] so remember them */
    encP = sbufs[0].pvBuffer;
    enclen = sbufs[0].cbBuffer;

    ss = DecryptMessage(&sech, &sbd, 0, NULL);
    switch (ss) {
    case SEC_E_INCOMPLETE_MESSAGE:
        objs[0] = STRING_LITERAL_OBJ("incomplete_message");
        objs[1] = ObjFromEmptyString();
        objs[2] = ObjFromByteArray(encP, enclen);
        ObjSetResult(interp, ObjNewList(ARRAYSIZE(objs), objs));
        res = TCL_OK;
        break;

    case SEC_E_OK:
    case SEC_I_CONTEXT_EXPIRED:
    case SEC_I_RENEGOTIATE:
        /* We do not know which of the output buffers 1-3 contain
         * decrypted data and extra data
         * TBD - Is handling of context_expired and renegoitate correct ?
        */
        for (i = 1; i < ARRAYSIZE(sbufs); ++i) {
            if (sbufs[i].BufferType == SECBUFFER_DATA)
                break;
        }
        if (i < ARRAYSIZE(sbufs))
            objs[1] = ObjFromByteArray(sbufs[i].pvBuffer, sbufs[i].cbBuffer);
        else
            objs[1] = ObjFromEmptyString(); /* No leftover data */

        for (i = 1; i < ARRAYSIZE(sbufs); ++i) {
            if (sbufs[i].BufferType == SECBUFFER_EXTRA)
                break;
        }
        if (i < ARRAYSIZE(sbufs)) {
            TWAPI_ASSERT(enclen > sbufs[i].cbBuffer);
            /* MSDN docs say "pvBuffer" does not contain a copy of the
               data. What they probably mean is that it points into
               the original data, not to a copy of it. But the sample
               code in MSDN and curl etc. does not make use of this field
               so neither do we, just to be safe. Instead calculate base
               on counts and the original data pointer
            */
            objs[2] = ObjFromByteArray(encP + (enclen - sbufs[i].cbBuffer), sbufs[i].cbBuffer);
        } else
            objs[2] = ObjFromEmptyString();

        switch (ss) {
        case SEC_E_OK:
            objs[0] = STRING_LITERAL_OBJ("ok");
            break;
        case SEC_I_CONTEXT_EXPIRED:
            objs[0] = STRING_LITERAL_OBJ("expired");
            break;
        case SEC_I_RENEGOTIATE:
            objs[0] = STRING_LITERAL_OBJ("renegotiate");
            break;
        }
        ObjSetResult(interp, ObjNewList(ARRAYSIZE(objs), objs));
        res = TCL_OK;
        break;

    default:
        res = Twapi_AppendSystemError(ticP->interp, ss);
        break;
    }

    MemLifoPopFrame(ticP->memlifoP);
    return res;
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
    int is_unisp, n;
    void *pv;

    pv = NULL;
    mark = MemLifoPushMark(ticP->memlifoP);
    luidP = &luid;
    if (TwapiGetArgsEx(ticP, objc-1, objv+1,
                       GETEMPTYASNULL(principalP), GETWSTR(packageP),
                       GETINT(cred_use), GETVAR(luidP, ObjToLUID_NULL),
                       GETOBJ(authObj), ARGEND) != TCL_OK)
        goto vamoose;

    if (ObjListLength(interp, authObj, &n) != TCL_OK)
        goto vamoose;

    if (n == 0)
        pv = NULL;
    else {
        if (WSTREQ(packageP, UNISP_NAME_W) ||
            WSTREQ(packageP, SCHANNEL_NAME_W)) {
            if (ParseSCHANNEL_CRED(ticP, authObj, &(SCHANNEL_CRED *)pv) != TCL_OK)
                goto vamoose;
            is_unisp = 1;
        } else if (WSTREQ(packageP, WDIGEST_SP_NAME_W) ||
                   WSTREQ(packageP, NTLMSP_NAME) ||
                   WSTREQ(packageP, NEGOSSP_NAME_W) ||
                   WSTREQ(packageP, MICROSOFT_KERBEROS_NAME_W)) {
            /* TBD - SDK has the comment that for RPC this memory
               must be valid for the lifetime of the binding handle. Is
               that true for SSPI also. Test */
            if (ParseSEC_WINNT_AUTH_IDENTITY(ticP, authObj, &(SEC_WINNT_AUTH_IDENTITY_W *)pv) != TCL_OK)
                goto vamoose;
            is_unisp = 0;
        } else {
            TwapiReturnErrorEx(interp, TWAPI_UNSUPPORTED_TYPE,
                               Tcl_ObjPrintf("Non-default authentication is not supported for SSPI package %s.", ObjToString(objv[2])));
            goto vamoose;
        }
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
    DWORD dw, dw2;
    SecHandle sech;
    SecBuffer sb;
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
        case 105: // Twapi_ApplyControlToken_SCHANNEL_SHUTDOWN
            dw = SCHANNEL_SHUTDOWN;
            sb.pvBuffer = &dw;
            sb.BufferType = SECBUFFER_TOKEN;
            sb.cbBuffer = sizeof(dw);
            sbd.cBuffers = 1;
            sbd.pBuffers = &sb;
            sbd.ulVersion = SECBUFFER_VERSION;
            result.type = TRT_EXCEPTION_ON_ERROR;
            result.value.ival = ApplyControlToken(&sech, &sbd);
            break;
        }
    } else {
        /* Free-for-all - each func responsible for checking arguments */
        switch (func) {
        
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
    static struct fncode_dispatch_s SspiDispatch[] = {
        DEFINE_FNCODE_CMD(EnumerateSecurityPackages, 1),
        DEFINE_FNCODE_CMD(QuerySecurityContextToken, 101),
        DEFINE_FNCODE_CMD(FreeCredentialsHandle, 102),
        DEFINE_FNCODE_CMD(DeleteSecurityContext, 103),
        DEFINE_FNCODE_CMD(ImpersonateSecurityContext, 104),
        DEFINE_FNCODE_CMD(Twapi_ApplyControlToken_SCHANNEL_SHUTDOWN, 105),
        DEFINE_FNCODE_CMD(QueryContextAttributes, 10023),
        DEFINE_FNCODE_CMD(VerifySignature, 10024),
        DEFINE_FNCODE_CMD(DecryptMessage, 10025),
    };

    static struct tcl_dispatch_s TclDispatch[] = {
        DEFINE_TCL_CMD(AcquireCredentialsHandle, Twapi_AcquireCredentialsHandleObjCmd),
        DEFINE_TCL_CMD(InitializeSecurityContext, Twapi_InitializeSecurityContextObjCmd),
        DEFINE_TCL_CMD(AcceptSecurityContext, Twapi_AcceptSecurityContextObjCmd),
        DEFINE_TCL_CMD(EncryptMessage, Twapi_EncryptMessageObjCmd),
        DEFINE_TCL_CMD(MakeSignature, Twapi_MakeSignatureObjCmd),
        DEFINE_TCL_CMD(EncryptStream, Twapi_EncryptStreamObjCmd),
        DEFINE_TCL_CMD(DecryptStream, Twapi_DecryptStreamObjCmd),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(SspiDispatch), SspiDispatch, Twapi_SspiCallObjCmd);
    TwapiDefineTclCmds(interp, ARRAYSIZE(TclDispatch), TclDispatch, ticP);

    return TCL_OK;
}



