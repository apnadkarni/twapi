#ifndef TWAPI_CRYPTO_H
#define TWAPI_CRYPTO_H


int TwapiSspiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP);
#ifndef UNISP_NAME_W
#include <schannel.h>            /* For VC6 */
#endif
#ifndef WDIGEST_SP_NAME_W
#include <wdigest.h>            /* For VC6 */
#endif

#ifndef X509_ALGORITHM_IDENTIFIER
# define X509_ALGORITHM_IDENTIFIER           ((LPCSTR) 74)
#endif

#ifndef szOID_SUBJECT_INFO_ACCESS
# define szOID_SUBJECT_INFO_ACCESS       "1.3.6.1.5.5.7.1.11"
#endif

#ifndef CRYPT_OID_INFO_OID_KEY_FLAGS_MASK
# define CRYPT_OID_INFO_OID_KEY_FLAGS_MASK           0xFFFF0000
# define CRYPT_OID_INFO_PUBKEY_SIGN_KEY_FLAG         0x80000000
# define CRYPT_OID_INFO_PUBKEY_ENCRYPT_KEY_FLAG      0x40000000
#endif

#ifndef CRYPT_OID_DISABLE_SEARCH_DS_FLAG
# define CRYPT_OID_DISABLE_SEARCH_DS_FLAG            0x80000000
#endif

void TwapiRegisterPCCERT_CONTEXT(Tcl_Interp *, PCCERT_CONTEXT);
void TwapiRegisterPCCERT_CONTEXTTic(TwapiInterpContext *, PCCERT_CONTEXT );
TCL_RESULT TwapiUnregisterPCCERT_CONTEXT(Tcl_Interp *, PCCERT_CONTEXT);
TCL_RESULT TwapiUnregisterPCCERT_CONTEXTTic(TwapiInterpContext *, PCCERT_CONTEXT);
Tcl_Obj *ObjFromCERT_NAME_BLOB(CERT_NAME_BLOB *blobP, DWORD flags);

#endif
