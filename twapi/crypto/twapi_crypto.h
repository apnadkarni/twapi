#ifndef TWAPI_CRYPTO_H
#define TWAPI_CRYPTO_H


int TwapiSspiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP);
#ifndef UNISP_NAME_W
#include <schannel.h>            /* For VC6 */
#endif
#ifndef WDIGEST_SP_NAME_W  /* Needed for MingW and VC6 */
#define WDIGEST_SP_NAME_W             L"WDigest"
#endif
#include <softpub.h>            /* WinVerifyTrust */

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

/* VC6 Wintrust.h does not define the dwUIContext field so define our
   own version of the structure
*/
#if _MSC_VER <= 1400
typedef struct _TWAPI_WINTRUST_DATA
{
    DWORD           cbStruct;                   // = sizeof(WINTRUST_DATA)

    LPVOID          pPolicyCallbackData;        // optional: used to pass data between the app and policy
    LPVOID          pSIPClientData;             // optional: used to pass data between the app and SIP.

    DWORD           dwUIChoice;
    DWORD           fdwRevocationChecks;
    DWORD           dwUnionChoice;
    union
    {
        struct WINTRUST_FILE_INFO_      *pFile;
        struct WINTRUST_CATALOG_INFO_   *pCatalog;
        struct WINTRUST_BLOB_INFO_      *pBlob;
        struct WINTRUST_SGNR_INFO_      *pSgnr;
        struct WINTRUST_CERT_INFO_      *pCert;
    };
    DWORD           dwStateAction;
    HANDLE          hWVTStateData;
    WCHAR           *pwszURLReference;
    DWORD           dwProvFlags;
    DWORD           dwUIContext;
} TWAPI_WINTRUST_DATA, *PTWAPI_WINTRUST_DATA;
#else
typedef WINTRUST_DATA TWAPI_WINTRUST_DATA;
typedef PWINTRUST_DATA PTWAPI_WINTRUST_DATA;
#endif

typedef struct _TWAPI_PLAINTEXTKEYBLOB {
    BLOBHEADER hdr;
    DWORD dwKeySize;
    BYTE  rgbKeyData[1]; /* Actually [dwKeySize] */
} TWAPI_PLAINTEXTKEYBLOB;
#define TWAPI_PLAINTEXTKEYBLOB_SIZE(klen_) \
    ((klen_) + offsetof(struct _TWAPI_PLAINTEXTKEYBLOB, rgbKeyData))

void TwapiRegisterPCCERT_CONTEXT(Tcl_Interp *, PCCERT_CONTEXT);
void TwapiRegisterPCCERT_CONTEXTTic(TwapiInterpContext *, PCCERT_CONTEXT );
TCL_RESULT TwapiUnregisterPCCERT_CONTEXT(Tcl_Interp *, PCCERT_CONTEXT);
TCL_RESULT TwapiUnregisterPCCERT_CONTEXTTic(TwapiInterpContext *, PCCERT_CONTEXT);
Tcl_Obj *ObjFromCERT_NAME_BLOB(CERT_NAME_BLOB *blobP, DWORD flags);

#endif
