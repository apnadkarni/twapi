#ifndef PBKDF2_H
#define PBKDF2_H

// Pseudo Random Function (PRF) prototype

/* generic context used in HMAC calculation */
typedef struct
{
   DWORD	magic;			/* used to help check that we are using the correct context */
   void*	pParam;	      /* hold a custom pointer known to the implementation  */
} PRF_CTX;

typedef BOOL (WINAPI* PRF_HmacInitPtr)(
                           PRF_CTX*       pContext,   /* PRF context used in HMAC computation */
                           unsigned char* pbKey,      /* pointer to authentication key */
                           DWORD          cbKey       /* length of authentication key */                        
                           );

typedef BOOL (WINAPI* PRF_HmacPtr)(
                           PRF_CTX*       pContext,   /* PRF context initialized by HmacInit */
                           unsigned char*  pbData,    /* pointer to data stream */
                           DWORD          cbData,     /* length of data stream */                           
                           unsigned char* pbDigest    /* caller digest to be filled in */                           
                           );

typedef BOOL (WINAPI* PRF_HmacFreePtr)(
                           PRF_CTX*       pContext	/* PRF context initialized by HmacInit */                        
                           );


/* PRF type definition */
typedef struct
{
   PRF_HmacInitPtr   hmacInit;
   PRF_HmacPtr       hmac;
   PRF_HmacFreePtr	hmacFree;
   DWORD             cbHmacLength;
} PRF;

extern PRF sha1Prf;

BOOL PBKDF2(PRF pPrf,
            unsigned char* pbPassword,
            DWORD cbPassword,
            unsigned char* pbSalt,
            DWORD cbSalt,
            DWORD dwIterationCount,
            unsigned char* pbDerivedKey,
            DWORD          cbDerivedKey);
#endif
