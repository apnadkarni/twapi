/*
 * An implementation of PBKDF2 (Password-Based Key Derivation Function) using Crypto API 
 * (RFC2898 / PKCS#5)
 *
 * Copyright (c) 2010 Mounir IDRASSI <mounir.idrassi@idrix.fr>. All rights reserved.
 *
 * This program is distributed in the hope that it will be useful, 
 * but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
 * or FITNESS FOR A PARTICULAR PURPOSE.
 * 
 */

#ifndef _WIN32_WINNT         
#define _WIN32_WINNT 0x0501
#endif

#include <windows.h>
#include <wincrypt.h>
#include <stdio.h>
#include <math.h>

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


/* Implementation of HMAC-SHA1 using CAPI */

#define HMAC_SHA1_MAGIC 0x53484131

typedef struct
{
   HCRYPTPROV hProv;
   HCRYPTKEY hKey;
} CAPI_CTX_PARAM;

// Structure used by CAPI for HMAC computation
typedef struct {
   BLOBHEADER hdr;
   DWORD cbKeySize;
} HMAC_KEY_BLOB;

BOOL WINAPI hmacInit_sha1(
   PRF_CTX*       pContext,   /* PRF context used in HMAC computation */
   unsigned char* pbKey,      /* pointer to authentication key */
   DWORD          cbKey       /* length of authentication key */                        
)
{
   HCRYPTPROV hProv = 0;
   HCRYPTKEY hKey = 0;
   HMAC_KEY_BLOB *pKeyBlob = (HMAC_KEY_BLOB *) LocalAlloc(0, sizeof(HMAC_KEY_BLOB) + cbKey + 20); // we put enough room for 0's padding
   BOOL bStatus = FALSE;
   DWORD dwError = 0, dwLen;
   CAPI_CTX_PARAM* pParam = NULL;
   
   pKeyBlob->hdr.bType = PLAINTEXTKEYBLOB;
   pKeyBlob->hdr.bVersion = CUR_BLOB_VERSION;
   pKeyBlob->hdr.reserved = 0;
   pKeyBlob->hdr.aiKeyAlg = CALG_RC2;
   pKeyBlob->cbKeySize = cbKey;
   memcpy(((LPBYTE) pKeyBlob) + sizeof(HMAC_KEY_BLOB), pbKey, cbKey);

   if (!pContext)
   {
      dwError = ERROR_BAD_ARGUMENTS;
      goto hmacInit_end;
   }

   if (!CryptAcquireContext(&hProv, NULL, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)) 
   {
      dwError = GetLastError();
      goto hmacInit_end;
   }

   dwLen = sizeof(HMAC_KEY_BLOB) + cbKey;
   if (dwLen < 20)
   {
      // we pad with zeros till the size of SHA1 digest.
      // this is to avoid erros under Windows 8.1 wich doesn't accept 1-byte RC2 keys
	  // the result will be the same since the HMAC-SHA1 will perform the same padding
	  DWORD dwPad = 20 - dwLen;
	  memset(((LPBYTE) pKeyBlob) + dwLen, 0, dwPad);
	  dwLen += dwPad;
	  pKeyBlob->cbKeySize += dwPad;
   }

   if (!CryptImportKey(hProv, (LPBYTE) pKeyBlob, dwLen, 0, CRYPT_IPSEC_HMAC_KEY, &hKey))
   {
      dwError = GetLastError();
      goto hmacInit_end;
   }

   pParam = (CAPI_CTX_PARAM*) LocalAlloc(0, sizeof(CAPI_CTX_PARAM));
   pParam->hProv = hProv;
   pParam->hKey = hKey;

   pContext->magic = HMAC_SHA1_MAGIC;
   pContext->pParam = (void*) pParam;

   hProv = 0;
   hKey = 0;

   bStatus = TRUE;

hmacInit_end:

   if (hKey) CryptDestroyKey(hKey);
   if (hProv) CryptReleaseContext(hProv, 0);

   if (pKeyBlob) LocalFree(pKeyBlob);

   SetLastError(dwError);
   return bStatus;
}

BOOL WINAPI hmac_sha1(
   PRF_CTX*       pContext,               /* PRF context used in HMAC computation */  
   unsigned char*  pbData,                /* pointer to data stream */
   DWORD           cbData,                /* length of data stream */
   unsigned char   pbDigest[20]           /* caller digest to be filled in */
)
{
   HCRYPTPROV hProv = 0;
   HCRYPTHASH hHash = 0;
   HCRYPTKEY hKey = 0;
   DWORD cbDigest = 20;
   HMAC_INFO   HmacInfo;
   BOOL bStatus = FALSE;
   DWORD dwError = 0;

   ZeroMemory(&HmacInfo, sizeof(HmacInfo));
   HmacInfo.HashAlgid = CALG_SHA1;

   if (!pContext || (pContext->magic != HMAC_SHA1_MAGIC) || (!pContext->pParam))
   {
      dwError = ERROR_BAD_ARGUMENTS;
      goto hmac_end;
   }

   hProv = ((CAPI_CTX_PARAM*) pContext->pParam)->hProv;
   hKey = ((CAPI_CTX_PARAM*) pContext->pParam)->hKey;

   if (!CryptCreateHash(hProv, CALG_HMAC, hKey, 0, &hHash))
   {
      dwError = GetLastError();
      goto hmac_end;
   }

   if (!CryptSetHashParam( hHash, HP_HMAC_INFO, (BYTE*)&HmacInfo, 0))
   {
      dwError = GetLastError();
      goto hmac_end;
   }

   if (!CryptHashData(hHash, pbData, cbData, 0))
   {
      dwError = GetLastError();
      goto hmac_end;
   }

   if (!CryptGetHashParam(hHash, HP_HASHVAL, pbDigest, &cbDigest, 0))
   {
      dwError = GetLastError();
      goto hmac_end;
   }

   bStatus = TRUE;

hmac_end:

   if (hHash) CryptDestroyHash(hHash);

   SetLastError(dwError);
   return bStatus;
}


BOOL WINAPI hmacFree_sha1(
   PRF_CTX*       pContext          /* PRF context used in HMAC computation */  
)
{
   HCRYPTPROV hProv = 0;
   HCRYPTKEY hKey = 0;

   if (!pContext || (pContext->magic != HMAC_SHA1_MAGIC) || (!pContext->pParam))
   {
      SetLastError(ERROR_BAD_ARGUMENTS);
      return FALSE;
   }

   hProv = ((CAPI_CTX_PARAM*) pContext->pParam)->hProv;
   hKey = ((CAPI_CTX_PARAM*) pContext->pParam)->hKey;

   CryptDestroyKey(hKey);
   CryptReleaseContext(hProv, 0);

   LocalFree(pContext->pParam);
   SecureZeroMemory(pContext, sizeof(PRF_CTX));

   return TRUE;
}

/*
 * Definition of the HMAC-SHA1 PRF
 */
PRF sha1Prf = {hmacInit_sha1, hmac_sha1, hmacFree_sha1, 20};


static __inline void xor(LPBYTE ptr1, LPBYTE ptr2, DWORD dwLen)
{
   if (dwLen)
      while (dwLen--) *ptr1++ ^= *ptr2++;
}

/*
 * PBKDF2 implementation
 */
BOOL PBKDF2(PRF pPrf,
            unsigned char* pbPassword,
            DWORD cbPassword,
            unsigned char* pbSalt,
            DWORD cbSalt,
            DWORD dwIterationCount,
            unsigned char* pbDerivedKey,
            DWORD          cbDerivedKey)
{
   BOOL bStatus = FALSE;
   DWORD dwError = 0;
   DWORD l, r, i,j;
   DWORD hlen = pPrf.cbHmacLength;
   LPBYTE Ti = (LPBYTE) LocalAlloc(0, hlen);
   LPBYTE V = (LPBYTE) LocalAlloc(0, hlen);
   LPBYTE U = (LPBYTE) LocalAlloc(0, max((cbSalt + 4), hlen));
   DWORD dwULen;
   PRF_CTX prfCtx = {0};

   if (!pbDerivedKey || !cbDerivedKey || (!pbPassword && cbPassword) )
   {
      dwError = ERROR_BAD_ARGUMENTS;
      goto PBKDF2_end;
   }

   if (!Ti || !U || !V)
   {
      dwError = ERROR_NOT_ENOUGH_MEMORY;
      goto PBKDF2_end;
   }

   l = (DWORD) ceil((double) cbDerivedKey / (double) hlen);
   r = cbDerivedKey - (l - 1) * hlen;

   if (!pPrf.hmacInit(&prfCtx, pbPassword, cbPassword))
   {
      dwError = GetLastError();
      goto PBKDF2_end;
   }

   for (i = 1; i <= l; i++)
   {
      ZeroMemory(Ti, hlen);
      for (j = 0; j < dwIterationCount; j++)
      {
         if (j == 0)
         {
            // construct first input for PRF
            memcpy(U, pbSalt, cbSalt);
            U[cbSalt] = (BYTE) ((i & 0xFF000000) >> 24);
            U[cbSalt + 1] = (BYTE) ((i & 0x00FF0000) >> 16);
            U[cbSalt + 2] = (BYTE) ((i & 0x0000FF00) >> 8);
            U[cbSalt + 3] = (BYTE) ((i & 0x000000FF));
            dwULen = cbSalt + 4;
         }
         else
         {
            memcpy(U, V, hlen);
            dwULen = hlen;
         }

         if (!pPrf.hmac(&prfCtx, U, dwULen, V))
         {
            dwError = GetLastError();
            goto PBKDF2_end;
         }

         xor(Ti, V, hlen);
      }

      if (i != l)
      {
         memcpy(&pbDerivedKey[(i-1) * hlen], Ti, hlen);
      }
      else
      {
         // Take only the first r bytes
         memcpy(&pbDerivedKey[(i-1) * hlen], Ti, r);
      }
   }

   bStatus = TRUE;

PBKDF2_end:

   pPrf.hmacFree(&prfCtx);

   if (Ti) LocalFree(Ti);
   if (U) LocalFree(U);
   if (V) LocalFree(V);
   SetLastError(dwError);
   return bStatus;
}
   
#if 0
int main(int argc, char* argv[])
{  
   unsigned char pbDerivedKey[32];

   // Test 1
   //  Input:
   //    P = "password" (8 bytes)
   //    S = "salt" (4 bytes)
   //    c = 1
   //   dkLen = 20
   //
   //  Output:
   //    DK = 0c 60 c8 0f 96 1f 0e 71 f3 a9 b5 24 af 60 12 06 2f e0 37 a6
   unsigned char pbExpectedKey1[20] = {0x0c, 0x60, 0xc8, 0x0f, 0x96, 0x1f, 0x0e, 0x71, 0xf3, 0xa9, 0xb5, 0x24, 0xaf, 0x60, 0x12, 0x06, 0x2f, 0xe0, 0x37, 0xa6};   
   // Test 2
   //  Input:
   //    P = "password" (8 bytes)
   //    S = "salt" (4 bytes)
   //    c = 2
   //   dkLen = 20
   //
   //  Output:
   //    DK = ea 6c 01 4d c7 2d 6f 8c cd 1e d9 2a ce 1d 41 f0 d8 de 89 57
   unsigned char pbExpectedKey2[20] = {0xea, 0x6c, 0x01, 0x4d, 0xc7, 0x2d, 0x6f, 0x8c, 0xcd, 0x1e, 0xd9, 0x2a, 0xce, 0x1d, 0x41, 0xf0, 0xd8, 0xde, 0x89, 0x57};   

   // Test 3
   //  Input:
   //    P = "password" (8 bytes)
   //    S = "salt" (4 bytes)
   //    c = 4096
   //   dkLen = 20
   //
   //  Output:
   //    DK = 4b 00 79 01 b7 65 48 9a be ad 49 d9 26 f7 21 d0 65 a4 29 c1
   unsigned char pbExpectedKey3[20] = {0x4b, 0x00, 0x79, 0x01, 0xb7, 0x65, 0x48, 0x9a, 0xbe, 0xad, 0x49, 0xd9, 0x26, 0xf7, 0x21, 0xd0, 0x65, 0xa4, 0x29, 0xc1};   
   // Test 4
   //  Input:
   //    P = "password" (8 bytes)
   //    S = 78 57 8E 5A 5D 63 CB 06 (8 bytes)
   //    c = 2048
   //   dkLen = 24
   //
   //  Output:
   //    DK = BF DE 6B E9 4D F7 E1 1D D4 09 BC E2 0A 02 55 EC 32 7C B9 36 FF E9 36 43
   unsigned char pbExpectedKey4[24] = {0xBF, 0xDE, 0x6B, 0xE9, 0x4D, 0xF7, 0xE1, 0x1D, 0xD4, 0x09, 0xBC, 0xE2, 0x0A, 0x02, 0x55, 0xEC, 0x32, 0x7C, 0xB9, 0x36, 0xFF, 0xE9, 0x36, 0x43};   
   unsigned char pbSalt[8] = { 0x78, 0x57, 0x8E, 0x5A, 0x5D, 0x63, 0xCB, 0x06};

   // Test 5
   //  Input:
   //    P = "Hello World" (11 bytes)
   //    S = DB 2F 38 A2 24 39 03 DE (8 bytes)
   //    c = 1000
   //   dkLen = 20
   //
   //  Output:
   //    DK = EC D1 54 2B 2E FB 3F 52 0F 12 A5 27 E3 86 B3 3C F4 AD E3 A2
   unsigned char pbExpectedKey5[20] = {0xEC, 0xD1, 0x54, 0x2B, 0x2E, 0xFB, 0x3F, 0x52, 0x0F, 0x12, 0xA5, 0x27, 0xE3, 0x86, 0xB3, 0x3C, 0xF4, 0xAD, 0xE3, 0xA2};      
   unsigned char pbOtherSalt[8] = {0xDB, 0x2F, 0x38, 0xA2, 0x24, 0x39, 0x03, 0xDE};

   // Test 6
   //  Input:
   //    P = "password" (8 bytes)
   //    S = "ATHENA.MIT.EDUraeburn" (21 bytes)
   //    c = 1
   //   dkLen = 32
   //
   //  Output:
   //    DK =  0xcd, 0xed, 0xb5, 0x28, 0x1b, 0xb2, 0xf8, 0x01,
	//		      0x56, 0x5a, 0x11, 0x22, 0xb2, 0x56, 0x35, 0x15,
	//		      0x0a, 0xd1, 0xf7, 0xa0, 0x4b, 0xb9, 0xf3, 0xa3,
	//		      0x33, 0xec, 0xc0, 0xe2, 0xe1, 0xf7, 0x08, 0x37

   unsigned char pbExpectedKey6[32] = {0xcd, 0xed, 0xb5, 0x28, 0x1b, 0xb2, 0xf8, 0x01,
			                              0x56, 0x5a, 0x11, 0x22, 0xb2, 0x56, 0x35, 0x15,
			                              0x0a, 0xd1, 0xf7, 0xa0, 0x4b, 0xb9, 0xf3, 0xa3,
			                              0x33, 0xec, 0xc0, 0xe2, 0xe1, 0xf7, 0x08, 0x37
                                       };

   // Test 7
   //  Input:
   //    P = "password" (8 bytes)
   //    S = "ATHENA.MIT.EDUraeburn" (21 bytes)
   //    c = 2
   //   dkLen = 32
   //
   //  Output:
   //    DK =  0x01, 0xdb, 0xee, 0x7f, 0x4a, 0x9e, 0x24, 0x3e, 
	//		      0x98, 0x8b, 0x62, 0xc7, 0x3c, 0xda, 0x93, 0x5d,
	//		      0xa0, 0x53, 0x78, 0xb9, 0x32, 0x44, 0xec, 0x8f,
	//		      0x48, 0xa9, 0x9e, 0x61, 0xad, 0x79, 0x9d, 0x86


   unsigned char pbExpectedKey7[32] = {0x01, 0xdb, 0xee, 0x7f, 0x4a, 0x9e, 0x24, 0x3e, 
	                        		      0x98, 0x8b, 0x62, 0xc7, 0x3c, 0xda, 0x93, 0x5d,
	                        		      0xa0, 0x53, 0x78, 0xb9, 0x32, 0x44, 0xec, 0x8f,
	                        		      0x48, 0xa9, 0x9e, 0x61, 0xad, 0x79, 0x9d, 0x86
                                       };

   // Test 8
   //  Input:
   //    P = "password" (8 bytes)
   //    S = "ATHENA.MIT.EDUraeburn" (21 bytes)
   //    c = 1200
   //   dkLen = 32
   //
   //  Output:
   //    DK =  0x5c, 0x08, 0xeb, 0x61, 0xfd, 0xf7, 0x1e, 0x4e,
   //			   0x4e, 0xc3, 0xcf, 0x6b, 0xa1, 0xf5, 0x51, 0x2b,
	//		      0xa7, 0xe5, 0x2d, 0xdb, 0xc5, 0xe5, 0x14, 0x2f,
	//		      0x70, 0x8a, 0x31, 0xe2, 0xe6, 0x2b, 0x1e, 0x13



   unsigned char pbExpectedKey8[32] = {0x5c, 0x08, 0xeb, 0x61, 0xfd, 0xf7, 0x1e, 0x4e,
			                              0x4e, 0xc3, 0xcf, 0x6b, 0xa1, 0xf5, 0x51, 0x2b,
			                              0xa7, 0xe5, 0x2d, 0xdb, 0xc5, 0xe5, 0x14, 0x2f,
			                              0x70, 0x8a, 0x31, 0xe2, 0xe6, 0x2b, 0x1e, 0x13
                                       };

   // Test 9
   //  Input:
   //    P = "password" (8 bytes)
   //    S = "\0224VxxV4\022" (8 bytes)
   //    c = 5
   //   dkLen = 32
   //
   //  Output:
   //    DK =  0xd1, 0xda, 0xa7, 0x86, 0x15, 0xf2, 0x87, 0xe6,
	//		      0xa1, 0xc8, 0xb1, 0x20, 0xd7, 0x06, 0x2a, 0x49,
	//		      0x3f, 0x98, 0xd2, 0x03, 0xe6, 0xbe, 0x49, 0xa6,
	//		      0xad, 0xf4, 0xfa, 0x57, 0x4b, 0x6e, 0x64, 0xee




   unsigned char pbExpectedKey9[32] = {0xd1, 0xda, 0xa7, 0x86, 0x15, 0xf2, 0x87, 0xe6,
			                              0xa1, 0xc8, 0xb1, 0x20, 0xd7, 0x06, 0x2a, 0x49,
			                              0x3f, 0x98, 0xd2, 0x03, 0xe6, 0xbe, 0x49, 0xa6,
			                              0xad, 0xf4, 0xfa, 0x57, 0x4b, 0x6e, 0x64, 0xee
                                       };

   // Test 10
   //  Input:
   //    P = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" (64 bytes)
   //    S = "pass phrase equals block size" (29 bytes)
   //    c = 1200
   //   dkLen = 32
   //
   //  Output:
   //    DK =  0x13, 0x9c, 0x30, 0xc0, 0x96, 0x6b, 0xc3, 0x2b,
	//		      0xa5, 0x5f, 0xdb, 0xf2, 0x12, 0x53, 0x0a, 0xc9,
	//		      0xc5, 0xec, 0x59, 0xf1, 0xa4, 0x52, 0xf5, 0xcc,
	//		      0x9a, 0xd9, 0x40, 0xfe, 0xa0, 0x59, 0x8e, 0xd1
   unsigned char pbExpectedKey10[32] = {0x13, 0x9c, 0x30, 0xc0, 0x96, 0x6b, 0xc3, 0x2b,
			                              0xa5, 0x5f, 0xdb, 0xf2, 0x12, 0x53, 0x0a, 0xc9,
			                              0xc5, 0xec, 0x59, 0xf1, 0xa4, 0x52, 0xf5, 0xcc,
			                              0x9a, 0xd9, 0x40, 0xfe, 0xa0, 0x59, 0x8e, 0xd1
                                       };

   // Test 11
   //  Input:
   //    P = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" (65 bytes)
   //    S = "pass phrase exceeds block size" (30 bytes)
   //    c = 1200
   //   dkLen = 32
   //
   //  Output:
   //    DK =  0x9c, 0xca, 0xd6, 0xd4, 0x68, 0x77, 0x0c, 0xd5,
	//		      0x1b, 0x10, 0xe6, 0xa6, 0x87, 0x21, 0xbe, 0x61,
	//		      0x1a, 0x8b, 0x4d, 0x28, 0x26, 0x01, 0xdb, 0x3b,
	//		      0x36, 0xbe, 0x92, 0x46, 0x91, 0x5e, 0xc8, 0x2a

   unsigned char pbExpectedKey11[32] = {0x9c, 0xca, 0xd6, 0xd4, 0x68, 0x77, 0x0c, 0xd5,
		                                 0x1b, 0x10, 0xe6, 0xa6, 0x87, 0x21, 0xbe, 0x61,
		                                 0x1a, 0x8b, 0x4d, 0x28, 0x26, 0x01, 0xdb, 0x3b,
		                                 0x36, 0xbe, 0x92, 0x46, 0x91, 0x5e, 0xc8, 0x2a
                                       };

   // Test 12
   //  Input:
   //    P = "\360\235\204\236" (4 bytes)
   //    S = "EXAMPLE.COMpianist" (18 bytes)
   //    c = 50
   //   dkLen = 32
   //
   //  Output:
   //    DK =  0x6b, 0x9c, 0xf2, 0x6d, 0x45, 0x45, 0x5a, 0x43,
	//		      0xa5, 0xb8, 0xbb, 0x27, 0x6a, 0x40, 0x3b, 0x39,
	//		      0xe7, 0xfe, 0x37, 0xa0, 0xc4, 0x1e, 0x02, 0xc2,
	//		      0x81, 0xff, 0x30, 0x69, 0xe1, 0xe9, 0x4f, 0x52


   unsigned char pbExpectedKey12[32] = {0x6b, 0x9c, 0xf2, 0x6d, 0x45, 0x45, 0x5a, 0x43,
		                                 0xa5, 0xb8, 0xbb, 0x27, 0x6a, 0x40, 0x3b, 0x39,
		                                 0xe7, 0xfe, 0x37, 0xa0, 0xc4, 0x1e, 0x02, 0xc2,
		                                 0x81, 0xff, 0x30, 0x69, 0xe1, 0xe9, 0x4f, 0x52
                                       };

   
   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, (LPBYTE) "salt", 4, 1, pbDerivedKey, 20))
   {
      printf("Test 1 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey1, 20))
   {
      printf("Test 1 failed: invalid value for derived key\n");
      goto main_end;
   }

   
   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, (LPBYTE) "salt", 4, 2, pbDerivedKey, 20))
   {
      printf("Test 2 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey2, 20))
   {
      printf("Test 2 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, (LPBYTE) "salt", 4, 4096, pbDerivedKey, 20))
   {
      printf("Test 3 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey3, 20))
   {
      printf("Test 3 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, pbSalt, 8, 2048, pbDerivedKey, 24))
   {
      printf("Test 4 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey4, 24))
   {
      printf("Test 4 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "Hello World", 11, pbOtherSalt, 8, 1000, pbDerivedKey, 20))
   {
      printf("Test 5 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey5, 20))
   {
      printf("Test 5 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, (LPBYTE) "ATHENA.MIT.EDUraeburn", 21, 1, pbDerivedKey, 32))
   {
      printf("Test 6 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey6, 32))
   {
      printf("Test 6 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, (LPBYTE) "ATHENA.MIT.EDUraeburn", 21, 2, pbDerivedKey, 32))
   {
      printf("Test 7 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey7, 32))
   {
      printf("Test 7 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, (LPBYTE) "ATHENA.MIT.EDUraeburn", 21, 1200, pbDerivedKey, 32))
   {
      printf("Test 8 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey8, 32))
   {
      printf("Test 8 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "password", 8, (LPBYTE) "\0224VxxV4\022", 8, 5, pbDerivedKey, 32))
   {
      printf("Test 9 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey9, 32))
   {
      printf("Test 9 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", 64, 
                        (LPBYTE) "pass phrase equals block size", 29, 1200, pbDerivedKey, 32))
   {
      printf("Test 10 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey10, 32))
   {
      printf("Test 10 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", 65, 
                        (LPBYTE) "pass phrase exceeds block size", 30, 1200, pbDerivedKey, 32))
   {
      printf("Test 11 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey11, 32))
   {
      printf("Test 11 failed: invalid value for derived key\n");
      goto main_end;
   }

   if (!PBKDF2(sha1Prf, (LPBYTE) "\360\235\204\236", 4, 
                        (LPBYTE) "EXAMPLE.COMpianist", 18, 50, pbDerivedKey, 32))
   {
      printf("Test 12 failed: PBKDF2 returned FALSE (Error 0x%.8X)\n", GetLastError());
      goto main_end;
   }
   if (0 != memcmp(pbDerivedKey, pbExpectedKey12, 32))
   {
      printf("Test 12 failed: invalid value for derived key\n");
      goto main_end;
   }

   printf("All tests OK!\n");


main_end:
	return 0;
}
#endif
