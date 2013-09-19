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

void TwapiRegisterCertPointer(Tcl_Interp *, PCCERT_CONTEXT);
void TwapiRegisterCertPointerTic(TwapiInterpContext *, PCCERT_CONTEXT );
TCL_RESULT TwapiUnregisterCertPointer(Tcl_Interp *, PCCERT_CONTEXT);
TCL_RESULT TwapiUnregisterCertPointerTic(TwapiInterpContext *, PCCERT_CONTEXT);

#endif
