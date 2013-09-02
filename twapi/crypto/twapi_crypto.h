#ifndef TWAPI_CRYPTO_H
#define TWAPI_CRYPTO_H


int TwapiSspiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP);
#ifndef UNISP_NAME_W
#include <schannel.h>            /* For VC6 */
#endif
#ifndef WDIGEST_SP_NAME_W
#include <wdigest.h>            /* For VC6 */
#endif

void TwapiRegisterCertPointer(Tcl_Interp *, PCCERT_CONTEXT);
void TwapiRegisterCertPointerTic(TwapiInterpContext *, PCCERT_CONTEXT );
TCL_RESULT TwapiUnregisterCertPointer(Tcl_Interp *, PCCERT_CONTEXT);
TCL_RESULT TwapiUnregisterCertPointerTic(TwapiInterpContext *, PCCERT_CONTEXT);

#endif
