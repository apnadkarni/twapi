#ifndef TWAPI_CRYPTO_H
#define TWAPI_CRYPTO_H


int TwapiSspiInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP);
#ifndef UNISP_NAME_W
#include <schnlsp.h>            /* For VC6 */
#endif
#ifndef WDIGEST_SP_NAME_W
#include <wdigest.h>            /* For VC6 */
#endif


#endif
