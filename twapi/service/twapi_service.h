#ifndef TWAPI_SERVICE_H
#define TWAPI_SERVICE_H

#include <winsvc.h>

int Twapi_BecomeAService(TwapiInterpContext *, int objc, Tcl_Obj *CONST objv[]);

int Twapi_SetServiceStatus(TwapiInterpContext *, int objc, Tcl_Obj *CONST objv[]);




#endif
