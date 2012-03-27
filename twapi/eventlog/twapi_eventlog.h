#ifndef TWAPI_EVENTLOG_H
#define TWAPI_EVENTLOG_H

void TwapiInitEvtStubs(Tcl_Interp *interp);
int Twapi_EvtCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);

#endif
