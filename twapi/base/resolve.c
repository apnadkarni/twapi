/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Routines for asynchronous IP resolution */

#include "twapi.h"

typedef struct _TwapiHostnameEvent {
    Tcl_Event tcl_ev;           /* Must be first field */
    TwapiInterpContext *ticP;
    TwapiId    id;             /* Passed from script as a request id */
    DWORD  status;         /* 0 -> success, else Win32 error code */
    union {
        struct addrinfo *addrinfolist; /* Returned by getaddrinfo, to be
                                          freed via freeaddrinfo
                                          Used for host->addr */
        char *hostname;      /* Tcl_Alloc'ed (used for addr->hostname) */
    };
    int family;                 /* AF_UNSPEC, AF_INET or AF_INET6 */
    char name[1];           /* Holds query for hostname->addr */
    /* VARIABLE SIZE SINCE name[] IS ARBITRARY SIZE */
} TwapiHostnameEvent;
/*
 * Macro to calculate struct size. Note terminating null and the sizeof
 * the name[] array cancel each other out. (namelen_) does not include
 * terminating null.
 */
#define SIZE_TwapiHostnameEvent(namelen_) \
    (sizeof(TwapiHostnameEvent) + (namelen_))



/* Called from the Tcl event loop with the result of a hostname lookup */
static int TwapiHostnameEventProc(Tcl_Event *tclevP, int flags)
{
    TwapiHostnameEvent *theP = (TwapiHostnameEvent *) tclevP;

    if (theP->ticP->interp != NULL &&
        ! Tcl_InterpDeleted(theP->ticP->interp)) {
        /* Invoke the script */
        Tcl_Interp *interp = theP->ticP->interp;
        Tcl_Obj *objP = Tcl_NewListObj(0, NULL);

        Tcl_ListObjAppendElement(
            interp, objP, STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_hostname_resolve_handler"));
        Tcl_ListObjAppendElement(interp, objP, ObjFromTwapiId(theP->id));
        if (theP->status == ERROR_SUCCESS) {
            /* Success */
            Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("success"));
            Tcl_ListObjAppendElement(interp, objP, TwapiCollectAddrInfo(theP->addrinfolist, theP->family));
        } else {
            /* Failure */
            Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("fail"));
            Tcl_ListObjAppendElement(interp, objP,
                                     Tcl_NewLongObj(theP->status));
        }
        /* Invoke the script */
        Tcl_IncrRefCount(objP);
        (void) Tcl_EvalObjEx(interp, objP, TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);
        Tcl_DecrRefCount(objP);
        /* TBD - check for error and add to background ? */
    }

    /* Done with the interp context */
    TwapiInterpContextUnref(theP->ticP, 1);

    /* Assumes we can free this from different thread than allocated it ! */
    if (theP->addrinfolist)
        freeaddrinfo(theP->addrinfolist);

    return 1;                   /* So Tcl removes from queue */
}


/* Called from the Win2000 thread pool */
static DWORD WINAPI TwapiHostnameHandler(TwapiHostnameEvent *theP)
{
    struct addrinfo hints;

    TwapiZeroMemory(&hints, sizeof(hints));
    hints.ai_family = theP->family;

    theP->tcl_ev.proc = TwapiHostnameEventProc;
    theP->status = getaddrinfo(theP->name, "0", &hints, &theP->addrinfolist);
    TwapiEnqueueTclEvent(theP->ticP, &theP->tcl_ev);
    return 0;               /* Return value does not matter */
}


int Twapi_ResolveHostnameAsync(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    TwapiId id;
    char *name;
    int   len;
    TwapiHostnameEvent *theP;
    DWORD winerr;
    int family;

    ERROR_IF_UNTHREADED(ticP->interp);

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETASTRN(name, len), ARGUSEDEFAULT, GETINT(family),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    id =  TWAPI_NEWID(ticP);
    /* Allocate the callback context, must be allocated via Tcl_Alloc
     * as it will be passed to Tcl_QueueEvent.
     */
    theP = (TwapiHostnameEvent *) Tcl_Alloc(SIZE_TwapiHostnameEvent(len));
    theP->tcl_ev.proc = NULL;
    theP->tcl_ev.nextPtr = NULL;
    theP->id = id;
    theP->status = ERROR_SUCCESS;
    theP->ticP = ticP;
    TwapiInterpContextRef(ticP, 1); /* So it does not go away */
    theP->addrinfolist = NULL;
    theP->family = family;
    CopyMemory(theP->name, name, len+1);

    if (QueueUserWorkItem(TwapiHostnameHandler, theP, WT_EXECUTEDEFAULT)) {
        Tcl_SetObjResult(ticP->interp, ObjFromTwapiId(id));
        return TCL_OK;
    }

    winerr = GetLastError();    /* Remember the error */

    TwapiInterpContextUnref(ticP, 1); /* Undo above ref */
    Tcl_Free((char*) theP);
    return Twapi_AppendSystemError(ticP->interp, winerr);
}


/* Called from the Tcl event loop with the result of a address lookup */
static int TwapiAddressEventProc(Tcl_Event *tclevP, int flags)
{
    TwapiHostnameEvent *theP = (TwapiHostnameEvent *) tclevP;

    if (theP->ticP->interp != NULL &&
        ! Tcl_InterpDeleted(theP->ticP->interp)) {
        /* Invoke the script */
        Tcl_Interp *interp = theP->ticP->interp;
        Tcl_Obj *objP = Tcl_NewListObj(0, NULL);

        Tcl_ListObjAppendElement(
            interp, objP, STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_address_resolve_handler"));
        Tcl_ListObjAppendElement(interp, objP, ObjFromTwapiId(theP->id));
        if (theP->status == ERROR_SUCCESS) {
            /* Success. Note theP->hostname may still be NULL */
            Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("success"));
            Tcl_ListObjAppendElement(
                interp, objP,
                Tcl_NewStringObj((theP->hostname ? theP->hostname : ""), -1));
        } else {
            /* Failure */
            Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("fail"));
            Tcl_ListObjAppendElement(interp, objP,
                                     Tcl_NewLongObj(theP->status));
        }
        /* Invoke the script */
        /* Do we need TclSave/RestoreResult ? */
        Tcl_IncrRefCount(objP);
        (void) Tcl_EvalObjEx(interp, objP, TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);
        Tcl_DecrRefCount(objP);
        /* TBD - check for error and add to background ? */
    }
    
    /* Done with the interp context */
    TwapiInterpContextUnref(theP->ticP, 1);
    if (theP->hostname)
        TwapiFree(theP->hostname);
    
    return 1;                   /* So Tcl removes from queue */
}


/* Called from the Win2000 thread pool */
static DWORD WINAPI TwapiAddressHandler(TwapiHostnameEvent *theP)
{
    SOCKADDR_STORAGE ss;
    char hostname[NI_MAXHOST];
    char portname[NI_MAXSERV];
    int family;

    theP->tcl_ev.proc = TwapiAddressEventProc;
    family = TwapiStringToSOCKADDR_STORAGE(theP->name, &ss, theP->family);
    if (family == AF_UNSPEC) {
        // Fail, invalid address string
        theP->status = 10022;         /* WSAINVAL error code */
    } else {    
        theP->status = getnameinfo((struct sockaddr *)&ss,
                                   ss.ss_family == AF_INET6 ? sizeof(SOCKADDR_IN6) : sizeof(SOCKADDR_IN),
                                   hostname, sizeof(hostname)/sizeof(hostname[0]),
                                   portname, sizeof(portname)/sizeof(portname[0]),
                                   NI_NUMERICSERV);
    }
    if (theP->status == 0) {
        /* If the function just returned back the address, then there
           was really no name found so return empty string (NULL) */
        theP->hostname = NULL;
        if (lstrcmpA(theP->name, hostname)) {
            /* Really do have a name */
            theP->hostname = TwapiAllocAString(hostname, -1);
        }
    }

    TwapiEnqueueTclEvent(theP->ticP, &theP->tcl_ev);
    return 0;                   /* Return value ignored anyways */
}

int Twapi_ResolveAddressAsync(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    TwapiId id;
    char *addrstr;
    int   len;
    TwapiHostnameEvent *theP;
    DWORD winerr;
    int family;

    ERROR_IF_UNTHREADED(ticP->interp);

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETASTRN(addrstr, len), ARGUSEDEFAULT, GETINT(family),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    id =  TWAPI_NEWID(ticP);

    /* Allocate the callback context, must be allocated via Tcl_Alloc
     * as it will be passed to Tcl_QueueEvent.
     */
    theP = (TwapiHostnameEvent *) Tcl_Alloc(SIZE_TwapiHostnameEvent(len));
    theP->tcl_ev.proc = NULL;
    theP->tcl_ev.nextPtr = NULL;
    theP->id = id;
    theP->status = ERROR_SUCCESS;
    theP->ticP = ticP;
    TwapiInterpContextRef(ticP, 1); /* So it does not go away */
    theP->hostname = NULL;
    theP->family = family;

    /* We do not syntactically validate address string here. All failures
       are delivered asynchronously */
    CopyMemory(theP->name, addrstr, len+1);

    if (QueueUserWorkItem(TwapiAddressHandler, theP, WT_EXECUTEDEFAULT)) {
        Tcl_SetObjResult(ticP->interp, ObjFromTwapiId(id));
        return TCL_OK;
    }
    winerr = GetLastError();    /* Remember the error */

    TwapiInterpContextUnref(ticP, 1); /* Undo above ref */
    Tcl_Free((char*) theP);
    return Twapi_AppendSystemError(ticP->interp, winerr);
}
