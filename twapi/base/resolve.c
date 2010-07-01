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
    int    id;             /* Passed from script as a request id */
    DWORD  status;         /* 0 -> success, else Win32 error code */
    union {
        struct in_addr *addrs;      /* Dynamically allocated using TwapiAlloc */
        char *hostname;      /* Ditto (used for addr->hostname) */
    };
    int    naddrs;              /* Number of addresses in addrs[] */
    char name[1];               /* Actually more */
} TwapiHostnameEvent;
/* Macro to calculate struct size. Note the terminating null (not included
   in namelen_) and existing space of name[] cancel each other out */
#define SIZE_TwapiHostnameEvent(namelen_)  \
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
        int i;

        Tcl_ListObjAppendElement(
            interp, objP, STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_hostname_resolve_handler"));
        Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(theP->id));
        if (theP->status == ERROR_SUCCESS) {
            /* Success */
            Tcl_Obj *addrsObj = Tcl_NewListObj(0, NULL);
            Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("success"));
            for (i=0; i < theP->naddrs; ++i) {
                Tcl_ListObjAppendElement(
                    interp, addrsObj,
                    Tcl_NewStringObj(inet_ntoa(theP->addrs[i]), -1));
            }
            Tcl_ListObjAppendElement(interp, objP, addrsObj);
        } else {
            /* Failure */
            Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("fail"));
            Tcl_ListObjAppendElement(interp, objP,
                                     Tcl_NewLongObj(theP->status));
        }
        /* Invoke the script */
        /* Do we need a Tcl_SaveResult/RestoreResult ? */
        Tcl_IncrRefCount(objP);
        (void) Tcl_EvalObjEx(interp, objP, TCL_EVAL_DIRECT|TCL_EVAL_GLOBAL);
        Tcl_DecrRefCount(objP);
        /* TBD - check for error and add to background ? */
    }

    /* Done with the interp context */
    TwapiInterpContextUnref(theP->ticP, 1);
    if (theP->addrs)
        TwapiFree(theP->addrs);

    return 1;                   /* So Tcl removes from queue */
}


/* Called from the Win2000 thread pool */
static DWORD WINAPI TwapiHostnameHandler(TwapiHostnameEvent *theP)
{
    struct addrinfo hints;
    struct addrinfo *addrP;
    struct addrinfo *saved_addrP;
    int i;

    memset(&hints, 0, sizeof(hints));
    hints.ai_family = PF_INET;

    theP->tcl_ev.proc = TwapiHostnameEventProc;
    theP->status = getaddrinfo(theP->name, "0", &hints, &addrP);
    if (theP->status) {
        TwapiEnqueueTclEvent(theP->ticP, &theP->tcl_ev);
        return 0;               /* Return value does not matter */
    }

    /* Loop and collect addresses. Assume at most 50 entries */
    theP->addrs = TwapiAlloc(50*sizeof(*(theP->addrs)));
    saved_addrP = addrP;
    for (i = 0; i < 50 && addrP; addrP = addrP->ai_next) {
        struct sockaddr_in *saddrP = (struct sockaddr_in *)addrP->ai_addr;
        if (addrP->ai_family != PF_INET ||
            addrP->ai_addrlen != sizeof(struct sockaddr_in) ||
            saddrP->sin_family != AF_INET) {
            /* Not IP V4 */
            continue;
        }
        theP->addrs[i++] = saddrP->sin_addr;
    }
    theP->naddrs = i;

    if (saved_addrP)
        freeaddrinfo(saved_addrP);

    TwapiEnqueueTclEvent(theP->ticP, &theP->tcl_ev);

    return 0;                   /* Return value is ignored by thread pool */
}


int Twapi_ResolveHostnameAsync(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    int   id;
    char *name;
    int   len;
    TwapiHostnameEvent *theP;
    DWORD winerr;

    ERROR_IF_UNTHREADED(ticP->interp);

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETINT(id), ARGSKIP, ARGEND) != TCL_OK)
        return TCL_ERROR;

    name = Tcl_GetStringFromObj(objv[1], &len);

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
    theP->addrs = NULL;
    theP->naddrs = 0;
    MoveMemory(theP->name, name, len+1);

    if (QueueUserWorkItem(TwapiHostnameHandler, theP, WT_EXECUTEDEFAULT))
        return TCL_OK;

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
        Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(theP->id));
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
    struct sockaddr_in saddr;
    char hostname[NI_MAXHOST];
    char portname[NI_MAXSERV];

    saddr.sin_family = AF_INET;
    saddr.sin_addr.s_addr = 0;
    saddr.sin_port = 0;
    saddr.sin_addr.s_addr = inet_addr(theP->name);

    theP->tcl_ev.proc = TwapiAddressEventProc;
    if (saddr.sin_addr.s_addr == INADDR_NONE && lstrcmpA(theP->name, "255.255.255.255")) {
        // Fail, invalid address string
        theP->status = 10022;         /* WSAINVAL error code */
    } else {
        theP->status = getnameinfo((struct sockaddr *)&saddr, sizeof(saddr),
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
    int   id;
    char *addrstr;
    int   len;
    TwapiHostnameEvent *theP;
    DWORD winerr;

    ERROR_IF_UNTHREADED(ticP->interp);

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETINT(id), ARGSKIP, ARGEND) != TCL_OK)
        return TCL_ERROR;

    addrstr = Tcl_GetStringFromObj(objv[1], &len);


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
    theP->naddrs = 0;

    /* We do not syntactically validate address string here. All failures
       are delivered asynchronously */
    MoveMemory(theP->name, addrstr, len+1);

    if (QueueUserWorkItem(TwapiAddressHandler, theP, WT_EXECUTEDEFAULT))
        return TCL_OK;

    winerr = GetLastError();    /* Remember the error */

    TwapiInterpContextUnref(ticP, 1); /* Undo above ref */
    Tcl_Free((char*) theP);
    return Twapi_AppendSystemError(ticP->interp, winerr);
}
