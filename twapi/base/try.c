/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>
#if (TCL_MAJOR_VERSION == 8) && (TCL_MINOR_VERSION == 4)
#include "tclInt.h" /* Needed for Twapi_SaveResultErrorInfo in 8.4 */
#endif

static int GlobalImport (Tcl_Interp *interp);

int Twapi_TryObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Obj **objv_copy;
    int       i, result, final;
    Tcl_Obj *errorCodeVar;
    Tcl_Obj **errorCodeObjv;
    int      errorCodeObjc;
    char    *facilityP;
    char    *codeP;

    if (objc < 2)
        goto badargs;

    /*
     * To quote from implementation of Tcl_CatchObjCmd -
     * "Save a pointer to the variable name object, if any, in case the
     * Tcl_EvalObj reallocates the bytecode interpreter's evaluation
     * stack rendering objv invalid."
     * - we do the same. I dunno if this is really necessary - TBD
     */
    if (objc > 128)
        goto badargs;           /* Just to prevent malicious stack overflow */
    objv_copy = _alloca(objc*sizeof(Tcl_Obj *));
    for (i = 0; i < objc; ++i)
        objv_copy[i] = objv[i];
    objv = objv_copy;

    /* First parse syntax without evaluating */
    final = 0;
    i = 2;
    while (i < objc) {
        char *s = Tcl_GetString(objv[i]);
        if (STREQ("onerror", s)) {
            /* There should be at least two more args */
            if ((i+2) >= objc)
                goto badargs;
            i += 3;             /* Skip over this onerror element */
        }
        else if (STREQ("finally", s)) {
            /* There should be at least one more args */
            if ((i+1) >= objc)
                goto badargs;
            /* Verify that we do not already have a finally element */
            if (final)
                goto badsyntax;
            final = i; /* Remember finalization script */

            i += 2;             /* Skip over this finally element */
        }
        else {
            goto badsyntax;
        }
    }

    /* Ok, now eval the first script */
    result = Tcl_EvalObjEx(interp, objv[1], 0);
    /*
     * If result is anything other than error, simply evaluate the
     * finalization script and return
     */
    if (result != TCL_ERROR)
        goto finalize;

    /* Get the errorCode variable */
    errorCodeVar = Tcl_GetVar2Ex(interp, "errorCode", NULL, TCL_GLOBAL_ONLY);
    if (errorCodeVar == NULL)
        goto finalize;          /* No errorCode set so cannot match */

    /* Get the error code facility and code */
    if (Tcl_ListObjGetElements(interp, errorCodeVar, &errorCodeObjc,
                               &errorCodeObjv) != TCL_OK) {
        /* Not in list format. Will not match any error patterns below */
        goto finalize;
    }

    facilityP = "";
    codeP     = "";
    if (errorCodeObjc > 0) {
        facilityP = Tcl_GetString(errorCodeObjv[0]);
        if (errorCodeObjc > 1)
            codeP = Tcl_GetString(errorCodeObjv[1]);
    }

    /* Check if any error conditions match */
    for (i = 2; i < (objc-2); i+= 3) {
        Tcl_Obj **codeObjv;
        int       codeObjc;
        int       match;

        if (i == final) {
            /* Skip the finally clause */
            --i;                /* finally has only 2 args, not 3 so fix up
                                   for next loop increment */
            continue;
        }

        /* Get the error code patterns and see if they matche */
        match = 0;
        if (Tcl_ListObjGetElements(interp, objv[i+1], &codeObjc, &codeObjv) != TCL_OK) {
            return TCL_ERROR;
        }

        if (codeObjc == 0)
            match = 1;        /* Note empty patterns matches any thing */
        else if (! lstrcmpA(facilityP, Tcl_GetString(codeObjv[0]))) {
            /* Facility matches. */
            if (codeObjc == 1)
                match = 1;  /* No code specified so facility need match */
            else if (! lstrcmpA(codeP, Tcl_GetString(codeObjv[1]))) {
                /* Code specified and matches */
                match = 1;
            } else {
                /* Code does not match as string. Try for integer match */
                int errorCodeInt, patInt;
                if (Tcl_GetIntFromObj(interp, codeObjv[1], &patInt) == TCL_OK &&
                    codeP[0] &&
                    Tcl_GetIntFromObj(interp, errorCodeObjv[1], &errorCodeInt) == TCL_OK &&
                    errorCodeInt == patInt) {
                    match = 1; /* Matched as integers */
                }
            }
        }


        if (match) {
            Tcl_Obj *errorResultObjP;
            /*
             * This onerror clause matches. Execute the code after importing
             * the errorCode, errorResult and errorInfo into local scope
             * We are basically cloning TclX's try_eval command code here
             */
            errorResultObjP = Tcl_DuplicateObj (Tcl_GetObjResult (interp));
            Tcl_IncrRefCount (errorResultObjP);
            Tcl_ResetResult (interp);

            /* Import errorResult, errorInfo, errorCode */
            result = GlobalImport (interp);
            if (result != TCL_ERROR) {
                /* Set errorResult variable so script can access it */
                if (Tcl_SetVar2Ex(interp, "errorResult", NULL,
                                  errorResultObjP, TCL_LEAVE_ERR_MSG) == NULL) {
                    result = TCL_ERROR;
                }
            }
            /* If no errors, eval the error handling script */
            if (result != TCL_ERROR)
                result = Tcl_EvalObjEx(interp, objv[i+2], 0);

            Tcl_DecrRefCount (errorResultObjP);

            break;              /* Error handling all done */
        }
    }


finalize:
    /*
     * If a finalize script is specified, execute it. We must preserve
     * both the result and interpreter value whether there was an error
     * or not
     */
    if (final) {
        void *savedResultsP;
        /* Again, basically cloned from TclX's try_eval */
        savedResultsP = Twapi_SaveResultErrorInfo (interp, result);
        Tcl_ResetResult (interp);

        result = Tcl_EvalObjEx (interp, objv[final+1], 0);
        if (result == TCL_ERROR)
            Twapi_DiscardResultErrorInfo(interp, savedResultsP);
        else
            result = Twapi_RestoreResultErrorInfo (interp, savedResultsP);
    }

    return result;


badsyntax:
    Tcl_ResetResult(interp);
    Tcl_AppendResult(interp, "Invalid syntax: should be ",
                     Tcl_GetString(objv[0]),
                     " SCRIPT ?onerror ERROR ERRORSCRIPT? ...?finally FINALSCRIPT?",
        NULL);
    return TCL_ERROR;

    badargs:
    Tcl_ResetResult(interp);
    Tcl_WrongNumArgs(interp, 1, objv, "script ?onerror ERROR errorscript? ...?finally FINALSCRIPT?");
    return TCL_ERROR;
}

/*-----------------------------------------------------------------------------
 * Copied from the TclX code
 * GlobalImport --
 *   Import the errorResult, errorInfo, and errorCode global variable into the
 * current environment by calling the global command directly.
 *
 * Parameters:
 *   o interp (I) - Current interpreter,  Result is preserved.
 * Returns:
 *   TCL_OK or TCL_ERROR.
 *-----------------------------------------------------------------------------
 */
static int
GlobalImport (interp)
    Tcl_Interp *interp;
{
    static char global [] = "global";
    Tcl_Obj *savedResult;
    Tcl_CmdInfo cmdInfo;
#define globalObjc (4)
    Tcl_Obj *globalObjv [globalObjc];
    int idx, code = TCL_OK;

    savedResult = Tcl_DuplicateObj (Tcl_GetObjResult (interp));

    if (!Tcl_GetCommandInfo (interp, global, &cmdInfo)) {
        Tcl_AppendResult (interp, "can't find \"global\" command",
                              (char *) NULL);
        goto errorExit;
    }

    globalObjv [0] = Tcl_NewStringObj (global, -1);
    globalObjv [1] = Tcl_NewStringObj ("errorResult", -1);
    globalObjv [2] = Tcl_NewStringObj ("errorInfo", -1);
    globalObjv [3] = Tcl_NewStringObj ("errorCode", -1);

    for (idx = 0; idx < globalObjc; idx++) {
        Tcl_IncrRefCount (globalObjv [idx]);
    }

    code = (*cmdInfo.objProc) (cmdInfo.objClientData,
                               interp,
                               globalObjc,
                               globalObjv);
    for (idx = 0; idx < globalObjc; idx++) {
        Tcl_DecrRefCount (globalObjv [idx]);
    }

    if (code == TCL_ERROR)
        goto errorExit;

    Tcl_SetObjResult (interp, savedResult);
    return TCL_OK;

  errorExit:
    Tcl_DecrRefCount (savedResult);
    return TCL_ERROR;
}

/* Copied from TclX
 *-----------------------------------------------------------------------------
 * TclX_SaveResultErrorInfo --
 *
 *   Saves the Tcl interp result plus errorInfo and errorCode in a structure.
 *
 * Parameters:
 *   o interp - Interpreter to save state for.
 * Returns:
 *   A list object containing the state.
 *-----------------------------------------------------------------------------
 */
void *
Twapi_SaveResultErrorInfo (Tcl_Interp *interp, int status)
{
#if (TCL_MAJOR_VERSION == 8) && (TCL_MINOR_VERSION == 4)

    /* Even when built against 8.4, if we are running against 8.5, we
       will use the new 8.5 routines through the stubs table. Note
       we don't run against older than 8.4 anyways */

    if (gTclVersion.major == 8 && gTclVersion.minor == 4) {
        Tcl_Obj *saveObjv [5];
        Tcl_Obj *listObj;

        long flags = ((Interp *)interp)->flags &
            (ERR_ALREADY_LOGGED | ERR_IN_PROGRESS | ERROR_CODE_SET);

        saveObjv [0] = Tcl_DuplicateObj (Tcl_GetObjResult (interp));

        saveObjv [1] = Tcl_GetVar2Ex(interp, "errorInfo", NULL, TCL_GLOBAL_ONLY);
        if (saveObjv [1] == NULL) {
            saveObjv [1] = Tcl_NewObj ();
        }

        saveObjv [2] = Tcl_GetVar2Ex(interp, "errorCode", NULL, TCL_GLOBAL_ONLY);
        if (saveObjv [2] == NULL) {
            saveObjv [2] = Tcl_NewObj ();
        }

        saveObjv [3] = Tcl_NewLongObj(flags);

        saveObjv [4] = Tcl_NewIntObj(status);

        Tcl_IncrRefCount(listObj = Tcl_NewListObj (5, saveObjv));

        return listObj;
    }
    else {
        /* Use 8.5 run time call */
        return TWAPI_TCL85_STUB(tcl_SaveInterpState) (interp, status);
    }

#else    /* Building against 8.5 or later */

    return (void *) Tcl_SaveInterpState(interp, status);

#endif

}


/* Copied from TclX
 *-----------------------------------------------------------------------------
 * TclX_RestoreResultErrorInfo --
 *
 *   Restores the Tcl interp state from TclX_SaveResultErrorInfo.
 *
 * Parameters:
 *   o interp - Interpreter to save state for.
 *   o saveObjPtr - Object returned from TclX_SaveResultErrorInfo.  Ref count
 *     will be decremented.
 *-----------------------------------------------------------------------------
 */
int Twapi_RestoreResultErrorInfo (Tcl_Interp *interp, void *savePtr)
{
#if (TCL_MAJOR_VERSION == 8) && (TCL_MINOR_VERSION == 4)
    /* Building against 8.4 only */

    /* Even when built against 8.4, if we are running against 8.5, we
       will use the new 8.5 routines through the stubs table. Note
       we don't run against older than 8.4 anyways */

    if (gTclVersion.major == 8 && gTclVersion.minor == 4) {
        Tcl_Obj *saveObjPtr = (Tcl_Obj *)savePtr;
        Tcl_Obj **saveObjv;
        int saveObjc;
        long flags;
        int status;

        if ((Tcl_ListObjGetElements (NULL, saveObjPtr, &saveObjc,
                                     &saveObjv) != TCL_OK) ||
            (saveObjc != 5) ||
            (Tcl_GetLongFromObj (NULL, saveObjv[3], &flags) != TCL_OK) ||
            (Tcl_GetIntFromObj (NULL, saveObjv[4], &status) != TCL_OK)
            ) {
            /*
             * This should never happen
             */
            panic ("invalid TclX result save object");
        }

        Tcl_SetVar2Ex(interp, "errorCode", NULL, saveObjv[2], TCL_GLOBAL_ONLY);
        Tcl_SetVar2Ex(interp, "errorInfo", NULL, saveObjv[1], TCL_GLOBAL_ONLY);

        Tcl_SetObjResult (interp, saveObjv[0]);

        ((Interp *)interp)->flags |= flags;

        Tcl_DecrRefCount (saveObjPtr);

        return status;
    }
    else {
        /* Use 8.5 run time call */
        return TWAPI_TCL85_STUB(tcl_RestoreInterpState) (interp, savePtr);
    }

#else

    return Tcl_RestoreInterpState(interp, (Tcl_InterpState) savePtr);

#endif
}

void Twapi_DiscardResultErrorInfo (Tcl_Interp *interp, void *savePtr)
{
#if (TCL_MAJOR_VERSION == 8) && (TCL_MINOR_VERSION == 4)
    /* Even when built against 8.4, if we are running against 8.5, we
       will use the new 8.5 routines through the stubs table. Note
       we don't run against older than 8.4 anyways */

    if (gTclVersion.major == 8 && gTclVersion.minor == 4) {
        Tcl_DecrRefCount ((Tcl_Obj *)savePtr);
    }
    else {
        /* Use 8.5 run time call */
        TWAPI_TCL85_STUB(tcl_DiscardInterpState) (savePtr);
    }
#else
    Tcl_DiscardInterpState((Tcl_InterpState)savePtr);
#endif
}
