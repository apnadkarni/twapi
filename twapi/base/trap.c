/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>

static int GlobalImport (Tcl_Interp *interp);

TCL_RESULT Twapi_TrapObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    int       i, final;
    Tcl_Obj *errorCodeVar;
    Tcl_Obj **errorCodeObjv;
    int      errorCodeObjc;
    char    *facilityP;
    char    *codeP;
    TCL_RESULT result = TCL_ERROR;


    if (objc < 2)
        goto badargs;


    /* First parse syntax without evaluating */
    final = 0;
    i = 2;
    while (i < objc) {
        char *s = ObjToString(objv[i]);
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
    if (ObjGetElements(interp, errorCodeVar, &errorCodeObjc,
                               &errorCodeObjv) != TCL_OK) {
        /* Not in list format. Will not match any error patterns below */
        goto finalize;
    }

    facilityP = "";
    codeP     = "";
    if (errorCodeObjc > 0) {
        facilityP = ObjToString(errorCodeObjv[0]);
        if (errorCodeObjc > 1)
            codeP = ObjToString(errorCodeObjv[1]);
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
        if (ObjGetElements(interp, objv[i+1], &codeObjc, &codeObjv) != TCL_OK) {
            goto pop_and_return;
        }

        if (codeObjc == 0)
            match = 1;        /* Note empty patterns matches any thing */
        else if (! lstrcmpA(facilityP, ObjToString(codeObjv[0]))) {
            /* Facility matches. */
            if (codeObjc == 1)
                match = 1;  /* No code specified so facility need match */
            else if (! lstrcmpA(codeP, ObjToString(codeObjv[1]))) {
                /* Code specified and matches */
                match = 1;
            } else {
                /* Code does not match as string. Try for integer match */
                int errorCodeInt, patInt;
                if (ObjToInt(interp, codeObjv[1], &patInt) == TCL_OK &&
                    codeP[0] &&
                    ObjToInt(interp, errorCodeObjv[1], &errorCodeInt) == TCL_OK &&
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
        Tcl_InterpState savedState;
        savedState = Tcl_SaveInterpState(interp, result);
        Tcl_ResetResult (interp);

        result = Tcl_EvalObjEx (interp, objv[final+1], 0);
        if (result == TCL_ERROR)
            Tcl_DiscardInterpState(savedState);
        else
            result = Tcl_RestoreInterpState(interp, savedState);
    }

    goto pop_and_return;

badsyntax:
    Tcl_ResetResult(interp);
    Tcl_AppendResult(interp, "Invalid syntax: should be ",
                     ObjToString(objv[0]),
                     " SCRIPT ?onerror ERROR ERRORSCRIPT? ...?finally FINALSCRIPT?",
        NULL);
    goto pop_and_return;

    badargs:
    Tcl_ResetResult(interp);
    Tcl_WrongNumArgs(interp, 1, objv, "script ?onerror ERROR errorscript? ...?finally FINALSCRIPT?");
    /* Fall thru */

pop_and_return:
    return result;
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

    globalObjv [0] = ObjFromString (global);
    globalObjv [1] = ObjFromString ("errorResult");
    globalObjv [2] = ObjFromString ("errorInfo");
    globalObjv [3] = ObjFromString ("errorCode");

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

    return TwapiSetObjResult (interp, savedResult);

  errorExit:
    Tcl_DecrRefCount (savedResult);
    return TCL_ERROR;
}
