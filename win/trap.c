/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

TCL_RESULT Twapi_TrapObjCmd(
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext *) clientData;
    int       i, final;
    Tcl_Obj *errorCodeVar;
    Tcl_Obj **errorCodeObjv;
    Tcl_Size errorCodeObjc;
    char    *facilityP;
    char    *codeP;
    TCL_RESULT result = TCL_ERROR;
    int      trapstackpushed = 0;

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

    ObjAppendElement(NULL, BASE_CONTEXT(ticP)->trapstack, ObjGetResult(interp));
    ObjAppendElement(NULL, BASE_CONTEXT(ticP)->trapstack, Tcl_GetReturnOptions(interp, result));
    trapstackpushed = 1;

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
        Tcl_Size  codeObjc;
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
            result = Tcl_EvalObjEx(interp, objv[i+2], 0);
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
    if (trapstackpushed) {
        Tcl_Obj *objP;
        Tcl_Size n = 0;
        objP = BASE_CONTEXT(ticP)->trapstack;
        ObjListLength(NULL, objP, &n);
        TWAPI_ASSERT(n >= 2);
        if (Tcl_IsShared(objP)) {
            objP = ObjDuplicate(objP);
            ObjIncrRefs(objP);
            ObjDecrRefs(BASE_CONTEXT(ticP)->trapstack);
            BASE_CONTEXT(ticP)->trapstack = objP;
        }
        Tcl_ListObjReplace(interp, objP, n-2, 2, 0, NULL);
    }

    return result;
}

