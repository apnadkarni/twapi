/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>


int Twapi_KlGetObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Obj **klObj;
    int       count;
    int       i;
    char     *key;


    if (objc < 3 || objc > 4) {
        Tcl_WrongNumArgs(interp, 1, objv, "KEYLIST KEY ?DEFAULT?");
        return TCL_ERROR;
    }

    if (Tcl_ListObjGetElements(interp, objv[1], &count, &klObj) != TCL_OK) {
        return TCL_ERROR;
    }

    if (count & 1) {
        Tcl_SetResult(interp, "Invalid keyed list format. Must have even number of elements.", TCL_STATIC);
        return TCL_ERROR;
    }

    /* Search for the key. */
    key = Tcl_GetString(objv[2]);
    for (i = 0; i < count; i += 2) {
        char *entry = Tcl_GetString(klObj[i]);
        if (STREQ(key, entry)) {
            Tcl_SetObjResult(interp, klObj[i+1]);
            return TCL_OK;
        }
    }

    /* Not found. see if a default was specified */
    if (objc == 4) {
        Tcl_SetObjResult(interp, objv[3]);
        return TCL_OK;
    }

    Tcl_AppendResult(interp, "No field ", key, " found in keyed list.", NULL);
    return TCL_ERROR;
}
