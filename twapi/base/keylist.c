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

    if (ObjGetElements(interp, objv[1], &count, &klObj) != TCL_OK) {
        return TCL_ERROR;
    }

    if (count & 1) {
        TwapiSetStaticResult(interp, "Keyed list must have even number of elements.");
        return TCL_ERROR;
    }

    /* Search for the key. */
    key = ObjToString(objv[2]);
    for (i = 0; i < count; i += 2) {
        char *entry = ObjToString(klObj[i]);
        if (STREQ(key, entry)) {
            return TwapiSetObjResult(interp, klObj[i+1]);
        }
    }

    /* Not found. see if a default was specified */
    if (objc == 4) {
        return TwapiSetObjResult(interp, objv[3]);
    }

    Tcl_AppendResult(interp, "No field ", key, " found in keyed list.", NULL);
    return TCL_ERROR;
}
