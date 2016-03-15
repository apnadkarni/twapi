/*
 * Copyright (c) 2010-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

int Twapi_TwineObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Obj *list1;
    Tcl_Obj *list2;
    Tcl_Obj *resultObj;
    int n;

    if (objc == 2) {
        /* Single argument - must be a nested list of two lists */
        Tcl_Obj **nested_list;
        if (ObjGetElements(interp, objv[1], &n, &nested_list) != TCL_OK)
            return TCL_ERROR;
        if (n != 2)
            return TwapiReturnError(interp, TWAPI_INVALID_ARGS);
        list1 = nested_list[0];
        list2 = nested_list[1];
    } else if (objc == 3) {
        list1 = objv[1];
        list2 = objv[2];
    } else {
        Tcl_WrongNumArgs(interp, 1, objv, "LIST1 ?LIST2?");
        return TCL_ERROR;
    }
    
    resultObj = TwapiTwine(interp, list1, list2);
    if (resultObj == NULL)
        return TCL_ERROR;

    return ObjSetResult(interp, resultObj);
}

Tcl_Obj *TwapiTwine(Tcl_Interp *interp, Tcl_Obj *first, Tcl_Obj *second)
{
    Tcl_Obj **list1;
    Tcl_Obj **list2;
    int i, n1, n2, nmin;
    Tcl_Obj *resultObj;

    if (ObjGetElements(interp, first, &n1, &list1) != TCL_OK ||
        ObjGetElements(interp, second, &n2, &list2) != TCL_OK) {
        return NULL;
    }

    nmin = n1 > n2 ? n2 : n1;

    resultObj = TwapiTwineObjv(list1, list2, nmin);
    if (nmin < n1) {
        /* n1 != n2 and n2 was the minimum. Use an empty object for list2 */
        Tcl_Obj *empty = ObjFromEmptyString();
        for (i = nmin ; i < n1; ++i) {
            ObjAppendElement(interp, resultObj, list1[i]);
            ObjAppendElement(interp, resultObj, empty);
        }
    } else if (nmin < n2) {
        /* n1 != n2 and n1 was the minimum. Use an empty object for list1 */
        Tcl_Obj *empty = ObjFromEmptyString();
        for (i = nmin ; i < n2; ++i) {
            ObjAppendElement(interp, resultObj, empty);
            ObjAppendElement(interp, resultObj, list2[i]);
        }
    }

    return resultObj;
}


/* Twine (merge) objv arrays */
Tcl_Obj *TwapiTwineObjv(Tcl_Obj **first, Tcl_Obj **second, int n)
{
    int i;
    Tcl_Obj *resultObj;
    Tcl_Obj *objv[2*100];
    Tcl_Obj **objs;

    if ((2*n) > ARRAYSIZE(objv))
        objs = SWSPushFrame(2 * n * sizeof(*objs), NULL);
    else
        objs = objv;

    for (i = 0; i < n; ++i) {
        objs[2*i] = first[i];
        objs[1 + 2*i] = second[i];
    }

    resultObj = ObjNewList(2*n, objs);

    if (objs != objv)
        SWSPopFrame();

    return resultObj;
}
