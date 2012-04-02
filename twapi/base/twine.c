/*
 * Copyright (c) 2010-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>

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
        if (Tcl_ListObjGetElements(interp, objv[1], &n, &nested_list) != TCL_OK)
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

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}

/* Twine (merge) two lists */
Tcl_Obj *TwapiTwine(Tcl_Interp *interp, Tcl_Obj *first, Tcl_Obj *second)
{
    Tcl_Obj **list1;
    Tcl_Obj **list2;
    int i, n1, n2, nmin;
    Tcl_Obj *resultObj;

    if (Tcl_ListObjGetElements(interp, first, &n1, &list1) != TCL_OK ||
        Tcl_ListObjGetElements(interp, second, &n2, &list2) != TCL_OK) {
        return NULL;
    }

    /* TBD - much faster to build lists, not dicts and let Tcl shimmer when necessary */

    resultObj = Tcl_NewDictObj();
    nmin = n1 > n2 ? n2 : n1;
    for (i = 0;  i < nmin; ++i) {
        Tcl_DictObjPut(interp, resultObj, list1[i], list2[i]);
    }

    if (i < n1) {
        /* n1 != n2 and n2 was the minimum. Use an empty object for list2 */
        Tcl_Obj *empty = Tcl_NewStringObj("", 0);
        for ( ; i < n1; ++i) {
            Tcl_DictObjPut(interp, resultObj, list1[i], empty);
        }
    } else if (i < n2) {
        /* n1 != n2 and n1 was the minimum. Use an empty object for list1 */
        Tcl_Obj *empty = Tcl_NewStringObj("", 0);
        for ( ; i < n2; ++i) {
            Tcl_DictObjPut(interp, resultObj, empty, list2[i]);
        }
    }

    return resultObj;
}

/* Twine (merge) objv arrays */
Tcl_Obj *TwapiTwineObjv(Tcl_Obj **first, Tcl_Obj **second, int n)
{
    int i;
    Tcl_Obj *resultObj;

#if 0
    Commented out - much faster to build a list and let
    Tcl automatically shimmer as needed.
    resultObj = Tcl_NewDictObj();
    for (i = 0;  i < n; ++i) {
        Tcl_DictObjPut(NULL, resultObj, first[i], second[i]);
    }

    return resultObj;
#else
    Tcl_Obj *objv[2*100];
    Tcl_Obj **objs;

    if ((2*n) > ARRAYSIZE(objv))
        objs = TwapiAlloc(2 * n * sizeof(*objs));
    else
        objs = objv;

    for (i = 0; i < n; ++i) {
        objs[2*i] = first[i];
        objs[1 + 2*i] = second[i];
    }

    resultObj = Tcl_NewListObj(2*n, objs);

    if (objs != objv)
        TwapiFree(objs);

    return resultObj;

#endif
}
