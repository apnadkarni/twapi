/*
 * Copyright (c) 2010-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

struct OptionDescriptor {
    Tcl_Obj    *name; // TBD - should this store name without the .type suffix?
    Tcl_Obj    *def_value;
    Tcl_Obj    *valid_values;
    unsigned short name_len;    /* Length of option name excluding
                                   any .type suffix */
    char        type;
#define OPT_ANY    0
#define OPT_BOOL   1
#define OPT_INT    2
#define OPT_SWITCH 3
#define OPT_SYM    4
    char        first;          /* First char of name[] */
};


static int SetParseargsOptFromAny(Tcl_Interp *interp, Tcl_Obj *objP);
static void DupParseargsOpt(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void FreeParseargsOpt(Tcl_Obj *objP);
static void UpdateStringParseargsOpt(Tcl_Obj *objP);

static struct Tcl_ObjType gParseargsOptionType = {
    "TwapiParseargsOpt",
    FreeParseargsOpt,
    DupParseargsOpt,
    UpdateStringParseargsOpt,
    NULL, // SetParseargsOptFromAny    /* jenglish says keep this NULL */
};

static void CleanupOptionDescriptor(struct OptionDescriptor *optP)
{
    if (optP->name) {
        ObjDecrRefs(optP->name);
        optP->name = NULL;
    }
    if (optP->def_value) {
        ObjDecrRefs(optP->def_value);
        optP->def_value = NULL;
    }
    if (optP->valid_values) {
        ObjDecrRefs(optP->valid_values);
        optP->valid_values = NULL;
    }
}

static void UpdateStringParseargsOpt(Tcl_Obj *objP)
{
    /* Not the most efficient but not likely to be called often */
    unsigned long i;
    Tcl_Obj *listObj = ObjEmptyList();
    struct OptionDescriptor *optP;

    TWAPI_ASSERT(objP->bytes == NULL);
    TWAPI_ASSERT(objP->typePtr == &gParseargsOptionType);

    for (i = 0, optP = (struct OptionDescriptor *) objP->internalRep.ptrAndLongRep.ptr;
         i < objP->internalRep.ptrAndLongRep.value;
         ++i, ++optP) {
        Tcl_Obj *elems[3];
        int nelems;
        elems[0] = optP->name;
        nelems = 1;
        if (optP->def_value) {
            ++nelems;
            elems[1] = optP->def_value;
            if (optP->valid_values) {
                ++nelems;
                elems[2] = optP->valid_values;
            }
        }
        ObjAppendElement(NULL, listObj, ObjNewList(nelems, elems));
    }

    ObjToString(listObj);     /* Ensure string rep */

    /* We could just shift the bytes field from listObj to objP resetting
       the former to NULL. But I'm nervous about doing that behind Tcl's back */
    objP->length = listObj->length; /* Note does not include terminating \0 */
    objP->bytes = ckalloc(listObj->length + 1);
    CopyMemory(objP->bytes, listObj->bytes, listObj->length+1);
    ObjDecrRefs(listObj);
}

static void FreeParseargsOpt(Tcl_Obj *objP)
{
    unsigned long i;
    struct OptionDescriptor *optsP;

    if ((optsP = (struct OptionDescriptor *)objP->internalRep.ptrAndLongRep.ptr) != NULL) {
        for (i = 0; i < objP->internalRep.ptrAndLongRep.value; ++i) {
            CleanupOptionDescriptor(&optsP[i]);
        }
        ckfree((char *) optsP);
    }

    objP->internalRep.ptrAndLongRep.ptr = NULL;
    objP->internalRep.ptrAndLongRep.value = 0;
    objP->typePtr = NULL;
}

static void DupParseargsOpt(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    int i;
    struct OptionDescriptor *doptsP;
    struct OptionDescriptor *soptsP;
    
    dstP->typePtr = &gParseargsOptionType;
    soptsP = srcP->internalRep.ptrAndLongRep.ptr;
    if (soptsP == NULL) {
        dstP->internalRep.ptrAndLongRep.ptr = NULL;
        dstP->internalRep.ptrAndLongRep.value = 0;
        return;
    }

    doptsP = (struct OptionDescriptor *) ckalloc(srcP->internalRep.ptrAndLongRep.value * sizeof(*doptsP));
    dstP->internalRep.ptrAndLongRep.ptr = doptsP;
    i = srcP->internalRep.ptrAndLongRep.value;
    dstP->internalRep.ptrAndLongRep.value = i;
    while (i--) {
        if ((doptsP->name = soptsP->name) != NULL)
            ObjIncrRefs(doptsP->name);
        if ((doptsP->def_value = soptsP->def_value) != NULL)
            ObjIncrRefs(doptsP->def_value);
        if ((doptsP->valid_values = soptsP->valid_values) != NULL)
            ObjIncrRefs(doptsP->valid_values);
        doptsP->name_len = soptsP->name_len;
        doptsP->type = soptsP->type;
        doptsP->first = soptsP->first;
        ++doptsP;
        ++soptsP;
    }

}


static int SetParseargsOptFromAny(Tcl_Interp *interp, Tcl_Obj *objP)
{
    int k, nopts;
    Tcl_Obj **optObjs;
    struct OptionDescriptor *optsP;
    struct OptionDescriptor *curP;
    int len;

    if (objP->typePtr == &gParseargsOptionType)
        return TCL_OK;          /* Already in correct format */

    if (ObjGetElements(interp, objP, &nopts, &optObjs) != TCL_OK)
        return TCL_ERROR;
    
    
    optsP = nopts ? (struct OptionDescriptor *) ckalloc(nopts * sizeof(*optsP)) : NULL;

    for (k = 0; k < nopts ; ++k) {
        Tcl_Obj **elems;
        int       nelems;
        const char     *type;
        const char *p;

        curP = &optsP[k];

        /* Init to NULL first so error handling frees correctly */
        curP->name = NULL;
        curP->def_value = NULL;
        curP->valid_values = NULL;

        if (ObjGetElements(interp, optObjs[k], &nelems, &elems) != TCL_OK ||
            nelems == 0) {
            goto error_handler;
        }

        curP->type = OPT_SWITCH; /* Assumed option type */
        curP->name = elems[0];
        ObjIncrRefs(elems[0]);
        p = ObjToStringN(elems[0], &len);
        curP->first = *p;
        type = Tcl_UtfFindFirst(p, '.');
        if (type == NULL)
            curP->name_len = (unsigned short) len;
        else {
            if (type == p)
                goto error_handler;

            curP->name_len = (unsigned short) (type - p);
            ++type;          /* Point to type descriptor */
            if (STREQ(type, "int"))
                curP->type = OPT_INT;
            else if (STREQ(type, "arg"))
                curP->type = OPT_ANY;
            else if (STREQ(type, "bool"))
                curP->type = OPT_BOOL;
            else if (STREQ(type, "switch"))
                curP->type = OPT_SWITCH;
            else if (STREQ(type, "sym"))
                curP->type = OPT_SYM;
            else
                goto error_handler;
        }
        if (nelems > 1) {
            /* Squirrel away specified default */
            curP->def_value = elems[1];
            ObjIncrRefs(elems[1]);

            if (nelems > 2) {
                Tcl_Obj **validObjs;
                int nvalid;
                if (ObjGetElements(interp, elems[2], &nvalid, &validObjs) != TCL_OK)
                    goto error_handler;
                if (nvalid == 0) {
                    ObjSetStaticResult(interp, "Empty validity list.");
                    goto error_handler;
                }
                if (curP->type == OPT_SYM && (nvalid & 1)) {
                    ObjSetStaticResult(interp, "Dictionary must have even number of elements");
                    goto error_handler;
                }
                /* Value must be in the specified list or for BOOL and SWITCH
                   value to use for 'true' */
                curP->valid_values = elems[2];
                ObjIncrRefs(elems[2]);
            }
        }
    }
    
    /* OK, options are in order. Convert the passed object's internal rep */
    if (objP->typePtr && objP->typePtr->freeIntRepProc) {
        objP->typePtr->freeIntRepProc(objP);
        objP->typePtr = NULL;
    }

#if 0
    /* 
     * Commented out - as per msofer:
     * First reaction, from memory, is:
     * a literal should ALWAYS have a string rep - the exact string rep it
     * had when stored as a literal. As that string rep is not guaranteed to
     * be regenerated exactly, it should never be cleared.
     * 
     * Second reaction: Tcl_InvalidateStringRep must not be called on
     * a shared object, ever. This is because ... because ... why was
     * it?
     * 
     * Note that you can shimmer a shared Tcl_Obj, ie, change its
     * internal rep. You can also generate a string rep if there was
     * none to begin with. But a Tcl_Obj that has a string rep must
     * retain it forever, until its last ref is gone and the obj is
     * returned to free mem.
     * END QUOTE
     *
     * Since original intent was to just save memory, do not need this.
     */
    
    Tcl_InvalidateStringRep(objP);
#endif

    objP->internalRep.ptrAndLongRep.ptr = optsP;
    objP->internalRep.ptrAndLongRep.value = nopts;
    objP->typePtr = &gParseargsOptionType;

    return TCL_OK;

error_handler: /* Tcl error result must have been set */
    /* k holds highest index that has been processed and is the error */
    if (interp) {
        Tcl_AppendResult(interp, "Badly formed option descriptor: '",
                         ObjToString(optObjs[k]), "'", NULL);
        Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(TWAPI_INVALID_ARGS));
    }
    if (optsP) {
        while (k >= 0) {
            CleanupOptionDescriptor(&optsP[k]);
            --k;
        }
        ckfree((char *)optsP);
    }
    return TCL_ERROR;
}


static void TwapiParseargsSetResultBadValue (Tcl_Interp *interp, const char *error_type,
                                    Tcl_Obj *value,
                                    Tcl_Obj *opt_name, int opt_name_len)
{
    /* TBD - limit length of value */
    ObjSetResult(interp,
                 Tcl_ObjPrintf("%s value '%s' specified for option '-%.*s'.",
                               error_type ? error_type : "Invalid",
                               ObjToString(value),
                               opt_name_len, ObjToString(opt_name)));
}

static void TwapiParseargsUnknownOption(Tcl_Interp *interp, char *badopt, struct OptionDescriptor *opts, int nopts)
{
    Tcl_Obj *objP;
    int j;
    char *sep = "-";

    objP = Tcl_ObjPrintf("Invalid option '%s'. Must be one of ", badopt);

    for (j = 0; j < nopts; ++j) {
        Tcl_AppendPrintfToObj(objP, "%s%.*s", sep, opts[j].name_len, ObjToString(opts[j].name));
        sep = ", -";
    }

    ObjSetResult(interp, objP);
    return;
}


/*
 * Argument parsing command
 */
int Twapi_ParseargsObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Obj    *argvObj;
    int         argc, iarg;
    Tcl_Obj   **argv;
    int         nopts;
    int         j, k;
    Tcl_WideInt wide;
    struct OptionDescriptor *opts;
    int         ignoreunknown = 0;
    int         nulldefault = 0;
    int         hyphenated = 0;
    int         setvars = 0;
    Tcl_Obj    *newargvObj = NULL;
    int         maxleftover = INT_MAX;
#define TWAPI_PARSEARGS_STATIC 20
    Tcl_Obj    *values[TWAPI_PARSEARGS_STATIC];
    Tcl_Obj    **valuesP = NULL;
    Tcl_Obj    *retObjs[2*TWAPI_PARSEARGS_STATIC];
    Tcl_Obj    **retP = NULL;
    int         nret = 0;

    if (objc < 3) {
        Tcl_WrongNumArgs(interp, 1, objv, "argvVar optlist ?-ignoreunknown? ?-nulldefault? ?-hyphenated? ?-maxleftover COUNT? ?--?");
        Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(TWAPI_BAD_ARG_COUNT));
        return TCL_ERROR;
    }

    /* Now construct the option descriptors */
    if (objv[2]->typePtr != &gParseargsOptionType) {
        if (SetParseargsOptFromAny(interp, objv[2]) != TCL_OK)
            return TCL_ERROR;
    }

    opts =  objv[2]->internalRep.ptrAndLongRep.ptr;
    nopts = objv[2]->internalRep.ptrAndLongRep.value;

    if (nopts > TWAPI_PARSEARGS_STATIC) {
        valuesP = MemLifoPushFrame(ticP->memlifoP, nopts * sizeof(*valuesP), NULL);
        retP = MemLifoAlloc(ticP->memlifoP, 2*nopts*sizeof(Tcl_Obj*), NULL);
    } else {
        valuesP = values;
        retP = retObjs;
    }

    for (k = 0; k < nopts; ++k)
        valuesP[k] = NULL;      /* Values corresponding to each option */

    /* TBD - use Tcl_GetIndexFromObj to parse parseargs options. Faster */
    for (j = 3 ; j < objc ; ++j) {
        char *s = ObjToString(objv[j]);
        if (STREQ("-nulldefault", s)) {
            nulldefault = 1;
        }
        else if (STREQ("-maxleftover", s)) {
            ++j;
            if (j == objc) {
                ObjSetStaticResult(interp, "Missing value for -maxleftover");
                goto invalid_args_error;
            }

            if (ObjToInt(interp, objv[j], &maxleftover) != TCL_OK) {
                goto invalid_args_error;
            }
        }
        else if (STREQ("-setvars", s)) {
            setvars = 1;
        }
        else if (STREQ("-ignoreunknown", s)) {
            ignoreunknown = 1;
        }
        else if (STREQ("-hyphenated", s)) {
            hyphenated = 1;
        }
        else {
            Tcl_AppendResult(interp, "Extra argument or unknown option '",
                             ObjToString(objv[j]), "'", NULL);
            goto invalid_args_error;
        }
    }

    /* Collect the arguments into an array */
    argvObj = Tcl_ObjGetVar2(interp, objv[1], NULL, TCL_LEAVE_ERR_MSG);
    if (argvObj == NULL)
        goto error_return;

    if (ObjGetElements(interp, argvObj, &argc, &argv) != TCL_OK)
        goto error_return;

    newargvObj = ObjNewList(0, NULL);

    /* OK, now go through the passed arguments */
    for (iarg = 0; iarg < argc; ++iarg) {
        int   argp_len;
        char *argp = ObjToStringN(argv[iarg], &argp_len);

        /* Non-option arg or a '-' or a "--" signals end of arguments */
        if (*argp != '-')
            break;
        if ((argp[1] == 0) ||
            (argp[1] == '-' && argp[2] == 0)) {
            ++iarg;             /* Skip the - or -- */
            break;
        }

        /* Check against each option in turn */
        for (j = 0; j < nopts; ++j) {
            if (opts[j].name_len == (argp_len-1) &&
                opts[j].first == argp[1] &&
                ! Tcl_UtfNcmp(ObjToString(opts[j].name), argp+1, (argp_len-1))) {
                break;          /* Match ! */
            }
        }

        if (j < nopts) {
            /*
             *  Matches option j. Remember the option value.
             */
            if (opts[j].type == OPT_SWITCH) {
                valuesP[j] = ObjFromBoolean(1);
            }
            else {
                if (iarg >= (argc-1)) {
                    /* No more args! */
                    Tcl_AppendResult(interp, "No value supplied for option '", argp, "'", NULL);
                    Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(TWAPI_INVALID_ARGS));
                    goto error_return;
                }
                valuesP[j] = argv[iarg+1];
                ++iarg;            /* Move on to next arg in array */
            }
        }
        else {
            /* Does not match any option. */
            if (! ignoreunknown) {
                TwapiParseargsUnknownOption(interp, argp, opts, nopts);
                Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(TWAPI_INVALID_ARGS));
                goto error_return;
            }
            /*
             * We have to ignore this option. But we do not know if
             * it takes an argument value or not. So we check the next
             * argument. If it does not begin with a "-" assume it
             * is the value of the unknown option. Else it is the next option
             */
            ObjAppendElement(interp, newargvObj, argv[iarg]);
            if (iarg < (argc-1)) {
                argp = ObjToString(argv[iarg+1]);
                if (*argp != '-') {
                    /* Assume this is the value for the option */
                    ++iarg;
                    ObjAppendElement(interp, newargvObj, argv[iarg]);
                }
            }
        }
    }

    /*
     * Now loop through the option definitions, collecting the option
     * values, using defaults for unspecified options
     */
    for (k = 0; k < nopts ; ++k) {
        Tcl_Obj  *objP;

        /* OPT_SWITCH ignores defaults */
        if (opts[k].type != OPT_SWITCH) {
            if (valuesP[k] == NULL) {
                /* Option not specified. Check for a default and use it */
                if (opts[k].def_value == NULL && !nulldefault)
                    continue;       /* No default, so skip */
                valuesP[k] = opts[k].def_value;
            }
        }

        /* Do value type checking */
        switch (opts[k].type) {
        case OPT_INT:
            /* If no explicit default, but have a -nulldefault switch,
             * (else we would have continued above), return 0
             */
            if (valuesP[k] == NULL)
                wide = 0;
            else if (ObjToWideInt(interp, valuesP[k], &wide) == TCL_ERROR) {
                TwapiParseargsSetResultBadValue(interp, "Non-integer",
                                       valuesP[k],
                                       opts[k].name,
                                       opts[k].name_len);
                goto error_return;
            }

            /* Check list of allowed values if specified */
            if (opts[k].valid_values) {
                Tcl_Obj **validObjs;
                int nvalid, ivalid;
                if (ObjGetElements(interp, opts[k].valid_values, &nvalid, &validObjs) != TCL_OK)
                    goto error_return;
                for (ivalid = 0; ivalid < nvalid; ++ivalid) {
                    Tcl_WideInt valid_wide;
                    if (ObjToWideInt(interp, validObjs[ivalid], &valid_wide) == TCL_ERROR) {
                        TwapiParseargsSetResultBadValue(
                            interp,
                            "Non-integer enumeration",
                            validObjs[ivalid],
                            opts[k].name, opts[k].name_len);
                        goto error_return;
                    }
                    if (valid_wide == wide)
                        break;
                }
                if (ivalid == nvalid) {
                    TwapiParseargsSetResultBadValue(interp, "Invalid",
                                                    valuesP[k],
                                                    opts[k].name,
                                                    opts[k].name_len);
                    goto error_return;
                }
            }
            if (valuesP[k] == NULL)
                valuesP[k] = ObjFromInt(0); /* Deals with -nulldefault case. */
            break;

        case OPT_ANY:
            /* If no explicit default, but have a -nulldefault switch,
             * (else we would have continued above), return ""
             */
            /* Check list of allowed values if specified */
            if (opts[k].valid_values) {
                Tcl_Obj **validObjs;
                int nvalid, ivalid;
                char *s;
                if (ObjGetElements(interp, opts[k].valid_values, &nvalid, &validObjs) != TCL_OK)
                    goto error_return;
                s = valuesP[k] ? ObjToString(valuesP[k]) : "";
                for (ivalid = 0; ivalid < nvalid; ++ivalid) {
                    if (!lstrcmpA(ObjToString(validObjs[ivalid]), s))
                        break;
                }
                if (ivalid == nvalid) {
                    TwapiParseargsSetResultBadValue(interp, "Invalid",
                                                    valuesP[k],
                                                    opts[k].name,
                                                    opts[k].name_len);
                    goto error_return;
                }
            }
            if (valuesP[k] == NULL)
                valuesP[k] = ObjFromEmptyString(); /* Deals with -nulldefault */

            break;

        case OPT_SYM:
            /* If no explicit default, but have a -nulldefault switch,
             * (else we would have continued above), return ""
             */
            
            /* Check list of allowed values if specified */
            if (opts[k].valid_values) {
                Tcl_Obj *symvalObj;
                if (ObjDictGet(interp, opts[k].valid_values, valuesP[k], &symvalObj) != TCL_OK)
                    goto error_return; // Really should not happen
                if (symvalObj) {
                    valuesP[k] = symvalObj;
                } else {
                    TwapiParseargsSetResultBadValue(interp, NULL,
                                                    valuesP[k],
                                                    opts[k].name,
                                                    opts[k].name_len);
                    goto error_return;
                }
            }
            if (valuesP[k] == NULL)
                valuesP[k] = ObjFromEmptyString(); /* Deals with -nulldefault */
            break;

        case OPT_SWITCH:
            /* Fallthru */
        case OPT_BOOL:
            if (valuesP[k] == NULL)
                valuesP[k] = ObjFromBoolean(0);
            else {
                if (ObjToBoolean(interp, valuesP[k], &j) == TCL_ERROR) {
                    TwapiParseargsSetResultBadValue(interp, "Non-boolean",
                                                    valuesP[k],
                                                    opts[k].name,
                                                    opts[k].name_len);
                    goto error_return;
                }
                if (j && opts[k].valid_values) {
                    /* Note the AppendElement below will incr its ref count */
                    valuesP[k] = opts[k].valid_values;
                } else {
                    /*
                     * Note: Can't just do a SetBoolean as object is shared
                     * Need to allocate a new obj
                     * BAD  - Tcl_SetBooleanObj(opts[k].value, j); 
                     */
                    valuesP[k] = ObjFromBoolean(j);
                }
            }
            break;

        default:
            break;
        }

        /* Tack it on to result */
        if (hyphenated) {
            objP = STRING_LITERAL_OBJ("-");
            Tcl_AppendToObj(objP, Tcl_GetString(opts[k].name), opts[k].name_len);
        } else {
            objP = ObjFromStringN(ObjToString(opts[k].name), opts[k].name_len);
        }

        retP[nret++] = objP;
        retP[nret++] = valuesP[k];
    }


    if (maxleftover < (argc - iarg)) {
        Tcl_Obj *tooManyErrorObj;
        Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(TWAPI_EXTRA_ARGS));
        tooManyErrorObj = STRING_LITERAL_OBJ("Command has extra arguments specified: ");
        TwapiAppendObjArray(tooManyErrorObj, argc-iarg,
                            &argv[iarg], " ");
        ObjSetResult(interp, tooManyErrorObj);
        goto error_return;
    }

    /* Tack on the remaining items in the argument list to new argv */
    while (iarg < argc) {
        ObjAppendElement(interp, newargvObj, argv[iarg]);
        ++iarg;
    }

    if (setvars) {
        for (j = 0; j < nret; j += 2) {
            if (Tcl_ObjSetVar2(interp, retP[j], NULL, retP[j+1], TCL_LEAVE_ERR_MSG) == NULL)
                goto error_return;
        }
    } else
        ObjSetResult(interp, ObjNewList(nret, retP));

    /* VERY IMPORTANT: Note we do this LAST!! Because retP[] may hold
       references to some of the objects in the variable objv[1], we
       do not want those going away when the variable's value changes.
       So only update the variable after we create a list from retP above
    */
    if (Tcl_ObjSetVar2(interp, objv[1], NULL, newargvObj, TCL_LEAVE_ERR_MSG)
        == NULL) {
        goto error_return;
    }

    if (valuesP && valuesP != values)
        MemLifoPopFrame(ticP->memlifoP);

    return TCL_OK;

invalid_args_error:
    Tcl_SetObjErrorCode(interp, Twapi_MakeTwapiErrorCodeObj(TWAPI_INVALID_ARGS));

error_return:
    /* Free up allocated resources that were not used because of error */
    if (newargvObj)
        ObjDecrRefs(newargvObj);
    if (retP && nret) {
        /* Note we cannot just ObjDecrRefs retP[] objects since
           some will have ref 0 and some that came from argv[] 1 or more.
           To free, we need to incr and then decr. Otherwise we will
           land up freeing a ref of 1 belong to someone else
        */
        for (j = 0; j < nret; ++j) {
            ObjIncrRefs(retP[j]);
            ObjDecrRefs(retP[j]);
        }
    }
    if (valuesP && valuesP != values)
        MemLifoPopFrame(ticP->memlifoP);

    return TCL_ERROR;
}


