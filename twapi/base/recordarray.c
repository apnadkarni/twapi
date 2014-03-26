/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

int Twapi_RecordArrayHelperObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    int i, j;
    Tcl_Obj **raObj;
    static const char *opts[] = {
        "-format",              /* FORMAT */
        "-slice",               /* FIELDNAMES */
        "-filter",              /* OPER FIELDNAME VALUE */
        "-key",                 /* FIELDNAME */
        "-first",               /* no args */
        NULL
    };
    enum opts_enum {RA_FORMAT, RA_SLICE, RA_FILTER, RA_KEY, RA_FIRST};
    int opt;
    /* Format of each record */
    static const char *formats[] = {
        "recordarray", "flat", "list", "dict", NULL
    };
    enum format_enum {RA_ARRAY, RA_FLAT, RA_LIST, RA_DICT};
    int format = RA_ARRAY;
    static const char *filter_ops[] = {
        "eq", "ne", "==", "!=", "~", "!~", NULL
    };
    enum filter_ops_enum {RA_EQ, RA_NE, RA_EQ_INT, RA_NE_INT, RA_MATCH, RA_NOMATCH};

    Tcl_Obj *sliceObj = NULL,
        *filterObj = NULL,
        *keyfieldObj = NULL;
    Tcl_Obj **filterElems = NULL;
    Tcl_Obj *recsObj = NULL;     /* Dup of records passed in */
    Tcl_Obj **recs;              /* Contents of recsObj */
    int      nrecs;              /* Number records */
    Tcl_Obj *fieldsObj = NULL;   /* Dup of record definition passed in */
    Tcl_Obj **fields;            /* Contents of recsObj */
    int      nfields;            /* Number of fields in record */
    int keyfield_pos;            /* Position of the key field */
    int first = 0;               /* If true, only first match returned */
    struct {
        union {
            Tcl_WideInt wide;
            char *string;
        } operand;
        int (WINAPI *cmpfn) (const char *, const char *);
        int filter_pos;
        int filter_op;
        int nocase;
        int negate;
    } *filters;
    int nfilters;
    TCL_RESULT res;

    Tcl_Obj *new_rec[2];        /* 2 because we may need one for the key */
    Tcl_Obj **output = NULL;
    int output_count;

    Tcl_Obj **slice_fields = NULL;   /* Names of the fields to retrieve */
    int       *slice_fieldindices = NULL; /* Positions of the fields to retrieve */
    int      nslice_fields;   /* Count of above */
    Tcl_Obj  **slice_values = NULL;
    Tcl_Obj  *newfieldsObj = NULL;

    MemLifoMarkHandle mark = NULL;
    
    if (objc < 2) {
        Tcl_WrongNumArgs(interp, 1, objv, "RECORDARRAY ?OPTIONS?");
        return TCL_ERROR;
    }

    /*
     * recordarray REC
     *  Returns the values list
     *
     * recordarray options REC
     *   -slice FIELDNAMES
     *      Returns only those fields that are included in FIELDNAMES
     *      in the order specified
     *   -format [recordarray | flat | list | dict]
     *      recordarray - return value is in recordarray format (default)
     *      flat - all records are concatenated and returned as
     *             a flat list of values
     *      list - each returned record list
     *      dict - each returned record is a dict with keys being field names
     *   -filter {{FIELDNAME OPERATOR OPERAND ?-nocase?}....}
     *      Only those records whose field FIELDNAME match OPERAND using
     *      the given OPERATOR are returned
     *   -key KEYFIELD
     *      Only used if -format is specified as 'list' or 'dict'.
     *      The returned value is a dictionary with KEYFIELD as the key
     *   -first
     *      Only returns the first matching record
     */ 

    /* Figure out the command options */
    for (i = 1 ; i < objc-1; ++i) {
        if (Tcl_GetIndexFromObj(interp, objv[i], opts, "option", TCL_EXACT, &opt) != TCL_OK)
            return TCL_ERROR;
        switch (opt) {
        case RA_FORMAT:                 /* -format FORMAT */
            if (++i == (objc-1)) {
            missing_value:
                return TwapiReturnErrorEx(interp, TWAPI_INVALID_ARGS, Tcl_ObjPrintf("Missing value for option %s", ObjToString(objv[i-1])));
            }
            if (Tcl_GetIndexFromObj(interp, objv[i], formats, "format", TCL_EXACT, &format) != TCL_OK)
                return TCL_ERROR;
            break;
        case RA_SLICE:          /* -slice FIELDNAMEPAIRS */
            if (++i == (objc-1))
                goto missing_value;
            sliceObj = objv[i];
            break;
        case RA_FILTER:         /* -filter FILTER */
            if (++i == (objc-1))
                goto missing_value;
            filterObj = ObjDuplicate(objv[i]); /* Protect against shimmer */
            break;
        case RA_KEY:
            if (++i == (objc-1))
                goto missing_value;
            keyfieldObj = objv[i];
            break;
        case RA_FIRST:
            first = 1;
            break;
        }
    }

    if (ObjGetElements(interp, objv[objc-1], &i, &raObj) != TCL_OK)
        return TCL_ERROR;

    if (i == 0)
        return TCL_OK; /* TBD - check  empty result is valid for all commands */

    if (i != 2)
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "Invalid recordarray format");

    /* No commands -> return as is */
    if (objc == 2) {
        ObjSetResult(interp, objv[1]);
        return TCL_OK;
    }

    keyfield_pos = -1;
    /* Key field is ignored unless output is RA_LIST or RA_DICT */
    if (keyfieldObj && (format == RA_LIST || format == RA_DICT)) {
        if ((res=ObjToEnum(interp, raObj[0], keyfieldObj, &keyfield_pos)) != TCL_OK)
            return res;
    }

    /* Any exits hereon must MemLifoPopFrame */
    mark = MemLifoPushMark(ticP->memlifoP);

    /* If selection criteria are given, find index of field to match on */
    nfilters = 0;
    if (filterObj) {
        res = ObjGetElements(interp, filterObj, &nfilters, &filterElems);
        if (res != TCL_OK)
            goto vamoose;
        filters = MemLifoAlloc(ticP->memlifoP, nfilters * sizeof(*filters), NULL);
        for (i = 0; i < nfilters; ++i) {
            Tcl_Obj **filterElem;
            res = ObjGetElements(interp, filterElems[i], &j, &filterElem);
            if (res != TCL_OK)
                goto vamoose;
            if (j < 3 || j > 4) {
                res = TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid -filter argument value");
                goto vamoose;
            }

            filters[i].negate = 0;
            filters[i].nocase = 0;
            if (j == 4) {
                char *s = ObjToString(filterElem[3]);
                if (STREQ("-nocase", s))
                    filters[i].nocase = 1;
                else {
                    res = TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid -filter argument value");
                    goto vamoose;
                }
            }
            if ((res=ObjToEnum(interp, raObj[0], filterElem[0], &filters[i].filter_pos)) != TCL_OK
            ||
                (res = Tcl_GetIndexFromObj(interp, filterElem[1], filter_ops, "operator", TCL_EXACT, &filters[i].filter_op)) != TCL_OK) {
                goto vamoose;
            }
            switch (filters[i].filter_op) {
            case RA_NE: filters[i].negate = 1; /* FALLTHRU */
            case RA_EQ: /* TBD - should we do unicode compares? */
                filters[i].cmpfn = filters[i].nocase ? lstrcmpiA : lstrcmpA;
                filters[i].operand.string = ObjToString(filterElem[2]);
            break;
            case RA_NE_INT: filters[i].negate = 1; /* FALLTHRU */
            case RA_EQ_INT:
                if ((res = ObjToWideInt(interp, filterElem[2], &filters[i].operand.wide)) != TCL_OK)
                    goto vamoose;
                break;
            case RA_NOMATCH: filters[i].negate = 1; /* FALLTHRU */
            case RA_MATCH:
                filters[i].cmpfn = filters[i].nocase ? TwapiGlobCmpCase : TwapiGlobCmp;
                filters[i].operand.string = ObjToString(filterElem[2]);
                break;
            }
        }
    }

    if (sliceObj) {
        /* Get list of fields to include in slice */
        if ((res = ObjGetElements(interp, sliceObj,
                                  &nslice_fields, &slice_fields)) != TCL_OK)
            goto vamoose;

        slice_fieldindices = MemLifoAlloc(ticP->memlifoP, nslice_fields*sizeof(int), NULL);
        slice_values = MemLifoAlloc(ticP->memlifoP, nslice_fields*sizeof(Tcl_Obj*), NULL);

        for (i=0; i < nslice_fields; ++i) {
            /* Note use raObj[0] here, as (a) fieldsObj not init'ed yet,
               AND we want to pass the original, not dup, to ObjToEnum
               for performance reasons (it checks for cached values)
            */
            res = ObjToEnum(interp, raObj[0], slice_fields[i], &slice_fieldindices[i]);
            if (res != TCL_OK)
                goto vamoose;
        }
        /*
         * At this point,
         * slice_fieldindices[] maps field pos in slice to field pos in
         *   original recordarray
         */
        newfieldsObj = ObjNewList(nslice_fields, slice_fields);
        ObjIncrRefs(newfieldsObj); /* NEEDED for correct vamoosing! */
    }

    
    /*
     * We do not want recs[] and fields shimmering so dup first
     * and then only access via dups, not originals.
     */
    recsObj = ObjDuplicate(raObj[1]);
    ObjIncrRefs(recsObj);
    if ((res = ObjGetElements(interp, recsObj, &nrecs, &recs)) != TCL_OK)
        goto vamoose;
    if (nrecs == 0) {
        /* Return empty result. res is already TCL_OK */
        goto vamoose;           /* TBD - is empty result valid for all cases? */
    }
    fieldsObj = ObjDuplicate(raObj[0]);
    ObjIncrRefs(fieldsObj);
    if ((res = ObjGetElements(interp, fieldsObj, &nfields, &fields)) != TCL_OK)
        goto vamoose;
    raObj = NULL;              /* So we do not inadvertently use it */

    if (first)
        output = new_rec;
    else {
        i = nrecs * sizeof(Tcl_Obj*);
        if (keyfield_pos >= 0)
            i *= 2; /* Need twice the space for a dictionary output */
        output = MemLifoAlloc(ticP->memlifoP, i , NULL);
    }

    for (output_count = 0, i = 0; i < nrecs; ++i) {
        int matched;
        matched = 1;
        for (j = 0; j < nfilters; ++j) {
            Tcl_Obj *valueObj;
            int filter_op;

            TWAPI_ASSERT(filters);
            filter_op = filters[j].filter_op;

            /* filter_pos gives position of field to match */
            res = ObjListIndex(interp, recs[i], filters[j].filter_pos, &valueObj);
            if (res != TCL_OK)
                break;
            if (valueObj == NULL) {
                res = TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "too few values in record");
                break;
            }

            if (filter_op == RA_EQ_INT || filter_op == RA_NE_INT) {
                Tcl_WideInt wide;
                /* Note not-an-int is treated as no match, not as error */
                if (ObjToWideInt(NULL, valueObj, &wide) != TCL_OK ||
                    ((wide == filters[j].operand.wide) == filters[j].negate)) {
                    matched = 0;
                    break;
                }
            } else {
                if ((0 == filters[j].cmpfn(ObjToString(valueObj), filters[j].operand.string)) == filters[j].negate) {
                    matched = 0;
                    break;
                }
            }
        }

        if (res != TCL_OK)
            break;

        if (! matched)
            continue;

        /* We have a match */

        /* Add in the dictionary key if so specified */
        if (keyfield_pos >= 0) {
            Tcl_Obj *keyObj;
            TWAPI_ASSERT(format == RA_DICT || format == RA_LIST);
            res = ObjListIndex(interp, recs[i], keyfield_pos, &keyObj);
            if (res != TCL_OK)
                break;
            if (keyObj == NULL) {
                res = TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "too few values in record");
                break;
            }
            output[output_count++] = keyObj;
        }
        if (sliceObj == NULL) {
            if (format == RA_DICT) {
                output[output_count++] = TwapiTwine(NULL, fieldsObj, recs[i]);
                TWAPI_ASSERT(output[output_count]);
            } else
                output[output_count++] = recs[i];
        } else {
            /* Make a new obj based on slice. recs[] contains source records.
               slice_fieldindices[] contains the field position in the source
               records to be picked up.
            */
            Tcl_Obj **values;
            int nvalues;
            res = ObjGetElements(interp, recs[i], &nvalues, &values);
            if (res != TCL_OK)
                break;
            if (nvalues < nfields) {
                res = TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "too few values in record");
                break;
            }
            for (j = 0; j < nslice_fields; ++j) {
                TWAPI_ASSERT(slice_fieldsindices[j] < nvalues); /* Follows from prior checks */
                slice_values[j] = values[slice_fieldindices[j]];
            }
            if (format == RA_DICT)
                output[output_count++] = TwapiTwineObjv(slice_fields, slice_values, nslice_fields);
            else
                output[output_count++] = ObjNewList(nslice_fields, slice_values);;
        }
        if (first) break;
    }

    if (res != TCL_OK) {
        if (sliceObj)
            ObjDecrArrayRefs(output_count, output);
        goto vamoose;
    }

    /* output[] contains output_count records to be returned */
    /* Figure out output format */
    if (0 && first) {
        TWAPI_ASSERT(output == &new_rec[0]);
        TWAPI_ASSERT(output_count <= 2);
        if (output_count)
            ObjSetResult(interp, ObjNewList(output_count, new_rec));
        /* else empty result */
    } else {
        Tcl_Obj *resultObjs[2];
        Tcl_Obj *resultObj;
        switch (format) {
        case RA_FLAT:
            resultObj = ObjNewList(output_count * (sliceObj ? nslice_fields : nfields), NULL);
            for (i = 0; i < output_count; ++i)
                Tcl_ListObjAppendList(NULL, resultObj, output[i]);
            break;
        case RA_DICT: /* FALLTHRU as output constructed for RA_DICT above */
        case RA_LIST:
            /* Note output includes key fields also if so specified earlier */
            resultObj = ObjNewList(output_count, output);
            break;
        case RA_ARRAY:
        default:
            if (sliceObj) {
                TWAPI_ASSERT(newfieldsObj);
                resultObjs[0] = newfieldsObj;
            } else
                resultObjs[0] = fieldsObj;
            resultObjs[1] = ObjNewList(output_count, output);
            resultObj = ObjNewList(2, resultObjs);
            break;
        }
        ObjSetResult(interp, resultObj);
    }

vamoose:
    if (filterObj)
        ObjDecrRefs(filterObj);
    if (recsObj)
        ObjDecrRefs(recsObj);
    if (fieldsObj)
        ObjDecrRefs(fieldsObj);
    if (newfieldsObj)
        ObjDecrRefs(newfieldsObj);
    if (mark)
        MemLifoPopMark(mark);
    return res;
}

static Tcl_Obj *RecordGetField(Tcl_Interp *interp, Tcl_Obj *fieldsObj, Tcl_Obj *recObj, Tcl_Obj *fieldObj)
{
    int field_index;
    Tcl_Obj *objP;

    if (ObjToEnum(interp, fieldsObj, fieldObj, &field_index) == TCL_OK &&
        ObjListIndex(interp, recObj, field_index, &objP) == TCL_OK) {
        if (objP)
            return objP;
        TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "too few values in record");
    }
    return NULL;
}

static TCL_RESULT RecordInstanceObjCmd(
    ClientData clientdata,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Obj *fieldsObj = clientdata;
    Tcl_Obj **fields;
    Tcl_Obj *objP;
    int field_index;
    int i, j, cmd;
    static const char *cmds[] = {
        "get",
        "set",
        "select",
        NULL,
    };

    objP = NULL;
    switch (objc) {
    case 1:
        objP = fieldsObj;
        break;
    case 2:
        if (ObjListLength(interp, fieldsObj, &i) != TCL_OK ||
            ObjListLength(interp, objv[1], &j) != TCL_OK)
            break;
        if (i != j) {
            ObjSetStaticResult(interp, "Number of field values not equal to number of fields");
            break;
        }
        objP = TwapiTwine(interp, fieldsObj, objv[1]);
        break;
    case 3:
        objP = RecordGetField(interp, fieldsObj, objv[2], objv[1]);
        break;
    default:
        if (Tcl_GetIndexFromObj(interp, objv[1], cmds, "command", TCL_EXACT, &cmd) != TCL_OK)
            break;
        switch (cmd) {
        case 0: // get RECORD FIELD
            if (objc != 4)
                goto nargs_error;
            objP = RecordGetField(interp, fieldsObj, objv[2], objv[3]);
            break;
        case 1: // set RECORD FIELD VALUE
            if (objc != 5)
                goto nargs_error;
            if (ObjToEnum(interp, fieldsObj, objv[3], &field_index) != TCL_OK ||
                ObjListLength(interp, objv[2], &i) != TCL_OK)
                break;
            if (i <= field_index) {
                ObjSetStaticResult(interp, "too few values in record");
                break;
            }
            if (Tcl_IsShared(objv[2]))
                objP = ObjDuplicate(objv[2]);
            else
                objP = objv[2];
            Tcl_ListObjReplace(interp, objP, field_index, 1, 1, &objv[4]);
            break;
        case 2: // select RECORD FIELDLIST
            if (objc != 4)
                goto nargs_error;
            if (ObjGetElements(interp, objv[3], &j, &fields) != TCL_OK)
                break;
            objP = ObjNewList(j, NULL);
            for (i = 0; i < j; ++i) {
                Tcl_Obj *obj2P = RecordGetField(interp, fieldsObj, objv[2], fields[i]);
                if (obj2P == NULL) {
                    ObjDecrRefs(objP);
                    objP = NULL;
                    break;
                }
                ObjAppendElement(NULL, objP, obj2P);
            }
            break;
        }
        break;
    }

    if (objP) {
        ObjSetResult(interp, objP);
        return TCL_OK;;
    } else
        return TCL_ERROR;       /* interp should already hold error */

nargs_error:
    Tcl_WrongNumArgs(interp, 1, objv, "command RECORD PARAM ?PARAM ...?");
    return TCL_ERROR;
}

static void RecordInstanceObjCmdDelete(ClientData fieldsObj)
{
    ObjDecrRefs(fieldsObj);
}


/* TBD - document */
TCL_RESULT Twapi_RecordObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    int len;
    TCL_RESULT res;
    Tcl_Obj *nameObj, *fieldsObj;
    Tcl_Namespace *nsP;
    char *sep;

    switch (objc) {
    case 2: fieldsObj = objv[1]; break;
    case 3: fieldsObj = objv[2]; break;
    default:
        Tcl_WrongNumArgs(interp, 1, objv, "?RECORDNAME? FIELDS");
        return TCL_ERROR;
    }
    
    res = ObjListLength(interp, fieldsObj, &len); 
    if (res != TCL_OK)
        return res;
    if (len == 0) {
        ObjSetStaticResult(interp, "empty record definition");
        return TCL_ERROR;
    }

    nsP = Tcl_GetCurrentNamespace(interp);
    sep = nsP->parentPtr == NULL ? "" : "::";
    if (objc < 3) 
        nameObj = Tcl_ObjPrintf("%s%srecord%d", nsP->fullName, sep, Twapi_NewId());
    else {
        char *nameP = ObjToString(objv[1]);
        if (nameP[0] == ':' && nameP[1] == ':')
            nameObj = objv[1];
        else
            nameObj = Tcl_ObjPrintf("%s%s%s", nsP->fullName, sep, nameP);
    }

    /*
     * The record instance command will be passed the fields object
     * as clientdata so make sure it does not go away. Corresponding
     * ObjDecrRefs is in RecordInstanceObjCmdDelete
     */
    ObjIncrRefs(fieldsObj);

    Tcl_CreateObjCommand(interp, ObjToString(nameObj), RecordInstanceObjCmd,
                         fieldsObj, RecordInstanceObjCmdDelete);

    ObjSetResult(interp, nameObj);
    return TCL_OK;
}
