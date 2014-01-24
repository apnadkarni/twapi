/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

int Twapi_RecordArrayObjCmd(
    TwapiInterpContext *ticP,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    int i, j, negate = 0, ignore_case = 0;
    Tcl_Obj **raObj;
    static const char *opts[] = {
        "-format",              /* FORMAT */
        "-slice",               /* FIELDNAMES */
        "-select",              /* OPER FIELDNAME VALUE */
        "-key",                 /* FIELDNAME */
        "-first",               /* no args */
        "-nocase",              /* no args */
        NULL
    };
    enum opts_enum {RA_FORMAT, RA_SLICE, RA_SELECT, RA_KEY, RA_FIRST, RA_NOCASE};
    int opt;
    static const char *formats[] = {
        "recordarray", "flat", "nested", "dict", "index", NULL
    };
    enum format_enum {RA_ARRAY, RA_FLAT, RA_NESTED, RA_DICT, RA_INDEX};
    int format = RA_ARRAY;
    static const char *select_ops[] = {
        "eq", "ne", "==", "!=", "~", "!~", NULL
    };
    enum select_ops_enum {RA_EQ, RA_NE, RA_EQ_INT, RA_NE_INT, RA_MATCH, RA_NOMATCH};
    int select_op = RA_EQ;

    Tcl_Obj *sliceObj = NULL,
        *selectObj = NULL,
        *operatorObj = NULL,
        *operandObj = NULL,
        *keyfieldObj = NULL;
    Tcl_Obj *recsObj = NULL;     /* Dup of records passed in */
    Tcl_Obj **recs;              /* Contents of recsObj */
    int      nrecs;              /* Number records */
    Tcl_Obj *fieldsObj = NULL;   /* Dup of record definition passed in */
    Tcl_Obj **fields;            /* Contents of recsObj */
    int      nfields;            /* Number of fields in record */
    int select_pos;              /* Position of field to use for compares */
    int first = 0;               /* If true, only first match returned */
    union {
        Tcl_WideInt wide;
        char *string;
    } operand;
    int (WINAPI *cmpfn) (const char *, const char *) = lstrcmpA;
    TCL_RESULT res;

    Tcl_Obj *new_rec;
    Tcl_Obj **new_recs = &new_rec;
    int new_count;

    Tcl_Obj **slice_fields = NULL;   /* Names of the fields to retrieve */
    int       *slice_fieldindices = NULL; /* Positions of the fields to retrieve */
    int      nslice_fields;   /* Count of above */
    Tcl_Obj  **slice_newfields = NULL;/* New names of the fields to be returned */
    Tcl_Obj  **slice_values = NULL;
    Tcl_Obj  *newfieldsObj = NULL;

    MemLifoMarkHandle mark = NULL;
    
    if (objc < 2) {
        Tcl_WrongNumArgs(interp, 1, objv, "RECORDARRAY ?OPTIONS?");
        return TCL_ERROR;
    }

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
        case RA_SELECT:         /* -select FIELDNAME OPERATOR OPERAND */
            ++i;                /* So missing value error picks correct obj */
            if ((i+2) == (objc-1))
                goto missing_value;
            selectObj = objv[i++];
            operatorObj = objv[i++];
            operandObj = objv[i];
            break;
        case RA_KEY:
            if (++i == (objc-1))
                goto missing_value;
            keyfieldObj = objv[i];
            break;
        case RA_FIRST:
            first = 1;
            break;
        case RA_NOCASE:
            ignore_case = 1;
            break;
        }
    }

    if (ObjGetElements(interp, objv[objc-1], &i, &raObj) != TCL_OK)
        return TCL_ERROR;

    if (i == 0)
        return TCL_OK; /* TBD - check  empty result is valid for all commands */

    if (i != 2)
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "Invalid recordarray format");

    /* No commands -> return values */
    if (objc == 2) {
        ObjSetResult(interp, raObj[1]);
        return TCL_OK;
    }

    /* If selection criteria are given, find index of field to match on */
    if (selectObj) {
        if ((res=ObjToEnum(interp, raObj[0], selectObj, &select_pos)) != TCL_OK
            ||
            (res = Tcl_GetIndexFromObj(interp, operatorObj, select_ops, "operator", TCL_EXACT, &select_op)) != TCL_OK)
            return res;
        switch (select_op) {
        case RA_NE: negate = 1; /* FALLTHRU */
        case RA_EQ: /* TBD - should we do unicode compares? */
            cmpfn = ignore_case ? lstrcmpiA : lstrcmpA;
            operand.string = ObjToString(operandObj);
            break;
        case RA_NE_INT: negate = 1; /* FALLTHRU */
        case RA_EQ_INT:
            if ((res = ObjToWideInt(interp, operandObj, &operand.wide)) != TCL_OK)
                goto vamoose;
            break;
        case RA_NOMATCH: negate = 1; /* FALLTHRU */
        case RA_MATCH:
            cmpfn = ignore_case ? TwapiGlobCmpCase : TwapiGlobCmp;
            operand.string = ObjToString(operandObj);
            break;
        }
    }

    if (ObjGetElements(interp, objv[objc-1], &i, &raObj) != TCL_OK)
        return TCL_ERROR;

    if (i == 0)
        return TCL_OK; /* TBD - check  empty result is valid for all commands */

    if (i != 2)
        return TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "Invalid recordarray format");

    /* No commands -> return values */
    if (objc == 2) {
        ObjSetResult(interp, raObj[1]);
        return TCL_OK;
    }

    /* Any exits hereon must MemLifoPopFrame */
    mark = MemLifoPushMark(ticP->memlifoP);

    if (sliceObj) {
        /* Get list of fields to include in slice */
        if ((res = ObjGetElements(interp, sliceObj,
                                  &nslice_fields, &slice_fields)) != TCL_OK)
            goto vamoose;

        slice_newfields = MemLifoAlloc(ticP->memlifoP, nslice_fields*sizeof(Tcl_Obj*), NULL);
        slice_fieldindices = MemLifoAlloc(ticP->memlifoP, nslice_fields*sizeof(int), NULL);
        slice_values = MemLifoAlloc(ticP->memlifoP, nslice_fields*sizeof(Tcl_Obj*), NULL);

        /* Figure out which columns go into the slice, and their names */
        for (i=0; i < nslice_fields; ++i) {
            int slen;
            char *s;
            Tcl_Obj **names;
            res = ObjGetElements(interp, slice_fields[i], &j, &names);
            if (res != TCL_OK)
                goto vamoose;
            if (j == 1)
                slice_newfields[i] = names[0]; /* Field name unchanged */
            else if (j == 2)
                slice_newfields[i] = names[1]; /* Field name is to be changed */
            else {
                /* Should be just 1 or 2 elements in renaming entry */
                res = TwapiReturnErrorMsg(interp, TWAPI_INVALID_ARGS, "Invalid slice field renaming entry.");
            }
            /* Note use raObj[0] here, as (a) fieldObj not init'ed yet,
               AND we want to pass the original, not dup, to ObjToEnum
               for performance reasons (it checks for cached values)
            */
            res = ObjToEnum(interp, raObj[0], names[0], &slice_fieldindices[i]);
            if (res != TCL_OK)
                goto vamoose;
            
        }
        /*
         * At this point,
         * slice_fieldindices[] maps field pos in slice to field pos in
         *   original recordarray
         * slice_newfields[] contains the new names of the fields to
         *   be returned in the slice.
         */
        newfieldsObj = ObjNewList(nslice_fields, slice_newfields);
        ObjIncrRefs(newfieldsObj); /* NEEDED for correct vamoosing! */
    }

    /*
     * From this point, do not want recs[] and fields shimmering so dup first
     * and then only access via dups, not originals.
     */
    recsObj = ObjDuplicate(raObj[1]);
    ObjIncrRefs(recsObj);
    if ((res = ObjGetElements(interp, recsObj, &nrecs, &recs)) != TCL_OK)
        goto vamoose;
    fieldsObj = ObjDuplicate(raObj[0]);
    ObjIncrRefs(fieldsObj);
    if ((res = ObjGetElements(interp, fieldsObj, &nfields, &fields)) != TCL_OK)
        goto vamoose;
    raObj = NULL;              /* So we do not inadvertently use it */
    
    if (first)
        new_recs = &new_rec;
    else
        new_recs = MemLifoAlloc(ticP->memlifoP,
                                nrecs * sizeof(Tcl_Obj*),
                                NULL);

    for (new_count = 0, i = 0; i < nrecs; ++i) {
        if (selectObj) {
            Tcl_Obj *valueObj;
            /* select_pos gives position of field to match */
            res = ObjListIndex(interp, recs[i], select_pos, &valueObj);
            if (res != TCL_OK)
                break;
            if (valueObj == NULL) {
                res = TwapiReturnErrorMsg(interp, TWAPI_INVALID_DATA, "too few values in record");
                break;
            }

            if (select_op == RA_EQ_INT || select_op == RA_NE_INT) {
                Tcl_WideInt wide;
                /* Note not-an-int is treated as no match, not as error */
                if (ObjToWideInt(NULL, valueObj, &wide) != TCL_OK)
                    continue;
                if ((wide == operand.wide) == negate)
                    continue;
            } else {
                if (cmpfn(ObjToString(valueObj), operand.string) != negate)
                    continue;
            }
        }
        /* We have a match */
        if (sliceObj == NULL)
            new_recs[new_count++] = recs[i];
        else {
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
            new_recs[new_count++] = ObjNewList(nslice_fields, slice_values);
        }
        if (first) break;
    }

    if (res != TCL_OK) {
        if (sliceObj)
            ObjDecrArrayRefs(new_count, new_recs);
        goto vamoose;
    }

    /* new_recs[] contains new_count records to be returned */
    /* Figure out output format */
    if (first) {
        TWAPI_ASSERT(new_recs == &new_rec);
        if (new_count)
            ObjSetResult(interp, new_rec);
        /* else empty result */
    } else {
        Tcl_Obj *resultObjs[2];
        if (sliceObj) {
            TWAPI_ASSERT(newfieldsObj);
            resultObjs[0] = newfieldsObj;
        } else
            resultObjs[0] = fieldsObj;
        resultObjs[1] = ObjNewList(new_count, new_recs);
        ObjSetResult(interp, ObjNewList(2, resultObjs));
    }

vamoose:
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

int Twapi_OBSOLETERecordArrayObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    Tcl_Obj *resultObj = NULL;
    int      raindex;
    int      keyindex;
    int      integer_key;
    char    *skey;
    Tcl_Obj **raObj;
    Tcl_Obj **recs;
    int      nrecs;
    int      i, j;
    Tcl_Obj **fields;
    int       nfields;
    char     *cmdstr;
    enum {RA_GET, RA_FIELD, RA_EXISTS, RA_KEYS, RA_FILTER, RA_SLICE} cmd;
    int   cmptype;
    int (WINAPI *cmpfn) (const char *, const char *) = lstrcmpA;
    Tcl_Obj *new_ra[2];
    Tcl_Obj **slice_fields;
    int      nslice_fields;
#if 0
    Tcl_Obj **slice_renamedfields;
#endif
#define MAX_SLICE_WIDTH 64
    Tcl_Obj  *slice_newfields[MAX_SLICE_WIDTH];
    Tcl_Obj  *slice_values[MAX_SLICE_WIDTH];
    int       slice_fieldindices[MAX_SLICE_WIDTH];
    Tcl_Obj  *emptyObj;

    /*
     * A record array consists of an empty list or a list with two sublists.
     * The first sublist is an ordered list of field names in a record. The
     * second sublist is a list of alternating key and record value
     * elements. The record value is a list of values in the same order
     * as the field names.
     *
     * Command forms:
     *   recordarray get RECORDARRAY
     *     Returns entire record array as a list of key and record values,
     *     the latter formatted as keyed lists
     *
     *   recordarray exists ?OPTIONS? RECORDARRAY KEY
     *     Returns 1 if KEY exists and 0 otherwise.
     *
     *   recordarray get ?OPTIONS? RECORDARRAY KEY
     *     Returns the record value corresponding to KEY as a keyed list.
     *
     *   recordarray keys RECORDARRAY
     *     Returns the list of keys.
     *
     *   recordarray field ?OPTIONS? RECORDARRAY ?KEY? FIELD
     *     Returns the corresponding field in a record as a scalar value
     *     if KEY is specified. Otherwise returns a list containing
     *     key and value for FIELD for all records. Note if ?OPTIONS?
     *     specified, ?KEY? must also be specified.
     *
     *   recordarray filter ?OPTIONS? RECORDARRAY FIELD FIELDVALUE
     *     Returns a record array containing only those records
     *     whose FIELD has value FIELDVALUE.
     *
     *   recordarray slice RECORDARRAY FIELDS
     *     Returns a record array containing all records but with only 
     *     the specified fields. FIELDS is a list of fields to include
     *     each element may be a pair if field is to be renamed.
     *     
     * Following options available:
     *  -integer : compare as integer
     *  -string  : compare as string
     *  -nocase  : ignore case in comparisons (implies -string)
     *  -glob  : do glob matching (implies -string)
     *
     * Notes:
     * - implementation is simplistic and not mean to scale particularly
     *   well. It's not possible to do so given that we do not place
     *   restrictions on how the record array is created, key ordering etc.
     * - if duplicate keys exist, first occurence wins.
     * - the above command syntax is such that the RECORDARRAY field
     *   can always be located irrespective of any options. This is
     *   important as we do not want to shimmer the potentially large
     *   RECORDARRAY list into a string format.
     */
    
    if (objc < 3) {
    badargcount:
        Tcl_WrongNumArgs(interp, 1, objv, "COMMAND ?OPTIONS? RECORDARRAY ?ARGS?");
        return TCL_ERROR;
    }

    /* TBD - use Tcl_GetObjIndex if possible for efficiency */

    keyindex = -1;
    cmdstr = ObjToString(objv[1]);
    if (STREQ("get", cmdstr)) {
        cmd = RA_GET;
        if (objc == 3) {
            raindex = 2;
        } else {
            if (objc < 4)
                goto badargcount;
            raindex = objc-2;
            keyindex = objc-1;
        }
    } else if (STREQ("field", cmdstr)) {
        cmd = RA_FIELD;
        if (objc == 4) {
            /* recordarray field RECORDARRAY FIELDNAME */
            /* No key specified, get all fields */
            raindex = objc-2;
        } else if (objc == 5) {
            /* recordarray field RECORDARRAY KEY FIELDNAME
             * Note this is not ambiguous because
             * "recordarray field OPTION RECORDARRAY FIELDNAME"
             * having OPTION without KEY would not make sense.
             */
            raindex = objc-3;
            keyindex = objc-2;
        } else if (objc == 6) {
            /* recordarray field OPTION RECORDARRAY KEY FIELDNAME */
            raindex = objc-3;
            keyindex = objc-2;
        } else
            goto badargcount;
    } else if (STREQ("exists", cmdstr)) {
        if (objc < 4)
            goto badargcount;
        cmd = RA_EXISTS;
        raindex = objc-2;
        keyindex = objc-1;
    } else if (STREQ("keys", cmdstr)) {
        cmd = RA_KEYS;
        raindex = 2;
    } else if (STREQ("filter", cmdstr)) {
        if (objc < 5)
            goto badargcount;
        cmd = RA_FILTER;
        raindex = objc-3;
    } else if (STREQ("slice", cmdstr)) {
        if (objc != 4)
            goto badargcount;
        cmd = RA_SLICE;
        raindex = 2;
    } else {
        ObjSetStaticResult(interp, "Invalid command. Must be one of 'exists', 'field', 'filter', 'get', 'keys' or 'slice'.");
        return TCL_ERROR;
    }

    if (ObjGetElements(interp, objv[raindex], &i, &raObj) != TCL_OK)
        return TCL_ERROR;

    /* Special case - empty list */
    if (i == 0) {
        if (cmd == RA_EXISTS)
            ObjSetResult(interp, ObjFromBoolean(0));
        else
            Tcl_ResetResult(interp);
        return TCL_OK;
    }
    if (i != 2 ||
         ObjGetElements(interp, raObj[1], &nrecs, &recs) != TCL_OK ||
         (nrecs & 1) != 0) {
        /* Record array must have exactly two elements
           the second of which (the key,record list) must have an even
           number of elements */
        ObjSetStaticResult(interp, "Invalid format.");
        return TCL_ERROR;
    }

    /* Parse any options */
    cmptype = 0;
    for (i=2; i < raindex; ++i) {
        char *s = ObjToString(objv[i]);
        if (STREQ("-integer", s))
            cmptype = 1;
        else if (STREQ("-string", s))
            cmptype = 0;
        else if (STREQ("-nocase", s))
            cmptype |= 0x2;
        else if (STREQ("-glob", s))
            cmptype |= 0x4;
        else {
            ObjSetStaticResult(interp, "Option must be -string, -nocase, -glob, -integer.");
            return TCL_ERROR;
        }
    }

    switch (cmptype) {
    case 0:  cmpfn = lstrcmpA; break; /* Exact string compare */
    case 1:  break;           /* Int compare, cmpfn does not matter */
    case 2:  cmpfn = lstrcmpiA; break;     /* Case insensitive string compare */
    case 4:  cmpfn = TwapiGlobCmp; break; /* Case sensitive glob match */
    case 6:  cmpfn = TwapiGlobCmpCase; break;  /* Case insensitive flob match */
    default:
        ObjSetStaticResult(interp, "Invalid combination of options specified.");
        return TCL_ERROR;
    }

    /*
     * raObj[0] - List of fields
     * raObj[1] - List of records
     * recs[] - List of KEY RECORD extracted from raObj[1]
     * nrecs - size of recs[] 
     */

    if (cmd == RA_KEYS || (cmd == RA_GET && objc == 3)) {
        resultObj = ObjNewList(0, NULL);
        for (i=0; i < nrecs; i += 2) {
            ObjAppendElement(interp, resultObj, recs[i]); // Key
            if (cmd == RA_GET) {
                Tcl_Obj *objP = TwapiTwine(interp, raObj[0], recs[i+1]);
                if (objP == NULL)
                    goto error_return;
                ObjAppendElement(interp, resultObj, objP); // Value
            }
        }
        return ObjSetResult(interp, resultObj);
    }

    /* Locate the key. */
    if (keyindex == -1) {
        /* Command does not operate on a key, so no need to locate */
        i = nrecs;
    } else {
        if (cmptype == 1) {
            if (ObjToInt(interp, objv[keyindex], &integer_key) != TCL_OK)
                return TCL_ERROR;
            for (i=0; i < nrecs; i += 2) {
                int ival;
                /* If a key is not integer, ignore it, don't generate error */
                if (ObjToInt(interp, recs[i], &ival) == TCL_OK &&
                    ival == integer_key) {
                    /* Found it */
                    break;
                }
            }
        } else {
            skey = ObjToString(objv[keyindex]);
            for (i=0; i < nrecs; i += 2) {
                if (!cmpfn(ObjToString(recs[i]), skey))
                    break;
            }
        }
    }

    /* i is index of matched key, if any, or >= nrecs on no match or
     * we are not looking for a key
     */

    switch (cmd) {
    case RA_EXISTS:
        resultObj = ObjFromBoolean(i < nrecs ? 1 : 0);
        break;
    case RA_GET:
        /* i - index of matched key, i+1 - index of corresponding record */
        if (i < nrecs) {
            resultObj = TwapiTwine(interp, raObj[0], recs[i+1]);
            if (resultObj == NULL)
                return TCL_ERROR;
        }
        break;

    case RA_FILTER:
        /* Find position of the requested field as j */
        if (ObjGetElements(interp, raObj[0], &nfields, &fields) != TCL_OK)
            return TCL_ERROR;

        /* For field names we don't use Unicode compares as they are
           likely to be ASCII */
        skey = ObjToString(objv[objc-2]); /* Field name we want */
        for (j=0; j < nfields; ++j) {
            char *s = ObjToString(fields[j]);
            if (STREQ(skey, s))
                break;
        }
        if (j >= nfields) {
            ObjSetStaticResult(interp, "Field name not found.");
            return TCL_ERROR;
        }

        /* j now holds position of the field we have to check */

        resultObj = ObjNewList(0, NULL); /* Will hold records, is released on error */
        if (cmptype == 1) {
            /* Do comparisons as integer */
            if (ObjToInt(interp, objv[objc-1], &integer_key) != TCL_OK)
                goto error_return; /* Need to deallocate resultObj */
            /* Iterate over all records matching on the field */
            for (i = 0; i < nrecs; i += 2) {
                Tcl_Obj *objP;
                int ival;
                if (ObjListIndex(interp, recs[i+1], j, &objP) != TCL_OK)
                    goto error_return;
                if (objP == NULL)
                    continue;   /* This field not in the record */
                if (ObjToInt(interp, objP, &ival) != TCL_OK)
                    continue;   /* Not int value hence no match */
                if (ival != integer_key)
                    continue;   /* Field value does not match */
                /* Matched. Add the record to filtered list */
                ObjAppendElement(interp, resultObj, recs[i]);
                ObjAppendElement(interp, resultObj, recs[i+1]);
            }
        } else {
            skey = ObjToString(objv[objc-1]);
            /* Iterate over all records matching on the field */
            for (i = 0; i < nrecs; i += 2) {
                Tcl_Obj *objP;
                if (ObjListIndex(interp, recs[i+1], j, &objP) != TCL_OK)
                    goto error_return;
                if (objP == NULL)
                    continue;   /* This field not in the record */
                if (cmpfn(ObjToString(objP), skey))
                    continue;   /* Field value does not match */
                /* Matched. Add the record to filtered list */
                ObjAppendElement(interp, resultObj, recs[i]);
                ObjAppendElement(interp, resultObj, recs[i+1]);
            }
        }

        // Construct the filtered recordarray
        // raObj[0] contains field names
        // NOTE WE CANNOT JUST REUSE raobj ARRAY AND OVERWRITE raObj[1] AS
        // raObj POINTS INTO THE LIST object. We have to use a separate array.
        new_ra[0] = raObj[0];
        new_ra[1] = resultObj;
        resultObj = ObjNewList(2, new_ra);
        break;

    case RA_FIELD:
        /* Find position of the requested field as j */
        /* TBD - use ObjToEnum ? */
        if (ObjGetElements(interp, raObj[0], &nfields, &fields) != TCL_OK)
            return TCL_ERROR;
        skey = ObjToString(objv[objc-1]); /* Field name we want */
        for (j=0; j < nfields; ++j) {
            char *s = ObjToString(fields[j]);
            if (STREQ(skey, s))
                break;
        }
        if (j >= nfields) {
            ObjSetStaticResult(interp, "Field name not found.");
            return TCL_ERROR;
        }
        if (objc == 4) {
            /* Return ALL values */
            resultObj = ObjNewList(0, NULL);
            for (i=0; i < nrecs; i += 2) {
                Tcl_Obj *objP;
                // Append key
                ObjAppendElement(interp, resultObj, recs[i]);
                if (ObjListIndex(interp, recs[i+1], j, &objP) != TCL_OK)
                    goto error_return;
                // Append field value
                ObjAppendElement(interp, resultObj,
                                         objP ? objP : STRING_LITERAL_OBJ(""));
            }
        } else {
            /* Get value of a single element field (empty if does not exist) */
            /* Get the corresponding element from the record (or empty) */
            if (ObjListIndex(interp, recs[i+1], j, &resultObj) != TCL_OK)
                return TCL_ERROR;
            /* Note resultObj may be NULL even on TCL_OK */
        }
        break;

    case RA_SLICE:
        /* Return a vertical slice of the array */

        /* Get list of fields in current recordarray */
        if (ObjGetElements(interp, raObj[0], &nfields, &fields) != TCL_OK)
            return TCL_ERROR;

        /* Get list of fields to include in slice */
        if (ObjGetElements(interp, objv[3],
                                   &nslice_fields, &slice_fields) != TCL_OK)
            return TCL_ERROR;
        if (nslice_fields > MAX_SLICE_WIDTH) {
            ObjSetStaticResult(interp, "Limit on fields in slice exceeded.");
            return TCL_ERROR;
        }

        /* Figure out which columns go into the slice, and their names */
        for (i=0; i < nslice_fields; ++i) {
            int slen;
            char *s;
            Tcl_Obj **names;
            if (ObjGetElements(interp, slice_fields[i], &j, &names) != TCL_OK) {
                return TCL_ERROR;
            }
            if (j == 1) {
                /* Field name in slice is same as original */
                slice_newfields[i] = names[0];
            } else if (j == 2) {
                /* Field name is to be changed */
                slice_newfields[i] = names[1];
            } else {
                /* Should be just 1 or 2 elements in renaming entry */
                ObjSetStaticResult(interp, "Invalid slice field renaming entry.");
                return TCL_ERROR;
            }
            /* TBD - maybe use ObjToEnum ? */
            s = ObjToStringN(names[0], &slen);
            for (j = 0; j < nfields; ++j) {
                char *f;
                int flen;
                f = ObjToStringN(fields[j], &flen);
                if (flen == slen && !strcmp(s, f)) {
                    slice_fieldindices[i] = j;
                    break;
                }
            }
            if (j == nfields) {
                ObjSetStaticResult(interp, "Field not found.");
                return TCL_ERROR;
            }
        }
        /*
         * At this point,
         * slice_fieldindices[] maps field pos in slice to field pos in
         *   original recordarray
         * slice_newfields[] contains the names of the fields to
         *   be returned in the slice.
         * Now just loop through records and collect everything.
         */

        // Construct the sliced recordarray
        // raObj[0] contains field names
        // NOTE WE CANNOT JUST REUSE raobj ARRAY AND OVERWRITE raObj[1] AS
        // raObj POINTS INTO THE LIST object. We have to use a separate array.
        new_ra[0] = ObjNewList(nslice_fields, slice_newfields);
        new_ra[1] = ObjNewList(0, NULL);
        emptyObj = STRING_LITERAL_OBJ(""); /* So we don't create multiple empty objs */
        ObjIncrRefs(emptyObj);
        for (i=0; i < nrecs; i += 2) {
            Tcl_Obj **values;
            int nvalues;
            ObjAppendElement(interp, new_ra[1], recs[i]); // Key
            if (ObjGetElements(interp, recs[i+1], &nvalues, &values)
                != TCL_OK) {
                return TCL_ERROR;
            }

            for (j = 0; j < nslice_fields; ++j) {
                if (slice_fieldindices[j] >= nvalues) {
                    slice_values[j] = emptyObj;
                } else {
                    slice_values[j] = values[slice_fieldindices[j]];
                }
            }
            ObjAppendElement(interp, new_ra[1],
                                     ObjNewList(nslice_fields, slice_values));
        }
        ObjDecrRefs(emptyObj);
        resultObj = ObjNewList(2, new_ra);
        break;
    }

        
    if (resultObj)              /* May legitimately be NULL */
        ObjSetResult(interp, resultObj);
    else
        Tcl_ResetResult(interp); /* Because some intermediate checks might
                                    have left error messages */

    return TCL_OK;

error_return:
    if (resultObj)
        Twapi_FreeNewTclObj(resultObj);
    return TCL_ERROR;
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
    Tcl_Obj **values;
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
