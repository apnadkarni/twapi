/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>

int Twapi_RecordArrayObjCmd(
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
        TwapiSetStaticResult(interp, "Invalid command. Must be one of 'exists', 'field', 'filter', 'get', 'keys' or 'slice'.");
        return TCL_ERROR;
    }

    if (Tcl_ListObjGetElements(interp, objv[raindex], &i, &raObj) != TCL_OK)
        return TCL_ERROR;

    /* Special case - empty list */
    if (i == 0) {
        if (cmd == RA_EXISTS)
            TwapiSetObjResult(interp, ObjFromBoolean(0));
        else
            Tcl_ResetResult(interp);
        return TCL_OK;
    }
    if (i != 2 ||
         Tcl_ListObjGetElements(interp, raObj[1], &nrecs, &recs) != TCL_OK ||
         (nrecs & 1) != 0) {
        /* Record array must have exactly two elements
           the second of which (the key,record list) must have an even
           number of elements */
        TwapiSetStaticResult(interp, "Invalid format.");
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
            TwapiSetStaticResult(interp, "Option must be -string, -nocase, -glob, -integer.");
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
        TwapiSetStaticResult(interp, "Invalid combination of options specified.");
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
            Tcl_ListObjAppendElement(interp, resultObj, recs[i]); // Key
            if (cmd == RA_GET) {
                Tcl_Obj *objP = TwapiTwine(interp, raObj[0], recs[i+1]);
                if (objP == NULL)
                    goto error_return;
                Tcl_ListObjAppendElement(interp, resultObj, objP); // Value
            }
        }
        return TwapiSetObjResult(interp, resultObj);
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
        if (Tcl_ListObjGetElements(interp, raObj[0], &nfields, &fields) != TCL_OK)
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
            TwapiSetStaticResult(interp, "Field name not found.");
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
                if (Tcl_ListObjIndex(interp, recs[i+1], j, &objP) != TCL_OK)
                    goto error_return;
                if (objP == NULL)
                    continue;   /* This field not in the record */
                if (ObjToInt(interp, objP, &ival) != TCL_OK)
                    continue;   /* Not int value hence no match */
                if (ival != integer_key)
                    continue;   /* Field value does not match */
                /* Matched. Add the record to filtered list */
                Tcl_ListObjAppendElement(interp, resultObj, recs[i]);
                Tcl_ListObjAppendElement(interp, resultObj, recs[i+1]);
            }
        } else {
            skey = ObjToString(objv[objc-1]);
            /* Iterate over all records matching on the field */
            for (i = 0; i < nrecs; i += 2) {
                Tcl_Obj *objP;
                if (Tcl_ListObjIndex(interp, recs[i+1], j, &objP) != TCL_OK)
                    goto error_return;
                if (objP == NULL)
                    continue;   /* This field not in the record */
                if (cmpfn(ObjToString(objP), skey))
                    continue;   /* Field value does not match */
                /* Matched. Add the record to filtered list */
                Tcl_ListObjAppendElement(interp, resultObj, recs[i]);
                Tcl_ListObjAppendElement(interp, resultObj, recs[i+1]);
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
        if (Tcl_ListObjGetElements(interp, raObj[0], &nfields, &fields) != TCL_OK)
            return TCL_ERROR;
        skey = ObjToString(objv[objc-1]); /* Field name we want */
        for (j=0; j < nfields; ++j) {
            char *s = ObjToString(fields[j]);
            if (STREQ(skey, s))
                break;
        }
        if (j >= nfields) {
            TwapiSetStaticResult(interp, "Field name not found.");
            return TCL_ERROR;
        }
        if (objc == 4) {
            /* Return ALL values */
            resultObj = ObjNewList(0, NULL);
            for (i=0; i < nrecs; i += 2) {
                Tcl_Obj *objP;
                // Append key
                Tcl_ListObjAppendElement(interp, resultObj, recs[i]);
                if (Tcl_ListObjIndex(interp, recs[i+1], j, &objP) != TCL_OK)
                    goto error_return;
                // Append field value
                Tcl_ListObjAppendElement(interp, resultObj,
                                         objP ? objP : STRING_LITERAL_OBJ(""));
            }
        } else {
            /* Get value of a single element field (empty if does not exist) */
            /* Get the corresponding element from the record (or empty) */
            if (Tcl_ListObjIndex(interp, recs[i+1], j, &resultObj) != TCL_OK)
                return TCL_ERROR;
            /* Note resultObj may be NULL even on TCL_OK */
        }
        break;

    case RA_SLICE:
        /* Return a vertical slice of the array */

        /* Get list of fields in current recordarray */
        if (Tcl_ListObjGetElements(interp, raObj[0], &nfields, &fields) != TCL_OK)
            return TCL_ERROR;

        /* Get list of fields to include in slice */
        if (Tcl_ListObjGetElements(interp, objv[3],
                                   &nslice_fields, &slice_fields) != TCL_OK)
            return TCL_ERROR;
        if (nslice_fields > MAX_SLICE_WIDTH) {
            TwapiSetStaticResult(interp, "Limit on fields in slice exceeded.");
            return TCL_ERROR;
        }

        /* Figure out which columns go into the slice, and their names */
        for (i=0; i < nslice_fields; ++i) {
            int slen;
            char *s;
            Tcl_Obj **names;
            if (Tcl_ListObjGetElements(interp, slice_fields[i], &j, &names) != TCL_OK) {
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
                TwapiSetStaticResult(interp, "Invalid slice field renaming entry.");
                return TCL_ERROR;
            }
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
                TwapiSetStaticResult(interp, "Field not found.");
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
        Tcl_IncrRefCount(emptyObj);
        for (i=0; i < nrecs; i += 2) {
            Tcl_Obj **values;
            int nvalues;
            Tcl_ListObjAppendElement(interp, new_ra[1], recs[i]); // Key
            if (Tcl_ListObjGetElements(interp, recs[i+1], &nvalues, &values)
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
            Tcl_ListObjAppendElement(interp, new_ra[1],
                                     ObjNewList(nslice_fields, slice_values));
        }
        Tcl_DecrRefCount(emptyObj);
        resultObj = ObjNewList(2, new_ra);
        break;
    }

        
    if (resultObj)              /* May legitimately be NULL */
        TwapiSetObjResult(interp, resultObj);
    else
        Tcl_ResetResult(interp); /* Because some intermediate checks might
                                    have left error messages */

    return TCL_OK;

error_return:
    if (resultObj)
        Twapi_FreeNewTclObj(resultObj);
    return TCL_ERROR;
}

