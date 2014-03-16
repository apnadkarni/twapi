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
        "recordarray", "flat", "list", "dict", "index", NULL
    };
    enum format_enum {RA_ARRAY, RA_FLAT, RA_LIST, RA_DICT, RA_INDEX};
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
    int keyfield_pos;            /* Position of the key field */
    int first = 0;               /* If true, only first match returned */
    union {
        Tcl_WideInt wide;
        char *string;
    } operand;
    int (WINAPI *cmpfn) (const char *, const char *) = lstrcmpA;
    TCL_RESULT res;

    Tcl_Obj *new_rec;
    Tcl_Obj **output = &new_rec;
    int output_count;

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

    /*
     * recordarray REC
     *  Returns the values list
     *
     * recordarray options REC
     *   -slice FIELDNAMEPAIRS
     *      Returns only those fields that are included in FIELDNAMEPAIRS
     *      each element of which is {FIELDNAME ?RENAMEDFIELD?) in
     *      the order specified
     *   -format [recordarray | flat | list | dict | index]
     *      recordarray - return value is in recordarray format (default)
     *      flat - all records are concatenated and returned as
     *             a flat list of values
     *      list - each returned record list
     *      dict - each returned record is a dict with keys being field names
     *   -select FIELDNAME OPERATOR OPERAND
     *      Only those records whose field FIELDNAME match OPERAND using
     *      the given OPERATOR are returned
     *   -key KEYFIELD
     *      Only used if -format is specified as 'list' or 'dict'.
     *      The returned value is a dictionary with KEYFIELD as the key
     *   -first
     *      Only returns the first matching record
     *   -nocase
     *      Text based comparisons are done in case-insensitive manner
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

    keyfield_pos = -1;
    /* Key field is ignored unless output is RA_LIST or RA_DICT */
    if (keyfieldObj && (format == RA_LIST || format == RA_DICT)) {
        if ((res=ObjToEnum(interp, raObj[0], keyfieldObj, &keyfield_pos)) != TCL_OK)
            return res;
    }

    /* If selection criteria are given, find index of field to match on */
    select_pos = -1;
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
            /* Note use raObj[0] here, as (a) fieldsObj not init'ed yet,
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
     * We do not want recs[] and fields shimmering so dup first
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
        output = &new_rec;
    else {
        i = nrecs * sizeof(Tcl_Obj*);
        if (keyfield_pos >= 0)
            i *= 2; /* Need twice the space for a dictionary output */
        output = MemLifoAlloc(ticP->memlifoP, i , NULL);
    }

    for (output_count = 0, i = 0; i < nrecs; ++i) {
        if (select_pos >= 0) {
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
                if ((0 == cmpfn(ObjToString(valueObj), operand.string)) == negate)
                    continue;
            }
        }
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
                output[output_count++] = TwapiTwineObjv(slice_newfields, slice_values, nslice_fields);
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
    if (first) {
        TWAPI_ASSERT(output == &new_rec);
        if (output_count)
            ObjSetResult(interp, new_rec);
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
