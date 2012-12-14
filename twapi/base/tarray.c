#include "tcl.h"
#define TARRAY_ENABLE_ASSERT 1
#include "tarray.h"


/*
 * Various operations on bit arrays use the following definitions. 
 * Since we are using the bitarray code from stackoverflow, bit arrays
 * are treated as a sequence of unsigned chars each containing CHAR_BIT
 * bits ordered from most significant bit to least significant bit
 * within a char.
 */

/* Return a mask containing a 1 at a bit position (MSB being bit 0) 
   BITMASK(2) -> 00100000 */
#define BITMASK(pos_) ((unsigned char) ((1 << (CHAR_BIT-1)) >> (pos_)))

/* Return a mask where all bit positions up to, but not including pos
 * are 1, remaining are 0. For example, BITMASKBELOW(2) -> 11000000
 * (remember again, that MS Bit is bit 0)
 */
#define BITMASKBELOW(pos_) (- (BITMASK(pos_) + 1))
/* Similarly, a bit mask beyond, and not including bit at pos
 * BITMASKABOVE(2) -> 00011111
 */
#define BITMASKABOVE(pos_) (BITMASK(pos_) - 1)

/*
 * TArray is a Tcl "type" used for densely storing arrays of elements
 * of a specific type.
 * The Tcl_Obj.internalRep.ptrAndLongRep.value type of an element.
 * and Tcl_Obj.internalRep.ptrAndLongRep.ptr holds a pointer to 
 * an allocated array of that type.
 */
static void TArrayDupObj(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void TArrayFreeRep(Tcl_Obj *objP);
static void TArrayUpdateStringRep(Tcl_Obj *objP);
struct Tcl_ObjType gTArrayType = {
    "TArray",
    TArrayFreeRep,
    TArrayDupObj,
    TArrayUpdateStringRep,
    NULL,     /* jenglish advises to keep this NULL */
};


/* Must match definitions in tarray.h ! */
const char *gTArrayTypeTokens[] = {
    "boolean",
    "int",
    "uint",
    "wide",
    "double",
    "byte",
    "tclobj",
    NULL
};    

/*
 * Options for 'tarray search'
 */
static const char *TArraySearchSwitches[] = {
    "-all", "-inline", "-not", "-start", "-eq", "-gt", "-lt", "-pat", "-re", "-nocase", NULL
};
enum TArraySearchSwitches {
    TARRAY_SEARCH_OPT_ALL, TARRAY_SEARCH_OPT_INLINE, TARRAY_SEARCH_OPT_INVERT, TARRAY_SEARCH_OPT_START, TARRAY_SEARCH_OPT_EQ, TARRAY_SEARCH_OPT_GT, TARRAY_SEARCH_OPT_LT, TARRAY_SEARCH_OPT_PAT, TARRAY_SEARCH_OPT_RE, TARRAY_SEARCH_OPT_NOCASE
};
/* Search flags */
#define TARRAY_SEARCH_INLINE 1  /* Return values, not indices */
#define TARRAY_SEARCH_INVERT 2  /* Invert matching expression */
#define TARRAY_SEARCH_ALL    4  /* Return all matches */
#define TARRAY_SEARCH_NOCASE 8  /* Ignore case */


void TArrayTypePanic(unsigned char tatype)
{
    Tcl_Panic("Unknown or unexpected tarray type %d", tatype);
}

TCL_RESULT TArrayBadArgError(Tcl_Interp *interp, const char *optname)
{
    Tcl_SetObjResult(interp, Tcl_ObjPrintf("Missing or invalid argument to option '%s'", optname));
    Tcl_SetErrorCode(interp, "TARRAY", "ARGUMENT", NULL);
    return TCL_ERROR;
}

TCL_RESULT TArrayBadSearchOpError(Tcl_Interp *interp, enum TArraySearchSwitches op)
{
    if (interp) {
        const char *ops = NULL;
        if ((int)op < (sizeof(TArraySearchSwitches)/sizeof(TArraySearchSwitches[0])))
            ops = TArraySearchSwitches[(int)op];
        if (ops == NULL)
            Tcl_SetObjResult(interp, Tcl_ObjPrintf("Unknown or invalid search operator (%d).", op));
        else
            Tcl_SetObjResult(interp, Tcl_ObjPrintf("Unknown or invalid search operator (%s).", ops));
    }
    return TCL_ERROR;
}


/* Increments the ref counts of Tcl_Objs in a tarray making sure not
   to run past end of array */
void TArrayIncrObjRefs(TArrayHdr *thdrP, int first, int count)
{
    register int i;
    register Tcl_Obj **objPP;

    if (thdrP->type == TARRAY_OBJ) {
        if ((first + count) > thdrP->used)
            count = thdrP->used - first;
        if (count <= 0)
            return;
        objPP = TAHDRELEMPTR(thdrP, Tcl_Obj *, first);
        for (i = 0; i < count; ++i, ++objPP) {
            Tcl_IncrRefCount(*objPP);
        }
    }
}

/* Decrements the ref counts of Tcl_Objs in a tarray.
   Does NOT CLEAR ANY OTHER HEADER FIELDS. CALLER MUST DO THAT 
*/
void TArrayDecrObjRefs(TArrayHdr *thdrP, int first, int count)
{
    register int i;
    register Tcl_Obj **objPP;

    if (thdrP->type == TARRAY_OBJ) {
        if ((first + count) > thdrP->used)
            count = thdrP->used - first;
        if (count <= 0)
            return;
        objPP = TAHDRELEMPTR(thdrP, Tcl_Obj *, first);
        for (i = 0; i < count; ++i, ++objPP) {
            Tcl_DecrRefCount(*objPP);
        }
    }
}

void TArrayFreeHdr(TArrayHdr *thdrP)
{
    if (--thdrP->nrefs <= 0) {
        if (thdrP->type == TARRAY_OBJ) {
            TArrayDecrObjRefs(thdrP, 0, thdrP->used);
        }
        TARRAY_FREEMEM(thdrP);
    }
}

TCL_RESULT TArrayVerifyType(Tcl_Interp *interp, Tcl_Obj *objP)
{
    if (objP->typePtr == &gTArrayType)
        return TCL_OK;
    else {
        if (interp)
            Tcl_SetResult(interp, "Value is not a tarray", TCL_STATIC);
        return TCL_ERROR;
    }
}

static void TArrayFreeRep(Tcl_Obj *objP)
{
    TArrayHdr *thdrP;

    TARRAY_ASSERT(objP->typePtr == &gTArrayType);

    thdrP = TARRAYHDR(objP); 
    TARRAY_ASSERT(thdrP);

    TArrayFreeHdr(thdrP);
    TARRAYHDR(objP) = NULL;
    objP->typePtr = NULL;
}

static void TArrayDupObj(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    TARRAY_ASSERT(srcP->typePtr == &gTArrayType);
    TARRAY_ASSERT(TARRAYDATA(srcP) != NULL);
        
    TARRAY_OBJ_SETREP(dstP, TARRAYHDR(srcP));
}


/* Called to generate a string implementation from an array of Tcl_Obj */
static void UpdateObjArrayString(
    Tcl_Obj *objP,
    Tcl_Obj **objv,             /* Must NOT be NULL */
    int objc                    /* Must NOT be 0 */
    )
{
    /* Copied almost verbatim from the Tcl's UpdateStringOfList */
#   define LOCAL_SIZE 20
    int localFlags[LOCAL_SIZE], *flagPtr = NULL;
    int i, length, bytesNeeded = 0;
    const char *elem;
    char *dst;

    TARRAY_ASSERT(objv);
    TARRAY_ASSERT(objc > 0);

    /*
     * Pass 1: estimate space, gather flags.
     */

    if (objc <= LOCAL_SIZE) {
        flagPtr = localFlags;
    } else {
        /*
         * We know objc <= TARRAY_MAX_OBJC, so this is safe.
         */

        flagPtr = (int *) ckalloc(objc * sizeof(int));
    }
    for (i = 0; i < objc; i++) {
        flagPtr[i] = (i ? TCL_DONT_QUOTE_HASH : 0);
        elem = Tcl_GetStringFromObj(objv[i], &length);
        bytesNeeded += Tcl_ScanCountedElement(elem, length, flagPtr+i);
        if (bytesNeeded < 0) {
            Tcl_Panic("max size for a Tcl value (%d bytes) exceeded", INT_MAX);
        }
    }
    if (bytesNeeded > INT_MAX - objc + 1) {
        Tcl_Panic("max size for a Tcl value (%d bytes) exceeded", INT_MAX);
    }
    bytesNeeded += objc;

    /*
     * Pass 2: copy into string rep buffer.
     */

    objP->length = bytesNeeded - 1;
    objP->bytes = ckalloc(bytesNeeded);
    dst = objP->bytes;
    for (i = 0; i < objc; i++) {
        flagPtr[i] |= (i ? TCL_DONT_QUOTE_HASH : 0);
        elem = Tcl_GetStringFromObj(objv[i], &length);
        dst += Tcl_ConvertCountedElement(elem, length, dst, flagPtr[i]);
        *dst++ = ' ';
    }
    objP->bytes[objP->length] = '\0';

    if (flagPtr != localFlags) {
        ckfree((char *) flagPtr);
    }
}


static void TArrayUpdateStringRep(Tcl_Obj *objP)
{
    int i, n, count;
    int allocated, unused;
    char *cP;
    int max_elem_space;  /* Max space to print one element including
                            either terminating null or space */
    TArrayHdr *thdrP;
        
    TARRAY_ASSERT(objP->typePtr == &gTArrayType);

    objP->bytes = NULL;

    count = TARRAYELEMCOUNT(objP);
    if (count == 0) {
        objP->bytes = ckalloc(sizeof(objP->bytes[0]));
        objP->bytes[0] = 0;
        objP->length = 0;
        return;
    }

    thdrP = TARRAYHDR(objP);

    /* Code below based on count > 0 else terminating \0 will blow memory */

    /*
     * Special case Boolean since we know exactly how many chars will
     * be required 
     */

    /*
     * When output size cannot be calculated exactly, we allocate using
     * some estimate based on the type.
     */
        
    switch (TARRAYTYPE(objP)) {
    case TARRAY_BOOLEAN:
        {
            unsigned char *ucP = TAHDRELEMPTR(thdrP, unsigned char, 0);
            register unsigned char uc = *ucP;
            register unsigned char uc_mask;

            /* For BOOLEANS, we know how long a buffer needs to be */
            cP = ckalloc(2*count);
            objP->bytes = cP;
            n = count / CHAR_BIT;
            for (i = 0; i < n; ++i, ++ucP) {
                for (uc_mask = BITMASK(0); uc_mask ; uc_mask >>= 1) {
                    *cP++ = (uc & uc_mask) ? '1' : '0';
                    *cP++ = ' ';
                }
            }
            n = count - n*CHAR_BIT;    /* Left over bits in last byte */
            if (n) {
                uc = *ucP;
                for (i = 0, uc_mask = BITMASK(0); i < n; ++i, uc_mask >>= 1) {
                    *cP++ = (uc & uc_mask) ? '1' : '0';
                    *cP++ = ' ';
                }
            }
            cP[-1] = 0;         /* Overwrite last space with terminating \0 */
            objP->length = 2*count - 1;
        }
        return;
                
    case TARRAY_OBJ:
        UpdateObjArrayString(objP, TAHDRELEMPTR(thdrP, Tcl_Obj *, 0), count);
        return;
                
    case TARRAY_UINT:
    case TARRAY_INT:
        TARRAY_ASSERT(sizeof(int) == 4); /* So max string space needed is 11 */
        max_elem_space = 11+1;
        break;
    case TARRAY_WIDE:
        max_elem_space = TCL_INTEGER_SPACE+1;
        break;
    case TARRAY_DOUBLE:
        max_elem_space = TCL_DOUBLE_SPACE+1;
        break;
    case TARRAY_BYTE:
        max_elem_space = 3+1;
        break;
    default:
        TArrayTypePanic(thdrP->type);
    }
            
    allocated = 0;
    unused = 0;
    objP->bytes= NULL;
    /* TBD - do Nested loop for efficiency reasons to avoid switch on every iter */
    for (i = 0; i < count; ++i) {
        if (unused < max_elem_space) {
            n = allocated - unused; /* Used space */
            /* Increase assuming remaining take half max space on average */
            allocated += ((max_elem_space + 1)/2)*(count - i);
            objP->bytes = ckrealloc(objP->bytes, allocated);
            cP = n + (char *) objP->bytes;
            unused = allocated - n;
        }
        switch (thdrP->type) {
        case TARRAY_UINT:
            _snprintf(cP, unused, "%u", *TAHDRELEMPTR(thdrP, unsigned int, i));
            break;
        case TARRAY_INT:
            _snprintf(cP, unused, "%d", *TAHDRELEMPTR(thdrP, int, i));
            break;
        case TARRAY_WIDE:
            _snprintf(cP, unused, "%" TCL_LL_MODIFIER "d", *TAHDRELEMPTR(thdrP, Tcl_WideInt, i));
            break;
        case TARRAY_DOUBLE:
            /* Do not use _snprintf because of slight difference
               it does not include decimal point for whole ints. For
               consistency with Tcl, use Tcl_PrintDouble instead */
            Tcl_PrintDouble(NULL, *TAHDRELEMPTR(thdrP, double, i), cP);
            break;
        case TARRAY_BYTE:
            _snprintf(cP, unused, "%u", *TAHDRELEMPTR(thdrP, unsigned char, i));
            break;
        }
        n = strlen(cP);
        cP += n;
        *cP++ = ' ';
        unused -= n+1;
    }

    cP[-1] = 0;         /* Overwrite last space with terminating \0 */
    objP->length = allocated - unused - 1; /* Terminating null not included in length */
            
    /* Only shrink array if unused space is comparatively too large */
    if (unused > (allocated / 8) && unused > 10)
        objP->bytes = ckrealloc(objP->bytes, allocated - unused);
    return;
}

Tcl_Obj *TArrayNewObj(TArrayHdr *thdrP)
{
    Tcl_Obj *objP = Tcl_NewObj();
    Tcl_InvalidateStringRep(objP);
    TARRAY_OBJ_SETREP(objP, thdrP);
    return objP;
}
    
/* thdrP must NOT be shared and must have enough slots */
/* interp may be NULL (only used for errors) */
TCL_RESULT TArraySetFromObjs(Tcl_Interp *interp, TArrayHdr *thdrP,
                                 int first, int nelems,
                                 Tcl_Obj * const elems[])
{
    int i, ival;
    Tcl_WideInt wide;
    double dval;

    TARRAY_ASSERT(thdrP->nrefs < 2);

    if ((first + nelems) > thdrP->allocated) {
        /* Should really panic but not a fatal error (ie. no memory
         * corruption etc.). Most likely some code path did not check
         * size and allocate space accordingly.
         */
        if (interp)
            Tcl_SetResult(interp, "Internal error: TArray too small.", TCL_STATIC);
        return TCL_ERROR;
    }

    /*
     * In case of conversion errors, we have to keep the old values
     * so we loop through first to verify there are no errors and then
     * a second time to actually store the values. The arrays can be
     * very large so we do not want to allocate a temporary
     * holding area for saving old values to be restored in case of errors.
     *
     * As a special optimization, when appending to the end, we do
     * not need to first check. We directly store the values and in case
     * of errors, simply restore the old size.
     *
     * Also for TARRAY_OBJ there is no question of conversion and hence
     * no question of conversion errors.
     */

    if (first < thdrP->used && thdrP->type != TARRAY_OBJ) {
        /* Not appending, need to verify conversion */
        switch (thdrP->type) {
        case TARRAY_BOOLEAN:
            for (i = 0; i < nelems; ++i) {
                if (Tcl_GetBooleanFromObj(interp, elems[i], &ival) != TCL_OK)
                    goto convert_error;
            }
            break;

        case TARRAY_UINT:
            for (i = 0; i < nelems; ++i) {
                if (Tcl_GetWideIntFromObj(interp, elems[i], &wide) != TCL_OK)
                    goto convert_error;
                if (wide < 0 || wide > 0xFFFFFFFF) {
                    if (interp)
                        Tcl_SetObjResult(interp,
                                         Tcl_ObjPrintf("Integer \"%s\" too large for type \"uint\" typearray.", Tcl_GetString(elems[i])));
                    goto convert_error;
                }
            }
            break;

        case TARRAY_INT:
            for (i = 0; i < nelems; ++i) {
                if (Tcl_GetIntFromObj(interp, elems[i], &ival) != TCL_OK)
                    goto convert_error;
            }
            break;

        case TARRAY_WIDE:
            for (i = 0; i < nelems; ++i) {
                if (Tcl_GetWideIntFromObj(interp, elems[i], &wide) != TCL_OK)
                    goto convert_error;
            }
            break;

        case TARRAY_DOUBLE:
            for (i = 0; i < nelems; ++i) {
                if (Tcl_GetDoubleFromObj(interp, elems[i], &dval) != TCL_OK)
                    goto convert_error;
            }
            break;

        case TARRAY_BYTE:
            for (i = 0; i < nelems; ++i) {
                if (Tcl_GetIntFromObj(interp, elems[i], &ival) != TCL_OK)
                    goto convert_error;
                if (ival > 255 || ival < 0) {
                    if (interp)
                        Tcl_SetObjResult(interp,
                                         Tcl_ObjPrintf("Integer \"%d\" does not fit type \"byte\" typearray.", ival));
                    goto convert_error;
                }
            }
            break;
        default:
            TArrayTypePanic(thdrP->type);
        }
    }

    /*
     * Now actually store the values. Note we still have to check
     * status on conversion since we did not do checks when we are appending
     * to the end.
     */

    switch (thdrP->type) {
    case TARRAY_BOOLEAN:
        {
            register unsigned char *ucP;
            unsigned int uc, uc_mask;

            /* Take care of the initial condition where the first bit
               may not be aligned on a char boundary */
            ucP = TAHDRELEMPTR(thdrP, unsigned char, first / CHAR_BIT);
            uc = first % CHAR_BIT; /* Offset of bit within a char */
            uc_mask = BITMASK(uc); /* The bit position corresponding to 'first' */
            if (uc != 0) {
                /*
                 * Offset is uc within a char. Get the byte at that location
                 * preserving the preceding bits within the char.
                 */
                uc = *ucP & BITMASKBELOW(uc);
            } else {
                /* uc = 0; Already so */
            }
            for (i = 0; i < nelems; ++i) {
                if (Tcl_GetBooleanFromObj(interp, elems[i], &ival) != TCL_OK)
                    goto convert_error;
                if (ival)
                    uc |= uc_mask;
                uc_mask >>= 1;
                if (uc_mask == 0) {
                    *ucP++ = uc;
                    uc = 0;
                    uc_mask = BITMASK(0);
                }
            }
            if (uc_mask != BITMASK(0)) {
                /* We have some leftover bits in ui that need to be stored.
                 * We need to *merge* these into the corresponding word
                 * keeping the existing high index bits.
                 * Note the bit indicated by ui_mask also has to be preserved,
                 * not overwritten.
                 */
                *ucP = uc | (*ucP & (uc_mask | (uc_mask-1)));
            }
        }
        break;

    case TARRAY_UINT:
        {
            register unsigned int *uintP;
            uintP = TAHDRELEMPTR(thdrP, unsigned int, first);
            for (i = 0; i < nelems; ++i, ++uintP) {
                if (Tcl_GetWideIntFromObj(interp, elems[i], &wide) != TCL_OK)
                    goto convert_error;
                if (wide < 0 || wide > 0xFFFFFFFF) {
                    if (interp)
                        Tcl_SetObjResult(interp,
                                         Tcl_ObjPrintf("Integer \"%s\" too large for type \"uint\" typearray.", Tcl_GetString(elems[i])));
                    goto convert_error;
                }
                *uintP = (unsigned int) wide;
            }
        }
        break;
    case TARRAY_INT:
        {
            register int *intP;
            intP = TAHDRELEMPTR(thdrP, int, first);
            for (i = 0; i < nelems; ++i, ++intP) {
                if (Tcl_GetIntFromObj(interp, elems[i], intP) != TCL_OK)
                    goto convert_error;
            }
        }
        break;

    case TARRAY_WIDE:
        {
            register Tcl_WideInt *wideP;
            wideP = TAHDRELEMPTR(thdrP, Tcl_WideInt, first);
            for (i = 0; i < nelems; ++i, ++wideP) {
                if (Tcl_GetWideIntFromObj(interp, elems[i], wideP) != TCL_OK)
                    goto convert_error;
            }
        }
        break;

    case TARRAY_DOUBLE:
        {
            register double *dblP;
            dblP = TAHDRELEMPTR(thdrP, double, first);
            for (i = 0; i < nelems; ++i, ++dblP) {
                if (Tcl_GetDoubleFromObj(interp, elems[i], dblP) != TCL_OK)
                    goto convert_error;
            }
        }
        break;

    case TARRAY_OBJ:
        {
            register Tcl_Obj **objPP;
            objPP = TAHDRELEMPTR(thdrP, Tcl_Obj *, first);
            for (i = 0; i < nelems; ++i, ++objPP) {
                /* Careful about the order here! */
                Tcl_IncrRefCount(elems[i]);
                if ((first + i) < thdrP->used) {
                    /* Deref what was originally in that slot */
                    Tcl_DecrRefCount(*objPP);
                }
                *objPP = elems[i];
            }
        }
        break;

    case TARRAY_BYTE:
        {
            register unsigned char *byteP;
            byteP = TAHDRELEMPTR(thdrP, unsigned char, first);
            for (i = 0; i < nelems; ++i, ++byteP) {
                if (Tcl_GetIntFromObj(interp, elems[i], &ival) != TCL_OK)
                    goto convert_error;
                if (ival > 255 || ival < 0) {
                    if (interp)
                        Tcl_SetObjResult(interp,
                                         Tcl_ObjPrintf("Integer \"%d\" does not fit type \"byte\" typearray.", ival));
                    goto convert_error;
                }
                *byteP = (unsigned char) ival;
            }
        }
        break;

    default:
        TArrayTypePanic(thdrP->type);
    }

    if ((first + nelems) > thdrP->used)
        thdrP->used = first + nelems;

    return TCL_OK;

convert_error:                  /* Interp should already contain errors */
    TARRAY_ASSERT(thdrP->type != TARRAY_OBJ); /* Else we may need to deal with ref counts */

    return TCL_ERROR;

}

int TArrayCalcSize(unsigned char tatype, int count)
{
    int space;

    switch (tatype) {
    case TARRAY_BOOLEAN:
        space = (count + CHAR_BIT - 1) / CHAR_BIT;
        break;
    case TARRAY_UINT:
    case TARRAY_INT:
        space = count * sizeof(int);
        break;
    case TARRAY_WIDE:
        space = count * sizeof(Tcl_WideInt);
        break;
    case TARRAY_DOUBLE:
        space = count * sizeof(double);
        break;
    case TARRAY_OBJ:
        space = count * sizeof(Tcl_Obj *);
        break;
    case TARRAY_BYTE:
        space = count * sizeof(unsigned char);
        break;
    default:
        TArrayTypePanic(tatype);
    }

    return sizeof(TArrayHdr) + space;
}

TArrayHdr *TArrayRealloc(TArrayHdr *oldP, int new_count)
{
    TArrayHdr *thdrP;

    TARRAY_ASSERT(oldP->nrefs < 2);
    TARRAY_ASSERT(oldP->used <= new_count);

    thdrP = (TArrayHdr *) TARRAY_REALLOCMEM((char *) oldP, TArrayCalcSize(oldP->type, new_count));
    thdrP->allocated = new_count;
    return thdrP;
}

TArrayHdr * TArrayAlloc(unsigned char tatype, int count)
{
    unsigned char nbits;
    TArrayHdr *thdrP;

    if (count == 0)
            count = TARRAY_DEFAULT_NSLOTS;
    thdrP = (TArrayHdr *) TARRAY_ALLOCMEM(TArrayCalcSize(tatype, count));
    thdrP->nrefs = 0;
    thdrP->allocated = count;
    thdrP->used = 0;
    thdrP->type = tatype;
    switch (tatype) {
    case TARRAY_BOOLEAN: nbits = 1; break;
    case TARRAY_UINT: nbits = sizeof(unsigned int) * CHAR_BIT; break;
    case TARRAY_INT: nbits = sizeof(int) * CHAR_BIT; break;
    case TARRAY_WIDE: nbits = sizeof(Tcl_WideInt) * CHAR_BIT; break;
    case TARRAY_DOUBLE: nbits = sizeof(double) * CHAR_BIT; break;
    case TARRAY_OBJ: nbits = sizeof(Tcl_Obj *) * CHAR_BIT; break;
    case TARRAY_BYTE: nbits = sizeof(unsigned char) * CHAR_BIT; break;
    default:
        TArrayTypePanic(tatype);
    }
    thdrP->elem_bits = nbits;
    
    return thdrP;
}

TArrayHdr * TArrayAllocAndInit(Tcl_Interp *interp, unsigned char tatype,
                           int nelems, Tcl_Obj * const elems[],
                           int init_size)
{
    TArrayHdr *thdrP;

    if (elems) {
        /*
         * Initialization provided. If explicit size specified, fix
         * at that else leave some extra space.
         */
        if (init_size) {
            if (init_size < nelems)
                init_size = nelems;
        } else {
            init_size = nelems + TARRAY_EXTRA(nelems);
        }
    } else {
        nelems = 0;
    }

    thdrP = TArrayAlloc(tatype, init_size);

    if (elems != NULL && nelems != 0) {
        if (TArraySetFromObjs(interp, thdrP, 0, nelems, elems) != TCL_OK) {
            TARRAY_FREEMEM(thdrP);
            return NULL;
        }
    }

    return thdrP;

}

/* dstP must not be shared and must be large enough */
TCL_RESULT TArraySet(Tcl_Interp *interp, TArrayHdr *dstP, int dst_first,
                             TArrayHdr *srcP, int src_first, int count)
{
    int nbytes;
    void *s, *d;

    TARRAY_ASSERT(dstP->type == srcP->type);
    TARRAY_ASSERT(dstP->nrefs < 2); /* Must not be shared */

    if (src_first < 0)
        src_first = 0;

    if (src_first >= srcP->used || dstP == srcP)
        return TCL_OK;          /* Nothing to be copied */

    if ((src_first + count) > srcP->used)
        count = srcP->used - src_first;

    if (count <= 0)
        return TCL_OK;

    if (dst_first < 0)
        dst_first = 0;
    else if (dst_first > dstP->used)
        dst_first = dstP->used;

    if ((dst_first + count) > dstP->allocated) {
        if (interp)
            Tcl_SetResult(interp, "Internal error: TArray too small.", TCL_STATIC);
        return TCL_ERROR;
    }

    /*
     * For all types other than BOOLEAN and OBJ, we can just memcpy
     * Those two types have complication in that BOOLEANs are compacted
     * into bytes and the copy may not be aligned on a byte boundary.
     * For OBJ types, we have to deal with reference counts.
     */
    switch (srcP->type) {
    case TARRAY_BOOLEAN:
        bitarray_copy(TAHDRELEMPTR(srcP, unsigned char, 0),
                      src_first, count,
                      TAHDRELEMPTR(dstP, unsigned char, 0),
                      dst_first);
        if ((dst_first + count) > dstP->used)
            dstP->used = dst_first + count;
        return TCL_OK;

    case TARRAY_OBJ:
        /*
         * We have to deal with reference counts here. For the objects
         * we are copying (source) we need to increment the reference counts.
         * For objects in destination that we are overwriting, we need
         * to decrement reference counts.
         */

        TArrayIncrObjRefs(srcP, src_first, count); /* Do this first */
        /* Note this call take care of the case where count exceeds
         * actual number in dstP
         */
        TArrayDecrObjRefs(dstP, dst_first, count);
         
        /* Now we can just memcpy like the other types */
        nbytes = count * sizeof(Tcl_Obj *);
        s = TAHDRELEMPTR(srcP, Tcl_Obj *, src_first);
        d = TAHDRELEMPTR(dstP, Tcl_Obj *, dst_first);
        break;

    case TARRAY_UINT:
    case TARRAY_INT:
        nbytes = count * sizeof(int);
        s = TAHDRELEMPTR(srcP, int, src_first);
        d = TAHDRELEMPTR(dstP, int, dst_first);
        break;
    case TARRAY_WIDE:
        nbytes = count * sizeof(Tcl_WideInt);
        s = TAHDRELEMPTR(srcP, Tcl_WideInt, src_first);
        d = TAHDRELEMPTR(dstP, Tcl_WideInt, dst_first);
        break;
    case TARRAY_DOUBLE:
        nbytes = count * sizeof(double);
        s = TAHDRELEMPTR(srcP, double, src_first);
        d = TAHDRELEMPTR(dstP, double, dst_first);
        break;
    case TARRAY_BYTE:
        nbytes = count * sizeof(unsigned char);
        s = TAHDRELEMPTR(srcP, unsigned char, src_first);
        d = TAHDRELEMPTR(dstP, unsigned char, dst_first);
        break;
    default:
        TArrayTypePanic(srcP->type);
    }

    memcpy(d, s, nbytes);

    if ((dst_first + count) > dstP->used)
        dstP->used = dst_first + count;

    return TCL_OK;
}

/* Note: nrefs of cloned array is 0 */
TArrayHdr *TArrayClone(TArrayHdr *srcP, int minsize)
{
    TArrayHdr *thdrP;

    if (minsize == 0)
        minsize = srcP->allocated;
    else if (minsize < srcP->used)
        minsize = srcP->used;

    /* TBD - optimize these two calls */
    thdrP = TArrayAlloc(srcP->type, minsize);
    if (TArraySet(NULL, thdrP, 0, srcP, 0, srcP->used) != TCL_OK) {
        TArrayFreeHdr(thdrP);
        return NULL;
    }
    return thdrP;
}

void bitarray_set(unsigned char *ucP, int offset, int count, int ival)
{
    unsigned uc;
    unsigned char uc_mask; /* The bit position corresponding to 'first' */
    int i;

    if (count == 0)
        return;

    /* First set the bits to get to a char boundary */
    ucP += offset/CHAR_BIT;
    uc = offset % CHAR_BIT; /* Offset of bit within a char */
    i = 0;                     /* Will track num partial bits copied */
    if (uc != 0) {
        /* Offset is uc within a char. Get the byte at that location */
        uc_mask = BITMASK(uc);
        uc = *ucP;
        for (; i < count && uc_mask; ++i, uc_mask >>= 1) {
            if (ival)
                uc |= uc_mask;
            else
                uc &= ~ uc_mask;
        }
        *ucP++ = uc;
    }
    /* Copied the first i bits. Now copy full bytes with memset */
    memset(ucP, ival ? 0xff : 0, (count-i)/CHAR_BIT);
    ucP += (count-i)/CHAR_BIT;
    i = i + 8*(count-i); /* Number of bits left to be copied */
    TARRAY_ASSERT(i < 8);
    /* Now leftover bits */
    if (i) {
        uc = *ucP;
        for (uc_mask = BITMASK(0); i; --i, uc_mask >>= 1) {
            if (ival)
                uc |= uc_mask;
            else
                uc &= ~ uc_mask;
        }
        *ucP = uc;
    }
}


/* dstP must not be shared and must be large enough */
TCL_RESULT TArraySetRange(Tcl_Interp *interp, TArrayHdr *dstP, int dst_first,
                          int count, Tcl_Obj *objP)
{
    int i, n, ival;
    unsigned char *ucP;

    /* TBD - optimize when value is 0 by using memset */

    TARRAY_ASSERT(dstP->nrefs < 2); /* Must not be shared */

    if (count <= 0)
        return TCL_OK;

    if (dst_first < 0)
        dst_first = 0;
    else if (dst_first > dstP->used)
        dst_first = dstP->used;

    if ((dst_first + count) > dstP->allocated) {
        if (interp)
            Tcl_SetResult(interp, "Internal error: TArray too small.", TCL_STATIC);
        return TCL_ERROR;
    }

    switch (dstP->type) {
    case TARRAY_BOOLEAN:
        if (Tcl_GetBooleanFromObj(interp, objP, &ival) != TCL_OK)
            return TCL_ERROR;
        else {
            bitarray_set(TAHDRELEMPTR(dstP, unsigned char, 0), dst_first,
                         count, ival);
        }
        break;

    case TARRAY_OBJ:
        {
            Tcl_Obj **objPP;

            /*
             * We have to deal with reference counts here. For the object
             * we are copying we need to increment the reference counts
             * that many times. For objects being overwritten,
             * we need to decrement reference counts. Note we c
             */
            /* First loop overwriting existing elements */
            n = dst_first + count;
            if (n > dstP->used)
                n = dstP->used;
            objPP = TAHDRELEMPTR(dstP, Tcl_Obj *, dst_first);
            for (i = dst_first; i < n; ++i) {
                /* Be careful of the order */
                Tcl_IncrRefCount(objP);
                Tcl_DecrRefCount(*objPP);
                *objPP = objP;
            }

            /* Now loop over new elements being appended */
            for (; i < dst_first+count; ++i) {
                Tcl_IncrRefCount(objP);
                *objPP = objP;
            }
        }
        break;

    case TARRAY_UINT:           /* TBD - test specifically for UINT_MAX etc. */
    case TARRAY_INT:
        if (Tcl_GetIntFromObj(interp, objP, &ival) != TCL_OK)
            return TCL_ERROR;
        else {
            int *iP;
            iP = TAHDRELEMPTR(dstP, int, dst_first);
            for (i = 0; i < count; ++i, ++iP)
                *iP = ival;
        }
        break;
    case TARRAY_WIDE:
        {
            Tcl_WideInt wide, *wideP;

            if (Tcl_GetWideIntFromObj(interp, objP, &wide) != TCL_OK)
                return TCL_ERROR;

            wideP = TAHDRELEMPTR(dstP, Tcl_WideInt, dst_first);
            for (i = 0; i < count; ++i, ++wideP)
                *wideP = wide;
        }
        break;

    case TARRAY_DOUBLE:
        {
            double dval, *dvalP;
            if (Tcl_GetDoubleFromObj(interp, objP, &dval) != TCL_OK)
                return TCL_ERROR;

            dvalP = TAHDRELEMPTR(dstP, double, dst_first);
            for (i = 0; i < count; ++i, ++dvalP)
                *dvalP = dval;
        }
        break;

    case TARRAY_BYTE:
        if (Tcl_GetIntFromObj(interp, objP, &ival) != TCL_OK)
            return TCL_ERROR;
        else {
            if (ival > 255 || ival < 0) {
                if (interp)
                    Tcl_SetObjResult(interp,
                                     Tcl_ObjPrintf("Integer \"%d\" does not fit type \"byte\" typearray.", ival));
                return TCL_ERROR;
            }
            ucP = TAHDRELEMPTR(dstP, unsigned char, dst_first);
            for (i = 0; i < count; ++i, ++ucP)
                *ucP = (unsigned char) ival;
        }
        break;

    default:
        TArrayTypePanic(dstP->type);
    }

    if ((dst_first + count) > dstP->used)
        dstP->used = dst_first + count;

    return TCL_OK;
}

/* Returns a Tcl_Obj for a TArray slot. NOTE: WITHOUT its ref count incremented */
Tcl_Obj * TArrayIndex(Tcl_Interp *interp, TArrayHdr *thdrP, int index)
{
    int offset;

    if (index >= thdrP->used) {
        if (interp)
            Tcl_SetResult(interp, "tarray index out of bounds", TCL_STATIC);
        return NULL;
    }

    switch (thdrP->type) {
    case TARRAY_BOOLEAN:
        offset = index / CHAR_BIT;
        index = index % CHAR_BIT;
        return Tcl_NewIntObj(0 != (BITMASK(index) & *TAHDRELEMPTR(thdrP, unsigned char, offset)));
    case TARRAY_UINT:
        return Tcl_NewWideIntObj(*TAHDRELEMPTR(thdrP, unsigned int, index));
    case TARRAY_INT:
        return Tcl_NewIntObj(*TAHDRELEMPTR(thdrP, int, index));
    case TARRAY_WIDE:
        return Tcl_NewWideIntObj(*TAHDRELEMPTR(thdrP, Tcl_WideInt, index));
    case TARRAY_DOUBLE:
        return Tcl_NewDoubleObj(*TAHDRELEMPTR(thdrP, double, index));
    case TARRAY_BYTE:
        return Tcl_NewIntObj(*TAHDRELEMPTR(thdrP, unsigned char, index));
    case TARRAY_OBJ:
        return *TAHDRELEMPTR(thdrP, Tcl_Obj *, index);
    }
}

/* Returns a TArrayHdr of type int. The header's ref count is incremented
 * so caller should call TArrayFreeHdr as appropriate
 */
TArrayHdr *TArrayConvertToIndices(Tcl_Interp *interp, Tcl_Obj *objP)
{
    TArrayHdr *thdrP;
    Tcl_Obj **elems;
    int       nelems;

    /* Indices should be a tarray of ints. If not, treat as a list
     * and convert it that way. Though that is slower, it should be rare
     * as all tarray indices are returned already in the proper format.
     */
    if (objP->typePtr == &gTArrayType && TARRAYTYPE(objP) == TARRAY_INT) {
        thdrP = TARRAYHDR(objP);
        thdrP->nrefs++;
        return thdrP;
    }

    if (Tcl_ListObjGetElements(interp, objP, &nelems, &elems) != TCL_OK)
        return NULL;

    thdrP = TArrayAllocAndInit(interp, TARRAY_INT, nelems, elems, 0);
    if (thdrP)
        thdrP->nrefs++;
    return thdrP;
}

/* Returns a newly allocated TArrayHdr (with ref count 0) containing the
   values from the specified indices */
TArrayHdr *TArrayGetValues(Tcl_Interp *interp, TArrayHdr *srcP, TArrayHdr *indicesP)
{
    TArrayHdr *thdrP;
    int i, count;
    int *indexP;

    if (indicesP->type != TARRAY_INT) {
        if (interp)
            Tcl_SetResult(interp, "Invalid type for tarray indices", TCL_STATIC);
        return NULL;
    }

    count = indicesP->used;
    thdrP = TArrayAlloc(srcP->type, count);
    if (thdrP == 0 || count == 0)
        return thdrP;

    indexP = TAHDRELEMPTR(indicesP, int, 0);

    switch (srcP->type) {
    case TARRAY_BOOLEAN:
        {
            unsigned char *ucP = TAHDRELEMPTR(thdrP, unsigned char, 0);
            int off;
            unsigned char uc, src_uc, uc_mask;
            for (i = 0, uc = 0, uc_mask = BITMASK(0); i < count; ++i, ++indexP) {
                off = *indexP / CHAR_BIT;/* Offset to char containing the bit */
                src_uc = *TAHDRELEMPTR(srcP, unsigned char, off); /* The char */
                if (src_uc & BITMASK(*indexP/CHAR_BIT))
                    uc |= uc_mask;
                uc_mask >>= 1;
                if (uc_mask == 0) {
                    *ucP++ = uc;
                    uc = 0;
                    uc_mask = BITMASK(0);
                }
            }
            if (uc_mask != BITMASK(0)) {
                /* We have some leftover bits in ui that need to be stored. */
                *ucP = uc;
            }
        }
        break;
    case TARRAY_UINT:
    case TARRAY_INT:
        {
            unsigned int *uiP = TAHDRELEMPTR(thdrP, unsigned int, 0);
            for (i = 0; i < count; ++i, ++indexP, ++uiP)
                *uiP = *TAHDRELEMPTR(srcP, unsigned int, *indexP);
        }
        break;
    case TARRAY_WIDE:
        {
            Tcl_WideInt *wideP = TAHDRELEMPTR(thdrP, Tcl_WideInt, 0);
            for (i = 0; i < count; ++i, ++indexP, ++wideP)
                *wideP = *TAHDRELEMPTR(srcP, Tcl_WideInt, *indexP);
        }
        break;
    case TARRAY_DOUBLE:
        {
            double *dblP = TAHDRELEMPTR(thdrP, double, 0);
            for (i = 0; i < count; ++i, ++indexP, ++dblP)
                *dblP = *TAHDRELEMPTR(srcP, double, *indexP);
        }
        break;
    case TARRAY_BYTE:
        {
            unsigned char *ucP = TAHDRELEMPTR(thdrP, unsigned char, 0);
            for (i = 0; i < count; ++i, ++indexP, ++ucP)
                *ucP = *TAHDRELEMPTR(srcP, unsigned char, *indexP);
        }
        break;
    case TARRAY_OBJ:
        {
            Tcl_Obj **objPP = TAHDRELEMPTR(thdrP, Tcl_Obj *, 0);
            for (i = 0; i < count; ++i, ++indexP, ++objPP) {
                *objPP = *TAHDRELEMPTR(srcP, Tcl_Obj *, *indexP);
                Tcl_IncrRefCount(*objPP);
            }
        }
        break;
    default:
        TArrayTypePanic(srcP->type);
    }

    thdrP->used = count;
    return thdrP;
}


static TCL_RESULT TArraySearchBoolean(Tcl_Interp *interp, TArrayHdr * haystackP,
                                      Tcl_Obj *needleObj, int start, enum TArraySearchSwitches op, int flags)
{
    int bval;
    unsigned char *ucP;
    unsigned char uc, uc_mask, skip;
    int offset;
    Tcl_Obj *resultObj;

    TARRAY_ASSERT(haystackP->type == TARRAY_BOOLEAN);

    if (Tcl_GetBooleanFromObj(interp, needleObj, &bval) != TCL_OK)
        return TCL_ERROR;
    
    if (op != TARRAY_SEARCH_OPT_EQ)
        return TArrayBadSearchOpError(interp, op);

    if (flags & TARRAY_SEARCH_INVERT)
        bval = !bval;

    skip = bval ? 0 : 0xff;     /* Skip entire bytes of this value */

    /* First locate the starting point for the search */
    ucP = TAHDRELEMPTR(haystackP, unsigned char, 0);
    ucP += start/CHAR_BIT;
    uc_mask = BITMASK(start % CHAR_BIT);

    /* TBD - optimize this code */

    /*
     * At this point,
     * ucP points to the memory location containing bit at offset start
     * uc_mask is the mask for the start bit within that location
     */
    if (flags & TARRAY_SEARCH_ALL) {
        TArrayHdr *thdrP;

        thdrP = TArrayAlloc(
            flags & TARRAY_SEARCH_INLINE ? TARRAY_BOOLEAN : TARRAY_INT,
            10);                /* Assume 10 hits */

        for (offset = start; offset < haystackP->used; uc_mask = BITMASK(0)) {
            /*
             * At top of loop, *ucP potentially has a matching
             * bit, uc_mask contains position at which to
             * begin match
             */
            uc = *ucP++;
            if (uc == skip) {
                /* Looking for 1's and uc is all 0's or vice versa */
                offset += CHAR_BIT;
                continue;
            }
            while (uc_mask) {
                /* Compare bit against 1 or 0 as appropriate */
                if ((bval && (uc_mask & uc)) ||
                    !(bval || (uc_mask & uc))) {
                    /* Note this may be beyond haystackP->used, so check */
                    if (offset >= haystackP->used) {
                        /* Yep, beyond end */
                        break; /* Stop inner loop */
                    }
                    /* Ensure enough space in target array */
                    if (thdrP->used >= thdrP->allocated)
                        thdrP = TArrayRealloc(thdrP, thdrP->used + TARRAY_EXTRA(thdrP->used));
                    if (flags & TARRAY_SEARCH_INLINE) {
                        unsigned char *uc2P = TAHDRELEMPTR(thdrP, unsigned char, thdrP->used/CHAR_BIT);
                        unsigned char uc = BITMASK(thdrP->used % CHAR_BIT);
                        if (bval)
                            *uc2P |= uc;
                        else
                            *uc2P &= ~uc;
                    } else {
                        *TAHDRELEMPTR(thdrP, int, thdrP->used) = offset;
                    }
                    thdrP->used++;
                }
                uc_mask >>= 1;
                ++offset;
            }
        }

        resultObj = TArrayNewObj(thdrP);

    } else {
        /* Return first found element */
        int pos = -1;

        for (offset = start; offset < haystackP->used; uc_mask = BITMASK(0)) {
            /*
             * At top of loop, *ucP potentially has a matching
             * bit, uc_mask contains position at which to
             * begin match
             */
            uc = *ucP++;
            if (uc == skip) {
                /* Looking for 1's and uc is all 0's or vice versa */
                offset += CHAR_BIT;
                continue;
            }
            while (uc_mask) {
                /* Compare bit against 1 or 0 as appropriate */
                if ((bval && (uc_mask & uc)) ||
                    !(bval || (uc_mask & uc))) {
                    /* Note this may be beyond haystackP->used, but
                     * that's ok, we'll reset later rather than
                     * add another test to this loop
                     */
                    pos = offset;
                    break;
                }
                uc_mask >>= 1;
                ++offset;
            }
            if (pos >= 0)
                break;
        }

        if (pos >= haystackP->used)
            pos = -1;       /* We matched on extraneous bits in last byte */
        if (pos == -1)
            resultObj = Tcl_NewObj();
        else {
            resultObj =
                Tcl_NewIntObj((flags & TARRAY_SEARCH_INLINE) ? bval : pos);
        }
    }

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}
                        
/* TBD - see how much performance is gained by separating this search function into
   type-specific functions */
static TCL_RESULT TArraySearchEntier(Tcl_Interp *interp, TArrayHdr * haystackP,
                                   Tcl_Obj *needleObj, int start, enum TArraySearchSwitches op, int flags)
{
    int offset;
    Tcl_Obj *resultObj;
    Tcl_WideInt needle, elem, min_val, max_val;
    int compare_result;
    int compare_wanted;
    int elem_size;
    char *p;

    switch (op) {
    case TARRAY_SEARCH_OPT_GT:
    case TARRAY_SEARCH_OPT_LT: 
    case TARRAY_SEARCH_OPT_EQ:
        break;
    default:
        return TArrayBadSearchOpError(interp, op);
    }

    p = TAHDRELEMPTR(haystackP, char, 0);
    switch (haystackP->type) {
    case TARRAY_INT:
        max_val = INT_MAX;
        min_val = INT_MIN;
        p += start * sizeof(int);
        elem_size = sizeof(int);
        break;
    case TARRAY_UINT:
        max_val = UINT_MAX;
        min_val = 0;
        p += start * sizeof(unsigned int);
        elem_size = sizeof(unsigned int);
        break;
    case TARRAY_WIDE:
        max_val = needle; /* No-op */
        min_val = needle;
        p += start * sizeof(Tcl_WideInt);
        elem_size = sizeof(Tcl_WideInt);
        break;
    case TARRAY_BYTE:
        max_val = UCHAR_MAX;
        min_val = 0;
        p += start * sizeof(unsigned char);
        elem_size = sizeof(unsigned char);
        break;
    default:
        TArrayTypePanic(haystackP->type);
    }

    if (Tcl_GetWideIntFromObj(interp, needleObj, &needle) != TCL_OK)
        return TCL_ERROR;
    else {
        if (needle > max_val || needle < min_val) {
            Tcl_SetObjResult(interp,
                             Tcl_ObjPrintf("Integer \"%s\" type mismatch for typearray (type %d).", Tcl_GetString(needleObj), haystackP->type));
            return TCL_ERROR;
        }
    }

    compare_wanted = flags & TARRAY_SEARCH_INVERT ? 0 : 1;

    if (flags & TARRAY_SEARCH_ALL) {
        TArrayHdr *thdrP;

        thdrP = TArrayAlloc(
            flags & TARRAY_SEARCH_INLINE ? haystackP->type : TARRAY_INT,
            10);                /* Assume 10 hits TBD */

        for (offset = start; offset < haystackP->used; ++offset, p += elem_size) {
            switch (haystackP->type) {
            case TARRAY_INT:  elem = *(int *)p; break;
            case TARRAY_UINT: elem = *(unsigned int *)p; break;
            case TARRAY_WIDE: elem = *(Tcl_WideInt *)p; break;
            case TARRAY_BYTE: elem = *(unsigned char *)p; break;
            }
            switch (op) {
            case TARRAY_SEARCH_OPT_GT: compare_result = (elem > needle); break;
            case TARRAY_SEARCH_OPT_LT: compare_result = (elem < needle); break;
            case TARRAY_SEARCH_OPT_EQ:
            default: compare_result = (elem == needle); break;
            }

            if (compare_result == compare_wanted) {
                /* Have a match */
                /* Ensure enough space in target array */
                if (thdrP->used >= thdrP->allocated)
                    thdrP = TArrayRealloc(thdrP, thdrP->used + TARRAY_EXTRA(thdrP->used));
                if (flags & TARRAY_SEARCH_INLINE) {
                    switch (thdrP->type) {
                    case TARRAY_INT:  *TAHDRELEMPTR(thdrP, int, thdrP->used) = (int) elem; break;
                    case TARRAY_UINT: *TAHDRELEMPTR(thdrP, unsigned int, thdrP->used) = (unsigned int) elem; break;
                    case TARRAY_WIDE: *TAHDRELEMPTR(thdrP, Tcl_WideInt, thdrP->used) = elem; break;
                    case TARRAY_BYTE:  *TAHDRELEMPTR(thdrP, unsigned char, thdrP->used) = (unsigned char) elem; break;
                    }
                } else {
                    *TAHDRELEMPTR(thdrP, int, thdrP->used) = offset;
                }
                thdrP->used++;
            }
        }

        resultObj = TArrayNewObj(thdrP);

    } else {
        /* Return first found element */
        for (offset = start; offset < haystackP->used; ++offset, p += elem_size) {
            switch (haystackP->type) {
            case TARRAY_INT:  elem = *(int *)p; break;
            case TARRAY_UINT: elem = *(unsigned int *)p; break;
            case TARRAY_WIDE: elem = *(Tcl_WideInt *)p; break;
            case TARRAY_BYTE: elem = *(unsigned char *)p; break;
            }
            switch (op) {
            case TARRAY_SEARCH_OPT_GT: compare_result = (elem > needle); break;
            case TARRAY_SEARCH_OPT_LT: compare_result = (elem < needle); break;
            case TARRAY_SEARCH_OPT_EQ:
            default: compare_result = (elem == needle); break;
            }
            if (compare_result == compare_wanted)
                break;
        }
        if (offset >= haystackP->used) {
            /* No match */
            resultObj = Tcl_NewObj();
        } else {
            if (flags & TARRAY_SEARCH_INLINE)
                resultObj = Tcl_NewWideIntObj(elem);
            else
                resultObj = Tcl_NewIntObj(offset);
        }
    }

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}


static TCL_RESULT TArraySearchDouble(Tcl_Interp *interp, TArrayHdr * haystackP,
                                          Tcl_Obj *needleObj, int start, enum TArraySearchSwitches op, int flags)
{
    int offset;
    Tcl_Obj *resultObj;
    double dval, *dvalP;
    int compare_result;
    int compare_wanted;

    TARRAY_ASSERT(haystackP->type == TARRAY_DOUBLE);

    switch (op) {
    case TARRAY_SEARCH_OPT_GT:
    case TARRAY_SEARCH_OPT_LT: 
    case TARRAY_SEARCH_OPT_EQ:
        break;
    default:
        return TArrayBadSearchOpError(interp, op);
    }

    if (Tcl_GetDoubleFromObj(interp, needleObj, &dval) != TCL_OK)
        return TCL_ERROR;
    
    compare_wanted = flags & TARRAY_SEARCH_INVERT ? 0 : 1;

    /* First locate the starting point for the search */
    dvalP = TAHDRELEMPTR(haystackP, double, start);

    if (flags & TARRAY_SEARCH_ALL) {
        TArrayHdr *thdrP;

        thdrP = TArrayAlloc(
            flags & TARRAY_SEARCH_INLINE ? TARRAY_DOUBLE : TARRAY_INT,
            10);                /* Assume 10 hits */

        for (offset = start; offset < haystackP->used; ++offset, ++dvalP) {
            switch (op) {
            case TARRAY_SEARCH_OPT_GT: compare_result = (*dvalP > dval); break;
            case TARRAY_SEARCH_OPT_LT: compare_result = (*dvalP < dval); break;
            case TARRAY_SEARCH_OPT_EQ: compare_result = (*dvalP == dval); break;
            }

            if (compare_result == compare_wanted) {
                /* Have a match */
                /* Ensure enough space in target array */
                if (thdrP->used >= thdrP->allocated)
                    thdrP = TArrayRealloc(thdrP, thdrP->used + TARRAY_EXTRA(thdrP->used));
                if (flags & TARRAY_SEARCH_INLINE) {
                    *TAHDRELEMPTR(thdrP, double, thdrP->used) = *dvalP;
                } else {
                    *TAHDRELEMPTR(thdrP, int, thdrP->used) = offset;
                }
                thdrP->used++;
            }
        }

        resultObj = TArrayNewObj(thdrP);

    } else {
        /* Return first found element */
        for (offset = start; offset < haystackP->used; ++offset, ++dvalP) {
            switch (op) {
            case TARRAY_SEARCH_OPT_GT: compare_result = (*dvalP > dval); break;
            case TARRAY_SEARCH_OPT_LT: compare_result = (*dvalP < dval); break;
            case TARRAY_SEARCH_OPT_EQ: compare_result = (*dvalP == dval); break;
            }
            if (compare_result == compare_wanted)
                break;
        }
        if (offset >= haystackP->used) {
            /* No match */
            resultObj = Tcl_NewObj();
        } else {
            if (flags & TARRAY_SEARCH_INLINE)
                resultObj = Tcl_NewDoubleObj(*dvalP);
            else
                resultObj = Tcl_NewIntObj(offset);
        }
    }

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}

static TCL_RESULT TArraySearchObj(Tcl_Interp *interp, TArrayHdr * haystackP,
                                  Tcl_Obj *needleObj, int start, enum TArraySearchSwitches op, int flags)
{
    int offset;
    Tcl_Obj **objPP;
    Tcl_Obj *resultObj;
    int compare_result;
    int compare_wanted;
    int nocase;
    Tcl_RegExp re;

    /* TBD - do we need to increment the haystacP ref to guard against shimmering */
    TARRAY_ASSERT(haystackP->type == TARRAY_OBJ);
    
    compare_wanted = flags & TARRAY_SEARCH_INVERT ? 0 : 1;
    nocase = flags & TARRAY_SEARCH_NOCASE;

    switch (op) {
    case TARRAY_SEARCH_OPT_GT:
    case TARRAY_SEARCH_OPT_LT: 
    case TARRAY_SEARCH_OPT_EQ:
    case TARRAY_SEARCH_OPT_PAT:
        break;
    case TARRAY_SEARCH_OPT_RE:
        /* Following lsearch implementation, get the regexp before any
           shimmering can take place, and try to compile for the efficient
           NOSUB case
        */
        re = Tcl_GetRegExpFromObj(NULL, needleObj,
                                  TCL_REG_ADVANCED|(nocase ? TCL_REG_NOCASE : 0)|TCL_REG_NOSUB );
        if (re == NULL) {
            /* That failed, so try without the NOSUB flag */
            re = Tcl_GetRegExpFromObj(interp, needleObj,
                                      TCL_REG_ADVANCED|(nocase ? TCL_REG_NOCASE : 0));
            if (re == NULL)
                return TCL_ERROR;
        }
        break;
    default:
        return TArrayBadSearchOpError(interp, op);
    }

    /* First locate the starting point for the search */
    objPP = TAHDRELEMPTR(haystackP, Tcl_Obj *, start);


    if (flags & TARRAY_SEARCH_ALL) {
        TArrayHdr *thdrP;

        thdrP = TArrayAlloc(
            flags & TARRAY_SEARCH_INLINE ? TARRAY_OBJ : TARRAY_INT,
            10);                /* Assume 10 hits */

        for (offset = start; offset < haystackP->used; ++offset, ++objPP) {
            switch (op) {
            case TARRAY_SEARCH_OPT_GT:
                compare_result = TArrayCompareObjs(*objPP, needleObj, nocase) > 0; break;
            case TARRAY_SEARCH_OPT_LT: 
                compare_result = TArrayCompareObjs(*objPP, needleObj, nocase) < 0; break;
            case TARRAY_SEARCH_OPT_EQ:
                compare_result = TArrayCompareObjs(*objPP, needleObj, nocase) == 0; break;
            case TARRAY_SEARCH_OPT_PAT:
                compare_result = Tcl_StringCaseMatch(Tcl_GetString(*objPP),
                                                     Tcl_GetString(needleObj),
                                                     nocase ? TCL_MATCH_NOCASE : 0);
                break;
            case TARRAY_SEARCH_OPT_RE:
                compare_result = Tcl_RegExpExecObj(interp, re, *objPP,
                                                   0, 0, 0);
                if (compare_result < 0) {
                    TArrayFreeHdr(thdrP); /* Note this unrefs embedded Tcl_Objs if needed */
                    return TCL_ERROR;
                }
                break;
            }
            if (compare_result == compare_wanted) {
                /* Have a match */
                /* Ensure enough space in target array */
                if (thdrP->used >= thdrP->allocated)
                    thdrP = TArrayRealloc(thdrP, thdrP->used + TARRAY_EXTRA(thdrP->used));
                if (flags & TARRAY_SEARCH_INLINE) {
                    Tcl_IncrRefCount(*objPP);
                    *TAHDRELEMPTR(thdrP, Tcl_Obj *, thdrP->used) = *objPP;
                } else {
                    *TAHDRELEMPTR(thdrP, int, thdrP->used) = offset;
                }
                thdrP->used++;
            }
        }

        resultObj = TArrayNewObj(thdrP);

    } else {
        /* Return first found element */
        for (offset = start; offset < haystackP->used; ++offset, ++objPP) {
            switch (op) {
            case TARRAY_SEARCH_OPT_GT:
                compare_result = TArrayCompareObjs(*objPP, needleObj, nocase) > 0; break;
            case TARRAY_SEARCH_OPT_LT: 
                compare_result = TArrayCompareObjs(*objPP, needleObj, nocase) < 0; break;
            case TARRAY_SEARCH_OPT_EQ:
                compare_result = TArrayCompareObjs(*objPP, needleObj, nocase) == 0; break;
            case TARRAY_SEARCH_OPT_PAT:
                compare_result = Tcl_StringCaseMatch(Tcl_GetString(*objPP),
                                                     Tcl_GetString(needleObj),
                                                     nocase ? TCL_MATCH_NOCASE : 0);
                break;
            case TARRAY_SEARCH_OPT_RE:
                compare_result = Tcl_RegExpExecObj(interp, re, *objPP,
                                                   0, 0, 0);
                if (compare_result < 0)
                    return TCL_ERROR;
                break;
            }
            if (compare_result == compare_wanted)
                break;
        }
        if (offset >= haystackP->used) {
            /* No match */
            resultObj = Tcl_NewObj();
        } else {
            if (flags & TARRAY_SEARCH_INLINE)
                resultObj = *objPP; /* No need to incr ref, the SetObjResult does it */
            else
                resultObj = Tcl_NewIntObj(offset);
        }
    }

    Tcl_SetObjResult(interp, resultObj);
    return TCL_OK;
}

TCL_RESULT TArray_SearchObjCmd(ClientData clientdata, Tcl_Interp *interp,
                              int objc, Tcl_Obj *const objv[])
{
    int flags;
    int start_index;
    int i, n, opt;
    TArrayHdr *haystackP;
    enum TArraySearchSwitches op;

    if (objc < 3) {
	Tcl_WrongNumArgs(interp, 1, objv, "?options? tarray pattern");
	return TCL_ERROR;
    }

    if (TArrayVerifyType(interp, objv[objc-2]) != TCL_OK)
        return TCL_ERROR;

    flags = 0;
    start_index = 0;
    op = TARRAY_SEARCH_OPT_EQ;
    for (i = 1; i < objc-2; ++i) {
	if (Tcl_GetIndexFromObj(interp, objv[i], TArraySearchSwitches, "option", 0, &opt)
            != TCL_OK) {
            return TCL_ERROR;
	}
        switch ((enum TArraySortSwitches) opt) {
        case TARRAY_SEARCH_OPT_ALL: flags |= TARRAY_SEARCH_ALL; break;
        case TARRAY_SEARCH_OPT_INLINE: flags |= TARRAY_SEARCH_INLINE; break;
        case TARRAY_SEARCH_OPT_INVERT: flags |= TARRAY_SEARCH_INVERT; break;
        case TARRAY_SEARCH_OPT_NOCASE: flags |= TARRAY_SEARCH_NOCASE; break;
        case TARRAY_SEARCH_OPT_START:
            if (i > objc-4)
                return TArrayBadArgError(interp, "-start");
            ++i;
            /*
             * To prevent shimmering, check if the index object is same
             * as tarray object.
             */
            if (objv[i] == objv[objc-2]) {
                Tcl_Obj *dupObj = Tcl_DuplicateObj(objv[i]);
                n = Tcl_GetIntFromObj(interp, dupObj, &start_index);
                Tcl_DecrRefCount(dupObj);
                if (n != TCL_OK)
                    return TCL_ERROR;
            } else {
                if (Tcl_GetIntFromObj(interp, objv[i], &start_index) != TCL_OK)
                    return TCL_ERROR;
            }
            break;
        case TARRAY_SEARCH_OPT_EQ:
        case TARRAY_SEARCH_OPT_GT:
        case TARRAY_SEARCH_OPT_LT:
        case TARRAY_SEARCH_OPT_PAT:
        case TARRAY_SEARCH_OPT_RE:
            op = (enum TArraySortSwitches) opt;
        }
    }

    haystackP = TARRAYHDR(objv[objc-2]);
    switch (haystackP->type) {
    case TARRAY_BOOLEAN:
        return TArraySearchBoolean(interp, haystackP, objv[objc-1], start_index,op,flags);
    case TARRAY_INT:
    case TARRAY_UINT:
    case TARRAY_BYTE:
    case TARRAY_WIDE:
        return TArraySearchEntier(interp, haystackP, objv[objc-1], start_index, op, flags);
    case TARRAY_DOUBLE:
        return TArraySearchDouble(interp, haystackP, objv[objc-1], start_index, op, flags);
    case TARRAY_OBJ:
        return TArraySearchObj(interp, haystackP, objv[objc-1], start_index, op, flags);
    default:
        Tcl_SetResult(interp, "Not implemented", TCL_STATIC);
        return TCL_ERROR;
    }

}


int TArrayCompareObjs(Tcl_Obj *oaP, Tcl_Obj *obP, int ignorecase)
{
    char *a, *b;
    int alen, blen, len;
    int comparison;

    a = Tcl_GetStringFromObj(oaP, &alen);
    alen = Tcl_NumUtfChars(a, alen); /* Num bytes -> num chars */
    b = Tcl_GetStringFromObj(obP, &blen);
    blen = Tcl_NumUtfChars(b, blen); /* Num bytes -> num chars */

    len = alen < blen ? alen : blen; /* len is the shorter length */
    
    comparison = (ignorecase ? Tcl_UtfNcasecmp : Tcl_UtfNcmp)(a, b, len);

    if (comparison == 0) {
        comparison = alen-blen;
    }
    return (comparison > 0) ? 1 : (comparison < 0) ? -1 : 0;
}



/* Find number bits set in a bit array */
int TArrayNumSetBits(TArrayHdr *thdrP)
{
    int v, count;
    int i, n;
    int *iP;

    TARRAY_ASSERT(thdrP->type == TARRAY_BOOLEAN);
    
    n = thdrP->used / (sizeof(int)*CHAR_BIT); /* Number of ints */
    for (count = 0, i = 0, iP = TAHDRELEMPTR(thdrP, int, 0); i < n; ++i, ++iP) {
        // See http://graphics.stanford.edu/~seander/bithacks.html#CountBitsSetNaive
        v = *iP;
        v = v - ((v >> 1) & 0x55555555); // reuse input as temporary
        v = (v & 0x33333333) + ((v >> 2) & 0x33333333);     // temp
        count += ((v + (v >> 4) & 0xF0F0F0F) * 0x1010101) >> 24; // count
    }
//    TBD - in alloc code make sure allocations are aligned on longest elem size;
    /* Remaining bits */
    n = thdrP->used % (sizeof(int)*CHAR_BIT); /* Number of left over bits */
    if (n) {
        /* *iP points to next int, however, not all bytes in that int are
           valid. Mask off invalid bits */
        unsigned char *ucP = (unsigned char *) iP;
        /* Note value of v will change depending on endianness but no matter
           as we only care about number of 1's */
        for (i = 0, v = 0; n >= CHAR_BIT; ++i, n -= CHAR_BIT) {
            v = (v << 8) | ucP[i];
        }
        if (n) {
            v = (v << 8) | (ucP[i] & (-BITMASK(n-1))  );
        }

        v = v - ((v >> 1) & 0x55555555);
        v = (v & 0x33333333) + ((v >> 2) & 0x33333333);
        count += ((v + (v >> 4) & 0xF0F0F0F) * 0x1010101) >> 24;
    }

    return count;
}

/*
 * Following bitcopy code from stackoverflow.com. Bits are indexed
 * from MSB (0) to LSB (7)
 */
#define PREPARE_FIRST_COPY()                                      \
    do {                                                          \
    if (src_len >= (CHAR_BIT - dst_offset_modulo)) {              \
        *dst     &= reverse_mask[dst_offset_modulo];              \
        src_len -= CHAR_BIT - dst_offset_modulo;                  \
    } else {                                                      \
        *dst     &= reverse_mask[dst_offset_modulo]               \
              | reverse_mask_xor[dst_offset_modulo + src_len + 1];\
         c       &= reverse_mask[dst_offset_modulo + src_len    ];\
        src_len = 0;                                              \
    } } while (0)


void
bitarray_copy(const unsigned char *src_org, int src_offset, int src_len,
                    unsigned char *dst_org, int dst_offset)
{
    static const unsigned char mask[] =
        { 0x55, 0x01, 0x03, 0x07, 0x0f, 0x1f, 0x3f, 0x7f, 0xff };
    static const unsigned char mask_xor[] =
        { 0x55, 0xfe, 0xfc, 0xf8, 0xf0, 0xe0, 0xc0, 0x80, 0x00 };
    static const unsigned char reverse_mask[] =
        { 0x55, 0x80, 0xc0, 0xe0, 0xf0, 0xf8, 0xfc, 0xfe, 0xff };
    static const unsigned char reverse_mask_xor[] =
        { 0xff, 0x7f, 0x3f, 0x1f, 0x0f, 0x07, 0x03, 0x01, 0x00 };

    if (src_len) {
        const unsigned char *src;
              unsigned char *dst;
        int                  src_offset_modulo,
                             dst_offset_modulo;

        src = src_org + (src_offset / CHAR_BIT);
        dst = dst_org + (dst_offset / CHAR_BIT);

        src_offset_modulo = src_offset % CHAR_BIT;
        dst_offset_modulo = dst_offset % CHAR_BIT;

        if (src_offset_modulo == dst_offset_modulo) {
            int              byte_len;
            int              src_len_modulo;
            if (src_offset_modulo) {
                unsigned char   c;

                c = reverse_mask_xor[dst_offset_modulo]     & *src++;

                PREPARE_FIRST_COPY();
                *dst++ |= c;
            }

            byte_len = src_len / CHAR_BIT;
            src_len_modulo = src_len % CHAR_BIT;

            if (byte_len) {
                memcpy(dst, src, byte_len);
                src += byte_len;
                dst += byte_len;
            }
            if (src_len_modulo) {
                *dst     &= reverse_mask_xor[src_len_modulo];
                *dst |= reverse_mask[src_len_modulo]     & *src;
            }
        } else {
            int             bit_diff_ls,
                            bit_diff_rs;
            int             byte_len;
            int             src_len_modulo;
            unsigned char   c;
            /*
             * Begin: Line things up on destination. 
             */
            if (src_offset_modulo > dst_offset_modulo) {
                bit_diff_ls = src_offset_modulo - dst_offset_modulo;
                bit_diff_rs = CHAR_BIT - bit_diff_ls;

                c = *src++ << bit_diff_ls;
                c |= *src >> bit_diff_rs;
                c     &= reverse_mask_xor[dst_offset_modulo];
            } else {
                bit_diff_rs = dst_offset_modulo - src_offset_modulo;
                bit_diff_ls = CHAR_BIT - bit_diff_rs;

                c = *src >> bit_diff_rs     &
                    reverse_mask_xor[dst_offset_modulo];
            }
            PREPARE_FIRST_COPY();
            *dst++ |= c;

            /*
             * Middle: copy with only shifting the source. 
             */
            byte_len = src_len / CHAR_BIT;

            while (--byte_len >= 0) {
                c = *src++ << bit_diff_ls;
                c |= *src >> bit_diff_rs;
                *dst++ = c;
            }

            /*
             * End: copy the remaing bits; 
             */
            src_len_modulo = src_len % CHAR_BIT;
            if (src_len_modulo) {
                c = *src++ << bit_diff_ls;
                c |= *src >> bit_diff_rs;
                c     &= reverse_mask[src_len_modulo];

                *dst     &= reverse_mask_xor[src_len_modulo];
                *dst |= c;
            }
        }
    }
}

/*
 * Comparison functions for qsort. For a stable sort, if two items are
 * equal, we compare the addresses. Note expression is a bit more complicated
 * because we cannot just return result of a subtraction when the type
 * is wider than an int.
 */
#define RETCMP(a_, b_, t_)                              \
    do {                                                \
        if (*(t_*)(a_) > *(t_*)(b_))                    \
            return 1;                                   \
        else if (*(t_*)(a_) < *(t_*)(b_))               \
            return -1;                                  \
        else {                                          \
            if ((char *)(a_) == (char *)(b_))           \
                return 0;                               \
            else if ((char *)(a_) > (char *)(b_))       \
                return 1;                               \
            else                                        \
                return -1;                              \
        }                                               \
    } while (0)

/* Compare in reverse (for decreasing order). Note we cannot just
   reverse arguments to RETCMP since we are also comparing actual
   addresses and the sense for those should be fixed.
*/
#define RETCMPREV(a_, b_, t_)                           \
    do {                                                \
        if (*(t_*)(a_) > *(t_*)(b_))                    \
            return -1;                                  \
        else if (*(t_*)(a_) < *(t_*)(b_))               \
            return 1;                                   \
        else {                                          \
            if ((char *)(a_) == (char *)(b_))           \
                return 0;                               \
            else if ((char *)(a_) > (char *)(b_))       \
                return 1;                               \
            else                                        \
                return -1;                              \
        }                                               \
    } while (0)

int intcmp(const void *a, const void *b) { RETCMP(a,b,int); }
int intcmprev(const void *a, const void *b) { RETCMPREV(a,b,int); }
int uintcmp(const void *a, const void *b) { RETCMP(a,b,unsigned int); }
int uintcmprev(const void *a, const void *b) { RETCMPREV(a,b,unsigned int); }
int widecmp(const void *a, const void *b) { RETCMP(a,b,Tcl_WideInt); }
int widecmprev(const void *a, const void *b) { RETCMPREV(a,b,Tcl_WideInt); }
int doublecmp(const void *a, const void *b) { RETCMP(a,b,double); }
int doublecmprev(const void *a, const void *b) { RETCMPREV(a,b,double); }
int bytecmp(const void *a, const void *b) { RETCMPREV(a,b,unsigned char); }
int bytecmprev(const void *a, const void *b) { RETCMPREV(a,b,unsigned char); }
int tclobjcmp(const void *a, const void *b)
{
    int n;
    n = TArrayCompareObjs(*(Tcl_Obj **)a, *(Tcl_Obj **)b, 0);
    if (n)
        return n;
    else
        return a > b ? 1 : (a < b ? -1 : 0);
}
int tclobjcmprev(const void *a, const void *b)
{
    int n;
    n = TArrayCompareObjs(*(Tcl_Obj **)b, *(Tcl_Obj **)a, 0);
    if (n)
        return n;
    else
        return a > b ? 1 : (a < b ? -1 : 0);
}

/*
 * Comparison functions for tarray_qsort_r. 
 * The passed values are integer indexes into a TArrayHdr
 * For a stable sort, if two items are
 * equal, we compare the indexes.
 * ai_ and bi_ are pointers to int indexes into the TArrayHdr
 * pointed to by v_ which is of type t_
 */
#define RETCMPINDEXED(ai_, bi_, t_, v_)                 \
    do {                                                \
        t_ a_ = *(*(int*)(ai_) + (t_ *)v_); \
        t_ b_ = *(*(int*)(bi_) + (t_ *)v_);             \
        if (a_ < b_)                                    \
            return -1;                                  \
        else if (a_ > b_)                               \
            return 1;                                   \
        else                                            \
            return *(int *)(ai_) - *(int *)(bi_);       \
    } while (0)


/* Compare in reverse (for decreasing order). Note we cannot just
   reverse arguments to RETCMPINDEXED since we are also comparing actual
   addresses and the sense for those should be fixed.
*/
#define RETCMPINDEXEDREV(ai_, bi_, t_, v_)              \
    do {                                                \
        t_ a_ = *(*(int*)(ai_) + (t_ *)v_); \
        t_ b_ = *(*(int*)(bi_) + (t_ *)v_);             \
        if (a_ < b_)                                    \
            return 1;                                   \
        else if (a_ > b_)                               \
            return -1;                                  \
        else                                            \
            return *(int *)(ai_) - *(int *)(bi_);       \
    } while (0)

int intcmpindexed(void *ctx, const void *a, const void *b) { RETCMPINDEXED(a,b,int, ctx); }
int intcmpindexedrev(void *ctx, const void *a, const void *b) { RETCMPINDEXEDREV(a,b,int, ctx); }
int uintcmpindexed(void *ctx, const void *a, const void *b) { RETCMPINDEXED(a,b,unsigned int, ctx); }
int uintcmpindexedrev(void *ctx, const void *a, const void *b) { RETCMPINDEXEDREV(a,b,unsigned int, ctx); }
int widecmpindexed(void *ctx, const void *a, const void *b) { RETCMPINDEXED(a,b,Tcl_WideInt, ctx); }
int widecmpindexedrev(void *ctx, const void *a, const void *b) { RETCMPINDEXEDREV(a,b,Tcl_WideInt, ctx); }
int doublecmpindexed(void *ctx, const void *a, const void *b) { RETCMPINDEXED(a,b,double, ctx); }
int doublecmpindexedrev(void *ctx, const void *a, const void *b) { RETCMPINDEXEDREV(a,b,double, ctx); }
int bytecmpindexed(void *ctx, const void *a, const void *b) { RETCMPINDEXEDREV(a,b,unsigned char, ctx); }
int bytecmpindexedrev(void *ctx, const void *a, const void *b) { RETCMPINDEXEDREV(a,b,unsigned char, ctx); }
int tclobjcmpindexed(void *ctx, const void *ai, const void *bi)
{
    int n;
    Tcl_Obj *a = *(*(int *)ai + (Tcl_Obj **)ctx);
    Tcl_Obj *b = *(*(int *)bi + (Tcl_Obj **)ctx);
    n = TArrayCompareObjs(a, b, 0);
    if (n)
        return n;
    else
        return *(int *)ai - *(int *)bi;
}
int tclobjcmpindexedrev(void *ctx, const void *ai, const void *bi)
{
    int n;
    Tcl_Obj *a = *(*(int *)ai + (Tcl_Obj **)ctx);
    Tcl_Obj *b = *(*(int *)bi + (Tcl_Obj **)ctx);
    n = TArrayCompareObjs(b, a, 0);
    if (n)
        return n;
    else
        return *(int *)ai - *(int *)bi;
}
int booleancmpindexed(void *ctx, const void *ai, const void *bi)
{
    unsigned char *ucP = (unsigned char *)ctx;
    unsigned char uca, ucb;

    uca = ucP[(*(int*) ai)/CHAR_BIT] & BITMASK((*(int*)ai) % CHAR_BIT);
    ucb = ucP[(*(int*) bi)/CHAR_BIT] & BITMASK((*(int*)bi) % CHAR_BIT);

    if (!uca == !ucb) {
        /* Both bits set or unset, use index position to differentiate */
        return *(int *)ai - *(int *)bi;
    } else if (uca)
        return 1;
    else
        return -1;
}

int booleancmpindexedrev(void *ctx, const void *ai, const void *bi)
{
    unsigned char *ucP = (unsigned char *)ctx;
    unsigned char uca, ucb;

    uca = ucP[(*(int*) ai)/CHAR_BIT] & BITMASK((*(int*)ai) % CHAR_BIT);
    ucb = ucP[(*(int*) bi)/CHAR_BIT] & BITMASK((*(int*)bi) % CHAR_BIT);

    if (!uca == !ucb) {
        /* Both bits set or unset, use index position to differentiate */
        return *(int *)ai - *(int *)bi;
    } else if (uca)
        return -1;
    else
        return 1;
}



/*
 * qsort code - from BSD via speedtables
 */

/*-
 * Copyright (c) 1992, 1993
 *	The Regents of the University of California.  All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions
 * are met:
 * 1. Redistributions of source code must retain the above copyright
 *    notice, this list of conditions and the following disclaimer.
 * 2. Redistributions in binary form must reproduce the above copyright
 *    notice, this list of conditions and the following disclaimer in the
 *    documentation and/or other materials provided with the distribution.
 * 4. Neither the name of the University nor the names of its contributors
 *    may be used to endorse or promote products derived from this software
 *    without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND
 * ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED.  IN NO EVENT SHALL THE REGENTS OR CONTRIBUTORS BE LIABLE
 * FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
 * DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS
 * OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
 * HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
 * LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY
 * OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF
 * SUCH DAMAGE.
 */

#if defined(LIBC_SCCS) && !defined(lint)
static char sccsid[] = "@(#)qsort.c	8.1 (Berkeley) 6/4/93";
#endif /* LIBC_SCCS and not lint */

#if !defined(_MSC_VER)
#include <sys/cdefs.h>
#endif

#include <stdlib.h>

#define I_AM_QSORT_R

#ifdef _MSC_VER
// MS VC 6 does not like "inline" in C files
#define inline __inline
#endif

#ifdef I_AM_QSORT_R
typedef int		 cmp_t(void *, const void *, const void *);
#else
typedef int		 cmp_t(const void *, const void *);
#endif
static inline char	*med3(char *, char *, char *, cmp_t *, void *);
static inline void	 swapfunc(char *, char *, int, int);

#ifndef _MSC_VER
#define min(a, b)	(((a) < (b)) ? (a) : (b))
#endif

/*
 * Qsort routine from Bentley & McIlroy's "Engineering a Sort Function".
 */
#define swapcode(TYPE, parmi, parmj, n) { 		\
	long i = (n) / sizeof (TYPE); 			\
	TYPE *pi = (TYPE *) (parmi); 		\
	TYPE *pj = (TYPE *) (parmj); 		\
	do { 						\
		TYPE	t = *pi;		\
		*pi++ = *pj;				\
		*pj++ = t;				\
        } while (--i > 0);				\
}

#define SWAPINIT(a, es) swaptype = ((char *)a - (char *)0) % sizeof(long) || \
	es % sizeof(long) ? 2 : es == sizeof(long)? 0 : 1;

static inline void
swapfunc(a, b, n, swaptype)
	char *a, *b;
	int n, swaptype;
{
	if(swaptype <= 1)
		swapcode(long, a, b, n)
	else
		swapcode(char, a, b, n)
}

#define swap(a, b)					\
	if (swaptype == 0) {				\
		long t = *(long *)(a);			\
		*(long *)(a) = *(long *)(b);		\
		*(long *)(b) = t;			\
	} else						\
		swapfunc(a, b, es, swaptype)

#define vecswap(a, b, n) 	if ((n) > 0) swapfunc(a, b, n, swaptype)

#ifdef I_AM_QSORT_R
#define	CMP(t, x, y) (cmp((t), (x), (y)))
#else
#define	CMP(t, x, y) (cmp((x), (y)))
#endif

static inline char *
med3(char *a, char *b, char *c, cmp_t *cmp, void *thunk
#ifndef I_AM_QSORT_R
__unused
#endif
)
{
	return CMP(thunk, a, b) < 0 ?
	       (CMP(thunk, b, c) < 0 ? b : (CMP(thunk, a, c) < 0 ? c : a ))
              :(CMP(thunk, b, c) > 0 ? b : (CMP(thunk, a, c) < 0 ? a : c ));
}

#ifdef I_AM_QSORT_R
void
tarray_qsort_r(void *a, size_t n, size_t es, void *thunk, cmp_t *cmp)
#else
#define thunk NULL
void
qsort(void *a, size_t n, size_t es, cmp_t *cmp)
#endif
{
	char *pa, *pb, *pc, *pd, *pl, *pm, *pn;
	size_t d, r;
	int cmp_result;
	int swaptype, swap_cnt;

loop:	SWAPINIT(a, es);
	swap_cnt = 0;
	if (n < 7) {
		for (pm = (char *)a + es; pm < (char *)a + n * es; pm += es)
			for (pl = pm; 
			     pl > (char *)a && CMP(thunk, pl - es, pl) > 0;
			     pl -= es)
				swap(pl, pl - es);
		return;
	}
	pm = (char *)a + (n / 2) * es;
	if (n > 7) {
		pl = a;
		pn = (char *)a + (n - 1) * es;
		if (n > 40) {
			d = (n / 8) * es;
			pl = med3(pl, pl + d, pl + 2 * d, cmp, thunk);
			pm = med3(pm - d, pm, pm + d, cmp, thunk);
			pn = med3(pn - 2 * d, pn - d, pn, cmp, thunk);
		}
		pm = med3(pl, pm, pn, cmp, thunk);
	}
	swap(a, pm);
	pa = pb = (char *)a + es;

	pc = pd = (char *)a + (n - 1) * es;
	for (;;) {
		while (pb <= pc && (cmp_result = CMP(thunk, pb, a)) <= 0) {
			if (cmp_result == 0) {
				swap_cnt = 1;
				swap(pa, pb);
				pa += es;
			}
			pb += es;
		}
		while (pb <= pc && (cmp_result = CMP(thunk, pc, a)) >= 0) {
			if (cmp_result == 0) {
				swap_cnt = 1;
				swap(pc, pd);
				pd -= es;
			}
			pc -= es;
		}
		if (pb > pc)
			break;
		swap(pb, pc);
		swap_cnt = 1;
		pb += es;
		pc -= es;
	}
	if (swap_cnt == 0) {  /* Switch to insertion sort */
		for (pm = (char *)a + es; pm < (char *)a + n * es; pm += es)
			for (pl = pm; 
			     pl > (char *)a && CMP(thunk, pl - es, pl) > 0;
			     pl -= es)
				swap(pl, pl - es);
		return;
	}

	pn = (char *)a + n * es;
	r = min(pa - (char *)a, pb - pa);
	vecswap(a, pb - r, r);
	r = min(pd - pc, pn - pd - es);
	vecswap(pb, pn - r, r);
	if ((r = pb - pa) > es)
#ifdef I_AM_QSORT_R
                tarray_qsort_r(a, r / es, es, thunk, cmp);
#else
		qsort(a, r / es, es, cmp);
#endif
	if ((r = pd - pc) > es) {
		/* Iterate rather than recurse to save stack space */
		a = pn - r;
		n = r / es;
		goto loop;
	}
/*		qsort(pn - r, r / es, es, cmp);*/
}

