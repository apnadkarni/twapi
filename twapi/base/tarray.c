#include <limits.h>
#include "tcl.h"
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

static int bitarray_copy(const unsigned char *src_org, int src_offset,
                         int src_len, unsigned char *dst_org, int dst_offset);

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

void TArrayTypePanic(unsigned char tatype)
{
    Tcl_Panic("Unknown tarray type %d", tatype);
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

    TARRAY_ASSERT(srcP->typePtr == &gTArrayType);

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
                if (ival > 255 || ival < 255) {
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

TArrayHdr *TArrayRealloc(Tcl_Interp *interp, TArrayHdr *oldP, int new_count)
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
        if (init_size == 0)
            init_size = TARRAY_DEFAULT_NSLOTS;
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
TArrayHdr *TArrayClone(TArrayHdr *srcP, int init_size)
{
    TArrayHdr *thdrP;

    if (init_size == 0)
        init_size = srcP->allocated;
    else if (init_size < srcP->used)
        init_size = srcP->used;

    /* TBD - optimize these two calls */
    thdrP = TArrayAlloc(srcP->type, init_size);
    if (TArraySet(NULL, thdrP, 0, srcP, 0, srcP->used) != TCL_OK) {
        TArrayFreeHdr(thdrP);
        return NULL;
    }
    return thdrP;
}

/* dstP must not be shared and must be large enough */
TCL_RESULT TArraySetRange(Tcl_Interp *interp, TArrayHdr *dstP, int dst_first,
                          int count, Tcl_Obj *objP)
{
    int i, n, ival;
    unsigned char *ucP;

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
            unsigned uc;
            unsigned char uc_mask; /* The bit position corresponding to 'first' */

            /* First set the bits to get to a char boundary */
            ucP = TAHDRELEMPTR(dstP, unsigned char, dst_first / CHAR_BIT);
            uc = dst_first % CHAR_BIT; /* Offset of bit within a char */
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


static int
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

