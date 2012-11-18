/*
 * TArray is a Tcl "type" used for densely storing arrays of elements
 * of a specific type.
 * The Tcl_Obj.internalRep.ptrAndLongRep.value type of an element.
 * and Tcl_Obj.internalRep.ptrAndLongRep.ptr holds a pointer to 
 * an allocated array of that type.
 */

static char *gTArrayTypes[] = {
    "boolean",
#define TARRAY_BOOLEAN 0
    "int",
#define TARRAY_UINT 1
    "uint",
#define TARRAY_INT 2
    "wide",
#define TARRAY_WIDE 3
    "double",
#define TARRAY_DOUBLE 4
    "byte",
#define TARRAY_BYTE 5
    "tclobj",
#define TARRAY_OBJ 6
};    

#define TARRAY_MAX_ELEM_SIZE \
    (sizeof(double) > sizeof(int) ? (sizeof(double) > sizeof(void*) ? sizeof(double) : sizeof(void*)) : sizeof(int))
#define TARRAY_MAX_COUNT \
    (1 + (int)(((size_t)UINT_MAX - sizeof(TArrayHdr))/TARRAY_MAX_ELEM_SIZE))

typedef union TArrayHdr_s {
    void *aligner;
    struct {
        int nrefs;              /* Ref count when shared between Tcl_Objs */
        int allocated;
        int used;
    };
} TArrayHdr;

#define TARRAYTYPE(optr_) ((optr_)->internalRep.ptrAndLongRep.value)
#define TARRAYDATA(optr_)  ((optr_)->internalRep.ptrAndLongRep.ptr)
#define TARRAYHDR(optr_) ((TArrayHdr *)TARRAYDATA(optr_))
#define TARRAYELEMSLOTS(optr_) ((TARRAYHDR(optr_))->allocated)
#define TARRAYELEMCOUNT(optr_) ((TARRAYHDR(optr_))->used)
#define TARRAYELEMPTR(optr_, type_, index_) \
    ((index_) + (type_ *)(sizeof(TArrayHdr) + (char *) TARRAYDATA(optr_)))




static void DupTArray(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void FreeTArray(Tcl_Obj *objP);
static void UpdateTArrayString(Tcl_Obj *objP);
static struct Tcl_ObjType gTArrayType = {
    "TArray",
    FreeTArray,
    DupTArray,
    UpdateTArrayString,
    NULL,     /* jenglish advises to keep this NULL */
};


/* ALLOCATE_ARRAY call should panic on failure to allocate */
#define TARRAY_ALLOC ckalloc
#define TARRAY_FREE(p_) if (p_) ckfree(p_)

static void FreeTArray(Tcl_Obj *objP)
{
    void *thdrP;

    TARRAY_ASSERT(srcP->typePtr == &gTArrayType);

    thdrP = TARRAYHDR(objP); 
    TARRAY_ASSERT(thdrP);

    if (--thdrP->nrefs == 0)
        TARRAY_FREE(thdrP);
    TARRAYHDR(objP) = NULL;
    objP->typePtr = NULL;
}

static void DupTArray(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    TARRAY_ASSERT(srcP->typePtr == &gTArrayType);
    TARRAY_ASSERT(TARRAYDATA(srcP) != NULL);
    
    TARRAYTYPE(dstP) = TARRAYTYPE(srcP);
    TARRAYHDR(dstP) = TARRAYHDR(srcP);
    TARRAYHDR(dstP)->nrefs += 1;

    dstP->typePtr = &gTArrayType;
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

	flagPtr = ckalloc(objc * sizeof(int));
    }
    for (i = 0; i < objc; i++) {
	flagPtr[i] = (i ? TCL_DONT_QUOTE_HASH : 0);
	elem = TclGetStringFromObj(objv[i], &length);
	bytesNeeded += TclScanElement(elem, length, flagPtr+i);
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
	elem = TclGetStringFromObj(objv[i], &length);
	dst += TclConvertElement(elem, length, dst, flagPtr[i]);
	*dst++ = ' ';
    }
    objP->bytes[objP->length] = '\0';

    if (flagPtr != localFlags) {
	ckfree(flagPtr);
    }
}


static void UpdateTArrayString(Tcl_Obj *objP)
{
    int i, n, count;
    int allocated, unused;
    char *cP;
    unsigned char uc;
    unsigned char *ucP;
    int max_elem_space;  /* Max space to print one element including
                            either terminating null or space */
    
    TARRAY_ASSERT(objP->typePtr == &gTArrayType);

    objP->bytes = NULL;

    count = TARRAYELEMCOUNT(objP);
    if (count == 0) {
	objP->bytes = ckalloc(sizeof(objP->bytes[0]));
        objP->bytes[0] = 0;
	objP->length = 0;
	return;
    }

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
        /* For BOOLEANS, we know how long a buffer needs to be */
        cP = ckalloc(2*count);
        objP->bytes = cP;
        ucP = TARRAYELEMPTR(objP, unsigned char, 0);
        n = count / 8;
        for (i = 0; i < n; ++i, ++ucP) {
            for (uc = 1; uc ; uc <<= 1) {
                *cP++ = (*ucP & uc) ? '1' : '0';
                *cP++ = ' ';
            }
        }
        n = count - n*8;    /* Left over bits in last byte */
        if (n) {
            for (i = 0, uc = 1; i < n; ++i, uc <<= 1) {
                *cP++ = (uc & j) ? '1' : '0';
                *cP++ = ' ';
            }
        }
        cP[-1] = 0;         /* Overwrite last space with terminating \0 */
        objP->length = 2*count - 1;
        return;

    TARRAY_OBJ:
        UpdateObjArrayString(objP, TARRAYELEMPTR(objP, Tcl_Obj *, 0), count);
        return;
        
    TARRAY_UINT:
    TARRAY_INT:
        TARRAY_ASSERT(sizeof(int) == 4); /* So max string space needed is 11 */
        max_elem_space = 11+1;
        break;
    TARRAY_WIDE:
        max_elem_space = TCL_INTEGER_SPACE+1;
        break;
    TARRAY_DOUBLE:
        max_elem_space = TCL_DOUBLE_SPACE+1;
        break;
    TARRAY_BYTE:
        max_elem_space = 3+1;
        break;
    default:
        Tcl_Panic("Unknown TypedArray type %d", TARRAYTYPE(objP));
    }

    allocated = 0;
    unused = 0;
    objP->bytes= NULL;
    /* Nested loop for efficiency reasons to avoid switch on every iter */
    for (i = 0; i < count; ++i) {
        if (unused < max_elem_space) {
            n = allocated - unused; /* Used space */
            /* Increase assuming remaining take half max space on average */
            allocated += ((max_elem_size + 1)/2)*(count - i);
            objP->bytes = ckrealloc(objP->bytes, allocated);
            cP = n + (char *) objP->bytes;
            unused = allocated - n;
        }
        switch (TARRAYTYPE(objP)) {
        TARRAY_UINT:
            _snprintf(cP, unused, "%u", *TARRAYELEMPTR(objP, unsigned int, i));
            break;
        TARRAY_INT:
            _snprintf(cP, unused, "%d", *TARRAYELEMPTR(objP, int, i));
            break;
        TARRAY_WIDE:
            _snprintf(cP, unused, "%" TCL_LL_MODIFIER "d", *TARRAYELEMPTR(objP, Tcl_WideInt, i));
            break;
        TARRAY_DOUBLE:
            /* Do not use _snprintf because of slight difference
               it does not include decimal point for whole ints. For
               consistency with Tcl, use Tcl_PrintDouble instead */
            Tcl_PrintDouble(NULL, *TARRAYELEMPTR(objP, Tcl_WideInt, i), cP);
            break;
        TARRAY_BYTE:
            _snprintf(cP, unused, "%u", *TARRAYELEMPTR(objP, unsigned char, i));
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
    if (unused > (allocated / 8))
        objP->bytes = ckrealloc(objP->bytes, allocated - unused);
    return;
}

static void XXXDupTArray(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    void *newP;
    int sz;
    int do_incrref = 0;

    TARRAY_ASSERT(srcP->typePtr == &gTArrayType);
    TARRAY_ASSERT(TARRAYDATA(srcP) != NULL);

    if (TARRAYTYPE(srcP) == TARRAY_BOOLEAN)
        sz = sizeof(*thdrP) + (TARRAYELEMSLOTS(srcP) / sizeof(unsigned char));
    else {
        switch (TARRAYTYPE(srcP)) {
        case TARRAY_INT: sz = sizeof(long); break;
        case TARRAY_WIDE: sz = sizeof(Tcl_WideInt); break;
        case TARRAY_DOUBLE: sz = sizeof(double); break;
        case TARRAY_BYTE: sz = sizeof(unsigned char); break;
        default:
            sz = sizeof(Tcl_Obj *);
            do_incrref = 1;
            break;
        }
        sz = sizeof(*thdrP) + (TARRAYELEMSLOTS(srcP) * sz);
    }
    newP = ALLOCATE_ARRAY(sz);
    memmove(newP, TARRAYDATA(srcP), sz);
    if (do_incrref) {
        /* Tcl_Obj's are now pointed from both srcP and dstP */
        int i;
        Tcl_Obj **objPP = TARRAYELEMPTR(thdrP, Tcl_Obj *, i);
        for (i = 0; i < thdrP->used; ++i) {
            Tcl_IncrRefCount(*objPP++);
        }
    }


    dstP->typePtr = &gTArrayType;
    TARRAYTYPE(dstP) = TARRAYTYPE(srcP);
    TARRAYDATA(dstP) = newP;
}

