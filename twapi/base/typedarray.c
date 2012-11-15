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
#define TARRAY_INT 1
    "wide",
#define TARRAY_WIDE 2
    "double",
#define TARRAY_DOUBLE 3
    "byte",
#define TARRAY_BYTE 4
};    

#define TARRAY_MAX_ELEM_SIZE \
    (sizeof(double) > sizeof(long) ? (sizeof(double) > sizeof(void*) ? sizeof(double) : sizeof(void*)) : sizeof(long))
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
    void *thdrP = TARRAYHDR(objP);

    TARRAY_ASSERT(srcP->typePtr == &gTArrayType);

    thdrP = TARRAYHDR(objP); 
    TARRAY_ASSERT(thdrP);

    if (--thdrP->nrefs)
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
    char *cP;
    
    TARRAY_ASSERT(objP->bytes == NULL);
    TARRAY_ASSERT(objP->typePtr == &gTArrayType);

    count = TARRAYELEMCOUNT(objP);

    if (count == 0) {
	objP->bytes = ckalloc(sizeof(objP->bytes[0]));
        objP->bytes[0] = 0;
	objP->length = 0;
	return;
    }

    /* Code below based on count > 0 else terminating \0 will blow memory */

    /*
     * When output size cannot be calculated exactly, we allocate using
     * some estimage 
    
    switch (TARRAYTYPE(objP)) {
    case TARRAY_BOOLEAN:
        {
            unsigned char *boolP;
            unsigned char uc;
            /* For BOOLEANS, we know how long a buffer needs to be */
            cP = ckalloc(2*count);
            objP->bytes = cP;
            boolP = TARRAYELEMPTR(objP, unsigned char, 0);
            n = count / 8;
            for (i = 0; i < n; ++i, ++boolP) {
                for (uc = 1; uc ; uc <<= 1) {
                    *cP++ = (*boolP & uc) ? '1' : '0';
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
        }
        break;
    case TARRAY_INT:

    case TARRAY_WIDE: sz = sizeof(Tcl_WideInt); break; TBD;
    case TARRAY_DOUBLE: sz = sizeof(double); break; TBD;
    case TARRAY_BYTE: sz = sizeof(unsigned char); break; TBD;
    default:
        UpdateObjArrayString(objP, TARRAYELEMPTR(objP, Tcl_Obj *, 0), count);
        break;
    }
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

