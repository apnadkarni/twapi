#ifndef TARRAY_H
#define TARRAY_H

#include <limits.h>
#include <stdlib.h>

typedef int TCL_RESULT;

/* Must match gTArrayTypeTokens definition in tarray.c ! */
#define TARRAY_BOOLEAN 0
#define TARRAY_UINT 1
#define TARRAY_INT 2
#define TARRAY_WIDE 3
#define TARRAY_DOUBLE 4
#define TARRAY_BYTE 5
#define TARRAY_OBJ 6

/* If building out of twapi pool, use its settings */
#if defined(TWAPI_ENABLE_ASSERT) && !defined(TARRAY_ENABLE_ASSERT)
#define TARRAY_ENABLE_ASSERT TWAPI_ENABLE_ASSERT
#endif


#if TARRAY_ENABLE_ASSERT
#  if TARRAY_ENABLE_ASSERT == 1
#    define TARRAY_ASSERT(bool_) (void)( (bool_) || (Tcl_Panic("Assertion (%s) failed at line %d in file %s.", #bool_, __LINE__, __FILE__), 0) )
#  elif TARRAY_ENABLE_ASSERT == 2
#    define TARRAY_ASSERT(bool_) (void)( (bool_) || (DebugOutput("Assertion (" #bool_ ") failed at line " MAKESTRINGLITERAL2(__LINE__) " in file " __FILE__ "\n"), 0) )
#  elif TARRAY_ENABLE_ASSERT == 3
#    define TARRAY_ASSERT(bool_) do { if (! (bool_)) { __asm int 3 } } while (0)
#  else
#    error Invalid value for TARRAY_ENABLE_ASSERT
#  endif
#else
#define TARRAY_ASSERT(bool_) ((void) 0)
#endif

typedef Tcl_Obj *TArrayObjPtr;

extern const char *gTArrayTypeTokens[];

/* How many slots to allocate by default */
#define TARRAY_DEFAULT_NSLOTS 1000

#define TARRAY_MAX_ELEM_SIZE (sizeof(double) > sizeof(int) ? (sizeof(double) > sizeof(void*) ? sizeof(double) : sizeof(void*)) : sizeof(int))
#define TARRAY_MAX_COUNT (1 + (int)(((size_t)UINT_MAX - sizeof(TArrayHdr))/TARRAY_MAX_ELEM_SIZE))

typedef union TArrayHdr_s {
    void *pointer_aligner;
    double double_aligner;
    struct {
        int nrefs;              /* Ref count when shared between Tcl_Objs */
        int allocated;
        int used;
        unsigned char type;
        unsigned char elem_bits; /* Size of element in bits */
    };
} TArrayHdr;
#define TAHDRELEMPTR(thdr_, type_, index_) ((index_) + (type_ *)(sizeof(TArrayHdr) + (char *) (thdr_)))
#define TAHDRELEMUSEDBYTES(thdr_) ((((thdr_)->used * (thdr_)->elem_bits) + CHAR_BIT-1) / CHAR_BIT)

#define TARRAYDATA(optr_)  ((optr_)->internalRep.ptrAndLongRep.ptr)
#define TARRAYHDR(optr_) ((TArrayHdr *)TARRAYDATA(optr_))
#define TARRAYTYPE(optr_) (TARRAYHDR(optr_)->type)
#define TARRAYELEMSLOTS(optr_) ((TARRAYHDR(optr_))->allocated)
#define TARRAYELEMCOUNT(optr_) ((TARRAYHDR(optr_))->used)
#define TARRAYELEMPTR(optr_, type_, index_) TAHDRELEMPTR(TARRAYHDR(optr_), type_, index_)

/* How much extra slots to allocate when allocating memory. n_ should
 * be number of elements currently.
 */
#define TARRAY_EXTRA(n_)  \
    ((n_) < 10 ? 10 : ((n_) < 100 ? (n_) : ((n_) < 800 ? 100 : ((n_)/8))))

/* Can a TArrayHdr block be modified ? Must be unshared and large enough */
#define TARRAYHDR_SHARED(th_) ((th_)->nrefs > 1)
#define TARRAYHDR_WRITABLE(th_, size_) (TARRAYHDR_SHARED(th_) && (th_)->allocated >= (size_))

extern struct Tcl_ObjType gTArrayType;
#define TARRAY_OBJ_SETREP(optr_, thdr_) \
    do {                                        \
        (thdr_)->nrefs++;                       \
        TARRAYHDR(optr_) = thdr_;               \
        (optr_)->typePtr = &gTArrayType;         \
    } while (0)


/* ALLOCATE_ARRAY call should panic on failure to allocate */
#define TARRAY_ALLOCMEM ckalloc
#define TARRAY_FREEMEM(p_) if (p_) ckfree((char *)p_)
#define TARRAY_REALLOCMEM ckrealloc

/* Search flags */
#define TARRAY_SEARCH_INLINE 1  /* Return values, not indices */
#define TARRAY_SEARCH_INVERT 2  /* Invert matching expression */
#define TARRAY_SEARCH_ALL    4  /* Return all matches */

/* Prototypes - generated by cl */
extern void __cdecl TArrayTypePanic(unsigned char tatype);
extern void __cdecl TArrayIncrObjRefs(TArrayHdr *thdrP,int first,int count);
extern void __cdecl TArrayDecrObjRefs(TArrayHdr *thdrP,int first,int count);
extern void __cdecl TArrayFreeHdr(TArrayHdr *thdrP);
extern int __cdecl TArrayVerifyType(struct Tcl_Interp *interp,struct Tcl_Obj *objP);
extern Tcl_Obj * __cdecl TArrayNewObj(TArrayHdr *thdrP);
extern int __cdecl TArraySetFromObjs(struct Tcl_Interp *interp,TArrayHdr *thdrP,int first,int nelems,struct Tcl_Obj *const *elems );
extern int __cdecl TArrayCalcSize(unsigned char tatype,int count);
extern TArrayHdr *__cdecl TArrayRealloc(TArrayHdr *oldP,int new_count);
extern TArrayHdr *__cdecl TArrayAlloc(unsigned char tatype, int count);
extern TArrayHdr *__cdecl TArrayAllocAndInit(struct Tcl_Interp *interp,unsigned char tatype,int nelems,struct Tcl_Obj *const *elems ,int init_size);
extern int __cdecl TArraySet(struct Tcl_Interp *interp,TArrayHdr *dstP,int dst_first,TArrayHdr *srcP,int src_first,int count);
extern TArrayHdr *__cdecl TArrayClone(TArrayHdr *srcP, int init_size);
extern struct Tcl_Obj *__cdecl TArrayIndex(struct Tcl_Interp *interp,TArrayHdr *thdrP,int index);
extern TArrayHdr *__cdecl TArrayConvertToIndices(struct Tcl_Interp *interp,struct Tcl_Obj *objP);
extern TArrayHdr *__cdecl TArrayGetValues(struct Tcl_Interp *interp,TArrayHdr *srcP,TArrayHdr *indicesP);
extern int __cdecl TArrayNumSetBits(TArrayHdr *thdrP);
extern TCL_RESULT __cdecl TArraySetRange(Tcl_Interp *interp, TArrayHdr *dstP, int dst_first, int count, Tcl_Obj *objP);
extern TCL_RESULT __cdecl TArraySearchBoolean(Tcl_Interp *interp, TArrayHdr *haystackP,
                                              Tcl_Obj *needleObj, int start, int flags);

extern void __cdecl bitarray_set(unsigned char *ucP, int offset, int count, int ival);
extern void __cdecl bitarray_copy(const unsigned char *src_org, int src_offset,
                                  int src_len, unsigned char *dst_org, int dst_offset);

extern void __cdecl tarray_qsort_r(void *a, size_t n, size_t es, void *thunk, int (__cdecl *cmp)(void *, const void *, const void *));
extern int __cdecl intcmp(const void *a, const void *b);
extern int __cdecl intcmprev(const void *a, const void *b);
extern int __cdecl uintcmp(const void *a, const void *b);
extern int __cdecl uintcmprev(const void *a, const void *b);
extern int __cdecl widecmp(const void *a, const void *b);
extern int __cdecl widecmprev(const void *a, const void *b);
extern int __cdecl doublecmp(const void *a, const void *b);
extern int __cdecl doublecmprev(const void *a, const void *b);
extern int __cdecl bytecmp(const void *a, const void *b);
extern int __cdecl bytecmprev(const void *a, const void *b);
extern int __cdecl tclobjcmp(const void *a, const void *b);
extern int __cdecl tclobjcmprev(const void *a, const void *b);

extern int __cdecl intcmpindexed(void *, const void *a, const void *b);
extern int __cdecl intcmpindexedrev(void *, const void *a, const void *b);
extern int __cdecl uintcmpindexed(void *, const void *a, const void *b);
extern int __cdecl uintcmpindexedrev(void *, const void *a, const void *b);
extern int __cdecl widecmpindexed(void *, const void *a, const void *b);
extern int __cdecl widecmpindexedrev(void *, const void *a, const void *b);
extern int __cdecl doublecmpindexed(void *, const void *a, const void *b);
extern int __cdecl doublecmpindexedrev(void *, const void *a, const void *b);
extern int __cdecl bytecmpindexed(void *, const void *a, const void *b);
extern int __cdecl bytecmpindexedrev(void *, const void *a, const void *b);
extern int __cdecl tclobjcmpindexed(void *, const void *a, const void *b);
extern int __cdecl tclobjcmpindexedrev(void *, const void *a, const void *b);

extern void __cdecl TArrayBadArgError(Tcl_Interp *interp, const char *optname);

#endif
