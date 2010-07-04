#ifndef MEMLIFO_H
#define MEMLIFO_H

/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <windows.h>
#include <stdlib.h>
#include "memlifo.h"

#define TWAPI_MEMLIFO_ASSERT(x) ((void) 0) /* TBD */

typedef struct _TwapiMemLifo TwapiMemLifo;
typedef struct _TwapiMemLifoMarkDesc TwapiMemLifoMarkDesc;
typedef TwapiMemLifoMarkDesc *TwapiMemLifoMarkHandle;

typedef void *TwapiMemLifoChunkAllocFn(size_t sz, void *alloc_data, size_t *actual_szP);
typedef void TwapiMemLifoChunkFreeFn(void *p, void *alloc_data);

struct _TwapiMemLifo {
    void *lifo_allocator_data;           /* For use by allocation functions as
                                            they see fit */
    TwapiMemLifoChunkAllocFn *lifo_allocFn;
    TwapiMemLifoChunkFreeFn *lifo_freeFn;
    TwapiMemLifoMarkHandle lifo_top_mark;  /* Topmost mark */
    TwapiMemLifoMarkHandle lifo_bot_mark; /* Bottommost mark */
    size_t	lifo_chunk_size;   /* Size of each chunk to allocate.
                                      Note this size might not be a multiple
                                      of the alignment size */
    LONG		lifo_magic;	/* Only used in debug mode */
#define TWAPI_MEMLIFO_MAGIC 0xb92c610a
};

#define LIFODSCSZ (TwapiMem_ALIGNSIZEUP(sizeof(TwapiMemLifoDsc_t)))


/*f
Create a Last-In-First-Out memory pool

Creates a memory pool from which memory can be allocated in Last-In-First-Out
fashion.

On success, returns a handle for the LIFO memory pool. On failure, returns 0. 
In this case, the caller can obtain the error code through APNgetError().

See also TwapiMemLifoAlloc, TwapiMemLifoMark, TwapiMemLifoPopMark, 
TwapiMemLifoReset

Returns a handle for the LIFO memory pool on success, 0 on failure.
*/
int TwapiMemLifoInit
(
    TwapiMemLifo *lifoP,
    void *allocator_data,
    TwapiMemLifoChunkAllocFn *allocFunc,
                                /* Pointer to routine to be used to allocate
				   memory. Must return aligned memory.
				   The parameter indicates the amount of
				   memory to allocate. Note this function
				   must be callable at ANY time including
				   during the TwapiMemLifoInit call itself.
				   Naturally, this function must not make a
				   call to LIFO memory allocation routines
				   from the same pool. This parameter may be
				   indicated as 0, in which case allocation
				   will be done using an internal default. */
    TwapiMemLifoChunkFreeFn *freeFunc,
				/* Pointer to routine to be used for freeing
				   memory allocated through allocFunc. The
				   parameter points to the memory to be
				   freed. This routine must be specified
				   unless allocFunc is 0. If allocFunc is 0,
				   the value of this parameter is ignored. */
     size_t chunkSz             /* Default unit size for allocating memory
				   from (*allocFunc) as needed. This does
				   not include space for descriptor at start
				   of each allocation. The implementation
				   has a minimum default which will be used
				   if chunkSz is too small. */
     );


/*f
Free up resources associated with a LIFO memory pool

Frees up various resources allocated for a LIFO memory pool. The pool must not
be used after the function is called.
*/
void TwapiMemLifoClose(TwapiMemLifo  *lifoP);

/*f
Allocate memory from LIFO memory pool

Allocates memory from a LIFO memory pool and returns a pointer to it.

See also TwapiMemLifoInit, TwapiMemLifoMark, TwapiMemLifoPopMark,
TwapiMemLifoPushFrame

Returns pointer to allocated memory on success, a null pointer on failure.
*/
void* TwapiMemLifoAlloc
    (
     TwapiMemLifo *lifoP,       /* LIFO pool to allocate from */
     size_t sz			/* Number of bytes to allocate */
    );

/*f
Allocate a software stack frame in a LIFO memory pool

Allocates space in a LIFO memory pool being used as a software stack.
This provides a means of maintaining a software stack for temporary structures
that are too large to be allocated on the hardware stack.

Both TwapiMemLifoMark and TwapiMemLifoPushFrame may be used on the same
TwapiMemLifo_t. The latter function in effect creates a anonymous mark
that is maintained internally and which may be released (along with
the associated user memory) through the function TwapiMemLifoPopFrame
which releases the last allocated mark, irrespective of whether it was
allocated through TwapiMemLifoMark or TwapiMemLifoPushFrame.
Alternatively, the mark and associated memory are also freed when a
previosly allocated mark is released.

See also TwapiMemLifoAlloc, TwapiMemLifoPopFrame

Returns pointer to allocated memory on success, a null pointer on failure.
*/
void *TwapiMemLifoPushFrame
(
    TwapiMemLifo *lifoP,		/* LIFO pool to allocate from */
    size_t sz			/* Number of bytes to allocate */
    );

/*
Release the topmost mark from a TwapiMemLifo_t

Releases the topmost (last allocated) mark from a TwapiMemLifo_t. The
mark may have been allocated using either TwapiMemLifoMark or
TwapiMemLifoPushFrame.

See also TwapiMemLifoAlloc, TwapiMemLifoPushFrame, TwapiMemLifoMark

Returns APNer_SUCCESS or APNer_xxx status code on error
*/
int TwapiMemLifoPopFrame(TwapiMemLifo *lifoP);

/*
Mark current state of a LIFO memory pool

Stores the current state of a LIFO memory pool on a stack. The state can be
restored later by calling TwapiMemLifoPopMark. Any number of marks may be
created but they must be popped in reverse order. However, not all marks
need be popped since popping a mark automatically pops all marks created
after it.

See also TwapiMemLifoAlloc, TwapiMemLifoPopMark, TwapiMemLifoMarkedAlloc,
TwapiMemLifoAllocFrame

Returns a handle for the mark if successful, 0 on failure.
*/
TwapiMemLifoMarkHandle TwapiMemLifoPushMark(TwapiMemLifo *lifoP);


/*f
Restore state of a LIFO memory pool

Restores the state of a LIFO memory pool that was previously saved
using TwapiMemLifoMark or TwapiMemLifoMarkedAlloc.  Memory allocated
from the pool between the push and this pop is freed up. Caller must
not subsequently call this routine with marks created between the
TwapiMemLifoMark and this TwapiMemLifoPopMark as they will have been
freed as well. The mark being passed to this routine is freed as well
and hence must not be reused.

See also TwapiMemLifoAlloc, TwapiMemLifoMark, TwapiMemLifoMarkedAlloc

Returns 0 on success, 1 on failure
*/
int TwapiMemLifoPopMark(TwapiMemLifoMarkHandle mark);


/*f
Expand the last memory block allocated from a LIFO memory pool

Expands the last memory block allocated from a LIFO memory pool. If no
memory was allocated in the pool since the last mark, just allocates
a new block of size incr.

The function may move the block if necessary unless the dontMove parameter
is non-0 in which case the function will only attempt to expand the block in
place. When dontMove is 0, caller must be careful to update pointers that
point into the original block since it may have been moved on return.

On success, the size of the block is guaranteed to be increased by at
least the requested increment. The function may fail if the last
allocated block cannot be expanded without moving and dontMove is set,
if it is not the last allocation from the LIFO memory pool, or if a mark
has been allocated after this block, or if there is insufficient memory.

On failure, the function return a NULL pointer.

Returns pointer to new block position on success, else a NULL pointer.
*/
void * TwapiMemLifoExpandLast
    (
     TwapiMemLifo *lifoP,		/* Lifo pool from which alllocation
					   was made */
     size_t incr,			/* The amount by which the
					   block is to be expanded. */
     int dontMove			/* If 0, block may be moved if
					   necessary to expand. If non-0, the
					   function will fail if the block
					   cannot be expanded in place. */
    );



/*f
Shrink the last memory block allocated from a LIFO memory pool

Shrinks the last memory block allocated from a LIFO memory pool. No marks
must have been allocated in the pool since the last memory allocation else
the function will fail.

The function may move the block to compact memory unless the dontMove parameter
is non-0. When dontMove is 0, caller must be careful to update pointers that
point into the original block since it may have been moved on return.

On success, the size of the block is not guaranteed to have been decreased.

Returns pointer to address of relocated block or a null pointer if a mark
was allocated after the last allocation.
*/
void *TwapiMemLifoShrinkLast
    (
     TwapiMemLifo *lifoP,       /* Lifo pool from which alllocation
                                   was made */
     size_t decr,               /* The amount by which the
                                   block is to be shrunk. */
     int dontMove              /* If 0, block may be moved in order
                                   to reduce memory fragmentation.
                                   If non-0, the the block will be
                                   shrunk in place or be left as is. */
    );


/*f
Resize the last memory block allocated from a LIFO memory pool

Resizes the last memory block allocated from a LIFO memory pool.  No marks
must have been allocated in the pool since the last memory allocation else
the function will fail.  The block may be moved if necessary unless the
dontMove parameter is non-0 in which the function will only attempt to resiz
e the block in place. When dontMove is 0, caller must be careful to update
pointers that point into the original block since it may have been moved on
return.

On success, the block is guaranteed to be at least as large as the requested
size. The function may fail if the block cannot be resized without moving
and dontMove is set, or if there is insufficient memory.

Returns pointer to new block position on success, else a NULL pointer.
*/
void * TwapiMemLifoResizeLast
(
    TwapiMemLifo *lifoP,        /* Lifo pool from which allocation was made */
    size_t newSz,		/* New size of the block */
    int dontMove                /* If 0, block may be moved if
                                   necessary to expand. If non-0, the
                                   function will fail if the block
                                   cannot be expanded in place. */
    );

#endif
