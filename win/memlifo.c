/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#define MEMLIFO_MAX_ALLOC TCL_SIZE_MAX

typedef struct _MemLifoChunk MemLifoChunk;

/*
Each region is composed of a linked list of contiguous chunks of memory. Each
chunk is prefixed by a descriptor which is also used to link the chunks.
*/
typedef struct _MemLifoChunk {
    MemLifoChunk *lc_prev;	/* Pointer to next chunk */
    void         *lc_end;	/* One beyond end of chunk */
} MemLifoChunk;

/*
A mark keeps current state information about a MemLifo which can
be used to restore it to a previous state. On initialization, a mark
is internally created to hold the initial state for MemLifo.
Later, marks are added and deleted through application calls.

The topmost mark (last one allocated) holds the current state of the
MemLifo.

When chunks are allocated, they are put a single list shared among all
the marks which will point to sublists within the list. When a mark is
"popped", all chunks on the list that are not part of the sublist
pointed to by the previous mark are freed.  "Big block" allocations,
are handled the same way.

Since marks also keeps track of the free space in the current chunk
from which allocations are made, the state of the MemLifo_t can be
completely restored when popping a mark by simply making the previous
mark the topmost (current) mark.
*/


typedef struct _MemLifoMark {
    int			lm_magic;	/* Only used in debug mode */
#define MEMLIFO_MARK_MAGIC	0xa0193d4f
    int			lm_seq;		/* Only used in debug mode */
    MemLifo       *lm_lifo; /* MemLifo for this mark */
    MemLifoMark *lm_prev; /* Previous mark */
    void *		lm_last_alloc; /* last allocation from the lifo */
    MemLifoChunk *lm_big_blocks; /* Pointer to list of large block
                                         allocations. Note marks should
                                         never be allocated from here since
                                         big blocks may be deleted during
                                         some reallocations */
    MemLifoChunk *lm_chunks; /* Current chunk used for allocation
                                        and the head of the chunk list */
    void *lm_freeptr;           /* Ptr to unused space */
} MemLifoMark;

static void *MemLifoDefaultAlloc(MemLifoSize sz, HANDLE heap, MemLifoSize *actual)
{
    void *p = HeapAlloc(heap, 0, sz);
    if (actual)
        *actual = (MemLifoSize)HeapSize(heap, 0, p);
    return p;
}

static void MemLifoDefaultFree(void *p, HANDLE heap)
{
    HeapFree(heap, 0, p);
}


int MemLifoInit(
    MemLifo *l,
    void *allocator_data,
    MemLifoChunkAllocFn *allocFunc,
    MemLifoChunkFreeFn *freeFunc,
    MemLifoSize chunk_sz,
    int flags
    )
{
    MemLifoChunk *c;
    MemLifoMarkHandle m;
    MemLifoSize actual_chunk_sz;

    if (allocFunc == 0) {
        allocator_data = HeapCreate(0, 0, 0);
        if (allocator_data == NULL)
            return GetLastError();
	allocFunc = MemLifoDefaultAlloc;
	freeFunc = MemLifoDefaultFree;
    } else {
        MEMLIFO_ASSERT(freeFunc);	/* If allocFunc was not 0, freeFunc
					   should not be either */
    }

    if (chunk_sz < 1000)
        chunk_sz = 1000;

    /* Allocate a chunk and allocate space for the lifo descriptor from it */
    c = allocFunc(chunk_sz, allocator_data, &actual_chunk_sz);
    if (c == 0) {
        if (flags & MEMLIFO_F_PANIC_ON_FAIL)
            Tcl_Panic("Could not initialize memlifo");
        return ERROR_OUTOFMEMORY;
    }

    c->lc_prev = NULL;
    c->lc_end = ADDPTR(c, actual_chunk_sz, void*);

    l->lifo_allocator_data = allocator_data;
    l->lifo_allocFn = allocFunc;
    l->lifo_freeFn = freeFunc;
    l->lifo_chunk_size = ROUNDUP(chunk_sz); /* What caller asked, not actual_chunk_sz */
    l->lifo_flags = flags;
    l->lifo_magic = MEMLIFO_MAGIC;

    /* Allocate mark from chunk itself */
    m = ALIGNPTR(c, sizeof(*c), MemLifoMark*);

    m->lm_magic = MEMLIFO_MARK_MAGIC;
    m->lm_seq = 1;

    m->lm_freeptr = ALIGNPTR(m, sizeof(*m), void*);
    m->lm_lifo = l;
    m->lm_prev = m;/* Point back to itself. Effectively will never be popped */
    m->lm_big_blocks = 0;
    m->lm_last_alloc = 0;
    m->lm_chunks = c;

    l->lifo_top_mark = m;
    l->lifo_bot_mark = m;

    return ERROR_SUCCESS;
}

void MemLifoClose(MemLifo *l)
{
    /*
     * Note as a special case, popping the bottommost mark does not release
     * the mark itself
     */
    MemLifoPopMark(l->lifo_bot_mark);

    MEMLIFO_ASSERT(l->lifo_bot_mark);
    MEMLIFO_ASSERT(l->lifo_bot_mark->lm_chunks);

    /* Finally free the chunk containing the bottom mark */
    l->lifo_freeFn(l->lifo_bot_mark->lm_chunks,
                       l->lifo_allocator_data);
    TwapiZeroMemory(l, sizeof(*l));
}

void* MemLifoAlloc(MemLifo *l,  MemLifoSize sz, MemLifoSize *actual_szP)
{
    MemLifoChunk *c;
    MemLifoMarkHandle m;
    MemLifoSize chunk_sz;
    void *p;

    if (sz > MEMLIFO_MAX_ALLOC) {
        if (l->lifo_flags & MEMLIFO_F_PANIC_ON_FAIL)
            Tcl_Panic("Attempt to allocate %" MEMLIFO_SIZE_MODIFIER "u bytes for memlifo", sz);
        return NULL;
    }

    /*
     * NOTE: note that when called from MemLifoExpandLast(), on entry the
     * lm_last_alloc field may not be set correctly since that function
     * may remove the last allocation from the big block list under
     * some circumstances. We do not use that field here, only set it.
     */

    MEMLIFO_ASSERT(l->lifo_magic == MEMLIFO_MAGIC);
    MEMLIFO_ASSERT(l->lifo_bot_mark);

    m = l->lifo_top_mark;
    MEMLIFO_ASSERT(m);
    MEMLIFO_ASSERT(ALIGNED(m->lm_freeptr));

    sz = ROUNDUP(sz);
    p = ADDPTR(m->lm_freeptr, sz, void*); /* New end of used space */
    if (p > (void*) m->lm_chunks                  /* Ensure no wrap-around! */
        && p <= m->lm_chunks->lc_end) {
        /* OK, not beyond the end of the chunk */
	m->lm_last_alloc = m->lm_freeptr;
        MEMLIFO_ASSERT(ALIGNED(m->lm_last_alloc));
        /*
         * If actual_szP is non-NULL, caller wants to allocate at least sz
         * but as much as possible without allocating a new chunk
         */
        if (actual_szP) {
            m->lm_freeptr = m->lm_chunks->lc_end;
            *actual_szP = PTRDIFF32(m->lm_freeptr, m->lm_last_alloc);
        } else
            m->lm_freeptr = p;
	return 	m->lm_last_alloc;
    }

    /*
     * Insufficient space in current chunk.
     * Decide whether to allocate a new chunk or allocate a separate
     * block for the request. We allocate a chunk if
     * if there is little space available in the current chunk,
     * `little' being defined as less than 1/8th chunk size. Note
     * this also ensures we will not allocate separate blocks that
     * are smaller than 1/8th the chunk size.
     * Otherwise, we satisfy the request by allocating a separate
     * block.
     *
     * This strategy is intended to balance conflicting goals: on one
     * hand, we do not want to allocate a new chunk aggressively as
     * this will result in too much wasted space in the current chunk.
     * On the other hand, allocating a separate block results in wasted
     * space due to external fragmentation as well slower execution.
     *
     * TBD - if using our default windows based heap, try and HeapReAlloc
     * in place.
     */
    /* TBD - make it /2 instead of /8 ? */
    if (PTRDIFF32(m->lm_chunks->lc_end, m->lm_freeptr) < (int) l->lifo_chunk_size/8) {
        /* Little space in the current chunk
         * Allocate a new chunk and suballocate from it.
         *
	 * As a heuristic, we will allocate extra space in the chunk
	 * if the requested size is greater than half the default chunk size.
	 */
	if (sz > (l->lifo_chunk_size/2))
            chunk_sz = sz + l->lifo_chunk_size;
	else
	    chunk_sz = l->lifo_chunk_size;

        MEMLIFO_ASSERT(ROUNDED(chunk_sz));
        chunk_sz += ROUNDUP(sizeof(MemLifoChunk));

	c = l->lifo_allocFn(chunk_sz, l->lifo_allocator_data, &chunk_sz);
	if (c == 0) {
            if (l->lifo_flags & MEMLIFO_F_PANIC_ON_FAIL)
                Tcl_Panic("Attempt to allocate %" MEMLIFO_SIZE_MODIFIER "u bytes for memlifo", chunk_sz);
	    return 0;
        }

	c->lc_end = ADDPTR(c, chunk_sz, void*);

	c->lc_prev = m->lm_chunks;	/* Place on the list of chunks */
	m->lm_chunks = c;
	
	m->lm_last_alloc = ALIGNPTR(c, sizeof(*c), void*);
        m->lm_freeptr = ALIGNPTR(m->lm_last_alloc, sz, void*);
        /* Notice that when we have to allocate a new chunk, we do not
         * give more than caller asked (modulo some rounding)
         */
        if (actual_szP)
            *actual_szP = PTRDIFF32(m->lm_freeptr, m->lm_last_alloc);
    }
    else {
	/* Allocate a separate big block. */
        MemLifoSize actual_size;
        chunk_sz = sz + ROUNDUP(sizeof(MemLifoChunk));

	c = (MemLifoChunk *) l->lifo_allocFn(chunk_sz,
                                                  l->lifo_allocator_data,
                                                  &actual_size);
	if (c == 0) {
            if (l->lifo_flags & MEMLIFO_F_PANIC_ON_FAIL)
                Tcl_Panic("Attempt to allocate %" MEMLIFO_SIZE_MODIFIER "u bytes for memlifo", chunk_sz);
	    return 0;
        }
	c->lc_end = ADDPTR(c, actual_size, void*);

	c->lc_prev = m->lm_big_blocks;	/* Place on the list of big blocks */
	m->lm_big_blocks = c;
	/* 
	 * Note we do not modify m->m_freeptr since it still refers to 
	 * the current "mainstream" chunk.
	 */
	m->lm_last_alloc = ALIGNPTR(c, sizeof(*c), void*);
        if (actual_szP) {
            *actual_szP = PTRDIFF32(c->lc_end, m->lm_last_alloc);
#ifdef TWAPI_MEMLIFO_DEBUG
            MEMLIFO_ASSERT(*actual_szP >= sz);
#endif
        }
    }

    return m->lm_last_alloc;
}

void* MemLifoCopy(MemLifo *l, void *srcP, MemLifoSize nbytes)
{
    void *dstP = MemLifoAlloc(l, nbytes, NULL);
    if (dstP)
        memcpy(dstP, srcP, nbytes);
    return dstP;
}

void* MemLifoZeroes(MemLifo *l, MemLifoSize nbytes)
{
    void *dstP = MemLifoAlloc(l, nbytes, NULL);
    if (dstP)
        memset(dstP, 0, nbytes);
    return dstP;
}

MemLifoMarkHandle MemLifoPushMark(MemLifo *l)
{
    MemLifoMarkHandle m;			/* Previous (existing) mark */
    MemLifoMarkHandle n;			/* New mark */
    MemLifoChunk *c;
    MemLifoSize chunk_sz;
    MemLifoSize mark_sz;
    void *p;
    
    m = l->lifo_top_mark;

    /* NOTE - marks must never be allocated from big block list  */

    /*
     * Check for common case first - enough space in current chunk to
     * hold the mark.
     */
    mark_sz = ROUNDUP(sizeof(MemLifoMark));
    p = ADDPTR(m->lm_freeptr, mark_sz, void*); /* Potential end of mark */
    if (p > (void*) m->lm_chunks                  /* Ensure no wrap-around! */
        && p <= m->lm_chunks->lc_end) {
        /* Enough room for the mark in this chunk */
        n = (MemLifoMark *) m->lm_freeptr;
	n->lm_freeptr = p;
	n->lm_chunks = m->lm_chunks;
    }
    else {
	/* 
	 * No room in current chunk. Have to allocate a new chunk. Note 
	 * we do not use MemLifoAlloc to allocate the mark since that 
	 * would change the state of the previous mark.
	 */
	c = l->lifo_allocFn(l->lifo_chunk_size,
                            l->lifo_allocator_data,
                            &chunk_sz);
	if (c == 0) {
            if (l->lifo_flags & MEMLIFO_F_PANIC_ON_FAIL)
                Tcl_Panic("Attempt to allocate %" MEMLIFO_SIZE_MODIFIER "u bytes for memlifo", l->lifo_chunk_size);
	    return 0;
        }	
	c->lc_end = ADDPTR(c, chunk_sz, void*);	

	/* 
	 * Place on the list of chunks. Note however, that we do NOT 
	 * modify m->lm_chunkList since that should hold the original lifo 
	 * state. We'll put this chunk on the list headed by the new mark.
	 */
	c->lc_prev = m->lm_chunks;	/* Place on the list of chunks */

	n = ALIGNPTR(c, sizeof(*c), MemLifoMark*);
	n->lm_chunks = c;
	
	n->lm_freeptr = ALIGNPTR(n, sizeof(*n), void*);
    }

#ifdef MEMLIFO_DEBUG
    n->lm_magic = MEMLIFO_MARK_MAGIC;
    n->lm_seq = m->lm_seq + 1;
#endif
    n->lm_big_blocks = m->lm_big_blocks;
    n->lm_prev = m;
    n->lm_last_alloc = 0;
    n->lm_lifo = l;
    l->lifo_top_mark = n;          /* Of course, bottom mark stays the same */

    return n;
}


int MemLifoPopMark(MemLifoMarkHandle m)
{
    MemLifoMarkHandle n;

#ifdef TWAPI_MEMLIFO_DEBUG
    MEMLIFO_ASSERT(m->lm_magic == MEMLIFO_MARK_MAGIC);
#endif

    n = m->lm_prev;          /* Note n, m may be the same (first mark) */
    MEMLIFO_ASSERT(n);
    MEMLIFO_ASSERT(n->lm_lifo == m->lm_lifo);

#ifdef TWAPI_MEMLIFO_DEBUG
    MEMLIFO_ASSERT(n->lm_seq < m->lm_seq || n == m);
#endif

    if (m->lm_big_blocks != n->lm_big_blocks || m->lm_chunks != n->lm_chunks) {
	MemLifoChunk *c1, *c2, *end;
	MemLifo *l = m->lm_lifo;

	/* 
	 * Free big block lists before freeing chunks since freeing up 
	 * chunks might free up the mark m itself.
	 */
	c1 = m->lm_big_blocks;
	end = n->lm_big_blocks;
	while (c1 != end) {
	    MEMLIFO_ASSERT(c1);
	    c2 = c1->lc_prev;
	    l->lifo_freeFn(c1, l->lifo_allocator_data);
	    c1 = c2;
	}

	/* Free up chunks. Once chunks are freed up, do NOT access m since
	 * it might have been freed as well.
	 */
	c1 = m->lm_chunks;
	end = n->lm_chunks;
	while (c1 != end) {
	    MEMLIFO_ASSERT(c1);
	    c2 = c1->lc_prev;
	    l->lifo_freeFn(c1, l->lifo_allocator_data);
	    c1 = c2;
	}
    }
    n->lm_lifo->lifo_top_mark = n;
    return ERROR_SUCCESS;
}


void *MemLifoPushFrame(MemLifo *l, MemLifoSize sz, MemLifoSize *actual_szP)
{
    void *p;
    MemLifoMarkHandle m, n;
    MemLifoSize total;

    MEMLIFO_ASSERT(l->lifo_magic == MEMLIFO_MAGIC);

    if (sz > MEMLIFO_MAX_ALLOC) {
        if (l->lifo_flags & MEMLIFO_F_PANIC_ON_FAIL)
            Tcl_Panic("Attempt to allocate %" MEMLIFO_SIZE_MODIFIER "u bytes for memlifo", sz);
        return NULL;
    }

    m = l->lifo_top_mark;
    MEMLIFO_ASSERT(m);
    MEMLIFO_ASSERT(ALIGNED(m->lm_freeptr));
    MEMLIFO_ASSERT(ALIGNED(m->lm_chunks->lc_end));

    /* 
     * Optimize for the case that the request can be satisfied from 
     * the current block. Also, we do two compares with 
     * m->lm_free to guard against possible overflows if we simply do 
     * the add and a single compare.
     */
    MEMLIFO_ASSERT(ALIGNED(m->lm_freeptr));
    sz = ROUNDUP(sz);
    total = sz + ROUNDUP(sizeof(*m));        /* Note: ROUNDUP separately */
    p = ADDPTR(m->lm_freeptr, total, void*); /* Potential end of mark */
    if (p > (void*) m->lm_chunks                  /* Ensure no wrap-around! */
        && p <= m->lm_chunks->lc_end) {
        n = (MemLifoMark*)m->lm_freeptr;
        n->lm_chunks = m->lm_chunks;
        n->lm_big_blocks = m->lm_big_blocks;
#ifdef MEMLIFO_DEBUG
        n->lm_magic = MEMLIFO_MARK_MAGIC;
        n->lm_seq = m->lm_seq + 1;
#endif
        n->lm_prev = m;
        n->lm_lifo = l;
        n->lm_last_alloc = ALIGNPTR(n, sizeof(*n), void*);
        /*
         * If actual_szP is non-NULL, caller wants to allocate at least sz
         * but as much as possible without allocating a new chunk
         */
        if (actual_szP) {
            n->lm_freeptr = m->lm_chunks->lc_end;
            *actual_szP = PTRDIFF32(n->lm_freeptr, n->lm_last_alloc);
        } else
            n->lm_freeptr = p;

        l->lifo_top_mark = n;

        return n->lm_last_alloc;
    }

    /* Slow path. Allocate mark, them memory. */
    n = MemLifoPushMark(l);
    if (n) {
        p = MemLifoAlloc(l, sz, actual_szP);
        if (p)
            return p;
        MemLifoPopMark(n);
    }
    if (l->lifo_flags & MEMLIFO_F_PANIC_ON_FAIL)
        Tcl_Panic("Attempt to allocate %" MEMLIFO_SIZE_MODIFIER "u bytes for memlifo", sz);
    return NULL;
}

void *MemLifoExpandLast(MemLifo *l, MemLifoSize incr, int fix)
{
    MemLifoMarkHandle m;
    MemLifoSize old_sz, sz, chunk_sz;
    void *p, *p2;
    char is_big_block;

    m = l->lifo_top_mark;
    p = m->lm_last_alloc;

    if (p == 0) {
        /* Last alloc was a mark, Just allocate new */
        return MemLifoAlloc(l, incr, NULL);
    }

    incr = ROUNDUP(incr);

    /* 
     * Fast path. Allocation can be satisfied in place if the last 
     * allocation was not a big block and there is enough room in the 
     * current chunk
     */
    is_big_block = (p == ADDPTR(m->lm_big_blocks, sizeof(MemLifoChunk), void*));
    if ((!is_big_block) && (PTRDIFF32(m->lm_chunks->lc_end, m->lm_freeptr) >= (int) incr)) {
	m->lm_freeptr = ADDPTR(m->lm_freeptr, incr, void*);
	return p;
    }

    /* If we are not allowed to move the block, not much we can do. */
    if (fix)
        return 0;

    /* Need to allocate new block and copy to it. */
    /* TBD - use HeapRealloc if our default allocator */
    if (is_big_block)
	old_sz = PTRDIFF32(m->lm_big_blocks->lc_end, m->lm_big_blocks) - ROUNDUP(sizeof(MemLifoChunk));
    else
	old_sz = PTRDIFF32(m->lm_freeptr, m->lm_last_alloc);

    MEMLIFO_ASSERT(ROUNDED(old_sz));
    sz = old_sz + incr;

    /* Note so far state of memlifo has not been modified. */

    if (sz > MEMLIFO_MAX_ALLOC)
        return NULL;

    if (is_big_block) {
        MemLifoChunk *c;
        MemLifoSize actual_size;
	/*
	 * Unlink the big block from the big block list. 
	 * TBD - when we call MemLifoAlloc here we have to call it 
	 * with an inconsistent state of m.
	 * Note we do not need to update previous marks since only 
	 * topmost mark could point to allocations after the top mark.
	 */
        chunk_sz = sz + ROUNDUP(sizeof(MemLifoChunk));
        c = (MemLifoChunk *) l->lifo_allocFn(chunk_sz,
                                                  l->lifo_allocator_data,
                                                  &actual_size);
        if (c == NULL) {
            return NULL;
        }
        MEMLIFO_ASSERT(ROUNDED(actual_size));
        
	c->lc_end = ADDPTR(c, actual_size, void*);
        p2 = ADDPTR(c, sizeof(*c), void*);
	CopyMemory(p2, p, old_sz);

	/* Place on the list of big blocks, unlinking previous block */
	c->lc_prev = m->lm_big_blocks->lc_prev;
        l->lifo_freeFn(m->lm_big_blocks, l->lifo_allocator_data);
	m->lm_big_blocks = c;
	/* 
	 * Note we do not modify m->m_freeptr since it still refers to 
	 * the current "mainstream" chunk.
	 */
        m->lm_last_alloc = p2;
	return p2;
    } else {
        /*
         * Allocation was not from a big block. Just allocate new space.
         * Note previous space is not freed. TBD - should we just allocate
         * a big block as above in this case and free up the main chunk space?
         */
        p2 = MemLifoAlloc(l, sz, NULL);
        if (p2 == NULL)
            return NULL;
	CopyMemory(p2, p, old_sz);
	return p2;
    }
}


void * MemLifoShrinkLast(MemLifo *l, 
                              MemLifoSize decr,
                              int fix)
{
    MemLifoSize old_sz;
    char is_big_block;
    MemLifoMarkHandle m;

    m = l->lifo_top_mark;

    if (m->lm_last_alloc == 0)
	return 0;

    is_big_block = (m->lm_last_alloc == ADDPTR(m->lm_big_blocks, sizeof(MemLifoChunk), void*));
    if (!is_big_block) {
	old_sz = PTRDIFF32(m->lm_freeptr, m->lm_last_alloc);
	/* do a size check but ignore if invalid */
        decr = ROUNDDOWN(decr);
	if (decr <= old_sz)
	    m->lm_freeptr = SUBPTR(m->lm_freeptr, decr, void*);
    }
    else {
        /* Big block. Don't bother.
           TBD - may be use HeapReAlloc in default case */
    }
    return m->lm_last_alloc;
}


void *MemLifoResizeLast(MemLifo *l, MemLifoSize new_sz, int fix)
{
    MemLifoSize old_sz;
    char is_big_block;
    MemLifoMarkHandle m;

    m = l->lifo_top_mark;

    if (m->lm_last_alloc == 0)
	return 0;

    is_big_block = (m->lm_last_alloc == ADDPTR(m->lm_big_blocks, sizeof(MemLifoChunk), void*));

    /* 
     * Special fast path when allocation is not a big block and can be 
     * done from current chunk
     */
    new_sz = ROUNDUP(new_sz);
    if (is_big_block) {
	old_sz = PTRDIFF32(m->lm_big_blocks->lc_end, m->lm_big_blocks) - ROUNDUP(sizeof(MemLifoChunk));
    } else {
	old_sz = PTRDIFF32(m->lm_freeptr, m->lm_last_alloc);
	if (new_sz <= old_sz) {
	    m->lm_freeptr = SUBPTR(m->lm_freeptr, old_sz-new_sz, void*);
	    return m->lm_last_alloc;
	}
    }

    return (old_sz >= new_sz ?
	    MemLifoShrinkLast(l, old_sz-new_sz, fix) :
	    MemLifoExpandLast(l, new_sz-old_sz, fix));
}


int MemLifoValidate(MemLifo *l)
{
    MemLifoMark *m;
#ifdef MEMLIFO_DEBUG
    int last_seq;
#endif

    if (l->lifo_magic != MEMLIFO_MAGIC)
        return -1;

    /* First validate underlying allocations */
    if (l->lifo_allocFn == MemLifoDefaultAlloc)
        if (! HeapValidate(l->lifo_allocator_data, 0, NULL))
            return -2;

    /* Some basic validation for marks */
    if (l->lifo_top_mark == NULL || l->lifo_bot_mark == NULL)
        return -3;

    m = l->lifo_top_mark;
#ifdef MEMLIFO_DEBUG
    last_seq = 0;
#endif
    do {
#ifdef MEMLIFO_DEBUG
        if (m->lm_magic != MEMLIFO_MARK_MAGIC)
            return -4;
        if (m->lm_seq != last_seq+1)
            return -5;
        last_seq = m->lm_seq;
#endif
        if (m->lm_lifo != l)
            return -6;
        
        /* last_alloc must be 0 or within m->lm_chunks range */
        if (m->lm_last_alloc) {
            if ((char*) m->lm_last_alloc < (char *) m->lm_chunks)
                return -8;
            if (m->lm_last_alloc >= m->lm_chunks->lc_end) {
                /* last alloc is not in chunk. See if it is a big block */
                if (m->lm_big_blocks == NULL ||
                    (m->lm_last_alloc != ADDPTR(m->lm_big_blocks, ROUNDUP(sizeof
                                                                          (MemLifoChunk)), void*))) {
                    /* Not a big block allocation */
                    return -9;
                }
            }
        }

        if (m->lm_freeptr > m->lm_chunks->lc_end)
            return -10;

        if (m == m->lm_prev) {
            /* Last mark */
            if (m != l->lifo_bot_mark)
                return -7;
            break;
        }
        m = m->lm_prev;
    } while (1);
    

    return 0;
}


int Twapi_MemLifoDump(Tcl_Interp *interp, MemLifo *l)
{
    Tcl_Obj *objs[16];
    MemLifoMark *m;

    objs[0] = STRING_LITERAL_OBJ("allocator_data");
    objs[1] = ObjFromLPVOID(l->lifo_allocator_data);
    objs[2] = STRING_LITERAL_OBJ("allocFn");
    objs[3] = ObjFromLPVOID(l->lifo_allocFn);
    objs[4] = STRING_LITERAL_OBJ("freeFn");
    objs[5] = ObjFromLPVOID(l->lifo_freeFn);
    objs[6] = STRING_LITERAL_OBJ("chunk_size");
    objs[7] = ObjFromDWORD_PTR(l->lifo_chunk_size);
    objs[8] = STRING_LITERAL_OBJ("magic");
    objs[9] = ObjFromLong(l->lifo_magic);
    objs[10] = STRING_LITERAL_OBJ("top_mark");
    objs[11] = ObjFromDWORD_PTR(l->lifo_top_mark);
    objs[12] = STRING_LITERAL_OBJ("bot_mark");
    objs[13] = ObjFromDWORD_PTR(l->lifo_bot_mark);

    objs[14] = STRING_LITERAL_OBJ("marks");
    objs[15] = ObjNewList(0, NULL);

    m = l->lifo_top_mark;
    do {
        Tcl_Obj *mobjs[16];
        mobjs[0] = STRING_LITERAL_OBJ("magic");
        mobjs[1] = ObjFromLong(m->lm_magic);
        mobjs[2] = STRING_LITERAL_OBJ("seq");
        mobjs[3] = ObjFromLong(m->lm_seq);
        mobjs[4] = STRING_LITERAL_OBJ("lifo");
        mobjs[5] = ObjFromOpaque(m->lm_lifo, "MemLifo*");
        mobjs[6] = STRING_LITERAL_OBJ("prev");
        mobjs[7] = ObjFromOpaque(m->lm_prev, "MemLifoMark*");
        mobjs[8] = STRING_LITERAL_OBJ("last_alloc");
        mobjs[9] = ObjFromLPVOID(m->lm_last_alloc);
        mobjs[10] = STRING_LITERAL_OBJ("big_blocks");
        mobjs[11] = ObjFromLPVOID(m->lm_big_blocks);
        mobjs[12] = STRING_LITERAL_OBJ("chunks");
        mobjs[13] = ObjFromLPVOID(m->lm_chunks);
        mobjs[14] = STRING_LITERAL_OBJ("lm_freeptr");
        mobjs[15] = ObjFromDWORD_PTR(m->lm_freeptr);
        ObjAppendElement(interp, objs[15], ObjNewList(ARRAYSIZE(mobjs), mobjs));
        
        if (m == m->lm_prev)
            break;
        m = m->lm_prev;
    } while (1);
    
    return ObjSetResult(interp, ObjNewList(ARRAYSIZE(objs),objs));
}

#if 0
proc mark {l} {return [twapi::Twapi_MemLifoPushMark $l]}
proc unmark m {twapi::Twapi_MemLifoPopMark $m}
proc alloc {l n} {return [::twapi::Twapi_MemLifoAlloc $l $n]}
proc lifo n {twapi::Twapi_MemLifoInit $n}
proc lifodump l {twapi::Twapi_MemLifoDump $l}
proc validate l {twapi::Twapi_MemLifoValidate $l}
proc frame {l n} {twapi::Twapi_MemLifoPushFrame $l $n}
proc unframe l {twapi::Twapi_MemLifoPopFrame $l}

#endif
