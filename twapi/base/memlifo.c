/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

/*
 * Macros for alignment
 */
#define ALIGNMENT sizeof(__int64)
#define ALIGNMASK (~(INT_PTR)(ALIGNMENT-1))
/* Round up to alignment size */
#define ROUNDUP(x_) (( ALIGNMENT - 1 + (INT_PTR)(x_)) & ALIGNMASK)
#define ROUNDED(x_) (ROUNDUP(x_) == (x_))
#define ROUNDDOWN(x_) (ALIGNMASK & (INT_PTR)(x_))

#define ALIGNPTR(base_, offset_, type_) \
    (type_) ROUNDUP((offset_) + (DWORD_PTR)(base_))
#define ALIGNED(p_) (ROUNDED((DWORD_PTR)(p_)))

#define ADDPTR(p_, incr_, type_) \
    ((type_)((incr_) + (char *)(p_)))
#define SUBPTR(p_, decr_, type_) \
    ((type_)(((char *)(p_)) - (decr_)))
#define PTRDIFF(p_, q_) ((char*)(p_) - (char *)(q_))

/* MAX alloc is mainly to catch errors */
#define TWAPI_MEMLIFO_MAX_ALLOC ((size_t) 1000000)

typedef struct _TwapiMemLifoChunk TwapiMemLifoChunk;

/*
Each region is composed of a linked list of contiguous chunks of memory. Each
chunk is prefixed by a descriptor which is also used to link the chunks.
*/
typedef struct _TwapiMemLifoChunk {
    TwapiMemLifoChunk *lc_prev;	/* Pointer to next chunk */
    size_t	lc_size;	/* Size of this chunk.
                                   Do we really need to track this? */
} TwapiMemLifoChunk;
#define CHUNKDSCSZ (ROUNDUP(sizeof(TwapiMemLifoChunk)))
#define TwapiMem_LIFOMAXALLOC	(TwapiMem_MAXALLOC - CHUNKDSCSZ)


/*
A mark keeps current state information about a TwapiMemLifo which can
be used to restore it to a previous state. On initialization, a mark
is internally created to hold the initial state for TwapiMemLifo.
Later, marks are added and deleted through application calls.

The topmost mark (last one allocated) holds the current state of the
TwapiMemLifo.

When chunks are allocated, they are put a single list shared among all
the marks which will point to sublists within the list. When a mark is
"popped", all chunks on the list that are not part of the sublist
pointed to by the previous mark are freed.  "Big block" allocations,
are handled the same way.

Since marks also keeps track of the free space in the current chunk
from which allocations are made, the state of the TwapiMemLifo_t can be
completely restored when popping a mark by simply making the previous
mark the topmost (current) mark.
*/

typedef struct _TwapiMemLifoMark {
    int			lm_magic;	/* Only used in debug mode */
#define TWAPI_MEMLIFO_MARK_MAGIC	0xa0193d4f
    int			lm_seq;		/* Only used in debug mode */
    TwapiMemLifo       *lm_lifo; /* TwapiMemLifo for this mark */
    TwapiMemLifoMark *lm_prev; /* Previous mark */
    void *		lm_last_alloc; /* last allocation from the lifo */
    TwapiMemLifoChunk *lm_big_blocks; /* Pointer to list of large block
                                         allocations. Note marks should
                                         never be allocated from here since
                                         big blocks may be deleted during
                                         some reallocations */
    TwapiMemLifoChunk *lm_chunks; /* Current chunk used for allocation
                                        and the head of the chunk list */
    void  *lm_last_addr;/* Aligned addr 1 beyond usable space in lm_chunkList */
    size_t lm_free;     /* Available space */
} TwapiMemLifoMark;
#define MARKDSCSZ (TwapiMem_ALIGNSIZEUP(sizeof(TwapiMemLifoMarkDsc_t)))

static void *TwapiMemLifoDefaultAlloc(size_t sz, HANDLE heap, size_t *actual)
{
    void *p = HeapAlloc(heap, 0, sz);
    if (actual)
        *actual = HeapSize(heap, 0, p);
    return p;
}

static void TwapiMemLifoDefaultFree(void *p, HANDLE heap)
{
    HeapFree(heap, 0, p);
}


int TwapiMemLifoInit(
    TwapiMemLifo *l,
    void *allocator_data,
    TwapiMemLifoChunkAllocFn *allocFunc,
    TwapiMemLifoChunkFreeFn *freeFunc,
    size_t chunk_sz)
{
    TwapiMemLifoChunk *chunkP;
    TwapiMemLifoMarkHandle m;
    size_t actual_chunk_sz;

    if (allocFunc == 0) {
        allocator_data = HeapCreate(0, 0, 0);
        if (allocator_data == NULL)
            return GetLastError();
	allocFunc = TwapiMemLifoDefaultAlloc;
	freeFunc = TwapiMemLifoDefaultFree;
    } else {
        TWAPI_MEMLIFO_ASSERT(freeFunc);	/* If allocFunc was not 0, freeFunc
					   should not be either */
    }
	
    if (chunk_sz < 1000)
        chunk_sz = 1000;

    /* Allocate a chunk and allocate space for the lifo descriptor from it */
    chunkP = allocFunc(chunk_sz, allocator_data, &actual_chunk_sz);
    if (chunkP == 0)
        return ERROR_OUTOFMEMORY;

    chunkP->lc_prev = NULL;
    chunkP->lc_size = actual_chunk_sz;

    l->lifo_allocator_data = allocator_data;
    l->lifo_allocFn = allocFunc;
    l->lifo_freeFn = freeFunc;
    l->lifo_chunk_size = chunk_sz;
    l->lifo_magic = TWAPI_MEMLIFO_MAGIC;

    /* Allocate mark from chunk itself */
    m = ALIGNPTR(chunkP, sizeof(*chunkP), TwapiMemLifoMarkHandle);

    m->lm_magic = TWAPI_MEMLIFO_MARK_MAGIC;
    m->lm_seq = 1;
    
    m->lm_free = chunk_sz - ROUNDUP(sizeof(*chunkP)) - ROUNDUP(sizeof(*m));
    m->lm_last_addr = ADDPTR(chunkP, chunk_sz, void*);

    m->lm_lifo = l;	/* TBD - do we need this ? */
    m->lm_prev = m;	/* Point back to itself. Effectively will never be
                           popped */
    m->lm_big_blocks = 0;
    m->lm_last_alloc = 0;
    m->lm_chunks = chunkP;

    l->lifo_top_mark = m;
    l->lifo_bot_mark = m;

    return ERROR_SUCCESS;
}

void TwapiMemLifoClose(TwapiMemLifo *l)
{
    l->lifo_magic = 0;

    TWAPI_MEMLIFO_ASSERT(l->lifo_bot_mark->lm_chunks);

    TwapiMemLifoPopMark(l->lifo_bot_mark);

    /* Finally free the chunk containing the bottom mark */
    l->lifo_freeFn(l->lifo_bot_mark->lm_chunks,
                       l->lifo_allocator_data);
    ZeroMemory(l, sizeof(*l));
}


void* TwapiMemLifoAlloc(TwapiMemLifo *l,  size_t sz)
{
    TwapiMemLifoChunk *c;
    TwapiMemLifoMarkHandle m;
    size_t chunk_sz;

    if (sz > TWAPI_MEMLIFO_MAX_ALLOC)
        return NULL;

    /* 
     * NOTE: note that when called from TwapiMemLifoExpandLast(), the 
     * lm_last_alloc field may not be set correctly since that function 
     * may remove the last allocation from the big block list under 
     * some circumstances.
     */

    TWAPI_MEMLIFO_ASSERT(l->lifo_magic == TWAPI_MEMLIFO_MAGIC);
    TWAPI_MEMLIFO_ASSERT(l->lifo_bot_mark);
    
    m = l->lifo_top_mark;
    TWAPI_MEMLIFO_ASSERT(m);
    TWAPI_MEMLIFO_ASSERT(ROUNDUP(m->lm_free) == m->lm_free);
    TWAPI_MEMLIFO_ASSERT(ALIGNPTR(m->lm_last_addr, 0, void *) == m->lm_last_addr);

    sz = ROUNDUP(sz);

    if (sz <= m->lm_free) {
	m->lm_last_alloc = SUBPTR(m->lm_last_addr, m->lm_free, void *);
	m->lm_free -= sz;
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
     * space due to external fragmentation and possibly waste of 
     * global memory handles under Windows.
     *
     * TBD - if using our default windows based heap, try and HeapReAlloc
     * in place.
     */
    /* TBD - make it /2 instead of /8 ? */
    if (m->lm_free < l->lifo_chunk_size/8) {
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

        chunk_sz += ROUNDUP(sizeof(TwapiMemLifoChunk));

	c = l->lifo_allocFn(chunk_sz, l->lifo_allocator_data, &chunk_sz);
	if (c == 0)
	    return 0;

	c->lc_size = chunk_sz;

	c->lc_prev = m->lm_chunks;	/* Place on the list of chunks */
	m->lm_chunks = c;
	
	m->lm_free = chunk_sz - ROUNDUP(sizeof(TwapiMemLifoChunk)) - sz;
        m->lm_last_addr = ADDPTR(c, chunk_sz, void*);
    }
    else {
	/* Allocate a separate big block. */
        size_t actual_size;
        chunk_sz = sz + ROUNDUP(sizeof(TwapiMemLifoChunk));

	c = (TwapiMemLifoChunk *) l->lifo_allocFn(chunk_sz,
                                                  l->lifo_allocator_data,
                                                  &actual_size);
	if (c == 0)
	    return 0;

	c->lc_size = actual_size;

	c->lc_prev = m->lm_big_blocks;	/* Place on the list of big blocks */
	m->lm_big_blocks = c;
	/* 
	 * Note we do not modify m->m_free since it still refers to 
	 * the current "mainstream" chunk.
	 */
    }
    m->lm_last_alloc = ADDPTR(c, ROUNDUP(sizeof(TwapiMemLifoChunk)), void*);
    return m->lm_last_alloc;
}


TwapiMemLifoMarkHandle TwapiMemLifoPushMark(TwapiMemLifo *l)
{
    TwapiMemLifoMarkHandle m;			/* Previous (existing) mark */
    TwapiMemLifoMarkHandle n;			/* New mark */
    TwapiMemLifoChunk *c;
    size_t chunk_sz;
    size_t mark_sz;
    
    m = l->lifo_top_mark;

    /* NOTE - marks must never be allocated from big block list  */

    /*
     * Check for common case first - enough space in current chunk to
     * hold the mark.
     */
    mark_sz = ROUNDUP(sizeof(TwapiMemLifoMark));
    if (mark_sz <= m->lm_free) {
	TWAPI_MEMLIFO_ASSERT(ROUNDED(m->lm_free));
	TWAPI_MEMLIFO_ASSERT(ALIGNED(m->lm_lastAddr));
	n = SUBPTR(m->lm_last_addr, m->lm_free, TwapiMemLifoMarkHandle);
	n->lm_free = m->lm_free - mark_sz;
	n->lm_last_addr = m->lm_last_addr;
	n->lm_chunks = m->lm_chunks;
    }
    else {
	/* 
	 * No room in current chunk. Have to allocate a new chunk. Note 
	 * we do not use TwapiMemLifoAlloc to allocate the mark since that 
	 * would change the state of the previous mark.
	 */
	c = l->lifo_allocFn(l->lifo_chunk_size,
                            l->lifo_allocator_data,
                            &chunk_sz);
	if (c == 0)
	    return 0;
	
	c->lc_size = chunk_sz;
	
	/* 
	 * Place on the list of chunks. Note however, that we do NOT 
	 * modify m->lm_chunkList since that should hold the original lifo 
	 * state. We'll put this chunk on the list headed by the new mark.
	 */
	c->lc_prev = m->lm_chunks;	/* Place on the list of chunks */

	n = ADDPTR(c, ROUNDUP(sizeof(*c)), TwapiMemLifoMarkHandle);
	n->lm_chunks = c;
	
	n->lm_last_addr = ADDPTR(c, chunk_sz, void*);
	n->lm_free = chunk_sz - ROUNDUP(sizeof(*c)) - ROUNDUP(sizeof(*n));
    }

#ifdef TWAPI_MEMLIFO_DEBUG
    n->lm_magic = TWAPI_MEMLIFO_MARK_MAGIC;
    n->lm_seq = m->lm_seq + 1;
#endif
    n->lm_big_blocks = m->lm_big_blocks;
    n->lm_prev = m;
    n->lm_last_alloc = 0;
    n->lm_lifo = l;
    l->lifo_top_mark = n;          /* Of course, bottom mark stays the same */

    return n;
}


int TwapiMemLifoPopMark(TwapiMemLifoMarkHandle m)
{
    TwapiMemLifoMarkHandle n;


#ifdef TWAPI_MEMLIF_ODEBUG
    TWAPI_MEMLIFO_ASSERT(m->lm_magic == TWAPI_MEMLIFO_MARK_MAGIC);
#endif

    n = m->lm_prev;          /* Note n, m may be the same (first mark) */
    TWAPI_MEMLIFO_ASSERT(n);
    TWAPI_MEMLIFO_ASSERT(n->lm_lifo == m->lm_lifo);

#ifdef TWAPI_MEMLIF_ODEBUG
    TWAPI_MEMLIFO_ASSERT(n->lm_seq < m->lm_seq || n == m);
#endif

    /*
     * Do a quick check to see if any blocks or chunks need to be freed up.
     * 
     */
    if (m->lm_big_blocks == n->lm_big_blocks && m->lm_chunks == n->lm_chunks) {
        /* Note it is possible that m == n when it is the bottommost mark */
        n->lm_lifo->lifo_top_mark = n;
    } else {
	TwapiMemLifoChunk *c1, *c2, *end;
	TwapiMemLifo *l = m->lm_lifo;

	/* 
	 * Free big block lists before freeing chunks since freeing up 
	 * chunks might free up the mark m itself.
	 */
	c1 = m->lm_big_blocks;
	end = n->lm_big_blocks;
	while (c1 != end) {
	    TWAPI_MEMLIFO_ASSERT(c1);
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
	    TWAPI_MEMLIFO_ASSERT(c1);
	    c2 = c1->lc_prev;
	    l->lifo_freeFn(c1, l->lifo_allocator_data);
	    c1 = c2;
	}
    }

    return ERROR_SUCCESS;
}


void *TwapiMemLifoPushFrame(TwapiMemLifo *l, size_t sz)
{
    void *p;
    TwapiMemLifoMarkHandle m, n;
       
    TWAPI_MEMLIFO_ASSERT(l->lifo_magic == TWAPI_MEMLIFO_MAGIC);
    
    m = l->lifo_top_mark;
    TWAPI_MEMLIFO_ASSERT(m);
    TWAPI_MEMLIFO_ASSERT(ROUNDED(m->lm_free));
    TWAPI_MEMLIFO_ASSERT(ALIGNED(m->lm_last_addr));

    /* 
     * Optimize for the case that the request can be satisfied from 
     * the current block. Also, we do two compares with 
     * m->lm_free to guard against possible overflows if we simply do 
     * the add and a single compare.
     */
    TWAPI_MEMLIFO_ASSERT(ROUNDED(m->lm_free));
    sz = ROUNDUP(sz);
    if (sz <= m->lm_free) {
	size_t total = sz + ROUNDUP(sizeof(*m));
	if (total <= m->lm_free) {
	    n = SUBPTR(m->lm_last_addr, m->lm_free, TwapiMemLifoMarkHandle);
	    n->lm_last_addr = m->lm_last_addr;
	    n->lm_chunks = m->lm_chunks;
	    n->lm_big_blocks = m->lm_big_blocks;
	    n->lm_free = m->lm_free - total;
#ifdef TWAPI_MEMLIFO_DEBUG
	    n->lm_magic = TWAPI_MEMLIFO_MARK_MAGIC;
	    n->lm_seq = m->lm_seq + 1;
#endif
	    n->lm_prev = m;
	    n->lm_lifo = l;
	    l->lifo_top_mark = n;
	    n->lm_last_alloc = ADDPTR(n, sizeof(*n), void*);
	    return n->lm_last_alloc;
	}
    }

    /* Slow path. Allocate mark, them memory. */
    n = TwapiMemLifoPushMark(l);
    if (n == 0)
	return 0;
    p = TwapiMemLifoAlloc(l, sz);
    if (p == 0)
	TwapiMemLifoPopMark(n);
    return p;
}

int TwapiMemLifoPopFrame(TwapiMemLifo *l)
{
    TwapiMemLifoMarkHandle m = l->lifo_top_mark;
    TwapiMemLifoMarkHandle n;

    TWAPI_MEMLIFO_ASSERT(m->lm_magic == TWAPI_MEMLIFO_MARK_MAGIC);

    n = m->lm_prev;             /* m == n => first mark */
    TWAPI_MEMLIFO_ASSERT(n);
    TWAPI_MEMLIFO_ASSERT(n->lm_lifo == m->lm_lifo);
    TWAPI_MEMLIFO_ASSERT((n->lm_seq < m->lm_seq) || (n == m));

    /* Do a fastpath check and if that does not apply, pass on request */
    /* TBD - use a bitflag that is set at big block or chunk 
     *       allocation time and check for this bit instead of doing 
     *       two compares. (Will that be more efficient? Given that we 
     *       will have to init that field?
     */
    if (m->lm_big_blocks == n->lm_big_blocks && m->lm_chunks == n->lm_chunks) {
        l->lifo_top_mark = n;
	return ERROR_SUCCESS;
    }
    else
	return TwapiMemLifoPopMark(m);
}


void *TwapiMemLifoExpandLast(TwapiMemLifo *l, size_t incr, int fix)
{
    TwapiMemLifoChunk *c;
    TwapiMemLifoMarkHandle m;
    size_t old_sz, sz, chunk_sz;
    void *p, *p2;
    char is_big_block;

    m = l->lifo_top_mark;
    p = m->lm_last_alloc;

    if (p == 0) {
        /* Last alloc was a mark, Just allocate new */
        return TwapiMemLifoAlloc(l, incr);
    }

    incr = ROUNDUP(incr);

    /* 
     * Fast path. Allocation can be satisfied in place if the last 
     * allocation was not a big block and there is enough room in the 
     * current chunk
     */
    is_big_block = (p == ADDPTR(m->lm_big_blocks, sizeof(TwapiMemLifoChunk), void*));
    if ((!is_big_block) && (m->lm_free >= incr)) {
	m->lm_free -= incr;
	return p;
    }

    /* If we are not allowed to move the block, not much we can do. */
    if (fix)
        return 0;

    /* Need to allocate new block and copy to it. */
    if (is_big_block) {
        /* TBD - use HeapRealloc if our default allocator */
	c = m->lm_big_blocks;
	old_sz = c->lc_size - ROUNDUP(sizeof(*c));
    }
    else {
	old_sz = (size_t) PTRDIFF(m->lm_last_addr, m->lm_last_alloc);
	TWAPI_MEMLIFO_ASSERT(old_sz >= m->lm_free); /* Actual alloc+free space*/
	old_sz -= m->lm_free;
    }

    sz = old_sz + incr;
    TWAPI_MEMLIFO_ASSERT(ROUNDED(sz));

    /* Note so far state of memlifo has not been modified. */

    if (sz > TWAPI_MEMLIFO_MAX_ALLOC)
        return NULL;
 
    if (is_big_block) {
        size_t actual_size;
	/*
	 * Unlink the big block from the big block list. 
	 * TBD - when we call TwapiMemLifoAlloc here we have to call it 
	 * with an inconsistent state of m.
	 * Note we do not need to update previous marks since only 
	 * topmost mark could point to allocations after the top mark.
	 */
        chunk_sz = sz + ROUNDUP(sizeof(TwapiMemLifoChunk));
        c = (TwapiMemLifoChunk *) l->lifo_allocFn(chunk_sz,
                                                  l->lifo_allocator_data,
                                                  &actual_size);
        if (c == NULL)
            return NULL;
        
	c->lc_size = actual_size;
        p2 = ADDPTR(c, sizeof(*c), void*);
	CopyMemory(p2, p, old_sz);

	/* Place on the list of big blocks, unlinking previous block */
	c->lc_prev = m->lm_big_blocks->lc_prev;
        l->lifo_freeFn(m->lm_big_blocks, l->lifo_allocator_data);
	m->lm_big_blocks = c;
	/* 
	 * Note we do not modify m->m_free since it still refers to 
	 * the current "mainstream" chunk.
	 */
        m->lm_last_alloc = p2;
	return p2;

    } else {
        /* Allocation was not from a big block. Note last alloc is not freed */
        p2 = TwapiMemLifoAlloc(l, sz);
        if (p2 == NULL)
            return NULL;
	CopyMemory(p2, p, old_sz);
	return p2;
    }
}


void * TwapiMemLifoShrinkLast(TwapiMemLifo *l, 
                              size_t decr,
                              int fix)
{
    size_t old_sz;
    void *p;
    char is_big_block;
    TwapiMemLifoMarkHandle m;

    m = l->lifo_top_mark;
    p = m->lm_last_alloc;

    if (p == 0)
	return 0;

    is_big_block = (p == ADDPTR(m->lm_big_blocks, sizeof(TwapiMemLifoChunk), void*));
    if (!is_big_block) {
	old_sz = PTRDIFF(m->lm_last_addr, p);
	TWAPI_MEMLIFO_ASSERT(old_sz >= m->lm_free);

	old_sz -= m->lm_free;

	/* do a size check but ignore if invalid */
	if (decr <= old_sz)
	    m->lm_free += ROUNDDOWN(decr);
	return p;
    }
    else {
        /* Big block. Don't bother. TBD - may be use HeapReAlloc in default case
 */
        return p;
    }
}


void *TwapiMemLifoResizeLast(TwapiMemLifo *l, size_t new_sz, int fix)
{
    size_t old_sz;
    void *p;
    char is_big_block;
    TwapiMemLifoMarkHandle m;

    m = l->lifo_top_mark;
    p = m->lm_last_alloc;

    if (p == 0)
	return 0;

    is_big_block = (p == ADDPTR(m->lm_big_blocks, sizeof(TwapiMemLifoChunk), void*));

    /* 
     * Special fast path when allocation is not a big block and can be 
     * done from current chunk
     */
    new_sz = ROUNDUP(new_sz);
    if (is_big_block) {
	old_sz = m->lm_big_blocks->lc_size - ROUNDUP(sizeof(TwapiMemLifoChunk));
    } else {
	old_sz = PTRDIFF(m->lm_last_addr, p);
	TWAPI_MEMLIFO_ASSERT(old_sz >= m->lm_free);
	old_sz -= m->lm_free;
	if (new_sz <= old_sz) {
	    m->lm_free += (old_sz - new_sz);
	    return p;
	}
    }

    return (old_sz > new_sz ?
	    TwapiMemLifoShrinkLast(l, old_sz-new_sz, fix) :
	    TwapiMemLifoExpandLast(l, new_sz-old_sz, fix));
}


int TwapiMemLifoValidate(TwapiMemLifo *l)
{
    TwapiMemLifoMark *m;
#ifdef TWAPI_MEMLIFO_DEBUG
    int last_seq;
#endif

    if (l->lifo_magic != TWAPI_MEMLIFO_MAGIC)
        return -1;

    /* First validate underlying allocations */
    if (l->lifo_allocFn == TwapiMemLifoDefaultAlloc)
        if (! HeapValidate(l->lifo_allocator_data, 0, NULL))
            return -2;

    /* Some basic validation for marks */
    if (l->lifo_top_mark == NULL || l->lifo_bot_mark == NULL)
        return -3;

    m = l->lifo_top_mark;
#ifdef TWAPI_MEMLIFO_DEBUG
    last_seq = 0;
#endif
    do {
#ifdef TWAPI_MEMLIFO_DEBUG
        if (m->lm_magic != TWAPI_MEMLIFO_MARK_MAGIC)
            return -4;
        if (m->lm_seq != last_seq+1)
            return -5;
        last_seq = m->lm_seq;
#endif
        if (m->lm_lifo != l)
            return -6;
        
        if (m->lm_last_addr != ADDPTR(m->lm_chunks, m->lm_chunks->lc_size, void*))
            return -10;

        /* last_alloc must be 0 or within m->lm_chunks range */
        if (m->lm_last_alloc) {
            if ((char*) m->lm_last_alloc < (char *) m->lm_chunks)
                return -8;
            if (m->lm_last_alloc >= m->lm_last_addr) {
                /* last alloc is not in chunk. See if it is a big block */
                if (m->lm_big_blocks == NULL ||
                    (m->lm_last_alloc != ADDPTR(m->lm_big_blocks, ROUNDUP(sizeof
                                                                          (TwapiMemLifoChunk)), void*))) {
                    /* Not a big block allocation */
                    return -9;
                }
            }
        }

        if (m->lm_free > (size_t) PTRDIFF(m->lm_last_addr, m->lm_chunks))
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

int Twapi_MemLifoDump(TwapiInterpContext *ticP, TwapiMemLifo *l)
{
    Tcl_Obj *objs[16];
    TwapiMemLifoMark *m;

    objs[0] = STRING_LITERAL_OBJ("allocator_data");
    objs[1] = ObjFromLPVOID(l->lifo_allocator_data);
    objs[2] = STRING_LITERAL_OBJ("allocFn");
    objs[3] = ObjFromLPVOID(l->lifo_allocFn);
    objs[4] = STRING_LITERAL_OBJ("freeFn");
    objs[5] = ObjFromLPVOID(l->lifo_freeFn);
    objs[6] = STRING_LITERAL_OBJ("chunk_size");
    objs[7] = ObjFromDWORD_PTR(l->lifo_chunk_size);
    objs[8] = STRING_LITERAL_OBJ("magic");
    objs[9] = Tcl_NewLongObj(l->lifo_magic);
    objs[10] = STRING_LITERAL_OBJ("top_mark");
    objs[11] = ObjFromDWORD_PTR(l->lifo_top_mark);
    objs[12] = STRING_LITERAL_OBJ("bot_mark");
    objs[13] = ObjFromDWORD_PTR(l->lifo_bot_mark);

    objs[14] = STRING_LITERAL_OBJ("marks");
    objs[15] = Tcl_NewListObj(0, NULL);

    m = l->lifo_top_mark;
    do {
        Tcl_Obj *mobjs[18];
        mobjs[0] = STRING_LITERAL_OBJ("magic");
        mobjs[1] = Tcl_NewLongObj(m->lm_magic);
        mobjs[2] = STRING_LITERAL_OBJ("seq");
        mobjs[3] = Tcl_NewLongObj(m->lm_seq);
        mobjs[4] = STRING_LITERAL_OBJ("lifo");
        mobjs[5] = ObjFromOpaque(m->lm_lifo, "TwapiMemLifo*");
        mobjs[6] = STRING_LITERAL_OBJ("prev");
        mobjs[7] = ObjFromOpaque(m->lm_prev, "TwapiMemLifoMark*");
        mobjs[8] = STRING_LITERAL_OBJ("last_alloc");
        mobjs[9] = ObjFromLPVOID(m->lm_last_alloc);
        mobjs[10] = STRING_LITERAL_OBJ("big_blocks");
        mobjs[11] = ObjFromLPVOID(m->lm_big_blocks);
        mobjs[12] = STRING_LITERAL_OBJ("chunks");
        mobjs[13] = ObjFromLPVOID(m->lm_chunks);
        mobjs[14] = STRING_LITERAL_OBJ("last_addr");
        mobjs[15] = ObjFromLPVOID(m->lm_last_addr);
        mobjs[16] = STRING_LITERAL_OBJ("lm_free");
        mobjs[17] = ObjFromDWORD_PTR(m->lm_free);
        Tcl_ListObjAppendElement(ticP->interp, objs[15], Tcl_NewListObj(ARRAYSIZE(mobjs), mobjs));
        
        if (m == m->lm_prev)
            break;
        m = m->lm_prev;
    } while (1);
    
    Tcl_SetObjResult(ticP->interp, Tcl_NewListObj(ARRAYSIZE(objs),objs));
    return TCL_OK;
}

#if 0
proc mark {l} {return [twapi::Twapi_MemLifoPushMark $l]}
proc unmark m {twapi::Twapi_MemLifoPopMark $m}
proc alloc {l n} {return [::twapi::Twapi_MemLifoAlloc $l $n]}
proc lifo n {twapi::Twapi_MemLifoInit $n}
proc lifodump l {twapi::Twapi_MemLifoDump $l}
#endif
