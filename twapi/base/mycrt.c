/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#if defined(TWAPI_MINIMIZE_CRT) || defined(TWAPI_REPLACE_CRT)

int _fltused;

void * __cdecl malloc(size_t sz)
{
    return TwapiAlloc(sz);
}

void * __cdecl calloc(size_t nelems, size_t elem_sz)
{
    return TwapiAllocZero(nelems*elem_sz);
}

void __cdecl free(void *p)
{
    TwapiFree(p);
}

char *  __cdecl strncpy(char *dst, const char *src, size_t sz)
{
    StringCchCopyNExA(dst, sz, src, sz, NULL, NULL, STRSAFE_FILL_BEHIND_NULL);
    return dst;
}


#endif /* TWAPI_MINIMIZE_CRT */
