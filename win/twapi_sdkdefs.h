#ifndef TWAPI_SDKDEFS_H
#define TWAPI_SDKDEFS_H

/* 
 * Depending on compiler / platform, ZeroMemory translates a call to memset.
 * If we want to avoid the C rtl, then define appropriately
 */
#ifdef __GNUC__
# define TwapiZeroMemory(p_, count_) memset((p_), 0, (count_))
# undef SecureZeroMemory
# define SecureZeroMemory(p_, count_) memset((p_), 0, (count_))
#else
#ifdef TWAPI_REPLACE_CRT
# ifdef _M_AMD64
#  define TwapiZeroMemory(p_, count_) RtlSecureZeroMemory((p_), (count_))
# else
#  define TwapiZeroMemory(p_, count_) ZeroMemory((p_), (count_))
# endif
#else
# define TwapiZeroMemory(p_, count_) memset((p_), 0, (count_))
#endif
#endif

#endif