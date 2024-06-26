#ifndef TWAPI_H
#define TWAPI_H

/*
 * Copyright (c) 2010-2024, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#if _WIN32_WINNT < 0x0601
#error _WIN32_WINNT too low
#endif

#ifndef _UNICODE
#error -D_UNICODE is not defined on compile command.
#endif

/* Enable prototype-less extern functions warnings even at warning level 1 */
#ifdef _MSC_VER
#pragma warning (1 : 13)
#endif

/*
 * The following is along the lines of the Tcl model for exporting
 * functions.
 *
 * TWAPI_STATIC_BUILD - the TWAPI build should define this to build
 * a static library version of TWAPI. Clients should define this when
 * linking to the static library version of TWAPI.
 *
 * twapi_base_BUILD - the TWAPI core build and ONLY the TWAPI core build
 * should define this, both for static as well dll builds. Clients should never
 * define it for their builds, nor should other twapi components. (lower
 * case "twapi_base" because it is derived from the build system)
 *
 * USE_TWAPI_STUBS - for future use should not be currently defined.
 */

#if defined(TWAPI_STATIC_BUILD)
# define TWAPI_EXPORT
# define TWAPI_IMPORT
#else
# if defined(TWAPI_SINGLE_MODULE)
// We still want to export but modules are built-in so TWAPI_IMPORT is no-op
#  define TWAPI_EXPORT __declspec(dllexport)
#  define TWAPI_IMPORT
# else
#  define TWAPI_EXPORT __declspec(dllexport)
#  define TWAPI_IMPORT __declspec(dllimport)
# endif
#endif

#ifdef twapi_base_BUILD
#   define TWAPI_STORAGE_CLASS TWAPI_EXPORT
#else
#   ifdef USE_TWAPI_STUBS
#      error USE_TWAPI_STUBS not implemented yet.
#      define TWAPI_STORAGE_CLASS
#   else
#      define TWAPI_STORAGE_CLASS TWAPI_IMPORT
#   endif
#endif

#ifdef __cplusplus
#   define TWAPI_EXTERN extern "C" TWAPI_STORAGE_CLASS
#   ifdef __GNUC__
#       define TWAPI_EXTERN_VA extern "C" __cdecl
#   else
#       define TWAPI_EXTERN_VA extern "C" TWAPI_STORAGE_CLASS
#   endif
#else
#   define TWAPI_EXTERN extern TWAPI_STORAGE_CLASS
#   ifdef __GNUC__
#       define TWAPI_EXTERN_VA extern __cdecl
#   else
#       define TWAPI_EXTERN_VA extern TWAPI_STORAGE_CLASS
#   endif
#endif

/* TBD - how to define TWAPI_EXTERN */
#undef TWAPI_EXTERN
#undef TWAPI_EXTERN_VA
#undef TWAPI_EXPORT
#undef TWAPI_IMPORT
#define TWAPI_EXTERN
#define TWAPI_EXTERN_VA
#define TWAPI_EXPORT
#define TWAPI_IMPORT

/*
 * TWAPI_INLINE should be used for defining inline functions that are
 * used across multiple files. TWAPI_STATIC_INLINE should be used for
 * functions that are intended to be used within a single file.
 */
#ifndef TWAPI_INLINE
# ifdef _MSC_VER
#  define TWAPI_INLINE __inline  /* Because VC++ 6 only accepts "inline" in C++  */
#  define TWAPI_STATIC_INLINE static __inline
# elif __GNUC__ && !__GNUC_STDC_INLINE__
#  error foo
#  define TWAPI_INLINE extern inline
#  define TWAPI_STATIC_INLINE static inline
# else
#  define TWAPI_INLINE static inline
#  define TWAPI_STATIC_INLINE static inline
# endif
#endif


/* Define macros that affect structure sizes before any includes */
#define CRYPT_DECRYPT_MESSAGE_PARA_HAS_EXTRA_FIELDS
#define CRYPT_SIGN_MESSAGE_PARA_HAS_CMS_FIELDS

#include <stddef.h>
#include <stdlib.h>
#include <stdarg.h>
#include <assert.h>
#if _MSC_VER >= 1300
/* TBD - functions deprecated in newer platforms have no XP equivalent so
 * we have to live with them for now.
 */
#define _WINSOCK_DEPRECATED_NO_WARNINGS
#endif

#include <winsock2.h>
#include <windows.h>

//#define _WSPIAPI_COUNTOF // Needed to use VC++ 6.0 with Platform SDK 2003 SP1
#ifndef ARRAYSIZE
/* Older SDK's do not define this */
#define ARRAYSIZE(A) RTL_NUMBER_OF(A)
#endif
#include <ws2tcpip.h>
#include <psapi.h>
#include <sddl.h>
#include <lmerr.h>
#include <lm.h>
#include <limits.h>
#include <errno.h>
#include <lmat.h>
#include <lm.h>
#include <sddl.h>        /* For SECURITY_DESCRIPTOR <-> string conversions */
#include <aclapi.h>
#include <winnetwk.h>
#include <iphlpapi.h>
#include <objidl.h>
#include <shlobj.h>  /* Need for ITEMIDLIST */
#include <shlwapi.h> /* Need for DLLVERSIONINFO */
# include <ntsecapi.h>
#include <wtsapi32.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <intshcut.h>
#include <dispex.h>
#include <ocidl.h>
#include <dsgetdc.h>
#include <powrprof.h>
#if _MSC_VER <= 1400
/* Not present in newer compiler/sdks as it is subsumed by winuser.h */
# include <winable.h>
#endif
#define SECURITY_WIN32 1
#include <security.h>
#include <userenv.h>
#include <wmistr.h>             /* Needed for WNODE_HEADER for evntrace */
#include <wincred.h>

#include "tcl.h"

#ifndef TCL_SIZE_MAX
# define Tcl_GetSizeIntFromObj Tcl_GetIntFromObj
# define TCL_SIZE_MAX      INT_MAX
# define TCL_SIZE_MODIFIER ""
# ifndef Tcl_Size
typedef int Tcl_Size;
# endif
#endif

#include "twapi_sdkdefs.h"
#include "twapi_ddkdefs.h"
#include "zlist.h"
#include "memlifo.h"

#if 0
// Do not use for now as it pulls in C RTL _vsnprintf AND docs claim
// needs XP *SP2*.
// Note has to be included after all headers
//#define STRSAFE_LIB             /* Use lib, not inline strsafe functions */
//#include <strsafe.h>
#endif

/* Make wide char version of a string literal macro */
#define WLITERAL2(x) L ## x
#define WLITERAL(x) WLITERAL2(x)

typedef DWORD WIN32_ERROR;
typedef int TCL_RESULT;

#define TWAPI_TCL_NAMESPACE "::twapi"
#define TWAPI_SCRIPT_RESOURCE_TYPE "tclscript"
#define TWAPI_SCRIPT_RESOURCE_TYPE_LZMA "tclscriptlzma"
#define TWAPI_SETTINGS_VAR  TWAPI_TCL_NAMESPACE "::settings"
#define TWAPI_LOG_VAR TWAPI_TCL_NAMESPACE "::log_messages"
#define TWAPI_LOG_CONFIG_VAR TWAPI_TCL_NAMESPACE "::log_config"

/* Note:
 * when calling that pkg_ must be title-ized as documented in Tcl_StaticPackage
 */
/* TBD - in 8.6 Tcl_StaticPackage is no longer defined so we call it
   through the stubs structure. Is this the right thing to do ?
   TBD - if twapi gets loaded into multiple interps, is it ok to pass
   NULL to the function the second time it is called?
*/
#define TWAPI_INIT_MODULE(ip_, initfn_)                                        \
    do {                                                                       \
        TCL_RESULT ret = initfn_(ip_);                                         \
        if (ret != TCL_OK)                                                     \
            return ret;                                                        \
    } while (0)

#define MAKESTRINGLITERAL(s_) # s_
 /*
  * Stringifying special CPP symbols (__LINE__) needs another level of macro
  */
#define MAKESTRINGLITERAL2(s_) MAKESTRINGLITERAL(s_)

#if TWAPI_ENABLE_ASSERT
#  if TWAPI_ENABLE_ASSERT == 1
#    define TWAPI_ASSERT(bool_) (void)( (bool_) || (Tcl_Panic("Assertion (%s) failed at line %d in file %s.", #bool_, __LINE__, __FILE__), 0) )
#  elif TWAPI_ENABLE_ASSERT == 2
#    define TWAPI_ASSERT(bool_) (void)( (bool_) || (DebugOutput("Assertion (" #bool_ ") failed at line " MAKESTRINGLITERAL2(__LINE__) " in file " __FILE__ "\n"), 0) )
#  elif TWAPI_ENABLE_ASSERT == 3
#    define TWAPI_ASSERT(bool_) do { if (! (bool_)) { __asm int 3 } } while (0)
#  else
#    error Invalid value for TWAPI_ENABLE_ASSERT
#  endif
#else
#define TWAPI_ASSERT(bool_) ((void) 0)
#endif

TWAPI_EXTERN void TwapiUnreachablePanic(void);
#define UNREACHABLE do { \
        TwapiUnreachablePanic(); \
    } while (0)

#if TWAPI_ENABLE_LOG
#define TWAPI_OBJ_LOG_IF(cond_, interp_, objp_) if (cond_) TWAPI_OBJ_LOG(interp_, objp_)
#define TWAPI_OBJ_LOG(interp_, objp_) Twapi_AppendObjLog(interp, objp_)
#define TWAPI_LOG(interp_, s_) Twapi_AppendLog(interp, L ## s_)
#define TWAPI_LOG_IF(cond_, interp_, s_) if (cond_) TWAPI_LOG(interp_, s_)
#define TWAPI_LOG_BLOCK(cond_) if (cond_)
#else
#define TWAPI_OBJ_LOG_IF(cond_, interp_, objp_) ((void) 0)
#define TWAPI_OBJ_LOG(interp_, objp_) ((void) 0)
#define TWAPI_LOG(interp_, s_) ((void) 0)
#define TWAPI_LOG_IF(cond_, interp_, s_) ((void) 0)
#define TWAPI_LOG_BLOCK(cond_) if (0)
#endif

/*
 * Macro to create a stub to load and return a pointer to a function
 * fnname in dll dllname
 * Example:
 *     MAKE_DYNLOAD_FUNC(ConvertSidToStringSidA, advapi32, FARPROC)
 * will define a function Twapi_GetProc_ConvertSidToStringSidA
 * which can be called to return a pointer to ConvertSidToStringSidA
 * in DLL advapi32.dll or NULL if the function does not exist in the dll
 *
 * Then it can be called as:
 *      FARPROC func = Twapi_GetProc_ConvertSidToStringSidA();
 *      if (func) { (*func)(parameters...);}
 *      else { ... function not present, do whatever... }
 *
 * Note you can pass any function prototype typedef instead of FARPROC
 * for better type safety.
 */

#define MAKE_DYNLOAD_FUNC(fnname, dllname, fntype)       \
    fntype Twapi_GetProc_ ## fnname (void)               \
    { \
        static HINSTANCE dllname ## _H; \
        static fntype  fnname ## dllname ## _F;       \
        static int load_attempted;                                     \
        if (load_attempted) return fnname ## dllname ## _F;            \
        if ((fnname ## dllname ## _F) == NULL) { \
            if ((dllname ## _H) == NULL) { \
                dllname ## _H = LoadLibraryA(#dllname ".dll"); \
            } \
            if (dllname ## _H) { \
                fnname ## dllname ## _F = \
                    (fntype) GetProcAddress( dllname ## _H, #fnname); \
                if ((fnname ## dllname ## _F) == NULL) {               \
                    FreeLibrary(dllname ## _H);                        \
                    dllname ## _H = NULL;                              \
            } \
        } \
        }                                                              \
        load_attempted = 1;                                            \
        return fnname ## dllname ## _F; \
    }

/* Like MAKE_DYNLOAD_FUNC but function name includes dll. Useful in
 * cases where function may be in one of several different DLL's
 */
#define MAKE_DYNLOAD_FUNC2(fnname, dllname) \
    FARPROC Twapi_GetProc_ ## fnname ##_ ## dllname (void) \
    { \
        static HINSTANCE dllname ## _H; \
        static FARPROC   fnname ## dllname ## _F; \
        static int load_attempted; \
        if (load_attempted) return fnname ## dllname ## _F; \
        if ((fnname ## dllname ## _F) == NULL) { \
            if ((dllname ## _H) == NULL) { \
                dllname ## _H = LoadLibraryA(#dllname ".dll"); \
            } \
            if (dllname ## _H) { \
                fnname ## dllname ## _F = \
                    (FARPROC) GetProcAddress( dllname ## _H, #fnname); \
                if ((fnname ## dllname ## _F) == NULL) {               \
                    FreeLibrary(dllname ## _H);                        \
                    dllname ## _H = NULL;                              \
                }                                                      \
            } \
        } \
        load_attempted = 1; \
        return fnname ## dllname ## _F; \
    }

#define MAKE_DYNLOAD_FUNC_ORDINAL(ord, dllname) \
    FARPROC Twapi_GetProc_ ## dllname ## _ ## ord (void) \
    { \
        static HINSTANCE dllname ## _H; \
        static FARPROC   ord_ ## ord ## dllname ## _F; \
        static int load_attempted; \
        if (load_attempted) return ord_ ## ord ## dllname ## _F; \
        if ((ord_ ## ord ## dllname ## _F) == NULL) { \
            if ((dllname ## _H) == NULL) { \
                dllname ## _H = LoadLibraryA(#dllname ".dll"); \
            } \
            if (dllname ## _H) { \
                ord_ ## ord ## dllname ## _F = \
                    (FARPROC) GetProcAddress( dllname ## _H, (LPCSTR) ord); \
                if ((ord_ ## ord ## dllname ## _F) == NULL) {          \
                    FreeLibrary(dllname ## _H);                        \
                    dllname ## _H = NULL;                              \
                }                                                      \
            } \
        } \
        load_attempted = 1; \
        return ord_ ## ord ## dllname ## _F; \
    }

/* 64 bittedness needs a BOOL version of the FARPROC def */
typedef BOOL (WINAPI *FARPROC_BOOL)();

/*
 * Macros for alignment
 */
#define ALIGNMENT sizeof(__int64)
#define ALIGNMASK (~(INT_PTR)(ALIGNMENT-1))
/* Round up to alignment size */
#define ROUNDUP(x_) (( ALIGNMENT - 1 + (x_)) & ALIGNMASK)
#define ROUNDED(x_) (ROUNDUP(x_) == (x_))
#define ROUNDDOWN(x_) (ALIGNMASK & (x_))

/* Note diff between ALIGNPTR and ADDPTR is that former aligns the pointer */
#define ALIGNPTR(base_, offset_, type_) \
    (type_) ROUNDUP((offset_) + (DWORD_PTR)(base_))
#define ADDPTR(p_, incr_, type_) \
    ((type_)((incr_) + (char *)(p_)))
#define SUBPTR(p_, decr_, type_) \
    ((type_)(((char *)(p_)) - (decr_)))
#define ALIGNED(p_) (ROUNDED((DWORD_PTR)(p_)))

/*
 * Pointer diff assuming difference fits in 32 bits. That should be always
 * true even on 64-bit systems because of our limits on alloc size
 */
#define PTRDIFF32(p_, q_) ((int)((char*)(p_) - (char *)(q_)))

/*
 * Support for one-time initialization
 */
typedef volatile LONG TwapiOneTimeInitState;
#define TWAPI_INITSTATE_NOT_DONE    0
#define TWAPI_INITSTATE_IN_PROGRESS 1
#define TWAPI_INITSTATE_DONE        2
#define TWAPI_INITSTATE_ERROR       3

/*************************
 * Error code definitions
 *************************/

/*
 * String to use as first element of a errorCode
 */
#define TWAPI_WIN32_ERRORCODE_TOKEN "TWAPI_WIN32"  /* For Win32 errors */
#define TWAPI_ERRORCODE_TOKEN       "TWAPI"        /* For TWAPI errors */

/* TWAPI error codes - used with the Tcl error facility */
typedef int TWAPI_ERROR;
#define TWAPI_NO_ERROR       0
#define TWAPI_INVALID_ARGS   1
#define TWAPI_BUFFER_OVERRUN 2
#define TWAPI_EXTRA_ARGS     3
#define TWAPI_BAD_ARG_COUNT  4
#define TWAPI_INTERNAL_LIMIT 5
#define TWAPI_INVALID_OPTION 6
#define TWAPI_INVALID_FUNCTION_CODE 7
#define TWAPI_BUG            8
#define TWAPI_UNKNOWN_OBJECT 9
#define TWAPI_SYSTEM_ERROR  10
#define TWAPI_REGISTER_WAIT_FAILED 11
#define TWAPI_BUG_INVALID_STATE_FOR_OP 12
#define TWAPI_OUT_OF_RANGE 13
#define TWAPI_UNSUPPORTED_TYPE 14
#define TWAPI_REGISTERED_POINTER_EXISTS 15
#define TWAPI_REGISTERED_POINTER_TAG_MISMATCH 16
#define TWAPI_REGISTERED_POINTER_NOTFOUND 17
#define TWAPI_NULL_POINTER 18
#define TWAPI_REGISTERED_POINTER_IS_NOT_COUNTED 19
#define TWAPI_INVALID_COMMAND_SCOPE 20
#define TWAPI_SCRIPT_ERROR 21
#define TWAPI_INVALID_DATA 22
#define TWAPI_INVALID_PTR 23

/*
 * Map TWAPI error codes into Win32 error code format.
 */
#define TWAPI_WIN32_ERROR_FAC 0xABC /* 12-bit facility to distinguish from app */
#define TWAPI_ERROR_TO_WIN32(code) (0xE0000000 | (TWAPI_WIN32_ERROR_FAC < 16) | (code))
#define IS_TWAPI_WIN32_ERROR(code) (((code) >> 16) == (0xe000 | TWAPI_WIN32_ERROR_FAC))
#define TWAPI_WIN32_ERROR_TO_CODE(winerr) ((winerr) & 0x0000ffff)

/**********************
 * Misc utility macros
 **********************/

#define CHECK_RESULT(value_or_fncall)           \
    do {                                        \
        int result_ = value_or_fncall;           \
        if (result_ != TCL_OK)                   \
            return result_;                      \
    } while (0)

/* Verify a Tcl_Obj is an integer/long and return error if not */
#define CHECK_INTEGER_OBJ(interp_, intvar_, objp_)       \
    do {                                                                \
        if (ObjToInt((interp_), (objp_), &(intvar_)) != TCL_OK) \
            return TCL_ERROR;                                           \
    } while (0)

#define CHECK_DWORD_OBJ(interp_, dwvar_, objp_)       \
    do {                                                                \
        if (ObjToDWORD((interp_), (objp_), &(dwvar_)) != TCL_OK) \
            return TCL_ERROR;                                           \
    } while (0)


/* Check number of arguments */
#define CHECK_NARGS(interp_, n_, m_)                        \
    do {                                                                \
        if ((n_) != (m_))                                               \
            return TwapiReturnError((interp_), TWAPI_BAD_ARG_COUNT);    \
    } while (0)

#define CHECK_NARGS_RANGE(interp_, nargs_, min_, max_)                  \
    do {                                                                \
        if ((nargs_) < (min_) || (nargs_) > (max_))                     \
            return TwapiReturnError((interp_), TWAPI_BAD_ARG_COUNT);    \
    } while (0)

/* Below are NOT an integer overflow check. Rather it only checks that
   a Tcl_Size value first in a DWORD. The len_ Tcl_Size expr itself may
   overflow! */
#define DWORD_LIMIT_CHECK(interp_, len_) \
    ((len_) <= UINT_MAX ? TCL_OK : TwapiReturnErrorUIntMax(interp_))

#define CHECK_DWORD(interp_, len_) \
    do { \
        TCL_RESULT ret_ = DWORD_LIMIT_CHECK((interp_), (len_)); \
        if (ret_ != TCL_OK) \
            return ret_; \
    } while (0)

#define ULONG_LIMIT_CHECK DWORD_LIMIT_CHECK
#define CHECK_ULONG CHECK_DWORD

/* String equality test - check first char before calling strcmp */
#define STREQ(x, y) ( (((x)[0]) == ((y)[0])) && ! lstrcmpA((x), (y)) )
#define STREQUN(x, y, n) \
    (((x)[0] == (y)[0]) && (strncmp(x, y, n) == 0))
#define WSTREQ(x, y) ( (((x)[0]) == ((y)[0])) && ! lstrcmpW((x), (y)) )

/* Make a pointer null if it points to empty element (generally used
   when we want to treat empty strings as null pointers */
#define NULLIFY_EMPTY(s_) if ((s_) && ((s_)[0] == 0)) (s_) = NULL
/* And the converse */
#define EMPTIFY_NULL(s_) if ((s_) == NULL) (s_) = L"";


/**********************************************
 * Macros and definitions dealing with Tcl_Obj
 **********************************************/
enum TwapiTclType {
    TWAPI_TCLTYPE_NONE  = 0,
    TWAPI_TCLTYPE_STRING,
    TWAPI_TCLTYPE_BOOLEAN,
    TWAPI_TCLTYPE_INT,
    TWAPI_TCLTYPE_DOUBLE,
    TWAPI_TCLTYPE_BYTEARRAY,
    TWAPI_TCLTYPE_LIST,
    TWAPI_TCLTYPE_DICT,
    TWAPI_TCLTYPE_WIDEINT,
    TWAPI_TCLTYPE_BOOLEANSTRING,
    TWAPI_TCLTYPE_NATIVE_END,
    TWAPI_TCLTYPE_OPAQUE = TWAPI_TCLTYPE_NATIVE_END, /* Added by Twapi */
    TWAPI_TCLTYPE_VARIANT,      /* Added by Twapi */
    TWAPI_TCLTYPE_BOUND
};

/* Create a string obj from a string literal. */
#define STRING_LITERAL_OBJ(x) ObjFromStringN((x), sizeof(x)-1)

/*
 *  Macros to append field name and values to a list
 */

/* Append a struct DWORD field name and value to a Tcl list object */
#define Twapi_APPEND_DWORD_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_), (listp_), ObjFromDWORD(((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_LONG_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_), (listp_), ObjFromLong(((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_WORD_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_), (listp_), ObjFromLong((WORD)((structp_)->field_))); \
  } while (0)

/* Append a struct ULONGLONG field name and value to a Tcl list object */
#define Twapi_APPEND_ULONGLONG_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_), (listp_), ObjFromWideInt((ULONGLONG)(structp_)->field_)); \
  } while (0)

/* Append a struct SIZE_T field name and value pair to a Tcl list object */
#ifdef _WIN64
#define Twapi_APPEND_SIZE_T_FIELD_TO_LIST Twapi_APPEND_ULONGLONG_FIELD_TO_LIST
#else
#define Twapi_APPEND_SIZE_T_FIELD_TO_LIST Twapi_APPEND_DWORD_FIELD_TO_LIST
#endif
#define Twapi_APPEND_ULONG_PTR_FIELD_TO_LIST Twapi_APPEND_SIZE_T_FIELD_TO_LIST

/* Appends a struct LARGE_INTEGER field name and value pair to a given Tcl list object */
#define Twapi_APPEND_LARGE_INTEGER_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_), (listp_), ObjFromWideInt((structp_)->field_.QuadPart)); \
  } while (0)

/* Appends a struct Unicode field name and string pair to a Tcl list object */
#define Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_),(listp_), ObjFromWinCharsN(((structp_)->field_ ? (structp_)->field_ : L""), -1)); \
  } while (0)

/* Appends a struct char string field name and string pair to a Tcl list object */
#define Twapi_APPEND_LPCSTR_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_),(listp_), ObjFromString(((structp_)->field_ ? (structp_)->field_ : ""))); \
  } while (0)


/* Appends a struct Unicode field name and string to a Tcl list object */
#define Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_),(listp_), ObjFromLSA_UNICODE_STRING(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_UUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_),(listp_), ObjFromUUID(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_LUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_),(listp_), ObjFromLUID(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_PSID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_Obj *obj = ObjFromSIDNoFail((structp_)->field_); \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_),(listp_), obj ? obj : ObjFromEmptyString()); \
  } while (0)

#define Twapi_APPEND_GUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    ObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    ObjAppendElement((interp_),(listp_), ObjFromGUID(&((structp_)->field_))); \
  } while (0)

/*
 * Macros to build ordered list of names and values of the fields
 * in a struct while maintaining consistency in the order.
 * See services.i for examples of usage
 */
#define FIELD_NAME_OBJ(x, unused, unused2) STRING_LITERAL_OBJ(# x)
#define FIELD_VALUE_OBJ(field, func, structp) func(structp->field)

/******************************
 * Tcl version dependent stuff
 ******************************/
struct TwapiTclVersion {
    int major;
    int minor;
    int patchlevel;
    int reltype;
};


#if TCL_MAJOR_VERSION < 9
/*
 * We need to access platform-dependent internal stubs. For
 * example, the Tcl channel system relies on specific values to be used
 * for EAGAIN, EWOULDBLOCK etc. These are actually compiler-dependent.
 * so the only way to make sure we are using a consistent Win32->Posix
 * mapping is to use the internal Tcl mapping function.
 */
struct TwapiTcl85IntPlatStubs {
    int   magic;
    void *hooks;
    void (*tclWinConvertError) (DWORD errCode); /* 0 */
    int (*fn2[29])(); /* Totally 30 fns, (index 0-29) */
};
extern struct TwapiTcl85IntPlatStubs *tclIntPlatStubsPtr;
#define TWAPI_TCL85_INT_PLAT_STUB(fn_) (((struct TwapiTcl85IntPlatStubs *)tclIntPlatStubsPtr)->fn_)
#endif


/*******************
 * Misc definitions
 *******************/

/*
 * Type for generating ids. Because they are passed around in windows
 * messages, make them the same size as DWORD_PTR though we would have
 * liked them to be 64 bit even on 32-bit platforms.
 */
typedef LONG_PTR TwapiId;
#define ObjFromTwapiId ObjFromLONG_PTR
#define ObjToTwapiId ObjToLONG_PTR
#define INVALID_TwapiId    0
#define TWAPI_NEWID Twapi_NewId

/* Used to maintain context for common NetEnum* interfaces */
typedef struct _TwapiEnumCtx {
    Tcl_Interp *interp;
    Tcl_Obj    *objP;
} TwapiEnumCtx;

typedef struct {
    LPBYTE netbufP;     /* If non-NULL, points to buffer to be freed
                           with NetApiBufferFree */
    NET_API_STATUS status;
    DWORD level;
    DWORD entriesread;
    DWORD totalentries;
    DWORD_PTR  hresume;
} TwapiNetEnumContext;

typedef void TWAPI_FILEVERINFO;

/****************************************************************
 * Defintitions used for conversion from Tcl_Obj to C types
 ****************************************************************/

/*
 * Used to pass a typed result to TwapiSetResult
 * Do NOT CHANGE VALUES AS THEY MAY ALSO BE REFLECTED IN TCL CODE
 */
typedef enum {
    TRT_BADFUNCTIONCODE = 0,
    TRT_BOOL = 1,
    TRT_EXCEPTION_ON_FALSE = 2,
    TRT_HWND = 3,
    TRT_UNICODE = 4,
    TRT_OBJV = 5,            /* Array of Tcl_Obj * */
    TRT_RECT = 6,
    TRT_HANDLE = 7,
    TRT_CHARS = 8,           /* char string */
    TRT_BINARY = 9,
    TRT_CHARS_DYNAMIC = 10,  /* Char string to be freed through TwapiFree */
    TRT_DWORD = 11,
    TRT_HGLOBAL = 12,
    TRT_NONZERO_RESULT = 13,
    TRT_EXCEPTION_ON_ERROR = 14,
    TRT_HDC = 15,
    TRT_HMONITOR = 16,
    TRT_FILETIME = 17,
    TRT_EMPTY = 18,
    TRT_EXCEPTION_ON_MINUSONE = 19,
    TRT_UUID = 20,
    TRT_LUID = 21,
    TRT_SC_HANDLE = 22,
    TRT_HDESK = 23,
    TRT_HWINSTA = 24,
    TRT_POINT = 25,
    TRT_VALID_HANDLE = 26, // Unlike TRT_HANDLE, INVALID_HANDLE is an error, not NULL
    TRT_GETLASTERROR = 27,   /* Set result as error code from GetLastError() */
    TRT_EXCEPTION_ON_WNET_ERROR = 28,
    TRT_DWORD_PTR = 29,
    TRT_PTR = 30,
    TRT_NONNULL_PTR = 31,
    TRT_INTERFACE = 32,         /* COM interface */
    TRT_OBJ = 33,
    TRT_UNINITIALIZED = 34,     /* Error - result not initialized */
    TRT_VARIANT = 35,           /* Must VarientInit before use ! */
    TRT_LPOLESTR = 36,    /* WCHAR string to be freed through CoTaskMemFree
                             Note these are NOT BSTR's
                           */
    TRT_SYSTEMTIME = 37,
    TRT_DOUBLE = 38,
    TRT_GUID = 39,  /* Also use for IID, CLSID; string rep differs from TRT_UUID
 */
    TRT_HKEY = 40,
    TRT_TCL_RESULT = 41,             /* Interp result already set. Return ival
                                        field as status */
    TRT_NTSTATUS = 42,
    TRT_LSA_HANDLE = 43,
    TRT_INT = 44,
    TRT_HDEVINFO = 45,
    TRT_PIDL = 46,              /* Freed using CoTaskMemFree */
    TRT_WIDE = 47,              /* Tcl_WideInt */
    TRT_UNICODE_DYNAMIC = 48,     /* Unicode to be freed through TwapiFree */
    TRT_TWAPI_ERROR = 49,          /* TWAPI error code in ival*/
    TRT_HRGN = 50,
    TRT_HMODULE = 51,
    TRT_HMACHINE = 52,
    TRT_CONFIGRET = 53,         /* PnP Manager status (CM_* calls) */
    TRT_GETLASTERROR_SETUPAPI = 54, /* Setup* calls */
    TRT_LONG = 55,
} TwapiResultType;

typedef struct {
    TwapiResultType type;
    union {
        int ival;
        LONG lval;
        DWORD uval;
        double dval;
        BOOL bval;
        DWORD_PTR dwp;
        Tcl_WideInt wide;
        LPVOID pv;
        HANDLE hval;
        HMODULE hmodule;
        HWND hwin;
        HKEY hkey;
        struct {
            WCHAR *str;
            Tcl_Size len;         /* len == -1 if str is null terminated */
        } unicode;
        struct {
            char *str;
            Tcl_Size len;         /* len == -1 if str is null terminated */
        } chars;
        struct {
            unsigned char  *p;
            Tcl_Size len;
        } binary;
        Tcl_Obj *obj;
        struct {
            Tcl_Size nobj;
            Tcl_Obj **objPP;
        } objv;
        RECT rect;
        POINT point;
        FILETIME filetime;
        UUID uuid;
        GUID guid;              /* Formatted differently from uuid */
        LUID luid;
        struct {
            void *p;
            char *name;
        } ifc;
        struct {
            void *p;
            char *name;
        } ptr;                  /* TRT_NONNULL_PTR and TRT_PTR */
        VARIANT var;            /* Must VariantInit before use!! */
        LPOLESTR lpolestr; /* WCHAR string to be freed through CoTaskMemFree */
        SYSTEMTIME systime;
        LPITEMIDLIST pidl;
    } value;
} TwapiResult;

#define TwapiResult_SET_NONNULL_PTR(res_, name_, val_) \
    do {                                           \
        (res_).type = TRT_NONNULL_PTR;             \
        (res_).value.ptr.p = (val_);       \
        (res_).value.ptr.name = # name_;           \
    } while (0)

#define TwapiResult_SET_PTR(res_, name_, val_)  \
    do {                                        \
        (res_).type = TRT_PTR;                  \
        (res_).value.ptr.p = (val_);    \
        (res_).value.ptr.name = # name_;        \
    } while (0)


/*
 * Macros for passing arguments to TwapiGetArgs.
 * v is a variable of the appropriate type.
 * I is an int variable or const
 * n is a variable of type int
 * typestr - is any type string such as "HSERVICE" that indicates the type
 * fn - is a function to call to convert the value. The function
 *   should have the prototype TwapiGetArgFn
 */
#define ARGEND      0
#define ARGTERM     1
#define ARGBOOL    'b'
#define ARGBA     'B'
#define ARGDOUBLE  'd'
#define ARGEMPTYASNULL 'E'
#define ARGINT     'i'
#define ARGWIDE    'I'
#define ARGSIZE    'j'
#define ARGHKEY    'k'
#define ARGTOKENNULL 'N'
#define ARGOBJ     'o'
#define ARGPTR      'p'
#define ARGDWORD_PTR 'P'
#define ARGVA   'r'
#define ARGVW   'R'
#define ARGASTR      's'
#define ARGASTRN     'S'
#define ARGWSTR     'u'
#define ARGWSTRN    'U'
#define ARGVAR     'v'
#define ARGVARWITHDEFAULT 'V'
#define ARGWORD     'w'
#define ARGDWORD     'W'
#define ARGSKIP     'x'   /* Leave arg parsing to caller */
#define ARGUNUSED ARGSKIP /* For readability when even caller does not care */
#define ARGVERIFIEDPTR 'z'
#define ARGVERIFIEDORNULL 'Z'
#define ARGUSEDEFAULT '?'

/*
 * These silly functions are to ensure pointer type safety when
 * using the varargs argument parsing functions
 */
TWAPI_INLINE int *reflectIntPtr(int *p) { return p; }
TWAPI_INLINE unsigned int *reflectUIntPtr(unsigned int *p) { return p; }
TWAPI_INLINE long *reflectLongPtr(long *p) { return p; }
TWAPI_INLINE WORD *reflectWordPtr(WORD *p) { return p; }
TWAPI_INLINE DWORD *reflectDWordPtr(DWORD *p) { return p; }
TWAPI_INLINE double *reflectDoublePtr(double *p) { return p; }
TWAPI_INLINE Tcl_Size *reflectTclSizePtr(Tcl_Size *p) { return p; }
TWAPI_INLINE Tcl_WideInt *reflectTclWidePtr(Tcl_WideInt *p) { return p; }
TWAPI_INLINE Tcl_Obj **reflectTclObjPtr(Tcl_Obj **p) { return p; }
TWAPI_INLINE WCHAR **reflectWCharPtr(WCHAR **p) { return p; }
TWAPI_INLINE char **reflectCharPtr(char **p) { return p; }


#define GETBOOL(v)    ARGBOOL, reflectIntPtr(&(v))
#define GETBA(v, n)  ARGBA, &(v), reflectTclSizePtr(&(n))
#define GETINT(v)     ARGINT, reflectIntPtr(&(v))
#define GETLONG(v)     ARGINT, reflectLongPtr(&(v))
#define GETDWORD(v)     ARGDWORD, reflectDWordPtr(&(v))
#define GETUINT(v)     ARGDWORD, reflectUIntPtr(&(v))
#define GETWIDE(v)    ARGWIDE, reflectTclWidePtr(&(v))
#define GETSIZE(v)    ARGSIZE, reflectTclSizePtr(&(v))
#define GETDOUBLE(v)  ARGDOUBLE, reflectDoublePtr(&(v))
#define GETOBJ(v)     ARGOBJ, reflectTclObjPtr(&(v))
#define GETDWORD_PTR(v) ARGDWORD_PTR, &(v)
#define GETASTR(v)      ARGASTR, reflectCharPtr(&(v))
#define GETASTRN(v, n)  ARGASTRN, reflectCharPtr(&(v)), reflectTclSizePtr(&(n))
#define GETWSTR(v)     ARGWSTR, reflectWCharPtr(&(v))
#define GETWSTRN(v, n) ARGWSTRN, reflectWCharPtr(&(v)), reflectTclSizePtr(&(n))
#define GETEMPTYASNULL(v) ARGEMPTYASNULL, &(v)
#define GETTOKENNULL(v) ARGTOKENNULL, &(v)
#define GETWORD(v)     ARGWORD, reflectWordPtr(&(v))
#define GETPTR(v, typesym) ARGPTR, &(v), #typesym
#define GETVOIDP(v)    ARGPTR, &(v), NULL
#define GETVERIFIEDPTR(v, typesym, verifier)    ARGVERIFIEDPTR, &(v), #typesym, (verifier)
#define GETVERIFIEDVOIDP(v, verifier)    ARGVERIFIEDPTR, &(v), NULL, (verifier)
#define GETVERIFIEDORNULL(v, typesym, verifier)    ARGVERIFIEDORNULL, &(v), #typesym, (verifier)
#define GETHANDLE(v)   GETVOIDP(v)
#define GETHANDLET(v, typesym) GETPTR(v, typesym)
#define GETHKEY(v)     ARGHKEY, &(v)
#define GETHWND(v) GETHANDLET(v, HWND)
#define GETVAR(v, fn)  ARGVAR, &(v), fn
#define GETVARWITHDEFAULT(v, fn)  ARGVARWITHDEFAULT, &(v), fn
#define GETGUID(v)     GETVAR(v, ObjToGUID)
#define GETUUID(v)     GETVAR(v, ObjToUUID)
/* For GETARGVA/GETWARGVW, v is of type char **, or WCHAR ** */
#define GETARGVA(v, n) ARGVA, (&v), reflectTclSizePtr(&(n))
#define GETARGVW(v, n) ARGVW, (&v), reflectTclSizePtr(&(n))

typedef int (*TwapiGetArgsFn)(Tcl_Interp *, Tcl_Obj *, void *);

/*
 * Registry value type
 */
typedef struct TwapiRegValue {
    void *bytes;
    DWORD size;
    DWORD type; /* REG_* */
    union {
        LONG           lval;
        Tcl_WideInt    wide;
    } u;
} TwapiRegValue;

/*
 * Forward decls
 */
typedef struct _TwapiInterpContext TwapiInterpContext;
ZLINK_CREATE_TYPEDEFS(TwapiInterpContext);
ZLIST_CREATE_TYPEDEFS(TwapiInterpContext);

typedef struct _TwapiCallback TwapiCallback;
ZLINK_CREATE_TYPEDEFS(TwapiCallback);
ZLIST_CREATE_TYPEDEFS(TwapiCallback);


/*
 * We need to keep track of handles that are being tracked by the
 * thread pool so they can be released on interp deletion even if
 * the application code does not explicitly release them.
 * NOTE: currently not all modules make use of this but probably should - TBD.
 */
typedef struct _TwapiThreadPoolRegistration TwapiThreadPoolRegistration;
ZLINK_CREATE_TYPEDEFS(TwapiThreadPoolRegistration);
ZLIST_CREATE_TYPEDEFS(TwapiThreadPoolRegistration);
typedef struct _TwapiThreadPoolRegistration {
    HANDLE handle;              /* Handle being waited on by thread pool */
    HANDLE tp_handle;           /* Corresponding handle returned by pool */
    TwapiInterpContext *ticP;
    ZLINK_DECL(TwapiThreadPoolRegistration); /* Link for tracking list */

    /* To be called when a HANDLE is signalled */
    void (*signal_handler) (TwapiInterpContext *ticP, TwapiId id, HANDLE h, DWORD);

    /*
     * To be called when handle wait is being unregistered. Routine should
     * take care to handle the case where ticP, ticP->interp and/or h is NULL.
     */
    void (*unregistration_handler)(TwapiInterpContext *ticP, TwapiId id, HANDLE h);

    TwapiId id;                 /* We need an id because OS handles can be
                                   reused and therefore cannot be used
                                   to filter stale events that have been
                                   queued for older handles with same value
                                   that have since been closed.
                                 */

    /* Only accessed from Interp thread so no locking */
    ULONG nrefs;
} TwapiThreadPoolRegistration;



/*
 * For asynchronous notifications of events, there is a common framework
 * that passes events from the asynchronous handlers into the Tcl event
 * dispatch loop. From there, the framework calls a function of type
 * TwapiCallbackFn. On entry to this function,
 *  - a non-NULL pointer to the callback structure (cbP) is passed in
 *  - the cbP->ticP, which contains the interp context is also guaranteed
 *    to be non-NULL.
 *  - the cbP->ticP->interp is the Tcl interp if non-NULL. This may be NULL
 *    if the original associated interp has been logically or physically
 *    deleted.
 *  - the cbP->clientdata* fields may contain any callback-specific data
 *    set by the enqueueing module.
 *  - the cbP->winerr field is set by the enqueuing module to ERROR_SUCCESS
 *    or a Win32 error code. It is up to the callback and enqueuing module
 *    to figure out what to do with it.
 *  - the cbP pointer may actually point to a "derived" structure where
 *    the callback structure is just the header. The enqueuing module
 *    should use the TwapiCallbackNew function to allocate
 *    and initialize. This function allows the size of allocated storage
 *    to be specified.
 *
 * If the Tcl interp is valid (non-NULL), the callback function is expected
 * to invoke an appropriate script in the interp and store an appropriate
 * result in the cbP->response field. Generally, callbacks build a script
 * and make use of the TwapiEvalAndUpdateCallback utility function to
 * invoke the script and store the result in cbP->response.
 *
 * If the callback function returns TCL_OK, the cbP->status and cbP->response
 * fields are returned to the enqueuing code if it is waiting for a response.
 * Note that the cbP->status may itself contain a Win32 error code to
 * indicate an error. It is entirely up to the enqueuing module to interpret
 * the response.
 *
 * If the callback fails with TCL_ERROR, the framework will set cbP->status
 * to an appropriate Win32 error code and set cbP->response to TRT_EMPTY.
 *
 * In all cases (success or fail), any additional resources attached
 * to cbP, for example buffers, should be freed, or arranged to be freed,
 * by the callback. Obviously, the framework cannot arrange for this.
 */
typedef int TwapiCallbackFn(struct _TwapiCallback *cbP);

/*
 * Definitions relating to queue of pending callbacks. All pending callbacks
 * structure definitions must start with this structure as the header.
 */

/* Creates list link definitions */
typedef struct _TwapiCallback {
    struct _TwapiInterpContext *ticP; /* Interpreter context */
    TwapiCallbackFn  *callback;  /* Function to call back - see notes
                                       in the TwapiCallbackFn typedef */
    LONG volatile     nrefs;       /* Ref count - use InterlockedIncrement */
    ZLINK_DECL(TwapiCallback); /* Link for list */
    HANDLE            completion_event;
    DWORD             winerr;         /* Win32 error code. Used in both
                                         callback request and response */
    /*
     * Associates with a particular notification handle. Not used by all
     * notifications.
     */
    TwapiId           receiver_id;
    DWORD_PTR         clientdata;     /* For use by client code */
    DWORD_PTR         clientdata2;    /* == ditto == */
    union {
        TwapiResult response;
        struct {
            POINTS message_pos;
            DWORD  ticks;
        } wm_state;             /* Used for Window message notifications
                                   (where there is no response) */
    } ;
} TwapiCallback;


/*
 * Thread local storage area. This is initialized when the extension
 * is loaded in the thread by an interpreter if no other interpreter
 * in that thread has already done so. It is deallocated when all
 * interps AND TwapiInterpContexts in that thread are deleted. This is
 * tracked through reference counts which are incr/decr'ed whenever
 * an interp or a new TwapiInterpContext is allocated/freed.
 *
 * IT IS NOT ACCESSIBLE IN A THREAD UNLESS AT LEAST ONE INTERP IS
 * STILL ALIVE
 */
typedef struct _TwapiTls {
    Tcl_ThreadId thread;

    /*
     * Every thread will have a memlifo to be used as a software stack.
     * This is initialized when the TwapiTls blob is allocated and
     * released when the thread terminates.
     */
    MemLifo memlifo;

    /*
     * ffiObj points to a Tcl_DictObj that maps the DLL and function
     * to its address. The key is the DLL name/path and function name.
     * The value being the address of the function
     * if found and 0 otherwise.
     */
    Tcl_Obj *ffiObj;

    int nrefs;                  /* Reference count */

#define TWAPI_TLS_SLOTS 8
    DWORD_PTR slots[TWAPI_TLS_SLOTS];
/* Unsafe access to a slot */
#define TWAPI_TLS_SLOT_UNSAFE(slot_) (Twapi_GetTls()->slots[slot_])
} TwapiTls;

/*
 * Static information associated with a Twapi module
 */
typedef void TwapiInterpContextCleanup(TwapiInterpContext *);
typedef TCL_RESULT TwapiModuleCallInitializer(Tcl_Interp *interp, TwapiInterpContext *ticP);
typedef struct _TwapiModuleDef {
    /* The name of the module */
    CONST char *name;

    /*
     * Function to call to initialize the module commands. The second
     * arg can be NULL if the module has not asked for a TwapiInterpContext.
     */
    TCL_RESULT (*initializer)(Tcl_Interp *, TwapiInterpContext *); // NULL ok

    /* Function to call to clean up the module. */
    void (*finalizer)(TwapiInterpContext *); // NULL ok

    /* Debug / trace flags for the module */
    unsigned long log_flags;
} TwapiModuleDef;

/*
 * TwapiInterpContext keeps track of a per-interpreter context.
 * This is allocated when twapi is loaded into an interpreter and
 * passed around as ClientData to most commands. It is reference counted
 * for deletion purposes and also placed on a global list for cleanup
 * purposes when a thread exits.
 *
 * Each twapi module can optionally allocate a TwapiInterpContext.
 * It can also get access to the base TwapiInterpContext via the
 * TwapiGetBaseContext function.
 *
 * The global list of contexts therefore contains multiple contexts for
 * various combinations of interpreters and modules.
 */
typedef struct _TwapiInterpContext {
    ZLINK_DECL(TwapiInterpContext); /* Links all the contexts, primarily
                                       to track cleanup requirements */

    LONG volatile         nrefs;   /* Reference count for alloc/free. */

    /* List of pending callbacks. Accessed controlled by the lock field */
    int              pending_suspended;       /* If true, do not pend events */
    ZLIST_DECL(TwapiCallback) pending;

    /*
     * List of handles registered with the Windows thread pool.
     * NOTE: TO BE ACCESSED ONLY FROM THE INTERP THREAD.
     */
    ZLIST_DECL(TwapiThreadPoolRegistration) threadpool_registrations;

    struct {
        HMODULE      hmod;        /* Handle of allocating module */
        TwapiModuleDef *modP;
        union {
            void *pval;
            int   ival;
            HWND  hwnd;
        } data;              /* For use by module, initialized to 0 */
    } module;

    /* Back pointer to the associated interp. This must only be modified or
     * accessed in the Tcl thread. THIS IS IMPORTANT AS THERE IS NO
     * SYNCHRONIZATION PROTECTION AND Tcl interp's ARE NOT MEANT TO BE
     * ACCESSED FROM OTHER THREADS
     */
    Tcl_Interp *interp;

    Tcl_ThreadId thread;     /* Id of interp thread */

    /*
     * Every thread will have a memlifo to be used as a software stack
     * stored in the thread local storage (TLS) blob.
     * This caches a pointer to it so if a TwapiInterpContext is
     * available, we do not need to look up the TLS.
     * The memlifo is cleaned up only when the thread exits at which
     * point this context and the attached interp will no longer exist.
     *
     * NOTE:THIS FIELD MUST ONLY BE ACCESSED IN THE INTERP THREAD.
     */
    MemLifo *memlifoP;

    /*
     * A single lock that is shared among multiple lists attached to this
     * structure as contention is expected to be low.
     */
    CRITICAL_SECTION lock;

    HWND          notification_win; /* Window used for various notifications */
} TwapiInterpContext;

/*
 * Structure used for passing events into the Tcl core. In some instances
 * you can directly inherit from Tcl_Event and do not have to go through
 * the expense of an additional allocation. However, Tcl_Event based
 * structures have to be allocated using Tcl_Alloc and we prefer not
 * to do that from outside Tcl threads. In such cases, pending_callback
 * is allocated using TwapiAlloc, passed between threads, and attached
 * to a Tcl_Alloc'ed TwapiTclEvent in a Tcl thread. See async.c
 */
typedef struct _TwapiTclEvent {
    Tcl_Event event;            /* Must be first field */
    TwapiCallback *pending_callback;
} TwapiTclEvent;



/****************************************
 * Definitions related to hidden windows
 * TBD - should this be in twapi_wm.h ?
 ****************************************/
/* Name of hidden window class */
#define TWAPI_HIDDEN_WINDOW_CLASS_L L"TwapiHiddenWindow"

/* Define offsets in window data */
#define TWAPI_HIDDEN_WINDOW_CONTEXT_OFFSET     0
#define TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET   (TWAPI_HIDDEN_WINDOW_CONTEXT_OFFSET + sizeof(LONG_PTR))
#define TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET (TWAPI_HIDDEN_WINDOW_CALLBACK_OFFSET + sizeof(LONG_PTR))
#define TWAPI_HIDDEN_WINDOW_DATA_SIZE       (TWAPI_HIDDEN_WINDOW_CLIENTDATA_OFFSET + sizeof(LONG_PTR))



/*****************************************************************
 * Prototypes and globals
 *****************************************************************/

typedef TCL_RESULT TwapiTclObjCmd(
    ClientData dummy,           /* Not used. */
    Tcl_Interp *interp,         /* Current interpreter. */
    int objc,                   /* Number of arguments. */
    Tcl_Obj *CONST objv[]);     /* Argument objects. */

#ifdef __cplusplus
extern "C" {
#endif

/* GLOBALS */
extern HMODULE gTwapiModuleHandle;     /* DLL handle to ourselves */
#if defined(TWAPI_STATIC_BUILD) || defined(TWAPI_SINGLE_MODULE)
#define MODULE_HANDLE gTwapiModuleHandle
#else
/* When building separate components, each module has its own DLL handle */
#define MODULE_HANDLE gModuleHandle
#endif
extern GUID gTwapiNullGuid;
extern struct TwapiTclVersion gTclVersion;

#define RETURN_ERROR_IF_UNTHREADED(interp_)    \
  do { \
    int result = Twapi_CheckThreadedTcl(interp_); \
    if (result != TCL_OK) \
        return result; \
} while (0)

typedef NTSTATUS (WINAPI *NtQuerySystemInformation_t)(int, PVOID, ULONG, PULONG);

/* Thread pool handle registration */
TCL_RESULT TwapiThreadPoolRegister(
    TwapiInterpContext *ticP,
    HANDLE h,
    ULONG timeout,
    DWORD flags,
    void (*signal_handler)(TwapiInterpContext *ticP, TwapiId, HANDLE, DWORD),
    void (*unregistration_handler)(TwapiInterpContext *ticP, TwapiId, HANDLE)
    );
void TwapiThreadPoolUnregister(
    TwapiInterpContext *ticP,
    TwapiId id
    );
void TwapiCallRegisteredWaitScript(TwapiInterpContext *ticP, TwapiId id, HANDLE h, DWORD timeout);
void TwapiThreadPoolRegistrationShutdown(TwapiThreadPoolRegistration *tprP);


TWAPI_EXTERN int Twapi_GenerateWin32Error(Tcl_Interp *interp, DWORD error, char *msg);

LRESULT TwapiEvalWinMessage(TwapiInterpContext *ticP, UINT msg, WPARAM wParam, LPARAM lParam);

/* Tcl_Obj manipulation and conversion - basic Windows types */

TWAPI_EXTERN Tcl_Obj *ObjNewList(Tcl_Size objc, Tcl_Obj * const objv[]);
TWAPI_EXTERN Tcl_Obj *ObjEmptyList(void);
TWAPI_EXTERN TCL_RESULT ObjListLength(Tcl_Interp *interp, Tcl_Obj *l, Tcl_Size *lenP);
TWAPI_EXTERN TCL_RESULT ObjListIndex(Tcl_Interp *interp, Tcl_Obj *l, Tcl_Size ix, Tcl_Obj **);

TWAPI_EXTERN TCL_RESULT ObjAppendElement(Tcl_Interp *interp, Tcl_Obj *l, Tcl_Obj *e);
TWAPI_EXTERN TCL_RESULT ObjGetElements(Tcl_Interp *interp, Tcl_Obj *l, Tcl_Size *objcP, Tcl_Obj ***objvP);
TWAPI_EXTERN TCL_RESULT ObjListReplace(Tcl_Interp *interp, Tcl_Obj *l, Tcl_Size first, Tcl_Size count, Tcl_Size objc, Tcl_Obj *const objv[]);

TWAPI_EXTERN Tcl_Obj *ObjNewDict(void);
TWAPI_EXTERN TCL_RESULT ObjDictGet(Tcl_Interp *interp, Tcl_Obj *dictObj, Tcl_Obj *keyObj, Tcl_Obj **valueObjP);
TWAPI_EXTERN TCL_RESULT ObjDictPut(Tcl_Interp *interp, Tcl_Obj *dictObj, Tcl_Obj *keyObj, Tcl_Obj *valueObj);


#define Twapi_FreeNewTclObj(o_) do { if (o_) { ObjDecrRefs(o_); } } while (0)

Tcl_Obj *TwapiAppendObjArray(Tcl_Obj *resultObj, Tcl_Size objc, Tcl_Obj **objv,
                         char *join_string);
Tcl_Obj *ObjFromPOINTS(POINTS *ptP);
int ObjToFLASHWINFO (Tcl_Interp *interp, Tcl_Obj *obj, FLASHWINFO *fwP);
Tcl_Obj *ObjFromWINDOWINFO (WINDOWINFO *wiP);
Tcl_Obj *ObjFromWINDOWPLACEMENT(WINDOWPLACEMENT *wpP);
int ObjToWINDOWPLACEMENT(Tcl_Interp *, Tcl_Obj *objP, WINDOWPLACEMENT *wpP);
Tcl_Obj *ObjFromDISPLAY_DEVICE(DISPLAY_DEVICEW *ddP);
Tcl_Obj *ObjFromMONITORINFOEX(MONITORINFO *miP);
Tcl_Obj *ObjFromSYSTEM_POWER_STATUS(SYSTEM_POWER_STATUS *spsP);

TWAPI_EXTERN Tcl_Obj *ObjFromACE (Tcl_Interp *interp, void *aceP);
TWAPI_EXTERN TCL_RESULT ObjToACESWS (Tcl_Interp *interp, Tcl_Obj *aceobj, void **acePP);
TWAPI_EXTERN Tcl_Obj *ObjFromACL(Tcl_Interp *interp, ACL *aclP);

/* Window stuff */
int Twapi_EnumWindows(Tcl_Interp *interp);

/* System related */
int Twapi_TclGetChannelHandle(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
BOOL Twapi_IsWow64Process(HANDLE h, BOOL *is_wow64P);

/* Shares and LANMAN */
int Twapi_NetGetDCName(Tcl_Interp *interp, LPCWSTR server, LPCWSTR domain);

/* Security related */
int Twapi_LookupAccountName (Tcl_Interp *interp, LPCWSTR sysname, LPCWSTR name);
int Twapi_LookupAccountSid (Tcl_Interp *interp, LPCWSTR sysname, PSID sidP);
int Twapi_InitializeSecurityDescriptor(Tcl_Interp *interp);

/* ADSI related */
int Twapi_DsGetDcName(Tcl_Interp *interp, LPCWSTR systemnameP,
                      LPCWSTR domainnameP, UUID *guidP,
                      LPCWSTR sitenameP, ULONG flags);


/* COM stuff */

/* Note - because ifcp_ is typed, this has to be a macro */
#define TWAPI_STORE_COM_ERROR(interp_, hr_, ifcp_, iidp_)  \
    do { \
        ISupportErrorInfo *sei = NULL; \
        (ifcp_)->lpVtbl->QueryInterface((ifcp_), &IID_ISupportErrorInfo, (LPVOID*)&sei); \
        /* Twapi_AppendCOMError will accept NULL sei so no check for error */ \
        Twapi_AppendCOMError((interp_), (hr_), sei, (iidp_));           \
        if (sei) sei->lpVtbl->Release(sei);                              \
    } while (0)

#define TWAPI_GET_ISupportErrorInfo(sei_,ifcp_)    \
    do { \
        if (FAILED((ifcp_)->lpVtbl->QueryInterface((ifcp_), &IID_ISupportErrorInfo, (LPVOID*)&sei_))) { \
                sei_ = NULL; \
            } \
    } while (0)

int Twapi_CommandLineToArgv(Tcl_Interp *interp, LPCWSTR cmdlineP);
int Twapi_GetGUIThreadInfo(Tcl_Interp *interp, DWORD idThread);

#define Twapi_WTSFreeMemory(p_) do { if (p_) WTSFreeMemory(p_); } while (0)
int Twapi_WTSEnumerateProcesses(Tcl_Interp *interp, HANDLE wtsH);


/* Built-in commands */

/* Dispatcher routines */
int Twapi_InitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP);

/* General utility functions */
int WINAPI TwapiGlobCmp (const char *s, const char *pat);
int WINAPI TwapiGlobCmpCase (const char *s, const char *pat);

int Twapi_MemLifoDump(Tcl_Interp *, MemLifo *l);

#ifdef __cplusplus
} // extern "C"
#endif

/*
 * Exported functions
 */

TWAPI_EXTERN BOOL TwapiRtlGetVersion(LPOSVERSIONINFOW verP);
TWAPI_EXTERN int TwapiMinOSVersion(DWORD major, DWORD minor);

/* Memory allocation */

TWAPI_EXTERN void *TwapiAlloc(size_t sz);
TWAPI_EXTERN void *TwapiAllocSize(size_t sz, size_t *);
TWAPI_EXTERN void *TwapiAllocZero(size_t sz);
TWAPI_EXTERN void TwapiFree(void *p);
TWAPI_EXTERN WCHAR *TwapiAllocWString(WCHAR *, Tcl_Size len);
TWAPI_EXTERN WCHAR *TwapiAllocWStringFromObj(Tcl_Obj *, Tcl_Size *lenP);
TWAPI_EXTERN char *TwapiAllocAString(char *, Tcl_Size len);
TWAPI_EXTERN char *TwapiAllocAStringFromObj(Tcl_Obj *, Tcl_Size *lenP);
TWAPI_EXTERN void *TwapiReallocTry(void *p, size_t sz);
TWAPI_EXTERN void *TwapiAllocRegisteredPointer(Tcl_Interp *, size_t, void *tag);
TWAPI_EXTERN void TwapiFreeRegisteredPointer(Tcl_Interp *, void *, void *tag);


/* C - Tcl result and parameter conversion  */
TWAPI_EXTERN TCL_RESULT TwapiSetResult(Tcl_Interp *interp, TwapiResult *result);
TWAPI_EXTERN void TwapiClearResult(TwapiResult *resultP);
/* TBD - TwapiGetArgs* could also obe used to parse lists into C structs */
TWAPI_EXTERN_VA TCL_RESULT TwapiGetArgsVA(Tcl_Interp *interp, Tcl_Size objc, Tcl_Obj *CONST objv[], int fmt, va_list ap);
TWAPI_EXTERN_VA TCL_RESULT TwapiGetArgs(Tcl_Interp *interp, Tcl_Size objc, Tcl_Obj *CONST objv[], int fmt, ...);
TWAPI_EXTERN_VA TCL_RESULT TwapiGetArgsObj(Tcl_Interp *interp, Tcl_Obj *, int fmt, ...);
TWAPI_EXTERN_VA TCL_RESULT TwapiGetArgsExVA(TwapiInterpContext *ticP, Tcl_Size objc, Tcl_Obj *CONST objv[], int fmt, va_list ap);
TWAPI_EXTERN_VA TCL_RESULT TwapiGetArgsEx(TwapiInterpContext *ticP, Tcl_Size objc, Tcl_Obj *CONST objv[], int fmt, ...);
TWAPI_EXTERN_VA TCL_RESULT TwapiGetArgsExObj(TwapiInterpContext *ticP, Tcl_Obj *, int fmt, ...);
TWAPI_EXTERN void ObjSetStaticResult(Tcl_Interp *interp, CONST char s[]);
#define TwapiSetStaticResult ObjSetStaticResult
TWAPI_EXTERN TCL_RESULT ObjSetResult(Tcl_Interp *interp, Tcl_Obj *objP);
#define TwapiSetObjResult ObjSetResult
TWAPI_EXTERN Tcl_Obj *ObjGetResult(Tcl_Interp *interp);
TWAPI_EXTERN Tcl_Obj *ObjDuplicate(Tcl_Obj *);

/* errors.c */
TWAPI_EXTERN TCL_RESULT TwapiReturnSystemError(Tcl_Interp *interp);
TWAPI_EXTERN TCL_RESULT TwapiReturnError(Tcl_Interp *interp, int code);
TWAPI_EXTERN TCL_RESULT TwapiReturnErrorEx(Tcl_Interp *interp, int code, Tcl_Obj *objP);
TWAPI_EXTERN TCL_RESULT TwapiReturnErrorMsg(Tcl_Interp *interp, int code, char *msg);
TWAPI_EXTERN TCL_RESULT TwapiReturnErrorUIntMax(Tcl_Interp *interp);

TWAPI_EXTERN DWORD TwapiNTSTATUSToError(NTSTATUS status);
TWAPI_EXTERN Tcl_Obj *Twapi_MakeTwapiErrorCodeObj(int err);
TWAPI_EXTERN Tcl_Obj *Twapi_MapWindowsErrorToString(DWORD err);
TWAPI_EXTERN Tcl_Obj *Twapi_MakeWindowsErrorCodeObj(DWORD err, Tcl_Obj *);
TWAPI_EXTERN TCL_RESULT Twapi_AppendWNetError(Tcl_Interp *interp, unsigned long err);
TWAPI_EXTERN TCL_RESULT Twapi_AppendSystemErrorEx(Tcl_Interp *, unsigned long err, Tcl_Obj *extra);
#define Twapi_AppendSystemError2 Twapi_AppendSystemErrorEx
TWAPI_EXTERN TCL_RESULT Twapi_AppendSystemError(Tcl_Interp *, unsigned long err);
TWAPI_EXTERN int Twapi_AppendCOMError(Tcl_Interp *interp, HRESULT hr, ISupportErrorInfo *sei, REFIID iid);
TWAPI_EXTERN TCL_RESULT Twapi_WrongLevelError(Tcl_Interp *interp, int level);
TWAPI_EXTERN void TwapiWriteEventLogError(const char *msg);


/* Async handling related */

TWAPI_EXTERN void TwapiEnqueueTclEvent(TwapiInterpContext *ticP, Tcl_Event *evP);
#define TwapiCallbackRef(pcb_, incr_) InterlockedExchangeAdd(&(pcb_)->nrefs, (incr_))
TWAPI_EXTERN void TwapiCallbackUnref(TwapiCallback *pcbP, int);
TWAPI_EXTERN void TwapiCallbackDelete(TwapiCallback *pcbP);
TWAPI_EXTERN TwapiCallback *TwapiCallbackNew(
    TwapiInterpContext *ticP, TwapiCallbackFn *callback, int sz);
TWAPI_EXTERN int TwapiEnqueueCallback(
    TwapiInterpContext *ticP, TwapiCallback *pcbP,
    int enqueue_method,
    int timeout,
    TwapiCallback **responseP
    );
#define TWAPI_ENQUEUE_DIRECT 0
#define TWAPI_ENQUEUE_ASYNC  1
TWAPI_EXTERN int TwapiEvalAndUpdateCallback(TwapiCallback *cbP, int objc, Tcl_Obj *objv[], TwapiResultType response_type);

/* Tcl_Obj manipulation and conversion - basic Windows types */
int TwapiInitTclTypes(void);
TWAPI_EXTERN int TwapiGetTclType(Tcl_Obj *objP);

TWAPI_EXTERN void ObjIncrRefs(Tcl_Obj *);
TWAPI_EXTERN void ObjDecrRefs(Tcl_Obj *);
TWAPI_EXTERN void ObjDecrArrayRefs(int, Tcl_Obj *objv[]);

TWAPI_EXTERN TCL_RESULT ObjToEnum(Tcl_Interp *interp, Tcl_Obj *enumsObj, Tcl_Obj *nameObj, int *valP);


TWAPI_EXTERN Tcl_Obj *ObjFromOpaque(void *pv, char *name);

TWAPI_INLINE Tcl_Obj *ObjFromHANDLE(HANDLE h) {
    return ObjFromOpaque(h, "HANDLE");
}
TWAPI_INLINE Tcl_Obj *ObjFromHWND(HWND hwnd) {
    return ObjFromOpaque(hwnd, "HWND");
}
TWAPI_INLINE Tcl_Obj *ObjFromLPVOID(void *p) {
    return ObjFromOpaque(p, NULL);
}
TWAPI_INLINE Tcl_Obj *ObjFromHMODULE(HMODULE hmod) {
    return ObjFromOpaque(hmod, "HMODULE");
}
TWAPI_INLINE Tcl_Obj *ObjFromFARPROC(FARPROC fn) {
    return ObjFromOpaque(fn, "FARPROC");
}

/* The following macros assume objP_ typePtr points to Twapi's gVariantType */
#define VARIANT_REP_VALUE(objP_) ((objP_)->internalRep.ptrAndLongRep.ptr)
#define VARIANT_REP_VT(objP_)  ((objP_)->internalRep.ptrAndLongRep.value)
/* TWAPI-specific hack to indicate contents to be treated as a variable name. */
#define VT_TWAPI_VARNAME (VT_EMPTY|VT_BYREF)

/* The following macros assume objP_ typePtr points to Twapi's gOpaqueType */
#define OPAQUE_REP_VALUE(objP_) ((objP_)->internalRep.twoPtrValue.ptr1)
#define OPAQUE_REP_VALUE_SET(objP_) (objP_)->internalRep.twoPtrValue.ptr1
#define OPAQUE_REP_CTYPE(objP_)  ((Tcl_Obj *) (objP_)->internalRep.twoPtrValue.ptr2)
#define OPAQUE_REP_CTYPE_SET(objP_)  ((Tcl_Obj *) (objP_))->internalRep.twoPtrValue.ptr2

TCL_RESULT SetOpaqueFromAny(Tcl_Interp *interp, Tcl_Obj *objP);
TWAPI_EXTERN TCL_RESULT ObjToOpaque(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, const char *name);
TWAPI_EXTERN TCL_RESULT ObjToOpaqueMulti(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, int ntypes, char **types);
TWAPI_EXTERN TCL_RESULT ObjToVerifiedPointer(Tcl_Interp *interp, Tcl_Obj *objP, void **pvP, const char *name, void *verifier);
TWAPI_EXTERN TCL_RESULT ObjToVerifiedPointerOrNull(Tcl_Interp *interp, Tcl_Obj *objP, void **pvP, const char *name, void *verifier);
TWAPI_EXTERN TCL_RESULT ObjToVerifiedPointerTic(TwapiInterpContext *, Tcl_Obj *objP, void **pvP, const char *name, void *verifier);
TWAPI_EXTERN TCL_RESULT ObjToVerifiedPointerOrNullTic(TwapiInterpContext *, Tcl_Obj *objP, void **pvP, const char *name, void *verifier);

TWAPI_EXTERN TCL_RESULT ObjToLPVOID(Tcl_Interp *interp, Tcl_Obj *objP, HANDLE *pvP);
#define ObjToHANDLE ObjToLPVOID

TWAPI_EXTERN TCL_RESULT ObjToHKEY(Tcl_Interp *interp, Tcl_Obj *objP, HKEY *hkeyP);
TWAPI_INLINE TCL_RESULT ObjToHWND(Tcl_Interp *ip, Tcl_Obj *objP, HWND *pvP) {
    return ObjToOpaque(ip, objP, (void **)pvP, "HWND");
}

TWAPI_INLINE TCL_RESULT ObjToHMODULE(Tcl_Interp *ip, Tcl_Obj *objP, HMODULE *pvP) {
    return ObjToOpaque(ip, objP, (void **)pvP, "HMODULE");
}

TWAPI_INLINE TCL_RESULT ObjToFARPROC(Tcl_Interp *ip, Tcl_Obj *objP, FARPROC *pvP) {
    return ObjToOpaque(ip, objP, (void **)pvP, "FARPROC");
}

TWAPI_EXTERN Tcl_Obj *ObjFromBoolean(int bval);
TWAPI_EXTERN TCL_RESULT ObjToBoolean(Tcl_Interp *, Tcl_Obj *, int *);
TWAPI_EXTERN Tcl_Obj *ObjFromLong(long val);
TWAPI_EXTERN TCL_RESULT ObjToLong(Tcl_Interp *interp, Tcl_Obj *objP, long *);
TWAPI_EXTERN Tcl_Obj *ObjFromInt(int val);
TWAPI_EXTERN TCL_RESULT ObjToInt(Tcl_Interp *interp, Tcl_Obj *objP, int *);

TWAPI_EXTERN Tcl_Obj *ObjFromWideInt(Tcl_WideInt val);
TWAPI_EXTERN Tcl_Obj *ObjFromWideUInt(Tcl_WideUInt val);
/* Unsigneds need to be promoted to wide ints when converting to Tcl_Obj*/
TWAPI_INLINE Tcl_Obj *ObjFromDWORD(DWORD dw) {
    return ObjFromWideInt(dw);
}
TWAPI_EXTERN TCL_RESULT ObjToDWORD(Tcl_Interp *interp, Tcl_Obj *objP, DWORD *dwP);
#define ObjFromULONG      ObjFromDWORD
TWAPI_INLINE TCL_RESULT ObjToUINT(Tcl_Interp *ip, Tcl_Obj *objP, UINT *uiP) {
    return ObjToDWORD(ip, objP, (DWORD *)uiP);
}

TWAPI_EXTERN TCL_RESULT ObjToWideInt(Tcl_Interp *interp, Tcl_Obj *objP, Tcl_WideInt *wideP);
TWAPI_EXTERN TCL_RESULT ObjToWideUInt(Tcl_Interp *interp, Tcl_Obj *objP, Tcl_WideUInt *wideP);
TWAPI_EXTERN TCL_RESULT ObjToBits64(Tcl_Interp *interp, Tcl_Obj *objP, Tcl_WideInt *wideP);
TWAPI_EXTERN Tcl_Obj *ObjFromDouble(double val);
TWAPI_EXTERN TCL_RESULT ObjToDouble(Tcl_Interp *interp, Tcl_Obj *objP, double *);
TWAPI_EXTERN Tcl_Obj *ObjFromFloat(float val);
TWAPI_EXTERN TCL_RESULT ObjToFloat(Tcl_Interp *interp, Tcl_Obj *objP, float *);
TWAPI_EXTERN TCL_RESULT ObjToDWORD_PTR(Tcl_Interp *interp, Tcl_Obj *objP, DWORD_PTR *dwP);
#ifdef _WIN64
#define ObjFromDWORD_PTR(p_)  ObjFromULONGLONG((ULONGLONG)(p_))
#define ObjToLONG_PTR         ObjToWideInt
#define ObjFromLONG_PTR       ObjFromWideInt
#else  // ! _WIN64
#define ObjFromDWORD_PTR(p_)  ObjFromDWORD((DWORD_PTR)(p_))
#define ObjToLONG_PTR         ObjToLong
#define ObjFromLONG_PTR       ObjFromLong
#endif // _WIN64
#define ObjToULONG_PTR    ObjToDWORD_PTR
#define ObjFromULONG_PTR  ObjFromDWORD_PTR
#define ObjFromSIZE_T     ObjFromDWORD_PTR

TWAPI_EXTERN TCL_RESULT ObjToUCHAR(Tcl_Interp *interp, Tcl_Obj *obj, UCHAR *ucP);
TWAPI_EXTERN TCL_RESULT ObjToCHAR(Tcl_Interp *interp, Tcl_Obj *obj, CHAR *ucP);
TWAPI_EXTERN TCL_RESULT ObjToUSHORT(Tcl_Interp *interp, Tcl_Obj *obj, WORD *wordP);
#define ObjToWord ObjToUSHORT
TWAPI_EXTERN TCL_RESULT ObjToSHORT(Tcl_Interp *interp, Tcl_Obj *obj, SHORT *wordP);

#define ObjFromLARGE_INTEGER(val_) ObjFromWideInt((val_).QuadPart)
TWAPI_EXTERN Tcl_Obj *ObjFromULONGLONG(ULONGLONG ull);

TWAPI_EXTERN Tcl_Obj *ObjFromUCHARHex(UCHAR);
TWAPI_EXTERN Tcl_Obj *ObjFromUSHORTHex(USHORT);
TWAPI_EXTERN Tcl_Obj *ObjFromULONGHex(ULONG ull);
TWAPI_EXTERN Tcl_Obj *ObjFromULONGLONGHex(ULONGLONG ull);

TWAPI_EXTERN int TwapiWinCharsToUtf8(CONST WCHAR *wsP, Tcl_Size nchars, char *buf, Tcl_Size buf_sz);
TWAPI_EXTERN Tcl_Obj *TwapiUtf8ObjFromWinChars(CONST WCHAR *p, int len);

TWAPI_EXTERN Tcl_Obj *ObjFromEmptyString();
TWAPI_EXTERN Tcl_Size ObjCharLength(Tcl_Obj *);

TWAPI_EXTERN char *ObjToString(Tcl_Obj *objP);
TWAPI_EXTERN char *ObjToStringN(Tcl_Obj *objP, Tcl_Size *lenP);
TWAPI_EXTERN Tcl_Obj *ObjFromStringN(const char *s, Tcl_Size len);
TWAPI_EXTERN Tcl_Obj *ObjFromString(const char *s);
TWAPI_EXTERN Tcl_Obj *ObjFromStringLimited(const char *strP, Tcl_Size max, Tcl_Size *remain);

TWAPI_EXTERN Tcl_UniChar *ObjToTclUniChar(Tcl_Obj *objP);
TWAPI_EXTERN Tcl_UniChar *ObjToTclUniCharN(Tcl_Obj *objP, Tcl_Size *lenP);
TWAPI_EXTERN Tcl_Obj *ObjFromTclUniCharN(const Tcl_UniChar *ws, Tcl_Size len);
TWAPI_EXTERN Tcl_Obj *ObjFromTclUniChar(const Tcl_UniChar *ws);

#if (TCL_UTF_MAX >= 4) || defined(TWAPI_FORCE_WINCHARS)
TWAPI_EXTERN WCHAR *ObjToWinChars(Tcl_Obj *objP);
TWAPI_EXTERN WCHAR *ObjToWinCharsN(Tcl_Obj *objP, Tcl_Size *lenP);
TWAPI_EXTERN Tcl_Obj *ObjFromWinCharsN(const WCHAR *ws, Tcl_Size len);
TWAPI_EXTERN Tcl_Obj *ObjFromWinChars(const WCHAR *ws);
#else
#define ObjToWinChars ObjToTclUniChar
#define ObjToWinCharsN ObjToTclUniCharN
#define ObjToWinCharsDW ObjToTclUniCharDW
#define ObjFromWinChars ObjFromTclUniChar
#define ObjFromWinCharsN ObjFromTclUniCharN
#endif
TWAPI_EXTERN TCL_RESULT ObjToWinCharsDW(Tcl_Interp *interp,
                                        Tcl_Obj    *objP,
                                        DWORD      *lenP,
                                        WCHAR     **wsPP);
TWAPI_EXTERN Tcl_Obj *ObjFromWinCharsLimited(const WCHAR *wstrP, Tcl_Size max, Tcl_Size *remain);
TWAPI_EXTERN Tcl_Obj *ObjFromWinCharsNoTrailingSpace(const WCHAR *strP);

TWAPI_EXTERN Tcl_Obj *ObjFromByteArray(const unsigned char *bytes, Tcl_Size len);
TWAPI_EXTERN Tcl_Obj *ObjAllocateByteArray(Tcl_Size len, void **);
TWAPI_EXTERN Tcl_Obj *ObjFromByteArrayHex(const unsigned char *bytes, Tcl_Size len);
TWAPI_EXTERN unsigned char *ObjToByteArray(Tcl_Obj *objP, Tcl_Size *lenP);
TWAPI_EXTERN TCL_RESULT     ObjToByteArrayDW(Tcl_Interp     *ip,
                                             Tcl_Obj        *objP,
                                             DWORD          *dwP,
                                             unsigned char **bufPP);

TWAPI_EXTERN TCL_RESULT TwapiEncryptData(Tcl_Interp *, BYTE *, ULONG, BYTE *, ULONG *);
TWAPI_EXTERN TCL_RESULT TwapiDecryptData(Tcl_Interp *, BYTE *, ULONG , BYTE *, ULONG *);
TWAPI_EXTERN BYTE *TwapiDecryptDataSWS(Tcl_Interp *, BYTE *, ULONG, ULONG *);
TWAPI_EXTERN Tcl_Obj *ObjEncryptBytes(Tcl_Interp *, void *, Tcl_Size);
TWAPI_EXTERN void *ObjDecryptBytesExSWS(Tcl_Interp *, Tcl_Obj *, Tcl_Size, Tcl_Size *);
TWAPI_EXTERN Tcl_Obj *ObjEncryptWinChars(Tcl_Interp *, WCHAR *, Tcl_Size);
TWAPI_EXTERN WCHAR * ObjDecryptWinCharsSWS(Tcl_Interp *, Tcl_Obj *, Tcl_Size *);
TWAPI_EXTERN char * ObjDecryptUtf8SWS(Tcl_Interp *, Tcl_Obj *, Tcl_Size *);
TWAPI_EXTERN WCHAR * ObjDecryptPasswordSWS(Tcl_Obj *objP, Tcl_Size *ncharsP);

TWAPI_EXTERN Tcl_Obj *ObjFromLSA_UNICODE_STRING(const LSA_UNICODE_STRING *lsauniP);
TWAPI_EXTERN TCL_RESULT ObjToLSA_UNICODE_STRING(Tcl_Interp         *ip,
                                                Tcl_Obj            *objP,
                                                LSA_UNICODE_STRING *lsauniP);
TWAPI_EXTERN int ObjToLSASTRINGARRAYSWS(Tcl_Interp *interp, Tcl_Obj *obj,
                        LSA_UNICODE_STRING **arrayP, ULONG *countP);
TWAPI_EXTERN PSID TwapiSidFromStringSWS(char *strP);
TWAPI_EXTERN TCL_RESULT ObjToPSIDSWS(Tcl_Interp *, Tcl_Obj *obj, PSID *sidPP);
TWAPI_EXTERN TCL_RESULT ObjToPSIDNonNullSWS(Tcl_Interp *, Tcl_Obj *, PSID *sidPP);
TWAPI_EXTERN int ObjFromSID (Tcl_Interp *interp, SID *sidP, Tcl_Obj **objPP);
TWAPI_EXTERN Tcl_Obj *ObjFromSIDNoFail(SID *sidP);

TWAPI_EXTERN LPWSTR *ObjToMemLifoArgvW(TwapiInterpContext *ticP, Tcl_Obj *objP, Tcl_Size *argcP);
TWAPI_EXTERN int ObjToArgvW(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *argv, Tcl_Size argc, Tcl_Size *argcP);
TWAPI_EXTERN Tcl_Obj *ObjFromArgvA(Tcl_Size argc, char **argv);
TWAPI_EXTERN int ObjToArgvA(Tcl_Interp *interp, Tcl_Obj *objP, char **argv, Tcl_Size argc, Tcl_Size *argcP);
TWAPI_EXTERN LPWSTR ObjToLPWSTR_NULL_IF_EMPTY(Tcl_Obj *objP);

#define NULL_TOKEN "__null__"
#define NULL_TOKEN_L L"__null__"
TWAPI_EXTERN LPWSTR ObjToLPWSTR_WITH_NULL(Tcl_Obj *objP);

TWAPI_EXTERN Tcl_Obj *ObjFromPIDL(LPCITEMIDLIST pidl);
TWAPI_EXTERN int ObjToPIDL(Tcl_Interp *interp, Tcl_Obj *objP, LPITEMIDLIST *idsPP);
TWAPI_EXTERN void TwapiFreePIDL(LPITEMIDLIST idlistP);

#define ObjFromIDispatch(p_) ObjFromOpaque((p_), "IDispatch")
TWAPI_EXTERN int ObjToIDispatch(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP);
#define ObjFromIUnknown(p_) ObjFromOpaque((p_), "IUnknown")
#define ObjToIUnknown(ip_, obj_, ifc_) \
    ObjToOpaque((ip_), (obj_), (ifc_), "IUnknown")

TWAPI_EXTERN int ObjToVT(Tcl_Interp *interp, Tcl_Obj *obj, VARTYPE *vtP);
TWAPI_EXTERN Tcl_Obj *ObjFromBSTR (BSTR bstr);
TWAPI_EXTERN int ObjToBSTR (Tcl_Interp *, Tcl_Obj *, BSTR *);
TWAPI_EXTERN int ObjToRangedInt(Tcl_Interp *, Tcl_Obj *obj, int low, int high, int *iP);
TWAPI_EXTERN Tcl_Obj *ObjFromSYSTEMTIME(const SYSTEMTIME *timeP);
TWAPI_EXTERN int ObjToSYSTEMTIME(Tcl_Interp *interp, Tcl_Obj *timeObj, LPSYSTEMTIME timeP);
TWAPI_EXTERN Tcl_Obj *ObjFromFILETIME(const FILETIME *ftimeP);
TWAPI_EXTERN int ObjToFILETIME(Tcl_Interp *interp, Tcl_Obj *obj, FILETIME *cyP);
TWAPI_EXTERN Tcl_Obj *ObjFromTIME_ZONE_INFORMATION(const TIME_ZONE_INFORMATION *tzP);
TWAPI_EXTERN TCL_RESULT ObjToTIME_ZONE_INFORMATION(Tcl_Interp *interp, Tcl_Obj *tzObj, TIME_ZONE_INFORMATION *tzP);
TWAPI_EXTERN Tcl_Obj *ObjFromCY(const CY *cyP);
TWAPI_EXTERN int ObjToCY(Tcl_Interp *interp, Tcl_Obj *obj, CY *cyP);
TWAPI_EXTERN Tcl_Obj *ObjFromDECIMAL(DECIMAL *cyP);
TWAPI_EXTERN int ObjToDECIMAL(Tcl_Interp *interp, Tcl_Obj *obj, DECIMAL *cyP);
TWAPI_EXTERN VARTYPE ObjTypeToVT(Tcl_Obj *objP);
TWAPI_EXTERN Tcl_Obj *ObjFromVARIANT(VARIANT *varP, int value_only);
TWAPI_EXTERN TCL_RESULT ObjToVARIANT(Tcl_Interp *interp, Tcl_Obj *objP, VARIANT *varP, VARTYPE vt);

/* Note: the returned multiszPP must be free()'ed */
TWAPI_EXTERN int ObjToMultiSzEx(Tcl_Interp *interp,
                                Tcl_Obj *   listPtr,
                                LPCWSTR *   multiszPP,
                                Tcl_Size *,
                                MemLifo *lifoP);
TWAPI_EXTERN int ObjToMultiSzSWS(Tcl_Interp *interp,
                                 Tcl_Obj *   listPtr,
                                 LPCWSTR *   multiszPP,
                                 Tcl_Size *);
TWAPI_EXTERN Tcl_Obj *ObjFromMultiSz (LPCWSTR lpcw, Tcl_Size maxlen);
#define ObjFromMultiSz_MAX(lpcw) ObjFromMultiSz(lpcw, TCL_SIZE_MAX)
TWAPI_EXTERN Tcl_Obj *ObjFromEXPAND_SZW(WCHAR *ws);
TWAPI_EXTERN Tcl_Obj *ObjFromRegType(int regtype);
TWAPI_EXTERN Tcl_Obj *ObjFromRegValue(Tcl_Interp *interp, int regtype,
                         BYTE *bufP, Tcl_Size count);
TWAPI_EXTERN Tcl_Obj *ObjFromRegValueCooked(Tcl_Interp *interp, int regtype,
                               BYTE *bufP, Tcl_Size count);
TWAPI_EXTERN TCL_RESULT ObjToRegValueSWS(Tcl_Interp *   interp,
                                         Tcl_Obj *      typeObj,
                                         Tcl_Obj *      objP,
                                         TwapiRegValue *valueP);
TWAPI_EXTERN int ObjToRECT (Tcl_Interp *interp, Tcl_Obj *obj, RECT *rectP);
TWAPI_EXTERN int ObjToRECT_NULL (Tcl_Interp *interp, Tcl_Obj *obj, RECT **rectPP);
TWAPI_EXTERN Tcl_Obj *ObjFromRECT(RECT *rectP);
TWAPI_EXTERN Tcl_Obj *ObjFromPOINT(POINT *ptP);
TWAPI_EXTERN int ObjToPOINT (Tcl_Interp *interp, Tcl_Obj *obj, POINT *ptP);

/* GUIDs and UUIDs */
TWAPI_EXTERN Tcl_Obj *ObjFromGUID(const GUID *guidP);
TWAPI_EXTERN int ObjToGUID(Tcl_Interp *interp, Tcl_Obj *objP, GUID *guidP);
TWAPI_EXTERN int ObjToGUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, GUID **guidPP);
TWAPI_EXTERN Tcl_Obj *ObjFromUUID (UUID *uuidP);
TWAPI_EXTERN int ObjToUUID(Tcl_Interp *interp, Tcl_Obj *objP, UUID *uuidP);
TWAPI_EXTERN int ObjToUUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, UUID **uuidPP);
TWAPI_EXTERN Tcl_Obj *ObjFromLUID (const LUID *luidP);
TWAPI_EXTERN int ObjToLUID(Tcl_Interp *interp, Tcl_Obj *objP, LUID *luidP);
TWAPI_EXTERN int ObjToLUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, LUID **luidPP);

/* Network stuff */
TWAPI_EXTERN Tcl_Obj *IPAddrObjFromDWORD(DWORD addr);
TWAPI_EXTERN int IPAddrObjToDWORD(Tcl_Interp *interp, Tcl_Obj *objP, DWORD *addrP);
TWAPI_EXTERN Tcl_Obj *ObjFromIPv6Addr(const BYTE *addrP, DWORD scope_id);
TWAPI_EXTERN Tcl_Obj *ObjFromIP_ADDR_STRING (Tcl_Interp *, const IP_ADDR_STRING *ipaddrstrP);
TWAPI_EXTERN Tcl_Obj *ObjFromSOCKADDR_address(SOCKADDR *saP);
TWAPI_EXTERN Tcl_Obj *ObjFromSOCKADDR(SOCKADDR *saP);


/* Security stuff */
#define TWAPI_SID_LENGTH(sid_) GetSidLengthRequired((sid_)->SubAuthorityCount)
TWAPI_EXTERN TCL_RESULT TwapiValidateSID(Tcl_Interp *interp, SID *sidP, DWORD len);
TWAPI_EXTERN int ObjToPACLSWS(Tcl_Interp *interp, Tcl_Obj *aclObj, ACL **aclPP);
TWAPI_EXTERN int ObjToPSECURITY_ATTRIBUTESSWS(Tcl_Interp *interp, Tcl_Obj *secattrObj,
                                 SECURITY_ATTRIBUTES **secattrPP);
TWAPI_EXTERN int ObjToPSECURITY_DESCRIPTORSWS(Tcl_Interp *, Tcl_Obj *, SECURITY_DESCRIPTOR **secdPP);
TWAPI_EXTERN Tcl_Obj *ObjFromSECURITY_DESCRIPTOR(Tcl_Interp *, SECURITY_DESCRIPTOR *);
TWAPI_EXTERN TCL_RESULT ParsePSEC_WINNT_AUTH_IDENTITY (TwapiInterpContext *, Tcl_Obj *, SEC_WINNT_AUTH_IDENTITY_W **);
TWAPI_EXTERN void SecureZeroSEC_WINNT_AUTH_IDENTITY(SEC_WINNT_AUTH_IDENTITY_W *);
int TwapiFormatMessageHelper( Tcl_Interp *interp, DWORD dwFlags,
                              LPCVOID lpSource, DWORD dwMessageId,
                              DWORD dwLanguageId, Tcl_Size argc, LPCWSTR *argv );
TWAPI_EXTERN Tcl_Obj *ObjFromCREDENTIAL_ATTRIBUTEW(const CREDENTIAL_ATTRIBUTEW *attrP);
TWAPI_EXTERN Tcl_Obj *ObjFromCREDENTIALW(const CREDENTIALW *credsP);

/* LZMA */
TWAPI_EXTERN unsigned char *TwapiLzmaUncompressBuffer(Tcl_Interp *,
                                         unsigned char *buf,
                                         DWORD sz, DWORD *outsz);
TWAPI_EXTERN void TwapiLzmaFreeBuffer(unsigned char *buf);

/* Window message related */

/* Typedef for callbacks invoked from the hidden window proc. Parameters are
 * those for a window procedure except for an additional interp pointer (which
 * may be NULL)
 */
typedef LRESULT TwapiHiddenWindowCallbackProc(TwapiInterpContext *, LONG_PTR, HWND, UINT, WPARAM, LPARAM);
TWAPI_EXTERN int Twapi_CreateHiddenWindow(TwapiInterpContext *,
                             TwapiHiddenWindowCallbackProc *winProc,
                             LONG_PTR clientdata, HWND *winP);
TWAPI_EXTERN DWORD Twapi_SetWindowLongPtr(HWND hWnd, int nIndex, LONG_PTR lValue, LONG_PTR *retP);
TWAPI_EXTERN HWND Twapi_GetNotificationWindow(TwapiInterpContext *ticP);

/* General utility */
TWAPI_EXTERN TCL_RESULT Twapi_SourceResource(Tcl_Interp *interp, HANDLE dllH, const char *name, int try_file);
TWAPI_EXTERN Tcl_Obj *TwapiTwine(Tcl_Interp *interp, Tcl_Obj *first, Tcl_Obj *second);
TWAPI_EXTERN Tcl_Obj *TwapiTwineObjv(Tcl_Obj **first, Tcl_Obj **second, Tcl_Size n);

TWAPI_EXTERN void TwapiDebugOutputObj(Tcl_Obj *);
TWAPI_EXTERN void TwapiDebugOutput(char *s);
typedef int TwapiOneTimeInitFn(void *);
TWAPI_EXTERN int TwapiDoOneTimeInit(TwapiOneTimeInitState *stateP, TwapiOneTimeInitFn *, ClientData);
TWAPI_EXTERN int Twapi_AppendObjLog(Tcl_Interp *interp, Tcl_Obj *msgObj);
TWAPI_EXTERN int Twapi_AppendLog(Tcl_Interp *interp, WCHAR *msg);
TWAPI_EXTERN TwapiId Twapi_NewId();
TWAPI_EXTERN void TwapiGetDllVersion(char *dll, DLLVERSIONINFO *verP);

/* Interp context */
#define TwapiInterpContextRef(ticP_, incr_) InterlockedExchangeAdd(&(ticP_)->nrefs, (incr_))
TWAPI_EXTERN void TwapiInterpContextUnref(TwapiInterpContext *ticP, int);
TWAPI_EXTERN TwapiTls *Twapi_GetTls();
TWAPI_EXTERN int Twapi_AssignTlsSubSlot();
TWAPI_EXTERN Tcl_Obj *TwapiGetAtom(TwapiInterpContext *ticP, const char *key);
TWAPI_EXTERN void TwapiPurgeAtoms(TwapiInterpContext *ticP);
TWAPI_EXTERN void Twapi_MakeCallAlias(Tcl_Interp *interp, char *fn, char *callcmd, char *code);
TWAPI_EXTERN TCL_RESULT Twapi_CheckThreadedTcl(Tcl_Interp *interp);

/* Cstruct calls */
TWAPI_EXTERN TCL_RESULT ObjCastToCStruct(Tcl_Interp *, Tcl_Obj *, int allow_empty);
TWAPI_EXTERN TCL_RESULT TwapiCStructSize(Tcl_Interp *, Tcl_Obj *, int *);
TWAPI_EXTERN TCL_RESULT TwapiCStructParse (Tcl_Interp *interp, MemLifo *memlifoP,
                                      Tcl_Obj *csvalObj, DWORD flags, DWORD *sizeP, void **ppv);
/* ParseCStruct flags definitions */
#define CSTRUCT_ALLOW_NULL 0x1

TWAPI_EXTERN TCL_RESULT ObjFromCStruct(Tcl_Interp *interp, void *pv, int nbytes, Tcl_Obj *csObj, DWORD flags, Tcl_Obj **objPP);
/* ObjFromCStruct flags definitions */
#define CSTRUCT_RETURN_DICT 0x1

/*
 * Wrappers for memlifo based s/w stack
 * NOTE these all panic on fail since twapi inits the memlifos with the
 * PANIC_ON_FAIL flag.
 */
typedef MemLifo *SWStack;
typedef MemLifoMarkHandle SWSMark;
TWAPI_STATIC_INLINE SWStack SWS(void) {
    TwapiTls *tlsP = Twapi_GetTls();
    return &tlsP->memlifo;
}

TWAPI_STATIC_INLINE SWSMark SWSPushMark(void) { return MemLifoPushMark(SWS()); }
TWAPI_STATIC_INLINE void SWSPopMark(SWSMark mark) { MemLifoPopMark(mark); }
TWAPI_STATIC_INLINE void *SWSPushFrame(MemLifoSize sz, MemLifoSize *szP) {
    return MemLifoPushFrame(SWS(), sz, szP);
}
TWAPI_STATIC_INLINE void SWSPopFrame(void) { MemLifoPopFrame(SWS()); }
TWAPI_STATIC_INLINE void *SWSAlloc(MemLifoSize sz, MemLifoSize *actual_szP) {
    return MemLifoAlloc(SWS(), sz, actual_szP);
}


/* Module management */
TWAPI_EXTERN TwapiInterpContext *TwapiRegisterModule(
    Tcl_Interp *interp,
    HMODULE hmod,
    TwapiModuleDef *modP, /* MUST BE STATIC/NEVER DEALLOCATED */
    int new_context   /* If ! DEFAULT_TIC, new context is allocated.
                         If DEFAULT_TIC, the base package context is
                         used and caller must not access the
                         module-specific fields in the context. */
#define DEFAULT_TIC 0
#define NEW_TIC     1
    );

/* Pointer management */
TWAPI_EXTERN TCL_RESULT TwapiRegisterPointerTic(TwapiInterpContext *, const void *p, void *typetag);
TWAPI_EXTERN TCL_RESULT TwapiRegisterCountedPointerTic(TwapiInterpContext *, const void *p, void *typetag);
TWAPI_EXTERN int TwapiUnregisterPointerTic(TwapiInterpContext *, const void *p, void *typetag);
TWAPI_EXTERN int TwapiVerifyPointerTic(TwapiInterpContext *, const void *p, void *typetag);
TWAPI_EXTERN TCL_RESULT TwapiRegisterPointer(Tcl_Interp *interp, const void *p, void *typetag);
TWAPI_EXTERN TCL_RESULT TwapiRegisterCountedPointer(Tcl_Interp *interp, const void *p, void *typetag);
TWAPI_EXTERN int TwapiUnregisterPointer(Tcl_Interp *interp, const void *p, void *typetag);
TWAPI_EXTERN int TwapiVerifyPointer(Tcl_Interp *interp, const void *p, void *typetag);

TWAPI_EXTERN TCL_RESULT TwapiDictLookupString(Tcl_Interp *interp, Tcl_Obj *dictObj, const char *key, Tcl_Obj **objPP);

TWAPI_EXTERN Tcl_Obj *TwapiGetInstallDir(Tcl_Interp *interp, HANDLE dllH);

TWAPI_EXTERN BOOL CALLBACK Twapi_EnumWindowsCallback(HWND hwnd, LPARAM p_ctx);

TWAPI_EXTERN Tcl_Obj *TwapiLowerCaseObj(Tcl_Obj *objP);
TWAPI_EXTERN TCL_RESULT TwapiReturnNonnullHandle(Tcl_Interp *, HANDLE, char *typestr);

TWAPI_INLINE unsigned short swap2(unsigned short us)
{
    return (us << 8) | (us >> 8);
}

TWAPI_INLINE unsigned long swap4(unsigned long ul)
{
    ul = ((ul << 8) & 0xFF00FF00 ) | ((ul >> 8) & 0xFF00FF );
    return (ul << 16) | (ul >> 16);
}

TWAPI_INLINE unsigned __int64 swap8(unsigned __int64 ull)
{
    ull = ((ull << 8) & 0xFF00FF00FF00FF00 ) | ((ull >> 8) & 0x00FF00FF00FF00FF );
    ull = ((ull << 16) & 0xFFFF0000FFFF0000 ) | ((ull >> 16) & 0x0000FFFF0000FFFF );
    return (ull << 32) | (ull >> 32);
}

/*
 * Definitions used for defining Tcl commands dispatched via a function code
 * passed as ClientData
 */
struct fncode_dispatch_s {
    const char *command_name;
    int fncode;
};
#define DEFINE_FNCODE_CMD(fn_, code_)  {#fn_, code_}
TWAPI_EXTERN void TwapiDefineFncodeCmds(Tcl_Interp *, int, struct fncode_dispatch_s *, TwapiTclObjCmd *);

/* Commands that take a TwapiInterpContext as ClientData param */
struct tcl_dispatch_s {
    char *command_name;
    TwapiTclObjCmd *command_ptr;
};
#define DEFINE_TCL_CMD(fn_, cmdptr_) {#fn_, cmdptr_}
TWAPI_EXTERN void TwapiDefineTclCmds(Tcl_Interp *, int, struct tcl_dispatch_s *, ClientData clientdata);

/* Command that are defined as an alias */
struct alias_dispatch_s {
    const char *command_name;
    char *fncode;
};
#define DEFINE_ALIAS_CMD(fn_, code_)  {#fn_, #code_}
TWAPI_EXTERN void TwapiDefineAliasCmds(Tcl_Interp *, int, struct alias_dispatch_s *, const char *);





#endif // TWAPI_H
