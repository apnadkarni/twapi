#ifndef TWAPI_H
#define TWAPI_H

/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Enable prototype-less extern functions warnings even at warning level 1 */
#pragma warning (1 : 13)

#include <stddef.h>
#include <stdlib.h>
#include <stdarg.h>
#include <assert.h>
#include <winsock2.h>
#include <windows.h>
#define _WSPIAPI_COUNTOF // Needed to use VC++ 6.0 with Platform SDK 2003 SP1
#include <ws2tcpip.h>
#include <winsvc.h>
#include <psapi.h>
#include <pdhmsg.h>
#include <sddl.h>
#include <lmerr.h>
#include <lm.h>
#include <limits.h>
#include <errno.h>
#include <lmat.h>
#include <lm.h>
#include <pdh.h>         /* Include AFTER lm.h due to HLOG def conflict */
#include <sddl.h>        /* For SECURITY_DESCRIPTOR <-> string conversions */
#include <aclapi.h>
#include <winnetwk.h>
#include <iphlpapi.h>
#include <objidl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <locale.h>
#include <ntsecapi.h>
#include <wtsapi32.h>
#include <uxtheme.h>
#include <tmschema.h>
#include <intshcut.h>
#include <dispex.h>
#include <ocidl.h>
#include <mstask.h>
#include <dsgetdc.h>
#include <powrprof.h>
#include <winable.h>
#define SECURITY_WIN32 1
#include <security.h>
#include <userenv.h>
#include <setupapi.h>

#include "tcl.h"
#include "twapi_sdkdefs.h"
#include "twapi_ddkdefs.h"
#include "zlist.h"

#define TWAPI_TCL_NAMESPACE "twapi"

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
        if ((fnname ## dllname ## _F) == NULL) { \
            if ((dllname ## _H) == NULL) { \
                dllname ## _H = LoadLibraryA(#dllname ".dll"); \
            } \
            if (dllname ## _H) { \
                fnname ## dllname ## _F = \
                    (fntype) GetProcAddress( dllname ## _H, #fnname); \
            } \
        } \
 \
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
        if ((fnname ## dllname ## _F) == NULL) { \
            if ((dllname ## _H) == NULL) { \
                dllname ## _H = LoadLibraryA(#dllname ".dll"); \
            } \
            if (dllname ## _H) { \
                fnname ## dllname ## _F = \
                    (FARPROC) GetProcAddress( dllname ## _H, #fnname); \
            } \
        } \
 \
        return fnname ## dllname ## _F; \
    }

#define MAKE_DYNLOAD_FUNC_ORDINAL(ord, dllname) \
    FARPROC Twapi_GetProc_ ## dllname ## _ ## ord (void) \
    { \
        static HINSTANCE dllname ## _H; \
        static FARPROC   ord_ ## ord ## dllname ## _F; \
        if ((ord_ ## ord ## dllname ## _F) == NULL) { \
            if ((dllname ## _H) == NULL) { \
                dllname ## _H = LoadLibraryA(#dllname ".dll"); \
            } \
            if (dllname ## _H) { \
                ord_ ## ord ## dllname ## _F = \
                    (FARPROC) GetProcAddress( dllname ## _H, (LPCSTR) ord); \
            } \
        } \
 \
        return ord_ ## ord ## dllname ## _F; \
    }

/* 64 bittedness needs a BOOL version of the FARPROC def */
typedef BOOL (WINAPI *FARPROC_BOOL)();


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
 *************************

/*
 * String to use as first element of a errorCode
 */
#define TWAPI_WIN32_ERRORCODE_TOKEN "TWAPI_WIN32"  /* For Win32 errors */
#define TWAPI_ERRORCODE_TOKEN       "TWAPI"        /* For TWAPI errors */

/* TWAPI error codes - used with the Tcl error facility */
#define TWAPI_NO_ERROR       0
#define TWAPI_INVALID_ARGS   1
#define TWAPI_BUFFER_OVERRUN 2
#define TWAPI_EXTRA_ARGS     3
#define TWAPI_BAD_ARG_COUNT  4
#define TWAPI_INTERNAL_LIMIT 5
#define TWAPI_INVALID_OPTION 6
#define TWAPI_INVALID_FUNCTION_CODE 7
#define TWAPI_BUG            8

/**********************
 * Misc utility macros
 **********************/

/* Verify a Tcl_Obj is an integer/long and return error if not */
#define CHECK_INTEGER_OBJ(interp_, intvar_, objp_)       \
    do {                                                                \
        if (Tcl_GetIntFromObj((interp_), (objp_), &(intvar_)) != TCL_OK) \
            return TCL_ERROR;                                           \
    } while (0)

/* String equality test - check first char before calling strcmp */
#define STREQ(x, y) ( (((x)[0]) == ((y)[0])) && ! strcmp((x), (y)) )
#define STREQUN(x, y, n) \
    (((x)[0] == (y)[0]) && (strncmp(x, y, n) == 0))
#define WSTREQ(x, y) ( (((x)[0]) == ((y)[0])) && ! wcscmp((x), (y)) )

/* Make a pointer null if it points to empty element (generally used
   when we want to treat empty strings as null pointers */
#define NULLIFY_EMPTY(s_) if ((s_) && ((s_)[0] == 0)) (s_) = NULL
/* And the converse */
#define EMPTIFY_NULL(s_) if ((s_) == NULL) (s_) = L"";


/**********************************
 * Macros dealing with Tcl_Obj 
 **********************************/

/* Create a string obj from a string literal. */
#define STRING_LITERAL_OBJ(x) Tcl_NewStringObj(x, sizeof(x)-1)



/* Macros to append field name and values to a list */

/* Appends a struct DWORD field name and value pair to a given Tcl list object */
#define Twapi_APPEND_DWORD_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_), (listp_), Tcl_NewLongObj((DWORD)((structp_)->field_))); \
  } while (0)

/* Appends a struct ULONGLONG field name and value pair to a given Tcl list object */
#define Twapi_APPEND_ULONGLONG_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_), (listp_), Tcl_NewWideIntObj((ULONGLONG)(structp_)->field_)); \
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
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_), (listp_), Tcl_NewWideIntObj((structp_)->field_.QuadPart)); \
  } while (0)

/* Appends a struct Unicode field name and string pair to a Tcl list object */
#define Twapi_APPEND_LPCWSTR_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), Tcl_NewUnicodeObj(((structp_)->field_ ? (structp_)->field_ : L""), -1)); \
  } while (0)

/* Appends a struct char string field name and string pair to a Tcl list object */
#define Twapi_APPEND_LPCSTR_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), Tcl_NewStringObj(((structp_)->field_ ? (structp_)->field_ : ""), -1)); \
  } while (0)


/* Appends a struct Unicode field name and string pair to a Tcl list object */
#define Twapi_APPEND_LSA_UNICODE_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), ObjFromLSA_UNICODE_STRING(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_UUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), ObjFromUUID(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_LUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), ObjFromLUID(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_PSID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_Obj *obj = ObjFromSIDNoFail((structp_)->field_); \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), obj ? obj : Tcl_NewStringObj("", 0)); \
  } while (0)

#define Twapi_APPEND_GUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), ObjFromGUID(&((structp_)->field_))); \
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

#if (TCL_MAJOR_VERSION == 8) && (TCL_MINOR_VERSION == 4)
/*
 * We may want to sometimes use 8.5 functions when running against 8.5
 * even though we built against 8.4. This structure emulates the 8.5
 * stubs structure. We define explicit fields for the routines we use.
 * Note these should be called only after verifying at run time we
 * are running against 8.5!
 */
struct TwapiTcl85Stubs {
    int   magic;
    void *hooks;
    int (*fn[535])();
    void * (*tcl_SaveInterpState) (Tcl_Interp * interp, int status); /* 535 */
    int (*tcl_RestoreInterpState) (Tcl_Interp * interp, void * state); /* 536 */
    void (*tcl_DiscardInterpState) (void * state); /* 537 */
    int (*fn2[42])(); /* Totally 580 fns, (index 0-579) */
};
#define TWAPI_TCL85_STUB(fn_) (((struct TwapiTcl85Stubs *)tclStubsPtr)->fn_)

#endif


/*******************
 * Misc definitions
 *******************/
struct Twapi_EnumCtx {
    Tcl_Interp *interp;
    Tcl_Obj    *win_listobj;
};

/* Used to maintain context for common NetEnum* interfaces */
typedef struct {
    int    tag;  /* Type of entries in netbufP[] */
#define TWAPI_NETENUM_USERINFO              0
#define TWAPI_NETENUM_GROUPINFO             1
#define TWAPI_NETENUM_LOCALGROUPINFO        2
#define TWAPI_NETENUM_GROUPUSERSINFO        3
#define TWAPI_NETENUM_LOCALGROUPUSERSINFO   4
#define TWAPI_NETENUM_LOCALGROUPMEMBERSINFO 5
#define TWAPI_NETENUM_SESSIONINFO           6
#define TWAPI_NETENUM_FILEINFO              7
#define TWAPI_NETENUM_CONNECTIONINFO        8
#define TWAPI_NETENUM_SHAREINFO             9
#define TWAPI_NETENUM_USEINFO              10
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
    TRT_OBJV = 5,                    /* Array of Tcl_Obj * */
    TRT_RECT = 6,
    TRT_HANDLE = 7,
    TRT_CHARS = 8,                  /* char string */
    TRT_BINARY = 9,
    TRT_ADDRESS_LITERAL = 10,           /* Pointer or address */
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
    TRT_VALID_HANDLE = 26, // Unlike TRT_HANDLE, NULL is not an error
    TRT_GETLASTERROR = 27,   /* Set result as error code from GetLastError() */
    TRT_EXCEPTION_ON_WNET_ERROR = 28,
    TRT_DWORD_PTR = 29,
    TRT_LPVOID = 30,
    TRT_NONNULL_LPVOID = 31,
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
    TRT_OPAQUE = 40,
    TRT_TCL_RESULT = 41,             /* Interp result already set. Return ival
                                        field as status */
    TRT_NTSTATUS = 42,
    TRT_LSA_HANDLE = 43,
    TRT_SEC_WINNT_AUTH_IDENTITY = 44,
    TRT_HDEVINFO = 45,
    TRT_PIDL = 46,              /* Freed using CoTaskMemFree */
    TRT_WIDE = 47,              /* Tcl_WideInt */
    TRT_UNICODE_DYNAMIC = 48,     /* Unicode to be freed through TwapiFree */
    TRT_CHARS_DYNAMIC = 49,       /* Char string to be freed through TwapiFree */
} TwapiResultType;

typedef struct {
    TwapiResultType type;
    union {
        int ival;
        double dval;
        BOOL bval;
        DWORD_PTR dwp;
        Tcl_WideInt wide;
        void *pval;
        HANDLE hval;
        HWND hwin;
        struct {
            WCHAR *str;
            int    len;         /* len == -1 if str is null terminated */
        } unicode;
        struct {
            char *str;
            int    len;         /* len == -1 if str is null terminated */
        } chars;
        struct {
            char  *p;
            int    len;
        } binary;
        Tcl_Obj *obj;
        struct {
            int nobj;
            Tcl_Obj **objPP;
        } objv;
        RECT rect;
        POINT point;
        FILETIME filetime;
        UUID uuid;
        GUID guid;              /* Formatted differently from uuid */
        LUID luid;
        void *pv;
        struct {
            void *p;
            char *name;
        } ifc;
        struct {
            void *p;
            char *name;
        } opaque;
        VARIANT var;            /* Must VariantInit before use!! */
        LPOLESTR lpolestr; /* WCHAR string to be freed through CoTaskMemFree */
        SYSTEMTIME systime;
        LPITEMIDLIST pidl;
    } value;
} TwapiResult;




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
#define ARGBIN     'B'
#define ARGDOUBLE  'd'
#define ARGNULLIFEMPTY 'E'
#define ARGINT     'i'
#define ARGWIDE    'I'
#define ARGNULLTOKEN 'N'
#define ARGOBJ     'o'
#define ARGPTR      'p'
#define ARGDWORD_PTR 'P'
#define ARGAARGV   'r'
#define ARGWARGV   'R'
#define ARGASTR      's'
#define ARGASTRN     'S'
#define ARGWSTR     'u'
#define ARGWSTRN    'U'
#define ARGVAR     'v'
#define ARGVARWITHDEFAULT 'V'
#define ARGWORD     'w'
#define ARGSKIP     'x'
#define ARGUSEDEFAULT '?'

#define GETBOOL(v)    ARGBOOL, &(v)
#define GETBIN(v, n)  ARGBIN, &(v), &(n)
#define GETINT(v)     ARGINT, &(v)
#define GETWIDE(v)    ARGWIDE, &(v)
#define GETDOUBLE(v)  ARGDOUBLE, &(v)
#define GETOBJ(v)     ARGOBJ, &(v)
#define GETDWORD_PTR(v) ARGDWORD_PTR, &(v)
#define GETASTR(v)      ARGASTR, &(v)
#define GETASTRN(v, n)  ARGASTRN, &(v), &(n)
#define GETWSTR(v)     ARGWSTR, &(v)
#define GETWSTRN(v, n) ARGWSTRN, &(v), &(n)
#define GETNULLIFEMPTY(v) ARGNULLIFEMPTY, &(v)
#define GETNULLTOKEN(v) ARGNULLTOKEN, &(v)
#define GETWORD(v)     ARGWORD, &(v)
#define GETPTR(v, typesym) ARGPTR, &(v), #typesym
#define GETVOIDP(v)    ARGPTR, &(v), NULL
#define GETHANDLE(v)   GETVOIDP(v)
#define GETHANDLET(v, typesym) GETPTR(v, typesym)
#define GETHWND(v) GETHANDLET(v, HWND)
#define GETVAR(v, fn)  ARGVAR, &(v), fn
#define GETVARWITHDEFAULT(v, fn)  ARGVARWITHDEFAULT, &(v), fn
#define GETGUID(v)     GETVAR(v, ObjToGUID)
/* For GETAARGV/GETWARGV, v is of type char *v[n], or WCHAR *v[n] */
#define GETAARGV(v, I, n) ARGAARGV, (v), (I), &(n)
#define GETWARGV(v, I, n) ARGWARGV, (v), (I), &(n)

typedef int (*TwapiGetArgsFn)(Tcl_Interp *, Tcl_Obj *, void *);


/*
 * Forward decls
 */
typedef struct _TwapiInterpContext TwapiInterpContext;
ZLINK_CREATE_TYPEDEFS(TwapiInterpContext); 
ZLIST_CREATE_TYPEDEFS(TwapiInterpContext);
typedef struct _TwapiPendingCallback TwapiPendingCallback;
ZLINK_CREATE_TYPEDEFS(TwapiPendingCallback); 
ZLIST_CREATE_TYPEDEFS(TwapiPendingCallback);
typedef struct _TwapiThreadPoolRegisteredHandle TwapiThreadPoolRegisteredHandle;
typedef struct _TwapiDirectoryMonitorContext TwapiDirectoryMonitorContext;
ZLINK_CREATE_TYPEDEFS(TwapiDirectoryMonitorContext); 
ZLIST_CREATE_TYPEDEFS(TwapiDirectoryMonitorContext);

#if 0
/*
 * We need to keep track of handles that are being tracked by the 
 * thread pool so they can be released on interp deletion even if
 * the application code does not explicitly release them.
 */
ZLINK_CREATE_TYPEDEFS(TwapiThreadPoolRegisteredHandle); 
ZLIST_CREATE_TYPEDEFS(TwapiThreadPoolRegisteredHandle); 
typedef struct _TwapiThreadPoolRegisteredHandle {
    HANDLE handle;              /* Handle being waited on by thread pool */
    HANDLE tp_handle;           /* Corresponding handle returned by pool */
    ZLINK_DECL(TwapiThreadPoolRegisteredHandle); /* Link for tracking list */
} TwapiThreadPoolRegisteredHandle;
#endif


/*
 * For asynchronous notifications of events, there is a common framework
 * that passes events from the asynchronous handlers into the Tcl event
 * dispatch loop. From there, the framework calls a function of type
 * TwapiCallbackFn. This function should parse the event into a Tcl script
 * and return ERROR_SUCCESS on success. On failure, the function should
 * return an appropriate Win32 error code. *cbP, including status codes,
 * should be set appropriately. TBD - clarify/expand on this
 */
typedef int TwapiCallbackFn(struct _TwapiPendingCallback *cbP);

/*
 * Definitions relating to queue of pending callbacks. All pending callbacks
 * structure definitions must start with this structure as the header.
 */

/* Creates list link definitions */
typedef struct _TwapiPendingCallback {
    struct _TwapiInterpContext *ticP; /* Interpreter context */
    TwapiCallbackFn  *callback;     /* Function to call back - see notes
                                       in the TwapiCallbackFn typedef */
    LONG volatile     nrefs;       /* Ref count - use InterlockedIncrement */
    ZLINK_DECL(TwapiPendingCallback); /* Link for list */
    DWORD             status;         /* Return status - Win32 error code.
                                         Currently only used to send status
                                         back to callback initiator, not
                                         the other way */
    DWORD_PTR         clientdata;     /* For use by client code */
    HANDLE            completion_event;
    TwapiResult response;
} TwapiPendingCallback;

/*
 * TwapiInterpContext keeps track of a per-interpreter context.
 * This is allocated when twapi is loaded into an interpreter and
 * passed around as ClientData to most commands. It is reference counted
 * for deletion purposes and also placed on a global list for cleanup
 * purposes when a thread exits.
 */
typedef struct _TwapiInterpContext {
    ZLINK_DECL(TwapiInterpContext); /* Links all the contexts, primarily
                                       to track cleanup requirements */

    LONG volatile         nrefs;   /* Reference count for alloc/free */

    /* Back pointer to the associated interp. This must only be modified or
     * accessed in the Tcl thread. THIS IS IMPORTANT AS THERE IS NO
     * SYNCHORNIZATION PROTECTION AND Tcl interp's ARE NOT MEANT TO BE
     * ACCESSED FROM OTHER THREADS
     */
    Tcl_Interp *interp;

    Tcl_ThreadId thread;     /* Id of interp thread */

    /*
     * A single lock that is shared among multiple lists attached to this
     * structure as contention is expected to be low.
     */
    CRITICAL_SECTION lock;

    /* List of pending callbacks. Accessed controlled by the lock field */
    ZLIST_DECL(TwapiPendingCallback) pending;
    int              pending_suspended;       /* If true, do not pend events */
    
    /*
     * List of directory change monitor contexts.
     * Accessed controlled by the lock field.
     */
    ZLIST_DECL(TwapiDirectoryMonitorContext) directory_monitors;

    /* Tcl Async callback token. This is created on initialization
     * Note this CANNOT be left to be done when the event actually occurs.
     */
    Tcl_AsyncHandler async_handler;

    HWND          notification_win; /* Window used for various notifications */
    HWND          clipboard_win;    /* Window used for clipboard notifications */
    int           power_events_on; /* True-> send through power notifications */
    int           console_ctrl_hooked; /* True -> This interp is handling
                                          console ctrol signals */
    DWORD         device_notification_tid; /* device notification thread id */
    

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
    TwapiPendingCallback *pending_callback;
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

#ifdef __cplusplus
extern "C" {
#endif

/* GLOBALS */
extern OSVERSIONINFO gTwapiOSVersionInfo;
extern GUID gTwapiNullGuid;
extern struct TwapiTclVersion gTclVersion;
extern int gTclIsThreaded;
#define ERROR_IF_UNTHREADED(interp_)        \
    do {                                        \
        if (! gTclIsThreaded) {                                          \
            if (interp_) Tcl_SetResult((interp_), "This command requires a threaded build of Tcl.", TCL_STATIC); \
            return TCL_ERROR;                                           \
        }                                                               \
    } while (0)


/* Memory allocation */
void *TwapiAlloc(size_t sz);
void *TwapiAllocZero(size_t sz);
#define TwapiFree(p_) free(p_)
WCHAR *TwapiAllocWString(WCHAR *, int len);
WCHAR *TwapiAllocWStringFromObj(Tcl_Obj *, int *lenP);
char *TwapiAllocAString(char *, int len);
char *TwapiAllocAStringFromObj(Tcl_Obj *, int *lenP);

/* C - Tcl result and parameter conversion  */
int TwapiSetResult(Tcl_Interp *interp, TwapiResult *result);
void TwapiClearResult(TwapiResult *resultP);
int TwapiGetArgs(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[],
                 char fmt, ...);

/* errors.c */
int TwapiReturnSystemError(Tcl_Interp *interp);
int TwapiReturnTwapiError(Tcl_Interp *interp, char *msg, int code);
DWORD TwapiNTSTATUSToError(NTSTATUS status);
Tcl_Obj *Twapi_MakeTwapiErrorCodeObj(int err);
LPWSTR Twapi_MapWindowsErrorToString(DWORD err);
Tcl_Obj *Twapi_MakeWindowsErrorCodeObj(DWORD err, Tcl_Obj *);
int Twapi_AppendWNetError(Tcl_Interp *interp, unsigned long err);
LPWSTR Twapi_FormatMsgFromModule(DWORD err, HANDLE hModule);
int Twapi_AppendSystemError2(Tcl_Interp *, unsigned long err, Tcl_Obj *extra);
#define Twapi_AppendSystemError(interp_, error_) \
    Twapi_AppendSystemError2(interp_, error_, NULL)
int Twapi_GenerateWin32Error(Tcl_Interp *interp, DWORD error, char *msg);


/* Wrappers around saving/restoring Tcl state */
void *Twapi_SaveResultErrorInfo (Tcl_Interp *interp, int status);
int Twapi_RestoreResultErrorInfo (Tcl_Interp *interp, void *savePtr);
void Twapi_DiscardResultErrorInfo (Tcl_Interp *interp, void *savePtr);

/* Async handling related */
#define TwapiPendingCallbackRef(pcb_, incr_) InterlockedExchangeAdd(&(pcb_)->nrefs, (incr_))
void TwapiPendingCallbackUnref(TwapiPendingCallback *pcbP, int);
void TwapiPendingCallbackDelete(TwapiPendingCallback *pcbP);
TwapiPendingCallback *TwapiPendingCallbackNew(
    TwapiInterpContext *ticP, TwapiCallbackFn *callback, size_t sz);
int TwapiEnqueueCallback(
    TwapiInterpContext *ticP, TwapiPendingCallback *pcbP,
    int enqueue_method,
    int timeout,
    TwapiPendingCallback **responseP
    );
#define TWAPI_ENQUEUE_DIRECT 0
#define TWAPI_ENQUEUE_ASYNC  1
int TwapiEvalAndUpdateCallback(TwapiPendingCallback *cbP, int objc, Tcl_Obj *objv[], TwapiResultType response_type);

int Twapi_TclAsyncProc(TwapiInterpContext *ticP, Tcl_Interp *interp, int code);
#define TwapiInterpContextRef(ticP_, incr_) InterlockedExchangeAdd(&(ticP_)->nrefs, (incr_))
void TwapiInterpContextUnref(TwapiInterpContext *ticP, int);


/* Tcl_Obj manipulation and conversion - basic Windows types */

void Twapi_FreeNewTclObj(Tcl_Obj *objPtr);
Tcl_Obj *TwapiAppendObjArray(Tcl_Obj *resultObj, int objc, Tcl_Obj **objv,
                         char *join_string);
Tcl_Obj *ObjFromOpaque(void *pv, char *name);
#define ObjFromHANDLE(h) ObjFromOpaque((h), "HANDLE")
#define ObjFromHWND(h) ObjFromOpaque((h), "HWND")
#define ObjFromLPVOID(p) ObjFromOpaque((p), "void*")

int ObjToOpaque(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, char *name);
#define ObjToLPVOID(interp, obj, vPP) ObjToOpaque((interp), (obj), (vPP), NULL)
#define ObjToHANDLE ObjToLPVOID
#define ObjToHWND(ip_, obj_, p_) ObjToOpaque((ip_), (obj_), (p_), "HWND")

#ifdef _WIN64
#define ObjToDWORD_PTR        Tcl_GetWideIntFromObj
#define ObjFromDWORD_PTR(p_)  Tcl_NewWideIntObj((DWORD_PTR)(p_))
#else  // _WIN64
#define ObjToDWORD_PTR        Tcl_GetLongFromObj
#define ObjFromDWORD_PTR(p_)  Tcl_NewLongObj((DWORD_PTR)(p_))
#endif // _WIN64
#define ObjToULONG_PTR    ObjToDWORD_PTR
#define ObjFromULONG_PTR  ObjFromDWORD_PTR
#define ObjFromSIZE_T     ObjFromDWORD_PTR

#define ObjFromLARGE_INTEGER(val_) Tcl_NewWideIntObj((val_).QuadPart)
#define ObjFromULONGLONG(val_)     Tcl_NewWideIntObj(val_)

#define ObjFromUnicode(p_)    Tcl_NewUnicodeObj((p_), -1)
int ObjToWord(Tcl_Interp *interp, Tcl_Obj *obj, WORD *wordP);



int ObjToArgvW(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *argv, int argc, int *argcP);
int ObjToArgvA(Tcl_Interp *interp, Tcl_Obj *objP, char **argv, int argc, int *argcP);
LPWSTR ObjToLPWSTR_NULL_IF_EMPTY(Tcl_Obj *objP);

#define NULL_TOKEN "__null__"
#define NULL_TOKEN_L L"__null__"
LPWSTR ObjToLPWSTR_WITH_NULL(Tcl_Obj *objP);

Tcl_Obj *ObjFromMODULEINFO(LPMODULEINFO miP);
Tcl_Obj *ObjFromPIDL(LPCITEMIDLIST pidl);
int ObjToPIDL(Tcl_Interp *interp, Tcl_Obj *objP, LPITEMIDLIST *idsPP);

#define ObjFromIDispatch(p_) ObjFromOpaque((p_), "IDispatch")
#define ObjFromIUnknown(p_) ObjFromOpaque((p_), "IUnknown")
#define ObjToIDispatch(ip_, obj_, ifc_) \
    ObjToOpaque((ip_), (obj_), (ifc_), "IDispatch")
#define ObjToIUnknown(ip_, obj_, ifc_) \
    ObjToOpaque((ip_), (obj_), (ifc_), "IUnknown")

int ObjToVT(Tcl_Interp *interp, Tcl_Obj *obj, VARTYPE *vtP);
Tcl_Obj *ObjFromBSTR (BSTR bstr);
int ObjToBSTR (Tcl_Interp *, Tcl_Obj *, BSTR *);
int ObjToRangedInt(Tcl_Interp *, Tcl_Obj *obj, int low, int high, int *iP);
Tcl_Obj *ObjFromSYSTEMTIME(LPSYSTEMTIME timeP);
int ObjToSYSTEMTIME(Tcl_Interp *interp, Tcl_Obj *timeObj, LPSYSTEMTIME timeP);
Tcl_Obj *ObjFromFILETIME(FILETIME *ftimeP);
int ObjToFILETIME(Tcl_Interp *interp, Tcl_Obj *obj, FILETIME *cyP);
Tcl_Obj *ObjFromCY(const CY *cyP);
int ObjToCY(Tcl_Interp *interp, Tcl_Obj *obj, CY *cyP);
Tcl_Obj *ObjFromDECIMAL(DECIMAL *cyP);
int ObjToDECIMAL(Tcl_Interp *interp, Tcl_Obj *obj, DECIMAL *cyP);
Tcl_Obj *ObjFromVARIANT(VARIANT *varP, int value_only);

/* Note: the returned multiszPP must be free()'ed */
int ObjToMultiSz (Tcl_Interp *interp, Tcl_Obj *listPtr, LPCWSTR *multiszPP);
#define Twapi_ConvertTclListToMultiSz ObjToMultiSz
Tcl_Obj *ObjFromMultiSz (LPCWSTR lpcw, int maxlen);
#define ObjFromMultiSz_MAX(lpcw) ObjFromMultiSz(lpcw, INT_MAX)
Tcl_Obj *ObjFromRegValue(Tcl_Interp *interp, int regtype,
                         BYTE *bufP, int count);
int ObjToRECT (Tcl_Interp *interp, Tcl_Obj *obj, RECT *rectP);
int ObjToRECT_NULL (Tcl_Interp *interp, Tcl_Obj *obj, RECT **rectPP);
Tcl_Obj *ObjFromRECT(RECT *rectP);
Tcl_Obj *ObjFromPOINT(POINT *ptP);
int ObjToPOINT (Tcl_Interp *interp, Tcl_Obj *obj, POINT *ptP);
Tcl_Obj *ObjFromCONSOLE_SCREEN_BUFFER_INFO(
    Tcl_Interp *interp,
    const CONSOLE_SCREEN_BUFFER_INFO *csbiP
    );
int ObjToCOORD(Tcl_Interp *interp, Tcl_Obj *coordObj, COORD *coordP);
Tcl_Obj *ObjFromCOORD(Tcl_Interp *interp, const COORD *coordP);
int ObjToSMALL_RECT(Tcl_Interp *interp, Tcl_Obj *obj, SMALL_RECT *rectP);
int ObjToCHAR_INFO(Tcl_Interp *interp, Tcl_Obj *obj, CHAR_INFO *ciP);
Tcl_Obj *ObjFromSMALL_RECT(Tcl_Interp *interp, const SMALL_RECT *rectP);
int ObjToFLASHWINFO (Tcl_Interp *interp, Tcl_Obj *obj, FLASHWINFO *fwP);
Tcl_Obj *ObjFromWINDOWINFO (WINDOWINFO *wiP);
Tcl_Obj *ObjFromWINDOWPLACEMENT(WINDOWPLACEMENT *wpP);
int ObjToWINDOWPLACEMENT(Tcl_Interp *, Tcl_Obj *objP, WINDOWPLACEMENT *wpP);
Tcl_Obj *ObjFromGUID(GUID *guidP);
int ObjToGUID(Tcl_Interp *interp, Tcl_Obj *objP, GUID *guidP);
int ObjToGUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, GUID **guidPP);
Tcl_Obj *ObjFromSecHandle(SecHandle *shP);
int ObjToSecHandle(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle *shP);
int ObjToSecHandle_NULL(Tcl_Interp *interp, Tcl_Obj *obj, SecHandle **shPP);
Tcl_Obj *ObjFromSecPkgInfo(SecPkgInfoW *spiP);
void TwapiFreeSecBufferDesc(SecBufferDesc *sbdP);
int ObjToSecBufferDesc(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP, int readonly);
int ObjToSecBufferDescRO(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP);
int ObjToSecBufferDescRW(Tcl_Interp *interp, Tcl_Obj *obj, SecBufferDesc *sbdP);
Tcl_Obj *ObjFromSecBufferDesc(SecBufferDesc *sbdP);
Tcl_Obj *ObjFromUUID (UUID *uuidP);
int ObjToUUID(Tcl_Interp *interp, Tcl_Obj *objP, UUID *uuidP);
Tcl_Obj *ObjFromLUID (const LUID *luidP);
int ObjToSP_DEVINFO_DATA(Tcl_Interp *, Tcl_Obj *objP, SP_DEVINFO_DATA *sddP);
int ObjToSP_DEVINFO_DATA_NULL(Tcl_Interp *interp, Tcl_Obj *objP,
                              SP_DEVINFO_DATA **sddPP);
Tcl_Obj *ObjFromSP_DEVINFO_DATA(SP_DEVINFO_DATA *sddP);
int ObjToSP_DEVICE_INTERFACE_DATA(Tcl_Interp *interp, Tcl_Obj *objP,
                                  SP_DEVICE_INTERFACE_DATA *sdidP);
Tcl_Obj *ObjFromSP_DEVICE_INTERFACE_DATA(SP_DEVICE_INTERFACE_DATA *sdidP);
Tcl_Obj *ObjFromDISPLAY_DEVICE(DISPLAY_DEVICEW *ddP);
Tcl_Obj *ObjFromMONITORINFOEX(MONITORINFO *miP);
Tcl_Obj *ObjFromSYSTEM_POWER_STATUS(SYSTEM_POWER_STATUS *spsP);

Tcl_Obj *ObjFromMIB_IPNETROW(Tcl_Interp *interp, const MIB_IPNETROW *netrP);
Tcl_Obj *ObjFromMIB_IPNETTABLE(Tcl_Interp *interp, MIB_IPNETTABLE *nettP);
Tcl_Obj *ObjFromMIB_IPFORWARDROW(Tcl_Interp *interp, const MIB_IPFORWARDROW *ipfrP);
Tcl_Obj *ObjFromMIB_IPFORWARDTABLE(Tcl_Interp *interp, MIB_IPFORWARDTABLE *fwdP);
Tcl_Obj *ObjFromIP_ADAPTER_INDEX_MAP(Tcl_Interp *interp, IP_ADAPTER_INDEX_MAP *iaimP);
Tcl_Obj *ObjFromIP_INTERFACE_INFO(Tcl_Interp *interp, IP_INTERFACE_INFO *iiP);
Tcl_Obj *ObjFromMIB_TCPROW(Tcl_Interp *interp, const MIB_TCPROW *row, int size);
int ObjToMIB_TCPROW(Tcl_Interp *interp, Tcl_Obj *listObj, MIB_TCPROW *row);
Tcl_Obj *ObjFromIP_ADDR_STRING (Tcl_Interp *, const IP_ADDR_STRING *ipaddrstrP);
Tcl_Obj *ObjFromMIB_IPADDRROW(Tcl_Interp *interp, const MIB_IPADDRROW *iparP);
Tcl_Obj *ObjFromMIB_IPADDRTABLE(Tcl_Interp *interp, MIB_IPADDRTABLE *ipatP);
Tcl_Obj *ObjFromMIB_IFROW(Tcl_Interp *interp, const MIB_IFROW *ifrP);
Tcl_Obj *ObjFromMIB_IFTABLE(Tcl_Interp *interp, MIB_IFTABLE *iftP);
Tcl_Obj *ObjFromIP_ADAPTER_INDEX_MAP(Tcl_Interp *, IP_ADAPTER_INDEX_MAP *iaimP);
Tcl_Obj *ObjFromMIB_UDPROW(Tcl_Interp *interp, MIB_UDPROW *row, int size);
Tcl_Obj *ObjFromMIB_TCPTABLE(Tcl_Interp *interp, MIB_TCPTABLE *tab);
Tcl_Obj *ObjFromMIB_TCPTABLE_OWNER_PID(Tcl_Interp *i, MIB_TCPTABLE_OWNER_PID *tab);
Tcl_Obj *ObjFromMIB_TCPTABLE_OWNER_MODULE(Tcl_Interp *, MIB_TCPTABLE_OWNER_MODULE *tab);
Tcl_Obj *ObjFromMIB_UDPTABLE(Tcl_Interp *, MIB_UDPTABLE *tab);
Tcl_Obj *ObjFromMIB_UDPTABLE_OWNER_PID(Tcl_Interp *, MIB_UDPTABLE_OWNER_PID *tab);
Tcl_Obj *ObjFromMIB_UDPTABLE_OWNER_MODULE(Tcl_Interp *, MIB_UDPTABLE_OWNER_MODULE *tab);
Tcl_Obj *ObjFromTcpExTable(Tcl_Interp *interp, void *buf);
Tcl_Obj *ObjFromUdpExTable(Tcl_Interp *interp, void *buf);
int ObjToTASK_TRIGGER(Tcl_Interp *interp, Tcl_Obj *obj, TASK_TRIGGER *triggerP);
Tcl_Obj *ObjFromTASK_TRIGGER(TASK_TRIGGER *triggerP);

int ObjToPSID(Tcl_Interp *interp, Tcl_Obj *obj, PSID *sidPP);
int ObjFromSID (Tcl_Interp *interp, SID *sidP, Tcl_Obj **objPP);
Tcl_Obj *ObjFromSIDNoFail(SID *sidP);
int ObjToSID_AND_ATTRIBUTES(Tcl_Interp *interp, Tcl_Obj *obj, SID_AND_ATTRIBUTES *sidattrP);
Tcl_Obj *ObjFromSID_AND_ATTRIBUTES (Tcl_Interp *, const SID_AND_ATTRIBUTES *);
int ObjToPACL(Tcl_Interp *interp, Tcl_Obj *aclObj, ACL **aclPP);
int ObjToPSECURITY_ATTRIBUTES(Tcl_Interp *interp, Tcl_Obj *secattrObj,
                                 SECURITY_ATTRIBUTES **secattrPP);
void TwapiFreeSECURITY_ATTRIBUTES(SECURITY_ATTRIBUTES *secattrP);
void TwapiFreeSECURITY_DESCRIPTOR(SECURITY_DESCRIPTOR *secdP);
int ObjToPSECURITY_DESCRIPTOR(Tcl_Interp *, Tcl_Obj *, SECURITY_DESCRIPTOR **secdPP);
Tcl_Obj *ObjFromSECURITY_DESCRIPTOR(Tcl_Interp *, SECURITY_DESCRIPTOR *);
Tcl_Obj *ObjFromLUID_AND_ATTRIBUTES (Tcl_Interp *, const LUID_AND_ATTRIBUTES *);
int ObjToLUID_AND_ATTRIBUTES (Tcl_Interp *interp, Tcl_Obj *listobj,
                              LUID_AND_ATTRIBUTES *luidattrP);
int ObjToPTOKEN_PRIVILEGES(Tcl_Interp *interp,
                          Tcl_Obj *tokprivObj, TOKEN_PRIVILEGES **tokprivPP);
Tcl_Obj *ObjFromTOKEN_PRIVILEGES(Tcl_Interp *interp,
                                 const TOKEN_PRIVILEGES *tokprivP);
void TwapiFreeTOKEN_PRIVILEGES (TOKEN_PRIVILEGES *tokPrivP);
int ObjToLUID(Tcl_Interp *interp, Tcl_Obj *objP, LUID *luidP);
int ObjToLUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, LUID **luidPP);
Tcl_Obj *ObjFromLSA_UNICODE_STRING(const LSA_UNICODE_STRING *lsauniP);
int ObjToLSASTRINGARRAY(Tcl_Interp *interp, Tcl_Obj *obj,
                        LSA_UNICODE_STRING **arrayP, ULONG *countP);
Tcl_Obj *ObjFromACE (Tcl_Interp *interp, void *aceP);
int ObjToACE (Tcl_Interp *interp, Tcl_Obj *aceobj, void **acePP);
Tcl_Obj *ObjFromACL(Tcl_Interp *interp, ACL *aclP);
Tcl_Obj *ObjFromCONNECTION_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromUSE_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromSHARE_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromFILE_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromSESSION_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromUSER_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromGROUP_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromLOCALGROUP_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);
Tcl_Obj *ObjFromGROUP_USERS_INFO(Tcl_Interp *interp, LPBYTE infoP, DWORD level);

/* System related */
typedef NTSTATUS (WINAPI *NtQuerySystemInformation_t)(int, PVOID, ULONG, PULONG);
NtQuerySystemInformation_t Twapi_GetProc_NtQuerySystemInformation();
int TwapiFormatMessageHelper( Tcl_Interp *interp, DWORD dwFlags,
                              LPCVOID lpSource, DWORD dwMessageId,
                              DWORD dwLanguageId, int argc, LPCWSTR *argv );
int Twapi_LoadUserProfile(Tcl_Interp *interp, HANDLE hToken, DWORD flags,
                          LPWSTR username, LPWSTR profilepath);
BOOLEAN Twapi_Wow64DisableWow64FsRedirection(LPVOID *oldvalueP);
BOOLEAN Twapi_Wow64RevertWow64FsRedirection(LPVOID addr);
BOOLEAN Twapi_Wow64EnableWow64FsRedirection(BOOLEAN enable_redirection);
BOOL Twapi_IsWow64Process(HANDLE h, BOOL *is_wow64P);
int Twapi_GetSystemWow64Directory(Tcl_Interp *interp);
int Twapi_GetSystemInfo(Tcl_Interp *interp);
int Twapi_GlobalMemoryStatus(Tcl_Interp *interp);
int Twapi_SystemProcessorTimes(Tcl_Interp *interp);
int Twapi_SystemPagefileInformation(Tcl_Interp *interp);
int Twapi_TclGetChannelHandle(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_GetPrivateProfileSection(Tcl_Interp *interp, LPCWSTR app, LPCWSTR fn);
int Twapi_GetPrivateProfileSectionNames(Tcl_Interp *interp,LPCWSTR lpFileName);
int Twapi_GetVersionEx(Tcl_Interp *interp);
void TwapiGetDllVersion(char *dll, DLLVERSIONINFO *verP);

/* Shell stuff */
void TwapiFreePIDL(LPITEMIDLIST idlistP);
HRESULT Twapi_SHGetFolderPath(HWND hwndOwner, int nFolder, HANDLE hToken,
                          DWORD flags, WCHAR *pathbuf);
BOOL Twapi_SHObjectProperties(HWND hwnd, DWORD dwType,
                              LPCWSTR szObject, LPCWSTR szPage);

int Twapi_GetThemeColor(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_GetThemeFont(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int TwapiGetThemeDefine(Tcl_Interp *interp, char *name);
HTHEME Twapi_OpenThemeData(HWND win, LPCWSTR classes);
BOOL Twapi_IsThemeActive(void);
BOOL Twapi_IsAppThemed(void);
int Twapi_GetCurrentThemeName(Tcl_Interp *interp);
void Twapi_CloseThemeData(HTHEME themeH);
int Twapi_GetShellVersion(Tcl_Interp *interp);
int Twapi_ShellExecuteEx(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ReadShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_WriteShortcut(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ReadUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath);
int Twapi_WriteUrlShortcut(Tcl_Interp *interp, LPCWSTR linkPath, LPCWSTR url, DWORD flags);
int Twapi_InvokeUrlShortcut(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_SHFileOperation(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_VerQueryValue_FIXEDFILEINFO(Tcl_Interp *interp, TWAPI_FILEVERINFO * verP);
int Twapi_VerQueryValue_STRING(Tcl_Interp *interp, TWAPI_FILEVERINFO * verP,
                               LPCWSTR lang_and_cp, LPWSTR name);
int Twapi_VerQueryValue_TRANSLATIONS(Tcl_Interp *interp, TWAPI_FILEVERINFO * verP);
TWAPI_FILEVERINFO * Twapi_GetFileVersionInfo(LPWSTR path);
void Twapi_FreeFileVersionInfo(TWAPI_FILEVERINFO * verP);


/* Processes and threads */
int Twapi_GetProcessList(Tcl_Interp *interp, DWORD pid, int flags);
int Twapi_EnumProcesses (Tcl_Interp *interp);
int Twapi_EnumDeviceDrivers(Tcl_Interp *interp);
int Twapi_EnumProcessModules(Tcl_Interp *interp, HANDLE phandle);
int TwapiCreateProcessHelper(Tcl_Interp *interp, int func, int objc, Tcl_Obj * CONST objv[]);
int Twapi_NtQueryInformationProcessBasicInformation(Tcl_Interp *interp,
                                                    HANDLE processH);
int Twapi_NtQueryInformationThreadBasicInformation(Tcl_Interp *interp,
                                                   HANDLE threadH);
int Twapi_CommandLineToArgv(Tcl_Interp *interp, LPCWSTR cmdlineP);

/* Shares and LANMAN */
int Twapi_WNetGetUniversalName(Tcl_Interp *interp, LPCWSTR localpathP);
int Twapi_WNetGetUser(Tcl_Interp *interp, LPCWSTR  lpName);
int Twapi_NetScheduleJobEnum(Tcl_Interp *interp, LPCWSTR servername);
int Twapi_NetShareEnum(Tcl_Interp *interp, LPWSTR server_name);
int Twapi_NetUseGetInfo(Tcl_Interp *interp, LPWSTR UncServer, LPWSTR UseName, DWORD level);
int Twapi_NetShareCheck(Tcl_Interp *interp, LPWSTR server, LPWSTR device);
int Twapi_NetShareGetInfo(Tcl_Interp *interp, LPWSTR server,
                          LPWSTR netname, DWORD level);
int Twapi_NetShareSetInfo(Tcl_Interp *interp, LPWSTR server_name,
                          LPWSTR net_name, LPWSTR remark, DWORD  max_uses,
                          SECURITY_DESCRIPTOR *secd);
int Twapi_NetConnectionEnum(Tcl_Interp    *interp, LPWSTR server,
                            LPWSTR qualifier, DWORD level);
int Twapi_NetFileEnum(Tcl_Interp *interp, LPWSTR server, LPWSTR basepath,
                      LPWSTR user, DWORD level);
int Twapi_NetFileGetInfo(Tcl_Interp    *interp, LPWSTR server,
                         DWORD fileid, DWORD level);
int Twapi_NetSessionEnum(Tcl_Interp    *interp, LPWSTR server, LPWSTR client,
                         LPWSTR user, DWORD level);
int Twapi_NetSessionGetInfo(Tcl_Interp *interp, LPWSTR server,
                            LPWSTR client, LPWSTR user, DWORD level);
int Twapi_NetGetDCName(Tcl_Interp *interp, LPCWSTR server, LPCWSTR domain);
int Twapi_WNetGetResourceInformation(Tcl_Interp *interp, LPWSTR remoteName,
                                     LPWSTR provider, DWORD   resourcetype);
int Twapi_WNetUseConnection(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_NetShareAdd(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);


/* Security related */
int Twapi_LookupAccountName (Tcl_Interp *interp, LPCWSTR sysname, LPCWSTR name);
int Twapi_LookupAccountSid (Tcl_Interp *interp, LPCWSTR sysname, PSID sidP);
int Twapi_NetUserEnum(Tcl_Interp *interp, LPWSTR server_name, DWORD filter);
int Twapi_NetGroupEnum(Tcl_Interp *interp, LPWSTR server_name);
int Twapi_NetLocalGroupEnum(Tcl_Interp *interp, LPWSTR server_name);
int Twapi_NetUserGetGroups(Tcl_Interp *interp, LPWSTR server, LPWSTR user);
int Twapi_NetUserGetLocalGroups(Tcl_Interp *interp, LPWSTR server,
                                LPWSTR user, DWORD flags);
int Twapi_NetLocalGroupGetMembers(Tcl_Interp *interp, LPWSTR server, LPWSTR group);
int Twapi_NetGroupGetUsers(Tcl_Interp *interp, LPCWSTR server, LPCWSTR group);
int Twapi_NetUserGetInfo(Tcl_Interp *interp, LPCWSTR server,
                         LPCWSTR user, DWORD level);
int Twapi_NetGroupGetInfo(Tcl_Interp *interp, LPCWSTR server,
                          LPCWSTR group, DWORD level);
int Twapi_NetLocalGroupGetInfo(Tcl_Interp *interp, LPCWSTR server,
                               LPCWSTR group, DWORD level);
int Twapi_LsaEnumerateLogonSessions(Tcl_Interp *interp);
int Twapi_LsaQueryInformationPolicy (Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]
);
int Twapi_InitializeSecurityDescriptor(Tcl_Interp *interp);
int Twapi_GetSecurityInfo(Tcl_Interp *interp, HANDLE h, int type, int wanted_fields);
int Twapi_GetNamedSecurityInfo (Tcl_Interp *, LPWSTR name,int type,int wanted);
int Twapi_LsaGetLogonSessionData(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);

int TwapiReturnNetEnum(Tcl_Interp *interp, TwapiNetEnumContext *necP);
int Twapi_NetUseEnum(Tcl_Interp *interp);
int Twapi_NetUserSetInfoDWORD(int fun, LPCWSTR server, LPCWSTR user, DWORD dw);
int Twapi_NetUserSetInfoLPWSTR(int fun, LPCWSTR server, LPCWSTR user, LPWSTR s);
int Twapi_NetUserAdd(Tcl_Interp *interp, LPCWSTR servername, LPWSTR name,
                     LPWSTR password, DWORD priv, LPWSTR home_dir,
                     LPWSTR comment, DWORD flags, LPWSTR script_path);
int Twapi_GetTokenInformation(Tcl_Interp *interp, HANDLE tokenH, int tclass);
int Twapi_SetTokenPrimaryGroup(HANDLE tokenH, PSID sidP);
int Twapi_SetTokenVirtualizationEnabled(HANDLE tokenH, DWORD enabled);
int Twapi_AdjustTokenPrivileges(Tcl_Interp *interp, HANDLE tokenH,
                                BOOL disableAll, TOKEN_PRIVILEGES *tokprivP);
DWORD Twapi_PrivilegeCheck(HANDLE tokenH, const TOKEN_PRIVILEGES *tokprivP,
                           int all_required, int *resultP);

int Twapi_LsaEnumerateAccountRights(Tcl_Interp *interp,
                                    LSA_HANDLE PolicyHandle, PSID AccountSid);
int Twapi_LsaEnumerateAccountsWithUserRight(
    Tcl_Interp *, LSA_HANDLE PolicyHandle, LSA_UNICODE_STRING *UserRights);


/* Crypto API */
int Twapi_EnumerateSecurityPackages(Tcl_Interp *interp);
int Twapi_InitializeSecurityContext(
    Tcl_Interp *interp,
    SecHandle *credentialP,
    SecHandle *contextP,
    LPWSTR     targetP,
    ULONG      contextreq,
    ULONG      reserved1,
    ULONG      targetdatarep,
    SecBufferDesc *sbd_inP,
    ULONG     reserved2);
int Twapi_AcceptSecurityContext(Tcl_Interp *interp, SecHandle *credentialP,
                                SecHandle *contextP, SecBufferDesc *sbd_inP,
                                ULONG contextreq, ULONG targetdatarep);
int Twapi_QueryContextAttributes(Tcl_Interp *interp, SecHandle *INPUT,
                                 ULONG attr);
SEC_WINNT_AUTH_IDENTITY_W *Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY (
    LPCWSTR user, LPCWSTR domain, LPCWSTR password);
void Twapi_Free_SEC_WINNT_AUTH_IDENTITY (SEC_WINNT_AUTH_IDENTITY_W *swaiP);
int Twapi_MakeSignature(Tcl_Interp *interp, SecHandle *INPUT,
                        ULONG qop, int BINLEN, void *BINDATA, ULONG seqnum);
int Twapi_EncryptMessage(Tcl_Interp *interp, SecHandle *INPUT,
                        ULONG qop, int BINLEN, void *BINDATA, ULONG seqnum);
int Twapi_CryptGenRandom(Tcl_Interp *interp, HCRYPTPROV hProv, DWORD dwLen);

/* Device related */
int Twapi_EnumDisplayMonitors(Tcl_Interp *interp, HDC hdc, const RECT *rectP);
int Twapi_QueryDosDevice(Tcl_Interp *interp, LPCWSTR lpDeviceName);
int Twapi_SetupDiGetDeviceRegistryProperty(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_SetupDiGetDeviceInterfaceDetail(Tcl_Interp *, int objc, Tcl_Obj *CONST objv[]);
int Twapi_SetupDiClassGuidsFromNameEx(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_RegisterDeviceNotification(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[]);
int Twapi_UnregisterDeviceNotification(TwapiInterpContext *ticP, int id);

/* File and disk related */
int TwapiFirstVolume(Tcl_Interp *interp, LPCWSTR path);
int TwapiNextVolume(Tcl_Interp *interp, int treat_as_mountpoint, HANDLE hFindVolume);
int Twapi_GetVolumeInformation(Tcl_Interp *interp, LPCWSTR path);
int Twapi_GetDiskFreeSpaceEx(Tcl_Interp *interp, LPCWSTR dir);
int Twapi_GetFileType(Tcl_Interp *interp, HANDLE h);

/* PDH */
void TwapiPdhRestoreLocale(void);
int Twapi_PdhParseCounterPath(Tcl_Interp *interp, LPCWSTR buf, DWORD dwFlags);
int Twapi_PdhGetFormattedCounterValue(Tcl_Interp *, HANDLE hCtr, DWORD fmt);
int Twapi_PdhLookupPerfNameByIndex(Tcl_Interp *,  LPCWSTR machine, DWORD ctr);
int Twapi_PdhMakeCounterPath (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_PdhBrowseCounters(Tcl_Interp *interp);
int Twapi_PdhEnumObjects(Tcl_Interp *interp, LPCWSTR source, LPCWSTR machine,
                         DWORD  dwDetailLevel, BOOL bRefresh);
int Twapi_PdhEnumObjectItems(Tcl_Interp *, LPCWSTR source, LPCWSTR machine,
                              LPCWSTR objname, DWORD detail, DWORD dwFlags);


/* Printers */
int Twapi_EnumPrinters_Level4(Tcl_Interp *interp, DWORD flags);

/* Console related */
int Twapi_ReadConsole(Tcl_Interp *interp, HANDLE conh, unsigned int numchars);

/* Clipboard related */
int Twapi_EnumClipboardFormats(Tcl_Interp *interp);
int Twapi_MonitorClipboardStart(TwapiInterpContext *ticP);
int Twapi_MonitorClipboardStop(TwapiInterpContext *ticP);
int Twapi_StartConsoleEventNotifier(TwapiInterpContext *ticP);
int Twapi_StopConsoleEventNotifier(TwapiInterpContext *ticP);


/* ADSI related */
int Twapi_DsGetDcName(Tcl_Interp *interp, LPCWSTR systemnameP,
                      LPCWSTR domainnameP, GUID *guidP,
                      LPCWSTR sitenameP, ULONG flags);

/* Network related */
int Twapi_GetNetworkParams(Tcl_Interp *interp);
int Twapi_GetAdaptersInfo(Tcl_Interp *interp);
int Twapi_GetInterfaceInfo(Tcl_Interp *interp);
int Twapi_GetPerAdapterInfo(Tcl_Interp *interp, int adapter_index);
int Twapi_GetIfEntry(Tcl_Interp *interp, int if_index);
int Twapi_GetIfTable(Tcl_Interp *interp, int sort);
int Twapi_GetIpAddrTable(Tcl_Interp *interp, int sort);
int Twapi_GetIpNetTable(Tcl_Interp *interp, int sort);
int Twapi_GetIpForwardTable(Tcl_Interp *interp, int sort);

int Twapi_GetBestRoute(Tcl_Interp *interp, DWORD addr, DWORD addr2);
int Twapi_AllocateAndGetTcpExTableFromStack(Tcl_Interp *,BOOL sort,DWORD flags);
int Twapi_AllocateAndGetUdpExTableFromStack(Tcl_Interp *,BOOL sort,DWORD flags);
int Twapi_FormatExtendedTcpTable(Tcl_Interp *, void *buf, int family, int table_class);
int Twapi_FormatExtendedUdpTable(Tcl_Interp *, void *buf, int family, int table_class);
int Twapi_GetExtendedTcpTable(Tcl_Interp *interp, void *buf, DWORD buf_sz,
                              BOOL sorted, ULONG family, int table_class);
int Twapi_GetExtendedUdpTable(Tcl_Interp *interp, void *buf, DWORD buf_sz,
                              BOOL sorted, ULONG family, int table_class);
Tcl_Obj *ObjFromIP_ADAPTER_INFO(Tcl_Interp *interp, IP_ADAPTER_INFO *ainfoP);
Tcl_Obj *ObjFromIP_ADAPTER_INFO_table(Tcl_Interp *, IP_ADAPTER_INFO *ainfoP);
int ObjToSOCKADDR_IN(Tcl_Interp *, Tcl_Obj *objP, struct sockaddr_in *sinP);
int Twapi_GetNameInfo(Tcl_Interp *, const struct sockaddr_in* saP, int flags);
int Twapi_GetAddrInfo(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ResolveAddressAsync(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ResolveHostnameAsync(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[]);

/* NLS */

int Twapi_GetNumberFormat(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_GetCurrencyFormat(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);



/* COM stuff */

int TwapiMakeVariantParam(
    Tcl_Interp *interp,
    Tcl_Obj *paramDescriptorP,
    VARIANT *varP,
    VARIANT *refvarP,
    USHORT  *paramflagsP,
    Tcl_Obj *valueObj
    );
void TwapiClearVariantParam(Tcl_Interp *interp, VARIANT *varP);

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


int Twapi_AppendCOMError(Tcl_Interp *interp, HRESULT hr, ISupportErrorInfo *sei, REFIID iid);


/* WTS */

int Twapi_WTSEnumerateSessions(Tcl_Interp *interp, HANDLE wtsH);
int Twapi_WTSEnumerateProcesses(Tcl_Interp *interp, HANDLE wtsH);
int Twapi_WTSQuerySessionInformation(Tcl_Interp *interp,  HANDLE wtsH,
                                     DWORD  sess_id, WTS_INFO_CLASS info_class);

/* Services */
int Twapi_CreateService(Tcl_Interp *interp, int objc,
                        Tcl_Obj *CONST objv[]);
int Twapi_StartService(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ChangeServiceConfig(Tcl_Interp *interp, int objc,
                              Tcl_Obj *CONST objv[]);
int Twapi_EnumServicesStatusEx(Tcl_Interp *interp, SC_HANDLE hService,
                               int infolevel, DWORD dwServiceType,
                               DWORD dwServiceState,  LPCWSTR groupname);
int Twapi_EnumDependentServices(Tcl_Interp *interp, SC_HANDLE hService, DWORD state);
int Twapi_QueryServiceStatusEx(Tcl_Interp *interp, SC_HANDLE h, SC_STATUS_TYPE level);
int Twapi_QueryServiceConfig(Tcl_Interp *interp, SC_HANDLE hService);
int Twapi_EnumServicesStatus(Tcl_Interp *interp, SC_HANDLE hService,
                             DWORD dwServiceType, DWORD dwServiceState);
int Twapi_BecomeAService(TwapiInterpContext *, int objc, Tcl_Obj *CONST objv[]);

int Twapi_SetServiceStatus(TwapiInterpContext *, int objc, Tcl_Obj *CONST objv[]);


/* Task scheduler related */
int Twapi_IEnumWorkItems_Next(Tcl_Interp *interp,
        IEnumWorkItems *ewiP, unsigned long count);
int Twapi_IScheduledWorkItem_GetRunTimes(Tcl_Interp *interp,
        IScheduledWorkItem *swiP, SYSTEMTIME *, SYSTEMTIME *, WORD );
int Twapi_IScheduledWorkItem_GetWorkItemData(Tcl_Interp *interp,
                                             IScheduledWorkItem *swiP);

/* Event log */
BOOL Twapi_IsEventLogFull(HANDLE hEventLog, int *fullP);
int Twapi_ReportEvent(Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int Twapi_ReadEventLog(Tcl_Interp *, HANDLE evlH, DWORD  flags, DWORD offset);


/* UI and window related */
int Twapi_SendUnicode(Tcl_Interp *interp, Tcl_Obj *input_obj);
int Twapi_SendInput(Tcl_Interp *interp, Tcl_Obj *input_obj);
Tcl_Obj *ObjFromLOGFONTW(LOGFONTW *lfP);
int Twapi_EnumWindowStations(Tcl_Interp *interp);
int Twapi_EnumWindows(Tcl_Interp *interp);
int Twapi_BlockInput(Tcl_Interp *interp, BOOL block);
int Twapi_WaitForInputIdle(Tcl_Interp *, HANDLE hProcess, DWORD dwMillisecs);
int Twapi_GetGUIThreadInfo(Tcl_Interp *interp, DWORD idThread);
int Twapi_EnumDesktops(Tcl_Interp *interp, HWINSTA hwinsta);
int Twapi_EnumDesktopWindows(Tcl_Interp *interp, HDESK desk_handle);
int Twapi_EnumChildWindows(Tcl_Interp *interp, HWND parent_handle);
DWORD Twapi_SetWindowLongPtr(HWND hWnd, int nIndex, LONG_PTR lValue, LONG_PTR *retP);
int Twapi_UnregisterHotKey(TwapiInterpContext *ticP, int id);
int Twapi_RegisterHotKey(TwapiInterpContext *ticP, int id, UINT modifiers, UINT vk);
LRESULT TwapiHotkeyHandler(TwapiInterpContext *, UINT, WPARAM, LPARAM);
HWND TwapiGetNotificationWindow(TwapiInterpContext *ticP);

/* Power management */
LRESULT TwapiPowerHandler(TwapiInterpContext *, UINT, WPARAM, LPARAM);
int Twapi_PowerNotifyStart(TwapiInterpContext *ticP);
int Twapi_PowerNotifyStop(TwapiInterpContext *ticP);

/* Typedef for callbacks invoked from the hidden window proc. Parameters are
 * those for a window procedure except for an additional interp pointer (which
 * may be NULL)
 */
typedef LRESULT TwapiHiddenWindowCallbackProc(TwapiInterpContext *, LONG_PTR, HWND, UINT, WPARAM, LPARAM);
int Twapi_CreateHiddenWindow(TwapiInterpContext *,
                             TwapiHiddenWindowCallbackProc *winProc,
                             LONG_PTR clientdata, HWND *winP);



/* Built-in commands */

typedef int TwapiTclObjCmd(
    ClientData dummy,           /* Not used. */
    Tcl_Interp *interp,         /* Current interpreter. */
    int objc,                   /* Number of arguments. */
    Tcl_Obj *CONST objv[]);     /* Argument objects. */

TwapiTclObjCmd Twapi_ParseargsObjCmd;
TwapiTclObjCmd Twapi_TryObjCmd;
TwapiTclObjCmd Twapi_KlGetObjCmd;
TwapiTclObjCmd Twapi_TwineObjCmd;
TwapiTclObjCmd Twapi_RecordArrayObjCmd;
TwapiTclObjCmd Twapi_GetTwapiBuildInfo;
TwapiTclObjCmd Twapi_IDispatch_InvokeObjCmd;
TwapiTclObjCmd Twapi_ComEventSinkObjCmd;
TwapiTclObjCmd Twapi_SHChangeNotify;

/* Dispatcher routines */
    int Twapi_InitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP);
TwapiTclObjCmd Twapi_CallObjCmd;
TwapiTclObjCmd Twapi_CallUObjCmd;
TwapiTclObjCmd Twapi_CallSObjCmd;
TwapiTclObjCmd Twapi_CallHObjCmd;
TwapiTclObjCmd Twapi_CallHSUObjCmd;
TwapiTclObjCmd Twapi_CallSSSDObjCmd;
TwapiTclObjCmd Twapi_CallWObjCmd;
TwapiTclObjCmd Twapi_CallWUObjCmd;
TwapiTclObjCmd Twapi_CallPSIDObjCmd;
TwapiTclObjCmd Twapi_CallNetEnumObjCmd;
TwapiTclObjCmd Twapi_CallPdhObjCmd;
TwapiTclObjCmd Twapi_CallCOMObjCmd;


/* General utility functions */
int TwapiGlobCmp (const char *s, const char *pat);
int TwapiGlobCmpCase (const char *s, const char *pat);
Tcl_Obj *TwapiTwine(Tcl_Interp *interp, Tcl_Obj *first, Tcl_Obj *second);
/* TBD - replace all calls to malloc with this */
int Twapi_malloc(Tcl_Interp *interp, char *msg, size_t size, void **pp);
void DebugOutput(char *s);
int TwapiReadMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int TwapiWriteMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);

typedef int TwapiOneTimeInitFn(void *);
int TwapiDoOneTimeInit(TwapiOneTimeInitState *stateP, TwapiOneTimeInitFn *, ClientData);


#ifdef __cplusplus
} // extern "C"
#endif



#endif // TWAPI_H
