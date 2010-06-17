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
/* TBD #include "callback.h" */
#include "twapi_ddkdefs.h"

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
    static fntype Twapi_GetProc_ ## fnname (void)      \
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
    static FARPROC Twapi_GetProc_ ## fnname ##_ ## dllname (void) \
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
    static FARPROC Twapi_GetProc_ ## dllname ## _ ## ord (void) \
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
    Tcl_ListObjAppendElement((interp_),(listp_), StringObjFromLSA_UNICODE_STRING(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_UUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), StringObjFromUUID(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_LUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), StringObjFromLUID(&((structp_)->field_))); \
  } while (0)

#define Twapi_APPEND_PSID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_Obj *obj = StringObjFromSID((structp_)->field_); \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), obj ? obj : Tcl_NewStringObj("", 0)); \
  } while (0)

#define Twapi_APPEND_GUID_FIELD_TO_LIST(interp_, listp_, structp_, field_) \
  do { \
    Tcl_ListObjAppendElement((interp_), (listp_), STRING_LITERAL_OBJ( # field_)); \
    Tcl_ListObjAppendElement((interp_),(listp_), ObjFromGUID(&((structp_)->field_))); \
  } while (0)


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
    TRT_WIDE = 47,
} TwapiResultType;

struct swig_type_info;
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
 * n is a variable of type int
 * typestr - is any type string such as "HSERVICE" that indicates the type
 * fn - is a function to call to convert the value. The function
 *   should have the prototype TwapiGetArgFn
 */
#define ARGBOOL    'b'
#define ARGBIN     'B'
#define ARGINT     'i'
#define ARGWIDE    'I'
#define ARGDOUBLE  'd'
#define ARGOBJ     'o'
#define ARGDWORD_PTR 'P'
#define ARGASTR      's'
#define ARGASTRN     'S'
#define ARGWSTR     'u'
#define ARGWSTRN    'U'
#define ARGWORD     'w'
#define ARGPTR      'p'
#define ARGVAR     'v'
#define ARGVARWITHDEFAULT 'V'
#define ARGEND      0
#define ARGTERM     1
#define ARGSKIP     'x'
#define ARGUSEDEFAULT '?'
#define ARGNULLIFEMPTY 'E'
#define ARGNULLTOKEN 'N'

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
#define GETPTR(v, typestr) ARGPTR, &(v), #typestr
#define GETVOIDP(v)    ARGPTR, &(v), NULL
#define GETHANDLE(v)   GETVOIDP(v)
#define GETHANDLET(v, typestr) GETPTR(v, typestr)
#define GETVAR(v, fn)  ARGVAR, &(v), fn
#define GETVARWITHDEFAULT(v, fn)  ARGVARWITHDEFAULT, &(v), fn
#define GETGUID(v)     GETVAR(v, ObjToGUID)

typedef int (*TwapiGetArgsFn)(Tcl_Interp *, Tcl_Obj *, void *);


/*****************************************************************
 * Prototypes and globals
 *****************************************************************/

#ifdef __cplusplus
extern "C" {
#endif

/* GLOBALS */
extern OSVERSIONINFO TwapiOSVersionInfo;
extern GUID TwapiNullGuid;
extern struct TwapiTclVersion TclVersion;
extern int TclIsThreaded;
#define ERROR_IF_UNTHREADED(interp_)        \
    do {                                        \
        if (! TclIsThreaded) {                                          \
            if (interp_) Tcl_SetResult((interp_), "This command requires a threaded build of Tcl.", TCL_STATIC); \
            return TCL_ERROR;                                           \
        }                                                               \
    } while (0)


/* C - Tcl result and parameter conversion  */
int TwapiSetResult(Tcl_Interp *interp, TwapiResult *result);
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
int Twapi_RestoreResultErrorInfo (
    Tcl_Interp *interp,
    void       *savePtr
);
void Twapi_DiscardResultErrorInfo (Tcl_Interp *interp, void *savePtr);

/* Tcl_Obj manipulation and conversion - basic Windows types */

void Twapi_FreeNewTclObj(Tcl_Obj *objPtr);
Tcl_Obj *TwapiAppendObjArray(Tcl_Obj *resultObj, int objc, Tcl_Obj **objv,
                         char *join_string);
Tcl_Obj *ObjFromOpaque(void *pv, char *name);
#define ObjFromHANDLE(h) ObjFromOpaque((h), "HANDLE")

int ObjToOpaque(Tcl_Interp *interp, Tcl_Obj *obj, void **pvP, char *name);
#define ObjToLPVOID(interp, obj, vPP) ObjToOpaque((interp), (obj), (vPP), NULL)
#define ObjToHANDLE ObjToLPVOID

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



int ObjToArgvW(Tcl_Interp *interp, Tcl_Obj *objP, LPCWSTR *argv, int argc, int *argcP);
int ObjToArgvA(Tcl_Interp *interp, Tcl_Obj *objP, char **argv, int argc, int *argcP);
LPWSTR ObjToLPWSTR_NULL_IF_EMPTY(Tcl_Obj *objP);

#define NULL_TOKEN "__null__"
#define NULL_TOKEN_L L"__null__"
LPWSTR ObjToLPWSTR_WITH_NULL(Tcl_Obj *objP);

Tcl_Obj *ObjFromMODULEINFO(LPMODULEINFO miP);
Tcl_Obj *ObjFromPIDL(LPCITEMIDLIST pidl);
int ObjToPIDL(Tcl_Interp *interp, Tcl_Obj *objP, LPITEMIDLIST *idsPP);

Tcl_Obj *ObjFromIDispatch(void *p);
Tcl_Obj *ObjFromIUnknown(void *p);
int ObjToIDispatch(Tcl_Interp *interp, Tcl_Obj *obj, IDispatch **idisp);
int ObjToIUnknown(Tcl_Interp *interp, Tcl_Obj *obj, IUnknown **iunk);
int ObjToVT(Tcl_Interp *interp, Tcl_Obj *obj, VARTYPE *vtP);
Tcl_Obj *ObjFromBSTR (BSTR bstr);
BSTR ObjToBSTR (Tcl_Obj *);
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
int ObjToPSID(Tcl_Interp *interp, Tcl_Obj *obj, PSID *sidPP);
int ObjFromSID (Tcl_Interp *interp, SID *sidP, Tcl_Obj **objPP);
int ObjToSID_AND_ATTRIBUTES(Tcl_Interp *interp, Tcl_Obj *obj, SID_AND_ATTRIBUTES *sidattrP);
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
int ObjToLUID(Tcl_Interp *interp, Tcl_Obj *objP, LUID *luidP);
int ObjToLUID_NULL(Tcl_Interp *interp, Tcl_Obj *objP, LUID **luidPP);
Tcl_Obj *ObjFromLSA_UNICODE_STRING(const LSA_UNICODE_STRING *lsauniP);
int ObjToLSASTRINGARRAY(Tcl_Interp *interp, Tcl_Obj *obj,
                        LSA_UNICODE_STRING **arrayP, ULONG *countP);
int ObjToPTOKEN_PRIVILEGES(Tcl_Interp *interp,
                          Tcl_Obj *tokprivObj, TOKEN_PRIVILEGES **tokprivPP);
int ObjToSP_DEVINFO_DATA(Tcl_Interp *, Tcl_Obj *objP, SP_DEVINFO_DATA *sddP);
int ObjToSP_DEVINFO_DATA_NULL(Tcl_Interp *interp, Tcl_Obj *objP,
                              SP_DEVINFO_DATA **sddPP);
Tcl_Obj *ObjFromSP_DEVINFO_DATA(SP_DEVINFO_DATA *sddP);
int ObjToSP_DEVICE_INTERFACE_DATA(Tcl_Interp *interp, Tcl_Obj *objP,
                                  SP_DEVICE_INTERFACE_DATA *sdidP);
Tcl_Obj *ObjFromSP_DEVICE_INTERFACE_DATA(SP_DEVICE_INTERFACE_DATA *sdidP);
Tcl_Obj *ObjFromDISPLAY_DEVICE(DISPLAY_DEVICEW *ddP);
Tcl_Obj *ObjFromMONITORINFOEX(MONITORINFO *miP);

int Twapi_GetBestRoute(Tcl_Interp *interp, DWORD addr, DWORD addr2);
int Twapi_AllocateAndGetTcpExTableFromStack(Tcl_Interp *,BOOL sort,DWORD flags);
int Twapi_AllocateAndGetUdpExTableFromStack(Tcl_Interp *,BOOL sort,DWORD flags);
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



/* Network related */
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



int Twapi_GetWordFromObj(Tcl_Interp *interp, Tcl_Obj *obj, WORD *wordP);


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


/* UI and window related */

/* Typedef for callbacks invoked from the hidden window proc. Parameters are
 * those for a window procedure except for an additional interp pointer (which
 * may be NULL)
 */
typedef LRESULT (*TwapiHiddenWindowCallbackProc)(Tcl_Interp *, LONG_PTR, HWND, UINT, WPARAM, LPARAM);
int Twapi_CreateHiddenWindow(Tcl_Interp *interp,
                             TwapiHiddenWindowCallbackProc winProc,
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
int Twapi_IDispatch_InvokeObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]);
int Twapi_SHChangeNotify(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]);

int Twapi_ComEventSinkObjCmd(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]);

/* General utility functions */
int TwapiGlobCmp (const char *s, const char *pat);
int TwapiGlobCmpCase (const char *s, const char *pat);
Tcl_Obj *TwapiTwine(Tcl_Interp *interp, Tcl_Obj *first, Tcl_Obj *second);
/* TBD - replace all calls to malloc with this */
int Twapi_malloc(Tcl_Interp *interp, char *msg, size_t size, void **pp);
void DebugOutput(char *s);
int TwapiReadMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);
int TwapiWriteMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[]);


#ifdef __cplusplus
} // extern "C"
#endif



#endif // TWAPI_H
