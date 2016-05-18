/*
 * Copyright (c) 2016, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"

typedef struct WinChars {
    int nchars; /* Num of characters not counting terminating \0.
                   Always >=0 (i.e. -1 not used to indicate null termination)*/
    WCHAR chars[1]; /* Variable length array holding the string */
} WinChars;

static void DupWinCharsType(Tcl_Obj *srcP, Tcl_Obj *dstP);
static void FreeWinCharsType(Tcl_Obj *objP);
static void UpdateWinCharsTypeString(Tcl_Obj *objP);
struct Tcl_ObjType gWinCharsType = {
    "TwapiWinChars",
    FreeWinCharsType,
    DupWinCharsType,
    UpdateWinCharsTypeString,
    NULL,     /* jenglish says keep this NULL */
};

/* 
 * Allocate a intrep of appropriate size. Note nchars does not include \0
 * but the defined array size of 1 takes care of that.
 */
TWAPI_INLINE WinChars *WinCharsAlloc(int nchars) {
    return (WinChars *) ckalloc(sizeof(WinChars) + sizeof(WCHAR)*(nchars));
}

/* Free an internal rep */
TWAPI_INLINE void WinCharsFree(WinChars *rep) {
    ckfree((char *) rep);
}

/* Get a pointer to the internal rep from a Tcl_Obj */
TWAPI_INLINE WinChars *WinCharsGet(Tcl_Obj *objP) {
    return (WinChars *) objP->internalRep.twoPtrValue.ptr1;
}
        
/* Set the internal rep for a Tcl_Obj */
TWAPI_INLINE void WinCharsSet(Tcl_Obj *objP, WinChars *rep) {
    objP->typePtr = &gWinCharsType;
    objP->internalRep.twoPtrValue.ptr1 = (void *) rep ;
}

static WinChars *WinCharsNew(const WCHAR *wsP, int len)
{
    WinChars *rep;
    /*
     * Unlike Tcl's String object, we never directly operate on this
     * representation so we don't need to allocate extra space and so on.
     */
    if (wsP == NULL)
        len = 0;
    else if (len == -1)
        len = lstrlenW(wsP);
    rep = WinCharsAlloc(len); /* Will include space for terminating \0 */
    if (len)
        memmove(rep->chars, wsP, len * sizeof(WCHAR));
    rep->chars[len] = 0;
    rep->nchars = len;
    return rep;
}

static void FreeWinCharsType(Tcl_Obj *objP)
{
    TWAPI_ASSERT(objP->typePtr == &gWinCharsType);
    WinCharsFree(WinCharsGet(objP));
    objP->typePtr = NULL;
}

static void DupWinCharsType(Tcl_Obj *srcP, Tcl_Obj *dstP)
{
    WinChars *rep;

    TWAPI_ASSERT(srcP->typePtr == &gWinCharsType);
    rep = WinCharsGet(srcP);
    WinCharsSet(dstP, WinCharsNew(rep->chars, rep->nchars));
}

static void UpdateWinCharsTypeString(Tcl_Obj *objP)
{
    int nbytes;
    char *utf8;
    WinChars *rep;
    Tcl_DString ds;

    rep = WinCharsGet(objP);
    if (rep->nchars == 0) {
        objP->bytes = ckalloc(1);
        objP->bytes[0] = '\0';
        objP->length = 0;
        return;
    }

#if 0
    /* Disabled because XP does not support WC_ERR_INVALID_CHARS and
     *  will silently discard those characters
     */
    
    /* Note rep->nchars does not include terminating \0 so return values
     * will not include it either
     */
    nbytes = WideCharToMultiByte(
        CP_UTF8, /* CodePag */
        WC_ERR_INVALID_CHARS,       /* dwFlags */
        rep->chars,     /* lpWideCharStr */
        rep->nchars,  /* cchWideChar */
        NULL,     /* lpMultiByteStr */
        NULL,  /* cbMultiByte */
        NULL,    /* lpDefaultChar */
        NULL     /* lpUsedDefaultChar */
        );
    if (nbytes != 0) {
        utf8 = ckalloc(nbytes+1); /* One extra for terminating \0 */
        nbytes = WideCharToMultiByte(
            CP_UTF8, /* CodePag */
            WC_ERR_INVALID_CHARS,       /* dwFlags */
            rep->chars,     /* lpWideCharStr */
            rep->nchars,  /* cchWideChar */
            utf8,     /* lpMultiByteStr */
            nbytes,  /* cbMultiByte */
            NULL,    /* lpDefaultChar */
            NULL     /* lpUsedDefaultChar */
            );
        if (nbytes != 0) {
            utf8[nbytes] = '\0';
            objP->bytes = utf8;
            objP->length = nbytes;
            return;
        }
        ckfree(utf8);
    }
    /* 
     * Failures are possible because of invalid characters like embedded
     * nulls which are not illegal in Tcl, so do it the (slower) Tcl way then.
     */

#endif // #if 0
    
    Tcl_WinTCharToUtf(rep->chars, rep->nchars * sizeof(WCHAR), &ds);
    nbytes = Tcl_DStringLength(&ds);
    utf8 = ckalloc(nbytes+1);
    memmove(utf8, Tcl_DStringValue(&ds), nbytes);
    Tcl_DStringFree(&ds);
    utf8[nbytes] = '\0';

    objP->bytes = utf8;
    objP->length = nbytes;
}

TWAPI_EXTERN WCHAR *ObjToWinChars(Tcl_Obj *objP)
{
    WinChars *rep;
    Tcl_DString ds;
    int nbytes, len;
    char *utf8;
    
    if (objP->typePtr == &gWinCharsType)
        return WinCharsGet(objP)->chars;

    utf8 = ObjToStringN(objP, &nbytes);
    Tcl_WinUtfToTChar(utf8, nbytes, &ds);
    len = Tcl_DStringLength(&ds) / sizeof(WCHAR);
    rep = WinCharsNew((WCHAR *) Tcl_DStringValue(&ds), len);
    Tcl_DStringFree(&ds);
    
    /* Convert the passed object's internal rep */
    if (objP->typePtr && objP->typePtr->freeIntRepProc)
        objP->typePtr->freeIntRepProc(objP);
    WinCharsSet(objP, rep);
    return rep->chars;
}

TWAPI_EXTERN WCHAR *ObjToWinCharsN(Tcl_Obj *objP, int *lenP)
{
    WCHAR *wsP;
    wsP = ObjToWinChars(objP); /* Will convert as needed */
    if (lenP)
        *lenP = WinCharsGet(objP)->nchars;
    return wsP;
}

/* Identical to ObjToWinCharsN except length pointer type.
   Just to keep gcc happy without needing a cast */
TWAPI_EXTERN WCHAR *ObjToWinCharsNDW(Tcl_Obj *objP, DWORD *lenP)
{
    WCHAR *wsP;
    wsP = ObjToWinChars(objP); /* Will convert as needed */
    if (lenP)
        *lenP = WinCharsGet(objP)->nchars;
    return wsP;
}

TWAPI_EXTERN Tcl_Obj *ObjFromWinCharsN(const WCHAR *wsP, int nchars)
{
    Tcl_Obj *objP;
    WinChars *rep;
    
    if (wsP == NULL)
        return ObjFromEmptyString();
    
    if (! gBaseSettings.use_unicode_obj)
        return TwapiUtf8ObjFromWinChars(wsP, -1);
    
    rep = WinCharsNew(wsP, nchars);
    objP = Tcl_NewObj();
    Tcl_InvalidateStringRep(objP);
    WinCharsSet(objP, rep);
    return objP;
}

TWAPI_EXTERN Tcl_Obj *ObjFromWinChars(const WCHAR *wsP)
{
    return ObjFromWinCharsN(wsP, -1);
}

