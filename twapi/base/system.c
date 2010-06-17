/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include <twapi.h>



typedef NTSTATUS (WINAPI *NtQuerySystemInformation_t)(int, PVOID, ULONG, PULONG);
MAKE_DYNLOAD_FUNC(NtQuerySystemInformation, ntdll, NtQuerySystemInformation_t)


int TwapiReadMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    int restype;
    char *p;
    int offset;
    int len;
    TwapiResult result;

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(restype), GETVOIDP(p), GETINT(offset),
                     ARGUSEDEFAULT, GETINT(len),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    p += offset;
    // Note: restype matches func dispatch code in Twapi_CallObjCmd
    switch (restype) {
    case 10101:
        result.type = TRT_DWORD;
        result.value.ival = *(int UNALIGNED *) p;
        break;

    case 10102:
        result.type = TRT_BINARY;
        result.value.binary.p = p;
        result.value.binary.len = len;
        break;

    case 10103:
        result.type = TRT_CHARS;
        result.value.chars.str = p;
        result.value.chars.len = len;
        break;
        
    case 10104:
        result.type = TRT_UNICODE;
        result.value.unicode.str = (WCHAR *) p;
        result.value.unicode.len =  (len == -1) ? -1 : (len / sizeof(WCHAR));
        break;
        
    case 10105:
        result.type = TRT_ADDRESS_LITERAL;
        result.value.pval =  *(void* *) p;
        break;

    case 10106:
        result.type = TRT_WIDE;
        result.value.wide = *(Tcl_WideInt UNALIGNED *)p;
        break;

    default:
        Tcl_SetResult(interp, "Unknown result type passed to TwapiReadMemory", TCL_STATIC);
        return TCL_ERROR;
    }

    return TwapiSetResult(interp, &result);
}


int TwapiWriteMemory (Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    char *bufP;
    DWORD_PTR offset, buf_size;
    int func;
    int sz;
    int val;
    char *cp;
    WCHAR *wp;
    void  *pv;
    Tcl_WideInt wide;

    if (TwapiGetArgs(interp, objc, objv,
                     GETINT(func), GETVOIDP(bufP), GETDWORD_PTR(offset),
                     GETDWORD_PTR(buf_size), ARGSKIP,
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* Note the ARGSKIP ensures the objv[] in that position exists */
    switch (func) {
    case 10111: // Int
        if ((offset + sizeof(int)) > buf_size)
            goto overrun;
        if (Tcl_GetIntFromObj(interp, objv[4], &val) != TCL_OK)
            return TCL_ERROR;
        *(int UNALIGNED *)(offset + bufP) = val;
        break;
    case 10112: // Binary
        cp = Tcl_GetByteArrayFromObj(objv[4], &sz);
        if ((offset + sz) > buf_size)
            goto overrun;
        memmove(offset + bufP, cp, sz);
        break;
    case 10113: // Chars
        cp = Tcl_GetStringFromObj(objv[4], &sz);
        /* Note we also include the terminating null */
        if ((offset + sz + 1) > buf_size)
            goto overrun;
        memmove(offset + bufP, cp, sz+1);
        break;
    case 10114: // Unicode
        wp = Tcl_GetUnicodeFromObj(objv[4], &sz);
        /* Note we also include the terminating null */
        if ((offset + (sizeof(WCHAR)*(sz + 1))) > buf_size)
            goto overrun;
        memmove(offset + bufP, wp, sizeof(WCHAR)*(sz+1));
        break;
    case 10115: // Pointer
        if ((offset + sizeof(void*)) > buf_size)
            goto overrun;
        if (ObjToLPVOID(interp, objv[4], &pv) != TCL_OK)
            return TCL_ERROR;
        memmove(offset + bufP, &pv, sizeof(void*));
        break;
    case 10116:
        if ((offset + sizeof(Tcl_WideInt)) > buf_size)
            goto overrun;
        if (Tcl_GetWideIntFromObj(interp, objv[4], &wide) != TCL_OK)
            return TCL_ERROR;
        *(Tcl_WideInt UNALIGNED *)(offset + bufP) = wide;
        break;
    }        

    return TCL_OK;

overrun:
    return TwapiReturnTwapiError(interp, "Buffer too small.",
                                 TWAPI_BUFFER_OVERRUN);
}
