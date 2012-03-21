/* 
 * Copyright (c) 2004-2012, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

static TwapiInterpContext * volatile console_control_ticP;

static int ObjToCHAR_INFO(Tcl_Interp *interp, Tcl_Obj *obj, CHAR_INFO *ciP)
{
    Tcl_Obj **objv;
    int       objc;
    int i;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) != TCL_OK ||
        objc != 2 ||
        Tcl_GetIntFromObj(interp, objv[1], &i) != TCL_OK) {
        Tcl_SetResult(interp, "Invalid CHAR_INFO structure.", TCL_STATIC);
        return TCL_ERROR;
    }

    ciP->Char.UnicodeChar = * Tcl_GetUnicode(objv[0]);
    ciP->Attributes = (WORD) i;
    return TCL_OK;
}

static int ObjToCOORD(Tcl_Interp *interp, Tcl_Obj *coordObj, COORD *coordP)
{
    int objc, x, y;
    Tcl_Obj **objv;
    if (Tcl_ListObjGetElements(interp, coordObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if (objc != 2)
        goto format_error;
    
    if ((Tcl_GetIntFromObj(interp, objv[0], &x) != TCL_OK) ||
        (Tcl_GetIntFromObj(interp, objv[1], &y) != TCL_OK))
        goto format_error;

    if (x < 0 || x > 32767 || y < 0 || y > 32767)
        goto format_error;

    coordP->X = (SHORT) x;
    coordP->Y = (SHORT) y;

    return TCL_OK;

 format_error:
    if (interp)
        Tcl_SetResult(interp,
                      "Invalid Console coordinates format. Should have exactly 2 integer elements between 0 and 65535",
                      TCL_STATIC);
    return TCL_ERROR;
}

static Tcl_Obj *ObjFromCOORD(
    Tcl_Interp *interp,
    const COORD *coordP
)
{
    Tcl_Obj *objv[2];

    objv[0] = Tcl_NewIntObj(coordP->X);
    objv[1] = Tcl_NewIntObj(coordP->Y);

    return Tcl_NewListObj(2, objv);
}

static int ObjToSMALL_RECT(Tcl_Interp *interp, Tcl_Obj *obj, SMALL_RECT *rectP)
{
    Tcl_Obj **objv;
    int       objc;
    int l, t, r, b;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 4) {
        Tcl_SetResult(interp, "Need to specify exactly 4 integers for a SMALL_RECT structure", TCL_STATIC);
        return TCL_ERROR;
    }
    if ((Tcl_GetIntFromObj(interp, objv[0], &l) != TCL_OK) ||
        (Tcl_GetIntFromObj(interp, objv[1], &t) != TCL_OK) ||
        (Tcl_GetIntFromObj(interp, objv[2], &r) != TCL_OK) ||
        (Tcl_GetIntFromObj(interp, objv[3], &b) != TCL_OK)) {
        return TCL_ERROR;
    }
    rectP->Left   = (SHORT) l;
    rectP->Top    = (SHORT) t;
    rectP->Right  = (SHORT) r;
    rectP->Bottom = (SHORT) b;
    return TCL_OK;
}

static Tcl_Obj *ObjFromSMALL_RECT(
    Tcl_Interp *interp,
    const SMALL_RECT *rectP
)
{
    Tcl_Obj *objv[4];

    objv[0] = Tcl_NewIntObj(rectP->Left);
    objv[1] = Tcl_NewIntObj(rectP->Top);
    objv[2] = Tcl_NewIntObj(rectP->Right);
    objv[3] = Tcl_NewIntObj(rectP->Bottom);

    return Tcl_NewListObj(4, objv);
}

static Tcl_Obj *ObjFromCONSOLE_SCREEN_BUFFER_INFO(
    Tcl_Interp *interp,
    const CONSOLE_SCREEN_BUFFER_INFO *csbiP
)
{
    Tcl_Obj *objv[5];

    objv[0] = ObjFromCOORD(interp, &csbiP->dwSize);
    objv[1] = ObjFromCOORD(interp, &csbiP->dwCursorPosition);
    objv[2] = Tcl_NewIntObj(csbiP->wAttributes);
    objv[3] = ObjFromSMALL_RECT(interp, &csbiP->srWindow);
    objv[4] = ObjFromCOORD(interp, &csbiP->dwMaximumWindowSize);

    return Tcl_NewListObj(5, objv);
}


static int Twapi_ReadConsole(TwapiInterpContext *ticP, HANDLE conh, unsigned int numchars)
{
    WCHAR *bufP;
    DWORD  len;
    int status;

    bufP = MemLifoPushFrame(&ticP->memlifo, sizeof(WCHAR) * numchars, NULL);

    if (ReadConsoleW(conh, bufP, numchars, &len, NULL)) {
        Tcl_SetObjResult(ticP->interp, ObjFromUnicodeN(bufP, len));
        status = TCL_OK;
    } else {
        TwapiReturnSystemError(ticP->interp);
        status = TCL_ERROR;
    }

    MemLifoPopFrame(&ticP->memlifo);
    return status;
}

static int TwapiConsoleCtrlCallbackFn(TwapiCallback *cbP)
{
    char *event_str;
    Tcl_Obj *objs[3];

    /*
     * Note - event if interp is gone, we let TwapiEvalAndUpdateCallback
     * deal with it.
     */

    switch (cbP->clientdata) {
    case CTRL_C_EVENT:
        event_str = "ctrl-c";
        break;
    case CTRL_BREAK_EVENT:
        event_str = "ctrl-break";
        break;
    case CTRL_CLOSE_EVENT:
        event_str = "close";
        break;
    case CTRL_LOGOFF_EVENT:
        event_str = "logoff";
        break;
    case CTRL_SHUTDOWN_EVENT:
        event_str = "shutdown";
        break;
    default:
        // Unknown event type
        return ERROR_INVALID_PARAMETER;
    }

    objs[0] = Tcl_NewStringObj(TWAPI_TCL_NAMESPACE "::_console_ctrl_handler", -1);
    objs[1] = Tcl_NewStringObj(event_str, -1);
    return TwapiEvalAndUpdateCallback(cbP, 2, objs, TRT_BOOL);
}

/* Directly called by Windows in a separate thread */
static BOOL WINAPI TwapiConsoleCtrlHandler(DWORD ctrl)
{
    TwapiCallback *cbP;
    BOOL handled = FALSE;

    /* TBD - there is a race here? */
    if (console_control_ticP == NULL)
        return FALSE;

    cbP = TwapiCallbackNew(
        console_control_ticP, TwapiConsoleCtrlCallbackFn, sizeof(*cbP));

    cbP->clientdata = ctrl;
    if (TwapiEnqueueCallback(console_control_ticP,
                             cbP,
                             TWAPI_ENQUEUE_DIRECT,
                             100, /* Timeout (ms) */
                             &cbP)
        == ERROR_SUCCESS) {

        if (cbP && cbP->response.type == TRT_BOOL)
            handled = cbP->response.value.bval;
    }

    if (cbP)
        TwapiCallbackUnref((TwapiCallback *)cbP, 1);

    return handled;
}

static int Twapi_StartConsoleEventNotifier(TwapiInterpContext *ticP)
{
    void *pv;

    ERROR_IF_UNTHREADED(ticP->interp);
    pv = InterlockedCompareExchangePointer(&console_control_ticP,
                                           ticP, NULL);
    if (pv) {
        Tcl_SetResult(ticP->interp, "Console control handler is already set.", TCL_STATIC);
        return TCL_ERROR;
    }

    if (SetConsoleCtrlHandler(TwapiConsoleCtrlHandler, TRUE)) {
        ticP->module.data.ival = 1; /* Indicates console is hooked in this interp */
        TwapiInterpContextRef(ticP, 1);
        return TCL_OK;
    }
    else {
        InterlockedExchangePointer(&console_control_ticP, NULL);
        return TwapiReturnSystemError(ticP->interp);
    }
}

    
static int Twapi_StopConsoleEventNotifier(TwapiInterpContext *ticP)
{
    void *pv;
    pv = InterlockedCompareExchangePointer(&console_control_ticP,
                                           NULL, ticP);
    if (pv != (void*) ticP) {
        Tcl_SetResult(ticP->interp, "Console control handler not set by this interpreter.", TCL_STATIC);
        return TCL_ERROR;
    }
    SetConsoleCtrlHandler(TwapiConsoleCtrlHandler, FALSE);
    ticP->module.data.ival = 0; /* Indicates console is not hooked in this interp */

    TwapiInterpContextUnref(ticP, 1);
    return TCL_OK;
}

static void TwapiConsoleCleanup(TwapiInterpContext *ticP)
{
    if (ticP->module.data.ival)
        Twapi_StopConsoleEventNotifier(ticP);
}

static int Twapi_ConsoleCallObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiResult result;
    int func;
    union {
        CONSOLE_SCREEN_BUFFER_INFO csbi;
        WCHAR buf[MAX_PATH+1];
        SMALL_RECT srect[2];
    } u;
    COORD coord;
    CHAR_INFO chinfo;
    DWORD dw, dw2, dw3;
    LPWSTR s;
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h;
    WORD w;

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 100) {
        /* Functions taking no arguments */
        if (objc != 2)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = AllocConsole();
            break;
        case 2:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FreeConsole();
            break;
        case 3:
            result.type = TRT_DWORD;
            result.value.uval = GetConsoleCP();
            break;
        case 4:
            result.type = TRT_DWORD;
            result.value.uval = GetConsoleOutputCP();
            break;
        case 5:
            result.type = GetNumberOfConsoleMouseButtons(&result.value.uval) ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 6:
            /* Note : GetLastError == 0 means title is empty string */
            if ((result.value.unicode.len = GetConsoleTitleW(u.buf, sizeof(u.buf)/sizeof(u.buf[0]))) != 0 || GetLastError() == 0) {
                result.type = TRT_UNICODE;
                result.value.unicode.str = u.buf;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 7:
            result.value.hwin = GetConsoleWindow();
            result.type = result.value.hwin ? TRT_HWND : TRT_GETLASTERROR;
            break;
        case 8:
            return Twapi_StopConsoleEventNotifier(ticP);
        case 9:
            return Twapi_StartConsoleEventNotifier(ticP);
        }
    } else if (func < 200) {
        /* First arg integer, maybe two more */
        if (objc < 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_INTEGER_OBJ(interp, dw, objv[2]);
        switch (func) {
        case 101:
            result.value.ival = SetConsoleCP(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 102:
            result.value.ival = SetConsoleOutputCP(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 103:
            if (objc != 4)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            CHECK_INTEGER_OBJ(interp, dw2, objv[3]);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = GenerateConsoleCtrlEvent(dw, dw2);
            break;
        }
    } else if (func < 300) {
        /* A single handle arguments */
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToHANDLE(interp, objv[2], &h) != TCL_OK)
                return TCL_ERROR;
        switch (func) {
        case 201:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = FlushConsoleInputBuffer(h);
            break;
        case 202:
            result.type = GetConsoleMode(h, &result.value.uval)
                ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 203:
            if (GetConsoleScreenBufferInfo(h, &u.csbi) == 0)
                result.type = TRT_GETLASTERROR;
            else {
                Tcl_SetObjResult(interp, ObjFromCONSOLE_SCREEN_BUFFER_INFO(interp, &u.csbi));
                return TCL_OK;
            }
            break;
        case 204:
            coord = GetLargestConsoleWindowSize(h);
            Tcl_SetObjResult(interp, ObjFromCOORD(interp, &coord));
            return TCL_OK;
        case 205:
            result.type = GetNumberOfConsoleInputEvents(h, &result.value.uval) ? TRT_DWORD : TRT_GETLASTERROR;
            break;
        case 206:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleActiveScreenBuffer(h);
            break;
            
        }
    } else {
        /* Free-for-all - each func responsible for checking arguments */
        /* At least one arg present */
        if (objc < 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1001: // SetConsoleWindowInfo
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETBOOL(dw), GETVAR(u.srect[0], ObjToSMALL_RECT),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleWindowInfo(h, dw, &u.srect[0]);
            break;
        case 1002: // FillConsoleOutputAttribute
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWORD(w), GETINT(dw),
                             GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            if (FillConsoleOutputAttribute(h, w, dw, coord, &result.value.uval))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1003: // ScrollConsoleScreenBuffer
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h),
                             GETVAR(u.srect[0], ObjToSMALL_RECT),
                             GETVAR(u.srect[1], ObjToSMALL_RECT),
                             GETVAR(coord, ObjToCOORD),
                             GETVAR(chinfo, ObjToCHAR_INFO),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = ScrollConsoleScreenBufferW(
                h, &u.srect[0], &u.srect[1], coord, &chinfo);
            break;
        case 1004:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWSTRN(s, dw),
                             GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (WriteConsoleOutputCharacterW(h, s, dw, coord, &result.value.uval))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;    
            break;
        case 1005: // SetConsoleCursorPosition
        case 1006: // SetConsoleScreenBufferSize
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival =
                (func == 1005 ? SetConsoleCursorPosition : SetConsoleScreenBufferSize)
                (h, coord);
            break;
        case 1007: // CreateConsoleScreenBuffer
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETINT(dw),
                             GETINT(dw2),
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTES),
                             GETINT(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            result.value.hval = CreateConsoleScreenBuffer(dw, dw2, secattrP, dw3, NULL);
            TwapiFreeSECURITY_ATTRIBUTES(secattrP);
            break;
        case 1008:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleTitleW(Tcl_GetUnicode(objv[2]));
            break;
        case 1009:
        case 1010:
        case 1011:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETINT(dw), ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (func == 1011)
                return Twapi_ReadConsole(ticP, h, dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival =
                func == 1009 ? SetConsoleMode(h, dw) : SetConsoleTextAttribute(h, (WORD) dw);
            break;
        case 1012:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWSTR(s), ARGUSEDEFAULT, GETINT(dw), ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (WriteConsoleW(h, s, dw, &result.value.uval, NULL))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1013:
            if (TwapiGetArgs(interp, objc-2, objv+2,
                             GETHANDLE(h), GETWSTR(s), GETINT(dw), ARGSKIP, ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (ObjToCOORD(interp, objv[5], &coord) != TCL_OK)
                return TCL_ERROR;
            if (FillConsoleOutputCharacterW(h, s[0], dw, coord, &result.value.uval))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;
        }
    }

    return TwapiSetResult(interp, &result);
}


static int Twapi_ConsoleInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::ConsoleCall", Twapi_ConsoleCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
#define CALL_(fn_, call_, code_)                                         \
    do {                                                                \
        Twapi_MakeCallAlias(interp, "twapi::" #fn_, "twapi::Console" #call_, # code_); \
    } while (0);

    CALL_(AllocConsole, Call, 1);
    CALL_(FreeConsole, Call, 2);
    CALL_(GetConsoleCP, Call, 3);
    CALL_(GetConsoleOutputCP, Call, 4);
    CALL_(GetNumberOfConsoleMouseButtons, Call, 5);
    CALL_(GetConsoleTitle, Call, 6);
    CALL_(GetConsoleWindow, Call, 7);
    CALL_(Twapi_StopConsoleEventNotifier, Call, 8);
    CALL_(Twapi_StartConsoleEventNotifier, Call, 9);

    CALL_(SetConsoleCP, Call, 101);
    CALL_(SetConsoleOutputCP, Call, 102);
    CALL_(GenerateConsoleCtrlEvent, Call, 103);

    CALL_(FlushConsoleInputBuffer, Call, 201);
    CALL_(GetConsoleMode, Call, 202);
    CALL_(GetConsoleScreenBufferInfo, Call, 203);
    CALL_(GetLargestConsoleWindowSize, Call, 204);
    CALL_(GetNumberOfConsoleInputEvents, Call, 205);
    CALL_(SetConsoleActiveScreenBuffer, Call, 206);

    CALL_(SetConsoleWindowInfo, Call, 1001);
    CALL_(FillConsoleOutputAttribute, Call, 1002);
    CALL_(ScrollConsoleScreenBuffer, Call, 1003);
    CALL_(WriteConsoleOutputCharacter, Call, 1004);
    CALL_(SetConsoleCursorPosition, Call, 1005);
    CALL_(SetConsoleScreenBufferSize, Call, 1006);
    CALL_(CreateConsoleScreenBuffer, Call, 1007);
    CALL_(SetConsoleTitle, Call, 1008);
    CALL_(SetConsoleMode, Call, 1009);
    CALL_(SetConsoleTextAttribute, Call, 1010);
    CALL_(ReadConsole, Call, 1011);
    CALL_(WriteConsole, Call, 1012);
    CALL_(FillConsoleOutputCharacter, Call, 1013);

#undef CALL_

    return TCL_OK;
}


#ifndef TWAPI_SINGLE_MODULE
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gModuleHandle = hmod;
    return TRUE;
}
#endif

/* Main entry point */
#ifndef TWAPI_SINGLE_MODULE
__declspec(dllexport) 
#endif
int Twapi_console_Init(Tcl_Interp *interp)
{
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    /* Note ticP->module.data.ival is initialized to 0 which we use
     * as an indicator whether this interp has hooked ctrl-c or not
     */
    return Twapi_ModuleInit(interp, WLITERAL(MODULENAME), MODULE_HANDLE,
                            Twapi_ConsoleInitCalls, TwapiConsoleCleanup) ? TCL_OK : TCL_ERROR;
}

