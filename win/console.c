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

#ifndef MODULENAME
#define MODULENAME "twapi_console"
#endif

#if __GNUC__
static void * volatile console_control_ticP;
#else
static TwapiInterpContext * volatile console_control_ticP;
#endif
static int ObjToCHAR_INFO(Tcl_Interp *interp, Tcl_Obj *obj, CHAR_INFO *ciP)
{
    Tcl_Obj **objv;
    Tcl_Size objc;
    int i;

    if (ObjGetElements(interp, obj, &objc, &objv) != TCL_OK ||
        objc != 2 ||
        ObjToInt(interp, objv[1], &i) != TCL_OK) {
        ObjSetStaticResult(interp, "Invalid CHAR_INFO structure.");
        return TCL_ERROR;
    }

    ciP->Char.UnicodeChar = * ObjToWinChars(objv[0]);
    ciP->Attributes = (WORD) i;
    return TCL_OK;
}

static int ObjToCOORD(Tcl_Interp *interp, Tcl_Obj *coordObj, COORD *coordP)
{
    int x, y;
    Tcl_Size  objc;
    Tcl_Obj **objv;
    if (ObjGetElements(interp, coordObj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;
    if (objc != 2)
        goto format_error;
    
    if ((ObjToInt(interp, objv[0], &x) != TCL_OK) ||
        (ObjToInt(interp, objv[1], &y) != TCL_OK))
        goto format_error;

    if (x < 0 || x > 32767 || y < 0 || y > 32767)
        goto format_error;

    coordP->X = (SHORT) x;
    coordP->Y = (SHORT) y;

    return TCL_OK;

 format_error:
    if (interp)
        ObjSetStaticResult(interp, "Invalid Console coordinates format.");
    return TCL_ERROR;
}

static Tcl_Obj *ObjFromCOORD(const COORD *coordP)
{
    Tcl_Obj *objv[2];

    objv[0] = ObjFromInt(coordP->X);
    objv[1] = ObjFromInt(coordP->Y);

    return ObjNewList(2, objv);
}

static int ObjToSMALL_RECT(Tcl_Interp *interp, Tcl_Obj *obj, SMALL_RECT *rectP)
{
    Tcl_Obj **objv;
    Tcl_Size  objc;
    int l, t, r, b;

    if (ObjGetElements(interp, obj, &objc, &objv) == TCL_ERROR) {
        return TCL_ERROR;
    }

    if (objc != 4) {
        ObjSetStaticResult(interp, "Invalid SMALL_RECT structure");
        return TCL_ERROR;
    }
    if ((ObjToInt(interp, objv[0], &l) != TCL_OK) ||
        (ObjToInt(interp, objv[1], &t) != TCL_OK) ||
        (ObjToInt(interp, objv[2], &r) != TCL_OK) ||
        (ObjToInt(interp, objv[3], &b) != TCL_OK)) {
        return TCL_ERROR;
    }
    rectP->Left   = (SHORT) l;
    rectP->Top    = (SHORT) t;
    rectP->Right  = (SHORT) r;
    rectP->Bottom = (SHORT) b;
    return TCL_OK;
}

static Tcl_Obj *ObjFromSMALL_RECT(const SMALL_RECT *rectP)
{
    Tcl_Obj *objv[4];

    objv[0] = ObjFromInt(rectP->Left);
    objv[1] = ObjFromInt(rectP->Top);
    objv[2] = ObjFromInt(rectP->Right);
    objv[3] = ObjFromInt(rectP->Bottom);

    return ObjNewList(4, objv);
}

static Tcl_Obj *ObjFromCONSOLE_SCREEN_BUFFER_INFO(
    const CONSOLE_SCREEN_BUFFER_INFO *csbiP
)
{
    Tcl_Obj *objv[5];

    objv[0] = ObjFromCOORD(&csbiP->dwSize);
    objv[1] = ObjFromCOORD(&csbiP->dwCursorPosition);
    objv[2] = ObjFromInt(csbiP->wAttributes);
    objv[3] = ObjFromSMALL_RECT(&csbiP->srWindow);
    objv[4] = ObjFromCOORD(&csbiP->dwMaximumWindowSize);

    return ObjNewList(5, objv);
}

static Tcl_Obj *ObjFromINPUT_RECORD(const INPUT_RECORD *recP)
{
    Tcl_Obj *objs[6];
    Tcl_Obj *recObj;

    switch (recP->EventType) {
    case KEY_EVENT:
        objs[0] = ObjFromBoolean(recP->Event.KeyEvent.bKeyDown);
        objs[1] = ObjFromInt(recP->Event.KeyEvent.wRepeatCount);
        objs[2] = ObjFromInt(recP->Event.KeyEvent.wVirtualKeyCode);
        objs[3] = ObjFromInt(recP->Event.KeyEvent.wVirtualScanCode);
        if (recP->Event.KeyEvent.uChar.UnicodeChar)
            objs[4] = ObjFromWinCharsN(&recP->Event.KeyEvent.uChar.UnicodeChar, 1);
        else
            objs[4] = ObjFromEmptyString();
        objs[5] = ObjFromDWORD(recP->Event.KeyEvent.dwControlKeyState);
        recObj = ObjNewList(6, objs);
        break;

    case MOUSE_EVENT:
        objs[0] = ObjFromCOORD(&recP->Event.MouseEvent.dwMousePosition);
        /* NOTE: dwButtonState is actually signed because that is how
           direction of wheel movement is indicated */
        objs[1] = ObjFromInt(recP->Event.MouseEvent.dwButtonState);
        objs[2] = ObjFromDWORD(recP->Event.MouseEvent.dwControlKeyState);
        objs[3] = ObjFromDWORD(recP->Event.MouseEvent.dwEventFlags);
        recObj = ObjNewList(4, objs);
        break;

    case WINDOW_BUFFER_SIZE_EVENT:
        recObj = ObjFromCOORD(&recP->Event.WindowBufferSizeEvent.dwSize);
        break;

    case MENU_EVENT:
        recObj = ObjFromDWORD(recP->Event.MenuEvent.dwCommandId);
        break;

    case FOCUS_EVENT:
        recObj = ObjFromBoolean(recP->Event.FocusEvent.bSetFocus);
        break;
        
    default:
        recObj = ObjFromByteArray((BYTE *)recP, sizeof(*recP));
        break;
    }

    objs[0] = ObjFromInt(recP->EventType);
    objs[1] = recObj;

    return ObjNewList(2, objs);
}

static int Twapi_ReadConsole(Tcl_Interp *interp, HANDLE conh, unsigned int numchars)
{
    WCHAR buf[300];
    WCHAR *bufP;
    DWORD  len;
    int status;

    if (numchars > ARRAYSIZE(buf))
        bufP = SWSPushFrame(numchars * sizeof(WCHAR), NULL);
    else
        bufP = buf;

    if (ReadConsoleW(conh, bufP, numchars, &len, NULL)) {
        ObjSetResult(interp, ObjFromWinCharsN(bufP, len));
        status = TCL_OK;
    } else {
        TwapiReturnSystemError(interp);
        status = TCL_ERROR;
    }

    if (bufP != buf)
        SWSPopFrame();

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

    objs[0] = ObjFromString(TWAPI_TCL_NAMESPACE "::_console_ctrl_handler");
    objs[1] = ObjFromString(event_str);
    return TwapiEvalAndUpdateCallback(cbP, 2, objs, TRT_BOOL);
}

/* Directly called by Windows in a separate thread */
static BOOL WINAPI TwapiConsoleCtrlHandler(DWORD ctrl)
{
    TwapiCallback *cbP;
    BOOL handled = FALSE;
    DWORD winerr;

    /* TBD - there is a race here? */
    if (console_control_ticP == NULL)
        return FALSE;

    cbP = TwapiCallbackNew(
        console_control_ticP, TwapiConsoleCtrlCallbackFn, sizeof(*cbP));

    cbP->clientdata = ctrl;
    winerr = TwapiEnqueueCallback(console_control_ticP,
                                  cbP,
                                  TWAPI_ENQUEUE_DIRECT,
                                  100, /* Timeout (ms) */
                                  &cbP);
    if (winerr == ERROR_SUCCESS) {
        if (cbP && cbP->response.type == TRT_BOOL)
            handled = cbP->response.value.bval;
    } else if (winerr == WAIT_TIMEOUT)
        handled = 1;

    if (cbP)
        TwapiCallbackUnref((TwapiCallback *)cbP, 1);

    return handled;
}

static int Twapi_StartConsoleEventNotifier(TwapiInterpContext *ticP)
{
    void *pv;

    RETURN_ERROR_IF_UNTHREADED(ticP->interp);
    pv = InterlockedCompareExchangePointer(&console_control_ticP,
                                           ticP, NULL);
    if (pv) {
        ObjSetStaticResult(ticP->interp, "Console control handler is already set.");
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
        ObjSetStaticResult(ticP->interp, "Console control handler not set by this interpreter.");
        return TCL_ERROR;
    }
    SetConsoleCtrlHandler(TwapiConsoleCtrlHandler, FALSE);
    ticP->module.data.ival = 0; /* Indicates console is not hooked in this interp */

    TwapiInterpContextUnref(ticP, 1);
    return TCL_OK;
}

static int Twapi_ConsoleEventNotifierObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    int func;

    if (objc != 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);
    
    if (func)
        return Twapi_StartConsoleEventNotifier(ticP);
    else
        return Twapi_StopConsoleEventNotifier(ticP);
}

static void TwapiConsoleCleanup(TwapiInterpContext *ticP)
{
    if (ticP->module.data.ival)
        Twapi_StopConsoleEventNotifier(ticP);
}

static int Twapi_ConsoleCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    union {
        CONSOLE_SCREEN_BUFFER_INFO csbi;
        WCHAR buf[MAX_PATH+1];
        SMALL_RECT srect[2];
        INPUT_RECORD *inrecP;
    } u;
    COORD coord;
    CHAR_INFO chinfo;
    DWORD dw, dw2, dw3;
    int ival;
    SECURITY_ATTRIBUTES *secattrP;
    HANDLE h;
    WORD w;
    int func = PtrToInt(clientdata);
    TwapiResult result;
    Tcl_Obj *sObj;
    LPWSTR s;
    SWSMark mark = NULL;
    Tcl_Size len;

    --objc;
    ++objv;

    result.type = TRT_BADFUNCTIONCODE;

    if (func < 100) {
        /* Functions taking no arguments */
        if (objc != 0)
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
                /* On some platforms, the returned value includes the 
                 * terminating null, on others it does not.
                 */
                if (result.value.unicode.len > 0 &&
                    u.buf[result.value.unicode.len-1] == 0)
                    result.value.unicode.len -= 1;

                result.value.unicode.str = u.buf;
            } else
                result.type = TRT_GETLASTERROR;
            break;
        case 7:
            result.value.hwin = GetConsoleWindow();
            result.type = result.value.hwin ? TRT_HWND : TRT_GETLASTERROR;
            break;
        }
    } else if (func < 200) {
        /* First arg integer, maybe two more */
        if (objc == 0)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        CHECK_DWORD_OBJ(interp, dw, objv[0]);
        switch (func) {
        case 100:
            result.value.ival = AttachConsole(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 101:
            result.value.ival = SetConsoleCP(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 102:
            result.value.ival = SetConsoleOutputCP(dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            break;
        case 103:
            if (objc != 2)
                return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
            CHECK_DWORD_OBJ(interp, dw2, objv[1]);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = GenerateConsoleCtrlEvent(dw, dw2);
            break;
        }
    } else if (func < 300) {
        /* A single handle arguments */
        if (objc != 1)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        if (ObjToHANDLE(interp, objv[0], &h) != TCL_OK)
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
                result.type = TRT_OBJ;
                result.value.obj =
                    ObjFromCONSOLE_SCREEN_BUFFER_INFO(&u.csbi);
            }
            break;
        case 204:
            coord = GetLargestConsoleWindowSize(h);
            result.type = TRT_OBJ;
            result.value.obj = ObjFromCOORD(&coord);
            break;
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
        if (objc == 0)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

        switch (func) {
        case 1001: // SetConsoleWindowInfo
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h), GETBOOL(ival), GETVAR(u.srect[0], ObjToSMALL_RECT),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleWindowInfo(h, ival, &u.srect[0]);
            break;
        case 1002: // FillConsoleOutputAttribute
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h), GETWORD(w), GETDWORD(dw),
                             GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;

            if (FillConsoleOutputAttribute(h, w, dw, coord, &result.value.uval))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1003: // ScrollConsoleScreenBuffer
            if (TwapiGetArgs(interp, objc, objv,
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
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h), GETOBJ(sObj),
                             GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            s = ObjToWinCharsN(sObj, &len);
            CHECK_DWORD(interp, len);
            if (WriteConsoleOutputCharacterW(h, s,
                                             (DWORD)len, coord, &result.value.uval))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;    
            break;
        case 1005: // SetConsoleCursorPosition
        case 1006: // SetConsoleScreenBufferSize
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h), GETVAR(coord, ObjToCOORD),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival =
                (func == 1005 ? SetConsoleCursorPosition : SetConsoleScreenBufferSize)
                (h, coord);
            break;
        case 1007: // CreateConsoleScreenBuffer
            mark = SWSPushMark();
            if (TwapiGetArgs(interp, objc, objv,
                             GETDWORD(dw),
                             GETDWORD(dw2),
                             GETVAR(secattrP, ObjToPSECURITY_ATTRIBUTESSWS),
                             GETDWORD(dw3),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_HANDLE;
            result.value.hval = CreateConsoleScreenBuffer(dw, dw2, secattrP, dw3, NULL);
            break;
        case 1008:
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival = SetConsoleTitleW(ObjToWinChars(objv[0]));
            break;
        case 1009:
        case 1010:
        case 1011:
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h), GETDWORD(dw), ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (func == 1011)
                return Twapi_ReadConsole(interp, h, dw);
            result.type = TRT_EXCEPTION_ON_FALSE;
            result.value.ival =
                func == 1009 ? SetConsoleMode(h, dw) : SetConsoleTextAttribute(h, (WORD) dw);
            break;
        case 1012:
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h), GETOBJ(sObj),
                             ARGUSEDEFAULT, GETDWORD(dw), ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (WriteConsoleW(h, ObjToWinChars(sObj),
                              dw, &result.value.uval, NULL))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1013:
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h), GETOBJ(sObj),
                             GETDWORD(dw), ARGSKIP, ARGEND) != TCL_OK)
                return TCL_ERROR;
            if (ObjToCOORD(interp, objv[3], &coord) != TCL_OK)
                return TCL_ERROR;
            if (FillConsoleOutputCharacterW(h, *ObjToWinChars(sObj), dw, coord, &result.value.uval))
                result.type = TRT_DWORD;
            else
                result.type = TRT_GETLASTERROR;
            break;
        case 1014: 
        case 1015:
            if (TwapiGetArgs(interp, objc, objv,
                             GETHANDLE(h),
                             GETDWORD(dw), ARGEND) != TCL_OK)
                return TCL_ERROR;

            mark = SWSPushMark();
            u.inrecP = SWSAlloc(dw * sizeof(INPUT_RECORD), NULL);
            if ((func == 1014 ? ReadConsoleInputW : PeekConsoleInputW)(h, u.inrecP, dw, &dw2) == 0)
                result.type = TRT_GETLASTERROR;
            else if (dw2 == 0)
                result.type = TRT_EMPTY;
            else {
                result.type = TRT_OBJ;
                result.value.obj = ObjNewList(dw2, NULL);
                for (dw = 0; dw < dw2; ++dw)
                    ObjAppendElement(interp, result.value.obj,
                                     ObjFromINPUT_RECORD(&u.inrecP[dw]));
            }
            break;
        }
    }

    if (mark)
        SWSPopMark(mark);
    
    return TwapiSetResult(interp, &result);
}


static int TwapiConsoleInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s ConsoleDispatch[] = {
        DEFINE_FNCODE_CMD(AllocConsole, 1),
        DEFINE_FNCODE_CMD(FreeConsole, 2),
        DEFINE_FNCODE_CMD(GetConsoleCP, 3),
        DEFINE_FNCODE_CMD(GetConsoleOutputCP, 4),
        DEFINE_FNCODE_CMD(GetNumberOfConsoleMouseButtons, 5),
        DEFINE_FNCODE_CMD(GetConsoleTitle, 6),
        DEFINE_FNCODE_CMD(GetConsoleWindow, 7),

        DEFINE_FNCODE_CMD(AttachConsole, 100), /* TBD - Tcl */
        DEFINE_FNCODE_CMD(SetConsoleCP, 101),
        DEFINE_FNCODE_CMD(SetConsoleOutputCP, 102),
        DEFINE_FNCODE_CMD(GenerateConsoleCtrlEvent, 103),

        DEFINE_FNCODE_CMD(FlushConsoleInputBuffer, 201),
        DEFINE_FNCODE_CMD(GetConsoleMode, 202),
        DEFINE_FNCODE_CMD(GetConsoleScreenBufferInfo, 203),
        DEFINE_FNCODE_CMD(GetLargestConsoleWindowSize, 204),
        DEFINE_FNCODE_CMD(GetNumberOfConsoleInputEvents, 205),
        DEFINE_FNCODE_CMD(SetConsoleActiveScreenBuffer, 206),

        DEFINE_FNCODE_CMD(SetConsoleWindowInfo, 1001),
        DEFINE_FNCODE_CMD(FillConsoleOutputAttribute, 1002),
        DEFINE_FNCODE_CMD(ScrollConsoleScreenBuffer, 1003),
        DEFINE_FNCODE_CMD(WriteConsoleOutputCharacter, 1004),
        DEFINE_FNCODE_CMD(SetConsoleCursorPosition, 1005),
        DEFINE_FNCODE_CMD(SetConsoleScreenBufferSize, 1006),
        DEFINE_FNCODE_CMD(CreateConsoleScreenBuffer, 1007),
        DEFINE_FNCODE_CMD(SetConsoleTitle, 1008),
        DEFINE_FNCODE_CMD(SetConsoleMode, 1009),
        DEFINE_FNCODE_CMD(SetConsoleTextAttribute, 1010),
        DEFINE_FNCODE_CMD(ReadConsole, 1011),
        DEFINE_FNCODE_CMD(WriteConsole, 1012),
        DEFINE_FNCODE_CMD(FillConsoleOutputCharacter, 1013),
        DEFINE_FNCODE_CMD(ReadConsoleInput, 1014),
        DEFINE_FNCODE_CMD(PeekConsoleInput, 1015),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(ConsoleDispatch), ConsoleDispatch,
                          Twapi_ConsoleCallObjCmd);

    /* The following command requires a ticP so cannot be included above */
    Tcl_CreateObjCommand(interp, "twapi::Twapi_ConsoleEventNotifier", Twapi_ConsoleEventNotifierObjCmd, ticP, NULL);

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
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiConsoleInitCalls,
        TwapiConsoleCleanup
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    /* Note ticP->module.data.ival is initialized to 0 which we use
     * as an indicator whether this interp has hooked ctrl-c or not.
     * Moreover, we are specifying a finalizer so we have to request
     * a private context
     */
    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, NEW_TIC) ? TCL_OK : TCL_ERROR;
}

