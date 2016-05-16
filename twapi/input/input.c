/* 
 * Copyright (c) 2004-2012 Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Define interface to Windows API related to user input APIs */

#include "twapi.h"

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

#ifndef MODULENAME
#define MODULENAME "twapi_input"
#endif

static void init_keyboard_input(INPUT *pin, WORD vkey, DWORD flags);

int Twapi_UnregisterHotKey(TwapiInterpContext *ticP, int id);
int Twapi_RegisterHotKey(TwapiInterpContext *ticP, int id, UINT modifiers, UINT vk);
int Twapi_BlockInput(Tcl_Interp *interp, BOOL block);
int Twapi_SendUnicode(TwapiInterpContext *ticP, Tcl_Obj *input_obj);
int Twapi_SendInput(TwapiInterpContext *ticP, Tcl_Obj *input_obj);

/* TBD - move to script level code */
int Twapi_RegisterHotKey(TwapiInterpContext *ticP, int id, UINT modifiers, UINT vk)
{
    HWND hwnd;

    // Get the common notification window.
    hwnd = Twapi_GetNotificationWindow(ticP);
    if (hwnd == NULL)
        return TCL_ERROR;

    if (RegisterHotKey(hwnd, id, modifiers, vk))
        return TCL_OK;
    else
        return TwapiReturnSystemError(ticP->interp);
}


int Twapi_UnregisterHotKey(TwapiInterpContext *ticP, int id)
{
    // Note since we are using the common window for notifications,
    // we do not destroy it. Just unregister the hot key.
    HWND hwnd;
    hwnd = Twapi_GetNotificationWindow(ticP);
    
    if (UnregisterHotKey(hwnd, id))
        return TCL_OK;
    else
        return TwapiReturnSystemError(ticP->interp);
}

int Twapi_BlockInput(Tcl_Interp *interp, BOOL block)
{
    BOOL result = BlockInput(block);
    if (result || (GetLastError() == 0)) {
        ObjSetResult(interp, ObjFromInt(result));
        return TCL_OK;
    } else {
        return TwapiReturnSystemError(interp);
    }
}

int Twapi_SendInput(TwapiInterpContext *ticP, Tcl_Obj *input_obj) {
    int num_inputs;
    struct tagINPUT   *input;
    int i, j;
    int result = TCL_ERROR;
    Tcl_Interp *interp = ticP->interp;
    
    if (ObjListLength(interp, input_obj, &num_inputs) != TCL_OK) {
        return TCL_ERROR;
    }

    input = MemLifoPushFrame(ticP->memlifoP, num_inputs * sizeof(*input), NULL);
    /* Loop through each element, parsing it and storing its descriptor */
    for (i = 0; i < num_inputs; ++i) {
        Tcl_Obj *event_obj;
        Tcl_Obj *field_obj[5];
        LONG     value[5];
        static const char *options[] = {"key", "mouse", NULL};
        int   option;


        if (ObjListIndex(interp, input_obj, i, &event_obj) != TCL_OK)
            goto done;

        if (event_obj == NULL)
            break;

        /* This element is itself a list, parse it to get input type etc. */
        if (ObjListIndex(interp, event_obj, 0, &field_obj[0]) != TCL_OK)
            goto  done;
        
        if (field_obj[0] == NULL)
            break;

        /* Figure out the input type and parse remaining fields */
        if (Tcl_GetIndexFromObj(interp, field_obj[0], options,
                                "input event type", TCL_EXACT, &option) != TCL_OK)
            goto done;

        switch (option) {
        case 0:
            /* A single key stroke. Fields are:
             *  virtualkey(1-254), scancode (0-65535), flags
             * Extra arguments ignored
             */
            for (j = 1; j < 4; ++j) {
                if (ObjListIndex(interp, event_obj, j, &field_obj[j]) != TCL_OK)
                    goto done;
                if (field_obj[j] == NULL) {
                    ObjSetStaticResult(interp, "Missing field in key event");
                    goto done;
                }
                if (ObjToLong(interp, field_obj[j], &value[j]) != TCL_OK)
                    goto done;
            }

            /* OK, our three fields have been parsed.
             * Validate and add to input
             */
            if (value[1] < 0 || value[1] > 254) {
                ObjSetStaticResult(interp, "Invalid virtual key code");
                goto done;
            }
            if (value[2] < 0 || value[2] > 65535) {
                ObjSetStaticResult(interp, "Invalid scan code value.");
                goto done;
            }
            init_keyboard_input(&input[i], (WORD) value[1], value[3]);
            input[i].ki.wScan   = (WORD) value[2];
            break;

        case 1:
            /* Mouse event
             *  xpos ypos mousedata flags
             * Extra arguments ignored
             */
            for (j = 1; j < 5; ++j) {
                if (ObjListIndex(interp, event_obj, j, &field_obj[j]) != TCL_OK)
                    goto done;
                if (field_obj[j] == NULL) {
                    ObjSetStaticResult(interp, "Missing field in mouse event");
                    goto done;
                }
                if (ObjToLong(interp, field_obj[j], &value[j]) != TCL_OK)
                    goto done;
            }
            
            input[i].type           = INPUT_MOUSE;
            input[i].mi.dx          = value[1];
            input[i].mi.dy          = value[2];
            input[i].mi.mouseData   = value[3];
            input[i].mi.dwFlags     = value[4];
            input[i].mi.time        = 0;
            input[i].mi.dwExtraInfo = 0;
            break;

         default:
            /* Shouldn't happen else Tcl_GetIndexFromObj would return error */
             ObjSetStaticResult(interp, "Unknown field event type");
            goto done;
        }

    }

    
    /* i is actual number of elements found */
    if (i != num_inputs) {
        ObjSetStaticResult(interp, "Invalid or empty element in input event list");
        goto done;
    }

    /* OK, we have everything in the input[] array. Send it along */
    if (i) {
        num_inputs = SendInput(i, input, sizeof(input[0]));
        if (num_inputs == 0) {
            j = GetLastError();
            ObjSetStaticResult(interp, "Error sending input events: ");
            Twapi_AppendSystemError(interp, j);
            goto done;
        }
    }    

    ObjSetResult(interp, ObjFromInt(num_inputs));
    result = TCL_OK;

 done:

    MemLifoPopFrame(ticP->memlifoP);

    return result;
}


int Twapi_SendUnicode(TwapiInterpContext *ticP, Tcl_Obj *input_obj) {
    int num_chars;
    struct tagINPUT   *input = NULL;
    int i, j;
    int result = TCL_ERROR;
    int max_input_records;
    int sent_inputs;
    
    num_chars = Tcl_GetCharLength(input_obj);

    /* Now loop through every character adding it to the input event array */
    /* Win2K and up, accepts unicode characters */

    /* NUmber of events is twice number of chars (keydown + keyup) */
    max_input_records = 2 * num_chars;
    input = MemLifoAlloc(ticP->memlifoP, max_input_records * sizeof(*input), NULL);
    for (i = 0, j = 0; i < num_chars; ++i) {
        WCHAR wch;
            
#if TCL_UTF_MAX <= 4
        wch = Tcl_GetUniChar(input_obj, i);
#else
        wch = (WCHAR) Tcl_GetUniChar(input_obj, i);
#endif

#ifndef KEYEVENTF_UNICODE
#define KEYEVENTF_UNICODE     0x0004
#endif
        init_keyboard_input(&input[j], 0, KEYEVENTF_UNICODE);
        input[j].ki.wScan = wch;
        ++j;
        init_keyboard_input(&input[j], 0, KEYEVENTF_UNICODE|KEYEVENTF_KEYUP);
        input[j].ki.wScan  = wch;
        ++j;
    }
    
    /* j is actual number of input events created */
    assert (j <= max_input_records);

    /* OK, we have everything in the input[] array. Send it along */
    if (j) {
        sent_inputs = SendInput(j, input, sizeof(input[0]));
        if (sent_inputs == 0) {
            i = GetLastError();
            ObjSetStaticResult(ticP->interp, "Error sending input events: ");
            Twapi_AppendSystemError(ticP->interp, i);
            goto done;
        }
        /* TBD - what if we send fewer than expected, should we retry ? */
    } else {
        sent_inputs = 0;
    }

    ObjSetResult(ticP->interp, ObjFromInt(sent_inputs));
    result = TCL_OK;

 done:
    MemLifoPopFrame(ticP->memlifoP);

    return result;
}

static void init_keyboard_input(INPUT *pin, WORD vkey, DWORD flags)
{
    pin->type       = INPUT_KEYBOARD;
    pin->ki.wVk     = vkey;
    pin->ki.wScan   = 0;
    pin->ki.dwFlags = flags;
    pin->ki.time    = 0;
    pin->ki.dwExtraInfo = 0;
}

static int Twapi_InputCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiInterpContext *ticP = (TwapiInterpContext*) clientdata;
    int func;
    DWORD dw, dw2, dw3;
    LASTINPUTINFO lastin;
    TwapiResult result;
    WCHAR kl[KL_NAMELENGTH+1];

    if (objc < 2)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    CHECK_INTEGER_OBJ(interp, func, objv[1]);

    result.type = TRT_BADFUNCTIONCODE;
    switch (func) {
    case 1:
        result.type = TRT_DWORD;
        result.value.uval = GetDoubleClickTime();
        break;
    case 2:
        lastin.cbSize = sizeof(lastin);
        if (GetLastInputInfo(&lastin)) {
            result.type = TRT_DWORD;
            result.value.uval = lastin.dwTime;
        } else {
            result.type = TRT_GETLASTERROR;
        }
        break;
    case 3:
    case 4:
    case 5:
    case 6:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), ARGEND) != TCL_OK)
            return TCL_ERROR;
        switch (func) {
        case 3:
            result.type = TRT_DWORD;
            result.value.uval = GetAsyncKeyState(dw);
            break;
        case 4:
            result.type = TRT_DWORD;
            result.value.uval = GetKeyState(dw);
            break;
        case 5:
            return Twapi_BlockInput(interp, dw);
        case 6:
            return Twapi_UnregisterHotKey(ticP, dw);
        }
        break;
    case 7:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), GETINT(dw2),
                         ARGEND) != TCL_OK)
            return TCL_ERROR;
        result.type = TRT_DWORD;
        result.value.uval = MapVirtualKey(dw, dw2);
        break;
    case 8:
        if (TwapiGetArgs(interp, objc-2, objv+2, GETINT(dw), GETINT(dw2),
                         GETINT(dw3), ARGEND) != TCL_OK)
            return TCL_ERROR;
        return Twapi_RegisterHotKey(ticP, dw, dw2, dw3);
    case 9:
    case 10:
        if (objc != 3)
            return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
        return (func == 9 ? Twapi_SendInput : Twapi_SendUnicode) (ticP, objv[2]);
    case 11:
        if (! GetKeyboardLayoutNameW(kl))
            result.type = TRT_GETLASTERROR;
        else {
            result.type = TRT_UNICODE;
            result.value.unicode.str = kl;
            result.value.unicode.len = -1;
        }
        break;
    }
    return TwapiSetResult(interp, &result);
}


static int TwapiInputInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct alias_dispatch_s InputAliasDispatch[] = {
        DEFINE_ALIAS_CMD(GetDoubleClickTime, 1),
        DEFINE_ALIAS_CMD(GetLastInputInfo, 2),
        DEFINE_ALIAS_CMD(GetAsyncKeyState, 3), // TBD - Tcl
        DEFINE_ALIAS_CMD(GetKeyState, 4),  // TBD - Tcl
        DEFINE_ALIAS_CMD(BlockInput, 5),
        DEFINE_ALIAS_CMD(UnregisterHotKey, 6),
        DEFINE_ALIAS_CMD(MapVirtualKey, 7),
        DEFINE_ALIAS_CMD(RegisterHotKey, 8),
        DEFINE_ALIAS_CMD(SendInput, 9),
        DEFINE_ALIAS_CMD(Twapi_SendUnicode, 10),
        DEFINE_ALIAS_CMD(get_keyboard_layout_name, 11),
    };

    /* Create the underlying call dispatch commands */
    Tcl_CreateObjCommand(interp, "twapi::InputCall", Twapi_InputCallObjCmd, ticP, NULL);

    /* Now add in the aliases for the Win32 calls pointing to the dispatcher */
    TwapiDefineAliasCmds(interp, ARRAYSIZE(InputAliasDispatch), InputAliasDispatch, "twapi::InputCall");

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
int Twapi_input_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiInputInitCalls,
        NULL
    };

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

