/* 
 * Copyright (c) 2006-2013, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/*
 * Define interface to Windows API related MS Task
 */

#include "twapi.h"
#include <mstask.h>
#include <initguid.h> /* GUIDs in all included files below this will be instantiated */
DEFINE_GUID(IID_IEnumWorkItems, 0x148BD528L, 0xA2AB, 0x11CE, 0xB1, 0x1F, 0x00, 0xAA, 0x00, 0x53, 0x05, 0x03);
DEFINE_GUID(IID_IScheduledWorkItem, 0xa6b952f0L, 0xa4b1, 0x11d0, 0x99, 0x7d, 0x00, 0xaa, 0x00, 0x68, 0x87, 0xec);

#ifndef TWAPI_SINGLE_MODULE
static HMODULE gModuleHandle;     /* DLL handle to ourselves */
#endif

int ObjToTASK_TRIGGER(Tcl_Interp *interp, Tcl_Obj *obj, TASK_TRIGGER *triggerP);
Tcl_Obj *ObjFromTASK_TRIGGER(TASK_TRIGGER *triggerP);

int Twapi_IEnumWorkItems_Next(Tcl_Interp *interp,
        IEnumWorkItems *ewiP, unsigned long count);
int Twapi_IScheduledWorkItem_GetRunTimes(Tcl_Interp *interp,
        IScheduledWorkItem *swiP, SYSTEMTIME *, SYSTEMTIME *, WORD );
int Twapi_IScheduledWorkItem_GetWorkItemData(Tcl_Interp *interp,
                                             IScheduledWorkItem *swiP);


Tcl_Obj *ObjFromTASK_TRIGGER(TASK_TRIGGER *triggerP)
{
    Tcl_Obj *resultObj;
    Tcl_Obj *typeObj[4];
    int      ntype;

    resultObj = ObjNewList(0, 0);

    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, Reserved1);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wBeginYear);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wBeginMonth);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wBeginDay);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wEndYear);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wEndMonth);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wEndDay);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wStartHour);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wStartMinute);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, MinutesDuration);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, MinutesInterval);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, rgFlags);

    ntype = 1;
    typeObj[0] = ObjFromInt(triggerP->TriggerType);

    switch (triggerP->TriggerType) {
    case 1:
        typeObj[1] = ObjFromInt(triggerP->Type.Daily.DaysInterval);
        ntype = 2;
        break;
    case 2:
        typeObj[1] = ObjFromInt(triggerP->Type.Weekly.WeeksInterval);
        typeObj[2] = ObjFromInt(triggerP->Type.Weekly.rgfDaysOfTheWeek);
        ntype = 3;
        break;
    case 3:
        typeObj[1] = ObjFromDWORD(triggerP->Type.MonthlyDate.rgfDays);
        typeObj[2] = ObjFromInt(triggerP->Type.MonthlyDate.rgfMonths);
        ntype = 3;
        break;
    case 4:
        typeObj[1] = ObjFromInt(triggerP->Type.MonthlyDOW.wWhichWeek);
        typeObj[2] = ObjFromInt(triggerP->Type.MonthlyDOW.rgfDaysOfTheWeek);
        typeObj[3] = ObjFromInt(triggerP->Type.MonthlyDOW.rgfMonths);
        ntype = 4;
        break;
    }
    ObjAppendElement(NULL, resultObj, STRING_LITERAL_OBJ("type"));
    ObjAppendElement(NULL, resultObj, ObjNewList(ntype, typeObj));

    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, Reserved2);
    Twapi_APPEND_LONG_FIELD_TO_LIST(NULL, resultObj, triggerP, wRandomMinutesInterval);
    
    return resultObj;
}


int ObjToTASK_TRIGGER(Tcl_Interp *interp, Tcl_Obj *obj, TASK_TRIGGER *triggerP)
{
    int       i;
    Tcl_Obj **objv;
    int       objc;
    long      dval;

    if (ObjGetElements(interp, obj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc & 1) {
        ObjSetStaticResult(interp, "Invalid TASK_TRIGGER format - must have even number of elements");
        return TCL_ERROR;
    }

    TwapiZeroMemory(triggerP, sizeof(*triggerP));
    triggerP->cbTriggerSize = sizeof(*triggerP);
    
    for (i=0; i < (objc-1); i+=2) {
        char *name = ObjToString(objv[i]);
        if (STREQ(name, "wBeginYear")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wBeginYear) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "Reserved1")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->Reserved1) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wBeginMonth")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wBeginMonth) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wBeginDay")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wBeginDay) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wEndYear")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wEndYear) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wEndMonth")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wEndMonth) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wEndDay")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wEndDay) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wStartHour")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wStartHour) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wStartMinute")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wStartMinute) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "MinutesDuration")) {
            if (ObjToLong(interp, objv[i+1], &triggerP->MinutesDuration) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "MinutesInterval")) {
            if (ObjToLong(interp, objv[i+1], &triggerP->MinutesInterval) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "rgFlags")) {
            if (ObjToLong(interp, objv[i+1], &triggerP->rgFlags) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "Reserved2")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->Reserved2) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "wRandomMinutesInterval")) {
            if (ObjToWord(interp, objv[i+1], &triggerP->wRandomMinutesInterval) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "type")) {
            Tcl_Obj **typeObj;
            int      ntype;

            /* Object should be a list */
            if (ObjGetElements(interp, objv[i+1], &ntype, &typeObj) != TCL_OK) {
                return TCL_ERROR;
            }
            if (ntype == 0) {
            trigger_type_count_error:
                ObjSetStaticResult(interp, "Invalid task trigger type format");
                return TCL_ERROR;
            }

            /* First element is the type */
            if (ObjToLong(interp, typeObj[0], &dval) != TCL_OK) {
                return TCL_ERROR;
            }
            triggerP->TriggerType = dval;
            switch (triggerP->TriggerType) {
            case 1:
                if (ntype != 2)
                    goto trigger_type_count_error;
                if (ObjToWord(interp, typeObj[1], &triggerP->Type.Daily.DaysInterval) != TCL_OK)
                    goto trigger_type_count_error;
                break;

            case 2:
                if (ntype != 3)
                    goto trigger_type_count_error;
                if (ObjToWord(interp, typeObj[1], &triggerP->Type.Weekly.WeeksInterval) != TCL_OK)
                    goto trigger_type_count_error;
                if (ObjToWord(interp, typeObj[2], &triggerP->Type.Weekly.rgfDaysOfTheWeek) != TCL_OK)
                    goto trigger_type_count_error;
                break;

            case 3:
                if (ntype != 3)
                    goto trigger_type_count_error;
                if (ObjToLong(interp, typeObj[1], &triggerP->Type.MonthlyDate.rgfDays) != TCL_OK)
                    goto trigger_type_count_error;
                if (ObjToWord(interp, typeObj[2], &triggerP->Type.MonthlyDate.rgfMonths) != TCL_OK)
                    goto trigger_type_count_error;
                break;

            case 4:
                if (ntype != 4)
                    goto trigger_type_count_error;
                if (ObjToWord(interp, typeObj[1], &triggerP->Type.MonthlyDOW.wWhichWeek) != TCL_OK)
                    goto trigger_type_count_error;
                if (ObjToWord(interp, typeObj[2], &triggerP->Type.MonthlyDOW.rgfDaysOfTheWeek) != TCL_OK)
                    goto trigger_type_count_error;
                if (ObjToWord(interp, typeObj[3], &triggerP->Type.MonthlyDOW.rgfMonths) != TCL_OK)
                    goto trigger_type_count_error;
                break;
            }
        }
        else {
            Tcl_AppendResult(interp, "Unknown TASK_TRIGGER field '", name, "'", NULL);
            return TCL_ERROR;
        }
    }
    return TCL_OK;
}


int Twapi_IEnumWorkItems_Next(Tcl_Interp *interp,
                              IEnumWorkItems *ewiP,
                              unsigned long count)
{
    Tcl_Obj *objv[2];           // {More, List_of_elements}
    unsigned long ret_count;
    LPWSTR *jobsP;
    HRESULT  hr;
    unsigned long i;

    if (count == 0) {
        objv[0] = ObjFromBoolean(1);
        objv[1] = ObjNewList(0, NULL);
        ObjSetResult(interp, ObjNewList(2, objv));
        return TCL_OK;
    }

    hr = ewiP->lpVtbl->Next(ewiP, count, &jobsP, &ret_count);

    /*
     * hr - S_OK ret_count elements returned, more to come
     *      S_FALSE returned elements returned, no more
     *      else error
     */
    if (hr != S_OK && hr != S_FALSE) {
        TWAPI_STORE_COM_ERROR(interp, hr, ewiP, &IID_IEnumWorkItems);
        return TCL_ERROR;
    }

    objv[0] = ObjFromBoolean(hr == S_OK); // More to come?
    objv[1] = ObjNewList(0, NULL);
    if (jobsP) {
        for (i = 0; i < ret_count; ++i) {
            ObjAppendElement(interp, objv[1],
                                     ObjFromUnicode(jobsP[i]));
            /* Free the string */
            CoTaskMemFree(jobsP[i]);
        }
        /* Free the array itself */
        CoTaskMemFree(jobsP);
    }    

    ObjSetResult(interp, ObjNewList(2, objv));
    return TCL_OK;
}

int Twapi_IScheduledWorkItem_GetRunTimes (
    Tcl_Interp *interp,
    IScheduledWorkItem *swiP,
    SYSTEMTIME *beginP,
    SYSTEMTIME *endP,
    WORD        count
)
{
    Tcl_Obj *objv[2];           /* status + list of runtimes */
    HRESULT  hr;
    WORD     i;
    SYSTEMTIME *systimeP = NULL;

    hr = swiP->lpVtbl->GetRunTimes(swiP, beginP, endP, &count, &systimeP);

    objv[1] = NULL;

    switch (hr) {
    case S_OK: /* FALLTHRU */
    case S_FALSE:
        /* Both these cases are success. */
        objv[0] = STRING_LITERAL_OBJ("success");
        objv[1] = ObjNewList(0, NULL);
        for (i = 0; i < count; ++i) {
            ObjAppendElement(interp, objv[1],
                                     ObjFromSYSTEMTIME(&systimeP[i]));
        }
        if (systimeP)
            CoTaskMemFree(systimeP);
        break;

    case SCHED_S_EVENT_TRIGGER:
        objv[0] = STRING_LITERAL_OBJ("oneventonly");
        break;

    case SCHED_S_TASK_NO_VALID_TRIGGERS:
        objv[0] = STRING_LITERAL_OBJ("notriggers");
        break;

    case SCHED_S_TASK_DISABLED:
        objv[0] = STRING_LITERAL_OBJ("disabled");
        break;

    default:
        TWAPI_STORE_COM_ERROR(interp, hr, swiP, &IID_IScheduledWorkItem);
        return TCL_ERROR;
    }

    ObjSetResult(interp, ObjNewList(objv[1] ? 2 : 1, objv));
    return TCL_OK;
    
}


int Twapi_IScheduledWorkItem_GetWorkItemData (
    Tcl_Interp *interp,
    IScheduledWorkItem *swiP)
{
    HRESULT hr;
    BYTE *dataP;
    WORD count;

    hr = swiP->lpVtbl->GetWorkItemData(swiP, &count, &dataP);
    if (SUCCEEDED(hr)) {
        ObjSetResult(interp, ObjFromByteArray(dataP, count));
        return TCL_OK;
    }
    else {
        TWAPI_STORE_COM_ERROR(interp, hr, swiP, &IID_IScheduledWorkItem);
        return TCL_ERROR;
    }
}

/* Dispatcher for calling COM object methods */
int Twapi_MstaskCallObjCmd(ClientData clientdata, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    union {
        IUnknown *unknown;
        IDispatch *dispatch;
        IDispatchEx *dispatchex;
        ITaskScheduler *taskscheduler;
        IEnumWorkItems *enumworkitems;
        IScheduledWorkItem *scheduledworkitem;
        ITask *task;
        ITaskTrigger *tasktrigger;
    } ifc;
    HRESULT hr;
    TwapiResult result;
    DWORD dw1;
    HANDLE h;
    void *pv;
    GUID guid, guid2;
    Tcl_Obj *sObj;
    WORD w, w2;
    Tcl_Obj *objs[4];
    SYSTEMTIME systime, systime2;
    TASK_TRIGGER tasktrigger;
    int func = PtrToInt(clientdata);
    WCHAR *s, *passwordP;

    hr = S_OK;
    result.type = TRT_BADFUNCTIONCODE;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETVOIDP(pv), ARGTERM) != TCL_OK)
        return TCL_ERROR;

    if (pv == NULL) {
        ObjSetStaticResult(interp, "NULL interface pointer.");
        return TCL_ERROR;
    }

    /* We want stronger type checking so we have to convert the interface
       pointer on a per interface type basis even though it is repetitive. */

    --objc;
    ++objv;
    /* function codes are split into ranges assigned to interfaces */
    if (func < 5100) {
        /* ITaskScheduler */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.taskscheduler,
                        "ITaskScheduler") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5001: // Activate
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETOBJ(sObj), GETVAR(guid, ObjToGUID),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IUnknown";
            hr = ifc.taskscheduler->lpVtbl->Activate(
                ifc.taskscheduler,   ObjToUnicode(sObj),   &guid,
                (IUnknown **) &result.value.ifc.p);
            break;
        case 5002: // AddWorkItem
            if (objc != 3)
                goto badargs;
            if (ObjToOpaque(interp, objv[2], &pv, "IScheduledWorkItem") != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.taskscheduler->lpVtbl->AddWorkItem(
                ifc.taskscheduler,
                ObjToUnicode(objv[1]),
                (IScheduledWorkItem *) pv );
            break;
        case 5003: // Delete
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.taskscheduler->lpVtbl->Delete(ifc.taskscheduler,
                                                   ObjToUnicode(objv[1]));
            break;
        case 5004: // Enum
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumWorkItems";
            hr = ifc.taskscheduler->lpVtbl->Enum(
                ifc.taskscheduler,
                (IEnumWorkItems **) &result.value.ifc.p);
            break;
        case 5005: // IsOfType
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETOBJ(sObj), GETVAR(guid, ObjToGUID),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_BOOL;
            result.value.bval = ifc.taskscheduler->lpVtbl->IsOfType(
                ifc.taskscheduler, ObjToUnicode(sObj), &guid) == S_OK;
            break;
        case 5006: // NewWorkItem
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETOBJ(sObj),
                             GETGUID(guid), GETGUID(guid2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IUnknown";
            hr = ifc.taskscheduler->lpVtbl->NewWorkItem(
                ifc.taskscheduler, ObjToUnicode(sObj), &guid, &guid2,
                (IUnknown **) &result.value.ifc.p);
            break;
        case 5007: // SetTargetComputer
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.taskscheduler->lpVtbl->SetTargetComputer(
                ifc.taskscheduler, ObjToLPWSTR_NULL_IF_EMPTY(objv[1]));
            break;
        case 5008: // GetTargetComputer
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.taskscheduler->lpVtbl->GetTargetComputer(
                ifc.taskscheduler, &result.value.lpolestr);
            break;
        }        
    } else if (func < 5200) {
        /* IEnumWorkItems */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.enumworkitems,
                        "IEnumWorkItems") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5101: // Clone
            if (objc != 1)
                goto badargs;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "IEnumWorkItems";
            hr = ifc.enumworkitems->lpVtbl->Clone(
                ifc.enumworkitems, (IEnumWorkItems **)&result.value.ifc.p);
            break;
        case 5102: // Reset
            if (objc != 1)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.enumworkitems->lpVtbl->Reset(ifc.enumworkitems);
            break;
        case 5103: // Skip
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.enumworkitems->lpVtbl->Skip(ifc.enumworkitems, dw1);
            break;
        case 5104: // Next
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            return Twapi_IEnumWorkItems_Next(interp,ifc.enumworkitems,dw1);
        }
    } else if (func < 5300) {
        /* IScheduledWorkItem - ITask inherits from this and is also allowed */
        char *allowed_types[] = {"ITask", "IScheduledWorkItem"};

        if (ObjToOpaqueMulti(interp, objv[0], (void **)&ifc.scheduledworkitem,
                             ARRAYSIZE(allowed_types), allowed_types) != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5201: // CreateTrigger
            if (objc != 1)
                goto badargs;
            result.type = TRT_OBJV;
            hr = ifc.scheduledworkitem->lpVtbl->CreateTrigger(
                ifc.scheduledworkitem, &w, (ITaskTrigger **) &pv);
            if (hr != S_OK)
                break;
            objs[0] = ObjFromLong(w);
            objs[1] = ObjFromOpaque(pv, "ITaskTrigger");
            result.type = TRT_OBJV;
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            break;
        case 5202: // DeleteTrigger
            if (objc != 2)
                goto badargs;
            if (ObjToWord(interp, objv[1], &w) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->DeleteTrigger(
                ifc.scheduledworkitem, w);
            break;
        case 5203: // EditWorkItem
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETHANDLET(h, HWND), GETINT(dw1),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->EditWorkItem(
                ifc.scheduledworkitem, h, dw1);
            break;
        case 5204: // GetAccountInformation
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetAccountInformation(
                ifc.scheduledworkitem, &result.value.lpolestr);
            break;
        case 5205: // GetComment
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetComment(
                ifc.scheduledworkitem, &result.value.lpolestr);
            break;
        case 5206: // GetCreator
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetCreator(
                ifc.scheduledworkitem, &result.value.lpolestr);
            break;
        case 5207: // GetErrorRetryCount
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetErrorRetryCount(
                ifc.scheduledworkitem, &w);
            result.value.uval = w;
            break;
        case 5208: // GetErrorRetryInterval
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetErrorRetryInterval(
                ifc.scheduledworkitem, &w);
            result.value.uval = w;
            break;
        case 5209: // GetExitCode
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetExitCode(
                ifc.scheduledworkitem, &result.value.uval);
            break;
        case 5210: // GetFlags
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetFlags(
                ifc.scheduledworkitem, &result.value.uval);
            break;
        case 5211: // GetIdleWait
            if (objc != 1)
                goto badargs;
            hr = ifc.scheduledworkitem->lpVtbl->GetIdleWait(
                ifc.scheduledworkitem, &w, &w2);
            if (hr != S_OK)
                break;
            objs[0] = ObjFromLong(w);
            objs[1] = ObjFromLong(w2);
            result.type = TRT_OBJV;
            result.value.objv.nobj = 2;
            result.value.objv.objPP = objs;
            break;
        case 5212: // GetMostRecentRunTime
            if (objc != 1)
                goto badargs;
            result.type = TRT_SYSTEMTIME;
            hr = ifc.scheduledworkitem->lpVtbl->GetMostRecentRunTime(
                ifc.scheduledworkitem, &result.value.systime);
            break;
        case 5213: // GetNextRunTime
            if (objc != 1)
                goto badargs;
            result.type = TRT_SYSTEMTIME;
            hr = ifc.scheduledworkitem->lpVtbl->GetNextRunTime(
                ifc.scheduledworkitem, &result.value.systime);
            break;
        case 5214: // GetStatus
            if (objc != 1)
                goto badargs;
            result.type = TRT_LONG;
            hr = ifc.scheduledworkitem->lpVtbl->GetStatus(
                ifc.scheduledworkitem, &result.value.ival);
            break;
        case 5215: // GetTrigger
            if (objc != 2)
                goto badargs;
            if (ObjToWord(interp, objv[1], &w) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_INTERFACE;
            result.value.ifc.name = "ITaskTrigger";
            hr = ifc.scheduledworkitem->lpVtbl->GetTrigger(
                ifc.scheduledworkitem, w, (ITaskTrigger **)&result.value.ifc.p);
            break;
        case 5216: // GetTriggerCount
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.scheduledworkitem->lpVtbl->GetTriggerCount(
                ifc.scheduledworkitem, &w);
            result.value.uval = w;
            break;
        case 5217: // GetTriggerString
            if (objc != 2)
                goto badargs;
            if (ObjToWord(interp, objv[1], &w) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_LPOLESTR;
            hr = ifc.scheduledworkitem->lpVtbl->GetTriggerString(
                ifc.scheduledworkitem, w, &result.value.lpolestr);
            break;
        case 5218: // Run
            if (objc != 1)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->Run(ifc.scheduledworkitem);
            break;
        case 5219: // SetAccountInformation
            CHECK_NARGS_RANGE(interp, objc, 2, 3);
            s = ObjToUnicode(objv[1]);
            if (objc == 2 || s[0] == 0)
                passwordP = NULL;
            else
                passwordP = ObjDecryptPassword(objv[2], &dw1);
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetAccountInformation(
                ifc.scheduledworkitem, s, passwordP);
            TwapiFreeDecryptedPassword(passwordP, dw1);
            break;
        case 5220: // SetComment
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetComment(
                ifc.scheduledworkitem, ObjToUnicode(objv[1]));
            break;
        case 5221: // SetCreator
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetCreator(
                ifc.scheduledworkitem, ObjToUnicode(objv[1]));
            break;
        case 5222: // SetErrorRetryCount
            if (objc != 2)
                goto badargs;
            if (ObjToWord(interp, objv[1], &w) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetErrorRetryCount(
                ifc.scheduledworkitem, w);
            break;
        case 5223: // SetErrorRetryInterval
            if (objc != 2)
                goto badargs;
            if (ObjToWord(interp, objv[1], &w) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetErrorRetryInterval(
                ifc.scheduledworkitem, w);
            break;
        case 5224: // SetFlags
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetFlags(
                ifc.scheduledworkitem, dw1);
            break;
        case 5225: // SetIdleWait
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETWORD(w), GETWORD(w2),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetIdleWait(
                ifc.scheduledworkitem, w, w2);
            break;
        case 5226: // Terminate
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->Terminate(ifc.scheduledworkitem);
            break;
        case 5227: // SetWorkItemData
            if (objc != 2)
                goto badargs;
            pv = ObjToByteArray(objv[1], &dw1);
            if (dw1 > MAXWORD) {
                ObjSetStaticResult(interp, "Binary data exceeds MAXWORD");
                return TCL_ERROR;
            }
            result.type = TRT_EMPTY;
            hr = ifc.scheduledworkitem->lpVtbl->SetWorkItemData(
                ifc.scheduledworkitem, (WORD) dw1, pv);
            break;
        case 5228: // GetWorkItemData
            if (objc != 1)
                goto badargs;
            return Twapi_IScheduledWorkItem_GetWorkItemData(
                interp, ifc.scheduledworkitem);
        case 5229: // GetRunTimes
            if (TwapiGetArgs(interp, objc-1, objv+1,
                             GETVAR(systime, ObjToSYSTEMTIME),
                             GETVAR(systime2, ObjToSYSTEMTIME),
                             GETWORD(w),
                             ARGEND) != TCL_OK)
                return TCL_ERROR;
            return Twapi_IScheduledWorkItem_GetRunTimes(
                interp, ifc.scheduledworkitem, &systime, &systime2, w);
        }
    } else if (func < 5400) {
        /* ITask */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.task, "ITask") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5301: // GetApplicationName
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.task->lpVtbl->GetApplicationName(
                ifc.task, &result.value.lpolestr);
            break;
        case 5302: // GetMaxRunTime
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.task->lpVtbl->GetMaxRunTime(ifc.task, &result.value.uval);
            break;
        case 5303: // GetParameters
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.task->lpVtbl->GetParameters(
                ifc.task, &result.value.lpolestr);
            break;
        case 5304: // GetPriority
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.task->lpVtbl->GetPriority(ifc.task, &result.value.uval);
            break;
        case 5305: // GetTaskFlags
            if (objc != 1)
                goto badargs;
            result.type = TRT_DWORD;
            hr = ifc.task->lpVtbl->GetTaskFlags(ifc.task, &result.value.uval);
            break;
        case 5306: // GetWorkingDirectory
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.task->lpVtbl->GetWorkingDirectory(
                ifc.task, &result.value.lpolestr);
            break;
        case 5307: // SetApplicationName
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.task->lpVtbl->SetApplicationName(
                ifc.task, ObjToUnicode(objv[1]));
            break;
        case 5308: // SetParameters
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.task->lpVtbl->SetParameters(
                ifc.task, ObjToUnicode(objv[1]));
            break;
        case 5309: // SetWorkingDirectory
            if (objc != 2)
                goto badargs;
            result.type = TRT_EMPTY;
            hr = ifc.task->lpVtbl->SetWorkingDirectory(
                ifc.task, ObjToUnicode(objv[1]));
            break;
        case 5310: // SetMaxRunTime
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.task->lpVtbl->SetMaxRunTime(ifc.task, dw1);
            break;
        case 5311: // SetPriority
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.task->lpVtbl->SetPriority(ifc.task, dw1);
            break;
        case 5312: // SetTaskFlags
            if (objc != 2)
                goto badargs;
            CHECK_INTEGER_OBJ(interp, dw1, objv[1]);
            result.type = TRT_EMPTY;
            hr = ifc.task->lpVtbl->SetTaskFlags(ifc.task, dw1);
            break;
        }
    } else if (func < 5500) {
        /* ITaskTrigger */
        if (ObjToOpaque(interp, objv[0], (void **)&ifc.tasktrigger,
                        "ITaskTrigger") != TCL_OK)
            return TCL_ERROR;

        switch (func) {
        case 5401: // GetTrigger
            if (objc != 1)
                goto badargs;
            hr = ifc.tasktrigger->lpVtbl->GetTrigger(ifc.tasktrigger,
                                                     &tasktrigger);
            if (hr != S_OK)
                break;
            result.type = TRT_OBJ;
            result.value.obj = ObjFromTASK_TRIGGER(&tasktrigger);
            break;
        case 5402: // GetTriggerString
            if (objc != 1)
                goto badargs;
            result.type = TRT_LPOLESTR;
            hr = ifc.tasktrigger->lpVtbl->GetTriggerString(
                ifc.tasktrigger, &result.value.lpolestr);
            break;
        case 5403: // SetTrigger
            if (objc != 2)
                goto badargs;
            if (ObjToTASK_TRIGGER(interp, objv[1], &tasktrigger) != TCL_OK)
                return TCL_ERROR;
            result.type = TRT_EMPTY;
            hr = ifc.tasktrigger->lpVtbl->SetTrigger(ifc.tasktrigger,
                                                     &tasktrigger);
            break;
        }
    }

    if (hr != S_OK) {
        result.type = TRT_EXCEPTION_ON_ERROR;
        result.value.ival = hr;
    }

    /* Note when hr == 0, result.type can be BADFUNCTION code! */
    return TwapiSetResult(interp, &result);

badargs:
    TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);
    return TCL_ERROR;
}

static int TwapiMstaskInitCalls(Tcl_Interp *interp, TwapiInterpContext *ticP)
{
    static struct fncode_dispatch_s MstaskDispatch[] = {
        DEFINE_FNCODE_CMD(ITaskScheduler_Activate, 5001),
        DEFINE_FNCODE_CMD(ITaskScheduler_AddWorkItem, 5002),
        DEFINE_FNCODE_CMD(itaskscheduler_delete_task, 5003), //ITaskScheduler_Delete
        DEFINE_FNCODE_CMD(ITaskScheduler_Enum, 5004),
        DEFINE_FNCODE_CMD(ITaskScheduler_IsOfType, 5005),
        DEFINE_FNCODE_CMD(ITaskScheduler_NewWorkItem, 5006),
        DEFINE_FNCODE_CMD(itaskscheduler_set_target_system, 5007), //ITaskScheduler_SetTargetComputer
        DEFINE_FNCODE_CMD(itaskscheduler_get_target_system, 5008), //ITaskScheduler_GetTargetComputer

        DEFINE_FNCODE_CMD(IEnumWorkItems_Clone, 5101),
        DEFINE_FNCODE_CMD(IEnumWorkItems_Reset, 5102),
        DEFINE_FNCODE_CMD(IEnumWorkItems_Skip, 5103),
        DEFINE_FNCODE_CMD(IEnumWorkItems_Next, 5104),

        DEFINE_FNCODE_CMD(itask_new_itasktrigger, 5201), //IScheduledWorkItem_CreateTrigger
        DEFINE_FNCODE_CMD(itask_delete_itasktrigger, 5202), //IScheduledWorkItem_DeleteTrigger
        DEFINE_FNCODE_CMD(IScheduledWorkItem_EditWorkItem, 5203),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetAccountInformation, 5204),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetComment, 5205),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetCreator, 5206),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetErrorRetryCount, 5207),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetErrorRetryInterval, 5208),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetExitCode, 5209),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetFlags, 5210),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetIdleWait, 5211),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetMostRecentRunTime, 5212),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetNextRunTime, 5213),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetStatus, 5214),
        DEFINE_FNCODE_CMD(itask_get_itasktrigger, 5215), //IScheduledWorkItem_GetTrigger
        DEFINE_FNCODE_CMD(itask_get_itasktrigger_count, 5216), //IScheduledWorkItem_GetTriggerCount
        DEFINE_FNCODE_CMD(itask_get_itasktrigger_string, 5217), //IScheduledWorkItem_GetTriggerString
        DEFINE_FNCODE_CMD(itask_run, 5218), //IScheduledWorkItem_Run
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetAccountInformation, 5219),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetComment, 5220),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetCreator, 5221),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetErrorRetryCount, 5222),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetErrorRetryInterval, 5223),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetFlags, 5224),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetIdleWait, 5225),
        DEFINE_FNCODE_CMD(itask_end, 5226), //IScheduledWorkItem_Terminate
        DEFINE_FNCODE_CMD(IScheduledWorkItem_SetWorkItemData, 5227),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetWorkItemData, 5228),
        DEFINE_FNCODE_CMD(IScheduledWorkItem_GetRunTimes, 5229),

        DEFINE_FNCODE_CMD(ITask_GetApplicationName, 5301),
        DEFINE_FNCODE_CMD(ITask_GetMaxRunTime, 5302),
        DEFINE_FNCODE_CMD(ITask_GetParameters, 5303),
        DEFINE_FNCODE_CMD(ITask_GetPriority, 5304),
        DEFINE_FNCODE_CMD(ITask_GetTaskFlags, 5305),
        DEFINE_FNCODE_CMD(ITask_GetWorkingDirectory, 5306),
        DEFINE_FNCODE_CMD(ITask_SetApplicationName, 5307),
        DEFINE_FNCODE_CMD(ITask_SetParameters, 5308),
        DEFINE_FNCODE_CMD(ITask_SetWorkingDirectory, 5309),
        DEFINE_FNCODE_CMD(ITask_SetMaxRunTime, 5310),
        DEFINE_FNCODE_CMD(ITask_SetPriority, 5311),
        DEFINE_FNCODE_CMD(ITask_SetTaskFlags, 5312),

        DEFINE_FNCODE_CMD(ITaskTrigger_GetTrigger, 5401),
        DEFINE_FNCODE_CMD(ITaskTrigger_GetTriggerString, 5402),
        DEFINE_FNCODE_CMD(ITaskTrigger_SetTrigger, 5403),
    };

    TwapiDefineFncodeCmds(interp, ARRAYSIZE(MstaskDispatch), MstaskDispatch, Twapi_MstaskCallObjCmd);

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
int Twapi_mstask_Init(Tcl_Interp *interp)
{
    static TwapiModuleDef gModuleDef = {
        MODULENAME,
        TwapiMstaskInitCalls,
        NULL
    };
    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }

    return TwapiRegisterModule(interp, MODULE_HANDLE, &gModuleDef, DEFAULT_TIC) ? TCL_OK : TCL_ERROR;
}

