/* 
 * Copyright (c) 2006-2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/*
 * Define interface to Windows API related MS Task
 */

#include "twapi.h"

Tcl_Obj *ObjFromTASK_TRIGGER(TASK_TRIGGER *triggerP)
{
    Tcl_Obj *resultObj;
    Tcl_Obj *typeObj[4];
    int      ntype;

    resultObj = Tcl_NewListObj(0, 0);

    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, Reserved1);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wBeginYear);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wBeginMonth);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wBeginDay);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wEndYear);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wEndMonth);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wEndDay);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wStartHour);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wStartMinute);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, MinutesDuration);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, MinutesInterval);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, rgFlags);

    ntype = 1;
    typeObj[0] = Tcl_NewIntObj(triggerP->TriggerType);

    switch (triggerP->TriggerType) {
    case 1:
        typeObj[1] = Tcl_NewIntObj(triggerP->Type.Daily.DaysInterval);
        ntype = 2;
        break;
    case 2:
        typeObj[1] = Tcl_NewIntObj(triggerP->Type.Weekly.WeeksInterval);
        typeObj[2] = Tcl_NewIntObj(triggerP->Type.Weekly.rgfDaysOfTheWeek);
        ntype = 3;
        break;
    case 3:
        typeObj[1] = Tcl_NewIntObj(triggerP->Type.MonthlyDate.rgfDays);
        typeObj[2] = Tcl_NewIntObj(triggerP->Type.MonthlyDate.rgfMonths);
        ntype = 3;
        break;
    case 4:
        typeObj[1] = Tcl_NewIntObj(triggerP->Type.MonthlyDOW.wWhichWeek);
        typeObj[2] = Tcl_NewIntObj(triggerP->Type.MonthlyDOW.rgfDaysOfTheWeek);
        typeObj[3] = Tcl_NewIntObj(triggerP->Type.MonthlyDOW.rgfMonths);
        ntype = 4;
        break;
    }
    Tcl_ListObjAppendElement(NULL, resultObj, STRING_LITERAL_OBJ("type"));
    Tcl_ListObjAppendElement(NULL, resultObj, Tcl_NewListObj(ntype, typeObj));

    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, Reserved2);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(NULL, resultObj, triggerP, wRandomMinutesInterval);
    
    return resultObj;
}


int ObjToTASK_TRIGGER(Tcl_Interp *interp, Tcl_Obj *obj, TASK_TRIGGER *triggerP)
{
    int       i;
    Tcl_Obj **objv;
    int       objc;
    long      dval;

    if (Tcl_ListObjGetElements(interp, obj, &objc, &objv) != TCL_OK)
        return TCL_ERROR;

    if (objc & 1) {
        Tcl_SetResult(interp, "Invalid TASK_TRIGGER format - must have even number of elements", TCL_STATIC);
        return TCL_ERROR;
    }

    TwapiZeroMemory(triggerP, sizeof(*triggerP));
    triggerP->cbTriggerSize = sizeof(*triggerP);
    
    for (i=0; i < (objc-1); i+=2) {
        char *name = Tcl_GetString(objv[i]);
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
            if (Tcl_GetLongFromObj(interp, objv[i+1], &triggerP->MinutesDuration) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "MinutesInterval")) {
            if (Tcl_GetLongFromObj(interp, objv[i+1], &triggerP->MinutesInterval) != TCL_OK)
                return TCL_ERROR;
        }
        else if (STREQ(name, "rgFlags")) {
            if (Tcl_GetLongFromObj(interp, objv[i+1], &triggerP->rgFlags) != TCL_OK)
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
            if (Tcl_ListObjGetElements(interp, objv[i+1], &ntype, &typeObj) != TCL_OK) {
                return TCL_ERROR;
            }
            if (ntype == 0) {
            trigger_type_count_error:
                Tcl_SetResult(interp, "Invalid task trigger type format", TCL_STATIC);
                return TCL_ERROR;
            }

            /* First element is the type */
            if (Tcl_GetLongFromObj(interp, typeObj[0], &dval) != TCL_OK) {
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
                if (Tcl_GetLongFromObj(interp, typeObj[1], &triggerP->Type.MonthlyDate.rgfDays) != TCL_OK)
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
        objv[0] = Tcl_NewBooleanObj(1);
        objv[1] = Tcl_NewListObj(0, NULL);
        Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
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

    objv[0] = Tcl_NewBooleanObj(hr == S_OK); // More to come?
    objv[1] = Tcl_NewListObj(0, NULL);
    if (jobsP) {
        for (i = 0; i < ret_count; ++i) {
            Tcl_ListObjAppendElement(interp, objv[1],
                                     Tcl_NewUnicodeObj(jobsP[i], -1));
            /* Free the string */
            CoTaskMemFree(jobsP[i]);
        }
        /* Free the array itself */
        CoTaskMemFree(jobsP);
    }    

    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
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
        objv[1] = Tcl_NewListObj(0, NULL);
        for (i = 0; i < count; ++i) {
            Tcl_ListObjAppendElement(interp, objv[1],
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

    Tcl_SetObjResult(interp, Tcl_NewListObj(objv[1] ? 2 : 1, objv));
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
        Tcl_SetObjResult(interp, Tcl_NewByteArrayObj(dataP, count));
        return TCL_OK;
    }
    else {
        TWAPI_STORE_COM_ERROR(interp, hr, swiP, &IID_IScheduledWorkItem);
        return TCL_ERROR;
    }
}

