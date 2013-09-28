/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Routines for directory change notifications */

#include "twapi.h"
#include "twapi_storage.h"

#define MAXPATTERNS 32


/*
 * Static prototypes
 */
static TwapiDirectoryMonitorContext *TwapiDirectoryMonitorContextNew(
    LPWSTR pathP, int path_len, int include_subtree,
    DWORD  filter, WCHAR **patterns, int npatterns);
static void TwapiDirectoryMonitorContextDelete(TwapiDirectoryMonitorContext *);
static DWORD TwapiDirectoryMonitorInitiateRead(TwapiDirectoryMonitorContext *);
#define TwapiDirectoryMonitorContextRef(p_, incr_) InterlockedExchangeAdd(&(p_)->nrefs, (incr_))
void TwapiDirectoryMonitorContextUnref(TwapiDirectoryMonitorContext *dcmP, int decr);
static void CALLBACK TwapiDirectoryMonitorThreadPoolFn(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
);
static int TwapiDirectoryMonitorCallbackFn(TwapiCallback *p);
static int TwapiDirectoryMonitorPatternMatch(WCHAR *path, WCHAR *pattern);

TCL_RESULT Twapi_RegisterDirectoryMonitorObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    TwapiDirectoryMonitorContext *dmcP;
    LPWSTR pathP;
    int    path_len;
    int    include_subtree;
    int    npatterns;
    WCHAR  **patterns;
    DWORD  winerr;
    DWORD filter;
    MemLifoMarkHandle mark;

    ERROR_IF_UNTHREADED(interp);

    mark = MemLifoPushMark(&ticP->memlifo);
    if (TwapiGetArgsEx(ticP, objc-1, objv+1,
                     GETWSTRN(pathP, path_len), GETBOOL(include_subtree),
                     GETINT(filter),
                     GETARGVW(patterns, npatterns),
                     ARGEND)
        != TCL_OK) {
        MemLifoPopMark(mark);
        return TCL_ERROR;
    }        
    dmcP = TwapiDirectoryMonitorContextNew(pathP, path_len, include_subtree, filter, patterns, npatterns);

    MemLifoPopMark(mark);       /* Don't need parsed args any more */
    
    /* 
     * TBD - Should we add FILE_SHARE_DELETE to allow deleting of
     * the directory? For now, no because it causes confusing behaviour.
     * The directory being monitored can be deleted successfully but
     * an attempt to create a directory of that same name will then
     * fail mysteriously (from the user point of view) with access
     * denied errors. Also, no notification is sent about the deletion
     * unless the parent dir is also being monitored.
     * TBD - caller has to have the SE_BACKUP_NAME and SE_RESTORE_NAME
     * privs
     */
    dmcP->directory_handle = CreateFileW(
        dmcP->pathP,
        FILE_LIST_DIRECTORY,
        FILE_SHARE_READ | FILE_SHARE_WRITE,
        NULL,
        OPEN_EXISTING,
        FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED,
        NULL);

    if (dmcP->directory_handle == INVALID_HANDLE_VALUE) {
        winerr = GetLastError();
        goto system_error;
    }

    /*
     * Create an event to use for notification of completion. The event
     * must be auto-reset to prevent multiple callback queueing on a single 
     * input notification. See MSDN docs for RegisterWaitForSingleObject.
     */
    dmcP->completion_event = CreateEvent(
        NULL,                   /* No security attrs */
        FALSE,                  /* Auto reset */
        FALSE,                  /* Not Signaled */
        NULL);                  /* Unnamed event */

    if (dmcP->completion_event == NULL) {
        winerr = GetLastError();
        goto system_error;
    }

    winerr = TwapiDirectoryMonitorInitiateRead(dmcP);
    if (winerr != ERROR_SUCCESS)
        goto system_error;

    /*
     * Add to list of registered handles, BEFORE we register the wait.
     * We Ref by 2 - one for the list, and one for passing it to the
     * thread pool below.
     * Note we do not lock for access to the list as it will be
     * accessed only from this interp thread, never from thread pool
     * or another interp.
     */
    TwapiDirectoryMonitorContextRef(dmcP, 2);
    ZLIST_PREPEND(&(((TwapiStorageInterpContext *)ticP->module.data.pval)->directory_monitors), dmcP);
    dmcP->ticP = ticP;
    TwapiInterpContextRef(ticP, 1);

    /* Finally, ask thread pool to wait on the event */
    if (RegisterWaitForSingleObject(
            &dmcP->thread_pool_registry_handle,
            dmcP->completion_event,
            TwapiDirectoryMonitorThreadPoolFn,
            dmcP,
            INFINITE,           /* No timeout */
            WT_EXECUTEINIOTHREAD
            )) {
        ObjSetResult(interp, ObjFromHANDLE(dmcP->directory_handle));
        return TCL_OK;
    }

    /* Uh-oh, undo everything */
    winerr = GetLastError();
    
    ZLIST_REMOVE(&(((TwapiStorageInterpContext *)ticP->module.data.pval)->directory_monitors), dmcP);

system_error:
    /* winerr should contain system error, waits should not registered */
    TwapiShutdownDirectoryMonitor(dmcP);
    return Twapi_AppendSystemError(interp, winerr);
}


TCL_RESULT Twapi_UnregisterDirectoryMonitorObjCmd(TwapiInterpContext *ticP, Tcl_Interp *interp, int objc, Tcl_Obj *CONST objv[])
{
    HANDLE dirhandle;
    TwapiDirectoryMonitorContext *dmcP;

    if (TwapiGetArgs(interp, objc-1, objv+1, GETHANDLE(dirhandle),
                     ARGEND) != TCL_OK)
        return TCL_ERROR;

    /* 
     * Look up the handle in list of directory monitors. No locking
     * required as list access always done in thread of interpreter
     */

    // ASSERT ticP->thread == current thread
    
    ZLIST_LOCATE(dmcP, &(((TwapiStorageInterpContext *)ticP->module.data.pval)->directory_monitors), directory_handle, dirhandle);
    if (dmcP == NULL)
        return TwapiReturnError(ticP->interp, TWAPI_UNKNOWN_OBJECT);

    // ASSERT ticP == dmcP->ticP

    TwapiShutdownDirectoryMonitor(dmcP);

    return TCL_OK;
}



/* Always returns non-NULL, or panics */
static TwapiDirectoryMonitorContext *TwapiDirectoryMonitorContextNew(
    LPWSTR pathP,                /* May NOT be null terminated if path_len==-1 */
    int    path_len,            /* -1 -> null terminated */
    int    include_subtree,
    DWORD  filter,
    WCHAR **patterns,
    int    npatterns
    )
{
    int sz;
    TwapiDirectoryMonitorContext *dmcP;
    int i;
    int bytelengths[MAXPATTERNS];
    char *cP;

    if (npatterns > ARRAYSIZE(bytelengths)) {
        /* It was caller's responsibility to check */
        Tcl_Panic("Internal error: caller exceeded pattern limit.");
    }

    /* Calculate the size of the structure required */
    sz = sizeof(TwapiDirectoryMonitorContext);
    if (path_len < 0)
        path_len = lstrlenW(pathP);

    sz += npatterns * sizeof(dmcP->patterns[0]); /* Space for patterns array */
    for (i=0; i < npatterns; ++i) {
        bytelengths[i] = sizeof(WCHAR) * (lstrlenW(patterns[i]) + 1);
        sz += bytelengths[i];
    }
    sz += sizeof(WCHAR) * (path_len+1); /* Space for path to be monitored */

    dmcP = (TwapiDirectoryMonitorContext *) TwapiAlloc(sz);
    dmcP->ticP = NULL;
    dmcP->directory_handle = INVALID_HANDLE_VALUE;
    dmcP->completion_event = NULL;
    dmcP->thread_pool_registry_handle = INVALID_HANDLE_VALUE;
    dmcP->nrefs = 0;
    dmcP->filter = filter;
    dmcP->include_subtree = include_subtree;
    dmcP->npatterns = npatterns;
    cP = sizeof(TwapiDirectoryMonitorContext) +
        (npatterns*sizeof(dmcP->patterns[0])) +
        (char *)dmcP;
    for (i=0; i < npatterns; ++i) {
        dmcP->patterns[i] = (WCHAR *) cP;
        CopyMemory(cP, patterns[i], bytelengths[i]);
        cP += bytelengths[i];
    }
    TWAPI_ASSERT((((DWORD_PTR)cP) & 1) == 0); /* Alignment check */
    dmcP->pathP = (WCHAR *) cP;
    CopyMemory(dmcP->pathP, pathP, sizeof(WCHAR)*path_len);
    dmcP->pathP[path_len] = 0;
    dmcP->iobP = NULL;
    
    ZLINK_INIT(dmcP);
    return dmcP;
}

/*
 * Deletes a context. Only does deallocation. Does not do any unlinking,
 * caller must do that before calling.
 */
static void TwapiDirectoryMonitorContextDelete(TwapiDirectoryMonitorContext *dmcP)
{
    TWAPI_ASSERT(dmcP->ticP == NULL);
    TWAPI_ASSERT(dmcP->nrefs <= 0);
    TWAPI_ASSERT(dmcP->thread_pool_registry_handle == INVALID_HANDLE_VALUE);
    TWAPI_ASSERT(dmcP->directory_handle == INVALID_HANDLE_VALUE);
    TWAPI_ASSERT(dmcP->completion_event == NULL);

    if (dmcP->iobP) {
        TWAPI_ASSERT(dmcP->iobP->ovl.hEvent == NULL); /* Else I/O could be in progress */
        TwapiFree(dmcP->iobP);
    }

    TwapiFree(dmcP);
}

static DWORD TwapiDirectoryMonitorInitiateRead(
    TwapiDirectoryMonitorContext *dmcP
    )
{
    TwapiDirectoryMonitorBuffer *iobP = dmcP->iobP;
#define DIRMON_BUFSZ 8000
    if (iobP == NULL) {
        /* TBD - add config var for buffer size. Note larger buffer
         * potential waste as well as use up precious non-paged pool
         */
        iobP = (TwapiDirectoryMonitorBuffer *)TwapiAlloc(sizeof(*iobP) + DIRMON_BUFSZ);
        iobP->buf_sz = DIRMON_BUFSZ;
        dmcP->iobP = iobP;
    }

    iobP->ovl.Internal = 0;
    iobP->ovl.InternalHigh = 0;
    iobP->ovl.Offset = 0;
    iobP->ovl.OffsetHigh = 0;
    iobP->ovl.hEvent = dmcP->completion_event;
    if (ReadDirectoryChangesW(
            dmcP->directory_handle,
            iobP->buf,
            iobP->buf_sz,
            dmcP->include_subtree,
            dmcP->filter,
            NULL,
            &iobP->ovl,
            NULL))
        return ERROR_SUCCESS;
    else {
        DWORD winerr = GetLastError();
        if (winerr == ERROR_IO_PENDING) {
            /* Can this error even be returned ? */
            return ERROR_SUCCESS;
        }

        /* Don't just hang on to memory. Also (iobP && iobP->ovl.hEvent) is
           used as the flag to indicate I/O is pending so we need to make
           sure that evaluates to false
        */
        TwapiFree(dmcP->iobP);
        dmcP->iobP = NULL;

        return winerr;
    }
}

void TwapiDirectoryMonitorContextUnref(TwapiDirectoryMonitorContext *dcmP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized pcbP */
    if (InterlockedExchangeAdd(&dcmP->nrefs, -decr) <= decr)
        TwapiDirectoryMonitorContextDelete(dcmP);
}


/*
 * Called from the Windows thread pool when a dir change notification is 
 * signaled.
 */
static void CALLBACK TwapiDirectoryMonitorThreadPoolFn(
    PVOID pv,
    BOOLEAN timeout
)
{
    TwapiDirectoryMonitorContext *dmcP = (TwapiDirectoryMonitorContext *) pv;
    DWORD bytes_read;
    DWORD winerr;
    TwapiCallback *cbP;
    TwapiDirectoryMonitorBuffer *iobP;

    /*
     * We can safely access fields in *dmcP because the owning interp
     * thread will not pull it out from under us without unregisetring
     * the handle from the thread pool, which the thread pool will block
     * until we return
     */

    if (timeout)
        Tcl_Panic("Unexpected timeout in directory monitor callback.");

    /*
     * The thread pool requires that the event we used was auto-reset to
     * prevent multiple callbacks being queued for a single operation.
     * This means the event is now in a non-signaled state. So make sure
     * the last bWait param to GetOverlappedResult is FALSE else the
     * call will hang forever (since we will not issue another read
     * until later)
     */
    if (GetOverlappedResult(dmcP->directory_handle, &dmcP->iobP->ovl, &bytes_read, 0) == FALSE) {
        /* Error. */
        winerr = GetLastError();
        if (winerr == ERROR_IO_INCOMPLETE) {
            /* Huh? then why were we signaled? But don't treat as error */
            return;
        }
        goto error_handler;
    }
    
    /*
     * Success, send the current buffer over to the interp.
     */
    iobP = dmcP->iobP;
    dmcP->iobP = NULL;
    iobP->ovl.hEvent = NULL;    /* Just to make sure event not accessed from
                                   this structure */

    cbP = TwapiCallbackNew(dmcP->ticP, TwapiDirectoryMonitorCallbackFn,
                                  sizeof(*cbP));
    cbP->winerr = ERROR_SUCCESS;
    TwapiDirectoryMonitorContextRef(dmcP, 1); /* Since iobP is being queued */
    cbP->clientdata = (DWORD_PTR) dmcP;
    cbP->clientdata2 = (DWORD_PTR) iobP;
    TwapiEnqueueCallback(dmcP->ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);
    cbP = NULL;                 /* So we do not access it below */

    /* Set up for next read */
    if ((winerr = TwapiDirectoryMonitorInitiateRead(dmcP)) != ERROR_SUCCESS) {
        goto error_handler;
    }

    return;

error_handler:
    /* Queue an error notification. winerr must hold error code */
    /* Do NOT COME HERE IF OVERLAPPED IO STILL IN PROGRESS AS THE IOBUF MAY BE FREED */
    if (dmcP->iobP) {
        TwapiFree(dmcP->iobP);
        dmcP->iobP = NULL;
    }

    cbP = TwapiCallbackNew(dmcP->ticP, TwapiDirectoryMonitorCallbackFn,
                                  sizeof(*cbP));
    cbP->winerr = winerr;
    cbP->clientdata = (DWORD_PTR) dmcP;
    cbP->clientdata2 = 0;
    /*
     * We need to Ref again, even if we might have done so above
     * since this is a second callback being queued.
     */
    TwapiDirectoryMonitorContextRef(dmcP, 1);
    TwapiEnqueueCallback(dmcP->ticP, cbP, TWAPI_ENQUEUE_DIRECT, 0, NULL);
    cbP = NULL;                 /* So we do not access it below */

    return;
}

static int TwapiDirectoryMonitorCallbackFn(TwapiCallback *cbP)
{
    Tcl_Obj *scriptObj;
    Tcl_Obj *changesObj;
    Tcl_Obj *fnObj[2];
    Tcl_Obj *actionObj[6];
    int      notify;
    FILE_NOTIFY_INFORMATION *fniP;
    TwapiDirectoryMonitorBuffer *iobP;
    char      *endP;
    int        i;
    Tcl_Interp *interp;
    TwapiDirectoryMonitorContext *dmcP;
    int        tcl_status;

    fnObj[0] = NULL;
    fnObj[1] = NULL;

    dmcP = (TwapiDirectoryMonitorContext *) cbP->clientdata;
    /*
     * Note - dmcP->iobP points to i/o buffer currently in use, do not access.
     * cbP->clientdata2 points to the i/o buffer being passed to us.
     */
    iobP = (TwapiDirectoryMonitorBuffer *) cbP->clientdata2; /* May be NULL! in error cases */
    if (dmcP->ticP == NULL ||
        dmcP->ticP->interp == NULL ||
        Tcl_InterpDeleted(dmcP->ticP->interp)) {

        /* There is no interp left. We must close the handle ourselves */
        /* iobP may be NULL if error notification */
        if (iobP)
            TwapiFree(iobP);

        cbP->clientdata2 = 0;        /* iobP */

        if (dmcP->ticP) {
            /*
             * We are holding a ref associated with the callback
             * so we know dmcP will not be deallocated in this call.
             */
            TwapiShutdownDirectoryMonitor(dmcP);
        }

        /* Unref to match ref when the callback was queued */
        TwapiDirectoryMonitorContextUnref(dmcP, 1); /* dmcP may be GONE! */
        dmcP = NULL;

        cbP->winerr = ERROR_INVALID_FUNCTION; // TBD
        cbP->response.type = TRT_EMPTY;
        return TCL_ERROR;
    }

    interp = cbP->ticP->interp;
    notify = 0;
    // TBD - can iobP be null ?

    /* The object that will hold the change list */
    changesObj = ObjEmptyList();
    ObjIncrRefs(changesObj);

    /* Object containing the callback script */
    scriptObj = ObjEmptyList();
    ObjIncrRefs(scriptObj);
    ObjAppendElement(interp, scriptObj, STRING_LITERAL_OBJ(TWAPI_TCL_NAMESPACE "::_filesystem_monitor_handler"));
    ObjAppendElement(interp, scriptObj, ObjFromHANDLE(dmcP->directory_handle));

    if (cbP->winerr != ERROR_SUCCESS) {
        /* Error notification. Script should close the monitor */
        ObjAppendElement(interp, changesObj, STRING_LITERAL_OBJ("error"));
        ObjAppendElement(interp, changesObj,
                                 Tcl_NewLongObj(cbP->winerr)); /* Error code */
        notify = 1;         /* So we notify script */
    } else {
        /* Collect the list of matching notifications */
        // TBD - assert iobP
        fniP = (FILE_NOTIFY_INFORMATION *) iobP->buf;
        /* InternalHigh is byte count */
        endP = ADDPTR(iobP->buf, iobP->ovl.InternalHigh, char*);
        notify = 0;
        fnObj[1] = NULL;        /* When looping this can contain a reusable object */
        for (i=0; i < ARRAYSIZE(actionObj); ++i) {
            actionObj[i] = NULL;
        }
        while ((endP - offsetof(FILE_NOTIFY_INFORMATION,FileName)) > (char *)fniP) {
            int pattern_matched;

            if ((fniP->FileNameLength == 0) || (fniP->FileNameLength & 1)) {
                /*
                 * Number of bytes should be positive and even. Ignore all
                 * remaining. TBD - error
                 */
                break;
            }

            /* Double check lengths are OK.  Note FileNameLength is in bytes. */
            if ((fniP->FileNameLength + (char *)fniP->FileName) > endP) {
                /* Suspect length. TBD - error */
                break;
            }

            /*
             * Skip if pattern specified and do not match. We first create
             * a Tcl_Obj and then do the match because unfortunately, the
             * the filenames in iobP are not null terminated so we would have
             * to copy them somewhere anyways to compare. We could temporarily
             * overwrite the next char with \0 but that would not work for
             * the last filename.
             */
            if (fnObj[1]) {
                /* Left over from last iteration so use it */
                Tcl_SetUnicodeObj(fnObj[1], fniP->FileName, fniP->FileNameLength/2);
            } else 
                fnObj[1] = ObjFromUnicodeN(fniP->FileName, fniP->FileNameLength/2);
            if (dmcP->npatterns == 0)
                pattern_matched = 1; /* No pattern so match everything */
            else {
                /* Need to match on pattern */
                int i;
                for (i = 0; i < dmcP->npatterns; ++i) {
                    int matched = TwapiDirectoryMonitorPatternMatch(
                        ObjToUnicode(fnObj[1]),
                        dmcP->patterns[i]);
                    if (matched) {
                        /* Matches can be inclusive or exclusive */
                        pattern_matched = (matched > 0);
                        break;
                    }
                }
                if (i == dmcP->npatterns) {
                    /*
                     * No match. Try next name. Note fnObj will be released
                     * later if necessary or reused in next iteration
                     */
                    pattern_matched = 0;
                }
            }

            if (pattern_matched) {
                /*
                 * We reuse the action names instead of allocating new ones. Much
                 * more efficient in space and time when many notifications.
                 */
                switch (fniP->Action) {
                case FILE_ACTION_ADDED:
                    if (actionObj[0] == NULL) {
                        actionObj[0] = STRING_LITERAL_OBJ("added");
                        ObjIncrRefs(actionObj[0]);
                    }
                    fnObj[0] = actionObj[0];
                    break;
                case FILE_ACTION_REMOVED:
                    if (actionObj[1] == NULL) {
                        actionObj[1] = STRING_LITERAL_OBJ("removed");
                        ObjIncrRefs(actionObj[1]);
                    }
                    fnObj[0] = actionObj[1];
                    break;
                case FILE_ACTION_MODIFIED:
                    if (actionObj[2] == NULL) {
                        actionObj[2] = STRING_LITERAL_OBJ("modified");
                        ObjIncrRefs(actionObj[2]);
                    }
                    fnObj[0] = actionObj[2];
                    break;
                case FILE_ACTION_RENAMED_OLD_NAME:
                    if (actionObj[3] == NULL) {
                        actionObj[3] = STRING_LITERAL_OBJ("renameold");
                        ObjIncrRefs(actionObj[3]);
                    }
                    fnObj[0] = actionObj[3];
                    break;
                case FILE_ACTION_RENAMED_NEW_NAME:
                    if (actionObj[4] == NULL) {
                        actionObj[4] = STRING_LITERAL_OBJ("renamenew");
                        ObjIncrRefs(actionObj[4]);
                    }
                    fnObj[0] = actionObj[4];
                    break;
                default:
                    if (actionObj[5] == NULL) {
                        actionObj[5] = STRING_LITERAL_OBJ("unknown");
                        ObjIncrRefs(actionObj[5]);
                    }
                    fnObj[0] = actionObj[5];
                    break;
                }
                ObjAppendElement(interp, changesObj, fnObj[0]);
                ObjAppendElement(interp, changesObj, fnObj[1]);
                fnObj[1] = NULL;           /* So we don't mistakenly reuse it */
                notify = 1;
            }

            if (fniP->NextEntryOffset == 0)
                break;          // No more entries

            fniP = (FILE_NOTIFY_INFORMATION *) (fniP->NextEntryOffset + (char *)fniP);
        } /* while */

        /* Clean up left over objects */
        if (fnObj[1])
            Twapi_FreeNewTclObj(fnObj[1]); /* Allocated but not used */
        /* Deref the action objs. Note if in use by lists this will not free them */
        for (i=0; i < ARRAYSIZE(actionObj); ++i) {
            if (actionObj[i])
                ObjDecrRefs(actionObj[i]);
            actionObj[i] = NULL;
        }
    } /* if error/not error */

    /* Matches Ref from when iobP was queued */
    TwapiDirectoryMonitorContextUnref(dmcP, 1);
    TwapiFree(iobP);
    dmcP = NULL;
    iobP = NULL;

    cbP->clientdata = 0;        /* dmcP */
    cbP->clientdata2 = 0;        /* iobP */

    if (notify) {
        /* File or error notification */
        int objc;
        Tcl_Obj **objv;
        ObjAppendElement(interp, scriptObj, changesObj);
        ObjGetElements(interp, scriptObj, &objc, &objv);
        tcl_status = TwapiEvalAndUpdateCallback(cbP, objc, objv, TRT_EMPTY);
        if (tcl_status != TCL_OK)
            Twapi_AppendLog(interp, L"CALLBACK FAIL");
    } else {
        /* No files matched and no error so no need to invoke callback */
        cbP->winerr = ERROR_SUCCESS;
        cbP->response.type = TRT_EMPTY;
        tcl_status = TCL_OK;
    }
    ObjDecrRefs(changesObj); /* Free up list elements */
    ObjDecrRefs(scriptObj); /* Free up list elements */
    return tcl_status;

}


/*
 * Initiates shut down of a directory monitor. It unregisters the dmc from
 * the ticP and the thread pool. Hence dmcP may be deallocated before returning
 * unless caller holds some other ref to the dmc.
 */
int TwapiShutdownDirectoryMonitor(TwapiDirectoryMonitorContext *dmcP)
{
    int unrefs = 0;             /* How many times we need to unref  */
    TwapiInterpContext *ticP;

    /*
     * We need to do things in a specific order.
     *
     * Note all unrefs for the dmc are done at the end.
     */

    /*
     * Stop the thread pool for this dmc. We need to do that before
     * closing handles. Note the UnregisterWaitEx can result in thread pool
     * callbacks running while it is blocked. The callbacks might queue
     * additional events to the interp thread. That's ok, they will be
     * discarded when removed from the event queue.
     */
    if (dmcP->thread_pool_registry_handle != INVALID_HANDLE_VALUE) {
        UnregisterWaitEx(dmcP->thread_pool_registry_handle,
                         INVALID_HANDLE_VALUE /* Wait for callbacks to finish */
            );
        dmcP->thread_pool_registry_handle = INVALID_HANDLE_VALUE;
        ++unrefs;               /* Remove the ref coming from the thread pool */
    }

    /* Now that threads have stopped, unlink the dmcP and ticP */
    ticP = dmcP->ticP;
    if (ticP && ticP->module.data.pval) {
        ZLIST_REMOVE(&(((TwapiStorageInterpContext *)ticP->module.data.pval)->directory_monitors), dmcP);
        dmcP->ticP = NULL;
        ++unrefs;   /* Since directory_monitors no longer refs dmcP */
    }

    /* Third, now that handles are unregistered, close them. */
    if (dmcP->directory_handle != INVALID_HANDLE_VALUE) {
        CloseHandle(dmcP->directory_handle);
        dmcP->directory_handle = INVALID_HANDLE_VALUE;
    }
    if (dmcP->completion_event != NULL) {
        if (dmcP->iobP && dmcP->iobP->ovl.hEvent) {
            /* Read was in progress. Wait for it to complete */
            TWAPI_ASSERT(dmcP->iobP->ovl.hEvent == dmcP->completion_event);
            if (WaitForSingleObject(dmcP->completion_event, 1000) != WAIT_OBJECT_0) {
                if (ticP && ticP->interp) {
                    Twapi_AppendLog(ticP->interp, L"WaitForSingleObject did not return WAIT_OBJECT_0 while shutting down a directory monitor");
                }
            }
            CloseHandle(dmcP->completion_event);
            dmcP->iobP->ovl.hEvent = NULL;
            dmcP->completion_event = NULL;
        }
    }

    if (ticP)
        TwapiInterpContextUnref(ticP, 1);

    if (unrefs)
        TwapiDirectoryMonitorContextUnref(dmcP, unrefs); /* May be GONE! */

    return TCL_OK;
}

/* Returns 0 if no match, 1 if inclusion match, -1 if exclusion match */
static int TwapiDirectoryMonitorPatternMatch(WCHAR *path, WCHAR *pattern)
{
    int include;

    /* Pattern can be prefixed with + or - to indicate inclusion/exclusion */
    if (pattern[0] == L'-') {
        include = -1;
        ++pattern;
    } else if (pattern[0] == L'+') {
        include = 1;
        ++pattern;
    } else
        include = 1;            /* Default is inclusive pattern */

    return Tcl_UniCharCaseMatch(path, pattern, 1) ? include : 0;
}
