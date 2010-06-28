/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

/* Routines for directory change notifications */

#include "twapi.h"

#define MAXPATTERNS 32

typedef struct _TwapiDirectoryMonitorBuffer {
    OVERLAPPED ovl;
    int        buf_sz;          /* Actual size of buf[] */
    char       buf[1];          /* Variable sized */
}  TwapiDirectoryMonitorBuffer;

/*
 * Struct used to hold dir change notification context.
 */
typedef struct _TwapiDirectoryMonitorContext {
    TwapiInterpContext *ticP;
    int     id;
    HANDLE  directory_handle;   /* Handle to dir we are monitoring */
    HANDLE  thread_pool_registry_handle; /* Handle returned by thread pool */
    HANDLE  completion_event;            /* Used for i/o signaling */
    ZLINK_DECL(TwapiDirectoryMonitorContext);
    int     nrefs;               /* Ref count - not interlocked because
                                   only accessed from one thread at a time */
    TwapiDirectoryMonitorBuffer *iobP; /* Used for actual reads */
    WCHAR   *pathP;
    DWORD   filter;
    int     include_subtree;
    int     npatterns;
    char   *patterns[1];
    /* VARIABLE SIZE AREA FOLLOWS */
} TwapiDirectoryMonitorContext;


/*
 * Static prototypes
 */
static TwapiDirectoryMonitorContext *TwapiDirectoryMonitorContextNew(
    TwapiInterpContext *ticP,
    int id,
    LPWSTR pathP,                /* May NOT be null terminated if path_len==-1 */
    int    path_len,            /* -1 -> null terminated */
    int    include_subtree,
    DWORD  filter,
    char  **patterns,
    int    npatterns
    );
static void TwapiDirectoryMonitorContextDelete(TwapiDirectoryMonitorContext *);
static DWORD TwapiDirectoryMonitorInitiateRead(TwapiDirectoryMonitorContext *);
#define TwapiDirectoryMonitorContextRef(p_, incr_) InterlockedExchangeAdd(&(p_)->nrefs, (incr_))
void TwapiDirectoryMonitorContextUnref(TwapiDirectoryMonitorContext *dcmP, int decr);
VOID CALLBACK TwapiDirectoryMonitorThreadPoolCallback(
    PVOID lpParameter,
    BOOLEAN TimerOrWaitFired
);



int Twapi_RegisterDirectoryMonitor(TwapiInterpContext *ticP, int objc, Tcl_Obj *CONST objv[])
{
    TwapiDirectoryMonitorContext *dmcP;
    int    id;
    LPWSTR pathP;
    int    path_len;
    int    include_subtree;
    int    npatterns;
    char  *patterns[MAXPATTERNS];
    DWORD  winerr;
    DWORD filter;

    ERROR_IF_UNTHREADED(ticP->interp);

    if (TwapiGetArgs(ticP->interp, objc, objv,
                     GETINT(id),
                     GETWSTRN(pathP, path_len), GETBOOL(include_subtree),
                     GETINT(filter),
                     GETAARGV(patterns, ARRAYSIZE(patterns), npatterns),
                     ARGEND)
        != TCL_OK)
        return TCL_ERROR;
        
    dmcP = TwapiDirectoryMonitorContextNew(ticP, id, pathP, path_len, include_subtree, filter, patterns, npatterns);
    
    /* 
     * Should we add FILE_SHARE_DELETE to allow deleting of
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

    /* Create an event to use for notification of completion */
    dmcP->completion_event = CreateEvent(
        NULL,                   /* No security attrs */
        TRUE,                   /* Manual reset */
        FALSE,                  /* Not Signaled */
        NULL);                  /* Unnamed event */

    if (dmcP->completion_event == INVALID_HANDLE_VALUE) {
        winerr = GetLastError();
        goto system_error;
    }

    winerr = TwapiDirectoryMonitorInitiateRead(dmcP);
    if (winerr != ERROR_SUCCESS)
        goto system_error;

    /*
     * Add to list of registered handles, BEFORE we register the wait.
     * We Ref by 2 - one for the list, and one for passing it to the
     * thread pool below
     */
    TwapiDirectoryMonitorContextRef(dmcP, 2);
    EnterCriticalSection(&ticP->lock);
    ZLIST_PREPEND(&ticP->directory_monitors, dmcP);
    LeaveCriticalSection(&ticP->lock);

    /* Finally, ask thread pool to wait on the event */
    if (RegisterWaitForSingleObject(
            &dmcP->thread_pool_registry_handle,
            dmcP->completion_event,
            TwapiDirectoryMonitorThreadPoolCallback,
            dmcP,
            INFINITE,           /* No timeout */
            WT_EXECUTEINIOTHREAD
            ))
        return TCL_OK;

    /* Uh-oh, undo everything */
    winerr = GetLastError();
    
    EnterCriticalSection(&ticP->lock);
    ZLIST_REMOVE(&ticP->directory_monitors, dmcP);
    LeaveCriticalSection(&ticP->lock);

system_error:
    /* winerr should contain system error, waits should not registered */
    /* Need to close handles before deleting */
    if (dmcP->directory_handle)
        CloseHandle(dmcP->directory_handle);
    if (dmcP->completion_event)
        CloseHandle(dmcP->completion_event);

    TwapiDirectoryMonitorContextDelete(dmcP);
    return Twapi_AppendSystemError(ticP->interp, winerr);
}


/* Always returns non-NULL, or panics */
static TwapiDirectoryMonitorContext *TwapiDirectoryMonitorContextNew(
    TwapiInterpContext *ticP,
    int id,
    LPWSTR pathP,                /* May NOT be null terminated if path_len==-1 */
    int    path_len,            /* -1 -> null terminated */
    int    include_subtree,
    DWORD  filter,
    char  **patterns,
    int    npatterns
    )
{
    int sz;
    TwapiDirectoryMonitorContext *dmcP;
    int i;
    int lengths[MAXPATTERNS];
    char *cP;

    if (npatterns > ARRAYSIZE(lengths)) {
        /* It was caller's responsibility to check */
        Tcl_Panic("Internal error: caller exceeded pattern limit.");
    }

    /* Calculate the size of the structure required */
    sz = sizeof(TwapiDirectoryMonitorContext);
    if (path_len < 0)
        path_len = lstrlenW(pathP);

    sz += npatterns * sizeof(dmcP->patterns[0]); /* Space for patterns array */
    for (i=0; i < npatterns; ++i) {
        lengths[i] = lstrlenA(patterns[i]) + 1;
        sz += lengths[i];
    }
    sz += sizeof(WCHAR);                /* Sufficient space for alignment pad */
    sz += sizeof(WCHAR) * (path_len+1); /* Space for path to be monitored */

    dmcP = (TwapiDirectoryMonitorContext *) TwapiAlloc(sz);
    dmcP->ticP = ticP;
    if (ticP)
        TwapiInterpContextRef(ticP, 1);
    dmcP->id = id;
    dmcP->directory_handle = INVALID_HANDLE_VALUE;
    dmcP->thread_pool_registry_handle = INVALID_HANDLE_VALUE;
    dmcP->nrefs = 0;
    dmcP->filter = filter;
    dmcP->include_subtree = include_subtree;
    dmcP->npatterns = npatterns;
    cP = sizeof(TwapiDirectoryMonitorContext) +
        (npatterns*sizeof(dmcP->patterns[0])) +
        (char *)dmcP;
    for (i=0; i < npatterns; ++i) {
        dmcP->patterns[i] = cP;
        MoveMemory(cP, patterns[i], lengths[i]);
        cP += lengths[i];
    }
    /* Align up to WCHAR boundary */
    dmcP->pathP = (WCHAR *)((sizeof(WCHAR)-1 + (DWORD_PTR)cP) & ~ (sizeof(WCHAR)-1));
    MoveMemory(dmcP->pathP, pathP, path_len);
    dmcP->pathP[path_len] = 0;
    
    ZLINK_INIT(dmcP);
    return dmcP;
}

static void TwapiDirectoryMonitorContextDelete(TwapiDirectoryMonitorContext *dmcP)
{
    // TBD - assert nrefs <= 0
    // TBD - assert thread_pool_registry_handle == NULL
    // TBD - assert (dmcP->directory_handle == INVALID_HANDLE_VALUE)
    //       else overlapped i/o may be pending
    // dmcP->completion_event == INVALID_HANDLE_VALUE)

    if (dmcP->iobP)
        TwapiFree(dmcP->iobP);

    if (dmcP->ticP)
        TwapiInterpContextUnref(dmcP->ticP, 1);

    TwapiFree(dmcP);
}

static DWORD TwapiDirectoryMonitorInitiateRead(
    TwapiDirectoryMonitorContext *dmcP
    )
{
    TwapiDirectoryMonitorBuffer *iobP = dmcP->iobP;
    if (iobP == NULL) {
        /* TBD - add config var for buffer size. Note larger buffer
         * potential waste as well as use up precious non-paged pool
         */
        iobP = (TwapiDirectoryMonitorBuffer *)TwapiAlloc(sizeof(*iobP) + 4000);
        iobP->buf_sz = 4000;
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
    else
        return GetLastError();
}

void TwapiDirectoryMonitorContextUnref(TwapiDirectoryMonitorContext *dcmP, int decr)
{
    /* Note the ref count may be < 0 if this function is called
       on newly initialized pcbP */
    if (InterlockedExchangeAdd(&dcmP->nrefs, -decr) <= decr)
        TwapiDirectoryMonitorContextDelete(dcmP);
}


