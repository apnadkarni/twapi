#ifndef TWAPI_STORAGE_H
#define TWAPI_STORAGE_H

typedef struct _TwapiDirectoryMonitorContext TwapiDirectoryMonitorContext;
ZLINK_CREATE_TYPEDEFS(TwapiDirectoryMonitorContext); 
ZLIST_CREATE_TYPEDEFS(TwapiDirectoryMonitorContext);

/* We hang this off TwapiInterpContext to hold this module's data */
typedef struct _TwapiStorageInterpContext {
    /*
     * List of directory change monitor contexts. ONLY ACCESSED
     * FROM Tcl THREAD SO ACCESSED WITHOUT A LOCK.
     */
    ZLIST_DECL(TwapiDirectoryMonitorContext) directory_monitors;
} TwapiStorageInterpContext;

typedef struct _TwapiDirectoryMonitorBuffer {
    OVERLAPPED ovl;
    int        buf_sz;          /* Actual size of buf[] */
    __int64    buf[1];       /* Variable sized area. __int64 to force align
                                   to 8 bytes */
} TwapiDirectoryMonitorBuffer;

/*
 * Struct used to hold dir change notification context.
 *
 * Life of a TwapiDirectoryMonitorContext (dmc) -
 *
 *   Because of the use of the Windows thread pool, some care has to be taken
 * when managing a dmc. In particular, the dmc must not be deallocated and
 * the corresponding directory handle closed until it is unregistered with
 * the thread pool. We also have to be careful if the interpreter is deleted
 * without an explicit call to close the notification stream. The following 
 * outlines the life of a dmc -
 *   A dmc is allocated when a script calls Twapi_RegisterDirectoryMonitor.
 * Assuming no errors in setting up the notifications, it is added to the
 * interp context's directory monitor list so it can be freed if the 
 * interp is deleted. It is then passed to the thread pool to wait on
 * the directory notification event to be signaled. The reference count (nrefs)
 * is 2 at this point, one for the interp context list and one because the
 * thread pool holds a reference to it. The first of these has a corresponding
 * unref when the interp explicitly, or implicitly through interp deletion,
 * close the notification at which point the dmc is also removed from the
 * interp context's list. The second is matched with an unref when it
 * is unregistered from the thread pool, either because of one of the
 * aforementioned reasons or an error.
 *   When the thread pool reads the directory changes, it places a callback
 * on the Tcl event queue. The callback contains a pointer to the dmc
 * and therefore the dmc's ref count is updated. The corresponding dmc
 * unref is done when the event handler dispatches the callback.
 *   Note that when an error is encountered by the thread pool thread
 * when reading directory changes, it queues a callback which results in
 * the script being notified, which will then close the notification.
 *
 * Locking and synchronization -
 *
 *   The dmc may be accessed from either the interp thread, or one of
 * the thread pool threads. Moreover, since only one directory read is
 * outstanding at any time for a dmc only one thread pool thread can
 * potentially be accessing the dmc at any instant. Because of the ref
 * counting described above, a thread does not have to worry about
 * a dmc disappearing while it still has a reference to it. Only access
 * to fields within the dmc has to be synchronized. It turns out
 * no explicit syncrhnoization is necessary because all fields except
 * nrefs and iobP are initialized in the interp thread and not modified
 * again until the dmc has been unregistered from the thread pool (at
 * which time no other thread will access them). The nrefs field is of
 * course synchronized as interlocked ref count operations. Finally, the
 * the iobP field is only stored or accessed in the thread pool, never
 * in the interp thread EXCEPT when dmc is being deallocated at which
 * point the thread pool access is already shut down.
 */
typedef struct _TwapiDirectoryMonitorContext {
    TwapiInterpContext *ticP;
    HANDLE  directory_handle;   /* Handle to dir we are monitoring */
    HANDLE  thread_pool_registry_handle; /* Handle returned by thread pool */
    HANDLE  completion_event;            /* Used for i/o signaling */
    ZLINK_DECL(TwapiDirectoryMonitorContext);
    long volatile nrefs;              /* Ref count */
    TwapiDirectoryMonitorBuffer *iobP; /* Used for actual reads. IMPORTANT -
                                          if not NULL, and iobP->ovl.hEvent
                                          is not NULL, READ is in progress
                                          and even when closing handle
                                          we have to wait for event to be
                                          signalled.
                                       */
    WCHAR   *pathP;
    DWORD   filter;
    int     include_subtree;
    int     npatterns;
    WCHAR  *patterns[1];
    /* VARIABLE SIZE AREA FOLLOWS */
} TwapiDirectoryMonitorContext;


TwapiTclObjCmd Twapi_RegisterDirectoryMonitorObjCmd;
TwapiTclObjCmd Twapi_UnregisterDirectoryMonitorObjCmd;
int TwapiShutdownDirectoryMonitor(TwapiDirectoryMonitorContext *);

#endif
