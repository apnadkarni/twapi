/*
 * Copyright (c) 2010, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include <ntverp.h>             /* Needed for VER_PRODUCTBUILD SDK version */
#include "tclTomMath.h"

#if _WIN32_WINNT < 0x0500
#error _WIN32_WINNT too low
#endif

        
#define TWAPI_TCL_MAJOR 8
#ifdef _WIN64
#define TWAPI_MIN_TCL_MINOR 5
#else
#define TWAPI_MIN_TCL_MINOR 4
#endif

/*
 * Globals
 */
HMODULE gTwapiModuleHandle;     /* DLL handle to ourselves */

static const char *gTwapiEmbedType = "none"; /* Must point to static string */

OSVERSIONINFO gTwapiOSVersionInfo;
GUID gTwapiNullGuid;             /* Initialized to all zeroes */
struct TwapiTclVersion gTclVersion;
static int gTclIsThreaded;
static DWORD gTlsIndex;         /* As returned by TlsAlloc */
static ULONG volatile gTlsNextSlot;  /* Index into private slots in Tls area. */

/* List of allocated interpreter - used primarily for unnotified cleanup */
CRITICAL_SECTION gTwapiInterpContextsCS; /* To protect the same */
ZLIST_DECL(TwapiInterpContext) gTwapiInterpContexts;

/* Used to generate unique id's */
TwapiId volatile gIdGenerator;

/*
 * Whether the callback dll/libray has been initialized.
 * The value must be managed using the InterlockedCompareExchange functions to
 * ensure thread safety. The value returned by InterlockedCompareExhange
 * 0 -> first to call, do init,  1 -> init in progress by some other thread
 * 2 -> Init done
 */
static TwapiOneTimeInitState gTwapiInitialized;

static void Twapi_Cleanup(ClientData clientdata);
static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp);
static TwapiInterpContext *TwapiInterpContextNew(Tcl_Interp *, HMODULE, TwapiInterpContextCleanup * );
static void TwapiInterpContextDelete(TwapiInterpContext *ticP);
static int TwapiOneTimeInit(Tcl_Interp *interp);
static TCL_RESULT TwapiLoadInitScript(TwapiInterpContext *ticP);

#ifndef TWAPI_STATIC_BUILD
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gTwapiModuleHandle = hmod;
    return TRUE;
}
#endif

/* Why is this decl needed ? TBD */
const char * __cdecl Tcl_InitStubs(Tcl_Interp *interp, const char *,int);

/* Main entry point */
#ifndef TWAPI_STATIC_BUILD
__declspec(dllexport) 
#endif
int Twapi_base_Init(Tcl_Interp *interp)
{
    static LONG twapi_initialized; /* TBD - where used ? */
    TwapiInterpContext *ticP;

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs - should this be the
       done for EVERY interp creation or move into one-time above ? TBD
     */
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }
    if (Tcl_TomMath_InitStubs(interp, 0) == NULL) {
        return TCL_ERROR;
    }

    /* Init unless already done. */
    if (! TwapiDoOneTimeInit(&gTwapiInitialized, TwapiOneTimeInit, interp))
        return TCL_ERROR;

    /* NOTE: no point setting Tcl_SetResult for errors as they are not
       looked at when DLL is being loaded */

    /*
     * Per interp initialization
     */

    /*
     * Create the name space and some variables. 
     * Needed for some scripts bound into the dll
     */
    Tcl_Eval(interp, "namespace eval " TWAPI_TCL_NAMESPACE " { variable settings ; set settings(log_limit) 100}");

    /* Allocate a context that will be passed around in all interpreters */
    /* TBD - last param should not be NULL - add a cleanup instead of
       hardcoding in Twapi_InterpCleanup */
    ticP = Twapi_AllocateInterpContext(interp, gTwapiModuleHandle, NULL);
    if (ticP == NULL)
        return TCL_ERROR;

    /* Do our own commands. */
    if (Twapi_InitCalls(interp, ticP) != TCL_OK) {
        return TCL_ERROR;
    }

    /* TBD - move these commands into the calls.c infrastructure */
    Tcl_CreateObjCommand(interp, "twapi::parseargs", Twapi_ParseargsObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::trap", Twapi_TryObjCmd,
                         ticP, NULL);
    Tcl_CreateAlias(interp, "twapi::try", interp, "twapi::trap", 0, NULL);
    Tcl_CreateObjCommand(interp, "twapi::kl_get", Twapi_KlGetObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::twine", Twapi_TwineObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::recordarray", Twapi_RecordArrayObjCmd,
                         ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::GetTwapiBuildInfo",
                         Twapi_GetTwapiBuildInfo, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::tclcast", Twapi_InternalCastObjCmd, ticP, NULL);
    Tcl_CreateObjCommand(interp, "twapi::tcltype", Twapi_GetTclTypeObjCmd, ticP, NULL);
    
    if (TwapiLoadInitScript(ticP) != TCL_OK) {
        /* We keep going as scripts might be external, not bound into DLL */
        /* return TCL_ERROR; */
        Tcl_ResetResult(interp); /* Get rid of any error messages */
    }

    return TCL_OK;
}

/*
 * Loads the initialization script from image file resource
 */
TCL_RESULT Twapi_SourceResource(TwapiInterpContext *ticP, HANDLE dllH, const char *name)
{
    HRSRC hres = NULL;
    unsigned char *dataP;
    DWORD sz;
    HGLOBAL hglob;
    int result;
    int compressed;

    /*
     * Locate the twapi resource and load it if found. First check for
     * uncompressed type. Then compressed.
     */
    compressed = 0;
    hres = FindResource(dllH,
                        name,
                        TWAPI_SCRIPT_RESOURCE_TYPE);
    if (!hres) {
        hres = FindResource(dllH,
                            name,
                            TWAPI_SCRIPT_RESOURCE_TYPE_LZMA);
        compressed = 1;
    }

    if (!hres)
        return Twapi_AppendSystemError(ticP->interp, GetLastError());

    sz = SizeofResource(dllH, hres);
    hglob = LoadResource(dllH, hres);
    if (sz && hglob) {
        dataP = LockResource(hglob);
        if (dataP) {
            /* If compressed, we need to uncompress it first */
            if (compressed) {
                dataP = TwapiLzmaUncompressBuffer(ticP, dataP, sz, &sz);
                if (dataP == NULL)
                    return TCL_ERROR; /* ticP->interp already has error */
            }

            /* The resource is expected to be UTF-8 (actually strict ASCII) */
            /* TBD - double check use of GLOBAL and DIRECT */
            result = Tcl_EvalEx(ticP->interp, dataP, sz, TCL_EVAL_GLOBAL | TCL_EVAL_DIRECT);
            if (compressed)
                TwapiLzmaFreeBuffer(dataP);
            if (result == TCL_OK)
                Tcl_ResetResult(ticP->interp);
            return result;
        }
    }

    return Twapi_AppendSystemError(ticP->interp, GetLastError());
    
}

/*
 * Loads the initialization script from image file resource
 */
static TCL_RESULT TwapiLoadInitScript(TwapiInterpContext *ticP)
{
    int result;
    gTwapiEmbedType = "embedded";
    result = Twapi_SourceResource(ticP, gTwapiModuleHandle, MODULENAME);
    if (result != TCL_OK) {
        gTwapiEmbedType = "none"; /* Reset */
    }
    return result;
}

int Twapi_GetTwapiBuildInfo(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    Tcl_Obj *objP;
    Tcl_Obj *elemP;

    if (objc != 1)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    /* Return a keyed list */
    
    objP = Tcl_NewListObj(0, NULL);

    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("compiler"));
#if defined(_MSC_VER)
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("vc++"));
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("compiler_version"));
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(_MSC_VER));
#elif defined(__GNUC__)
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("gcc"));
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("compiler_version"));
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewStringObj(__VERSION__, -1));
#else
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("unknown"));
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("compiler_version"));
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("unknown"));
#endif

    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("sdk_version"));
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(VER_PRODUCTBUILD));

    /* Are we building with TEA ? */
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("tea"));
#if defined(HAVE_SYS_TYPES_H)
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(1));
#else
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(0));
#endif

    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("opts"));
    elemP = Tcl_NewListObj(0, NULL);
#ifdef TWAPI_NODESKTOP
    Tcl_ListObjAppendElement(interp, elemP, STRING_LITERAL_OBJ("nodesktop"));
#endif
#ifdef TWAPI_NOSERVER
    Tcl_ListObjAppendElement(interp, elemP, STRING_LITERAL_OBJ("noserver"));
#endif
#ifdef TWAPI_LEAN
    Tcl_ListObjAppendElement(interp, elemP, STRING_LITERAL_OBJ("lean"));
#endif
    Tcl_ListObjAppendElement(interp, objP, elemP);

    /* Tcl 8.4 has no indication of 32/64 builds at Tcl level so we have to. */
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("platform"));
#ifdef _WIN64
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("x64"));
#else
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("x86"));
#endif    
    
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("embed_type"));
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewStringObj(gTwapiEmbedType, -1));

    /* Which Tcl did we build against ? (As opposed to run time) */
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("tcl_header_version"));
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewStringObj(TCL_PATCH_LEVEL, -1));

    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}


static void Twapi_Cleanup(ClientData clientdata)
{
    /* TBD - do we need to protect against more than one call ? */

    // Commented out CoUninitialize for the time being.
    // If there are event sinks in use, and the application exits
    // when the main window is closed, then Twapi_Cleanup gets
    // called BEFORE the window destroy. We then call CoUninitialize
    // and then if subsequently a COM call is made as part of the
    // Tk window destroy binding, we crash.
    // We want to do this last but there seems to be no way to
    // control order of initialization/finalization other than
    // the documented LIFO unloading of packages. That cannot
    // guarantee that Tk will not be loaded before us. So for now
    // we do not call this.
    // Note that Tk destroy binding runs as thread finalization
    // which happens AFTER process finalization (where we get called)
#if 0
    CoUninitialize();
#endif
    // TBD - clean up allocated interp context lists, threads etc.

    DeleteCriticalSection(&gTwapiInterpContextsCS);
    WSACleanup();
}


static TwapiInterpContext *TwapiInterpContextNew(Tcl_Interp *interp, HMODULE hmodule, TwapiInterpContextCleanup *cleaner)
{
    DWORD winerr;
    TwapiInterpContext* ticP = TwapiAlloc(sizeof(*ticP));

    winerr = MemLifoInit(&ticP->memlifo, NULL, NULL, NULL, 8000,
                         MEMLIFO_F_PANIC_ON_FAIL);
    if (winerr != ERROR_SUCCESS) {
        Twapi_AppendSystemError(interp, winerr);
        return NULL;
    }

    ticP->nrefs = 0;
    ticP->interp = interp;
    ticP->thread = Tcl_GetCurrentThread();
    ticP->module.hmod = hmodule;
    ticP->module.cleaner = cleaner;
    ticP->module.data.pval = NULL;

    /* Initialize the critical section used for controlling
     * various attached lists
     *
     * TBD - what's an appropriate spin count? Default of 0 is not desirable
     * As per MSDN, Windows heap manager uses 4000 so we do too.
     */
    InitializeCriticalSectionAndSpinCount(&ticP->lock, 4000);

    ticP->pending_suspended = 0;
    ZLIST_INIT(&ticP->pending);
    ZLIST_INIT(&ticP->threadpool_registrations);

    /* Register a async callback with Tcl. */
    /* TBD - do we really need a separate Tcl_AsyncCreate call per interp?
     * or should it be per process ? Or per thread ? Do we need this at all?
     */
    ticP->async_handler = Tcl_AsyncCreate(Twapi_TclAsyncProc, ticP);

    ticP->notification_win = NULL; /* Created only on demand */
    ticP->device_notification_tid = 0;

    return ticP;
}

TwapiInterpContext *Twapi_AllocateInterpContext(Tcl_Interp *interp, HMODULE hmodule, TwapiInterpContextCleanup *cleaner)
{
    TwapiInterpContext *ticP;

    /*
     * Allocate a context that will be passed around in all commands
     * Different modules may call this for the same interp
     */
    ticP = TwapiInterpContextNew(interp, hmodule, cleaner);
    if (ticP == NULL)
        return NULL;

    /* For all the commands we register with the Tcl interp, we add a single
     * ref for the context, not one per command. This is sufficient since
     * when the interp gets deleted, all the commands get deleted as well.
     * The corresponding Unref happens when the interp is deleted.
     *
     * In addition, we add one more ref because we will place it on the global
     * queue.
     */
    TwapiInterpContextRef(ticP, 1+1);
    EnterCriticalSection(&gTwapiInterpContextsCS);
    ZLIST_PREPEND(&gTwapiInterpContexts, ticP);
    LeaveCriticalSection(&gTwapiInterpContextsCS);

    Tcl_CallWhenDeleted(interp, Twapi_InterpCleanup, ticP);

    return ticP;
}



static void TwapiInterpContextDelete(TwapiInterpContext *ticP)
{
    TWAPI_ASSERT(ticP->interp == NULL);

    /* TBD - does this need to be done only from the Tcl thread ? */
    if (ticP->async_handler)
        Tcl_AsyncDelete(ticP->async_handler);
    ticP->async_handler = 0;    /* Just in case */

    DeleteCriticalSection(&ticP->lock);

    /* TBD - should rest of this be in the Twapi_InterpCleanup instead ? */
    if (ticP->notification_win) {
        DestroyWindow(ticP->notification_win);
        ticP->notification_win = 0;
    }

    // TBD - what about pipes ?

    MemLifoClose(&ticP->memlifo);

    // TBD - what about freeing the memory?
}

/* Decrement ref count and free if 0 */
void TwapiInterpContextUnref(TwapiInterpContext *ticP, int decr)
{
    if (InterlockedExchangeAdd(&ticP->nrefs, -decr) <= decr)
        TwapiInterpContextDelete(ticP);
}

static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp)
{
    TwapiInterpContext *tic2P;
    TwapiThreadPoolRegistration *tprP;

    TWAPI_ASSERT(ticP->interp == interp);

    /* Should this be called from TwapiInterpContextDelete instead ? */
    if (ticP->module.cleaner) {
        ticP->module.cleaner(ticP);
        ticP->module.cleaner = NULL;
    }

    EnterCriticalSection(&gTwapiInterpContextsCS);
    ZLIST_LOCATE(tic2P, &gTwapiInterpContexts, interp, ticP->interp);
    if (tic2P != ticP) {
        LeaveCriticalSection(&gTwapiInterpContextsCS);
        /* Either not found, or linked to a different interp. */
        Tcl_Panic("TWAPI interpreter context not found or attached to the wrong Tcl interpreter.");
    }
    ZLIST_REMOVE(&gTwapiInterpContexts, ticP);
    LeaveCriticalSection(&gTwapiInterpContextsCS);

    ticP->interp = NULL;        /* Must not access during cleanup */
    
    EnterCriticalSection(&ticP->lock);
    ticP->pending_suspended = 1;
    LeaveCriticalSection(&ticP->lock);

    /*
     * Clean up the thread pool registrations. No need to lock but do
     */
    while ((tprP = ZLIST_HEAD(&ticP->threadpool_registrations)) != NULL) {
        /* Note this also unlinks from threadpool_registrations */
        TwapiThreadPoolRegistrationShutdown(tprP);
    }


    /* TBD - call other callback module clean up procedures */
    /* TBD - terminate device notification thread; */
    
    /* Unref for unlinking interp, +1 for removal from gTwapiInterpContexts */
    TwapiInterpContextUnref(ticP, 1+1);
}


/* One time (per process) initialization */
static int TwapiOneTimeInit(Tcl_Interp *interp)
{
    HRESULT hr;
    WSADATA ws_data;
    WORD    ws_ver = MAKEWORD(1,1);

    gTlsIndex = TlsAlloc();
    if (gTlsIndex == TLS_OUT_OF_INDEXES)
        return TCL_ERROR;       /* No point storing error message.
                                   Discarded anyways by Tcl */

    InitializeCriticalSection(&gTwapiInterpContextsCS);
    ZLIST_INIT(&gTwapiInterpContexts);

    if (Tcl_GetVar2Ex(interp, "tcl_platform", "threaded", TCL_GLOBAL_ONLY))
        gTclIsThreaded = 1;
    else
        gTclIsThreaded = 0;

    Tcl_GetVersion(&gTclVersion.major,
                   &gTclVersion.minor,
                   &gTclVersion.patchlevel,
                   &gTclVersion.reltype);

    /*
     * Check if running against an older Tcl version compared to what we 
     * built against.
     */
    if (gTclVersion.major < TCL_MAJOR_VERSION ||
        (gTclVersion.major == TCL_MAJOR_VERSION && gTclVersion.minor < TCL_MINOR_VERSION)) {
        return TCL_ERROR;
    }

    /* Next check is for minimal supported version. Probably not necessary
       given above check but ... */
    if (gTclVersion.major ==  TWAPI_TCL_MAJOR &&
        gTclVersion.minor >= TWAPI_MIN_TCL_MINOR) {
        TwapiInitTclTypes();
        gTwapiOSVersionInfo.dwOSVersionInfoSize =
            sizeof(gTwapiOSVersionInfo);
        if (GetVersionEx(&gTwapiOSVersionInfo)) {
            /* Sockets */
            if (WSAStartup(ws_ver, &ws_data) == 0) {
                /*
                 * Single-threaded COM model - note some Shell extensions
                 * require this if functions such as ShellExecute are
                 * invoked.
                 */
                hr = CoInitializeEx(
                    NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
                if (hr == S_OK || hr == S_FALSE) {
                    /* All init worked. */
                    Tcl_CreateExitHandler(Twapi_Cleanup, NULL);
                    return TCL_OK;
                } else {
                    WSACleanup();
                }
            }
        }
    }

    return TCL_ERROR;
}

int Twapi_AssignTlsSlot()
{
    DWORD slot;
    slot = InterlockedIncrement(&gTlsNextSlot);
    if (slot > TWAPI_TLS_SLOTS) {
        InterlockedDecrement(&gTlsNextSlot); /* So it does not grow unbounded */
        return -1;
    }
    return slot-1;
}

TwapiTls *Twapi_GetTls()
{
    TwapiTls *tlsP;

    tlsP = (TwapiTls *) TlsGetValue(gTlsIndex);
    if (tlsP)
        return tlsP;

    /* TLS for this thread not initialized yet */
    tlsP = (TwapiTls *) TwapiAllocZero(sizeof(*tlsP));
    tlsP->thread = Tcl_GetCurrentThread();
    TlsSetValue(gTlsIndex, tlsP);
    return tlsP;
}

TwapiId Twapi_NewId(TwapiInterpContext *ticP)
{
#ifdef _WIN64
    return InterlockedIncrement64(&gIdGenerator);
#else
    return InterlockedIncrement(&gIdGenerator);
#endif

}

TCL_RESULT Twapi_CheckThreadedTcl(Tcl_Interp *interp)
{
    if (! gTclIsThreaded) {
        if (interp)
            Tcl_SetResult(interp,
                          "This command requires a threaded build of Tcl.",
                          TCL_STATIC);
        return TCL_ERROR;
    }
    return TCL_OK;
}


/* Does basic default initialization of a module */
TwapiInterpContext * Twapi_ModuleInit(Tcl_Interp *interp, const char *nameP, HMODULE hmod, TwapiModuleCallInitializer *initFn, TwapiInterpContextCleanup *cleanerFn)
{
    TwapiInterpContext *ticP;

    /* NOTE: no point setting Tcl_SetResult for errors as they are not
       looked at when DLL is being loaded */

    /* Allocate a context that will be passed around in all interpreters */
    ticP = Twapi_AllocateInterpContext(interp, hmod, cleanerFn);
    if (ticP == NULL)
        return NULL;

    /* Do our own commands. */
    if (initFn && initFn(interp, ticP) != TCL_OK)
        return NULL;

    if (Twapi_SourceResource(ticP, hmod, nameP) != TCL_OK) {
        /* We keep going as scripts might be external, not bound into DLL */
        /* Twapi_AppendLog tbd */
        /* return NULL; */
        Tcl_ResetResult(interp); /* Get rid of any error messages */
    }

    return ticP;
}

