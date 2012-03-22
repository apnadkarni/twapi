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
#define TWAPI_MIN_TCL_MINOR 5

/*
 * Globals
 */
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
static TwapiInterpContext *TwapiInterpContextNew(Tcl_Interp *, HMODULE, TwapiModuleDef * );
static void TwapiInterpContextDelete(TwapiInterpContext *ticP);
static TwapiInterpContext *Twapi_AllocateInterpContext(Tcl_Interp *interp, HMODULE hmodule, TwapiModuleDef *);
static int TwapiOneTimeInit(Tcl_Interp *interp);

HMODULE gTwapiModuleHandle;     /* DLL handle to ourselves */
static TwapiModuleDef gBaseModule = {
    MODULENAME,
    Twapi_InitCalls,
    NULL    /* TBD -  should not be NULL - add a cleanup instead of
               hardcoding in Twapi_InterpCleanup unless we need to control
               to be after all other module cleanups ?
            */
};



#if !defined(TWAPI_STATIC_BUILD)
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gTwapiModuleHandle = hmod;
    return TRUE;
}
#endif

/* Why is this decl needed ? TBD */
const char * __cdecl Tcl_InitStubs(Tcl_Interp *interp, const char *,int);

int TwapiLoadStaticModules(Tcl_Interp *interp)
{
#if defined(TWAPI_SINGLE_MODULE)
/*
 * These files are generated at build time to call twapi module initializers 
 * Note two separate files needed because all prototypes need to come before
 * the calls.
 */
#include "twapi_module_static_proto.h"
#include "twapi_module_static_init.h"
#endif
    /* In case any module left over some crap, clean it out. Note
       errors will not get to this point */
    Tcl_ResetResult(interp);
    return TCL_OK;
}

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

    /* Create the name space and some variables. Not sure if this is explicitly needed */
    Tcl_CreateNamespace(interp, "::twapi", NULL, NULL);
    Tcl_SetVar2(interp, "::twapi::version", MODULENAME, MODULEVERSION, 0);
    Tcl_SetVar2(interp, "::twapi::settings", "log_limit", "100", 0);

    /* Allocate a context that will be passed around in all interpreters */
    ticP = TwapiRegisterModule(interp,  gTwapiModuleHandle, &gBaseModule, 1);
    if (ticP == NULL)
        return TCL_ERROR;

    return TwapiLoadStaticModules(interp);
}

/* Alternate entry point for when the DLL is called twapi */
#ifndef TWAPI_STATIC_BUILD
__declspec(dllexport) 
#endif
int Twapi_Init(Tcl_Interp *interp)
{
    return Twapi_base_Init(interp);
}

/*
 * Loads the initialization script from image file resource
 */
TCL_RESULT Twapi_SourceResource(TwapiInterpContext *ticP, HANDLE dllH, const char *name, int try_file)
{
    HRSRC hres = NULL;
    unsigned char *dataP;
    DWORD sz;
    HGLOBAL hglob;
    int result;
    int compressed;
    Tcl_Interp *interp = ticP->interp;
    Tcl_Obj *pathObj;

    /*
     * Locate the twapi resource and load it if found. First check for
     * compressed type. Then uncompressed.
     */
    compressed = 1;
    hres = FindResourceA(dllH,
                         name,
                         TWAPI_SCRIPT_RESOURCE_TYPE_LZMA);
    if (!hres) {
        hres = FindResourceA(dllH,
                             name,
                             TWAPI_SCRIPT_RESOURCE_TYPE);
        compressed = 0;
    }

    if (hres) {
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
                result = Tcl_EvalEx(interp, dataP, sz, TCL_EVAL_GLOBAL | TCL_EVAL_DIRECT);
                if (compressed)
                    TwapiLzmaFreeBuffer(dataP);
                if (result == TCL_OK)
                    Tcl_ResetResult(interp);
                return result;
            }
        }
        return Twapi_AppendSystemError(interp, GetLastError());
    }

    /* No resource found. Try loading external file from the DLL directory */
    pathObj = TwapiGetInstallDir(ticP, dllH);
    if (pathObj == NULL)
        return TCL_ERROR;
    Tcl_AppendToObj(pathObj, name, -1);
    Tcl_IncrRefCount(pathObj);  /* Must before calling any Tcl_FS functions */
    result = Tcl_FSEvalFile(interp, pathObj);
    Tcl_DecrRefCount(pathObj);

    return result;
}

#ifdef OBSOLETE
/*
 * Loads the initialization script from image file resource
 */
static TCL_RESULT TwapiLoadInitScript(TwapiInterpContext *ticP)
{
    int result;
    result = Twapi_SourceResource(ticP, gTwapiModuleHandle,
                                  WLITERAL(MODULENAME), 1);
    if (result == TCL_OK) {
        gTwapiEmbedType = "embedded";
    }
    return result;
}
#endif

Tcl_Obj *TwapiGetInstallDir(TwapiInterpContext *ticP, HANDLE dllH)
{
    DWORD sz;
    WCHAR path[MAX_PATH+1+1]; /* Extra one char to detect errors */

    if (dllH == NULL)
        dllH = gTwapiModuleHandle;

    /* No resource found. Try loading external file from the DLL directory */
    sz = GetModuleFileNameW(dllH, path, ARRAYSIZE(path));
    if (sz == 0 || sz == ARRAYSIZE(path)) {
        sz = GetLastError();
        if (sz == ERROR_SUCCESS)
            sz = ERROR_INSUFFICIENT_BUFFER;
        if (ticP && ticP->interp)
            Twapi_AppendSystemError(ticP->interp, sz);
        return NULL;
    }

    /* Look for the preceding / or \\ */
    while (sz--) {
        if (path[sz] == L'/' || path[sz] == L'\\') {
            ++sz;
            break;
        }
    }
    path[sz] = 0;
    return Tcl_NewUnicodeObj(path, sz);
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

#if 0
    /* No point to this. Not used and not reliable since in the current
       build, tcl scripts are sourced via a command in the resource script */
    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("embed_type"));
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewStringObj(gTwapiEmbedType, -1));
#endif

    Tcl_ListObjAppendElement(interp, objP, STRING_LITERAL_OBJ("single_module"));
#if defined(TWAPI_SINGLE_MODULE)
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(1));
#else
    Tcl_ListObjAppendElement(interp, objP, Tcl_NewLongObj(0));
#endif

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


static TwapiInterpContext *TwapiInterpContextNew(
    Tcl_Interp *interp, HMODULE hmodule, TwapiModuleDef *modP)
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
    ticP->module.modP = modP;
    ticP->module.data.pval = NULL;

    /* Cache of commonly used objects */
    Tcl_InitHashTable(&ticP->atoms, TCL_STRING_KEYS);

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
#ifdef OBSOLETE
    ticP->device_notification_tid = 0;
#endif
    return ticP;
}

TwapiInterpContext *Twapi_AllocateInterpContext(Tcl_Interp *interp, HMODULE hmodule, TwapiModuleDef *modP)
{
    TwapiInterpContext *ticP;

    /*
     * Allocate a context that will be passed around in all commands
     * Different modules may call this for the same interp
     */
    ticP = TwapiInterpContextNew(interp, hmodule, modP);
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
    Tcl_HashSearch hs;
    Tcl_HashEntry *he;

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

    /* Free up the atoms */
    for (he = Tcl_FirstHashEntry(&ticP->atoms, &hs) ;
         he != NULL;
         he = Tcl_NextHashEntry(&hs)) {
        /* It is safe to delete this element and only this element */
        Tcl_Obj *objP = Tcl_GetHashValue(he);
        Tcl_DeleteHashEntry(he);
        Tcl_DecrRefCount(objP);
    }
    Tcl_DeleteHashTable(&ticP->atoms);

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

/* Note this cleans up only one TwapiInterpContext for interp, not the whole
   interp */
static void Twapi_InterpCleanup(TwapiInterpContext *ticP, Tcl_Interp *interp)
{
    TwapiInterpContext *tic2P;
    TwapiThreadPoolRegistration *tprP;

    TWAPI_ASSERT(ticP->interp == interp);

    /* Should this be called from TwapiInterpContextDelete instead ? */
    if (ticP->module.modP->finalizer) {
        ticP->module.modP->finalizer(ticP);
        ticP->module.modP->finalizer = NULL;
    }

    EnterCriticalSection(&gTwapiInterpContextsCS);
    /* CMP should return 0 on a match */
#define CMP(x, y) ((x) != (y))
    ZLIST_FIND(tic2P, &gTwapiInterpContexts, CMP, ticP);
#undef CMP
    if (tic2P == NULL) {
        LeaveCriticalSection(&gTwapiInterpContextsCS);
        /* Either not found, or linked to a different interp. */
        Tcl_Panic("TWAPI interpreter context not found attached to a Tcl interpreter.");
    }
    ZLIST_REMOVE(&gTwapiInterpContexts, ticP);
    ticP->interp = NULL;        /* Must not access now on */
    LeaveCriticalSection(&gTwapiInterpContextsCS);
    
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

TwapiInterpContext *TwapiGetBaseContext(Tcl_Interp *interp)
{
    TwapiInterpContext *ticP;

    EnterCriticalSection(&gTwapiInterpContextsCS);
    /* CMP should return 0 on a match */
#define CMP(lelem, ip) ((lelem)->interp != (ip) || ! STREQ((lelem)->module.modP->name, "twapi_base"))
    ZLIST_FIND(ticP, &gTwapiInterpContexts, CMP, interp);
#undef CMP
    LeaveCriticalSection(&gTwapiInterpContextsCS);
    return ticP;
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
                 * TBD - should this be in twapi_com module ?
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
TwapiInterpContext *TwapiRegisterModule(
    Tcl_Interp *interp,
    HMODULE hmod,
    TwapiModuleDef *modP, /* MUST BE STATIC/NEVER DEALLOCATED */
    int context_type
    )
{
    TwapiInterpContext *ticP;

    if (modP->finalizer && ! context_type) {
        Tcl_SetResult(interp, "Non-private context cannot be requested if finalalizer is specified", TCL_STATIC);
    }

    if (context_type != DEFAULT_TIC) {
        ticP = Twapi_AllocateInterpContext(interp, hmod, modP);
        if (ticP == NULL)
            return NULL;
    } else {
        
        ticP = TwapiGetBaseContext(interp);

        /* We do not need to increment the ticP ref count because
         * we can rely on the base module ref count which will go
         * away only when the whole interp is deleted at which point
         * the calling module will not be invoked any more. Of course
         * if the calling module might access it AFTER the interp is
         * gone, it has to do an ticP ref increment itself just like
         * for a private ticP
         */
    }

    TWAPI_ASSERT(ticP);

    if (modP->initializer && modP->initializer(interp, ticP) != TCL_OK ||
        Twapi_SourceResource(ticP, hmod, modP->name, 1) != TCL_OK ||
        Tcl_PkgProvide(interp, modP->name, MODULEVERSION) != TCL_OK
        ) {
        if (context_type)
            TwapiInterpContextUnref(ticP, 1);
        return NULL;
    }

    return ticP;
}

int Twapi_GetVersionEx(Tcl_Interp *interp)
{
    OSVERSIONINFOEXW vi;
    Tcl_Obj *objP;

    vi.dwOSVersionInfoSize = sizeof(vi);
    if (GetVersionExW((OSVERSIONINFOW *)&vi) == 0) {
        return TwapiReturnSystemError(interp);
    }

    objP = Tcl_NewListObj(0, NULL);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwOSVersionInfoSize);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwMajorVersion);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwMinorVersion);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwBuildNumber);
    Twapi_APPEND_DWORD_FIELD_TO_LIST(interp, objP, &vi, dwPlatformId);
    //TCHAR szCSDVersion[128];
    Twapi_APPEND_WORD_FIELD_TO_LIST(interp, objP, &vi,  wServicePackMajor);
    Twapi_APPEND_WORD_FIELD_TO_LIST(interp, objP, &vi,  wServicePackMinor);
    Twapi_APPEND_WORD_FIELD_TO_LIST(interp, objP, &vi,  wSuiteMask);
    Twapi_APPEND_WORD_FIELD_TO_LIST(interp, objP, &vi,  wProductType);
    Twapi_APPEND_WORD_FIELD_TO_LIST(interp, objP, &vi,  wReserved);

    Tcl_SetObjResult(interp, objP);
    return TCL_OK;
}

int Twapi_WTSEnumerateProcesses(Tcl_Interp *interp, HANDLE wtsH)
{
    WTS_PROCESS_INFOW *processP = NULL;
    DWORD count;
    DWORD i;
    Tcl_Obj *records = NULL;
    Tcl_Obj *fields = NULL;
    Tcl_Obj *objv[4];

    /* Note wtsH == NULL means current server */
    if (! (BOOL) (WTSEnumerateProcessesW)(wtsH, 0, 1, &processP, &count)) {
        return Twapi_AppendSystemError(interp, GetLastError());
    }


    /* Now create the data records */
    records = Tcl_NewListObj(0, NULL);
    for (i = 0; i < count; ++i) {
        Tcl_Obj *sidObj;

        /* Create a list corresponding to the fields for the process entry */
        objv[0] = Tcl_NewLongObj(processP[i].SessionId);
        objv[1] = Tcl_NewLongObj(processP[i].ProcessId);
        objv[2] = ObjFromUnicode(processP[i].pProcessName);
        if (processP[i].pUserSid) {
            if (ObjFromSID(interp, processP[i].pUserSid, &sidObj) != TCL_OK) {
                Twapi_WTSFreeMemory(processP);
                Twapi_FreeNewTclObj(records);
                return TCL_ERROR;
            }
            objv[3] = sidObj;
        }
        else {
            /* NULL SID pointer. */
            objv[3] = STRING_LITERAL_OBJ("");
        }

        Tcl_ListObjAppendElement(interp, records, Tcl_NewLongObj(processP[i].ProcessId));
        Tcl_ListObjAppendElement(interp, records, Tcl_NewListObj(4, objv));
    }

    Twapi_WTSFreeMemory(processP);

    /* Make up the field name descriptor */
    objv[0] = STRING_LITERAL_OBJ("SessionId");
    objv[1] = STRING_LITERAL_OBJ("ProcessId");
    objv[2] = STRING_LITERAL_OBJ("pProcessName");
    objv[3] = STRING_LITERAL_OBJ("pUserSid");
    fields = Tcl_NewListObj(4, objv);

    /* Put field names and records to make up the recordarray */
    objv[0] = fields;
    objv[1] = records;
    Tcl_SetObjResult(interp, Tcl_NewListObj(2, objv));
    return TCL_OK;
}


/*
 * Returns the Tcl_Obj corresponding to the given string.
 * Caller MUST NOT call Tcl_DecrRefCount on the object without
 * a prior Tcl_IncrRefCount. Moreover, if it wants to hang on to it
 * it must do a Tcl_IncrRefCount itself directly, or implicitly via
 * a call such as Tcl_ListObjAppendElement.
 * (This is similar to Tcl_ListObjIndex)
 */
Tcl_Obj *TwapiGetAtom(TwapiInterpContext *ticP, const char *key)
{
    Tcl_HashEntry *he;
    int new_entry;
    
    he = Tcl_CreateHashEntry(&ticP->atoms, key, &new_entry);
    if (new_entry) {
        Tcl_Obj *objP = Tcl_NewStringObj(key, -1);
        Tcl_IncrRefCount(objP);
        Tcl_SetHashValue(he, objP);
        return objP;
    } else {
        return (Tcl_Obj *) Tcl_GetHashValue(he);
    }
}

Tcl_Obj *Twapi_GetAtomStats(TwapiInterpContext *ticP) 
{
    char *stats = Tcl_HashStats(&ticP->atoms);
    Tcl_Obj *objP = Tcl_NewStringObj(stats, -1);
    ckfree(stats);
    return objP;
}
