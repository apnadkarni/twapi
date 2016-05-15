/*
 * Copyright (c) 2010-2014, Ashok P. Nadkarni
 * All rights reserved.
 *
 * See the file LICENSE for license
 */

#include "twapi.h"
#include "twapi_base.h"
#include <ntverp.h>             /* Needed for VER_PRODUCTBUILD SDK version */
#include "tclTomMath.h"

#define TWAPI_TCL_MAJOR 8
#define TWAPI_MIN_TCL_MINOR 5

/* Following two definitions required for MinGW builds */
#ifndef MODULENAME
#define MODULENAME "twapi_base"
#endif

#ifndef MODULEVERSION
#define MODULEVERSION PACKAGE_VERSION
#endif

/*
 * Struct to keep track of registered pointers.
 *
 * Twapi keeps track of pointers passed to the script level to lessen the
 * probability of double frees. At the same time, some Win32 API's return
 * the same pointer multiple times (e.g. Cert* APIs) so they need to be
 * ref counted as well in some cases.
 *
 * There are two sets of API's - ref counted and non-refcounted. For
 * any pointer only one of the two must be used else error/panic will result
 *  Twapi*Pointer* - non-refcounted API
 *  Twapi*CountedPointer* - refcounted API
 *
 * The tag is to verify that the pointer is of the appropriate kind. Usually
 * the address of a free routine is used as the tag.
 */
typedef struct _TwapiRegisteredPointer {
    void *tag;                  /* Type tag */
    int   nrefs;                /* Reference count. -1 for non-refcounted */
} TwapiRegisteredPointer;

/*
 * Globals
 */
OSVERSIONINFOW gTwapiOSVersionInfo;
TwapiBaseSettings gBaseSettings = {
    1,                          /* use_unicode_obj, controlled via Tcl_LinkVar */
};
GUID gTwapiNullGuid;             /* Initialized to all zeroes */
struct TwapiTclVersion gTclVersion;
static int gTclIsThreaded;
static DWORD gTlsIndex = TLS_OUT_OF_INDEXES; /* As returned by TlsAlloc */
static LONG volatile gTlsNextSlot;  /* Index into private slots in Tls area. */

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

static void TwapiBaseModuleCleanup(TwapiInterpContext *ticP);
static void Twapi_Cleanup(ClientData clientdata);
static void Twapi_InterpCleanup(ClientData clientdata, Tcl_Interp *interp);
static void Twapi_InterpContextCleanup(void*, Tcl_Interp *interp);
static TwapiInterpContext *TwapiInterpContextNew(Tcl_Interp *, HMODULE, TwapiModuleDef * );
static void TwapiInterpContextDelete(TwapiInterpContext *ticP);
static TwapiInterpContext *Twapi_AllocateInterpContext(Tcl_Interp *interp, HMODULE hmodule, TwapiModuleDef *);
static int TwapiOneTimeInit(void *);

HMODULE gTwapiModuleHandle;     /* DLL handle to ourselves */
static TwapiModuleDef gBaseModule = {
    MODULENAME,
    Twapi_InitCalls,
    TwapiBaseModuleCleanup,
};



#if !defined(TWAPI_STATIC_BUILD)
BOOL WINAPI DllMain(HINSTANCE hmod, DWORD reason, PVOID unused)
{
    if (reason == DLL_PROCESS_ATTACH)
        gTwapiModuleHandle = hmod;
    return TRUE;
}
#endif

static void TwapiTlsUnref()
{
    if (gTlsIndex != TLS_OUT_OF_INDEXES) {
        TwapiTls *tlsP = TlsGetValue(gTlsIndex);
        if (tlsP) {
            TWAPI_ASSERT(tlsP->nrefs > 0);
            tlsP->nrefs -= 1;
            if (tlsP->nrefs == 0) {
                MemLifoClose(&tlsP->memlifo);
                ObjDecrRefs(tlsP->ffiObj);
                TwapiFree(tlsP);
                TlsSetValue(gTlsIndex, NULL);
            }
        }
    }
}

static TCL_RESULT TwapiTlsInit()
{
    TwapiTls *tlsP;

    TWAPI_ASSERT(gTlsIndex != TLS_OUT_OF_INDEXES);
    tlsP = TlsGetValue(gTlsIndex);
    if (tlsP == NULL) {
        tlsP = (TwapiTls *) TwapiAllocZero(sizeof(*tlsP));
        if (! TlsSetValue(gTlsIndex, tlsP)) {
            TwapiFree(tlsP);
            return TCL_ERROR;
        }
        tlsP->thread = Tcl_GetCurrentThread();
        /* TBD - should we raise alloc from 8000 ? Too small ? */
        if (MemLifoInit(&tlsP->memlifo, NULL, NULL, NULL, 8000,
                             MEMLIFO_F_PANIC_ON_FAIL) != ERROR_SUCCESS) {
            TwapiFree(tlsP);
            return TCL_ERROR;
        }
        tlsP->ffiObj = ObjNewDict();
        ObjIncrRefs(tlsP->ffiObj);
    }

    tlsP->nrefs += 1;
    return TCL_OK;
}

int Twapi_AssignTlsSubSlot()
{
    DWORD slot;
    slot = InterlockedIncrement(&gTlsNextSlot);
    if (slot > TWAPI_TLS_SLOTS) {
        InterlockedDecrement(&gTlsNextSlot); /* So it does not grow unbounded */
        return -1;
    }
    return slot-1;
}

/* TBD - make inline */
TwapiTls *Twapi_GetTls()
{
    TwapiTls *tlsP;

    tlsP = (TwapiTls *) TlsGetValue(gTlsIndex);
    if (tlsP == NULL)
        Tcl_Panic("TLS pointer is NULL");
    return tlsP;
}

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
    TwapiInterpContext *ticP;
    HRESULT hr;

    /* IMPORTANT */
    /* MUST BE FIRST CALL as it initializes Tcl stubs - should this be the
       done for EVERY interp creation or move into one-time above ? TBD
     */
    /* TBD dgp says this #ifdef USE_TCL_STUBS is not needed and indeed
       that seems to be the case for Tcl_InitStubs. But Tcl_TomMath_InitStubs
       crashes on a static build not using stubs
    */
#ifdef USE_TCL_STUBS
    if (Tcl_InitStubs(interp, TCL_VERSION, 0) == NULL) {
        return TCL_ERROR;
    }
    if (Tcl_TomMath_InitStubs(interp, 0) == NULL) {
        return TCL_ERROR;
    }
#endif

    /* Init unless already done. */
    if (! TwapiDoOneTimeInit(&gTwapiInitialized, TwapiOneTimeInit, interp))
        return TCL_ERROR;

    /* NOTE: no point setting Tcl_SetResult for errors as they are not
       looked at when DLL is being loaded */

    /*
     * Per interp initialization
     */

    if (TwapiTlsInit() != TCL_OK)
        return TCL_ERROR;

    /*
     * Single-threaded COM model - note some Shell extensions
     * require this if functions such as ShellExecute are
     * invoked. TBD - should we do this lazily in com and mstask modules ?
     *
     * TBD - recent MSDN docs states:
     * "Avoid the COM single-threaded apartment model, as it is incompatible
     * with the thread pool. STA creates thread state which can affect the
     * next work item for the thread. STA is generally long-lived and has
     * thread affinity, which is the opposite of the thread pool."
     * Since we use thread pools, does this mean we should not be
     * using STA? Or does that only apply when making COM calls from
     * a thread pool thread in which case it would not apply to us?
     */
    hr = CoInitializeEx(
        NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (hr != S_OK && hr != S_FALSE)
        return TCL_ERROR;

    /* Create the name space and some variables. Not sure if this is explicitly needed */
    Tcl_CreateNamespace(interp, "::twapi", NULL, NULL);
    Tcl_SetVar2(interp, "::twapi::version", MODULENAME, MODULEVERSION, 0);
    Tcl_SetVar2(interp, "::twapi::settings", "log_limit", "100", 0);
    Tcl_LinkVar(interp, "::twapi::settings(use_unicode_obj)", (char *)&gBaseSettings.use_unicode_obj, TCL_LINK_ULONG);

    /* Allocate a context that will be passed around in all interpreters */
    ticP = TwapiRegisterModule(interp,  gTwapiModuleHandle, &gBaseModule, NEW_TIC);
    if (ticP == NULL)
        return TCL_ERROR;

    ticP->module.data.pval = TwapiAlloc(sizeof(TwapiBaseSpecificContext));
    /* Cache of commonly used objects */
    Tcl_InitHashTable(&BASE_CONTEXT(ticP)->atoms, TCL_STRING_KEYS);
    /* Pointer registration table */
    Tcl_InitHashTable(&BASE_CONTEXT(ticP)->pointers, TCL_ONE_WORD_KEYS);
    /* Trap stack */
    BASE_CONTEXT(ticP)->trapstack = ObjNewList(0, NULL);
    ObjIncrRefs(BASE_CONTEXT(ticP)->trapstack);

    Tcl_CallWhenDeleted(interp, Twapi_InterpCleanup, NULL);

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
TCL_RESULT Twapi_SourceResource(Tcl_Interp *interp, HANDLE dllH, const char *name, int try_file)
{
    HRSRC hres = NULL;
    unsigned char *dataP;
    DWORD sz;
    HGLOBAL hglob;
    int result;
    int compressed;
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
                    dataP = TwapiLzmaUncompressBuffer(interp, dataP, sz, &sz);
                    if (dataP == NULL)
                        return TCL_ERROR; /* interp already has error */
                }
                
                /* The resource is expected to be UTF-8 (actually strict ASCII) */
                /* TBD - double check use of GLOBAL and DIRECT */
                result = Tcl_EvalEx(interp, (char *)dataP, sz, TCL_EVAL_GLOBAL | TCL_EVAL_DIRECT);
                if (compressed)
                    TwapiLzmaFreeBuffer(dataP);
                if (result == TCL_OK)
                    Tcl_ResetResult(interp);
                return result;
            }
        }
        return Twapi_AppendSystemError(interp, GetLastError());
    }

    if (!try_file) {
        Tcl_AppendResult(interp, "Resource ", name,  " not found.", NULL);
        return TCL_ERROR;
    }    

    /* No resource found. Try loading external file from the DLL directory */
    pathObj = TwapiGetInstallDir(interp, dllH);
    if (pathObj == NULL)
        return TCL_ERROR;
    Tcl_AppendToObj(pathObj, name, -1);
    ObjIncrRefs(pathObj);  /* Must before calling any Tcl_FS functions */
    result = Tcl_FSEvalFile(interp, pathObj);
    ObjDecrRefs(pathObj);

    return result;
}

Tcl_Obj *TwapiGetInstallDir(Tcl_Interp *interp, HANDLE dllH)
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
        if (interp)
            Twapi_AppendSystemError(interp, sz);
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
    return ObjFromWinCharsN(path, sz);
}

int Twapi_GetTwapiBuildInfo(
    ClientData dummy,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[]
)
{
    Tcl_Obj *objs[16];

    if (objc != 1)
        return TwapiReturnError(interp, TWAPI_BAD_ARG_COUNT);

    /* Return a keyed list */
    
    objs[0] = STRING_LITERAL_OBJ("compiler");
#if defined(_MSC_VER)
    objs[1] = STRING_LITERAL_OBJ("vc++");
    objs[2] = STRING_LITERAL_OBJ("compiler_version");
    objs[3] = ObjFromLong(_MSC_VER);
#elif defined(__GNUC__)
    objs[1] = STRING_LITERAL_OBJ("gcc");
    objs[2] = STRING_LITERAL_OBJ("compiler_version");
    objs[3] = ObjFromString(__VERSION__);
#else
    objs[1] = STRING_LITERAL_OBJ("unknown");
    objs[2] = STRING_LITERAL_OBJ("compiler_version");
    objs[3] = STRING_LITERAL_OBJ("unknown");
#endif

    objs[4] = STRING_LITERAL_OBJ("sdk_version");
    objs[5] = ObjFromLong(VER_PRODUCTBUILD);

    /* Are we building with TEA ? */
    objs[6] = STRING_LITERAL_OBJ("tea");
#if defined(HAVE_SYS_TYPES_H)
    objs[7] = ObjFromLong(1);
#else
    objs[7] = ObjFromLong(0);
#endif

    objs[8] = STRING_LITERAL_OBJ("opts");

    objs[9] = ObjEmptyList();
#ifdef NOOPTIMIZE
    ObjAppendElement(NULL, objs[9], STRING_LITERAL_OBJ("nooptimize"));
#endif
#ifdef TWAPI_ENABLE_LOG
    ObjAppendElement(NULL, objs[9], STRING_LITERAL_OBJ("enable_log"));
#endif
#ifdef TWAPI_ENABLE_ASSERT
    ObjAppendElement(NULL, objs[9], STRING_LITERAL_OBJ("enable_assert"));
#endif

    objs[10] = STRING_LITERAL_OBJ("single_module");
#if defined(TWAPI_SINGLE_MODULE)
    objs[11] = ObjFromLong(1);
#else
    objs[11] = ObjFromLong(0);
#endif

    /* Which Tcl did we build against ? (As opposed to run time) */
    objs[12] = STRING_LITERAL_OBJ("tcl_header_version");
    objs[13] = ObjFromString(TCL_PATCH_LEVEL);

    objs[14] = STRING_LITERAL_OBJ("source_id");
#ifdef HGID
    objs[15] = STRING_LITERAL_OBJ(HGID);
#else
    objs[15] = Tcl_NewObj();
#endif

    return ObjSetResult(interp, ObjNewList(16, objs));
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

    /* TBD - may be call Tcl_CreateThreadExitHandler in Twapi_Init to
       CoUninitialize on thread exit since thread exit handlers are
       called AFTER process exit handlers. See which other calls 
       from below should go there.
    */

#if 0
    CoUninitialize();
#endif
    // TBD - clean up allocated interp context lists, threads etc.

    DeleteCriticalSection(&gTwapiInterpContextsCS);
    WSACleanup();
}

static void Twapi_InterpCleanup(ClientData unused, Tcl_Interp *interp)
{
    TwapiTlsUnref();            /* Matches one in Twapi_Init */
}


static TwapiInterpContext *TwapiInterpContextNew(
    Tcl_Interp *interp, HMODULE hmodule, TwapiModuleDef *modP)
{
    TwapiInterpContext* ticP = TwapiAlloc(sizeof(*ticP));

    ticP->nrefs = 0;
    ticP->interp = interp;
    ticP->thread = Tcl_GetCurrentThread();
    ticP->module.hmod = hmodule;
    ticP->module.modP = modP;
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

    ticP->notification_win = NULL; /* Created only on demand */

    return ticP;
}

TwapiInterpContext *Twapi_AllocateInterpContext(Tcl_Interp *interp, HMODULE hmodule, TwapiModuleDef *modP)
{
    TwapiInterpContext *ticP;
    TwapiTls *tlsP;

    /*
     * Allocate a context that will be passed around in all commands
     * Different modules may call this for the same interp
     */
    ticP = TwapiInterpContextNew(interp, hmodule, modP);
    if (ticP == NULL)
        return NULL;

    /* 
     * Cache a pointer to the TLS memlifo. Note there are assumptions
     * in the code that the ticP->memlifoP is the same as the TLS memlifo
     * whereby the SWS* operations and ticP->memlifo operations 
     * impact the same memlifo.
     */
    tlsP = Twapi_GetTls();
    tlsP->nrefs += 1;           /* Will be unrefed when ticP is deleted */
    ticP->memlifoP = &tlsP->memlifo;

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
    /*
     * We have to search for base module more often than others so 
     * prepend that and append others. Note with multiple interps there
     * can be more than on base module entry in the list.
     */
    if (hmodule == gTwapiModuleHandle) {
        ZLIST_PREPEND(&gTwapiInterpContexts, ticP);
    } else {
        ZLIST_APPEND(&gTwapiInterpContexts, ticP);
    }
    LeaveCriticalSection(&gTwapiInterpContextsCS);

    Tcl_CallWhenDeleted(interp, Twapi_InterpContextCleanup, (void *) ticP);

    return ticP;
}

/* Note CALLER MAY BE SOME OTHER THREAD, NOT NECESSARILY THE INTERP ONE */
/* Most cleanup should have happened via Twapi_InterpContextCleanup */
static void TwapiInterpContextDelete(TwapiInterpContext *ticP)
{
    TWAPI_ASSERT(ticP->interp == NULL);

    DeleteCriticalSection(&ticP->lock);

    /* TBD - should rest of this be in the Twapi_InterpContextCleanup instead ? */
    if (ticP->notification_win) {
        DestroyWindow(ticP->notification_win);
        ticP->notification_win = 0;
    }

    TwapiTlsUnref();

    // TBD - what about freeing the memory?
}

/* Decrement ref count and free if 0 */
/* Note CALLER MAY BE SOME OTHER THREAD, NOT NECESSARILY THE INTERP ONE */
void TwapiInterpContextUnref(TwapiInterpContext *ticP, int decr)
{
    if (InterlockedExchangeAdd(&ticP->nrefs, -decr) <= decr)
        TwapiInterpContextDelete(ticP);
}

/* Note this cleans up only one TwapiInterpContext for interp, not the whole
   interp */
static void Twapi_InterpContextCleanup(void *pv, Tcl_Interp *interp)
{
    TwapiInterpContext *ticP = (TwapiInterpContext *)pv;
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


/* One time (per process) initialization for base module */
static int TwapiOneTimeInit(void *pv)
{
    Tcl_Interp *interp = (Tcl_Interp *) pv;
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
        if (TwapiRtlGetVersion(&gTwapiOSVersionInfo)) {
            /* Sockets */
            if (WSAStartup(ws_ver, &ws_data) == 0) {
                Tcl_CreateExitHandler(Twapi_Cleanup, NULL);
                return TCL_OK;
            }
        }
    }

    return TCL_ERROR;
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
            ObjSetStaticResult(interp, "Tcl build is not threaded.");
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
    char buf[100];
    Tcl_Obj *objP;

    if (modP->finalizer && ! context_type) {
        /* Non-private context cannot be requested if finalizer is specified */
        ObjSetStaticResult(interp, "Finalizer mandates private context");
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

    /* TBD - may be the last param to SourceResource should be 0? */
    if ((modP->initializer && modP->initializer(interp, ticP) != TCL_OK) ||
        Twapi_SourceResource(interp, hmod, modP->name, 1) != TCL_OK ||
        Tcl_PkgProvide(interp, modP->name, MODULEVERSION) != TCL_OK
        ) {
        if (context_type)
            TwapiInterpContextUnref(ticP, 1);
        return NULL;
    }

    /* Link the trace flags. Note these are shared amongs interps */
    _snprintf(buf, ARRAYSIZE(buf), "%s(%s)", TWAPI_LOG_CONFIG_VAR, modP->name);
    objP = Tcl_GetVar2Ex(interp, buf, NULL, TCL_GLOBAL_ONLY);
    if (objP) {
        /*
         * Variable already exists, copy its value to trace flags.
         * Errors are ignored as flags will remain 0
         */
        ObjToDWORD(NULL, objP, &modP->log_flags);
    }
    Tcl_LinkVar(interp, buf, (char *) &modP->log_flags, TCL_LINK_ULONG);
    return ticP;
}

int Twapi_GetVersionEx(Tcl_Interp *interp)
{
    OSVERSIONINFOEXW vi;
    Tcl_Obj *objP;

    vi.dwOSVersionInfoSize = sizeof(vi);
    if (TwapiRtlGetVersion((OSVERSIONINFOW *)&vi) == 0) {
        return TwapiReturnSystemError(interp);
    }

    objP = ObjNewList(0, NULL);
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

    return ObjSetResult(interp, objP);
}

int Twapi_WTSEnumerateProcesses(Tcl_Interp *interp, HANDLE wtsH)
{
    WTS_PROCESS_INFOW *processP = NULL;
    DWORD count;
    DWORD i;
    Tcl_Obj **records;
    Tcl_Obj *fields = NULL;
    Tcl_Obj *objv[4];

    /* Note wtsH == NULL means current server */
    if (! (BOOL) (WTSEnumerateProcessesW)(wtsH, 0, 1, &processP, &count)) {
        return Twapi_AppendSystemError(interp, GetLastError());
    }

    records = SWSPushFrame(count * sizeof(Tcl_Obj*), NULL);
    for (i = 0; i < count; ++i) {
        Tcl_Obj *sidObj;

        /* Create a list corresponding to the fields for the process entry */
        objv[0] = ObjFromLong(processP[i].SessionId);
        objv[1] = ObjFromLong(processP[i].ProcessId);
        objv[2] = ObjFromWinChars(processP[i].pProcessName);
        if (processP[i].pUserSid) {
            if (ObjFromSID(interp, processP[i].pUserSid, &sidObj) != TCL_OK) {
                Twapi_WTSFreeMemory(processP);
                ObjDecrArrayRefs(i, records);
                SWSPopFrame();
                return TCL_ERROR;
            }
            objv[3] = sidObj;
        }
        else {
            /* NULL SID pointer. */
            objv[3] = STRING_LITERAL_OBJ("");
        }
        records[i] = ObjNewList(4, objv);
    }

    Twapi_WTSFreeMemory(processP);

    /* Make up the field name descriptor */
    objv[0] = STRING_LITERAL_OBJ("SessionId");
    objv[1] = STRING_LITERAL_OBJ("ProcessId");
    objv[2] = STRING_LITERAL_OBJ("pProcessName");
    objv[3] = STRING_LITERAL_OBJ("pUserSid");
    fields = ObjNewList(4, objv);

    /* Put field names and records to make up the recordarray */
    objv[0] = fields;
    objv[1] = ObjNewList(count, records);
    SWSPopFrame();
    return ObjSetResult(interp, ObjNewList(2, objv));
}


/*
 * Returns the Tcl_Obj corresponding to the given string.
 * Caller MUST NOT call ObjDecrRefs on the object without
 * a prior ObjIncrRefs. Moreover, if it wants to hang on to it
 * it must do a ObjIncrRefs itself directly, or implicitly via
 * a call such as ObjAppendElement.
 * (This is similar to ObjListIndex)
 */
Tcl_Obj *TwapiGetAtom(TwapiInterpContext *ticP, const char *key)
{
    Tcl_HashEntry *he;
    int new_entry;
    
    if (ticP->module.hmod != gTwapiModuleHandle)
        ticP = TwapiGetBaseContext(ticP->interp);

    he = Tcl_CreateHashEntry(&BASE_CONTEXT(ticP)->atoms, key, &new_entry);
    if (new_entry) {
        Tcl_Obj *objP = ObjFromString(key);
        ObjIncrRefs(objP);
        Tcl_SetHashValue(he, objP);
        return objP;
    } else {
        return (Tcl_Obj *) Tcl_GetHashValue(he);
    }
}

void TwapiPurgeAtoms(TwapiInterpContext *ticP)
{
    Tcl_HashEntry *he;
    Tcl_HashSearch hs;
    
    if (ticP->module.hmod != gTwapiModuleHandle)
        ticP = TwapiGetBaseContext(ticP->interp);

    if (BASE_CONTEXT(ticP)) {
        for (he = Tcl_FirstHashEntry(&(BASE_CONTEXT(ticP)->atoms), &hs) ;
             he != NULL;
             he = Tcl_NextHashEntry(&hs)) {
            /* It is safe to delete this and only this hash element */
            Tcl_Obj *objP = Tcl_GetHashValue(he);
            /* The expectation is that when this routine is called,
               the caller is done with its use of atoms and released
               its use of them. If any other component is using the
               atom, ref count will be at least 2 (since the atom
               table itself contributes 1). If this is not the case
               remove from the atom table
            */
            if (! Tcl_IsShared(objP)) {
                Tcl_DeleteHashEntry(he);
                ObjDecrRefs(objP);
            }
        }
    }
}

#if TWAPI_ENABLE_INSTRUMENTATION
Tcl_Obj *Twapi_GetAtoms(TwapiInterpContext *ticP)
{
    Tcl_HashEntry *he;
    Tcl_HashSearch hs;
    Tcl_Obj *atomsObj;
    
    if (ticP->module.hmod != gTwapiModuleHandle)
        ticP = TwapiGetBaseContext(ticP->interp);

    atomsObj = ObjNewList(0, NULL);
    if (BASE_CONTEXT(ticP)) {
        for (he = Tcl_FirstHashEntry(&(BASE_CONTEXT(ticP)->atoms), &hs) ;
             he != NULL;
             he = Tcl_NextHashEntry(&hs)) {
            /* It is safe to delete this and only this hash element */
            Tcl_Obj *objP = Tcl_GetHashValue(he);
            ObjAppendElement(NULL, atomsObj, objP);
            ObjAppendElement(NULL, atomsObj, ObjFromLong(objP->refCount));
        }
    }
    return atomsObj;
}
#endif

#if TWAPI_ENABLE_INSTRUMENTATION
Tcl_Obj *Twapi_GetAtomStats(TwapiInterpContext *ticP) 
{
    char *stats;
    Tcl_Obj *objP;

    if (ticP->module.hmod != gTwapiModuleHandle)
        ticP = TwapiGetBaseContext(ticP->interp);

    stats = Tcl_HashStats(&BASE_CONTEXT(ticP)->atoms);
    objP = ObjFromString(stats);
    ckfree(stats);
    return objP;
}
#endif

static void TwapiBaseModuleCleanup(TwapiInterpContext *ticP)
{
    Tcl_HashSearch hs;
    Tcl_HashEntry *he;

    if (BASE_CONTEXT(ticP)) {
        for (he = Tcl_FirstHashEntry(&(BASE_CONTEXT(ticP)->atoms), &hs) ;
             he != NULL;
             he = Tcl_NextHashEntry(&hs)) {
            /* It is safe to delete this and only this hash element */
            Tcl_Obj *objP = Tcl_GetHashValue(he);
            Tcl_DeleteHashEntry(he);
            ObjDecrRefs(objP);
        }
        Tcl_DeleteHashTable(&(BASE_CONTEXT(ticP)->atoms));

        for (he = Tcl_FirstHashEntry(&(BASE_CONTEXT(ticP)->pointers), &hs) ;
             he != NULL;
             he = Tcl_NextHashEntry(&hs)) {
            /* It is safe to delete this and only this hash element */
            TwapiRegisteredPointer *rP = Tcl_GetHashValue(he);
            Tcl_DeleteHashEntry(he);
            ckfree((char *)rP);
        }
        Tcl_DeleteHashTable(&(BASE_CONTEXT(ticP)->pointers));
    }
}

int TwapiVerifyPointerTic(TwapiInterpContext *ticP, const void *p, void *typetag)
{
    Tcl_HashEntry *he;

    TWAPI_ASSERT(BASE_CONTEXT(ticP));

    he = Tcl_FindHashEntry(&BASE_CONTEXT(ticP)->pointers, p);
    if (he) {
        /* There are some corner cases where caller sets typetag to NULL
           to indicate only that pointer of *some* tag is registered. 
           So do not check tags in that case
        */
        if (typetag) {
            TwapiRegisteredPointer *rP = Tcl_GetHashValue(he);
            if (rP->tag && rP->tag != typetag)
                return TWAPI_REGISTERED_POINTER_TAG_MISMATCH;
        }
        return TWAPI_NO_ERROR;
    }

    return p ? TWAPI_REGISTERED_POINTER_NOTFOUND : TWAPI_NULL_POINTER;
}

int TwapiVerifyPointer(Tcl_Interp *interp, const void *p, void *typetag)
{
    TwapiInterpContext *ticP = TwapiGetBaseContext(interp);
    TWAPI_ASSERT(ticP);
    return TwapiVerifyPointerTic(ticP, p, typetag);
}

TCL_RESULT TwapiRegisterPointerTic(TwapiInterpContext *ticP, const void *p, void *typetag)
{
    Tcl_HashEntry *he;
    int new_entry;

    TWAPI_ASSERT(BASE_CONTEXT(ticP));

    if (p == NULL)
        return TwapiReturnError(ticP->interp, TWAPI_NULL_POINTER);

    he = Tcl_CreateHashEntry(&BASE_CONTEXT(ticP)->pointers, p, &new_entry);
    if (he && new_entry) {
        TwapiRegisteredPointer *rP = (TwapiRegisteredPointer *) ckalloc(sizeof(*rP));
        rP->tag = typetag;
        rP->nrefs = -1;         /* non-refcounted pointer */
        Tcl_SetHashValue(he, rP);
        return TCL_OK;
    } else {
        return TwapiReturnError(ticP->interp, TWAPI_REGISTERED_POINTER_EXISTS);
    }
}

TCL_RESULT TwapiRegisterPointer(Tcl_Interp *interp, const void *p, void *typetag)
{
    TwapiInterpContext *ticP = TwapiGetBaseContext(interp);
    TWAPI_ASSERT(ticP);
    return TwapiRegisterPointerTic(ticP, p, typetag);
}

TCL_RESULT TwapiRegisterCountedPointerTic(TwapiInterpContext *ticP, const void *p, void *typetag)
{
    Tcl_HashEntry *he;
    int new_entry;
    TwapiRegisteredPointer *rP;

    TWAPI_ASSERT(BASE_CONTEXT(ticP));

    if (p == NULL)
        return TwapiReturnError(ticP->interp, TWAPI_NULL_POINTER);

    he = Tcl_CreateHashEntry(&BASE_CONTEXT(ticP)->pointers, p, &new_entry);
    TWAPI_ASSERT(he);
    if (new_entry) {
        rP = (TwapiRegisteredPointer *) ckalloc(sizeof(*rP));
        rP->tag = typetag;
        rP->nrefs = 1;
        Tcl_SetHashValue(he, rP);
    } else {
        rP = Tcl_GetHashValue(he);
        if (rP->nrefs < 0)
            return TwapiReturnError(ticP->interp, TWAPI_REGISTERED_POINTER_IS_NOT_COUNTED);
        if (rP->tag != typetag)
            return TwapiReturnError(ticP->interp, TWAPI_REGISTERED_POINTER_TAG_MISMATCH);
        rP->nrefs += 1;
    }
    return TCL_OK;
}

TCL_RESULT TwapiRegisterCountedPointer(Tcl_Interp *interp, const void *p, void *typetag)
{
    TwapiInterpContext *ticP = TwapiGetBaseContext(interp);
    TWAPI_ASSERT(ticP);
    return TwapiRegisterCountedPointerTic(ticP, p, typetag);
}

TCL_RESULT TwapiUnregisterPointerTic(TwapiInterpContext *ticP, const void *p, void *typetag)
{
    Tcl_HashEntry *he;
    int code;

    TWAPI_ASSERT(BASE_CONTEXT(ticP));

    if (p == NULL)
        return TwapiReturnError(ticP->interp, TWAPI_NULL_POINTER);

    he = Tcl_FindHashEntry(&BASE_CONTEXT(ticP)->pointers, p);
    code = TWAPI_REGISTERED_POINTER_NOTFOUND;
    if (he) {
        TwapiRegisteredPointer *rP = Tcl_GetHashValue(he);
        if (rP->tag != typetag)
            code = TWAPI_REGISTERED_POINTER_TAG_MISMATCH;
        else {
            /* For counted pointers, free if ref count reaches 0.
               For uncounted pointers ref count is set to -1 already */
            if (--(rP->nrefs) <= 0) {
                ckfree((char *) rP);
                Tcl_DeleteHashEntry(he);
            }
            return TCL_OK;
        }
    }

    return TwapiReturnError(ticP->interp, code);
}

TCL_RESULT TwapiUnregisterPointer(Tcl_Interp *interp, const void *p, void *typetag)
{
    TwapiInterpContext *ticP = TwapiGetBaseContext(interp);
    TWAPI_ASSERT(ticP);
    return TwapiUnregisterPointerTic(ticP, p, typetag);
}

