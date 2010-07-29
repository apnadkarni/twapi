TWAPI_SRCROOT = $(MAKEDIR)\..
!include ..\common.inc

!IFDEF SERVERONLY
!IFDEF DESKTOPONLY
!ERROR Must define only one of SERVERONLY or DESKTOPONLY
!ENDIF
CONFIGFLAGS = -DTWAPI_NODESKTOP -DTWAPI_LEAN
!ELSE
!IFDEF DESKTOPONLY
CONFIGFLAGS = -DTWAPI_NOSERVER -DTWAPI_LEAN
!ELSE
# No special flags
!ENDIF
!ENDIF

!if "$(MACHINE)" == "AMD64"
MODULE        = $(PACKAGE)64.dll
!else
MODULE        = $(PACKAGE).dll
!endif

# Output directories
!IFNDEF OUTDIR
OUTDIR = build\$(MACHINE)
!ENDIF

!IFDEF DEBUG
OBJDIR = $(OUTDIR)\debug$(CONFIGNAME)
!ELSE
OBJDIR = $(OUTDIR)\release$(CONFIGNAME)
!ENDIF

OBJS  = \
	$(OBJDIR)\adsi.obj \
	$(OBJDIR)\async.obj \
	$(OBJDIR)\calls.obj \
	$(OBJDIR)\clipboard.obj \
	$(OBJDIR)\com.obj \
	$(OBJDIR)\console.obj \
	$(OBJDIR)\crypto.obj \
	$(OBJDIR)\device.obj \
	$(OBJDIR)\dirmonitor.obj \
	$(OBJDIR)\disk.obj \
	$(OBJDIR)\errors.obj \
	$(OBJDIR)\eventlog.obj \
	$(OBJDIR)\file.obj \
	$(OBJDIR)\hotkey.obj \
	$(OBJDIR)\gdi.obj \
	$(OBJDIR)\guids.obj \
	$(OBJDIR)\keylist.obj \
	$(OBJDIR)\memlifo.obj \
	$(OBJDIR)\mstask.obj \
	$(OBJDIR)\nls.obj \
	$(OBJDIR)\network.obj \
	$(OBJDIR)\parseargs.obj \
	$(OBJDIR)\pdh.obj \
	$(OBJDIR)\pipe.obj \
	$(OBJDIR)\power.obj \
	$(OBJDIR)\printer.obj \
	$(OBJDIR)\process.obj \
	$(OBJDIR)\recordarray.obj \
	$(OBJDIR)\resolve.obj \
	$(OBJDIR)\security.obj \
	$(OBJDIR)\scm.obj \
	$(OBJDIR)\service.obj \
	$(OBJDIR)\share.obj \
	$(OBJDIR)\shell.obj \
	$(OBJDIR)\system.obj \
	$(OBJDIR)\tclobjs.obj \
	$(OBJDIR)\try.obj \
	$(OBJDIR)\twapi.obj \
	$(OBJDIR)\twine.obj \
	$(OBJDIR)\ui.obj \
	$(OBJDIR)\util.obj \
	$(OBJDIR)\win.obj \
	$(OBJDIR)\wts.obj

!IFNDEF SERVERONLY
OBJS  = $(OBJS) \
!ENDIF


IFILE         = $(MODULENAME)
CINCLUDES     = twapi.h twapi_ddkdefs.h twapi_sdkdefs.h zlist.h memlifo.h


# General tools
DEL           = del
COPY          = copy

# Make the directory
!if ![if not exist $(OBJDIR) mkdir $(OBJDIR)]
!message Output directory is $(OBJDIR)
!else
!message Output directory $(OBJDIR) exists
!endif

# Location of the Visual C++ tools (32 bit assumed)
VCVER=0
!if ![echo VCVERSION=_MSC_VER > $(OBJDIR)\vercl.x] \
    && ![cl -nologo -TC -E $(OBJDIR)\vercl.x > $(OBJDIR)\vercl.p 2>NUL]
!include $(OBJDIR)\vercl.p
!endif

!if $(VCVERSION) >= 1500
VCVER=9
!elseif $(VCVERSION) >= 1400
VCVER=8
!elseif $(VCVERSION) >= 1300
VCVER=7
!elseif $(VCVERSION) >= 1200
VCVER=6
!endif

!if $(VCVER) == 0
!error Could not determine compiler version
!else
!message Compiling with Visual C++ $(VCVER).
!endif

!ifdef MSSDK
SDK           = "$(MSSDK)"
!else
SDK           = "c:\program files\microsoft platform sdk"
!endif
CC            = cl.exe
LINK          = link.exe

# TCL resources. By default, we will use 8.5 for 64-bit builds and 8.4
# for 32-bit builds.
!if "x$(TCLVER)" == "x"
!if "$(MACHINE)" == "AMD64"
TCLVER        = 85
!else
TCLVER        = 84
!endif
!endif
!if "x$(TCLDIR)" == "x"
TCLDIR        = "$(TWAPI_SRCROOT)\..\tcl-tk\$(TCLVER)"
!endif
TCLLIB        = $(TCLDIR)\lib\Release_$(MACHINE)\tclstub$(TCLVER).lib


# Windows libraries
# If building with Visual C++ 6.0, we need SDK libraries for newer API's.
# However, we cannot use all SDK libraries since the 6.0 linker does not
# understand some of the newer SDK library format in which case we have
# to stick with the compiler libraries. For newer compilers, we never
# use the SDK libraries.
!if $(VCVER) > 6
SDKLIB =
!else
SDKLIB = $(SDK)\lib\ # This comment to ensure line does not end in backslash
LIB=$(MSVCDIR)\lib
!endif

WINLIB        = \
	$(SDKLIB)kernel32.lib \
	$(SDKLIB)advapi32.lib \
	$(SDKLIB)user32.lib \
	rpcrt4.lib \
	gdi32.lib \
	$(SDKLIB)psapi.lib \
	$(SDKLIB)netapi32.lib \
	$(SDKLIB)pdh.lib \
	winmm.lib \
	mpr.lib \
	$(SDKLIB)ws2_32.lib \
	ole32.lib \
	oleaut32.lib \
	shell32.lib  \
	$(SDKLIB)winspool.lib \
	$(SDKLIB)version.lib \
	$(SDKLIB)iphlpapi.lib \
	$(SDKLIB)powrprof.lib \
	$(SDKLIB)secur32.lib \
	$(SDKLIB)userenv.lib \
	$(SDKLIB)wtsapi32.lib \
	$(SDKLIB)setupapi.lib \
	$(SDKLIB)strsafe.lib

!if $(VCVER) > 6
WINLIB = $(WINLIB) bufferoverflowU.lib
!endif

# Libraries common to all DLLs
LIBS          = $(TCLLIB) $(DLLIBC) $(WINLIB)

!IFDEF DEBUG
CDEBUG = -Z7 -Od -D_DEBUG
!ELSE
!IFDEF NOOPTIMIZE
CDEBUG = -Zi -Od
!ELSE
CDEBUG = -O1 -Zi -DNDEBUG -W3
!ENDIF
!ENDIF

!IFDEF ENABLEASSERTS
CDEBUG = $(CDEBUG) -DTWAPI_ENABLE_ASSERT
!ENDIF

# C compiler flags

INCFLAGS       =  -I. -I$(TCLDIR)\include

CFLAGS    = $(CDEBUG) /c /nologo $(CRUNTIME) /DWIN32 /D_WIN32 /D_WINDOWS -DTCL_THREADS=1 -DUSE_TCL_STUBS -D_WIN32_WINNT=0x0500 -DPSAPI_VERSION=1 $(INCFLAGS) $(CONFIGFLAGS)
!if "$(MACHINE)" == "AMD64"
CFLAGS = $(CFLAGS) /Wp64
!endif

# Note CRT=none option does not work. When building with _WINNT_VER=0x5000
# (ie. Win2K) the winsocket headers define inline functions that make use
# of the C runtime. So CRT=none needs to be postponed until such a time
# that we build for XP and later only.
!IF "$(CRT)x" == "nonex"
CFLAGS = $(CFLAGS) /MT /Oi /Zl
!ELSE
!IF "$(CRT)x" == "staticx"
CFLAGS = $(CFLAGS) /MT /Oi
!ELSE
CFLAGS = $(CFLAGS) /D_DLL /MD
!ENDIF
!ENDIF

# Linker options
LOPT      = /RELEASE /NOLOGO /MACHINE:$(MACHINE) /DLL /OPT:NOWIN98 /OPT:REF

!if $(VCVER) == 6
LOPT = $(LOPT) /debug:full /debugtype:cv 
!endif

all: $(OBJDIR)\$(MODULE)

$(OBJDIR)\$(MODULE): $(OBJDIR) $(OBJS) $(OBJDIR)\twapi.res makefile
#        rc $(INCFLAGS) -DMODULEMAJOR=$(MODULEMAJOR) -DMODULEMINOR=$(MODULEMINOR) -DMODULEBUILD=$(MODULEBUILD) -DFILE_VERSION_STR=\"$(FILE_VERSION_STR)\" -DPRODUCT_VERSION_STR=\"$(MODULE_VERSION_STR)\" twapi.rc
        $(LINK) $(LOPT) -out:$@ $(LIBS) $(OBJS) $(OBJDIR)\twapi.res

$(OBJS): $(CINCLUDES) makefile


$(OBJDIR)\twapi.res: $(OBJDIR) twapi.rc makefile
        rc $(INCFLAGS) -DMODULEMAJOR=$(MODULEMAJOR) -DMODULEMINOR=$(MODULEMINOR) -DMODULEBUILD=$(MODULEBUILD) -DFILE_VERSION_STR=\"$(FILE_VERSION_STR)\" -DPRODUCT_VERSION_STR=\"$(MODULE_VERSION_STR)\" -DMODULE_STR=\"$(MODULE)\" -fo $@ twapi.rc

$(OBJDIR):
	@if not exist $(OBJDIR)\nul mkdir $(OBJDIR)

distribution:
        $(COPY) $(OBJDIR)\$(MODULE) "$(PACKAGEDIR)"

printioctl: printioctl.exe

printioctl.exe: printioctl.c makefile
        $(CC) -D_WIN32_WINNT=0x0500 -nologo $(INCFLAGS) printioctl.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib

printdevices: printdevices.exe

printdevices.exe: printdevices.c makefile
        $(CC) -D_WIN32_WINNT=0x0500 -nologo $(INCFLAGS) printdevices.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib

makefile: ../common.inc

clean:
        -$(DEL) *.i~ *.c~ *.h~
        -rmdir /s /q $(OBJDIR)
        -$(DEL) printioctl.obj

.c{$(OBJDIR)}.obj::
       $(CC) $(CFLAGS) -Fo$(OBJDIR)\ $<

