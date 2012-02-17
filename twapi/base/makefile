MODULENAME=twapi

!include ..\include\common.inc

CONFIGFLAGS = -DTWAPI_CORE_BUILD

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
	$(OBJDIR)\etw.obj \
	$(OBJDIR)\eventlog.obj \
	$(OBJDIR)\file.obj \
	$(OBJDIR)\hotkey.obj \
	$(OBJDIR)\gdi.obj \
	$(OBJDIR)\guids.obj \
	$(OBJDIR)\keylist.obj \
	$(OBJDIR)\lzmadec.obj \
	$(OBJDIR)\lzmainterface.obj \
	$(OBJDIR)\memlifo.obj \
	$(OBJDIR)\mstask.obj \
	$(OBJDIR)\mycrt.obj \
	$(OBJDIR)\namedpipe.obj \
	$(OBJDIR)\nls.obj \
	$(OBJDIR)\network.obj \
	$(OBJDIR)\os.obj \
	$(OBJDIR)\parseargs.obj \
	$(OBJDIR)\pdh.obj \
	$(OBJDIR)\power.obj \
	$(OBJDIR)\printer.obj \
	$(OBJDIR)\process.obj \
	$(OBJDIR)\recordarray.obj \
	$(OBJDIR)\resolve.obj \
	$(OBJDIR)\resource.obj \
	$(OBJDIR)\security.obj \
	$(OBJDIR)\scm.obj \
	$(OBJDIR)\service.obj \
	$(OBJDIR)\share.obj \
	$(OBJDIR)\shell.obj \
	$(OBJDIR)\tclobjs.obj \
	$(OBJDIR)\threadpool.obj \
	$(OBJDIR)\trap.obj \
	$(OBJDIR)\twapi.obj \
	$(OBJDIR)\twine.obj \
	$(OBJDIR)\ui.obj \
	$(OBJDIR)\util.obj \
	$(OBJDIR)\win.obj \
	$(OBJDIR)\wmi.obj \
	$(OBJDIR)\wts.obj



# Libraries common to all DLLs
LIBS          = $(TCLLIB) $(DLLIBC) $(WINLIB)


!IF "x$(TWAPI_STATIC_BUILD)" == "x"
CFLAGS = $(CFLAGS) -DUSE_TCL_STUBS
!ELSE
CFLAGS = $(CFLAGS) -DTWAPI_STATIC_BUILD -DSTATIC_BUILD=1
!ENDIF


# CRT=none support is not complete as yet
!IF "$(CRT)x" == "nonex"
CFLAGS = $(CFLAGS) /MT /Oi /Zl /DTWAPI_REPLACE_CRT
LENTRYFLAG = /ENTRY:DllMain
!ELSE
!IF "$(CRT)x" == "staticx"
CFLAGS = $(CFLAGS) /MT /Oi /DTWAPI_MINIMIZE_CRT
!ELSE
CFLAGS = $(CFLAGS) /D_DLL /MD
!ENDIF
!ENDIF

# Linker options
LOPT      = /RELEASE /NOLOGO /MACHINE:$(MACHINE) /DLL /OPT:REF $(LENTRYFLAG)

!IFDEF ENABLE_DELAYLOAD
LIBS = $(LIBS) delayimp.lib
LOPT = $(LOPT) /DELAYLOAD:winmm.dll /DELAYLOAD:netapi32.dll /DELAYLOAD:mpr.dll /DELAYLOAD:iphlpapi.dll /DELAYLOAD:powrprof.dll /DELAYLOAD:userenv.dll /DELAYLOAD:setupapi.dll /DELAYLOAD:psapi.dll /DELAYLOAD:wtsapi32.dll /DELAYLOAD:pdh.dll
!ENDIF


!if $(VCVER) == 6
# VC 6 - /OPT:NOWIN98 reduces code size by using 512 byte section alignment
LOPT = $(LOPT) /debug:full /debugtype:cv /OPT:NOWIN98
!else
LOPT = $(LOPT) /debug
!endif

# Resource definitions
RCFLAGS = -DMODULEMAJOR=$(MODULEMAJOR) -DMODULEMINOR=$(MODULEMINOR) -DMODULEBUILD=$(MODULEBUILD) -DFILE_VERSION_STR=\"$(FILE_VERSION_STR)\" -DPRODUCT_VERSION_STR=\"$(MODULE_VERSION_STR)\" -DMODULE_STR=\"$(MODULE)\"

!IF "x$(EMBED_SCRIPT)" == "xlzma"
TWAPI_SCRIPT_RESOURCE_LZMA=combined_twapi_script.lzma
TWAPI_SCRIPT_RESOURCE_DIR=$(OBJDIR)\tcl
TWAPI_SCRIPT_RESOURCE_PATH=$(TWAPI_SCRIPT_RESOURCE_DIR)\$(TWAPI_SCRIPT_RESOURCE_LZMA)
RCFLAGS = $(RCFLAGS) -DTWAPI_SCRIPT_RESOURCE_LZMA="$(TWAPI_SCRIPT_RESOURCE_PATH)"
!ELSEIF "x$(EMBED_SCRIPT)" == "xplain"
TWAPI_SCRIPT_RESOURCE=combined_twapi_script.tcl
TWAPI_SCRIPT_RESOURCE_DIR=$(OBJDIR)\tcl
TWAPI_SCRIPT_RESOURCE_PATH=$(TWAPI_SCRIPT_RESOURCE_DIR)\$(TWAPI_SCRIPT_RESOURCE)
RCFLAGS = $(RCFLAGS) -DTWAPI_SCRIPT_RESOURCE="$(TWAPI_SCRIPT_RESOURCE_PATH)"
!ENDIF


all: $(OBJDIR)\$(MODULE)

$(OBJDIR)\$(MODULE): $(OBJDIR) $(OBJS) $(OBJDIR)\twapi.res makefile
        $(LINK) $(LOPT) -out:$@ $(LIBS) $(OBJS) $(OBJDIR)\twapi.res
!if $(VCVER) > 8
	cd $(OBJDIR) && $(MT) -manifest $(MODULE).manifest -outputresource:$(MODULE);2
!endif

$(OBJS): $(CINCLUDES) makefile

$(OBJDIR)\twapi.res: $(OBJDIR) twapi.rc makefile $(TWAPI_SCRIPT_RESOURCE_PATH)
        rc $(INCFLAGS) $(RCFLAGS) -fo $@ twapi.rc

!ifdef TWAPI_SCRIPT_RESOURCE
$(TWAPI_SCRIPT_RESOURCE_PATH):
	cd ..\tcl && nmake PACKAGEDIR="..\base\$(TWAPI_SCRIPT_RESOURCE_DIR)" COMBINEOUTPUT="$(TWAPI_SCRIPT_RESOURCE)" combine
!endif

!ifdef TWAPI_SCRIPT_RESOURCE_LZMA
$(TWAPI_SCRIPT_RESOURCE_PATH):
	cd ..\tcl && nmake PACKAGEDIR="..\base\$(TWAPI_SCRIPT_RESOURCE_DIR)" COMBINEOUTPUT="$(TWAPI_SCRIPT_RESOURCE_LZMA).raw" combine
	"$(TOOLDIR)\lzma.exe" e "$(TWAPI_SCRIPT_RESOURCE_PATH).raw" "$(TWAPI_SCRIPT_RESOURCE_PATH)"
!endif

$(OBJDIR):
	@if not exist $(OBJDIR)\nul mkdir $(OBJDIR)

lib: $(OBJDIR)\$(MODULELIB)

$(OBJDIR)\$(MODULELIB): $(OBJDIR) $(OBJS) makefile
	cd ..\tcl && nmake PACKAGEDIR="..\base\$(OBJDIR)\tcl" COMBINEOUTPUT="combined_twapi_script.lzma.raw" combine
	"$(TOOLDIR)\lzma.exe" e "$(OBJDIR)\tcl\combined_twapi_script.lzma.raw" "$(OBJDIR)\combined_twapi_script.lzma"
	$(LIBTOOL) /OUT:$@ $(OBJS)

distribution:
        $(COPY) $(OBJDIR)\$(MODULE) "$(PACKAGEDIR)"

printioctl: printioctl.exe

printioctl.exe: printioctl.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printioctl.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib

printdevices: printdevices.exe

printdevices.exe: printdevices.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printdevices.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib

makefile: ../include/common.inc

clean:
        -$(DEL) *.i~ *.c~ *.h~
        -rmdir /s /q $(OBJDIR)
        -$(DEL) printioctl.obj


