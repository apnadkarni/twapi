MODULENAME=twapi_base

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



# Targets


$(OBJDIR)\$(MODULE): $(OBJS) 

$(OBJS): $(CINCLUDES) makefile

!ifdef TWAPI_SCRIPT_RESOURCE
$(TWAPI_SCRIPT_RESOURCE_PATH):
	cd ..\tcl && nmake PACKAGEDIR="..\base\$(TWAPI_SCRIPT_RESOURCE_DIR)" COMBINEOUTPUT="$(TWAPI_SCRIPT_RESOURCE)" combine
!endif

!ifdef TWAPI_SCRIPT_RESOURCE_LZMA
$(TWAPI_SCRIPT_RESOURCE_PATH):
	cd ..\tcl && nmake PACKAGEDIR="..\base\$(TWAPI_SCRIPT_RESOURCE_DIR)" COMBINEOUTPUT="$(TWAPI_SCRIPT_RESOURCE_LZMA).raw" combine
	"$(TOOLDIR)\lzma.exe" e "$(TWAPI_SCRIPT_RESOURCE_PATH).raw" "$(TWAPI_SCRIPT_RESOURCE_PATH)"
!endif

$(OBJDIR)\$(MODULELIB): $(OBJS)

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


