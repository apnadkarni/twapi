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


!include ..\include\rules.inc

printioctl: printioctl.exe

printioctl.exe: printioctl.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printioctl.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printioctl.exe"

printdevices: printdevices.exe

printdevices.exe: printdevices.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printdevices.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printdevices.exe"



