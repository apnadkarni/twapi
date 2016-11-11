MODULENAME=twapi_base

!include ..\include\common.inc

OBJS  = \
	$(OBJDIR)\adsi.obj \
	$(OBJDIR)\async.obj \
	$(OBJDIR)\calls.obj \
	$(OBJDIR)\errors.obj \
	$(OBJDIR)\ffi.obj \
	$(OBJDIR)\keylist.obj \
	$(OBJDIR)\lzmadec.obj \
	$(OBJDIR)\lzmainterface.obj \
	$(OBJDIR)\memlifo.obj \
	$(OBJDIR)\mycrt.obj \
	$(OBJDIR)\parseargs.obj \
	$(OBJDIR)\printer.obj \
	$(OBJDIR)\recordarray.obj \
	$(OBJDIR)\tclobjs.obj \
	$(OBJDIR)\threadpool.obj \
	$(OBJDIR)\trap.obj \
	$(OBJDIR)\twapi.obj \
	$(OBJDIR)\twine.obj \
	$(OBJDIR)\util.obj \
	$(OBJDIR)\win.obj \
	$(OBJDIR)\winchars.obj

TCLFILES=..\tcl\twapi.tcl \
	..\tcl\base.tcl \
	..\tcl\handle.tcl \
	..\tcl\win.tcl \
	..\tcl\adsi.tcl

EXTRALIBS = gdi32.lib  $(SDKLIB)credui.lib

!include ..\include\rules.inc

printioctl: printioctl.exe

printioctl.exe: printioctl.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printioctl.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printioctl.exe"

printdevices: printdevices.exe

printdevices.exe: printdevices.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printdevices.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printdevices.exe"



