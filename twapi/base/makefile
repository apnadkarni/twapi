MODULENAME=twapi_base

!include ..\include\common.inc

OBJS  = \
	$(OBJDIR)\adsi.obj \
	$(OBJDIR)\async.obj \
	$(OBJDIR)\calls.obj \
	$(OBJDIR)\com.obj \
	$(OBJDIR)\errors.obj \
	$(OBJDIR)\guids.obj \
	$(OBJDIR)\keylist.obj \
	$(OBJDIR)\lzmadec.obj \
	$(OBJDIR)\lzmainterface.obj \
	$(OBJDIR)\memlifo.obj \
	$(OBJDIR)\mycrt.obj \
	$(OBJDIR)\nls.obj \
	$(OBJDIR)\os.obj \
	$(OBJDIR)\parseargs.obj \
	$(OBJDIR)\printer.obj \
	$(OBJDIR)\recordarray.obj \
	$(OBJDIR)\security.obj \
	$(OBJDIR)\tclobjs.obj \
	$(OBJDIR)\threadpool.obj \
	$(OBJDIR)\trap.obj \
	$(OBJDIR)\twapi.obj \
	$(OBJDIR)\twine.obj \
	$(OBJDIR)\util.obj \
	$(OBJDIR)\win.obj \
	$(OBJDIR)\wmi.obj \
	$(OBJDIR)\wts.obj


TCLFILES=..\tcl\twapi.tcl \
	..\tcl\handle.tcl \
	..\tcl\metoo.tcl \
	..\tcl\osinfo.tcl \
	..\tcl\security.tcl \
	..\tcl\nls.tcl \
	..\tcl\com.tcl \
	..\tcl\adsi.tcl \
	..\tcl\synch.tcl \
        ..\tcl\printer.tcl \
	..\tcl\msi.tcl \
	..\tcl\power.tcl \
	..\tcl\rds.tcl \
	..\tcl\wmi.tcl

!include ..\include\rules.inc

printioctl: printioctl.exe

printioctl.exe: printioctl.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printioctl.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printioctl.exe"

printdevices: printdevices.exe

printdevices.exe: printdevices.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printdevices.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printdevices.exe"



