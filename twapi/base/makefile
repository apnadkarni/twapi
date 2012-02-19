MODULENAME=twapi_base

!include ..\include\common.inc

OBJS  = \
	$(OBJDIR)\adsi.obj \
	$(OBJDIR)\async.obj \
	$(OBJDIR)\calls.obj \
	$(OBJDIR)\clipboard.obj \
	$(OBJDIR)\com.obj \
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
	$(OBJDIR)\os.obj \
	$(OBJDIR)\parseargs.obj \
	$(OBJDIR)\pdh.obj \
	$(OBJDIR)\power.obj \
	$(OBJDIR)\printer.obj \
	$(OBJDIR)\process.obj \
	$(OBJDIR)\recordarray.obj \
	$(OBJDIR)\resource.obj \
	$(OBJDIR)\security.obj \
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


TCLFILES=..\tcl\twapi.tcl \
	..\tcl\handle.tcl \
	..\tcl\metoo.tcl \
	..\tcl\osinfo.tcl \
	..\tcl\security.tcl \
	..\tcl\process.tcl \
	..\tcl\disk.tcl \
	..\tcl\ui.tcl \
	..\tcl\clipboard.tcl \
	..\tcl\nls.tcl \
	..\tcl\shell.tcl \
	..\tcl\com.tcl \
	..\tcl\eventlog.tcl \
	..\tcl\adsi.tcl \
	..\tcl\pdh.tcl \
	..\tcl\process2.tcl \
	..\tcl\accounts.tcl \
	..\tcl\share.tcl \
	..\tcl\synch.tcl \
	..\tcl\desktop.tcl \
        ..\tcl\printer.tcl \
	..\tcl\mstask.tcl \
	..\tcl\msi.tcl \
	..\tcl\crypto.tcl \
        ..\tcl\device.tcl \
	..\tcl\power.tcl \
	..\tcl\namedpipe.tcl \
	..\tcl\resource.tcl \
	..\tcl\rds.tcl \
	..\tcl\wmi.tcl \
	..\tcl\etw.tcl

!include ..\include\rules.inc

printioctl: printioctl.exe

printioctl.exe: printioctl.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printioctl.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printioctl.exe"

printdevices: printdevices.exe

printdevices.exe: printdevices.c makefile
        $(CC) -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -nologo $(INCFLAGS) printdevices.c $(SDKLIB)ole32.lib $(SDKLIB)setupapi.lib /Fe"$(OBJDIR)\printdevices.exe"



