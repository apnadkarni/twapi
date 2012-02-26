MODULENAME=twapi_ui

!include ..\include\common.inc

OBJS  = $(OBJDIR)\ui.obj $(OBJDIR)\gdi.obj
TCLFILES=..\tcl\ui.tcl
LIBS =	$(LIBS) gdi32.lib $(SDKLIB)uxtheme.lib

!include ..\include\rules.inc

$(OBJS) : twapi_ui.h

