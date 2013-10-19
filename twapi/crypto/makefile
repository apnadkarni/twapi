MODULENAME=twapi_crypto

EXTRALIBS = crypt32.lib

!include ..\include\common.inc

OBJS  = $(OBJDIR)\crypto.obj $(OBJDIR)\sspi.obj
TCLFILES=..\tcl\crypto.tcl ..\tcl\sspi.tcl ..\tcl\tls.tcl

!include ..\include\rules.inc

$(OBJS): twapi_crypto.h
