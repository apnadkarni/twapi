MODULENAME=twapi_crypto

EXTRALIBS = crypt32.lib wintrust.lib

!include ..\include\common.inc

OBJS  = $(OBJDIR)\crypto.obj $(OBJDIR)\sspi.obj $(OBJDIR)\pbkdf2.obj
TCLFILES=..\tcl\crypto.tcl ..\tcl\sspi.tcl ..\tcl\tls.tcl

!include ..\include\rules.inc

$(OBJS): twapi_crypto.h
