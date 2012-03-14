
# Common definition file for paths to various tools

# Following the Tcl build style, we assume that either the vcvars32.bat
# from Visual C++ (optional) and setenv.cmd have been run
# to set up the compiler and target environment. The "official" build
# of TWAPI uses the Visual C++ 6.0 compiler for 32-bit builds and Platform SDK
# for Windows 2003 SP1 compiler for 64-bit builds.
# So for example, to build the 32-bit TWAPI, from the command line
#   "c:\program files\microsoft visual studio\vc98\bin\vcvars32"
#   "c:\program files\microsoft platform sdk\setenv.cmd /2000 /RETAIL
# To build 64-bit TWAPI,
#   "c:\program files\microsoft platform sdk\setenv.cmd /X64 /RETAIL

# Set the compilation platform
!ifndef MACHINE
!if "$(CPU)" == "" || "$(CPU)" == "i386"
MACHINE		= X86
!else
MACHINE         = $(CPU)
!endif
!endif

!ifndef TOOLDIR
TOOLDIR=$(SRCROOT)\..\tools
!endif

################# COMPILER AND SDK #####################

# Now the SDK

!ifdef MSSDK
SDK           = "$(MSSDK)"
!else
SDK           = "c:\program files\microsoft platform sdk"
!endif

CC            = cl.exe
LINK          = link.exe
LIBTOOL       = lib.exe
MT            = mt.exe


# TCLSH is the Tcl shell path at build time, NOT THE TARGET. It is used
# to invoke build scripts.
TCLSH         = "$(SRCROOT)\..\tools\tclkitsh.exe"

# TCLSH_TARGET is the Tcl shell path on target installations. Used to install
# twapi. Need not be defined unless you are using the makefile to install
!ifndef TCLSH_TARGET
TCLSH_TARGET = tclsh86t
!endif



################################################################
# Definitions for general tools
DEL           = del
COPY          = copy /y >NUL
RMDIR	      = rmdir /S /Q
ZIP           = "$(TOOLDIR)"\zip.exe
MKDIR         = mkdir

