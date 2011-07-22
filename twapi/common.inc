# Common definitions for all directories

!ifndef TWAPI_SRCROOT
TWAPI_SRCROOT=$(MAKEDIR)
!endif

!ifndef TOOLDIR
TOOLDIR=$(TWAPI_SRCROOT)\..\tools
!endif

!ifndef HOME
HOME=%HOMEDRIVE%\%HOMEPATH%
!endif

################################################################
# Compiler and target environment

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


################################################################
# TWAPI feature options

# If neither SERVERONLY nor DESKTOPONLY are defined, the whole
# shebang is built. Else only the server components, or the
# desktop components are built.
!IFDEF SERVERONLY
!IFDEF DESKTOPONLY
!ERROR Must not define both SERVERONLY and DESKTOPONLY
!ENDIF
CONFIGNAME  = _server
LEAN        = 1
!ELSE
!IFDEF DESKTOPONLY
CONFIGNAME  = _desktop
LEAN        = 1
!ELSE
CONFIGNAME  =
!ENDIF
!ENDIF

################################################################
# Module name and version info

MODULENAME    = twapi
MODULEMAJOR   = 3
MODULEMINOR   = 1
MODULEBUILD   = 14
# RELEASETYPE: a is alpha, b is beta
RELEASETYPE = a

PACKAGE=$(MODULENAME)$(CONFIGNAME)

!if "x$(RELEASETYPE)" == "x"
RELEASETYPE = .
!endif

# Internal version as used within Tcl code
MODULEVERSION = $(MODULEMAJOR).$(MODULEMINOR)$(RELEASETYPE)$(MODULEBUILD)
FILE_VERSION_STR = "$(MODULEVERSION)"

# External version string as stored in a resource file
!if "x$(RELEASETYPE)" == "xa"
MODULE_VERSION_STR = "$(FILE_VERSION_STR) Alpha"
!elseif "x$(RELEASETYPE)" == "xb"
MODULE_VERSION_STR = "$(FILE_VERSION_STR) Beta"
!else
MODULE_VERSION_STR = $(FILE_VERSION_STR)
!endif

################################################################

# TCLSH is the Tcl shell path at build time, NOT THE TARGET. It is used
# to invoke build scripts.
TCLSH         = "$(TWAPI_SRCROOT)\..\tools\tclkitsh.exe"

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
