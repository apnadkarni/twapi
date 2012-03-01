# Include file for building all modules
# The defines assume nmake is executed in the directory of the module

!include version.inc

# ================ MANDATORY DEFINES ================

# Each module should define the following before including this file

!ifndef MODULENAME
# TBD
# !error MODULENAME not defined
!endif

########################## OPTIONAL DEFINES #######################
#
# In addition, the following may be *optionally* defined by each module

# Define SRCROOT if module is in a non-standard place in directory tree
!ifndef SRCROOT
SRCROOT=$(MAKEDIR)\..
!endif

# Major version of module defaults to TWAPI major version
!ifndef MODULEMAJOR
MODULEMAJOR = $(TWAPI_MAJOR)
!endif

# Minor version of module defaults to TWAPI minor version
!ifndef MODULEMINOR
MODULEMINOR = $(TWAPI_MINOR)
!endif

# Build number of module defaults to TWAPI build number
!ifndef MODULEBUILD
MODULEBUILD = $(TWAPI_BUILD)
!endif

# Release type module defaults to TWAPI release type
!ifndef RELEASETYPE
RELEASETYPE = $(TWAPI_RELEASETYPE)
!endif
!if "x$(RELEASETYPE)" == "x"
RELEASETYPE = .
!endif

# Name of the package defaults to module name
!ifndef PACKAGE
!ifndef MODULENAME
!error Neither PACKAGE nor MODULENAME are defined
!else
PACKAGE=$(MODULENAME)
!endif
!endif

# Some modules that only run on later versions of Windows may
# define this appropriately. By default, you get XP (5.1)
!ifndef TWAPI_WIN_HEADER_VERSION
TWAPI_WIN_HEADER_VERSION = 0x0501
!endif

# Selects whether we create Unicode UCS-2 or UTF-8 objects.
# Note both support same range of characters. The choice depends
# on speed of object creation versus memory.
# Note it is OK to mix modules that define this with those that do not
!ifndef USE_UNICODE_OBJ
USE_UNICODE_OBJ = 1
!endif


################ Command line Defines ################
# Optionally define these on the command line

# Top level directory where output is written
!ifndef OUTDIR
OUTDIR = $(SRCROOT)\build
!endif

# Define TWAPI_STATIC_BUILD=1 if building static twapi libraries
!ifdef TWAPI_STATIC_BUILD
!if ! ("$(TWAPI_STATIC_BUILD)" == "1")
!undef TWAPI_STATIC_BUILD
!endif
!endif

# Define EMBED_SCRIPT as "plain" or "lzma" to embed Tcl scripts
# as a resource
# EMBED_SCRIPT=

# Define TWAPI_SINGLE_MODULE=1 if all modules are built as part of a single DLL
!ifndef TWAPI_SINGLE_MODULE
!error Oh no
!endif

!ifdef TWAPI_SINGLE_MODULE
!if ! ("$(TWAPI_SINGLE_MODULE)" == "1")
!undef TWAPI_SINGLE_MODULE
!endif
!endif


# Define a BUILDID for each build. This is passed to a Tcl script.
!ifndef BUILDID
BUILDID = [clock seconds]
!endif

################ Product version definitions ################

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



################# TARGET ARCHITECTURE ###################

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

################# OUTPUT DIRECTORIES / FILES ###################

# Output directories
!ifdef DEBUG
OBJDIR = $(OUTDIR)\$(MACHINE)\debug
!else
OBJDIR = $(OUTDIR)\$(MACHINE)\release
!endif

!if "$(EMBED_SCRIPT)" == "lzma"
OBJDIR = $(OBJDIR)_embed_lzma
!elseif "$(EMBED_SCRIPT)" == "plain"
OBJDIR = $(OBJDIR)_embed
!endif

!ifdef TWAPI_STATIC_BUILD
OBJDIR = $(OBJDIR)_static
!else
!ifdef TWAPI_SINGLE_MODULE
OBJDIR = $(OBJDIR)_single
!endif
!endif

OBJDIR = $(OBJDIR)\$(MODULENAME)

# Make the directory
!if ![if not exist $(OBJDIR) mkdir $(OBJDIR)]
# !message Output directory is $(OBJDIR)
!else
# !message Output directory $(OBJDIR) exists
!endif

# 
!if "$(MACHINE)" == "AMD64"
MODULE        = $(PACKAGE)64.dll
MODULELIB     = $(PACKAGE)64.lib
!else
MODULE        = $(PACKAGE).dll
MODULELIB     = $(PACKAGE).lib
!endif


###################### TOOL PATHS ########################

!ifndef TOOLDIR
TOOLDIR=$(SRCROOT)\..\tools
!endif

# TBD - why required ?
!ifndef HOME
HOME=%HOMEDRIVE%\%HOMEPATH%
!endif

# TCL resources. By default, we will use 8.5 
!if "x$(TCLVER)" == "x"
TCLVER        = 85
!endif
!if "x$(TCLDIR)" == "x"
TCLDIR        = $(SRCROOT)\..\tcl-tk\$(TCLVER)
!endif
TCLLIB        = $(TCLDIR)\lib\Release_$(MACHINE)\tclstub$(TCLVER).lib


# TCLSH is the Tcl shell path at build time, NOT THE TARGET. It is used
# to invoke build scripts.
TCLSH         = "$(SRCROOT)\..\tools\tclkitsh.exe"

# TCLSH_TARGET is the Tcl shell path on target installations. Used to install
# twapi. Need not be defined unless you are using the makefile to install
!ifndef TCLSH_TARGET
TCLSH_TARGET = tclsh86t
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

# Compiler
# --------

# Visual C++ version
VCVER=0
!if ![echo VCVERSION=_MSC_VER > $(OBJDIR)\vercl.x] \
    && ![$(CC) -nologo -TC -E $(OBJDIR)\vercl.x > $(OBJDIR)\vercl.p 2>NUL]
!include $(OBJDIR)\vercl.p
!endif

!if $(VCVERSION) >= 1600
VCVER=10
!elseif $(VCVERSION) >= 1500
VCVER=9
!elseif $(VCVERSION) >= 1400
VCVER=8
!elseif $(VCVERSION) >= 1300
VCVER=7
!elseif $(VCVERSION) >= 1200
VCVER=6
!endif

!if $(VCVER) == 0
!error Could not determine compiler version
!else
# !message Compiling with Visual C++ $(VCVER).
!endif

# Compiler flags 

!ifdef DEBUG
CDEBUG = -Zi -Od -D_DEBUG  -Fd"$(OBJDIR)"
!else
!ifdef NOOPTIMIZE
CDEBUG = -Zi -Od -Fd"$(OBJDIR)"
!else
CDEBUG = -O1 -Zi -DNDEBUG -W3 -Fd"$(OBJDIR)"
!endif
!endif

!ifdef ENABLE_ASSERTS
CDEBUG = $(CDEBUG) -DTWAPI_ENABLE_ASSERT=$(ENABLE_ASSERTS)
!endif

INCFLAGS       =  -I. -I..\include -I$(TCLDIR)\include

CINCLUDES     = $(SRCROOT)\include\twapi.h \
		$(SRCROOT)\include\twapi_wm.h \
		$(SRCROOT)\include\twapi_ddkdefs.h \
		$(SRCROOT)\include\twapi_sdkdefs.h \
		$(SRCROOT)\include\zlist.h \
		$(SRCROOT)\include\memlifo.h

CFLAGS    = $(CDEBUG) /c /nologo /DWIN32 /D_WIN32 /D_WINDOWS -DTCL_THREADS=1 -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -DPSAPI_VERSION=1 $(INCFLAGS) -DMODULENAME=\"$(MODULENAME)\" -D$(MODULENAME)_BUILD

!if "$(MACHINE)" == "AMD64"
CFLAGS = $(CFLAGS) /Wp64
!endif

!ifdef USE_UNICODE_OBJ
CFLAGS = $(CFLAGS) /DUSE_UNICODE_OBJ=$(USE_UNICODE_OBJ)
!endif

# What version of shell32.dll to use. XP or later -> IE 6
!if $(TWAPI_WIN_HEADER_VERSION) >= 0x0501
CFLAGS  = $(CFLAGS) -D_WIN32_IE=0x0600
!else
CFLAGS  = $(CFLAGS) -D_WIN32_IE=0x0500
!endif

# CRT=none support is not complete as yet
!if "$(CRT)x" == "nonex"
CFLAGS = $(CFLAGS) /MT /Oi /Zl /DTWAPI_REPLACE_CRT
LENTRYFLAG = /ENTRY:DllMain
!else
!if "$(CRT)x" == "staticx"
CFLAGS = $(CFLAGS) /MT /Oi /DTWAPI_MINIMIZE_CRT
!else
CFLAGS = $(CFLAGS) /D_DLL /MD
!endif
!endif

!if "$(TWAPI_STATIC_BUILD)" == ""
CFLAGS = $(CFLAGS) -DUSE_TCL_STUBS
!else
CFLAGS = $(CFLAGS) -DTWAPI_STATIC_BUILD
!endif

!if "$(TWAPI_SINGLE_MODULE)" != ""
CFLAGS = $(CFLAGS) -DTWAPI_SINGLE_MODULE
!endif

# Libraries
# ---------

# Windows libraries
# If building with Visual C++ 6.0, we need SDK libraries for newer API's.
# However, we cannot use all SDK libraries since the 6.0 linker does not
# understand some of the newer SDK library format in which case we have
# to stick with the compiler libraries. For newer compilers, we never
# use the SDK libraries.
!if $(VCVER) > 6
SDKLIB =
!else
SDKLIB = $(SDK)\lib\ # This comment to ensure line does not end in backslash
LIB=$(MSVCDIR)\lib
!endif


WINLIB        = \
	$(SDKLIB)kernel32.lib \
	$(SDKLIB)advapi32.lib \
	$(SDKLIB)user32.lib \
	rpcrt4.lib \
	$(SDKLIB)psapi.lib \
	$(SDKLIB)netapi32.lib \
	mpr.lib \
	$(SDKLIB)ws2_32.lib \
	ole32.lib \
	oleaut32.lib \
	$(SDKLIB)winspool.lib \
	$(SDKLIB)version.lib \
	$(SDKLIB)powrprof.lib \
	$(SDKLIB)secur32.lib \
	$(SDKLIB)wtsapi32.lib 

# Disabled for now since it pulls in _vsnprintf - $(SDKLIB)strsafe.lib

!if $(VCVER) > 6 && $(VCVER) < 9
WINLIB = $(WINLIB) bufferoverflowU.lib
!endif

LIBS          = $(TCLLIB) $(WINLIB) $(EXTRALIBS)

# Modules have to link with the twapi_base dll
!if "$(MODULENAME)" != "twapi_base" && !defined(TWAPI_SINGLE_MODULE)
LIBS = "$(OBJDIR)\..\twapi_base\twapi_base.lib" $(LIBS)
!endif

# Resource definitions

!ifndef RC_FILE_NAME
RC_FILE_NAME = $(SRCROOT)/include/common.rc
!endif

RCFLAGS = -DMODULEMAJOR=$(MODULEMAJOR) -DMODULEMINOR=$(MODULEMINOR) -DMODULEBUILD=$(MODULEBUILD) -DFILE_VERSION_STR=\"$(FILE_VERSION_STR)\" -DPRODUCT_VERSION_STR=\"$(MODULE_VERSION_STR)\" -DMODULE_STR=\"$(MODULE)\" -DMODULENAME=$(MODULENAME)

!if defined(TWAPI_STATIC_BUILD) && ! defined(EMBED_SCRIPT)
EMBED_SCRIPT=lzma
!endif

!if defined(EMBED_SCRIPT)
!if "x$(EMBED_SCRIPT)" == "xlzma"
TWAPI_SCRIPT_RESOURCE_TYPE=tclscriptlzma
TWAPI_SCRIPT_RESOURCE=combined_$(MODULENAME)_script.lzma
!elseif "x$(EMBED_SCRIPT)" == "xplain"
TWAPI_SCRIPT_RESOURCE_TYPE=tclscript
TWAPI_SCRIPT_RESOURCE=combined_$(MODULENAME)_script
!else
!error EMBED_SCRIPT must be undefined, "lzma" or "plain"
!endif

TWAPI_SCRIPT_RESOURCE_DIR=$(OBJDIR)\tcl
TWAPI_SCRIPT_RESOURCE_PATH=$(TWAPI_SCRIPT_RESOURCE_DIR)\$(TWAPI_SCRIPT_RESOURCE)
!ifndef TWAPI_STATIC_BUILD
RCFLAGS = $(RCFLAGS) -DTWAPI_SCRIPT_RESOURCE_PATH="$(TWAPI_SCRIPT_RESOURCE_PATH)" -DTWAPI_SCRIPT_RESOURCE_TYPE=$(TWAPI_SCRIPT_RESOURCE_TYPE)
!endif
!endif

# Linking
# -------
LOPT      = /RELEASE /NOLOGO /MACHINE:$(MACHINE) /DLL /OPT:REF $(LENTRYFLAG)

!ifdef ENABLE_DELAYLOAD
LIBS = $(LIBS) delayimp.lib
LOPT = $(LOPT) /DELAYLOAD:netapi32.dll /DELAYLOAD:mpr.dll  /DELAYLOAD:powrprof.dll /DELAYLOAD:userenv.dll /DELAYLOAD:psapi.dll /DELAYLOAD:wtsapi32.dll
!endif


!if $(VCVER) == 6
# VC 6 - /OPT:NOWIN98 reduces code size by using 512 byte section alignment
LOPT = $(LOPT) /debug:full /debugtype:cv /OPT:NOWIN98
!else
LOPT = $(LOPT) /debug
!endif






################################################################
# Definitions for general tools
DEL           = del
COPY          = copy /y >NUL
RMDIR	      = rmdir /S /Q
ZIP           = "$(TOOLDIR)"\zip.exe
MKDIR         = mkdir

