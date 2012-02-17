# Include file for building all modules
# The defines assume nmake is executed in the directory of the module

!include version.inc

# ================ MANDATORY DEFINES ================

# Each module should define the following before including this file

!ifndef MODULENAME
!error MODULENAME not defined
!endif

!ifndef OBJS
#!error OBJS not defined.
!endif

########################## OPTIONAL DEFINES #######################
#
# In addition, the following may be *optionally* defined by each module

# Define TWAPI_SRCROOT if module is in a non-standard place in directory tree
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
PACKAGE=$(MODULENAME)
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

!if "x$(EMBED_SCRIPT)" == "xlzma"
OBJDIR = $(OBJDIR)_embed_lzma
!elseif "x$(EMBED_SCRIPT)" == "xplain"
OBJDIR = $(OBJDIR)_embed
!endif

!if "x$(TWAPI_STATIC_BUILD)" != "x"
OBJDIR = $(OBJDIR)_static
!endif

# Make the directory
!if ![if not exist $(OBJDIR) mkdir $(OBJDIR)]
!message Output directory is $(OBJDIR)
!else
!message Output directory $(OBJDIR) exists
!endif

# 
!if "$(MACHINE)" == "AMD64"
MODULE        = $(PACKAGE)64.dll
MODULELIB     = $(PACKAGE)64.lib
!else
MODULE        = $(PACKAGE).dll
MODULELIB     = $(PACKAGE).lib
!endif


################# COMPILER AND SDK #####################

CC            = cl.exe
LINK          = link.exe
LIBTOOL       = lib.exe
MT            = mt.exe

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
!message Compiling with Visual C++ $(VCVER).
!endif

# Compiler flags 

!ifdef DEBUG
CDEBUG = -Z7 -Od -D_DEBUG
!else
!ifdef NOOPTIMIZE
CDEBUG = -Zi -Od
!else
CDEBUG = -O1 -Zi -DNDEBUG -W3
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

CFLAGS    = $(CDEBUG) /c /nologo /DWIN32 /D_WIN32 /D_WINDOWS -DTCL_THREADS=1 -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -DPSAPI_VERSION=1 $(INCFLAGS) $(CONFIGFLAGS) 

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


# Now the SDK

!ifdef MSSDK
SDK           = "$(MSSDK)"
!else
SDK           = "c:\program files\microsoft platform sdk"
!endif

# Windows libraries
# If building with Visual C++ 6.0, we need SDK libraries for newer API's.
# However, we cannot use all SDK libraries since the 6.0 linker does not
# understand some of the newer SDK library format in which case we have
# to stick with the compiler libraries. For newer compilers, we never
# use the SDK libraries.
!if $(VCVER) > 6
SDKLIB =
!else
SDKLIB = $(SDK)\lb\ # This comment to ensure line does not end in backslash
LIB=$(MSVCDIR)\lib
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
TCLDIR        = "$(SRCROOT)\..\tcl-tk\$(TCLVER)"
!endif
TCLLIB        = $(TCLDIR)\lib\Release_$(MACHINE)\tclstub$(TCLVER).lib


# General tools
DEL           = del
COPY          = copy




################################################################
# TWAPI feature options


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



{.}.c{$(OBJDIR)}.obj::
       $(CC) $(CFLAGS) -Fo$(OBJDIR)\ $<
