# Include file for building all modules
# The defines assume nmake is executed in the directory of the module

!include version.inc
!include tooldefs.inc

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

# Release type module defaults to TWAPI release type defined in version.inc
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

# If the module has dependencies on any other TWAPI modules,
# define this in the module makefile. No need to include twapi_base
# as it is implied
# e.g. PREREQUISITES=twapi_com

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

# Top level directory where distributions are written
!ifndef DISTDIR
DISTDIR = $(SRCROOT)\dist
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
# EMBED_SCRIPT

!ifdef TWAPI_SINGLE_MODULE
!if "$(TWAPI_SINGLE_MODULE)" != "1"
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


# Distribution directories
!if "$(PACKAGEDIR)" == ""

! if "$(EMBED_SCRIPT)" == ""
PACKAGEDIR = $(DISTDIR)\noembed
! else
PACKAGEDIR = $(DISTDIR)\embed_$(EMBED_SCRIPT)
! endif

! ifdef TWAPI_SINGLE_MODULE
PACKAGEDIR = $(PACKAGEDIR)_single
! endif TWAPI_SINGLE_MODULE

!endif $(PACKAGEDIR) == ""

# 
!if "$(MACHINE)" == "AMD64"
MODULE        = $(PACKAGE)64.dll
MODULELIB     = $(PACKAGE)64.lib
!else
MODULE        = $(PACKAGE).dll
MODULELIB     = $(PACKAGE).lib
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

INCFLAGS       =  -I"$(OBJDIR)" -I. -I..\include -I$(TCLDIR)\include

CINCLUDES     = $(SRCROOT)\include\twapi.h \
		$(SRCROOT)\include\twapi_wm.h \
		$(SRCROOT)\include\twapi_ddkdefs.h \
		$(SRCROOT)\include\twapi_sdkdefs.h \
		$(SRCROOT)\include\zlist.h \
		$(SRCROOT)\include\memlifo.h

CFLAGS    = $(CDEBUG) /c /nologo /DWIN32 /D_WIN32 /D_WINDOWS -DTCL_THREADS=1 -D_WIN32_WINNT=$(TWAPI_WIN_HEADER_VERSION) -DPSAPI_VERSION=1 $(INCFLAGS) -DMODULENAME=\"$(MODULENAME)\" -D$(MODULENAME)_BUILD -DMODULEVERSION=\"$(MODULEVERSION)\"

!if "$(MACHINE)" == "AMD64" && $(VCVER) < 9
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

# TBD - remove this at some point once unsafe C RTL functions are removed
CFLAGS = $(CFLAGS) /D_CRT_SECURE_NO_WARNINGS

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
! if "$(MACHINE)" == "AMD64"
LIBS = "$(OBJDIR)\..\twapi_base\twapi_base64.lib" $(LIBS)
! else
LIBS = "$(OBJDIR)\..\twapi_base\twapi_base.lib" $(LIBS)
! endif
!endif

# Resource definitions

!ifndef RC_FILE_PATH
RC_FILE_PATH = "$(SRCROOT)\include\common.rc"
!endif

RCFLAGS = -DMODULEMAJOR=$(MODULEMAJOR) -DMODULEMINOR=$(MODULEMINOR) -DMODULEBUILD=$(MODULEBUILD) -DFILE_VERSION_STR=\"$(FILE_VERSION_STR)\" -DPRODUCT_VERSION_STR=\"$(MODULE_VERSION_STR)\" -DMODULE_STR=\"$(MODULE)\" -DMODULENAME=$(MODULENAME)

# If building a lib, default to lzma 
!if defined(TWAPI_STATIC_BUILD) && ! defined(EMBED_SCRIPT)
EMBED_SCRIPT=lzma
!endif

TWAPI_SCRIPT_RESOURCE_DIR=$(OBJDIR)\tcl

!if "$(EMBED_SCRIPT)" == "" || "$(EMBED_SCRIPT)" == "none"

# Not embedding the scripts themselves. Embed a stub that will source the files
TWAPI_SCRIPT_RESOURCE_TYPE=tclscript
TWAPI_SCRIPT_RESOURCE=$(MODULENAME)_source_stub

# Undefine to indicate actual scripts are not embedded
!  undef EMBED_SCRIPT

!elseif "$(EMBED_SCRIPT)" == "lzma"

TWAPI_SCRIPT_RESOURCE_TYPE=tclscriptlzma
TWAPI_SCRIPT_RESOURCE=combined_$(MODULENAME)_script.lzma

!  elseif "$(EMBED_SCRIPT)" == "plain"

TWAPI_SCRIPT_RESOURCE_TYPE=tclscript
TWAPI_SCRIPT_RESOURCE=combined_$(MODULENAME)_script

!  else

!    error EMBED_SCRIPT must be undefined, "none", "lzma" or "plain"

!  endif

# Note: Do NOT enclose TWAPI_SCRIPT_RESOURCE_PATH in quotes even though
# it is a file name, not a path because when passed to the rc compiler
# as an include, it should not have quotes
TWAPI_SCRIPT_RESOURCE_PATH=$(TWAPI_SCRIPT_RESOURCE_DIR)\$(TWAPI_SCRIPT_RESOURCE)

# Linking
# -------
LOPT      = /RELEASE /NOLOGO /MACHINE:$(MACHINE) /DLL /OPT:REF $(LENTRYFLAG)

!ifdef ENABLE_DELAYLOAD
LIBS = $(LIBS) delayimp.lib
LOPT = $(LOPT) /DELAYLOAD:netapi32.dll /DELAYLOAD:mpr.dll  /DELAYLOAD:powrprof.dll /DELAYLOAD:userenv.dll /DELAYLOAD:psapi.dll /DELAYLOAD:wtsapi32.dll /DELAYLOAD:pdh.dll /DELAYLOAD:setupapi.dll /DELAYLOAD:iphlpapi.dll /DELAYLOAD:winmm.dll
!endif

!if $(VCVER) == 6
# VC 6 - /OPT:NOWIN98 reduces code size by using 512 byte section alignment
LOPT = $(LOPT) /debug:full /debugtype:cv /OPT:NOWIN98
!else
LOPT = $(LOPT) /debug
!endif






