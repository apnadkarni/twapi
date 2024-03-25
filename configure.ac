#!/bin/bash -norc
dnl	This file is an input file used by the GNU "autoconf" program to
dnl	generate the file "configure", which is run during Tcl installation
dnl	to configure the system for the local environment.

#-----------------------------------------------------------------------
# Sample configure.in for Tcl Extensions.  The only places you should
# need to modify this file are marked by the string __CHANGE__
#-----------------------------------------------------------------------

#-----------------------------------------------------------------------
# This initializes the environment with PACKAGE_NAME and PACKAGE_VERSION
# set as provided.  These will also be added as -D defs in your Makefile
# so you can encode the package version directly into the source files.
# This will also define a special symbol for Windows (BUILD_sample in
# this case) so that we create the export library with the dll.
#-----------------------------------------------------------------------

AC_INIT([twapi],[5.0a0])

#--------------------------------------------------------------------
# Call TEA_INIT as the first TEA_ macro to set up initial vars.
# This will define a ${TEA_PLATFORM} variable == "unix" or "windows"
# as well as PKG_LIB_FILE and PKG_STUB_LIB_FILE.
# --------------------------------------------------------------------

TEA_INIT([3.13])

AC_CONFIG_AUX_DIR(tclconfig)

#--------------------------------------------------------------------
# Load the tclConfig.sh file
#--------------------------------------------------------------------

TEA_PATH_TCLCONFIG
TEA_LOAD_TCLCONFIG

#--------------------------------------------------------------------
# Load the tkConfig.sh file if necessary (Tk extension)
#--------------------------------------------------------------------

#TEA_PATH_TKCONFIG
#TEA_LOAD_TKCONFIG

#-----------------------------------------------------------------------
# Handle the --prefix=... option by defaulting to what Tcl gave.
# Must be called after TEA_LOAD_TCLCONFIG and before TEA_SETUP_COMPILER.
#-----------------------------------------------------------------------

TEA_PREFIX

#-----------------------------------------------------------------------
# Standard compiler checks.
# This sets up CC by using the CC env var, or looks for gcc otherwise.
# This also calls AC_PROG_CC, AC_PROG_INSTALL and a few others to create
# the basic setup necessary to compile executables.
#-----------------------------------------------------------------------

TEA_SETUP_COMPILER

#-----------------------------------------------------------------------
# __CHANGE__
# Specify the C source files to compile in TEA_ADD_SOURCES,
# public headers that need to be installed in TEA_ADD_HEADERS,
# stub library C source files to compile in TEA_ADD_STUB_SOURCES,
# and runtime Tcl library files in TEA_ADD_TCL_SOURCES.
# This defines PKG(_STUB)_SOURCES, PKG(_STUB)_OBJECTS, PKG_HEADERS
# and PKG_TCL_SOURCES.
#-----------------------------------------------------------------------

TEA_ADD_LIBS([])
TEA_ADD_STUB_SOURCES([])

#--------------------------------------------------------------------
# You can add more files to clean if your extension creates any extra
# files by extending CLEANFILES.
# Add pkgIndex.tcl if it is generated in the Makefile instead of ./configure
# and change Makefile.in to move it from CONFIG_CLEAN_FILES to BINARIES var.
#
# A few miscellaneous platform-specific items:
# TEA_ADD_* any platform specific compiler/build info here.
#--------------------------------------------------------------------

# CLEANFILES="$CLEANFILES yourfile ..."

if test "${TEA_PLATFORM}" = "windows" ; then

    AC_CHECK_MEMBER([struct _SYSTEM_POWER_STATUS.SystemStatusFlag], [
                        AC_DEFINE(HAVE_SYSTEM_STATUS_FLAG)
                    ], [], [[#include <windows.h>]])

    TEA_ADD_INCLUDES([-I. -I\"$(${CYGPATH} ${srcdir}/twapi/include)\" -I\"${srcdir}/dyncall/dyncall-0.9/include\"])

    TEA_ADD_CFLAGS([-DUNICODE -D_UNICODE -DWINVER=0x0601 -D_WIN32_WINNT=0x0601 -DPSAPI_VERSION=1])

    TEA_ADD_SOURCES([
	    win/adsi.c
	    win/async.c
	    win/calls.c
	    win/errors.c
	    win/ffi.c
	    win/keylist.c
	    win/lzmadec.c
	    win/lzmainterface.c
	    win/memlifo.c
	    win/mycrt.c
	    win/parseargs.c
	    win/printer.c
	    win/recordarray.c
	    win/tclobjs.c
	    win/threadpool.c
	    win/trap.c
	    win/twapi.c
	    win/twine.c
	    win/util.c
	    win/win.c
	    win/winchars.c
	    win/account.c
	    win/apputil.c
	    win/clipboard.c
	    win/com.c
	    win/comserver.c
	    win/console.c
	    win/crypto.c
	    win/sspi.c
	    win/pbkdf2.c
	    win/device.c
	    win/etw.c
	    win/eventlog.c
	    win/evt.c
	    win/input.c
	    win/mstask.c
	    win/multimedia.c
	    win/namedpipe.c
	    win/network.c
	    win/nls.c
	    win/os.c
	    win/pdh.c
	    win/process.c
	    win/rds.c
            win/registry.c
	    win/resource.c
	    win/security.c
	    win/service.c
	    win/scm.c
	    win/share.c
	    win/shell.c
	    win/storage.c
	    win/dirmonitor.c
	    win/ui.c
	    win/gdi.c
	    win/winsta.c
	    win/wmi.c
    ])

    TEA_ADD_HEADERS([
            win/twapi.h
	    win/twapi_ddkdefs.h
	    win/twapi_sdkdefs.h
	    win/zlist.h
	    win/memlifo.h
    ])

    TEA_ADD_LIBS([
		    kernel32.lib advapi32.lib user32.lib
		    rpcrt4.lib gdi32.lib psapi.lib netapi32.lib
		    pdh.lib winmm.lib mpr.lib ws2_32.lib uuid.lib
		    ole32.lib oleaut32.lib shell32.lib shlwapi.lib
		    winspool.lib version.lib iphlpapi.lib
		    powrprof.lib secur32.lib userenv.lib
		    wtsapi32.lib setupapi.lib uxtheme.lib
		    credui.lib crypt32.lib wintrust.lib tdh.lib
		])

    TEA_ADD_TCL_SOURCES([
	    library/twapi.tcl
	    library/base.tcl
	    library/handle.tcl
	    library/win.tcl
	    library/adsi.tcl
	    library/account.tcl
	    library/apputil.tcl
	    library/clipboard.tcl
	    library/com.tcl
	    library/console.tcl
	    library/crypto.tcl
	    library/sspi.tcl
	    library/tls.tcl
	    library/device.tcl
	    library/etw.tcl
	    library/eventlog.tcl
	    library/evt.tcl
	    library/winlog.tcl
	    library/input.tcl
	    library/msi.tcl
	    library/mstask.tcl
	    library/multimedia.tcl
	    library/namedpipe.tcl
	    library/network.tcl
	    library/nls.tcl
	    library/os.tcl
	    library/pdh.tcl
	    library/power.tcl
	    library/printer.tcl
	    library/process.tcl
	    library/rds.tcl
	    library/registry.tcl
	    library/resource.tcl
	    library/security.tcl
	    library/service.tcl
	    library/share.tcl
	    library/shell.tcl
	    library/storage.tcl
	    library/synch.tcl
	    library/ui.tcl
	    library/winsta.tcl
	    library/wmi.tcl
	    LICENSE
            README.md
    ])

else
    # Ensure no empty else clauses
    :
    #TEA_ADD_SOURCES([unix/unixFile.c])
    #TEA_ADD_LIBS([-lsuperfly])
fi

#--------------------------------------------------------------------
# Choose which headers you need.  Extension authors should try very
# hard to only rely on the Tcl public header files.  Internal headers
# contain private data structures and are subject to change without
# notice.
# This MUST be called after TEA_LOAD_TCLCONFIG / TEA_LOAD_TKCONFIG
#--------------------------------------------------------------------

TEA_PUBLIC_TCL_HEADERS
#TEA_PRIVATE_TCL_HEADERS

#TEA_PUBLIC_TK_HEADERS
#TEA_PRIVATE_TK_HEADERS
#TEA_PATH_X

#--------------------------------------------------------------------
# Check whether --enable-threads or --disable-threads was given.
# This auto-enables if Tcl was compiled threaded.
#--------------------------------------------------------------------

TEA_ENABLE_THREADS

#--------------------------------------------------------------------
# The statement below defines a collection of symbols related to
# building as a shared library instead of a static library.
#--------------------------------------------------------------------

TEA_ENABLE_SHARED

#--------------------------------------------------------------------
# This macro figures out what flags to use with the compiler/linker
# when building shared/static debug/optimized objects.  This information
# can be taken from the tclConfig.sh file, but this figures it all out.
#--------------------------------------------------------------------

TEA_CONFIG_CFLAGS

# Link to either the 64- or 32-bit dyncall libs. This check has to happen
# after TEA_CONFIG_CFLAGS because that sets the do64bit variable
AS_IF([test "$do64bit" = "yes"], [
	  AC_SUBST(DYNCALL_LIBS, "${srcdir}/dyncall/dyncall-0.9/lib/release_amd64/libdyncall_s.lib")
      ], [
	  AC_SUBST(DYNCALL_LIBS, "${srcdir}/dyncall/dyncall-0.9/lib/release_x86/libdyncall_s.lib")
      ])

#--------------------------------------------------------------------
# Set the default compiler switches based on the --enable-symbols option.
#--------------------------------------------------------------------

TEA_ENABLE_SYMBOLS

#--------------------------------------------------------------------
# Everyone should be linking against the Tcl stub library.  If you
# can't for some reason, remove this definition.  If you aren't using
# stubs, you also need to modify the SHLIB_LD_LIBS setting below to
# link against the non-stubbed Tcl library.  Add Tk too if necessary.
#--------------------------------------------------------------------

AC_DEFINE(USE_TCL_STUBS, 1, [Use Tcl stubs])
#AC_DEFINE(USE_TK_STUBS, 1, [Use Tk stubs])

#--------------------------------------------------------------------
# This macro generates a line to use when building a library.  It
# depends on values set by the TEA_ENABLE_SHARED, TEA_ENABLE_SYMBOLS,
# and TEA_LOAD_TCLCONFIG macros above.
#--------------------------------------------------------------------

TEA_MAKE_LIB

#--------------------------------------------------------------------
# Determine the name of the tclsh and/or wish executables in the
# Tcl and Tk build directories or the location they were installed
# into. These paths are used to support running test cases only,
# the Makefile should not be making use of these paths to generate
# a pkgIndex.tcl file or anything else at extension build time.
#--------------------------------------------------------------------

TEA_PROG_TCLSH
#TEA_PROG_WISH

#--------------------------------------------------------------------
# Finally, substitute all of the various values into the Makefile.
# You may alternatively have a special pkgIndex.tcl.in or other files
# which require substituting th AC variables in.  Include these here.
#--------------------------------------------------------------------

AC_CONFIG_FILES([Makefile pkgIndex.tcl])
AC_OUTPUT
