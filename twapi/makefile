# Master TWAPI makefile

!include common.inc

!ifndef EMBED_SCRIPT
EMBED_SCRIPT=none
!endif

!if "$(MACHINE)" == "AMD64"
DISTDIR=$(MAKEDIR)\temp\embed_$(EMBED_SCRIPT)\win64
DISTBASENAME=$(PACKAGE)-x64-$(MODULEVERSION)
DLLNAME=$(PACKAGE)64
!else
DISTDIR=$(MAKEDIR)\temp\embed_$(EMBED_SCRIPT)\x86
DISTBASENAME=$(PACKAGE)-x86-$(MODULEVERSION)
DLLNAME=$(PACKAGE)
!endif

PACKAGEDIR="$(DISTDIR)\$(PACKAGE)"
STARKITDIR="$(DISTDIR)\$(PACKAGE).vfs"
!if "x$(EMBED_SCRIPT)" == "xnone"
# Zip distro contains both x86 and x64
ZIPBASENAME=$(PACKAGE)-$(MODULEVERSION)
!else
ZIPBASENAME=$(PACKAGE)-$(MODULEVERSION)-bin
!endif

all: base doc

base: dummy
        cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)

doc: dummy
        cd doc && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)

distribution:
        -rmdir/s/q $(PACKAGEDIR)
        -mkdir $(PACKAGEDIR)
        copy LICENSE $(PACKAGEDIR)
	copy doc\announce.txt $(PACKAGEDIR)\RELNOTES.TXT
	cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) PACKAGEDIR=$(PACKAGEDIR) EMBED_SCRIPT=$(EMBED_SCRIPT) distribution
!IF "$(MACHINE)" != "AMD64"
# Also copy the 64bit version
	cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) PACKAGEDIR=$(PACKAGEDIR) EMBED_SCRIPT=$(EMBED_SCRIPT) MACHINE=AMD64 distribution

!ENDIF
!IF "x$(EMBED_SCRIPT)" != "xlzma" && "x$(EMBED_SCRIPT)" != "xplain"
        cd tcl && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) distribution
#        cd doc && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) distribution
        cd tests && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) distribution
!ELSE
	copy << $(PACKAGEDIR)\pkgIndex.tcl
# When script is sourced, the variable $dir must contain the
# full path name of this file's directory.
if {$$::tcl_platform(os) eq "Windows NT" &&
    ($$::tcl_platform(machine) eq "intel" ||
     $$::tcl_platform(machine) eq "amd64") &&
    [string index $$::tcl_platform(osVersion) 0] >= 5} {
    package ifneeded $(PACKAGE) $(MODULEVERSION) [list load [file join $$dir [expr {$$::tcl_platform(machine) eq "amd64" ? "twapi64.dll" : "twapi.dll"}]]]
}
<<KEEP
!ENDIF
        cd $(DISTDIR) && $(ZIP) -r "$(DISTDIR)"/$(ZIPBASENAME).zip $(PACKAGE)
	if exist "%HOME%\My Documents\Downloads" $(COPY) "$(DISTDIR)"\$(ZIPBASENAME).zip "%HOME%\My Documents\Downloads"

#NOTE - the starkit is no longer supported/built regularly
starkit: distribution
        -rmdir/s/q $(STARKITDIR)
        -mkdir $(STARKITDIR)\lib
        -rmdir/s/q $(PACKAGEDIR)\doc
        -rmdir/s/q $(PACKAGEDIR)\tests
        move $(PACKAGEDIR) $(STARKITDIR)\lib
        echo package require starkit > $(STARKITDIR)\main.tcl
        echo starkit::startup >> $(STARKITDIR)\main.tcl
        cd $(DISTDIR) && "$(TOOLDIR)\tclkitsh.exe" "$(TOOLDIR)\sdx.kit" wrap $(PACKAGE).kit
	move "$(DISTDIR)\$(PACKAGE).kit" "$(DISTDIR)\$(DISTBASENAME).kit"

tmdistribution:
        -rmdir/s/q $(PACKAGEDIR)
        -mkdir $(PACKAGEDIR)
        cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) PACKAGEDIR=$(PACKAGEDIR)  EMBED_SCRIPT=$(EMBED_SCRIPT) distribution
!IF "x$(EMBED_SCRIPT)" != "xlzma" && "x$(EMBED_SCRIPT)" != "xplain"
        cd tcl && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) PACKAGESTYLE=tm distribution
	$(TCLSH) "$(TOOLDIR)\commentify.tcl" LICENSE $(PACKAGEDIR)\LICENSE
        $(TCLSH) "$(TOOLDIR)\createtmfile.tcl" -compact $(PACKAGE) $(MODULEVERSION) $(PACKAGEDIR)\LICENSE $(PACKAGEDIR)\twapi_version.tcl $(PACKAGEDIR)\twapi_buildid.tcl $(PACKAGEDIR)\win32calls.tcl $(PACKAGEDIR)\twapi.tcl $(PACKAGEDIR)\*.tcl $(PACKAGEDIR)\$(DLLNAME).dll
!ELSE
        cd tcl && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) PACKAGESTYLE=tm twapi_buildid.tcl
	$(TCLSH) "$(TOOLDIR)\commentify.tcl" LICENSE $(PACKAGEDIR)\LICENSE
        $(TCLSH) "$(TOOLDIR)\createtmfile.tcl" -autoload $(PACKAGE) $(MODULEVERSION) $(PACKAGEDIR)\LICENSE $(PACKAGEDIR)\twapi_buildid.tcl $(PACKAGEDIR)\$(DLLNAME).dll
!ENDIF
        move $(PACKAGE)-$(MODULEVERSION).tm "$(DISTDIR)"\$(DISTBASENAME).tm
	if exist "%HOME%\My Documents\Downloads" $(COPY) "$(DISTDIR)"\$(DISTBASENAME).tm "%HOME%\My Documents\Downloads"

!IF "x$(EMBED_SCRIPT)" == "xlzma" || "x$(EMBED_SCRIPT)" == "xplain"
dlldistribution:
        -rmdir/s/q $(PACKAGEDIR)
        -mkdir $(PACKAGEDIR)
	cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) PACKAGEDIR=$(PACKAGEDIR)  EMBED_SCRIPT=$(EMBED_SCRIPT) distribution
        move $(PACKAGEDIR)\$(DLLNAME).dll "$(DISTDIR)"\$(DISTBASENAME).dll
	if exist "%HOME%\My Documents\Downloads" $(COPY) "$(DISTDIR)"\$(DISTBASENAME).dll "%HOME%\My Documents\Downloads"
!ENDIF


install:
	$(TCLSH_TARGET) <<
exec unzip -o [file join {$(DISTDIR)} $(ZIPBASENAME).zip] -d [lindex [set auto_path] 0]
<<keep

website:
        cd doc && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  website

makefile: common.inc

dummy:

clean:
        cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) clean
        cd doc && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) clean
        cd tcl && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) clean
        -rmdir/s/q "$(DISTDIR)"
