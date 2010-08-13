# Master TWAPI makefile

!include common.inc

!if "$(MACHINE)" == "AMD64"
DISTDIR=$(MAKEDIR)\temp\win64
DISTBASENAME=$(PACKAGE)-x64-$(MODULEVERSION)
!else
DISTDIR=$(MAKEDIR)\temp
DISTBASENAME=$(PACKAGE)-x86-$(MODULEVERSION)
!endif

PACKAGEDIR="$(DISTDIR)\$(PACKAGE)"
DISTFILE=$(DISTBASENAME).zip
STARKITDIR="$(DISTDIR)\$(PACKAGE).vfs"

!ifndef EMBED_SCRIPT
EMBED_SCRIPT=no
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
	cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) PACKAGEDIR=$(PACKAGEDIR) EMBED_SCRIPT=$(EMBED_SCRIPT) distribution
!IF "x$(EMBED_SCRIPT)" != "xlzma" && "x$(EMBED_SCRIPT)" != "xplain"
        cd tcl && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) distribution
!ENDIF
!IFNDEF LEAN
        cd tests && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) distribution
        cd doc && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) distribution
!ENDIF
        cd $(DISTDIR) && zip -r "$(DISTDIR)/$(DISTFILE)" $(PACKAGE)
	if exist "%HOME%\My Documents\Downloads" $(COPY) "$(DISTDIR)\$(DISTFILE)" "%HOME%\My Documents\Downloads"

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
        $(TCLSH) "$(TOOLDIR)\createtmfile.tcl" -compact $(PACKAGE) $(MODULEVERSION) $(PACKAGEDIR)\LICENSE $(PACKAGEDIR)\twapi_version.tcl $(PACKAGEDIR)\twapi_buildid.tcl $(PACKAGEDIR)\win32calls.tcl $(PACKAGEDIR)\twapi.tcl $(PACKAGEDIR)\*.tcl $(PACKAGEDIR)\$(PACKAGE).dll
!ELSE
        cd tcl && $(MAKE) MAKEFLAGS=$(MAKEFLAGS)  PACKAGEDIR=$(PACKAGEDIR) PACKAGESTYLE=tm twapi_buildid.tcl
	$(TCLSH) "$(TOOLDIR)\commentify.tcl" LICENSE $(PACKAGEDIR)\LICENSE
        $(TCLSH) "$(TOOLDIR)\createtmfile.tcl" -autoload $(PACKAGE) $(MODULEVERSION) $(PACKAGEDIR)\LICENSE $(PACKAGEDIR)\twapi_buildid.tcl $(PACKAGEDIR)\$(PACKAGE).dll
!ENDIF
        move $(PACKAGE)-$(MODULEVERSION).tm "$(DISTDIR)"\$(DISTBASENAME).tm

!IF "x$(EMBED_SCRIPT)" == "xlzma" || "x$(EMBED_SCRIPT)" == "xplain"
dlldistribution:
        -rmdir/s/q $(PACKAGEDIR)
        -mkdir $(PACKAGEDIR)
	cd base && $(MAKE) MAKEFLAGS=$(MAKEFLAGS) PACKAGEDIR=$(PACKAGEDIR)  EMBED_SCRIPT=$(EMBED_SCRIPT) distribution
        move $(PACKAGEDIR)\$(PACKAGE).dll "$(DISTDIR)"\$(DISTBASENAME).dll
!ENDIF


install:
	$(TCLSH_TARGET) <<
exec unzip -o [file join {$(DISTDIR)} $(DISTFILE)] -d [lindex [set auto_path] 0]
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
