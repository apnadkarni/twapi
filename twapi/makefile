SRCROOT=$(MAKEDIR)
DISTDIR=$(MAKEDIR)\dist

# If 1, the generated pkgindex file will do lazy loading
LAZYPACKAGELOAD=1

!include "$(SRCROOT)\include\tooldefs.inc"


# base must be built before everything else as the rest depend on its exports
# Remaining are in order of package dependencies since the build process
# loads them in order when generating the pkgIndex.tcl file.
PURETCLMODULES = metoo msi
MODULES = account apputil clipboard metoo com console crypto device \
	etw eventlog input msi mstask multimedia namedpipe network nls \
	os pdh process rds resource security service share shell storage \
	ui winsta wmi

# Note this can be overridden on the command line
TARGETS = base $(MODULES)
ACTIONS = build

default: $(TARGETS)

base $(MODULES): force
	cd $@ && $(MAKE) /$(MAKEFLAGS) $(ACTIONS) TWAPI_STATIC_BUILD=$(TWAPI_STATIC_BUILD) TWAPI_SINGLE_MODULE=$(TWAPI_SINGLE_MODULE) EMBED_SCRIPT=$(EMBED_SCRIPT) PACKAGEDIR="$(PACKAGEDIR)"

force:

!ifdef TWAPI_SINGLE_MODULE
# When building a single module, all modules are embedded in base so that
# has to be built last
base : $(MODULES)
!else
# When building multiple modules, pure Tcl scripts are embedded in base
# so make sure those are built first
base : $(PURETCLMODULES)
!endif

# Traditional build with single dll and not embedding
traditional: 
!if "$(MACHINE)" != "AMD64"
	-$(DEL) "$(DISTDIR)\$@\pkgindex.modules"
!endif
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(DISTDIR)\$@"
	$(COPY) tcl\pkgIndexTemplate.tcl "$(DISTDIR)\$@\pkgIndex.tcl"
!if "$(MACHINE)" != "AMD64"
	$(TCLSH) "$(TOOLDIR)\makepkgindex.tcl" "$(DISTDIR)\$@" $(LAZYPACKAGELOAD) >> "$(DISTDIR)\$@\pkgIndex.tcl"
	-$(DEL) "$(DISTDIR)\$@\pkgindex.modules"
!endif

# Build multiple DLLs with embedded scripts
multi_embedded:
!if "$(MACHINE)" != "AMD64"
	-$(DEL) "$(DISTDIR)\$@\pkgindex.modules"
!endif
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma PACKAGEDIR="$(DISTDIR)\$@"
	$(COPY) tcl\pkgIndexTemplate.tcl "$(DISTDIR)\$@\pkgIndex.tcl"
!if "$(MACHINE)" != "AMD64"
	$(TCLSH) "$(TOOLDIR)\makepkgindex.tcl" "$(DISTDIR)\$@" $(LAZYPACKAGELOAD) >> "$(DISTDIR)\$@\pkgIndex.tcl"
	-$(DEL) "$(DISTDIR)\$@\pkgindex.modules"
!endif

single_embedded:
!if "$(MACHINE)" != "AMD64"
	-$(DEL) "$(DISTDIR)\$@\pkgindex.modules"
!endif
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(DISTDIR)\$@"
	$(COPY) tcl\pkgIndexTemplate.tcl "$(DISTDIR)\$@\pkgIndex.tcl"
!if "$(MACHINE)" != "AMD64"
	$(TCLSH) "$(TOOLDIR)\makepkgindex.tcl" "$(DISTDIR)\$@" $(LAZYPACKAGELOAD) >> "$(DISTDIR)\$@\pkgIndex.tcl"
	-$(DEL) "$(DISTDIR)\$@\pkgindex.modules"
!endif
