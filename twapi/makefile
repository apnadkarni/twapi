SRCROOT=$(MAKEDIR)
!include "$(SRCROOT)\include\tooldefs.inc"


# base must be built before everything else as the rest depend on its exports
# Remaining are in order of package dependencies since the build process
# loads them in order when generating the pkgIndex.tcl file.
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
base : $(MODULES)
!endif

# Traditional build with single dll and not embedding
traditional:
!if "$(MACHINE)" != "AMD64"
	-$(DEL) "$(MAKEDIR)\dist\$@\pkgindex.modules"
!endif
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(MAKEDIR)\dist\$@"
	$(COPY) tcl\pkgIndexTemplate.tcl "$(MAKEDIR)\dist\$@\pkgIndex.tcl"
!if "$(MACHINE)" != "AMD64"
	$(TCLSH) "$(TOOLDIR)\makepkgindex.tcl" "$(MAKEDIR)\dist\$@" >> "$(MAKEDIR)\dist\$@\pkgIndex.tcl"
	-$(DEL) "$(MAKEDIR)\dist\$@\pkgindex.modules"
!endif

# Build multiple DLLs with embedded scripts
multi_embedded:
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma PACKAGEDIR="$(MAKEDIR)\dist\$@"
	copy /y >NUL tcl\pkgIndexTemplate.tcl "$(MAKEDIR)\dist\$@\pkgIndex.tcl"

single_embedded:
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(MAKEDIR)\dist\$@"
	copy /y >NUL tcl\pkgIndexTemplate.tcl "$(MAKEDIR)\dist\$@\pkgIndex.tcl"
