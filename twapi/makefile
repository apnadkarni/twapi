SRCROOT=$(MAKEDIR)
DISTDIR=$(MAKEDIR)\dist

# If 1, the generated pkgindex file will do lazy loading
LAZYPACKAGELOAD=1

!include "$(SRCROOT)\include\version.inc"
!include "$(SRCROOT)\include\tooldefs.inc"

# base must be built before everything else as the rest depend on its exports
# Dependencies at load time are handled via the PREREQUISITES macro (see 
# common.inc)
PURETCLMODULES = metoo msi power printer synch
BINMODULES = account apputil clipboard com console crypto device \
	etw eventlog mstask multimedia namedpipe network nls \
	os pdh process rds resource security service share shell \
	storage ui input winsta wmi
MODULES= $(PURETCLMODULES) $(BINMODULES)

# Note this can be overridden on the command line
TARGETS = base $(MODULES)
ACTIONS = build

default: $(TARGETS)

base $(MODULES): force
	echo Making $@
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
twapi-build: 
!if "$(MACHINE)" != "AMD64"
	-$(DEL) "$(DISTDIR)\twapi\pkgindex.modules"
!endif
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(DISTDIR)\twapi"

twapi-bin-build: 
!if "$(MACHINE)" != "AMD64"
	-$(DEL) "$(DISTDIR)\twapi-bin\pkgindex.modules"
!endif
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(DISTDIR)\twapi-bin"

twapi-modular-build: 
!if "$(MACHINE)" != "AMD64"
	-$(DEL) "$(DISTDIR)\twapi-modular\pkgindex.modules"
!endif
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma PACKAGEDIR="$(DISTDIR)\twapi-modular"


"$(DISTDIR)\twapi\pkgIndex.tcl" "$(DISTDIR)\twapi-bin\pkgIndex.tcl" "$(DISTDIR)\twapi-modular\pkgIndex.tcl": force
	$(COPY) tcl\pkgIndexTemplate.tcl $@
	$(TCLSH) "$(TOOLDIR)\makepkgindex.tcl" $(@D) $(LAZYPACKAGELOAD) >> $@
	-$(DEL) $(@D)\pkgindex.modules

!if "$(MACHINE)" == "AMD64"
twapi twapi-bin twapi-modular: $(@)-build
!else
twapi twapi-bin twapi-modular: $(@)-build "$(DISTDIR)\$@\pkgindex.tcl"
        $(COPY) LICENSE "$(DISTDIR)\$@"
	$(COPY) doc\announce.txt "$(DISTDIR)\$@\RELNOTES.TXT"
	cd "$(DISTDIR)" && $(ZIP) -r "$@-$(TWAPI_VERSION).zip" $@
	if exist "$(USERPROFILE)\My Documents\Downloads\twapi" $(COPY) "$(DISTDIR)\$@-$(TWAPI_VERSION).zip" "$(USERPROFILE)\My Documents\Downloads\twapi"

!endif

