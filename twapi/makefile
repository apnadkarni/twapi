
# base must be built before everything else as the rest depend on its exports
MODULES = account apputil clipboard com console crypto device \
	etw eventlog input metoo msi mstask multimedia namedpipe network nls \
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
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(MAKEDIR)\dist\$@"
	copy /y >NUL tcl\pkgIndexTemplate.tcl "$(MAKEDIR)\dist\$@\pkgIndex.tcl"

# Build multiple DLLs with embedded scripts
multi_embedded:
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma PACKAGEDIR="$(MAKEDIR)\dist\$@"
	copy /y >NUL tcl\pkgIndexTemplate.tcl "$(MAKEDIR)\dist\$@\pkgIndex.tcl"

single_embedded:
	$(MAKE) /$(MAKEFLAGS) ACTIONS="build distribution" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma TWAPI_SINGLE_MODULE=1 PACKAGEDIR="$(MAKEDIR)\dist\$@"
	copy /y >NUL tcl\pkgIndexTemplate.tcl "$(MAKEDIR)\dist\$@\pkgIndex.tcl"
