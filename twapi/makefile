
# base must be built before everything else as the rest depend on its exports
MODULES = account apputil clipboard com console crypto device \
	etw eventlog input mstask multimedia namedpipe network nls \
	os pdh process rds resource security service share shell storage \
	ui winsta

# Note this can be overridden on the command line
TARGETS = base $(MODULES)
ACTIONS = build

default: $(ACTIONS)

build: $(TARGETS)

base $(MODULES): force
	cd $@ && $(MAKE) /$(MAKEFLAGS) $(ACTIONS) TWAPI_STATIC_BUILD=$(TWAPI_STATIC_BUILD) TWAPI_SINGLE_MODULE=$(TWAPI_SINGLE_MODULE) EMBED_SCRIPT=$(EMBED_SCRIPT)

force:

!ifdef TWAPI_SINGLE_MODULE
base : $(MODULES)
!endif

# Build multiple DLLs with embedded scripts
multi_embedded: force
	cd $@ && $(MAKE) /$(MAKEFLAGS) $(ACTIONS)

single_embedded:
	$(MAKE) /$(MAKEFLAGS) ACTIONS="$(ACTIONS)" TWAPI_SINGLE_MODULE=1 EMBED_SCRIPT=lzma TARGETS="$(TARGETS)"
