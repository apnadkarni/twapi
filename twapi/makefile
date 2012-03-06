
# base must be built before everything else as the rest depend on its exports
MODULES = account apputil clipboard com console crypto device \
	etw eventlog input msi mstask multimedia namedpipe network nls \
	os pdh process rds resource security service share shell storage \
	ui winsta wmi

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

# Traditional build with single dll and not embedding
traditional: 
	$(MAKE) /$(MAKEFLAGS) ACTIONS="$(ACTIONS)" TARGETS="$(TARGETS)" TWAPI_SINGLE_MODULE=1

# Build multiple DLLs with embedded scripts
multi_embedded:
	$(MAKE) /$(MAKEFLAGS) ACTIONS="$(ACTIONS)" TARGETS="$(TARGETS)"

single_embedded:
	$(MAKE) /$(MAKEFLAGS) ACTIONS="$(ACTIONS)" TARGETS="$(TARGETS)" EMBED_SCRIPT=lzma TWAPI_SINGLE_MODULE=1 
