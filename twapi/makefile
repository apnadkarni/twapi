
# base must be built before everything else as the rest depend on its exports
MODULES = account apputil clipboard com console crypto device \
	etw eventlog input mstask multimedia namedpipe network nls \
	os pdh process resource security service share shell storage \
	ui winsta

# Note this can be overridden on the command line
TARGETS = base $(MODULES)
ACTIONS = build

build: $(TARGETS)

base $(MODULES): force
	cd $@ && $(MAKE) /$(MAKEFLAGS) $(ACTIONS) TWAPI_STATIC_BUILD=$(TWAPI_STATIC_BUILD)

force:
