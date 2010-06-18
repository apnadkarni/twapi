TWAPI_SRCROOT = $(MAKEDIR)\..
!include ..\common.inc

!ifndef PACKAGE
!error Macro PACKAGE is not defined.
!endif

!ifndef PACKAGEDIR
PACKAGEDIR=$(TWAPI_SRCROOT)\dist
!endif

TESTFILES=all.tcl \
	testutil.tcl \
	base.test \
	clipboard.test \
	com.test \
	console.test \
	disk.test \
	eventlog.test \
	network.test \
	nls.test \
	osinfo.test \
	parseargs.test \
	process.test \
	security.test \
	services.test \
	share.test \
	shell.test \
	try.test \
	ui.test \
	disk.vbs \
	osinfo.vbs \
	process.vbs \
	services.vbs \
	share.vbs \
	thread.vbs


distribution: $(TESTFILES)
	-mkdir "$(PACKAGEDIR)\tests"
	!copy $** "$(PACKAGEDIR)\tests"

clean:
	-del *.*~
