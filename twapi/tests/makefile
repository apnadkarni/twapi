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
	device.test \
	disk.test \
	eventlog.test \
	namedpipe.test \
	network.test \
	nls.test \
	osinfo.test \
	parseargs.test \
	pdh.test \
	power.test \
	process.test \
	resource.test \
	security.test \
	services.test \
	share.test \
	shell.test \
	try.test \
	ui.test \
	mstask.test \
	npecho.tcl \
	echoservice.tcl \
	client_share.vbs \
	disk.vbs \
	osinfo.vbs \
	process.vbs \
	services.vbs \
	share.vbs \
	thread.vbs \
	rctest.dll

distribution: $(TESTFILES)
	-mkdir "$(PACKAGEDIR)\tests\rctest"
	!copy $** "$(PACKAGEDIR)\tests"

clean:
	-del *.*~
