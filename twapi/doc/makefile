# Documentation generator
#
# To build CHM help file (note this randomly crashes the MS help compiler)
#   nmake
#
# To build website
#   nmake ADFILE=google-ads.js ADFILE2=google-ads-horizontal-link-unit.js website

TWAPI_SRCROOT = $(MAKEDIR)\..
!include ../common.inc

# Comment line below to see command lines being executed
# Or just invoke nmake with SILENT=
SILENT=@
!if "x$(SILENT)" == "x@"
DTPQUIET = -quiet
!endif

!ifndef TOOLDIR
TOOLDIR= $(TWAPI_SRCROOT)\..\tools
!endif

TIDY="$(TOOLDIR)\tidy.exe"
DTP="$(TOOLDIR)\twapi-dtp.kit"
!ifndef HHC
HHC="$(PROGRAMFILES)\HTML Help Workshop\hhc.exe"
!endif

# Documentation input files
!ifndef MANFILES
MANFILES= \
	overview.man \
	packages.man \
	installation.man \
	raw.man \
	ui.man \
	desktop.man \
	com.man \
	shell.man \
	input.man \
	sound.man \
	clipboard.man \
	osinfo.man \
	process.man \
	handles.man \
	services.man \
	network.man \
	disk.man \
	share.man \
	users.man \
	security.man \
	crypto.man \
	eventlog.man \
	inifile.man \
	console.man \
	schedule.man \
	msi.man \
	nls.man \
	printer.man \
	device.man \
	power.man \
	errors.man \
	misc.man \
	examples.man \
	versionhistory.man 
!endif

IMGFILES = twapi-32.png twapi-48.png magicsplat.png navarrow.png download.png download_active.png


!ifndef FMT
FMT=htm
!endif

HTMLOUTPUTFILES=$(MANFILES:.man=.html)
HTFOUTPUTFILES=$(MANFILES:.man=.htf)

all: twapi.chm

twapi.chm : $(MANFILES) styleschm.css makefile html_head.inc generate_mshelp_hhp.tcl generate_mshelp_index.tcl generate_mshelp_contents.tcl
	@echo =========Building CHM file...
	@echo Making directories...
	-$(SILENT)if not exist chm-build mkdir chm-build
	@echo Copying man files
	$(SILENT)xcopy /q /y *.man chm-build
	@echo Processing man files...
	$(SILENT)cd chm-build && $(SILENT)$(TCLSH) $(DTP) map -ext html $(MANFILES) > iomap.tmp
	$(SILENT)cd chm-build && $(SILENT)$(TCLSH) $(DTP) meta iomap.tmp > meta.tmp
	$(SILENT)cd chm-build && $(SILENT)$(TCLSH) $(DTP) idx -title "Tcl Windows API Extension - Index" meta.tmp > idx.tmp
	$(SILENT)cd chm-build && $(SILENT)$(TCLSH) $(DTP) gen-doc $(DTPQUIET) -varfile meta ../html_head.inc $(FMT) iomap.tmp
	@echo Generating MSHELP input files...
	$(SILENT)cd chm-build && $(SILENT)$(TCLSH) ../generate_mshelp_hhp.tcl $(HTMLOUTPUTFILES) styleschm.css > twapi.hhp
	$(SILENT)cd chm-build && $(SILENT)$(TCLSH) ../generate_mshelp_index.tcl idx.tmp > twapi.hhk
	$(SILENT)cd chm-build && $(SILENT)$(TCLSH) ../generate_mshelp_contents.tcl meta.tmp > twapi.hhc
	$(SILENT)xcopy /q /y styleschm.css chm-build
	$(SILENT)for %f in ($(IMGFILES)) do $(SILENT)xcopy /q /y %f chm-build
	$(SILENT)xcopy /q /y *.example chm-build
# HHC always seems to return exit code 1!!
	@echo Generating CHM file...
	-$(SILENT)cd chm-build && $(SILENT)$(HHC) twapi.hhp

htmdistribution: readme.txt twapi.chm
	-if not exist "$(PACKAGEDIR)\doc" mkdir "$(PACKAGEDIR)\doc"
	!copy /y $? "$(PACKAGEDIR)\doc"

distribution: $(FMT)distribution

website: website.zip

website.zip : $(MANFILES) styles.css index.htf generate_web_page.tcl generate_www_index.tcl makefile
	@echo =========Building Web site...
	@echo Making directories...
	-$(SILENT)if not exist website mkdir website
	-$(SILENT)if not exist website-build mkdir website-build
	@echo Copying man files
	$(SILENT)xcopy /q /y *.man website-build
	@echo Processing man files...
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) map -ext htf $(MANFILES) idx.man > iomap.tmp
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) meta iomap.tmp > meta.tmp
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) toc -title "Tcl Windows API Extension - Table of Contents"  meta.tmp > toc.tmp
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) idx -title "Tcl Windows API Extension - Index" meta.tmp > idx.tmp
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) gen-doc $(DTPQUIET) htf iomap.tmp
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) gen-toc htf iomap.tmp toc.tmp > toc.htf
	@echo Generating web pages...
	$(SILENT)$(TCLSH) generate_www_index.tcl website-build\idx.tmp > website-build\idx.htf
        $(SILENT)for %f in ($(HTFOUTPUTFILES) idx.htf) do $(SILENT)$(TCLSH) generate_web_page.tcl website-build\%f $(MODULEVERSION) website-build\toc.htf $(ADFILE) $(ADFILE2) > website-build\%~nf.html
	@echo Tidying web pages...
        $(SILENT)for %f in ($(HTMLOUTPUTFILES) idx.html) do $(SILENT)$(TIDY) -asxhtml -q website-build\%f > website\%f
        $(SILENT)$(TCLSH) generate_web_page.tcl index.htf $(MODULEVERSION) website-build\toc.htf $(ADFILE) $(ADFILE2) | $(TIDY) -asxhtml -q > website\index.html
	@echo Copying files...
	$(SILENT)xcopy /q /y styles.css website
	$(SILENT)for %f in ($(IMGFILES)) do $(SILENT)xcopy /q /y %f website
	$(SILENT)xcopy /q /y *.example website
	@echo Zipping distribution...
	$(SILENT)zip -q -r twapi-website.zip website

clean:
	-del *.tmp
	-del *.*~
	-del idx.html
	-del toc.html
	-del $(HTMLOUTPUTFILES)
	-rmdir/s/q website-build
	-rmdir/s/q website
	-del *.chm
	-del *.hhc
	-del *.hhk
	-del twapi.hhp
