# Documentation generator
#
# To build CHM help file (note this randomly crashes the MS help compiler)
#   nmake
#
# To build website
#   nmake ADFILE=google-ads.js ADFILE2=google-ads-horizontal-link-unit.js website

!include ../include/version.inc
SRCROOT=$(MAKEDIR)\..
!include ../include/tooldefs.inc

# Comment line below to see command lines being executed
# Or just invoke nmake with SILENT=
SILENT=@
!if "x$(SILENT)" == "x@"
DTPQUIET = -quiet
!endif

!ifndef TOOLDIR
TOOLDIR= $(SRCROOT)\..\tools
!endif

TIDY="$(TOOLDIR)\tidy.exe"
DTP="$(TOOLDIR)\twapi-dtp.kit"
!ifndef HHC
HHC="$(PROGRAMFILES)\HTML Help Workshop\hhc.exe"
!endif

# Documentation input files. Note this also specifies the order in the ToC
!ifndef MANFILES
MANFILES= \
	overview.man \
	packages.man \
	installation.man \
	raw.man \
	base.man \
	ui.man \
	com.man \
	shell.man \
	input.man \
	sound.man \
	clipboard.man \
	osinfo.man \
	process.man \
	services.man \
	storage.man \
	share.man \
	users.man \
	network.man \
	namedpipe.man \
	security.man \
	crypto.man \
	eventlog.man \
	etw.man \
	apputil.man \
	console.man \
	winsta.man \
	schedule.man \
	resource.man \
	msi.man \
	nls.man \
	printer.man \
	device.man \
	synch.man \
	power.man \
	versionhistory.man 
!endif

IMGFILES = twapi-32.png twapi-48.png magicsplat.png navarrow.png download.png download_active.png


!ifndef FMT
FMT=htm
!endif

HTMLOUTPUTFILES=$(MANFILES:.man=.html)
HTFOUTPUTFILES=$(MANFILES:.man=.htf)
CHMDIR=chm-build

all: $(CHMDIR)\twapi.chm makefile

$(CHMDIR)\twapi.chm : $(MANFILES) styleschm.css makefile html_head.inc generate_mshelp_hhp.tcl generate_mshelp_index.tcl generate_mshelp_contents.tcl
	@echo =========Building CHM file...
	@echo Making directories...
	-$(SILENT)if not exist $(CHMDIR) mkdir chm-build
	@echo Copying man files
	$(SILENT)xcopy /q /y *.man $(CHMDIR)
	@echo Processing man files...
	$(SILENT)cd $(CHMDIR) && $(SILENT)$(TCLSH) $(DTP) map -ext html $(MANFILES) > iomap.tmp
	$(SILENT)cd $(CHMDIR) && $(SILENT)$(TCLSH) $(DTP) meta iomap.tmp > meta.tmp
	$(SILENT)cd $(CHMDIR) && $(SILENT)$(TCLSH) $(DTP) idx -title "Tcl Windows API Extension - Index" meta.tmp > idx.tmp
	$(SILENT)cd $(CHMDIR) && $(SILENT)$(TCLSH) $(DTP) gen-doc -varfile meta ../html_head.inc $(FMT) iomap.tmp
	@echo Generating MSHELP input files...
	$(SILENT)cd $(CHMDIR) && $(SILENT)$(TCLSH) ../generate_mshelp_hhp.tcl $(HTMLOUTPUTFILES) styleschm.css > twapi.hhp
	$(SILENT)cd $(CHMDIR) && $(SILENT)$(TCLSH) ../generate_mshelp_index.tcl idx.tmp > twapi.hhk
	$(SILENT)cd $(CHMDIR) && $(SILENT)$(TCLSH) ../generate_mshelp_contents.tcl meta.tmp > twapi.hhc
	$(SILENT)xcopy /q /y styleschm.css $(CHMDIR)
	$(SILENT)for %f in ($(IMGFILES)) do $(SILENT)xcopy /q /y %f $(CHMDIR)
#	$(SILENT)xcopy /q /y *.example $(CHMDIR)
# HHC always seems to return exit code 1!!
	@echo Generating CHM file...
	-$(SILENT)cd $(CHMDIR) && $(SILENT)$(HHC) twapi.hhp

htmdistribution: readme.txt $(CHMDIR)\twapi.chm
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
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) gen-doc htf iomap.tmp
	$(SILENT)cd website-build && $(SILENT)$(TCLSH) $(DTP) gen-toc htf iomap.tmp toc.tmp > toc.htf
	@echo Generating web pages...
	$(SILENT)$(TCLSH) generate_www_index.tcl website-build\idx.tmp > website-build\idx.htf
        $(SILENT)for %f in ($(HTFOUTPUTFILES) idx.htf) do $(SILENT)$(TCLSH) generate_web_page.tcl website-build\%f $(MODULEVERSION) website-build\toc.htf $(ADFILE) $(ADFILE2) > website-build\%~nf.html
	@echo Tidying web pages...
        $(SILENT)for %f in ($(HTMLOUTPUTFILES) idx.html) do echo %f && $(SILENT)$(TIDY) -asxhtml -q website-build\%f > website\%f
        $(SILENT)$(TCLSH) generate_web_page.tcl index.htf $(MODULEVERSION) website-build\toc.htf $(ADFILE) $(ADFILE2) | $(TIDY) -asxhtml -q > website\index.html
	@echo Copying files...
	$(SILENT)xcopy /q /y styles.css website
	$(SILENT)for %f in ($(IMGFILES)) do $(SILENT)xcopy /q /y %f website
#	$(SILENT)xcopy /q /y *.example website
	@echo Zipping distribution...
	$(SILENT)$(ZIP) -q -r twapi-website.zip website

clean:
	-del *.tmp
	-del *.*~
	-del idx.html
	-del toc.html
	-del $(HTMLOUTPUTFILES)
	-rmdir/s/q chm-build
	-rmdir/s/q website-build
	-rmdir/s/q website
