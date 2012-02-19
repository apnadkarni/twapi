# Common file defining build rules
# This is split off from common.inc because we want to give the parent
# makefile a chance to define $(OBJS)

all: $(OBJDIR)\$(MODULE)

$(OBJDIR):
	@if not exist $(OBJDIR)\nul mkdir $(OBJDIR)

# Link the module
$(OBJDIR)\$(MODULE): $(OBJDIR)\$(MODULENAME).res $(OBJS)
        $(LINK) $(LOPT) -out:$@ $(LIBS) $**
!if $(VCVER) > 8
	cd $(OBJDIR) && $(MT) -manifest $(MODULE).manifest -outputresource:$(MODULE);2
!endif

$(OBJS): $(CINCLUDES) makefile

$(OBJDIR)\$(MODULENAME).res: $(OBJDIR) $(RC_FILE_NAME) makefile $(TWAPI_SCRIPT_RESOURCE_PATH)
        rc $(INCFLAGS) $(RCFLAGS) -fo $@ $(RC_FILE_NAME)

# Static library
lib: $(OBJDIR)\$(MODULELIB)

makefile: $(SRCROOT)\include\common.inc

distribution:
        $(COPY) $(OBJDIR)\$(MODULE) "$(PACKAGEDIR)"

$(OBJDIR)\$(MODULELIB): $(OBJDIR) $(OBJS)
	cd ..\tcl && nmake PACKAGEDIR="$(OBJDIR)\tcl" COMBINEOUTPUT="combined_$(MODULENAME)_script.lzma.raw" PACKAGE="$(PACKAGE)" combine
	"$(TOOLDIR)\lzma.exe" e "$(OBJDIR)\tcl\combined_$(MODULENAME)_script.lzma.raw" "$(OBJDIR)\combined_$(MODULENAME)_script.lzma"
	$(LIBTOOL) /OUT:$@ $**

# Compile a file
.c{$(OBJDIR)}.obj::
       $(CC) $(CFLAGS) -Fo$(OBJDIR)\ $<


!ifdef TWAPI_SCRIPT_RESOURCE
$(TWAPI_SCRIPT_RESOURCE_PATH):
!if "$(TWAPI_SCRIPT_RESOURCE_TYPE)" == "tclscript"
	cd ..\tcl && nmake PACKAGEDIR="$(TWAPI_SCRIPT_RESOURCE_DIR)" COMBINEOUTPUT="$(TWAPI_SCRIPT_RESOURCE)" PACKAGE=$(PACKAGE) combine
!else
	cd ..\tcl && nmake PACKAGEDIR="$(TWAPI_SCRIPT_RESOURCE_DIR)" COMBINEOUTPUT="$(TWAPI_SCRIPT_RESOURCE).raw" PACKAGE=$(PACKAGE) combine
	"$(TOOLDIR)\lzma.exe" e "$(TWAPI_SCRIPT_RESOURCE_PATH).raw" "$(TWAPI_SCRIPT_RESOURCE_PATH)"
!endif
!endif


clean:
        -$(RMDIR) "$(OBJDIR)"
