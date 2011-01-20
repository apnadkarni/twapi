# Exported from Visual Studio Project and hacked so 64-bit can be built
!if "$(CPU)" == "AMD64"
OUTDIR=.\x64
INTDIR=.\x64\Release
!else
OUTDIR=.\x86
INTDIR=.\x86\Release
!endif

ALL : "$(OUTDIR)\rctest.dll"

CLEAN :
	-@erase "$(INTDIR)\rctest.obj"
	-@erase "$(INTDIR)\rctest.res"
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(OUTDIR)\rctest.dll"
	-@erase "$(OUTDIR)\rctest.exp"
	-@erase "$(OUTDIR)\rctest.lib"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

"$(INTDIR)" :
    if not exist "$(INTDIR)/$(NULL)" mkdir "$(INTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MT /W3 /Oi /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "RCTEST_EXPORTS" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/nologo /D "NDEBUG" /mktyplib203 /win32 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\rctest.res" /d "NDEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\rctest.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /incremental:no /pdb:"$(OUTDIR)\rctest.pdb" /machine:$(CPU) /out:"$(OUTDIR)\rctest.dll" /implib:"$(OUTDIR)\rctest.lib" 

!if "$(CPU)" == "AMD64"
LINK32_FLAGS = $(LINK32_FLAGS) bufferoverflowU.lib
!endif

LINK32_OBJS= \
	"$(INTDIR)\rctest.obj" \
	"$(INTDIR)\rctest.res"

"$(OUTDIR)\rctest.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

.\rctest.rc : \
	".\arrow.cur"\
	".\bitmap.bmp"\
	".\cursor.cur"\
	".\html1.htm"\
	".\icon1.ico"\
	".\icon2.ico"

SOURCE=.\rctest.c

"$(INTDIR)\rctest.obj" : $(SOURCE) "$(INTDIR)"


SOURCE=.\rctest.rc

"$(INTDIR)\rctest.res" : $(SOURCE) "$(INTDIR)"
	$(RSC) $(RSC_PROJ) $(SOURCE)


