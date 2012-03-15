rem NOTE: MUST BE RUN FROM the directory containing this file
rem Builds all release configurations of TWAPI
rem For 64-bit builds, should be called from buildall64.cmd, not directly
rem To use the custom TWAPI compiler setup instead of the standard VC 6 and SDK,
rem define TWAPI_COMPILER_DIR environment var appropriately before running this.

rem We first clean up existing builds, then call out to build the 64 bit
rem binaries and then continue here to build the 32 bit binaries and
rem distributions. Note the 64 bit build invokes another instance of
rem this batch file to do the actual build after setting up the environ.

IF "x%CPU%" == "xAMD64" goto dobuild

rem Clean out existing builds
pause Existing build and dist directories will be deleted. Ctrl-C to abort
rmdir/s/q build dist

rem Set up 32-bit build environment. If we are using the TWAPI custom
rem environment point there, else the standard Microsoft paths
IF %TWAPI_COMPILER_DIR%. == . goto setupsdk

@call "%TWAPI_COMPILER_DIR%"\x86\setup.bat
goto call64build

:setupsdk

@call "%ProgramFiles%\Microsoft Visual Studio\VC98\Bin\vcvars32.bat"
@call "%ProgramFiles%\Microsoft Platform SDK\setenv.cmd" /2000 /RETAIL

:call64build
rem Do 64-bit build first so dll can be included in full distro
cmd /c buildall64.cmd

:dobuild
rem Doing actual build
nmake /nologo /a twapi
nmake /nologo /a twapi-bin
nmake /nologo /a twapi-modular

rem Libraries
rem cd base && nmake /s /nologo /a EMBED_SCRIPT=lzma TWAPI_STATIC_BUILD=1 clean lib 
rem cd ..
