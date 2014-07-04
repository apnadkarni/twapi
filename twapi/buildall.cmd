:: NOTE: MUST BE RUN FROM the directory containing this file
:: Builds all release configurations of TWAPI
:: For 64-bit builds, should be called from buildall64.cmd, not directly
:: To use the custom TWAPI compiler setup instead of the standard VC 6 and SDK,
:: define TWAPI_COMPILER_DIR environment var appropriately before running this.

:: We first clean up existing builds, then call out to build the 64 bit
:: binaries and then continue here to build the 32 bit binaries and
:: distributions. Note the 64 bit build invokes another instance of
:: this batch file to do the actual build after setting up the environ.

IF "x%CPU%" == "xAMD64" goto dobuild

:: Clean out existing builds
@echo Existing build and dist directories will be deleted. Ctrl-C to abort
@pause 
rmdir/s/q build dist 2>NUL
mkdir build

:: Generate the mercurial ID
:: The first echo is a hack to write to a file without a terminating newline
:: We do it this way and not through the makefile because the compiler
:: build env is pristine and does not have a path to hg
if exist build\hgid.tmp del build\hgid.tmp
echo|set /P=HGID=>build\hgid.tmp
hg identify -i >>build\hgid.tmp

:: Set up 32-bit build environment. If we are using the TWAPI custom
:: environment point there, else the standard Microsoft paths

IF NOT %TWAPI_COMPILER_DIR%. == . goto setuptwapicompiler

if NOT EXIST "c:\bin\x86\twapi-tcl-vc6" goto setupsdk
set TWAPI_COMPILER_DIR=c:\bin\x86\twapi-tcl-vc6

:setuptwapicompiler
@call "%TWAPI_COMPILER_DIR%"\x86\setup.bat
goto call64build

:setupsdk

@call "%ProgramFiles%\Microsoft Visual Studio\VC98\Bin\vcvars32.bat"
@call "%ProgramFiles%\Microsoft Platform SDK\setenv.cmd" /XP32 /RETAIL

:call64build
:: Do 64-bit build first so dll can be included in full distro
cmd /c buildall64.cmd %1
@if NOT ERRORLEVEL 1 goto dobuild
@exit /B 1

:dobuild
:: Doing actual build
@if "x%1" == "x" goto build_twapi
@if NOT "x%1" == "xtwapi" goto check_twapi_bin
:build_twapi
@echo BUILDING twapi %CPU%
nmake /nologo /a /s twapi

:check_twapi_bin
@if "x%1" == "x" goto build_twapi_bin
@if NOT "x%1" == "xtwapi_bin" goto check_twapi_mod
:build_twapi_bin
@echo BUILDING twapi-bin %CPU%
nmake /nologo /a /s twapi-bin

:check_twapi_mod
@if "x%1" == "x" goto build_twapi_mod
@if NOT "x%1" == "xtwapi_modular" goto check_twapi_lib
:build_twapi_mod
@echo BUILDING twapi-modular %CPU%
nmake /nologo /a /s twapi-modular

:check_twapi_lib
@if "x%1" == "x" goto build_lib
@if NOT "x%1" == "xtwapi_lib" goto vamoose
:build_lib
@echo BUILDING twapi-lib %CPU%
nmake /nologo /a /s twapi-lib

:vamoose