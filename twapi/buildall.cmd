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

:: Set up 32-bit build environment. If we are using the TWAPI custom
:: environment point there, else the standard Microsoft paths
IF %TWAPI_COMPILER_DIR%. == . goto setupsdk

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
IF "x%1" == "x" goto build_twapi
IF NOT "x%1" == "xtwapi" goto check_twapi_bin
:build_twapi
@echo BUILDING twapi
nmake /nologo /a /s twapi

:check_twapi_bin
IF "x%1" == "x" goto build_twapi_bin
IF NOT "x%1" == "xtwapi_bin" goto check_twapi_mod
:build_twapi_bin
@echo BUILDING twapi-bin
nmake /nologo /a /s twapi-bin


:check_twapi_mod
IF "x%1" == "x" goto build_twapi_mod
IF NOT "x%1" == "xtwapi_bin" goto build_lib
:build_twapi_mod
@echo BUILDING twapi-modular
nmake /nologo /a /s twapi-modular

IF "x%1" == "x" goto build_lib
IF NOT "x%1" == "xtwapi_lib" goto vamoose
:build_lib
nmake /nologo /a /s twapi-lib

:vamoose