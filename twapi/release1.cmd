:: This build all specified files for a single architecture. It mucks with
:: environment, paths etc. so should not be generally invoked within
:: the interactive command shell directly. The master release file
:: invokes it in a separate command shell.
:: See the BUILD_INSTRUCTIONS file for more details.

setlocal

:: Make output directories if they do not exist
mkdir build
mkdir dist

:: Generate the mercurial ID. Need to do this before resetting PATH below
:: The first echo is a hack to write to a file without a terminating newline
:: We do it this way and not through the makefile because the compiler
:: build env is pristine and does not have a path to hg
if exist build\hgid.tmp goto init
echo|set /P=HGID=>build\hgid.tmp
hg identify -i >>build\hgid.tmp

:init
set arch=%1
set target=%2
:: If target is unspecified, or an empty string, treat as all
IF ".%target%" == "." set target=all
IF ".%target%" == ".""" set target=all

:: Clean out build environment before calling sdk setup
set INCLUDE=
SET LIB=
SET MSDEVDIR=
set MSVCDIR=
SET PATH=%WINDIR%\SYSTEM32


:: Check if we are using our customized compiler setup
IF NOT ".%TWAPI_COMPILER_DIR%" == "." goto check_arch
if NOT EXIST "c:\bin\x86\twapi-tcl-vc6" goto check_arch 
set TWAPI_COMPILER_DIR=c:\bin\x86\twapi-tcl-vc6

:check_arch

if /i ".%arch%" == ".x86" goto setup_x86
if /i ".%arch%" == ".amd64" goto setup_amd64
if /i ".%arch%" == ".x64" goto setup_amd64

@echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
@echo Syntax: %0 x86/x64/amd64 ?config?
@echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
@exit /B 1

:: =================================================
:: Setup the x86 environment
:setup_x86

:: Setup build environment
IF ".%TWAPI_COMPILER_DIR%" == "." goto setup_sdk_x86
:: We are using our custom compiler setup
@call "%TWAPI_COMPILER_DIR%"\x86\setup.bat
goto do_builds

:setup_sdk_x86
:: We are using the standard compiler setup
if NOT EXIST "%ProgramFiles%\Microsoft Visual Studio\VC98\Bin\vcvars32.bat" goto sdkerror
if NOT EXIST "%ProgramFiles%\Microsoft Platform SDK\setenv.cmd" goto sdkerror
@call "%ProgramFiles%\Microsoft Visual Studio\VC98\Bin\vcvars32.bat"
@call "%ProgramFiles%\Microsoft Platform SDK\setenv.cmd" /XP32 /RETAIL
goto do_builds

:: =================================================
:: Setup the x64/amd64 environment
:setup_amd64

:: Setup build environment
IF ".%TWAPI_COMPILER_DIR%" == "." goto setup_sdk_amd64
:: We are using our custom compiler setup
@call "%TWAPI_COMPILER_DIR%"\x64\setup.bat
goto do_builds

:setup_sdk_amd64
:: We are using the standard compiler setup
if NOT EXIST "%ProgramFiles%\Microsoft Platform SDK\SetEnv.cmd" goto setup_sdk_amd64_2
call "%ProgramFiles%\Microsoft Platform SDK\SetEnv.cmd" /XP64 /RETAIL
goto do_builds

:setup_sdk_amd64_2
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Platform SDK\SetEnv.cmd" goto sdkerror
call "%ProgramFiles(x86)%\Microsoft Platform SDK\SetEnv.cmd" /XP64 /RETAIL
goto do_builds

:sdkerror
@echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
@echo ERROR: Could not set up %arch% compiler and SDK
@echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
@exit /B 1

:do_builds

:: NOTE: we use %3... etc and not %* because that would include all
:: arguments even if we were to use shift

@if "x%target%" == "xall" goto build_twapi
@if NOT "x%target%" == "xtwapi" goto check_twapi_bin
:build_twapi
@echo BUILDING twapi %3 %4 %5 %6
nmake /nologo /a /s twapi %3 %4 %5 %6 %7 %8 %9

:check_twapi_bin
@if "x%target%" == "xall" goto build_twapi_bin
@if NOT "x%target%" == "xtwapi_bin" goto check_twapi_mod
:build_twapi_bin
@echo BUILDING twapi-bin
nmake /nologo /a /s twapi-bin %3 %4 %5 %6 %7 %8 %9

:check_twapi_mod
@if "x%target%" == "xall" goto build_twapi_mod
@if NOT "x%target%" == "xtwapi_modular" goto check_twapi_lib
:build_twapi_mod
@echo BUILDING twapi-modular
nmake /nologo /a /s twapi-modular %3 %4 %5 %6 %7 %8 %9

:check_twapi_lib
@if "x%target%" == "xall" goto build_lib
@if NOT "x%target%" == "xtwapi_lib" goto vamoose
:build_lib
@echo BUILDING twapi-lib
nmake /nologo /a /s twapi-lib %3 %4 %5 %6 %7 %8 %9

:vamoose
