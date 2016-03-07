:: Builds all 64-bit release configurations of TWAPI

:: Generate the mercurial ID. Need to do this before resetting PATH below
:: The first echo is a hack to write to a file without a terminating newline
:: We do it this way and not through the makefile because the compiler
:: build env is pristine and does not have a path to hg
if exist build\hgid.tmp goto init
echo|set /P=HGID=>build\hgid.tmp
hg identify -i >>build\hgid.tmp

:init
:: Clean out build environment before calling sdk setup
set INCLUDE=
SET LIB=
SET MSDEVDIR=
set MSVCDIR=
SET PATH=%WINDIR%\SYSTEM32

:: Setup build environment
IF NOT %TWAPI_COMPILER_DIR%. == . goto setuptwapicompiler

if NOT EXIST "c:\bin\x86\twapi-tcl-vc6" goto setupsdk
set TWAPI_COMPILER_DIR=c:\bin\x86\twapi-tcl-vc6

:setuptwapicompiler
@call "%TWAPI_COMPILER_DIR%"\x64\setup.bat
goto dobuild

:setupsdk
if NOT EXIST "%ProgramFiles%\Microsoft Platform SDK\SetEnv.cmd" goto setupsdk2
call "%ProgramFiles%\Microsoft Platform SDK\SetEnv.cmd" /XP64 /RETAIL
goto dobuild

:setupsdk2
if NOT EXIST "%ProgramFiles(x86)%\Microsoft Platform SDK\SetEnv.cmd" goto sdkerror
call "%ProgramFiles(x86)%\Microsoft Platform SDK\SetEnv.cmd" /XP64 /RETAIL
goto dobuild

:sdkerror
@echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
@echo ERROR: Could not set up 64-bit compiler and SDK
@echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
@exit /B 1

:dobuild
call buildone.cmd %1


