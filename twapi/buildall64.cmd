:: Builds all 64-bit release configurations of TWAPI

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
call buildall.cmd %1


