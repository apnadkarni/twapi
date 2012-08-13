rem Builds all 64-bit release configurations of TWAPI

rem Clean out build environment before calling sdk setup
set INCLUDE=
SET LIB=
SET MSDEVDIR=
set MSVCDIR=
SET PATH=%WINDIR%\SYSTEM32

rem Setup build environment
IF %TWAPI_COMPILER_DIR%. == . goto setupsdk
@call "%TWAPI_COMPILER_DIR%"\x64\setup.bat

goto dobuild

:setupsdk
call "%ProgramFiles%\Microsoft Platform SDK\SetEnv.cmd" /XP64 /RETAIL

:dobuild
call buildall.cmd %1


