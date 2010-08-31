rem Builds all 64-bit release configurations of TWAPI

rem Clean out build environment before calling sdk setup
set INCLUDE=
SET LIB=
SET MSDEVDIR=
set MSVCDIR=
SET PATH=%WINDIR%\SYSTEM32


rem Setup build environment
call C:\PROGRA~1\MICROS~3\SetEnv.cmd /XP64 /RETAIL

call buildall.cmd


