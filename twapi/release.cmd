:: Builds release versions twapi using the "official" twapi compiler
:: See BUILD_INSTRUCTIONS for more information

:: release1.cmd manipulates env and paths so call in child shell.

setlocal

SET target=%1
IF ".%target%" == "." SET target=all 

cmd /c release1.cmd amd64 %target% 
@if ERRORLEVEL 1 goto error_exit

:: Note MAKEDIST=1 has to be in quotes else cmd passes it as two params
:: "MAKEDIST" and "1". Arghh!
cmd /c release1.cmd x86 %target% "MAKEDIST=1"
@if NOT ERRORLEVEL 1 goto vamoose

:error_exit
@echo "ERROR: Build failed"
exit /B 1

:vamoose


