:: Builds release versions twapi using the "official" twapi compiler
:: See BUILD_INSTRUCTIONS for more information

:: release1.cmd manipulates env and paths so call in child shell.

cmd /c release1.cmd x86 %1
@if ERRORLEVEL 1 goto error_exit

cmd /c release1.cmd amd64 %1
@if NOT ERRORLEVEL 1 goto vamoose

:error_exit
@echo "ERROR: Build failed"

:vamoose


