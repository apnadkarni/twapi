:: Builds release versions twapi using the "official" twapi compiler

setlocal

set dir90x64=d:\Tcl\90\x64-debug
set dir90x86=d:\Tcl\90\x86-debug
set dir86x64=d:\Tcl\86\x64-debug
set dir86x86=d:\Tcl\86\x86-debug

powershell .\release.ps1 %dir90x64% %~dp0\dist\latest x64
@if ERRORLEVEL 1 goto error_exit

powershell .\release.ps1 %dir90x86% %~dp0\dist\latest x86
@if ERRORLEVEL 1 goto error_exit

powershell .\release.ps1 %dir86x64% %~dp0\dist\latest x64
@if ERRORLEVEL 1 goto error_exit

powershell .\release.ps1 %dir86x86% %~dp0\dist\latest x86
@if ERRORLEVEL 1 goto error_exit

echo lappend auto_path dist/latest; puts [package require twapi] ; exit | %dir90x64%\bin\tclsh90.exe
@if ERRORLEVEL 1 goto error_exit

echo lappend auto_path dist/latest; puts [package require twapi] ; exit | %dir90x86%\bin\tclsh90.exe
@if ERRORLEVEL 1 goto error_exit

echo lappend auto_path dist/latest; puts [package require twapi] ; exit | %dir86x64%\bin\tclsh86t.exe
@if ERRORLEVEL 1 goto error_exit

echo lappend auto_path dist/latest; puts [package require twapi] ; exit | %dir86x86%\bin\tclsh86t.exe
@if ERRORLEVEL 1 goto error_exit

goto vamoose

:error_exit
@echo "ERROR: Build failed"
exit /B 1

:vamoose


