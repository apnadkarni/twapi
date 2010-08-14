rem Builds all release configurations of TWAPI

rmdir/s/q temp

rem Traditional distribution
nmake /s /nologo /a
nmake /s /nologo /a distribution

rem LZMA distributions
nmake /s /nologo /a EMBED_SCRIPT=lzma
rem DLL based package
nmake /s /nologo /a distribution
rem Tcl 8.5 module
nmake /s /nologo /a EMBED_SCRIPT=lzma tmdistribution
rem Tcl single file DLL
nmake /s /nologo /a EMBED_SCRIPT=lzma dlldistribution

