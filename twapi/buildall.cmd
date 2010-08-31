rem Builds all release configurations of TWAPI

nmake /s /nologo /a clean
nmake /s /nologo /a EMBED_SCRIPT=lzma clean

rem Do not remove existing distributions if doing 64 bit build
rem and do 64-bit build first so dll can be included in full distro
IF "x%CPU%" == "xAMD64" goto dobuild
rmdir/s/q temp
cmd /c buildall64.cmd

:dobuild

rem Doing build
nmake /s /nologo /a
nmake /s /nologo /a EMBED_SCRIPT=lzma

rem Skip distribution in 64bit build - included in 32-bit builds
IF "x%CPU%" == "xAMD64" goto singlefile
nmake /s /nologo /a distribution
nmake /s /nologo /a EMBED_SCRIPT=lzma distribution

:singlefile
nmake /s /nologo /a EMBED_SCRIPT=lzma tmdistribution
nmake /s /nologo /a EMBED_SCRIPT=lzma dlldistribution

