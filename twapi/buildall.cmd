rem Builds all release configurations of TWAPI

echo BUILDALL ENTER %CPU%

nmake /s /nologo /a clean
nmake /s /nologo /a EMBED_SCRIPT=lzma clean

rem Do not remove existing distributions if doing 64 bit build
rem and do 64-bit build first so dll can be included in full distro
IF "x%CPU%" == "xAMD64" goto dobuild
echo REMOVE TEMP %CPU%
rmdir/s/q temp
echo REMOVE TEMP %CPU%
cmd /c buildall64.cmd

:dobuild

echo DOING BUILD %CPU%
rem Doing build
nmake /s /nologo /a
echo DOING BUILD LZMA %CPU%
nmake /s /nologo /a EMBED_SCRIPT=lzma

rem Skip distribution in 64bit build - included in 32-bit builds
IF "x%CPU%" == "xAMD64" goto singlefile
echo DOING DISTRIBUTION %CPU%
nmake /s /nologo /a distribution
echo DOING DISTRIBUTION LZMA %CPU%
nmake /s /nologo /a EMBED_SCRIPT=lzma distribution

:singlefile
echo DOING TM %CPU%
nmake /s /nologo /a EMBED_SCRIPT=lzma tmdistribution
echo DOING DLL %CPU%
nmake /s /nologo /a EMBED_SCRIPT=lzma dlldistribution

