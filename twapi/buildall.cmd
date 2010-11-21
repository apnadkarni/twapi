rem Builds all release configurations of TWAPI
rem For 64-bit builds, should be called from buildall64.cmd, not directly
rem To use the custom TWAPI compiler setup instead of the standard VC 6 and SDK,
rem define TWAPI_COMPILER_DIR environment var appropriately before running this.

IF "x%CPU%" == "xAMD64" goto cleanbin

rem Set up 32-bit build environment. If we are using the TWAPI custom
rem environment point there, else the standard Microsoft paths
IF %TWAPI_COMPILER_DIR%. == . goto setupsdk

@call "%TWAPI_COMPILER_DIR%"\x86\setup.bat
goto cleanbin

:setupsdk

@call "%ProgramFiles%\Microsoft Visual Studio\VC98\Bin\vcvars32.bat"
@call "%ProgramFiles%\Microsoft Platform SDK\setenv.cmd" /2000 /RETAIL

:cleanbin
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

rem Libraries
cd base && nmake /s /nologo /a clean lib EMBED_SCRIPT=lzma TWAPI_STATIC_BUILD=1
