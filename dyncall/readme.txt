The dyncall library build is as follows:

In the dyncall source tree, buildsys\nmake\tool_msvc.nmak needs the
following patch:

40c40
< CFLAGS   = /EHsc /GR- /GS- /Ox /c /nologo /I$(TOP)\dyncall /I$(TOP)\dyncallback
---
> CFLAGS   = /MD /EHsc /GR- /GS- /Ox /c /nologo /I$(TOP)\dyncall /I$(TOP)\dyncallback

The above adds the /MD switch. Without this the dyncall static library
links to LIBCMT which causes conflicts since twapi links to MSVCRT.

Then, for the x64 build, at the toplevel dyncall dir,

nmake /f Nmakefile clean
.\configure /target-x64 /tool-msvc /config-release
nmake /f Nmakefile


NOTE: The x86 build using Visual C++ 6 will not work if you do not
have Microsoft 32-bit assembler (the 64-bit assembler does not work fo 32-bit).
If so, you have to install the MingW tool chain and follow the directions
in the dynload documentation for gcc/mingw builds.

For the x86 build, the dyncall distribution needs additional
modifications. First, the VC++ 6 compiler does not understand "long long"
so the following patch is needed
diff dyncall\dyncall_config.h~ dyncall\dyncall_config.h
42a43,45
> #ifdef _WIN32
> #define DC_LONG_LONG    __int64
> #else
43a47
> #endif

diff dyncallback\dyncall_args_x86.c~ dyncallback\dyncall_args_x86.c
65c65
< static long long default_i64(DCArgs* args)
---
> static DC_LONG_LONG default_i64(DCArgs* args)
67c67
<       long long result = * (long long*) args->stack_ptr;
---
>       DC_LONG_LONG result = * (DC_LONG_LONG*) args->stack_ptr;
116c116
< static long long fast_gnu_i64(DCArgs* args)
---
> static DC_LONG_LONG fast_gnu_i64(DCArgs* args)

Further, make sure a working x86 assembler ml.exe is in your path or
change the tool_msvc.mk definition AS to point to it.

Then

nmake /f Nmakefile clean
.\configure /target-x86 /tool-msvc /config-release
nmake /f Nmakefile

