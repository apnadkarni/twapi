
comtestps.dll: dlldata.obj comtest_p.obj comtest_i.obj
	link /dll /out:comtestps.dll /def:comtestps.def /entry:DllMain dlldata.obj comtest_p.obj comtest_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del comtestps.dll
	@del comtestps.lib
	@del comtestps.exp
	@del dlldata.obj
	@del comtest_p.obj
	@del comtest_i.obj
