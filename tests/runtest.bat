set matchparam=%2
if "x%matchparam%" == "x" set matchparam="*"
if "x%TCLSH%" == "x" set TCLSH=tclsh86t
%TCLSH% %1.test -constraint userInteraction -verbose pe -match %matchparam% %3 %4 %5 %6 %7 %8 %9
