# Given a file, sticks Tcl comments characters in front of each line
set usage {
    usage: commentify.tcl INFILE OUTFILE
}

if {[llength $argv] != 2} {puts stderr $usage; exit 1}

set in  [open [lindex $argv 0] r]
set out [open [lindex $argv 1] w]

foreach line [split [read $in] \n] {
    puts $out "# $line"
}

close $in
close $out
