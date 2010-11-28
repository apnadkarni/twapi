# Copies the test environment to target 
# tclsh setuptestenv.tcl TARGETPATH

if {[llength $argv] != 1} {
    puts stderr "Usage: [info nameofexecutable] [info script] TARGETPATH"
}

set target [file normalize [lindex $argv 0]]
if {[file exists $target]} {
    puts stderr "Error: Target path $target already exists. Will not overwrite."
    exit 1
}

file mkdir [set target_bindir [file join $target bin]]
file mkdir [set target_libdir [file join $target lib]]

set tcl_root [file dirname [file dirname [file normalize [info nameofexecutable]]]]
foreach fpat {*.dll *.exe} {
    foreach fn [glob [file join $tcl_root bin $fpat]] {
        file copy $fn $target_bindir
    }
}

# Our distribution
foreach dirpat {dde* reg* tcl8* tk*} {
    if {![catch {glob [file join $tcl_root lib $dirpat]} dirs]} {
        foreach dir $dirs {
            file copy $dir $target_libdir
        }
    }
}

# Finally copy the testscripts
file copy [file dirname [file normalize [info script]]] [file join $target tests]

puts "Remember to copy appropriate twapi distribution"
