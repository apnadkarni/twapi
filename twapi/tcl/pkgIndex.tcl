# WHen script is sourced, the variable $dir must contain the
# full path name of this file's directory.

if {$::tcl_platform(os) eq "Windows NT" &&
    ($::tcl_platform(machine) eq "intel" ||
     $::tcl_platform(machine) eq "amd64") &&
    [string index $::tcl_platform(osVersion) 0] >= 5} {
    source [file join $dir twapi_version.tcl]
    source [file join $dir twapi_buildid.tcl]
    package ifneeded $::twapi::package_name $twapi::patchlevel [list source [file join $dir twapi.tcl]]
}
