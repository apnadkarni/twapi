# This file is automatically generated and will be overwritten
namespace eval twapi {
    variable version 3.0
    variable patchlevel 3.0.6
    variable package_name twapi
    if {$::tcl_platform(machine) eq "amd64"} {
        variable dll_base_name twapi64
    } else {
        variable dll_base_name twapi
    }
}
