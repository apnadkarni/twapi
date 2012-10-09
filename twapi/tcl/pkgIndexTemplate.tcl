if {$::tcl_platform(os) ne "Windows NT" ||
    ($::tcl_platform(machine) ne "intel" &&
     $::tcl_platform(machine) ne "amd64")} {
    return
}

namespace eval twapi {}
proc twapi::package_setup {dir pkg version type {file {}} {commands {}}} {
    global auto_index

    if {$file eq ""} {
        set file $pkg
    }
    if {$::tcl_platform(pointerSize) == 8} {
        set fn [file join $dir "${file}64.dll"]
    } else {
        set fn [file join $dir "${file}.dll"]
    }

    if {$fn ne ""} {
        if {![file exists $fn]} {
            set fn "";          # Assume twapi statically linked in
        }
    }

    if {$pkg eq "twapi_base"} {
        # Need the twapi base of the same version
        # In tclkit builds, twapi_base is statically linked in
        foreach pair [info loaded] {
            if {$pkg eq [lindex $pair 1]} {
                set fn [lindex $pair 0]; # Possibly statically loaded
                break
            }
        }
        set loadcmd [list load $fn $pkg]
    } else {
        package require twapi_base $version
        if {$type eq "load"} {
            # Package could be statically linked or to be loaded
            if {[twapi::get_build_config single_module]} {
                # Modules are statically bound. Reset fn
                set fn {}
            }
            set loadcmd [list load $fn $pkg]
        } else {
            # A pure Tcl script package
            set loadcmd [list twapi::Twapi_SourceResource $file 1]
        }
    }

    if {[llength $commands] == 0} {
        # No commands specified, load the package right away
        # TBD - what about the exports table?
        uplevel #0 $loadcmd
    } else {
        # Set up the load for when commands are actually accessed
        # TBD - add a line to export commands here ?
        foreach {ns cmds} $commands {
            foreach cmd $cmds {
                if {[string index $cmd 0] ne "_"} {
                    dict lappend ::twapi::exports $ns $cmd
                }
                set auto_index(${ns}::$cmd) $loadcmd
            }
        }
    }

    # TBD - really necessary? The C modules do this on init anyways.
    # Maybe needed for pure scripts
    package provide $pkg $version
}

# The build process will append package ifneeded commands below
# to create an appropriate pkgIndex.tcl file for included modules
