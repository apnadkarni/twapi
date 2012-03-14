#
# Generates a TWAPI pkgIndex.tcl file and writes it to standard output
# First arg is path to the directory containing the package.
# Second arg is whether to enable lazy
# loading. The module names are picked up from the pkgindex.modules
# file in the directory.

proc get_twapi_commands {} {
    return [concat [info commands ::twapi::*] [info commands ::metoo::*]]
}

proc makeindex {pkgdir lazy ver} {

    # Read the list of modules and types
    set fd [open [file join $pkgdir pkgindex.modules]]
    set mods [read $fd]
    close $fd

    # We need to figure out the commands in this module. See what we
    # have so far and then make a diff
    set modinfo [dict create]
    if {$lazy} {
        array set commands {}
        foreach cmd [get_twapi_commands] {
            set commands($cmd) ""
        }
    }

    foreach {mod type} $mods {
        if {$mod eq "twapi_base"} continue
        if {! $lazy} {
            if {$type eq "load"} {
                if {[file exists [file join $pkgdir ${mod}.dll]]} {
                    # Separate dll
                    dict set modinfo $mod load_command "package ifneeded $mod $ver \"load \[file join \$dir ${mod}.dll \] $mod ; package provide $mod $ver\""
                } else {
                    # Bin objects are statically bound
                    dict set modinfo $mod load_command "package ifneeded $mod $ver \"load {} $mod ; package provide $mod $ver\""
                }
            } else {
                # A pure Tcl script bound into the main twapi_base dll
                dict set modinfo $mod load_command "package ifneeded $mod $ver \"twapi::Twapi_SourceResource $mod ; package provide $mod $ver\""
            }            

            continue
        }
        # We want lazy loading
        set dll {}
        if {$type eq "load"} {
            # Binary module. See if statically bound or not
            if {[file exists [file join $pkgdir ${mod}.dll]]} {
                set dll [file join $pkgdir ${mod}.dll]
            }
            uplevel #0 [list load $dll $mod]
        } else {
            # Pure script
            twapi::Twapi_SourceResource $mod
        }
        set mod_cmds {}
        foreach cmd [get_twapi_commands] {
            if {![regexp {^::(twapi|metoo)::[a-z].*} $cmd]} {
                continue
            }
            if {![info exists commands($cmd)]} {
                set commands($cmd) $mod
                lappend mod_cmds $cmd
            }
        }
        if {[llength $mod_cmds] == 0} {
            error "No public commands found in twapi module $mod"
        }
        if {$dll ne ""} {
            # $dll is path on build system. Change to path on target system
            # when package is loaded.
            set dll "\[file join \$dir [file tail $dll]\]"
        } else {
            set dll "{}"
        }
        dict set modinfo $mod load_command "package ifneeded $mod $ver \[list twapi::package_setup \$dir $mod $ver $type $dll {$mod_cmds}\]"
    }

    return $modinfo
}

if {[info script] eq $::argv0} {
    set ::argv [lassign $::argv dir lazy]
    load [file join $dir twapi_base.dll] twapi
    set ver [twapi::get_version -patchlevel]
    puts "package ifneeded twapi_base $ver \[list apply {{d} {load \[file join \$d twapi_base.dll\] ; package provide twapi_base $ver}} \$dir\]"
    dict for {mod info} [makeindex $dir $lazy $ver] {
        puts [dict get $info load_command]
        lappend mods $mod
    }
    puts "package ifneeded twapi $ver {"
    puts "  package require twapi_base $ver"
    foreach mod $mods {
        puts "  package require $mod $ver"
    }
    puts "\n  package provide twapi $ver"
    puts "}"
}