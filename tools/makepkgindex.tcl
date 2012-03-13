#
# Generates a TWAPI pkgIndex.tcl file and writes it to standard output
# First arg is path to an embedded twapi dll. Remaining args are
# module names

proc get_twapi_commands {} {
    return [concat [info commands ::twapi::*] [info commands ::metoo::*]]
}

proc makeindex {pkgdir ver} {
global commands
    # buildtype must be "single" or "multi"

    # Read the list of modules and types
    set fd [open [file join $pkgdir pkgindex.modules]]
    set mods [read $fd]
    close $fd

    # We need to figure out the commands in this module. See what we
    # have so far and then make a diff
    array set commands {}
    set modinfo [dict create]
    foreach cmd [get_twapi_commands] {
        set commands($cmd) ""
    }

    foreach {mod type} $mods {
        if {$mod eq "twapi_base"} continue
puts $mod
        dict set modinfo $mod type $type
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
        set cmds [get_twapi_commands]
        if {[llength $cmds] == 0} {
            puts "No commands ($mod)"
        }
        foreach cmd $cmds {
            if {![regexp {^::(twapi|metoo)::[a-z].*} $cmd]} {
                if {$mod eq "twapi_wmi"} {
                    puts "Skipping $cmd"
                }
                continue
            }
            if {![info exists commands($cmd)]} {
                set commands($cmd) $mod
                lappend mod_cmds $cmd
            } else {
                if {$mod eq "twapi_wmi"} {
                    puts "Skipping (2) $cmd"
                }
            }
        }
        if {[llength $mod_cmds] == 0} {
            error "No public commands found in twapi module $mod"
        }
        dict set modinfo $mod load_command "package ifneeded $mod $ver \[list twapi::package_setup \$dir $mod $ver $type $dll {$mod_cmds}\]"
    }

    return $modinfo
}

if {[info script] eq $::argv0} {
    set ::argv [lassign $::argv dir]
    load [file join $dir twapi_base.dll] twapi
    set ver [twapi::get_version -patchlevel]
    puts "package ifneeded twapi_base $ver \[list apply {{d} {load \[file join \$d twapi_base.dll\] ; package provide twapi_base $ver}} \$dir\]"
    dict for {mod info} [makeindex $dir $ver] {
        puts [dict get $info load_command]
    }
    puts "package ifneeded twapi $ver {"
    puts "  package require twapi_base $ver"
    foreach mod [lrange $::argv 1 end] {
        puts "  package require $mod $ver"
    }
    puts "\n  package provide twapi $ver"
    puts "}"
}