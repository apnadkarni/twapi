#
# Generates a TWAPI pkgIndex.tcl file and writes it to standard output
# First arg is path to the directory containing the package.
# Second arg is whether to enable lazy
# loading. The module names are picked up from the pkgindex.modules
# file in the directory.

proc get_ns_commands {ns} {
    if {![namespace exists ${ns}]} {
        return {}
    }
    set cmds [info commands ${ns}::*]
    foreach alias [interp aliases {}] {
        if {[string match ${ns}::* $alias]} {
            lappend cmds $alias
        }
    }
    foreach childns [namespace children $ns] {
        lappend cmds {*}[get_ns_commands $childns]
    }

    # Nested aliases are duplicated so remove them
    return [lsort -unique $cmds]
}

proc get_twapi_commands {} {
    return [concat [get_ns_commands ::twapi] [get_ns_commands ::metoo]]
}

proc dependents {mods modpath} {
    set deps {}
    foreach dep [dict get $mods [lindex $modpath end] dependencies] {
        if {$dep in $modpath} {
            error "Circular dependency: [join $modpath ->]->$dep"
        }
        if {$dep in $deps} {
            continue;           # Already dealt with this
        }
        lappend modpath $dep
        foreach dep2 [dependents $mods $modpath] {
            if {$dep2 ni $deps} {
                lappend deps $dep2
            }
        }
        lappend deps $dep
    }
    return $deps
}

proc makeindex {pkgdir lazy} {

    # Read the list of modules and types
    set fd [open [file join $pkgdir pkgindex.modules]]
    set lines [split [read $fd] \n]
    close $fd

    # Arrange the modules such that dependencies are loaded
    # before dependents
    set mods [dict create]
    foreach line $lines {
        set line [string trim $line]
        if {$line eq ""} continue
        set dependencies [lassign $line mod type]
        if {$mod eq "" || $type eq ""} {
            error "Empty module name or type in pkgindex.modules file"
        }
        dict set mods $mod type $type
        dict set mods $mod dependencies $dependencies
    }

    # $order contains order in which to load so that dependencies are
    # loaded first. Even though twapi_base is explicitly loaded,
    # include it here because we want to enumerate its exports
    # TBD - is this why it shows up twice in the pkgIndex.tcl entry
    # for package twapi
    set order [list twapi_base]
    foreach mod [dict keys $mods] {
        # Don't bother removing duplicates, we will just ignore them
        # in loop below
        lappend order {*}[dependents $mods $mod] $mod
    }

    # We need to figure out the commands in this module. See what we
    # have so far and then make a diff
    set modinfo [dict create]
    if {$lazy} {
        array set commands {}
        foreach cmd [get_twapi_commands] {
            set commands($cmd) ""
        }
    }

    foreach mod $order {
        set type [dict get $mods $mod type]
        if {[dict exists $modinfo $mod]} {
            # Modules may already be loaded due to dependencies
            continue
        }
        set dll {}
        if {$type eq "load"} {
            # Binary module. See if statically bound or not
            set modfile $mod
            if {$::tcl_platform(pointerSize) == 8} {
                append modfile "64"
            }
            if {[file exists [file join $pkgdir ${modfile}.dll]]} {
                set dll [file join $pkgdir ${mod}.dll]
            }
            uplevel #0 [list load $dll $mod]
            if {$mod eq "twapi_base"} {
                set ver [twapi::get_version -patchlevel]
            }
        } else {
            # Pure script
            twapi::Twapi_SourceResource $mod
        }
        set mod_cmds [dict create]
        if {$lazy} {
            foreach cmd [get_twapi_commands] {
                set ns [namespace qualifiers $cmd]
                set tail [namespace tail $cmd]
                
                # Only include twapi extension commands (inc. metoo)
                if {![regexp {^::(twapi|metoo)::[_a-zA-Z].*} $cmd]} {
                    continue
                }
                
                if {![info exists commands($cmd)]} {
                    set commands($cmd) $mod
                    dict lappend mod_cmds $ns $tail
                }
            }
            if {[dict size $mod_cmds] == 0} {
                error "No public commands found in twapi module $mod"
            }
        }
        # Set dll to be base name of dll or empty if no dll
        if {$dll ne ""} {
            set dll $mod
        } else {
            set dll "{}"
        }
        dict set modinfo $mod load_command "package ifneeded $mod $ver \[list twapi::package_setup \$dir $mod $ver $type $dll {$mod_cmds}\]"
    }

    return $modinfo
}

if {[info script] eq $::argv0} {
    set ::argv [lassign $::argv dir lazy]
#    puts "package ifneeded twapi_base $ver \[list apply {{d} {load \[file join \$d twapi_base.dll\] ; package provide twapi_base $ver}} \$dir\]"
    dict for {mod info} [makeindex $dir $lazy] {
        puts [dict get $info load_command]
        lappend mods $mod
    }
    set ver [twapi::get_version -patchlevel]
    puts "package ifneeded twapi $ver {"
    foreach mod $mods {
        puts "  package require $mod $ver"
    }
    puts "\n  package provide twapi $ver"
    puts "}"
}
