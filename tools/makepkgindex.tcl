#
# Generates a TWAPI pkgIndex.tcl file and writes it to standard output
# First arg is path to an embedded twapi dll. Remaining args are
# module names

proc makeindex {mods} {
    array set commands {}
    set modules [dict create]
    foreach cmd [info commands ::twapi::*] {
        set commands($cmd) ""
    }

    foreach mod $mods {
        # Try loading as a static module. On error, try as a script only
        dict set modules $mod type load
        if {[catch {
            uplevel #0 [list load {} $mod]
        }]} {
            twapi::Twapi_SourceResource $mod
            dict set modules $mod type source
        }
        set mod_cmds {}
        foreach cmd [concat [info commands ::twapi::*] [info commands ::metoo::*]] {
            if {![regexp {^::(twapi|metoo)::[a-z].*} $cmd]} {continue}
            if {![info exists commands($cmd)]} {
                set commands($cmd) $mod
                lappend mod_cmds $cmd
            }
        }
        if {[llength $mod_cmds] == 0} {
            error "No public commands found in twapi module $mod"
        }
        dict set modules $mod commands $mod_cmds
    }

    return $modules
}

if {[info script] eq $::argv0} {
    uplevel #0 [list load [lindex $::argv 0] twapi]
    set ver [twapi::get_version -patchlevel]
    dict for {mod info} [makeindex [lrange $::argv 1 end]] {
        puts "package ifneeded $mod $ver \[list twapi::package_setup \$dir $mod $ver [dict get $info type] {} {[dict get $info commands]}\]"
    }
}