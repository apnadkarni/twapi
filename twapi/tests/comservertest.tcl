array set comserver {
    clsid {{332B8252-2249-4B34-BAD3-81259F2A2842}}
    progid TwapiComserver.Test
}
set comserver(script) [file normalize [info script]]

set dist_dir [file normalize [file join [file dirname [info script]] .. dist twapi]]
if {[file exists [file join $dist_dir pkgIndex.tcl]]} {
    lappend auto_path $dist_dir
}
package require twapi


proc install {} {
    twapi::install_comserver $::comserver(progid) $::comserver(clsid) 1 -script $::comserver(script)
}

proc uninstall {} {
    twapi::uninstall_comserver $::comserver(progid)
}

proc run {} {
    set ::argv [lassign $::argv command]

    switch -exact -nocase -- $command {
        -embedding -
        /embedding {
            # Just fall through to run the component
        }
        -regserver -
        /regserver {
            install
            return
        }
        -unregserver -
        /unregserver {
            uninstall
            return
        }
        default {
            error "Unknown command. Must be one of /regserver, /unregserver or /embedding"
        }
    }
    twapi::class create TestComserver {
        method Sum args {
            return [tcl::mathop::+ {*}$args]
        }
        export Sum
    }
    twapi::comserver_factory $::comserver(clsid) {0 Sum} {TestComserver new} factory
    factory register
    twapi::run_comservers
    factory destroy
}

run
