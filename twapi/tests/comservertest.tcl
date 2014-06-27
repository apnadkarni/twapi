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

proc find_wish {} {
    if {[info commands wm] eq ""} {
        set dir [file dirname [info nameofexecutable]]
        set wishes [glob -directory $dir wish*.exe]
        if {[llength $wishes] != 1} {
            error "Multiple wish wishes found"
        }
        set wish [file normalize [lindex $wishes 0]]
    } else {
        # We are running wish already
        set wish [info nameofexecutable]
    }

    return [file nativename [file attributes $wish -shortname]]
}

proc install {} {
    twapi::install_comserver $::comserver(progid) $::comserver(clsid) 1 -command "[find_wish] [file attributes $::comserver(script) -shortname]"
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
    twapi::class create TestComServer {
        constructor {} {
            my variable _interp
            set _interp [interp create -safe]
        }
        method Eval script {
            my variable _interp
            return [$_interp eval $script]
        }
        method EvalArgs args {
            my variable _interp
            return [$_interp eval $args]
        }
        export Eval EvalArgs
    }
    twapi::comserver_factory $::comserver(clsid) {0 Eval 1 EvalArgs} {TestComServer new} factory
    factory register
    twapi::run_comservers
    factory destroy
}

catch {wm withdraw .}
run
# Since we might be in Wish, explicitly exit
exit
