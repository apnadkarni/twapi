# Named pipe test routines

if {[llength [info commands load_twapi_package]] == 0} {
    source [file join [file dirname [info script]] testutil.tcl]
}

if {[llength [info commands ::twapi::get_version]] == 0} {
    load_twapi_package
}

proc np_echo_usage {} {
    puts stderr {
Usage:
     tclsh npecho.tcl syncserver ?PIPENAME? ?TIMEOUT?
        -- runs the synchronous echo server
     tclsh npecho.tcl asyncserver ?PIPENAME? ?TIMEOUT?
        -- runs the asynchronous echo server
    }
    exit 1
}

proc np_echo_server_sync_accept {chan} {
    set ::np_echo_server_status connected
    fileevent $chan writable {}
}

proc np_echo_server_sync {{name {\\.\pipe\twapiecho}} {timeout 20000}} {
    set timer [after $timeout "set ::np_echo_server_status timeout"]
    set echo_fd [::twapi::namedpipe_server $name]

    fconfigure $echo_fd -buffering line -translation crlf -eofchar {} -encoding utf-8
    fileevent $echo_fd writable [list ::np_echo_server_sync_accept $echo_fd]
    # Following line is important as it is used by automated test scripts
    puts "READY"
    vwait ::np_echo_server_status
    after cancel $timer
    set msgs 0
    set last_size 0
    set total 0
    if {$::np_echo_server_status eq "connected"} {
        while {1} {
            if {[gets $echo_fd line] >= 0} {
                if {$line eq "exit"} {
                    break
                }
                puts $echo_fd $line
                incr msgs
                set last_size [string length $line]
                incr total $last_size
            } else {
                puts stderr "Unexpected eof from echo client"
                break
            }
        }
    } else {
        puts stderr "echo_server: $::np_echo_server_status"
    }
    close $echo_fd
    return [list $msgs $total $last_size]
}

proc np_echo_server_async_echoline {chan} {
    set ::np_echo_server_status connected

    # Check end of file or abnormal connection drop,
    # then echo data back to the client.

    set error [catch {gets $chan line} count]
    if {$error || $count < 0} {
        if {$error || [eof $chan]} {
            set ::np_async_end eof
            close $chan
        }
    } else {
        if {$line eq "exit"} {
            set ::np_async_end exit
        } else {
            incr ::np_msgs
            set ::np_last_size [string length $line]
            incr ::np_total $::np_last_size
            puts $chan $line
        }
    }
}

proc np_echo_server_async {{name {\\.\pipe\twapiecho}} {timeout 20000}} {
    set ::np_msgs 0
    set ::np_last_size 0
    set ::np_total 0

    set ::np_timer [after $timeout "set ::np_echo_server_status timeout"]
    set echo_fd [::twapi::namedpipe_server $name]
    fconfigure $echo_fd -buffering line -translation crlf -eofchar {} -encoding utf-8 -blocking 0
    fileevent $echo_fd readable [list ::np_echo_server_async_echoline $echo_fd]
    # Following line is important as it is used by automated test scripts
    puts "READY"
    vwait ::np_echo_server_status
    after cancel $::np_timer

    if {$::np_echo_server_status eq "connected"} {
        vwait ::np_async_end
        if {$::np_async_end ne "exit"} {
            puts stderr "Unexpected status $::np_async_end"
        }
    } else {
        puts stderr "echo_server: $::np_echo_server_status"
    }
    return [list $::np_msgs $::np_total $::np_last_size]
}

proc np_echo_client {args} {
    array set opts [twapi::parseargs args {
        {name.arg {\\.\pipe\twapiecho}}
        {density.int 2}
        {limit.int 10000}
    }]

    set alphabet "0123456789abcdefghijklmnopqrstuvwxyz"
    set alphalen [string length $alphabet]
    set msgs 0
    set last 0
    set total 0
    set fd [twapi::namedpipe_client $opts(name)]
    fconfigure $fd -buffering line -translation crlf -eofchar {} -encoding utf-8
    for {set i 1} {$i < $opts(limit)} {incr i [expr {1+($i+1)/$opts(density)}]} {
        set c [string index $alphabet [expr {$i % $alphalen}]]
        set request [string repeat $c $i]
        puts $fd $request
        set response [gets $fd]
        if {$request ne $response} {
            puts "Mismatch in message of size $i"
            set n [string length $response]
            if {$i != $n} {
                puts "Sent $i chars, received $n chars"
            }
        }
        incr msgs
        incr total $i
        set last $i
    }
    puts $fd "exit"
    close $fd
    return [list $msgs $total $last]
}



# Main code
if {[string equal -nocase [file normalize $argv0] [file normalize [info script]]]} {

    # We are being directly sourced, not as a library
    if {[llength $argv] == 0} {
        np_echo_usage
    }

    switch -exact -- [lindex $argv 0] {
        syncserver {
            if {[catch {
                foreach {nmsgs nbytes last} [eval np_echo_server_sync [lrange $argv 1 end]] break
            }]} {
                testlog $::errorInfo
            }
        }
        asyncserver {
            if {[catch {
                foreach {nmsgs nbytes last} [eval np_echo_server_async [lrange $argv 1 end]] break
            }]} {
                testlog $::errorInfo
            }
        }
        default {
            np_echo_usage
        }
    }

    puts [list $nmsgs $nbytes $last]
}
