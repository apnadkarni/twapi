#
# Copyright (c) 2010, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Implementation of named pipes

proc twapi::namedpipe {name args} {
    array set opts [parseargs args {
        {type.arg client {server client}}
    }]
    if {$opts(type) eq "server"} {
        return [twapi::namedpipe_server $name {*}$args]
    } else {
        return [twapi::namedpipe_client $name {*}$args]
    }
}

proc twapi::namedpipe_server {name args} {
    set name [file nativename $name]

    # Only byte mode currently supported. Message mode does
    # not mesh well with Tcl channel infrastructure.
    # readmode.arg
    # writemode.arg

    array set opts [twapi::parseargs args {
        {mode.arg {read write}}
        writedacl
        writeowner
        writesacl
        writethrough
        denyremote
        {timeout.int 50}
        {maxinstances.int 255}
        {secd.arg {}}
        {inherit.bool 0}
    } -maxleftover 0]

    set open_mode [twapi::_parse_symbolic_bitmask $opts(mode) {read 1 write 2}]
    foreach {opt mask} {
        writedacl  0x00040000
        writeowner 0x00080000
        writesacl  0x01000000
        writethrough 0x80000000
    } {
        if {$opts($opt)} {
            set open_mode [expr {$open_mode | $mask}]
        }
    }
        
    set open_mode [expr {$open_mode | 0x40000000}]; # OVERLAPPPED I/O
        
    set pipe_mode 0
    if {$opts(denyremote)} {
        if {! [twapi::min_os_version 6]} {
            error "Option -denyremote not supported on this operating system."
        }
        set pipe_mode [expr {$pipe_mode | 8}]
    }

    return [twapi::Twapi_NPipeServer $name $open_mode $pipe_mode \
                $opts(maxinstances) 4000 4000 $opts(timeout) \
                [_make_secattr $opts(secd) $opts(inherit)]]
}

proc twapi::namedpipe_client {name args} {
    set name [file nativename $name]

    # Only byte mode currently supported. Message mode does
    # not mesh well with Tcl channel infrastructure.
    # readmode.arg
    # writemode.arg

    array set opts [twapi::parseargs args {
        {mode.arg {read write}}
        {secattr.arg {}}
    } -maxleftover 0]

    set desired_access [twapi::_parse_symbolic_bitmask $opts(mode) {
        read 0x80000000 write 0x40000000
    }]
        
    set share_mode 0;           # Share none
    set flags 0
    set create_disposition 3;   # OPEN_EXISTING
    return [twapi::Twapi_NPipeClient $name $desired_access $share_mode \
                $opts(secattr) $create_disposition $flags]
}


proc twapi::echo_server_accept {chan} {
    set ::twapi::echo_server_status connected
    fileevent $chan writable {}
}

proc twapi::echo_server {{name {\\.\pipe\twapiecho}} {timeout 20000}} {
    set timer [after $timeout "set ::twapi::echo_server_status timeout"]
    set echo_fd [::twapi::namedpipe_server $name]
    fconfigure $echo_fd -buffering line -translation crlf -eofchar {} -encoding utf-8
    fileevent $echo_fd writable [list ::twapi::echo_server_accept $echo_fd]
    vwait ::twapi::echo_server_status
    after cancel $timer
    set size 0
    if {$::twapi::echo_server_status eq "connected"} {
        while {1} {
            if {[gets $echo_fd line] >= 0} {
                puts $echo_fd $line
                if {$line eq "exit"} {
                    break
                }
                set size [string length $line]
            } else {
                puts stderr "Unexpected eof from echo client"
                break
            }
        }
    } else {
        puts stderr "echo_server: $::twapi::echo_server_status"
    }
    close $echo_fd
    return $size
}

proc twapi::echo_client {args} {
    array set opts [parseargs args {
        {name.arg {\\.\pipe\twapiecho}}
        {density.int 2}
        {limit.int 10000}
    }]

    set msgs 0
    set last 0
    set fd [twapi::namedpipe_client $opts(name)]
    fconfigure $fd -buffering line -translation crlf -eofchar {} -encoding utf-8
    for {set i 1} {$i < 100000} {incr i [expr {1+($i+1)/$opts(density)}]} {
        set request [string repeat X $i]
        puts $fd $request
        set response [gets $fd]
        if {$request ne $response} {
            puts "Mismatch in message of size $i"
        }
        incr msgs
        set last $i
    }
    puts $fd "exit"
    close $fd
    puts "Messages: $msgs, Last: $last"
}
