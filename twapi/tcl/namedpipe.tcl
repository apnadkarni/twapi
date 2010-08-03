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
        return [twapi::_namedpipe_server $name {*}$args]
    } else {
        return [twapi::_namedpipe_client $name {*}$args]
    }
}

proc twapi::_namedpipe_server {name args} {
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
        {secattr.arg {}}
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

    set options [dict create \
                     open_mode $open_mode \
                     pipe_mode $pipe_mode \
                     max_instances $opts(maxinstances) \
                     timeout $opts(timeout) \
                     secattr $opts(secattr)]

    return [twapi::Twapi_NPipeServer $name $open_mode $pipe_mode \
                $opts(maxinstances) 4000 4000 $opts(timeout) \
                $opts(secattr)]
}
