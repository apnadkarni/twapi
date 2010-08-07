#
# Copyright (c) 2010, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Implementation of named pipes

proc twapi::namedpipe_server {name args} {
    set name [file nativename $name]

    # Only byte mode currently supported. Message mode does
    # not mesh well with Tcl channel infrastructure.
    # readmode.arg
    # writemode.arg

    array set opts [twapi::parseargs args {
        {access.arg {read write}}
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

    set open_mode [twapi::_parse_symbolic_bitmask $opts(access) {read 1 write 2}]
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
        {access.arg {read write}}
        impersonationlevel.arg
        {impersonateeffectiveonly.bool false}
        {impersonatecontexttracking.bool false}
    } -maxleftover 0]

    # FILE_READ_DATA              0x00000001
    # FILE_WRITE_DATA             0x00000002
    # Note - use _parse_symbolic_bitmask because we allow user to specify
    # numeric masks as well
    set desired_access [twapi::_parse_symbolic_bitmask $opts(access) {
        read  1
        write 2
    }]
        
    set flags 0
    if {[info exists opts(impersonationlevel)]} {
        switch -exact -- $opts(impersonationlevel) {
            anonymous      { set flags 0x00100000 }
            identification { set flags 0x00110000 }
            impersonation  { set flags 0x00120000 }
            delegation     { set flags 0x00130000 }
            default {
                # ERROR_BAD_IMPERSONATION_LEVEL
                win32_error 1346 "Invalid impersonation level '$opts(impersonationlevel)'."
            }
        }
        if {$opts(impersonateeffectiveonly)} {
            set flags [expr {$flags | 0x00080000}]
        }
        if {$opts(impersonatecontexttracking)} {
            set flags [expr {$flags | 0x00040000}]
        }
    }

    set share_mode 0;           # Share none
    set secattr {};             # At some point use this for "inherit" flag
    set create_disposition 3;   # OPEN_EXISTING
    return [twapi::Twapi_NPipeClient $name $desired_access $share_mode \
                $secattr $create_disposition $flags]
}

