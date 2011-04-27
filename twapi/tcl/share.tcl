#
# Copyright (c) 2003, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {}

# TBD - is there a Tcl wrapper around NetShareCheck?

# Create a network share
proc twapi::new_share {sharename path args} {
    array set opts [parseargs args {
        {system.arg ""}
        {type.arg "file"}
        {comment.arg ""}
        {max_conn.int -1}
        secd.arg
    } -maxleftover 0]

    # If no security descriptor specified, default to "Everyone,
    # read permission". Levaing it empty will give everyone all permissions
    # which is probably not a good idea!
    if {![info exists opts(secd)]} {
        set opts(secd) [new_security_descriptor -dacl [new_acl [list [new_ace allow S-1-1-0 1179817]]]]
    }
    
    NetShareAdd $opts(system) \
        $sharename \
        [_share_type_symbols_to_code $opts(type)] \
        $opts(comment) \
        $opts(max_conn) \
        [file nativename $path] \
        $opts(secd)
}

# Delete a network share
proc twapi::delete_share {sharename args} {
    array set opts [parseargs args {system.arg} -nulldefault]
    NetShareDel $opts(system) $sharename 0
}

# Enumerate network shares
proc twapi::get_shares {args} {
    variable windefs

    array set opts [parseargs args {
        {system.arg ""}
        {type.arg ""}
        excludespecial
        level.int
    } -maxleftover 0]

    if {$opts(type) != ""} {
        set type_filter [_share_type_symbols_to_code $opts(type) 1]
    }

    if {[info exists opts(level)] && $opts(level) > 0} {
        set level $opts(level)
    } else {
        # Either -level not specified or specified as 0
        # We need at least level 1 to filter on type
        set level 1
    }

    set raw_data [_net_enum_helper NetShareEnum -system $opts(system) -level $level]
    set shares [list ]
    foreach elem $raw_data {
        array set share $elem
        set special [expr {$share(type) & ($windefs(STYPE_SPECIAL) | $windefs(STYPE_TEMPORARY))}]
        if {$special && $opts(excludespecial)} {
            continue
        }
        # We need the special cast to int because else operands get promoted
        # to 64 bits as the hex is treated as an unsigned value
        set share(type) [expr {int($share(type) & ~ $special)}]
        if {[info exists type_filter] && $share(type) != $type_filter} {
            continue
        }
        if {[info exists opts(level)]} {
            if {$opts(level) == 0} {
                # We had bumped level to 1 so only extract the level 0 field
                lappend shares [list netname $share(netname)]
            } else {
                lappend shares $elem
            }
        } else {
            lappend shares $share(netname)
        }
    }

    return $shares
}


# Get details about a share
proc twapi::get_share_info {sharename args} {
    array set opts [parseargs args {
        system.arg
        all
        name
        type
        path
        comment
        max_conn
        current_conn
        secd
    } -nulldefault]

    if {$opts(all)} {
        foreach opt {name type path comment max_conn current_conn secd} {
            set opts($opt) 1
        }
    }

    set level 0

    if {$opts(name) || $opts(type) || $opts(comment)} {
        set level 1
    }

    if {$opts(max_conn) || $opts(current_conn) || $opts(path)} {
        set level 2
    }

    if {$opts(secd)} {
        set level 502
    }

    if {! $level} {
        return
    }

    array set shareinfo [NetShareGetInfo $opts(system) $sharename $level]

    set result [list ]
    if {$opts(name)} {
        lappend result -name $shareinfo(netname)
    }
    if {$opts(type)} {
        lappend result -type [_share_type_code_to_symbols $shareinfo(type)]
    }
    if {$opts(comment)} {
        lappend result -comment $shareinfo(remark)
    }
    if {$opts(max_conn)} {
        lappend result -max_conn $shareinfo(max_uses)
    }
    if {$opts(current_conn)} {
        lappend result -current_conn $shareinfo(current_uses)
    }
    if {$opts(path)} {
        lappend result -path $shareinfo(path)
    }
    if {$opts(secd)} {
        lappend result -secd $shareinfo(security_descriptor)
    }
    
    return $result
}


# Set a share configuration
proc twapi::set_share_info {sharename args} {
    array set opts [parseargs args {
        {system.arg ""}
        comment.arg
        max_conn.int
        secd.arg
    }]

    # First get the current config so we can change specified fields
    # and write back
    array set shareinfo [get_share_info $sharename -system $opts(system) \
                             -comment -max_conn -secd]
    foreach field {comment max_conn secd} {
        if {[info exists opts($field)]} {
            set shareinfo(-$field) $opts($field)
        }
    }

    NetShareSetInfo $opts(system) $sharename $shareinfo(-comment) \
        $shareinfo(-max_conn) $shareinfo(-secd)
}


# Get list of remote shares
interp alias {} twapi::get_connected_shares {} twapi::get_client_shares
proc twapi::get_client_shares {args} {
    array set opts [parseargs args {
        {system.arg ""}
        level.int
    } -maxleftover 0]

    if {[info exists opts(level)]} {
        set shares {}
        foreach share [_net_enum_helper NetUseEnum -system $opts(system) -level $opts(level)] {
            lappend shares [_map_USE_INFO $share]
        }
        return $shares
    }

    # For backwards compatibility, if no level specified, returned
    # format is slightly different.
    set cshares {}
    foreach elem [_net_enum_helper NetUseEnum -system $opts(system) -level 0] {
        lappend cshares [list [kl_get $elem local] [kl_get $elem remote]]
    }
    return $cshares
}


# Connect to a share
proc twapi::connect_share {remoteshare args} {
    array set opts [parseargs args {
        {type.arg  "disk"} 
        localdevice.arg
        provider.arg
        password.arg
        nopassword
        defaultpassword
        user.arg
        {window.arg 0}
        {interactive {} 0x8}
        {prompt      {} 0x10}
        {updateprofile {} 0x1}
        {commandline {} 0x800}
    } -nulldefault]

    set flags 0

    switch -exact -- $opts(type) {
        "any"       {set type 0}
        "disk"      -
        "file"      {set type 1}
        "printer"   {set type 2}
        default {
            error "Invalid network share type '$opts(type)'"
        }
    }

    # localdevice - "" means no local device, * means pick any, otherwise
    # it's a local device to be mapped
    if {$opts(localdevice) == "*"} {
        set opts(localdevice) ""
        setbits flags 0x80;             # CONNECT_REDIRECT
    }

    if {$opts(defaultpassword) && $opts(nopassword)} {
        error "Options -defaultpassword and -nopassword may not be used together"
    }
    if {$opts(nopassword)} {
        set opts(password) ""
        set ignore_password 1
    } else {
        set ignore_password 0
        if {$opts(defaultpassword)} {
            set opts(password) ""
        }
    }

    set flags [expr {$flags | $opts(interactive) | $opts(prompt) |
                     $opts(updateprofile) | $opts(commandline)}]

    return [Twapi_WNetUseConnection $opts(window) $type $opts(localdevice) \
                $remoteshare $opts(provider) $opts(user) $ignore_password \
                $opts(password) $flags]
}

# Disconnects an existing share
proc twapi::disconnect_share {sharename args} {
    array set opts [parseargs args {updateprofile force}]

    set flags [expr {$opts(updateprofile) ? 0x1 : 0}]
    WNetCancelConnection2 $sharename $flags $opts(force)
}


# Get information about a connected share
proc twapi::get_client_share_info {sharename args} {

    if {$sharename eq ""} {
        error "A share name cannot be the empty string"
    }

    # We have to use a combination of NetUseGetInfo and 
    # WNetGetResourceInformation as neither gives us the full information
    # THe former takes the local device name if there is one and will
    # only accept a UNC if there is an entry for the UNC with
    # no local device mapped. The latter
    # always wants the UNC. So we need to figure out exactly if there
    # is a local device mapped to the sharename or not
    # TBD _ see if this is really the case. Also, NetUse only works with
    # LANMAN, not WebDAV. So see if there is a way to only use WNet*
    # variants
    
    # There may be multiple entries for the same UNC
    # If there is an entry for the UNC with no device mapped, select
    # that else select any of the local devices mapped to it
    # TBD - any better way of finding out a mapping than calling
    # get_client_shares?
    foreach elem [get_client_shares] {
        lassign $elem elem_device elem_unc
        if {[string equal -nocase $sharename $elem_unc]} {
            if {$elem_device eq ""} {
                # Found an entry without a local device. Use it
                set unc $elem_unc
                unset -nocomplain local; # In case we found a match earlier
                break
            } else {
                # Found a matching device
                set local $elem_device
                set unc $elem_unc
                # Keep looping in case we find an entry with no local device
                # (which we will prefer)
            }
        } else {
            # See if the sharename is actually a local device name
            if {[string equal -nocase [string trimright $elem_device :] [string trimright $sharename :]]} {
                # Device name matches. Use it
                set local $elem_device
                set unc $elem_unc
                break
            }
        }
    }

    if {![info exists unc]} {
        win32_error 2250 "Share '$sharename' not found."
    }

    # At this point $unc is the UNC form of the share and
    # $local is either undefined or the local mapped device if there is one

    array set opts [parseargs args {
        user
        localdevice
        remoteshare
        status
        type
        opencount
        usecount
        domain
        provider
        comment
        all
    } -maxleftover 0]


    # Call Twapi_NetGetInfo always to get status. If we are not connected,
    # we will not call WNetGetResourceInformation as that will time out
    if {[info exists local]} {
        array set shareinfo [_map_USE_INFO [NetUseGetInfo "" $local 2]]
    } else {
        array set shareinfo [_map_USE_INFO [NetUseGetInfo "" $unc 2]]
    }

    if {$opts(all) || $opts(comment) || $opts(provider)} {
        # Only get this information if we are connected
        if {$shareinfo(-status) eq "connected"} {
            array set wnetinfo [lindex [Twapi_WNetGetResourceInformation $unc "" 0] 0]
            set shareinfo(-comment) $wnetinfo(lpComment)
            set shareinfo(-provider) $wnetinfo(lpProvider)
        } else {
            set shareinfo(-comment) ""
            set shareinfo(-provider) ""
        }
    }

    if {$opts(all)} {
        return [array get shareinfo]
    }

    # Get rid of unwanted fields
    foreach opt {
        user
        localdevice
        remoteshare
        status
        type
        opencount
        usecount
        domain
        provider
        comment
    } {
        if {! $opts($opt)} {
            if {[info exists shareinfo(-$opt)]} {
                unset shareinfo(-$opt)
            }
        }
    }

    return [array get shareinfo]
}


# Enumerate sessions
proc twapi::find_lm_sessions args {
    array set opts [parseargs args {
        all
        {client.arg ""}
        {system.arg ""}
        {user.arg ""}
        transport
        clientname
        username
        clienttype
        opencount
        idleseconds
        activeseconds
        attrs
    } -maxleftover 0]

    set level [_calc_minimum_session_info_level opts]
    if {![min_os_version 5]} {
        # System name is specified. If NT, make sure it is UNC form
        set opts(system) [_make_unc_computername $opts(system)]
    }
    
    # On all platforms, client must be in UNC format
    set opts(client) [_make_unc_computername $opts(client)]

    trap {
        set sessions [_net_enum_helper NetSessionEnum -system $opts(system) -preargs [list $opts(client) $opts(user)] -level $level]
    } onerror {TWAPI_WIN32 2312} {
        # No session matching the specified client
        return [list ]
    } onerror {TWAPI_WIN32 2221} {
        # No session matching the user
        return [list ]
    }

    set retval [list ]
    foreach sess $sessions {
        lappend retval [_format_lm_session $sess opts]
    }

    return $retval
}


# Get information about a session 
proc twapi::get_lm_session_info {client user args} {
    array set opts [parseargs args {
        all
        {system.arg ""}
        transport
        clientname
        username
        clienttype
        opencount
        idleseconds
        activeseconds
        attrs
    } -maxleftover 0]

    set level [_calc_minimum_session_info_level opts]
    if {$level == -1} {
        # No data requested so return empty list
        return [list ]
    }

    if {![min_os_version 5]} {
        # System name is specified. If NT, make sure it is UNC form
        set opts(system) [_make_unc_computername $opts(system)]
    }
    
    # On all platforms, client must be in UNC format
    set client [_make_unc_computername $client]

    # Note an error is generated if no matching session exists
    set sess [NetSessionGetInfo $opts(system) $client $user $level]

    return [_format_lm_session $sess opts]
}

# Delete sessions
proc twapi::end_lm_sessions args {
    array set opts [parseargs args {
        {client.arg ""}
        {system.arg ""}
        {user.arg ""}
    } -maxleftover 0]

    if {![min_os_version 5]} {
        # System name is specified. If NT, make sure it is UNC form
        set opts(system) [_make_unc_computername $opts(system)]
    }

    if {$opts(client) eq "" && $opts(user) eq ""} {
        win32_error 87 "At least one of -client and -user must be specified."
    }

    # On all platforms, client must be in UNC format
    set opts(client) [_make_unc_computername $opts(client)]

    trap {
        NetSessionDel $opts(system) $opts(client) $opts(user)
    } onerror {TWAPI_WIN32 2312} {
        # No session matching the specified client - ignore error
    } onerror {TWAPI_WIN32 2221} {
        # No session matching the user - ignore error
    }
    return
}

# Enumerate open files
proc twapi::find_lm_open_files args {
    array set opts [parseargs args {
        {basepath.arg ""}
        {system.arg ""}
        {user.arg ""}
        all
        permissions
        id
        lockcount
        path
        username
    } -maxleftover 0]

    if {![min_os_version 5]} {
        # System name is specified. If NT, make sure it is UNC form
        set opts(system) [_make_unc_computername $opts(system)]
    }
    
    set level 3
    if {! ($opts(all) || $opts(permissions) || $opts(lockcount) ||
           $opts(path) || $opts(username))} {
        # Only id's required
        set level 2
    }

    # TBD - change to use -resume option to _net_enum_helper as there
    # might be a lot of files
    trap {
        set files [_net_enum_helper NetFileEnum -system $opts(system) -preargs [list [file nativename $opts(basepath)] $opts(user)] -level $level]
    } onerror {TWAPI_WIN32 2221} {
        # No files matching the user
        return [list ]
    }

    set retval [list ]
    foreach file $files {
        lappend retval [_format_lm_open_file $file opts]
    }

    return $retval
}

# Get information about an open LM file
proc twapi::get_lm_open_file_info {fid args} {
    array set opts [parseargs args {
        {system.arg ""}
        all
        permissions
        id
        lockcount
        path
        username
    } -maxleftover 0]

    # System name is specified. If NT, make sure it is UNC form
    if {![min_os_version 5]} {
        set opts(system) [_make_unc_computername $opts(system)]
    }
    
    set level 3
    if {! ($opts(all) || $opts(permissions) || $opts(lockcount) ||
           $opts(path) || $opts(username))} {
        # Only id's required. We actually already have this but don't
        # return it since we want to go ahead and make the call in case
        # the id does not exist
        set level 2
    }

    return [_format_lm_open_file [NetFileGetInfo $opts(system) $fid $level] opts]
}

# Close an open LM file
proc twapi::close_lm_open_file {fid args} {
    array set opts [parseargs args {
        {system.arg ""}
    } -maxleftover 0]
    trap {
        NetFileClose $opts(system) $fid
    } onerror {TWAPI_WIN32 2314} {
        # No such fid. Ignore, perhaps it was closed in the meanwhile
    }
}


# Enumerate open connections
proc twapi::find_lm_connections args {
    array set opts [parseargs args {
        client.arg
        {system.arg ""}
        share.arg
        all
        id
        type
        opencount
        usercount
        activeseconds
        username
        clientname
        sharename
    } -maxleftover 0]

    if {![min_os_version 5]} {
        # System name is specified. If NT, make sure it is UNC form
        set opts(system) [_make_unc_computername $opts(system)]
    }
    
    if {! ([info exists opts(client)] || [info exists opts(share)])} {
        win32_error 87 "Must specify either -client or -share option."
    }

    if {[info exists opts(client)] && [info exists opts(share)]} {
        win32_error 87 "Must not specify both -client and -share options."
    }

    if {[info exists opts(client)]} {
        set qualifier [_make_unc_computername $opts(client)]
    } else {
        set qualifier $opts(share)
    }

    set level 1
    if {! ($opts(all) || $opts(type) || $opts(opencount) ||
           $opts(usercount) || $opts(username) ||
           $opts(activeseconds) || $opts(clientname) || $opts(sharename))} {
        # Only id's required
        set level 0
    }

    # TBD - change to use -resume option to _net_enum_helper since
    # there might be a log of connections
    set conns [_net_enum_helper NetConnectionEnum -system $opts(system) -preargs [list $qualifier] -level $level]
    set retval [list ]
    foreach conn $conns {
        set item [list ]
        foreach {opt fld} {
            id            id
            opencount     num_opens
            usercount     num_users
            activeseconds time
            username      username
        } {
            if {$opts(all) || $opts($opt)} {
                lappend item -$opt [kl_get $conn $fld]
            }
        }
        if {$opts(all) || $opts(type)} {
            lappend item -type [_share_type_code_to_symbols [kl_get $conn type]]
        }
        # What's returned in the netname field depends on what we
        # passed as the qualifier
        if {$opts(all) || $opts(clientname) || $opts(sharename)} {
            if {[info exists opts(client)]} {
                set sharename [kl_get $conn netname]
                set clientname [_make_unc_computername $opts(client)]
            } else {
                set sharename $opts(share)
                set clientname [_make_unc_computername [kl_get $conn netname]]
            }
            if {$opts(all) || $opts(clientname)} {
                lappend item -clientname $clientname
            }
            if {$opts(all) || $opts(sharename)} {
                lappend item -sharename $sharename
            }
        }
        lappend retval $item
    }

    return $retval
}


################################################################
# Utility functions

# Common code to figure out what SESSION_INFO level is required
# for the specified set of requested fields. v_opts is name
# of array indicating which fields are required
proc twapi::_calc_minimum_session_info_level {v_opts} {
    upvar $v_opts opts

    # Set the information level requested based on options specified.
    # We set the level to the one that requires the lowest possible
    # privilege level and still includes the data requested.
    if {$opts(all) || $opts(transport)} {
        return 502
    } elseif {$opts(clienttype)} {
        return 2
    } elseif {$opts(opencount) || $opts(attrs)} {
        return 1
    } elseif {$opts(clientname) || $opts(username) ||
        $opts(idleseconds) || $opts(activeseconds)} {
        return 10
    } else {
        return 0
    }
}

# Common code to format a session record. v_opts is name of array
# that controls which fields are returned
proc twapi::_format_lm_session {sess v_opts} {
    upvar $v_opts opts

    set retval [list ]
    foreach {opt fld} {
        transport     transport
        username      username
        opencount     num_opens
        idleseconds   idle_time
        activeseconds time
        clienttype    cltype_name
    } {
        if {$opts(all) || $opts($opt)} {
            lappend retval -$opt [kl_get $sess $fld]
        }
    }
    if {$opts(all) || $opts(clientname)} {
        # Since clientname is always required to be in UNC on input
        # also pass it back in UNC format
        lappend retval -clientname [_make_unc_computername [kl_get $sess cname]]
    }
    if {$opts(all) || $opts(attrs)} {
        set attrs [list ]
        set flags [kl_get $sess user_flags]
        if {$flags & 1} {
            lappend attrs guest
        }
        if {$flags & 2} {
            lappend attrs noencryption
        }
        lappend retval -attrs $attrs
    }
    return $retval
}

# Common code to format a lm open file record. v_opts is name of array
# that controls which fields are returned
proc twapi::_format_lm_open_file {file v_opts} {
    upvar $v_opts opts

    set retval [list ]
    foreach {opt fld} {
        id          id
        lockcount   num_locks
        path        pathname
        username    username
    } {
        if {$opts(all) || $opts($opt)} {
            lappend retval -$opt [kl_get $file $fld]
        }
    }

    if {$opts(all) || $opts(permissions)} {
        set permissions [list ]
        set perms [kl_get $file permissions]
        foreach {flag perm} {1 read 2 write 4 create} {
            if {$perms & $flag} {
                lappend permissions $perm
            }
        }
        lappend retval -permissions $permissions
    }

    return $retval
}

# NOTE: THIS ONLY MAPS FOR THE Net* functions, NOT THE WNet*
proc twapi::_share_type_symbols_to_code {typesyms {basetypeonly 0}} {
    variable windefs

    switch -exact -- [lindex $typesyms 0] {
        file    { set code $windefs(STYPE_DISKTREE) }
        printer { set code $windefs(STYPE_PRINTQ) }
        device  { set code $windefs(STYPE_DEVICE) }
        ipc     { set code $windefs(STYPE_IPC) }
        default {
            error "Unknown type network share type symbol [lindex $typesyms 0]"
        }
    }

    if {$basetypeonly} {
        return $code
    }

    set special 0
    foreach sym [lrange $typesyms 1 end] {
        switch -exact -- $sym {
            special   { setbits special $windefs(STYPE_SPECIAL) }
            temporary { setbits special $windefs(STYPE_TEMPORARY) }
            file    -
            printer -
            device  -
            ipc     {
                error "Base share type symbol '$sym' cannot be used as a share attribute type"
            }
            default {
                error "Unknown type network share type symbol '$sym'"
            }
        }
    }

    return [expr {$code | $special}]
}


# First element is always the base type of the share
# NOTE: THIS ONLY MAPS FOR THE Net* functions, NOT THE WNet*
proc twapi::_share_type_code_to_symbols {type} {
    variable windefs


    set special [expr {$type & ($windefs(STYPE_SPECIAL) | $windefs(STYPE_TEMPORARY))}]

    # We need the special cast to int because else operands get promoted
    # to 64 bits as the hex is treated as an unsigned value
    switch -exact -- [expr {int($type & ~ $special)}] \
        [list \
             $windefs(STYPE_DISKTREE) {set sym "file"} \
             $windefs(STYPE_PRINTQ)   {set sym "printer"} \
             $windefs(STYPE_DEVICE)   {set sym "device"} \
             $windefs(STYPE_IPC)      {set sym "ipc"} \
             default                  {set sym $type}
            ]

    set typesyms [list $sym]

    if {$special & $windefs(STYPE_SPECIAL)} {
        lappend typesyms special
    }

    if {$special & $windefs(STYPE_TEMPORARY)} {
        lappend typesyms temporary
    }
    
    return $typesyms
}

# Make sure a computer name is in unc format unless it is an empty
# string (local computer)
proc twapi::_make_unc_computername {name} {
    if {$name eq ""} {
        return ""
    } else {
        return "\\\\[string trimleft $name \\]"
    }
}

# Maps a USE_INFO struct as returned from C level
proc twapi::_map_USE_INFO {useinfo} {
    array set shareinfo $useinfo

    array set result {}
    foreach {opt field} {
        -user           username
        -localdevice    local
        -remoteshare    remote
        -usecount       usecount
        -opencount      refcount
        -status         status
        -type           asg_type
        -domain         domainname
    } {
        if {[info exists shareinfo($field)]} {
            set result($opt) $shareinfo($field)
        }
    }

    # Map values to symbols

    if {[info exists result(-status)]} {
        # Map code 0-5
        set temp [lindex {connected paused lostsession disconnected networkerror connecting reconnecting} $result(-status)]
        if {$temp ne ""} {
            set result(-status) $temp
        } else {
            set result(-status) "unknown"
        }
    }

    if {[info exists result(-type)]} {
        set temp [lindex {file printer char ipc} $result(-type)]
        if {$temp ne ""} {
            set result(-type) $temp
        } else {
            set result(-type) "unknown"
        }
    }

    return [array get result]
}
