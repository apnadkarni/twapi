#
# Copyright (c) 2009, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license


# Add a new user account
proc twapi::new_user {username args} {
    

    array set opts [parseargs args [list \
                                        system.arg \
                                        password.arg \
                                        comment.arg \
                                        [list priv.arg "user" [array names twapi::priv_level_map]] \
                                        home_dir.arg \
                                        script_path.arg \
                                       ] \
                        -nulldefault]

    if {$opts(priv) ne "user"} {
        error "Option -priv is deprecated and values other than 'user' are not allowed"
    }

    # 1 -> priv level 'user'. NetUserAdd mandates this as only allowed value
    NetUserAdd $opts(system) $username $opts(password) 1 \
        $opts(home_dir) $opts(comment) 0 $opts(script_path)


    # Backward compatibility - add to 'Users' local group
    # but only if -system is local
    if {$opts(system) eq "" ||
        ([info exists ::env(COMPUTERNAME)] &&
         [string equal -nocase $opts(system) $::env(COMPUTERNAME)])} {
        trap {
            _set_user_priv_level $username $opts(priv) -system $opts(system)
        } onerror {} {
            # Remove the previously created user account
            set ecode $errorCode
            set einfo $errorInfo
            catch {delete_user $username -system $opts(system)}
            error $errorResult $einfo $ecode
        }
    }
}


# Delete a user account
proc twapi::delete_user {username args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    # Remove the user from the LSA rights database.
    _delete_rights $username $opts(system)

    NetUserDel $opts(system) $username
}


# Define various functions to set various user account fields
foreach twapi::_field_ {
    {name LPWSTR 0}
    {password LPWSTR 1003}
    {home_dir LPWSTR 1006}
    {comment LPWSTR 1007}
    {script_path LPWSTR 1009}
    {full_name LPWSTR 1011}
    {country_code DWORD 1024}
    {profile LPWSTR 1052}
    {home_dir_drive LPWSTR 1053}
} {
    proc twapi::set_user_[lindex $::twapi::_field_ 0] {username fieldval args} "
        array set opts \[parseargs args {
            system.arg
        } -nulldefault \]
        Twapi_NetUserSetInfo[lindex $::twapi::_field_ 1] [lindex $::twapi::_field_ 2] \$opts(system) \$username \$fieldval"
}
unset twapi::_field_

# Set account expiry time
proc twapi::set_user_expiration {username time args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    if {![string is integer -strict $time]} {
        if {[string equal $time "never"]} {
            set time -1
        } else {
            set time [clock scan $time]
        }
    }
    Twapi_NetUserSetInfoDWORD 1017 $opts(system) $username $time
}

# Unlock a user account
proc twapi::unlock_user {username args} {
    # UF_LOCKOUT -> 0x10
    _change_user_info_flags $username 0x10 0 {*}$args
}

# Enable a user account
proc twapi::enable_user {username args} {
    # UF_ACCOUNTDISABLE -> 0x2
    _change_user_info_flags $username 0x2 0 {*}$args
}

# Disable a user account
proc twapi::disable_user {username args} {
    # UF_ACCOUNTDISABLE -> 0x2
    _change_user_info_flags $username 0x2 0x2 {*}$args
}


# Return the specified fields for a user account
proc twapi::get_user_account_info {account args} {

    # Define each option, the corresponding field, and the 
    # information level at which it is returned
    array set fields {
        comment 1
        password_expired 3
        full_name 2
        parms 2
        units_per_week 2
        primary_group_id 3
        status 1
        logon_server 2
        country_code 2
        home_dir 1
        password_age 1
        home_dir_drive 3
        num_logons 2
        acct_expires 2
        last_logon 2
        user_id 3
        usr_comment 2
        bad_pw_count 2
        code_page 2
        logon_hours 2
        workstations 2
        last_logoff 2
        name 0
        script_path 1
        profile 3
        max_storage 2
    }
    # Left out - auth_flags 2
    # Left out (always returned as NULL) - password {usri3_password 1}

    array set opts [parseargs args \
                        [concat [array names fields] \
                             [list sid local_groups global_groups system.arg all]] \
                        -nulldefault]

    if {$opts(all)} {
        set level 3
        set opts(local_groups) 1
        set opts(global_groups) 1
        set opts(sid) 1
    } else {
        # Based on specified fields, figure out what level info to ask for
        set level -1
        foreach {opt optval} [array get opts] {
            if {[info exists fields($opt)] &&
                $optval &&
                $fields($opt) > $level
            } {
                set level $fields($opt)
            }
        }                
    }
    
    array set result [list ]

    if {$level > -1} {
        array set data [NetUserGetInfo $opts(system) $account $level]
        # Extract the requested data
        foreach {opt optval} [array get opts] {
            if {[info exists fields($opt)] && ($optval || $opts(all))} {
                if {$opt eq "status"} {
                    set result(status) $data(flags)
                } else {
                    set result($opt) $data($opt)
                }
            }
        }

        # Map internal values to more friendly formats
        if {[info exists result(status)]} {
            # UF_LOCKOUT -> 0x10, UF_ACCOUNTDISABLE -> 0x2
            if {$result(status) & 0x2} {
                set result(status) "disabled"
            } elseif {$result(status) & 0x10} {
                set result(status) "locked"
            } else {
                set result(status) "enabled"
            }
        }

        if {[info exists result(logon_hours)]} {
            binary scan $result(logon_hours) b* result(logon_hours)
        }

        foreach time_field {acct_expires last_logon last_logoff} {
            if {[info exists result($time_field)]} {
                if {$result($time_field) == -1} {
                    set result($time_field) "never"
                } elseif {$result($time_field) == 0} {
                    set result($time_field) "unknown"
                }
            }
        }
    
    }

    if {$opts(local_groups)} {
        set result(local_groups) [kl_flatten [lindex [NetUserGetLocalGroups $opts(system) $account 0 0] 3] name]
    }

    if {$opts(global_groups)} {
        set result(global_groups) [kl_flatten [lindex [NetUserGetGroups $opts(system) $account 0] 3] name]
    }

    if {$opts(sid)} {
        set result(sid) [lookup_account_name $account -system $opts(system)]
    }

    return [_get_array_as_options result]
}

proc twapi::get_user_local_groups_recursive {account args} {
    array set opts [parseargs args {
        system.arg
    } -nulldefault -maxleftover 0]

    return [kl_flatten [lindex [NetUserGetLocalGroups $opts(system) [map_account_to_name $account -system $opts(system)] 0 1] 3] name]
}


# Set the specified fields for a user account
proc twapi::set_user_account_info {account args} {

    set notspecified "3kjafnq2or2034r12"; # Some junk

    # Define each option, the corresponding field, and the 
    # information level at which it is returned
    array set opts [parseargs args {
        {system.arg ""}
        comment.arg
        full_name.arg
        country_code.arg
        home_dir.arg
        home_dir.arg
        acct_expires.arg
        name.arg
        script_path.arg
        profile.arg
    }]

    # TBD - rewrite this to be atomic

    if {[info exists opts(comment)]} {
        set_user_comment $account $opts(comment) -system $opts(system)
    }

    if {[info exists opts(full_name)]} {
        set_user_full_name $account $opts(full_name) -system $opts(system)
    }

    if {[info exists opts(country_code)]} {
        set_user_country_code $account $opts(country_code) -system $opts(system)
    }

    if {[info exists opts(home_dir)]} {
        set_user_home_dir $account $opts(home_dir) -system $opts(system)
    }

    if {[info exists opts(home_dir_drive)]} {
        set_user_home_dir_drive $account $opts(home_dir_drive) -system $opts(system)
    }

    if {[info exists opts(acct_expires)]} {
        set_user_expiration $account $opts(acct_expires) -system $opts(system)
    }

    if {[info exists opts(name)]} {
        set_user_name $account $opts(name) -system $opts(system)
    }

    if {[info exists opts(script_path)]} {
        set_user_script_path $account $opts(script_path) -system $opts(system)
    }

    if {[info exists opts(profile)]} {
        set_user_profile $account $opts(profile) -system $opts(system)
    }
}
                    

proc twapi::get_global_group_info {name args} {
    array set opts [parseargs args {
        {system.arg ""}
        comment
        name
        members
        sid
        attributes
        all
    } -maxleftover 0]

    set result [list ]
    if {$opts(all) || $opts(sid)} {
        # Level 3 in NetGroupGetInfo would get us the SID as well but
        # sadly, Win2K does not support level 3 - TBD
        lappend result -sid [lookup_account_name $name -system $opts(system)]
    }
    if {$opts(all) || $opts(comment) || $opts(name) || $opts(attributes)} {
        set info [NetGroupGetInfo $opts(system) $name 2]
        if {$opts(all) || $opts(name)} {
            lappend result -name [kl_get $info name]
        }
        if {$opts(all) || $opts(comment)} {
            lappend result -comment [kl_get $info comment]
        }
        if {$opts(all) || $opts(attributes)} {
            lappend result -attributes [_map_token_group_attr [kl_get $info attributes]]
        }
    }
    if {$opts(all) || $opts(members)} {
        lappend result -members [get_global_group_members $name -system $opts(system)]
    }
    return $result
}

# Get info about a local or global group
proc twapi::get_local_group_info {name args} {
    array set opts [parseargs args {
        {system.arg ""}
        comment
        name
        members
        sid
        all
    } -maxleftover 0]

    set result [list ]
    if {$opts(all) || $opts(sid)} {
        lappend result -sid [lookup_account_name $name -system $opts(system)]
    }
    if {$opts(all) || $opts(comment) || $opts(name)} {
        set info [NetLocalGroupGetInfo $opts(system) $name 1]
        if {$opts(all) || $opts(name)} {
            lappend result -name [kl_get $info name]
        }
        if {$opts(all) || $opts(comment)} {
            lappend result -comment [kl_get $info comment]
        }
    }
    if {$opts(all) || $opts(members)} {
        lappend result -members [get_local_group_members $name -system $opts(system)]
    }
    return $result
}

# Get list of users on a system
proc twapi::get_users {args} {
    lappend args -filter 0; # Filter. TBD -allow user to specify filter
    return [_net_enum_helper NetUserEnum $args]
}

# Get list of global groups on a system
proc twapi::get_global_groups {args} {
    return [_net_enum_helper NetGroupEnum $args]
}

# Get list of local groups on a system
proc twapi::get_local_groups {args} {
    return [_net_enum_helper NetLocalGroupEnum $args]
    array set opts [parseargs args {system.arg} -nulldefault]
    return [NetLocalGroupEnum $opts(system)]
}

# Create a new global group
proc twapi::new_global_group {grpname args} {
    array set opts [parseargs args {
        system.arg
        comment.arg
    } -nulldefault]

    NetGroupAdd $opts(system) $grpname $opts(comment)
}

# Create a new local group
proc twapi::new_local_group {grpname args} {
    array set opts [parseargs args {
        system.arg
        comment.arg
    } -nulldefault]

    NetLocalGroupAdd $opts(system) $grpname $opts(comment)
}


# Delete a global group
proc twapi::delete_global_group {grpname args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    # Remove the group from the LSA rights database.
    _delete_rights $grpname $opts(system)

    NetGroupDel $opts(system) $grpname
}

# Delete a local group
proc twapi::delete_local_group {grpname args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    # Remove the group from the LSA rights database.
    _delete_rights $grpname $opts(system)

    NetLocalGroupDel $opts(system) $grpname
}


# Enumerate members of a global group
proc twapi::get_global_group_members {grpname args} {
    lappend args -preargs [list $grpname] -namelevel 1
    return [_net_enum_helper NetGroupGetUsers $args]
}

# Enumerate members of a local group
proc twapi::get_local_group_members {grpname args} {
    lappend args -preargs [list $grpname] -namelevel 1
    return [_net_enum_helper NetLocalGroupGetMembers $args]
}

# Add a user to a global group
proc twapi::add_user_to_global_group {grpname username args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    # No error if already member of the group
    trap {
        NetGroupAddUser $opts(system) $grpname $username
    } onerror {TWAPI_WIN32 1320} {
        # Ignore
    }
}


# Add a user to a local group
proc twapi::add_member_to_local_group {grpname username args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    # No error if already member of the group
    trap {
        Twapi_NetLocalGroupAddMember $opts(system) $grpname $username
    } onerror {TWAPI_WIN32 1378} {
        # Ignore
    }
}


# Remove a user from a global group
proc twapi::remove_user_from_global_group {grpname username args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    trap {
        NetGroupDelUser $opts(system) $grpname $username
    } onerror {TWAPI_WIN32 1321} {
        # Was not in group - ignore
    }
}


# Remove a user from a local group
proc twapi::remove_member_from_local_group {grpname username args} {
    array set opts [parseargs args {system.arg} -nulldefault]

    trap {
        Twapi_NetLocalGroupDelMember $opts(system) $grpname $username
    } onerror {TWAPI_WIN32 1377} {
        # Was not in group - ignore
    }
}

# Get a token for a user
proc twapi::open_user_token {username password args} {
    variable windefs

    array set opts [parseargs args {
        domain.arg
        {type.arg batch}
        {provider.arg default}
    } -nulldefault]

    set typedef "LOGON32_LOGON_[string toupper $opts(type)]"
    if {![info exists windefs($typedef)]} {
        error "Invalid value '$opts(type)' specified for -type option"
    }

    set providerdef "LOGON32_PROVIDER_[string toupper $opts(provider)]"
    if {![info exists windefs($typedef)]} {
        error "Invalid value '$opts(provider)' specified for -provider option"
    }
    
    # If username is of the form user@domain, then domain must not be specified
    # If username is not of the form user@domain, then domain is set to "."
    # if it is empty
    if {[regexp {^([^@]+)@(.+)} $username dummy user domain]} {
        if {[string length $opts(domain)] != 0} {
            error "The -domain option must not be specified when the username is in UPN format (user@domain)"
        }
    } else {
        if {[string length $opts(domain)] == 0} {
            set opts(domain) "."
        }
    }

    return [LogonUser $username $opts(domain) $password $windefs($typedef) $windefs($providerdef)]
}


# Impersonate a user given a token
proc twapi::impersonate_token {token} {
    ImpersonateLoggedOnUser $token
}


# Impersonate a user
proc twapi::impersonate_user {args} {
    set token [open_user_token {*}$args]
    trap {
        impersonate_token $token
    } finally {
        close_token $token
    }
}

# Impersonate a named pipe client
proc twapi::impersonate_namedpipe_client {chan} {
    set h [get_tcl_channel_handle $chan read]
    ImpersonateNamedPipeClient $h
}

# Revert to process token
proc twapi::revert_to_self {{opt ""}} {
    RevertToSelf
}


# Impersonate self
proc twapi::impersonate_self {level} {
    switch -exact -- $level {
        anonymous      { set level 0 }
        identification { set level 1 }
        impersonation  { set level 2 }
        delegation     { set level 3 }
        default {
            error "Invalid impersonation level $level"
        }
    }
    ImpersonateSelf $level
}

# Set a thread token - currently only for current thread
proc twapi::set_thread_token {token} {
    SetThreadToken NULL $token
}

# Reset a thread token - currently only for current thread
proc twapi::reset_thread_token {} {
    SetThreadToken NULL NULL
}

# Get a handle to a LSA policy
proc twapi::get_lsa_policy_handle {args} {
    array set opts [parseargs args {
        {system.arg ""}
        {access.arg policy_read}
    } -maxleftover 0]

    set access [_access_rights_to_mask $opts(access)]
    return [Twapi_LsaOpenPolicy $opts(system) $access]
}

# Close a LSA policy handle
proc twapi::close_lsa_policy_handle {h} {
    LsaClose $h
    return
}

# Get rights for an account
proc twapi::get_account_rights {account args} {
    array set opts [parseargs args {
        {system.arg ""}
    } -maxleftover 0]

    set sid [map_account_to_sid $account -system $opts(system)]

    trap {
        set lsah [get_lsa_policy_handle -system $opts(system) -access policy_lookup_names]
        return [Twapi_LsaEnumerateAccountRights $lsah $sid]
    } onerror {TWAPI_WIN32 2} {
        # No specific rights for this account
        return [list ]
    } finally {
        if {[info exists lsah]} {
            close_lsa_policy_handle $lsah
        }
    }
}

# Get accounts having a specific right
proc twapi::find_accounts_with_right {right args} {
    array set opts [parseargs args {
        {system.arg ""}
        name
    } -maxleftover 0]

    trap {
        set lsah [get_lsa_policy_handle \
                      -system $opts(system) \
                      -access {
                          policy_lookup_names
                          policy_view_local_information
                      }]
        set accounts [list ]
        foreach sid [Twapi_LsaEnumerateAccountsWithUserRight $lsah $right] {
            if {$opts(name)} {
                if {[catch {lappend accounts [lookup_account_sid $sid -system $opts(system)]}]} {
                    # No mapping for SID - can happen if account has been
                    # deleted but LSA policy not updated accordingly
                    lappend accounts $sid
                }
            } else {
                lappend accounts $sid
            }
        }
        return $accounts
    } onerror {TWAPI_WIN32 259} {
        # No accounts have this right
        return [list ]
    } finally {
        if {[info exists lsah]} {
            close_lsa_policy_handle $lsah
        }
    }

}

# Add/remove rights to an account
proc twapi::_modify_account_rights {operation account rights args} {
    set switches {
        system.arg
        handle.arg
    }    

    switch -exact -- $operation {
        add {
            # Nothing to do
        }
        remove {
            lappend switches all
        }
        default {
            error "Invalid operation '$operation' specified"
        }
    }

    array set opts [parseargs args $switches -maxleftover 0]

    if {[info exists opts(system)] && [info exists opts(handle)]} {
        error "Options -system and -handle may not be specified together"
    }

    if {[info exists opts(handle)]} {
        set lsah $opts(handle)
        set sid $account
    } else {
        if {![info exists opts(system)]} {
            set opts(system) ""
        }

        set sid [map_account_to_sid $account -system $opts(system)]
        # We need to open a policy handle ourselves. First try to open
        # with max privileges in case the account needs to be created
        # and then retry with lower privileges if that fails
        catch {
            set lsah [get_lsa_policy_handle \
                          -system $opts(system) \
                          -access {
                              policy_lookup_names
                              policy_create_account
                          }]
        }
        if {![info exists lsah]} {
            set lsah [get_lsa_policy_handle \
                          -system $opts(system) \
                          -access policy_lookup_names]
        }
    }

    trap {
        if {$operation == "add"} {
            LsaAddAccountRights $lsah $sid $rights
        } else {
            LsaRemoveAccountRights $lsah $sid $opts(all) $rights
        }
    } finally {
        # Close the handle if we opened it
        if {! [info exists opts(handle)]} {
            close_lsa_policy_handle $lsah
        }
    }
}

interp alias {} twapi::add_account_rights {} twapi::_modify_account_rights add
interp alias {} twapi::remove_account_rights {} twapi::_modify_account_rights remove

# Return list of logon sesionss
proc twapi::find_logon_sessions {args} {
    array set opts [parseargs args {
        user.arg
        type.arg
        tssession.arg
    } -maxleftover 0]

    set luids [LsaEnumerateLogonSessions]
    if {! ([info exists opts(user)] || [info exists opts(type)] ||
           [info exists opts(tssession)])} {
        return $luids
    }


    # Need to get the data for each session to see if it matches
    set result [list ]
    if {[info exists opts(user)]} {
        set sid [map_account_to_sid $opts(user)]
    }
    if {[info exists opts(type)]} {
        set logontypes [list ]
        foreach logontype $opts(type) {
            lappend logontypes [_logon_session_type_code $logontype]
        }
    }

    foreach luid $luids {
        trap {
            unset -nocomplain session
            array set session [LsaGetLogonSessionData $luid]

            # For the local system account, no data is returned on some
            # platforms
            if {[array size session] == 0} {
                set session(Sid) S-1-5-18; # SYSTEM
                set session(Session) 0
                set session(LogonType) 0
            }
            if {[info exists opts(user)] && $session(Sid) ne $sid} {
                continue;               # User id does not match
            }

            if {[info exists opts(type)] && [lsearch -exact $logontypes $session(LogonType)] < 0} {
                continue;               # Type does not match
            }

            if {[info exists opts(tssession)] && $session(Session) != $opts(tssession)} {
                continue;               # Term server session does not match
            }

            lappend result $luid

        } onerror {TWAPI_WIN32 1312} {
            # Session no longer exists. Just skip
            continue
        }
    }

    return $result
}


# Return data for a logon session
proc twapi::get_logon_session_info {luid args} {
    array set opts [parseargs args {
        all
        authpackage
        dnsdomain
        logondomain
        logonid
        logonserver
        logontime
        type
        sid
        user
        tssession
        userprincipal
    } -maxleftover 0]

    array set session [LsaGetLogonSessionData $luid]

    # Some fields may be missing on Win2K
    foreach fld {LogonServer DnsDomainName Upn} {
        if {![info exists session($fld)]} {
            set session($fld) ""
        }
    }

    array set result [list ]
    foreach {opt index} {
        authpackage AuthenticationPackage
        dnsdomain   DnsDomainName
        logondomain LogonDomain
        logonid     LogonId
        logonserver LogonServer
        logontime   LogonTime
        type        LogonType
        sid         Sid
        user        UserName
        tssession   Session
        userprincipal Upn
    } {
        if {$opts(all) || $opts($opt)} {
            set result(-$opt) $session($index)
        }
    }

    if {[info exists result(-type)]} {
        set result(-type) [_logon_session_type_symbol $result(-type)]
    }

    return [array get result]
}


# Set/reset the given bits in the usri3_flags field for a user account
# mask indicates the mask of bits to set. values indicates the values
# of those bits
proc twapi::_change_user_info_flags {username mask values args} {
    array set opts [parseargs args {
        system.arg
    } -nulldefault -maxleftover 0]

    # Get current flags
    array set data [NetUserGetInfo $opts(system) $username 1]

    # Turn off mask bits and write flags back
    set flags [expr {$data(flags) & (~ $mask)}]
    # Set the specified bits
    set flags [expr {$flags | ($values & $mask)}]

    # Write new flags back
    Twapi_NetUserSetInfoDWORD 1008 $opts(system) $username $flags
}

# Map impersonation level to symbol
proc twapi::_map_impersonation_level ilevel {
    switch -exact -- $ilevel {
        0 { return "anonymous" }
        1 { return "identification" }
        2 { return "impersonation" }
        3 { return "delegation" }
        default { return $ilevel }
    }
}

# Returns the logon session type value for a symbol
proc twapi::_logon_session_type_code {type} {
    # Type may be an integer or one of the strings below
    set code [lsearch -exact $::twapi::logon_session_type_map $type]
    if {$code >= 0} {
        return $code
    }

    if {![string is integer -strict $type]} {
        error "Invalid logon session type '$type' specified"
    }
    return $type
}

# Returns the logon session type symbol for an integer value
proc twapi::_logon_session_type_symbol {code} {
    set symbol [lindex $::twapi::logon_session_type_map $code]
    if {$symbol eq ""} {
        return $code
    } else {
        return $symbol
    }
}

proc twapi::_set_user_priv_level {username priv_level args} {

    array set opts [parseargs args {system.arg} -nulldefault]

    if {0} {
        # FOr some reason NetUserSetInfo cannot change priv level
        # Tried it separately with a simple C program. So this code
        # is commented out and we use group membership to achieve
        # the desired result
        # Note: - latest MSDN confirms above
        if {![info exists twapi::priv_level_map($priv_level)]} {
            error "Invalid privilege level value '$priv_level' specified. Must be one of [join [array names twapi::priv_level_map] ,]"
        }
        set priv $twapi::priv_level_map($priv_level)

        Twapi_NetUserSetInfo_priv $opts(system) $username $priv
    } else {
        # Don't hardcode group names - reverse map SID's instead for 
        # non-English systems. Also note that since
        # we might be lowering privilege level, we have to also
        # remove from higher privileged groups
        variable builtin_account_sids
        switch -exact -- $priv_level {
            guest {
                set outgroups {administrators users}
                set ingroup guests
            }
            user  {
                set outgroups {administrators}
                set ingroup users
            }
            admin {
                set outgroups {}
                set ingroup administrators
            }
            default {error "Invalid privilege level '$priv_level'. Must be one of 'guest', 'user' or 'admin'"}
        }
        # Remove from higher priv groups
        foreach outgroup $outgroups {
            # Get the potentially localized name of the group
            set group [lookup_account_sid $builtin_account_sids($outgroup) -system $opts(system)]
            # Catch since may not be member of that group
            catch {remove_member_from_local_group $group $username -system $opts(system)}
        }

        # Get the potentially localized name of the group to be added
        set group [lookup_account_sid $builtin_account_sids($ingroup) -system $opts(system)]
        add_member_to_local_group $group $username -system $opts(system)
    }
}
