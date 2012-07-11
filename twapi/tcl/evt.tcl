#
# Copyright (c) 2012, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Event log handling for Vista and later

namespace eval twapi {}

proc twapi::evt_open_session {server args} {
    array set opts [parseargs args {
        user.arg
        domain.arg
        {password.arg ""}
        {authtype.arg 0}
    } -maxleftover 0]

    if {![string is integer -strict $opts(authtype)]} {
        set opts(authtype) [dict get {default 0 negotiate 1 kerberos 2 ntlm 3} [string tolower $opts(authtype)]]
    }

    if {![info exists opts(user)]} {
        set opts(user) $::env(USERNAME)
    }

    if {![info exists opts(domain)]} {
        set opts(domain) $::env(USERDOMAIN)
    }

    return [EvtOpenSession 1 [list $server $opts(user) $opts(domain) $opts(password) $opts(authtype)] 0 0]
}

proc twapi::evt_close {hevt} {
    return [EvtClose $hevt]
}

proc twapi::evt_channels {{hevtsess NULL}} {
    set chnames {}
    set hevt [EvtOpenChannelEnum $hevtsess 0]
    trap {
        while {[set chname [EvtNextChannelPath $hevt]] ne ""} {
            lappend chnames $chname
        }
    } finally {
        EvtClose $hevt
    }

    return $chnames
}
