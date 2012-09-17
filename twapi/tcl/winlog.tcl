#
# Copyright (c) 2012, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Routines to unify old and new Windows event log APIs

namespace eval twapi {
    # Dictionary to map eventlog consumer handles to various related info
    # The primary key is the read handle to the event channel/source.
    # Nested keys depend on OS version
    variable _winlog_handles
}

proc twapi::winlog_open {args} {
    variable _winlog_handles

    # TBD - document -authtype
    array set opts [parseargs args {
        {system.arg ""}
        channel.arg
        file.arg
        {authtype.arg 0}
        {direction.arg forward {forward backward}}
    } -maxleftover 0]

    if {[info exists opts(file)] &&
        ($opts(system) ne "" || [info exists opts(channel)])} {
        error "Option '-file' cannot be used with '-channel' or '-system'"
    } else {
        if {![info exists opts(channel)]} {
            set opts(channel) "Application"
        }
    }
    
    if {[min_os_version 6]} {
        # Use new Vista APIs
        if {[info exists opts(file)]} {
            set hsess NULL
            set hq [evt_query -file $opts(file) -ignorequeryerrors]
        } else {
            if {$opts(system) eq ""} {
                set hsess [twapi::evt_local_session]
            } else {
                set hsess [evt_open_session $opts(system) -authtype $opts(authtype)]
            }
            set hq [evt_query -session $hsess -channel $opts(channel) -ignorequeryerrors -direction $opts(direction)]
        }
        
        dict set _winlog_handles $hq session $hsess
    } else {
        if {[info exists opts(file)]} {
            set hq [eventlog_open -file $opts(file)]
        } else {
            set hq [eventlog_open -system $opts(system) -source $opts(channel)]
        }
        dict set _winlog_handles $hq direction $opts(direction)
    }
    return $hq
}

proc twapi::winlog_close {hq} {
    variable _winlog_handles

    if {! [dict exists $_winlog_handles $hq]} {
        error "Invalid event consumer handler '$hq'"
    }

    if {[min_os_version 6]} {
        set hsess [dict get $_winlog_handles $hq session]
        if {![pointer_null? $hsess]} {
            # Note this closes $hq as well
            evt_close $hsess
        }
    } else {
        eventlog_close $hq
    }

    dict unset _winlog_handles $hq
    return
}

if {[twapi::min_os_version 6]} {

    proc twapi::winlog_read {hq {lcid 0}} {
        set result {}
        # TBD - is 10 an appropriate number of events to read?
        foreach hevt [evt_next $hq -timeout 0 -count 10] {
            lappend result [evt_decode_event $hevt -lcid $lcid -ignorestring "" -message -levelname -taskname]
            evt_close $hevt
        }
        return $result
    }

} else {

    proc twapi::winlog_read {hq {langid 0}} {
        variable _winlog_handles
        set result {}
        foreach evl [eventlog_read $hq -direction [dict get $_winlog_handles $hq direction]] {
            lappend result \
                [list \
                     -taskname [eventlog_format_category $evl -langid $langid] \
                     -message [eventlog_format_message $evl -langid $langid] \
                     -providername [dict get $evl -source] \
                     -eventid [dict get $evl -eventid] \
                     -level [dict get $evl -level] \
                     -levelname [dict get $evl -type] \
                     -eventrecordid [dict get $evl -recnum] \
                     -computer [dict get $evl -system] \
                     -timecreated [secs_since_1970_to_large_system_time [dict get $evl -timewritten]]]
        }
        return $result
    }
}

proc twapi::_winlog_dump {{channel Application} {fd stdout}} {
    set hevl [winlog_open -channel $channel]
    while {[llength [set events [winlog_read $hevl]]]} {
        # print out each record
        foreach ev $events {
            puts $fd "[dict get $ev -timecreated] [dict get $ev -providername]: [dict get $ev -message]"
        }
    }
    winlog_close $hevl
}
