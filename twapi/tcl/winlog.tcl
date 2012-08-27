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

    array set opts [parseargs args {
        {system.arg ""}
        channel.arg
        file.arg
        {authtype.arg 0}
        {direction.arg forward {forward backward}}
    } -maxleftover 0]

    if {[info exists opts(file)] &&
        $opts(system) ne "" || [info exists opts(channel)]} {
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
        # Note this closes $hq as well
        evt_close [dict get $_winlog_handles $hq]
    } else {
        eventlog_close $hq
    }

    dict unset _winlog_handles $hq
}

if {[twapi::min_os_version 6]} {

    proc twapi::winlog_read {hq} {
        # TBD - is 10 an appropriate number of events to read?
        return [evt_next $hq -timeout 0 -count 10]
    }

} else {
    proc twapi::winlog_read {hq} {
        variable _winlog_handles
        return [eventlog_read $hq -direction [dict get $_winlog_handles $hq direction]]
    }


}

proc twapi::winlog_decode_events {events {langid 0}} {
    set result {}
    if {[min_os_version 6]} {
        foreach evh $events {
            set ev {}
            lappend ev -message [evt_format_message $evh -lcid $langid]

        }
    } else {
        foreach ev $events {
            dict set ev -task [eventlog_format_category $ev -langid $langid]
            dict set ev -message [eventlog_format_message $ev -langid $langid]
            lappend result $ev
        }
    }

    return $result
}
