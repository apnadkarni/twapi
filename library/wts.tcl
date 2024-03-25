
# Copyright (c) 2021 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {
    variable _wts_session_monitors
    set _wts_session_monitors [dict create]
}

proc twapi::start_wts_session_monitor {script args} {
    variable _wts_session_monitors

    parseargs args {
        all
    } -maxleftover 0 -setvars]

    set script [lrange $script 0 end]; # Verify syntactically a list

    set id "wts#[TwapiId]"
    if {[dict size $_wts_session_monitors] == 0} {
        # No monitoring in progress. Start it
        # 0x2B1 -> WM_WTSSESSION_CHANGE
        Twapi_WTSRegisterSessionNotification $all
        _register_script_wm_handler 0x2B1 [list [namespace current]::_wts_session_change_handler] 0
    }

    dict set  _wts_session_monitors $id $script
    return $id
}


proc twapi::stop_wts_session_monitor {id} {
    variable _wts_session_monitors

    if {![dict exists $_wts_session_monitors $id]} {
        return
    }

    dict unset _wts_session_monitors $id
    if {[dict size $_wts_session_monitors] == 0} {
        # 0x2B1 -> WM_WTSSESSION_CHANGE
        _unregister_script_wm_handler 0x2B1 [list [namespace current]::_wts_session_handler]
        Twapi_WTSUnRegisterSessionNotification
    }
}

proc twapi::_wts_session_change_handler {msg change session_id msgpos ticks} {
    variable _wts_session_monitors

    if {[dict size $_wts_session_monitors] == 0} {
        return; # Not an error, could have deleted while already queued
    }

    dict for {id script} $_wts_session_monitors {
        set code [catch {uplevel #0 [linsert $script end $change $session_id]} msg]
        if {$code == 1} {
            # Error - put in background but we do not abort
            after 0 [list error $msg $::errorInfo $::errorCode]
        }
    }
    return
}
