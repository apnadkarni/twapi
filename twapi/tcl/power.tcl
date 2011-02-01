#
# Copyright (c) 2003-2010 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Suspend system
proc twapi::suspend_system {args} {
    array set opts [parseargs args {
        {state.arg standby {standby hibernate}}
        force.bool
        disablewakeevents.bool
    } -maxleftover 0 -nulldefault]

    eval_with_privileges {
        SetSuspendState [expr {$opts(state) eq "hibernate"}] $opts(force) $opts(disablewakeevents)
    } SeShutdownPrivilege
}

# Indicate is a device is powered down
interp alias {} twapi::get_device_power_state {} twapi::GetDevicePowerState

# Get the power status of the system
proc twapi::get_power_status {} {
    lassign  [GetSystemPowerStatus] ac battery lifepercent reserved lifetime fulllifetime

    set acstatus unknown
    if {$ac == 0} {
        set acstatus off
    } elseif {$ac == 1} {
        # Note only value 1 is "on", not just any non-0 value
        set acstatus on
    }

    set batterycharging unknown
    if {$battery == -1} {
        set batterystate unknown
    } elseif {$battery & 128} {
        set batterystate notpresent;  # No battery
    } else {
        if {$battery & 8} {
            set batterycharging true
        } else {
            set batterycharging false
        }
        if {$battery & 4} {
            set batterystate critical
        } elseif {$battery & 2} {
            set batterystate low
        } else {
            set batterystate high
        }
    }

    set batterylifepercent unknown
    if {$lifepercent >= 0 && $lifepercent <= 100} {
        set batterylifepercent $lifepercent
    }

    set batterylifetime $lifetime
    if {$lifetime == -1} {
        set batterylifetime unknown
    }

    set batteryfulllifetime $fulllifetime
    if {$fulllifetime == -1} {
        set batteryfulllifetime unknown
    }

    return [kl_create2 {
        -acstatus
        -batterystate
        -batterycharging
        -batterylifepercent
        -batterylifetime
        -batteryfulllifetime
    } [list $acstatus $batterystate $batterycharging $batterylifepercent $batterylifetime $batteryfulllifetime]]
}


# Start monitoring of the clipboard
proc twapi::_power_handler {power_event lparam} {
    variable _power_monitors

    if {![info exists _power_monitors] ||
        [llength $_power_monitors] == 0} {
        return; # Not an error, could have deleted while already queued
    }

    if {![kl_vget {
        0 apmquerysuspend
        2 apmquerysuspendfailed
        4 apmsuspend
        6 apmresumecritical
        7 apmresumesuspend
        9 apmbatterylow
        10 apmpowerstatuschange
        11 apmoemevent
        18 apmresumeautomatic
    } $power_event power_event]} {
        return;                 # Do not support this event
    }

    foreach {id script} $_power_monitors {
        set code [catch {eval [linsert $script end $power_event $lparam]} msg]
        if {$code == 1} {
            # Error - put in background but we do not abort
            after 0 [list error $msg $::errorInfo $::errorCode]
        }
    }
    return
}

proc twapi::start_power_monitor {script} {
    variable _power_monitors

    set script [lrange $script 0 end]; # Verify syntactically a list

    set id "power#[TwapiId]"
    if {![info exists _power_monitors] ||
        [llength $_power_monitors] == 0} {
        # No power monitoring in progress. Start it
        Twapi_PowerNotifyStart
    }

    lappend _power_monitors $id $script
    return $id
}



# Stop monitoring of the power
proc twapi::stop_power_monitor {clipid} {
    variable _power_monitors

    if {![info exists _power_monitors]} {
        return;                 # Should we raise an error instead?
    }

    set new_monitors {}
    foreach {id script} $_power_monitors {
        if {$id ne $clipid} {
            lappend new_monitors $id $script
        }
    }

    set _power_monitors $new_monitors
    if {[llength $_power_monitors] == 0} {
        Twapi_PowerNotifyStop
    }
}
