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
    foreach {ac battery lifepercent reserved lifetime fulllifetime} [GetSystemPowerStatus] break

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
