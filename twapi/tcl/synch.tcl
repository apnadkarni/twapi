#
# Copyright (c) 2004, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license
#
# TBD - document

namespace eval twapi {
}

#
# Create and return a handle to a mutex
proc twapi::create_mutex {args} {
    array set opts [parseargs args {
        {name.arg ""}
        {secd.arg ""}
        {inherit.bool 0}
        lock
    }]

    return [CreateMutex [_make_secattr $opts(secd) $opts(inherit)] $opts(lock) $opts(name)]
}

# Get handle to an existing mutex
proc twapi::get_mutex_handle {name args} {
    array set opts [parseargs args {
        {inherit.bool 0}
        {access.arg {mutex_all_access}}
    }]
    
    return [OpenMutex [_access_rights_to_mask $opts(access)] $opts(inherit) $name]
}

# Lock the mutex
proc twapi::lock_mutex {h args} {
    array set opts [parseargs args {
        {wait.int 1000}
    }]

    return [wait_on_handle $h -wait $opts(wait)]
}


# Unlock the mutex
proc twapi::unlock_mutex {h} {
    ReleaseMutex $h
}

#
# Wait on a handle
proc twapi::wait_on_handle {h args} {
    variable _wait_handles

    array set opts [parseargs args {
        {wait.int -1}
        async.arg
        {executeonce.bool false}
    }]

    if {![info exists opts(async)]} {
        if {[info exists _wait_handles($h)]} {
            error "Attempt to synchronously wait on handle that is registered for an asynchronous wait."
        }

        set ret [WaitForSingleObject $h $opts(wait)]
        if {$ret == 0x80} {
            return abandoned
        } elseif {$ret == 0} {
            return signalled
        } elseif {$ret == 0x102} {
            return timeout
        } else {
            error "Unexpected value $ret returned from WaitForSingleObject"
        }
    }

    # async option specified

    # If already registered, only the callback script is replaced.
    # Wait time and -singlewait settings are not changed
    if {[info exists _wait_handles($h)]} {
        set _wait_handles($h) $opts(async)
        return
    }

    # New registration.
    set flags 0
    if {$opts(executeonce)} {
        set flags 0x00000008;   # WT_EXECUTEONCEONLY
    }
    Twapi_RegisterWaitOnHandle $h $opts(wait) $flags

    # Set now that successfully registered
    set _wait_handles($h) $opts(async)

    return
}

#
# Cancel an async wait on a handle
proc twapi::cancel_wait_on_handle {h} {
    variable _wait_handles
    if {[info exists _wait_handles($h)]} {
        Twapi_UnregisterWaitOnHandle $h
    }
}

#
# Called from C when a handle is signalled or times out
proc twapi::_wait_handler {h event} {
    variable _wait_handles

    if {[info exists _wait_handles($h)]} {
        eval $_wait_handles($h) [list $h $event]
    }

    return
}
