#
# Copyright (c) 2010, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {
    # Array maps handles we are waiting on to the ids of the registered waits
    variable _wait_handle_ids
    # Array maps id of registered wait to the corresponding callback scripts
    variable _wait_handle_scripts
    
}

proc twapi::cast_handle {h type} {
    # Don't convert untyped values like 0 and NULL
    if {[llength $h] > 1} {
        return [list [lindex $h 0] $type]
    } else {
        return $h
    }
}

proc twapi::close_handle {h} {

    # Cancel waits on the handle, if any
    cancel_wait_on_handle $h
    
    # Then close it
    CloseHandle $h
}

# Close multiple handles. In case of errors, collects them but keeps
# closing remaining handles and only raises the error at the end.
proc twapi::close_handles {args} {
    # The original definition for this was broken in that it would
    # gracefully accept non list parameters as a list of one. In 3.0
    # the handle format has changed so this does not happen
    # naturally. We have to try and decipher whether it is a list
    # of handles or a single handle.

    foreach arg $args {
        if {[Twapi_IsPtr $arg]} {
            # Looks like a single handle
            if {[catch {close_handle $arg} msg]} {
                set erinfo $::errorInfo
                set ercode $::errorCode
                set ermsg $msg
            }
        } else {
            # Assume a list of handles
            foreach h $arg {
                if {[catch {close_handle $h} msg]} {
                    set erinfo $::errorInfo
                    set ercode $::errorCode
                    set ermsg $msg
                }
            }
        }
    }

    if {[info exists erinfo]} {
        error $msg $erinfo $ercode
    }
}

#
# Wait on a handle
proc twapi::wait_on_handle {hwait args} {
    variable _wait_handle_ids
    variable _wait_handle_scripts

    # When we are invoked from callback, handle is always typed as HANDLE
    # so convert it so lookups succeed
    set h [cast_handle $hwait HANDLE]

    # 0x00000008 ->   # WT_EXECUTEONCEONLY
    array set opts [parseargs args {
        {wait.int -1}
        async.arg
        {executeonce.bool false 0x00000008}
    }]

    if {![info exists opts(async)]} {
        if {[info exists _wait_handle_ids($h)]} {
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

    # Do not wait on manual reset events as cpu will spin continuously
    # queueing events
    if {[Twapi_IsPtr $hwait HANDLE_MANUALRESETEVENT] &&
        ! $opts(executeonce)
    } {
        error "A handle to a manual reset event cannot be waited on asynchronously unless -executeonce is specified."
    }

    # If handle already registered, cancel previous registration.
    if {[info exists _wait_handle_ids($h)]} {
        cancel_wait_on_handle $h
    }


    set id [Twapi_RegisterWaitOnHandle $h $opts(wait) $opts(executeonce)]

    # Set now that successfully registered
    set _wait_handle_scripts($id) $opts(async)
    set _wait_handle_ids($h) $id

    return
}

#
# Cancel an async wait on a handle
proc twapi::cancel_wait_on_handle {h} {
    variable _wait_handle_ids
    variable _wait_handle_scripts

    if {[info exists _wait_handle_ids($h)]} {
        Twapi_UnregisterWaitOnHandle $_wait_handle_ids($h)
        unset _wait_handle_scripts($_wait_handle_ids($h))
        unset _wait_handle_ids($h)
    }
}

#
# Called from C when a handle is signalled or times out
proc twapi::_wait_handler {id h event} {
    variable _wait_handle_ids
    variable _wait_handle_scripts

    # We ignore the following stale event cases -
    #  - _wait_handle_ids($h) does not exist : the wait was canceled while
    #    and event was queued
    #  - _wait_handle_ids($h) exists but is different from $id - same
    #    as prior case, except that a new wait has since been initiated
    #    on the same handle value (which might have be for a different
    #    resource

    if {[info exists _wait_handle_ids($h)] &&
        $_wait_handle_ids($h) == $id} {
        eval $_wait_handle_scripts($id) [list $h $event]
    }

    return
}
