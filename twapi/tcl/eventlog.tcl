#
# Copyright (c) 2004, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

package require registry

namespace eval twapi {
    # Keep track of event log handles - values are "r" or "w"
    variable eventlog_handles
    array set eventlog_handles {}

    # We maintain caches so we do not do lookups all the time
    # Toplevel keys are the source values. Below are subdictionaries:
    #   regkey -> Registry key for the given source
    #   messagefile  -  lang id + event -> file
    #   categoryfile - lang id + cat -> file
    # TBD - have a means of clearing this out
    variable _eventlog_message_paths;
    set _eventlog_message_paths {}

}

# Open an eventlog for reading or writing
proc twapi::eventlog_open {args} {
    variable eventlog_handles

    array set opts [parseargs args {
        system.arg
        source.arg
        file.arg
        write
    } -nulldefault -maxleftover 0]
    if {$opts(source) == ""} {
        # Source not specified
        if {$opts(file) == ""} {
            # No source or file specified, default to current event log 
            # using executable name as source
            set opts(source) [file rootname [file tail [info nameofexecutable]]]
        } else {
            if {$opts(write)} {
                error "Option -file may not be used with -write"
            }
        }
    } else {
        # Source explicitly specified
        if {$opts(file) != ""} {
            error "Option -file may not be used with -source"
        }
    }

    if {$opts(write)} {
        set handle [RegisterEventSource $opts(system) $opts(source)]
        set mode write
    } else {
        if {$opts(source) != ""} {
            set handle [OpenEventLog $opts(system) $opts(source)]
        } else {
            set handle [OpenBackupEventLog $opts(system) $opts(file)]
        }
        set mode read
    }

    set eventlog_handles($handle) $mode
    return $handle
}

# Close an event log opened for writing
proc twapi::eventlog_close {hevl} {
    variable eventlog_handles

    if {[_eventlog_valid_handle $hevl read]} {
        CloseEventLog $hevl
    } else {
        DeregisterEventSource $hevl
    }

    unset eventlog_handles($hevl)
}


# Log an event
proc twapi::eventlog_write {hevl id args} {
    _eventlog_valid_handle $hevl write raise

    array set opts [parseargs args {
        {type.arg information {success error warning information auditsuccess auditfailure}}
        {category.int 1}
        loguser
        params.arg
        data.arg
    } -nulldefault]


    switch -exact -- $opts(type) {
        success          {set opts(type) 0}
        error            {set opts(type) 1}
        warning          {set opts(type) 2}
        information      {set opts(type) 4}
        auditsuccess     {set opts(type) 8}
        auditfailure     {set opts(type) 16}
        default {error "Invalid value '$opts(type)' for option -type"}
    }
    
    if {$opts(loguser)} {
        set user [get_current_user -sid]
    } else {
        set user ""
    }

    ReportEvent $hevl $opts(type) $opts(category) $id \
        $user $opts(params) $opts(data)
}


# Log a message 
proc twapi::eventlog_log {message args} {
    array set opts [parseargs args {
        system.arg
        source.arg
        {type.arg information}
        {category.int 1}
    } -nulldefault]

    set hevl [eventlog_open -write -source $opts(source) -system $opts(system)]

    trap {
        eventlog_write $hevl 1 -params [list $message] -type $opts(type) -category $opts(category)
    } finally {
        eventlog_close $hevl
    }
    return
}


# Read the event log
proc twapi::eventlog_read {hevl args} {
    _eventlog_valid_handle $hevl read raise

    array set opts [parseargs args {
        seek.int
        {direction.arg forward}
    }]

    if {[info exists opts(seek)]} {
        set flags 2;                    # Seek
        set offset $opts(seek)
    } else {
        set flags 1;                    # Sequential read
        set offset 0
    }

    switch -glob -- $opts(direction) {
        ""    -
        forw* {
            setbits flags 4
        }
        back* {
            setbits flags 8
        }
        default {
            error "Invalid value '$opts(direction)' for -direction option"
        }
    }

    set results [list ]

    trap {
        set recs [ReadEventLog $hevl $flags $offset]
    } onerror {TWAPI_WIN32 38} {
        # EOF - no more
        set recs [list ]
    }
    foreach rec $recs {
        set event [twine {
            -source -system -reserved -recordnum -timegenerated
            -timewritten -eventid -type -category -reservedflags
            -recnum -params -sid -data 
        } $rec]
        dict set event -type [string map {0 success 1 error 2 warning 4 information 8 auditsuccess 16 auditfailure} [dict get $event -type]]
        lappend results $event
    }

    return $results
}


# Get the oldest event log record index. $hevl must be read handle
proc twapi::eventlog_oldest {hevl} {
    _eventlog_valid_handle $hevl read raise
    return [GetOldestEventLogRecord $hevl]
}

# Get the event log record count. $hevl must be read handle
proc twapi::eventlog_count {hevl} {
    _eventlog_valid_handle $hevl read raise
    return [GetNumberOfEventLogRecords $hevl]
}

# Check if the event log is full. $hevl may be either read or write handle
# (only win2k plus)
proc twapi::eventlog_is_full {hevl} {
    # Does not matter if $hevl is read or write, but verify it is a handle
    _eventlog_valid_handle $hevl read
    return [Twapi_IsEventLogFull $hevl]
}

# Backup the event log
proc twapi::eventlog_backup {hevl file} {
    _eventlog_valid_handle $hevl read raise
    BackupEventLog $hevl $file
}

# Clear the event log
proc twapi::eventlog_clear {hevl args} {
    _eventlog_valid_handle $hevl read raise
    array set opts [parseargs args {backup.arg} -nulldefault]
    ClearEventLog $hevl $opts(backup)
}


# Formats the given event log record message
# 
proc twapi::eventlog_format_message {rec args} {
    variable _eventlog_message_paths

    array set opts [parseargs args {
        width.int
        langid.int
    } -nulldefault]

    # TBD - See if we can cache the actual message string
    # At least cache the DLL handle ?
    # Maybe associate cache with the open eventlog handle? And close accordingly?

    set source  [dict get $rec -source]
    set eventid [dict get $rec -eventid]

    # Find the registry key if we do not have it already
    if {[dict exists $_eventlog_message_paths $source regkey]} {
        set regkey [dict get $_eventlog_message_paths $source regkey]
        dict incr _eventlog_message_paths __regkey_hits; # TBD
    } else {
        set regkey [_find_eventlog_regkey $source]
        dict set _eventlog_message_paths $source regkey $regkey
        dict incr _eventlog_message_paths __regkey_misses; # TBD
    }

    set found 0
    if {[dict exists $_eventlog_message_paths $source messagefile $opts(langid) $eventid]} {
        dict incr _eventlog_message_paths __messagefile_hits; # TBD
        if {! [catch {
            format_message -module [dict get $_eventlog_message_paths $source messagefile $opts(langid) $eventid] -messageid $eventid -params [dict get $rec -params] -width $opts(width) -langid $opts(langid)
        } msg]} {
            set found 1
        }
    } else {
        dict incr _eventlog_message_paths __messagefile_misses; # TBD
        # Get the message file, if there is one
        if {! [catch {registry get $regkey "EventMessageFile"} path]} {
            # Try each file listed in turn
            foreach dll [split $path \;] {
                set dll [expand_environment_strings $dll]
                if {! [catch {
                    format_message -module $dll -messageid $eventid -params [dict get $rec -params] -width $opts(width) -langid $opts(langid)
                } msg]} {
                    dict set _eventlog_message_paths $source messagefile $opts(langid) $eventid $dll
                    set found 1
                    break
                }
            }
        }
    }

    if {! $found} {
        set fmt "The message file or event definition for event id [dict get $rec -eventid] from source [dict get $rec -source] was not found. The following information was part of the event: "
        set flds [list ]
        for {set i 1} {$i <= [llength [dict get $rec -params]]} {incr i} {
            lappend flds %$i
        }
        append fmt [join $flds ", "]
        return [format_message -fmtstring $fmt  \
                    -params [dict get $rec -params] -width $opts(width)]
    }

    # We'd found a message from the message file and replaced the string
    # parameters. Now fill in the parameter file values if any. Note these are
    # separate from the string parameters passed in through rec(-params)

    # First check if the formatted string itself still has placeholders
    # Place holder for the parameters file are supposed to start
    # with two % chars. Unfortunately, not all apps, even Microsoft's own
    # DCOM obey this. So check for both % and %%
    set placeholder_indices [regexp -indices -all -inline {%?%\d+} $msg]
    if {[llength $placeholder_indices] == 0} {
        # No placeholders.
        return $msg
    }

    # Need to get strings from the parameter file
    # TBD - add parameterfile to cache
    dict incr _eventlog_message_paths __parameterfile_misses; # TBD
    if {! [catch {registry get $regkey "ParameterMessageFile"} path]} {
        # Loop through every placeholder, look for the entry in the
        # parameters file and replace it if found

        set msgfiles {}
        foreach msgfile [split $path \;] {
            lappend msgfiles [expand_environment_strings $msgfile]
        }
        set msg2 ""
        set prev_end 0
        foreach placeholder $placeholder_indices {
            lassign $placeholder start end
            # Append the stuff between previous placeholder and this one
            append msg2 [string range $msg $prev_end [expr {$start-1}]]
            set repl [string range $msg $start $end]; # Default if not found
            set msgid [string trimleft $repl %];     # Skip "%"
            # Try each file listed in turn
            foreach msgfile $msgfiles {
                if {! [catch {
                    set repl [string trimright [format_message -module $msgfile -messageid $msgid -params [dict get $rec -params] -langid $opts(langid)] \r\n]
                } ]} {
                    # Found the replacement
                    break
                }
            }
            append msg2 $repl
            set prev_end [incr end]
        }
        append msg2 [string range $msg $prev_end end]
        set msg $msg2
    }

    return $msg
}

# Format the category
proc twapi::eventlog_format_category {rec args} {

    array set opts [parseargs args {
        width.int
        langid.int
    } -nulldefault]

    set category [dict get $rec -category]
    if {$category == 0} {
        return ""
    }

    variable _eventlog_message_paths

    # TBD - See if we can cache the actual message string
    # At least cache the DLL handle ?
    # Maybe associate cache with the open eventlog handle? And close accordingly?
    set source  [dict get $rec -source]

    # Get the message file, if there is one
    if {[dict exists $_eventlog_message_paths $source categoryfile $opts(langid) $category]} {
        dict incr _eventlog_message_paths __categoryfile_hits; # TBD
        if {! [catch {
            format_message -module [dict get $_eventlog_message_paths $source categoryfile $opts(langid) $category] -messageid $category -params [dict get $rec -params] -width $opts(width) -langid $opts(langid)
        } msg]} {
            return $msg
        }
    } else {
        dict incr _eventlog_message_paths __categoryfile_misses; # TBD

        # Find the registry key if we do not have it already
        if {[dict exists $_eventlog_message_paths $source regkey]} {
            dict incr _eventlog_message_paths __regkey_hits; # TBD
            set regkey [dict get $_eventlog_message_paths $source regkey]
        } else {
            dict incr _eventlog_message_paths __regkey_misses; # TBD
            set regkey [_find_eventlog_regkey $source]
            dict set _eventlog_message_paths $source regkey $regkey
        }

        if {! [catch {registry get $regkey "CategoryMessageFile"} path]} {
            # Try each file listed in turn
            foreach dll [split $path \;] {
                set dll [expand_environment_strings $dll]
                if {! [catch {
                    format_message -module $dll -messageid $category -params [dict get $rec -params] -width $opts(width) -langid $opts(langid)
                } msg]} {
                    dict set _eventlog_message_paths $source categoryfile $opts(langid) $category $dll
                    return $msg
                }
            }
        }
    }

    return "Category $category"
}

proc twapi::eventlog_monitor_start {hevl script} {
    variable _eventlog_notification_scripts

    set hevent [create_event]
    if {[catch {NotifyChangeEventLog $hevl $hevent} msg]} {
        CloseHandle $hevent
        error $msg $::errorInfo $::errorCode
    }

    wait_on_handle $hevent -async twapi::_eventlog_notification_handler
    set _eventlog_notification_scripts($hevent) $script

    # We do not want the application mistakenly closing the event
    # while being waited on by the thread pool. That would be a big NO-NO
    # so change the handle type so it cannot be passed to close_handle.
    return [list evl $hevent]
}

# Stop any notifications. Note these will stop even if the event log
# handle is closed but leave the event dangling.
proc twapi::eventlog_monitor_stop {hevent} {
    variable _eventlog_notification_scripts
    set hevent [lindex $hevent 1]
    if {[info exists _eventlog_notification_scripts($hevent)]} {
        unset _eventlog_notification_scripts($hevent)
        cancel_wait_on_handle $hevent
        CloseHandle $hevent
    }
}

proc twapi::_eventlog_notification_handler {hevent event} {
    variable _eventlog_notification_scripts
    if {[info exists _eventlog_notification_scripts($hevent)] &&
        $event eq "signalled"} {
        uplevel #0 $_eventlog_notification_scripts($hevent) [list [list evl $hevent]]
    }
}

# Utility procs

# Validate a handle for a mode. Always raises error if handle is invalid
# If handle valid but not for that mode, will raise error iff $raise_error
# is non-empty. Returns 1 if valid, 0 otherwise
proc twapi::_eventlog_valid_handle {hevl mode {raise_error ""}} {
    variable eventlog_handles
    if {![info exists eventlog_handles($hevl)]} {
        error "Invalid event log handle '$hevl'"
    }

    if {[string compare $eventlog_handles($hevl) $mode]} {
        if {$raise_error != ""} {
            error "Eventlog handle '$hevl' not valid for $mode"
        }
        return 0
    } else {
        return 1
    }
}

# Find the registry key corresponding the given event log source
proc twapi::_find_eventlog_regkey {source} {
    set topkey {HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog}

    # Set a default list of children to work around an issue in
    # the Tcl [registry keys] command where a ERROR_MORE_DATA is returned
    # instead of a retry with a larger buffer.
    set keys {Application Security System}
    catch {set keys [registry keys $topkey]}
    # Get all keys under this key and look for a source under that
    foreach key $keys {
        # See above Tcl issue
        set srckeys {}
        catch {set srckeys [registry keys "${topkey}\\$key"]}
        foreach srckey $srckeys {
            if {[string equal -nocase $srckey $source]} {
                return "${topkey}\\${key}\\$srckey"
            }
        }
    }

    # Default to Application - TBD
    return "${topkey}\\Application"
}

proc twapi::_eventlog_dump {source chan} {
    set hevl [eventlog_open -source $source]
    while {[llength [set events [eventlog_read $hevl]]]} {
        # print out each record
        foreach eventrec $events {
            array set event $eventrec
            set timestamp [clock format $event(-timewritten) -format "%x %X"]
            set source   $event(-source)
            set category [twapi::eventlog_format_category $eventrec -width -1]
            set message  [twapi::eventlog_format_message $eventrec -width -1]
            puts -nonewline $chan "Time: $timestamp\r\nSource: $source\r\nCategory: $category\r\n$message\r\n\r\n"
        }
    }
    eventlog_close $hevl
}