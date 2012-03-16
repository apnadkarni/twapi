#
# Copyright (c) 2012, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Commands in twapi_base module

# Return major minor servicepack as a quad list
proc twapi::get_os_version {} {
    array set verinfo [GetVersionEx]
    return [list $verinfo(dwMajorVersion) $verinfo(dwMinorVersion) \
                $verinfo(wServicePackMajor) $verinfo(wServicePackMinor)]
}

# Returns true if the OS version is at least $major.$minor.$sp
proc twapi::min_os_version {major {minor 0} {spmajor 0} {spminor 0}} {
    lassign  [twapi::get_os_version]  osmajor osminor osspmajor osspminor

    if {$osmajor > $major} {return 1}
    if {$osmajor < $major} {return 0}
    if {$osminor > $minor} {return 1}
    if {$osminor < $minor} {return 0}
    if {$osspmajor > $spmajor} {return 1}
    if {$osspmajor < $spmajor} {return 0}
    if {$osspminor > $spminor} {return 1}
    if {$osspminor < $spminor} {return 0}

    # Same version, ok
    return 1
}

# Convert a LARGE_INTEGER time value (100ns since 1601) to a formatted date
# time
interp alias {} twapi::large_system_time_to_secs {} twapi::large_system_time_to_secs_since_1970
proc twapi::large_system_time_to_secs_since_1970 {ns100 {fraction false}} {
    # No. 100ns units between 1601 to 1970 = 116444736000000000
    set ns100_since_1970 [expr {wide($ns100)-wide(116444736000000000)}]

    if {0} {
        set secs_since_1970 [expr {wide($ns100_since_1970)/wide(10000000)}]
        if {$fraction} {
            append secs_since_1970 .[expr {wide($ns100_since_1970)%wide(10000000)}]
        }
    } else {
        # Equivalent to above but faster
        if {[string length $ns100_since_1970] > 7} {
            set secs_since_1970 [string range $ns100_since_1970 0 end-7]
            if {$fraction} {
                set frac [string range $ns100_since_1970 end-6 end]
                append secs_since_1970 .$frac
            }
        } else {
            set secs_since_1970 0
            if {$fraction} {
                set frac [string range "0000000${ns100_since_1970}" end-6 end]
                append secs_since_1970 .$frac
            }
        }
    }
    return $secs_since_1970
}

proc twapi::secs_since_1970_to_large_system_time {secs} {
    # No. 100ns units between 1601 to 1970 = 116444736000000000
    return [expr {($secs * 10000000) + wide(116444736000000000)}]
}

interp alias {} ::twapi::get_system_time {} ::twapi::GetSystemTimeAsFileTime
interp alias {} ::twapi::large_system_time_to_timelist {} ::twapi::FileTimeToSystemTime
interp alias {} ::twapi::timelist_to_large_system_time {} ::twapi::SystemTimeToFileTime

# Map a Windows error code to a string
proc twapi::map_windows_error {code} {
    # Trim trailing CR/LF
    return [string trimright [twapi::Twapi_MapWindowsErrorToString $code] "\r\n"]
}

# Load given library
proc twapi::load_library {path args} {
    array set opts [parseargs args {
        dontresolverefs
        datafile
        alteredpath
    }]

    set flags 0
    if {$opts(dontresolverefs)} {
        setbits flags 1;                # DONT_RESOLVE_DLL_REFERENCES
    }
    if {$opts(datafile)} {
        setbits flags 2;                # LOAD_LIBRARY_AS_DATAFILE
    }
    if {$opts(alteredpath)} {
        setbits flags 8;                # LOAD_WITH_ALTERED_SEARCH_PATH
    }

    # LoadLibrary always wants backslashes
    set path [file nativename $path]
    return [LoadLibraryEx $path $flags]
}

# Free library opened with load_library
proc twapi::free_library {libh} {
    FreeLibrary $libh
}


# Format message string
proc twapi::format_message {args} {
    array set opts [parseargs args {
        params.arg
        fmtstring.arg
        width.int
    } -ignoreunknown]

    # TBD - document - if no params specified, different from params = {}

    # If a format string is specified, other options do not matter
    # except for -width. In that case, we do not call FormatMessage
    # at all
    if {[info exists opts(fmtstring)]} {
        # If -width specifed, call FormatMessage
        if {[info exists opts(width)] && $opts(width)} {
            set msg [_unsafe_format_message -ignoreinserts -fmtstring $opts(fmtstring) -width $opts(width) {*}$args]
        } else {
            set msg $opts(fmtstring)
        }
    } else {
        # Not -fmtstring, retrieve from message file
        if {[info exists opts(width)]} {
            set msg [_unsafe_format_message -ignoreinserts -width $opts(width) {*}$args]
        } else {
            set msg [_unsafe_format_message -ignoreinserts {*}$args]
        }
    }

    # If not param list, do not replace placeholder. This is NOT
    # the same as empty param list
    if {![info exists opts(params)]} {
        return $msg
    }

    set placeholder_indices [regexp -indices -all -inline {%(?:.|(?:[1-9][0-9]?(?:![^!]+!)?))} $msg]

    if {[llength $placeholder_indices] == 0} {
        # No placeholders.
        return $msg
    }

    # Use of * in format specifiers will change where the actual parameters
    # are positioned
    set num_asterisks 0
    set msg2 ""
    set prev_end 0
    foreach placeholder $placeholder_indices {
        lassign $placeholder start end
        # Append the stuff between previous placeholder and this one
        append msg2 [string range $msg $prev_end [expr {$start-1}]]
        set spec [string range $msg $start+1 $end]
        switch -exact -- [string index $spec 0] {
            % { append msg2 % }
            r { append msg2 \r }
            n { append msg2 \n }
            t { append msg2 \t }
            0 { 
                # No-op - %0 means to not add trailing newline
            }
            default {
                if {! [string is integer -strict [string index $spec 0]]} {
                    # Not a insert parameter. Just append the character
                    append msg2 $spec
                } else {
                    # Insert parameter
                    set fmt ""
                    scan $spec %d%s param_index fmt
                    # Note params are numbered starting with 1
                    incr param_index -1
                    # Format spec, if present, is enclosed in !. Get rid of them
                    set fmt [string trim $fmt "!"]
                    if {$fmt eq ""} {
                        # No fmt spec
                    } else {
                        # Since everything is a string in Tcl, we happily
                        # do not have to worry about type. However, the
                        # format spec could have * specifiers which will
                        # change the parameter indexing for subsequent
                        # arguments
                        incr num_asterisks [expr {[llength [split $fmt *]]-1}]
                        incr param_index $num_asterisks
                    }
                    # TBD - we ignore the actual format type
                    append msg2 [lindex $opts(params) $param_index]
                }                        
            }
        }                    
        set prev_end [incr end]
    }
    append msg2 [string range $msg $prev_end end]
    return $msg2
}

# Revert to process token. In base package because used across many modules
proc twapi::revert_to_self {{opt ""}} {
    RevertToSelf
}

