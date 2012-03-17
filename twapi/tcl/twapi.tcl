#
# Copyright (c) 2003-2012, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# General definitions and procs used by all TWAPI modules

package require Tcl 8.5
package require registry

namespace eval twapi {
    variable nullptr "__null__"
    variable scriptdir [file dirname [info script]]

    # Accessing global environ in ::env is expensive so cache
    # it. Not updated even if real environ changes
    proc getenv {varname} {
        variable envcache
        set varname [string toupper $varname]
        if {[info exists envcache($varname)]} {
            return $envcache($varname)
        }
        return [set envcache($varname) $::env($varname)]
    }

}

# Make twapi versions the same as the base module versions
set twapi::version(twapi) $::twapi::version(twapi_base)

# log for tracing / debug messages.
proc twapi::debuglog {args} {
    variable log_messages
    if {[llength $args] == 0} {
        if {[info exists log_messages]} {
            return $log_messages
        }
        return [list ]
    }
    foreach msg $args {
        Twapi_AppendLog $msg
    }
}

proc twapi::debuglog_clear {} {
    variable log_messages
    set log_messages {}
}

proc twapi::get_build_config {{key ""}} {
    variable build_ids
    array set config [GetTwapiBuildInfo]

    # This is actually a runtime config and might not have been initialized
    if {[info exists ::twapi::use_tcloo_for_com]} {
        if {$::twapi::use_tcloo_for_com} {
            set config(comobj_ootype) tcloo
        } else {
            set config(comobj_ootype) metoo
        }
    } else {
        set config(comobj_ootype) uninitialized
    }

    if {$key eq ""} {
        return [array get config]
    } else {
        return $config($key)
    }
}


# Construct a security attributes structure out of a security descriptor
# and inheritance. The command is here because we do not want to
# have to load the twapi_security package for the common case of
# null security attributes.
proc twapi::_make_secattr {secd inherit} {
    if {$inherit} {
        set sec_attr [list $secd 1]
    } else {
        if {[llength $secd] == 0} {
            # If a security descriptor not specified, keep
            # all security attributes as an empty list (ie. NULL)
            set sec_attr [list ]
        } else {
            set sec_attr [list $secd 0]
        }
    }
    return $sec_attr
}



# Returns a list of raw Windows API functions supported
proc twapi::list_raw_api {} {
    set rawapi [list ]
    foreach fn [info commands ::twapi::*] {
         if {[regexp {^::twapi::([A-Z][^_]*)$} $fn ignore fn]} {
             lappend rawapi $fn
         }
    }
    return $rawapi
}


# Wait for $wait_ms milliseconds or until $script returns $guard. $gap_ms is
# time between retries to call $script
# TBD - write a version that will allow other events to be processed
proc twapi::wait {script guard wait_ms {gap_ms 10}} {
    if {$gap_ms == 0} {
        set gap_ms 10
    }
    set end_ms [expr {[clock clicks -milliseconds] + $wait_ms}]
    while {[clock clicks -milliseconds] < $end_ms} {
        set script_result [uplevel $script]
        if {[string equal $script_result $guard]} {
            return 1
        }
        after $gap_ms
    }
    # Reached limit, one last try
    return [string equal [uplevel $script] $guard]
}

# Get twapi version
proc twapi::get_version {args} {
    variable version
    array set opts [parseargs args {patchlevel}]
    if {$opts(patchlevel)} {
        return $version(twapi)
    } else {
        # Only return major, minor
        set ver $version(twapi)
        regexp {^([[:digit:]]+\.[[:digit:]]+)[.ab]} $version(twapi) - ver
        return $ver
    }
}

# Set all elements of the array to specified value
proc twapi::_array_set_all {v_arr val} {
    upvar $v_arr arr
    foreach e [array names arr] {
        set arr($e) $val
    }
}

# Check if any of the specified array elements are non-0
proc twapi::_array_non_zero_entry {v_arr indices} {
    upvar $v_arr arr
    foreach i $indices {
        if {$arr($i)} {
            return 1
        }
    }
    return 0
}

# Check if any of the specified array elements are non-0
# and return them as a list of options (preceded with -)
proc twapi::_array_non_zero_switches {v_arr indices all} {
    upvar $v_arr arr
    set result [list ]
    foreach i $indices {
        if {$all || ([info exists arr($i)] && $arr($i))} {
            lappend result -$i
        }
    }
    return $result
}


# Bitmask operations on 32bit values
# The int() casts are to deal with hex-decimal sign extension issues
proc twapi::setbits {v_bits mask} {
    upvar $v_bits bits
    set bits [expr {int($bits) | int($mask)}]
    return $bits
}
proc twapi::resetbits {v_bits mask} {
    upvar $v_bits bits
    set bits [expr {int($bits) & int(~ $mask)}]
    return $bits
}

# Return a bitmask corresponding to a list of symbolic and integer values
# If symvals is a single item, it is an array else a list of sym bitmask pairs
proc twapi::_parse_symbolic_bitmask {syms symvals} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
    } else {
        array set lookup $symvals
    }
    set bits 0
    foreach sym $syms {
        if {[info exists lookup($sym)]} {
            set bits [expr {$bits | $lookup($sym)}]
        } else {
            set bits [expr {$bits | $sym}]
        }
    }
    return $bits
}

# Return a list of symbols corresponding to a bitmask
proc twapi::_make_symbolic_bitmask {bits symvals {append_unknown 1}} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
        set map [array get lookup]
    } else {
        set map $symvals
    }
    set symbits 0
    set symmask [list ]
    foreach {sym val} $map {
        if {$bits & $val} {
            set symbits [expr {$symbits | $val}]
            lappend symmask $sym
        }
    }

    # Get rid of bits that mapped to symbols
    set bits [expr {$bits & ~$symbits}]
    # If any left over, add them
    if {$bits && $append_unknown} {
        lappend symmask $bits
    }
    return $symmask
}

# Return a bitmask corresponding to a list of symbolic and integer values
# If symvals is a single item, it is an array else a list of sym bitmask pairs
# Ditto for switches - an array or flat list of switch boolean pairs
proc twapi::_switches_to_bitmask {switches symvals {bits 0}} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
    } else {
        array set lookup $symvals
    }
    if {[llength $switches] == 1} {
        upvar $switches swtable
    } else {
        array set swtable $switches
    }

    foreach {switch bool} [array get swtable] {
        if {$bool} {
            set bits [expr {$bits | $lookup($switch)}]
        } else {
            set bits [expr {$bits & ~ $lookup($switch)}]
        }
    }
    return $bits
}

# Return a list of switche bool pairs corresponding to a bitmask
proc twapi::_bitmask_to_switches {bits symvals} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
        set map [array get lookup]
    } else {
        set map $symvals
    }
    set symbits 0
    set symmask [list ]
    foreach {sym val} $map {
        if {$bits & $val} {
            set symbits [expr {$symbits | $val}]
            lappend symmask $sym 1
        } else {
            lappend symmask $sym 0
        }
    }

    return $symmask
}

# Make and return a keyed list
proc twapi::kl_create {args} {
    if {[llength $args] & 1} {
        error "No value specified for keyed list field [lindex $args end]. A keyed list must have an even number of elements."
    }
    return $args
}

# Make a keyed list given fields and values
interp alias {} twapi::kl_create2 {} twapi::twine

# Return a field from a keyed list or a default if not present
# This routine is now obsolete since the C version of kl_get takes
# an optional default parameter
# kl_get_default KEYEDLIST KEY DEFAULT
interp alias {} ::twapi::kl_get_default {} ::twapi::kl_get

# Set a key value
proc twapi::kl_set {kl field newval} {
   set i 0
   foreach {fld val} $kl {
        if {[string equal $fld $field]} {
            incr i
            return [lreplace $kl $i $i $newval]
        }
        incr i 2
    }
    lappend kl $field $newval
    return $kl
}

# Check if a field exists in the keyed list
proc twapi::kl_vget {kl field varname} {
    upvar $varname var
    return [expr {! [catch {set var [kl_get $kl $field]}]}]
}

# Remote/unset a key value
proc twapi::kl_unset {kl field} {
    array set arr $kl
    unset -nocomplain arr($field)
    return [array get arr]
}

# Compare two keyed lists
proc twapi::kl_equal {kl_a kl_b} {
    array set a $kl_a
    foreach {kb valb} $kl_b {
        if {[info exists a($kb)] && ($a($kb) == $valb)} {
            unset a($kb)
        } else {
            return 0
        }
    }
    if {[array size a]} {
        return 0
    } else {
        return 1
    }
}

# Return the field names in a keyed list in the same order that they
# occured
proc twapi::kl_fields {kl} {
    set fields [list ]
    foreach {fld val} $kl {
        lappend fields $fld
    }
    return $fields
}

# Returns a flat list of the $field fields from a list
# of keyed lists
proc twapi::kl_flatten {list_of_kl args} {
    set result {}
    foreach kl $list_of_kl {
        foreach field $args {
            lappend result [kl_get $kl $field]
        }
    }
    return $result
}


# Print the specified fields of a keyed list
proc twapi::kl_print {kl args} {
    # If only one arg, just print value without label
    if {[llength $args] == 1} {
        puts [kl_get $kl [lindex $args 0]]
        return
    }
    if {[llength $args] == 0} {
        array set arr $kl
        parray arr
    }
    foreach field $args {
        puts "$field: [kl_get $kl $field]"
    }
    return
}


# Return an array as a list of -index value pairs
proc twapi::_get_array_as_options {v_arr} {
    upvar $v_arr arr
    set result [list ]
    foreach {index value} [array get arr] {
        lappend result -$index $value
    }
    return $result
}

# Parse a list of two integers or a x,y pair and return a list of two integers
# Generate exception on format error using msg
proc twapi::_parse_integer_pair {pair {msg "Invalid integer pair"}} {
    if {[llength $pair] == 2} {
        lassign $pair first second
        if {[string is integer -strict $first] &&
            [string is integer -strict $second]} {
            return [list $first $second]
        }
    } elseif {[regexp {^([[:digit:]]+),([[:digit:]]+)$} $pair dummy first second]} {
        return [list $first $second]
    }

    error "$msg: '$pair'. Should be a list of two integers or in the form 'x,y'"
}


# Convert file names by substituting \SystemRoot and \??\ sequences
proc twapi::_normalize_path {path} {
    # Get rid of \??\ prefixes
    regsub {^[\\/]\?\?[\\/](.*)} $path {\1} path

    # Replace leading \SystemRoot with real system root
    if {[string match -nocase {[\\/]Systemroot*} $path] &&
        ([string index $path 11] in [list "" / \\])} {
        return [file join [twapi::GetSystemWindowsDirectory] [string range $path 12 end]]
    } else {
        return [file normalize $path]
    }
}


# Convert seconds to a list {Year Month Day Hour Min Sec Ms}
# (Ms will always be zero). Always return local time
proc twapi::_seconds_to_timelist {secs} {
    # For each field, we need to trim the leading zeroes
    set result [list ]
    foreach x [clock format $secs -format "%Y %m %e %k %M %S 0" -gmt false] {
        lappend result [scan $x %d]
    }
    return $result
}

# Convert local time list {Year Month Day Hour Min Sec Ms} to seconds
# (Ms field is ignored)
proc twapi::_timelist_to_seconds {timelist} {
    return [clock scan [_timelist_to_timestring $timelist] -gmt false]
}

# Convert local time list {Year Month Day Hour Min Sec Ms} to a time string
# (Ms field is ignored)
proc twapi::_timelist_to_timestring {timelist} {
    if {[llength $timelist] < 6} {
        error "Invalid time list format"
    }

    return "[lindex $timelist 0]-[lindex $timelist 1]-[lindex $timelist 2] [lindex $timelist 3]:[lindex $timelist 4]:[lindex $timelist 5]"
}

# Convert a time string to a time list
proc twapi::_timestring_to_timelist {timestring} {
    return [_seconds_to_timelist [clock scan $timestring -gmt false]]
}

# Parse raw memory like binary scan command
proc twapi::mem_binary_scan {mem off mem_sz args} {
    uplevel [list binary scan [Twapi_ReadMemoryBinary $mem $off $mem_sz]] $args
}


# Validate guid syntax
proc twapi::_validate_guid {guid} {
    if {![Twapi_IsValidGUID $guid]} {
        error "Invalid GUID syntax: '$guid'"
    }
}

# Validate uuid syntax
proc twapi::_validate_uuid {uuid} {
    if {![regexp {^[[:xdigit:]]{8}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{12}$} $uuid]} {
        error "Invalid UUID syntax: '$uuid'"
    }
}

# Extract a UCS-16 string from a binary. Cannot directly use
# encoding convertfrom because that will not stop at the terminating
# null. The UCS-16 assumed to be little endian.
proc twapi::_ucs16_binary_to_string {bin {off 0}} {
    return [encoding convertfrom unicode [string range $bin $off [string first \0\0\0 $bin]]]
}

# Given a binary, return a GUID. The formatting is done as per the
# Windows StringFromGUID2 convention used by COM
proc twapi::_binary_to_guid {bin {off 0}} {
    if {[binary scan $bin "@$off i s s H4 H12" g1 g2 g3 g4 g5] != 5} {
        error "Invalid GUID binary"
    }

    return [format "{%8.8X-%2.2hX-%2.2hX-%s}" $g1 $g2 $g3 [string toupper "$g4-$g5"]]
}

# Given a guid string, return a GUID in binary form
proc twapi::_guid_to_binary {guid} {
    _validate_guid $guid
    lassign [split [string range $guid 1 end-1] -] g1 g2 g3 g4 g5
    return [binary format "i s s H4 H12" 0x$g1 0x$g2 0x$g3 $g4 $g5]
}

# Return a guid from raw memory
proc twapi::_decode_mem_guid {mem {off 0}} {
    return [_binary_to_guid [Twapi_ReadMemoryBinary $mem $off 16]]
}

# Convert a Windows registry value to Tcl form. mem is a raw
# memory object. off is the offset into the memory object to read.
# $type is a integer corresponding
# to the registry types
proc twapi::_decode_mem_registry_value {type mem len {off 0}} {
    set type [expr {$type}];    # Convert hex etc. to decimal form
    switch -exact -- $type {
        1 -
        2 {
            # Note - pass in -1, not $len since we do not
            # want terminating nulls
            return [list [expr {$type == 2 ? "expand_sz" : "sz"}] \
                        [Twapi_ReadMemoryUnicode $mem $off -1]]
        }
        7 {
            # Collect strings until we come across an empty string
            # Note two nulls right at the start will result in
            # an empty list. Should it result in a list with
            # one empty string element? Most code on the web treats
            # it as the former so we do too.
           set multi [list ]
            while {1} {
                set str [Twapi_ReadMemoryUnicode $mem $off -1]
                set n [string length $str]
                # Check for out of bounds. Cannot check for this before
                # actually reading the string since we do not know size
                # of the string.
                if {($len != -1) && ($off+$n+1) > $len} {
                    error "Possible memory corruption: read memory beyond specified memory size."
                }
                if {$n == 0} {
                    return [list multi_sz $multi]
                }
                lappend multi $str
                # Move offset by length of the string and terminating null
                # (times 2 since unicode and we want byte offset)
                incr off [expr {2*($n+1)}]
            }
        }
        4 {
            if {$len < 4} {
                error "Insufficient number of bytes to convert to integer."
            }
            return [list dword [Twapi_ReadMemoryInt $mem $off]]
        }
        5 {
            if {$len < 4} {
                error "Insufficient number of bytes to convert to big-endian integer."
            }
            set type "dword_big_endian"
            set scanfmt "I"
            set len 4
        }
        11 {
            if {$len < 8} {
                error "Insufficient number of bytes to convert to wide integer."
            }
            set type "qword"
            set scanfmt "w"
            set len 8
        }
        0 { set type "none" }
        6 { set type "link" }
        8 { set type "resource_list" }
        3 { set type "binary" }
        default {
            error "Unsupported registry value type '$type'"
        }
    }

    set val [Twapi_ReadMemoryBinary $mem $off $len]
    if {[info exists scanfmt]} {
        if {[binary scan $val $scanfmt val] != 1} {
            error "Could not convert from binary value using scan format $scanfmt"
        }
    }

    return [list $type $val]
}

proc twapi::Twapi_PtrToAddress {p} {
    if {[Twapi_IsPtr $p]} {
        set addr [lindex $p 0]
        if {$addr eq "NULL"} {
            return 0
        } else {
            return $addr
        }
    } else {
        error "'$p' is not a valid pointer value."
    }
}

proc twapi::Twapi_PtrType {p} {
    if {[Twapi_IsPtr $p]} {
        set type [lindex $p 1]
        if {$type eq ""} {
            set type void*
        }
    } else {
        error "'$p' is not a valid pointer value."
    }
    return $type
}

proc twapi::Twapi_AddressToPtr {addr type} {
    return [list $addr $type]
}


proc twapi::_log_timestamp {} {
    return [clock format [clock seconds] -format "%a %T"]
}


# Helper for Net*Enum type functions taking a common set of arguments
proc twapi::_net_enum_helper {function args} {
    if {[llength $args] == 1} {
        set args [lindex $args 0]
    }

    # -namelevel is used internally to indicate what level is to be used
    # to retrieve names. -preargs and -postargs are used internally to
    # add additional arguments at specific positions in the generic call.
    array set opts [parseargs args {
        {system.arg ""}
        level.int
        resume.int
        filter.int
        {namelevel.int 0}
        {preargs.arg {}}
        {postargs.arg {}}
        {namefield.arg name}
    } -maxleftover 0]

    if {[info exists opts(level)]} {
        set level $opts(level)
    } else {
        set level $opts(namelevel)
    }
    if {[info exists opts(resume)]} {
        set resumehandle $opts(resume)
    } else {
        set resumehandle 0
    }

    set moredata 1
    set result {}
    while {$moredata} {
        if {[info exists opts(filter)]} {
            lassign  [$function $opts(system) {*}$opts(preargs) $level $opts(filter) {*}$opts(postargs) $resumehandle] moredata resumehandle totalentries groups
        } else {
            lassign [$function $opts(system) {*}$opts(preargs) $level {*}$opts(postargs) $resumehandle] moredata resumehandle totalentries groups
        }
        # If caller does not want all data in one lump stop here
        if {[info exists opts(resume)]} {
            if {[info exists opts(level)]} {
                return [list $moredata $resumehandle $totalentries $groups]
            } else {
                # Return flat list of names
                return [list $moredata $resumehandle $totalentries [kl_flatten $groups name]]
            }
        }
        # Append to existing result
        # TBD - can the K operator makes this concatnation faster ?
        set result [concat $result $groups]
    }

    # Return what we have. Format depend on caller options.
    if {[info exists opts(level)]} {
        return $result
    } else {
        return [kl_flatten $result $opts(namefield)]
    }
}

# If we are being sourced ourselves, then we need to source the remaining files.
# The apply is just to use vars without polluting global namespace
apply {{filelist} {
    if {[file tail [info script]] eq "twapi.tcl"} {
        # We are being sourced so source the remaining files

        set dir [file dirname [info script]]
        foreach f $filelist {
            uplevel #0 [list source [file join $dir $f]]
        }
    }
}} {base.tcl handle.tcl win.tcl adsi.tcl}

# Returns a list of twapi procs that are currently defined and should
# be exported. SHould be called after completely loading twapi
proc twapi::_get_public_procs {} {

    set public_procs [info procs]

    # Init with C built-ins - there does not seem an easy auto way
    # of getting these. Also, ensembles although probably there is
    # a way of doing this.
    lappend public_procs {*}{
        canonicalize_guid
        is_valid_sid_syntax
        kl_get
        parseargs
        recordarray
        systemtray
        trap
        twine
    }

    # Also export aliases but not "try" as it conflicts
    # with 8.6 try
    foreach p [interp aliases] {
        if {[regexp {twapi::([a-z][^:]*)$} $p _ tail]} {
            if {$tail ne "try"} {
                lappend public_procs $tail
            }
        }
    }

    return $public_procs
}

# Used in various matcher callbacks to signify always include etc.
# TBD - document
proc twapi::true {args} {
    return true
}

namespace eval twapi {
    # Get a handle to ourselves. This handle never need be closed
    variable my_process_handle [GetCurrentProcess]

    # TBD - To improve start-up times, we should really enumerate exports
    # at build time,

    # eval namespace export [::twapi::_get_public_procs]
}

proc twapi::export_public_commands {} {
    uplevel #0 [list namespace eval twapi [list eval namespace export [::twapi::_get_public_procs]]]
}

proc twapi::import_commands {} {
    export_public_commands
    uplevel namespace import twapi::*
}

