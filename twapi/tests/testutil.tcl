# Various utility routines used in the TWAPI tests

package require tcltest

global psinfo;                    # Array storing process information

global thrdinfo;                  # Array storing thread informations

global twapi_test_dir
set twapi_test_script_dir [file dirname [info script]]

proc load_twapi {} {

    # If in source dir, we load that twapi in preference to installed package
    if {[file exists ../tcl/twapi.tcl]} {
        uplevel #0 source ../tcl/twapi.tcl
    } else {
        uplevel #0 package require twapi
    }
}

# From http://mini.net/tcl/460
#
# If you need to split string into list using some more complicated rule
# than builtin split command allows, use following function
# It mimics Perl split operator which allows regexp as element
# separator, but, like builtin split, it expects string to split as
# first arg and regexp as second (optional) By default, it splits by any
# amount of whitespace.

# Note that if you add parenthesis into regexp, parenthesed part of
# separator would be added into list as additional element. Just like in
# Perl. -- cary
proc xsplit [list str [list regexp "\[\t \r\n\]+"]] {
    set list  {}
    while {[regexp -indices -- $regexp $str match submatch]} {
        lappend list [string range $str 0 [expr [lindex $match 0] -1]]
        if {[lindex $submatch 0]>=0} {
            lappend list [string range $str [lindex $submatch 0]\
                              [lindex $submatch 1]]
        }
        set str [string range $str [expr [lindex $match 1]+1] end]
    }
    lappend list $str
    return $list
}

# Validate IP address
proc valid_ip_address {ipaddr} {
    # (Copied from Mastering Regular Expression)
    # Expression to match 0-255
    set sub {([01]?\d\d?|2[0-4]\d|25[0-5])}

    return [regexp "^$sub\.$sub\.$sub\.$sub\$" $ipaddr]
}

# Validate list of ip addresses
proc validate_ip_addresses {addrlist} {
    foreach addr $addrlist {
        if {![valid_ip_address $addr]} {return 0}
    }
    return 1
}

proc valid_handle {h} {
    return [twapi::Twapi_IsPtr $h]
}

# Validate SIDs
proc valid_sids {sids} {
    foreach sid $sids {
        if {[catch {twapi::lookup_account_sid $sid}]} {
            return 0
        }
    }
    return 1
}

proc valid_account_names {names} {
    foreach name $names {
        if {[catch {twapi::lookup_account_name $name}]} {
            if {$name ne "Logon SID"} {
                return 0
            }
        }
    }
    return 1
}

proc system_drive_root {} {
    return [file dirname $::env(WINDIR)]
}

# Start notepad and wait till it's up and running.
proc notepad_exec {args} {
    set pid [eval [list exec notepad.exe] $args &]
    if {![twapi::process_waiting_for_input $pid -wait 5000]} {
        error "Timeout waiting for notepad to be ready for input"
    }
    return $pid
}

# Start notepad, make it store something in the clipboard and exit
proc notepad_copy {text} {
    set pid [notepad_exec]
    set hwin [lindex [twapi::find_windows -pids [list $pid] -class Notepad] 0]
    twapi::set_foreground_window $hwin
    after 100;                          # Wait for it to become foreground
    twapi::send_keys $text
    twapi::send_keys ^a^c;                 # Select all and copy
    after 100
    twapi::end_process $pid -force
}

# Start notepad, make it add text and return its pid
proc notepad_exec_and_insert {{text "Some junk"}} {
    set pid [notepad_exec]
    set hwins [twapi::find_windows -pids [list $pid] -class Notepad]
    twapi::set_foreground_window [lindex $hwins 0]
    after 100;                          # Wait for it to become foreground
    twapi::send_keys $text
    after 100
    return $pid
}

# Find the notepad window for a notepad process
proc notepad_top {np_pid} {
    return [twapi::find_windows -class Notepad -pids [list $np_pid] -single]
}

# Find the popup window for a notepad process
proc notepad_popup {np_pid} {
    return [twapi::find_windows -text Notepad -pids [list $np_pid] -single]
}


proc get_processes {{refresh 0}} {
    global psinfo
    
    if {[info exists psinfo(0)] && ! $refresh} return
    
    catch {unset psinfo}
    array set psinfo {} 
    set fd [open "| cscript.exe /nologo process.vbs"]
    while {[gets $fd line] >= 0} {
        if {[string length $line] == 0} continue
        if {[catch {array set processinfo [split $line "*"]} msg]} {
            puts stderr "Error parsing line: '$line': $msg"
            error $msg
        }

        set pid $processinfo(ProcessId)
        set psinfo($pid) \
            [list \
                 -basepriority $processinfo(Priority) \
                 -handlecount  $processinfo(HandleCount) \
                 -name         $processinfo(Name) \
                 -pagefilebytes $processinfo(PageFileUsage) \
                 -pagefilebytespeak $processinfo(PeakPageFileUsage) \
                 -parent       $processinfo(ParentProcessId) \
                 -path         $processinfo(ExecutablePath) \
                 -pid          $pid \
                 -poolnonpagedbytes $processinfo(QuotaNonPagedPoolUsage) \
                 -poolpagedbytes $processinfo(QuotaPagedPoolUsage) \
                 -privatebytes $processinfo(PrivatePageCount) \
                 -threadcount  $processinfo(ThreadCount) \
                 -virtualbytes     $processinfo(VirtualSize) \
                 -virtualbytespeak $processinfo(PeakVirtualSize) \
                 -workingset       $processinfo(WorkingSetSize) \
                 -workingsetpeak   $processinfo(PeakWorkingSetSize) \
                 -user             $processinfo(User) \
                ]
    }
    close $fd
}

# Get given field for the given pid. Error if pid does not exist
proc get_process_field {pid field {refresh 0}} {
    global psinfo
    get_processes $refresh
    if {![info exists psinfo($pid)]} {
        error "Pid $pid does not exist"
    }
    return [get_kl_field $psinfo($pid) $field]
}

# Get the first pid with the given value (case insensitive)
# in the given field
proc get_process_with_field_value {field value {refresh 0}} {
    global psinfo
    get_processes $refresh
    foreach pid [array names psinfo] {
        if {[string equal -nocase [get_process_field $pid $field] $value]} {
            return $pid
        }
    }
    error "No process with $field=$value"
}

proc get_winlogon_path {} {
    set winlogon_path [file join $::env(WINDIR) "system32" "winlogon.exe"]
    return [string tolower [file nativename $winlogon_path]]
}

proc get_winlogon_pid {} {
    global winlogon_pid
    if {! [info exists winlogon_pid]} {
        set winlogon_pid [get_process_with_field_value -name "winlogon.exe"]
    }
    return $winlogon_pid
}

proc get_explorer_path {} {
    set explorer_path [file join $::env(WINDIR) "explorer.exe"]
    return [string tolower [file nativename $explorer_path]]
}

proc get_explorer_pid {} {
    global explorer_pid
    if {! [info exists explorer_pid]} {
        set explorer_pid [get_process_with_field_value -name "explorer.exe"]
    }
    return $explorer_pid
}

proc get_explorer_tid {} {
    return [get_thread_with_field_value -pid [get_explorer_pid]]
}

proc get_notepad_path {} {
    set path [auto_execok notepad.exe]
    return [string tolower [file nativename $path]]
}

proc get_cmd_path {} {
    set path [auto_execok cmd.exe]
    return [string tolower [file nativename $path]]
}

proc get_temp_path {{name ""}} {
    return [file join $::tcltest::temporaryDirectory $name]
}

proc get_system_pid {} {
    global system_pid
    global psinfo
    if {! [info exists system_pid]} {
        set system_pid [get_process_with_field_value -name "System"]
    }
    return $system_pid
}

proc get_idle_pid {} {
    global idle_pid
    global psinfo
    if {! [info exists idle_pid]} {
        set idle_pid [get_process_with_field_value -name "System Idle Process"]
    }
    return $idle_pid
}

proc get_threads {{refresh 0}} {
    global thrdinfo
    
    if {[info exists thrdinfo] && ! $refresh} return
    catch {unset thrdinfo}
    array set thrdinfo {} 
    set fd [open "| cscript.exe /nologo thread.vbs"]
    while {[gets $fd line] >= 0} {
        if {[string length $line] == 0} continue
        array set threadrec [split $line "*"]
        set tid $threadrec(Handle)
        set thrdinfo($tid) \
            [list \
                 -tid $tid \
                 -basepriority $threadrec(PriorityBase) \
                 -pid          $threadrec(ProcessHandle) \
                 -priority     $threadrec(Priority) \
                 -startaddress $threadrec(StartAddress) \
                 -state        $threadrec(ThreadState) \
                 -waitreason   $threadrec(ThreadWaitReason) \
                ]
    }
    close $fd
}

# Get given field for the given tid. Error if tid does not exist
proc get_thread_field {tid field {refresh 0}} {
    global thrdinfo
    get_threads $refresh
    if {![info exists thrdinfo($tid)]} {
        error "Thread $tid does not exist"
    }
    return [get_kl_field $thrdinfo($tid) $field]
}

# Get the first tid with the given value (case insensitive)
# in the given field
proc get_thread_with_field_value {field value {refresh 0}} {
    global thrdinfo
    get_threads $refresh
    foreach tid [array names thrdinfo] {
        if {[string equal -nocase [get_thread_field $tid $field] $value]} {
            return $tid
        }
    }
    error "No thread with $field=$value"
}

# Get list of threads for the given process
proc get_process_tids {pid {refresh 0}} {
    global thrdinfo
    
    get_threads $refresh
    set tids [list ]
    foreach {tid rec} [array get thrdinfo] {
        array set thrd $rec
        if {$thrd(-pid) == $pid} {
            lappend tids $tid
        }
    }
    return $tids
}


# Start the specified program and return its pid
proc start_program {exepath args} {
    set pid [eval exec [list $exepath] $args &]
    # Wait to ensure it has started up
    if {![twapi::wait {twapi::process_exists $pid} 1 1000]} {
        error "Could not start $exepath"
    }
    # delay to let it get fully initialized.
    after 100
    return $pid
}

# Compare two strings as paths
proc equal_paths {p1 p2} {
    # Use file join to convert \ to /
    return [string equal -nocase [file join $p1] [file join $p2]]
}
tcltest::customMatch path equal_paths

# Compare two sets (dup elements are treated as same)
proc equal_sets {s1 s2} {
    set s1 [lsort -unique $s1]
    set s2 [lsort -unique $s2]
    if {[llength $s1] != [llength $s2]} {
        return 0
    }

    foreach e1 $s1 e2 $s2 {
        if {[string compare $e1 $e2]} {
            return 0
        }
    }

    return 1
}
#
# Custom proc for matching file paths
tcltest::customMatch set equal_sets


# Note - use single quotes, not double quotes to pass values to wmic from exec
proc wmic_delete {obj clause} {
    # The cmd echo is required because otherwise wmic hangs for some obscure
    # reason when spawned from a non-interactive tclsh
    exec cmd /c echo . | wmic path $obj where $clause delete
}

# Note - use single quotes, not double quotes to pass values to wmic from exec
proc wmic_get {obj fields {clause ""}} {

    # On some systems when invoking wmic
    # the "cmd echo..." is required because otherwise wmic hangs for some 
    # reason when spawned from a non-interactive tclsh
    # On the other hand, if this is done, tests cannot be excuted from 
    # a read-only dir. So we have both versions here, with one or the other
    # commented out
    if {$clause eq ""} {
        set lines [exec cmd /c echo . | wmic path $obj get [join $fields ,] /format:csv]
        #set lines [exec wmic path $obj get [join $fields ,] /format:csv]
    } else {
        set lines [exec cmd /c echo . | wmic path $obj where $clause get [join $fields ,] /format:csv]
        #set lines [exec wmic path $obj where $clause get [join $fields ,] /format:csv]
    }


    set data {}
    foreach line [split $lines \n] {
        set line [string trim $line]
        if {$line eq ""} continue
        # Assumes no "," in content
        lappend data [split $line ,]
    }

    # First element is field names, not in same order as $fields. Also,
    # Case might be different. Make them consistent with what caller
    # expects. Code below assumes no duplicate names
    set fieldnames {}
    foreach fname [lindex $data 0] {
        set fieldname $fname
        foreach fname2 $fields {
            if {[string equal -nocase $fname $fname2]} {
                set fieldname $fname2
                break
            }
        }
        lappend fieldnames $fieldname
    }

    set result {}
    foreach values [lrange $data 1 end] {
        # wmic seems to html-encode when outputting in csv format
        # We do minimal necessary for our test scripts
        set decoded_values {}
        foreach value $values {
            lappend decoded_values [string map {&amp; &} $value]
        }
        lappend result [twapi::twine $fieldnames $decoded_values]
    }

    return $result
}

# Return 1/0 depending on whether at least one record with specified field
# value exists in wmic table
# Note - use single quotes, not double quotes to pass values to wmic from exec
proc wmic_exists {obj field value} {
    if {[catch {
        wmic_get $obj [list $field] "$field='$value'"
    } msg]} {
        if {[string match -nocase "*No Instance(s) available*" $msg]} {
            return 0
        }
        error $msg $::errorInfo $::errorCode
    }

    return 1
}

# Returns value of a field in the first record with specified key
proc wmic_value {obj field key keyval} {
    # Some WMI fields have to be retrieved using *, not the name. For example
    # the FullName field from Win32_Account. Dunno why

    array set rec [lindex [wmic_get $obj [list $field] "$key='$keyval'"] 0]
    return $rec($field)
}

# Return true if $a is close to $b (within 5%)
proc approx {a b {adjust 0}} {
    set max [expr {$a > $b ? $a : $b}]; # Tcl 8.4 does not have a max() func
    if {[expr {abs($b-$a) < ($max/20)}]} {
        return 1
    }
    if {! $adjust} {
        return 0
    }

    # Scale whichever one is smaller
    if {$a < $b} {
        set a [expr {$a * $adjust}]
    } else {
        set b [expr {$b * $adjust}]
    }

    # See if they match up after adjustment
    set max [expr {$a > $b ? $a : $b}]; # Tcl 8.4 does not have a max() func
    return [expr {abs($b-$a) < ($max/20)}]
}


# Return a field is keyed list
proc get_kl_field {kl field} {
    foreach {fld val} $kl {
        if {$fld == $field} {
            return $val
        }
    }
    error "No field $field found in keyed list"
}

#
# Verify that a keyed list has the specified fields
# Raises an error otherwise
proc verify_kl_fields {kl fields {ignoreextra 0}} {
    array set data $kl
    foreach field $fields {
        if {![info exists data($field)]} {
            error "Field $field not found keyed list <$kl>"
        }
        unset data($field)
    }
    if {$ignoreextra} {
        return
    }
    set extra [array names data]
    if {[llength $extra]} {
        puts stderr "Extra fields ([join $extra ,]) found in keyed list"
        error "Extra fields ([join $extra ,]) found in keyed list"
    }
    return
}

#
# Verify that all elements in a list of keyed lists have
# the specified fields
# Raises an error otherwise
proc verify_list_kl_fields {l fields {ignoreextra 0}} {
    foreach kl $l {
        verify_kl_fields $kl $fields $ignoreextra
    }
}

#
# Verify is an integer pair
proc verify_integer_pair {pair} {
    if {([llength $pair] != 2) || 
        (![string is integer [lindex $pair 0]]) ||
        (![string is integer [lindex $pair 1]]) } {
        error "'$pair' is not a pair of integers"
    }
    return
}

# Return true if all items in a list look like privileges
proc verify_priv_list {privs} {
    set match 1
    foreach priv $privs {
        set match [expr {$match && [string match Se* $priv]}]
    }
    return $match
}


# Read commands from standard input and execute them.
# From Welch.
proc start_commandline {} {
    set ::command_line ""
    fileevent stdin readable [list eval_commandline]
    # We need a vwait for events to fire!
    vwait ::exit_command_loop
}

proc eval_commandline {} {
    if {[eof stdin]} {
        exit
    }
    
    append ::command_line [gets stdin]
    if {[info complete $::command_line]} {
        catch {uplevel \#0 $::command_line[set ::command_line ""]} result
    } else {
        # Command not complete
        append ::command_line "\n"
    }
}

# Stops the command line loop
proc stop_commandline {} {
    set ::exit_command_loop 1
    set ::command_line ""
    fileevent stdin readable {}
}

# Starts a Tcl shell that will read commands and execute them
proc tclsh_slave_start {} {
    set fd [open "| [list [::tcltest::interpreter]]" r+]
    fconfigure $fd -buffering line -blocking 0 -eofchar {}
    tclsh_slave_verify_started $fd
    #puts $fd [list source [file join $::twapi_test_script_dir testutil.tcl]]
    #puts $fd start_commandline
    return $fd
}
proc tclsh_slave_verify_started {fd} {
    # Verify started. Note we need the puts because tclsh does
    # not output result unless it is a tty.
    tclsh_slave_puts $fd {
        source testutil.tcl
        load_twapi
        if {[catch {
            fconfigure stdout -buffering line -encoding utf-8
            fconfigure stdin -buffering line -encoding utf-8 -eofchar {}
            puts [info tclversion]
            flush stdout
        }]} {
            testlog Error
            testlog $::errorInfo
        }
    }

    if {[catch {
        set ver [gets_timeout $fd]
    } msg]} {
        #close $fd
        testlog $msg
        error $msg $::errorInfo $::errorCode
    }
    if {$ver ne [info tclversion]} {
        error "Slave Tcl version $ver does not match."
    }

    return $fd
}

# Send a command to the slave
proc tclsh_slave_puts {fd cmd} {
    if {[string index $cmd end] == "\n"} {
        puts -nonewline $fd $cmd
    } else {
        puts $fd $cmd
    }
    flush $fd
    # Need an update to get around a Tcl bug (Bug 3059220)
    update
}

proc tclsh_slave_stop {fd} {
    tclsh_slave_puts $fd "exit"
    close $fd
}

# Read a line from the specified fd
# fd expected to be non-blocking and line buffered
# Raises error after timeout
proc gets_timeout {fd {ms 1000}} {
    set elapsed 0
    while {$elapsed < $ms} {
        if {[gets $fd line] == -1} {
            if {[eof $fd]} {
                testlog "get_timeout: unexpected eof"
                error "Unexpected EOF reading from $fd."
            }
            after 50;           # Wait a bit and then retry
            incr elapsed 50
        } else {
            return $line
        }
    }

    error "Time out reading from $fd."
}

# Wait until $fd returns specified output. Discards any intermediate input.
# ms is not total timeout, rather it's max time to wait for single read.
# As long as remote keeps writing, we will keep reading.
# fd expected to be non-blocking and line buffered
proc expect {fd expected {ms 1000}} {
    set elapsed 0
    while {true} {
        set data [gets_timeout $fd $ms]
        if {$data eq $expected} {
            return
        }
        # Keep going
    }
}

# Wait for the slave to get ready. Discards any intermediate input.
# ms is not total timeout, rather it's max time to wait for single read.
# As long as slave keeps writing, we will keep reading.
proc tclsh_slave_wait {fd {ms 1000}} {
    set marker "Ready: [clock clicks]"
    tclsh_slave_puts $fd "puts {$marker}"
    set elapsed 0
    while {$elapsed < $ms} {
        set data [gets_timeout $fd $ms]
        if {$data eq $marker} {
            return
        }
        # Keep going
    }
}

# Used for matching results
proc oneof {allowed_values value} {
    return [expr {[lsearch -exact $allowed_values $value] >= 0}]
}
tcltest::customMatch oneof oneof

proc inrange {range value} {
    foreach {low high} $range break
    expr {$value >= $low && $value <= $high}
}
tcltest::customMatch inrange inrange


# Log a test debug message
proc testlog {msg} {
    if {![info exists ::testlog_fd]} {
        set ::testlog_fd [open testlog-[pid].log w+]
        set ::testlog_time [clock clicks]
    }
    puts $::testlog_fd "[expr {[clock clicks]-$::testlog_time}]: $msg"
    flush $::testlog_fd
}



#####
#
# "SetOps, Code, 8.x v2"
# http://wiki.tcl.tk/1763
#
#
#####


# ---------------------------------------------
# SetOps -- Set operations for Tcl
#
# (C) c.l.t. community, 1999
# (C) TclWiki community, 2001
#
# ---------------------------------------------
# Implementation variant for tcl 8.x and beyond.
# Uses namespaces and 'unset -nocomplain'
# ---------------------------------------------
# NOTE: [set][array names] in the {} array is faster than
#   [set][info locals] for local vars; it is however slower
#   for [info exists] or [unset] ...

namespace eval ::setops {
    namespace export {[a-z]*}
}

proc ::setops::create {args} {
    cleanup $args
}

proc ::setops::cleanup {A} {
    # unset A to avoid collisions
    foreach [lindex [list $A [unset A]] 0] {.} {break}
    info locals
}

proc ::setops::union {args} {
    switch [llength $args] {
	 0 {return {}}
	 1 {return [lindex $args 0]}
    }

   foreach setX $args {
	foreach x $setX {set ($x) {}}
   }
   array names {}
}

proc ::setops::diff {A B} {
    if {[llength $A] == 0} {
	 return {}
    }
    if {[llength $B] == 0} {
	 return $A
    }

    # get the variable B out of the way, avoid collisions
    # prepare for "pure list optimisation"
    set ::setops::tmp [lreplace $B -1 -1 unset -nocomplain]
    unset B

    # unset A early: no local variables left
    foreach [lindex [list $A [unset A]] 0] {.} {break}

    eval $::setops::tmp

    info locals
}

proc ::setops::contains {set element} {
   expr {[lsearch -exact $set $element] < 0 ? 0 : 1}
}

proc ::setops::symdiff {A B} {
    union [diff $A $B] [diff $B $A]
}

proc ::setops::empty {set} {
   expr {[llength $set] == 0}
}

proc ::setops::intersect {args} {
   set res  [lindex $args 0]
   foreach set [lrange $args 1 end] {
	if {[llength $res] && [llength $set]} {
	    set res [Intersect $res $set]
	} else {
	    break
	}
   }
   set res
}

proc ::setops::Intersect {A B} {
# This is slower than local vars, but more robust
    if {[llength $B] > [llength $A]} {
	 set res $A
	 set A $B
	 set B $res
    }
    set res {}
    foreach x $A {set ($x) {}}
    foreach x $B {
	 if {[info exists ($x)]} {
	     lappend res $x
	 }
    }
    set res
}


# If this is the first argument to the shell and there are more arguments
# execute them
if {[string equal -nocase [file normalize $argv0] [file normalize [info script]]]} {
    load_twapi
    if {[catch {
        foreach arg $argv {
            eval $arg
        }
    } msg]} {
        twapi::eventlog_log "testutil error: $msg"
    }
}
