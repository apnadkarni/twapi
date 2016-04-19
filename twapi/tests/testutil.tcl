# Various utility routines used in the TWAPI tests

package require tcltest

global psinfo;                    # Array storing process information

global thrdinfo;                  # Array storing thread informations

interp alias {} testdir {} lindex [file normalize [file dirname [info script]]]

proc new_name {} {
    variable _test_name_counter
    # Just using [clock microseconds] is not enough granularity
    return TwapiTest-[clock microseconds]-[incr _test_name_counter]
}

proc errorcode {} {
    return [lrange $::errorCode 0 1]
}

proc getenv {envvar {default ""}} {
    if {[info exists ::env($envvar)]} {
        return $::env($envvar)
    } else {
        return $default
    }
}

# Returns if current system is part of domain
proc indomain {} {
    if {[info exists ::env(USERDNSDOMAIN)] ||
        [string compare -nocase \\\\$::env(COMPUTERNAME) $::env(LOGONSERVER)]} {
        return 1
    } else {
        return 0
    }
}


proc testconfig {item} {
    global testconfig
    if {![info exists testconfig($item)]} {
        switch -exact -- $item {
            domain_user {
                set testconfig($item) [getenv TWAPI_TEST_DOMAIN_USER aduser]
            }
            domain_name     {
                set testconfig($item) [getenv TWAPI_TEST_DOMAIN TEST]
            }
            domain_dnsname  {
                set testconfig($item) [getenv TWAPI_TEST_DNSDOMAIN [string tolower [testconfig domain_name]].twapi]
            }
            domain_controller {
                # Returns name of domain controller if there is one on the
                # network. Used even if current system is not in the domain

                if {[indomain]} {
                    set testconfig($item) [string trimleft $::env(LOGONSERVER) \\]
                } else {
                    # Not in domain, try to find one on network
                    set dc [getenv TWAPI_TEST_DOMAINCONTROLLER win2k8-adserver]
                    if {[catch {
                        exec ping $dc -n 1
                    }]} {
                        set testconfig($item) ""
                    } else {
                        set testconfig($item) $dc
                    }
                }
            }
            default { error "Unknown config item '$item'" }
        }
    }
    return $testconfig($item)
}

proc find_unused_drive {} {
    set drives [wmic_values Win32_LogicalDisk name]
    foreach drv [split KLMNOPQRSTUVWXYZABCDEFGHIJ ""] {
        if {[lsearch -exact $drives "${drv}:"] < 0} {
            return ${drv}:
        }
    }
    error "No free drive letters found."
}

proc equal_boolean {a b} {
    return [expr {(! $a) == (! $b)}]
}
tcltest::customMatch boolean equal_boolean

proc load_twapi_package {{pkg twapi}} {
    global _twapi_test_loaded_packages

    if {[info exists _twapi_test_loaded_packages($pkg)]} {
        return
    }

    # If in source dir, we load that twapi in preference to installed package
    if {[file exists ../dist] &&
        ! [info exists _twapi_test_loaded_packages]} {
        set ::auto_path [linsert $::auto_path 0 [file normalize ../dist]]
    }

    package require $pkg
    set _twapi_test_loaded_packages($pkg) 1
}

proc write_test_file {content {mode wb}} {
    set path [tcltest::makeFile "" twapitest-[clock microseconds]]
    set fd [open $path $mode]
    puts -nonewline $fd $content
    close $fd
    return $path
}

proc read_file {path {mode r}} {
    set fd [open $path $mode]
    set data [read $fd]
    close $fd
    return $data
}

proc read_binary {path} {
    return [read_file $path rb]
}

proc write_file {path content {mode w}} {
    set fd [open $path $mode]
    puts -nonewline $fd $content
    close $fd
}

proc write_binary {path content} {
    return [write_file $path $content wb]
}

# Create a new user with a random password
proc create_user_with_password {uname {system ""}} {
    # Sometimes the password will be rejected because it contains
    # a substring of the user name
    while {1} {
        twapi::trap {
            twapi::new_user $uname -password [twapi::new_uuid] -system $system
            break
        } onerror {TWAPI_WIN32 2245} {
            # Loop and retry
        }
    }
}

proc name2sid {name} {
    variable testutil_sids
    if {![info exists testutil_sids($name)]} {
        set testutil_sids($name) [wmic_value Win32_Account sid name $name]
    }
    return $testutil_sids($name)
}

# Get the localized account name for a well known account
proc get_localized_account {name} {
    switch -exact -- [string tolower $name] {
        administrator {set sddl "O:LA"}
        administrators {set sddl "O:BA"}
        guest {set sddl "O:LG"}
        guests {set sddl "O:BG"}
        default {
            error "Do not know how to localize account $name"
        }
    }

    return [twapi::map_account_to_name [twapi::get_security_descriptor_owner [twapi::sddl_to_security_descriptor $sddl]]]
}

# Populate users on a system
proc populate_accounts {{system ""}} {
    variable populated_accounts

    if {! [info exists populated_accounts($system)]} {
        patience "Creation of test accounts"
        set uname TWAPI_[clock seconds]
        set gname ${uname}_GROUP
        twapi::new_local_group $gname -system $system
        lappend populated_accounts($system) $gname

        for {set i 0} {$i < 500} {incr i} {
            set u ${uname}_$i
            create_user_with_password $u $system
            twapi::add_member_to_local_group $gname $u -system $system
            lappend populated_accounts($system) $u
        }
    }

    # Return number of accounts (first elem is group name)
    return [expr {[llength $populated_accounts($system)] - 1}]
}

# Populate global groups on a system
proc populate_global_groups {{system ""}} {
    variable populated_global_groups

    if {! [info exists populated_global_groups($system)]} {
        patience "Creation of test global groups"
        set gname TWAPI_[clock seconds]
        for {set i 0} {$i < 500} {incr i} {
            set g ${gname}_$i
            twapi::new_global_group $g -system $system -comment "TwAPI Test group $g"
            lappend populated_global_groups($system) $g
        }
    }

    # Return number of groups created
    return [llength $populated_global_groups($system)]
}


# Populate local groups on a system
proc populate_local_groups {{system ""}} {
    variable populated_local_groups

    if {! [info exists populated_local_groups($system)]} {
        patience "Creation of test local groups"
        set gname TWAPI_[clock seconds]
        for {set i 0} {$i < 500} {incr i} {
            set g ${gname}_$i
            twapi::new_local_group $g -system $system -comment "TwAPI Test group $g"
            lappend populated_local_groups($system) $g
        }
    }

    # Return number of groups created
    return [llength $populated_local_groups($system)]
}

proc cleanup_populated_accounts_and_groups {} {
    variable populated_accounts
    variable populated_local_groups
    variable populated_global_groups
    
    patience "Cleaning up test accounts"

    set failures 0
    foreach {system accounts} [array get populated_accounts] {
        # First elem is group, remaining user accounts
        foreach uname [lrange $accounts 1 end] {
            incr failures [catch {twapi::delete_user $uname -system $system}]
        }
        incr failures [catch {twapi::delete_local_group [lindex $accounts 0] -system $system}]
        unset populated_accounts($system)
    }

    foreach {system groups} [array get populated_local_groups] {
        foreach gname $groups {
            incr failures [catch {twapi::delete_local_group $gname -system $system}]
        }
        unset populated_local_groups($system)
    }

    foreach {system groups} [array get populated_global_groups] {
        foreach gname $groups {
            incr failures [catch {twapi::delete_global_group $gname -system $system}]
        }
        unset populated_global_groups($system)
    }


    catch {unset populated_accounts}
    catch {unset populated_local_groups}
    catch {unset populated_global_groups}

    if {$failures} {
        error "$failures failures cleaning up test accounts"
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
proc valid_ip_address {ipaddr {ipver 0}} {
    load_twapi_package twapi_network

    set addrver [twapi::get_ipaddr_version $ipaddr]
    if {$addrver == 0} { return 0 }
    if {$ipver && $addrver != $ipver} { return 0 }
    return 1
}

# Validate list of ip addresses
proc validate_ip_addresses {addrlist {ipver 0}} {
    foreach addr $addrlist {
        if {![valid_ip_address $addr $ipver]} {return 0}
    }
    return 1
}
interp alias {} valid_ip_addresses {} validate_ip_addresses

proc valid_handle {h} {
    return [twapi::pointer? $h]
}

proc valid_typed_handle {type h} {
    return [twapi::pointer? $h $type]
}
tcltest::customMatch handle valid_typed_handle


# Validate SIDs
proc valid_sids {sids} {
    foreach sid $sids {
        if {[catch {twapi::lookup_account_sid $sid}]} {
            return 0
        }
    }
    return 1
}

proc valid_account_names {names {system ""}} {
    if {[llength $names] == 0} {
        error "List of accounts passed in is empty."
    }
    foreach name $names {
        if {[catch {twapi::lookup_account_name $name -system $system}]} {
            if {$name ne "Logon SID"} {
                return 0
            }
        }
    }
    return 1
}

proc valid_account_sids {sids {system ""}} {
    if {[llength $sids] == 0} {
        error "List of accounts passed in is empty."
    }
    foreach sid $sids {
        if {[catch {twapi::lookup_account_sid $sid -system $system}]} {
            return 0
        }
    }
    return 1
}

proc system_drive_root {} {
    return [file dirname $::env(WINDIR)]
}

# $pid may be PID or image name
proc kill {pid args} {
    # Note we do not want to use twapi to keep testing of modules
    # not dependent on each other

    if {[string is integer $pid]} {
        set opt /PID
    } else {
        set opt /IM
    }

    if {[lsearch -exact $args "-force"] >= 0} {
        exec taskkill $opt $pid /F
    } else {
        exec taskkill $opt $pid
    }

    set wait [lsearch -exact $args "-wait"]
    if {$wait < 0} { return }
    set wait [lindex $args $wait+1]

    # Wait in a loop checking for process existence WITHOUT using twapi
    while {$wait > 0 && [process_exists? $pid]} {
        after 50
        incr wait -50
    }
}

# Start notepad and wait till it's up and running.
proc notepad_exec {args} {
    set pid [eval [list exec [get_notepad_path]] $args &]
    if {[info commands twapi::process_waiting_for_input] ne ""} {
        if {![twapi::process_waiting_for_input $pid -wait 5000]} {
            error "Timeout waiting for notepad to be ready for input"
        }
    } else {
        # Assume ready after 2000
        set wait 2000
        while {$wait > 0 && ! [process_exists? $pid]} {
            after 50
            incr wait -50
        }
        # Once pid is there, wait another 100 for window to be ready for input
    }
    return $pid
}

# Start notepad, make it store something in the clipboard and exit
proc notepad_copy {text} {
    set pid [notepad_exec]
    if {[info commands twapi::find_windows] ne ""} {
        set hwin [lindex [twapi::find_windows -pids [list $pid] -class Notepad] 0]
        twapi::set_foreground_window $hwin
    } else  {
        # Assume the exec put it in foreground
    }
    twapi::send_keys $text
    twapi::send_keys ^a^c;                 # Select all and copy
    after 100
    kill $pid
}

# Start notepad, make it add text and return its pid
proc notepad_exec_and_insert {{text "Some junk"}} {
    # TBD - modify to get rid of both these package requirements
    package require twapi_input
    package require twapi_ui

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
    return [wait_for_window -class Notepad -pids [list $np_pid]]
    return [twapi::find_windows -class Notepad -pids [list $np_pid] -single]
}

# Find the popup window for a notepad process
proc notepad_popup {np_pid} {
    return [twapi::find_windows -text Notepad -pids [list $np_pid] -single]
}

# Find the PID for a window with specified title 
proc window_to_pid {args} {
    return [twapi::get_window_process [eval wait_for_window $args]]
}

# Wait for a window to become visible and return it
proc wait_for_window args {
    set elapsed 0
    while {$elapsed < 15000} {
        set wins [eval twapi::find_windows $args]
        if {[llength $wins] > 1} {
            error "More than one matching window found for '$args'"
        }

        if {[llength $wins] == 1} {
            return [lindex $wins 0]
        }
        
        # When waiting for some windows, like those created through
        # shell_execute, the event loop must be running so that window
        # messages can be processed (COM dispatching for shell controls
        # for instance)
        set after_id [after 100 "set ::wait_for_window_flag 1"]
        vwait ::wait_for_window_flag
        after cancel $after_id; # Probably no need but ...
        incr elapsed 100
    }

    error "No matching window found for '$args'"
}

proc wait_for_visible_toplevel args {
    return [eval [list wait_for_window -toplevel 1 -visible 1] $args]
}

proc get_processes {{refresh 0}} {
    global psinfo

    if {[info exists psinfo(0)] && ! $refresh} return
    
    catch {unset psinfo}
    array set psinfo {} 
    # WHen  tests are run in quick succession cscript sometimes fails
    # with a SWBemblah blah not found error. So just delay a bit
    after 500
    set fd [open "| cscript.exe /nologo process.vbs"]
    while {[gets $fd line] >= 0} {
        if {[string length $line] == 0} continue
        if {[catch {array set processinfo [split $line "*"]} msg]} {
            puts stderr "Error parsing line: '$line': $msg"
            error $msg
        }

        # Note some values are returned in kilobytes by wmi so we scale
        # them accordingly
        set pid $processinfo(ProcessId)
        set psinfo($pid) \
            [list \
                 -basepriority $processinfo(Priority) \
                 -handlecount  $processinfo(HandleCount) \
                 -name         $processinfo(Name) \
                 -pagefilebytes [expr {1024*$processinfo(PageFileUsage)}] \
                 -pagefilebytespeak [expr {1024*$processinfo(PeakPageFileUsage)}] \
                 -parent       $processinfo(ParentProcessId) \
                 -path         $processinfo(ExecutablePath) \
                 -pid          $pid \
                 -poolnonpagedbytes [expr {1024*$processinfo(QuotaNonPagedPoolUsage)}] \
                 -poolpagedbytes [expr {1024*$processinfo(QuotaPagedPoolUsage)}] \
                 -privatebytes $processinfo(PrivatePageCount) \
                 -threadcount  $processinfo(ThreadCount) \
                 -virtualbytes     $processinfo(VirtualSize) \
                 -virtualbytespeak $processinfo(PeakVirtualSize) \
                 -workingset       $processinfo(WorkingSetSize) \
                 -workingsetpeak   [expr {1024*$processinfo(PeakWorkingSetSize)}] \
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

proc process_exists? pid {
    # TBD - maybe use wmic instead
    set ret [catch {get_process_with_field_value -pid $pid 1}]
    return [expr {! $ret}]
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

# See if two files are the same
proc same_file {fnA fnB} {
    set fnA [file normalize $fnA]
    set fnB [file normalize $fnB]

    # Try to convert to full name. This works only if file exists
    if {[file exists $fnA]} {
        if {![file exists $fnB]} {
            return 0
        }
        set fnA [file attributes $fnA -longname]
        set fnB [file attributes $fnB -longname]
    } else {
        if {[file exists $fnB]} {
            return 0
        }
    }

    return [equal_paths $fnA $fnB]
}

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

proc valid_guid {guid} {
     return [regexp {^\{[[:xdigit:]]{8}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{12}\}$} $guid]
}


# Exec a wmic command
proc wmic_exec {wmi_cmdline} {
    # Unfortunately, different variations have been tried to get this working
    # The current version is first. Remaining versions below are commented
    # out with their problems.

    if {1} {
        set fd [open "| cmd /c echo . | $wmi_cmdline"]
        fconfigure $fd -translation binary
        set lines [read $fd]
        close $fd
        return $lines
    } else {
        # On some systems when invoking wmic
        # the "cmd echo..." is required because otherwise wmic hangs for some 
        # reason when spawned from a non-interactive tclsh
        # On the other hand, if this is done, tests cannot be excuted from 
        # a read-only dir. So we have both versions here, with one or the other
        # commented out
        
        set lines [exec cmd /c echo . | {*}$wmi_cmdline]
         #set lines [exec wmic path $obj get [join $fields ,] /format:list]
     }
}

# Note - use single quotes, not double quotes to pass values to wmic from exec
proc wmic_delete {obj clause} {
    # The cmd echo is required because otherwise wmic hangs for some obscure
    # reason when spawned from a non-interactive tclsh
    exec cmd /c echo . | wmic path $obj where $clause delete
}

# Note - use single quotes, not double quotes to pass values to wmic from exec
proc wmic_get {obj {fields *} {clause ""}} {

    if {0} {
        # Commented out because wmic does not properly escape comma separators
        # if they appear in values
        if {$clause eq ""} {
            set lines [wmic_exec "wmic path $obj get [join $fields ,] /format:csv"]
        } else {
            set lines [wmic_exec "wmic path $obj where $clause get [join $fields ,] /format:csv"]
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
        # expects.
        # Code below assumes no duplicate names
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
            set dict {}
            foreach fieldname $fieldnames decoded_value $decoded_values {
                lappend dict $fieldname $decoded_value
            }
            lappend result $dict
        }

        return $result

    } else {

        if {$clause eq ""} {
            set lines [wmic_exec "wmic path $obj get [join $fields ,] /format:list"]
        } else {
            set lines [wmic_exec "wmic path $obj where $clause get [join $fields ,] /format:list"]
        }

        # Data is returned with blank lines separating records. Each
        # record is a sequence of lines of the form FIELDNAME=...
        # Code below assumes no duplicate names
        set records {}
        set rec {}
        foreach line [split $lines \n] {
            set line [string trim $line]
            if {[string length $line] == 0} {
                # End of a record or just a blank line
                if {[llength $rec]} {
                    # End of a record since we have collected some fields.
                    lappend records $rec
                    set rec {}
                }
                continue
            }
            set pos [string first = $line]
            if {$pos < 1} {
                error "Invalid field value format in wmic output line '$line'"
            }
            set wmicfield [string range $line 0 $pos-1]
            # wmic seems to html-encode values. Do minimal mapping
            set wmicvalue [string map {&amp; &} [string range $line $pos+1 end]]

            # The field name may differ in upper/lower case with what was 
            # passed in. If it matches, then use the name that was passed
            # in. Else use as is (e.g. if passed in was *)
            set pos [lsearch -ascii -nocase $fields $wmicfield]
            if {$pos >= 0} {
                lappend rec [lindex $fields $pos] $wmicvalue
            } else {
                lappend rec $wmicfield $wmicvalue
            }
        }
        
        return $records
    }

}

# Gets all fields of specified class with all field names in lower case.
proc wmic_records {wmiclass {clause ""}} {
    set recs {}
    foreach elem [wmic_get $wmiclass * $clause] {
        set rec {}
        foreach {fld val} $elem {
            lappend rec [string tolower $fld] $val
        }
        lappend recs $rec
    }
    return $recs
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

# Returns value of a field in all matching records
proc wmic_values {obj field {clause ""}} {
    # Some WMI fields have to be retrieved using *, not the name. For example
    # the FullName field from Win32_Account. Dunno why

    set values {}
    foreach rec [wmic_get $obj [list $field] $clause] {
        lappend values [dict get $rec $field]
    }
    return $values
}

# Return true if $a is close to $b (within 10%)
proc approx {a b {adjust 0}} {
    set max [expr {$a > $b ? $a : $b}]; # Tcl 8.4 does not have a max() func
    if {[expr {abs($b-$a) < ($max/10)}]} {
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
    return [expr {abs($b-$a) < ($max/10)}]
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
# the specified fields and list is not empty
# Raises an error otherwise
proc verify_list_kl_fields {l fields {ignoreextra 0}} {
    if {[llength $l] == 0} {
        error "List is empty."
    }
    foreach kl $l {
        verify_kl_fields $kl $fields $ignoreextra
    }
}

#
# Verify record array contains specified fields and that the 
# record array is not empty and each element has correct number of
# field values
proc verify_recordarray {ra fields {ignoreextra 0}} {
    set ra_fields [twapi::recordarray fields $ra]
    set ra_list [twapi::recordarray getlist $ra]
    if {[llength $ra_list] == 0} {
        error "Record array is empty"
    }
    foreach field $fields {
        if {$field ni $ra_fields} {
            error "Field <$field> not in record array fields <$ra_fields>"
        }
    }
    if {! $ignoreextra} {
        foreach field $ra_fields {
            if {$field ni $fields} {
                error "Record array has extra field <$field>"
            }
        }
    }

    set nfields [llength $ra_fields]
    foreach e $ra_list {
        if {[llength $e] != $nfields} {
            error "Record array record size [llength $e] != record array width $nfields"
        }
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
interp alias {} valid_privs {} verify_priv_list

# Verify evey element satisfies a condition
proc verify_list_elements {l cond} {
    foreach elem $l {
        if {! [eval $cond [list $elem]]} {
            return 0
        }
    }
    return 1
}

# Verify min list count and every list element matches a regexp
proc verify_list_match_regexp {cond l} {
    foreach {r min max} $cond break
    if {$min ne "" &&
        [llength $l] < $min} {
        return 0
    }
    if {$max ne "" &&
        [llength $l] > $max} {
        return 0
    }
    return [verify_list_elements $l [list regexp $r]]
}
tcltest::customMatch listregexp verify_list_match_regexp


# Verify two lists are equal
proc equal_lists {l1 l2} {
    if {[llength $l1] != [llength $l2]} {
        return 0
    }
    foreach e1 $l1 e2 $l2 {
        if {$e1 ne $e2} {
            return 0
        }
    }
    return 1
}
tcltest::customMatch list equal_lists

# Verify two lists are equal
proc equal_dicts {l1 l2} {
    if {[dict size $l1] != [dict size $l2]} {
        return 0
    }
    dict for {k val} $l1 {
        if {![dict exists $l2 $k] || $val != [dict get $l2 $k]} {
            return 0
        }
        dict unset l2 $k
    }
    return [expr {[dict size $l2] == 0}]
}
tcltest::customMatch dict equal_dicts

# Prompt the user
proc yesno {question {default "Y"}} {
    set answer ""
    # Make sure we are seen but catch because ui and console
    # packages may not be available
    catch {twapi::set_foreground_window [twapi::get_console_window]}
    while {![string is boolean -strict $answer]} {
        # We would have liked to use -nonewline here but that
        # causes output not to be displayed when running from 
        # tcltest::runAllTests. I believe this is the same 
        # bug Tcl seems to have with pipes and new lines
        # flushing on Windows (SF buf whatever)
        puts stdout "$question Type Y/N followed by Enter \[$default\] : "
        flush stdout
        set answer [string trim [gets stdin]]
#        puts $answer
        if {$answer eq ""} {
            set answer $default
        }
    }
    return [expr {!! $answer}]
}


# Pause to allow reader to read a message
proc pause {message} {
    # Make sure we are seen
    if {[info commands twapi::set_foreground_window] ne "" &&
        [info commands twapi::get_console_window] ne ""} {
        twapi::set_foreground_window [twapi::get_console_window]
    }
    # Would like -nonewline here but see comments in proc yesno
    puts "\n$message Hit Return to continue..."
    flush stdout
    gets stdin
    return
}

proc patience {task} {
    puts "$task may take a little while, patience please..."
}

proc hexdump {data {width 1} {count -1}} {
    # Adapted from AMG at http://wiki.tcl.tk/1599
    switch -exact -- $width {
        1 {
            set regex "(..)"
            set repl {\1 }
        }
        2 {
            set regex "(..)(..)"
            set repl {\2\1 }
        }
        4 {
            set regex "(..)(..)(..)(..)"
            set repl {\4\3\2\1 }
        }
    }
    set regex [string repeat (..) $width]
    set repl "[string range {\4\3\2\1} end-[expr {1+($width * 2)}] end] "
    if {$count < 1} {
        set count [string length $data]
    }
    for {set i 0} {$i < $count} {incr i 16} {
        set row [string range $data $i [expr {$i + 15}]]
        binary scan $row H* hex
        set hex [regsub -all $regex [format %-32s $hex] $repl]
        set row [regsub -all {[^[:print:]]} $row .]
        puts [format "%08x: %s %-16s" $i $hex $row]
    }
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
    # testlog "Starting slave"
    set fd [open "| [list [::tcltest::interpreter]]" r+]
    fconfigure $fd -buffering line -blocking 0 -eofchar {}
    if {[catch {tclsh_slave_verify_started $fd} msg]} {
        error $msg
    }
    #puts $fd [list source [file join $::twapi_test_script_dir testutil.tcl]]
    #puts $fd start_commandline
    return $fd
}
proc tclsh_slave_verify_started {fd} {
    # Verify started. Note we need the puts because tclsh does
    # not output result unless it is a tty.
    tclsh_slave_puts $fd {
        source testutil.tcl
        # testlog "SLAVE process [pid] STARTED"
        if {[catch {
            load_twapi_package
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
        testlog "Got version $ver"
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
proc gets_timeout {fd {ms 10000}} {
    set elapsed 0
    while {$elapsed < $ms} {
        if {[gets $fd line] == -1} {
            if {[eof $fd]} {
                testlog "get_timeout: unexpected eof"
                error "Unexpected EOF reading from $fd."
            }
            after 50;           # Wait a bit and then retry
            update;             # Required for non-blocking I/O
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
    set marker "Ready:[clock clicks]"
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

# Return expected path to the wish shell (assumed called from tclsh)
proc wish_path {} {
    set path [info nameofexecutable]
    set fn [file tail $path]
    if {![regsub tclsh $fn wish fn]} {
        if {![regsub tclkit-cli $fn tclkit-gui fn]} {
            error "Could not locate wish."
        }
    }
    set path [file join [file dirname $path] $fn]
    if {![file exists $path]} {
        error "Could not locate wish."
    }
    return $path
}

# Intended to be called as a separate wish process else allocate_console
# will fail.
proc allocate_console_in_wish {{title "wish dos console"}} {
    twapi::allocate_console
    twapi::set_console_title $title
}

# Intended to be called as a separate wish process else allocate_console
# will fail.
proc free_console_in_wish {{title "wish dos console"}} {
    twapi::allocate_console
    twapi::set_console_title $title
    # For some inexplicable reason, the patent tclsh does not return
    # from exec until the event loop is run here so do "update"
    update
    after 3000
    twapi::free_console
}


# Used for matching results
proc oneof {allowed_values value} {
    return [expr {[lsearch -exact $allowed_values $value] >= 0}]
}
tcltest::customMatch oneof oneof

# Result is a superset
proc superset {subset result} {
    foreach elem $subset {
        if {[lsearch -exact $result $elem] < 0} {
            return 0
        }
    }
    return 1
}
tcltest::customMatch superset superset

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


# Used for synch testing. Lock the mutex and signal the event_out and
# then wait for event_in
proc lock_mutex_and_signal_event {mutex event_out event_in} {
    set hmutex [twapi::open_mutex $mutex]
    set hev_in [twapi::create_event -name $event_in]
    set hev_out [twapi::create_event -name $event_out]
    twapi::lock_mutex $hmutex
    twapi::set_event $hev_out
    twapi::wait_on_handle $hev_in
    twapi::unlock_mutex $hmutex
    twapi::close_handle $hev_in
    twapi::close_handle $hev_out
    twapi::close_handle $hmutex

}

proc attempt_lock_mutex {mutex open_call} {
    if {[catch {
        if {$open_call eq "open_mutex"} {
            set hmutex [twapi::open_mutex $mutex]
        } else {
            set hmutex [twapi::create_mutex -name $mutex]
        }
    }]} {
        puts "error $::errorCode"
        return
    }
    set result [twapi::lock_mutex $hmutex -wait 0]
    puts $result
    if {$result eq "signalled"} {
        twapi::unlock_mutex $hmutex
    }
    twapi::close_handle $hmutex
}


####
# Certificate stuff

# Returns the sample store context handle. Must not be released by caller
# Note this contains only certs, NO private keys
proc samplestore {} {
    global samplestore
    if {![info exists samplestore]} {
        # TBD - may be just read in the PFX file instead ?
        set samplestore [twapi::cert_temporary_store]
        foreach suff {ca intermediate server client altserver full min} {
            twapi::cert_release [twapi::cert_store_add_encoded_certificate $samplestore [sampleencodedcert $suff]]
        }
    }
    return $samplestore
}

proc storewithkeys {} {
    global storewithkeys
    if {![info exists storewithkeys]} {
        set storewithkeys [twapi::make_test_certs]
    }
    return $storewithkeys
}

# Returns revoked cert. Must be released by caller
proc revokedcert {} {
    set enc [read_file [file join [tcltest::testsDirectory] certs grcrevoked.cer] rb]
    return [twapi::cert_import $enc]
}

# Returns google cert. Must be released by caller
proc googlecert {} {
    set enc [read_file [file join [tcltest::testsDirectory] certs www.google.com.pem] r]
    return [twapi::cert_import $enc -encoding pem]
}


# Returns google encoded cert.
proc googleencodedcert {} {
    return [read_file [file join [tcltest::testsDirectory] certs www.google.com.pem] r]
}

proc expiredcert {} {
    set enc [read_file [file join [tcltest::testsDirectory] certs www.google.com-expired.pem] r]
    return [twapi::cert_import $enc -encoding pem]
}

# Returns context for one of the sample certs in the sample store
# Must be released by caller
proc samplecert {{which full}} {
    return [twapi::cert_store_find_certificate [samplestore] subject_substring twapitest$which]
}

# Returns one of the sample certs in encoded form
proc sampleencodedcert {{which full}} {
    variable _sampleencodedcert
    if {![info exists _sampleencodedcert($which)]} {
        set _sampleencodedcert($which) [read_file [file join [tcltest::testsDirectory] certs twapitest${which}.cer] rb]
    }
    return $_sampleencodedcert($which)
}

proc temp_system_store_path {} {
    variable temp_system_store_path

    if {![info exists temp_system_store_path]} {
        set temp_system_store_path TwapiTest-[clock microseconds]
        set hstore [twapi::cert_system_store_open $temp_system_store_path user]
        set cer [sampleencodedcert]
        twapi::cert_release [twapi::cert_store_add_encoded_certificate $hstore $cer]
        twapi::cert_store_release $hstore
    }
    return $temp_system_store_path
}

proc temp_crypto_dir_path {} {
    variable temp_crypto_dir_path
    if {![info exists temp_crypto_dir_path]} {
        set temp_crypto_dir_path [tcltest::makeDirectory twapicryptotest]
    }
    return $temp_crypto_dir_path
}


proc temp_file_store_path {} {
    variable temp_file_store_path

    if {![info exists temp_file_store_path]} {
        set temp_file_store_path [file join [temp_crypto_dir_path] [clock microseconds].store]
        set hstore [twapi::cert_file_store_open $temp_file_store_path]
        set cer [sampleencodedcert]
        twapi::cert_release [twapi::cert_store_add_encoded_certificate $hstore $cer]
        twapi::cert_store_release $hstore
    }
    return $temp_file_store_path
}

proc cleanup_test_cert_files {} {
    variable temp_file_store_path
    variable temp_system_store_path

    if {[info exists temp_file_store_path] &&
        [file exists $temp_file_store_path]} {
        file delete $temp_file_store_path
    }

    foreach store [twapi::cert_system_stores user] {
        if {[string match -nocase TwapiTest-* $store]} {
            twapi::cert_system_store_delete $store user
        }
    }

    twapi::crypt_test_container_cleanup
}

proc equal_certs {certa certb} {
    return [expr {[twapi::cert_thumbprint $certa] eq [twapi::cert_thumbprint $certb]}]
}

# Returns cert from a store at index n for testing searches
# Errors if not enough certs
proc pick_cert {hstore {n 4}} {
    set hcert NULL
    time {set hcert [twapi::cert_store_find_certificate $hstore any "" $hcert]} [incr n]
    return $hcert
}

proc openssl_path {args} {
    set path [file join [pwd] .. openssl bin openssl.exe]
    if {![file exists $path]} {
        # Try from the source pool. We do it this way because 
        # of problems loading in VmWare virtual machines from the
        # host's drive share
        set path [list ../../tools/openssl/bin/openssl.exe]
        if {![file exists $path]} {
            error "Could not locate openssl.exe"
        }
    }
    return [file normalize [file join [file dirname [file dirname $path]] {*}$args]]
}

# WARNING: return value combines stdout and stderr so the two
# can be intermixed. Take into consideration when matching
proc openssl {args} {
    set cmd [openssl_path bin openssl.exe]
    set ::env(OPENSSL_CONF) [openssl_path ssl openssl.cnf]
    set stderr_temp [file join [temp_crypto_dir_path] twapi-openssl-stderr.tmp]
    set status 0
    if {[catch {exec $cmd {*}$args 2> $stderr_temp} stdout options]} {
        if {[lindex [dict get $options -errorcode] 0] eq "CHILDSTATUS"} {
            error [read_binary $stderr_temp]
        } else {
            return -options $options -level 0 $stdout
        }
    }
    return [list $stdout [read_binary $stderr_temp]]
}

# WARNING: return value combines stdout and stderr so the two
# can be intermixed. Take into consideration when matching data read
# from the returned channel
proc openssl& {args} {
    set cmd [openssl_path bin openssl.exe]
    set ::env(OPENSSL_CONF) [openssl_path ssl openssl.cnf]
    set stderr_temp [file join [temp_crypto_dir_path] twapi-openssl-stderr.tmp]
    set status 0
    return [open |[list $cmd {*}$args 2>@1]]
}

proc openssl_port {} {
    return 4433
}

proc openssl_ca_store {} {
    set store [twapi::cert_temporary_store]
    foreach ca {root-ca signing-ca} {
        set cert [twapi::cert_import [read_file [openssl_path ca ${ca}.crt]] -encoding pem]
        twapi::cert_release [twapi::cert_store_add_certificate $store $cert]
        twapi::cert_release $cert
    }
    proc openssl_ca_store {} [list return $store]
    return $store
}

proc openssl_ca_cert {} {
    set cert [twapi::cert_import [read_file [openssl_path ca root-ca.crt]] -encoding pem]
    proc openssl_ca_cert {} [list return $cert]
    return $cert
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

proc setops::equal {A B} {
    return [empty [symdiff $A $B]]
}
tcltest::customMatch set setops::equal

# If this is the first argument to the shell and there are more arguments
# execute them

if {[string equal -nocase [file normalize $argv0] [file normalize [info script]]]} {
    load_twapi_package
    if {[catch {
        foreach arg $argv {
            eval $arg
        }
    } msg]} {
        twapi::eventlog_log "testutil error: $msg\n$::errorInfo"
    }
} else {
    # Only define these if running inside a test script else
    # wish errors out
    tcltest::testConstraint domain [indomain]
    tcltest::testConstraint dcexists [expr {[testconfig domain_controller] ne ""}]
}
