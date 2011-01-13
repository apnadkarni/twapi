# Report which commands do not have a test suite

#package require twapi

package require tcltest

::tcltest::testConstraint systemmodificationok 1
::tcltest::testConstraint userInteraction 1
::tcltest::testConstraint powertests 1

set implemented 0
set placeholders {}
set missing {}

proc tcltest::test args {
    set test_id [lindex $args 0]
    lappend ::tests $test_id

    # If a constraint is TBD, mark as a placeholder
    set i [lsearch -exact $args "-constraints"]
    if {$i >= 0} {
        if {[lsearch -exact [lindex $args [incr i]] "TBD"] >= 0} {
            lappend ::placeholders "$test_id - [string trim [lindex $args 1]]"
        } else {
            incr ::implemented
        }
    } else {
        incr ::implemented
    }
}
proc tcltest::cleanupTests {} {}

foreach fn [glob *.test] {
    source $fn
}



# What we are skipping and why
array set do_not_test {
    load_twapi                fundamental
    class "Covered by COM"
    eventlog_monitor_start eventlog_monitor
    eventlog_monitor_stop  eventlog_monitor
    large_system_time_to_secs deprecated_alias
    getaddrinfo               hostname_to_address
    getnameinfo               address_to_hostname
    get_drive_info            deprecated_alias
    get_logical_drives        deprecated_alias
    control_service           "stop_service and friends"
    get_lcid                  deprecated_alias
    get_process_handle        get_process_info
    get_connected_shares      deprecated_alias
    kl_get_default            deprecated_alias
    set_console_output_mode   deprecated_alias
    set_console_input_mode   deprecated_alias
    read_console             deprecated_alias
    write_console            deprecated_alias
}
set missing {}
foreach cmd [twapi::_get_public_procs] {
    # Skip internal commands
    if {[regexp {^(_|[A-Z]).*} $cmd]} continue

    # Skip commands that are tested in tests whose test names do not match
    if {[info exists do_not_test($cmd)]} continue

    if {[lsearch -glob $::tests ${cmd}-*] < 0} {
        lappend missing $cmd
    }
}

puts "Implemented tests: $implemented"
puts "\nPlaceholders: [llength $placeholders]"
puts [join $placeholders \n];   # Don't sort - want in file order
puts "\nMissing tests:[llength $::missing]"
puts [join [lsort $missing] \n]


