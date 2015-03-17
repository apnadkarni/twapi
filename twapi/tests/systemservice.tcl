# A Windows service implemented in Tcl using TWAPI's windows
# services module.
# Runs as a system service

lappend auto_path c:/src/twapi/twapi/dist
package require twapi

namespace eval systcl {
    variable Script [info script]
    variable Name SysTcl
    variable State stopped
}

proc systcl::usage {} {
    variable Script
    puts stderr [format {
Usage:
    %1$s %2$s install
       -- installs the service
    %1$s %2$s uninstall
       -- uninstalls the service
Then start/stop the service using either "net start" or the services control
manager GUI.
    } [file tail [info nameofexecutable]] $Script]
    exit 1
}

proc systcl::run_child_orig {} {
    # Get our access token and duplicate it so we can point to the
    # console window station.
    set my_tok [twapi::open_process_token -access {token_query token_duplicate token_assign_primary}]
    twapi::trap {
        set tok [twapi::duplicate_token $my_tok -access {token_query token_duplicate token_assign_primary token_adjust_sessionid token_adjust_default} -type primary]
        twapi::set_token_tssession $tok 1
        twapi::create_process c:/windows/system32/cmd.exe -startdir c:\\ -token $tok
        # twapi::create_process c:/tcl/864/x64/bin/wish86t.exe -token $tok
    } finally {
        twapi::close_token $my_tok
        if {[info exists tok]} {
            twapi::close_token $tok
        }
    }

    
}

proc systcl::steal_identity {tok} {
    twapi::trap {
        return [twapi::duplicate_token $tok -access {
            token_query token_duplicate token_assign_primary
            token_adjust_sessionid token_adjust_default
        } -type primary]
    } finally {
        twapi::close_token $tok
    }
}

proc systcl::find_victim {name} {
    set tssession [expr {[twapi::min_os_version 6] ? 1 : 0}]

    # Find the victim whose identity we want to  steal
    foreach pid [twapi::get_process_ids -name $name] {
        if {[dict get [twapi::get_process_info $pid -tssession] -tssession] == $tssession} {
            set tok [twapi::open_process_token -pid $pid -access {
                token_query token_duplicate token_assign_primary
            }]
            return [steal_identity $tok]
        }
    }
    error "Could not find a victim process $name in session $tssession"
}

proc systcl::shell_path {} {
    return [file nativename [file join $::env(SYSTEMROOT) system32 cmd.exe]]
}

proc systcl::run_child {tok environ} {
    twapi::create_process [shell_path] -startdir c:\\ -token $tok -env $environ
}

proc systcl::system_shell {} {
    # Find the winlogon process
    
    # Get winlogon access token and duplicate it so we can point to the
    # console window station.
    set tok [systcl::steal_identity winlogon.exe]
    twapi::trap {
        run_child $tok [twapi::get_system_environment_vars]
    } finally {
        twapi::close_token $tok
    }
}

proc systcl::user_shell {who} {
    set tssession [expr {[twapi::min_os_version 6] ? 1 : 0}]
    set tok [steal_identity [twapi::WTSQueryUserToken $tssession]]
    twapi::trap {
        twapi::load_user_profile $tok; # Not really necessary since we are stealing from existing process
        twapi::impersonate_token $tok
        twapi::trap {
            run_child $tok [twapi::get_user_environment_vars $tok]
        } finally {
            twapi::revert_to_self
        }
    } finally {
        twapi::close_token $tok
    }
}

proc systcl::stop_children {} {
    # TBD
}

#
# Update the SCM with our state
proc systcl::report_state {seq} {
    variable State
    variable Name
    if {[catch {
        set ret [twapi::update_service_status $Name $seq $State]
    } msg]} {
        ::twapi::eventlog_log "Service $Name failed to update status: $msg\r\n$::errorInfo"
    }
}

# Callback handler
proc systcl::service_control_handler {control {name ""} {seq 0} args} {
    variable State
    variable Done
    switch -exact -- $control {
        start {
            if {[catch {
                set State running
                report_state $seq
            } msg]} {
                twapi::eventlog_log "Could not start $name server: $msg\r\n$::errorInfo"
                set Done 1
            }
        }
        stop {
            stop_children
            set State stopped
            report_state $seq
            set Done 1
        }
        userdefined {
            set arg 128
            if {[llength $args]} {
                set arg [lindex $args 0]
            }
            if {[catch {
                if {$arg == 128} {
                    system_shell
                } else {
                    user_shell
                }
            } msg]} {
                twapi::eventlog_log "Error starting child: $msg\r\n$::errorInfo"
            }
        }
        default {
            # Ignore
        }
    }
}


################################################################
# Main code

# Parse arguments
if {[llength $argv] > 1} {
    systcl::usage
}

switch -exact -- [lindex $argv 0] {
    service {
        # We are running as a service
        if {[catch {
            twapi::run_as_service [list [list $systcl::Name ::systcl::service_control_handler]] -controls [list stop]
        } msg]} {
            twapi::eventlog_log "Service error: $msg\n$::errorInfo"
        }
        # We sit in the event loop until service control stop us through
        # the event handler
        vwait ::systcl::Done
    }
    install {
        if {[twapi::service_exists $::systcl::Name]} {
            puts stderr "Service $::systcl::Name already exists"
            exit 1
        }

        # Make the names a short name to not have to deal with
        # quoting of spaces in the path

        set exe [file nativename [file attributes [info nameofexecutable] -shortname]]
        set script [file nativename [file attributes [file normalize [info script]] -shortname]]
        twapi::create_service $::systcl::Name "$exe $script service" -interactive 1
    }
    uninstall {
        if {[twapi::service_exists $::systcl::Name]} {
            twapi::delete_service $::systcl::Name
        }
    }
    default {
        systcl::usage
    }
}

exit 0
