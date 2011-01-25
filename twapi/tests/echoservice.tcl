# A sample Windows service implemented in Tcl using TWAPI's windows
# services module.
#
# A sample client session looks like this
#   set s [echo_client localhost 2020]
#   puts $s "Hello!"
#   gets $s line

proc usage {} {
    puts stderr {
Usage:
     tclsh echoservice.tcl install SERVICENAME ?PORT?
       -- installs the service as SERVICENAME listening on PORT
     tclsh echoservice.tcl uninstall SERVICENAME
       -- uninstalls the service
Then start/stop the service using either "net start" or the services control
manager GUI.
    }
    exit 1
}

# If in source dir, we load that twapi in preference to installed package
set echo_script_dir [file dirname [file dirname [file normalize [info script]]]]
if {[file exists [file join $echo_script_dir tcl twapi.tcl]]} {
    uplevel #0 [list source [file join $echo_script_dir tcl twapi.tcl]]
} else {
    uplevel #0 package require twapi
}


################################################################
# The echo_server code is almost verbatim from the Tcl Developers
# Exchange samples.

set echo(server_state) stopped;       # State of the server

# echo_server --
#       Open the server listening socket
#       and enter the Tcl event loop
#
# Arguments:
#       port    The server's port number

proc echo_server {} {
    global echo
    if {[catch {
        set echo(server_socket) [socket -server echo_accept $echo(server_port)]
    } msg]} {
        ::twapi::eventlog_log "Error binding to port $echo(server_port): $msg"
        exit 1
    }
}

# echo_accept --
#       Accept a connection from a new client.
#       This is called after a new socket connection
#       has been created by Tcl.
#
# Arguments:
#       sock    The new socket connection to the client
#       addr    The client's IP address
#       port    The client's port number

proc echo_accept {sock addr port} {
    global echo

    if {$echo(server_state) ne "running"} {
        close $sock
        return
    }

    # Record the client's information

    set echo(addr,$sock) [list $addr $port $sock]

    # Ensure that each "puts" by the server
    # results in a network transmission

    fconfigure $sock -buffering line

    # Set up a callback for when the client sends data

    fileevent $sock readable [list echo $sock]
}

# echo --
#       This procedure is called when the server
#       can read data from the client
#
# Arguments:
#       sock    The socket connection to the client

proc echo {sock} {
    global echo

    # Check end of file or abnormal connection drop,
    # then echo data back to the client.

    if {[eof $sock] || [catch {gets $sock line}]} {
        close $sock
        unset -nocomplain echo(addr,$sock)
    } else {
        puts $sock $line
    }
}

#
# Close all sockets
proc echo_close_shop {{serveralso true}} {
    global echo

    # Loop and close all client connections
    foreach {index conn} [array get echo addr,*] {
        close [lindex $conn 2]; # 3rd element is socket handle
        unset -nocomplain echo($index)
    }

    if {$serveralso} {
        close $echo(server_socket)
        unset -nocomplain echo(server_socket)
    }
}

#
# Send the message to all connected clients
proc echo_bcast {msg} {
    global echo

    # Loop and close all client connections
    foreach {index conn} [array get echo addr,*] {
        catch {puts [lindex $conn 2] $msg}; # 3rd element is socket handle
    }
}




################################################################
# The actual service related code

#
# Update the SCM with our state
proc report_state {name seq} {
    if {[catch {
        set ret [twapi::update_service_status $name $seq $::echo(server_state)]
    } msg]} {
        ::twapi::eventlog_log "Service $name failed to update status: $msg"
    }
}

# Callback handler
proc service_control_handler {control {name ""} {seq 0} args} {
    global echo
    switch -exact -- $control {
        start {
            if {[catch {
                # Start the echo server
                echo_server
                set echo(server_state) running
            } msg]} {
                twapi::eventlog_log "Could not start echo server: $msg"
            }
            report_state $name $seq
        }
        stop {
            echo_close_shop
            set echo(server_state) stopped
            report_state $name $seq
        }
        pause {
            # Close all client connections but leave server socket open
            echo_close_shop false
            set echo(server_state) paused
            report_state $name $seq
        }
        continue {
            set echo(server_state) running
            report_state $name $seq
        }
        userdefined {
            # Note we do not need to call update_service_status
            echo_bcast "CONTROL: $control,$name,$seq,[join $args ,]"
        }
        all_stopped {
            # Mark we are all done so we can exit at global level
            set ::done 1
        }
        default {
            # Ignore
            echo_bcast "CONTROL: $control,$name,$seq,[join $args ,]"
        }
    }
}


################################################################
# Main code

# Parse arguments
if {[llength $argv] < 2 || [llength $argv] > 3} {
    usage
}

set service_name [lindex $argv 1]
if {[llength $argv] > 2} {
    set echo(server_port) [lindex $argv 2]
} else {
    set echo(server_port) 2020
}

switch -exact -- [lindex $argv 0] {
    service {
        # We are running as a service
        if {[catch {
            set controls {stop pause_continue paramchange shutdown powerevent}
            if {[twapi::min_os_version 5 1]} {
                lappend controls sessionchange
            }
            twapi::run_as_service [list [list $service_name ::service_control_handler]] -controls $controls
        } msg]} {
            twapi::eventlog_log "Service error: $msg\n$::errorInfo"
        }
        # We sit in the event loop until service control stop us through
        # the event handler
        vwait ::done
    }
    install {
        if {[twapi::service_exists $service_name]} {
            puts stderr "Service $service_name already exists"
            exit 1
        }

        # Make the names a short name to not have to deal with
        # quoting of spaces in the path

        set exe [file nativename [file attributes [info nameofexecutable] -shortname]]
        set script [file nativename [file attributes [file normalize [info script]] -shortname]]
        twapi::create_service $service_name "$exe $script service $service_name $echo(server_port)"
    }
    uninstall {
        if {[twapi::service_exists $service_name]} {
            twapi::delete_service $service_name
        }
    }
    default {
        usage
    }
}

exit 0
