# TLS stress test routines - does NOT test certs
# The code is a bit of a mess and will take some work to operate outside
# of the twapi test scaffolding.

# Hide argv otherwise the test package swallows the arguments
set __argv $argv
set argv [list ]
if {[llength [info commands load_twapi_package]] == 0} {
    source [file join [file dirname [info script]] testutil.tcl]
}
set argv $__argv

if {[llength [info commands ::twapi::tls_socket]] == 0} {
    load_twapi_package twapi_crypto
}


proc tls_echo_usage {} {
    puts stderr {
Usage:
     tclsh tlsecho.tcl syncserver ?PORT? ?TIMEOUT?
        -- runs the synchronous echo server
     tclsh tlsecho.tcl asyncserver ?PORT? ?TIMEOUT?
        -- runs the asynchronous echo server
Does NOT verify certs - only for testing throughput and message sizes.
    }
    exit 1
}

proc tls_verifier args {
    # Don't care about the cert. Only used for testing throughput
    return 1
}

proc tls_init_echo_server_creds {} {
    # Where the certificates are
    set pfxFile [file join [file dirname [info script]] certs twapitest.pfx]
    # Set up the store containing the certificates
    set certStore [twapi::cert_temporary_store -pfx [read_file $pfxFile rb]]
    # Set up the client and server credentials
    set serverCert [twapi::cert_store_find_certificate $certStore subject_substring twapitestserver]
    # TBD - check if certs can be released as soon as we obtain credentials
    set ::tls_server_creds [twapi::sspi_acquire_credentials -credentials [twapi::sspi_schannel_credentials -certificates [list $serverCert]] -package unisp -role server]
    twapi::cert_release $serverCert
    twapi::cert_store_release $certStore
}

proc tls_echo_syncserver_accept {echo_fd clientaddr clientport} {
    after cancel $::timer
    fconfigure $echo_fd -buffering line -translation crlf -eofchar {} -encoding utf-8
    set msgs 0
    set last_size 0
    set total 0
    while {1} {
        if {[gets $echo_fd line] >= 0} {
            if {$line eq "exit"} {
                break
            }
            puts $echo_fd $line
            incr msgs
            set last_size [string length $line]
            incr total $last_size
        } else {
            puts stderr "Unexpected eof from echo client"
            break
        }
    }
    close $echo_fd
    set ::tls_echo_server_status [list $msgs $total $last_size]
}

proc tls_echo_server {type {port 4433} {timeout 20000}} {
    set ::timer [after $timeout "set ::tls_echo_server_status timeout"]
    set listen_fd [::twapi::tls_socket -credentials $::tls_server_creds -server tls_echo_${type}_accept 4433]
    # Following line is important as it is used by automated test scripts
    puts "READY"
    vwait ::tls_echo_server_status
    return $::tls_echo_server_status
}

proc tls_echo_server_async_echoline {chan} {

    # Check end of file or abnormal connection drop,
    # then echo data back to the client.

    set error [catch {gets $chan line} count]
    if {$error || $count < 0} {
        if {$error || [eof $chan]} {
            set ::tls_echo_server_status eof
            close $chan
        }
    } else {
        if {$line eq "exit"} {
            set ::tls_echo_server_status [list $::tls_msgs $::tls_total $::tls_last_size]
        } else {
            incr ::tls_msgs
            set ::tls_last_size [string length $line]
            incr ::tls_total $::tls_last_size
            puts $chan $line
        }
    }
}

proc tls_echo_asyncserver_accept {echo_fd clientaddr clientport} {
    after cancel $::timer

    set ::tls_msgs 0
    set ::tls_last_size 0
    set ::tls_total 0

    fconfigure $echo_fd -buffering line -translation crlf -eofchar {} -encoding utf-8 -blocking 0
    fileevent $echo_fd readable [list ::tls_echo_server_async_echoline $echo_fd]
    vwait ::tls_echo_server_status
    return $::tls_echo_server_status
}

proc tls_echo_client {args} {
    array set opts [twapi::parseargs args {
        {port.int 4433}
        {density.int 1}
        {limit.int 10000}
    }]

    set alphabet "0123456789abcdefghijklmnopqrstuvwxyz"
    set alphalen [string length $alphabet]
    set msgs 0
    set last 0
    set total 0
    set fd [twapi::tls_socket -verifier tls_verifier 127.0.0.1 $opts(port)]
    fconfigure $fd -buffering line -translation crlf -eofchar {} -encoding utf-8
    for {set i 1} {$i < $opts(limit)} {incr i [expr {1+($i+1)/$opts(density)}]} {
        set c [string index $alphabet [expr {$i % $alphalen}]]
        set request [string repeat $c $i]
        puts $fd $request
        set response [gets $fd]
        if {$request ne $response} {
            if {$response eq "" && [eof $fd]} {
                break
            }
            puts "Mismatch in message of size $i"
            set n [string length $response]
            if {$i != $n} {
                puts "Sent $i chars, received $n chars"
            }
        }
        incr msgs
        incr total $i
        set last $i
    }
    puts $fd "exit"
    close $fd
    return [list $msgs $total $last]
}


# Main code
if {[string equal -nocase [file normalize $argv0] [file normalize [info script]]]} {
    # We are being directly sourced, not as a library
    if {[llength $argv] == 0} {
        tls_echo_usage
    }

    tls_init_echo_server_creds
    switch -exact -- [lindex $argv 0] {
        syncserver -
        asyncserver {
            if {[catch {
                foreach {nmsgs nbytes last} [eval tls_echo_server [lindex $argv 0] [lrange $argv 1 end]] break
            }]} {
                testlog $::errorInfo
            }
        }
        default {
            tls_echo_usage
        }
    }

    twapi::sspi_free_credentials $::tls_server_creds
    puts [list $nmsgs $nbytes $last]
}

