namespace eval twapi::ssl {
    # Each element of _channels is dictionary with the following keys
    #  Socket - the underlying socket
    #  State - SERVERINIT, CLIENTINIT, LISTENERINIT, OPEN, NEGOTIATING, CLOSED
    #  Type - SERVER, CLIENT, LISTENER
    #  Blocking - 0/1 indicating whether blocking or non-blocking channel
    #  WatchMask - list of {read write} indicating what events to post
    #  Credentials - credentials handle to use for local end of connection
    #  AcceptCallback - application callback on a listener socket
    #  SspiContext - SSPI context for the connection
    #  Input  - plaintext data to pass to app
    #  Output - plaintext data to encrypt and output

    variable _channels
    array set _channels {}

    namespace path [linsert [namespace path] 0 [namespace parent]]
}

interp alias {} twapi::ssl_socket {} twapi::ssl::_socket
proc twapi::ssl::_socket {args} {

    parseargs args {
        myaddr.arg
        myport.int
        async
        {server.arg {}}
        {credentials.arg {}}
        {verifier.arg {}}
    } -setvars

    set chan [chan create {read write} [list [namespace current]]]

    set socket_args {}
    foreach opt {myaddr myport} {
        if {[info exists $opt]} {
            lappend socket_args -$opt [set $opt]
        }
    }

    if {[info exists server]} {
        set type LISTENER
        lappend socket_args -server [list [namespace current]::_accept $chan]
    } else {
        lappend socket_args -async; # Always async, we will explicitly block
        set type CLIENT
    }

    trap {
        set so [socket {*}$socket_args {*}$args]
        _init $chan $type $so $credentials $verifier $server
        if {$type eq "CLIENT"} {
            if {! $async} {
                _client_blocking_negotiate $chan
            }
        }
    } onerror {} {
        catch {_cleanup $chan}
        rethrow
    }

    return $chan
}


proc twapi::ssl::_accept {listener so raddr raport} {
    variable _channels

    trap {
        set chan [chan create {read write} [list [namespace current]::ssl $id]]
        _init $chan SERVER $so [dict get $_channels($listener) Credentials] [dict get $_channels($listener) Verifier]
        {*}[dict get _channels($listener) AcceptCallback] $chan $raddr $raport
    } onerror {
        catch {_cleanup $chan}
        rethrow
    }
    return
}

proc twapi::ssl::initialize {chan mode} {
    # All init is done in chan creation routine after base socket is created
    return {initialize finalize watch blocking read write configure cget cgetall}
}

proc twapi::ssl::finalize {chan} {
    _cleanup $chan
    return
}

proc twapi::ssl::blocking {chan mode} {
    variable _channels

    chan configure [_chansocket $chan] -blocking $mode
    dict set _channels($chan) Blocking $mode
    if {$mode == 0} {
        # Since we need to negotiate SSL we always have socket event
        # handlers irrespective of the state of the watch mask
        chan event $Socket readable [list [namespace current]::_so_read_handler $chan]
        chan event $Socket writable [list [namespace current]::_so_write_handler $chan]
    } else {
        chan event $Socket readable {}
        chan event $Socket writable {}
    }
    return
}

proc twapi::ssl::watch {chan watchmask} {
    variable _channels
    dict with _channels($chan) {
        set WatchMask $watchmask
        if {"read" in $watchmask} {
            if {[string length $Input]} {
                chan postevent read
            }
        }

        if {"write" in $watchmask} {
            if {$State eq "OPEN"} {
                chan postevent write
            }
        }
    }

    return
}

proc twapi::ssl::read {chan nbytes} {
    variable _channels

    if {$nbytes == 0} {
        return {}
    }

    # This is not inside the dict with because _negotiate will update the dict
    if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
        _negotiate $chan
        if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            # If a blocking channel, should have come back with negotiation
            # complete. If non-blocking, return EAGAIN to indicate no
            # data yet
            if {[dict get $_channels($chan) Blocking]} {
                error "SSL negotiation failed on blocking channel" 
            } else {
                return -code error EAGAIN
            }
        }
    }

    dict with _channels($chan) {
        if {[string length $Input] >= $nbytes} {
            # We already have enough decrypted bytes
            # TBD - see if return [lindex [list [string range $Input 0 $nbytes-1] [set Input [string range $Input $nbytes end]]] 0]
            # is significantly faster (inline K operator)

            set ret [string range $Input 0 $nbytes-1]
            set Input [string range $Input $nbytes end]
            return $ret
        }

        if {$State eq "OPEN"} {
            if {$Blocking} {
                # The channel does not compress so we need to read in
                # at least $needed bytes. Because of SSL overhead, we may actually
                # need even more
                while {[set ninput [string length $Input]] < $nbytes} {
                    set needed [expr {$nbytes-$ninput}]
                    set data [chan read $Socket $needed]
                    if {[string length $data] < $needed} {
                        set ret $Input; # ok to return < nbytes on eof
                        set Input ""
                        set State CLOSED
                        return $ret; # May be "" indicating eof
                    }
                    append Input [sspi_decrypt_stream $SspiContext $data]
                }
                set ret [string range $Input 0 $nbytes-1]
                set Input [string range $Input $nbytes end]
                return $ret
            } else {
                # Non-blocking - read all that we can
                set data [chan read $Socket]
                if {[string length $data]} {
                    append Input [sspi_decrypt_stream $SspiContext $data]
                }
                if {[string length $Input]} {
                    # Return whatever we have (may be less than $nbytes)
                    set ret [string range $Input 0 $nbytes-1]
                    set Input [string range $Input $nbytes end]
                    return $ret
                }

                # Do not have enough data. See if connection closed
                if {[chan eof $Socket]} {
                    set State CLOSED
                    return "";          # EOF
                }
                # Not closed, just waiting for data
                return -code error EAGAIN
            }
        } elseif {$State eq "CLOSED"} {
            # Return whatever we have (less than $nbytes)
            return $Input[set Input ""]
        }
    }

}

proc twapi::ssl::write {chan data} {
    variable _channels
    dict with _channels($chan) {
        switch $State {
            CLOSED {
                # TBD - for now throw it away. Should we raise an error ?
            }
            OPEN {
                # There might be pending output if channel has just
                # transitioned to OPEN state
                if {[string length $Output]} {
                    chan puts -nonewline $Socket [sspi_encrypt_stream $SspiContext $Output]
                    set Output ""
                }
                chan puts -nonewline $Socket [sspi_encrypt_stream $SspiContext $data]
                flush $Socket
            }
            default {
                append Output $data
            }
        }
    }
    return [string length $data]
}

proc twapi::ssl::configure {chan opt val} {
    if {$opt eq "-credentials"} {
        error "Option -credentials is read-only."
    }

    chan configure [_chansocket $chan] $opt $val
    return
}

proc twapi::ssl::cget {chan opt} {
    variable _channels

    set so [_chansocket $chan]

    if {$opt eq "-credentials"} {
        return [dict get $_channels($chan) Credentials]
    }

    return [chan configure $so $opt]
}

proc twapi::ssl::cgetall {chan} {
    set so [_chansocket $chan]
    set config [chan configure $so]
    lappend config -credentials [dict get $_channels($chan) Credentials]
    return $config
}

proc twapi::ssl::_chansocket {chan} {
    variable _channels
    if {![info exists _channels($chan)]} {
        error "Channel $chan not found."
    }
    return [dict get $_channels($chan) Socket]
}

proc twapi::ssl::_init {chan type so creds verifier {accept_callback {}}} {
    variable _channels

    # TBD - verify that -buffering none is the right thing to do
    # as the scripted channel interface takes care of this itself
    # We always set -blocking to 0 and take care of blocking ourselves
    chan configure $so -translation binary -buffering none -blocking 0
    set _channels($chan) [list Socket $so \
                              State ${type}INIT \
                              Type $type \
                              Blocking [chan configure $so -blocking] \
                              WatchMask {} \
                              Verifier $Verifier \
                              Credentials $creds \
                              SspiContext {} \
                              Input {} Output {}]
    if {$type eq "LISTENER"} {
        dict set _channels($chan) AcceptCallback $accept_callback
    }
}

proc twapi::ssl::_cleanup {chan} {
    variable _channels
    if {[info exists _channels($chan)]} {
        # Note _cleanup can be called in inconsistent state so not all
        # keys may be set up
        dict with _channels($chan) {
            if {[info exists Socket]} {
                if {[catch {chan close $Socket} msg]} {
                    # TBD - debug log socket close error
                }
            }
            if {[info exists SspiContext]} {
                # TBD - call sspi_shutdown_context first ?
                if {[catch {sspi_delete_context $SspiContext} msg]} {
                    # TBD - debug log
                }
            }
        }
        unset _channels($chan)
    }
}

proc twapi::ssl::_so_read_handler {chan} {
    variable _channels

    if {[info exists _channels($chan)]} {
        dict with _channels($chan) {}
        if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            _negotiate $chan
        }

        if {"read" in [dict get $_channels($chan) WatchMask] && [string length [dict get $_channels($chan) Input]]} {
            chan postevent $chan read
        }
    }
    return
}

proc twapi::ssl::_so_write_handler {chan} {
    variable _channels
    if {[info exists _channels($chan)]} {
        if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            _negotiate $chan
        }

        if {"write" in [dict get $_channels($chan) WatchMask] && [dict get $_channels($chan) State] eq "OPEN"} {
            chan postevent $chan write
        }
    }
    return
}

proc twapi::ssl::_negotiate {chan} {
    variable _channels
        
    switch [dict get $_channels($chan) State] {
        NEGOTIATING {
            dict with _channels($chan) {
                if {$Blocking} {
                    error "Internal error: NEGOTIATING state not expected on a blocking ssl socket"
                }

                set data [chan read $Socket]
                if {[string length $data] == 0} {
                    if {[chan eof $Socket]} {
                        error "Unexpected EOF during SSL negotiation"
                    } else {
                        # No data yet, just keep waiting
                    }
                } else {
                    lassign [sspi_step $SspiContext $data] status output leftover
                    if {[string length $output]} {
                        chan puts -nonewline $Socket $output
                        chan flush $Socket
                    }
                    switch $status {
                        done {
                            set State OPEN
                            if {[string length $leftover]} {
                                set Input [sspi_decrypt_stream $SspiContext $leftover]
                            }
                        }
                        continue {
                            # Keep waiting for next input
                        }
                        default {
                            error "Unexpected status $status from sspi_step"
                        }
                    }
                }
            }
        }

        CLIENTINIT {
            if {[dict get $_channels($chan) Blocking]} {
                _client_blocking_negotiate $chan
            } else {
                set State NEGOTIATING
                set SspiContext [sspi_client_context $Credentials -stream 1]
                lassign [sspi_step $SspiContext] status output
                if {[string length $output]} {
                    chan puts -nonewline $Socket $output
                    chan flush $Socket
                }
                if {$status ne "continue"} {
                    error "Unexpected status $status from sspi_step"
                }
            }
        }
        
        SERVERINIT {
            if {[dict get $_channels($chan) Blocking]} {
                _server_blocking_negotiate $chan
            } else {
                set State NEGOTIATING
                
            }
        }
    }

    return
}

proc twapi::ssl::_client_blocking_negotiate {chan} {
    variable _channels
    dict with _channels($chan) {
        set State NEGOTIATING
        set SspiContext [sspi_client_context $Credentials -stream 1]
        lassign [sspi_step $SspiContext] status output
        # Keep looping as long as the SSPI state machine tells us to continue negotiation
        while {$status eq "continue"} {
            # If the previous step required any output to be sent to remote end, do so
            if {[string length $output]} {
                chan puts -nonewline $Socket $output
                chan flush $Socket
            }

            set input [chan read $Socket 1]
            set more [chan pending input $Socket]
            if {$more > 0} {
                append input [chan read $Socket $more]
            }
            if {[chan eof $Socket]} {
                error "Unexpected EOF during SSL negotiation."
            }
            lassign [sspi_step $SspiContext $input] status output leftover
        }
        # Send output irrespective of status
        if {[string length $output]} {
            chan puts -nonewline $Socket $output
            chan flush $Socket
        }

        if {$status eq "done"} {
            set State OPEN
            if {[string length $leftover]} {
                set Input [sspi_decrypt_stream $SspiContext $leftover]
            }
        } else {
            # Should not happen. Negotiation failures will raise an error,
            # not return a value
            error "SSL negotiation failed: status $status."
        }
    }
    return
}



# Finite state machine for connection.
proc twapi::ssl::_fsm chan {
    trap {
        _fsm_run $chan
    } onerror {
        variable _channels
        if {[info exists _channels($chan)]} {
            dict set _channels($chan) State FAIL
            dict set _channels($chan) ErrorOptions [trapoptions]
            dict set _channels($chan) ErrorResult [trapresult]
        }
        rethrow
    }
}

proc twapi::ssl::_fsm_run chan {
    variable _channels

    dict with _channels($chan) {
        switch $State {
            CLIENTINIT {
                set SspiContext [sspi_client_context $Credentials -stream 1 -manualvalidation [expr {$Verifier eq ""}]]
                lassign [sspi_step $SspiContext] action output
                if {[string length $output]} {
                    lappend CipherOutQ $output
                }
                set State [dict! {done OPEN continue NEGOTIATING disconnected CLOSED}]
            }
            NEGOTIATING {
                _soinput $chan
                lassign [sspi_step $SspiContext [join $CipherInQ ""]] action output extra
                if {[string length $output]} {
                    lappend CipherOutQ $output
                }
                if {[string length $extra]} {
                    set CipherIn [list $extra]
                } else {
                    set CipherIn {}
                }
                set State [dict! {done OPEN continue NEGOTIATING disconnected CLOSED}]
            }
            default {
                error "Illegal state $State"
            }
        }
        _sooutput $chan
    }
    return
}



namespace eval twapi::ssl {
    namespace export initialize finalize blocking watch read write configure cget cgetall
    namespace ensemble create
}
