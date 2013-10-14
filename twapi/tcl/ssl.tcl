namespace eval twapi::ssl {
    # Each element of _channels is dictionary with the following keys
    #  Socket - the underlying socket
    #  State - SERVERINIT, CLIENTINIT, LISTENERINIT, OPEN, NEGOTIATING, CLOSED
    #  Type - SERVER, CLIENT, LISTENER
    #  Blocking - 0/1 indicating whether blocking or non-blocking channel
    #  WatchMask - list of {read write} indicating what events to post
    #  Target - Name for server cert
    #  Credentials - credentials handle to use for local end of connection
    #  FreeCredentials - if credentials should be freed on connection cleanup
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
        server.arg
        peersubject.arg
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
        if {[info exists peersubject]} {
            error "Option -peersubject cannot be specified for with -server"
        }
        set peersubject ""
        set type LISTENER
        lappend socket_args -server [list [namespace current]::_accept $chan]
        if {[llength $credentials] == 0} {
            error "Option -credentials must be specified for server sockets"
        }
    } else {
        if {![info exists peersubject]} {
            set peersubject [lindex $args 0]
        }
        set server ""
        lappend socket_args -async; # Always async, we will explicitly block
        set type CLIENT
    }

    trap {
        set so [socket {*}$socket_args {*}$args]
        _init $chan $type $so $credentials $peersubject [lrange $verifier 0 end] $server
        if {$type eq "CLIENT"} {
            if {! $async} {
                _client_blocking_negotiate $chan
            }
        }
    } onerror {} {
        variable _channels
        if {![info exists _channel($chan)]} {
            catch {chan close $so}
        }
        catch {chan close $chan}
        rethrow
    }

    return $chan
}


proc twapi::ssl::_accept {listener so raddr raport} {
    variable _channels

    trap {
        set chan [chan create {read write} [list [namespace current]]]
        _init $chan SERVER $so [dict get $_channels($listener) Credentials] "" [dict get $_channels($listener) Verifier]
        # TBD - should we call this AFTER verification / negotiation is done?
        {*}[dict get $_channels($listener) AcceptCallback] $chan $raddr $raport
    } onerror {} {
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

    dict with _channels($chan) {
        set Blocking $mode
        chan configure $Socket -blocking $mode
        if {$mode == 0} {
            # Since we need to negotiate SSL we always have socket event
            # handlers irrespective of the state of the watch mask
            chan event $Socket readable [list [namespace current]::_so_read_handler $chan]
            chan event $Socket writable [list [namespace current]::_so_write_handler $chan]
        } else {
            chan event $Socket readable {}
            chan event $Socket writable {}
        }
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
        # Try to read more bytes if don't have enough AND conn is open
        set status ok
        if {[string length $Input] < $nbytes && $State eq "OPEN"} {
            if {$Blocking} {
                # The channel does not compress so we need to read in
                # at least $needed bytes. Because of SSL overhead, we may actually
                # need even more
                set status ok
                while {[set ninput [string length $Input]] < $nbytes} {
                    set needed [expr {$nbytes-$ninput}]
                    set data [chan read $Socket $needed]
                    if {[string length $data]} {
                        lassign [sspi_decrypt_stream $SspiContext $data] status plaintext
                        append Input $plaintext
                        if {$status ne "ok"} break
                    }
                    
                    if {[string length $data] < $needed} {
                        set status eof
                        break
                    }
                }
            } else {
                # Non-blocking - read all that we can
                set status ok
                set data [chan read $Socket]
                if {[string length $data]} {
                    lassign [sspi_decrypt_stream $SspiContext $data] status plaintext
                    append Input $plaintext
                } else {
                    if {[chan eof $Socket]} {
                        set status eof
                    }
                }
                if {[string length $Input] == 0} {
                    # Do not have enough data. See if connection closed
                    # TBD - also handle status == renegotiate
                    if {$status eq "ok"} {
                        # Not closed, just waiting for data
                        return -code error EAGAIN
                    }
                }
            }
        }

        # TBD - use inline K operator to make this faster? Probably no use
        # since Input is also referred to from _channels($chan)
        set ret [string range $Input 0 $nbytes-1]
        set Input [string range $Input $nbytes end]
        if {"read" in [dict get $_channels($chan) WatchMask] && [string length $Input]} {
            chan postevent $chan read
        }
        if {$status ne "ok"} {
            # TBD - handle renegotiate
            lassign [sspi_shutdown_context $SspiContext] _ outdata
            if {[string length $outdata]} {
                puts -nonewline $Socket $outdata
            }
            set State CLOSED
        }
        return $ret;            # Note ret may be ""
    }
}

proc twapi::ssl::write {chan data} {
    variable _channels

    # This is not inside the dict with because _negotiate will update the dict
    if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
        _negotiate $chan
        if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            # If a blocking channel, should have come back with negotiation
            # complete. If non-blocking, return EAGAIN to indicate channel
            # not open yet.
            if {[dict get $_channels($chan) Blocking]} {
                error "SSL negotiation failed on blocking channel" 
            } else {
                # TBD - should we just accept the data ?
                return -code error EAGAIN
            }
        }
    }

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

proc twapi::ssl::_init {chan type so creds peersubject verifier {accept_callback {}}} {
    variable _channels

    # TBD - verify that -buffering none is the right thing to do
    # as the scripted channel interface takes care of this itself
    # We always set -blocking to 0 and take care of blocking ourselves
    chan configure $so -translation binary -buffering none
    set _channels($chan) [list Socket $so \
                              State ${type}INIT \
                              Type $type \
                              Blocking 1 \
                              WatchMask {} \
                              Verifier $verifier \
                              SspiContext {} \
                              PeerSubject $peersubject \
                              Input {} Output {}]

    if {[llength $creds]} {
        set free_creds 0
    } else {
        set creds [sspi_acquire_credentials -package ssl -role client -credentials [sspi_schannel_credentials]]
        set free_creds 1
    }
    dict set _channels($chan) Credentials $creds
    dict set _channels($chan) FreeCredentials $free_creds

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
            if {[info exists SspiContext]} {
                if {$State eq "OPEN"} {
                    lassign [sspi_shutdown_context $SspiContext] _ outdata
                    if {[string length $outdata] && [info exists Socket]} {
                        if {[catch {puts -nonewline $Socket $outdata} msg]} {
                            # TBD - debug log
                        }
                    }
                }
                if {[catch {sspi_delete_context $SspiContext} msg]} {
                    # TBD - debug log
                }
            }
            if {[info exists Socket]} {
                if {[catch {chan close $Socket} msg]} {
                    # TBD - debug log socket close error
                }
            }
            if {[info exists Credentials] && $FreeCredentials} {
                if {[catch {sspi_free_credentials $Credentials} msg]} {
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
        if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            _negotiate $chan
        }

        if {"read" in [dict get $_channels($chan) WatchMask]} {
            chan postevent $chan read
        } else {
            # We are not asked to generate read events, turn off the read
            # event handler unless we are negotiating
            if {[dict get $_channels($chan) State] ni {SERVERINIT CLIENTINIT NEGOTIATING}} {
                chan event [dict get $_channels($chan) Socket] readable {}
            }
        }
    }
    return
}

proc twapi::ssl::_so_write_handler {chan} {
    variable _channels

    if {[info exists _channels($chan)]} {
        dict with _channels($chan) {}

        # If we are not actually asked to generate write events,
        # the only time we want a write handler is on a client -async
        # Once it runs, we never want it again else it will keep triggering
        # as sockets are always writable
        if {"write" ni $WatchMask} {
            chan event $Socket writable {}
        }

        if {$State in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            _negotiate $chan
        }

        # Do not use local var $State because _negotiate might have updated it
        if {"write" in $WatchMask && [dict get $_channels($chan) State] eq "OPEN"} {
            chan postevent $chan write
        }
    }
    return
}

proc twapi::ssl::_negotiate chan {
    trap {
        _negotiate2 $chan
    } onerror {} {
        variable _channels
        if {[info exists _channels($chan)]} {
            dict set _channels($chan) State FAIL
            dict set _channels($chan) ErrorOptions [trapoptions]
            dict set _channels($chan) ErrorResult [trapresult]
        }
        rethrow
    }
}

proc twapi::ssl::_negotiate2 {chan} {
    variable _channels
        
    dict with _channels($chan) {}; # dict -> local vars
    switch $State {
        NEGOTIATING {
            if {$Blocking} {
                error "Internal error: NEGOTIATING state not expected on a blocking ssl socket"
            }

            set data [chan read $Socket]
            if {[string length $data] == 0} {
                if {[chan eof $Socket]} {
                    error "Unexpected EOF during SSL negotiation"
                } else {
                    # No data yet, just keep waiting
                    return
                }
            } else {
                lassign [sspi_step $SspiContext $data] status outdata leftover
                if {[string length $outdata]} {
                    chan puts -nonewline $Socket $outdata
                    chan flush $Socket
                }
                switch $status {
                    done {
                        if {[string length $leftover]} {
                            lassign [sspi_decrypt_stream $SspiContext $leftover] status plaintext
                            dict append _channels($chan) Input $plaintext
                            if {$status ne "ok"} {
                                # TBD - shutdown channel or let _cleanup do it?
                            }
                        }
                        _open $chan
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

        CLIENTINIT {
            if {$Blocking} {
                _client_blocking_negotiate $chan
            } else {
                dict set _channels($chan) State NEGOTIATING
                set SspiContext [sspi_client_context $Credentials -stream 1 -target $PeerSubject -manualvalidation [expr {[llength $Verifier] > 0}]]
                dict set _channels($chan) SspiContext $SspiContext
                lassign [sspi_step $SspiContext] status outdata
                if {[string length $outdata]} {
                    chan puts -nonewline $Socket $outdata
                    chan flush $Socket
                }
                if {$status ne "continue"} {
                    error "Unexpected status $status from sspi_step"
                }
            }
        }
        
        SERVERINIT {
            if {$Blocking} {
                _server_blocking_negotiate $chan
            } else {
                dict set _channels($chan) State NEGOTIATING
                set data [chan read $Socket]
                if {[string length $data] == 0} {
                    if {[chan eof $Socket]} {
                        error "Unexpected EOF during SSL negotiation"
                    } else {
                        # No data yet, just keep waiting
                        return
                    }
                } else {
                    set SspiContext [sspi_server_context $Credentials $data -stream 1]
                    dict set _channels($chan) SspiContext $SspiContext
                    lassign [sspi_step $SspiContext] status outdata leftover
                    if {[string length $outdata]} {
                        chan puts -nonewline $Socket $outdata
                        chan flush $Socket
                    }
                    switch $status {
                        done {
                            if {[string length $leftover]} {
                                lassign [sspi_decrypt_stream $SspiContext $leftover] status plaintext
                                dict append _channels($chan) Input $plaintext
                                if {$status ne "ok"} {
                                    # TBD - shut down channel
                                }
                            }
                            _open $chan
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

        default {
            error "Internal error: _negotiate called in state [dict get $_channels($chan) State]"
        }
    }

    return
}

proc twapi::ssl::_client_blocking_negotiate {chan} {
    variable _channels
    dict with _channels($chan) {
        set State NEGOTIATING
        set SspiContext [sspi_client_context $Credentials -stream 1 -target $PeerSubject -manualvalidation [expr {[llength $Verifier] > 0}]]
    }
    return [_blocking_negotiate_loop $chan]
}

proc twapi::ssl::_server_blocking_negotiate {chan} {
    variable _channels
    dict set _channels($chan) State NEGOTIATING
    set so [dict get $_channels($chan) Socket]
    set indata [_blocking_read $so]
    if {[chan eof $so]} {
        error "Unexpected EOF during SSL negotiation."
    }
    dict set _channels($chan) SspiContext [sspi_server_context [dict get $_channels($chan) Credentials] $indata -stream 1]
    return [_blocking_negotiate_loop $chan]
}

proc twapi::ssl::_blocking_negotiate_loop {chan} {
    variable _channels
    dict with _channels($chan) {}; # dict -> local vars

    lassign [sspi_step $SspiContext] status outdata
    # Keep looping as long as the SSPI state machine tells us to 
    while {$status eq "continue"} {
        # If the previous step had any output, send it out
        if {[string length $outdata]} {
            chan puts -nonewline $Socket $outdata
            chan flush $Socket
        }

        set indata [_blocking_read $Socket]
        if {[chan eof $Socket]} {
            error "Unexpected EOF during SSL negotiation."
        }
        lassign [sspi_step $SspiContext $indata] status outdata leftover
    }

    # Send output irrespective of status
    if {[string length $outdata]} {
        chan puts -nonewline $Socket $outdata
        chan flush $Socket
    }

    if {$status eq "done"} {
        if {[string length $leftover]} {
            lassign [sspi_decrypt_stream $SspiContext $leftover] status plaintext
            dict append _channels($chan) Input $plaintext
            if {$status ne "ok"} {
                # TBD - shut down channel
            }
        }
        _open $chan
    } else {
        # Should not happen. Negotiation failures will raise an error,
        # not return a value
        error "SSL negotiation failed: status $status."
    }
    return
}

proc twapi::ssl::_blocking_read {so} {
    # Read from a blocking socket. We do not know how much data is needed
    # so read a single byte and then read any pending
    set input [chan read $so 1]
    if {[string length $input]} {
        set more [chan pending input $so]
        if {$more > 0} {
            append input [chan read $so $more]
        }
    }    
    return $input
}

proc twapi::ssl::_open {chan} {
    variable _channels

    dict with _channels($chan) {}; # dict -> local vars

    if {[llength $Verifier] == 0} {
        # No verifier specified. In this case, we would not have specified
        # -manualvalidation in creating the context and the system would
        # have done the verification already for client. For servers,
        # there is no verification of clients to be done by default

        # TBD - what about server accept callback ?
        dict set _channels($chan) State OPEN
        return 1
    }

    trap {
        if {[{*}$Verifier $SspiContext]} {
            dict set _channels($chan) State OPEN
            return 1
        }
    } onerror {} {
        # TBD - debug log
    }

    dict set _channels($chan) State CLOSED
    catch {close $Socket}
    return 0
}

namespace eval twapi::ssl {
    namespace export initialize finalize blocking watch read write configure cget cgetall
    namespace ensemble create
}
