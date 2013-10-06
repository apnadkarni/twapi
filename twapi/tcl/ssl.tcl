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
    #  CipherOutQ - list of encrypted data to send to socket
    #  CipherInQ - list of encrypted data read from socket
    #  PlainOutQ - list of plaintext data from app
    #  PlainInQ  - list of plaintext data to pass to app

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
        if {$type eq "CLIENT" && $async} {
            variable _channels
            while {[dict get $_channels($chan) State] eq "NEGOTIATING"} {
                _fsm $chan
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
        _init $chan SERVER $so [dict get $_channels($listener) Credentials] [dict get $_channels($listener) Verifier
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

    if {0} {
        Keep underlying socket always in non-blocking mode
        chan configure [_chansocket $chan] -blocking $mode
    }

    dict set _channels($chan) Blocking $mode
    return
}

proc twapi::ssl::watch {chan watchmask} {
    variable _channels
    dict set _channels($chan) WatchMask $watchmask
    set so [dict get $_channels($chan) Socket]
    if {"read" in $watchmask} {
        chan event $so readable [list [namespace current]::_so_readable_handler $chan]
    } else {
        chan event $so readable {}
    }

    if {"write" in $watchmask} {
        chan event $so writable [list [namespace current]::_so_writable_handler $chan]
    } else {
        chan event $so writable {}
    }
    
    return
}

proc twapi::ssl::read {chan nbytes} {
    variable _channels

    set so [_chansocket $chan]

    set data [chan read $so $nbytes]
    if {[string length $data]} {
        return $data
    }
    if {[chan eof $so]} {
        return ""
    }
    if {![dict get $_channels($chan) Blocking]} {
        return -code error EAGAIN
    }
    error "Unknown error. Data could not be read from channel $so"
}

proc twapi::ssl::write {chan data} {
    set so [_chansocket $chan]
    chan puts -nonewline $so $data
    flush $so
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
                              CipherInQ {} CipherOutQ {} \
                              PlainInQ {} PlainOutQ {}]
    if {$type eq "LISTENER"} {
        dict set _channels($chan) AcceptCallback $accept_callback
    } else {
        _fsm $chan
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
                if {[catch {sspi_delete_context $SspiContext} msg]} {
                    # TBD - debug log
                }
            }
        }
        unset _channels($chan)
    }
}

proc twapi::ssl::_so_readable_handler {chan} {
    variable _channels

    if {[info exists _channels($chan)]} {
        if {"read" in [dict get $_channels($chan) WatchMask]} {
            chan postevent $chan read
        }
    }
    return
}

proc twapi::ssl::_so_writable_handler {chan} {
    variable _channels
    if {[info exists _channels($chan)]} {
        if {"write" in [dict get $_channels($chan) WatchMask]} {
            chan postevent $chan write
        }
    }
    return
}

# Reads from specified socket and stores in input queue
# Assumes socket is non-blocking.
# Returns 1 if socket state is eof
proc twapi::ssl::_soinput {chan} { 
    variable _channels

    dict with _channels($chan) {
        set in [chan read $Socket]
        if {[string length $in]} {
            lappend CipherInQ $in
        }
        return [eof $Socket]
    }
}

proc twapi::ssl::_sooutput {chan} {
    variable _channels
    dict with _channels($chan) {
        foreach out $CipherOutQ {
            puts -nonewline $Socket $out
        }
        flush $Socket
        set CipherOutQ {}
    }
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
                    set CipherInQ [list $extra]
                } else {
                    set CipherInQ {}
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
