namespace eval twapi::ssl {
    variable _channels
    array set _channels {}

    namespace path [linsert [namespace path] 0 [namespace parent]]
}

interp alias {} twapi::ssl_socket {} twapi::ssl::_socket
proc twapi::ssl::_socket {args} {
    variable _channels

    parseargs args {
        myaddr.arg
        myport.int
        async
        server.arg
        certificate.arg
    } -setvars

    set chan [chan create {read write} [list [namespace current]]]

    set socket_args {}
    if {[info exists server]} {
        set type LISTENER
        lappend socket_args -server [list [namespace current]::_accept $chan]
    } else {
        set type CLIENT
    }

    foreach opt {myaddr myport} {
        if {[info exists $opt]} {
            lappend socket_args -$opt [set $opt]
        }
    }

    trap {
        set so [socket {*}$socket_args {*}$args]
        _init $chan $type $so
    } onerror {} {
        unset -nocomplain _channels($chan)
        catch { chan close $chan }
        catch { chan close $so }
        rethrow
    }

    return $chan
}

proc twapi::ssl::_accept {chan so raddr raport} {
    variable _channels

    trap {
        set chan [chan create {read write} [list [namespace current]::ssl $id]]
        _init $chan SERVER $so
        {*}[dict get _channels($chan) accept_callback] $chan $raddr $raport
    } onerror {
        catch {chan close $chan}
        unset -nocomplain _channels($chan)
        # No need to close socket - Tcl does that on error return
        rethrow
    }
    return
}

proc twapi::ssl::initialize {chan mode} {
    # All init is done in chan creation routine after base socket is created
    return {initialize finalize watch blocking read write configure cget cgetall}
}

proc twapi::ssl::finalize {chan} {
    variable _channels
    if {[info exists _channels($chan)]} {
        if {[dict exists $_channels($chan) socket]} {
            chan close [dict get $_channels($chan) socket]
        }
        unset _channels($chan)
    }
    return
}

proc twapi::ssl::blocking {chan mode} {
    variable _channels

    chan configure [_chansocket $chan] -blocking $mode
    dict set _channels($chan) blocking $mode
    return
}

proc twapi::ssl::watch {chan watchmask} {
    variable _channels
    dict set _channels($chan) watchmask $watchmask
    set so [dict get $_channels($chan) socket]
    if {"read" in $watchmask} {
        chan event $so readable [list [namespace current]::_readable_handler $chan]
    } else {
        chan event $so readable {}
    }

    if {"write" in $watchmask} {
        chan event $so writable [list [namespace current]::_writable_handler $chan]
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
    if {![dict get $_channels($chan) blocking]} {
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
    if {$opt eq "-certificates"} {
        error "Option -certificates is read-only."
    }

    chan configure [_chansocket $chan] $opt $val
    return
}

proc twapi::ssl::cget {chan opt} {
    variable _channels

    set so [_chansocket $chan]

    if {$opt eq "-certificates"} {
        return [dict get $_channels($chan) certificates]
    }

    return [chan configure $so $opt]
}

proc twapi::ssl::cgetall {chan} {
    set so [_chansocket $chan]
    set config [chan configure $so]
    lappend config -certificates [dict get $_channels($chan) certificates]
    return $config
}

proc twapi::ssl::_chansocket {chan} {
    variable _channels
    if {![info exists _channels($chan)]} {
        error "Channel $chan not found."
    }
    return [dict get $_channels($chan) socket]
}

proc twapi::ssl::_init {chan type so} {
    variable _channels

    # TBD - verify that -buffering none is the right thing to do
    # as the scripted channel interface takes care of this itself
    chan configure $so -translation binary -buffering none
    set _channels($chan) [list socket $so state INIT type $type blocking [chan configure $so -blocking] watchmask {}]
    if {$type eq "LISTENER"} {
        dict set _channels($chan) accept_callback $server
    }
}

proc twapi::ssl::_readable_handler {chan} {
    variable _channels
    if {[info exists _channels($chan)]} {
        if {"read" in [dict get $_channels($chan) watchmask]} {
            chan postevent $chan read
        }
    }
    return
}

proc twapi::ssl::_writable_handler {chan} {
    variable _channels
    if {[info exists _channels($chan)]} {
        if {"write" in [dict get $_channels($chan) watchmask]} {
            chan postevent $chan write
        }
    }
    return
}


namespace eval twapi::ssl {
    namespace export initialize finalize blocking watch read write configure cget cgetall
    namespace ensemble create
}
