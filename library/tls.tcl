#
# Copyright (c) 2012-2020, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license
namespace eval twapi::tls {
    # Each element of _channels is dictionary with the following keys
    #  Socket - the underlying socket. This key will not exist if
    #   socket has been closed.
    #  State - SERVERINIT, CLIENTINIT, LISTENERINIT, OPEN, NEGOTIATING, CLOSED
    #  Type - SERVER, CLIENT, LISTENER
    #  Blocking - 0/1 indicating whether blocking or non-blocking channel
    #  WatchMask - list of {read write} indicating what events to post
    #  Target - Name for server cert
    #  Credentials - credentials handle to use for local end of connection
    #  FreeCredentials - if credentials should be freed on connection cleanup
    #  AcceptCallback - application callback on a listener and server socket.
    #    On listener, it is the accept command prefix. On a server 
    #    (accepted socket) it is the prefix plus arguments passed to
    #    accept callback. On client and on servers sockets initialized
    #    with starttls, this key must NOT be present
    #  SspiContext - SSPI context for the connection
    #  Input  - plaintext data to pass to app
    #  Output - plaintext data to encrypt and output
    #  ReadEventPosted - if this key exists, a chan postevent for read
    #    is already in progress and a second one should not be posted
    #  WriteEventPosted - if this key exists, a chan postevent for write
    #    is already in progress and a second one should not be posted
    #  WriteDisabled - 0 normally. Set to 1 on a half-close

    variable _channels
    array set _channels {}

    # Socket command - Tcl socket by default.
    variable _socket_cmd ::socket

    namespace path [linsert [namespace path] 0 [namespace parent]]

}

proc twapi::tls_socket_command {args} {
    set orig_command $tls::_socket_cmd
    if {[llength $args] == 1} {
        set tls::_socket_cmd [lindex $args 0]
    } elseif {[llength $args] != 0} {
        error "wrong # args: should be \"tls_socket_command ?cmd?\""
    }
    return $orig_command
}

interp alias {} twapi::tls_socket {} twapi::tls::_socket
proc twapi::tls::_socket {args} {
    variable _channels
    variable _socket_cmd

    debuglog [info level 0]

    parseargs args {
        myaddr.arg
        myport.int
        async
        socketcmd.arg
        server.arg
        peersubject.arg
        requestclientcert
        {credentials.arg {}}
        {verifier.arg {}}
    } -setvars

    set chan [chan create {read write} [list [namespace current]]]
    # NOTE: We were originally using badargs! instead of error to raise
    # exceptions. However that lands up bypassing the trap because of
    # the way badargs! is implemented. So stick to error.
    trap {
        set socket_args {}
        foreach opt {myaddr myport} {
            if {[info exists $opt]} {
                lappend socket_args -$opt [set $opt]
            }
        }
        if {$async} {
            lappend socket_args -async
        }

        if {[info exists server]} {
            if {$server eq ""} {
                error "Cannot specify an empty value for -server."
            }

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
            set requestclientcert 0; #  Not valid for client side
            set server ""
            set type CLIENT
        }

        if {[info exists socketcmd]} {
            if {$socketcmd eq ""} {
                set socketcmd ::socket
            }
        } else {
            set socketcmd $_socket_cmd
        }
    } onerror {} {
        catch {chan close $chan}
        rethrow
    }
    trap {
        set so [$socketcmd {*}$socket_args {*}$args]
        _init $chan $type $so $credentials $peersubject $requestclientcert [lrange $verifier 0 end] $server

        if {$type eq "CLIENT"} {
            if {$async} {
                chan event $so writable [list [namespace current]::_so_write_handler $chan]
            } else {
                _client_blocking_negotiate $chan
                if {(![info exists _channels($chan)]) ||
                    [dict get $_channels($chan) State] ne "OPEN"} {
                    if {[info exists _channels($chan)] &&
                        [dict exists $_channels($chan) ErrorResult]} {
                        error [dict get $_channels($chan) ErrorResult]
                    } else {
                        error "TLS negotiation aborted"
                    }
                }
            }
        }
    } onerror {} {
        # If _init did not even go as far initializing _channels($chan),
        # close socket ourselves. If it was initialized, the socket
        # would have been closed even on error
        if {![info exists _channels($chan)]} {
            catch {chan close $so}
        }
        catch {chan close $chan}
        # DON'T ACCESS _channels HERE ON
        if {[string match "wrong # args*" [trapresult]]} {
            badargs! "wrong # args: should be \"tls_socket ?-credentials creds? ?-verifier command? ?-peersubject peer? ?-myaddr addr? ?-myport myport? ?-async? host port\" or \"tls_socket ?-credentials creds? ?-verifier command? -server command ?-myaddr addr? port\""
        } else {
            rethrow
        }
    }

    return $chan
}

interp alias {} twapi::starttls {} twapi::tls::_starttls
proc twapi::tls::_starttls {so args} {
    variable _channels

    debuglog [info level 0]

    trap {
        parseargs args {
            server
            requestclientcert
            peersubject.arg
            {credentials.arg {}}
            {verifier.arg {}}
        } -setvars -maxleftover 0

        if {$server} {
            if {[info exists peersubject]} {
                badargs! "Option -peersubject cannot be specified with -server."
            }
            if {[llength $credentials] == 0} {
                error "Option -credentials must be specified for server sockets."
            }
            set peersubject ""
            set type SERVER
        } else {
            set requestclientcert 0; # Ignored for client side
            if {![info exists peersubject]} {
                # TBD - even if verifier is specified ?
                badargs! "Option -peersubject must be specified for client connections."
            }
            set type CLIENT
        }
        set chan [chan create {read write} [list [namespace current]]]
    } onerror {} {
        chan close $so
        rethrow
    }
    trap {
        # Get config from the wrapped socket and reset its handlers
        # Do not get all options because that results in reverse name
        # lookups for -peername and -sockname causing a stall.
        foreach opt {
            -blocking -buffering -buffersize -encoding -eofchar -translation
        } {
            lappend so_opts $opt [chan configure $so $opt]
        }

        # NOTE: we do NOT save read and write handlers and attach
        # them to the new channel because the channel name is different.
        # Thus in most cases the callbacks, which often are passed the
        # channel name as an arg, would not be valid. It is up
        # to the caller to reestablish handlers
        # TBD - maybe keep handlers but replace $so with $chan in them ?
        chan event $so readable {}
        chan event $so writable {}
        _init $chan $type $so $credentials $peersubject $requestclientcert [lrange $verifier 0 end] ""
        # Copy saved config to wrapper channel
        chan configure $chan {*}$so_opts
        if {$type eq "CLIENT"} {
            if {[dict get $_channels($chan) Blocking]} {
                _client_blocking_negotiate $chan
                if {(![info exists _channels($chan)]) ||
                    [dict get $_channels($chan) State] ne "OPEN"} {
                    if {[info exists _channels($chan)] &&
                        [dict exists $_channels($chan) ErrorResult]} {
                        error [dict get $_channels($chan) ErrorResult]
                    } else {
                        error "TLS negotiation aborted"
                    }
                }
            } else {
                _negotiate $chan
            }
        } else {
            # Note: unlike the tls_socket server case, here we
            # do not need to switch a blocking socket to non-blocking
            # and then switch back, primarily because the socket
            # is already open and there is no need for a callback
            # when connection opens.
            if {! [dict get $_channels($chan) Blocking]} {
                chan configure $so -blocking 0
                chan event $so readable [list [namespace current]::_so_read_handler $chan]
            }
            _negotiate $chan
        }
    } onerror {} {
        # If _init did not even go as far initializing _channels($chan),
        # close socket ourselves. If it was initialized, the socket
        # would have been closed even on error
        if {![info exists _channels($chan)]} {
            catch {chan close $so}
        }
        catch {chan close $chan}
        # DON'T ACCESS _channels HERE ON
        if {[string match "wrong # args*" [trapresult]]} {
            badargs! "wrong # args: should be \"tls_socket ?-credentials creds? ?-verifier command? ?-peersubject peer? ?-myaddr addr? ?-myport myport? ?-async? host port\" or \"tls_socket ?-credentials creds? ?-verifier command? -server command ?-myaddr addr? port\""
        } else {
            rethrow
        }
    }

    return $chan
}

interp alias {} twapi::tls_state {} twapi::tls::_state
proc twapi::tls::_state {chan} {
    variable _channels
    if {![info exists _channels($chan)]} {
       twapi::badargs! "Not a valid TLS channel." 
    }
    return [dict get $_channels($chan) State]
}

interp alias {} twapi::tls_handshake {} twapi::tls::_handshake
proc twapi::tls::_handshake {chan} {
    variable _channels
    if {![info exists _channels($chan)]} {
       twapi::badargs "Not a valid TLS channel." 
    }

    dict with _channels($chan) {}

    # For a blocking channel, complete the handshake before returning
    if {$Blocking} {
        switch -exact $State {
            NEGOTIATING - CLIENTINIT - SERVERINIT {
                _negotiate2 $chan
            }
            OPEN {}
            LISTERNERINIT {
                error "Cannot do a TLS handshake on a listening socket."
            }
            CLOSED -
            default {
                error "Channel has been closed or in error state."
            }
        }
    } else {
        # For non-blocking channels, simply return the state
        switch -exact -- $State {
            OPEN {}
            CLIENTINIT - SERVERINIT - LISTENERINIT - NEGOTIATING {
                return 0
            }
            CLOSED - default {
                error "Channel has been closed or in error state."
            }
        }
    }
    return 1
}

proc twapi::tls::_accept {listener so raddr raport} {
    variable _channels

    debuglog [info level 0]

    trap {
        set chan [chan create {read write} [list [namespace current]]]
        _init $chan SERVER $so [dict get $_channels($listener) Credentials] "" [dict get $_channels($listener) RequestClientCert] [dict get $_channels($listener) Verifier] [linsert [dict get $_channels($listener) AcceptCallback] end $chan $raddr $raport]
        # If we negotiate the connection, the socket is blocking so
        # will hang the whole operation. Instead we mark it non-blocking
        # and the switch back to blocking when the connection gets opened.
        # For accepts to work, the event loop has to be running anyways.
        chan configure $so -blocking 0
        chan event $so readable [list [namespace current]::_so_read_handler $chan]
        _negotiate $chan
    } onerror {} {
        catch {_cleanup $chan}
        rethrow
    }
    return
}

proc twapi::tls::initialize {chan mode} {
    debuglog [info level 0]

    # All init is done in chan creation routine after base socket is created
    return {initialize finalize watch blocking read write configure cget cgetall}
}

proc twapi::tls::finalize {chan} {
    debuglog [info level 0]
    _cleanup $chan
    return
}

proc twapi::tls::blocking {chan mode} {
    debuglog [info level 0]

    variable _channels

    dict set _channels($chan) Blocking $mode

    if {![dict exists $_channels($chan) Socket]} {
        # We do not currently generate an error because the Tcl socket
        # command does not either on a fconfigure when remote has
        # closed connection
        return
    }
    set so [dict get $_channels($chan) Socket]
    fconfigure $so -blocking $mode

    # There is an issue with Tcl sockets created with -async switching
    # from blocking->non-blocking->blocking and writing to the socket
    # before connection is fully open. The internal buffers containing
    # data that was written before the connection was open do not get
    # flushed even if there was an explicit flush call by the application.
    # Doing a flush after changing blocking mode seems to fix this
    # problem. TBD - check if still the case
    flush $so

    # TBD - Should we change handlers BEFORE flushing?

    # The flush may recursively call event handler (possibly) which
    # may change state so have to retrieve values from _channels again.
    if {![dict exists $_channels($chan) Socket]} {
        return
    }
    set so [dict get $_channels($chan) Socket]

    if {[dict get $_channels($chan) Blocking] == 0} {
        # Non-blocking
        # Since we need to negotiate TLS we always have socket event
        # handlers irrespective of the state of the watch mask
        chan event $so readable [list [namespace current]::_so_read_handler $chan]
        chan event $so writable [list [namespace current]::_so_write_handler $chan]
    } else {
        # TBD - is this right? Application may have file event handlers even
        # on blocking sockets
        chan event $so readable {}
        chan event $so writable {}
    }
    return
}

proc twapi::tls::watch {chan watchmask} {
    debuglog [info level 0]
    variable _channels

    dict set _channels($chan) WatchMask $watchmask

    if {"read" in $watchmask} {
        debuglog "[info level 0]: read"
        # Post a read even if we already have input or if the
        # underlying socket has gone away.
        # TBD - do we have a mechanism for continuously posting
        # events when socket has gone away ? Do we even post once
        # when socket is closed (on error for example)
        if {[string length [dict get $_channels($chan) Input]] || ![dict exists $_channels($chan) Socket]} {
            _post_read_event $chan
        }
        # Turn read handler back on in case it had been turned off.
        chan event [dict get $_channels($chan) Socket] readable [list [namespace current]::_so_read_handler $chan]
    }

    if {"write" in [dict get $_channels($chan) WatchMask]} {
        debuglog "[info level 0]: write"
        if {[dict get $_channels($chan) State] in {OPEN NEGOTIATING CLOSED} } {
            _post_write_event $chan
        }
        # TBD - do we need to turn write handler back on?
    }

    return
}

proc twapi::tls::read {chan nbytes} {
    variable _channels

    debuglog [info level 0]

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
                error "TLS negotiation failed on blocking channel" 
            } else {
                return -code error EAGAIN
            }
        }
    }

    dict with _channels($chan) {
        # Either in OPEN or CLOSED state. For the latter, if an error is
        # present, immediately raise it else go on to return any pending data.
        if {$State eq "CLOSED" && [info exists ErrorResult]} {
            error $ErrorResult
        }
        # Try to read more bytes if don't have enough AND conn is open
        set status ok
        if {[string length $Input] < $nbytes && $State eq "OPEN"} {
            if {$Blocking} {
                # For blocking channels, we do not want to block if some
                # bytes are already available. The refchan will call us
                # with number of bytes corresponding to its buffer size,
                # not what app's read call has asked. It expects us
                # to return whatever we have (but at least one byte)
                # and block only if nothing is available
                while {[string length $Input] == 0 && $status eq "ok"} {
                    # The channel does not compress so we need to read in
                    # at least $needed bytes. Because of TLS overhead, we may
                    # actually need even more
                    set status ok
                    set data [_blocking_read $Socket]
                    if {[string length $data]} {
                        lassign [sspi_decrypt_stream $SspiContext $data] status plaintext
                        # Note plaintext might be "" if complete cipher block
                        # was not received
                        append Input $plaintext
                    } else {
                        set status eof
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
            _post_read_event $chan
        }
        if {$status ne "ok"} {
            # TBD - handle renegotiate
            debuglog "read: setting State CLOSED"

            # Need a EOF event even if read event posted. See Bug #203
            _post_eof_event $chan
            set State CLOSED
            lassign [sspi_shutdown_context $SspiContext] _ outdata
            if {[info exists Socket]} {
                if {[string length $outdata] && $status ne "eof"} {
                    puts -nonewline $Socket $outdata
                }
                catch {close $Socket}
                unset Socket
            }
        }
        return $ret;            # Note ret may be ""
    }
}

proc twapi::tls::write {chan data} {
    variable _channels

    set datalen [string length $data]
    debuglog "twapi::tls::write: $chan, $datalen bytes"

    # This is not inside the dict with below because _negotiate will update the dict
    if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
        _negotiate $chan
        if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            if {[dict get $_channels($chan) Blocking]} {
                # If a blocking channel, negotiation should have completed
                error "TLS negotiation failed on blocking channel" 
            } else {
                # TBD - which of the following alternatives to use?
                if {1} {
                    # Store for later output once connection is open
                    debuglog "twapi::tls::write conn not open, appending $datalen bytes to pending output"
                    dict append _channels($chan) Output $data
                    return $datalen
                } else {
                    # If non-blocking, return EAGAIN to indicate channel
                    # not open yet.
                    debuglog "twapi::tls::write returning EAGAIN"
                    return -code error EAGAIN
                }
            }
        }
    }

    dict with _channels($chan) {
        debuglog "twapi::tls::write state $State"
        switch $State {
            CLOSED {
                # Just like a Tcl socket, we do not raise an error on a
                # write to a closed socket. Simply throw away the data/
                # However, if an error already exists (negotiation fail)
                # raise it.
                if {[info exists ErrorResult]} {
                    error $ErrorResult
                }
            }
            OPEN {
                if {$WriteDisabled} {
                    error "Channel closed for output."
                }
                # There might be pending output if channel has just
                # transitioned to OPEN state
                _flush_pending_output $chan
                # TBD - use sspi_encrypt_and_write instead
                chan puts -nonewline $Socket [sspi_encrypt_stream $SspiContext $data]
                flush $Socket
            }
            default {
                append Output $data
            }
        }
    }
    debuglog "twapi::tls::write returning $datalen"
    return $datalen
}

proc twapi::tls::configure {chan opt val} {
    debuglog [info level 0]
    # Does not make sense to change creds and verifier after creation
    switch $opt {
        -context -
        -verifier -
        -credentials {
            error "$opt is a read-only option."
        }
        default {
            chan configure [_chansocket $chan] $opt $val
        }
    }

    return
}

proc twapi::tls::cget {chan opt} {
    debuglog [info level 0]
    variable _channels

    switch $opt {
        -credentials {
            return [dict get $_channels($chan) Credentials]
        }
        -verifier {
            return [dict get $_channels($chan) Verifier]
        }
        -context {
            return [dict get $_channels($chan) SspiContext]
        }
        -error {
            if {[dict exists $_channels($chan) ErrorResult]} {
                set result "[dict get $_channels($chan) ErrorResult]"
                if {$result ne ""} {
                    return $result
                }
            }
            # Get -error from underlying socket
            # -error should not raise an error but return the error as result
            catch {chan configure [_chansocket $chan] -error} result
            return $result
        }
        default {
            return [chan configure [_chansocket $chan] $opt]
        }
    }
}

proc twapi::tls::cgetall {chan} {
    debuglog [info level 0]
    variable _channels
    dict with _channels($chan) {
        if {[info exists Socket]} {
            # First get all options underlying socket supports. Note this may
            # or may not a Tcl native socket.
            array set so_config [chan configure $Socket]
            # Only return options that are not owned by the core channel code
            # and apply to the $chan itself.
            foreach {opt val} [chan configure $Socket] {
                if {$opt ni {-blocking -buffering -buffersize -encoding -eofchar -translation}} {
                    lappend config $opt $val
                }
            }
        }
        lappend config -credentials $Credentials \
        -verifier $Verifier \
        -context $SspiContext
    }
    return $config
}

# Implement a half-close command since Tcl does not support it for
# reflected channels.
interp alias {} twapi::tls_close {} twapi::tls::_close
proc twapi::tls::_close {chan {direction ""}} {

    if {$direction in {read r re rea}} {
        error "Half close of input side not currently supported for TLS sockets."
    }

    # We handle write-side half-closes. Let Tcl close handle everything else.
    if {$direction ni {write w wr wri writ}} {
        return [close $chan]
    }

    # Closing the write side of the channel

    variable _channels

    dict with _channels($chan) {}
    if {$State eq "CLOSED"} return
    if {$State ne "OPEN"} {
        error "Connection not in OPEN state."
    }
    flush $chan
    # Note state may have changed
    if {[dict get $_channels($chan) State] ne "OPEN"} {
        return
    }
    # Flush internally buffered, if any. Can happen if we buffered
    # data before TLS negotiation was complete.
    _flush_pending_output $chan
    close $Socket write
    dict set _channels($chan) WriteDisabled 1
    return
}

proc twapi::tls::_chansocket {chan} {
    debuglog [info level 0]
    variable _channels
    if {![info exists _channels($chan)]} {
        error "Channel $chan not found."
    }
    if {![dict exists $_channels($chan) Socket]} {
        set error "Socket not connected."
        if {[dict exists $_channels($chan) ErrorResult]} {
            append error " [dict get $_channels($chan) ErrorResult]"
        }
        error $error
    }
    return [dict get $_channels($chan) Socket]
}

proc twapi::tls::_init {chan type so creds peersubject requestclientcert verifier {accept_callback {}}} {
    debuglog [info level 0]
    variable _channels

    # TBD - verify that -buffering none is the right thing to do
    # as the scripted channel interface takes care of this itself
    chan configure $so -translation binary -buffering none
    set _channels($chan) [list Socket $so \
                              State ${type}INIT \
                              Type $type \
                              Blocking [chan configure $so -blocking] \
                              WatchMask {} \
                              WriteDisabled 0 \
                              RequestClientCert $requestclientcert \
                              Verifier $verifier \
                              SspiContext {} \
                              PeerSubject $peersubject \
                              Input {} Output {}]

    if {[llength $creds]} {
        set free_creds 0
    } else {
        set creds [sspi_acquire_credentials -package tls -role client -credentials [sspi_schannel_credentials]]
        set free_creds 1
    }
    dict set _channels($chan) Credentials $creds
    dict set _channels($chan) FreeCredentials $free_creds

    # See SF issue #178. Need to supply -usesuppliedcreds to sspi_client_context
    # else servers that request (even optionally) client certs might fail since
    # we do not currently implement incomplete credentials handling. This
    # option will prevent schannel from trying to automatically look up client
    # certificates.
    dict set _channels($chan) UseSuppliedCreds 0; # TBD - make this use settable option

    if {[string length $accept_callback] &&
        ($type eq "LISTENER" || $type eq "SERVER")} {
        dict set _channels($chan) AcceptCallback $accept_callback
    }
}

proc twapi::tls::_cleanup {chan} {
    debuglog [info level 0]
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

proc twapi::tls::_cleanup_failed_accept {chan} {
    debuglog [info level 0]
    variable _channels
    # This proc is called from the event loop when negotiation fails
    # on a server TLS channel that is not yet open (and hence not
    # known to the application). For some protection against
    # channel name re-use (which does not happen as of 8.6)
    # check the state before cleaning up.
    if {[info exists _channels($chan)] &&
        [dict get $_channels($chan) Type] eq "SERVER" &&
        [dict get $_channels($chan) State] eq "CLOSED"} {
        close $chan;            # Really close
    }
}

if {[llength [info commands ::twapi::tls_background_error]] == 0} {
    proc twapi::tls_background_error {result ropts} {
        return -options $ropts $result
    }
}

proc twapi::tls::_negotiate_from_handler {chan} {
    # Called from socket read / write handlers if
    # negotiation is still in progress.
    # Returns the error code from next step of
    # negotiation.
    # 1 -> ok,
    # 0 -> some error occured, most likely negotiation failure
    variable _channels
    if {[catch {_negotiate $chan} result ropts]} {
        if {![dict exists $_channels($chan) ErrorResult]} {
            dict set _channels($chan) ErrorResult $result
        }
        if {"read" in [dict get $_channels($chan) WatchMask]} {
            _post_read_event $chan
        }
        if {"write" in [dict get $_channels($chan) WatchMask]} {
            _post_write_event $chan
        }
        # For SERVER sockets, force error because no other way
        # to record some error happened.
        if {[dict get $_channels($chan) Type] eq "SERVER"} {
            ::twapi::tls_background_error $result $ropts
            # Above should raise an error, else do it ourselves
            # since stack needs to be rewound
            return -options $ropts $result
        }
        return 0
    }
    return 1
}

proc twapi::tls::_so_read_handler {chan} {
    debuglog [info level 0]
    variable _channels

    if {[info exists _channels($chan)]} {
        if {[dict get $_channels($chan) State] in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            if {![_negotiate_from_handler $chan]} {
                return
            }
        }

        if {"read" in [dict get $_channels($chan) WatchMask]} {
            _post_read_event $chan
        } else {
            # We are not asked to generate read events, turn off the read
            # event handler unless we are negotiating
            if {[dict get $_channels($chan) State] ni {SERVERINIT CLIENTINIT NEGOTIATING}} {
                if {[dict exists $_channels($chan) Socket]} {
                    chan event [dict get $_channels($chan) Socket] readable {}
                }
            }
        }
    }
    return
}

proc twapi::tls::_so_write_handler {chan} {
    debuglog [info level 0]
    variable _channels

    if {[info exists _channels($chan)]} {
        debuglog "[info level 0]: channel exists"
        dict with _channels($chan) {}

        # If we are not actually asked to generate write events,
        # the only time we want a write handler is on a client -async
        # Once it runs, we never want it again else it will keep triggering
        # as sockets are always writable
        if {"write" ni $WatchMask} {
            debuglog "[info level 0]: write not in writemask"
            if {[info exists Socket]} {
                chan event $Socket writable {}
            }
        }

        if {$State in {SERVERINIT CLIENTINIT NEGOTIATING}} {
            debuglog "[info level 0]: Calling _negotiate_from_handler, State=$State"
            if {![_negotiate_from_handler $chan]} {
                # TBD - should we throw so bgerror gets run?
                debuglog "[info level 0]: _negotiate_from_handler returned non-zero."
                return
            }
        }
        debuglog "[info level 0]: State = $State, newstate=[dict get $_channels($chan) State]"
        # Do not use local var $State because _negotiate might have updated it
        if {"write" in $WatchMask && [dict get $_channels($chan) State] eq "OPEN"} {
            debuglog "[info level 0]: posting write event"
            _post_write_event $chan
        } else {
            debuglog "[info level 0]: NOT posting write event"
        }
    }
    debuglog "[info level 0]: returning"
    return
}

proc twapi::tls::_negotiate chan {
    debuglog [info level 0]
    trap {
        _negotiate2 $chan
    } onerror {} {
        variable _channels
        if {[info exists _channels($chan)]} {
            if {[dict get $_channels($chan) Type] eq "SERVER" &&
                [dict get $_channels($chan) State] in {SERVERINIT NEGOTIATING}} {
                # There is no one to clean up accepted sockets (server) that
                # fail verification (or error out) since application does
                # not know about them. So queue some garbage
                # cleaning.
                after 0 [namespace current]::_cleanup_failed_accept $chan
            }
            dict set _channels($chan) State CLOSED
            dict set _channels($chan) ErrorOptions [trapoptions]
            dict set _channels($chan) ErrorResult [trapresult]
            if {[dict exists $_channels($chan) Socket]} {
                catch {close [dict get $_channels($chan) Socket]}
                dict unset _channels($chan) Socket
            }
        }
        rethrow
    }
}

proc twapi::tls::_negotiate2 {chan} {
    variable _channels
        
    dict with _channels($chan) {}; # dict -> local vars

    debuglog "[info level 0]: State=$State"
    switch $State {
        NEGOTIATING {
            if {$Blocking && ![info exists AcceptCallback]} {
                debuglog "[info level 0]: Blocking"
                return [_blocking_negotiate_loop $chan]
            }

            set data [chan read $Socket]
            if {[string length $data] == 0} {
                debuglog "[info level 0]: No data from socket"
                if {[chan eof $Socket]} {
                    debuglog "[info level 0]: EOF on socket"
                    throw {TWAPI TLS NEGOTIATE EOF} "Unexpected EOF during TLS negotiation (NEGOTIATING)"
                } else {
                    # No data yet, just keep waiting
                    debuglog "Waiting (chan $chan) for more data on Socket $Socket"
                    return
                }
            } else {
                debuglog "[info level 0]: Read data from socket"
                lassign [sspi_step $SspiContext $data] status outdata leftover
                debuglog "[info level 0]: sspi_step returned $status"
                debuglog "sspi_step returned status $status with [string length $outdata] bytes"
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
                        debuglog "sspi_step returned $status"
                        error "Unexpected status $status from sspi_step"
                    }
                }
            }
        }

        CLIENTINIT {
            if {$Blocking} {
                debuglog "[info level 0]: CLIENTINIT - blocking negotiate"
                _client_blocking_negotiate $chan
            } else {
                dict set _channels($chan) State NEGOTIATING
                set SspiContext [sspi_client_context $Credentials -usesuppliedcreds $UseSuppliedCreds -stream 1 -target $PeerSubject -manualvalidation [expr {[llength $Verifier] > 0}]]
                dict set _channels($chan) SspiContext $SspiContext
                lassign [sspi_step $SspiContext] status outdata
                debuglog "[info level 0]: sspi_step returned $status"
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
            # For server sockets created from tls_socket, we
            # always take the non-blocking path as we set the socket
            # to be non-blocking so as to not hold up the whole app
            # For server sockets created with starttls 
            # (AcceptCallback will not exist), we can do a blocking
            # negotiate.
            if {$Blocking && ![info exists AcceptCallback]} {
                _server_blocking_negotiate $chan
            } else {
                set data [chan read $Socket]
                if {[string length $data] == 0} {
                    if {[chan eof $Socket]} {
                        throw {TWAPI TLS NEGOTIATE EOF} "Unexpected EOF during TLS negotiation (SERVERINIT)"
                    } else {
                        # No data yet, just keep waiting
                        debuglog "$chan: no data from socket $Socket. Waiting..."
                        return
                    }
                } else {
                    debuglog "Setting $chan State=NEGOTIATING"

                    dict set _channels($chan) State NEGOTIATING
                    set SspiContext [sspi_server_context $Credentials $data -stream 1 -mutualauth $RequestClientCert]
                    dict set _channels($chan) SspiContext $SspiContext
                    lassign [sspi_step $SspiContext] status outdata leftover
                    debuglog "sspi_step returned status $status with [string length $outdata] bytes"
                    if {[string length $outdata]} {
                        debuglog "Writing [string length $outdata] bytes to socket $Socket"
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
                            debuglog "Marking channel $chan open"
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
    debuglog "[info level 0]: returning with state [dict get $_channels($chan) State]"
    return
}

proc twapi::tls::_client_blocking_negotiate {chan} {
    debuglog [info level 0]
    variable _channels
    dict with _channels($chan) {
        set State NEGOTIATING
        set SspiContext [sspi_client_context $Credentials -usesuppliedcreds $UseSuppliedCreds -stream 1 -target $PeerSubject -manualvalidation [expr {[llength $Verifier] > 0}]]
    }
    return [_blocking_negotiate_loop $chan]
}

proc twapi::tls::_server_blocking_negotiate {chan} {
    debuglog [info level 0]
    variable _channels
    dict set _channels($chan) State NEGOTIATING
    set so [dict get $_channels($chan) Socket]
    set indata [_blocking_read $so]
    if {[chan eof $so]} {
        throw {TWAPI TLS NEGOTIATE EOF} "Unexpected EOF during TLS negotiation (server)."
    }
    dict set _channels($chan) SspiContext [sspi_server_context [dict get $_channels($chan) Credentials] $indata -stream 1 -mutualauth [dict get $_channels($chan) RequestClientCert]]
    return [_blocking_negotiate_loop $chan]
}

proc twapi::tls::_blocking_negotiate_loop {chan} {
    debuglog [info level 0]
    variable _channels

    dict with _channels($chan) {}; # dict -> local vars

    lassign [sspi_step $SspiContext] status outdata
    debuglog "sspi_step status $status"
    # Keep looping as long as the SSPI state machine tells us to 
    while {$status eq "continue"} {
        # If the previous step had any output, send it out
        if {[string length $outdata]} {
            debuglog "Writing [string length $outdata] to socket $Socket"
            chan puts -nonewline $Socket $outdata
            chan flush $Socket
        }

        set indata [_blocking_read $Socket]
        debuglog "Read [string length $indata] from socket $Socket"
        if {[chan eof $Socket]} {
            throw {TWAPI TLS NEGOTIATE EOF} "Unexpected EOF during TLS negotiation."
        }
        trap {
            lassign [sspi_step $SspiContext $indata] status outdata leftover
        } onerror {} {
            debuglog "sspi_step returned error: [trapresult]"
            close $Socket
            unset Socket
            rethrow
        }
        debuglog "sspi_step status $status"
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
                error "Error status $status decrypting data"
            }
        }
        _open $chan
    } else {
        # Should not happen. Negotiation failures will raise an error,
        # not return a value
        error "TLS negotiation failed: status $status."
    }

    return
}

proc twapi::tls::_blocking_read {so} {
    debuglog [info level 0]
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

proc twapi::tls::_flush_pending_output {chan} {
    variable _channels

    dict with _channels($chan) {
        if {[string length $Output]} {
            debuglog "_flush_pending_output: flushing output"
            puts -nonewline $Socket [sspi_encrypt_stream $SspiContext $Output]
            set Output ""
	}
    }
    return
}

# Transitions connection to OPEN or throws error if verifier returns false
# or fails
proc twapi::tls::_open {chan} {
    debuglog [info level 0]
    variable _channels

    dict with _channels($chan) {}; # dict -> local vars

    if {[llength $Verifier] == 0} {
        # No verifier specified. In this case, we would not have specified
        # -manualvalidation in creating the context and the system would
        # have done the verification already for client. For servers,
        # there is no verification of clients to be done by default

        # For compatibility with TLS we call accept callbacks AFTER verification
        dict set _channels($chan) State OPEN
        if {[info exists AcceptCallback]} {
            # Server sockets are set up to be non-blocking during negotiation
            # Change them back to original state before notifying app
            chan configure $Socket -blocking [dict get $_channels($chan) Blocking]
            chan event $Socket readable {}
            after 0 $AcceptCallback
        }
        # If there is any pending output waiting for the connection to
        # open, write it out
        _flush_pending_output $chan

        return
    }

    # TBD - what if verifier closes the channel
    if {[{*}$Verifier $chan $SspiContext]} {
        dict set _channels($chan) State OPEN
        # For compatibility with TLS we call accept callbacks AFTER verification
        if {[info exists AcceptCallback]} {
            # Server sockets are set up to be non-blocking during 
            # negotiation. Change them back to original state
            # before notifying app
            chan configure $Socket -blocking [dict get $_channels($chan) Blocking]
            chan event $Socket readable {}
            after 0 $AcceptCallback
        }
        _flush_pending_output $chan
        return
    } else {
        error "SSL/TLS negotiation failed. Verifier callback returned false." "" [list TWAPI TLS VERIFYFAIL]
    }
}

# Calling [chan postevent] results in filevent handlers being called right
# away which can recursively call back into channel code making things
# more than a bit messy. So we always schedule them through the event loop
proc twapi::tls::_post_read_event_callback {chan} {
    debuglog [info level 0]
    variable _channels
    if {[info exists _channels($chan)]} {
        dict unset _channels($chan) ReadEventPosted
        if {"read" in [dict get $_channels($chan) WatchMask]} {
            chan postevent $chan read
        }
    }
}
proc twapi::tls::_post_read_event {chan} {
    debuglog [info level 0]
    variable _channels
    if {![dict exists $_channels($chan) ReadEventPosted]} {
        # Note after 0 after idle does not work - (never get called)
        # not sure why so just do after 0
        dict set _channels($chan) ReadEventPosted \
            [after 0 [namespace current]::_post_read_event_callback $chan]
    }
}
proc twapi::tls::_post_eof_event_callback {chan} {
    debuglog [info level 0]
    variable _channels
    if {[info exists _channels($chan)]} {
        if {"read" in [dict get $_channels($chan) WatchMask]} {
            chan postevent $chan read
        }
    }
}
proc twapi::tls::_post_eof_event {chan} {
    # EOF events are always generated event if a read event is already posted.
    # See Bug #203
    debuglog [info level 0]
    after 0 [namespace current]::_post_eof_event_callback $chan
}


proc twapi::tls::_post_write_event_callback {chan} {
    debuglog [info level 0]
    variable _channels
    if {[info exists _channels($chan)]} {
        dict unset _channels($chan) WriteEventPosted
        if {"write" in [dict get $_channels($chan) WatchMask]} {
            # NOTE: we do not check state here as we should generate an event
            # even in the CLOSED state - see Bug #206
            chan postevent $chan write
        }
    }
}
proc twapi::tls::_post_write_event {chan} {
    debuglog [info level 0]
    variable _channels
    if {![dict exists $_channels($chan) WriteEventPosted]} {
        # Note after 0 after idle does not work - (never get called)
        # not sure why so just do after 0
        dict set _channels($chan) WriteEventPosted \
            [after 0 [namespace current]::_post_write_event_callback $chan]
    }
}

namespace eval twapi::tls {
    namespace ensemble create -subcommands {
        initialize finalize blocking watch read write configure cget cgetall
    }
}

proc twapi::tls::sample_server_creds pfxFile {
    set fd [open $pfxFile rb]
    set pfx [read $fd]
    close $fd
    # Set up the store containing the certificates
    set certStore [twapi::cert_temporary_store -pfx $pfx]
    # Set up the client and server credentials
    set serverCert [twapi::cert_store_find_certificate $certStore subject_substring twapitestserver]
    # TBD - check if certs can be released as soon as we obtain credentials
    set creds [twapi::sspi_acquire_credentials -credentials [twapi::sspi_schannel_credentials -certificates [list $serverCert]] -package unisp -role server]
    twapi::cert_release $serverCert
    twapi::cert_store_release $certStore
    return $creds
}
