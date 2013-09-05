#
# Copyright (c) 2007-2013, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

if {0} {
TBD - from curl -
schannel: Removed extended error connection setup flag
According KB975858 this flag may cause problems on Windows 7 and
Windows Server 2008 R2 systems. Extended error information is not
currently used by libcurl and therefore not a requirement.

}

namespace eval twapi {

    # Holds SSPI security context handles
    variable _sspi_state
    array set _sspi_state {}

    proc* _init_security_context_syms {} {
        variable _server_security_context_syms
        variable _client_security_context_syms
        variable _secpkg_capability_syms

        # Symbols used for mapping server security context flags
        array set _server_security_context_syms {
            confidentiality      0x10
            connection           0x800
            delegate             0x1
            extendederror        0x8000
            integrity            0x20000
            mutualauth           0x2
            replaydetect         0x4
            sequencedetect       0x8
            stream               0x10000
        }

        # Symbols used for mapping client security context flags
        array set _client_security_context_syms {
            confidentiality      0x10
            connection           0x800
            delegate             0x1
            extendederror        0x4000
            integrity            0x10000
            manualvalidation     0x80000
            mutualauth           0x2
            replaydetect         0x4
            sequencedetect       0x8
            stream               0x8000
            usesessionkey        0x20
            usesuppliedcreds     0x80
        }

        # Symbols used for mapping security package capabilities
        array set _secpkg_capability_syms {
            integrity                   0x00000001
            privacy                     0x00000002
            tokenonly                  0x00000004
            datagram                    0x00000008
            connection                  0x00000010
            multirequired              0x00000020
            clientonly                 0x00000040
            extendederror              0x00000080
            impersonation               0x00000100
            acceptwin32name           0x00000200
            stream                      0x00000400
            negotiable                  0x00000800
            gsscompatible              0x00001000
            logon                       0x00002000
            asciibuffers               0x00004000
            fragment                    0x00008000
            mutualauth                 0x00010000
            delegation                  0x00020000
            readonlywithchecksum      0x00040000
            restrictedtokens           0x00080000
            negoextender               0x00100000
            negotiable2                 0x00200000
            appcontainerpassthrough  0x00400000
            appcontainerchecks  0x00800000
        }
    } {}
}

# Return list of security packages
proc twapi::sspi_enumerate_packages {args} {
    set pkgs [EnumerateSecurityPackages]
    if {[llength $args] == 0} {
        set names [list ]
        foreach pkg $pkgs {
            lappend names [kl_get $pkg Name]
        }
        return $names
    }

    array set opts [parseargs args {
        all capabilities version rpcid maxtokensize name comment
    } -maxleftover 0 -hyphenated]

    _init_security_context_syms
    variable _secpkg_capability_syms
    set retdata {}
    foreach pkg $pkgs {
        set rec {}
        if {$opts(-all) || $opts(-capabilities)} {
            lappend rec -capabilities [_make_symbolic_bitmask [kl_get $pkg fCapabilities] _secpkg_capability_syms]
        }
        foreach {opt field} {-version wVersion -rpcid wRPCID -maxtokensize cbMaxToken -name Name -comment Comment} {
            if {$opts(-all) || $opts($opt)} {
                lappend rec $opt [kl_get $pkg $field]
            }
        }
        dict set recdata [kl_get $pkg Name] $rec
    }
    return $recdata
}


# Return sspi credentials
proc twapi::sspi_new_credentials {args} {
    array set opts [parseargs args {
        {principal.arg ""}
        {package.arg NTLM}
        {usage.arg both {inbound outbound both}}
        getexpiration
    } -maxleftover 0]

    set creds [AcquireCredentialsHandle $opts(principal) $opts(package) \
                   [kl_get {inbound 1 outbound 2 both 3} $opts(usage)] \
                   "" {}]

    if {$opts(getexpiration)} {
        return [kl_create2 {-handle -expiration} $creds]
    } else {
        return [lindex $creds 0]
    }
}

# Frees credentials
proc twapi::sspi_free_credentials {cred} {
    FreeCredentialsHandle $cred
}

# Return a client context
proc ::twapi::sspi_client_new_context {cred args} {
    _init_security_context_syms
    variable _client_security_context_syms

    array set opts [parseargs args {
        target.arg
        {datarep.arg network {native network}}
        confidentiality.bool
        connection.bool
        delegate.bool
        extendederror.bool
        integrity.bool
        manualvalidation.bool
        mutualauth.bool
        replaydetect.bool
        sequencedetect.bool
        stream.bool
        usesessionkey.bool
        usesuppliedcreds.bool
    } -maxleftover 0 -nulldefault]

    set context_flags 0
    foreach {opt flag} [array get _client_security_context_syms] {
        if {$opts($opt)} {
            set context_flags [expr {$context_flags | $flag}]
        }
    }

    set drep [kl_get {native 0x10 network 0} $opts(datarep)]
    return [_construct_sspi_security_context \
                sspiclient#[TwapiId] \
                [InitializeSecurityContext \
                     $cred \
                     "" \
                     $opts(target) \
                     $context_flags \
                     0 \
                     $drep \
                     [list ] \
                     0] \
                client \
                $context_flags \
                $opts(target) \
                $cred \
                $drep \
               ]
}

# Delete a security context
proc twapi::sspi_close_security_context {ctx} {
    variable _sspi_state

    _sspi_validate_handle $ctx
    
    DeleteSecurityContext [dict get $_sspi_state($ctx) Handle]
    unset _sspi_state($ctx)
}

# Take the next step in client side authentication
# Returns
#   {done data newctx}
#   {continue data newctx}
proc twapi::sspi_security_context_next {ctx {response ""}} {
    variable _sspi_state

    _sspi_validate_handle $ctx

    dict with _sspi_state($ctx) {
        # Note the dictionary content variables are
        #   State, Handle, Output, Outattr, Expiration,
        #   Ctxtype, Inattr, Target, Datarep, Credentials
        switch -exact -- $State {
            ok {
                # Should not be passed remote response data in this state
                if {[string length $response]} {
                    error "Unexpected remote response data passed."
                }
                # See if there is any data to send.
                set data ""
                foreach buf $Output {
                    # First element is buffer type, which we do not care
                    # Second element is actual data
                    append data [lindex $buf 1]
                }

                set Output {}
                # We return the context handle as third element for backwards
                # compatibility reasons - TBD
                return [list done $data $ctx]
            }
            continue {
                # See if there is any data to send.
                set data ""
                foreach buf $Output {
                    append data [lindex $buf 1]
                }
                # Either we are receiving response from the remote system
                # or we have to send it data. Both cannot be empty
                if {[string length $response] != 0} {
                    # We are given a response. Pass it back in
                    # to SSPI.
                    # "2" buffer type is SECBUFFER_TOKEN
                    set inbuflist [list [list 2 $response]]
                    if {$Ctxtype eq "client"} {
                        set rawctx [InitializeSecurityContext \
                                        $Credentials \
                                        $Handle \
                                        $Target \
                                        $Inattr \
                                        0 \
                                        $Datarep \
                                        $inbuflist \
                                        0]
                    } else {
                        set rawctx [AcceptSecurityContext \
                                        $Credentials \
                                        $Handle \
                                        $inbuflist \
                                        $Inattr \
                                        $Datarep]
                    }
                    lassign $rawctx State Handle Output Outattr Expiration
                    # Will recurse at proc end
                } elseif {[string length $data] != 0} {
                    # We have to send data to the remote system and await its
                    # response. Reset output buffers to empty
                    set Output {}
                    return [list continue $data $ctx]
                } else {
                    # TBD - is this really an error ?
                    error "No token data available to send to remote system (SSPI context $ctx)"
                }
            }
            complete -
            complete_and_continue -
            incomplete_message {
                # TBD
                error "State $State handling not implemented."
            }
        }
    }

    # Recurse to return next state
    # Note this has to be OUTSIDE the [dict with] else it will not
    # see the updated values
    return [sspi_security_context_next $ctx]
}

# Return a server context
proc ::twapi::sspi_server_new_context {cred clientdata args} {
    _init_security_context_syms
    variable _server_security_context_syms

    array set opts [parseargs args {
        {datarep.arg network {native network}}
        confidentiality.bool
        connection.bool
        delegate.bool
        extendederror.bool
        integrity.bool
        mutualauth.bool
        replaydetect.bool
        sequencedetect.bool
        stream.bool
    } -maxleftover 0 -nulldefault]

    set context_flags 0
    foreach {opt flag} [array get _server_security_context_syms] {
        if {$opts($opt)} {
            set context_flags [expr {$context_flags | $flag}]
        }
    }

    set drep [kl_get {native 0x10 network 0} $opts(datarep)]
    return [_construct_sspi_security_context \
                sspiserver#[TwapiId] \
                [AcceptSecurityContext \
                     $cred \
                     "" \
                     [list [list 2 $clientdata]] \
                     $context_flags \
                     $drep] \
                server \
                $context_flags \
                "" \
                $cred \
                $drep \
               ]
}


# Get the security context flags after completion of request
proc ::twapi::sspi_get_security_context_features {ctx} {
    variable _sspi_state

    _sspi_validate_handle $ctx

    _init_security_context_syms

    # We could directly look in the context itself but intead we make
    # an explicit call, just in case they change after initial setup
    set flags [QueryContextAttributes [dict get $_sspi_state($ctx) Handle] 14]

        # Mapping of symbols depends on whether it is a client or server
        # context
    if {[dict get $_sspi_state($ctx) Ctxtype] eq "client"} {
        upvar 0 [namespace current]::_client_security_context_syms syms
    } else {
        upvar 0 [namespace current]::_server_security_context_syms syms
    }

    set result [list -raw $flags]
    foreach {sym flag} [array get syms] {
        lappend result -$sym [expr {($flag & $flags) != 0}]
    }

    return $result
}

# Get the user name for a security context
proc twapi::sspi_get_security_context_username {ctx} {
    variable _sspi_state
    _sspi_validate_handle $ctx
    return [QueryContextAttributes [dict get $_sspi_state($ctx) Handle] 1]
}

# Get the field size information for a security context
# TBD - update for SSL
proc twapi::sspi_get_security_context_sizes {ctx} {
    variable _sspi_state
    _sspi_validate_handle $ctx

    set sizes [QueryContextAttributes [dict get $_sspi_state($ctx) Handle] 0]
    return [twine {-maxtoken -maxsig -blocksize -trailersize} $sizes]
}

# Returns a signature
proc twapi::sspi_generate_signature {ctx data args} {
    variable _sspi_state
    _sspi_validate_handle $ctx

    array set opts [parseargs args {
        {seqnum.int 0}
        {qop.int 0}
    } -maxleftover 0]

    return [MakeSignature \
                [dict get $_sspi_state($ctx) Handle] \
                $opts(qop) \
                $data \
                $opts(seqnum)]
}

# Verify signature
proc twapi::sspi_verify_signature {ctx sig data args} {
    variable _sspi_state
    _sspi_validate_handle $ctx

    array set opts [parseargs args {
        {seqnum.int 0}
    } -maxleftover 0]

    # Buffer type 2 - Token, 1- Data
    return [VerifySignature \
                [dict get $_sspi_state($ctx) Handle] \
                [list [list 2 $sig] [list 1 $data]] \
                $opts(seqnum)]
}

# Encrypts a data as per a context
# Returns {securitytrailer encrypteddata padding}
proc twapi::sspi_encrypt {ctx data args} {
    variable _sspi_state
    _sspi_validate_handle $ctx

    array set opts [parseargs args {
        {seqnum.int 0}
        {qop.int 0}
    } -maxleftover 0]

    return [EncryptMessage \
                [dict get $_sspi_state($ctx) Handle] \
                $opts(qop) \
                $data \
                $opts(seqnum)]
}

# Decrypts a message
proc twapi::sspi_decrypt {ctx sig data padding args} {
    variable _sspi_state
    _sspi_validate_handle $ctx

    array set opts [parseargs args {
        {seqnum.int 0}
    } -maxleftover 0]

    # Buffer type 2 - Token, 1- Data, 9 - padding
    set decrypted [DecryptMessage \
                       [dict get $_sspi_state($ctx) Handle] \
                       [list [list 2 $sig] [list 1 $data] [list 9 $padding]] \
                       $opts(seqnum)]
    set plaintext {}
    # Pick out only the data buffers, ignoring pad buffers and signature
    # Optimize copies by keeping as a list so in the common case of a 
    # single buffer can return it as is. Multiple buffers are expensive
    # because Tcl will shimmer each byte array into a list and then
    # incur additional copies during joining
    foreach buf $decrypted {
        # SECBUFFER_DATA -> 1
        if {[lindex $buf 0] == 1} {
            lappend plaintext [lindex $buf 1]
        }
    }

    if {[llength $plaintext] < 2} {
        return [lindex $plaintext 0]
    } else {
        return [join $plaintext ""]
    }
}


################################################################
# Utility procs


# Construct a high level SSPI security context structure
# rawctx is context as returned from C level code
proc twapi::_construct_sspi_security_context {id rawctx ctxtype inattr target credentials datarep} {
    variable _sspi_state
    
    set _sspi_state($id) [twine \
                                {State Handle Output Outattr Expiration} \
                                $rawctx]
    dict set _sspi_state($id) Ctxtype $ctxtype
    dict set _sspi_state($id) Inattr $inattr
    dict set _sspi_state($id) Target $target
    dict set _sspi_state($id) Datarep $datarep
    dict set _sspi_state($id) Credentials $credentials

    return $id
}

proc twapi::_sspi_validate_handle {ctx} {
    variable _sspi_state

    if {![info exists _sspi_state($ctx)]} {
        badargs! "Invalid SSPI security context handle $ctx" 3
    }
}
