#
# Copyright (c) 2007, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {

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
}

# Return list of security packages
proc twapi::sspi_enumerate_packages {} {
    set packages [list ]
    foreach pkg [EnumerateSecurityPackages] {
        lappend packages [kl_get $pkg Name]
    }
    return $packages
}


# Return sspi credentials
proc twapi::sspi_new_credentials {args} {
    array set opts [parseargs args {
        {principal.arg ""}
        {package.arg NTLM}
        {usage.arg both {inbound outbound both}}
        getexpiration
        user.arg
        {domain.arg ""}
        {password.arg ""}
    } -maxleftover 0]

    # Do not want error stack to include password but how to scrub it? - TBD
    if {[info exists opts(user)]} {
        set auth [Twapi_Allocate_SEC_WINNT_AUTH_IDENTITY $opts(user) $opts(domain) $opts(password)]
    } else {
        set auth NULL
    }

    trap {
        set creds [AcquireCredentialsHandle $opts(principal) $opts(package) \
                       [kl_get {inbound 1 outbound 2 both 3} $opts(usage)] \
                       "" $auth]
    } finally {
        Twapi_Free_SEC_WINNT_AUTH_IDENTITY $auth; # OK if NULL
    }
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
    foreach {opt flag} [array get ::twapi::_client_security_context_syms] {
        if {$opts($opt)} {
            set context_flags [expr {$context_flags | $flag}]
        }
    }

    set drep [kl_get {native 0x10 network 0} $opts(datarep)]
    return [_construct_sspi_security_context \
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
    DeleteSecurityContext [kl_get $ctx -handle]
}

# Take the next step in client side authentication
# Returns
#   {done data newctx}
#   {continue data newctx}
proc twapi::sspi_security_context_next {ctx {response ""}} {
    switch -exact -- [kl_get $ctx -state] {
        ok {
            # Should not be passed remote response data in this state
            if {[string length $response]} {
                error "Unexpected remote response data passed."
            }
            # See if there is any data to send.
            set data ""
            foreach buf [kl_get $ctx -output] {
                append data [lindex $buf 1]
            }
            return [list done $data [kl_set $ctx -output [list ]]]
        }
        continue {
            # See if there is any data to send.
            set data ""
            foreach buf [kl_get $ctx -output] {
                append data [lindex $buf 1]
            }
            # Either we are receiving response from the remote system
            # or we have to send it data. Both cannot be empty
            if {[string length $response] != 0} {
                # We are given a response. Pass it back in
                # to SSPI.
                # "2" buffer type is SECBUFFER_TOKEN
                set inbuflist [list [list 2 $response]]
                if {[kl_get $ctx -type] eq "client"} {
                    set rawctx [InitializeSecurityContext \
                                    [kl_get $ctx -credentials] \
                                    [kl_get $ctx -handle] \
                                    [kl_get $ctx -target] \
                                    [kl_get $ctx -inattr] \
                                    0 \
                                    [kl_get $ctx -datarep] \
                                    $inbuflist \
                                    0]
                } else {
                    set rawctx [AcceptSecurityContext \
                                    [kl_get $ctx -credentials] \
                                    [kl_get $ctx -handle] \
                                    $inbuflist \
                                    [kl_get $ctx -inattr] \
                                    [kl_get $ctx -datarep] \
                                   ]
                }
                set newctx [_construct_sspi_security_context \
                                $rawctx \
                                [kl_get $ctx -type] \
                                [kl_get $ctx -inattr] \
                                [kl_get $ctx -target] \
                                [kl_get $ctx -credentials] \
                                [kl_get $ctx -datarep] \
                               ]
                # Recurse to return next state
                return [sspi_security_context_next $newctx]
            } elseif {[string length $data] != 0} {
                # We have to send data to the remote system and await its
                # response. Reset output buffers to empty
                return [list continue $data [kl_set $ctx -output [list ]]]
            } else {
                error "No token data available to send to remote system"
            }
        }
        complete -
        complete_and_continue -
        incomplete_message {
            # TBD
            error "State '[kl_get $ctx -state]' handling not implemented."
        }
    }
}

# Return a server context
proc ::twapi::sspi_server_new_context {cred clientdata args} {
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
    foreach {opt flag} [array get ::twapi::_server_security_context_syms] {
        if {$opts($opt)} {
            set context_flags [expr {$context_flags | $flag}]
        }
    }

    set drep [kl_get {native 0x10 network 0} $opts(datarep)]
    return [_construct_sspi_security_context \
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
    # We could directly look in the context itself but intead we make
    # an explicit call, just in case they change after initial setup
    set flags [QueryContextAttributes [kl_get $ctx -handle] 14]

    # Mapping of symbols depends on whether it is a client or server
    # context
    if {[kl_get $ctx -type] eq "client"} {
        upvar 0 ::twapi::_client_security_context_syms syms
    } else {
        upvar 0 ::twapi::_server_security_context_syms syms
    }

    set result [list -raw $flags]
    foreach {sym flag} [array get syms] {
        lappend result -$sym [expr {($flag & $flags) != 0}]
    }

    return $result
}

# Get the user name for a security context
proc twapi::sspi_get_security_context_username {ctx} {
    return [QueryContextAttributes [kl_get $ctx -handle] 1]
}

# Get the field size information for a security context
proc twapi::sspi_get_security_context_sizes {ctx} {
    if {![kl_vget $ctx -sizes sizes]} {
        set sizes [QueryContextAttributes [kl_get $ctx -handle] 0]
    }

    return [kl_create2 {-maxtoken -maxsig -blocksize -trailersize} $sizes]
}

# Returns a signature
proc twapi::sspi_generate_signature {ctx data args} {
    array set opts [parseargs args {
        {seqnum.int 0}
        {qop.int 0}
    } -maxleftover 0]

    return [MakeSignature \
                [kl_get $ctx -handle] \
                $opts(qop) \
                $data \
                $opts(seqnum)]
}

# Verify signature
proc twapi::sspi_verify_signature {ctx sig data args} {
    array set opts [parseargs args {
        {seqnum.int 0}
    } -maxleftover 0]

    # Buffer type 2 - Token, 1- Data
    return [VerifySignature \
                [kl_get $ctx -handle] \
                [list [list 2 $sig] [list 1 $data]] \
                $opts(seqnum)]
}

# Encrypts a data as per a context
# Returns {securitytrailer encrypteddata padding}
proc twapi::sspi_encrypt {ctx data args} {
    array set opts [parseargs args {
        {seqnum.int 0}
        {qop.int 0}
    } -maxleftover 0]

    return [EncryptMessage \
                [kl_get $ctx -handle] \
                $opts(qop) \
                $data \
                $opts(seqnum)]
}

# Decrypts a message
proc twapi::sspi_decrypt {ctx sig data padding args} {
    array set opts [parseargs args {
        {seqnum.int 0}
    } -maxleftover 0]

    # Buffer type 2 - Token, 1- Data, 9 - padding
    set decrypted [DecryptMessage \
                       [kl_get $ctx -handle] \
                       [list [list 2 $sig] [list 1 $data] [list 9 $padding]] \
                       $opts(seqnum)]
    set plaintext ""
    # Pick out only the data buffers, ignoring pad buffers and signature
    foreach buf $decrypted {
        if {[lindex $buf 0] == 1} {
            append plaintext [lindex $buf 1]
        }
    }
    return $plaintext
}

# Utility procs

# Construct a high level SSPI security context structure
# ctx is context as returned from C level code
proc twapi::_construct_sspi_security_context {ctx ctxtype inattr target credentials datarep} {
    set result [kl_create2 \
                    {-state -handle -output -outattr -expiration} \
                    $ctx]
    set result [kl_set $result -type $ctxtype]
    set result [kl_set $result -inattr $inattr]
    set result [kl_set $result -target $target]
    set result [kl_set $result -datarep $datarep]
    return [kl_set $result -credentials $credentials]

}



proc twapi::_sspi_sample {} {
    set ccred [sspi_new_credentials -usage outbound]
    set scred [sspi_new_credentials -usage inbound]
    set cctx [sspi_client_new_context $ccred -target LUNA -confidentiality true -connection true]
    lassign  [sspi_security_context_next $cctx] step data cctx
    set sctx [sspi_server_new_context $scred $data]
    lassign  [sspi_security_context_next $sctx] step data sctx
    lassign  [sspi_security_context_next $cctx $data] step data cctx
    lassign  [sspi_security_context_next $sctx $data] step data sctx
    sspi_free_credentials $scred
    sspi_free_credentials $ccred
    return [list $cctx $sctx]
}

