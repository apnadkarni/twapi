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

    # Maps OID mnemonics
    variable _oid_map
    array set _oid_map {
        oid_common_name                   "2.5.4.3"
        oid_sur_name                      "2.5.4.4"
        oid_device_serial_number          "2.5.4.5"
        oid_country_name                  "2.5.4.6"
        oid_locality_name                 "2.5.4.7"
        oid_state_or_province_name        "2.5.4.8"
        oid_street_address                "2.5.4.9"
        oid_organization_name             "2.5.4.10"
        oid_organizational_unit_name      "2.5.4.11"
        oid_title                         "2.5.4.12"
        oid_description                   "2.5.4.13"
        oid_search_guide                  "2.5.4.14"
        oid_business_category             "2.5.4.15"
        oid_postal_address                "2.5.4.16"
        oid_postal_code                   "2.5.4.17"
        oid_post_office_box               "2.5.4.18"
        oid_physical_delivery_office_name "2.5.4.19"
        oid_telephone_number              "2.5.4.20"
        oid_telex_number                  "2.5.4.21"
        oid_teletext_terminal_identifier  "2.5.4.22"
        oid_facsimile_telephone_number    "2.5.4.23"
        oid_x21_address                   "2.5.4.24"
        oid_international_isdn_number     "2.5.4.25"
        oid_registered_address            "2.5.4.26"
        oid_destination_indicator         "2.5.4.27"
        oid_user_password                 "2.5.4.35"
        oid_user_certificate              "2.5.4.36"
        oid_ca_certificate                "2.5.4.37"
        oid_authority_revocation_list     "2.5.4.38"
        oid_certificate_revocation_list   "2.5.4.39"
        oid_cross_certificate_pair        "2.5.4.40"
    }

    variable _provider_names
    array set _provider_names {
        rsa_full           1
        rsa_sig            2
        dss                3
        fortezza           4
        ms_exchange        5
        ssl                6
        rsa_schannel       12
        dss_dh             13
        ec_ecdsa_sig       14
        ec_ecnra_sig       15
        ec_ecdsa_full      16
        ec_ecnra_full      17
        dh_schannel        18
        spyrus_lynks       20
        rng                21
        intel_sec          22
        replace_owf        23
        rsa_aes            24
    }

    # Certificate property menomics
    variable _cert_prop_ids
    array set _cert_prop_ids {
        key_prov_handle        1
        key_prov_info          2
        sha1_hash              3
        hash                   3
        md5_hash               4
        key_context            5
        key_spec               6
        ie30_reserved          7
        pubkey_hash_reserved   8
        enhkey_usage           9
        ctl_usage              9
        next_update_location   10
        friendly_name          11
        pvk_file               12
        description            13
        access_state           14
        signature_hash         15
        smart_card_data        16
        efs                    17
        fortezza_data          18
        archived               19
        key_identifier         20
        auto_enroll            21
        pubkey_alg_para        22
        cross_cert_dist_points 23
        issuer_public_key_md5_hash     24
        subject_public_key_md5_hash    25
        id             26
        date_stamp             27
        issuer_serial_number_md5_hash  28
        subject_name_md5_hash  29
        extended_error_info    30

        renewal                64
        archived_key_hash      65
        auto_enroll_retry      66
        aia_url_retrieved      67
        authority_info_access  68
        backed_up              69
        ocsp_response          70
        request_originator     71
        source_location        72
        source_url             73
        new_key                74
        ocsp_cache_prefix      75
        smart_card_root_info   76
        no_auto_expire_check   77
        ncrypt_key_handle      78
        hcryptprov_or_ncrypt_key_handle   79

        subject_info_access    80
        ca_ocsp_authority_info_access  81
        ca_disable_crl         82
        root_program_cert_policies    83
        root_program_name_constraints 84
        subject_ocsp_authority_info_access  85
        subject_disable_crl    86
        cep                    87

        sign_hash_cng_alg      89

        scard_pin_id           90
        scard_pin_info         91
    }
}

proc twapi::_crypto_provider_type prov {
    variable _provider_names

    set key [string tolower $prov]

    if {[info exists _provider_names($key)]} {
        return $_provider_names($key)
    }

    if {[string is integer -strict $prov]} {
        return $prov
    }

    badargs! "Invalid or unknown provider name '$prov'"
}

proc twapi::oid {oid} {
    variable _oid_map

    if {[info exists _oid_map($oid)]} {
        return $_oid_map($oid)
    }
    
    if {[regexp {^\d([\d\.]*\d)?$} $oid]} {
        return $oid
    } else {
        error "Invalid OID '$oid'"
    }
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

################################################################
# Certificate procs

# Close a certificate store
# TBD - document
proc twapi::cert_close_store {hstore} {
    CertCloseStore $hstore 0
}

proc twapi::cert_find {hstore args} {
    # Currently only subject names supported
    array set opts [parseargs args {
        subject.arg
        {hcert.arg NULL}
    } -maxleftover 0]
                    
    if {![info exists opts(subject)]} {
        error "Option -subject must be specified"
    }

    return [TwapiFindCertBySubjectName $hstore $opts(subject) $opts(hcert)]
}

proc twapi::cert_enum_store {hstore {hcert NULL}} {
    return [CertEnumCertificatesInStore $hstore $hcert]
}

proc twapi::cert_get_name {hcert args} {
    array set opts [parseargs args {
        issuer
        {name.arg oid_common_name}
        {separator.arg comma {comma semi newline}}
        {reverse 0 0x02000000}
        {noquote 0 0x10000000}
        {noplus  0 0x20000000}
        {format.arg x500 {x500 oid simple}}
    } -maxleftover 0]

    set arg ""
    switch $opts(name) {
        email { set what 1 }
        simpledisplay { set what 4 }
        friendlydisplay {set what 5 }
        dns { set what 6 }
        url { set what 7 }
        upn { set what 8 }
        rdn {
            set what 2
            switch $opts(format) {
                simple {set arg 1}
                oid {set arg 2}
                x500 -
                default {set arg 3}
            }
            set arg [expr {$arg | $opts(reverse) | $opts(noquote) | $opts(noplus)}]
            switch $opts(separator) {
                semi    { set arg [expr {$arg | 0x40000000}] }
                newline { set arg [expr {$arg | 0x08000000}] }
            }
        }
        default {
            set what 3;         # Assume OID
            set arg [oid $opts(name)]
        }
    }

    return [CertGetNameString $hcert $what $opts(issuer) $arg]

}

proc twapi::cert_blob_to_name {blob args} {
    array set opts [parseargs args {
        {format.arg x500 {x500 oid simple}}
        {separator.arg comma {comma semi newline}}
        {reverse 0 0x02000000}
        {noquote 0 0x10000000}
        {noplus  0 0x20000000}
    } -maxleftover 0]

    switch $opts(format) {
        x500   {set arg 3}
        simple {set arg 1}
        oid    {set arg 2}
    }

    set arg [expr {$arg | $opts(reverse) | $opts(noquote) | $opts(noplus)}]
    switch $opts(separator) {
        semi    { set arg [expr {$arg | 0x40000000}] }
        newline { set arg [expr {$arg | 0x08000000}] }
    }

    return [CertNameToStr $blob $arg]
}

proc twapi::cert_name_to_blob {name args} {
    array set opts [parseargs args {
        {format.arg x500 {x500 oid simple}}
        {separator.arg comma {comma semi newline}}
        {reverse 0 0x02000000}
        {noquote 0 0x10000000}
        {noplus  0 0x20000000}
    } -maxleftover 0]

    switch $opts(format) {
        x500   {set arg 3}
        simple {set arg 1}
        oid    {set arg 2}
    }

    set arg [expr {$arg | $opts(reverse) | $opts(noquote) | $opts(noplus)}]
    switch $opts(separator) {
        comma   { set arg [expr {$arg | 0x04000000}] }
        semi    { set arg [expr {$arg | 0x40000000}] }
        newline { set arg [expr {$arg | 0x08000000}] }
    }

    return [CertStrToName $name $arg]
}

proc twapi::crypt_acquire_context {args} {
    array set opts [parseargs args {
        container.arg
        provider.arg
        {providertype.arg rsa_full}
        {storage.arg user {machine user}}
        {create 0 0x8}
        {silent 0 0x40}
        {verifycontext 0 0xf0000000}
    } -maxleftover 0 -nulldefault]
    
    # Based on http://support.microsoft.com/kb/238187    
    if {$opts(verifycontext) && $opts(container) eq ""} {
        badargs! "Option -verifycontext must be specified if -container option is unspecified or empty"
    }

    set flags [expr {$opts(create) | $opts(silent) | $opts(verifycontext)}]
    if {$opts(storage) eq "machine"} {
        incr flags 0x20;        # CRYPT_MACHINE_KEYSET
    }

    return [CryptAcquireContext $opts(container) $opts(provider) [_crypto_provider_type $opts(providertype) $flags]]
}

proc twapi::crypt_delete_key_container args {
    array set opts [parseargs args {
        container.arg
        provider.arg
        {providertype.arg rsa_full}
        {storage.arg user {machine user}}
    } -maxleftover 0 -nulldefault]

    set flags 0x10;             # CRYPT_DELETEKEYSET
    if {$opts(storage) eq "machine"} {
        incr flags 0x20;        # CRYPT_MACHINE_KEYSET
    }

    return [CryptAcquireContext $opts(container) $opts(provider) [_crypto_provider_type $opts(providertype) $flags]]
}

proc twapi::crypt_generate_key {hprov algid args} {

    array set opts [parseargs args {
        {archivable 0 0x4000}
        {salt 0 4}
        {exportable 0 1}
        {pregen 0x40}
        {userprotected 0 2}
        {nosalt40 0 0x10}
        {size 0}
    } -maxleftover 0]

    if {![string is integer -strict $algid]} {
        # See wincrypt.h in SDK
        switch -nocase -exact -- $algid {
            keyexhange {set algid 1}
            signature {set algid 2}
            dh_ephemeral {set algid [_algid 5 5 2]}
            dh_sandf {set algid [_algid 5 5 1]}
            md5 {set algid [_algid 4 0 3]}
            sha -
            sha1 {set algid [_algid 4 0 4]}
            mac {set algid [_algid 4 0 5]}
            ripemd {set algid [_algid 4 0 6]}
            rimemd160 {set algid [_algid 4 0 7]}
            ssl3_shamd5 {set algid [_algid 4 0 8]}
            hmac {set algid [_algid 4 0 9]}
            sha256 {set algid [_algid 4 0 12]}
            sha384 {set algid [_algid 4 0 13]}
            sha512 {set algid [_algid 4 0 14]}
            rsa_sign {set algid [_algid 1 2 0]}
            dss_sign {set algid [_algid 1 1 0]}
            rsa_keyx {set algid [_algid 5 2 0]}
            des {set algid [_algid 3 3 1]}
            3des {set algid [_algid 3 3 3]}
            3des112 {set algid [_algid 3 3 9]}
            desx {set algid [_algid 3 3 4]}
            default {badargs! "Invalid value '$algid' for parameter algid"}
        }
    }

    if {$opts(size) < 0 || $opts(size) > 65535} {
        badargs! "Bad key size parameter '$size':  must be positive integer less than 65536"
    }

    return [CryptGenKey $hprov $algid [expr {($opts(size) << 16) | $opts(archivable) | $opts(salt) | $opts(exportable) | $opts(pregen) | $opts(userprotected) | $opts(nosalt40)}]]
}

proc twapi::cert_enum_context_properties {hcert} {
    variable _cert_prop_ids

    foreach {k val} [array get _cert_prop_ids] {
        set reverse($val) $k
    }

    set id 0
    set ids {}
    while {[set id [CertEnumCertificateContextProperties $hcert $id]]} {
        if {[info exists reverse($id)]} {
            lappend ids [list $id $reverse($id)]
        } else {
            lappend ids [list $id $id]
        }
    }
    return $ids
}

proc twapi::cert_get_context_property {hcert prop} {
    variable _cert_prop_ids

    if {[string is integer -strict $prop]} {
        return [CertGetCertificateContextProperty $hcert $prop]
    } 
                                  
    return [switch $prop {
        access_state -
        key_spec {
            binary scan [CertGetCertificateContextProperty $hcert $_cert_prop_ids($prop)] i i
            set i
        }
        extended_error_info -
        friendly_name -
        pvk_file {
            string range [encoding convertfrom unicode [CertGetCertificateContextProperty $hcert $_cert_prop_ids($prop)]] 0 end-1
        }
        default {
            CertGetCertificateContextProperty $hcert $_cert_prop_ids($prop)
        }
    }]
}


################################################################
# Utility procs

proc twapi::_algid {class type alg} {
    return [expr {($class << 13) | ($type << 9) | $alg}]
}

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

