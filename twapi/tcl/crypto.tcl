#
# Copyright (c) 2007-2013, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {}

# Maps OID mnemonics
proc twapi::oid {oid} {
    set map {
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

    if {[dict exists $map $oid]} {
        return [dict get $map $oid]
    }
    if {[regexp {^\d([\d\.]*\d)?$} $oid]} {
        return $oid
    } else {
        badargs! "Invalid OID '$oid'"
    }
}

################################################################
# Certificate procs
# TBD - document

# Close a certificate store
proc twapi::cert_store_close {hstore} {
    CertCloseStore $hstore 0
}

proc twapi::cert_memory_store_open {} {
    return [CertOpenStore 2 0 NULL 0 ""]
}

proc twapi::cert_file_store_open {path args} {
    array set opts [parseargs args {
        {readonly        0 0x00008000}
        {commitenable    0 0x00010000}
        {existing        0 0x00004000}
        {create          0 0x00002000}
        {includearchived 0 0x00000200}
        {maxpermissions  0 0x00001000}
        {deferclose      0 0x00000004}
        {backupprivilege 0 0x00000800}
    } -maxleftover 0 -nulldefault]

    if {$opts(readonly) && $opts(commitenable)} {
        badargs! "Options -commitenable and -readonly are mutually exclusive."
    }

    set flags 0
    foreach {opt val} [array get opts] {
        incr flags $val
    }

    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    return [CertOpenStore 8 0x10001 NULL $flags [file nativename [file normalize $path]]]
}

proc twapi::cert_file_store_delete {name} {
    return [CertOpenStore 8 0 NULL 0x10 [file nativename [file normalize $path]]]
}

proc twapi::cert_physical_store_open {name location args} {
    variable _system_stores

    array set opts [parseargs args {
        {readonly        0 0x00008000}
        {commitenable    0 0x00010000}
        {existing        0 0x00004000}
        {create          0 0x00002000}
        {includearchived 0 0x00000200}
        {maxpermissions  0 0x00001000}
        {deferclose      0 0x00000004}
        {backupprivilege 0 0x00000800}
    } -maxleftover 0 -nulldefault]

    if {$opts(readonly) && $opts(commitenable)} {
        badargs! "Options -commitenable and -readonly are mutually exclusive."
    }

    set flags 0
    foreach {opt val} [array get opts] {
        incr flags $val
    }
    incr flags [_system_store_id $location]
    return [CertOpenStore 14 0 NULL $flags $name]
}

proc twapi::cert_physical_store_delete {name} {
    return [CertOpenStore 14 0 NULL 0x10 $name]
}

proc twapi::cert_system_store_open {name args} {
    variable _system_stores

    if {[llength $args] == 0} {
        return [CertOpenSystemStore $name]
    }

    set args [lassign $args location]
    array set opts [parseargs args {
        {readonly        0 0x00008000}
        {commitenable    0 0x00010000}
        {existing        0 0x00004000}
        {create          0 0x00002000}
        {includearchived 0 0x00000200}
        {maxpermissions  0 0x00001000}
        {deferclose      0 0x00000004}
        {backupprivilege 0 0x00000800}
    } -maxleftover 0 -nulldefault]

    if {$opts(readonly) && $opts(commitenable)} {
        badargs! "Options -commitenable and -readonly are mutually exclusive."
    }

    set flags 0
    foreach {opt val} [array get opts] {
        incr flags $val
    }
    incr flags [_system_store_id $location]
    return [CertOpenStore 10 0 NULL $flags $name]
}

proc twapi::cert_system_store_delete {name} {
    return [CertOpenStore 10 0 NULL 0x10 $name]
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

proc twapi::cert_enum_properties {hcert} {
    set id 0
    set ids {}
    while {[set id [CertEnumCertificateContextProperties $hcert $id]]} {
        lappend ids [list $id [_cert_prop_name $id]]
    }
    return $ids
}

proc twapi::cert_get_property {hcert prop} {

    if {[string is integer -strict $prop]} {
        return [CertGetCertificateContextProperty $hcert $prop]
    } 
                                  
    return [switch $prop {
        access_state -
        key_spec {
            binary scan [CertGetCertificateContextProperty $hcert [_cert_prop_id $prop]] i i
            set i
        }
        extended_error_info -
        friendly_name -
        pvk_file {
            string range [encoding convertfrom unicode [CertGetCertificateContextProperty $hcert [_cert_prop_id $prop]]] 0 end-1
        }
        default {
            CertGetCertificateContextProperty $hcert [_cert_prop_id $prop]
        }
    }]
}




################################################################
# Provider contexts

proc twapi::crypt_context_acquire {args} {
    array set opts [parseargs args {
        container.arg
        csp.arg
        {type.arg rsa_full}
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

    return [CryptAcquireContext $opts(container) $opts(csp) [_crypto_provider_type $opts(type)] $flags]
}

proc twapi::crypt_delete_key_container args {
    array set opts [parseargs args {
        container.arg
        csp.arg
        {type.arg rsa_full}
        {storage.arg user {machine user}}
    } -maxleftover 0 -nulldefault]

    set flags 0x10;             # CRYPT_DELETEKEYSET
    if {$opts(storage) eq "machine"} {
        incr flags 0x20;        # CRYPT_MACHINE_KEYSET
    }

    return [CryptAcquireContext $opts(container) $opts(csp) [_crypto_provider_type $opts(type)] $flags]
}

proc twapi::crypt_context_generate_key {hprov algid args} {

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

proc twapi::crypt_context_security_descriptor {hprov args} {
    if {[llength $args] == 1} {
        CryptSetProvParam $hprov 8 [lindex $args 0]
    } elseif {[llength $args] == 0} {
        return [CryptGetProvParam $hprov 8 7]
    } else {
        badargs! "wrong # args: should be \"[lindex [info level 0] 0] hprov ?secd?\""
    }
}

proc twapi::crypt_context_container {hprov} {
    return [_ascii_binary_to_string [CryptGetProvParam $hprov 6 0]]
}

proc twapi::crypt_context_unique_container {hprov} {
    return [_ascii_binary_to_string [CryptGetProvParam $hprov 36 0]]
}

proc twapi::crypt_context_csp {hprov} {
    return [_ascii_binary_to_string [CryptGetProvParam $hprov 4 0]]
}

proc twapi::crypt_context_containers {hprov} {
    return [CryptGetProvParam $hprov 2 0]
}



################################################################
# Utility procs

proc twapi::_algid {class type alg} {
    return [expr {($class << 13) | ($type << 9) | $alg}]
}

twapi::proc* twapi::_cert_prop_id {prop} {
    # Certificate property menomics
    variable _cert_prop_id2name
    array set _cert_prop_id2name {
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
} {
    variable _cert_prop_id2name

    if {[string is integer -strict $prop]} {
        return $prop
    }
    if {![info exists _cert_prop_id2name($prop)]} {
        badargs! "Unknown certificate property id '$prop'" 3
    }

    return $_cert_prop_id2name($prop)
}

twapi::proc* twapi::_cert_prop_name {id} {
    variable _cert_prop_name2id
    variable _cert_prop_id2name

    _cert_prop_id key_prov_handle; # Just to init _cert_prop_id2name
    foreach {k val} [array get _cert_prop_id2name] {
        set _cert_prop_name2id($val) $k
    }
} {
    variable _cert_prop_id2name
    if {[info exists _cert_prop_id2name($id)]} {
        return $_cert_prop_id2name($id)
    }
    if {[string is integer -strict $id]} {
        return $id
    }
    badargs! "Unknown certificate property id '$id'" 3
}

proc twapi::_system_store_id {name} {
    if {[string is integer -strict $name]} {
        if {$name < 65536} {
            badargs! "Invalid system store name $name" 3
        }
        return $name
    }

    set system_stores {
        service          0x40000
        ""               0x10000
        user             0x10000
        usergrouppolicy  0x70000
        localmachine     0x20000
        localmachineenterprise  0x90000
        localmachinegrouppolicy 0x80000
        services 0x50000
        users    0x60000
    }

    if {![dict exists $system_stores $name]} {
        badargs! "Invalid system store name $name" 3
    }

    return [dict get $system_stores $name]
}

proc twapi::_crypto_provider_type prov {
    set provider_names {
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

    set key [string tolower $prov]

    if {[dict exists $provider_names $key]} {
        return [dict get $provider_names $key]
    }

    if {[string is integer -strict $prov]} {
        return $prov
    }

    badargs! "Invalid or unknown provider name '$prov'" 3
}

# If we are being sourced ourselves, then we need to source the remaining files.
if {[file tail [info script]] eq "crypto.tcl"} {
    source [file join [file dirname [info script]] sspi.tcl]
}
