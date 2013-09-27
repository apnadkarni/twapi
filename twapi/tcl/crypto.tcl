#
# Copyright (c) 2007-2013, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {}


if {0} {
    TBD - associating private key with cert
    from http://stackoverflow.com/questions/749654/associate-private-key-to-certificate-for-pfxexportcertstoreex:

Apparently, CertSetCertificateContextProperty(p, CERT_KEY_PROV_HANDLE_PROP_ID ...)

is not good. Need to do this instead:

CRYPT_KEY_PROV_INFO kpi;
ZeroMemory( & kpi, sizeof(kpi) );
kpi.pwszContainerName = "my-container-name";
kpi.dwProvType = PROV_RSA_FULL;
kpi.dwKeySpec = AT_KEYEXCHANGE;
kpi.dwFlags = CRYPT_MACHINE_KEYSET;
CertSetCertificateContextProperty( pCert, CERT_KEY_PROV_INFO_PROP_ID, 0, & kpi);
It's critical that provider name and other crap match the information that was used to generate actual key. It's not needed to set provider handle or any of that stuff. It also must be done before CertAddCertificateContextToStore.

This is the only way that I found to attach private key to a certificate.
}

################################################################
# Certificate Stores

# Close a certificate store
proc twapi::cert_store_close {hstore} {
    CertCloseStore $hstore 0
}

proc twapi::cert_memory_store_open {} {
    return [CertOpenStore 2 0 NULL 0 ""]
}

proc twapi::cert_file_store_open {path args} {
    array set opts [parseargs args {
        {readonly.bool        0 0x00008000}
        {commitenable.bool    0 0x00010000}
        {existing.bool        0 0x00004000}
        {create.bool          0 0x00002000}
        {includearchived.bool 0 0x00000200}
        {maxpermissions.bool  0 0x00001000}
        {deferclose.bool      0 0x00000004}
        {backupprivilege.bool 0 0x00000800}
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

proc twapi::cert_physical_store_open {name location args} {
    variable _system_stores

    array set opts [parseargs args {
        {readonly.bool        0 0x00008000}
        {commitenable.bool    0 0x00010000}
        {existing.bool        0 0x00004000}
        {create.bool          0 0x00002000}
        {includearchived.bool 0 0x00000200}
        {maxpermissions.bool  0 0x00001000}
        {deferclose.bool      0 0x00000004}
        {backupprivilege.bool 0 0x00000800}
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

proc twapi::cert_physical_store_delete {name location} {
    set flags 0x10;             # CERT_STORE_DELETE_FLAG
    incr flags [_system_store_id $location]
    
    return [CertOpenStore 14 0 NULL $flags $name]
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

proc twapi::cert_system_store_delete {name location} {
    set flags 0x10;             # CERT_STORE_DELETE_FLAG
    incr flags [_system_store_id $location]
    return [CertOpenStore 10 0 NULL $flags $name]
}

proc twapi::cert_store_find_certificate {hstore {type any} {term {}} {hcert NULL}} {

    set term_types {
        any 0
        existing 13<<16
        subject_cert 11<<16
        key_identifier 15<<16
        md5_hash 4<<16
        pubkey_md5_hash 18<<16
        sha1_hash 1<<16
        signature_hash 14<<16
        issuer_name (2<<16)|4
        subject_name  (2<<16)|7
        issuer_substring (8<<16)|4
        subject_substring (8<<16)|7
        property 5<<16
        key_spec 9<<16
        public_key 6<<16
    }

    set type [expr [dict! $term_types $type 1]]

    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    return [CertFindCertificateInStore $hstore 0x10001 0 $type $term $hcert]
}

proc twapi::cert_store_enum_contents {hstore {hcert NULL}} {
    return [CertEnumCertificatesInStore $hstore $hcert]
}

proc twapi::cert_store_add_certificate {hstore hcert args} {
    array set opts [_cert_add_parseargs args]
    return [CertAddCertificateContextToStore $hstore $hcert $opts(disposition)]
}

proc twapi::cert_store_add_encoded_certificate {hstore enccert args} {
    array set opts [_cert_add_parseargs args]
    return [CertAddEncodedCertificateToStore $hstore 0x10001 $enccert $opts(disposition)]
}

proc twapi::cert_store_export {hstore password args} {
    array set opts [parseargs args {
        {exportprivatekeys.bool 0 0x4}
        {failonmissingkey.bool 0 0x1}
        {failonunexportablekey.bool 0 0x2}
    } -maxleftover 0]

    set flags [tcl::mathop::| $opts(exportprivatekeys) $opts(failonunexportablekey) $opts(failonmissingkey)]

    return [::twapi::PFXExportCertStoreEx $hstore $password {} $flags]
}

proc twapi::cert_store_commit {hstore args} {
    array set opts [parseargs args {
        {force.bool 0}
    } -maxleftover 0]
    
    return [Twapi_CertStoreCommit $hstore $opts(force)]
}

################################################################
# Certificates


interp alias {} twapi::cert_subject_name {} twapi::_cert_get_name subject
interp alias {} twapi::cert_issuer_name {} twapi::_cert_get_name issuer
proc twapi::_cert_get_name {field hcert args} {

    switch $field {
        subject { set field 0 }
        issuer  { set field 1 }
        default { badargs! "Invalid name type '$field': must be \"subject\" or \"issuer\"."
        }
    }
    array set opts [parseargs args {
        {name.arg oid_common_name}
        {separator.arg comma {comma semicolon newline}}
        {reverse.bool 0 0x02000000}
        {noquote.bool 0 0x10000000}
        {noplus.bool  0 0x20000000}
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
                semicolon    { set arg [expr {$arg | 0x40000000}] }
                newline { set arg [expr {$arg | 0x08000000}] }
            }
        }
        default {
            set what 3;         # Assume OID
            set arg [oid $opts(name)]
        }
    }

    return [CertGetNameString $hcert $what $field $arg]

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
        {separator.arg any {any comma semicolon newline}}
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
        semicolon    { set arg [expr {$arg | 0x40000000}] }
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

proc twapi::cert_property {hcert prop} {
    # TBD - need to cook some properties

    if {[string is integer -strict $prop]} {
        return [CertGetCertificateContextProperty $hcert $prop]
    } else {
        return [CertGetCertificateContextProperty $hcert [_cert_prop_id $prop] 1]
    }
}

# TBD - Also add cert_set_key_prov_from_crypt_context
proc twapi::cert_set_key_prov {hcert args} {
    # TB - make keycontainer explicit arg
    parseargs args {
        keycontainer.arg
        csp.arg
        {csptype.arg prov_rsa_full}
        {keysettype.arg user {user machine}}
        {silent.bool 0 0x40}
        {keyspec.arg signature {keyexchange signature}}
    } -maxleftover 0 -nulldefault -setvars

    set flags $silent
    if {$keysettype eq "machine"} {
        incr flags 0x20;        # CRYPT_KEYSET_MACHINE
    }

    # TBD - does the keyspec matter ? In case of self signed cert
    # which (keyexchange/signature) or both have to be specified ?
    Twapi_SetCertContextKeyProvInfo $hcert \
        [list $keycontainer $csp [_csp_type_name_to_id $csptype] $flags {} [_crypt_keyspec $keyspec]]
    return
}

proc twapi::cert_export {hcert} {
    return [lindex [Twapi_CertGetEncoded $hcert] 1]
}

proc twapi::cert_enhkey_usage {hcert {loc both}} {
    return [_cert_decode_enhkey [CertGetEnhancedKeyUsage $hcert [dict! {property 4 extension 2 both 0} $loc 1]]]
}

proc twapi::cert_key_usage {hcert} {
    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    return [_cert_decode_keyusage [Twapi_CertGetIntendedKeyUsage 0x10001 $hcert]]
}

proc twapi::cert_thumbprint {hcert} {
    binary scan [cert_property $hcert sha1_hash] H* hash
    return $hash
}

proc twapi::cert_info {hcert} {
    return [twine {-version -serialnumber -signaturealgorithm -issuer
        -start -end -subject -publickey -issuerid -subjectid -extensions} \
                [Twapi_CertGetInfo $hcert]]
}

proc twapi::cert_extension {hcert oid} {
    set ext [CertFindExtension $hcert [oid $oid]]
    if {[llength $ext] == 0} {
        return $ext
    }
    lassign $ext oid critical val
    return [list $critical [_cert_decode_extension $oid $val]]
}

proc twapi::cert_create_self_signed {subject args} {
    set args [_cert_create_parse_options $args opts]

    # TB - make keycontainer explicit arg
    array set opts [parseargs args {
        {keyspec.arg signature {keyexchange signature}}
        {keycontainer.arg {}}
        {keysettype.arg user {machine user}}
        {silent.bool 0 0x40}
        {csp.arg {}}
        {csptype.arg {prov_rsa_full}}
        {signaturealgorithm.arg {}}
    } -maxleftover 0 -ignoreunknown]

    set name_blob [cert_name_to_blob $subject]

    set kiflags $opts(silent)
    if {$opts(keysettype) eq "machine"} {
        incr kiflags 0x20;  # CRYPT_MACHINE_KEYSET
    }
    set keyinfo [list \
                     $opts(keycontainer) \
                     $opts(csp) \
                     [_csp_type_name_to_id $opts(csptype)] \
                     $kiflags \
                     {} \
                     [_crypt_keyspec $opts(keyspec)]]
    
    set flags 0;                # Always 0 for now
    return [CertCreateSelfSignCertificate NULL $name_blob $flags $keyinfo \
                [_make_algorithm_identifier $opts(signaturealgorithm)] \
                $opts(start) $opts(end) $opts(extensions)]
}

proc twapi::cert_create_self_signed_from_crypt_context {subject hprov args} {
    set args [_cert_create_parse_options $args opts]

    array set opts [parseargs args {
        {signaturealgorithm.arg {}}
    } -maxleftover 0]

    set name_blob [cert_name_to_blob $subject]

    set flags 0;                # Always 0 for now
    return [CertCreateSelfSignCertificate $hprov $name_blob $flags {} \
                [_make_algorithm_identifier $opts(signaturealgorithm)] \
                $opts(start) $opts(end) $opts(extensions)]
}

proc twapi::cert_create {subject hprov cissuer args} {
    set args [_cert_create_parse_options $args opts]

    parseargs args {
        {keyspec.arg signature {keyexchange signature}}
    } -maxleftover 0 -setvars
    
    # TBD - check that issuer is a CA

    set issuer_info [cert_info $cissuer]
    set issuer_blob [cert_name_to_blob [dict get $issuer_info -subject] -format x500]
    set sigoid [dict get $issuer_info -signaturealgorithm]

    # TBD Issuer altnames - get from issuer cert
    # lappend exts [_make_altnames_ext $opts(altnames) $critical 1]


    # The subject key id in issuer's cert will become the
    # authority key id in the new cert
    # TBD - if fail, get the CERT_KEY_IDENTIFIER_PROP_ID
    # 2.5.29.14 -> oid_subject_key_identifier
    set issuer_subject_key_id [cert_extension $cissuer 2.5.29.14]
    if {[string length [lindex $issuer_subject_key_id 1]] } {
        # 2.5.29.35 -> oid_authority_key_identifier
        lappend opts(extensions) [list 2.5.29.35 0 [list [lindex $issuer_subject_key_id 1] {} {}]]
    }

    # Generate a subject key identifier for this cert based on a hash
    # of the public key
    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    # TBD - Do not hard code oid_rsa_rsa.
    set pubkey [CryptExportPublicKeyInfoEx $hprov \
                    [_crypt_keyspec $keyspec] \
                    0x10001 \
                    [oid oid_rsa_rsa] \
                    0]
    set subject_key_id [Twapi_HashPublicKeyInfo $pubkey]
    lappend opts(extensions) [list 2.5.29.14 0 $subject_key_id]


    set start [timelist_to_large_system_time $opts(start)]
    set end [timelist_to_large_system_time $opts(end)]

    # 2 -> CERT_V3
    # issuer_id and subject_id for the certificate are left empty
    # as recommended by gutman's X.509 paper
    set cert_info [list 2 $opts(serialnumber) $sigoid $issuer_blob \
                       $start $end \
                       [cert_name_to_blob $subject] \
                       $pubkey {} {} \
                       $opts(extensions)]

    # We need to get the crypt provider for the issuer cert since
    # that is what will sign the new cert
    lassign [cert_property $cissuer key_prov_info] issuer_container issuer_provname issuer_provtype issuer_flags dontcare issuer_keyspec
    set hissuerprov [crypt_acquire $issuer_container -csp $issuer_provname -csptype $issuer_provtype -keysettype [expr {$issuer_flags & 0x20 ? "machine" : "user"}]]
    trap {
        # 0x10001 -> X509_ASN_ENCODING, 2 -> X509_CERT_TO_BE_SIGNED
        return [CryptSignAndEncodeCertificate $hissuerprov $issuer_keyspec \
                0x10001 2 $cert_info $sigoid]
    } finally {
        # TBD - test to make sure ok to close this if caller had
        # it open
        crypt_free $hissuerprov
    }
}

proc twapi::cert_ssl_verify {hcert args} {

    parseargs args {
        {ignoreerrors.arg {}}
        {cacheendcert.bool 0 0x1}
        {revocationcheckcacheonly.bool 0 0x80000000}
        {urlretrievalcacheonly.bool 0 0x4}
        {disablepass1qualityfiltering.bool 0 0x40}
        {returnlowerqualitycontexts.bool 0 0x80}
        {disableauthrootautoupdate.bool 0 0x100}
        {revocationcheck.arg all {none all leaf excluderoot}}
        usageall.arg
        usageany.arg 
        {engine.arg user {user machine}}
        {timestamp.arg ""}
        {hstore.arg NULL}
        server.arg
    } -setvars -maxleftover 0

    set flags [dict! {none 0 all 0x20000000 leaf 0x10000000 excluderoot 0x40000000} $revocationcheck]
    set flags [tcl::mathop::| $flags $cacheendcert $revocationcheckcacheonly $urlretrievalcacheonly $disablepass1qualityfiltering $returnlowerqualitycontexts $disableauthrootautoupdate]

    set usage_op 1;             # USAGE_MATCH_TYPE_OR
    if {[info exists usageall]} {
        if {[info exists usageany]} {
            error "Only one of -usageall and -usageany may be specified"
        }
        set usage_op 0;         # USAGE_MATCH_TYPE_AND
        set usage [_get_enhkey_usage_oids $usageall]
    } else {
        lappend usageany;       # Ensure it exists
        set usage [_get_enhkey_usage_oids $usageany]
    }

    set chainh [CertGetCertificateChain \
                    [dict* {user NULL machine {1 HCERTCHAINENGINE}} $engine] \
                    $hcert $timestamp $hstore \
                    [list [list $usage_op $usage]] $flags]

    trap {
        set verify_flags 0
        foreach ignore $ignoreerrors {
            set verify_flags [expr {$verify_flags | [dict! {
                time             0x07
                basicconstraints 0x08
                unknownca        0x10
                usage            0x20
                name             0x40
                policy           0x80
                revocation       0xf00
                criticalextensions 0x2000
            } $ignore]}]
        }

        if {[info exists server]} {
            set role 2;         # AUTHTYPE_SERVER
        } else {
            set role 1;         # AUTHTYPE_CLIENT
            set server ""
        }

        # I have no clue as to why some of these options have to
        # be specified in two different places
        set checks 0
        foreach {verify check} {
            0x7 0x2000
            0xf00 0x80
            0x10 0x100
            0x20 0x200
            0x40 0x1000
        } {
            if {$verify_flags & $verify} {
                set checks [expr {$checks | $check}]
            }
        }

        set status [Twapi_CertVerifyChainPolicySSL $chainh [list $verify_flags [list $role $checks $server]]]
        return [dict*  {
            0x00000000 ok
            0x80096004 signature
            0x80092010 revoked
            0x800B0109 untrustedroot
            0x800B010D untrustedtestroot
            0x800B010A chaining
            0x800B0110 wrongusage
            0x800B0101 expired
            0x800B0114 name
            0x800B0113 policy
            0x80096019 basicconstraints
            0x800B0105 criticalextension
            0x800B0102 validityperiodnesting
            0x80092012 norevocationcheck
            0x80092013 revocationoffline
            0x800B010F cnmatch
            0x800B0106 purpose
            0x800B0103 carole
        } [format 0x%8.8X $status]]

    } finally {
        CertFreeCertificateChain $chainh
    }
}

################################################################
# Cryptographic context commands

proc twapi::crypt_acquire {keycontainer args} {
    parseargs args {
        csp.arg
        {csptype.arg prov_rsa_full}
        {keysettype.arg user {user machine}}
        {create.bool 0 0x8}
        {silent.bool 0 0x40}
        {verifycontext.bool 0 0xf0000000}
    } -maxleftover 0 -nulldefault -setvars
    
    # Based on http://support.microsoft.com/kb/238187, if verifycontext
    # is not specified, default container must not be used as keys
    # from different applications might overwrite. The docs for
    # CryptAcquireContext say keycontainer must be empty if verifycontext
    # is specified. Thus they are mutually exclusive.
    if {! $verifycontext} {
        if {$keycontainer eq ""} {
            badargs! "Option -verifycontext must be specified for the default key container."
        }
    }

    set flags [expr {$create | $silent | $verifycontext}]
    if {$keysettype eq "machine"} {
        incr flags 0x20;        # CRYPT_KEYSET_MACHINE
    }

    return [CryptAcquireContext $keycontainer $csp [_csp_type_name_to_id $csptype] $flags]
}

proc twapi::crypt_free {hcrypt} {
    twapi::CryptReleaseContext $hcrypt
}

proc twapi::crypt_key_container_delete {keycontainer args} {
    parseargs args {
        csp.arg
        {csptype.arg prov_rsa_full}
        {keysettype.arg user {machine user}}
        force
    } -maxleftover 0 -nulldefault -setvars

    if {$keycontainer eq "" && ! $force} {
        error "Default container cannot be deleted unless the -force option is specified"
    }

    set flags 0x10;             # CRYPT_DELETEKEYSET
    if {$keysettype eq "machine"} {
        incr flags 0x20;        # CRYPT_MACHINE_KEYSET
    }

    return [CryptAcquireContext $keycontainer $csp [_csp_type_name_to_id $csptype] $flags]
}

proc twapi::crypt_key_generate {hprov algid args} {

    array set opts [parseargs args {
        {archivable.bool 0 0x4000}
        {salt.bool 0 4}
        {exportable.bool 0 1}
        {pregen.bool 0x40}
        {userprotected.bool 0 2}
        {nosalt40.bool 0 0x10}
        {size.int 0}
    } -maxleftover 0]

    if {![string is integer -strict $algid]} {
        # See wincrypt.h in SDK
        switch -nocase -exact -- $algid {
            keyexchange {set algid 1}
            signature {set algid 2}
            default {
                set id [CertOIDToAlgId [oid $algid]]
                if {$id == 0} {
                    badargs! "Invalid algorithm id '$algid'"
                }
                set algid $id
            }
        }
    }

    if {$opts(size) < 0 || $opts(size) > 65535} {
        badargs! "Bad key size value '$size':  must be positive integer less than 65536"
    }

    return [CryptGenKey $hprov $algid [expr {($opts(size) << 16) | $opts(archivable) | $opts(salt) | $opts(exportable) | $opts(pregen) | $opts(userprotected) | $opts(nosalt40)}]]
}

proc twapi::crypt_keypair {hprov keyspec} {
    return [CryptGetUserKey $hprov [dict! {keyexchange 1 signature 2} $keyspec]]
}

proc twapi::crypt_get_security_descriptor {hprov} {
    return [CryptGetProvParam $hprov 8 7]
}

proc twapi::crypt_set_security_descriptor {hprov secd} {
    CryptSetProvParam $hprov 8 $secd
}

proc twapi::crypt_key_container_name {hprov} {
    return [_ascii_binary_to_string [CryptGetProvParam $hprov 6 0]]
}

proc twapi::crypt_key_container_unique_name {hprov} {
    return [_ascii_binary_to_string [CryptGetProvParam $hprov 36 0]]
}

proc twapi::crypt_csp {hprov} {
    return [_ascii_binary_to_string [CryptGetProvParam $hprov 4 0]]
}

proc twapi::crypt_csptype {hprov} {
    binary scan [CryptGetProvParam $hprov 16 0] i i
    return [_csp_type_id_to_name $i]
}

proc twapi::crypt_key_container_names {hprov} {
    return [CryptGetProvParam $hprov 2 0]
}

proc twapi::crypt_session_key_size {hprov} {
    binary scan [CryptGetProvParam $hprov 20 0] i i
    return $i
}

proc twapi::crypt_keyset_type {hprov} {
    binary scan [CryptGetProvParam $hprov 27 0] i i
    return [expr {$i & 0x20 ? "machine" : "user"}]
}

proc twapi::crypt_symmetric_key_size {hprov} {
    binary scan [CryptGetProvParam $hprov 19 0] i i
    return $i
}

################################################################
# Utility procs

proc twapi::_algid {class type alg} {
    return [expr {($class << 13) | ($type << 9) | $alg}]
}

proc twapi::_make_algorithm_identifier {oid {param {}}} {
    if {[string length $oid] == 0} {
        return ""
    }
    set oid [oid $oid]
    if {[string length $param]} {
        return [list $oid $param]
    } else {
        return [list $oid]
    }
}

twapi::proc* twapi::_cert_prop_id {prop} {
    # Certificate property menomics
    variable _cert_prop_name_id_map
    array set _cert_prop_name_id_map {
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
    variable _cert_prop_name_id_map

    if {[string is integer -strict $prop]} {
        return $prop
    }
    if {![info exists _cert_prop_name_id_map($prop)]} {
        badargs! "Unknown certificate property id '$prop'" 3
    }

    return $_cert_prop_name_id_map($prop)
}

twapi::proc* twapi::_cert_prop_name {id} {
    variable _cert_prop_name_id_map
    variable _cert_prop_id_name_map

    _cert_prop_id key_prov_handle; # Just to init _cert_prop_name_id_map
    array set _cert_prop_id_name_map [swapl [array get _cert_prop_name_id_map]]
} {
    variable _cert_prop_id_name_map
    if {[info exists _cert_prop_id_name_map($id)]} {
        return $_cert_prop_id_name_map($id)
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

    return [dict! $system_stores $name 2]
}

twapi::proc* twapi::_csp_type_name_to_id prov {
    variable _csp_name_id_map

    array set _csp_name_id_map {
        prov_rsa_full           1
        prov_rsa_sig            2
        prov_dss                3
        prov_fortezza           4
        prov_ms_exchange        5
        prov_ssl                6
        prov_rsa_schannel       12
        prov_dss_dh             13
        prov_ec_ecdsa_sig       14
        prov_ec_ecnra_sig       15
        prov_ec_ecdsa_full      16
        prov_ec_ecnra_full      17
        prov_dh_schannel        18
        prov_spyrus_lynks       20
        prov_rng                21
        prov_intel_sec          22
        prov_replace_owf        23
        prov_rsa_aes            24
    }
} {
    variable _csp_name_id_map

    set key [string tolower $prov]

    if {[info exists _csp_name_id_map($key)]} {
        return $_csp_name_id_map($key)
    }

    if {[string is integer -strict $prov]} {
        return $prov
    }

    badargs! "Invalid or unknown provider name '$prov'" 3
}

twapi::proc* twapi::_csp_type_id_to_name prov {
    variable _csp_name_id_map
    variable _csp_id_name_map

    _csp_type_name_to_id prov_rsa_full; # Just to ensure _csp_name_id_map exists
    array set _csp_id_name_map [swapl [array get _csp_name_id_map]]
} {
    variable _csp_id_name_map
    if {[info exists _csp_id_name_map($prov)]} {
        return $_csp_id_name_map($prov)
    }

    if {[string is integer -strict $prov]} {
        return $prov
    }

    badargs! "Invalid or unknown provider id '$prov'" 3
}

twapi::proc* twapi::oid {name} {
    variable _name_oid_map
    if {![info exists _name_oid_map]} {
        oids;                       # To init the map
    }
} {
    variable _name_oid_map

    if {[info exists _name_oid_map($name)]} {
        return $_name_oid_map($name)
    }
    if {[regexp {^\d([\d\.]*\d)?$} $name]} {
        return $name
    } else {
        badargs! "Invalid OID '$name'"
    }

}

twapi::proc* twapi::oidname {oid} {
    variable _oid_name_map
    if {![info exists _oid_name_map]} {
        oids;                       # To init the map
    }
} {
    variable _oid_name_map

    if {[info exists _oid_name_map($oid)]} {
        return $_oid_name_map($oid)
    }
    if {[regexp {^\d([\d\.]*\d)?$} $oid]} {
        return $oid
    } else {
        badargs! "Invalid OID '$name'"
    }
}

twapi::proc* twapi::oids {{pattern *}} {
    variable _oid_name_map
    variable _name_oid_map

    # TBD - clean up table for rarely used OIDs
    array set _name_oid_map {
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

        oid_rsa               "1.2.840.113549"
        oid_pkcs              "1.2.840.113549.1"
        oid_rsa_hash          "1.2.840.113549.2"
        oid_rsa_encrypt       "1.2.840.113549.3"

        oid_pkcs_1            "1.2.840.113549.1.1"
        oid_pkcs_2            "1.2.840.113549.1.2"
        oid_pkcs_3            "1.2.840.113549.1.3"
        oid_pkcs_4            "1.2.840.113549.1.4"
        oid_pkcs_5            "1.2.840.113549.1.5"
        oid_pkcs_6            "1.2.840.113549.1.6"
        oid_pkcs_7            "1.2.840.113549.1.7"
        oid_pkcs_8            "1.2.840.113549.1.8"
        oid_pkcs_9            "1.2.840.113549.1.9"
        oid_pkcs_10           "1.2.840.113549.1.10"
        oid_pkcs_12           "1.2.840.113549.1.12"

        oid_rsa_rsa           "1.2.840.113549.1.1.1"
        oid_rsa_md2rsa        "1.2.840.113549.1.1.2"
        oid_rsa_md4rsa        "1.2.840.113549.1.1.3"
        oid_rsa_md5rsa        "1.2.840.113549.1.1.4"
        oid_rsa_sha1rsa       "1.2.840.113549.1.1.5"
        oid_rsa_setoaep_rsa   "1.2.840.113549.1.1.6"

        oid_rsa_dh            "1.2.840.113549.1.3.1"

        oid_rsa_data          "1.2.840.113549.1.7.1"
        oid_rsa_signeddata    "1.2.840.113549.1.7.2"
        oid_rsa_envelopeddata "1.2.840.113549.1.7.3"
        oid_rsa_signenvdata   "1.2.840.113549.1.7.4"
        oid_rsa_digesteddata  "1.2.840.113549.1.7.5"
        oid_rsa_hasheddata    "1.2.840.113549.1.7.5"
        oid_rsa_encrypteddata "1.2.840.113549.1.7.6"

        oid_rsa_emailaddr     "1.2.840.113549.1.9.1"
        oid_rsa_unstructname  "1.2.840.113549.1.9.2"
        oid_rsa_contenttype   "1.2.840.113549.1.9.3"
        oid_rsa_messagedigest "1.2.840.113549.1.9.4"
        oid_rsa_signingtime   "1.2.840.113549.1.9.5"
        oid_rsa_countersign   "1.2.840.113549.1.9.6"
        oid_rsa_challengepwd  "1.2.840.113549.1.9.7"
        oid_rsa_unstructaddr  "1.2.840.113549.1.9.8"
        oid_rsa_extcertattrs  "1.2.840.113549.1.9.9"
        oid_rsa_certextensions "1.2.840.113549.1.9.14"
        oid_rsa_smimecapabilities "1.2.840.113549.1.9.15"
        oid_rsa_prefersigneddata "1.2.840.113549.1.9.15.1"

        oid_rsa_smimealg              "1.2.840.113549.1.9.16.3"
        oid_rsa_smimealgesdh          "1.2.840.113549.1.9.16.3.5"
        oid_rsa_smimealgcms3deswrap   "1.2.840.113549.1.9.16.3.6"
        oid_rsa_smimealgcmsrc2wrap    "1.2.840.113549.1.9.16.3.7"

        oid_rsa_md2           "1.2.840.113549.2.2"
        oid_rsa_md4           "1.2.840.113549.2.4"
        oid_rsa_md5           "1.2.840.113549.2.5"

        oid_rsa_rc2cbc        "1.2.840.113549.3.2"
        oid_rsa_rc4           "1.2.840.113549.3.4"
        oid_rsa_des_ede3_cbc  "1.2.840.113549.3.7"
        oid_rsa_rc5_cbcpad    "1.2.840.113549.3.9"


        oid_ansi_x942         "1.2.840.10046"
        oid_ansi_x942_dh      "1.2.840.10046.2.1"

        oid_x957              "1.2.840.10040"
        oid_x957_dsa          "1.2.840.10040.4.1"
        oid_x957_sha1dsa      "1.2.840.10040.4.3"

        oid_ds                "2.5"
        oid_dsalg             "2.5.8"
        oid_dsalg_crpt        "2.5.8.1"
        oid_dsalg_hash        "2.5.8.2"
        oid_dsalg_sign        "2.5.8.3"
        oid_dsalg_rsa         "2.5.8.1.1"

        oid_pkix_kp_server_auth "1.3.6.1.5.5.7.3.1"
        oid_pkix_kp_client_auth "1.3.6.1.5.5.7.3.2"
        oid_pkix_kp_code_signing   "1.3.6.1.5.5.7.3.3"
        oid_pkix_kp_email_protection      "1.3.6.1.5.5.7.3.4"
        oid_pkix_kp_ipsec_end_system "1.3.6.1.5.5.7.3.5"
        oid_pkix_kp_ipsec_tunnel "1.3.6.1.5.5.7.3.6"
        oid_pkix_kp_ipsec_user "1.3.6.1.5.5.7.3.7"
        oid_pkix_kp_timestamp_signing "1.3.6.1.5.5.7.3.8"
        oid_pkix_kp_ocsp_signing      "1.3.6.1.5.5.7.3.9"

        oid_oiw               "1.3.14"

        oid_oiwsec            "1.3.14.3.2"
        oid_oiwsec_md4rsa     "1.3.14.3.2.2"
        oid_oiwsec_md5rsa     "1.3.14.3.2.3"
        oid_oiwsec_md4rsa2    "1.3.14.3.2.4"
        oid_oiwsec_desecb     "1.3.14.3.2.6"
        oid_oiwsec_descbc     "1.3.14.3.2.7"
        oid_oiwsec_desofb     "1.3.14.3.2.8"
        oid_oiwsec_descfb     "1.3.14.3.2.9"
        oid_oiwsec_desmac     "1.3.14.3.2.10"
        oid_oiwsec_rsasign    "1.3.14.3.2.11"
        oid_oiwsec_dsa        "1.3.14.3.2.12"
        oid_oiwsec_shadsa     "1.3.14.3.2.13"
        oid_oiwsec_mdc2rsa    "1.3.14.3.2.14"
        oid_oiwsec_sharsa     "1.3.14.3.2.15"
        oid_oiwsec_dhcommmod  "1.3.14.3.2.16"
        oid_oiwsec_desede     "1.3.14.3.2.17"
        oid_oiwsec_sha        "1.3.14.3.2.18"
        oid_oiwsec_mdc2       "1.3.14.3.2.19"
        oid_oiwsec_dsacomm    "1.3.14.3.2.20"
        oid_oiwsec_dsacommsha "1.3.14.3.2.21"
        oid_oiwsec_rsaxchg    "1.3.14.3.2.22"
        oid_oiwsec_keyhashseal "1.3.14.3.2.23"
        oid_oiwsec_md2rsasign "1.3.14.3.2.24"
        oid_oiwsec_md5rsasign "1.3.14.3.2.25"
        oid_oiwsec_sha1       "1.3.14.3.2.26"
        oid_oiwsec_dsasha1    "1.3.14.3.2.27"
        oid_oiwsec_dsacommsha1 "1.3.14.3.2.28"
        oid_oiwsec_sha1rsasign "1.3.14.3.2.29"

        oid_oiwdir            "1.3.14.7.2"
        oid_oiwdir_crpt       "1.3.14.7.2.1"
        oid_oiwdir_hash       "1.3.14.7.2.2"
        oid_oiwdir_sign       "1.3.14.7.2.3"
        oid_oiwdir_md2        "1.3.14.7.2.2.1"
        oid_oiwdir_md2rsa     "1.3.14.7.2.3.1"

        oid_infosec                       "2.16.840.1.101.2.1"
        oid_infosec_sdnssignature         "2.16.840.1.101.2.1.1.1"
        oid_infosec_mosaicsignature       "2.16.840.1.101.2.1.1.2"
        oid_infosec_sdnsconfidentiality   "2.16.840.1.101.2.1.1.3"
        oid_infosec_mosaicconfidentiality "2.16.840.1.101.2.1.1.4"
        oid_infosec_sdnsintegrity         "2.16.840.1.101.2.1.1.5"
        oid_infosec_mosaicintegrity       "2.16.840.1.101.2.1.1.6"
        oid_infosec_sdnstokenprotection   "2.16.840.1.101.2.1.1.7"
        oid_infosec_mosaictokenprotection "2.16.840.1.101.2.1.1.8"
        oid_infosec_sdnskeymanagement     "2.16.840.1.101.2.1.1.9"
        oid_infosec_mosaickeymanagement   "2.16.840.1.101.2.1.1.10"
        oid_infosec_sdnskmandsig          "2.16.840.1.101.2.1.1.11"
        oid_infosec_mosaickmandsig        "2.16.840.1.101.2.1.1.12"
        oid_infosec_suiteasignature       "2.16.840.1.101.2.1.1.13"
        oid_infosec_suiteaconfidentiality "2.16.840.1.101.2.1.1.14"
        oid_infosec_suiteaintegrity       "2.16.840.1.101.2.1.1.15"
        oid_infosec_suiteatokenprotection "2.16.840.1.101.2.1.1.16"
        oid_infosec_suiteakeymanagement   "2.16.840.1.101.2.1.1.17"
        oid_infosec_suiteakmandsig        "2.16.840.1.101.2.1.1.18"
        oid_infosec_mosaicupdatedsig      "2.16.840.1.101.2.1.1.19"
        oid_infosec_mosaickmandupdsig     "2.16.840.1.101.2.1.1.20"
        oid_infosec_mosaicupdatedinteg    "2.16.840.1.101.2.1.1.21"
    }

    # OIDs for certificate extensions
    array set _name_oid_map {
        oid_authority_key_identifier_old  "2.5.29.1"
        oid_key_attributes            "2.5.29.2"
        oid_cert_policies_95          "2.5.29.3"
        oid_key_usage_restriction     "2.5.29.4"
        oid_subject_alt_name_old          "2.5.29.7"
        oid_issuer_alt_name_old           "2.5.29.8"
        oid_basic_constraints_old     "2.5.29.10"
        oid_key_usage                 "2.5.29.15"
        oid_privatekey_usage_period   "2.5.29.16"
        oid_basic_constraints        "2.5.29.19"

        oid_cert_policies             "2.5.29.32"
        oid_any_cert_policy           "2.5.29.32.0"
        oid_inhibit_any_policy        "2.5.29.54"

        oid_authority_key_identifier "2.5.29.35"
        oid_subject_key_identifier    "2.5.29.14"
        oid_subject_alt_name2         "2.5.29.17"
        oid_issuer_alt_name          "2.5.29.18"
        oid_crl_reason_code           "2.5.29.21"
        oid_reason_code_hold          "2.5.29.23"
        oid_crl_dist_points           "2.5.29.31"
        oid_enhanced_key_usage        "2.5.29.37"

        oid_any_enhanced_key_usage    "2.5.29.37.0"

        oid_crl_number                "2.5.29.20"
        oid_delta_crl_indicator       "2.5.29.27"
        oid_issuing_dist_point        "2.5.29.28"
        oid_freshest_crl              "2.5.29.46"
        oid_name_constraints          "2.5.29.30"

        oid_policy_mappings           "2.5.29.33"
        oid_legacy_policy_mappings    "2.5.29.5"
        oid_policy_constraints        "2.5.29.36"
    }

    array set _oid_name_map [swapl [array get _name_oid_map]]
} {
    variable _name_oid_map
    return [array get _name_oid_map $pattern]
}


proc twapi::_make_altnames_ext {altnames critical {issuer 0}} {
    set names {}
    foreach {alttype altname} $altnames {
        lappend names [list \
                           [dict get {
                               other 1
                               email 2
                               dns   3
                               directory 5
                               url 7
                               ip  8
                               registered 9
                           } $alttype] $altname]
    }

    return [list [expr {$issuer ? "2.5.29.18" : "2.5.29.17"}] $critical $names]
}

proc twapi::_get_enhkey_usage_oids {names} {
    array set map [oids oid_pkix_kp_*]

    # We use an array to remove duplicates
    array set oids {}
    foreach name $names {
        if {[info exists map($name)]} {
            set oids($map($name)) 1
        } elseif {[info exists map(oid_pkix_kp_$name)]} {
            set oids($map(oid_pkix_kp_$name)) 1
        } elseif {[regexp {^\d([\d\.]*\d)?$} $name]} {
            # Any OID will do
            set oids($name) 1
        } else {
            error "Invalid Enhanced Key Usage OID \"$name\""
        }
    }
    return [array names oids]
}

proc twapi::_make_enhkeyusage_ext {enhkeyusage critical} {
    return [list "2.5.29.37" $critical [_get_enhkey_usage_oids $enhkeyusage]]
}

twapi::proc* twapi::_init_keyusage_names {} {
    variable _keyusage_byte1
    variable _keyusage_byte2
    set _keyusage_byte1 {
        digital_signature     0x80
        non_repudiation       0x40
        key_encipherment      0x20
        data_encipherment     0x10
        key_agreement         0x08
        key_cert_sign         0x04
        crl_sign              0x02
        encipher_only         0x01
    }
    set _keyusage_byte2 {
        decipher_only         0x80
    }
} {}

proc twapi::_make_basic_constraints_ext {isca capathlen} {
    if {$isca} {
        if {$capathlen < 0} {
            set basic {1 0 0};  # No path length constraint
        } else {
            set basic [list 1 1 $capathlen]
        }
    } else {
        set basic {0 0 0}
    }
    return [list "2.5.29.19" 1 $basic]
}

proc twapi::_make_keyusage_ext {keyusage critical} {
    variable _keyusage_byte1
    variable _keyusage_byte2

    _init_keyusage_names
    set byte1 0
    set byte2 0
    foreach usage $keyusage {
        if {[dict exists $_keyusage_byte1 $usage]} {
            set byte1 [expr {$byte1 | [dict get $_keyusage_byte1 $usage]}]
        } elseif {[dict exists $_keyusage_byte2 $usage]} {
            set byte2 [expr {$byte2 | [dict get $_keyusage_byte2 $usage]}]
        } else {
            error "Invalid key usage value \"$keyusage\""
        }
    }

    set bin [binary format cc $byte1 $byte2]
    # 7 -> # unused bits in last byte
    return [list "2.5.29.15" $critical [list $bin 7]]
}

# Given a byte array, decode to key usage flags
proc twapi::_cert_decode_keyusage {bin} {
    variable _keyusage_byte1
    variable _keyusage_byte2
    
    _init_keyusage_names

    binary scan $bin c* bytes

    if {[llength $bytes] == 0} {
        return *;               # Field not present, TBD
    }

    set usages {}
    set byte [lindex $bytes 0]
    dict for {key val} $_keyusage_byte1 {
        if {$byte & $val} {
            lappend usages $key
        }
    } 

    set byte [lindex $bytes 1]
    dict for {key val} $_keyusage_byte2 {
        if {$byte & $val} {
            lappend usages $key
            set byte [expr {$byte & ~$val}]
        }
    } 

    if {0} {
        # Commented out because some certificates seem to contain
        # bits not defined by RF5280. Do not barf on these

        # For the second byte, not all bits are defined. Error if any
        # that we do not understand
        if {$byte} {
            error "Key usage sequence $bytes includes unsupported bits"
        }

        # If there are more bytes, they should all be 0 as well
        foreach byte [lrange $bytes 2 end] {
            if {$byte} {
                error "Key usage sequence $bytes includes unsupported bits"
            }
        }
    }

    return $usages
}

proc twapi::_cert_decode_enhkey {vals} {
    set result {}
    foreach val $vals {
        lappend result [dict* [swapl [oids oid_pkix_kp_*]] $val]
    }
    return $result
}

proc twapi::_cert_decode_extension {oid val} {
    # TBD - see what other types need to be decoded
    switch $oid {
        2.5.29.15 { return [_cert_decode_keyusage $val] }
        2.5.29.37 { return [_cert_decode_enhkey $val] }
    }
    return $val
}

proc twapi::_crypt_keyspec {keyspec} {
    return [dict* {keyexchange 1 signature 2} $keyspec]
}

proc twapi::_cert_create_parse_options {optvals optsvar} {
    upvar 1 $optsvar opts

    parseargs optvals {
        start.int
        end.int
        serialnumber.arg
        altnames.arg
        {critical.arg {}}
        enhkeyusage.arg
        keyusage.arg
        {purpose.arg {} {}}
        {capathlen.int -1}
    } -ignoreunknown -setvars

    set ca [expr {"ca" in $purpose}]
    set sslserver [expr {"sslserver" in $purpose}]
    set sslclient [expr {"sslclient" in $purpose}]

    if {[info exists serialnumber]} {
        if {$serialnumber <= 0 || $serialnumber > 0x7fffffffffffffff} {
            badargs! "Serial number must be specified as a positive wide integer."
        }
        # Format as little endian
        set opts(serialnumber) [binary format w $serialnumber]
    } else {
        # Generate 15 byte random and add high byte (little endian)
        # to 0x01 to ensure it is treated as positive
        set opts(serialnumber) "[random_bytes 15]\x01"
    }
    
    # Validity period
    if {![info exists start]} {
        set opts(start) [_seconds_to_timelist [clock seconds] 1]
    }
    if {![info exists end]} {
        set opts(end) $opts(start)
        lset opts(end) 0 [expr {[lindex $opts(end) 0] + 1}]
    }

    # Generate the extensions list
    set exts {}
    lappend exts [_make_basic_constraints_ext $ca $capathlen]
    if {$ca} {
        lappend keyusage key_cert_sign crl_sign
    }
    if {$sslserver || $sslclient} {
        lappend keyusage digital_signature key_encipherment key_agreement
        if {$sslserver} {
            lappend enhkeyusage oid_pkix_kp_server_auth
        }
        if {$sslclient} {
            lappend enhkeyusage oid_pkix_kp_client_auth
        }
    }
    if {[info exists keyusage]} {
        lappend exts [_make_keyusage_ext $keyusage [expr {"keyusage" in $critical}]]
    }
    if {[info exists enhkeyusage]} {
        lappend exts [_make_enhkeyusage_ext $enhkeyusage [expr {"enhkeyusage" in $critical}]]
    }
    if {[info exists altnames]} {
        lappend exts [_make_altnames_ext $altnames [expr {"altnames" in $critical}] 0]
    }

    set opts(extensions) $exts

    return $optvals
}

proc twapi::_cert_add_parseargs {vargs} {
    upvar 1 $vargs optvals
    parseargs optvals {
        {disposition.arg preserve {overwrite duplicate update preserve}}
    } -maxleftover 0 -setvars

    # 4 -> CERT_STORE_ADD_ALWAYS
    # 3 -> CERT_STORE_ADD_REPLACE_EXISTING
    # 6 -> CERT_STORE_ADD_NEWER
    # 1 -> CERT_STORE_ADD_NEW

    return [list disposition \
                [dict get {
                    duplicate 4
                    overwrite 3
                    update 6
                    preserve 1
                } $disposition]]
}

# If we are being sourced ourselves, then we need to source the remaining files.
if {[file tail [info script]] eq "crypto.tcl"} {
    source [file join [file dirname [info script]] sspi.tcl]
}

