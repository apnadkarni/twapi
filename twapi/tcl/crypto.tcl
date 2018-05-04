#
# Copyright (c) 2007-2018, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {
    variable wintrust_guids
    # Array key names match those in softpub.h in SDK
    array set wintrust_guids {
        action_generic_verify_v2    00AAC56B-CD44-11d0-8CC2-00C04FC295EE
        action_trust_provider_test  573E31F8-DDBA-11d0-8CCB-00C04FC295EE
        action_generic_cert_verify  189A3842-3041-11d1-85E1-00C04FC295EE
        action_generic_chain_verify fc451c16-ac75-11d1-b4b8-00c04fb66ea0
        httpsprov_action            573E31F8-AABA-11d0-8CCB-00C04FC295EE
        driver_action_verify        F750E6C3-38EE-11d1-85E5-00C04FC295EE
    }

    # Dictionaries used by capi_encrypt|decrypt_bytes to store partial blocks of data
    # First level key is Crypto key handle
    # Second level keys are Blocklen (block size in bytes) and Data (data bytes left over)
    variable _capi_encrypt_partials
    variable _capi_decrypt_partials
    set _capi_encrypt_partials {}
    set _capi_decrypt_partials {}
}

### Hash functions

proc twapi::capi_hash_create {hcrypt algid {hkey NULL}} {
    return [CryptCreateHash $hcrypt [capi_algid $algid] $hkey]
}

proc twapi::capi_hash_string {hhash s {enc utf-8}} {
    return [capi_hash_bytes $hhash [encoding convertto $enc $s] 0]
}

proc twapi::capi_hash_value {hhash} {
    return [CryptGetHashParam $hhash 2]; # HP_HASHVAL
}

proc twapi::capi_hash_sign {hhash keyspec args} {
    # -pad not documented because new Windows version do not support X.931
    # and there are some openssl incompatibilities I cannot figure out
    parseargs args {
        {nohashoid.bool 0 1}
        {pad.arg pkcs1 {pkcs1 x931}}
    } -maxleftover 0 -setvars
    set flags [expr {[dict get {pkcs1 0 x931 4} $pad] | $nohashoid}]
    return [CryptSignHash $hhash [_crypt_keyspec $keyspec] "" $flags]
}

proc twapi::capi_hash_verify {hhash sig hkey args} {
    # -pad not documented because new Windows version do not support X.931
    # and there are some openssl incompatibilities I cannot figure out
    parseargs args {
        {nohashoid.bool 0 1}
        {pad.arg pkcs1 {pkcs1 x931}}
    } -maxleftover 0 -setvars
    set flags [expr {[dict get {pkcs1 0 x931 4} $pad] | $nohashoid}]
    return [CryptVerifySignature $hhash $sig $hkey "" $flags]
}

proc twapi::_do_hash {csptype alg s {enc ""}} {
    if {$enc ne ""} {
        set s [encoding convertto $enc $s]
    }
    set hcrypt [crypt_acquire -csptype $csptype]
    trap {
        set hhash [capi_hash_create $hcrypt $alg]
        capi_hash_bytes $hhash $s
        return [capi_hash_value $hhash]
    } finally {
        if {[info exists hhash]} {
            capi_hash_free $hhash
        }
        crypt_free $hcrypt
    }
}

interp alias {} twapi::md5 {} twapi::_do_hash prov_rsa_full md5
interp alias {} twapi::sha1 {} twapi::_do_hash prov_rsa_full sha1
interp alias {} twapi::sha256 {} twapi::_do_hash prov_rsa_aes sha_256
interp alias {} twapi::sha384 {} twapi::_do_hash prov_rsa_aes sha_384
interp alias {} twapi::sha512 {} twapi::_do_hash prov_rsa_aes sha_512

proc twapi::hmac {data key {prf sha1} {charset {}}} {
    if {$charset ne ""} {
        set data [encoding convertto $charset $data]
    }

    # Choose prov_rsa_aes because older CSP's do not support sha256
    set hcrypt [crypt_acquire -csptype prov_rsa_aes]
    try {
        # The algorithm specified for importing the key actually is not
        # executed at all. It's only used for importing the key.
        # However it has to be something that will accept any key size.
        # On Windows 8 at least, RC4 seems to require at least 5 byte keys.
        # RC2 on the other hand, if the -ipsechmac flag is specifie
        # will accept any number. TBD - the pbkdf2 source code implies
        # on Win8.1 single byte keys will not be accepted by rc2 and
        # keys need to be padded with 0's. Need to check that.
        set hkey [crypt_import_key $hcrypt [capi_keyblob_concealed rc2 $key] -ipsechmac 1]
        set hhash [capi_hash_create $hcrypt hmac $hkey]
        # 5 -> HP_HMAC_INFO
        CryptSetHashParam $hhash 5 [list [capi_algid $prf] "" ""]
        capi_hash_bytes $hhash $data
        return [capi_hash_value $hhash]
    } finally {
        if {[info exists hhash]} {
            capi_hash_free $hhash
        }
        if {[info exists hkey]} {
            capi_key_free $hkey
        }
        crypt_free $hcrypt
    }
}

                  
### Data protection

proc twapi::protect_data {data args} {

    # Not used because doesn't seem to have any effect 
    # {promptonunprotect.bool 0 0x1}
    parseargs args {
        {description.arg ""}
        {localmachine.bool 0 0x4}
        {noui.bool 0 0x1}
        {audit.bool 0 0x10}
        {hwnd.arg NULL}
        prompt.arg
    } -setvars -maxleftover 0

    if {[info exists prompt]} {
        # 2 -> PROMPTONPROTECT
        set prompt [list 2 $hwnd $prompt]
    } else {
        set prompt {}
    }

    return [CryptProtectData $data $description "" "" $prompt [expr {$localmachine | $noui | $audit}]]
}

proc twapi::unprotect_data {data args} {
    # Do not seem to have any effect
    # {promptonunprotect.bool 0 0x1}
    # {promptonprotect.bool 0 0x2}
    parseargs args {
        {withdescription.bool 0}
        {noui.bool 0 0x1}
        {hwnd.arg NULL}
        prompt.arg
    } -setvars -maxleftover 0

    if {[info exists prompt]} {
        # 2 -> PROMPTONPROTECT
        set prompt [list 2 $hwnd $prompt]
    } else {
        set prompt {}
    }

    set data [CryptUnprotectData $data "" "" $prompt $noui]
    if {$withdescription} {
        return $data
    } else {
        return [lindex $data 0]
    }
}



################################################################
# Certificate Stores

# Close a certificate store
proc twapi::cert_store_release {hstore} {
    CertCloseStore $hstore 0
    return
}

proc twapi::cert_temporary_store {args} {
    # TBD - add support for PKCS12_NO_PERSIST_KEY post-XP. If not
    # specified and on XP document a means of getting rid of the key
    # containers. See https://msdn.microsoft.com/en-us/library/ms867088.aspx#pk_topic6
    # Also CryptAcquireCertificatePrivateKey and GetCryptProvFromCert
    # might be useful in this regard
    parseargs args {
        {encoding.arg {} {der pem {}}}
        serialized.arg
        pkcs7.arg
        {password.arg ""}
        pfx.arg
        pkcs12.arg
        {exportableprivatekeys.bool 0 1}
        {userprotected.bool 0 2}
        keysettype.arg
    } -setvars -maxleftover 0
    
    set nformats 0
    foreach format {serialized pkcs7 pfx pkcs12} {
        if {[info exists $format]} {
            set data [set $format]
            incr nformats
        }
    }
    if {$nformats > 1} {
        badargs! "At most one of -pfx, -pkcs12, -pkcs7 or -serialized may be specified."
    }
    if {$nformats == 0} {
        # 2 -> CERT_STORE_PROV_MEMORY 
        return [CertOpenStore 2 0 NULL 0 ""]
    }
    
    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING

    if {[info exists serialized]} {
        # 6 -> CERT_STORE_PROV_SERIALIZED
        return [CertOpenStore 6 0x10001 NULL 0 $data]
    }

    if {[info exists pkcs7]} {
        # 5 -> CERT_STORE_PROV_PKCS7
        return [CertOpenStore 5 0x10001 NULL 0 [_pem_decode $data $encoding]]
    }

    # PFX/PKCS12
    if {[string length $password] == 0} {
        set password [conceal ""]
    }
    set flags 0
    if {[info exists keysettype]} {
        set flags [dict! {user 0x1000 machine 0x20} $keysettype]
    }

    set flags [tcl::mathop::| $flags $exportableprivatekeys $userprotected]
    return [PFXImportCertStore $data $password $flags]
}

proc twapi::cert_file_store_open {path args} {
    set flags [_parse_store_open_opts $args]

    if {! ($flags & 0x00008000)} {
        # If not readonly, set commitenable
        set flags [expr {$flags | 0x00010000}]
    }

    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    # 8 -> CERT_STORE_PROV_FILENAME_W
    return [CertOpenStore 8 0x10001 NULL $flags [file nativename [file normalize $path]]]
}

proc twapi::cert_serialized_store_open {data args} {
    set flags [_parse_store_open_opts $args]

    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    # 6 -> CERT_STORE_PROV_SERIALIZED
    return [CertOpenStore 6 0x10001 NULL $flags $data]
}



proc twapi::cert_physical_store_open {name location args} {
    variable _system_stores

    set flags [_parse_store_open_opts $args]
    incr flags [_system_store_id $location]
    # 14 -> CERT_STORE_PROV_PHYSICAL_W
    return [CertOpenStore 14 0 NULL $flags $name]
}

proc twapi::cert_physical_store_delete {name location} {
    set flags 0x10;             # CERT_STORE_DELETE_FLAG
    incr flags [_system_store_id $location]
    
    # 14 -> CERT_STORE_PROV_PHYSICAL_W
    return [CertOpenStore 14 0 NULL $flags $name]
}

# TBD - document and figure out what format to return data in
proc twapi::cert_physical_stores {system_store_name location} {
    return [CertEnumPhysicalStore $system_store_name [_system_store_id $location]]
}

proc twapi::cert_system_store_open {name args} {
    variable _system_stores

    if {[llength $args] == 0} {
        return [CertOpenSystemStore $name]
    }

    set flags [_parse_store_open_opts [lassign $args location]]
    incr flags [_system_store_id $location]
    return [CertOpenStore 10 0 NULL $flags $name]
}

proc twapi::cert_system_store_delete {name location} {
    set flags 0x10;             # CERT_STORE_DELETE_FLAG
    incr flags [_system_store_id $location]
    return [CertOpenStore 10 0 NULL $flags $name]
}

proc twapi::cert_system_store_locations {} {
    set l {}
    foreach e [CertEnumSystemStoreLocation 0] {
        lappend l [lindex $e 0]
    }
    return $l
}

proc twapi::cert_system_stores {location} {
    set l {}
    foreach e [CertEnumSystemStore [_system_store_id $location] ""] {
        lappend l [lindex $e 0]
    }
    return $l
}

# TBD - document?
proc twapi::cert_store_iterate {hstore varname script {type any} {term {}}} {
    upvar 1 $varname cert
    set cert NULL
    while {1} {
        set cert [cert_store_find_certificate $hstore $type $term $cert]
        if {$cert eq ""} break
        switch [catch {uplevel 1 $script} result options] {
            0 -
            4 {}
            3 {
                cert_release $cert
                set cert ""
                return
            }
            1 -
            default {
                cert_release $cert
                set cert ""
                return -options $options $result
            }
        }
    }
    return
}

proc twapi::cert_store_find_certificate {hstore {type any} {term {}} {hcert NULL}} {

    # TBD subject_cert 11<<16
    # TBD key_spec 9<<16

    set term_types {
        any 0
        existing 13<<16
        key_identifier 15<<16
        md5_hash 4<<16
        subject_public_key_md5_hash 18<<16
        sha1_hash 1<<16
        signature_hash 14<<16
        issuer_name (2<<16)|4
        subject_name  (2<<16)|7
        issuer_substring (8<<16)|4
        subject_substring (8<<16)|7
        property 5<<16
        public_key 6<<16
    }

    if {$type eq "property"} {
        set term [_cert_prop_id $term]
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
    parseargs args {
        {encoding.arg {} {der pem {}}}
    } -ignoreunknown -setvars
    array set opts [_cert_add_parseargs args]
    return [CertAddEncodedCertificateToStore $hstore 0x10001 [_pem_decode $enccert $encoding] $opts(disposition)]
}

proc twapi::cert_store_export_pem {hstore} {
    set pem {}
    cert_store_iterate $hstore c {append pem [cert_export $c]\n}
    return $pem
}

proc twapi::cert_store_export_pfx {hstore password args} {
    parseargs args {
        {exportprivatekeys.bool 0 0x4}
        {failonmissingkey.bool 0 0x1}
        {failonunexportablekey.bool 0 0x2}
    } -maxleftover 0 -setvars

    if {[string length $password] == 0} {
        set password [conceal ""]
    }

    # NOTE: the -fail* flags only take effect iff the certificate in the store
    # claims to have a private key but does not actually have one. It will
    # not fail if the cert does not actually claim to have a private key

    set flags [tcl::mathop::| $exportprivatekeys $failonunexportablekey $failonmissingkey]

    return [PFXExportCertStoreEx $hstore $password {} $flags]
}
interp alias {} twapi::cert_store_export_pkcs12 {} twapi::cert_store_export_pfx

proc twapi::cert_store_commit {hstore args} {
    array set opts [parseargs args {
        {force.bool 0}
    } -maxleftover 0]
    
    return [Twapi_CertStoreCommit $hstore $opts(force)]
}

proc twapi::cert_store_serialize {hstore} {
    return [Twapi_CertStoreSerialize $hstore 1]
}

proc twapi::cert_store_export_pkcs7 {hstore args} {
    parseargs args {
        {encoding.arg pem {der pem}}
    } -setvars -maxleftover 0
    
    return [_as_pem_or_der [Twapi_CertStoreSerialize $hstore 2] "PKCS7" $encoding]
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
        {reverse.bool 0 0x02000000}
        {noquote.bool 0 0x10000000}
        {noplus.bool  0 0x20000000}
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
        {reverse.bool 0 0x02000000}
        {noquote.bool 0 0x10000000}
        {noplus.bool  0 0x20000000}
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

proc twapi::cert_enum_properties {hcert args} {
    parseargs args {
        names
    } -setvars -maxleftover 0
    
    set id 0
    set ids {}
    while {[set id [CertEnumCertificateContextProperties $hcert $id]]} {
        if {$names} {
            lappend ids [_cert_prop_name $id]
        } else {
            lappend ids $id
        }
    }
    return $ids
}

proc twapi::cert_property {hcert prop} {
    # TBD - need to cook some properties - enhkey_usage

    if {[string is integer -strict $prop]} {
        return [CertGetCertificateContextProperty $hcert $prop]
    } else {
        return [CertGetCertificateContextProperty $hcert [_cert_prop_id $prop] 1]
    }
}

proc twapi::cert_property_set {hcert prop propval} {
    switch $prop {
        pvk_file -
        friendly_name -
        description {
            set val [encoding convertto unicode "${propval}\0"]
        }
        enhkey_usage {
            set val [::twapi::CryptEncodeObjectEx 2.5.29.37 [_get_enhkey_usage_oids $propval]]
        }
        default {
            badargs! "Invalid or unsupported property name \"$prop\". Must be one of [join $unicode_props {, }]."
        }
    }

    CertSetCertificateContextProperty $hcert [_cert_prop_id $prop] 0 $val
}

proc twapi::cert_property_delete {hcert prop} {
    CertSetCertificateContextProperty $hcert [_cert_prop_id $prop] 0
}

# TBD - Also add cert_set_key_prov_from_crypt_context
proc twapi::cert_set_key_prov {hcert keycontainer keyspec args} {
    parseargs args {
        csp.arg
        {csptype.arg prov_rsa_full}
        {keysettype.arg user {user machine}}
        {silent.bool 0 0x40}
    } -maxleftover 0 -nulldefault -setvars

    set flags $silent
    if {$keysettype eq "machine"} {
        incr flags 0x20;        # CRYPT_KEYSET_MACHINE
    }

    # 2 -> CERT_KEY_PROV_INFO_PROP_ID
    # TBD - the provider param is hardcoded as {}. Should that be an option ?
    CertSetCertificateContextProperty $hcert 2 0 \
        [list $keycontainer $csp [_csp_type_name_to_id $csptype] $flags {} [_crypt_keyspec $keyspec]]
    return
}

proc twapi::cert_export {hcert args} {
    parseargs args {
        {encoding.arg pem {der pem}}
    } -maxleftover 0 -setvars

    return [_as_pem_or_der [lindex [Twapi_CertGetEncoded $hcert] 1] CERTIFICATE $encoding]
}

proc twapi::cert_import {enccert args} {
    parseargs args {
        {encoding.arg {} {der pem {}}}
    } -maxleftover 0 -setvars
    return [CertCreateCertificateContext 0x10001 [_pem_decode $enccert $encoding]]
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
    # TBD - add option to cook extensions using _cert_decode_extension
    # instead of returning the raw form
    set info [twine {
        -version -serialnumber -signaturealgorithm -issuer
        -start -end -subject -publickey -issuerid -subjectid -extensions} \
                [Twapi_CertGetInfo $hcert]]
    dict set info -start \
        [clock format \
             [large_system_time_to_secs_since_1970 [dict get $info -start]] \
             -timezone :UTC \
             -format "%Y-%m-%d %H:%M:%S"]
    dict set info -end \
        [clock format \
             [large_system_time_to_secs_since_1970 [dict get $info -end]] \
             -timezone :UTC \
             -format "%Y-%m-%d %H:%M:%S"]

    return $info
}

proc twapi::cert_extension {hcert oid} {
    # TBD - add option to "cook" OID
    set ext [CertFindExtension $hcert [oid $oid]]
    if {[llength $ext] == 0} {
        return $ext
    }
    lassign $ext oid critical val
    return [list $critical [_cert_decode_extension $oid $val]]
}

proc twapi::cert_create_self_signed {subject keycontainer keyspec args} {
    set args [_cert_create_parse_options $args opts]

    array set opts [parseargs args {
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
                     $keycontainer \
                     $opts(csp) \
                     [_csp_type_name_to_id $opts(csptype)] \
                     $kiflags \
                     {} \
                     [_crypt_keyspec $keyspec]]
    
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

proc twapi::cert_create {subject pubkey cissuer args} {
    set args [_cert_create_parse_options $args opts]

    parseargs args {
        {encoding.arg pem {der pem}}
    } -maxleftover 0 -setvars
    
    # TBD - check that issuer is a CA

    set issuer_info [cert_info $cissuer]
    set issuer_blob [cert_name_to_blob [dict get $issuer_info -subject] -format x500]
    set sigalgo [dict get $issuer_info -signaturealgorithm]

    # If issuer cert has altnames, use they as issuer altnames for new cert
    set issuer_altnames [lindex [cert_extension $cissuer 2.5.29.17] 1]
    if {[llength $issuer_altnames]} {
        lappend opts(extensions) [_make_altnames_ext $issuer_altnames 0 1]
    }

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
    set subject_key_id [Twapi_HashPublicKeyInfo $pubkey]
    lappend opts(extensions) [list 2.5.29.14 0 $subject_key_id]

    set start [timelist_to_large_system_time $opts(start)]
    set end [timelist_to_large_system_time $opts(end)]

    # 2 -> CERT_V3
    # issuer_id and subject_id for the certificate are left empty
    # as recommended by gutman's X.509 paper
    set cert_info [list 2 $opts(serialnumber) $sigalgo $issuer_blob \
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
        return [_as_pem_or_der [CryptSignAndEncodeCertificate $hissuerprov \
                             $issuer_keyspec \
                             0x10001 2 $cert_info $sigalgo] \
                    CERTIFICATE $encoding]
    } finally {
        # TBD - test to make sure ok to close this if caller had
        # it open
        crypt_free $hissuerprov
    }
}

# TBD - test
proc twapi::cert_chain_build {hcert args} {
    # -timestamp not documented because not clear exactly how it behaves
    # -disablepass1*, -returnlower* not documented because not clear how
    # useful.
    # TBD - what about CERT_CHAIN_REVOCATION_ACCUMULATIVE_TIMEOUT 
    parseargs args {
        {cacheendcert.bool 0 0x1}
        {disableauthrootautoupdate.bool 0 0x100}
        {disablepass1qualityfiltering.bool 0 0x40}
        {engine.arg user {user machine}}
        {hstore.arg NULL}
        {returnlowerqualitycontexts.bool 0 0x80}
        {revocationcheck.arg all {none all leaf excluderoot}}
        {revocationcheckcacheonly.bool 0 0x80000000}
        {timestamp.arg ""}
        {urlretrievalcacheonly.bool 0 0x4}
        usageall.arg
        usageany.arg 
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
    } elseif {[info exists usageany]} {
        set usage [_get_enhkey_usage_oids $usageany]
    } else {
        set usage {}
    }

    return [CertGetCertificateChain \
                [dict* {user NULL machine {1 HCERTCHAINENGINE}} $engine] \
                $hcert $timestamp $hstore \
                [list [list $usage_op $usage]] $flags]
}

# TBD - document
proc twapi::cert_ancestors {hcert args} {
    # Note - does not care if certs are valid or not
    set certs {}
    set hchain [cert_chain_build $hcert {*}$args]
    trap {
        set simple_chain [twapi::Twapi_CertChainSimpleChain $hchain 0]
    } finally {
        cert_chain_release $hchain
    }
    foreach elem [dict get $simple_chain chain] {
        lappend certs [dict get $elem hcert]
    }
    return $certs
}

proc twapi::cert_chain_simple_chain {hchain index} {
    set simple_chain [twapi::Twapi_CertChainSimpleChain $hchain $index]
    set errors [_map_trust_error [dict get $simple_chain trust_errors]]
    dict set simple_chain trust_errors $errors
    if {[llength $errors]} {
        dict set simple_chain status fail
    } else {
        dict set simple_chain status ok
    }
    dict set simple_chain trust_info [_map_trust_info [dict get $simple_chain trust_info]]
    set chain_elements {}
    foreach elem [dict get $simple_chain chain] {
        set errors [_map_trust_error [dict get $elem trust_errors]]
        dict set elem trust_errors $errors
        if {[llength $errors]} {
            dict set elem status fail
        } else {
            dict set elem status ok
        }
        dict set elem trust_info [_map_trust_info [dict get $elem trust_info]]
        if {[dict exists $elem revocation]} {
            set revocation [dict get $elem revocation]
            if {$revocation == 0} {
                dict unset elem revocation
            } else {
                dict set elem revocation [_map_cert_verify_error $revocation]
            }
        }
        if {[dict exists $elem application_usage]} {
            dict set elem application_usage [_cert_decode_enhkey [dict get $elem application_usage]]
        }
        lappend chain_elements $elem
    }
    dict set simple_chain chain $chain_elements
    return $simple_chain
}

# TBD - test
proc twapi::cert_chain_trust_info {hchain} {
    return [_map_trust_info [Twapi_CertChainInfo $hchain]]
}

proc twapi::_map_trust_info {info} {
    return [_make_symbolic_bitmask $info {
        hasexactmatchissuer               0x00000001
        haskeymatchissuer                 0x00000002
        hasnamematchissuer                0x00000004
        isselfsigned                       0x00000008
        haspreferredissuer                 0x00000100
        hasissuancechainpolicy            0x00000200
        hasvalidnameconstraints           0x00000400
        ispeertrusted                      0x00000800
        hascrlvalidityextended            0x00001000
        isfromexclusivetruststore        0x00002000
        iscomplexchain                     0x00010000
    }]
}

# TBD - test
proc twapi::cert_chain_trust_errors {hchain} {
    return [_map_trust_error [Twapi_CertChainError $hchain]]
}

proc twapi::_map_trust_error {errbits} {
    return [_make_symbolic_bitmask $errbits {
        time 1
        revoked 4
        signature 8
        wrongusage 0x10
        untrustedroot 0x20
        revocationunknown 0x40
        trustcycle 0x80
        extension 0x100
        policy 0x200
        basiconstraints 0x400
        nameconstraints 0x800
        unsupportednameconstraint 0x1000
        undefinednameconstraint 0x2000
        unpermittednameconstraint 0x4000
        excludednameconstraint 0x8000
        revocationoffline 0x01000000
        noissuancechainpolicy 0x02000000
        distrust 0x04000000
        criticalextension 0x08000000
        weaksignature 0x00100000
        partialchain 0x00010000
        ctltime 0x00020000
        ctlsignature 0x00040000
        ctlusage 0x00080000
    }]
}
                            
proc twapi::cert_verify {hcert policy args} {
    # TBD - should we explicitly look for nulls in the subject name?
    # The Chrome source at
    # https://src.chromium.org/svn/branches/455/src/net/base/x509_certificate_win.cc
    # does this though it also uses the same calls as below. See
    # CertSubjectCommonNameHasNull in that code.
    set policy_id [dict! {
        authenticode 2   authenticodets 3   base 1   basicconstraints 5
        extendedvalidation 8   microsoftroot 7   ntauth 6
        ssl 4   tls 4
    } $policy]

    # Construct policy specific options
    set optdefs {
        {ignoreerrors.arg {}}
        policyparams.arg
        {trustedroots.arg}
    }
    switch -exact -- $policy_id {
        4 {
            # SSL/TLS
            lappend optdefs server.arg
        }
        5 {
            # basicconstraints
            lappend optdefs isa.arg
        }
        6 {
            # ntauth also accepts -isa as it includes basic constraints checks
            lappend optdefs isa.arg
        }
        7 {
            # microsoftroot
            lappend optdefs enabletestroot.bool
        }
    }

    array set opts [parseargs args $optdefs -ignoreunknown -setvars]

    if {![dict exists $args -usageall] && ![dict exists $args -usageany]} {
        switch -exact -- $policy {
            authenticodets -
            authenticode {
                dict lappend args -usageany code_signing
            }
            ssl -
            tls {
                if {[info exists server]} {
                    dict lappend args -usageany server_auth
                } else {
                    dict lappend args -usageany client_auth
                }
            }
        }
    }

    set verify_flags 0
    if {[info exists isa]} {
        switch -exact -- $isa {
            ca { set verify_flags [expr {$verify_flags | 0x80000000}] }
            endentity { set verify_flags [expr {$verify_flags | 0x40000000}] }
            default {
                error "Invalid value \"$isa\" specified for option -isa."
            }
        }
    }
    if {[info exists enabletestroot]} {
        set verify_flags [expr {$verify_flags | 0x00010000}]
    }
    
    if {$policy eq "basicconstraints"} {
        # TBD - peertrust 0x1000, see below
        set ignore_options {}
    } else {
        # Any other policy
        # TBD - the meaning of these is not clear. Are they ignore
        # error flags or options?
        #    peertrust        0x1000
        #    trusttestroot    0x4000
        #    allowtestroot    0x8000
        set ignore_options {
            time             0x07
            basicconstraints 0x08
            unknownca        0x10
            usage            0x20
            name             0x40
            policy           0x80
            revocation       0xf00
            criticalextensions 0x2000
        }
    }
    
    foreach ignore $ignoreerrors {
        if {![dict exists $ignore_options $ignore]} {
            error "Value $ignore for option -ignoreerrors cannot be used with policy $policy."
        }
        set verify_flags [expr {$verify_flags | [dict get $ignore_options $ignore]}]
    }

    if {![info exists policyparams]} {
        switch -exact -- $policy_id {
            4 {
                # ssl/tls
                if {[info exists server]} {
                    set policyparams [cert_policy_params_tls -ignoreerrors $ignoreerrors -server $server]
                } else {
                    set policyparams [cert_policy_params_tls -ignoreerrors $ignoreerrors]
                }
            }
            default {
                set policyparams {}
            }
        }
    }
    
    if {[info exists ignoreerrors] && "revocation" in $ignoreerrors} {
        lappend args -revocationcheck none
    }
    set chainh [cert_chain_build $hcert {*}$args]

    trap {
        # Actually verification is a bit tricky because the caller might
        # have asked for certain errors to be ignored. 
        # Note that CertVerifyChainPolicy below does NOT check for revocation 
        # of certificates in the certificate chain as per Microsoft docs.
        # We therefore check for revocation errors here and abort if present.
        set chain_errors [cert_chain_trust_errors $chainh]
        if {[llength $chain_errors]} {
            if {"revoked" in $chain_errors} {
                return revoked
            }
            if {"revocationoffline" in $chain_errors} {
                return revocationoffline
            }
            if {"revocationunknown" in $chain_errors} {
                return revocationunknown
            }

            if {0} {
                # For other kind of errors, caller might have indicated
                # some types are to be ignored. In that case we will proceed
                # to use CertVerifyTrustPolicy since that will allow
                # control of which errors are to be ignored. As a
                # special case, if caller has specified additional trusted
                # roots, we will proceed to call CertVerifyTrustPolicy
                # even when caller is not ignoring errors but only if
                # there are no errors indicated.
                if {[llength $chain_errors] > 1 ||
                    [lindex $chain_errors 0] ne "untrustedroot" ||
                    ![info exists trustedroots]} {
                    return $chain_errors
                }
            }
        }
        
        set status [Twapi_CertVerifyChainPolicy $policy_id $chainh [list $verify_flags $policyparams]]

        # If caller had provided additional trusted roots that are not
        # in the Windows trusted store, and the error is that the root is
        # untrusted, see if the root cert is one of the passed trusted ones
        # We will only deal when there is a single possible chain else
        # the recheck becomes very complicated as we are not sure if
        # the recheck will employ the same chain or not.
        if {$status == 0x800B0109 &&
            [info exists trustedroots] && [llength $trustedroots] &&
            [cert_chain_simple_chain_count $chainh] == 1} {
            set simple_chain [cert_chain_simple_chain $chainh 0]
            # Double check no errors listed for this chain
            set trust_errors [dict get $simple_chain trust_errors]
            if {[llength $trust_errors] == 1 &&
                [lindex $trust_errors 0] eq "untrustedroot"} {
                set certs_in_chain [dict get $simple_chain chain]
                set root_cert [dict get [lindex $certs_in_chain end] hcert]
                set thumbprint [cert_thumbprint $root_cert]
                # Match against each trusted root
                set trusted 0
                foreach trusted_cert $trustedroots {
                    if {$thumbprint eq [cert_thumbprint $trusted_cert]} {
                        set trusted 1
                        break
                    }
                }
                if {$trusted} {
                    # Yes, the root is trusted. It is not enough to
                    # say validation is ok because even if root
                    # is trusted, other errors might show up
                    # once untrusted roots are ignored. So we have
                    # to call the verification again.
                    # 0x10 -> CERT_CHAIN_POLICY_ALLOW_UNKNOWN_CA_FLAG
                    set verify_flags [expr {$verify_flags | 0x10}]
                    if {0} {
                        TBD - need to redo the policy params?
                        # 0x100 -> SECURITY_FLAG_IGNORE_UNKNOWN_CA
                        set checks [expr {$checks | 0x100}]
                    }
                    # Retry the call ignoring root errors
                    set status [Twapi_CertVerifyChainPolicy $policy_id $chainh [list $verify_flags $policyparams]]
                }
            }
        }

        return [_map_cert_verify_error $status]
    } finally {
        if {[info exists simple_chain]} {
            foreach cert [dict get $simple_chain chain] {
                cert_release [dict get $cert hcert]
            }
        }
        cert_chain_release $chainh
    }

    return $status
}

proc twapi::_map_cert_verify_error {err} {
    if {![string is integer -strict $err]} {
        return $err
    }
    return [dict*  {
        0x00000000 ok
        0x80096004 signature
        0x80092010 revoked
        0x800b0109 untrustedroot
        0x800b010d untrustedtestroot
        0x800b010a partialchain
        0x800b0110 wrongusage
        0x800b0101 time
        0x800b0114 name
        0x800b0113 policy
        0x80096019 basicconstraints
        0x800b0105 criticalextension
        0x800b0102 validityperiodnesting
        0x80092011 norevocationdll
        0x80092012 norevocationcheck
        0x80092013 revocationoffline
        0x800b010f cnmatch
        0x800b0106 purpose
        0x800b010e revocationunknown
        0x800b0103 carole
    } [format 0x%8.8x $err]]
}

# TBD - document
proc twapi::cert_policy_params_tls {args} {
    
    parseargs args {
        ignoreerrors.arg
        server.arg
    } -maxleftover 0 -setvars -ignoreunknown
    
    if {[info exists server]} {
        set role 2;         # AUTHTYPE_SERVER
    } else {
        set role 1;         # AUTHTYPE_CLIENT
        set server ""
    }

    set ignore_options {
        time             0x2000
        unknownca        0x100
        usage            0x200
        name             0x1000
        revocation       0x80
    }
    set checks 0
    foreach ignore $ignoreerrors {
        # Note we use dict*, not dict! so we can skip any ignore tokens
        # that we don't know
        set checks [expr {$checks | [dict* $ignore_options $ignore 0]}]
    }
    return [list $role $checks $server]
}

proc twapi::cert_tls_verify {hcert args} {
    return [cert_verify $hcert tls {*}$args]
}

proc twapi::cert_fetch {addr {port 443}} {
    set so [tls_socket $addr $port]
    trap {
        set sspi_ctx [chan configure $so -context]
        return [sspi_remote_cert $sspi_ctx]
    } finally {
        close $so
    }
}
                        
proc twapi::cert_locate_private_key {hcert args} {
    parseargs args {
        {keysettype.arg any {any user machine}}
        {silent 0 0x40}
    } -maxleftover 0 -setvars
    
    return [CryptFindCertificateKeyProvInfo $hcert \
                [expr {$silent | [dict get {any 0 user 1 machine 2} $keysettype]}]]
}

proc twapi::cert_request_parse {req args} {
    parseargs args {
        {encoding.arg {} {der pem {}}}
    } -setvars -maxleftover 0

    # 3 -> CRYPT_STRING_BASE64REQUESTHEADER 
    # 4 -> X509_CERT_REQUEST_TO_BE_SIGNED 
    lassign [::twapi::CryptDecodeObjectEx 4 [_pem_decode $req $encoding 3]] ver subject pubkey attrs
    lappend reqdict version $ver pubkey $pubkey attributes $attrs
    lappend reqdict subject [cert_blob_to_name $subject]
    foreach attr $attrs {
        lassign $attr oid values
        if {$oid eq "1.2.840.113549.1.9.14"} {
            # Extensions
            set extensions {}
            foreach ext [lindex $values 0] {
                lassign $ext oid critical value
                set value [_cert_decode_extension $oid $value]
                switch -exact -- $oid {
                    2.5.29.15 { set oidname -keyusage }
                    2.5.29.17 { set oidname -altnames }
                    2.5.29.19 { set oidname -basicconstraints }
                    2.5.29.37 { set oidname -enhkeyusage }
                    default { set oidname $oid }
                }
                lappend extensions $oidname [list $value $critical]
            }
            lappend reqdict extensions $extensions
        }
    }

    return $reqdict
}


proc twapi::cert_request_create {subject hprov keyspec args} {
    set args [_cert_create_parse_options $args opts]
    # TBD - barf if any elements other than extensions is set
    # TBD - document signaturealgorithmid
    parseargs args {
        {signaturealgorithmid.arg oid_rsa_sha1rsa}
        {encoding.arg pem {der pem}}
    } -setvars -maxleftover 0
    
    set sigoid [oid $signaturealgorithmid]
    if {$sigoid ni [list [oid oid_rsa_sha1rsa] [oid oid_rsa_md5rsa] [oid oid_x957_sha1dsa]]} {
        badargs! "Invalid signature algorithm '$sigalg'"
    }
    set keyspec [twapi::_crypt_keyspec $keyspec]
    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    # Pass oid_rsa_rsa as that seems to be what OPENSSL understands in
    # a CSR
    set pubkeyinfo [crypt_public_key $hprov $keyspec oid_rsa_rsa]
    set attrs [list 0 [cert_name_to_blob $subject] $pubkeyinfo]
    if {[llength $opts(extensions)]} {
        lappend attrs [list [list [oid oid_rsa_certextensions] [list $opts(extensions)]]]
    } else {
        lappend attrs {}
    }
    return [_as_pem_or_der [CryptSignAndEncodeCertificate $hprov $keyspec 0x10001 4 $attrs $sigoid] "NEW CERTIFICATE REQUEST" $encoding]
}


################################################################
# Cryptographic context commands

proc twapi::crypt_acquire {args} {
    # Backward compatibility - keycontainer can be specified as first arg
    if {[llength $args] &  1} {
        set args [lassign $args keycontainer]
    } else {
        set keycontainer ""
    }
    
    parseargs args {
        {csp.arg {}}
        {csptype.arg prov_rsa_full}
        keycontainer.arg
        {keysettype.arg user {user machine}}
        {create.bool 0 0x8}
        {silent.bool 0 0x40}
        verifycontext.bool
    } -maxleftover 0 -setvars

    # The defaults for verifycontext are a little confusing. For a named
    # key container, at least the MS CSP's require -verifycontext to be 0.
    # For the frequent case where private keys are not required, MS recommends
    # using the null key container with -verifycontext 1. So accordingly,
    # if the keycontainer is empty (or unspecified), then it 
    # defaults to 1, else defaults to 0.
    if {![info exists verifycontext]} {
        if {$keycontainer eq ""} {
            set verifycontext 1
        } else {
            set verifycontext 0
        }
    }

    if {$verifycontext} {
        set verifycontext 0xf0000000
    }

    set flags [expr {$silent | $verifycontext}]
    if {$keysettype eq "machine"} {
        incr flags 0x20;        # CRYPT_KEYSET_MACHINE
    }

    trap {
        return [CryptAcquireContext $keycontainer $csp [_csp_type_name_to_id $csptype] $flags]
    } onerror {TWAPI_WIN32 0x80090016} {
        # NTE_BAD_KEYSET - does not exist. Try to create it.
        if {$create} {
            set flags [expr {$flags | $create}]
            return [CryptAcquireContext $keycontainer $csp [_csp_type_name_to_id $csptype] $flags]            
        } else {
            rethrow
        }
    }
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

proc twapi::crypt_generate_key {hprov algid args} {

    array set opts [parseargs args {
        {archivable.bool 0 0x4000}
        {salt.bool 0 4}
        {exportable.bool 0 1}
        {pregen.bool 0x40}
        {userprotected.bool 0 2}
        {nosalt40.bool 0 0x10}
        {size.int 0}
    } -maxleftover 0]

    set algid [capi_algid $algid]

    if {$opts(size) < 0 || $opts(size) > 65535} {
        badargs! "Bad key size value '$size':  must be positive integer less than 65536"
    }

    return [CryptGenKey $hprov $algid [expr {($opts(size) << 16) | $opts(archivable) | $opts(salt) | $opts(exportable) | $opts(pregen) | $opts(userprotected) | $opts(nosalt40)}]]
}

proc twapi::crypt_keypair {hprov keyspec} {
    return [CryptGetUserKey $hprov [dict! {keyexchange 1 signature 2} $keyspec]]
}

proc twapi::crypt_public_key_import {hprov key args} {
    parseargs args {
        {algid.arg 0}
        {encoding.arg {} {native pem der {}}}
    } -setvars

    if {$encoding eq "native"} {
        set pub $key
    } elseif {$encoding eq "der"} {
        set pub [CryptDecodeObjectEx 8 $key]
    } elseif {$encoding eq "pem" ||
              ($encoding eq "" && [string match -nocase "-----BEGIN*" $key])} {
        set pub [CryptDecodeObjectEx 8 [CryptStringToBinary $key 0]]
    } else {
        # encoding is unspecified and is either der or native
        if {[catch {set pub [CryptDecodeObjectEx 8 $key]}]} {
            # Not DER, assume native
            set pub $key
        }
    }
    
    return [CryptImportPublicKeyInfoEx $hprov 0x10001 $pub [capi_algid $algid]]
}
                    
proc twapi::crypt_public_key_export {hprov keyspec args} {
    parseargs args {
        algoid.arg
        {encoding.arg pem {pem der native}}
    } -setvars -nulldefault
    
    if {$algoid ne ""} {
        set algoid [oid $algoid]
    }
    set pubkey [CryptExportPublicKeyInfoEx $hprov \
                    [_crypt_keyspec $keyspec] \
                    0x10001 \
                    $algoid \
                    0]
    if {$encoding eq "native"} {
        return $pubkey
    }
    # Generate SubjectPublicKeyInfo
    set der [CryptEncodeObjectEx 8 $pubkey]
    if {$encoding eq "der"} {
        return $der
    }
    # 0x80000001 -> No CR (only LF) and headers
    return "-----BEGIN PUBLIC KEY-----\n[CryptBinaryToString $der 0x80000001]-----END PUBLIC KEY-----\n"
}

# For back compat - undocumented
proc twapi::crypt_public_key {hcrypt algid oid} {
    return [crypt_public_key_export $hcrypt $algid -encoding native -algoid $oid]
}

proc twapi::crypt_get_security_descriptor {hprov} {
    return [CryptGetProvParam $hprov 8 7]
}

proc twapi::crypt_set_security_descriptor {hprov secd} {
    CryptSetProvParam $hprov 8 $secd
}

proc twapi::crypt_key_container_name {hprov} {
    return [CryptGetProvParam $hprov 6 0]
}

proc twapi::crypt_key_container_unique_name {hprov} {
    return [CryptGetProvParam $hprov 36 0]
}

proc twapi::crypt_csp {hprov} {
    return [CryptGetProvParam $hprov 4 0]
}

proc twapi::csps {} {
    set i 0
    set result {}
    while {[llength [set csp [::twapi::CryptEnumProviders $i]]]} {
        lappend result [lreplace $csp 0 0 [_csp_type_id_to_name [lindex $csp 0]]]
        incr i
    }
    return $result
}
interp alias {} twapi::crypt_csps {} twapi::csps

proc twapi::crypt_csp_type {hprov} {
    return [_csp_type_id_to_name [CryptGetProvParam $hprov 16 0]]
}

proc twapi::csp_types {} {
    set i 0
    set result {}
    while {[llength [set csptype [::twapi::CryptEnumProviderTypes $i]]]} {
        lappend result [lreplace $csptype 0 0 [_csp_type_id_to_name [lindex $csptype 0]]]
        incr i
    }
    return $result
}
interp alias {} twapi::crypt_csptypes {} twapi::csp_types

proc twapi::crypt_key_container_names {hcrypt} {
    return [CryptGetProvParam $hcrypt 2 0]
}

proc twapi::crypt_session_key_size {hcrypt} {
    return [CryptGetProvParam $hcrypt 20 0]
}

proc twapi::crypt_keyx_keysize_increment {hcrypt} {
    return [CryptGetProvParam $hcrypt 35 0]
}

proc twapi::crypt_sig_keysize_increment {hcrypt} {
    return [CryptGetProvParam $hcrypt 34 0]
}

# TBD - Doc and test
proc twapi::crypt_admin_pin {hcrypt} {
    return [CryptGetProvParam $hcrypt 31 0]
}

# TBD - Doc and test
proc twapi::crypt_keyx_pin {hcrypt} {
    return [CryptGetProvParam $hcrypt 32 0]
}

# TBD - Doc and test
proc twapi::crypt_sig_pin {hcrypt} {
    return [CryptGetProvParam $hcrypt 33 0]
}

proc twapi::crypt_csp_version {hcrypt} {
    set ver [CryptGetProvParam $hcrypt 5 0]
    return [format %d.%d [expr {($ver & 0xff00)>>8}] [expr {$ver & 0xff}]]
}

proc twapi::crypt_keyset_type {hcrypt} {
    return [expr {[CryptGetProvParam $hcrypt 27 0] & 0x20 ? "machine" : "user"}]
}

proc twapi::crypt_key_specifiers {hcrypt} {
    set keyspec [CryptGetProvParam $hcrypt 39 0]
    set keyspecs {}
    if {$keyspec & 1} {
        lappend keyspecs keyexchange
    }
    if {$keyspec & 2} {
        lappend keyspecs signature
    }
    return $keyspecs
}

proc twapi::crypt_symmetric_key_size {hcrypt} {
    return [CryptGetProvParam $hcrypt 19 0]
}

proc twapi::capi_key_export {hkey blob_type args} {
    parseargs args {
        {wrapper.arg NULL}
        {v3.bool 0 0x80}
        {oeap.bool 0 0x40}
        {destroy.bool 0 0x04}
    } -setvars -maxleftover 0

    return [CryptExportKey $hkey $wrapper [_capi_keyblob_type_id $blob_type] [expr {$v3|$oeap}]]
}
interp alias {} twapi::crypt_export_key {} twapi::capi_key_export


proc twapi::crypt_import_key {hcrypt keyblob args} {
    parseargs args {
        {wrapper.arg NULL}
        {exportable.bool 1 0x01}
        {oaep.bool 0 0x40}
        {userprotected.bool 0 0x02}
        {ipsechmac.bool 0 0x100}
    } -setvars -maxleftover 0
    return [CryptImportKey $hcrypt $keyblob $wrapper \
                [expr {$exportable|$oaep|$userprotected|$ipsechmac}]]
}
interp alias {} twapi::capi_key_import {} twapi::crypt_import_key

proc twapi::crypt_derive_key {hcrypt algid passphrase args} {
    parseargs args {
        {size.int 0}
        {exportable.bool 1 0x01}
        {prf.arg sha1}
        {method.arg pbkdf2}
        {iterations.int 100000}
        {salt.arg ""}
    } -maxleftover 0 -setvars

    if {$method eq "pbkdf2"} {
        set algnum [capi_algid $algid]
        if {$size == 0} {
            # Need to figure out the default key size for the algorithm
            # The loop below does not work for des/3des/3des_112 because
            # it will get the actual key size whereas CryptImportKey
            # wants key size with pad/parity bits. So hardcode these
            if {$algnum == 0x6601} {
                set size 64;    #
            } elseif {$algnum == 0x6603} {
                set size 192;   # 3des
            } elseif {$algnum == 0x6609} {
                set size 128;   # 3des_112
            } else {
                foreach alg [crypt_algorithms $hcrypt] {
                    if {[dict get $alg algid] == $algnum} {
                        set size [dict get $alg defkeylen]
                        break
                    }
                }
            }
            if {$size == 0} {
                error "Could not figure out default key size for algorithm $algid. Please use the -size option."
            }
        }
        set pbkdf2 [PBKDF2 $passphrase $size [capi_algid $prf] $salt $iterations]
        set keyblob [list 0 2 0 $algnum $pbkdf2]
        return [crypt_import_key $hcrypt $keyblob -exportable $exportable]
    } else {
        if {$size < 0 || $size > 65535} {
            # Key size of 0 is default. Else it must be within 1-65535
            badargs! "Option -size value \"$size\" is not between 0 and 65535."
        }
        set hhash [capi_hash_create $hcrypt [capi_algid $method]]
        twapi::trap {
            capi_hash_password $hhash $passphrase
            return [CryptDeriveKey $hcrypt [capi_algid $algid] $hhash \
                        [expr {($size << 16) | $exportable}]]
        } finally {
            capi_hash_free $hhash
        }
    }
}

proc twapi::pbkdf2 {pass nbits alg_id salt niters} {
    return [PBKDF2 $pass $nbits [capi_algid $alg_id] $salt $niters]
}
                    

proc twapi::capi_encrypt_bytes {bytes hkey args} {
    variable _capi_encrypt_partials
    parseargs args {
        {hhash.arg NULL}
        {final.bool 1}
        {pad.arg oaep {oaep pkcs1}}
    } -setvars -maxleftover 0
        
    if {[dict exists $_capi_encrypt_partials $hkey Data]} {
        append plaintext \
            [dict get $_capi_encrypt_partials $hkey Data] \
            $bytes
    } else {
        set plaintext $bytes
    }

    if {$final} {
        dict unset _capi_encrypt_partials $hkey
        return [CryptEncrypt $hkey $hhash $final [dict! {pkcs1 0 oaep 0x40} $pad] $plaintext]
    }

    # If not the final segment, we have to split it up into the block size multiple.
    if {[dict exists $_capi_encrypt_partials $hkey Blocklen]} {
        set blocklen [dict get $_capi_encrypt_partials $hkey Blocklen]
    } else {
        set blocklen [expr {[capi_key_blocklen $hkey] / 8}]; # Bits -> bytes
    }
    
    # len is largest multiple of block size less than data length
    set len [expr {([string length $plaintext] / $blocklen) * $blocklen}]
    set enc [CryptEncrypt $hkey $hhash $final [dict! {pkcs1 0 oaep 0x40} $pad] [string range $plaintext 0 $len-1]]
    # Note following will not happen if CryptEncrypt throws an error. As desired
    set remain [string range $plaintext $len end]
    if {[string length $remain]} {
        # Remember additional data 
        dict set _capi_encrypt_partials $hkey Data $remain
        dict set _capi_encrypt_partials $hkey Blocklen $blocklen
    } else {
        dict unset _capi_encrypt_partials $hkey
    }
    
    return $enc
}

proc twapi::capi_encrypt_string {s hkey args} {
    # Explicitly parse args, not just pass on because this command
    # does not support -final for symmetry with capi_decrypt_string
    parseargs args {
        {hhash.arg NULL}
        {pad.arg oaep {oaep pkcs1}}
    } -setvars -maxleftover 0
    return [capi_encrypt_bytes [encoding convertto utf-8 $s] $hkey -hhash $hhash -pad $pad]
}

proc twapi::capi_decrypt_bytes {bytes hkey args} {
    variable _capi_decrypt_partials
    parseargs args {
        {pad.arg oaep {oaep pkcs1 nopadcheck}}
        {final.bool 1}
        {hhash.arg NULL}
    } -setvars -maxleftover 0

    if {[dict exists $_capi_decrypt_partials $hkey Data]} {
        append enc \
            [dict get $_capi_decrypt_partials $hkey Data] \
            $bytes
    } else {
        set enc $bytes
    }

    if {$final} {
        dict unset _capi_decrypt_partials $hkey
        return [CryptDecrypt $hkey $hhash $final [dict! {pkcs1 0 oaep 0x40 nopadcheck 0x20} $pad] $enc]
    }

    # If not the final segment, we have to split it up into the block size multiple.
    if {[dict exists $_capi_decrypt_partials $hkey Blocklen]} {
        set blocklen [dict get $_capi_decrypt_partials $hkey Blocklen]
    } else {
        set blocklen [expr {[capi_key_blocklen $hkey] / 8}]; # Bits -> bytes
    }

    # len is largest multiple of block size less than data length
    set len [expr {([string length $enc] / $blocklen) * $blocklen}]
    set plaintext [CryptDecrypt $hkey $hhash $final [dict! {pkcs1 0 oaep 0x40} $pad] [string range $enc 0 $len-1]]
    # Note following will not happen if CryptDecrypt throws an error. As desired
    set remain [string range $enc $len end]
    if {[string length $remain]} {
        # Remember additional data 
        dict set _capi_decrypt_partials $hkey Data $remain
        dict set _capi_decrypt_partials $hkey Blocklen $blocklen
    } else {
        dict unset _capi_decrypt_partials $hkey
    }

    return $plaintext
}

proc twapi::capi_decrypt_string {s hkey args} {
    # Explicitly parse args, not just pass on because this command
    # does not support -final for symmetry with capi_decrypt_string
    parseargs args {
        {hhash.arg NULL}
        {pad.arg oaep {oaep pkcs1}}
    } -setvars -maxleftover 0
    return [encoding convertfrom utf-8 [capi_decrypt_bytes $s $hkey -hhash $hhash -pad $pad]]
}

# Returns the most capable CSP
proc twapi::_crypt_acquire_default {} {
    if {[catch {crypt_acquire -csptype prov_rsa_aes} hcrypt] &&
        [catch {crypt_acquire -csptype prov_rsa_full -csp {Microsoft Enhanced Cryptographic Provider v1.0}} hcrypt]} {
        set hcrypt [crypt_acquire]
    }
    set cspname [crypt_csp $hcrypt]
    set csptype [crypt_csp_type $hcrypt]
    # Redefine ourselves for next call
    proc [namespace current]::_crypt_acquire_default {} "crypt_acquire -csp {$cspname} -csptype $csptype"
    return $hcrypt
}

proc twapi::_block_cipher {algo direction bytes keybytes args} {

    # Note: padding mode is not documented since MS providers only support
    # one mode anyway
    parseargs args {
        mode.arg
        iv.arg
        padding.arg
    } -setvars -maxleftover 0
    
    set hcrypt [_crypt_acquire_default]
    try {
        set hkey [crypt_import_key $hcrypt [capi_keyblob_concealed $algo $keybytes]]
        if {[info exists mode]} {
            capi_key_mode $hkey $mode
        }
        if {[info exists iv]} {
            capi_key_iv $hkey $iv
        }
        if {$direction eq "encrypt"} {
            if {[info exists padding]} {
                capi_key_padding $hkey $padding
            }
            set ciphertext [capi_encrypt_bytes $bytes $hkey]
        } else {
            set ciphertext [capi_decrypt_bytes $bytes $hkey]
        }
    } finally {
        if {[info exists hkey]} {
            capi_key_free $hkey
        }
        crypt_free $hcrypt
    }
    return $ciphertext
}

# apply to avoid global variable pollution
apply {{} {
    foreach {algo blocklen} {des 8 3des 8 aes_128 16 aes_192 16 aes_256 16} {
        namespace eval twapi::$algo {}
        interp alias {} twapi::${algo}::encrypt {} twapi::_block_cipher $algo encrypt
        interp alias {} twapi::${algo}::decrypt {} twapi::_block_cipher $algo decrypt
        interp alias {} twapi::${algo}::iv {} twapi::random_bytes $blocklen
        namespace eval twapi::$algo {
            namespace export encrypt decrypt iv
            namespace ensemble create
        }
    }
}}

###
# PKCS7 commands

proc twapi::pkcs7_encrypt {bytes recipients encalg args} {
    parseargs args {
        {encoding.arg pem {pem der}}
        {innertype.arg 0}
    } -setvars -maxleftover 0

    # TBD - add support for the following
    set flags 0
    set encauxinfo {}

    set params [list \
                    0x10001 \
                    NULL \
                    [_make_algorithm_identifier $encalg] \
                    $encauxinfo \
                    $flags \
                    $innertype]
    return [_as_pem_or_der [CryptEncryptMessage $params $recipients $bytes] PKCS7 $encoding]
}

proc twapi::pkcs7_decrypt {bytes stores args} {
    parseargs args {
        {encoding.arg {} {der pem {}}}
        {silent.bool 0 0x40}
        {certvar.arg ""}
    } -maxleftover 0 -setvars

    set params [list \
                    0x10001 \
                    $stores \
                    $silent]
    if {$certvar ne ""} {
        upvar 1 $certvar hcert
        set certvar hcert
    }
    
    return [CryptDecryptMessage $params [_pem_decode $bytes $encoding] $certvar]
}
 
proc twapi::pkcs7_sign {bytes hcert hashalg args} {
    # TBD - document crls?
    parseargs args {
        {detached.bool 0}
        {encoding.arg pem {pem der}}
        {includecerts.arg all {none leaf all}}
        {silent.bool 0 0x40}
        {usesignerkeyid.bool 0 0x4}
        {crls.arg {}}
        {innercontenttype.arg 0}
    } -setvars -maxleftover 0

    set flags [expr {$usesignerkeyid | $silent}]

    switch -exact -- $includecerts {
        leaf { set certs [list [cert_duplicate $hcert]] }
        none { set certs {} }
        all { set certs [cert_ancestors $hcert] }
    }
    # TBD - add support for the following
    set hashaux {}
    set authattrs {}
    set unauthattrs {}
    set encalg ""
    set hashencaux ""
    # 0x10001 -> PKCS_7_ASN_ENCODING|X509_ASN_ENCODING
    set params [list \
                    0x10001 \
                    $hcert \
                    [_make_algorithm_identifier $hashalg] \
                    $hashaux \
                    $certs \
                    $crls \
                    $authattrs \
                    $unauthattrs \
                    $flags \
                    $innercontenttype \
                    $encalg \
                    $hashencaux]
    trap {
        return [_as_pem_or_der [CryptSignMessage $params $detached [list $bytes]] PKCS7 $encoding]
    } finally {
        foreach c $certs {
            cert_release $c
        }
    }
}

proc twapi::pkcs7_verify {bytes args} {
    parseargs args {
        {encoding.arg {} {der pem {}}}
        {contentvar.arg ""}
        {certvar.arg ""}
    } -maxleftover 0 -setvars -ignoreunknown

    if {$contentvar ne ""} {
        upvar 1 $contentvar content
        set contentvar content
    }
    set status [CryptVerifyMessageSignature [list 0x10001 NULL] 0 [_pem_decode $bytes $encoding] $contentvar hcert]
    if {$status == 0} {
        trap {
            set status [cert_verify $hcert base {*}$args]
            if {$status eq "ok"} {
                if {$certvar ne ""} {
                    upvar 1 $certvar cert
                    set cert $hcert
                    unset hcert;    # So we do not release it below
                }
                if {$contentvar ne ""} {
                    upvar 1 $contentvar con
                    set con content
                }
            }
        } finally {
            if {[info exists hcert]} {
                cert_release $hcert
            }
        }
    } else {
        # Note these codes are different from those in _map_cert_verify_error
        if {$status == 0x80090006} {
            set status "signature"
        } elseif {$status == 0x80090008} {
            set status "invalidalgorithm"
        }
    }
        
    return $status
}


# For backwards compat - deprecated
interp alias {} twapi::crypt_key_free {} twapi::capi_key_free

proc twapi::crypt_algorithms {hcrypt} {
    set algs {}
    foreach alg [CryptGetProvParam $hcrypt 22 0] {
        lassign $alg algid defaultlen minlen maxlen protos name description
        set protos [_make_symbolic_bitmask $protos {
            ipsec 0x10 pct1 0x01 signing 0x20 ssl2 0x02 ssl3 0x04 tls1 0x08
        }]
        lappend algs [list algid $algid defkeylen $defaultlen minkeylen $minlen maxkeylen $maxlen protocols $protos name $name description $description]
    }
    return $algs
}

proc twapi::crypt_implementation_type {hcrypt} {
    return [dict* {1 hardware 2 software 3 mixed 4 unknown 8 removable} [CryptGetProvParam $hcrypt 3 0]]
}

proc twapi::capi_algid {s} {
    if {[string is integer -strict $s]} {
        return [expr {$s}];     # Return in decimal form
    }
    set algid [dict* {
        3des 0x00006603
        3des_112 0x00006609
        aes 0x00006611
        aes_128 0x0000660e
        aes_192 0x0000660f
        aes_256 0x00006610
        agreedkey_any 0x0000aa03
        keyexchange 1
        signature 2
        cylink_mek 0x0000660c
        des 0x00006601
        desx 0x00006604
        dh_ephem 0x0000aa02
        dh_sf 0x0000aa01
        dss_sign 0x00002200
        ecdh 0x0000aa05
        ecdsa 0x00002203
        ecmqv 0x0000a001
        hash_replace_owf 0x0000800b
        hughes_md5 0x0000a003
        hmac 0x00008009
        kea_keyx 0x0000aa04
        mac 0x00008005
        md2 0x00008001
        md4 0x00008002
        md5 0x00008003
        no_sign 0x00002000
        pct1_master 0x00004c04
        rc2 0x00006602
        rc4 0x00006801
        rc5 0x0000660d
        rsa_keyx 0x0000a400
        rsa_sign 0x00002400
        schannel_enc_key 0x00004c07
        schannel_mac_key 0x00004c03
        schannel_master_hash 0x00004c02
        sha 0x00008004
        sha1 0x00008004
        sha_256 0x0000800c
        sha_384 0x0000800d
        sha_512 0x0000800e
        ssl2_master 0x00004c05
        ssl3_master 0x00004c01
        ssl3_shamd5 0x00008008
        tls1_master 0x00004c06
        tls1prf 0x0000800a
    } $s ""]

    if {$algid eq ""} {
        set oid [oid $s]
        set algid [CertOIDToAlgId $oid]
        if {$algid == 0} {
            error "Could not map \"$s\" to algorithm id"
        }
    }
    # Return the decimal form
    return [expr {$algid}]
}

# TBD - document
proc twapi::crypt_find_oid_info {key args} {
    array set opts [parseargs args {
        {restrict.arg any {sign encrypt any}}
        keylen.int
        {searchds.bool 0}
        {oidgroup.arg  0}
    } -maxleftover 0]

    # We will try key to be an OID, Alg Id, sign id or a simple
    # name in turn
    if {[catch {
        set key [oid $key]
        set keytype 1;          # OID
    }]} {
        if {[catch {
            set key [capi_algid $key]
            set keytype 3;      # Alg Id
        }]} {
            if {[catch {
                # Sign - list of two alg id's
                if {[llength $key] == 2} {
                    set key [list [capi_algid [lindex $key 0]] [capi_algid [lindex $key 1]]]
                    set keytype 4
                } else {
                    set keytype 2 ;# Name
                }
            }]} {
                set keytype 2 ;# Name
            }
        }
    }

    set oidgroup [oidgroup $opts(oidgroup)]
    if {$opts(restrict) ne "any"} {
        if {$oidgroup != 0 && $oidgroup != 3} {
            error "The -restrict option can only be used with the oidgroup_pubkey_alg OID group"
        }
        if {$opts(restrict) eq "sign"} {
            set keytype [expr {$keytype | 0x80000000}]
        } else {
            set keytype [expr {$keytype | 0x40000000}]
        }
    }
    
    if {[info exists opts(keylen)]} {
        set oidgroup [expr {$oidgroup | ($opts(keylen) << 16)}]
    }

    # Because search of active dir can be slow, turn it off unless
    # caller explicitly requests it
    if {! $opts(searchds)} {
        set oidgroup [expr {$oidgroup | 0x80000000}]
    }

    return [CryptFindOIDInfo $keytype $key $oidgroup]
}

# TBD - document
proc twapi::crypt_enumerate_oid_info {{oidgroup 0}} {
    # TBD - parse extra based on OID group
    set ret {}
    foreach info [CryptEnumOIDInfo [oidgroup $oidgroup]] {
        lappend ret [twine {oid name oidgroup value extra} $info]
    }
    return $ret
}

# TBD - test
proc twapi::_capi_parse {type arg args} {
    parseargs args {
        {contenttype.arg any}
        {formattype.arg any}
        {typesonly.bool 0}
    } -setvars -maxleftover 0
        
    # First try the formats not supported by CryptQueryObject
    if {$contenttype in {any rsapublickey subjectpublickeyinfo}} {
        if {$formattype eq "binary"} {
            set encoding der
        } elseif {$formattype eq "base64"} {
            set encoding pem
        } else {
            set encoding ""
        }
        if {$type == 1} {
            # arg is a file
            set fd [open $arg]
            trap {
                fconfigure $fd -translation binary
                set content [_pem_decode [read $fd] $encoding]
                set is_pem  [_is_pem $content]
            } finally {
                close $fd
            }
        }
        if {$contenttype in {any subjectpublickeyinfo}} {
            trap {
                set data [CryptDecodeObjectEx 8 $content]
                dict set ret contenttype subjectpublickeyinfo
                dict set ret formattype [lindex {binary base64} $is_pem]
                if {! $typesonly} {
                    dict set ret subjectpublickeyinfo $data
                }
                return $ret
            } onerror {} {
                if {$contenttype eq "subjectpublickeyinfo"} {
                    rethrow
                }
                # Go on to try other types
            }
        }
        if {$contenttype in {any rsapublickey}} {
            trap {
                set data [CryptDecodeObjectEx 19 $content]
                dict set ret contenttype rsapublickey
                dict set ret formattype [lindex {binary base64} $is_pem]
                if {! $typesonly} {
                    dict set ret rsapublickey $data
                }
                return $ret
            } onerror {} {
                if {$contenttype eq "rsapublickey"} {
                    rethrow
                }
                # Go on to try other types
            }
        }
    }

    # No joy. Go on to try CryptQueryObject

    # Note - CERT_QUERY_CONTENT_FLAG_PFX_AND_LOAD not supported
    # on XP/2k3 hence not included in expected_content_type
    set contenttype [dict! {
        cert 2
        ctl  4
        crl  8
        serializedstore 16
        serializedcert 32
        serializedctl 64
        serializedcrl 128
        pkcs7signed   256
        pkcs7unsigned 512
        pkcs7signedembed 1024
        pkcs10 2048
        pfx 4096
        certpair 8192
        any 0x3FFE
    } $contenttype]
    
    set formattype [dict! {
        binary 2
        base64 4
        asn1hex 8
        any    14
    } $formattype]

    set ret [CryptQueryObject $type $arg \
                 $contenttype $formattype 0 $typesonly]
    # We don't mention PKCS7_ASN v/s X509_ASN anywhere and use encoding
    # to refer to PEM/DER so leave it off for now
    dict unset ret encoding
    dict set ret formattype [dict* {
        1 binary
        2 base64
        3 asn1hex
    } [dict get $ret formattype]]
    dict set ret contenttype [dict* {
        1 cert
        2 ctl 
        3 crl 
        4 serializedstore 
        5 serializedcert 
        6 serializedctl 
        7 serializedcrl
        8 pkcs7signed  
        9 pkcs7unsigned
        10 pkcs7signedembed 
        11 pkcs10 
        12 pfx 
        13 certpair 
    } [dict get $ret contenttype]]

    return $ret
}
interp alias {} twapi::capi_parse_file {} twapi::_capi_parse 1
interp alias {} twapi::capi_parse {} twapi::_capi_parse 2

###
# ASN.1 procs

# TBD - document
proc twapi::asn1_decode_string {bin} {
    # 24 -> X509_UNICODE_ANY_STRING
    return [lindex [twapi::CryptDecodeObjectEx 24 $bin] 1]
}

# TBD - document
proc twapi::asn1_encode_string {s {encformat utf8}} {
    # 24 -> X509_UNICODE_ANY_STRING
    return [twapi::CryptEncodeObjectEx 24 [list [dict! {
        numeric 3 printable 4 teletex 5 t61 5 videotex 6 ia5 7 graphic 8
        visible 9 iso646 9 general 10 universal 11 int4 11
        bmp 12 unicode 12 utf8 13
    } $encformat] $s]]
}

###
# Key procs

proc twapi::_capi_key_param {param_id hkey args} {
    if {[llength $args] == 0} {
        return [CryptGetKeyParam $hkey $param_id]
    }
    if {[llength $args] == 1} {
        return [CryptSetKeyParam $hkey $param_id [lindex $args 0]]
    }
    badargs! "Invalid syntax. Should be [lindex [info level -1] 0] HKEY ?VALUE?" 3
}

proc twapi::capi_key_iv {args} {return [_capi_key_param 1 {*}$args]} 
proc twapi::capi_key_mode_bits {args} {return [_capi_key_param 5 {*}$args]} 
proc twapi::capi_key_dss_p {args} {return [_capi_key_param 11 {*}$args]} 
proc twapi::capi_key_dss_q {args} {return [_capi_key_param 13 {*}$args]} 
proc twapi::capi_key_dss_g {args} {return [_capi_key_param 12 {*}$args]} 
proc twapi::capi_key_effective_keylen {args} {return [_capi_key_param 19 {*}$args]} 

proc twapi::capi_key_blocklen {hkey} {return [CryptGetKeyParam $hkey 8]}
proc twapi::capi_key_certificate {hkey} {return [CryptGetKeyParam $hkey 26]}
proc twapi::capi_key_keylen {hkey} {return [CryptGetKeyParam $hkey 9]}

proc twapi::capi_key_algid {hkey args} {
    if {[llength $args] == 0} {
        return [CryptGetKeyParam $hkey 7]
    }
    set args [lassign $args algid]
    set algid [capi_algid $algid]
    array set opts [parseargs args {
        {archivable.bool 0 0x4000}
        {salt.bool 0 4}
        {exportable.bool 0 1}
        {pregen.bool 0x40}
        {userprotected.bool 0 2}
        {nosalt40.bool 0 0x10}
        {size.int 0}
    } -maxleftover 0]
    if {$opts(size) < 0 || $opts(size) > 65535} {
        badargs! "Bad key size value '$size':  must be positive integer less than 65536"
    }
    set flags [expr {($opts(size) << 16) | $opts(archivable) | $opts(salt) | $opts(exportable) | $opts(pregen) | $opts(userprotected) | $opts(nosalt40)}]
    return [CryptSetKeyParam $hkey 7 $algid $flags]
}

proc twapi::capi_key_mode {hkey args} {
    if {[llength $args] == 0} {
        return [dict* {1 cbc 2 ecb 3 ofb 4 cfb 5 cts} [CryptGetKeyParam $hkey 4]]
    }
    if {[llength $args] == 1} {
        set val [dict* {cbc 1 ecb 2 ofb 3 cfb 4 cts 5} [lindex $args 0]]
        return [CryptSetKeyParam $hkey 4 $val]
    }
    badargs! "Invalid syntax. Should be [lindex [info level 0] 0] HKEY ?VALUE?"
}

proc twapi::capi_key_padding {hkey args} {
    if {[llength $args] == 0} {
        return [dict* {1 pkcs5 2 random 3 zeroes} [CryptGetKeyParam $hkey 3]]
    }
    if {[llength $args] == 1} {
        set val [dict* {pkcs5 1 random 2 zeroes 3} [lindex $args 0]]
        return [CryptSetKeyParam $hkey 3 $val]
    }
    badargs! "Invalid syntax. Should be [lindex [info level 0] 0] HKEY ?VALUE?"
}

proc twapi::capi_key_permissions {hkey args} {
    set bitmasks {
        encrypt 0x01 decrypt 0x02 export 0x04 read 0x08 write 0x10
        mac 0x20 export_key 0x40 import_key 0x80 archive 0x100
    }
    if {[llength $args] == 0} {
        return [_make_symbolic_bitmask [CryptGetKeyParam $hkey 6] $bitmasks]
    }
    if {[llength $args] == 1} {
        set val [_parse_symbolic_bitmask [lindex $args 0] $bitmasks]
        return [CryptSetKeyParam $hkey 6 $val]
    }
    badargs! "Invalid syntax. Should be [lindex [info level 0] 0] HKEY ?VALUE?"
}

proc twapi::capi_key_salt {hkey args} {
    if {[llength $args] == 0} {
        # 2 -> KP_SALT
        return [CryptGetKeyParam $hkey 2]
    }
    if {[llength $args] == 1} {
        # 10 -> KP_SALT_EX
        return [CryptSetKeyParam $hkey 10 [lindex $args 0]]
    } 
    badargs! "Invalid syntax. Should be [lindex [info level 0] 0] HKEY ?VALUE?"
}

proc twapi::capi_keyblob_create {ver algid blob_type key} {
    # 0 -> reserved field
    return [list [_capi_keyblob_type_id $blob_type] $ver 0 [capi_algid $algid] $key]
}

proc twapi::capi_keyblob_concealed {algid concealed_key} {
    # 2 -> bVersion
    # 0 -> concealed plaintextkeyblob
    # Note: for our own home grown concealed type there is no
    # BLOBHEADER
    return [capi_keyblob_create 2 $algid concealed $concealed_key]
}

proc twapi::capi_keyblob_plaintext {algid binkey} {
    # typedef struct _PUBLICKEYSTRUC {
    #   BYTE   bType;
    #   BYTE   bVersion;
    #   WORD   reserved;
    #   ALG_ID aiKeyAlg;
    # } BLOBHEADER;
    # 2 -> bVersion
    set algnum [capi_algid $algid]
    set blob_type [_capi_keyblob_type_id plaintext]
    set len [string length $binkey]
    set blob "[binary format ccsii $blob_type 2 0 $algnum $len]$binkey"
    return [capi_keyblob_create 2 $algid plaintext $blob]
}

proc twapi::capi_keyblob_version {kblob} {
    return [lindex $kblob 1]
}

proc twapi::capi_keyblob_algid {kblob} {
    return [lindex $kblob 3]
}

proc twapi::capi_keyblob_type {kblob} {
    return [_capi_keyblob_type_name [lindex $kblob 0]]
}

proc twapi::capi_keyblob_blob {kblob} {
    return [lindex $kblob 4]
}

proc twapi::_capi_keyblob_type_id {name} {
    set blob_type [dict* {
        concealed    0
        keystate    12
        opaque       9
        plaintext    8
        privatekey   7
        publickey    6
        publickeyex 10
        rfc3217     11
        simple       1
    } $name]
}

proc twapi::_capi_keyblob_type_name {id} {
    set blob_type [dict* {
        0  concealed
        1  simple
        6  publickey
        7  privatekey
        8  plaintext
        9  opaque
        10 publickeyex
        11 rfc3217
        12 keystate
    } [incr id 0]];              # incr to convert hex etc. to decimal
    
}

###
# Utility procs

proc twapi::_make_algorithm_identifier {oid {param {}}} {
    if {[string length $oid] == 0} {
        return ""
    }
    if {0} {
        # TBD - what modes to default to ?
        switch -exact -- $oid {
#define szOID_NIST_AES128_CBC        "2.16.840.1.101.3.4.1.2"
#define szOID_NIST_AES192_CBC        "2.16.840.1.101.3.4.1.22"
#define szOID_NIST_AES256_CBC        "2.16.840.1.101.3.4.1.42"

#// For the above Algorithms, the AlgorithmIdentifier parameters must be
#// present and the parameters field MUST contain an AES-IV:
#//
#//  AES-IV ::= OCTET STRING (SIZE(16))

#// NIST AES WRAP Algorithms
#define szOID_NIST_AES128_WRAP       "2.16.840.1.101.3.4.1.5"
#define szOID_NIST_AES192_WRAP       "2.16.840.1.101.3.4.1.25"
#define szOID_NIST_AES256_WRAP       "2.16.840.1.101.3.4.1.45"
            des { set oid "oid_rsa_des_ede3_cbc" }
            des { set oid "oid_oiwsec_descbc" }
            aes128 { TBD }
            aes192 { TBD }
            aes256 { TBD }
            rc2 { set oid "oid_rsa_rc2cbc" }
            rc4 { set oid "oid_rsa_rc4" }
        }
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

twapi::proc* twapi::_system_store_id {name} {
    variable _system_store_locations
    
    set _system_store_locations {
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

    foreach loc [CertEnumSystemStoreLocation 0] {
        dict set _system_store_locations {*}$loc
    }
} {
    variable _system_store_locations

    if {[string is integer -strict $name]} {
        if {$name < 65536} {
            badargs! "Invalid system store name $name" 3
        }
        return $name
    }

    return [dict! $_system_store_locations $name 2]
}

twapi::proc* twapi::_csp_type_name_to_id prov {
    variable _csp_type_name_id_map

    array set _csp_type_name_id_map {
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
    variable _csp_type_name_id_map

    set key [string tolower $prov]

    if {[info exists _csp_type_name_id_map($key)]} {
        return $_csp_type_name_id_map($key)
    }

    if {[string is integer -strict $prov]} {
        return $prov
    }

    badargs! "Invalid or unknown provider type '$prov'" 3
}

twapi::proc* twapi::_csp_type_id_to_name prov {
    variable _csp_type_name_id_map
    variable _csp_id_type_name_map

    _csp_type_name_to_id prov_rsa_full; # Just to ensure _csp_type_name_id_map exists
    array set _csp_id_type_name_map [swapl [array get _csp_type_name_id_map]]
} {
    variable _csp_id_type_name_map
    if {[info exists _csp_id_type_name_map($prov)]} {
        return $_csp_id_type_name_map($prov)
    }

    if {[string is integer -strict $prov]} {
        return $prov
    }

    badargs! "Invalid or unknown CSP type id '$prov'" 3
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
    if {[regexp {^\d+\.\d+(\.\d+)*$} $name]} {
        return $name;           # OID literal n.n...
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

# TBD - change OID mnemonics to those in RFC (see pki.tcl in tcllib)
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

# TBD - document
proc twapi::oidgroup {oidgroup} {
    if {[string is integer -strict $oidgroup]} {
        return $oidgroup
    }
    return [dict! {
        oidgroup_hash_alg             1
        oidgroup_encrypt_alg          2
        oidgroup_pubkey_alg           3
        oidgroup_sign_alg             4
        oidgroup_rdn_attr             5
        oidgroup_ext_or_attr          6
        oidgroup_enhkey_usage         7
        oidgroup_policy               8
        oidgroup_template             9
    } $oidgroup]
}

# TBD - document
proc twapi::oidgroup_token {oidgroup} {
    return [lindex {
        {}
        oidgroup_hash_alg
        oidgroup_encrypt_alg
        oidgroup_pubkey_alg
        oidgroup_sign_alg
        oidgroup_rdn_attr
        oidgroup_ext_or_attr
        oidgroup_enhkey_usage
        oidgroup_policy
        oidgroup_template
    } $oidgroup]
}

proc twapi::_make_altnames_ext {altnames {critical 0} {issuer 0}} {
    set names {}
    foreach pair $altnames {
        lassign $pair alttype altname
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

proc twapi::_make_enhkeyusage_ext {enhkeyusage {critical 0}} {
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

proc twapi::_make_basic_constraints_ext {basicconstraints {critical 1}} {
    lassign $basicconstraints isca capathlenvalid capathlen
    if {[string is boolean $isca] && [string is boolean $capathlenvalid] &&
        [string is integer -strict $capathlen] && $capathlen >= 0} {
        return [list "2.5.29.19" $critical [list $isca $capathlenvalid $capathlen]]
    }
    error "Invalid basicconstraints value"
}

proc twapi::_make_keyusage_ext {keyusage {critical 0}} {
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
    set symmap [swapl [oids oid_pkix_kp_*]]
    foreach val $vals {
        if {[dict exists $symmap $val]} {
            lappend result [string range [dict get $symmap $val] 12 end]
        } else {
            lappend result $val
        }
    }
    return $result
}

proc twapi::_cert_decode_extension {oid val} {
    # TBD - see what other types need to be decoded
    # 2.5.29.19 - basic constraints
    # 
    switch $oid {
        2.5.29.15 { return [_cert_decode_keyusage $val] }
        2.5.29.37 { return [_cert_decode_enhkey $val] }
        2.5.29.17 -
        2.5.29.18 {
            set names {}
            foreach elem $val {
                lappend names [list [dict* {
                    1 other 2 email 3 dns 5 directory 7 url 8 ip 9 registered
                } [lindex $elem 0]] [lindex $elem 1]]
            }
            return $names
        }
    }
    return $val
}

proc twapi::_crypt_keyspec {keyspec} {
    return [dict* {keyexchange 1 signature 2} $keyspec]
}

proc twapi::_cert_create_parse_options {optvals optsvar} {
    upvar 1 $optsvar opts

    # TBD - add -issueraltnames
    parseargs optvals {
        start.arg
        end.arg
        serialnumber.arg
        altnames.arg
        enhkeyusage.arg
        keyusage.arg
        basicconstraints.arg
        {purpose.arg {}}
        {capathlen.int -1}
    } -ignoreunknown -setvars

    set ca [expr {"ca" in $purpose}]
    if {$ca} {
        if {[info exists basicconstraints]} {
            badargs! "Option -basicconstraints cannot be specified if \"ca\" is included in the -purpose option"
        }
        if {$capathlen < 0} {
            set basicconstraints {{1 0 0} 1};  # No path length constraint
        } else {
            set basicconstraints [list [list 1 1 $capathlen] 1]
        }
    } else {
        if {![info exists basicconstraints]} {
            set basicconstraints {{0 0 0} 1}
        }
    }
    set sslserver [expr {"server" in $purpose}]
    set sslclient [expr {"client" in $purpose}]

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
    if {[info exists start]} {
        set opts(start) $start
    } else {
        set opts(start) [_seconds_to_timelist [clock seconds] 1]
    }
    if {[info exists end]} {
        set opts(end) $end
    } else {
        set opts(end) $opts(start)
        lset opts(end) 0 [expr {[lindex $opts(end) 0] + 1}]
        # Ensure valid date (Feb 29 leap year -> non-leap year for example)
        set opts(end) [clock format [clock scan [lrange $opts(end) 0 2] -format "%Y %N %e"] -format "%Y %N %e"]
        lappend opts(end) 23 59 59 0
    }

    # Generate the extensions list
    set exts {}
    lappend exts [_make_basic_constraints_ext {*}$basicconstraints ]
    if {$ca} {
        lappend extra_keyusage key_cert_sign crl_sign
    }
    if {$sslserver || $sslclient} {
        # TBD - not clear key_agreement is needed for SSL certs for
        # either client or server. See https://access.redhat.com/documentation/en-US/Red_Hat_Certificate_System/8.0/html/Admin_Guide/Standard_X.509_v3_Certificate_Extensions.html#Discussion-Certificate_Uses_and_Corresponding_Key_Usage_Bits
        lappend extra_keyusage digital_signature key_encipherment key_agreement
        if {$sslserver} { 
           lappend extra_enhkeyusage oid_pkix_kp_server_auth
        }
        if {$sslclient} {
            lappend extra_enhkeyusage oid_pkix_kp_client_auth
        }
    }

    if {[info exists extra_keyusage]} {
        if {[info exists keyusage]} {
            # TBD - should it be marked critical or not ?
            lset keyusage 0 [concat [lindex $keyusage 0] $extra_keyusage]
        } else {
            # TBD - should it be marked critical or not ?
            set keyusage [list $extra_keyusage 1]
        }
    }

    if {[info exists keyusage]} {
        lappend exts [_make_keyusage_ext {*}$keyusage]
    }

    if {[info exists extra_enhkeyusage]} {
        if {[info exists enhkeyusage]} {
            # TBD - should it be marked critical or not ?
            lset enhkeyusage 0 [concat [lindex $enhkeyusage 0] $extra_enhkeyusage]
        } else {
            # TBD - should it be marked critical or not ?
            set enhkeyusage [list $extra_enhkeyusage 1]
        }
    }
    if {[info exists enhkeyusage]} {
        lappend exts [_make_enhkeyusage_ext {*}$enhkeyusage]
    }

    if {[info exists altnames]} {
        lappend exts [_make_altnames_ext {*}$altnames]
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

proc twapi::_parse_store_open_opts {optvals} {
    array set opts [parseargs optvals  {
        {commitenable.bool    0 0x00010000}
        {readonly.bool        0 0x00008000}
        {existing.bool        0 0x00004000}
        {create.bool          0 0x00002000}
        {includearchived.bool 0 0x00000200}
        {maxpermissions.bool  0 0x00001000}
        {deferclose.bool      0 0x00000004}
        {backupprivilege.bool 0 0x00000800}
    } -maxleftover 0 -nulldefault]

    set flags 0
    foreach {opt val} [array get opts] {
        incr flags $val
    }
    return $flags
}

# Helper to return as der/pem based on encoding option
proc twapi::_as_pem_or_der {bin tag encoding} {
    if {$encoding eq "pem"} {
        # 1 -> CRYPT_STRING_BASE64
        # 0x80000000 -> LF-only, not CRLF
        return "-----BEGIN $tag-----\n[CryptBinaryToString $bin 0x80000001]-----END $tag-----\n"
    } else {
        return $bin
    }
}

# Helper for converting input parameters if they are in PEM format
# pem_or_der is the data
# enc specifies the type of pem_or_der. If empty, we guess.
# pemtype should generally be
# 0 -> CRYPT_STRING_BASE64HEADER for certificates
# 1 -> CRYPT_STRING_BASE64 (no header)
# 3 -> CRYPT_STRING_BASE64REQUESTHEADER
# 6 -> CRYPT_STRING_BASE64_ANY (actually same as 0 or 1)
proc twapi::_pem_decode {pem_or_der enc {pemtype 6}} {
    if {$enc eq "der"} {
        return $pem_or_der
    }
    if {$enc eq "pem" || 
        [regexp -nocase {^\s*-----\s*BEGIN\s+} $pem_or_der]} {
        return [CryptStringToBinary $pem_or_der $pemtype]
    }
    return $pem_or_der
}

proc twapi::_is_pem {pem_or_der} {
    return [regexp -nocase {^\s*-----\s*BEGIN\s+} $pem_or_der]
}

# Utility proc to generate certs in a memory store - 
# one self signed which is used to sign a client and a server cert
proc twapi::make_test_certs {{hstore {}} args} {
    crypt_test_container_cleanup

    parseargs args {
        {csp.arg {Microsoft Strong Cryptographic Provider}}
        {csptype.arg prov_rsa_full}
        unique
        {duration.int 5}
    } -maxleftover 0 -setvars

    set enddate [clock format [clock seconds] -format "%Y %N %e"]
    lset enddate 0 [expr {[lindex $enddate 0]+$duration}]
    # Ensure valid date e.g. Feb 29 non-leap year
    set enddate [clock format [clock scan $enddate -format "%Y %N %e"] -format "%Y %N %e"]

    if {$unique} {
        set uuid [twapi::new_uuid]
    } else {
        set uuid ""
    }

    # Create the self signed CA cert
    set container twapitestca$uuid
    set crypt [twapi::crypt_acquire $container -csp $csp -csptype $csptype -create 1]
    twapi::crypt_key_free [twapi::crypt_generate_key $crypt signature -exportable 1]
    set ca_altnames [list [list [list email ${container}@twapitest.com] [list dns ${container}.twapitest.com] [list url http://${container}.twapitest.com] [list directory [cert_name_to_blob "CN=${container}altname"]] [list ip [binary format c4 {127 0 0 2}]]]]
    set cert [twapi::cert_create_self_signed_from_crypt_context "CN=$container, C=IN, O=Tcl, OU=twapi" $crypt -purpose {ca} -altnames $ca_altnames -end $enddate]
    if {[llength $hstore] == 0} {
        set hstore [twapi::cert_temporary_store]
    }
    set ca_certificate [twapi::cert_store_add_certificate $hstore $cert]
    twapi::cert_release $cert
    twapi::cert_set_key_prov $ca_certificate $container signature -csp $csp -csptype $csptype
    crypt_free $crypt

    # Create the client and server certs
    foreach cert_type {intermediate server client altserver full min} {
        set container twapitest${cert_type}$uuid
        set subject $container
        set crypt [twapi::crypt_acquire $container -csp $csp -csptype $csptype -create 1]
        twapi::crypt_key_free [twapi::crypt_generate_key $crypt keyexchange -exportable 1]
        switch $cert_type {
            intermediate {
                set req [cert_request_create "CN=$container, C=IN, O=Tcl, OU=twapi" $crypt keyexchange -purpose ca]
                set signing_cert $ca_certificate
            }
            altserver {
                # No COMMON name. Used for testing use of DNS altname
                set altnames [list [list [list dns ${cert_type}.twapitest.com] [list dns ${cert_type}2.twapitest.com]]]
                set req [cert_request_create "C=IN, O=Tcl, OU=twapi, OU=$container" $crypt keyexchange -purpose $cert_type -altnames $altnames]
                set signing_cert $ca_certificate
            }
            client -
            server {
                set req [cert_request_create "CN=$container, C=IN, O=Tcl, OU=twapi" $crypt keyexchange -purpose $cert_type]
                set signing_cert $intermediate_certificate
            }
            full {
                set altnames [list [list [list email ${container}@twapitest.com] [list dns ${cert_type}.twapitest.com] [list url http://${container}.twapitest.com] [list directory [cert_name_to_blob "CN=${container}altname"]] [list ip [binary format c4 {127 0 0 1}]]]]
                set req [cert_request_create \
                             "CN=$container, C=IN, O=Tcl, OU=twapi" \
                             $crypt keyexchange \
                             -keyusage [list {crl_sign data_encipherment digital_signature key_agreement key_cert_sign key_encipherment non_repudiation} 1]\
                             -enhkeyusage [list {client_auth code_signing email_protection ipsec_end_system  ipsec_tunnel ipsec_user server_auth timestamp_signing ocsp_signing} 1] \
                             -altnames $altnames]
                set signing_cert $ca_certificate
            }
            min {
                set req [cert_request_create "CN=$container" $crypt keyexchange]
                set signing_cert $ca_certificate
            }
        }
        crypt_free $crypt
        set parsed_req [cert_request_parse $req]
        set subject [dict get $parsed_req subject]
        set pubkey [dict get $parsed_req pubkey]
        set opts {}
        foreach optname {-basicconstraints -keyusage -enhkeyusage -altnames} {
            if {[dict exists $parsed_req extensions $optname]} {
                lappend opts $optname [dict get $parsed_req extensions $optname]
            }
        }
        set encoded_cert [cert_create $subject $pubkey $signing_cert {*}$opts -end $enddate]
        set certificate [twapi::cert_store_add_encoded_certificate $hstore $encoded_cert]
        twapi::cert_set_key_prov $certificate $container keyexchange -csp $csp -csptype $csptype
        if {$cert_type eq "intermediate"} {
            set intermediate_certificate $certificate
        } else {
            cert_release $certificate
        }
    }

    cert_release $ca_certificate
    cert_release $intermediate_certificate
    return $hstore
}

proc twapi::dump_test_certs {hstore dir {pfxfile twapitest.pfx}} {
    set fd [open [file join $dir $pfxfile] wb]
    puts -nonewline $fd [cert_store_export_pfx $hstore "" -exportprivatekeys 1]
    close $fd
    cert_store_iterate $hstore c {
        set fd [open [file join $dir [cert_subject_name $c -name simpledisplay].cer] wb]
        puts -nonewline $fd [cert_export $c]
        close $fd
    }
}

proc twapi::crypt_test_containers {} {
    set crypt [crypt_acquire "" -verifycontext 1]
    twapi::trap {
        set names {}
        foreach name [crypt_key_container_names $crypt] {
            if {[string match -nocase twapitest* $name]} {
                lappend names $name
            }
        }
    } finally {
        crypt_free $crypt
    }
    return $names
}

proc twapi::crypt_test_container_cleanup {} {
    foreach c [crypt_test_containers] {
        crypt_key_container_delete $c
    }
}


# If we are not being sourced from a executable resource, need to
# source the remaining support files. In the former case, they are
# automatically combined into one so the sourcing is not needed.
if {![info exists twapi::twapi_crypto_rc_sourced]} {
    source [file join [file dirname [info script]] sspi.tcl]
    source [file join [file dirname [info script]] tls.tcl]
}

