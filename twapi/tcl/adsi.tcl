#
# Copyright (c) 2010, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# ADSI routines

# TBD - document
proc twapi::adsi_translate_name {name to {from 0}} {
    if {! [string is integer -strict $to]} {
        switch -exact -- $to {
            fqdn          { set to 1 }
            samcompatible { set to 2 }
            display       { set to 3 }
            uniqueid      { set to 6 }
            canonical     { set to 7 }
            userprincipal { set to 8 }
            canonicalex   { set to 9 }
            serviceprincipal {set to 10 }
            dnsdomain     { set to 12 }
            unknown -
            default {
                error "Invalid target format specifier '$to'"
            }
        }
    }

    if {! [string is integer -strict $from]} {
        switch -exact -- $from {
            unknown       { set from 0 }
            fqdn          { set from 1 }
            samcompatible { set from 2 }
            display       { set from 3 }
            uniqueid      { set from 6 }
            canonical     { set from 7 }
            userprincipal { set from 8 }
            canonicalex   { set from 9 }
            serviceprincipal {set from 10 }
            dnsdomain     { set from 12 }
            default {
                error "Invalid source format specifier '$from'"
            }
        }
    }
        
    return [TranslateName $name $from $to]
}