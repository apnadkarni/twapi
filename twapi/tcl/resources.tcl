# Commands related to resource manipulation

#
# Copyright (c) 2003, 2004, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

proc twapi::begin_resource_update {path args} {
    array set opts [parseargs args {
        deleteall
    } -maxleftover 0]

    return [BeginUpdateResource $path $opts(deleteall)]
}

proc twapi::delete_resource {hmod restype resname langid} {
    # String resources have to be handled differently
    if {$restype != 6} {
        UpdateResource $hmod $restype $resname $langid
    } else {
        TBD - locate string resource and delete it
    }
}

proc twapi::update_resource {hmod restype resname langid bindata} {
    # String resources have to be handled differently
    if {$restype != 6} {
        UpdateResource $hmod $restype $resname $langid $bindata
    } else {
        TBD - locate string resource and update it
    }
}

proc twapi::end_resource_update {hmod args} {
    array set opts [parseargs args {
        discard
    } -maxleftover 0]

    return [EndUpdateResource $hmod $opts(discard)]
}

proc twapi::read_resource {hmod restype resname langid} {
    # Strings (type 6) have to be handled differently as they are stored blocks
    if {$restype != 6} {
        return [Twapi_LoadResource $hmod [FindResourceEx $hmod $restype $resname $langid]]
    }

    # As an aside, note that we do not use a LoadString call
    # because it does not allow for specification of a langid
    
    # For a reference to how strings are stored, see
    # http://blogs.msdn.com/b/oldnewthing/archive/2004/01/30/65013.aspx
    # or http://support.microsoft.com/kb/196774

    if {![string is integer -strict $resname]} {
        error "String resources must have an integer id"
    }

    # Strings are stored in blocks of 16, so figure out the block id
    set block_id [expr {($resname / 16) + 1}]
    set block [Twapi_LoadResource $hmod [FindResourceEx $hmod 6 $block_id $langid]]
    return [lindex [Twapi_SplitStringResource $block] [expr {$resname & 15}]]
}
