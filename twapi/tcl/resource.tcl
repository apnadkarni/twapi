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

# Note this is not an alias because we want to control arguments
# to UpdateResource (which can take more args that specified here)
proc twapi::delete_resource {hmod restype resname langid} {
    UpdateResource $hmod $restype $resname $langid
}


# Note this is not an alias because we want to make sure $bindata is specified
# as an argument else it will have the effect of deleting a resource
proc twapi::update_resource {hmod restype resname langid bindata} {
    UpdateResource $hmod $restype $resname $langid $bindata
}

proc twapi::end_resource_update {hmod args} {
    array set opts [parseargs args {
        discard
    } -maxleftover 0]

    return [EndUpdateResource $hmod $opts(discard)]
}

proc twapi::read_resource {hmod restype resname langid} {
    return [Twapi_LoadResource $hmod [FindResourceEx $hmod $restype $resname $langid]]
}

proc twapi::read_resource_string {hmod resname langid} {
    # As an aside, note that we do not use a LoadString call
    # because it does not allow for specification of a langid
    
    # For a reference to how strings are stored, see
    # http://blogs.msdn.com/b/oldnewthing/archive/2004/01/30/65013.aspx
    # or http://support.microsoft.com/kb/196774

    if {![string is integer -strict $resname]} {
        error "String resources must have an integer id"
    }

    foreach {block_id index_within_block} [resource_stringid_to_stringblockid $resname] break

    return [lindex \
                [resource_stringblock_to_strings \
                     [read_resource $hmod 6 $block_id $langid] ] \
                $index_within_block]
}

interp alias {} twapi::resource_stringblock_to_strings {} twapi::Twapi_SplitStringResource

# Give a list of strings, formats it as a string block. Number of strings
# must not be greater than 16. If less than 16 strings, remaining are
# treated as empty.
proc twapi::strings_to_resource_stringblock {strings} {
    if {[llength $strings] > 16} {
        error "Cannot have more than 16 strings in a resource string block."
    }

    for {set i 0} {$i < 16} {incr i} {
        set s [lindex $strings $i]
        set n [string length $s]
        append bin [binary format sa* $n [encoding convertto unicode $s]]
    }

    return $bin
}

proc twapi::resource_stringid_to_stringblockid {id} {
    # Strings are stored in blocks of 16, with block id's beginning at 1, not 0
    return [list [expr {($id / 16) + 1}] [expr {$id & 15}]]
}

interp alias {} twapi::enumerate_resource_languages {} twapi::Twapi_EnumResourceLanguages
interp alias {} twapi::enumerate_resource_names {} twapi::Twapi_EnumResourceNames
interp alias {} twapi::enumerate_resource_types {} twapi::Twapi_EnumResourceTypes
