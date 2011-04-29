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

    lassign [resource_stringid_to_stringblockid $resname]  block_id index_within_block

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

proc twapi::extract_resources {hmod {withdata 0}} {
    set result [dict create]
    foreach type [enumerate_resource_types $hmod] {
        set typedict [dict create]
        foreach name [enumerate_resource_names $hmod $type] {
            set namedict [dict create]
            foreach lang [enumerate_resource_languages $hmod $type $name] {
                if {$withdata} {
                    dict set namedict $lang [read_resource $hmod $type $name $lang]
                } else {
                    dict set namedict $lang {}
                }
            }
            dict set typedict $name $namedict
        }
        dict set result $type $typedict
    }
    return $result
}

proc twapi::_load_image {flags type hmod path args} {
    # The flags arg is generally 0x10 (load from file), or 0 (module)
    # or'ed with 0x8000 (shared). The latter can be overridden by
    # the -shared option but should not be except when loading from module.
    array set opts [parseargs args {
        {createdibsection.bool 0 0x2000}
        {defaultsize.bool  0  0x40}
        height.int
        {loadtransparent.bool 0 0x20}
        {monochrome.bool  0  0x1}
        {shared.bool  0  0x8000}
        {vgacolor.bool  0  0x80}
        width.int
    } -maxleftover 0 -nulldefault]

    set flags [expr {$flags | $opts(defaultsize) | $opts(loadtransparent) | $opts(monochrome) | $opts(shared) | $opts(vgacolor)}]

    set h [LoadImage $hmod $path $type $opts(width) $opts(height) $flags]
    # Cast as _SHARED if required to offer some protection against
    # being freed using DestroyIcon etc.
    set type [lindex {HGDIOBJ HICON HCURSOR} $type]
    if {$flags & 0x8000} {
        append type _SHARED
    }
    return [cast_handle $h $type]
}


proc twapi::_load_image_from_system {type id args} {
    variable _oem_image_syms

    if {![string is integer -strict $id]} {
        if {![info exists _oem_image_syms]} {
            # Bitmap symbols (type 0)
            dict set _oem_image_syms 0 {
                CLOSE           32754            UPARROW         32753
                DNARROW         32752            RGARROW         32751
                LFARROW         32750            REDUCE          32749
                ZOOM            32748            RESTORE         32747
                REDUCED         32746            ZOOMD           32745
                RESTORED        32744            UPARROWD        32743
                DNARROWD        32742            RGARROWD        32741
                LFARROWD        32740            MNARROW         32739
                COMBO           32738            UPARROWI        32737
                DNARROWI        32736            RGARROWI        32735
                LFARROWI        32734            OLD_CLOSE       32767
                SIZE            32766            OLD_UPARROW     32765
                OLD_DNARROW     32764            OLD_RGARROW     32763
                OLD_LFARROW     32762            BTSIZE          32761
                CHECK           32760            CHECKBOXES      32759
                BTNCORNERS      32758            OLD_REDUCE      32757
                OLD_ZOOM        32756            OLD_RESTORE     32755
            }            
            # Icon symbols (type 1)
            dict set _oem_image_syms 1 {
                SAMPLE          32512            HAND            32513
                QUES            32514            BANG            32515
                NOTE            32516            WINLOGO         32517
                WARNING         32515            ERROR           32513
                INFORMATION     32516            SHIELD          32518
            }
            # Cursor symbols (type 2)
            dict set _oem_image_syms 2 {
                NORMAL          32512            IBEAM           32513
                WAIT            32514            CROSS           32515
                UP              32516            SIZENWSE        32642
                SIZENESW        32643            SIZEWE          32644
                SIZENS          32645            SIZEALL         32646
                NO              32648            HAND            32649
                APPSTARTING     32650
            }

        }
    }
        
    set id [dict get $_oem_image_syms $type [string toupper $id]]
    # Built-in system images must always be loaded shared (0x8000)
    return [_load_image 0x8000 $type NULL $id {*}$args]
}


# 0x10 -> LR_LOADFROMFILE. Also 0x8000 not set (meaning unshared)
interp alias {} twapi::load_bitmap_from_file {} twapi::_load_image 0x10 0 NULL
interp alias {} twapi::load_icon_from_file {} twapi::_load_image 0x10 1 NULL
interp alias {} twapi::load_cursor_from_file {} twapi::_load_image 0x10 2 NULL

interp alias {} twapi::load_bitmap_from_module {} twapi::_load_image 0 0
interp alias {} twapi::load_icon_from_module {} twapi::_load_image   0 1
interp alias {} twapi::load_cursor_from_module {} twapi::_load_image 0 2

interp alias {} twapi::load_bitmap_from_system {} twapi::_load_image_from_system 0
interp alias {} twapi::load_icon_from_system {} twapi::_load_image_from_system   1
interp alias {} twapi::load_cursor_from_system {} twapi::_load_image_from_system 2

interp alias {} twapi::free_icon {} twapi::DestroyIcon
interp alias {} twapi::free_bitmap {} twapi::DeleteObject
interp alias {} twapi::free_cursor {} twapi::DestroyCursor
