#
# Copyright (c) 2020 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {}


proc twapi::reg_key_open {hkey subkey args} {
    parseargs args {
        {link.bool 0}
        {access.arg generic_read}
        32bit
        64bit
    } -maxleftover 0 -setvars

    set access [_access_rights_to_mask $access]
    # Note: Following might be set via -access as well. The -32bit and -64bit
    # options just make it a little more convenient for caller
    if {$32bit} {
        set access [expr {$access | 0x200}]
    }
    if {$64bit} {
        set access [expr {$access | 0x100}]
    }
    return [RegOpenKeyEx $hkey $subkey $link $access]
}

proc twapi::reg_key_create {hkey subkey args} {
    parseargs args {
        {access.arg generic_read}
        {inherit.bool 0}
        {secd.arg ""}
        {volatile.bool 0 0x1}
        {backup.bool 0 0x4}
        32bit
        64bit
        disposition.arg
    } -maxleftover 0 -setvars

    set access [_access_rights_to_mask $access]
    # Note: Following might be set via -access as well. The -32bit and -64bit
    # options just make it a little more convenient for caller
    if {$32bit} {
        set access [expr {$access | 0x200}]
    }
    if {$64bit} {
        set access [expr {$access | 0x100}]
    }
    lassign [RegCreateKeyEx \
                 $hkey \
                 $subkey \
                 0 \
                 "" \
                 [expr {$volatile | $backup}] \
                 $access \
                 [_make_secattr $opts(secd) $inherit] \
                ] hkey created
    if {[info exists disposition]} {
        upvar 1 $disposition created_flags
        set created_flag $created
    }
    return $hkey
}

proc twapi::reg_value_delete {hkey value_name {subkey {}}} {
    if {$subkey eq ""} {
        RegDeleteValue $hkey $value_name
    } else {
        RegDeleteKeyValue $hkey $subkey $value_name
    }
}

proc twapi::reg_set_sz {hkey value_name value} {
    RegSetValueEx $hkey $value_name sz $value
}

proc twapi::reg_set_expand_sz {hkey value_name value} {
    RegSetValueEx $hkey $value_name expand_sz $value
}
proc twapi::reg_set_binary {hkey value_name value} {
    RegSetValueEx $hkey $value_name binary $value
}
proc twapi::reg_set_dword {hkey value_name value} {
    RegSetValueEx $hkey $value_name dword $value
}
proc twapi::reg_set_dword_be {hkey value_name value} {
    RegSetValueEx $hkey $value_name dword_be $value
}
proc twapi::reg_set_link {hkey value_name value} {
    RegSetValueEx $hkey $value_name link $value
}
proc twapi::reg_set_multi_sz {hkey value_name value} {
    RegSetValueEx $hkey $value_name multi_sz $value
}
proc twapi::reg_set_resource_list {hkey value_name value} {
    RegSetValueEx $hkey $value_name resource_list $value
}
proc twapi::reg_set_full_resource_descriptor {hkey value_name value} {
    RegSetValueEx $hkey $value_name full_resource_descriptor $value
}
proc twapi::reg_set_resource_requirements_list {hkey value_name value} {
    RegSetValueEx $hkey $value_name resource_requirements_list $value
}
proc twapi::reg_set_qword {hkey value_name value} {
    RegSetValueEx $hkey $value_name qword $value
}

proc twapi::reg_key_delete {hkey subkey args} {
    parseargs args {
        32bit
        64bit
    } -maxleftover 0 -setvars
    set access 0
    if {$32bit} {
        set access [expr {$access | 0x200}]
    }
    if {$64bit} {
        set access [expr {$access | 0x100}]
    }

    RegDeleteKeyEx $hkey $subkey $access
}

proc twapi::reg_enum_value_names {hkey} {
    return [RegEnumValue $hkey 0]
}

proc twapi::reg_enum_values {hkey} {
    return [RegEnumValue $hkey 1]
}

proc twapi::reg_key_copy {from_hkey subkey to_hkey} {
    SHCopyKey $from_hkey $subkey $to_hkey
}

proc twapi::reg_open_user {args} {
    parseargs args {
        {access.arg generic_read}
        32bit
        64bit
    } -maxleftover 0 -setvars

    set access [_access_rights_to_mask $access]
    # Note: Following might be set via -access as well. The -32bit and -64bit
    # options just make it a little more convenient for caller
    if {$32bit} {
        set access [expr {$access | 0x200}]
    }
    if {$64bit} {
        set access [expr {$access | 0x100}]
    }
    return [RegOpenCurrentUser $access]
}

proc twapi::reg_key_user_classes_root {usertoken args} {
    parseargs args {
        {access.arg generic_read}
        32bit
        64bit
    } -maxleftover 0 -setvars

    set access [_access_rights_to_mask $access]
    # Note: Following might be set via -access as well. The -32bit and -64bit
    # options just make it a little more convenient for caller
    if {$32bit} {
        set access [expr {$access | 0x200}]
    }
    if {$64bit} {
        set access [expr {$access | 0x100}]
    }
    return [RegOpenUserClassesRoot 0 $access]
}

proc twapi::reg_save_to_file {hkey filepath args} {
    parseargs args {
        {secd.arg {}}
        {format.int 2 {1 2}}
        {compress.bool 1}
    } -setvars

    if {! $compress} {
        set format [expr {$format | 4}]
    }
    RegSaveKeyEx $hkey $filepath [_make_secattr $secd 0] $format
}

proc twapi::reg_restore_from_file {hkey filepath args} {
    parseargs args {
        {volatile.bool 0 0x1}
        {force.bool 0 0x8}
    } -setvars
    RegRestoreKey $hkey $filepath [expr {$force | $volatile}]
}

proc twapi::reg_key_monitor {hkey args} {
    parseargs arg {
        {keys.bool 0 0x1}
        {attr.bool 0 0x2}
        {values.bool 0 0x4}
        {secd.bool 0 0x8}
        {subtree.bool 0}
        hevent.arg
    } -setvars

    set filter [expr {$keys | $attr | $values | $secd}]
    if {$filter == 0} {
        set filter $0xf
    }

    if {[info exists hevent]} {
        set async 1
    } else {
        set async 0
        set hevent $::twapi::nullptr
    }
    RegNotifyChangeKeyValue $hkey $subtree $filter $hevent $async
}
