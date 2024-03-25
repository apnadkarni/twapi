#
# Copyright (c) 2020 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {}

#
# TBD -32bit and -64bit options are not documented
# pending test cases

proc twapi::reg_key_copy {hkey to_hkey args} {
    parseargs args {
        subkey.arg
        copysecd.bool
    } -setvars -maxleftover 0 -nulldefault

    if {$copysecd} {
        RegCopyTree $hkey $subkey $to_hkey
    } else {
        SHCopyKey $hkey $subkey $to_hkey
    }
}

proc twapi::reg_key_create {hkey subkey args} {
    # TBD - document -link
    # [opt_def [cmd -link] [arg BOOL]] If [const true], [arg SUBKEY] is stored as the
    # value of the [const SymbolicLinkValue] value under [arg HKEY]. Default is
    # [const false].
    parseargs args {
        {access.arg generic_read}
        {inherit.bool 0}
        {secd.arg ""}
        {volatile.bool 0 0x1}
        {link.bool 0 0x2}
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
                 [_make_secattr $secd $inherit] \
                ] hkey disposition_value
    if {[info exists disposition]} {
        upvar 1 $disposition created_or_existed
        if {$disposition_value == 1} {
            set created_or_existed created
        } else {
            # disposition_value == 2
            set created_or_existed existed
        }
    }
    return $hkey
}

proc twapi::reg_key_delete {hkey subkey args} {
    parseargs args {
        32bit
        64bit
    } -maxleftover 0 -setvars

    # TBD - document options after adding tests
    set access 0
    if {$32bit} {
        set access [expr {$access | 0x200}]
    }
    if {$64bit} {
        set access [expr {$access | 0x100}]
    }

    RegDeleteKeyEx $hkey $subkey $access
}

proc twapi::reg_keys {hkey {subkey {}}} {
    if {$subkey ne ""} {
        set hkey [reg_key_open $hkey $subkey]
    }
    try {
        return [RegEnumKeyEx $hkey 0]
    } finally {
        if {$subkey ne ""} {
            reg_key_close $hkey
        }
    }
}

proc twapi::reg_key_open {hkey subkey args} {
    # Not documented: -link, -32bit, -64bit
    # [opt_def [cmd -link] [arg BOOL]] If [const true], specifies the key is a
    # symbolic link. Defaults to [const false].
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

proc twapi::reg_value_delete {hkey args} {
    if {[llength $args] == 1} {
        RegDeleteValue $hkey [lindex $args 0]
    } elseif {[llength $args] == 2} {
        RegDeleteKeyValue $hkey {*}$args
    } else {
        error "Wrong # args: should be \"reg_value_delete ?SUBKEY? VALUENAME\""
    }
}

proc twapi::reg_key_current_user {args} {
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
    return [RegOpenUserClassesRoot $usertoken 0 $access]
}

proc twapi::reg_key_export {hkey filepath args} {
    parseargs args {
        {secd.arg {}}
        {format.arg xp {win2k xp}}
        {compress.bool 1}
    } -setvars

    set format [dict get {win2k 1 xp 2} $format]
    if {! $compress} {
        set format [expr {$format | 4}]
    }
    twapi::eval_with_privileges {
        RegSaveKeyEx $hkey $filepath [_make_secattr $secd 0] $format
    } SeBackupPrivilege
}

proc twapi::reg_key_import {hkey filepath args} {
    parseargs args {
        {volatile.bool 0 0x1}
        {force.bool 0 0x8}
    } -setvars
    twapi::eval_with_privileges {
        RegRestoreKey $hkey $filepath [expr {$force | $volatile}]
    } {SeBackupPrivilege SeRestorePrivilege}
}

proc twapi::reg_key_load {hkey hivename filepath} {
    twapi::eval_with_privileges {
        RegLoadKey $hkey $subkey $filepath
    } {SeBackupPrivilege SeRestorePrivilege}
}

proc twapi::reg_key_unload {hkey hivename} {
    twapi::eval_with_privileges {
        RegUnLoadKey $hkey $subkey
    } {SeBackupPrivilege SeRestorePrivilege}
}

proc twapi::reg_key_monitor {hkey hevent args} {
    parseargs args {
        {keys.bool 0 0x1}
        {attr.bool 0 0x2}
        {values.bool 0 0x4}
        {secd.bool 0 0x8}
        {subtree.bool 0}
    } -setvars

    set filter [expr {$keys | $attr | $values | $secd}]
    if {$filter == 0} {
        set filter 0xf
    }

    RegNotifyChangeKeyValue $hkey $subtree $filter $hevent 1
}

proc twapi::reg_value_names {hkey {subkey {}}} {
    if {$subkey eq ""} {
        # 0 - value names only
        return [RegEnumValue $hkey 0]
    }
    set hkey [reg_key_open $hkey $subkey]
    try {
        # 0 - value names only
        return [RegEnumValue $hkey 0]
    } finally {
        reg_key_close $hkey
    }
}

proc twapi::reg_values {hkey {subkey {}}} {
    if {$subkey eq ""} {
        #  3 -> 0x1 - return data values, 0x2 - cooked data
        return [RegEnumValue $hkey 3]
    }
    set hkey [reg_key_open $hkey $subkey]
    try {
        #  3 -> 0x1 - return data values, 0x2 - cooked data
        return [RegEnumValue $hkey 3]
    } finally {
        reg_key_close $hkey
    }
}

proc twapi::reg_values_raw {hkey {subkey {}}} {
    if {$subkey eq ""} {
        #  0x1 - return data values
        return [RegEnumValue $hkey 1]
    }
    set hkey [reg_key_open $hkey $subkey]
    try {
        return [RegEnumValue $hkey 1]
    } finally {
        reg_key_close $hkey
    }
}

proc twapi::reg_value_raw {hkey args} {
    if {[llength $args] == 1} {
        return [RegQueryValueEx $hkey [lindex $args 0] false]
    } elseif {[llength $args] == 2} {
        return [RegGetValue $hkey {*}$args 0x1000ffff false]
    } else {
        error "wrong # args: should be \"reg_value_get HKEY ?SUBKEY? VALUENAME\""
    }
}

proc twapi::reg_value {hkey args} {
    if {[llength $args] == 1} {
        return [RegQueryValueEx $hkey [lindex $args 0] true]
    } elseif {[llength $args] == 2} {
        return [RegGetValue $hkey {*}$args 0x1000ffff true]
    } else {
        error "wrong # args: should be \"reg_value_get HKEY ?SUBKEY? VALUENAME\""
    }
}

if {[twapi::min_os_version 6]} {
    proc twapi::reg_value_set {hkey args} {
        if {[llength $args] == 3} {
            return [RegSetValueEx $hkey {*}$args]
        } elseif {[llength $args] == 4} {
            return [RegSetKeyValue $hkey {*}$args]
        } else {
            error "wrong # args: should be \"reg_value_set HKEY ?SUBKEY? VALUENAME TYPE VALUE\""
        }
    }
} else {
    proc twapi::reg_value_set {hkey args} {
        if {[llength $args] == 3} {
            lassign $args value_name value_type value
        } elseif {[llength $args] == 4} {
            lassign $args subkey value_name value_type value
            set hkey [reg_key_open $hkey $subkey -access key_set_value]
        } else {
            error "wrong # args: should be \"reg_value_set HKEY ?SUBKEY? VALUENAME TYPE VALUE\""
        }
        try {
            RegSetValueEx $hkey $value_name $value_type $value
        } finally {
            if {[info exists subkey]} {
                # We opened hkey
                reg_close_key $hkey
            }
        }
    }
}

proc twapi::reg_key_override_undo {hkey} {
    RegOverridePredefKey $hkey 0
}

proc twapi::_reg_walker {hkey path callback cbdata} {
    # Callback for the key
    set code [catch {
        {*}$callback $cbdata $hkey $path
    } cbdata ropts]
    if {$code != 0} {
        if {$code == 4} {
            # Continue - skip children, continue with siblings
            return $cbdata
        } elseif {$code == 3} {
            # Skip siblings as well
            return -code break $cbdata
        } elseif {$code == 2} {
            # Stop complete iteration
            return -code return $cbdata
        } else {
            return -options $ropts $cbdata
        }
    }

    # Iterate over child keys
    foreach child_key [reg_keys $hkey] {
        set child_hkey [reg_key_open $hkey $child_key]
        try {
            # Recurse to call into children
            set code [catch {
                _reg_walker $child_hkey [linsert $path end $child_key] $callback $cbdata
            } cbdata ropts]
            if {$code != 0 && $code != 4} {
                if {$code == 3} {
                    # break - skip remaining child keys
                    return $cbdata
                } elseif {$code == 2} {
                    # return - stop all iteration all up the tree
                    return -code return $cbdata
                } else {
                    return -options $ropts $cbdata
                }
            }
        } finally {
            reg_key_close $child_hkey
        }
    }

    return $cbdata
}

proc twapi::reg_walk {hkey args} {
    parseargs args {
        {subkey.arg {}}
        callback.arg
        {cbdata.arg ""}
    } -maxleftover 0 -setvars


    if {$subkey ne ""} {
        set hkey [reg_key_open $hkey $subkey]
        set path [list $subkey]
    } else {
        set path [list ]
    }

    if {![info exists callback]} {
        set callback [lambda {cbdata hkey path} {puts [join $path \\]}]
    }
    try {
        set code [catch {_reg_walker $hkey $path $callback $cbdata } result ropts]
        # Codes 2 (return), 3 (break) and 4 (continue) are just early terminations
        if {$code == 1} {
            return -options $ropts $result
        }
    } finally {
        if {$subkey ne ""} {
            reg_key_close $hkey
        }
    }
    return $result
}

proc twapi::_reg_iterator_callback {cbdata hkey path args} {
    set cmd [yield [list $hkey $path {*}$args]]
    # Loop until valid argument
    while {1} {
        switch -exact -- $cmd {
            "" -
            next { return $cbdata }
            stop { return -code return $cbdata }
            parentsibling { return -code break $cbdata }
            sibling { return -code continue $cbdata }
            default {
                set ret [yieldto return -level 0 -code error "Invalid argument \"$cmd\"."]
            }
        }
    }
}

proc twapi::_reg_iterator_coro {hkey subkey} {
    set cmd [yield [info coroutine]]
    switch -exact -- $cmd {
        "" -
        next {
            # Drop into reg_walk
        }
        stop -
        parentsibling -
        sibling {
            return {}
        }
        default {
            error "Invalid argument \"$cmd\"."
        }
    }
    if {$subkey ne ""} {
        set hkey [reg_key_open $hkey $subkey]
    }
    try {
        reg_walk $hkey -callback [namespace current]::_reg_iterator_callback
    } finally {
        if {$subkey ne ""} {
            reg_key_close $hkey
        }
    }
    return
}

proc twapi::reg_iterator {hkey {subkey {}}} {
    variable reg_walk_counter

    return [coroutine "regwalk#[incr reg_walk_counter]" _reg_iterator_coro $hkey $subkey]
}

proc twapi::reg_tree {hkey {subkey {}}} {

    set iter [reg_iterator $hkey $subkey]

    set paths {}
    while {[llength [set item [$iter next]]]} {
        lappend paths [join [lindex $item 1] \\]
    }
    return $paths
}

proc twapi::reg_tree_values {hkey {subkey {}}} {

    set iter [reg_iterator $hkey $subkey]

    set tree {}
    # Note here we cannot ignore the first empty node corresponding
    # to the root because we have to return any values it contains.
    while {[llength [set item [$iter next]]]} {
        dict set tree [join [lindex $item 1] \\] [reg_values [lindex $item 0]]
    }
    return $tree
}

proc twapi::reg_tree_values_raw {hkey {subkey {}}} {
    set iter [reg_iterator $hkey $subkey]

    set tree {}
    while {[llength [set item [$iter next]]]} {
        dict set tree [join [lindex $item 1] \\] [reg_values_raw [lindex $item 0]]
    }
    return $tree
}

