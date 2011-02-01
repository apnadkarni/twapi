#
# Copyright (c) 2004-2006 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {}


# Get the specified shell folder
proc twapi::get_shell_folder {csidl args} {
    variable csidl_lookup

    array set opts [parseargs args {create} -maxleftover 0]

    # Following are left out because they refer to virtual folders
    # and will return error if used here
    #    CSIDL_BITBUCKET - 0xa
    if {![info exists csidl_lookup]} {
        array set csidl_lookup {
            CSIDL_ADMINTOOLS 0x30
            CSIDL_COMMON_ADMINTOOLS 0x2f
            CSIDL_APPDATA 0x1a
            CSIDL_COMMON_APPDATA 0x23
            CSIDL_COMMON_DESKTOPDIRECTORY 0x19
            CSIDL_COMMON_DOCUMENTS 0x2e
            CSIDL_COMMON_FAVORITES 0x1f
            CSIDL_COMMON_MUSIC 0x35
            CSIDL_COMMON_PICTURES 0x36
            CSIDL_COMMON_PROGRAMS 0x17
            CSIDL_COMMON_STARTMENU 0x16
            CSIDL_COMMON_STARTUP 0x18
            CSIDL_COMMON_TEMPLATES 0x2d
            CSIDL_COMMON_VIDEO 0x37
            CSIDL_COOKIES 0x21
            CSIDL_DESKTOPDIRECTORY 0x10
            CSIDL_FAVORITES 0x6
            CSIDL_HISTORY 0x22
            CSIDL_INTERNET_CACHE 0x20
            CSIDL_LOCAL_APPDATA 0x1c
            CSIDL_MYMUSIC 0xd
            CSIDL_MYPICTURES 0x27
            CSIDL_MYVIDEO 0xe
            CSIDL_NETHOOD 0x13
            CSIDL_PERSONAL 0x5
            CSIDL_PRINTHOOD 0x1b
            CSIDL_PROFILE 0x28
            CSIDL_PROFILES 0x3e
            CSIDL_PROGRAMS 0x2
            CSIDL_PROGRAM_FILES 0x26
            CSIDL_PROGRAM_FILES_COMMON 0x2b
            CSIDL_RECENT 0x8
            CSIDL_SENDTO 0x9
            CSIDL_STARTMENU 0xb
            CSIDL_STARTUP 0x7
            CSIDL_SYSTEM 0x25
            CSIDL_TEMPLATES 0x15
            CSIDL_WINDOWS 0x24
        }
    }

    if {![string is integer $csidl]} {
        set csidl_key [string toupper $csidl]
        if {![info exists csidl_lookup($csidl_key)]} {
            # Try by adding a CSIDL prefix
            set csidl_key "CSIDL_$csidl_key"
            if {![info exists csidl_lookup($csidl_key)]} {
                error "Invalid CSIDL value '$csidl'"
            }
        }
        set csidl $csidl_lookup($csidl_key)
    }

    trap {
        set path [SHGetSpecialFolderPath 0 $csidl $opts(create)]
    } onerror {} {
        # Try some other way to get the information
        set code $errorCode
        set msg $errorResult
        set info $errorInfo
        switch -exact -- [format %x $csidl] {
            1a { catch {set path $::env(APPDATA)} }
            2b { catch {set path $::env(CommonProgramFiles)} }
            26 { catch {set path $::env(ProgramFiles)} }
            24 { catch {set path $::env(windir)} }
            25 { catch {set path [file join $::env(systemroot) system32]} }
        }
        if {![info exists path]} {
            #error $msg $info $code
            return ""
        }
    }

    return $path
}

# Displays a shell property dialog for the given object
proc twapi::shell_object_properties_dialog {path args} {
    array set opts [parseargs args {
        {type.arg "" {"" file printer volume}}
        {hwin.int 0}
        {page.arg ""}
    } -maxleftover 0]

    if {$opts(type) eq ""} {
        # Try figure out object type
        if {[file exists $path]} {
            set opts(type) file
        } elseif {[lsearch -exact [string tolower [find_volumes]] [string tolower $path]] >= 0} {
            set opts(type) volume
        } else {
            # Check if printer
            foreach printer [enumerate_printers] {
                if {[string equal -nocase [kl_get $printer name] $path]} {
                    set opts(type) printer
                    break
                }
            }
            if {$opts(type) eq ""} {
                error "Could not figure out type of object '$path'"
            }
        }
    }


    if {$opts(type) eq "file"} {
        set path [file nativename [file normalize $path]]
    }

    SHObjectProperties $opts(hwin) \
        [string map {printer 1 file 2 volume 4} $opts(type)] \
        $path \
        $opts(page)
}

# Show property dialog for a file
proc twapi::file_properties_dialog {name args} {
    array set opts [parseargs args {
        {hwin.int 0}
        {page.arg ""}
    } -maxleftover 0]

    shell_object_properties_dialog $name -type file -hwin $opts(hwin) -page $opts(page)
}


# Writes a shell shortcut
proc twapi::write_shortcut {link args} {
    
    array set opts [parseargs args {
        path.arg
        idl.arg
        args.arg
        desc.arg
        hotkey.arg
        iconpath.arg
        iconindex.int
        {showcmd.arg normal}
        workdir.arg
        relativepath.arg
    } -nulldefault -maxleftover 0]

    # Map hot key to integer if needed
    if {![string is integer -strict $opts(hotkey)]} {
        if {$opts(hotkey) eq ""} {
            set opts(hotkey) 0
        } else {
            # Try treating it as symbolic
            lassign [_hotkeysyms_to_vk $opts(hotkey)]  modifiers vk
            set opts(hotkey) $vk
            if {$modifiers & 1} {
                set opts(hotkey) [expr {$opts(hotkey) | (4<<8)}]
            }
            if {$modifiers & 2} {
                set opts(hotkey) [expr {$opts(hotkey) | (2<<8)}]
            }
            if {$modifiers & 4} {
                set opts(hotkey) [expr {$opts(hotkey) | (1<<8)}]
            }
            if {$modifiers & 8} {
                set opts(hotkey) [expr {$opts(hotkey) | (8<<8)}]
            }
        }
    }

    # IF a known symbol translate it. Note caller can pass integer
    # values as well which will be kept as they are. Bogus valuse and
    # symbols will generate an error on the actual call so we don't
    # check here.
    switch -exact -- $opts(showcmd) {
        minimized { set opts(showcmd) 7 }
        maximized { set opts(showcmd) 3 }
        normal    { set opts(showcmd) 1 }
    }

    Twapi_WriteShortcut $link $opts(path) $opts(idl) $opts(args) \
        $opts(desc) $opts(hotkey) $opts(iconpath) $opts(iconindex) \
        $opts(relativepath) $opts(showcmd) $opts(workdir)
}


# Read a shortcut
proc twapi::read_shortcut {link args} {
    array set opts [parseargs args {
        shortnames
        uncpath
        rawpath
        timeout.int
        {hwin.int 0}
        install
        nosearch
        notrack
        noui
        nolinkinfo
        anymatch
    } -maxleftover 0]

    set pathfmt 0
    foreach {opt val} {shortnames 1 uncpath 2 rawpath 4} {
        if {$opts($opt)} {
            setbits pathfmt $val
        }
    }

    set resolve_flags 4;                # SLR_UPDATE
    foreach {opt val} {
        install      128
        nolinkinfo    64
        notrack       32
        nosearch      16
        anymatch       2
        noui           1
    } {
        if {$opts($opt)} {
            setbits resolve_flags $val
        }
    }

    array set shortcut [twapi::Twapi_ReadShortcut $link $pathfmt $opts(hwin) $resolve_flags]

    switch -exact -- $shortcut(-showcmd) {
        1 { set shortcut(-showcmd) normal }
        3 { set shortcut(-showcmd) maximized }
        7 { set shortcut(-showcmd) minimized }
    }

    return [array get shortcut]
}



# Writes a url shortcut
proc twapi::write_url_shortcut {link url args} {
    
    array set opts [parseargs args {
        {missingprotocol.arg 0}
    } -nulldefault -maxleftover 0]

    switch -exact -- $opts(missingprotocol) {
        guess {
            set opts(missingprotocol) 1; # IURL_SETURL_FL_GUESS_PROTOCOL
        }
        usedefault {
            # 3 -> IURL_SETURL_FL_GUESS_PROTOCOL | IURL_SETURL_FL_USE_DEFAULT_PROTOCOL
            # The former must also be specified (based on experimentation)
            set opts(missingprotocol) 3
        }
        default {
            if {![string is integer -strict $opts(missingprotocol)]} {
                error "Invalid value '$opts(missingprotocol)' for -missingprotocol option."
            }
        }
    }

    Twapi_WriteUrlShortcut $link $url $opts(missingprotocol)
}

# Read a url shortcut
proc twapi::read_url_shortcut {link} {
    return [Twapi_ReadUrlShortcut $link]
}

# Invoke a url shortcut
proc twapi::invoke_url_shortcut {link args} {
    
    array set opts [parseargs args {
        verb.arg
        {hwin.int 0}
        allowui
    } -maxleftover 0]

    set flags 0
    if {$opts(allowui)} {setbits flags 1}
    if {! [info exists opts(verb)]} {
        setbits flags 2
        set opts(verb) ""
    }
    

    Twapi_InvokeUrlShortcut $link $opts(verb) $flags $opts(hwin)
}

# Send a file to the recycle bin
proc twapi::recycle_file {fn args} {
    array set opts [parseargs args {
        confirm.bool
        showerror.bool
    } -maxleftover 0 -nulldefault]

    set fn [file nativename [file normalize $fn]]

    if {$opts(confirm)} {
        set flags 0x40;         # FOF_ALLOWUNDO
    } else {
        set flags 0x50;         # FOF_ALLOWUNDO | FOF_NOCONFIRMATION
    }

    if {! $opts(showerror)} {
        set flags [expr {$flags | 0x0400}]; # FOF_NOERRORUI
    }

    return [expr {[lindex [Twapi_SHFileOperation 0 3 [list $fn] __null__ $flags ""] 0] ? false : true}]
}

proc twapi::shell_execute args {
    array set opts [parseargs args {
        asyncok.bool
        class.arg
        connect.bool
        dir.arg
        getprocesshandle.bool
        {hicon.arg NULL}
        {hkeyclass.arg NULL}
        {hmonitor.arg NULL}
        hotkey.int
        hwin.int
        idl.arg
        invokeidlist.bool
        logusage.bool
        noconsole.bool
        noui.bool
        nozonechecks.bool
        params.arg
        path.arg
        {show.arg 1}
        substenv.bool
        unicode.bool
        verb.arg
        {wait.bool true}
        waitforinputidle.bool
    } -maxleftover 0 -nulldefault]

    set fmask 0

    foreach {opt mask} {
        class     1
        idl       4
    } {
        if {$opts($opt) ne ""} {
            setbits fmask $mask
        }
    }

    if {$opts(hkeyclass) ne "NULL"} {
        setbits fmask 3
    }

    foreach {opt mask} {
        getprocesshandle 0x00000040
        connect          0x00000080
        wait             0x00000100
        substenv         0x00000200
        noui             0x00000400
        unicode          0x00004000
        noconsole        0x00008000
        asyncok          0x00100000
        nozonechecks     0x00800000
        waitforinputidle 0x02000000
        logusage         0x04000000
        invokeidlist     0x0000000C
    } {
        if {$opts($opt)} {
            setbits fmask $mask
        }
    }


    if {$opts(hicon) ne "NULL" && $opts(hmonitor) ne "NULL"} {
        error "Cannot specify -hicon and -hmonitor options together."
    }

    set hiconormonitor NULL
    if {$opts(hicon) ne "NULL"} {
        set hiconormonitor $opts(hicon)
        setbits flags 0x00000010
    } elseif {$opts(hmonitor) ne "NULL"} {
        set hiconormonitor $opts(hmonitor)
        setbits flags 0x00200000
    }

    if {![string is integer -strict $opts(show)]} {
        variable windefs
        set def SW_[string toupper $opts(show)]
        if {[info exists windefs($def)]} {
            set opts(show) $windefs($def)
        } else {
            error "Invalid value $opts(show) specified for option -show"
        }
    }

    return [Twapi_ShellExecuteEx \
                $fmask \
                $opts(hwin) \
                $opts(verb) \
                $opts(path) \
                $opts(params) \
                $opts(dir) \
                $opts(show) \
                $opts(idl) \
                $opts(class) \
                $opts(hkeyclass) \
                $opts(hotkey) \
                $hiconormonitor]
}
