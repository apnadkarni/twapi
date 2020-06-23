#
# Copyright (c) 2012 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

package require twapi_ui;       # SetCursorPos etc.

# Enable window input
proc twapi::enable_window_input {hwin} {
    return [expr {[EnableWindow $hwin 1] != 0}]
}

# Disable window input
proc twapi::disable_window_input {hwin} {
    return [expr {[EnableWindow $hwin 0] != 0}]
}

# CHeck if window input is enabled
proc twapi::window_input_enabled {hwin} {
    return [IsWindowEnabled $hwin]
}

# Simulate user input
proc twapi::send_input {inputlist} {
    array set input_defs {
        MOUSEEVENTF_MOVE        0x0001
        MOUSEEVENTF_LEFTDOWN    0x0002
        MOUSEEVENTF_LEFTUP      0x0004
        MOUSEEVENTF_RIGHTDOWN   0x0008
        MOUSEEVENTF_RIGHTUP     0x0010
        MOUSEEVENTF_MIDDLEDOWN  0x0020
        MOUSEEVENTF_MIDDLEUP    0x0040
        MOUSEEVENTF_XDOWN       0x0080
        MOUSEEVENTF_XUP         0x0100
        MOUSEEVENTF_WHEEL       0x0800
        MOUSEEVENTF_VIRTUALDESK 0x4000
        MOUSEEVENTF_ABSOLUTE    0x8000
        
        KEYEVENTF_EXTENDEDKEY 0x0001
        KEYEVENTF_KEYUP       0x0002
        KEYEVENTF_UNICODE     0x0004
        KEYEVENTF_SCANCODE    0x0008

        XBUTTON1      0x0001
        XBUTTON2      0x0002
    }

    set inputs [list ]
    foreach input $inputlist {
        if {[string equal [lindex $input 0] "mouse"]} {
            lassign $input mouse xpos ypos
            set mouseopts [lrange $input 3 end]
            array unset opts
            array set opts [parseargs mouseopts {
                relative moved
                ldown lup rdown rup mdown mup x1down x1up x2down x2up
                wheel.int
            }]
            set flags 0
            if {! $opts(relative)} {
                set flags $input_defs(MOUSEEVENTF_ABSOLUTE)
            }

            if {[info exists opts(wheel)]} {
                if {($opts(x1down) || $opts(x1up) || $opts(x2down) || $opts(x2up))} {
                    error "The -wheel input event attribute may not be specified with -x1up, -x1down, -x2up or -x2down events"
                }
                set mousedata $opts(wheel)
                set flags $input_defs(MOUSEEVENTF_WHEEL)
            } else {
                if {$opts(x1down) || $opts(x1up)} {
                    if {$opts(x2down) || $opts(x2up)} {
                        error "The -x1down, -x1up mouse input attributes are mutually exclusive with -x2down, -x2up attributes"
                    }
                    set mousedata $input_defs(XBUTTON1)
                } else {
                    if {$opts(x2down) || $opts(x2up)} {
                        set mousedata $input_defs(XBUTTON2)
                    } else {
                        set mousedata 0
                    }
                }
            }
            foreach {opt flag} {
                moved MOVE
                ldown LEFTDOWN
                lup   LEFTUP
                rdown RIGHTDOWN
                rup   RIGHTUP
                mdown MIDDLEDOWN
                mup   MIDDLEUP
                x1down XDOWN
                x1up   XUP
                x2down XDOWN
                x2up   XUP
            } {
                if {$opts($opt)} {
                    set flags [expr {$flags | $input_defs(MOUSEEVENTF_$flag)}]
                }
            }

            lappend inputs [list mouse $xpos $ypos $mousedata $flags]

        } else {
            lassign $input inputtype vk scan keyopts
            if {"-extended" ni $keyopts} {
                set extended 0
            } else {
                set extended $input_defs(KEYEVENTF_EXTENDEDKEY)
            }
            if {"-usescan" ni $keyopts} {
                set usescan 0
            } else {
                set usescan $input_defs(KEYEVENTF_SCANCODE)
            }
            switch -exact -- $inputtype {
                keydown {
                    lappend inputs [list key $vk $scan [expr {$extended|$usescan}]]
                }
                keyup {
                    lappend inputs [list key $vk $scan \
                                        [expr {$extended
                                               | $usescan
                                               | $input_defs(KEYEVENTF_KEYUP)
                                           }]]
                }
                key {
                    lappend inputs [list key $vk $scan [expr {$extended|$usescan}]]
                    lappend inputs [list key $vk $scan \
                                        [expr {$extended
                                               | $usescan
                                               | $input_defs(KEYEVENTF_KEYUP)
                                           }]]
                }
                unicode {
                    lappend inputs [list key 0 $scan $input_defs(KEYEVENTF_UNICODE)]
                    lappend inputs [list key 0 $scan \
                                        [expr {$input_defs(KEYEVENTF_UNICODE)
                                               | $input_defs(KEYEVENTF_KEYUP)
                                           }]]
                }
                default {
                    error "Unknown input type '$inputtype'"
                }
            }
        }
    }

    SendInput $inputs
}

# Block the input
proc twapi::block_input {} {
    return [BlockInput 1]
}

# Unblock the input
proc twapi::unblock_input {} {
    return [BlockInput 0]
}

# Send the given set of characters to the input queue
proc twapi::send_input_text {s} {
    return [Twapi_SendUnicode $s]
}

# send_keys - uses same syntax as VB SendKeys function
proc twapi::send_keys {keys} {
    set inputs [_parse_send_keys $keys]
    send_input $inputs
}


# Handles a hotkey notification
proc twapi::_hotkey_handler {msg atom key msgpos ticks} {
    variable _hotkeys

    # Note it is not an error if a hotkey does not exist since it could
    # have been deregistered in the time between hotkey input and receiving it.
    set code 0
    if {[info exists _hotkeys($atom)]} {
        foreach handler $_hotkeys($atom) {
            set code [catch {uplevel #0 $handler} msg]
            switch -exact -- $code {
                0 {
                    # Normal, keep going
                }
                1 {
                    # Error - put in background and abort
                    after 0 [list error $msg $::errorInfo $::errorCode]
                    break
                }
                3 {
                    break;      # Ignore remaining handlers
                }
                default {
                    # Keep going
                }
            }
        }
    }
    return -code $code ""
}

proc twapi::register_hotkey {hotkey script args} {
    variable _hotkeys

    # 0x312 -> WM_HOTKEY
    _register_script_wm_handler 0x312 [list [namespace current]::_hotkey_handler] 1

    array set opts [parseargs args {
        append
    } -maxleftover 0]

#    set script [lrange $script 0 end]; # Ensure a valid list

    lassign  [_hotkeysyms_to_vk $hotkey]  modifiers vk
    set hkid "twapi_hk_${vk}_$modifiers"
    set atom [GlobalAddAtom $hkid]
    if {[info exists _hotkeys($atom)]} {
        GlobalDeleteAtom $atom; # Undo above AddAtom since already there
        if {$opts(append)} {
            lappend _hotkeys($atom) $script
        } else {
            set _hotkeys($atom) [list $script]; # Replace previous script
        }
        return $atom
    }
    trap {
        RegisterHotKey $atom $modifiers $vk
    } onerror {} {
        GlobalDeleteAtom $atom; # Undo above AddAtom
        rethrow
    }
    set _hotkeys($atom) [list $script]; # Replace previous script
    return $atom
}

proc twapi::unregister_hotkey {atom} {
    variable _hotkeys
    if {[info exists _hotkeys($atom)]} {
        UnregisterHotKey $atom
        GlobalDeleteAtom $atom
        unset _hotkeys($atom)
    }
}


# Simulate clicking a mouse button
proc twapi::click_mouse_button {button} {
    switch -exact -- $button {
        1 -
        left { set down -ldown ; set up -lup}
        2 -
        right { set down -rdown ; set up -rup}
        3 -
        middle { set down -mdown ; set up -mup}
        x1     { set down -x1down ; set up -x1up}
        x2     { set down -x2down ; set up -x2up}
        default {error "Invalid mouse button '$button' specified"}
    }

    send_input [list \
                    [list mouse 0 0 $down] \
                    [list mouse 0 0 $up]]
    return
}

# Simulate mouse movement
proc twapi::move_mouse {xpos ypos {mode ""}} {
    # If mouse trails are enabled, it leaves traces when the mouse is
    # moved and does not clear them until mouse is moved again. So
    # we temporarily disable mouse trails if we can

    if {[llength [info commands ::twapi::get_system_parameters_info]] != 0} {
        set trail [get_system_parameters_info SPI_GETMOUSETRAILS]
        set_system_parameters_info SPI_SETMOUSETRAILS 0
    }
    switch -exact -- $mode {
        -relative {
            lappend cmd -relative
            lassign [GetCursorPos] curx cury
            incr xpos $curx
            incr ypos $cury
        }
        -absolute -
        ""        { }
        default   { error "Invalid mouse movement mode '$mode'" }
    }

    SetCursorPos $xpos $ypos

    # Restore trail setting if we had disabled it and it was originally enabled
    if {[info exists trail] && $trail} {
        set_system_parameters_info SPI_SETMOUSETRAILS $trail
    }
}

# Simulate turning of the mouse wheel
proc twapi::turn_mouse_wheel {wheelunits} {
    send_input [list [list mouse 0 0 -relative -wheel $wheelunits]]
    return
}

# Get the mouse/cursor position
proc twapi::get_mouse_location {} {
    return [GetCursorPos]
}

proc twapi::get_input_idle_time {} {
    # The formats are to convert wrapped 32bit signed to unsigned
    set last_event [format 0x%x [GetLastInputInfo]]
    set now [format 0x%x [GetTickCount]]

    # Deal with wrap around
    if {$now >= $last_event} {
        return [expr {$now - $last_event}]
    } else {
        return [expr {$now + (0xffffffff - $last_event) + 1}]
    }
}

# Initialize the virtual key table
proc twapi::_init_vk_map {} {
    variable vk_map

    if {![info exists vk_map]} {
        # Map tokens to VK_* key codes
        array set vk_map {
            BACK {0x08 0}
            BACKSPACE {0x08 0}   BS {0x08 0}   BKSP {0x08 0}   TAB {0x09 0}
            CLEAR {0x0C 0}   RETURN {0x0D 0}   ENTER {0x0D 0}   SHIFT {0x10 0}
            CONTROL {0x11 0}   MENU {0x12 0}   ALT {0x12 0}   PAUSE {0x13 0}
            BREAK {0x13 0}   CAPITAL {0x14 0}   CAPSLOCK {0x14 0}
            KANA {0x15 0}   HANGEUL {0x15 0}   HANGUL {0x15 0}   JUNJA {0x17 0}
            FINAL {0x18 0}   HANJA {0x19 0}   KANJI {0x19 0}   ESCAPE {0x1B 0}
            ESC {0x1B 0}   CONVERT {0x1C 0}   NONCONVERT {0x1D 0}
            ACCEPT {0x1E 0}   MODECHANGE {0x1F 0}   SPACE {0x20 0}
            PRIOR {0x21 0}   PGUP {0x21 0}   NEXT {0x22 0}   PGDN {0x22 0}
            END {0x23 0}   HOME {0x24 0}   LEFT {0x25 0}   UP {0x26 0}
            RIGHT {0x27 0}   DOWN {0x28 0}   SELECT {0x29 0}
            PRINT {0x2A 0}   PRTSC {0x2C 0}   EXECUTE {0x2B 0}
            SNAPSHOT {0x2C 0}   INSERT {0x2D 0}   INS {0x2D 0}
            DELETE {0x2E 0}   DEL {0x2E 0}   HELP {0x2F 0}   LWIN {0x5B 0}
            RWIN {0x5C 0}   APPS {0x5D 0}   SLEEP {0x5F 0}   NUMPAD0 {0x60 0}
            NUMPAD1 {0x61 0}   NUMPAD2 {0x62 0}   NUMPAD3 {0x63 0}
            NUMPAD4 {0x64 0}   NUMPAD5 {0x65 0}   NUMPAD6 {0x66 0}
            NUMPAD7 {0x67 0}   NUMPAD8 {0x68 0}   NUMPAD9 {0x69 0}
            MULTIPLY {0x6A 0}   ADD {0x6B 0}   SEPARATOR {0x6C 0}
            SUBTRACT {0x6D 0}   DECIMAL {0x6E 0}   DIVIDE {0x6F 0}
            F1 {0x70 0}   F2 {0x71 0}   F3 {0x72 0}   F4 {0x73 0}
            F5 {0x74 0}   F6 {0x75 0}   F7 {0x76 0}   F8 {0x77 0}
            F9 {0x78 0}   F10 {0x79 0}   F11 {0x7A 0}   F12 {0x7B 0}
            F13 {0x7C 0}   F14 {0x7D 0}   F15 {0x7E 0}   F16 {0x7F 0}
            F17 {0x80 0}   F18 {0x81 0}   F19 {0x82 0}   F20 {0x83 0}
            F21 {0x84 0}   F22 {0x85 0}   F23 {0x86 0}   F24 {0x87 0}
            NUMLOCK {0x90 0}   SCROLL {0x91 0}   SCROLLLOCK {0x91 0}
            LSHIFT {0xA0 0}   RSHIFT {0xA1 0 -extended}   LCONTROL {0xA2 0}
            RCONTROL {0xA3 0 -extended}   LMENU {0xA4 0}   LALT {0xA4 0}
            RMENU {0xA5 0 -extended}   RALT {0xA5 0 -extended}
            BROWSER_BACK {0xA6 0}   BROWSER_FORWARD {0xA7 0}
            BROWSER_REFRESH {0xA8 0}   BROWSER_STOP {0xA9 0}
            BROWSER_SEARCH {0xAA 0}   BROWSER_FAVORITES {0xAB 0}
            BROWSER_HOME {0xAC 0}   VOLUME_MUTE {0xAD 0}
            VOLUME_DOWN {0xAE 0}   VOLUME_UP {0xAF 0}
            MEDIA_NEXT_TRACK {0xB0 0}   MEDIA_PREV_TRACK {0xB1 0}
            MEDIA_STOP {0xB2 0}   MEDIA_PLAY_PAUSE {0xB3 0}
            LAUNCH_MAIL {0xB4 0}   LAUNCH_MEDIA_SELECT {0xB5 0}
            LAUNCH_APP1 {0xB6 0}   LAUNCH_APP2 {0xB7 0}
        }
    }
}

# Find the next token from a send_keys argument
# Returns pair token,position after token
proc twapi::_parse_send_key_token {keys start} {
    set char [string index $keys $start]
    if {$char ne "\{"} {
        return [list $char [incr start]]
    }
    # Need to find the matching end brace. Note special case of
    # start/end brace enclosed within braces
    set n [string length $keys]
    # Jump past brace and succeeding character (which may be end brace)
    set terminator [string first "\}" $keys $start+2]
    if {$terminator < 0} {
        error "Unterminated or empty braced key token."
    }
    return [list [string range $keys $start $terminator] [incr terminator]]
}

# Appends to inputs the trailer in reverse order. trailer is reset
proc twapi::_flush_send_keys_trailer {vinputs vtrailer} {
    upvar 1 $vinputs inputs
    upvar 1 $vtrailer trailer

    lappend inputs {*}[lreverse $trailer]
    set trailer {}
}

# Constructs a list of input events by parsing a string in the format
# used by Visual Basic's SendKeys function. See that documentation
# for syntax.
proc twapi::_parse_send_keys {keys} {
    variable vk_map

    _init_vk_map
    array set modifier_vk {+ 0x10 ^ 0x11 % 0x12}

    # Array state holds state of the parse. An atom refers to a single
    # character or a () group.
    #  modifiers - list of current modifiers in order they were added including
    #     those coming from containing groups.
    #  group_modifiers - stack of modifiers state when parsing groups.
    #     When a group begins, state(modifiers) is pushed on this stack.
    #     The top of the stack is used to initialize state(modifiers)
    #     for every atom within the group. When the group ends,
    #     the top of the stack is popped and discarded and state(modifiers)
    #     is reinitialized to new top of stack.
    #  trailer - list of trailing input records to add after next atom. Note
    #    these are stored in order of occurence but need to be reversed
    #    when emitted
    #  group_trailers - stack of trailers to add after group ends. Each
    #    element is a trailer which is a list of input records.
    #  cleanup_trailer - to be emitted right at the end if we have to
    #    reset CAPSLOCK/NUMLOCK/SCROLL
    set state(modifiers) {}
    set state(group_modifiers) [list $state(modifiers)]; # "Global" group
    set state(trailer) {}
    set state(group_trailers) {}
    set state(cleanup_trailer) {}

    set inputs {}

    # If {CAPS,NUM,SCROLL}LOCK are set, need to reset them and then
    # set them back
    foreach vk {20 144 145} {
        if {[GetKeyState $vk]} {
            lappend inputs [list key $vk 0]
            lappend state(cleanup_trailer) [list key $vk 0]
        }
    }

    set keyslen [string length $keys]
    set pos 0;                  # Current parse position
    while {$pos < $keyslen} {
        lassign [_parse_send_key_token $keys $pos] token pos
        switch -exact -- $token {
            + -
            ^ -
            % {
                if {$token in $state(modifiers)} {
                    # Following VB SendKeys
                    error "Modifier state for $token already set."
                }
                lappend state(modifiers) $token
                lappend inputs [list keydown $modifier_vk($token) 0]
                lappend state(trailer) [list keyup $modifier_vk($token) 0]
            }
            "(" {
                # Start a group
                lappend state(group_modifiers) $state(modifiers)
                lappend state(group_trailers) $state(trailer)
                set state(trailer) {}
            }
            ")" {
                # Terminates group. Illegal if no group collection in progress
                if {[llength $state(group_trailers)] == 0} {
                    error "Unmatched \")\" in send_keys string."
                }
                # If there is a live trailer inside group, emit it e.g. +(ab^)
                _flush_send_keys_trailer inputs state(trailer)
                # Now emit the group trailer
                set trailer [lpop state(group_trailers)]
                _flush_send_keys_trailer inputs trailer
                # Discard the initial modifier state for this group
                lpop state(group_modifiers)
                # Set the current modifiers to outer group state
                set state(modifiers) [lindex $state(group_modifiers) end]
            }
            default {
                if {$token eq "~"} {
                    set token "{ENTER}"
                }
                # May be a single character to send, a braced virtual key
                # or a braced single char with count
                if {[string length $token] == 1} {
                    # Single character.
                    set key $token
                    set nch 1
                } elseif {[string index $token 0] eq "\{"} {
                    # NOTE: a ~ inside a brace is treated as a literal ~
                    # and not the ENTER key
                    # Look for space skipping the starting brace and following
                    # character which may be itself a space (to be repeated)
                    set space_pos [string first " " $token 2]
                    if {$space_pos < 0} {
                        # No space found
                        set nch 1
                        set key [string range $token 1 end-1]
                    } else {
                        # A key followed by a count
                        # Note space_pos >= 2
                        set key [string range $token 1 $space_pos-1]
                        set nch [string trim [string range $token $space_pos+1 end-1]]
                        if {![string is integer -strict $nch] || $nch < 0} {
                            error "Invalid count \"$nch\" in send_keys."
                        } 
                    }
                } else {
                    # Problem in token parsing. Would be a bug.
                    error "Internal error: invalid token \"$token\" parsing send_keys string."
                }

                set vk_leader  {}
                set vk_trailer {}
                if {[string length $key] == 1} {
                    # Single character
                    lassign [VkKeyScan $key] modifiers vk
                    if {$modifiers == -1 || $vk == -1} {
                        scan $key %c code_point
                        set vk_rec [list unicode 0 $code_point]
                    } else {
                        # Generates input records for modifiers that are set
                        # unless they are already set. NOTE: Do NOT set the
                        # state(modifier) state since they will be in effect
                        # only for the current character. This is for correctly
                        # showing A-Z with shift and Ctrl-A etc. with control.
                        if {($modifiers & 0x1) && ("+" ni $state(modifiers))} {
                            lappend vk_leader  [list keydown 0x10 0]
                            lappend vk_trailer [list keyup 0x10 0]
                        }
                        if {($modifiers & 0x2) && ("^" ni $state(modifiers))} {
                            lappend vk_leader  [list keydown 0x11 0]
                            lappend vk_trailer [list keyup 0x11 0]
                        }

                        if {($modifiers & 0x4) && ("%" ni $state(modifiers))} {
                            lappend vk_leader  [list keydown 0x12 0]
                            lappend vk_trailer [list keyup 0x12 0]
                        }
                        set vk_rec [list key $vk 0]
                    }
                } else {
                    # Virtual key string. Note modifiers ignored here
                    # as for VB SendKeys
                    if {[info exists vk_map($key)]} {
                        # Virtual key
                        set vk_rec [list key {*}$vk_map($key)]
                    } else {
                        error "Unknown braced virtual key \"$token\"."
                    }
                }
                lappend inputs {*}$vk_leader
                lappend inputs {*}[lrepeat $nch $vk_rec]
                # vk_trailer arises from the character itself, e.g. A
                # has shift set, Ctrl-A has control set.
                _flush_send_keys_trailer inputs vk_trailer
                # state(trailer) arises from preceding +,^,% This is also
                # emitted and reset as it applied only to this character
                _flush_send_keys_trailer inputs state(trailer)
                set state(modifiers) [lindex $state(group_modifiers) end]
            }
        }
    }
    # Emit left over trailer
    _flush_send_keys_trailer inputs state(trailer)

    # Restore capslock/numlock
    _flush_send_keys_trailer inputs state(cleanup_trailer)

    return $inputs
}

# utility procedure to map symbolic hotkey to {modifiers virtualkey}
# We allow modifier map to be passed in because different api's use
# different bits for key modifiers
proc twapi::_hotkeysyms_to_vk {hotkey {modifier_map {ctrl 2 control 2 alt 1 menu 1 shift 4 win 8}}} {
    variable vk_map

    _init_vk_map

    set keyseq [split [string tolower $hotkey] -]
    set key [lindex $keyseq end]

    # Convert modifiers to bitmask
    set modifiers 0
    foreach modifier [lrange $keyseq 0 end-1] {
        setbits modifiers [dict! $modifier_map [string tolower $modifier]]
    }
    # Map the key to a virtual key code
    if {[string length $key] == 1} {
        # Single character
        scan $key %c unicode

        # Only allow alphanumeric keys and a few punctuation symbols
        # since keyboard layouts are not standard
        if {$unicode >= 0x61 && $unicode <= 0x7A} {
            # Lowercase letters - change to upper case virtual keys
            set vk [expr {$unicode-32}]
        } elseif {($unicode >= 0x30 && $unicode <= 0x39)
                  || ($unicode >= 0x41 && $unicode <= 0x5A)} {
            # Digits or upper case
            set vk $unicode
        } else {
            error "Only alphanumeric characters may be specified for the key. For non-alphanumeric characters, specify the virtual key code"
        }
    } elseif {[info exists vk_map($key)]} {
        # It is a virtual key name
        set vk [lindex $vk_map($key) 0]
    } elseif {[info exists vk_map([string toupper $key])]} {
        # It is a virtual key name
        set vk [lindex $vk_map([string toupper $key]) 0]
    } elseif {[string is integer -strict $key]} {
        # Actual virtual key specification
        set vk $key
    } else {
        error "Unknown or invalid key specifier '$key'"
    }

    return [list $modifiers $vk]
}
