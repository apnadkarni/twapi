#
# Copyright (c) 2012 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

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
    variable windefs

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
                set flags $windefs(MOUSEEVENTF_ABSOLUTE)
            }

            if {[info exists opts(wheel)]} {
                if {($opts(x1down) || $opts(x1up) || $opts(x2down) || $opts(x2up))} {
                    error "The -wheel input event attribute may not be specified with -x1up, -x1down, -x2up or -x2down events"
                }
                set mousedata $opts(wheel)
                set flags $windefs(MOUSEEVENTF_WHEEL)
            } else {
                if {$opts(x1down) || $opts(x1up)} {
                    if {$opts(x2down) || $opts(x2up)} {
                        error "The -x1down, -x1up mouse input attributes are mutually exclusive with -x2down, -x2up attributes"
                    }
                    set mousedata $windefs(XBUTTON1)
                } else {
                    if {$opts(x2down) || $opts(x2up)} {
                        set mousedata $windefs(XBUTTON2)
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
                    set flags [expr {$flags | $windefs(MOUSEEVENTF_$flag)}]
                }
            }

            lappend inputs [list mouse $xpos $ypos $mousedata $flags]

        } else {
            lassign $input inputtype vk scan keyopts
            if {"-extended" ni $keyopts} {
                set extended 0
            } else {
                set extended $windefs(KEYEVENTF_EXTENDEDKEY)
            }
            if {"-usescan" ni $keyopts} {
                set usescan 0
            } else {
                set usescan $windefs(KEYEVENTF_SCANCODE)
            }
            switch -exact -- $inputtype {
                keydown {
                    lappend inputs [list key $vk $scan [expr {$extended|$usescan}]]
                }
                keyup {
                    lappend inputs [list key $vk $scan \
                                        [expr {$extended
                                               | $usescan
                                               | $windefs(KEYEVENTF_KEYUP)
                                           }]]
                }
                key {
                    lappend inputs [list key $vk $scan [expr {$extended|$usescan}]]
                    lappend inputs [list key $vk $scan \
                                        [expr {$extended
                                               | $usescan
                                               | $windefs(KEYEVENTF_KEYUP)
                                           }]]
                }
                unicode {
                    lappend inputs [list key 0 $scan $windefs(KEYEVENTF_UNICODE)]
                    lappend inputs [list key 0 $scan \
                                        [expr {$windefs(KEYEVENTF_UNICODE)
                                               | $windefs(KEYEVENTF_KEYUP)
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
        error "Error registering hotkey." $errorInfo $errorCode
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
    # we temporarily disable mouse trails

    if {[min_os_version 5 1]} {
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

    # Restore trail setting
    if {[min_os_version 5 1]} {
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

    if {$now >= $last_event} {
        return [expr {$now - $last_event}]
    } else {
        return [expr {$now + (0xffffffff - $last_event) + 1}]
    }
}

# Initialize the virtual key table
proc twapi::_init_vk_map {} {
    variable windefs
    variable vk_map

    if {![info exists vk_map]} {
        array set vk_map [list \
                              "+" [list $windefs(VK_SHIFT) 0]\
                              "^" [list $windefs(VK_CONTROL) 0] \
                              "%" [list $windefs(VK_MENU) 0] \
                              "BACK" [list $windefs(VK_BACK) 0] \
                              "BACKSPACE" [list $windefs(VK_BACK) 0] \
                              "BS" [list $windefs(VK_BACK) 0] \
                              "BKSP" [list $windefs(VK_BACK) 0] \
                              "TAB" [list $windefs(VK_TAB) 0] \
                              "CLEAR" [list $windefs(VK_CLEAR) 0] \
                              "RETURN" [list $windefs(VK_RETURN) 0] \
                              "ENTER" [list $windefs(VK_RETURN) 0] \
                              "SHIFT" [list $windefs(VK_SHIFT) 0] \
                              "CONTROL" [list $windefs(VK_CONTROL) 0] \
                              "MENU" [list $windefs(VK_MENU) 0] \
                              "ALT" [list $windefs(VK_MENU) 0] \
                              "PAUSE" [list $windefs(VK_PAUSE) 0] \
                              "BREAK" [list $windefs(VK_PAUSE) 0] \
                              "CAPITAL" [list $windefs(VK_CAPITAL) 0] \
                              "CAPSLOCK" [list $windefs(VK_CAPITAL) 0] \
                              "KANA" [list $windefs(VK_KANA) 0] \
                              "HANGEUL" [list $windefs(VK_HANGEUL) 0] \
                              "HANGUL" [list $windefs(VK_HANGUL) 0] \
                              "JUNJA" [list $windefs(VK_JUNJA) 0] \
                              "FINAL" [list $windefs(VK_FINAL) 0] \
                              "HANJA" [list $windefs(VK_HANJA) 0] \
                              "KANJI" [list $windefs(VK_KANJI) 0] \
                              "ESCAPE" [list $windefs(VK_ESCAPE) 0] \
                              "ESC" [list $windefs(VK_ESCAPE) 0] \
                              "CONVERT" [list $windefs(VK_CONVERT) 0] \
                              "NONCONVERT" [list $windefs(VK_NONCONVERT) 0] \
                              "ACCEPT" [list $windefs(VK_ACCEPT) 0] \
                              "MODECHANGE" [list $windefs(VK_MODECHANGE) 0] \
                              "SPACE" [list $windefs(VK_SPACE) 0] \
                              "PRIOR" [list $windefs(VK_PRIOR) 0] \
                              "PGUP" [list $windefs(VK_PRIOR) 0] \
                              "NEXT" [list $windefs(VK_NEXT) 0] \
                              "PGDN" [list $windefs(VK_NEXT) 0] \
                              "END" [list $windefs(VK_END) 0] \
                              "HOME" [list $windefs(VK_HOME) 0] \
                              "LEFT" [list $windefs(VK_LEFT) 0] \
                              "UP" [list $windefs(VK_UP) 0] \
                              "RIGHT" [list $windefs(VK_RIGHT) 0] \
                              "DOWN" [list $windefs(VK_DOWN) 0] \
                              "SELECT" [list $windefs(VK_SELECT) 0] \
                              "PRINT" [list $windefs(VK_PRINT) 0] \
                              "PRTSC" [list $windefs(VK_SNAPSHOT) 0] \
                              "EXECUTE" [list $windefs(VK_EXECUTE) 0] \
                              "SNAPSHOT" [list $windefs(VK_SNAPSHOT) 0] \
                              "INSERT" [list $windefs(VK_INSERT) 0] \
                              "INS" [list $windefs(VK_INSERT) 0] \
                              "DELETE" [list $windefs(VK_DELETE) 0] \
                              "DEL" [list $windefs(VK_DELETE) 0] \
                              "HELP" [list $windefs(VK_HELP) 0] \
                              "LWIN" [list $windefs(VK_LWIN) 0] \
                              "RWIN" [list $windefs(VK_RWIN) 0] \
                              "APPS" [list $windefs(VK_APPS) 0] \
                              "SLEEP" [list $windefs(VK_SLEEP) 0] \
                              "NUMPAD0" [list $windefs(VK_NUMPAD0) 0] \
                              "NUMPAD1" [list $windefs(VK_NUMPAD1) 0] \
                              "NUMPAD2" [list $windefs(VK_NUMPAD2) 0] \
                              "NUMPAD3" [list $windefs(VK_NUMPAD3) 0] \
                              "NUMPAD4" [list $windefs(VK_NUMPAD4) 0] \
                              "NUMPAD5" [list $windefs(VK_NUMPAD5) 0] \
                              "NUMPAD6" [list $windefs(VK_NUMPAD6) 0] \
                              "NUMPAD7" [list $windefs(VK_NUMPAD7) 0] \
                              "NUMPAD8" [list $windefs(VK_NUMPAD8) 0] \
                              "NUMPAD9" [list $windefs(VK_NUMPAD9) 0] \
                              "MULTIPLY" [list $windefs(VK_MULTIPLY) 0] \
                              "ADD" [list $windefs(VK_ADD) 0] \
                              "SEPARATOR" [list $windefs(VK_SEPARATOR) 0] \
                              "SUBTRACT" [list $windefs(VK_SUBTRACT) 0] \
                              "DECIMAL" [list $windefs(VK_DECIMAL) 0] \
                              "DIVIDE" [list $windefs(VK_DIVIDE) 0] \
                              "F1" [list $windefs(VK_F1) 0] \
                              "F2" [list $windefs(VK_F2) 0] \
                              "F3" [list $windefs(VK_F3) 0] \
                              "F4" [list $windefs(VK_F4) 0] \
                              "F5" [list $windefs(VK_F5) 0] \
                              "F6" [list $windefs(VK_F6) 0] \
                              "F7" [list $windefs(VK_F7) 0] \
                              "F8" [list $windefs(VK_F8) 0] \
                              "F9" [list $windefs(VK_F9) 0] \
                              "F10" [list $windefs(VK_F10) 0] \
                              "F11" [list $windefs(VK_F11) 0] \
                              "F12" [list $windefs(VK_F12) 0] \
                              "F13" [list $windefs(VK_F13) 0] \
                              "F14" [list $windefs(VK_F14) 0] \
                              "F15" [list $windefs(VK_F15) 0] \
                              "F16" [list $windefs(VK_F16) 0] \
                              "F17" [list $windefs(VK_F17) 0] \
                              "F18" [list $windefs(VK_F18) 0] \
                              "F19" [list $windefs(VK_F19) 0] \
                              "F20" [list $windefs(VK_F20) 0] \
                              "F21" [list $windefs(VK_F21) 0] \
                              "F22" [list $windefs(VK_F22) 0] \
                              "F23" [list $windefs(VK_F23) 0] \
                              "F24" [list $windefs(VK_F24) 0] \
                              "NUMLOCK" [list $windefs(VK_NUMLOCK) 0] \
                              "SCROLL" [list $windefs(VK_SCROLL) 0] \
                              "SCROLLLOCK" [list $windefs(VK_SCROLL) 0] \
                              "LSHIFT" [list $windefs(VK_LSHIFT) 0] \
                              "RSHIFT" [list $windefs(VK_RSHIFT) 0 -extended] \
                              "LCONTROL" [list $windefs(VK_LCONTROL) 0] \
                              "RCONTROL" [list $windefs(VK_RCONTROL) 0 -extended] \
                              "LMENU" [list $windefs(VK_LMENU) 0] \
                              "LALT" [list $windefs(VK_LMENU) 0] \
                              "RMENU" [list $windefs(VK_RMENU) 0 -extended] \
                              "RALT" [list $windefs(VK_RMENU) 0 -extended] \
                              "BROWSER_BACK" [list $windefs(VK_BROWSER_BACK) 0] \
                              "BROWSER_FORWARD" [list $windefs(VK_BROWSER_FORWARD) 0] \
                              "BROWSER_REFRESH" [list $windefs(VK_BROWSER_REFRESH) 0] \
                              "BROWSER_STOP" [list $windefs(VK_BROWSER_STOP) 0] \
                              "BROWSER_SEARCH" [list $windefs(VK_BROWSER_SEARCH) 0] \
                              "BROWSER_FAVORITES" [list $windefs(VK_BROWSER_FAVORITES) 0] \
                              "BROWSER_HOME" [list $windefs(VK_BROWSER_HOME) 0] \
                              "VOLUME_MUTE" [list $windefs(VK_VOLUME_MUTE) 0] \
                              "VOLUME_DOWN" [list $windefs(VK_VOLUME_DOWN) 0] \
                              "VOLUME_UP" [list $windefs(VK_VOLUME_UP) 0] \
                              "MEDIA_NEXT_TRACK" [list $windefs(VK_MEDIA_NEXT_TRACK) 0] \
                              "MEDIA_PREV_TRACK" [list $windefs(VK_MEDIA_PREV_TRACK) 0] \
                              "MEDIA_STOP" [list $windefs(VK_MEDIA_STOP) 0] \
                              "MEDIA_PLAY_PAUSE" [list $windefs(VK_MEDIA_PLAY_PAUSE) 0] \
                              "LAUNCH_MAIL" [list $windefs(VK_LAUNCH_MAIL) 0] \
                              "LAUNCH_MEDIA_SELECT" [list $windefs(VK_LAUNCH_MEDIA_SELECT) 0] \
                              "LAUNCH_APP1" [list $windefs(VK_LAUNCH_APP1) 0] \
                              "LAUNCH_APP2" [list $windefs(VK_LAUNCH_APP2) 0] \
                             ]
    }

}


# Constructs a list of input events by parsing a string in the format
# used by Visual Basic's SendKeys function
proc twapi::_parse_send_keys {keys {inputs ""}} {
    variable vk_map

    _init_vk_map

    set n [string length $keys]
    set trailer [list ]
    for {set i 0} {$i < $n} {incr i} {
        set key [string index $keys $i]
        switch -exact -- $key {
            "+" -
            "^" -
            "%" {
                lappend inputs [concat keydown $vk_map($key)]
                set trailer [linsert $trailer 0 [concat keyup $vk_map($key)]]
            }
            "~" {
                lappend inputs [concat key $vk_map(RETURN)]
                set inputs [concat $inputs $trailer]
                set trailer [list ]
            }
            "(" {
                # Recurse for paren expression
                set nextparen [string first ")" $keys $i]
                if {$nextparen == -1} {
                    error "Invalid key sequence - unterminated ("
                }
                set inputs [concat $inputs [_parse_send_keys [string range $keys [expr {$i+1}] [expr {$nextparen-1}]]]]
                set inputs [concat $inputs $trailer]
                set trailer [list ]
                set i $nextparen
            }
            "\{" {
                set nextbrace [string first "\}" $keys $i]
                if {$nextbrace == -1} {
                    error "Invalid key sequence - unterminated $key"
                }

                if {$nextbrace == ($i+1)} {
                    # Look for the next brace
                    set nextbrace [string first "\}" $keys $nextbrace]
                    if {$nextbrace == -1} {
                        error "Invalid key sequence - unterminated $key"
                    }
                }

                set key [string range $keys [expr {$i+1}] [expr {$nextbrace-1}]]
                set bracepat [string toupper $key]
                if {[info exists vk_map($bracepat)]} {
                    lappend inputs [concat key $vk_map($bracepat)]
                } else {
                    # May be pattern of the type {C} or {C N} where
                    # C is a single char and N is a count
                    set c [string index $key 0]
                    set count [string trim [string range $key 1 end]]
                    scan $c %c unicode
                    if {[string length $count] == 0} {
                        set count 1
                    } else {
                        # Note if $count is not an integer, an error
                        # will be generated as we want
                        incr count 0
                        if {$count < 0} {
                            error "Negative character count specified in braced key input"
                        }
                    }
                    for {set j 0} {$j < $count} {incr j} {
                        lappend inputs [list unicode 0 $unicode]
                    }
                }
                set inputs [concat $inputs $trailer]
                set trailer [list ]
                set i $nextbrace
            }
            default {
                scan $key %c unicode
                # Alphanumeric keys are treated separately so they will
                # work correctly with control modifiers
                if {$unicode >= 0x61 && $unicode <= 0x7A} {
                    # Lowercase letters
                    lappend inputs [list key [expr {$unicode-32}] 0]
                } elseif {$unicode >= 0x30 && $unicode <= 0x39} {
                    # Digits
                    lappend inputs [list key $unicode 0]
                } else {
                    lappend inputs [list unicode 0 $unicode]
                }
                set inputs [concat $inputs $trailer]
                set trailer [list ]
            }
        }
    }
    return $inputs
}

# utility procedure to map symbolic hotkey to {modifiers virtualkey}
proc twapi::_hotkeysyms_to_vk {hotkey} {
    variable vk_map

    _init_vk_map

    set keyseq [split [string tolower $hotkey] -]
    set key [lindex $keyseq end]

    # Convert modifiers to bitmask
    set modifiers 0
    foreach modifier [lrange $keyseq 0 end-1] {
        switch -exact -- [string tolower $modifier] {
            ctrl -
            control {
                setbits modifiers 2
            }

            alt -
            menu {
                setbits modifiers 1
            }

            shift {
                setbits modifiers 4
            }

            win {
                setbits modifiers 8
            }

            default {
                error "Unknown key modifier $modifier"
            }
        }
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
    } elseif {[string is integer $key]} {
        # Actual virtual key specification
        set vk $key
    } else {
        error "Unknown or invalid key specifier '$key'"
    }

    return [list $modifiers $vk]
}
