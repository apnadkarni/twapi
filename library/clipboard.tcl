#
# Copyright (c) 2004, 2008 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Clipboard related commands

namespace eval twapi {}

# Open the clipboard
# TBD - why no mechanism to pass window handle to OpenClipboard?
proc twapi::open_clipboard {} {
    OpenClipboard 0
}

# Close the clipboard
proc twapi::close_clipboard {} {
    catch {CloseClipboard}
    return
}

# Empty the clipboard
proc twapi::empty_clipboard {} {
    EmptyClipboard
}

proc twapi::_read_clipboard {fmt} {
    # Always catch errors and close clipboard before passing exception on
    # Also ensure memory unlocked
    trap {
        set h [GetClipboardData $fmt]
        set p [GlobalLock $h]
        set data [Twapi_ReadMemory 1 $p 0 [GlobalSize $h]]
    } onerror {} {
        catch {close_clipboard}
        rethrow
    } finally {
        # If p exists, then we must have locked the handle
        if {[info exists p]} {
            GlobalUnlock $h
        }
    }
    return $data
}

proc twapi::read_clipboard {fmt} {
    trap {
        set data [_read_clipboard $fmt]
    } onerror {TWAPI_WIN32 1418} {
        # Caller did not have clipboard open. Do it on its behalf
        open_clipboard
        trap {
            set data [_read_clipboard $fmt]
        } finally {
            catch {close_clipboard}
        }
    }
    return $data
}

# Read text data from the clipboard
proc twapi::read_clipboard_text {args} {
    array set opts [parseargs args {
        {raw.bool 0}
    }]

    set bin [read_clipboard 13]; # 13 -> Unicode
    # Decode Unicode and discard trailing nulls
    set data [string trimright [encoding convertfrom unicode $bin] \0]
    if {! $opts(raw)} {
        set data [string map {"\r\n" "\n"} $data]
    }

    return $data
}

proc twapi::_write_clipboard {fmt data} {
    # Always catch errors and close
    # clipboard before passing exception on
    trap {
        # For byte arrays, string length does return correct size
        # (DO NOT USE string bytelength - see Tcl docs!)
        set len [string length $data]

        # Allocate global memory
        set mem_h [GlobalAlloc 2 $len]
        set mem_p [GlobalLock $mem_h]

        Twapi_WriteMemory 1 $mem_p 0 $len $data

        # The rest of this code just to ensure we do not free
        # memory beyond this point irrespective of error/success
        set h $mem_h
        unset mem_p mem_h
        GlobalUnlock $h
        SetClipboardData $fmt $h
    } onerror {} {
        catch close_clipboard
        rethrow
    } finally {
        if {[info exists mem_p]} {
            GlobalUnlock $mem_h
        }
        if {[info exists mem_h]} {
            GlobalFree $mem_h
        }
    }
    return
}

proc twapi::write_clipboard {fmt data} {
    trap {
        _write_clipboard $fmt $data
    } onerror {TWAPI_WIN32 1418} {
        # Caller did not have clipboard open. Do it on its behalf
        open_clipboard
        empty_clipboard
        trap {
            _write_clipboard $fmt $data
        } finally {
            catch close_clipboard
        }
    }
    return
}

# Write text to the clipboard
proc twapi::write_clipboard_text {data args} {
    array set opts [parseargs args {
        {raw.bool 0}
    }]

    # Convert \n to \r\n leaving existing \r\n alone
    if {! $opts(raw)} {
        set data [regsub -all {(^|[^\r])\n} $data[set data ""] \\1\r\n]
    }
    append data \0
    write_clipboard 13 [encoding convertto unicode $data]; # 13 -> Unicode
    return
}

# Get current clipboard formats
proc twapi::get_clipboard_formats {} {
    return [Twapi_EnumClipboardFormats]
}

# Get registered clipboard format name. Clipboard does not have to be open
proc twapi::get_registered_clipboard_format_name {fmt} {
    return [GetClipboardFormatName $fmt]
}

# Register a clipboard format
proc twapi::register_clipboard_format {fmt_name} {
    RegisterClipboardFormat $fmt_name
}

# Returns 1/0 depending on whether a format is on the clipboard. Clipboard
# does not have to be open
proc twapi::clipboard_format_available {fmt} {
    return [IsClipboardFormatAvailable $fmt]
}

proc twapi::read_clipboard_paths {} { 
    set bin [read_clipboard 15]
    # Extract the DROPFILES header 
    if {[binary scan $bin iiiii offset - - - unicode] != 5} { 
        error "Invalid or unsupported clipboard CF_DROP data." 
    } 
    # Sanity check 
    if {$offset >= [string length $bin]} { 
        error "Truncated clipboard data." 
    } 
    if {$unicode} { 
        set paths [encoding convertfrom unicode [string range $bin $offset end]] 
    } else { 
        set paths [encoding convertfrom ascii [string range $bin $offset end]] 
    } 
    set ret {}
    foreach path [split $paths \0] { 
        if {[string length $path] == 0} break; # Empty string -> end of list 
        lappend ret [file join $path]
    }
    return $ret
} 

proc twapi::write_clipboard_paths {paths} { 
    # The header for a DROPFILES path list in hex 
    set fheader "1400000000000000000000000000000001000000" 
    set bin [binary format H* $fheader] 
    foreach path $paths { 
        # Note explicit \0 so the encoded binary includes the null terminator 
        append bin [encoding convertto unicode "[file nativename [file normalize $path]]\0"] 
    } 
    # A Unicode null char to terminate the list of paths 
    append bin [encoding convertto unicode \0] 
    write_clipboard 15 $bin 
}

# Start monitoring of the clipboard
proc twapi::_clipboard_handler {} {
    variable _clipboard_monitors

    if {![info exists _clipboard_monitors] ||
        [llength $_clipboard_monitors] == 0} {
        return; # Not an error, could have deleted while already queued
    }

    foreach {id script} $_clipboard_monitors {
        set code [catch {uplevel #0 $script} msg]
        if {$code == 1} {
            # Error - put in background but we do not abort
            after 0 [list error $msg $::errorInfo $::errorCode]
        }
    }
    return
}

proc twapi::start_clipboard_monitor {script} {
    variable _clipboard_monitors

    set id "clip#[TwapiId]"
    if {![info exists _clipboard_monitors] ||
        [llength $_clipboard_monitors] == 0} {
        # No clipboard monitoring in progress. Start it
        Twapi_ClipboardMonitorStart
    }

    lappend _clipboard_monitors $id $script
    return $id
}



# Stop monitoring of the clipboard
proc twapi::stop_clipboard_monitor {clipid} {
    variable _clipboard_monitors

    if {![info exists _clipboard_monitors]} {
        return;                 # Should we raise an error instead?
    }

    set new_monitors {}
    foreach {id script} $_clipboard_monitors {
        if {$id ne $clipid} {
            lappend new_monitors $id $script
        }
    }

    set _clipboard_monitors $new_monitors
    if {[llength $_clipboard_monitors] == 0} {
        Twapi_ClipboardMonitorStop
    }
}
