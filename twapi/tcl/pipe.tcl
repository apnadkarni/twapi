#
# Copyright (c) 2010, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Implementation of named pipes

if {! [package vsatisfies [package require Tcl] 8.5]} {
    proc twapi::pipe args {
        error "This command requires V8.5 or later of Tcl."
    }
} else {
    proc twapi::pipe {name args} {
        array set opts [parseargs args {
            server.arg
        }]
        if {[info exists opts(server)]} {
            return [twapi::pipechan server $name $opts(server) {*}$args]
        } else {
            return [twapi::pipechan client $name {*}$args]
        }
    }
}


namespace eval twapi::pipechan {

    # Initializes the pipe module
    proc _initialize_module {} {
        # A pipe is identified by its name. Variable pipes tracks these.
        # There can be multiple instances of a pipe. These tracked
        # in through the handles variable, which maps OS handles to
        # pipe instances, and channels, which maps Tcl channels to
        # pipe instances.
        variable pipes
        set pipes [dict create]
        variable handles
        set handles [dict create]
        variable channels
        set channels [dict create]

        # Create the ensemble used for reflected channels
        namespace ensemble create \
            -command [namespace current] \
            -subcommands {
                initialize finalize watch read write configure cget
                cgetall blocking
            }
        # Redefine ourselves to be a no-op
        proc [namespace current]::_initialize_module {} {}
    }

    proc server {name script args} {

        _initialize_module

        variable pipes
        variable channels
        variable handles
        
        set name [file nativename $name]

        if {[dict exists $pipes $name]} {
            if {[llength $args]} {
                # TBD - should we allow caller to respecify -secattr
                error "Options may only be specified for the first creation of a named pipe."
            }
            set options [dict get $pipes options]
        } else {
            # Only byte mode currently supported. Message mode does
            # not mesh well with Tcl channel infrastructure.
            # readmode.arg
            # writemode.arg

            array set opts [twapi::parseargs args {
                {mode.arg {read write}}
                writedacl
                writeowner
                writesacl
                writethrough
                denyremote
                {timeout.int 50}
                {maxinstances.int 255}
                {secattr.arg {}}
            } -maxleftover 0]

            set open_mode [twapi::_parse_symbolic_bitmask $opts(mode) {read 1 write 2}]
            foreach {opt mask} {
                writedacl  0x00040000
                writeowner 0x00080000
                writesacl  0x01000000
                writethrough 0x80000000
            } {
                if {$opts($opt)} {
                    set open_mode [expr {$open_mode | $mask}]
                }
            }
        
            set open_mode [expr {$open_mode | 0x40000000}]; # OVERLAPPPED I/O
        
            set pipe_mode 0
            if {$opts(denyremote)} {
                if {! [twapi::min_os_version 6]} {
                    error "Option -denyremote not supported on this operating system."
                }
                set pipe_mode [expr {$pipe_mode | 8}]
            }

            set options [dict create \
                             chan_mode $opts(mode) \
                             open_mode $open_mode \
                             pipe_mode $pipe_mode \
                             max_instances $opts(maxinstances) \
                             timeout $opts(timeout) \
                             secattr $opts(secattr)]

            dict set pipes $name options $options
        }

        set hpipe [twapi::Twapi_PipeServer $name \
                       [dict get $options open_mode] \
                       [dict get $options pipe_mode] \
                       [dict get $options max_instances] \
                       4096 4096 \
                       [dict get $options timeout] \
                       [dict get $options secattr]]
        
        # Now create a Tcl channel based on the OS handle
        if {[catch {
            chan create [dict get $pipes $name options chan_mode] [list [namespace current]]} chan]} {
            set einfo $::errorInfo
            set ecode $::errorCode
            catch {twapi::Twapi_PipeClose $hpipe}
            error $chan $einfo $ecode
        }
        
        if {![dict exists $pipes $name]} {
            dict set pipes $name options $options
        }

        # Cross link the various structures
        dict set pipes $name channels $chan {}
        dict set channels $chan pipe_name $name
        dict set channels $chan handle $hpipe
        dict set channels $chan direction server
        dict set handles $hpipe channel $chan

        # Register the callback script for server connection notification
        dict set channels $chan connect_notification_script $script
        
        # Now indicate readiness for a client connection.
        if {[catch {
            twapi::Twapi_PipeAccept $hpipe
        } msg]} {
            set erinfo $::errorInfo
            set ercode $::errorCode
            chan close $chan
            error $msg $erinfo $ercode
        }

        return $chan
    }

    proc initialize {chan mode} {
        # Whatever initialization is to be done is already done in the
        # server command. Nothing much more to be done here except init
        # some channel state.

        variable channels

        dict set channels $chan watch_read 0
        dict set channels $chan watch_write 0

        # Note commands cget, cgetall and configure are not currently
        # supported.
        set commands [list initialize finalize watch read write blocking]
    }

    proc finalize {chan} {
        variable channels
        variable pipes
        variable handles
        
        # Unlink this instance from the named pipe. If the named pipe
        # has no more instances, delete it as well
        if {[dict exists $channels $chan pipe_name]} {
            set name [dict get $channels $chan pipe_name]
            dict unset pipes $name channels $chan
            if {[dict size [dict get $pipes $name channels]] == 0} {
                # No more instances, get rid of the whole named pipe
                dict unset pipes $name
            }
        }
        
        set handle [dict get $channels $chan handle]
        dict unset handles $handle

        dict unset channels $chan

        # Close the OS handle at the end
        if {[catch {twapi::Twapi_PipeClose $handle} msg]} {
            twapi::log "Error closing named pipe: $msg"
        }

        return
    }

    proc watch {chan events} {
        variable channels

        set read [expr {"read" in $events}]
        dict set channels $chan watch_read $read

        set write [expr {"write" in $events}]
        dict set channels $chan watch_write $write
        
        twapi::Twapi_PipeWatch [chan2handle $chan] $read $write
    }

    proc read {chan count} {
        return [twapi::Twapi_PipeRead [chan2handle $chan] $count]
    }

    proc write {chan data} {
        return [twapi::Twapi_PipeWrite [chan2handle $chan] $data]
    }

    proc blocking {chan blocking} {
        return [twapi::Twapi_PipeSetBlockMode [chan2handle $chan] $blocking]
    }

    # Callback from C level event handler in the background
    proc _pipe_handler {hpipe event} {
        variable handles
        variable channels
twapi::log "_pipe_handler: $hpipe $event"
        if {![dict exists $handles $hpipe channel]} {
            return;             # Assume async channel close
        }
        set chan [dict get $handles $hpipe channel]
        switch -exact -- $event {
            connect {
                eval [dict get $channels $chan connect_notification_script] [list $chan connect]
            }
            read {
                if {[dict get $channels $chan watch_read]} {
                    chan postevent $chan read
                }
            }
            write {
                if {[dict get $channels $chan watch_write]} {
                    chan postevent $chan write
                }
            }
        }
    }

    proc chan2handle {chan} {
        variable channels

        if {! [dict exists $channels $chan handle]} {
            twapi::win32_error ERROR_BAD_PIPE "Channel does not exist."

        }
        return [dict get $channels $chan handle]
    }

    namespace export -clear *
    if {[package vsatisfies [info tclversion] 8.5]} {
        namespace ensemble create -subcommands {}
    }
}