#
# Copyright (c) 2010, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Implementation of named pipes


namespace eval twapi::pipe {


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
        array set handles {}
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

    proc server {name args} {

        _initialize_module

        variable pipes
        variable channels
        variable handles
        
        if {[dict exists $names $name]} {
            if {[llength $args]} {
                # TBD - should we allow caller to respecify -secattr
                error "Options may only be specified for the first creation of a named pipe."
            }
            set options [dict get $pipes options]
        } else {
            array set opts [parseargs args {
                {mode.arg {read write}}
                writedacl
                writeowner
                writesacl
                writethrough
                {readmode.arg byte {byte message}}
                {writemode.arg byte {byte message}}
                {timeout.int 50}
                {maxinstances.int 255}
                {secattr.arg {}}
            } -maxleftover 0]

            set open_mode [twapi::_parse_symbolic_bitmask $opts(mode) {read 1 write 2}]
            set open_mode [twapi::_switches_to_bitmask opts {
                writedacl  0x00040000
                writeowner 0x00080000
                writesacl  0x01000000
                writethrough 0x80000000
            } $open_mode]
        
            set open_mode [expr {$open_mode | 0x40000000}]; # OVERLAPPPED I/O
        
            set pipe_mode 0
            if {$opts(readmode) eq "message"} {
                setbits pipe_mode 2
            }
            if {$opts(writemode) eq "message"} {
                setbits pipe_mode 4
            }

            set options [dict create \
                             chan_mode $opts(mode) \
                             open_mode $open_mode \
                             pipe_mode $pipe_mode \
                             max_instances $opts(maxinstances) \
                             timeout $opts(timeout) \
                             secattr $opts(secattr)]
        }

        set hpipe [twapi::Twapi_PipeServer $name \
                       [dict get $options open_mode] \
                       [dict get $options pipe_mode] \
                       [dict get $options max_instances] \
                       4096 4096 \
                       [dict get $options timeout] \
                       [dict get $options(secattr)]]
        
        # Now create a Tcl channel based on the OS handle
        if {[catch {
            chan create [dict get $pipes $name options chan_mode] [list [namespace current]]} chan]} {
            set einfo $::errorInfo
            set ecode $::errorcode
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
        dict set handles $hpipe channel $chan
        
        return $chan
    }

    proc initialize {chan mode} {
        # Whatever initialization is to be done is already done in the
        # server command. Nothing much more to be done here except init
        # some channel state.
        variable channels

        dict set channels $chan watch_read 0
        dict set channels $chan watch_write 0

        return [list initialize finalize watch read write configure cget cgetall blocking]
    }

    proc finalize {chan} {
        variable channels
        variable pipes
        variable handles
        
        # Unlink this instance from the named pipe. If the named pipe
        # has no more instances, delete it as well
        if {[dict exists $channels $chan pipe_name]} {
            set name [dict get $channels $chan pipe_name]
            dict unset $pipes $name channels $chan
            if {[dict size [dict get $pipes $name channels]] == 0} {
                # No more instances, get rid of the whole named pipe
                dict unset $pipes $name
            }
        }

        set handle [dict get $channels $chan handle]
        dict unset handles $handle
            
        dict unset channels $chan

        # Close the OS handle at the end
        if {[catch {twapi::Twapi_ClosePipe $handle} msg]} {
            twapi::log "Error closing named pipe: $msg"
        }

        return
    }

    proc watch {chan events} {
        variable channels

        dict set channels $chan watch_read  0
        dict set channels $chan watch_write 0
        foreach event $events {
            dict set channels $chan watch_$event 1
        }
    }

    proc read {chanid count} {
        variable chan
        puts [info level 0]
        if {[string length $chan($chanid)] < $count} {
            set result $chan($chanid); set chan($chanid) ""
        } else {
            set result [string range $chan($chanid) 0 $count-1]
            set chan($chanid) [string range $chan($chanid) $count end]
        }

        # implement max buffering
        variable watching
        variable max
        if {$watching(write) && ([string length $chan($chanid)] < $max)} {
            chan postevent $chanid write
        }

        return $result
    }

    variable max 1048576        ;# maximum size of the reflected channel

    proc write {chanid data} {
        variable chan
        variable max
        variable watching

        puts [info level 0]

        set left [expr {$max - [string length $chan($chanid)]}]        ;# bytes left in buffer
        set dsize [string length $data]
        if {$left >= $dsize} {
            append chan($chanid) $data
            if {$watching(write) && ([string length $chan($chanid)] < $max)} {
                # inform the app that it may still write
                chan postevent $chanid write
            }
        } else {
            set dsize $left
            append chan($chanid) [string range $data $left]
        }

        # inform the app that there's something to read
        if {$watching(read) && ($chan($chanid) ne "")} {
            puts "post event read"
            chan postevent $chanid read
        }

        return $dsize        ;# number of bytes actually written
    }

    proc blocking { chanid args } {
        variable chan

        puts [info level 0]
    }

    proc cget { chanid args } {
        variable chan

        puts [info level 0]
    }

    proc cgetall { chanid args } {
        variable chan

        puts [info level 0]
    }

    proc configure { chanid args } {
        variable chan

        puts [info level 0]
    }

    namespace export -clear *
    namespace ensemble create -subcommands {}
}