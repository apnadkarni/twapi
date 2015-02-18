package require vix
vix::initialize
package require twapi
twapi::import_commands

array set Config {
    target_tcl_dir c:/twapitest/tcl
}
set Config(source_test_dir) [file dirname [info script]]
set Config(source_root_dir) [file dirname [file dirname $Config(source_test_dir)]]

proc config key {
    return $::Config($key)
}

proc usertask {message} {
    # Make sure we are seen
    if {[info commands twapi::set_foreground_window] ne "" &&
        [info commands twapi::get_console_window] ne ""} {
        twapi::set_foreground_window [twapi::get_console_window]
    }
    # Would like -nonewline here but see comments in proc yesno
    puts -nonewline "\n$message\nThen Hit Return to continue..."
    flush stdout
    gets stdin
    return
}

interp alias {} progress {} puts

oo::class create TestTarget {
    variable VM VMpath
    constructor {vmpath} {
        set VMpath $vmpath
        set VM [::host open $vmpath]
    }

    destructor {
        if {[info exists $VM]} {
            $VM destroy
        }
    }

    method vm {} { return $VM }
    method setup args {
        set vm_name [$VM name]

        progress "Powering on $vm_name"
        $VM power_on
        $VM wait_for_tools 60
        usertask "Please login manually at the console to the virtual machine $vm_name\nand wait for the login to complete."
        # Use the GUI dialog so it will work in both tclsh and tkcon
        lassign [credentials_dialog -showsaveoption 0 -target $vm_name -message "Enter the same credentials that you used to login at the console"] user password
        $VM login $user [reveal $password] -interactive 1

        set tcl_root [file dirname [file dirname [file normalize [info nameofexecutable]]]]
        set target_dir [config target_tcl_dir]
        set target_bindir [file join $target_dir bin]
        set target_libdir [file join $target_dir lib]

        # Create dir structure
        $VM rmdir $target_dir
        $VM mkdir $target_dir

        # Copy binaries
        progress "Copying binaries"
        $VM mkdir $target_bindir
        foreach fpat {*.dll *.exe} {
            foreach fn [glob [file join $tcl_root bin $fpat]] {
                $VM copy_to_vm [file nativename $fn] [file nativename [file join $target_bindir [file tail $fn]]]
            }
        }

        # Copy libraries
        progress "Copying libraries"
        $VM mkdir $target_libdir
        foreach dirpat {dde* reg* tcl8* tk*} {
            if {![catch {glob [file join $tcl_root lib $dirpat]} dirs]} {
                foreach dir $dirs {
                    set to_dir [file nativename [file join $target_libdir [file tail $dir]]]
                    progress "Copying $dir to $to_dir"
                    $VM copy_to_vm [file nativename $dir] $to_dir
                }
            }
        }

        # Copy test scripts
        progress "Copying test scripts"
        $VM copy_to_vm [file nativename [config source_test_dir]] [file nativename [file join $target_dir tests]]

    }
}

#

# Reset VM to base state
# Power it on
# Do what setuptarget.tcl does
# Run desired tests


proc main {args} {
    vix::Host create host
    twapi::trap {
        twapi::trap {
            TestTarget create target "C:/virtual machines/vm0-xpsp3/vm0-xpsp3.vmx"
            target setup
        } finally {
            target destroy
        }
    } finally {
        host destroy
    }
}

trap {
    main
} finally {
    vix::finalize
}
