package require twapi
package require vix

vix::initialize
twapi::import_commands
        
array set Config {
    target_test_dir c:/twapitest
    virtual_machine_directory {C:/Virtual Machines}
}
set Config(source_test_dir) [file dirname [file normalize [info script]]]
set Config(source_root_dir) [file dirname [file dirname $Config(source_test_dir)]]

proc config {key args} {
    if {[llength $args]} {
        set ::Config($key) [lindex $args 0]
    }
    return $::Config($key)
}

proc vm_path {os platform} {
    dict for {name info} {
        vm0-xpsp3 {
            os xppro
            platform x86
        }
        vm1-win8pro {
            os win8pro
            platform x86
        }
        vm2-w2k12r2 {
            os w2k12r2
            platform x64
        }
        vm3-win81pro {
            os win81pro
            platform x86
        }
    } {
        if {[dict get $info os] eq $os &&
            [dict get $info platform] eq $platform} {
            return [file join [config virtual_machine_directory] $name ${name}.vmx]
        }
    }
    error "No VM found matching OS=$os, platform=$platform"
}


proc tclkits {{ver 8.6} {platform x86}} {
    set kitdir [file join $::Config(source_root_dir) tools tclkits]
    switch -exact -- $ver {
        8.6 { set ver 8.6.3 }
        8.5 { set ver 8.5.17 }
    }
    return [list \
                [file join $kitdir tclkit-cli-${ver}-${platform}.exe] \
                [file join $kitdir tclkit-gui-${ver}-${platform}.exe]]
}

proc usertask {message} {
    # Make sure we are seen
    if {[info commands twapi::set_foreground_window] ne "" &&
        [info commands twapi::get_console_window] ne ""} {
        catch {twapi::set_foreground_window [twapi::get_console_window]}
    }
    puts -nonewline "\n$message\nThen Hit Return to continue..."
    flush stdout
    gets stdin
    return
}

proc nativepath {args} {
    return [file nativename [file join {*}$args]]
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
    method setup {tclver platform twapi_distribution_format} {
        set vm_name [$VM name]

        progress "Powering on $vm_name"
        $VM power_on
        $VM wait_for_tools 60
        usertask "Please login manually at the console to the virtual machine $vm_name\nand wait for the login to complete."
        # Use the GUI dialog so it will work in both tclsh and tkcon
        lassign [credentials_dialog -showsaveoption 0 -target $vm_name -message "Enter the same credentials that you used to login at the console"] user password
        $VM login $user [reveal $password] -interactive 1

        switch -exact -- $twapi_distribution_format {
            "" { set twapidir twapi }
            mod -
            modular { set twapidir twapi-modular }
            bin -
            binary { set twapidir twapi-bin }
            default {
                error "Unknown twapi distro format $twapi_distribution_format"
            }
        }

        set target_dir [nativepath [config target_test_dir]]
        set target_bindir [nativepath $target_dir bin]

        # Create dir structure
        $VM rmdir $target_dir
        $VM mkdir $target_dir
        $VM mkdir $target_bindir
        $VM mkdir [nativepath $target_dir dist]
       #  $VM mkdir [nativepath $target_dir dist $twapidir]

        # Copy binaries
        progress "Copying binaries"
        lassign [tclkits $tclver $platform] tclsh wish
        config target_tclsh [nativepath $target_bindir [file tail $tclsh]]
        config target_wish [nativepath $target_bindir [file tail $wish]]
        $VM copy_to_vm $tclsh [config target_tclsh]
        $VM copy_to_vm $wish [config target_wish]

        # Copy TWAPI. The testutil load_twapi_package expects it in
        # the dist directory
        progress "Copying [nativepath [config source_root_dir] twapi dist $twapidir] to [nativepath $target_dir dist $twapidir]"
        $VM copy_to_vm [nativepath [config source_root_dir] twapi dist $twapidir] [nativepath $target_dir dist $twapidir]
        
        progress "Copying openssl"
        $VM copy_to_vm [nativepath [config source_root_dir] tools openssl] [nativepath $target_dir openssl]

        # Copy test scripts
        progress "Copying test scripts"
        $VM copy_to_vm [nativepath [config source_test_dir]] \
            [nativepath $target_dir tests]

        progress "Registering comtest DLL"
        set script {
            if {[catch {
                set testdir [file normalize [file join [file dirname [info nameofexecutable]] ..]]
                lappend auto_path [file join $testdir dist]
                package require twapi
                set comtest_dll [file nativename [file join $testdir tests comtest comtest.dll]]
                twapi::shell_execute -path regsvr32.exe -verb runas -params $comtest_dll
            } msg]} {
                exit 1
            }
        }
        lassign [$VM script [config target_tclsh] $script -wait 1] pid exit_code elapsed_time
        if {$exit_code} {
            error "Registration of comtest DLL existed with code $exit_code"
        }
        progress "Target setup completed"
    }

    method run {testfile args} {
        parseargs args {
            {constraints.arg {}}
            outfile
        } -setvars -maxleftover 0

        # Note file paths are passed in Unix format using 
        progress "Running test file $testfile"
        set script [format {
            set constraints "%s"
            set testfile "%s"
            if {[catch {
                set testdir [file normalize [file join [file dirname [info nameofexecutable]] ..]]
                lappend auto_path [file join $testdir dist]
                package require twapi
                
                set cmdargs "/C cd [file nativename [file join $testdir tests]] && [file nativename [info nameofexecutable]] $testfile -outfile results.txt"
                if {[llength $constraints]} {
                    append cmdargs " -constraints \"${constraints}\""
                }

                twapi::shell_execute -path $::env(COMSPEC) -verb runas -params $cmdargs
            } msg]} {
                set fd [open c:/twapitest/error.log w]
                puts $fd "$msg\n$::errorInfo"
                close $fd
                exit 1
            }
        } $constraints $testfile]
        lassign [$VM script [config target_tclsh] $script -wait 1] pid exit_code elapsed_time
        progress "Test run completed with exit code $exit_code"
        if {[info exists outfile]} {
            $VM copy_from_vm [nativepath [config target_test_dir] tests results.txt] $outfile
        } else {
            close [file tempname result_file]
            exec notepad.exe [file nativename $result_file] &
            after 1000;         # Give notepad a change to read it
            file delete $result_file
        }
    }
}

#

# Reset VM to base state
# Power it on
# Do what setuptarget.tcl does
# Run desired tests


proc main {args} {
    array set opts [parseargs args {
        {os.arg win8pro}
        {platform.arg x86 {x86 x64}}
        {tclversion.arg 8.6}
        {distribution.arg ""}
        {constraints.arg ""}
        {testfile.arg all.tcl}
    } -maxleftover 0 -setvars]

    vix::Host create host
    try {
        TestTarget create target [vm_path $os $platform]
        try {
            target setup $tclversion $platform $distribution
            target run $testfile -constraints $constraints
        } finally {
            target destroy
        }
    } finally {
        host destroy
    }
}

if {[file tail $argv0] eq [file tail [info script]]} {
    try {
        main {*}$argv
    } finally {
        vix::finalize
    }
}
