# Adapted from http://wiki.tcl.tk/19801

# foreachLine -- COPIED FROM tcllib's fileutil package (we don't want
#    to be dependent on any packages being installed)
#
#	Executes a script for every line in a file.
#
# Arguments:
#	var		name of the variable to contain the lines
#	filename	name of the file to read.
#	cmd		The script to execute.
#
# Results:
#	None.

proc foreachLine {var filename cmd} {
    upvar 1 $var line
    set fp [open $filename r]

    # -future- Use try/eval from tcllib/control
    catch {
	set code 0
	set result {}
	while {[gets $fp line] >= 0} {
	    set code [catch {uplevel 1 $cmd} result]
	    if {($code != 0) && ($code != 4)} {break}
	}
    }
    close $fp

    if {($code == 0) || ($code == 3) || ($code == 4)} {
        return $result
    }
    if {$code == 1} {
        global errorCode errorInfo
        return \
		-code      $code      \
		-errorcode $errorCode \
		-errorinfo $errorInfo \
		$result
    }
    return -code $code $result
}

set usage {
    usage: createtmfile ?-compact? ?-outfile OUTPUTFILE? package version ?tclfile...? ?dllfile?
    Creates a Tcl module file from the specified tclfiles
    and/or maximally one DLL. Arguments are globbed. If a file appears
    multiple times, it is only included once (this helps with controlling
    order of files by explicitly listing them before globbed args)
}

array set opts {
    -compact 0
    -force 0
}
while {[string index [lindex $argv 0] 0] eq "-"} {
    set opt [lindex $argv 0]
    set argv [lrange $argv 1 end]
    switch -exact -- $opt {
        -compact { set opts(-compact) 1 }
        -outfile {
            set opts(-outfile) [lindex $argv 0]
            set argv [lrange $argv 1 end]
        }
        -force { set opts(-force) 1 }
        default {
            error "Unknown option '$opt'."
        }
    }
}

if {[llength $argv] == 0} {puts stderr $usage; exit 1}

proc main {package version header files} {
    global opts

    if {[info exists opts(-outfile)] && $opts(-outfile) ne ""} {
        if {[file exists $opts(-outfile)] && ! $opts(-force)} {
            error "Output file $opts(-outfile) exists. Use -force to overwrite."
        }
        set outf [open $opts(-outfile) w]
    } else {
        set outf [open ${package}-${version}.tm w]
    }
    fconfigure $outf -translation lf

    if {$header ne ""} {
        set f [open $header]
        fcopy $f $outf
        close $f
    }

    # This proc has to be at the beginning of the file since the app
    # Tcl code may call it at any time
    puts $outf "proc copy_dll_from_tm {{path {}}} {"
    puts $outf "if {\$path eq {}} {set path \[file join \$env(TMP) ${package}-${version}.dll \]}"
    puts $outf "set tmp \[open \$path w\]"
    puts $outf {
        set f [open [info script]]
        fconfigure $f -translation binary
        set data [read $f][close $f]
        set ctrlz [string first \u001A $data]
        fconfigure $tmp -translation binary
        puts -nonewline $tmp [string range $data [incr ctrlz] end]
        close $tmp
    }
    puts $outf "}"

    # Commented out package provide - let package itself do this
    #puts $outf "package provide [lindex $argv 0] [lindex $argv 1]"


    foreach a $files {
        switch -- [file extension $a] {
            .dll {
                set f [open $a]
                fconfigure $f    -translation binary
                puts $outf "\#-- from [file tail $a]"
                puts -nonewline $outf \u001A
                fconfigure $outf -translation binary
                fcopy $f $outf
                close $f
                break
            }
            .tcl {
                puts $outf "\#-- from [file tail $a]"
                set f [open $a]
                if {$opts(-compact)} {
                    set outdata ""
                    foreachLine line $a {
                        # Skip pure comments
                        if {[regexp {^\s*#} $line]} continue
                        # SKip blank lines
                        if {[regexp {^\s*$} $line]} continue
                        # Compact leading spaces
                        regsub {^(\s+)} $line "" line
                        append outdata "${line}\n"
                    }
                    puts -nonewline $outf $outdata
                } else {
                    fcopy $f $outf
                }
                close $f
            }
            default {
                set f [open $a]
                fcopy $f $outf
                close $f
            }
        }
    }
    close $outf
}

set package [lindex $argv 0]
set version [lindex  $argv 1]
set header  [lindex $argv 2]
set files [list ]
foreach arg [lrange $argv 3 end] {
    foreach f [glob -nocomplain [file normalize $arg]] {
        set f [string tolower $f]
        if {[lsearch $files $f] < 0} {
            lappend files $f
        }
    }
}

main $package $version $header $files
