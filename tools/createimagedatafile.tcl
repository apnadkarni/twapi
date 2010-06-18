# Creates a Tcl file containing base64 encoded images

package require base64

set usage {
    usage: createimagedatafile.tcl outfilename arrayname ?imagefile...?
    Creates a Tcl source file 'outfilename' from the specified image files.
    Image file arguments are globbed. If a file appears
    multiple times, it is only included once (this helps with controlling
    order of files by explicitly listing them before globbed args).
    Image data is stored in arrayname indexed by the specified path
    to the file, including the file name itself. If the filespec is
    a directory, recurses to find all files below it.
}
if {[llength $argv] < 2} {puts stderr $usage; exit 1}

# Recursive glob - from the wiki
proc globr {{dir .} {filespec "*"} {types {b c f l p s}}} {
    set files [glob -nocomplain -types $types -dir $dir -- $filespec]
    foreach x [glob -nocomplain -types {d} -dir $dir -- *] {
        set files [concat $files [globr $x $filespec $types]]
    }
    set filelist {}
    foreach x $files {
        while {[string range $x 0 1]=="./"} {
            set x [string range $x 2 end]
        }
        lappend filelist $x
    }
    return $filelist;
}

proc create_imagedata_file {outfilename imagearray files} {
    set outf [open $outfilename w]

    puts $outf "uplevel #0 {"

    foreach file $files {
        set file [string map {\\ /} [string tolower $file]]
        puts $outf "    set ${imagearray}($file) {"
        set fd [open $file]
        fconfigure $fd -translation binary
        puts $outf [base64::encode [read $fd]]
        close $fd
        puts $outf "    }"
    }

    puts $outf "}"

    close $outf
}

set files [list ]
foreach arg [lrange $argv 2 end] {
    set arg [file join $arg];   # Convert any \ to / as required by glob
    set fargs [glob -nocomplain $arg]
    set candidates [list ]
    foreach farg $fargs {
        if {[file type $farg] eq "file"} {
            lappend candidates $farg
        } else {
            set candidates [concat $candidates [globr $farg * f]]
        }
    }
    foreach f $candidates {
        set f [string tolower $f]
        if {[lsearch -exact {.png .gif} [file extension $f]] < 0} {
            continue
        }

        if {[lsearch $files $f] < 0} {
            lappend files $f
        }
    }
}

create_imagedata_file [lindex $argv 0] [lindex $argv 1] $files
