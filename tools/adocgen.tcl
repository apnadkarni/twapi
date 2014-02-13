package require textutil
package require cmdline

namespace eval docgen {
    variable options
    array set options {}
}

proc docgen::docgen {docdef} {
    variable options

    if {$options(unsafe)} {
        interp create -- slave
    } else {
        interp create -safe -- slave
    }

    set output ""
    while {[llength $docdef]} {
        set docdef [lassign $docdef command]
        switch -exact -- $command {
            shell -
            script -
            text {
                append output [$command [getarg $command docdef]]
            }
        }
    }

    interp delete slave
    return $output
}

# Parses callouts from a script. Returns list of 2 elems -
# script text to be displayed (will not be valid Tcl)
# and list of callouts
proc docgen::callouts {script} {
    # Superfluous \'s are for emacs indentation workarounds
    #regsub -line -all {\;\s*\#(\s*<\d+>).*$} $script {\1} display
    regsub -line -all {\;(\s*)\#(\s*<\d+>).*$} $script { \1 \2} display
    return [list $display [regexp -line -all -inline {<\d+>.*$} $script]]
}

# Returns text verbatim except left aligned to shortest common blank prefix
proc docgen::text {text} {
    return [textutil::undent $text]
}

# Returns script marked for output, verifying it runs without errors
proc docgen::script {script} {
    variable options
    set script [textutil::undent [string trim $script \n]]
    lassign [callouts $script] display callouts
    if {[catch {
        slave eval $script
    } msg]} {
        error "Error in example script \"[string range $script 0 99]...\" ($msg)"
    }
    return "$options(highlight)----\n$display\n----\n[join $callouts \n]\n"
}

# Returns script as though run interactively
proc docgen::shell {script} {
    variable options
    set cmd ""
    set output ""
    set callouts {}
    foreach line [split [textutil::undent [string trim $script \n]] \n] {
        append cmd $line
        if {[info complete $cmd]} {
            lassign [callouts $cmd] display callouts2
            lappend callouts {*}$callouts2
            if {[catch {
                interp eval slave $cmd
            } result]} {
                error "Error in example script \"[string range $cmd 0 99]...\" ($result)"
            }
            append output "$options(prompt)$display\n$result\n"
            set cmd ""
        }
    }
    if {[string length $cmd]} {
        error "Incomplete line \"$cmd\""
    }
    # TBD - Is there a "console" block type instead of [source] ?
    return "$options(highlight)----\n[string trim $output]\n----\n[join $callouts \n]\n"
}

proc docgen::getarg {command argsvar} {
    upvar 1 $argsvar var

    if {[llength $var] == 0} {
        error "No value specified for command $command"
    }

    set var [lassign $var arg]
    return $arg
}


proc usage {} {
    return "Usage: [file tail [info nameofexecutable]] $::argv0 ?options? INPUTFILE1 ?INPUTFILE2...?"
}

proc exit1 {msg} {
    if {$msg ne ""} {
        puts stderr $msg
    }
    exit 1
}

proc docgen::process_files files {
    variable options

    if {$options(highlight) ne ""} {
        set options(highlight) "\[source,$options(highlight)\]\n"
    }
    if {[llength $files] == 0} {
        # Treat as error since empty list may be because no files
        # match due to mistyping
        error "No matching files"
    }

    foreach fn $files {
        if {$options(outdir) eq ""} {
            set outfn [file rootname $fn].ad
        } else {
            file mkdir $options(outdir); # Make sure it exists
            set outfn [file join $options(outdir) [file rootname [file tail $fn]].ad]
        }
        if {! $options(overwrite) && [file exists $outfn]} {
            error "File $outfn already exists"
        }

        set fd [open $fn]
        set content [read $fd]
        close $fd

        set fd [open $outfn w]
        puts $fd [docgen::docgen $content]
        close $fd
    }
}

if {[catch {
    array set docgen::options [cmdline::getoptions argv {
        {outdir.arg  ""  "Directory for output files"}
        {overwrite       "Force - overwrite any existing files"}
        {unsafe          "Unsafe mode - do not run in a safe interpreter"}
        {prompt.arg "% " "Prompt to show in shell commands"}
        {highlight.arg "" "Source highlighter to use"}
    } [usage]]

    docgen::process_files [cmdline::getfiles $argv 1]

} msg]} {
    exit1 $msg
}

exit 0
