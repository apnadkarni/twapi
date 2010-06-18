# Given a directory tree of HTML files, convert them to a Microsoft
# help .CHM file.
#
# Example: (in the h2c directory)
#  tclsh86t ..\html2chm.tcl -name TclWinHelp -title "Tcl for Windows"  -overwrite -tcl -noregexp -hidewarnings -compile

package require Tcl 8.5
package require cmdline
package require fileutil
catch {package require tdom}

namespace eval html2chm {
    # Program options
    variable optlist
    set optlist {
        {noregexp            "Do not use the regexp parser. Abort on tdom failures"}
        {cfile.arg "h2c.cfg" "Name of per-directory configuration file"}
        {compile          "Compile using the HTML Help workshop compiler"}
        {homepage.arg "index.html" "Name to use for the toplevel page for each directory"}
        {name.arg   "" "Name of the help module"}
        {hidewarnings           "Do not show warnings"}
        {output.arg "" "Base name of output files"}
        {overwrite        "Overwrite existing files as needed"}
        {tcl              "Enable Tcl documentation extensions"}
        {title.arg "Help" "Main title to use for the help file"}
        {verbose          "Verbose progress messages"}
    }
    variable opts

    # Meta information collected about files and documentation nodes
    # Array indexed by file/dir id (full path), containing a dict
    # with meta information for the file
    variable meta

    # Keeps a full index across all files
    # List of sublists containing the index text, the target link and
    # the path where it occurs (used to disambiguate between entries
    # from different paths.
    variable fullindex {}

    # Dummy root for Table of Contents
    variable tocroot
}

namespace eval html2chm::app {
    # Procs in this namespace can be redefined by the application

    proc progress {str} {
        puts stderr $str
    }

    proc warn {str} {
        if {! [set [namespace parent [namespace current]]::opts(hidewarnings)]} {
            puts stderr "Warning: $str"
        }
    }

    proc usage {msg} {
        if {$msg ne ""} {
            puts stderr $msg
        }
        puts stderr "Usage:"
        puts stderr "[info nameofexecutable] html2chm.tcl ?options? ?DIRECTORY?"
        puts stderr [cmdline::usage [set [namespace parent]::optlist]]
        exit 1
    }

}

proc html2chm::tocheaders_alias {dir args} {
    # Configuration callback - defines index pages
    variable meta

    if {[llength $args] == 0} {
        set args {1 2};         # Default - top 3 header levels
    }
    array set headers {};         # To remove duplicates, indexed in XPATH form
    foreach arg $args {
        if {[regexp {^([1-9])(-[1-9])?$} dontcare low high]} {
            set headers(//h$low) ""
            if {$high ne ""} {
                while {[incr low] < $high} {
                    set headers(//h$low) ""
                }
            }
        } else {
            error "Invalid tocheaders specification $arg in configuration for directory $dir"
        }
    }

    # Store the XPATH expressions. Note the linkage command may also append
    # to this field.
    dict lappend meta($dir) toc_xpaths {*}[array names headers]
}

proc html2chm::toc_alias {dir toc} {
    variable meta

    # $toc is pairs of filename title, one per line
    set lines [split $toc \n]
    # Convert the flat list into a nested list so we can sort easier.
    # Note Tcl8.5 lsort does not have the -stride option, but does have -index
    foreach line $lines {
        if {[catch {lrange $line 0 end} toc_pair] ||
            [llength $toc_pair] > 2} {
            # Not well formed list or too many elements
            error "Badly formatted toc line '$line' in configuration for directory $dir."
        }
        if {[llength $toc_pair] == 0} continue; # Skip blank line
        foreach {fn title} $toc_pair break
        dict lappend meta($dir) toc [list [file join $dir $fn] $title]
    }    
}

proc html2chm::homepage_alias {dir args} {
    # Configuration callback - defines home page for a directory
    variable meta

    switch -exact -- [lindex $args 0] {
        generate {
            error "homepage generate command not implemented"
        }
        file {
            # For file - name of file that should be the home file
            # For generate - a file of this name will be generated
            if {[llength $args] != 2 } {
                error "Incorrect number of arguments for 'homepage' command in configuration for directory $dir."
            }
            dict set meta($dir) homepage [lindex $args 1]
        }
        default {
            error "Unknown home page command type [lindex $args 0]."
        }
    }
}

proc html2chm::linkage_alias {dir defs} {
    # Configuration callback - defines links from content and indexes
    # Note there can be more than one link command

    variable meta
    # Definitions are of the form (one per line)
    #   xpath ?toc? ?index?
    set lines [split $defs \n]
    foreach line $lines {
        if {[catch {lrange $line 0 end} def]} {
            error "Badly formatted linkage line '$line' in configuration for directory $dir."
        }
        if {[llength $def] == 0} continue; # Skip blank line
        set xpath [lindex $def 0]
        set linkages [lrange $def 1 end]
        if {[llength $linkages] == 0} {
            set linkages {toc index}; # default
        }
        foreach linkage $linkages {
            switch -exact -- $linkage {
                toc   -
                index { set key ${linkage}_xpaths }
                default {
                    error "Invalid option '$opt' for 'link' command in configuration for directory $dir."
                }
            }
            dict lappend meta($dir) $key $xpath
        }
    }
}

proc html2chm::read_dir_config {dir} {
    # Parses configuration files and and sets per directory options
    variable opts

    set path [file join $dir $opts(cfile)]
    if {![file exists $path]} {
        return
    }

    # Create a safe interpreter to protect against arbitrary (malicious)
    # input in config files. Creating/destroying safe interpreters is
    # cheap enough so we recreate on every call. Keeping it around would
    # require state from previous invocation to be cleaned up which is
    # a pain.

    set cinterp [interp create -safe]
    $cinterp alias toc [namespace current]::toc_alias $dir
    $cinterp alias homepage [namespace current]::homepage_alias $dir
    $cinterp alias linkage [namespace current]::linkage_alias $dir
    $cinterp alias tocheaders [namespace current]::tocheaders_alias $dir

    if {[catch {
        $cinterp invokehidden -namespace settings source $path
    } msg]} {
        interp delete $cinterp
        error $msg
    }

    interp delete $cinterp
}

proc html2chm::entity_decode s {
    # TBD
    return $s
}

proc html2chm::make_hhp_path {path} {
    variable tocroot
    # Convert paths to a format that the MS Help compiler likes
    #    return [file nativename [file attributes $path -shortname]]
    return [fileutil::stripPath $tocroot [file nativename $path]]
}

proc html2chm::read_file {path} {
    set fd [open $path r]
    set data [read $fd]
    close $fd
    return $data
}

proc html2chm::process_content {path root} {
    # Process a file containing content. HTML files
    # are parsed to retrive titles and other meta data.
    # Other files are let alone but marked for
    # storage into the help file.

    variable meta
    variable tocroot
    variable fullindex
    variable opts

    if {$opts(verbose)} {
        app::progress "Processing $path..."
    }

    if {$path eq $root} {
        set parent $tocroot
    } else {
        set parent [file dirname $path]
    }

    set ext [file extension $path] 
    if {[string compare -nocase $ext ".html"] &&
        [string compare -nocase $ext ".htm"]} {
        # Note we do not add these files to the parent
        # child list as they will not show up in ToC
        # Also, we skip help files assuming they are leftover
        if {[string tolower $ext] ni {.chm .hhp .hhk .hhc}} {
            set meta($path) [dict create type other]
        }
        return
    }
    
    set meta($path) [dict create type content parent $parent]

    # Parse using tdom to get its content. On failure
    # we'll be crude and use a regexp
    set title [file rootname [file tail $path]]; # Default title
    set data [read_file $path]
    set ctr 0;                  # Just to generate unique targets
    if {[catch {
        # Note the -html flag is important, not just for being less
        # strict but also for the selectNodes later to work without
        # qualifiers (what the heck was I talking about here? - TBD)
        set doc [dom parse -keepEmpties -html $data]
        set node [lindex [$doc selectNodes /html/head/title] 0]
        if {$node ne ""} {
            set title [$node asText]
        }
        # Check if any rules are defined for ToC and indexes
        foreach xref {toc index} {
            if {[dict exists $meta($parent) ${xref}_xpaths]} {
                set nodes [$doc selectNodes [join [dict get $meta($parent) ${xref}_xpaths] "|"]]
                set warned false
                foreach node $nodes {
                    # Add each link as a ToC/index item
                    set text [$node asText]
                    # We do not want to create entries for strings like
                    # (1), a) etc. as are present in Tcl man pages
                    if {[regexp {\)\s*$} $text]} {
                        continue
                    }
                    set target ${path}
                    if {[$node hasAttribute "name"]} {
                        append target "#[$node getAttribute name]"
                    } else {
                        # TBD - try if parent or any child has a name attribute
                        if {! $warned} {
                            app::warn "$path: One or more xpath expressions for $xref do not have name attribute. Linking to page."
                            set warned true
                        }
                        # MS Help gets confused if multiple targets in ToC
                        # are the same. It uses the same name for all the
                        # ToC entries. So make them unique by adding dummy
                        # target names. Since they do not exist (probably)
                        # target will just be top of the page.
                        append target "#[incr ctr]"
                    }
                    if {$xref eq "toc"} {
                        dict lappend meta($path) $xref \
                            [list $text $target]
                    } else {
                        # There is only one global index, not per path
                        lappend fullindex [list $text $target $path]
                    }
                }
            }
        }

        # TBD - add other rules here for indices etc.
        $doc delete
    } msg]} {
        if {$opts(noregexp)} {
            app::progress "tdom error ([string range $msg 0 60]...). Aborting."
            exit 1
        }
        app::progress "tdom error ([string range $msg 0 60]...). Using regexp to parse file"

        # Clean up
        if {[info exists doc]} {
            catch {$doc delete}
        }

        # Try simplistic regexp
        if {[regexp {<title>(.*)</title>} $data dontcare title]} {
            set title [entity_decode $title]
        }
    }

    dict set meta($path) title $title
    # Also add title to the index
    lappend fullindex [list $title $path $path]

    dict lappend meta($parent) children [list $path $title]
}

proc html2chm::process_directory {dir root} {

    # Given a directory, reads meta information from it
    # for purposes of creating a ToC etc.
    # dir - path to the directory
    # root - the toplevel root to which this belongs
    # Both $dir and $root must be normalized

    variable opts
    variable meta
    variable tocroot

    read_dir_config $dir;       # Per-dir configuration

    if {$dir eq $root} {
        set parent $tocroot
    } else {
        set parent [file dirname $dir]
    }

    dict lappend meta($parent) children [list $dir [file rootname [file tail $dir]]]

    dict set meta($dir) type dir
    dict set meta($dir) parent $parent
    dict set meta($dir) title [file rootname [file tail $dir]]

    # Inherit configuration (some of it) from parent unless we already
    # read it from the directory.
    foreach metakey {toc_xpaths index_xpaths homepage} {
        if {![dict exists $meta($dir) $metakey] &&
            [dict exists $meta($parent) $metakey]} {
            # Inherit from parent
            dict set meta($dir) $metakey [dict get $meta($parent) $metakey]
        }
    }

    if {$opts(homepage) ne "" &&
        (![dict exists $meta($dir) homepage]) &&
        [file exists [file join $dir $opts(homepage)]]} {
        dict set meta($dir) homepage $opts(homepage)
    }    
}

proc ::html2chm::write_hhp_object {fd title args} {
    # Write a single ToC or index link as a list item
    # If $args specified, a link is written to
    # the first element of $args. Else it is
    # written as a ToC node with no content link.

    # TAG IS NOT ON THE SAME LINE AS THE LI TAG. Yes, I know, hard to
    # believe. But confirmed by hand editing the file. Not sure
    # exactly when this happens or if it is related to other structure,
    # but for now do not start <object> tag on a new line.
    set title [string map {\n " " \t " "} $title]
    puts $fd "<object type=\"text/sitemap\">
    <param name=\"Name\" value=\"$title\">
"
    if {[llength $args]} {
        puts $fd "\
    <param name=\"Local\" value=\"[make_hhp_path [lindex $args 0]]\">
"
    }
    puts $fd "\
  </object>
"
}

proc ::html2chm::write_toc_entry {fd path {title ""}} {
    # Writes the ToC entries for the ToC subtree under $path
    variable tocroot
    variable meta
    variable opts

    # Write the entry for this element first, and then its children
    if {$path ne $tocroot} {
        if {$title eq ""} {
            set title [dict get $meta($path) title]
        }

        # NOTE: UNDER SOME CIRCUMSTANCES, THE HTML HELP BARFS IF THE OBJECT
        # TAG IS NOT ON THE SAME LINE AS THE LI TAG. Yes, I know, hard to
        # believe. But confirmed by hand editing the file. Not sure
        # exactly when this happens or if it is related to other structure,
        # but for now always put object on the same line as li.
        puts -nonewline $fd "<li> "
        if {[dict get $meta($path) type] ne "dir"} {
            # For now, set the default page for the help file to be
            # the first page we encounter in any directory
            # TBD - fix this once, we generate index page for directories
            if {![info exists opts(defaultpage)]} {
                set opts(defaultpage) $path
            }
            write_hhp_object $fd $title $path
        } else {
            if {[dict exists $meta($path) homepage]} {
                write_hhp_object $fd $title \
                    [make_hhp_path [file join $path [dict get $meta($path) homepage]]]
            } else {
                write_hhp_object $fd $title
            }
        }
    }

    # Write the children if any. If there is a toc field defined,
    # that contains the order and titles of the children. Otherwise
    # sort alphabetically based on titles.
    if {[dict exists $meta($path) children]} {
        if {[dict exists $meta($path) toc]} {
            # toc field overrides children field
            set children [dict get $meta($path) toc]
            set default_toc false
        } else {
            # Sort the kids based on the second element (title)
            set children [lsort -dictionary -index 1 [dict get $meta($path) children]]
            set default_toc true
        }
        puts $fd "<ul>";        # Children are a sublist
        foreach child $children {
            # If we are generating the default ToC, then exclude a
            # file if it is the default home page for the directory
            # (since it is linked to the directory node)
            if {$default_toc && 
                [dict exists $meta($path) homepage] &&
                [string equal -nocase \
                     [file tail [lindex $child 0]] \
                     [file tail [dict get $meta($path) homepage]]]} {
                # Skip the toc for this as it is linked to the dir node
                continue
            }
            write_toc_entry $fd {*}$child
        }
        puts $fd "</ul>"
    } else {
        # No kids, but there may still be a ToC for link targets
        if {[dict exists $meta($path) toc]} {
            puts $fd "<ul>";        # Children are a sublist
            foreach toc_entry [dict get $meta($path) toc] {
                # See comments above for why we need -nonewline
                puts -nonewline $fd "<li> "
                write_hhp_object $fd {*}$toc_entry
                puts -nonewline $fd "</li> "
            }
            puts $fd "</ul>"
        }
    }

    # Terminate this list item
    if {$path ne $tocroot} {
        puts $fd "</li>"
    }
}

proc ::html2chm::write_toc {hhc_path} {
    # Write out the table of contents file
    variable meta
    variable opts
    variable tocroot

    app::progress "Writing table of contents file [fileutil::stripPath [pwd] $hhc_path]"

    set fd [open $hhc_path w]
    puts $fd {
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<meta name="GENERATOR" content="Microsoft&reg; HTML Help Workshop 4.1">
<!-- Sitemap 1.0 -->
</head>
<body>
<object type="text/site properties">
        <param name="Auto Generated" value="Yes">
</object>
}
    
    # Write out the actual table of contents
    write_toc_entry $fd $tocroot

    puts $fd {
</body>
</html>
    }

    close $fd
}

proc ::html2chm::write_index {hhk_path} {
    # Write out the table of index file
    variable meta
    variable opts
    variable fullindex

    app::progress "Writing index file [fileutil::stripPath [pwd] $hhk_path]"

    set fd [open $hhk_path w]
    puts $fd {
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<meta name="GENERATOR" content="Microsoft&reg; HTML Help Workshop 4.1">
<!-- Sitemap 1.0 -->
</head>
<body>
<ul>
}
    
    # Write out the index
    set indexlist [list]
    foreach entry $fullindex {
        foreach {phrase path origin} $entry break
        # If Tcl extensions are enabled, then strip namespace qualifiers
        # and also add entries for tail of namespace
        if {$opts(tcl) &&
            [regexp {^\s*([^\s]+)($|\s+.*$)} $phrase _ firstword rest]} {
            # See if it looks like a namespace prefixed command.
            # If so, remove leading ::, if any and also 
            # add entries without the namespace
            if {[string first :: $firstword] >= 0} {
                # Looks like a namespace,
                lappend indexlist "[string trimleft $firstword :]$rest" $path $origin
                # Add an entry for the command without namespaces
                lappend indexlist [namespace tail $firstword] $path $origin
            } else {
                # Keep original entry
                lappend indexlist $phrase $path $origin
            }
        } else {
            # Not Tcl, keep the original
            lappend indexlist $phrase $path $origin
        }
    }

    # If some entries occur more than once then we need to disambiguate them
    # as MS Help does not do that for us making it hard for users to
    # distinguish between them.
    array set counts {}
    foreach {phrase path origin} $indexlist {
        if {[info exists counts($phrase)]} {
            incr counts($phrase)
        } else {
            set counts($phrase) 1
        }
    }

    # Now actually write out the stuff
    foreach {phrase path origin} $indexlist {
        # Disambiguate if necessary!
        if {$counts($phrase) > 1} {
            append phrase " ([file rootname [file tail $origin]])"
        }
        # See comments above for why we need -nonewline
        puts -nonewline $fd "<li> "
        write_hhp_object $fd $phrase $path
        puts $fd "</li> "
    }

    puts $fd {
</ul>
</body>
</html>
    }

    close $fd
}

proc ::html2chm::write_project {hhppath} {
    # Writes out the Help project file
    variable opts
    variable meta
    variable tocroot

    app::progress "Writing project file [fileutil::stripPath [pwd] $hhppath]"

    set basename [file tail [file rootname $hhppath]]
    set chmfile ${basename}.chm; # Help file
    set hhcfile ${basename}.hhc; # Contents file
    set hhkfile ${basename}.hhk; # Index file

    set fd [open $hhppath w]

    # The Help compiler gets confused if the default page
    # is below the directory containing the project file. In
    # that case, use a relative file name.
    set hhpdir [file normalize [file dirname $hhppath]]
    set defpagedir [file normalize [file dirname $opts(defaultpage)]]
    set hhpdir_len [string len $hhpdir]
    if {[string equal -nocase $hhpdir [string range $defpagedir 0 [expr {$hhpdir_len-1}]]]} {
        set defpage [string range $opts(defaultpage) [incr hhpdir_len] end]
    } else {
        set defpage $opts(defaultpage)
    }

    puts $fd "
\[OPTIONS]
Compatibility=1.1 or later
Display compile progress=No
Full-text search=Yes
Binary TOC=Yes
Language=0x409 English (United States)
Compiled file=$chmfile
Contents file=$hhcfile
Index file=$hhkfile
Title=$opts(title)
Default topic=[make_hhp_path $defpage]
Default Window=$opts(title)
\[WINDOWS]
$opts(title)=\"$opts(title)\",\"$hhcfile\",\"$hhkfile\",\"[make_hhp_path $defpage]\",,,,,,0x62521,,0x380e,,,,,,,,0

\[FILES]
"

    foreach path [array names meta] {
        if {$path ne $tocroot &&
            [dict get $meta($path) type] ne "dir"} {
            puts $fd [make_hhp_path $path]
        }
    }
    puts $fd "\n\[INFOTYPES]"
    close $fd
}

proc html2chm::compile_project {path} {
    # Compile a HTML help project file

    variable opts

    # For some reason, the help compiler crashes occasionally when
    # exec'ed. But works fine from the command line. So for now by default
    # just ask user to do it.

    if {! $opts(compile)} {
        app::progress "\nPlease compile file \"$path\" with the Microsoft HTML Help workshop."
        return
    }

    app::progress "Compiling help project [fileutil::stripPath [pwd] $path]"

    set exe [auto_execok hhc.exe]
    if {$exe eq ""} {
        # Not in path. See if we can find it in Program files
        if {[info exists ::env(PROGRAMFILES)]} {
            set dir [file join $::env(PROGRAMFILES) "HTML Help Workshop"]
            if {[file isdirectory $dir]} {
                set exe [file join $dir hhc.exe]
            }
        }
    }

    if {$exe eq ""} {
        error "Could not find help compiler."
    }

    # hhc always returns error even on successful compiles!
    if {[catch {
#        exec [file attributes $exe -shortname] [file nativename [file attributes $path -shortname]]
        exec [file attributes $exe -shortname] [file nativename $path]
    } msg]} {
        if {[lindex $::errorCode 0] eq "CHILDSTATUS" &&
            [lindex $::errorCode 2] eq "1"} {
            # Not really an error.
            # Chop off last line (child process exit with error)
            set msg [join [lrange [split $msg \n] 0 end-1] \n]
            app::progress $msg
        } else {
            error $msg
        }
    }
}

proc html2chm::main {args} {
    variable opts
    variable optlist
    variable meta
    variable tocroot

    array set opts [::cmdline::getoptions args $optlist]

    # NOTE:
    # The original code would accept multiple arguments pointing to
    # directories anywhere on the file system. However, the Help Workshop
    # has problems handling links in the table of contents in such
    # a case. In particular, files outside the hhp file directory
    # are stored at the top level (without the path) resulting in broken
    # links and worse, potential file name clashes (eg. index.html)
    # So we restrict arguments to a single directory under which
    # all files must lie.
    # TBD - maybe we could copy files to a temporary build tree
    # here instead of leaving it to the user?
    if {[llength $args] > 1} {
        app::usage "Invalid number of arguments, must specify a single input directory."
    }
    if {[llength $args] == 0} {
        set topdir [file normalize [pwd]]
    } else {
        set topdir [file normalize [lindex $args 0]]
    }
        
    if {![file isdirectory $topdir]} {
        error "$topdir does not exist or is not a directory."
    }

    set tocroot $topdir
    read_dir_config $topdir; # "Defaults" for the tree

    # Put in a list to match old code. That way we can switch back
    # easily enough later.
    set specs [list [file join $topdir *]]

    if {$opts(name) eq ""} {
        set opts(name) [file rootname [file tail [lindex $specs 0]]]
    }
    if {$opts(output) eq ""} {
        set opts(output) $opts(name)
    }
    set output_path_base [file join $topdir $opts(output)]


    # Write the table of contents and project files
    if {! $opts(overwrite)} {
        foreach ext {hhp hhc hhk chm} {
            if {[file exists "${output_path_base}.$ext"]} {
                error "File ${output_path_base}.$ext exists. Use -overwrite option to overwrite."
            }
        }
    }

    # glob treats \ specially, so convert all paths to use /
    set paths {}
    foreach spec $specs {
        lappend paths [file join $spec]
    }
    set paths [glob -nocomplain {*}$paths]
    if {[llength $paths] == 0} {
        app::warn "No matching files."
        exit
    }

    # Note non-existent files are silently ignored
    app::progress "Reading input files"
    foreach spec $paths {
        set spec [file normalize $spec]
        if {[file isdirectory $spec]} {
            process_directory $spec $spec
            foreach path [::fileutil::find $spec] {
                set path [file normalize $path]
                if {[file isdirectory $path]} {
                    process_directory $path $spec
                } else {
                    process_content $path $spec
                }
            }
        } else {
            # Do not include the per-dir configuration files
            if {[string compare -nocase $opts(cfile) [file tail $spec]]} {
                process_content $spec $spec
            }
        }
    }


    # Exclude help project files from help file
    foreach ext {hhc hhp hhk chm} {
        unset -nocomplain meta(${output_path_base}.${ext})
    }

    # Write it out
    write_toc ${output_path_base}.hhc
    write_index ${output_path_base}.hhk
    write_project ${output_path_base}.hhp
    compile_project ${output_path_base}.hhp
}

if {[string equal -nocase [info script] $::argv0]} {
    if {[catch {html2chm::main {*}$::argv} msg]} {
        puts stderr "Error: $msg"
    }
}
