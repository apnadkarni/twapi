# This script must be invoked by the Tcl shell from the Tcl distribution
# for which the MSI (Microsoft Installer) package is to be built.

# Make sure we are not picking up anything outside our installation
# even if TCLLIBPATH etc. is set
if {0 && [llength [array names env TCL*]] || [llength [array names env tcl*]]} {
    error "[array names env TCL*] environment variables must not be set"
}

package require fileutil

namespace eval msibuild {
    # Define included packages. A dictionary keyed by the MSI package
    # (not Tcl package) name. The dictionary values are themselves
    # dictionaries with the keys below:
    #
    # TclPackages - list of Tcl packages to include in this MSI package (optional)
    # Description - Description to be shown in the Installer
    # Documentation - link to the documentation (optional)
    # Paths - List of glob paths for  files belonging to the package.
    #   Optional. If not specified, the script will try figuring it out.
    #   Directories will include all files under that directory. Paths must
    #   be relative to the root of the Tcl installation.
    # Mandatory - if 2, the package must be installed. 1 User selectable
    #   defaulting to yes. 0 User selectable defaulting to no.  Optional.
    #   If unspecified, defaults to 1
    #
    # The following additional keys are computed from the above as we
    # go along
    # Files - list of *normalized* file paths
    # Version - version of the package if available
    variable msi_packages

    # This should be read from a config file. Oh well. Later ...
    set msi_packages {
        {Tcl/Tk Core} {
            TclPackages { Tcl Tk }
            Description {Base Tcl/Tk package}
            Paths {bin include lib/tcl8 lib/tcl8.* lib/dde* lib/reg* lib/tk8.* lib/*.lib}
            Mandatory 2
        }
        {Incr Tcl} {
            TclPackages Itcl
            Description {Incr Tcl}
        }
        TDBC {
            TclPackages tdbc
            Description {Tcl Database Connectivity extension}
            Paths {lib/tdbc* lib/sqlite*}
        }
        {Tcl Windows API} {
            TclPackages twapi
            Description {Extension for accessing the Windows API}
        }
        Threads {
            TclPackages Thread
            Description {Extension for script-level access to Tcl threads}
        }
        {Development libraries} {
            Description {C libraries for building your own binary extensions}
            Paths {lib/*.lib}
            Mandatory 0
        }
    }        
    dict unset msi_packages "Development libraries"
    dict unset msi_packages "Tcl Windows API"
    
    # Root of the Tcl installation.
    # Must be normalized else fileutil::relative etc. will not work
    # because they do not handle case differences correctly
    variable tcl_root [file normalize [file dirname [file dirname [info nameofexecutable]]]]

    # Contains a dictionary of all directory paths mapping them to an id
    # Again normalized for same reason as above.
    variable directories {}

    # Used for keeping track of tags in xml generation
    variable xml_tags {}
}

# Generates a unique id
proc msibuild::id {{path {}}} {
    variable id_counter
    if {$path eq ""} {
        return "ID[incr id_counter]"
    } else {
        return "ID[incr id_counter]_[string map {/ _ : _ . _} $path]"
    }
}

# Build a file path list for a MSI package. Returned value is a nested
# list consisting of file paths only (no directories)
proc msibuild::build_file_paths_for_one {msipack} {
    variable msi_packages
    variable tcl_root

    set files {}
    set dirs {}
    dict with msi_packages $msipack {
        if {[info exists Paths]} {
            foreach glob $Paths {
                foreach path [glob [file join $tcl_root $glob]] {
                    if {[file isfile $path]} {
                        lappend files [file normalize $path]
                    } else {
                        lappend dirs $path
                    }
                }                    
            }
        } else {
            # If no Paths dictionary entry, build it based on the package name
            # and version number.
            if {![info exists TclPackages] || [llength TclPackages] == 0} {
                error "No TclPackages or Paths entry for \"$msipack\""
            }
            foreach pack $TclPackages {
                lappend dirs {*}[glob [file join $tcl_root lib ${pack}*]]
            }
        }
    }

    foreach dir $dirs {
        lappend files {*}[lmap path [fileutil::find $dir {file isfile}] {
            file normalize $path
        }]
    }
    dict set msi_packages $msipack Files [lsort -unique $files]
}

proc msibuild::add_parent_directory {path} {
    variable directories
    variable tcl_root
    
    if {[file pathtype $path] ne "absolute"} {
        error "Internal error: $path is not a absolute normalized path"
    }
    set parent [file normalize [file dirname $path]]
    if {$parent eq $tcl_root} {
        return; # Top level
    }
    if {![dict exists $directories $parent]} {
        add_parent_directory $parent
        dict set directories $parent [id [fileutil::relative $tcl_root $parent]]
    }
}

# Builds the file paths for all files contained in all MSI package
proc msibuild::build_file_paths {} {
    variable msi_packages

    dict set directory_tree . Subdirs {}
    foreach msipack [dict keys $msi_packages] {
        build_file_paths_for_one $msipack
        foreach path [dict get $msi_packages $msipack Files] {
            add_parent_directory $path
        }
    }                             
}

# Dumps the list of directories
proc msibuild::dump_dir {dir} {
    variable directories
    variable tcl_root

    if {[file pathtype $dir] ne "absolute"} {
        error "Path \"$dir\" is not an absolute path"
    }

    set indent [string repeat {  } [info level]]
    
    # Note assumes no links

    set dir [file normalize $dir]

    set reldir [fileutil::relative $tcl_root $dir]
    set subdirs [glob -nocomplain -types d -dir $dir -- *]

    # To get indentation right, we have to generate the outer tags
    # before inner tags
    if {$reldir eq "."} {
        set name ProgramFiles
        set id TBD
    } else {
        if {![dict exists $directories $dir]} {
            puts stderr "Directory \"$dir\" not in a package. Skipping..."
            return
        }
        set name [file tail $dir]
        set id [dict get $directories $dir]
    }
    if {[llength $subdirs] == 0} {
        return [tag/ DIRECTORY Name $name Id $id]
    }
    set xml [tag DIRECTORY Name $name Id $id]
    foreach subdir $subdirs {
        append xml [dump_dir $subdir]
    }
    append xml [tag_close]
    return $xml
}

# Buggy XML generator (does not encode special chars)
proc msibuild::tag {tag args} {
    variable xml_tags

    set indent [string repeat {  } [llength $xml_tags]]
    append xml "${indent}<$tag"
    set prefix " "
    dict for {attr val} $args {
        append xml "$prefix$attr='$val'"
        set prefix "\n${indent}[string repeat { }  [string length $tag]]  "
    }
    append xml ">\n"
    lappend xml_tags $tag
    
    return $xml
}

# Like tag but closes it as well
proc msibuild::tag/ {tag args} {
    variable xml_tags

    set indent [string repeat {  } [llength $xml_tags]]
    append xml "${indent}<$tag"
    set prefix " "
    dict for {attr val} $args {
        append xml "$prefix$attr='$val'"
        set prefix "\n${indent}[string repeat { }  [string length $tag]]  "
    }
    append xml "/>\n"
    return $xml
}

# Close last xml tag 
proc msibuild::tag_close {} {
    variable xml_tags
    if {[llength $xml_tags] == 0} {
        error "XML tag stack empty"
    }
    set tag [lindex $xml_tags end]
    set xml_tags [lrange $xml_tags 0 end-1]
    return "[string repeat {  } [llength $xml_tags]]</$tag>\n"
}

proc msibuild::tag_close_all {} {
    variable xml_tags
    set xml ""
    while {[llength $xml_tags]} {
        append xml [close_tag]
    }
    return $xml
}

proc msibuild::build {{chan stdout}} {
    puts $chan {<?xml version="1.0" encoding="UTF-8"?>}

    set xml [tag Wix xmlns http://schemas.microsoft.com/wix/2006/wi]

    # Product - info about Tcl itself
    # Name - Tcl/Tk for Windows
    # Id - "*" -> always generate a new one on every run. Makes upgrades
    #    much easier
    # UpgradeCode - must never change between releases else upgrades won't work.
    # Language, Codepage - Currently only English
    # Version - picked up from Tcl
    # Manufacturer - TCT? Tcl Community?
    append xml [tag Product \
                    Name        "Tcl/Tk for Windows" \
                    Id          * \
                    UpgradeCode "413F733E-BBB8-47C7-AD49-D9E4B039438C" \
                    Language    1033 \
                    Codepage    1252 \
                    Version     [info patchlevel].0 \
                    Manufacturer "Tcl Community"]

    # Package - describes the MSI package itself
    # Compressed - always set to "yes".
    # InstallerVersion - version of Windows Installer required. Not sure
    #     the minimum required here but XP SP3 has 301 (I think)
    append xml [tag/ Package \
                    Compressed       Yes \
                    InstallerVersion 301 \
                    Description      "Installer for Tcl/Tk"]

    append xml [tag_close_all]
    return $xml
}
