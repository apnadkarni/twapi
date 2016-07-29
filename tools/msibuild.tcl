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
    # Files - list of file paths relative to the Tcl root (no directories)
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
    
    variable tcl_root [file dirname [file dirname [info nameofexecutable]]]

    # Contains a dictionary of all directory paths mapping them to an id
    variable directories {}
}

# Generates a unique id
proc msibuild::id {{path Id}} {
    variable id_counter
    return "[string map {/ _} $path]_[incr id_counter]"
}

# Build a file path list for a MSI package. Returned value is a nested
# list consisting of file paths only (no directories)
proc msibuild::build_file_paths_for_one {msipack} {
    variable msi_packages
    variable tcl_root

    set files {}
    set dirs {}
    dict with msi_packages $msipack {
        # If no Paths dictionary entry, build it based on the package name
        # and version number.
        if {[info exists Paths]} {
            foreach glob $Paths {
                foreach path [glob [file join $tcl_root $glob]] {
                    if {[file isfile $path]} {
                        lappend files $path
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
        lappend files {*}[fileutil::find $dir {file isfile}]
    }
    dict set msi_packages $msipack Files [lmap path [lsort -unique $files] {
        fileutil::relative $tcl_root $path
    }]
}

proc msibuild::add_parent_directory {path} {
    variable directories

    if {[file pathtype $path] ne "relative"} {
        error "Internal error: $path is not a relative path"
    }
    set parent [file dirname $path]
    if {$parent eq "."} {
        return; # Top level
    }
    if {![dict exists $directories $parent]} {
        add_parent_directory $parent
        dict set directories $parent [id $parent]
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
proc msibuild::dump_dirs {} {
    variable directories

    # First build a tree of directories.
    dict for {dir meta} $directories {
        set parent [file dirname $path]
            
    }
}
