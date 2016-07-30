# This script must be invoked by the Tcl shell from the Tcl distribution
# for which the MSI (Microsoft Installer) package is to be built.

# Make sure we are not picking up anything outside our installation
# even if TCLLIBPATH etc. is set
if {0 && [llength [array names env TCL*]] || [llength [array names env tcl*]]} {
    error "[array names env TCL*] environment variables must not be set"
}

package require fileutil

namespace eval msibuild {
    # Define included features. A dictionary keyed by the MSI feature Id
    # The dictionary values are themselves
    # dictionaries with the keys below:
    #
    # TclPackages - list of Tcl packages to include in this MSI package (optional)
    # Name - The Title to be shown in the MSI feature tree
    # Description - Description to be shown in the Installer (optional)
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
    variable feature_definitions

    # This should be read from a config file. Oh well. Later ...
    set feature_definitions {
        core {
            Name {Tcl/Tk Core}
            TclPackages { Tcl Tk }
            Description {Includes Tcl and Tk core, Windows registry and DDE extensions.}
            Paths {bin include lib/tcl8 lib/tcl8.* lib/dde* lib/reg* lib/tk8.* lib/*.lib}
            Mandatory 2
        }
        itcl {
            Name {Incr Tcl}
            TclPackages Itcl
            Description {Incr Tcl object oriented extension}
        }
        tdbc {
            Name TDBC
            TclPackages tdbc
            Description {Tcl Database Connectivity extension}
            Paths {lib/tdbc* lib/sqlite*}
        }
        twapi {
            Name {Tcl Windows API}
            TclPackages twapi
            Description {Extension for accessing the Windows API}
        }
        threads {
            Name Threads
            TclPackages Thread
            Description {Extension for script-level access to Tcl threads}
        }
        clibs {
            {C libraries} 
            Description {C libraries for building your own binary extensions}
            Paths {lib/*.lib}
            Mandatory 0
        }
    }        

    # selected_features contains the actual selected features from the above.
    variable selected_features
    
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
    variable tcl_root
    if {$path eq ""} {
        return "ID[incr id_counter]"
    } else {
        set path [fileutil::relative $tcl_root $path]
        return "ID[incr id_counter]_[string map {/ _ : _ - _ + _} $path]"
    }
}

# Build a file path list for a MSI package. Returned value is a nested
# list consisting of file paths only (no directories)
proc msibuild::build_file_paths_for_feature {feature} {
    variable selected_features
    variable tcl_root

    set files {}
    set dirs {}
    dict with selected_features $feature {
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
                error "No TclPackages or Paths entry for \"$feature\""
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
    dict set selected_features $feature Files [lsort -unique $files]
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
        dict set directories $parent [id $parent]
    }
}

# Builds the file paths for all files contained in all MSI package
proc msibuild::build_file_paths {} {
    variable selected_features

    dict set directory_tree . Subdirs {}
    foreach feature [dict keys $selected_features] {
        build_file_paths_for_feature $feature
        foreach path [dict get $selected_features $feature Files] {
            add_parent_directory $path
        }
    }                             
}

# Generate the Directory nodes
proc msibuild::generate_directory {dir} {
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
        # Top level must have this exact values
        set name SourceDir
        set id   TARGETDIR
    } else {
        if {![dict exists $directories $dir]} {
            puts stderr "Directory \"$dir\" not in a package. Skipping..."
            return
        }
        set name [file tail $dir]
        set id [dict get $directories $dir]
    }
    if {[llength $subdirs] == 0} {
        return [tag/ Directory Name $name Id $id]
    }
    set xml [tag Directory Name $name Id $id]
    foreach subdir $subdirs {
        append xml [generate_directory $subdir]
    }
    append xml [tag_close]
    return $xml
}

proc msibuild::generate_file {path} {
    variable directories

    if {[file pathtype $path] ne "absolute"} {
        error "generate_file passed a non-absolute path"
    }
    set dir [file dirname $path]
    if {![dict exists $directories $dir]} {
        error "Could not find directory \"$dir\" in directories dictionary"
    }
    set dir_id [dict get $directories $dir]
    set file_id [id $path]

    # Every FILE must be enclosed in a Component and a Component should
    # have only one file.
    set xml [tag Component \
                 Id CMP_$file_id \
                 Guid * \
                 Directory $dir_id]
    append xml [tag/ File \
                    Id $file_id \
                    Source $path \
                    KeyPath yes]
    
    append xml [tag_close];                  # Component
                
}

proc msibuild::generate_features {} {
    variable selected_features

    set xml ""
    dict for {fid feature} $selected_features {
        set absent allow
        if {[dict exists $feature Mandatory]} {
            set mandatory [dict get $feature Mandatory]
            if {$mandatory == 2} {
                set absent disallow
            } elseif {$mandatory == 1} {
            } else {
            }
        }
        if {[dict exists $feature Version]} {
            set version " Version [dict get $feature Version]"
        } else {
            set version ""
        }
        if {[dict exists $feature Description]} {
            set description "[dict get $feature Description]$version"
        } else {
            set description $version
        }

        append xml [tag Feature \
                        Id $fid \
                        Level 1 \
                        Title [dict get $feature Name] \
                        Description $description \
                        Absent $absent]
        foreach path [dict get $feature Files] {
            append xml [generate_file $path]
        }

        append xml [tag_close]; # Feature
                        
    }
    return $xml
}

proc msibuild::generate {} {
    variable tcl_root
    
    set xml "<?xml version='1.0' encoding='UTF-8'?>\n"

    append xml [tag Wix xmlns http://schemas.microsoft.com/wix/2006/wi]

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
                    Compressed       yes \
                    Id               * \
                    InstallerVersion 301 \
                    Description      "Installer for Tcl/Tk"]

    # Media - does not really matter. We don't have multiple media.
    # EmbedCab - yes because the files will be embedded inside the MSI,
    #   and not stored separately
    append xml [tag/ Media \
                    Id       1 \
                    Cabinet  media1.cab \
                    EmbedCab yes]

    append xml [generate_directory $tcl_root]
    append xml [generate_features]
    append xml [tag_close_all]

    return $xml
}

proc msibuild::main {args} {
    variable selected_features
    variable feature_definitions
    variable tcl_root

    if {[llength $args] == 0} {
        set selected_features $feature_definitions
    } else {
        set selected_features [dict filter $feature_definitions key {*}$args]
    }

    build_file_paths 
    return [generate]
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
        append xml [tag_close]
    }
    return $xml
}

puts [msibuild::main {*}$::argv]
