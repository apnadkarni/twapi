# This script must be invoked by the Tcl shell from the Tcl distribution
# for which the MSI (Microsoft Installer) package is to be built.

# Make sure we are not picking up anything outside our installation
# even if TCLLIBPATH etc. is set
if {0 && [llength [array names env TCL*]] || [llength [array names env tcl*]]} {
    error "[array names env TCL*] environment variables must not be set"
}

package require fileutil

namespace eval msibuild {
    variable script_dir
    set script_dir [file attributes [file dirname [info script]] -shortname]

    # Strings and values used in the MSI to identify the product.
    # Dictionary keyed by platform x86 or x64
    variable architecture [expr {$::tcl_platform(pointerSize) == 8 ? "x64" : "x86"}]
    variable msi_strings
    array set msi_strings {
        ProductName        "Tcl/Tk for Windows"
        ProgramMenuDir     ProgramMenuDir
    }
    if {$architecture eq "x86"} {
        array set msi_strings {
            UpgradeCode        9888EC4F-7EB8-40EF-8506-7230E811AFE9
            ProgramFilesFolder ProgramFilesFolder
            Win64              no
            ArchString         32-bit
        }
    } else {
        array set msi_strings {
            UpgradeCode        1AE719B3-0895-4913-B8BF-1117944A7046
            ProgramFilesFolder ProgramFiles64Folder
            Win64              yes
            ArchString         64-bit
        }
    }
    
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
    # Files - the list of files for that feature
    # Version - version of the package if available
    variable feature_definitions

    # This should be read from a config file. Oh well. Later ...
    set feature_definitions {
        core {
            Name {Tcl/Tk Core}
            TclPackages { Tcl Tk }
            Description {Includes Tcl and Tk core, Windows registry and DDE extensions.}
            Paths {bin lib/tcl8 lib/tcl8.* lib/dde* lib/reg* lib/tk8.*}
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
        thread {
            Name {Thread package}
            TclPackages Thread
            Description {Extension for script-level access to Tcl threads}
        }
        clibs {
            {C libraries} 
            Description {C libraries for building your own binary extensions}
            Paths {include lib/*.lib}
            Mandatory 0
        }
    }        

    # selected_features contains the actual selected features from the above.
    variable selected_features
    
    # Root of the Tcl installation.
    # Must be normalized else fileutil::relative etc. will not work
    # because they do not handle case differences correctly
    variable tcl_root [file normalize [file dirname [file dirname [info nameofexecutable]]]]

    # Contains an array of all directory paths. We use an array instead
    # of a nested dict structure so that we can make use of dict lappend
    # on the contained dictionaries.
    # Again keys are normalized for same reason as above.
    # Elements are dictionaries with the following keys:
    #  Id - the Wix Id for the directory
    #  Files - list of normalized file paths contained in the directory
    variable directories
    array set directories {}
    
    # Directory tree structure. Maintained as a dictionary where nested
    # keys are subdirectories. Tree is rooted relative to $tcl_root
    variable directory_tree {}

    # Array to track file ids
    variable file_ids
    array set file_ids {}
    
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

proc msibuild::component_id {file_id} {
    return CMP_$file_id
}

# Returns the id of the Tcl bin directory
proc msibuild::bin_dir_id {} {
    variable directories
    variable tcl_root
    return [dict get $directories([file join $tcl_root bin]) Id]
}

# Returns the id of the Tcl exe
proc msibuild::tclsh_id {} {
    variable file_ids
    variable tcl_root
    return $file_ids([file join $tcl_root bin tclsh.exe])
}

# Returns the id of the wish exe
proc msibuild::wish_id {} {
    variable file_ids
    return $file_ids([file join $tcl_root bin wish.exe])
}

# Build a file path list for a MSI package.
# The files are added to the corresponding directory
proc msibuild::build_file_paths_for_feature {feature} {
    variable selected_features
    variable tcl_root
    variable directories
    variable file_ids

    log Building file paths for $feature
    
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

    dict set selected_features $feature Files $files
    foreach path $files {
        set file_ids($path) [id $path]
        add_parent_directory $path
        dict lappend directories([file dirname $path]) Files $path
    }
}

# Adds all ancestors of $path (must be normalized) to the directories array
proc msibuild::add_parent_directory {path} {
    variable directories
    variable directory_tree
    variable tcl_root
    
    if {[file pathtype $path] ne "absolute"} {
        error "Internal error: $path is not a absolute normalized path"
    }
    if {[file type $path] ne "file"} {
        error "Internal error: $path is not an ordinary file"
    }

    set parent [file dirname $path]

    # Add parent and ancestore (up to tcl root) to the directory tree
    set relpath [fileutil::relative $tcl_root $parent]
    if {$relpath eq "."} {
        # No need to do anything with the root. It is always preserved.
        return
    }
    dict set directory_tree {*}[split $relpath /] {}

    # Generate ids for all ancestors
    while {$parent ne $tcl_root} {
        if {[info exists directories($parent)]} {
            return;             # We have already seen this
        }
        dict set directories($parent) Id [id $parent]
        set parent [file dirname $parent]
    }
}

# Builds the file paths for all files contained in all MSI package
proc msibuild::build_file_paths {} {
    variable selected_features

    foreach feature [dict keys $selected_features] {
        build_file_paths_for_feature $feature
    }                             
}

# Generate the Directory and child nodes for all directories other
# than tcl_root itself which is special cased.
# dirpath is a list that is a path through $directory_tree
# Note assumes no links
proc msibuild::generate_directory_tree {dirpath} {
    variable directories
    variable tcl_root
    variable directory_tree

    set dir [file join $tcl_root {*}$dirpath]

    set xml [tag Directory \
                 Id [dict get $directories($dir) Id] \
                 Name [lindex $dirpath end]]

    # First write out all the files in this directory, if any
    # Note the directory will always exist in $directories but may
    # not have files
    if {[dict exists $directories($dir) Files]} {
        foreach path [dict get $directories($dir) Files] {
            append xml [generate_file $path]
        }
    }

    # Now write out all the subdirectories
    foreach subdir [dict keys [dict get $directory_tree {*}$dirpath]] {
        append xml [generate_directory_tree [linsert $dirpath end $subdir]]
    }
    
    append xml [tag_close Directory]; # Top level for this dirpath

    return $xml
}

# Generates the <Directory> entries for installation.
proc msibuild::generate_directory {} {
    variable tcl_root
    variable msi_strings
    variable directories
    variable directory_tree

    # Use of the WixUI_Advanced dialogs requires the following
    # Directory element structures.
    append xml [tag Directory Id TARGETDIR Name SourceDir]
    
    append xml [tag Directory Id $msi_strings(ProgramFilesFolder) Name PFiles]

    # Generate the top level ($tcl_root) which is not included in
    # directory_tree and also needs special handling.
    append xml [tag Directory Id APPLICATIONFOLDER Name Tcl[info tclversion]]

    # Generate the file components for the top level
    if {[info exists directories($tcl_root)] &&
        [dict exists $directories($tcl_root) Files]} {
        foreach path [dict get $directories($tcl_root) Files] {
            append xml [generate_file $path]
        }
    }
    
    # Now generate each of the subdirectories of the top level
    foreach subdir [dict keys $directory_tree] {
        append xml [generate_directory_tree [list $subdir]]
    }

    append xml [generate_file_associations]
    
    append xml [tag_close Directory]; # APPLICATIONFOLDER

    append xml [tag_close];     # ProgramFilesFolder

    # NOTE: We use ProgramMenuFolder because StartMenuFolder is not available
    # on XP Windows installer
    append xml [tag Directory Id ProgramMenuFolder]
    append xml [tag Directory Id TclStartMenuFolder Name Tcl[info tclversion]]
    append xml [tag/ Directory Id TclDocMenuFolder Name Documentation]
    append xml [tag_close Directory Directory]; # TclStartMenuFolder ProgramMenuFolder

    
    append xml [tag_close Directory];     # TARGETDIR

    return $xml
}

proc msibuild::generate_file_associations {} {
    # NOTE: This has to be under the directory structure and *indirectly*
    # referenced from the file association feature. Definining it
    # under that feature will result in errors as the fragment
    # refers to tclsh_id which is in a different feature.

    append xml [tag Component Id CMP_TclFileAssoc]

    # To associate a file, create a ProgId for Tcl. Then associate an
    # extension with it. HKMU -> HKCU for per-user and HKLM for per-machine
    set tcl_prog_id "Tcl.Application"
    append xml [tag/ RegistryValue \
                    Root HKMU \
                    Key "SOFTWARE\\Classes\\$tcl_prog_id" \
                    Name "FriendlyTypeName" \
                    Value "Tcl application" \
                    Type "string"]
    # TBD - Icon attribute for ProgId
    # TBD - Not sure of value for Advertise
    append xml [tag ProgId \
                    Id $tcl_prog_id \
                    Description "Tcl application" \
                    Advertise no]
    append xml [tag Extension Id "tcl"]
    append xml [tag/ Verb \
                    Id open \
                    TargetFile [tclsh_id] \
                    Command "Run as a Tcl application" \
                    Argument "&quot;%1&quot;"]
    append xml [tag_close Extension ProgId Component]
}

proc msibuild::generate_file {path} {
    variable file_ids

    if {[file pathtype $path] ne "absolute"} {
        error "generate_file passed a non-absolute path"
    }
    set file_id $file_ids($path)
    
    # Every FILE must be enclosed in a Component and a Component should
    # have only one file.
    set xml [tag Component \
                 Id [component_id $file_id] \
                 Guid *]
    append xml [tag/ File \
                    Id $file_id \
                    Source $path \
                    KeyPath yes]
    append xml [tag_close Component]
}

proc msibuild::generate_features {} {
    variable selected_features
    variable tcl_root
    variable file_ids

    set xml ""
    dict for {fid feature} $selected_features {
        log Generating feature $fid
        set absent allow
        if {[dict exists $feature Mandatory]} {
            switch -exact -- [dict get $feature Mandatory] {
                2 { set absent disallow }
                1 { set absent allow }
                0 {
                    # TBD - how to set the default state to NOT INSTALL ?
                    set absent allow
                }
                default {
                    error "Unknown value [dict get $feature Mandatory] for Mandatory feature definition key."
                }
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
            append xml [tag/ ComponentRef Id [component_id $file_ids($path)]]
        }
        if {$fid eq "core"} {
            # These features are subservient to the core feature
            append xml [generate_start_menu_feature]; # Option to add to Start menu
            append xml [generate_path_feature];       # Option to modify PATH
            append xml [generate_file_assoc_feature]; # Associate .tcl etc.
        }
        append xml [tag_close Feature]
    }
    return $xml
}

# Generate the UI elements
proc msibuild::generate_ui {} {
    variable script_dir
    
    append xml [tag/ UIRef Id WixUI_Advanced]
    # Following property provides the default dir name for install
    # when using WixUI_Advanced dialog.
    append xml [tag/ Property Id ApplicationFolderName Value "Tcl"]
    
    # Set the default when showing the per-user/per-machine dialog
    # Note that if per-user is chosen, the user cannot choose
    # location of install
    append xml [tag/ Property Id WixAppFolder Value WixPerUserFolder]

    # License file
    append xml [tag/ WixVariable Id WixUILicenseRtf Value [file join $script_dir license.rtf]]
    
    # The background for Install dialogs
    append xml [tag/ WixVariable Id WixUIDialogBmp Value [file join $script_dir msidialog.bmp]]
    
    # The horizontal banner for Install dialogs
    append xml [tag/ WixVariable Id WixUIBannerBmp Value [file join $script_dir msibanner.bmp]]
    
    return $xml
}

# Generate pre-launch conditions in terms of platform requirements.
proc msibuild::generate_launch_conditions {} {
    append xml [tag Condition \
                    Message "This program is only supported on Windows XP and later versions of Windows."]
    append xml {<![CDATA[VersionNT >= 501]]>}
    append xml [tag_close Condition]

    append xml [tag Condition \
                    Message "This program is requires at least Service Pack 3 on Windows XP."]
    append xml {<![CDATA[VersionNT > 501 OR ServicePackLevel >= 3]]>}
    append xml [tag_close Condition]

    return $xml
}

# Generate the Add/Remove program properties
proc msibuild::generate_arp {} {
    variable script_dir
    string cat \
        [tag/ Property Id ARPURLINFOABOUT Value http://www.tcl.tk] \
        [tag/ Property Id ARPHELPLINK Value http://www.tcl.tk/man/tcl[info tclversion]] \
        [tag/ Property Id ARPCOMMENTS Value "The Tcl programming language and Tk graphical toolkit"] \
        [tag/ Property Id ARPPRODUCTICON Value [file join $script_dir tcl.ico]]
}

# Allow user to modify path.
proc msibuild::generate_path_feature {} {
    variable tcl_root

    append xml [tag Feature \
                    Id [id] \
                    Level 1 \
                    Title {Modify Paths} \
                    Description {Modify PATH environment variable to include the Tcl/Tk directory}]

    append xml [tag Component Id [id] Guid 5C4574A9-ECE5-4565-BA0D-38AC38755C4E KeyPath yes Directory [bin_dir_id]]
    # TBD - Should System be set to "yes" for machine installs?
    append xml [tag/ Environment \
                    Action set \
                    Id [id] \
                    Name Path \
                    Value {[APPLICATIONFOLDER]bin} \
                    System no \
                    Permanent no \
                    Part last \
                    Separator ";"]
                
    append xml [tag_close Component Feature]

    return $xml
}

proc msibuild::generate_start_menu_feature {} {
    append xml [tag Feature \
                    Id [id] \
                    Level 1 \
                    Title {Start menu} \
                    Description {Install Start menu shortcuts}]

    # TBD - should WorkingDirectory be set to PersonalFolder/
    append xml [tag Component Id [id] Guid * Directory TclStartMenuFolder] 
    append xml [tag/ Shortcut Id [id] \
                    Name "Tcl command shell" \
                    Description "Console for interactive execution of commands in the Tcl language" \
                    WorkingDirectory [bin_dir_id] \
                    Target {[APPLICATIONFOLDER]bin\tclsh.exe}]
    append xml [tag/ Shortcut Id [id] \
                    Name "Tk graphical shell" \
                    Description "Graphical console for interactive execution of commands using the Tcl/Tk toolkit" \
                    WorkingDirectory [bin_dir_id] \
                    Target {[APPLICATIONFOLDER]bin\wish.exe}]
    
    # Arrange for the folder to be removed on an uninstall
    # We only include this for one shortcut component
    append xml [tag/ RemoveFolder Id RemoveTclStartMenuFolder \
                    Directory TclStartMenuFolder \
                    On uninstall]
    append xml [tag/ RegistryValue Root HKCU \
                    Key "Software\\Tcl\\[info tclversion]" \
                    Name installed \
                    Type integer \
                    Value 1 \
                    KeyPath yes]
    # APN append xml [tag_close Component]

    # APN append xml [tag Component Id [id] Guid 57009BF7-3E8D-49C8-A557-26F86943233F Directory TclStartMenuFolder]
    append xml [tag/ util:InternetShortcut \
                    Id TclManPage \
                    Name "Tcl documentation" \
                    Directory TclDocMenuFolder \
                    Target http://tcl.tk/man/tcl[info tclversion]/contents.htm \
                    Type url]
    if {1} {
    append xml [tag/ RemoveFolder Id RemoveTclDocMenuFolder \
                    Directory TclDocMenuFolder \
                    On uninstall]
    }
    if {0} {
    append xml [tag/ RegistryValue Root HKCU \
                    Key "Software\\Tcl\\[info tclversion]\\Doc" \
                    Name installed \
                    Type integer \
                    Value 1 \
                    KeyPath yes]
    }
    append xml [tag_close Component]

    if {0} {
    append xml [tag_close Component]

    # TBD - should be tied to tkcon feature
    append xml [tag Component Id [id] Guid *]
    append xml [tag/ Shortcut Id [id] \
                    Name "tkcon" \
                    Description "TkCon enhanced graphical console" \
                    Directory TclStartMenuFolder \
                    Target {[APPLICATIONFOLDER]bin\wish.exe} \
                    Arguments {[APPLICATIONFOLDER]bin\tkcon.tcl}]
    append xml [tag_close Component]
    }
    
    # TBD - icons on shortcuts? See Wix Tools book

    # Comment out desktop shortcut for now. Not really warranted
    if {0} {
        append xml [tag Component Id [id] Guid *]
        append xml [tag/ Shortcut Id [id] \
                        Name "tkcon" \
                        Description "TkCon enhanced graphical console" \
                        Directory DesktopFolder \
                        Target {[APPLICATIONFOLDER]bin\wish.exe} \
                        Arguments {[APPLICATIONFOLDER]bin\tkcon.tcl}]
        # TBD - need a registry keypath here too?
        append xml [tag_close Component]
    }
    
    append xml [tag_close Feature]
}

# Option to associate .tcl and .tk files with tclsh and tk
proc msibuild::generate_file_assoc_feature {} {
    append xml [tag Feature \
                    Id TclFileAssoc \
                    Level 1 \
                    Title {File associations} \
                    Description {Associate .tcl and .tk files with tclsh and wish}]
    # NOTE: the tclsh component must be referenced here as well otherwise
    # there will be a error in the Verb element below which refers to it.
    append xml [tag/ ComponentRef Id [component_id [tclsh_id]]]
    append xml [tag/ ComponentRef Id CMP_TclFileAssoc]

    append xml [tag_close Feature]
    return $xml
}

proc msibuild::generate {} {
    variable tcl_root
    variable msi_strings
    
    log Generating Wix XML
    
    set xml "<?xml version='1.0' encoding='windows-1252'?>\n"

    append xml [tag Wix \
                    xmlns http://schemas.microsoft.com/wix/2006/wi \
                    xmlns:util "http://schemas.microsoft.com/wix/UtilExtension"]

    # Product - info about Tcl itself
    # Name - Tcl/Tk for Windows
    # Id - "*" -> always generate a new one on every run. Makes upgrades
    #    much easier
    # UpgradeCode - must never change between releases else upgrades won't work.
    # Language, Codepage - Currently only English
    # Version - picked up from Tcl
    # Manufacturer - TCT? Tcl Community?
    append xml [tag Product \
                    Name        "Tcl/Tk for Windows ($msi_strings(ArchString))" \
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
    # NOTE: we do not set the following because it does not then allow
    #     per-machine installs
    # InstallPrivileges - "limited" if no elevation required, "elevated"
    #     if elevation required.
    # TBD - change installer version to be xp compatible
    append xml [tag/ Package \
                    Compressed       yes \
                    Id               * \
                    InstallerVersion 500 \
                    Description      "Installer for Tcl/Tk ($msi_strings(ArchString))"]

    # Checks for platforms
    append xml [generate_launch_conditions]
    
    # Upgrade behaviour. There is probably no reason to disallow downgrades
    # but I don't want to do the additional testing...
    append xml [tag/ MajorUpgrade \
                    AllowDowngrades no \
                    DowngradeErrorMessage "A later version of \[ProductName\] is already installed. Setup will now exit." \
                    AllowSameVersionUpgrades yes]
                
    # Media - does not really matter. We don't have multiple media.
    # EmbedCab - yes because the files will be embedded inside the MSI,
    #   and not stored separately
    append xml [tag/ Media \
                    Id       1 \
                    Cabinet  media1.cab \
                    EmbedCab yes]

    append xml [generate_arp];  # Information that shows in Add/Remove Programs
    
    # NOTE:Despite what the Wix reference says that it can be placed anywhere,
    # UIRef can't be the first child of Product element. So dump it here.
    append xml [generate_ui];   # Installation dialogs

    # The MSIINSTALLPERUSER only has effect on Win 7 and later using
    # Windows Installer 5.0 and is only effective if ALLUSERS=2
    # See https://msdn.microsoft.com/en-us/library/aa371865.aspx
    # The ALLUSERS property will be set based on the WixUI_Advanced dialog
    # user selection so I'm not quite sure if this is needed.
    append xml [tag/ Property Id MSIINSTALLPERUSER Secure yes Value 1]
    
    append xml [generate_directory]; # Directory structure
    append xml [generate_features];  # Feature tree

    if {0} {
        # TBD Can't get file assoc to compile
        append xml [generate_file_assoc_feature]; # Option to associate .tcl etc. with tclsh/wish
    }
    
    append xml [tag_close Product Wix]

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

# Close xml tag(s). For each argument n,
# if n is an integer, the last that many tags are popped
# off the tag stack. Otherwise, n must be the name of the topmost tag
# and that tag is popped (this is to catch tag matching errors early)
proc msibuild::tag_close {args} {
    variable xml_tags
    if {[llength $args] == 0} {
        set args [list 1]
    }
    set xml {}
    foreach n $args {
        if {![string is integer -strict $n]} {
            set expected_tag $n
            set n 1
        } else {
            unset -nocomplain expected_tag
        }

        # Pop n tags
        if {[llength $xml_tags] < $n} {
            error "XML tag stack empty"
        }
        while {$n > 0} {
            set tag [lindex $xml_tags end]
            if {[info exists expected_tag] && $tag ne $expected_tag} {
                error "Tag nesting error. Attempt to terminate $expected_tag but innermost tag is $tag"
            }
            set xml_tags [lrange $xml_tags 0 end-1]
            append xml "[string repeat {  } [llength $xml_tags]]</$tag>\n"
            incr n -1
        }
    }
    return $xml
}

proc msibuild::tag_close_all {} {
    variable xml_tags
    return [tag_close [llength $xml_tags]]
}

# Gets the next arg from ::argv and raises error if no more arguments
# argv is modified accordingly.
proc msibuild::nextarg {} {
    global argv
    if {[llength $argv] == 0} {
        error "Missing argument"
    }
    set argv [lassign $argv arg]
    return $arg
}

proc msibuild::parse_command_line {} {
    global argv
    variable feature_definitions
    variable selected_features
    variable options
    variable architecture

    array set options {
        silent 0
    }
    
    while {[llength $argv]} {
        set arg [nextarg]
        if {$arg eq "--"} {
            break;   # Rest are all passed to subcommand as component names
        }
        switch -glob -- $arg {
            -silent { set options(silent) 1 }
            -outdir {
                set options([string range $arg 1 end]) [nextarg]
            }
            -* {
                error "Unknown option \"$arg\""
            }
            default {
                lappend options(features) $arg
            }
        }
    }

    if {![info exists options(features)] || [llength $options(features)] == 0} {
        set selected_features $feature_definitions
    } else {
        set selected_features [dict filter $feature_definitions key {*}$options(features)]
    }
    
    if {![info exists options(outdir)]} {
        set options(outdir) $architecture
    }
    set options(outdir) [file normalize $options(outdir)]
}

proc msibuild::log {args} {
    variable options
    if {!$options(silent)} {
        puts [join $args { }]
    }
}

# Some pre-build steps to fix up the pool.
proc msibuild::prebuild {} {
    variable tcl_root

    # We want tclsh and wish not, tclsh86t etc.
    file copy -force [file join $tcl_root bin tclsh86t.exe] \
        [file join $tcl_root bin tclsh.exe]
    file copy -force [file join $tcl_root bin wish86t.exe] \
        [file join $tcl_root bin wish.exe]
    
}

proc msibuild::main {} {
    variable selected_features
    variable feature_definitions
    variable tcl_root
    variable options
    variable architecture
    variable msi_strings
    
    parse_command_line
    log Building $architecture MSI for [join [dict keys $selected_features] {, }]
    prebuild
    build_file_paths 
    set xml [generate]

    file mkdir $options(outdir)

    set wxs [file join $options(outdir) tcl$architecture.wxs]
    log Writing Wix XML file $wxs
    set fd [open $wxs w]
    fconfigure $fd -encoding utf-8
    puts $fd $xml
    close $fd

    if {[info exists ::env(WIX)]} {
        set candle [file join $::env(WIX) bin candle.exe]
        set light  [file join $::env(WIX) bin light.exe]
    } else {
        # Assume on the path
        set candle candle.exe
        set light  light.exe
    }

    set outdir [file attributes $options(outdir) -shortname]
    set wixobj [file join $outdir tcl$architecture.wixobj]
    log Generating Wix object file $wixobj
    if {$architecture eq "x86"} {
        set arch {}
    } else {
        set arch [list -arch x64]
    }
    exec $candle -nologo {*}$arch -ext WixUIExtension.dll -ext WixUtilExtension.dll -out $wixobj $wxs
    
    set msi [file join $outdir "Tcl Installer ($msi_strings(ArchString)).msi"]
    log Generating MSI file $msi
    exec $light -out $msi -ext WixUIExtension.dll -ext WixUtilExtension.dll $wixobj

    log MSI file $msi created.
}


# If we are the command line script run as main program.
# The ... is to ensure last path component gets normalized if it is a link
if {[info exists argv0] && 
    [file dirname [file normalize [info script]/...]] eq [file dirname [file normalize $argv0/...]]} {
    msibuild::main
}

