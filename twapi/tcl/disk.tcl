#
# Copyright (c) 2003, 2008 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# TBD - convert file spec to drive root path

# Get info associated with a drive
proc twapi::get_volume_info {drive args} {
    variable windefs

    set drive [_drive_rootpath $drive]

    array set opts [parseargs args {
        all size freespace used useravail type serialnum label maxcomponentlen fstype attr device extents
    } -maxleftover 0]

    if {$opts(all)} {
        # -all option does not cover -type, -extents and -device
        foreach opt {
            all size freespace used useravail serialnum label maxcomponentlen fstype attr
        } {
            set opts($opt) 1
        }
    }

    set result [list ]
    if {$opts(size) || $opts(freespace) || $opts(used) || $opts(useravail)} {
        lassign  [GetDiskFreeSpaceEx $drive] useravail size freespace
        foreach opt {size freespace useravail}  {
            if {$opts($opt)} {
                lappend result -$opt [set $opt]
            }
        }
        if {$opts(used)} {
            lappend result -used [expr {$size - $freespace}]
        }
    }

    if {$opts(type)} {
        set drive_type [get_drive_type $drive]
        lappend result -type $drive_type
    }
    if {$opts(device)} {
        if {[_is_unc $drive]} {
            # UNC paths cannot be used with QueryDosDevice
            lappend result -device ""
        } else {
            lappend result -device [QueryDosDevice [string range $drive 0 1]]
        }
    }

    if {$opts(extents)} {
        set extents {}
        if {! [_is_unc $drive]} {
            trap {
                set device_handle [create_file "\\\\.\\[string range $drive 0 1]" -createdisposition open_existing]
                set bin [device_ioctl $device_handle 0x560000]
                if {[binary scan $bin i nextents] != 1} {
                    error "Truncated information returned from ioctl 0x560000"
                }
                set off 8
                for {set i 0} {$i < $nextents} {incr i} {
                    if {[binary scan $bin "@$off i x4 w w" extent(-disknumber) extent(-startingoffset) extent(-extentlength)] != 3} {
                        error "Truncated information returned from ioctl 0x560000"
                    }
                    lappend extents [array get extent]
                    incr off 24; # Size of one extent element
                }
            } onerror {} {
                # Do nothing, device does not support extents or access denied
                # Empty list is returned
            } finally {
                if {[info exists device_handle]} {
                    CloseHandle $device_handle
                }
            }
        }

        lappend result -extents $extents
    }

    if {$opts(serialnum) || $opts(label) || $opts(maxcomponentlen)
        || $opts(fstype) || $opts(attr)} {
        foreach {label serialnum maxcomponentlen attr fstype} \
            [GetVolumeInformation $drive] { break }
        foreach opt {label maxcomponentlen fstype}  {
            if {$opts($opt)} {
                lappend result -$opt [set $opt]
            }
        }
        if {$opts(serialnum)} {
            set low [expr {$serialnum & 0x0000ffff}]
            set high [expr {($serialnum >> 16) & 0x0000ffff}]
            lappend result -serialnum [format "%.4X-%.4X" $high $low]
        }
        if {$opts(attr)} {
            set attrs [list ]
            foreach val {
                case_preserved_names
                unicode_on_disk
                persistent_acls
                file_compression
                volume_quotas
                supports_sparse_files
                supports_reparse_points
                supports_remote_storage
                volume_is_compressed
                supports_object_ids
                supports_encryption
                named_streams
                read_only_volume
            } {
                # Coincidentally, the attribute values happen to match
                # the corresponding constant defines
                set cdef "FILE_[string toupper $val]"
                if {$attr & $windefs($cdef)} {
                    lappend attrs $val
                }
            }
            lappend result -attr $attrs
        }
    }

    return $result
}
interp alias {} twapi::get_drive_info {} twapi::get_volume_info


# Check if disk has at least n bytes available for the user (NOT total free)
proc twapi::user_drive_space_available {drv space} {
    return [expr {$space <= [lindex [get_drive_info $drv -useravail] 1]}]
}

# Get the drive type
proc twapi::get_drive_type {drive} {
    # set type [GetDriveType "[string trimright $drive :/\\]:\\"]
    set type [GetDriveType [_drive_rootpath $drive]]
    switch -exact -- $type {
        0 { return unknown}
        1 { return invalid}
        2 { return removable}
        3 { return fixed}
        4 { return remote}
        5 { return cdrom}
        6 { return ramdisk}
    }
}

# Get list of drives
proc twapi::find_logical_drives {args} {
    array set opts [parseargs args {type.arg}]

    set drives [list ]
    foreach drive [_drivemask_to_drivelist [GetLogicalDrives]] {
        if {(![info exists opts(type)]) ||
            [lsearch -exact $opts(type) [get_drive_type $drive]] >= 0} {
            lappend drives $drive
        }
    }
    return $drives
}
interp alias {} twapi::get_logical_drives {} twapi::find_logical_drives

# Set the drive label
proc twapi::set_drive_label {drive label} {
    SetVolumeLabel [_drive_rootpath $drive] $label
}

# Maps a drive letter to the given path
proc twapi::map_drive_local {drive path args} {
    array set opts [parseargs args {raw}]

    set drive [string range [_drive_rootpath $drive] 0 1]
    DefineDosDevice $opts(raw) $drive [file nativename $path]
}


# Unmaps a drive letter
proc twapi::unmap_drive_local {drive args} {
    array set opts [parseargs args {
        path.arg
        raw
    } -nulldefault]

    set drive [string range [_drive_rootpath $drive] 0 1]

    set flags $opts(raw)
    setbits flags 0x2;                  # DDD_REMOVE_DEFINITION
    if {$opts(path) ne ""} {
        setbits flags 0x4;              # DDD_EXACT_MATCH_ON_REMOVE
    }
    DefineDosDevice $flags $drive [file nativename $opts(path)]
}


# Callback from C code
proc twapi::_filesystem_monitor_handler {id changes} {
    variable _filesystem_monitor_scripts
    if {[info exists _filesystem_monitor_scripts($id)]} {
        return [eval $_filesystem_monitor_scripts($id) [list $id $changes]]
    } else {
        # Callback queued after close. Ignore
    }
}

# Monitor file changes
proc twapi::begin_filesystem_monitor {path script args} {
    variable _filesystem_monitor_scripts

    array set opts [parseargs args {
        {subtree.bool  0}
        {filename.bool 0 0x1}
        {dirname.bool  0 0x2}
        {attr.bool     0 0x4}
        {size.bool     0 0x8}
        {write.bool    0 0x10}
        {access.bool   0 0x20}
        {create.bool   0 0x40}
        {secd.bool     0 0x100}
        {pattern.arg ""}
        {patterns.arg ""}
    } -maxleftover 0]

    if {[string length $opts(pattern)] &&
        [llength $opts(patterns)]} {
        error "Options -pattern and -patterns are mutually exclusive. Note option -pattern is deprecated."
    }

    if {[string length $opts(pattern)]} {
        # Old style single pattern. Convert to new -patterns
        set opts(patterns) [list "+$opts(pattern)"]
    }

    # Change to use \ style path separator as that is what the file monitoring functions return
    if {[llength $opts(patterns)]} {
        foreach pat $opts(patterns) {
            # Note / is replaced by \\ within the pattern
            # since \ needs to be escaped with another \ within
            # string match patterns
            lappend pats [string map [list / \\\\] $pat]
        }
        set opts(patterns) $pats
    }

    set flags [expr { $opts(filename) | $opts(dirname) | $opts(attr) |
                      $opts(size) | $opts(write) | $opts(access) |
                      $opts(create) | $opts(secd)}]

    if {! $flags} {
        # If no options specified, default to all
        set flags 0x17f
    }

    set id [Twapi_RegisterDirectoryMonitor $path $opts(subtree) $flags $opts(patterns)]
    set _filesystem_monitor_scripts($id) $script
    return $id
}

# Stop monitoring of files
proc twapi::cancel_filesystem_monitor {id} {
    variable _filesystem_monitor_scripts
    if {[info exists _filesystem_monitor_scripts($id)]} {
        Twapi_UnregisterDirectoryMonitor $id
        unset _filesystem_monitor_scripts($id)
    }
}


# Get list of volumes
proc twapi::find_volumes {} {
    set vols [list ]
    set found 1
    # Assumes there has to be at least one volume
    lassign [FindFirstVolume] handle vol
    while {$found} {
        lappend vols $vol
        lassign [FindNextVolume $handle] found vol
    }
    FindVolumeClose $handle
    return $vols
}

# Get list of volume mount points
proc twapi::find_volume_mount_points {vol} {
    set mntpts [list ]
    set found 1
    trap {
        lassign  [FindFirstVolumeMountPoint $vol] handle mntpt
    } onerror {TWAPI_WIN32 18} {
        # ERROR_NO_MORE_FILES
        # No volume mount points
        return [list ]
    } onerror {TWAPI_WIN32 3} {
        # Volume does not support them
        return [list ]
    }

    # At least one volume found
    while {$found} {
        lappend mntpts $mntpt
        lassign  [FindNextVolumeMountPoint $handle] found mntpt
    }
    FindVolumeMountPointClose $handle
    return $mntpts
}

# Set volume mount point
proc twapi::mount_volume {volpt volname} {
    # Note we don't use _drive_rootpath for trimming since may not be root path
    SetVolumeMountPoint "[string trimright $volpt /\\]\\" "[string trimright $volname /\\]\\"
}

# Delete volume mount point
proc twapi::unmount_volume {volpt} {
    # Note we don't use _drive_rootpath for trimming since may not be root path
    DeleteVolumeMountPoint "[string trimright $volpt /\\]\\"
}

# Get the volume mounted at a volume mount point
proc twapi::get_mounted_volume_name {volpt} {
    # Note we don't use _drive_rootpath for trimming since may not be root path
    return [GetVolumeNameForVolumeMountPoint "[string trimright $volpt /\\]\\"]
}

# Get the mount point corresponding to a given path
proc twapi::get_volume_mount_point_for_path {path} {
    return [GetVolumePathName [file nativename $path]]
}

# Show property dialog for a volume
proc twapi::volume_properties_dialog {name args} {
    array set opts [parseargs args {
        {hwin.int 0}
        {page.arg ""}
    } -maxleftover 0]

    shell_object_properties_dialog $name -type volume -hwin $opts(hwin) -page $opts(page)
}


# Retrieve version info for a file
proc twapi::get_file_version_resource {path args} {
    # TBD add -datetime opt to return date and time from fixed version struct
    array set opts [parseargs args {
        all
        datetime
        signature
        structversion
        fileversion
        productversion
        flags
        fileos
        filetype
        foundlangid
        foundcodepage
        langid.arg
        codepage.arg
    }]


    set ver [Twapi_GetFileVersionInfo $path]

    trap {
        array set verinfo [Twapi_VerQueryValue_FIXEDFILEINFO $ver]

        set result [list ]
        if {$opts(all) || $opts(signature)} {
            lappend result -signature [format 0x%x $verinfo(dwSignature)]
        }

        if {$opts(all) || $opts(structversion)} {
            lappend result -structversion "[expr {0xffff & ($verinfo(dwStrucVersion) >> 16)}].[expr {0xffff & $verinfo(dwStrucVersion)}]"
        }

        if {$opts(all) || $opts(fileversion)} {
            lappend result -fileversion "[expr {0xffff & ($verinfo(dwFileVersionMS) >> 16)}].[expr {0xffff & $verinfo(dwFileVersionMS)}].[expr {0xffff & ($verinfo(dwFileVersionLS) >> 16)}].[expr {0xffff & $verinfo(dwFileVersionLS)}]"
        }

        if {$opts(all) || $opts(productversion)} {
            lappend result -productversion "[expr {0xffff & ($verinfo(dwProductVersionMS) >> 16)}].[expr {0xffff & $verinfo(dwProductVersionMS)}].[expr {0xffff & ($verinfo(dwProductVersionLS) >> 16)}].[expr {0xffff & $verinfo(dwProductVersionLS)}]"
        }

        if {$opts(all) || $opts(flags)} {
            set flags [expr {$verinfo(dwFileFlags) & $verinfo(dwFileFlagsMask)}]
            lappend result -flags \
                [_make_symbolic_bitmask \
                     [expr {$verinfo(dwFileFlags) & $verinfo(dwFileFlagsMask)}] \
                     {
                         debug 1
                         prerelease 2
                         patched 4
                         privatebuild 8
                         infoinferred 16
                         specialbuild 32
                     } \
                     ]
        }

        if {$opts(all) || $opts(fileos)} {
            switch -exact -- [format %08x $verinfo(dwFileOS)] {
                00010000 {set os dos}
                00020000 {set os os216}
                00030000 {set os os232}
                00040000 {set os nt}
                00050000 {set os wince}
                00000001 {set os windows16}
                00000002 {set os pm16}
                00000003 {set os pm32}
                00000004 {set os windows32}
                00010001 {set os dos_windows16}
                00010004 {set os dos_windows32}
                00020002 {set os os216_pm16}
                00030003 {set os os232_pm32}
                00040004 {set os nt_windows32}
                default {set os $verinfo(dwFileOS)}
            }
            lappend result -fileos $os
        }

        if {$opts(all) || $opts(filetype)} {
            switch -exact -- [expr {0+$verinfo(dwFileType)}] {
                1 {set type application}
                2 {set type dll}
                3 {
                    set type "driver."
                    switch -exact -- [expr {0+$verinfo(dwFileSubtype)}] {
                        1 {append type printer}
                        2 {append type keyboard}
                        3 {append type language}
                        4 {append type display}
                        5 {append type mouse}
                        6 {append type network}
                        7 {append type system}
                        8 {append type installable}
                        9  {append type sound}
                        10 {append type comm}
                        11 {append type inputmethod}
                        12 {append type versionedprinter}
                        default {append type $verinfo(dwFileSubtype)}
                    }
                }
                4 {
                    set type "font."
                    switch -exact -- [expr {0+$verinfo(dwFileSubtype)}] {
                        1 {append type raster}
                        2 {append type vector}
                        3 {append type truetype}
                        default {append type $verinfo(dwFileSubtype)}
                    }
                }
                5 { set type "vxd.$verinfo(dwFileSubtype)" }
                7 {set type staticlib}
                default {
                    set type "$verinfo(dwFileType).$verinfo(dwFileSubtype)"
                }
            }
            lappend result -filetype $type
        }

        if {$opts(all) || $opts(datetime)} {
            lappend result -datetime [expr {(wide($verinfo(dwFileDateMS)) << 32) + $verinfo(dwFileDateLS)}]
        }

        # Any remaining arguments are treated as string names

        if {[llength $args] || $opts(foundlangid) || $opts(foundcodepage) || $opts(all)} {
            # Find list of langid's and codepages and do closest match
            set langid [expr {[info exists opts(langid)] ? $opts(langid) : [get_user_ui_langid]}]
            set primary_langid [extract_primary_langid $langid]
            set sub_langid     [extract_sublanguage_langid $langid]
            set cp [expr {[info exists opts(codepage)] ? $opts(codepage) : 0}]

            # Find a match in the following order:
            # 0 Exact match for both langid and codepage
            # 1 Exact match for langid
            # 2 Primary langid matches (sublang does not) and exact codepage
            # 3 Primary langid matches (sublang does not)
            # 4 Language neutral
            # 5 English
            # 6 First langcp in list or "00000000"
            set match(7) "00000000";    # In case list is empty
            foreach langcp [Twapi_VerQueryValue_TRANSLATIONS $ver] {
                set verlangid 0x[string range $langcp 0 3]
                set vercp 0x[string range $langcp 4 7]
                if {$verlangid == $langid && $vercp == $cp} {
                    set match(0) $langcp
                    break;              # No need to look further
                }
                if {[info exists match(1)]} continue
                if {$verlangid == $langid} {
                    set match(1) $langcp
                    continue;           # Continue to look for match(0)
                }
                if {[info exists match(2)]} continue
                set verprimary [extract_primary_langid $verlangid]
                if {$verprimary == $primary_langid && $vercp == $cp} {
                    set match(2) $langcp
                    continue;       # Continue to look for match(1) or better
                }
                if {[info exists match(3)]} continue
                if {$verprimary == $primary_langid} {
                    set match(3) $langcp
                    continue;       # Continue to look for match(2) or better
                }
                if {[info exists match(4)]} continue
                if {$verprimary == 0} {
                    set match(4) $langcp; # LANG_NEUTRAL
                    continue;       # Continue to look for match(3) or better
                }
                if {[info exists match(5)]} continue
                if {$verprimary == 9} {
                    set match(5) $langcp; # English
                    continue;       # Continue to look for match(4) or better
                }
                if {![info exists match(6)]} {
                    set match(6) $langcp
                }
            }

            # Figure out what is the best match we have
            for {set i 0} {$i <= 7} {incr i} {
                if {[info exists match($i)]} {
                    break
                }
            }

            if {$opts(foundlangid) || $opts(all)} {
                set langid 0x[string range $match($i) 0 3] 
                lappend result -foundlangid [list $langid [VerLanguageName $langid]]
            }

            if {$opts(foundcodepage) || $opts(all)} {
                lappend result -foundcodepage 0x[string range $match($i) 4 7]
            }

            foreach sname $args {
                lappend result $sname [Twapi_VerQueryValue_STRING $ver $match($i) $sname]
            }

        }

    } finally {
        Twapi_FreeFileVersionInfo $ver
    }

    return $result
}

# Return the times associated with a file
proc twapi::get_file_times {fd args} {
    array set opts [parseargs args {
        all
        mtime
        ctime
        atime
    } -maxleftover 0]

    # Figure out if fd is a file path, Tcl channel or a handle
    set close_handle false
    if {[file exists $fd]} {
        # It's a file name
        # 0x02000000 -> FILE_FLAG_BACKUP_SEMANTICS, always required in case 
        # opening a directory (even if SeBackupPrivilege is not held
        set h [create_file $fd -createdisposition open_existing -flags 0x02000000]
        set close_handle true
    } elseif {[catch {fconfigure $fd}]} {
        # Not a Tcl channel, See if handle
        if {[Twapi_IsPtr $fd]} {
            set h $fd
        } else {
            error "$fd is not an existing file, handle or Tcl channel."
        }
    } else {
        # Tcl channel
        set h [get_tcl_channel_handle $fd read]
    }

    set result [list ]

    foreach opt {ctime atime mtime} time [GetFileTime $h] {
        if {$opts(all) || $opts($opt)} {
            lappend result -$opt $time
        }
    }

    if {$close_handle} {
        CloseHandle $h
    }

    return $result
}


# Set the times associated with a file
proc twapi::set_file_times {fd args} {

    array set opts [parseargs args {
        mtime.arg
        ctime.arg
        atime.arg
        preserveatime
    } -maxleftover 0 -nulldefault]

    if {$opts(atime) ne "" && $opts(preserveatime)} {
        win32_error 87 "Cannot specify -atime and -preserveatime at the same time."
    }
    if {$opts(preserveatime)} {
        set opts(atime) -1;             # Meaning preserve access to original
    }

    # Figure out if fd is a file path, Tcl channel or a handle
    set close_handle false
    if {[file exists $fd]} {
        if {$opts(preserveatime)} {
            win32_error 87 "Cannot specify -preserveatime unless file is specified as a Tcl channel or a Win32 handle."
        }

        # It's a file name
        # 0x02000000 -> FILE_FLAG_BACKUP_SEMANTICS, always required in case 
        # opening a directory (even if SeBackupPrivilege is not held
        set h [create_file $fd -access {generic_write} -createdisposition open_existing -flags 0x02000000]
        set close_handle true
    } elseif {[catch {fconfigure $fd}]} {
        # Not a Tcl channel, assume a handle
        set h $fd
    } else {
        # Tcl channel
        set h [get_tcl_channel_handle $fd read]
    }

    SetFileTime $h $opts(ctime) $opts(atime) $opts(mtime)

    if {$close_handle} {
        CloseHandle $h
    }

    return
}

# Return a list of physical disks. Note CD-ROMs and floppies not included
proc twapi::find_physical_disks {} {
    # Disk interface class guid
    set guid {{53F56307-B6BF-11D0-94F2-00A0C91EFB8B}}
    set hdevinfo [update_devinfoset \
                      -guid $guid \
                      -presentonly true \
                      -classtype interface]
    trap {
        return [kl_flatten [get_devinfoset_interface_details $hdevinfo $guid -devicepath] -devicepath]
    } finally {
        close_devinfoset $hdevinfo
    }
}

# Return information about a physical disk
proc twapi::get_physical_disk_info {disk args} {
    set result [list ]

    array set opts [parseargs args {
        geometry
        layout
        all
    } -maxleftover 0]

    if {$opts(all) || $opts(geometry) || $opts(layout)} {
        set h [create_file $disk -createdisposition open_existing]
    }

    trap {
        if {$opts(all) || $opts(geometry)} {
            # IOCTL_DISK_GET_DRIVE_GEOMETRY - 0x70000
            if {[binary scan [device_ioctl $h 0x70000] "wiiii" geom(-cylinders) geom(-mediatype) geom(-trackspercylinder) geom(-sectorspertrack) geom(-bytespersector)] != 5} {
                error "DeviceIoControl 0x70000 on disk '$disk' returned insufficient data."
            }
            lappend result -geometry [array get geom]
        }

        if {$opts(all) || $opts(layout)} {
            # IOCTL_DISK_GET_DRIVE_LAYOUT_EX is not supported on Win2K
            if {[min_os_version 5 1] && ![info exists ::twapi::_use_win2k_disk_ioctls]} {
                # XP and later - IOCTL_DISK_GET_DRIVE_LAYOUT_EX
                set data [device_ioctl $h 0x70050]
                if {[binary scan $data "i i" partstyle layout(-partitioncount)] != 2} {
                    error "DeviceIoControl 0x70050 on disk '$disk' returned insufficient data."
                }
                set layout(-partitionstyle) [_partition_style_sym $partstyle]
                switch -exact -- $layout(-partitionstyle) {
                    mbr {
                        if {[binary scan $data "@8 i" layout(-signature)] != 1} {
                            error "DeviceIoControl 0x70050 on disk '$disk' returned insufficient data."
                        }
                    }
                    gpt {
                        set pi(-diskid) [_binary_to_guid $data 32]
                        if {[binary scan $data "@8 w w i" layout(-startingusableoffset) layout(-usablelength) layout(-maxpartitioncount)] != 3} {
                            error "DeviceIoControl 0x70050 on disk '$disk' returned insufficient data."
                        }
                    }
                    raw -
                    unknown {
                        # No fields to add
                    }
                }

                set layout(-partitions) [list ]
                for {set i 0} {$i < $layout(-partitioncount)} {incr i} {
                    # Decode each partition in turn. Sizeof of PARTITION_INFORMATION_EX is 144
                    lappend layout(-partitions) [_decode_PARTITION_INFORMATION_EX_binary $data [expr {48 + (144*$i)}]]
                }
            } else {
                # Win2K - IOCTL_DISK_GET_DRIVE_LAYOUT
                set data [device_ioctl $h 0x7400c]
                if {[binary scan $data "i i" layout(-partitioncount) layout(-signature)] != 2} {
                    error "DeviceIoControl 0x7400C on disk '$disk' returned insufficient data."
                }
                set layout(-partitions) [list ]
                for {set i 0} {$i < $layout(-partitioncount)} {incr i} {
                    # Devode each partition in turn. Sizeof of PARTITION_INFORMATION is 32
                    lappend layout(-partitions) [_decode_PARTITION_INFORMATION_binary $data [expr {8 + (32*$i)}]]
                }
            }
            lappend result -layout [array get layout]
        }

    } finally {
        if {[info exists h]} {
            CloseHandle $h
        }
    }

    return $result
}

# Wrapper around CreateFile
proc twapi::create_file {path args} {
    array set opts [parseargs args {
        {access.arg {generic_read}}
        {share.arg {read write delete}}
        {inherit.bool 0}
        {secd.arg ""}
        {createdisposition.arg open_always}
        {flags.int 0}
        {templatefile.arg NULL}
    } -maxleftover 0]

    set access_mode [_access_rights_to_mask $opts(access)]
    set share_mode [_share_mode_to_mask $opts(share)]
    set create_disposition [_create_disposition_to_code $opts(createdisposition)]
    return [CreateFile $path \
                $access_mode \
                $share_mode \
                [_make_secattr $opts(secd) $opts(inherit)] \
                $create_disposition \
                $opts(flags) \
                $opts(templatefile)]
}


# Convert a device based path to a normalized Win32 path with drive letters
# TBD - document
proc twapi::normalize_device_rooted_path {path args} {
    # For example, we need to map \Device\HarddiskVolume1 to C:
    # Can only do that by enumerating logical drives
    set npath [file nativename $path]
    if {![string match -nocase {\\Device\\*} $npath]} {
        error "$path is not a valid device based path."
    }
    array set device_map {}
    foreach drive [find_logical_drives] {
        set device_path [lindex [lindex [get_volume_info $drive -device] 1] 0]
        if {$device_path ne ""} {
            set len [string length $device_path]
            if {[string equal -nocase -length $len $path $device_path]} {
                # Prefix matches, must be terminated by end or path separator
                set ch [string index $npath $len]
                if {$ch eq "" || $ch eq "\\"} {
                    set path ${drive}[string range $npath $len end]
                    if {[llength $args]} {
                        upvar [lindex $args 0] retvar
                        set retvar $path
                        return 1
                    } else {
                        return $path
                    }
                }
            }
        }
    }

    if {[llength $args]} {
        return 0
    } else {
        error "Could not map device based path '$path'"
    }
}

# Utility functions

proc twapi::_drive_rootpath {drive} {
    if {[_is_unc $drive]} {
        # UNC
        return "[string trimright $drive ]\\"
    } else {
        return "[string trimright $drive :/\\]:\\"
    }
}

proc twapi::_is_unc {path} {
    return [expr {[string match {\\\\*} $path] || [string match //* $path]}]
}

# Map bit mask to list of drive letters
proc twapi::_drivemask_to_drivelist {drivebits} {
    set drives [list ]
    set i 0
    foreach drive {A B C D E F G H I J K L M N O P Q R S T U V W X Y Z} {
        if {[expr {$drivebits & (1 << $i)}]} {
            lappend drives $drive:
        }
        incr i
    }
    return $drives
}

# Given a Tcl binary and offset, decode the PARTITION_INFORMATION record
proc twapi::_decode_PARTITION_INFORMATION_binary {bin off} {
    if {[binary scan $bin "@$off w w i i c c c c" \
             pi(-startingoffset) \
             pi(-partitionlength) \
             pi(-hiddensectors) \
             pi(-partitionnumber) \
             pi(-partitiontype) \
             pi(-bootindicator) \
             pi(-recognizedpartition) \
             pi(-rewritepartition)] != 8} {
        error "Truncated partition structure."
    }

    # Show partition type in hex, not negative number
    set pi(-partitiontype) [format 0x%2.2x [expr {0xff & $pi(-partitiontype)}]]

    return [array get pi]
}

# Given a Tcl binary and offset, decode the PARTITION_INFORMATION_EX record
proc twapi::_decode_PARTITION_INFORMATION_EX_binary {bin off} {
    if {[binary scan $bin "@$off i x4 w w i c" \
             pi(-partitionstyle) \
             pi(-startingoffset) \
             pi(-partitionlength) \
             pi(-partitionnumber) \
             pi(-rewritepartition)] != 5} {
        error "Truncated partition structure."
    }

    set pi(-partitionstyle) [_partition_style_sym $pi(-partitionstyle)]

    # MBR/GPT are at offset 32 in the structure
    switch -exact -- $pi(-partitionstyle) {
        mbr {
            if {[binary scan $bin "@$off x32 c c c x i" pi(-partitiontype) pi(-bootindicator) pi(-recognizedpartition) pi(-hiddensectors)] != 4} {
                error "Truncated partition structure."
            }
            # Show partition type in hex, not negative number
            set pi(-partitiontype) [format 0x%2.2x [expr {0xff & $pi(-partitiontype)}]]
        }
        gpt {
            set pi(-partitiontype) [_binary_to_guid $bin [expr {$off+32}]]
            set pi(-partitionif)   [_binary_to_guid $bin [expr {$off+48}]]
            if {[binary scan $bin "@$off x64 w" pi(-attributes)] != 1} {
                error "Truncated partition structure."
            }
            set pi(-name) [_ucs16_binary_to_string [string range $bin [expr {$off+72} end]]]
        }
        raw -
        unknown {
            # No fields to add
        }

    }


    return [array get pi]
}

# Map a partition style code to a symbol
proc twapi::_partition_style_sym {partstyle} {
    set partstyle [lindex {mbr gpt raw} $partstyle]
    if {$partstyle ne ""} {
        return $partstyle
    }
    return "unknown"
}

# Map a set of share mode symbols to a flag. Concatenates
# all the arguments, and then OR's the individual elements. Each
# element may either be a integer or one of the share modes
proc twapi::_share_mode_to_mask {modelist} {
    # Values correspond to FILE_SHARE_* defines
    return [_parse_symbolic_bitmask $modelist {read 1 write 2 delete 4}]
}

# Map the symbolic CreateDisposition parameter of CreateFile to integer values
proc twapi::_create_disposition_to_code {sym} {
    if {[string is integer -strict $sym]} {
        return $sym
    }
    set code [lsearch -exact {dummy create_new create_always open_existing open_always truncate_existing} $sym]
    if {$code == -1} {
        error "Invalid create disposition value '$sym'"
    }
    return $code
}
