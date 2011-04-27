#
# Copyright (c) 2008 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# TBD - document

# Callback invoked for device changes.
# Does some processing of passed data and then invokes the
# real callback script

proc twapi::_device_notification_handler {id args} {
    variable _device_notifiers
    set idstr "devnotifier#$id"
    if {![info exists _device_notifiers($idstr)]} {
        # Notifications that expect a response default to "true"
        return 1
    }
    set script [lindex $_device_notifiers($idstr) 1]

    # For volume notifications, change drive bitmask to
    # list of drives before passing back to script
    set event [lindex $args 0]
    if {[lindex $args 1] eq "volume" &&
        ($event eq "deviceremovecomplete" || $event eq "devicearrival")} {
        lset args 2 [_drivemask_to_drivelist [lindex $args 2]]

        # Also indicate whether network volume and whether change is a media
        # change or physical change
        set attrs [list ]
        set flags [lindex $args 3]
        if {$flags & 1} {
            lappend attrs mediachange
        }
        if {$flags & 2} {
            lappend attrs networkvolume
        }
        lset args 3 $attrs
    }

    return [eval $script [list $idstr] $args]
}

proc twapi::start_device_notifier {script args} {
    variable _device_notifiers

    set script [lrange $script 0 end]; # Verify syntactically a list

    array set opts [parseargs args {
        deviceinterface.arg
        handle.arg
    } -maxleftover 0]

    # For reference - some common device interface classes
    # NOTE: NOT ALL HAVE BEEN VERIFIED!
    # Network Card      {ad498944-762f-11d0-8dcb-00c04fc3358c}
    # Human Interface Device (HID)      {4d1e55b2-f16f-11cf-88cb-001111000030}
    # GUID_DEVINTERFACE_DISK          - {53f56307-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_CDROM         - {53f56308-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_PARTITION     - {53f5630a-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_TAPE          - {53f5630b-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_WRITEONCEDISK - {53f5630c-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_VOLUME        - {53f5630d-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_MEDIUMCHANGER - {53f56310-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_FLOPPY        - {53f56311-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_CDCHANGER     - {53f56312-b6bf-11d0-94f2-00a0c91efb8b}
    # GUID_DEVINTERFACE_STORAGEPORT   - {2accfe60-c130-11d2-b082-00a0c91efb8b}
    # GUID_DEVINTERFACE_KEYBOARD      - {884b96c3-56ef-11d1-bc8c-00a0c91405dd}
    # GUID_DEVINTERFACE_MOUSE         - {378de44c-56ef-11d1-bc8c-00a0c91405dd}
    # GUID_DEVINTERFACE_PARALLEL      - {97F76EF0-F883-11D0-AF1F-0000F800845C}
    # GUID_DEVINTERFACE_COMPORT       - {86e0d1e0-8089-11d0-9ce4-08003e301f73}
    # GUID_DEVINTERFACE_DISPLAY_ADAPTER - {5b45201d-f2f2-4f3b-85bb-30ff1f953599}
    # GUID_DEVINTERFACE_USB_HUB       - {f18a0e88-c30c-11d0-8815-00a0c906bed8}
    # GUID_DEVINTERFACE_USB_DEVICE    - {A5DCBF10-6530-11D2-901F-00C04FB951ED}
    # GUID_DEVINTERFACE_USB_HOST_CONTROLLER - {3abf6f2d-71c4-462a-8a92-1e6861e6af27}


    if {[info exists opts(deviceinterface)] && [info exists opts(handle)]} {
        error "Options -deviceinterface and -handle are mutually exclusive."
    }

    if {![info exists opts(deviceinterface)]} {
        set opts(deviceinterface) ""
    }
    if {[info exists opts(handle)]} {
        set type 6
    } else {
        set opts(handle) NULL
        switch -exact -- $opts(deviceinterface) {
            port            { set type 3 ; set opts(deviceinterface) "" }
            volume          { set type 2 ; set opts(deviceinterface) "" }
            default {
                # device interface class guid or empty string (for all device interfaces)
                set type 5
            }
        }
    }

    set id [Twapi_RegisterDeviceNotification $type $opts(deviceinterface) $opts(handle)]
    set idstr "devnotifier#$id"

    set _device_notifiers($idstr) [list $id $script]
    return $idstr
}

# Stop monitoring of device changes
proc twapi::stop_device_notifier {idstr} {
    variable _device_notifiers

    if {![info exists _device_notifiers($idstr)]} {
        return;
    }

    Twapi_UnregisterDeviceNotification [lindex $_device_notifiers($idstr) 0]
    unset _device_notifiers($idstr)
}


# Retrieve a device information set for a device setup or interface class
proc twapi::update_devinfoset {args} {
    array set opts [parseargs args {
        {guid.arg ""}
        {classtype.arg setup {interface setup}}
        {presentonly.bool false 0x2}
        {currentprofileonly.bool false 0x8}
        {deviceinfoset.arg NULL}
        {hwin.int 0}
        {system.arg ""}
        {pnpname.arg ""}
    } -maxleftover 0]

    # DIGCF_ALLCLASSES is bitmask 4
    set flags [expr {$opts(guid) eq "" ? 0x4 : 0}]
    if {$opts(classtype) eq "interface"} {
        # DIGCF_DEVICEINTERFACE
        set flags [expr {$flags | 0x10}]
    }

    # DIGCF_PRESENT
    set flags [expr {$flags | $opts(presentonly)}]

    # DIGCF_PRESENT
    set flags [expr {$flags | $opts(currentprofileonly)}]

    return [SetupDiGetClassDevsEx \
                $opts(guid) \
                $opts(pnpname) \
                $opts(hwin) \
                $flags \
                $opts(deviceinfoset) \
                $opts(system)]
}


# Close and release resources for a device information set
interp alias {} twapi::close_devinfoset {} twapi::SetupDiDestroyDeviceInfoList


# Given a device information set, returns the device elements within it
proc twapi::get_devinfoset_elements {hdevinfo} {
    set result [list ]
    set i 0
    trap {
        while {true} {
            lappend result [SetupDiEnumDeviceInfo $hdevinfo $i]
            incr i
        }
    } onerror {TWAPI_WIN32 259} {
        # Fine, Just means no more items
    }

    return $result
}

# Given a device information set, returns a list of specified registry
# properties for all elements of the set
# args is list of properties to retrieve
proc twapi::get_devinfoset_registry_properties {hdevinfo args} {
    set result [list ]
    trap {
        # Keep looping until there is an error saying no more items
        set i 0
        while {true} {

            # First element is the DEVINFO_DATA element
            set devinfo_data [SetupDiEnumDeviceInfo $hdevinfo $i]
            set item [list -deviceelement $devinfo_data ]

            # Get all specified property values
            foreach prop $args {
                set prop [_device_registry_sym_to_code $prop]
                trap {
                    lappend item $prop \
                        [list success \
                             [Twapi_SetupDiGetDeviceRegistryProperty \
                                  $hdevinfo $devinfo_data $prop]]
                } onerror {} {
                    lappend item $prop [list fail $errorCode]
                }
            }
            lappend result $item

            incr i
        }
    } onerror {TWAPI_WIN32 259} {
        # Fine, Just means no more items
    }

    return $result
}


# Given a device information set, returns specified device interface
# properties
# TBD - document ?
proc twapi::get_devinfoset_interface_details {hdevinfo guid args} {
    set result [list ]

    array set opts [parseargs args {
        {matchdeviceelement.arg {}}
        interfaceclass
        flags
        devicepath
        deviceelement
        ignoreerrors
    } -maxleftover 0]

    trap {
        # Keep looping until there is an error saying no more items
        set i 0
        while {true} {
            set interface_data [SetupDiEnumDeviceInterfaces $hdevinfo \
                                    $opts(matchdeviceelement) $guid $i]
            set item [list ]
            if {$opts(interfaceclass)} {
                lappend item -interfaceclass [lindex $interface_data 0]
            }
            if {$opts(flags)} {
                set flags    [lindex $interface_data 1]
                set symflags [_make_symbolic_bitmask $flags {active 1 default 2 removed 4} false]
                lappend item -flags [linsert $symflags 0 $flags]
            }

            if {$opts(devicepath) || $opts(deviceelement)} {
                # Need to get device interface detail.
                trap {
                    foreach {devicepath deviceelement} \
                        [SetupDiGetDeviceInterfaceDetail \
                             $hdevinfo \
                             $interface_data \
                             $opts(matchdeviceelement)] \
                        break

                    if {$opts(deviceelement)} {
                        lappend item -deviceelement $deviceelement
                    }
                    if {$opts(devicepath)} {
                        lappend item -devicepath $devicepath
                    }
                } onerror {} {
                    if {! $opts(ignoreerrors)} {
                        error $errorResult $errorInfo $errorCode
                    }
                }
            }
            lappend result $item

            incr i
        }
    } onerror {TWAPI_WIN32 259} {
        # Fine, Just means no more items
    }

    return $result
}


# Return the guids associated with a device class set name. Note
# the latter is not unique so multiple guids may be associated.
proc twapi::device_setup_class_name_to_guids {name args} {
    array set opts [parseargs args {
        system.arg
    } -maxleftover 0 -nulldefault]

    return [twapi::SetupDiClassGuidsFromNameEx $name $opts(system)]
}

# Given a setup class guid, return the name of the setup class
interp alias {} twapi::device_setup_class_guid_to_name {} twapi::SetupDiClassNameFromGuidEx

# REturn the device instance id from a device element
interp alias {} twapi::get_device_element_instance_id {} twapi::SetupDiGetDeviceInstanceId


# Utility functions

proc twapi::_init_device_registry_code_maps {} {
    variable _device_registry_syms
    variable _device_registry_codes

    # Note this list is ordered based on the corresponding integer codes
    set _device_registry_code_syms {
        devicedesc hardwareid compatibleids unused0 service unused1
        unused2 class classguid driver configflags mfg friendlyname
        location physical capabilities ui upperfilters lowerfilters
        bustypeguid legacybustype busnumber enumerator security
        security devtype exclusive characteristics address ui device
        removal removal removal install location
    }

    set i 0
    foreach sym $_device_registry_code_syms {
        set _device_registry_codes($sym) $i
        incr i
    }
}

# Map a device registry property to a symbol
proc twapi::_device_registry_code_to_sym {code} {
    _init_device_registry_code_maps

    # Once we have initialized, redefine ourselves so we do not do so
    # every time. Note define at global ::twapi scope!
    proc ::twapi::_device_registry_code_to_sym {code} {
        variable _device_registry_code_syms
        if {$code >= [llength $_device_registry_code_syms]} {
            return $code
        } else {
            return [lindex $_device_registry_code_syms $code]
        }
    }
    # Call the redefined proc
    return [_device_registry_code_to_sym $code]
}

# Map a device registry property symbol to a numeric code
proc twapi::_device_registry_sym_to_code {sym} {
    _init_device_registry_code_maps

    # Once we have initialized, redefine ourselves so we do not do so
    # every time. Note define at global ::twapi scope!
    proc ::twapi::_device_registry_sym_to_code {sym} {
        variable _device_registry_codes
        # Return the value. If non-existent, an error will be raised
        if {[info exists _device_registry_codes($sym)]} {
            return $_device_registry_codes($sym)
        } elseif {[string is integer -strict $sym]} {
            return $sym
        } else {
            error "Unknown or unsupported device registry property symbol '$sym'"
        }
    }
    # Call the redefined proc
    return [_device_registry_sym_to_code $sym]
}

# Do a device ioctl, returning result as a binary
proc twapi::device_ioctl {h code args} {
    variable _ioctl_membuf;     # Memory buffer is reused so we do not allocate every time
    variable _ioctl_membuf_size

    array set opts [parseargs args {
        {inputbuffer.arg NULL}
        {inputcount.int 0}
    } -maxleftover 0]

    if {![info exists _ioctl_membuf]} {
        set _ioctl_membuf_size 128
        set _ioctl_membuf [malloc $_ioctl_membuf_size]
    }

    # Note on an exception error, the output buffer stays allocated.
    # That is not a bug.
    while {true} {
        trap {
            # IMPORTANT NOTE: the last parameter OVERLAPPED must be NULL
            # since device_ioctl is not reentrant (because of the use
            # of the "static" output buffer) and hence must not
            # be used in asynchronous operation where it might be called
            # before the previous call has finished.
            set outcount [DeviceIoControl $h $code $opts(inputbuffer) $opts(inputcount) $_ioctl_membuf $_ioctl_membuf_size NULL]
        } onerror {TWAPI_WIN32 122} {
            # Need to reallocate buffer. The seq below is such that
            # _ioctl_membuf is valid at all times, even when
            # there is a no memory exception.
            set newsize [expr {$_ioctl_membuf_size * 2}]
            set newbuf [malloc $newsize]
            # Now that we got a new buffer without an exception, set the
            # buffer variables
            set _ioctl_membuf $newbuf
            set _ioctl_membuf_size $newsize
            # Loop back to retry
            continue
        }

        # Fine, got what we wanted. Break out of the loop
        break
    }

    set bin [Twapi_ReadMemoryBinary $_ioctl_membuf 0 $outcount]

    # Do not hold on to cache memory is it is too big
    if {$_ioctl_membuf_size >= 1000} {
        free $_ioctl_membuf
        unset _ioctl_membuf
        set _ioctl_membuf_size 0
    }

    return $bin
}
