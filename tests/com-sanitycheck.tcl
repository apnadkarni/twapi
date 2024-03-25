#
# Some simple sanity checks for COM

proc twapi::_com_tests {{tests {ie word excel wmi tracker}}} {

    if {"ie" in $tests} {
        puts "Invoking Internet Explorer"
        set ie [comobj InternetExplorer.Application -enableaaa true]
        $ie Visible 1
        $ie Navigate http://www.google.com
        after 2000
        puts "Exiting Internet Explorer"
        $ie Quit
        $ie -destroy
        puts "Internet Explorer done."
        puts "------------------------------------------"
    }

    if {"word" in $tests} {
        puts "Invoking Word"
        set word [comobj Word.Application]
        set doc [$word -with Documents Add]
        $word Visible 1
        puts "Inserting text"
        $word -with {selection font} name "Courier New"
        $word -with {selection font} size 10.0
        $doc -with content text "Text in Courier 10 point"
        after 2000
        puts "Exiting Word"
        $doc -destroy
        $word Quit 0
        $word -destroy
        puts "Word done."
        
        puts "------------------------------------------"
    }

    if {"excel" in $tests} {
        puts "Invoking Excel"
        # This tests property sets with multiple parameters
        set xl [comobj Excel.Application]
        $xl -set Visible True

        $xl WindowState -4137;        # Test for enum params

        set workbooks [$xl Workbooks]
        set workbook [$workbooks Add]
        set sheets [$workbook Sheets]
        set sheet [$sheets Item 1]
        $sheet Activate
        set r [$sheet Range "A1:B2"]
        $r Value 10 "helloworld"
        after 2000
        $r Value 11 [string map {helloworld hellouniverse} [$r Value 11]]
        set vals [$r Value 10]
        if {$vals ne "hellouniverse hellouniverse hellouniverse hellouniverse"} {
            puts "EXcel mismatch"
        }

        # clean up
        $xl DisplayAlerts 0
        $xl Quit
        foreach obj {r sheet sheets workbook workbooks xl} {
            [set $obj] -destroy
        }

    }

    if {"wmi" in $tests} {
        puts "WMI BIOS test"
        puts [_get_bios_info]
        puts "WMI BIOS done."

        puts "------------------------------------------"
    
        puts "WMI direct property access test (get bios version)"
        set wmi [twapi::_wmi]
        $wmi -with {{ExecQuery "select * from Win32_BIOS"}} -iterate biosobj {
            puts "BIOS version: [$biosobj BiosVersion]"
            $biosobj -destroy
        }
        $wmi -destroy

        puts "------------------------------------------"
    }

    if {"tracker" in $tests} {
        puts " Starting process tracker. Type 'twapi::_stop_process_tracker' to stop it."
        twapi::_start_process_tracker
        vwait ::twapi::_stop_tracker
    }
}


#
proc twapi::_wmi_read_popups {} {
    set res {}
    set wmi [twapi::_wmi]
    set wql {select * from Win32_NTLogEvent where LogFile='System' and \
                 EventType='3'    and \
                 SourceName='Application Popup'}
    set svcs [$wmi ExecQuery $wql]

    # Iterate all records
    $svcs -iterate instance {
        set propSet [$instance Properties_]
        # only the property (object) 'Message' is of interest here
        set msgVal [[$propSet Item Message] Value]
        lappend res $msgVal
    }
    return $res
}

#
proc twapi::_wmi_read_popups_succint {} {
    set res [list ]
    set wmi [twapi::_wmi]
    $wmi -with {
        {ExecQuery "select * from Win32_NTLogEvent where LogFile='System' and EventType='3' and SourceName='Application Popup'"}
    } -iterate event {
        lappend res [$event Message]
    }
    return $res
}

# Returns a list of records returned by WMI. The name of each field in
# the record is in lower case to make it easier to extract without
# worrying about case.
proc twapi::_wmi_records {wmi_class} {
    set wmi [twapi::_wmi]
    set records [list ]
    $wmi -with {{ExecQuery "select * from $wmi_class"}} -iterate elem {
        set record {}
        set propset [$elem Properties_]
        $propset -iterate itemobj {
            # Note how we get the default property
            lappend record [string tolower [$itemobj Name]] [$itemobj -default]
            $itemobj -destroy
        }
        $elem -destroy
        $propset -destroy
        lappend records $record
    }
    $wmi -destroy
    return $records
}

#
proc twapi::_wmi_get_autostart_services {} {
    set res [list ]
    set wmi [twapi::_wmi]
    $wmi -with {
        {ExecQuery "select * from Win32_Service where StartMode='Auto'"}
    } -iterate svc {
        lappend res [$svc DisplayName]
    }
    return $res
}

proc twapi::_get_bios_info {} {
    set wmi [twapi::_wmi]
    set entries [list ]
    $wmi -with {{ExecQuery "select * from Win32_BIOS"}} -iterate elem {
        set propset [$elem Properties_]
        $propset -iterate itemobj {
            # Note how we get the default property
            lappend entries [$itemobj Name] [$itemobj -default]
            $itemobj -destroy
        }
        $elem -destroy
        $propset -destroy
    }
    $wmi -destroy
    return $entries
}

# Handler invoked when a process is started.  Will print exe name of process.
proc twapi::_process_start_handler {wmi_event args} {
    if {$wmi_event eq "OnObjectReady"} {
        # First arg is a IDispatch interface of the event object
        # Create a TWAPI COM object out of it
        set ifc [lindex $args 0]
        IUnknown_AddRef $ifc;   # Must hold ref before creating comobj
        set event_obj [comobj_idispatch $ifc]

        # Get and print the Name property
        puts "Process [$event_obj ProcessID] [$event_obj ProcessName] started at [clock format [large_system_time_to_secs [$event_obj TIME_CREATED]] -format {%x %X}]"

        # Get rid of the event object
        $event_obj -destroy
    }
}

# Call to begin tracking of processes.
proc twapi::_start_process_tracker {} {
    # Get local WMI root provider
    set ::twapi::_process_wmi [twapi::_wmi]

    # Create an WMI event sink
    set ::twapi::_process_event_sink [comobj wbemscripting.swbemsink]

    # Attach our handler to it
    set ::twapi::_process_event_sink_id [$::twapi::_process_event_sink -bind twapi::_process_start_handler]

    # Associate the sink with a query that polls every 1 sec for process
    # starts.
    set sink_ifc [$::twapi::_process_event_sink -interface]; # Does AddRef
    trap {
        $::twapi::_process_wmi ExecNotificationQueryAsync $sink_ifc "select * from Win32_ProcessStartTrace"
        # WMI will internally do a AddRef, so we can release our AddRef on sink_ifc
    } finally {
        IUnknown_Release $sink_ifc
    }
}

# Stop tracking of process starts
proc twapi::_stop_process_tracker {} {
    # Cancel event notifications
    $::twapi::_process_event_sink Cancel

    # Unbind our callback
    $::twapi::_process_event_sink -unbind $::twapi::_process_event_sink_id

    # Get rid of all objects
    $::twapi::_process_event_sink -destroy
    $::twapi::_process_wmi -destroy

    set ::twapi::_stop_tracker 1
    return
}


# Handler invoked when a service status changes.
proc twapi::_service_change_handler {wmi_event args} {
    if {$wmi_event eq "OnObjectReady"} {
        # First arg is a IDispatch interface of the event object
        # Create a TWAPI COM object out of it
        set ifc [lindex $args 0]
        IUnknown_AddRef $ifc;   # Needed before passing to comobj
        set event_obj [twapi::comobj_idispatch $ifc]

        puts "Previous: [$event_obj PreviousInstance]"
        #puts "Target: [$event_obj -with TargetInstance State]"

        # Get rid of the event object
        $event_obj -destroy
    }
}

# Call to begin tracking of service state
proc twapi::_start_service_tracker {} {
    # Get local WMI root provider
    set ::twapi::_service_wmi [twapi::_wmi]

    # Create an WMI event sink
    set ::twapi::_service_event_sink [twapi::comobj wbemscripting.swbemsink]

    # Attach our handler to it
    set ::twapi::_service_event_sink_id [$::twapi::_service_event_sink -bind twapi::_service_change_handler]

    # Associate the sink with a query that polls every 1 sec for service
    # starts.
    $::twapi::_service_wmi ExecNotificationQueryAsync [$::twapi::_service_event_sink -interface] "select * from __InstanceModificationEvent within 1 where TargetInstance ISA 'Win32_Service'"
}

# Stop tracking of services
proc twapi::_stop_service_tracker {} {
    # Cancel event notifications
    $::twapi::_service_event_sink Cancel

    # Unbind our callback
    $::twapi::_service_event_sink -unbind $::twapi::_service_event_sink_id

    # Get rid of all objects
    $::twapi::_service_event_sink -destroy
    $::twapi::_service_wmi -destroy
}
