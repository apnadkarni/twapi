#
# Copyright (c) 2012 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {
    # GUID's and event types for ETW.
    variable _etw_provider_uuid "{B358E9D9-4D82-4A82-A129-BAC098C54746}"
    variable _etw_event_class_uuid "{D5B52E95-8447-40C1-B316-539894449B36}"
    
    # Maps event field type strings to enums to pass to the C code
    variable _etw_fieldtypes
    # 0 should be unmapped. Note some are duplicates because they
    # are the same format. Some are legacy formats not explicitly documented
    # in MSDN but found in the sample code.
    array set _etw_fieldtypes {
        string  1
        stringnullterminated 1
        wstring 2
        wstringnullterminated 2
        stringcounted 3
        stringreversecounted 4
        wstringcounted 5
        wstringreversecounted 6
        boolean 7
        sint8 8
        uint8 9
        csint8 10
        cuint8 11
        sint16 12
        uint16 13
        uint32 14
        sint32 15
        sint64 16
        uint64 17
        xsint16 18
        xuint16 19
        xsint32 20
        xuint32 21
        xsint64 22
        xuint64 23
        real32 24
        real64 25
        object 26
        char16 27
        uint8guid 28
        objectguid 29
        objectipaddrv4 30
        uint32ipaddr 30
        objectipaddr 30
        objectipaddrv6 31
        objectvariant 32
        objectsid 33
        uint64wmitime 34
        objectwmitime 35
        uint16port 38
        objectport 39
        datetime 40
        stringnotcounted 41
        wstringnotcounted 42
        pointer 43
        sizet   43
    }

    # Cache of event definitions for parsing event. Nested dictionary
    # with the following structure (uppercase keys are variables,
    # lower case are constant/tokens, "->" is nested dict, "-" is scalar):
    #  EVENTCLASSGUID ->
    #    classname - name of the class
    #    definitions ->
    #      VERSION ->
    #        EVENTTYPE ->
    #          eventtype - same as EVENTTYPE
    #          eventtypename - name / description for the event type
    #          fieldtypes - ordered list of field types for that event
    #          fields ->
    #            FIELDINDEX ->
    #              type - the field type in string format
    #              fieldtype - the corresponding field type numeric value
    #              extension - the MoF extension qualifier for the field
    #
    # The cache assumes that MOF event definitions are globally identical
    # (ie. same on local and remote systems)
    variable _etw_event_defs
    set _etw_event_defs [dict create]

    # Keeps track of open trace handles
    variable _etw_open_traces
    array set _etw_open_traces {}
}


proc twapi::etw_install_mof {} {
    variable _etw_provider_uuid
    variable _etw_event_class_uuid
    
    # MOF definition for our ETW trace event. This is loaded into
    # the system WMI registry so event readers can decode our events
    set mof [format {
        #pragma namespace("\\\\.\\root\\wmi")

        [dynamic: ToInstance, Description("Tcl Windows API ETW Provider"),
         Guid("%s")]
        class TwapiETWProvider : EventTrace
        {
        };

        [dynamic: ToInstance, Description("Tcl Windows API ETW event class"): Amended,
         Guid("%s")]
        class TwapiETWEventClass : TwapiETWProvider
        {
        };

        // NOTE: The EventTypeName is REQUIRED else the MS LogParser app
        // crashes (even though it should not)
        [dynamic: ToInstance, Description("Tcl Windows API child event type class used to describe the standard event."): Amended,
         EventType(10), EventTypeName("Twapi Trace")]
        class TwapiETWEventClass_TwapiETWEventTypeClass : TwapiETWEventClass
        {
            [WmiDataId(1), Description("Log message"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Message;
        };
    } $_etw_provider_uuid $_etw_event_class_uuid]

    set mofc [twapi::IMofCompilerProxy new]
    twapi::trap {
        $mofc CompileBuffer $mof
    } finally {
        $mofc Release
    }
}

proc twapi::etw_register_provider {} {
    variable _etw_provider_uuid
    variable _etw_event_class_uuid

    twapi::RegisterTraceGuids $_etw_provider_uuid $_etw_event_class_uuid
}

interp alias {} twapi::etw_unregister_provider {} twapi::UnregisterTraceGuids

interp alias {} twapi::etw_trace {} twapi::TraceEvent

proc twapi::etw_parse_mof_event_class {ocls} {
    # Returns a dict 
    # First level key - event type (integer)
    # See description of _etw_event_defs for rest of the structure

    set result [dict create]

    # Iterate over the subclasses, collecting the event metadata
    # Create a forward only enumerator for efficiency
    # wbemFlagUseAmendedQualifiers|wbemFlagReturnImmediately|wbemFlagForwardOnly
    # wbemQueryFlagsShallow
    # -> 0x20031
    $ocls -with {{SubClasses_ 0x20031}} -iterate -cleanup osub {
        # The subclass must have the eventtype property
        # We fetch as a raw value so we can tell the
        # original type
        if {![catch {
            $osub -with {
                Qualifiers_
                {Item EventType}
            } -invoke Value 2 {} -raw 1
        } event_types]} {

            # event_types is a raw value with a type descriptor as elem 0
            if {[variant_type $event_types] & 0x2000} {
                # It is VT_ARRAY so value is already a list
                set event_types [variant_value $event_types]
            } else {
                set event_types [list [variant_value $event_types]]
            }

            set event_type_names {}
            catch {
                set event_type_names [$osub -with {
                    Qualifiers_
                    {Item EventTypeName}
                } -invoke Value 2 {} -raw 1]
                # event_type_names is a raw value with a type descriptor as elem 0
                # It is IMPORTANT to check this else we cannot distinguish
                # between a array (list) and a string with spaces
                if {[variant_type $event_type_names] & 0x2000} {
                    # It is VT_ARRAY so value is already a list
                    set event_type_names [variant_value $event_type_names]
                } else {
                    # Scalar value. Make into a list
                    set event_type_names [list [variant_value $event_type_names]]
                }
            }

            # The subclass has a EventType property. Pick up the
            # field definitions.
            set fields [dict create]
            $osub -with Properties_ -iterate -cleanup oprop {
                set quals [$oprop Qualifiers_]
                # Event fields will have a WmiDataId qualifier
                if {![catch {$quals -with {{Item WmiDataId}} Value} wmidataid]} {
                    # Yep this is a field, figure out its type
                    set type [_etw_decipher_mof_event_field_type $oprop $quals]
                    dict set type -fieldname [$oprop -get Name]
                    dict set fields $wmidataid $type
                }
                $quals destroy
            }
                    
            # Process the records to put the fields in order based on
            # their wmidataid. If any info is missing or inconsistent
            # we will mark the whole event type class has undecodable.
            # Ids begin from 1.
            set fieldtypes {}
            for {set id 1} {$id <= [dict size $fields]} {incr id} {
                if {![dict exists $fields $id]} {
                    # Discard all type info - missing type info
                    debuglog "Missing id $id for event type(s) $event_types for  EventTrace Mof Class [$ocls -with {{SystemProperties_} {Item __CLASS}} Value]"
                    set fieldtypes {}
                    break;
                }
                lappend fieldtypes [dict get $fields $id -fieldname] [dict get $fields $id -fieldtype]
            }

            foreach event_type $event_types event_type_name $event_type_names {
                dict set result -definitions $event_type [dict create -eventtype $event_type -eventtypename $event_type_name -fields $fields -fieldtypes $fieldtypes]
            }
        }
    }

    if {[dict size $result] == 0} {
        return {}
    } else {
        dict set result -classname [$ocls -with {SystemProperties_ {Item __CLASS}} Value]
        return $result
    }
}

# Deciphers an event  field type
proc twapi::_etw_decipher_mof_event_field_type {oprop oquals} {
    variable _etw_fieldtypes

    # On any errors, we will set type to unknown or unsupported
    set type unknown
    set quals(extension)  "";   # Hint for formatting for display

    if {![catch {
        $oquals -with {{Item Pointer}} Value
    }]} {
        # Actual value does not matter
        # If the Pointer qualifier exists, ignore everything else
        set type pointer
    } elseif {![catch {
        $oquals -with {{Item PointerType}} Value
    }]} {
        # Actual value does not matter
        # Some apps mistakenly use PointerType instead of Pointer
        set type pointer
    } else {
        catch {
            set type [string tolower [$oquals -with {{Item CIMTYPE}} Value]]

            # The following qualifiers may or may not exist
            # TBD - not all may be required to be retrieved
            # NOTE: MSDN says some qualifiers are case sensitive!
            foreach qual {BitMap BitValues Extension Format Pointer StringTermination ValueMap Values ValueType XMLFragment} {
                # catch in case it does not exist
                set lqual [string tolower $qual]
                set quals($lqual) ""
                catch {
                    set quals($lqual) [$oquals -with [list [list Item $qual]] Value]
                }
            }
            set type [string tolower "$quals(format)${type}$quals(stringtermination)"]
            set quals(extension) [string tolower $quals(extension)]
            # Not all extensions affect how the event field is extracted
            # e.g. the noprint value
            if {$quals(extension) in {ipaddr ipaddrv4 ipaddrv6 port variant wmitime guid sid}} {
                append type $quals(extension)
            } elseif {$quals(extension) eq "sizet"} {
                set type sizet
            }
        }
    }

    # Cannot handle arrays yet - TBD
    if {[$oprop -get IsArray]} {
        set type "arrayof$type"
    }

    if {![info exists _etw_fieldtypes($type)]} {
        set fieldtype 0
    } else {
        set fieldtype $_etw_fieldtypes($type)
    }

    return [dict create -type $type -fieldtype $fieldtype -extension $quals(extension)]
}

proc twapi::etw_find_mof_event_classes {oswbemservices args} {
    # Return all classes where a GUID or name matches

    # To avoid iterating the tree multiple times, separate out the guids
    # and the names and use separator comparators

    set guids {}
    set names {}

    foreach arg $args {
        if {[Twapi_IsValidGUID $arg]} {
            lappend guids $arg
        } else {
            lappend names $arg
        }
    }

    # Note there can be multiple versions sharing a single guid so
    # we cannot use the wmi_collect_classes "-first" option to stop the search when
    # one is found.

    set name_matcher [list apply {{names val} {::tcl::mathop::>= [lsearch -exact -nocase $names $val] 0}} $names]
    # Note GUIDS can be specified in multiple format so we cannot just lsearch
    set guid_matcher [list apply {{guids val} {
        foreach guid $guids {
            if {[twapi::IsEqualGUID $guid $val]} { return 1 }
        }
        return 0
    }} $guids]

    set named_classes {}
    if {[llength $names]} {
        foreach name $names {
            catch {lappend named_classes [$oswbemservices Get $name]}
        }
    }

    if {[llength $guids]} {
        set guid_classes [wmi_collect_classes $oswbemservices -ancestor EventTrace -matchqualifiers [list Guid $guid_matcher]]
    } else {
        set guid_classes {}
    }

    return [concat $guid_classes $named_classes]
}

proc twapi::etw_get_all_mof_event_classes {oswbemservices} {
    return [twapi::wmi_collect_classes $oswbemservices -ancestor EventTrace -matchqualifiers [list Guid ::twapi::true]]
}

proc twapi::etw_load_mof_event_class_obj {oswbemservices ocls} {
    variable _etw_event_defs
    set quals [$ocls Qualifiers_]
    trap {
        set guid [$quals -with {{Item Guid}} Value]
        set vers ""
        catch {set vers [$quals -with {{Item EventVersion}} Value]}
        set def [etw_parse_mof_event_class $ocls]
        # Class may be a provider, not a event class in which case
        # def will be empty
        if {[dict size $def]} {
            dict set _etw_event_defs [canonicalize_guid $guid] $vers $def
        }
    } finally {
        $quals destroy
    }
}

proc twapi::etw_load_mof_event_classes {oswbemservices args} {
    if {[llength $args] == 0} {
        set oclasses [etw_get_all_mof_event_classes $oswbemservices]
    } else {
        set oclasses [etw_find_mof_event_classes $oswbemservices {*}$args]
    }

    foreach ocls $oclasses {
        trap {
            etw_load_mof_event_class_obj $oswbemservices $ocls
        } finally {
            $ocls destroy
        }
    }
}

proc twapi::etw_open_trace {path args} {
    variable _etw_open_traces

    array set opts [parseargs args {
        realtime
    } -maxleftover 0]

    if {! $opts(realtime)} {
        set path [file normalize $path]
    }

    set htrace [OpenTrace $path $opts(realtime)]
    set _etw_open_traces($htrace) $path
    return $htrace
}

proc twapi::etw_close_trace {htrace} {
    variable _etw_open_traces

    if {[info exists _etw_open_traces($htrace)]} {
        CloseTrace $htrace
        unset _etw_open_traces($htrace)
    }
    return
}


proc twapi::etw_process_events {htrace args} {
    array set opts [parseargs args {
        callback.arg
        start.arg
        end.arg
    } -maxleftover 0 -nulldefault]

    return [ProcessTrace $htrace $opts(callback) $opts(start) $opts(end)]
}

proc twapi::etw_format_events {oswbemservices bufdesc events} {
    variable _etw_event_defs
    array set missing {}
    foreach event $events {
        if {! [dict exists $_etw_event_defs [dict get $event -guid]]} {
            set missing([dict get $event -guid]) 1
        }
    }

    if {[array size missing]} {
parray missing
        etw_load_mof_event_classes $oswbemservices {*}[array names missing]
    }

    set formatted {}
    foreach event $events {
        set guid [dict get $event -guid]
        set vers [dict get $event -version]
        set type [dict get $event -eventtype]
        if {[dict exists $_etw_event_defs $guid $vers -definitions $type]} {
            set mof [dict get $_etw_event_defs $guid $vers -definitions $type]
            set eventclass [dict get $_etw_event_defs $guid $vers -classname]
        } elseif {[dict exists $_etw_event_defs $guid "" -definitions $type]} {
            # If exact version not present, use one without
            # a version
            set mof [dict get $_etw_event_defs $guid "" -definitions $type]
            set eventclass [dict get $_etw_event_defs $guid "" -classname]
        } else {
            # No definition.
            # Nothing we can add to the event. Pass on with defaults
            dict set event -eventtypename [dict get $event -eventtype]
            # Try to get at least the class name
            if {[dict exists $_etw_event_defs $guid $vers -classname]} {
                dict set event -classname [dict get $_etw_event_defs $guid $vers -classname]
            } elseif {[dict exists $_etw_event_defs $guid "" -classname]} {
                dict set event -classname [dict get $_etw_event_defs $guid "" -classname]
            } else {
                dict set event -classname ""
            }
            lappend formatted $event
            continue
        }

        dict set event -eventtypename [dict get $mof -eventtypename]
        dict set event -classname $eventclass
        dict set event -mofformatteddata [Twapi_ParseEventMofData [dict get $event -mofdata] [dict get $mof -fieldtypes] [dict get $bufdesc -hdr_pointersize]]
        lappend formatted $event
    }

    return $formatted
}


proc twapi::etw_dump_file {path args} {
    array set opts [parseargs args {
        {outfd.arg stdout}
        {limit.int -1}
    } -maxleftover 0]

    trap {
        set wmi [twapi::_wmi wmi]
        set htrace [etw_open_trace $path]
        set varname ::twapi::_etw_dump_ctr[TwapiId]
        set $varname 0;         # Yes, set $varname, not set varname
        etw_process_events $htrace -callback [list apply {
            {fd counter_varname max wmi bufd events}
            {
                foreach event [etw_format_events $wmi $bufd $events] {
                    if {$max >= 0 && [set $counter_varname] >= $max} { return -code break }
                    incr $counter_varname
                    if {[dict exists $event -mofformatteddata]} {
                        set fmtdata [dict get $event -mofformatteddata]
                    } else {
                        binary scan [string range [dict get $event -mofdata] 0 31] H* hex
                        set fmtdata [dict create MofData [regsub -all (..) $hex {\1 }]]
                    }
                    puts $fd "[dict get $event -timestamp] [dict get $event -classname]/[dict get $event -eventtypename] $fmtdata"
                }
            }
        } $opts(outfd) $varname $opts(limit) $wmi]
    } finally {
        unset -nocomplain $varname
        if {[info exists htrace]} {etw_close_trace $htrace}
        if {[info exists wmi]} {$wmi destroy}
        flush $opts(outfd)
    }
}
