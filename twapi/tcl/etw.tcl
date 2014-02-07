#
# Copyright (c) 2012-2014 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {
    # GUID's and event types for ETW.
    variable _etw_mof
    array set _etw_mof {
        provider_name "TwapiETWProvider"
        provider_guid "{B358E9D9-4D82-4A82-A129-BAC098C54746}"
        eventclass_name "TwapiETWEventClass"
        eventclass_guid "{D5B52E95-8447-40C1-B316-539894449B36}"
    }

    # So we don't pollute namespace with temp vars
    apply { {defs} {
        foreach {key val} $defs {
            proc [namespace current]::etw_twapi_$key {} "return $val"
        }
    }} [array get _etw_mof]

    # Cache of event definitions for parsing MOF  events. Nested dictionary
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


    #
    # These record definitions match the lists constructed in the ETW C code

    # Buffer header (EVENT_TRACE_LOGFILE)
    record etw_event_trace_logfile {
        logfile logger_name current_time buffers_read trace_logfile_header
        buffer_size filled kernel_trace
    }

    # TRACE_LOGFILE_HEADER
    record etw_trace_logfile_header {
        buffer_size
        version_major version_minor version_submajor version_subminor
        provider_version processor_count end_time timer_resolution 
        max_file_size logfile_mode buffers_written pointer_size events_lost
        cpu_mhz time_zone boot_time perf_frequency start_time reserved_flags
        buffers_lost
    }

    # TDH based event definitions

    record tdh_event { header buffer_context extended_data data }

    record tdh_event_header { flags event_property tid pid timestamp
        kernel_time user_time processor_time activity_id }

    record tdh_event_buffer_context { processor logger_id }

    record tdh_event_data { provider_guid event_guid descriptor
        decoder provider_name level_name channel_name keyword_names
        task_name opcode_name message localized_provider_name
        activity_id related_activity_id properties }

    record tdh_event_data_descriptor {id version channel level opcode task keywords}

    # Definitions for EVENT_TRACE_LOGFILE
    record tdh_buffer { logfile logger current_time buffers_read
        header buffer_size filled kernel_trace }

    record tdh_logfile_header { size major_version minor_version
        sub_version subminor_version provider_version processor_count
        end_time resolution max_file_size logfile_mode buffers_written
        pointer_size events_lost cpu_mhz timezone boot_time
        perf_frequency start_time reserved_flags buffers_lost }

    # MOF based event definitions
    record mof_event {header instance_id parent_instance_id parent_guid data}
    record mof_event_header {type level version tid pid timestamp guid
        kernel_time user_time processor_time}

    # Standard app visible event definitions
    record etw_event {id version channel level opcode task keywords
        timestamp tid pid 
        user_time kernel_time
        provider_guid provider_name event_guid event_name
        channel_name level_name opcode_name task_name keyword_names
        properties
        message
    }
}


proc twapi::etw_install_twapi_mof {} {
    variable _etw_mof
    
    # MOF definition for our ETW trace event. This is loaded into
    # the system WMI registry so event readers can decode our events
    #
    # Note all strings are NullTerminated and not Counted so embedded nulls
    # will not be handled correctly. The problem with using Counted strings
    # is that the MSDN docs are inconsistent as to whether the count
    # is number of *bytes* or number of *characters* and the existing tools
    # are similarly confused. We avoid this by choosing null terminated
    # strings despite the embedded nulls drawback.
    # TBD - revisit this and see if counted can always be treated as
    # bytes and not characters.
    set mof_template {
        #pragma namespace("\\\\.\\root\\wmi")

        [dynamic: ToInstance, Description("TWAPI ETW Provider"),
         Guid("@provider_guid")]
        class @provider_name : EventTrace
        {
        };

        [dynamic: ToInstance, Description("TWAPI ETW event class"): Amended,
         Guid("@eventclass_guid")]
        class @eventclass_name : @provider_name
        {
        };

        // NOTE: The EventTypeName is REQUIRED else the MS LogParser app
        // crashes (even though it should not)

        [dynamic: ToInstance, Description("TWAPI log message"): Amended,
         EventType(1), EventTypeName("Message")]
        class @eventclass_name_Message : @eventclass_name
        {
            [WmiDataId(1), Description("Log message"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Message;
        };

        [dynamic: ToInstance, Description("TWAPI variable trace"): Amended,
         EventType(2), EventTypeName("VariableTrace")]
        class @eventclass_name_VariableTrace : @eventclass_name
        {
            [WmiDataId(1), Description("Operation"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Operation;
            [WmiDataId(2), Description("Variable name"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Name;
            [WmiDataId(3), Description("Array index"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Index;
            [WmiDataId(4), Description("Value"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Value;
            [WmiDataId(5), Description("Context"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Context;
        };

        [dynamic: ToInstance, Description("TWAPI execution trace"): Amended,
         EventType(3), EventTypeName("ExecutionTrace")]
        class @eventclass_name_ExecutionTrace : @eventclass_name
        {
            [WmiDataId(1), Description("Operation"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Operation;
            [WmiDataId(2), Description("Executed command"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Command;
            [WmiDataId(3), Description("Status code"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Code;
            [WmiDataId(4), Description("Result"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Result;
            [WmiDataId(5), Description("Context"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Context;
        };

        [dynamic: ToInstance, Description("TWAPI command trace"): Amended,
         EventType(4), EventTypeName("CommandTrace")]
        class @eventclass_name_CommandTrace : @eventclass_name
        {
            [WmiDataId(1), Description("Operation"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Operation;
            [WmiDataId(2), Description("Old command name"): Amended, read, StringTermination("NullTerminated"), Format("w")] string OldName;
            [WmiDataId(3), Description("New command name"): Amended, read, StringTermination("NullTerminated"), Format("w")] string NewName;
            [WmiDataId(4), Description("Context"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Context;
        };
    }

    set mof [string map \
                 [list @provider_name $_etw_mof(provider_name) \
                      @provider_guid $_etw_mof(provider_guid) \
                      @eventclass_name $_etw_mof(eventclass_name) \
                      @eventclass_guid $_etw_mof(eventclass_guid) \
                     ] $mof_template]

    set mofc [twapi::IMofCompilerProxy new]
    twapi::trap {
        $mofc CompileBuffer $mof
    } finally {
        $mofc Release
    }
}

proc twapi::etw_uninstall_twapi_mof {} {
    variable _etw_mof

    set wmi [twapi::_wmi wmi]
    trap {
        set omof [$wmi Get $_etw_mof(provider_name)]
        $omof Delete_
    } finally {
        if {[info exists omof]} {
            $omof destroy
        }
        $wmi destroy
    }
}

proc twapi::etw_twapi_provider_register {} {
    variable _etw_mof
    return [twapi::RegisterTraceGuids $_etw_mof(provider_guid) $_etw_mof(eventclass_guid)]
}

proc twapi::etw_log_message {htrace message} {
    # Must match Message event type in MoF definition
    TraceEvent $htrace 1 0 [encoding convertto unicode "$message\0"]
    return
}

proc twapi::etw_variable_tracker {htrace name1 name2 op} {
    switch -exact -- $op {
        array -
        unset { set var "" }
        default {
            if {$name2 eq ""} {
                upvar 1 $name1 var
            } else {
                upvar 1 $name1($name2) var
            }
        }
    }

    if {[info level] > 1} {
        set context [info level -1]
    } else {
        set context ""
    }

    # Must match VariableTrace event type in MoF definition
    TraceEvent $htrace 2 0 \
        [encoding convertto unicode "$op\0$name1\0$name2\0$var\0"] \
        [_etw_make_limited_unicode $context]
}


proc twapi::etw_execution_tracker {htrace command args} {
    set op [lindex $args end]

    switch -exact -- $op {
        enter -
        enterstep {
            set code ""
            set result ""
        }
        leave -
        leavestep {
            lassign $args code result
        }
    }

    if {[info level] > 1} {
        set context [info level -1]
    } else {
        set context ""
    }

    # Must match Execution event type in MoF definition
    TraceEvent $htrace 3 0 \
        [encode convertto unicode "$op\0"] \
        [_etw_make_limited_unicode $command] \
        [encode convertto unicode "$code\0"] \
        [_etw_make_limited_unicode $result] \
        [_etw_make_limited_unicode $context]
}


proc twapi::etw_command_tracker {htrace oldname newname op} {
    if {[info level] > 1} {
        set context [info level -1]
    } else {
        set context ""
    }

    # Must match CommandTrace event type in MoF definition
    TraceEvent $htrace 4 0 \
        [encode converto unicode "$op\0$oldname\0$newname\0"] \
        [_etw_make_limited_unicode $context]
}

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
    # Maps event field type strings to enums to pass to the C code
    # 0 should be unmapped. Note some are duplicates because they
    # are the same format. Some are legacy formats not explicitly documented
    # in MSDN but found in the sample code.
    # Reference - Event Tracing MOF Qualifiers http://msdn.microsoft.com/en-us/library/windows/desktop/aa363800(v=vs.85).aspx
    set etw_fieldtypes {
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

    if {![dict exists $etw_fieldtypes $type]} {
        set fieldtype 0
    } else {
        set fieldtype [dict get $etw_fieldtypes $type]
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
            # GUID's can be multiple format, canonicalize for lsearch
            lappend guids [canonicalize_guid $arg]
        } else {
            lappend names $arg
        }
    }

    # Note there can be multiple versions sharing a single guid so
    # we cannot use the wmi_collect_classes "-first" option to stop the
    # search when one is found.

    set name_matcher [lambda* {names val} {
        ::tcl::mathop::>= [lsearch -exact -nocase $names $val] 0
    } :: $names]
    set guid_matcher [lambda* {guids val} {
        ::tcl::mathop::>= [lsearch -exact -nocase $guids $val] 0
    } :: $guids]

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

proc twapi::etw_open_file {path} {
# TBD - PROCESS_TRACE_MODE_RAW_TIMESTAMP
    variable _etw_open_traces

    set path [file normalize $path]

    set htrace [OpenTrace $path 0]
    set _etw_open_traces($htrace) $path
    return $htrace
}

proc twapi::etw_open_session {sessionname} {
# TBD - PROCESS_TRACE_MODE_RAW_TIMESTAMP
    variable _etw_open_traces

    set htrace [OpenTrace $sessionname 1]
    set _etw_open_traces($htrace) $sessionname
    return $htrace
}

proc twapi::etw_close_session {htrace} {
    variable _etw_open_traces

    if {[info exists _etw_open_traces($htrace)]} {
        CloseTrace $htrace
        unset _etw_open_traces($htrace)
    }
    return
}


proc twapi::etw_process_events {args} {
    array set opts [parseargs args {
        callback.arg
        start.arg
        end.arg
    } -nulldefault]

    if {[llength $args] == 0} {
        error "At least one trace handle must be specified."
    }

    return [ProcessTrace $args $opts(callback) $opts(start) $opts(end)]
}

proc twapi::etw_formatter_open {} {
    variable _etw_formatters

    if {[etw_force_mof] || ![twapi::min_os_version 6 0]} {
        # Need WMI MOF definitions
        set id mof[TwapiId]
        dict set _etw_formatters $id OSWBemServices [wmi_root -root wmi]
    } else {
        # Just a dummy if using a TDH based api
        set id tdh[TwapiId]
        # Nothing to set as yet but for consistency with MOF implementation
        dict set _etw_formatters $id {}
    }
    return $id
}

proc twapi::etw_formatter_close {formatter} {
    variable _etw_formatters
    if {[dict exists $_etw_formatters $formatter OSWBemServices]} {
        [dict get $_etw_formatters $formatter OSWBemServices] -destroy
    }

    dict unset _etw_formatters $formatter
    if {[dict size $_etw_formatters] == 0} {
        variable _etw_event_defs
        # No more formatters
        # Clear out event defs cache which can be quite large
        # Really only needed for mof but doesn't matter
        set _etw_event_defs {}
    }

    return
}

proc twapi::etw_format_events {formatter args} {
    variable _etw_formatters

    if {![dict exists $_etw_formatters $formatter]} {
        # We could actually just init ourselves but we want to force
        # consistency and caller to release wmi COM object
        badargs! "Invalid ETW formatter id \"$formatter\""
    }

    set events {}
    if {[dict exists $_etw_formatters $formatter OSWBemServices]} {
        set oswbemservices [dict get $_etw_formatters $formatter OSWBemServices]
        foreach {bufd rawevents} $args {
            lappend events [_etw_format_mof_events $oswbemservices $bufd $rawevents]
        }
    } else {
        foreach {bufd rawevents} $args {
            lappend events [_etw_format_tdh_events $bufd $rawevents]
        }
    }

    # TBD - is concat the best way to do this ?
    return [concat {*}$events]
}

proc twapi::_etw_format_mof_events {oswbemservices bufdesc events} {
    variable _etw_event_defs

    # TBD - it may be faster to special case NT kernel events as per
    # the structures defined in http://msdn.microsoft.com/en-us/library/windows/desktop/aa364083(v=vs.85).aspx
    # However, the MSDN warns that structures should not be created from
    # MOF classes as alignment restrictions might be different
    array set missing {}
    foreach event $events {
        set guid [mof_event_header guid [mof_event header $event]]
        if {! [dict exists $_etw_event_defs $guid]} {
            set missing($guid) ""
        }
    }

    if {[array size missing]} {
        etw_load_mof_event_classes $oswbemservices {*}[array names missing]
    }

    set bufhdr [etw_event_trace_logfile trace_logfile_header $bufdesc]
    set timer_resolution [etw_trace_logfile_header timer_resolution $bufhdr]
    set private_session [expr {0x800 & [etw_trace_logfile_header logfile_mode $bufhdr]}]
    set pointer_size [etw_trace_logfile_header pointer_size $bufhdr]

    set formatted_events {}
    foreach event $events {
        array set hdr [mof_event_header [mof_event header $event]]
        
        # Formatted event must match field sequence in etw_event record
        set formatted_event [list 0 $hdr(version) 0 $hdr(level) $hdr(type) 0 0 \
                                 $hdr(timestamp) $hdr(tid) $hdr(pid)]
        
        if {$private_session} {
            lappend formatted_event [expr {$hdr(processor_time) * $timer_resolution}] 0
        } else {
            lappend formatted_event [expr {$hdr(user_time) * $timer_resolution}] [expr {$hdr(kernel_time) * $timer_resolution}]
        }

        if {[dict exists $_etw_event_defs $hdr(guid) $hdr(version) -definitions $hdr(type)]} {
            set eventclass [dict get $_etw_event_defs $hdr(guid) $hdr(version) -classname]
            set mof [dict get $_etw_event_defs $hdr(guid) $hdr(version) -definitions $hdr(type)]
            set eventtypename [dict get $mof -eventtypename]
            set properties [Twapi_ParseEventMofData \
                                [mof_event data $event] \
                                [dict get $mof -fieldtypes] \
                                $pointer_size]
        } elseif {[dict exists $_etw_event_defs $hdr(guid) "" -definitions $hdr(type)]} {
            # If exact version not present, use one without
            # a version
            set eventclass [dict get $_etw_event_defs $hdr(guid) "" -classname]
            set mof [dict get $_etw_event_defs $hdr(guid) "" -definitions $hdr(type)]
            set eventtypename [dict get $mof -eventtypename]
            set properties [Twapi_ParseEventMofData \
                                [mof_event data $event] \
                                [dict get $mof -fieldtypes] \
                                $pointer_size]
        } else {
            # No definition. Create an entry so we know we already tried
            # looking this up and don't keep retrying later
            dict set _etw_event_defs $hdr(guid) {}

            # Nothing we can add to the event. Pass on with defaults
            set eventtypename $hdr(type)
            # Try to get at least the class name
            if {[dict exists $_etw_event_defs $hdr(guid) $hdr(version) -classname]} {
                set eventclass [dict get $_etw_event_defs $hdr(guid) $hdr(version) -classname]
            } elseif {[dict exists $_etw_event_defs $hdr(guid) "" -classname]} {
                set eventclass [dict get $_etw_event_defs $hdr(guid) "" -classname]
            } else {
                set eventclass ""
            }
            set properties [list mofdata [mof_event data $event]]
        }

        # eventclass -> provider_name
        # TBD - should we get the Provider qualifier from Mof as provider_name?
        # TBD - what should provider_guid be ?
        set provider_guid ""
        # mofformatteddata -> properties
        # TBD - construct dummy level
        set level_name ""
        lappend formatted_event $provider_guid $eventclass $hdr(guid) $eventclass "" $level_name $eventtypename "" "" $properties ""

        lappend formatted_events $formatted_event
    }

    return $formatted_events
}


twapi::proc* twapi::etw_dump_to_file {args} {
    package require twapi_wmi
    package require csv
} {
    array set opts [parseargs args {
        {output.arg stdout}
        {limit.int -1}
        {format.arg list {csv list}}
        {separator.arg ,}
    }]

    if {$opts(output) in [chan names]} {
        # Writing to a channel
        set outfd $opts(output)
        set do_close 0
    } else {
        if {[file exists $opts(output)]} {
            error "File $opts(output) already exists."
        }
        set outfd [open $opts(output) a]
        set do_close 1
    }

    set formatter [etw_formatter_open]
    trap {
        set varname ::twapi::_etw_dump_ctr[TwapiId]
        set $varname 0;         # Yes, set $varname, not set varname
        set htraces {}
        foreach arg $args {
            if {[file exists $arg]} {
                lappend htraces [etw_open_file $arg]
            } else {
                lappend htraces [etw_open_session $arg]
            }
        }
        # This is written using a callback to basically test the callback path
        set callback [list apply {
            {options outfd counter_varname max formatter bufd events}
            {
                array set opts $options
                foreach event [etw_format_events $formatter $bufd $events] {
                    if {$max >= 0 && [set $counter_varname] >= $max} { return -code break }
                    incr $counter_varname

                    array set fields [etw_event $event]
                    set message $fields(message)
                    if {$message ne ""} {
                        set params {}
                        foreach {propname propval} $fields(properties) {
                            lappend params $propval
                        }
                        catch {set message [format_message -fmtstring $message -params $params]}
                    }
                    set fmtdata $fields(properties)
                    if {[dict exists $fmtdata mofdata]} {
                        # Only show 32 bytes
                        binary scan [string range [dict get $fmtdata mofdata] 0 31] H* hex
                        dict set fmtdata mofdata [regsub -all (..) $hex {\1 }]
                    }
                    set fmtlist [list $fields(timestamp) $fields(pid) $fields(tid) $fields(provider_name) $fields(event_name) {*}$fmtdata]
                    if {$opts(format) eq "csv"} {
                        puts $outfd [csv::join $fmtlist $opts(separator)]
                    } else {
                        puts $outfd $fmtlist
                    }
                }
            }
        } [array get opts] $outfd $varname $opts(limit) $formatter]

        # Process the events using the callback
        etw_process_events -callback $callback {*}$htraces

    } finally {
        unset -nocomplain $varname
        foreach htrace $htraces {
            etw_close_session $htrace
        }
        if {$do_close} {
            close $outfd
        } else {
            flush $outfd
        }
        etw_formatter_close $formatter
    }
}

twapi::proc* twapi::etw_dump_to_list {args} {
    package require twapi_wmi
} {
    set htraces {}
    set formatter [etw_formatter_open]
    trap {
        set wmi [wmi_root -root wmi]
        foreach arg $args {
            if {[file exists $arg]} {
                lappend htraces [etw_open_file $arg]
            } else {
                lappend htraces [etw_open_session $arg]
            }
        }
        return [etw_format_events $formatter {*}[etw_process_events {*}$htraces]]
    } finally {
        foreach htrace $htraces {
            etw_close_session $htrace
        }
        etw_formatter_close $formatter
    }
}

proc twapi::etw_start_trace {session_name args} {
    
    # Specialized for kernel debugging - {bufferingmode {} 0x400}
    # Not supported until Win7 {noperprocessorbuffering {} 0x10000000}
    # Not clear what conditions it can be used {usekbytesforsize {} 0x2000}
    array set opts [parseargs args {
        sessionguid.arg
        logfile.arg
        buffersize.int
        minbuffers.int
        maxbuffers.int
        maxfilesize.int
        flushtimer.int
        enableflags.int
        {filemode.arg circular {sequential append rotate circular}}
        {clockresolution.arg qpc}
        {private.bool 0 0x800}
        {realtime.bool 0 0x100}
        {secure.bool 0 0x80}
        {privateinproc.bool 0 0x20800}
        {sequence.arg none {none local global}}
        {paged.bool 0 0x01000000}
        {preallocate.bool 0 0x20}
    } -maxleftover 0]

    set logfilemode 0

    switch -exact $opts(filemode) {
        sequential {
            if {[info exists opts(maxfilesize)]} {
                # 1 -> EVENT_TRACE_FILE_MODE_SEQUENTIAL 
                set logfilemode [expr {$logfilemode | 1}]
            } else {
                # 0 -> EVENT_TRACE_FILE_MODE_NONE
                # set logfilemode [expr {$logfilemode | 0}]
            }
        }
        circular {
            # 2 -> EVENT_TRACE_FILE_MODE_CIRCULAR
            set logfilemode [expr {$logfilemode | 2}]
            if {![info exists opts(maxfilesize)]} {
                set opts(maxfilesize) 1; # 1MB default
            }
        }
        rotate {
            if {$opts(private) || $opts(privateinproc)} {
                if {![min_os_version 6 2]} {
                    badargs! "Option -filemode must not be \"rotate\" for private traces"
                }
            }

            # 8 -> EVENT_TRACE_FILE_MODE_NEWFILE
            set logfilemode [expr {$logfilemode | 8}]
            if {![info exists opts(maxfilesize)]} {
                set opts(maxfilesize) 1; # 1MB default
            }
        }
        append {
            if {$opts(private) || $opts(privateinproc) || $opts(realtime)} {
                badargs! "Option -filemode must not be \"append\" for private or realtime traces"
            }
            # 4 -> EVENT_TRACE_FILE_MODE_APPEND
            # Not clear what to do about maxfilesize. Keep as is for now
            set logfilemode [expr {$logfilemode | 4}]
        }
    }

    if {![info exists opts(maxfilesize)]} {
        set opts(maxfilesize) 0
    }

    if {$opts(realtime) && ($opts(private) || $opts(privateinproc))} {
        badargs! "Option -realtime is incompatible with options -private and -privateinproc"
    }

    foreach opt {sessionguid logfile buffersize minbuffers maxbuffers flushtimer enableflags maxfilesize} {
        if {[info exists opts($opt)]} {
            lappend params -$opt $opts($opt)
        }
    }

    set logfilemode [expr {$logfilemode | [dict! {none 0 local 0x8000 global 0x4000} $opts(sequence)]}]

    set logfilemode [tcl::mathop::| $logfilemode $opts(realtime) $opts(private) $opts(privateinproc) $opts(secure) $opts(paged) $opts(preallocate)]

    lappend params -logfilemode $logfilemode

    if {$opts(filemode) eq "append" && $opts(clockresolution) ne "system"} {
        error "Option -clockresolution must be set to 'system' if -filemode is append"
    }

    if {($opts(filemode) eq "rotate" || $opts(preallocate)) &&
        $opts(maxfilesize) == 0} {
        error "Option -maxfilesize must also be specified with -preallocate or -filemodenewfile."
    }

    lappend params -clockresolution [dict! {qpc 1 system 2 cpucycle 3} $opts(clockresolution)]

    trap {
        return [StartTrace $session_name $params]
    } onerror {TWAPI_WIN32 5} {
        return -options [trapoptions] "Access denied. This may be because the process does not have permission to create the specified logfile or because it is not running under an account permitted to control ETW traces."
    }
}

proc twapi::etw_start_kernel_trace {events args} {
    
    set enableflags 0

    # Note sysconfig is a dummy event. It is always logged.
    set eventmap {
        process 0x00000001
        thread 0x00000002
        imageload 0x00000004
        diskio 0x00000100
        diskfileio 0x00000200
        pagefault 0x00001000
        hardfault 0x00002000
        tcpip 0x00010000
        registry 0x00020000
        dbgprint 0x00040000
        sysconfig 0x00000000
    }

    if {"diskfileio" in $events} {
        lappend events diskio;  # Required by diskfileio
    }

    if {[min_os_version 6]} {
        lappend eventmap {*}{
            processcounter 0x00000008
            contextswitch 0x00000010
            dpc 0x00000020
            interrupt 0x00000040
            systemcall 0x00000080
            diskioinit 0x00000400
            alpc 0x00100000
            splitio 0x00200000
            driver 0x00800000
            profile 0x01000000
            fileio 0x02000000
            fileioinit 0x04000000
        }

        if {"diskio" in $events} {
            lappend events diskioinit
        }
    }

    if {[min_os_version 6 1]} {
        lappend eventmap {*}{
            dispatcher 0x00000800
            virtualalloc 0x00004000
        }
    }

    if {[min_os_version 6 2]} {
        lappend eventmap {*}{
            vamap 0x00008000
        }
        if {"sysconfig" ni $events} {
            # EVENT_TRACE_FLAG_NO_SYSCONFIG 
            set enableflags [expr {$enableflags | 0x10000000}]
        }
    }

    foreach event $events {
        set enableflags [expr {$enableflags | [dict! $eventmap $event]}]
    }

    # Name "NT Kernel Logger" is hardcoded in Windows
    # GUID is 9e814aad-3204-11d2-9a82-006008a86939 but does not need to be
    # specified. Note kernel logger cannot use paged memory so 
    # -paged 0 is required
    return [etw_start_trace "NT Kernel Logger" -enableflags $enableflags {*}$args -paged 0]
}

proc twapi::etw_enable_trace {hsession guid enableflags level} {
    return [EnableTrace 1 $enableflags [dict* {critical 1 error 2 warning 3 information 4 verbose 5} $level] $guid $hsession]
}

proc twapi::etw_disable_trace {hsession guid enableflags level} {
    return [EnableTrace 0 $enableflags [dict* {critical 1 error 2 warning 3 information 4 verbose 5} $level] $guid $hsession]
}

proc twapi::etw_control_trace {action args} {

    set action [dict get {
        query  0
        stop   1
        update 2
        flush  3
    } $action]

    array set opts [parseargs args {
        {sessionhandle.arg 0}
        sessionname.arg
        sessionguid.arg
        logfile.arg
        maxbuffers.int
        flushtimer.int
        enableflags.int
        realtime.bool
    } -maxleftover 0]

    set params {}

    if {(![info exists opts(sessionhandle)]) &&
        (![info exists opts(sessionname)])} {
        error "One of -sessionhandle or -sessionname must be specified."
    }

    if {[info exists opts(realtime)]} {
        if {$opts(realtime)} {
            lappend params -logfilemode 0x100
        } else {
            lappend params -logfilemode 0
        }
    }
        
    foreach opt {sessionguid sessionname} {
        if {[info exists opts($opt)]} {
            lappend params -$opt $opts($opt)
        }
    }

    if {$action eq "update"} {
        foreach opt {logfile flushtimer enableflags maxbuffers} {
            if {[info exists opts($opt)]} {
                lappend params -$opt $opts($opt)
            }
        }
    }

    return [ControlTrace $action $opts(sessionhandle) $params]
}

interp alias {} twapi::etw_stop_trace {} twapi::etw_control_trace stop
interp alias {} twapi::etw_flush_trace {} twapi::etw_control_trace flush
interp alias {} twapi::etw_query_trace {} twapi::etw_control_trace query
interp alias {} twapi::etw_update_trace {} twapi::etw_control_trace update

proc twapi::etw_query_trace {args} {
    set d [etw_control_trace query {*}$args]
    set cr [lindex  {{} qpc system cpucycle} [dict get $d -clockresolution]]
    if {$cr ne ""} {
        dict set d -clockresolution $cr
    }

    #TBD - check whether -maxfilesize needs to be massaged

    return $d
}

#
# Helper functions
#


# Return binary unicode with truncation if necessary
proc twapi::_etw_make_limited_unicode {s {max 80}} {
    if {[string length $s] > $max} {
        set s "[string range $s 0 $max-3]..."
    }
    return [encoding convertto unicode "$s\0"]
}

# Used for development/debug to see what all types are in use
proc twapi::_etw_get_types {} {
    dict for {g gval} $::twapi::_etw_event_defs {
        dict for {ver verval} $gval {
            dict for {eventtype eval} [dict get $verval -definitions] {
                dict for {id idval} [dict get $eval -fields] {
                    dict set types [dict get $idval -type] [dict get $verval -classname] $eventtype $id
                }
            }
        }
    }
    return $types
}

