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
    variable _etw_typeenums
    # 0 should be unmapped. Note some are duplicates because they
    # are the same format. Some are legacy formats not explicitly documented
    # in MSDN but found in the sample code.
    array set _etw_typeenums {
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
    }
}


proc twapi::etw_register_mof {} {
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

proc twapi::etw_register {} {
    variable _etw_provider_uuid
    variable _etw_event_class_uuid

    twapi::RegisterTraceGuids $_etw_provider_uuid $_etw_event_class_uuid
}

interp alias {} twapi::etw_unregister {} twapi::UnregisterTraceGuids

interp alias {} twapi::etw_trace {} twapi::TraceEvent


proc twapi::etw_find_event_classes {oswbemservices guid} {
    # Return all classes where the GUID matches
    # Note there can be multiple versions sharing a single guid so
    # we cannot use the "-first" option to stop the search when
    # one is found.
    return [wmi_collect_classes $oswbemservices -ancestor EventTrace -matchqualifiers [list Guid [list twapi::IsEqualGUID $guid]]]
}

proc twapi::etw_parse_event_class {ocls} {
    # Returns a dict 
    # First level key - event type (integer)
    # Second level keys:
    #   eventtypename - string (may be empty)
    #   eventtype - same as first level key
    #   record - list of field types indexed by wmiid
    # The record is an ordered list of enums that describe
    # the type of each field.

    set result [dict create]

    # Iterate over the subclasses, collecting the event metadata
    # Create a forward only enumerator for efficiency
    # wbemFlagUseAmendedQualifiers|wbemFlagReturnImmediately|wbemFlagForwardOnly
    # wbemQueryFlagsShallow
    # -> 0x20031
    $ocls -with {{SubClasses_ 0x20031}} -iterate osub {
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
            set record [dict create]
            $osub -with Properties_ -iterate oprop {
                set quals [$oprop Qualifiers_]
                # Event fields will have a WmiDataId qualifier
                if {![catch {$quals -with {{Item WmiDataId}} Value} wmidataid]} {
                    # Yep this is a field, figure out its type
                    set type [_etw_decipher_event_field_type $oprop $quals]
                    dict set record $wmidataid $type
                }
                $quals destroy
            }
                    
            # Process the records to put the fields in order based on
            # their wmidataid. If any info is missing or inconsistent
            # we will mark the whole event type class has undecodable.
            # Ids begin from 1.
            set typeenums {}
            for {set id 1} {$id <= [dict size $record]} {incr id} {
                if {![dict exists $record $id]} {
                    # Discard all type info - missing type info
                    debuglog "Missing id $id for event type(s) $event_types for  EventTrace Mof Class [$ocls -with {{SystemProperties_} {Item __CLASS}} Value]"
                    set typeenums {}
                    break;
                }
                lappend typeenums [dict get $record $id typeenum]
            }

            foreach event_type $event_types event_type_name $event_type_names {
                dict set result $event_type [dict create eventtype $event_type eventtypename $event_type_name record $record typeenums $typeenums]
            }
        }
    }

    return $result
}

# Deciphers an event record field type
proc twapi::_etw_decipher_event_field_type {oprop oquals} {
    variable _etw_typeenums

    # On any errors, we will set type to unknown or unsupported
    set type unknown
    set quals(extension)  "";   # Hint for formatting for display

    if {! [$oprop -get IsArray]} {
        # Cannot handle arrays yet - TBD

        # If the Pointer qualifier exists, ignore everything else
        if {![catch {
            $oquals -with {{Item Pointer}} Value
        }]} {
            # Actual value does not matter
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
                }
            }
        }
    }

    if {![info exists _etw_typeenums($type)]} {
        set typeenum 0
    } else {
        set typeenum $_etw_typeenums($type)
    }

    return [dict create type $type typeenum $typeenum extension $quals(extension)]
}

proc twapi::etw_get_all_event_classes {oswbemservices} {
    return [twapi::wmi_collect_classes $wmi -ancestor EventTrace -matchqualifiers [list Guid ::twapi::true]]
}

