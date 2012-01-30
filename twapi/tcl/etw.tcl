#
# Copyright (c) 2012 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {
    # GUID's and event types for ETW.
    variable _etw_provider_uuid "{B358E9D9-4D82-4A82-A129-BAC098C54746}"
    variable _etw_event_class_uuid "{D5B52E95-8447-40C1-B316-539894449B36}"
    
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
            [WmiDataId(1), Description("Message"): Amended, read, StringTermination("NullTerminated"), Format("w")] string Message;
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



