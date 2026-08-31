#
# Copyright (c) 2012-2026, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Event log handling for Vista and later

package require twapi_synch

namespace eval twapi {
    variable _evt;              # See _evt_init

    # System event fields in order returned by _evt_event_decode_system_fields
    twapi::record evt_system_properties  {
        -providername -providerguid -eventid -qualifiers -level -task
        -opcode -keywordmask -timecreated -eventrecordid -activityid
        -relatedactivityid -pid -tid -channel
        -computer -sid -version
    }
    interp alias {} ::twapi::evt_system_fields {} ::twapi::evt_system_properties

    proc _evt_init {} {
        variable _evt

        # Various structures that we maintain / cache for efficiency as they
        # are commonly used are kept in the _evt array with the following keys:

        # system_render_context_handle - is the handle to a rendering
        #    context for the system portion of an event
        set _evt(system_render_context_handle) [evt_render_context_system]

        # user_render_context_handle - is the handle to a rendering
        #    context for the user data portion of an event
        set _evt(user_render_context_handle) [evt_render_context_userdata]

        # render_buffer - is NULL or holds a pointer to the buffer used to
        #    retrieve values so does not have to be reallocated every time.
        set _evt(render_buffer) NULL

        # publisher_handles - caches publisher names to their meta information.
        #    This is a dictionary indexed with nested keys -
        #     publisher, session, lcid. TBD - need a mechanism to clear ?
        set _evt(publisher_handles) [dict create]

        # -levelname - dict of publisher name / level number to level names
        set _evt(-levelname) {}

        # -taskname - dict of publisher name / task number to task name
        set _evt(-taskname) {}

        # -opcodename - dict of publisher name / opcode number to opcode name
        set _evt(-opcodename) {}

        # No-op the proc once init is done
        proc _evt_init {} {}
    }
}

proc twapi::evt_close {args} {
    EvtClose {*}$args
}

proc twapi::evt_bookmark_render {hbm} {
    # 2 -> EvtRenderBookmark
    return [Twapi_EvtRenderUnicode NULL $hbm 2]
}

proc twapi::evt_event_xml {hevt} {
    # 1 -> EvtRenderEventXml
    return [Twapi_EvtRenderUnicode NULL $hevt 1]
}

proc twapi::evt_event_render {hevt hctx} {
    set hbuf [Twapi_EvtRenderValues $hctx $hevt NULL]
    try {
        return [Twapi_ExtractEVT_RENDER_VALUES $hbuf]
    } finally {
        evt_free_EVT_RENDER_VALUES $hbuf
    }
}

proc twapi::evt_render_context_xpaths {xpaths} {
    return [EvtCreateRenderContext $xpaths 0]
}

proc twapi::evt_render_context_system {} {
    return [EvtCreateRenderContext {} 1]
}

proc twapi::evt_render_context_userdata {} {
    return [EvtCreateRenderContext {} 2]
}

proc twapi::evt_event_logpath {hevt} {
    return [EvtGetEventInfo $hevt 1]
}

twapi::proc* twapi::_evt_event_decode_system_fields {hevt} {
    _evt_init
} {
    variable _evt
    set _evt(render_buffer) [Twapi_EvtRenderValues $_evt(system_render_context_handle) $hevt $_evt(render_buffer)]
    set rec [Twapi_ExtractEVT_RENDER_VALUES $_evt(render_buffer)]
    return [evt_system_properties set $rec \
                -providername [atomize [evt_system_properties -providername $rec]] \
                -providerguid [atomize [evt_system_properties -providerguid $rec]] \
                -channel [atomize [evt_system_properties -channel $rec]] \
                -computer [atomize [evt_system_properties -computer $rec]]]
}

# TBD - document. Returns a list of user data values
twapi::proc* twapi::evt_event_decode_userdata {hevt} {
    _evt_init
} {
    variable _evt
    set _evt(render_buffer) [Twapi_EvtRenderValues $_evt(user_render_context_handle) $hevt $_evt(render_buffer)]
    return [Twapi_ExtractEVT_RENDER_VALUES $_evt(render_buffer)]
}

# TBD - document
# Where is this used?
proc twapi::evt_free_EVT_RENDER_VALUES {p} {
    evt_free $p
}

# TBD - test
proc twapi::evt_publisher_install {manifest resource_file message_file} {
    set wevutil [auto_execok wevtutil]
    if {[get_process_elevation] ne "full"} {
        set params "im"
        append params " \"[file nativename [file normalize $manifest]]\""
        append params " \"/rf:[file nativename [file normalize $resource_file]]\""
        append params " \"/mf:[file nativename [file normalize $message_file]]\""
        set wevutil [lindex $wevutil 0]
        shell_execute -verb runas -show hide -path $wevutil -params $params
        return
    }
    exec {*}$wevutil im \
        [file nativename [file normalize $manifest]] \
        "/rf:[file nativename [file normalize $resource_file]]" \
        "/mf:[file nativename [file normalize $message_file]]"
}

# TBD - test
proc twapi::evt_publisher_uninstall {manifest} {
    set wevutil [auto_execok wevtutil]
    if {[get_process_elevation] ne "full"} {
        set params "um"
        append params " \"[file nativename [file normalize $manifest]]\""
        set wevutil [lindex $wevutil 0]
        shell_execute -verb runas -show hide -path $wevutil -params $params
        return
    }
    exec {*}$wevutil um [file nativename [file normalize $manifest]]
}

# TBD - test
proc twapi::evt_twapi_install {} {
    set path [get_twapi_dll_path]
    if {$path eq ""} {
        # Assume statically linked
        set path [info nameofexecutable]
    }
    evt_publisher_install [file join [get_twapi_script_dir] twapi_events.man] $path $path
}

# TBD - test
proc twapi::evt_twapi_uninstall {} {
    evt_publisher_uninstall [file join [get_twapi_script_dir] twapi_events.man]
}

proc twapi::_evt_parse_query_options {argvvar args} {
    upvar 1 $argvvar argv
    parseargs argv {
        channel.arg
        logfile.arg
        {query.arg {}}
        {ignorequeryerrors 0 0x1000}
    } -setvars {*}$args

    if {[info exists channel]} {
        if {[info exists logfile]} {
            error "At most one of -channel and -logfile may be specified."
        }
        set source $channel
        set flags 1
    } elseif {[info exists logfile]} {
        set source $logfile
        set flags 2
    } elseif {$query ne ""} {
        set source ""
        set flags 1
    } else {
        set source Application
        set flags 1
    }
    return [list $source $query [tcl::mathop::| $flags $ignorequeryerrors]]
}

proc twapi::_evt_native_path {path} {
    # Do not want to rely on [file normalize] returning "" for ""
    if {$path eq ""} {
        return ""
    } else {
        return [file nativename [file normalize $path]]
    }
}

proc twapi::_evt_dump {args} {
    parseargs args {
        {outfd.arg stdout}
        count.int
    } -ignoreunknown -setvars


    set osess [EvtSession new]
    try {
        set oreader [$osess newReader {*}$args]
        set ofmt [$osess newFormatter]
        while {[llength [set hevts [$oreader getEvents -count 100]]]} {
            try {
                foreach evt [recordarray getlist \
                                 [$ofmt decodeEvents $hevts -properties {
                                     -providername -eventid -level -eventrecordid
                                     -timecreated -message
                                 }] -format dict] {
                    if {[info exists count] && [incr count -1] < 0} {
                        return
                    }
                    puts $outfd "[dict get $evt -timecreated] [dict get $evt -eventid] [dict get $evt -providername] [dict get $evt -eventrecordid]: [dict get $evt -message]"
                }
            } finally {
                evt_close {*}$hevts
            }
        }
    } finally {
        # Also destroyed dependent objects like oquery
        $osess destroy
    }
}

catch {twapi::EvtSession destroy}
catch {twapi::EvtPublisher destroy}
catch {twapi::EvtReader destroy}
catch {twapi::EvtChannelConfig destroy}
catch {twapi::EvtChannelInfo destroy}
catch {twapi::EvtLogInfo destroy}
catch {twapi::EvtResultSet destroy}
catch {twapi::EvtFormatter destroy}

oo::class create twapi::EvtSession {
    variable hSession
    variable nameCounter

    # Track objects that need to be destroyed when session is destroyed.
    # Since objects can be renamed, we track their namespaces as the key.
    variable dependentNamespaces

    constructor {args} {

        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set dependentNamespaces {}

        if {[llength $args] == 0} {
            set hSession NULL
            return
        }

        parseargs args {
            {system.arg ""}
            user.arg
            domain.arg
            password.arg
            {authtype.arg 0}
        } -nulldefault -maxleftover 0 -setvars

        if {![string is integer -strict $authtype]} {
            set authtype [dict get {default 0 negotiate 1 kerberos 2 ntlm 3} [string tolower $authtype]]
        }

        set hSession [EvtOpenSession 1 [list $system $user $domain $password $authtype] 0 0]
    }
    destructor {
        foreach dependent [dict keys $dependentNamespaces] {
            # Will destroy object implemented by the namespace
            catch {
                namespace delete $dependent
            }
        }
        if {![my isLocal]} {
            EvtClose $hSession
        }
    }
    method handle {} {return $hSession}
    method isLocal {} {return [string equal $hSession NULL]}
    method registeredChannels {} {
        set channels {}
        set hce [EvtOpenChannelEnum $hSession 0]
        try {
            while {[set chname [EvtNextChannelPath $hce]] ne ""} {
                lappend channels $chname
            }
        } finally {
            EvtClose $hce
        }

        return $channels
    }
    method clearChannel {channel args} {
        parseargs args {{backup.arg ""}} -maxleftover 0 -setvars
        return [EvtClearLog $hSession $channel [_evt_native_path $backup] 0]
    }
    method archiveLogFile {logpath args} {
        parseargs args {{lcid.int 0}} -maxleftover 0 -setvars
        return [EvtArchiveExportedLog $hSession \
                    [_evt_native_path $logpath] $lcid 0]
    }
    method exportEvents {outfile args} {
        lassign [_evt_parse_query_options args -maxleftover 0] source query flags
        EvtExportLog $hSession $source $query \
            [_evt_native_path $outfile] $flags
    }
    method createChannelConfig {objname channel} {
        set obj [uplevel 1 [list [namespace which -command EvtChannelConfig] \
                                create $objname [self] $channel]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newChannelConfig {channel} {
        return [my createChannelConfig [my NewName channel-config] $channel]
    }
    method createLogInfo {objname args} {
        set obj [uplevel 1 [list [namespace which -command EvtLogInfo] \
                                create $objname [self] {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newLogInfo {args} {
        return [my createLogInfo [my NewName log-info] {*}$args]
    }
    method publishers {} {
        set pubs {}
        set henum [EvtOpenPublisherEnum $hSession]
        try {
            while {[set pub [EvtNextPublisherId $henum]] ne ""} {
                lappend pubs $pub
            }
        } finally {
            evt_close $henum
        }
        return $pubs
    }
    method createPublisher {objname publisher args} {
        parseargs args {
            {lcid.int 0}
            {logarchive.arg ""}
        } -setvars -maxleftover 0
        set obj [uplevel 1 [list [namespace which -command EvtPublisher] \
                                create $objname [self] $publisher $lcid $logarchive]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newPublisher {publisher args} {
        return [my createPublisher [my NewName pub] $publisher {*}$args]
    }
    method createReader {objname args} {
        set obj [uplevel 1 [list [namespace which -command EvtReader] \
                                create $objname $hSession {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newReader {args} {
        return [my createReader [my NewName reader] {*}$args]
    }
    method createSubscription {objname channel args} {
        set obj [uplevel 1 [list [namespace which -command EvtSubscription] \
                                create $objname $hSession $channel {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newSubscription {channel args} {
        return [my createSubscription [my NewName subscription] $channel {*}$args]
    }
    method createFormatter {objname args} {
        set obj [uplevel 1 [list [namespace which -command EvtFormatter] \
                                create $objname [self] {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newFormatter {args} {
        return [my createFormatter [my NewName fmt] {*}$args]
    }
    method with {factory targetcall} {
        set obj [my {*}$factory]
        try {
            uplevel 1 [list $obj {*}$targetcall]
        } finally {
            $obj destroy
        }
    }
    method unregister {objs} {
        foreach obj $objs {
            set obj_ns [info object namespace $obj]
            dict unset dependentNamespace $obj_ns
        }
    }

    # Private methods
    method NewName {{name_part obj}} {
        return [string cat evt- $name_part - [incr nameCounter]]
    }

}

oo::class create twapi::EvtPublisher {
    variable hPublisher
    variable publisherName
    variable levelLabelMap
    variable taskLabelMap
    variable opcodeLabelMap
    variable keywordLabelMap
    variable localeId

    constructor {osess publisher {lcid 0} {logarchive {}}} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set publisherName $publisher
        set localeId $lcid
        set hPublisher [EvtOpenPublisherMetadata [$osess handle] \
                            $publisher $logarchive $lcid 0]
    }
    method name          {} {return $publisherName}
    method handle        {} {return $hPublisher}
    method lcid          {} {return $localeId}
    method guid          {} {EvtGetPublisherMetadataProperty $hPublisher 0}
    method resourceFile  {} {EvtGetPublisherMetadataProperty $hPublisher 1}
    method parameterFile {} {EvtGetPublisherMetadataProperty $hPublisher 2}
    method messageFile   {} {EvtGetPublisherMetadataProperty $hPublisher 3}
    method helpLink      {} {EvtGetPublisherMetadataProperty $hPublisher 4}
    method messageId     {} {EvtGetPublisherMetadataProperty $hPublisher 5}
    method channels {} {
        return [my GetPropertiesArray 6 {
            -channelpath 7 -channelindex 8 -channelid 9
            -channelflags 10 -channelmessageid 11
        } {-channelindex -channelmessageid}]
    }
    method levels {} {
        return [my GetPropertiesArray 12 {
            -name 13 -value 14 -messageid 15
        } {-messageid}]
    }
    method tasks {} {
        return [my GetPropertiesArray 16 {
            -name 17 -eventguid 18 -value 19
            -messageid 20
        } {-messageid}]
    }
    method opcodes {} {
        return [my GetPropertiesArray 21 {
            -name 22 -value 23 -messageid 24
        } {-messageid}]
    }
    method keywords {} {
        return [my GetPropertiesArray 25 {
            -name 26 -value 27 -messageid 28
        } {-messageid}]
    }
    method events {property_names} {
        set henum [EvtOpenEventMetadataEnum $hPublisher]

        # It is faster to build a list and then have Tcl shimmer to a dict when
        # required
        set meta {}
        try {
            while {[set hmeta [EvtNextEventMetadata $henum 0]] ne ""} {
                try {
                    set properties {}
                    foreach prop $property_names {
                        lappend properties $prop \
                            [EvtGetEventMetadataProperty $hmeta \
                                 [dict get {
                                     -eventid 0 -version 1 -channel 2 -level 3
                                     -opcode 4 -task 5 -keywords 6 -messageid 7 -template 8
                                 } $prop]]
                    }
                    lappend meta $properties
                } finally {
                    evt_close $hmeta
                }
            }
        } finally {
            evt_close $henum
        }

        return $meta
    }
    method message {msg_id} {
        # TBD - cache message id's
        # 8 -> EvtFormatMessageId
        return [EvtFormatMessage $hPublisher NULL $msg_id NULL 8]
    }
    method levelLabels {} {
        if {![info exists levelLabelMap]} {
            my InitLabels levelLabelMap levels
        }
        return $levelLabelMap
    }
    method taskLabels {} {
        if {![info exists taskLabelMap]} {
            my InitLabels taskLabelMap tasks
        }
        return $taskLabelMap
    }
    method opcodeLabels {} {
        if {![info exists opcodeLabelMap]} {
            my InitLabels opcodeLabelMap opcodes
        }
        return $opcodeLabelMap
    }
    method keywordLabels {} {
        if {![info exists keywordLabelMap]} {
            my InitLabels keywordLabelMap keywords
        }
        return $keywordLabelMap
    }

    # Private methods
    method InitLabels {name_map_var method_name} {
        set $name_map_var [dict create]
        foreach elem [my $method_name] {
            set value [dict get $elem -value]
            set msgid [dict get $elem -messageid]
            if {$msgid != -1} {
                if {![catch {my message $msgid} name]} {
                    dict set $name_map_var $value $name
                    continue
                }
            }
            dict set $name_map_var $value $value
        }
    }

    method GetPropertiesArray {property_enum definitions {minus_one_map {}}} {
        set harray [EvtGetPublisherMetadataProperty $hPublisher $property_enum]
        try {
            set n [EvtGetObjectArraySize $harray]
            set elems {}
            for {set i 0} {$i < $n} {incr i} {
                set elem [dict create]
                foreach {opt enum} $definitions {
                    set value [EvtGetObjectArrayProperty $harray $enum $i]
                    if {$opt in $minus_one_map && $value == 4294967295} {
                        set value -1
                    }
                    dict set elem $opt $value
                }
                lappend elems $elem
            }
            return $elems
        } finally {
            EvtClose $harray
        }
    }
}

# Intended to be used as mixin
oo::class create twapi::EvtResultSet {
    variable hResultSet
    constructor args {
        next {*}$args
    }
    destructor {
        if {[llength [self next]]} {
            next
        }
        EvtClose $hResultSet
    }
    method handle {} {return $hResultSet}
    method info {} {
        # Don't return as dictionary because in case of multiple references to
        # the same channel in a query list, not clear if it will be returned
        # just once or multiple types.
        lmap channel_name [EvtGetQueryInfo $hResultSet 0] \
            channel_status [EvtGetQueryInfo $hResultSet 1] {
                list $channel_name $channel_status
            }
    }
    method getEvents {args} {
        parseargs args {
            {timeout.int -1}
            {count.int 1}
            {statusvar.arg}
        } -maxleftover 0 -setvars

        if {[info exists statusvar]} {
            upvar 1 $statusvar status
            set hevts [EvtNext $hResultSet $count $timeout 0 status]
        } else {
            set hevts [EvtNext $hResultSet $count $timeout 0]
        }
        if {[llength $hevts]} {
            return $hevts
        }
        my EofHandler
        return $hevts
    }
    method SetHandle h {
        set hResultSet $h
    }
}

oo::class create twapi::EvtReader {
    mixin twapi::EvtResultSet
    variable hSession
    constructor {hsess args} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        parseargs args {
            {direction.sym forward {forward 0x100 backward 0x200}}
        } -ignoreunknown -setvars
        lassign [_evt_parse_query_options args -maxleftover 0] source query flags

        set hSession $hsess
        my SetHandle [EvtQuery $hsess $source $query \
                          [tcl::mathop::| $flags $direction]]
    }
    destructor {}
    method seek {offset args} {
        parseargs args {
            {origin.arg current {first last current}}
            hbookmark.arg
            {strict 0 0x10000}
        } -maxleftover 0 -setvars

        if {[info exists hbookmark]} {
            if {[info exists origin]} {
                error "At most one of options -hbookmark and -origin may be specified."
            }
            set flags 4
        } else {
            set flags [dict get {first 1 last 2 current 3} $origin]
            set hbookmark NULL
        }

        incr flags $strict

        EvtSeek [my handle] $pos $hbookmark 0 $flags
    }
    method EofHandler {} {}
}

oo::class create twapi::EvtSubscription {
    mixin twapi::EvtResultSet

    variable hSession
    variable hSignal
    variable commandPrefix
    variable callbackTimeout

    constructor {hsess channel args} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        parseargs args {
            {query.arg {}}
            hbookmark.arg
            includeexisting
            {ignorequeryerrors 0 0x1000}
            {strict 0 0x10000}
        } -maxleftover 0 -setvars

        if {[info exists hbookmark]} {
            set flags 3
        } else {
            set hbookmark NULL
            set flags [expr {$includeexisting ? 2 : 1}]
        }
        set flags [expr {$flags | $ignorequeryerrors | $strict}]
        # MUST be manual rest and initially signalled as per the SDK
        # example, else subscriptions do not work.
        set hsig [create_event -manualreset 1 -signalled 1]
        try {
            set hsub [EvtSubscribe $hsess $hsig $channel $query $hbookmark $flags]
        } on error {message ropts} {
            CloseHandle $hsig
            return -options $ropts $message
        }

        set asyncMode 0
        set hSession $hsess
        set hSignal $hsig
        my SetHandle $hsub
        # Do not init commandPrefix until callback is registered
    }
    destructor {
        if {[info exists commandPrefix]} {
            cancel_wait_on_handle $hSignal
        }
        CloseHandle $hSignal
    }
    method waitForEvent {{ms -1}} {
        if {[info exists commandPrefix]} {
            error "Cannot wait on a subscription that has registered callbacks."
        }
        # Don't really care why the wait completed (signalled, timeout, abandoned)
        # Caller simply needs to call the getEvents method in all cases
        wait_on_handle $hSignal -wait $ms
    }
    method registerCallback {cb {ms -1}} {
        if {$cb eq ""} {
            error "Cannot register empty callback."
        }
        # Verify well formed list
        llength $cb

        # If callback already exists, no need to register handler again
        if {![info exists commandPrefix]} {
            wait_on_handle $hSignal -async [mymethod SignalHandler] \
                -executeonce 1 -timeout $ms
        }
        set commandPrefix $cb
        set callbackTimeout [incr ms 0]
    }
    method unregisterCallback {} {
        if {[info exists commandPrefix]} {
            cancel_wait_on_handle $hSignal
            unset commandPrefix
        }
    }
    method SignalHandler {hsig trigger} {
        # Irrespective of whether trigger is signalled, timeout, abandoned,
        # action to be taken is the same.
        if {[info exists commandPrefix]} {
            uplevel #0 $commandPrefix
        }
    }
    method EofHandler {} {
        twapi::reset_event $hSignal
        if {[info exists commandPrefix]} {
            wait_on_handle $hSignal -async [mymethod SignalHandler] \
                -executeonce 1 -timeout $callbackTimeout
        }
    }
}

oo::class create twapi::EvtFormatter {
    # Owning session object
    variable oSession

    # LCID for formatting messages
    variable localeId

    # metadata file to look up before system registry
    variable logArchive

    # Dictionary mapping publisher names to their wrapper objects
    # publisher names are case-insensitive. However, we do not bother
    # normalizing names to use as keys as duplicate objects cause no harm
    variable publisherObjs

    # Dictionary mapping publisher names to their handles
    variable publisherHandles

    # Rendering context for system fields
    variable hSystemContext

    # Rendering context for user fields
    variable hUserContext

    # Reusable buffer for rendering values
    variable renderBuffer

    # Map of (publisher, level/task/opcode/keywords) -> names
    variable levelLabelMap
    variable taskLabelMap
    variable opcodeLabelMap
    variable keywordLabelMap

    initialize {
        # system properties mapped to their position in an event
        variable systemPropertyNameMap
        array set systemPropertyNameMap {
            -providername 0 -providerguid 1 -eventid 2 -qualifiers 3 -level 4
            -task 5 -opcode 6 -keywords 7 -timecreated 8 -eventrecordid 9
            -activityid 10 -relatedactivityid 11 -pid 12 -tid 13 -channel 14
            -computer 15 -sid 16 -version 17
        }
        variable eventPropertyNames [concat \
                                         [array names systemPropertyNameMap] \
                                         {-userdata}]
    }
    constructor {osess args} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        parseargs args {
            {logarchive.arg ""}
            {lcid.int 0}
        } -maxleftover 0 -setvars
        set oSession $osess
        set localeId $lcid
        set publisherObjs [dict create]
        set publisherHandles [dict create]
        set logArchive [_evt_native_path $logarchive]
        set hSystemContext [EvtCreateRenderContext {} 1]
        set hUserContext [EvtCreateRenderContext {} 2]
        set levelLabelMap [dict create]
        set taskLabelMap [dict create]
        set opcodeLabelMap [dict create]
        set keywordLabelMap [dict create]
        set renderBuffer NULL
    }
    destructor {
        if {$renderBuffer ne "NULL"} {
            evt_free_EVT_RENDER_VALUES $renderBuffer
        }
        EvtClose $hSystemContext
        EvtClose $hUserContext
        $oSession unregister [dict values $publisherObjs]
        $oSession unregister [self]
    }
    method decodeEvents {hevts {properties {-providername -eventid -level -timecreated -message}}} {
        classvariable systemPropertyNameMap eventPropertyNames

        return [list $properties [lmap hevt $hevts {
            set system_properties [my EventSystemProperties $hevt]
            set publisher [lindex $system_properties 0]
            lmap prop_name $properties {
                switch -exact -- $prop_name {
                    -level {
                        my LevelLabel $publisher [lindex $system_properties 4]
                    }
                    -task {
                        my TaskLabel $publisher [lindex $system_properties 5]
                    }
                    -opcode {
                        my OpcodeLabel $publisher [lindex $system_properties 6]
                    }
                    -keywords {
                        my KeywordLabels $publisher [lindex $system_properties 7]
                    }
                    -userdata {
                        my EventUserProperties $hevt
                    }
                    -message {
                        my EventMessage $publisher $hevt
                    }
                    default {
                        lindex $system_properties $systemPropertyNameMap($prop_name)
                    }
                }
            }
        }]]
    }
    method decodeEvent {hevt properties} {
        return [recordarray index \
                    [my decodeEvents [list $hevt] $properties] \
                    0 -format dict]
    }
    method formatEvent hevt {
        my EventMessage [lindex [my EventSystemProperties $hevt] 0] $hevt
    }
    method formatEventAsXml hevt {
        set publisher [lindex [my EventSystemProperties $hevt] 0]
        ## 9 -> EvtFormatMessageXml
        return [EvtFormatMessage [my PublisherHandle $publisher] $hevt 0 NULL 9]
    }

    method PublisherObj publisher {
        if {[dict exists $publisherObjs $publisher]} {
            return [dict get $publisherObjs $publisher]
        }
        set obj [$oSession newPublisher $publisher \
                     -lcid $localeId -logarchive $logArchive]
        dict set publisherObjs $publisher $obj
        return $obj
    }
    method PublisherHandle publisher {
        if {[dict exists $publisherHandles $publisher]} {
            return [dict get $publisherHandles $publisher]
        }

        if {[catch {my PublisherObj $publisher} obj]} {
            dict set publisherHandles $publisher NULL
            return NULL
        } else {
            set h [$obj handle]
            dict set publisherHandles $publisher $h
            return $h
        }
    }
    method EventSystemProperties {hevt} {
        set renderBuffer [Twapi_EvtRenderValues $hSystemContext $hevt $renderBuffer]
        return [Twapi_ExtractEVT_RENDER_VALUES $renderBuffer]
    }
    method EventUserProperties -export {hevt} {
        set renderBuffer [Twapi_EvtRenderValues $hUserContext $hevt $renderBuffer]
        return [Twapi_ExtractEVT_RENDER_VALUES $renderBuffer]
    }
    method LevelLabel -export {publisher level} {
        if {[dict exists $levelLabelMap $publisher]} {
            return [dict getdef $levelLabelMap $publisher $level $level]
        }
        # Not in cache. Get from publisher. May fail because there is no such
        # publisher registered.
        if {![catch {
            dict set levelLabelMap $publisher [[my PublisherObj $publisher] levelLabels]
        }]} {
            return [dict getdef $levelLabelMap $publisher $level $level]
        }
        if {![dict exists $levelLabelMap ""]} {
            # Default localized level names
            # Get default level names used by Windows Eventlog
            if {[catch {
                set map [[my PublisherObj Microsoft-Windows-Eventlog] levelLabels]
            }]} {
                # Even that failed, so use English mappings
                set map {1 Critical 2 Error 3 Warning 4 Information 5 Verbose}
            }
            dict set levelLabelMap "" $map
        }
        dict set levelLabelMap $publisher [dict get $levelLabelMap ""]
        return [dict getdef $levelLabelMap $publisher $level $level]
    }
    method TaskLabel -export {publisher task} {
        if {[dict exists $taskLabelMap $publisher]} {
            return [dict getdef $taskLabelMap $publisher $task $task]
        }
        # Not in cache. Get from publisher. May fail because there is no such
        # publisher registered.
        if {[catch {
            dict set taskLabelMap $publisher [[my PublisherObj $publisher] taskLabels]
        }]} {
            # Default to the task id. Unlike for levels, we do not try
            # Microsoft-Windows-Eventlog or have any predefined names for tasks.
            dict set taskLabelMap $publisher $task $task
        }
        return [dict getdef $taskLabelMap $publisher $task $task]
    }
    method KeywordLabels -export {publisher keywords} {
        # Treat as a bit mask else loop below will continue forever
        # on negative 64-bit values
        set keywords [expr {$keywords & 0xffffffffffffffff}]
        set names {}
        # keywords are a bitmask, each bit being a keyword
        while {$keywords} {
            set keyword [expr {$keywords & -$keywords}]
            set keywords [expr {$keywords & ~$keyword}]
            lappend names [my KeywordLabel $publisher $keyword]
        }
        return $names
    }
    method KeywordLabel {publisher keyword} {
        if {[dict exists $keywordLabelMap $publisher]} {
            return [dict getdef $keywordLabelMap $publisher $keyword $keyword]
        }
        # Not in cache. Get from publisher. May fail because there is no such
        # publisher registered.
        if {[catch {
            dict set keywordLabelMap $publisher [[my PublisherObj $publisher] keywordLabels]
        }]} {
            # Default to the keyword id. Unlike for levels, we do not try
            # Microsoft-Windows-Eventlog or have any predefined names for keywords.
            dict set keywordLabelMap $publisher $keyword $keyword
        }
        return [dict getdef $keywordLabelMap $publisher $keyword $keyword]
    }

    method EventMessage {publisher hevt} {
        if {[EvtFormatMessage [my PublisherHandle $publisher] \
                 $hevt 0 NULL 1 message]} {
            return $message
        } elseif {[EvtFormatMessage NULL $hevt 0 NULL 1 message]} {
            # try with NULL publisher handler. In this case EvtFormatMessage
            # will use the rendering info stored within the event in case it is
            # a forwarded event from another system.
            return $message
        } else {
            # TBD - make sure we have a test for this case.
            set message "Message for event could not be found."
            catch {
                append message " Event user data: " [join [my EventUserProperties $hevt] ", "]
            }
            return $message
        }
    }
}

oo::class create twapi::EvtLogInfo {
    variable hInfo
    variable logName
    variable logType
    constructor {osess args} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        parseargs args {
            channel.arg
            logfile.arg
        } -setvars -maxleftover 0

        if {[info exists logfile]} {
            if {[info exists channel]} {
                error "At most one of -channel and -logfile may be specified."
            }
            set logName $logfile
            set flags 2
            set logType file
        } else {
            set flags 1
            set logType channel
            if {[info exists channel]} {
                set logName $channel
            } else {
                set logName Application
            }
        }
        set hInfo [twapi::EvtOpenLog [$osess handle] $logName $flags]
    }
    method creationTime       {} {EvtGetLogInfo $hInfo 0}
    method lastAccessTime     {} {EvtGetLogInfo $hInfo 1}
    method lastWriteTime      {} {EvtGetLogInfo $hInfo 2}
    method fileSize           {} {EvtGetLogInfo $hInfo 3}
    method attributes         {} {EvtGetLogInfo $hInfo 4}
    method recordCount        {} {EvtGetLogInfo $hInfo 5}
    method oldestRecordNumber {} {EvtGetLogInfo $hInfo 6}
    method isFull             {} {EvtGetLogInfo $hInfo 7}
    method logType {} { return $logType }
    method logName {} { return $logName }
    destructor {
        EvtClose $hInfo
    }
}

oo::class create twapi::EvtChannelConfig {
    variable hConfig
    variable oSession
    variable channelName
    constructor {osess channel} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set channelName $channel
        set oSession $osess
        set hConfig [twapi::EvtOpenChannelConfig [$osess handle] $channel 0]
    }
    destructor {
        EvtClose $hConfig
        $oSession unregister [list [self]]
    }
    method name {} {return $channelName}
    method save {} {EvtSaveChannelConfig $hConfig}

    method isEnabled        {}    {EvtGetChannelConfigProperty $hConfig 0}
    method setEnabled       {val} {EvtSetChannelConfigProperty $hConfig 0 0 $val}
    method isolation        {}    {EvtGetChannelConfigProperty $hConfig 1}
    method setIsolation     {val} {EvtSetChannelConfigProperty $hConfig 1 0 $val}
    method type             {}    {EvtGetChannelConfigProperty $hConfig 2}
    method publisher        {}    {EvtGetChannelConfigProperty $hConfig 3}
    method isClassic        {}    {EvtGetChannelConfigProperty $hConfig 4}
    method access           {}    {EvtGetChannelConfigProperty $hConfig 5}
    method setAccess        {val} {EvtSetChannelConfigProperty $hConfig 5 0 $val}
    method retention        {}    {EvtGetChannelConfigProperty $hConfig 6}
    method setRetention     {val} {EvtSetChannelConfigProperty $hConfig 6 0 $val}
    method hasAutoBackup    {}    {EvtGetChannelConfigProperty $hConfig 7}
    method setAutoBackup    {val} {EvtSetChannelConfigProperty $hConfig 7 0 $val}
    method maxSize          {}    {EvtGetChannelConfigProperty $hConfig 8}
    method setMaxSize       {val} {EvtSetChannelConfigProperty $hConfig 8 0 $val}
    method filePath         {}    {EvtGetChannelConfigProperty $hConfig 9}
    method setFilePath      {val} {EvtSetChannelConfigProperty $hConfig 9 0 $val}
    method levelFilter      {}    {EvtGetChannelConfigProperty $hConfig 10}
    method setLevelFilter   {val} {EvtGetChannelConfigProperty $hConfig 10 0 $val}
    method keywordFilter    {}    {EvtGetChannelConfigProperty $hConfig 11}
    method setKeywordFilter {val} {EvtGetChannelConfigProperty $hConfig 11 0 $val}
    method controlGuid      {}    {EvtGetChannelConfigProperty $hConfig 12}
    method bufferSize       {}    {EvtGetChannelConfigProperty $hConfig 13}
    method minBuffers       {}    {EvtGetChannelConfigProperty $hConfig 14}
    method maxBuffers       {}    {EvtGetChannelConfigProperty $hConfig 15}
    method latency          {}    {EvtGetChannelConfigProperty $hConfig 16}
    method clockType        {}    {EvtGetChannelConfigProperty $hConfig 17}
    method sidType          {}    {EvtGetChannelConfigProperty $hConfig 18}
    method publishers       {}    {EvtGetChannelConfigProperty $hConfig 19}
    method maxFiles         {}    {EvtGetChannelConfigProperty $hConfig 20}
    method setMaxFiles      {val} {EvtSetChannelConfigProperty $hConfig 20 0 $val}
}
