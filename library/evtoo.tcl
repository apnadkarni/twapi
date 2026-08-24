#
# Copyright (c) 2012-2026, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Event log handling for Vista and later

namespace eval twapi {}

catch {twapi::EventLogSession destroy}
catch {twapi::EventLogPublisher destroy}
catch {twapi::EventLogChannelConfig destroy}
catch {twapi::EventLogChannelInfo destroy}
catch {twapi::EventLogInfo destroy}
catch {twapi::EventResultSet destroy}
catch {twapi::EventLogQuery destroy}
catch {twapi::EventLogFormatter destroy}

oo::class create twapi::EventLogSession {
    variable hSession
    variable nameCounter

    # Track objects that need to be destroyed when session is destroyed.
    # Since objects can be renamed, we track their namespaces as the key.
    variable dependentNamespaces

    constructor args {

        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set dependentNamespaces {}

        if {[llength $args] == 0} {
            set hSession NULL
            return
        }

        parseargs args {
            user.arg
            domain.arg
            password.arg
            {authtype.arg 0}
        } -nulldefault -maxleftover 0 -setvars

        if {![string is integer -strict $authtype]} {
            set authtype [dict get {default 0 negotiate 1 kerberos 2 ntlm 3} [string tolower $authtype]]
        }

        set hSession [EvtOpenSession 1 [list $server $user $domain $password $authtype] 0 0]
    }
    destructor {
        foreach dependent [dict keys $dependentNamespaces] {
            # Will destroy object implemented by the namespace
            catch {namespace delete $dependent}
        }
        if {![my isLocal]} {
            EvtClose $hSession
        }
    }
    method handle {} {return $hSession}
    method isLocal {} {return [string equal $hSession NULL]}
    method channels {} {
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
    method archiveLogfile {logpath args} {
        parseargs args {{lcid.int 0}} -maxleftover 0 -setvars
        return [EvtArchiveExportedLog $hSession \
                    [_evt_native_path $logpath] $lcid 0]
    }
    method exportChannelEvents {channel outfile args} {
        parseargs args {
            {query.arg *}
            {ignorequeryerrors 0 0x1000}
        } -maxleftover 0 -setvars

        set flags [expr {$ignorequeryerrors | 1}]
        EvtExportLog $hSession $channel $query \
            [_evt_native_path $outfile] $flags
    }
    method exportFileEvents {logpath outfile args} {
        parseargs args {
            {query.arg *}
            {ignorequeryerrors 0 0x1000}
        } -maxleftover 0 -setvars

        set flags [expr {$ignorequeryerrors | 2}]
        EvtExportLog $hSession [_evt_native_path $logpath] $query \
            [_evt_native_path $outfile] $flags
    }
    method createChannelConfig {objname channel} {
        set obj [uplevel 1 [list [namespace which -command EventLogChannelConfig] \
                                create $objname [self] $channel]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newChannelConfig {channel} {
        return [my createChannelConfig [my NewName channel-config] $channel]
    }
    method createChannelInfo {objname channel} {
        set obj [uplevel 1 [list [namespace which -command EventLogChannelInfo] \
                                create $objname [self] $channel]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newChannelInfo {channel} {
        return [my createChannelInfo [my NewName channel-info] $channel]
    }
    method createLogFileInfo {objname logpath} {
        set obj [uplevel 1 [list [namespace which -command EventLogFileInfo] \
                                create $objname [self] $logpath]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newLogFileInfo {logpath} {
        return [my createLogFileInfo [my NewName logfile-info] $logpath]
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
        set obj [uplevel 1 [list [namespace which -command EventLogPublisher] \
                                create $objname [self] $publisher $lcid $logarchive]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newPublisher {publisher args} {
        return [my createPublisher [my NewName pub] $publisher {*}$args]
    }
    method createChannelQuery {objname channel args} {
        set obj [uplevel 1 [list [namespace which -command EventLogQuery] \
                                create $objname $hSession $channel 1 {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newChannelQuery {channel args} {
        return [my createChannelQuery [my NewName chan-query] $channel {*}$args]
    }
    method createFileQuery {objname logpath args} {
        set obj [uplevel 1 [list [namespace which -command EventLogQuery] \
                                create $objname $hSession $logpath 2 {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newFileQuery {logpath args} {
        return [my createFileQuery [my NewName file-query] $logpath {*}$args]
    }
    method createSubscription {objname channel args} {
        set obj [uplevel 1 [list [namespace which -command EventLogSubscription] \
                                create $objname $hSession $channel {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newSubscription {channel args} {
        return [my createSubscription [my NewName subscription] $channel {*}$args]
    }
    method createFormatter {objname args} {
        set obj [uplevel 1 [list [namespace which -command EventLogFormatter] \
                                create $objname [self] {*}$args]]
        dict set dependentNamespaces [info object namespace $obj] $obj
        return $obj
    }
    method newFormatter {args} {
        return [my createFormatter [my NewName fmt] {*}$args]
    }
    method unregister {objs} {
        foreach obj $objs {
            set obj_ns [info object namespace $obj]
            $obj destroy
            dict unset dependentNamespace $obj_ns
        }
    }

    # Private methods
    method NewName {{name_part obj}} {
        return [string cat evt- $name_part - [incr nameCounter]]
    }

}

oo::class create twapi::EventLogPublisher {
    variable hPublisher
    variable publisherName
    variable levelNameMap
    variable taskNameMap
    variable opcodeNameMap
    variable keywordNameMap

    constructor {osess publisher {lcid 0} {logarchive {}}} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set publisherName $publisher
        set hPublisher [EvtOpenPublisherMetadata [$osess handle] \
                            $publisher $logarchive $lcid 0]
    }
    method name          {} {return $publisherName}
    method handle        {} {return $hPublisher}
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
    method eventDefinitions {property_names} {
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
                                     -id 0 -version 1 -channel 2 -level 3
                                     -opcode 4 -task 5 -keyword 6 -messageid 7 -template 8
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
    method levelNames {} {
        if {![info exists levelNameMap]} {
            my InitNames levelNameMap levels
        }
        return $levelNameMap
    }
    method taskNames {} {
        if {![info exists taskNameMap]} {
            my InitNames taskNameMap tasks
        }
        return $taskNameMap
    }
    method opcodeNames {} {
        if {![info exists opcodeNameMap]} {
            my InitNames opcodeNameMap opcodes
        }
        return $opcodeNameMap
    }
    method keywordNames {} {
        if {![info exists keywordNameMap]} {
            my InitNames keywordNameMap keywords
        }
        return $keywordNameMap
    }

    # Private methods
    method InitNames {name_map_var method_name} {
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
oo::class create twapi::EventResultSet {
    variable hResultSet
    constructor args {
        next {*}$args
    }
    destructor {
        next
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

oo::class create twapi::EventLogQuery {
    mixin twapi::EventResultSet
    variable hSession
    constructor {hsess source flags args} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        parseargs args {
            {query.arg {}}
            {ignorequeryerrors 0 0x1000}
            {direction.sym forward {forward 0x100 backward 0x200}}
        } -maxleftover 0 -setvars
        set hSession $hsess
        my SetHandle [EvtQuery $hsess $source $query \
                          [tcl::mathop::| $flags $ignorequeryerrors $direction]]

    }
    destructor {}
    method EofHandler {} {}
}

oo::class create twapi::EventLogSubscription {
    mixin twapi::EventResultSet

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
    method wait {{ms -1}} {
        if {[info exists commandPrefix]} {
            error "Cannot wait on a subscription that has registered callbacks."
        }
        # Don't really care why the wait completed (signalled, timeout, abandoned)
        # Caller simply needs to call the next method in all cases
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
        # action to be taken is the same. Invoke the callback
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

oo::class create twapi::EventLogFormatter {
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
    variable levelNameMap
    variable taskNameMap
    variable opcodeNameMap
    variable keywordNameMap

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
        set levelNameMap [dict create]
        set taskNameMap [dict create]
        set opcodeNameMap [dict create]
        set keywordNameMap [dict create]
        set renderBuffer NULL
    }
    destructor {
        evt_free_render_values $hRenderValuesBuffer
        EvtClose $hSystemContext
        EvtClose $hUserContext
        $oSession unregister [dict values $publisherObjs]
    }
    method decodeEvents {hevts args} {
        classvariable systemPropertyNameMap eventPropertyNames
        # TBD - ignorestring, raw
        parseargs args {
            ignorestring.arg
            {properties.arg {-providername -eventid -level -task -timecreated -pid}}
            {raw 0}
        } -setvars -maxleftover 0

        return [list $properties [lmap hevt $hevts {
            set system_properties [my EventSystemProperties $hevt]
            set publisher [lindex $system_properties 0]
            set rec [lmap prop_name $properties {
                switch -exact -- $prop_name {
                    -level {
                        my LevelName $publisher [lindex $system_properties 4]
                    }
                    -task {
                        my TaskName $publisher [lindex $system_properties 5]
                    }
                    -opcode {
                        my OpcodeName $publisher [lindex $system_properties 6]
                    }
                    -keywords {
                        my KeywordNames $publisher [lindex $system_properties 7]
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
            }]
        }]]
    }
    method decodeEvent {hevt args} {
        return [recordarray index \
                    [my decodeEvents [list $hevt] {*}$args] \
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
    method seek {offset args} {
        parseargs args {
            {origin.arg current {first last current}}
            bookmark.arg
            {strict 0 0x10000}
        } -maxleftover 0 -setvars

        if {[info exists bookmark]} {
            set flags 4
        } else {
            set flags [dict get {first 1 last 2 current 3} $origin]
            set bookmark NULL
        }

        incr flags $strict

        EvtSeek $hresults $pos $bookmark 0 $flags
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
    method LevelName -export {publisher level} {
        if {[dict exists $levelNameMap $publisher]} {
            return [dict getdef $levelNameMap $publisher $level $level]
        }
        # Not in cache. Get from publisher. May fail because there is no such
        # publisher registered.
        if {![catch {
            dict set levelNameMap $publisher [[my PublisherObj $publisher] levelNames]
        }]} {
            return [dict getdef $levelNameMap $publisher $level $level]
        }
        if {![dict exists $levelNameMap ""]} {
            # Default localized level names
            # Get default level names used by Windows Eventlog
            if {[catch {
                set map [[my PublisherObj Microsoft-Windows-Eventlog] levelNames]
            }]} {
                # Even that failed, so use English mappings
                set map {1 Critical 2 Error 3 Warning 4 Information 5 Verbose}
            }
            dict set levelNameMap "" $map
        }
        dict set levelNameMap $publisher [dict get $levelNameMap ""]
        return [dict getdef $levelNameMap $publisher $level $level]
    }
    method TaskName -export {publisher task} {
        if {[dict exists $taskNameMap $publisher]} {
            return [dict getdef $taskNameMap $publisher $task $task]
        }
        # Not in cache. Get from publisher. May fail because there is no such
        # publisher registered.
        if {[catch {
            dict set taskNameMap $publisher [[my PublisherObj $publisher] taskNames]
        }]} {
            # Default to the task id. Unlike for levels, we do not try
            # Microsoft-Windows-Eventlog or have any predefined names for tasks.
            dict set taskNameMap $publisher $task $task
        }
        return [dict getdef $taskNameMap $publisher $task $task]
    }
    method KeywordNames -export {publisher keywords} {
        # Treat as a bit mask else loop below will continue forever
        # on negative 64-bit values
        set keywords [expr {$keywords & 0xffffffffffffffff}]
        set names {}
        # keywords are a bitmask, each bit being a keyword
        while {$keywords} {
            set keyword [expr {$keywords & -$keywords}]
            set keywords [expr {$keywords & ~$keyword}]
            lappend names [my KeywordName $publisher $keyword]
        }
        return $names
    }
    method KeywordName {publisher keyword} {
        if {[dict exists $keywordNameMap $publisher]} {
            return [dict getdef $keywordNameMap $publisher $keyword $keyword]
        }
        # Not in cache. Get from publisher. May fail because there is no such
        # publisher registered.
        if {[catch {
            dict set keywordNameMap $publisher [[my PublisherObj $publisher] keywordNames]
        }]} {
            # Default to the keyword id. Unlike for levels, we do not try
            # Microsoft-Windows-Eventlog or have any predefined names for keywords.
            dict set keywordNameMap $publisher $keyword $keyword
        }
        return [dict getdef $keywordNameMap $publisher $keyword $keyword]
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

oo::class create twapi::EventLogInfo {
    variable hInfo
    constructor {h} {
        set hInfo $h
    }
    method creationTime       {} {EvtGetLogInfo $hInfo 0}
    method lastAccessTime     {} {EvtGetLogInfo $hInfo 1}
    method lastWriteTime      {} {EvtGetLogInfo $hInfo 2}
    method fileSize           {} {EvtGetLogInfo $hInfo 3}
    method attributes         {} {EvtGetLogInfo $hInfo 4}
    method recordCount        {} {EvtGetLogInfo $hInfo 5}
    method oldestRecordNumber {} {EvtGetLogInfo $hInfo 6}
    method isFull             {} {EvtGetLogInfo $hInfo 7}
    destructor {
        EvtClose $hInfo
    }
}

oo::class create twapi::EventLogChannelInfo {
    superclass twapi::EventLogInfo
    variable channelName
    constructor {osess channel} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set channelName $channel
        next [twapi::EvtOpenLog [$osess handle] $channel 1]
    }
    method channel {} {return $channelName}
}

oo::class create twapi::EventLogFileInfo {
    superclass twapi::EventLogInfo
    variable filePath
    constructor {osess logfile} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set filePath $logfile
        next [twapi::EvtOpenLog [$osess handle] $logfile 1]
    }
    method filePath {} {return $filePath}
}

oo::class create twapi::EventLogChannelConfig {
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
        $oSession 
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
    method setRetention     {val} {EvtGetChannelConfigProperty $hConfig 6 0 $val}
    method hasAutoBackup    {}    {EvtGetChannelConfigProperty $hConfig 7}
    method setAutoBackup    {val} {EvtGetChannelConfigProperty $hConfig 7 0 $val}
    method maxSize          {}    {EvtGetChannelConfigProperty $hConfig 8}
    method setMaxSize       {val} {EvtGetChannelConfigProperty $hConfig 8 0 $val}
    method filePath         {}    {EvtGetChannelConfigProperty $hConfig 9}
    method setFilePath      {val} {EvtGetChannelConfigProperty $hConfig 9 0 $val}
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
    method setMaxFiles      {val} {EvtGetChannelConfigProperty $hConfig 20 0 $val}
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

proc twapi::evt_event_logpath {hevt} {
    return [EvtGetEventInfo $hevt 1]
}

proc twapi::evt_render_context_xpaths {xpaths} {
    return [EvtCreateRenderContext $xpaths 0]
}

proc twapi::evt_render_context_system {} {
    return [EvtCreateRenderContext {} 1]
}
proc twapi::_evt_native_path {path} {
    # Do not want to rely on [file normalize] returning "" for ""
    if {$path eq ""} {
        return ""
    } else {
        return [file nativename [file normalize $path]]
    }
}

proc twapi::evt_dump {args} {
    parseargs args {
        channel.arg
        logfile.arg
        {outfd.arg stdout}
        count.int
    } -ignoreunknown -setvars

    if {[info exists channel]} {
        if {[info exists logfile]} {
            error "At most one of -channel and -file may be specified."
        }
    } elseif {![info exists logfile]} {
        set channel Application
    }

    set osess [EventLogSession new]
    try {
        if {[info exists channel]} {
            set oquery [$osess newChannelQuery $channel {*}$args]
        } else {
            set oquery [$osess newFileQuery $logfile {*}$args]
        }
        set ofmt [$osess newFormatter]
        while {[llength [set hevts [$oquery getEvents -count 100]]]} {
            try {
                foreach evt [recordarray getlist \
                                 [$ofmt decodeEvents $hevts -properties {
                                     -providername -eventid -level -eventrecordid
                                     -timecreated -message
                                 }] -format dict] {
                    if {[info exists count] && [incr count -1] < 0} {
                        return
                    }
                    puts $outfd "[dict get $evt -timecreated] [dict get $evt -eventid] [dict get $evt -providername]: [dict get $evt -eventrecordid] [dict get $evt -message]"
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


