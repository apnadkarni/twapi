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
    method newLogFileQuery {logpath args} {
        return [my createFileQuery [my NewName file-query] $logpath {*}$args]
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
            -levelname 13 -levelvalue 14 -levelmessageid 15
        } {-levelmessageid}]
    }
    method tasks {} {
        return [my GetPropertiesArray 16 {
            -taskname 17 -taskeventguid 18 -taskvalue 19
            -taskmessageid 20
        } {-taskmessageid}]
    }
    method opcodes {} {
        return [my GetPropertiesArray 21 {
            -opcodename 22 -opcodevalue 23 -opcodemessageid 24
        } {-opcodemessageid}]
    }
    method keywords {} {
        return [my GetPropertiesArray 25 {
            -keywordname 26 -keywordvalue 27 -keywordmessageid 28
        } {-keywordmessageid}]
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

oo::class create twapi::EventLogQuery {
    variable hResultSet
    variable hSession
    constructor {hsess source flags args} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        parseargs args {
            {query.arg {}}
            {ignorequeryerrors 0 0x1000}
            {direction.sym forward {forward 0x100 reverse 0x200 backward 0x200}}
        } -maxleftover 0 -setvars
        set hSession $hsess
        set hResultSet [EvtQuery $hsess $source $query \
                    [tcl::mathop::| $flags $ignorequeryerrors $direction]]

    }
    method next {args} {
        parseargs args {
            {timeout.int -1}
            {count.int 1}
            {statusvar.arg}
        } -maxleftover 0 -setvars

        if {[info exists statusvar]} {
            upvar 1 $statusvar status
            return [EvtNext $hResultSet $count $timeout 0 status]
        } else {
            return [EvtNext $hResultSet $count $timeout 0]
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
    variable publisherObjs

    # Rendering context for system fields
    variable hSystemContext

    # Reusable buffer for rendering values
    variable renderBuffer

    constructor {osess args} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        parseargs args {
            {logarchive.arg ""}
            {lcid.int 0}
        } -maxleftover 0 -setvars
        set oSession $osess
        set localeId $lcid
        set publisherObjs [dict create]
        set logArchive [_evt_native_path $logarchive]
        set hSystemContext [EvtCreateRenderContext {} 1]

        set renderBuffer NULL
    }
    destructor {
        evt_free_render_values $hRenderValuesBuffer
        EvtClose $hSystemContext
        $oSession unregister [dict values $publisherObjs]
    }
    method formatEventAsXml {hevt} {
        set publisher [lindex [my EventSystemProperties $hevt] 0]
        ## 9 -> EvtFormatMessageXml
        return [EvtFormatMessage [my PublisherHandle $publisher] $hevt 0 NULL 9]
    }
    method PublisherHandle publisher {
        if {![dict exists $publisherObjs $publisher]} {
            dict set publisherObjs $publisher \
                [$oSession newPublisher $publisher \
                     -lcid $localeId -logarchive $logArchive]
        }
        return [[dict get $publisherObjs $publisher] handle]
    }
    method EventSystemProperties {hevt} {
        set renderBuffer [Twapi_EvtRenderValues $hSystemContext $hevt $renderBuffer]
        return [Twapi_ExtractEVT_RENDER_VALUES $renderBuffer]
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

proc twapi::evt_free_render_values {p} {
    evt_free $p
}
