#
# Copyright (c) 2012-2026, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Event log handling for Vista and later

namespace eval twapi {}

catch {twapi::EventLogSession destroy}
catch {twapi::EventLogChannelConfig destroy}
catch {twapi::EventLogChannelInfo destroy}
catch {twapi::EventLogInfo destroy}

oo::class create twapi::EventLogSession {
    variable hSession
    variable nameCounter
    # Track objects that need to be destroyed when session is destroyed.
    # Since objects can be renamed, we track their namespaces
    variable dependentNamespaces

    constructor args {

        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        puts nspath:[namespace path]
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
        foreach dependent $dependentNamespace {
            catch {namespace delete $dependent}
        }
        if {![my isLocal]} {
            EvtClose $hSession
        }
    }
    method isLocal {} {
        return [string equal $hSession NULL]
    }
    method channels {} {
        set channels {}
        set hce [EvtOpenChannelEnum $hSession 0]
        trap {
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
    method exportChannel {channel outfile args} {
        parseargs args {
            {query.arg *}
            {ignorequeryerrors 0 0x1000}
        } -maxleftover 0 -setvars

        set flags [expr {$ignorequeryerrors | 1}]
        EvtExportLog $hSession $channel $query \
            [_evt_native_path $outfile] $flags
    }
    method exportFile {logpath outfile args} {
        parseargs args {
            {query.arg *}
            {ignorequeryerrors 0 0x1000}
        } -maxleftover 0 -setvars

        set flags [expr {$ignorequeryerrors | 2}]
        EvtExportLog $hSession [_evt_native_path $logpath] $query \
            [_evt_native_path $outfile] $flags
    }
    method openChannelConfig {channel {objname {}}} {
        if {$objname eq ""} {
            set objname evt-chan-config-[incr nameCounter]
        }
        set obj [uplevel 1 [list [namespace which -command EventLogChannelConfig] \
                                create $objname $hSession $channel]]
        lappend dependentNamespaces [info object namespace $obj]
        return $obj
    }
    method openChannelInfo {channel {objname {}}} {
        if {$objname eq ""} {
            set objname evt-chan-info-[incr nameCounter]
        }
        set obj [uplevel 1 [list [namespace which -command EventLogChannelInfo] \
                                create $objname $hSession $channel]]
        lappend dependentNamespaces [info object namespace $obj]
        return $obj
    }
    method openLogfileInfo {logpath {objname {}}} {
        if {$objname eq ""} {
            set objname evt-logfile-info-[incr nameCounter]
        }
        set obj [uplevel 1 [list [namespace which -command EventLogFileInfo] \
                                create $objname $hSession $logpath]]
        lappend dependentNamespaces [info object namespace $obj]
        return $obj
    }
    method openPublisher {publisher {lcid 0} {objname {}}} {
        if {$objname eq ""} {
            set objname evt-publisher-info-[incr nameCounter]
        }
        set obj [uplevel 1 [list [namespace which -command EventLogPublisher] \
                                create $objname $hSession $publisher $lcid ""]]
        lappend dependentNamespaces [info object namespace $obj]
        return $obj
    }
    method openPublisherFromArchive {publisher archive {lcid 0} {objname {}}} {
        if {$objname eq ""} {
            set objname evt-publisher-info-[incr nameCounter]
        }
        set obj [uplevel 1 [list [namespace which -command EventLogPublisher] \
                                create $objname $hSession $publisher $lcid $archive]]
        lappend dependentNamespaces [info object namespace $obj]
        return $obj
    }
}

oo::class create twapi::EventLogPublisher {
    variable hPublisher
    constructor {hsess publisher {lcid 0} {logarchive {}}} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set hPublisher [EvtOpenPublisherMetadata $hsess $publisher $logarchive $lcid 0]
    }
}

oo::class create twapi::EventLogInfo {
    variable hInfo
    constructor {h} {
        set hInfo $h
    }
    method get {args} {
        set result {}
        foreach opt $args {
            lappend result $opt [EvtGetLogInfo $hInfo [dict get {
                -creationtime 0 -lastaccesstime 1 -lastwritetime 2
                -filesize 3 -attributes 4 -numberoflogrecords 5
                -oldestrecordnumber 6 -full 7
            } $opt]]
        }
        return $result
    }
    destructor {
        EvtClose $hInfo
    }
}

oo::class create twapi::EventLogChannelInfo {
    superclass twapi::EventLogInfo
    constructor {hsess channel} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        next [twapi::EvtOpenLog $hsess $channel 1]
    }
}

oo::class create twapi::EventLogFileInfo {
    superclass twapi::EventLogInfo
    constructor {hsess logfile} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        next [twapi::EvtOpenLog $hsess $logfile 1]
    }
}

oo::class create twapi::EventLogChannelConfig {
    variable hConfig
    variable channelName
    classmethod propertyId {property} {
        return [dict getdef {
            -enabled       0 -isolation     1 -type          2 -publisher     3
            -classic       4 -access        5 -logretention  6 -autobackup    7
            -logmaxsize    8 -logfilepath   9 -level        10 -keywords     11
            -controlguid  12 -buffersize   13 -minbuffers   14 -maxbuffers   15
            -latency      16 -clocktype    17 -sidtype      18 -publishers   19
            -filemax      20
        } $property $property]
    }
    constructor {hsess channel} {
        namespace path [linsert [namespace path] 0 [namespace qualifiers [self class]]]
        set channelName $channel
        set hConfig [twapi::EvtOpenChannelConfig $hsess $channel 0]
    }
    destructor {
        EvtClose $hConfig
    }
    method get {propid} {
        return [EvtGetChannelConfigProperty $hConfig \
                    [my propertyId $propid]]
    }
    method set {propid val} {
        return [EvtSetChannelConfigProperty $hConfig \
                    [my propertyId $propid] 0 $val]
    }
    method name {} {
        return $channelName
    }
    method save {} {
        EvtSaveChannelConfig $hConfig
    }
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
