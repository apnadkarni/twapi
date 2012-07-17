#
# Copyright (c) 2012, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# Event log handling for Vista and later

namespace eval twapi {}

proc twapi::evt_local_session {} {
    return NULL
}

proc twapi::evt_open_session {server args} {
    array set opts [parseargs args {
        user.arg
        domain.arg
        {password.arg ""}
        {authtype.arg 0}
    } -maxleftover 0]

    if {![string is integer -strict $opts(authtype)]} {
        set opts(authtype) [dict get {default 0 negotiate 1 kerberos 2 ntlm 3} [string tolower $opts(authtype)]]
    }

    if {![info exists opts(user)]} {
        set opts(user) $::env(USERNAME)
    }

    if {![info exists opts(domain)]} {
        set opts(domain) $::env(USERDOMAIN)
    }

    return [EvtOpenSession 1 [list $server $opts(user) $opts(domain) $opts(password) $opts(authtype)] 0 0]
}

# proc twapi::evt_close - defined in C

proc twapi::evt_channels {{hevtsess NULL}} {
    set chnames {}
    set hevt [EvtOpenChannelEnum $hevtsess 0]
    trap {
        while {[set chname [EvtNextChannelPath $hevt]] ne ""} {
            lappend chnames $chname
        }
    } finally {
        EvtClose $hevt
    }

    return $chnames
}

proc twapi::evt_clear_log {chanpath args} {
    array set opts [parseargs args {
        {session.arg NULL}
        {backup.arg ""}
    } -maxleftover 0]

    return [EvtClearLog $opts(session) $chanpath $opts(backup) 0]
}

proc twapi::evt_archive_exported_log {logpath args} {
    array set opts [parseargs args {
        {session.arg NULL}
        {lcid.int 0}
    } -maxleftover 0]

    return [EvtArchiveExportedLog $opts(session) $logpath $opts(lcid) 0]
}

proc twapi::evt_export_log {outfile args} {
    array set opts [parseargs args {
        {session.arg NULL}
        file.arg
        channel.arg
        {query.arg *}
        {ignorequeryerrors 0 0x1000}
    } -maxleftover 0]

    if {([info exists opts(file)] && [info exists opts(channel)]) ||
        ! ([info exists opts(file)] || [info exists opts(channel)])} {
        error "Exactly one of -file or -channel must be specified."
    }

    if {[info exists opts(file)]} {
        set path $opts(file)
        incr opts(ignorequeryerrors) 2
    } else {
        set path $opts(channel)
        incr opts(ignorequeryerrors) 1
    }

    return [EvtExportLog $opts(session) $path $opts(query) $outfile $opts(ignorequeryerrors)]
}

proc twapi::evt_create_bookmark {{mark ""}} {
    return [EvtCreateBookmark $mark]
}

proc twapi::evt_render_context_xpaths {xpaths} {
    return [EvtCreateRenderContext $xpaths 0]
}

proc twapi::evt_render_context_system {} {
    return [EvtCreateRenderContext {} 1]
}

proc twapi::evt_render_context_user {} {
    return [EvtCreateRenderContext {} 2]
}

proc twapi::evt_open_channel_config {chanpath args} {
    array set opts [parseargs args {
        {session.arg NULL}
    } -maxleftover 0]

    return [EvtOpenChannelConfig $opts(session) $chanpath 0]
}

proc twapi::evt_get_channel_config {hevt args} {
    set result {}
    foreach opt $args {
        lappend result $opt \
            [EvtGetChannelConfigProperty $hevt \
                 [_evt_map_channel_config_property $hevt $propid]]
    }
    return $result
}

proc twapi::evt_set_channel_config {hevt propid val} {
    return [EvtSetChannelConfigProperty $hevt [_evt_map_channel_config_property $propid 0 $val]]
}


proc twapi::_evt_map_channel_config_property {propid} {
    if {[string is integer -strict $propid]} {
        return $propid
    }
    
    # Note: values are from winevt.h, Win7 SDK has typos for last few
    return [dict get {
        -enabled                  0
        -isolation                1
        -type                     2
        -owningpublisher          3
        -classiceventlog          4
        -access                   5
        -loggingretention         6
        -loggingautobackup        7
        -loggingmaxsize           8
        -logginglogfilepath       9
        -publishinglevel          10
        -publishingkeywords       11
        -publishingcontrolguid    12
        -publishingbuffersize     13
        -publishingminbuffers     14
        -publishingmaxbuffers     15
        -publishinglatency        16
        -publishingclocktype      17
        -publishingsidtype        18
        -publisherlist            19
        -publishingfilemax        20
    } $propid]
}

proc twapi::evt_event_info {hevt args} {
    set result {}
    foreach opt $args {
        lappend result $opt [EvtGetEventInfo $hevt \
                                 [dict get {-queryids 0 -path 1}]]
    }
    return $result
}


proc twapi::evt_event_metadata_property {hevt args} {
    set result {}
    foreach opt $args {
        lappend result $opt \
            [EvtGetEventMetadataProperty $hevt \
                 [dict get {
                     -id 0 version 1 channel 2 level 3
                     -opcode 4 task 5 keyword 6 -messageid 7 template 8
                 } $opt]]
    }
    return $result
}



proc twapi::evt_log_info {hevt args} {
    set result {}
    foreach opt $args {
        lappend result $opt  [EvtGetLogInfo $hevt [dict get {
            -creationtime 0 -lastaccesstime 1 -lastwritetime 2
            -filesize 3 -attributes 4 -numberoflogrecords 5
            -oldestrecordnumber 6 -full 7
        } $opt]]
    }
    return $result
}

proc twapi::evt_publisher_metadata_property {hevt args} {
    set result {}
    foreach opt $args {
        lappend result $opt \
            [EvtGetPublisherMetadataProperty [dict get {
                -publisherguid 0   -resourcefilepath 1 -parameterfilepath 2
                -messagefilepath 3 -helplink 4 -publishermessageid 5
                -channelreferences 6 -channelreferencepath 7
                -channelreferenceindex 8 -channelreferenceid 9
                -channelreferenceflags 10 -channelreferencemessageid 11
                -levels 12 -levelname 13 -levelvalue 14 -levelmessageid 15
                -tasks 16 -taskname 17 -taskeventguid 18 -taskvalue 19
                -taskmessageid 20 -opcodes 21 -opcodename 22
                -opcodevalue 23 -opcodemessageid 24 -keywords 25
                -keywordname 26 -keywordvalue 27 -keywordmessageid 28
            } $opt]]
    }
    return $result
}

proc twapi::evt_query_info {hevt args} {
    set result {}
    foreach opt $args {
        lappend result $opt  [EvtGetQueryInfo $hevt [dict get {
            -names 1 statuses 2
        } $opt]]
    }
    return $result
}

proc twapi::evt_event_array_size {hevt} {
    return [EvtGetObjectArraySize $hevt]
}


proc twapi::evt_publishers {{hevtsess NULL}} {
    set pubs {}
    set hevt [EvtOpenPublisherEnum $hevtsess 0]
    trap {
        while {[set pub [EvtNextPublisherId $hevt]] ne ""} {
            lappend pubs $pub
        }
    } finally {
        EvtClose $hevt
    }

    return $pubs
}




# TBD - EvtFormatMessage
# TBD - EvtGetObjectArrayProperty
# TBD - EvtNext
