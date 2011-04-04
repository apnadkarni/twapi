#
# Copyright (c) 2004, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# TBD _ maybe more information is available through PDH? Perhaps even on NT?

namespace eval twapi {

    array set IfTypeTokens {
        1  other
        6  ethernet
        9  tokenring
        15 fddi
        23 ppp
        24 loopback
        28 slip
    }

    array set IfOperStatusTokens {
        0 nonoperational
        1 wanunreachable
        2 disconnected
        3 wanconnecting
        4 wanconnected
        5 operational
    }

    # Various pieces of information come from different sources. Moreover,
    # the same information may be available from multiple APIs. In addition
    # older versions of Windows may not have all the APIs. So we try
    # to first get information from older API's whenever we have a choice
    # These tables map fields to positions in the corresponding API result.
    # -1 means rerieving is not as simple as simply indexing into a list

    # GetIfEntry is available from NT4 SP4 onwards
    array set GetIfEntry_opts {
        type                2
        mtu                 3
        speed               4
        physicaladdress     5
        adminstatus         6
        operstatus          7
        laststatuschange    8
        inbytes             9
        inunicastpkts      10
        innonunicastpkts   11
        indiscards         12
        inerrors           13
        inunknownprotocols 14
        outbytes           15
        outunicastpkts     16
        outnonunicastpkts  17
        outdiscards        18
        outerrors          19
        outqlen            20
        description        21
    }

    # GetIpAddrTable also exists in NT4 SP4+
    array set GetIpAddrTable_opts {
        ipaddresses -1
        ifindex     -1
        reassemblysize -1
    }

    # Win2K and up
    array set GetAdaptersInfo_opts {
        adaptername     0
        adapterdescription     1
        adapterindex    3
        dhcpenabled     5
        defaultgateway  7
        dhcpserver      8
        havewins        9
        primarywins    10
        secondarywins  11
        dhcpleasestart 12
        dhcpleaseend   13
    }

    # Win2K and up
    array set GetPerAdapterInfo_opts {
        autoconfigenabled 0
        autoconfigactive  1
        dnsservers        2
    }

    # Win2K and up
    array set GetInterfaceInfo_opts {
        ifname  -1
    }

}

# TBD - Tcl interface to GetIfTable ?

# Get the list of local IP addresses
proc twapi::get_ip_addresses {args} {
    array set opts [parseargs args {
        {ipversion.arg 0}
        {type.arg unicast}
    } -maxleftover 0]

    # 0x20 -> SKIP_FRIENDLYNAME
    set flags 0x27
    if {"all" in $opts(type)} {
        set flags 0x20
    } else {
        if {"unicast" in $opts(type)} {incr flags -1}
        if {"anycast" in $opts(type)} {incr flags -2}
        if {"multicast" in $opts(type)} {incr flags -4}
    }

    set addrs {}
    foreach entry [GetAdaptersAddresses [_ipversion_to_af $opts(ipversion)] $flags] {
        foreach fld {-unicastaddresses -anycastaddresses -multicastaddresses} {
            foreach addrset [kl_get $entry $fld] {
                lappend addrs [kl_get $addrset -address]
            }
        }
    }

    return [lsort -unique $addrs]
}

# Get the list of interfaces
proc twapi::get_netif_indices {} {
    return [lindex [get_network_info -interfaces] 1]
}

proc twapi::get_netif6_indices {} {
    return [twapi::kl_flatten [twapi::GetAdaptersAddresses 23 8] -ipv6ifindex]
}

# Get network related information
proc twapi::get_network_info {args} {
    # Map options into the positions in result of GetNetworkParams
    array set getnetworkparams_opts {
        hostname     0
        domain       1
        dnsservers   2
        dhcpscopeid  4
        routingenabled  5
        arpproxyenabled 6
        dnsenabled      7
    }

    array set opts [parseargs args \
                        [concat [list all ipaddresses interfaces] \
                             [array names getnetworkparams_opts]]]
    set result [list ]
    foreach opt [array names getnetworkparams_opts] {
        if {!$opts(all) && !$opts($opt)} continue
        if {![info exists netparams]} {
            set netparams [GetNetworkParams]
        }
        lappend result -$opt [lindex $netparams $getnetworkparams_opts($opt)]
    }

    if {$opts(all) || $opts(ipaddresses) || $opts(interfaces)} {
        set addrs     [list ]
        set interfaces [list ]
        foreach entry [GetIpAddrTable 0] {
            set addr [lindex $entry 0]
            if {[string compare $addr "0.0.0.0"]} {
                lappend addrs $addr
            }
            lappend interfaces [lindex $entry 1]
        }
        if {$opts(all) || $opts(ipaddresses)} {
            lappend result -ipaddresses $addrs
        }
        if {$opts(all) || $opts(interfaces)} {
            lappend result -interfaces $interfaces
        }
    }

    return $result
}


proc twapi::get_netif_info {interface args} {
    variable IfTypeTokens
    variable GetIfEntry_opts
    variable GetIpAddrTable_opts
    variable GetAdaptersInfo_opts
    variable GetPerAdapterInfo_opts
    variable GetInterfaceInfo_opts

    array set opts [parseargs args \
                        [concat [list all unknownvalue.arg] \
                             [array names GetIfEntry_opts] \
                             [array names GetIpAddrTable_opts] \
                             [array names GetAdaptersInfo_opts] \
                             [array names GetPerAdapterInfo_opts] \
                             [array names GetInterfaceInfo_opts]]]

    array set result [list ]

    set nif $interface
    if {![string is integer $nif]} {
        set nif [GetAdapterIndex $nif]
    }

    if {$opts(all) || $opts(ifindex)} {
        # This really is only useful if $interface had been specified as a name
        set result(-ifindex) $nif
    }

    # Several options need the type so make the GetIfEntry call always
    set values [GetIfEntry $nif]
    foreach opt [array names GetIfEntry_opts] {
        if {$opts(all) || $opts($opt)} {
            set result(-$opt) [lindex $values $GetIfEntry_opts($opt)]
        }
    }
    set type [lindex $values $GetIfEntry_opts(type)]

    if {$opts(all) ||
        [_array_non_zero_entry opts [array names GetIpAddrTable_opts]]} {
        # Collect all the entries, sort by index, then pick out what
        # we want. This assumes there may be multiple entries with the
        # same ifindex
        foreach entry [GetIpAddrTable 0] {
            lassign  $entry  addr ifindex netmask broadcast reasmsize
            lappend ipaddresses($ifindex) [list $addr $netmask $broadcast]
            set reassemblysize($ifindex) $reasmsize
        }
        foreach opt {ipaddresses reassemblysize} {
            if {$opts(all) || $opts($opt)} {
                if {![info exists ${opt}($nif)]} {
                    error "No interface exists with index $nif"
                }
                set result(-$opt) [set ${opt}($nif)]
            }
        }
    }

    if {$opts(all) ||
        [_array_non_zero_entry opts [array names GetAdaptersInfo_opts]]} {
        foreach entry [GetAdaptersInfo] {
            if {$nif != [lindex $entry 3]} continue; # Different interface
            foreach opt [array names GetAdaptersInfo_opts] {
                if {$opts(all) || $opts($opt)} {
                    set result(-$opt) [lindex $entry $GetAdaptersInfo_opts($opt)]
                }
            }
        }
    }

    if {$opts(all) ||
        [_array_non_zero_entry opts [array names GetPerAdapterInfo_opts]]} {
        if {$type == 24} {
            # Loopback - we have to make this info up
            set values {0 0 {}}
        } else {
            set values [GetPerAdapterInfo $nif]
        }
        foreach opt [array names GetPerAdapterInfo_opts] {
            if {$opts(all) || $opts($opt)} {
                set result(-$opt) [lindex $values $GetPerAdapterInfo_opts($opt)]
            }
        }
    }

    if {$opts(all) || $opts(ifname)} {
        array set ifnames [eval concat [GetInterfaceInfo]]
        if {$type == 24} {
            set result(-ifname) "loopback"
        } else {
            if {![info exists ifnames($nif)]} {
                error "No interface exists with index $nif"
            }
            set result(-ifname) $ifnames($nif)
        }
    }

    # Some fields need to be translated to more mnemonic names
    if {[info exists result(-type)]} {
        if {[info exists IfTypeTokens($result(-type))]} {
            set result(-type) $IfTypeTokens($result(-type))
        } else {
            set result(-type) "other"
        }
    }
    if {[info exists result(-physicaladdress)]} {
        set result(-physicaladdress) [_hwaddr_binary_to_string $result(-physicaladdress)]
    }
    foreach opt {-primarywins -secondarywins} {
        if {[info exists result($opt)]} {
            if {[string equal $result($opt) "0.0.0.0"]} {
                set result($opt) ""
            }
        }
    }
    if {[info exists result(-operstatus)] &&
        [info exists twapi::IfOperStatusTokens($result(-operstatus))]} {
        set result(-operstatus) $twapi::IfOperStatusTokens($result(-operstatus))
    }

    return [array get result]
}

proc twapi::get_netif6_info {interface args} {
    array set opts [parseargs args {
        all
        ipv6ifindex
        adaptername
        unicastaddresses
        anycastaddresses
        multicastaddresses
        dnsservers
        dnssuffix
        description
        friendlyname
        physicaladdress
        type
        operstatus
        zoneindices
        prefixes
        dhcpenabled
    } -maxleftover 0]
    
    set haveindex [string is integer -strict $interface]

    set flags 0
    if {! $opts(all)} {
        if {! $opts(unicastaddresses)} { incr flags 0x1 }
        if {! $opts(anycastaddresses)} { incr flags 0x2 }
        if {! $opts(multicastaddresses)} { incr flags 0x4 }
        if {! $opts(dnsservers)} { incr flags 0x8 }
        if {(! $opts(friendlyname)) && $haveindex} { incr flags 0x20 }

        if {$opts(prefixes)} { incr flags 0x10 }
    } else {
        incr flags 0x10;        # Want prefixes also
    }
    
    if {$haveindex} {
        foreach entry [GetAdaptersAddresses 23 $flags] {
            if {[kl_get $entry -ipv6ifindex] == $interface} {
                set found $entry
                break
            }
        }
    } else {
        foreach entry [GetAdaptersAddresses 23 $flags] {
            if {[string equal -nocase [kl_get $entry -adaptername] $interface] ||
                [string equal -nocase [kl_get $entry -friendlyname] $interface]} {
                if {[info exists found]} {
                    error "More than one interface found matching '$interface'"
                }
                set found $entry
            }
        }
    }

    # $found is the matching entry
    if {![info exists found]} {
        error "No interface matching '$interface'."
    }

    set result {}
    foreach opt {
        ipv6ifindex adaptername unicastaddresses anycastaddresses
        multicastaddresses dnsservers dnssuffix description
        friendlyname physicaladdress type operstatus zoneindices prefixes
    } {
        if {$opts(all) || $opts($opt)} {
            lappend result -$opt [kl_get $found -$opt]
        }
    }

    if {$opts(all) || $opts(dhcpenabled)} {
        lappend result -dhcpenabled [expr {([kl_get $found -flags] & 0x4) != 0}]
    }

    return $result
}


# Get the number of network interfaces
proc twapi::get_netif_count {} {
    return [GetNumberOfInterfaces]
}

proc twapi::get_netif6_count {} {
    return [llength [GetAdaptersAddresses 23 8]]
}

# Get the address->h/w address table
proc twapi::get_arp_table {args} {
    array set opts [parseargs args {
        sort
        ifindex.int
        validonly
    }]

    set arps [list ]

    foreach arp [GetIpNetTable $opts(sort)] {
        lassign $arp  ifindex hwaddr ipaddr type
        if {$opts(validonly) && $type == 2} continue
        if {[info exists opts(ifindex)] && $opts(ifindex) != $ifindex} continue
        # Token for enry   0     1      2      3        4
        set type [lindex {other other invalid dynamic static} $type]
        if {$type == ""} {
            set type other
        }
        lappend arps [list $ifindex [_hwaddr_binary_to_string $hwaddr] $ipaddr $type]
    }
    return $arps
}

# Return IP address for a hw address
proc twapi::ipaddr_to_hwaddr {ipaddr {varname ""}} {
    if {![Twapi_IPAddressFamily $ipaddr]} {
        error "$ipaddr is not a valid IP V4 address"
    }

    foreach arp [GetIpNetTable 0] {
        if {[lindex $arp 3] == 2} continue;       # Invalid entry type
        if {[string equal $ipaddr [lindex $arp 2]]} {
            set result [_hwaddr_binary_to_string [lindex $arp 1]]
            break
        }
    }

    # If could not get from ARP table, see if it is one of our own
    # Ignore errors
    if {![info exists result]} {
        foreach ifindex [get_netif_indices] {
            catch {
                array set netifinfo [get_netif_info $ifindex -ipaddresses -physicaladdress]
                # Search list of ipaddresses
                foreach elem $netifinfo(-ipaddresses) {
                    if {[lindex $elem 0] eq $ipaddr && $netifinfo(-physicaladdress) ne ""} {
                        set result $netifinfo(-physicaladdress)
                        break
                    }
                }
            }
            if {[info exists result]} {
                break
            }
        }
    }

    if {[info exists result]} {
        if {$varname == ""} {
            return $result
        }
        upvar $varname var
        set var $result
        return 1
    } else {
        if {$varname == ""} {
            error "Could not map IP address $ipaddr to a hardware address"
        }
        return 0
    }
}

# Return hw address for a IP address
proc twapi::hwaddr_to_ipaddr {hwaddr {varname ""}} {
    set hwaddr [string map {- "" : ""} $hwaddr]
    foreach arp [GetIpNetTable 0] {
        if {[lindex $arp 3] == 2} continue;       # Invalid entry type
        if {[string equal $hwaddr [_hwaddr_binary_to_string [lindex $arp 1] ""]]} {
            set result [lindex $arp 2]
            break
        }
    }

    # If could not get from ARP table, see if it is one of our own
    # Ignore errors
    if {![info exists result]} {
        foreach ifindex [get_netif_indices] {
            catch {
                array set netifinfo [get_netif_info $ifindex -ipaddresses -physicaladdress]
                # Search list of ipaddresses
                set ifhwaddr [string map {- ""} $netifinfo(-physicaladdress)]
                if {[string equal -nocase $hwaddr $ifhwaddr]} {
                    set result [lindex [lindex $netifinfo(-ipaddresses) 0] 0]
                    break
                }
            }
            if {[info exists result]} {
                break
            }
        }
    }

    if {[info exists result]} {
        if {$varname == ""} {
            return $result
        }
        upvar $varname var
        set var $result
        return 1
    } else {
        if {$varname == ""} {
            error "Could not map hardware address $hwaddr to an IP address"
        }
        return 0
    }
}

# Flush the arp table for a given interface
proc twapi::flush_arp_tables {args} {
    if {[llength $args] == 0} {
        set args [get_netif_indices]
    }
    foreach ix $args {
        if {[lindex [get_netif_info $ix -type] 1] ne "loopback"} {
            FlushIpNetTable $ix
        }
    }
}
interp alias {} twapi::flush_arp_table {} twapi::flush_arp_tables

# Return the list of TCP connections
proc twapi::get_tcp_connections {args} {
    variable tcp_statenames
    variable tcp_statevalues
    if {![info exists tcp_statevalues]} {
        array set tcp_statevalues {
            closed            1
            listen            2
            syn_sent          3
            syn_rcvd          4
            estab             5
            fin_wait1         6
            fin_wait2         7
            close_wait        8
            closing           9
            last_ack         10
            time_wait        11
            delete_tcb       12
        }
        foreach {name val} [array get tcp_statevalues] {
            set tcp_statenames($val) $name
        }
    }
    array set opts [parseargs args {
        state
        {ipversion.arg 0}
        localaddr
        remoteaddr
        localport
        remoteport
        pid
        modulename
        modulepath
        bindtime
        all
        matchstate.arg
        matchlocaladdr.arg
        matchremoteaddr.arg
        matchlocalport.int
        matchremoteport.int
        matchpid.int
    } -maxleftover 0]

    set opts(ipversion) [_ipversion_to_af $opts(ipversion)]

    if {! ($opts(state) || $opts(localaddr) || $opts(remoteaddr) || $opts(localport) || $opts(remoteport) || $opts(pid) || $opts(modulename) || $opts(modulepath) || $opts(bindtime))} {
        set opts(all) 1
    }

    # Convert state to appropriate symbol if necessary
    if {[info exists opts(matchstate)]} {
        set matchstates [list ]
        foreach stateval $opts(matchstate) {
            if {[info exists tcp_statevalues($stateval)]} {
                lappend matchstates $stateval
                continue
            }
            if {[info exists tcp_statenames($stateval)]} {
                lappend matchstates $tcp_statenames($stateval)
                continue
            }
            error "Unrecognized connection state '$stateval' specified for option -matchstate"
        }
    }

    foreach opt {matchlocaladdr matchremoteaddr} {
        if {[info exists opts($opt)]} {
            # Note this also normalizes the address format
            set $opt [_hosts_to_ip_addrs $opts($opt)]
            if {[llength [set $opt]] == 0} {
                return [list ]; # No addresses, so no connections will match
            }
        }
    }

    # Get the complete list of connections
    if {$opts(modulename) || $opts(modulepath) || $opts(bindtime) || $opts(all)} {
        set level 8
    } else {
        set level 5
    }
    set conns [list ]
    foreach entry [_get_all_tcp 0 $level $opts(ipversion)] {
        foreach {state localaddr localport remoteaddr remoteport pid bindtime modulename modulepath} $entry {
            break
        }
        if {[string equal $remoteaddr 0.0.0.0]} {
            # Socket not connected. WIndows passes some random value
            # for remote port in this case. Set it to 0
            set remoteport 0
        }
        if {[info exists opts(matchpid)]} {
            # See if this platform even returns the PID
            if {$pid == ""} {
                error "Connection process id not available on this system."
            }
            if {$pid != $opts(matchpid)} {
                continue
            }
        }
        if {[info exists matchlocaladdr] &&
            [lsearch -exact $matchlocaladdr $localaddr] < 0} {
            # Not in match list
            continue
        }
        if {[info exists matchremoteaddr] &&
            [lsearch -exact $matchremoteaddr $remoteaddr] < 0} {
            # Not in match list
            continue
        }
        if {[info exists opts(matchlocalport)] &&
            $opts(matchlocalport) != $localport} {
            continue
        }
        if {[info exists opts(matchremoteport)] &&
            $opts(matchremoteport) != $remoteport} {
            continue
        }
        if {[info exists tcp_statenames($state)]} {
            set state $tcp_statenames($state)
        }
        if {[info exists matchstates] && [lsearch -exact $matchstates $state] < 0} {
            continue
        }

        # OK, now we have matched. Include specified fields in the result
        set conn [list ]
        foreach opt {localaddr localport remoteaddr remoteport state pid bindtime modulename modulepath} {
            if {$opts(all) || $opts($opt)} {
                lappend conn -$opt [set $opt]
            }
        }
        lappend conns $conn
    }
    return $conns
}


# Return the list of UDP connections
proc twapi::get_udp_connections {args} {
    array set opts [parseargs args {
        {ipversion.arg 0}
        localaddr
        localport
        pid
        modulename
        modulepath
        bindtime
        all
        matchlocaladdr.arg
        matchlocalport.int
        matchpid.int
    } -maxleftover 0]

    set opts(ipversion) [_ipversion_to_af $opts(ipversion)]

    if {! ($opts(localaddr) || $opts(localport) || $opts(pid) || $opts(modulename) || $opts(modulepath) || $opts(bindtime))} {
        set opts(all) 1
    }

    if {[info exists opts(matchlocaladdr)]} {
        # Note this also normalizes the address format
        set matchlocaladdr [_hosts_to_ip_addrs $opts(matchlocaladdr)]
        if {[llength $matchlocaladdr] == 0} {
            return [list ]; # No addresses, so no connections will match
        }
    }

    # Get the complete list of connections
    # Get the complete list of connections
    if {$opts(modulename) || $opts(modulepath) || $opts(bindtime) || $opts(all)} {
        set level 2
    } else {
        set level 1
    }
    set conns [list ]
    foreach entry [_get_all_udp 0 $level $opts(ipversion)] {
        foreach {localaddr localport pid bindtime modulename modulepath} $entry {
            break
        }
        if {[info exists opts(matchpid)]} {
            # See if this platform even returns the PID
            if {$pid == ""} {
                error "Connection process id not available on this system."
            }
            if {$pid != $opts(matchpid)} {
                continue
            }
        }
        if {[info exists matchlocaladdr] &&
            [lsearch -exact $matchlocaladdr $localaddr] < 0} {
            continue
        }
        if {[info exists opts(matchlocalport)] &&
            $opts(matchlocalport) != $localport} {
            continue
        }

        # OK, now we have matched. Include specified fields in the result
        set conn [list ]
        foreach opt {localaddr localport pid bindtime modulename modulepath} {
            if {$opts(all) || $opts($opt)} {
                lappend conn -$opt [set $opt]
            }
        }
        lappend conns $conn
    }
    return $conns
}

# Terminates a TCP connection. Does not generate an error if connection
# does not exist
proc twapi::terminate_tcp_connections {args} {
    array set opts [parseargs args {
        matchstate.arg
        matchlocaladdr.arg
        matchremoteaddr.arg
        matchlocalport.int
        matchremoteport.int
        matchpid.int
    } -maxleftover 0]

    # TBD - ignore 'no such connection' errors

    # If local and remote endpoints fully specified, just directly call
    # SetTcpEntry. Note pid must NOT be specified since we must then
    # fall through and check for that pid
    if {[info exists opts(matchlocaladdr)] && [info exists opts(matchlocalport)] &&
        [info exists opts(matchremoteaddr)] && [info exists opts(matchremoteport)] &&
        ! [info exists opts(matchpid)]} {
        # 12 is "delete" code
        catch {
            SetTcpEntry [list 12 $opts(matchlocaladdr) $opts(matchlocalport) $opts(matchremoteaddr) $opts(matchremoteport)]
        }
        return
    }

    # Get connection list and go through matching on each
    foreach conn [eval get_tcp_connections [_get_array_as_options opts]] {
        array set aconn $conn
        # TBD - should we handle integer values of opts(state) ?
        if {[info exists opts(matchstate)] &&
            $opts(matchstate) != $aconn(-state)} {
            continue
        }
        if {[info exists opts(matchlocaladdr)] &&
            $opts(matchlocaladdr) != $aconn(-localaddr)} {
            continue
        }
        if {[info exists opts(matchlocalport)] &&
            $opts(matchlocalport) != $aconn(-localport)} {
            continue
        }
        if {[info exists opts(matchremoteaddr)] &&
            $opts(matchremoteaddr) != $aconn(-remoteaddr)} {
            continue
        }
        if {[info exists opts(remoteport)] &&
            $opts(matchremoteport) != $aconn(-remoteport)} {
            continue
        }
        if {[info exists opts(matchpid)] &&
            $opts(matchpid) != $aconn(-pid)} {
            continue
        }
        # Matching conditions fulfilled
        # 12 is "delete" code
        catch {
            SetTcpEntry [list 12 $aconn(-localaddr) $aconn(-localport) $aconn(-remoteaddr) $aconn(-remoteport)]
        }
    }
    return
}


# Flush cache of host names and ports.
proc twapi::flush_network_name_cache {} {
    array unset ::twapi::port2name
    array unset ::twapi::addr2name
    array unset ::twapi::name2port
    array unset ::twapi::name2addr
}

# IP addr -> hostname
proc twapi::address_to_hostname {addr args} {
    variable addr2name

    array set opts [parseargs args {
        flushcache
        async.arg
    } -maxleftover 0]

    # Note as a special case, we treat 0.0.0.0 explicitly since
    # win32 getnameinfo translates this to the local host name which
    # is completely bogus.
    if {$addr eq "0.0.0.0"} {
        set addr2name($addr) $addr
        set opts(flushcache) 0
        # Now just fall thru to deal with async option etc.
    }


    if {[info exists addr2name($addr)]} {
        if {$opts(flushcache)} {
            unset addr2name($addr)
        } else {
            if {[info exists opts(async)]} {
                after idle [list after 0 $opts(async) [list $addr success $addr2name($addr)]]
                return ""
            } else {
                return $addr2name($addr)
            }
        }
    }

    # If async option, we will call back our internal function which
    # will update the cache and then invoke the caller's script
    if {[info exists opts(async)]} {
        variable _address_handler_scripts
        set id [Twapi_ResolveAddressAsync $addr]
        set _address_handler_scripts($id) [list $addr $opts(async)]
        return ""
    }

    # Synchronous
    set name [lindex [twapi::getnameinfo [list $addr] 8] 0]
    if {$name eq $addr} {
        # Could not resolve.
        set name ""
    }

    set addr2name($addr) $name
    return $name
}

# host name -> IP addresses
proc twapi::hostname_to_address {name args} {
    variable name2addr

    set name [string tolower $name]

    array set opts [parseargs args {
        flushcache
        async.arg
    } -maxleftover 0]

    if {[info exists name2addr($name)]} {
        if {$opts(flushcache)} {
            unset name2addr($name)
        } else {
            if {[info exists opts(async)]} {
                after idle [list after 0 $opts(async) [list $name success $name2addr($name)]]
                return ""
            } else {
                return $name2addr($name)
            }
        }
    }

    # Do not have resolved name

    # If async option, we will call back our internal function which
    # will update the cache and then invoke the caller's script
    if {[info exists opts(async)]} {
        variable _hostname_handler_scripts
        set id [Twapi_ResolveHostnameAsync $name]
        set _hostname_handler_scripts($id) [list $name $opts(async)]
        return ""
    }

    # Resolve address synchronously
    set addrs [list ]
    trap {
        foreach endpt [twapi::getaddrinfo $name 0 0] {
            lassign $endpt addr port
            lappend addrs $addr
        }
    } onerror {TWAPI_WIN32 11001} {
        # Ignore - 11001 -> no such host, so just return empty list
    } onerror {TWAPI_WIN32 11002} {
        # Ignore - 11002 -> no such host, non-authoritative
    } onerror {TWAPI_WIN32 11003} {
        # Ignore - 11001 -> no such host, non recoverable
    } onerror {TWAPI_WIN32 11004} {
        # Ignore - 11004 -> no such host, though valid syntax
    }

    set name2addr($name) $addrs
    return $addrs
}

# Look up a port name
proc twapi::port_to_service {port} {
    variable port2name

    if {[info exists port2name($port)]} {
        return $port2name($port)
    }

    set name ""
    trap {
        set name [lindex [twapi::getnameinfo [list 0.0.0.0 $port] 2] 1]
	if {[string is integer $name] && $name == $port} {
	    # Some platforms return the port itself if no name exists
	    set name ""
	}
    } onerror {TWAPI_WIN32 11001} {
        # Ignore - 11001 -> no such host, so just return empty list
    } onerror {TWAPI_WIN32 11002} {
        # Ignore - 11002 -> no such host, non-authoritative
    } onerror {TWAPI_WIN32 11003} {
        # Ignore - 11001 -> no such host, non recoverable
    } onerror {TWAPI_WIN32 11004} {
        # Ignore - 11004 -> no such host, though valid syntax
    }

    # If we did not get a name back, check for some well known names
    # that windows does not translate. Note some of these are names
    # that windows does translate in the reverse direction!
    if {$name eq ""} {
        foreach {p n} {
            123 ntp
            137 netbios-ns
            138 netbios-dgm
            500 isakmp
            1900 ssdp
            4500 ipsec-nat-t
        } {
            if {$port == $p} {
                set name $n
                break
            }
        }
    }

    set port2name($port) $name
    return $name
}


# Port name -> number
proc twapi::service_to_port {name} {
    variable name2port

    # TBD - add option for specifying protocol
    set protocol 0

    if {[info exists name2port($name)]} {
        return $name2port($name)
    }

    if {[string is integer $name]} {
        return $name
    }

    if {[catch {
        # Return the first port
        set port [lindex [lindex [twapi::getaddrinfo "" $name $protocol] 0] 1]
    }]} {
        set port ""
    }
    set name2port($name) $port
    return $port
}

# Get the routing table
proc twapi::get_routing_table {args} {
    array set opts [parseargs args {
        sort
    } -maxleftover 0]

    set routes [list ]
    foreach route [twapi::GetIpForwardTable $opts(sort)] {
        lappend routes [_format_route $route]
    }

    return $routes
}

# Get the best route for given destination
proc twapi::get_route {args} {
    array set opts [parseargs args {
        {dest.arg 0.0.0.0}
        {source.arg 0.0.0.0}
    } -maxleftover 0]
    return [_format_route [GetBestRoute $opts(dest) $opts(source)]]
}

# Get the interface for a destination
proc twapi::get_outgoing_interface {{dest 0.0.0.0}} {
    return [GetBestInterface $dest]
}

################################################################
# Utility procs

# Convert a route as returned by C code to Tcl format route
proc twapi::_format_route {route} {
    foreach fld {
        addr
        mask
        policy
        nexthop
        ifindex
        type
        protocol
        age
        nexthopas
        metric1
        metric2
        metric3
        metric4
        metric5
    } val $route {
        set r(-$fld) $val
    }

    switch -exact -- $r(-type) {
        2       { set r(-type) invalid }
        3       { set r(-type) local }
        4       { set r(-type) remote }
        1       -
        default { set r(-type) other }
    }

    switch -exact -- $r(-protocol) {
        2 { set r(-protocol) local }
        3 { set r(-protocol) netmgmt }
        4 { set r(-protocol) icmp }
        5 { set r(-protocol) egp }
        6 { set r(-protocol) ggp }
        7 { set r(-protocol) hello }
        8 { set r(-protocol) rip }
        9 { set r(-protocol) is_is }
        10 { set r(-protocol) es_is }
        11 { set r(-protocol) cisco }
        12 { set r(-protocol) bbn }
        13 { set r(-protocol) ospf }
        14 { set r(-protocol) bgp }
        1       -
        default { set r(-protocol) other }
    }

    return [array get r]
}


# Convert binary hardware address to string format
proc twapi::_hwaddr_binary_to_string {b {joiner -}} {
    if {[binary scan $b H* str]} {
        set s ""
        foreach {x y} [split $str ""] {
            lappend s $x$y
        }
        return [join $s $joiner]
    } else {
        error "Could not convert binary hardware address"
    }
}

# Callback for address resolution
proc twapi::_address_resolve_handler {id status hostname} {
    variable _address_handler_scripts

    if {![info exists _address_handler_scripts($id)]} {
        # Queue a background error
        after 0 [list error "Error: No entry found for id $id in address request table"]
        return
    }
    lassign  $_address_handler_scripts($id)  addr script
    unset _address_handler_scripts($id)

    # Before invoking the callback, store result if available
    if {$status eq "success"} {
        set ::twapi::addr2name($addr) $hostname
    }
    eval [linsert $script end $addr $status $hostname]
    return
}

# Callback for hostname resolution
proc twapi::_hostname_resolve_handler {id status addrandports} {
    variable _hostname_handler_scripts

    if {![info exists _hostname_handler_scripts($id)]} {
        # Queue a background error
        after 0 [list error "Error: No entry found for id $id in hostname request table"]
        return
    }
    lassign  $_hostname_handler_scripts($id)  name script
    unset _hostname_handler_scripts($id)

    # Before invoking the callback, store result if available
    if {$status eq "success"} {
        set addrs {}
        foreach addr $addrandports {
            lappend addrs [lindex $addr 0]
        }
        set ::twapi::name2addr($name) $addrs
    } elseif {$addrs == 11001} {
        # For compatibility with the sync version and address resolution,
        # We return an success if empty list if in fact the failure was
        # that no name->address mapping exists
        set status success
        set addrs [list ]
    }

    eval [linsert $script end $name $status $addrs]
    return
}

# Return list of all TCP connections
# Uses GetExtendedTcpTable if available, else AllocateAndGetTcpExTableFromStack
# $level is passed to GetExtendedTcpTable and dtermines format of returned
# data. Level 5 (default) matches what AllocateAndGetTcpExTableFromStack
# returns. Note level 6 and higher is two orders of magnitude more expensive
# to get.
proc twapi::_get_all_tcp {sort level ipversion} {

    if {$ipversion == 0} {
        return [concat [_get_all_tcp $sort $level 2] [_get_all_tcp $sort $level 23]]
    }

    # Get required size of buffer. This also verifies that the
    # GetExtendedTcpTable API exists on this system
    # TBD - modify to do this check only once and not on every call

    if {[catch {twapi::GetExtendedTcpTable NULL 0 $sort $ipversion $level} bufsz]} {
        # No workee, try AllocateAndGetTcpExTableFromStack
        # Note if GetExtendedTcpTable is not present, ipv6 is not
        # available
        if {$ipversion == 2} {
            return [AllocateAndGetTcpExTableFromStack $sort 0]
        } else {
            return {}
        }
    }

    # Allocate the required buffer
    set buf [twapi::malloc $bufsz]
    trap {
        # The required buffer size might change as connections
        # are added or deleted. So we sit in a loop until
        # the required size that we get back from the command
        # is less than or equal to what we supplied
        while {true} {
            set reqsz [twapi::GetExtendedTcpTable $buf $bufsz $sort $ipversion $level]
            if {$reqsz <= $bufsz} {
                # Buffer was large enough. Return the formatted data
                # Note the finally clause below automatically frees
                # the buffer so don't do that here!
                return [Twapi_FormatExtendedTcpTable $buf $ipversion $level]
            }
            # Need bigger buffer
            set bufsz $reqsz
            twapi::free $buf
            unset buf;          # So if malloc fails, we do not free buf again
                                # in the finally clause below
            set buf [twapi::malloc $bufsz]
            # Loop around and try again
        }
    } finally {
        if {[info exists buf]} {
            twapi::free $buf
        }
    }

}

# See comments for _get_all_tcp above except this is for _get_all_udp
proc twapi::_get_all_udp {sort level ipversion} {
    if {$ipversion == 0} {
        return [concat [_get_all_udp $sort $level 2] [_get_all_udp $sort $level 23]]
    }

    # Get required size of buffer. This also verifies that the
    # GetExtendedTcpTable API exists on this system
    if {[catch {twapi::GetExtendedUdpTable NULL 0 $sort $ipversion $level} bufsz]} {
        # No workee, try AllocateAndGetUdpExTableFromStack
        if {$ipversion == 2} {
            return [AllocateAndGetUdpExTableFromStack $sort 0]
        } else {
            return {}
        }
    }

    # Allocate the required buffer
    set buf [twapi::malloc $bufsz]
    trap {
        # The required buffer size might change as connections
        # are added or deleted. So we sit in a loop until
        # the required size that we get back from the command
        # is less than or equal to what we supplied
        while {true} {
            set reqsz [twapi::GetExtendedUdpTable $buf $bufsz $sort $ipversion $level]
            if {$reqsz <= $bufsz} {
                # Buffer was large enough. Return the formatted data
                # Note the finally clause below automatically frees
                # the buffer so don't do that here!
                return [Twapi_FormatExtendedUdpTable $buf $ipversion $level]
            }
            # Need bigger buffer
            set bufsz $reqsz
            twapi::free $buf
            unset buf;          # So if malloc fails, we do not free buf again
                                # in the finally clause below
            set buf [twapi::malloc $bufsz]
            # Loop around and try again
        }
    } finally {
        if {[info exists buf]} {
            twapi::free $buf
        }
    }

}


# valid IP address
proc twapi::_valid_ipaddr_format {ipaddr} {
    return [expr {[Twapi_IPAddressFamily $ipaddr] != 0}]
}

# Given lists of IP addresses and DNS names, returns
# a list purely of IP addresses in normalized form
proc twapi::_hosts_to_ip_addrs hosts {
    set addrs [list ]
    foreach host $hosts {
        if {[_valid_ipaddr_format $host]} {
            lappend addrs [Twapi_NormalizeIPAddress $host]
        } else {
            # Not IP address. Try to resolve, ignoring errors
            if {![catch {hostname_to_address $host -flushcache} hostaddrs]} {
                foreach addr $hostaddrs {
                    lappend addrs [Twapi_NormalizeIPAddress $addr]
                }
            }
        }
    }
    return $addrs
}

proc twapi::_ipversion_to_af {opt} {
    if {[string is integer -strict $opt]} {
        incr opt 0;             # Normalize ints for switch
    }
    switch -exact -- [string tolower $opt] {
        4 -
        inet  { return 2 }
        6 -
        inet6 { return 23 }
        0 -
        all   { return 0 }
    }
    error "Invalid IP version '$opt'"
}
