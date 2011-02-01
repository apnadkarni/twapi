#
# Copyright (c) 2003-2007, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

#package require twapi

namespace eval twapi {
}


# Returns an keyed list with the following elements:
#   os_major_version
#   os_minor_version
#   os_build_number
#   platform - currently always NT
#   sp_major_version
#   sp_minor_version
#   suites - one or more from backoffice, blade, datacenter, enterprise,
#            smallbusiness, smallbusiness_restricted, terminal, personal
#   system_type - workstation, server
proc twapi::get_os_info {} {
    variable windefs
    variable _osinfo

    if {[info exists _osinfo]} {
        return [array get _osinfo]
    }

    array set verinfo [GetVersionEx]
    set _osinfo(os_major_version) $verinfo(dwMajorVersion)
    set _osinfo(os_minor_version) $verinfo(dwMinorVersion)
    set _osinfo(os_build_number)  $verinfo(dwBuildNumber)
    set _osinfo(platform)         "NT"

    set _osinfo(sp_major_version) $verinfo(wServicePackMajor)
    set _osinfo(sp_minor_version) $verinfo(wServicePackMinor)

    set _osinfo(suites) [list ]
    set suites $verinfo(wSuiteMask)
    foreach suite {
        BACKOFFICE BLADE COMMUNICATIONS COMPUTE_SERVER DATACENTER
        EMBEDDEDNT EMBEDDED_RESTRICTED ENTERPRISE
        PERSONAL SECURITY_APPLIANCE SINGLEUSERTS
        SMALLBUSINESS
        SMALLBUSINESS_RESTRICTED STORAGE_SERVER TERMINAL WH_SERVER
    } {
        set def "VER_SUITE_$suite"
        if {$suites & $windefs($def)} {
            lappend _osinfo(suites) [string tolower $suite]
        }
    }

    set system_type $verinfo(wProductType)
    if {$system_type == $windefs(VER_NT_WORKSTATION)} {
        set _osinfo(system_type) "workstation"
    } elseif {$system_type == $windefs(VER_NT_SERVER)} {
        set _osinfo(system_type) "server"
    } elseif {$system_type == $windefs(VER_NT_DOMAIN_CONTROLLER)} {
        set _osinfo(system_type) "domain_controller"
    } else {
        set _osinfo(system_type) "unknown"
    }

    return [array get _osinfo]
}

# Return a text string describing the OS version and options
# If specified, osinfo should be a keyed list containing
# data returned by get_os_info
proc twapi::get_os_description {} {

    array set osinfo [get_os_info]

    # Assume not terminal server
    set tserver ""

    # Version
    set osversion "$osinfo(os_major_version).$osinfo(os_minor_version)"

    # Base OS name
    if {$osinfo(os_major_version) < 5} {
        set osname "Windows NT"
        if {[string equal $osinfo(system_type) "workstation"]} {
            set systype "Workstation"
        } else {
            if {[lsearch -exact $osinfo(suites) "terminal"] >= 0} {
                set systype "Terminal Server Edition"
            } elseif {[lsearch -exact $osinfo(suites) "enterprise"] >= 0} {
                set systype "Advanced Server"
            } else {
                set systype "Server"
            }
        }
    } else {
        switch -exact -- $osversion {
            "5.0" {
                set osname "Windows 2000"
                if {[string equal $osinfo(system_type) "workstation"]} {
                    set systype "Professional"
                } else {
                    if {[lsearch -exact $osinfo(suites) "datacenter"] >= 0} {
                        set systype "Datacenter Server"
                    } elseif {[lsearch -exact $osinfo(suites) "enterprise"] >= 0} {
                        set systype "Advanced Server"
                    } else {
                        set systype "Server"
                    }
                }
            }
            "5.1" {
                set osname "Windows XP"
                if {[lsearch -exact $osinfo(suites) "personal"] >= 0} {
                    set systype "Home Edition"
                } else {
                    set systype "Professional"
                }
            }
            "5.2" {
                set osname "Windows Server 2003"
                if {[string equal $osinfo(system_type) "workstation"]} {
                    set systype "Professional"
                } else {
                    if {[lsearch -exact $osinfo(suites) "datacenter"] >= 0} {
                        set systype "Datacenter Edition"
                    } elseif {[lsearch -exact $osinfo(suites) "enterprise"] >= 0} {
                        set systype "Enterprise Edition"
                    } elseif {[lsearch -exact $osinfo(suites) "blade"] >= 0} {
                        set systype "Web Edition"
                    } else {
                        set systype "Standard Edition"
                    }
                }
            }
            default {
                # Future release - can't really name, just make something up
                set osname "Windows"
                if {[string equal $osinfo(system_type) "workstation"]} {
                    set systype "Professional"
                } else {
                    set systype "Server"
                }
            }
        }
        if {[lsearch -exact $osinfo(suites) "terminal"] >= 0} {
            set tserver " with Terminal Services"
        }
    }

    # Service pack
    if {$osinfo(sp_major_version) != 0} {
        set spver " Service Pack $osinfo(sp_major_version)"
    } else {
        set spver ""
    }

    return "$osname $systype ${osversion} (Build $osinfo(os_build_number))${spver}${tserver}"
}

# Return major minor servicepack as a quad list
proc twapi::get_os_version {} {
    array set verinfo [GetVersionEx]
    return [list $verinfo(dwMajorVersion) $verinfo(dwMinorVersion) \
                $verinfo(wServicePackMajor) $verinfo(wServicePackMinor)]
}

# Returns true if the OS version is at least $major.$minor.$sp
proc twapi::min_os_version {major {minor 0} {spmajor 0} {spminor 0}} {
    foreach {osmajor osminor osspmajor osspminor} [twapi::get_os_version] {break}

    if {$osmajor > $major} {return 1}
    if {$osmajor < $major} {return 0}
    if {$osminor > $minor} {return 1}
    if {$osminor < $minor} {return 0}
    if {$osspmajor > $spmajor} {return 1}
    if {$osspmajor < $spmajor} {return 0}
    if {$osspminor > $spminor} {return 1}
    if {$osspminor < $spminor} {return 0}

    # Same version, ok
    return 1
}

# Returns proc information
#  $processor should be processor number or "" for "total"
proc twapi::get_processor_info {processor args} {

    if {![info exists ::twapi::get_processor_info_base_opts]} {
        array set ::twapi::get_processor_info_base_opts {
            idletime    IdleTime
            privilegedtime  KernelTime
            usertime    UserTime
            dpctime     DpcTime
            interrupttime InterruptTime
            interrupts    InterruptCount
        }
    }

    # Note the PDH options match those of
    # twapi::get_processor_perf_counter_paths
    set pdh_opts {
        dpcutilization
        interruptutilization
        privilegedutilization
        processorutilization
        userutilization
        dpcrate
        dpcqueuerate
        interruptrate
    }
    # apcbypassrate - does not exist on XP
    # dpcbypassrate - does not exist on XP

    set sysinfo_opts {
        arch
        processorlevel
        processorrev
        processorname
        processormodel
        processorspeed
    }

    array set opts [parseargs args \
                        [concat [list all \
                                     currentprocessorspeed \
                                     [list interval.int 100]] \
                             [array names ::twapi::get_processor_info_base_opts] \
                             $pdh_opts $sysinfo_opts]]

    # Registry lookup for processor description
    # If no processor specified, use 0 under the assumption all processors
    # are the same
    set reg_hwkey "HKEY_LOCAL_MACHINE\\HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\[expr {$processor == "" ? 0 : $processor}]"

    set results [list ]

    set processordata [Twapi_SystemProcessorTimes]
    if {$processor ne ""} {
        if {[llength $processordata] <= $processor} {
            error "Invalid processor number '$processor'"
        }
        array set times [lindex $processordata $processor]
        foreach {opt field} [array get ::twapi::get_processor_info_base_opts] {
            if {$opts(all) || $opts($opt)} {
                lappend results -$opt $times($field)
            }
        }
    } else {
        # Need information across all processors
        foreach instancedata $processordata {
            foreach {opt field} [array get ::twapi::get_processor_info_base_opts] {
                if {[info exists times($field)]} {
                    # We use expr, and not incr here so as to deal with wides
                    set times($field) [expr {wide($times($field)) + [kl_get $instancedata $field]}]
                } else {
                    set times($field) [kl_get $instancedata $field]
                }
            }
            foreach {opt field} [array get ::twapi::get_processor_info_base_opts] {
                if {$opts(all) || $opts($opt)} {
                    lappend results -$opt $times($field)
                }
            }
        }
    }

    if {$opts(all) || $opts(currentprocessorspeed)} {
        # This might fail if counter is not present. We return
        # the rated setting in that case
        if {[catch {
            set ctr_path [make_perf_counter_path ProcessorPerformance "Processor Frequency" -instance Processor_Number_$processor -localize true]
            lappend results -currentprocessorspeed [get_counter_path_value $ctr_path -interval $opts(interval)]
        }]} {
            if {[catch {registry get $reg_hwkey "~MHz"} val]} {
                set val "unknown"
            }
            lappend results -currentprocessorspeed $val
        }
    }
    # Now retrieve each PDH counter
    set requested_opts [list ]
    foreach pdh_opt $pdh_opts {
        if {$opts(all) || $opts($pdh_opt)} {
            lappend requested_opts "-$pdh_opt"
        }
    }


    if {[llength $requested_opts]} {
        set counter_list [eval [list get_perf_processor_counter_paths $processor] \
                          $requested_opts]
        foreach {opt processor value} [get_perf_values_from_metacounter_info $counter_list -interval $opts(interval)] {
            lappend results -$opt $value
        }

    }

    if {$opts(all) || $opts(arch) || $opts(processorlevel) || $opts(processorrev)} {
        set sysinfo [GetSystemInfo]
        if {$opts(all) || $opts(arch)} {
            switch -exact -- [lindex $sysinfo 0] {
                0 {set arch intel}
                6 {set arch ia64}
                9 {set arch amd64}
                10 {set arch ia32_win64}
                default {set arch unknown}
            }
            lappend results -arch $arch
        }

        if {$opts(all) || $opts(processorlevel)} {
            lappend results -processorlevel [lindex $sysinfo 8]
        }

        if {$opts(all) || $opts(processorrev)} {
            lappend results -processorrev [format %x [lindex $sysinfo 9]]
        }
    }

    if {$opts(all) || $opts(processorname)} {
        if {[catch {registry get $reg_hwkey "ProcessorNameString"} val]} {
            set val "unknown"
        }
        lappend results -processorname $val
    }

    if {$opts(all) || $opts(processormodel)} {
        if {[catch {registry get $reg_hwkey "Identifier"} val]} {
            set val "unknown"
        }
        lappend results -processormodel $val
    }

    if {$opts(all) || $opts(processorspeed)} {
        if {[catch {registry get $reg_hwkey "~MHz"} val]} {
            set val "unknown"
        }
        lappend results -processorspeed $val
    }

    return $results
}

# Get number of active processors
proc twapi::get_processor_count {} {
    return [lindex [GetSystemInfo] 5]
}

# Get mask of active processors
proc twapi::get_active_processor_mask {} {
    return [format 0x%x [lindex [GetSystemInfo] 4]]
}

# Get system memory information
proc twapi::get_memory_info {args} {
    array set opts [parseargs args {
        all
        allocationgranularity
        availcommit
        availphysical
        kernelpaged
        kernelnonpaged
        minappaddr
        maxappaddr
        pagesize
        peakcommit
        physicalmemoryload
        processavailcommit
        processcommitlimit
        processtotalvirtual
        processavailvirtual
        swapfiles
        swapfiledetail
        systemcache
        totalcommit
        totalphysical
        usedcommit
    }]

    # TBD - some of the fields are not system fields, but the min of 
    # system and process limits - use GetPerformanceInfo instead - see
    # MSDN

    set results [list ]

    set mem [GlobalMemoryStatus]
    foreach {opt fld} {
        physicalmemoryload     dwMemoryLoad
        totalphysical  ullTotalPhys
        availphysical  ullAvailPhys
        processcommitlimit    ullTotalPageFile
        processavailcommit    ullAvailPageFile
        processtotalvirtual   ullTotalVirtual
        processavailvirtual   ullAvailVirtual
    } {
        if {$opts(all) || $opts($opt)} {
            lappend results -$opt [kl_get $mem $fld]
        }
    }

    if {$opts(all) || $opts(swapfiles) || $opts(swapfiledetail)} {
        set swapfiles [list ]
        set swapdetail [list ]

        foreach item [Twapi_SystemPagefileInformation] {
            array set swap $item
            set swap(FileName) [_normalize_path $swap(FileName)]
            lappend swapfiles $swap(FileName)
            lappend swapdetail $swap(FileName) [list $swap(CurrentSize) $swap(TotalUsed) $swap(PeakUsed)]
        }
        if {$opts(all) || $opts(swapfiles)} {
            lappend results -swapfiles $swapfiles
        }
        if {$opts(all) || $opts(swapfiledetail)} {
            lappend results -swapfiledetail $swapdetail
        }
    }

    if {$opts(all) || $opts(allocationgranularity) ||
        $opts(minappaddr) || $opts(maxappaddr) || $opts(pagesize)} {
        set sysinfo [twapi::GetSystemInfo]
        foreach {opt fmt index} {
            pagesize %u 1 minappaddr 0x%lx 2 maxappaddr 0x%lx 3 allocationgranularity %u 7} {
            if {$opts(all) || $opts($opt)} {
                lappend results -$opt [format $fmt [lindex $sysinfo $index]]
            }
        }
    }

    # This call is slightly expensive so check if it is really needed 
    if {$opts(all) || $opts(totalcommit) || $opts(usedcommit) ||
        $opts(availcommit) ||
        $opts(kernelpaged) || $opts(kernelnonpaged)
    } {
        set mem [GetPerformanceInformation]
        set page_size [kl_get $mem PageSize]
        foreach {opt fld} {
            totalcommit CommitLimit
            usedcommit  CommitTotal
            peakcommit  CommitPeak
            systemcache SystemCache
            kernelpaged KernelPaged
            kernelnonpaged KernelNonpaged
        } {
            if {$opts(all) || $opts($opt)} {
                lappend results -$opt [expr {[kl_get $mem $fld] * $page_size}]
            }
        }
        if {$opts(all) || $opts(availcommit)} {
            lappend results -availcommit [expr {$page_size * ([kl_get $mem CommitLimit]-[kl_get $mem CommitTotal])}]
        }
    }
        
    return $results
}

# Get the netbios name
proc twapi::get_computer_netbios_name {} {
    return [GetComputerName]
}

# Get the computer name
proc twapi::get_computer_name {{typename netbios}} {
    if {[string is integer $typename]} {
        set type $typename
    } else {
        set type [lsearch -exact {netbios dnshostname dnsdomain dnsfullyqualified physicalnetbios physicaldnshostname physicaldnsdomain physicaldnsfullyqualified} $typename]
        if {$type < 0} {
            error "Unknown computer name type '$typename' specified"
        }
    }
    return [GetComputerNameEx $type]
}

# Shut down the system
proc twapi::shutdown_system {args} {
    array set opts [parseargs args {
        system.arg
        {message.arg "System shutdown has been initiated"}
        {timeout.int 60}
        force
        restart
    } -nulldefault]

    eval_with_privileges {
        InitiateSystemShutdown $opts(system) $opts(message) \
            $opts(timeout) $opts(force) $opts(restart)
    } SeShutdownPrivilege
}

# Abort a system shutdown
proc twapi::abort_system_shutdown {args} {
    array set opts [parseargs args {system.arg} -nulldefault]
    eval_with_privileges {
        AbortSystemShutdown $opts(system)
    } SeShutdownPrivilege
}

# Get system uptime
proc twapi::get_system_uptime {} {
    variable _system_start_time
    set now [clock seconds]
    if {![info exists _system_start_time]} {
        set ctr_path [make_perf_counter_path System "System Up Time" -localize true]
        set uptime [get_counter_path_value $ctr_path -interval 0 -format double]
        set _system_start_time [expr {$now - round($uptime+0.5)}]
    }
    return [expr {$now - $_system_start_time}]
}

# Get system information
proc twapi::get_system_info {args} {
    array set opts [parseargs args {
        all
        sid
        uptime
        handlecount
        eventcount
        mutexcount
        processcount
        sectioncount
        semaphorecount
        threadcount
    } -maxleftover 0]

    set result [list ]
    if {$opts(all) || $opts(uptime)} {
        lappend result -uptime [get_system_uptime]
    }

    if {$opts(all) || $opts(sid)} {
        set lsah [get_lsa_policy_handle -access policy_view_local_information]
        trap {
            lappend result -sid [lindex [LsaQueryInformationPolicy $lsah 5] 1]
        } finally {
            close_lsa_policy_handle $lsah
        }
    }

    # If we don't need any PDH based values, return
    # TBD - many of these are available without PDH ? Check and replace
    if {! ($opts(all) || $opts(handlecount) || $opts(processcount) || $opts(threadcount) || $opts(eventcount) || $opts(mutexcount) || $opts(sectioncount) || $opts(semaphorecount))} {
        return $result
    }

    set hquery [open_perf_query]
    trap {
        # Create the counters
        if {$opts(all) || $opts(handlecount)} {
            set handlecount_ctr [add_perf_counter $hquery [make_perf_counter_path Process "Handle Count" -instance _Total -localize true]]
        }
        foreach {opt ctrname} {
            eventcount   Events
            mutexcount   Mutexes
            processcount Processes
            sectioncount Sections
            semaphorecount Semaphores
            threadcount  Threads
        } {
            if {$opts(all) || $opts($opt)} {
                set ${opt}_ctr [add_perf_counter $hquery [make_perf_counter_path Objects $ctrname -localize true]]
            }
        }
        # Collect the data
        collect_perf_query_data $hquery

        foreach opt {
            handlecount
            eventcount
            mutexcount
            processcount
            sectioncount
            semaphorecount
            threadcount
        } {
            if {[info exists ${opt}_ctr]} {
                lappend result -$opt [get_hcounter_value [set ${opt}_ctr] -format long -scale "" -full 0]
            }
        }
    } finally {
        foreach opt {
            handlecount
            eventcount
            mutexcount
            processcount
            sectioncount
            semaphorecount
            threadcount
        } {
            if {[info exists ${opt}_ctr]} {
                remove_perf_counter [set ${opt}_ctr]
            }
        }
        close_perf_query $hquery
    }
    return $result
}

# Map a Windows error code to a string
proc twapi::map_windows_error {code} {
    # Trim trailing CR/LF
    return [string trimright [twapi::Twapi_MapWindowsErrorToString $code] "\r\n"]
}

# Return $s with all environment strings expanded
proc twapi::expand_environment_strings {s} {
    return [ExpandEnvironmentStrings $s]
}

# Load given library
proc twapi::load_library {path args} {
    array set opts [parseargs args {
        dontresolverefs
        datafile
        alteredpath
    }]

    set flags 0
    if {$opts(dontresolverefs)} {
        setbits flags 1;                # DONT_RESOLVE_DLL_REFERENCES
    }
    if {$opts(datafile)} {
        setbits flags 2;                # LOAD_LIBRARY_AS_DATAFILE
    }
    if {$opts(alteredpath)} {
        setbits flags 8;                # LOAD_WITH_ALTERED_SEARCH_PATH
    }

    # LoadLibrary always wants backslashes
    set path [file nativename $path]
    return [LoadLibraryEx $path $flags]
}

# Free library opened with load_library
proc twapi::free_library {libh} {
    FreeLibrary $libh
}


# Format message string
proc twapi::format_message {args} {
    if {[catch {eval _unsafe_format_message $args} result]} {
        set erinfo $::errorInfo
        set ercode $::errorCode
        if {[lindex $ercode 0] == "POSIX" && [lindex $ercode 1] == "EFAULT"} {
            # Number of string params do not match % specifiers
            # Retry without replacing % specifiers
            return [eval _unsafe_format_message -ignoreinserts $args]
        } else {
            error $result $erinfo $ercode
        }
    }

    return $result
}


# Read an ini file int
proc twapi::read_inifile_key {section key args} {
    array set opts [parseargs args {
        {default.arg ""}
        inifile.arg
    } -maxleftover 0]

    if {[info exists opts(inifile)]} {
        set values [read_inifile_section $section -inifile $opts(inifile)]
    } else {
        set values [read_inifile_section $section]
    }

    # Cannot use kl_get or arrays here because we want case insensitive compare
    foreach {k val} $values {
        if {[string equal -nocase $key $k]} {
            return $val
        }
    }
    return $opts(default)
}

# Write an ini file string
proc twapi::write_inifile_key {section key value args} {
    array set opts [parseargs args {
        inifile.arg
    } -maxleftover 0]

    if {[info exists opts(inifile)]} {
        WritePrivateProfileString $section $key $value $opts(inifile)
    } else {
        WriteProfileString $section $key $value
    }
}

# Delete an ini file string
proc twapi::delete_inifile_key {section key args} {
    array set opts [parseargs args {
        inifile.arg
    } -maxleftover 0]

    if {[info exists opts(inifile)]} {
        WritePrivateProfileString $section $key $twapi::nullptr $opts(inifile)
    } else {
        WriteProfileString $section $key $twapi::nullptr
    }
}

# Get names of the sections in an inifile
proc twapi::read_inifile_section_names {args} {
    array set opts [parseargs args {
        inifile.arg
    } -nulldefault -maxleftover 0]

    return [GetPrivateProfileSectionNames $opts(inifile)]
}

# Get keys and values in a section in an inifile
proc twapi::read_inifile_section {section args} {
    array set opts [parseargs args {
        inifile.arg
    } -nulldefault -maxleftover 0]

    set result [list ]
    foreach line [GetPrivateProfileSection $section $opts(inifile)] {
        set pos [string first "=" $line]
        if {$pos >= 0} {
            lappend result [string range $line 0 [expr {$pos-1}]] [string range $line [incr pos] end]
        }
    }
    return $result
}


# Delete an ini file section
proc twapi::delete_inifile_section {section args} {
    variable nullptr

    array set opts [parseargs args {
        inifile.arg
    }]

    if {[info exists opts(inifile)]} {
        WritePrivateProfileString $section $nullptr $nullptr $opts(inifile)
    } else {
        WriteProfileString $section $nullptr $nullptr
    }
}


# Get the primary domain controller
proc twapi::get_primary_domain_controller {args} {
    array set opts [parseargs args {system.arg domain.arg} -nulldefault -maxleftover 0]
    return [NetGetDCName $opts(system) $opts(domain)]
}

# Get a domain controller for a domain
proc twapi::find_domain_controller {args} {
    array set opts [parseargs args {
        system.arg
        avoidself.bool
        domain.arg
        domainguid.arg
        site.arg
        rediscover.bool
        allowstale.bool
        require.arg
        prefer.arg
        justldap.bool
        {inputnameformat.arg any {dns flat netbios any}}
        {outputnameformat.arg any {dns flat netbios any}}
        {outputaddrformat.arg any {ip netbios any}}
        getdetails
    } -maxleftover 0 -nulldefault]


    set flags 0

    if {$opts(outputaddrformat) eq "ip"} {
        setbits flags 0x200
    }

    # Set required bits.
    foreach req $opts(require) {
        if {[string is integer $req]} {
            setbits flags $req
        } else {
            switch -exact -- $req {
                directoryservice { setbits flags 0x10 }
                globalcatalog    { setbits flags 0x40 }
                pdc              { setbits flags 0x80 }
                kdc              { setbits flags 0x400 }
                timeserver       { setbits flags 0x800 }
                writable         { setbits flags 0x1000 }
                default {
                    error "Invalid token '$req' specified in value for option '-require'"
                }
            }
        }
    }

    # Set preferred bits.
    foreach req $opts(prefer) {
        if {[string is integer $req]} {
            setbits flags $req
        } else {
            switch -exact -- $req {
                directoryservice {
                    # If required flag is already set, don't set this
                    if {! ($flags & 0x10)} {
                        setbits flags 0x20
                    }
                }
                timeserver {
                    # If required flag is already set, don't set this
                    if {! ($flags & 0x800)} {
                        setbits flags 0x2000
                    }
                }
                default {
                    error "Invalid token '$req' specified in value for option '-prefer'"
                }
            }
        }
    }

    if {$opts(rediscover)} {
        setbits flags 0x1
    } else {
        # Only look at this option if rediscover is not set
        if {$opts(allowstale)} {
            setbits flags 0x100
        }
    }

    if {$opts(avoidself)} {
        setbits flags 0x4000
    }

    if {$opts(justldap)} {
        setbits flags 0x8000
    }

    switch -exact -- $opts(inputnameformat) {
        any  { }
        netbios -
        flat { setbits flags 0x10000 }
        dns  { setbits flags 0x20000 }
        default {
            error "Invalid value '$opts(inputnameformat)' for option '-inputnameformat'"
        }
    }

    switch -exact -- $opts(outputnameformat) {
        any  { }
        netbios -
        flat { setbits flags 0x80000000 }
        dns  { setbits flags 0x40000000 }
        default {
            error "Invalid value '$opts(outputnameformat)' for option '-outputnameformat'"
        }
    }

    array set dcinfo [DsGetDcName $opts(system) $opts(domain) $opts(domainguid) $opts(site) $flags]

    if {! $opts(getdetails)} {
        return $dcinfo(DomainControllerName)
    }

    set result [list \
                    -dcname $dcinfo(DomainControllerName) \
                    -dcaddr [string trimleft $dcinfo(DomainControllerAddress) \\] \
                    -domainguid $dcinfo(DomainGuid) \
                    -domain $dcinfo(DomainName) \
                    -dnsforest $dcinfo(DnsForestName) \
                    -dcsite $dcinfo(DcSiteName) \
                    -clientsite $dcinfo(ClientSiteName) \
                   ]


    if {$dcinfo(DomainControllerAddressType) == 1} {
        lappend result -dcaddrformat ip
    } else {
        lappend result -dcaddrformat netbios
    }

    if {$dcinfo(Flags) & 0x20000000} {
        lappend result -dcnameformat dns
    } else {
        lappend result -dcnameformat netbios
    }

    if {$dcinfo(Flags) & 0x40000000} {
        lappend result -domainformat dns
    } else {
        lappend result -domainformat netbios
    }

    if {$dcinfo(Flags) & 0x80000000} {
        lappend result -dnsforestformat dns
    } else {
        lappend result -dnsforestformat netbios
    }

    set features [list ]
    foreach {flag feature} {
        0x1    pdc
        0x4    globalcatalog
        0x8    ldap
        0x10   directoryservice
        0x20   kdc
        0x40   timeserver
        0x80   closest
        0x100  writable
        0x200  goodtimeserver
    } {
        if {$dcinfo(Flags) & $flag} {
            lappend features $feature
        }
    }

    lappend result -features $features

    return $result
}



# Get the primary domain info
proc twapi::get_primary_domain_info {args} {
    array set opts [parseargs args {
        all
        name
        dnsdomainname
        dnsforestname
        domainguid
        sid
        type
    } -maxleftover 0]

    set result [list ]
    set lsah [get_lsa_policy_handle -access policy_view_local_information]
    trap {
        foreach {name dnsdomainname dnsforestname domainguid sid} [LsaQueryInformationPolicy $lsah 12] break
        if {[string length $sid] == 0} {
            set type workgroup
            set domainguid ""
        } else {
            set type domain
        }
        foreach opt {name dnsdomainname dnsforestname domainguid sid type} {
            if {$opts(all) || $opts($opt)} {
                lappend result -$opt [set $opt]
            }
        }
    } finally {
        close_lsa_policy_handle $lsah
    }

    return $result
}

# Get the handle for a Tcl channel
proc twapi::get_tcl_channel_handle {chan direction} {
    set direction [expr {[string equal $direction "write"] ? 1 : 0}]
    return [Tcl_GetChannelHandle $chan $direction]
}


# Duplicate a OS handle
proc twapi::duplicate_handle {h args} {
    variable my_process_handle

    array set opts [parseargs args {
        sourcepid.int
        targetpid.int
        access.arg
        inherit
        closesource
    } -maxleftover 0]

    # Assume source and target processes are us
    set source_ph $my_process_handle
    set target_ph $my_process_handle

    if {![string is integer $h]} {
        set h [HANDLE2ADDRESS_LITERAL $h]
    }

    trap {
        set me [pid]
        # If source pid specified and is not us, get a handle to the process
        if {[info exists opts(sourcepid)] && $opts(sourcepid) != $me} {
            set source_ph [get_process_handle $opts(sourcepid) -access process_dup_handle]
        }

        # Ditto for target process...
        if {[info exists opts(targetpid)] && $opts(targetpid) != $me} {
            set target_ph [get_process_handle $opts(targetpid) -access process_dup_handle]
        }

        # Do we want to close the original handle (DUPLICATE_CLOSE_SOURCE)
        set flags [expr {$opts(closesource) ? 0x1: 0}]

        if {[info exists opts(access)]} {
            set access [_access_rights_to_mask $opts(access)]
        } else {
            # If no desired access is indicated, we want the same access as
            # the original handle
            set access 0
            set flags [expr {$flags | 0x2}]; # DUPLICATE_SAME_ACCESS
        }


        set dup [DuplicateHandle $source_ph $h $target_ph $access $opts(inherit) $flags]

        # IF targetpid specified, return handle else literal
        # (even if targetpid is us)
        if {![info exists opts(targetpid)]} {
            set dup [ADDRESS_LITERAL2HANDLE $dup]
        }
    } finally {
        if {$source_ph != $my_process_handle} {
            CloseHandle $source_ph
        }
        if {$target_ph != $my_process_handle} {
            CloseHandle $source_ph
        }
    }

    return $dup
}




# Get a element from SystemParametersInfo
proc twapi::get_system_parameters_info {uiaction} {
    variable SystemParametersInfo_uiactions_get
    # Format of an element is
    #  uiaction_indexvalue uiparam binaryscanstring malloc_size modifiers
    # uiparam may be an int or "sz" in which case the malloc size
    # is substribnuted for it.
    # If modifiers contains "cbsize" the first dword is initialized
    # with malloc_size
    if {![info exists SystemParametersInfo_uiactions_get]} {
        array set SystemParametersInfo_uiactions_get {
            SPI_GETDESKWALLPAPER {0x0073 2048 unicode 4096}
            SPI_GETBEEP  {0x0001 0 i 4}
            SPI_GETMOUSE {0x0003 0 i3 12}
            SPI_GETBORDER {0x0005 0 i 4}
            SPI_GETKEYBOARDSPEED {0x000A 0 i 4}
            SPI_ICONHORIZONTALSPACING {0x000D 0 i 4}
            SPI_GETSCREENSAVETIMEOUT {0x000E 0 i 4}
            SPI_GETSCREENSAVEACTIVE {0x0010 0 i 4}
            SPI_GETKEYBOARDDELAY {0x0016 0 i 4}
            SPI_ICONVERTICALSPACING {0x0018 0 i 4}
            SPI_GETICONTITLEWRAP {0x0019 0 i 4}
            SPI_GETMENUDROPALIGNMENT {0x001B 0 i 4}
            SPI_GETDRAGFULLWINDOWS {0x0026 0 i 4}
            SPI_GETNONCLIENTMETRICS {0x0029 sz i125 500 cbsize}
            SPI_GETMINIMIZEDMETRICS {0x002B sz i5 20 cbsize}
            SPI_GETWORKAREA {0x0030 0 i4 16}
            SPI_GETKEYBOARDPREF {0x0044 0 i 4 }
            SPI_GETSCREENREADER {0x0046 0 i 4}
            SPI_GETANIMATION {0x0048 sz i2 8 cbsize}
            SPI_GETFONTSMOOTHING {0x004A 0 i 4}
            SPI_GETLOWPOWERTIMEOUT {0x004F 0 i 4}
            SPI_GETPOWEROFFTIMEOUT {0x0050 0 i 4}
            SPI_GETLOWPOWERACTIVE {0x0053 0 i 4}
            SPI_GETPOWEROFFACTIVE {0x0054 0 i 4}
            SPI_GETMOUSETRAILS {0x005E 0 i 4}
            SPI_GETSCREENSAVERRUNNING {0x0072 0 i 4}
            SPI_GETFILTERKEYS {0x0032 sz i6 24 cbsize}
            SPI_GETTOGGLEKEYS {0x0034 sz i2 8 cbsize}
            SPI_GETMOUSEKEYS {0x0036 sz i7 28 cbsize}
            SPI_GETSHOWSOUNDS {0x0038 0 i 4}
            SPI_GETSTICKYKEYS {0x003A sz i2 8 cbsize}
            SPI_GETACCESSTIMEOUT {0x003C 12 i3 12 cbsize}
            SPI_GETSNAPTODEFBUTTON {0x005F 0 i 4}
            SPI_GETMOUSEHOVERWIDTH {0x0062 0 i 4}
            SPI_GETMOUSEHOVERHEIGHT {0x0064 0 i 4 }
            SPI_GETMOUSEHOVERTIME {0x0066 0 i 4}
            SPI_GETWHEELSCROLLLINES {0x0068 0 i 4}
            SPI_GETMENUSHOWDELAY {0x006A 0 i 4}
            SPI_GETSHOWIMEUI {0x006E 0 i 4}
            SPI_GETMOUSESPEED {0x0070 0 i 4}
            SPI_GETACTIVEWINDOWTRACKING {0x1000 0 i 4}
            SPI_GETMENUANIMATION {0x1002 0 i 4}
            SPI_GETCOMBOBOXANIMATION {0x1004 0 i 4}
            SPI_GETLISTBOXSMOOTHSCROLLING {0x1006 0 i 4}
            SPI_GETGRADIENTCAPTIONS {0x1008 0 i 4}
            SPI_GETKEYBOARDCUES {0x100A 0 i 4}
            SPI_GETMENUUNDERLINES            {0x100A 0 i 4}
            SPI_GETACTIVEWNDTRKZORDER {0x100C 0 i 4}
            SPI_GETHOTTRACKING {0x100E 0 i 4}
            SPI_GETMENUFADE {0x1012 0 i 4}
            SPI_GETSELECTIONFADE {0x1014 0 i 4}
            SPI_GETTOOLTIPANIMATION {0x1016 0 i 4}
            SPI_GETTOOLTIPFADE {0x1018 0 i 4}
            SPI_GETCURSORSHADOW {0x101A 0 i 4}
            SPI_GETMOUSESONAR {0x101C 0 i 4 }
            SPI_GETMOUSECLICKLOCK {0x101E 0 i 4}
            SPI_GETMOUSEVANISH {0x1020 0 i 4}
            SPI_GETFLATMENU {0x1022 0 i 4}
            SPI_GETDROPSHADOW {0x1024 0 i 4}
            SPI_GETBLOCKSENDINPUTRESETS {0x1026 0 i 4}
            SPI_GETUIEFFECTS {0x103E 0 i 4}
            SPI_GETFOREGROUNDLOCKTIMEOUT {0x2000 0 i 4}
            SPI_GETACTIVEWNDTRKTIMEOUT {0x2002 0 i 4}
            SPI_GETFOREGROUNDFLASHCOUNT {0x2004 0 i 4}
            SPI_GETCARETWIDTH {0x2006 0 i 4}
            SPI_GETMOUSECLICKLOCKTIME {0x2008 0 i 4}
            SPI_GETFONTSMOOTHINGTYPE {0x200A 0 i 4}
            SPI_GETFONTSMOOTHINGCONTRAST {0x200C 0 i 4}
            SPI_GETFOCUSBORDERWIDTH {0x200E 0 i 4}
            SPI_GETFOCUSBORDERHEIGHT {0x2010 0 i 4}
        }
    }

    set key [string toupper $uiaction]

    # TBD -
    # SPI_GETHIGHCONTRAST {0x0042 }
    # SPI_GETSOUNDSENTRY {0x0040 }
    # SPI_GETICONMETRICS {0x002D }
    # SPI_GETICONTITLELOGFONT {0x001F }
    # SPI_GETDEFAULTINPUTLANG {0x0059 }
    # SPI_GETFONTSMOOTHINGORIENTATION {0x2012}

    if {![info exists SystemParametersInfo_uiactions_get($key)]} {
        set key SPI_$key
        if {![info exists SystemParametersInfo_uiactions_get($key)]} {
            error "Unknown SystemParametersInfo index symbol '$uiaction'"
        }
    }

    foreach {index uiparam fmt sz modifiers} $SystemParametersInfo_uiactions_get($key) break
    if {$uiparam eq "sz"} {
        set uiparam $sz
    }
    set mem [malloc $sz]
    trap {
        if {[lsearch -exact $modifiers cbsize] >= 0} {
            # A structure that needs first field set to its size
            Twapi_WriteMemoryBinary $mem 0 $sz [binary format i $sz]
        }
        SystemParametersInfo $index $uiparam $mem 0
        if {$fmt eq "unicode"} {
            set val [Twapi_ReadMemoryUnicode $mem 0 -1]
        } else {
            binary scan [Twapi_ReadMemoryBinary $mem 0 $sz] $fmt val
        }
    } finally {
        free $mem
    }
    return $val
}

proc twapi::set_system_parameters_info {uiaction val args} {
    variable SystemParametersInfo_uiactions_set

    # Format of an element is
    #  uiaction_indexvalue uiparam binaryscanstring malloc_size modifiers
    # uiparam may be an int or "sz" in which case the malloc size
    # is substribnuted for it.
    # If modifiers contains "cbsize" the first dword is initialized
    # with malloc_size
    if {![info exists SystemParametersInfo_uiactions_set]} {
        array set SystemParametersInfo_uiactions_set {
            SPI_SETBEEP                 {0x0002 bool}
            SPI_SETMOUSE                {0x0004 unsupported}
            SPI_SETBORDER               {0x0006 int}
            SPI_SETKEYBOARDSPEED        {0x000B int}
            SPI_ICONHORIZONTALSPACING   {0x000D int}
            SPI_SETSCREENSAVETIMEOUT    {0x000F int}
            SPI_SETSCREENSAVEACTIVE     {0x0011 bool}
            SPI_SETDESKWALLPAPER        {0x0014 unsupported}
            SPI_SETDESKPATTERN          {0x0015 int}
            SPI_SETKEYBOARDDELAY        {0x0017 int}
            SPI_ICONVERTICALSPACING     {0x0018 int}
            SPI_SETICONTITLEWRAP        {0x001A bool}
            SPI_SETMENUDROPALIGNMENT    {0x001C bool}
            SPI_SETDOUBLECLKWIDTH       {0x001D int}
            SPI_SETDOUBLECLKHEIGHT      {0x001E int}
            SPI_SETDOUBLECLICKTIME      {0x0020 int}
            SPI_SETMOUSEBUTTONSWAP      {0x0021 bool}
            SPI_SETICONTITLELOGFONT     {0x0022 LOGFONT}
            SPI_SETDRAGFULLWINDOWS      {0x0025 bool}
            SPI_SETNONCLIENTMETRICS     {0x002A NONCLIENTMETRICS}
            SPI_SETMINIMIZEDMETRICS     {0x002C MINIMIZEDMETRICS}
            SPI_SETICONMETRICS          {0x002E ICONMETRICS}
            SPI_SETWORKAREA             {0x002F RECT}
            SPI_SETPENWINDOWS           {0x0031}
            SPI_SETHIGHCONTRAST         {0x0043 HIGHCONTRAST}
            SPI_SETKEYBOARDPREF         {0x0045 bool}
            SPI_SETSCREENREADER         {0x0047 bool}
            SPI_SETANIMATION            {0x0049 ANIMATIONINFO}
            SPI_SETFONTSMOOTHING        {0x004B bool}
            SPI_SETDRAGWIDTH            {0x004C int}
            SPI_SETDRAGHEIGHT           {0x004D int}
            SPI_SETHANDHELD             {0x004E}
            SPI_SETLOWPOWERTIMEOUT      {0x0051 int}
            SPI_SETPOWEROFFTIMEOUT      {0x0052 int}
            SPI_SETLOWPOWERACTIVE       {0x0055 bool}
            SPI_SETPOWEROFFACTIVE       {0x0056 bool}
            SPI_SETCURSORS              {0x0057 int}
            SPI_SETICONS                {0x0058 int}
            SPI_SETDEFAULTINPUTLANG     {0x005A HKL}
            SPI_SETLANGTOGGLE           {0x005B int}
            SPI_SETMOUSETRAILS          {0x005D int}
            SPI_SETFILTERKEYS          {0x0033 FILTERKEYS}
            SPI_SETTOGGLEKEYS          {0x0035 TOGGLEKEYS}
            SPI_SETMOUSEKEYS           {0x0037 MOUSEKEYS}
            SPI_SETSHOWSOUNDS          {0x0039 bool}
            SPI_SETSTICKYKEYS          {0x003B STICKYKEYS}
            SPI_SETACCESSTIMEOUT       {0x003D ACCESSTIMEOUT}
            SPI_SETSERIALKEYS          {0x003F SERIALKEYS}
            SPI_SETSOUNDSENTRY         {0x0041 SOUNDSENTRY}
            SPI_SETSNAPTODEFBUTTON     {0x0060 bool}
            SPI_SETMOUSEHOVERWIDTH     {0x0063 int}
            SPI_SETMOUSEHOVERHEIGHT    {0x0065 int}
            SPI_SETMOUSEHOVERTIME      {0x0067 int}
            SPI_SETWHEELSCROLLLINES    {0x0069 int}
            SPI_SETMENUSHOWDELAY       {0x006B int}
            SPI_SETSHOWIMEUI          {0x006F bool}
            SPI_SETMOUSESPEED         {0x0071 castint}
            SPI_SETACTIVEWINDOWTRACKING         {0x1001 castbool}
            SPI_SETMENUANIMATION                {0x1003 castbool}
            SPI_SETCOMBOBOXANIMATION            {0x1005 castbool}
            SPI_SETLISTBOXSMOOTHSCROLLING       {0x1007 castbool}
            SPI_SETGRADIENTCAPTIONS             {0x1009 castbool}
            SPI_SETKEYBOARDCUES                 {0x100B castbool}
            SPI_SETMENUUNDERLINES               {0x100B castbool}
            SPI_SETACTIVEWNDTRKZORDER           {0x100D castbool}
            SPI_SETHOTTRACKING                  {0x100F castbool}
            SPI_SETMENUFADE                     {0x1013 castbool}
            SPI_SETSELECTIONFADE                {0x1015 castbool}
            SPI_SETTOOLTIPANIMATION             {0x1017 castbool}
            SPI_SETTOOLTIPFADE                  {0x1019 castbool}
            SPI_SETCURSORSHADOW                 {0x101B castbool}
            SPI_SETMOUSESONAR                   {0x101D castbool}
            SPI_SETMOUSECLICKLOCK               {0x101F bool}
            SPI_SETMOUSEVANISH                  {0x1021 castbool}
            SPI_SETFLATMENU                     {0x1023 castbool}
            SPI_SETDROPSHADOW                   {0x1025 castbool}
            SPI_SETBLOCKSENDINPUTRESETS         {0x1027 bool}
            SPI_SETUIEFFECTS                    {0x103F castbool}
            SPI_SETFOREGROUNDLOCKTIMEOUT        {0x2001 castint}
            SPI_SETACTIVEWNDTRKTIMEOUT          {0x2003 castint}
            SPI_SETFOREGROUNDFLASHCOUNT         {0x2005 castint}
            SPI_SETCARETWIDTH                   {0x2007 castint}
            SPI_SETMOUSECLICKLOCKTIME           {0x2009 int}
            SPI_SETFONTSMOOTHINGTYPE            {0x200B castint}
            SPI_SETFONTSMOOTHINGCONTRAST        {0x200D unsupported}
            SPI_SETFOCUSBORDERWIDTH             {0x200F castint}
            SPI_SETFOCUSBORDERHEIGHT            {0x2011 castint}
        }
    }


    array set opts [parseargs args {
        persist
        notify
    } -nulldefault]

    set flags 0
    if {$opts(persist)} {
        setbits flags 1
    }

    if {$opts(notify)} {
        # Note that actually the notify flag has no effect if persist
        # is not set.
        setbits flags 2
    }

    set key [string toupper $uiaction]

    if {![info exists SystemParametersInfo_uiactions_set($key)]} {
        set key SPI_$key
        if {![info exists SystemParametersInfo_uiactions_set($key)]} {
            error "Unknown SystemParametersInfo index symbol '$uiaction'"
        }
    }

    foreach {index fmt} $SystemParametersInfo_uiactions_set($key) break

    switch -exact -- $fmt {
        int  { SystemParametersInfo $index $val NULL $flags }
        bool {
            set val [expr {$val ? 1 : 0}]
            SystemParametersInfo $index $val NULL $flags
        }
        castint {
            # We have to pass the value as a cast pointer
            SystemParametersInfo $index 0 [Twapi_AddressToPointer $val] $flags
        }
        castbool {
            # We have to pass the value as a cast pointer
            set val [expr {$val ? 1 : 0}]
            SystemParametersInfo $index 0 [Twapi_AddressToPointer $val] $flags
        }
        default {
            error "The data format for $uiaction is not currently supported"
        }
    }

    return
}

################################################################
# Utility procs

# Format message string - will raise exception if insufficient number
# of arguments
proc twapi::_unsafe_format_message {args} {
    array set opts [parseargs args {
        module.arg
        fmtstring.arg
        messageid.arg
        langid.arg
        params.arg
        includesystem
        ignoreinserts
        width.int
    } -nulldefault -maxleftover 0]

    set flags 0

    if {$opts(module) == ""} {
        if {$opts(fmtstring) == ""} {
            # If neither -module nor -fmtstring specified, message is formatted
            # from the system
            set opts(module) NULL
            setbits flags 0x1000;       # FORMAT_MESSAGE_FROM_SYSTEM
        } else {
            setbits flags 0x400;        # FORMAT_MESSAGE_FROM_STRING
            if {$opts(includesystem) || $opts(messageid) != "" || $opts(langid) != ""} {
                error "Options -includesystem, -messageid and -langid cannot be used with -fmtstring"
            }
        }
    } else {
        if {$opts(fmtstring) != ""} {
            error "Options -fmtstring and -module cannot be used together"
        }
        setbits flags 0x800;        # FORMAT_MESSAGE_FROM_HMODULE
        if {$opts(includesystem)} {
            # Also include system in search
            setbits flags 0x1000;       # FORMAT_MESSAGE_FROM_SYSTEM
        }
    }

    if {$opts(ignoreinserts)} {
        setbits flags 0x200;            # FORMAT_MESSAGE_IGNORE_INSERTS
    }

    if {$opts(width) > 254} {
        error "Invalid value for option -width. Must be -1, 0, or a positive integer less than 255"
    }
    if {$opts(width) < 0} {
        # Negative width means no width restrictions
        set opts(width) 255;                  # 255 -> no restrictions
    }
    incr flags $opts(width);                  # Width goes in low byte of flags

    if {$opts(fmtstring) != ""} {
        return [FormatMessageFromString $flags $opts(fmtstring) $opts(params)]
    } else {
        if {![string is integer -strict $opts(messageid)]} {
            error "Unspecified or invalid value for -messageid option. Must be an integer value"
        }
        if {$opts(langid) == ""} { set opts(langid) 0 }
        if {![string is integer -strict $opts(langid)]} {
            error "Unspecfied or invalid value for -langid option. Must be an integer value"
        }

        # Check if $opts(module) is a file or module handle (pointer)
        if {[Twapi_IsPtr $opts(module)]} {
            return  [FormatMessageFromModule $flags $opts(module) \
                         $opts(messageid) $opts(langid) $opts(params)]
        } else {
            set hmod [load_library $opts(module) -datafile]
            trap {
                set message  [FormatMessageFromModule $flags $hmod \
                                  $opts(messageid) $opts(langid) $opts(params)]
            } finally {
                free_library $hmod
            }
            return $message
        }
    }
}


# Helper for Net*Enum type functions taking a common set of arguments
proc twapi::_net_enum_helper {function args} {
    if {[llength $args] == 1} {
        set args [lindex $args 0]
    }

    # -namelevel is used internally to indicate what level is to be used
    # to retrieve names. -preargs and -postargs are used internally to
    # add additional arguments at specific positions in the generic call.
    array set opts [parseargs args {
        {system.arg ""}
        level.int
        resume.int
        filter.int
        {namelevel.int 0}
        {preargs.arg {}}
        {postargs.arg {}}
        {namefield.arg name}
    } -maxleftover 0]

    if {[info exists opts(level)]} {
        set level $opts(level)
    } else {
        set level $opts(namelevel)
    }
    if {[info exists opts(resume)]} {
        set resumehandle $opts(resume)
    } else {
        set resumehandle 0
    }

    set moredata 1
    set result {}
    while {$moredata} {
        if {[info exists opts(filter)]} {
            foreach {moredata resumehandle totalentries groups} [eval [list $function $opts(system)] $opts(preargs) [list $level $opts(filter)] $opts(postargs) [list $resumehandle]] break
        } else {
            foreach {moredata resumehandle totalentries groups} [eval [list $function $opts(system)] $opts(preargs) [list $level] $opts(postargs) [list $resumehandle]] break
        }
        # If caller does not want all data in one lump stop here
        if {[info exists opts(resume)]} {
            if {[info exists opts(level)]} {
                return [list $moredata $resumehandle $totalentries $groups]
            } else {
                # Return flat list of names
                return [list $moredata $resumehandle $totalentries [kl_flatten $groups name]]
            }
        }
        # Append to existing result
        # TBD - can the K operator makes this concatnation faster ?
        set result [concat $result $groups]
    }

    # Return what we have. Format depend on caller options.
    if {[info exists opts(level)]} {
        return $result
    } else {
        return [kl_flatten $result $opts(namefield)]
    }
}
