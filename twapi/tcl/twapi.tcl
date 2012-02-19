#
# Copyright (c) 2003-2012, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

# General definitions and procs used by all TWAPI modules

package require Tcl 8.5
package require registry

namespace eval twapi {
    variable nullptr "__null__"
    variable scriptdir [file dirname [info script]]

    # Accessing global environ in ::env is expensive so cache
    # it. Not updated even if real environ changes
    proc getenv {varname} {
        variable envcache
        set varname [string toupper $varname]
        if {[info exists envcache($varname)]} {
            return $envcache($varname)
        }
        return [set envcache($varname) $::env($varname)]
    }

}

if {![info exists twapi::twapi_base_version]} {
    # set dir $twapi::scriptdir;          # Needed by pkgIndex
    source [file join $twapi::scriptdir twapi_base_version.tcl]
}

# Make twapi versions the same as the base module versions
set twapi::version $::twapi::twapi_base_version
set twapi::patchlevel $::twapi::twapi_base_patchlevel

# log for tracing / debug messages.
proc twapi::debuglog {args} {
    variable log_messages
    if {[llength $args] == 0} {
        if {[info exists log_messages]} {
            return $log_messages
        }
        return [list ]
    }
    foreach msg $args {
        Twapi_AppendLog $msg
    }
}

proc twapi::debuglog_clear {} {
    variable log_messages
    set log_messages {}
}

# Given a proc, wraps some initialization code around it
proc twapi::init_wrapper {procname initcode} {
    set procname [uplevel 1 [list namespace which -command $procname]]
    set arglist {}
    foreach argname [info args $procname] {
        if {[info default $procname $argname argdefault]} {
            lappend arglist [list $argname $argdefault]
        } else {
            lappend arglist $argname
        }
    }
    set body [info body $procname]
    set orig_proc_def [list proc $procname $arglist $body]
    set new_proc_def [format {proc %s args {%s ; %s ; eval [list %s] $args}} $procname $initcode $orig_proc_def $procname]

    uplevel 1 $new_proc_def
}

# Defines a proc with some initialization code
proc twapi::initialized_proc {procname arglist initcode body} {
    set proc_def [format {proc %s {%s} {%s ; proc %s {%s} {%s} ; eval [list %s] [lindex [info level 0] 1]}} $procname $arglist $initcode $procname $arglist $body $procname]
    uplevel 1 $proc_def
}


# Utility proc to load required DLL. Always try the script dir first
# and then the fallback directories. Note the proc is not under th
# twapi:: namespace because we want to load the dll in the caller's
# namespace itself. Huh? why don't we just uplevel the load then? TBD
proc load_twapi_dll {fallback_dirs} {
    if {![info exists ::twapi::dll_base_name]} {
        switch -exact -- $::tcl_platform(machine) {
            intel { set ::twapi::dll_base_name twapi }
            amd64 { set ::twapi::dll_base_name twapi64 }
            default { error "TWAPI not supported on this platform." }
        }
    }

    # If we are a starkit or 8.5 Tcl module, we may need to
    # copy to an external directory before loading

    set tmpdir [pwd]
    catch {set tmpdir $::env(TEMP)}; # Use TEMP if available
    # We do not randomize the directory path since we don't want to
    # clutter up the disk. Unfortunately, there is no easy way of
    # deleting the copied files. Even with atexit type functions
    # the OS will lock the loaded DLLs until process exits.
    # TBD - this here is not a good thing from the security perspective

    # If application has set twapi::temp_dll_dir, that overrides
    # everything.
    if {[info exists twapi::temp_dll_dir]} {
        set tmpdir $twapi::temp_dll_dir
    }

    if {[llength [info commands copy_dll_from_tm]]} {
        set dest [file join $tmpdir "${::twapi::dll_base_name}-${::twapi::build_id}.dll"]
        # We are a running as a tcl 8.5 style Tcl module
        # built using the twapi tools createtmfile.tcl script
        if {![file exists $dest]} {
            file mkdir $tmpdir
            copy_dll_from_tm $dest
        }
        load $dest Twapi
    } elseif {[info exists ::starkit::topdir]} {
        set dest [file join $tmpdir "${::twapi::dll_base_name}-${::twapi::build_id}.dll"]
        if {![file exists $dest]} {
            file mkdir $tmpdir
            file copy [file join $twapi::scriptdir "${::twapi::dll_base_name}.dll"] $dest
        }
        load $dest Twapi
    } else {
        if {[catch {load [file join $twapi::scriptdir "${::twapi::dll_base_name}.dll"]}]} {
            set loaded 0
            foreach dir $fallback_dirs {
                if {[catch {load [file join $dir "${::twapi::dll_base_name}.dll"]}] == 0} {
                    set loaded 1
                    break
                }
            }
            if {! $loaded} {
                error "Could not load ${::twapi::dll_base_name}.dll"
            }
        }
    }
}

proc ::twapi::load_twapi {} {
    if {[llength [info commands GetTwapiBuildInfo]]} {
        return;                 # DLL already loaded or script embedded in dll
    }

    if {[catch {
        if {$::tcl_platform(machine) eq "amd64"} {
            set subdir AMD64
        } else {
            set subdir X86
        }
        load_twapi_dll [list [file join $twapi::scriptdir ../base/build/${subdir}/release]]
    } msg]} {
        set ercode $::errorCode
        set erinfo $::errorInfo
        # Failed to load twapi. Check that dll's we depend on are present
        if {[info exists ::env(SystemRoot)]} {
            set dir $::env(SystemRoot)
        } elseif {[info exists ::env(WINDIR)]} {
            set dir $::env(WINDIR)
        } else {
            # Don't really know where to look. Just pass on original error
            error $msg $erinfo $ercode
        }
        set dir [file join $dir SYSTEM32]
        # TBD -  MSVCP60 no longer needed ?
        foreach dll {
            KERNEL32.dll ADVAPI32.dll USER32.dll RPCRT4.dll
            GDI32.dll PSAPI.DLL NETAPI32.dll pdh.dll WINMM.dll
            MPR.dll WS2_32.dll ole32.dll OLEAUT32.dll SHELL32.dll
            WINSPOOL.DRV VERSION.dll iphlpapi.dll POWRPROF.dll Secur32.dll
            USERENV.dll WTSAPI32.dll SETUPAPI.dll MSVCRT.dll
        } {
            if {![file exists [file join $dir $dll]]} {
                lappend missing $dll
            }
        }
        if {[info exists missing]} {
            set msg "$msg The error might be because the file(s) [join $missing {, }] are missing from the Windows SYSTEM32 directory."
        }
        error $msg $erinfo $ercode
    }
}

twapi::load_twapi

proc twapi::get_build_config {{key ""}} {
    variable build_id
    array set config [GetTwapiBuildInfo]
    if {[info exists build_id]} {
        set config(build_id) $build_id
    } else {
        set config(build_id) 0; # Running from source directory (development)
    }

    # This is actually a runtime config and might not have been initialized
    if {[info exists ::twapi::use_tcloo_for_com]} {
        if {$::twapi::use_tcloo_for_com} {
            set config(comobj_ootype) tcloo
        } else {
            set config(comobj_ootype) metoo
        }
    } else {
        set config(comobj_ootype) uninitialized
    }

    if {$key eq ""} {
        return [array get config]
    } else {
        return $config($key)
    }
}

# Adds the specified Windows header defines into a global array
# We will set a trace to do a lazy initialization on first read of array.
array set twapi::windefs {}
trace add variable twapi::windefs {read array} ::twapi::init_windefs
proc twapi::add_defines {deflist} {
    variable windefs
    array set windefs $deflist
}
proc twapi::init_windefs args {
    variable windefs
    trace remove variable [namespace current]::windefs {read array} ::twapi::init_windefs
    # Redefine ourselves - this is just to guard against coding bugs
    # where we do not remove the trace on the array
    proc ::twapi::init_windefs args {
        error "twapi::init_windefs called multiple times"
    }

    twapi::add_defines {
        VER_NT_WORKSTATION              0x0000001
        VER_NT_DOMAIN_CONTROLLER        0x0000002
        VER_NT_SERVER                   0x0000003

        VER_SERVER_NT                       0x80000000
        VER_WORKSTATION_NT                  0x40000000
        VER_SUITE_SMALLBUSINESS             0x00000001
        VER_SUITE_ENTERPRISE                0x00000002
        VER_SUITE_BACKOFFICE                0x00000004
        VER_SUITE_COMMUNICATIONS            0x00000008
        VER_SUITE_TERMINAL                  0x00000010
        VER_SUITE_SMALLBUSINESS_RESTRICTED  0x00000020
        VER_SUITE_EMBEDDEDNT                0x00000040
        VER_SUITE_DATACENTER                0x00000080
        VER_SUITE_SINGLEUSERTS              0x00000100
        VER_SUITE_PERSONAL                  0x00000200
        VER_SUITE_BLADE                     0x00000400
        VER_SUITE_EMBEDDED_RESTRICTED       0x00000800
        VER_SUITE_SECURITY_APPLIANCE        0x00001000
        VER_SUITE_STORAGE_SERVER            0x00002000
        VER_SUITE_COMPUTE_SERVER            0x00004000
        VER_SUITE_WH_SERVER                 0x00008000

        DELETE                         0x00010000
        READ_CONTROL                   0x00020000
        WRITE_DAC                      0x00040000
        WRITE_OWNER                    0x00080000
        SYNCHRONIZE                    0x00100000

        STANDARD_RIGHTS_REQUIRED       0x000F0000
        STANDARD_RIGHTS_READ           0x00020000
        STANDARD_RIGHTS_WRITE          0x00020000
        STANDARD_RIGHTS_EXECUTE        0x00020000
        STANDARD_RIGHTS_ALL            0x001F0000
        SPECIFIC_RIGHTS_ALL            0x0000FFFF

        GENERIC_READ                   0x80000000
        GENERIC_WRITE                  0x40000000
        GENERIC_EXECUTE                0x20000000
        GENERIC_ALL                    0x10000000

        DESKTOP_READOBJECTS         0x0001
        DESKTOP_CREATEWINDOW        0x0002
        DESKTOP_CREATEMENU          0x0004
        DESKTOP_HOOKCONTROL         0x0008
        DESKTOP_JOURNALRECORD       0x0010
        DESKTOP_JOURNALPLAYBACK     0x0020
        DESKTOP_ENUMERATE           0x0040
        DESKTOP_WRITEOBJECTS        0x0080
        DESKTOP_SWITCHDESKTOP       0x0100

        DF_ALLOWOTHERACCOUNTHOOK    0x0001

        WINSTA_ENUMDESKTOPS         0x0001
        WINSTA_READATTRIBUTES       0x0002
        WINSTA_ACCESSCLIPBOARD      0x0004
        WINSTA_CREATEDESKTOP        0x0008
        WINSTA_WRITEATTRIBUTES      0x0010
        WINSTA_ACCESSGLOBALATOMS    0x0020
        WINSTA_EXITWINDOWS          0x0040
        WINSTA_ENUMERATE            0x0100
        WINSTA_READSCREEN           0x0200

        WINSTA_ALL_ACCESS           0x37f

        FILE_READ_DATA                 0x00000001
        FILE_LIST_DIRECTORY            0x00000001
        FILE_WRITE_DATA                0x00000002
        FILE_ADD_FILE                  0x00000002
        FILE_APPEND_DATA               0x00000004
        FILE_ADD_SUBDIRECTORY          0x00000004
        FILE_CREATE_PIPE_INSTANCE      0x00000004
        FILE_READ_EA                   0x00000008
        FILE_WRITE_EA                  0x00000010
        FILE_EXECUTE                   0x00000020
        FILE_TRAVERSE                  0x00000020
        FILE_DELETE_CHILD              0x00000040
        FILE_READ_ATTRIBUTES           0x00000080
        FILE_WRITE_ATTRIBUTES          0x00000100

        FILE_ALL_ACCESS                0x001F01FF
        FILE_GENERIC_READ              0x00120089
        FILE_GENERIC_WRITE             0x00120116
        FILE_GENERIC_EXECUTE           0x001200A0

        FILE_SHARE_READ                    0x00000001
        FILE_SHARE_WRITE                   0x00000002
        FILE_SHARE_DELETE                  0x00000004

        FILE_ATTRIBUTE_READONLY             0x00000001
        FILE_ATTRIBUTE_HIDDEN               0x00000002
        FILE_ATTRIBUTE_SYSTEM               0x00000004
        FILE_ATTRIBUTE_DIRECTORY            0x00000010
        FILE_ATTRIBUTE_ARCHIVE              0x00000020
        FILE_ATTRIBUTE_DEVICE               0x00000040
        FILE_ATTRIBUTE_NORMAL               0x00000080
        FILE_ATTRIBUTE_TEMPORARY            0x00000100
        FILE_ATTRIBUTE_SPARSE_FILE          0x00000200
        FILE_ATTRIBUTE_REPARSE_POINT        0x00000400
        FILE_ATTRIBUTE_COMPRESSED           0x00000800
        FILE_ATTRIBUTE_OFFLINE              0x00001000
        FILE_ATTRIBUTE_NOT_CONTENT_INDEXED  0x00002000
        FILE_ATTRIBUTE_ENCRYPTED            0x00004000

        FILE_NOTIFY_CHANGE_FILE_NAME    0x00000001
        FILE_NOTIFY_CHANGE_DIR_NAME     0x00000002
        FILE_NOTIFY_CHANGE_ATTRIBUTES   0x00000004
        FILE_NOTIFY_CHANGE_SIZE         0x00000008
        FILE_NOTIFY_CHANGE_LAST_WRITE   0x00000010
        FILE_NOTIFY_CHANGE_LAST_ACCESS  0x00000020
        FILE_NOTIFY_CHANGE_CREATION     0x00000040
        FILE_NOTIFY_CHANGE_SECURITY     0x00000100

        FILE_ACTION_ADDED                   0x00000001
        FILE_ACTION_REMOVED                 0x00000002
        FILE_ACTION_MODIFIED                0x00000003
        FILE_ACTION_RENAMED_OLD_NAME        0x00000004
        FILE_ACTION_RENAMED_NEW_NAME        0x00000005

        FILE_CASE_SENSITIVE_SEARCH      0x00000001
        FILE_CASE_PRESERVED_NAMES       0x00000002
        FILE_UNICODE_ON_DISK            0x00000004
        FILE_PERSISTENT_ACLS            0x00000008
        FILE_FILE_COMPRESSION           0x00000010
        FILE_VOLUME_QUOTAS              0x00000020
        FILE_SUPPORTS_SPARSE_FILES      0x00000040
        FILE_SUPPORTS_REPARSE_POINTS    0x00000080
        FILE_SUPPORTS_REMOTE_STORAGE    0x00000100
        FILE_VOLUME_IS_COMPRESSED       0x00008000
        FILE_SUPPORTS_OBJECT_IDS        0x00010000
        FILE_SUPPORTS_ENCRYPTION        0x00020000
        FILE_NAMED_STREAMS              0x00040000
        FILE_READ_ONLY_VOLUME           0x00080000

        CREATE_NEW          1
        CREATE_ALWAYS       2
        OPEN_EXISTING       3
        OPEN_ALWAYS         4
        TRUNCATE_EXISTING   5


        KEY_QUERY_VALUE                0x00000001
        KEY_SET_VALUE                  0x00000002
        KEY_CREATE_SUB_KEY             0x00000004
        KEY_ENUMERATE_SUB_KEYS         0x00000008
        KEY_NOTIFY                     0x00000010
        KEY_CREATE_LINK                0x00000020
        KEY_WOW64_32KEY                0x00000200
        KEY_WOW64_64KEY                0x00000100
        KEY_WOW64_RES                  0x00000300

        KEY_READ                       0x00020019
        KEY_WRITE                      0x00020006
        KEY_EXECUTE                    0x00020019

        KEY_ALL_ACCESS                 0x000F003F

        SERVICE_QUERY_CONFIG           0x00000001
        SERVICE_CHANGE_CONFIG          0x00000002
        SERVICE_QUERY_STATUS           0x00000004
        SERVICE_ENUMERATE_DEPENDENTS   0x00000008
        SERVICE_START                  0x00000010
        SERVICE_STOP                   0x00000020
        SERVICE_PAUSE_CONTINUE         0x00000040
        SERVICE_INTERROGATE            0x00000080
        SERVICE_USER_DEFINED_CONTROL   0x00000100

        SERVICE_ALL_ACCESS             0x000F01FF

        POLICY_VIEW_LOCAL_INFORMATION   0x00000001
        POLICY_VIEW_AUDIT_INFORMATION   0x00000002
        POLICY_GET_PRIVATE_INFORMATION  0x00000004
        POLICY_TRUST_ADMIN              0x00000008
        POLICY_CREATE_ACCOUNT           0x00000010
        POLICY_CREATE_SECRET            0x00000020
        POLICY_CREATE_PRIVILEGE         0x00000040
        POLICY_SET_DEFAULT_QUOTA_LIMITS 0x00000080
        POLICY_SET_AUDIT_REQUIREMENTS   0x00000100
        POLICY_AUDIT_LOG_ADMIN          0x00000200
        POLICY_SERVER_ADMIN             0x00000400
        POLICY_LOOKUP_NAMES             0x00000800
        POLICY_NOTIFICATION             0x00001000

        POLICY_ALL_ACCESS               0X000F0FFF
        POLICY_READ                     0X00020006
        POLICY_WRITE                    0X000207F8
        POLICY_EXECUTE                  0X00020801


        PROCESS_TERMINATE              0x0001
        PROCESS_CREATE_THREAD          0x0002
        PROCESS_SET_SESSIONID          0x0004
        PROCESS_VM_OPERATION           0x0008
        PROCESS_VM_READ                0x0010
        PROCESS_VM_WRITE               0x0020
        PROCESS_DUP_HANDLE             0x0040
        PROCESS_CREATE_PROCESS         0x0080
        PROCESS_SET_QUOTA              0x0100
        PROCESS_SET_INFORMATION        0x0200
        PROCESS_QUERY_INFORMATION      0x0400
        PROCESS_SUSPEND_RESUME         0x0800

        THREAD_TERMINATE               0x00000001
        THREAD_SUSPEND_RESUME          0x00000002
        THREAD_GET_CONTEXT             0x00000008
        THREAD_SET_CONTEXT             0x00000010
        THREAD_SET_INFORMATION         0x00000020
        THREAD_QUERY_INFORMATION       0x00000040
        THREAD_SET_THREAD_TOKEN        0x00000080
        THREAD_IMPERSONATE             0x00000100
        THREAD_DIRECT_IMPERSONATION    0x00000200
        THREAD_SET_LIMITED_INFORMATION   0x00000400
        THREAD_QUERY_LIMITED_INFORMATION 0x00000800

        EVENT_MODIFY_STATE             0x00000002
        EVENT_ALL_ACCESS               0x001F0003

        SEMAPHORE_MODIFY_STATE         0x00000002
        SEMAPHORE_ALL_ACCESS           0x001F0003

        MUTANT_QUERY_STATE             0x00000001
        MUTANT_ALL_ACCESS              0x001F0001

        MUTEX_MODIFY_STATE             0x00000001
        MUTEX_ALL_ACCESS               0x001F0001

        TIMER_QUERY_STATE              0x00000001
        TIMER_MODIFY_STATE             0x00000002
        TIMER_ALL_ACCESS               0x001F0003

        TOKEN_ASSIGN_PRIMARY           0x00000001
        TOKEN_DUPLICATE                0x00000002
        TOKEN_IMPERSONATE              0x00000004
        TOKEN_QUERY                    0x00000008
        TOKEN_QUERY_SOURCE             0x00000010
        TOKEN_ADJUST_PRIVILEGES        0x00000020
        TOKEN_ADJUST_GROUPS            0x00000040
        TOKEN_ADJUST_DEFAULT           0x00000080
        TOKEN_ADJUST_SESSIONID         0x00000100

        TOKEN_ALL_ACCESS_WINNT         0x000F00FF
        TOKEN_ALL_ACCESS_WIN2K         0x000F01FF
        TOKEN_ALL_ACCESS               0x000F01FF
        TOKEN_READ                     0x00020008
        TOKEN_WRITE                    0x000200E0
        TOKEN_EXECUTE                  0x00020000

        SYSTEM_MANDATORY_LABEL_NO_WRITE_UP         0x1
        SYSTEM_MANDATORY_LABEL_NO_READ_UP          0x2
        SYSTEM_MANDATORY_LABEL_NO_EXECUTE_UP       0x4

        OBJECT_INHERIT_ACE                0x1
        CONTAINER_INHERIT_ACE             0x2
        NO_PROPAGATE_INHERIT_ACE          0x4
        INHERIT_ONLY_ACE                  0x8
        INHERITED_ACE                     0x10
        VALID_INHERIT_FLAGS               0x1F

        ACL_REVISION     2
        ACL_REVISION_DS  4

        ACCESS_ALLOWED_ACE_TYPE                 0x0
        ACCESS_DENIED_ACE_TYPE                  0x1
        SYSTEM_AUDIT_ACE_TYPE                   0x2
        SYSTEM_ALARM_ACE_TYPE                   0x3
        ACCESS_ALLOWED_COMPOUND_ACE_TYPE        0x4
        ACCESS_ALLOWED_OBJECT_ACE_TYPE          0x5
        ACCESS_DENIED_OBJECT_ACE_TYPE           0x6
        SYSTEM_AUDIT_OBJECT_ACE_TYPE            0x7
        SYSTEM_ALARM_OBJECT_ACE_TYPE            0x8
        ACCESS_ALLOWED_CALLBACK_ACE_TYPE        0x9
        ACCESS_DENIED_CALLBACK_ACE_TYPE         0xA
        ACCESS_ALLOWED_CALLBACK_OBJECT_ACE_TYPE 0xB
        ACCESS_DENIED_CALLBACK_OBJECT_ACE_TYPE  0xC
        SYSTEM_AUDIT_CALLBACK_ACE_TYPE          0xD
        SYSTEM_ALARM_CALLBACK_ACE_TYPE          0xE
        SYSTEM_AUDIT_CALLBACK_OBJECT_ACE_TYPE   0xF
        SYSTEM_ALARM_CALLBACK_OBJECT_ACE_TYPE   0x10
        SYSTEM_MANDATORY_LABEL_ACE_TYPE         0x11

        ACCESS_MAX_MS_V2_ACE_TYPE               0x3
        ACCESS_MAX_MS_V3_ACE_TYPE               0x4
        ACCESS_MAX_MS_V4_ACE_TYPE               0x8
        ACCESS_MAX_MS_V5_ACE_TYPE               0x11


        OWNER_SECURITY_INFORMATION              0x00000001
        GROUP_SECURITY_INFORMATION              0x00000002
        DACL_SECURITY_INFORMATION               0x00000004
        SACL_SECURITY_INFORMATION               0x00000008
        LABEL_SECURITY_INFORMATION       	0x00000010

        PROTECTED_DACL_SECURITY_INFORMATION     0x80000000
        PROTECTED_SACL_SECURITY_INFORMATION     0x40000000
        UNPROTECTED_DACL_SECURITY_INFORMATION   0x20000000
        UNPROTECTED_SACL_SECURITY_INFORMATION   0x10000000

        TokenUser                      1
        TokenGroups                    2
        TokenPrivileges                3
        TokenOwner                     4
        TokenPrimaryGroup              5
        TokenDefaultDacl               6
        TokenSource                    7
        TokenType                      8
        TokenImpersonationLevel        9
        TokenStatistics               10
        TokenRestrictedSids           11
        TokenSessionId                12
        TokenGroupsAndPrivileges      13
        TokenSessionReference         14
        TokenSandBoxInert             15
        TokenAuditPolicy              16
        TokenOrigin                   17
        TokenElevationType            18
        TokenLinkedToken              19
        TokenElevation                20
        TokenHasRestrictions          21
        TokenAccessInformation        22
        TokenVirtualizationAllowed    23
        TokenVirtualizationEnabled    24
        TokenIntegrityLevel           25
        TokenUIAccess                 26
        TokenMandatoryPolicy          27
        TokenLogonSid                 28

        SE_GROUP_MANDATORY              0x00000001
        SE_GROUP_ENABLED_BY_DEFAULT     0x00000002
        SE_GROUP_ENABLED                0x00000004
        SE_GROUP_OWNER                  0x00000008
        SE_GROUP_USE_FOR_DENY_ONLY      0x00000010
        SE_GROUP_LOGON_ID               0xC0000000
        SE_GROUP_RESOURCE               0x20000000
        SE_GROUP_INTEGRITY              0x00000020
        SE_GROUP_INTEGRITY_ENABLED      0x00000040

        SE_PRIVILEGE_ENABLED_BY_DEFAULT 0x00000001
        SE_PRIVILEGE_ENABLED            0x00000002
        SE_PRIVILEGE_USED_FOR_ACCESS    0x80000000

        SC_MANAGER_CONNECT             0x00000001
        SC_MANAGER_CREATE_SERVICE      0x00000002
        SC_MANAGER_ENUMERATE_SERVICE   0x00000004
        SC_MANAGER_LOCK                0x00000008
        SC_MANAGER_QUERY_LOCK_STATUS   0x00000010
        SC_MANAGER_MODIFY_BOOT_CONFIG  0x00000020
        SC_MANAGER_ALL_ACCESS          0x000F003F

        SERVICE_NO_CHANGE              0xffffffff

        SERVICE_KERNEL_DRIVER          0x00000001
        SERVICE_FILE_SYSTEM_DRIVER     0x00000002
        SERVICE_ADAPTER                0x00000004
        SERVICE_RECOGNIZER_DRIVER      0x00000008
        SERVICE_WIN32_OWN_PROCESS      0x00000010
        SERVICE_WIN32_SHARE_PROCESS    0x00000020

        SERVICE_INTERACTIVE_PROCESS    0x00000100

        SERVICE_BOOT_START             0x00000000
        SERVICE_SYSTEM_START           0x00000001
        SERVICE_AUTO_START             0x00000002
        SERVICE_DEMAND_START           0x00000003
        SERVICE_DISABLED               0x00000004

        SERVICE_ERROR_IGNORE           0x00000000
        SERVICE_ERROR_NORMAL           0x00000001
        SERVICE_ERROR_SEVERE           0x00000002
        SERVICE_ERROR_CRITICAL         0x00000003

        SERVICE_CONTROL_STOP                   0x00000001
        SERVICE_CONTROL_PAUSE                  0x00000002
        SERVICE_CONTROL_CONTINUE               0x00000003
        SERVICE_CONTROL_INTERROGATE            0x00000004
        SERVICE_CONTROL_SHUTDOWN               0x00000005
        SERVICE_CONTROL_PARAMCHANGE            0x00000006
        SERVICE_CONTROL_NETBINDADD             0x00000007
        SERVICE_CONTROL_NETBINDREMOVE          0x00000008
        SERVICE_CONTROL_NETBINDENABLE          0x00000009
        SERVICE_CONTROL_NETBINDDISABLE         0x0000000A
        SERVICE_CONTROL_DEVICEEVENT            0x0000000B
        SERVICE_CONTROL_HARDWAREPROFILECHANGE  0x0000000C
        SERVICE_CONTROL_POWEREVENT             0x0000000D
        SERVICE_CONTROL_SESSIONCHANGE          0x0000000E

        SERVICE_ACTIVE                 0x00000001
        SERVICE_INACTIVE               0x00000002
        SERVICE_STATE_ALL              0x00000003

        SERVICE_STOPPED                        0x00000001
        SERVICE_START_PENDING                  0x00000002
        SERVICE_STOP_PENDING                   0x00000003
        SERVICE_RUNNING                        0x00000004
        SERVICE_CONTINUE_PENDING               0x00000005
        SERVICE_PAUSE_PENDING                  0x00000006
        SERVICE_PAUSED                         0x00000007

        GA_PARENT       1
        GA_ROOT         2
        GA_ROOTOWNER    3

        GW_HWNDFIRST        0
        GW_HWNDLAST         1
        GW_HWNDNEXT         2
        GW_HWNDPREV         3
        GW_OWNER            4
        GW_CHILD            5
        GW_ENABLEDPOPUP     6

        GWL_WNDPROC         -4
        GWL_HINSTANCE       -6
        GWL_HWNDPARENT      -8
        GWL_STYLE           -16
        GWL_EXSTYLE         -20
        GWL_USERDATA        -21
        GWL_ID              -12

        SW_HIDE             0
        SW_SHOWNORMAL       1
        SW_NORMAL           1
        SW_SHOWMINIMIZED    2
        SW_SHOWMAXIMIZED    3
        SW_MAXIMIZE         3
        SW_SHOWNOACTIVATE   4
        SW_SHOW             5
        SW_MINIMIZE         6
        SW_SHOWMINNOACTIVE  7
        SW_SHOWNA           8
        SW_RESTORE          9
        SW_SHOWDEFAULT      10
        SW_FORCEMINIMIZE    11

        WS_OVERLAPPED       0x00000000
        WS_TILED            0x00000000
        WS_POPUP            0x80000000
        WS_CHILD            0x40000000
        WS_MINIMIZE         0x20000000
        WS_ICONIC           0x20000000
        WS_VISIBLE          0x10000000
        WS_DISABLED         0x08000000
        WS_CLIPSIBLINGS     0x04000000
        WS_CLIPCHILDREN     0x02000000
        WS_MAXIMIZE         0x01000000
        WS_BORDER           0x00800000
        WS_DLGFRAME         0x00400000
        WS_CAPTION          0x00C00000
        WS_VSCROLL          0x00200000
        WS_HSCROLL          0x00100000
        WS_SYSMENU          0x00080000
        WS_THICKFRAME       0x00040000
        WS_SIZEBOX          0x00040000
        WS_GROUP            0x00020000
        WS_TABSTOP          0x00010000

        WS_MINIMIZEBOX      0x00020000
        WS_MAXIMIZEBOX      0x00010000

        WS_EX_DLGMODALFRAME     0x00000001
        WS_EX_NOPARENTNOTIFY    0x00000004
        WS_EX_TOPMOST           0x00000008
        WS_EX_ACCEPTFILES       0x00000010
        WS_EX_TRANSPARENT       0x00000020
        WS_EX_MDICHILD          0x00000040
        WS_EX_TOOLWINDOW        0x00000080
        WS_EX_WINDOWEDGE        0x00000100
        WS_EX_CLIENTEDGE        0x00000200
        WS_EX_CONTEXTHELP       0x00000400

        WS_EX_RIGHT             0x00001000
        WS_EX_LEFT              0x00000000
        WS_EX_RTLREADING        0x00002000
        WS_EX_LTRREADING        0x00000000
        WS_EX_LEFTSCROLLBAR     0x00004000
        WS_EX_RIGHTSCROLLBAR    0x00000000

        WS_EX_CONTROLPARENT     0x00010000
        WS_EX_STATICEDGE        0x00020000
        WS_EX_APPWINDOW         0x00040000

        CS_VREDRAW          0x0001
        CS_HREDRAW          0x0002
        CS_DBLCLKS          0x0008
        CS_OWNDC            0x0020
        CS_CLASSDC          0x0040
        CS_PARENTDC         0x0080
        CS_NOCLOSE          0x0200
        CS_SAVEBITS         0x0800
        CS_BYTEALIGNCLIENT  0x1000
        CS_BYTEALIGNWINDOW  0x2000
        CS_GLOBALCLASS      0x4000

        SWP_NOSIZE          0x0001
        SWP_NOMOVE          0x0002
        SWP_NOZORDER        0x0004
        SWP_NOREDRAW        0x0008
        SWP_NOACTIVATE      0x0010
        SWP_FRAMECHANGED    0x0020
        SWP_DRAWFRAME       0x0020
        SWP_SHOWWINDOW      0x0040
        SWP_HIDEWINDOW      0x0080
        SWP_NOCOPYBITS      0x0100
        SWP_NOOWNERZORDER   0x0200
        SWP_NOREPOSITION    0x0200
        SWP_NOSENDCHANGING  0x0400

        SWP_DEFERERASE      0x2000
        SWP_ASYNCWINDOWPOS  0x4000

        SMTO_NORMAL         0x0000
        SMTO_BLOCK          0x0001
        SMTO_ABORTIFHUNG    0x0002

        HWND_TOP         0
        HWND_BOTTOM      1
        HWND_TOPMOST    -1
        HWND_NOTOPMOST  -2


        WM_NULL                         0x0000
        WM_CREATE                       0x0001
        WM_DESTROY                      0x0002
        WM_MOVE                         0x0003
        WM_SIZE                         0x0005
        WM_ACTIVATE                     0x0006
        WM_SETFOCUS                     0x0007
        WM_KILLFOCUS                    0x0008
        WM_ENABLE                       0x000A
        WM_SETREDRAW                    0x000B
        WM_SETTEXT                      0x000C
        WM_GETTEXT                      0x000D
        WM_GETTEXTLENGTH                0x000E
        WM_PAINT                        0x000F
        WM_CLOSE                        0x0010
        WM_QUERYENDSESSION              0x0011
        WM_QUERYOPEN                    0x0013
        WM_ENDSESSION                   0x0016
        WM_QUIT                         0x0012
        WM_ERASEBKGND                   0x0014
        WM_SYSCOLORCHANGE               0x0015
        WM_SHOWWINDOW                   0x0018
        WM_WININICHANGE                 0x001A
        WM_SETTINGCHANGE                WM_WININICHANGE
        WM_DEVMODECHANGE                0x001B
        WM_ACTIVATEAPP                  0x001C
        WM_FONTCHANGE                   0x001D
        WM_TIMECHANGE                   0x001E
        WM_CANCELMODE                   0x001F
        WM_SETCURSOR                    0x0020
        WM_MOUSEACTIVATE                0x0021
        WM_CHILDACTIVATE                0x0022
        WM_QUEUESYNC                    0x0023
        WM_GETMINMAXINFO                0x0024

        PERF_DETAIL_NOVICE          100
        PERF_DETAIL_ADVANCED        200
        PERF_DETAIL_EXPERT          300
        PERF_DETAIL_WIZARD          400

        PDH_FMT_RAW     0x00000010
        PDH_FMT_ANSI    0x00000020
        PDH_FMT_UNICODE 0x00000040
        PDH_FMT_LONG    0x00000100
        PDH_FMT_DOUBLE  0x00000200
        PDH_FMT_LARGE   0x00000400
        PDH_FMT_NOSCALE 0x00001000
        PDH_FMT_1000    0x00002000
        PDH_FMT_NODATA  0x00004000
        PDH_FMT_NOCAP100 0x00008000

        PERF_DETAIL_COSTLY   0x00010000
        PERF_DETAIL_STANDARD 0x0000FFFF

        UF_SCRIPT                          0x0001
        UF_ACCOUNTDISABLE                  0x0002
        UF_HOMEDIR_REQUIRED                0x0008
        UF_LOCKOUT                         0x0010
        UF_PASSWD_NOTREQD                  0x0020
        UF_PASSWD_CANT_CHANGE              0x0040
        UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED 0x0080
        UF_TEMP_DUPLICATE_ACCOUNT       0x0100
        UF_NORMAL_ACCOUNT               0x0200
        UF_INTERDOMAIN_TRUST_ACCOUNT    0x0800
        UF_WORKSTATION_TRUST_ACCOUNT    0x1000
        UF_SERVER_TRUST_ACCOUNT         0x2000
        UF_DONT_EXPIRE_PASSWD           0x10000
        UF_MNS_LOGON_ACCOUNT            0x20000
        UF_SMARTCARD_REQUIRED           0x40000
        UF_TRUSTED_FOR_DELEGATION       0x80000
        UF_NOT_DELEGATED               0x100000
        UF_USE_DES_KEY_ONLY            0x200000
        UF_DONT_REQUIRE_PREAUTH        0x400000
        UF_PASSWORD_EXPIRED            0x800000
        UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION 0x1000000

        FILE_CASE_PRESERVED_NAMES       0x00000002
        FILE_UNICODE_ON_DISK            0x00000004
        FILE_PERSISTENT_ACLS            0x00000008
        FILE_FILE_COMPRESSION           0x00000010
        FILE_VOLUME_QUOTAS              0x00000020
        FILE_SUPPORTS_SPARSE_FILES      0x00000040
        FILE_SUPPORTS_REPARSE_POINTS    0x00000080
        FILE_SUPPORTS_REMOTE_STORAGE    0x00000100
        FILE_VOLUME_IS_COMPRESSED       0x00008000
        FILE_SUPPORTS_OBJECT_IDS        0x00010000
        FILE_SUPPORTS_ENCRYPTION        0x00020000
        FILE_NAMED_STREAMS              0x00040000
        FILE_READ_ONLY_VOLUME           0x00080000

        KEYEVENTF_EXTENDEDKEY 0x0001
        KEYEVENTF_KEYUP       0x0002
        KEYEVENTF_UNICODE     0x0004
        KEYEVENTF_SCANCODE    0x0008

        MOUSEEVENTF_MOVE        0x0001
        MOUSEEVENTF_LEFTDOWN    0x0002
        MOUSEEVENTF_LEFTUP      0x0004
        MOUSEEVENTF_RIGHTDOWN   0x0008
        MOUSEEVENTF_RIGHTUP     0x0010
        MOUSEEVENTF_MIDDLEDOWN  0x0020
        MOUSEEVENTF_MIDDLEUP    0x0040
        MOUSEEVENTF_XDOWN       0x0080
        MOUSEEVENTF_XUP         0x0100
        MOUSEEVENTF_WHEEL       0x0800
        MOUSEEVENTF_VIRTUALDESK 0x4000
        MOUSEEVENTF_ABSOLUTE    0x8000

        XBUTTON1      0x0001
        XBUTTON2      0x0002

        VK_BACK           0x08
        VK_TAB            0x09
        VK_CLEAR          0x0C
        VK_RETURN         0x0D
        VK_SHIFT          0x10
        VK_CONTROL        0x11
        VK_MENU           0x12
        VK_PAUSE          0x13
        VK_CAPITAL        0x14
        VK_KANA           0x15
        VK_HANGEUL        0x15
        VK_HANGUL         0x15
        VK_JUNJA          0x17
        VK_FINAL          0x18
        VK_HANJA          0x19
        VK_KANJI          0x19
        VK_ESCAPE         0x1B
        VK_CONVERT        0x1C
        VK_NONCONVERT     0x1D
        VK_ACCEPT         0x1E
        VK_MODECHANGE     0x1F
        VK_SPACE          0x20
        VK_PRIOR          0x21
        VK_NEXT           0x22
        VK_END            0x23
        VK_HOME           0x24
        VK_LEFT           0x25
        VK_UP             0x26
        VK_RIGHT          0x27
        VK_DOWN           0x28
        VK_SELECT         0x29
        VK_PRINT          0x2A
        VK_EXECUTE        0x2B
        VK_SNAPSHOT       0x2C
        VK_INSERT         0x2D
        VK_DELETE         0x2E
        VK_HELP           0x2F
        VK_LWIN           0x5B
        VK_RWIN           0x5C
        VK_APPS           0x5D
        VK_SLEEP          0x5F
        VK_NUMPAD0        0x60
        VK_NUMPAD1        0x61
        VK_NUMPAD2        0x62
        VK_NUMPAD3        0x63
        VK_NUMPAD4        0x64
        VK_NUMPAD5        0x65
        VK_NUMPAD6        0x66
        VK_NUMPAD7        0x67
        VK_NUMPAD8        0x68
        VK_NUMPAD9        0x69
        VK_MULTIPLY       0x6A
        VK_ADD            0x6B
        VK_SEPARATOR      0x6C
        VK_SUBTRACT       0x6D
        VK_DECIMAL        0x6E
        VK_DIVIDE         0x6F
        VK_F1             0x70
        VK_F2             0x71
        VK_F3             0x72
        VK_F4             0x73
        VK_F5             0x74
        VK_F6             0x75
        VK_F7             0x76
        VK_F8             0x77
        VK_F9             0x78
        VK_F10            0x79
        VK_F11            0x7A
        VK_F12            0x7B
        VK_F13            0x7C
        VK_F14            0x7D
        VK_F15            0x7E
        VK_F16            0x7F
        VK_F17            0x80
        VK_F18            0x81
        VK_F19            0x82
        VK_F20            0x83
        VK_F21            0x84
        VK_F22            0x85
        VK_F23            0x86
        VK_F24            0x87
        VK_NUMLOCK        0x90
        VK_SCROLL         0x91
        VK_LSHIFT         0xA0
        VK_RSHIFT         0xA1
        VK_LCONTROL       0xA2
        VK_RCONTROL       0xA3
        VK_LMENU          0xA4
        VK_RMENU          0xA5
        VK_BROWSER_BACK        0xA6
        VK_BROWSER_FORWARD     0xA7
        VK_BROWSER_REFRESH     0xA8
        VK_BROWSER_STOP        0xA9
        VK_BROWSER_SEARCH      0xAA
        VK_BROWSER_FAVORITES   0xAB
        VK_BROWSER_HOME        0xAC
        VK_VOLUME_MUTE         0xAD
        VK_VOLUME_DOWN         0xAE
        VK_VOLUME_UP           0xAF
        VK_MEDIA_NEXT_TRACK    0xB0
        VK_MEDIA_PREV_TRACK    0xB1
        VK_MEDIA_STOP          0xB2
        VK_MEDIA_PLAY_PAUSE    0xB3
        VK_LAUNCH_MAIL         0xB4
        VK_LAUNCH_MEDIA_SELECT 0xB5
        VK_LAUNCH_APP1         0xB6
        VK_LAUNCH_APP2         0xB7

        SND_SYNC            0x0000
        SND_ASYNC           0x0001
        SND_NODEFAULT       0x0002
        SND_MEMORY          0x0004
        SND_LOOP            0x0008
        SND_NOSTOP          0x0010
        SND_NOWAIT      0x00002000
        SND_ALIAS       0x00010000
        SND_ALIAS_ID    0x00110000
        SND_FILENAME    0x00020000
        SND_RESOURCE    0x00040004
        SND_PURGE           0x0040
        SND_APPLICATION     0x0080


        STYPE_DISKTREE          0
        STYPE_PRINTQ            1
        STYPE_DEVICE            2
        STYPE_IPC               3
        STYPE_TEMPORARY         0x40000000
        STYPE_SPECIAL           0x80000000

        LOGON32_LOGON_INTERACTIVE       2
        LOGON32_LOGON_NETWORK           3
        LOGON32_LOGON_BATCH             4
        LOGON32_LOGON_SERVICE           5
        LOGON32_LOGON_UNLOCK            7
        LOGON32_LOGON_NETWORK_CLEARTEXT 8
        LOGON32_LOGON_NEW_CREDENTIALS   9

        LOGON32_PROVIDER_DEFAULT    0
        LOGON32_PROVIDER_WINNT35    1
        LOGON32_PROVIDER_WINNT40    2
        LOGON32_PROVIDER_WINNT50    3

    }

    if {[min_os_version 6]} {
        twapi::add_defines {
            PROCESS_QUERY_LIMITED_INFORMATION      0x00001000
            PROCESS_ALL_ACCESS             0x001fffff
            THREAD_ALL_ACCESS              0x001fffff
        }
    } else {
        twapi::add_defines {
            PROCESS_ALL_ACCESS             0x001f0fff
            THREAD_ALL_ACCESS              0x001f03ff
        }
    }
}


# Returns a list of raw Windows API functions supported
proc twapi::list_raw_api {} {
    set rawapi [list ]
    foreach fn [info commands ::twapi::*] {
         if {[regexp {^::twapi::([A-Z][^_]*)$} $fn ignore fn]} {
             lappend rawapi $fn
         }
    }
    return $rawapi
}



# Get the handle for a Tcl channel
proc twapi::get_tcl_channel_handle {chan direction} {
    set direction [expr {[string equal $direction "write"] ? 1 : 0}]
    return [Tcl_GetChannelHandle $chan $direction]
}

# Wait for $wait_ms milliseconds or until $script returns $guard. $gap_ms is
# time between retries to call $script
# TBD - write a version that will allow other events to be processed
proc twapi::wait {script guard wait_ms {gap_ms 10}} {
    if {$gap_ms == 0} {
        set gap_ms 10
    }
    set end_ms [expr {[clock clicks -milliseconds] + $wait_ms}]
    while {[clock clicks -milliseconds] < $end_ms} {
        set script_result [uplevel $script]
        if {[string equal $script_result $guard]} {
            return 1
        }
        after $gap_ms
    }
    # Reached limit, one last try
    return [string equal [uplevel $script] $guard]
}

# Get twapi version
proc twapi::get_version {args} {
    array set opts [parseargs args {patchlevel}]
    if {$opts(patchlevel)} {
        return $twapi::patchlevel
    } else {
        return $twapi::version
    }
}



# Set all elements of the array to specified value
proc twapi::_array_set_all {v_arr val} {
    upvar $v_arr arr
    foreach e [array names arr] {
        set arr($e) $val
    }
}

# Check if any of the specified array elements are non-0
proc twapi::_array_non_zero_entry {v_arr indices} {
    upvar $v_arr arr
    foreach i $indices {
        if {$arr($i)} {
            return 1
        }
    }
    return 0
}

# Check if any of the specified array elements are non-0
# and return them as a list of options (preceded with -)
proc twapi::_array_non_zero_switches {v_arr indices all} {
    upvar $v_arr arr
    set result [list ]
    foreach i $indices {
        if {$all || ([info exists arr($i)] && $arr($i))} {
            lappend result -$i
        }
    }
    return $result
}


# Bitmask operations on 32bit values
# The int() casts are to deal with hex-decimal sign extension issues
proc twapi::setbits {v_bits mask} {
    upvar $v_bits bits
    set bits [expr {int($bits) | int($mask)}]
    return $bits
}
proc twapi::resetbits {v_bits mask} {
    upvar $v_bits bits
    set bits [expr {int($bits) & int(~ $mask)}]
    return $bits
}

# Return a bitmask corresponding to a list of symbolic and integer values
# If symvals is a single item, it is an array else a list of sym bitmask pairs
proc twapi::_parse_symbolic_bitmask {syms symvals} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
    } else {
        array set lookup $symvals
    }
    set bits 0
    foreach sym $syms {
        if {[info exists lookup($sym)]} {
            set bits [expr {$bits | $lookup($sym)}]
        } else {
            set bits [expr {$bits | $sym}]
        }
    }
    return $bits
}

# Return a list of symbols corresponding to a bitmask
proc twapi::_make_symbolic_bitmask {bits symvals {append_unknown 1}} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
        set map [array get lookup]
    } else {
        set map $symvals
    }
    set symbits 0
    set symmask [list ]
    foreach {sym val} $map {
        if {$bits & $val} {
            set symbits [expr {$symbits | $val}]
            lappend symmask $sym
        }
    }

    # Get rid of bits that mapped to symbols
    set bits [expr {$bits & ~$symbits}]
    # If any left over, add them
    if {$bits && $append_unknown} {
        lappend symmask $bits
    }
    return $symmask
}

# Return a bitmask corresponding to a list of symbolic and integer values
# If symvals is a single item, it is an array else a list of sym bitmask pairs
# Ditto for switches - an array or flat list of switch boolean pairs
proc twapi::_switches_to_bitmask {switches symvals {bits 0}} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
    } else {
        array set lookup $symvals
    }
    if {[llength $switches] == 1} {
        upvar $switches swtable
    } else {
        array set swtable $switches
    }

    foreach {switch bool} [array get swtable] {
        if {$bool} {
            set bits [expr {$bits | $lookup($switch)}]
        } else {
            set bits [expr {$bits & ~ $lookup($switch)}]
        }
    }
    return $bits
}

# Return a list of switche bool pairs corresponding to a bitmask
proc twapi::_bitmask_to_switches {bits symvals} {
    if {[llength $symvals] == 1} {
        upvar $symvals lookup
        set map [array get lookup]
    } else {
        set map $symvals
    }
    set symbits 0
    set symmask [list ]
    foreach {sym val} $map {
        if {$bits & $val} {
            set symbits [expr {$symbits | $val}]
            lappend symmask $sym 1
        } else {
            lappend symmask $sym 0
        }
    }

    return $symmask
}

# Make and return a keyed list
proc twapi::kl_create {args} {
    if {[llength $args] & 1} {
        error "No value specified for keyed list field [lindex $args end]. A keyed list must have an even number of elements."
    }
    return $args
}

# Make a keyed list given fields and values
interp alias {} twapi::kl_create2 {} twapi::twine

# Return a field from a keyed list or a default if not present
# This routine is now obsolete since the C version of kl_get takes
# an optional default parameter
# kl_get_default KEYEDLIST KEY DEFAULT
interp alias {} ::twapi::kl_get_default {} ::twapi::kl_get

# Set a key value
proc twapi::kl_set {kl field newval} {
   set i 0
   foreach {fld val} $kl {
        if {[string equal $fld $field]} {
            incr i
            return [lreplace $kl $i $i $newval]
        }
        incr i 2
    }
    lappend kl $field $newval
    return $kl
}

# Check if a field exists in the keyed list
proc twapi::kl_vget {kl field varname} {
    upvar $varname var
    return [expr {! [catch {set var [kl_get $kl $field]}]}]
}

# Remote/unset a key value
proc twapi::kl_unset {kl field} {
    array set arr $kl
    unset -nocomplain arr($field)
    return [array get arr]
}

# Compare two keyed lists
proc twapi::kl_equal {kl_a kl_b} {
    array set a $kl_a
    foreach {kb valb} $kl_b {
        if {[info exists a($kb)] && ($a($kb) == $valb)} {
            unset a($kb)
        } else {
            return 0
        }
    }
    if {[array size a]} {
        return 0
    } else {
        return 1
    }
}

# Return the field names in a keyed list in the same order that they
# occured
proc twapi::kl_fields {kl} {
    set fields [list ]
    foreach {fld val} $kl {
        lappend fields $fld
    }
    return $fields
}

# Returns a flat list of the $field fields from a list
# of keyed lists
proc twapi::kl_flatten {list_of_kl args} {
    set result {}
    foreach kl $list_of_kl {
        foreach field $args {
            lappend result [kl_get $kl $field]
        }
    }
    return $result
}


# Print the specified fields of a keyed list
proc twapi::kl_print {kl args} {
    # If only one arg, just print value without label
    if {[llength $args] == 1} {
        puts [kl_get $kl [lindex $args 0]]
        return
    }
    if {[llength $args] == 0} {
        array set arr $kl
        parray arr
    }
    foreach field $args {
        puts "$field: [kl_get $kl $field]"
    }
    return
}


# Return an array as a list of -index value pairs
proc twapi::_get_array_as_options {v_arr} {
    upvar $v_arr arr
    set result [list ]
    foreach {index value} [array get arr] {
        lappend result -$index $value
    }
    return $result
}



# Parse a list of two integers or a x,y pair and return a list of two integers
# Generate exception on format error using msg
proc twapi::_parse_integer_pair {pair {msg "Invalid integer pair"}} {
    if {[llength $pair] == 2} {
        lassign $pair first second
        if {[string is integer -strict $first] &&
            [string is integer -strict $second]} {
            return [list $first $second]
        }
    } elseif {[regexp {^([[:digit:]]+),([[:digit:]]+)$} $pair dummy first second]} {
        return [list $first $second]
    }

    error "$msg: '$pair'. Should be a list of two integers or in the form 'x,y'"
}


# Map console color name to integer attribute
proc twapi::_map_console_color {colors background} {
    set attr 0
    foreach color $colors {
        switch -exact -- $color {
            blue   {setbits attr 1}
            green  {setbits attr 2}
            red    {setbits attr 4}
            white  {setbits attr 7}
            bright {setbits attr 8}
            black  { }
            default {error "Unknown color name $color"}
        }
    }
    if {$background} {
        set attr [expr {$attr << 4}]
    }
    return $attr
}


# Convert file names by substituting \SystemRoot and \??\ sequences
proc twapi::_normalize_path {path} {
    # Get rid of \??\ prefixes
    regsub {^[\\/]\?\?[\\/](.*)} $path {\1} path

    # Replace leading \SystemRoot with real system root
    if {[string match -nocase {[\\/]Systemroot*} $path] &&
        ([string index $path 11] in [list "" / \\])} {
        return [file join [twapi::GetSystemWindowsDirectory] [string range $path 12 end]]
    } else {
        return [file normalize $path]
    }
}

# Convert a LARGE_INTEGER time value (100ns since 1601) to a formatted date
# time
interp alias {} twapi::large_system_time_to_secs {} twapi::large_system_time_to_secs_since_1970
proc twapi::large_system_time_to_secs_since_1970 {ns100 {fraction false}} {
    # No. 100ns units between 1601 to 1970 = 116444736000000000
    set ns100_since_1970 [expr {wide($ns100)-wide(116444736000000000)}]

    if {0} {
        set secs_since_1970 [expr {wide($ns100_since_1970)/wide(10000000)}]
        if {$fraction} {
            append secs_since_1970 .[expr {wide($ns100_since_1970)%wide(10000000)}]
        }
    } else {
        # Equivalent to above but faster
        if {[string length $ns100_since_1970] > 7} {
            set secs_since_1970 [string range $ns100_since_1970 0 end-7]
            if {$fraction} {
                set frac [string range $ns100_since_1970 end-6 end]
                append secs_since_1970 .$frac
            }
        } else {
            set secs_since_1970 0
            if {$fraction} {
                set frac [string range "0000000${ns100_since_1970}" end-6 end]
                append secs_since_1970 .$frac
            }
        }
    }
    return $secs_since_1970
}

proc twapi::secs_since_1970_to_large_system_time {secs} {
    # No. 100ns units between 1601 to 1970 = 116444736000000000
    return [expr {($secs * 10000000) + wide(116444736000000000)}]
}

interp alias {} ::twapi::get_system_time {} ::twapi::GetSystemTimeAsFileTime
interp alias {} ::twapi::large_system_time_to_timelist {} ::twapi::FileTimeToSystemTime
interp alias {} ::twapi::timelist_to_large_system_time {} ::twapi::SystemTimeToFileTime

# Convert seconds to a list {Year Month Day Hour Min Sec Ms}
# (Ms will always be zero). Always return local time
proc twapi::_seconds_to_timelist {secs} {
    # For each field, we need to trim the leading zeroes
    set result [list ]
    foreach x [clock format $secs -format "%Y %m %e %k %M %S 0" -gmt false] {
        lappend result [scan $x %d]
    }
    return $result
}

# Convert local time list {Year Month Day Hour Min Sec Ms} to seconds
# (Ms field is ignored)
proc twapi::_timelist_to_seconds {timelist} {
    return [clock scan [_timelist_to_timestring $timelist] -gmt false]
}

# Convert local time list {Year Month Day Hour Min Sec Ms} to a time string
# (Ms field is ignored)
proc twapi::_timelist_to_timestring {timelist} {
    if {[llength $timelist] < 6} {
        error "Invalid time list format"
    }

    return "[lindex $timelist 0]-[lindex $timelist 1]-[lindex $timelist 2] [lindex $timelist 3]:[lindex $timelist 4]:[lindex $timelist 5]"
}

# Convert a time string to a time list
proc twapi::_timestring_to_timelist {timestring} {
    return [_seconds_to_timelist [clock scan $timestring -gmt false]]
}

# Parse raw memory like binary scan command
proc twapi::mem_binary_scan {mem off mem_sz args} {
    uplevel [list binary scan [Twapi_ReadMemoryBinary $mem $off $mem_sz]] $args
}


# Validate guid syntax
proc twapi::_validate_guid {guid} {
    if {![Twapi_IsValidGUID $guid]} {
        error "Invalid GUID syntax: '$guid'"
    }
}

# Validate uuid syntax
proc twapi::_validate_uuid {uuid} {
    if {![regexp {^[[:xdigit:]]{8}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{12}$} $uuid]} {
        error "Invalid UUID syntax: '$uuid'"
    }
}

# Extract a UCS-16 string from a binary. Cannot directly use
# encoding convertfrom because that will not stop at the terminating
# null. The UCS-16 assumed to be little endian.
proc twapi::_ucs16_binary_to_string {bin {off 0}} {
    return [encoding convertfrom unicode [string range $bin $off [string first \0\0\0 $bin]]]
}

# Given a binary, return a GUID. The formatting is done as per the
# Windows StringFromGUID2 convention used by COM
proc twapi::_binary_to_guid {bin {off 0}} {
    if {[binary scan $bin "@$off i s s H4 H12" g1 g2 g3 g4 g5] != 5} {
        error "Invalid GUID binary"
    }

    return [format "{%8.8X-%2.2hX-%2.2hX-%s}" $g1 $g2 $g3 [string toupper "$g4-$g5"]]
}

# Given a guid string, return a GUID in binary form
proc twapi::_guid_to_binary {guid} {
    _validate_guid $guid
    lassign [split [string range $guid 1 end-1] -] g1 g2 g3 g4 g5
    return [binary format "i s s H4 H12" 0x$g1 0x$g2 0x$g3 $g4 $g5]
}

# Return a guid from raw memory
proc twapi::_decode_mem_guid {mem {off 0}} {
    return [_binary_to_guid [Twapi_ReadMemoryBinary $mem $off 16]]
}

# Convert a Windows registry value to Tcl form. mem is a raw
# memory object. off is the offset into the memory object to read.
# $type is a integer corresponding
# to the registry types
proc twapi::_decode_mem_registry_value {type mem len {off 0}} {
    set type [expr {$type}];    # Convert hex etc. to decimal form
    switch -exact -- $type {
        1 -
        2 {
            # Note - pass in -1, not $len since we do not
            # want terminating nulls
            return [list [expr {$type == 2 ? "expand_sz" : "sz"}] \
                        [Twapi_ReadMemoryUnicode $mem $off -1]]
        }
        7 {
            # Collect strings until we come across an empty string
            # Note two nulls right at the start will result in
            # an empty list. Should it result in a list with
            # one empty string element? Most code on the web treats
            # it as the former so we do too.
           set multi [list ]
            while {1} {
                set str [Twapi_ReadMemoryUnicode $mem $off -1]
                set n [string length $str]
                # Check for out of bounds. Cannot check for this before
                # actually reading the string since we do not know size
                # of the string.
                if {($len != -1) && ($off+$n+1) > $len} {
                    error "Possible memory corruption: read memory beyond specified memory size."
                }
                if {$n == 0} {
                    return [list multi_sz $multi]
                }
                lappend multi $str
                # Move offset by length of the string and terminating null
                # (times 2 since unicode and we want byte offset)
                incr off [expr {2*($n+1)}]
            }
        }
        4 {
            if {$len < 4} {
                error "Insufficient number of bytes to convert to integer."
            }
            return [list dword [Twapi_ReadMemoryInt $mem $off]]
        }
        5 {
            if {$len < 4} {
                error "Insufficient number of bytes to convert to big-endian integer."
            }
            set type "dword_big_endian"
            set scanfmt "I"
            set len 4
        }
        11 {
            if {$len < 8} {
                error "Insufficient number of bytes to convert to wide integer."
            }
            set type "qword"
            set scanfmt "w"
            set len 8
        }
        0 { set type "none" }
        6 { set type "link" }
        8 { set type "resource_list" }
        3 { set type "binary" }
        default {
            error "Unsupported registry value type '$type'"
        }
    }

    set val [Twapi_ReadMemoryBinary $mem $off $len]
    if {[info exists scanfmt]} {
        if {[binary scan $val $scanfmt val] != 1} {
            error "Could not convert from binary value using scan format $scanfmt"
        }
    }

    return [list $type $val]
}

proc twapi::Twapi_PtrToAddress {p} {
    if {[Twapi_IsPtr $p]} {
        set addr [lindex $p 0]
        if {$addr eq "NULL"} {
            return 0
        } else {
            return $addr
        }
    } else {
        error "'$p' is not a valid pointer value."
    }
}

proc twapi::Twapi_PtrType {p} {
    if {[Twapi_IsPtr $p]} {
        set type [lindex $p 1]
        if {$type eq ""} {
            set type void*
        }
    } else {
        error "'$p' is not a valid pointer value."
    }
    return $type
}

proc twapi::Twapi_AddressToPtr {addr type} {
    return [list $addr $type]
}


proc twapi::_log_timestamp {} {
    return [clock format [clock seconds] -format "%a %T"]
}

# If we have a .tm extension, we are a 8.5 Tcl module or embedded script,
# we expect all source files to have been appended to this file. So do not
# source them.
if {([file extension [info script]] ne ".tm") && [twapi::get_build_config embed_type] eq "none"} {

    # When running from the source dir (while developing), there is
    # no buildid file but we do not really care
    if {[file exists [file join [file dirname [info script]] twapi_buildid.tcl]]} {
        source [file join [file dirname [info script]] twapi_buildid.tcl]
    }

    # Source files based on build configuration
    # First, all the base files
    foreach ::twapi::_field_ {
        handle.tcl
        metoo.tcl
        osinfo.tcl
        security.tcl
        process.tcl
        disk.tcl
    } {
        source [file join [file dirname [info script]] $::twapi::_field_]
    }

    # Now the desktop files
    if {[lsearch [::twapi::get_build_config opts] nodesktop] < 0} {
        foreach ::twapi::_field_ {
            ui.tcl
            clipboard.tcl
            shell.tcl
            nls.tcl
            com.tcl
        } {
            source [file join [file dirname [info script]] $::twapi::_field_]
        }
    }


    # Now the server files
    if {[lsearch [::twapi::get_build_config opts] noserver] < 0} {
        foreach ::twapi::_field_ {
            services.tcl
            eventlog.tcl
        } {
            source [file join [file dirname [info script]] $::twapi::_field_]
        }
    }


    # Now the extras
    if {[lsearch [::twapi::get_build_config opts] lean] < 0} {
        foreach ::twapi::_field_ {
            adsi.tcl
            process2.tcl
            accounts.tcl
            pdh.tcl
            share.tcl
            network.tcl
            console.tcl
            synch.tcl
            desktop.tcl
            printer.tcl
            mstask.tcl
            msi.tcl
            crypto.tcl
            device.tcl
            power.tcl
            namedpipe.tcl
            resource.tcl
            rds.tcl
            wmi.tcl
            etw.tcl
        } {
            source [file join [file dirname [info script]] $::twapi::_field_]
        }
    }


    # Get rid of temp variable
    unset twapi::_field_
}

# Returns a list of twapi procs that are currently defined and should
# be exported. SHould be called after completely loading twapi
proc twapi::_get_public_procs {} {

    set public_procs [info procs]

    # Init with C built-ins - there does not seem an easy auto way
    # of getting these. Also, ensembles although probably there is
    # a way of doing this.
    lappend public_procs {*}{
        canonicalize_guid
        is_valid_sid_syntax
        kl_get
        parseargs
        recordarray
        systemtray
        trap
        twine
    }

    # Also export aliases but not "try" as it conflicts
    # with 8.6 try
    foreach p [interp aliases] {
        if {[regexp {twapi::([a-z][^:]*)$} $p _ tail]} {
            if {$tail ne "try"} {
                lappend public_procs $tail
            }
        }
    }

    return $public_procs
}

# Used in various matcher callbacks to signify always include etc.
# TBD - document
proc twapi::true {args} {
    return true
}

namespace eval twapi {
    # Get a handle to ourselves. This handle never need be closed
    variable my_process_handle [GetCurrentProcess]

    # TBD - To improve start-up times, we should really enumerate exports
    # at build time,

    # eval namespace export [::twapi::_get_public_procs]
}

proc twapi::export_public_commands {} {
    uplevel #0 [list namespace eval twapi [list eval namespace export [::twapi::_get_public_procs]]]
}

proc twapi::import_commands {} {
    export_public_commands
    uplevel namespace import twapi::*
}

if {![info exists ::twapi::package_name]} {
    set ::twapi::package_name "twapi"
}

package provide $::twapi::package_name $twapi::patchlevel

# Disabled auto-import of commands as it can cause confusion
if {0 && [llength [info commands tkcon*]]} {
    twapi::import_commands
}
