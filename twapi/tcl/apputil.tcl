#
# Copyright (c) 2003-2012, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

namespace eval twapi {}

# Get the command line
proc twapi::get_command_line {} {
    return [GetCommandLineW]
}

# Parse the command line
proc twapi::get_command_line_args {cmdline} {
    # Special check for empty line. CommandLinetoArgv returns process
    # exe name in this case.
    if {[string length $cmdline] == 0} {
        return [list ]
    }
    return [CommandLineToArgv $cmdline]
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

proc twapi::_env_block_to_dict {block normalize} {
    set env_dict {}
    foreach env_str $block {
        set pos [string first = $env_str]
        set key [string range $env_str 0 $pos-1]
        if {$normalize} {
            set key [string toupper $key]
        }
        lappend env_dict $key [string range $env_str $pos+1 end]
    }
    return $env_dict
}

proc twapi::get_system_environment_vars {args} {
    parseargs args {normalize.bool} -nulldefault -setvars -maxleftover 0
    return [_env_block_to_dict [CreateEnvironmentBlock 0 0] $normalize]
}

proc twapi::get_user_environment_vars {token args} {
    parseargs args {inherit.bool normalize.bool} -nulldefault -setvars -maxleftover 0
    return [_env_block_to_dict [CreateEnvironmentBlock $token $inherit] $normalize]
}

proc twapi::expand_system_environment_vars {s} {
    return [ExpandEnvironmentStringsForUser 0 $s]
}

proc twapi::expand_user_environment_vars {tok s} {
    return [ExpandEnvironmentStringsForUser $tok $s]
}


