#
# Copyright (c) 2012 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

twapi::class create ::twapi::IMofCompilerProxy {
    superclass ::twapi::IUnknownProxy

    constructor {args} {
        if {[llength $args] == 0} {
            set args [list [::twapi::com_create_instance "{6daf9757-2e37-11d2-aec9-00c04fb68820}" -interface IMofCompiler -raw]]
        }
        next {*}$args
    }

    method CompileBuffer args {
        my variable _ifc
        return [::twapi::IMofCompiler_CompileBuffer $_ifc {*}$args]
    }

    method CompileFile args {
        my variable _ifc
        return [::twapi::IMofCompiler_CompileFile $_ifc {*}$args]
    }

    method CreateBMOF args {
        my variable _ifc
        return [::twapi::IMofCompiler_CreateBMOF $_ifc {*}$args]
    }

    twapi_exportall
}


#
# Get WMI service
proc twapi::_wmi {{top cimv2}} {
    return [comobj_object "winmgmts:{impersonationLevel=impersonate}!//./root/$top"]
}

proc twapi::wmi_find_classes {swbemservices args} {
    array set opts [parseargs args {
        {parent.arg {}}
        shallow
        first
        matchproperties.arg
        matchsystemproperties.arg
        matchqualifiers.arg
    } -maxleftover 0]
    
    
    # Create a forward only enumerator for efficiency
    # wbemFlagUseAmendedQualifiers | wbemFlagReturnImmediately | wbemFlagForwardOnly
    set flags 0x20030
    if {$opts(shallow)} {
        incr flags 1;           # 0x1 -> wbemQueryFlagsShallow
    }

    set classes [$swbemservices SubclassesOf $opts(parent) $flags]
    set matches {}
    twapi::trap {
        $classes -iterate class {
            set matched 1
            foreach {opt fn} {
                matchproperties Properties_
                matchsystemproperties SystemProperties_
                matchqualifiers Qualifiers_
            } {
                if {[info exists opts($opt)]} {
                    foreach {name matcher} $opts($opt) {
                        if {[catch {
                            if {! [{*}$matcher [$class -with [list [list -get $fn] [list Item $name]] Value]]} {
                                set matched 0
                                break; # Value does not match
                            }
                        } msg ]} {
                            # No such property or no access
                            set matched 0
                            break
                        }
                    }
                }
                if {! $matched} {
                    # Already failed to match, no point continuing looping
                    break
                }
            }

            if {$matched} {
                if {$opts(first)} {
                    return $class
                } else {
                    lappend matches $class
                }
            }
        }
    } onerror {} {
        foreach class $matches {
            $class destroy
        }
        error $::errorResult $::errorInfo $::errorCode
    } finally {
        $classes destroy
    }

    return $matches

}

