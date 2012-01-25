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
