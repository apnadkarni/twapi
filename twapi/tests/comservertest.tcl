console show
source [file join [file dirname [info script]] testutil.tcl]
load_twapi

twapi::class create Adder {
    method Sum args {
        return [tcl::mathop::+ {*}$args]
    }
    export Sum
}

set adder_clsid {{332B8252-2249-4B34-BAD3-81259F2A2842}}
twapi::comserver_factory $adder_clsid {0 Sum} {Adder new} factory
factory register
run_comservers
factory destroy

