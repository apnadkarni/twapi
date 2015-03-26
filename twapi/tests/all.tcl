# Important constraints to be controlled from command line:
# systemmodificationok - will run tests that modify system configuration
#  such as file shares. They are supposed to clean up after themselves
#  but that may itself fail after test failures
# userInteraction - tcltest builtin - expect user to respond in some way
#  for the test case to run

package require Tcl
package require tcltest 2.2
source [file join [file dirname [info script]] testutil.tcl]

# TBD - should we load twapi here ? Probably not
load_twapi_package

# Test configuration options that may be set are:
#  systemmodificationok - will include tests that modify the system
#      configuration (eg. add users, share a disk etc.)

puts "Test environment: Tcl [info patchlevel], [expr {[info exists ::env(TWAPI_PACKAGE)] ? $::env(TWAPI_PACKAGE) : "twapi" }] [twapi::get_version -patchlevel]"

tcltest::configure -testdir [file dirname [file normalize [info script]]]
tcltest::configure -tmpdir $::env(TEMP)/twapi-test/[clock seconds] {*}$argv
tcltest::runAllTests
puts "All done."
