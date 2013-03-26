# File to setup the vfs for a kitgen build where twapi is statically
# linked into the tclkit executable. This script should be placed in the
# same directory as the twapi static libraries.

# We just copy our package index file into the kitgen vfs area
# vfs variable is set up by the setupvfs.tcl kitgen build script
file mkdir [file join $vfs lib twapi]
file copy [file join [file dirname [info script]] pkgIndex.tcl] [file join $vfs lib twapi]
