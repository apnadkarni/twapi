#
# Copyright (c) 2003, Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

package require doctools

set doc_files {
osinfo 
process
}

proc generate_html {} {
    foreach docfile $::doc_files {
        set fd [open "${docfile}.man" r]
        set indata [read $fd]
        close $fd

        set doctool [doctools::new doctool -format htm -file $docfile]
        set outdata [$doctool format $indata]
        $doctool destroy
        
        set fd [open "${docfile}.htm" w]
        puts $fd $outdata
        close $fd
    }
}
