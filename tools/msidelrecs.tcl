#
# Copyright (c) 2004-2006 Ashok P. Nadkarni
# All rights reserved.
#
# See the file LICENSE for license

#
# Script to delete records from MSI files
#
#  tclsh msidelrecs.tcl MSIFILE TABLE COLUMN VALUE1 VALUE2...

package require twapi 2.0

# Delete records matching a value in a field
proc delete_record {db table column value} {
    # Open a view with rows for the value
    # Yes, I know there is sql injection vulnerability here
    # but this is just a script for building and I can't be bothered
    # to figure out how to use parameter token with MSI
    set view [$db -call OpenView "select $column from $table where $column='$value'"]
    twapi::load_msi_prototypes $view view
    twapi::try {
        $view -call Execute
        # Now we loop through and delete all the records we find
        while {true} {
            set rec [$view Fetch]
            if {[$rec -isnull]} {
                break;                  # All done, no more records
            }
            twapi::load_msi_prototypes $rec record
            # Delete the record. 6 == delete record
            $view -call Modify 6 [$rec -interface]
            $rec -destroy
        }
    } finally {
        $view -call Close
        $view -destroy
    }
}

# Script starts here

# Parse arguments
if {[llength $argv] < 3} {
    puts stderr "Usage: [info nameofexecutable] $argv0 MSIFILE TABLE COLUMN VALUE1 VALUR2..."
    exit 1
}
foreach {msifile msitable msicolumn} $argv break
set msivalues [lrange $argv 3 end]
set msiinstaller [twapi::new_msi]
twapi::try {
    twapi::load_msi_prototypes $msiinstaller installer
    set msidb [$msiinstaller -call OpenDatabase $msifile 1]
    twapi::load_msi_prototypes $msidb database
    
    # Create a view for each value and delete one at a time
    # I'm sure there are more efficient ways but I don't know
    # exactly what level of SQL the MSI engine understands
    foreach val $msivalues {
        delete_record $msidb $msitable $msicolumn $val
    }

    # No errors, commit the db
    $msidb -call Commit

} finally {
    twapi::delete_msi $msiinstaller
    if {[info exists msidb]} {
        $msidb -destroy
    }
}