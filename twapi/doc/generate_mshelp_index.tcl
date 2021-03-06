# Create a linkable name that matches the one created by the doctools
# htm formatter.
# This proc is duplicated all over the place, for example in
# fmt.htf in twapi version of doctools...oh well
proc make_linkable_cmd {name} {
    # Map starting - to x- as the former is not legal in XHTML "id" attribute
    if {[string index $name 0] eq "-"} {
        set name x$name
    }

    # If name is all upper case, attach another underscore
    if {[regexp {[[:alnum:]]+} $name] &&
        [string toupper $name] eq $name} {
        append name _uc
    }

    # Massage some other funky chars
    set name [string map {{ } {} ? __} $name]    

    # Hopefully all the above will not lead to name clashes
    return [string tolower $name]
}

# Define the commands to implement the commands in the file
array set keywords {}
proc index_begin args {}
proc index_end args {}
proc key keyword {
    set ::current_keyword [eval concat $keyword]
    return
}
proc manpage {filename description} {
    set filename [file rootname $filename].html
    set entry {<LI> <OBJECT type="text/sitemap">}
    append entry "\n\t<param name=\"Name\" value=\"$::current_keyword\">"
    if {$description != ""} {
        append entry "\n\t<param name=\"Name\" value=\"$description\">"
    }
    # We stick in #current_keyword as a locator. If it does not exist,
    # no problem, we will just get the whole page instead. To match
    # the manpage htm format generator, we replace - with _ in keyword
    # since the former is not valid in XHTML id fields.
    if {1 || [string first " " $::current_keyword] < 0} {
        append entry "\n\t<param name=\"Local\" value=\"$filename#[make_linkable_cmd $::current_keyword]\">"
    } else {
        append entry "\n\t<param name=\"Local\" value=\"$filename\">"
    }
    append entry "\n\t</OBJECT>\n"
    return $entry
}

#
# Processing being here
#

# Output the standard header
puts {
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<HTML>
<HEAD>
<meta name="GENERATOR" content="Microsoft&reg; HTML Help Workshop 4.1">
<!-- Sitemap 1.0 -->
</HEAD><BODY>
<UL>
}

# Open the first argument as a index file generated by dtp
set fd [open [lindex $argv 0]]
set data [read $fd]
close $fd

# Eval the file - should really be in a safe interpreter! TBD
puts [subst $data]

# Write the trailer
puts {
</UL>
</BODY></HTML>
}
