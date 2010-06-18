# Output the Windows help project file
puts {
[OPTIONS]
Compatibility=1.1 or later
Compiled file=twapi.chm
Contents file=twapi.hhc
Default Window=Tcl Windows API Help
Default topic=overview.html
Display compile progress=No
Full-text search=Yes
Index file=twapi.hhk
Language=0x409 English (United States)
Title=Tcl Windows API Help

[WINDOWS]
Tcl Windows API Help="Tcl Windows API Help","twapi.hhc","twapi.hhk","overview.html",,,,,,0x42520,,0x380e,,,,,,,,0

[FILES]
}

puts [join $argv \n]

puts {
[INFOTYPES]
}