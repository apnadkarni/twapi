#
# Generates a WiTS Web page.
#   tclsh generate_web_page.tcl INPUTFILE VERSION [NAVFILE] [ADFILE] [ADFILE2] [OUTPUTFILE]
# where INPUTFILE is the html fragment that should go into the content
# section of the web page, NAVFILE is the navigation link fragment file,
# and the two ADFILE are right side and top ad content.
# TBD - fix hardcoded version numbers and copyright years

# TBD - for IE8, we may need to insert the following as it defaults to full
# standards mode
# <meta http-equiv="X-UA-Compatible" content="IE=7" />

#
# Read the given file and write out the HTML
proc transform_file {infile {navfile ""} {adfile ""} {adfile2 ""} {outfile ""}} {
    set infd [open $infile r]
    set frag [read $infd]
    close $infd
    if {$outfile eq ""} {
        set outfd stdout
    } else {
        set outfd [open $outfile w]
    }

    puts $outfd {<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN">}
    puts $outfd {
<html>
  <head>
    <title>Tcl Windows API extension</title>
    <link rel="shortcut icon" href="favicon.ico" />
    <link rel="stylesheet" type="text/css" href="http://yui.yahooapis.com/2.5.1/build/reset-fonts-grids/reset-fonts-grids.css"/>
    <link rel="stylesheet" type="text/css" href="styles.css" />
  </head>
  <body>
    <div id="doc" class="yui-t4">
    }


    # Put the page header
    puts $outfd "<div id='hd'>"

    # TBD - remove hardcoding once we know we want it
    set fd [open google-searchbox.js r]
    set searchbox [read $fd]
    close $fd

    puts "<div class='searchbox'>$searchbox</div>"

    puts $outfd {
        <div class='headingbar'>
        <a href='http://www.magicsplat.com'><img style='float:right; padding-right: 5px;' src='magicsplat.png' alt='site logo'/></a>
        <p><a href='index.html'>Tcl Windows API extension</a></p>
        </div>
    }

    # Insert the horizontal ad
    if {$adfile2 ne ""} {
        set adfd [open $adfile2 r]
        set addata [read $adfd]
        close $adfd
        puts $outfd "<div class='headingads'>$addata</div>"
    }

    # Terminate "hd"
    puts $outfd {
      </div>
    }

    # Put the main area headers
    puts $outfd {
      <div id="bd">
        <div id="yui-main">
          <div class="yui-b">
            <div class="yui-gf">
    }

    # Put the actual text
    puts $outfd {
        <div class="yui-u content">
    }

    puts -nonewline $outfd  $frag
    puts "</div>"


    # Put the navigation pane
    if {$navfile ne ""} {
        set navfd [open $navfile r]
        set navdata [read $navfd]
        close $navfd
        puts $outfd "<div class='yui-u first navigation'>"
        puts $outfd "<a class='imgbutton' href='http://sourceforge.net/project/showfiles.php?group_id=90123'><img title='Download button' alt='Download' src='download.png' onmouseover='javascript:this.src=\"download_active.png\"' onmouseout='javascript:this.src=\"download.png\"' /></a>"
        puts $outfd "<hr style='width: 100px; margin-left: 0pt;'/>"
        puts $outfd "<h2>TWAPI $::twapi_version Documentation</h2>"
        puts $outfd "<ul>\n$navdata\n</ul></div>"
    }

    # Terminate the yui-main, yui-b and yui-gf above
    puts $outfd {
        </div>
        </div>
        </div>
    }

    # Insert the ad pane
    if {$adfile ne ""} {
        set adfd [open $adfile r]
        set addata [read $adfd]
        close $adfd
        puts $outfd "<div class='yui-b'>"
        puts -nonewline $outfd "<div class='sideads'>$addata</div>"
        puts $outfd "</div>"
    }

    # Terminate the main body bd
    puts $outfd "</div>"

    # Insert the footer
    puts $outfd "<div id='ft'>"
    puts $outfd "Tcl Windows API $::twapi_version"
    puts $outfd {
        <div class='copyright'>
          &copy; 2002-2011 Ashok P. Nadkarni
        </div>
        <a href='http://www.magicsplat.com/privacy.html'>Privacy policy</a>
    }
    puts $outfd "</div>"


    # Finally terminate the whole div and body and html
    puts $outfd {
        </div>
        </body>
        </html>
    }

    flush $outfd
    if {$outfile ne ""} {
        close $outfd
    }
}

set twapi_version [lindex $argv 1]
transform_file [lindex $argv 0] [lindex $argv 2] [lindex $argv 3] [lindex $argv 4]
