# Tcl Windows API (TWAPI) extension

The Tcl Windows API (TWAPI) extension provides access to the Windows API
from within the Tcl scripting language.

  * Project source repository is at https://github.com/apnadkarni/twapi
  * Documentation is at https://twapi.magicsplat.com
  * Binary distribution is at https://sourceforge.net/projects/twapi/files/Current%20Releases/Tcl%20Windows%20API/

## Supported platforms

TWAPI 5.0 requires

  * Windows 7 SP1 or later
  * Tcl 8.6.10+ or Tcl 9.x

### Binary distribution

The single binary distribution supports Tcl 8.6 and Tcl 9 for both 32-
and 64-bit platforms.

It requires the VC++ runtime to already be installed on the system.
Download from
https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist
if necessary.

Windows 7 and 8.x also require the Windows UCRT runtime to be installed
if not present. Download from
https://support.microsoft.com/en-gb/topic/update-for-universal-c-runtime-in-windows-c0514201-7fe6-95a3-b0a5-287930f3560c.

In most cases, both the above should already be present on the system.

Note that the *modular* and single file *bin* in 4.x distributions are
no longer available and will not be supported in 5.0.

## TWAPI Summary

The Tcl Windows API (TWAPI) extension provides access to the Windows API
from within the Tcl scripting language.

Functions in the following areas are implemented:

  * System functions including OS and CPU information,
    shutdown and message formatting
  * User and group management
  * COM client and server support
  * Security and resource access control
  * Window management
  * User input: generate key/mouse input and hotkeys
  * Basic sound playback functions
  * Windows services
  * Windows event log access
  * Windows event tracing
  * Process and thread management
  * Directory change monitoring
  * Lan Manager and file and print shares
  * Drive information, file system types etc.
  * Network configuration and statistics
  * Network connection monitoring and control
  * Named pipes
  * Clipboard access
  * Taskbar icons and notifications
  * Console mode functions
  * Window stations and desktops
  * Internationalization
  * Task scheduling
  * Shell functions
  * Registry
  * Windows Management Instrumentation
  * Windows Installer
  * Synchronization
  * Power management
  * Device I/O and management
  * Crypto API and certificates
  * SSL/TLS
  * Windows Performance Counters
