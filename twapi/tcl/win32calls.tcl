#
# Definitions for Win32 API calls through a standard handler
# Note most are defined at C level in twapi_calls.c and this file
# only contains those that could not be defined there for whatever
# reason.

# NOTE CALLS SHOULD BE DEFINED IN twapi_calls.c AS MUCH AS POSSIBLE
# SINCE WE WANT twapi.dll TO BE USABLE BY ITSELF WITHOUT ANY TCL SCRIPT
# WRAPPERS.

namespace eval twapi {}

