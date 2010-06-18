#
# Definitions for Win32 API calls through a standard handler
# Note most are defined at C level in twapi_calls.c and this file
# only contains those that could not be defined there for whatever
# reason.

# NOTE CALLS SHOULD BE DEFINED IN twapi_calls.c AS MUCH AS POSSIBLE
# SINCE WE WANT twapi.dll TO BE USABLE BY ITSELF WITHOUT ANY TCL SCRIPT
# WRAPPERS.

namespace eval twapi {}

proc twapi::IUnknown_QueryInterface {ifc iid} {
    set iidname void
    catch {set iidname [registry get HKEY_CLASSES_ROOT\\Interface\\$iid ""]}
    return [Twapi_IUnknown_QueryInterface $ifc $iid $iidname]
}

proc twapi::CoGetObject {name bindopts iid} {
    set iidname void
    catch {set iidname [registry get HKEY_CLASSES_ROOT\\Interface\\$iid ""]}
    return [Twapi_CoGetObject $name $bindopts $iid $iidname]
}

proc twapi::CoCreateInstance {clsid iunknown context iid} {
    set iidname void
    catch {set iidname [registry get HKEY_CLASSES_ROOT\\Interface\\$iid ""]}
    return [Twapi_CoCreateInstance $clsid $iunknown $context $iid $iidname]
}
