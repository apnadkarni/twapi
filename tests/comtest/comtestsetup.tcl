exec c:/windows/syswow64/regsvr32 comtest.dll
set guid {{310FEA61-BC62-4944-84BE-D9DB986701DC}}
package require registry
registry -64bit set HKEY_CLASSES_ROOT\\Wow6432Node\\CLSID\\$guid AppID $guid
registry -64bit set HKEY_CLASSES_ROOT\\Wow6432Node\\AppID\\$guid
registry -64bit set HKEY_CLASSES_ROOT\\Wow6432Node\\AppID\\$guid DllSurrogate "" sz
registry -64bit set HKEY_LOCAL_MACHINE\\Software\\Wow6432Node\\Classes\\AppID\\$guid DllSurrogate "" sz
