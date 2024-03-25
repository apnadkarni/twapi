set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
set osList = wmiObj.ExecQuery ("SELECT * FROM Win32_OperatingSystem")
for each osObj in osList
    Wscript.Echo "Name*" & osObj.Name _
    & "*Version*" & osObj.Version _
    & "*ServicePackMajorVersion*" & osObj.ServicePackMajorVersion _
    & "*ServicePackMinorVersion*" & osObj.ServicePackMinorVersion _
    & "*BuildNumber*" & osObj.BuildNumber _
    & "*WindowsDirectory*" & osObj.WindowsDirectory _
    & "*Locale*" & osObj.Locale _
    & "*FreePhysicalMemory*" & osObj.FreePhysicalMemory _
    & "*TotalVirtualMemorySize*" & osObj.TotalVirtualMemorySize _
    & "*FreeVirtualMemory*" & osObj.FreeVirtualMemory _
    & "*SizeStoredInPagingFiles*" & osObj.SizeStoredInPagingFiles
Next

set osList = wmiObj.ExecQuery _
    ("SELECT * FROM Win32_ComputerSystem")
for Each objComputer in osList 
    Wscript.Echo "System Name: " & objComputer.Name
    Wscript.Echo "System Manufacturer: " & objComputer.Manufacturer
    Wscript.Echo "System Model: " & objComputer.Model
    Wscript.Echo "Time Zone: " & objComputer.CurrentTimeZone
    Wscript.Echo "Total Physical Memory: " & _
        objComputer.TotalPhysicalMemory
Next
