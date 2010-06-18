Set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
Set diskObjs = wmiObj.ExecQuery ("SELECT * FROM Win32_LogicalDisk")
For each diskObj in diskObjs
    Wscript.Echo diskObj.DeviceID & _
        "*" & diskObj.DriveType	& _
        "*" & diskObj.FileSystem	& _
        "*" & diskObj.FreeSpace	& _
        "*" & diskObj.MediaType	& _
        "*" & diskObj.Name	& _
        "*" & diskObj.Size	& _
        "*" & diskObj.SupportsFileBasedCompression	& _
        "*" & diskObj.VolumeName	& _
        "*" & diskObj.VolumeSerialNumber
Next
