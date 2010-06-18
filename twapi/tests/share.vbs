Set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
Set shareObjs = wmiObj.ExecQuery ("SELECT * FROM Win32_Share")
For Each shareObj in shareObjs
    Wscript.Echo "AllowMaximum*" & shareObj.AllowMaximum  _
    & "*Caption*" & shareObj.Caption   _
    & "*MaximumAllowed*" & shareObj.MaximumAllowed _
    & "*Name*" & shareObj.Name   _
    & "*Path*" & shareObj.Path   _
    & "*Type*" & shareObj.Type
Next
