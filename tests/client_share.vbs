Set netObj = CreateObject("Wscript.Network")
Set driveObjs = netObj.EnumNetworkDrives
For i = 0 to driveObjs.Count - 1 Step 2
    Wscript.Echo "drive" & driveObjs.Item(i) & _
        "*" & driveObjs.Item(i+1)
Next

Set prnObjs = netObj.EnumPrinterConnections
For i = 0 to prnObjs.Count - 1 Step 2
    Wscript.Echo "printer" & prnObjs.Item(i) & _
        "*" & prnObjs.Item(i+1)
Next
