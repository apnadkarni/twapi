Set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
Set processObjs = wmiObj.ExecQuery ("SELECT * FROM Win32_Process")
For each processObj in processObjs
    properties = processObj.GetOwner(user, domain)       
    Wscript.Echo _
        "ProcessId*" & processObj.ProcessId & _
        "*Domain*" & domain & _
        "*ExecutablePath*" & processObj.ExecutablePath & _
        "*ExecutionState*" & processObj.ExecutionState & _
        "*HandleCount*" & processObj.HandleCount & _
        "*Name*" & processObj.Name & _
        "*OtherOperationCount*" & processObj.OtherOperationCount & _
        "*OtherTransferCount*" & processObj.OtherTransferCount & _
        "*PageFileUsage*" & processObj.PageFileUsage & _
        "*PeakPageFileUsage*" & processObj.PeakPageFileUsage & _
        "*ParentProcessId*" & processObj.ParentProcessId & _
        "*PeakVirtualSize*" & processObj.PeakVirtualSize & _
        "*Priority*" & processObj.Priority & _
        "*PrivatePageCount*" & processObj.PrivatePageCount & _
        "*QuotaNonPagedPoolUsage*" & processObj.QuotaNonPagedPoolUsage & _
        "*QuotaPagedPoolUsage*" & processObj.QuotaPagedPoolUsage & _
        "*ReadOperationCount*" & processObj.ReadOperationCount & _
        "*ReadTransferCount*" & processObj.ReadTransferCount & _
        "*ThreadCount*" & processObj.ThreadCount & _
        "*User*" & user & _
        "*VirtualSize*" & processObj.VirtualSize & _
        "*WorkingSetSize*" & processObj.WorkingSetSize & _
        "*WriteOperationCount*" & processObj.WriteOperationCount & _
        "*WriteTransferCount*" & processObj.WriteTransferCount & _
        "*PeakWorkingSetSize*" & processObj.PeakWorkingSetSize
Next
