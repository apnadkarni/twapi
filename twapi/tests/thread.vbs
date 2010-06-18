Set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
Set threadObjs = wmiObj.ExecQuery ("SELECT * FROM Win32_Thread")
For each threadObj in threadObjs
    Wscript.Echo _
        "Handle*" & threadObj.Handle & _
        "*Priority*" & threadObj.Priority & _
        "*PriorityBase*" & threadObj.PriorityBase & _
        "*ProcessHandle*" & threadObj.ProcessHandle & _
        "*StartAddress*" & threadObj.StartAddress & _
        "*ThreadState*" & threadObj.ThreadState & _
        "*ThreadWaitReason*" & threadObj.ThreadWaitReason
Next
