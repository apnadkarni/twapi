set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
Set servList = wmiObj.ExecQuery("SELECT * FROM Win32_Service")
For Each servObj in servList
    Wscript.StdOut.Write "Name<*>" & servObj.Name _
    & "<*>ServiceType<*>" & servObj.ServiceType _
    & "<*>State<*>" & servObj.State _
    & "<*>ExitCode<*>" & servObj.ExitCode _
    & "<*>ProcessID<*>" & servObj.ProcessID _
    & "<*>AcceptPause<*>" & servObj.AcceptPause _
    & "<*>AcceptStop<*>" & servObj.AcceptStop _
    & "<*>Caption<*>" & servObj.Caption _
    & "<*>Description<*>" & servObj.Description _
    & "<*>DesktopInteract<*>" & servObj.DesktopInteract _
    & "<*>DisplayName<*>" & servObj.DisplayName _
    & "<*>ErrorControl<*>" & servObj.ErrorControl _
    & "<*>PathName<*>" & servObj.PathName _
    & "<*>Started<*>" & servObj.Started _
    & "<*>StartMode<*>" & servObj.StartMode _
    & "<*>StartName<*>" & servObj.StartName _
    & "<@>"
        
Next

