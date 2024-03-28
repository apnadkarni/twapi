# exec powershell.exe -Command {Get-CimInstance win32_service | select * | Export-Csv -Path y -NoTypeInformation -Delimiter "`t"}
Get-CimInstance Win32_Service | foreach-object -Process {
    $rec = "Name<*>" + $_.Name `
      + "<*>ServiceType<*>" + $_.ServiceType `
      + "<*>State<*>" + $_.State `
      + "<*>ExitCode<*>" + $_.ExitCode `
      + "<*>ProcessID<*>" + $_.ProcessID `
      + "<*>AcceptPause<*>" + $_.AcceptPause `
      + "<*>AcceptStop<*>" + $_.AcceptStop `
      + "<*>Caption<*>" + $_.Caption `
      + "<*>Description<*>" + $_.Description `
      + "<*>DesktopInteract<*>" + $_.DesktopInteract `
      + "<*>DisplayName<*>" + $_.DisplayName `
      + "<*>ErrorControl<*>" + $_.ErrorControl `
      + "<*>PathName<*>" + $_.PathName `
      + "<*>Started<*>" + $_.Started `
      + "<*>StartMode<*>" + $_.StartMode `
      + "<*>StartName<*>" + $_.StartName `
      + "<@>"

    Write-Host $rec
}
