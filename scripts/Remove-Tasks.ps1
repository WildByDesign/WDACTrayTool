Get-ScheduledTask | where TaskPath -eq "\WDACTrayTool\" | Unregister-ScheduledTask -Confirm:$false

$scheduleObject = New-Object -ComObject Schedule.Service
$scheduleObject.connect()
$rootFolder = $scheduleObject.GetFolder("\")
$rootFolder.DeleteFolder("WDACTrayTool",$null)
