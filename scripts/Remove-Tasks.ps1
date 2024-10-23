Get-ScheduledTask | where TaskPath -eq "\AppControlTray\" | Unregister-ScheduledTask -Confirm:$false

$scheduleObject = New-Object -ComObject Schedule.Service
$scheduleObject.connect()
$rootFolder = $scheduleObject.GetFolder("\")
$rootFolder.DeleteFolder("AppControlTray",$null)
