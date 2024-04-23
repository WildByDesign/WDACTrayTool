Dim shell,command
command = "powershell.exe -nologo -noprofile D:\Tools\ToastNotificationScript\New-ToastNotification.ps1 -config D:\Tools\ToastNotificationScript\config-toast-policyrefresh.xml"
Set shell = CreateObject("WScript.Shell")
shell.Run command,0