# This script was created using the following blog post from Xavier Plantefeve as a base:
# Source: https://xplantefeve.io/posts/SchdTskOnEvent
# I added the "WorkingDirectory = $PWD" for passing current working directory to Tasks' Start in entry.

$class = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler

$trigger = $class | New-CimInstance -ClientOnly

$trigger.Enabled = $true

$trigger.Subscription = "<QueryList><Query Id='0' Path='Microsoft-Windows-CodeIntegrity/Operational'><Select Path='Microsoft-Windows-CodeIntegrity/Operational'>*[System[(EventID=3077)]]</Select></Query></QueryList>"

$ActionParameters = @{
    Execute  = 'wscript.exe'
    Argument = 'New-WDACToastNotification.vbs'
    WorkingDirectory = $PWD
}

$Action = New-ScheduledTaskAction @ActionParameters
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

$RegSchTaskParameters = @{
    TaskName    = 'WDAC-ToastBlocked'
    Description = 'Task that triggers a toast notification on WDAC block events.'
    TaskPath    = '\'
    Action      = $Action
    Settings    = $Settings
    Trigger     = $Trigger
}

Register-ScheduledTask @RegSchTaskParameters

$class2 = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler

$trigger2 = $class2 | New-CimInstance -ClientOnly

$trigger2.Enabled = $true

$trigger2.Subscription = "<QueryList><Query Id='0' Path='Microsoft-Windows-CodeIntegrity/Operational'><Select Path='Microsoft-Windows-CodeIntegrity/Operational'>*[System[(EventID=3102)]]</Select></Query></QueryList>"

$ActionParameters2 = @{
    Execute  = 'wscript.exe'
    Argument = 'New-WDACPolicyRefresh.vbs'
    WorkingDirectory = $PWD
}

$Action2 = New-ScheduledTaskAction @ActionParameters2
$Settings2 = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

$RegSchTaskParameters2 = @{
    TaskName    = 'WDAC-PolicyRefresh'
    Description = 'Task that triggers a toast notification on WDAC policy refresh events.'
    TaskPath    = '\'
    Action      = $Action2
    Settings    = $Settings2
    Trigger     = $Trigger2
}

Register-ScheduledTask @RegSchTaskParameters2
