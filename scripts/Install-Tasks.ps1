﻿# This script was created using the following blog post from Xavier Plantefeve as a base:
# Source: https://xplantefeve.io/posts/SchdTskOnEvent
# I added the "WorkingDirectory = $PWD" for passing current working directory to Tasks' Start in entry.

# Go back one directory to set $CurrentDirectory for WorkingDirectory
cd..
$CurrentDirectory = $PWD

$class = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler

$trigger = $class | New-CimInstance -ClientOnly

$trigger.Enabled = $true

$trigger.Subscription = "<QueryList><Query Id='0' Path='Microsoft-Windows-CodeIntegrity/Operational'><Select Path='Microsoft-Windows-CodeIntegrity/Operational'>*[System[(EventID=3077)]]</Select></Query></QueryList>"

$ActionParameters = @{
    Execute  = 'WDACTaskHelper.exe'
    Argument = 'blocked'
    WorkingDirectory = $CurrentDirectory
}

$Action = New-ScheduledTaskAction @ActionParameters
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

$RegSchTaskParameters = @{
    TaskName    = 'WDAC-ToastBlocked'
    Description = 'Task that triggers a toast notification on WDAC block events.'
    TaskPath    = '\WDACTrayTool\'
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
    Execute  = 'WDACTaskHelper.exe'
    Argument = 'refresh'
    WorkingDirectory = $CurrentDirectory
}

$Action2 = New-ScheduledTaskAction @ActionParameters2
$Settings2 = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

$RegSchTaskParameters2 = @{
    TaskName    = 'WDAC-PolicyRefresh'
    Description = 'Task that triggers a toast notification on WDAC policy refresh events.'
    TaskPath    = '\WDACTrayTool\'
    Action      = $Action2
    Settings    = $Settings2
    Trigger     = $Trigger2
}

Register-ScheduledTask @RegSchTaskParameters2

$class3 = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler

$trigger3 = $class3 | New-CimInstance -ClientOnly

$trigger3.Enabled = $true

$trigger3.Subscription = "<QueryList><Query Id='0' Path='Microsoft-Windows-CodeIntegrity/Operational'><Select Path='Microsoft-Windows-CodeIntegrity/Operational'>*[System[(EventID=3076)]]</Select></Query></QueryList>"

$ActionParameters3 = @{
    Execute  = 'WDACTaskHelper.exe'
    Argument = 'audit'
    WorkingDirectory = $CurrentDirectory
}

$Action3 = New-ScheduledTaskAction @ActionParameters3
$Settings3 = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

$RegSchTaskParameters3 = @{
    TaskName    = 'WDAC-ToastAudit'
    Description = 'Task that triggers a toast notification on WDAC audit events.'
    TaskPath    = '\WDACTrayTool\'
    Action      = $Action3
    Settings    = $Settings3
    Trigger     = $Trigger3
}

Register-ScheduledTask @RegSchTaskParameters3
