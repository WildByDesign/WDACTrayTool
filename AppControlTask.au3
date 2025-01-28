#Region
#include <MsgBoxConstants.au3>
#include <String.au3>
#include <Array.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <String.au3>
#include <StringConstants.au3>
#EndRegion

#include "includes\TaskScheduler.au3"

#NoTrayIcon

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Task Manager
#AutoIt3Wrapper_Res_Fileversion=5.2.0.0
#AutoIt3Wrapper_Res_ProductVersion=5.2.0
#AutoIt3Wrapper_Res_ProductName=AppControlTaskManager
#AutoIt3Wrapper_Outfile_x64=AppControlTask.exe
#AutoIt3Wrapper_Res_LegalCopyright=@ 2025 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
$sTitle = "AppControlTaskHelper"

If @Compiled = 0 Then
	; System aware DPI awareness
	;DllCall("User32.dll", "bool", "SetProcessDPIAware")
	; Per-monitor V2 DPI awareness
	DllCall("User32.dll", "bool", "SetProcessDpiAwarenessContext" , "HWND", "DPI_AWARENESS_CONTEXT" -4)
EndIf

Func task_create_folder()
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	Global $sFolder = "\AppControlTray"
	Global $aTaskProperties = _TS_FolderCreate($oService, $sFolder)
	_TS_Close($oService)
Endfunc

Func task_create_blocked()
	Global $sFolder = "\AppControlTray"    ; Folder where to create the task
	Global $sName = "ToastBlocked"  ; Name of the task to create
	; *****************************************************************************
	; Prepare start and end date of the trigger. Format must be YYYY-MM-DDTHH:MM:SS
	; *****************************************************************************
	Global $sStartDateTime = _DateAdd("n", 2, _NowCalc())
	$sStartDateTime = StringReplace($sStartDateTime, "/", "-")
	$sStartDateTime = StringReplace($sStartDateTime, " ", "T")
	Global $sEndDateTime = _DateAdd("M", 4, _NowCalc())
	$sEndDateTime = StringReplace($sEndDateTime, "/", "-")
	$sEndDateTime = StringReplace($sEndDateTime, " ", "T")

	; *****************************************************************************
	; Connect to the Task Scheduler Service
	; *****************************************************************************
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Delete a task in the same folder with the same name
	; *****************************************************************************
	_TS_TaskDelete($oService, $sFolder & "\" & $sName)
	; If @error <> 0 And @error <> 2 Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskDelete returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Create a new task
	; *****************************************************************************
	; Create the Task Definition object
	Global $oTaskDefinition = _TS_TaskCreate($oService)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskCreate returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Set all task properties
	Global $aProperties[] = [ _
			"ACTIONS|Path|AppControlTask.exe", _
			"SETTINGS|DisallowStartIfOnBatteries|False", _
			"SETTINGS|StopIfGoingOnBatteries|False", _
			"SETTINGS|Compatibility|4", _
			"SETTINGS|ExecutionTimeLimit|PT0S", _
			"REGISTRATIONINFO|Author|" & @ComputerName & "\" & @UserName, _
			"REGISTRATIONINFO|Description|Task that triggers a toast notification on App Control block events.", _
			"ACTIONS|WorkingDirectory|" & @ScriptDir, _
			"ACTIONS|Arguments|blocked", _
			"TRIGGERS|Enabled|True", _
			"TRIGGERS|Type|" & $TASK_TRIGGER_EVENT, _
			"TRIGGERS|Subscription|<QueryList><Query Id='0' Path='Microsoft-Windows-CodeIntegrity/Operational'><Select Path='Microsoft-Windows-CodeIntegrity/Operational'>*[System[(EventID=3077)]]</Select></Query></QueryList>" _
			]
	_TS_TaskPropertiesSet($oTaskDefinition, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the TaskDefinition returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Create an Event trigger
	Global $oTrigger = _TS_TriggerCreate($oTaskDefinition, $TASK_TRIGGER_EVENT)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "Creating the Trigger returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	_TS_TaskPropertiesSet($oTrigger, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the Trigger returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Create an Action
	Global $oAction = _TS_ActionCreate($oTaskDefinition, $TASK_ACTION_EXEC)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "Creating the Action returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	_TS_TaskPropertiesSet($oAction, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the Action returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; List properties of the Task Definition
	Global $aTaskProperties = _TS_TaskPropertiesGet($oService, $oTaskDefinition,1 , True)
	If Not @error Then
		;_ArrayDisplay($aTaskProperties, "Properties of the task to be created")
	Else
		MsgBox($MB_ICONERROR, "_TS_TaskPropertiesGet", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	EndIf

	; Register the task
	Global $oTask = _TS_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskRegister returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	;MsgBox($MB_ICONINFORMATION, "_TS_TaskCreate", "Task " & $sName & " has been created!")

	_TS_Close($oService)
Endfunc

Func task_create_audit()
	Global $sFolder = "\AppControlTray"    ; Folder where to create the task
	Global $sName = "ToastAudit"  ; Name of the task to create
	; *****************************************************************************
	; Prepare start and end date of the trigger. Format must be YYYY-MM-DDTHH:MM:SS
	; *****************************************************************************
	Global $sStartDateTime = _DateAdd("n", 2, _NowCalc())
	$sStartDateTime = StringReplace($sStartDateTime, "/", "-")
	$sStartDateTime = StringReplace($sStartDateTime, " ", "T")
	Global $sEndDateTime = _DateAdd("M", 4, _NowCalc())
	$sEndDateTime = StringReplace($sEndDateTime, "/", "-")
	$sEndDateTime = StringReplace($sEndDateTime, " ", "T")

	; *****************************************************************************
	; Connect to the Task Scheduler Service
	; *****************************************************************************
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Delete a task in the same folder with the same name
	; *****************************************************************************
	_TS_TaskDelete($oService, $sFolder & "\" & $sName)
	; If @error <> 0 And @error <> 2 Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskDelete returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Create a new task
	; *****************************************************************************
	; Create the Task Definition object
	Global $oTaskDefinition = _TS_TaskCreate($oService)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskCreate returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Set all task properties
	Global $aProperties[] = [ _
			"ACTIONS|Path|AppControlTask.exe", _
			"SETTINGS|DisallowStartIfOnBatteries|False", _
			"SETTINGS|StopIfGoingOnBatteries|False", _
			"SETTINGS|Compatibility|4", _
			"SETTINGS|ExecutionTimeLimit|PT0S", _
			"REGISTRATIONINFO|Author|" & @ComputerName & "\" & @UserName, _
			"REGISTRATIONINFO|Description|Task that triggers a toast notification on App Control audit events.", _
			"ACTIONS|WorkingDirectory|" & @ScriptDir, _
			"ACTIONS|Arguments|audit", _
			"TRIGGERS|Enabled|True", _
			"TRIGGERS|Type|" & $TASK_TRIGGER_EVENT, _
			"TRIGGERS|Subscription|<QueryList><Query Id='0' Path='Microsoft-Windows-CodeIntegrity/Operational'><Select Path='Microsoft-Windows-CodeIntegrity/Operational'>*[System[(EventID=3076)]]</Select></Query></QueryList>" _
			]
	_TS_TaskPropertiesSet($oTaskDefinition, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the TaskDefinition returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Create an Event trigger
	Global $oTrigger = _TS_TriggerCreate($oTaskDefinition, $TASK_TRIGGER_EVENT)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "Creating the Trigger returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	_TS_TaskPropertiesSet($oTrigger, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the Trigger returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Create an Action
	Global $oAction = _TS_ActionCreate($oTaskDefinition, $TASK_ACTION_EXEC)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "Creating the Action returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	_TS_TaskPropertiesSet($oAction, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the Action returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; List properties of the Task Definition
	Global $aTaskProperties = _TS_TaskPropertiesGet($oService, $oTaskDefinition,1 , True)
	If Not @error Then
		;_ArrayDisplay($aTaskProperties, "Properties of the task to be created")
	Else
		MsgBox($MB_ICONERROR, "_TS_TaskPropertiesGet", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	EndIf

	; Register the task
	Global $oTask = _TS_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskRegister returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	;MsgBox($MB_ICONINFORMATION, "_TS_TaskCreate", "Task " & $sName & " has been created!")

	_TS_Close($oService)
Endfunc

Func task_create_refresh()
	Global $sFolder = "\AppControlTray"    ; Folder where to create the task
	Global $sName = "PolicyRefresh"  ; Name of the task to create
	; *****************************************************************************
	; Prepare start and end date of the trigger. Format must be YYYY-MM-DDTHH:MM:SS
	; *****************************************************************************
	Global $sStartDateTime = _DateAdd("n", 2, _NowCalc())
	$sStartDateTime = StringReplace($sStartDateTime, "/", "-")
	$sStartDateTime = StringReplace($sStartDateTime, " ", "T")
	Global $sEndDateTime = _DateAdd("M", 4, _NowCalc())
	$sEndDateTime = StringReplace($sEndDateTime, "/", "-")
	$sEndDateTime = StringReplace($sEndDateTime, " ", "T")

	; *****************************************************************************
	; Connect to the Task Scheduler Service
	; *****************************************************************************
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Delete a task in the same folder with the same name
	; *****************************************************************************
	_TS_TaskDelete($oService, $sFolder & "\" & $sName)
	; If @error <> 0 And @error <> 2 Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskDelete returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Create a new task
	; *****************************************************************************
	; Create the Task Definition object
	Global $oTaskDefinition = _TS_TaskCreate($oService)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskCreate returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Set all task properties
	Global $aProperties[] = [ _
			"ACTIONS|Path|AppControlTask.exe", _
			"SETTINGS|DisallowStartIfOnBatteries|False", _
			"SETTINGS|StopIfGoingOnBatteries|False", _
			"SETTINGS|Compatibility|4", _
			"SETTINGS|ExecutionTimeLimit|PT0S", _
			"REGISTRATIONINFO|Author|" & @ComputerName & "\" & @UserName, _
			"REGISTRATIONINFO|Description|Task that triggers a toast notification on App Control policy refresh events.", _
			"ACTIONS|WorkingDirectory|" & @ScriptDir, _
			"ACTIONS|Arguments|refresh", _
			"TRIGGERS|Enabled|True", _
			"TRIGGERS|Type|" & $TASK_TRIGGER_EVENT, _
			"TRIGGERS|Subscription|<QueryList><Query Id='0' Path='Microsoft-Windows-CodeIntegrity/Operational'><Select Path='Microsoft-Windows-CodeIntegrity/Operational'>*[System[(EventID=3102)]]</Select></Query></QueryList>" _
			]
	_TS_TaskPropertiesSet($oTaskDefinition, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the TaskDefinition returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Create an Event trigger
	Global $oTrigger = _TS_TriggerCreate($oTaskDefinition, $TASK_TRIGGER_EVENT)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "Creating the Trigger returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	_TS_TaskPropertiesSet($oTrigger, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the Trigger returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; Create an Action
	Global $oAction = _TS_ActionCreate($oTaskDefinition, $TASK_ACTION_EXEC)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "Creating the Action returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	_TS_TaskPropertiesSet($oAction, $aProperties)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskPropertiesSet for the Action returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; List properties of the Task Definition
	Global $aTaskProperties = _TS_TaskPropertiesGet($oService, $oTaskDefinition,1 , True)
	If Not @error Then
		;_ArrayDisplay($aTaskProperties, "Properties of the task to be created")
	Else
		MsgBox($MB_ICONERROR, "_TS_TaskPropertiesGet", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	EndIf

	; Register the task
	Global $oTask = _TS_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition)
	If @error Then Exit MsgBox($MB_ICONERROR, "_TS_TaskCreate", "_TS_TaskRegister returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	;MsgBox($MB_ICONINFORMATION, "_TS_TaskCreate", "Task " & $sName & " has been created!")

	_TS_Close($oService)
Endfunc

Func task_delete_blocked()
	; *****************************************************************************
	; Connect to the Task Scheduler Service
	; *****************************************************************************
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Delete task "Test-Task" from folder "\Test"
	; *****************************************************************************
	Global $sTaskPath = "\AppControlTray\ToastBlocked"    ; Folder and name of the task to be deleted
	_TS_TaskDelete($oService, $sTaskPath)
	If Not @error Then
		;MsgBox($MB_ICONINFORMATION, "_TS_TaskDelete", "Task '" & $sTaskPath & "' successfully deleted!")
	Else
		MsgBox($MB_ICONERROR, "_TS_TaskDelete", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	EndIf
	_TS_Close($oService)
Endfunc

Func task_delete_audit()
	; *****************************************************************************
	; Connect to the Task Scheduler Service
	; *****************************************************************************
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Delete task "Test-Task" from folder "\Test"
	; *****************************************************************************
	Global $sTaskPath = "\AppControlTray\ToastAudit"    ; Folder and name of the task to be deleted
	_TS_TaskDelete($oService, $sTaskPath)
	If Not @error Then
		;MsgBox($MB_ICONINFORMATION, "_TS_TaskDelete", "Task '" & $sTaskPath & "' successfully deleted!")
	Else
		MsgBox($MB_ICONERROR, "_TS_TaskDelete", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	EndIf
	_TS_Close($oService)
Endfunc

Func task_delete_refresh()
	; *****************************************************************************
	; Connect to the Task Scheduler Service
	; *****************************************************************************
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Delete task "Test-Task" from folder "\Test"
	; *****************************************************************************
	Global $sTaskPath = "\AppControlTray\PolicyRefresh"    ; Folder and name of the task to be deleted
	_TS_TaskDelete($oService, $sTaskPath)
	If Not @error Then
		;MsgBox($MB_ICONINFORMATION, "_TS_TaskDelete", "Task '" & $sTaskPath & "' successfully deleted!")
	Else
		MsgBox($MB_ICONERROR, "_TS_TaskDelete", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	EndIf
	_TS_Close($oService)
Endfunc

Func task_delete_folder()
	; *****************************************************************************
	; Connect to the Task Scheduler Service
	; *****************************************************************************
	Global $oService = _TS_Open()
	If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

	; *****************************************************************************
	; Delete a folder
	; $sFolder has always to start at the root folder (means you have to
	; specify all the folders from the root down even when they already exist)
	; *****************************************************************************
	Global $sFolder = "\AppControlTray"
	_TS_FolderDelete($oService, $sFolder)
	If Not @error Then
		;MsgBox($MB_ICONINFORMATION, "_TS_FolderDelete", "Folder: " & $sFolder & " successfully deleted!")
	Else
		MsgBox($MB_ICONERROR, "_TS_FolderDelete", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	EndIf
	_TS_Close($oService)
Endfunc

If $CmdLine[0] = 0 Then Exit MsgBox(16, $sTitle, "No parameters passed!")

If $CmdLine[1] = "install" Then
; install special shortcut with AppID
$cmdInstall1 = ' -install "App Control Tray Tool.lnk" '
$cmdInstall2 = @ScriptDir & '\AppControlTray.exe '
$cmdInstall3 = '"App Control Tray Tool"'
Run(@ScriptDir & "\toasts\snoretoast.exe" & $cmdInstall1 & $cmdInstall2 & $cmdInstall3, @ScriptDir & "\toasts", @SW_HIDE)
EndIf

If $CmdLine[1] = "blocked" Then
$cmdBlocked = ' -t "An application was stopped from running." -m "For security and stability reasons, unknown applications are prevented from running on this device." -appID "App Control Tray Tool" -p toast-blocked.png -silent -d short -b Dismiss'
Run(@ScriptDir & "\toasts\snoretoast.exe" & $cmdBlocked, @ScriptDir & "\toasts", @SW_HIDE)
EndIf

If $CmdLine[1] = "audit" Then
$cmdAudit = ' -t "An application would have been stopped from running." -m "Unknown applications are prevented from running. However, this was able to run due to Audit Mode." -appID "App Control Tray Tool" -p toast-audit.png -silent -d short -b Dismiss'
Run(@ScriptDir & "\toasts\snoretoast.exe" & $cmdAudit, @ScriptDir & "\toasts", @SW_HIDE)
EndIf

If $CmdLine[1] = "refresh" Then
$cmdRefresh = ' -t "Your App Control policies have been refreshed." -m "Code Integrity policy refresh finished successfully." -appID "App Control Tray Tool" -p toast-refresh.png -silent -d short -b Dismiss'
Run(@ScriptDir & "\toasts\snoretoast.exe" & $cmdRefresh, @ScriptDir & "\toasts", @SW_HIDE)
EndIf

If $CmdLine[1] = "convert" Then
	$mFile = FileOpenDialog("Select XML Policy File for Conversion", @ScriptDir, "Policy Files (*.xml)", 1)
	If @error Then
		Exit
	Else

	Local $sFileRead = FileRead($mFile)
	$aVersionEx = _StringBetween($sFileRead, "<VersionEx>", "</VersionEx>")
	$aExtract = _StringBetween($sFileRead, "<PolicyID>", "</PolicyID>")

	Local $versionsplit = StringSplit($aVersionEx[0], '.')
	;MsgBox($MB_SYSTEMMODAL, "version", $versionsplit[1] & '.' & $versionsplit[2] & '.' & $versionsplit[3] & '.' & $versionsplit[4])
	Local $versionplus = $versionsplit[4] + 1
	Local $versionincreased = $versionsplit[1] & '.' & $versionsplit[2] & '.' & $versionsplit[3] & '.' & $versionplus
	;MsgBox($MB_SYSTEMMODAL, "version increased", $versionincreased)

	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	Local $aPathSplit = _PathSplit($mFile, $sDrive, $sDir, $sFileName, $sExtension)

	Local $xmlfiledir = $sDrive & $sDir
	Local $xmlfilenew = $sFileName
	Local $xmlfilename2 = StringInStr($xmlfilenew, "-v")
	Local $aDays = StringSplit($xmlfilenew, "-v", 1)
	Local $xmlfilename3 = $aDays[1]
	Local $xmlupdatedname = $sDrive & $sDir & $aDays[1] & '-v' & $versionincreased & '.xml'
	;MsgBox($MB_SYSTEMMODAL, "Title", $xmlupdatedname)

	FileCopy($mFile, $xmlupdatedname, $FC_OVERWRITE)

	Local $binarysave = $sFileName & '.cip'
	Local $binaryname = $xmlfiledir & $sFileName & '.cip'
	Local $binarynameGUIDsave = $aExtract[0] & '.cip'
	Local $binarynameGUID = $xmlfiledir & $aExtract[0] & '.cip'

	EndIf

	Local Const $sMessage = "Choose a filename for saving your policy binary"
	Local $sFileSaveDialog = FileSaveDialog($sMessage, "::{450D8FBA-AD25-11D0-98A8-0800361B1103}", "Policy Files (*.cip)", $FD_PATHMUSTEXIST+$FD_PROMPTOVERWRITE, $binarynameGUIDsave)
	If @error Then
	Else
		Local $sFileName = StringTrimLeft($sFileSaveDialog, StringInStr($sFileSaveDialog, "\", $STR_NOCASESENSEBASIC, -1))

		Local $iExtension = StringInStr($sFileName, ".", $STR_NOCASESENSEBASIC)

		If $iExtension Then
			If Not (StringTrimLeft($sFileName, $iExtension - 1) = ".cip") Then $sFileSaveDialog &= ".cip"
		Else
			$sFileSaveDialog &= ".cip"
		EndIf
	EndIf

	Local $bDrive = "", $bDir = "", $bFileName = "", $bExtension = ""
	Local $bPathSplit = _PathSplit($sFileSaveDialog, $bDrive, $bDir, $bFileName, $bExtension)
	Local $xmlfiledir2 = $bDrive & $bDir
	Local $binarynameonly = $bFileName & $bExtension

	Local $binarynamechosen = $sFileSaveDialog
	; .\Convert-Policy.ps1 -XmlPolicyFile $mFile -BinaryDir $xmlfiledir -XmlOutName $xmlfilenew -BinaryFile $binarynameonly

	; Set-CIPolicyVersion -FilePath $XmlPolicyFileNew -Version $PolicyVersionNew
	; ConvertFrom-CIPolicy -XmlFilePath $XmlPolicyFileNew -BinaryFilePath "$BinaryDir\$BinaryFile"

	Local $quote = "'"
	Local $quote2 = '"'
	Local $cmd1 = ' -NoProfile -Command ' & $quote2 & '$CurrentDate = Get-Date -UFormat "%Y-%m-%d"; Set-CIPolicyIdInfo -FilePath ' & $quote & $xmlupdatedname & $quote & ' -PolicyId $CurrentDate; Set-CIPolicyVersion -FilePath ' & $quote & $xmlupdatedname & $quote & ' -Version ' & $quote & $versionincreased & $quote & '; ConvertFrom-CIPolicy -XmlFilePath ' & $quote & $xmlupdatedname & $quote & ' -BinaryFilePath ' & $quote & $binarynamechosen & $quote & $quote2
	Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
	Local $o_Pid1 = Run($o_powershell & $cmd1, @ScriptDir, @SW_Hide)
EndIf

If $CmdLine[1] = "installtasks" Then
	task_create_folder()
	Sleep(1000)
	task_create_blocked()
	task_create_audit()
	task_create_refresh()
EndIf

If $CmdLine[1] = "removetasks" Then
	task_delete_blocked()
	task_delete_audit()
	task_delete_refresh()
	Sleep(1000)
	task_delete_folder()
EndIf
