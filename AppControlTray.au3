#Region
#include <MsgBoxConstants.au3>
#include <TrayConstants.au3>
#include <Array.au3>
#include <Misc.au3>
#EndRegion

#include "includes\ExtMsgBox.au3"

#include "includes\GuiCtrls_HiDpi.au3"
#include "includes\GUIDarkMode_v0.02mod.au3"

#include "includes\TaskScheduler.au3"

#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Tray Tool
#AutoIt3Wrapper_Res_Fileversion=3.1.0.0
#AutoIt3Wrapper_Res_ProductVersion=3.1.0
#AutoIt3Wrapper_Res_ProductName=AppControlTrayTool
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)

Global $softName = "App Control Tray Tool"
Global $trayIcon = "AppControl.ico"
Global $idRegTitleKey = "App Control Tray Tool"
Global $createdBy = "WildByDesign"
Global $programVersion = "3.1"

Global $LastTheme = ""
Global $tasksExist = ""

If _Singleton("AppControlTray", 1) = 0 Then
        $sMsg = " App Control Tray Tool is already running. " & @CRLF
		$sMsg &= " " & @CRLF
		$sMsg &= " Check for it in the system tray or hidden icon menu. " & @CRLF
		;MsgBox($MB_ICONWARNING, "Warning", $sMsg)
		_ExtMsgBoxSet(1, 4, -1, -1, -1, -1, 800)
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlTray.exe", 0, "App Control Tray Tool", $sMsg)
        Exit
EndIf

initial_set_theme()

Func initial_set_theme()
    ;$isDarkMode = regread('HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', 'AppsUseLightTheme') == 0 ? True : False
	Local $isDarkMode = _WinAPI_ShouldAppsUseDarkMode()
If $isDarkMode = True Then
	_ExtMsgBoxSet(Default)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, 10, "Cascadia Mono", 800)
	$LastTheme = "Dark"
	Local $hGUI = _HiDpi_GUICreate("Tray Tool", 100, 100)
    GuiDarkmodeApply($hGUI)
Else
	_ExtMsgBoxSet(Default)
	_ExtMsgBoxSet(1, 4, -1, -1, 10, "Cascadia Mono", 800)
	$LastTheme = "Light"
	Local $hGUI = _HiDpi_GUICreate("Tray Tool", 100, 100)
	Local $bEnableDarkTheme = False
    GuiLightmodeApply($hGUI)
EndIf
Endfunc

ThemeChanges()

Func ThemeChanges()
        ; Check for theme changes every 10 seconds
		AdlibRegister("query_theme_changes", 10000)
EndFunc

;query_theme_changes()

Func query_theme_changes()
    $isDarkMode = regread('HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', 'AppsUseLightTheme') == 0 ? True : False
	;Local $isDarkMode = _WinAPI_ShouldAppsUseDarkMode() ; can't get this method working for subsequent theme changes
If $isDarkMode = True And $LastTheme = "Light" Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 4, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, -1, "Consolas", 800)
	$LastTheme = "Dark"
	Local $hGUI = _HiDpi_GUICreate("Tray Tool", 100, 100)
    GuiDarkmodeApply($hGUI)
ElseIf $isDarkMode <> True And $LastTheme = "Dark" Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 4, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, -1, -1, -1, "Consolas", 800)
	$LastTheme = "Light"
	Local $hGUI = _HiDpi_GUICreate("Tray Tool", 100, 100)
	Local $bEnableDarkTheme = False
    GuiLightmodeApply($hGUI)
EndIf
Endfunc

_TasksExists()

Func _TasksExists()
; *****************************************************************************
; Connect to the Task Scheduler Service
; *****************************************************************************
Local $oService = _TS_Open()
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "Task Scheduler UDF", "Error connecting to the Task Scheduler Service. @error = " & @error & ", @extended = " & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))

; *****************************************************************************
; Check if a folder exists
; *****************************************************************************
Local $sFolder = "\AppControlTray"
Local $iExists = _TS_FolderExists($oService, $sFolder)
If Not @error Then
	;MsgBox($MB_ICONINFORMATION, "_TS_FolderExists", "Checked folder: " & $sFolder & @CRLF & "Exists: " & $iExists)
	Global $tasksExist = $iExists
Else
	;MsgBox($MB_ICONERROR, "_TS_FolderExists", "Returned @error=" & @error & ", @extended=" & @extended & @CRLF & @CRLF & _TS_ErrorText(@error))
	Global $tasksExist = $iExists
EndIf
_TS_Close($oService)
EndFunc

ShortcutExists()

Func ShortcutExists()
        Local $sFilePath = @AppDataDir & "\Microsoft\Windows\Start Menu\Programs\App Control Tray Tool.lnk"
		Local $iFileExists = FileExists($sFilePath)

        If $iFileExists Then
                ;MsgBox($MB_SYSTEMMODAL, "", "The file exists." & @CRLF & "FileExist returned: " & $iFileExists)
        Else
                ;MsgBox($MB_SYSTEMMODAL, "", "The file doesn't exist." & @CRLF & "FileExist returned: " & $iFileExists)
				Run(@ScriptDir & "\AppControlTask.exe install", @ScriptDir, @SW_HIDE)
        EndIf

EndFunc

; Set Current Working Directory
Local $setCWD = True
If $setCWD Then FileChangeDir(@ScriptDir)

; Menu
Local $idPStest = TrayCreateItem("App Control Policy Status")
Local $idCiTool = TrayCreateItem("CiTool Status (-lp)")
TrayCreateItem("")
Local $idWDACWizard = TrayCreateItem("App Control Wizard")
Local $idMsinfo32 = TrayCreateItem("System Information")
TrayCreateItem("")
Local $idViewLogs = TrayCreateMenu("View Logs")
Local $idLogsCI = TrayCreateItem("Code Integrity", $idViewLogs)
Local $idLogsScript = TrayCreateItem("MSI and Script", $idViewLogs)
TrayCreateItem("")
Local $idAddRemove = TrayCreateMenu("Change Policies")
Local $idAddPol = TrayCreateItem("Add or Update Policies (*.cip)", $idAddRemove)
Local $idRemovePol = TrayCreateItem("Remove Policies (*.cip)", $idAddRemove)
TrayCreateItem("")
Local $idToasts = TrayCreateItem("Enable Notifications")
;$idRegRead2 = RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", $idRegTitleKey)
If $tasksExist = '1' Then TrayItemSetState ($idToasts, $TRAY_CHECKED)
;If $tasksExist = '0' Then TrayItemSetState ($idToasts, $TRAY_UNCHECKED)
TrayItemSetOnEvent ($idToasts, "Notifications")
Local $idStartItem = TrayCreateItem("Start at Logon")
$idRegRead = RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", $idRegTitleKey)
If $idRegRead <> '' Then TrayItemSetState ($idStartItem, $TRAY_CHECKED)
TrayItemSetOnEvent ($idStartItem, "StartWithWindows")
;TrayCreateItem("")
Local $idAbout = TrayCreateItem("About")
Local $idExit = TrayCreateItem("Exit")

; Tray Icon - 201 is the resource number of embedded icon
; Set Dark Mode tray icon
TraySetIcon(@ScriptFullPath, 201)
;TraySetIcon($trayIcon)
;TraySetIcon(@ScriptFullPath, 203)
; Set Light Mode tray icon
;TraySetIcon(@ScriptFullPath, 202)

Func WDACWizard()
    Run(@ComSpec & " /c " & 'Explorer.Exe Shell:AppsFolder\Microsoft.WDAC.WDACWizard_8wekyb3d8bbwe!WDACWizard', "", @SW_HIDE)
EndFunc
Func Msinfo32()
    Run("C:\Windows\System32\msinfo32.exe", "", @SW_SHOWMAXIMIZED)
EndFunc
Func CiTool()
    Run(@ScriptDir & "\AppControlHelper.exe /CiTool")
EndFunc
Func LogsCI()
	Run(@ScriptDir & "\AppControlHelper.exe /LogsCI")
EndFunc
Func LogsScript()
	Run(@ScriptDir & "\AppControlHelper.exe /LogsScript")
EndFunc
Func AddPolicies()
    Run(@ScriptDir & "\AppControlHelper.exe /AddPolicies")
EndFunc
Func RemovePolicies()
    Run(@ScriptDir & "\AppControlHelper.exe /RemovePolicies")
EndFunc
Func Notifications()
    If @Compiled Then
        $_ItemGetState = TrayItemGetState ($idToasts)
        If $_ItemGetState = 64+1 Then
            ;Local $o_CmdString1 = ' ./Remove-Tasks.ps1'
			;Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
			;Local $o_Pid = Run($o_powershell & $o_CmdString1 , @ScriptDir & '\scripts', @SW_Hide)
			Run("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe ./Remove-Tasks.ps1", @ScriptDir & '\scripts', @SW_Hide)
            TrayItemSetState ($idToasts, $TRAY_UNCHECKED)
        Else
            ;Local $o_CmdString2 = ' ./Install-Tasks.ps1'
			;Local $o_powershell2 = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
			;Local $o_Pid2 = Run($o_powershell2 & $o_CmdString2 , @ScriptDir & '\scripts', @SW_Hide)
			Run("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe ./Install-Tasks.ps1", @ScriptDir & '\scripts', @SW_Hide)
            TrayItemSetState ($idToasts, $TRAY_CHECKED)
        EndIf
    EndIf
EndFunc
Func StartWithWindows()
    If @Compiled Then
        $_ItemGetState = TrayItemGetState ($idStartItem)
        If $_ItemGetState = 64+1 Then
            RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", $idRegTitleKey)
            TrayItemSetState ($idStartItem, $TRAY_UNCHECKED)
        Else
            RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", $idRegTitleKey,"REG_SZ",@ScriptFullPath)
            TrayItemSetState ($idStartItem, $TRAY_CHECKED)
        EndIf
    EndIf
EndFunc
Func PStest()
	Run(@ScriptDir & "\AppControlHelper.exe /status")
EndFunc


TraySetToolTip($softName)

While 1
    Switch TrayGetMsg()
        Case $idWDACWizard
            WDACWizard()
        Case $idMsinfo32
            Msinfo32()
        Case $idCiTool
            CiTool()
        Case $idLogsCI
            LogsCI()
        Case $idLogsScript
            LogsScript()
		Case $idAddPol
            AddPolicies()
		Case $idRemovePol
            RemovePolicies()
        Case $idPStest
            PStest()
		Case $idStartItem
            StartWithWindows()
		Case $idToasts
            Notifications()
		Case $idAbout
            ;_ExtMsgBoxSet(1, 4, -1, -1, -1, "Consolas", 800, 800)
			_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlTray.exe", 0, "About App Control Tray Tool", " Version: " & $programVersion & @CRLF & _
                            " Created by: " & $createdBy & @CRLF & @CRLF & _
                            " This program is free and open source. " & @CRLF)
        Case $idExit
        ExitLoop
    EndSwitch

WEnd
