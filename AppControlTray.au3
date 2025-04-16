#NoTrayIcon

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Tray Tool
#AutoIt3Wrapper_Res_Fileversion=5.5.0.0
#AutoIt3Wrapper_Res_ProductVersion=5.5.0
#AutoIt3Wrapper_Res_ProductName=AppControlTrayTool
#AutoIt3Wrapper_Outfile_x64=AppControlTray.exe
#AutoIt3Wrapper_Res_LegalCopyright=@ 2025 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#AutoIt3Wrapper_Res_Icon_Add=AppControl-Audit.ico
#AutoIt3Wrapper_Res_Icon_Add=AppControl-Disabled.ico
#AutoIt3Wrapper_Res_Icon_Add=AppControl-Blocked.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region

#include <MsgBoxConstants.au3>
#include <TrayConstants.au3>
#include <Array.au3>
#include <Misc.au3>
#include <FileConstants.au3>
#include <File.au3>
#include <String.au3>
#include <StringConstants.au3>
#EndRegion

#include "includes\ExtMsgBox.au3"

#include "includes\GuiCtrls_HiDpi.au3"
#include "includes\GUIDarkMode_v0.02mod.au3"

#include "includes\TaskScheduler.au3"

If @Compiled = 0 Then
	; System aware DPI awareness
	;DllCall("User32.dll", "bool", "SetProcessDPIAware")
	; Per-monitor V2 DPI awareness
	DllCall("User32.dll", "bool", "SetProcessDpiAwarenessContext" , "HWND", "DPI_AWARENESS_CONTEXT" -4)
EndIf

Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("TrayOnEventMode", 1)

Global $programVersion = "5.5"
;Global $softName = "App Control Tray Tool"
Global $trayIcon = "AppControl.ico"
Global $idRegTitleKey = "App Control Tray Tool"
Global $createdBy = "WildByDesign"

Global $LastTheme = ""
Global $tasksExist = ""

Global $WDACWizardExists, $MainFont

;$softName = 'App Control Tray Tool' & @CRLF & @CRLF
;$softName &= 'App Control policy' & @TAB & @TAB
;$softName &= ': Enforced Mode' & @CRLF
;$softName &= 'App Control user mode policy' & @TAB
;$softName &= ': Enforced Mode'

$tipTitle = 'App Control Tray Tool' & @CRLF & @CRLF
$tipKernelTitle = 'App Control policy' & @TAB & @TAB & ':'
$tipKernelStatus = ' Enforced Mode' & @CRLF
$tipUserTitle = 'App Control user mode policy' & @TAB & ':'
$tipUserStatus = ' Enforced Mode'

If _Singleton("AppControlTray", 1) = 0 Then
        $sMsg = " App Control Tray Tool is already running. " & @CRLF
		$sMsg &= " " & @CRLF
		$sMsg &= " Check for it in the system tray or hidden icon menu. " & @CRLF
		;MsgBox($MB_ICONWARNING, "Warning", $sMsg)
		_ExtMsgBoxSet(1, 4, -1, -1, -1, -1, 800)
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlTray.exe", 0, "App Control Tray Tool", $sMsg)
        Exit
EndIf


Local Const $sSegUIVar = @WindowsDir & "\fonts\SegUIVar.ttf"
Local $SegUIVarExists = FileExists($sSegUIVar)

If $SegUIVarExists Then
	Global $MainFont = "Segoe UI Variable Display"
	GUISetFont(8.5, $FW_NORMAL, -1, $MainFont)
Else
	Global $MainFont = "Segoe UI"
	GUISetFont(8.5, $FW_NORMAL, -1, $MainFont)
EndIf


initial_set_theme()

Func initial_set_theme()
    ;$isDarkMode = regread('HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', 'AppsUseLightTheme') == 0 ? True : False
	Local $isDarkMode = _WinAPI_ShouldAppsUseDarkMode()
If $isDarkMode = True Then
	_ExtMsgBoxSet(Default)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, 10, $MainFont, 800)
	$LastTheme = "Dark"
	Local $hGUI = _HiDpi_GUICreate("Tray Tool", 100, 100)
    GuiDarkmodeApply($hGUI)
Else
	_ExtMsgBoxSet(Default)
	_ExtMsgBoxSet(1, 4, -1, -1, 10, $MainFont, 800)
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

PolicyChanges()

Func PolicyChanges()
        ; Check for theme changes every 7 seconds
		AdlibRegister("policy_changes", 7000)
EndFunc


BlockEvents()
Func BlockEvents()
	; Check for theme changes every 3 seconds
	AdlibRegister("block_events", 3000)
EndFunc

;query_theme_changes()

Func query_theme_changes()
    $isDarkMode = regread('HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', 'AppsUseLightTheme') == 0 ? True : False
	;Local $isDarkMode = _WinAPI_ShouldAppsUseDarkMode() ; can't get this method working for subsequent theme changes
If $isDarkMode = True And $LastTheme = "Light" Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 4, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, 10, "Cascadia Mono", 800)
	$LastTheme = "Dark"
	Local $hGUI = _HiDpi_GUICreate("Tray Tool", 100, 100)
    GuiDarkmodeApply($hGUI)
ElseIf $isDarkMode <> True And $LastTheme = "Dark" Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 4, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, -1, -1, 10, "Cascadia Mono", 800)
	$LastTheme = "Light"
	Local $hGUI = _HiDpi_GUICreate("Tray Tool", 100, 100)
	Local $bEnableDarkTheme = False
    GuiLightmodeApply($hGUI)
EndIf
Endfunc

policy_changes()
Func policy_changes()
	Global $dString14 = ""
	Local $o_CmdString1 = " Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity*"
	Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Command"
	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	$test = stringsplit($out , @CR , 2)
	_ArrayDelete($test, 4)
	_ArrayDelete($test, 4)
	$test2 = _ArrayToString($test)
	Local $dString = StringStripWS($test2, $STR_STRIPSPACES)
	Local $dString1 = StringReplace($dString, " |", "")
	Local $dString2 = StringReplace($dString1, "|", "")
	Local $dString3 = StringStripWS($dString2, $STR_STRIPLEADING + $STR_STRIPTRAILING)
	Local $dString4 = StringReplace($dString3, " ", "  ")
	Local $dString5 = StringReplace($dString4, "test  output", "")
	Local $dString6 = StringReplace($dString5, "Currently  Active  Policies:", "Currently Active Policies:")
	Local $dString7 = StringReplace($dString6, "PolicyID  FriendlyName", "Policy ID                             Policy Name")
	Local $dString8 = StringReplace($dString7, "--------  ------------", "---------                             -----------")
	Local $dString9 = StringReplace($dString8, "UsermodeCodeIntegrityPolicyEnforcementStatus  :  0", 'App Control user mode policy' & @TAB & ':' & ' Not Configured')
	Local $dString10 = StringReplace($dString9, "UsermodeCodeIntegrityPolicyEnforcementStatus  :  1", 'App Control user mode policy' & @TAB & ':' & ' Audit Mode')
	Local $dString11 = StringReplace($dString10, "UsermodeCodeIntegrityPolicyEnforcementStatus  :  2", 'App Control user mode policy' & @TAB & ':' & ' Enforced Mode')
	Local $dString12 = StringReplace($dString11, "CodeIntegrityPolicyEnforcementStatus  :  0", 'App Control policy' & @TAB & @TAB & ':' & ' Not Configured')
	Local $dString13 = StringReplace($dString12, "CodeIntegrityPolicyEnforcementStatus  :  1", 'App Control policy' & @TAB & @TAB & ':' & ' Audit Mode')
	Global $dString14 = StringReplace($dString13, "CodeIntegrityPolicyEnforcementStatus  :  2", 'App Control policy' & @TAB & @TAB & ':' & ' Enforced Mode')
	TraySetToolTip($tipTitle & $dString14)
	Local $iAudit = StringInStr($dString14, "Audit Mode")
	Local $iEnforced = StringInStr($dString14, "Enforced Mode")
	Local $iDisabled = StringInStr($dString14, "Not Configured")
	$iSignal = RegRead("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked")
	If $iSignal = '1' Then
		; don't change icon
	Else
		If $iAudit <> '0' Then
			TraySetIcon(@ScriptFullPath, 202)
		ElseIf $iEnforced <> '0' Then
			TraySetIcon(@ScriptFullPath, 201)
		ElseIf $iDisabled <> '0' Then
			TraySetIcon(@ScriptFullPath, 203)
		EndIf
	EndIf
Endfunc


Func block_events()
	$iSignal = RegRead("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked")
	If $iSignal = '1' Then
		TraySetIcon(@ScriptFullPath, 204)
	EndIf
EndFunc


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


; Set Current Working Directory
Local $setCWD = True
If $setCWD Then FileChangeDir(@ScriptDir)

; Menu
;Local $idPStest = TrayCreateItem("App Control Policy Status")
Local $idPStest = TrayCreateItem("App Control Policy Manager")
TrayItemSetOnEvent(-1, "PStest")
Local $idCiTool = TrayCreateItem("CiTool Status (-lp)")
TrayItemSetOnEvent(-1, "CiTool")
TrayCreateItem("")
Local $idWDACWizard = TrayCreateItem("App Control Wizard")
TrayItemSetOnEvent(-1, "WDACWizard")
Local $idMsinfo32 = TrayCreateItem("System Information")
TrayItemSetOnEvent(-1, "Msinfo32")
TrayCreateItem("")
Local $idViewLogs = TrayCreateMenu("View Logs")
Local $idLogsCI = TrayCreateItem("Code Integrity", $idViewLogs)
TrayItemSetOnEvent(-1, "LogsCI")
Local $idLogsScript = TrayCreateItem("MSI and Script", $idViewLogs)
TrayItemSetOnEvent(-1, "LogsScript")
TrayCreateItem("")
Local $idAddRemove = TrayCreateMenu("Change Policies")
Local $idAddPol = TrayCreateItem("Add or Update Policies", $idAddRemove)
TrayItemSetOnEvent(-1, "AddPolicies")
Local $idRemovePol = TrayCreateItem("Remove Policies", $idAddRemove)
TrayItemSetOnEvent(-1, "RemovePolicies")
TrayCreateItem("")
Local $idConvertPol = TrayCreateItem("Convert (xml to binary)")
TrayItemSetOnEvent(-1, "ConvertPolicy")
TrayCreateItem("")
Local $idToasts = TrayCreateItem("Enable Notifications")
If $tasksExist = '1' Then TrayItemSetState ($idToasts, $TRAY_CHECKED)
TrayItemSetOnEvent ($idToasts, "Notifications")
Local $idStartItem = TrayCreateItem("Start at Logon")
$idRegRead = RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", $idRegTitleKey)
If $idRegRead <> '' Then TrayItemSetState ($idStartItem, $TRAY_CHECKED)
TrayItemSetOnEvent ($idStartItem, "StartWithWindows")
Local $idAbout = TrayCreateItem("About")
TrayItemSetOnEvent(-1, "About")
Local $idExit = TrayCreateItem("Exit")
TrayItemSetOnEvent(-1, "ExitScript")

; Tray Icon - 201 is the resource number of embedded icon
;TraySetIcon(@ScriptFullPath, 201)
; Audit Mode icon
;TraySetIcon(@ScriptFullPath, 202)


Func WDACWizardExists()
	Local $o_CmdString1 = " Get-AppxPackage -Name 'Microsoft.WDAC.WDACWizard'"
	Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Command"
	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	Local $iString = StringInStr($out, "Microsoft.WDAC.WDACWizard")
	If $iString = 0 Then
		$WDACWizardExists = False
	Else
		$WDACWizardExists = True
	EndIf
EndFunc


Func WDACWizard()
	If $WDACWizardExists = True Then
		Run(@ComSpec & " /c " & 'Explorer.Exe Shell:AppsFolder\Microsoft.WDAC.WDACWizard_8wekyb3d8bbwe!WDACWizard', "", @SW_HIDE)
	Else
		$sMsg = " App Control Wizard doesn't appear to be installed. " & @CRLF
		$sMsg &= " " & @CRLF
		$sMsg &= " Would you like to visit the App Control Wizard download page? " & @CRLF
		$sMsg &= " " & @CRLF
		$iRetValue = _ExtMsgBox (0, 4, "App Control Wizard", $sMsg)

		If $iRetValue = 1 Then
			;ConsoleWrite("Yes" & @CRLF)
			ShellExecute("https://webapp-wdac-wizard.azurewebsites.net/")
		ElseIf $iRetValue = 2 Then
			;ConsoleWrite("No" & @CRLF)
		EndIf
		; check if user installs Wizard
		CheckWizardInstall()
	EndIf
EndFunc

Func CheckWizardInstall()
	; Check for Wizard install every 11 seconds
	AdlibRegister("WDACWizardExists", 11000)
EndFunc

Func Msinfo32()
    Run("C:\Windows\System32\msinfo32.exe", "", @SW_SHOWMAXIMIZED)
EndFunc

Func CiTool()
    ;Run(@ScriptDir & "\AppControlHelper.exe /CiTool")
	ShellExecute("AppControlHelper.exe", "/CiTool", @ScriptDir)
EndFunc

Func LogsCI()
	;Run(@ScriptDir & "\AppControlHelper.exe /LogsCI")
	ShellExecute("AppControlHelper.exe", "/LogsCI", @ScriptDir)
EndFunc

Func LogsScript()
	;Run(@ScriptDir & "\AppControlHelper.exe /LogsScript")
	ShellExecute("AppControlHelper.exe", "/LogsScript", @ScriptDir)
EndFunc

Func AddPolicies()
    ;Run(@ScriptDir & "\AppControlHelper.exe /AddPolicies")
	ShellExecute("AppControlHelper.exe", "/AddPolicies", @ScriptDir)
EndFunc

Func RemovePolicies()
    ;Run(@ScriptDir & "\AppControlHelper.exe /RemovePolicies")
	ShellExecute("AppControlHelper.exe", "/RemovePolicies", @ScriptDir)
EndFunc

Func ConvertPolicy()
	Run(@ScriptDir & "\AppControlTask.exe convert")
EndFunc

Func Notifications()
    If @Compiled Then
        $_ItemGetState = TrayItemGetState ($idToasts)
        If $_ItemGetState = 64+1 Then
			Run("AppControlTask.exe removetasks", @ScriptDir, @SW_Hide)
            TrayItemSetState ($idToasts, $TRAY_UNCHECKED)
        Else
			Run("AppControlTask.exe installtasks", @ScriptDir, @SW_Hide)
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
	;Run(@ScriptDir & "\AppControlManager.exe")
	ShellExecute("AppControlManager.exe", "", @ScriptDir)
EndFunc

TraySetToolTip($tipTitle & $dString14)
;TraySetToolTip($tipTitle & $tipKernelTitle & $tipKernelStatus & $tipUserTitle & $tipUserStatus)

; check whether or not App Control Wizard is installed
WDACWizardExists()


Func ProcessWaitCloseEx($iPID)
	While ProcessExists($iPID) And Sleep(10)
	WEnd
EndFunc


TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "TrayEvent")
TraySetOnEvent($TRAY_EVENT_SECONDARYUP, "TrayEvent")


While 1
	Sleep(100)
WEnd


Func TrayEvent()
	Switch @TRAY_ID
			Case $TRAY_EVENT_PRIMARYUP
				$iSignal = RegRead("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked")
				If $iSignal = '1' Then
					$iRegWrite = RegWrite("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked", "REG_DWORD", "0")
					policy_changes()
				EndIf
			Case $TRAY_EVENT_SECONDARYUP
				$iSignal = RegRead("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked")
				If $iSignal = '1' Then
					$iRegWrite = RegWrite("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked", "REG_DWORD", "0")
					policy_changes()
				EndIf
	EndSwitch
EndFunc


Func ExitScript()
	Exit
EndFunc


Func About()
	_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlTray.exe", 0, "About App Control Tray Tool", " Version: " & $programVersion & @CRLF & _
                            " Created by: " & $createdBy & @CRLF & @CRLF & _
                            " This program is free and open source. " & @CRLF)
EndFunc
