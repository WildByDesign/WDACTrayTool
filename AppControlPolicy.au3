#NoTrayIcon
#RequireAdmin

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl-App.ico
#AutoIt3Wrapper_Outfile_x64=AppControlPolicy.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Policy Manager
#AutoIt3Wrapper_Res_Fileversion=6.0.2
#AutoIt3Wrapper_Res_ProductName=AppControlPolicyManager
#AutoIt3Wrapper_Res_ProductVersion=6.0.2
#AutoIt3Wrapper_Res_LegalCopyright=@ 2025 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl-App.ico
#AutoIt3Wrapper_Res_Icon_Add=unchecked.ico
#AutoIt3Wrapper_Res_Icon_Add=checked.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; *** Start added Standard Include files by AutoIt3Wrapper ***
#include <Array.au3>
#include <GDIPlus.au3>
#include <GuiListView.au3>
#include <GuiStatusBar.au3>
#include <GuiToolTip.au3>
#include <File.au3>
#include <String.au3>
#include <WinAPIGdi.au3>
#include <WinAPISysWin.au3>
#include <WinAPISysInternals.au3>
#include <WinAPIRes.au3>
#include <WinAPIGdiDC.au3>
#include <WinAPIInternals.au3>
#include <WinAPIGdiInternals.au3>
#include <WinAPITheme.au3>
#include <APIResConstants.au3>
#include <AutoItConstants.au3>
#include <FileConstants.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <HeaderConstants.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <StatusBarConstants.au3>
#include <StringConstants.au3>
#include <StructureConstants.au3>
#include <WindowsConstants.au3>
; *** End added Standard Include files by AutoIt3Wrapper ***

#include "includes\ExtMsgBox.au3"
#include "includes\GUIDarkMode_v0.02mod.au3"
#include "includes\GUIListViewEx.au3"
#include "includes\XML.au3"
#include "includes\ModernMenuRaw.au3"
#include "includes\_GUICtrlListView_SaveCSV.au3"
#include "includes\libnotif.au3"

Global $programversion = "6.0.1"

If @Compiled = 0 Then
	; System aware DPI awareness
	;DllCall("User32.dll", "bool", "SetProcessDPIAware")
	; Per-monitor V2 DPI awareness
	DllCall("User32.dll", "bool", "SetProcessDpiAwarenessContext" , "HWND", "DPI_AWARENESS_CONTEXT" -4)
EndIf


isDarkMode()
Func isDarkMode()
	Global $isDarkMode = _WinAPI_ShouldAppsUseDarkMode()
Endfunc

If $isDarkMode = True Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, 9, -1, 1200)
Else
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, -1, -1, 9, -1, 1200)
EndIf

Global $hGUI, $cListView, $hListView, $EFIarray, $hGUI2, $ExitButtonEFI, $EFIListView, $hEFIListView, $PolicyStatusInfo
Global $out, $arpol, $arraycount, $policycount, $policyoutput, $policycorrect, $CountTotal, $ExitButton, $TestButton, $Label2, $aContent, $iLV_Index
Global $CountEnforced = 0
Global $WDACWizardExists

Global Const $SBS_SIZEBOX = 0x08, $SBS_SIZEGRIP = 0x10

$idLightBk = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))

If $isDarkMode = True Then
	Global $g_iBkColor = 0x2c2c2c, $g_iTextColor = 0xffffff
Else
	Global $g_iBkColor = $idLightBk, $g_iTextColor = 0x000000
EndIf

Global $g_hSizebox, $g_hOldProc, $g_hStatus, $g_iHeight, $g_aText, $g_aRatioW, $g_hDots


Local Const $sSegUIVar = @WindowsDir & "\fonts\SegUIVar.ttf"
Local $SegUIVarExists = FileExists($sSegUIVar)

If $SegUIVarExists Then
	Global $MainFont = "Segoe UI Variable Display"
	;GUISetFont(8.5, $FW_NORMAL, -1, $MainFont)
Else
	Global $MainFont = "Segoe UI"
	;GUISetFont(8.5, $FW_NORMAL, -1, $MainFont)
EndIf


;---

; configure library
_LibNotif_Config($LIBNOTIF_CFG_TIMEOUT, 10000)

_LibNotif_Config($LIBNOTIF_CFG_MINTOASTWIDTH, 128)
;_LibNotif_Config($LIBNOTIF_CFG_POSITION, $LIBNOTIF_POS_UR)
_LibNotif_Config($LIBNOTIF_CFG_MAXTEXTWIDTH, @DesktopWidth - 100)

_LibNotif_Config($LIBNOTIF_CFG_CLOSEBUTTON, True)

$idLightBk = _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))

If $isDarkMode = True Then
	_LibNotif_Config($LIBNOTIF_CFG_BKCOLOR, 0x202020)
	_LibNotif_Config($LIBNOTIF_CFG_TEXTCOLOR, 0xe0e0e0)
Else
	_LibNotif_Config($LIBNOTIF_CFG_BKCOLOR, $idLightBk)
	_LibNotif_Config($LIBNOTIF_CFG_TEXTCOLOR, 0x000000)
EndIf

_LibNotif_Config($LIBNOTIF_CFG_TEXTFONTNAME, $MainFont)
_LibNotif_Config($LIBNOTIF_CFG_TEXTSIZE, 10)
_LibNotif_Config($LIBNOTIF_CFG_TEXTWEIGHT, 400)
_LibNotif_Config($LIBNOTIF_CFG_TEXTATTRIBUTES, 0)

_LibNotif_Config($LIBNOTIF_CFG_TEXTMARGIN, 15)
_LibNotif_Config($LIBNOTIF_CFG_BUTTONSIZE, 30)

;_LibNotif_Config($LIBNOTIF_CFG_SHOW_ANIMSTYLE, $AW_VER_NEGATIVE)
;_LibNotif_Config($LIBNOTIF_CFG_HIDE_ANIMSTYLE, $AW_VER_POSITIVE + $AW_HIDE)

; init library
_LibNotif_Init()

;---


If $isDarkMode = True Then
Global $iDllGDI = DllOpen("gdi32.dll")
Global $iDllUSER32 = DllOpen("user32.dll")

;Three column colours
Global $aCol[11][2] = [[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff]]

;Convert RBG to BGR for SetText/BkColor()
For $i = 0 To UBound($aCol)-1
    $aCol[$i][0] = _BGR2RGB($aCol[$i][0])
    $aCol[$i][1] = _BGR2RGB($aCol[$i][1])
Next
EndIf

; Fake older build for testing
;Global $WinBuild = "22621"

Global $WinBuild = @OSBuild
If $WinBuild >= 26100 Then
	Global $is24H2 = True
	;MsgBox($MB_SYSTEMMODAL, "Title", "This is 24H2 or newer")
Else
	Global $is24H2 = False
	;MsgBox($MB_SYSTEMMODAL, "Title", "This is not 24H2")
EndIf

Global $ifPS7Exists = FileExists(@ProgramFilesDir & '\PowerShell\7\pwsh.exe')
If $ifPS7Exists Then
	Global $o_powershell = @ProgramFilesDir & '\PowerShell\7\pwsh.exe -NoProfile -Command'
Else
	Global $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Command"
EndIf

GetPolicyInfo()
Func GetPolicyInfo()
	If $is24H2 = True Then
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
	Else
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
		; Pre-24H2 Simulation using older CiTool
		;Local $o_CmdString1 = " (.\CiTool\22621\CiTool.exe -lp -json | ConvertFrom-Json).Policies | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
	EndIf

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	CreatePolicyTable($out)
EndFunc

Func GetPolicyEnforced()
	If $is24H2 = True Then
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
	Else
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
		; Pre-24H2 Simulation using older CiTool
		;Local $o_CmdString1 = " (.\CiTool\22621\CiTool.exe -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
	EndIf

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)

	;;;

	; Testing in case of parsing issues

	;Local Const $sFilePath = @DesktopDir & "\test-output.txt"
	;Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
	;FileWrite($hFileOpen, $out)
	;FileClose($hFileOpen)

	;Local Const $sFilePath2 = @DesktopDir & "\test-parsed.txt"
	;Local $hFileOpen2 = FileOpen($sFilePath2, $FO_OVERWRITE)
	;FileWrite($hFileOpen2, $finalparse)
	;FileClose($hFileOpen2)

	;;;

	CreatePolicyTable($out)
EndFunc

Func GetPolicyBase()
	If $is24H2 = True Then
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -eq $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
	Else
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -eq $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
		; Pre-24H2 Simulation using older CiTool
		;Local $o_CmdString1 = " (.\CiTool\22621\CiTool.exe -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -eq $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
	EndIf

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	CreatePolicyTable($out)
EndFunc

Func GetPolicySupp()
	If $is24H2 = True Then
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -ne $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
	Else
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -ne $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
		; Pre-24H2 Simulation using older CiTool
		;Local $o_CmdString1 = " (.\CiTool\22621\CiTool.exe -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -ne $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
	EndIf

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	CreatePolicyTable($out)
EndFunc

Func GetPolicySigned()
	Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsSignedPolicy -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	CreatePolicyTable($out)
EndFunc

Func GetPolicySystem()
	If $is24H2 = True Then
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsSystemPolicy -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
	Else
		Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsSystemPolicy -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
		; Pre-24H2 Simulation using older CiTool
		;Local $o_CmdString1 = " (.\CiTool\22621\CiTool.exe -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsSystemPolicy -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,Version,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
	EndIf

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	CreatePolicyTable($out)
EndFunc

Func CreatePolicyTable($out_passed)
	If $is24H2 = True Then

		Local $string0 = StringStripWS($out_passed, $STR_STRIPSPACES + $STR_STRIPLEADING + $STR_STRIPTRAILING)
		Local $string1 = StringReplace($string0, "FriendlyName : ", "")
		Local $string2 = StringReplace($string1, "BasePolicyID : ", "")
		Local $string3 = StringReplace($string2, "PolicyID : ", "")
		Local $string4 = StringReplace($string3, "VersionString : ", "")
		Local $string5 = StringReplace($string4, "IsEnforced : ", "")
		Local $string6 = StringReplace($string5, "IsOnDisk : ", "")
		Local $string7 = StringReplace($string6, "IsSignedPolicy : ", "")
		Local $string8 = StringReplace($string7, "IsSystemPolicy : ", "")
		Local $string9 = StringReplace($string8, "IsAuthorized : ", "")
		Local $string10 = StringReplace($string9, "PolicyOptions : ", "")
		Local $string11 = StringReplace($string10, "[32;1m", "")
		Local $string12 = StringReplace($string11, "[0m", "")
		Local $string13 = StringReplace($string12, "{", "")
		Local $string14 = StringReplace($string13, "}", "")
		Local $string15 = StringReplace($string14, "Version : ", "")
		Local $string16 = StringReplace($string15, "...", "")
		Local $string17 = StringReplace($string16, "Default Policy.", "Default Policy")
		Local $string18 = StringReplace($string17, "Integrity Policy.", "Integrity Policy")
		Local $string19 = StringReplace($string18, "Supplemental Policies.", "Supplemental Policies")
		Local $string20 = StringReplace($string19, "Code Trust.", "Code Trust")
		Local $string21 = StringReplace($string20, "Rule Protection.", "Rule Protection")
		Local $finalparse = $string21

	Else

		;Local $string0 = StringStripWS($out_passed, $STR_STRIPSPACES + $STR_STRIPLEADING + $STR_STRIPTRAILING)
		Local $string0 = StringStripWS($out_passed, $STR_STRIPLEADING + $STR_STRIPTRAILING)
		Local $string1 = StringReplace($string0, "FriendlyName   : ", "")
		Local $string2 = StringReplace($string1, "BasePolicyID   : ", "")
		Local $string3 = StringReplace($string2, "PolicyID       : ", "")
		Local $string4 = StringReplace($string3, "VersionString  : ", "*")
		Local $string5 = StringReplace($string4, "IsEnforced     : ", "")
		Local $string6 = StringReplace($string5, "IsOnDisk       : ", "")
		Local $string7 = StringReplace($string6, "IsSignedPolicy : ", "*")
		Local $string8 = StringReplace($string7, "IsSystemPolicy : ", "")
		Local $string9 = StringReplace($string8, "IsAuthorized   : ", "")
		Local $string10 = StringReplace($string9, "PolicyOptions  : ", "*")
		Local $string11 = StringReplace($string10, "[32;1m", "")
		Local $string12 = StringReplace($string11, "[0m", "")
		Local $string13 = StringReplace($string12, "{", "")
		Local $string14 = StringReplace($string13, "}", "")
		Local $string15 = StringReplace($string14, "Version        : ", "")
		Local $string16 = StringStripWS($string15, $STR_STRIPSPACES)
		Local $string17 = StringReplace($string16, "PolicyOptions :", "*")
		Local $string18 = StringReplace($string17, "...", "")
		Local $string19 = StringReplace($string18, "Default Policy.", "Default Policy")
		Local $string20 = StringReplace($string19, "Integrity Policy.", "Integrity Policy")
		Local $string21 = StringReplace($string20, "Supplemental Policies.", "Supplemental Policies")
		Local $string22 = StringReplace($string21, "Code Trust.", "Code Trust")
		Local $string23 = StringReplace($string22, "Rule Protection.", "Rule Protection")
		Local $finalparse = $string23

	EndIf


	;MsgBox($MB_SYSTEMMODAL, "Title", "Final parse = " & $finalparse)

	;;;

	; Testing in case of parsing issues

	;Local Const $sFilePath = @DesktopDir & "\test-output.txt"
	;Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
	;FileWrite($hFileOpen, $out_passed)
	;FileClose($hFileOpen)

	;Local Const $sFilePath2 = @DesktopDir & "\test-parsed.txt"
	;Local $hFileOpen2 = FileOpen($sFilePath2, $FO_OVERWRITE)
	;FileWrite($hFileOpen2, $finalparse)
	;FileClose($hFileOpen2)

	;;;

	If $finalparse = "" Then
		Global $policycount = 0
	Else
		$arpol = stringsplit($finalparse, @CR , 0)

		;_ArrayDisplay($arpol, "test")

		Global $policyoutput = UBound ($arpol, $UBOUND_ROWS)
		Global $policycorrect = $policyoutput - 1
		Global $policycount = $policycorrect / 10
	EndIf

	; create policy array (pixelsearch)
	; support more policies by increasing $policycount <= 64
	Select
		Case $policycount = 0
			Global $aWords[1][11]
		Case IsInt($policycount) And $policycount >= 1 And $policycount <= 64
			Global $aWords[$policycount][11]
			For $i = 0 To $policycount - 1
				For $j = 1 To 10
					$aWords[$i][$j] = $arpol[$i*10 + $j]
				Next
			Next
		Case Else
			Global $aWords[1][11] = [["", "Error:", "Error " & $policycorrect & " &" & " Error " & $policycount, "", "", "", "", "", "", "", ""]]
	EndSelect


	If $is24H2 = False Then
		For $i = 0 To UBound($aWords) -1
			$iVersion = $aWords[$i][5]
			Local $tUInt64  = DllStructCreate("uint64 value;"), _
				$tVersion = DllStructCreate("word revision; word build; word minor; word major", DllStructGetPtr($tUInt64))

			$tUInt64.value = $iVersion

			Local $iVersionString = StringFormat('%i.%i.%i.%i', $tVersion.major, $tVersion.minor, $tVersion.build, $tVersion.revision)

			$aWords[$i][5] = $iVersionString
		Next
	EndIf
Endfunc


VulnerableDriver()
Func VulnerableDriver()
	Local $sVulnDriver = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\CI\Config", "VulnerableDriverBlocklistEnable")
	If $sVulnDriver = '1' Then
		Global $sVulnDrivermsg = " Vulnerable Driver Blocklist  : Enabled"
	ElseIf $sVulnDriver = '0' Then
		Global $sVulnDrivermsg = " Vulnerable Driver Blocklist  : Disabled"
	Else
		Global $sVulnDrivermsg = " Vulnerable Driver Blocklist  : Unknown"
	EndIf
Endfunc


;GetPolicyStatus()

Func GetPolicyStatus()
	Local $sSmartAppControl = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\CI\Policy", "VerifiedAndReputablePolicyState")
	If $sSmartAppControl = 0 Then
		$sSACStatus = "  Smart App Control (Off)"
	ElseIf $sSmartAppControl = 1 Then
		$sSACStatus = "  Smart App Control (Enforced)"
	ElseIf $sSmartAppControl = 2 Then
		$sSACStatus = "  Smart App Control (Evaluation)"
	Else
		$sSACStatus = "  Smart App Control (undetermined)"
	EndIf

	Local $sVulnerableDriver = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\CI\Config", "VulnerableDriverBlocklistEnable")
	If $sVulnerableDriver = 0 Then
		$sVDriverStatus = "Vulnerable Driver Blocklist (Disabled)"
	ElseIf $sVulnerableDriver = 1 Then
		$sVDriverStatus = "Vulnerable Driver Blocklist (Enabled)"
	Else
		$sVDriverStatus = "Vulnerable Driver Blocklist (undetermined)"
	EndIf	

	Local $o_CmdString1 = " Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity*"
	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)

	Local $topstatus1 = StringReplace($out, "[32;1m", "")
	Local $topstatus2 = StringReplace($topstatus1, "[0m", "")
	Local $topstatus3 = StringReplace($topstatus2, "UsermodeCodeIntegrityPolicyEnforcementStatus : 0", "User Mode Code Integrity " & "(Not Configured)")
	Local $topstatus4 = StringReplace($topstatus3, "UsermodeCodeIntegrityPolicyEnforcementStatus : 1", "User Mode Code Integrity " & "(Audit)")
	Local $topstatus5 = StringReplace($topstatus4, "UsermodeCodeIntegrityPolicyEnforcementStatus : 2", "User Mode Code Integrity " & "(Enforced)")
	Local $topstatus6 = StringReplace($topstatus5, "CodeIntegrityPolicyEnforcementStatus         : 0", "Kernel Mode Code Integrity " & "(Not Configured)")
	Local $topstatus7 = StringReplace($topstatus6, "CodeIntegrityPolicyEnforcementStatus         : 1", "Kernel Mode Code Integrity " & "(Audit)")
	Local $topstatus8 = StringReplace($topstatus7, "CodeIntegrityPolicyEnforcementStatus         : 2", "Kernel Mode Code Integrity " & "(Enforced)")
	Global $topstatus9 = StringStripWS($topstatus8, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)

	Local $aPolicyStatus = StringSplit($topstatus9, @CR)
	Global $CurrentPolicyStatus = $sVDriverStatus & "   |   " & $sSACStatus & "   |   " & $aPolicyStatus[1] & "   |   " & $aPolicyStatus[2]

	; update status bar parts with current info
	;
	; update smart app control status
	$g_aText[0] = $sSACStatus
	; update vulnerable driver blocklist status
	$g_aText[1] = $sVDriverStatus
	; update user mode code integrity status
	$g_aText[2] = $aPolicyStatus[2]
	; update kernel mode code integrity status
	$g_aText[3] = $aPolicyStatus[1]

	; redraw status bar to update values
	_WinAPI_RedrawWindow($g_hStatus)
	
Endfunc


Func GetPolicyStatus_old()
	Local $o_CmdString1 = " Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity*"
	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)

	Local $topstatus1 = StringReplace($out, "[32;1m", "")
	Local $topstatus2 = StringReplace($topstatus1, "[0m", "")
	Local $topstatus3 = StringReplace($topstatus2, "UsermodeCodeIntegrityPolicyEnforcementStatus : 0", " App Control user mode policy : " & "Not Configured ")
	Local $topstatus4 = StringReplace($topstatus3, "UsermodeCodeIntegrityPolicyEnforcementStatus : 1", " App Control user mode policy : " & "Audit Mode     ")
	Local $topstatus5 = StringReplace($topstatus4, "UsermodeCodeIntegrityPolicyEnforcementStatus : 2", " App Control user mode policy : " & "Enforced Mode  ")
	Local $topstatus6 = StringReplace($topstatus5, "CodeIntegrityPolicyEnforcementStatus         : 0", "App Control policy           : " & "Not Configured ")
	Local $topstatus7 = StringReplace($topstatus6, "CodeIntegrityPolicyEnforcementStatus         : 1", "App Control policy           : " & "Audit Mode     ")
	Local $topstatus8 = StringReplace($topstatus7, "CodeIntegrityPolicyEnforcementStatus         : 2", "App Control policy           : " & "Enforced Mode  ")
	Global $topstatus9 = StringStripWS($topstatus8, $STR_STRIPLEADING + $STR_STRIPTRAILING)

	;MsgBox($MB_SYSTEMMODAL, "Title", $topstatus9)
Endfunc


_GDIPlus_Startup()

;Global $iW = @DesktopWidth - 420, $iH = $hGUIHeight + 200
Global $iW = @DesktopWidth / 1.5, $iH = @DesktopHeight / 2
$hGUI = GUICreate("App Control Policy Manager", $iW, $iH, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX + $WS_MAXIMIZEBOX)

$aWinSize = WinGetClientSize($hGUI)
GUISetIcon(@ScriptFullPath, 201)

$SBARPartSize = $iW / 4

;GUISetBkColor($g_iBkColor)

;-----------------
; Create a sizebox window (Scrollbar class) BEFORE creating the StatusBar control
$g_hSizebox = _WinAPI_CreateWindowEx(0, "Scrollbar", "", $WS_CHILD + $WS_VISIBLE + $SBS_SIZEBOX, _
0, 0, 0, 0, $hGUI) ; $SBS_SIZEBOX or $SBS_SIZEGRIP

; Subclass the sizebox (by changing the window procedure associated with the Scrollbar class)
Local $hProc = DllCallbackRegister('ScrollbarProc', 'lresult', 'hwnd;uint;wparam;lparam')
$g_hOldProc = _WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, DllCallbackGetPtr($hProc))

Local $hCursor = _WinAPI_LoadCursor(0, $OCR_SIZENWSE)
_WinAPI_SetClassLongEx($g_hSizebox, -12, $hCursor) ; $GCL_HCURSOR = -12

;$g_hBrush = _WinAPI_CreateSolidBrush($g_iBkColor)

;-----------------
$g_hStatus = _GUICtrlStatusBar_Create($hGUI, -1, "", $WS_CLIPSIBLINGS) ; ClipSiblings style +++
Local $aParts[4] = [$SBARPartSize - 50, ($SBARPartSize * 2) - 50, ($SBARPartSize * 3) - 20, -1]
If $aParts[Ubound($aParts) - 1] = -1 Then $aParts[Ubound($aParts) - 1] = $iW ; client width size
_MyGUICtrlStatusBar_SetParts($g_hStatus, $aParts)

Dim $g_aText[Ubound($aParts)] = ["  Smart App Control ()", "Vulnerable Driver Blocklist ()", "User Mode Code Integrity ()", "Kernel Mode Code Integrity ()"]
Dim $g_aRatioW[Ubound($aParts)]
For $i = 0 To UBound($g_aText) - 1
_GUICtrlStatusBar_SetText($g_hStatus, "", $i, $SBT_OWNERDRAW + $SBT_NOBORDERS)
; _GUICtrlStatusBar_SetText($g_hStatus, "", $i, $SBT_OWNERDRAW + $SBT_NOBORDERS) ; interesting ?
$g_aRatioW[$i] = $aParts[$i] / $iW
Next


; get status bar height for GUI and listview height
$StatusBarCtrlID = _WinAPI_GetDlgCtrlID($g_hStatus)
$aPos = ControlGetPos($hGUI, "", $StatusBarCtrlID)
$StatusBarCtrlIDV = $aPos[1]
$StatusBarCtrlIDHeight = $aPos[3]


$aGUI_Pos = WinGetPos($hGUI)
$aGUI_ClientSize = WinGetClientSize($hGUI)
$iCaptionHeight = $aGUI_Pos[3] - $aGUI_ClientSize[1]


If $isDarkMode = True Then
	_SetMenuBkColor(0x2c2c2c)
	_SetMenuTextColor(0xffffff)
	_SetMenuSelectTextColor(0xffffff)
	_SetMenuSelectBkColor(0x404040)
	_SetMenuSelectRectColor(0x404040)
	_SetMenuIconBkColor(0x2c2c2c)
	_SetMenuIconBkGrdColor(0x2c2c2c)
Else
	_SetMenuBkColor($idLightBk)
	_SetMenuTextColor(0x000000)
	_SetMenuSelectTextColor(0x000000)
	_SetMenuSelectBkColor(0xf5cba7)
	_SetMenuSelectRectColor(0xf5cba7)
	_SetMenuIconBkColor($idLightBk)
	_SetMenuIconBkGrdColor($idLightBk)
EndIf


Local $iFileMenu5 = _GUICtrlCreateODTopMenu( "&File", $hGui )
Local $iFileMenu4 = _GUICtrlCreateODTopMenu( "&Logs", $hGui )
Local $iFileMenu3 = _GUICtrlCreateODTopMenu( "&Tools", $hGui )
Local $iFileMenu = _GUICtrlCreateODTopMenu( "&Policy Actions", $hGui )
Local $iFileMenu2 = _GUICtrlCreateODTopMenu( "&Filtering Options", $hGui )


If $isDarkMode = True Then
	$menuPolicyAddUpdate = GUICtrlCreateMenuItem( "Add or Update Policies", $iFileMenu)
	$menuPolicyRemove = GUICtrlCreateMenuItem( "Remove Checked Policies", $iFileMenu)
	$menuPolicyConvert = GUICtrlCreateMenuItem( "Convert XML to Binary", $iFileMenu)
	$menuPolicyRefresh = GUICtrlCreateMenuItem( "Refresh Policy List", $iFileMenu)

	$menuFilterEnforced = GUICtrlCreateMenuItem( "Enforced Policies", $iFileMenu2)
	$menuFilterBase = GUICtrlCreateMenuItem( "Base Policies", $iFileMenu2)
	$menuFilterSupp = GUICtrlCreateMenuItem( "Supplemental Policies", $iFileMenu2)
	$menuFilterSigned = GUICtrlCreateMenuItem( "Signed Policies", $iFileMenu2)
	; disable filtering by Signed policies on pre-24H2 systems
	If $is24H2 = False Then GUICtrlSetState($menuFilterSigned, $GUI_DISABLE)
	$menuFilterSystem = GUICtrlCreateMenuItem( "System Policies", $iFileMenu2)
	$menuEFIPartition = GUICtrlCreateMenuItem( "EFI Partition", $iFileMenu2)
	$menuFilterReset = GUICtrlCreateMenuItem( "Reset", $iFileMenu2)

	$menuAppWizard = GUICtrlCreateMenuItem( "App Control Wizard", $iFileMenu3)
	$menuMsInfo32 = GUICtrlCreateMenuItem( "System Information", $iFileMenu3)

	$menuLogsCI = GUICtrlCreateMenuItem( "Code Integrity", $iFileMenu4)
	$menuLogsScript = GUICtrlCreateMenuItem( "MSI and Script", $iFileMenu4)

	$menuExportCSV = GUICtrlCreateMenuItem( "Export as CSV", $iFileMenu5)
	$menuFileExit = GUICtrlCreateMenuItem( "Exit", $iFileMenu5)
Else
	$menuPolicyAddUpdate = _GUICtrlCreateODMenuItem( "Add or Update Policies", $iFileMenu)
	$menuPolicyRemove = _GUICtrlCreateODMenuItem( "Remove Checked Policies", $iFileMenu)
	$menuPolicyConvert = _GUICtrlCreateODMenuItem( "Convert XML to Binary", $iFileMenu)
	$menuPolicyRefresh = _GUICtrlCreateODMenuItem( "Refresh Policy List", $iFileMenu)

	$menuFilterEnforced = _GUICtrlCreateODMenuItem( "Enforced Policies", $iFileMenu2)
	$menuFilterBase = _GUICtrlCreateODMenuItem( "Base Policies", $iFileMenu2)
	$menuFilterSupp = _GUICtrlCreateODMenuItem( "Supplemental Policies", $iFileMenu2)
	$menuFilterSigned = _GUICtrlCreateODMenuItem( "Signed Policies", $iFileMenu2)
	; disable filtering by Signed policies on pre-24H2 systems
	If $is24H2 = False Then GUICtrlSetState($menuFilterSigned, $GUI_DISABLE)
	$menuFilterSystem = _GUICtrlCreateODMenuItem( "System Policies", $iFileMenu2)
	$menuEFIPartition = _GUICtrlCreateODMenuItem( "EFI Partition", $iFileMenu2)
	$menuFilterReset = _GUICtrlCreateODMenuItem( "Reset", $iFileMenu2)

	$menuAppWizard = _GUICtrlCreateODMenuItem( "App Control Wizard", $iFileMenu3)
	$menuMsInfo32 = _GUICtrlCreateODMenuItem( "System Information", $iFileMenu3)

	$menuLogsCI = _GUICtrlCreateODMenuItem( "Code Integrity", $iFileMenu4)
	$menuLogsScript = _GUICtrlCreateODMenuItem( "MSI and Script", $iFileMenu4)

	$menuExportCSV = _GUICtrlCreateODMenuItem( "Export as CSV", $iFileMenu5)
	$menuFileExit = _GUICtrlCreateODMenuItem( "Exit", $iFileMenu5)
EndIf

$aGUI_ClientSizeLV = WinGetClientSize($hGUI)


Local Const $sCascadiaPath = @WindowsDir & "\fonts\CascadiaCode.ttf"
Local $iCascadiaExists = FileExists($sCascadiaPath)

If $iCascadiaExists Then
	GUISetFont(10, $FW_NORMAL, -1, "Cascadia Mono")
Else
	GUISetFont(10, $FW_NORMAL, -1, "Consolas")
EndIf


Local $hToolTip2 = _GUIToolTip_Create(0)
;_GUIToolTip_SetMaxTipWidth($hToolTip2, 400)

If $isDarkMode = True Then
	_WinAPI_SetWindowTheme($hToolTip2, "", "")
	_GUIToolTip_SetTipBkColor($hToolTip2, 0x202020)
	_GUIToolTip_SetTipTextColor($hToolTip2, 0xe0e0e0)
Else
	_WinAPI_SetWindowTheme($hToolTip2, "", "")
	_GUIToolTip_SetTipBkColor($hToolTip2, 0xffffff)
	_GUIToolTip_SetTipTextColor($hToolTip2, 0x000000)
EndIf

_GUIToolTip_SetMaxTipWidth($hToolTip2, 260)


Local $exStyles = BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_CHECKBOXES, $LVS_EX_DOUBLEBUFFER), $cListView
$cListView = GUICtrlCreateListView("|Enforced|Policy Name|Policy ID|Base Policy ID|Version|On Disk|Signed Policy|System Policy|Authorized|Policy Options", 0, 0, $aWinSize[0], $aGUI_ClientSizeLV[1] - $StatusBarCtrlIDHeight)
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
$hListView = GUICtrlGetHandle($cListView)

_GUICtrlListView_SetExtendedListViewStyle($hListView, $exStyles)

_GUICtrlListView_AddArray($hListView,$aWords)


; header text fix for dark mode

HeaderFix()
Func HeaderFix()
	If $isDarkMode = True Then
		;get handle to child SysHeader32 control of ListView
		Global $hHeader = HWnd(GUICtrlSendMsg($cListView, $LVM_GETHEADER, 0, 0))
		;Turn off theme for header
		DllCall("uxtheme.dll", "int", "SetWindowTheme", "hwnd", $hHeader, "wstr", "", "wstr", "")
		;subclass ListView to get at NM_CUSTOMDRAW notification sent to ListView
		Global $wProcNew = DllCallbackRegister("_LVWndProc", "ptr", "hwnd;uint;wparam;lparam")
		Global $wProcOld = _WinAPI_SetWindowLong($hListView, $GWL_WNDPROC, DllCallbackGetPtr($wProcNew))

		;Optional: Flat Header - remove header 3D button effect
		Global $iStyle = _WinAPI_GetWindowLong($hHeader, $GWL_STYLE)
		_WinAPI_SetWindowLong($hHeader, $GWL_STYLE, BitOR($iStyle, $HDS_FLAT))
	EndIf
Endfunc

CountTotal()
Func CountTotal()
	Global $CountTotal = 0
	Global $CountTotal = _GUICtrlListView_GetItemCount($cListView)
	;ConsoleWrite("Count = " & $CountEnforced & @CRLF)
	;MsgBox($MB_SYSTEMMODAL, "Title", "Count = " & $CountEnforced)
Endfunc

CountEnforced()
Func CountEnforced()
	Global $CountEnforced = 0
	For $I = 0 To _GUICtrlListView_GetItemCount($cListView) - 1
		$Array = _GUICtrlListView_GetItemTextArray($cListView, $I)
		If $Array[2] = "True" Then $CountEnforced += 1
	Next
	;ConsoleWrite("Count = " & $CountEnforced & @CRLF)
	;MsgBox($MB_SYSTEMMODAL, "Title", "Count = " & $CountEnforced)
Endfunc


;UpdatePolicyStatus()
Func UpdatePolicyStatus()
	Global $PolicyStatusInfo = " " & $topstatus9 & @CRLF & @CRLF & $sVulnDrivermsg & @CRLF & @CRLF & " Policies (total)      : " & $CountTotal & @CRLF & " Policies (enforced)   : " & $CountEnforced
EndFunc

$PolicyStatus = GUICtrlCreateLabel($PolicyStatusInfo, 100, 100, -1, -1, $WS_BORDER + $SS_LEFTNOWORDWRAP)
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT + $GUI_DOCKLEFT)
GUICtrlSetState($PolicyStatus, $GUI_HIDE)


; Read ListView content into an array
$aContent = _GUIListViewEx_ReadToArray($cListView)


If $isDarkMode = True Then
	;$hImageList = _GUICtrlListView_GetImageList($hListView, 2)
EndIf


ApplyThemeColor()
Func ApplyThemeColor()
	If $isDarkMode = True Then
		GuiDarkmodeApply($hGUI)
	Else
		Local $bEnableDarkTheme = False
		GuiLightmodeApply($hGUI)
	EndIf
Endfunc


If $isDarkMode = True Then
	$hImageList = _GUICtrlListView_GetImageList($hListView, 2)
	_GUIImageList_Remove($hImageList)

	If @Compiled = 0 Then
		_GUIImageList_AddIcon($hImageList, "unchecked.ico")
		_GUIImageList_AddIcon($hImageList, "checked.ico")
	Else
		_GUIImageList_AddIcon($hImageList, @ScriptFullPath, 3)
		_GUIImageList_AddIcon($hImageList, @ScriptFullPath, 4)
	EndIf
EndIf


; Initiate ListView
$iLV_Index = _GUIListViewEx_Init($cListView, $aContent, 0, Default, False, 1)

; Register required messages
_GUIListViewEx_MsgRegister()

hListViewRefreshColWidth()
Func hListViewRefreshColWidth()
	_GUICtrlListView_SetColumnWidth($hListView, 0, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hListView, 1, $LVSCW_AUTOSIZE_USEHEADER)

	If $policycount = 0 Then
	_GUICtrlListView_SetColumnWidth($hListView, 2, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hListView, 3, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hListView, 4, $LVSCW_AUTOSIZE_USEHEADER)
	Else
	_GUICtrlListView_SetColumnWidth($hListView, 2, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hListView, 3, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hListView, 4, $LVSCW_AUTOSIZE)
	EndIf
	;_GUICtrlListView_SetColumnWidth($hListView, 5, $LVSCW_AUTOSIZE)

	_GUICtrlListView_SetColumnWidth($hListView, 5, $LVSCW_AUTOSIZE)
	Local $aTmp1 = _GUICtrlListView_GetColumn($hListView, 5)
	_GUICtrlListView_SetColumnWidth($hListView, 5, $LVSCW_AUTOSIZE_USEHEADER)
	Local $aTmp2 = _GUICtrlListView_GetColumn($hListView, 5)

	If $aTmp1[4] < $aTmp2[4] Then _GUICtrlListView_SetColumnWidth($hListView, 5, $LVSCW_AUTOSIZE_USEHEADER)
	If $aTmp1[4] > $aTmp2[4] Then _GUICtrlListView_SetColumnWidth($hListView, 5, $LVSCW_AUTOSIZE)

	_GUICtrlListView_SetColumnWidth($hListView, 6, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hListView, 7, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hListView, 8, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hListView, 9, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hListView, 10, $LVSCW_AUTOSIZE_USEHEADER)

	GUICtrlSendMsg( $cListView, $WM_CHANGEUISTATE, 65537, 0 )
EndFunc



; to allow the setting of StatusBar BkColor at least under Windows 10
_WinAPI_SetWindowTheme($g_hStatus, "", "")

; Set status bar background color
_GUICtrlStatusBar_SetBkColor($g_hStatus, $g_iBkColor)

;$g_iHeight = _GUICtrlStatusBar_GetHeight($g_hStatus) + 3 ; change the constant (+3) if necessary
$g_iHeight = $StatusBarCtrlIDHeight
$g_hDots = CreateDots($g_iHeight, $g_iHeight, 0xFF000000 + $g_iBkColor, 0xFF000000 + $g_iTextColor)

GUIRegisterMsg($WM_SIZE, "WM_SIZE")
GUIRegisterMsg($WM_MOVE, "WM_MOVE")
GUIRegisterMsg($WM_DRAWITEM, "WM_DRAWITEM")

GetPolicyStatus()


GUISetState()


Example()

Func Example()
        Local $aWindow_Size = WinGetPos($hGUI)
		ConsoleWrite('X position    = ' & $aWindow_Size[0] & @CRLF)
		ConsoleWrite('Y position    = ' & $aWindow_Size[1] & @CRLF)
        ConsoleWrite('Window Width  = ' & $aWindow_Size[2] & @CRLF)
        ConsoleWrite('Window Height = ' & $aWindow_Size[3] & @CRLF)
        Local $aWindowClientArea_Size = WinGetClientSize($hGUI)
        ConsoleWrite('Window Client Area Width  = ' & $aWindowClientArea_Size[0] & @CRLF)
        ConsoleWrite('Window Client Area Height = ' & $aWindowClientArea_Size[1] & @CRLF)
EndFunc   ;==>Example

#cs
WinMove($hGUI,'',(@Desktopwidth - WinGetPos($hGUI)[2]) / 2,(@Desktopheight - WinGetPos($hGUI)[3]) / 2, @DesktopWidth - 420 + 6, $GUIHeightMeasured + 6)

;Sleep(500)

;GUISetState()

GUISetStyle($GUI_SS_DEFAULT_GUI, -1)

WinMove($hGUI,'', (@Desktopwidth - WinGetPos($hGUI)[2]) / 2,(@Desktopheight - WinGetPos($hGUI)[3]) / 2)

GUISetState(@SW_SHOWMINIMIZED)
GUISetState(@SW_RESTORE)

WinSetOnTop($hGUI, "", $WINDOWS_ONTOP)
WinSetOnTop($hGUI, "", $WINDOWS_NOONTOP)
#ce
WDACWizardExists()
;GUISetState(@SW_SHOW)



Func _LVWndProc($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg, $wParam
    If $iMsg = $WM_NOTIFY Then
        Local $tNMHDR, $hWndFrom, $iCode, $iItem, $hDC
        $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
        $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
        $iCode = DllStructGetData($tNMHDR, "Code")
        ;Local $IDFrom = DllStructGetData($tNMHDR, "IDFrom")

        Switch $hWndFrom
            Case $hHeader
                Switch $iCode
                    Case $NM_CUSTOMDRAW
                        Local $tCustDraw = DllStructCreate($tagNMLVCUSTOMDRAW, $lParam)
                        Switch DllStructGetData($tCustDraw, "dwDrawStage")
                            Case $CDDS_PREPAINT
                                Return $CDRF_NOTIFYITEMDRAW
                            Case $CDDS_ITEMPREPAINT
                                $hDC = DllStructGetData($tCustDraw, "hDC")
                                $iItem = DllStructGetData($tCustDraw, "dwItemSpec")
                                DllCall($iDllGDI, "int", "SetTextColor", "handle", $hDC, "dword", $aCol[$iItem][0])
                                DllCall($iDllGDI, "int", "SetBkColor", "handle", $hDC, "dword", $aCol[$iItem][1])
                                Return $CDRF_NEWFONT
                                Return $CDRF_SKIPDEFAULT
                        EndSwitch
                EndSwitch
        EndSwitch
    EndIf
    ;pass the unhandled messages to default WindowProc
    Local $aResult = DllCall($iDllUSER32, "lresult", "CallWindowProcW", "ptr", $wProcOld, _
            "hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, "lparam", $lParam)
    If @error Then Return -1
    Return $aResult[0]
EndFunc   ;==>_LVWndProc

Func _BGR2RGB($iColor)
    ;Author: Wraithdu
    Return BitOR(BitShift(BitAND($iColor, 0x0000FF), -16), BitAND($iColor, 0x00FF00), BitShift(BitAND($iColor, 0xFF0000), 16))
EndFunc   ;==>_BGR2RGB

Func AddPolicies()
	Local $spFile

	$mFile = FileOpenDialog("Select Policy File(s) to Add or Update", @ScriptDir, "Policy Files (*.cip)", 1 + 4 )
	If @error Then
		ConsoleWrite("error")
	Else
		$spFile = StringSplit($mFile, "|")
		If UBound($spFile) = 2 Then
			$path = $spFile[1]
			_ArrayDelete($spFile, 0)
			Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
			Local $aPathSplit = _PathSplit($spFile[0], $sDrive, $sDir, $sFileName, $sExtension)
			Local $cmd1 = ' (citool.exe -up '
			Local $cmd2 = '"'
			Local $cmd3 = $aPathSplit[3]
			Local $cmd4 = $aPathSplit[4]
			Local $cmd5 = '"'
			Local $cmd6 = ' -json)'
			Run(@ComSpec & " /c " & $cmd1 & $cmd2 & $cmd3 & $cmd4 & $cmd5 & $cmd6, "", @SW_HIDE)
			$sMsg = "Selected policy has been added." & @CRLF
			$sMsg &= "" & @CRLF
			$sMsg &= "The policy list and status information will automatically" & @CRLF
			$sMsg &= "refresh in about 5 seconds."
			;_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
			_LibNotif_Notif($sMsg)
			Sleep(5000)
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicyInfo()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			CountTotal()
			CountEnforced()
			UpdatePolicyStatus()
			hListViewRefreshColWidth()
			GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
		Else
			$path = $spFile[1]
			_ArrayDelete($spFile, 0)
			_ArrayDelete($spFile, 0)
			For $x = 0 to UBound($spFile)-1
				Local $cmd1 = ' (citool.exe -up '
				Local $cmd2 = '"'
				Local $cmd3 = $spFile[$x]
				Local $cmd4 = '"'
				Local $cmd5 = ' -json)'
				Run(@ComSpec & " /c " & $cmd1 & $cmd2 & $cmd3 & $cmd4 & $cmd5, "", @SW_HIDE)
			Next
				$sMsg = "Selected policies have been added." & @CRLF
				$sMsg &= "" & @CRLF
				$sMsg &= "The policy list and status information will automatically" & @CRLF
				$sMsg &= "refresh in about 5 seconds."
				;_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
				_LibNotif_Notif($sMsg)
				Sleep(5000)
				_GUICtrlListView_DeleteAllItems($cListView)
				GetPolicyInfo()
				GetPolicyStatus()
				_GUICtrlListView_AddArray($hListView,$aWords)
				CountTotal()
				CountEnforced()
				UpdatePolicyStatus()
				hListViewRefreshColWidth()
				GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
		EndIf
	EndIf
EndFunc

Func LogsCI()
	$oXML=ObjCreate("Microsoft.XMLDOM")
	$stest = FileRead(@AppDataDir & "\Microsoft\MMC\eventvwr")
	$oXML.LoadXML($stest)

	If Not _XML_NodeExists($oXML, "//DynamicPath") Then
			Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-CodeIntegrity/Operational"', "", @SW_HIDE)
	Else
	$oXMLChild1 = $oXML.selectSingleNode( "//DynamicPath" )
	$parent = $oXMLChild1.ParentNode

	$test = $parent.removeChild ( $oXMLChild1 )

	$oXML.Save(@AppDataDir & "\Microsoft\MMC\eventvwr")
	Sleep(500)
	Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-CodeIntegrity/Operational"', "", @SW_HIDE)
	EndIf
Endfunc

Func LogsScript()
	$oXML=ObjCreate("Microsoft.XMLDOM")
	$stest = FileRead(@AppDataDir & "\Microsoft\MMC\eventvwr")
	$oXML.LoadXML($stest)

	If Not _XML_NodeExists($oXML, "//DynamicPath") Then
			Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-AppLocker/MSI and Script"', "", @SW_HIDE)
	Else
	$oXMLChild1 = $oXML.selectSingleNode( "//DynamicPath" )
	$parent = $oXMLChild1.ParentNode

	$test = $parent.removeChild ( $oXMLChild1 )

	$oXML.Save(@AppDataDir & "\Microsoft\MMC\eventvwr")
	Sleep(500)
	Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-AppLocker/MSI and Script"', "", @SW_HIDE)
	EndIf
Endfunc

Func About()
	_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", " Version: " & $programversion & @CRLF & _
                            " Created by: " & "WildByDesign" & @CRLF & @CRLF & _
                            " This program is free and open source. " & @CRLF)
Endfunc

Func CountChecked()
    Local $sReturn = ''
    For $i = 0 To _GUICtrlListView_GetItemCount($cListView) - 1
        If _GUICtrlListView_GetItemChecked($cListView, $i) Then
            $sReturn &= $i & '|'
        EndIf
	Next
	If $sReturn = '' Then
		;_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", " No policies have been selected for removal. " & @CRLF)
		_LibNotif_Notif("No policies have been selected for removal.")
	Else
		CheckIsSystem()
	EndIf
Endfunc

Func CheckIsSystem()
    Local $sReturn = ''
    For $i = 0 To _GUICtrlListView_GetItemCount($cListView) - 1
        ; IsSystemPolicy is column 8
		$systempol = _GUICtrlListView_GetItemText($cListView, $i, 8)
		If _GUICtrlListView_GetItemChecked($cListView, $i) Then
            $sReturn &= $systempol & '|'
        EndIf
	Next
	Local $SystemPolicyTrue = StringInStr($sReturn, "True")
	If $SystemPolicyTrue <> 0 Then
		$sMsg = "System policies cannot be removed with this method and" & @CRLF
		$sMsg &= "are protected in the EFI partition." & @CRLF & @CRLF
		$sMsg &= "Please uncheck any system policies and try again."
		;_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
		_LibNotif_Notif($sMsg)
	Else
		;PolicyRemoval()
		CheckIsSigned()
	EndIf
Endfunc

Func CheckIsSigned()
    Local $sReturn = ''
    For $i = 0 To _GUICtrlListView_GetItemCount($cListView) - 1
        ; IsSignedPolicy is column 7
		$signedpol = _GUICtrlListView_GetItemText($cListView, $i, 7)
		If _GUICtrlListView_GetItemChecked($cListView, $i) Then
            $sReturn &= $signedpol & '|'
        EndIf
	Next
	Local $SignedPolicyTrue = StringInStr($sReturn, "True")
	If $SignedPolicyTrue <> 0 Then
		$sMsg = "Signed base policies cannot be removed with this" & @CRLF
		$sMsg &= "method but Signed supplemental policies can. (not implemented yet)"
		;_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
		_LibNotif_Notif($sMsg)
		PolicyRemoval()
	Else
		PolicyRemoval()
	EndIf
Endfunc

Func PolicyRemoval()
	;For $i = 1 To $policycount
	For $i = 1 To _GUICtrlListView_GetItemCount($cListView)
	;For $i = 1 To $policycount
		Local $CheckedStatus = _GUICtrlListView_GetItemChecked($cListView, $i - 1)
		If _GUICtrlListView_GetItemChecked($cListView, $i - 1) Then
			Local $cmd1 = ' (citool.exe -rp '
			Local $cmd2 = '"{'
			Local $cmd3 = _GUICtrlListView_GetItemText($cListView, $i - 1, 3)
			Local $cmd4 = '}"'
			Local $cmd5 = ' -json)'
			Run(@ComSpec & " /c " & $cmd1 & $cmd2 & $cmd3 & $cmd4 & $cmd5, "", @SW_HIDE)
			;MsgBox($MB_SYSTEMMODAL, "Title", $cmd1 & $cmd2 & $cmd3 & $cmd4 & $cmd5)
		EndIf
		Next
		$sMsg = "Selected policies have been removed." & @CRLF
		$sMsg &= "" & @CRLF
		$sMsg &= "The policy list and status information will automatically" & @CRLF
		$sMsg &= "refresh in about 5 seconds."
		;_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
		_LibNotif_Notif($sMsg)
		Sleep(5000)
		_GUICtrlListView_DeleteAllItems($cListView)
		GetPolicyInfo()
		GetPolicyStatus()
		_GUICtrlListView_AddArray($hListView,$aWords)
		CountTotal()
		CountEnforced()
		UpdatePolicyStatus()
		hListViewRefreshColWidth()
		GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
Endfunc

Func ConvertPolicy()
	$mFile = FileOpenDialog("Select XML Policy File for Conversion", @ScriptDir, "Policy Files (*.xml)", 1)
	If @error Then
	;	Exit
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

;	EndIf

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
Endfunc


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


Func PolicyWizard()
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
Endfunc


Func CheckWizardInstall()
	; Check for Wizard install every 10 seconds
	AdlibRegister("WDACWizardExists", 10000)
EndFunc


Func UnmountEFI()
	Local $o_CmdString1 = " $MountPoint = $env:SystemDrive+'\EFIMount'; mountvol $MountPoint /D"
	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide)
	ProcessWaitCloseEx($o_Pid)

	Local $systemdrive = EnvGet('SystemDrive')
	Local Const $sFilePath = $systemdrive & '\EFIMount\EFI\Microsoft'
	Local Const $sFilePath2 = $systemdrive & '\EFIMount'
	Local $iFileExists = FileExists($sFilePath)
	If $iFileExists Then
	Else
		DirRemove($sFilePath2)
	EndIf
Endfunc


Func Msinfo32()
    Run(@WindowsDir & "\System32\msinfo32.exe", "", @SW_SHOWMAXIMIZED)
EndFunc


Func ProcessWaitCloseEx($iPID)
	While ProcessExists($iPID) And Sleep(10)
	WEnd
EndFunc


Func EFIPartition()
	; Refresh policy list
	GetPolicyInfo()

	; mount EFI

	Local $o_CmdString1 = " $MountPoint = $env:SystemDrive+'\EFIMount'; $EFIPartition = (Get-Partition | Where-Object IsSystem).AccessPaths[0]; if (-Not (Test-Path $MountPoint)) { New-Item -Path $MountPoint -Type Directory -Force }; mountvol $MountPoint $EFIPartition"
	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide)
	ProcessWaitCloseEx($o_Pid)
	Local $systemdrive = EnvGet('SystemDrive')
	;MsgBox(0, '$systemdrive', $systemdrive)
	Local Const $sFilePath = $systemdrive & '\EFIMount\EFI\Microsoft'
	Local $iFileExists = FileExists($sFilePath)
	If $iFileExists Then
		;GUICtrlSetState($RemoveEFI,$GUI_ENABLE)
		;ListEFI()
	EndIf


	; list EFI multi-policy

	Local $o_CmdString1 = " $MountPoint = $env:SystemDrive+'\EFIMount'; $MultiPolicyDir = $MountPoint+'\EFI\Microsoft\Boot\CiPolicies\Active'; $CIPFiles = Get-ChildItem $MultiPolicyDir\*.cip -Name; $CIPFiles"

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	;MsgBox(0, 'test', $out)
	Local $CIPFiles1 = StringReplace($out, "[32;1m", "")
	Local $CIPFiles2 = StringReplace($CIPFiles1, "[0m", "")
	Local $CIPFiles3 = StringReplace($CIPFiles2, ".cip", "")
	Local $CIPFiles4 = StringStripWS($CIPFiles3, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
	Local $finalparse = $CIPFiles4
	;MsgBox(0, 'test', $CIPFiles4)
	$EFIsplit = stringsplit($finalparse, @CR , 0)
	;_ArrayDisplay($EFIsplit, "test")
	Global $EFIarray[0][11]
	;_ArrayDisplay($CIParray, "test")

	For $i = 0 To UBound($aWords) -1
		; column 3 is PolicyID
		$iEFI = $aWords[$i][3]
		;MsgBox(0, 'test', $iCIP)
		Local $iPosition = StringInStr($finalparse , $iEFI)
		;Local $iPosition2 = StringInStr($finalparse , "{82443e1e-8a39-4b4a-96a8-f40ddc00b9f3}")
		;Local $iPosition3 = StringInStr($finalparse , "{CDD5CB55-DB68-4D71-AA38-3DF2B6473A52}")
		;MsgBox(0, 'test', $iPosition)
		Global $aExtract
		If $iPosition <> '0' Then
			;_ArrayAdd($CIParray, $aWords[1])
			Local $aExtract = _ArrayExtract($aWords, $i, $i)
			_ArrayAdd($EFIarray, $aExtract)
			;Local $CIParray[1][11] = $aExtract
			;$CIParray[1] = $aWords[$i]
		EndIf

	Next


	; list EFI single policy

	Local $o_CmdString1 = " $MountPoint = $env:SystemDrive+'\EFIMount'; $SinglePolicyDir = $MountPoint+'\EFI\Microsoft\Boot'; $p7bFiles = Get-ChildItem $SinglePolicyDir\*.p7b -Name; $p7bFiles"

	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
	ProcessWaitCloseEx($o_Pid)
	$out = StdoutRead($o_Pid)
	;MsgBox(0, 'test', $out)
	Local $p7bFiles1 = StringReplace($out, "[32;1m", "")
	Local $p7bFiles2 = StringReplace($p7bFiles1, "[0m", "")
	Local $p7bFiles3 = StringReplace($p7bFiles2, "driversipolicy.p7b", "{d2bda982-ccf6-4344-ac5b-0b44427b6816}")
	Local $p7bFiles4 = StringReplace($p7bFiles3, "winsipolicy.p7b", "{5951a96a-e0b5-4d3d-8fb8-3e5b61030784}")
	Local $p7bFiles5 = StringStripWS($p7bFiles4, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
	Local $finalparse = $p7bFiles5
	;MsgBox(0, 'test', $CIPFiles4)
	$EFIp7bsplit = stringsplit($finalparse, @CR , 0)
	;_ArrayDisplay($EFIsplit, "test")
	;Global $EFIarray[0][11]
	;_ArrayDisplay($CIParray, "test")

	For $i = 0 To UBound($aWords) -1
		; column 3 is PolicyID
		$iEFI = $aWords[$i][3]
		;MsgBox(0, 'test', $iCIP)
		Local $iPosition = StringInStr($finalparse , $iEFI)
		;Local $iPosition2 = StringInStr($finalparse , "{82443e1e-8a39-4b4a-96a8-f40ddc00b9f3}")
		;Local $iPosition3 = StringInStr($finalparse , "{CDD5CB55-DB68-4D71-AA38-3DF2B6473A52}")
		;MsgBox(0, 'test', $iPosition)
		Global $aExtract
		If $iPosition <> '0' Then
			;_ArrayAdd($CIParray, $aWords[1])
			Local $aExtract = _ArrayExtract($aWords, $i, $i)
			_ArrayAdd($EFIarray, $aExtract)
			;Local $CIParray[1][11] = $aExtract
			;$CIParray[1] = $aWords[$i]
		EndIf

	Next


	; EFI listview

	_GUICtrlListView_AddArray($hListView,$EFIarray)

	; Read ListView content into an array
	$aContent = _GUIListViewEx_ReadToArray($cListView)
EndFunc


;==============================================
Func ScrollbarProc($hWnd, $iMsg, $wParam, $lParam) ; Andreik

    If $hWnd = $g_hSizebox And $iMsg = $WM_PAINT Then
        Local $tPAINTSTRUCT
        Local $hDC = _WinAPI_BeginPaint($hWnd, $tPAINTSTRUCT)
        Local $iWidth = DllStructGetData($tPAINTSTRUCT, 'rPaint', 3) - DllStructGetData($tPAINTSTRUCT, 'rPaint', 1)
        Local $iHeight = DllStructGetData($tPAINTSTRUCT, 'rPaint', 4) - DllStructGetData($tPAINTSTRUCT, 'rPaint', 2)
        Local $hGraphics = _GDIPlus_GraphicsCreateFromHDC($hDC)
        _GDIPlus_GraphicsDrawImageRect($hGraphics, $g_hDots, 0, 0, $iWidth, $iHeight)
        _GDIPlus_GraphicsDispose($hGraphics)
        _WinAPI_EndPaint($hWnd, $tPAINTSTRUCT)
        Return 0
    EndIf
    Return _WinAPI_CallWindowProc($g_hOldProc, $hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>ScrollbarProc

;==============================================
Func CreateDots($iWidth, $iHeight, $iBackgroundColor, $iDotsColor) ; Andreik

    Local $iDotSize = Int($iHeight / 10)
    Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iWidth, $iHeight)
    Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
    Local $hBrush = _GDIPlus_BrushCreateSolid($iDotsColor)
    _GDIPlus_GraphicsClear($hGraphics, $iBackgroundColor)
    Local $a[6][2] = [[2,6], [2,4], [2,2], [4,4], [4,2], [6,2]]
    For $i = 0 To UBound($a) - 1
        _GDIPlus_GraphicsFillRect($hGraphics, $iWidth - $iDotSize * $a[$i][0], $iHeight - $iDotSize * $a[$i][1], $iDotSize, $iDotSize, $hBrush)
    Next
    _GDIPlus_BrushDispose($hBrush)
    _GDIPlus_GraphicsDispose($hGraphics)
    Return $hBitmap
EndFunc   ;==>CreateDots

;==============================================
Func _MyGUICtrlStatusBar_SetParts($hWnd, $aPartEdge) ; Pixelsearch

    If Not IsArray($aPartEdge) Then Return False
    Local $iParts = UBound($aPartEdge)
    Local $tParts = DllStructCreate("int[" & $iParts & "]")
    For $i = 0 To $iParts - 1
        DllStructSetData($tParts, 1, $aPartEdge[$i], $i + 1)
    Next
    DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $SB_SETPARTS, "wparam", $iParts, "struct*", $tParts)
    _GUICtrlStatusBar_Resize($hWnd)
    Return True
EndFunc   ;==>_MyGUICtrlStatusBar_SetParts

;==============================================
Func WM_SIZE($hWnd, $iMsg, $wParam, $lParam) ; Pixelsearch
    #forceref $iMsg, $wParam, $lParam

    If $hWnd = $hGUI Then
        Local Static $bIsSizeBoxShown = True
        Local $aSize = WinGetClientSize($hGUI)
        Local $aGetParts = _GUICtrlStatusBar_GetParts($g_hStatus)
        Local $aParts[$aGetParts[0]]
        For $i = 0 To $aGetParts[0] - 1
            $aParts[$i] = Int($aSize[0] * $g_aRatioW[$i])
        Next
        If BitAND(WinGetState($hGUI), $WIN_STATE_MAXIMIZED) Then
            _GUICtrlStatusBar_SetParts($g_hStatus, $aParts) ; set parts until GUI right border
            _WinAPI_ShowWindow($g_hSizebox, @SW_HIDE)
            $bIsSizeBoxShown = False
        Else
            _MyGUICtrlStatusBar_SetParts($g_hStatus, $aParts) ; set parts as user scripted them
            WinMove($g_hSizebox, "", $aSize[0] - $g_iHeight, $aSize[1] - $g_iHeight, $g_iHeight, $g_iHeight)
            If Not $bIsSizeBoxShown Then
                _WinAPI_ShowWindow($g_hSizebox, @SW_SHOW)
                $bIsSizeBoxShown = True
            EndIf
        EndIf
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_SIZE

;==============================================
Func WM_MOVE($hWnd, $iMsg, $wParam, $lParam)
    #forceref $iMsg, $wParam, $lParam

    If $hWnd = $hGUI Then
        _WinAPI_RedrawWindow($g_hSizebox)
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_MOVE

;==============================================
Func WM_DRAWITEM($hWnd, $iMsg, $wParam, $lParam) ; Kafu
    #forceref $hWnd, $iMsg, $wParam

    Local Static $tagDRAWITEM = "uint CtlType;uint CtlID;uint itemID;uint itemAction;uint itemState;hwnd hwndItem;handle hDC;long rcItem[4];ulong_ptr itemData"
    Local $tDRAWITEM = DllStructCreate($tagDRAWITEM, $lParam)
    If $tDRAWITEM.hwndItem = $g_hStatus Then

		Local $itemID = $tDRAWITEM.itemID ; status bar part number (0, 1, ...)
		Local $hDC = $tDRAWITEM.hDC
		Local $tRect = DllStructCreate("long left;long top;long right;long bottom", DllStructGetPtr($tDRAWITEM, "rcItem"))
		;_WinAPI_FillRect($hDC, DllStructGetPtr($tRect), $g_hBrush) ; backgound color
		_WinAPI_SetTextColor($hDC, $g_iTextColor) ; text color
		_WinAPI_SetBkMode($hDC, $TRANSPARENT)
		DllStructSetData($tRect, "top", $tRect.top + 1)
		DllStructSetData($tRect, "left", $tRect.left + 1)
		_WinAPI_DrawText($hDC, $g_aText[$itemID], $tRect, $DT_LEFT)

		Return $GUI_RUNDEFMSG
	Else
		_WM_DRAWITEM($hWnd, $iMsg, $wParam, $lParam)
	EndIf
EndFunc   ;==>WM_DRAWITEM


Func ExportToCSV()
    Local Const $sMessage = "Choose a filename."
    Local $sFileSaveDialog = FileSaveDialog($sMessage, @DesktopDir, "Comma-separated values (*.csv)", $FD_PATHMUSTEXIST + $FD_PROMPTOVERWRITE, "ExportAppControl")

	If @error Then
		;MsgBox($MB_SYSTEMMODAL, "", "No file was saved.")
	Else
		Local $sFileName = StringTrimLeft($sFileSaveDialog, StringInStr($sFileSaveDialog, "\", $STR_NOCASESENSEBASIC, -1))

		Local $iExtension = StringInStr($sFileName, ".", $STR_NOCASESENSEBASIC)

		If $iExtension Then
			If Not (StringTrimLeft($sFileName, $iExtension - 1) = ".csv") Then $sFileSaveDialog &= ".csv"
		Else
			$sFileSaveDialog &= ".csv"
		EndIf

		_GUICtrlListView_SaveCSV($cListView, $sFileSaveDialog)
	EndIf
EndFunc


Local $aMsg

While 1
    $aMsg = GUIGetMsg(1)
	If Not IsHWnd($aMsg[1]) Then ContinueLoop
		Switch $aMsg[1]
			Case $hGUI
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						ExitLoop
					Case $menuPolicyAddUpdate
						AddPolicies()
					Case $menuPolicyConvert
						ConvertPolicy()
					Case $menuPolicyRefresh
						WinSetTitle($hGUI, "", "App Control Policy Manager")
						GUICtrlSetState($menuPolicyRemove, $GUI_ENABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						GetPolicyInfo()
						GetPolicyStatus()
						_GUICtrlListView_AddArray($hListView,$aWords)
						hListViewRefreshColWidth()
						UpdatePolicyStatus()
						;GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
					Case $menuPolicyRemove
						CountChecked()
					Case $menuAppWizard
						PolicyWizard()
					Case $menuLogsCI
						LogsCI()
					Case $menuLogsScript
						LogsScript()
					Case $menuFilterEnforced
						WinSetTitle($hGUI, "", "App Control Policy Manager - Enforced Policies")
						GUICtrlSetState($menuPolicyRemove, $GUI_ENABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						GetPolicyEnforced()
						GetPolicyStatus()
						_GUICtrlListView_AddArray($hListView,$aWords)
						hListViewRefreshColWidth()
						UpdatePolicyStatus()
						;GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
					Case $menuFilterBase
						WinSetTitle($hGUI, "", "App Control Policy Manager - Base Policies")
						GUICtrlSetState($menuPolicyRemove, $GUI_ENABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						GetPolicyBase()
						GetPolicyStatus()
						_GUICtrlListView_AddArray($hListView,$aWords)
						hListViewRefreshColWidth()
						UpdatePolicyStatus()
						;GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
					Case $menuFilterSupp
						WinSetTitle($hGUI, "", "App Control Policy Manager - Supplemental Policies")
						GUICtrlSetState($menuPolicyRemove, $GUI_ENABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						GetPolicySupp()
						GetPolicyStatus()
						_GUICtrlListView_AddArray($hListView,$aWords)
						hListViewRefreshColWidth()
						UpdatePolicyStatus()
						;GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
					Case $menuFilterSigned
						WinSetTitle($hGUI, "", "App Control Policy Manager - Signed Policies")
						GUICtrlSetState($menuPolicyRemove, $GUI_DISABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						GetPolicySigned()
						GetPolicyStatus()
						_GUICtrlListView_AddArray($hListView,$aWords)
						hListViewRefreshColWidth()
						UpdatePolicyStatus()
						;GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
					Case $menuFilterSystem
						WinSetTitle($hGUI, "", "App Control Policy Manager - System Policies")
						GUICtrlSetState($menuPolicyRemove, $GUI_DISABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						GetPolicySystem()
						GetPolicyStatus()
						_GUICtrlListView_AddArray($hListView,$aWords)
						hListViewRefreshColWidth()
						UpdatePolicyStatus()
						;GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
					Case $menuEFIPartition
						WinSetTitle($hGUI, "", "App Control Policy Manager - EFI Partition")
						GUICtrlSetState($menuPolicyRemove, $GUI_DISABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						EFIPartition()
						hListViewRefreshColWidth()
						UnmountEFI()
					Case $menuFilterReset
						WinSetTitle($hGUI, "", "App Control Policy Manager")
						GUICtrlSetState($menuPolicyRemove, $GUI_ENABLE)
						_GUICtrlListView_DeleteAllItems($cListView)
						GetPolicyInfo()
						GetPolicyStatus()
						_GUICtrlListView_AddArray($hListView,$aWords)
						hListViewRefreshColWidth()
						UpdatePolicyStatus()
						;GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
					Case $menuMsInfo32
						Msinfo32()
					Case $menuExportCSV
						ExportToCSV()
					Case $menuFileExit
						_GDIPlus_BitmapDispose($g_hDots)
						_GUICtrlStatusBar_Destroy($g_hStatus)
						_WinAPI_DestroyCursor($hCursor)
						;_WinAPI_DeleteObject($g_hBrush)
						_WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, $g_hOldProc)
						DllCallbackFree($hProc)
						_GDIPlus_Shutdown()
						Exit
				EndSwitch
				;GUIDelete($hGui)
				;ScrollbarDots_Cleanup()
		Case $hGUI2
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE, $ExitButtonEFI
					UnmountEFI()
					GUIDelete($hGUI2)
					GUISetState(@SW_SHOW, $hGUI)
					If $isDarkMode = True Then
						If $wProcOld Then _WinAPI_SetWindowLong($hListView, $GWL_WNDPROC, $wProcOld)
					; Delete callback function
						If $wProcNew Then DllCallbackFree($wProcNew)
						HeaderFix()
						ApplyThemeColor()
					Else
					ApplyThemeColor()
					EndIf
					;HeaderFix()
					;GUICtrlSetState($g_idButton2, $GUI_ENABLE) ; ... enable button (previously disabled)
				;Case $g_idButton3
					;MsgBox($MB_OK, "MsgBox", "Test from GUI 2")
			EndSwitch
		EndSwitch

;_GUIListViewEx_EventMonitor()

	;If $wProcOld Then _WinAPI_SetWindowLong($hListView, $GWL_WNDPROC, $wProcOld)
	; Delete callback function
	;If $wProcNew Then DllCallbackFree($wProcNew)
WEnd

_GDIPlus_BitmapDispose($g_hDots)
_GUICtrlStatusBar_Destroy($g_hStatus)
_WinAPI_DestroyCursor($hCursor)
;_WinAPI_DeleteObject($g_hBrush)
_WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, $g_hOldProc)
DllCallbackFree($hProc)
_GDIPlus_Shutdown()