#NoTrayIcon
#RequireAdmin

#Region ; *** Dynamically added Include files ***
#include <Array.au3>                                         ; added:01/27/25 07:16:20
#include <AutoItConstants.au3>                               ; added:01/27/25 07:16:20
#include <ButtonConstants.au3>                               ; added:01/27/25 07:16:20
#include <File.au3>                                          ; added:01/27/25 07:16:20
#include <FileConstants.au3>                                 ; added:01/27/25 07:16:20
#include <FontConstants.au3>                                 ; added:01/27/25 07:16:20
#include <GUIConstantsEx.au3>                                ; added:01/27/25 07:16:20
#include <GuiListView.au3>                                   ; added:01/27/25 07:16:20
#include <GuiToolTip.au3>                                    ; added:01/27/25 07:16:20
#include <HeaderConstants.au3>                               ; added:01/27/25 07:16:20
#include <ListViewConstants.au3>                             ; added:01/27/25 07:16:20
#include <StaticConstants.au3>                               ; added:01/27/25 07:16:20
#include <String.au3>                                        ; added:01/27/25 07:16:20
#include <StringConstants.au3>                               ; added:01/27/25 07:16:20
#include <StructureConstants.au3>                            ; added:01/27/25 07:16:20
#include <WinAPISysInternals.au3>                            ; added:01/27/25 07:16:20
#include <WinAPISysWin.au3>                                  ; added:01/27/25 07:16:20
#include <WinAPITheme.au3>                                   ; added:01/27/25 07:16:20
#include <WindowsConstants.au3>                              ; added:01/27/25 07:16:20
#EndRegion ; *** Dynamically added Include files ***

#include "includes\ExtMsgBox.au3"
#include "includes\GUIDarkMode_v0.02mod.au3"
#include "includes\GUIListViewEx.au3"
#include "includes\XML.au3"

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Policy Manager
#AutoIt3Wrapper_res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Fileversion=5.5.0.0
#AutoIt3Wrapper_Res_ProductVersion=5.5.0
#AutoIt3Wrapper_Res_ProductName=AppControlPolicyManager
#AutoIt3Wrapper_Outfile_x64=AppControlManager.exe
#AutoIt3Wrapper_Res_LegalCopyright=@ 2025 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region

Global $programversion = "5.5"

If @Compiled = 0 Then
	; System aware DPI awareness
	;DllCall("User32.dll", "bool", "SetProcessDPIAware")
	; Per-monitor V2 DPI awareness
	DllCall("User32.dll", "bool", "SetProcessDpiAwarenessContext" , "HWND", "DPI_AWARENESS_CONTEXT" -4)
EndIf


;Opt('MustDeclareVars', 1)

$GetDPI = _GetDPI()

; 96 DPI = 100% scaling
; 120 DPI = 125% scaling
; 144 DPI = 150% scaling
; 168 DPI = 175% scaling
; 192 DPI = 200% scaling


If $GetDPI = 96 Then
	$DPIScale = 100
	$ListViewWidth = 1000
	$ListViewHeight = 340
	$hGUIWidth = 1020
	$hGUIHeight = 580
	$PolicyInfoWidth = -1
	$PolicyInfoHeight = -1
	$HorizontalSpace = 25
	$HorizontalSpaceSml = 15
	$VerticalSpaceSml = 8
	$VerticalSpace = 48
	$PolicyInfoLabelWidth = 340
	$PolicyActionLabelWidth = 200
	$FilteringLabelWidth = 180
	$ToolslogsLabelWidth = 154
	$PolicyInfoLabelHeight= 30
	$PolicyActionLabelHeight = 30
	$FilteringLabelHeight = 30
	$ToolslogsLabelHeight = 30
	$ClientAreaTitlebar = 24
ElseIf $GetDPI = 120 Then
	$DPIScale = 125
	$ListViewWidth = 1200
	$ListViewHeight = 340
	$hGUIWidth = 1220
	$hGUIHeight = 634
	$PolicyInfoWidth = -1
	$PolicyInfoHeight = -1
	$HorizontalSpace = 34
	$HorizontalSpaceSml = 10
	$VerticalSpaceSml = 14
	$VerticalSpace = 48
	$PolicyInfoLabelWidth = 340
	$PolicyActionLabelWidth = 200
	$FilteringLabelWidth = 180
	$ToolslogsLabelWidth = 154
	$PolicyInfoLabelHeight= 30
	$PolicyActionLabelHeight = 30
	$FilteringLabelHeight = 30
	$ToolslogsLabelHeight = 30
	$ClientAreaTitlebar = 30
ElseIf $GetDPI = 144 Then
	$DPIScale = 150
	$ListViewWidth = 1400
	$ListViewHeight = 380
	$hGUIWidth = 1420
	$hGUIHeight = 750
	$PolicyInfoWidth = -1
	$PolicyInfoHeight = -1
	$HorizontalSpace = 25
	$HorizontalSpaceSml = 20
	$VerticalSpaceSml = 14
	$VerticalSpace = 68
	$PolicyInfoLabelWidth = 340
	$PolicyActionLabelWidth = 200
	$FilteringLabelWidth = 180
	$ToolslogsLabelWidth = 154
	$PolicyInfoLabelHeight= 30
	$PolicyActionLabelHeight = 30
	$FilteringLabelHeight = 30
	$ToolslogsLabelHeight = 30
	$ClientAreaTitlebar = 34
ElseIf $GetDPI = 168 Then
	$DPIScale = 175
	$ListViewWidth = 1520
	$ListViewHeight = 380
	$hGUIWidth = 1540
	$hGUIHeight = 770
	$PolicyInfoWidth = -1
	$PolicyInfoHeight = -1
	$HorizontalSpace = 25
	$HorizontalSpaceSml = 20
	$VerticalSpaceSml = 14
	$VerticalSpace = 68
	$PolicyInfoLabelWidth = 340
	$PolicyActionLabelWidth = 200
	$FilteringLabelWidth = 180
	$ToolslogsLabelWidth = 154
	$PolicyInfoLabelHeight= 30
	$PolicyActionLabelHeight = 30
	$FilteringLabelHeight = 30
	$ToolslogsLabelHeight = 30
	$ClientAreaTitlebar = 40
ElseIf $GetDPI = 192 Then
	$DPIScale = 200
	$ListViewWidth = 1640
	$ListViewHeight = 380
	$hGUIWidth = 1660
	$hGUIHeight = 800
	$PolicyInfoWidth = -1
	$PolicyInfoHeight = -1
	$HorizontalSpace = 25
	$HorizontalSpaceSml = 20
	$VerticalSpaceSml = 14
	$VerticalSpace = 68
	$PolicyInfoLabelWidth = 340
	$PolicyActionLabelWidth = 200
	$FilteringLabelWidth = 180
	$ToolslogsLabelWidth = 154
	$PolicyInfoLabelHeight= 30
	$PolicyActionLabelHeight = 30
	$FilteringLabelHeight = 30
	$ToolslogsLabelHeight = 30
	$ClientAreaTitlebar = 48
Else
	$ListViewWidth = 1200
	$ListViewHeight = 340
	$hGUIWidth = 1220
	$hGUIHeight = 634
	$PolicyInfoWidth = -1
	$PolicyInfoHeight = -1
	$HorizontalSpace = 25
	$HorizontalSpaceSml = 15
	$VerticalSpaceSml = 8
	$VerticalSpace = 48
	$PolicyInfoLabelWidth = 340
	$PolicyActionLabelWidth = 200
	$FilteringLabelWidth = 180
	$ToolslogsLabelWidth = 154
	$PolicyInfoLabelHeight= 30
	$PolicyActionLabelHeight = 30
	$FilteringLabelHeight = 30
	$ToolslogsLabelHeight = 30
	$ClientAreaTitlebar = 48
EndIf

Func _GetDPI()
    Local $iDPI, $iDPIRat, $Logpixelsy = 90, $hWnd = 0
    Local $hDC = DllCall("user32.dll", "long", "GetDC", "long", $hWnd)
    Local $aRet = DllCall("gdi32.dll", "long", "GetDeviceCaps", "long", $hDC[0], "long", $Logpixelsy)
    DllCall("user32.dll", "long", "ReleaseDC", "long", $hWnd, "long", $hDC)
    $iDPI = $aRet[0]
    ;; Set a ratio for the GUI dimensions based upon the current DPI value.
    If $iDPI < 145 And $iDPI > 121 Then
        $iDPIRat = $iDPI / 95
    ElseIf $iDPI < 121 And $iDPI > 84 Then
        $iDPIRat = $iDPI / 96
    ElseIf $iDPI < 84 And $iDPI > 0 Then
        $iDPIRat = $iDPI / 105
    ElseIf $iDPI = 0 Then
        $iDPI = 96
        $iDPIRat = 94
    Else
        $iDPIRat = $iDPI / 94
    EndIf
    Return SetError(0, $iDPIRat, $iDPI)
EndFunc


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

If $policycount = 0 Then
Global $aWords[1][11] = [["", "", "", "", "", "", "", "", "", "", ""]]


ElseIf $policycount = 1 Then
Global $aWords[1][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]]]


ElseIf $policycount = 2 Then
Global $aWords[2][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]]]


ElseIf $policycount = 3 Then
Global $aWords[3][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]]]


ElseIf $policycount = 4 Then
Global $aWords[4][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]]]


ElseIf $policycount = 5 Then
Global $aWords[5][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]]]


ElseIf $policycount = 6 Then
Global $aWords[6][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]]]


ElseIf $policycount = 7 Then
Global $aWords[7][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]]]


ElseIf $policycount = 8 Then
Global $aWords[8][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]]]


ElseIf $policycount = 9 Then
Global $aWords[9][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]]]


ElseIf $policycount = 10 Then
Global $aWords[10][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]]]


ElseIf $policycount = 11 Then
Global $aWords[11][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]]]


ElseIf $policycount = 12 Then
Global $aWords[12][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]]]


ElseIf $policycount = 13 Then
Global $aWords[13][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]]]


ElseIf $policycount = 14 Then
Global $aWords[14][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]]]


ElseIf $policycount = 15 Then
Global $aWords[15][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]]]


ElseIf $policycount = 16 Then
Global $aWords[16][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]]]


ElseIf $policycount = 17 Then
Global $aWords[17][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]]]


ElseIf $policycount = 18 Then
Global $aWords[18][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]]]


ElseIf $policycount = 19 Then
Global $aWords[19][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]]]


ElseIf $policycount = 20 Then
Global $aWords[20][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]]]


ElseIf $policycount = 21 Then
Global $aWords[21][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]]]


ElseIf $policycount = 22 Then
Global $aWords[22][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]]]


ElseIf $policycount = 23 Then
Global $aWords[23][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]]]


ElseIf $policycount = 24 Then
Global $aWords[24][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]]]


ElseIf $policycount = 25 Then
Global $aWords[25][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]]]


ElseIf $policycount = 26 Then
Global $aWords[26][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]] _
		, ["", $arpol[251], $arpol[252], $arpol[253], $arpol[254], $arpol[255], $arpol[256], $arpol[257], $arpol[258], $arpol[259], $arpol[260]]]


ElseIf $policycount = 27 Then
Global $aWords[27][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]] _
		, ["", $arpol[251], $arpol[252], $arpol[253], $arpol[254], $arpol[255], $arpol[256], $arpol[257], $arpol[258], $arpol[259], $arpol[260]] _
		, ["", $arpol[261], $arpol[262], $arpol[263], $arpol[264], $arpol[265], $arpol[266], $arpol[267], $arpol[268], $arpol[269], $arpol[270]]]


ElseIf $policycount = 28 Then
Global $aWords[28][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]] _
		, ["", $arpol[251], $arpol[252], $arpol[253], $arpol[254], $arpol[255], $arpol[256], $arpol[257], $arpol[258], $arpol[259], $arpol[260]] _
		, ["", $arpol[261], $arpol[262], $arpol[263], $arpol[264], $arpol[265], $arpol[266], $arpol[267], $arpol[268], $arpol[269], $arpol[270]] _
		, ["", $arpol[271], $arpol[272], $arpol[273], $arpol[274], $arpol[275], $arpol[276], $arpol[277], $arpol[278], $arpol[279], $arpol[280]]]


ElseIf $policycount = 29 Then
Global $aWords[29][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]] _
		, ["", $arpol[251], $arpol[252], $arpol[253], $arpol[254], $arpol[255], $arpol[256], $arpol[257], $arpol[258], $arpol[259], $arpol[260]] _
		, ["", $arpol[261], $arpol[262], $arpol[263], $arpol[264], $arpol[265], $arpol[266], $arpol[267], $arpol[268], $arpol[269], $arpol[270]] _
		, ["", $arpol[271], $arpol[272], $arpol[273], $arpol[274], $arpol[275], $arpol[276], $arpol[277], $arpol[278], $arpol[279], $arpol[280]] _
		, ["", $arpol[281], $arpol[282], $arpol[283], $arpol[284], $arpol[285], $arpol[286], $arpol[287], $arpol[288], $arpol[289], $arpol[290]]]


ElseIf $policycount = 30 Then
Global $aWords[30][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]] _
		, ["", $arpol[251], $arpol[252], $arpol[253], $arpol[254], $arpol[255], $arpol[256], $arpol[257], $arpol[258], $arpol[259], $arpol[260]] _
		, ["", $arpol[261], $arpol[262], $arpol[263], $arpol[264], $arpol[265], $arpol[266], $arpol[267], $arpol[268], $arpol[269], $arpol[270]] _
		, ["", $arpol[271], $arpol[272], $arpol[273], $arpol[274], $arpol[275], $arpol[276], $arpol[277], $arpol[278], $arpol[279], $arpol[280]] _
		, ["", $arpol[281], $arpol[282], $arpol[283], $arpol[284], $arpol[285], $arpol[286], $arpol[287], $arpol[288], $arpol[289], $arpol[290]] _
		, ["", $arpol[291], $arpol[292], $arpol[293], $arpol[294], $arpol[295], $arpol[296], $arpol[297], $arpol[298], $arpol[299], $arpol[300]]]


ElseIf $policycount = 31 Then
Global $aWords[31][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]] _
		, ["", $arpol[251], $arpol[252], $arpol[253], $arpol[254], $arpol[255], $arpol[256], $arpol[257], $arpol[258], $arpol[259], $arpol[260]] _
		, ["", $arpol[261], $arpol[262], $arpol[263], $arpol[264], $arpol[265], $arpol[266], $arpol[267], $arpol[268], $arpol[269], $arpol[270]] _
		, ["", $arpol[271], $arpol[272], $arpol[273], $arpol[274], $arpol[275], $arpol[276], $arpol[277], $arpol[278], $arpol[279], $arpol[280]] _
		, ["", $arpol[281], $arpol[282], $arpol[283], $arpol[284], $arpol[285], $arpol[286], $arpol[287], $arpol[288], $arpol[289], $arpol[290]] _
		, ["", $arpol[291], $arpol[292], $arpol[293], $arpol[294], $arpol[295], $arpol[296], $arpol[297], $arpol[298], $arpol[299], $arpol[300]] _
		, ["", $arpol[301], $arpol[302], $arpol[303], $arpol[304], $arpol[305], $arpol[306], $arpol[307], $arpol[308], $arpol[309], $arpol[310]]]


ElseIf $policycount = 32 Then
Global $aWords[32][11] = [["", $arpol[1], $arpol[2], $arpol[3], $arpol[4], $arpol[5], $arpol[6], $arpol[7], $arpol[8], $arpol[9], $arpol[10]] _
		, ["", $arpol[11], $arpol[12], $arpol[13], $arpol[14], $arpol[15], $arpol[16], $arpol[17], $arpol[18], $arpol[19], $arpol[20]] _
		, ["", $arpol[21], $arpol[22], $arpol[23], $arpol[24], $arpol[25], $arpol[26], $arpol[27], $arpol[28], $arpol[29], $arpol[30]] _
		, ["", $arpol[31], $arpol[32], $arpol[33], $arpol[34], $arpol[35], $arpol[36], $arpol[37], $arpol[38], $arpol[39], $arpol[40]] _
		, ["", $arpol[41], $arpol[42], $arpol[43], $arpol[44], $arpol[45], $arpol[46], $arpol[47], $arpol[48], $arpol[49], $arpol[50]] _
		, ["", $arpol[51], $arpol[52], $arpol[53], $arpol[54], $arpol[55], $arpol[56], $arpol[57], $arpol[58], $arpol[59], $arpol[60]] _
		, ["", $arpol[61], $arpol[62], $arpol[63], $arpol[64], $arpol[65], $arpol[66], $arpol[67], $arpol[68], $arpol[69], $arpol[70]] _
		, ["", $arpol[71], $arpol[72], $arpol[73], $arpol[74], $arpol[75], $arpol[76], $arpol[77], $arpol[78], $arpol[79], $arpol[80]] _
		, ["", $arpol[81], $arpol[82], $arpol[83], $arpol[84], $arpol[85], $arpol[86], $arpol[87], $arpol[88], $arpol[89], $arpol[90]] _
		, ["", $arpol[91], $arpol[92], $arpol[93], $arpol[94], $arpol[95], $arpol[96], $arpol[97], $arpol[98], $arpol[99], $arpol[100]] _
		, ["", $arpol[101], $arpol[102], $arpol[103], $arpol[104], $arpol[105], $arpol[106], $arpol[107], $arpol[108], $arpol[109], $arpol[110]] _
		, ["", $arpol[111], $arpol[112], $arpol[113], $arpol[114], $arpol[115], $arpol[116], $arpol[117], $arpol[118], $arpol[119], $arpol[120]] _
		, ["", $arpol[121], $arpol[122], $arpol[123], $arpol[124], $arpol[125], $arpol[126], $arpol[127], $arpol[128], $arpol[129], $arpol[130]] _
		, ["", $arpol[131], $arpol[132], $arpol[133], $arpol[134], $arpol[135], $arpol[136], $arpol[137], $arpol[138], $arpol[139], $arpol[140]] _
		, ["", $arpol[141], $arpol[142], $arpol[143], $arpol[144], $arpol[145], $arpol[146], $arpol[147], $arpol[148], $arpol[149], $arpol[150]] _
		, ["", $arpol[151], $arpol[152], $arpol[153], $arpol[154], $arpol[155], $arpol[156], $arpol[157], $arpol[158], $arpol[159], $arpol[160]] _
		, ["", $arpol[161], $arpol[162], $arpol[163], $arpol[164], $arpol[165], $arpol[166], $arpol[167], $arpol[168], $arpol[169], $arpol[170]] _
		, ["", $arpol[171], $arpol[172], $arpol[173], $arpol[174], $arpol[175], $arpol[176], $arpol[177], $arpol[178], $arpol[179], $arpol[180]] _
		, ["", $arpol[181], $arpol[182], $arpol[183], $arpol[184], $arpol[185], $arpol[186], $arpol[187], $arpol[188], $arpol[189], $arpol[190]] _
		, ["", $arpol[191], $arpol[192], $arpol[193], $arpol[194], $arpol[195], $arpol[196], $arpol[197], $arpol[198], $arpol[199], $arpol[200]] _
		, ["", $arpol[201], $arpol[202], $arpol[203], $arpol[204], $arpol[205], $arpol[206], $arpol[207], $arpol[208], $arpol[209], $arpol[210]] _
		, ["", $arpol[211], $arpol[212], $arpol[213], $arpol[214], $arpol[215], $arpol[216], $arpol[217], $arpol[218], $arpol[219], $arpol[220]] _
		, ["", $arpol[221], $arpol[222], $arpol[223], $arpol[224], $arpol[225], $arpol[226], $arpol[227], $arpol[228], $arpol[229], $arpol[230]] _
		, ["", $arpol[231], $arpol[232], $arpol[233], $arpol[234], $arpol[235], $arpol[236], $arpol[237], $arpol[238], $arpol[239], $arpol[240]] _
		, ["", $arpol[241], $arpol[242], $arpol[243], $arpol[244], $arpol[245], $arpol[246], $arpol[247], $arpol[248], $arpol[249], $arpol[250]] _
		, ["", $arpol[251], $arpol[252], $arpol[253], $arpol[254], $arpol[255], $arpol[256], $arpol[257], $arpol[258], $arpol[259], $arpol[260]] _
		, ["", $arpol[261], $arpol[262], $arpol[263], $arpol[264], $arpol[265], $arpol[266], $arpol[267], $arpol[268], $arpol[269], $arpol[270]] _
		, ["", $arpol[271], $arpol[272], $arpol[273], $arpol[274], $arpol[275], $arpol[276], $arpol[277], $arpol[278], $arpol[279], $arpol[280]] _
		, ["", $arpol[281], $arpol[282], $arpol[283], $arpol[284], $arpol[285], $arpol[286], $arpol[287], $arpol[288], $arpol[289], $arpol[290]] _
		, ["", $arpol[291], $arpol[292], $arpol[293], $arpol[294], $arpol[295], $arpol[296], $arpol[297], $arpol[298], $arpol[299], $arpol[300]] _
		, ["", $arpol[301], $arpol[302], $arpol[303], $arpol[304], $arpol[305], $arpol[306], $arpol[307], $arpol[308], $arpol[309], $arpol[310]] _
		, ["", $arpol[311], $arpol[312], $arpol[313], $arpol[314], $arpol[315], $arpol[316], $arpol[317], $arpol[318], $arpol[319], $arpol[320]]]


Else

Global $aWords[1][11] = [["", "Error:", "Error " & $policycorrect & " &" & " Error " & $policycount, "", "", "", "", "", "", "", ""]]

EndIf

; TODO: Add more policy arrays later

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

GetPolicyStatus()
Func GetPolicyStatus()
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


Func GUI2()

	$hGUI2 = GUICreate("App Control Policy Manager - EFI [Work-in-Progress]", 1220, 430, -1, -1)

	GUISetIcon(@ScriptFullPath, 201)
	GUISetFont(10, -1, -1, "Cascadia Mono")

	Local $exStyles = BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_CHECKBOXES, $LVS_EX_DOUBLEBUFFER), $EFIListView

	Local $hToolTip2 = _GUIToolTip_Create(0)
	_GUIToolTip_SetMaxTipWidth($hToolTip2, 286)

	$ExitButtonEFI = GUICtrlCreateButton("Close", 1048, 380, 88, 30)
	Local $hExitButtonEFI = GUICtrlGetHandle($ExitButtonEFI)

	$EFIListView = GUICtrlCreateListView("|Enforced|Policy Name|Policy ID|Base Policy ID|Version|On Disk|Signed Policy|System Policy|Authorized|Policy Options", 10, 10, $ListViewWidth, 340)

	$hEFIListView = GUICtrlGetHandle($EFIListView)

	_GUICtrlListView_SetExtendedListViewStyle($hEFIListView, $exStyles)



	If $isDarkMode = True Then
	;get handle to child SysHeader32 control of ListView
	Global $hHeader = HWnd(GUICtrlSendMsg($EFIListView, $LVM_GETHEADER, 0, 0))
	;Turn off theme for header
	DllCall("uxtheme.dll", "int", "SetWindowTheme", "hwnd", $hHeader, "wstr", "", "wstr", "")
	;subclass ListView to get at NM_CUSTOMDRAW notification sent to ListView
	Global $wProcNew = DllCallbackRegister("_LVWndProc", "ptr", "hwnd;uint;wparam;lparam")
	Global $wProcOld = _WinAPI_SetWindowLong($hEFIListView, $GWL_WNDPROC, DllCallbackGetPtr($wProcNew))

	;Optional: Flat Header - remove header 3D button effect
	Global $iStyle = _WinAPI_GetWindowLong($hHeader, $GWL_STYLE)
	_WinAPI_SetWindowLong($hHeader, $GWL_STYLE, BitOR($iStyle, $HDS_FLAT))
	EndIf

	;GUISetState(@SW_SHOW)

	If $isDarkMode = True Then
		GuiDarkmodeApply($hGUI2)
	Else
		Local $bEnableDarkTheme = False
		GuiLightmodeApply($hGUI2)
	EndIf

	_GUICtrlListView_SetColumnWidth($hEFIListView, 0, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 1, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 2, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 3, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 4, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 5, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 6, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 7, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 8, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 9, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 10, $LVSCW_AUTOSIZE_USEHEADER)

	GUISetState(@SW_SHOW)

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

	_GUICtrlListView_AddArray($hEFIListView,$EFIarray)

	; Read ListView content into an array
	$aContent = _GUIListViewEx_ReadToArray($EFIListView)

	_GUICtrlListView_SetColumnWidth($hEFIListView, 0, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 1, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 2, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 3, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 4, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 5, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 6, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 7, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 8, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 9, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($hEFIListView, 10, $LVSCW_AUTOSIZE_USEHEADER)

	If $isDarkMode = True Then
		GuiDarkmodeApply($hGUI2)
	Else
		Local $bEnableDarkTheme = False
		GuiLightmodeApply($hGUI2)
	EndIf

	; Initiate ListView
	$iLV_Index = _GUIListViewEx_Init($EFIListView, $aContent, 0, Default, False, 1)

	; Register required messages
	_GUIListViewEx_MsgRegister()

	;GUISetState(@SW_SHOW)


Endfunc


;$hGUI = GUICreate("App Control Policy Manager", 1220, 560, -1, -1, 0x00CF0000)
;$hGUI = GUICreate("App Control Policy Manager", 1220, 594, -1, -1, 0x00CF0000)
$hGUI = GUICreate("App Control Policy Manager", @DesktopWidth - 420, $hGUIHeight, -1, -1, $WS_SIZEBOX + $WS_SYSMENU + $WS_MINIMIZEBOX + $WS_MAXIMIZEBOX)

GUISetIcon(@ScriptFullPath, 201)

Local Const $sCascadiaPath = @WindowsDir & "\fonts\CascadiaCode.ttf"
Local $iCascadiaExists = FileExists($sCascadiaPath)

If $iCascadiaExists Then
	GUISetFont(10.5, $FW_NORMAL, -1, "Cascadia Code")
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
$cListView = GUICtrlCreateListView("|Enforced|Policy Name|Policy ID|Base Policy ID|Version|On Disk|Signed Policy|System Policy|Authorized|Policy Options", 10, $VerticalSpaceSml + $VerticalSpaceSml, @DesktopWidth - 440, @DesktopHeight / 3)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
$hListView = GUICtrlGetHandle($cListView)

$aPos = ControlGetPos($hGUI, "", $cListView)
;MsgBox($MB_SYSTEMMODAL, "", "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
$cListViewPosV = $aPos[1]
$cListViewPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$cListViewLength = $aPos[2]
$cListViewHeight = $aPos[3]

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


;$PolicyStatus = GUICtrlCreateLabel(" Policies (total)" & @TAB & @TAB & ": " & $CountTotal & @CRLF & " Policies (enforced)" & @TAB & @TAB & ": " & $CountEnforced & @CRLF & @CRLF & " " & $topstatus9 & @CRLF & @CRLF & $sVulnDrivermsg, 35, 410, 340, 156, $WS_BORDER)

UpdatePolicyStatus()
Func UpdatePolicyStatus()
Global $PolicyStatusInfo = " " & $topstatus9 & @CRLF & @CRLF & $sVulnDrivermsg & @CRLF & @CRLF & " Policies (total)      : " & $CountTotal & @CRLF & " Policies (enforced)   : " & $CountEnforced
EndFunc

If $iCascadiaExists Then
	GUISetFont(10, $FW_NORMAL, -1, "Cascadia Code")
Else
	GUISetFont(10, $FW_NORMAL, -1, "Consolas")
EndIf

;$PolicyStatus = GUICtrlCreateLabel($PolicyStatusInfo, 35, $Label2PosV + $Label2Height + $VerticalSpaceSml, $PolicyInfoWidth, $PolicyInfoHeight, $WS_BORDER + $SS_LEFTNOWORDWRAP)
$PolicyStatus = GUICtrlCreateLabel($PolicyStatusInfo, 35, $cListViewPosV + $cListViewHeight + $VerticalSpace, -1, -1, $WS_BORDER + $SS_LEFTNOWORDWRAP)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hPolicyStatus = GUICtrlGetHandle($PolicyStatus)
$aPos = ControlGetPos($hGUI, "", $PolicyStatus)

$PolicyStatusPosV = $aPos[1] + 4
$PolicyStatusPosV2 = $aPos[1]
$PolicyStatusPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$PolicyStatusWidth = $aPos[2]
$PolicyStatusHeight = $aPos[3]


GUISetFont(10, $FW_BOLD, -1, "Segoe UI")

$Label2 = GUICtrlCreateLabel("Current Policy Information:", 35, $cListViewPosV + $cListViewHeight + $VerticalSpaceSml, $PolicyStatusWidth, -1, $SS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hLabel2 = GUICtrlGetHandle($Label2)

$aPos = ControlGetPos($hGUI, "", $Label2)
;MsgBox($MB_SYSTEMMODAL, "", "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
$Label2PosV = $aPos[1]
$Label2PosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$Label2Length = $aPos[2]
$Label2Height = $aPos[3]

GUISetFont(8.5, $FW_NORMAL, -1, "Segoe UI")

If $is24H2 = True Then
	$pre24H2Lablel = GUICtrlCreateLabel(" ", 35, $PolicyStatusPosV2 + $PolicyStatusHeight + $VerticalSpaceSml + 4, $PolicyStatusWidth, -1, $SS_CENTER)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
Else
	$pre24H2Lablel = GUICtrlCreateLabel("* Pre-24H2 OS will show less information.", 35, $PolicyStatusPosV2 + $PolicyStatusHeight + $VerticalSpaceSml + 4, $PolicyStatusWidth, -1, $SS_CENTER)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	Local $hpre24H2Lablel = GUICtrlGetHandle($pre24H2Lablel)
EndIf

Local $hpre24H2Lablel = GUICtrlGetHandle($pre24H2Lablel)

$aPos = ControlGetPos($hGUI, "", $pre24H2Lablel)
;MsgBox($MB_SYSTEMMODAL, "", "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
$pre24H2LablelPosV = $aPos[1]
$pre24H2LablelPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$pre24H2LablelLength = $aPos[2]
$pre24H2LablelHeight = $aPos[3]
$pre24H2LablelbottomPos = $pre24H2LablelPosV + $pre24H2LablelHeight

;MsgBox($MB_SYSTEMMODAL, "", $pre24H2LablelbottomPos)

GUISetFont(10, $FW_NORMAL, -1, "Segoe UI")

$AddButton = GUICtrlCreateButton("    Add or Update Policies    ", $PolicyStatusPosH, $cListViewPosV + $cListViewHeight + $VerticalSpace, -1, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hAddButton = GUICtrlGetHandle($AddButton)

$aPos = ControlGetPos($hGUI, "", $AddButton)
;MsgBox($MB_SYSTEMMODAL, "", "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
$AddButtonPosV = $aPos[1]
$AddButtonPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$AddButtonLength = $aPos[2]
$AddButtonHeight = $aPos[3]

GUISetFont(10, $FW_BOLD, -1, "Segoe UI")

$PolicyActions = GUICtrlCreateLabel("Policy Actions:", $PolicyStatusPosH, $cListViewPosV + $cListViewHeight + $VerticalSpaceSml, $AddButtonLength, $PolicyActionLabelHeight, $SS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hPolicyActions = GUICtrlGetHandle($PolicyActions)

GUISetFont(10, $FW_NORMAL, -1, "Segoe UI")

$RemoveButton = GUICtrlCreateButton(" Remove Selected Policies ", $PolicyStatusPosH, $AddButtonPosV + $VerticalSpaceSml + $AddButtonHeight, $AddButtonLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hRemoveButton = GUICtrlGetHandle($RemoveButton)

$aPos = ControlGetPos($hGUI, "", $RemoveButton)

$RemoveButtonPosV = $aPos[1]
$RemoveButtonPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$RemoveButtonHeight = $aPos[3]

;$ConvertButton = GUICtrlCreateButton(" Convert (xml to binary) ", $PolicyStatusPosH, 530, -1, -1)
$ConvertButton = GUICtrlCreateButton("Convert XML to Binary", $PolicyStatusPosH, $RemoveButtonPosV + $RemoveButtonHeight + $VerticalSpaceSml, $AddButtonLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hConvertButton = GUICtrlGetHandle($ConvertButton)

$aPos = ControlGetPos($hGUI, "", $ConvertButton)

$ConvertButtonPosV = $aPos[1]
$ConvertButtonPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$ConvertButtonHeight = $aPos[3]

$RefreshButton = GUICtrlCreateButton("Refresh Policy List", $PolicyStatusPosH, $ConvertButtonPosV + $ConvertButtonHeight + $VerticalSpaceSml, $AddButtonLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hRefreshButton = GUICtrlGetHandle($RefreshButton)


$FilterEnforced = GUICtrlCreateButton(" Enforced ", $RemoveButtonPosH, $cListViewPosV + $cListViewHeight + $VerticalSpace, -1, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterEnforced = GUICtrlGetHandle($FilterEnforced)

$aPos = ControlGetPos($hGUI, "", $FilterEnforced)

$FilterEnforcedPosV = $aPos[1] + 4
$FilterEnforcedPosV2 = $aPos[1]
$FilterEnforcedPosH = $aPos[0] + $aPos[2] + $HorizontalSpaceSml
$FilterEnforcedLength = $aPos[2]
$FilterEnforcedHeight = $aPos[3]
;MsgBox($MB_SYSTEMMODAL, "", $FilterEnforcedLength)

GUISetFont(10, $FW_BOLD, -1, "Segoe UI")

$FilterLogs = GUICtrlCreateLabel("Filtering Options:", $RemoveButtonPosH, $cListViewPosV + $cListViewHeight + $VerticalSpaceSml, $FilterEnforcedLength + $FilterEnforcedLength + $HorizontalSpaceSml, $FilteringLabelHeight, $SS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterLogs = GUICtrlGetHandle($FilterLogs)

GUISetFont(10, $FW_NORMAL, -1, "Segoe UI")

$FilterBase = GUICtrlCreateButton("Base", $RemoveButtonPosH, $FilterEnforcedPosV2 + $FilterEnforcedHeight + $VerticalSpaceSml, $FilterEnforcedLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterBase = GUICtrlGetHandle($FilterBase)

$aPos = ControlGetPos($hGUI, "", $FilterBase)

$FilterBasePosV = $aPos[1]
$FilterBasePosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$FilterBaseHeight = $aPos[3]

$FilterSupp = GUICtrlCreateButton("Suppl.", $RemoveButtonPosH, $FilterBasePosV + $FilterBaseHeight + $VerticalSpaceSml, $FilterEnforcedLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterSupp = GUICtrlGetHandle($FilterSupp)
_GUIToolTip_AddTool($hToolTip2, 0, "Supplemental", $hFilterSupp)

$FilterSigned = GUICtrlCreateButton("Signed", $FilterEnforcedPosH, $cListViewPosV + $cListViewHeight + $VerticalSpace, $FilterEnforcedLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterSigned = GUICtrlGetHandle($FilterSigned)

$aPos = ControlGetPos($hGUI, "", $FilterSigned)

$FilterSignedPosV = $aPos[1]
$FilterSignedPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$FilterSignedHeight = $aPos[3]

If $is24H2 = False Then
$FilterSignedLabelOffset = $Label2PosV + $Label2Height + $VerticalSpaceSml
$FilterSignedLabel = GUICtrlCreateLabel("", $FilterEnforcedPosH - 2, $cListViewPosV + $cListViewHeight + $VerticalSpace - 2, $FilterEnforcedLength + 4, $FilterEnforcedHeight + 4, $WS_CLIPSIBLINGS)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterSignedLabel = GUICtrlGetHandle($FilterSignedLabel)
GUICtrlSetState($FilterSigned,$GUI_DISABLE)
_GUIToolTip_AddTool($hToolTip2, 0, "Policy list cannot be filtered by Signed policies on pre-24H2 builds of Windows 11.", $hFilterSignedLabel)
EndIf

$FilterSystem = GUICtrlCreateButton("System", $FilterEnforcedPosH, $FilterSignedPosV + $FilterSignedHeight + $VerticalSpaceSml, $FilterEnforcedLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterSystem = GUICtrlGetHandle($FilterSystem)

$aPos = ControlGetPos($hGUI, "", $FilterSystem)

$FilterSystemPosV = $aPos[1] + 4
$FilterSystemPosH = $aPos[0] + $aPos[2] + $HorizontalSpace

$FilterReset = GUICtrlCreateButton("Reset", $FilterEnforcedPosH, $FilterBasePosV + $FilterBaseHeight + $VerticalSpaceSml, $FilterEnforcedLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hFilterReset = GUICtrlGetHandle($FilterReset)

$aPos = ControlGetPos($hGUI, "", $FilterReset)

$FilterResetPosV = $aPos[1]
$FilterResetPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$FilterResetHeight = $aPos[3]

$EFIButton = GUICtrlCreateButton("EFI (Signed/System)", $RemoveButtonPosH, $FilterResetPosV + $FilterResetHeight + $VerticalSpaceSml, $FilterEnforcedLength * 2 + $HorizontalSpaceSml, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hEFIButton = GUICtrlGetHandle($EFIButton)

$aPos = ControlGetPos($hGUI, "", $EFIButton)
;MsgBox($MB_SYSTEMMODAL, "", "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
$EFIButtonPosV = $aPos[1] + 4
$EFIButtonPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$EFIButtonLength = $aPos[2]
;MsgBox($MB_SYSTEMMODAL, "", $EFIButtonLength)


$PolicyWizard = GUICtrlCreateButton(" App Control Wizard ", $FilterSystemPosH, $cListViewPosV + $cListViewHeight + $VerticalSpace, -1, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hPolicyWizard = GUICtrlGetHandle($PolicyWizard)

$aPos = ControlGetPos($hGUI, "", $PolicyWizard)
;MsgBox($MB_SYSTEMMODAL, "", "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
$PolicyWizardPosV = $aPos[1] + 4
$PolicyWizardPosV2 = $aPos[1]
$PolicyWizardPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$PolicyWizardLength = $aPos[2]
$PolicyWizardHeight = $aPos[3]

GUISetFont(10, $FW_BOLD, -1, "Segoe UI")

$EventLogs = GUICtrlCreateLabel("Tools && Logs:", $FilterSystemPosH, $cListViewPosV + $cListViewHeight + $VerticalSpaceSml, $PolicyWizardLength, $ToolslogsLabelHeight, $SS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hEventLogs = GUICtrlGetHandle($EventLogs)

GUISetFont(10, $FW_NORMAL, -1, "Segoe UI")

$CodeIntegrity = GUICtrlCreateButton("Code Integrity", $FilterSystemPosH, $PolicyWizardPosV2 + $PolicyWizardHeight + $PolicyWizardHeight + $VerticalSpaceSml + $VerticalSpaceSml, $PolicyWizardLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hCodeIntegrity = GUICtrlGetHandle($CodeIntegrity)

$aPos = ControlGetPos($hGUI, "", $CodeIntegrity)
;MsgBox($MB_SYSTEMMODAL, "", "Position: " & $aPos[0] & ", " & $aPos[1] & @CRLF & "Size: " & $aPos[2] & ", " & $aPos[3])
$CodeIntegrityPosV = $aPos[1] + 4
$CodeIntegrityPosV2 = $aPos[1]
$CodeIntegrityPosH = $aPos[0] + $aPos[2] + $HorizontalSpace
$CodeIntegrityLength = $aPos[2]
$CodeIntegrityHeight = $aPos[3]

$MSIandScript = GUICtrlCreateButton("MSI and Script", $FilterSystemPosH, $CodeIntegrityPosV2 + $CodeIntegrityHeight + $VerticalSpaceSml, $PolicyWizardLength, -1, $BS_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
Local $hMSIandScript = GUICtrlGetHandle($MSIandScript)


; Read ListView content into an array
$aContent = _GUIListViewEx_ReadToArray($cListView)


ApplyThemeColor()
Func ApplyThemeColor()

If $isDarkMode = True Then
	GuiDarkmodeApply($hGUI)
Else
	Local $bEnableDarkTheme = False
    GuiLightmodeApply($hGUI)
EndIf

Endfunc


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
EndFunc



$GUIHeightMeasured = $ClientAreaTitlebar + $pre24H2LablelbottomPos + 10


WinMove($hGUI,'',(@Desktopwidth - WinGetPos($hGUI)[2]) / 2,(@Desktopheight - WinGetPos($hGUI)[3]) / 2, @DesktopWidth - 420 + 6, $GUIHeightMeasured + 6)

;Sleep(500)

;GUISetState()

GUISetStyle($GUI_SS_DEFAULT_GUI, -1)

WinMove($hGUI,'', (@Desktopwidth - WinGetPos($hGUI)[2]) / 2,(@Desktopheight - WinGetPos($hGUI)[3]) / 2)

GUISetState(@SW_SHOWMINIMIZED)
GUISetState(@SW_RESTORE)

WinSetOnTop($hGUI, "", $WINDOWS_ONTOP)
WinSetOnTop($hGUI, "", $WINDOWS_NOONTOP)

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
		$sMsg = " Selected policy has been added. " & @CRLF
		$sMsg &= "" & @CRLF
		$sMsg &= " The policy list and status information will automatically " & @CRLF
		$sMsg &= " refresh in about 5 seconds. " & @CRLF
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
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
			$sMsg = " Selected policies have been added. " & @CRLF
			$sMsg &= "" & @CRLF
			$sMsg &= " The policy list and status information will automatically " & @CRLF
			$sMsg &= " refresh in about 5 seconds. " & @CRLF
			_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
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
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", " No policies have been selected for removal. " & @CRLF)
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
		$sMsg = " System policies cannot be removed with this method and " & @CRLF
		$sMsg &= " are protected in the EFI partition. " & @CRLF & @CRLF
		$sMsg &= " Please uncheck any system policies and try again. " & @CRLF
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
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
		$sMsg = " Signed base policies cannot be removed with this " & @CRLF
		$sMsg &= " method but Signed supplemental policies can. " & @CRLF & @CRLF
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
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
		$sMsg = " Selected policies have been removed. " & @CRLF
		$sMsg &= "" & @CRLF
		$sMsg &= " The policy list and status information will automatically " & @CRLF
		$sMsg &= " refresh in about 5 seconds. " & @CRLF
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlManager.exe", 0, "App Control Policy Manager", $sMsg)
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


Func MountEFI()
	Local $o_CmdString1 = " $MountPoint = $env:SystemDrive+'\EFIMount'; $EFIPartition = (Get-Partition | Where-Object IsSystem).AccessPaths[0]; if (-Not (Test-Path $MountPoint)) { New-Item -Path $MountPoint -Type Directory -Force }; mountvol $MountPoint $EFIPartition"
	Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide)
	ProcessWaitCloseEx($o_Pid)
	Local $systemdrive = EnvGet('SystemDrive')
	;MsgBox(0, '$systemdrive', $systemdrive)
	Local Const $sFilePath = $systemdrive & '\EFIMount\EFI\Microsoft'
	Local $iFileExists = FileExists($sFilePath)
	If $iFileExists Then
		ListEFI()
	EndIf
Endfunc

Func ListEFI()
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

	;$iCIPtest = $aWords[2][3]
	;MsgBox(0, 'test', $iCIPtest)
	;_ArrayDisplay($EFIarray, "test")
	ListEFIp7b()

Endfunc

Func ListEFIp7b()
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

	;$iCIPtest = $aWords[2][3]
	;MsgBox(0, 'test', $iCIPtest)
	_ArrayDisplay($EFIarray, "test")
	EFIlistview()

Endfunc

Func EFIlistview()

	_GUICtrlListView_DeleteAllItems($EFIListView)
	_GUICtrlListView_AddArray($hEFIListView,$EFIarray)

Endfunc

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


Func ProcessWaitCloseEx($iPID)
	While ProcessExists($iPID) And Sleep(10)
	WEnd
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
				Case $AddButton
					AddPolicies()
				Case $ConvertButton
					ConvertPolicy()
				Case $RefreshButton
					_GUICtrlListView_DeleteAllItems($cListView)
					GetPolicyInfo()
					GetPolicyStatus()
					_GUICtrlListView_AddArray($hListView,$aWords)
					hListViewRefreshColWidth()
					UpdatePolicyStatus()
					GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
				Case $RemoveButton
					CountChecked()
				Case $EFIButton
					;MountEFI()
					;GetPolicyInfo()
					GUISetState(@SW_HIDE, $hGUI)
					GUI2()
				Case $PolicyWizard
					PolicyWizard()
				Case $CodeIntegrity
					LogsCI()
				Case $MSIandScript
					LogsScript()
				Case $FilterEnforced
					_GUICtrlListView_DeleteAllItems($cListView)
					GetPolicyEnforced()
					GetPolicyStatus()
					_GUICtrlListView_AddArray($hListView,$aWords)
					hListViewRefreshColWidth()
					UpdatePolicyStatus()
					GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
				Case $FilterBase
					_GUICtrlListView_DeleteAllItems($cListView)
					GetPolicyBase()
					GetPolicyStatus()
					_GUICtrlListView_AddArray($hListView,$aWords)
					hListViewRefreshColWidth()
					UpdatePolicyStatus()
					GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
				Case $FilterSupp
					_GUICtrlListView_DeleteAllItems($cListView)
					GetPolicySupp()
					GetPolicyStatus()
					_GUICtrlListView_AddArray($hListView,$aWords)
					hListViewRefreshColWidth()
					UpdatePolicyStatus()
					GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
				Case $FilterSigned
					_GUICtrlListView_DeleteAllItems($cListView)
					GetPolicySigned()
					GetPolicyStatus()
					_GUICtrlListView_AddArray($hListView,$aWords)
					hListViewRefreshColWidth()
					UpdatePolicyStatus()
					GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
				Case $FilterSystem
					_GUICtrlListView_DeleteAllItems($cListView)
					GetPolicySystem()
					GetPolicyStatus()
					_GUICtrlListView_AddArray($hListView,$aWords)
					hListViewRefreshColWidth()
					UpdatePolicyStatus()
					GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
				Case $FilterReset
					_GUICtrlListView_DeleteAllItems($cListView)
					GetPolicyInfo()
					GetPolicyStatus()
					_GUICtrlListView_AddArray($hListView,$aWords)
					hListViewRefreshColWidth()
					UpdatePolicyStatus()
					GUICtrlSetData($PolicyStatus, $PolicyStatusInfo)
			EndSwitch
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