#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Policy Manager
#AutoIt3Wrapper_Res_Fileversion=4.5.0.0
#AutoIt3Wrapper_Res_ProductName=AppControlPolicyManager
#AutoIt3Wrapper_Res_ProductVersion=4.5.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#AutoIt3Wrapper_Res_Icon_Add=AppControl-Audit.ico
#AutoIt3Wrapper_Res_Icon_Add=AppControl-Disabled.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region
#include <AutoItConstants.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <String.au3>
#include <StringConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ColorConstants.au3>
#include <FontConstants.au3>
#include <GUIToolTip.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
#include <Constants.au3>
#include <ListViewConstants.au3>
#EndRegion

#include "includes\ExtMsgBox.au3"
#include "includes\GuiCtrls_HiDpi.au3"
#include "includes\GUIDarkMode_v0.02mod.au3"
#include "includes\GUIListViewEx.au3"
#include "includes\XML.au3"

Global $programversion = "4.5"

;Opt('MustDeclareVars', 1)

Global $isDarkMode = _WinAPI_ShouldAppsUseDarkMode()

If $isDarkMode = True Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, 9, -1, 1200)
Else
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, -1, -1, 9, -1, 1200)
EndIf

Global $hGUI, $cListView, $hListView
Global $out, $arpol, $arraycount, $policycount, $policyoutput, $policycorrect, $ExitButton, $TestButton, $Label2, $aContent, $iLV_Index

Global $iDllGDI = DllOpen("gdi32.dll")
Global $iDllUSER32 = DllOpen("user32.dll")

;Three column colours
Global $aCol[11][2] = [[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff],[0xffffff, 0xffffff]]

;Convert RBG to BGR for SetText/BkColor()
For $i = 0 To UBound($aCol)-1
    $aCol[$i][0] = _BGR2RGB($aCol[$i][0])
    $aCol[$i][1] = _BGR2RGB($aCol[$i][1])
Next

Global $ifPS7Exists = FileExists(@ProgramFilesDir & '\PowerShell\7\pwsh.exe')
If $ifPS7Exists Then
	Global $o_powershell = @ProgramFilesDir & '\PowerShell\7\pwsh.exe -NoProfile -Command'
Else
	Global $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Command"
EndIf

GetPolicyInfo()
Func GetPolicyInfo()
Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
ProcessWaitClose($o_Pid)
$out = StdoutRead($o_Pid)
CreatePolicyTable($out)
EndFunc

Func GetPolicyEnforced()
Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
ProcessWaitClose($o_Pid)
$out = StdoutRead($o_Pid)
CreatePolicyTable($out)
EndFunc

Func GetPolicyBase()
Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -eq $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
ProcessWaitClose($o_Pid)
$out = StdoutRead($o_Pid)
CreatePolicyTable($out)
EndFunc

Func GetPolicySupp()
Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.PolicyID -ne $_.BasePolicyID} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property FriendlyName | FL"
Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
ProcessWaitClose($o_Pid)
$out = StdoutRead($o_Pid)
CreatePolicyTable($out)
EndFunc

Func GetPolicySigned()
Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsSignedPolicy -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
ProcessWaitClose($o_Pid)
$out = StdoutRead($o_Pid)
CreatePolicyTable($out)
EndFunc

Func GetPolicySystem()
Local $o_CmdString1 = " (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsSystemPolicy -eq 'True'} | Select-Object -Property IsEnforced,FriendlyName,PolicyID,BasePolicyID,VersionString,IsOnDisk,IsSignedPolicy,IsSystemPolicy,IsAuthorized,PolicyOptions | Sort-Object -Descending -Property IsEnforced | Sort-Object -Property BasePolicyID,FriendlyName | FL"
Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
ProcessWaitClose($o_Pid)
$out = StdoutRead($o_Pid)
CreatePolicyTable($out)
EndFunc

Func CreatePolicyTable($out_passed)
Local $string1 = StringReplace($out_passed, "FriendlyName   : ", "")
Local $string2 = StringReplace($string1, "PolicyID       : ", "")
Local $string3 = StringReplace($string2, "BasePolicyID   : ", "")
Local $string4 = StringReplace($string3, "VersionString  : ", "")
Local $string5 = StringReplace($string4, "IsEnforced     : ", "")
Local $string6 = StringReplace($string5, "IsOnDisk       : ", "")
Local $string7 = StringReplace($string6, "IsSignedPolicy : ", "")
Local $string8 = StringReplace($string7, "IsSystemPolicy : ", "")
Local $string9 = StringReplace($string8, "IsAuthorized   : ", "")
Local $string10 = StringReplace($string9, "PolicyOptions  : ", "")
Local $string11 = StringReplace($string10, "[32;1m", "")
Local $string12 = StringReplace($string11, "[0m", "")
Local $string13 = StringReplace($string12, "                 ", " ")
Local $string14 = StringStripWS($string13, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
$arpol = stringsplit($string14 , @CR , 0)
Global $policyoutput = UBound ($arpol, $UBOUND_ROWS)
Global $policycorrect = $policyoutput - 1
Global $policycount = $policycorrect / 10


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

Else

Global $aWords[1][11] = [["", "Error:", "Error " & $policycorrect & " &" & " Error " & $policycount, "", "", "", "", "", "", "", ""]]

EndIf

; TODO: Add more policy arrays later

Endfunc

GetPolicyStatus()
Func GetPolicyStatus()
Local $o_CmdString1 = " Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity*"
Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
ProcessWaitClose($o_Pid)
$out = StdoutRead($o_Pid)

Local $topstatus1 = StringReplace($out, "[32;1m", "")
Local $topstatus2 = StringReplace($topstatus1, "[0m", "")
Local $topstatus3 = StringReplace($topstatus2, "UsermodeCodeIntegrityPolicyEnforcementStatus : 0", " App Control user mode policy" & @TAB & ": Not Configured")
Local $topstatus4 = StringReplace($topstatus3, "UsermodeCodeIntegrityPolicyEnforcementStatus : 1", " App Control user mode policy" & @TAB & ": Audit Mode")
Local $topstatus5 = StringReplace($topstatus4, "UsermodeCodeIntegrityPolicyEnforcementStatus : 2", " App Control user mode policy" & @TAB & ": Enforced Mode")
Local $topstatus6 = StringReplace($topstatus5, "CodeIntegrityPolicyEnforcementStatus         : 0", "App Control policy" & @TAB & @TAB & ": Not Configured")
Local $topstatus7 = StringReplace($topstatus6, "CodeIntegrityPolicyEnforcementStatus         : 1", "App Control policy" & @TAB & @TAB & ": Audit Mode")
Local $topstatus8 = StringReplace($topstatus7, "CodeIntegrityPolicyEnforcementStatus         : 2", "App Control policy" & @TAB & @TAB & ": Enforced Mode")
Global $topstatus9 = StringStripWS($topstatus8, $STR_STRIPLEADING + $STR_STRIPTRAILING)

;MsgBox($MB_SYSTEMMODAL, "Title", $topstatus9)

Endfunc

Local $exStyles = BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_CHECKBOXES, $LVS_EX_INFOTIP), $cListView

$hGUI = GUICreate("App Control Policy Manager", 1220, 550, -1, -1, 0x00CF0000)

GUISetIcon(@ScriptFullPath, 201)

Local $hToolTip2 = _GUIToolTip_Create(0)

$AboutButton = GUICtrlCreateButton("About", 1085, 490, 100, 30)
Local $hAboutButton = GUICtrlGetHandle($AboutButton)

$ExitButton = GUICtrlCreateButton("Close", 1085, 410, 100, 30)
Local $hExitButton = GUICtrlGetHandle($ExitButton)

$PolicyActions = GUICtrlCreateLabel("Policy Actions:", 468, 370, 220, 30)
Local $hPolicyActions = GUICtrlGetHandle($PolicyActions)

$AddButton = GUICtrlCreateButton("Add or Update Policies", 425, 410, 200, 30)
Local $hAddButton = GUICtrlGetHandle($AddButton)

$RemoveButton = GUICtrlCreateButton("Remove Selected Policies", 425, 450, 200, 30)
Local $hRemoveButton = GUICtrlGetHandle($RemoveButton)

$RefreshButton = GUICtrlCreateButton("Refresh Policy List", 425, 490, 200, 30)
Local $hRefreshButton = GUICtrlGetHandle($RefreshButton)

$FilterLogs = GUICtrlCreateLabel("Filtering Options:", 692, 370, 220, 30)
Local $hFilterLogs = GUICtrlGetHandle($FilterLogs)

$FilterEnforced = GUICtrlCreateButton("Enforced", 675, 410, 85, 30)
Local $hFilterEnforced = GUICtrlGetHandle($FilterEnforced)

$FilterBase = GUICtrlCreateButton("Base", 675, 450, 85, 30)
Local $hFilterBase = GUICtrlGetHandle($FilterBase)

$FilterSupp = GUICtrlCreateButton("Suppl.", 675, 490, 85, 30)
Local $hFilterSupp = GUICtrlGetHandle($FilterSupp)
_GUIToolTip_AddTool($hToolTip2, 0, "Supplemental", $hFilterSupp)

$FilterSigned = GUICtrlCreateButton("Signed", 770, 410, 85, 30)
Local $hFilterSigned = GUICtrlGetHandle($FilterSigned)

$FilterSystem = GUICtrlCreateButton("System", 770, 450, 85, 30)
Local $hFilterSystem = GUICtrlGetHandle($FilterSystem)

$FilterReset = GUICtrlCreateButton("Reset", 770, 490, 85, 30)
Local $hFilterReset = GUICtrlGetHandle($FilterReset)

$EventLogs = GUICtrlCreateLabel("Event Logs:", 923, 370, 100, 30)
Local $hEventLogs = GUICtrlGetHandle($EventLogs)

$CodeIntegrity = GUICtrlCreateButton("Code Integrity", 905, 410, 130, 30)
Local $hCodeIntegrity = GUICtrlGetHandle($CodeIntegrity)

$MSIandScript = GUICtrlCreateButton("MSI and Script", 905, 450, 130, 30)
Local $hMSIandScript = GUICtrlGetHandle($MSIandScript)

$cListView = GUICtrlCreateListView("|Enforced|Policy Name|Policy ID|Base Policy ID|Version|On Disk|Signed Policy|System Policy|Authorized|Policy Options", 10, 10, 1200, 340, $LVS_EX_DOUBLEBUFFER)

$hListView = GUICtrlGetHandle($cListView)

_GUICtrlListView_SetExtendedListViewStyle($hListView, $exStyles)

_GUICtrlListView_AddArray($hListView,$aWords)

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

$Label2 = GUICtrlCreateLabel("Policy Information:", 120, 370, 200, 25)
Local $hLabel2 = GUICtrlGetHandle($Label2)
$PolicyStatus = GUICtrlCreateLabel(" " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9, 35, 410, 340, 94, $WS_BORDER)
Local $hPolicyStatus = GUICtrlGetHandle($PolicyStatus)

_GUICtrlListView_SetColumnWidth($hListView, 0, $LVSCW_AUTOSIZE_USEHEADER)
_GUICtrlListView_SetColumnWidth($hListView, 1, $LVSCW_AUTOSIZE_USEHEADER)
_GUICtrlListView_SetColumnWidth($hListView, 2, 355)
_GUICtrlListView_SetColumnWidth($hListView, 3, 295)
_GUICtrlListView_SetColumnWidth($hListView, 4, 295)
_GUICtrlListView_SetColumnWidth($hListView, 5, 95)
_GUICtrlListView_SetColumnWidth($hListView, 6, $LVSCW_AUTOSIZE_USEHEADER)
_GUICtrlListView_SetColumnWidth($hListView, 7, $LVSCW_AUTOSIZE_USEHEADER)
_GUICtrlListView_SetColumnWidth($hListView, 8, $LVSCW_AUTOSIZE_USEHEADER)
_GUICtrlListView_SetColumnWidth($hListView, 9, $LVSCW_AUTOSIZE_USEHEADER)
_GUICtrlListView_SetColumnWidth($hListView, 10, $LVSCW_AUTOSIZE)

If $isDarkMode = True Then
GUICtrlSetColor($Label2, $COLOR_WHITE)
GUICtrlSetFont($Label2, 11, -1, -1, "Segoe UI")
GUICtrlSetColor($EventLogs, $COLOR_WHITE)
GUICtrlSetFont($EventLogs, 11, -1, -1, "Segoe UI")
GUICtrlSetColor($PolicyActions, $COLOR_WHITE)
GUICtrlSetColor($FilterLogs, $COLOR_WHITE)
GUICtrlSetFont($FilterLogs, 11, -1, -1, "Segoe UI")

Else
GUICtrlSetColor($Label2, $COLOR_BLACK)
GUICtrlSetFont($Label2, 11, -1, -1, "Segoe UI")
GUICtrlSetColor($EventLogs, $COLOR_BLACK)
GUICtrlSetFont($EventLogs, 11, -1, -1, "Segoe UI")
GUICtrlSetColor($PolicyActions, $COLOR_BLACK)
GUICtrlSetColor($FilterLogs, $COLOR_BLACK)
GUICtrlSetFont($FilterLogs, 11, -1, -1, "Segoe UI")
EndIf
GUICtrlSetFont($PolicyActions, 11, -1, -1, "Segoe UI")
GUICtrlSetFont($AboutButton, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($ExitButton, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($RemoveButton, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($AddButton, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($RefreshButton, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($PolicyStatus, 10, -1, -1, "Segoe UI")
GUICtrlSetFont($CodeIntegrity, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($MSIandScript, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($FilterEnforced, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($FilterBase, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($FilterSupp, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($FilterSigned, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($FilterSystem, 9, -1, -1, "Segoe UI")
GUICtrlSetFont($FilterReset, 9, -1, -1, "Segoe UI")

; Read ListView content into an array
$aContent = _GUIListViewEx_ReadToArray($cListView)

If $isDarkMode = True Then
	GuiDarkmodeApply($hGUI)
Else
	Local $bEnableDarkTheme = False
    GuiLightmodeApply($hGUI)
EndIf

; Initiate ListView
$iLV_Index = _GUIListViewEx_Init($cListView, $aContent, 0, Default, False, 1)

; Register required messages
_GUIListViewEx_MsgRegister()

GUISetState(@SW_SHOW)

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
		GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
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
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
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
		GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
Endfunc

Global $CountEnforced = 0
Func CountEnforced()
	Global $CountEnforced = 0
	For $I = 0 To _GUICtrlListView_GetItemCount($cListView) - 1
		$Array = _GUICtrlListView_GetItemTextArray($cListView, $I)
		If $Array[2] = "True" Then $CountEnforced += 1
	Next
	;ConsoleWrite("Count = " & $CountEnforced & @CRLF)
	MsgBox($MB_SYSTEMMODAL, "Title", "Count = " & $CountEnforced)
Endfunc

While 1
    Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE, $ExitButton
			Exit
		Case $AddButton
			AddPolicies()
		Case $RefreshButton
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicyInfo()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
		Case $RemoveButton
			CountChecked()
		Case $AboutButton
			About()
			;CountEnforced()
		Case $CodeIntegrity
			LogsCI()
		Case $MSIandScript
			LogsScript()
		Case $FilterEnforced
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicyEnforced()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
		Case $FilterBase
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicyBase()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
		Case $FilterSupp
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicySupp()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
		Case $FilterSigned
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicySigned()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
		Case $FilterSystem
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicySystem()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
		Case $FilterReset
			_GUICtrlListView_DeleteAllItems($cListView)
			GetPolicyInfo()
			GetPolicyStatus()
			_GUICtrlListView_AddArray($hListView,$aWords)
			GUICtrlSetData($PolicyStatus, " " & "App Control Policies" & @TAB & @TAB & ": " & $policycount & @CRLF & @CRLF & " " & $topstatus9)
	EndSwitch
;_GUIListViewEx_EventMonitor()

	;If $wProcOld Then _WinAPI_SetWindowLong($hListView, $GWL_WNDPROC, $wProcOld)
	; Delete callback function
	;If $wProcNew Then DllCallbackFree($wProcNew)
WEnd