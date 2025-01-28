#Region
#include <AutoItConstants.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <String.au3>
#include <StringConstants.au3>
#EndRegion

#include "includes\ExtMsgBox.au3"
#include "includes\XML.au3"

#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Tray Helper
#AutoIt3Wrapper_res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Fileversion=5.2.0.0
#AutoIt3Wrapper_Res_ProductVersion=5.2.0
#AutoIt3Wrapper_Res_ProductName=AppControlTrayHelper
#AutoIt3Wrapper_Outfile_x64=AppControlHelper.exe
#AutoIt3Wrapper_Res_LegalCopyright=@ 2025 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
$sTitle = "AppControlTrayHelper"

If @Compiled = 0 Then
	; System aware DPI awareness
	;DllCall("User32.dll", "bool", "SetProcessDPIAware")
	; Per-monitor V2 DPI awareness
	DllCall("User32.dll", "bool", "SetProcessDpiAwarenessContext" , "HWND", "DPI_AWARENESS_CONTEXT" -4)
EndIf

Global $isDarkMode = is_app_dark_theme()

If $isDarkMode = True Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, 10, "Cascadia Mono", 1200)
Else
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, -1, -1, 10, "Cascadia Mono", 1200)
EndIf

#cs ----------------------------------------------------------------------------
 Function    : is_app_dark_theme()
 Description : returns if the user has enabled the dark theme for applications in the Windows settings (0 on / 1 off)
               if OS too old (key does not exist) the key returns nothing, so function returns False
#ce ----------------------------------------------------------------------------
func is_app_dark_theme()
    return(regread('HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize', 'AppsUseLightTheme') == 0) ? True : False
endfunc

If $CmdLine[0] = 0 Then Exit MsgBox(16, $sTitle, "No parameters passed!")
If $CmdLine[1] = "/CiTool" Then
Run('C:\Windows\System32\CiTool.exe --list-policies', "")
EndIf
If $CmdLine[1] = "/status" Then
	Local $iFileExists = FileExists(@ProgramFilesDir & '\PowerShell\7\pwsh.exe')
	If $iFileExists Then
		Local $o_CmdString1 = " Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity*; Write-Output ''; Write-Output 'Active Base Policies:'; (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True' -and $_.PolicyID -eq $_.BasePolicyID} | Select-Object -Property FriendlyName,PolicyID,VersionString | Sort-Object -Property FriendlyName | FT; Write-Output ''; Write-Output 'Active Supplemental Policies:'; (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True' -and $_.PolicyID -ne $_.BasePolicyID} | Select-Object -Property FriendlyName,PolicyID,VersionString | Sort-Object -Property FriendlyName | FT; Write-Output ''; Write-Output 'Inactive Policies:'; (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -ne 'True'} | Select-Object -Property FriendlyName,PolicyID,VersionString | Sort-Object -Property FriendlyName | FT"
		Local $o_powershell = @ProgramFilesDir & '\PowerShell\7\pwsh.exe -NoProfile -Command'
		Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
		ProcessWaitClose($o_Pid)
		$out = StdoutRead($o_Pid)
		$test = stringsplit($out, @CR, 0)
		;ReDim $test[$test[0] - 1]
		Local $topstatus = _ArrayToString($test, "", 2, 3)
		Local $topstatus1 = StringReplace($topstatus, "[32;1m", "")
		Local $topstatus2 = StringReplace($topstatus1, "[0m", "")
		Local $topstatus3 = StringReplace($topstatus2, "UsermodeCodeIntegrityPolicyEnforcementStatus : 0", "App Control user mode policy    : Not Configured")
		Local $topstatus4 = StringReplace($topstatus3, "UsermodeCodeIntegrityPolicyEnforcementStatus : 1", "App Control user mode policy    : Audit Mode")
		Local $topstatus5 = StringReplace($topstatus4, "UsermodeCodeIntegrityPolicyEnforcementStatus : 2", "App Control user mode policy    : Enforced Mode")
		Local $topstatus6 = StringReplace($topstatus5, "CodeIntegrityPolicyEnforcementStatus         : 0", "App Control policy              : Not Configured")
		Local $topstatus7 = StringReplace($topstatus6, "CodeIntegrityPolicyEnforcementStatus         : 1", "App Control policy              : Audit Mode")
		Local $topstatus8 = StringReplace($topstatus7, "CodeIntegrityPolicyEnforcementStatus         : 2", "App Control policy              : Enforced Mode")
		Local $topstatus9 = StringStripWS($topstatus8, $STR_STRIPLEADING + $STR_STRIPTRAILING)

		Local $rangetest = "0-5"
		_ArrayDelete($test, $rangetest)
		Local $test2 = _ArrayToString($test, "")
		Local $test3 = StringStripWS($test2, $STR_STRIPLEADING + $STR_STRIPTRAILING)
		Local $test4 = StringReplace($test3, "[32;1m", "")
		Local $test5 = StringReplace($test4, "[0m", "")
		Local $test6 = StringReplace($test5, "FriendlyName", "Policy Name")
		Local $test7 = StringReplace($test6, " VersionString", "Version")
		Local $test8 = StringReplace($test7, "PolicyID", " Policy ID")
		Local $test9 = StringReplace($test8, " -------- ", "   ---------")
		Local $test10 = StringReplace($test9, "------------ ", "-----------")
		Local $test11 = StringReplace($test10, "-------------", "-------")
		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlHelper.exe", 0, "App Control Policy List", $test11 & @CRLF)
	Else
		Local $o_CmdString1 = " Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity*; Write-Output 'Active Base Policies:'; (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True' -and $_.PolicyID -eq $_.BasePolicyID} | Select-Object -Property FriendlyName,PolicyID,VersionString | Sort-Object -Property FriendlyName | FT; Write-Output 'Active Supplemental Policies:'; (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True' -and $_.PolicyID -ne $_.BasePolicyID} | Select-Object -Property FriendlyName,PolicyID,VersionString | Sort-Object -Property FriendlyName | FT; Write-Output 'Inactive Policies:'; (CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -ne 'True'} | Select-Object -Property FriendlyName,PolicyID,VersionString | Sort-Object -Property FriendlyName | FT"
		Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Command"
		Local $o_Pid = Run($o_powershell & $o_CmdString1 , "", @SW_Hide, $STDOUT_CHILD)
		ProcessWaitClose($o_Pid)
		$out = StdoutRead($o_Pid)
		$test = stringsplit($out , @CR , 0)

		Local $topstatus = _ArrayToString($test, "", 3, 4)
		Local $topstatus1 = StringReplace($topstatus, "[32;1m", "")
		Local $topstatus2 = StringReplace($topstatus1, "[0m", "")
		Local $topstatus3 = StringReplace($topstatus2, "UsermodeCodeIntegrityPolicyEnforcementStatus : 0", "App Control user mode policy    : Not Configured")
		Local $topstatus4 = StringReplace($topstatus3, "UsermodeCodeIntegrityPolicyEnforcementStatus : 1", "App Control user mode policy    : Audit Mode")
		Local $topstatus5 = StringReplace($topstatus4, "UsermodeCodeIntegrityPolicyEnforcementStatus : 2", "App Control user mode policy    : Enforced Mode")
		Local $topstatus6 = StringReplace($topstatus5, "CodeIntegrityPolicyEnforcementStatus         : 0", "App Control policy              : Not Configured")
		Local $topstatus7 = StringReplace($topstatus6, "CodeIntegrityPolicyEnforcementStatus         : 1", "App Control policy              : Audit Mode")
		Local $topstatus8 = StringReplace($topstatus7, "CodeIntegrityPolicyEnforcementStatus         : 2", "App Control policy              : Enforced Mode")
		Local $topstatus9 = StringStripWS($topstatus8, $STR_STRIPLEADING + $STR_STRIPTRAILING)

		Local $rangetest = "0-5"
		_ArrayDelete($test, $rangetest)
		Local $test2 = _ArrayToString($test, "")
		Local $test3 = StringStripWS($test2, $STR_STRIPLEADING + $STR_STRIPTRAILING)
		Local $test4 = StringReplace($test3, "[32;1m", "")
		Local $test5 = StringReplace($test4, "[0m", "")
		Local $test6 = StringReplace($test5, "FriendlyName", "Policy Name")
		Local $test7 = StringReplace($test6, " VersionString", "Version")
		Local $test8 = StringReplace($test7, "PolicyID", " Policy ID")
		Local $test9 = StringReplace($test8, " -------- ", "   ---------")
		Local $test10 = StringReplace($test9, "------------ ", "-----------")
		Local $test11 = StringReplace($test10, "-------------", "-------")

		_ExtMsgBox (0 & ";" & @ScriptDir & "\AppControlHelper.exe", 0, "App Control Policy List", $test11 & @CRLF)
	EndIf
EndIf
If $CmdLine[1] = "/LogsCI" Then
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
EndIf
If $CmdLine[1] = "/LogsScript" Then
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
EndIf

If $CmdLine[1] = "/AddPolicies" Then

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
    EndIf
EndIf

EndIf
If $CmdLine[1] = "/RemovePolicies" Then

Local $spFile

$mFile = FileOpenDialog("Select Policy File(s) for Removal", "C:\Windows\System32\CodeIntegrity\CIPolicies\Active\", "Policy Files (*.cip)", 1 + 4 )
If @error Then
    ConsoleWrite("error")
Else
	$spFile = StringSplit($mFile, "|")
     If UBound($spFile) = 2 Then
		$path = $spFile[1]
        _ArrayDelete($spFile, 0)
		$aExtract = _StringBetween($spFile[0], "{", "}")
		Local $cmd1 = ' (citool.exe -rp '
		Local $cmd2 = '"{'
		Local $cmd3 = $aExtract[0]
		Local $cmd4 = '}"'
		Local $cmd5 = ' -json)'
		Run(@ComSpec & " /c " & $cmd1 & $cmd2 & $cmd3 & $cmd4 & $cmd5, "", @SW_HIDE)
     Else
		$path = $spFile[1]
        _ArrayDelete($spFile, 0)
        _ArrayDelete($spFile, 0)
        For $x = 0 to UBound($spFile)-1
			$aExtract = _StringBetween($spFile[$x], "{", "}")
			Local $cmd1 = ' (citool.exe -rp '
			Local $cmd2 = '"{'
			Local $cmd3 = $aExtract[0]
			Local $cmd4 = '}"'
			Local $cmd5 = ' -json)'
			Run(@ComSpec & " /c " & $cmd1 & $cmd2 & $cmd3 & $cmd4 & $cmd5, "", @SW_HIDE)
        Next
    EndIf
EndIf

EndIf
