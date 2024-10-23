#Region
#include <AutoItConstants.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <String.au3>
#EndRegion

#include "includes\ExtMsgBox.au3"
#include "includes\XML.au3"

#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Tray Tool
#AutoIt3Wrapper_Res_Fileversion=3.1.0.0
#AutoIt3Wrapper_Res_ProductVersion=3.1.0
#AutoIt3Wrapper_Res_ProductName=AppControlTrayTool
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
$sTitle = "AppControlTrayHelper"

Global $isDarkMode = is_app_dark_theme()

If $isDarkMode = True Then
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, 0x202020, 0xFFFFFF, 10, "Consolas", 800)
Else
	_ExtMsgBoxSet(Default)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	_ExtMsgBoxSet(1, 4, -1, -1, 10, "Consolas", 800)
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
                Local $o_powershell = @ProgramFilesDir & '\PowerShell\7\pwsh.exe'
        Else
                Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        EndIf

	Local $o_CmdString1 = ' ./Get-CiPolicy.ps1'
	;Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
	;Local $o_powershell = "pwsh.exe"
    Local $o_Pid = Run($o_powershell & $o_CmdString1 , @ScriptDir & '\scripts', @SW_Hide)
    ProcessWaitClose($o_Pid)

	Local $testopen = @TempDir & '\CiPolicy.txt'
	Local $testread = FileRead($testopen)
	;MsgBox($MB_SYSTEMMODAL, "Title", $testread)
	;_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
	$iRetValue = _ExtMsgBox (0 & ";" & @ScriptDir & "\WDACTrayTool.exe", 0, "App Control Policy Status", $testread)
	FileDelete(@TempDir & '\CiPolicy.txt')
	FileDelete(@TempDir & '\CiPolicy1.txt')
	FileDelete(@TempDir & '\CiPolicy2.txt')
EndIf
If $CmdLine[1] = "/LogsCI" Then
	$oXML=ObjCreate("Microsoft.XMLDOM")
	$stest = FileRead(@AppDataDir & "\Microsoft\MMC\eventvwr")
	$oXML.LoadXML($stest)

	If Not _XML_NodeExists($oXML, "//DynamicPath") Then
			Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-CodeIntegrity/Operational" /f:"*[System/EventID=3004] or *[System/EventID=3033] or *[System/EventID=3034] or *[System/EventID=3076] or *[System/EventID=3077] or *[System/EventID=3089]"', "", @SW_HIDE)
	Else
	$oXMLChild1 = $oXML.selectSingleNode( "//DynamicPath" )
	$parent = $oXMLChild1.ParentNode

	$test = $parent.removeChild ( $oXMLChild1 )

	$oXML.Save(@AppDataDir & "\Microsoft\MMC\eventvwr")
	Sleep(500)
	Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-CodeIntegrity/Operational" /f:"*[System/EventID=3004] or *[System/EventID=3033] or *[System/EventID=3034] or *[System/EventID=3076] or *[System/EventID=3077] or *[System/EventID=3089]"', "", @SW_HIDE)
	EndIf
EndIf
If $CmdLine[1] = "/LogsScript" Then
	$oXML=ObjCreate("Microsoft.XMLDOM")
	$stest = FileRead(@AppDataDir & "\Microsoft\MMC\eventvwr")
	$oXML.LoadXML($stest)

	If Not _XML_NodeExists($oXML, "//DynamicPath") Then
			Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-AppLocker/MSI and Script" /f:"*[System/EventID=8028] or *[System/EventID=8029] or *[System/EventID=8036] or *[System/EventID=8037] or *[System/EventID=8038] or *[System/EventID=8039] or *[System/EventID=8040]"', "", @SW_HIDE)
	Else
	$oXMLChild1 = $oXML.selectSingleNode( "//DynamicPath" )
	$parent = $oXMLChild1.ParentNode

	$test = $parent.removeChild ( $oXMLChild1 )

	$oXML.Save(@AppDataDir & "\Microsoft\MMC\eventvwr")
	Sleep(500)
	Run(@ComSpec & " /c " & 'eventvwr /c:"Microsoft-Windows-AppLocker/MSI and Script" /f:"*[System/EventID=8028] or *[System/EventID=8029] or *[System/EventID=8036] or *[System/EventID=8037] or *[System/EventID=8038] or *[System/EventID=8039] or *[System/EventID=8040]"', "", @SW_HIDE)
	EndIf
EndIf

If $CmdLine[1] = "/AddPolicies" Then

Local $spFile

$mFile = FileOpenDialog("Select Policy File(s)", @ScriptDir & "\policies\", "Policy Files (*.cip)", 1 + 4 )
If @error Then
    ConsoleWrite("error")
Else
	$spFile = StringSplit($mFile, "|")
     If UBound($spFile) = 2 Then
		$path = $spFile[1]
        _ArrayDelete($spFile, 0)
		$aExtract = _StringBetween($spFile[0], "{", "}")
		Local $cmd1 = ' (citool.exe -up '
		Local $cmd2 = '"{'
		Local $cmd3 = $aExtract[0]
		Local $cmd4 = '}.cip"'
		Local $cmd5 = ' -json)'
		Run(@ComSpec & " /c " & $cmd1 & $cmd2 & $cmd3 & $cmd4 & $cmd5, "", @SW_HIDE)
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

$mFile = FileOpenDialog("Select Policy File(s)", "C:\Windows\System32\CodeIntegrity\CIPolicies\Active\", "Policy Files (*.cip)", 1 + 4 )
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
