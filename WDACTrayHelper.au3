#Region ; *** Dynamically added Include files ***
#include <AutoItConstants.au3>                               ; added:08/18/24 07:37:14
#include <File.au3>                                          ; added:08/18/24 07:37:14
#include <MsgBoxConstants.au3>                               ; added:08/18/24 07:37:14
#include <XML.au3>                                           ; added:08/18/24 07:37:14
#EndRegion ; *** Dynamically added Include files ***
#include <ExtMsgBox.au3>
#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=WDAC.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=WDAC Tray Tool
#AutoIt3Wrapper_Res_Fileversion=2.5.0.0
#AutoIt3Wrapper_Res_ProductName=WDACTrayTool
#AutoIt3Wrapper_Res_ProductVersion=2.5.0
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=WDAC.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
$sTitle = "WDACTrayHelper"
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
	_ExtMsgBoxSet(1, 5, -1, -1, -1, "Consolas", 800, 800)
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
If $CmdLine[1] = "/AllowAll" Then
	Local $ps_Cipath = @SystemDir & "\CodeIntegrity\CiPolicies\Active\"
	Local $ps_copy = ' Copy-Item *.cip -Destination ' & $ps_Cipath & ' -Force; Start-Sleep -Seconds 1; (citool.exe -r -json) '
	Local $ps_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    Local $ps_Pid = Run($ps_powershell & $ps_copy, @ScriptDir & '\policies\AllowAllMode', @SW_Hide)
EndIf
If $CmdLine[1] = "/Audit" Then
	Local $ps_Cipath = @SystemDir & "\CodeIntegrity\CiPolicies\Active\"
	Local $ps_copy = ' Copy-Item *.cip -Destination ' & $ps_Cipath & ' -Force; Start-Sleep -Seconds 1; (citool.exe -r -json) '
	Local $ps_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    Local $ps_Pid = Run($ps_powershell & $ps_copy, @ScriptDir & '\policies\AuditMode', @SW_Hide)
EndIf
If $CmdLine[1] = "/Enforce" Then
	Local $ps_Cipath = @SystemDir & "\CodeIntegrity\CiPolicies\Active\"
	Local $ps_copy = ' Copy-Item *.cip -Destination ' & $ps_Cipath & ' -Force; Start-Sleep -Seconds 1; (citool.exe -r -json) '
	Local $ps_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    Local $ps_Pid = Run($ps_powershell & $ps_copy, @ScriptDir & '\policies\EnforcedMode', @SW_Hide)
EndIf
