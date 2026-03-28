DllCall("user32.dll", "bool", "SetProcessDpiAwarenessContext", @AutoItX64 ? "int64" : "int", -2)

#include <File.au3>
#include <MsgBoxConstants.au3>
#include <String.au3>
#include "includes\XML.au3"

#NoTrayIcon
#RequireAdmin

#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Tray Helper
#AutoIt3Wrapper_res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Fileversion=6.1.0
#AutoIt3Wrapper_Res_ProductVersion=6.1.0
#AutoIt3Wrapper_Res_ProductName=AppControlTrayHelper
#AutoIt3Wrapper_Outfile_x64=AppControlHelper.exe
#AutoIt3Wrapper_Res_LegalCopyright=@ 2026 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=N
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#AutoIt3Wrapper_UseUpx=N
#AutoIt3Wrapper_Compression=0
#AutoIt3Wrapper_Run_Au3Stripper=y

$sTitle = "AppControlTrayHelper"

Global $isDarkMode = is_app_dark_theme()

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
