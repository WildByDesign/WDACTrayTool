#Region
#include <MsgBoxConstants.au3>
#include <String.au3>
#include <Array.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#EndRegion

#NoTrayIcon

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AppControl.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=App Control Tray Tool
#AutoIt3Wrapper_Res_Fileversion=3.2.0.0
#AutoIt3Wrapper_Res_ProductVersion=3.2.0
#AutoIt3Wrapper_Res_ProductName=AppControlTrayTool
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=AppControl.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
$sTitle = "AppControlTaskHelper"

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
	$mFile = FileOpenDialog("Select XML Policy File for Conversion", @ScriptDir & "\policies\", "Policy Files (*.xml)", 1)
	If @error Then
		Exit
	Else

	Local $sFileRead = FileRead($mFile)
	$aExtract = _StringBetween($sFileRead, "<PolicyID>", "</PolicyID>")

	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	Local $aPathSplit = _PathSplit($mFile, $sDrive, $sDir, $sFileName, $sExtension)

	Local $xmlfiledir = $sDrive & $sDir
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

	Local $binarynamechosen = $sFileSaveDialog

	Local $quote = "'"
	Local $cmd1 = ' ConvertFrom-CIPolicy -XmlFilePath '
	Local $cmd2 = $mFile
	Local $cmd3 = ' -BinaryFilePath '
	Local $cmd4 = $binarynamechosen
	Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
	Local $o_Pid = Run($o_powershell & $cmd1 & $quote & $cmd2 & $quote & $cmd3 & $quote & $cmd4 & $quote, @ScriptDir & '\scripts', @SW_Hide)
EndIf
