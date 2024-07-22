#Region ; *** Dynamically added Include files ***
#include <MsgBoxConstants.au3>                               ; added:05/19/24 07:20:09
#include <TrayConstants.au3>                                 ; added:05/19/24 07:20:09
#include <File.au3>                                          ; added:07/20/24 07:01:29
#include <AutoItConstants.au3>                               ; added:07/20/24 21:22:30
#include <XML.au3>                                           ; added:07/20/24 21:22:30
#EndRegion ; *** Dynamically added Include files ***
;#include <WinAPIFiles.au3>
;#include <FileConstants.au3>
#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=WDAC.ico
#AutoIt3Wrapper_Res_Icon_Add=WDAC.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=WDAC Tray Tool
#AutoIt3Wrapper_Res_Fileversion=2.0.0.0
#AutoIt3Wrapper_Res_ProductVersion=2.0.0
#AutoIt3Wrapper_Res_ProductName=WDACTrayTool
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)

Global $softName = "WDAC Tray Tool"
Global $trayIcon = "WDAC.ico"
Global $idRegTitleKey = "WDAC Tray Tool"
Global $createdBy = "WildByDesign"
Global $programVersion = "2.0"

; Set Current Working Directory
Local $setCWD = True
If $setCWD Then FileChangeDir(@ScriptDir)

; Create Directory for binaries
Local $CreateDir = True
If $CreateDir Then DirCreate(@ScriptDir & ".\bin")

; Use FileInstall to bundle binaries
Local $sudoFileInstall = True
Local $RefreshPolicyFileInstall = True

; Use FileInstall to extract binaries
If $sudoFileInstall Then FileInstall(".\embed\sudo.exe", @ScriptDir & ".\bin\sudo.exe")
If $RefreshPolicyFileInstall Then FileInstall(".\embed\RefreshPolicy.exe", @ScriptDir & ".\bin\RefreshPolicy.exe")

; Menu
Local $idPStest = TrayCreateItem("App Control Policy Status")
TrayCreateItem("")
Local $idWDACWizard = TrayCreateItem("WDAC Policy Wizard")
Local $idMsinfo32 = TrayCreateItem("System Information")
Local $idCiTool = TrayCreateItem("CiTool Status")
TrayCreateItem("")
Local $idViewLogs = TrayCreateMenu("View Logs")
Local $idLogsCI = TrayCreateItem("Code Integrity", $idViewLogs)
Local $idLogsScript = TrayCreateItem("MSI and Script", $idViewLogs)
TrayCreateItem("")
Local $idStatus = TrayCreateMenu("Change Policy")
Local $idAllowAll = TrayCreateItem("Allow All", $idStatus)
Local $idAudit = TrayCreateItem("Audit Mode", $idStatus)
Local $idEnforce = TrayCreateItem("Enforced Mode", $idStatus)
TrayCreateItem("")
Local $idStartItem = TrayCreateItem("Start at Logon")
$idRegRead = RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", $idRegTitleKey)
If $idRegRead <> '' Then TrayItemSetState ($idStartItem, $TRAY_CHECKED)
TrayItemSetOnEvent ($idStartItem, "StartWithWindows")
;TrayCreateItem("")
Local $idAbout = TrayCreateItem("About")
Local $idExit = TrayCreateItem("Exit")

; Tray Icon - 201 is the resource number of embedded icon
TraySetIcon(@ScriptFullPath, 201)

Func WDACWizard()
    Run(@ComSpec & " /c " & 'Explorer.Exe Shell:AppsFolder\Microsoft.WDAC.WDACWizard_8wekyb3d8bbwe!WDACWizard', "", @SW_HIDE)
EndFunc
Func Msinfo32()
    Run("C:\Windows\System32\msinfo32.exe", "", @SW_SHOWMAXIMIZED)
EndFunc
Func CiTool()
    Run(@ComSpec & " /c " & '.\sudo.exe run --new-window C:\Windows\System32\CiTool.exe --list-policies', ".\bin", @SW_HIDE)
EndFunc
Func LogsCI()
	$oXML=ObjCreate("Microsoft.XMLDOM")
	$stest = FileRead(@AppDataDir & "\Microsoft\MMC\eventvwr")
	$oXML.LoadXML($stest)

	If Not _XML_NodeExists($oXML, "//DynamicPath") Then
			Run(@ComSpec & " /c " & '.\sudo.exe run eventvwr /c:"Microsoft-Windows-CodeIntegrity/Operational" /f:"*[System/EventID=3004] or *[System/EventID=3033] or *[System/EventID=3034] or *[System/EventID=3076] or *[System/EventID=3077] or *[System/EventID=3089]"', ".\bin", @SW_HIDE)
	Else
	$oXMLChild1 = $oXML.selectSingleNode( "//DynamicPath" )
	$parent = $oXMLChild1.ParentNode

	$test = $parent.removeChild ( $oXMLChild1 )

	$oXML.Save(@AppDataDir & "\Microsoft\MMC\eventvwr")
	Sleep(500)
	Run(@ComSpec & " /c " & '.\sudo.exe run eventvwr /c:"Microsoft-Windows-CodeIntegrity/Operational" /f:"*[System/EventID=3004] or *[System/EventID=3033] or *[System/EventID=3034] or *[System/EventID=3076] or *[System/EventID=3077] or *[System/EventID=3089]"', ".\bin", @SW_HIDE)
	EndIf
EndFunc
Func LogsScript()
	$oXML=ObjCreate("Microsoft.XMLDOM")
	$stest = FileRead(@AppDataDir & "\Microsoft\MMC\eventvwr")
	$oXML.LoadXML($stest)

	If Not _XML_NodeExists($oXML, "//DynamicPath") Then
			Run(@ComSpec & " /c " & '.\sudo.exe run eventvwr /c:"Microsoft-Windows-AppLocker/MSI and Script" /f:"*[System/EventID=8028] or *[System/EventID=8029] or *[System/EventID=8036] or *[System/EventID=8037] or *[System/EventID=8038] or *[System/EventID=8039] or *[System/EventID=8040]"', ".\bin", @SW_HIDE)
	Else
	$oXMLChild1 = $oXML.selectSingleNode( "//DynamicPath" )
	$parent = $oXMLChild1.ParentNode

	$test = $parent.removeChild ( $oXMLChild1 )

	$oXML.Save(@AppDataDir & "\Microsoft\MMC\eventvwr")
	Sleep(500)
	Run(@ComSpec & " /c " & '.\sudo.exe run eventvwr /c:"Microsoft-Windows-AppLocker/MSI and Script" /f:"*[System/EventID=8028] or *[System/EventID=8029] or *[System/EventID=8036] or *[System/EventID=8037] or *[System/EventID=8038] or *[System/EventID=8039] or *[System/EventID=8040]"', ".\bin", @SW_HIDE)
	EndIf
EndFunc
Func AllowAll()
    RunWait(@ComSpec & " /c " & '.\sudo.exe run --chdir ..\policies\AllowAllMode copy /y *.cip C:\Windows\System32\CodeIntegrity\CiPolicies\Active\', ".\bin", @SW_HIDE)
    Sleep(500)
    Run(@ComSpec & " /c " & '.\RefreshPolicy.exe', ".\bin", @SW_HIDE)
EndFunc
Func Audit()
    RunWait(@ComSpec & " /c " & '.\sudo.exe run --chdir ..\policies\AuditMode copy /y *.cip C:\Windows\System32\CodeIntegrity\CiPolicies\Active\', ".\bin", @SW_HIDE)
    Sleep(500)
    Run(@ComSpec & " /c " & '.\RefreshPolicy.exe', ".\bin", @SW_HIDE)
EndFunc
Func Enforce()
    RunWait(@ComSpec & " /c " & '.\sudo.exe run --chdir ..\policies\EnforcedMode copy /y *.cip C:\Windows\System32\CodeIntegrity\CiPolicies\Active\', ".\bin", @SW_HIDE)
    Sleep(500)
    Run(@ComSpec & " /c " & '.\RefreshPolicy.exe', ".\bin", @SW_HIDE)
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
	Local $o_CmdString = ' Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity*'
    Local $o_powershell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    Local $o_Pid = Run($o_powershell & $o_CmdString, '', @SW_Hide, $STDERR_CHILD + $STDOUT_CHILD)
    ProcessWaitClose($o_Pid)

	Local Const $sFilePath = @TempDir & '\ci-policy.txt'
	Local $hFileOpen = FileOpen($sFilePath, 2)
	Local $o_Output = StdoutRead($o_Pid)
	_FileWriteLog($sFilePath, $o_Output)
	While 1
    $sFileReadLine3 = FileReadLine($hFileOpen, 3)
	Local $hFileOpen = FileOpen($sFilePath, 0)
	$sFileReadLine4 = FileReadLine($hFileOpen, 4)
	ExitLoop
    ;If @error = -1 Then ExitLoop
	WEnd
	FileClose($hFileOpen)
	FileDelete($sFilePath)
	$FilterLine3 = StringRight($sFileReadLine3, 1)
	$FilterLine4 = StringRight($sFileReadLine4, 1)
	; FilterName = 1 means idDisplayName contains an @ symbol
				If $FilterLine3 = 2 Then
				$StatusLine3 = "App Control policy: Enforced"
				Else
				$StatusLine3 = "App Control policy: Audit"
				EndIf
	; FilterName = 1 means idDisplayName contains an @ symbol
				If $FilterLine4 = 2 Then
				$StatusLine4 = "App Control user mode policy: Enforced"
				Else
				$StatusLine4 = "App Control user mode policy: Audit"
				EndIf
	Local $sText = $StatusLine3
	Local $sText2 = $StatusLine4
	MsgBox($MB_ICONNONE, "App Control Policy Status", $StatusLine3 & @CRLF & @CRLF & $StatusLine4)
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
		Case $idAllowAll
            AllowAll()
        Case $idAudit
            Audit()
        Case $idEnforce
            Enforce()
        Case $idPStest
            PStest()
		Case $idStartItem
            StartWithWindows()
		Case $idAbout
            MsgBox($MB_ICONNONE, "About WDAC Tray Tool", "Version: " & $programVersion & @CRLF & _
                            "Created by: " & $createdBy & @CRLF & @CRLF & _
                            "This program is free and open source.")
        Case $idExit
        ExitLoop
    EndSwitch

WEnd
