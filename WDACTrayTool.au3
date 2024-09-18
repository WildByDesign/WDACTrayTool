#Region ; *** Dynamically added Include files ***
#include <MsgBoxConstants.au3>                               ; added:08/18/24 07:36:26
#include <TrayConstants.au3>                                 ; added:08/18/24 07:36:26
#EndRegion ; *** Dynamically added Include files ***
#include <ExtMsgBox.au3>
#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=WDAC.ico
#AutoIt3Wrapper_Res_Icon_Add=WDAC.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=WDAC Tray Tool
#AutoIt3Wrapper_Res_Fileversion=2.5.0.0
#AutoIt3Wrapper_Res_ProductVersion=2.5.0
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
Global $programVersion = "2.5"

; Set Current Working Directory
Local $setCWD = True
If $setCWD Then FileChangeDir(@ScriptDir)

; Menu
Local $idPStest = TrayCreateItem("App Control Policy Status")
Local $idCiTool = TrayCreateItem("CiTool Status (-lp)")
TrayCreateItem("")
Local $idWDACWizard = TrayCreateItem("WDAC Policy Wizard")
Local $idMsinfo32 = TrayCreateItem("System Information")
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
    Run(@ScriptDir & "\WDACTrayHelper.exe /CiTool")
EndFunc
Func LogsCI()
	Run(@ScriptDir & "\WDACTrayHelper.exe /LogsCI")
EndFunc
Func LogsScript()
	Run(@ScriptDir & "\WDACTrayHelper.exe /LogsScript")
EndFunc
Func AllowAll()
    Run(@ScriptDir & "\WDACTrayHelper.exe /AllowAll")
EndFunc
Func Audit()
    Run(@ScriptDir & "\WDACTrayHelper.exe /Audit")
EndFunc
Func Enforce()
    Run(@ScriptDir & "\WDACTrayHelper.exe /Enforce")
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
	Run(@ScriptDir & "\WDACTrayHelper.exe /status")
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
            _ExtMsgBoxSet(1, 4, -1, -1, -1, "Consolas", 800, 800)
			_ExtMsgBox (0 & ";" & @ScriptDir & "\WDACTrayTool.exe", 0, "About WDAC Tray Tool", "Version: " & $programVersion & @CRLF & _
                            "Created by: " & $createdBy & @CRLF & @CRLF & _
                            "This program is free and open source." & @CRLF)
        Case $idExit
        ExitLoop
    EndSwitch

WEnd
