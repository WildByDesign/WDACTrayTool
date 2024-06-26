﻿#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=WDAC.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=WDAC Tray Tool
#AutoIt3Wrapper_Res_Fileversion=1.5.0.0
#AutoIt3Wrapper_Res_ProductVersion=1.5.0
#AutoIt3Wrapper_Res_ProductName=WDACTrayTool
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <APIRegConstants.au3>
#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <EventLog.au3>
#include <WinAPIFiles.au3>
#include <WinAPIDiag.au3>
#include <WinAPIReg.au3>
#include <File.au3>
#include <EventLog.au3>
#include <Date.au3>
#include <Crypt.au3>


Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)

Global $softName = "WDAC Tray Tool"
Global $trayIcon = "WDAC.ico"
Global $idRegTitleKey = "WDAC Tray Tool"
Global $createdBy = "WildByDesign"

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
Local $idWDACWizard = TrayCreateItem("WDAC Wizard")
Local $idMsinfo32 = TrayCreateItem("System Information")
Local $idCiTool = TrayCreateItem("CiTool Status")
Local $idViewLogs = TrayCreateItem("View Logs")
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
TrayCreateItem("")
Local $idAbout = TrayCreateItem("About")
Local $idExit = TrayCreateItem("Exit")

; Tray Icon
TraySetIcon($trayIcon)

Func WDACWizard()
    Run(@ComSpec & " /c " & 'Explorer.Exe Shell:AppsFolder\Microsoft.WDAC.WDACWizard_8wekyb3d8bbwe!WDACWizard', "", @SW_HIDE)
EndFunc
Func Msinfo32()
    Run("C:\Windows\System32\msinfo32.exe", "", @SW_SHOWMAXIMIZED)    
EndFunc
Func CiTool()
    Run(@ComSpec & " /c " & '.\sudo.exe run --new-window C:\Windows\System32\CiTool.exe --list-policies', ".\bin", @SW_HIDE)
EndFunc
Func ViewLogs()
    Run(@ComSpec & " /c " & '.\sudo.exe run eventvwr /c:"Microsoft-Windows-CodeIntegrity/Operational"', ".\bin", @SW_HIDE)
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

TraySetToolTip($softName)

While 1
    Switch TrayGetMsg()
        Case $idWDACWizard
            WDACWizard()
        Case $idMsinfo32
            Msinfo32()
        Case $idCiTool
            CiTool()
        Case $idViewLogs
            ViewLogs()
        Case $idAllowAll
            AllowAll()
        Case $idAudit
            Audit()
        Case $idEnforce
            Enforce()
        Case $idStartItem
            StartWithWindows()
        Case $idAbout
            MsgBox($MB_SYSTEMMODAL, "About WDAC Tray Tool", "Version: " & 1.5 & @CRLF & _
                            "Created by: " & $createdBy & @CRLF & @CRLF & _
                            "This program is free and open source.")
        Case $idExit
        ExitLoop
    EndSwitch

WEnd
