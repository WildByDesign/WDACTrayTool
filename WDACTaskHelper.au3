
#NoTrayIcon

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=WDAC-color.ico
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=WDAC Tray Tool
#AutoIt3Wrapper_Res_Fileversion=3.0.0.0
#AutoIt3Wrapper_Res_ProductVersion=3.0.0
#AutoIt3Wrapper_Res_ProductName=WDACTrayTool
#AutoIt3Wrapper_Res_LegalCopyright=@ 2024 WildByDesign
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_HiDpi=P
#AutoIt3Wrapper_Res_Icon_Add=WDAC-color.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
$sTitle = "WDACTaskHelper"

If $CmdLine[0] = 0 Then Exit MsgBox(16, $sTitle, "No parameters passed!")

If $CmdLine[1] = "install" Then
; install special shortcut with AppID
$cmdInstall1 = ' -install "App Control Tray Tool.lnk" '
$cmdInstall2 = @ScriptDir & '\WDACTrayTool.exe '
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
$cmdRefresh = ' -t "Your WDAC policies have been refreshed." -m "Code Integrity policy refresh finished successfully." -appID "App Control Tray Tool" -p toast-refresh.png -silent -d short -b Dismiss'
Run(@ScriptDir & "\toasts\snoretoast.exe" & $cmdRefresh, @ScriptDir & "\toasts", @SW_HIDE)
EndIf
