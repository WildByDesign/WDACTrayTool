#include-once

#include <ButtonConstants.au3>
#include <Timers.au3>
#include <WinAPI.au3>
#include <WinAPITheme.au3>
#include <WindowsConstants.au3>
#include <WinAPIGdi.au3>

;Store COLOR_* as coloref (consistancy with _WinAPI_GetSysColor())
Global $MSGBOX_BG_TOP = _WinAPI_SwitchColor(0x323232)
Global $MSGBOX_BG_BOTTOM = _WinAPI_SwitchColor(0x202020)
Global $MSGBOX_BG_BUTTON = $MSGBOX_BG_BOTTOM
Global $MSGBOX_TEXT = _WinAPI_SwitchColor(0xFFFFFF)

Global Const $HCBT_MOVESIZE = 0
Global Const $HCBT_MINMAX = 1
Global Const $HCBT_QS = 2
Global Const $HCBT_CREATEWND = 3
Global Const $HCBT_DESTROYWND = 4
Global Const $HCBT_ACTIVATE = 5
Global Const $HCBT_CLICKSKIPPED = 6
Global Const $HCBT_KEYSKIPPED = 7
Global Const $HCBT_SYSCOMMAND = 8
Global Const $HCBT_SETFOCUS = 9

Global Const $DWMSBT_AUTO = 0               ; Default (Auto)
Global Const $DWMSBT_NONE = 1               ; None
Global Const $DWMSBT_MAINWINDOW = 2         ; Mica
Global Const $DWMSBT_TRANSIENTWINDOW = 3    ; Acrylic
Global Const $DWMSBT_TABBEDWINDOW = 4       ; Mica Alt (Tabbed)

Global Const $tagNMCUSTOMDRAWINFO = $tagNMHDR & ";dword DrawStage;handle hdc;" & $tagRECT & ";dword_ptr ItemSpec;uint ItemState;lparam lItemParam;"

Global $g_iRevision = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR")
Global $g_b24H2Plus = False
If @OSBuild >= 26100 And $g_iRevision >= 6899 Then $g_b24H2Plus = True

Global $g_hMsgBoxHook, $g_idTImer
Global $g_UseDarkMode
Global $g_Timeout = 0, $g_hMsgBoxOldProc, $g_hMsgBoxBrush, $g_hMsgBoxBtn = 0, $g_bMsgBoxClosing = False, $g_bNCLButtonDown = False, $g_bMsgBoxInitialized = False
Global $g_hMsgBoxSubProc = DllCallbackRegister("_MsgBoxProc", "lresult", "hwnd;uint;wparam;lparam;uint_ptr;dword_ptr")
Global $g_pMsgBoxSubProc = DllCallbackGetPtr($g_hMsgBoxSubProc)
Global $g_bUseMica = False, $g_iMaterial
Global $g_aButtonText[12]
Global $g_bShowCount = True, $g_bShowUnderline = True

Global Enum $PAM_Default, $PAM_AllowDark, $PAM_ForceDark, $PAM_ForceLight, $PAM_Max

Global $g_iMsgBoxDpi = Round(_WinAPI_GetDPI() / 96, 2)
If @error Then $g_iMsgBoxDpi = 1

OnAutoItExitRegister("_MsgBoxExCleaup")

;From testing, using $PAM_AllowDark will cause _WinAPI_ShouldAppsUseDarkMode() to follow the actual darkmode setting.
;Otherwise you can force _WinAPI_ShouldAppsUseDarkMode to return either way with $PAM_ForceDark & $PAM_ForceLight.
;_WinAPI_SetPreferredAppMode($PAM_AllowDark)

Func _MsgBoxExCleaup()
    DllCallbackFree($g_hMsgBoxSubProc)
EndFunc   ;==>_MsgBoxExCleaup

Func _MsgBoxProc($hWnd, $iMsg, $wParam, $lParam, $iSubClsID, $pData)
    Local Static $hBkColorBrush, $hFooterBrush, $hTimer

    Switch $iMsg

        Case $WM_NCCREATE
            If $g_UseDarkMode Then _WinAPI_DwmSetWindowAttribute($hWnd, $DWMWA_USE_IMMERSIVE_DARK_MODE, True)

        Case $WM_INITDIALOG
            $hBkColorBrush = _WinAPI_CreateSolidBrush($MSGBOX_BG_TOP)
            $hFooterBrush = _WinAPI_CreateSolidBrush($MSGBOX_BG_BOTTOM)

            ; Store array of button text and determine button availability
            For $i = 1 To 11
                $g_aButtonText[$i] = _WinAPI_GetDlgItemText($hWnd, $i)
            Next

            If $g_Timeout Then
                $hTimer = _Timer_SetTimer($hWnd, 1000, "_TimerProc")

                If $g_bShowCount Then
                    Select
                        Case $g_aButtonText[1] And Not $g_aButtonText[2]
                            ; OK button
                            _WinAPI_SetDlgItemText($hWnd, $IDOK, StringFormat("%s [%d]", $g_aButtonText[1], $g_Timeout))

                        Case $g_aButtonText[1] And $g_aButtonText[2]
                            ; OK and Cancel
                            _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

                        Case $g_aButtonText[2] And $g_aButtonText[4]
                            ; Retry and Cancel
                            _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

                        Case $g_aButtonText[6] And $g_aButtonText[7] And $g_aButtonText[2]
                            ; Yes, No and Cancel
                            _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

                        Case $g_aButtonText[2] And $g_aButtonText[10] And $g_aButtonText[11]
                            ; Cancel, Try Again, Continue
                            _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

                        Case $g_aButtonText[3] And $g_aButtonText[4] And $g_aButtonText[5]
                            ; Abort, Retry, and Ignore
                            _WinAPI_SetDlgItemText($hWnd, $IDABORT, StringFormat("%s [%d]", $g_aButtonText[3], $g_Timeout))
                    EndSelect
                EndIf
            EndIf

            If $g_bUseMica And @OSBuild >= 22621 Then
                _WinAPI_DwmSetWindowAttributeEx($hWnd, $DWMWA_SYSTEMBACKDROP_TYPE, $g_iMaterial)
                _WinAPI_DwmExtendFrameIntoClientArea($hWnd, _WinAPI_CreateMargins(-1, -1, -1, -1))
            EndIf

        Case $WM_CTLCOLORSTATIC,  $WM_CTLCOLORDLG
            If $g_UseDarkMode Then
                _WinAPI_SetBkMode($wParam, $TRANSPARENT)
                _WinAPI_SetTextColor($wParam, $MSGBOX_TEXT)
                Return $hBkColorBrush
            EndIf

        ;Case $WM_CTLCOLORBTN
        ;    If $g_UseDarkMode Then Return $hFooterBrush

        Case $WM_PAINT
            If $g_UseDarkMode Then
                Local $tPS = DllStructCreate($tagPAINTSTRUCT)
                Local $hDC = _WinAPI_BeginPaint($hWnd, $tPS)
                Local $tPaintRect = DllStructCreate($tagRECT, DllStructGetPtr($tPS, "rPaint"))
                $tPaintRect.Top = ($tPaintRect.Bottom - (48 * $g_iMsgBoxDpi)) ; this depends on DPI scale
                _WinAPI_FillRect($hDC, $tPaintRect, $hFooterBrush)
                _WinAPI_EndPaint($hWnd, $tPS)
                Return True
            EndIf

        Case $WM_COMMAND
            Local $iNotifCode = _WinAPI_HiWord($wParam)
            Local $iItemId = _WinAPI_LoWord($wParam)
            If (Not $lParam) Or ($iNotifCode = $BN_CLICKED) Then
                Return _Dialog_EndDialog($hWnd, $iItemId)
            EndIf

        Case $WM_DESTROY
            _Timer_KillTimer($hWnd, $hTimer)
            _WinAPI_RemoveWindowSubclass($hWnd, $g_pMsgBoxSubProc, $iSubClsID)
            _WinAPI_SetActiveWindow(_WinAPI_GetParent($hWnd))
            _WinAPI_DeleteObject($hBkColorBrush)
            _WinAPI_DeleteObject($hFooterBrush)
;#cs
        Case $WM_NOTIFY
        Local $tInfo = DllStructCreate($tagNMCUSTOMDRAWINFO, $lParam)
        Local $iMsgBoxDpiPct = $g_iMsgBoxDpi * 100
        If _WinAPI_GetClassName($tInfo.hWndFrom) = "Button" And $tInfo.Code = $NM_CUSTOMDRAW Then
            Local $tRECT = DllStructCreate($tagRECT, DllStructGetPtr($tInfo, "left"))
            Switch $tInfo.DrawStage
			    Case $CDDS_PREPAINT
                    Local $hBrush
                    If BitAND($tInfo.ItemState, $CDIS_HOT) Then
                        $hBrush = _WinAPI_CreateSolidBrush(_WinAPI_ColorAdjustLuma($MSGBOX_BG_BUTTON, 20))
                    EndIf
                    If BitAND($tInfo.ItemState, $CDIS_SELECTED) Then
                        $hBrush = _WinAPI_CreateSolidBrush(_WinAPI_ColorAdjustLuma($MSGBOX_BG_BUTTON, 15))
                    EndIf
                    If BitAND($tInfo.ItemState, $CDIS_DISABLED) Then
                        $hBrush = _WinAPI_CreateSolidBrush(_WinAPI_ColorAdjustLuma($MSGBOX_BG_BUTTON, 5))
                    EndIf
                    If Not BitAND($tInfo.ItemState, $CDIS_HOT) And Not BitAND($tInfo.ItemState, $CDIS_SELECTED) And Not BitAND($tInfo.ItemState, $CDIS_DISABLED) Then
                        $hBrush = _WinAPI_CreateSolidBrush(_WinAPI_ColorAdjustLuma($MSGBOX_BG_BUTTON, 10))
                    EndIf
                    _WinAPI_FillRect($tInfo.hDC, $tRECT, $hBrush)
                    _WinAPI_DeleteObject($hBrush)
                    Return $CDRF_NOTIFYPOSTPAINT
                #cs
                Case $CDDS_POSTPAINT
                    ; calculate InflateRect parameters based on current DPI scale
                    ; for each 25% increase in scale, the Y param seems to negate by 1 (eg. -5, -6, -7, etc.)
                    ; -5 seems best for 100% scale as starting point
                    ; TODO: not sure yet if X needs changes for higher DPI (eg. 200% or higher)
                    Local $iY = -5
                    Local Const $iSteps = Int(($iMsgBoxDpiPct - 100) / 25)
	                $iY = $iY - (1 * $iSteps)
                    _WinAPI_InflateRect($tRECT, -4, $iY)
                    _WinAPI_SetBkMode($tInfo.hDC, $TRANSPARENT)
                    _WinAPI_SetTextColor($tInfo.hDC, $MSGBOX_TEXT)
                    Local $iFlags = BitOR($DT_CENTER, $DT_VCENTER, $DT_NOCLIP)
                    Switch $g_bShowUnderline
                        Case True
                            _WinAPI_DrawText($tInfo.hDC, _WinAPI_GetDlgItemText($hWnd, $tInfo.IDFrom), $tRECT, $iFlags)
                        Case False
                            _WinAPI_DrawText($tInfo.hDC, StringReplace(_WinAPI_GetDlgItemText($hWnd, $tInfo.IDFrom), "&", ""), $tRECT, $iFlags)
                    EndSwitch
                #ce
            EndSwitch
        EndIf
    ;#ce

    EndSwitch
    Return _WinAPI_DefSubclassProc($hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>_MsgBoxProc

Func _TimerProc($hWnd, $iMsg, $wParam, $lParam)
    If Not _WinAPI_IsWindow($hWnd) Then Return
    $g_Timeout -= 1

    If $g_bShowCount Then
        Select
            Case $g_aButtonText[1] And Not $g_aButtonText[2]
                ; OK button
                _WinAPI_SetDlgItemText($hWnd, $IDOK + 1, StringFormat("%s [%d]", $g_aButtonText[1], $g_Timeout))

            Case $g_aButtonText[1] And $g_aButtonText[2]
                ; OK and Cancel
                _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

            Case $g_aButtonText[2] And $g_aButtonText[4]
                ; Retry and Cancel
                _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

            Case $g_aButtonText[6] And $g_aButtonText[7] And $g_aButtonText[2]
                ; Yes, No and Cancel
                _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

            Case $g_aButtonText[2] And $g_aButtonText[10] And $g_aButtonText[11]
                ; Cancel, Try Again, Continue
                _WinAPI_SetDlgItemText($hWnd, $IDCANCEL, StringFormat("%s [%d]", $g_aButtonText[2], $g_Timeout))

            Case $g_aButtonText[3] And $g_aButtonText[4] And $g_aButtonText[5]
                ; Abort, Retry, and Ignore
                _WinAPI_SetDlgItemText($hWnd, $IDABORT, StringFormat("%s [%d]", $g_aButtonText[3], $g_Timeout))
        EndSelect
    EndIf
EndFunc   ;==>_TimerProc

Func _CBTHookProc($nCode, $wParam, $lParam)
    Local Const $hWnd = HWnd($wParam)
    If $nCode < 0 Then Return _WinAPI_CallNextHookEx($g_hMsgBoxHook, $nCode, $wParam, $lParam)
    Switch $nCode
        Case $HCBT_CREATEWND
            Switch _WinAPI_GetClassName($hWnd)
                Case "#32770"
                    _WinAPI_SetWindowSubclass($hWnd, $g_pMsgBoxSubProc, 1000)
                Case "Button"
                    ;If $g_b24H2Plus Then
                    ;    If $g_UseDarkMode Then _WinAPI_SetWindowTheme($hWnd, "DarkMode_DarkTheme")
                    ;Else
                        If $g_UseDarkMode Then _WinAPI_SetWindowTheme($hWnd, "DarkMode_Explorer")
                    ;EndIf
            EndSwitch
    EndSwitch
    Return _WinAPI_CallNextHookEx($g_hMsgBoxHook, $nCode, $wParam, $lParam)
EndFunc   ;==>_CBTHookProc

Func MsgBoxEx($iFlag, $sTitle, $sText, $iTimeout = 0, $hParentHWND = "")
    ; check for any DPI changes
    $g_iMsgBoxDpi = Round(_WinAPI_GetDPI() / 96, 2)
    If @error Then $g_iMsgBoxDpi = 1
    $g_UseDarkMode = _WinAPI_ShouldAppsUseDarkMode()
    $g_bMsgBoxInitialized = False
    If $iFlag = Default Then $iFlag = BitOR($MB_TOPMOST, $MB_ICONINFORMATION)
    $g_Timeout = $iTimeout
    Local $hMsgProc = DllCallbackRegister("_CBTHookProc", "int", "uint;wparam;lparam")
    Local Const $hThreadID = _WinAPI_GetCurrentThreadId()
    $g_hMsgBoxHook = _WinAPI_SetWindowsHookEx($WH_CBT, DllCallbackGetPtr($hMsgProc), Null, $hThreadID)
    If $sTitle = Default Then $sTitle = "Information"
    Local Const $iReturn = MsgBox($iFlag, $sTitle, $sText, $iTimeout, $hParentHWND)
    If $g_hMsgBoxHook Then _WinAPI_UnhookWindowsHookEx($g_hMsgBoxHook)
    DllCallbackFree($hMsgProc)
    Return $iReturn
EndFunc   ;==>MsgBoxEx

; #FUNCTION# ====================================================================================================================
; Name ..........: _MsgBoxExSetColors
; Description ...: Sets custom colors and materials
; Syntax ........: _MsgBoxExSetColors($iBgColorTop, $iBgColorBottom, $iBgColorButton, $iTextColor, $iMaterial)
; Parameters ....: $iBgColorTop       - [optional] 0xRRGGBB (RGB value). Default is $MSGBOX_BG_TOP.
;                  $iBgColorBottom    - [optional] 0xRRGGBB (RGB value). Default is $MSGBOX_BG_BOTTOM.
;                  $iBgColorButton 	  - [optional] 0xRRGGBB (RGB value). Default is $MSGBOX_BG_BUTTON.
;                  $iTextColor 	      - [optional] 0xRRGGBB (RGB value). Default is $MSGBOX_TEXT.
;                  $iMaterial 	      - [optional] Integer. Default is blank. Options below:
;                                     - $DWMSBT_AUTO               ; Default (Auto)
;                                     - $DWMSBT_NONE               ; None
;                                     - $DWMSBT_MAINWINDOW         ; Mica
;                                     - $DWMSBT_TRANSIENTWINDOW    ; Acrylic
;                                     - $DWMSBT_TABBEDWINDOW       ; Mica Alt (Tabbed)
; Return values .: None
; Author ........: WildByDesign
; Remarks .......: Windows 11 material effects look best with background values between 0x000000 and 0x202020
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _MsgBoxExSetColors($iBgColorTop = Default, $iBgColorBottom = Default, $iBgColorButton = Default, $iTextColor = Default, $iMaterial = Default)
    If $iBgColorTop <> Default Then $MSGBOX_BG_TOP = _WinAPI_SwitchColor($iBgColorTop)
    If $iBgColorBottom <> Default Then $MSGBOX_BG_BOTTOM = _WinAPI_SwitchColor($iBgColorBottom)
    If $iBgColorButton <> Default Then $MSGBOX_BG_BUTTON = _WinAPI_SwitchColor($iBgColorButton)
    If $iTextColor <> Default Then $MSGBOX_TEXT = _WinAPI_SwitchColor($iTextColor)
    If $iMaterial <> Default And $iMaterial <> "" And @OSBuild >= 22621 Then
        $g_bUseMica = True
        $g_iMaterial = $iMaterial
    EndIf
EndFunc   ;==>_MsgBoxExSetColors

; #FUNCTION# ====================================================================================================================
; Name ..........: _MsgBoxExSetOptions
; Description ...: Sets custom settings.
; Syntax ........: _MsgBoxExSetOptions($bShowCount, $bShowUnderline)
; Parameters ....: $bShowCount        - [optional] Boolean. Default is True. Show countdown timer on buttons.
;                  $bShowUnderline    - [optional] Boolean. Default is True. Show underline for letters with ampersand on buttons.
; Return values .: None
; Author ........: WildByDesign
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _MsgBoxExSetOptions($bShowCount = Default, $bShowUnderline = Default)
    If $bShowCount <> Default Then $g_bShowCount = $bShowCount
    If $bShowUnderline <> Default Then $g_bShowUnderline = $bShowUnderline
EndFunc   ;==>_MsgBoxExSetOptions

Func _Dialog_EndDialog($hWnd, $iReturn)
    Local $aCall = DllCall("User32.dll", "bool", "EndDialog", "hwnd", $hWnd, "int_ptr", $iReturn)
    If @error Then Return SetError(@error, @extended, False)
    Return $aCall[0]
EndFunc   ;==>_Dialog_EndDialog

Func _WinAPI_ShouldAppsUseDarkMode()
    Local $aResult = DllCall("UxTheme.dll", "bool", 132)
    If @error Then Return SetError(1, 0, False)
    Return $aResult[0]
EndFunc   ;==>_WinAPI_ShouldAppsUseDarkMode

Func _WinAPI_SetPreferredAppMode($iMode)
    Local $aResult = DllCall("UxTheme.dll", "ptr", 135, "int", $iMode)
    If @error Then Return SetError(1, 0, False)
    Return $aResult[0]
EndFunc   ;==>_WinAPI_SetPreferredAppMode

Func _WinAPI_GetDPI($hWnd = 0) ; UEZ
    $hWnd = Not $hWnd ? _WinAPI_GetDesktopWindow() : $hWnd
    Local Const $hDC = _WinAPI_GetDC($hWnd)
    If @error Then Return SetError(1, 0, 0)
    Local Const $iDPI = _WinAPI_GetDeviceCaps($hDC, $LOGPIXELSX)
    If @error Or Not $iDPI Then
        _WinAPI_ReleaseDC($hWnd, $hDC)
        Return SetError(2, 0, 0)
    EndIf
    _WinAPI_ReleaseDC($hWnd, $hDC)
    Return $iDPI
EndFunc   ;==>_WinAPI_GetDPI

Func _WinAPI_DwmSetWindowAttributeEx($hWnd, $iAttribute, $iData)
	Switch $iAttribute
		Case 2, 3, 4, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, $DWMWA_USE_IMMERSIVE_DARK_MODE, 33, 34, 35, 36, 37, 38, 39, 40

		Case Else
			Return SetError(1, 0, 0)
	EndSwitch

	Local $aCall = DllCall('dwmapi.dll', 'long', 'DwmSetWindowAttribute', 'hwnd', $hWnd, 'dword', $iAttribute, _
			'dword*', $iData, 'dword', 4)
	If @error Then Return SetError(@error + 10, @extended, 0)
	If $aCall[0] Then Return SetError(10, $aCall[0], 0)

	Return 1
EndFunc   ;==>_WinAPI_DwmSetWindowAttributeEx

Func _WinAPI_GetDlgItemText($hDlg, $iDlgItem)
    Local $hCtrl, $iBuffLen, $tBuff, $sResult = ""
    $hCtrl = _WinAPI_GetDlgItem($hDlg, $iDlgItem)
    If Not @error Then $iBuffLen = _SendMessage($hCtrl, $WM_GETTEXTLENGTH)
    If Not @error Then
        $tBuff = DllStructCreate(StringFormat("wchar[%d]", $iBuffLen + 1))
        _SendMessage($hCtrl, $WM_GETTEXT, $iBuffLen + 1, DllStructGetPtr($tBuff))
        If Not @error Then $sResult = DllStructGetData($tBuff, 1)
    EndIf
    Return SetError(@error, @extended, $sResult)
EndFunc   ;==>_WinAPI_GetDlgItemText

Func _WinAPI_SetDlgItemText($hDlg, $nIDDlgItem, $lpString) ;https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setdlgitemtextw
    Local $aRet = DllCall("user32.dll", "int", "SetDlgItemText", "hwnd", $hDlg, "int", $nIDDlgItem, "str", $lpString)
    If @error Then Return SetError(@error, @extended, 0)
    Return $aRet[0]
EndFunc   ;==>_WinAPI_SetDlgItemText
