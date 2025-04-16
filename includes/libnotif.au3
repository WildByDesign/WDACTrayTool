#include-once
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>
#include <MsgBoxConstants.au3>

#include "StringSize.au3"

; Notification position
Enum _
	$LIBNOTIF_POS_LR, _
	$LIBNOTIF_POS_LL, _
	$LIBNOTIF_POS_UR, _
	$LIBNOTIF_POS_UL

; Animations style (passed to AnimateWindow)
Global Const $AW_ACTIVATE = 0x00020000
Global Const $AW_BLEND = 0x00080000
Global Const $AW_CENTER = 0x00000010
Global Const $AW_HIDE = 0x00010000
Global Const $AW_HOR_POSITIVE = 0x00000001
Global Const $AW_HOR_NEGATIVE = 0x00000002
Global Const $AW_SLIDE = 0x00040000
Global Const $AW_VER_POSITIVE = 0x00000004
Global Const $AW_VER_NEGATIVE = 0x00000008

Enum _ ; _LibNotif_Config - $iCfgOpt
	$LIBNOTIF_CFG_TIMEOUT, _        ; default: 0 ms (no timeout)
	$LIBNOTIF_CFG_CLOSEBUTTON, _    ; default: [*]true
	$LIBNOTIF_CFG_POSITION, _       ; default: lower-right corner
	$LIBNOTIF_CFG_MINTOASTWIDTH, _  ; default: 0 (pixels)
	$LIBNOTIF_CFG_MAXTEXTWIDTH, _   ; default: @DesktopWidth / 3 (pixels)
	$LIBNOTIF_CFG_TEXTMARGIN, _     ; default: 5 (pixels)
	$LIBNOTIF_CFG_BUTTONSIZE, _     ; default: 15 (pixels)
	$LIBNOTIF_CFG_BKCOLOR, _        ; default: [*]system default (passed to GUISetBkColor)
	$LIBNOTIF_CFG_TEXTCOLOR, _      ; default: [*]system default (passed to GUICtrlSetColor)
	$LIBNOTIF_CFG_TEXTSIZE, _       ; default: [*]system default (passed to GUICtrlSetFont)
	$LIBNOTIF_CFG_TEXTWEIGHT, _     ; default: [*]system default (passed to GUICtrlSetFont)
	$LIBNOTIF_CFG_TEXTATTRIBUTES, _ ; default: [*]system default (passed to GUICtrlSetFont)
	$LIBNOTIF_CFG_TEXTFONTNAME, _   ; default: [*]system default (passed to GUICtrlSetFont)
	$LIBNOTIF_CFG_SHOW_ANIMSTYLE, _ ; default: $AW_SLIDE + $AW_HOR_NEGATIVE (0 to deactivate)
	$LIBNOTIF_CFG_SHOW_ANIMTIME, _  ; default: 250
	$LIBNOTIF_CFG_HIDE_ANIMSTYLE, _ ; default: $AW_BLEND + $AW_HIDE (0 to deactivate)
	$LIBNOTIF_CFG_HIDE_ANIMTIME     ; default: 200

; [*] options that need a call to _LibNotif_Init to be validated

; ===============================================================================================================================
; internal globals

; default options
Global $__gLibNotif_aDefaultCfg[17] = [0, True, $LIBNOTIF_POS_LR, 0, @DesktopWidth / 3, 5, 15, Default, Default, 9, 400, 0, "", $AW_SLIDE + $AW_HOR_NEGATIVE, 250, $AW_BLEND + $AW_HIDE, 200]

; set default font name and size in _aDefaultCfg
__libNotif_GetDefaultFont()

; config values used by the library
Global $__gLibNotif_aCfg[17]

For $i = 0 To 16
	$__gLibNotif_aCfg[$i] = $__gLibNotif_aDefaultCfg[$i]
Next

;
Global $__gLibNotif_hWnd = -1
Global $__gLibNotif_iTextLabel = -1, $__gLibNotif_hTextLabel = -1
Global $__gLibNotif_iCloseButton = -1, $__gLibNotif_hCloseButton = 1
Global $__gLibNotif_iIsVisible = 0

; ===============================================================================================================================

; Configure the library. Must call before _LibNotif_Init for [*] config options.
; Possible config options ($iCfgOpt) are the $LIBNOTIF_CFG_XXX Enums.
;
Func _LibNotif_Config($iCfgOpt, $iCfgValue = Default)
	If $iCfgOpt < 0 Or $iCfgOpt >= UBound($__gLibNotif_aCfg) Then Return SetError(-1, 0, 0)
	If $iCfgValue = Default Then
		$__gLibNotif_aCfg[$iCfgOpt] = $__gLibNotif_aDefaultCfg[$iCfgOpt]
	Else
		$__gLibNotif_aCfg[$iCfgOpt] = $iCfgValue
	EndIf
	Return 1
EndFunc

; Init the library. You must call this befor starting to display notifications.
; If you modifie any config option that has [*], you must re-call this function.
;
Func _LibNotif_Init()
	; re-init global variables
	If $__gLibNotif_hWnd <> -1 Then GUIDelete($__gLibNotif_hWnd)
	$__gLibNotif_hWnd = -1
	$__gLibNotif_iTextLabel = -1
	$__gLibNotif_hTextLabel = -1
	$__gLibNotif_iCloseButton = -1
	$__gLibNotif_hCloseButton = -1
	$__gLibNotif_iIsVisible = 0
	; ---
	; create the window
	Local $iOldOpt = Opt("GUIOnEventMode", 1) ; TODO: is it really usefull?
	; ---
	$__gLibNotif_hWnd = GUICreate("LibNotifWindow", 50, 50, -100, -100, $WS_POPUPWINDOW, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
	If $__gLibNotif_aCfg[$LIBNOTIF_CFG_BKCOLOR] <> Default Then
		GUISetBkColor($__gLibNotif_aCfg[$LIBNOTIF_CFG_BKCOLOR], $__gLibNotif_hWnd)
	EndIf
	If $__gLibNotif_aCfg[$LIBNOTIF_CFG_CLOSEBUTTON] Then GUIRegisterMsg($WM_COMMAND, "__libNotif_WM_Handler")
	; ---
	Opt("GUIOnEventMode", $iOldOpt)
	; ---
	; create text label
	$__gLibNotif_iTextLabel = GUICtrlCreateLabel("", 10, 10, 30, 30, -1, $SS_NOTIFY)
	$__gLibNotif_hTextLabel = GUICtrlGetHandle($__gLibNotif_iTextLabel)
	GUICtrlSetFont($__gLibNotif_iTextLabel, $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTSIZE], $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTWEIGHT], $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTATTRIBUTES], $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTFONTNAME])
	If $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTCOLOR] <> Default Then
		GUICtrlSetColor($__gLibNotif_iTextLabel, $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTCOLOR])
	EndIf
	; ---
	; create the close button if activated in configuration
	If $__gLibNotif_aCfg[$LIBNOTIF_CFG_CLOSEBUTTON] Then
		$__gLibNotif_iCloseButton = GUICtrlCreateButton("x", 50 - 20, 5, 15, 15, $BS_CENTER + $BS_FLAT)
		GUICtrlSetOnEvent($__gLibNotif_iCloseButton, "ClosePressed")
		$__gLibNotif_hCloseButton = GUICtrlGetHandle($__gLibNotif_iCloseButton)
		DllCall("UxTheme.dll", "int", "SetWindowTheme", "hwnd", $__gLibNotif_hCloseButton, "wstr", 0, "wstr", 0)
	EndIf
EndFunc


Func ClosePressed()
	MsgBox($MB_SYSTEMMODAL, "Close Pressed", "Close pressed yay!")
EndFunc


; Display a notification.
; Set $iAppend to append $sText to the texte already displayed in the notification
; Set $iHideFirst to hide any visible notification before displaying the new one
;
Func _LibNotif_Notif($sText, $iAppend = False, $iSeparator = @CRLF)
	If $iAppend Then $sText = GUICtrlRead($__gLibNotif_iTextLabel) & $iSeparator & $sText
	; ---
	; get text size
	Local $aTextSize = _StringSize($sText, $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTSIZE], $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTWEIGHT], $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTATTRIBUTES], $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTFONTNAME], $__gLibNotif_aCfg[$LIBNOTIF_CFG_MAXTEXTWIDTH])
	$sText = $aTextSize[0]
	; ---
	__libNotif_SetPosAndSize($aTextSize[2], $aTextSize[3])
	GUICtrlSetData($__gLibNotif_iTextLabel, $sText)
	; ---
	__libNotif_Show()
	; ---
	If $__gLibNotif_aCfg[$LIBNOTIF_CFG_TIMEOUT] > 0 Then
		AdlibRegister("__libNotif_TimedHide", $__gLibNotif_aCfg[$LIBNOTIF_CFG_TIMEOUT])
	EndIf
EndFunc

; Hide notification.
; Set $iWithoutAnimation to hide the notification immediatly
;
Func _LibNotif_Hide($iWithoutAnimation = False)
	__libNotif_Hide($iWithoutAnimation)
	AdlibUnRegister("__libNotif_TimedHide")
EndFunc

; Check if there is a visible notification.
;
Func _LibNotif_IsVisible()
	Return $__gLibNotif_iIsVisible
EndFunc

; ===============================================================================================================================
; internals

Func __libNotif_TimedHide()
	__libNotif_Hide()
	AdlibUnRegister("__libNotif_TimedHide")
EndFunc

Func __libNotif_SetPosAndSize($iTextWidth, $iTextHeight)
	Local $iTextMargin = $__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTMARGIN], $iButtonSize = $__gLibNotif_aCfg[$LIBNOTIF_CFG_BUTTONSIZE]
	; ---
	; calculate window size and pos
	Local $iMinWinWidth = $__gLibNotif_aCfg[$LIBNOTIF_CFG_MINTOASTWIDTH]
	Local $iMinWinHeight = $iButtonSize + (2 * $iTextMargin)

	Local $iWinWidth = $iTextWidth + $iTextMargin * 2 ; left + right margin
	Local $iWinHeight = $iTextHeight + $iTextMargin * 2 ; top + bottom margin
	If $__gLibNotif_iCloseButton <> -1 Then $iWinWidth += $iButtonSize + $iTextMargin

	If $iWinWidth < $iMinWinWidth Then $iWinWidth = $iMinWinWidth
	If $iWinHeight < $iMinWinHeight Then $iWinHeight = $iMinWinHeight

	Local $aWinPos = _WinSizeToPos($iWinWidth, $iWinHeight, $__gLibNotif_aCfg[$LIBNOTIF_CFG_POSITION])
	; ---
	; resize window
	WinMove($__gLibNotif_hWnd, "", $aWinPos[0], $aWinPos[1], $iWinWidth, $iWinHeight)
	; ---
	; resize text label
	WinMove($__gLibNotif_hTextLabel, "", $iTextMargin, $iTextMargin, $iTextWidth, $iTextHeight)
	; ---
	; resize close button (if it exists)
	If $__gLibNotif_iCloseButton <> -1 Then
		WinMove($__gLibNotif_hCloseButton, "", $iWinWidth - ($iTextMargin + $iButtonSize), $iTextMargin, $iButtonSize, $iButtonSize)
	EndIf
EndFunc

Func __libNotif_Show($iInstant = False)
	If $__gLibNotif_iIsVisible Then Return
	; ---
	If Not $iInstant And $__gLibNotif_aCfg[$LIBNOTIF_CFG_SHOW_ANIMSTYLE] And $__gLibNotif_aCfg[$LIBNOTIF_CFG_SHOW_ANIMTIME] Then
		_AnimateWindow($__gLibNotif_hWnd, $__gLibNotif_aCfg[$LIBNOTIF_CFG_SHOW_ANIMTIME], $__gLibNotif_aCfg[$LIBNOTIF_CFG_SHOW_ANIMSTYLE])
	EndIf
	GUISetState(@SW_SHOWNOACTIVATE)
	$__gLibNotif_iIsVisible = 1
EndFunc

Func __libNotif_Hide($iInstant = False)
	If Not $__gLibNotif_iIsVisible Then Return
	; ---
	If Not $iInstant And $__gLibNotif_aCfg[$LIBNOTIF_CFG_HIDE_ANIMSTYLE] And $__gLibNotif_aCfg[$LIBNOTIF_CFG_HIDE_ANIMTIME] Then
		_AnimateWindow($__gLibNotif_hWnd, $__gLibNotif_aCfg[$LIBNOTIF_CFG_HIDE_ANIMTIME], $__gLibNotif_aCfg[$LIBNOTIF_CFG_HIDE_ANIMSTYLE])
	EndIf
	GUICtrlSetData($__gLibNotif_iTextLabel, "")
	If Not BitAND($__gLibNotif_aCfg[$LIBNOTIF_CFG_HIDE_ANIMSTYLE], $AW_HIDE) Then GUISetState(@SW_HIDE)
	$__gLibNotif_iIsVisible = 0
EndFunc

Func __libNotif_WM_Handler($hWnd, $iMsg, $wParam, $lParam)
;~ 	ConsoleWrite($hWnd & ", 0x" & Hex($iMsg, 4) & ", " & $wParam & ", " & $lParam & @CRLF)
	If $hWnd = $__gLibNotif_hWnd Or $hWnd = $__gLibNotif_hTextLabel Then
		Switch $iMsg
			Case $WM_COMMAND
;~ 				Local $nNotifyCode = BitShift($wParam, 16) ; high word
;~ 				Local $nID = BitAND($wParam, 0x0000FFFF) ; low word
;~ 				Local $hCtrl = $lParam
				If $__gLibNotif_iCloseButton <> -1 And $lParam = $__gLibNotif_hCloseButton And BitShift($wParam, 16) = $BN_CLICKED Then
					AdlibRegister("__libNotif_TimedHide", 1)
					$iSignal = RegRead("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked")
					If $iSignal = '1' Then
						$iRegWrite = RegWrite("HKEY_CURRENT_USER\Software\AppControlTray", "Blocked", "REG_DWORD", "0")
					EndIf
				EndIf
;~ 			Case $WM_NCLBUTTONDOWN, $WM_LBUTTONDOWN
;~ 				ConsoleWrite("LBUTTON" & @CRLF)
;~ 			Case $WM_NCRBUTTONDOWN, $WM_RBUTTONDOWN
;~ 				ConsoleWrite("RBUTTON" & @CRLF)
;~ 			Case Else
;~ 				ConsoleWrite("WM_Handler " & $iMsg & @CRLF)
		EndSwitch
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc

Func _AnimateWindow($hWnd, $iTime, $iFlags)
	Local $ret = DllCall("user32.dll", "bool", "AnimateWindow", "hwnd", $hWnd, "dword", $iTime, "dword", $iFlags)
	If @error Then
		Local $err = @error
		If Not @Compiled Then ConsoleWrite("! DllCall error " & $err & " (AnimateWindow)" & @CRLF)
		Return SetError($err, 0, 0)
	EndIf
	Return $ret[0]
EndFunc

; get the position of a window of size $iW,$iH that will position it in bottom-right corner of the screen
Func _WinSizeToPos($iW, $iH, $iPos = $LIBNOTIF_POS_LR)
	Local $tRect = DllStructCreate("long;long;long;long") ; left,top,right,bottom
    Local $aRet = DllCall('user32.dll', 'int', 'SystemParametersInfo', 'uint', 48, 'uint', 0, 'struct*', $tRect, 'uint', 0)
    If @error Then Return SetError(@error, 0, 0)
	; ---
	Switch $iPos
		Case $LIBNOTIF_POS_LR
			Local $aRet[2] = [DllStructGetData($tRect, 3) - $iW, DllStructGetData($tRect, 4) - $iH]
		Case $LIBNOTIF_POS_LL
			Local $aRet[2] = [0, DllStructGetData($tRect, 4) - $iH]
		Case $LIBNOTIF_POS_UR
			Local $aRet[2] = [DllStructGetData($tRect, 3) - $iW, 0]
		Case $LIBNOTIF_POS_UL
			Local $aRet[2] = [0, 0]
	EndSwitch
    Return $aRet
EndFunc

; from Melba23's toast.au3
Func __libNotif_GetDefaultFont()
	; Get default system font data
	Local $tNONCLIENTMETRICS = DllStructCreate("uint;int;int;int;int;int;byte[60];int;int;byte[60];int;int;byte[60];byte[60];byte[60]")
	DllStructSetData($tNONCLIENTMETRICS, 1, DllStructGetSize($tNONCLIENTMETRICS))
	DllCall("user32.dll", "int", "SystemParametersInfo", "int", 41, "int", DllStructGetSize($tNONCLIENTMETRICS), "ptr", DllStructGetPtr($tNONCLIENTMETRICS), "int", 0)
	; Read font data for MsgBox font
	Local $tLOGFONT = DllStructCreate("long;long;long;long;long;byte;byte;byte;byte;byte;byte;byte;byte;char[32]", DllStructGetPtr($tNONCLIENTMETRICS, 15))
	; ---
	$__gLibNotif_aDefaultCfg[$LIBNOTIF_CFG_TEXTFONTNAME] = DllStructGetData($tLOGFONT, 14)
	$__gLibNotif_aDefaultCfg[$LIBNOTIF_CFG_TEXTSIZE] = Int((Abs(DllStructGetData($tLOGFONT, 1)) + 1) * .75)

;~ 	$__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTFONTNAME] = DllStructGetData($tLOGFONT, 14)
;~ 	$__gLibNotif_aCfg[$LIBNOTIF_CFG_TEXTSIZE] = Int((Abs(DllStructGetData($tLOGFONT, 1)) + 1) * .75)
EndFunc
