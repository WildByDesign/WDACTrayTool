#include-once

; #INDEX# ============================================================================================================
; Title .........: ExtMsgBox
; AutoIt Version : v3.2.12.1 or higher
; Language ......: English
; Description ...: Generates user defined message boxes centred on a GUI, on screen or at defined coordinates
; Remarks .......:
; Note ..........:
; Author(s) .....: Melba23, based on some original code by photonbuddy & YellowLab, and KaFu (default font code)
; ====================================================================================================================

;#AutoIt3Wrapper_au3check_parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w- 7

; #INCLUDES# =========================================================================================================
#include "StringSize.au3"
#include <GDIPlus.au3>
#include <WinAPIGdi.au3>
#include <WinAPISys.au3>

; #GLOBAL CONSTANTS# =================================================================================================
Global Const $EMB_ICONSTOP = 16 ; Stop-sign icon
Global Const $EMB_ICONQUERY = 32 ; Question-mark icon
Global Const $EMB_ICONEXCLAM = 48 ; Exclamation-point icon
Global Const $EMB_ICONINFO = 64 ; Icon consisting of an 'i' in a circle

; #GLOBAL VARIABLES# =================================================================================================

Global $g_aEMB_Settings[15]
; [0] = Style			[6]  = Max Width              [12] = Title bar reduction
; [1] = Justification	[7]  = Absolute Width         [13] = Focused button character
; [2] = Back Colour		[8]  = Default Back Colour    [14] = Default titlebar icon
; [3] = Text Colour		[9]  = Default Text Colour
; [4] = Font Size		[10] = Default Font Size
; [5] = Font Name		[11] = Default Font Name
;

; Default settings
; Font
Global $g_aEMB_TempArray = __EMB_GetDefaultFont()
$g_aEMB_Settings[10] = $g_aEMB_TempArray[0]
$g_aEMB_Settings[11] = $g_aEMB_TempArray[1]
; Colours
$g_aEMB_TempArray = DllCall("User32.dll", "int", "GetSysColor", "int", 15) ; $COLOR_3DFACE
$g_aEMB_Settings[8] = BitAND(BitShift(String(Binary($g_aEMB_TempArray[0])), 8), 0xFFFFFF)
$g_aEMB_TempArray = DllCall("User32.dll", "int", "GetSysColor", "int", 8) ; $COLOR_WINDOWTEXT
$g_aEMB_Settings[9] = BitAND(BitShift(String(Binary($g_aEMB_TempArray[0])), 8), 0xFFFFFF)
; Title bar width reduction by icon and [X] button in various themes
$g_aEMB_TempArray = DllCall("user32.dll", "int", "GetSystemMetrics", "int", 11) ; Title bar icon width
$g_aEMB_Settings[12] = $g_aEMB_TempArray[0]
$g_aEMB_TempArray = DllCall("user32.dll", "int", "GetSystemMetrics", "int", 30) ; Title bar button width
$g_aEMB_Settings[12] += ( ($g_aEMB_TempArray[0] < 30) ? ($g_aEMB_TempArray[0] * 3) : ($g_aEMB_TempArray[0]) ) ; Compensate for small buttons in some themes
$g_aEMB_TempArray = 0
$g_aEMB_TempArray = DllCall("dwmapi.dll", "uint", "DwmIsCompositionEnabled", "int*", $g_aEMB_TempArray) ; Check for Aero enabled
If Not @error And $g_aEMB_TempArray[1] = True Then
	$g_aEMB_TempArray = DllCall("user32.dll", "int", "GetSystemMetrics", "int", 7) ; GUI button frame width
	$g_aEMB_Settings[12] += ($g_aEMB_TempArray[0] * 4) ; Add frames to compensate for incorrect Aero return
EndIf
$g_aEMB_TempArray = 0

; Set default values
$g_aEMB_Settings[0] = 0
$g_aEMB_Settings[1] = 0
$g_aEMB_Settings[2] = $g_aEMB_Settings[8]
$g_aEMB_Settings[3] = $g_aEMB_Settings[9]
$g_aEMB_Settings[4] = $g_aEMB_Settings[10]
$g_aEMB_Settings[5] = $g_aEMB_Settings[11]
$g_aEMB_Settings[6] = 370
$g_aEMB_Settings[7] = 500
$g_aEMB_Settings[13] = "~"
$g_aEMB_Settings[14] = ""

; #CURRENT# ==========================================================================================================
; _ExtMsgBoxSet: Sets the GUI style, justification, colours, font and max width for subsequent _ExtMsgBox function calls
; _ExtMsgBox:    Generates user defined message boxes centred on a GUI, on screen or at defined coordinates
; ====================================================================================================================

; #INTERNAL_USE_ONLY#=================================================================================================
; __EMB_GetDefaultFont: Determines Windows default MsgBox font size and name
; ====================================================================================================================

; #FUNCTION# =========================================================================================================
; Name...........: _ExtMsgBoxSet
; Description ...: Sets the GUI style, justification, colours, font and max width for subsequent _ExtMsgBox function calls
; Syntax.........: _ExtMsgBoxSet($iStyle, $iJust, [$iBkCol, [$iCol, [$sFont_Size, [$iFont_Name, [$iWidth, [$iWidth_Abs, [$sFocus_Char, [$sTitlebar_Icon]]]]]]]])
; Parameters ....: $iStyle          -> 0 (Default) - Taskbar Button, TOPMOST, button in user font, no tab expansion,
;                                          no checkbox, titlebar icon, active closure [X] and SysMenu close
;                                      Combine following to change:
;                                          1   = No Taskbar Button
;                                          2   = TOPMOST Style not set
;                                          4   = Buttons use default font
;                                          8   = Expand Tabs to ensure adequate sizing of GUI
;                                          16  = "Do not display again" checkbox
;                                          32  = Show no icon on title bar
;                                          64  = Disable EMB closure [X] and SysMenu Close
;                   $iJust          -> 0 = Left justified (Default), 1 = Centred , 2 = Right justified
;                                          + 4 = Centred single button.  Note: multiple buttons are always centred
;                                          ($SS_LEFT, $SS_CENTER, $SS_RIGHT can also be used - must #include <StaticConstants.au3>)
;                   $iBkCol		    -> The colour for the message box background.  Default = system colour
;                   $iCol		    -> The colour for the message box text.  Default = system colour
;                   $iFont_Size     -> The font size in points to use for the message box. Default = system font size
;                   $sFont_Name     -> The font to use for the message box. Default = system font
;                   $iWidth         -> Normal max width for EMB.   Default/min = 370 pixels - max = @DesktopWidth - 20
;                   $iWidth_Abs     -> Absolute max width for EMB. Default/min = 370 pixels - max = @DesktopWidth - 20
;                                       EMB will expand to this value to accommodate long unbroken character strings
;                                       Forced to $iWidth value if less
;                   $sFocus_Char    -> Character to define focused button. Default = "~"
;                   $sTitlebar_Icon -> Icon to use on ExtMsgBox titlebar - default is standard AutoIt icon
;                                          When set to name of an ico or exe file, the main icon within will be displayed
;                                          If another icon from the file is required, add a trailing "|" followed by the icon index
; Requirement(s).: v3.2.12.1 or higher
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error to 1 with @extended set to incorrect parameter index number
; Remarks .......; - Setting any parameter to -1 leaves the current value unchanged
;                    Setting the $iStyle parameter to 'Default' resets ALL parameters to default values <<<<<<<<<<<<<<<<<<<<<<<
;                    Setting any other parameter to "Default" only resets that parameter
;                  - Setting a titlebar icon overrides any "show no icon" numeric setting
; Author ........: Melba23
; Example........; Yes
;=====================================================================================================================
Func _ExtMsgBoxSet($iStyle = -1, $iJust = -1, $iBkCol = -1, $iCol = -1, $iFont_Size = -1, $sFont_Name = -1, $iWidth = -1, $iWidth_Abs = -1, $sFocus_Char = "~", $sTitlebar_Icon = "")

	; Set global EMB variables to required values
	Switch $iStyle
		Case Default
			$g_aEMB_Settings[0] = 0
			$g_aEMB_Settings[1] = 0
			$g_aEMB_Settings[2] = $g_aEMB_Settings[8]
			$g_aEMB_Settings[3] = $g_aEMB_Settings[9]
			$g_aEMB_Settings[5] = $g_aEMB_Settings[11]
			$g_aEMB_Settings[4] = $g_aEMB_Settings[10]
			$g_aEMB_Settings[6] = 370
			$g_aEMB_Settings[7] = 370
			$g_aEMB_Settings[13] = "~"
			$g_aEMB_Settings[14] = ""
			Return
		Case -1
			; Do nothing
		Case 0 To 127
			$g_aEMB_Settings[0] = Int($iStyle)
		Case Else
			Return SetError(1, 1, 0)
	EndSwitch

	Switch $iJust
		Case Default
			$g_aEMB_Settings[1] = 0
		Case -1
			; Do nothing
		Case 0, 1, 2, 4, 5, 6
			$g_aEMB_Settings[1] = $iJust
		Case Else
			Return SetError(1, 2, 0)
	EndSwitch

	Switch $iBkCol
		Case Default
			$g_aEMB_Settings[2] = $g_aEMB_Settings[8]
		Case -1
			; Do nothing
		Case 0 To 0xFFFFFF
			$g_aEMB_Settings[2] = Int($iBkCol)
		Case Else
			Return SetError(1, 3, 0)
	EndSwitch

	Switch $iCol
		Case Default
			$g_aEMB_Settings[3] = $g_aEMB_Settings[9]
		Case -1
			; Do nothing
		Case 0 To 0xFFFFFF
			$g_aEMB_Settings[3] = Int($iCol)
		Case Else
			Return SetError(1, 4, 0)
	EndSwitch

	Switch $iFont_Size
		Case Default
			$g_aEMB_Settings[4] = $g_aEMB_Settings[10]
		Case -1
			; Do nothing
		Case 8 To 72
			$g_aEMB_Settings[4] = Int($iFont_Size)
		Case Else
			Return SetError(1, 5, 0)
	EndSwitch

	Switch $sFont_Name
		Case Default
			$g_aEMB_Settings[5] = $g_aEMB_Settings[11]
		Case -1
			; Do nothing
		Case Else
			If IsString($sFont_Name) Then
				$g_aEMB_Settings[5] = $sFont_Name
			Else
				Return SetError(1, 6, 0)
			EndIf
	EndSwitch

	Switch $iWidth
		Case Default
			$g_aEMB_Settings[6] = 370
		Case -1
			; Do nothing
		Case 370 To @DesktopWidth - 20
			$g_aEMB_Settings[6] = Int($iWidth)
		Case Else
			Return SetError(1, 7, 0)
	EndSwitch

	Switch $iWidth_Abs
		Case Default
			$g_aEMB_Settings[7] = 370
		Case -1
			; Do nothing
		Case 370 To @DesktopWidth - 20
			$g_aEMB_Settings[7] = Int($iWidth_Abs)
		Case Else
			Return SetError(1, 8, 0)
	EndSwitch

	; Check absolute width is at least max width
	If $g_aEMB_Settings[7] < $g_aEMB_Settings[6] Then
		$g_aEMB_Settings[7] = $g_aEMB_Settings[6]
	EndIf

	Switch $sFocus_Char
		Case Default
			$g_aEMB_Settings[13] = "~"
		Case -1
			; Do nothing
		Case Else
			If IsString($sFocus_Char) Then
				$g_aEMB_Settings[13] = StringLeft($sFocus_Char, 1)
			Else
				$g_aEMB_Settings[13] = "~"
			EndIf
	EndSwitch

	Switch $sTitlebar_Icon
		Case Default
			$g_aEMB_Settings[14] = ""
		Case -1
			; Do nothing
		Case Else
			If IsString($sTitlebar_Icon) Then
				$g_aEMB_Settings[14] = $sTitlebar_Icon
			Else
				$g_aEMB_Settings[14] = ""
			EndIf
	EndSwitch

	Return 1

EndFunc   ;==>_ExtMsgBoxSet

; #FUNCTION# =========================================================================================================
; Name...........: _ExtMsgBox
; Description ...: Generates user defined message boxes centred on a GUI, the desktop, or at defined coordinates
; Syntax.........: _ExtMsgBox ($vIcon, $vButton, $sTitle, $sText, [$vTimeout, [$hWin, [$iVPos, [$bMain = True]]]])
; Parameters ....: $vIcon   -> Sets the required icon in the GUI display and optionally in the titlebar of the ExtMsgBox
;                                  GUI display icon:
;                                      Numeric value:
;                                          0   - No icon
;                                          8   - UAC
;                                          16  - Stop         )
;                                          32  - Query        ) or equivalent $MB/$EMB_ICON constant
;                                          48  - Exclamation  )
;                                          64  - Information  )
;                                          128 - Countdown digits if $iTimeOut set
;                                              Any other numeric value returns Error 1
;                                      String value:
;                                          If set to the name of an ico or exe file, the main icon within will be displayed
;                                              If another icon from the file is required, add a trailing "|" followed by the icon index
;                                          If set to the name of an image file, that image will be displayed
;                                  Titlebar icon:
;                                      Use a semicolon delimiter followed by the name/index of the required exe/ico file
;                                      Default (no delimiter or name) = use global icon setting from _ExtMsgBoxSet
;                   $vButton  -> Button text separated with "|" character. " " = no buttons.
;                                      Putting a user-defined character (default = "~") before the text indicates button to be focused
;                                         Two focused buttons returns Error 2. A single button is always focused
;                                      Use of standard Windows keyboard shortcut by adding "&" before the button text is still available
;                                          Note UDF focused button character must be placed first, e.g. "~&Yes|&No"
;                                      Can also use $MB_ button numeric constants to define buttons: 0 = "OK", 1 = "~OK|Cancel",
;                                          2 = "~Abort|Retry|Ignore", 3 = "~Yes|No|Cancel", 4 = "~Yes|No", 5 = "~Retry|Cancel",
;                                          6 = "~Cancel|Try Again|Continue".  Other values return Error 3
;                                      Default max width of 370 gives 1-4 buttons @ width 80, 5 @ width 60, 6 @ width 50
;                                      Min button width set at 50, so unless default widths changed 7 buttons will return Error 4
;                   $sTitle   -> The title of the message box.
;                                      Procrustean truncation if too long to fit
;                   $sText    -> The text to be displayed. Long lines will wrap. The box depth is adjusted to fit.
;                                      If unbroken character strings in $sText too long for set max width,
;                                      EMB expands to set absolute width. Error 6 if still not able to fit
;                   $vTimeout -> Single integer sets timeout delay in secs before EMB closes. 0 = no timeout (Default).
;                                      If no buttons and no timeout set, timeout automatically set to 5
;                                Colon-delimited string sets EMB timeout delay and initial button(s) disabled delay ("EMB:Button")
;                                      Default = button(s) not disabled
;                   $hWin     -> Handle of the GUI in which EMB is centred
;                                      If GUI  hidden or no handle passed - EMB centred in desktop (Default)
;                                      If not valid window handle, interpreted as horizontal coordinate for EMB location
;                   $iVPos    -> Vertical coordinate for EMB location
;                                      Only valid if $hWin parameter interpreted as horizontal coordinate (Default = 0)
;                   $bMain    -> True (default) = Adjust dialog position to ensure dialog positioned on main screen
;                                      False = Dialog positioned at user-defined coords, which can be on other screens
; Requirement(s).: v3.2.12.1 or higher
; Return values .: Success:	Returns 1-based index of the button pressed, counting from the LEFT.
;                           Returns 0 if closed by a "CloseGUI" event (i.e. click [X] or press Escape)
;                           Returns 9 if timed out
;                           If "Not again" checkbox is present and checked, return value is negated
;                  Failure:	Returns -1 and sets @error as follows:
;                               1 - Icon parameter error
;                               2 - Multiple default button error
;                               3 - Button constant error
;                               4 - Too many buttons to fit in max available EMB width
;                               5 - Button text too long for max available button width
;                               6 - StringSize error
;                               7 - GUI creation error
; Remarks .......; - If $bMain set EMB adjusted to appear on main screen closest to required position.
;                  - Setting a titlebar icon overrides any _ExtMsgBoxSet global setting - either "no-icon" or a specific file.
;                  - When buttons disabled, closure "X" still functions. Use $iStyle parameter of _ExtMsgBoxSet to prevent this.
; Author ........: Melba23, based on some original code by photonbuddy & YellowLab
; Example........; Yes
;=====================================================================================================================
Func _ExtMsgBox($vIcon, $vButton, $sTitle, $sText, $vTimeOut = 0, $hWin = "", $iVPos = 0, $bMain = True)

	; Set default sizes for message box
	Local $iMsg_Width_Max = $g_aEMB_Settings[6], $iMsg_Width_Min = 150, $iMsg_Width_Abs = $g_aEMB_Settings[7]
	Local $iMsg_Height_Min = 100, $iButton_Width_Def = 80, $iButton_Width_Min = 50, $iTimeOut = 0, $iTimeIn = 0

	; Declare local variables
	Local $iParent_Win = 0, $fCountdown = False, $cCheckbox, $aLabel_Size, $aRet, $iRet_Value, $iHpos
	Local $sButton_Text, $iButton_Width, $iButton_Xpos

	; Check for icon(s)
	Local $vTitlebar_Icon
	Local $iTitlebar_Icon_Style = 0
	Local $sTitlebar_DLL = ""
	Local $iIcon_Style = 0
	Local $iIcon_Reduction = 42
	Local $sDLL = "user32.dll"
	Local $sImg = ""
	Local $aSplit

	; Validate timeout value
	$iTimeOut = Int($vTimeOut)
	; If negative then increase timer font size if required
	If $iTimeOut < 0 Then
		$iTimeOut = Abs($iTimeOut)
		$iIcon_Reduction = 74
	EndIf
	; Check for button disabling
    If StringInStr($vTimeOut, ":") Then
        $iTimeIn = Abs(Int(StringSplit($vTimeOut, ":")[2]))
    EndIf
	; Set automatic timeout if no buttons and no timeout set
	If $vButton == " " And $iTimeOut = 0 Then
		$iTimeOut = 5
	EndIf

	; Check icon parameter for titlebar icon data
	If StringInStr($vIcon, ";") Then
		; Extract GUI icon data
		$aSplit = StringSplit($vIcon, ";")
		; Keep display icon data
		$vIcon = $aSplit[1]
		; Get required titlebar icon data
		$sTitlebar_DLL = $aSplit[2]
		; Parse data
		If StringInStr($vTitlebar_Icon, "|") Then
			$iTitlebar_Icon_Style = StringRegExpReplace($sTitlebar_DLL, "(.*)\|", "")
            $sTitlebar_DLL = StringRegExpReplace($sTitlebar_DLL, "\|.*$", "")
		EndIf
	Else
		; Use global setting if available
		$sTitlebar_DLL = $g_aEMB_Settings[14]
	EndIf

	; Cancel numeric countdown if no timeout
	If $iTimeOut = 0 And $vIcon = 128 Then
		$vIcon = 0
	EndIf

	; Get required display icon data
	If StringIsDigit($vIcon) Then
		Switch $vIcon
			Case 0
				$iIcon_Reduction = 0
			Case 8
				$sDLL = "imageres.dll"
				$iIcon_Style = 78
			Case 16 ; Stop
				$iIcon_Style = -4
			Case 32 ; Query
				$iIcon_Style = -3
			Case 48 ; Exclam
				$iIcon_Style = -2
			Case 64 ; Info
				$iIcon_Style = -5
			Case 128 ; Countdown
				If $iTimeOut > 0 Then
					$fCountdown = True
				EndIf
			Case Else
				Return SetError(1, 0, -1)
		EndSwitch
	Else
		If StringInStr($vIcon, "|") Then
			$iIcon_Style = StringRegExpReplace($vIcon, "(.*)\|", "")
            $vIcon = StringRegExpReplace($vIcon, "\|.*$", "")
		EndIf
		; Parse data
		Switch StringLower(StringRight($vIcon, 3))
			Case "exe", "ico"
				$sDLL = $vIcon
			Case "bmp", "jpg", "gif", "png"
				$sImg = $vIcon
		EndSwitch
	EndIf

	; Check if two buttons are seeking focus
	Local $sFocus_Char = $g_aEMB_Settings[13]
	StringReplace($vButton, $sFocus_Char, "")
	If @extended > 1 Then
		Return SetError(2, 0, -1)
	EndIf

	; Check if using constants or text
	If IsNumber($vButton) Then
		Switch $vButton
			Case 0
				$vButton = "OK"
			Case 1
				$vButton = $sFocus_Char & "OK|Cancel"
			Case 2
				$vButton = $sFocus_Char & "Abort|Retry|Ignore"
			Case 3
				$vButton = $sFocus_Char & "Yes|No|Cancel"
			Case 4
				$vButton = $sFocus_Char & "Yes|No"
			Case 5
				$vButton = $sFocus_Char & "Retry|Cancel"
			Case 6
				$vButton = $sFocus_Char & "Cancel|Try Again|Continue"
			Case Else
				Return SetError(3, 0, -1)
		EndSwitch
	EndIf

	; Set default values
	Local $aButton_Text[1] = [0]
	Local $iButton_Width_Req = 0
	; Get required button size
	If $vButton <> " " Then
		; Split button text into individual strings
		$aButton_Text = StringSplit($vButton, "|")

		; Get absolute available width for each button
		Local $iButton_Width_Abs = Floor((($iMsg_Width_Max - 10) / $aButton_Text[0]) - 10)
		; Error if below min button size
		If $iButton_Width_Abs < $iButton_Width_Min Then
			Return SetError(4, 0, -1)
		EndIf
		; Determine required size of buttons to fit text
		Local $iButton_Width_Text = 0
		; Loop through button text
		For $i = 1 To $aButton_Text[0]
			; Remove a possible leading focus character
			$sButton_Text = StringRegExpReplace($aButton_Text[$i], "^" & $sFocus_Char & "?(.*)$", "$1")
			; Check on font to use
			If BitAND($g_aEMB_Settings[0], 4) Then
				$aRet = _StringSize($sButton_Text, $g_aEMB_Settings[10], Default, Default, $g_aEMB_Settings[11])
			Else
				$aRet = _StringSize($sButton_Text, $g_aEMB_Settings[4], Default, Default, $g_aEMB_Settings[5])
			EndIf
			If IsArray($aRet) And $aRet[2] + 10 > $iButton_Width_Text Then
				; Find max button width required for text
				$iButton_Width_Text = $aRet[2] + 10
			EndIf
		Next
		; Error if text would make required button width > absolute available
		If $iButton_Width_Text > $iButton_Width_Abs Then
			Return SetError(5, 0, -1)
		EndIf
		; Determine button size to use - assume default
		$iButton_Width = $iButton_Width_Def
		; If text requires wider then default
		If $iButton_Width_Text > $iButton_Width_Def Then
			; Increase - cannot be > absolute
			$iButton_Width = $iButton_Width_Text
		EndIf
		; If absolute < default
		If $iButton_Width_Abs < $iButton_Width_Def Then
			; If text > min (text must be < abs)
			If $iButton_Width_Text > $iButton_Width_Min Then
				; Set text width
				$iButton_Width = $iButton_Width_Text
			Else
				; Set min width
				$iButton_Width = $iButton_Width_Min
			EndIf
		EndIf
		; Determine GUI width required for all buttons at this width
		$iButton_Width_Req = (($iButton_Width + 10) * $aButton_Text[0]) + 10
	EndIf

	; Set tab expansion flag if required
	Local $iExpTab = Default
	If BitAND($g_aEMB_Settings[0], 8) Then
		$iExpTab = 1
	EndIf

	; Get message label size
	While 1
		Local $aLabel_Pos = _StringSize($sText, $g_aEMB_Settings[4], Default, $iExpTab, $g_aEMB_Settings[5], $iMsg_Width_Max - 20 - $iIcon_Reduction)
		If @error Then
			If $iMsg_Width_Max >= $iMsg_Width_Abs Then
				Return SetError(6, 0, -1)
			Else
				$iMsg_Width_Max += 10
			EndIf
		Else
			ExitLoop
		EndIf
	WEnd
	; Reset text to wrapped version
	$sText = $aLabel_Pos[0]
	; Set label size
	Local $iLabel_Width = $aLabel_Pos[2]
	Local $iLabel_Height = $aLabel_Pos[3]

	; Set GUI size
	Local $iMsg_Width = $iLabel_Width + 20 + $iIcon_Reduction
	; Increase width to fit buttons if needed
	If $iButton_Width_Req > $iMsg_Width Then $iMsg_Width = $iButton_Width_Req
	If $iMsg_Width < $iMsg_Width_Min Then
		$iMsg_Width = $iMsg_Width_Min
		$iLabel_Width = $iMsg_Width_Min - 20
	EndIf

	; Check if title sets width
	Local $iDialog_Width = $iMsg_Width
	; Size title
	Local $aTitleSize = _StringSize($sTitle, $g_aEMB_Settings[10], Default, Default, $g_aEMB_Settings[11])

	; Check if title wider than text
	If $aTitleSize[2] > ($iMsg_Width - 70) Then ; Assume icon reduction of 50 regardless of icon setting
		; Adjust dialog width up to absolute dialog width value
		$iDialog_Width = ( ($aTitleSize[2] < ($g_aEMB_Settings[7] - $g_aEMB_Settings[12])) ? ($aTitleSize[2] + $g_aEMB_Settings[12]) : ($g_aEMB_Settings[7]) )
	EndIf

	Local $iMsg_Height = $iLabel_Height + 35
	; Increase height if buttons present
	If $vButton <> " " Then
		$iMsg_Height += 30
	EndIf
	; Increase height if checkbox required
	If BitAND($g_aEMB_Settings[0], 16) Then
		$iMsg_Height += 40
	EndIf
	If $iMsg_Height < $iMsg_Height_Min Then $iMsg_Height = $iMsg_Height_Min

	; If only single line, lower label to to centre text on icon
	Local $iLabel_Vert = 20
	If StringInStr($sText, @CRLF) = 0 Then $iLabel_Vert = 27

	; Check for taskbar button style required
	If Mod($g_aEMB_Settings[0], 2) = 1 Then ; Hide taskbar button so create as child
		If IsHWnd($hWin) Then
			$iParent_Win = $hWin ; Make child of that window
		Else
			$iParent_Win = WinGetHandle(AutoItWinGetTitle()) ; Make child of AutoIt window
		EndIf
	EndIf

	; Determine EMB location
	If $hWin = "" Then
		; No handle or position passed so centre on screen
		$iHpos = (@DesktopWidth - $iDialog_Width) / 2
		$iVPos = (@DesktopHeight - $iMsg_Height) / 2
	Else
		If IsHWnd($hWin) Then
			; Get parent GUI pos if visible
			If BitAND(WinGetState($hWin), 2) Then
				; Set EMB to centre on parent
				Local $aPos = WinGetPos($hWin)
				$iHpos = ($aPos[2] - $iDialog_Width) / 2 + $aPos[0] - 3
				$iVPos = ($aPos[3] - $iMsg_Height) / 2 + $aPos[1] - 20
			Else
				; Set EMB to centre om screen
				$iHpos = (@DesktopWidth - $iDialog_Width) / 2
				$iVPos = (@DesktopHeight - $iMsg_Height) / 2
			EndIf
		Else
			; Assume parameter is horizontal coord
			$iHpos = $hWin ; $iVpos already set
		EndIf
	EndIf

	; If dialog is to appear on main display
	If $bMain Then
		; Dialog is visible horizontally
		If $iHpos < 10 Then $iHpos = 10
		If $iHpos + $iDialog_Width > @DesktopWidth - 20 Then $iHpos = @DesktopWidth - 20 - $iDialog_Width
		; Then vertically
		If $iVPos < 10 Then $iVPos = 10
		If $iVPos + $iMsg_Height > @DesktopHeight - 60 Then $iVPos = @DesktopHeight - 60 - $iMsg_Height
	EndIf

	; Remove TOPMOST extended style if required
	Local $iExtStyle = 0x00000008 ; $WS_TOPMOST
	If BitAND($g_aEMB_Settings[0], 2) Then $iExtStyle = -1

	; Create GUI with $WS_POPUPWINDOW, $WS_CAPTION style and required extended style
	Local $hMsgGUI = GUICreate($sTitle, $iDialog_Width, $iMsg_Height, $iHpos, $iVPos, BitOR(0x80880000, 0x00C00000), $iExtStyle, $iParent_Win)
	If @error Then
		Return SetError(7, 0, -1)
	EndIf

	; Check if custom or no titlebar icon required
	Select
		; Specific icon set
		Case $sTitlebar_DLL <> ""
			GUISetIcon($sTitlebar_DLL, $iTitlebar_Icon_Style)
		; No icon - actually uses transparent icon from AutoIt executable
		Case BitAND($g_aEMB_Settings[0], 32)
			If @Compiled Then
				GUISetIcon(@ScriptName, -2, $hMsgGUI)
			Else
				GUISetIcon(@AutoItExe, -2, $hMsgGUI)
			EndIf
		; In all other cases use standard Autoit icon
	EndSelect

	If $g_aEMB_Settings[2] <> Default Then GUISetBkColor($g_aEMB_Settings[2])

	; Check if user closure permitted
	If BitAND($g_aEMB_Settings[0], 64) Then
		$aRet = DllCall("User32.dll", "hwnd", "GetSystemMenu", "hwnd", $hMsgGUI, "int", 0)
		Local $hSysMenu = $aRet[0]
		DllCall("User32.dll", "int", "RemoveMenu", "hwnd", $hSysMenu, "int", 0xF060, "int", 0) ; $SC_CLOSE
		DllCall("User32.dll", "int", "DrawMenuBar", "hwnd", $hMsgGUI)
	EndIf

	; Set centring parameter
	Local $iLabel_Style = 0 ; $SS_LEFT
	If BitAND($g_aEMB_Settings[1], 1) = 1 Then
		$iLabel_Style = 1 ; $SS_CENTER
	ElseIf BitAND($g_aEMB_Settings[1], 2) = 2 Then
		$iLabel_Style = 2 ; $SS_RIGHT
	EndIf

	; Create label
	GUICtrlCreateLabel($sText, 10 + $iIcon_Reduction, $iLabel_Vert, $iLabel_Width, $iLabel_Height, $iLabel_Style)
	GUICtrlSetFont(-1, $g_aEMB_Settings[4], Default, Default, $g_aEMB_Settings[5])
	If $g_aEMB_Settings[3] <> Default Then GUICtrlSetColor(-1, $g_aEMB_Settings[3])

	; Create checkbox if required
	If BitAND($g_aEMB_Settings[0], 16) Then
		Local $sAgain = " Do not show again"
		Local $iY = $iLabel_Vert + $iLabel_Height + 10
		; Create checkbox
		$cCheckbox = GUICtrlCreateCheckbox("", 10 + $iIcon_Reduction, $iY, 20, 20)
		; Write text in separate checkbox label
		Local $cCheckLabel = GUICtrlCreateLabel($sAgain, 20, 20, 20, 20)
		GUICtrlSetColor($cCheckLabel, $g_aEMB_Settings[3])
		GUICtrlSetBkColor($cCheckLabel, $g_aEMB_Settings[2])
		; Set font if required and size checkbox label text
		If BitAND($g_aEMB_Settings[0], 4) Then
			$aLabel_Size = _StringSize($sAgain)
		Else
			$aLabel_Size = _StringSize($sAgain, $g_aEMB_Settings[4], 400, 0, $g_aEMB_Settings[5])
			GUICtrlSetFont($cCheckLabel, $g_aEMB_Settings[4], 400, 0, $g_aEMB_Settings[5])
		EndIf
		; Move and resize checkbox label to fit
		$iY = ($iY + 10) - ($aLabel_Size[3] - 4) / 2
		ControlMove($hMsgGUI, "", $cCheckLabel, 30 + $iIcon_Reduction, $iY, $iMsg_Width - (30 + $iIcon_Reduction), $aLabel_Size[3])
	EndIf

	; Create icon, image or countdown timer
	If $fCountdown = True Then
		Local $cCountdown_Label = GUICtrlCreateLabel(StringFormat("%2s", $iTimeOut), 10, 20, $iIcon_Reduction - 10, $iIcon_Reduction - 10)
		GUICtrlSetFont(-1, ($iIcon_Reduction > 42) ? (36) : (18), Default, Default, $g_aEMB_Settings[5])
		GUICtrlSetColor(-1, $g_aEMB_Settings[3])
	Else
		If $iIcon_Reduction Then
			Switch StringLower(StringRight($sImg, 3))
				Case "bmp", "jpg", "gif"
					GUICtrlCreatePic($sImg, 10, 20, 32, 32)
				Case "png"
					__EMB_ShowPNG($sImg)
				Case Else
					GUICtrlCreateIcon($sDLL, $iIcon_Style, 10, 20)
			EndSwitch
		EndIf
	EndIf

	; Create buttons
	Local $aButtonCID[$aButton_Text[0] + 1] = [9999] ; Placeholder prevent problems if no buttons
    If $vButton <> " " Then

        ; Calculate button horizontal start
        If $aButton_Text[0] = 1 Then
            If BitAND($g_aEMB_Settings[1], 4) = 4 Then
                ; Single centred button
                $iButton_Xpos = ($iMsg_Width - $iButton_Width) / 2
            Else
                ; Single offset button
                $iButton_Xpos = $iMsg_Width - $iButton_Width - 10
            EndIf
        Else
            ; Multiple centred buttons
            $iButton_Xpos = ($iMsg_Width - ($iButton_Width_Req - 20)) / 2
        EndIf

		; Set default button style
        Local $iDef_Button_Style = 0
        ; Work through button list
        For $i = 0 To $aButton_Text[0] - 1
            Local $iButton_Text = $aButton_Text[$i + 1]
            ; Set default button
            If $aButton_Text[0] = 1 Then ; Only 1 button
                $iDef_Button_Style = 0x0001
            ElseIf StringLeft($iButton_Text, 1) = $sFocus_Char Then ; Look for focus character
                $iDef_Button_Style = 0x0001
                $aButton_Text[$i + 1] = StringTrimLeft($iButton_Text, 1)
            EndIf
            ; Draw button
            $aButtonCID[$i + 1] = GUICtrlCreateButton($aButton_Text[$i + 1], $iButton_Xpos + ($i * ($iButton_Width + 10)), $iMsg_Height - 35, $iButton_Width, 25, $iDef_Button_Style)
			; Set focus if default
			If $iDef_Button_Style Then
				GUICtrlSetState($aButtonCID[$i + 1], 256) ; $GUI_FOCUS
			EndIf
            ; Set font if required
            If Not BitAND($g_aEMB_Settings[0], 4) Then GUICtrlSetFont(-1, $g_aEMB_Settings[4], 400, 0, $g_aEMB_Settings[5])
            ; Reset default style parameter
            $iDef_Button_Style = 0
			; Disable if required
			If $iTimeIn Then
				GUICtrlSetState($aButtonCID[$i + 1], 128) ; $GUI_DISABLE
			EndIf
        Next
    EndIf

	; Show GUI
	GUISetState(@SW_SHOW, $hMsgGUI)

	; Begin timeout counter
	Local $iTimeout_Begin = TimerInit()
	Local $iCounter = 0

	; Declare GUIGetMsg return array here and not in loop
	Local $aMsg

	; Set MessageLoop mode
	Local $iOrgMode = Opt('GUIOnEventMode', 0)

	While 1
		$aMsg = GUIGetMsg(1)

		If $aMsg[1] = $hMsgGUI Then
			Select
				Case $aMsg[0] = -3 ; $GUI_EVENT_CLOSE
					$iRet_Value = 0
					ExitLoop
				Case Else
					; Check for other buttons
					For $i = 1 To UBound($aButtonCID) - 1
						If $aMsg[0] = $aButtonCID[$i] Then
							$iRet_Value = $i
							; No point in looking further
							ExitLoop 2
						EndIf
					Next
			EndSelect
		EndIf

		; Timeout if required
		If TimerDiff($iTimeout_Begin) / 1000 >= $iTimeOut And $iTimeOut > 0 Then
			$iRet_Value = 9
			ExitLoop
		EndIf

		; Check for button TimeIn
		If $iTimeIn > 0 And TimerDiff($iTimeout_Begin) / 1000 >= $iTimeIn Then
			For $i = 1 To $aButton_Text[0]
				GUICtrlSetState($aButtonCID[$i], 64) ; $GUI_ENABLE
			Next
			$iTimeIn = 0
		EndIf

		; Show countdown if required
		If $fCountdown = True Then
			Local $iTimeRun = Int(TimerDiff($iTimeout_Begin) / 1000)
			If $iTimeRun <> $iCounter Then
				$iCounter = $iTimeRun
				GUICtrlSetData($cCountdown_Label, StringFormat("%2s", $iTimeOut - $iCounter))
			EndIf
		EndIf

	WEnd

	; Reset original mode
	Opt('GUIOnEventMode', $iOrgMode)

	If GUICtrlRead($cCheckbox) = 1 Then
		; Negate the return value
		$iRet_Value *= -1
	EndIf

	GUIDelete($hMsgGUI)

	Return $iRet_Value

EndFunc   ;==>_ExtMsgBox

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: _EMB_GetDefaultFont
; Description ...: Determines Windows default MsgBox font size and name
; Syntax.........: _EMB_GetDefaultFont()
; Return values .: Success - Array holding determined font data
;                : Failure - Array holding default values
;                  Array elements - [0] = Size, [1] = Weight, [2] = Style, [3] = Name, [4] = Quality
; Author ........: KaFu
; Remarks .......: Used internally by ExtMsgBox UDF
; ===============================================================================================================================
Func __EMB_GetDefaultFont()

	; Fill array with standard default data
	Local $aDefFontData[2] = [9, "Tahoma"]

	; Get AutoIt GUI handle
	Local $hWnd = WinGetHandle(AutoItWinGetTitle())
	; Open Theme DLL
	Local $hThemeDLL = DllOpen("uxtheme.dll")
	; Get default theme handle
	Local $hTheme = DllCall($hThemeDLL, 'ptr', 'OpenThemeData', 'hwnd', $hWnd, 'wstr', "Static")
	If @error Then Return $aDefFontData
	$hTheme = $hTheme[0]
	; Create LOGFONT structure
	Local $tFont = DllStructCreate("long;long;long;long;long;byte;byte;byte;byte;byte;byte;byte;byte;wchar[32]")
	Local $pFont = DllStructGetPtr($tFont)
	; Get MsgBox font from theme
	DllCall($hThemeDLL, 'long', 'GetThemeSysFont', 'HANDLE', $hTheme, 'int', 805, 'ptr', $pFont) ; TMT_MSGBOXFONT
	If @error Then Return $aDefFontData
	; Get default DC
	Local $hDC = DllCall("user32.dll", "handle", "GetDC", "hwnd", $hWnd)
	If @error Then Return $aDefFontData
	$hDC = $hDC[0]
	; Get font vertical size
	Local $iPixel_Y = DllCall("gdi32.dll", "int", "GetDeviceCaps", "handle", $hDC, "int", 90) ; LOGPIXELSY
	If Not @error Then
		$iPixel_Y = $iPixel_Y[0]
		$aDefFontData[0] = Int(2 * (.25 - DllStructGetData($tFont, 1) * 72 / $iPixel_Y)) / 2
	EndIf
	; Close DC
	DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
	; Extract font data from LOGFONT structure
	$aDefFontData[1] = DllStructGetData($tFont, 14)

	Return $aDefFontData

EndFunc   ;==>__EMB_GetDefaultFont

; #INTERNAL_USE_ONLY#=================================================================================================
; Name...........: __EMB_ShowPNG
; Description ...: Set PNG as image
; Author ........: UEZ
; Modified.......: Melba23, guinness
; Remarks .......:
; ====================================================================================================================
Func __EMB_ShowPNG($sImg)

	_GDIPlus_Startup()
    Local $hPic = GUICtrlCreatePic("", 10, 20, 32, 32)
    Local $hBitmap = _GDIPlus_BitmapCreateFromFile($sImg)
    Local $hBitmap_Resized = _GDIPlus_BitmapCreateFromScan0(32, 32)
	Local $hBMP_Ctxt = _GDIPlus_ImageGetGraphicsContext($hBitmap_Resized)
	_GDIPlus_GraphicsSetInterpolationMode($hBMP_Ctxt, 7)
    _GDIPlus_GraphicsDrawImageRect($hBMP_Ctxt, $hBitmap, 0, 0, 32, 32)
    Local $hHBitmap = __EMB_BitmapCreateDIB($hBitmap_Resized)
    _WinAPI_DeleteObject(GUICtrlSendMsg($hPic, 0x0172, 0, $hHBitmap)) ; $STM_SETIMAGE
    _GDIPlus_BitmapDispose($hBitmap)
    _GDIPlus_BitmapDispose($hBitmap_Resized)
    _GDIPlus_GraphicsDispose($hBMP_Ctxt)
    _WinAPI_DeleteObject($hHBitmap)
    _GDIPlus_Shutdown()

EndFunc   ;==>__EMB_ShowPNG

; #INTERNAL_USE_ONLY#=================================================================================================
; Name...........: __EMB_BitmapCreateDIB
; Description ...: Create bitmap
; Author ........: UEZ
; Modified.......:
; Remarks .......:
; ====================================================================================================================
Func __EMB_BitmapCreateDIB($hBitmap)

	Local $hRet = 0

	Local $aRet1 = DllCall($__g_hGDIPDll, "uint", "GdipGetImageDimension", "ptr", $hBitmap, "float*", 0, "float*", 0)
	If (@error) Or ($aRet1[0]) Then Return 0
	Local $tData = _GDIPlus_BitmapLockBits($hBitmap, 0, 0, $aRet1[2], $aRet1[3], $GDIP_ILMREAD, $GDIP_PXF32ARGB)
	Local $pBits = DllStructGetData($tData, "Scan0")
	If Not $pBits Then Return 0
	Local $tBIHDR = DllStructCreate("dword;long;long;ushort;ushort;dword;dword;long;long;dword;dword")
	DllStructSetData($tBIHDR, 1, DllStructGetSize($tBIHDR))
	DllStructSetData($tBIHDR, 2, $aRet1[2])
	DllStructSetData($tBIHDR, 3, $aRet1[3])
	DllStructSetData($tBIHDR, 4, 1)
	DllStructSetData($tBIHDR, 5, 32)
	DllStructSetData($tBIHDR, 6, 0)
	Local $aRet2 = DllCall("gdi32.dll", "ptr", "CreateDIBSection", "hwnd", 0, "ptr", DllStructGetPtr($tBIHDR), "uint", 0, "ptr*", 0, "ptr", 0, "dword", 0)
	If (Not @error) And ($aRet2[0]) Then
		DllCall("gdi32.dll", "dword", "SetBitmapBits", "ptr", $aRet2[0], "dword", $aRet1[2] * $aRet1[3] * 4, "ptr", DllStructGetData($tData, "Scan0"))
		$hRet = $aRet2[0]
	EndIf
	_GDIPlus_BitmapUnlockBits($hBitmap, $tData)
	Return $hRet
EndFunc   ;==>__EMB_BitmapCreateDIB
