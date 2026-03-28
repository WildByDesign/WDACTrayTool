#include-once
#include <GuiMenu.au3>
#include <GUIConstantsEx.au3>
#include <APIGdiConstants.au3>
#include <WindowsNotifsConstants.au3>
#include <FontConstants.au3>
#include "GUIDarkAPI.au3"

; #INDEX# =======================================================================================================================
; Title .........: GUIDarkMenu UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: UDF library for applying dark theme to menubar
; Author(s) .....: WildByDesign (including previous code from ahmet, argumentum, UEZ)
; Version .......: 1.0.0
; Note ..........: Window messages used: WM_DRAWITEM, WM_ACTIVATE, WM_MEASUREITEM, WM_WINDOWPOSCHANGED
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _GUIDarkMenu_Register
; _GUIDarkMenu_SetColors
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __GUIDarkMenu_WM_DRAWITEM
; __GUIDarkMenu_WM_ACTIVATE
; __GUIDarkMenu_WM_MEASUREITEM
; __GUIDarkMenu_WM_WINDOWPOSCHANGED
; __GUIDarkMenu_CalcMenuWidth
; __GUIDarkMenu_GetTopMenuItems
; __GUIDarkMenu_PaintWhiteLine
; __GUIDarkMenu_MenuBarBKColor
; __GUIDarkMenu_ColorToCOLORREF
; __GUIDarkMenu_GUICtrlGetFont
; __GUIDarkMenu_GUIGetFontSize
; __GUIDarkMenu_CreateFont
; __GUIDarkMenu_WndProc
; __GUIDarkMenu_OnExit
; ===============================================================================================================================

; #GLOBAL VARIABLES# ============================================================================================================
Global $COLOR_BG_DARK = 0x121212
Global $COLOR_TEXT_LIGHT = 0xE0E0E0
Global $COLOR_CONTROL_BG = 0x202020
Global $COLOR_BORDER = 0x3F3F3F
Global $COLOR_MENU_BG = __WinAPI_ColorAdjustLuma($COLOR_BG_DARK, 5)
Global $COLOR_MENU_HOT = __WinAPI_ColorAdjustLuma($COLOR_MENU_BG, 20)
Global $COLOR_MENU_SEL = __WinAPI_ColorAdjustLuma($COLOR_MENU_BG, 10)
Global $COLOR_MENU_TEXT = $COLOR_TEXT_LIGHT
; ===============================================================================================================================

; #GLOBAL CONSTANTS# ============================================================================================================
Global Const $ODT_MENU = 1
Global Const $ODS_SELECTED = 0x0001
Global Const $ODS_DISABLED = 0x0004
Global Const $ODS_HOTLIGHT = 0x0040
; ===============================================================================================================================

; #INTERNAL_USE_ONLY GLOBAL VARIABLES # =========================================================================================
Global $g_aMenuText = []
Global $g_iDpi = 100
Global $g_hGui
Global $g_hMenuFont = 0
Global $g_hMenuProc = 0, $g_hPrevProc = 0
; ===============================================================================================================================

OnAutoItExitRegister("__GUIDarkMenu_OnExit")

Func __GUIDarkMenu_OnExit()
	If $g_hPrevProc Then __WinAPI_SetWindowLong($g_hGui, $GWL_WNDPROC, $g_hPrevProc)
	If $g_hMenuProc Then DllCallbackFree($g_hMenuProc)
	If $g_hMenuFont Then __WinAPI_DeleteObject($g_hMenuFont)
EndFunc   ;==>__GUIDarkMenu_OnExit

Func __GUIDarkMenu_WM_MEASUREITEM($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam
	Local $tagMEASUREITEM = "uint CtlType;uint CtlID;uint itemID;uint itemWidth;uint itemHeight;ulong_ptr itemData"
	Local $t = DllStructCreate($tagMEASUREITEM, $lParam)
	If Not IsDllStruct($t) Then Return $GUI_RUNDEFMSG

	If $t.CtlType <> $ODT_MENU Then Return $GUI_RUNDEFMSG

	Local $itemID = $t.itemID

	; itemID is the control ID, not the position!
	; We must derive the position from the itemID
	Local $iPos = -1
	For $i = 0 To UBound($g_aMenuText) - 1
		If $itemID = $g_aMenuText[$i][0] Then
			$iPos = $i
			ExitLoop
		EndIf
	Next

	; Fallback: try the itemID directly
	If $iPos < 0 Then $iPos = $itemID
	If $iPos < 0 Or $iPos >= UBound($g_aMenuText) Then $iPos = 0

	Local $sText = $g_aMenuText[$iPos][1]

	; Calculate text dimensions
	Local $hDC = __WinAPI_GetDC($hWnd)
	__WinAPI_SelectObject($hDC, $g_hMenuFont)
	Local $tSize = __WinAPI_GetTextExtentPoint32($hDC, $sText)
	Local $iTextWidth = $tSize.X
	Local $iTextHeight = $tSize.Y

	__WinAPI_ReleaseDC($hWnd, $hDC)

	; Set dimensions with padding (with high DPI)
	$t.itemWidth = __GUIDarkMenu_CalcMenuWidth($g_iDpi, $iTextWidth)
	$t.itemHeight = $iTextHeight + 1

	Return $GUI_RUNDEFMSG
EndFunc   ;==>__GUIDarkMenu_WM_MEASUREITEM

Func __GUIDarkMenu_CalcMenuWidth($g_iDpi, $iTextWidth)
	If $g_iDpi < 100 Or $g_iDpi > 400 Then
		Return $iTextWidth - 5
	EndIf

	Local Const $iSteps = Int(($g_iDpi - 100) / 25)
	Return $iTextWidth - (4 * $iSteps)
EndFunc   ;==>__GUIDarkMenu_CalcMenuWidth

Func __GUIDarkMenu_WM_DRAWITEM($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam
	Local Const $SM_CXDLGFRAME = 7
	Local $tagDRAWITEM = "uint CtlType;uint CtlID;uint itemID;uint itemAction;uint itemState;ptr hwndItem;handle hDC;" & _
			"long left;long top;long right;long bottom;ulong_ptr itemData"
	Local $t = DllStructCreate($tagDRAWITEM, $lParam)
	If Not IsDllStruct($t) Then Return $GUI_RUNDEFMSG

	If $t.CtlType <> $ODT_MENU Then Return $GUI_RUNDEFMSG

	Local $hDC = $t.hDC
	Local $left = $t.left
	Local $top = $t.top
	Local $right = $t.right
	Local $bottom = $t.bottom
	Local $state = $t.itemState
	Local $itemID = $t.itemID

	; convert itemID to position
	Local $iPos = -1
	For $i = 0 To UBound($g_aMenuText) - 1
		If $itemID = $g_aMenuText[$i][0] Then
			$iPos = $i
			ExitLoop
		EndIf
	Next

	If $iPos < 0 Then $iPos = $itemID
	If $iPos < 0 Or $iPos >= UBound($g_aMenuText) Then $iPos = 0

	Local $sText = $g_aMenuText[$iPos][1]
	$sText = StringReplace($sText, "&", "")

	; Colors
	Local $clrBG = __GUIDarkMenu_ColorToCOLORREF($COLOR_MENU_BG)
	Local $clrSel = __GUIDarkMenu_ColorToCOLORREF($COLOR_MENU_SEL)
	Local $clrText = __GUIDarkMenu_ColorToCOLORREF($COLOR_MENU_TEXT)

	Static $iDrawCount = 0
	Static $bFullBarDrawn = False

	; Count how many items were drawn in this "draw cycle"
	$iDrawCount += 1

	; argumentum ; pre-declare all the "Local" in those IF-THEN that could be needed
	Local $tClient, $iFullWidth, $tFullMenuBar, $hFullBrush
	Local $tEmptyArea, $hEmptyBrush

	; If we are at the first item AND the bar has not yet been drawn
	If $iPos = 0 And Not $bFullBarDrawn Then
		; Get the full window width
		$tClient = __WinAPI_GetClientRect($hWnd)
		$iFullWidth = $tClient.right

		; Fill the entire menu bar
		$tFullMenuBar = DllStructCreate($tagRECT)
		With $tFullMenuBar
			.left = 0
			.top = $top - 1
			.right = $iFullWidth + 3
			.bottom = $bottom
		EndWith

		$hFullBrush = __WinAPI_CreateSolidBrush($clrBG)
		__WinAPI_FillRect($hDC, $tFullMenuBar, $hFullBrush)
		__WinAPI_DeleteObject($hFullBrush)
	EndIf

	; After drawing all items, mark as "drawn"
	If $iDrawCount >= UBound($g_aMenuText) Then
		$bFullBarDrawn = True
		$iDrawCount = 0
	EndIf

	; Draw background for the area AFTER the last menu item
	If $iPos = (UBound($g_aMenuText) - 1) Then ; Last menu
		$tClient = __WinAPI_GetClientRect($hWnd)
		$iFullWidth = $tClient.right

		; Fill only the area to the RIGHT of the last menu item
		If $right < $iFullWidth Then
			$tEmptyArea = DllStructCreate($tagRECT)
			With $tEmptyArea
				.left = $right
				.top = $top ;        argumentum ; replace magic numbers with it's parameter name when possible
				.right = $iFullWidth + __WinAPI_GetSystemMetrics($SM_CXDLGFRAME) ; 7 = $SM_CXDLGFRAME
				.bottom = $bottom
			EndWith

			$hEmptyBrush = __WinAPI_CreateSolidBrush($clrBG)
			__WinAPI_FillRect($hDC, $tEmptyArea, $hEmptyBrush)
			__WinAPI_DeleteObject($hEmptyBrush)
		EndIf
	EndIf

	; Draw item background (selected = lighter)
	Local $bSelected = BitAND($state, $ODS_SELECTED)
	Local $bHot = BitAND($state, $ODS_HOTLIGHT)
	Local $hBrush

	If $bSelected Then
		$hBrush = __WinAPI_CreateSolidBrush($clrSel)
	ElseIf $bHot Then
		$hBrush = __WinAPI_CreateSolidBrush($COLOR_MENU_HOT)
	Else
		$hBrush = __WinAPI_CreateSolidBrush($clrBG)
	EndIf

	Local $tItemRect = DllStructCreate($tagRECT)
	With $tItemRect
		.left = $left
		.top = $top
		.right = $right
		.bottom = $bottom
	EndWith

	__WinAPI_FillRect($hDC, $tItemRect, $hBrush)
	__WinAPI_DeleteObject($hBrush)

	; Setup font
	__WinAPI_SelectObject($hDC, $g_hMenuFont)

	__WinAPI_SetBkMode($hDC, $TRANSPARENT)
	__WinAPI_SetTextColor($hDC, $clrText)

	; Draw text
	Local $tTextRect = DllStructCreate($tagRECT)
	With $tTextRect
		.left = $left + 10
		.top = $top + 4
		.right = $right - 10
		.bottom = $bottom - 4
	EndWith

	DllCall($hUser32Dll, "int", "DrawTextW", "handle", $hDC, "wstr", $sText, "int", -1, "ptr", _
			DllStructGetPtr($tTextRect), "uint", BitOR($DT_SINGLELINE, $DT_VCENTER, $DT_LEFT))

	;If $hOldFont Then __WinAPI_SelectObject($hDC, $hOldFont)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>__GUIDarkMenu_WM_DRAWITEM

Func _GUIDarkMenu_Register($hWnd)
	$g_hGui = $hWnd

	$g_hMenuProc = DllCallbackRegister('__GUIDarkMenu_WndProc', 'ptr', 'hwnd;uint;wparam;lparam')
	$g_hPrevProc = __WinAPI_SetWindowLong($g_hGui, $GWL_WNDPROC, DllCallbackGetPtr($g_hMenuProc))

	; get top menu handle
	Local $hMenu = _GUICtrlMenu_GetMenu($hWnd)
	If Not $hMenu Then Return False

	; Create font
	;$g_hMenuFont = __GUIDarkMenu_CreateFont("Segoe UI", __GUIDarkMenu_GUIGetFontSize()[0])
	$g_hMenuFont = __GUIDarkMenu_CreateFont("Segoe UI", 9)
	;$g_hMenuFont = __GUIDarkMenu_CreateFont("Segoe UI")

	; get window DPI for measurement adjustments
	$g_iDpi = Round(__WinAPI_GetDpiForWindow($g_hGui) / 96, 2) * 100
	If @error Then $g_iDpi = 100

	$g_aMenuText = __GUIDarkMenu_GetTopMenuItems($g_hGui)

	For $i = 0 To UBound($g_aMenuText) - 1
		_GUICtrlMenu_SetItemType($hMenu, $i, $MFT_OWNERDRAW, True)
	Next
	__GUIDarkMenu_MenuBarBKColor($hMenu, $COLOR_MENU_BG)
EndFunc   ;==>_GUIDarkMenu_Register

Func __GUIDarkMenu_GetTopMenuItems($hWnd)
	Local $iItemID = 10000
	Local $hMenu = _GUICtrlMenu_GetMenu($hWnd)
	Local $nItem = _GUICtrlMenu_GetItemCount($hMenu)
	Local $aList[$nItem][2], $tInfo
	Local $tText, $iLen
	For $i = 0 To $nItem - 1
		$tInfo = _GUICtrlMenu_GetItemInfo($hMenu, $i)
		If Not $tInfo.ID Then
			_GUICtrlMenu_SetItemID($hMenu, $i, $iItemID)
			$aList[$i][0] = _GUICtrlMenu_GetItemID($hMenu, $i)
			$iItemID += 1
		Else
			$aList[$i][0] = $tInfo.ID
		EndIf
		;$aList[$i][1] = _GUICtrlMenu_GetItemText($hMenu, $i)
		; retrieve text via GetMenuStringW (works better than _GUICtrlMenu_GetItemText)
		$tText = DllStructCreate("wchar s[256]")
		$iLen = DllCall($hUser32Dll, "int", "GetMenuStringW", _
				"handle", $hMenu, _
				"uint", $i, _
				"struct*", $tText, _
				"int", 255, _
				"uint", $MF_BYPOSITION)

		If IsArray($iLen) And $iLen[0] > 0 Then
			$aList[$i][1] = $tText.s
		Else
			$aList[$i][1] = ""
		EndIf
	Next
	Return $aList
EndFunc   ;==>__GUIDarkMenu_GetTopMenuItems

Func _GUIDarkMenu_SetColors($hWnd, $MenuBG, $MenuHot, $MenuSel, $MenuText)
	Local $hMenu = _GUICtrlMenu_GetMenu($hWnd)
	$COLOR_MENU_BG = $MenuBG
	$COLOR_MENU_HOT = $MenuHot
	$COLOR_MENU_SEL = $MenuSel
	$COLOR_MENU_TEXT = $MenuText
	; redraw menubar background area
	__GUIDarkMenu_MenuBarBKColor($hMenu, $COLOR_MENU_BG)
	; redraw menubar and force refresh
	_GUICtrlMenu_DrawMenuBar($hWnd)
	__WinAPI_RedrawWindow($hWnd, 0, 0, BitOR($RDW_INVALIDATE, $RDW_UPDATENOW))
EndFunc   ;==>_GUIDarkMenu_SetColors

Func __GUIDarkMenu_WM_WINDOWPOSCHANGED($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG
	__GUIDarkMenu_PaintWhiteLine($hWnd)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>__GUIDarkMenu_WM_WINDOWPOSCHANGED

Func __GUIDarkMenu_PaintWhiteLine($hWnd)
	Local $rcClient = __WinAPI_GetClientRect($hWnd)

	DllCall($hUser32Dll, "int", "MapWindowPoints", _
			"hwnd", $hWnd, _ ; hWndFrom
			"hwnd", 0, _ ; hWndTo
			"ptr", DllStructGetPtr($rcClient), _
			"uint", 2)   ;number of points - 2 for RECT structure

	If @error Then
		;MsgBox($MB_ICONERROR, "Error", @error)
		Exit
	EndIf

	Local $rcWindow = __WinAPI_GetWindowRect($hWnd)

	__WinAPI_OffsetRect($rcClient, -$rcWindow.left, -$rcWindow.top)

	Local $rcAnnoyingLine = DllStructCreate($tagRECT)
	$rcAnnoyingLine.left = $rcClient.left
	$rcAnnoyingLine.top = $rcClient.top
	$rcAnnoyingLine.right = $rcClient.right
	$rcAnnoyingLine.bottom = $rcClient.bottom

	$rcAnnoyingLine.bottom = $rcAnnoyingLine.top
	$rcAnnoyingLine.top = $rcAnnoyingLine.top - 1

	Local $hRgn = __WinAPI_CreateRectRgn(0, 0, 8000, 8000)

	Local $hDC = __WinAPI_GetDCEx($hWnd, $hRgn, BitOR($DCX_WINDOW, $DCX_INTERSECTRGN))
	Local $hFullBrush = __WinAPI_CreateSolidBrush($COLOR_MENU_BG)
	__WinAPI_FillRect($hDC, $rcAnnoyingLine, $hFullBrush)
	__WinAPI_ReleaseDC($hWnd, $hDC)
	__WinAPI_DeleteObject($hFullBrush)

EndFunc   ;==>__GUIDarkMenu_PaintWhiteLine

Func __GUIDarkMenu_WM_ACTIVATE($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	If $hWnd <> $g_hGui Then Return $GUI_RUNDEFMSG
	__GUIDarkMenu_PaintWhiteLine($hWnd)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>__GUIDarkMenu_WM_ACTIVATE

Func __GUIDarkMenu_MenuBarBKColor($hMenu, $nColor)
	Local $tInfo, $aResult
	Local $hBrush = DllCall($hGdi32Dll, 'hwnd', 'CreateSolidBrush', 'int', $nColor)
	If @error Then Return
	;$tInfo = DllStructCreate("int Size;int Mask;int Style;int YMax;int hBack;int ContextHelpID;ptr MenuData")
	$tInfo = DllStructCreate("int Size;int Mask;int Style;int YMax;handle hBack;int ContextHelpID;ptr MenuData")
	DllStructSetData($tInfo, "Mask", 2)
	DllStructSetData($tInfo, "hBack", $hBrush[0])
	DllStructSetData($tInfo, "Size", DllStructGetSize($tInfo))
	$aResult = DllCall($hUser32Dll, "int", "SetMenuInfo", "hwnd", $hMenu, "ptr", DllStructGetPtr($tInfo))
	Return $aResult[0] <> 0
EndFunc   ;==>__GUIDarkMenu_MenuBarBKColor

Func __GUIDarkMenu_ColorToCOLORREF($iColor) ;RGB to BGR
	Local $iR = BitAND(BitShift($iColor, 16), 0xFF)
	Local $iG = BitAND(BitShift($iColor, 8), 0xFF)
	Local $iB = BitAND($iColor, 0xFF)
	Return BitOR(BitShift($iB, -16), BitShift($iG, -8), $iR)
EndFunc   ;==>__GUIDarkMenu_ColorToCOLORREF

Func __GUIDarkMenu_GUICtrlGetFont($hWnd)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local Const $LOGPIXELSY = 90
	Local $aFont[6], $hDC = __WinAPI_GetDC($hWnd)
	Local $hFont = _SendMessage($hWnd, $WM_GETFONT), $tFont = DllStructCreate($tagLOGFONT)

	__WinAPI_GetObject($hFont, DllStructGetSize($tFont), $tFont)

	$aFont[0] = Round(-(($tFont.Height * 72) / __WinAPI_GetDeviceCaps($hDC, $LOGPIXELSY)) / 0.5) * 0.5
	$aFont[1] = $tFont.Weight
	$aFont[2] = BitOR(2 * ($tFont.Italic <> 0), 4 * ($tFont.Underline <> 0), 8 * ($tFont.Strikeout) <> 0)
	$aFont[3] = $tFont.FaceName
	$aFont[4] = $tFont.Quality
	$aFont[5] = $hFont

	__WinAPI_ReleaseDC($hWnd, $hDC)
	Return $aFont
EndFunc   ;==>__GUIDarkMenu_GUICtrlGetFont

Func __GUIDarkMenu_GUIGetFontSize()
	Local $idTest = GUICtrlCreateLabel("Test", -100, -100, -1, -1)
	Local $aFont = __GUIDarkMenu_GUICtrlGetFont($idTest)
	GUICtrlDelete($idTest)
	;Return Round($aTest[0] / 0.5) * 0.5
	Return $aFont
EndFunc   ;==>__GUIDarkMenu_GUIGetFontSize

Func __GUIDarkMenu_CreateFont($sFontName, $nHeight = 9, $nWidth = 400)
	Local $stFontName = DllStructCreate("char[260]")
	DllStructSetData($stFontName, 1, $sFontName)
	Local $hDC = __WinAPI_GetDC(0)        ; Get the Desktops DC
	Local $nPixel = __WinAPI_GetDeviceCaps($hDC, 90)
	$nHeight = 0 - _WinAPI_MulDiv($nHeight, $nPixel, 72)
	__WinAPI_ReleaseDC(0, $hDC)
	Local $hFont = __WinAPI_CreateFont($nHeight, 0, 0, 0, $nWidth, False, False, False, _
			$DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $PROOF_QUALITY, $DEFAULT_PITCH, $sFontName)
	Return $hFont
EndFunc   ;==>__GUIDarkMenu_CreateFont

Func __GUIDarkMenu_WndProc($hWnd, $iMsg, $wParam, $lParam)
	Local $sContinue = $GUI_RUNDEFMSG
	Switch $iMsg
		Case $WM_ACTIVATE
			$sContinue = __GUIDarkMenu_WM_ACTIVATE($hWnd, $iMsg, $wParam, $lParam)
		Case $WM_DRAWITEM
			$sContinue = __GUIDarkMenu_WM_DRAWITEM($hWnd, $iMsg, $wParam, $lParam)
		Case $WM_MEASUREITEM
			$sContinue = __GUIDarkMenu_WM_MEASUREITEM($hWnd, $iMsg, $wParam, $lParam)
		Case $WM_WINDOWPOSCHANGED
			$sContinue = __GUIDarkMenu_WM_WINDOWPOSCHANGED($hWnd, $iMsg, $wParam, $lParam)
	EndSwitch
	If $sContinue = $GUI_RUNDEFMSG Then Return __WinAPI_CallWindowProc($g_hPrevProc, $hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>__GUIDarkMenu_WndProc
