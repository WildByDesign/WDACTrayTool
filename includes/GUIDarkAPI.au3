#include-once
#include <StructureConstants.au3>
#include <AutoItConstants.au3>
#include <WinAPIGdiInternals.au3>

; #INDEX# =======================================================================================================================
; Title .........: GUIDarkAPI UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: API support library for GUIDarkTheme UDF
; Author(s) .....: WildByDesign (including previous code from NoNameCode, argumentum, UEZ)
; Version .......: 1.0.0
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; TODO: Need to fill list of all functions.
; ===============================================================================================================================

; #GLOBAL VARIABLES# ============================================================================================================
Global $hGdi32Dll = DllOpen("gdi32.dll")
Global $hUser32Dll = DllOpen("user32.dll")
Global $hKernel32Dll = DllOpen("kernel32.dll")
Global $hShlwapiDll = DllOpen("shlwapi.dll")
Global $hUxthemeDll = DllOpen("uxtheme.dll")
; ===============================================================================================================================

; #GLOBAL CONSTANTS# ============================================================================================================
Global Const $APPMODE_DEFAULT = 0
Global Const $APPMODE_ALLOWDARK = 1
Global Const $APPMODE_FORCEDARK = 2
Global Const $APPMODE_FORCELIGHT = 3
Global Const $APPMODE_MAX = 4
; ===============================================================================================================================

OnAutoItExitRegister(_UnloadDLLs)

Func _UnloadDLLs()
	If $hGdi32Dll Then DllClose($hGdi32Dll)
	If $hUser32Dll Then DllClose($hUser32Dll)
	If $hKernel32Dll Then DllClose($hKernel32Dll)
	If $hShlwapiDll Then DllClose($hShlwapiDll)
	If $hUxthemeDll Then DllClose($hUxthemeDll)
EndFunc   ;==>_UnloadDLLs

Func __WinAPI_GetDpiForWindow($hWnd)
	Local $aResult = DllCall($hUser32Dll, "uint", "GetDpiForWindow", "hwnd", $hWnd) ;requires Win10 v1607+ / no server support
	If Not IsArray($aResult) Or @error Then Return SetError(1, @extended, 0)
	If Not $aResult[0] Then Return SetError(2, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>__WinAPI_GetDpiForWindow

Func __WinAPI_GetWindowRect($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($hUser32Dll, "bool", "GetWindowRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>__WinAPI_GetWindowRect

Func __WinAPI_IsWindowVisible($hWnd)
	Local $aCall = DllCall($hUser32Dll, "bool", "IsWindowVisible", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_IsWindowVisible

Func __WinAPI_GetWindow($hWnd, $iCmd)
	Local $aCall = DllCall($hUser32Dll, "hwnd", "GetWindow", "hwnd", $hWnd, "uint", $iCmd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetWindow

Func __WinAPI_CallWindowProc($pPrevWndFunc, $hWnd, $iMsg, $wParam, $lParam)
	Local $aCall = DllCall($hUser32Dll, "lresult", "CallWindowProc", "ptr", $pPrevWndFunc, "hwnd", $hWnd, "uint", $iMsg, _
			"wparam", $wParam, "lparam", $lParam)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CallWindowProc

Func __WinAPI_EndPaint($hWnd, ByRef $tPAINTSTRUCT)
	Local $aCall = DllCall($hUser32Dll, 'bool', 'EndPaint', 'hwnd', $hWnd, 'struct*', $tPAINTSTRUCT)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_EndPaint

Func __WinAPI_RedrawWindow($hWnd, $tRECT = 0, $hRegion = 0, $iFlags = 5)
	Local $aCall = DllCall($hUser32Dll, "bool", "RedrawWindow", "hwnd", $hWnd, "struct*", $tRECT, "handle", $hRegion, _
			"uint", $iFlags)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_RedrawWindow

Func __WinAPI_GetSystemMetrics($iIndex)
	Local $aCall = DllCall($hUser32Dll, "int", "GetSystemMetrics", "int", $iIndex)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetSystemMetrics

Func __WinAPI_GetClientRect($hWnd)
	Local $tRECT = DllStructCreate($tagRECT)
	Local $aCall = DllCall($hUser32Dll, "bool", "GetClientRect", "hwnd", $hWnd, "struct*", $tRECT)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tRECT
EndFunc   ;==>__WinAPI_GetClientRect

Func __WinAPI_OffsetRect(ByRef $tRECT, $iDX, $iDY)
	Local $aCall = DllCall($hUser32Dll, 'bool', 'OffsetRect', 'struct*', $tRECT, 'int', $iDX, 'int', $iDY)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_OffsetRect

Func __WinAPI_GetDCEx($hWnd, $hRgn, $iFlags)
	Local $aCall = DllCall($hUser32Dll, 'handle', 'GetDCEx', 'hwnd', $hWnd, 'handle', $hRgn, 'dword', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetDCEx

Func __WinAPI_FillRect($hDC, $tRECT, $hBrush)
	Local $aCall
	If IsPtr($hBrush) Then
		$aCall = DllCall($hUser32Dll, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "handle", $hBrush)
	Else
		$aCall = DllCall($hUser32Dll, "int", "FillRect", "handle", $hDC, "struct*", $tRECT, "dword_ptr", $hBrush)
	EndIf
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_FillRect

Func __WinAPI_ReleaseDC($hWnd, $hDC)
	Local $aCall = DllCall($hUser32Dll, "int", "ReleaseDC", "hwnd", $hWnd, "handle", $hDC)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_ReleaseDC

Func __WinAPI_DefWindowProc($hWnd, $iMsg, $wParam, $lParam)
	Local $aCall = DllCall($hUser32Dll, "lresult", "DefWindowProc", "hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, _
			"lparam", $lParam)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DefWindowProc

Func __WinAPI_GetParent($hWnd)
	Local $aCall = DllCall($hUser32Dll, "hwnd", "GetParent", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetParent

Func __WinAPI_GetDC($hWnd)
	Local $aCall = DllCall($hUser32Dll, "handle", "GetDC", "hwnd", $hWnd)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetDC

Func __WinAPI_BeginPaint($hWnd, ByRef $tPAINTSTRUCT)
	Local Const $tagPAINTSTRUCT = 'hwnd hDC;int fErase;dword rPaint[4];int fRestore;int fIncUpdate;byte rgbReserved[32]'
	$tPAINTSTRUCT = DllStructCreate($tagPAINTSTRUCT)
	Local $aCall = DllCall($hUser32Dll, 'handle', 'BeginPaint', 'hwnd', $hWnd, 'struct*', $tPAINTSTRUCT)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_BeginPaint

Func __WinAPI_GetClassName($hWnd)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $aCall = DllCall($hUser32Dll, "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
	If @error Or Not $aCall[0] Then Return SetError(@error, @extended, '')

	Return SetExtended($aCall[0], $aCall[2])
EndFunc   ;==>__WinAPI_GetClassName

Func __WinAPI_InvertRect($hDC, ByRef $tRECT)
	Local $aCall = DllCall($hUser32Dll, 'bool', 'InvertRect', 'handle', $hDC, 'struct*', $tRECT)
	If @error Then Return SetError(@error, @extended, False)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_InvertRect

Func __WinAPI_InvalidateRect($hWnd, $tRECT = 0, $bErase = True)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $aCall = DllCall($hUser32Dll, "bool", "InvalidateRect", "hwnd", $hWnd, "struct*", $tRECT, "bool", $bErase)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_InvalidateRect

Func __WinAPI_GetFocus()
	Local $aCall = DllCall($hUser32Dll, "hwnd", "GetFocus")
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetFocus

Func __WinAPI_GetWindowLong($hWnd, $iIndex)
	Local $sFuncName = "GetWindowLongW"
	If @AutoItX64 Then $sFuncName = "GetWindowLongPtrW"
	Local $aCall = DllCall($hUser32Dll, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetWindowLong

Func __WinAPI_SetWindowLong($hWnd, $iIndex, $iValue)
	__WinAPI_SetLastError(0) ; as suggested in MSDN
	Local $sFuncName = "SetWindowLongW"
	If @AutoItX64 Then $sFuncName = "SetWindowLongPtrW"
	Local $aCall = DllCall($hUser32Dll, "long_ptr", $sFuncName, "hwnd", $hWnd, "int", $iIndex, "long_ptr", $iValue)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetWindowLong

Func __WinAPI_SetWindowPos($hWnd, $hAfter, $iX, $iY, $iCX, $iCY, $iFlags)
	Local $aCall = DllCall($hUser32Dll, "bool", "SetWindowPos", "hwnd", $hWnd, "hwnd", $hAfter, "int", $iX, "int", $iY, _
			"int", $iCX, "int", $iCY, "uint", $iFlags)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetWindowPos

Func __WinAPI_SetClassLongEx($hWnd, $iIndex, $iNewLong)
	Local $aCall
	If @AutoItX64 Then
		$aCall = DllCall($hUser32Dll, 'ulong_ptr', 'SetClassLongPtrW', 'hwnd', $hWnd, 'int', $iIndex, 'long_ptr', $iNewLong)
	Else
		$aCall = DllCall($hUser32Dll, 'dword', 'SetClassLongW', 'hwnd', $hWnd, 'int', $iIndex, 'long', $iNewLong)
	EndIf
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $aCall[0] Then Return SetError(1000, 0, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetClassLongEx

Func __WinAPI_FindWindowEx($hParent, $hAfter, $sClass, $sTitle = "")
	Local $ret = DllCall($hUser32Dll, "hwnd", "FindWindowExW", "hwnd", $hParent, "hwnd", $hAfter, "wstr", $sClass, "wstr", $sTitle)
	If @error Or Not IsArray($ret) Then Return 0
	Return $ret[0]
EndFunc   ;==>__WinAPI_FindWindowEx

Func __WinAPI_EnumProcessWindows($iPID = 0, $bVisible = True)
	Local $aThreads = __WinAPI_EnumProcessThreads($iPID)
	If @error Then Return SetError(@error, @extended, 0)

	Local $hEnumProc = DllCallbackRegister('__EnumWindowsProc', 'bool', 'hwnd;lparam')

	Dim $__g_vEnum[101][2] = [[0]]
	For $i = 1 To $aThreads[0]
		DllCall($hUser32Dll, 'bool', 'EnumThreadWindows', 'dword', $aThreads[$i], 'ptr', DllCallbackGetPtr($hEnumProc), _
				'lparam', $bVisible)
		If @error Then
			ExitLoop
		EndIf
	Next
	DllCallbackFree($hEnumProc)
	If Not $__g_vEnum[0][0] Then Return SetError(11, 0, 0)

	___Inc($__g_vEnum, -1)
	Return $__g_vEnum
EndFunc   ;==>__WinAPI_EnumProcessWindows

Func __WinAPI_EnumChildWindows($hWnd, $bVisible = True)
	If Not __WinAPI_GetWindow($hWnd, 5) Then Return SetError(2, 0, 0) ; $GW_CHILD

	Local $hEnumProc = DllCallbackRegister('__EnumWindowsProc', 'bool', 'hwnd;lparam')

	Dim $__g_vEnum[101][2] = [[0]]
	DllCall($hUser32Dll, 'bool', 'EnumChildWindows', 'hwnd', $hWnd, 'ptr', DllCallbackGetPtr($hEnumProc), 'lparam', $bVisible)
	If @error Or Not $__g_vEnum[0][0] Then
		$__g_vEnum = @error + 10
	EndIf
	DllCallbackFree($hEnumProc)
	If $__g_vEnum Then Return SetError($__g_vEnum, 0, 0)

	___Inc($__g_vEnum, -1)
	Return $__g_vEnum
EndFunc   ;==>__WinAPI_EnumChildWindows

Func ___Inc(ByRef $aData, $iIncrement = 100)
	Select
		Case UBound($aData, $UBOUND_COLUMNS)
			If $iIncrement < 0 Then
				ReDim $aData[$aData[0][0] + 1][UBound($aData, $UBOUND_COLUMNS)]
			Else
				$aData[0][0] += 1
				If $aData[0][0] > UBound($aData) - 1 Then
					ReDim $aData[$aData[0][0] + $iIncrement][UBound($aData, $UBOUND_COLUMNS)]
				EndIf
			EndIf
		Case UBound($aData, $UBOUND_ROWS)
			If $iIncrement < 0 Then
				ReDim $aData[$aData[0] + 1]
			Else
				$aData[0] += 1
				If $aData[0] > UBound($aData) - 1 Then
					ReDim $aData[$aData[0] + $iIncrement]
				EndIf
			EndIf
		Case Else
			Return 0
	EndSelect
	Return 1
EndFunc   ;==>___Inc

Func __WinAPI_SetLastError($iErrorCode, Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
	DllCall($hKernel32Dll, "none", "SetLastError", "dword", $iErrorCode)
	Return SetError($_iCallerError, $_iCallerExtended, Null)
EndFunc   ;==>__WinAPI_SetLastError

Func __WinAPI_EnumProcessThreads($iPID = 0)
	If Not $iPID Then $iPID = @AutoItPID

	Local Const $TH32CS_SNAPTHREAD = 0x00000004
	Local $hSnapshot = DllCall($hKernel32Dll, 'handle', 'CreateToolhelp32Snapshot', 'dword', $TH32CS_SNAPTHREAD, 'dword', 0)
	If @error Or Not $hSnapshot[0] Then Return SetError(@error + 10, @extended, 0)

	Local Const $tagTHREADENTRY32 = 'dword Size;dword Usage;dword ThreadID;dword OwnerProcessID;long BasePri;long DeltaPri;dword Flags'
	Local $tTHREADENTRY32 = DllStructCreate($tagTHREADENTRY32)
	Local $aRet[101] = [0]

	$hSnapshot = $hSnapshot[0]
	DllStructSetData($tTHREADENTRY32, 'Size', DllStructGetSize($tTHREADENTRY32))
	Local $aCall = DllCall($hKernel32Dll, 'bool', 'Thread32First', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
	While Not @error And $aCall[0]
		If DllStructGetData($tTHREADENTRY32, 'OwnerProcessID') = $iPID Then
			___Inc($aRet)
			$aRet[$aRet[0]] = DllStructGetData($tTHREADENTRY32, 'ThreadID')
		EndIf
		$aCall = DllCall($hKernel32Dll, 'bool', 'Thread32Next', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
	WEnd
	DllCall($hKernel32Dll, "bool", "CloseHandle", "handle", $hSnapshot)
	If Not $aRet[0] Then Return SetError(1, 0, 0)

	___Inc($aRet, -1)
	Return $aRet
EndFunc   ;==>__WinAPI_EnumProcessThreads

Func __SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
	Local $aCall = DllCall($hUser32Dll, $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
	If @error Then Return SetError(@error, @extended, "")
	If $iReturn >= 0 And $iReturn <= 4 Then Return $aCall[$iReturn]
	Return $aCall
EndFunc   ;==>__SendMessage

Func __WinAPI_CreateRectRgn($iLeftRect, $iTopRect, $iRightRect, $iBottomRect)
	Local $aCall = DllCall($hGdi32Dll, "handle", "CreateRectRgn", "int", $iLeftRect, "int", $iTopRect, "int", $iRightRect, _
			"int", $iBottomRect)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateRectRgn

Func __WinAPI_CreateSolidBrush($iColor)
	Local $aCall = DllCall($hGdi32Dll, "handle", "CreateSolidBrush", "INT", $iColor)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateSolidBrush

Func __WinAPI_DeleteObject($hObject)
	Local $aCall = DllCall($hGdi32Dll, "bool", "DeleteObject", "handle", $hObject)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DeleteObject

Func __WinAPI_CreatePen($iPenStyle, $iWidth, $iColor)
	Local $aCall = DllCall($hGdi32Dll, "handle", "CreatePen", "int", $iPenStyle, "int", $iWidth, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreatePen

Func __WinAPI_MoveTo($hDC, $iX, $iY)
	Local $aCall = DllCall($hGdi32Dll, "bool", "MoveToEx", "handle", $hDC, "int", $iX, "int", $iY, "ptr", 0)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_MoveTo

Func __WinAPI_LineTo($hDC, $iX, $iY)
	Local $aCall = DllCall($hGdi32Dll, "bool", "LineTo", "handle", $hDC, "int", $iX, "int", $iY)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_LineTo

Func __WinAPI_BitBlt($hDestDC, $iXDest, $iYDest, $iWidth, $iHeight, $hSrcDC, $iXSrc, $iYSrc, $iROP)
	Local $aCall = DllCall($hGdi32Dll, "bool", "BitBlt", "handle", $hDestDC, "int", $iXDest, "int", $iYDest, "int", $iWidth, _
			"int", $iHeight, "handle", $hSrcDC, "int", $iXSrc, "int", $iYSrc, "dword", $iROP)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_BitBlt

Func __WinAPI_DeleteDC($hDC)
	Local $aCall = DllCall($hGdi32Dll, "bool", "DeleteDC", "handle", $hDC)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_DeleteDC

Func __WinAPI_SelectObject($hDC, $hGDIObj)
	Local $aCall = DllCall($hGdi32Dll, "handle", "SelectObject", "handle", $hDC, "handle", $hGDIObj)
	If @error Then Return SetError(@error, @extended, False)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SelectObject

Func __WinAPI_SetBkMode($hDC, $iBkMode)
	Local $aCall = DllCall($hGdi32Dll, "int", "SetBkMode", "handle", $hDC, "int", $iBkMode)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetBkMode

Func __WinAPI_SetTextColor($hDC, $iColor)
	Local $aCall = DllCall($hGdi32Dll, "INT", "SetTextColor", "handle", $hDC, "INT", $iColor)
	If @error Then Return SetError(@error, @extended, -1)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_SetTextColor

Func __WinAPI_GetStockObject($iObject)
	Local $aCall = DllCall($hGdi32Dll, "handle", "GetStockObject", "int", $iObject)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetStockObject

Func __WinAPI_CreateFont($iHeight, $iWidth, $iEscape = 0, $iOrientn = 0, $iWeight = $__WINAPICONSTANT_FW_NORMAL, $bItalic = False, $bUnderline = False, $bStrikeout = False, $iCharset = $__WINAPICONSTANT_DEFAULT_CHARSET, $iOutputPrec = $__WINAPICONSTANT_OUT_DEFAULT_PRECIS, $iClipPrec = $__WINAPICONSTANT_CLIP_DEFAULT_PRECIS, $iQuality = $__WINAPICONSTANT_DEFAULT_QUALITY, $iPitch = $__WINAPICONSTANT_FF_DONTCARE, $sFace = 'Arial')
	Local $aCall = DllCall($hGdi32Dll, "handle", "CreateFontW", "int", $iHeight, "int", $iWidth, "int", $iEscape, _
			"int", $iOrientn, "int", $iWeight, "dword", $bItalic, "dword", $bUnderline, "dword", $bStrikeout, _
			"dword", $iCharset, "dword", $iOutputPrec, "dword", $iClipPrec, "dword", $iQuality, "dword", $iPitch, "wstr", $sFace)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateFont

Func __WinAPI_CreateCompatibleDC($hDC)
	Local $aCall = DllCall($hGdi32Dll, "handle", "CreateCompatibleDC", "handle", $hDC)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateCompatibleDC

Func __WinAPI_CreateCompatibleBitmap($hDC, $iWidth, $iHeight)
	Local $aCall = DllCall($hGdi32Dll, "handle", "CreateCompatibleBitmap", "handle", $hDC, "int", $iWidth, "int", $iHeight)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateCompatibleBitmap

Func __WinAPI_GetObject($hObject, $iSize, $pObject)
	Local $aCall = DllCall($hGdi32Dll, "int", "GetObjectW", "handle", $hObject, "int", $iSize, "struct*", $pObject)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetObject

Func __WinAPI_GetDeviceCaps($hDC, $iIndex)
	Local $aCall = DllCall($hGdi32Dll, "int", "GetDeviceCaps", "handle", $hDC, "int", $iIndex)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aCall[0]
EndFunc   ;==>__WinAPI_GetDeviceCaps

Func __WinAPI_GetTextExtentPoint32($hDC, $sText)
	Local $tSize = DllStructCreate($tagSIZE)
	Local $iSize = StringLen($sText)
	Local $aCall = DllCall($hGdi32Dll, "bool", "GetTextExtentPoint32W", "handle", $hDC, "wstr", $sText, "int", $iSize, "struct*", $tSize)
	If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)

	Return $tSize
EndFunc   ;==>__WinAPI_GetTextExtentPoint32

Func __WinAPI_ColorAdjustLuma($iRGB, $iPercent, $bScale = True)
	If $iRGB = -1 Then Return SetError(10, 0, -1)

	If $bScale Then
		$iPercent = Floor($iPercent * 10)
	EndIf

	Local $aCall = DllCall($hShlwapiDll, 'dword', 'ColorAdjustLuma', 'dword', ___RGB($iRGB), 'int', $iPercent, 'bool', $bScale)
	If @error Then Return SetError(@error, @extended, -1)

	Return ___RGB($aCall[0])
EndFunc   ;==>__WinAPI_ColorAdjustLuma

Func ___RGB($iColor)
	Local $__g_iRGBMode = 1
	If $__g_iRGBMode Then
		$iColor = __WinAPI_SwitchColor($iColor)
	EndIf
	Return $iColor
EndFunc   ;==>___RGB

Func __WinAPI_SwitchColor($iColor)
	If $iColor = -1 Then Return $iColor
	Return BitOR(BitAND($iColor, 0x00FF00), BitShift(BitAND($iColor, 0x0000FF), -16), BitShift(BitAND($iColor, 0xFF0000), 16))
EndFunc   ;==>__WinAPI_SwitchColor

Func __WinAPI_SetWindowTheme($hWnd, $sName = Default, $sList = Default)
	If Not IsString($sName) Then $sName = Null
	If Not IsString($sList) Then $sList = Null

	Local $sResult = DllCall($hUxthemeDll, 'long', 'SetWindowTheme', 'hwnd', $hWnd, 'wstr', $sName, 'wstr', $sList)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return 1
EndFunc   ;==>__WinAPI_SetWindowTheme

Func __WinAPI_GetThemeAppProperties()
	Local $sResult = DllCall($hUxthemeDll, 'dword', 'GetThemeAppProperties')
	If @error Then Return SetError(@error, @extended, 0)

	Return $sResult[0]
EndFunc   ;==>__WinAPI_GetThemeAppProperties

Func __WinAPI_GetThemePartSize($hTheme, $iPartID, $iStateID, $hDC, $tRECT, $iType)
	Local $tSize = DllStructCreate($tagSIZE)
	Local $sResult = DllCall($hUxthemeDll, 'long', 'GetThemePartSize', 'handle', $hTheme, 'handle', $hDC, 'int', $iPartID, _
			'int', $iStateID, 'struct*', $tRECT, 'int', $iType, 'struct*', $tSize)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return $tSize
EndFunc   ;==>__WinAPI_GetThemePartSize

Func __WinAPI_SetThemeAppProperties($iFlags)
	DllCall($hUxthemeDll, 'none', 'SetThemeAppProperties', 'dword', $iFlags)
	If @error Then Return SetError(@error, @extended, 0)

	Return 1
EndFunc   ;==>__WinAPI_SetThemeAppProperties

Func __WinAPI_DrawThemeBackground($hTheme, $iPartID, $iStateID, $hDC, $tRECT, $tCLIP = 0)
	Local $sResult = DllCall($hUxthemeDll, 'long', 'DrawThemeBackground', 'handle', $hTheme, 'handle', $hDC, 'int', $iPartID, _
			'int', $iStateID, 'struct*', $tRECT, 'struct*', $tCLIP)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return 1
EndFunc   ;==>__WinAPI_DrawThemeBackground

Func __WinAPI_CloseThemeData($hTheme)
	Local $sResult = DllCall($hUxthemeDll, 'long', 'CloseThemeData', 'handle', $hTheme)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)

	Return 1
EndFunc   ;==>__WinAPI_CloseThemeData

Func __WinAPI_OpenThemeData($hWnd, $sClass)
	Local $sResult = DllCall($hUxthemeDll, 'handle', 'OpenThemeData', 'hwnd', $hWnd, 'wstr', $sClass)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $sResult[0] Then Return SetError(1000, 0, 0)

	Return $sResult[0]
EndFunc   ;==>__WinAPI_OpenThemeData

Func __WinAPI_OpenThemeDataForDpi($hWnd, $sClass, $iDPI)
	Local $sResult = DllCall($hUxthemeDll, 'handle', 'OpenThemeDataForDpi', 'hwnd', $hWnd, 'wstr', $sClass, "uint", $iDPI)
	If @error Then Return SetError(@error, @extended, 0)
	; If Not $sResult[0] Then Return SetError(1000, 0, 0)

	Return $sResult[0]
EndFunc   ;==>__WinAPI_OpenThemeDataForDpi

Func __WinAPI_ShouldAppsUseDarkMode()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnShouldAppsUseDarkMode = 132
	Local $aResult = DllCall($hUxthemeDll, 'bool', $fnShouldAppsUseDarkMode)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>__WinAPI_ShouldAppsUseDarkMode

Func __WinAPI_AllowDarkModeForWindow($hWnd, $bAllow = True)
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnAllowDarkModeForWindow = 133
	Local $aResult = DllCall($hUxthemeDll, 'bool', $fnAllowDarkModeForWindow, 'hwnd', $hWnd, 'bool', $bAllow)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>__WinAPI_AllowDarkModeForWindow

Func __WinAPI_FlushMenuThemes()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnFlushMenuThemes = 136
	DllCall($hUxthemeDll, 'none', $fnFlushMenuThemes)
	If @error Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>__WinAPI_FlushMenuThemes

Func __WinAPI_RefreshImmersiveColorPolicyState()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnRefreshImmersiveColorPolicyState = 104
	DllCall($hUxthemeDll, 'none', $fnRefreshImmersiveColorPolicyState)
	If @error Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>__WinAPI_RefreshImmersiveColorPolicyState

Func __WinAPI_SetPreferredAppMode($PREFERREDAPPMODE)
	If @OSBuild < 18362 Then Return SetError(-1, 0, False)
	Local $fnSetPreferredAppMode = 135
	Local $aResult = DllCall($hUxthemeDll, 'int', $fnSetPreferredAppMode, 'int', $PREFERREDAPPMODE)
	If @error Then Return SetError(@error, @extended, '')
	Return $aResult[0]
EndFunc   ;==>__WinAPI_SetPreferredAppMode

Func __WinAPI_DwmSetWindowAttribute($hWnd, $iAttribute, $iData) ; #include <WinAPIGdi.au3> ; unthoughtful unrestricting mod.
	Local $aCall = DllCall('dwmapi.dll', 'long', 'DwmSetWindowAttribute', 'hwnd', $hWnd, 'dword', $iAttribute, _
			'dword*', $iData, 'dword', 4)
	If @error Then Return SetError(@error, @extended, 0)
	If $aCall[0] Then Return SetError(10, $aCall[0], 0)
	Return 1
EndFunc   ;==>__WinAPI_DwmSetWindowAttribute

Func __WinAPI_CreateRectEx($iX, $iY, $iWidth, $iHeight)
	Local $tRECT = DllStructCreate($tagRECT)
	DllStructSetData($tRECT, 1, $iX)
	DllStructSetData($tRECT, 2, $iY)
	DllStructSetData($tRECT, 3, $iX + $iWidth)
	DllStructSetData($tRECT, 4, $iY + $iHeight)

	Return $tRECT
EndFunc   ;==>__WinAPI_CreateRectEx

Func __WinAPI_Base64Decode($sB64String)
	Local $aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "ptr", 0, "dword*", 0, "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(1, 0, "")
	Local $bBuffer = DllStructCreate("byte[" & $aCrypt[5] & "]")
	$aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "struct*", $bBuffer, "dword*", $aCrypt[5], "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($bBuffer, 1)
EndFunc   ;==>__WinAPI_Base64Decode
