#include-once ; GuiCtrls_HiDpi.au3 ; v 0.2023.7.29(a) ; https://www.autoitscript.com/forum/topic/210613-guictrls_hidpi-udf-in-progress/

Global $___gn__HiDpi_ratio = 1
Global $___gi__HiDpi_DPI = 96

Func _HiDpi_GUICreate($sTitle = "", $iWidth = Default, $iHeight = Default, $iLeft = -1, $iTop = -1, $iStyle = -1, $iExStyle = -1, $hParent = 0)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $hRet = GUICreate($sTitle, $iWidth, $iHeight, $iLeft, $iTop, $iStyle, $iExStyle, $hParent) ;  Create a GUI window.
    Return SetError(@error, @extended, $hRet)
EndFunc   ;==>_HiDpi_GUICreate

#Region GUICtrlCreate__

Func _HiDpi_SplashTextOn($sTitle, $sText, $iWidth = 500, $iHeight = 400, $iLeft = Default, $iTop = Default, $iOpt = Default, $sFontname = Default, $iFontSz = 12, $iFfontWt = 400)
	__HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
	Local $iRet = SplashTextOn($sTitle, $sText, $iWidth, $iHeight, $iLeft, $iTop, $iOpt, $sFontname, $iFontSz, $iFfontWt)
	Return SetError(@error, @extended, $iRet)
EndFunc

Func _HiDpi_GUICtrlCreateAvi($sFilename, $iSubfileId, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateAvi($sFilename, $iSubfileId, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates an AVI video control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateAvi

Func _HiDpi_GUICtrlCreateButton($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateButton($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Button control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateButton

Func _HiDpi_GUICtrlCreateCheckbox($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateCheckbox($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Checkbox control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateCheckbox

Func _HiDpi_GUICtrlCreateCombo($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateCombo($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a ComboBox control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateCombo

Func _HiDpi_GUICtrlCreateDate($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateDate($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a date control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateDate

Func _HiDpi_GUICtrlCreateEdit($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateEdit($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates an Edit control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateEdit

Func _HiDpi_GUICtrlCreateGraphic($iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateGraphic($iLeft, $iTop, $iWidth, $iHeight, $iStyle) ;  Creates a Graphic control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateGraphic

Func _HiDpi_GUICtrlCreateGroup($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateGroup($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Group control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateGroup

Func _HiDpi_GUICtrlCreateIcon($sFilename, $vIconName, $iLeft, $iTop, $iWidth = 32, $iHeight = 32, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateIcon($sFilename, $vIconName, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates an Icon control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateIcon

Func _HiDpi_GUICtrlCreateInput($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateInput($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates an Input control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateInput

Func _HiDpi_GUICtrlCreateLabel($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateLabel($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a static Label control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateLabel

Func _HiDpi_GUICtrlCreateList($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateList($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a List control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateList

Func _HiDpi_GUICtrlCreateListView($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateListView($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a ListView control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateListView

Func _HiDpi_GUICtrlCreateMonthCal($sText, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateMonthCal($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a month calendar control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateMonthCal

Func _HiDpi_GUICtrlCreateObj($ObjectVar, $iLeft, $iTop, $iWidth = Default, $iHeight = Default)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateObj($ObjectVar, $iLeft, $iTop, $iWidth, $iHeight) ;  Creates an ActiveX control in the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateObj

Func _HiDpi_GUICtrlCreatePic($sFilename, $iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreatePic($sFilename, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Picture control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreatePic

Func _HiDpi_GUICtrlCreateProgress($iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateProgress($iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Progress control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateProgress

Func _HiDpi_GUICtrlCreateRadio($sText, $iLeft, $iTop = Default, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateRadio($sText, $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Radio button control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateRadio

Func _HiDpi_GUICtrlCreateSlider($iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateSlider($iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Slider control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateSlider

Func _HiDpi_GUICtrlCreateTab($iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateTab($iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a Tab control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateTab

Func _HiDpi_GUICtrlCreateTreeView($iLeft, $iTop, $iWidth = Default, $iHeight = Default, $iStyle = -1, $iExStyle = -1)
    __HiDpi_CtrlBulkNewSize($iLeft, $iTop, $iWidth, $iHeight)
    Local $iRet = GUICtrlCreateTreeView($iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle) ;  Creates a TreeView control for the GUI.
    Return SetError(@error, @extended, $iRet)
EndFunc   ;==>_HiDpi_GUICtrlCreateTreeView

#EndRegion GUICtrlCreate__

;https://learn.microsoft.com/en-us/windows/win32/hidpi/dpi-awareness-context
Func _HiDpi_Ctrl_LazyInit($iCONTEXT = -2) ; under the current state of developnment, -2 is the best option for this UDF.
	If $iCONTEXT <> Null Then
		If Not ($iCONTEXT < -1 And $iCONTEXT > -5) Then $iCONTEXT = -2
		DllCall("user32.dll", "bool", "SetProcessDpiAwarenessContext", "int", $iCONTEXT)
	EndIf
    _HiDpi_CtrlSetVarDPI( _HiDpi_GetDpiForWindow())
    _HiDpi_CtrlSetVarRatio(_HiDpi_CtrlSetVarDPI() / 96)
EndFunc   ;==>_HiDpi_Ctrl_LazyInit

Func _HiDpi_CtrlSetVarDPI($iDPI = Default)
    If $iDPI <> Default Then $___gi__HiDpi_DPI = $iDPI
    Return $___gi__HiDpi_DPI
EndFunc   ;==>_HiDpi_CtrlSetVarDPI

Func _HiDpi_CtrlSetVarRatio($nRadio = Default)
    If $nRadio <> Default Then $___gn__HiDpi_ratio = $nRadio
    Return $___gn__HiDpi_ratio
EndFunc   ;==>_HiDpi_CtrlSetVarRatio

Func _HiDpi_CtrlAdjRatio($iSize, $iSubstractRadio = 0, Const $_iCallerError = @error, Const $_iCallerExtended = @extended) ; @error and @extended are preserved on return.
    If $___gn__HiDpi_ratio = 1 Then Return SetError($_iCallerError, $_iCallerExtended, $iSize)
    If $iSubstractRadio Then
        $iSize /= $___gn__HiDpi_ratio
    Else
        $iSize *= $___gn__HiDpi_ratio
    EndIf
    Return SetError($_iCallerError, $_iCallerExtended, $iSize)
EndFunc   ;==>_HiDpi_CtrlAdjRatio

;~ Func _HiDpi_CtrlCount_Add($vRet, $iType = 2, Const $_iCallerError = @error, Const $_iCallerExtended = @extended) ; @error and @extended are preserved on return.
;~     ; working on it. Shouln't post. oops. too late.
;~     Return SetError($_iCallerError, $_iCallerExtended, $vRet)
;~ EndFunc   ;==>_HiDpi_CtrlCount_Add

Func __HiDpi_CtrlNewSize(ByRef $iSize) ; internal use
    If $___gn__HiDpi_ratio = 1 Then Return SetError(0, 0, $iSize)
    $iSize *= $___gn__HiDpi_ratio
    Return SetError(0, 1, $iSize) ; @extended = 1 ; radio is not 1 to 1 ;
EndFunc   ;==>__HiDpi_CtrlNewSize

Func __HiDpi_CtrlBulkNewSize(ByRef $iLeft, ByRef $iTop, ByRef $iWidth, ByRef $iHeight) ; internal use
    If $iLeft <> -1 And $iLeft <> Default Then __HiDpi_CtrlNewSize($iLeft)
    If $iTop <> -1 And $iTop <> Default Then __HiDpi_CtrlNewSize($iTop)
    If $iWidth <> Default Then __HiDpi_CtrlNewSize($iWidth)
    If $iHeight <> Default Then __HiDpi_CtrlNewSize($iHeight)
EndFunc   ;==>__HiDpi_CtrlBulkNewSize

;https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdpiforwindow
Func _HiDpi_GetDpiForWindow($hWnd = Default)
    Local $iErr, $iExt, $aResult, $gui = 0
    If $hWnd = Default Then $gui = GUICreate("")
    $aResult = DllCall("user32.dll", "uint", "GetDpiForWindow", "hwnd", ($gui ? $gui : $hWnd)) ;requires Win10 v1607+ / no server support
    $iErr = @error ; adaptation from https://www.autoitscript.com/forum/topic/139260-autoit-snippets/?do=findComment&comment=1521526
    $iExt = @extended ; were the original code can be found.
    If $gui Then GUIDelete($gui)
    If $iErr Or Not IsArray($aResult) Then Return SetError($iErr + Int($iErr ? 0 : 100), $iExt, 96)
    Return $aResult[0] ; should be retutning an array[2]
EndFunc   ;==>_HiDpi_GetDpiForWindow

Func _HiDpi_WinGetPos($sTitle, $sText = "") ; Retrieves the position and size of a given window.
    Local $n, $iErr, $iExt, $aArray = WinGetPos($sTitle, $sText)
    $iErr = @error
    $iExt = @extended
    If $iErr Then Return SetError($iErr, $iExt, $aArray)
    For $n = 0 To UBound($aArray) - 1 ; since _HiDpi_GUICreate() will adjust the sizes, "unadjusting" to what would have been is needed.
        $aArray[$n] = _HiDpi_CtrlAdjRatio($aArray[$n], 1) ; 1 = remove DPI radio rather than add
    Next
    Return SetError($iErr, $iExt, $aArray)
EndFunc   ;==>_HiDpi_WinGetPos

Func _HiDpi_ControlGetPos($sTitle, $sText, $sControlID) ; Retrieves the position and size of a control relative to its window.
    Local $n, $iErr, $iExt, $aArray = ControlGetPos($sTitle, $sText, $sControlID)
    $iErr = @error
    $iExt = @extended
    If $iErr Then Return SetError($iErr, $iExt, $aArray)
    For $n = 0 To UBound($aArray) - 1
        $aArray[$n] = _HiDpi_CtrlAdjRatio($aArray[$n], 1)
    Next
    Return SetError($iErr, $iExt, $aArray) ; added on 2023.7.29
EndFunc   ;==>_HiDpi_ControlGetPos