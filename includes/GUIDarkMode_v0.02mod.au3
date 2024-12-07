#include-once ; #include <GUIDarkMode_v0.02mod.au3>

#include <WindowsConstants.au3>
#include <WinAPISysWin.au3>

Global $isDarkMode = _WinAPI_ShouldAppsUseDarkMode()

#Region ; APIThemeConstantsEx.au3
;~ #include "APIThemeConstantsEx.au3"
; _WinAPI_GetIsImmersiveColorUsingHighContrast($IMMERSIVE_HC_CACHE_MODE)
Global Const $IHCM_USE_CACHED_VALUE = 0
Global Const $IIHCM_REFRESH = 1
; _WinAPI_SetPreferredAppMode($PREFERREDAPPMODE)
Global Const $APPMODE_DEFAULT = 0
Global Const $APPMODE_ALLOWDARK = 1
Global Const $APPMODE_FORCEDARK = 2
Global Const $APPMODE_FORCELIGHT = 3
Global Const $APPMODE_MAX = 4
#EndRegion ; APIThemeConstantsEx.au3

#Region ; WinAPIThemeEx.au3
;~ #include "WinAPIThemeEx.au3"


; #CURRENT# =====================================================================================================================
; _WinAPI_ShouldAppsUseDarkMode
; _WinAPI_AllowDarkModeForWindow
; _WinAPI_AllowDarkModeForApp
; _WinAPI_FlushMenuThemes
; _WinAPI_RefreshImmersiveColorPolicyState
; _WinAPI_IsDarkModeAllowedForWindow
; _WinAPI_GetIsImmersiveColorUsingHighContrast
; _WinAPI_OpenNcThemeData
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_ShouldAppsUseDarkMode
; Description ...: Checks if apps should use the dark mode.
; Syntax ........: _WinAPI_ShouldAppsUseDarkMode()
; Parameters ....: None
; Return values .: Success: Returns True if apps should use dark mode.
;                  Failure: Returns False and sets @error:
;                           -1: Operating system version is earlier than Windows 10 (version 1809, build 17763).
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: Requires Windows 10 (version 1809, build 17763) or later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_ShouldAppsUseDarkMode()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnShouldAppsUseDarkMode = 132
	Local $aResult = DllCall('uxtheme.dll', 'bool', $fnShouldAppsUseDarkMode)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_ShouldAppsUseDarkMode

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_AllowDarkModeForWindow
; Description ...: Allows or disallows dark mode for a specific window handle.
; Syntax ........: _WinAPI_AllowDarkModeForWindow($hWnd, $bAllow = True)
; Parameters ....: $hWnd    - Handle to the window.
;                  $bAllow  - [optional] If True, allows dark mode; if False, disallows dark mode. Default is True.
; Return values .: Success: Returns True if the operation succeeded.
;                  Failure: Returns False and sets @error:
;                           -1: Operating system version is earlier than Windows 10 (version 1809, build 17763).
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: Requires Windows 10 (version 1809, build 17763) or later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_AllowDarkModeForWindow($hWnd, $bAllow = True)
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnAllowDarkModeForWindow = 133
	Local $aResult = DllCall('uxtheme.dll', 'bool', $fnAllowDarkModeForWindow, 'hwnd', $hWnd, 'bool', $bAllow)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_AllowDarkModeForWindow

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_AllowDarkModeForApp
; Description ...: Allows or disallows dark mode for the entire application.
; Syntax ........: _WinAPI_AllowDarkModeForApp($bAllow = True)
; Parameters ....: $bAllow  - [optional] If True, allows dark mode for the application; if False, disallows dark mode. Default is True.
; Return values .: Success: Returns True if the operation succeeded.
;                  Failure: Returns False and sets @error:
;                           -1: Operating system version is earlier than Windows 10 (version 1809, build 17763).
;                           -2: Operating system version is later than or equal to Windows 10 (version 1903, build 18362). (Use _WinAPI_SetPreferredAppMode instat!)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: Requires Windows 10 (version 1809, build 17763) and earlier than Windows 10 (version 1903, build 18362).
; Related .......: _WinAPI_SetPreferredAppMode
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_AllowDarkModeForApp($bAllow = True)
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	If @OSBuild >= 18362 Then Return SetError(-2, 0, False)
	Local $fnAllowDarkModeForApp = 135
	Local $aResult = DllCall('uxtheme.dll', 'bool', $fnAllowDarkModeForApp, 'bool', $bAllow)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_AllowDarkModeForApp

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_FlushMenuThemes
; Description ...: Refreshes the system's immersive color policy state, allowing changes to take effect.
; Syntax ........: _WinAPI_FlushMenuThemes()
; Parameters ....: None
; Return values .: Success: True
;                  Failure: False and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 17763)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 17763) and later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_FlushMenuThemes()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnFlushMenuThemes = 136
	Local $aResult = DllCall('uxtheme.dll', 'none', $fnFlushMenuThemes)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>_WinAPI_FlushMenuThemes

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_RefreshImmersiveColorPolicyState
; Description ...: Refreshes the system's immersive color policy state, allowing changes to take effect.
; Syntax ........: _WinAPI_RefreshImmersiveColorPolicyState()
; Parameters ....: None
; Return values .: Success: True
;                  Failure: False and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 17763)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 17763) and later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_RefreshImmersiveColorPolicyState()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnRefreshImmersiveColorPolicyState = 104
	Local $aResult = DllCall('uxtheme.dll', 'none', $fnRefreshImmersiveColorPolicyState)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, False)
	Return True
EndFunc   ;==>_WinAPI_RefreshImmersiveColorPolicyState

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_IsDarkModeAllowedForWindow
; Description ...: Checks if the dark mode is allowed for the specified window.
; Syntax ........: _WinAPI_IsDarkModeAllowedForWindow()
; Parameters ....: None
; Return values .: Success: True if dark mode is allowed for the window, False otherwise.
;                  Failure: False and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 17763)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 17763) and later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_IsDarkModeAllowedForWindow()
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnIsDarkModeAllowedForWindow = 137
	Local $aResult = DllCall('uxtheme.dll', 'bool', $fnIsDarkModeAllowedForWindow)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_IsDarkModeAllowedForWindow

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_GetIsImmersiveColorUsingHighContrast
; Description ...: Retrieves whether immersive color is using high contrast.
; Syntax ........: _WinAPI_GetIsImmersiveColorUsingHighContrast($IMMERSIVE_HC_CACHE_MODE)
; Parameters ....: $IMMERSIVE_HC_CACHE_MODE - The cache mode. Use one of the following values:
;                    $IHCM_USE_CACHED_VALUE (0) - Use the cached value. (Default)
;                    $IHCM_REFRESH (1) - Refresh the value.
; Return values .: Success: True if immersive color is using high contrast.
;                  Failure: False and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 17763)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 17763) and later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_GetIsImmersiveColorUsingHighContrast($IMMERSIVE_HC_CACHE_MODE = $IHCM_USE_CACHED_VALUE)
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnGetIsImmersiveColorUsingHighContrast = 106
	Local $aResult = DllCall('uxtheme.dll', 'bool', $fnGetIsImmersiveColorUsingHighContrast, 'int', $IMMERSIVE_HC_CACHE_MODE)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_GetIsImmersiveColorUsingHighContrast

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_OpenNcThemeData
; Description ...: Opens the theme data for a window.
; Syntax ........: _WinAPI_OpenNcThemeData($hWnd, $pClassList)
; Parameters ....: $hWnd - Handle to the window.
;                  $sClassList - String that contains a semicolon-separated list of classes.
; Return values .: Success: A handle to the theme data.
;                  Failure: 0 and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 17763)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://github.com/ysc3839/win32-darkmode/blob/master/win32-darkmode/DarkMode.h#L69
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_OpenNcThemeData($hWnd, $sClassList)
	If @OSBuild < 17763 Then Return SetError(-1, 0, False)
	Local $fnOpenNcThemeData = 49
	Local $aResult = DllCall('uxtheme.dll', 'hwnd', $fnOpenNcThemeData, 'hwnd', $hWnd, 'wstr', $sClassList)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, 0)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_OpenNcThemeData

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_ShouldSystemUseDarkMode
; Description ...: Checks if system should use the dark mode.
; Syntax ........: _WinAPI_ShouldSystemUseDarkMode()
; Parameters ....: None
; Return values .: Success: Returns True if system should use dark mode.
;                  Failure: Returns False and sets @error:
;                           -1: Operating system version is earlier than Windows 10 (version 1903, build 18362).
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: Requires Windows 10 (version 1903, build 18362) or later.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_ShouldSystemUseDarkMode()
	If @OSBuild < 18362 Then Return SetError(-1, 0, False)
	Local $fnShouldSystemUseDarkMode = 138
	Local $aResult = DllCall('uxtheme.dll', 'bool', $fnShouldSystemUseDarkMode)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_ShouldSystemUseDarkMode

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_SetPreferredAppMode
; Description ...: Sets the preferred application mode for Windows 10 (version 1903, build 18362) and later.
; Syntax ........: _WinAPI_SetPreferredAppMode($PREFERREDAPPMODE)
; Parameters ....: $PREFERREDAPPMODE - The preferred application mode. See enum PreferredAppMode for possible values.
;                    $APPMODE_DEFAULT (0)
;                    $APPMODE_ALLOWDARK (1)
;                    $APPMODE_FORCEDARK (2)
;                    $APPMODE_FORCELIGHT (3)
;                    $APPMODE_MAX (4)
; Return values .: Success: The PreferredAppMode retuned by the DllCall
;                  Failure: '' and sets the @error flag:
;                           -1: Operating system version is earlier than Windows 10 (version 18362)
;                           Other values: DllCall error, check @error @extended for more information.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function is applicable for Windows 10 (version 18362) and later.
; Related .......: _WinAPI_AllowDarkModeForApp
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_SetPreferredAppMode($PREFERREDAPPMODE)
	If @OSBuild < 18362 Then Return SetError(-1, 0, False)
	Local $fnSetPreferredAppMode = 135
	Local $aResult = DllCall('uxtheme.dll', 'int', $fnSetPreferredAppMode, 'int', $PREFERREDAPPMODE)
	If @error Or Not IsArray($aResult) Then Return SetError(@error, @extended, '')
	Return $aResult[0]
EndFunc   ;==>_WinAPI_SetPreferredAppMode


#EndRegion ;b  WinAPIThemeEx.au3

#Region ; GUIStyles.inc.au3
;~ #include "GUIStyles.inc.au3"

Global Const $_g__Style_Gui[32][2] = _
   [[0x80000000, 'WS_POPUP'], _
    [0x40000000, 'WS_CHILD'], _
    [0x20000000, 'WS_MINIMIZE'], _
    [0x10000000, 'WS_VISIBLE'], _
    [0x08000000, 'WS_DISABLED'], _
    [0x04000000, 'WS_CLIPSIBLINGS'], _
    [0x02000000, 'WS_CLIPCHILDREN'], _
    [0x01000000, 'WS_MAXIMIZE'], _
    [0x00CF0000, 'WS_OVERLAPPEDWINDOW'], _ ; (WS_CAPTION | WS_SYSMENU | WS_SIZEBOX | WS_MINIMIZEBOX | WS_MAXIMIZEBOX) aka 'WS_TILEDWINDOW'
    [0x00C00000, 'WS_CAPTION'], _          ; (WS_BORDER | WS_DLGFRAME)
    [0x00800000, 'WS_BORDER'], _
    [0x00400000, 'WS_DLGFRAME'], _
    [0x00200000, 'WS_VSCROLL'], _
    [0x00100000, 'WS_HSCROLL'], _
    [0x00080000, 'WS_SYSMENU'], _
    [0x00040000, 'WS_SIZEBOX'], _
    [0x00020000, '! WS_MINIMIZEBOX ! WS_GROUP'], _   ; ! GUI ! Control
    [0x00010000, '! WS_MAXIMIZEBOX ! WS_TABSTOP'], _ ; ! GUI ! Control
    [0x00002000, 'DS_CONTEXTHELP'], _
    [0x00001000, 'DS_CENTERMOUSE'], _
    [0x00000800, 'DS_CENTER'], _
    [0x00000400, 'DS_CONTROL'], _
    [0x00000200, 'DS_SETFOREGROUND'], _
    [0x00000100, 'DS_NOIDLEMSG'], _
    [0x00000080, 'DS_MODALFRAME'], _
    [0x00000040, 'DS_SETFONT'], _
    [0x00000020, 'DS_LOCALEDIT'], _
    [0x00000010, 'DS_NOFAILCREATE'], _
    [0x00000008, 'DS_FIXEDSYS'], _
    [0x00000004, 'DS_3DLOOK'], _
    [0x00000002, 'DS_SYSMODAL'], _
    [0x00000001, 'DS_ABSALIGN']]
    ;
    ; [0x80880000, 'WS_POPUPWINDOW']
    ; [0x20000000, 'WS_ICONIC']
    ; [0x00040000, 'WS_THICKFRAME']
    ;
    ; [0x00000000, 'WS_OVERLAPPED'] ; also named 'WS_TILED'


Global Const $_g__Style_GuiExtended[21][2] = _
   [[0x08000000, 'WS_EX_NOACTIVATE'], _
    [0x02000000, 'WS_EX_COMPOSITED'], _
    [0x00400000, 'WS_EX_LAYOUTRTL'], _
    [0x00100000, '! WS_EX_NOINHERITLAYOUT ! GUI_WS_EX_PARENTDRAG'], _ ; ! GUI ! Control (label or pic, AutoIt "draggable" feature on 2 controls)
    [0x00080000, 'WS_EX_LAYERED'], _
    [0x00040000, 'WS_EX_APPWINDOW'], _
    [0x00020000, 'WS_EX_STATICEDGE'], _
    [0x00010000, 'WS_EX_CONTROLPARENT'], _ ; AutoIt adds a "draggable" feature to this GUI extended style behavior
    [0x00004000, 'WS_EX_LEFTSCROLLBAR'], _
    [0x00002000, 'WS_EX_RTLREADING'], _
    [0x00001000, 'WS_EX_RIGHT'], _
    [0x00000400, 'WS_EX_CONTEXTHELP'], _
    [0x00000200, 'WS_EX_CLIENTEDGE'], _
    [0x00000100, 'WS_EX_WINDOWEDGE'], _
    [0x00000080, 'WS_EX_TOOLWINDOW'], _
    [0x00000040, 'WS_EX_MDICHILD'], _
    [0x00000020, 'WS_EX_TRANSPARENT'], _
    [0x00000010, 'WS_EX_ACCEPTFILES'], _
    [0x00000008, 'WS_EX_TOPMOST'], _
    [0x00000004, 'WS_EX_NOPARENTNOTIFY'], _
    [0x00000001, 'WS_EX_DLGMODALFRAME']]
    ;
    ; [0x00000300, 'WS_EX_OVERLAPPEDWINDOW']
    ; [0x00000188, 'WS_EX_PALETTEWINDOW']
    ;
    ; [0x00000000, 'WS_EX_LEFT']
    ; [0x00000000, 'WS_EX_LTRREADING']
    ; [0x00000000, 'WS_EX_RIGHTSCROLLBAR']


Global Const $_g__Style_Avi[5][2] = _
   [[0x0010, 'ACS_NONTRANSPARENT'], _
    [0x0008, 'ACS_TIMER'], _
    [0x0004, 'ACS_AUTOPLAY'], _
    [0x0002, 'ACS_TRANSPARENT'], _
    [0x0001, 'ACS_CENTER']]


Global Const $_g__Style_Button[20][2] = _
   [[0x8000, 'BS_FLAT'], _
    [0x4000, 'BS_NOTIFY'], _
    [0x2000, 'BS_MULTILINE'], _
    [0x1000, 'BS_PUSHLIKE'], _
    [0x0C00, 'BS_VCENTER'], _
    [0x0800, 'BS_BOTTOM'], _
    [0x0400, 'BS_TOP'], _
    [0x0300, 'BS_CENTER'], _
    [0x0200, 'BS_RIGHT'], _
    [0x0100, 'BS_LEFT'], _
    [0x0080, 'BS_BITMAP'], _
    [0x0040, 'BS_ICON'], _
    [0x0020, 'BS_RIGHTBUTTON'], _
    [0x0009, 'BS_AUTORADIOBUTTON'] , _
    [0x0007, 'BS_GROUPBOX'], _
    [0x0006, 'BS_AUTO3STATE'], _
    [0x0005, 'BS_3STATE'], _
    [0x0003, 'BS_AUTOCHECKBOX'], _
    [0x0002, 'BS_CHECKBOX'], _
    [0x0001, 'BS_DEFPUSHBUTTON']]


Global Const $_g__Style_Combo[13][2] = _
   [[0x4000, 'CBS_LOWERCASE'], _
    [0x2000, 'CBS_UPPERCASE'], _
    [0x0800, 'CBS_DISABLENOSCROLL'], _
    [0x0400, 'CBS_NOINTEGRALHEIGHT'], _
    [0x0200, 'CBS_HASSTRINGS'], _
    [0x0100, 'CBS_SORT'], _
    [0x0080, 'CBS_OEMCONVERT'], _
    [0x0040, 'CBS_AUTOHSCROLL'], _
    [0x0020, 'CBS_OWNERDRAWVARIABLE'], _
    [0x0010, 'CBS_OWNERDRAWFIXED'], _
    [0x0003, 'CBS_DROPDOWNLIST'], _
    [0x0002, 'CBS_DROPDOWN'], _
    [0x0001, 'CBS_SIMPLE']]


Global Const $_g__Style_Common[12][2] = _ ; "for rebar controls, toolbar controls, and status windows (msdn)"
   [[0x0083, 'CCS_RIGHT'], _
    [0x0082, 'CCS_NOMOVEX'], _
    [0x0081, 'CCS_LEFT'], _
    [0x0080, 'CCS_VERT'], _
    [0x0040, 'CCS_NODIVIDER'], _
    [0x0020, 'CCS_ADJUSTABLE'], _
    [0x0010, 'CCS_NOHILITE'], _
    [0x0008, 'CCS_NOPARENTALIGN'], _
    [0x0004, 'CCS_NORESIZE'], _
    [0x0003, 'CCS_BOTTOM'], _
    [0x0002, 'CCS_NOMOVEY'], _
    [0x0001, 'CCS_TOP']]


Global Const $_g__Style_DateTime[7][2] = _
   [[0x0020, 'DTS_RIGHTALIGN'], _
    [0x0010, 'DTS_APPCANPARSE'], _
    [0x000C, 'DTS_SHORTDATECENTURYFORMAT'], _
    [0x0009, 'DTS_TIMEFORMAT'], _
    [0x0004, 'DTS_LONGDATEFORMAT'], _
    [0x0002, 'DTS_SHOWNONE'], _
    [0x0001, 'DTS_UPDOWN']]
    ;
    ; [0x0000, 'DTS_SHORTDATEFORMAT']


Global Const $_g__Style_Edit[13][2] = _
   [[0x2000, 'ES_NUMBER'], _
    [0x1000, 'ES_WANTRETURN'], _
    [0x0800, 'ES_READONLY'], _
    [0x0400, 'ES_OEMCONVERT'], _
    [0x0100, 'ES_NOHIDESEL'], _
    [0x0080, 'ES_AUTOHSCROLL'], _
    [0x0040, 'ES_AUTOVSCROLL'], _
    [0x0020, 'ES_PASSWORD'], _
    [0x0010, 'ES_LOWERCASE'], _
    [0x0008, 'ES_UPPERCASE'], _
    [0x0004, 'ES_MULTILINE'], _
    [0x0002, 'ES_RIGHT'], _
    [0x0001, 'ES_CENTER']]


Global Const $_g__Style_Header[10][2] = _
   [[0x1000, 'HDS_OVERFLOW'], _
    [0x0800, 'HDS_NOSIZING'], _
    [0x0400, 'HDS_CHECKBOXES'], _
    [0x0200, 'HDS_FLAT'], _
    [0x0100, 'HDS_FILTERBAR'], _
    [0x0080, 'HDS_FULLDRAG'], _
    [0x0040, 'HDS_DRAGDROP'], _
    [0x0008, 'HDS_HIDDEN'], _
    [0x0004, 'HDS_HOTTRACK'], _
    [0x0002, 'HDS_BUTTONS']]
    ;
    ; [0x0000, '$HDS_HORZ']


Global Const $_g__Style_ListBox[16][2] = _
   [[0x8000, 'LBS_COMBOBOX'], _
    [0x4000, 'LBS_NOSEL'], _
    [0x2000, 'LBS_NODATA'], _
    [0x1000, 'LBS_DISABLENOSCROLL'], _
    [0x0800, 'LBS_EXTENDEDSEL'], _
    [0x0400, 'LBS_WANTKEYBOARDINPUT'], _
    [0x0200, 'LBS_MULTICOLUMN'], _
    [0x0100, 'LBS_NOINTEGRALHEIGHT'], _
    [0x0080, 'LBS_USETABSTOPS'], _
    [0x0040, 'LBS_HASSTRINGS'], _
    [0x0020, 'LBS_OWNERDRAWVARIABLE'], _
    [0x0010, 'LBS_OWNERDRAWFIXED'], _
    [0x0008, 'LBS_MULTIPLESEL'], _
    [0x0004, 'LBS_NOREDRAW'], _
    [0x0002, 'LBS_SORT'], _
    [0x0001, 'LBS_NOTIFY']]
    ;
    ; [0xA00003, 'LBS_STANDARD'] ; i.e. (LBS_NOTIFY | LBS_SORT | WS_VSCROLL | WS_BORDER) help file correct, ListBoxConstants.au3 incorrect


Global Const $_g__Style_ListView[17][2] = _
   [[0x8000, 'LVS_NOSORTHEADER'], _
    [0x4000, 'LVS_NOCOLUMNHEADER'], _
    [0x2000, 'LVS_NOSCROLL'], _
    [0x1000, 'LVS_OWNERDATA'], _
    [0x0800, 'LVS_ALIGNLEFT'], _
    [0x0400, 'LVS_OWNERDRAWFIXED'], _
    [0x0200, 'LVS_EDITLABELS'], _
    [0x0100, 'LVS_AUTOARRANGE'], _
    [0x0080, 'LVS_NOLABELWRAP'], _
    [0x0040, 'LVS_SHAREIMAGELISTS'], _
    [0x0020, 'LVS_SORTDESCENDING'], _
    [0x0010, 'LVS_SORTASCENDING'], _
    [0x0008, 'LVS_SHOWSELALWAYS'], _
    [0x0004, 'LVS_SINGLESEL'], _
    [0x0003, 'LVS_LIST'], _
    [0x0002, 'LVS_SMALLICON'], _
    [0x0001, 'LVS_REPORT']]
    ;
    ; [0x0000, 'LVS_ICON']
    ; [0x0000, 'LVS_ALIGNTOP']


Global Const $_g__Style_ListViewExtended[20][2] = _
   [[0x00100000, 'LVS_EX_SIMPLESELECT'], _
    [0x00080000, 'LVS_EX_SNAPTOGRID'], _
    [0x00020000, 'LVS_EX_HIDELABELS'], _
    [0x00010000, 'LVS_EX_DOUBLEBUFFER'], _
    [0x00008000, 'LVS_EX_BORDERSELECT'], _
    [0x00004000, 'LVS_EX_LABELTIP'], _
    [0x00002000, 'LVS_EX_MULTIWORKAREAS'], _
    [0x00001000, 'LVS_EX_UNDERLINECOLD'], _
    [0x00000800, 'LVS_EX_UNDERLINEHOT'], _
    [0x00000400, 'LVS_EX_INFOTIP'], _
    [0x00000200, 'LVS_EX_REGIONAL'], _
    [0x00000100, 'LVS_EX_FLATSB'], _
    [0x00000080, 'LVS_EX_TWOCLICKACTIVATE'], _
    [0x00000040, 'LVS_EX_ONECLICKACTIVATE'], _
    [0x00000020, 'LVS_EX_FULLROWSELECT'], _
    [0x00000010, 'LVS_EX_HEADERDRAGDROP'], _
    [0x00000008, 'LVS_EX_TRACKSELECT'], _
    [0x00000004, 'LVS_EX_CHECKBOXES'], _
    [0x00000002, 'LVS_EX_SUBITEMIMAGES'], _
    [0x00000001, 'LVS_EX_GRIDLINES']]


Global Const $_g__Style_MonthCal[8][2] = _
   [[0x0100, 'MCS_NOSELCHANGEONNAV'], _
    [0x0080, 'MCS_SHORTDAYSOFWEEK'], _
    [0x0040, 'MCS_NOTRAILINGDATES'], _
    [0x0010, 'MCS_NOTODAY'], _
    [0x0008, 'MCS_NOTODAYCIRCLE'], _
    [0x0004, 'MCS_WEEKNUMBERS'], _
    [0x0002, 'MCS_MULTISELECT'], _
    [0x0001, 'MCS_DAYSTATE']]


Global Const $_g__Style_Pager[3][2] = _
   [[0x0004, 'PGS_DRAGNDROP'], _
    [0x0002, 'PGS_AUTOSCROLL'], _
    [0x0001, 'PGS_HORZ']]
    ;
    ; [0x0000, 'PGS_VERT']


Global Const $_g__Style_Progress[4][2] = _
   [[0x0010, 'PBS_SMOOTHREVERSE'], _
    [0x0008, 'PBS_MARQUEE'], _
    [0x0004, 'PBS_VERTICAL'], _
    [0x0001, 'PBS_SMOOTH']]


Global Const $_g__Style_Rebar[8][2] = _
   [[0x8000, 'RBS_DBLCLKTOGGLE'], _
    [0x4000, 'RBS_VERTICALGRIPPER'], _
    [0x2000, 'RBS_AUTOSIZE'], _
    [0x1000, 'RBS_REGISTERDROP'], _
    [0x0800, 'RBS_FIXEDORDER'], _
    [0x0400, 'RBS_BANDBORDERS'], _
    [0x0200, 'RBS_VARHEIGHT'], _
    [0x0100, 'RBS_TOOLTIPS']]


Global Const $_g__Style_RichEdit[8][2] = _      ; will also use plenty (not all) of Edit styles
   [[0x01000000, 'ES_SELECTIONBAR'], _
    [0x00400000, 'ES_VERTICAL'], _          ; Asian-language support only (msdn)
    [0x00080000, 'ES_NOIME'], _             ; ditto
    [0x00040000, 'ES_SELFIME'], _           ; ditto
    [0x00008000, 'ES_SAVESEL'], _
    [0x00004000, 'ES_SUNKEN'], _
    [0x00002000, 'ES_DISABLENOSCROLL'], _   ; same value as 'ES_NUMBER' => issue ?
    [0x00000008, 'ES_NOOLEDRAGDROP']]       ; same value as 'ES_UPPERCASE' but RichRdit controls do not support 'ES_UPPERCASE' style (msdn)


Global Const $_g__Style_Scrollbar[5][2] = _
   [[0x0010, 'SBS_SIZEGRIP'], _
    [0x0008, 'SBS_SIZEBOX'], _
    [0x0004, 'SBS_RIGHTALIGN or SBS_BOTTOMALIGN'], _ ; i.e. use SBS_RIGHTALIGN with SBS_VERT, use SBS_BOTTOMALIGN with SBS_HORZ (msdn)
    [0x0002, 'SBS_LEFTALIGN or SBS_TOPALIGN'], _     ; i.e. use SBS_LEFTALIGN  with SBS_VERT, use SBS_TOPALIGN    with SBS_HORZ (msdn)
    [0x0001, 'SBS_VERT']]
    ;
    ; [0x0000, 'SBS_HORZ']


Global Const $_g__Style_Slider[13][2] = _ ; i.e. trackbar
   [[0x1000, 'TBS_TRANSPARENTBKGND'], _
    [0x0800, 'TBS_NOTIFYBEFOREMOVE'], _
    [0x0400, 'TBS_DOWNISLEFT'], _
    [0x0200, 'TBS_REVERSED'], _
    [0x0100, 'TBS_TOOLTIPS'], _
    [0x0080, 'TBS_NOTHUMB'], _
    [0x0040, 'TBS_FIXEDLENGTH'], _
    [0x0020, 'TBS_ENABLESELRANGE'], _
    [0x0010, 'TBS_NOTICKS'], _
    [0x0008, 'TBS_BOTH'], _
    [0x0004, 'TBS_LEFT or TBS_TOP'], _ ; i.e. TBS_LEFT tick marks when vertical slider, or TBS_TOP tick marks when horizontal slider
    [0x0002, 'TBS_VERT'], _
    [0x0001, 'TBS_AUTOTICKS']]
    ;
    ; [0x0000, 'TBS_RIGHT']
    ; [0x0000, 'TBS_BOTTOM']
    ; [0x0000, 'TBS_HORZ']


Global Const $_g__Style_Static[18][2] = _
   [[0x1000, 'SS_SUNKEN'], _
    [0x0400, 'SS_RIGHTJUST'], _
    [0x0200, 'SS_CENTERIMAGE'], _
    [0x0100, 'SS_NOTIFY'], _
    [0x0080, 'SS_NOPREFIX'], _
    [0x0012, 'SS_ETCHEDFRAME'], _
    [0x0011, 'SS_ETCHEDVERT'], _
    [0x0010, 'SS_ETCHEDHORZ'], _
    [0x000C, 'SS_LEFTNOWORDWRAP'], _
    [0x000B, 'SS_SIMPLE'], _
    [0x0009, 'SS_WHITEFRAME'], _
    [0x0008, 'SS_GRAYFRAME'], _
    [0x0007, 'SS_BLACKFRAME'], _
    [0x0006, 'SS_WHITERECT'], _
    [0x0005, 'SS_GRAYRECT'], _
    [0x0004, 'SS_BLACKRECT'], _
    [0x0002, 'SS_RIGHT'], _
    [0x0001, 'SS_CENTER']]
    ;
    ; [0x0000, 'SS_LEFT']


Global Const $_g__Style_StatusBar[2][2] = _
   [[0x0800, 'SBARS_TOOLTIPS'], _
    [0x0100, 'SBARS_SIZEGRIP']]
    ;
    ; [0x0800, 'SBT_TOOLTIPS']


Global Const $_g__Style_Tab[17][2] = _
   [[0x8000, 'TCS_FOCUSNEVER'], _
    [0x4000, 'TCS_TOOLTIPS'], _
    [0x2000, 'TCS_OWNERDRAWFIXED'], _
    [0x1000, 'TCS_FOCUSONBUTTONDOWN'], _
    [0x0800, 'TCS_RAGGEDRIGHT'], _
    [0x0400, 'TCS_FIXEDWIDTH'], _
    [0x0200, 'TCS_MULTILINE'], _
    [0x0100, 'TCS_BUTTONS'], _
    [0x0080, 'TCS_VERTICAL'], _
    [0x0040, 'TCS_HOTTRACK'], _
    [0x0020, 'TCS_FORCELABELLEFT'], _
    [0x0010, 'TCS_FORCEICONLEFT'], _
    [0x0008, 'TCS_FLATBUTTONS'], _
    [0x0004, 'TCS_MULTISELECT'], _
    [0x0002, 'TCS_RIGHT'], _
    [0x0002, 'TCS_BOTTOM'], _
    [0x0001, 'TCS_SCROLLOPPOSITE']]
    ;
    ; [0x0000, 'TCS_TABS']
    ; [0x0000, 'TCS_SINGLELINE']
    ; [0x0000, 'TCS_RIGHTJUSTIFY']


Global Const $_g__Style_Toolbar[8][2] = _
   [[0x8000, 'TBSTYLE_TRANSPARENT'], _
    [0x4000, 'TBSTYLE_REGISTERDROP'], _
    [0x2000, 'TBSTYLE_CUSTOMERASE'], _
    [0x1000, 'TBSTYLE_LIST'], _
    [0x0800, 'TBSTYLE_FLAT'], _
    [0x0400, 'TBSTYLE_ALTDRAG'], _
    [0x0200, 'TBSTYLE_WRAPABLE'], _
    [0x0100, 'TBSTYLE_TOOLTIPS']]


Global Const $_g__Style_TreeView[16][2] = _
   [[0x8000, 'TVS_NOHSCROLL'], _
    [0x4000, 'TVS_NONEVENHEIGHT'], _
    [0x2000, 'TVS_NOSCROLL'], _
    [0x1000, 'TVS_FULLROWSELECT'], _
    [0x0800, 'TVS_INFOTIP'], _
    [0x0400, 'TVS_SINGLEEXPAND'], _
    [0x0200, 'TVS_TRACKSELECT'], _
    [0x0100, 'TVS_CHECKBOXES'], _
    [0x0080, 'TVS_NOTOOLTIPS'], _
    [0x0040, 'TVS_RTLREADING'], _
    [0x0020, 'TVS_SHOWSELALWAYS'], _
    [0x0010, 'TVS_DISABLEDRAGDROP'], _
    [0x0008, 'TVS_EDITLABELS'], _
    [0x0004, 'TVS_LINESATROOT'], _
    [0x0002, 'TVS_HASLINES'], _
    [0x0001, 'TVS_HASBUTTONS']]


Global Const $_g__Style_UpDown[9][2] = _
   [[0x0100, 'UDS_HOTTRACK'], _
    [0x0080, 'UDS_NOTHOUSANDS'], _
    [0x0040, 'UDS_HORZ'], _
    [0x0020, 'UDS_ARROWKEYS'], _
    [0x0010, 'UDS_AUTOBUDDY'], _
    [0x0008, 'UDS_ALIGNLEFT'], _
    [0x0004, 'UDS_ALIGNRIGHT'], _
    [0x0002, 'UDS_SETBUDDYINT'], _
    [0x0001, 'UDS_WRAP']]


#EndRegion ; GUIStyles.inc.au3

#Region ; GUIStyles.au3
;~ #include "GUIStyles.au3"
Func hWnd2Styles($hWnd)
	Return _GetCtrlStyleString(_WinAPI_GetWindowLong($hWnd, $GWL_STYLE), _WinAPI_GetWindowLong($hWnd, $GWL_EXSTYLE), _WinAPI_GetClassName($hWnd))
EndFunc

Func _GetStyleString($iStyle, $fExStyle)
	ConsoleWrite('+ Func _GetStyleString(' & $iStyle & ', ' & $fExStyle & ')' & @CRLF)
	Local $Text = '', $Data = $fExStyle ? $_g__Style_GuiExtended : $_g__Style_Gui

	For $i = 0 To UBound($Data) - 1
		If BitAND($iStyle, $Data[$i][0]) = $Data[$i][0] Then
			$iStyle = BitAND($iStyle, BitNOT($Data[$i][0]))
			If StringLeft($Data[$i][1], 1) <> "!" Then
				$Text &= $Data[$i][1] & ', '
			Else
				; ex. '! WS_MINIMIZEBOX ! WS_GROUP'  =>  'WS_MINIMIZEBOX, '
				$Text &= StringMid($Data[$i][1], 3, StringInStr($Data[$i][1], "!", 2, 2) - 4) & ', '
			EndIf
		EndIf
	Next

	If $iStyle Then $Text = '0x' & Hex($iStyle, 8) & ', ' & $Text

	Return StringRegExpReplace($Text, ',\s\z', '')
EndFunc   ;==>_GetStyleString

Func _GetCtrlStyleString($iStyle, $fExStyle, $sClass, $iLVExStyle = 0)

	If $sClass = "AutoIt v3 GUI" Or $sClass = "#32770" Or $sClass = "MDIClient" Then ; control = child GUI, dialog box (msgbox) etc...
		Return _GetStyleString($iStyle, 0)
	EndIf

	If StringLeft($sClass, 8) = "RichEdit" Then $sClass = "RichEdit" ; RichEdit, RichEdit20A, RichEdit20W, RichEdit50A, RichEdit50W

	Local $Text = ''

	_GetCtrlStyleString2($iStyle, $Text, $sClass, $iLVExStyle) ; 4th param. in case $sClass = "Ex_SysListView32" (special treatment)

	If $sClass = "ReBarWindow32" Or $sClass = "ToolbarWindow32" Or $sClass = "msctls_statusbar32" Then
		$sClass = "Common" ; "for rebar controls, toolbar controls, and status windows" (msdn)
		_GetCtrlStyleString2($iStyle, $Text, $sClass)
	ElseIf $sClass = "RichEdit" Then
		$sClass = "Edit" ; "Richedit controls also support many edit control styles (not all)" (msdn)
		_GetCtrlStyleString2($iStyle, $Text, $sClass)
	EndIf

	Local $Data = $fExStyle ? $_g__Style_GuiExtended : $_g__Style_Gui

	For $i = 0 To UBound($Data) - 1
		If BitAND($iStyle, $Data[$i][0]) = $Data[$i][0] Then
			If (Not BitAND($Data[$i][0], 0xFFFF)) Or ($fExStyle) Then
				$iStyle = BitAND($iStyle, BitNOT($Data[$i][0]))
				If StringLeft($Data[$i][1], 1) <> "!" Then
					$Text &= $Data[$i][1] & ', '
				Else
					; ex. '! WS_MINIMIZEBOX ! WS_GROUP'  =>  'WS_GROUP, '
					$Text &= StringMid($Data[$i][1], StringInStr($Data[$i][1], "!", 2, 2) + 2) & ', '
				EndIf
			EndIf
		EndIf
	Next

	If $iStyle Then $Text = '0x' & Hex($iStyle, 8) & ', ' & $Text

	Return StringRegExpReplace($Text, ',\s\z', '')
EndFunc   ;==>_GetCtrlStyleString

;=====================================================================
Func _GetCtrlStyleString2(ByRef $iStyle, ByRef $Text, $sClass, $iLVExStyle = 0)

	Local $Data

	Switch $sClass  ; $Input[16]
		Case "Button"
			$Data = $_g__Style_Button
		Case "ComboBox", "ComboBoxEx32"
			$Data = $_g__Style_Combo
		Case "Common"
			$Data = $_g__Style_Common ; "for rebar controls, toolbar controls, and status windows (msdn)"
		Case "Edit"
			$Data = $_g__Style_Edit
		Case "ListBox"
			$Data = $_g__Style_ListBox
		Case "msctls_progress32"
			$Data = $_g__Style_Progress
		Case "msctls_statusbar32"
			$Data = $_g__Style_StatusBar
		Case "msctls_trackbar32"
			$Data = $_g__Style_Slider
		Case "msctls_updown32"
			$Data = $_g__Style_UpDown
		Case "ReBarWindow32"
			$Data = $_g__Style_Rebar
		Case "RichEdit"
			$Data = $_g__Style_RichEdit
		Case "Scrollbar"
			$Data = $_g__Style_Scrollbar
		Case "Static"
			$Data = $_g__Style_Static
		Case "SysAnimate32"
			$Data = $_g__Style_Avi
		Case "SysDateTimePick32"
			$Data = $_g__Style_DateTime
		Case "SysHeader32"
			$Data = $_g__Style_Header
		Case "SysListView32"
			$Data = $_g__Style_ListView
		Case "Ex_SysListView32" ; special treatment below
			$Data = $_g__Style_ListViewExtended
		Case "SysMonthCal32"
			$Data = $_g__Style_MonthCal
		Case "SysPager"
			$Data = $_g__Style_Pager
		Case "SysTabControl32", "SciTeTabCtrl"
			$Data = $_g__Style_Tab
		Case "SysTreeView32"
			$Data = $_g__Style_TreeView
		Case "ToolbarWindow32"
			$Data = $_g__Style_Toolbar
		Case Else
			Return
	EndSwitch

	If $sClass <> "Ex_SysListView32" Then
		For $i = 0 To UBound($Data) - 1
			If BitAND($iStyle, $Data[$i][0]) = $Data[$i][0] Then
				$iStyle = BitAND($iStyle, BitNOT($Data[$i][0]))
				$Text = $Data[$i][1] & ', ' & $Text
			EndIf
		Next
	Else
		For $i = 0 To UBound($Data) - 1
			If BitAND($iLVExStyle, $Data[$i][0]) = $Data[$i][0] Then
				$iLVExStyle = BitAND($iLVExStyle, BitNOT($Data[$i][0]))
				$Text = $Data[$i][1] & ', ' & $Text
				If BitAND($iStyle, $Data[$i][0]) = $Data[$i][0] Then
					$iStyle = BitAND($iStyle, BitNOT($Data[$i][0]))
				EndIf
			EndIf
		Next
		If $iLVExStyle Then $Text = 'LVex: 0x' & Hex($iLVExStyle, 8) & ', ' & $Text
		; next test bc LVS_EX_FULLROWSELECT (default AutoIt LV ext style) and WS_EX_TRANSPARENT got both same value 0x20 (hard to solve in some cases)
		If BitAND($iStyle, $WS_EX_TRANSPARENT) = $WS_EX_TRANSPARENT Then ; note that $WS_EX_TRANSPARENT has nothing to do with listview
			$iStyle = BitAND($iStyle, BitNOT($WS_EX_TRANSPARENT))
		EndIf
	EndIf
EndFunc   ;==>_GetCtrlStyleString2



#EndRegion ; GUIStyles.au3


;~ #include "_FindDelayLoadThunkInModule.au3"

#include <WinAPIGdi.au3>
#include <WinAPIDlg.au3> ; _WinAPI_GetDlgCtrlID


#include <SendMessage.au3>
;~ #include <WindowsConstants.au3>
#include <WinAPITheme.au3>
;~ #include <WinAPISysWin.au3>

#include <Array.au3> ; Just For _ArrayDisplay()
#include <GuiListView.au3> ; Debug ListView Header


; #INDEX# =======================================================================================================================
; Title .........: GUIDarkmode UDF Library for AutoIt3
; AutoIt Version : 3.3.16.1
; Description ...: Additional variables, constants and functions for the WinAPITheme.au3
; Author(s) .....: NoNameCode
; ===============================================================================================================================

#Region Global Variables and Constants

; #VARIABLES# ===================================================================================================================

; Darkmode default colors for GUI / -Ctrls
Global $GUIDARKMODE_COLOR_GUIBK = 0x202020 ; 0x383838 ; 0x202020 ;Orig argumentum: 0x1B1B1B
Global $GUIDARKMODE_COLOR_GUICTRL = 0xe0e0e0 ;Orig argumentum: 0xFBFBFB
Global $GUIDARKMODE_COLOR_GUICTRLBK = 0x202020 ; 0x383838 ; 0x202020 ; 0x3f3f3f ;Orig argumentum: 0x2B2B2B

; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================

Global Const $DWMWA_USE_IMMERSIVE_DARK_MODE = (@OSBuild <= 18985) ? 19 : 20            ; before this build set to 19, otherwise set to 20, no thanks Windaube to document anything ??

; ===============================================================================================================================
#EndRegion Global Variables and Constants

#Region Functions list
; #CURRENT# =====================================================================================================================
; _GUISetDarkTheme
; _GUICtrlSetDarkTheme
; _GUICtrlAllSetDarkTheme
; ===============================================================================================================================
#EndRegion Functions list



#Region Public Functions

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUISetDarkTheme
; Description ...: Sets the theme for a specified window to either dark or light mode on Windows 10.
; Syntax ........: _GUISetDarkTheme($hwnd, $dark_theme = True)
; Parameters ....: $hwnd          - The handle to the window.
;                  $dark_theme    - If True, sets the dark theme; if False, sets the light theme.
;                                   (Default is True for dark theme.)
; Return values .: None
; Author ........: DK12000, NoNameCode
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/211196-gui-title-bar-dark-theme-an-elegant-solution-using-dwmapi/
; Example .......: No
; ===============================================================================================================================
Func _GUISetDarkTheme($hWnd, $bEnableDarkTheme = True)
	Local $iPreferredAppMode = ($bEnableDarkTheme == True) ? $APPMODE_FORCEDARK : $APPMODE_FORCELIGHT
	Local $iGUI_BkColor = ($bEnableDarkTheme == True) ? $GUIDARKMODE_COLOR_GUIBK : _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_3DFACE))
	_WinAPI_SetPreferredAppMode($iPreferredAppMode)
	_WinAPI_RefreshImmersiveColorPolicyState()
	_WinAPI_FlushMenuThemes()
	GUISetBkColor($iGUI_BkColor, $hWnd)
	_GUICtrlSetDarkTheme($hWnd, $bEnableDarkTheme)            ;To Color the GUI's own Scrollbar
;~ 	DllCall('dwmapi.dll', 'long', 'DwmSetWindowAttribute', 'hwnd', $hWnd, 'dword', $DWMWA_USE_IMMERSIVE_DARK_MODE, 'dword*', Int($bEnableDarkTheme), 'dword', 4)
	_WinAPI_DwmSetWindowAttribute_unr($hWnd, $DWMWA_USE_IMMERSIVE_DARK_MODE, $bEnableDarkTheme)
EndFunc   ;==>_GUISetDarkTheme

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUICtrlAllSetDarkTheme
; Description ...: Sets the dark theme to all existing sub Controls from a GUI
; Syntax ........: _GUICtrlAllSetDarkTheme($hGUI[, $bEnableDarkTheme = True])
; Parameters ....: $hGUI                - GUI handle
;                  $bEnableDarkTheme    - [optional] a boolean value. Default is True.
; Return values .: None
; Author ........: NoName
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GUICtrlAllSetDarkTheme($hGUI, $bEnableDarkTheme = True)
	Local $aCtrls = _WinAPI_EnumChildWindows($hGUI, False)
	If @error Then
		Dim $aCtrls[1][2] = [[0]]
		Return SetError(1, UBound($aCtrls) - 1, $aCtrls)
	EndIf
	For $i = 1 To $aCtrls[0][0]
		_GUICtrlSetDarkTheme($aCtrls[$i][0], $bEnableDarkTheme)
	Next
	Return $aCtrls
EndFunc   ;==>_GUICtrlAllSetDarkTheme

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUICtrlSetDarkTheme
; Description ...: Sets the dark theme for a specified control.
; Syntax ........: _GUICtrlSetDarkTheme($vCtrl, $bEnableDarkTheme = True)
; Parameters ....: $vCtrl            - The control handle or identifier.
;                  $bEnableDarkTheme - If True, enables the dark theme; if False, disables it.
;                                      (Default is True for enabling dark theme.)
; Return values .: Success: True
;                  Failure: False and sets the @error flag:
;                           1: Invalid control handle or identifier.
;                           2: Error while allowing dark mode for the window.
;                           3: Error while setting the window theme.
;                           4: Error while sending the WM_THEMECHANGED message.
; Author ........: NoNameCode
; Modified ......:
; Remarks .......: This function requires the _WinAPI_SetWindowTheme and _WinAPI_AllowDarkModeForWindow functions.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: Yes
; ===============================================================================================================================
Global 	$iGUI_Ctrl_Color = -2, $iGUI_Ctrl_BkColor = -2
Func _GUICtrlSetDarkTheme($vCtrl, $bEnableDarkTheme = True)
	Local $sThemeName = Null, $sThemeList = Null
	$iGUI_Ctrl_Color = ($bEnableDarkTheme == True) ? $GUIDARKMODE_COLOR_GUICTRL : _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT))
	$iGUI_Ctrl_BkColor = ($bEnableDarkTheme == True) ? $GUIDARKMODE_COLOR_GUICTRLBK : _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))
;~ 	If Not IsHWnd($vCtrl) Then $vCtrl = GUICtrlGetHandle($vCtrl)
;~ 	If Not IsHWnd($vCtrl) Then Return SetError(1, 0, False)
	_WinAPI_AllowDarkModeForWindow($vCtrl, $bEnableDarkTheme)
	If @error <> 0 Then
		ConsoleWrite('! _GUICtrlSetDarkTheme: _WinAPI_AllowDarkModeForWindow: FAILED @ ' & $vCtrl & @CRLF)
		Return SetError(2, @error, False)
	EndIf

	;=========
;~ 	_WinAPI_SetWindowTheme_unr($vCtrl, 0, 0) ; argu Testing
;~ 	ConsoleWrite(@CRLF & _WinAPI_GetClassName($vCtrl))
	Local $sStyles = hWnd2Styles($vCtrl)

	Switch _WinAPI_GetClassName($vCtrl)
		Case 'Button'
			If StringInStr($sStyles, "BS_AUTORADIOBUTTON") Or StringInStr($sStyles, "BS_GROUPBOX") Or StringInStr($sStyles, "BS_AUTOCHECKBOX")   Then
				$sThemeName = ""
				ConsoleWrite('>' & $sStyles & '<' & @CRLF)
				If Not StringInStr($sStyles, "BS_PUSHLIKE") Then _WinAPI_SetWindowTheme($vCtrl, "", "")
;~ 				DllCall("UxTheme.dll", "int", "SetWindowTheme", "hwnd", $vCtrl, "wstr", 0, "wstr", 0)
				GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
				GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)
			Else
                If $isDarkMode = True Then
                $sThemeName = 'DarkMode_Explorer'
                Else
                $sThemeName = 'Explorer'
                EndIf
			EndIf


		Case 'AutoIt v3 GUI'
			$sThemeName = 'DarkMode_Explorer'

		Case 'ComboBox' ;,'Edit'
			$sThemeName = 'DarkMode_CFD'
			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)
		Case 'SysHeader32'
			$sThemeName = 'ItemsView'
			$sThemeList = 'Header'
		Case 'ListBox', 'SysTreeView32', 'SysListView32', 'Edit', 'msctls_trackbar32'
			If $isDarkMode = True Then
            $sThemeName = 'DarkMode_Explorer'
            Else
            $sThemeName = 'Explorer'
            EndIf
;~ 			$sThemeList = 0 ; failed
			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)
		Case 'Static'
			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color )  ; $iGUI_Ctrl_Color
		Case 'SysDateTimePick32'
			;Right now; no solution
;~ 			_WinAPI_SetWindowTheme($vCtrl, "", "")
		Case 'SysTabControl32'
			_WinAPI_SetWindowTheme($vCtrl, "", "")

			$sThemeName = 'DarkMode'
;~ 			$sThemeName = 'DarkMode_Explorer'
;~ 			$sThemeList = 'ExplorerStatusBar'				;=> Favorite
			$sThemeList = 'FileExplorerBannerContainer'        ;=> Favorite
;~ 			$sThemeList = 'ItemsView'
;~ 			$sThemeList = 'Menu'
;~ 			$sThemeList = 'Toolbar'
			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)


			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)
			;Right now; no solution
		Case 'Scrollbar'
			$sThemeName = 'DarkMode_Explorer'
;~ 			$sThemeName = 'DarkMode_CFD'
;~ 			$sThemeName = 'DarkMode'
;~ 			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)
;~ 			GUICtrlSetBkColor($vCtrl, $iGUI_Ctrl_BkColor)
;~ 			GUISetBkColor($iGUI_Ctrl_BkColor, $vCtrl)

		Case Else
			If $sThemeName = Null Then $sThemeName = 'Explorer'
	EndSwitch
	ConsoleWrite(@CRLF & 'Class:' & _WinAPI_GetClassName($vCtrl) & ' Theme:' & $sThemeName & '::' & $sThemeList & @TAB & ControlGetText($vCtrl, "", 0) & @TAB & hWnd2Styles($vCtrl))
	;=========
	If $sThemeName = "" Then Return True
	_WinAPI_SetWindowTheme_unr($vCtrl, $sThemeName, $sThemeList)
	If @error <> 0 Then Return SetError(3, @error, False)
	_SendMessage($vCtrl, $WM_THEMECHANGED, 0, 0)
	If @error <> 0 Then Return SetError(4, @error, False)
	Return True
EndFunc   ;==>_GUICtrlSetDarkTheme

;//Testting
Func _GUICtrlSetDarkThemeEx($vCtrl, $sThemeName = Null, $sThemeList = Null, $bEnableDarkTheme = True)
	If Not IsHWnd($vCtrl) Then $vCtrl = GUICtrlGetHandle($vCtrl)
	If Not IsHWnd($vCtrl) Then Return SetError(1, 0, False)
	_WinAPI_AllowDarkModeForWindow($vCtrl, $bEnableDarkTheme)
	If @error <> 0 Then Return SetError(2, @error, False)
	_WinAPI_SetWindowTheme_unr($vCtrl, $sThemeName, $sThemeList)
	If @error <> 0 Then Return SetError(3, @error, False)
	_SendMessage($vCtrl, $WM_THEMECHANGED, 0, 0)
	If @error <> 0 Then Return SetError(4, @error, False)
	Return True
EndFunc   ;==>_GUICtrlSetDarkThemeEx

#EndRegion Public Functions

#Region Internal Functions


; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_SetWindowTheme_unr
; Description ...:  Dose the same as _WinAPI_SetWindowTheme; But has no Restrictions
; Syntax ........: _WinAPI_SetWindowTheme_unr($hWnd[, $sName = Null[, $sList = Null]])
; Parameters ....: $hWnd                - a handle value.
;                  $sName               - [optional] a string value. Default is Null.
;                  $sList               - [optional] a string value. Default is Null.
; Return values .: Success: 1 Failure: @error, @extended & False
; Author ........: argumentum
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/211475-winapithemeex-darkmode-for-autoits-win32guis/?do=findComment&comment=1530103
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_SetWindowTheme_unr($hWnd, $sName = Null, $sList = Null) ; #include <WinAPITheme.au3> ; unthoughtful unrestricting mod.
	Local $sResult = DllCall('UxTheme.dll', 'long', 'SetWindowTheme', 'hwnd', $hWnd, 'wstr', $sName, 'wstr', $sList)
	If @error Then Return SetError(@error, @extended, 0)
	If $sResult[0] Then Return SetError(10, $sResult[0], 0)
	Return 1
EndFunc   ;==>_WinAPI_SetWindowTheme_unr

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_DwmSetWindowAttribute_unr
; Description ...: Dose the same as _WinAPI_DwmSetWindowAttribute; But has no Restrictions
; Syntax ........: _WinAPI_DwmSetWindowAttribute_unr($hWnd, $iAttribute, $iData)
; Parameters ....: $hWnd                - a handle value.
;                  $iAttribute          - an integer value.
;                  $iData               - an integer value.
; Return values .: Success: 1 Failure: @error, @extended & False
; Author ........: argumentum
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/211475-winapithemeex-darkmode-for-autoits-win32guis/?do=findComment&comment=1530103
; Example .......: No
; ===============================================================================================================================
Func _WinAPI_DwmSetWindowAttribute_unr($hWnd, $iAttribute, $iData) ; #include <WinAPIGdi.au3> ; unthoughtful unrestricting mod.
	Local $aCall = DllCall('dwmapi.dll', 'long', 'DwmSetWindowAttribute', 'hwnd', $hWnd, 'dword', $iAttribute, _
			'dword*', $iData, 'dword', 4)
	If @error Then Return SetError(@error, @extended, 0)
	If $aCall[0] Then Return SetError(10, $aCall[0], 0)
	Return 1
EndFunc   ;==>_WinAPI_DwmSetWindowAttribute_unr

#EndRegion Internal Functions

#Region Experimental Functions

;~ Func FixDarkScrollBar()
;~     Local $hComctl = _WinAPI_GetModuleHandle("comctl32.dll")
;~     If $hComctl Then
;~         Local $addr = _FindDelayLoadThunkInModule($hComctl, $UxThemeDLL, $OpenNcThemeDataOrdinal)
;~         If $addr Then
;~             Local $oldProtect, $MyOpenThemeData

;~             If _WinAPI_VirtualProtect($addr, DllStructGetSize(DllStructCreate("ptr")), $PAGE_READWRITE, $oldProtect) Then
;~                 $MyOpenThemeData = DLLCallbackRegister("_Modifyed_OpenNcThemeData", "lresult", "hwnd;wstr")
;~                 DllStructSetData(DllStructCreate("ptr", $addr), 1, $MyOpenThemeData)
;~                 _WinAPI_VirtualProtect($addr, DllStructGetSize(DllStructCreate("ptr")), $oldProtect, 0)
;~             EndIf
;~         EndIf
;~     EndIf
;~ EndFunc


;~ Func _Modifyed_OpenNcThemeData($hWnd, $classList)
;~     If StringCompare($classList, "ScrollBar") = 0 Then
;~         $hWnd = 0
;~         $classList = "Explorer::ScrollBar"
;~     EndIf
;~     Return _WinAPI_OpenNcThemeData($hWnd, $classList)
;~ EndFunc

#EndRegion Experimental Functions


#Region Enable GUI DARKMODE

Func GuiDarkmodeApply($hGUI)
	_GUISetDarkTheme($hGUI)
	_GUICtrlAllSetDarkTheme($hGUI)
	ConsoleWrite(@CRLF & '+ _GUICtrlAllSetDarkTheme: ' & @error & ',' & @extended & @CRLF)
EndFunc

#EndRegion Enable GUI DARKMODE

#Region Enable GUI LIGHTMODE

Func GuiLightmodeApply($hGUI)
    $bEnableDarkTheme = False
    _GUISetDarkTheme($hGUI, $bEnableDarkTheme)
    _GUICtrlAllSetDarkTheme($hGUI, $bEnableDarkTheme)
	ConsoleWrite(@CRLF & '+ _GUICtrlAllSetDarkTheme: ' & @error & ',' & @extended & @CRLF)
EndFunc

#EndRegion Enable GUI LIGHTMODE

