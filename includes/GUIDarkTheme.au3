#include-once
#include <GuiMonthCal.au3>
#include <GuiDateTimePicker.au3>
#include <WindowsConstants.au3>
#include <GuiStatusBar.au3>
#include <GuiListView.au3>
#include <TabConstants.au3>
#include <GuiImageList.au3>
#include <GuiTreeView.au3>
#include <GuiMenu.au3>
#include <GDIPlus.au3>
#include "GUIDarkAPI.au3"
#include "GUIDarkMenu.au3"

; #INDEX# =======================================================================================================================
; Title .........: GUIDarkTheme UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Language ......: English
; Description ...: UDF library for applying dark theme to win32 controls
; Author(s) .....: WildByDesign (including previous code from NoNameCode, argumentum, UEZ, pixelsearch)
; Version .......: 1.0.0
; Note ..........: Window messages used: WM_CTLCOLOR, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX, WM_NOTIFY, WM_SIZE
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _GUIDarkTheme_ApplyDark
; _GUIDarkTheme_ApplyLight
; _GUIDarkTheme_GUISetDarkTheme
; _GUIDarkTheme_GUICtrlAllSetDarkTheme
; _GUIDarkTheme_GUICtrlSetDarkTheme
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __GUIDarkTheme_WM_CTLCOLOR
; __GUIDarkTheme_WM_NOTIFY
; __GUIDarkTheme_WM_SIZE
; __GUIDarkTheme_OnExit
; __GUIDarkTheme_TabProc
; __GUIDarkTheme_ListViewProc
; __GUIDarkTheme_DateProc
; __GUIDarkTheme_hWnd2Styles
; __GUIDarkTheme_GetStyleString
; __GUIDarkTheme_GetCtrlStyleString
; __GUIDarkTheme_GetCtrlStyleString2
; __GUIDarkTheme_CheckedPNG
; __GUIDarkTheme_UncheckedPNG
; __GUIDarkTheme_GetImages
; __GUIDarkTheme_SizeboxProc
; __GUIDarkTheme_CreateDots
; __GUIDarkTheme_StatusRatio
; __GUIDarkTheme_CreateSizebox
; ===============================================================================================================================

; #GLOBAL VARIABLES# ============================================================================================================
Global $g_iBkColor = 0x1c1c1c
; ===============================================================================================================================

; #INTERNAL_USE_ONLY GLOBAL VARIABLES # =========================================================================================
Global $g_iRevision = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR")
Global $g_b24H2Plus = False
If @OSBuild >= 26100 And $g_iRevision >= 6899 Then $g_b24H2Plus = True
Global $g_hGuiWnd
Global $g_hBrushEdit = 0
Global $g_hTab_CB, $g_pTab_CB, $g_hProc, $g_hTab
Global $g_hWndListViewHeader, $g_hWndListView
Global $g__ListView_wProcNew = 0, $g__ListView_wProcOld = 0
Global $g_hDateProc_CB, $g_pDateProc_CB, $g_hDateOldProc, $g_hDate
Global $g_bHover = False
Global $g_iGripPos = 1
Global $g_hOldProc = 0, $g_hSizeboxProc = 0
Global $g_hStatus, $g_hGripSize, $g_hSizebox, $g_hDots, $g_iHeight, $g_aRatioW, $g_hCursor
; ===============================================================================================================================

; #GLOBAL CONSTANTS# ============================================================================================================
Global Const $PRF_CLIENT = 0x0004, $DTP_BORDER = 0x404040, $DTP_BG_DARK = $COLOR_CONTROL_BG, $DTP_TEXT_LIGHT = $COLOR_TEXT_LIGHT, $DTP_BORDER_LIGHT = 0xD8D8D8

Global Const $tagNMCUSTOMDRAW = _
		$tagNMHDR & ";" & _                                    ; Contains NM_CUSTOMDRAW / NMHDR header among other things
		"dword dwDrawStage;" & _                               ; Current drawing stage (CDDS_*)
		"handle hdc;" & _                                      ; Device Context Handle
		"long left;long top;long right;long bottom;" & _       ; Drawing rectangle
		"dword_ptr dwItemSpec;" & _                            ; Item index or other info (depending on the control)
		"uint uItemState;" & _                                 ; State Flags (CDIS_SELECTED, CDIS_FOCUS etc.)
		"lparam lItemlParam"                                   ; lParam set by the item (e.g., via LVITEM.lParam)

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
		[0x00C00000, 'WS_CAPTION'], _      ; (WS_BORDER | WS_DLGFRAME)
		[0x00800000, 'WS_BORDER'], _
		[0x00400000, 'WS_DLGFRAME'], _
		[0x00200000, 'WS_VSCROLL'], _
		[0x00100000, 'WS_HSCROLL'], _
		[0x00080000, 'WS_SYSMENU'], _
		[0x00040000, 'WS_SIZEBOX'], _
		[0x00020000, '! WS_MINIMIZEBOX ! WS_GROUP'], _ ; ! GUI ! Control
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
		[0x0009, 'BS_AUTORADIOBUTTON'], _
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
		[0x00400000, 'ES_VERTICAL'], _      ; Asian-language support only (msdn)
		[0x00080000, 'ES_NOIME'], _         ; ditto
		[0x00040000, 'ES_SELFIME'], _       ; ditto
		[0x00008000, 'ES_SAVESEL'], _
		[0x00004000, 'ES_SUNKEN'], _
		[0x00002000, 'ES_DISABLENOSCROLL'], _ ; same value as 'ES_NUMBER' => issue ?
		[0x00000008, 'ES_NOOLEDRAGDROP']]   ; same value as 'ES_UPPERCASE' but RichRdit controls do not support 'ES_UPPERCASE' style (msdn)

Global Const $_g__Style_Scrollbar[5][2] = _
		[[0x0010, 'SBS_SIZEGRIP'], _
		[0x0008, 'SBS_SIZEBOX'], _
		[0x0004, 'SBS_RIGHTALIGN or SBS_BOTTOMALIGN'], _ ; i.e. use SBS_RIGHTALIGN with SBS_VERT, use SBS_BOTTOMALIGN with SBS_HORZ (msdn)
		[0x0002, 'SBS_LEFTALIGN or SBS_TOPALIGN'], _ ; i.e. use SBS_LEFTALIGN  with SBS_VERT, use SBS_TOPALIGN    with SBS_HORZ (msdn)
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

; ===============================================================================================================================

; GDI+ Startup
_GDIPlus_Startup()

OnAutoItExitRegister(__GUIDarkTheme_OnExit)

Func __GUIDarkTheme_OnExit()
	If $g_hDateOldProc Then __WinAPI_SetWindowLong($g_hDate, $GWL_WNDPROC, $g_hDateOldProc)
	If $g_hDateProc_CB Then DllCallbackFree($g_hDateProc_CB)
	If $g_hBrushEdit Then __WinAPI_DeleteObject($g_hBrushEdit)
	If $g__ListView_wProcOld Then __WinAPI_SetWindowLong($g_hWndListView, $GWL_WNDPROC, $g__ListView_wProcOld)
	If $g__ListView_wProcNew Then DllCallbackFree($g__ListView_wProcNew)
	; Restore the original window procedure for the tab control
	If $g_hProc Then __WinAPI_SetWindowLong($g_hTab, $GWL_WNDPROC, $g_hProc)
	If $g_hTab_CB Then DllCallbackFree($g_hTab_CB)
	; statusbar
	If $g_hDots Then _GDIPlus_BitmapDispose($g_hDots)
	If $g_hCursor Then _WinAPI_DestroyCursor($g_hCursor)
	If $g_hOldProc Then __WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, $g_hOldProc)
	If $g_hSizeboxProc Then DllCallbackFree($g_hSizeboxProc)
	_GDIPlus_Shutdown()
EndFunc   ;==>__GUIDarkTheme_OnExit

Func __GUIDarkTheme_WM_CTLCOLOR($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg
	Local $hDC = $wParam
	Local $hCtrl = $lParam

	Switch __WinAPI_GetClassName($hCtrl)
		Case 'Static'
			; set transparent background
			__WinAPI_SetBkMode($hDC, $TRANSPARENT)
			; set text color (if necessary) - e.g., white
			__WinAPI_SetTextColor($hDC, __GUIDarkMenu_ColorToCOLORREF($COLOR_TEXT_LIGHT))
			; return NULL_BRUSH (stock object), so Windows does NOT fill with your dark brush
			Local $hNull = __WinAPI_GetStockObject(5) ; 5 = NULL_BRUSH
			If $hNull Then Return $hNull
			; Fallback if not available:
			Return $GUI_RUNDEFMSG
	EndSwitch

	; --- Default behavior for all other statics / controls ---
	__WinAPI_SetTextColor($hDC, __GUIDarkMenu_ColorToCOLORREF($COLOR_TEXT_LIGHT))

	Local $hBrush = $g_hBrushEdit

	_WinAPI_SetBkColor($hDC, __GUIDarkMenu_ColorToCOLORREF($COLOR_CONTROL_BG))
	__WinAPI_SetBkMode($hDC, $TRANSPARENT)

	Return $hBrush
EndFunc   ;==>__GUIDarkTheme_WM_CTLCOLOR

Func __GUIDarkTheme_TabProc($hWnd, $iMsg, $wParam, $lParam)
	; Custom window procedure for tab control with Dark Mode
	Switch $iMsg
		Case $WM_ERASEBKGND
			Return 1 ; Prevent background erase to avoid flicker

		Case $WM_PAINT
			Local $tPaint = DllStructCreate($tagPAINTSTRUCT)
			Local $hDC = DllCall($hUser32Dll, "handle", "BeginPaint", "hwnd", $hWnd, "struct*", $tPaint)
			If @error Or Not $hDC[0] Then Return __WinAPI_CallWindowProc($g_hProc, $hWnd, $iMsg, $wParam, $lParam)
			$hDC = $hDC[0]

			; Get client rectangle
			Local $tClient = __WinAPI_GetClientRect($hWnd)
			If Not IsDllStruct($tClient) Then
				__WinAPI_EndPaint($hWnd, $tPaint)
				Return 0
			EndIf

			Local $iWidth = $tClient.Right
			Local $iHeight = $tClient.Bottom

			; Create memory DC for double buffering
			Local $hMemDC = __WinAPI_CreateCompatibleDC($hDC)
			Local $hBitmap = __WinAPI_CreateCompatibleBitmap($hDC, $iWidth, $iHeight)
			Local $hOldBmp = __WinAPI_SelectObject($hMemDC, $hBitmap)

			; Fill background but exclude overlapping GUI controls from painting
			Local $hParent = __WinAPI_GetParent($hWnd)
			Local $hChild = __WinAPI_GetWindow($hParent, $GW_CHILD)
			Local $tCR, $tPR = __WinAPI_GetWindowRect($hWnd)
			Local $left, $top, $right, $bottom

			While $hChild
				If $hChild <> $hWnd And __WinAPI_IsWindowVisible($hChild) Then
					$tCR = __WinAPI_GetWindowRect($hChild)
					; Only exclude controls that lie fully within the tab control area
					If $tCR.left >= $tPR.left And $tCR.right <= $tPR.right And $tCR.top >= $tPR.top And $tCR.bottom <= $tPR.bottom Then
						$left = $tCR.left - $tPR.left
						$top = $tCR.top - $tPR.top
						$right = $tCR.right - $tPR.left
						$bottom = $tCR.bottom - $tPR.top
						; Exclude from offscreen bitmap (prevents black fill)
						DllCall($hGdi32Dll, "int", "ExcludeClipRect", "handle", $hMemDC, "int", $left, "int", $top, "int", $right, "int", $bottom)
						; Exclude from screen DC (prevents BitBlt overwrite)
						DllCall($hGdi32Dll, "int", "ExcludeClipRect", "handle", $hDC, "int", $left, "int", $top, "int", $right, "int", $bottom)
					EndIf
				EndIf
				$hChild = __WinAPI_GetWindow($hChild, $GW_HWNDNEXT)
			WEnd

			; Fill background
			Local $hBrush = __WinAPI_CreateSolidBrush(__GUIDarkMenu_ColorToCOLORREF($COLOR_CONTROL_BG))
			__WinAPI_FillRect($hMemDC, $tClient, $hBrush)
			__WinAPI_DeleteObject($hBrush)

			; Get tab info
			Local $iTabCount = __SendMessage($hWnd, $TCM_GETITEMCOUNT, 0, 0)
			Local $iCurSel = __SendMessage($hWnd, $TCM_GETCURSEL, 0, 0)

			; Setup font
			Local $hFont = __SendMessage($hWnd, $WM_GETFONT, 0, 0)
			If Not $hFont Then $hFont = __WinAPI_GetStockObject($DEFAULT_GUI_FONT)
			Local $hOldFont = __WinAPI_SelectObject($hMemDC, $hFont)

			__WinAPI_SetBkMode($hMemDC, $TRANSPARENT)
			__WinAPI_SetTextColor($hMemDC, __GUIDarkMenu_ColorToCOLORREF($COLOR_TEXT_LIGHT))

			; Draw each tab
			Local $tRect, $iLeft, $iTop, $iRight, $iBottom, $tItem, $tText, $bSelected, $iTabColor, $hTabBrush, $tTabRect, _
					$sText, $hPen, $hOldPen, $hPenSep, $hOldPenSep, $tTextRect, $hBorderPen, $hOldBorderPen, $hNullBrush, $hOldBorderBrush

			For $i = 0 To $iTabCount - 1
				; Get tab rectangle using TCM_GETITEMRECT
				$tRect = DllStructCreate($tagRECT)
				Local $aResult = DllCall($hUser32Dll, "lresult", "SendMessageW", _
						"hwnd", $hWnd, _
						"uint", $TCM_GETITEMRECT, _
						"wparam", $i, _
						"struct*", $tRect)
				If @error Or Not $aResult[0] Then ContinueLoop

				$iLeft = $tRect.Left
				$iTop = $tRect.Top
				$iRight = $tRect.Right
				$iBottom = $tRect.Bottom

				; Skip if rectangle is invalid
				If $iLeft >= $iRight Or $iTop >= $iBottom Then ContinueLoop

				; Get tab text
				$tItem = DllStructCreate("uint Mask;dword dwState;dword dwStateMask;ptr pszText;int cchTextMax;int iImage;lparam lParam")
				$tText = DllStructCreate("wchar Text[256]")
				With $tItem
					.Mask = 0x0001 ; TCIF_TEXT
					.pszText = DllStructGetPtr($tText)
					.cchTextMax = 256
				EndWith

				DllCall($hUser32Dll, "lresult", "SendMessageW", _
						"hwnd", $hWnd, _
						"uint", $TCM_GETITEMW, _
						"wparam", $i, _
						"struct*", $tItem)

				$sText = DllStructGetData($tText, "Text")

				; Draw tab background
				$bSelected = ($i = $iCurSel)
				$iTabColor = $bSelected ? __WinAPI_ColorAdjustLuma($COLOR_CONTROL_BG, 20) : __WinAPI_ColorAdjustLuma($COLOR_CONTROL_BG, 10)
				$hTabBrush = __WinAPI_CreateSolidBrush($iTabColor)

				$tTabRect = DllStructCreate($tagRECT)
				With $tTabRect
					.Left = $iLeft
					.Top = $iTop
					.Right = $iRight
					.Bottom = $iBottom
				EndWith

				__WinAPI_FillRect($hMemDC, $tTabRect, $hTabBrush)
				__WinAPI_DeleteObject($hTabBrush)

				; Draw selection indicator (top border for selected tab)
				If $bSelected Then
					$hPen = __WinAPI_CreatePen(0, 2, __GUIDarkMenu_ColorToCOLORREF(0x0078D4)) ; Blue accent
					$hOldPen = __WinAPI_SelectObject($hMemDC, $hPen)
					__WinAPI_MoveTo($hMemDC, $iLeft, $iTop)
					__WinAPI_LineTo($hMemDC, $iRight - 2, $iTop)
					__WinAPI_SelectObject($hMemDC, $hOldPen)
					__WinAPI_DeleteObject($hPen)
				EndIf

				; Draw separator between tabs
				If $i < $iTabCount - 1 Then
					$hPenSep = __WinAPI_CreatePen(0, 1, __GUIDarkMenu_ColorToCOLORREF($COLOR_BORDER))
					$hOldPenSep = __WinAPI_SelectObject($hMemDC, $hPenSep)
					__WinAPI_MoveTo($hMemDC, $iRight - 1, $iTop + 4)
					__WinAPI_LineTo($hMemDC, $iRight - 1, $iBottom - 4)
					__WinAPI_SelectObject($hMemDC, $hOldPenSep)
					__WinAPI_DeleteObject($hPenSep)
				EndIf

				; Draw text centered in tab
				$tTextRect = DllStructCreate($tagRECT)
				With $tTextRect
					.Left = $iLeft + 6
					.Top = $iTop + 3
					.Right = $iRight - 6
					.Bottom = $iBottom - 3
				EndWith
				DllCall($hUser32Dll, "int", "DrawTextW", _
						"handle", $hMemDC, _
						"wstr", $sText, _
						"int", -1, _
						"struct*", $tTextRect, _
						"uint", BitOR($DT_CENTER, $DT_VCENTER, $DT_SINGLELINE))
			Next

			; Draw border around entire control
			$hBorderPen = __WinAPI_CreatePen(0, 1, __GUIDarkMenu_ColorToCOLORREF($COLOR_BORDER))
			$hOldBorderPen = __WinAPI_SelectObject($hMemDC, $hBorderPen)
			$hNullBrush = __WinAPI_GetStockObject(5) ; NULL_BRUSH
			$hOldBorderBrush = __WinAPI_SelectObject($hMemDC, $hNullBrush)

			DllCall($hGdi32Dll, "bool", "Rectangle", "handle", $hMemDC, "int", 0, "int", 0, "int", $iWidth, "int", $iHeight)

			__WinAPI_SelectObject($hMemDC, $hOldBorderPen)
			__WinAPI_SelectObject($hMemDC, $hOldBorderBrush)
			__WinAPI_DeleteObject($hBorderPen)

			; Copy to screen
			__WinAPI_BitBlt($hDC, 0, 0, $iWidth, $iHeight, $hMemDC, 0, 0, $SRCCOPY)

			; Cleanup
			__WinAPI_SelectObject($hMemDC, $hOldFont)
			__WinAPI_SelectObject($hMemDC, $hOldBmp)
			__WinAPI_DeleteObject($hBitmap)
			__WinAPI_DeleteDC($hMemDC)

			__WinAPI_EndPaint($hWnd, $tPaint)
			Return 0
	EndSwitch

	Return __WinAPI_CallWindowProc($g_hProc, $hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>__GUIDarkTheme_TabProc

Func __GUIDarkTheme_ListViewProc($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	If $iMsg = $WM_NOTIFY Then
		Local $tNMHDR, $hWndFrom, $iCode, $hDC
		$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
		$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
		$iCode = DllStructGetData($tNMHDR, "Code")

		Switch $hWndFrom
			Case $g_hWndListViewHeader
				Switch $iCode
					Case $NM_CUSTOMDRAW
						Local $tCustDraw = DllStructCreate($tagNMLVCUSTOMDRAW, $lParam)
						Switch DllStructGetData($tCustDraw, "dwDrawStage")
							Case $CDDS_PREPAINT
								Return $CDRF_NOTIFYITEMDRAW
							Case $CDDS_ITEMPREPAINT
								$hDC = DllStructGetData($tCustDraw, "hDC")
								DllCall($hGdi32Dll, "int", "SetTextColor", "handle", $hDC, "dword", __WinAPI_SwitchColor($COLOR_TEXT_LIGHT))
								Return $CDRF_NEWFONT
								Return $CDRF_SKIPDEFAULT
						EndSwitch
				EndSwitch
		EndSwitch
	EndIf
	;pass the unhandled messages to default WindowProc
	Local $aResult = DllCall($hUser32Dll, "lresult", "CallWindowProcW", "ptr", $g__ListView_wProcOld, _
			"hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, "lparam", $lParam)
	If @error Then Return -1
	Return $aResult[0]
EndFunc   ;==>__GUIDarkTheme_ListViewProc

Func __GUIDarkTheme_DateProc($hWnd, $iMsg, $wParam, $lParam)
	Local $iRet
	Switch $iMsg

		Case $WM_PAINT
			Local $tPaint = DllStructCreate($tagPAINTSTRUCT)
			Local $hDC = __WinAPI_BeginPaint($hWnd, $tPaint)

			Local $tClient = __WinAPI_GetClientRect($hWnd)
			Local $iW = $tClient.Right
			Local $iH = $tClient.Bottom

			; --- Memory DC for flicker-free rendering ---
			Local $hMemDC = __WinAPI_CreateCompatibleDC($hDC)
			Local $hBitmap = __WinAPI_CreateCompatibleBitmap($hDC, $iW, $iH)
			Local $hOldBmp = __WinAPI_SelectObject($hMemDC, $hBitmap)

			; 1. Let Windows draw the light-mode control into memory DC
			__WinAPI_CallWindowProc($g_hDateOldProc, $hWnd, $WM_PRINTCLIENT, $hMemDC, $PRF_CLIENT)

			; 2. Invert all pixels (background becomes black, text white, selection orange)
			Local $tRect = DllStructCreate($tagRECT)
			$tRect.right = $iW
			$tRect.bottom = $iH
			__WinAPI_InvertRect($hMemDC, $tRect)

			; --- 3. PIXEL HACK: destroy orange highlight & set background color ---
			Local $iSize = $iW * $iH
			Local $tPixels = DllStructCreate("dword c[" & $iSize & "]")
			; Load pixel array directly from bitmap memory
			Local $iBytes = DllCall($hGdi32Dll, "long", "GetBitmapBits", "handle", $hBitmap, "long", $iSize * 4, "ptr", DllStructGetPtr($tPixels))[0]

			If $iBytes = $iSize * 4 Then
				Local $iPixel, $r, $g, $b, $iGray
				For $i = 1 To $iSize
					$iPixel = $tPixels.c(($i))

					; Split into color channels
					$b = BitAND($iPixel, 0xFF)
					$g = BitAND(BitShift($iPixel, 8), 0xFF)
					$r = BitAND(BitShift($iPixel, 16), 0xFF)

					; Convert to grayscale (orange becomes mid-gray)
					$iGray = Int(($r + $g + $b) / 3)

					; Very dark pixel = inverted white background
					If $iGray < 15 Then
						$iPixel = $DTP_BG_DARK ; Replace with exact GUI background color
					Else
						; Grayscale value for text (white) and selection (gray)
						; (negative BitShift shifts left in AutoIt)
						$iPixel = BitOR(BitShift($iGray, -16), BitShift($iGray, -8), $iGray)
					EndIf

					$tPixels.c(($i)) = $iPixel
				Next
				; Write cleaned pixels back into the bitmap
				DllCall($hGdi32Dll, "long", "SetBitmapBits", "handle", $hBitmap, "long", $iSize * 4, "ptr", DllStructGetPtr($tPixels))
			EndIf
			; --- END PIXEL HACK ---

			; --- Border color (hover effect) ---
			Local $iBorderColor = $DTP_BORDER
			If __WinAPI_GetFocus() = $hWnd Then $iBorderColor = $DTP_BORDER
			Local $tCursorPos = DllStructCreate($tagPOINT)
			DllCall($hUser32Dll, "bool", "GetCursorPos", "struct*", $tCursorPos)
			DllCall($hUser32Dll, "bool", "ScreenToClient", "hwnd", $hWnd, "struct*", $tCursorPos)
			If $tCursorPos.X >= 0 And $tCursorPos.X <= $iW And $tCursorPos.Y >= 0 And $tCursorPos.Y <= $iH Then
				$iBorderColor = $DTP_BORDER_LIGHT
			EndIf

			; --- Draw border ---
			Local $hPen = __WinAPI_CreatePen(0, 1, __GUIDarkMenu_ColorToCOLORREF($iBorderColor))
			Local $hNullBr = __WinAPI_GetStockObject(5)
			Local $hOldPen = __WinAPI_SelectObject($hMemDC, $hPen)
			Local $hOldBr = __WinAPI_SelectObject($hMemDC, $hNullBr)
			DllCall($hGdi32Dll, "bool", "Rectangle", "handle", $hMemDC, "int", 0, "int", 0, "int", $iW, "int", $iH)
			__WinAPI_SelectObject($hMemDC, $hOldPen)
			__WinAPI_SelectObject($hMemDC, $hOldBr)
			__WinAPI_DeleteObject($hPen)

			; --- Copy finished result to screen in one step (no flicker) ---
			__WinAPI_BitBlt($hDC, 0, 0, $iW, $iH, $hMemDC, 0, 0, $SRCCOPY)

			; --- Cleanup ---
			__WinAPI_SelectObject($hMemDC, $hOldBmp)
			__WinAPI_DeleteObject($hBitmap)
			__WinAPI_DeleteDC($hMemDC)
			__WinAPI_EndPaint($hWnd, $tPaint)
			Return 0

		Case $WM_ERASEBKGND
			Return 1

		Case $WM_SETFOCUS, $WM_KILLFOCUS, $WM_LBUTTONDOWN, $WM_LBUTTONUP
			$iRet = __WinAPI_CallWindowProc($g_hDateOldProc, $hWnd, $iMsg, $wParam, $lParam)
			__WinAPI_InvalidateRect($hWnd, 0, False)
			Return $iRet

		Case $WM_MOUSEMOVE
			$iRet = __WinAPI_CallWindowProc($g_hDateOldProc, $hWnd, $iMsg, $wParam, $lParam)
			If Not $g_bHover Then
				$g_bHover = True
				__WinAPI_InvalidateRect($hWnd, 0, False)
			EndIf
			Return $iRet

		Case $WM_MOUSELEAVE
			$g_bHover = False
			__WinAPI_InvalidateRect($hWnd, 0, False)

	EndSwitch

	Return __WinAPI_CallWindowProc($g_hDateOldProc, $hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>__GUIDarkTheme_DateProc

Func __GUIDarkTheme_WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	Local Static $iTheme  ; argumentum
	Local $hWndFrom, $iCode, $tNMHDR, $iCtrl
	$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	$hWndFrom = HWnd($tNMHDR.hWndFrom)
	$iCode = $tNMHDR.Code

	; --- Slider (msctls_trackbar32) custom drawing ---
	If $iCode = $NM_CUSTOMDRAW And StringLower(__WinAPI_GetClassName($hWndFrom)) = "msctls_trackbar32" Then
		Local $tNMCD = DllStructCreate($tagNMCUSTOMDRAW, $lParam)
		Local $dwStage = $tNMCD.dwDrawStage
		Local $hDC = $tNMCD.hdc
		Local $dwItemSpec = $tNMCD.dwItemSpec

		Switch $dwStage
			Case $CDDS_PREPAINT
				; Remove focus rectangle from slider
				DllStructSetData($tNMCD, 'ItemState', BitXOR(DllStructGetData($tNMCD, 'ItemState'), $CDIS_FOCUS))
				Return $CDRF_NOTIFYITEMDRAW ; request per-item notifications

			Case $CDDS_ITEMPREPAINT
				Local Const $TBCD_TICS = 0x0001    ; tick marks
				Local Const $TBCD_THUMB = 0x0002   ; draggable thumb
				Local Const $TBCD_CHANNEL = 0x0003 ; slider channel/track

				Local $tRect = DllStructCreate($tagRECT)
				$tRect.left = $tNMCD.left
				$tRect.top = $tNMCD.top
				$tRect.right = $tNMCD.right
				$tRect.bottom = $tNMCD.bottom

				Switch $dwItemSpec
					Case $TBCD_TICS
						; Let Windows draw tick marks normally
						Return $CDRF_DODEFAULT

					Case $TBCD_THUMB
						; Draw thumb as a pentagon (rectangle + downward arrow)
						Local $iL = $tNMCD.left
						Local $iT = $tNMCD.top
						Local $iR = $tNMCD.right - 1     ; -1 to stay within bounds and avoid paint artifacts
						Local $iB = $tNMCD.bottom
						Local $iMid = ($iL + $iR) / 2         ; horizontal center (tip of arrow)
						Local $iSplit = $iB - ($iR - $iL) / 2 ; y-position where rectangle ends and arrow begins

						; Pentagon points: top-left, top-right, right-middle, bottom-tip, left-middle
						Local $tPoints = DllStructCreate("int p[10]")
						$tPoints.p((1)) = $iL
						$tPoints.p((2)) = $iT
						$tPoints.p((3)) = $iR
						$tPoints.p((4)) = $iT
						$tPoints.p((5)) = $iR
						$tPoints.p((6)) = $iSplit
						$tPoints.p((7)) = $iMid
						$tPoints.p((8)) = $iB
						$tPoints.p((9)) = $iL
						$tPoints.p((10)) = $iSplit

						; Fill and outline thumb with blue accent color
						Local $hBrush = __WinAPI_CreateSolidBrush(__GUIDarkMenu_ColorToCOLORREF(0x0078D4))
						Local $hPen = __WinAPI_CreatePen(0, 1, __GUIDarkMenu_ColorToCOLORREF($COLOR_CONTROL_BG))
						Local $hOldBrush = __WinAPI_SelectObject($hDC, $hBrush)
						Local $hOldPen = __WinAPI_SelectObject($hDC, $hPen)
						DllCall($hGdi32Dll, "bool", "Polygon", "handle", $hDC, "struct*", $tPoints, "int", 5)
						__WinAPI_SelectObject($hDC, $hOldBrush)
						__WinAPI_SelectObject($hDC, $hOldPen)
						__WinAPI_DeleteObject($hBrush)
						__WinAPI_DeleteObject($hPen)
						Return $CDRF_SKIPDEFAULT ; skip default drawing

					Case $TBCD_CHANNEL
						; Fill channel with border color
						Local $hBrushChan = __WinAPI_CreateSolidBrush(__GUIDarkMenu_ColorToCOLORREF($COLOR_BORDER))
						__WinAPI_FillRect($hDC, $tRect, $hBrushChan)
						__WinAPI_DeleteObject($hBrushChan)
						Return $CDRF_SKIPDEFAULT ; skip default drawing

				EndSwitch
		EndSwitch
	EndIf

	If __WinAPI_GetClassName($hWndFrom) = "SysDateTimePick32" Then
		Switch $iCode
			Case $NM_SETFOCUS
				$iTheme = __WinAPI_GetThemeAppProperties() ; argumentum
				; Disable the visual theme when the DateTime control receives focus
				__WinAPI_SetThemeAppProperties(0)

			Case $DTN_DROPDOWN ;, $EVENT_OBJECT_CREATE
				; Apply dark colors when the calendar dropdown appears
				$iCtrl = _GUICtrlDTP_GetMonthCal($hWndFrom)
				__WinAPI_SetWindowTheme($iCtrl, "", "")
				_GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TEXT, $COLOR_TEXT_LIGHT)
				_GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TITLEBK, $COLOR_CONTROL_BG)
				_GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TITLETEXT, $COLOR_TEXT_LIGHT)
				_GUICtrlMonthCal_SetColor($iCtrl, $MCSC_BACKGROUND, $COLOR_CONTROL_BG)
				_GUICtrlMonthCal_SetColor($iCtrl, $MCSC_MONTHBK, $COLOR_CONTROL_BG)
				_GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TRAILINGTEXT, $COLOR_TEXT_LIGHT)

			Case $DTN_CLOSEUP
				; Reset visual theme
				__WinAPI_SetThemeAppProperties($iTheme)  ; argumentum
				; Calendar will closed

		EndSwitch
	EndIf
	#cs
	    ; --- Per-control notification handling ---
	    Switch $hWndFrom
	        Case $g_hDate ;thanks to argumentum for the code :-)
	            Switch $iCode
	                Case $NM_SETFOCUS
	                    $iTheme = __WinAPI_GetThemeAppProperties() ; argumentum
	                    ; Disable the visual theme when the DateTime control receives focus
	                    __WinAPI_SetThemeAppProperties(0)

	                Case $DTN_DROPDOWN;, $EVENT_OBJECT_CREATE
	                    ; Apply dark colors when the calendar dropdown appears
	                    $iCtrl = _GUICtrlDTP_GetMonthCal($hWndFrom)
	                    __WinAPI_SetWindowTheme($iCtrl, "", "")
	                    _GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TEXT, $COLOR_TEXT_LIGHT)
	                    _GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TITLEBK, $COLOR_CONTROL_BG)
	                    _GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TITLETEXT, $COLOR_TEXT_LIGHT)
	                    _GUICtrlMonthCal_SetColor($iCtrl, $MCSC_BACKGROUND, $COLOR_CONTROL_BG)
	                    _GUICtrlMonthCal_SetColor($iCtrl, $MCSC_MONTHBK, $COLOR_CONTROL_BG)
	                    _GUICtrlMonthCal_SetColor($iCtrl, $MCSC_TRAILINGTEXT, $COLOR_TEXT_LIGHT)

	                Case $DTN_CLOSEUP
	                    ; Reset visual theme
	                    __WinAPI_SetThemeAppProperties($iTheme)  ; argumentum
	                    ; Calendar will closed

	            EndSwitch
	    EndSwitch
	#ce
	Return $GUI_RUNDEFMSG
EndFunc   ;==>__GUIDarkTheme_WM_NOTIFY

Func __GUIDarkTheme_hWnd2Styles($hWnd)
	Return __GUIDarkTheme_GetCtrlStyleString(__WinAPI_GetWindowLong($hWnd, $GWL_STYLE), __WinAPI_GetWindowLong($hWnd, $GWL_EXSTYLE), __WinAPI_GetClassName($hWnd))
EndFunc   ;==>__GUIDarkTheme_hWnd2Styles

Func __GUIDarkTheme_GetStyleString($iStyle, $fExStyle)
	ConsoleWrite('+ Func __GUIDarkTheme_GetStyleString(' & $iStyle & ', ' & $fExStyle & ')' & @CRLF)
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
EndFunc   ;==>__GUIDarkTheme_GetStyleString

Func __GUIDarkTheme_GetCtrlStyleString($iStyle, $fExStyle, $sClass, $iLVExStyle = 0)

	If $sClass = "AutoIt v3 GUI" Or $sClass = "#32770" Or $sClass = "MDIClient" Then ; control = child GUI, dialog box (msgbox) etc...
		Return __GUIDarkTheme_GetStyleString($iStyle, 0)
	EndIf

	If StringLeft($sClass, 8) = "RichEdit" Then $sClass = "RichEdit" ; RichEdit, RichEdit20A, RichEdit20W, RichEdit50A, RichEdit50W

	Local $Text = ''

	__GUIDarkTheme_GetCtrlStyleString2($iStyle, $Text, $sClass, $iLVExStyle) ; 4th param. in case $sClass = "Ex_SysListView32" (special treatment)

	If $sClass = "ReBarWindow32" Or $sClass = "ToolbarWindow32" Or $sClass = "msctls_statusbar32" Then
		$sClass = "Common" ; "for rebar controls, toolbar controls, and status windows" (msdn)
		__GUIDarkTheme_GetCtrlStyleString2($iStyle, $Text, $sClass)
	ElseIf $sClass = "RichEdit" Then
		$sClass = "Edit" ; "Richedit controls also support many edit control styles (not all)" (msdn)
		__GUIDarkTheme_GetCtrlStyleString2($iStyle, $Text, $sClass)
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
EndFunc   ;==>__GUIDarkTheme_GetCtrlStyleString

;=====================================================================
Func __GUIDarkTheme_GetCtrlStyleString2(ByRef $iStyle, ByRef $Text, $sClass, $iLVExStyle = 0)

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
EndFunc   ;==>__GUIDarkTheme_GetCtrlStyleString2

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUIDarkTheme_GUISetDarkTheme
; Description ...: Sets the theme for a specified window to either dark or light mode on Windows 10.
; Syntax ........: _GUIDarkTheme_GUISetDarkTheme($hwnd, $dark_theme = True)
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
Func _GUIDarkTheme_GUISetDarkTheme($hWnd, $bEnableDarkTheme = True)
	Local $iPreferredAppMode = ($bEnableDarkTheme == True) ? $APPMODE_FORCEDARK : $APPMODE_FORCELIGHT
	Local $iGUI_BkColor = ($bEnableDarkTheme == True) ? $COLOR_BG_DARK : _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_3DFACE))
	; update global GUI handle
	$g_hGuiWnd = $hWnd
	__WinAPI_SetPreferredAppMode($iPreferredAppMode)
	__WinAPI_RefreshImmersiveColorPolicyState()
	__WinAPI_FlushMenuThemes()
	GUISetBkColor($iGUI_BkColor, $hWnd)
	_GUIDarkTheme_GUICtrlSetDarkTheme($hWnd, $bEnableDarkTheme)            ;To Color the GUI's own Scrollbar
	__WinAPI_DwmSetWindowAttribute($hWnd, $DWMWA_USE_IMMERSIVE_DARK_MODE, $bEnableDarkTheme)
EndFunc   ;==>_GUIDarkTheme_GUISetDarkTheme

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUIDarkTheme_GUICtrlAllSetDarkTheme
; Description ...: Sets the dark theme to all existing sub Controls from a GUI
; Syntax ........: _GUIDarkTheme_GUICtrlAllSetDarkTheme($g_hGuiWnd[, $bEnableDarkTheme = True, $bPreferNewTheme = False])
; Parameters ....: $g_hGuiWnd                - GUI handle
;                  $bEnableDarkTheme    - [optional] a boolean value. Default is True.
;                  $bPreferNewTheme 	- Prefer the newer DarkMode_DarkTheme theme over DarkMode_Explorer when possible.
;                                         (Default is False. DarkMode_DarkTheme is only available on Win11 26100.6899 and higher)
; Return values .: None
; Author ........: NoName
; Modified ......: WildByDesign
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GUIDarkTheme_GUICtrlAllSetDarkTheme($g_hGuiWnd, $bEnableDarkTheme = True, $bPreferNewTheme = False)
	Local $aCtrls = __WinAPI_EnumChildWindows($g_hGuiWnd, False)
	For $i = 1 To $aCtrls[0][0]
		_GUIDarkTheme_GUICtrlSetDarkTheme($aCtrls[$i][0], $bEnableDarkTheme, $bPreferNewTheme)
	Next
	Local $aCtrlsEx = __WinAPI_EnumProcessWindows(0, False) ; allows getting handles for tooltips_class32, ComboLBox, etc.
	For $i = 1 To $aCtrlsEx[0][0]
		_GUIDarkTheme_GUICtrlSetDarkTheme($aCtrlsEx[$i][0], $bEnableDarkTheme, $bPreferNewTheme)
	Next
	Return $aCtrls
EndFunc   ;==>_GUIDarkTheme_GUICtrlAllSetDarkTheme


; #FUNCTION# ====================================================================================================================
; Name ..........: _GUIDarkTheme_GUICtrlSetDarkTheme
; Description ...: Sets the dark theme for a specified control.
; Syntax ........: _GUIDarkTheme_GUICtrlSetDarkTheme($vCtrl, $bEnableDarkTheme = True, $bPreferNewTheme = False)
; Parameters ....: $vCtrl            - The control handle or identifier.
;                  $bEnableDarkTheme - If True, enables the dark theme; if False, disables it.
;                                      (Default is True for enabling dark theme.)
;                  $bPreferNewTheme  - Prefer the newer DarkMode_DarkTheme theme over DarkMode_Explorer when possible.
;                                      (Default is False. DarkMode_DarkTheme is only available on Win11 26100.6899 and higher)
; Return values .: Success: True
;                  Failure: False and sets the @error flag:
;                           1: Invalid control handle or identifier.
;                           2: Error while allowing dark mode for the window.
;                           3: Error while setting the window theme.
;                           4: Error while sending the WM_THEMECHANGED message.
; Author ........: NoNameCode
; Modified ......: WildByDesign
; Remarks .......: This function requires the _WinAPI_SetWindowTheme and __WinAPI_AllowDarkModeForWindow functions.
; Related .......:
; Link ..........: http://www.opengate.at/blog/2021/08/dark-mode-win32/
; Example .......: Yes
; ===============================================================================================================================
Func _GUIDarkTheme_GUICtrlSetDarkTheme($vCtrl, $bEnableDarkTheme = True, $bPreferNewTheme = False)
	Local $sThemeName = Null, $sThemeList = Null
	Local $iGUI_Ctrl_Color = ($bEnableDarkTheme == True) ? $COLOR_TEXT_LIGHT : _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_WINDOWTEXT))
	Local $iGUI_Ctrl_BkColor = ($bEnableDarkTheme == True) ? $COLOR_CONTROL_BG : _WinAPI_SwitchColor(_WinAPI_GetSysColor($COLOR_BTNFACE))
	Local $bSpecialBtn = False, $bSpecialLV = False, $bSpecialTV = False
	Local $sStyles
	If Not IsHWnd($vCtrl) Then $vCtrl = GUICtrlGetHandle($vCtrl)
	If Not IsHWnd($vCtrl) Then Return SetError(1, 0, False)
	__WinAPI_AllowDarkModeForWindow($vCtrl, $bEnableDarkTheme)
	If @error <> 0 Then Return SetError(2, @error, False)
	;=========
	;ConsoleWrite(@CRLF & __WinAPI_GetClassName($vCtrl))
	Switch __WinAPI_GetClassName($vCtrl)
		Case 'Button'
			$sStyles = __GUIDarkTheme_hWnd2Styles($vCtrl)
			Switch $bEnableDarkTheme
				Case True
					If $bPreferNewTheme And $g_b24H2Plus Then
						$sThemeName = 'DarkMode_DarkTheme'
					Else
						If StringInStr($sStyles, "BS_GROUPBOX") Or StringInStr($sStyles, "BS_AUTORADIOBUTTON") Then
							$bSpecialBtn = True
						Else
							$sThemeName = 'DarkMode_Explorer'
						EndIf
					EndIf
				Case False
					$sThemeName = 'Explorer'
			EndSwitch

		Case 'msctls_trackbar32'
			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)
			If $bEnableDarkTheme Then GUIRegisterMsg($WM_NOTIFY, "__GUIDarkTheme_WM_NOTIFY")

		Case 'msctls_updown32'
			If $bEnableDarkTheme Then
				$sThemeName = 'DarkMode_Explorer'
			EndIf

		Case 'ListBox'
			Switch $bEnableDarkTheme
				Case True
					If $bPreferNewTheme And $g_b24H2Plus Then
						$sThemeName = 'DarkMode_DarkTheme'
					Else
						$sThemeName = 'DarkMode_Explorer'
					EndIf

					; create brush and register GUI message
					If Not $g_hBrushEdit Then $g_hBrushEdit = __WinAPI_CreateSolidBrush(__GUIDarkMenu_ColorToCOLORREF($COLOR_CONTROL_BG))
					GUIRegisterMsg($WM_CTLCOLORLISTBOX, "__GUIDarkTheme_WM_CTLCOLOR")
				Case False
					$sThemeName = 'Explorer'
			EndSwitch

		Case 'SysTreeView32'
			$sStyles = __GUIDarkTheme_hWnd2Styles($vCtrl)
			Switch $bEnableDarkTheme
				Case True
					If $bPreferNewTheme And $g_b24H2Plus Then
						$sThemeName = 'DarkMode_DarkTheme'
					Else
						$sThemeName = 'DarkMode_Explorer'
					EndIf

					; dark mode checkboxes
					If StringInStr($sStyles, "TVS_CHECKBOXES") Then
						$bSpecialTV = True
					EndIf
				Case False
					$sThemeName = 'Explorer'
			EndSwitch

			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)

		Case 'SysListView32'
			Switch $bEnableDarkTheme
				Case True
					If $bPreferNewTheme And $g_b24H2Plus Then
						$sThemeName = 'DarkMode_DarkTheme'
					Else
						$sThemeName = 'DarkMode_Explorer'
					EndIf

					; register callback to subclass ListView header text color
					$g_hWndListViewHeader = _GUICtrlListView_GetHeader($vCtrl)
					$g_hWndListView = $vCtrl
					$g__ListView_wProcNew = DllCallbackRegister("__GUIDarkTheme_ListViewProc", "ptr", "hwnd;uint;wparam;lparam")
					$g__ListView_wProcOld = __WinAPI_SetWindowLong($vCtrl, $GWL_WNDPROC, DllCallbackGetPtr($g__ListView_wProcNew))

					; checkbox dark mode
					If (BitAND(_GUICtrlListView_GetExtendedListViewStyle($vCtrl), $LVS_EX_CHECKBOXES) = $LVS_EX_CHECKBOXES) Then
						$bSpecialLV = True
					EndIf
				Case False
					$sThemeName = 'Explorer'
			EndSwitch

			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)

		Case 'Edit'
			Switch $bEnableDarkTheme
				Case True
					If $bPreferNewTheme And $g_b24H2Plus Then
						$sThemeName = 'DarkMode_DarkTheme'
					Else
						$sThemeName = 'DarkMode_Explorer'
					EndIf

					; create brush and register GUI message
					If Not $g_hBrushEdit Then $g_hBrushEdit = __WinAPI_CreateSolidBrush(__GUIDarkMenu_ColorToCOLORREF($COLOR_CONTROL_BG))
					GUIRegisterMsg($WM_CTLCOLOREDIT, "__GUIDarkTheme_WM_CTLCOLOR")
				Case False
					$sThemeName = 'Explorer'
			EndSwitch

		Case 'SysHeader32'
			Switch $bEnableDarkTheme
				Case True
					If $bPreferNewTheme And $g_b24H2Plus Then
						$sThemeName = 'DarkMode_DarkTheme'
					Else
						$sThemeName = 'DarkMode_ItemsView'
					EndIf
				Case False
					$sThemeName = 'ItemsView'
			EndSwitch
			$sThemeList = 'Header'

		Case 'Static'
			If $bEnableDarkTheme Then
				GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			EndIf

		Case 'SysDateTimePick32'
			; if SysDateTimePick32 exists, obtain handle for SysDateTimePick32 and register WM_NOTIFY
			If $bEnableDarkTheme Then
				$g_hDate = $vCtrl
				$g_hDateProc_CB = DllCallbackRegister('__GUIDarkTheme_DateProc', 'ptr', 'hwnd;uint;wparam;lparam')
				$g_pDateProc_CB = DllCallbackGetPtr($g_hDateProc_CB)
				$g_hDateOldProc = __WinAPI_SetWindowLong($g_hDate, $GWL_WNDPROC, $g_pDateProc_CB)
				GUIRegisterMsg($WM_NOTIFY, "__GUIDarkTheme_WM_NOTIFY")
			EndIf

		Case 'msctls_progress32'
			If $bEnableDarkTheme Then
				If $bPreferNewTheme And $g_b24H2Plus Then
					$sThemeName = 'DarkMode_CopyEngine'
				Else
					$sThemeName = 'DarkMode'
				EndIf
				$sThemeList = 'Progress'
			EndIf

		Case 'Scrollbar'
			If $bEnableDarkTheme Then $sThemeName = 'DarkMode_Explorer'

		Case 'AutoIt v3 GUI'
			If $bEnableDarkTheme Then $sThemeName = 'DarkMode_Explorer'

		Case 'msctls_statusbar32'
			If $bEnableDarkTheme Then
				; get handle for statusbar
				$g_hStatus = $vCtrl
				Local Const $SP_GRIPPER = 3
				Local Const $TS_TRUE = 1
				Local $hTheme = __WinAPI_OpenThemeData($g_hGuiWnd, 'Status')
				Local $tSIZE = __WinAPI_GetThemePartSize($hTheme, $SP_GRIPPER, 0, Null, Null, $TS_TRUE)
				$g_hGripSize = $tSIZE.X
				__WinAPI_CloseThemeData($hTheme)
				If $bPreferNewTheme And $g_b24H2Plus Then
					$sThemeName = 'DarkMode_DarkTheme'
					$sThemeList = 'Status'
					$g_iGripPos = 0
					$g_iBkColor = 0x3b3b3b
				Else
					$sThemeName = 'DarkMode'
					$sThemeList = 'ExplorerStatusBar'
					$g_iGripPos = 1
					$g_iBkColor = 0x1c1c1c
				EndIf
				GUIRegisterMsg($WM_SIZE, "__GUIDarkTheme_WM_SIZE")
				__GUIDarkTheme_StatusRatio()
				__GUIDarkTheme_CreateSizebox()
			Else
				;
			EndIf

		Case 'tooltips_class32'
			If $bEnableDarkTheme Then
				If $bPreferNewTheme And $g_b24H2Plus Then
					;$sThemeName = 'DarkMode_DarkTheme' ; works but is faded (MS still developing DarkMode_DarkTheme parts)
					$sThemeName = 'DarkMode_Explorer' ; use for now
					$sThemeList = 'ToolTip'
				Else
					$sThemeName = 'DarkMode_Explorer'
					$sThemeList = 'ToolTip'
				EndIf
			Else
				;
			EndIf

		Case 'ComboLBox', 'ComboBox'
			If $bEnableDarkTheme Then
				If $bPreferNewTheme And $g_b24H2Plus Then
					$sThemeName = 'DarkMode_DarkTheme'
					$sThemeList = 'Combobox'
				Else
					$sThemeName = 'DarkMode_CFD'
					$sThemeList = 'Combobox'
				EndIf

				; create brush and register GUI message
				If Not $g_hBrushEdit Then $g_hBrushEdit = __WinAPI_CreateSolidBrush(__GUIDarkMenu_ColorToCOLORREF($COLOR_CONTROL_BG))
				GUIRegisterMsg($WM_CTLCOLORLISTBOX, "__GUIDarkTheme_WM_CTLCOLOR")
			Else
				;
			EndIf

			GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
			GUICtrlSetBkColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_BkColor)

			#cs
				Local $hEdit = __WinAPI_FindWindowEx($vCtrl, 0, "Edit", "")
				         If $hEdit Then
				             __WinAPI_SetWindowTheme($hEdit, "DarkMode_CFD", 0)
				              __WinAPI_AllowDarkModeForWindow($hEdit, True)
				         EndIf
				         ; ComboBox dropdown list
				         Local $hComboLBox = __WinAPI_FindWindowEx($vCtrl, 0, "ComboLBox", "")
				         If $hComboLBox Then
				             __WinAPI_SetWindowTheme($hComboLBox, "DarkMode_Explorer", 0)
				             __WinAPI_AllowDarkModeForWindow($hComboLBox, True)
				         EndIf
			#ce

		Case 'SysTabControl32'
			If $bEnableDarkTheme Then
				If $bPreferNewTheme And $g_b24H2Plus Then
					$sThemeName = 'DarkMode_DarkTheme'
					$sThemeList = 'Tab'
				Else
					$sThemeName = 'DarkMode_Explorer'
				EndIf
				; Register a custom window procedure for the tab control for owner-drawing
				$g_hTab_CB = DllCallbackRegister('__GUIDarkTheme_TabProc', 'ptr', 'hwnd;uint;wparam;lparam')
				$g_pTab_CB = DllCallbackGetPtr($g_hTab_CB)
				$g_hProc = __WinAPI_SetWindowLong($vCtrl, $GWL_WNDPROC, $g_pTab_CB)
				$g_hTab = $vCtrl
			Else
				$sThemeName = 'Explorer'
			EndIf

		Case Else
			$sThemeName = 'DarkMode_Explorer'

	EndSwitch
	;ConsoleWrite(@CRLF & 'Class:' & __WinAPI_GetClassName($vCtrl) & ' Theme:' & $sThemeName & '::' & $sThemeList)
	;=========
	__WinAPI_SetWindowTheme($vCtrl, $sThemeName, $sThemeList)
	If @error <> 0 Then Return SetError(3, @error, False)
	__SendMessage($vCtrl, $WM_THEMECHANGED, 0, 0)
	If @error <> 0 Then Return SetError(4, @error, False)
	If $bSpecialBtn Then
		; this is used to remove theme from group box and radio buttons which are not themed properly with older DarkMode_Explorer
		__WinAPI_SetWindowTheme($vCtrl, "", "")
		GUICtrlSetColor(_WinAPI_GetDlgCtrlID($vCtrl), $iGUI_Ctrl_Color)
	EndIf
	Local $iSize, $hChecked, $hUnchecked
	If $bSpecialLV Then
		Local $hImageListLV = _GUICtrlListView_GetImageList($vCtrl, 2)
		$iSize = _GUIImageList_GetIconHeight($hImageListLV)
		_GUIImageList_Remove($hImageListLV)
		If @OSBuild >= 22000 Then
			$hUnchecked = __GUIDarkTheme_GetImages(0, 3, $iSize, $iSize)
			$hChecked = __GUIDarkTheme_GetImages(5, 3, $iSize, $iSize)
			_GUIImageList_Add($hImageListLV, $hUnchecked)
			_GUIImageList_Add($hImageListLV, $hChecked)
			__WinAPI_DeleteObject($hChecked)
			__WinAPI_DeleteObject($hUnchecked)
		Else
			;_GDIPlus_Startup()
			$hChecked = _GDIPlus_BitmapCreateFromMemory(__GUIDarkTheme_CheckedPNG($iSize), True)
			$hUnchecked = _GDIPlus_BitmapCreateFromMemory(__GUIDarkTheme_UncheckedPNG($iSize), True)
			_GUIImageList_Add($hImageListLV, $hUnchecked)
			_GUIImageList_Add($hImageListLV, $hChecked)
			;_GDIPlus_Shutdown()
		EndIf
	EndIf
	If $bSpecialTV Then
		Local $hImageListTV = _GUICtrlTreeView_GetStateImageList($vCtrl)
		$iSize = _GUIImageList_GetIconHeight($hImageListTV)
		_GUIImageList_Remove($hImageListTV)
		If @OSBuild >= 22000 Then
			$hUnchecked = __GUIDarkTheme_GetImages(0, 3, $iSize, $iSize)
			$hChecked = __GUIDarkTheme_GetImages(5, 3, $iSize, $iSize)
			_GUIImageList_Add($hImageListTV, $hUnchecked)
			_GUIImageList_Add($hImageListTV, $hUnchecked)
			_GUIImageList_Add($hImageListTV, $hChecked)
			__WinAPI_DeleteObject($hChecked)
			__WinAPI_DeleteObject($hUnchecked)
		Else
			;_GDIPlus_Startup()
			$hChecked = _GDIPlus_BitmapCreateFromMemory(__GUIDarkTheme_CheckedPNG($iSize), True)
			$hUnchecked = _GDIPlus_BitmapCreateFromMemory(__GUIDarkTheme_UncheckedPNG($iSize), True)
			_GUIImageList_Add($hImageListTV, $hUnchecked)
			_GUIImageList_Add($hImageListTV, $hUnchecked)
			_GUIImageList_Add($hImageListTV, $hChecked)
			;_GDIPlus_Shutdown()
		EndIf
	EndIf
	Return True
EndFunc   ;==>_GUIDarkTheme_GUICtrlSetDarkTheme

Func _GUIDarkTheme_ApplyDark($hWnd, $bPreferNewTheme = False)
	_GUIDarkTheme_GUISetDarkTheme($hWnd, True)
	_GUIDarkTheme_GUICtrlAllSetDarkTheme($hWnd, True, $bPreferNewTheme)

	; GUIDarkMenu register
	_GUIDarkMenu_Register($hWnd)
	; redraw menus
	_GUICtrlMenu_DrawMenuBar($hWnd)
EndFunc   ;==>_GUIDarkTheme_ApplyDark

Func _GUIDarkTheme_ApplyLight($g_hGuiWnd)
	_GUIDarkTheme_GUISetDarkTheme($g_hGuiWnd, False)
	_GUIDarkTheme_GUICtrlAllSetDarkTheme($g_hGuiWnd, False)
EndFunc   ;==>_GUIDarkTheme_ApplyLight

Func __GUIDarkTheme_CheckedPNG($iSize = 16, $sSavePath = Default)
	Local $Base64String, $iExt = (IsKeyword($sSavePath) ? 0 : 2) ; "BitAND(@extended, 2)" meant that the user wanted to save to disk
	Switch $iSize
		Case 13
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAAA0AAAANCAYAAABy6+R8AAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcYOepp4t4AAADpSURBVCjPY2RgYGAwSG9o0IrMz2fjFRBgwAF+ff7w4c7mBQtO9RYWMpuVTJigm1BezszOwcGABzCzc3CI6lpYMDAwMDBGHXj/Hp8N2GxkIkUDAwMDAxuvgAATMQrvbF6Awieo6Uh9AgMDIyNuTY8ObMDQIGHqyKDiE49b06/PHxiONCTi1cDAwMDAgsxR8U1gYGBgYFjjq8hgkN6AVQMDAwMDY8LZ//+xBCsDGy/uQGX69fnDByzBijeemJnZODgkTBwciI2nm2tnzGB+cfbgQTZeAQF+BQ0NfEnp1+cPHy4v6Og4O6WyEgBYG0pIgbNupgAAAABJRU5ErkJggg=='
		Case 20
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcZJn563moAAAEdSURBVDjLY2RgYFDilVZUs66f3yqkpq/GxivAw0AieHRg46FTvQXtX58/vMXII6Vg5bfs/A42XgFeBgrAr88fPm2KNPRktmlYMFNI3UCbgULAzM7BLqRuoMGYcPb/fwYqgV+fP3xkYqAiYOMV4KeqgQwMDAyDw8B3ty4wbIo0xCrHQo5h+4oCGZz61lPuQmTDhNQMiDdwU6Qhw7tbF0g2DKeBNo3zGfYVB8INJdYwBgYGBpwJ+92tCwz7igMZDNLqGS7MbCTKMLwGkuoyogwcHAn71+cPH6hq4LtbFy5Qy7BHBzZsYOSRUlDwW3b+PBuvgACFRdeHTZGGhsy/Pn/4cH/XypU8UgoK/AoaGuQY9PryiRO7cz09vzx/8AAAGLWDCCvF1IEAAAAASUVORK5CYII='
		Case 26
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAABoAAAAaCAYAAACpSkzOAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcaCxCI0dwAAAF1SURBVEjH7da9a8JAGAbwJ1pSh1iyFGqmG5wKHaQgDh10c7FTF7e4uCr4D+heyKpOuhQculgH6WICTgU1q0Uwk41Dy0EChVvs0JamBa3mQyj0mY/8eLjcy8vhI4JEyGm+VIrnZJmPiiI8xF4YhjlSVb1Rq9lPhgEAHADEc7KcrCiKV+BnmEXpw3W5POu125wgEXJ5M5n4jTixbj6RCCcrinJ8lkohoIQPIxEA4K5687kQIwQBhlmUcvJotcIeEsKe8g+tu2z0i2kMq4W1Zw78QTJgNsVFrRVMIyeSbQ4gxIj/0C7IRkhvVNHNJ8As6hnZCMVz8tcHHZgbZCMkSATZ5uAb5hYBgF9HkL0w0C9mwEffh7sbZKufwdnMLbJVI2czcHCF7PRgBYn8kVnHLEqDRuyFYYRmd61W0JA5UtX9LSfMovT1ebk8OU+nPxcJv9ctc6xpYQB4edT1+X2nw0dFkT8SRa/tmEXp9LZeH1YLBXOsaQDwBivy2deCjE70AAAAAElFTkSuQmCC'
		Case 32
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxccGixiVqgAAAGaSURBVFjD7dcxaMJQFIXhk6Skdgh9tAiZSqcigmARxKFD7OTYVad2EbcguDdduolxjEu3bNLSyUy+oSAIgUyBdJKCNJMG7ORipxSxFKOGBIp3CiHh/wgXwmOwNGJOktIVWRZzksQLhCDEmc88zzUptfV22zUp9e8z/kW+oarpsiwjgrF1VR0263UA4KKOA0AyUyjwAiHjgWEwYk6SSp1+HzFMr1oscvlGq3V8nkrFAeAFQpgKnU7DXrhNFpONK+5/ARYxzx7wfwG2rmLYrMcDsDQFw2YdE8eKHmBpCqzOA04usrhuPkcLWI6XOn3wAokOsE08EOC1fLl2mbaNBwKcFW9g6yrelLvQ4wBwsO6BbPX+JwQAV8pTaPFAgL8QYcQDA1YRE8fC5N3aOb4RYBURRhwAmFtzsdj0pQ/6AjEn7RzfGrD/G4YKmM88L674fOZ57PIxKepxTUpZW2+34wLYuqpyX5+jES8QkswUClHHna6mcQAwHhhGlIhfh1Mf4ZqU8gIhR6eiyB0mEmEv3HjQ6w0eazWnq2n+/W+c6MNyCJEgYgAAAABJRU5ErkJggg=='
		Case 40
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxccOPkCF0wAAAHRSURBVFjD7dg9a8JAGAfwRwPBgkKwFNRB4mKhUAgIxTHZHFrM0KVbXdKtDv0Ams3NrPoBBMEh0kVdVCh0Em4SdPEIVDtJwAwiFDtZ0mCLUaNXmv8UuOF+PPeS4/GAJaEEz0f5dDoqiKI/zLJwgEyHCE0HCKGSLBsTjM1jntUHHWAY7iGfv7jLZuGI6VcUBZVleTHT9S8gHWCYVLndDsY5DgjIdIhQQxKExUzXKQCAxGOhEOVFEQjJyWkoRNE+39trs+nxR1j29nk0AgLTkASBunoqFoPnZCytNYuZrntJxQEARHlR9Nz3lksgOF4gPC7QBbpAkoDGGEPtOga1mxh5QGOMoSEJYEwweRU04/wRFlLlNjnAdTh/mCUDuCvOUeA+cLaAq9O3ySbfF852BTc5ifvE2QKuJvoNuW+cLaB5wnVIJ3C2l/gnpFM4AICtXtRWECzBEdzW14y1kk7hdroHzSCncFsvsfsedIH/CmhtGJKU6QAhr9ZWVWKBQ4S8WqdeJxWISrJMGROM6QDDnF0mkyTh+hVFGbWq1b/RAv5YzOejVrVK0T7fsSvZryjKSz6T+dZEt/xjWU7K5YJxjjtUc9MYY6x1VFXrqOp7r9s1j30CYTArpz88OJ4AAAAASUVORK5CYII='
		Case 52
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAADQAAAA0CAYAAADFeBvrAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcdEaKrvmEAAAKqSURBVGje7dq/b9pAFAfwh6kIg1FQJseDdQOKGIkcdWI4MjUDwWsnPFGWSo3yB4QsmdJmbhbK0ipTnKZD1ErNZYoUQTErQuqFARypRZawFIklXeqKumkbwPiX/N18guGjOz/7WS8CfwknYizgQoFbw5hdRiiWSCbBxYyGum70KdXqhHROa7VBW1Xv+13EusDyCGUr1SonYgweTpcoytX+1pbRp3R8PTp+kcrL8vrL4+NFlE6Dx7OI0ulUXpZvv2naoN1q/QFK5WU5W6lWowvxOPgk0YV4XMhJktGj1ERFzGO2+bbZdPs+meX+ev90ddXoU8oAADzePjjwKwYAIJZIJrO71SoAQGRpJZPZfNdsQgByVsrlmFS+WISARMCFAsOtebs8TxJOxJhhlxEKCojlEWL8XAzuKw4MBCwhKASFoBAUghzNoK3C520JukSZ6v+PvIY5K+VgNNRh0G6BgCX/7tA45mfD6d8jZ8VknlUgU9rxJ8hOjOsguzGuguaBcQ00L8zUoKv9F/BGjID6uuIpzNQg86GnHu5OhJo3ZmrQ+isFYonkRCgnMFODllYy8OTw/MEopzAzFYWHopzEzFzl/odyGgMAEJEbd3e2l+HSDgg5yXGMbaD7ULFE0nGMrQ9W6/FzA2P7m4IV5TTG1iM3HqNHwehT4ETs+GvVXDpWlkfA8ij8phCCQpAfQaOhrgcFMxrqOmOdxPBzjB6ljFYnJCggrUEI0yUnJ0EBdU5rNUZrEKI1/L9L3XNFGbRVNXCjMVFz4fb7zY2AJcmPoMu9cln7cnEBMDaN'
			$Base64String &= 'NWirqtG/vuZEjP0ykTUa6vrlXrnc+VCrmWtRS5Omfv14dMTyCHl9Zk6rE/Lp+caGuTO/2od/9DaZVL5Y5ESMWd4jI5o9SrUGIV2iKFrjd4iZH07DnAm7lTVAAAAAAElFTkSuQmCC'
		Case Else ; 16
			If $iSize <> 16 Then $iExt = BitOR($iExt, 1)
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcZFLatj+oAAAEPSURBVDjLY2RgYGDQiioo0IrKz+eRVFBgIAK8u3XhwrWlEybc2bJwIaOKb0KCTcP8+QxkgFO9hYWMfsvPnxdSMzAgx4Avzx48YEw4+/8/AwWAiYFCQJIBdzYvYPj1+QN5BlyY2cBwbdlEBjZeAdINuDCzgeHRgY0MHrP24/fCl2cPGNb4KjJ8ef4Aq2Z02zEM4JFSYDBIq2fYkebI8OX5A4KaGRgYGFjQBVR8ExgYGBgYNkUaMvBIKuDVzMDAwIAzHby7eYGBR0oBr2asLoABIXXiEifTl+cPHpCbiN7dvHCBmYGBkVHaysODHAPOTq6oYH59+cSJL88fPuSRVFDgFJGQIDYTXZjZ0HBz3cyZABhkZ7hwets3AAAAAElFTkSuQmCC'
	EndSwitch
	Local $bString = __WinAPI_Base64Decode($Base64String)
	If @error Then Return SetError(1, $iExt, 0)
	$bString = Binary($bString)
	If Not $iExt Then
		Local Const $hFile = FileOpen($sSavePath & "\checked.png", $FO_BINARY + $FO_OVERWRITE)
		If @error Then Return SetError(2, $iExt, $bString)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf
	Return SetError(0, $iExt, $bString)
EndFunc   ;==>__GUIDarkTheme_CheckedPNG

Func __GUIDarkTheme_UncheckedPNG($iSize = 16, $sSavePath = Default)
	Local $Base64String, $iExt = (IsKeyword($sSavePath) ? 0 : 2) ; "BitAND(@extended, 2)" meant that the user wanted to save to disk
	Switch $iSize
		Case 13
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAAA0AAAANCAYAAABy6+R8AAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcfHO4soF4AAAC3SURBVCjP7dIhDoMwGMXx1y0kX1pDFUdBsgRDgsb3BmThAuwEQyAmwWOruQQWjdxn2lROTbKR6f31+7knAKCqqrYsy1pKGWMn7z3P8zyM43gVxpguSZLaWosQwp6B1hp5nmNd11YMw/Ds+z5mZnyLiNA0DZ+klIcAAIQQoJSKT/ihP3oj7z0T0aGx1hrOOT5HUURpml62bfv4CCJCURRYluUhAMAYc8+yzCildr/nnGNrbTdN0+0FDuRGAoicao0AAAAASUVORK5CYII='
		Case 20
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcgFyEgU+oAAAD6SURBVDjL7dUhboRAFMbxB2smYAjJJMi5AVuSCYaEEEwlvQHFoirgEpDUYcsNukdoBWNIaBUWLBkzBjJiRNW6ihIq+R/gl++ppwEAYIxJnudvhJCrYRgW7Kzv+1vbti+c81nDGJOqqr4YYxZjDKSUuzCEEHieB1EUibIsH7SiKN6XZUkYY3CkOI7BcZwPnVKaDMMAR+u6DgghVx0Adp/5W1JKME3T0uGfO8ETPME/gdu2CYTQYehuXFzXfbRtm0zTdAj0fR+EELfLOI6fWZalSinEOQel1O5lYRhCEASirusn7f4C0jR9pZQme5et6yrmef5umuaZcz7/AC65XdIkaSqyAAAAAElFTkSuQmCC'
		Case 26
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAABoAAAAaCAYAAACpSkzOAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcgMfMt1hcAAAEzSURBVEjH7dahbsJAAIfxf28VCyTNnbgqRPWh5uo4LB5FxfEEm8Iu6AqCR3RJXTXo4wmmeqYGTUXPlGAaZkaCL12ypN8L/Ozn4DfOeTCbzd6n06kaDAYULSrL8pTnuc6ybF2W5QkAHACQUiql1IYQQvM8h7W2jQPGGIQQaJrGJknycTwevxzOeRDH8be1lqZp2hp5xKIogud5drVavb0opTaj0Sjc7XZPQwDger2iKAqEYfjqui7IeDyWxpinIveqqoIxBlJKRTjnQRfIIzYcDinBH9VDPdRDPdRD/xK6XC6WUtoZwBjD+Xw+Ea11IoRAF9h9Uowxmuz3++3tdrNRFD0VY4xhsVigaRqbZdnaAYDJZKKWy2W3u/U4kPP5/FMIIX3fD9pAdV1brXVyOBy294H8AUHOkTVbUaKIAAAAAElFTkSuQmCC'
		Case 32
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxciEmN8xecAAAF/SURBVFjD7ZexiqtAFIb/mFsYheA8gt0y8wDZLpIpAuNL2KTe7JOsWy23tMgjjEWQ4EJe4UxpLqRMoY1ikeJWguUWidvMVw2HA/8Hc5p/hhFCiEgp9SaEiDzPC/BAuq5riKjM8/yTiMphPhseSZKkSqm3vu9BRGia5pH5YIwhDEMwxqC1TrMseweAP+Pw8/mM0+mEvu/xLDabDeI43gNAlmXvcyFEtNvtvoqiwPF4xP1+xzO5XC5wXRdSylci+p4nSfLhed7L4XDAVFyvV6xWKyyXy8ARQkRVVWFK+r6HMQZCiMjxPC949MH9hLqu4ft+4OCXsQJWwApYAStgBayAFbACTtd1TRAEkwczxtC2beMQUck5h+u6kwqEYQgiKh2t9edisYCUcrJwKSUYY8jzPJ3fbrd/vu8HUsrXobk8C9d1sd1usV6vobVOi6L4Oy6nH0qpfV3XqKrqKeV0+OpxOZ2Nlzjn6ziO95zzyPf9h15m27aNMabUWqfGmO9h/h/YFp4sUeI9ggAAAABJRU5ErkJggg=='
		Case 40
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxciKdJ3LMMAAAIZSURBVFjD7dg9q9pQAMbx5/oC1SDEIXAzeZIxLiG+cDFLuxWXOqiLiwX3+g2u2br1gqAfoZQsdnDoVC5IoAbfBuMgmDgJOphB5YBLJ0UunVXo+U+Bs/zI4SzPA96UTCbfp9PpT9lstiAIAsEV8jxv7Hne2DRNY7PZeJdnD6ePaDTKl8vlRj6f/4Ib1u12X0zTNA6Hg38GRqNRvtFo/CaEqJRS9Ho9uK4L13WvghJFEbquQ9M0AIDrumPDMD4cDgc/CACVSuVrJpMpbLdbtNttzGYz+L5/tb+22+3gOA6GwyEURYEoio/hcPjdZDL5FRQEgdTr9e8A0Gw2rwp7G6UUs9kMqVQKiqI8TafT12C1Wv1GCFEHgwFGoxFuHaUUoVAIsixjv9/7AUmSVACwLAv30mKxAABks9lCIJFIqACwWq3uBnh6nIIgkADuPAZkQAZkQAZkQAZkQAZkQAZkQAZkwP8aeBoMJUm6G5QoiqeFYRywbbsDALIs3w1Q13UAwHK5HAf6/f5PAMjlcuB5/ua4eDx+HjJN0zSCm83G4ziOVxTlSVEUOI4DSunNcLVaDZFIBN1u98WyrB9BAJjP539UVf0oiuKjruvgeR6+72O3210FJssyNE1DsVhELBaD67rjVqv1+Xg80vOIznEcXyqVnvP5fP2WV/zPEf0yQRBIqVR6JoSohBD1Gqj1eu3Ztt2xbbvjOM7r5dlfsTLTT1G4ptgAAAAASUVORK5CYII='
		Case 52
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAADQAAAA0CAYAAADFeBvrAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcjAxDX1FQAAAKhSURBVGje7dqxbtpQGAXgg5tICBYzGMmejEd7SRPIksFXYmAMDwCK2apkaRZeIkOmsDAkFi+QPgDIlpLFCDWLPWKzxKo9xIuRxw7RRYQmkSpVAlf3bFw83E/31z+dAj6IpmmkXq+fappGqtWqXCqVeGwxy+UyiaIocF3Xsm3bDILg6b3vCpsHgiDIFxcXt6qqEuxwHMe5N03zMo7j4EMQIcQwDON6/TXm8znCMESWZVsFFItFiKIIRVFWZ2maJnd3d99t2zb/ABFCjPPz81v6++HhAZPJZOuQzVQqFTSbTRweHq7Obm5uDIoq0DG7urr6WSqV+CzLMBqN4Pv+Lk8carUaut0uisUi0jRN+v3+1ziOAw4A1scsDxgA8H0fo9EIAFAul3k6XZwsyweNRqMNALPZLBeYddTj4+NqK6uqSjhCyBn9YDweI29Zv3Oj0TjlNE0jdJslSZI7UJZlmM/nAABVVQknCIIMAGEYIq+hd69WqzJHl8Guree/fSW6HDj8Z2EgBmIgBmIgBmIgBmIgBmIgBmIgBmIgBmIgBmIgBmIgBmIgBvrHoOVymQCvPYC8ht49TdOEo00MURRzC5IkCQAQRVHAua5rAYCiKLl8pUqlglqtBgDwPM/iHMf5Qf9sNpu5A63f2bIsk/M8z6KvdHJystLmIYqirCoyjuPcLxaLJw4ABoNBjy6HTqeTC5SiKOh0OqtlYJrmJQB8AV67aC8vL7+Oj4/b+/v7ODo6As/zO9HCem+jtVottNtt7O3tAQCGw+E3z/NsYKNepuu60ev13tTLfN/H'
			$Base64String &= '8/PzTtTLJEl6Mz2f1stoBEGQDcO4pv2fXY3rutZgMOh9WgBcjyzLB7qun2maRgRBkMvl8lYrmmmaJnEcB67rWtPp9J6O2GZ+A6oXHMXfWhlTAAAAAElFTkSuQmCC'
		Case Else ; 16
			If $iSize <> 16 Then $iExt = BitOR($iExt, 1)
			$Base64String &= 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IB2cksfwAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAAd0SU1FB+oDAxcfO0smFTUAAADBSURBVDjL7cyhDYNAGIbhD1qFqzhxijsmYAOQDZcQgkIg2IDK7oBou0HrCYIEsGxQNvjPkYBlgK5wpJbXP68FAEqpWxRFJWNMwCCt9dT3/XMcx885DMMijuNHXdcgIhMPzrmf5/nbcZyLVVXVt21bf55n7ElKiSRJtO267m4MAEQExpiw8WfH4BgAgL2uq5ZS7oaccxDRdAJgpWl6JSJs22aEPc9DlmVomuZuAUAQBIVSqhRC+CaDZVl013XPYRheP/vGQ/mk4/+CAAAAAElFTkSuQmCC'
	EndSwitch
	Local $bString = __WinAPI_Base64Decode($Base64String)
	If @error Then Return SetError(1, $iExt, 0)
	$bString = Binary($bString)
	If Not $iExt Then
		Local Const $hFile = FileOpen($sSavePath & "\unchecked.png", $FO_BINARY + $FO_OVERWRITE)
		If @error Then Return SetError(2, $iExt, $bString)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf
	Return SetError(0, $iExt, $bString)
EndFunc   ;==>__GUIDarkTheme_UncheckedPNG

Func __GUIDarkTheme_SizeboxProc($hWnd, $iMsg, $wParam, $lParam) ; Andreik

	If $iMsg = $WM_PAINT Then
		Local $tPAINTSTRUCT
		Local $hDC = _WinAPI_BeginPaint($hWnd, $tPAINTSTRUCT)
		Local $hGraphics = _GDIPlus_GraphicsCreateFromHDC($hDC)
		_GDIPlus_GraphicsDrawImageRect($hGraphics, $g_hDots, 0, 0, $g_hGripSize + 2, $g_hGripSize + 2)
		_GDIPlus_GraphicsDispose($hGraphics)
		_WinAPI_EndPaint($hWnd, $tPAINTSTRUCT)
		Return 0
	EndIf
	Return __WinAPI_CallWindowProc($g_hOldProc, $hWnd, $iMsg, $wParam, $lParam)
EndFunc   ;==>__GUIDarkTheme_SizeboxProc

Func __GUIDarkTheme_CreateDots($iWidth, $iHeight, $iBackgroundColor)
	Local $hTheme = __WinAPI_OpenThemeData($g_hGuiWnd, 'Status')
	Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iWidth, $iHeight)
	Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
	_GDIPlus_GraphicsClear($hGraphics, $iBackgroundColor)
	; draw SP_GRIPPER directly on graphics object
	Local $hDC = _GDIPlus_GraphicsGetDC($hGraphics)
	Local $tRect = _WinAPI_CreateRectEx($g_hGripSize + 2 - $g_hGripSize + $g_iGripPos, $g_hGripSize + 2 - $g_hGripSize + $g_iGripPos, $g_hGripSize, $g_hGripSize)
	__WinAPI_DrawThemeBackground($hTheme, 3, 0, $hDC, $tRect)
	_GDIPlus_GraphicsReleaseDC($hGraphics, $hDC)
	_GDIPlus_GraphicsDispose($hGraphics)
	__WinAPI_CloseThemeData($hTheme)
	Return $hBitmap
EndFunc   ;==>__GUIDarkTheme_CreateDots

Func __GUIDarkTheme_WM_SIZE($hWnd, $iMsg, $wParam, $lParam) ; Pixelsearch
	#forceref $iMsg, $wParam, $lParam

	If $hWnd = $g_hGuiWnd Then
		Local Static $bIsSizeBoxShown = True
		Local $aSize = WinGetClientSize($g_hGuiWnd)
		Local $aGetParts = _GUICtrlStatusBar_GetParts($g_hStatus)
		Local $aParts[$aGetParts[0]]
		For $i = 0 To $aGetParts[0] - 1
			$aParts[$i] = Int($aSize[0] * $g_aRatioW[$i])
		Next
		If BitAND(WinGetState($g_hGuiWnd), $WIN_STATE_MAXIMIZED) Then
			_GUICtrlStatusBar_SetParts($g_hStatus, $aParts)
			_WinAPI_ShowWindow($g_hSizebox, @SW_HIDE)
			$bIsSizeBoxShown = False
		Else
			If $g_aRatioW[UBound($aParts) - 1] <> 1 Then $aParts[UBound($aParts) - 1] = $aSize[0] - $g_iHeight
			_GUICtrlStatusBar_SetParts($g_hStatus, $aParts)
			WinMove($g_hSizebox, "", $aSize[0] - $g_hGripSize - 2 - $g_iGripPos, $aSize[1] - $g_hGripSize - 2 - $g_iGripPos, $g_hGripSize + 2, $g_hGripSize + 2)
			If Not $bIsSizeBoxShown Then
				_WinAPI_ShowWindow($g_hSizebox, @SW_SHOW)
				$bIsSizeBoxShown = True
			EndIf
			_WinAPI_RedrawWindow($g_hSizebox)
		EndIf
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>__GUIDarkTheme_WM_SIZE

Func __GUIDarkTheme_StatusRatio()
	; calculate ratio need for the resizing of statusbar parts
	Local $iGuiWidth = WinGetClientSize($g_hGuiWnd)[0]
	Local $aParts = _GUICtrlStatusBar_GetParts($g_hStatus)
	_ArrayDelete($aParts, 0)
	Dim $g_aRatioW[UBound($aParts)]
	For $i = 0 To UBound($aParts) - 1
		$g_aRatioW[$i] = $aParts[$i] / $iGuiWidth
	Next
EndFunc   ;==>__GUIDarkTheme_StatusRatio

Func __GUIDarkTheme_CreateSizebox()
	; create grip
	$g_iHeight = WinGetPos($g_hStatus)[3] - 3
	$g_hDots = __GUIDarkTheme_CreateDots($g_hGripSize + 2, $g_hGripSize + 2, 0xFF000000 + $g_iBkColor)

	Local Const $SBS_SIZEBOX = 0x08
	; Create a sizebox window (Scrollbar class)
	$g_hSizebox = _WinAPI_CreateWindowEx(0, "Scrollbar", "", $WS_CHILD + $WS_VISIBLE + $SBS_SIZEBOX, _
			0, 0, 0, 0, $g_hGuiWnd) ; $SBS_SIZEBOX or $SBS_SIZEGRIP

	; Subclass the sizebox (by changing the window procedure associated with the Scrollbar class)
	$g_hSizeboxProc = DllCallbackRegister('__GUIDarkTheme_SizeboxProc', 'lresult', 'hwnd;uint;wparam;lparam')
	$g_hOldProc = __WinAPI_SetWindowLong($g_hSizebox, $GWL_WNDPROC, DllCallbackGetPtr($g_hSizeboxProc))

	$g_hCursor = _WinAPI_LoadCursor(0, $OCR_SIZENWSE)
	__WinAPI_SetClassLongEx($g_hSizebox, $GCL_HCURSOR, $g_hCursor)

	; Fix Z-order of Sizebox (needed for cursor)
	__WinAPI_SetWindowPos($g_hSizebox, $HWND_TOP, 0, 0, 0, 0, BitOR($SWP_NOMOVE, $SWP_NOREDRAW, $SWP_NOSIZE))

	; Add WS_CLIPSIBLINGS to statusbar (needed for cursor)
	__WinAPI_SetWindowLong($g_hStatus, $GWL_STYLE, BitOR(__WinAPI_GetWindowLong($g_hStatus, $GWL_STYLE), $WS_CLIPSIBLINGS))
EndFunc   ;==>__GUIDarkTheme_CreateSizebox

Func __GUIDarkTheme_GetImages($iState, $iPart = 3, $iWidth = 16, $iHeight = 16)
	Local $hTheme = __WinAPI_OpenThemeData($g_hGuiWnd, "Button")
	If @error Then Return SetError(1, 0, 0)
	Local $hDC = __WinAPI_GetDC($g_hGuiWnd)
	Local $hHBitmap = __WinAPI_CreateCompatibleBitmap($hDC, $iWidth, $iHeight)
	Local $hMemDC = __WinAPI_CreateCompatibleDC($hDC)
	Local $hObjOld = __WinAPI_SelectObject($hMemDC, $hHBitmap)
	Local $tRect = __WinAPI_CreateRectEx(0, 0, $iWidth, $iHeight)
	__WinAPI_DrawThemeBackground($hTheme, $iPart, $iState, $hMemDC, $tRect)
	If @error Then
		__WinAPI_CloseThemeData($hTheme)
		Return SetError(2, 0, 0)
	EndIf
	__WinAPI_SelectObject($hMemDC, $hObjOld)
	__WinAPI_ReleaseDC($g_hGuiWnd, $hDC)
	__WinAPI_DeleteDC($hMemDC)
	__WinAPI_CloseThemeData($hTheme)
	Return $hHBitmap
EndFunc   ;==>__GUIDarkTheme_GetImages
