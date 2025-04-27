#include <Constants.au3>
#include <GUIListView.au3>

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUICtrlListView_SaveCSV
; Description ...: Exports a listview to a CSV file.
; Syntax ........: _GUICtrlListView_SaveCSV($hListView, $sFilePath[, $sDelimiter = '[, $sQuote = '"']])
; Parameters ....: $hListView           - Control ID/Handle to the control
;                  $sFilePath           - Filepath to save the CSV data string to.
;                  $sDelimiter          - [optional] CSV delimiter. Default is ,.
;                  $sQuote              - [optional] CSV quote type. Default is ".
; Return values .: Success - True
;                  Failure - False and sets @error to non-zero.
; Author ........: guinness
; Remarks .......: Thanks to Prog@ndy for the idea of CSV creation. GUICtrlListView.au3 should be included.
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlListView_SaveCSV($hListView, $sFilePath, $sDelimiter = ',', $sQuote = '"')
	If $sDelimiter = Default Then
		$sDelimiter = ','
	EndIf
	If $sQuote = Default Then
		$sQuote = '"'
	EndIf

	Local Const $iColumnCount = _GUICtrlListView_GetColumnCount($hListView) - 1
	Local Const $iItemCount = _GUICtrlListView_GetItemCount($hListView) - 1
	Local $sReturn = ''
	For $i = 0 To $iItemCount
		For $j = 0 To $iColumnCount
			$sReturn &= $sQuote & StringReplace(_GUICtrlListView_GetItemText($hListView, $i, $j), $sQuote, $sQuote & $sQuote, 0, 1) & $sQuote
			If $j < $iColumnCount Then
				$sReturn &= $sDelimiter
			EndIf
		Next
		$sReturn &= @CRLF
	Next

	Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
	If $hFileOpen = -1 Then
		Return SetError(1, 0, False)
	EndIf
	FileWrite($hFileOpen, $sReturn)
	FileClose($hFileOpen)
	Return True
EndFunc   ;==>_GUICtrlListView_SaveCSV