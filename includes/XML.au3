#include-once
;~ #include <XMLConstants.au3>
#include <FileConstants.au3>
#include <StringConstants.au3>
#include <AutoItConstants.au3>
#include "ADO_CONSTANTS.au3"

#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w 7
#Tidy_Parameters=/sort_funcs /reel

#Region XML.au3 - UDF Header
; #INDEX# =======================================================================================================================
; Title .........: XML.au3
; AutoIt Version : 3.3.10.2++
; Language ......: English
; Description ...: Functions to use for reading and writing XML using msxml. (UDF based on XMLWrapper.au3)
; Remarks .......: BETA Version
; Author(s) .....: mLipok, Eltorro, Weaponx, drlava, Lukasz Suleja, oblique, Mike Rerick, Tom Hohmann, guinness, GMK
; Version .......: "1.1.1.13" ; _XML_MiscProperty_UDFVersion()

#CS
	This UDF is created on the basis of:
	https://www.autoitscript.com/forum/topic/19848-xml-dom-wrapper-com/
	For this reason, I attach also the last known (to me) previous version ($_XMLUDFVER = "1.0.3.98"  _XMLDomWrapper_1.0.3.98_CN.au3 )
	For the same reason I continue to recognize the achievements of the work of my predecessors (they are still noted in each Function header).
	.
	.
	.
	. !!!!!!!!! This is BETA VERSION (all could be changed) !!!!!!!!!
	.
	.
	.
	WORK IN PROGRESS INFORMATION:
	. For now 2015-09-01 the descripion (Function Header) can not entirely correctly describe the function.
	. TODO: in many places I used "TODO" as a keyword to find what should be done in future
	. TO REMOVE _XML_NodeExists(ByRef $oXmlDoc, $sXPath)as is duplicate for _XML_SelectSingleNode($oXmlDoc, $sXPath)
	.
	I want to: PREVENT THIS:
	. The unfortunate nature of both the scripts is that the func return results are strings or arrays instead of objects.
	.
	I want to: USE THIS CONCEPT:
	.   All function should use Refernce to the object as first Function parameter
	.   All function should return in most cases objects. There should be separate functions to Change Object collection to array
	.   All function should use COM Error Handler in local scope.
	.   All function should return @error which are defined in Region XMLWrapperEx.au3 - ERROR Enums
	.	All function should have the same naming convention
	.	All variables should have the same naming convention
	.	There should not to be any Global Variable - exception is $__g_oXMLDOM_EventsHandler
	.   It should be possible easy to use XML DOM Events
	.		https://msdn.microsoft.com/en-us/library/ms764697(v=vs.85).aspx
	.   It should be possible easy to Debug
	.	Ultimately, you should be able to do anything with your XML without having to use your own Error Handler.

	. MAIN MOTTO:  "Programs are meant to be read by humans, and only incidentally for computers to execute." - Donald Knuth

#CE

#CS
	; !!!!!!!!!!!!!!!!  the previous version of UDF
	; Author ........: Stephen Podhajecki - Eltorro <gehossafats@netmdc.com>
	; Author ........: Weaponx
	; Modified ......: drlava, Lukasz Suleja, oblique, Mike Rerick, Tom Hohmann
	; Initial release Dec. 15, 2005

	;; There is also _MSXML.au3 at http://code.google.com/p/my-autoit/source/browse/#svn/trunk/Scripts/MSXML
	;; http://my-autoit.googlecode.com/files/_MSXML.au3
	;; which return the xml object instead of using a global var.
	;;
	;; Therefore using the one above, more than one xml file can be opened at a time.
	;; The unfortunate nature of both the scripts is that the func return results are strings or arrays instead of objects.

	!!!!!!!!!!!! The following description is intended only to show the extent of the changes that have taken place in relation to the source version of the UDF
	!!!!!!!!!!!! PLEASE TREAT THIS AS QUITE NEW UDF
	2015/09/02
	"1.1.1.01"

	. Global Variable renaming: $fXMLAUTOSAVE > $g__bXMLAUTOSAVE - mLipok
	. Global Variable renaming: $fADDFORMATTING > $g__bADDFORMATTING - mLipok
	. Global Variable renaming: $DOMVERSION > $g__iDOMVERSION - mLipok
	. Global Variable renaming: $_XMLUDFVER > $g__sXML_UDF_VERSION - mLipok

	. Local Variable renaming: $fDebug > $bDebug - mLipok
	. Local Variable renaming: $strNameSpc > $sNameSpace - mLipok
	. Local Variable renaming: $strXPath > $sXPath - mLipok

	. added EventHandler for __XML_DOM_EVENT_onreadystatechange - mLipok
	. added EventHandler for __XML_DOM_EVENT_ondataavailable - mLipok
	. added EventHandler for __XML_DOM_EVENT_ontransformnode - mLipok

	. _XML_Misc* functions renamed and extended - mLipok
	. #Region -- grouping functions - mLipok
	. New Enums - Stream Object - ReadText Method - StreamReadEnum - mLipok
	. New Enums - readyState Property (DOMDocument) - mLipok
	. New Function: __XML_ComErrorHandler_UserFunction - mLipok
	. added ADO_CONSTANTS.au3 - mLipok

	??
	. __XML_ComErrorHandler_InternalFunction() CleanUp - mLipok
	. __XML_ComErrorHandler_InternalFunction() new parameters - mLipok

	SCRIPT BREAKING CHANGES
	. all function renamed from _XMLSomeName to _XML_SomeName ; TODO not yet all, but that is the plan - mLipok
	. Function _DebugWrite renamed to __XML_Misc_ConsoleNotifier - mLipok
	. Function _XSL_GetDefaultStyleSheet renamed to __XSL_GetDefaultStyleSheet - mLipok
	. Function _XMLGetDomVersion renamed to _XML_Misc_GetDomVersion - mLipok
	. Func _SetDebug renamed to _XML_MiscProperty_NotifyToConsole - mLipok
	. Removed $fDEBUGGING as is now used inside _XML_MiscProperty_NotifyToConsole - mLipok
	. Removed $bXMLAUTOSAVE as is now used inside _XML_MiscProperty_AutoSave - mLipok
	. Removed $bADDFORMATTING as is now used inside _XML_MiscProperty_AutoFormat - mLipok
	. Removed $g__iDOMVERSION as is now used inside __XML_MiscProperty_DomVersion - mLipok
	. standardarization for using Return Value for failure: return 0 and set @error to none 0 - mLipok
	. _XML_Load on success return $oXmlDoc instead 1 - mLipok
	. _XML_LoadXML on success return $oXmlDoc instead 1 - mLipok
	. $oXmlDoc is no longer global variable (must be passed to function as ByRef) - mLipok
	. removed many variable declaration used even for variable iteration "For $i to 10" - mLipok

	2015/09/04
	"1.1.1.02"
	. event handler renamed to $__g_oXMLDOM_EventsHandler - mLipok
	. added $XMLWRAPPER_ERROR_SAVEFILERO and checking in - mLipok
	. if variable is ObjectsCollection then variable name ends with sufix "Coll" - mLipok
	. variable used in loops are named with sufix "_idx" like this For $iAttribute_idx = 0 To $oAttributes_coll.length - 1 - mLipok
	. variable used in loops are named with sufix "Enum" like this For $oNode_enum in $oNodes_coll - mLipok
	. Local Variable renaming: $strQuery > $sQuery - mLipok - thanks to guinness
	. Local Variable renaming: $sComment > $sComment - mLipok - thanks to guinness
	. Function renamed fixed typo in function name _XML_CreateAttributeute => _XML_CreateAttribute - mLipok - thanks to guinness
	. Function renamed _XML_Misc_NodesList_GetNames >> _XML_Array_GetNodesNamesFromCollection - mLipok
	. Global Enum renamed from   $g__eXML_ .....    to    $__g_eXML_ ..... - mLipok
	. Function renamed _XML_ArrayAdd >> _XML_Array_AddName - mLipok
	. Function renamed _XML_Misc_NodesColl_GetNamesToArray >> _XML_Array_GetNodesNamesFromCollection - mLipok
	. Function renamed _XML_DeleteNode_XML_DeleteNode >> _XML_removeAll - mLipok
	. Function renamed _XML_FileOpen >> _XML_Load - mLipok
	. new Function _XML_Array_GetNodeList() - mLipok
	. UDF FileName: _XMLWrapperEx.au3 >> XMLWrapperEx.au3 - mLipok
	. !!! EXAMPLES FILE:    XMLWrapperEx__Examples.au3 - mLipok

	2015/09/06
	"1.1.1.03"
	. Renamed Function: _XMLCreateChildWAttr >> _XML_CreateChildWAttr - mLipok
	. Renamed ENUMs:  $__g_eXML_ERROR ... >> $XMLWRAPPER_ERROR ... - mLipok - Thanks to guinness
	. Renamed ENUMs:  $__g_eXML_RESULT ... >> $XMLWRAPPER_RESULT ... - mLipok - Thanks to guinness
	. Renamed CONST:  $XMLDOM_DOCUMENT_READYSTATE_ ... >> $XMLWRAPPER_DOCUMENT_READYSTATE_ ... - mLipok - Thanks to guinness
	. Renamed Variable : $oXML_Document >> $oXmlDoc  - for shortness, and compliance with examples of MSDN - mLipok
	. New Example: _Example_2__XML_CreateChildWAttr - mLipok
	. New Example: _Example_3__XML_Misc_ErrorDecription - mLipok
	. New Example: _Example_MSDN_1__setAttributeNode - mLipok
	. New Function: _XML_Misc_Viewer - mLipok
	. New: Function: _XML_Misc_ErrorParser() -- added as a replacment for using: _XML_Misc_ErrorDecription() - mLipok
	. NEW: @ERROR result: $XMLWRAPPER_ERROR_ARRAY - mLipok
	. NEW: @ERROR result: $XMLWRAPPER_ERROR_NODECREATEAPPEND - mLipok
	. Removed: all old _XML_Misc_ErrorDecription() - mLipok
	. Removed: _XML_Misc_ErrorDecription_Reset()- mLipok
	. Removed: all $sXML_error - mLipok
	. Removed: _XML_Error_Reason - mLipok
	. COMPLETED: @error checking for   .selectNodes  methods: @error is set to: $XMLWRAPPER_ERROR_XPATH - mLipok
	. COMPLETED: _XML_CreateChildWAttr - mLipok
	. COMPLETED: Standardarization for using Return Value @error codes returned from function --> ENUMERATED CONSTATNS - mLipok
	. 		Now all functions  Return like the following examples:
	.				Return SetError($XMLWRAPPER_ERROR_ .........
	.					or in this way:
	.				If @error Then Return SetError(@error, @extended, $XMLWRAPPER_RESULT_FAILURE)

	2015/09/07
	"1.1.1.04"
	. Renamed: $iXMLWrapper_Error_number >> $iXMLWrapper_Error - mLipok
	. Removed: almost all MsgBox() - leaves only in __XML_Misc_MsgBoxNotifier() - mLipok
	. Changed: in Examples : Global $oErrorHandler >> Local $oErrorHandler - mLipok
	. Renamed _XML_Array_GetNodeList >> _XML_Array_GetNodesProperties - mLipok
	. Removed: _XML_Array_GetNodesFromCollection as this was duplicate for _XML_Array_GetNodesProperties - mLipok
	. NEW: $__g_eARRAY_NODE_ATTRIBUTES in function _XML_Array_GetNodesProperties - now also display all atributes - mLipok
	. Renamed: _XML_MiscProperty_DomVersion >> __XML_MiscProperty_DomVersion - as this should be internal - as you can use _XML_Misc_GetDomVersion() - mLipok
	. Renamed: _XML_ComErrorHandler_MainFunction >> __XML_ComErrorHandler_UserFunction - as this should be internal - mLipok
	. Renamed: _XML_ErrorParser >> _XML_ErrorParser_GetDescription - mLipok
	. Removed: _XML_Misc_ErrorParser - will be in Examples - mLipok
	. Changed: Return value behavior - mLipok
	.	From:
	.		SetError($XMLWRAPPER_ERROR_PARSE, $oXmlDoc.parseError.errorCode, $oXmlDoc.parseError.reason)
	.		SetError($XMLWRAPPER_ERROR_PARSE, $oXmlDoc.parseError.errorCode, $XMLWRAPPER_RESULT_FAILURE)
	. Changed: _XML_Load and _XML_LoadXML require $oXmlDoc as first parameter, you must use _XML_CreateDOMDocument() as only in this way it is possible to use _XML_ErrorParser_GetDescription() - mLipok
	.		All examples changed to show how to use "new" _XML_Load and _XML_LoadXML
	. NEW: ENUMs: $XMLWRAPPER_ERROR_EMPTYCOLLECTION - mLipok
	. NEW: ENUMs: $XMLWRAPPER_ERROR_NONODESMATCH - mLipok
	. Changed: MAGIC NUMBERS: for StringStripWS in _XML_DeleteNode() - mLipok
	. COMPLETED: _XML_NodeExists - mLipok
	. COMPLETED: Refactroing all functions with "selectNodes" Method now have checking: If (Not IsObj($oNodes_coll)) Or $oNodes_coll.length = 0 Then ... SetError($XMLWRAPPER_ERROR_EMPTYCOLLECTION ..... - mLipok
	. COMPLETED: Refactroing all functions with "selectSingleNode" Method now have checking: If $oNode_Selected = Null Then .. SetError($XMLWRAPPER_ERROR_NONODESMATCH ..... - mLipok

	2015/09/09
	"1.1.1.05"
	. NEW: ENUMs: $XMLWRAPPER_ERROR_PARSE_XSL - mLipok
	. NEW: ENUMs: $XMLWRAPPER_ERROR_NOATTRMATCH - mLipok
	. COMPLETED: Function: _XML_RemoveAttribute() - mLipok
	. NEW EXAMPLE: _Example_4__XML_RemoveAttribute() - mLipok
	. Renamed: _XMLReplaceChild >> _XML_ReplaceChild - mLipok
	. COMPLETED: _XML_ReplaceChild - mLipok
	. NEW EXAMPLE: _Example_5__XML_ReplaceChild() - mLipok
	. NEW: ENUMs: $XMLWRAPPER_ERROR_NOCHILDMATCH - mLipok
	. NEW EXAMPLE: _Example_6__XML_GetChildNodes() - mLipok
	. COMPLETED: _XML_GetChildNodes() - mLipok
	. COMPLETED: _XML_GetAllAttribIndex() - mLipok
	. NEW Function: _XML_Array_GetAttributesProperties() - mLipok
	. NEW EXAMPLE: _Example_7__XML_GetAllAttribIndex() - mLipok
	. ADDED: many IsObj() >> https://www.autoitscript.com/forum/topic/177176-why-isobj-0-and-vargettype-object/

	2015/09/11
	"1.1.1.06"
	. Removed: Function: __XML_Misc_ConsoleNotifier() - mLipok
	. Removed: Function: __XML_Misc_MsgBoxNotifier() - mLipok
	. Removed: Function: _XML_MiscProperty_Notify() - mLipok
	. Removed: Function: _XML_MiscProperty_NotifyToConsole() - mLipok
	. Removed: Function: _XML_MiscProperty_NotifyToMsgBox() - mLipok
	. Removed: Function: _XML_MiscProperty_NotifyAll() - mLipok
	. Removed: Function: __XML_ComErrorHandler_MainFunction() - mLipok
	. Removed: Function: __XML_DOM_EVENT_ondataavailable() >> is in example - mLipok
	. Removed: Function: __XML_DOM_EVENT_onreadystatechange() >> is in example - mLipok
	. Removed: Function: __XML_DOM_EVENT_ontransformnode() >> is in example - mLipok
	. Removed: Function: _XML_UseEventHandler() - as event handler should be defined by user in main script - look in examlpes - mLipok
	. Removed: Function: _XML_ComErrorHandler_UseInternalAsUser() - mLipok
	. Renamed: Function: _XML_ComErrorHandler_UserFunction >> __XML_ComErrorHandler_UserFunction - is now internal - mLipok
	.		User Function are now passed as parameter to _XML_CreateDOMDocument()
	. Changed: Examples: XML_Misc_ErrorParser() >> XML_My_ErrorParser() - mLipok
	. Changed: Examples to fit to the changed UDF - mLipok
	. Modified: Examples: CleanUp +++ Comments - mLipok
	. Removed: Function: __AddFormat - as now is _XML_TIDY() function - mLipok
	. Removed: Function: _XML_MiscProperty_StaticCOMErrorHandler() - as was not used - mLipok
	.
	.
	. !!!!!! REMARKS: - mLipok
	.       It is user choice to set ERROR HANDLER or not.
	.		If user do not set it and something goes wrong with COM then this is USER PROBLEM and NOT UDF ISSUE/ERROR
	.			in such case UDF will call empty function for avoid AutoIt Error
	. 	FOR ERROR CHECKING
	.		1. check @error
	.		2. _XML_ErrorParser_GetDescription()
	.		3. setup your COM ERROR HANDLER and pass it as parameter to _XML_CreateDOMDocument()
	.		4. you can make in your main script function like XML_My_ErrorParser() and use it if you want

	2015/09/15
	"1.1.1.07"
	. Renamed: ENUMs: $eAttributeList_ ..... >> $__g_eARRAY_ATTR_ .... - as now are Global Enums - mLipok
	. Renamed: ENUMs: $eNodeList_ ..... >> $__g_eARRAY_NODE_ .... - as now are Global Enums - mLipok
	. Removed: Function: _XML_MiscProperty_AutoFormat() - as is not used - mLipok
	. Removed: Function: _XML_MiscProperty_EventHandling() - as is not used - mLipok
	. Removed: Function: _XML_MiscProperty_AutoSave() - as is not used - mLipok
	. ADDED NEW: Function: __XML_IsValidObject_Attributes() - mLipok
	. ADDED NEW: Function: __XML_IsValidObject_NodesColl() - mLipok
	. ADDED NEW: $XMLWRAPPER_EXT_ ... for proper handling @extended information - mLipok
	. Changed: Examples: XML_My_ErrorParser() - added support for $XMLWRAPPER_EXT_ ...  - mLipok
	. Removed: $sQuery      - DOM compliant query string (not really necessary as it becomes part of the path) - mLipok
	. Changed: All @extended are returned as $XMLWRAPPER_EXT_ .... or @extended - never as 0 or directly as number - mLipok
	. REFACTORED: _XML_Tidy() - proper @errors and @extended support - mLipok
	. REFACTORED: _XML_Array_GetAttributesProperties() - proper @errors and @extended support - mLipok
	. REFACTORED: _XML_Array_GetNodesProperties() - proper @errors and @extended support - mLipok
	. REFACTORED: all this following function uses _XML_SelectNodes() - mLipok
	.		_XMLCreateChildNode() 		_XML_DeleteNode() 			_XML_GetAllAttribIndex()
	.		_XML_GetParentNodeName() 	_XMLSetAttrib() 			_XML_UpdateField2()
	.		_XMLGetValue() 				_XML_GetNodesCount() 		_XML_ReplaceChild()
	.		_XML_GetNodesPath()			_XMLGetAllAttrib()
	. REFACTORED: all this following function uses _XML_SelectSingleNode() - mLipok
	.		_XML_CreateAttribute() 		_XML_CreateComment() 		_XML_GetAttrib()
	.		_XML_GetChildNodes() 		_XML_GetChildren() 			_XML_NodeExist()
	.		_XML_RemoveAttribute() 		_XML_GetChildText() 		_XML_UpdateField()
	.		_XMLGetField()
	. Renamed: Function: __XML_IsValidObject >> __XML_IsValidObject_DOMDocument - mLipok
	. COMPLETED: Function: _XML_SaveToFile - mLipok
	. Removed: $XMLWRAPPER_ERROR_NODECREATEAPPEND - mLipok
	. ADDED: $XMLWRAPPER_ERROR_NODECREATE - mLipok
	. ADDED: $XMLWRAPPER_ERROR_NODEAPPEND - mLipok
	. Changed: Functions to proper use $XMLWRAPPER_ERROR_NODECREATE  and $XMLWRAPPER_ERROR_NODEAPPEND - mLipok
	. Renamed: Function: __XML_IsValidObject_Nodes >> __XML_IsValidObject_NodesColl - mLipok
	. Renamed: ENUMs: $XMLWRAPPER_EXT_OK >> $XMLWRAPPER_EXT_DEFAULT - mLipok
	. Renamed: Function: _XML_GetNodeCount >> _XML_GetNodesCount - mLipok
	. Renamed: Function: _XML_GetAttrib >> _XML_GetNodeAttributeValue - mLipok
	. Changed: Function: _XML_GetNodeAttributeValue - parameters chagned - must pass $oNode instead $oXmlDoc - mLipok
	. COMPLETED: Function: _XML_LoadXML - mLipok
	. COMPLETED: Function: _XML_CreateDOMDocument - mLipok
	. COMPLETED: Function: _XML_Tidy - mLipok
	. Renamed: Function: _XML_CreateObject >> _XML_CreateDOMDocument - mLipok
	. NEW: Region: XMLWrapperEx.au3 - Functions - Not yet reviewed - mLipok
	. NEW: Region: XMLWrapperEx.au3 - Functions - COMPLETED - mLipok
	. COMPLETED: Function: _XML_GetNodeAttributeValue - mLipok
	. other CleanUp - mLipok

	2015/10/22
	"1.1.1.08"
	. ADDED: Description for _XML_ErrorParser_GetDescription() - mLipok
	. ADDED: Description for _XML_Array_GetNodesProperties() - mLipok
	. Fixed: Function: _XMLCreateFile () - mLipok
	.		Issue with usage of Tenary Operator in $oXmlDoc.createProcessingInstruction("xml", 'version="1.0"' & (($bUTF8) ? ' encoding="UTF-8"' : ''))
	. Renamed: Function: _XMLCreateFile >> _XML_CreateFile - mLipok
	. Completed: Function: _XML_CreateFile - mLipok
	. !!! EXAMPLES FILE: New Example: Example_9__XML_CreateFile() - mLipok
	. !!! EXAMPLES FILE: Changed to Keep the changed Script Breaking Changes - mLipok
	. Renamed: Function: __XML_ComErrorHandler_UserFunction >> _XML_ComErrorHandler_UserFunction : is now normal function - not internal - mLipok
	. NEW: methode of transfer UDF internal COM Error Handler to the user function - mLipok
	.		just use _XML_ComErrorHandler_UserFunction() like in example
	.
	SCRIPT BREAKING CHANGES
	. Renamed: Enums: $XMLWRAPPER_RESULT_ >> $XML_RET_ - mLipok
	. Renamed: Enums: $XMLWRAPPER_EXT_ >> $XML_EXT_ - mLipok
	. Renamed: Enums: $XMLWRAPPER_ERROR_ >> $XML_ERR_ - mLipok
	. Renamed: Enums: $NODE_ >> $XML_NODE_ - mLipok
	. Removed: Parameter: _XML_CreateDOMDocument > $vComErrFunc - mLipok


	2016/05/18
	"1.1.1.09"
	. !!!! UDF RENAMED XMLWrapperEx.au3 >> XML.au3 - mLipok
	. Changed: Error Handling: all SetError($XML_ERR_NODEAPPEND, - returns @error as extended - as this is COM ERROR - mLipok
	. Changed: Error Handling: all SetError($XML_ERR_NODECREATE, - returns @error as extended - as this is COM ERROR - mLipok
	. Removed: Function: _XMLCreateChildNode - as it was duplicate feature with _XML_CreateChildWAttr - mLipok
	.		Thanks to: @scila1996
	.		https://www.autoitscript.com/forum/topic/176895-xmlwrapperexau3-beta/?do=findComment&comment=1278825
	. Removed: Function: _XMLCreateChildNodeWAttr - as it was only duplicate/wrapper for _XML_CreateChildWAttr - mLipok
	.		Thanks to: @scila1996
	.		https://www.autoitscript.com/forum/topic/176895-xmlwrapperexau3-beta/?do=findComment&comment=1278825
	. ! EXAMPLES FILE: Modified: XML_My_ErrorParser - mLipok
	. ! EXAMPLES FILE: New: Example_2a__XML_CreateChildWAttr() - mLipok
	. Removed: Enums: $XML_ERR_SAVEFILERO - New Requirment for saving - File Can Not Exist - user should manage it by their own - mLipok
	. Renamed: Enums: $XML_ERR_ISNOTVALIDNODESE >> $XML_ERR_ISNOTVALIDNODETYPE - mLipok
	. Renamed: Enums: $XML_ERR_ISNOTVALIDNODETYPE >> $XML_ERR_INVALIDNODETYPE - mLipok
	. Renamed: Enums: $XML_ERR_ISNOTVALIDATTRIB >> $XML_ERR_INVALIDATTRIB - mLipok
	. Renamed: Enums: $XML_ERR_ISNOTVALIDDOMDOC >> $XML_ERR_INVALIDDOMDOC - mLipok
	. Removed: $XML_EXT_GENERAL >> $XML_EXT_DEFAULT - mLipok
	. Changed: $XML_EXT_.. are reordered - mLipok
	. Removed: Function: _XSL_GetDefaultStyleSheet - mLipok
	.	This was example from:
	.	http://www.xml.com/lpt/a/1681
	.	But it is: Copyright © 1998-2006 O'Reilly Media, Inc.
	. Renamed: Function: _XMLGetField >> _XML_GetField - mLipok
	. Renamed: Function: _XMLGetValue >> _XML_GetValue - mLipok
	. Renamed: Function: _XMLGetAllAttrib >> _XML_GetAllAttrib - mLipok
	. Renamed: Function: _XMLSetAttrib >> _XML_SetAttrib - mLipok
	. New: Function: _XML_InsertChildNode - GMK
	. New: Function: _XML_InsertChildWAttr - GMK
	. Changed: Function: _XML_CreateAttribute - numbers of parameters - mLipok
	.			now you must pass an Array with AttributeName and AttributeValue
	. Fixed: Function: _XMLCreateRootNode - GMK
	.		$oXmlDoc.documentElement.appendChild($oChild) >> $oXmlDoc.appendChild($oChild)
	. Fixed: Function: _XMLCreateRootNodeWAttr - GMK
	.		$oXmlDoc.documentElement.appendChild($oChild_Node) >> $oXmlDoc.appendChild($oChild_Node)
	. Renamed: Function: _XMLCreateRootChild >> _XML_CreateRootNode - GMK
	.			!!!!!!!! @TODO need to be revisited
	.
	. Renamed: Function: _XMLCreateRootNodeWAttr >> _XML_CreateRootNodeWAttr - GMK
	. ADDED: #CURRENT# - GMK
	. ADDED: #IN_PROGESS# - GMK
	. ADDED: #INTERNAL_USE_ONLY# - GMK
	. 		!!! Additional Thanks for GMK for testing and many changes in many Description
	. CleanUp: Function: _XML_GetNodesPath - removed $sNodePathTag - mLipok - thanks to GMK
	. CleanUp: Function: _XML_GetParentNodeName - removed $sNodePathTag - mLipok - thanks to GMK
	. CleanUp: Function: removed #include <MsgBoxConstants.au3> - mLipok - thanks to GMK
	. CleanUp: Function: _XML_GetField - removed $oChild - mLipok - thanks to GMK
	. CleanUp: Function: _XML_GetNodesPath - MagicNumber 0 replaced with $STR_NOCASESENSE - mLipok - thanks to GMK
	. CleanUp: Function: _XML_GetNodesPathInternal - MagicNumber 0 replaced with $STR_NOCASESENSE - mLipok - thanks to GMK
	. CleanUp: Function: _XML_GetParentNodeName - MagicNumber 0 replaced with $STR_NOCASESENSE - mLipok - thanks to GMK
	. Renamed: Function: _XMLCreateCDATA >> _XML_CreateCDATA - mLipok - thanks to GMK
	. Rewrite: Function: _XML_GetAllAttrib - Parameters : removed ByRef $aName, ByRef $aValue - mLipok
	. Fixed Typo: Descripton: Chceck >> Check - mLipok - thanks to GMK
	. Added: Descripton: _XML_GetNodesCount - mLipok - thanks to GMK
	. Changed: Descripton: _XML_TransformNode - mLipok - thanks to GMK
	. Changed: Descripton: _XML_CreateDOMDocument - mLipok - thanks to GMK
	. Added: Descripton: _XML_GetNodeAttributeValue - mLipok - thanks to GMK
	. Changed: Descripton: _XML_Misc_Viewer - mLipok - thanks to GMK
	.
	. Changed: Function: _XML_SelectNodes in case of success @extended = $oNodes_coll.length
	.
	.
	2016/05/18
	"1.1.1.10"
	. NEW: Feature: _XML_Tidy:   if Parameter $sEncoding = -1 then .omitXMLDeclaration = true - mLipok
	.			Feature asked by @GMK here:
	.			https://www.autoitscript.com/forum/topic/176895-xmlau3-v-11109-formerly-xmlwrapperexau3-beta-support-topic/?do=findComment&comment=1294688
	. Changed: _XML_Tidy(ByRef $oXmlDoc, $sEncoding = -1) - Default value is set to -1    - mLipok
	. New:	XML__Example_TIDY.au3 - mLipok
	.
	.
	2016/06/16
	"1.1.1.11"
	. NEW: Function: __XML_IsValidObject_DOMDocumentOrElement - mLipok
	. Changed: Function: _XML_SelectNodes - parameter validation __XML_IsValidObject_DOMDocumentOrElement - mLipok
	. 		Currently _XML_SelectNodes can use relative XPath
	. Changed: Function: _XML_SelectSingleNode - parameter validation __XML_IsValidObject_DOMDocumentOrElement - mLipok
	. 		Currently _XML_SelectSingleNode can use relative XPath
	. Refactored: Function: _XML_CreateComment - mLipok
	. Changed: Function: @error > $XML_ERR_COMERROR - mLipok
	. Fixed: Function: _XML_Array_GetNodesProperties - properly gets all attributes - mLipok
	.
	. EXAMPLES: New, and checked/refactored/fixed
	.	XML__Examples_TIDY.au3
	.	XML__Examples_User__asdf1nit.au3
	.	XML__Examples_User_coma.au3
	.	XML__Examples_User_Realm.au3
	.	XML__Examples_User_scila1996.au3
	.	XML__Examples_User_DarkAqua__Tasks.au3


	2016/10/27
	"1.1.1.12"
	. Changed: Function: _XML_SetAttrib - support for $vAttributeNameOrList - GMK
	. Added: Enums:  $XMLATTR_COLNAME, $XMLATTR_COLVALUE, $XMLATTR_COLCOUNTER - mLipok
	. Changed: Function: _XML_GetAllAttrib - !!! array result is reordered ROWS<>COLS - mLipok
	.			now are coherent manner for: _XML_InsertChildWAttr, _XML_CreateChildWAttr, _XML_SetAttrib, _XML_GetAllAttrib
	.              THIS IS !!! SCRIPT BREAKING CHANGE !!!
	. Added: Function parameter: _XML_Load new parameter $bPreserveWhiteSpace = True - GMK
	. Added: Function parameter: _XML_LoadXML new parameter $bPreserveWhiteSpace = True - GMK
	. Changed: Enums:  $XML_ERR_OK >> $XML_ERR_SUCCESS - for unification/coherence in relatation to some other UDF's - mLipok
	.
	. EXAMPLES: New, and checked/refactored/fixed
	.	XML__Examples_TIDY2.au3
	.	XML__Examples_User_BlaBlaFoo__Dellwarranty.au3
	.	XML__Examples_User_Shrapnel.au3
	.
	.

	2017/03/05
	"1.1.1.13"
	. Added: Function: _XML_Base64Encode() - mLipok
	. Added: Function: _XML_Base64Decode() - mLipok
	. Fixed: bug in: _XML_Array_AddName() - krupa
	. Changed: $ADOENUM_ad* >>> $ADO_ad** - to be coherent with ADO.au3 UDF - mLipok
	.
	2017/xx/xx
	"1.1.1.xx"
	.
	.
	.
	.
	.
	.

	@LAST - this keyword is usefull for quick jumping here

	TODO LIST:
	. TODO CHECK: _XML_GetField, _XML_GetValue
	. @WIP TODO: COUNT = 50
	. TODO: Description, Function Header CleanUp (are still old)
	. TODO: browse entire UDF for TODO "Keyword"
	. TODO: Return SetError($XML_ERR_GENERAL ... should be used only once per function
	. TODO: Return SetError($XML_ERR_GENERAL ... should be always ONLY as the last Error returned from function
	. TODO: $XML_ERR_ .... should be reordered it will be SCRIPT BREAKING CHANGES: only if used MAGIC NUMBERS for @error checking
	. TODO: GMK: What's a better way to insert a node before a specified node object or XPath for _XML_InsertChildNode and _XML_InsertChildWAttr?  Replace $iItem with $oInsertBeforeNode?
	. TODO: GMK: Rename _XML_Transform ==> _XML_TransformNodeToObj (?)
	. TODO: GMK: Why not combine _XML_UpdateField and _XML_UpdateField2?  Would inputting parameters for a single node XPath not work the same for _XML_UpdateField2 as it would for _XML_UpdateField?

#CE

; #CURRENT# =====================================================================================================================
; _XML_CreateAttribute
; _XML_CreateComment
; _XML_DeleteNote
; _XML_GetChildren
; _XML_GetChildText
; _XML_GetNodesPath
; _XML_GetNodesPathInternal
; _XML_GetParentNodeName
; _XML_RemoveAttributeNode
; _XML_ReplaceChild
; _XML_Transform
; _XML_UpdateField
; _XML_UpdateField2
; _XML_ValidateFile
; _XML_CreateCDATA
; _XML_CreateChildNode
; _XML_CreateRootNode
; _XML_CreateRootNodeWAttr
; _XML_GetAllAttrib
; _XML_GetAllAttribIndex
; _XML_GetField
; _XML_GetValue
; _XML_SetAttrib
; _XML_CreateChildWAttr
; _XML_CreateDOMDocument
; _XML_CreateFile
; _XML_GetChildNodes
; _XML_GetNodeAttributeValue
; _XML_Load
; _XML_LoadXML
; _XML_NodeExists
; _XML_SaveToFile
; _XML_SelectNodes
; _XML_SelectSingleNode
; _XML_Tidy
; _XML_Misc_GetDomVersion
; _XML_Misc_Viewer
; _XML_MiscProperty_UDFVersion
; _XML_Array_AddName
; _XML_Array_GetAttributesProperties
; _XML_Array_GetNodesProperties
; _XML_ErrorParser_GetDescription

; #IN_PROGESS# ==================================================================================================================
; _XML_InsertChildNode
; _XML_InsertChildWAttr
; _XML_GetNodesCount
; _XML_RemoveAttribute
; _XML_TransformNode
; _XML_MiscProperty_Encoding

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __XML_IsValidObject_Attributes
; __XML_IsValidObject_DOMDocument
; __XML_IsValidObject_DOMDocumentOrElement
; __XML_IsValidObject_Node
; __XML_IsValidObject_NodesColl
; __XML_MiscProperty_DomVersion
; __XML_ComErrorHandler_InternalFunction
; __XML_ComErrorHandler_UserFunction

#EndRegion XML.au3 - UDF Header
; #VARIABLES# ===================================================================================================================
#Region XML.au3 - Enumeration - ERROR EXTENDED RETURN
Global Enum _
		$XML_ERR_SUCCESS, _ ; 				All is ok
		$XML_ERR_GENERAL, _ ; 			The error which is not specifically defined.
		$XML_ERR_COMERROR, _ ; 			COM Error occured - Check @extended for COMERROR Number or check you COM Error Handler Function
		$XML_ERR_DOMVERSION, _ ; 		Check @extended for required DOM Version
		$XML_ERR_ISNOTOBJECT, _ ; 		No object passed to function.
		$XML_ERR_INVALIDDOMDOC, _ ; 	Invalid object passed to function - expected Document.
		$XML_ERR_INVALIDATTRIB, _ ; 	Invalid object passed to function - expected Attrib Element.
		$XML_ERR_INVALIDNODETYPE, _ ; 	Invalid object passed to function - expected Node Element.
		$XML_ERR_OBJCREATE, _ ; 		Object can not be created.
		$XML_ERR_NODECREATE, _ ; 		Can not create Node - check also extended for node type
		$XML_ERR_NODEAPPEND, _ ; 		Can not append Node - check also extended for node type
		$XML_ERR_PARSE, _ ; 			Error with Parsing objects use _XML_ErrorParser_GetDescription() for Get details
		$XML_ERR_PARSE_XSL, _ ; 		Error with Parsing XSL objects use _XML_ErrorParser_GetDescription() for Get details
		$XML_ERR_LOAD, _ ; 				Error opening specified file.
		$XML_ERR_SAVE, _ ; 				Error saving file.
		$XML_ERR_PARAMETER, _ ; 		Wrong parameter passed to function.
		$XML_ERR_ARRAY, _ ; 			Wrong array parameter passed to function. Check array dimension and conent.
		$XML_ERR_XPATH, _ ; 			XPath syntax error - you should check also COM Error Handler.
		$XML_ERR_NONODESMATCH, _ ;  	No nodes match the XPath expression
		$XML_ERR_NOCHILDMATCH, _ ; 		There is no Child in nodes matched by XPath expression.
		$XML_ERR_NOATTRMATCH, _ ; 		There is no such attribute in selected node.
		$XML_ERR_EMPTYCOLLECTION, _ ; 	Collections of objects was empty
		$XML_ERR_EMPTYOBJECT, _ ; 		Object is empty
		$XML_ERR_ENUMCOUNTER ; not used in UDF - just for other/future testing
Global Enum _
		$XML_EXT_DEFAULT, _ ; 					Default - Do not return any additional information, The extended which is not specifically defined.
		$XML_EXT_PARAM1, _ ;					Error Occurs in 1-Parameter
		$XML_EXT_PARAM2, _ ;					Error Occurs in 2-Parameter
		$XML_EXT_PARAM3, _ ;					Error Occurs in 3-Parameter
		$XML_EXT_PARAM4, _ ;					Error Occurs in 4-Parameter
		$XML_EXT_PARAM5, _ ;					Error Occurs in 5-Parameter
		$XML_EXT_PARAM6, _ ;					Error Occurs in 6-Parameter
		$XML_EXT_XMLDOM, _ ; 					"Microsoft.XMLDOM" related Error
		$XML_EXT_DOMDOCUMENT, _ ; 				"Msxml2.DOMDocument" related Error
		$XML_EXT_XSLTEMPLATE, _ ; 				"Msxml2.XSLTemplate" related Error
		$XML_EXT_SAXXMLREADER, _ ; 				"MSXML2.SAXXMLReader" related Error
		$XML_EXT_MXXMLWRITER, _ ; 				"MSXML2.MXXMLWriter" related Error
		$XML_EXT_FREETHREADEDDOMDOCUMENT, _ ; 	"Msxml2.FreeThreadedDOMDocument" related Error
		$XML_EXT_XMLSCHEMACACHE, _ ; 			"Msxml2.XMLSchemaCache." related Error
		$XML_EXT_STREAM, _ ; 					"ADODB.STREAM" related Error
		$XML_EXT_ENCODING, _ ; 					Encoding related Error
		$XML_EXT_ENUMCOUNTER ; not used in UDF - just for other/future testing
Global Enum _
		$XML_RET_FAILURE, _ ;			UDF Default Failure Return Value
		$XML_RET_SUCCESS, _ ;			UDF Default Success Return Value
		$XML_RET_ENUMCOUNTER ; not used in UDF - just for other/future testing
#EndRegion XML.au3 - Enumeration - ERROR EXTENDED RETURN

#Region XML.au3 - ARRAY Enums
; Enums for _XML_Array_GetAttributesProperties() function
Global Enum _
		$__g_eARRAY_ATTR_NAME, _
		$__g_eARRAY_ATTR_TYPESTRING, _
		$__g_eARRAY_ATTR_VALUE, _
		$__g_eARRAY_ATTR_TEXT, _
		$__g_eARRAY_ATTR_DATATYPE, _
		$__g_eARRAY_ATTR_XML, _
		$__g_eARRAY_ATTR_ARRAYCOLCOUNT
; Enums for _XML_Array_GetNodesProperties() function
Global Enum _
		$__g_eARRAY_NODE_NAME, _
		$__g_eARRAY_NODE_TYPESTRING, _
		$__g_eARRAY_NODE_VALUE, _
		$__g_eARRAY_NODE_TEXT, _
		$__g_eARRAY_NODE_DATATYPE, _
		$__g_eARRAY_NODE_XML, _
		$__g_eARRAY_NODE_ATTRIBUTES, _
		$__g_eARRAY_NODE_ARRAYCOLCOUNT

; Enums for _XML_InsertChildWAttr, _XML_CreateChildWAttr, _XML_SetAttrib, _XML_GetAllAttrib
Global Enum _
		$XMLATTR_COLNAME, _
		$XMLATTR_COLVALUE, _
		$XMLATTR_COLCOUNTER
#EndRegion XML.au3 - ARRAY Enums

#Region XML.au3 - XML DOM Enumerated Constants
;~ https://msdn.microsoft.com/en-us/library/ms766473(v=vs.85).aspx
Global Const $XML_NODE_ELEMENT = 1
Global Const $XML_NODE_ATTRIBUTE = 2
Global Const $XML_NODE_TEXT = 3
Global Const $XML_NODE_CDATA_SECTION = 4
Global Const $XML_NODE_ENTITY_REFERENCE = 5
Global Const $XML_NODE_ENTITY = 6
Global Const $XML_NODE_PROCESSING_INSTRUCTION = 7
Global Const $XML_NODE_COMMENT = 8
Global Const $XML_NODE_DOCUMENT = 9
Global Const $XML_NODE_DOCUMENT_TYPE = 10
Global Const $XML_NODE_DOCUMENT_FRAGMENT = 11
Global Const $XML_NODE_NOTATION = 12
#EndRegion XML.au3 - XML DOM Enumerated Constants

#Region XML.au3 - readyState Property (DOMDocument)
;~ https://msdn.microsoft.com/en-us/library/ms753702(v=vs.85).aspx
Global Const $XML_DOCUMENT_READYSTATE_LOADING = 1
Global Const $XML_DOCUMENT_READYSTATE_LOADED = 2
Global Const $XML_DOCUMENT_READYSTATE_INTERACTIVE = 3
Global Const $XML_DOCUMENT_READYSTATE_COMPLETED = 4
#EndRegion XML.au3 - readyState Property (DOMDocument)

#Region XML.au3 - Functions - Not yet reviewed
; ===============================================================================================================================
; XMLWrapper functions

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_CreateAttribute
; Description ...: Adds an XML Attribute to specified node.
; Syntax ........: _XML_CreateAttribute(Byref $oXmlDoc, $sXPath, $asAttributeList)
; Parameters ....: $oXmlDoc             - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath              - a string value. The XML tree path from root node (root/child/child..)
;                  $asAttributeList		- an array of strings. Column0=AtributeName, Column1=AtributeValue
; Return values .: On Success      	    - Returns $XML_RET_SUCCESS
;                  On Failure           - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok, GMK
; Remarks .......:
; Related .......: https://msdn.microsoft.com/en-us/library/ms754616(v=vs.85).aspx
; Link ..........: https://msdn.microsoft.com/en-us/library/ms764615(v=vs.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _XML_CreateAttribute(ByRef $oXmlDoc, $sXPath, $asAttributeList)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	If Not UBound($asAttributeList) Or UBound($asAttributeList, $UBOUND_COLUMNS) <> 2 Then
		Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM3, $XML_RET_FAILURE)
	EndIf

	Local Enum $eAttribute_Name, $eAttribute_Value
	Local $oAttribute = Null
	#forceref $oAttribute
	For $iAttribute_idx = 0 To UBound($asAttributeList) - 1
		$oAttribute = $oXmlDoc.createAttribute($asAttributeList[$iAttribute_idx][$eAttribute_Name]) ;, $sNameSpace) ; TODO Check why $sNameSpace
		If @error Then Return SetError($XML_ERR_COMERROR, @error, $XML_RET_FAILURE)

		$oNode_Selected.SetAttribute($asAttributeList[$iAttribute_idx][$eAttribute_Name], $asAttributeList[$iAttribute_idx][$eAttribute_Value])
		If $oXmlDoc.parseError.errorCode Then
			Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
		EndIf
	Next

	; CleanUp
	$oAttribute = Null
	$oNode_Selected = Null

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_CreateAttribute

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_CreateCDATA
; Description ...: Create a CDATA SECTION node directly under root.
; Syntax ........: _XML_CreateCDATA(ByRef $oXmlDoc, $sNode, $sCDATA[, $sNameSpace = ""])
; Parameters ....: $oXmlDoc    - [in/out] an object. A valid DOMDocument object.
;                  $sNode      - a string value. name of node to create
;                  $sCDATA     - a string value. CDATA value
;                  $sNameSpace - a string value. the namespace to specifiy if the xml uses one.
; Return values .: On Success  - Returns $XML_RET_SUCCESS
;                  On Failure  - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......: fixme, won't append to exisiting node. must create new node.
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_CreateCDATA(ByRef $oXmlDoc, $sNode, $sCDATA, $sNameSpace = "")
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oNode = $oXmlDoc.createNode($XML_NODE_ELEMENT, $sNode, $sNameSpace)
	If @error Then Return SetError($XML_ERR_NODECREATE, @error, $XML_RET_FAILURE)

	If IsObj($oNode) Then
		Local $oChild = $oXmlDoc.createCDATASection($sCDATA)
		$oNode.appendChild($oChild)
		If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)

		$oXmlDoc.documentElement.appendChild($oNode)
		If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)

		$oChild = Null
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_CreateCDATA

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_CreateComment
; Description ...: Create a COMMENT node at specified path.
; Syntax ........: _XML_CreateComment(ByRef $oXmlDoc, $sXPath, $sComment)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
;                  $sComment  - a string value. The comment to add the to the xml file.
; Return values .: On Success - Returns $XML_RET_SUCCESS
;                  On Failure - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_CreateComment(ByRef $oXmlDoc, $sXPath, $sComment)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oChild = $oXmlDoc.createComment($sComment)
	$oNode_Selected.insertBefore($oChild, $oNode_Selected.childNodes(0))
	If @error Then
		Return SetError($XML_ERR_COMERROR, @error, $XML_RET_SUCCESS)
	ElseIf $oXmlDoc.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS) ; TODO Check for what we need to return on success
EndFunc   ;==>_XML_CreateComment

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_CreateRootNode
; Description ...: Create node directly under root.
; Syntax ........: _XML_CreateRootNode(ByRef $oXmlDoc, $sNode[, $sData = ""[, $sNameSpace = ""]])
; Parameters ....: $oXmlDoc    - [in/out] an object. A valid DOMDocument object.
;                  $sNode      - a string value. The name of node to create.
;                  $sData      - a string value. The optional value to create
;                  $sNameSpace - a string value. the namespace to specifiy if the file uses one.
; Return values .: On Success  - Returns $XML_RET_SUCCESS
;                  On Failure  - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok, GMK
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_CreateRootNode(ByRef $oXmlDoc, $sNode, $sData = "", $sNameSpace = "")
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oChild = $oXmlDoc.createNode($XML_NODE_ELEMENT, $sNode, $sNameSpace)
	If @error Then Return SetError($XML_ERR_NODECREATE, @error, $XML_RET_FAILURE)

	If IsObj($oChild) Then
		If $sData <> "" Then $oChild.text = $sData
		$oXmlDoc.appendChild($oChild)
		If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)

		$oChild = 0
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_CreateRootNode

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_CreateRootNodeWAttr
; Description ...: Create a child node under root node with attributes.
; Syntax.........: _XML_CreateRootNodeWAttr($sNode, $aAttribute_Names, $aAttribute_Values[, $sData = ""[, $sNameSpace = ""]])
; Parameters ....: $sNode       - The node to add with attibute(s)
;                  $aAttribute_Names         - The attribute name(s) -- can be array
;                  $aAttribute_Values          - The	attribute value(s) -- can be array
;                  $sData       - The optional value to give the node.
; Return values .: Success        $XML_RET_SUCCESS
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......: This function requires that each attribute name has a corresponding value.
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_CreateRootNodeWAttr(ByRef $oXmlDoc, $sNode, $aAttribute_Names, $aAttribute_Values, $sData = "", $sNameSpace = "")
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oChild_Node = $oXmlDoc.createNode($XML_NODE_ELEMENT, $sNode, $sNameSpace)
	If @error Then Return SetError($XML_ERR_NODECREATE, @error, $XML_RET_FAILURE)

	If IsObj($oChild_Node) Then
		If $sData <> "" Then $oChild_Node.text = $sData

		Local $oAttribute = Null
		#forceref $oAttribute
		If IsArray($aAttribute_Names) And IsArray($aAttribute_Values) Then
			If UBound($aAttribute_Names) <> UBound($aAttribute_Values) Then
				Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
			EndIf

			For $iAttribute_idx = 0 To UBound($aAttribute_Names) - 1
				If $aAttribute_Names[$iAttribute_idx] = "" Then
					Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
				EndIf

				$oAttribute = $oXmlDoc.createAttribute($aAttribute_Names[$iAttribute_idx]) ;, $sNameSpace) ; TODO Check why $sNameSpace
				$oChild_Node.SetAttribute($aAttribute_Names[$iAttribute_idx], $aAttribute_Values[$iAttribute_idx])
			Next
		Else
			$oAttribute = $oXmlDoc.createAttribute($aAttribute_Names)
			$oChild_Node.SetAttribute($aAttribute_Names, $aAttribute_Values)
		EndIf
		$oXmlDoc.appendChild($oChild_Node)
		If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)

		$oChild_Node = Null
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_CreateRootNodeWAttr

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_DeleteNode
; Description ...: Deletes XML Node based on XPath input from root node.
; Syntax ........: _XML_DeleteNode(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success - Returns $XML_RET_SUCCESS
;                  On Failure - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
;                  |0 - No error
;                  |1 - Deletion error
;                  |2 - No object passed
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_DeleteNode(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	For $oNode_enum In $oNodes_coll
		If $oNode_enum.hasChildNodes Then
			For $oNode_enum_Child In $oNode_enum.childNodes
				If $oNode_enum_Child.nodeType = $XML_NODE_TEXT Then
					If StringStripWS($oNode_enum_Child.text, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES) = "" Then
						$oNode_enum.removeChild($oNode_enum_Child)
					EndIf
				EndIf
			Next
		EndIf
		$oNode_enum.parentNode.removeChild($oNode_enum)
	Next

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_DeleteNode

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetAllAttrib
; Description ...: Get all XML Field(s) attributes based on XPath input from root node.
; Syntax ........: _XML_GetAllAttrib(ByRef $oXmlDoc, $sXPath, ByRef $aName, ByRef $aValue)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success - Returns array of fields text values (number of items is in [0][0])
;                  On Failure - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetAllAttrib(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oAttributes_coll = Null
	Local $aResponse[1][$XMLATTR_COLCOUNTER]
	For $oNode_enum In $oNodes_coll
		$oAttributes_coll = $oNode_enum.attributes
		If ($oAttributes_coll.length) = 0 Then
			ContinueLoop
			Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
		EndIf

		ReDim $aResponse[$oAttributes_coll.length + 2][$XMLATTR_COLCOUNTER]
		For $iAttribute_idx = 0 To $oAttributes_coll.length - 1
			$aResponse[$iAttribute_idx + 1][$XMLATTR_COLNAME] = $oAttributes_coll.item($iAttribute_idx).nodeName
			$aResponse[$iAttribute_idx + 1][$XMLATTR_COLVALUE] = $oAttributes_coll.item($iAttribute_idx).Value
		Next
	Next
	$aResponse[0][0] = $oAttributes_coll.length

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $aResponse)
EndFunc   ;==>_XML_GetAllAttrib

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetChildren
; Description ...: Selects XML child Node(s) of an element based on XPath input from root node and returns there text values.
; Syntax ........: _XML_GetChildren(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc      - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath       - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success    - Returns an array where:
;                  |$array[0][0] = Size of array
;                  |$array[1][0] = Name
;                  |$array[1][1] = Text
;                  |$array[1][2] = NameSpaceURI
;                  |...
;                  |$array[n][0] = Name
;                  |$array[n][1] = Text
;                  |$array[n][2] = NamespaceURI
;                  On Failure    - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetChildren(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then
		Return SetError(@error, @extended, $XML_RET_FAILURE)
	ElseIf $oNode_Selected.hasChildNodes() Then
		Local $iResponseDimSize = 0
		Local $aResponse[1][3] = [[$iResponseDimSize, '', '']]
		For $oNode_enum_Child In $oNode_Selected.childNodes()
			If $oNode_enum_Child.nodeType() = $XML_NODE_ELEMENT And $oNode_enum_Child.hasChildNodes() Then
				For $oNode_enum_Descendant In $oNode_enum_Child.childNodes()
					If $oNode_enum_Descendant.nodeType() = $XML_NODE_TEXT Then
						$iResponseDimSize = UBound($aResponse, 1)
						ReDim $aResponse[$iResponseDimSize + 1][3]
						$aResponse[$iResponseDimSize][0] = $oNode_enum_Descendant.parentNode.baseName
						$aResponse[$iResponseDimSize][1] = $oNode_enum_Descendant.text
						$aResponse[$iResponseDimSize][2] = $oNode_enum_Descendant.NamespaceURI
						; TODO Check
						; _XML_Array_AddName($aResponse, $oNode_enum_Child.baseName)
					EndIf
				Next
			EndIf
		Next
		$aResponse[0][0] = $iResponseDimSize

		; TODO Description for @extended
		Return SetError($XML_ERR_SUCCESS, $iResponseDimSize, $aResponse)
	EndIf

	Return SetError($XML_ERR_NOCHILDMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_GetChildren

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetChildText
; Description ...: Selects XML child Node(s) of an element based on XPath input from root node.
; Syntax ........: _XML_GetChildText(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc    - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath     - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success  - Returns an array of Node's names and text.
;                  On Failure  - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetChildText(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then
		Return SetError(@error, @extended, $XML_RET_FAILURE)
	ElseIf $oNode_Selected.hasChildNodes() Then
		Local $aResponse[1] = [0]
		For $oNode_enum_Child In $oNode_Selected.childNodes()
			If $oNode_enum_Child.nodeType = $XML_NODE_ELEMENT Then
				_XML_Array_AddName($aResponse, $oNode_enum_Child.baseName)
			ElseIf $oNode_enum_Child.nodeType = $XML_NODE_TEXT Then
				_XML_Array_AddName($aResponse, $oNode_enum_Child.text)
			EndIf
		Next

		$aResponse[0] = UBound($aResponse) - 1
		Return SetError($XML_ERR_SUCCESS, $aResponse[0], $aResponse) ; TODO Description for @extended
	EndIf

	Return SetError($XML_ERR_NOCHILDMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_GetChildText

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetField
; Description ...: Get XML Field(s) based on XPath input from root node.
; Syntax ........: _XML_GetField(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success - Returns an array of fields text values (count is in first element)
;                  On Failure - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok, GMK
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetField(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then
		Return SetError(@error, @extended, $XML_RET_FAILURE)
	ElseIf $oNode_Selected.hasChildNodes() Then
		Local $aResponse[1] = [0], $sNodePath = ""
		Local $iNodeMaxCount = $oNode_Selected.childNodes.length
		Local $aRet
		For $iNode_idx = 1 To $iNodeMaxCount
			If $oNode_Selected.parentNode.nodeType = $XML_NODE_DOCUMENT Then
				$sNodePath = "/" & $oNode_Selected.baseName & "/*[" & $iNode_idx & "]"
			Else
				$sNodePath = $oNode_Selected.baseName & "/*[" & $iNode_idx & "]"
			EndIf

			$aRet = _XML_GetValue($oXmlDoc, $sNodePath)
			If UBound($aRet) > 1 Then
				_XML_Array_AddName($aResponse, $aRet[1])
			Else
				_XML_Array_AddName($aResponse, "")
			EndIf
		Next
		$aResponse[0] = UBound($aResponse) - 1
		Return SetError($XML_ERR_SUCCESS, $aResponse[0], $aResponse) ; TODO Description for @extended

	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_GetField

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetNodesPath
; Description ...: Return a nodes full path based on XPath input from root node.
; Syntax ........: _XML_GetNodesPath(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success - An array of node names from root, count in [0] element.
;                  On Failure - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetNodesPath(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	If _XML_Misc_GetDomVersion() < 4 Then
		Return SetError($XML_ERR_DOMVERSION, 4, $XML_RET_FAILURE) ; TODO @extended Description
	EndIf

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $aResponse[1], $sNodePath, $sNameSpace

	Local $oParent, $oNode_enum_Temp
	For $oNode_enum In $oNodes_coll
		$oNode_enum_Temp = $oNode_enum
		$sNodePath = ""
		If $oNode_enum.nodeType <> $XML_NODE_DOCUMENT Then
			$sNameSpace = $oNode_enum.namespaceURI()
			If $sNameSpace <> "" Then
				$sNameSpace = StringRight($sNameSpace, StringLen($sNameSpace) - StringInStr($sNameSpace, "/", $STR_NOCASESENSE, -1)) & ":"
			EndIf
			If $sNameSpace = 0 Then $sNameSpace = ""
			$sNodePath = "/" & $sNameSpace & $oNode_enum.nodeName() & $sNodePath
		EndIf

		Do
			$oParent = $oNode_enum_Temp.parentNode()
			If $oParent.nodeType <> $XML_NODE_DOCUMENT Then
				$sNameSpace = $oParent.namespaceURI()
				If $sNameSpace <> "" Then
					; $sNameSpace = StringRight($sNameSpace, StringLen($sNameSpace) - StringInStr($sNameSpace, "/", $STR_NOCASESENSE, -1)) & ":"
					$sNameSpace &= ":"
				EndIf
				If $sNameSpace = 0 Then $sNameSpace = ""
				$sNodePath = "/" & $sNameSpace & $oParent.nodeName() & $sNodePath
				$oNode_enum_Temp = $oParent
			Else
				$oNode_enum_Temp = Null
			EndIf
			$oParent = Null
		Until (Not (IsObj($oNode_enum_Temp)))

		_XML_Array_AddName($aResponse, $sNodePath)
	Next

	$aResponse[0] = UBound($aResponse) - 1
	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $aResponse)
EndFunc   ;==>_XML_GetNodesPath

; #FUNCTION# ===================================================================
; Name ..........: _XML_GetNodesPathInternal
; Description ...: Returns the path of a valid node object.
; Syntax ........: _XML_GetNodesPathInternal(ByRef $oXML_Node)
; Parameters ....: $oXML_Node - A valid node object
; Return values .: On Success - Path from root as string.
;                  On Failure - @TODO
;                  On Failure - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
;                  On Failure - An empty string and @error set to 1.
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetNodesPathInternal(ByRef $oXML_Node)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXML_Node)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $sNodePath = "/" & $oXML_Node.baseName
	Local $oParentNode = Null, $sNameSpace = ''

	Do
		$oParentNode = $oXML_Node.parentNode()

		If $oParentNode.nodeType <> $XML_NODE_DOCUMENT Then
			$sNameSpace = $oParentNode.namespaceURI()
			If $sNameSpace = 0 Then $sNameSpace = ""
			If $sNameSpace <> "" Then
				$sNameSpace = StringRight($sNameSpace, StringLen($sNameSpace) - StringInStr($sNameSpace, "/", $STR_NOCASESENSE, -1)) & ":"
			EndIf
			$sNodePath = "/" & $sNameSpace & $oParentNode.nodeName() & $sNodePath
			$oXML_Node = $oParentNode
		Else
			$oXML_Node = Null
		EndIf

		$oParentNode = Null
	Until (Not (IsObj($oXML_Node)))

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $sNodePath)
EndFunc   ;==>_XML_GetNodesPathInternal

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetParentNodeName
; Description ...: Gets the parent node name of the node pointed to by the XPath
; Syntax ........: _XML_GetParentNodeName(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success - @TODO
;                  On Failure - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Mike Rerick
; Modified.......:
; Remarks .......: Returns empty string if the XPath is not valid
; Related .......:
; Link ..........:
; Example .......: [yes/no]
; ===============================================================================================================================
Func _XML_GetParentNodeName(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	If _XML_Misc_GetDomVersion() < 4 Then
		Return SetError($XML_ERR_DOMVERSION, 4, $XML_RET_FAILURE) ; TODO @extended Description
	EndIf

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $aResponse[1], $sNodePath, $sNameSpace = ''
	Local $sParentNodeName = "", $oParent_Node = Null

	For $oNode_enum In $oNodes_coll
		Local $oNode_enum1 = $oNode_enum
		$sNodePath = ""
		If $oNode_enum.nodeType <> $XML_NODE_DOCUMENT Then
			$sNameSpace = $oNode_enum.namespaceURI()
			If $sNameSpace = 0 Then
				$sNameSpace = ""
			ElseIf $sNameSpace <> "" Then
				$sNameSpace = StringRight($sNameSpace, StringLen($sNameSpace) - StringInStr($sNameSpace, "/", $STR_NOCASESENSE, -1)) & ":"
			EndIf

			$sNodePath = "/" & $sNameSpace & $oNode_enum.nodeName() & $sNodePath
		EndIf

		$oParent_Node = $oNode_enum1.parentNode()
		If $oParent_Node.nodeType <> $XML_NODE_DOCUMENT Then
			$sNameSpace = $oParent_Node.namespaceURI()
			If $sNameSpace = 0 Then
				$sNameSpace = ""
			ElseIf $sNameSpace <> "" Then
				$sNameSpace &= ":"
			EndIf

			$sNodePath = "/" & $sNameSpace & $oParent_Node.nodeName() & $sNodePath
			$oNode_enum1 = $oParent_Node
			$sParentNodeName = $oParent_Node.nodeName()
		Else
			$oNode_enum1 = 0
		EndIf

		$oParent_Node = Null
		_XML_Array_AddName($aResponse, $sNodePath)
	Next

	$aResponse[0] = UBound($aResponse) - 1
	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $sParentNodeName)
EndFunc   ;==>_XML_GetParentNodeName

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetValue
; Description ...: Get XML values based on XPath input from root node.
; Syntax ........: _XML_GetValue(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
; Return values .: On Success - Returns an array of fields text values (count is in first element)
;                  On Failure - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
;                  |0 - No matching node. ; TODO
;                  |1 - No object passed. ; TODO
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetValue(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $aResponse[1]
	For $oNode_enum In $oNodes_coll
		If $oNode_enum.hasChildNodes() Then
			For $oNode_enum_Child In $oNode_enum.childNodes()
				If $oNode_enum_Child.nodeType = $XML_NODE_CDATA_SECTION Then
					_XML_Array_AddName($aResponse, $oNode_enum_Child.data)
				ElseIf $oNode_enum_Child.nodeType = $XML_NODE_TEXT Then
					_XML_Array_AddName($aResponse, $oNode_enum_Child.Text)
				EndIf
			Next
		Else
			_XML_Array_AddName($aResponse, $oNode_enum.nodeValue)
		EndIf
	Next

	$aResponse[0] = UBound($aResponse) - 1
	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $aResponse)
EndFunc   ;==>_XML_GetValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_InsertChildNode
; Description ...: Insert a child node under the specified XPath Node.
; Syntax ........: _XML_InsertChildNode(ByRef $oXmlDoc, $sXPath, $sNode[, $sData = ""[, $sNameSpace = ""]])
; Parameters ....: $oXmlDoc    - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath     - a string value. The XML tree path from root node (root/child/child..)
;                  $sNode      - Node name to add.
;                  $iItem      - [optional] 0-based child item before which to insert. (Default = 0
;                  $sData      - [optional] Value to give the node
;                  $sNameSpace - [optional] Name Space
; Return values .: On Success  - Returns $XML_RET_SUCCESS
;                  On Failure  - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: GMK
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================
Func _XML_InsertChildNode(ByRef $oXmlDoc, $sXPath, $sNode, $iItem = 0, $sData = "", $sNameSpace = "")
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oChild
	For $oNode_enum In $oNodes_coll
		If $sNameSpace = "" Then
			If Not ($oNode_enum.namespaceURI = 0 Or $oNode_enum.namespaceURI = "") Then $sNameSpace = $oNode_enum.namespaceURI
		EndIf

		$oChild = $oXmlDoc.createNode($XML_NODE_ELEMENT, $sNode, $sNameSpace)
		If @error Then Return SetError($XML_ERR_NODECREATE, @error, $XML_RET_FAILURE)

		If $sData <> "" Then $oChild.text = $sData
		$oNode_enum.insertBefore($oChild, $oNode_enum.childNodes.item($iItem))
		If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)

	Next

	$oNodes_coll = Null
	$oChild = Null
	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_InsertChildNode

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_RemoveAttributeNode
; Description ...: Delete XML Attribute node based on XPath input from root node.
; Syntax ........: _XML_RemoveAttributeNode(ByRef $oXmlDoc, $sXPath, $sAttribute)
; Parameters ....: $oXmlDoc    - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath     - a string value. The XML tree path from root node (root/child/child..)
;                  $sAttribute - a string value. The name of attribute node to delete
; Return values .: On Success  - Returns $XML_RET_SUCCESS
;                  On Failure  - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_RemoveAttributeNode(ByRef $oXmlDoc, $sXPath, $sAttribute)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oAttribute = $oNode_Selected.removeAttributeNode($oNode_Selected.getAttributeNode($sAttribute))
	If @error Then
		Return SetError($XML_ERR_COMERROR, @error, $XML_RET_FAILURE)
	ElseIf Not IsObj($oAttribute) Then
		Return SetError($XML_ERR_NOATTRMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_RemoveAttributeNode

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_ReplaceChild
; Description ...: Replaces selected nodes with another
; Syntax ........: _XML_ReplaceChild(ByRef $oXmlDoc, $sXPath, $sNodeNew_Name[, $sNameSpace = ""])
; Parameters ....: $oXmlDoc       - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath        - a string value. The XML tree path from root node (root/child/child..)
;                  $sNodeNew_Name - a string value. The replacement node name.
;                  $sNameSpace    - [optional] a string value. Default is "".
; Return values .: On Success  - Returns $XML_RET_SUCCESS
;                  On Failure  - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro adapted from http://www.perfectxml.com/msxmlAnswers.asp?Row_ID=65
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........: http://www.perfectxml.com/msxmlAnswers.asp?Row_ID=65
; Example .......; yes
; ===============================================================================================================================
Func _XML_ReplaceChild(ByRef $oXmlDoc, $sXPath, $sNodeNew_Name, $sNameSpace = "")
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oNodeNew = Null
	For $oNode_enum_old In $oNodes_coll
		; Create a New Node element
		$oNodeNew = $oXmlDoc.createNode($XML_NODE_ELEMENT, $sNodeNew_Name, $sNameSpace)
		If @error Then Return SetError($XML_ERR_NODECREATE, @error, $XML_RET_FAILURE)

		; Copy all attributes
		For $oAttributeEnum In $oNode_enum_old.Attributes
			$oNodeNew.Attributes.setNamedItem($oAttributeEnum.cloneNode(True))
		Next

		; Copy all Child Nodes
		For $oNode_enum_old_Child In $oNode_enum_old.childNodes
			$oNodeNew.appendChild($oNode_enum_old_Child)
			If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)
		Next

		; Replace the specified $oNode_enum_old with the supplied $oNodeNew
		$oNode_enum_old.parentNode.replaceChild($oNodeNew, $oNode_enum_old)
		If $oXmlDoc.parseError.errorCode Then
			Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
		EndIf

	Next
	$oNodes_coll = Null
	$oNodeNew = Null

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_ReplaceChild

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_SetAttrib
; Description ...: Set XML Field(s) based on XPath input from root node.
; Syntax ........: _XML_SetAttrib(ByRef $oXmlDoc, $sXPath, $vAttributeNameOrList[, $sValue = ""[, $iIndex = Default]])
; Parameters ....: $oXmlDoc         		- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath          		- a string value. The XML tree path from root node (root/child/child..)
;                  $vAttributeNameOrList	- An array of attributes and values to set, or just one attribute name.
;                  $sValue          		- The value to give the attribute (if $vAttributeNameOrList is a string); defaults to ""
;                  $iIndex          		- Used to specify a specific index for "same named" nodes.
; Return values .: Success          		- An array of fields text values
;                  Failure          		- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok, GMK
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_SetAttrib(ByRef $oXmlDoc, $sXPath, $vAttributeNameOrList, $sValue = "", $iIndex = Default)
	; Local Error handler declaration, it will be automatic CleanUp when returning from function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $aResponse[1]
	Local $iAttributes_count = UBound($vAttributeNameOrList)
	Local $iLastAttribute = $iAttributes_count - 1

	If IsInt($iIndex) And $iIndex > 0 Then
		If $iAttributes_count > 0 Then
			ReDim $aResponse[1][$iAttributes_count]
			For $iAttribute_idx = 0 To $iLastAttribute
				$aResponse[0][$iAttribute_idx] = $oNodes_coll.item($iIndex).SetAttribute($vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLNAME], $vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLVALUE])
				If @error Then Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
			Next
		Else
			$aResponse[0] = $oNodes_coll.item($iIndex).SetAttribute($vAttributeNameOrList, $sValue)
			If @error Then Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
		EndIf
	ElseIf $iIndex = Default Then
		If $iAttributes_count > 0 Then
			ReDim $aResponse[$oNodes_coll.length][$iAttributes_count]
			For $iNode_idx = 0 To $oNodes_coll.length - 1
				For $iAttribute_idx = 0 To $iLastAttribute
					$aResponse[0][$iAttribute_idx] = $oNodes_coll.item($iNode_idx).SetAttribute($vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLNAME], $vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLVALUE])
					If @error Then Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
				Next
			Next
		Else
			ReDim $aResponse[$oNodes_coll.length]
			For $iNode_idx = 0 To $oNodes_coll.length - 1
				$aResponse[$iNode_idx] = $oNodes_coll.item($iNode_idx).SetAttribute($vAttributeNameOrList, $sValue)
				If $oXmlDoc.parseError.errorCode Then
					Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
				EndIf
			Next
		EndIf
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $aResponse)
	EndIf
	Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)

EndFunc   ;==>_XML_SetAttrib

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Transform
; Description ...: Transform XML data
; Syntax ........: _XML_Transform(ByRef $oXmlDoc, $sXSL_FileFullPath)
; Parameters ....: $oXmlDoc           - [in/out] an object. A valid DOMDocument object.
;                  $sXSL_FileFullPath - a string value. The stylesheet to use
; Return values .: On Success         - Returns $XML_RET_SUCCESS
;                  On Failure         - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro , Modified by WeaponX
; Modified ......: mLipok
; Remarks .......: Ref XML Object will be overwriten - will contain Transformed Data
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_Transform(ByRef $oXmlDoc, $sXSL_FileFullPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	If Not (FileExists($sXSL_FileFullPath)) Then
		Return SetError($XML_ERR_LOAD, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oXSLT_Template = ObjCreate("Msxml2.XSLTemplate." & _XML_Misc_GetDomVersion() & ".0")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_XSLTEMPLATE, $XML_RET_FAILURE)

	Local $oXSL_Document = ObjCreate("Msxml2.FreeThreadedDOMDocument." & _XML_Misc_GetDomVersion() & ".0")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_FREETHREADEDDOMDOCUMENT, $XML_RET_FAILURE)

	Local $oXmlDoc_Temp = ObjCreate("Msxml2.DOMDocument." & _XML_Misc_GetDomVersion() & ".0")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_DOMDOCUMENT, $XML_RET_FAILURE)

	$oXSL_Document.async = False
	$oXSL_Document.load($sXSL_FileFullPath)
	If $oXSL_Document.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	$oXSLT_Template.stylesheet = $oXSL_Document
	Local $oXSL_Processor = $oXSLT_Template.createProcessor()
	$oXSL_Processor.input = $oXmlDoc

	$oXmlDoc_Temp.transformNodeToObject($oXSL_Document, $oXmlDoc)
	If $oXSL_Document.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc_Temp.parseError.errorCode, $oXmlDoc_Temp.parseError.reason)
	EndIf

	; Replace oryginal document obecject
	$oXmlDoc = $oXmlDoc_Temp

	; CleanUp
	$oXSL_Processor = Null
	$oXSLT_Template = Null
	$oXSL_Document = Null
	$oXmlDoc_Temp = Null

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_Transform

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_UpdateField
; Description ...: Update existing single node based on XPath specs.
; Syntax ........: _XML_UpdateField(ByRef $oXmlDoc, $sXPath, $sData)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
;                  $sData     - The data to update the node with.
; Return values .: On Success - Returns $XML_RET_SUCCESS
;                  On Failure - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: Weaponx, mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_UpdateField(ByRef $oXmlDoc, $sXPath, $sData)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then
		Return SetError(@error, @extended, $XML_RET_FAILURE)
	ElseIf $oNode_Selected.hasChildNodes Then
		Local $bUpdateStatus = False
		For $oNode_enum_Child In $oNode_Selected.childNodes()
			If $oNode_enum_Child.nodetype = $XML_NODE_TEXT Then
				$oNode_enum_Child.Text = $sData
				$bUpdateStatus = True
				ExitLoop
			EndIf
		Next

		If Not $bUpdateStatus Then
			Local $oNode_Created = $oXmlDoc.createTextNode($sData)
			$oNode_Selected.appendChild($oNode_Created)
			If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)
		EndIf

		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_UpdateField

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_UpdateField2
; Description ...: Update existing node(s) based on XPath specs.
; Syntax ........: _XML_UpdateField2(ByRef $oXmlDoc, $sXPath, $sData)
; Parameters ....: $oXmlDoc   - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath    - a string value. The XML tree path from root node (root/child/child..)
;                  $sData     - The data to update the node with.
; Return values .: On Success - Returns $XML_RET_SUCCESS
;                  On Failure - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified.......: Weaponx, mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: [yes/no]
; ===============================================================================================================================
Func _XML_UpdateField2(ByRef $oXmlDoc, $sXPath, $sData)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	For $oNode_enum In $oNodes_coll
		If $oNode_enum.hasChildNodes() Then
			For $oNode_enum_Child In $oNode_enum.childNodes()
				If $oNode_enum_Child.nodetype = $XML_NODE_TEXT Then
					$oNode_enum_Child.Text = $sData
					ExitLoop
				EndIf
			Next
		Else
			; TODO What here ???
		EndIf
	Next

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_UpdateField2

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Validate_File
; Description ...: Validates a document against a dtd.
; Syntax ........: _XML_Validate_File($sXMLFile, $sNameSpace, $sXSD_FileFullPath)
; Parameters ....: $sXMLFile          - The file to validate
;                  $sNameSpace        - xml namespace
;                  $sXSD_FileFullPath - DTD file to validate against.
; Return values .: On Success         - Returns $XML_RET_SUCCESS
;                  On Failure         - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......: 	; TODO: Add such function but to work with object instead files ( I mean in memory validation )
; Related .......:
; Link ..........; https://msdn.microsoft.com/en-us/library/ms760267(v=vs.85).aspx
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_Validate_File($sXMLFile, $sNameSpace, $sXSD_FileFullPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oXmlDoc = _XML_CreateDOMDocument()
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oXML_SchemaCache = ObjCreate("Msxml2.XMLSchemaCache." & _XML_Misc_GetDomVersion() & ".0")
	If Not IsObj($oXML_SchemaCache) Then
		Return SetError($XML_ERR_GENERAL, $XML_EXT_XMLSCHEMACACHE, $XML_RET_FAILURE)
	EndIf

	$oXML_SchemaCache.add($sNameSpace, $sXSD_FileFullPath)
	$oXmlDoc.schemas = $oXML_SchemaCache
	$oXmlDoc.async = False
	$oXmlDoc.ValidateOnParse = False

	$oXmlDoc.load($sXMLFile)
	$oXmlDoc.validate()

	If $oXmlDoc.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_Validate_File
#EndRegion XML.au3 - Functions - Not yet reviewed

#Region XML.au3 - Functions - Work in progress

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetNodesCount
; Description ...: Get node count based on $sXPath and selected $iNodeType
; Syntax ........: _XML_GetNodesCount(ByRef $oXmlDoc, $sXPath[, $iNodeType = Default])
; Parameters ....: $oXmlDoc				- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath				- a string value. The XML tree path from root node (root/child/child..)
;                  $iNodeType			- [optional] an integer value. Default value is Default which mean any type.
; Return values .: Success				- Number of nodes found (can be 0)
;                  Failure				- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro & DickB
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_GetNodesCount(ByRef $oXmlDoc, $sXPath, $iNodeType = Default)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	If $iNodeType = Default Then
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $oNodes_coll.length)
	ElseIf $iNodeType >= $XML_NODE_ELEMENT And $iNodeType <= $XML_NODE_NOTATION Then
		Local $iNodeCount = 0
		For $oNode_enum In $oNodes_coll
			If $oNode_enum.nodeType = $iNodeType Then $iNodeCount += 1
		Next
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $iNodeCount)
	EndIf

	Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_GetNodesCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_InsertChildWAttr
; Description ...: Inserts a child node(s) under the specified XPath NodeCollection with attributes.
; Syntax ........: _XML_InsertChildWAttr(ByRef $oXmlDoc, $sXPath, $sNodeName[, $aAttributeList = Default[, $sNodeText = ""[, $sNameSpace = ""]]])
; Parameters ....: $oXmlDoc              - [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath               - a string value. The XML tree path from root node (root/child/child..)
;                  $sNodeName            - a string value. The nodeName
;                  $iItem                - [optional] Item before which to insert the child. (Default = 0)
;                  $vAttributeNameOrList - [optional] Attribute name or an array of XML Attributes and Values. (Name|Value)
;                  $sAttribute_Value     - [optional] Attribute value if $vAttributeNameOrList is a string
;                  $sNodeText			 - [optional] a string value. Default is "".
;                  $sNameSpace           - [optional] a string value. Default is "".
; Return values .: Success				 - $XML_RET_SUCCESS
;                  Failure				 - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: GMK
; Modified ......:
; Remarks .......: This function requires that each attribute name has a corresponding value.
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_InsertChildWAttr(ByRef $oXmlDoc, $sXPath, $sNodeName, $iItem = 0, $vAttributeNameOrList = Default, $sAttribute_Value = "", $sNodeText = "", $sNameSpace = "")
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	If (UBound($vAttributeNameOrList) > 0 And UBound($vAttributeNameOrList, $UBOUND_COLUMNS) <> $XMLATTR_COLCOUNTER) Then Return SetError($XML_ERR_ARRAY, $XML_EXT_DEFAULT, $XML_RET_FAILURE)

	Local $iLastAttribute = UBound($vAttributeNameOrList) - 1
	For $iAttribute_idx = 0 To $iLastAttribute
		If _
				$vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLNAME] = '' _
				Or (Not IsString($vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLNAME])) _
				Or (Not IsString($vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLVALUE])) _
				Then
			Return SetError($XML_ERR_ARRAY, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
		EndIf
	Next

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oChild_Temp = Null, $oAttribute_Temp = Null
	For $oNode_enum In $oNodes_coll
		If $sNameSpace = "" Then
			If Not ($oNode_enum.namespaceURI = 0 Or $oNode_enum.namespaceURI = "") Then $sNameSpace = $oNode_enum.namespaceURI
		EndIf

		$oChild_Temp = $oXmlDoc.createNode($XML_NODE_ELEMENT, $sNodeName, $sNameSpace)
		If @error Then Return SetError($XML_ERR_NODECREATE, @error, $XML_RET_FAILURE)

		If $sNodeText <> "" Then $oChild_Temp.text = $sNodeText

		If UBound($vAttributeNameOrList) Then
			For $iAttribute_idx = 0 To UBound($vAttributeNameOrList) - 1
				$oAttribute_Temp = $oXmlDoc.createAttribute($vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLNAME]) ;, $sNameSpace) ; TODO Check this comment
				If @error Then ExitLoop ; TODO Description ?
				$oAttribute_Temp.value = $vAttributeNameOrList[$iAttribute_idx][$XMLATTR_COLVALUE]
				$oChild_Temp.setAttributeNode($oAttribute_Temp)
			Next
		Else
			$oAttribute_Temp = $oXmlDoc.createAttribute($vAttributeNameOrList) ;, $sNameSpace) ; TODO Check this comment
			$oAttribute_Temp.value = $sAttribute_Value
			$oChild_Temp.setAttributeNode($oAttribute_Temp)
		EndIf

		; Inserts a new child node before the given item of the parent node
		$oNode_enum.insertBefore($oChild_Temp, $oNode_enum.childNodes.item($iItem))
		If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)

	Next

	$oChild_Temp = Null
	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_InsertChildWAttr

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_RemoveAttribute
; Description ...: Delete XML Attribute based on XPath input from root node.
; Syntax ........: _XML_RemoveAttribute(ByRef $oXmlDoc, $sXPath, $sAttribute_name)
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath              - a string value. The XML tree path from root node (root/child/child..)
;                  $sAttribute_name     - a string value.
; Return values .: Success        		- $XML_RET_SUCCESS
;                  Failure             	- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: [yes/no]
; ===============================================================================================================================
Func _XML_RemoveAttribute(ByRef $oXmlDoc, $sXPath, $sAttribute_Name)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oAttribute = $oNode_Selected.getAttributeNode($sAttribute_Name)
	If Not (IsObj($oAttribute)) Then
		Return SetError($XML_ERR_NOATTRMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	; https://msdn.microsoft.com/en-us/library/ms757848(v=vs.85).aspx
	Local $iAttributesLength = $oNode_Selected.attributes.length
	Local $oAttribute_Removed = $oNode_Selected.removeAttributeNode($oAttribute)

	If Not IsObj($oAttribute_Removed) Or ($iAttributesLength = 1 + $oNode_Selected.attributes.length) Then
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_RemoveAttribute

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_TransformNode
; Description ...: Process given XML file with a given XSL stylesheet, in order to transform it to given HTML file
; Syntax ........: _XML_TransformNode($sXML_FileFullPath, $sXSL_FileFullPath, $sHTML_FileFullPath[, $iEncoding = $FO_UTF8_NOBOM])
; Parameters ....: $sXML_FileFullPath   - a string value.
;                  $sXSL_FileFullPath   - a string value.
;                  $sHTML_FileFullPath  - a string value.
;                  $iEncoding           - [optional] an integer value. Default is $FO_UTF8_NOBOM.
; Return values .: Success             - 1 and set @error = 0
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _XML_TransformNode($sXML_FileFullPath, $sXSL_FileFullPath, $sHTML_FileFullPath, $iEncoding = $FO_UTF8_NOBOM)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oXSL_Document = ObjCreate("Microsoft.XMLDOM")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_XMLDOM, $XML_RET_FAILURE)

	Local $oXmlDoc = ObjCreate("Microsoft.XMLDOM")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_XMLDOM, $XML_RET_FAILURE)

	$oXSL_Document.Async = False
	$oXSL_Document.Load($sXSL_FileFullPath)
	If $oXSL_Document.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE_XSL, $oXSL_Document.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	$oXmlDoc.Async = False
	$oXmlDoc.Load($sXML_FileFullPath)
	If $oXmlDoc.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	Local $sHTML = $oXmlDoc.transformNode($oXSL_Document)

	Local $hFile = FileOpen($sHTML_FileFullPath, $FO_OVERWRITE + $iEncoding)
	FileWrite($hFile, $sHTML)
	FileClose($hFile)

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_TransformNode
#EndRegion XML.au3 - Functions - Work in progress

#Region XML.au3 - Functions - COMPLETED

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __XML_IsValidObject_Attributes
; Description ...: Check if Object is valid IXMLDOMNamedNodeMap Object
; Syntax ........: __XML_IsValidObject_Attributes(ByRef $oAttributes)
; Parameters ....: $oAttributes                - [in/out] an object.
; Return values .: Success             - $XML_RET_SUCCESS
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __XML_IsValidObject_Attributes(ByRef $oAttributes)
	If Not IsObj($oAttributes) Then
		Return SetError($XML_ERR_ISNOTOBJECT, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf ObjName($oAttributes, $OBJ_NAME) <> 'IXMLDOMNamedNodeMap' Then
		Return SetError($XML_ERR_INVALIDATTRIB, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>__XML_IsValidObject_Attributes

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __XML_IsValidObject_DOMDocument
; Description ...: Check if Object is valid Msxml2.DOMDocument.xxxx Object
; Syntax ........: __XML_IsValidObject_DOMDocument(ByRef $oXML)
; Parameters ....: $oXML                - [in/out] an object.
; Return values .: Success             - $XML_RET_SUCCESS
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __XML_IsValidObject_DOMDocument(ByRef $oXML)
	If Not IsObj($oXML) Then
		Return SetError($XML_ERR_ISNOTOBJECT, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf StringInStr(ObjName($oXML, $OBJ_NAME), 'DOMDocument') = 0 Then
		Return SetError($XML_ERR_INVALIDDOMDOC, $XML_EXT_DOMDOCUMENT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>__XML_IsValidObject_DOMDocument

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __XML_IsValidObject_DOMDocumentOrElement
; Description ...: Check if Object is valid Msxml2.DOMDocument.xxxx Object or IXMLDOMElement
; Syntax ........: __XML_IsValidObject_DOMDocumentOrElement(Byref $oXML)
; Parameters ....: $oXML                - [in/out] an object.
; Return values .: Success             - $XML_RET_SUCCESS
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Return values .: None
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __XML_IsValidObject_DOMDocumentOrElement(ByRef $oXML)
	If Not IsObj($oXML) Then
		Return SetError($XML_ERR_ISNOTOBJECT, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf StringInStr(ObjName($oXML, $OBJ_NAME), 'DOMDocument') = 0 And StringInStr(ObjName($oXML, $OBJ_NAME), 'IXMLDOMElement') = 0 Then
		Return SetError($XML_ERR_INVALIDDOMDOC, $XML_EXT_DOMDOCUMENT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>__XML_IsValidObject_DOMDocumentOrElement

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __XML_IsValidObject_Node
; Description ...: Check if Object is valid IXMLDOMSelection Object
; Syntax ........: __XML_IsValidObject_Node(ByRef $oNode[, $iNodeType = Default])
; Parameters ....: $oNode               - [in/out] an object.
;                  $iNodeType           - [optional] an integer value. Default value is Default this mean any type.
; Return values .: Success             - $XML_RET_SUCCESS
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __XML_IsValidObject_Node(ByRef $oNode, $iNodeType = Default)
	If Not IsObj($oNode) Then
		Return SetError($XML_ERR_ISNOTOBJECT, $XML_EXT_PARAM1, $XML_RET_FAILURE)
	ElseIf ObjName($oNode, $OBJ_NAME) <> 'IXMLDOMNode' And ObjName($oNode, $OBJ_NAME) <> 'IXMLDOMElement' Then
		Return SetError($XML_ERR_INVALIDNODETYPE, $XML_EXT_DEFAULT, ObjName($oNode, $OBJ_NAME))
	ElseIf $iNodeType = Default Then
		; do not check type
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
	ElseIf $iNodeType >= $XML_NODE_ELEMENT And $iNodeType <= $XML_NODE_NOTATION Then
		If $iNodeType = $oNode.type Then
			Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
		EndIf
		Return SetError($XML_ERR_INVALIDNODETYPE, $oNode.type, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM2, $XML_RET_FAILURE)
EndFunc   ;==>__XML_IsValidObject_Node

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __XML_IsValidObject_NodesColl
; Description ...: Check if Object is valid IXMLDOMSelection Object
; Syntax ........: __XML_IsValidObject_NodesColl(ByRef $oNodes_coll)
; Parameters ....: $oNodes_coll                - [in/out] an object.
; Return values .: Success             - $XML_RET_SUCCESS
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __XML_IsValidObject_NodesColl(ByRef $oNodes_coll)
	If Not IsObj($oNodes_coll) Then
		Return SetError($XML_ERR_ISNOTOBJECT, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf ObjName($oNodes_coll, $OBJ_NAME) <> 'IXMLDOMSelection' Then
		Return SetError($XML_ERR_INVALIDNODETYPE, $XML_EXT_DEFAULT, ObjName($oNodes_coll, $OBJ_NAME))
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>__XML_IsValidObject_NodesColl

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_CreateChildWAttr
; Description ...: Create a child node(s) under the specified XPath NodeCollection with attributes.
; Syntax ........: _XML_CreateChildWAttr(ByRef $oXmlDoc, $sXPath, $sNodeName[, $aAttributeList = Default[, $sNodeText = ""[,
;                  $sNameSpace = ""]]])
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath              - a string value. The XML tree path from root node (root/child/child..)
;                  $sNodeName           - a string value. The nodeName
;                  $aAttributeList      - [optional] an array of XML Attributes. Column0=Name, Column1=Value
;                  $sNodeText			- [optional] a string value. Default is "".
;                  $sNameSpace          - [optional] a string value. Default is "".
; Return values .: Success				- $XML_RET_SUCCESS
;                  Failure				- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......: This function requires that each attribute name has a corresponding value.
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_CreateChildWAttr(ByRef $oXmlDoc, $sXPath, $sNodeName, $aAttributeList = Default, $sNodeText = "", $sNameSpace = "")
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	If IsArray($aAttributeList) Then
		If Not (UBound($aAttributeList) > 0 And UBound($aAttributeList, $UBOUND_COLUMNS ) = $XMLATTR_COLCOUNTER) Then
			Return SetError($XML_ERR_ARRAY, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
		EndIf

		For $iAttribute_idx = 0 To UBound($aAttributeList) - 1
			If _
					$aAttributeList[$iAttribute_idx][$XMLATTR_COLNAME] = '' _
					Or (Not IsString($aAttributeList[$iAttribute_idx][$XMLATTR_COLNAME])) _
					Or $aAttributeList[$iAttribute_idx][$XMLATTR_COLVALUE] = '' _ ;  TODO: QUESTION: is Value must be not empty string ?
					Or (Not IsString($aAttributeList[$iAttribute_idx][$XMLATTR_COLVALUE])) _
					Then
				Return SetError($XML_ERR_ARRAY, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
			EndIf
		Next
	ElseIf Not $aAttributeList = Default Then
		Return SetError($XML_ERR_ARRAY, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oChild_Temp = Null, $oAttribute_Temp = Null
	For $oNode_enum In $oNodes_coll
		If $sNameSpace = "" Then
			If Not ($oNode_enum.namespaceURI = 0 Or $oNode_enum.namespaceURI = "") Then $sNameSpace = $oNode_enum.namespaceURI
		EndIf

		$oChild_Temp = $oXmlDoc.createNode($XML_NODE_ELEMENT, $sNodeName, $sNameSpace)
		If @error Then Return SetError($XML_ERR_NODECREATE, @error, $XML_RET_FAILURE)

		If $sNodeText <> "" Then $oChild_Temp.text = $sNodeText

		If UBound($aAttributeList) Then
			For $iAttribute_idx = 0 To UBound($aAttributeList) - 1
				$oAttribute_Temp = $oXmlDoc.createAttribute($aAttributeList[$iAttribute_idx][$XMLATTR_COLNAME]) ;, $sNameSpace) ; TODO Check this comment
				If @error Then ExitLoop ; TODO Description ?

				$oAttribute_Temp.value = $aAttributeList[$iAttribute_idx][$XMLATTR_COLVALUE]
				$oChild_Temp.setAttributeNode($oAttribute_Temp)
			Next
		EndIf

		; Appends a new child node as the last child of the node.
		$oNode_enum.appendChild($oChild_Temp)
		If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)
	Next

	$oChild_Temp = Null
	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_CreateChildWAttr

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_CreateDOMDocument
; Description ...: Create DOMDocument Object
; Syntax ........: _XML_CreateDOMDocument([$iDOM_Version = Default])
; Parameters ....: $iDOM_Version        - [optional] an integer value. Default value is Default.
; Return values .: Success             - DOMDocument Object
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Modified ......: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: yes
; ===============================================================================================================================
Func _XML_CreateDOMDocument($iDOM_Version = Default)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	If $iDOM_Version = Default Then
		For $iDOM_Version_idx = 8 To 0 Step -1
			If FileExists(@SystemDir & "\msxml" & $iDOM_Version_idx & ".dll") Then
				$iDOM_Version = $iDOM_Version_idx
				ExitLoop
			EndIf
		Next
		; if not found $iDOM_Version (still is Default)
		If $iDOM_Version = Default Then
			Return SetError($XML_ERR_OBJCREATE, $XML_EXT_DOMDOCUMENT, $XML_RET_FAILURE)
		EndIf
	Else
		If Not ($iDOM_Version > 0 And $iDOM_Version < 7) Then
			Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
		EndIf
	EndIf

	Local $oXmlDoc = ObjCreate("Msxml2.DOMDocument." & $iDOM_Version & ".0")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_DOMDOCUMENT, $XML_RET_FAILURE)

	__XML_MiscProperty_DomVersion($iDOM_Version)
	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $oXmlDoc)
EndFunc   ;==>_XML_CreateDOMDocument

; #FUNCTION# ===================================================================
; Name ..........: _XML_CreateFile
; Description ...: Create a new blank metafile with header.
; Syntax.........: _XML_CreateFile($sXML_FileFullPath, $sRoot[, $bUTF8 = False[, $iDOM_Version = Default]])
; Parameters ....: $sXML_FileFullPath - The xml filename with full path to create
;                  $sRoot			  - The root of the xml file to create
;                  $bUTF8			  - boolean flag to specify UTF-8 encoding in header.
;                  $iDOM_Version	  - specifically try to use the version supplied here.
; Return values .: On Success		  - Returns $oXmlDoc
;                  On Failure		  - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ==============================================================================
Func _XML_CreateFile($sXML_FileFullPath, $sRoot, $bUTF8 = False, $iDOM_Version = Default)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	; TODO $bUTF8 = False -- change to $iEncoding = $FO_UTF8_NOBOM
	If Not (IsString($sXML_FileFullPath) And $sXML_FileFullPath <> "") Then
		Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM1, $XML_RET_FAILURE)
	ElseIf FileExists($sXML_FileFullPath) Then
		Return SetError($XML_ERR_SAVE, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf Not IsString($sRoot) Then
		Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM2, $XML_RET_FAILURE)
	EndIf

	Local $oXmlDoc = _XML_CreateDOMDocument($iDOM_Version)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oProcessingInstruction = $oXmlDoc.createProcessingInstruction("xml", 'version="1.0"' & (($bUTF8) ? ' encoding="UTF-8"' : ''))
	If @error Then SetError($XML_ERR_COMERROR, @error, $XML_RET_FAILURE)
	If $oXmlDoc.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	$oXmlDoc.appendChild($oProcessingInstruction)
	If @error Then Return SetError($XML_ERR_NODEAPPEND, @error, $XML_RET_FAILURE)

	If $sRoot <> '' Then
		Local $oXML_RootElement = $oXmlDoc.createElement($sRoot)
		$oXmlDoc.documentElement = $oXML_RootElement
	EndIf

	_XML_SaveToFile($oXmlDoc, $sXML_FileFullPath)
	If @error Then _
			Return SetError(@error, @extended, $XML_RET_FAILURE)
	If $oXmlDoc.parseError.errorCode Then _
			Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $oXmlDoc)
EndFunc   ;==>_XML_CreateFile

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetAllAttribIndex
; Description ...: Get XML attributes collection from node element based on Xpath and specific index.
; Syntax ........: _XML_GetAllAttribIndex(ByRef $oXmlDoc, $sXPath[, $iNodeIndex = 0])
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath              - a string value. The XML tree path from root node (root/child/child..)
;                  $iNodeIndex          - [optional] an integer value. Default is 0. Specify which node in collection should be used.
; Return values .: Success        		- Attributes collection of object,
;                  Failure             	- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: yes
; ===============================================================================================================================
Func _XML_GetAllAttribIndex(ByRef $oXmlDoc, $sXPath, $iNodeIndex = 0)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNodes_coll = _XML_SelectNodes($oXmlDoc, $sXPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oAttributes_coll = $oNodes_coll.item($iNodeIndex).attributes
	If @error Then
		Return SetError($XML_ERR_COMERROR, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf IsObj($oAttributes_coll) Then
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $oAttributes_coll)
	ElseIf $oAttributes_coll = Null Then
		Return SetError($XML_ERR_NOATTRMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf $oAttributes_coll.length = 0 Then
		Return SetError($XML_ERR_EMPTYCOLLECTION, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_GetAllAttribIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetChildNodes
; Description ...: Get XML child Node collection of node element based on XPath input from root node.
; Syntax ........: _XML_GetChildNodes(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath              - a string value. The XML tree path from root node (root/child/child..)
; Return values .: Success        		- Child Nodes collection, and @extended contains collection length (count)
;                  Failure             	- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: yes
; ===============================================================================================================================
Func _XML_GetChildNodes(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then
		Return SetError(@error, @extended, $XML_RET_FAILURE)
	ElseIf $oNode_Selected.hasChildNodes() Then
		Return SetError($XML_ERR_SUCCESS, $oNode_Selected.childNodes().length, $oNode_Selected.childNodes())
	EndIf

	Return SetError($XML_ERR_NOCHILDMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_GetChildNodes

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_GetNodeAttributeValue
; Description ...: Get attribute value of selected node
; Syntax ........: _XML_GetNodeAttributeValue(ByRef $oNode_Selected, $sAttribute_Name)
; Parameters ....: $oNode_Selected  - [in/out] A valid node object.
;                  $sAttribute_Name - A string value. Attribute name.
; Return values .: On Success       - Returns string value of selected node
;                  On Failure       - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _XML_GetNodeAttributeValue(ByRef $oNode_Selected, $sAttribute_Name)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_Node($oNode_Selected) ; TODO: , $XML_NODE_ELEMENT)  ??
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	If Not IsString($sAttribute_Name) Or $sAttribute_Name = '' Then
		Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Local $sAttribute_Value = $oNode_Selected.getAttribute($sAttribute_Name)
	If IsString($sAttribute_Value) Then
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $sAttribute_Value)
	ElseIf $sAttribute_Value = Null Then
		Return SetError($XML_ERR_NOATTRMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_GetNodeAttributeValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Load
; Description ...: Load XML file to the existing object or declared empty variable.
; Syntax ........: _XML_Load(Byref $oXmlDoc, $sXML_FileFullPath[, $sNameSpace = ""[, $bValidateOnParse = True[,
;                  $bPreserveWhiteSpace = True]]])
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument object.
;                  $sXML_FileFullPath   - a string value. The XML file to open
;                  $sNameSpace          - [optional] a string value. Default is "". The namespace to specifiy if the file uses one.
;                  $bValidateOnParse    - [optional] a boolean value. Default is True. Validate the document as it is being parsed
;                  $bPreserveWhiteSpace - [optional] a boolean value. Default is True.
; Return values .: Success        		- $oXmlDoc
;                  Failure             	- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: Tom Hohmann, mLipok, GMK
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_Load(ByRef $oXmlDoc, $sXML_FileFullPath, $sNameSpace = "", $bValidateOnParse = True, $bPreserveWhiteSpace = True)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	If Not FileExists($sXML_FileFullPath) Then Return SetError($XML_ERR_LOAD, $XML_EXT_DEFAULT, $XML_RET_FAILURE)

	If _XML_Misc_GetDomVersion() > 4 Then $oXmlDoc.setProperty("ProhibitDTD", False)
	$oXmlDoc.async = False
	$oXmlDoc.preserveWhiteSpace = $bPreserveWhiteSpace
	$oXmlDoc.validateOnParse = $bValidateOnParse
	$oXmlDoc.Load($sXML_FileFullPath)
	If $oXmlDoc.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	; SelectionLanguage do not use this as this cause a problem
	; $oXmlDoc.setProperty("SelectionLanguage", "XPath")

	If $sNameSpace <> "" Then $oXmlDoc.setProperty("SelectionNamespaces", $sNameSpace)

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $oXmlDoc)
EndFunc   ;==>_XML_Load

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_LoadXML
; Description ...: Load XML String to the DOMDocument object.
; Syntax ........: _XML_LoadXML(Byref $oXmlDoc, $sXML_Content[, $sNameSpace = ""[, $bValidateOnParse = True[,
;                  $bPreserveWhiteSpace = True]]])
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument object.
;                  $sXML_Content        - a string value. The XML string to load into the document
;                  $sNameSpace          - [optional] a string value. Default is "". The namespace to specifiy if the file uses one.
;                  $bValidateOnParse    - [optional] a boolean value. Default is True. Set the MSXML ValidateOnParse property
;                  $bPreserveWhiteSpace - [optional] a boolean value. Default is True.
; Return values .: Success				- $oXmlDoc
;                  Failure				- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro, Lukasz Suleja, Tom Hohmann
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_LoadXML(ByRef $oXmlDoc, $sXML_Content, $sNameSpace = "", $bValidateOnParse = True, $bPreserveWhiteSpace = True)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	If _XML_Misc_GetDomVersion() > 4 Then $oXmlDoc.setProperty("ProhibitDTD", False)
	$oXmlDoc.async = False
	$oXmlDoc.preserveWhiteSpace = $bPreserveWhiteSpace
	$oXmlDoc.validateOnParse = $bValidateOnParse
	$oXmlDoc.LoadXml($sXML_Content)
	If $oXmlDoc.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	; SelectionLanguage do not use this as this cause a problem
	; $oXmlDoc.setProperty("SelectionLanguage", "XPath")

	$oXmlDoc.setProperty("SelectionNamespaces", $sNameSpace)
	; TODO this following line was here, actualy I (mLipok) wondering why.
	; "xmlns:ms='urn:schemas-microsoft-com:xslt'"
	; here I put some reference to look in
	; https://msdn.microsoft.com/en-us/library/ms256186(v=vs.110).aspx

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $oXmlDoc)
EndFunc   ;==>_XML_LoadXML

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_NodeExists
; Description ...: Checks for the existence of a node matching the specified path
; Syntax ........: _XML_NodeExists(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc 		- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath          - a string value. The XML tree path from root node (root/child/child..)
; Return values .: Success        	- $XML_RET_SUCCESS
;                  Failure			- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_NodeExists(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	Local $oNode_Selected = _XML_SelectSingleNode($oXmlDoc, $sXPath)
	If @error Then
		Return SetError(@error, @extended, $XML_RET_FAILURE)
	EndIf
	#forceref $oNode_Selected

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_NodeExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_SaveToFile
; Description ...: Save the current $oXmlDoc content to XML file
; Syntax ........: _XML_SaveToFile(ByRef $oXmlDoc, $sXML_FileFullPath)
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument object.
;                  $sXML_FileFullPath   - a string value. The filename to save the $oXmlDoc content.
; Return values .: Success             - $XML_RET_SUCCESS
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_SaveToFile(ByRef $oXmlDoc, $sXML_FileFullPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	If Not (IsString($sXML_FileFullPath) And $sXML_FileFullPath <> '') Then
		Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM2, $XML_RET_FAILURE)
	ElseIf FileExists($sXML_FileFullPath) Then
		Return SetError($XML_ERR_SAVE, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	$oXmlDoc.save($sXML_FileFullPath)
	If $oXmlDoc.parseError.errorCode Then
		Return SetError($XML_ERR_PARSE, $oXmlDoc.parseError.errorCode, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_SaveToFile

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_SelectNodes
; Description ...: Selects XML Node(s) based on XPath input from root node.
; Syntax ........: _XML_SelectNodes(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath              - a string value. The XML tree path from root node (root/child/child..)
; Return values .: Success        		- $oNodes_coll - Nodes collection, and set @extended = $oNodes_coll.length
;                  Failure             	- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_SelectNodes(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocumentOrElement($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oNodes_coll = $oXmlDoc.selectNodes($sXPath)
	If @error Then
		Return SetError($XML_ERR_XPATH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf (Not IsObj($oNodes_coll)) Or $oNodes_coll.length = 0 Then
		Return SetError($XML_ERR_EMPTYCOLLECTION, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $oNodes_coll.length, $oNodes_coll)
EndFunc   ;==>_XML_SelectNodes

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_SelectSingleNode
; Description ...: Select single XML Node based on XPath input from root node.
; Syntax ........: _XML_SelectSingleNode(ByRef $oXmlDoc, $sXPath)
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument or IXMLDOMElement object.
;                  $sXPath              - a string value. The XML tree path from root node (root/child/child..)
; Return values .: Success        		- $oNode_Selected - single node
;                  Failure              - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......; [yes/no]
; ===============================================================================================================================
Func _XML_SelectSingleNode(ByRef $oXmlDoc, $sXPath)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocumentOrElement($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oNode_Selected = $oXmlDoc.selectSingleNode($sXPath)
	If @error Then
		Return SetError($XML_ERR_XPATH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf $oNode_Selected = Null Then
		Return SetError($XML_ERR_NONODESMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	ElseIf Not IsObj($oNode_Selected) Then ; https://www.autoitscript.com/forum/topic/177176-why-isobj-0-and-vargettype-object/
		; $XML_ERR_EMPTYOBJECT
		Return SetError($XML_ERR_NONODESMATCH, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
	EndIf

	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $oNode_Selected)
EndFunc   ;==>_XML_SelectSingleNode

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Tidy
; Description ...: Tidy XML structure
; Syntax ........: _XML_Tidy(ByRef $oXmlDoc[, $sEncoding = Default])
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument object.
;                  $sEncoding           - [optional] a string value. Default value is -1 (omitXMLDeclaration) .
; Return values .: Success				- XML with indent
;                  Failure				- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: mLipok
; Modified ......:
; Remarks .......: set $sEncoding = Default if you want to use default UDF encoding
; Related .......: _XML_MiscProperty_Encoding
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _XML_Tidy(ByRef $oXmlDoc, $sEncoding = -1)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_DOMDocument($oXmlDoc)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	If $sEncoding = Default Then $sEncoding = _XML_MiscProperty_Encoding()
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $oReader = ObjCreate("MSXML2.SAXXMLReader")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_SAXXMLREADER, $XML_RET_FAILURE)

	Local $oWriter = ObjCreate("MSXML2.MXXMLWriter")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_MXXMLWRITER, $XML_RET_FAILURE)

	Local $oStream = ObjCreate("ADODB.STREAM")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_STREAM, $XML_RET_FAILURE)

	Local $sXML_Return = ''

	With $oStream
		.Open
		If $sEncoding <> -1 Then .Charset = $sEncoding
		If @error Then Return SetError($XML_ERR_PARAMETER, $XML_EXT_ENCODING, $XML_RET_FAILURE)

		With $oWriter
			.indent = True
			If $sEncoding = -1 Then
				.omitXMLDeclaration = True
			Else
				.encoding = $sEncoding
			EndIf
			If @error Then Return SetError($XML_ERR_PARAMETER, $XML_EXT_ENCODING, $XML_RET_FAILURE)
			.output = $oStream
		EndWith

		With $oReader
			.contentHandler = $oWriter
			.errorHandler = $oWriter
			.Parse($oXmlDoc)
			If @error Then
				; TODO $XML_ERR_GENERAL replacement ??
				Return SetError($XML_ERR_GENERAL, $XML_EXT_SAXXMLREADER, $XML_RET_FAILURE)
			EndIf
		EndWith

		.Position = 0
		$sXML_Return = .ReadText($ADO_adReadAll)
		If $sXML_Return = Null Then $sXML_Return = ''

	EndWith

	Local $iSizeInBytes = $oStream.size

	; CleanUp
	$oStream = Null
	$oReader = Null
	$oWriter = Null

	; TODO Description for @error and @extended
	Return SetError($XML_ERR_SUCCESS, $iSizeInBytes, $sXML_Return)
EndFunc   ;==>_XML_Tidy
#EndRegion XML.au3 - Functions - COMPLETED

#Region XML.au3 - Functions - Misc
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __XML_MiscProperty_DomVersion
; Description ...: TODO
; Syntax ........: __XML_MiscProperty_DomVersion([$sDomVersion = Default])
; Parameters ....: $sDomVersion         - [optional] a string value. Default value is Default.
; Return values .: DOM Version
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __XML_MiscProperty_DomVersion($sDomVersion = Default)
	Local Static $sDomVersion_Static = -1

	If $sDomVersion = Default Then
		; just return stored static variable
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $sDomVersion_Static)
	ElseIf IsNumber($sDomVersion) Then
		; set and return static variable
		$sDomVersion_Static = $sDomVersion
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $sDomVersion_Static)
	EndIf

	; reset static variable
	$sDomVersion_Static = -1

	; return error as incorrect parameter was passed to this function
	Return SetError($XML_ERR_PARAMETER, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>__XML_MiscProperty_DomVersion

; #FUNCTION# ===================================================================
; Name ..........: _XML_Misc_GetDomVersion
; Description ...: Returns the version of msxml that is in use for the document.
; Syntax.........: _XML_Misc_GetDomVersion()
; Parameters ....: none
; Return values .: Success		- msxml version
;                  Failure		- $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........: Eltorro
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ==============================================================================
Func _XML_Misc_GetDomVersion()
	Return __XML_MiscProperty_DomVersion()
EndFunc   ;==>_XML_Misc_GetDomVersion

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Misc_Viewer
; Description ...: Display XML Document in default system viewer
; Syntax ........: _XML_Misc_Viewer(ByRef $oXmlDoc[, $sXML_FileFullPath = Default])
; Parameters ....: $oXmlDoc           - [in/out] an object. A valid DOMDocument object.
;                  $sXML_FileFullPath - [optional] a string value. Default is Default.
; Return values .: On Success         - Returns $XML_RET_SUCCESS
;                  On Failure         - Returns $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
;                  TODO: @extended description
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _XML_Misc_Viewer(ByRef $oXmlDoc, $sXML_FileFullPath = Default)
	If $sXML_FileFullPath = Default Then
		$sXML_FileFullPath = @ScriptDir & '\XMLDOM_Display_TestingFile.xml'
		FileDelete($sXML_FileFullPath)
	EndIf

	_XML_SaveToFile($oXmlDoc, $sXML_FileFullPath)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	ShellExecute($sXML_FileFullPath)
	Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
EndFunc   ;==>_XML_Misc_Viewer

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_MiscProperty_Encoding
; Description ...: Set/Get Encoding property
; Syntax ........: _XML_MiscProperty_Encoding([$sEncoding = Default])
; Parameters ....: $sEncoding             - [optional] a string value. Default value is Default.
; Return values .: Encoding as string
; Author ........: mLipok
; Modified ......:
; Remarks .......: For a list of the character set names that are known by a system, see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset in the Windows Registry.
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/ms681424(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _XML_MiscProperty_Encoding($sEncoding = Default)
	Local Static $sEncoding_Static = Default

	If $sEncoding = Default Then ; get current value from static variable
		If $sEncoding_Static = Default Then ; if not set yet, try to set default encoding
			Switch @OSLang
				Case 0415 ; 'pl-PL'
					$sEncoding_Static = 'ISO-8859-2'
				Case Else
					$sEncoding_Static = 'ISO-8859-1'
			EndSwitch
		EndIf
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $sEncoding_Static)

	ElseIf IsString($sEncoding) Then ; set and return current value
		$sEncoding_Static = $sEncoding
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $sEncoding_Static)
	EndIf

	Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM1, $sEncoding_Static)
EndFunc   ;==>_XML_MiscProperty_Encoding

; #FUNCTION# ===================================================================
; Name ..........: _XML_MiscProperty_UDFVersion
; Description ...: Returns UDF version number
; Syntax.........: _XML_MiscProperty_UDFVersion()
; Parameters ....: None
; Return values .: The UDF version number
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ==============================================================================
Func _XML_MiscProperty_UDFVersion()
	Return "1.1.1.13"
EndFunc   ;==>_XML_MiscProperty_UDFVersion
#EndRegion XML.au3 - Functions - Misc

#Region XML.au3 - Functions - Arrays

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Array_AddName
; Description ...: Adds an item to an array.
; Syntax.........: _XML_Array_AddName(ByRef $avArray, $sValue)
; Parameters ....: $avArray       - The array to modify.
;                  $sValue        - The value to add to the array.
; Return values .: Success        $XML_RET_SUCCESS and value added to array.
;                  Failure             - $XML_RET_FAILURE and sets the @error flag to non-zero (look in #Region XML.au3 - ERROR Enums)
; Author ........:
; Modified ......: mLipok, guinness
; Remarks .......: Local version of _ArrayAdd to remove dependency on Array.au3
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ==============================================================================
Func _XML_Array_AddName(ByRef $avArray, $sValue)
	Local $iUBound = UBound($avArray)
	If $iUBound Then
		; Cache function call results, as function calls are expensive compared to variable lookups.
		ReDim $avArray[$iUBound + 1]
		$avArray[$iUBound] = $sValue
		Return SetError($XML_ERR_SUCCESS, $XML_EXT_DEFAULT, $XML_RET_SUCCESS)
	EndIf

	Return SetError($XML_ERR_GENERAL, $XML_EXT_DEFAULT, $XML_RET_FAILURE)
EndFunc   ;==>_XML_Array_AddName

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Array_GetAttributesProperties
; Description ...: Get IXMLDOMAttribute Members - properties, and put the result to array
; Syntax ........: _XML_Array_GetAttributesProperties(ByRef $oAttributes_coll)
; Parameters ....: $oAttributes_coll           - [in/out] an object.
; Return values .: TODO
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO @error and description
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/ms767677(v=vs.85).aspx
; Example .......: Yes
; ===============================================================================================================================
Func _XML_Array_GetAttributesProperties(ByRef $oAttributes_coll)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_Attributes($oAttributes_coll)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $avArray[1][$__g_eARRAY_ATTR_ARRAYCOLCOUNT]
	$avArray[0][$__g_eARRAY_ATTR_NAME] = 'name'
	$avArray[0][$__g_eARRAY_ATTR_TYPESTRING] = 'nodeTypeString'
	$avArray[0][$__g_eARRAY_ATTR_VALUE] = 'value'
	$avArray[0][$__g_eARRAY_ATTR_TEXT] = 'text'
	$avArray[0][$__g_eARRAY_ATTR_DATATYPE] = 'dataType'
	$avArray[0][$__g_eARRAY_ATTR_XML] = 'xml'
	Local $iUBound = 0

	For $oAttributeEnum In $oAttributes_coll
		$iUBound = UBound($avArray)
		ReDim $avArray[$iUBound + 1][$__g_eARRAY_ATTR_ARRAYCOLCOUNT]
		$avArray[$iUBound][$__g_eARRAY_ATTR_NAME] = $oAttributeEnum.name
		$avArray[$iUBound][$__g_eARRAY_ATTR_TYPESTRING] = $oAttributeEnum.nodeTypeString
		$avArray[$iUBound][$__g_eARRAY_ATTR_VALUE] = $oAttributeEnum.value
		$avArray[$iUBound][$__g_eARRAY_ATTR_TEXT] = $oAttributeEnum.text
		$avArray[$iUBound][$__g_eARRAY_ATTR_DATATYPE] = $oAttributeEnum.dataType
		$avArray[$iUBound][$__g_eARRAY_ATTR_XML] = $oAttributeEnum.xml
	Next

	Return SetError($XML_ERR_SUCCESS, UBound($avArray), $avArray)
EndFunc   ;==>_XML_Array_GetAttributesProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Array_GetNodesProperties
; Description ...: Get IXMLDOMNode Members - properties, and put the result to array
; Syntax ........: _XML_Array_GetNodesProperties(ByRef $oNodeColl)
; Parameters ....: $oNodeColl           - [in/out] an object.
; Return values .: Array with attributes description
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/ms761386(v=vs.85).aspx
; Example .......: yes
; ===============================================================================================================================
Func _XML_Array_GetNodesProperties(ByRef $oNodes_coll)
	; Error handler, automatic cleanup at end of function
	Local $oXML_COM_ErrorHandler = ObjEvent("AutoIt.Error", __XML_ComErrorHandler_InternalFunction)
	#forceref $oXML_COM_ErrorHandler

	__XML_IsValidObject_NodesColl($oNodes_coll)
	If @error Then Return SetError(@error, @extended, $XML_RET_FAILURE)

	Local $avArray[1][$__g_eARRAY_NODE_ARRAYCOLCOUNT]
	$avArray[0][$__g_eARRAY_NODE_NAME] = 'nodeName'
	$avArray[0][$__g_eARRAY_NODE_TYPESTRING] = 'nodeTypeString'
	$avArray[0][$__g_eARRAY_NODE_VALUE] = 'nodeValue'
	$avArray[0][$__g_eARRAY_NODE_TEXT] = 'text'
	$avArray[0][$__g_eARRAY_NODE_DATATYPE] = 'dataType'
	$avArray[0][$__g_eARRAY_NODE_XML] = 'xml'
	$avArray[0][$__g_eARRAY_NODE_ATTRIBUTES] = 'attributes'
	Local $iUBound = 0

	For $oNode_enum In $oNodes_coll
		$iUBound = UBound($avArray)
		ReDim $avArray[$iUBound + 1][$__g_eARRAY_NODE_ARRAYCOLCOUNT]
		$avArray[$iUBound][$__g_eARRAY_NODE_NAME] = $oNode_enum.nodeName
		$avArray[$iUBound][$__g_eARRAY_NODE_TYPESTRING] = $oNode_enum.nodeTypeString
		$avArray[$iUBound][$__g_eARRAY_NODE_VALUE] = $oNode_enum.nodeValue
		$avArray[$iUBound][$__g_eARRAY_NODE_TEXT] = $oNode_enum.text
		$avArray[$iUBound][$__g_eARRAY_NODE_DATATYPE] = $oNode_enum.dataType
		$avArray[$iUBound][$__g_eARRAY_NODE_XML] = $oNode_enum.xml

		; check if node have any attributes
		If IsObj($oNode_enum.attributes) And $oNode_enum.attributes.length = 0 Then
			$avArray[$iUBound][$__g_eARRAY_NODE_ATTRIBUTES] = ''
		Else
			Local $sAttributes = ''
			For $oAttribute In $oNode_enum.attributes
				$sAttributes &= $oAttribute.name & '="' & $oAttribute.text & '" '
			Next
			$avArray[$iUBound][$__g_eARRAY_NODE_ATTRIBUTES] = StringTrimRight($sAttributes, 1)
		EndIf
	Next

	Return SetError($XML_ERR_SUCCESS, UBound($avArray), $avArray)
EndFunc   ;==>_XML_Array_GetNodesProperties
#EndRegion XML.au3 - Functions - Arrays

#Region XML.au3 - Functions - Error Handling

; #INTERNAL_USE_ONLY#==========================================================
; Name ..........: __XML_ComErrorHandler_InternalFunction
; Description ...: A COM error handling routine.
; Syntax.........: __XML_ComErrorHandler_InternalFunction()
; Parameters ....: None
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; [yes/no]
; ==============================================================================
Func __XML_ComErrorHandler_InternalFunction($oCOMError)
	; If not defined ComErrorHandler_UserFunction then this function do nothing special
	; In that case you only can check @error / @extended after suspect functions

	Local $sUserFunction = _XML_ComErrorHandler_UserFunction()
	If IsFunc($sUserFunction) Then $sUserFunction($oCOMError)
EndFunc   ;==>__XML_ComErrorHandler_InternalFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: __XML_ComErrorHandler_UserFunction
; Description ...: Set a UserFunctionWrapper to move the Fired COM Error Error outside UDF
; Syntax ........: _XML_ComErrorHandler_UserFunction([$fnUserFunction = Default])
; Parameters ....: $fnUserFunction- [optional] a Function. Default value is Default.
; Return values .: ErrorHandler Function
; Author ........: mLipok
; Modified ......:
; Remarks .......: Description TODO
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _XML_ComErrorHandler_UserFunction($fnUserFunction = Default)
	; in case when user do not set his own function UDF must use internal function to avoid AutoItError
	Local Static $fnUserFunction_Static = ''

	If $fnUserFunction = Default Then
		Return $fnUserFunction_Static ; just return stored static variable
	ElseIf IsFunc($fnUserFunction) Then
		$fnUserFunction_Static = $fnUserFunction ; set and return static variable
		Return $fnUserFunction_Static
	EndIf
	$fnUserFunction_Static = '' ; reset static variable

	; return error as incorrect parameter was passed to this function
	Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM1, $fnUserFunction_Static)
EndFunc   ;==>_XML_ComErrorHandler_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_ErrorParser_GetDescription
; Description ...: Parse Error to Console
; Syntax ........: _XML_ErrorParser_GetDescription(ByRef $oXmlDoc)
; Parameters ....: $oXmlDoc 			- [in/out] an object. A valid DOMDocument object.
; Return values .: Descripition for paresed error
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/ms767720(v=vs.85).aspx , https://msdn.microsoft.com/en-us/library/ms757019(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _XML_ErrorParser_GetDescription(ByRef $oXmlDoc)
	Local $sParseError_FullDescription = ''
	If IsObj($oXmlDoc) And $oXmlDoc.parseError.errorCode <> 0 Then
		$sParseError_FullDescription &= 'IXMLDOMParseError errorCode = ' & $oXmlDoc.parseError.errorCode & @CRLF ; Contains the error code of the last parse error.
		$sParseError_FullDescription &= 'IXMLDOMParseError filepos = ' & $oXmlDoc.parseError.filepos & @CRLF ; Contains the absolute file position where the error occurred.
		$sParseError_FullDescription &= 'IXMLDOMParseError line = ' & $oXmlDoc.parseError.line & @CRLF ; Specifies the line number that contains the error.
		$sParseError_FullDescription &= 'IXMLDOMParseError linepos = ' & $oXmlDoc.parseError.linepos & @CRLF ; Contains the character position within the line where the error occurred.
		$sParseError_FullDescription &= 'IXMLDOMParseError reason = ' & $oXmlDoc.parseError.reason & @CRLF ; Describes the reason for the error.
		$sParseError_FullDescription &= 'IXMLDOMParseError srcText = ' & $oXmlDoc.parseError.srcText & @CRLF ; Returns the full text of the line containing the error.
		$sParseError_FullDescription &= 'IXMLDOMParseError url = ' & $oXmlDoc.parseError.url & @CRLF ; Contains the URL of the XML document containing the last error.
	EndIf

	Return $sParseError_FullDescription
EndFunc   ;==>_XML_ErrorParser_GetDescription
#EndRegion XML.au3 - Functions - Error Handling

#Region XML.au3 - NEW TODO

#CS
	DOM Concepts
	https://msdn.microsoft.com/en-us/library/ms764620(v=vs.85).aspx

	DOM Reference
	https://msdn.microsoft.com/en-us/library/ms764730(v=vs.85).aspx

	XML DOM Enumerated Constants
	https://msdn.microsoft.com/en-us/library/ms766473(v=vs.85).aspx

	XML DOM Objects/Interfaces
	https://msdn.microsoft.com/en-us/library/ms760218(v=vs.85).aspx

	XML DOM Methods
	https://msdn.microsoft.com/en-us/library/ms757828(v=vs.85).aspx

	XML DOM Properties
	https://msdn.microsoft.com/en-us/library/ms763798(v=vs.85).aspx

	XML DOM Events
	https://msdn.microsoft.com/en-us/library/ms764697(v=vs.85).aspx


	Working with XML Document Parts
	https://msdn.microsoft.com/en-us/library/ms761381(v=vs.85).aspx


	Understanding XML Namespaces
	https://msdn.microsoft.com/en-us/library/aa468565.aspx

	Managing Namespaces in an XML Document
	https://msdn.microsoft.com/pl-pl/library/d6730bwt(v=vs.110).aspx

	XPath Examples
	https://msdn.microsoft.com/en-us/library/ms256086(v=vs.110).aspx

	XPath Examples
	https://msdn.microsoft.com/pl-pl/library/ms256086(v=vs.120).aspx

	XPath Tutorial (multi lingual: English | česky | Nederlands | Français | Español | По-русски | Deutsch | 中文 | Italiano | Polski )
	http://zvon.org/xxl/XPathTutorial/General/examples.html


	https://social.msdn.microsoft.com/Forums/en-US/b141c07f-4b2c-403a-9a0d-f64b219d316f/prettyprinting-using-mxxmlwriter-encoding-issue?forum=xmlandnetfx

#CE

Func _EncodeXML($sFileToEncode)
	; http://www.vb-helper.com/howto_encode_base64_hex.html
	; https://www.autoitscript.com/forum/topic/138443-image-to-base64-code/?do=findComment&comment=970372
	; http://stackoverflow.com/questions/496751/base64-encode-string-in-vbscript
	; https://support.microsoft.com/en-us/kb/254388
	; https://gist.github.com/wangye/1990522
	; Xroot 2011

	Local $hFile = FileOpen($sFileToEncode, $FO_BINARY)
	Local $dFileContent = FileRead($hFile)
	FileClose($hFile)
	Local $oXmlDoc = ObjCreate("MSXML2.DOMDocument")
	Local $oNode = $oXmlDoc.createElement("b64")
	$oNode.dataType = "bin.base64"
	$oNode.nodeTypedValue = $dFileContent

	Return $oNode.Text
EndFunc   ;==>_EncodeXML


; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Base64Encode
; Description ...: @TODO
; Syntax ........: _XML_Base64Encode(Byref $sData)
; Parameters ....: $sData               - [in/out] a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _XML_Base64Encode(ByRef $sData)
	Local $oXmlDoc = ObjCreate("Msxml2.DOMDocument")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_DEFAULT, $XML_RET_FAILURE)

	Local $oElement = $oXmlDoc.createElement("base64")
	If @error Then Return SetError($XML_ERR_NODECREATE, $XML_EXT_DEFAULT, $XML_RET_FAILURE)

	$oElement.dataType = "bin.base64"
	$oElement.nodeTypedValue = Binary($sData)
	Return $oElement.Text
EndFunc   ;==>_XML_Base64Encode

; #FUNCTION# ====================================================================================================================
; Name ..........: _XML_Base64Decode
; Description ...: @TODO
; Syntax ........: _XML_Base64Decode(Byref $dData[, $iEncoding = $SB_UTF8])
; Parameters ....: $dData               - [in/out] a binary variant value.
;                  $iEncoding           - [optional] an integer value. Default is $SB_UTF8.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _XML_Base64Decode(ByRef $dData, $iEncoding = $SB_UTF8)
	If Not IsBinary($dData) Then Return SetError($XML_ERR_PARAMETER, $XML_EXT_PARAM1, $XML_RET_FAILURE)

	Local $oXmlDoc = ObjCreate("Msxml2.DOMDocument")
	If @error Then Return SetError($XML_ERR_OBJCREATE, $XML_EXT_DEFAULT, $XML_RET_FAILURE)

	Local $oElement = $oXmlDoc.createElement("base64")
	If @error Then Return SetError($XML_ERR_NODECREATE, $XML_EXT_DEFAULT, $XML_RET_FAILURE)

	$oElement.dataType = "bin.base64"
	$oElement.Text = $dData
	Return BinaryToString($oElement.nodeTypedValue, $iEncoding)
EndFunc   ;==>_XML_Base64Decode

#EndRegion XML.au3 - NEW TODO
