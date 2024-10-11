#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w 7

;~ ADO Enumerated Constants
;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ms678353%28v=vs.85%29.aspx

; ADCPROP_ASYNCTHREADPRIORITY_ENUM
Global Const $ADO_adPriorityLowest = 1
Global Const $ADO_adPriorityBelowNormal = 2
Global Const $ADO_adPriorityNormal = 3
Global Const $ADO_adPriorityAboveNormal = 4
Global Const $ADO_adPriorityHighest = 5

; ADCPROP_AUTORECALC_ENUM
Global Const $ADO_adRecalcUpFront = 0
Global Const $ADO_adRecalcAlways = 1

; ADCPROP_UPDATECRITERIA_ENUM
Global Const $ADO_adCriteriaKey = 0
Global Const $ADO_adCriteriaAllCols = 1
Global Const $ADO_adCriteriaUpdCols = 2
Global Const $ADO_adCriteriaTimeStamp = 3

; ADCPROP_UPDATERESYNC_ENUM
Global Const $ADO_adResyncNone = 0
Global Const $ADO_adResyncAutoIncrement = 1
Global Const $ADO_adResyncConflicts = 2
Global Const $ADO_adResyncUpdates = 4
Global Const $ADO_adResyncInserts = 8
Global Const $ADO_adResyncAll = 15

; AffectEnum
Global Const $ADO_adAffectCurrent = 1
Global Const $ADO_adAffectGroup = 2
Global Const $ADO_adAffectAll = 3
Global Const $ADO_adAffectAllChapters = 4

; BookmarkEnum
Global Const $ADO_adBookmarkCurrent = 0
Global Const $ADO_adBookmarkFirst = 1
Global Const $ADO_adBookmarkLast = 2

; CommandTypeEnum
Global Const $ADO_adCmdUnspecified = -1
Global Const $ADO_adCmdText = 1
Global Const $ADO_adCmdTable = 2
Global Const $ADO_adCmdStoredProc = 4
Global Const $ADO_adCmdUnknown = 8
Global Const $ADO_adCmdFile = 256
Global Const $ADO_adCmdTableDirect = 512

; CompareEnum
Global Const $ADO_adCompareLessThan = 0
Global Const $ADO_adCompareEqual = 1
Global Const $ADO_adCompareGreaterThan = 2
Global Const $ADO_adCompareNotEqual = 3
Global Const $ADO_adCompareNotComparable = 4

; ConnectModeEnum
Global Const $ADO_adModeUnknown = 0
Global Const $ADO_adModeRead = 1
Global Const $ADO_adModeWrite = 2
Global Const $ADO_adModeReadWrite = 3
Global Const $ADO_adModeShareDenyRead = 4
Global Const $ADO_adModeShareDenyWrite = 8
Global Const $ADO_adModeShareExclusive = 12
Global Const $ADO_adModeShareDenyNone = 16
Global Const $ADO_adModeRecursive = 0x400000

; ConnectOptionEnum
Global Const $ADO_adConnectUnspecified = -1
Global Const $ADO_adAsyncConnect = 16

; ConnectPromptEnum
Global Const $ADO_adPromptAlways = 1
Global Const $ADO_adPromptComplete = 2
Global Const $ADO_adPromptCompleteRequired = 3
Global Const $ADO_adPromptNever = 4

; CopyRecordOptionsEnum
Global Const $ADO_adCopyUnspecified = -1
Global Const $ADO_adCopyOverWrite = 1
Global Const $ADO_adCopyNonRecursive = 2
Global Const $ADO_adCopyAllowEmulation = 4

; CursorLocationEnum
Global Const $ADO_adUseNone = 1
Global Const $ADO_adUseServer = 2
Global Const $ADO_adUseClient = 3

; CursorOptionEnum
Global Const $ADO_adAddNew = 0x1000400
Global Const $ADO_adApproxPosition = 0x4000
Global Const $ADO_adBookmark = 0x2000
Global Const $ADO_adDelete = 0x1000800
Global Const $ADO_adFind = 0x80000
Global Const $ADO_adHoldRecords = 0x100
Global Const $ADO_adIndex = 0x100000
Global Const $ADO_adMovePrevious = 0x200
Global Const $ADO_adNotify = 0x40000
Global Const $ADO_adResync = 0x20000
Global Const $ADO_adSeek = 0x200000
Global Const $ADO_adUpdate = 0x1008000
Global Const $ADO_adUpdateBatch = 0x10000

;~ CursorTypeEnum
;~ https://msdn.microsoft.com/en-us/library/windows/desktop/ms681771(v=vs.85).aspx
Global Const $ADO_adOpenUnspecified = -1 ; Does not specify the type of cursor.
Global Const $ADO_adOpenForwardOnly = 0 ; Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
Global Const $ADO_adOpenKeyset = 1 ; Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
Global Const $ADO_adOpenDynamic = 2 ; Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
Global Const $ADO_adOpenStatic = 3 ; Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.

; DataTypeEnum
Global Const $ADO_adArray = 0x2000
Global Const $ADO_adBigInt = 20
Global Const $ADO_adBinary = 128
Global Const $ADO_adBoolean = 11
Global Const $ADO_adBSTR = 8
Global Const $ADO_adChapter = 136
Global Const $ADO_adChar = 129
Global Const $ADO_adCurrency = 6
Global Const $ADO_adDate = 7
Global Const $ADO_adDBDate = 133
Global Const $ADO_adDBTime = 134
Global Const $ADO_adDBTimeStamp = 135
Global Const $ADO_adDecimal = 14
Global Const $ADO_adDouble = 5
Global Const $ADO_adEmpty = 0
Global Const $ADO_adError = 10
Global Const $ADO_adFileTime = 64
Global Const $ADO_adGUID = 72
Global Const $ADO_adIDispatch = 9
Global Const $ADO_adInteger = 3
Global Const $ADO_adIUnknown = 13
Global Const $ADO_adLongVarBinary = 205
Global Const $ADO_adLongVarChar = 201
Global Const $ADO_adLongVarWChar = 203
Global Const $ADO_adNumeric = 131
Global Const $ADO_adPropVariant = 138
Global Const $ADO_adSingle = 4
Global Const $ADO_adSmallInt = 2
Global Const $ADO_adTinyInt = 16
Global Const $ADO_adUnsignedBigInt = 21
Global Const $ADO_adUnsignedInt = 19
Global Const $ADO_adUnsignedSmallInt = 18
Global Const $ADO_adUnsignedTinyInt = 17
Global Const $ADO_adUserDefined = 132
Global Const $ADO_adVarBinary = 204
Global Const $ADO_adVarChar = 200
Global Const $ADO_adVariant = 12
Global Const $ADO_adVarNumeric = 139
Global Const $ADO_adVarWChar = 202
Global Const $ADO_adWChar = 130

; EditModeEnum
Global Const $ADO_adEditNone = 0
Global Const $ADO_adEditInProgress = 1
Global Const $ADO_adEditAdd = 2
Global Const $ADO_adEditDelete = 4

; ErrorValueEnum ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms681549(v=vs.85).aspx

#CS
	Global Const $ADO_adErrBoundToCommand = 3707
	Global Const $ADO_adErrCannotComplete = 3732
	Global Const $ADO_adErrCantChangeConnection = 3748
	Global Const $ADO_adErrCantChangeProvider = 3220
	Global Const $ADO_adErrCantConvertvalue = 3724
	Global Const $ADO_adErrCantCreate = 3725
	Global Const $ADO_adErrCatalogNotSet = 3747
	Global Const $ADO_adErrColumnNotOnThisRow = 3726
	Global Const $ADO_adErrDataConversion = 3421
	Global Const $ADO_adErrDataOverflow = 3721
	Global Const $ADO_adErrDelResOutOfScope = 3738
	Global Const $ADO_adErrDenyNotSupported = 3750
	Global Const $ADO_adErrDenyTypeNotSupported = 3751
	Global Const $ADO_adErrFeatureNotAvailable = 3251
	Global Const $ADO_adErrFieldsUpdateFailed = 3749
	Global Const $ADO_adErrIllegalOperation = 3219
	Global Const $ADO_adErrIntegrityViolation = 3719
	Global Const $ADO_adErrInTransaction = 3246
	Global Const $ADO_adErrInvalidArgument = 3001
	Global Const $ADO_adErrInvalidConnection = 3709
	Global Const $ADO_adErrInvalidParamInfo = 3708
	Global Const $ADO_adErrInvalidTransaction = 3714
	Global Const $ADO_adErrInvalidURL = 3729
	Global Const $ADO_adErrItemNotFound = 3265
	Global Const $ADO_adErrNoCurrentRecord = 3021
	Global Const $ADO_adErrNotExecuting = 3715
	Global Const $ADO_adErrNotReentrant = 3710
	Global Const $ADO_adErrObjectClosed = 3704
	Global Const $ADO_adErrObjectInCollection = 3367
	Global Const $ADO_adErrObjectNotSet = 3420
	Global Const $ADO_adErrObjectOpen = 3705
	Global Const $ADO_adErrOpeningFile = 3002
	Global Const $ADO_adErrOperationCancelled = 3712
	Global Const $ADO_adErrOutOfSpace = 3734
	Global Const $ADO_adErrPermissionDenied = 3720
	Global Const $ADO_adErrProviderFailed = 3000
	Global Const $ADO_adErrProviderNotFound = 3706
	Global Const $ADO_adErrReadFile = 3003
	Global Const $ADO_adErrResourceExists = 3731
	Global Const $ADO_adErrResourceLocked = 3730
	Global Const $ADO_adErrResourceOutOfScope = 3735
	Global Const $ADO_adErrSchemaViolation = 3722
	Global Const $ADO_adErrSignMismatch = 3723
	Global Const $ADO_adErrStillConnecting = 3713
	Global Const $ADO_adErrStillExecuting = 3711
	Global Const $ADO_adErrTreePermissionDenied = 3728
	Global Const $ADO_adErrUnavailable = 3736
	Global Const $ADO_adErrUnsafeOperation = 3716
	Global Const $ADO_adErrURLDoesNotExist = 3727
	Global Const $ADO_adErrURLNamedRowDoesNotExist = 3737
	Global Const $ADO_adErrVolumeNotFound = 3733
	Global Const $ADO_adErrWriteFile = 3004
	Global Const $ADO_adWrnSecurityDialog = 3717
	Global Const $ADO_adWrnSecurityDialogHeader = 3718
#CE
Global Const $ADO_adErrProviderFailed = 3000
Global Const $ADO_adErrInvalidArgument = 3001
Global Const $ADO_adErrOpeningFile = 3002
Global Const $ADO_adErrReadFile = 3003
Global Const $ADO_adErrWriteFile = 3004
Global Const $ADO_adErrNoCurrentRecord = 3021
Global Const $ADO_adErrCantChangeProvider = 3220
Global Const $ADO_adErrInTransaction = 3246
Global Const $ADO_adErrFeatureNotAvailable = 3251
Global Const $ADO_adErrItemNotFound = 3265
Global Const $ADO_adErrObjectInCollection = 3367
Global Const $ADO_adErrObjectNotSet = 3420
Global Const $ADO_adErrDataConversion = 3421
Global Const $ADO_adErrObjectClosed = 3704
Global Const $ADO_adErrObjectOpen = 3705
Global Const $ADO_adErrProviderNotFound = 3706
Global Const $ADO_adErrInvalidParamInfo = 3708
Global Const $ADO_adErrInvalidConnection = 3709
Global Const $ADO_adErrNotReentrant = 3710
Global Const $ADO_adErrStillExecuting = 3711
Global Const $ADO_adErrOperationCancelled = 3712
Global Const $ADO_adErrStillConnecting = 3713
Global Const $ADO_adErrInvalidTransaction = 3714
Global Const $ADO_adErrNotExecuting = 3715
Global Const $ADO_adErrUnsafeOperation = 3716
Global Const $ADO_adWrnSecurityDialog = 3717
Global Const $ADO_adWrnSecurityDialogHeader = 3718
Global Const $ADO_adErrIntegrityViolation = 3719
Global Const $ADO_adErrPermissionDenied = 3720
Global Const $ADO_adErrDataOverflow = 3721
Global Const $ADO_adErrSchemaViolation = 3722
Global Const $ADO_adErrSignMismatch = 3723
Global Const $ADO_adErrCantConvertvalue = 3724
Global Const $ADO_adErrCantCreate = 3725
Global Const $ADO_adErrColumnNotOnThisRow = 3726
Global Const $ADO_adErrURLDoesNotExist = 3727
Global Const $ADO_adErrTreePermissionDenied = 3728
Global Const $ADO_adErrInvalidURL = 3729
Global Const $ADO_adErrResourceLocked = 3730
Global Const $ADO_adErrResourceExists = 3731
Global Const $ADO_adErrCannotComplete = 3732
Global Const $ADO_adErrVolumeNotFound = 3733
Global Const $ADO_adErrOutOfSpace = 3734
Global Const $ADO_adErrResourceOutOfScope = 3735
Global Const $ADO_adErrUnavailable = 3736
Global Const $ADO_adErrURLNamedRowDoesNotExist = 3737
Global Const $ADO_adErrDelResOutOfScope = 3738
Global Const $ADO_adErrCatalogNotSet = 3747
Global Const $ADO_adErrCantChangeConnection = 3748
Global Const $ADO_adErrFieldsUpdateFailed = 3749
Global Const $ADO_adErrDenyTypeNotSupported = 3751

; EventReasonEnum
Global Const $ADO_adRsnAddNew = 1
Global Const $ADO_adRsnDelete = 2
Global Const $ADO_adRsnUpdate = 3
Global Const $ADO_adRsnUndoUpdate = 4
Global Const $ADO_adRsnUndoAddNew = 5
Global Const $ADO_adRsnUndoDelete = 6
Global Const $ADO_adRsnRequery = 7
Global Const $ADO_adRsnResynch = 8
Global Const $ADO_adRsnClose = 9
Global Const $ADO_adRsnMove = 10
Global Const $ADO_adRsnFirstChange = 11
Global Const $ADO_adRsnMoveFirst = 12
Global Const $ADO_adRsnMoveNext = 13
Global Const $ADO_adRsnMovePrevious = 14
Global Const $ADO_adRsnMoveLast = 15

; EventStatusEnum
Global Const $ADO_adStatusOK = 1
Global Const $ADO_adStatusErrorsOccurred = 2
Global Const $ADO_adStatusCantDeny = 3
Global Const $ADO_adStatusCancel = 4
Global Const $ADO_adStatusUnwantedEvent = 5

; ExecuteOptionEnum
Global Const $ADO_adAsyncExecute = 0x10
Global Const $ADO_adAsyncFetch = 0x20
Global Const $ADO_adAsyncFetchNonBlocking = 0x40
Global Const $ADO_adExecuteNoRecords = 0x80
Global Const $ADO_adExecuteStream = 0x400
Global Const $ADO_adExecuteRecord = 2048
Global Const $ADO_adOptionUnspecified = -1

; FieldEnum
Global Const $ADO_adDefaultStream = -1
Global Const $ADO_adRecordURL = -2

; FieldAttributeEnum
Global Const $ADO_adFldCacheDeferred = 0x1000
Global Const $ADO_adFldFixed = 0x10
Global Const $ADO_adFldIsChapter = 0x2000
Global Const $ADO_adFldIsCollection = 0x40000
Global Const $ADO_adFldKeyColumn = 0x8000
Global Const $ADO_adFldIsDefaultStream = 0x20000
Global Const $ADO_adFldIsNullable = 0x20
Global Const $ADO_adFldIsRowURL = 0x10000
Global Const $ADO_adFldLong = 0x80
Global Const $ADO_adFldMayBeNull = 0x40
Global Const $ADO_adFldMayDefer = 0x2
Global Const $ADO_adFldNegativeScalem = 0x4000
Global Const $ADO_adFldRowID = 0x100
Global Const $ADO_adFldRowVersion = 0x200
Global Const $ADO_adFldUnknownUpdatable = 0x8
Global Const $ADO_adFldUnspecified = -1
Global Const $ADO_adFldUpdatable = 0x4

; FieldStatusEnum
Global Const $ADO_adFieldOK = 0
Global Const $ADO_adFieldCantConvertValue = 2
Global Const $ADO_adFieldIsNull = 3
Global Const $ADO_adFieldTruncated = 4
Global Const $ADO_adFieldSignMismatch = 5
Global Const $ADO_adFieldDataOverflow = 6
Global Const $ADO_adFieldCantCreate = 7
Global Const $ADO_adFieldUnavailable = 8
Global Const $ADO_adFieldIntegrityViolation = 10
Global Const $ADO_adFieldSchemaViolation = 11
Global Const $ADO_adFieldBadStatus = 12
Global Const $ADO_adFieldDefault = 13
Global Const $ADO_adFieldIgnore = 15
Global Const $ADO_adFieldDoesNotExist = 16
Global Const $ADO_adFieldInvalidURL = 17
Global Const $ADO_adFieldResourceLocked = 18
Global Const $ADO_adFieldResourceExists = 19
Global Const $ADO_adFieldCannotComplete = 20
Global Const $ADO_adFieldVolumeNotFound = 21
Global Const $ADO_adFieldOutOfSpace = 22
Global Const $ADO_adFieldCannotDeleteSource = 23
Global Const $ADO_adFieldResourceOutOfScope = 25
Global Const $ADO_adFieldAlreadyExists = 26
Global Const $ADO_adFieldPendingChange = 0x40000
Global Const $ADO_adFieldPendingDelete = 0x20000
Global Const $ADO_adFieldPendingInsert = 0x10000
Global Const $ADO_adFieldPendingUnknown = 0x80000
Global Const $ADO_adFieldPendingUnknownDelete = 0x100000
Global Const $ADO_adFieldPermissionDenied = 0x9
Global Const $ADO_adFieldReadOnly = 0x24

; FilterGroupEnum
Global Const $ADO_adFilterNone = 0
Global Const $ADO_adFilterPendingRecords = 1
Global Const $ADO_adFilterAffectedRecords = 2
Global Const $ADO_adFilterFetchedRecords = 3
Global Const $ADO_adFilterConflictingRecords = 5

; GetRowsOptionEnum
Global Const $ADO_adGetRowsRest = -1

; IsolationLevelEnum
Global Const $ADO_adXactUnspecified = -1
Global Const $ADO_adXactChaos = 16
Global Const $ADO_adXactBrowse = 256
Global Const $ADO_adXactReadUncommitted = 256
Global Const $ADO_adXactCursorStability = 4096
Global Const $ADO_adXactReadCommitted = 4096
Global Const $ADO_adXactRepeatableRead = 65536
Global Const $ADO_adXactIsolated = 1048576
Global Const $ADO_adXactSerializable = 1048576

; LineSeparatorsEnum
Global Const $ADO_adCRLF = -1
Global Const $ADO_adLF = 10
Global Const $ADO_adCR = 13

; LockTypeEnum
Global Const $ADO_adLockUnspecified = -1
Global Const $ADO_adLockReadOnly = 1
Global Const $ADO_adLockPessimistic = 2
Global Const $ADO_adLockOptimistic = 3
Global Const $ADO_adLockBatchOptimistic = 4

; MarshalOptionsEnum
Global Const $ADO_adMarshalAll = 0
Global Const $ADO_adMarshalModifiedOnly = 1

; MoveRecordOptionsEnum
Global Const $ADO_adMoveUnspecified = -1
Global Const $ADO_adMoveOverWrite = 1
Global Const $ADO_adMoveDontUpdateLinks = 2
Global Const $ADO_adMoveAllowEmulation = 4

;~ ObjectStateEnum
;~ https://msdn.microsoft.com/en-us/library/windows/desktop/ms675546(v=vs.85).aspx
Global Const $ADO_adStateClosed = 0 ;   The object is closed
Global Const $ADO_adStateOpen = 1 ;   The object is open
Global Const $ADO_adStateConnecting = 2 ;   The object is connecting
Global Const $ADO_adStateExecuting = 4 ;   The object is executing a command
Global Const $ADO_adStateFetching = 8 ;   The rows of the object are being retrieved

; ParameterAttributesEnum
Global Const $ADO_adParamSigned = 16
Global Const $ADO_adParamNullable = 64
Global Const $ADO_adParamLong = 128

; ParameterDirectionEnum
Global Const $ADO_adParamUnknown = 0
Global Const $ADO_adParamInput = 1
Global Const $ADO_adParamOutput = 2
Global Const $ADO_adParamInputOutput = 3
Global Const $ADO_adParamReturnValue = 4

; PersistFormatEnum
Global Const $ADO_adPersistADTG = 0
;!!!!
Global Const $ADO_adPersistADO = 1
;!!!!
Global Const $ADO_adPersistXML = 1
Global Const $ADO_adPersistProviderSpecific = 2

; PositionEnum
Global Const $ADO_adPosEOF = -3
Global Const $ADO_adPosBOF = -2
Global Const $ADO_adPosUnknown = -1

; PropertyAttributesEnum
Global Const $ADO_adPropNotSupported = 0
Global Const $ADO_adPropRequired = 1
Global Const $ADO_adPropOptional = 2
Global Const $ADO_adPropRead = 512
Global Const $ADO_adPropWrite = 1024

; RecordCreateOptionsEnum
Global Const $ADO_adFailIfNotExists = -1
Global Const $ADO_adCreateNonCollection = 0
Global Const $ADO_adCreateCollection = 0x2000
Global Const $ADO_adCreateOverwrite = 0x4000000
Global Const $ADO_adCreateStructDoc = 0x80000000
Global Const $ADO_adOpenIfExists = 0x2000000

; RecordOpenOptionsEnum
Global Const $ADO_adDelayFetchFields = 0x8000
Global Const $ADO_adDelayFetchStream = 0x4000
Global Const $ADO_adOpenAsync = 0x1000
Global Const $ADO_adOpenExecuteCommand = 0x10000
Global Const $ADO_adOpenRecordUnspecified = -1
Global Const $ADO_adOpenOutput = 0x800000

; RecordStatusEnum
Global Const $ADO_adRecCanceled = 0x100
Global Const $ADO_adRecCantRelease = 0x400
Global Const $ADO_adRecConcurrencyViolation = 0x800
Global Const $ADO_adRecDBDeleted = 0x40000
Global Const $ADO_adRecDeleted = 0x4
Global Const $ADO_adRecIntegrityViolation = 0x1000
Global Const $ADO_adRecInvalid = 0x10
Global Const $ADO_adRecMaxChangesExceeded = 0x2000
Global Const $ADO_adRecModified = 0x2
Global Const $ADO_adRecMultipleChanges = 0x40
Global Const $ADO_adRecNew = 0x1
Global Const $ADO_adRecObjectOpen = 0x4000
Global Const $ADO_adRecOK = 0
Global Const $ADO_adRecOutOfMemory = 0x8000
Global Const $ADO_adRecPendingChanges = 0x80
Global Const $ADO_adRecPermissionDenied = 0x10000
Global Const $ADO_adRecSchemaViolation = 0x20000
Global Const $ADO_adRecUnmodified = 0x8

; RecordTypeEnum
Global Const $ADO_adSimpleRecord = 0
Global Const $ADO_adCollectionRecord = 1
Global Const $ADO_adRecordUnknown = -1
Global Const $ADO_adStructDoc = 2

; ResyncEnum
Global Const $ADO_adResyncUnderlyingValues = 1
Global Const $ADO_adResyncAllValues = 2

; SaveOptionsEnum
Global Const $ADO_adSaveCreateNotExist = 1
Global Const $ADO_adSaveCreateOverWrite = 2

; SchemaEnum
; https://msdn.microsoft.com/en-us/library/ms675274(v=vs.85).aspx
Global Const $ADO_adSchemaProviderSpecific = -1
Global Const $ADO_adSchemaAsserts = 0
Global Const $ADO_adSchemaCatalogs = 1
Global Const $ADO_adSchemaCharacterSets = 2
Global Const $ADO_adSchemaCollations = 3
Global Const $ADO_adSchemaCheckConstraints = 5
Global Const $ADO_adSchemaColumns = 4
Global Const $ADO_adSchemaConstraintColumnUsage = 6
Global Const $ADO_adSchemaConstraintTableUsage = 7
Global Const $ADO_adSchemaKeyColumnUsage = 8
Global Const $ADO_adSchemaReferentialConstraints = 9
Global Const $ADO_adSchemaTableConstraints = 10
Global Const $ADO_adSchemaColumnsDomainUsage = 11
Global Const $ADO_adSchemaIndexes = 12
Global Const $ADO_adSchemaColumnPrivileges = 13
Global Const $ADO_adSchemaTablePrivileges = 14
Global Const $ADO_adSchemaUsagePrivileges = 15
Global Const $ADO_adSchemaProcedures = 16
Global Const $ADO_adSchemaSchemata = 17
Global Const $ADO_adSchemaSQLLanguages = 18
Global Const $ADO_adSchemaStatistics = 19
Global Const $ADO_adSchemaTables = 20
Global Const $ADO_adSchemaTranslations = 21
Global Const $ADO_adSchemaProviderTypes = 22
Global Const $ADO_adSchemaViews = 23
Global Const $ADO_adSchemaViewColumnUsage = 24
Global Const $ADO_adSchemaViewTableUsage = 25
Global Const $ADO_adSchemaProcedureParameters = 26
Global Const $ADO_adSchemaForeignKeys = 27
Global Const $ADO_adSchemaPrimaryKeys = 28
Global Const $ADO_adSchemaProcedureColumns = 29
Global Const $ADO_adSchemaDBInfoKeywords = 30
Global Const $ADO_adSchemaDBInfoLiterals = 31
Global Const $ADO_adSchemaCubes = 32
Global Const $ADO_adSchemaDimensions = 33
Global Const $ADO_adSchemaHierarchies = 34
Global Const $ADO_adSchemaLevels = 35
Global Const $ADO_adSchemaMeasures = 36
Global Const $ADO_adSchemaProperties = 37
Global Const $ADO_adSchemaMembers = 38
Global Const $ADO_adSchemaTrustees = 39

; SearchDirectionEnum
; https://msdn.microsoft.com/en-us/library/ms676696(v=vs.85).aspx
Global Const $ADO_adSearchBackward = -1
Global Const $ADO_adSearchForward = 1

; SeekEnum
; https://msdn.microsoft.com/en-us/library/ms681524(v=vs.85).aspx
Global Const $ADO_adSeekFirstEQ = 1
Global Const $ADO_adSeekLastEQ = 2
Global Const $ADO_adSeekAfterEQ = 4
Global Const $ADO_adSeekAfter = 8
Global Const $ADO_adSeekBeforeEQ = 16
Global Const $ADO_adSeekBefore = 32

; StreamOpenOptionsEnum
; https://msdn.microsoft.com/en-us/library/ms676706(v=vs.85).aspx
Global Const $ADO_adOpenStreamUnspecified = -1
Global Const $ADO_adOpenStreamAsync = 1
Global Const $ADO_adOpenStreamFromRecord = 4

; StreamReadEnum
; https://msdn.microsoft.com/en-us/library/ms679794(v=vs.85).aspx
Global Const $ADO_adReadLine = -2
Global Const $ADO_adReadAll = -1

; StreamTypeEnum
; https://msdn.microsoft.com/en-us/library/ms675277(v=vs.85).aspx
Global Const $ADO_adTypeBinary = 1
Global Const $ADO_adTypeText = 2

; StreamWriteEnum
; https://msdn.microsoft.com/en-us/library/ms678072(v=vs.85).aspx
Global Const $ADO_adWriteChar = 0
Global Const $ADO_adWriteLine = 1

; StringFormatEnum
Global Const $ADO_adClipString = 2

; XactAttributeEnum
Global Const $ADO_adXactCommitRetaining = 131072
Global Const $ADO_adXactAbortRetaining = 262144


; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_ERROR_GetErrorInfo
; Description ...:
; Syntax ........: _ADO_ERROR_Description($iError)
; Parameters ....: $iError              - an integer value.
; Return values .: $sDescription
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681549(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_ERROR_GetErrorInfo($iError)
	Local $sDescription = ''
	Switch $iError
		Case $ADO_adErrProviderFailed
			$sDescription = "Provider failed to perform the requested operation."
		Case $ADO_adErrInvalidArgument
			$sDescription = "Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another. This error is often caused by a typographical error in an SQL SELECT statement. For example, a misspelled field name or table name can generate this error. This error can also occur when a field or table named in a SELECT statement does not exist in the data store."
		Case $ADO_adErrOpeningFile
			$sDescription = "File could not be opened. A misspelled file name was specified, or a file has been moved, renamed, or deleted. Over a network, the drive might be temporarily unavailable or network traffic might be preventing a connection."
		Case $ADO_adErrReadFile
			$sDescription = "File could not be read. The name of the file is specified incorrectly, the file might have been moved or deleted, or the file might have become corrupted."
		Case $ADO_adErrWriteFile
			$sDescription = "Write to file failed. You might have closed a file and then tried to write to it, or the file might be corrupted. If the file is located on a network drive, transient network conditions might prevent writing to a network drive."
		Case $ADO_adErrNoCurrentRecord
			$sDescription = "Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record. 3219	adErrIllegalOperation	Operation is not allowed in this context."
		Case $ADO_adErrCantChangeProvider
			$sDescription = "Supplied provider is different from the one already in use."
		Case $ADO_adErrInTransaction
			$sDescription = "Connection object cannot be explicitly closed while in a transaction. A Recordset or Connection object that is currently participating in a transaction cannot be closed. Call either RollbackTrans or CommitTrans before closing the object."
		Case $ADO_adErrFeatureNotAvailable
			$sDescription = "The object or provider is not capable of performing the requested operation. Some operations depend on a particular provider version."
		Case $ADO_adErrItemNotFound
			$sDescription = "Item cannot be found in the collection corresponding to the requested name or ordinal. An incorrect field or table name has been specified."
		Case $ADO_adErrObjectInCollection
			$sDescription = "Object is already in collection. Cannot append. An object cannot be added to the same collection twice."
		Case $ADO_adErrObjectNotSet
			$sDescription = "Object is no longer valid."
		Case $ADO_adErrDataConversion
			$sDescription = "Application uses a value of the wrong type for the current operation. You might have supplied a string to an operation that expects a stream, for example."
		Case $ADO_adErrObjectClosed
			$sDescription = "Operation is not allowed when the object is closed. TheConnection or Recordset has been closed. For example, some other routine might have closed a global object. You can prevent this error by checking the State property before you attempt an operation."
		Case $ADO_adErrObjectOpen
			$sDescription = "Operation is not allowed when the object is open. An object that is open cannot be opened. Fields cannot be appended to an open Recordset."
		Case $ADO_adErrProviderNotFound
			$sDescription = "Provider cannot be found. It may not be properly installed. 3707	adErrBoundToCommand	The ActiveConnection property of a Recordset object, which has a Command object as its source, cannot be changed. The application attempted to assign a newConnection object to a Recordset that has a Commandobject as its source."
		Case $ADO_adErrInvalidParamInfo
			$sDescription = "Parameter object is improperly defined. Inconsistent or incomplete information was provided."
		Case $ADO_adErrInvalidConnection
			$sDescription = "The connection cannot be used to perform this operation. It is either closed or invalid in this context."
		Case $ADO_adErrNotReentrant
			$sDescription = "Operation cannot be performed while processing event. An operation cannot be performed within an event handler that causes the event to fire again. For example, navigation methods should not be called from within aWillMove event handler."
		Case $ADO_adErrStillExecuting
			$sDescription = "Operation cannot be performed while executing asynchronously."
		Case $ADO_adErrOperationCancelled
			$sDescription = "Operation has been canceled by the user. The application has called the CancelUpdate or CancelBatch method and the current operation has been canceled."
		Case $ADO_adErrStillConnecting
			$sDescription = "Operation cannot be performed while connecting asynchronously."
		Case $ADO_adErrInvalidTransaction
			$sDescription = "Coordinating transaction is invalid or has not started."
		Case $ADO_adErrNotExecuting
			$sDescription = "Operation cannot be performed while not executing."
		Case $ADO_adErrUnsafeOperation
			$sDescription = "Safety settings on this computer prohibit accessing a data source on another domain."
		Case $ADO_adWrnSecurityDialog
			$sDescription = "For internal use only. Don't use. (Entry was included for the sake of completeness. This error should not appear in your code.)"
		Case $ADO_adWrnSecurityDialogHeader
			$sDescription = "For internal use only. Don't use. (Entry included for the sake of completeness. This error should not appear in your code.)"
		Case $ADO_adErrIntegrityViolation
			$sDescription = "Data value conflicts with the integrity constraints of the field. A new value for a Field would cause a duplicate key. A value that forms one side of a relationship between two records might not be updatable."
		Case $ADO_adErrPermissionDenied
			$sDescription = "Insufficient permission prevents writing to the field. The user named in the connection string does not have the proper permissions to write to a Field."
		Case $ADO_adErrDataOverflow
			$sDescription = "Data value is too large to be represented by the field data type. A numeric value that is too large for the intended field was assigned. For example, a long integer value was assigned to a short integer field."
		Case $ADO_adErrSchemaViolation
			$sDescription = "Data value conflicts with the data type or constraints of the field. The data store has validation constraints that differ from the Field value."
		Case $ADO_adErrSignMismatch
			$sDescription = "Conversion failed because the data value was signed and the field data type used by the provider was unsigned."
		Case $ADO_adErrCantConvertvalue
			$sDescription = "Data value cannot be converted for reasons other than sign mismatch or data overflow. For example, conversion would have truncated data."
		Case $ADO_adErrCantCreate
			$sDescription = "Data value cannot be set or retrieved because the field data type was unknown, or the provider had insufficient resources to perform the operation."
		Case $ADO_adErrColumnNotOnThisRow
			$sDescription = "Record does not contain this field. An incorrect field name was specified or a field not in the Fields collection of the current record was referenced."
		Case $ADO_adErrURLDoesNotExist
			$sDescription = "Either the source URL or the parent of the destination URL does not exist. There is a typographical error in either the source or destination URL. You might havehttp://mysite/photo/myphoto.jpg when you should actually have http://mysite/photos/myphoto.jpginstead. The typographical error in the parent URL (in this case, photo instead of photos) has caused the error."
		Case $ADO_adErrTreePermissionDenied
			$sDescription = "Permissions are insufficient to access tree or subtree. The user named in the connection string does not have the appropriate permissions."
		Case $ADO_adErrInvalidURL
			$sDescription = "URL contains invalid characters. Make sure the URL is typed correctly. The URL follows the scheme registered to the current provider (for example, Internet Publishing Provider is registered for http)."
		Case $ADO_adErrResourceLocked
			$sDescription = "Object represented by the specified URL is locked by one or more other processes. Wait until the process has finished and attempt the operation again. The object you are trying to access has been locked by another user or by another process in your application. This is most likely to arise in a multi-user environment."
		Case $ADO_adErrResourceExists
			$sDescription = "Copy operation cannot be performed. Object named by destination URL already exists. Specify adCopyOverwriteto replace the object. If you do not specifyadCopyOverwrite when copying the files in a directory, the copy fails when you try to copy an item that already exists in the destination location."
		Case $ADO_adErrCannotComplete
			$sDescription = "The server cannot complete the operation. This might be because the server is busy with other operations or it might be low on resources."
		Case $ADO_adErrVolumeNotFound
			$sDescription = "Provider cannot locate the storage device indicated by the URL. Make sure the URL is typed correctly. The URL of the storage device might be incorrect, but this error can occur for other reasons. The device might be offline or a large volume of network traffic might prevent the connection from being made."
		Case $ADO_adErrOutOfSpace
			$sDescription = "Operation cannot be performed. Provider cannot obtain enough storage space. There might not be enough RAM or hard-drive space for temporary files on the server."
		Case $ADO_adErrResourceOutOfScope
			$sDescription = "Source or destination URL is outside the scope of the current record."
		Case $ADO_adErrUnavailable
			$sDescription = "Operation failed to complete and the status is unavailable. The field may be unavailable or the operation was not attempted. Another user might have changed or deleted the field you are trying to access."
		Case $ADO_adErrURLNamedRowDoesNotExist
			$sDescription = "Record named by this URL does not exist. While attempting to open a file using a Record object, either the file name or the path to the file was misspelled."
		Case $ADO_adErrDelResOutOfScope
			$sDescription = "The URL of the object to be deleted is outside the scope of the current record."
		Case $ADO_adErrCatalogNotSet
			$sDescription = "Operation requires a valid ParentCatalog."
		Case $ADO_adErrCantChangeConnection
			$sDescription = "Connection was denied. The new connection you requested has different characteristics than the one already in use."
		Case $ADO_adErrFieldsUpdateFailed
			$sDescription = "Fields update failed. For further information, examine theStatus property of individual field objects. This error can occur in two situations: when changing a Field object's value in the process of changing or adding a record to the database; and when changing the properties of the Fieldobject itself. 3750	adErrDenyNotSupported	Provider does not support sharing restrictions. An attempt was made to restrict file sharing and your provider does not support the concept."
		Case $ADO_adErrDenyTypeNotSupported
			$sDescription = "Provider does not support the requested kind of sharing restriction. An attempt was made to establish a particular type of file-sharing restriction that is not supported by your provider. See the provider's documentation to determine what file-sharing restrictions are supported."
	EndSwitch
	Return $sDescription
EndFunc   ;==>_ADO_ERROR_GetErrorInfo
