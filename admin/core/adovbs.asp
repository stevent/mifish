<%
'--------------------------------------------------------------------
' Microsoft ADO
'
' (c) 1996 Microsoft Corporation.  All Rights Reserved.
'
'
'
' ADO constants include file for VBScript
'
'--------------------------------------------------------------------

'---- CursorTypeEnum Values ----
CONST adOpenForwardOnly	= 0
CONST adOpenKeyset		= 1
CONST adOpenDynamic		= 2
CONST adOpenStatic		= 3	

'---- CursorOptionEnum Values ----
CONST adHoldRecords		= &H00000100
CONST adMovePrevious		= &H00000200
CONST adAddNew				= &H01000400
CONST adDelete				= &H01000800
CONST adUpdate				= &H01008000
CONST adBookmark			= &H00002000
CONST adApproxPosition	= &H00004000
CONST adUpdateBatch		= &H00010000
CONST adResync				= &H00020000
CONST adNotify				= &H00040000

'---- LockTypeEnum Values ----
CONST adLockReadOnly				= 1
CONST adLockPessimistic			= 2
CONST adLockOptimistic			= 3
CONST adLockBatchOptimistic	= 4

'---- ExecuteOptionEnum Values ----
CONST adRunAsync = &H00000010

'---- ObjectStateEnum Values ----
CONST adStateClosed		= &H00000000
CONST adStateOpen			= &H00000001
CONST adStateConnecting	= &H00000002
CONST adStateExecuting	= &H00000004

'---- ExecuteOptionEnum Values ----
CONST adAsyncExecute				= &H00000010
CONST adAsyncFetch				= &H00000020
CONST adAsyncFetchNonBlocking	= &H00000040
CONST adExecuteNoRecords		= &H00000080
CONST adExecuteStream			= &H00000400

'---- CursorLocationEnum Values ----
CONST adUseServer	= 2
CONST adUseClient	= 3

'---- DataTypeEnum Values ----
'Type											SQL Field						MSAccess Field
CONST adEmpty					= 0
CONST adSmallInt				= 2		'smallint
CONST adInteger				= 3		'int								long integer
CONST adSingle					= 4		'real								single
CONST adDouble					= 5		'float							double
CONST adCurrency				= 6		'money, smallmoney			currency
CONST adDate					= 7		'datetime
CONST adBSTR					= 8
CONST adIDispatch				= 9
CONST adError					= 10
CONST adBoolean				= 11		'bit								boolean
CONST adVariant				= 12
CONST adIUnknown				= 13
CONST adDecimal				= 14
CONST adTinyInt				= 16
CONST adUnsignedTinyInt		= 17		'tinyint							byte
CONST adUnsignedSmallInt	= 18
CONST adUnsignedInt			= 19
CONST adBigInt					= 20		'bigint
CONST adUnsignedBigInt		= 21
CONST adGUID					= 72		'uniqueidentifier
CONST adBinary					= 128		'binary, timestamp
CONST adChar					= 129		'char
CONST adWChar					= 130		'nchar
CONST adNumeric				= 131		'decimal, numeric				decimal
CONST adUserDefined			= 132
CONST adDBDate					= 133
CONST adDBTime					= 134
CONST adDBTimeStamp			= 135		'datetime, smalldatetime
CONST adVarChar				= 200		'varchar
CONST adLongVarChar			= 201		'text
CONST adVarWChar				= 202		'nvarchar						text
CONST adLongVarWChar			= 203		'ntext							memo
CONST adVarBinary				= 204		'sql_variant
CONST adLongVarBinary		= 205		'image							ole object



'---- FieldAttributeEnum Values ----
CONST adFldMayDefer				= &H00000002
CONST adFldUpdatable				= &H00000004
CONST adFldUnknownUpdatable	= &H00000008
CONST adFldFixed					= &H00000010
CONST adFldIsNullable			= &H00000020
CONST adFldMayBeNull				= &H00000040
CONST adFldLong					= &H00000080
CONST adFldRowID					= &H00000100
CONST adFldRowVersion			= &H00000200
CONST adFldCacheDeferred		= &H00001000

'---- EditModeEnum Values ----
CONST adEditNone			= &H0000
CONST adEditInProgress	= &H0001
CONST adEditAdd			= &H0002
CONST adEditDelete		= &H0004

'---- RecordStatusEnum Values ----
CONST adRecOK							= &H0000000
CONST adRecNew							= &H0000001
CONST adRecModified					= &H0000002
CONST adRecDeleted					= &H0000004
CONST adRecUnmodified				= &H0000008
CONST adRecInvalid					= &H0000010
CONST adRecMultipleChanges			= &H0000040
CONST adRecPendingChanges			= &H0000080
CONST adRecCanceled					= &H0000100
CONST adRecCantRelease				= &H0000400
CONST adRecConcurrencyViolation	= &H0000800
CONST adRecIntegrityViolation		= &H0001000
CONST adRecMaxChangesExceeded		= &H0002000
CONST adRecObjectOpen				= &H0004000
CONST adRecOutOfMemory				= &H0008000
CONST adRecPermissionDenied		= &H0010000
CONST adRecSchemaViolation			= &H0020000
CONST adRecDBDeleted					= &H0040000

'---- GetRowsOptionEnum Values ----
CONST adGetRowsRest = -1

'---- PositionEnum Values ----
CONST adPosUnknown	= -1
CONST adPosBOF			= -2
CONST adPosEOF			= -3

'---- enum Values ----
CONST adBookmarkCurrent	= 0
CONST adBookmarkFirst	= 1
CONST adBookmarkLast		= 2

'---- MarshalOptionsEnum Values ----
CONST adMarshalAll				= 0
CONST adMarshalModifiedOnly	= 1

'---- AffectEnum Values ----
CONST adAffectCurrent	= 1
CONST adAffectGroup		= 2
CONST adAffectAll			= 3

'---- FilterGroupEnum Values ----
CONST adFilterNone				= 0
CONST adFilterPendingRecords	= 1
CONST adFilterAffectedRecords	= 2
CONST adFilterFetchedRecords	= 3
CONST adFilterPredicate			= 4

'---- SearchDirection Values ----
CONST adSearchForward	= 1
CONST adSearchBackward	= -1

'---- ConnectPromptEnum Values ----
CONST adPromptAlways					= 1
CONST adPromptComplete				= 2
CONST adPromptCompleteRequired	= 3
CONST adPromptNever					= 4

'---- ConnectModeEnum Values ----
CONST adModeUnknown			= 0
CONST adModeRead				= 1
CONST adModeWrite				= 2
CONST adModeReadWrite		= 3
CONST adModeShareDenyRead	= 4
CONST adModeShareDenyWrite	= 8
CONST adModeShareExclusive	= &Hc
CONST adModeShareDenyNone	= &H10

'---- IsolationLevelEnum Values ----
CONST adXactUnspecified			= &Hffffffff
CONST adXactChaos					= &H00000010
CONST adXactReadUncommitted	= &H00000100
CONST adXactBrowse				= &H00000100
CONST adXactCursorStability	= &H00001000
CONST adXactReadCommitted		= &H00001000
CONST adXactRepeatableRead		= &H00010000
CONST adXactSerializable		= &H00100000
CONST adXactIsolated				= &H00100000

'---- XactAttributeEnum Values ----
CONST adXactCommitRetaining	= &H00020000
CONST adXactAbortRetaining		= &H00040000

'---- PropertyAttributesEnum Values ----
CONST adPropNotSupported	= &H0000
CONST adPropRequired			= &H0001
CONST adPropOptional			= &H0002
CONST adPropRead				= &H0200
CONST adPropWrite				= &H0400

'---- ErrorValueEnum Values ----
CONST adErrInvalidArgument			= &Hbb9
CONST adErrNoCurrentRecord			= &Hbcd
CONST adErrIllegalOperation		= &Hc93
CONST adErrInTransaction			= &Hcae
CONST adErrFeatureNotAvailable	= &Hcb3
CONST adErrItemNotFound				= &Hcc1
CONST adErrObjectInCollection		= &Hd27
CONST adErrObjectNotSet				= &Hd5c
CONST adErrDataConversion			= &Hd5d
CONST adErrObjectClosed				= &He78
CONST adErrObjectOpen				= &He79
CONST adErrProviderNotFound		= &He7a
CONST adErrBoundToCommand			= &He7b
CONST adErrInvalidParamInfo		= &He7c
CONST adErrInvalidConnection		= &He7d
CONST adErrStillExecuting			= &He7f
CONST adErrStillConnecting			= &He81

'---- ParameterAttributesEnum Values ----
CONST adParamSigned		= &H0010
CONST adParamNullable	= &H0040
CONST adParamLong			= &H0080

'---- ParameterDirectionEnum Values ----
CONST adParamUnknown			= &H0000
CONST adParamInput			= &H0001
CONST adParamOutput			= &H0002
CONST adParamInputOutput	= &H0003
CONST adParamReturnValue	= &H0004

'---- CommandTypeEnum Values ----
CONST adCmdUnknown		= &H0008
CONST adCmdText			= &H0001
CONST adCmdTable			= &H0002
CONST adCmdStoredProc	= &H0004

'---- SchemaEnum Values ----
CONST adSchemaProviderSpecific		= -1
CONST adSchemaAsserts					= 0
CONST adSchemaCatalogs					= 1
CONST adSchemaCharacterSets			= 2
CONST adSchemaCollations				= 3
CONST adSchemaColumns					= 4
CONST adSchemaCheckConstraints		= 5
CONST adSchemaConstraintColumnUsage	= 6
CONST adSchemaConstraintTableUsage	= 7
CONST adSchemaKeyColumnUsage			= 8
CONST adSchemaReferentialContraints	= 9
CONST adSchemaTableConstraints		= 10
CONST adSchemaColumnsDomainUsage		= 11
CONST adSchemaIndexes					= 12
CONST adSchemaColumnPrivileges		= 13
CONST adSchemaTablePrivileges			= 14
CONST adSchemaUsagePrivileges			= 15
CONST adSchemaProcedures				= 16
CONST adSchemaSchemata					= 17
CONST adSchemaSQLLanguages				= 18
CONST adSchemaStatistics				= 19
CONST adSchemaTables						= 20
CONST adSchemaTranslations				= 21
CONST adSchemaProviderTypes			= 22
CONST adSchemaViews						= 23
CONST adSchemaViewColumnUsage			= 24
CONST adSchemaViewTableUsage			= 25
CONST adSchemaProcedureParameters	= 26
CONST adSchemaForeignKeys				= 27
CONST adSchemaPrimaryKeys				= 28
CONST adSchemaProcedureColumns		= 29
%>