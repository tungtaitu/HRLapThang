<%
'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3


'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- DataTypeEnum Values ----
Const adSmallInt = 2
Const adInteger = 3
Const adSingle = 4
Const adDouble = 5
Const adDecimal = 14
Const adNumeric = 131
Const adDate = 7
Const adDBDate = 133
Const adChar = 129
Const adVarChar = 200

'---- CommandTypeEnum Values ----
Const adCmdUnknown = 0
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004

'---- FilterGroupEnum Values ----
Const adFilterNone = 0
Const adFilterPendingRecords = 1
Const adFilterAffectedRecords = 2
Const adFilterFetchedRecords = 3
Const adFilterPredicate = 4
Const adFilterConflictingRecords = 5

'---- ParameterAttributesEnum Values ----
Const adParamSigned = &H0010
Const adParamNullable = &H0040
Const adParamLong = &H0080

'---- ParameterDirectionEnum Values ----
Const adParamUnknown = &H0000
Const adParamInput = &H0001
Const adParamOutput = &H0002
Const adParamInputOutput = &H0003
Const adParamReturnValue = &H0004
%>
