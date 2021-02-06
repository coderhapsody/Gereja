Attribute VB_Name = "mdlDataAccess"
Option Explicit

Private sSQL As String

Public Function IsBOFEOF(Recordset_Or_ADODC As Object) As Boolean
    On Error Resume Next
    If TypeOf Recordset_Or_ADODC Is ADODB.Recordset Then
        IsBOFEOF = (Recordset_Or_ADODC.BOF And Recordset_Or_ADODC.EOF)
    Else
        IsBOFEOF = (Recordset_Or_ADODC.Recordset.BOF And Recordset_Or_ADODC.Recordset.EOF)
    End If
End Function

Public Function CloneRecordset(DBConnection As ADODB.Connection, ByVal SQL As String, _
    Optional ByVal CursorType As ADODB.CursorTypeEnum = adOpenForwardOnly, _
    Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseClient, _
    Optional ByVal LockingMode As ADODB.LockTypeEnum = adLockReadOnly, _
    Optional ByVal CommandType As CommandTypeEnum = adCmdText, _
    Optional ByVal TimeOutExpiration As Long = 3000) As ADODB.Recordset
    
    Dim rsTemp As ADODB.Recordset
    
    If DBConnection.State <> adStateClosed Then DBConnection.Close
    DBConnection.CommandTimeout = TimeOutExpiration
    DBConnection.Open
    rsTemp.CursorLocation = CursorLocation
    rsTemp.Open SQL, DBConnection, CursorType, LockingMode, CommandType
    
    Set CloneRecordset = rsTemp.Clone
    
    rsTemp.Close
    Set rsTemp = Nothing
End Function

Public Function GetFieldValue(DBConnection As ADODB.Connection, ByVal TableName As String, ByVal FilterField As String, _
  ByVal FilterValue As String, ByVal SelectedField As String, Optional ByVal IsNumericField As Boolean = False)
    GetFieldValue = GetFieldValueByName(DBConnection, TableName, FilterField, FilterValue, SelectedField, IsNumericField)
End Function

Public Function GetFieldValueByName(DBConnection As ADODB.Connection, ByVal TableName As String, ByVal FilterField As String, _
  ByVal FilterValue As String, ByVal SelectedField As String, Optional ByVal IsNumericField As Boolean = False)
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
        
    If IsNumericField Then
        sSQL = "SELECT " & SelectedField & " FROM " & TableName & " WHERE " & FilterField & "=" & FilterValue
    Else
        sSQL = "SELECT " & SelectedField & " FROM " & TableName & " WHERE " & FilterField & "='" & FilterValue & "'"
    End If
    
    rsTemp.Open sSQL, DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not IsBOFEOF(rsTemp) Then
        GetFieldValueByName = GetDefaultFieldValue(rsTemp.Fields(0))
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
End Function

Public Function GetFieldValueByIndex(DBConnection As ADODB.Connection, ByVal TableName As String, ByVal FilterField As String, _
  ByVal FilterValue As String, ByVal SelectedFieldIndex As Integer, Optional ByVal IsNumericField As Boolean = False)
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
        
    rsTemp.Open TableName, DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
    If IsNumericField Then
        rsTemp.Filter = FilterField & "='" & FilterValue & "'"
    Else
        rsTemp.Filter = FilterField & "=" & FilterValue
    End If
    
    If Not IsBOFEOF(rsTemp) Then
        GetFieldValueByIndex = GetDefaultFieldValue(rsTemp.Fields(SelectedFieldIndex))
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
End Function

Public Function GetDefaultFieldValue(DBField As ADODB.Field)
    If Not IsNull(DBField) Then
        GetDefaultFieldValue = DBField.Value
    Else
        Select Case DBField.Type
            Case DataTypeEnum.adBoolean
                GetDefaultFieldValue = False
            Case DataTypeEnum.adDouble, DataTypeEnum.adInteger, DataTypeEnum.adTinyInt, DataTypeEnum.adDecimal, _
                 DataTypeEnum.adNumeric, DataTypeEnum.adUnsignedInt, DataTypeEnum.adUnsignedSmallInt, _
                 DataTypeEnum.adUnsignedTinyInt, DataTypeEnum.adSmallInt, DataTypeEnum.adSingle
                GetDefaultFieldValue = 0
            Case DataTypeEnum.adVarChar, DataTypeEnum.adVarWChar, DataTypeEnum.adLongVarChar, DataTypeEnum.adLongVarWChar
                GetDefaultFieldValue = BLANK
            Case DataTypeEnum.adDate, DataTypeEnum.adDBDate
                GetDefaultFieldValue = Date
        End Select
    End If
End Function

Public Function GetFieldValueMultipleSelection(DBConnection As ADODB.Connection, ByVal TableName As String, _
  ByVal SelectedField As String, ByVal WhereClause As String)
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
        
    sSQL = "SELECT " & SelectedField & " FROM " & TableName & " WHERE " & WhereClause
    rsTemp.Open sSQL, DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not IsBOFEOF(rsTemp) Then
        GetFieldValueMultipleSelection = GetDefaultFieldValue(rsTemp.Fields(SelectedField))
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
End Function

Public Function GetMultipleFieldValue(DBConnection As ADODB.Connection, ByVal TableName As String, _
    ByVal WhereClause As String, SelectedFields As Variant, ArrayResults As Variant) As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim sSelectList As String
    Dim iLoop As Integer
    
    sSelectList = Join(SelectedFields, ",")
    
    sSQL = "SELECT " & sSelectList & " FROM " & TableName & " WHERE " & WhereClause
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        .CursorLocation = adUseClient
        .Open sSQL, DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not .EOF And Not .BOF Then
            ReDim ArrayResults(.RecordCount - 1)
            For iLoop = 0 To UBound(ArrayResults)
                ArrayResults(iLoop) = GetDefaultFieldValue(rsTemp.Fields(iLoop))
            Next
        Else
            GetMultipleFieldValue = False
            Exit Function
        End If
        .Close
    End With
    Set rsTemp = Nothing
    GetMultipleFieldValue = True
End Function

Public Function IsValidFieldValue(DBConnection As ADODB.Connection, ByVal TableName As String, _
    ByVal WhereClause As String) As Boolean
    
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    
    Set rsTemp = New ADODB.Recordset
    sSQL = "SELECT * FROM " & TableName & " WHERE " & WhereClause
    rsTemp.Open sSQL, DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    IsValidFieldValue = Not IsBOFEOF(rsTemp)
End Function

Public Function GetScalarValue(DBConnection As ADODB.Connection, SQL As String) As Double
    Dim rsTemp As ADODB.Recordset
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open SQL, DBConnection, adOpenStatic, adLockReadOnly, adCmdText
    GetScalarValue = GetDefaultFieldValue(rsTemp(0))
End Function

Private Function CompactRepairDatabase(DBConnection As Variant, ByVal DestConnectionString As String) As Boolean
    Dim oJRO As JRO.JetEngine
    Dim DestDBConnection As ADODB.Connection
    
    Set oJRO = New JRO.JetEngine
    
    If IsObject(DBConnection) Then
        If TypeOf DBConnection Is ADODB.Connection Then
            Call oJRO.CompactDatabase(DBConnection.ConnectionString, DestConnectionString)
        End If
    Else
        Call oJRO.CompactDatabase(DBConnection, DestConnectionString)
    End If
    
    Set oJRO = Nothing
End Function

Public Function RefreshDatabaseCache(DBConnection As ADODB.Connection)
    Dim oJRO As JRO.JetEngine
    Set oJRO = New JRO.JetEngine
    
    Call oJRO.RefreshCache(DBConnection)
        
    Set oJRO = Nothing
End Function

Public Sub InitializeDataControl(ADODC As Object, _
                                ByVal ConnectionString As String, _
                                ByVal SQL As String, _
                                ByVal CommandType As CommandTypeEnum, _
                                ByVal LockType As LockTypeEnum, _
                                ByVal CursorType As CursorTypeEnum, _
                                ByVal CursorLocation As CursorLocationEnum)
    With ADODC
        .ConnectionString = ConnectionString
        .CommandType = CommandType
        .LockType = LockType
        .CursorType = CursorType
        .CursorLocation = CursorLocation
        .RecordSource = SQL
        .Refresh
    End With
End Sub

Public Sub RefreshDataControl(ActiveForm As Object)
    Dim Control As Object
        
    For Each Control In ActiveForm.Controls
        If TypeOf Control Is ADODC Then
            Control.Refresh
        End If
    Next
End Sub

Public Sub CompactDatabase(ByVal DatabasePath As String)
On Error Resume Next
    Dim sSourceConnectionString As String, sDestConnectionString As String
    Dim bSuccess As Boolean
    
    sSourceConnectionString = "Provider=Microsoft.JET.OLEDB.4.0; Data source=" & DatabasePath & ";"
    sDestConnectionString = "Provider=Microsoft.JET.OLEDB.4.0; Data source=" & App.Path & "\temp.mdb" & ";"
    bSuccess = CompactRepairDatabase(sSourceConnectionString, sDestConnectionString)
    'If Not bSuccess Then
    '    MsgBox "Failed to compact and repair database.", vbCritical, "Internal Error"
    '    Exit Sub
    'End If
    
    bSuccess = XCopyFile(App.Path & "\temp.mdb", DatabasePath, True)
    If Not bSuccess Then
        MsgBox "Failed to copy file", vbCritical, "Internal Error"
        Exit Sub
    End If
    Kill App.Path & "\temp.mdb"
End Sub
