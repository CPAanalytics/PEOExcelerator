Attribute VB_Name = "FormulaEngine"
Option Explicit

Private mConnection As Object
Private mWorkbook As Workbook
Private mSettings As String
Private mContext As Object
Private mRevision As Long
Private mResults As Object
Private mStates As Object

Private Sub Initialize()
    If mResults Is Nothing Then Set mResults = CreateObject("Scripting.Dictionary")
    If mStates Is Nothing Then Set mStates = CreateObject("Scripting.Dictionary")
End Sub

Public Function WorkbookKey(ByVal wb As Workbook) As String
    WorkbookKey = CStr(ObjPtr(wb))
End Function

Private Function CallerKey(ByVal caller As Range) As String
    CallerKey = WorkbookKey(caller.Worksheet.Parent) & ":" & caller.Worksheet.CodeName & ":" & caller.Cells(1, 1).Address
End Function

Private Function Setting(ByVal wb As Workbook, ByVal name As String) As String
    Dim target As Range
    On Error GoTo Missing
    Set target = wb.Names(name).RefersToRange
    If Not target.Worksheet.Parent Is wb Then GoTo Missing
    If target.CountLarge <> 1 Or IsError(target.Value2) Then GoTo Missing
    Setting = CStr(target.Value2)
    Exit Function
Missing:
    Err.Raise vbObjectError + 2110, , "Connect this workbook using the Snowflake Connect button to initialize its configuration."
End Function

Private Function Pack(ByVal value As String) As String
    Pack = CStr(Len(value)) & ":" & value
End Function

Private Function SettingsKey(ByVal wb As Workbook) As String
    Dim name As Variant
    For Each name In Array(sgRangeServer, sgRangeUserID, sgRangeAuthType, sgRangeDefaultDatabase, sgRangeDefaultSchema, sgRangeRole, sgRangeWarehouse, sgRangeDateInputFormat, sgRangeTimestampInputFormat, sgRangeTimeInputFormat)
        SettingsKey = SettingsKey & Pack(Setting(wb, CStr(name)))
    Next name
End Function

' Called only from interactive connection/session commands, never by a UDF.
Public Sub AttachConnection(ByVal wb As Workbook, ByVal connection As Object)
    Dim rs As Object, names As Variant, i As Long, context As Object, settings As String
    On Error GoTo Failed
    settings = SettingsKey(wb)
    Set context = CreateObject("Scripting.Dictionary")
    Set rs = connection.Execute("SELECT CURRENT_DATABASE(), CURRENT_SCHEMA(), CURRENT_ROLE(), CURRENT_WAREHOUSE(), CURRENT_USER()")
    names = Array("Database", "Schema", "Role", "Warehouse", "User")
    For i = 0 To 4
        If IsNull(rs.Fields(i).Value) Then
            context.Add names(i), ""
        Else
            context.Add names(i), CStr(rs.Fields(i).Value)
        End If
    Next i
    rs.Close
    context.Add "Server", Setting(wb, sgRangeServer)
    Set mWorkbook = wb
    Set mConnection = connection
    Set mContext = context
    mSettings = settings
    Invalidate
    Exit Sub
Failed:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    DetachConnection
End Sub

Public Sub DetachConnection()
    Set mConnection = Nothing
    Set mWorkbook = Nothing
    Set mContext = Nothing
    mSettings = ""
    Invalidate
End Sub

Private Sub Invalidate()
    Initialize
    mRevision = mRevision + 1
    mResults.RemoveAll
    ' Keep last-success timestamps but mark prior attempts as stale.
End Sub

Public Sub ReleaseWorkbook(ByVal wb As Workbook)
    Dim key As Variant, prefix As String
    Initialize
    prefix = WorkbookKey(wb) & ":"
    For Each key In mResults.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then mResults.Remove key
    Next key
    For Each key In mStates.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then mStates.Remove key
    Next key
    If Not mWorkbook Is Nothing Then
        If mWorkbook Is wb Then DetachConnection
    End If
End Sub

Private Function ConnectionProblem(ByVal wb As Workbook) As String
    On Error GoTo Missing
    If mConnection Is Nothing Or mWorkbook Is Nothing Then GoTo Missing
    If Not mWorkbook Is wb Then GoTo Missing
    If mConnection.State <> 1 Then
        ConnectionProblem = "Connection is closed. Connect this workbook again, then change the refresh cell."
    ElseIf SettingsKey(wb) <> mSettings Then
        ConnectionProblem = "Connection settings changed. Connect this workbook again."
    End If
    Exit Function
Missing:
    ConnectionProblem = "Connect this workbook using the Snowflake Connect button, then change the refresh cell."
End Function

Private Function InputValue(ByVal value As Variant) As Variant
    Dim target As Range
    Dim result As Variant
    If IsObject(value) Then
        If Not TypeOf value Is Range Then Err.Raise vbObjectError + 2111, , "Arguments must be values or single cells."
        Set target = value
        If target.CountLarge <> 1 Then Err.Raise vbObjectError + 2111, , "Arguments must reference a single cell."
        result = target.Value2
    Else
        result = value
    End If
    If IsError(result) Or IsArray(result) Or IsNull(result) Then Err.Raise vbObjectError + 2111, , "An input contains an Excel error or is not a single value."
    InputValue = result
End Function

Public Function RunQuery(ByVal caller As Range, ByVal sqlInput As Variant, ByVal scalar As Boolean, ByVal headersInput As Variant, ByVal refreshInput As Variant, ByVal maxRowsInput As Variant, Optional ByVal metadata As Boolean = False) As Variant
    Dim key As String, signature As String, sql As String, refresh As Variant
    Dim maxRows As Long, rowsInput As Variant, headers As Boolean, headerInput As Variant
    Dim problem As String, entry As Object, state As Object, rs As Object
    Dim result As Variant, rowCount As Long, colCount As Long, errorCode As Long
    On Error GoTo Failed
    Initialize
    key = CallerKey(caller)
    sql = CStr(InputValue(sqlInput))
    refresh = InputValue(refreshInput)
    headerInput = InputValue(headersInput)
    If VarType(headerInput) <> vbBoolean Then
        If Not IsNumeric(headerInput) Then Err.Raise vbObjectError + 2111, , "headers must be TRUE or FALSE."
        If headerInput <> 0 And headerInput <> 1 Then Err.Raise vbObjectError + 2111, , "headers must be TRUE or FALSE."
    End If
    headers = CBool(headerInput)
    rowsInput = InputValue(maxRowsInput)
    If Not IsNumeric(rowsInput) Then Err.Raise vbObjectError + 2111, , "maxRows must be a positive whole number."
    If CDbl(rowsInput) < 1 Or CDbl(rowsInput) > 1048575 Then Err.Raise vbObjectError + 2111, , "maxRows must be between 1 and 1,048,575."
    If CDbl(rowsInput) <> Fix(CDbl(rowsInput)) Then Err.Raise vbObjectError + 2111, , "maxRows must be a whole number."
    maxRows = CLng(rowsInput)
    signature = Pack(sql) & Pack(CStr(scalar)) & Pack(CStr(headers)) & Pack(CStr(VarType(refresh)) & ":" & CStr(refresh)) & Pack(CStr(maxRows)) & Pack(CStr(mRevision))
    problem = ConnectionProblem(caller.Worksheet.Parent)
    signature = signature & Pack(problem)
    If mResults.Exists(key) Then
        Set entry = mResults(key)
        If entry("Signature") = signature Then
            RunQuery = entry("Value")
            Exit Function
        End If
    End If
    If Not mStates.Exists(key) Then
        Set state = CreateObject("Scripting.Dictionary")
        state.Add "Last success", ""
        mStates.Add key, state
    Else
        Set state = mStates(key)
    End If
    If Not metadata Then
        problem = SQLValidation.ValidateSelect(sql)
        If problem <> "" Then Err.Raise vbObjectError + 2111, , problem
    End If
    problem = ConnectionProblem(caller.Worksheet.Parent)
    If problem <> "" Then
        errorCode = xlErrNA
        Err.Raise vbObjectError + 2110, , problem
    End If
    Set rs = mConnection.Execute(sql)
    If metadata And sql = "SHOW DATABASES" Then
        result = FormulaResults.ReadResult(rs, scalar, headers, maxRows, rowCount, colCount, "name")
    Else
        result = FormulaResults.ReadResult(rs, scalar, headers, maxRows, rowCount, colCount)
    End If
    rs.Close
    Set rs = Nothing
    state("Last success") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    state("State") = "Success"
    state("Message") = "Query completed. Change the refresh cell to fetch again."
    state("Rows") = rowCount
    state("Columns") = colCount
    SaveResult key, signature, result, state
    RunQuery = result
    Exit Function
Failed:
    problem = Err.Description
    If errorCode = 0 Then errorCode = xlErrValue
    If InStr(1, problem, "expired", vbTextCompare) > 0 Or InStr(1, problem, "authentication", vbTextCompare) > 0 Then
        errorCode = xlErrNA
        problem = "Session expired. Connect this workbook again, then change the refresh cell."
    End If
    result = CVErr(errorCode)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Len(key) > 0 Then
        If state Is Nothing Then
            If mStates.Exists(key) Then
                Set state = mStates(key)
            Else
                Set state = CreateObject("Scripting.Dictionary")
                state.Add "Last success", ""
                mStates.Add key, state
            End If
        End If
        state("State") = "Error"
        state("Message") = problem
        state("Rows") = 0
        state("Columns") = 0
        SaveResult key, signature, result, state
    End If
    RunQuery = result
End Function

Private Sub SaveResult(ByVal key As String, ByVal signature As String, ByVal value As Variant, ByVal state As Object)
    Dim entry As Object
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "Signature", signature
    entry.Add "Value", value
    state("Revision") = mRevision
    If mResults.Exists(key) Then mResults.Remove key
    mResults.Add key, entry
End Sub

Public Sub RecordInputError(ByVal caller As Range, ByVal message As String)
    Dim key As String, state As Object
    Initialize
    key = CallerKey(caller)
    If mStates.Exists(key) Then
        Set state = mStates(key)
    Else
        Set state = CreateObject("Scripting.Dictionary")
        state.Add "Last success", ""
        mStates.Add key, state
    End If
    state("State") = "Error"
    state("Message") = message
    state("Rows") = 0
    state("Columns") = 0
    If mResults.Exists(key) Then mResults.Remove key
    state("Revision") = mRevision
End Sub

Public Function Status(ByVal caller As Range, ByVal queryCell As Range) As Variant
    Dim output(1 To 11, 1 To 2) As Variant, names As Variant, i As Long
    Dim wb As Workbook, problem As String, state As Object, key As String
    Initialize
    Set wb = caller.Worksheet.Parent
    output(1, 1) = "State": output(1, 2) = "Connected"
    output(2, 1) = "Message": output(2, 2) = "Ready for worksheet queries."
    output(3, 1) = "Last success": output(3, 2) = ""
    output(4, 1) = "Rows": output(4, 2) = ""
    output(5, 1) = "Columns": output(5, 2) = ""
    If Not queryCell Is Nothing Then
        If Not queryCell.Worksheet.Parent Is wb Or queryCell.CountLarge <> 1 Then
            output(1, 2) = "Error"
            output(2, 2) = "Reference one query formula anchor in this workbook."
            Status = output
            Exit Function
        End If
        key = CallerKey(queryCell)
        If mStates.Exists(key) Then
            Set state = mStates(key)
            For i = 1 To 5
                output(i, 2) = state(output(i, 1))
            Next i
            If state("Revision") <> mRevision Then
                output(1, 2) = "Refresh needed"
                output(2, 2) = "Connection context changed. Recalculate the query or change its refresh cell."
            End If
        Else
            output(1, 2) = "Not evaluated"
            output(2, 2) = "Recalculate the referenced query formula."
        End If
    End If
    problem = ConnectionProblem(wb)
    If problem <> "" Then
        output(1, 2) = "Not connected"
        output(2, 2) = problem
    End If
    names = Array("Server", "Database", "Schema", "Role", "Warehouse", "User")
    For i = 0 To 5
        output(i + 6, 1) = names(i)
        output(i + 6, 2) = ""
        If problem = "" Then output(i + 6, 2) = mContext(names(i))
    Next i
    Status = output
End Function
