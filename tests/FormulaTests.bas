Attribute VB_Name = "FormulaTests"
Option Explicit

Private checks As Long

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2300, "FormulaTests", message
    checks = checks + 1
End Sub

Public Function Run() As String
    Dim wb As Workbook, other As Workbook, fake As FakeConnection, ws As Worksheet
    Dim result As Variant, status As Variant, before As Long, name As Variant, rs As Object
    On Error GoTo Failed
    checks = 0
    TestSQL
    TestResults
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set other = Workbooks.Add
    Configure wb
    Configure other
    Set fake = New FakeConnection
    FormulaEngine.AttachConnection wb, fake
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 42", True, False, 0, 1)
    Check result = 42, "Scalar result"
    Check fake.Executions = 1, "First evaluation executes once"
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 42", True, False, 0, 1)
    Check fake.Executions = 1, "Recalculation reuses result"
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 42", True, False, 1, 1)
    Check fake.Executions = 2, "Refresh token executes again"
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 43", True, False, 1, 1)
    Check fake.Executions = 3, "Changed SQL executes again"
    fake.FailQuery = True
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 43", True, False, 2, 1)
    Check IsError(result), "Failed refresh returns an error"
    status = FormulaEngine.Status(ws.Range("D1"), ws.Range("A1"))
    Check status(1, 2) = "Error", "Status records failure"
    Check Len(status(3, 2)) > 0, "Failed refresh retains last-success time"
    before = fake.Executions
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 43", True, False, 2, 1)
    Check fake.Executions = before, "Recalculation reuses failure"
    fake.FailQuery = False
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 43", True, False, 3, 1)
    Check result = 42, "Changed refresh retries failure"
    result = FormulaEngine.RunQuery(other.Worksheets(1).Range("A1"), "SELECT 42", True, False, 0, 1)
    Check IsError(result), "Other workbook cannot use session"
    Check fake.Executions = before + 1, "Other workbook does not execute SQL"
    wb.Names(sgRangeRole).RefersToRange.Value2 = "CHANGED_ROLE"
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 43", True, False, 4, 1)
    Check IsError(result), "Changed connection settings require reconnect"
    FormulaEngine.AttachConnection wb, fake
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 43", True, False, 4, 1)
    Check result = 42, "Reconnect invalidates cached errors"
    ws.Range("B1").Value2 = "SELECT 42"
    result = FormulaEngine.RunQuery(ws.Range("A2"), ws.Range("B1"), True, False, 0, 1)
    Check ws.Range("B1").Value2 = "SELECT 42", "Reading a range input does not modify it"
    before = fake.Executions
    result = FormulaEngine.RunQuery(ws.Range("A3"), "DELETE FROM t", False, True, 0, 10000)
    Check IsError(result) And fake.Executions = before, "Write SQL rejected before execution"
    result = FormulaEngine.RunQuery(ws.Range("A3"), "SELECT 42", False, True, 0, 1.5)
    Check IsError(result) And fake.Executions = before, "Fractional maxRows rejected"
    result = FormulaEngine.RunQuery(ws.Range("A4"), "SHOW DATABASES", False, True, 0, 10000, True)
    Check UBound(result, 2) = 1 And result(2, 1) = "TEST_DB", "SHOW projection selects database name"
    TestWorksheet wb, other, fake
    TestWorkbookCalculation wb, other
    fake.State = 0
    before = fake.Executions
    result = FormulaEngine.RunQuery(ws.Range("A1"), "SELECT 43", True, False, 4, 1)
    Check IsError(result) And fake.Executions = before, "Closed session cannot return cached success"
    FormulaEngine.ReleaseWorkbook wb
    wb.Close False
    other.Close False
    Run = CStr(checks) & " checks passed (offline Excel/ADO; no Snowflake queries)."
    Exit Function
Failed:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then FormulaEngine.ReleaseWorkbook wb
    If Not other Is Nothing Then FormulaEngine.ReleaseWorkbook other
    If Not wb Is Nothing Then wb.Close False
    If Not other Is Nothing Then other.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 2300, "FormulaTests", message
End Function

Private Sub TestWorkbookCalculation(ByVal wb As Workbook, ByVal other As Workbook)
    Dim first As Worksheet, second As Worksheet, cleanup As Worksheet, oldCalculation As XlCalculation
    Set first = wb.Worksheets(1)
    Set second = wb.Worksheets.Add(After:=first)
    second.Visible = xlSheetHidden
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    Application.Calculation = xlCalculationManual
    first.Range("P1").Value2 = 4
    first.Range("P2").Formula2 = "=P1*2"
    second.Range("A1").Formula2 = "='" & Replace(first.Name, "'", "''") & "'!P1*3"
    first.Range("P1").Value2 = 5
    other.Activate
    Utils.CalculateWorkbook wb
    Check first.Range("P2").Value2 = 10, "Workbook calculation updates visible sheet"
    Check second.Range("A1").Value2 = 15, "Workbook calculation updates hidden sheet with another workbook active"
    Utils.CalculateWorkbook Nothing
    Check True, "Workbook calculation safely ignores a missing workbook"
    Set cleanup = wb.Worksheets.Add(After:=second)
    cleanup.Range("A1").Value2 = "First"
    cleanup.Range("C1").Value2 = "Third"
    Load.DeleteAllEmptyColumns cleanup
    Check cleanup.Range("A1").Value2 = "First" And cleanup.Range("B1").Value2 = "Third", "Legacy empty-column helper preserves data and removes only the blank column"
    Application.Calculation = oldCalculation
    Exit Sub
Failed:
    Dim message As String
    message = Err.Description
    Application.Calculation = oldCalculation
    Err.Raise vbObjectError + 2300, "FormulaTests", message
End Sub

Private Sub Configure(ByVal wb As Workbook)
    Dim i As Long, name As Variant
    For Each name In Array(sgRangeServer, sgRangeUserID, sgRangeAuthType, sgRangeDefaultDatabase, sgRangeDefaultSchema, sgRangeRole, sgRangeWarehouse, sgRangeDateInputFormat, sgRangeTimestampInputFormat, sgRangeTimeInputFormat)
        i = i + 1
        wb.Worksheets(1).Cells(i, 12).Value2 = "TEST_" & i
        wb.Names.Add CStr(name), wb.Worksheets(1).Cells(i, 12)
    Next name
End Sub

Private Sub TestSQL()
    Dim sql As Variant
    For Each sql In Array("SELECT 1", "select ';' AS x; -- trailing comment", "WITH t AS (SELECT 1 AS n) SELECT n FROM t", "SELECT $$a;b$$", "SELECT 'it''s; fine'", "/* comment */ SELECT (1)", "SELECT ""semi;colon"" FROM t")
        Check SQLValidation.ValidateSelect(CStr(sql)) = "", "Expected SELECT accepted: " & sql
    Next sql
    For Each sql In Array("", "DELETE FROM t", "SELECT 1; DROP TABLE t", "WITH t AS (SELECT 1) DELETE FROM t", "SELECT 'unfinished", "SELECT /*unfinished", "SELECT (1", "SELECT 1;;", "SELECT $$unfinished")
        Check SQLValidation.ValidateSelect(CStr(sql)) <> "", "Expected SQL rejected: " & sql
    Next sql
    Check SQLValidation.Identifier("a""b") = """a""""b""", "Identifier escaping"
    Check SQLValidation.Literal("O'Brien") = "'O''Brien'", "Literal escaping"
End Sub

Private Function Rows(ByVal count As Long, Optional ByVal columns As Long = 2) As Object
    Dim rs As Object, i As Long, c As Long
    Set rs = CreateObject("ADODB.Recordset")
    For c = 1 To columns
        rs.Fields.Append "COL" & c, 200, 30, 32
    Next c
    rs.Open
    For i = 1 To count
        rs.AddNew
        For c = 1 To columns
            rs.Fields(c - 1).Value = "r" & i & "c" & c
        Next c
        rs.Update
    Next i
    If count > 0 Then rs.MoveFirst
    Set Rows = rs
End Function

Private Sub ExpectResultError(ByVal rs As Object, ByVal scalar As Boolean, ByVal headers As Boolean, ByVal maxRows As Long, ByVal message As String)
    Dim result As Variant, r As Long, c As Long, number As Long
    On Error Resume Next
    result = FormulaResults.ReadResult(rs, scalar, headers, maxRows, r, c)
    number = Err.Number
    On Error GoTo 0
    rs.Close
    Check number <> 0, message
End Sub

Private Sub TestResults()
    Dim rs As Object, result As Variant, r As Long, c As Long
    Set rs = Rows(2)
    result = FormulaResults.ReadResult(rs, False, True, 2, r, c)
    Check r = 2 And c = 2, "Result dimensions"
    Check LBound(result, 1) = 1 And LBound(result, 2) = 1, "One-based worksheet result"
    Check result(1, 1) = "COL1" And result(2, 2) = "r1c2" And result(3, 1) = "r2c1", "Headers and row orientation"
    rs.Close
    Set rs = Rows(2)
    result = FormulaResults.ReadResult(rs, False, False, 2, r, c)
    Check result(1, 2) = "r1c2" And UBound(result, 1) = 2, "No headers"
    rs.Close
    Set rs = Rows(0)
    result = FormulaResults.ReadResult(rs, False, True, 2, r, c)
    Check UBound(result, 1) = 1 And result(1, 1) = "COL1", "Empty table returns headers"
    rs.Close
    Set rs = Rows(0)
    result = FormulaResults.ReadResult(rs, False, False, 2, r, c)
    Check result = "", "Empty table without headers is blank"
    rs.Close
    Set rs = Rows(0, 1)
    result = FormulaResults.ReadResult(rs, True, False, 1, r, c)
    Check result = "", "Empty scalar is blank"
    rs.Close
    ExpectResultError Rows(2, 1), True, False, 1, "Multiple scalar rows rejected"
    ExpectResultError Rows(1, 2), True, False, 1, "Multiple scalar columns rejected"
    ExpectResultError Rows(3, 2), False, True, 2, "Row limit overflow rejected"
    Set rs = Rows(124999, 2)
    result = FormulaResults.ReadResult(rs, False, True, 200000, r, c)
    Check UBound(result, 1) * UBound(result, 2) = 250000, "Exact cell limit accepted"
    rs.Close
    ExpectResultError Rows(125000, 2), False, True, 200000, "Cell limit includes headers"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Fields.Append "VALUE", 3, , 32
    rs.Open
    rs.AddNew
    rs.Fields(0).Value = 0
    rs.Update
    rs.MoveFirst
    result = FormulaResults.ReadResult(rs, True, False, 1, r, c)
    Check VarType(result) <> vbString And result = 0, "Numeric zero retained"
    rs.Close
    Set rs = Rows(1, 1)
    rs.Fields(0).Value = Null
    rs.Update
    result = FormulaResults.ReadResult(rs, True, False, 1, r, c)
    Check result = "", "SQL NULL is blank"
    rs.Close
    Set rs = Rows(1, 1)
    rs.Fields(0).Value = "00123"
    rs.Update
    result = FormulaResults.ReadResult(rs, True, False, 1, r, c)
    Check result = "00123" And VarType(result) = vbString, "Leading zero identifier retained"
    rs.Close
    Set rs = CreateObject("ADODB.Recordset")
    rs.Fields.Append "AMOUNT", 131
    rs.Fields(0).Precision = 28
    rs.Fields(0).NumericScale = 2
    rs.Open
    rs.AddNew
    rs.Fields(0).Value = CDec("1234567890123456.78")
    rs.Update
    rs.MoveFirst
    result = FormulaResults.ReadResult(rs, True, False, 1, r, c)
    Check VarType(result) = vbString And result = CStr(CDec("1234567890123456.78")), "High precision decimal stays exact text"
    rs.Close
End Sub

Private Sub TestWorksheet(ByVal wb As Workbook, ByVal other As Workbook, ByVal fake As FakeConnection)
    Dim ws As Worksheet, prefix As String, before As Long
    Set ws = wb.Worksheets(1)
    prefix = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    ws.Range("F1").Value2 = "SELECT 42"
    ws.Range("F2").Value2 = 10
    ws.Range("E1").Formula2 = "=" & prefix & "SFQUERY(F1,TRUE,F2)"
    ws.Range("N1").Formula2 = "=SFVALUE(F1,F2)"
    ws.Range("G1").Value2 = "Accountant notes"
    ws.Range("G2").Formula2 = "=10+5"
    other.Activate
    ws.Calculate
    Check ws.Range("E1").Value2 = "VALUE" And ws.Range("E2").Value2 = 42, "Dynamic array works with another workbook active"
    Check ws.Range("N1").Value2 = 42, "Installed add-in functions resolve without file prefixes"
    Check ws.Range("G1").Value2 = "Accountant notes" And ws.Range("G2").Value2 = 15, "Surrounding cells preserved"
    before = fake.Executions
    ws.Calculate
    Check fake.Executions = before, "Worksheet recalculation is cached"
    ws.Range("F2").Value2 = 11
    ws.Calculate
    Check fake.Executions = before + 2, "Worksheet refresh cell refreshes both query formulas"
    ws.Range("H1").Formula2 = "=" & prefix & "SFSTATUS(E1)"
    ws.Calculate
    Check ws.Range("I1").Value2 = "Success", "Dependent worksheet status"
    ws.Range("J2").Value2 = "occupied"
    ws.Range("J1").Formula2 = "=" & prefix & "SFQUERY(F1,TRUE,F2)"
    ws.Calculate
    Check IsError(ws.Range("J1").Value2) And ws.Range("J2").Value2 = "occupied", "Spill collision preserves destination"
End Sub
