
Dim checks, candidate, rs, output, r, c, errNumber
checks = 0

Sub Check(condition, message)
    If Not condition Then Err.Raise vbObjectError + 2300, "Offline tests", message
    checks = checks + 1
End Sub

For Each candidate In Array("SELECT 1", "select ';' AS x; -- tail", "WITH t AS (SELECT 1) SELECT * FROM t", "SELECT $$a;b$$", "SELECT 'it''s; fine'", "SELECT 'back\'slash;still in string'", "/*lead*/ SELECT (1)", "SELECT ""semi;colon"" FROM t")
    Check ValidateSelect(candidate) = "", "Valid query: " & candidate
Next
For Each candidate In Array("", "DROP TABLE t", "SELECT 1; DROP TABLE t", "SELECT 1;;", "WITH t AS (SELECT 1) DELETE FROM t", "SELECT 'unfinished", "SELECT /*unfinished", "SELECT (1", "SELECT $$unfinished", "SELECT 1);", "-- only comment")
    Check ValidateSelect(candidate) <> "", "Rejected query: " & candidate
Next
Check Identifier("a""b") = """a""""b""", "Quoted identifier"
Check Literal("O'Brien") = "'O''Brien'", "Quoted literal"
Check Literal("a\b") = "'a\\b'", "Escaped backslash"

Function MakeRows(count, columns)
    Dim result, i, col
    Set result = CreateObject("ADODB.Recordset")
    For col = 1 To columns
        result.Fields.Append "COL" & col, 200, 30, 32
    Next
    result.Open
    For i = 1 To count
        result.AddNew
        For col = 1 To columns
            result.Fields(col - 1).Value = "r" & i & "c" & col
        Next
        result.Update
    Next
    If count > 0 Then result.MoveFirst
    Set MakeRows = result
End Function

Sub ExpectFailure(recordset, scalar, headers, maxRows, message)
    Dim value, rows, columns, number
    On Error Resume Next
    value = ReadResult(recordset, scalar, headers, maxRows, rows, columns, "")
    number = Err.Number
    On Error GoTo 0
    recordset.Close
    Check number <> 0, message
End Sub

Set rs = MakeRows(2, 2)
output = ReadResult(rs, False, True, 2, r, c, "")
Check r = 2 And c = 2, "Dimensions"
Check output(1, 1) = "COL1" And output(2, 2) = "r1c2" And output(3, 1) = "r2c1", "Row orientation and headers"
rs.Close
Set rs = MakeRows(2, 2)
output = ReadResult(rs, False, False, 2, r, c, "")
Check output(1, 2) = "r1c2" And UBound(output, 1) = 2, "No headers"
rs.Close
Set rs = MakeRows(0, 2)
output = ReadResult(rs, False, True, 2, r, c, "")
Check UBound(output, 1) = 1 And output(1, 1) = "COL1", "Empty table with headers"
rs.Close
Set rs = MakeRows(0, 2)
output = ReadResult(rs, False, False, 2, r, c, "")
Check output = "", "Empty table without headers"
rs.Close
Set rs = MakeRows(0, 1)
output = ReadResult(rs, True, False, 1, r, c, "")
Check output = "", "Empty scalar"
rs.Close
ExpectFailure MakeRows(2, 1), True, False, 1, "Scalar multiple rows"
ExpectFailure MakeRows(1, 2), True, False, 1, "Scalar multiple columns"
ExpectFailure MakeRows(3, 2), False, True, 2, "Row overflow"
Set rs = MakeRows(124999, 2)
output = ReadResult(rs, False, True, 200000, r, c, "")
Check UBound(output, 1) * UBound(output, 2) = 250000, "Exact cell boundary"
rs.Close
ExpectFailure MakeRows(125000, 2), False, True, 200000, "Cell overflow includes headers"
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "VALUE", 3, , 32
rs.Open
rs.AddNew
rs.Fields(0).Value = 0
rs.Update
rs.MoveFirst
output = ReadResult(rs, True, False, 1, r, c, "")
Check VarType(output) <> vbString And output = 0, "Numeric zero"
rs.Close
Set rs = MakeRows(1, 1)
rs.Fields(0).Value = Null
rs.Update
output = ReadResult(rs, True, False, 1, r, c, "")
Check output = "", "NULL"
rs.Close
Set rs = MakeRows(1, 1)
rs.Fields(0).Value = "00123"
rs.Update
output = ReadResult(rs, True, False, 1, r, c, "")
Check output = "00123" And VarType(output) = vbString, "Identifier leading zero"
rs.Close
Set rs = MakeRows(1, 2)
output = ReadResult(rs, False, True, 1, r, c, "COL2")
Check c = 1 And output(2, 1) = "r1c2", "Metadata column projection"
rs.Close
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "AMOUNT", 131
rs.Fields(0).Precision = 28
rs.Fields(0).NumericScale = 2
rs.Open
rs.AddNew
rs.Fields(0).Value = "1234567890123456.78"
rs.Update
rs.MoveFirst
output = ReadResult(rs, True, False, 1, r, c, "")
Check VarType(output) = vbString And output = "1234567890123456.78", "Exact high precision decimal"
rs.Close
WScript.Echo checks & " checks passed (production SQL/result code adapted to WSH; Excel integration requires VBA trust)."
