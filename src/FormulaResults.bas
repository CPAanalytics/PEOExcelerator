Attribute VB_Name = "FormulaResults"
Option Explicit

Public Const MaxResultCells As Long = 250000

' Fetch one extra row to detect overflow. Never return a partial schedule.
Public Function ReadResult(ByVal rs As Object, ByVal scalar As Boolean, ByVal headers As Boolean, ByVal maxRows As Long, ByRef rowCount As Long, ByRef colCount As Long, Optional ByVal onlyColumn As String = "") As Variant
    Dim raw As Variant, output() As Variant, r As Long, c As Long
    Dim allowance As Long, headerRows As Long, fetchRows As Long
    Dim sourceColumn As Long, selectedColumn As Long
    colCount = rs.Fields.Count
    If onlyColumn <> "" Then
        selectedColumn = -1
        For c = 0 To colCount - 1
            If StrComp(rs.Fields(c).Name, onlyColumn, vbTextCompare) = 0 Then selectedColumn = c
        Next c
        If selectedColumn = -1 Then Err.Raise vbObjectError + 2102, , "Database metadata did not include a name column."
        colCount = 1
    End If
    If colCount = 0 Then Err.Raise vbObjectError + 2102, , "The query did not return columns."
    If scalar And colCount <> 1 Then Err.Raise vbObjectError + 2102, , "SFVALUE requires exactly one column and at most one row. Use SFQUERY for tables."
    If headers And Not scalar Then headerRows = 1
    If colCount > MaxResultCells Then GoTo TooLarge
    allowance = MaxResultCells \ colCount - headerRows
    If scalar Then allowance = 1
    If allowance > maxRows Then allowance = maxRows
    fetchRows = allowance + 1
    rowCount = 0
    If Not rs.EOF Then
        raw = rs.GetRows(fetchRows)
        rowCount = UBound(raw, 2) + 1
    End If
    If rowCount > allowance Then
        If scalar Then Err.Raise vbObjectError + 2102, , "SFVALUE requires at most one row. Use SFQUERY or aggregate the query."
        GoTo TooLarge
    End If
    If scalar Then
        If rowCount = 0 Then
            ReadResult = ""
        Else
            ReadResult = CellValue(raw(0, 0), rs.Fields(0))
        End If
        Exit Function
    End If
    If rowCount = 0 And headerRows = 0 Then
        ReadResult = ""
        Exit Function
    End If
    ReDim output(1 To rowCount + headerRows, 1 To colCount)
    For c = 0 To colCount - 1
        sourceColumn = c
        If onlyColumn <> "" Then sourceColumn = selectedColumn
        If headers Then output(1, c + 1) = rs.Fields(sourceColumn).Name
        For r = 0 To rowCount - 1
            output(r + headerRows + 1, c + 1) = CellValue(raw(sourceColumn, r), rs.Fields(sourceColumn))
        Next r
    Next c
    ReadResult = output
    Exit Function
TooLarge:
    Err.Raise vbObjectError + 2103, , "Result exceeds maxRows or the 250,000-cell limit. Filter or aggregate the SQL; no partial result was returned."
End Function

Private Function CellValue(ByVal value As Variant, ByVal field As Object) As Variant
    If IsNull(value) Then
        CellValue = ""
    Else
        If VarType(value) = vbString Then
            If Len(value) > 32767 Then Err.Raise vbObjectError + 2104, , "A text value exceeds Excel's 32,767-character cell limit. Shorten it in the SQL."
        End If
        Select Case field.Type
            Case 14, 131, 139 ' adDecimal, adNumeric, adVarNumeric
                If field.Precision > 15 Then
                    ' CStr(Decimal) retains precision; conversion through Double does not.
                    If VarType(value) = vbDouble Or VarType(value) = vbSingle Then
                        Err.Raise vbObjectError + 2104, , "The ODBC driver returned a high-precision decimal as floating point. CAST that column AS VARCHAR in the SQL to preserve its value."
                    End If
                    CellValue = CStr(value)
                Else
                    CellValue = value
                End If
            Case 20, 21 ' signed/unsigned 64-bit integers can exceed Excel precision
                If VarType(value) = vbDouble Or VarType(value) = vbSingle Then
                    If Abs(value) > 999999999999999# Then Err.Raise vbObjectError + 2104, , "The ODBC driver returned a large integer as floating point. CAST it AS VARCHAR in the SQL."
                End If
                If Len(Replace(CStr(value), "-", "")) > 15 Then
                    CellValue = CStr(value)
                Else
                    CellValue = value
                End If
            Case Else
                If IsArray(value) Or IsObject(value) Then Err.Raise vbObjectError + 2104, , "Unsupported column value. CAST it AS VARCHAR in the SQL."
                CellValue = value
        End Select
    End If
End Function
