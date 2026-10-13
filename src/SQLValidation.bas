Attribute VB_Name = "SQLValidation"
Option Explicit

' Scan SQL rather than splitting at semicolons: strings, identifiers and comments
' may all contain punctuation. This is a statement guard, not a privilege boundary.
Public Function ValidateSelect(ByVal sql As String) As String
    Dim i As Long, depth As Long, n As Long, start As Long
    Dim ch As String, token As String, first As String, command As String
    Dim ended As Boolean, closed As Boolean, quote As String
    n = Len(sql)
    i = 1
    Do While i <= n
        ch = Mid$(sql, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            i = i + 1
        ElseIf Mid$(sql, i, 2) = "--" Then
            i = i + 2
            Do While i <= n
                If Mid$(sql, i, 1) = vbCr Or Mid$(sql, i, 1) = vbLf Then Exit Do
                i = i + 1
            Loop
        ElseIf Mid$(sql, i, 2) = "/*" Then
            start = InStr(i + 2, sql, "*/", vbBinaryCompare)
            If start = 0 Then GoTo Malformed
            i = start + 2
        Else
            If ended Then
                ValidateSelect = "Only one SELECT statement is allowed."
                Exit Function
            End If
            If ch = "'" Or ch = Chr$(34) Then
                quote = ch
                i = i + 1
                closed = False
                Do While i <= n
                    ch = Mid$(sql, i, 1)
                    If ch = quote Then
                        If Mid$(sql, i + 1, 1) = quote Then
                            i = i + 2
                        Else
                            i = i + 1
                            closed = True
                            Exit Do
                        End If
                    ElseIf ch = "\" And quote = "'" Then
                        i = i + 2
                    Else
                        i = i + 1
                    End If
                Loop
                If Not closed Then GoTo Malformed
            ElseIf Mid$(sql, i, 2) = "$$" Then
                start = InStr(i + 2, sql, "$$", vbBinaryCompare)
                If start = 0 Then GoTo Malformed
                i = start + 2
            ElseIf ch = "(" Then
                depth = depth + 1
                i = i + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth < 0 Then GoTo Malformed
                i = i + 1
            ElseIf ch = ";" Then
                If depth <> 0 Then GoTo Malformed
                ended = True
                i = i + 1
            ElseIf ch Like "[A-Za-z_]" Then
                start = i
                i = i + 1
                Do While i <= n
                    If Not Mid$(sql, i, 1) Like "[A-Za-z0-9_$]" Then Exit Do
                    i = i + 1
                Loop
                token = UCase$(Mid$(sql, start, i - start))
                If first = "" Then
                    first = token
                    If first <> "SELECT" And first <> "WITH" Then GoTo NotSelect
                End If
                If depth = 0 And command = "" Then
                    If token = "SELECT" Then
                        command = token
                    ElseIf IsWriteWord(token) Then
                        GoTo NotSelect
                    End If
                End If
            Else
                If first = "" Then GoTo NotSelect
                i = i + 1
            End If
        End If
    Loop
    If depth <> 0 Then GoTo Malformed
    If command <> "SELECT" Then GoTo NotSelect
    Exit Function
NotSelect:
    ValidateSelect = "Use a single SELECT or WITH ... SELECT query. Database changes are not allowed in formulas."
    Exit Function
Malformed:
    ValidateSelect = "SQL contains an unterminated quote/comment or unbalanced parentheses."
End Function

Private Function IsWriteWord(ByVal token As String) As Boolean
    Select Case token
        Case "INSERT", "UPDATE", "DELETE", "MERGE", "CREATE", "ALTER", "DROP", "TRUNCATE", "CALL", "COPY", "PUT", "GET", "USE", "GRANT", "REVOKE", "BEGIN", "COMMIT", "ROLLBACK"
            IsWriteWord = True
    End Select
End Function

Public Function Identifier(ByVal value As String) As String
    If Len(value) = 0 Then Err.Raise vbObjectError + 2101, , "Database, schema and table names must not be empty."
    Identifier = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Public Function Literal(ByVal value As String) As String
    Literal = "'" & Replace(Replace(value, "\", "\\"), "'", "''") & "'"
End Function
