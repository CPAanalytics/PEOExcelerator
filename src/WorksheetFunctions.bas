Attribute VB_Name = "WorksheetFunctions"
Option Explicit

Public Function SFVALUE(ByVal sql As Variant, Optional ByVal refresh As Variant = 0) As Variant
    On Error GoTo Failed
    SFVALUE = FormulaEngine.RunQuery(Application.Caller, sql, True, False, refresh, 1)
    Exit Function
Failed:
    SFVALUE = CVErr(xlErrValue)
End Function

Public Function SFQUERY(ByVal sql As Variant, Optional ByVal headers As Variant = True, Optional ByVal refresh As Variant = 0, Optional ByVal maxRows As Variant = 10000) As Variant
    On Error GoTo Failed
    SFQUERY = FormulaEngine.RunQuery(Application.Caller, sql, False, headers, refresh, maxRows)
    Exit Function
Failed:
    SFQUERY = CVErr(xlErrValue)
End Function

Public Function SFDATABASES(Optional ByVal refresh As Variant = 0) As Variant
    SFDATABASES = MetadataResult("databases", Empty, Empty, Empty, refresh)
End Function

Public Function SFSCHEMAS(ByVal database As Variant, Optional ByVal refresh As Variant = 0) As Variant
    SFSCHEMAS = MetadataResult("schemas", database, Empty, Empty, refresh)
End Function

Public Function SFTABLES(ByVal database As Variant, ByVal schema As Variant, Optional ByVal refresh As Variant = 0) As Variant
    SFTABLES = MetadataResult("tables", database, schema, Empty, refresh)
End Function

Public Function SFCOLUMNS(ByVal database As Variant, ByVal schema As Variant, ByVal table As Variant, Optional ByVal refresh As Variant = 0) As Variant
    SFCOLUMNS = MetadataResult("columns", database, schema, table, refresh)
End Function

Private Function MetadataResult(ByVal kind As String, ByVal database As Variant, ByVal schema As Variant, ByVal table As Variant, ByVal refresh As Variant) As Variant
    Dim sql As String, db As String, schemaName As String, tableName As String, caller As Range
    On Error GoTo Failed
    Set caller = Application.Caller
    If kind <> "databases" Then db = SQLValidation.Identifier(MetadataText(database))
    If kind = "tables" Or kind = "columns" Then schemaName = SQLValidation.Literal(MetadataText(schema))
    Select Case kind
        Case "databases"
            ' Execute SHOW directly; do not depend on mutable LAST_QUERY_ID state.
            sql = "SHOW DATABASES"
        Case "schemas"
            sql = "SELECT SCHEMA_NAME FROM " & db & ".INFORMATION_SCHEMA.SCHEMATA ORDER BY SCHEMA_NAME"
        Case "tables"
            sql = "SELECT TABLE_NAME, TABLE_TYPE FROM " & db & ".INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = " & schemaName & " ORDER BY TABLE_NAME"
        Case "columns"
            tableName = SQLValidation.Literal(MetadataText(table))
            sql = "SELECT COLUMN_NAME, DATA_TYPE, NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE FROM " & db & ".INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = " & schemaName & " AND TABLE_NAME = " & tableName & " ORDER BY ORDINAL_POSITION"
    End Select
    MetadataResult = FormulaEngine.RunQuery(caller, sql, False, True, refresh, 10000, True)
    Exit Function
Failed:
    If Not caller Is Nothing Then FormulaEngine.RecordInputError caller, Err.Description
    MetadataResult = CVErr(xlErrValue)
End Function

Private Function MetadataText(ByVal value As Variant) As String
    Dim cell As Range
    Dim result As Variant
    If IsObject(value) Then
        If Not TypeOf value Is Range Then Err.Raise vbObjectError + 2111, , "Object names must be text or single cells."
        Set cell = value
        If cell.CountLarge <> 1 Then Err.Raise vbObjectError + 2111, , "Object names must reference single cells."
        result = cell.Value2
    Else
        result = value
    End If
    If IsError(result) Or IsArray(result) Or IsNull(result) Then Err.Raise vbObjectError + 2111, , "Object names must be text without Excel errors."
    MetadataText = CStr(result)
    If Len(MetadataText) = 0 Then Err.Raise vbObjectError + 2111, , "Database, schema and table names must not be empty."
End Function

Public Function SFSTATUS(Optional ByVal queryCell As Range = Nothing) As Variant
    On Error GoTo Failed
    SFSTATUS = FormulaEngine.Status(Application.Caller, queryCell)
    Exit Function
Failed:
    SFSTATUS = CVErr(xlErrValue)
End Function

Public Sub RegisterFunctions()
    ' Registration is an interactive/startup task. Never call it from a formula.
    On Error Resume Next
    Application.MacroOptions Macro:="SFVALUE", Description:="Return one Snowflake value. Connect first; change refresh to fetch again.", Category:="Snowflake", ArgumentDescriptions:=Array("One SELECT returning one column and at most one row.", "Optional refresh value or cell; default 0.")
    Application.MacroOptions Macro:="SFQUERY", Description:="Spill a Snowflake SELECT into rows and columns without clearing the sheet.", Category:="Snowflake", ArgumentDescriptions:=Array("One SELECT or WITH ... SELECT; text or a single cell.", "Include column names; default TRUE.", "Change this value to fetch again; default 0.", "Maximum data rows; default 10000. Also limited to 250000 cells.")
    Application.MacroOptions Macro:="SFDATABASES", Description:="List accessible Snowflake database names.", Category:="Snowflake", ArgumentDescriptions:=Array("Change this value to refresh metadata; default 0.")
    Application.MacroOptions Macro:="SFSCHEMAS", Description:="List schemas accessible in a Snowflake database.", Category:="Snowflake", ArgumentDescriptions:=Array("Exact database name, without SQL quotes.", "Change this value to refresh metadata; default 0.")
    Application.MacroOptions Macro:="SFTABLES", Description:="List accessible Snowflake tables and views with their types.", Category:="Snowflake", ArgumentDescriptions:=Array("Exact database name, without SQL quotes.", "Exact schema name, without SQL quotes.", "Change this value to refresh metadata; default 0.")
    Application.MacroOptions Macro:="SFCOLUMNS", Description:="List columns, types, precision, scale and nullability in source order.", Category:="Snowflake", ArgumentDescriptions:=Array("Exact database name.", "Exact schema name.", "Exact table or view name.", "Change this value to refresh metadata; default 0.")
    Application.MacroOptions Macro:="SFSTATUS", Description:="Show connection status or a query's last attempt and successful refresh.", Category:="Snowflake", ArgumentDescriptions:=Array("Optional reference to one query formula anchor in this workbook.")
End Sub
