VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "FakeConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public State As Long
Public Executions As Long
Public FailQuery As Boolean
Public LastSQL As String

Private Sub Class_Initialize()
    State = 1
End Sub

Public Function Execute(ByVal sql As String) As Object
    Dim rs As Object, i As Long
    Set rs = CreateObject("ADODB.Recordset")
    If sql = "SELECT CURRENT_DATABASE(), CURRENT_SCHEMA(), CURRENT_ROLE(), CURRENT_WAREHOUSE(), CURRENT_USER()" Then
        For i = 0 To 4
            rs.Fields.Append "context" & i, 200, 100
        Next i
        rs.Open
        rs.AddNew
        rs.Fields(0).Value = "TEST_DB"
        rs.Fields(1).Value = "PUBLIC"
        rs.Fields(2).Value = "ACCOUNTANT"
        rs.Fields(3).Value = "TEST_WH"
        rs.Fields(4).Value = "TEST_USER"
    Else
        Executions = Executions + 1
        LastSQL = sql
        If FailQuery Then Err.Raise vbObjectError + 2200, , "Test database query failed."
        If sql = "SHOW DATABASES" Then
            rs.Fields.Append "created_on", 200, 100
            rs.Fields.Append "name", 200, 100
            rs.Open
            rs.AddNew
            rs.Fields(0).Value = "2026-01-01"
            rs.Fields(1).Value = "TEST_DB"
        Else
            rs.Fields.Append "VALUE", 3
            rs.Open
            rs.AddNew
            rs.Fields(0).Value = 42
        End If
    End If
    rs.Update
    rs.MoveFirst
    Set Execute = rs
End Function
