VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OracleConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private conn As Object

Public Sub Initialize(ByVal connection As Object)
    Set conn = connection
End Sub

Public Sub ExecuteNonQuery(ByVal sql As String)
    On Error GoTo ErrorHandler
    
    conn.Execute sql
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar a consulta: " & Err.Description, vbCritical
End Sub

Public Function ExecuteQuery(ByVal sql As String) As Object
    Dim rs As Object
    On Error GoTo ErrorHandler
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3 'adOpenStatic
    
    Set ExecuteQuery = rs
    Exit Function

ErrorHandler:
    MsgBox "Erro ao executar a consulta: " & Err.Description, vbCritical
    Set ExecuteQuery = Nothing
End Function

Public Sub CloseConnection()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub
