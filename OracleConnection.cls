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

Public Function ExecuteNonQuery(ByVal sql As String, Optional ByRef affectedRows As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    conn.Execute sql, affectedRows
    ExecuteNonQuery = True
    Exit Function

ErrorHandler:
    MsgBox "Erro ao executar a consulta: " & Err.Description, vbCritical
    ExecuteNonQuery = False
End Function

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

Public Sub BeginTransaction()
    On Error GoTo ErrorHandler
    
    conn.BeginTrans
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao iniciar a transação: " & Err.Description, vbCritical
End Sub

Public Sub CommitTransaction()
    On Error GoTo ErrorHandler
    
    conn.CommitTrans
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao confirmar a transação: " & Err.Description & ". Realizando rollback.", vbCritical
    RollbackTransaction
End Sub

Public Sub RollbackTransaction()
    On Error GoTo ErrorHandler
    
    conn.RollbackTrans
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao reverter a transação: " & Err.Description, vbCritical
End Sub
