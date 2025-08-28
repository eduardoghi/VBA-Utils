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

Public Function ExecuteQuery(ByVal sql As String, ParamArray params() As Variant) As Object
    On Error GoTo ErrorHandler
    
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1 ' adCmdText
    
    Dim i As Long
    For i = LBound(params) To UBound(params) Step 2
        Dim paramType As Long
        paramType = GetADOType(params(i + 1))
        
        Dim paramSize As Long
        If paramType = 200 Then
            paramSize = Len(params(i + 1))
        Else
            paramSize = 0
        End If
        
        Dim param As Object
        Set param = cmd.CreateParameter(params(i), paramType, 1, paramSize, params(i + 1)) ' 1 = adParamInput
        cmd.Parameters.Append param
    Next i
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorType = 3 ' adOpenStatic
    rs.Open cmd
    
    Set ExecuteQuery = rs
    Exit Function

ErrorHandler:
    MsgBox "Erro ao executar a consulta: " & Err.Description, vbCritical
    Set ExecuteQuery = Nothing
End Function

Private Function GetADOType(ByRef paramValue As Variant) As Integer
    Select Case VarType(paramValue)
        Case vbInteger
            GetADOType = 3 ' adInteger
        Case vbLong
            GetADOType = 3 ' adInteger
        Case vbSingle
            GetADOType = 4 ' adSingle
        Case vbDouble
            GetADOType = 5 ' adDouble
        Case vbCurrency
            GetADOType = 6 ' adCurrency
        Case vbDate
            GetADOType = 7 ' adDate
        Case vbString
            GetADOType = 200 ' adVarChar
        Case vbBoolean
            GetADOType = 11 ' adBoolean
        Case Else
            GetADOType = 200 ' adVarChar
    End Select
End Function

Public Sub CloseConnection()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub

Public Function IsConnected() As Boolean
    On Error Resume Next
    IsConnected = Not (conn Is Nothing Or conn.State = 0)
End Function

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