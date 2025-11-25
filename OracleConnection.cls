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

Public Function ExecuteQueryArray(ByVal sql As String, ParamArray params() As Variant) As Variant
    Dim rs As Object
    Set rs = ExecuteQuery(sql, params)

    ExecuteQueryArray = RecordSetToArray(rs)

    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close

        Set rs = Nothing
    End If
End Function

Public Function ExecuteNonQuery(ByVal sql As String, Optional ByRef affectedRows As Long = -1) As Boolean
    conn.Execute sql, affectedRows
    ExecuteNonQuery = True
End Function

Public Function ExecuteQuery(ByVal sql As String, ParamArray params() As Variant) As Object
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1 ' adCmdText
    
    If UBound(params) > 0 Then
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
    End If
        
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorType = 3 ' adOpenStatic
    rs.Open cmd
    
    Set ExecuteQuery = rs
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
    conn.BeginTrans
End Sub

Public Sub CommitTransaction()
    conn.CommitTrans
End Sub

Public Sub RollbackTransaction()
    conn.RollbackTrans
End Sub

Private Function RecordSetToArray(ByVal RecordSet As Object) As Variant
    With RecordSet
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            Dim arr As Variant
            ReDim arr(0 To .RecordCount - 1, 0 To .Fields.Count - 1)
            
            Do Until .EOF = True
                Dim i As Long

                Dim j As Long
                For j = 0 To .Fields.Count - 1
                    arr(i, j) = .Fields.Item(j).Value
                Next j
                i = i + 1
                .MoveNext
            Loop
        End If
    End With

    RecordSetToArray = arr
End Function

