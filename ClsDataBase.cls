VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Results array always start in (0, 0)
'IsArray is the way to check if recordset returned some value

'Some uses examples

'Private Sub SimpleSelectExample()
'    Dim DB As ClsDataBase
'    Set DB = New ClsDataBase
'
'    With DB
'        Dim SQL As String
'        SQL = "SELECT 1 FROM dual"
'
'        .ExecuteQuery SQL, "SELECT"
'
'        Dim Results As Variant
'        Results = .Results
'
'        If IsArray(Results) Then
'            MsgBox Results(0, 0)
'        End If
'    End With
'
'    Set DB = Nothing
'End Sub

'Private Sub SimpleUpdateExample()
'    Dim DB As ClsDataBase
'    Set DB = New ClsDataBase
'
'    With DB
'        Dim SQL As String
'               SQL = "UPDATE"
'            SQL = SQL & " e120ped"
'        SQL = SQL & " SET"
'            SQL = SQL & " e120ped.codfpg = 12"
'        SQL = SQL & " WHERE"
'            SQL = SQL & " e120ped.codemp = 1"
'            SQL = SQL & " AND e120ped.codfil = 1"
'            SQL = SQL & " AND e120ped.numped = 467750"
'
'        .ExecuteQuery SQL, "UPDATE"
'    End With
'
'    Set DB = Nothing
'End Sub

Private Connection As Object
Private RecordSet As Object

Private Const PROVIDER As String = "OraOLEDB.Oracle.1"
Private Const PROTOCOL As String = "TCP"
Private Const HOST As String = ""
Private Const PORT As String = ""
Private Const SERVER As String = "DEDICATED"
Private Const DB_SOURCE As String = ""
Private Const USER As String = ""
Private Const PASSWORD As String = ""

Private Type ClassType
    SQL As String
    Results As Variant
End Type

Private This As ClassType

Public Property Let Results(ByVal Value As Variant)
    This.Results = Value
End Property

Public Property Get Results() As Variant
    Results = This.Results
End Property

Public Sub ExecuteQuery(ByVal SQL As String, ByVal QueryType As String)
    Const adCmdText As Byte = 1
    Const adOpenStatic As Byte = 3
    Const adLockReadOnly As Byte = 1

    On Error GoTo ErrorHandler
    
    If QueryType = "SELECT" Then
        Set RecordSet = CreateObject("ADODB.Recordset")
        
        RecordSet.Open SQL, Connection, adOpenStatic, adLockReadOnly
        Results = RecordSetToArray(RecordSet)
        RecordSet.Close
        
        Set RecordSet = Nothing
    Else
        Connection.Execute SQL, , adCmdText
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Ocorreu um erro: " & Err.Description
End Sub

Private Sub Class_Initialize()
    Set Connection = CreateObject("ADODB.Connection")
    With Connection
        Const CONNECTION_STRING As String = "Provider=" & PROVIDER & ";Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=" & PROTOCOL & ")(HOST=" & HOST & ")(PORT=" & PORT & "))(CONNECT_DATA=(SERVER=" & SERVER & ")(SERVICE_NAME=" & DB_SOURCE & ")));user ID=" & USER & ";password=" & PASSWORD & ";"
        
        .ConnectionString = CONNECTION_STRING
        .Open
    End With
End Sub

Private Sub Class_Terminate()
    Connection.Close
    Set Connection = Nothing
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
                Next
                i = i + 1
                .MoveNext
            Loop
        End If
    End With

    RecordSetToArray = arr
End Function
