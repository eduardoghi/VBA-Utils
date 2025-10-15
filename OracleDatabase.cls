VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OracleDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetOracleConnection(ByVal user As String, ByVal password As String, ByVal connectString As String) As OracleConnection
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim connectionString As String
    
    On Error GoTo ErrorHandler
    
    connectionString = "Provider=OraOLEDB.Oracle;" & _
                       "Data Source=" & connectString & ";" & _
                       "User Id=" & user & ";" & _
                       "Password=" & password & ";"
    
    conn.Open connectionString
    
    Dim oracleConn As OracleConnection
    Set oracleConn = New OracleConnection
    oracleConn.Initialize conn
    
    Set GetOracleConnection = oracleConn
    Exit Function

ErrorHandler:
    MsgBox "Erro ao conectar ao banco de dados: " & Err.Description, vbCritical
    Set GetOracleConnection = Nothing
End Function
