Option Explicit
Option Private Module

Public Enum SpeedSetting
    Normal = 0
    Fast = 1
End Enum

Public Sub SetSpeed(ByVal Speed As SpeedSetting, Optional ByVal DisableAlerts As Boolean = False)
    With Application
        Select Case Speed
            Case SpeedSetting.Normal
                .ScreenUpdating = True
                .Calculation = xlCalculationAutomatic
                .EnableEvents = True
                .DisplayStatusBar = True
                
            Case SpeedSetting.Fast
                .ScreenUpdating = False
                .Calculation = xlCalculationManual
                .EnableEvents = False
                .DisplayStatusBar = False
        End Select
        
        .DisplayAlerts = Not DisableAlerts
    End With
End Sub

Public Function Delay(ByVal MilliSeconds As Long) As Variant
    Delay = Timer + MilliSeconds / 1000
    Do While Timer < Delay: DoEvents: Loop
End Function

Public Sub PasteDataIntoTable(ByVal Data As Variant, ByVal ws As Worksheet, ByVal TableName As String)
    ClearFilters ws

    Dim Table As ListObject
    Set Table = ws.ListObjects(TableName)

    With Table
        ' Check if the table has any data
        If .ListRows.Count > 0 Then
            .DataBodyRange.Value2 = vbNullString
        End If
          
        Dim HasTotal As Boolean
        If .ShowTotals = True Then
            HasTotal = True
            .ShowTotals = False
        End If
        
        ' Resize the table to fit the incoming data
        .Resize ws.Range(.Range.Cells(1, 1), ws.Cells(.HeaderRowRange.Row + UBound(Data) - LBound(Data) + 1, .ListColumns.Count + .Range.Cells(1, 1).Column - 1))

        ' Check if the incoming data is a single row
        If LBound(Data) = UBound(Data) Then
            Dim j As Long
            If LBound(Data, 2) = 0 Then j = 1
            
            Dim i As Long
            For i = 1 To Table.Range.Columns.Count
                .Range(2, i).Value2 = Data(LBound(Data), i - j)
            Next
        Else
            .DataBodyRange.Value = Data
        End If
        
        If HasTotal Then
            .ShowTotals = True
        End If
    End With
End Sub

Public Sub ExecuteShellWait(ByVal cmd As String)
    Dim Shell As Object
    Set Shell = CreateObject("WScript.Shell")
    
    Shell.Run cmd, 0, True
    
    Set Shell = Nothing
End Sub

Private Sub ClearFilters(ByVal ws As Worksheet)
    Dim Table As ListObject
    
    For Each Table In ws.ListObjects
        With Table
            If .ShowAutoFilter Then
                If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
            End If
            .ShowAutoFilter = False
            
            .Range.AutoFilter
            .Sort.SortFields.Clear
        End With
    Next
End Sub
