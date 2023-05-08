Attribute VB_Name = "Utils"
Option Explicit
Option Private Module

Public Enum SpeedSetting
    Normal = 0
    Fast = 1
End Enum

Public Sub SetSpeed(ByVal Speed As SpeedSetting, Optional ByVal disableAlerts As Boolean = False)
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
        
        .DisplayAlerts = Not disableAlerts
    End With
End Sub

Public Function Delay(ByVal MilliSeconds As Long) As Variant
    Delay = Timer + MilliSeconds / 1000
    Do While Timer < Delay: DoEvents: Loop
End Function

Private Sub PasteDataIntoTable(ByVal Data As Variant, ByVal ws As Worksheet, ByVal TableName As String)
    ClearFilters ws

    Dim Table As ListObject
    Set Table = ws.ListObjects(TableName)

    With Table
        ' Check if the table has any data
        If .ListRows.Count > 0 Then
            .DataBodyRange.Value2 = vbNullString
        End If

        ' Resize the table to fit the incoming data
        .Resize ws.Range(.Range.Cells(1, 1), ws.Cells(.HeaderRowRange.Row + UBound(Data) + 1, .ListColumns.Count + .Range.Cells(1, 1).Column - 1))

        ' Check if the incoming data is a single row
        If UBound(Data) = 0 Then
            Dim i As Long
            For i = 1 To Table.Range.Columns.Count
                .Range(2, i).Value2 = Data(0, i - 1)
            Next
        Else
            .DataBodyRange.Value = Data
        End If
    End With
End Sub

Private Sub ClearFilters(ByVal ws As Worksheet)
    Dim Table As ListObject
    
    For Each Table In ws.ListObjects
        If Table.ShowAutoFilter Then
            If Table.AutoFilter.FilterMode Then Table.AutoFilter.ShowAllData
        End If
        Table.ShowAutoFilter = False
        
        Table.Range.AutoFilter
        Table.Sort.SortFields.Clear
    Next
End Sub
