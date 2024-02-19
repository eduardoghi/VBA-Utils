Attribute VB_Name = "Utils"
Option Explicit
Option Private Module

Public Enum speed_setting
    normal = 0
    fast = 1
End Enum

Public Sub set_speed(ByVal speed As speed_setting, Optional ByVal disable_alerts As Boolean = False)
    With Application
        Select Case speed
            Case speed_setting.normal
                .ScreenUpdating = True
                .Calculation = xlCalculationAutomatic
                .EnableEvents = True
                .DisplayStatusBar = True
                
            Case speed_setting.fast
                .ScreenUpdating = False
                .Calculation = xlCalculationManual
                .EnableEvents = False
                .DisplayStatusBar = False
        End Select
        
        .DisplayAlerts = Not disable_alerts
    End With
End Sub

Public Function delay(ByVal milli_seconds As Long) As Variant
    delay = Timer + milli_seconds / 1000
    Do While Timer < delay: DoEvents: Loop
End Function

Public Sub paste_data_into_table(ByVal data As Variant, ByVal ws As Worksheet, ByVal table_name As String)
    clear_filters ws

    Dim table As ListObject
    Set table = ws.ListObjects(table_name)

    With table
        ' Check if the table has any data
        If .ListRows.Count > 0 Then
            .DataBodyRange.Value2 = vbNullString
        End If
          
        Dim has_total As Boolean
        If .ShowTotals = True Then
            has_total = True
            .ShowTotals = False
        End If
        
        ' Resize the table to fit the incoming data
        .Resize ws.Range(.Range.Cells(1, 1), ws.Cells(.HeaderRowRange.Row + UBound(data) - LBound(data) + 1, .ListColumns.Count + .Range.Cells(1, 1).Column - 1))

        ' Check if the incoming data is a single row
        If LBound(data) = UBound(data) Then
            Dim j As Long
            If LBound(data, 2) = 0 Then j = 1
            
            Dim i As Long
            For i = 1 To table.Range.Columns.Count
                .Range(2, i).Value2 = data(LBound(data), i - j)
            Next
        Else
            .DataBodyRange.Value = data
        End If
        
        If has_total Then
            .ShowTotals = True
        End If
    End With
End Sub

Public Sub execute_shell_wait(ByVal cmd As String)
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    shell.Run cmd, 0, True
    
    Set shell = Nothing
End Sub

Private Sub clear_filters(ByVal ws As Worksheet)
    Dim table As ListObject
    
    For Each table In ws.ListObjects
        With table
            If .ShowAutoFilter Then
                If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
            End If
            .ShowAutoFilter = False
            
            .Range.AutoFilter
            .Sort.SortFields.Clear
        End With
    Next
End Sub
