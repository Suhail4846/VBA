Attribute VB_Name = "AutomaticPivotTableGenerator"
Option Explicit

Public Sub GenerateAutomaticPivotTable()

    Dim sourceRange As Range
    Dim wsReport As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim i As Long
    Dim valueCol As Long

    On Error Resume Next
    Set sourceRange = Application.InputBox( _
        "Select data including headers", _
        "Pivot Table Generator", Type:=8)
    On Error GoTo ErrorHandler

    If sourceRange Is Nothing Then Exit Sub

    If sourceRange.Rows.Count < 2 Or sourceRange.Columns.Count < 2 Then
        MsgBox "Select at least 2 rows and 2 columns.", vbExclamation
        Exit Sub
    End If

    For i = 1 To sourceRange.Columns.Count
        If Len(Trim(CStr(sourceRange.Cells(1, i).Value))) = 0 Then
            MsgBox "All columns must have headers.", vbExclamation
            Exit Sub
        End If
    Next i

    valueCol = 0

    For i = 2 To sourceRange.Columns.Count
        If IsNumeric(sourceRange.Cells(2, i).Value) And _
           Len(CStr(sourceRange.Cells(2, i).Value)) > 0 Then
            valueCol = i
            Exit For
        End If
    Next i

    If valueCol = 0 Then
        MsgBox "No numeric column found.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Set wsReport = ThisWorkbook.Worksheets.Add
    wsReport.Name = "Pivot_" & Format(Now, "hhmmss")

    wsReport.Range("A1").Value = "Automatic Pivot Table Report"
    wsReport.Range("A1").Font.Bold = True

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sourceRange.Address( _
            ReferenceStyle:=xlR1C1, External:=True))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsReport.Range("A3"), _
        TableName:="AutoPivot_" & Format(Now, "hhmmss"))

    With pt.PivotFields(CStr(sourceRange.Cells(1, 1).Value))
        .Orientation = xlRowField
        .Position = 1
    End With

    pt.AddDataField _
        pt.PivotFields(CStr(sourceRange.Cells(1, valueCol).Value)), _
        "Sum of " & CStr(sourceRange.Cells(1, valueCol).Value), xlSum

    pt.RowAxisLayout xlTabularRow
    pt.TableStyle2 = "PivotStyleMedium9"
    pt.RefreshTable

    wsReport.Columns.AutoFit

    Application.ScreenUpdating = True

    MsgBox "Pivot Table created successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical

End Sub
