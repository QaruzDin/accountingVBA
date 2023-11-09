Attribute VB_Name = "AccountingModule"
Option Explicit

Public Sub Del_Cell_Automation()

    ' To delete the cells that is no needed with certain condition from top to bottom automatically

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    
        Set ws = ThisWorkbook.ActiveSheet
        
        ' select range that contain LABEL, for instance B2 contains Label
        Range("B2").Select
        Set rng = ws.Range("B2")
        
        'iterrating one cell at the time
        For i = 1 To ws.Rows.Count
            Set cell = rng.Offset(i - 1, 0) ' will be moving one cell above

            ' Exit the FOR loop if there is a Cell with the value "TOTAL AMOUNT", it can be changed as needed.
            If cell.Value = "TOTAL AMOUNT" Then
                Exit For

            'Deleting Cell with condition 1, it can be changed as needed.
            ElseIf cell.Value = "(Condition 1)" Then
                cell.EntireRow.Delete
                i = i - 1 ' Decrease the value of i because the row has been deleted
            ElseIf cell.Offset(0, 1) = 0 And cell.Offset(0, 1).Interior.Color <> RGB(242, 242, 242) Then
            'Deleting Cell with condition 1, it can be changed as needed.
                cell.EntireRow.Delete
                i = i - 1 ' Decrease the value of i because the row has been deleted

            ' It can be added more conditions as needed in this block of FOR loop
            End If
        Next i
End Sub

Public Sub cleaningData_perSheet()

    ' To delete coloum(s) or certain cell range that no needed with certain condition from sheet to another automatically

    Dim lastRow As Long
    Dim dataRange1 As Range
    Dim datarange2 As Range
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.ActiveSheet

    For Each ws In ThisWorkbook.Sheets

        ' (optional) Exit the FOR EACH loop if there is a sheet named "Name of Sheet", it can be changed as needed.
        if ws.Name = "Name of Sheet" Then
            Exit For
        End If

        'Find the last row of data on the sheet according column "B", it can be changed as needed.
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
        ' Set coloum(s) or range which is want to be deleted, example : coloum C until M, C2 to M2 which value that want to be deleted begins
        Set dataRange1 = ws.Range("C2:M2").Resize(lastRow - 1, 11) ' 11 = range count C to M
        
        ' Deleting "dataRange1"
        dataRange1.Delete
        
        ' Set coloum(s) range which is want to be deleted, example : coloum D until F, D2 to F2 which value that want to be deleted begins
        Set dataRange2 = ws.Range("D2:F2").Resize(lastRow - 1, 3) ' 3 = range count D to F
        
        ' Deleting "dataRange2"
        dataRange2.Delete

        ' It can be added more conditions as needed in this block of FOR EACH loop
    Next ws ' wil move to another sheet of this workbook
End Sub



