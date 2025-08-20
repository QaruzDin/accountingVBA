Attribute VB_Name = "DelCellAutomationV1"
Option Explicit

Public Sub Del_Cell_Automation()

    ' To delete the cells that is no needed with certain condition from top to bottom automatically

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim startTime As Single
    Dim timeout As Single
    Dim newName As String
    Dim inputColumn As String
    Dim response As VbMsgBoxResult
    Dim confirm As VbMsgBoxResult
    Dim abortMsg As VbMsgBoxResult
    Dim colNumb As Integer
    Dim indicator As String
    Dim currentName As String
    Dim currentWSName As String
    Dim indexws As Integer

    confirm = MsgBox("Are you sure want to continue this module?" & vbCrLf & _
                    "This is would be take a moment", vbYesNo)
    If confirm = vbNo Then
        Exit Sub
    End If

    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Cells

    ' to backup the origin sheet
    currentWSName = ws.Name
    ws.Copy Before:=ws
    indexws = ws.Index
    ws.Select
    newName = ws.Name & "_RESULT"
    ws.Name = newName
    ThisWorkbook.Sheets(indexws - 1).Name = currentWSName
    
    ' giving an exit for loop by adding a string
    indicator = "~"
    ws.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    Selection.Value = indicator
  
    ' formatting entire cells in the worksheet as values
    rng.UnMerge
    rng.WrapText = False
    rng.Rows.AutoFit
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    
    ' setting range starter (NOTE: This starter should not be deleted)
    Range("A1").Select
    Range("A1").Value = 0
    Set rng = ws.Range("A1")
    
    ' timeout setting for infinte loop prevention (timeout is in second)
    startTime = Timer
    timeout = 60

    'iterrating one cell at the time
    For i = 1 To ws.Rows.Count
        Set cell = rng.Offset(i - 1, 0) ' will be moving one cell above

        ' Exit the FOR loop if there is a Cell with the value "~", it can be changed as needed.
        If cell.Value = indicator Then
            cell.EntireRow.Delete
            Range("A1").EntireRow.Delete
            Exit For

        ElseIf IsEmpty(cell.Value) Then
            cell.EntireRow.Delete
            i = i - 1
        ElseIf (IsEmpty(cell.Offset(0, 1).Value) And cell.Value <> 0) Then
            cell.EntireRow.Delete
            i = i - 1
        ElseIf cell.Interior.Color <> RGB(255, 255, 255) Then
            cell.EntireRow.Delete
            i = i - 1
            
        ' Here! You can add more conditions as needed in this block of FOR loop

        ' Check if the timeout duration has been exceeded
        ElseIf Timer - startTime > timeout Then
            MsgBox "Time limit is reached out"
            Exit Sub
        End If
    Next i
    
    ' formatting : labeling columns and make them centered

    currentName = Range("A1").Value
    With Range("A1")
        .Select
        .Value = UserInput(currentName)
        .HorizontalAlignment = xlCenter
    End With
    currentName = Range("B1").Value
    With Range("B1")
        .Select
        .Value = UserInput(currentName)
        .HorizontalAlignment = xlCenter
    End With
    
    ' to configure column delete
    
    Do
        inputColumn = InputBox("Please, input the location of the column you want to keep it. The other columns from C will be erased:", "Input Column")
        If inputColumn = vbNullString Then
            abortMsg = MsgBox("Are you sure want to end the module?" & vbCrLf & _
                    "It would be used current formatting.", vbExclamation + vbYesNo)
            If abortMsg = vbYes Then
                Exit Sub
            End If
        End If
        
        If Not IsNumeric(inputColumn) Then
            On Error Resume Next
            colNumb = Columns(inputColumn).Column
            On Error GoTo 0
            
            If colNumb > 3 Then
                ws.Range(ws.Columns(3), ws.Columns(colNumb - 1)).Delete
                Exit Do
            Else
                MsgBox "The column given must be greater than column B. Please try again.", vbExclamation
            End If
        Else
            MsgBox "The column given must be alphabetical. Please try again.", vbExclamation
        End If
    Loop
    
    ws.Columns("D:Z").Delete
    
    ' formatting : the rest range
    
    currentName = Range("C1").Value
    With Range("C1")
        .Select
        .Value = UserInput(currentName)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(3).AutoFit
    finalSum
    MsgBox "The module is success.", vbInformation

End Sub

Function UserInput(defaultName As String) As String
    Dim inputUser As String
    
    inputUser = InputBox("Please insert the name:" & vbCrLf & _
                "Cancel to keep current name")
    
    ' when user cancel it
    If inputUser = vbNullString Then
        UserInput = defaultName
        Exit Function
    End If
    
    UserInput = inputUser
End Function

Sub copySheet()
    Dim ws As Worksheet
    Dim indexws As Integer
    Dim currentWSName As String
    
    Set ws = ThisWorkbook.ActiveSheet
    
    currentWSName = ws.Name
    ws.Copy Before:=ws
    indexws = ws.Index
    ThisWorkbook.Sheets(indexws - 1).Name = currentWSName & "_Origin"
    ws.Select
    

End Sub

Sub finalSum()
    Dim sumcells As Range
    
    Set sumcells = Range("A1").Offset(3, 0).End(xlDown).Offset(1, 2)
    With sumcells
        .Formula = "=SUM(" & Range(sumcells.Offset(-1, 0).End(xlUp).Offset(1, 0), sumcells.Offset(-1, 0)).Address & ")"
        .Offset(-1, 0).Copy
        .PasteSpecial Paste:=xlPasteFormats
        .Offset(0, -1).Value = "Total"
    End With
End Sub

