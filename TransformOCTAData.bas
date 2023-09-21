Attribute VB_Name = "TransformOCTAData"
Option Explicit
Sub TransformOCTAData()
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    ' Prompt the user to confirm data formatting
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Please confirm that the data is formatted correctly before continuing. Each subject should have four rows of data, with scans in order:" & _
    vbCrLf & "- OD 3x3" & Chr(10) & "- OD 6x6" & Chr(10) & "- OS 3x3" & Chr(10) & "- OS 6x6", vbQuestion + vbYesNo, "Data Formatting Confirmation")

    ' Check the user's response
    If confirmation <> vbYes Then
        Exit Sub ' Exit the macro if the user clicks No or cancels
    End If

    Dim subjectIDColumn As Long
    Dim previousSubjectID As Variant

    ' Set the column number for the subject ID column
    subjectIDColumn = 2

    ' Set the active worksheet as the target worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim currentRow As Long
    currentRow = 3 ' Start from row 3 to avoid row 2 being appedned to header row

    ' Initialize the previousSubjectID variable with the first subject ID
    previousSubjectID = ws.Cells(2, subjectIDColumn).Value

    ' Loop through each row from top to bottom
    Do While ws.Cells(currentRow, subjectIDColumn).Value <> "" ' Continue until an empty cell is encountered
        ' Check if the current subject ID matches the previous one
        If ws.Cells(currentRow, subjectIDColumn).Value = previousSubjectID Then
            ' Find the last column in the previous row
            Dim lastColumn As Long
            lastColumn = ws.Cells(currentRow - 1, ws.Columns.Count).End(xlToLeft).Column

            ' Get the range of the current row from column 1 to the last used column
            Dim copyRange As Range
            Set copyRange = ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, ws.Columns.Count).End(xlToLeft))

            ' Loop through each cell in the copy range and append the values to the previous row
            Dim cell As Range
            For Each cell In copyRange.Cells
                lastColumn = lastColumn + 1
                ws.Cells(currentRow - 1, lastColumn).Value = cell.Value
            Next cell

            ' Delete the current row
            ws.Rows(currentRow).Delete
        Else
            ' Update the previousSubjectID variable if it doesn't match
            previousSubjectID = ws.Cells(currentRow, subjectIDColumn).Value
            currentRow = currentRow + 1 ' Move to the next row
        End If
    Loop

    ' Copy the headers from A1:BP1 and paste them starting in BQ1
    ws.Range("A1:BP1").Copy
    ws.Range("BQ1").PasteSpecial Paste:=xlPasteValues
    
    'Paste again starting in EG1
    ws.Range("A1:BP1").Copy
    ws.Range("EG1").PasteSpecial Paste:=xlPasteValues
    
    'Paste again starting in GW1
    ws.Range("A1:BP1").Copy
    ws.Range("GW1").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False ' Clear the clipboard

    ' Add variable names with different endings
    Dim variableNames As Range
    Set variableNames = ws.Range("A1:JL1") ' Adjust the range as needed

    Dim newColumn As Long
    Dim columnName As String

    For newColumn = 1 To variableNames.Columns.Count
        columnName = variableNames.Cells(1, newColumn).Value
        Dim columnIndex As Long
        columnIndex = newColumn ' The column index starts from 1

        ' Append the appropriate suffix based on the columnIndex
        If columnIndex >= 1 And columnIndex <= 68 Then
            columnName = columnName & "_OD3x3"
        ElseIf columnIndex >= 69 And columnIndex <= 136 Then
            columnName = columnName & "_OD6x6"
        ElseIf columnIndex >= 137 And columnIndex <= 204 Then
            columnName = columnName & "_OS3x3"
        ElseIf columnIndex >= 205 And columnIndex <= 272 Then
            columnName = columnName & "_OS6x6"
        End If

        ws.Cells(1, newColumn).Value = columnName
    Next newColumn
    
    ' Find the last used row in column F
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Apply conditional formatting to highlight cells that indicate a data transformation error
    Dim rng As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim rng4 As Range
    Dim rng5 As Range
    Dim rng6 As Range
    Dim rng7 As Range
    Dim rng8 As Range
    
    Set rng = ws.Range("F2:F" & lastRow)
    Set rng2 = ws.Range("H2:H" & lastRow)
    Set rng3 = ws.Range("BV2:BV" & lastRow)
    Set rng4 = ws.Range("BX2:BX" & lastRow)
    Set rng5 = ws.Range("EL2:EL" & lastRow)
    Set rng6 = ws.Range("EN2:EN" & lastRow)
    Set rng7 = ws.Range("HB2:HB" & lastRow)
    Set rng8 = ws.Range("HD2:HD" & lastRow)

    ' Clear any previous conditional formatting rules
    rng.FormatConditions.Delete
    rng2.FormatConditions.Delete
    rng3.FormatConditions.Delete
    rng4.FormatConditions.Delete
    rng5.FormatConditions.Delete
    rng6.FormatConditions.Delete
    rng7.FormatConditions.Delete
    rng8.FormatConditions.Delete

    ' Add new conditional formatting rules to highlight cells that indicate data was transposed to the wrong place
    
    Dim condition As FormatCondition
    Dim condition2 As FormatCondition
    Dim condition3 As FormatCondition
    Dim condition4 As FormatCondition
    Dim condition5 As FormatCondition
    Dim condition6 As FormatCondition
    Dim condition7 As FormatCondition
    Dim condition8 As FormatCondition
    
    Set condition = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="OD")
    condition.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    Set condition2 = rng2.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="Angiography 3x3 mm")
    condition2.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    Set condition3 = rng3.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="OD")
    condition3.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    Set condition4 = rng4.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="Angiography 6x6 mm")
    condition4.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    Set condition5 = rng5.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="OS")
    condition5.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    Set condition6 = rng6.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="Angiography 3x3 mm")
    condition6.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    Set condition7 = rng7.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="OS")
    condition7.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    Set condition8 = rng8.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="Angiography 6x6 mm")
    condition8.Interior.Color = RGB(255, 0, 0) ' Set the highlight color (red)
    
    ' Resize the cells to the width of the header
    ws.Cells.EntireColumn.AutoFit
    
    ' Go to cell A1
    Application.Goto Reference:=ws.Range("A1")
    
    ' Resets ScreenUpdating toggle
    Application.ScreenUpdating = True
End Sub

