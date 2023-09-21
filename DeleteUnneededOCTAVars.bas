Attribute VB_Name = "DeleteUnneededOCTAVars"
Option Explicit

Sub ExtractSubsetOCTAVars()

    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
' Prompt the user to confirm that data file has been processed so there's one line per subject
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Please confirm that the data is formatted correctly before continuing. There should be a header row, then one row of data per subject" & _
    vbCrLf, vbQuestion + vbYesNo, "Data Formatting Confirmation")

    ' Check the user's response
    If confirmation <> vbYes Then
        Exit Sub ' Exit the macro if the user clicks No or cancels
    End If

    ' Set the active worksheet as the target worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Define an array of column letters or indices to delete
    Dim columnsToDelete As Variant
    columnsToDelete = Array("A")
    
    ' Delete all unnecessary columns
   
    ws.Columns("A:A").Delete Shift:=xlToLeft
    ws.Columns("B:P").Delete Shift:=xlToLeft
    ws.Columns("BA:BQ").Delete Shift:=xlToLeft
    ws.Columns("CZ:DP").Delete Shift:=xlToLeft
    ws.Columns("EY:FO").Delete Shift:=xlToLeft
    
    ws.Columns("G:G").Delete Shift:=xlToLeft
    ws.Columns("M:P").Delete Shift:=xlToLeft
    ws.Columns("O:O").Delete Shift:=xlToLeft
    ws.Columns("U:X").Delete Shift:=xlToLeft
    ws.Columns("AD:AG").Delete Shift:=xlToLeft
    ws.Columns("AI:AI").Delete Shift:=xlToLeft
    ws.Columns("AK:AK").Delete Shift:=xlToLeft
    ws.Columns("CF:CF").Delete Shift:=xlToLeft
    ws.Columns("CH:CH").Delete Shift:=xlToLeft
    ws.Columns("CM:CM").Delete Shift:=xlToLeft
    ws.Columns("CS:CV").Delete Shift:=xlToLeft
    ws.Columns("DB:DE").Delete Shift:=xlToLeft
    ws.Columns("DK:DN").Delete Shift:=xlToLeft
    ws.Columns("DP:DP").Delete Shift:=xlToLeft
    ws.Columns("DR:DR").Delete Shift:=xlToLeft
    ws.Columns("FP:FP").Delete Shift:=xlToLeft
    ws.Columns("FJ:FJ").Delete Shift:=xlToLeft
    ws.Columns("FL:FL").Delete Shift:=xlToLeft
    ws.Columns("DM:DM").Delete Shift:=xlToLeft
    ws.Columns("CU:CU").Delete Shift:=xlToLeft
    
    ' Message box to save as a new file
    Dim reminder As VbMsgBoxResult
    reminder = MsgBox("Please check data was processed correctly and save as a new file.", vbOKOnly, "Reminder")
    
    ' Go to cell A1
    Application.Goto Reference:=ws.Range("A1")
    
     Application.ScreenUpdating = False

End Sub
