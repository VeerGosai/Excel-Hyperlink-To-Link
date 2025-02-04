Attribute VB_Name = "Module2"
Sub RemoveDuplicatesColumns()
    Dim ws As Worksheet
    Dim lastCol As String
    Dim lastColNum As Integer
    Dim rng As Range
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Ask for the last column to process
    lastCol = InputBox("Enter the last column letter to process (e.g., D):", "Specify Last Column")
    If lastCol = "" Then Exit Sub ' Exit if no input
    
    ' Convert column letter to number
    lastColNum = Columns(lastCol).Column
    
    ' Remove duplicates in each column individually from B to the specified column
    For i = 2 To lastColNum
        Set rng = ws.Range(ws.Cells(1, i), ws.Cells(ws.Cells(Rows.Count, i).End(xlUp).Row, i))
        rng.RemoveDuplicates Columns:=1, Header:=xlYes
    Next i
    
    MsgBox "Duplicate removal complete!", vbInformation
End Sub

