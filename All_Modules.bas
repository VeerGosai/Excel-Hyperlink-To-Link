Sub SortColumns()
    Dim ws As Worksheet
    Dim lastCol As String
    Dim lastColNum As Integer
    Dim rng As Range
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Ask for the last column to sort until
    lastCol = InputBox("Enter the last column letter to sort until (e.g., D):", "Specify Last Column")
    If lastCol = "" Then Exit Sub ' Exit if no input
    
    ' Convert column letter to number
    lastColNum = Columns(lastCol).Column
    
    ' Sort each column individually from B to the specified column
    For i = 2 To lastColNum
        Set rng = ws.Range(ws.Cells(1, i), ws.Cells(ws.Cells(Rows.Count, i).End(xlUp).Row, i))
        rng.Sort Key1:=rng, Order1:=xlAscending, Header:=xlYes
    Next i
    
    MsgBox "Sorting complete!", vbInformation
End Sub


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


Sub ExtractHyperlinks()
    Dim ws As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim cell As Range
    Dim targetCell As Range
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Let user select the source range
    On Error Resume Next
    Set sourceRange = Application.InputBox("Select the range containing hyperlinks:", Type:=8)
    If sourceRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Let user select the starting cell for output
    On Error Resume Next
    Set targetRange = Application.InputBox("Select the starting cell for output:", Type:=8)
    If targetRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Extract hyperlinks
    For Each cell In sourceRange
        If cell.Hyperlinks.Count > 0 Then
            Set targetCell = targetRange.Offset(cell.Row - sourceRange.Row, cell.Column - sourceRange.Column)
            targetCell.Value = cell.Hyperlinks(1).Address
        End If
    Next cell
    
    MsgBox "Hyperlink extraction complete!", vbInformation
End Sub

Sub ReplaceGoogleDriveLinks()
    Dim ws As Worksheet
    Dim sourceRange As Range
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Let user select the source range
    On Error Resume Next
    Set sourceRange = Application.InputBox("Select the range containing Google Drive links:", Type:=8)
    If sourceRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Replace text in each cell
    For Each cell In sourceRange
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) = False Then
            cell.Value = Replace(cell.Value, "https://drive.google.com/open?id=", "https://drive.google.com/embeddedfolderview?id=")
            cell.Value = Replace(cell.Value, "&usp=drive_copy", "#list")
        End If
    Next cell
    
    MsgBox "Replacement complete!", vbInformation
End Sub

Sub BuildTextDataset()
    Dim ws As Worksheet
    Dim sourceRange As Range
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Let user select the source range
    On Error Resume Next
    Set sourceRange = Application.InputBox("Select the range to process:", Type:=8)
    If sourceRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Step 1: Replace blank cells with "n/a"
    For Each cell In sourceRange
        If Trim(cell.Value) = "" Then
            cell.Value = "n/a"
        End If
    Next cell
    
    ' Step 2: Find and clear all cells containing "https"
    For Each cell In sourceRange
        If InStr(cell.Value, "https") > 0 Then
            cell.ClearContents
        End If
    Next cell
    
    ' Step 3: Replace remaining blank cells with "Final Exam"
    For Each cell In sourceRange
        If Trim(cell.Value) = "" Then
            cell.Value = "Final Exam"
        End If
    Next cell
    
    MsgBox "Dataset processing complete!", vbInformation
End Sub


