Attribute VB_Name = "Module3"
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

