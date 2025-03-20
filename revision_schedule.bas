Sub CreateRevisionSchedule()
    Dim i As Integer
    Dim j As Integer
    Dim rowAddress As Long
    Dim columnAddress As Long
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim rng As Range
    Set rng = Selection
    
    Dim cell As Range
    
    rng.Value = Date  'Learning Day
    rng.NumberFormat = "MM-DD-YYYY"
    rng.Font.Color = vbBlack
    rowAddress = rng.Row
    columnAddress = rng.Column
    
    For i = 1 To 8
        Set cell = ws.Cells(rowAddress, columnAddress + i)
        Debug.Print "Cell Address:"
        Debug.Print cell.Row
        Debug.Print cell.Column
        If i < 3 Then
            cell.Value = ws.Cells(rowAddress, columnAddress).Value + i
            cell.NumberFormat = "MM-DD-YYYY"
            cell.Font.Color = vbBlack
        ElseIf i = 3 Then
            cell.Value = ws.Cells(rowAddress, columnAddress).Value + 5
            cell.NumberFormat = "MM-DD-YYYY"
            cell.Font.Color = vbBlack
        ElseIf i = 4 Then
            cell.Value = ws.Cells(rowAddress, columnAddress).Value + 12
            cell.NumberFormat = "MM-DD-YYYY"
            cell.Font.Color = vbBlack
        ElseIf i = 5 Then
            cell.Value = ws.Cells(rowAddress, columnAddress).Value + 26
            cell.NumberFormat = "MM-DD-YYYY"
            cell.Font.Color = vbBlack
        ElseIf i = 6 Then
            cell.Value = ws.Cells(rowAddress, columnAddress).Value + 56
            cell.NumberFormat = "MM-DD-YYYY"
            cell.Font.Color = vbBlack
        ElseIf i = 7 Then
            cell.Value = ws.Cells(rowAddress, columnAddress).Value + 116
            cell.NumberFormat = "MM-DD-YYYY"
            cell.Font.Color = vbBlack
        ElseIf i = 8 Then
            cell.Value = ws.Cells(rowAddress, columnAddress).Value + 146
            cell.NumberFormat = "MM-DD-YYYY"
            cell.Font.Color = vbBlack
        End If
    Next i
    
    Debug.Print "Row Address:"
    Debug.Print rng.Row
    Debug.Print "Column Address:"
    Debug.Print rng.Column
    
    Debug.Print rng.Rows.Count
    Debug.Print rng.Columns.Count
    
    Debug.Print "End Debug!"
    
    
End Sub
