Sub ProgressEvaluation()
    Dim i As Integer
    Dim j As Integer
    Dim ws As Worksheet
    Dim rng As Range
    Dim numberOfRows As Long
    Dim numberOfColumns As Long
    Dim cellValue As Variant
    Dim Count As Long
    Dim result As String
    Dim specialChar As String
    Dim startCell As Range
    Dim endCell As Range
    Dim rangeStr As String
    
    Set ws = ActiveSheet
    Set rng = Selection
    
    numberOfRows = rng.Row
    numberOfColumns = rng.Column
    
    Debug.Print numberOfRows
    Debug.Print numberOfColumns
    
    Set startCell = ws.Cells(numberOfRows, 1)
    Set endCell = ws.Cells(numberOfRows, numberOfColumns - 1)
    
    rangeStr = startCell.Address & ":" & endCell.Address
    
    
    Count = 0
    result = ""
    specialChar = ChrW(9608)
    
    For i = 1 To numberOfRows
        For j = 1 To numberOfColumns
            cellValue = ws.Cells(i, j).Value
            If cellValue = True Then
                Count = Count + 1
            End If
        Next j
    Next i
    
    For i = 1 To Count
        result = result & specialChar
    Next i
    
    ws.Cells(numberOfRows, numberOfColumns).Formula = "=REPT(""" & specialChar & """,COUNTIF(" & rangeStr & ",TRUE))"
    

End Sub
