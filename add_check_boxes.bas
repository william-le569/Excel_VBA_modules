Sub AddCheckBox01()
    Dim i As Integer
    Dim j As Integer
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim rng As Range
    Set rng = Selection

    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            Set cell = rng.Cells(i, j)
            If i Mod 2 = 1 Then 'Add Check Box
                ws.CheckBoxes.Add(cell.Left + 2 / 5 * cell.Width, _
                              cell.Top, _
                              cell.Width / 5, _
                              cell.Height * 0.9).Select
                With Selection
                    ' Clear the default text label of the checkbox
                    .Characters.Text = ""  ' ?? This might cause an error because Form Controls don't use .Characters.Text
                
                    ' Link the checkbox to the respective cell
                    ' - This means when the checkbox is checked, the cell will show TRUE
                    ' - When unchecked, the cell will show FALSE
                    .LinkedCell = cell.Address
                    cell.Font.Color = vbWhite
                End With
                
            ElseIf i Mod 2 = 0 Then
                cell.Font.Color = vbBlack
            End If
        Next j
    Next i
End Sub



