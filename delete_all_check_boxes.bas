Sub DeleteAllCheckboxes()
    Dim chk As Shape
    For Each chk In ActiveSheet.Shapes
        If chk.Type = msoFormControl Then
            If chk.FormControlType = xlCheckBox Then
                chk.Select Replace:=False
            End If
        End If
    Next chk
    Selection.Delete
End Sub