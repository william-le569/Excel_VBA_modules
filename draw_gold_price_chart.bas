Sub CreateChartFromSelection()
    Dim ws As Worksheet
    Dim rng As Range
    Dim chartObj As ChartObject
    ' Dim chartType As XlChartType
    
    ' Set the active sheet and selected range
    Set ws = ActiveSheet
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    
    ' Check if selection is valid
    If rng Is Nothing Then
        MsgBox "Please select a valid data range!", vbExclamation, "Selection Error"
        Exit Sub
    End If
    
    ' Add a new chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=300)
    
    ' Set chart type (change as needed)
    chartObj.Chart.chartType = xlLine ' Change xlLine to another type if needed
    
    ' Set data source
    chartObj.Chart.SetSourceData Source:=rng
    
    ' Set titles (optional)
    With chartObj.Chart
        .HasTitle = True
        .ChartTitle.Text = "Generated Chart"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "X Axis"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Y Axis"
    End With
    
    MsgBox "Chart created successfully!", vbInformation, "Done"
End Sub
