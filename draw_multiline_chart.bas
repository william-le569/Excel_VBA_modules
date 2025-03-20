Sub CreateMultiLineChart()
    Dim ws As Worksheet
    Dim rng As Range
    Dim chartObj As ChartObject
    Dim i As Integer
    
    ' Xác định trang tính đang hoạt động
    Set ws = ActiveSheet
    
    ' Kiểm tra xem người dùng đã chọn vùng dữ liệu hay chưa
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Vui lòng chọn vùng dữ liệu có nhiều cột!", vbExclamation, "Lỗi Chọn Dữ Liệu"
        Exit Sub
    End If
    
    ' Tạo biểu đồ mới
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=500, Height:=350)
    
    ' Cấu hình biểu đồ
    With chartObj.Chart
        .ChartType = xlLine ' Chọn biểu đồ đường
        .HasTitle = True
        .ChartTitle.Text = "Multi-Line Chart"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "X Axis"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Y Axis"
    End With
    
    ' Thêm từng series vào biểu đồ
    For i = 2 To rng.Columns.Count
        With chartObj.Chart.SeriesCollection.NewSeries
            .XValues = rng.Columns(1) ' Chọn cột đầu tiên làm trục X
            .Values = rng.Columns(i) ' Cột hiện tại làm dữ liệu Y
            '.Name = rng.Cells(1, i).Value ' Lấy tiêu đề từ hàng đầu tiên
        End With
    Next i
    
    MsgBox "Biểu đồ nhiều đường đã được tạo thành công!", vbInformation, "Hoàn tất"
End Sub