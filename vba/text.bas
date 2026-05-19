Sub TextSplitCharacter()
    ' 获取 PowerPoint 应用程序的活动对象
    Dim pptApp As Object
    Set pptApp = Application

    ' 检查是否有选中的形状，如果没有，则返回
    If pptApp.ActiveWindow.Selection.Type < 2 Then Exit Sub
    
    ' 获取选中的形状范围
    Dim selShapeRange As ShapeRange
    Set selShapeRange = pptApp.ActiveWindow.Selection.ShapeRange

    ' 获取选中的幻灯片
    Dim selSlide As Slide
    Set selSlide = pptApp.ActiveWindow.Selection.SlideRange(1)
    
    ' 遍历选中的每个形状
    Dim shp As Shape
    For Each shp In selShapeRange
        ' 检查形状是否有文本框，如果没有，则继续下一个形状
        If Not shp.HasTextFrame Then GoTo NextShape
        
        ' 检查文本框是否有文本，如果没有，则继续下一个形状
        If Not shp.TextFrame.HasText Then GoTo NextShape
        
        ' 遍历文本框中的每个字符
        Dim i As Long
        For i = 1 To shp.TextFrame2.TextRange.Characters.Count
            Dim rng As TextRange2
            Set rng = shp.TextFrame2.TextRange.Characters(i)
            
            ' 复制形状
            Dim shpDuplicate As Shape
            Set shpDuplicate = shp.Duplicate
            
            ' 将复制的形状的文本设置为当前字符
            shpDuplicate.TextFrame.TextRange.Text = rng.Text
            
            ' 设置复制的形状的位置和大小与当前字符相同
            shpDuplicate.Left = rng.BoundLeft
            shpDuplicate.Top = rng.BoundTop
            shpDuplicate.Width = rng.BoundWidth
            shpDuplicate.Height = rng.BoundHeight
        Next i
        
        ' 隐藏原始形状
        shp.Visible = msoFalse
        
NextShape:
    Next shp
    
End Sub
