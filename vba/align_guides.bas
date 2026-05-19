Sub add_guides(is_master As Boolean, Optional h1 As Single = 105, _
               Optional h2 As Single = 487, Optional v1 As Single = 67, _
               Optional v2 As Single = 893)
    ' 添加参考线
    Application.DisplayGuides = True
    Dim guides As Guides
    If is_master Then
        Set guides = Application.ActivePresentation.SlideMaster.Guides
    Else
        Set guides = Application.ActivePresentation.Guides
    End If
    ' 添加新的参考线
    guides.Add ppHorizontalGuide, h1  ' 水平参考线，位置 h1
    guides.Add ppHorizontalGuide, h2  ' 水平参考线，位置 h2
    guides.Add ppVerticalGuide, v1    ' 垂直参考线，位置 v1
    guides.Add ppVerticalGuide, v2    ' 垂直参考线，位置 v2
End Sub

Sub delete_guides(is_master As Boolean)
    ' 删除所有参考线
    Application.DisplayGuides = True
    Dim guides As Guides
    If is_master Then
        Set guides = Application.ActivePresentation.SlideMaster.Guides
    Else
        Set guides = Application.ActivePresentation.Guides
    End If
    Do While guides.Count > 0
        guides.Item(1).Delete
    Loop
End Sub

Sub align_group_to_guides(alignment As String, Optional h1 As Single = 105, _
                          Optional h2 As Single = 487, Optional v1 As Single = 67, _
                          Optional v2 As Single = 893)
    ' 对齐形状组到参考线
    Application.DisplayGuides = True
    If Application.ActiveWindow.Selection.Type < 2 Then Exit Sub
    Dim shape_range As ShapeRange
    Set shape_range = Application.ActiveWindow.Selection.ShapeRange
    Dim guides As Guides
    Set guides = Application.ActivePresentation.Guides

    ' 添加默认参考线
    If guides.Count < 4 Then
        add_guides False, h1, h2, v1, v2
    End If

    ' 计算选中形状集合的边界
    Dim leftmost As Single, rightmost As Single, topmost As Single, bottommost As Single
    Call calculate_boundaries(shape_range, leftmost, rightmost, topmost, bottommost)

    Dim move_x As Single, move_y As Single, scale_x As Single, scale_y As Single
    move_x = 0
    move_y = 0
    scale_x = 1
    scale_y = 1

    Select Case alignment
        Case "expand"
            Dim left_guide As Single, right_guide As Single, top_guide As Single, bottom_guide As Single
            left_guide = guides(3).Position
            right_guide = guides(4).Position
            top_guide = guides(1).Position
            bottom_guide = guides(2).Position
            move_x = left_guide - leftmost
            move_y = top_guide - topmost
            scale_x = (right_guide - left_guide) / (rightmost - leftmost)
            scale_y = (bottom_guide - top_guide) / (bottommost - topmost)
        Case "expand_width"
            left_guide = guides(3).Position
            right_guide = guides(4).Position
            move_x = left_guide - leftmost
            scale_x = (right_guide - left_guide) / (rightmost - leftmost)
        Case "expand_height"
            top_guide = guides(1).Position
            bottom_guide = guides(2).Position
            move_y = top_guide - topmost
            scale_y = (bottom_guide - top_guide) / (bottommost - topmost)
        Case "left"
            move_x = guides(3).Position - leftmost
        Case "center"
            move_x = (guides(3).Position + guides(4).Position) / 2 - (leftmost + rightmost) / 2
        Case "right"
            move_x = guides(4).Position - rightmost
        Case "top"
            move_y = guides(1).Position - topmost
        Case "middle"
            move_y = (guides(1).Position + guides(2).Position) / 2 - (topmost + bottommost) / 2
        Case "bottom"
            move_y = guides(2).Position - bottommost
    End Select

    ' 移动和缩放形状
    Dim shp As Shape
    For Each shp In shape_range
        shp.Left = (shp.Left - leftmost) * scale_x + leftmost + move_x
        shp.Top = (shp.Top - topmost) * scale_y + topmost + move_y
        shp.Width = shp.Width * scale_x
        shp.Height = shp.Height * scale_y
    Next shp
End Sub

Sub align_to_guide(alignment As String, Optional h1 As Single = 105, _
                   Optional h2 As Single = 487, Optional v1 As Single = 67, _
                   Optional v2 As Single = 893)
    ' 对齐单个形状到参考线
    Application.DisplayGuides = True
    If Application.ActiveWindow.Selection.Type <> 2 Then Exit Sub
    Dim shape_range As ShapeRange
    Set shape_range = Application.ActiveWindow.Selection.ShapeRange
    Dim guides As Guides
    Set guides = Application.ActivePresentation.Guides

    ' 添加默认参考线
    If guides.Count < 4 Then
        add_guides False, h1, h2, v1, v2
    End If

    Dim move_x As Single, move_y As Single
    move_x = 0
    move_y = 0

    Select Case alignment
        Case "left"
            move_x = guides(3).Position
        Case "center"
            move_x = (guides(3).Position + guides(4).Position) / 2
        Case "right"
            move_x = guides(4).Position
        Case "top"
            move_y = guides(1).Position
        Case "middle"
            move_y = (guides(1).Position + guides(2).Position) / 2
        Case "bottom"
            move_y = guides(2).Position
    End Select

    ' 移动形状
    Dim shp As Shape
    For Each shp In shape_range
        If alignment = "left" Or alignment = "center" Or alignment = "right" Then
            shp.Left = move_x - (IIf(alignment = "center", shp.Width / 2, IIf(alignment = "right", shp.Width, 0)))
        ElseIf alignment = "top" Or alignment = "middle" Or alignment = "bottom" Then
            shp.Top = move_y - (IIf(alignment = "middle", shp.Height / 2, IIf(alignment = "bottom", shp.Height, 0)))
        End If
    Next shp
End Sub

Sub calculate_boundaries(shape_range As ShapeRange, ByRef leftmost As Single, _
                         ByRef rightmost As Single, ByRef topmost As Single, _
                         ByRef bottommost As Single)
    ' 计算选中形状集合的边界
    leftmost = 999999
    rightmost = -999999
    topmost = 999999
    bottommost = -999999
    Dim shp As Shape
    For Each shp In shape_range
        If shp.Left < leftmost Then leftmost = shp.Left
        If shp.Left + shp.Width > rightmost Then rightmost = shp.Left + shp.Width
        If shp.Top < topmost Then topmost = shp.Top
        If shp.Top + shp.Height > bottommost Then bottommost = shp.Top + shp.Height
    Next shp
End Sub

' 添加母版参考线
Sub add_master_guides(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                      Optional v1 As Single = 67, Optional v2 As Single = 893)
    add_guides True, h1, h2, v1, v2
End Sub

' 删除母版参考线
Sub delete_master_guides()
    delete_guides True
End Sub

' 添加参考线
Sub add_standard_guides(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                        Optional v1 As Single = 67, Optional v2 As Single = 893)
    add_guides False, h1, h2, v1, v2
End Sub

' 删除参考线
Sub delete_standard_guides()
    delete_guides False
End Sub

' 扩展形状组以匹配参考线
Sub expand_group_to_guides(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                           Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "expand", h1, h2, v1, v2
End Sub

' 扩展形状组宽度以匹配参考线
Sub expand_group_width_to_guides(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                                 Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "expand_width", h1, h2, v1, v2
End Sub

' 扩展形状组高度以匹配参考线
Sub expand_group_height_to_guides(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                                  Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "expand_height", h1, h2, v1, v2
End Sub

' 左对齐到参考线（单个形状）
Sub align_left_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                        Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_to_guide "left", h1, h2, v1, v2
End Sub

' 中心对齐到参考线（单个形状）
Sub align_center_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                          Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_to_guide "center", h1, h2, v1, v2
End Sub

' 右对齐到参考线（单个形状）
Sub align_right_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                         Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_to_guide "right", h1, h2, v1, v2
End Sub

' 顶部对齐到参考线（单个形状）
Sub align_top_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                       Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_to_guide "top", h1, h2, v1, v2
End Sub

' 中部对齐到参考线（单个形状）
Sub align_middle_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                          Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_to_guide "middle", h1, h2, v1, v2
End Sub

' 底部对齐到参考线（单个形状）
Sub align_bottom_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                          Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_to_guide "bottom", h1, h2, v1, v2
End Sub

' 左对齐到参考线（形状组）
Sub align_group_left_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                               Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "left", h1, h2, v1, v2
End Sub


' 中心对齐到参考线（形状组）
Sub align_group_center_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                                Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "center", h1, h2, v1, v2
End Sub

' 右对齐到参考线（形状组）
Sub align_group_right_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                               Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "right", h1, h2, v1, v2
End Sub

' 顶部对齐到参考线（形状组）
Sub align_group_top_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                             Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "top", h1, h2, v1, v2
End Sub

' 中部对齐到参考线（形状组）
Sub align_group_middle_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                                Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "middle", h1, h2, v1, v2
End Sub

' 底部对齐到参考线（形状组）
Sub align_group_bottom_to_guide(Optional h1 As Single = 105, Optional h2 As Single = 487, _
                                Optional v1 As Single = 67, Optional v2 As Single = 893)
    align_group_to_guides "bottom", h1, h2, v1, v2
End Sub
