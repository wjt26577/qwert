; change_all_fonts(fontName := "") {
;     ; --- 主程序 ---
;     try {
;         if (fontName = "") {
;             ; 1. 弹出输入框获取字体名称
;             input := InputBox("请输入要修改的字体名称：", "修改PPT字体")

;             ; 如果用户点击取消或输入为空，则退出
;             if (input.Value == "") {
;                 ExitApp
;             }
;             fontName := Trim(input.Value) ; 去除首尾空格
;         }

;         selection := get_ppt_object()

;         if (selection.success == false) {
;             MsgBox(selection.msg, "Error", "OK Iconx")
;             return
;         }

;         slide_range := selection.slide_range

;         for slide in slide_range {
;             for shape in slide.Shapes {
;                 if (shape.HasTextFrame) {
;                     textFrame := shape.TextFrame2
;                     if (textFrame.HasText) {
;                         textRange := textFrame.TextRange
;                         textRange.Font.Name := fontName
;                         textRange.Font.NameFarEast := fontName
;                         textRange.Font.Spacing := 0
;                     }
;                 }
;             }
;         }

;         MsgBox "字体已成功修改为：'" fontName "'", "完成", "Ok"

;     } catch as err {
;         MsgBox(err.message, "Error", "OK Iconx")
;     }

; }

; 获取selection
get_ppt_object2() {
    try {
        global selection_type_map, SHAPE_TYPE_MAP
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection

        if (selection_type_map.Has(selection.Type)) {
            sel_type := selection_type_map[selection.Type]
        } else {
            return 0
        }

        slide_range := selection.SlideRange

        if (sel_type == "none" || sel_type == "slides") {
            return { sel_type: sel_type, slide_range: slide_range, shape_range: "", shape_type: "" }
        }

        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange

        shape_type_key := shape_range(1).Type

        if (shape_type_key && SHAPE_TYPE_MAP.Has(shape_type_key)) {
            shape_type := SHAPE_TYPE_MAP[shape_type_key]
        } else {
            shape_type := ""
        }
        ;         msg: "当前没有活动的演示文稿窗口，请打开一个 PPT 文件后再试。"
        ;         msg: "无法连接到 PowerPoint。请确保 PowerPoint 已经打开并处于活动状态。"
        ;         msg: "无法连接到 PowerPoint。请确保 PowerPoint 已经打开并处于活动状态。"

        return { sel_type: sel_type, slide_range: slide_range, shape_range: shape_range, shape_type: shape_type }

    } catch as err {
        Msgbox(err.Message, "Error", "OK Iconx")
        return 0
    }
}

; --- 辅助函数：找出数组中所有最小值的索引 ---
; @param arr 数值数组
; @return Array 包含所有最小值索引的数组 (索引从 1 开始)
find_min_index(arr) {
    if (arr.Length == 0) {
        return []
    }

    ; 找出最小值
    min_val := Min(arr*)

    ; 收集所有等于最小值的索引
    indices := []
    for val in arr {
        if (val == min_val) {
            indices.Push(A_Index)
        }
    }

    return indices
}

find_min_in_min(a_arr, b_arr) {
    if (a_arr.Length == 0 || b_arr.Length == 0) {
        return []
    }

    a_min_indices := find_min_index(a_arr)

    b_candidates := []
    for idx in a_min_indices {
        b_candidates.Push(b_arr[idx])
    }
    min_b_val := Min(b_candidates*)

    candidates := []
    for idx in a_min_indices {
        candidates.Push({ val: b_arr[idx], idx: idx })
    }

    result := []
    for obj in candidates
        if (obj.val == min_b_val)
            result.Push(obj.idx)

    return result
}

test_find_min_in_min() {
    ; 模拟数据
    ; A数组: [10, 5, 5, 8]  -> 最小值是5，索引是 2 和 3
    ; B数组: [1,  3, 1, 2]  -> 在索引 2,3 中，B的最小值是1 (索引3)
    ; 期望结果: [3]

    a_arr := [10, 5, 5, 5, 5, 5, 8]
    b_arr := [1, 3, 1, 2, 3, 1, 2]

    result := find_min_in_min(a_arr, b_arr)

    str_result := ""
    for value in result {
        str_result := str_result . "," . value
    }

    MsgBox(str_result, , "T20")
}

; 简单的冒泡排序函数
bubble_sort(arr) {
    loop (arr.Length) {
        loop (arr.Length - A_Index) {
            j := A_Index
            if (arr[j] > arr[j + 1]) {
                temp := arr[j]
                arr[j] := arr[j + 1]
                arr[j + 1] := temp
            }
        }
    }
    return arr
}

; 获取当前选区的详细信息
get_ppt_object3() {
    global selection_type_map, SHAPE_TYPE_MAP

    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection
    } catch as err {
        return { success: false, msg: err.message }
    }

    ; 检查当前演示文稿中是否有幻灯片
    if (!ppt_app.ActivePresentation.Slides.Count) {
        return {
            success: false,
            msg: "不能获取选择对象，因为幻灯片数量为零。"
        }
    }

    ; 根据预定义的映射表获取选区类型（如 "shapes", "slides", "none"）
    selection_type := selection_type_map[selection.Type]

    ; 初始化结果对象，包含成功标志、选区类型、幻灯片和形状计数等
    result := {
        success: true,
        selection_type: selection_type,
        slide_count: 0,
        shape_count: 0,
        slide_range: "",
        shape_range: "",
        shape_type: "",
        base_shape: "",
        left: 0,
        top: 0,
        width: 0,
        height: 0
    }

    ; 获取选中的幻灯片范围
    slide_range := selection.SlideRange
    result.slide_range := slide_range
    result.slide_count := slide_range.Count

    ; 如果选区类型为 "none" (无选中) 或 "slides" (仅选中幻灯片标签)，直接返回结果
    if (selection_type ~= "none|slides") {
        return result
    }

    ; 获取选中的形状范围
    ; 如果选中的是组合内的形状，优先获取子形状范围
    if (selection.HasChildShapeRange)
        shape_range := selection.ChildShapeRange
    else
        shape_range := selection.ShapeRange

    ; 将形状范围和数量写入结果对象
    result.shape_range := shape_range
    result.shape_count := shape_range.Count

    ; 初始化包围盒的边界值，用于计算所有形状的整体范围
    min_left := 999999
    min_top := 999999
    max_right := -999999
    max_bottom := -999999

    left_arr := []   ; 存储 shape.Left
    top_arr := []    ; 存储 shape.Top
    right_arr := []   ; 存储 shape.Right
    bottom_arr := []  ; 存储 shape.Bottom

    ; --- 2. 遍历生成数组 ---
    for shape in shape_range {
        left_arr.Push(shape.Left)
        top_arr.Push(shape.Top)
        right_arr.Push(shape.Left + shape.Width)
        bottom_arr.Push(shape.Top + shape.Height)
    }

    result.left := Min(left_arr*)
    result.top := Min(top_arr*)
    result.width := result.left + Max(right_arr*)
    result.height := result.top + Max(bottom_arr*)

    base_shape_index := find_min_in_min(left_arr, top_arr)[1]
    base_shape := shape_range(base_shape_index)
    result.base_shape := base_shape
    ; 获取第一个形状的类型作为代表类型
    result.shape_type := SHAPE_TYPE_MAP[base_shape.Type]

    ; 返回最终结果
    return result
}

test_get_ppt_object2() {
    result := get_ppt_object()
    if (!result.success) {
        MsgBox(result.msg, "Error", "OK Iconx")
    } else {
        MsgBox("sucess", "sucess", "OK T2")
    }
}

test_get_ppt_object() {
    result := get_ppt_object()

    if (!result.success) {
        ; 显示错误描述
        MsgBox(result.msg, "Error", "OK Iconx")

    } else {
        ; 显示成功信息和统计数据
        summary := "选区类型: " result.selection_type
            . "`nshape_type: " result.shape_type
            . "`n幻灯片数: " result.slide_count
            . "`n形状数量: " result.shape_count
            . "`n边界框: (" result.left ", " result.top ") " result.width "x" result.height

        MsgBox("✅ 获取选区成功`n`n" summary, "Success", "Ok")
    }
}

; { 调整字体大小 --------------------------------------------------------------------------------

; 格式化字体大小显示
format_font_size(size) {
    ; 如果是整数，显示为整数；否则显示小数
    if (Mod(size, 1) == 0) {
        return Round(size, 0)
    } else {
        return size
    }
}

; 增大字体
increase_font() {
    shape_range := get_ppt_object().shape_range
    shape_range.Select
    run_ppt_command("FontSizeIncrease")
    font_size := shape_range.TextFrame2.TextRange.Font.Size
    Notify.show("Font Size: " . format_font_size(font_size))
}

; 减小字体
decrease_font() {
    shape_range := get_ppt_object().shape_range
    shape_range.Select
    run_ppt_command("FontSizeDecrease")
    font_size := shape_range.TextFrame2.TextRange.Font.Size
    Notify.show("Font Size: " . format_font_size(font_size))
}

; } 调整字体大小 ---------------------------------------------------------------------------------

selection_is_shapes() {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        if !ppt_application {
            return 0
        }
        selection := ppt_application.ActiveWindow.Selection
        type := selection.Type  ; 获取选择类型值
    } catch {
        return 0
    }
    ; 获取PowerPoint应用程序
    ; if !(ppt_application := ComObjActive("PowerPoint.Application") ) {
    ;     return 0
    ; }
    ; ppt_app := ComObjActive("PowerPoint.Application")
    ; ; 获取选择对象
    ; try {
    ;     selection := ppt_app.ActiveWindow.Selection
    ;     type := selection.Type  ; 获取选择类型值
    ; } catch {
    ;     return 0
    ; }
    type := 1
    ppt_application := ""
    ; 返回选择类型的文本描述
    switch type {
        case 0:
            return 0     ; 无选择
        case 1:
            return 0    ; 幻灯片选择
        case 2:
            return 1    ; 形状选择
        case 3:
            return 0     ; 文本选择
        default:
            return 0
    }
}

get_ppt_object_type2() {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        if !ppt_application {
            return 0
        }
        selection := ppt_application.ActiveWindow.Selection
        type := selection.Type  ; 获取选择类型值
    } catch {
        return 0
    }

    switch type {
        case 0:
            return "none"      ; 无选择
        case 1:
            return "slides"    ; 幻灯片选择
        case 2:
            return "shapes"    ; 形状选择
        case 3:
            return "text"      ; 文本选择
        default:
            return 0
    }
}

; selection_is_shapes() {
;     ; 获取PowerPoint应用程序
;     ; if !(ppt_application := ComObjActive("PowerPoint.Application") ) {
;     ;     return 0
;     ; }
;     ppt_app := ComObjActive("PowerPoint.Application")
;     ; 获取选择对象
;     try {
;         selection := ppt_app.ActiveWindow.Selection
;         type := selection.Type  ; 获取选择类型值
;     } catch {
;         return 0
;     }
;     ppt_application := ""
;     ; 返回选择类型的文本描述
;     Switch type {
;     Case 0:
;         return 0     ; 无选择
;     Case 1:
;         return 0    ; 幻灯片选择
;     Case 2:
;         return 1    ; 形状选择
;     Case 3:
;         return 0     ; 文本选择
;     default:
;         return 0
;     }
; }

; 获取幻灯片范围（选中的或当前的）
get_slide_range() {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        if (!ppt_application.ActivePresentation) {
            Msgbox("没有活动的演示文稿")
            return 0
        }

        selection := ppt_application.ActiveWindow.Selection
        return selection.SlideRange
    }
    catch {
        Msgbox("无法获取幻灯片")
        return 0
    }
}

; 获取当前激活的幻灯片
get_current_slide() {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        if (!ppt_application.ActivePresentation) {
            Msgbox("没有活动的演示文稿")
            return 0
        }

        ; 直接返回当前激活的幻灯片
        return ppt_application.ActiveWindow.View.Slide
    }
    catch {
        Msgbox("无法获取当前幻灯片")
        return 0
    }
}

; 获取形状范围
get_shape_range2() {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
    } catch {
        Msgbox("PowerPoint 未运行或无法访问")
        return 0
    }

    if (!ppt_app.ActivePresentation) {
        Msgbox("没有活动的演示文稿")
        return 0
    }

    sel := ppt_app.ActiveWindow.Selection

    ; 允许形状选择(2)和文本选择(3)
    if (sel.Type < 2) {
        return 0
    }
    if (sel.HasChildShapeRange) {
        return sel.ChildShapeRange
    } else {
        return sel.ShapeRange
    }
}

; 一次性获取所有PowerPoint对象的函数
; 返回一个对象，包含所有可获取的PowerPoint对象
;
; 主要属性说明:
; - active_slide: 当前正在编辑的幻灯片（单个）
; - selected_slides: 用户选中的幻灯片集合（在幻灯片浏览视图中，可能多个）
; - selected_shapes: 当前选中的形状集合

close_ppt_process() {
    if winactive("ahk_exe POWERPNT.EXE") {
        WinClose
        WinWaitClose(, , 5)
    }
    if WinExist("ahk_exe POWERPNT.EXE") {
        WinClose("ahk_exe POWERPNT.EXE")
        WinWaitClose(, , 5)
    }
    if ProcessExist("POWERPNT.EXE") {
        Result := MsgBox("是否强行关闭PowerPoint程序?", "Warning", "YesNo")
        if Result = "Yes"
            ProcessClose("POWERPNT.EXE")
        return
    }
    Msgbox("已经没有正在运行的PowerPoint进程了。", 2)
}

get_all_ppt_objects() {
    ; 创建返回对象
    ppt_obj := {}

    try {
        ; 获取PowerPoint应用程序
        ppt_obj.application := ComObjActive("PowerPoint.Application")

        ; 获取活动演示文稿
        try {
            ppt_obj.active_presentation := ppt_obj.application.ActivePresentation
        } catch {
            ppt_obj.active_presentation := ""
        }

        ; 获取活动窗口
        try {
            ppt_obj.active_window := ppt_obj.application.ActiveWindow

            ; 获取活动视图
            try {
                ppt_obj.active_view := ppt_obj.active_window.View

                ; 获取活动幻灯片
                try {
                    ppt_obj.active_slide := ppt_obj.active_view.Slide

                    ; 获取活动幻灯片上的形状集合
                    try {
                        ppt_obj.active_slide_shapes := ppt_obj.active_slide.Shapes
                    } catch {
                        ppt_obj.active_slide_shapes := ""
                    }
                } catch {
                    ppt_obj.active_slide := ""
                    ppt_obj.active_slide_shapes := ""
                }
            } catch {
                ppt_obj.active_view := ""
                ppt_obj.active_slide := ""
                ppt_obj.active_slide_shapes := ""
            }

            ; 获取活动选择
            try {
                ppt_obj.active_selection := ppt_obj.active_window.Selection

                ; 获取选择类型
                try {
                    selection_type := ppt_obj.active_selection.Type
                    switch selection_type {
                        case 0: ppt_obj.selection_type := "none"
                        case 1: ppt_obj.selection_type := "slides"
                        case 2: ppt_obj.selection_type := "shapes"
                        case 3: ppt_obj.selection_type := "text"
                        default: ppt_obj.selection_type := "unknown"
                    }
                } catch {
                    ppt_obj.selection_type := "error"
                }

                ; 获取选中的幻灯片集合
                try {
                    ppt_obj.selected_slides := ppt_obj.active_selection.SlideRange
                } catch {
                    ppt_obj.selected_slides := ""
                }

                ; 获取选中的形状集合
                try {
                    if WinActive("ahk_class PPTFrameClass") {
                        if (ppt_obj.active_selection.HasChildShapeRange) {
                            shape_range := ppt_obj.active_selection.ChildShapeRange
                        } else {
                            shape_range := ppt_obj.active_selection.ShapeRange
                        }
                        ppt_obj.selected_shapes := shape_range

                        ; 获取形状范围类型和数量
                        try {
                            ppt_obj.shape_range_type := shape_range.Type
                            ppt_obj.shape_range_count := shape_range.Count
                        } catch {
                            ppt_obj.shape_range_type := 0
                            ppt_obj.shape_range_count := 0
                        }
                    } else {
                        ppt_obj.selected_shapes := ""
                        ppt_obj.shape_range_type := 0
                        ppt_obj.shape_range_count := 0
                    }
                } catch {
                    ppt_obj.selected_shapes := ""
                    ppt_obj.shape_range_type := 0
                    ppt_obj.shape_range_count := 0
                }
            } catch {
                ppt_obj.active_selection := ""
                ppt_obj.selection_type := "error"
                ppt_obj.selected_slides := ""
                ppt_obj.selected_shapes := ""
                ppt_obj.shape_range_type := 0
                ppt_obj.shape_range_count := 0
            }

            ; 获取活动视图类型
            try {
                view_type := ppt_obj.active_window.ViewType
                switch view_type {
                    case 1: ppt_obj.active_view_type := "普通视图"
                    case 2: ppt_obj.active_view_type := "大纲视图"
                    case 3: ppt_obj.active_view_type := "幻灯片视图"
                    case 4: ppt_obj.active_view_type := "幻灯片放映视图"
                    case 5: ppt_obj.active_view_type := "幻灯片浏览视图"
                    case 6: ppt_obj.active_view_type := "备注页视图"
                    case 7: ppt_obj.active_view_type := "母版视图"
                    case 8: ppt_obj.active_view_type := "备注母版视图"
                    case 9: ppt_obj.active_view_type := "页眉/页脚视图"
                    case 10: ppt_obj.active_view_type := "幻灯片母版视图"
                    case 11: ppt_obj.active_view_type := "备注母版视图"
                    case 12: ppt_obj.active_view_type := "大纲母版视图"
                    default: ppt_obj.active_view_type := "未知视图(" . view_type . ")"
                }
            } catch {
                ppt_obj.active_view_type := "error"
            }
        } catch {
            ppt_obj.active_window := ""
            ppt_obj.active_view := ""
            ppt_obj.active_slide := ""
            ppt_obj.active_slide_shapes := ""
            ppt_obj.active_selection := ""
            ppt_obj.selection_type := "error"
            ppt_obj.selected_slides := ""
            ppt_obj.selected_shapes := ""
            ppt_obj.shape_range_type := 0
            ppt_obj.shape_range_count := 0
            ppt_obj.active_view_type := "error"
        }

    } catch as err {
        Msgbox("获取PowerPoint对象失败: " . err.Message)
        return {}
    }

    return ppt_obj
}

; 使用示例:
; ppt := get_all_ppt_objects()
;
; ; 访问各个对象:
; app := ppt.application
; presentation := ppt.active_presentation
; window := ppt.active_window
; view := ppt.active_view
; slide := ppt.active_slide
; shapes := ppt.active_slide_shapes
; selection := ppt.active_selection
; selection_type := ppt.selection_type
; selected_slides := ppt.selected_slides
; selected_shapes := ppt.selected_shapes
; shape_type := ppt.shape_range_type
; shape_count := ppt.shape_range_count
; view_type := ppt.active_view_type

; 对象属性说明:
; ================
; active_slide: 当前正在编辑的幻灯片（单个对象）
; selected_slides: 用户在幻灯片浏览视图中选中的幻灯片集合（SlideRange对象，可能包含多个）
; selected_shapes: 当前选中的形状集合（ShapeRange对象）
; shape_range_type: 选中形状的类型
; shape_range_count: 选中形状的数量

; 辅助函数：显示所有获取到的对象信息
show_ppt_objects_info(ppt_obj := "") {
    ppt_obj := get_all_ppt_objects()
    info := "PowerPoint对象信息:`n`n"

    ; 遍历对象的所有属性
    for prop_name in ppt_obj.OwnProps() {
        value := ppt_obj.%prop_name%
        if (IsObject(value) && value != "") {
            info .= prop_name . ": [COM对象]`n"
        } else if (value == "") {
            info .= prop_name . ": [空]`n"
        } else {
            info .= prop_name . ": " . value . "`n"
        }
    }

    MsgBox(info, "PowerPoint对象信息", 0)
}

;  获取常用的PowerPoint对象
; 返回: 成功时返回{app, selection, selected_shapes, selected_slides}，失败时返回false
get_ppt_objects() {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection

        result := { app: ppt_app, selection: selection, selected_shapes: "", selected_slides: "" }

        ; 根据选择类型处理
        switch selection.Type {
            case 1:  ; 幻灯片选择
                try {
                    result.selected_slides := selection.SlideRange
                } catch {
                    result.selected_slides := ""
                }

            case 2:  ; 形状选择
                try {
                    result.selected_shapes := (selection.HasChildShapeRange) ? selection.ChildShapeRange : selection.ShapeRange
                } catch {
                    result.selected_shapes := ""
                }

            case 3:  ; 文本选择
                try {
                    result.selected_shapes := (selection.HasChildShapeRange) ? selection.ChildShapeRange : selection.ShapeRange
                } catch {
                    result.selected_shapes := ""
                }

            default:  ; 无选择或其他
                ; 保持空值
        }

        return result
    } catch as err {
        Msgbox("获取PowerPoint对象时发生错误: " . err.Message)
        return false
    }
}
; 计算形状集合的中心和边界
get_shapes_bounds(shape_range) {
    if (!shape_range || shape_range.Count < 1) {
        return 0
    }

    ; 初始化边界值
    left := 9999
    right := -9999
    top := 9999
    bottom := -9999

    ; 找到所有形状的整体边界
    for shape in shape_range {
        try {
            shape_left := shape.Left
            shape_top := shape.Top
            shape_right := shape_left + shape.Width
            shape_bottom := shape_top + shape.Height

            if (shape_left < left)
                left := shape_left
            if (shape_right > right)
                right := shape_right
            if (shape_top < top)
                top := shape_top
            if (shape_bottom > bottom)
                bottom := shape_bottom
        } catch {
            ; 忽略错误，继续处理其他形状
        }
    }

    ; 计算中心点
    center_x := left + ((right - left) / 2)
    center_y := top + ((bottom - top) / 2)

    return {
        left: left,
        top: top,
        right: right,
        bottom: bottom,
        width: right - left,
        height: bottom - top,
        center_x: center_x,
        center_y: center_y
    }
}

; 向右复制形状
copy_to_right(gap := 8) {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection

        ; 检查选择类型是否为形状
        if (selection.Type < 2) {  ; 2 = shapes
            Msgbox("请先选择要复制的形状")
            return
        }

        if (selection.HasChildShapeRange)
            selected_shape_range := selection.ChildShapeRange
        else
            selected_shape_range := selection.ShapeRange

        ; 获取形状边界
        bounds := get_shapes_bounds(selected_shape_range)
        if (!bounds) {
            Msgbox("无法获取形状边界")
            return
        }

        selection.Unselect()

        ; 复制每个选中的形状到右侧
        for shape in selected_shape_range {
            copy := shape.Duplicate()
            ; 保持相对位置，整体移动到右侧
            relative_pos := shape.Left - bounds.left
            copy.Left := bounds.right + gap + relative_pos
            copy.Top := shape.Top
            copy.Select(0)
        }
    } catch as err {
        Msgbox("向右复制时发生错误: " . err.Message)
    } finally {
        ppt_application := ""
    }
}

; 向左复制形状
copy_to_left(gap := 8) {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection

        ; 检查选择类型是否为形状
        if (selection.Type < 2) {  ; 2 = shapes
            Msgbox("请先选择要复制的形状")
            return
        }

        if (selection.HasChildShapeRange)
            selected_shape_range := selection.ChildShapeRange
        else
            selected_shape_range := selection.ShapeRange

        ; 获取形状边界
        bounds := get_shapes_bounds(selected_shape_range)
        if (!bounds) {
            Msgbox("无法获取形状边界")
            return
        }

        selection.Unselect()

        ; 复制每个选中的形状到左侧
        for shape in selected_shape_range {
            copy := shape.Duplicate()
            ; 保持相对位置，整体移动到左侧
            relative_pos := shape.Left - bounds.left
            copy.Left := bounds.left - gap - bounds.width + relative_pos
            copy.Top := shape.Top
            copy.Select(0)
        }
    } catch as err {
        Msgbox("向左复制时发生错误: " . err.Message)
    } finally {
        ppt_application := ""
    }
}

; 向上复制形状
copy_to_up(gap := 8) {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection

        ; 检查选择类型是否为形状
        if (selection.Type < 2) {  ; 2 = shapes
            Msgbox("请先选择要复制的形状")
            return
        }

        if (selection.HasChildShapeRange)
            selected_shape_range := selection.ChildShapeRange
        else
            selected_shape_range := selection.ShapeRange

        ; 获取形状边界
        bounds := get_shapes_bounds(selected_shape_range)
        if (!bounds) {
            Msgbox("无法获取形状边界")
            return
        }

        selection.Unselect()

        ; 复制每个选中的形状到上方
        for shape in selected_shape_range {
            copy := shape.Duplicate()
            ; 保持相对位置，整体移动到上方
            relative_pos := shape.Top - bounds.top
            copy.Left := shape.Left
            copy.Top := bounds.top - gap - bounds.height + relative_pos
            copy.Select(0)
        }
    } catch as err {
        Msgbox("向上复制时发生错误: " . err.Message)
    } finally {
        ppt_application := ""
    }
}

; 向下复制形状
copy_to_down(gap := 8) {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection

        ; 检查选择类型是否为形状
        if (selection.Type < 2) {
            Msgbox("请先选择要复制的形状")
            return
        }

        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange

        ; 获取形状边界
        bounds := get_shapes_bounds(shape_range)
        if (!bounds) {
            Msgbox("无法获取形状边界")
            return
        }

        selection.Unselect()

        ; 复制每个选中的形状到下方
        for shape in shape_range {
            copy := shape.Duplicate()
            ; 保持相对位置，整体移动到下方
            relative_pos := shape.Top - bounds.top
            copy.Left := shape.Left
            copy.Top := bounds.bottom + gap + relative_pos
            copy.Select(0)
        }
    } catch as err {
        Msgbox("向下复制时发生错误: " . err.Message)
    } finally {
        ppt_application := ""
    }
}

; 使用示例:
; copy_to_right()      ; 使用默认间距8像素
; copy_to_right(16)    ; 使用自定义间距16像素
; copy_to_left(5)      ; 向左复制，间距5像素
; copy_to_up(10)       ; 向上复制，间距10像素
; copy_to_down(12)     ; 向下复制，间距12像素

; 标准结构说明:
; 1. ppt_application := ComObjActive("PowerPoint.Application")
; 2. selection := ppt_application.ActiveWindow.Selection
; 3. if (selection.Type < 2) - 检查选择类型
; 4. HasChildShapeRange 的 if-else 处理
; 5. get_shapes_bounds() 计算边界
; 6. selection.Unselect() 取消选择
; 7. 复制逻辑，保持相对位置
; 8. finally 中清理 ppt_application

; 设置形状等大小（等宽且等高）- 增强版
set_equal_size2(shape_range := 0, ref_shape := 0, options := {}) {
    ; 如果没有提供形状范围，尝试获取当前选中的形状
    if !shape_range {
        try {
            selection := get_ppt_object()
            if !selection.shape_range {
                return false
            }
            shape_range := selection.shape_range
            ref_shape := selection.min_left_shape
        } catch as err {
            error_msg := "无法获取选中的形状: " . err.Message
            Msgbox(error_msg)
        }
    }

    ; 默认选项
    default_options := {
        preserve_aspect_ratio: false,
        width_only: false,
        height_only: false,
        min_size: 1,
        max_size: 9999
    }

    ; 合并选项
    for key, value in default_options.OwnProps() {
        if !options.HasProp(key)
            options.%key% := value
    }

    reference_width := 0
    reference_height := 0
    errors := []

    ; 获取参考尺寸
    ref_width := ref_shape.Width
    ref_height := ref_shape.Height

    ; 设置所有形状的尺寸
    success_count := 0
    total_count := 0

    for shape in shape_range {
        total_count++
        try {
            ; 保持宽高比选项
            if (options.preserve_aspect_ratio) {
                current_ratio := shape.Width / shape.Height
                reference_ratio := ref_width / ref_height

                if (current_ratio > reference_ratio) {
                    ; 以宽度为准
                    shape.Width := ref_width
                    shape.Height := ref_width / current_ratio
                } else {
                    ; 以高度为准
                    shape.Height := ref_height
                    shape.Width := ref_height * current_ratio
                }
            } else {
                ; 只设置宽度
                if (options.width_only && !options.height_only) {
                    shape.Width := ref_width
                }
                ; 只设置高度
                else if (options.height_only && !options.width_only) {
                    shape.Height := ref_height
                }
                ; 设置宽度和高度（默认行为）
                else if (!options.width_only && !options.height_only) {
                    shape.Width := ref_width
                    shape.Height := ref_height
                }
            }

            success_count++
        } catch as err {
            error_msg := "设置形状 " . total_count . " 尺寸时出错: " . err.Message
            errors.Push(error_msg)
            ; 可选择是否显示每个错误
            ; Msgbox(error_msg)
        }
    }

}

duplicate_aligned_to_origin(shape_range := 0, left_offset := -12, top_offset := -12) {
    try {
        if (!shape_range) {
            selection := get_ppt_object()
            shape_range := selection.shape_range
            if (!shape_range)
                return false
        }
        dup_range := shape_range.Duplicate()
        dup_range.Select(1)

        if (Abs(left_offset) >= 10000) {
            left_offset := selection.width + left_offset - 10012
        }

        if (Abs(top_offset) >= 10000) {
            top_offset := selection.height + top_offset - 10012
        }

        for shape in dup_range {
            shape.Left := shape.Left + left_offset
            shape.Top := shape.Top + top_offset

        }
        ; Msgbox("left_offset: " . shape_range2(1).Left . " `ntop_offset: " . shape_range2(1).Top)

        return true

    } catch as err {
        MsgBox(err.Message)
    }
}

duplicate_right_to_origin() {
    duplicate_aligned_to_origin(, left_offset := 10012, top_offset := -12)
}

duplicate_down_to_origin() {
    duplicate_aligned_to_origin(, left_offset := -12, top_offset := 10012)
}

; 设置等宽高 (同时设置宽和高)
; @param shape_range 要调整的形状范围
; @param ref_shape 参考形状
resize_by_ref(mode := "full", shape_range := 0, ref_shape := 0) {
    try {
        ; 获取全局选区
        if (!shape_range) {
            selection := get_ppt_object()

            shape_range := selection.shape_range
            if (!shape_range)
                return false
        }

        ; 确定参考形状
        if (!ref_shape)
            ref_shape := selection.base_shape

        ; 获取参考尺寸
        ref_width := ref_shape.Width
        ref_height := ref_shape.Height
        ref_type := ref_shape.Type

        ; 遍历设置 (同时设置宽和高)
        for shape in shape_range {
            try {
                if shape.Type != ref_type
                    continue
                if (mode ~= "i)^width|full")
                    shape.Width := ref_width
                if (mode ~= "i)^height|full")
                    shape.Height := ref_height
            } catch {
                continue
            }
        }
        return true

    } catch as err {
        MsgBox(err.Message)
    }
}

resize_by_ref_width() {
    resize_by_ref("width")
}

resize_by_ref_height() {
    resize_by_ref("height")
}

; ; 便捷函数 - 只设置宽度
; set_equal_width2(shape_range := 0, reference_shape := 0) {
;     return set_equal_size(shape_range, reference_shape, {width_only: true})
; }

; ; 便捷函数 - 只设置高度
; set_equal_height2(shape_range := 0, reference_shape := 0) {
;     return set_equal_size(shape_range, reference_shape, {height_only: true})
; }

; ; 便捷函数 - 保持宽高比
; set_equal_size_proportional(shape_range := 0, reference_shape := 0) {
;     return set_equal_size(shape_range, reference_shape, {preserve_aspect_ratio: true})
; }

; #region 😀锁定、隐藏形状

; { 😀😀😀锁定、隐藏形状 ---------------------------------------------------------------------------------

; 隐藏形状
hide_shape_range(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        if (ppt_app.ActiveWindow.Selection.HasChildShapeRange)
            ppt_app.ActiveWindow.Selection.ChildShapeRange.Visible := 0
        else
            ppt_app.ActiveWindow.Selection.ShapeRange.Visible := 0

        Notify.show("隐藏形状", , "500")

    } catch
        throw
}

show_shape_range(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        for slide in ppt_app.ActiveWindow.Selection.SlideRange {
            if !slide.Shapes.Count
                continue
            slide.Shapes.Range.Visible := -1
        }

        Notify.show("显示形状")

    } catch
        throw
}

toggle_hide_shape_range(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        selection := ppt_app.ActiveWindow.Selection
        selection_type := selection_type_map[Selection.Type]

        if selection_type == "shapes" {
            if (Selection.HasChildShapeRange)
                Selection.ChildShapeRange.Visible := 0
            else
                Selection.ShapeRange.Visible := 0

            Notify.show("隐藏形状", "info", "500")

        } else if selection_type ~= "i)^(slides|none)" {
            for slide in Selection.SlideRange {
                if !slide.Shapes.Count
                    continue
                slide.Shapes.Range.Visible := -1
            }

            Notify.show("显示形状", "success", "500")
        }
    } catch
        throw
}


; 锁定形状
lock_shape_range(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        if (ppt_app.ActiveWindow.Selection.HasChildShapeRange)
            ppt_app.ActiveWindow.Selection.ChildShapeRange.Locked := -1
        else
            ppt_app.ActiveWindow.Selection.ShapeRange.Locked := -1

        Notify.show("锁定形状")

    } catch
        throw
}

; 锁定形状
toggle_lock_shape_range(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        if ppt_app.ActiveWindow.Selection.HasChildShapeRange
            shape_range := ppt_app.ActiveWindow.Selection.ChildShapeRange
        else
            shape_range := ppt_app.ActiveWindow.Selection.ShapeRange

        if (shape_range.Locked == 0) {
            shape_range.Locked := -1
            Notify.show_red("锁定形状")
        }
        else {
            shape_range.Locked := 0
            Notify.show("解锁形状")
        }

    } catch
        throw
}

unlock_shape_range(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        if (ppt_app.ActiveWindow.Selection.HasChildShapeRange)
            ppt_app.ActiveWindow.Selection.ChildShapeRange.Locked := 0
        else
            ppt_app.ActiveWindow.Selection.ShapeRange.Locked := 0

        Notify.show("解锁形状")

    } catch
        throw
}

lock_slides_shapes(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        for slide in ppt_app.ActiveWindow.Selection.SlideRange {
            if !slide.Shapes.Count
                continue
            slide.Shapes.Range.Locked := -1
        }

        Notify.show("锁定形状")

    } catch
        throw
}

unlock_slides_shapes(ppt_app := "") {
    try {
        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        for slide in ppt_app.ActiveWindow.Selection.SlideRange {
            if !slide.Shapes.Count
                continue
            slide.Shapes.Range.Locked := 0
        }

        Notify.show("解锁形状")

    } catch
        throw
}

; } 锁定、隐藏形状 ---------------------------------------------------------------------------------

; #endregion

; #region 形状分割成多个独立形状

/**
 * 将PowerPoint形状水平分割成多个独立形状（列）
 * @param shape 要分割的PowerPoint形状对象
 * @param columns 要分割成的列数
 * @param gap_pts 形状之间的间隙(磅)
 * @param keep_original 是否保留原始形状(默认不保留)
 * @return 返回包含所有新创建形状的数组
 */
divide_shape_into_columns(shape, columns := 2, gap_pts := 8, keep_original := false) {
    if (columns < 2) {
        Msgbox("列数必须大于1")
        return 0
    }

    try {
        ; 获取形状所在幻灯片
        parent_slide := shape.Parent

        ; 获取形状的位置和尺寸
        shape_left := shape.Left
        shape_top := shape.Top
        shape_width := shape.Width
        shape_height := shape.Height

        ; 保存原始形状的属性
        shape_type := shape.Type
        shape_name := shape.Name

        ; 对于特殊形状可能需要替代方法
        if (shape_type != 1) {  ; 1 = msoAutoShape
            ; 尝试获取更多关于形状的信息
            shape_auto_shape_type := 0
            try shape_auto_shape_type := shape.AutoShapeType

            ; 这里可以添加特殊形状类型的处理
        }

        ; 计算每列的宽度(不含间隙)
        column_width := (shape_width - (gap_pts * (columns - 1))) / columns

        ; 创建新形状数组
        new_shapes := []

        ; 创建新的形状
        for i in generate_number_range(1, columns) {
            ; 计算新形状的x位置
            new_left := shape_left + (column_width + gap_pts) * (i - 1)

            ; 复制原始形状
            new_shape := shape.Duplicate()

            ; 设置新形状的位置和尺寸
            new_shape.Left := new_left
            new_shape.Top := shape_top
            new_shape.Width := column_width
            new_shape.Height := shape_height

            ; 添加到数组
            new_shapes.Push(new_shape)
        }

        ; 删除原始形状(如果不保留)
        if (!keep_original) {
            shape.Delete()
        }

        ; 返回新形状数组
        return new_shapes
    } catch as errrr {
        Msgbox("分割形状时出错: " . err.Message)
        return 0
    }
}

/**
 * 将PowerPoint形状垂直分割成多个独立形状（行）
 * @param shape 要分割的PowerPoint形状对象
 * @param rows 要分割成的行数
 * @param gap_pts 形状之间的间隙(磅)
 * @param keep_original 是否保留原始形状(默认不保留)
 * @return 返回包含所有新创建形状的数组
 */
divide_shape_into_rows(shape, rows := 2, gap_pts := 8, keep_original := false) {
    if (rows < 2) {
        Msgbox("行数必须大于1")
        return 0
    }

    try {
        ; 获取形状所在幻灯片
        parent_slide := shape.Parent

        ; 获取形状的位置和尺寸
        shape_left := shape.Left
        shape_top := shape.Top
        shape_width := shape.Width
        shape_height := shape.Height

        ; 保存原始形状的属性
        shape_type := shape.Type
        shape_name := shape.Name

        ; 计算每行的高度(不含间隙)
        row_height := (shape_height - (gap_pts * (rows - 1))) / rows

        ; 创建新形状数组
        new_shapes := []

        ; 创建新的形状
        for i in generate_number_range(1, rows) {
            ; 计算新形状的y位置
            new_top := shape_top + (row_height + gap_pts) * (i - 1)

            ; 复制原始形状
            new_shape := shape.Duplicate()

            ; 设置新形状的位置和尺寸
            new_shape.Left := shape_left
            new_shape.Top := new_top
            new_shape.Width := shape_width
            new_shape.Height := row_height

            ; 添加到数组
            new_shapes.Push(new_shape)
        }

        ; 删除原始形状(如果不保留)
        if (!keep_original) {
            shape.Delete()
        }

        ; 返回新形状数组
        return new_shapes
    } catch as errrr {
        Msgbox("分割形状时出错: " . err.Message)
        return 0
    }
}

/**
 * 将PowerPoint形状同时分割成行和列（创建网格）
 * @param shape 要分割的PowerPoint形状对象
 * @param rows 行数
 * @param columns 列数
 * @param gap_pts 形状之间的间隙(磅)
 * @param keep_original 是否保留原始形状(默认不保留)
 * @return 返回二维数组，包含所有网格单元格形状
 */
divide_shape_into_grid(shape, rows := 2, columns := 2, gap_pts := 8, keep_original := false) {
    if (rows < 1 || columns < 1) {
        Msgbox("行数和列数必须大于0")
        return 0
    }

    try {
        ; 获取形状的位置和尺寸
        shape_left := shape.Left
        shape_top := shape.Top
        shape_width := shape.Width
        shape_height := shape.Height

        ; 计算每个单元格的尺寸
        cell_width := (shape_width - (gap_pts * (columns - 1))) / columns
        cell_height := (shape_height - (gap_pts * (rows - 1))) / rows

        ; 创建二维网格数组
        grid_shapes := []

        ; 创建网格单元格
        for row in generate_number_range(1, rows) {
            row_shapes := []

            for col in generate_number_range(1, columns) {
                ; 计算新形状的位置
                new_left := shape_left + (cell_width + gap_pts) * (col - 1)
                new_top := shape_top + (cell_height + gap_pts) * (row - 1)

                ; 复制原始形状
                new_shape := shape.Duplicate()

                ; 设置新形状的位置和尺寸
                new_shape.Left := new_left
                new_shape.Top := new_top
                new_shape.Width := cell_width
                new_shape.Height := cell_height

                ; 添加到行数组
                row_shapes.Push(new_shape)
            }

            ; 添加行到网格数组
            grid_shapes.Push(row_shapes)
        }

        ; 删除原始形状(如果不保留)
        if (!keep_original) {
            shape.Delete()
        }

        ; 返回网格数组
        return grid_shapes
    } catch as errrr {
        Msgbox("创建形状网格时出错: " . err.Message)
        return 0
    }
}

; 辅助函数：生成数字范围
generate_number_range(start, end, step := 1) {
    result := []
    current := start

    while (current <= end) {
        result.Push(current)
        current += step
    }

    return result
}

divide_shape(direction := "rows", divisions := 2) {
    ; 获取当前选中的形状
    shape_range := get_ppt_object().shape_range
    if (!shape_range || shape_range.Count == 0) {
        Msgbox("请先选择一个形状")
        return
    }

    if (direction == "columns") {
        for shape in shape_range {
            column_shapes := divide_shape_into_columns(shape, divisions)
        }
        return
    }

    if (direction == "rows") {
        for shape in shape_range {
            row_shapes := divide_shape_into_rows(shape, divisions)
        }
        return
    }

}

divide_shape_into_2rows(direction := "rows", divisions := 2) {
    divide_shape(direction, divisions)
}

divide_shape_into_3rows(direction := "rows", divisions := 3) {
    divide_shape(direction, divisions)
}

divide_shape_into_4rows(direction := "rows", divisions := 4) {
    divide_shape(direction, divisions)
}

divide_shape_into_5rows(direction := "rows", divisions := 5) {
    divide_shape(direction, divisions)
}

divide_shape_into_6rows(direction := "rows", divisions := 6) {
    divide_shape(direction, divisions)
}

divide_shape_into_2columns(direction := "columns", divisions := 2) {
    divide_shape(direction, divisions)
}

divide_shape_into_3columns(direction := "columns", divisions := 3) {
    divide_shape(direction, divisions)
}

divide_shape_into_4columns(direction := "columns", divisions := 4) {
    divide_shape(direction, divisions)
}

divide_shape_into_5columns(direction := "columns", divisions := 5) {
    divide_shape(direction, divisions)
}

divide_shape_into_6columns(direction := "columns", divisions := 6) {
    divide_shape(direction, divisions)
}

/**
 * 弹出对话框让用户输入比例，然后按比例分割选中的形状
 * 支持两种输入方式:
 * 1. 连续数字，如"2457"表示按照2:4:5:7的比例分割
 * 2. 带分隔符的数字，如"2:4:5:7"或"2 4 5 7"等
 */
divide_shape_by_ratio() {
    ; 获取当前选中的形状
    shape_range := get_ppt_object().shape_range
    if (!shape_range || shape_range.Count == 0) {
        Msgbox("请先选择一个形状")
        return
    }

    ; 获取第一个选中的形状
    shape := shape_range.Item(1)

    ; 弹出对话框让用户输入比例
    user_input := InputBox("请输入分割比例:" . "`n"
        . "1. 单个数字如'2457'将按2:4:5:7比例分割" . "`n"
        . "2. 也可使用分隔符，如'2:4:5:7'或'2 4 5 7'",
        "按比例分割形状", "w450 h160")

    ; 检查用户是否取消了操作
    if (user_input.Result = "Cancel") {
        return
    }

    ; 获取用户输入的文本
    ratio_text := user_input.Value
    ratio_text := Trim(ratio_text)

    ; 检查输入是否为空
    if (ratio_text = "") {
        Msgbox("请输入分割比例")
        return
    }

    ; 解析输入文本为比例数组
    ratios := []

    ; 检查是否包含分隔符
    if (RegExMatch(ratio_text, "[\s:,;]")) {
        ; 有分隔符，进行常规分割
        ratio_text := RegExReplace(ratio_text, "[^\d\s:,;]+", "")  ; 移除非数字和分隔符
        ratio_text := RegExReplace(ratio_text, "[\s:,;]+", ":")    ; 标准化分隔符
        ratio_text := RegExReplace(ratio_text, "^:|:$", "")        ; 移除首尾分隔符

        ; 分割成数组
        ratio_parts := StrSplit(ratio_text, ":")

        ; 将字符串转换为数字
        for part in ratio_parts {
            num := Integer(part)
            if (num <= 0) {
                continue  ; 跳过无效数字
            }
            ratios.Push(num)
        }
    } else {
        ; 无分隔符，将每个数字字符视为一个比例
        for i, char in StrSplit(ratio_text) {
            num := Integer(char)
            if (num <= 0) {
                continue  ; 跳过非数字字符
            }
            ratios.Push(num)
        }
    }

    ; 检查是否有有效的比例
    if (ratios.Length < 2) {
        Msgbox("需要至少两个有效的比例数字")
        return
    }

    ; 调用函数按比例分割形状
    new_shapes := divide_shape_into_columns_by_ratio(shape, ratios, 8)

    ; 显示分割信息
    if (new_shapes) {
        ratio_str := ""
        for ratio in ratios {
            ratio_str .= ratio . (A_Index < ratios.Length ? ":" : "")
        }

        Notify.show("已按比例 " . ratio_str . " 成功分割形状为 " . ratios.Length . " 列")
    }
}

/**
 * 按比例将PowerPoint形状分割成多个列
 * @param shape 要分割的PowerPoint形状
 * @param ratios 包含各列宽度比例的数组
 * @param gap_pts 形状之间的间隙(磅)
 * @param keep_original 是否保留原始形状
 * @return 返回包含所有新创建形状的数组
 */
divide_shape_into_columns_by_ratio(shape, ratios, gap_pts := 8, keep_original := false) {
    try {
        ; 获取形状的位置和尺寸
        shape_left := shape.Left
        shape_top := shape.Top
        shape_width := shape.Width
        shape_height := shape.Height

        ; 计算比例总和
        total_ratio := 0
        for ratio in ratios {
            total_ratio += ratio
        }

        ; 计算总间隙宽度
        total_gap_width := gap_pts * (ratios.Length - 1)

        ; 计算可用宽度(总宽度减去间隙)
        available_width := shape_width - total_gap_width

        ; 创建新形状数组
        new_shapes := []
        current_left := shape_left

        ; 创建新的形状
        for i, ratio in ratios {

            ; 计算当前列的宽度
            column_width := (ratio / total_ratio) * available_width

            ; 复制原始形状
            new_shape := shape.Duplicate()

            ; 设置新形状的位置和尺寸
            new_shape.Left := current_left
            new_shape.Top := shape_top
            new_shape.Width := column_width
            new_shape.Height := shape_height

            ; 添加到数组
            new_shapes.Push(new_shape)

            ; 更新下一个形状的左侧位置
            current_left += column_width + gap_pts
        }

        ; 删除原始形状(如果不保留)
        if (!keep_original) {
            shape.Delete()
        }

        ; 选择所有新创建的形状
        for i, new_shape in new_shapes {
            new_shape.Select(i > 1)  ; 第一个形状用Select，其他用AddToSelection
        }

        ; 返回新形状数组
        return new_shapes
    } catch as errrr {
        Msgbox("按比例分割形状时出错: " . err.Message)
        return 0
    }
}

create_array() {
    my_array := [1, 2, 3, 4, 5]
    return my_array
}

; #endregion

; #region 对齐形状

; { 对齐形状 ---------------------------------------------------------------------------------

; ; 将所选形状垂直堆叠排列
; stack_vert() {
;     ; 获取当前选中的形状范围
;     shape_range := get_ppt_object().shape_range

;     ; 检查是否至少选中了两个形状
;     if (shape_range.Count < 2) {
;         Msgbox("没有选中形状或者形状数量少于2个, 无法完成垂直堆叠操作!")
;         return
;     }

;     ; 按照顶部位置对形状进行排序
;     shapes := sort_shapes(range_to_array(shape_range) , "Top")

;     ; 初始化第一个形状的底部位置
;     prev_bottom := shapes[1].Top + shapes[1].Height

;     ; 遍历剩余的形状并依次堆叠
;     Loop (shapes.Length - 1) {
;         shape := shapes[A_Index + 1]
;         shape.Top := prev_bottom  ; 将当前形状的顶部与上一个形状的底部对齐
;         prev_bottom := shape.Top + shape.Height  ; 更新底部位置
;     }
; }

; ; 将所选形状水平堆叠排列
; stack_horiz() {
;     ; 获取当前选中的形状范围
;     shape_range := get_ppt_object().shape_range

;     ; 检查是否至少选中了两个形状
;     if (shape_range.Count < 2) {
;         Msgbox("没有选中形状或者形状数量少于2个, 无法完成水平堆叠操作!")
;         return
;     }

;     ; 按照左侧位置对形状进行排序
;     shapes := sort_shapes(range_to_array(shape_range), "Left")

;     ; 初始化第一个形状的右侧位置
;     prev_right := shapes[1].Left + shapes[1].Width

;     ; 遍历剩余的形状并依次堆叠
;     Loop (shapes.Length - 1) {
;         shape := shapes[A_Index + 1]
;         shape.Left := prev_right  ; 将当前形状的左侧与上一个形状的右侧对齐
;         prev_right := shape.Left + shape.Width  ; 更新右侧位置
;     }
; }

; adjust_stack_spacing(direction := "vertical", way := "increase") {
;     if way == "increase"
;         spacing_increment := 2 ; 间隔递增单位
;     else
;         spacing_increment := -2 ; 间隔递增单位
;     ppt_app := ComObjActive("PowerPoint.Application")
;     selection := ppt_app.ActiveWindow.Selection
;     shape_range := selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange

;     if (shape_range.Count < 2) {
;         Msgbox("请至少选择两个形状。")
;         return
;     }

;     ; 根据方向选择排序属性和提示信息
;     if (direction = "vertical") {
;         property := "Top"
;     } else if (direction = "horizontal") {
;         property := "Left"
;     } else {
;         Msgbox("无效的方向。请使用 `"vertical`" 或 `"horizontal`"。")
;         return
;     }

;     ; 排序形状
;     shapes := sort_shapes(shape_range, property)
;     if (!shapes) {
;         return
;     }

;     ; 间隔加2
;     Try {
;         if (direction = "vertical") {
;             Loop (shapes.Length - 1) {
;                 shape := shapes[A_Index+1]
;                 current_increment := (A_Index) * spacing_increment
;                 shape.Top := shape.Top + current_increment
;             }
;         } else {
;             Loop (shapes.Length - 1) {
;                 shape := shapes[A_Index+1]
;                 current_increment := (A_Index) * spacing_increment
;                 shape.Left := shape.Left + current_increment
;             }
;         }
;     }

;     ppt_app := ""
;     selection := ""
;     shape_range := ""
; }

; } 对齐形状 ---------------------------------------------------------------------------------

; #endregion

; 设置形状标签
set_shape_tag(shape, tag_name, tag_value) {
    ; 验证参数
    if (!IsObject(shape)) {
        Msgbox("形状对象无效")
        return 0
    }

    if (tag_name == "") {
        Msgbox("标签名称不能为空")
        return 0
    }

    ; 尝试删除旧标签，忽略任何错误
    try {
        shape.Tags.Delete tag_name
    }

    ; 添加新标签
    try {
        shape.Tags.Add tag_name, tag_value
        return 1  ; 成功
    } catch as err {
        Msgbox("添加标签失败: " . err.Message)
        return 0  ; 失败
    }
}

; #region dialog

master_view(ppt_app := "") {
    try {
        if !ppt_app
            ppt_app := ComObjActive("PowerPoint.Application")

        cb := ppt_app.CommandBars("Slide Master View")
        if (cb.Visible) {
            ppt_app.CommandBars.ExecuteMso("ViewThumbnailViewPowerPoint")   ; 普通视图
            Notify.show_green("普通视图")
        } else {
            ppt_app.CommandBars.ExecuteMso("ViewSlideMasterView")   ; 母版视图
            Notify.show_blue("母版视图")
        }

    } catch
        throw
}

textleft_1row_slide_guides() {
    try {
        obj := get_ppt_object()
        app := obj.app

        if (obj.type ~= "i)^shapes|text") {
            Notify.show("文字左对齐")
            app.CommandBars.ExecuteMso("AlignLeft")
            return
        }
        add_1row_slide_guides()
    } catch as err {
        MsgBox(err.Message)
    }
}

textleft_2row_slide_guides() {
    try {
        obj := get_ppt_object()
        app := obj.app

        if (obj.type ~= "i)^shapes|text") {
            Notify.show("文字左对齐")
            app.CommandBars.ExecuteMso("AlignLeft")
            return
        }
        add_2row_slide_guides()
    } catch as err {
        MsgBox(err.Message)
    }
}

textleft_3row_slide_guides() {
    try {
        obj := get_ppt_object()
        app := obj.app

        if (obj.type ~= "i)^shapes|text") {
            Notify.show("文字左对齐")
            app.CommandBars.ExecuteMso("AlignLeft")
            return
        }
        add_3row_slide_guides()
    } catch as err {
        MsgBox(err.Message)
    }
}

textleft_4row_slide_guides() {
    try {
        obj := get_ppt_object()
        app := obj.app

        if (obj.type ~= "i)^shapes|text") {
            Notify.show("文字左对齐")
            app.CommandBars.ExecuteMso("AlignLeft")
            return
        }
        add_4row_slide_guides()
    } catch as err {
        MsgBox(err.Message)
    }
}

selection_pane(ppt_app := "") {
    try {
        if !ppt_app
            ppt_app := ComObjActive("PowerPoint.Application")

        cb := ppt_app.CommandBars("Selection")
        if (!cb.Enabled) {
            ppt_app.CommandBars.ExecuteMso("SelectionPane")
            Notify.show_green("打开选择窗格")
            return
        }

        if (cb.Visible)
            Notify.show_red("关闭选择窗格")
        else
            Notify.show_green("打开选择窗格")

        cb.Visible := !cb.Visible

    } catch
        throw
}

; 格式对话框 1
format_dialog(ppt_app := "") {
    if (ppt_app = "") {
        try {
            ppt_app := ComObjActive("PowerPoint.Application")
        } catch {
            MsgBox("错误：PPT 未运行，请先打开 PowerPoint。")
            return
        }
    }

    try {
        selection := ppt_app.ActiveWindow.Selection
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange

        shape_type := SHAPE_TYPE_MAP[shape_range(1).Type]

        switch shape_type {
            case "msoTextBox":
                ppt_app.CommandBars.ExecuteMso("TextFillMoreGradientsDialog")
                Notify.show("设置文本填充")
            case "msoPicture":
                ppt_app.CommandBars.ExecuteMso("PictureCorrectionsDialog")
                Notify.show("图片校正")
            default:
                ppt_app.CommandBars.ExecuteMso("ObjectFormatDialog")
                Notify.show("设置文本填充")
        }
    }
    catch as err
        Msgbox("打开格式对话框时出错: " . err.Message)
}

; 格式对话框 2
shadow_dialog(ppt_app := "") {
    if (ppt_app = "") {
        try {
            ppt_app := ComObjActive("PowerPoint.Application")
        } catch {
            MsgBox("错误：PPT 未运行，请先打开 PowerPoint。")
            return
        }
    }

    try {
        selection := ppt_app.ActiveWindow.Selection
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange

        shape_type := SHAPE_TYPE_MAP[shape_range(1).Type]

        switch shape_type {
            case "msoTextBox":
                ppt_app.CommandBars.ExecuteMso("TextEffectsMoreShadowsDialog")
                Notify.show("设置文本阴影")
            case "msoPicture":
                ppt_app.CommandBars.ExecuteMso("PictureColorDialog")
                Notify.show("图片颜色")
            default:
                ppt_app.CommandBars.ExecuteMso("ShadowOptionsDialog")
                Notify.show("设置形状阴影")
        }
    }
    catch as err
        Msgbox("打开格式阴影对话框时出错: " . err.Message)
}

; 格式对话框 3
size_dialog(ppt_app := "") {
    if (ppt_app = "") {
        try {
            ppt_app := ComObjActive("PowerPoint.Application")
        } catch {
            MsgBox("错误：PPT 未运行，请先打开 PowerPoint。")
            return
        }
    }

    try {
        selection := ppt_app.ActiveWindow.Selection
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange

        shape_type := SHAPE_TYPE_MAP[shape_range(1).Type]

        switch shape_type {
            case "msoTextBox":
                ppt_app.CommandBars.ExecuteMso("TextDirectionMoreOptionsDialog")
                Notify.show("设置文本尺寸")
            case "msoPicture":
                ppt_app.CommandBars.ExecuteMso("PictureMoreTransparency")
                Notify.show("图片透明度")
            default:
                ppt_app.CommandBars.ExecuteMso("ObjectSizeAndPropertiesDialog")
                Notify.show("设置形状大小和位置")
        }
    }
    catch as err
        Msgbox("打开格式尺寸框时出错: " . err.Message)
}

; #endregion

; #region  insert_text

insert_paragraph2(index := 1) {
    result := select_region_interactive("Purple")
    if (!result.region) {
        return { success: false, file_path: "", message: "用户取消了区域选择" }
    }
    parts := StrSplit(result.region, "|")
    x := parts[1]
    y := parts[2]
    w := parts[3]
    h := parts[4]

    ppt_app := ComObjActive("PowerPoint.Application")
    window := ppt_app.ActiveWindow

    origin_x := window.PointsToScreenPixelsX(0)
    origin_y := window.PointsToScreenPixelsY(0)
    zoom := window.View.Zoom / 100
    x := pixel_to_point((x - origin_x) / zoom, 88)
    y := pixel_to_point((y - origin_y) / zoom, 90)
    w := pixel_to_point(w / zoom, 88)
    h := pixel_to_point(h / zoom, 90)

    shape := window.View.Slide.Shapes.AddTextbox(1, x, y, w, h)
    font_size := 12
    shape.TextFrame.TextRange.Font.Size := font_size ; 字体大小
    shape.TextFrame.AutoSize := 0
    shape.select

    sample_shape_name := "para_" . index
    slides := ppt_app.ActivePresentation.Slides
    shape_exists := false
    for sample_shape in Slides(1).Shapes {
        if (sample_shape.Name = sample_shape_name) {
            shape_exists := true
            break
        }
    }

    if (shape_exists) {
        sample_shape := Slides(1).Shapes("para_" . index)
        if (sample_shape.HasTextFrame == -1 && sample_shape.TextFrame.HasText) {
            sample_char_count := StrLen(sample_shape.TextFrame.TextRange.Text)
            sample_width := sample_shape.Width
            sample_height := sample_shape.Height
            char_count := Round(sample_char_count * w * h / sample_width / sample_height)
            ; Notify.show char_count " | " sample_char_count " | " w " | " sample_width
        }
        shape.TextFrame.TextRange.Text := generate_random_paragraph(char_count)
        sample_shape.Pickup
        shape.Apply
        shape.TextFrame.AutoSize := 1
        shape.name := sample_shape_name . "_" . A_Index
        return
    }

    char_count := Ceil(w * 1 / font_size)
    shape.TextFrame.TextRange.Text := generate_random_text(char_count)
    shape.TextFrame.MarginLeft := 0
    shape.TextFrame.MarginRight := 0
    shape.TextFrame.MarginTop := 0
    shape.TextFrame.MarginBottom := 0
    shape.TextFrame.TextRange.Font.Name := "微软雅黑" ; 字体名称
    shape.TextFrame.WordWrap := -1
    shape.TextFrame.AutoSize := 0
    shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin := true
    shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin := 1
}

insert_text_paragraph_interactive3(index := 1) {
    ; global super_power := False
    ; 交互式选择区域
    result := select_region_interactive("Purple")
    if (!result.region) {
        return { success: false, file_path: "", message: "用户取消了区域选择" }
    }
}

; insert_text_line_interactive
insert_text_paragraph_interactive4(index := 1) {
    ; global super_power := False
    ; 交互式选择区域
    result := select_region_interactive("Purple")
    if (!result.region) {
        return { success: false, file_path: "", message: "用户取消了区域选择" }
    }

    ; 解析区域坐标
    parts := StrSplit(result.region, "|")
    if (parts.Length < 4) {
        return { success: false, file_path: "", message: "区域坐标格式错误" }
    }

    x := parts[1], y := parts[2], w := parts[3], h := parts[4]

    ; 错误处理：检查PowerPoint是否可用
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
    } catch {
        return { success: false, file_path: "", message: "无法连接到PowerPoint应用程序" }
    }

    ; 检查是否有活动窗口
    try {
        window := ppt_application.ActiveWindow
        if (!window) {
            throw "没有活动窗口"
        }
    } catch as err {
        return { success: false, file_path: "", message: "没有活动的PowerPoint窗口" }
    }

    ; 坐标转换
    try {
        origin_x := window.PointsToScreenPixelsX(0)
        origin_y := window.PointsToScreenPixelsY(0)
        zoom := window.View.Zoom / 100

        ; 转换为PowerPoint点位
        ppt_x := pixel_to_point((x - origin_x) / zoom, 88)  ; 使用标准DPI
        ppt_y := pixel_to_point((y - origin_y) / zoom, 90)
        ppt_w := pixel_to_point(w / zoom, 88)
        ppt_h := pixel_to_point(h / zoom, 90)

    } catch as err {
        return { success: false, file_path: "", message: "坐标转换失败: " . err.message }
    }

    ; 创建文本框
    try {
        shape := window.View.Slide.Shapes.AddTextbox(1, ppt_x, ppt_y, ppt_w, ppt_h)
        shape.TextFrame.AutoSize := 0
        font_size := 12
        shape.TextFrame.TextRange.Font.Size := font_size
        shape.Select
    } catch as err {
        return { success: false, file_path: "", message: "创建文本框失败: " . err.message }
    }

    ; 查找样本形状
    sample_shape_name := "para_" . index
    slides := ppt_application.ActivePresentation.Slides
    shape_exists := false
    sample_shape := ""
    char_count := 0  ; 初始化字符数

    try {
        ; 修复：正确访问幻灯片集合
        first_slide := slides.Item(1)
        shapes_collection := first_slide.Shapes

        ; 遍历第一张幻灯片的所有形状
        loop shapes_collection.Count {
            current_shape := shapes_collection.Item(A_Index)
            if (current_shape.Name = sample_shape_name) {
                sample_shape := current_shape
                shape_exists := true
                break
            }
        }
    } catch as err {
        ; 如果找不到样本形状，继续使用默认设置
        shape_exists := false
    }

    ; 如果找到样本形状，复制其格式和计算字符数
    if (shape_exists && sample_shape) {
        try {
            if (sample_shape.HasTextFrame = -1 && sample_shape.TextFrame.HasText = -1) {
                sample_text := sample_shape.TextFrame.TextRange.Text
                sample_char_count := StrLen(sample_text)
                sample_width := sample_shape.Width
                sample_height := sample_shape.Height

                ; 根据面积比例计算字符数
                if (sample_width > 0 && sample_height > 0) {
                    char_count := Round(sample_char_count * ppt_w * ppt_h / (sample_width * sample_height))
                    char_count := Max(char_count, 10)  ; 最少10个字符
                }

                ; 生成文本并应用格式
                shape.TextFrame.TextRange.Text := generate_random_paragraph(char_count)
                sample_shape.PickUp()
                shape.Apply()
                shape.TextFrame.AutoSize := 1
                shape.Name := sample_shape_name . "_" . A_TickCount  ; 使用时间戳避免重名

                return { success: true, file_path: "", message: "成功插入段落（使用样本格式）" }
            }
        } catch as err {
            ; 如果应用样本格式失败，继续使用默认格式
        }
    }

    ; char_count := Ceil(w * 1 / font_size)
    ; shape.TextFrame.TextRange.Text := generate_random_text(char_count)
    ; shape.TextFrame.MarginLeft := 0
    ; shape.TextFrame.MarginRight := 0
    ; shape.TextFrame.MarginTop := 0
    ; shape.TextFrame.MarginBottom := 0
    ; shape.TextFrame.TextRange.Font.Name := "微软雅黑" ; 字体名称
    ; shape.TextFrame.WordWrap := -1
    ; shape.TextFrame.AutoSize := 0
    ; shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin := true
    ; shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin := 1

    ; 默认格式设置
    try {
        ; 如果没有计算出字符数，使用默认算法
        ; if (char_count = 0) {
        ;     char_count := Ceil(ppt_w * ppt_h / (font_size * font_size) * 0.8)  ; 改进的字符数计算
        ;     char_count := Max(char_count, 20)  ; 最少20个字符
        ;     char_count := Min(char_count, 500) ; 最多500个字符
        ; }

        ; char_count := Ceil(ppt_w * 1 / font_size)
        char_count := Ceil(ppt_w * ppt_h / (font_size * font_size) * 0.5)  ; 改进的字符数计算

        ; MsgBox(char_count, , "T2")

        shape.TextFrame.TextRange.Text := generate_random_paragraph(char_count)

        ; 设置边距和格式
        shape.TextFrame.MarginLeft := 0  ; 2毫米
        shape.TextFrame.MarginRight := 0
        shape.TextFrame.MarginTop := 0
        shape.TextFrame.MarginBottom := 0

        ; 字体设置
        shape.TextFrame.TextRange.Font.Name := "微软雅黑"
        shape.TextFrame.TextRange.Font.Size := font_size

        ; 段落格式
        shape.TextFrame.WordWrap := True
        shape.TextFrame.AutoSize := 1
        shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin := True
        shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin := 1.3
        shape.TextFrame.TextRange.ParagraphFormat.Alignment := 4  ; 两端对齐

        ; 设置形状名称
        shape.Name := "para_" . index . "_" . A_TickCount

        WinActivate("ahk_exe POWERPNT.EXE")

        return { success: true, file_path: "", message: "成功插入段落（使用默认格式）" }

    } catch as err {
        return { success: false, file_path: "", message: "设置文本格式失败: " . err.message }
    }
}

; ============================================================================
; 第一部分：屏幕取点 - 获取用户选择的区域并转换为PPT坐标
; ============================================================================
get_ppt_region_from_screen(color := "Purple", button := "RButton") {
    ; 交互式选择区域
    result := select_region_interactive(color, button)
    if (!result.region) {
        return { success: false, message: "用户取消了区域选择" }
    }

    ; 解析区域坐标
    parts := StrSplit(result.region, "|")
    if (parts.Length < 4) {
        return { success: false, message: "区域坐标格式错误" }
    }

    screen_x := parts[1]
    screen_y := parts[2]
    screen_w := parts[3]
    screen_h := parts[4]

    ; 错误处理：检查PowerPoint是否可用
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
    } catch {
        return { success: false, message: "无法连接到PowerPoint应用程序" }
    }

    ; 检查是否有活动窗口
    try {
        window := ppt_application.ActiveWindow
        if (!window) {
            throw "没有活动窗口"
        }
    } catch as err {
        return { success: false, message: "没有活动的PowerPoint窗口" }
    }

    ; 坐标转换：屏幕像素 -> PPT点位
    try {
        origin_x := window.PointsToScreenPixelsX(0)
        origin_y := window.PointsToScreenPixelsY(0)
        zoom := window.View.Zoom / 100

        ; 转换为PowerPoint点位
        ppt_x := pixel_to_point((screen_x - origin_x) / zoom, 88)
        ppt_y := pixel_to_point((screen_y - origin_y) / zoom, 90)
        ppt_w := pixel_to_point(screen_w / zoom, 88)
        ppt_h := pixel_to_point(screen_h / zoom, 90)

        return {
            success: true,
            ppt_x: ppt_x,
            ppt_y: ppt_y,
            ppt_w: ppt_w,
            ppt_h: ppt_h,
            window: window,
            ppt_app: ppt_application
        }
    } catch as err {
        return { success: false, message: "坐标转换失败: " . err.message }
    }
}

; ============================================================================
; 第二部分：调取样本 - 查找并获取样本形状的格式信息（通用版本）
; ============================================================================
get_sample_format(ppt_application, sample_slide_id := 1, sample_name := "sample_shape_1") {
    try {
        slides := ppt_application.ActivePresentation.Slides

        ; 检查幻灯片ID是否有效
        if (sample_slide_id < 1 || sample_slide_id > slides.Count) {
            return { success: true, has_template: false, message: "幻灯片ID超出范围" }
        }

        target_slide := slides.Item(sample_slide_id)
        shapes_collection := target_slide.Shapes

        ; 遍历指定幻灯片的所有形状
        loop shapes_collection.Count {
            current_shape := shapes_collection.Item(A_Index)
            if (current_shape.Name = sample_name) {
                ; 检查是否有文本框
                if (current_shape.HasTextFrame = -1) {
                    sample_text := current_shape.TextFrame.TextRange.Text
                    sample_char_count := StrLen(sample_text)
                    sample_width := current_shape.Width
                    sample_height := current_shape.Height

                    return {
                        success: true,
                        shape: current_shape,
                        char_count: sample_char_count,
                        width: sample_width,
                        height: sample_height,
                        has_sample: true
                    }
                }
            }
        }

        ; 未找到样本
        return { success: true, has_sample: false }

    } catch as err {
        ; 查找样本失败，返回默认设置
        return { success: true, has_sample: false, message: "查找样本失败: " . err.message }
    }
}

; ============================================================================
; 第三部分：绘制图形 - 创建文本框并应用格式
; ============================================================================
draw_text_shape(region_data) {
    if (!region_data.success) {
        return region_data
    }

    window := region_data.window
    ppt_x := region_data.ppt_x
    ppt_y := region_data.ppt_y
    ppt_w := region_data.ppt_w
    ppt_h := region_data.ppt_h

    ; 创建文本框
    try {
        shape := window.View.Slide.Shapes.AddTextbox(1, ppt_x, ppt_y, ppt_w, ppt_h)
        shape.TextFrame.AutoSize := 0
        font_size := 12
        shape.TextFrame.TextRange.Font.Size := font_size
        shape.Select
    } catch as err {
        return { success: false, message: "创建文本框失败: " . err.message }
    }

    ; 使用默认格式
    try {
        ; 计算字符数
        char_count := Ceil(ppt_w * ppt_h / (font_size * font_size) * 0.5)
        char_count := Max(char_count, 20)   ; 最少20个字符
        char_count := Min(char_count, 500)  ; 最多500个字符

        ; 设置文本内容
        shape.TextFrame.TextRange.Text := generate_random_paragraph(char_count)

        ; 设置边距
        shape.TextFrame.MarginLeft := 0
        shape.TextFrame.MarginRight := 0
        shape.TextFrame.MarginTop := 0
        shape.TextFrame.MarginBottom := 0

        ; 字体设置
        shape.TextFrame.TextRange.Font.Name := "微软雅黑"
        shape.TextFrame.TextRange.Font.Size := font_size

        ; 段落格式
        shape.TextFrame.WordWrap := True
        shape.TextFrame.AutoSize := 1
        shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin := True
        shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin := 1.3
        shape.TextFrame.TextRange.ParagraphFormat.Alignment := 4  ; 两端对齐

        ; 设置形状名称
        ; shape.Name := sample_name . "_" . A_TickCount

        WinActivate("ahk_exe POWERPNT.EXE")
        return { success: true, message: "成功插入段落（使用默认格式）" }

    } catch as err {
        return { success: false, message: "设置文本格式失败: " . err.message }
    }
}

; ============================================================================
; 主函数：整合三个部分
; ============================================================================
; insert_paragraph(color := "Purple", button := "RButton") {
;     ; 第一步：屏幕取点
;     region_data := get_ppt_region_from_screen(color, button)
;     if (!region_data.success) {
;         return {success: false, file_path: "", message: region_data.message}
;     }

;     ; 第二步：绘制图形
;     result := draw_text_shape(region_data)

;     return {success: result.success, file_path: "", message: result.message}
; }

; insert_paragraph_1() {
;    insert_paragraph(1, "sample_paragraph_1")
; }

; ============================================================================
; insert_shape - 插入形状（矩形、圆形等）
; ============================================================================

; 第三部分：绘制形状
draw_shape(region_data, shape_type := 1) {
    ; shape_type: 1=矩形, 2=圆角矩形, 3=椭圆, 5=菱形, 9=圆形等
    if (!region_data.success) {
        return region_data
    }

    window := region_data.window
    ppt_x := region_data.ppt_x
    ppt_y := region_data.ppt_y
    ppt_w := region_data.ppt_w
    ppt_h := region_data.ppt_h

    ; 创建形状
    try {
        shape := window.View.Slide.Shapes.AddShape(shape_type, ppt_x, ppt_y, ppt_w, ppt_h)
        shape.Select
    } catch as err {
        return { success: false, message: "创建形状失败: " . err.message }
    }

    ; 使用默认格式
    try {
        ; 设置默认填充和边框
        shape.Fill.ForeColor.RGB := 0xDCDCDC  ; 浅灰色填充 (RGB: 220, 220, 220)
        shape.Line.Weight := 1.5
        shape.Line.ForeColor.RGB := 0x646464  ; 深灰色边框 (RGB: 100, 100, 100)

        WinActivate("ahk_exe POWERPNT.EXE")
        return { success: true, message: "成功插入形状（使用默认格式）" }

    } catch as err {
        return { success: false, message: "设置形状格式失败: " . err.message }
    }
}

; ; 主函数：插入形状
; insert_shape2(sample_slide_id := 1, sample_name := "sample_shape_1", shape_type := 1, color := "Purple", button := "RButton") {
;     ; 第一步：屏幕取点
;     region_data := get_ppt_region_from_screen(color, button)
;     if (!region_data.success) {
;         return {success: false, file_path: "", message: region_data.message}
;     }

;     ; 第二步：调取样本
;     sample_data := get_sample_format(region_data.ppt_app, sample_slide_id, sample_name)

;     ; 第三步：绘制形状
;     result := draw_shape(region_data, sample_data, sample_name, shape_type)

;     return {success: result.success, file_path: "", message: result.message}
; }

insert_rounded_rectangle() {
    insert_shape(1)
}

; ============================================================================
; insert_shape - 插入形状（矩形、圆形等）
; ============================================================================
insert_paragraph22(font_size := 12) {
    ; 第一步：屏幕取点
    region_data := get_ppt_region()
    if (!region_data.success) {
        return { success: false, file_path: "", message: region_data.message }
    }

    window := region_data.window
    ppt_x := region_data.ppt_x
    ppt_y := region_data.ppt_y
    ppt_w := region_data.ppt_w
    ppt_h := region_data.ppt_h

    ; 创建文本框
    try {
        shape := window.View.Slide.Shapes.AddTextbox(1, ppt_x, ppt_y, ppt_w, ppt_h)
        shape.TextFrame.AutoSize := 0  ; 先关闭自动大小
        shape.Select
    } catch as err {
        return { success: false, message: "创建文本框失败: " . err.message }
    }

    try {

        char_count := Ceil(ppt_w * ppt_h / (font_size * font_size) * 0.5)

        ; 设置字符数限制
        char_count := Max(char_count, 5)     ; 最少5个字符
        char_count := Min(char_count, 500)   ; 最多500个字符

        ; 设置文本内容
        if (IsSet(generate_random_paragraph)) {
            shape.TextFrame.TextRange.Text := generate_random_paragraph(char_count)
        } else {
            ; 如果没有随机文本生成函数，使用默认文本
            default_text := "这是一个文本框。您可以在PowerPoint中直接编辑此文本内容。"
            shape.TextFrame.TextRange.Text := default_text
        }

        ; 字体设置
        shape.TextFrame.TextRange.Font.Name := "微软雅黑"
        shape.TextFrame.TextRange.Font.Size := font_size
        shape.TextFrame.TextRange.Font.Color.RGB := 0x4D4D4D  ; 黑色文字

        ; 设置边距
        shape.TextFrame.MarginLeft := 0
        shape.TextFrame.MarginRight := 0
        shape.TextFrame.MarginTop := 0
        shape.TextFrame.MarginBottom := 0

        ; 段落格式
        shape.TextFrame.WordWrap := True
        shape.TextFrame.AutoSize := 1
        shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin := True
        shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin := 1.3
        shape.TextFrame.TextRange.ParagraphFormat.Alignment := 4  ; 两端对齐

        ; 设置形状名称
        shape.Name := "paragraph_" . A_TickCount

        WinActivate("ahk_exe POWERPNT.EXE")
        return { success: true, message: "成功插入文本框（字体大小：" font_size "）" }

    } catch as err {
        return { success: false, message: "设置文本格式失败: " . err.message }
    }
}

insert_paragraph(font_size := 12, text_color := 0x4D4D4D) {
    ; 参数验证
    if (font_size < 6 || font_size > 72) {
        return { success: false, message: "字体大小应在6-72之间" }
    }

    ; 第一步：屏幕取点
    region_data := get_ppt_region()
    if (!region_data.success) {
        return { success: false, file_path: "", message: region_data.message }
    }

    window := region_data.window
    ppt_x := region_data.ppt_x
    ppt_y := region_data.ppt_y
    ppt_w := region_data.ppt_w
    ppt_h := region_data.ppt_h

    ; 创建文本框
    try {
        shape := window.View.Slide.Shapes.AddTextbox(1, ppt_x, ppt_y, ppt_w, ppt_h)
        shape.TextFrame.AutoSize := 0  ; 先关闭自动大小
        shape.Select
    } catch as err {
        return { success: false, message: "创建文本框失败: " . err.message }
    }

    try {
        ; 更精确的字符数计算
        char_count := Ceil((ppt_w / font_size) * (ppt_h / font_size) * 0.5)
        char_count := Clamp(char_count, 10, 500)

        ; 设置文本内容
        if (IsSet(generate_random_paragraph)) {
            text_content := generate_random_paragraph(char_count)
        } else {
            text_content := "这是一个文本框。您可以在PowerPoint中直接编辑此文本内容。"
        }
        shape.TextFrame.TextRange.Text := text_content

        ; 字体设置（带回退）
        try {
            shape.TextFrame.TextRange.Font.Name := "微软雅黑"
        } catch {
            shape.TextFrame.TextRange.Font.Name := "Arial"
        }

        shape.TextFrame.TextRange.Font.Size := font_size
        shape.TextFrame.TextRange.Font.Color.RGB := text_color

        ; 布局设置
        shape.TextFrame.MarginLeft := 0
        shape.TextFrame.MarginRight := 0
        shape.TextFrame.MarginTop := 0
        shape.TextFrame.MarginBottom := 0

        ; 段落格式
        shape.TextFrame.WordWrap := True
        shape.TextFrame.AutoSize := 1  ; 重新启用自动大小
        shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin := True
        shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin := 1.3
        shape.TextFrame.TextRange.ParagraphFormat.Alignment := 4  ; 两端对齐

        ; 设置形状名称
        shape.Name := "paragraph_" . A_TickCount

        WinActivate("ahk_exe POWERPNT.EXE")
        return { success: true, message: "成功插入文本框（字体大小：" font_size "）" }

    } catch as err {
        return { success: false, message: "设置文本格式失败: " . err.message }
    }
}
insert_title_20() {
    insert_title()
}
insert_title(font_size := 20, text_color := 0x000000) {
    ; 参数验证
    if (font_size < 12 || font_size > 96) {
        return { success: false, message: "标题字体大小应在12-96之间" }
    }

    ; 第一步：屏幕取点
    region_data := get_ppt_region()
    if (!region_data.success) {
        return { success: false, file_path: "", message: region_data.message }
    }

    window := region_data.window
    ppt_x := region_data.ppt_x
    ppt_y := region_data.ppt_y
    ppt_w := region_data.ppt_w
    ppt_h := region_data.ppt_h

    ; 创建标题文本框（高度根据字体大小自适应）
    title_height := font_size * 1.5  ; 标题行高约为字体的1.5倍
    try {
        shape := window.View.Slide.Shapes.AddTextbox(1, ppt_x, ppt_y, ppt_w, title_height)
        shape.TextFrame.AutoSize := 0  ; 先关闭自动大小
        shape.Select
    } catch as err {
        return { success: false, message: "创建标题文本框失败: " . err.message }
    }

    try {
        ; 根据宽度和字体大小计算标题字数
        ; 假设中文字符宽度约为字体大小的0.8倍
        char_width := font_size * 1
        max_chars := Floor(ppt_w / char_width)

        ; 设置合理的字数范围
        char_count := Clamp(max_chars, 2, 20)  ; 标题通常2-20个字

        ; 设置标题文本内容
        if (IsSet(generate_random_paragraph)) {
            title_text := generate_random_paragraph(char_count)
        } else {
            ; 默认标题文本
            default_titles := ["项目标题", "主要内容", "核心观点", "重点分析", "总结报告"]
            random_index := Random(1, default_titles.Length)
            title_text := default_titles[random_index]
        }
        shape.TextFrame.TextRange.Text := title_text

        ; 字体设置（标题常用字体）
        try {
            shape.TextFrame.TextRange.Font.Name := "微软雅黑"
        } catch {
            try {
                shape.TextFrame.TextRange.Font.Name := "黑体"
            } catch {
                shape.TextFrame.TextRange.Font.Name := "Arial"
            }
        }

        shape.TextFrame.TextRange.Font.Size := font_size
        shape.TextFrame.TextRange.Font.Color.RGB := text_color
        shape.TextFrame.TextRange.Font.Bold := True  ; 标题通常加粗

        ; 布局设置
        shape.TextFrame.MarginLeft := 0
        shape.TextFrame.MarginRight := 0
        shape.TextFrame.MarginTop := 0
        shape.TextFrame.MarginBottom := 0

        ; 标题特定格式
        shape.TextFrame.WordWrap := False  ; 标题不自动换行
        shape.TextFrame.AutoSize := 1      ; 启用自动大小以适应文本长度

        ; 段落对齐（标题通常居中对齐）
        shape.TextFrame.TextRange.ParagraphFormat.Alignment := 2  ; 居中对齐

        ; 设置形状名称
        shape.Name := "title_" . A_TickCount

        WinActivate("ahk_exe POWERPNT.EXE")
        return { success: true, message: "成功插入标题（字体大小：" font_size "）", text: title_text }

    } catch as err {
        return { success: false, message: "设置标题格式失败: " . err.message }
    }
}

; 主函数：插入形状
insert_shape(shape_type := 1) {
    ; 第一步：屏幕取点
    region_data := get_ppt_region()
    if (!region_data.success) {
        return { success: false, file_path: "", message: region_data.message }
    }

    ; 第二步：绘制形状
    ; shape_type: 1=矩形, 2=圆角矩形, 3=椭圆, 5=菱形, 9=圆形等
    window := region_data.window
    ppt_x := region_data.ppt_x
    ppt_y := region_data.ppt_y
    ppt_w := region_data.ppt_w
    ppt_h := region_data.ppt_h

    ; 创建形状
    try {
        shape := window.View.Slide.Shapes.AddShape(shape_type, ppt_x, ppt_y, ppt_w, ppt_h)
        shape.Select
    } catch as err {
        return { success: false, message: "创建形状失败: " . err.message }
    }

    ; 使用默认格式
    try {
        ; 设置默认填充和边框
        shape.Fill.ForeColor.RGB := 0xDCDCDC  ; 浅灰色填充 (RGB: 220, 220, 220)
        shape.Line.Weight := 1.5
        shape.Line.ForeColor.RGB := 0x646464  ; 深灰色边框 (RGB: 100, 100, 100)

        ; WinActivate("ahk_exe POWERPNT.EXE")
        return { success: true, message: "成功插入形状（使用默认格式）" }

    } catch as err {
        return { success: false, message: "设置形状格式失败: " . err.message }
    }
}

/**
 * 从屏幕选择区域并转换为PowerPoint坐标
 * @param {String} color 选择框颜色（默认："Purple"）
 * @param {String} button 触发按键（默认："RButton"）
 * @param {Number} transparent 透明度（默认：100）
 * @param {Number} min_region_size 最小区域尺寸（默认：5）
 * @return {Object} 包含成功状态、PPT坐标和原始区域信息
 */
get_ppt_region(button := "RButton", color := "Purple", transparent := 100, min_region_size := 5) {
    start_time := A_TickCount    ; 记录开始时间
    region_width := region_height := region_x := region_y := 0        ; 初始化区域变量
    selection_gui := 0                       ; GUI对象
    info_gui := 0                ; 信息显示GUI

    try {
        ; 获取起始鼠标位置
        MouseGetPos(&begin_x, &begin_y)
        ; 创建选择框GUI
        selection_gui := Gui()
        selection_gui.Opt("+AlwaysOnTop -Caption +Border +ToolWindow +LastFound -DPIScale")
        selection_gui.BackColor := color
        WinSetTransparent(transparent)
        selection_gui.Show()

        ; 创建信息显示GUI
        info_gui := Gui()
        info_gui.Opt("+AlwaysOnTop -Caption +ToolWindow +LastFound -DPIScale")
        info_gui.BackColor := 0x000000
        info_text := info_gui.Add("Text", "cWhite", "选择区域中...")
        WinSetTransparent(200)
        info_gui.Show("x" . (begin_x + 10) . " y" . (begin_y - 30) . " AutoSize")

        ; 实时更新选择框，直到按键释放
        while GetKeyState(button, "P") {
            ; 检查ESC键取消
            if GetKeyState("Escape", "P") {
                return { success: false, message: "用户取消了区域选择" }
            }

            Sleep(10)
            MouseGetPos(&end_x, &end_y)

            ; 计算选择区域
            region_width := Abs(begin_x - end_x)
            region_height := Abs(begin_y - end_y)
            region_x := (begin_x < end_x) ? begin_x : end_x
            region_y := (begin_y < end_y) ? begin_y : end_y

            ; 移动选择框
            selection_gui.Move(region_x, region_y, region_width, region_height)

            ; 更新信息显示
            info_text.Text := Format("区域: {}x{} | 位置: ({}, {})", region_width, region_height, region_x, region_y)
            info_gui.Move(end_x + 15, end_y - 40)
        }

        ; 计算总按键时长
        total_duration := A_TickCount - start_time

        ; 检查区域大小
        if (region_width < min_region_size || region_height < min_region_size) {
            if (total_duration < 200) {  ; 短按认为是取消
                return { success: false, message: "用户取消了区域选择" }
            } else {
                return { success: false, message: "选择的区域太小" }
            }
        }

        ; 错误处理：检查PowerPoint是否可用
        try {
            ppt_application := ComObjActive("PowerPoint.Application")
        } catch {
            return { success: false, message: "无法连接到PowerPoint应用程序" }
        }

        ; 检查是否有活动窗口
        try {
            window := ppt_application.ActiveWindow
            if (!window) {
                throw "没有活动窗口"
            }
        } catch as err {
            return { success: false, message: "没有活动的PowerPoint窗口" }
        }

        ; 坐标转换：屏幕像素 -> PPT点位
        try {
            origin_x := window.PointsToScreenPixelsX(0)
            origin_y := window.PointsToScreenPixelsY(0)
            zoom := window.View.Zoom / 100

            ; 转换为PowerPoint点位
            ppt_x := pixel_to_point((region_x - origin_x) / zoom, 88)
            ppt_y := pixel_to_point((region_y - origin_y) / zoom, 90)
            ppt_w := pixel_to_point(region_width / zoom, 88)
            ppt_h := pixel_to_point(region_height / zoom, 90)

            return {
                success: true,
                ppt_x: ppt_x,
                ppt_y: ppt_y,
                ppt_w: ppt_w,
                ppt_h: ppt_h,
                screen_region: Format("{}|{}|{}|{}", region_x, region_y, region_width, region_height),
                duration: total_duration,
                window: window,
                ppt_app: ppt_application
            }
        } catch as err {
            return { success: false, message: "坐标转换失败: " . err.message }
        }

    } catch as err {
        Msgbox("区域选择错误: " . err.Message)
        return { success: false, message: "区域选择过程出错: " . err.Message }
    } finally {
        ; 确保GUI被销毁
        if (selection_gui) {
            selection_gui.Destroy()
        }
        if (info_gui) {
            info_gui.Destroy()
        }
    }
}

; ; 主函数：插入形状
; insert_shape(shape_type := 1, color := "Purple", button := "RButton") {
;     ; 第一步：屏幕取点
;     region_data := get_ppt_region(color, button)
;     if (!region_data.success) {
;         return {success: false, file_path: "", message: region_data.message}
;     }

;     ; 第三步：绘制形状
;     result := draw_shape(region_data, shape_type)

;     return {success: result.success, file_path: "", message: result.message}
; }

; #endregion

; #region copy_format_from_sample
/**
 * 从指定形状复制格式到当前选中的所有形状
 * @param source_name 源形状名称
 * @param source_slide 源形状所在幻灯片ID（默认为当前页）
 * @param options 选项对象，支持:
 *   - keep_font_size: 保留原字体大小（默认false）
 *   - rename: 是否重命名形状（默认true）
 *   - silent: 静默模式，不显示成功提示（默认false）
 * @return 成功返回应用格式的形状数量，失败返回0
 */
copy_format_from_sample2(source_name, source_slide := "", options := "") {
    ; 默认选项
    opts := {
        keep_font_size: false,
        rename: true,
        silent: false
    }

    ; 合并用户选项
    if (IsObject(options)) {
        for key, value in options.OwnProps() {
            opts.%key% := value
        }
    }

    try {
        ; 获取PowerPoint应用对象
        try {
            ppt_app := ComObjActive("PowerPoint.Application")
        } catch {
            MsgBox("PowerPoint未运行！", "错误", "Icon!")
            return 0
        }

        ; 检查是否有活动演示文稿
        if (ppt_app.Presentations.Count = 0) {
            MsgBox("没有打开的PowerPoint文档！", "错误", "Icon!")
            return 0
        }

        ; 获取活动窗口的选择
        selection := ppt_app.ActiveWindow.Selection

        ; 检查是否选中了形状
        if (selection.Type != 2 && selection.Type != 3) {  ; ppSelectionShapes=2, ppSelectionText=3
            Msgbox("请先选中至少一个形状！")
            return 0
        }

        ; 获取选中的形状范围（支持组内形状）
        shape_range := (selection.HasChildShapeRange) ? selection.ChildShapeRange : selection.ShapeRange

        ; 如果未指定幻灯片ID，使用当前页
        if (source_slide = "") {
            source_slide := ppt_app.ActiveWindow.View.Slide.SlideIndex
        }

        ; 获取样本形状并拾取格式
        slides := ppt_app.ActivePresentation.Slides

        ; 验证幻灯片是否存在
        if (source_slide < 1 || source_slide > slides.Count) {
            MsgBox("幻灯片ID超出范围！`n有效范围: 1-" slides.Count, "错误", "Icon!")
            return 0
        }

        ; 验证样本形状是否存在
        try {
            source_shape := slides(source_slide).Shapes(source_name)
        } catch {
            Msgbox("未找到源形状: " source_name "@slide_" source_slide)
            return 0
        }

        ; 拾取样本格式
        source_shape.PickUp()

        ; 应用格式到所有选中的形状
        count := 0
        failed := 0

        for shape in shape_range {
            try {
                ; 保存字体大小
                font_size := 0
                if (opts.keep_font_size) {
                    try {
                        if (shape.HasTextFrame) {
                            font_size := shape.TextFrame.TextRange.Font.Size
                        }
                    }
                }

                ; 应用格式
                shape.Apply()

                ; 恢复字体大小
                if (opts.keep_font_size && font_size > 0) {
                    try {
                        shape.TextFrame.TextRange.Font.Size := font_size
                    }
                }

                ; 重命名形状（每次循环获取新的TickCount确保唯一）
                if (opts.rename) {
                    shape.Name := source_name . "_" . Random(1000, 9999)
                }

                count++

            } catch as err {
                failed++
                ; 可选：记录失败的形状
                ; OutputDebug("应用格式失败: " shape.Name " - " err.Message)
            }
        }

        ; 显示结果
        if (!opts.silent) {
            msg := "格式复制完成！`n成功: " count " 个形状"
            if (failed > 0) {
                msg .= "`n失败: " failed " 个形状"
            }
            if (opts.keep_font_size) {
                msg .= "`n（已保留字体大小）"
            }
            ; MsgBox(msg, "成功", "Iconi")
        }

        return count

    } catch as err {
        MsgBox("错误: " err.Message "`n`n详情: " err.Extra, "错误", "Icon!")
        return 0
    }
}

/**
 * 快捷函数：复制格式并保留字体大小
 */
copy_format_keep_fontsize(source_name, source_slide := "") {
    return copy_format_from_sample(source_name, source_slide, { keep_font_size: true })
}

/**
 * 快捷函数：静默复制格式（不显示提示）
 */
copy_format_silent(source_name, source_slide := "") {
    return copy_format_from_sample(source_name, source_slide, { silent: true })
}

/**
 * 获取当前页所有形状名称列表
 * @param slide_id 幻灯片ID（默认当前页）
 * @return 形状名称数组
 */
get_shape_names(slide_id := "") {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")

        if (slide_id = "") {
            slide_id := ppt_app.ActiveWindow.View.Slide.SlideIndex
        }

        shapes := ppt_app.ActivePresentation.Slides(slide_id).Shapes
        names := []

        for shape in shapes {
            names.Push(shape.Name)
        }

        return names

    } catch {
        return []
    }
}

; ===== 使用示例 =====
/*
; 基础用法：从当前页复制
copy_format_from_sample("Rectangle 1")

; 从指定页复制
copy_format_from_sample("标准框", 3)

; 使用选项对象
copy_format_from_sample("标准框", 1, {
    keep_font_size: true,  ; 保留字体大小
    rename: false,         ; 不重命名
    silent: true           ; 静默模式
})

; 快捷函数
copy_format_keep_fontsize("标准框")  ; 保留字体大小
copy_format_silent("标准框", 2)      ; 静默模式

; 获取形状列表
names := get_shape_names()
for name in names {
    MsgBox(name)
}

; 快捷键绑定
^!f::copy_format_from_sample("标准样式")           ; Ctrl+Alt+F
^!g::copy_format_keep_fontsize("标准样式")        ; Ctrl+Alt+G
^!+f::copy_format_silent("标准样式")              ; Ctrl+Alt+Shift+F
*/

/**
 * 从指定形状复制格式到当前选中的所有形状
 * @param source_name 源形状名称
 * @param source_slide 源形状所在幻灯片ID（默认为当前页）
 * @param options 选项对象，支持:
 *   - keep_font_size: 保留原字体大小（默认false）
 *   - rename: 是否重命名形状（默认true）
 * @return 成功返回应用格式的形状数量，失败返回0
 */
copy_format_from_sample(source_name, source_slide := "", options := "") {
    ; 默认选项
    opts := {
        keep_font_size: false,
        rename: true
    }

    ; 合并用户选项
    if (IsObject(options)) {
        for key, value in options.OwnProps() {
            opts.%key% := value
        }
    }

    try {
        ; 获取PowerPoint应用对象
        try {
            ppt_app := ComObjActive("PowerPoint.Application")
        } catch {
            Msgbox("PowerPoint未运行！")
            return 0
        }

        ; 检查是否有活动演示文稿
        if (ppt_app.Presentations.Count = 0) {
            Msgbox("没有打开的PowerPoint文档！")
            return 0
        }

        ; 获取活动窗口的选择
        selection := ppt_app.ActiveWindow.Selection

        ; 检查是否选中了形状
        if (selection.Type != 2 && selection.Type != 3) {  ; ppSelectionShapes=2, ppSelectionText=3
            Msgbox("请先选中至少一个形状！")
            return 0
        }

        ; 获取选中的形状范围（支持组内形状）
        shape_range := (selection.HasChildShapeRange) ? selection.ChildShapeRange : selection.ShapeRange

        ; 如果未指定幻灯片ID，使用当前页
        if (source_slide = "") {
            source_slide := ppt_app.ActiveWindow.View.Slide.SlideIndex
        }

        ; 获取样本形状并拾取格式
        slides := ppt_app.ActivePresentation.Slides

        ; 验证幻灯片是否存在
        if (source_slide < 1 || source_slide > slides.Count) {
            Msgbox("幻灯片ID超出范围！`n有效范围: 1-" slides.Count)
            return 0
        }

        ; 验证样本形状是否存在
        try {
            source_shape := slides(source_slide).Shapes(source_name)
        } catch {
            Msgbox("未找到源形状: " source_name "@slide_" source_slide)
            return 0
        }

        ; 拾取样本格式
        source_shape.PickUp()

        ; 应用格式到所有选中的形状
        count := 0
        failed := 0

        for shape in shape_range {
            try {
                ; 保存字体大小
                font_size := 0
                if (opts.keep_font_size) {
                    try {
                        if (shape.HasTextFrame) {
                            font_size := shape.TextFrame.TextRange.Font.Size
                        }
                    }
                }

                ; 应用格式
                shape.Apply()

                ; 恢复字体大小
                if (opts.keep_font_size && font_size > 0) {
                    try {
                        shape.TextFrame.TextRange.Font.Size := font_size
                    }
                }

                ; 重命名形状（每次循环获取新的TickCount确保唯一）
                if (opts.rename) {
                    shape.Name := source_name . "_" . Random(1000, 9999)
                }

                count++

            } catch as err {
                failed++
                ; 可选：记录失败的形状
                ; OutputDebug("应用格式失败: " shape.Name " - " err.Message)
            }
        }

        ; 显示结果
        ; if (count > 0) {
        ;     msg := "格式复制完成！`n成功: " count " 个形状"
        ;     if (failed > 0) {
        ;         msg .= "`n失败: " failed " 个形状"
        ;     }
        ;     if (opts.keep_font_size) {
        ;         msg .= "`n（已保留字体大小）"
        ;     }
        ;     Msgbox(msg)
        ; } else {
        ;     Msgbox("格式复制失败！没有成功应用任何形状。")
        ; }

        return count

    } catch as err {
        Msgbox("错误: " err.Message "`n`n详情: " err.Extra)
        return 0
    }
}

; #endregion

; #region format_sample_shape_01

; ===== 形状样式函数 (sample_shape_01 ~ sample_shape_39) =====
format_sample_shape_01() {
    return copy_format_from_sample("sample_shape_01", 1)
}
format_sample_shape_02() {
    return copy_format_from_sample("sample_shape_02", 1)
}
format_sample_shape_03() {
    return copy_format_from_sample("sample_shape_03", 1)
}
format_sample_shape_04() {
    return copy_format_from_sample("sample_shape_04", 1)
}
format_sample_shape_05() {
    return copy_format_from_sample("sample_shape_05", 1)
}
format_sample_shape_06() {
    return copy_format_from_sample("sample_shape_06", 1)
}
format_sample_shape_07() {
    return copy_format_from_sample("sample_shape_07", 1)
}
format_sample_shape_08() {
    return copy_format_from_sample("sample_shape_08", 1)
}
format_sample_shape_09() {
    return copy_format_from_sample("sample_shape_09", 1)
}
format_sample_shape_10() {
    return copy_format_from_sample("sample_shape_10", 1)
}
format_sample_shape_11() {
    return copy_format_from_sample("sample_shape_11", 1)
}
format_sample_shape_12() {
    return copy_format_from_sample("sample_shape_12", 1)
}
format_sample_shape_13() {
    return copy_format_from_sample("sample_shape_13", 1)
}
format_sample_shape_14() {
    return copy_format_from_sample("sample_shape_14", 1)
}
format_sample_shape_15() {
    return copy_format_from_sample("sample_shape_15", 1)
}
format_sample_shape_16() {
    return copy_format_from_sample("sample_shape_16", 1)
}
format_sample_shape_17() {
    return copy_format_from_sample("sample_shape_17", 1)
}
format_sample_shape_18() {
    return copy_format_from_sample("sample_shape_18", 1)
}
format_sample_shape_19() {
    return copy_format_from_sample("sample_shape_19", 1)
}
format_sample_shape_20() {
    return copy_format_from_sample("sample_shape_20", 1)
}
format_sample_shape_21() {
    return copy_format_from_sample("sample_shape_21", 1)
}
format_sample_shape_22() {
    return copy_format_from_sample("sample_shape_22", 1)
}
format_sample_shape_23() {
    return copy_format_from_sample("sample_shape_23", 1)
}
format_sample_shape_24() {
    return copy_format_from_sample("sample_shape_24", 1)
}
format_sample_shape_25() {
    return copy_format_from_sample("sample_shape_25", 1)
}
format_sample_shape_26() {
    return copy_format_from_sample("sample_shape_26", 1)
}
format_sample_shape_27() {
    return copy_format_from_sample("sample_shape_27", 1)
}
format_sample_shape_28() {
    return copy_format_from_sample("sample_shape_28", 1)
}
format_sample_shape_29() {
    return copy_format_from_sample("sample_shape_29", 1)
}
format_sample_shape_30() {
    return copy_format_from_sample("sample_shape_30", 1)
}
format_sample_shape_31() {
    return copy_format_from_sample("sample_shape_31", 1)
}
format_sample_shape_32() {
    return copy_format_from_sample("sample_shape_32", 1)
}
format_sample_shape_33() {
    return copy_format_from_sample("sample_shape_33", 1)
}
format_sample_shape_34() {
    return copy_format_from_sample("sample_shape_34", 1)
}
format_sample_shape_35() {
    return copy_format_from_sample("sample_shape_35", 1)
}
format_sample_shape_36() {
    return copy_format_from_sample("sample_shape_36", 1)
}
format_sample_shape_37() {
    return copy_format_from_sample("sample_shape_37", 1)
}
format_sample_shape_38() {
    return copy_format_from_sample("sample_shape_38", 1)
}
format_sample_shape_39() {
    return copy_format_from_sample("sample_shape_39", 1)
}

; ===== 段落样式函数 (sample_paragraph_01 ~ sample_paragraph_09) =====
format_sample_paragraph_01() {
    return copy_format_from_sample("sample_paragraph_01", 1)
}
format_sample_paragraph_02() {
    return copy_format_from_sample("sample_paragraph_02", 1)
}
format_sample_paragraph_03() {
    return copy_format_from_sample("sample_paragraph_03", 1)
}
format_sample_paragraph_04() {
    return copy_format_from_sample("sample_paragraph_04", 1)
}
format_sample_paragraph_05() {
    return copy_format_from_sample("sample_paragraph_05", 1)
}
format_sample_paragraph_06() {
    return copy_format_from_sample("sample_paragraph_06", 1)
}
format_sample_paragraph_07() {
    return copy_format_from_sample("sample_paragraph_07", 1)
}
format_sample_paragraph_08() {
    return copy_format_from_sample("sample_paragraph_08", 1)
}
format_sample_paragraph_09() {
    return copy_format_from_sample("sample_paragraph_09", 1)
}

; ===== 标题样式函数 (sample_title_01 ~ sample_title_39) =====
format_sample_title_01() {
    return copy_format_from_sample("sample_title_01", 1)
}
format_sample_title_02() {
    return copy_format_from_sample("sample_title_02", 1)
}
format_sample_title_03() {
    return copy_format_from_sample("sample_title_03", 1)
}
format_sample_title_04() {
    return copy_format_from_sample("sample_title_04", 1)
}
format_sample_title_05() {
    return copy_format_from_sample("sample_title_05", 1)
}
format_sample_title_06() {
    return copy_format_from_sample("sample_title_06", 1)
}
format_sample_title_07() {
    return copy_format_from_sample("sample_title_07", 1)
}
format_sample_title_08() {
    return copy_format_from_sample("sample_title_08", 1)
}
format_sample_title_09() {
    return copy_format_from_sample("sample_title_09", 1)
}
format_sample_title_10() {
    return copy_format_from_sample("sample_title_10", 1)
}
format_sample_title_11() {
    return copy_format_from_sample("sample_title_11", 1)
}
format_sample_title_12() {
    return copy_format_from_sample("sample_title_12", 1)
}
format_sample_title_13() {
    return copy_format_from_sample("sample_title_13", 1)
}
format_sample_title_14() {
    return copy_format_from_sample("sample_title_14", 1)
}
format_sample_title_15() {
    return copy_format_from_sample("sample_title_15", 1)
}
format_sample_title_16() {
    return copy_format_from_sample("sample_title_16", 1)
}
format_sample_title_17() {
    return copy_format_from_sample("sample_title_17", 1)
}
format_sample_title_18() {
    return copy_format_from_sample("sample_title_18", 1)
}
format_sample_title_19() {
    return copy_format_from_sample("sample_title_19", 1)
}
format_sample_title_20() {
    return copy_format_from_sample("sample_title_20", 1)
}
format_sample_title_21() {
    return copy_format_from_sample("sample_title_21", 1)
}
format_sample_title_22() {
    return copy_format_from_sample("sample_title_22", 1)
}
format_sample_title_23() {
    return copy_format_from_sample("sample_title_23", 1)
}
format_sample_title_24() {
    return copy_format_from_sample("sample_title_24", 1)
}
format_sample_title_25() {
    return copy_format_from_sample("sample_title_25", 1)
}
format_sample_title_26() {
    return copy_format_from_sample("sample_title_26", 1)
}
format_sample_title_27() {
    return copy_format_from_sample("sample_title_27", 1)
}
format_sample_title_28() {
    return copy_format_from_sample("sample_title_28", 1)
}
format_sample_title_29() {
    return copy_format_from_sample("sample_title_29", 1)
}
format_sample_title_30() {
    return copy_format_from_sample("sample_title_30", 1)
}
format_sample_title_31() {
    return copy_format_from_sample("sample_title_31", 1)
}
format_sample_title_32() {
    return copy_format_from_sample("sample_title_32", 1)
}
format_sample_title_33() {
    return copy_format_from_sample("sample_title_33", 1)
}
format_sample_title_34() {
    return copy_format_from_sample("sample_title_34", 1)
}
format_sample_title_35() {
    return copy_format_from_sample("sample_title_35", 1)
}
format_sample_title_36() {
    return copy_format_from_sample("sample_title_36", 1)
}
format_sample_title_37() {
    return copy_format_from_sample("sample_title_37", 1)
}
format_sample_title_38() {
    return copy_format_from_sample("sample_title_38", 1)
}
format_sample_title_39() {
    return copy_format_from_sample("sample_title_39", 1)
}

; ===== 图标样式函数 (sample_icon_01 ~ sample_icon_09) =====
format_sample_icon_01() {
    return copy_format_from_sample("sample_icon_01", 1)
}
format_sample_icon_02() {
    return copy_format_from_sample("sample_icon_02", 1)
}
format_sample_icon_03() {
    return copy_format_from_sample("sample_icon_03", 1)
}
format_sample_icon_04() {
    return copy_format_from_sample("sample_icon_04", 1)
}
format_sample_icon_05() {
    return copy_format_from_sample("sample_icon_05", 1)
}
format_sample_icon_06() {
    return copy_format_from_sample("sample_icon_06", 1)
}
format_sample_icon_07() {
    return copy_format_from_sample("sample_icon_07", 1)
}
format_sample_icon_08() {
    return copy_format_from_sample("sample_icon_08", 1)
}
format_sample_icon_09() {
    return copy_format_from_sample("sample_icon_09", 1)
}

; ===== 图表样式函数 (sample_chart_01 ~ sample_chart_09) =====
format_sample_chart_01() {
    return copy_format_from_sample("sample_chart_01", 1)
}
format_sample_chart_02() {
    return copy_format_from_sample("sample_chart_02", 1)
}
format_sample_chart_03() {
    return copy_format_from_sample("sample_chart_03", 1)
}
format_sample_chart_04() {
    return copy_format_from_sample("sample_chart_04", 1)
}
format_sample_chart_05() {
    return copy_format_from_sample("sample_chart_05", 1)
}
format_sample_chart_06() {
    return copy_format_from_sample("sample_chart_06", 1)
}
format_sample_chart_07() {
    return copy_format_from_sample("sample_chart_07", 1)
}
format_sample_chart_08() {
    return copy_format_from_sample("sample_chart_08", 1)
}
format_sample_chart_09() {
    return copy_format_from_sample("sample_chart_09", 1)
}

; ===== 表格样式函数 (sample_table_01 ~ sample_table_09) =====
format_sample_table_01() {
    return copy_format_from_sample("sample_table_01", 1)
}
format_sample_table_02() {
    return copy_format_from_sample("sample_table_02", 1)
}
format_sample_table_03() {
    return copy_format_from_sample("sample_table_03", 1)
}
format_sample_table_04() {
    return copy_format_from_sample("sample_table_04", 1)
}
format_sample_table_05() {
    return copy_format_from_sample("sample_table_05", 1)
}
format_sample_table_06() {
    return copy_format_from_sample("sample_table_06", 1)
}
format_sample_table_07() {
    return copy_format_from_sample("sample_table_07", 1)
}
format_sample_table_08() {
    return copy_format_from_sample("sample_table_08", 1)
}
format_sample_table_09() {
    return copy_format_from_sample("sample_table_09", 1)
}

; #endregion

; #region  tag

/**
 * 对话框函数：通过对话框为选中对象添加标签（包含标签名和标签内容）
 * @param {String} default_name 默认标签名（默认：""）
 * @param {String} default_value 默认标签内容（默认：""）
 * @param {Boolean} show_append_option 是否显示追加模式选项（默认：true）
 * @return {Object} 包含成功状态和操作结果信息
 */
add_tag_via_dialog(default_name := "", default_value := "", show_append_option := true) {
    start_time := A_TickCount

    try {
        ; 创建对话框
        tag_dialog := Gui("+AlwaysOnTop +ToolWindow", "添加对象标签")
        tag_dialog.SetFont("s10", "Segoe UI")

        ; 标签名输入框
        tag_dialog.Add("Text", "w300", "标签名：")
        name_edit := tag_dialog.Add("Edit", "w300 h25 vTagName", default_name)

        ; 标签内容输入框
        tag_dialog.Add("Text", "w300", "标签内容：")
        value_edit := tag_dialog.Add("Edit", "w300 h80 vTagValue", default_value)

        ; 追加模式选项
        append_checkbox := 0
        if (show_append_option) {
            append_checkbox := tag_dialog.Add("Checkbox", "vAppendMode", "追加模式（不覆盖原有标签）")
        }

        ; 按钮区域
        button_row := tag_dialog.Add("Button", "x10 w140 h30 Default", "确定")
        button_row.OnEvent("Click", button_ok)
        tag_dialog.Add("Button", "x+10 w140 h30", "取消").OnEvent("Click", button_cancel)

        ; 对话框结果
        dialog_result := ""
        tag_name := ""
        tag_value := ""
        append_mode := false

        ; 按钮事件处理
        button_ok(*) {
            tag_name := Trim(name_edit.Value)
            tag_value := Trim(value_edit.Value)
            if (show_append_option) {
                append_mode := append_checkbox.Value
            }
            dialog_result := "ok"
            tag_dialog.Destroy()
        }

        button_cancel(*) {
            dialog_result := "cancel"
            tag_dialog.Destroy()
        }

        ; 显示对话框
        tag_dialog.Show("AutoSize Center")

        ; 等待对话框关闭
        while (dialog_result = "") {
            Sleep(10)
        }

        ; 处理取消情况
        if (dialog_result = "cancel") {
            return { success: false, message: "用户取消了标签添加" }
        }

        ; 验证输入
        if (tag_name = "") {
            return { success: false, message: "标签名不能为空" }
        }

        if (tag_value = "") {
            return { success: false, message: "标签内容不能为空" }
        }

        ; 调用核心函数添加标签
        result := add_tag_to_selection(tag_name, tag_value, append_mode)
        result.dialog_used := true
        result.duration := A_TickCount - start_time

        return result

    } catch as err {
        return { success: false, message: "对话框操作失败: " err.Message }
    }
}

/**
 * 核心函数：为PowerPoint选中对象添加标签（支持标签名和标签内容）
 * @param {String} tag_name 标签名
 * @param {String} tag_value 标签内容
 * @param {Boolean} append_mode 是否追加模式（默认：false，覆盖原有标签）
 * @return {Object} 包含成功状态和操作结果信息
 */
add_tag_to_selection(tag_name, tag_value, append_mode := false) {
    start_time := A_TickCount
    ppt_application := 0
    processed_count := 0

    try {
        ; 参数验证
        if (tag_name = "") {
            return { success: false, message: "标签名不能为空" }
        }

        if (tag_value = "") {
            return { success: false, message: "标签内容不能为空" }
        }

        ; 连接到PowerPoint应用程序
        ppt_application := ComObjActive("PowerPoint.Application")
        if (!ppt_application.ActivePresentation) {
            return { success: false, message: "没有活动的PowerPoint演示文稿" }
        }

        ; 检查是否有选中的对象
        try {
            selection := ppt_application.ActiveWindow.Selection
            if (!selection || selection.Type = 0) { ; ppSelectionNone
                return { success: false, message: "没有选中的对象" }
            }
        } catch as err {
            return { success: false, message: "无法获取选中对象: " err.Message }
        }

        ; 处理选中的对象
        processed_objects := []
        object_count := selection.Count

        loop object_count {
            try {
                shape := selection.Item(A_Index)

                ; 获取当前标签内容
                current_value := ""
                try {
                    current_value := shape.Tags[tag_name]
                } catch as err {
                    ; 如果没有该标签名的内容，继续处理
                }

                ; 构建新标签内容
                new_value := ""
                if (append_mode && current_value != "") {
                    new_value := current_value ";" tag_value
                } else {
                    new_value := tag_value
                }

                ; 删除原有标签（如果存在）
                try {
                    shape.Tags.Delete(tag_name)
                } catch as err {
                    ; 标签可能不存在，忽略错误
                }

                ; 添加新标签
                shape.Tags.Add(tag_name, new_value)
                processed_count += 1

                ; 记录处理的对象信息
                object_info := {
                    index: A_Index,
                    name: shape.Name,
                    type: shape.Type,
                    tag_name: tag_name,
                    old_value: current_value,
                    new_value: new_value
                }
                processed_objects.Push(object_info)

            } catch as err {
                ; 跳过无法处理的对象，继续处理下一个
                processed_objects.Push({
                    index: A_Index,
                    error: err.Message
                })
            }
        }

        ; 返回结果
        total_duration := A_TickCount - start_time

        return {
            success: true,
            message: "标签添加成功",
            tag_name: tag_name,
            tag_value: tag_value,
            append_mode: append_mode,
            processed_count: processed_count,
            total_objects: object_count,
            processed_objects: processed_objects,
            duration: total_duration,
            presentation: ppt_application.ActivePresentation.Name
        }

    } catch as err {
        return { success: false, message: "添加标签失败: " err.Message }
    }
}

/**
 * 查询函数：获取选中对象的所有标签信息
 * @return {Object} 包含选中对象的所有标签信息
 */
get_selection_all_tags() {
    start_time := A_TickCount
    ppt_application := 0

    try {
        ; 连接到PowerPoint应用程序
        ppt_application := ComObjActive("PowerPoint.Application")
        if (!ppt_application.ActivePresentation) {
            return { success: false, message: "没有活动的PowerPoint演示文稿" }
        }

        ; 检查选中的对象
        selection := ppt_application.ActiveWindow.Selection
        if (!selection || selection.Type = 0) {
            return { success: false, message: "没有选中的对象" }
        }

        ; 收集所有标签信息
        objects_with_tags := []
        object_count := selection.Count

        loop object_count {
            try {
                shape := selection.Item(A_Index)
                tags_info := []

                ; 获取所有标签
                try {
                    tags_count := shape.Tags.Count
                    loop tags_count {
                        tag_name := shape.Tags.Name(A_Index)
                        tag_value := shape.Tags.Value(A_Index)
                        tags_info.Push({
                            name: tag_name,
                            value: tag_value
                        })
                    }
                } catch as err {
                    ; 没有标签
                }

                object_info := {
                    index: A_Index,
                    name: shape.Name,
                    type: shape.Type,
                    tags: tags_info,
                    tags_count: tags_info.Length
                }
                objects_with_tags.Push(object_info)

            } catch as err {
                objects_with_tags.Push({
                    index: A_Index,
                    error: err.Message
                })
            }
        }

        total_duration := A_TickCount - start_time

        return {
            success: true,
            total_objects: object_count,
            objects_with_tags: objects_with_tags,
            total_tags_count: objects_with_tags.Sum(obj => obj.tags_count),
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "获取标签信息失败: " err.Message }
    }
}

/**
 * 查询函数：获取选中对象的指定标签信息
 * @param {String} tag_name 要查询的标签名
 * @return {Object} 包含指定标签的信息
 */
get_selection_tag_by_name(tag_name) {
    start_time := A_TickCount
    ppt_application := 0

    try {
        if (tag_name = "") {
            return { success: false, message: "标签名不能为空" }
        }

        ; 连接到PowerPoint应用程序
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection

        if (!selection || selection.Type = 0) {
            return { success: false, message: "没有选中的对象" }
        }

        objects_with_tag := []
        object_count := selection.Count

        loop object_count {
            try {
                shape := selection.Item(A_Index)
                tag_value := ""
                has_tag := false

                try {
                    tag_value := shape.Tags[tag_name]
                    has_tag := (tag_value != "")
                } catch as err {
                    ; 没有该标签
                }

                object_info := {
                    index: A_Index,
                    name: shape.Name,
                    type: shape.Type,
                    tag_name: tag_name,
                    tag_value: tag_value,
                    has_tag: has_tag
                }
                objects_with_tag.Push(object_info)

            } catch as err {
                objects_with_tag.Push({
                    index: A_Index,
                    error: err.Message
                })
            }
        }

        total_duration := A_TickCount - start_time

        return {
            success: true,
            tag_name: tag_name,
            total_objects: object_count,
            objects_with_tag: objects_with_tag,
            tagged_count: objects_with_tag.Filter(obj => obj.has_tag).Length,
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "获取标签信息失败: " err.Message }
    }
}

/**
 * 删除函数：删除选中对象的指定标签
 * @param {String} tag_name 要删除的标签名
 * @return {Object} 包含成功状态和操作结果信息
 */
delete_selection_tag(tag_name) {
    start_time := A_TickCount
    ppt_application := 0
    deleted_count := 0

    try {
        if (tag_name = "") {
            return { success: false, message: "标签名不能为空" }
        }

        ; 连接到PowerPoint应用程序
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection

        if (!selection || selection.Type = 0) {
            return { success: false, message: "没有选中的对象" }
        }

        object_count := selection.Count
        deleted_objects := []

        loop object_count {
            try {
                shape := selection.Item(A_Index)

                ; 删除指定标签
                try {
                    shape.Tags.Delete(tag_name)
                    deleted_count += 1
                    deleted_objects.Push({
                        index: A_Index,
                        name: shape.Name,
                        success: true
                    })
                } catch as err {
                    ; 标签可能不存在
                    deleted_objects.Push({
                        index: A_Index,
                        name: shape.Name,
                        success: false,
                        error: "标签不存在"
                    })
                }

            } catch as err {
                deleted_objects.Push({
                    index: A_Index,
                    error: err.Message
                })
            }
        }

        total_duration := A_TickCount - start_time

        return {
            success: true,
            message: "标签删除成功",
            tag_name: tag_name,
            deleted_count: deleted_count,
            total_objects: object_count,
            deleted_objects: deleted_objects,
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "删除标签失败: " err.Message }
    }
}

; #endregion

; #region  show tags
/**
 * 显示函数：显示选中对象的标签信息对话框
 * @param {Boolean} show_all 是否显示所有标签（默认：true），false时只显示常用标签
 * @return {Object} 包含成功状态和显示结果信息
 */
show_tags_info(show_all := true) {
    start_time := A_TickCount

    try {
        ; 获取标签信息
        tags_result := get_selection_all_tags()
        if (!tags_result.success) {
            return tags_result
        }

        ; 如果没有选中的对象
        if (tags_result.total_objects = 0) {
            return { success: false, message: "没有选中的对象" }
        }

        ; 创建显示对话框
        info_dialog := Gui("+AlwaysOnTop +Resize", "对象标签信息")
        info_dialog.SetFont("s9", "Segoe UI")

        ; 统计信息
        total_tags := tags_result.total_tags_count
        tagged_objects := tags_result.objects_with_tags.Filter(obj => obj.tags_count > 0).Length

        ; 标题区域
        info_dialog.Add("Text", "w400", "选中对象标签信息：")
        info_dialog.Add("Text", "w400", Format("选中对象: {}个 | 有标签的对象: {}个 | 总标签数: {}个",
            tags_result.total_objects, tagged_objects, total_tags))
        info_dialog.Add("Text", "w400", "────────────────────────────────────")

        ; 创建列表显示区域
        if (total_tags > 0) {
            ; 有标签的情况
            for obj_index, obj in tags_result.objects_with_tags {
                if (obj.tags_count > 0) {
                    ; 对象标题
                    obj_header := info_dialog.Add("Text", "w400 y+10", Format("对象 {}: {} (类型: {})",
                        obj.index, obj.name, get_shape_type_name(obj.type)))
                    obj_header.SetFont("s9 cNavy", "Segoe UI")

                    ; 标签列表
                    for tag_index, tag in obj.tags {
                        tag_text := info_dialog.Add("Text", "x20 w380", Format("  • {}: {}", tag.name, tag.value))
                        if (tag_index & 1) { ; 交替背景色
                            tag_text.BackColor := 0xF5F5F5
                        }
                    }
                } else {
                    ; 无标签的对象
                    no_tag_text := info_dialog.Add("Text", "w400 y+10", Format("对象 {}: {} - 无标签",
                        obj.index, obj.name))
                    no_tag_text.SetFont("s9 cGray", "Segoe UI")
                }

                ; 对象间分隔线
                if (obj_index < tags_result.objects_with_tags.Length) {
                    info_dialog.Add("Text", "w400", "────────────────────────────────────")
                }
            }
        } else {
            ; 没有标签的情况
            info_dialog.Add("Text", "w400 y+20 Center", "选中的对象都没有标签")
            info_dialog.Add("Text", "w400 Center", "使用添加标签功能为对象添加标签")
        }

        ; 按钮区域
        button_row := info_dialog.Add("Button", "x10 w120 h30", "添加标签")
        button_row.OnEvent("Click", button_add_tag)

        info_dialog.Add("Button", "x+10 w120 h30", "刷新")
        info_dialog.OnEvent("Click", button_refresh)

        info_dialog.Add("Button", "x+10 w120 h30", "关闭")
        info_dialog.OnEvent("Click", button_close)

        ; 按钮事件处理
        button_add_tag(*) {
            info_dialog.Destroy()
            add_tag_via_dialog()
        }

        button_refresh(*) {
            info_dialog.Destroy()
            show_tags_info(show_all)
        }

        button_close(*) {
            info_dialog.Destroy()
        }

        ; 显示对话框（根据内容调整大小）
        dialog_height := Min(600, 200 + (tags_result.total_objects * 40) + (total_tags * 20))
        info_dialog.Show("w420 h" dialog_height " Center")

        total_duration := A_TickCount - start_time

        return {
            success: true,
            message: "标签信息显示成功",
            total_objects: tags_result.total_objects,
            tagged_objects: tagged_objects,
            total_tags: total_tags,
            dialog_shown: true,
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "显示标签信息失败: " err.Message }
    }
}

/**
 * 显示函数：简洁显示标签信息（用于状态栏或提示）
 * @return {Object} 包含简洁的标签统计信息
 */
show_tags_summary() {
    start_time := A_TickCount

    try {
        tags_result := get_selection_all_tags()
        if (!tags_result.success) {
            return tags_result
        }

        total_tags := tags_result.total_tags_count
        tagged_objects := tags_result.objects_with_tags.Filter(obj => obj.tags_count > 0).Length

        ; 创建简洁信息对话框
        summary_dialog := Gui("+AlwaysOnTop +ToolWindow", "标签统计")
        summary_dialog.SetFont("s10", "Segoe UI")

        summary_dialog.Add("Text", "w300 Center", "📊 标签统计")
        summary_dialog.Add("Text", "w300", "────────────────────")

        summary_dialog.Add("Text", "w300", Format("选中对象: {}个", tags_result.total_objects))
        summary_dialog.Add("Text", "w300", Format("有标签对象: {}个", tagged_objects))
        summary_dialog.Add("Text", "w300", Format("总标签数: {}个", total_tags))

        if (total_tags > 0) {
            ; 统计常用标签名
            tag_names := Map()
            for obj in tags_result.objects_with_tags {
                for tag in obj.tags {
                    if (tag_names.Has(tag.name)) {
                        tag_names[tag.name] += 1
                    } else {
                        tag_names[tag.name] := 1
                    }
                }
            }

            summary_dialog.Add("Text", "w300", "────────────────────")
            summary_dialog.Add("Text", "w300", "标签类型分布:")

            for name, count in tag_names {
                summary_dialog.Add("Text", "x20 w280", Format("  • {}: {}个", name, count))
            }
        }

        summary_dialog.Add("Button", "w280 h30", "查看详情").OnEvent("Click", button_details)
        summary_dialog.OnEvent("Click", button_close_summary)

        button_details(*) {
            summary_dialog.Destroy()
            show_tags_info()
        }

        button_close_summary(*) {
            summary_dialog.Destroy()
        }

        summary_dialog.Show("AutoSize Center")

        total_duration := A_TickCount - start_time

        return {
            success: true,
            message: "标签统计显示成功",
            total_objects: tags_result.total_objects,
            tagged_objects: tagged_objects,
            total_tags: total_tags,
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "显示标签统计失败: " err.Message }
    }
}

/**
 * 内部函数：获取形状类型名称
 * @param {Number} shape_type 形状类型代码
 * @return {String} 形状类型名称
 */
get_shape_type_name(shape_type) {
    shape_types := Map(
        -2, "幻灯片",
        -1, "占位符",
        1, "表格",
        2, "图表",
        3, "图片",
        4, "剪贴画",
        5, "组织结构图",
        6, "媒体",
        7, "OLE对象",
        8, "SmartArt",
        9, "链接图片",
        10, "文本框",
        11, "线条",
        12, "矩形",
        13, "椭圆",
        14, "箭头",
        15, "多边形",
        16, "曲线",
        17, "自由形状",
        18, "连接符",
        19, "标注",
        20, "动作按钮",
        21, "组",
        22, "图表区",
        23, "注释",
        24, "墨迹",
        25, "墨迹注释",
        26, "脚本锚点",
        27, "图片框",
        28, "控件"
    )

    return shape_types.Has(shape_type) ? shape_types[shape_type] : "未知类型"
}

/**
 * 显示函数：显示特定标签名的对象列表
 * @param {String} tag_name 要显示的标签名
 * @return {Object} 包含成功状态和显示结果信息
 */
show_objects_by_tag(tag_name) {
    start_time := A_TickCount

    try {
        if (tag_name = "") {
            return { success: false, message: "标签名不能为空" }
        }

        tag_result := get_selection_tag_by_name(tag_name)
        if (!tag_result.success) {
            return tag_result
        }

        ; 创建标签特定显示对话框
        tag_dialog := Gui("+AlwaysOnTop", Format("标签 '{}' 的对象列表", tag_name))
        tag_dialog.SetFont("s9", "Segoe UI")

        tag_dialog.Add("Text", "w400", Format("标签 '{}' 的对象列表:", tag_name))
        tag_dialog.Add("Text", "w400", Format("共有 {} 个对象拥有此标签", tag_result.tagged_count))
        tag_dialog.Add("Text", "w400", "────────────────────────────────────")

        if (tag_result.tagged_count > 0) {
            for obj in tag_result.objects_with_tag {
                if (obj.has_tag) {
                    obj_text := tag_dialog.Add("Text", "w380", Format("• {} ({})", obj.name, get_shape_type_name(obj.type
                    )))
                    value_text := tag_dialog.Add("Text", "x20 w360", Format("  值: {}", obj.tag_value))
                    value_text.SetFont("s8 cGreen", "Segoe UI")

                    ; 分隔线
                    tag_dialog.Add("Text", "w400", "────────────────────────────────────")
                }
            }
        } else {
            tag_dialog.Add("Text", "w400 Center", Format("没有对象拥有标签 '{}'", tag_name))
        }

        ; 定义关闭按钮事件
        button_close(*) {
            tag_dialog.Destroy()
        }
        tag_dialog.Add("Button", "w380 h30", "关闭").OnEvent("Click", button_close)
        tag_dialog.OnEvent("Click", button_close)

        tag_dialog.Show("AutoSize Center")

        total_duration := A_TickCount - start_time

        return {
            success: true,
            message: "标签对象列表显示成功",
            tag_name: tag_name,
            tagged_count: tag_result.tagged_count,
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "显示标签对象列表失败: " err.Message }
    }
}

; #endregion

; #region  copy_header_shape
/**
 * 从指定幻灯片复制名称匹配的对象到当前页面
 * @param {String} name_pattern 名称匹配模式（支持通配符*）
 * @param {Number} source_slide 源幻灯片ID（默认为1）
 * @param {Object} options 选项对象，支持:
 *   - exact_match: 是否精确匹配（默认false，使用通配符匹配）
 *   - copy_format: 是否复制格式（默认true）
 *   - position_offset: 位置偏移量，如{x: 10, y: 10}（默认无偏移）
 *   - rename_copies: 是否重命名副本（默认true）
 * @return {Object} 包含成功状态和复制结果信息
 */
copy_shapes_by_name(name_pattern, source_slide := 1, options := "") {
    start_time := A_TickCount

    ; 默认选项
    opts := {
        exact_match: false,
        copy_format: true,
        position_offset: "",
        rename_copies: true
    }

    ; 合并用户选项
    if (IsObject(options)) {
        for key, value in options.OwnProps() {
            opts.%key% := value
        }
    }

    try {
        ; 获取PowerPoint应用对象
        ppt_app := ComObjActive("PowerPoint.Application")
        if (ppt_app.Presentations.Count = 0) {
            return { success: false, message: "没有打开的PowerPoint文档！" }
        }

        presentation := ppt_app.ActivePresentation
        slides := presentation.Slides

        ; 验证源幻灯片是否存在
        if (source_slide < 1 || source_slide > slides.Count) {
            return { success: false, message: "源幻灯片ID超出范围！有效范围: 1-" slides.Count }
        }

        source_slide_obj := slides(source_slide)
        current_slide := ppt_app.ActiveWindow.View.Slide

        ; 如果源幻灯片就是当前幻灯片，不需要复制
        if (source_slide_obj.SlideIndex = current_slide.SlideIndex) {
            return { success: false, message: "源幻灯片与当前幻灯片相同，无需复制" }
        }

        ; 查找匹配的形状
        matching_shapes := []
        shapes_count := source_slide_obj.Shapes.Count

        loop shapes_count {
            shape := source_slide_obj.Shapes(A_Index)
            shape_name := shape.Name

            ; 检查名称匹配
            is_match := false
            if (opts.exact_match) {
                is_match := (shape_name = name_pattern)
            } else {
                ; 通配符匹配：将*转换为正则表达式.*
                pattern := StrReplace(name_pattern, "*", ".*")
                is_match := RegExMatch(shape_name, "^" pattern "$")
            }

            if (is_match) {
                matching_shapes.Push({
                    name: shape_name,
                    shape: shape,
                    index: A_Index,
                    left: shape.Left,
                    top: shape.Top,
                    width: shape.Width,
                    height: shape.Height
                })
            }
        }

        ; 检查是否找到匹配的形状
        if (matching_shapes.Length = 0) {
            return {
                success: false,
                message: "未找到匹配的形状！",
                pattern: name_pattern,
                source_slide: source_slide,
                searched_count: shapes_count
            }
        }

        ; 复制形状到当前幻灯片
        copied_shapes := []
        failed_count := 0

        for match in matching_shapes {
            try {
                ; 复制形状
                match.shape.Copy()

                ; 粘贴到当前幻灯片
                current_slide.Shapes.Paste()

                ; 获取粘贴的形状（通常是最后一个形状）
                pasted_shape := current_slide.Shapes(current_slide.Shapes.Count)

                ; 应用位置偏移
                if (IsObject(opts.position_offset)) {
                    if (opts.position_offset.Has("x")) {
                        pasted_shape.Left := match.left + opts.position_offset.x
                    }
                    if (opts.position_offset.Has("y")) {
                        pasted_shape.Top := match.top + opts.position_offset.y
                    }
                } else {
                    ; 保持原位置
                    pasted_shape.Left := match.left
                    pasted_shape.Top := match.top
                }

                ; 重命名副本
                new_name := match.name
                if (opts.rename_copies) {
                    new_name := match.name . "_copy_" . Random(1000, 9999)
                    pasted_shape.Name := new_name
                }

                ; 复制格式（如果需要）
                if (opts.copy_format) {
                    try {
                        match.shape.PickUp()
                        pasted_shape.Apply()
                    } catch as format_err {
                        ; 格式复制失败不影响整体操作
                    }
                }

                copied_shapes.Push({
                    original_name: match.name,
                    new_name: new_name,
                    position: {
                        left: pasted_shape.Left,
                        top: pasted_shape.Top,
                        width: pasted_shape.Width,
                        height: pasted_shape.Height
                    },
                    success: true
                })

            } catch as err {
                failed_count++
                copied_shapes.Push({
                    original_name: match.name,
                    success: false,
                    error: err.Message
                })
            }
        }

        ; 返回结果
        total_duration := A_TickCount - start_time
        success_count := copied_shapes.Filter(shape => shape.success).Length

        return {
            success: true,
            message: "形状复制完成",
            pattern: name_pattern,
            source_slide: source_slide,
            current_slide: current_slide.SlideIndex,
            matching_count: matching_shapes.Length,
            copied_count: success_count,
            failed_count: failed_count,
            copied_shapes: copied_shapes,
            duration: total_duration,
            options: opts
        }

    } catch as err {
        return { success: false, message: "复制形状失败: " err.Message }
    }
}

/**
 * 便捷函数：复制header开头的对象从slide1到当前页面
 * @param {Object} options 选项对象
 * @return {Object} 包含成功状态和复制结果信息
 */
copy_header_shapes(options := "") {
    return copy_shapes_by_name("header*", 1, options)
}

/**
 * 便捷函数：复制footer开头的对象从slide1到当前页面
 * @param {Object} options 选项对象
 * @return {Object} 包含成功状态和复制结果信息
 */
copy_footer_shapes(options := "") {
    return copy_shapes_by_name("footer*", 1, options)
}

/**
 * 便捷函数：复制logo开头的对象从slide1到当前页面
 * @param {Object} options 选项对象
 * @return {Object} 包含成功状态和复制结果信息
 */
copy_logo_shapes(options := "") {
    return copy_shapes_by_name("logo*", 1, options)
}

/**
 * 查询函数：获取指定幻灯片中匹配名称的形状列表
 * @param {String} name_pattern 名称匹配模式
 * @param {Number} slide_index 幻灯片ID（默认为1）
 * @return {Object} 包含匹配的形状列表
 */
find_shapes_by_name(name_pattern, slide_index := 1) {
    start_time := A_TickCount

    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        presentation := ppt_app.ActivePresentation
        slide := presentation.Slides(slide_index)

        matching_shapes := []
        shapes_count := slide.Shapes.Count

        loop shapes_count {
            shape := slide.Shapes(A_Index)
            shape_name := shape.Name

            if (InStr(shape_name, name_pattern)) {
                matching_shapes.Push({
                    name: shape_name,
                    index: A_Index,
                    type: get_shape_type_name(shape.Type),
                    position: {
                        left: shape.Left,
                        top: shape.Top,
                        width: shape.Width,
                        height: shape.Height
                    },
                    visible: shape.Visible
                })
            }
        }

        total_duration := A_TickCount - start_time

        return {
            success: true,
            pattern: name_pattern,
            slide_index: slide_index,
            matching_count: matching_shapes.Length,
            total_shapes: shapes_count,
            matching_shapes: matching_shapes,
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "查找形状失败: " err.Message }
    }
}

; #endregion

; #region  resize as ref
/**
 * 根据基准形状设置选中对象的大小（纯形状数组版）
 * 基准形状：最上 → 最左 → 先选中
 * @param {Object} options 选项对象
 *   - match_width: 是否匹配宽度（默认true）
 *   - match_height: 是否匹配高度（默认true）
 *   - keep_ratio: 是否保持宽高比（默认false）
 *   - match_position: 是否匹配位置（默认false）
 * @return {Object} 操作结果
 */
resize_shapes_to_reference(options := "") {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection

        if (selection.Type != 2) {
            return { success: false, message: "请选中形状！" }
        }

        shape_range := selection.ShapeRange
        if (shape_range.Count < 2) {
            return { success: false, message: "请至少选中两个形状！" }
        }

        ; 纯形状数组
        shapes := []
        loop shape_range.Count {
            shapes.Push(shape_range(A_Index))
        }

        ; 选择基准形状（最上 → 最左 → 先选中）
        reference_shape := get_reference_shape(shapes)

        ; 设置选项
        opts := { match_width: true, match_height: true, keep_ratio: false, match_position: false }
        if (IsObject(options)) {
            for key, value in options.OwnProps() {
                opts.%key% := value
            }
        }

        ; 获取基准尺寸
        ref_width := reference_shape.Width
        ref_height := reference_shape.Height
        ref_left := reference_shape.Left
        ref_top := reference_shape.Top

        ; 调整其他形状
        adjusted_count := 0
        for shape in shapes {
            if (shape.Name = reference_shape.Name) {
                continue  ; 跳过基准形状
            }

            ; 计算新尺寸
            new_width := shape.Width
            new_height := shape.Height

            if (opts.match_width && opts.match_height) {
                if (opts.keep_ratio) {
                    ratio := Min(ref_width / shape.Width, ref_height / shape.Height)
                    new_width := shape.Width * ratio
                    new_height := shape.Height * ratio
                } else {
                    new_width := ref_width
                    new_height := ref_height
                }
            } else if (opts.match_width) {
                new_width := ref_width
                if (opts.keep_ratio) {
                    new_height := shape.Height * (ref_width / shape.Width)
                }
            } else if (opts.match_height) {
                new_height := ref_height
                if (opts.keep_ratio) {
                    new_width := shape.Width * (ref_height / shape.Height)
                }
            }

            ; 应用尺寸
            shape.Width := new_width
            shape.Height := new_height

            ; 应用位置
            if (opts.match_position) {
                shape.Left := ref_left
                shape.Top := ref_top
            }

            adjusted_count++
        }

        return {
            success: true,
            message: "调整完成",
            reference_shape: reference_shape.Name,
            adjusted_count: adjusted_count,
            total_shapes: shapes.Length
        }

    } catch as err {
        return { success: false, message: "操作失败: " err.Message }
    }
}

/**
 * 匹配宽度
 */
match_width() {
    return resize_shapes_to_reference({ match_width: true, match_height: false })
}

/**
 * 匹配高度
 */
match_height() {
    return resize_shapes_to_reference({ match_width: false, match_height: true })
}

/**
 * 匹配尺寸（保持比例）
 */
match_size_keep_ratio() {
    return resize_shapes_to_reference({ match_width: true, match_height: true, keep_ratio: true })
}

/**
 * 完全匹配
 */
match_completely() {
    return resize_shapes_to_reference({ match_width: true, match_height: true })
}

; #endregion

; #region  set_textbox_content_from_clipboard
/**
 * 为选中的文本框设置剪贴板文本内容
 * 文本分配顺序：最上 → 最左 → 先选中 或 按选中顺序
 * 剪贴板内容通过回车切分，按顺序分配给文本框
 * @param {Object} options 选项对象，支持:
 *   - clear_empty: 如果文本框数量多于文本段，是否清空多余文本框（默认true）
 *   - trim_text: 是否修剪文本前后的空白字符（默认true）
 *   - skip_empty_lines: 是否跳过空行（默认true）
 *   - sort_by_index: 是否按选中顺序排序（默认false，按位置排序）
 * @return {Object} 包含成功状态和操作结果信息
 */
set_textbox_content_from_clipboard(options := "") {
    start_time := A_TickCount

    ; 默认选项
    opts := {
        clear_empty: true,
        trim_text: true,
        skip_empty_lines: true,
        sort_by_index: false  ; 新增：是否按索引排序
    }

    ; 合并用户选项
    if (IsObject(options)) {
        for key, value in options.OwnProps() {
            opts.%key% := value
        }
    }

    try {
        ; 获取剪贴板内容
        try {
            clipboard_text := A_Clipboard
            if (clipboard_text = "") {
                return { success: false, message: "剪贴板为空！" }
            }
        } catch as err {
            return { success: false, message: "无法读取剪贴板内容: " err.Message }
        }

        ; 切分剪贴板文本（按回车）
        text_lines := StrSplit(clipboard_text, "`n", "`r")

        ; 处理文本行
        processed_lines := []
        for index, line in text_lines {
            if (opts.trim_text) {
                line := Trim(line)
            }

            if (opts.skip_empty_lines && line = "") {
                continue
            }

            processed_lines.Push(line)
        }

        if (processed_lines.Length = 0) {
            return { success: false, message: "剪贴板中没有有效的文本内容！" }
        }

        ; 获取PowerPoint应用对象
        ppt_app := ComObjActive("PowerPoint.Application")
        if (ppt_app.Presentations.Count = 0) {
            return { success: false, message: "没有打开的PowerPoint文档！" }
        }

        ; 获取选中的对象
        selection := ppt_app.ActiveWindow.Selection
        if (selection.Type != 2 && selection.Type != 3) {  ; ppSelectionShapes=2, ppSelectionText=3
            return { success: false, message: "请先选中至少一个文本框！" }
        }

        ; 获取选中的形状范围
        shape_range := (selection.HasChildShapeRange) ? selection.ChildShapeRange : selection.ShapeRange

        ; 筛选出文本框
        textboxes := []
        loop shape_range.Count {
            shape := shape_range(A_Index)

            ; 检查是否为文本框
            if (shape.HasTextFrame) {
                textboxes.Push({
                    shape: shape,
                    name: shape.Name,
                    top: shape.Top,
                    left: shape.Left,
                    index: A_Index
                })
            }
        }

        ; 检查是否找到文本框
        if (textboxes.Length = 0) {
            return { success: false, message: "选中的对象中没有文本框！" }
        }

        ; 对文本框排序
        sorted_textboxes := []
        if (opts.sort_by_index) {
            ; 按选中顺序排序（索引号）
            sorted_textboxes := sort_shapes_by_rules(textboxes, [{ property: "index", ascending: true }])
        } else {
            ; 按位置排序（最上 → 最左 → 先选中）
            sorted_textboxes := sort_shapes_by_rules(textboxes, [{ property: "Top", ascending: true }, { property: "Left",
                ascending: true }, { property: "index", ascending: true }
            ])
        }

        ; 分配文本内容
        updated_count := 0
        text_assignments := []

        for index, textbox in sorted_textboxes {
            text_content := ""

            if (index <= processed_lines.Length) {
                ; 分配对应的文本行
                text_content := processed_lines[index]
            } else if (opts.clear_empty) {
                ; 清空多余的文本框
                text_content := ""
            } else {
                ; 保留原有内容
                try {
                    text_content := textbox.shape.TextFrame.TextRange.Text
                } catch {
                    text_content := ""
                }
            }

            ; 设置文本内容
            try {
                textbox.shape.TextFrame.TextRange.Text := text_content
                updated_count++

                text_assignments.Push({
                    textbox_name: textbox.name,
                    textbox_index: textbox.index,
                    sort_index: index,
                    text_content: text_content,
                    assigned_line: (index <= processed_lines.Length) ? index : 0,
                    sort_method: opts.sort_by_index ? "按索引顺序" : "按位置顺序",
                    success: true
                })

            } catch as err {
                text_assignments.Push({
                    textbox_name: textbox.name,
                    textbox_index: textbox.index,
                    sort_index: index,
                    text_content: text_content,
                    assigned_line: (index <= processed_lines.Length) ? index : 0,
                    sort_method: opts.sort_by_index ? "按索引顺序" : "按位置顺序",
                    success: false,
                    error: err.Message
                })
            }
        }

        ; 返回结果
        total_duration := A_TickCount - start_time

        return {
            success: true,
            message: "文本内容设置完成",
            clipboard_line_count: processed_lines.Length,
            textbox_count: sorted_textboxes.Length,
            updated_count: updated_count,
            sort_method: opts.sort_by_index ? "按索引顺序" : "按位置顺序",
            text_assignments: text_assignments,
            options: opts,
            duration: total_duration
        }

    } catch as err {
        return { success: false, message: "设置文本内容失败: " err.Message }
    }
}

/**
 * 便捷函数：按选中顺序设置文本内容
 * @param {Object} options 其他选项
 * @return {Object} 包含成功状态和操作结果信息
 */
set_textbox_content_by_index(options := "") {
    base_options := { sort_by_index: true }

    ; 合并用户选项
    if (IsObject(options)) {
        for key, value in options.OwnProps() {
            base_options.%key% := value
        }
    }

    return set_textbox_content_from_clipboard(base_options)
}

/**
 * 便捷函数：按位置顺序设置文本内容
 * @param {Object} options 其他选项
 * @return {Object} 包含成功状态和操作结果信息
 */
set_textbox_content_by_position(options := "") {
    base_options := { sort_by_index: false }

    ; 合并用户选项
    if (IsObject(options)) {
        for key, value in options.OwnProps() {
            base_options.%key% := value
        }
    }

    return set_textbox_content_from_clipboard(base_options)
}

; #endregion

; #region  sort_shapes

range_to_array(range) {
    array := []
    if range.Count = 0 {
        return array
    }
    for element in range {
        element := range.Item(A_Index)
        array.Push(element)
    }
    return array
}

/**
 * 获取基准形状（最上 → 最左 → 先选中）
 */
get_reference_shape(shapes) {
    sorted_shapes := sort_shapes_by_rules(shapes, [{ property: "Top", ascending: true }, { property: "Left", ascending: true }, { property: "index",
        ascending: true }])
    return sorted_shapes[1]
}

/**
 * 根据规则对形状进行排序
 * @param {Array} source_shapes 源形状数组
 * @param {Array|String} rules 排序规则数组或单个属性字符串
 * @param {Boolean} ascending 当rules为字符串时的排序方向
 * @return {Array} 排序后的形状数组
 */
sort_shapes_by_rules(source_shapes := [], rules := [], ascending := true) {
    if (source_shapes.Length = 0) {
        return []
    }

    ; 处理字符串形式的简单排序
    if (Type(rules) = "String") {
        rules := [{ property: rules, ascending: ascending }]
    }

    ; 默认排序规则：最上 → 最左 → 先选中
    if (rules.Length = 0) {
        rules := [{ property: "Top", ascending: true }, { property: "Left", ascending: true }, { property: "index",
            ascending: true }]  ; 加上index确保完全确定顺序]

        sorted := source_shapes.Clone()

        ; 使用稳定的插入排序进行多条件排序
        loop sorted.Length {
            i := A_Index
            temp := sorted[i]
            j := i

            while (j > 1) {
                prev := sorted[j - 1]
                should_swap := false

                ; 按规则顺序比较
                for rule in rules {
                    prop := rule.property
                    rule_ascending := rule.ascending

                    if (rule_ascending) {
                        if (prev.%prop% > temp.%prop%) {
                            should_swap := true
                            break
                        } else if (prev.%prop% < temp.%prop%) {
                            should_swap := false
                            break
                        }
                        ; 相等则继续比较下一个条件
                    } else {
                        ; 降序
                        if (prev.%prop% < temp.%prop%) {
                            should_swap := true
                            break
                        } else if (prev.%prop% > temp.%prop%) {
                            should_swap := false
                            break
                        }
                        ; 相等则继续比较下一个条件
                    }
                }

                if (should_swap) {
                    sorted[j] := prev
                    j -= 1
                } else {
                    break
                }
            }

            sorted[j] := temp
        }

        return sorted
    }
}

; #endregion

; #region  insert_jiaonang
; insert_jiaonang(font_size := 12, text_color := 0xFFFFFF, fill_color := -5, char_count := 4) {
;     ; 参数验证
;     if (font_size < 10 || font_size > 36) {
;         return {success: false, message: "胶囊字体大小应在10-36之间"}
;     }

;     if (char_count < 2 || char_count > 8) {
;         return {success: false, message: "胶囊字数应在2-8之间"}
;     }

;     try {
;         ; 获取PowerPoint应用对象
;         ppt_app := ComObjActive("PowerPoint.Application")
;         window := ppt_app.ActiveWindow
;         if (ppt_app.Presentations.Count = 0) {
;             return {success: false, message: "没有打开的PowerPoint文档！"}
;         }

;         ; 获取当前幻灯片
;         current_slide := ppt_app.ActiveWindow.View.Slide

;         ; 确定插入位置
;         point_x := 0
;         point_y := 0
;         position_type := "mouse"  ; 位置类型

;         ; 检查是否有选中的对象
;         selection := ppt_app.ActiveWindow.Selection
;         if (selection.Type = 2 || selection.Type = 3) {  ; ppSelectionShapes=2, ppSelectionText=3
;             ; 有选中的对象，在对象下方4点位置插入
;             shape_range := (selection.HasChildShapeRange) ? selection.ChildShapeRange : selection.ShapeRange
;             if (shape_range.Count > 0) {
;                 ; 获取第一个选中对象的底部位置
;                 first_shape := shape_range(1)
;                 point_x := first_shape.Left
;                 point_y := first_shape.Top + first_shape.Height + 8  ; 对象下方4点
;                 position_type := "below_selection"
;             } else {
;                 ; 没有选中对象，使用鼠标位置
;                 MouseGetPos(&mouse_x, &mouse_y)
;                 point_x := ScreenPixelsToPointsX(mouse_x, window)
;                 point_y := ScreenPixelsToPointsY(mouse_y, window)
;                 position_type := "mouse"
;             }
;         } else {
;             ; 没有选中的对象，使用鼠标位置
;             MouseGetPos(&mouse_x, &mouse_y)
;             point_x := ScreenPixelsToPointsX(mouse_x, window)
;             point_y := ScreenPixelsToPointsY(mouse_y, window)
;             position_type := "mouse"
;         }

;         ; 计算胶囊尺寸（根据字数和字体大小）
;         char_width := font_size * 1.4  ; 中文字符宽度
;         text_width := char_width * char_count
;         text_height := font_size * 1.4

;         ; 胶囊尺寸（包含内边距）
;         ; capsule_width := text_width + font_size * 2  ; 左右各留一个字体大小的边距
;         capsule_width := text_width
;         capsule_height := text_height + font_size * 0.8  ; 上下各留0.4个字体大小的边距

;         ; 创建圆角矩形（5 = msoShapeRoundedRectangle）
;         shape := current_slide.Shapes.AddShape(5, point_x, point_y, capsule_width, capsule_height)

;         ; 设置圆角弧度（拉满）
;         shape.Adjustments.Item[1] := 1  ; 圆角弧度调整到最大

;         shape.Select

;         ; 设置填充颜色
;         if (fill_color < 0) {
;             ; 使用主题颜色
;             theme_color_index := Abs(fill_color)
;             shape.Fill.ForeColor.ObjectThemeColor := theme_color_index
;         } else {
;             ; 使用自定义RGB颜色
;             shape.Fill.ForeColor.RGB := fill_color
;         }
;         shape.Fill.Transparency := 0  ; 不透明

;         ; 设置边框（无边框）
;         shape.Line.Visible := 0

;         ; 添加文本
;         shape.TextFrame.TextRange.Text := generate_jiaonang_text(char_count)

;         ; 文本格式设置
;         shape.TextFrame.TextRange.Font.Name := "微软雅黑"
;         shape.TextFrame.TextRange.Font.Size := font_size
;         shape.TextFrame.TextRange.Font.Color.RGB := text_color

;         ; 文本对齐（水平垂直都居中）
;         shape.TextFrame.HorizontalAnchor := 3  ; msoAnchorCenter
;         shape.TextFrame.VerticalAnchor := 3    ; msoAnchorCenter

;         ; 文本边距设置
;         ; shape.TextFrame.MarginLeft := font_size * 0.5
;         ; shape.TextFrame.MarginRight := font_size * 0.5
;         shape.TextFrame.MarginTop := font_size * 0.2
;         shape.TextFrame.MarginBottom := font_size * 0.2

;         shape.TextFrame.MarginLeft := 0
;         shape.TextFrame.MarginRight := 0

;         ; 自动调整文本大小关闭
;         shape.TextFrame.AutoSize := 0

;         ; 设置形状名称
;         shape.Name := "jiaonang_" . A_TickCount

;         ; ; 选中新创建的胶囊
;         ; shape.Select

;         WinActivate("ahk_exe POWERPNT.EXE")
;         return {
;             success: true,
;             message: "成功插入胶囊形状",
;             text: shape.TextFrame.TextRange.Text,
;             position: {x: point_x, y: point_y},
;             size: {width: capsule_width, height: capsule_height},
;             font_size: font_size,
;             char_count: char_count,
;             position_type: position_type
;         }

;     } catch as err {
;         return {success: false, message: "插入胶囊失败: " . err.Message}
;     }
; }

; /**
;  * 在选中对象下方插入胶囊
;  */
; insert_jiaonang_below() {
;     return insert_jiaonang(16, 0xFFFFFF, -5, 4)
; }

; /**
;  * 在鼠标位置插入胶囊
;  */
; insert_jiaonang_at_cursor() {
;     ; 临时清除选择以确保使用鼠标位置
;     try {
;         ppt_app := ComObjActive("PowerPoint.Application")
;         ppt_app.ActiveWindow.Selection.Unselect()
;     } catch {
;         ; 忽略错误，继续执行
;     }
;     return insert_jiaonang(16, 0xFFFFFF, -5, 4)
; }

/**
 * 生成胶囊文本
 */
generate_jiaonang_text(char_count := 4) {
    ; 2字标签库
    tags_2char := ["开始", "体验", "试用", "购买", "查看", "下载", "更新", "设置", "登录", "注册"]

    ; 3字标签库
    tags_3char := ["立即开始", "免费体验", "点击购买", "查看详情", "下载安装", "设置配置", "用户登录", "新用户注册"]

    ; 4字标签库
    tags_4char := [
        "开始使用", "立即体验", "免费试用", "了解更多",
        "立即购买", "点击查看", "快速开始", "新手引导",
        "功能特性", "产品优势", "核心价值", "使用指南",
        "操作说明", "注意事项", "常见问题", "联系我们",
        "技术支持", "客户服务", "下载中心", "更新日志"
    ]

    ; 5字标签库
    tags_5char := ["立即开始使用", "免费体验版", "点击了解更多", "查看详细内容", "下载安装包", "设置个性化", "用户登录入口"]

    tags_map := Map(
        2, tags_2char,
        3, tags_3char,
        4, tags_4char,
        5, tags_5char
    )

    ; 如果指定字数不在映射中，使用4字
    available_tags := tags_map.Has(char_count) ? tags_map[char_count] : tags_4char
    random_index := Random(1, available_tags.Length)
    return available_tags[random_index]
}

; ==================== 便捷函数 ====================

; /**
;  * 插入蓝色胶囊（默认）
;  */
; insert_blue_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0x0070C0, 4)
; }

; /**
;  * 插入绿色胶囊
;  */
; insert_green_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0x00B050, 4)
; }

; /**
;  * 插入橙色胶囊
;  */
; insert_orange_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0xFF6600, 4)
; }

; /**
;  * 插入红色胶囊
;  */
; insert_red_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0xFF0000, 4)
; }

; /**
;  * 插入紫色胶囊
;  */
; insert_purple_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0x7030A0, 4)
; }

; /**
;  * 插入大胶囊
;  */
; insert_large_jiaonang() {
;     return insert_jiaonang(20, 0xFFFFFF, 0x0070C0, 4)
; }

; /**
;  * 插入小胶囊
;  */
; insert_small_jiaonang() {
;     return insert_jiaonang(14, 0xFFFFFF, 0x0070C0, 4)
; }

; /**
;  * 插入2字胶囊
;  */
; insert_2char_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0x0070C0, 2)
; }

; /**
;  * 插入3字胶囊
;  */
; insert_3char_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0x0070C0, 3)
; }

; /**
;  * 插入5字胶囊
;  */
; insert_5char_jiaonang() {
;     return insert_jiaonang(16, 0xFFFFFF, 0x0070C0, 5)
; }

; #endregion

; #region  ModifySlideNameShapes
ModifySlideNameShapes() {
    try {
        ; 获取 PowerPoint 应用程序对象
        ppt := ComObjActive("PowerPoint.Application")

        ; 检查是否在 PowerPoint 环境中
        if !ppt || !ppt.ActivePresentation {
            MsgBox "未检测到活动的 PowerPoint 演示文稿。", "错误", 16
            return
        }

        presentation := ppt.ActivePresentation
        selection := ppt.ActiveWindow.Selection

        ; 检查选中的是否是幻灯片
        if (selection.Type != 1) { ; ppSelectionSlides = 1
            MsgBox "请先选中一个或多个幻灯片。", "提示", 64
            return
        }

        slideRange := selection.SlideRange
        slidesCount := slideRange.Count

        if (slidesCount = 0) {
            MsgBox "没有选中任何幻灯片。", "信息", 64
            return
        }

        processedCount := 0
        errors := []

        ; 遍历所有选中的幻灯片
        loop slidesCount {
            slideIndex := A_Index
            try {
                slide := slideRange.Item(slideIndex)
                originalSlideNumber := slide.SlideNumber

                ; 查找名为 "slide_name" 的形状
                sourceShape := false
                loop slide.Shapes.Count {
                    shape := slide.Shapes.Item(A_Index)
                    if (shape.Name = "slide_name") {
                        sourceShape := shape
                        break
                    }
                }

                if !sourceShape {
                    errors.Push("幻灯片 " originalSlideNumber " 中未找到名为 'slide_name' 的形状")
                    continue
                }

                ; 步骤1: 重命名原始形状并设置为蓝色
                try {
                    sourceShape.Name := "set_slide_name"
                    ; 设置字体颜色为蓝色 (RGB: 0, 0, 255)
                    sourceShape.TextFrame.TextRange.Font.Color := RGB(0, 0, 255)
                    ; 或者设置形状填充色为蓝色（根据您的需求选择）
                    ; sourceShape.Fill.ForeColor.RGB := RGB(0, 0, 255)
                } catch as err {
                    errors.Push("幻灯片 " originalSlideNumber " 设置原始形状属性失败: " err.message)
                }

                ; 步骤2: 复制形状
                try {
                    ; 复制原始形状
                    sourceShape.Copy()

                    ; 粘贴到同一幻灯片
                    newShape := slide.Shapes.Paste()

                    ; 重命名新形状
                    newShape.Name := "get_slide_name"

                    ; 设置新形状位置 (左对齐，下移8点)
                    newShape.Left := sourceShape.Left
                    newShape.Top := sourceShape.Top + sourceShape.Height + 8  ; 下移8个点

                    ; 设置新形状字体颜色为灰色 (RGB: 128, 128, 128)
                    newShape.TextFrame.TextRange.Font.Color := RGB(128, 128, 128)
                    ; 或者设置形状填充色为灰色（根据您的需求选择）
                    ; newShape.Fill.ForeColor.RGB := RGB(128, 128, 128)

                    processedCount++

                } catch as err {
                    errors.Push("幻灯片 " originalSlideNumber " 复制形状失败: " err.message)
                }

            } catch as err {
                errors.Push("处理幻灯片 " originalSlideNumber " 时出错: " err.message)
            }
        }

        ; 显示操作结果
        resultMessage := "操作完成！`n`n成功处理: " processedCount " 个幻灯片"
        if (errors.Length > 0) {
            resultMessage .= "`n`n遇到 " errors.Length " 个错误:"
            loop Min(errors.Length, 5) {  ; 最多显示5个错误
                resultMessage .= "`n• " errors[A_Index]
            }
            if (errors.Length > 5) {
                resultMessage .= "`n• ... 还有 " (errors.Length - 5) " 个错误"
            }
        }

        MsgBox resultMessage, "处理结果", processedCount > 0 ? 64 : 48

    } catch as err {
        MsgBox "发生意外错误: " err.message, "严重错误", 16
    }
}

; RGB 颜色转换函数
RGB(red, green, blue) {
    return (blue << 16) | (green << 8) | red
}

; 设置热键：Ctrl+Alt+M 执行操作
; ^!m::ModifySlideNameShapes()

; 功能：将选中幻灯片的名称赋值给 get_slide_name 形状的文本

AssignSlideNameToShape() {
    try {
        ; 获取 PowerPoint 应用程序对象
        ppt := ComObjActive("PowerPoint.Application")

        ; 检查是否在 PowerPoint 环境中
        if !ppt || !ppt.ActivePresentation {
            MsgBox "未检测到活动的 PowerPoint 演示文稿。", "错误", 16
            return
        }

        presentation := ppt.ActivePresentation
        selection := ppt.ActiveWindow.Selection

        ; 检查选中的是否是幻灯片
        if (selection.Type != 1) { ; ppSelectionSlides = 1
            MsgBox "请先选中一个或多个幻灯片。", "提示", 64
            return
        }

        slideRange := selection.SlideRange
        slidesCount := slideRange.Count

        if (slidesCount = 0) {
            MsgBox "没有选中任何幻灯片。", "信息", 64
            return
        }

        processedCount := 0
        errors := []

        ; 遍历所有选中的幻灯片
        loop slidesCount {
            slideIndex := A_Index
            try {
                slide := slideRange.Item(slideIndex)
                originalSlideNumber := slide.SlideNumber
                slideName := slide.Name  ; 获取幻灯片的名称

                ; 如果幻灯片名称为空，使用默认名称
                if (Trim(slideName) = "") {
                    slideName := "Slide " originalSlideNumber
                }

                ; 查找名为 "get_slide_name" 的形状
                targetShape := false
                loop slide.Shapes.Count {
                    shape := slide.Shapes.Item(A_Index)
                    if (shape.Name = "get_slide_name") {
                        targetShape := shape
                        break
                    }
                }

                if !targetShape {
                    errors.Push("幻灯片 " originalSlideNumber " 中未找到名为 'get_slide_name' 的形状")
                    continue
                }

                ; 将幻灯片名称赋值给形状的文本
                try {
                    targetShape.TextFrame.TextRange.Text := slideName
                    processedCount++

                    ; 可选：在控制台显示操作信息
                    ; ToolTip "幻灯片 " originalSlideNumber " : " slideName
                    ; Sleep(300)

                } catch as err {
                    errors.Push("幻灯片 " originalSlideNumber " 设置文本失败: " err.message)
                }

            } catch as err {
                errors.Push("处理幻灯片 " originalSlideNumber " 时出错: " err.message)
            }
        }

        ; ToolTip  ; 清除提示（如果使用了ToolTip）

        ; 显示操作结果
        resultMessage := "操作完成！`n`n成功处理: " processedCount " 个幻灯片"
        if (errors.Length > 0) {
            resultMessage .= "`n`n遇到 " errors.Length " 个错误:"
            loop Min(errors.Length, 5) {  ; 最多显示5个错误
                resultMessage .= "`n• " errors[A_Index]
            }
            if (errors.Length > 5) {
                resultMessage .= "`n• ... 还有 " (errors.Length - 5) " 个错误"
            }
        }

        MsgBox resultMessage, "处理结果", processedCount > 0 ? 64 : 48

    } catch as err {
        MsgBox "发生意外错误: " err.message, "严重错误", 16
    }
}

; 设置热键：Ctrl+Alt+G 执行操作
; ^!g::AssignSlideNameToShape()

; 备用热键：Ctrl+Shift+G
; ^+g::AssignSlideNameToShape()

; 功能：根据选中幻灯片中指定形状的文本重新命名幻灯片
; 要求：形状的 Name 必须为 "slide_name"

RenameSelectedSlidesByShape() {
    try {
        ; 获取 PowerPoint 应用程序对象
        ppt := ComObjActive("PowerPoint.Application")

        ; 检查是否在 PowerPoint 环境中
        if !ppt || !ppt.ActivePresentation {
            MsgBox "未检测到活动的 PowerPoint 演示文稿。", "错误", 16
            return
        }

        presentation := ppt.ActivePresentation
        selection := ppt.ActiveWindow.Selection

        ; 检查选中的是否是幻灯片
        if (selection.Type != 1) { ; ppSelectionSlides = 1
            MsgBox "请先选中一个或多个幻灯片。", "提示", 64
            return
        }

        slideRange := selection.SlideRange
        slidesCount := slideRange.Count

        if (slidesCount = 0) {
            MsgBox "没有选中任何幻灯片。", "信息", 64
            return
        }

        renamedCount := 0
        errors := []

        ; 遍历所有选中的幻灯片
        loop slidesCount {
            slideIndex := A_Index
            try {
                slide := slideRange.Item(slideIndex)
                originalSlideNumber := slide.SlideNumber  ; 记录原始幻灯片编号用于显示

                ; 遍历幻灯片中的所有形状
                found := false
                loop slide.Shapes.Count {
                    shape := slide.Shapes.Item(A_Index)

                    ; 检查形状名称是否为 "slide_name"
                    if (shape.Name = "set_slide_name") {
                        ; 获取形状文本并清理
                        newName := Trim(shape.TextFrame.TextRange.Text)

                        ; 检查文本是否有效
                        if (newName != "") {
                            slide.Name := newName
                            renamedCount++
                            found := true
                            ; 输出成功信息到控制台
                            ToolTip "已重命名幻灯片 " originalSlideNumber " 为: " newName
                            Sleep(500)  ; 短暂显示提示
                            break  ; 找到后跳出内层循环
                        } else {
                            errors.Push("幻灯片 " originalSlideNumber " 中的形状文本为空")
                        }
                    }
                }

                if (!found) {
                    errors.Push("幻灯片 " originalSlideNumber " 中未找到名为 'slide_name' 的形状")
                }

            } catch as err {
                errors.Push("处理幻灯片 " originalSlideNumber " 时出错: " err.message)
            }
        }

        ToolTip  ; 清除提示

        ; 显示操作结果
        resultMessage := "操作完成！`n`n成功重命名: " renamedCount " 个幻灯片"
        if (errors.Length > 0) {
            resultMessage .= "`n`n遇到 " errors.Length " 个错误:"
            loop Min(errors.Length, 10) {  ; 最多显示10个错误
                resultMessage .= "`n• " errors[A_Index]
            }
            if (errors.Length > 10) {
                resultMessage .= "`n• ... 还有 " (errors.Length - 10) " 个错误"
            }
        }

        MsgBox resultMessage, "重命名结果", renamedCount > 0 ? 64 : 48

    } catch as err {
        MsgBox "发生意外错误: " err.message, "严重错误", 16
    }
}

; 设置热键：Ctrl+Alt+R 执行重命名操作
; ^!r::RenameSelectedSlidesByShape()

; 备用热键：Ctrl+Shift+S
; ^+s::RenameSelectedSlidesByShape()

; 功能：锁定选中幻灯片中的所有形状

LockAllShapesInSelectedSlides() {
    try {
        ; 获取 PowerPoint 应用程序对象
        ppt := ComObjActive("PowerPoint.Application")

        ; 检查是否在 PowerPoint 环境中
        if !ppt || !ppt.ActivePresentation {
            MsgBox "未检测到活动的 PowerPoint 演示文稿。", "错误", 16
            return
        }

        presentation := ppt.ActivePresentation
        selection := ppt.ActiveWindow.Selection

        ; 检查选中的是否是幻灯片
        if (selection.Type != 1) { ; ppSelectionSlides = 1
            MsgBox "请先选中一个或多个幻灯片。", "提示", 64
            return
        }

        slideRange := selection.SlideRange
        slidesCount := slideRange.Count

        if (slidesCount = 0) {
            MsgBox "没有选中任何幻灯片。", "信息", 64
            return
        }

        lockedShapesCount := 0
        processedSlides := 0
        errors := []

        ; 遍历所有选中的幻灯片
        loop slidesCount {
            slideIndex := A_Index
            try {
                slide := slideRange.Item(slideIndex)
                originalSlideNumber := slide.SlideNumber
                slideShapesCount := slide.Shapes.Count

                if (slideShapesCount = 0) {
                    errors.Push("幻灯片 " originalSlideNumber " 中没有形状")
                    continue
                }

                shapesLockedInSlide := 0

                ; 遍历幻灯片中的所有形状并锁定
                loop slideShapesCount {
                    shape := slide.Shapes.Item(A_Index)

                    try {
                        ; 锁定形状（设置 Locked 属性为 True）
                        shape.Locked := -1
                        shapesLockedInSlide++
                        lockedShapesCount++
                    } catch as err {
                        errors.Push("幻灯片 " originalSlideNumber " 形状 " A_Index " 锁定失败: " err.message)
                    }
                }

                if (shapesLockedInSlide > 0) {
                    processedSlides++
                    ; ToolTip "幻灯片 " originalSlideNumber " 已锁定 " shapesLockedInSlide " 个形状"
                    ; Sleep(300)
                }

            } catch as err {
                errors.Push("处理幻灯片 " originalSlideNumber " 时出错: " err.message)
            }
        }

        ToolTip  ; 清除提示

        ; 显示操作结果
        resultMessage := "操作完成！`n`n"
        resultMessage .= "已处理幻灯片: " processedSlides " 个`n"
        resultMessage .= "已锁定形状: " lockedShapesCount " 个"

        if (errors.Length > 0) {
            resultMessage .= "`n`n遇到 " errors.Length " 个错误:"
            loop Min(errors.Length, 5) {  ; 最多显示5个错误
                resultMessage .= "`n• " errors[A_Index]
            }
            if (errors.Length > 5) {
                resultMessage .= "`n• ... 还有 " (errors.Length - 5) " 个错误"
            }
        }

        MsgBox resultMessage, "锁定形状结果", lockedShapesCount > 0 ? 64 : 48

    } catch as err {
        MsgBox "发生意外错误: " err.message, "严重错误", 16
    }
}

; 设置热键：Ctrl+Alt+L 锁定形状
; ^!l::LockAllShapesInSelectedSlides()

; 功能：从 sample_slide 复制 header_1 形状到所有名称包含 "body" 的幻灯片，并在文本后添加幻灯片名称

CopyHeaderToBodySlidesWithName() {
    try {
        ; 获取 PowerPoint 应用程序对象
        ppt := ComObjActive("PowerPoint.Application")

        ; 检查是否在 PowerPoint 环境中
        if !ppt || !ppt.ActivePresentation {
            MsgBox "未检测到活动的 PowerPoint 演示文稿。", "错误", 16
            return
        }

        presentation := ppt.ActivePresentation
        slidesCount := presentation.Slides.Count

        if (slidesCount = 0) {
            MsgBox "演示文稿中没有幻灯片。", "信息", 64
            return
        }

        ; 1. 查找 sample_slide 和 header_1 形状
        sampleSlide := false
        headerShape := false

        loop slidesCount {
            slide := presentation.Slides.Item(A_Index)
            if (slide.Name = "sample_1") {
                sampleSlide := slide

                ; 在 sample_slide 中查找 header_1 形状
                loop slide.Shapes.Count {
                    shape := slide.Shapes.Item(A_Index)
                    if (shape.Name = "header_1") {
                        headerShape := shape
                        break
                    }
                }
                break
            }
        }

        if !sampleSlide {
            MsgBox "未找到名为 'sample_slide' 的幻灯片。", "错误", 16
            return
        }

        if !headerShape {
            MsgBox "在 'sample_slide' 中未找到名为 'header_1' 的形状。", "错误", 16
            return
        }

        ; 获取原始形状的文本
        originalText := headerShape.TextFrame.TextRange.Text

        ; 2. 查找所有名称包含 "body" 的幻灯片
        bodySlides := []
        loop slidesCount {
            slide := presentation.Slides.Item(A_Index)
            if (InStr(slide.Name, "body")) {
                bodySlides.Push(slide)
            }
        }

        if (bodySlides.Length = 0) {
            MsgBox "未找到任何名称包含 'body' 的幻灯片。", "信息", 64
            return
        }

        ; 3. 复制并处理每个目标幻灯片
        processedCount := 0
        errors := []

        loop bodySlides.Length {
            targetSlide := bodySlides[A_Index]
            try {
                ; 复制原始形状
                headerShape.Copy()

                ; 粘贴到目标幻灯片
                newShape := targetSlide.Shapes.Paste()

                ; 在原始文本后添加幻灯片名称
                newText := originalText . "_" . targetSlide.Name
                newShape.TextFrame.TextRange.Text := newText

                processedCount++

                ; 可选：显示进度
                ; ToolTip "正在处理幻灯片: " targetSlide.SlideNumber " - " targetSlide.Name
                ; Sleep(200)

            } catch as err {
                errors.Push("幻灯片 " targetSlide.SlideNumber " 处理失败: " err.message)
            }
        }

        ToolTip  ; 清除提示

        ; 显示操作结果
        resultMessage := "操作完成！`n`n"
        resultMessage .= "源幻灯片: sample_slide (幻灯片 " sampleSlide.SlideNumber ")`n"
        resultMessage .= "复制形状: header_1`n"
        resultMessage .= "原始文本: " originalText "`n"
        resultMessage .= "目标幻灯片: " processedCount " 个名称包含 'body' 的幻灯片"

        if (errors.Length > 0) {
            resultMessage .= "`n`n遇到 " errors.Length " 个错误:"
            loop Min(errors.Length, 5) {
                resultMessage .= "`n• " errors[A_Index]
            }
        }

        MsgBox resultMessage, "复制结果", processedCount > 0 ? 64 : 48

    } catch as err {
        MsgBox "发生意外错误: " err.message, "严重错误", 16
    }
}

; 设置热键：Ctrl+Alt+H 执行操作
; ^!h::CopyHeaderToBodySlidesWithName()

; 功能：将选中形状的格式复制给所有同名的其他形状

CopyFormatToSameNameShapes() {
    try {
        ; 获取 PowerPoint 应用程序对象
        ppt := ComObjActive("PowerPoint.Application")

        ; 检查是否在 PowerPoint 环境中
        if !ppt || !ppt.ActivePresentation {
            MsgBox "未检测到活动的 PowerPoint 演示文稿。", "错误", 16
            return
        }

        presentation := ppt.ActivePresentation
        selection := ppt.ActiveWindow.Selection

        ; 检查选中的是否是形状
        if (selection.Type != 2) { ; ppSelectionShapes = 2
            MsgBox "请先选中一个形状。", "提示", 64
            return
        }

        if (selection.ShapeRange.Count != 1) {
            MsgBox "请只选中一个形状。", "提示", 64
            return
        }

        ; 获取源形状
        sourceShape := selection.ShapeRange(1)
        sourceShapeName := sourceShape.Name
        sourceSlideNumber := sourceShape.Parent.SlideNumber

        if (sourceShapeName = "") {
            MsgBox "选中的形状没有名称，请先为形状命名。", "错误", 16
            return
        }

        ; 在整个演示文稿中查找同名的形状
        targetShapes := []
        slidesCount := presentation.Slides.Count

        loop slidesCount {
            slide := presentation.Slides.Item(A_Index)
            loop slide.Shapes.Count {
                shape := slide.Shapes.Item(A_Index)
                if (shape.Name = sourceShapeName && shape.Parent.SlideNumber != sourceSlideNumber) {
                    targetShapes.Push(shape)
                }
            }
        }

        if (targetShapes.Length = 0) {
            MsgBox "未找到其他同名的形状。形状名称: " sourceShapeName, "信息", 64
            return
        }

        ; 复制源形状格式并应用到目标形状
        processedCount := 0
        errors := []

        loop targetShapes.Length {
            targetShape := targetShapes[A_Index]
            try {
                ; 复制格式（使用 PickUp 和 Apply 方法）
                sourceShape.PickUp()  ; 复制源形状的格式
                targetShape.Apply()   ; 应用格式到目标形状

                processedCount++

                ; 可选：显示进度
                ; ToolTip "正在处理幻灯片: " targetShape.Parent.SlideNumber
                ; Sleep(100)

            } catch as err {
                errors.Push("幻灯片 " targetShape.Parent.SlideNumber " 形状应用格式失败: " err.message)
            }
        }

        ToolTip  ; 清除提示

        ; 显示操作结果
        resultMessage := "格式复制完成！`n`n"
        resultMessage .= "源形状: " sourceShapeName " (幻灯片 " sourceSlideNumber ")`n"
        resultMessage .= "成功应用格式: " processedCount " 个同名形状"

        if (errors.Length > 0) {
            resultMessage .= "`n`n遇到 " errors.Length " 个错误:"
            loop Min(errors.Length, 5) {
                resultMessage .= "`n• " errors[A_Index]
            }
        }

        MsgBox resultMessage, "格式复制结果", processedCount > 0 ? 64 : 48

    } catch as err {
        MsgBox "发生意外错误: " err.message, "严重错误", 16
    }
}

; 设置热键：Ctrl+Alt+F 执行格式复制
; ^!f::CopyFormatToSameNameShapes()

; 功能：将选中的形状复制到所有名称包含 "body" 的幻灯片

copy_selected_shapes_to_body_slides() {
    try {
        ; 获取 PowerPoint 应用程序对象
        ppt := ComObjActive("PowerPoint.Application")

        ; 检查是否在 PowerPoint 环境中
        if !ppt || !ppt.ActivePresentation {
            MsgBox "未检测到活动的 PowerPoint 演示文稿。", "错误", 16
            return
        }

        presentation := ppt.ActivePresentation
        selection := ppt.ActiveWindow.Selection

        ; 检查选中的是否是形状
        if (selection.Type != 2) { ; ppSelectionShapes = 2
            MsgBox "请先选中一个或多个形状。", "提示", 64
            return
        }

        shape_range := selection.ShapeRange
        selected_count := shape_range.Count

        if (selected_count = 0) {
            MsgBox "没有选中任何形状。", "信息", 64
            return
        }

        ; 查找所有名称包含 "body" 的幻灯片
        body_slides := []
        slides_count := presentation.Slides.Count

        loop slides_count {
            slide := presentation.Slides.Item(A_Index)
            if (InStr(slide.Name, "body")) {
                body_slides.Push(slide)
            }
        }

        if (body_slides.Length = 0) {
            MsgBox "未找到任何名称包含 'body' 的幻灯片。", "信息", 64
            return
        }

        ; 复制选中的形状到所有 body 幻灯片
        processed_slides := 0
        total_copies := 0
        errors := []

        ; 遍历所有 body 幻灯片
        loop body_slides.Length {
            target_slide := body_slides[A_Index]
            slide_copies := 0

            try {
                ; 遍历所有选中的形状
                loop selected_count {
                    shape := shape_range.Item(A_Index)

                    ; 复制形状
                    shape.Copy()

                    ; 粘贴到目标幻灯片
                    target_slide.Shapes.Paste()
                    slide_copies++
                    total_copies++
                }

                processed_slides++

            } catch as err {
                errors.Push("幻灯片 " target_slide.SlideNumber " 复制失败: " err.message)
            }
        }

        ; 显示操作结果
        result_message := "操作完成！`n`n"
        result_message .= "选中形状数量: " selected_count " 个`n"
        result_message .= "目标幻灯片: " processed_slides " 个名称包含 'body' 的幻灯片`n"
        result_message .= "总复制次数: " total_copies " 次"

        if (errors.Length > 0) {
            result_message .= "`n`n遇到 " errors.Length " 个错误:"
            loop Min(errors.Length, 5) {
                result_message .= "`n• " errors[A_Index]
            }
        }

        MsgBox result_message, "复制结果", total_copies > 0 ? 64 : 48

    } catch as err {
        MsgBox "发生意外错误: " err.message, "严重错误", 16
    }
}

; 设置热键：Ctrl+Alt+B 执行操作
; ^!b::copy_selected_shapes_to_body_slides()

; #endregion

; 获取 PPT 应用程序对象
; 获取 PPT 应用程序对象
get_ppt() {
    try {
        ppt := ComObjActive("PowerPoint.Application")
        return ppt
    } catch as err {
        Notify.show("错误：请先打开 PowerPoint 应用程序")
        return ""
    }
}

; 检查是否有活动演示文稿
has_active_presentation(ppt) {
    if (ppt.Presentations.Count = 0) {
        Notify.show("错误：没有打开的演示文稿")
        return false
    }
    return true
}

; 获取活动演示文稿
get_active_presentation() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return ""
    return ppt.ActivePresentation
}

; 在当前位置添加节
add_section() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        pres := ppt.ActivePresentation
        current_slide := ppt.ActiveWindow.View.Slide
        section_name := "新节_" . A_Now

        ; 在当前位置添加节
        pres.SectionProperties.AddBeforeSlide(current_slide.SlideIndex, section_name)

        Notify.show("已添加节: " . section_name)
    } catch as err {
        A_Clipboard := err.message
        Notify.show("错误：添加节失败 - " . err.message)
    }

    ppt := ""
}

; 重命名当前节
rename_current_section() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        pres := ppt.ActivePresentation
        current_slide := ppt.ActiveWindow.View.Slide
        section_props := pres.SectionProperties

        ; 获取当前节的索引
        section_index := section_props.GetSectionIndex(current_slide.SlideIndex)
        current_name := section_props.GetSectionName(section_index)

        ; 输入新名称
        new_name := InputBox("请输入新的节名称：", "重命名节", , current_name)

        if (new_name.Result = "OK" && new_name.Value != "") {
            section_props.RenameSection(section_index, new_name.Value)
            Notify.show("节已重命名为: " . new_name.Value)
        } else {
            Notify.show("操作已取消")
        }
    } catch as err {
        Notify.show("错误：重命名节失败 - " . err.message)
    }

    ppt := ""
}

; 删除当前节（保留幻灯片）
delete_current_section() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        pres := ppt.ActivePresentation
        current_slide := ppt.ActiveWindow.View.Slide
        section_props := pres.SectionProperties

        ; 获取当前节的索引
        section_index := section_props.GetSectionIndex(current_slide.SlideIndex)
        section_name := section_props.GetSectionName(section_index)

        result := MsgBox("确定要删除节 '" . section_name . "' 吗？`n（幻灯片将保留）", "确认删除", "YesNo")
        if (result = "Yes") {
            section_props.DeleteSection(section_index)
            Notify.show("节 '" . section_name . "' 已删除")
        } else {
            Notify.show("删除操作已取消")
        }
    } catch as err {
        Notify.show("错误：删除节失败 - " . err.message)
    }

    ppt := ""
}

; 折叠所有节
collapse_all_sections() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        ; 通过切换视图来刷新节状态
        current_view := ppt.ActiveWindow.View.Type
        ppt.ActiveWindow.View.Type := 1 ; ppViewNormal
        ppt.ActiveWindow.View.Type := current_view

        Notify.show("所有节已折叠")
    } catch as err {
        Notify.show("错误：折叠节失败 - " . err.message)
    }

    ppt := ""
}

; 展开所有节
expand_all_sections() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        ; 切换到节视图再切回来
        ppt.ActiveWindow.View.Type := 13 ; ppViewSlideSorter
        Sleep(100)
        ppt.ActiveWindow.View.Type := 1 ; ppViewNormal

        Notify.show("所有节已展开")
    } catch as err {
        Notify.show("错误：展开节失败 - " . err.message)
    }

    ppt := ""
}

; 获取所有节的信息
list_all_sections() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        pres := ppt.ActivePresentation
        section_props := pres.SectionProperties
        section_count := section_props.Count

        section_list := "当前演示文稿中的节：`n`n"

        loop section_count {
            current_index := A_Index

            name := section_props.Name(current_index)

            FirstSlide := section_props.FirstSlide(current_index)

            slide_count := section_props.slide_count(current_index)
            section_list .= current_index . ". " . name . " (FirstSlide: " . FirstSlide . ")`n"
        }

        MsgBox(section_list)
    } catch as err {
        A_Clipboard := err.message
        Notify.show("错误：获取节列表失败 - " . err.message)
    }

    ppt := ""
}

find_section_slides(section_name_to_find := "原稿") {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        pres := ppt.ActivePresentation
        section_props := pres.SectionProperties
        section_count := section_props.Count

        loop section_count {
            current_index := A_Index

            ; 尝试使用Name方法获取节名称
            name := section_props.Name(current_index)

            ; if (name = section_name_to_find) {
            ; 获取该节的第一个幻灯片编号
            first_slide := section_props.FirstSlide(section_name_to_find)
            ; 获取该节的幻灯片数量
            slide_count := section_props.SlidesCount(section_name_to_find)  ; 注意：这个方法名可能需要调整，根据实际情况可能是GetSectionSlidesCount

            ; 计算该节的幻灯片范围
            start_slide := first_slide
            end_slide := first_slide + slide_count - 1

            Notify.show("节 '" . section_name_to_find . "' 包含的幻灯片范围: 第 " . start_slide . " 页到第 " . end_slide . " 页")
            return
            ; }
        }

        Notify.show("未找到节: " . section_name_to_find)
    } catch as err {
        Notify.show("错误：查找节失败 - " . err.message)
    }

    ppt := ""
}

; 移动到指定节
go_to_section(section_index) {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        pres := ppt.ActivePresentation
        section_props := pres.SectionProperties

        if (section_index > section_props.Count || section_index < 1) {
            Notify.show("错误：节索引超出范围")
            return
        }

        ; 获取该节的第一个幻灯片
        slide_start := 1
        if (section_index > 1) {
            loop section_index - 1 {
                slide_start += section_props.GetSectionSlidesCount(A_Index)
            }
        }

        ; 跳转到该节的第一个幻灯片
        ppt.ActiveWindow.View.GotoSlide(slide_start)

        section_name := section_props.GetSectionName(section_index)
        Notify.show("已跳转到节: " . section_name)
    } catch as err {
        Notify.show("错误：跳转到节失败 - " . err.message)
    }

    ppt := ""
}

; 获取当前节信息
get_current_section_info() {
    ppt := get_ppt()
    if (!ppt || !has_active_presentation(ppt))
        return

    try {
        pres := ppt.ActivePresentation
        current_slide := ppt.ActiveWindow.View.Slide
        section_props := pres.SectionProperties

        section_index := section_props.GetSectionIndex(current_slide.SlideIndex)
        section_name := section_props.GetSectionName(section_index)
        slide_count := section_props.GetSectionSlidesCount(section_index)

        info_msg := "当前节信息：`n"
        info_msg .= "名称: " . section_name . "`n"
        info_msg .= "索引: " . section_index . "`n"
        info_msg .= "幻灯片数量: " . slide_count

        Notify.show(info_msg)
    } catch as err {
        Notify.show("错误：获取节信息失败 - " . err.message)
    }

    ppt := ""
}

; 热键配置 - 仅在PPT激活时生效
; #HotIf WinActive("ahk_class PPTFrameClass")

; ; 节管理热键
; F1:: add_section()                      ; 添加节
; F2:: rename_current_section()           ; 重命名节
; F3:: delete_current_section()           ; 删除节
; F4:: collapse_all_sections()            ; 折叠所有节
; F5:: expand_all_sections()              ; 展开所有节
; F6:: list_all_sections()                ; 列出所有节

; 快速跳转到节
; ^1:: go_to_section(1)                   ; Ctrl+1 跳转到第1节
; ^2:: go_to_section(2)                   ; Ctrl+2 跳转到第2节
; ^3:: go_to_section(3)                   ; Ctrl+3 跳转到第3节
; ^4:: go_to_section(4)                   ; Ctrl+4 跳转到第4节
; ^5:: go_to_section(5)                   ; Ctrl+5 跳转到第5节

; #HotIf

; #region  draw_shapes_from_patato_json

; 根据 patato.json 数据在 PPT 中绘制形状
draw_shapes_from_patato_json() {
    try {
        ; 读取保存的数据
        objects_data := read_patato_json()

        ; 获取当前活动的 PowerPoint 应用程序
        ppt_app := ComObjActive("PowerPoint.Application")
        if !ppt_app || !ppt_app.ActivePresentation {
            throw Error("未找到打开的 PowerPoint 演示文稿")
        }

        ; 获取当前活动的幻灯片
        active_slide := ppt_app.ActiveWindow.View.Slide
        if !active_slide {
            throw Error("未找到活动幻灯片")
        }

        ; 清空当前幻灯片的所有形状
        clear_slide_shapes(active_slide)

        ; 计数器
        chip_count := 1
        fry_count := 1

        ; 按数组顺序绘制（保持层级关系）
        shapes_created := 0
        for obj in objects_data["objects"] {
            if draw_shape_from_data(active_slide, obj, &chip_count, &fry_count) {
                shapes_created++
            }
        }

        MsgBox("成功绘制 " shapes_created " 个形状")
        return shapes_created

    } catch as err {
        MsgBox("绘制形状失败: " err.message)
        return 0
    }
}

; 清空幻灯片中的所有形状
clear_slide_shapes(slide) {
    try {
        ; 从后往前删除，避免索引变化问题
        while slide.Shapes.Count > 0 {
            slide.Shapes(slide.Shapes.Count).Delete()
        }
        return true
    } catch as err {
        OutputDebug("清空形状失败: " err.message)
        return false
    }
}

; 根据单个对象数据绘制形状
draw_shape_from_data(slide, obj_data, &chip_count, &fry_count) {
    try {
        left := obj_data["left"]
        top := obj_data["top"]
        width := obj_data["width"]
        height := obj_data["height"]

        ; 获取倾斜角度，如果不存在则默认为0
        rotation := obj_data.Has("rotation") ? obj_data["rotation"] : 0

        ; 使用固定 0.1 阈值分类
        if height <= 0.1 {
            ; 水平线
            shape := slide.Shapes.AddLine(left, top, left + width, top)
            shape.Name := "fry_" fry_count
            apply_fry_style(shape)

            ; 应用旋转角度
            if rotation != 0 {
                shape.Rotation := rotation
            }

            fry_count++
        } else if width <= 0.1 {
            ; 垂直线
            shape := slide.Shapes.AddLine(left, top, left, top + height)
            shape.Name := "fry_" fry_count
            apply_fry_style(shape)

            ; 应用旋转角度
            if rotation != 0 {
                shape.Rotation := rotation
            }

            fry_count++
        } else {
            ; 普通形状
            shape := slide.Shapes.AddShape(1, left, top, width, height) ; 1 = msoShapeRectangle
            shape.Name := "chip_" chip_count
            apply_chip_style(shape)

            ; 应用旋转角度
            if rotation != 0 {
                shape.Rotation := rotation
            }

            chip_count++
        }

        return true

    } catch as err {
        OutputDebug("绘制形状失败: " err.message)
        return false
    }
}

; 应用 chip 样式（形状）- 使用射线渐变，从右下角开始
apply_chip_style(shape) {
    try {
        ; 设置射线渐变类型，从右下角开始
        shape.Fill.TwoColorGradient(1, 1) ;

        ; 获取渐变停止点集合
        gradient_stops := shape.Fill.GradientStops

        ; 清除默认的停止点
        ; gradient_stops.Delete()

        ; 添加前景色停止点 (0% 位置) - 右下角
        stop1 := gradient_stops(1) ; 紫色
        stop1.Color.RGB := rgb2bgr("GREEN")
        stop1.Transparency := 0.9 ; 80%透明度

        ; 添加背景色停止点 (100% 位置) - 向外扩散
        stop2 := gradient_stops(2) ; 同样的紫色
        stop2.Color.RGB := rgb2bgr("GREEN")
        stop2.Transparency := 1 ; 95%透明度

        ; 设置虚线边框
        shape.Line.ForeColor.RGB := rgb2bgr("BLACK")
        shape.Line.Weight := 0.5
        shape.Line.DashStyle := 4 ; 虚线

        return true

    } catch as err {
        OutputDebug("应用 chip 样式失败: " err.message)
        return false
    }
}

; 应用 fry 样式（线条）
apply_fry_style(shape) {
    try {
        ; 设置虚线线条
        shape.Line.ForeColor.RGB := rgb2bgr("RED")
        shape.Line.Weight := 0.5
        shape.Line.DashStyle := 4 ; 虚线

        return true

    } catch as err {
        OutputDebug("应用 fry 样式失败: " err.message)
        return false
    }
}

; 测试函数：验证倾斜角度功能
test_rotation_function() {
    try {
        objects_data := read_patato_json()

        if !objects_data.Has("objects") || objects_data["objects"].Length = 0 {
            MsgBox("没有找到对象数据")
            return
        }

        rotation_count := 0
        for obj in objects_data["objects"] {
            if obj.Has("rotation") && obj["rotation"] != 0 {
                rotation_count++
                OutputDebug("对象旋转角度: " obj["rotation"])
            }
        }

        MsgBox("共找到 " rotation_count " 个有旋转角度的对象")

    } catch as err {
        MsgBox("测试失败: " err.message)
    }
}

; #endregion

; #region  extract_pptx_objects

; 主函数：根据活动窗口提取 PPT 对象
extract_pptx_objects() {
    try {
        ; 检查当前活动窗口
        if WinActive("ahk_exe Eagle.exe") {
            ; Eagle 操作：复制文件路径并提取 PPT 对象
            return extract_pptx_objects_from_eagle()
        } else if WinActive("ahk_class PPTFrameClass") {
            ; PowerPoint 操作：直接提取当前激活页面的对象
            return extract_pptx_objects_from_active_presentation()
        } else {
            ; 都不是，显示提示信息
            MsgBox("请先激活 Eagle 或 PowerPoint 窗口")
            return false
        }
    } catch as err {
        MsgBox("提取失败: " err.message)
        return false
    }
}

; 从 Eagle 提取 PPT 对象
extract_pptx_objects_from_eagle() {
    ; 保存当前剪贴板内容
    original_clipboard := ClipboardAll()

    try {
        ; 发送 Ctrl+Alt+C 复制 Eagle 中的文件路径
        Send("^!c")
        Sleep(500) ; 等待复制完成

        ; 获取剪贴板中的文件路径
        ppt_path := A_Clipboard

        ; 检查文件是否存在且是 pptx 文件
        if !FileExist(ppt_path) || !InStr(ppt_path, ".pptx") {
            throw Error("未能获取有效的 pptx 文件路径。请确保：`n1. 已选中一个 pptx 文件`n2. Eagle 的 Ctrl+Alt+C 快捷键未被修改")
        }

        ; 管理 PPT 应用
        ppt_app := ""
        app_was_running := false

        try {
            ppt_app := ComObjActive("PowerPoint.Application")
            app_was_running := true
        } catch {
            ; PowerPoint 没有运行，创建新实例
            ppt_app := ComObject("PowerPoint.Application")
            app_was_running := false
        }

        ; 隐藏 PowerPoint 窗口
        original_visible := ppt_app.Visible
        ppt_app.Visible := true

        ; 记录打开前的演示文稿数量
        presentations_count_before := ppt_app.Presentations.Count

        ; 打开演示文稿
        ppt_pres := ppt_app.Presentations.Open(ppt_path)

        ; 提取对象信息到映射
        data_map := extract_pptx_objects_to_map(ppt_pres)

        ; 关闭演示文稿
        ppt_pres.Close()

        ; 保存数据到文件
        output_path := extract_map_to_file(data_map)

        ; 恢复 PowerPoint 窗口状态
        ppt_app.Visible := original_visible

        ; 检查是否需要关闭 PowerPoint 应用
        presentations_count_after := ppt_app.Presentations.Count
        if !app_was_running && presentations_count_after = 0 {
            ; 应用是我们启动的，且没有其他打开的演示文稿，关闭应用
            ppt_app.Quit()
        }

        ; 恢复原始剪贴板内容
        Clipboard := original_clipboard

        MsgBox("PPT 对象信息提取完成！`n保存位置: " output_path)
        return output_path

    } catch as err {
        ; 恢复原始剪贴板内容
        Clipboard := original_clipboard

        ; 确保在出错时恢复 PowerPoint 可见性
        if IsSet(ppt_app) && ppt_app {
            try {
                ppt_app.Visible := original_visible
            } catch {
                ; 忽略恢复可见性时的错误
            }
        }

        throw err
    }
}

; 从当前激活的 PowerPoint 演示文稿提取对象
extract_pptx_objects_from_active_presentation() {
    try {
        ; 获取当前活动的 PowerPoint 应用程序
        ppt_app := ComObjActive("PowerPoint.Application")
        if !ppt_app || !ppt_app.ActivePresentation {
            throw Error("未找到打开的 PowerPoint 演示文稿")
        }

        ; 获取当前活动的演示文稿
        ppt_pres := ppt_app.ActivePresentation

        ; 提取对象信息到映射
        data_map := extract_pptx_objects_to_map(ppt_pres)

        ; 保存数据到文件
        output_path := extract_map_to_file(data_map)

        MsgBox("PPT 对象信息提取完成！`n保存位置: " output_path)
        return output_path

    } catch as err {
        throw err
    }
}

master_view1(ppt_app := "") {
    if (ppt_app = "") {
        try {
            ppt_app := ComObjActive("PowerPoint.Application")
        } catch {
            MsgBox("错误：PPT 未运行，请先打开 PowerPoint。")
            return
        }
    }

    try {
        app := ComObjActive("PowerPoint.Application")
        cb := app.CommandBars("Slide Master View")
        if (cb.Visible) {
            app.CommandBars.ExecuteMso("ViewThumbnailViewPowerPoint")   ; 普通视图
            Notify.show_green("普通视图")
        } else {
            app.CommandBars.ExecuteMso("ViewSlideMasterView")   ; 母版视图
            Notify.show_blue("母版视图")
        }
    }
    catch as err
        Msgbox("打开母版视图时出错: " . err.Message)
}

; 根据单个对象数据绘制形状
draw_shape_from_dat2a(slide, obj_data, &chip_count, &fry_count) {
    try {
        left := obj_data["left"]
        top := obj_data["top"]
        width := obj_data["width"]
        height := obj_data["height"]

        ; 获取倾斜角度，如果不存在则默认为0
        rotation := obj_data.Has("rotation") ? obj_data["rotation"] : 0

        ; 使用固定 0.1 阈值分类
        if height <= 0.1 {
            ; 水平线
            shape := slide.Shapes.AddLine(left, top, left + width, top)
            shape.Name := "fry_" fry_count
            apply_fry_style(shape)

            ; 应用旋转角度
            if rotation != 0 {
                shape.Rotation := rotation
            }

            fry_count++
        } else if width <= 0.1 {
            ; 垂直线
            shape := slide.Shapes.AddLine(left, top, left, top + height)
            shape.Name := "fry_" fry_count
            apply_fry_style(shape)

            ; 应用旋转角度
            if rotation != 0 {
                shape.Rotation := rotation
            }

            fry_count++
        } else {
            ; 普通形状
            shape := slide.Shapes.AddShape(1, left, top, width, height) ; 1 = msoShapeRectangle
            shape.Name := "chip_" chip_count
            apply_chip_style(shape)

            ; 应用旋转角度
            if rotation != 0 {
                shape.Rotation := rotation
            }

            chip_count++
        }

        return true

    } catch as err {
        OutputDebug("绘制形状失败: " err.message)
        return false
    }
}

; 从 PowerPoint 演示文稿提取对象数据到映射
extract_pptx_objects_to_map(ppt_pres) {
    try {
        ; 准备数据存储

        objects_data := Map()
        objects_data["slide_info"] := Map()

        ; 获取幻灯片尺寸 - 通过 SlideMaster 获取
        objects_data["slide_info"]["width"] := Round(ppt_pres.SlideMaster.Width, 1)
        objects_data["slide_info"]["height"] := Round(ppt_pres.SlideMaster.Height, 1)
        objects_data["slide_info"]["extraction_time"] := A_Now

        objects_data["objects"] := []

        ; 检查幻灯片数量
        if ppt_pres.Slides.Count = 0 {
            throw Error("PPT 文件中没有幻灯片")
        }

        ; 获取第一张幻灯片
        slide := ppt_pres.Slides(1)

        ; 遍历所有形状对象
        for shape in slide.Shapes {
            obj_info := Map()

            ; 基本几何属性
            obj_info["left"] := Round(shape.Left, 1)
            obj_info["top"] := Round(shape.Top, 1)
            obj_info["width"] := Round(shape.Width, 1)
            obj_info["height"] := Round(shape.Height, 1)

            ; 添加倾斜角度属性
            try {
                ; 尝试获取旋转角度（倾斜角度）
                obj_info["rotation"] := Round(shape.Rotation, 1)
            } catch {
                ; 如果无法获取旋转角度，设置为 0
                obj_info["rotation"] := 0
            }

            ; 添加到对象列表
            objects_data["objects"].Push(obj_info)
        }

        return objects_data

    } catch as err {
        throw err
    }
}

; 将数据映射保存到文件
extract_map_to_file(data_map) {
    try {
        ; 生成输出文件路径
        output_path := A_ScriptDir "\patato.json"

        ; 直接覆盖保存
        json_text := json.dumps(data_map, 2)
        FileOpen(output_path, "w", "UTF-8").Write(json_text)

        return output_path

    } catch as err {
        throw Error("保存文件失败: " err.message)
    }
}

; 读取 patato.json 文件
read_patato_json() {
    json_path := A_ScriptDir "\patato.json"

    if !FileExist(json_path) {
        throw Error("patato.json 文件不存在: " json_path)
    }

    try {
        json_text := FileRead(json_path, "UTF-8")
        objects_data := json.load(json_text)
        return objects_data
    } catch as err {
        throw Error("读取 patato.json 文件时出错: " err.message)
    }
}

; CopyFromEagle or File
add_slide_from_eagle(ppt_app := "") {
    try {
        SendInput("!^c")
        Sleep(200)
        file_name := A_Clipboard
        if (ppt_app = "") {
            ppt_app := ComObjActive("PowerPoint.Application")
        }
        slides := ppt_app.ActivePresentation.Slides
        layout := slides(1).CustomLayout
        count := slides.Count
        slides.InsertFromFile(file_name, count)
    } catch as err {
        Msgbox("从eagle中复制PPT页面时发生错误:" . err.Message)
    }
}

; #endregion

; #region  pickup_from_sample
; 从模板形状拾取并应用到当前选中形状
pickup_from_sample(sample_slide := "sample_cups", sample_shape := "cup_1") {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        if (!ppt_application) {
            Msgbox("无法创建PowerPoint应用!")
            return 0
        }
    } catch as err {
        Msgbox("无法创建PowerPoint应用!")
        return 0
    }

    try {
        selection := ppt_application.ActiveWindow.Selection         ; 获取当前选择对象
        if selection.type < 2 {
            Msgbox("没有选中任何形状!")
            return 0
        }

        if (selection.HasChildShapeRange) {
            shape_range := selection.ChildShapeRange
        } else {
            shape_range := selection.ShapeRange
        }
    } catch as err {
        Msgbox("选择对象时,发生意外错误!")
        return 0
    }

    try {
        source_shape := ppt_application.ActivePresentation.Slides(sample_slide).Shapes(sample_shape)
        ; source_shape := ppt_application.ActivePresentation.Slides(sample_slide).Shapes(sample_shape)
        source_shape.PickUp
    } catch as err {
        Msgbox("提取样本形状时,发生意外错误!")
        return 0
    }

    try {
        for shape in shape_range {
            shape.Apply
        }
    } catch as err {
        Msgbox("复制格式时,发生意外错误!")
        return 0
    }
}

pickup_from_slidemaster(sample_slide := "sample_cups", sample_shape := "text_style_heading_1") {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        if (!ppt_application) {
            Msgbox("无法创建PowerPoint应用!")
            return 0
        }
    } catch as err {
        Msgbox("无法创建PowerPoint应用!")
        return 0
    }

    try {
        selection := ppt_application.ActiveWindow.Selection         ; 获取当前选择对象
        if selection.type < 2 {
            Msgbox("没有选中任何形状!")
            return 0
        }

        if (selection.HasChildShapeRange) {
            shape_range := selection.ChildShapeRange
        } else {
            shape_range := selection.ShapeRange
        }
    } catch as err {
        Msgbox("选择对象时,发生意外错误!")
        return 0
    }

    try {
        source_shape := ppt_application.ActivePresentation.SlideMaster.Shapes(sample_shape)
        ; source_shape := ppt_application.ActivePresentation.Slides(sample_slide).Shapes(sample_shape)
        source_shape.PickUp
    } catch as err {
        Msgbox("提取样本形状时,发生意外错误!")
        return 0
    }

    try {
        for shape in shape_range {
            shape.Apply
        }
    } catch as err {
        Msgbox("复制格式时,发生意外错误!")
        return 0
    }
}

pickup_from_sample_cup_1() {
    pickup_from_sample(sample_slide := "sample_cups", sample_shape := "cup_1")
}

pickup_from_sample_cup_2() {
    pickup_from_sample(sample_slide := "sample_cups", sample_shape := "cup_2")
}

pickup_from_sample_cup_3() {
    pickup_from_sample(sample_slide := "sample_cups", sample_shape := "cup_3")
}

pickup_from_sample_cup_4() {
    pickup_from_sample(sample_slide := "sample_cups", sample_shape := "cup_4")
}

pickup_from_sample_cup_5() {
    pickup_from_sample(sample_slide := "sample_cups", sample_shape := "cup_5")
}

pickup_from_sample_cup_6() {
    pickup_from_sample(sample_slide := "sample_cups", sample_shape := "cup_6")
}

; 从模板形状拾取并应用到当前选中形状
copy_from_sample(sample_slide := "sample_cups", sample_shape := "cup_101") {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        if (!ppt_application) {
            Msgbox("无法创建PowerPoint应用!")
            return 0
        }
    } catch as err {
        Msgbox("无法创建PowerPoint应用!")
        return 0
    }

    try {
        selection := ppt_application.ActiveWindow.Selection         ; 获取当前选择对象
        if selection.type < 2 {
            Msgbox("没有选中任何形状!")
            return 0
        }

        if (selection.HasChildShapeRange) {
            shape_range := selection.ChildShapeRange
        } else {
            shape_range := selection.ShapeRange
        }
    } catch as err {
        Msgbox("选择对象时,发生意外错误!")
        return 0
    }

    try {
        source_shape := ppt_application.ActivePresentation.Slides(sample_slide).Shapes(sample_shape)
        ; source_shape := ppt_application.ActivePresentation.Slides(sample_slide).Shapes(sample_shape)
        source_shape.Copy
    } catch as err {
        Msgbox("提取样本形状时,发生意外错误!")
        return 0
    }

    try {
        for shape in shape_range {
            shape.PickUp
            copy_shape := ppt_application.ActiveWindow.View.Slide.Shapes.Paste
            copy_shape.Apply
            copy_shape.Select
            copy_shape.left := shape.left
            copy_shape.top := shape.top
            copy_shape.Width := shape.Width
            copy_shape.Height := shape.Height
            shape.Delete
        }
    } catch as err {
        Msgbox("复制格式时,发生意外错误!")
        return 0
    }
}

copy_from_sample_cup_101() {
    copy_from_sample(sample_slide := "sample_cups", sample_shape := "cup_101")
}

copy_from_sample_cup_102() {
    copy_from_sample(sample_slide := "sample_cups", sample_shape := "cup_102")
}

copy_from_sample_cup_103() {
    copy_from_sample(sample_slide := "sample_cups", sample_shape := "cup_103")
}

copy_from_sample_cup_104() {
    copy_from_sample(sample_slide := "sample_cups", sample_shape := "cup_104")
}

copy_from_sample_cup_105() {
    copy_from_sample(sample_slide := "sample_cups", sample_shape := "cup_105")
}

copy_from_sample_cup_106() {
    copy_from_sample(sample_slide := "sample_cups", sample_shape := "cup_106")
}

; #endregion

; #region get_ppt_object
; 获取ppt对象

get_selection_type() {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection
        return selection_type_map[selection.Type]
    } catch as err {
        return false
    }
}

selection_type_is_none() {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection
        if (selection.Type == 0)
            return true
        else
            return false
    } catch as err {
        return false
    }
}

selection_type_is_slides() {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection
        if (selection.Type == 1)
            return true
        else
            return false
    } catch as err {
        return false
    }
}

selection_type_is_shapes() {
    try {
        ppt_appaa := ComObjActive("PowerPoint.Application")

        selection := ppt_appaa.ActiveWindow.Selection

        if (selection.Type == 2)
            return true
        else
            return false
    } catch as err {

        return false
    }
}

selection_type_is_text() {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection
        if (selection.Type == 3)
            return true
        else
            return false
    } catch as err {
        return false
    }
}

get_ppt_object() {
    try {
        global selection_type_map, SHAPE_TYPE_MAP

        result := {
            success: true, app: "", pres: "", window: "", slide_master: "",
            selection: "", type: "", slide_range: "", slide: "", shape_range: "", shape: "", shape_type: "",
            base_shape: "",
            left: 0, top: 0, width: 0, height: 0, right: 0, bottom: 0, center_x: 0, center_y: 0,
            msg: ""
        }

        ppt_app := ComObjActive("PowerPoint.Application")

        result.app := ppt_app
        result.pres := ppt_app.ActivePresentation
        result.window := ppt_app.ActiveWindow
        result.slide_master := ppt_app.ActivePresentation.SlideMaster

        ; ----------------------------------------------------

        selection := ppt_app.ActiveWindow.Selection
        type := selection_type_map[selection.Type]

        result.selection := selection
        result.type := type
        result.slide_range := ppt_app.ActiveWindow.Selection.SlideRange
        result.slide := ppt_app.ActiveWindow.view.slide

        ; 如果选区类型为 "none" (无选中) 或 "slides" (仅选中幻灯片标签)，直接返回结果
        if (type ~= "i)^none|slides") {
            return result
        }

        ; 获取选中的形状范围
        ; 如果选中的是组合内的形状，优先获取子形状范围
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange

        ; 将形状范围和数量写入结果对象
        result.shape_range := shape_range
        result.shape := shape_range(1)
        result.shape_type := SHAPE_TYPE_MAP[shape_range(1).Type]

        ; ----------------------------------------------------

        left_arr := []   ; 存储 shape.Left
        top_arr := []    ; 存储 shape.Top
        right_arr := []   ; 存储 shape.Right
        bottom_arr := []  ; 存储 shape.Bottom

        for shape in shape_range {
            left_arr.Push(shape.Left)
            top_arr.Push(shape.Top)
            right_arr.Push(shape.Left + shape.Width)
            bottom_arr.Push(shape.Top + shape.Height)
        }

        result.left := Min(left_arr*)
        result.top := Min(top_arr*)
        result.right := Max(right_arr*)
        result.bottom := Max(bottom_arr*)
        result.width := result.right - result.left
        result.height := result.bottom - result.top
        result.center_x := result.left + result.width / 2
        result.center_y := result.top + result.height / 2

        ; ----------------------------------------------------

        for shape in shape_range {
            left_arr.Push(shape.Left)
            top_arr.Push(shape.Top)
        }
        base_shape_index := find_min_in_min(left_arr, top_arr)[1]
        base_shape := shape_range(base_shape_index)
        result.base_shape := base_shape

        ; 返回最终结果
        return result

    } catch as err {
        result.success := false
        result.msg := err.message
        return result
    }
}

; #endregion

; #region  add textbox

; 将屏幕像素坐标转换为PowerPoint点坐标
pixels_points(window, x_px := 0, y_px := 0, w_px := 0, h_px := 0) {
    try {
        ; --- 1. 获取 PPT 原点 (0,0) 在屏幕上的物理像素位置 ---
        ; PointsToScreenPixelsX/Y 会自动处理 DPI 和 Zoom
        origin_x_px := window.PointsToScreenPixelsX(0)
        origin_y_px := window.PointsToScreenPixelsY(0)

        ; --- 2. 获取当前视图缩放比例 ---
        zoom_factor := window.View.Zoom / 100

        ; --- 3. 计算相对于 PPT 左上角的“未缩放”像素差值 ---
        ; 鼠标位置 - 原点位置 = 屏幕上的物理像素偏移量
        screen_offset_x := x_px - origin_x_px
        screen_offset_y := y_px - origin_y_px

        ; --- 4. 逆向推导回 PPT 的 Point 坐标 ---
        ; 公式推导：
        ; ScreenPixels = (Points * Zoom) * (DPI_Scale)
        ; 但 PointsToScreenPixelsX 返回的已经是 (Points * Zoom) 后的物理像素值吗？
        ; 实际上：PointsToScreenPixelsX(P) 返回的是 P 点在屏幕上的像素位置。
        ; 所以：Offset_Pixels = (Target_Point - 0) * Zoom * DPI_Scale
        ; 而 PPT 内部单位是 Point (1/72 英寸)。
        ; 关键点：PointsToScreenPixelsX 返回的值已经包含了 DPI 缩放。
        ; 我们需要消除 Zoom 影响，然后利用 PPT 的 PixelsToPoints (如果有) 或者标准换算。

        ; 【更稳健的方法】：利用 PPT 的逆运算特性
        ; 我们知道：Origin_Pixels = PointsToScreenPixelsX(0)
        ; 假设目标点是 P，则 Target_Pixels = PointsToScreenPixelsX(P)
        ; 那么 Target_Pixels - Origin_Pixels = PointsToScreenPixelsX(P) - PointsToScreenPixelsX(0)
        ; 由于线性关系：Diff_Pixels ≈ P * Zoom * (PixelsPerPoint)
        ; 这里的 PixelsPerPoint 取决于系统 DPI。

        ; 既然不能直接用 API 反推，我们可以构造一个“测试点”来动态获取换算率！
        ; 这种方法比任何硬编码的 DPI 常量都准确。

        ; 动态获取：1 Point 在当前 Zoom 和 DPI 下等于多少屏幕像素？
        ; 取 P=100 (避免浮点误差)，计算其屏幕像素增量
        test_point := 100
        px_at_0 := window.PointsToScreenPixelsX(0)
        px_at_100 := window.PointsToScreenPixelsX(test_point)

        pixels_per_100_points := px_at_100 - px_at_0

        if (pixels_per_100_points == 0)
            throw Error("无法计算像素转换率，视图可能无效。")

        ; 计算 1 Point 对应的像素值 (已包含 Zoom 和 DPI)
        pixels_per_point := pixels_per_100_points / test_point

        ; --- 5. 执行转换 ---
        x_pt := screen_offset_x / pixels_per_point
        y_pt := screen_offset_y / pixels_per_point

        ; 处理宽高 (如果提供)
        w_pt := (w_px > 0) ? (w_px / pixels_per_point) : 0
        h_pt := (h_px > 0) ? (h_px / pixels_per_point) : 0

        return { success: true, x: x_pt, y: y_pt, w: w_pt, h: h_pt }
    } catch as err {
        return { success: false, message: "像素转换失败：" . err.Message, x: 0, y: 0, w: 0, h: 0 }
    }
}

position_by_frame(shape_range := "") {
    try {
        x_arr := []
        y_arr := []
        w_arr := []
        h_arr := []

        for shape in shape_range {
            if (shape.name ~= "i)^frame") {
                x_arr.Push(shape.Left)
                y_arr.Push(shape.Top)
                w_arr.Push(shape.width)
                h_arr.Push(shape.height)
                shape.Left += 960
            }
        }
        if (x_arr.Length)
            return { success: true, x_arr: x_arr, y_arr: y_arr, w_arr: w_arr, h_arr: h_arr, msg: "成功提取位置信息" }
        else
            return { success: false, x_arr: [], y_arr: [], w_arr: [], h_arr: [], msg: "提取位置信息失败" }

    } catch as err {
        return { success: false, x_arr: [], y_arr: [], w_arr: [], h_arr: [], msg: "err.Message" }
    }
}

; add_textbox2(sample := "", ppt_command := "") {
;     try {
;         obj := get_ppt_object()
;         obj.slide_master.Shapes(sample).PickUp
;         width := obj.slide_master.Shapes(sample).Width

;         pos := {success: true, x_arr: [], y_arr: [], w_arr: [], h_arr: [], msg: ""}

;         if RegExMatch(sample, "标题")
;             text := TEXT_8[Random(1, TEXT_8.Length)]
;         else
;             text := TEXT_30[Random(1, TEXT_30.Length)]

;         if (obj.type ~= "i)^shapes|text") {
;             pos := position_by_frame(obj.shape_range)
;             if !(pos.success) {
;                 for shape in obj.shape_range
;                     if (shape.Type == "17")
;                         shape.Apply
;                 return
;             }
;         } else {
;             if RegExMatch(sample, "一级标题") {
;                 h1_pos := obj.slide_master.Shapes("一级标题位置").textFrame.textRange.text
;                 array := StrSplit(h1_pos, "/")
;                 pos.x_arr.Push(array[1])
;                 pos.y_arr.Push(array[2])
;             } else {
;                 MouseGetPos(&x, &y)
;                 point := pixels_points(obj.window, x, y)
;                 pos.x_arr.Push(point.x)
;                 pos.y_arr.Push(point.y)
;             }
;         }

;         obj.selection.Unselect()

;         for x in pos.x_arr {
;             new_shape := obj.slide.Shapes.AddTextbox(1, pos.x_arr[A_Index], pos.y_arr[A_Index], 10, 10)
;             new_shape.Apply
;             new_shape.Width := width
;             new_shape.TextFrame.TextRange.Text := text
;             new_shape.Select(0)
;         }

;     } catch as err {
;         MsgBox("插入文本失败: " . err.Message)
;     }
; }

apply_textbox(sample := "") {
    try {
        obj := get_ppt_object()
        obj.slide_master.Shapes(sample).PickUp

        if (obj.type ~= "i)^shapes|text") {
            for shape in obj.shape_range
                if (shape.Type == "17")
                    shape.Apply
        }

    } catch as err {
        MsgBox("应用文本格式失败: " . err.Message)
    }
}

add_capsule(sample := "") {
    try {
        ; if (ppt_app = "") {
        ;     ppt_app := ComObjActive("PowerPoint.Application")
        ; }

        obj := get_ppt_object()

        ; 1. 拾取格式 (PickUp)
        if (sample != "")
            obj.slide_master.Shapes(sample).PickUp

        ; 2. 获取当前选中的对象
        selection := obj.selection.ShapeRange

        if (selection.Count == 0) {
            MsgBox("请先选中一个文本框！")
            return
        }

        target_shape := selection(1)

        ; 3. 计算胶囊的尺寸和位置
        padding_x := 20
        padding_y := 10

        new_width := target_shape.Width + (padding_x * 2)
        new_height := target_shape.Height + (padding_y * 2)

        ; 中心对齐计算
        new_left := target_shape.Left + (target_shape.Width / 2) - (new_width / 2)
        new_top := target_shape.Top + (target_shape.Height / 2) - (new_height / 2)

        ; 4. 添加形状 (2 = msoShapeRoundedRectangle)
        new_capsule := obj.slide.Shapes.AddShape(5, new_left, new_top, new_width, new_height)

        ; 5. 设置最大圆角 (1.0 = 胶囊状)
        new_capsule.Adjustments.Item[1] := 1.0

        ; 6. 应用格式
        new_capsule.Apply

        ; 7. 将胶囊移动到最底层 (1 = msoSendToBack)
        new_capsule.ZOrder(1)
        name1 := new_capsule.Name
        name2 := target_shape.Name
        ; shape_names := [new_capsule.Name, target_shape.Name]
        ; MsgBox(shape_names[2], , "T2")

        ; grouped_shape := obj.slide.Shapes.Range(name1, name2).Group
        new_capsule.Select(0)

        grouped_shape := obj.selection.ShapeRange.Group
        grouped_shape.Select(1)

    } catch as err {
        MsgBox("插入胶囊失败: " . err.Message)
    }
}

add_textbox(sample := "") {
    try {
        obj := get_ppt_object()
        obj.slide_master.Shapes(sample).PickUp
        width := obj.slide_master.Shapes(sample).Width

        pos := { success: true, x_arr: [], y_arr: [], w_arr: [], h_arr: [], msg: "" }

        if RegExMatch(sample, "标题")
            text := TEXT_8[Random(1, TEXT_8.Length)]
        else
            text := TEXT_30[Random(1, TEXT_30.Length)]

        pos := position_by_frame(obj.shape_range)
        if !(pos.success) {
            MouseGetPos(&x, &y)
            point := pixels_points(obj.window, x, y)
            pos.x_arr.Push(point.x)
            pos.y_arr.Push(point.y)
        }

        obj.selection.Unselect()

        for x in pos.x_arr {
            new_shape := obj.slide.Shapes.AddTextbox(1, pos.x_arr[A_Index], pos.y_arr[A_Index], 10, 10)
            new_shape.Apply
            new_shape.Width := width
            new_shape.TextFrame.TextRange.Text := text
            new_shape.Select(0)
        }

    } catch as err {
        MsgBox("插入文本失败: " . err.Message)
    }
}

; --- 一级标题 ---
add_h11() => add_textbox("一级标题_默认")
add_h12() => add_textbox("一级标题_品牌")
add_h13() => add_textbox("一级标题_强调")
add_h14() => add_textbox("一级标题_反色")

apply_h11() => apply_textbox("一级标题_默认")
apply_h12() => apply_textbox("一级标题_品牌")
apply_h13() => apply_textbox("一级标题_强调")
apply_h14() => apply_textbox("一级标题_反色")

; --- 二级标题 ---
add_h21() => add_textbox("二级标题_默认")
add_h22() => add_textbox("二级标题_品牌")
add_h23() => add_textbox("二级标题_强调")
add_h24() => add_textbox("二级标题_反色")

apply_h21() => apply_textbox("二级标题_默认")
apply_h22() => apply_textbox("二级标题_品牌")
apply_h23() => apply_textbox("二级标题_强调")
apply_h24() => apply_textbox("二级标题_反色")

; --- 三级标题 ---
add_h31() => add_textbox("三级标题_默认")
add_h32() => add_textbox("三级标题_品牌")
add_h33() => add_textbox("三级标题_强调")
add_h34() => add_textbox("三级标题_反色")

apply_h31() => apply_textbox("三级标题_默认")
apply_h32() => apply_textbox("三级标题_品牌")
apply_h33() => apply_textbox("三级标题_强调")
apply_h34() => apply_textbox("三级标题_反色")

; --- 四级标题 ---
add_h41() => add_textbox("四级标题_默认")
add_h42() => add_textbox("四级标题_品牌")
add_h43() => add_textbox("四级标题_强调")
add_h44() => add_textbox("四级标题_反色")

apply_h41() => apply_textbox("四级标题_默认")
apply_h42() => apply_textbox("四级标题_品牌")
apply_h43() => apply_textbox("四级标题_强调")
apply_h44() => apply_textbox("四级标题_反色")

; --- 五级标题 ---
add_h51() => add_textbox("五级标题_默认")
add_h52() => add_textbox("五级标题_品牌")
add_h53() => add_textbox("五级标题_强调")
add_h54() => add_textbox("五级标题_反色")

apply_h51() => apply_textbox("五级标题_默认")
apply_h52() => apply_textbox("五级标题_品牌")
apply_h53() => apply_textbox("五级标题_强调")
apply_h54() => apply_textbox("五级标题_反色")

; --- 正文 ---
add_b11() => add_textbox("正文_默认")
add_b12() => add_textbox("正文_浅色")
add_b13() => add_textbox("正文_品牌")
add_b14() => add_textbox("正文_反色")

apply_b11() => apply_textbox("正文_默认")
apply_b12() => apply_textbox("正文_浅色")
apply_b13() => apply_textbox("正文_品牌")
apply_b14() => apply_textbox("正文_反色")

add_shape(sample := "") {
    try {
        obj := get_ppt_object()
        obj.slide_master.Shapes(sample).PickUp

        pos := { success: true, x_arr: [], y_arr: [], w_arr: [], h_arr: [], msg: "" }

        if (obj.type ~= "i)^shapes") {
            pos := position_by_frame(obj.shape_range)
            if !(pos.success) {
                for shape in obj.shape_range
                    if (shape.Type == "1")
                        shape.Apply
                return
            }
        } else {
            MouseGetPos(&x, &y)
            point := pixels_points(obj.window, x, y)
            pos.x_arr.Push(point.x)
            pos.y_arr.Push(point.y)
            pos.w_arr.Push(500)
            pos.h_arr.Push(400)
        }

        obj.selection.Unselect()

        for x in pos.x_arr {
            new_shape := obj.slide.Shapes.AddShape(5, pos.x_arr[A_Index], pos.y_arr[A_Index], pos.w_arr[A_Index],
                pos.h_arr[
                    A_Index])
            new_shape.Adjustments.Item[1] := 0.05
            new_shape.Apply
            new_shape.Select(0)
        }
    } catch
        throw
}

add_shapess(ppt_app := "", sample := "") {
    try {
        obj := get_ppt_object()
        obj.slide_master.Shapes(sample).PickUp

        pos := { success: true, x_arr: [], y_arr: [], w_arr: [], h_arr: [], msg: "" }

        if (obj.type ~= "i)^shapes") {
            pos := position_by_frame(obj.shape_range)
            if !(pos.success) {
                for shape in obj.shape_range
                    if (shape.Type == "1")
                        shape.Apply
                return
            }
        } else {
            MouseGetPos(&x, &y)
            point := pixels_points(obj.window, x, y)
            pos.x_arr.Push(point.x)
            pos.y_arr.Push(point.y)
            pos.w_arr.Push(500)
            pos.h_arr.Push(400)
        }

        obj.selection.Unselect()

        for x in pos.x_arr {
            new_shape := obj.slide.Shapes.AddShape(5, pos.x_arr[A_Index], pos.y_arr[A_Index], pos.w_arr[A_Index],
                pos.h_arr[
                    A_Index])
            new_shape.Adjustments.Item[1] := 0.05
            new_shape.Apply
            new_shape.Select(0)
        }
    } catch
        throw
}

add_shape_1() {
    add_shape("容器_默认")
}

add_shape_2() {
    add_shape("容器_品牌")
}
add_shape_3() {
    add_shape("容器_强调")
}

add_shape_4() {
    add_shape("容器_反色")
}

add_capsule_1() {
    add_capsule("容器_默认")
}

add_capsule_2() {
    add_capsule("容器_品牌")
}
add_capsule_3() {
    add_capsule("容器_强调")
}

add_capsule_4() {
    add_capsule("容器_反色")
}

; insert_heading_13() {
;     insert_heading_at_cursor("一级标题_强调")
; }

; insert_heading_14() {
;     insert_heading_at_cursor("一级标题_反色")
; }

; insert_heading_21() {
;     insert_heading_at_cursor("二级标题_默认")
; }

; insert_heading_22() {
;     insert_heading_at_cursor("二级标题_品牌")
; }

; insert_heading_23() {
;     insert_heading_at_cursor("二级标题_强调")
; }

; insert_heading_24() {
;     insert_heading_at_cursor("二级标题_反色")
; }

apply_shape_001(shape) {
    try {
        shape.Fill.ForeColor.RGB := 0x0000FF
        shape.Fill.BackColor.RGB := 0x0000FF
        shape.Fill.Patterned(40)
        shape.Fill.Transparency := 0.9
        shape.Line.ForeColor.RGB := 0x0000FF
        shape.Line.Weight := 1.5
        shape.Line.DashStyle := 4
        shape.locked := -1
    } catch
        throw
}

; 函数：获取下一个可用的框架索引
; 参数: ppt_shapes - PowerPoint 的形状集合对象
; 返回: 整数 (例如 1, 2, 3)
get_min_shape_index(shapes, prefix := "layout_frame_") {
    try {
        used_numbers := Map() ; 创建一个映射表来存储已存在的数字

        ; 1. 遍历所有形状，收集已使用的编号
        for shape in shapes {
            if (InStr(shape.Name, prefix) == 1) {
                ; 提取前缀后面的部分
                suffix := StrReplace(shape.Name, prefix)

                ; 尝试转换为整数
                num := 0
                try num := Integer(suffix)

                ; 只记录正整数
                if (num > 0) {
                    used_numbers[num] := true
                }
            }
        }

        ; 2. 从 1 开始寻找第一个未被使用的数字 (填补空缺逻辑)
        next_index := 1
        while (used_numbers.Has(next_index)) {
            next_index++
        }

        return next_index

    } catch
        throw
}

add_frame(ppt_app := "", position := "") {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")

        if !position {
            region := get_screen_region()
            pt := pixels_points(ppt_app.ActiveWindow, region.x, region.y, region.w, region.h)
        } else {
            pt := {}
            pt.x := position["left"]
            pt.y := position["top"]
            pt.w := position["width"]
            pt.h := position["height"]
        }

        shapes := ppt_app.ActiveWindow.View.Slide.Shapes
        name_prefix := "layout_frame_"

        ; 调用函数获取最小可用数字
        next_num := get_min_shape_index(shapes, name_prefix)

        ; 格式化为三位数 (例如 001, 002)
        final_shape_name := name_prefix . Format("{:03d}", next_num)

        shape := shapes.AddShape(1, pt.x, pt.y, pt.w, pt.h)
        shape.Name := final_shape_name
        shape.Select(1)
        apply_shape_001(shape)
    } catch
        throw
}

; PATH_PRESET_LAYOUTS_JSON

position_h1() {
    try {
        obj := get_ppt_object()
        text := obj.pres.SlideMaster.Shapes("一级标题位置").textFrame.textRange.text
        array := StrSplit(text, "/")
        shape := obj.shape

        if array[1]
            shape.Left := array[1]
        else
            shape.Left := 60

        if array[2]
            shape.Top := array[2]
        else
            shape.Top := 30

    } catch as err {
        MsgBox("定位失败: " . err.Message)
    }
}

; #endregion

; #region guides

; ==============================================================================
; 【普通幻灯片参考线操作函数】
; 作用：直接操作当前活动幻灯片的参考线集合 (Slide.Guides)。
; 注意：这些参考线仅存在于当前幻灯片，不会影响母版或其他幻灯片。
; ==============================================================================

; 1. 添加垂直参考线到当前幻灯片
; 参数: x_array* - 可变参数，接收多个 X 轴坐标值 (单位：磅 points)
; 逻辑: 遍历坐标数组，在指定 X 位置创建垂直参考线 (Orientation=2)
add_vertical_slide_guides(x_array*) {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        guides := obj.pres.Guides
        for x in x_array
            guides.Add(2, x)
    } catch as err {
        MsgBox("插入垂直幻灯片参考线失败: " err.Message)
    }
}

; 2. 添加水平参考线到当前幻灯片
; 参数: y_array* - 可变参数，接收多个 Y 轴坐标值 (单位：磅 points)
; 逻辑: 遍历坐标数组，在指定 Y 位置创建水平参考线 (Orientation=1)
add_horizontal_slide_guides(y_array*) {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        guides := obj.pres.Guides
        for y in y_array
            guides.Add(1, y)
    } catch as err {
        MsgBox("插入水平幻灯片参考线失败: " err.Message)
    }
}

; 3. 删除当前幻灯片的所有参考线
; 逻辑: 循环删除 Guides 集合中的第一个元素，直到集合为空
; 注意: 必须从索引 1 开始删除，因为删除后后续元素索引会自动前移
delete_slide_guides() {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        guides := obj.pres.Guides
        while (guides.Count > 0) {
            guides(1).Delete
        }
    } catch as err {
        MsgBox("删除幻灯片参考线失败: " err.Message)
    }
}

delete_horizontal_guides() {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        v_guides := []
        for guide in obj.pres.Guides {
            if (guide.Orientation == 1) {
                v_guides.Push(guide)
            }
        }
        while (v_guides.Length > 0) {
            v_guides[1].Delete
            v_guides.RemoveAt(1)

        }
    } catch as err {
        MsgBox("删除幻灯片参考线失败: " err.Message)
    }
}

delete_vertical_guides() {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        v_guides := []
        for guide in obj.pres.Guides {
            if (guide.Orientation == 2) {
                v_guides.Push(guide)
            }
        }
        while (v_guides.Length > 0) {
            v_guides[1].Delete
            v_guides.RemoveAt(1)

        }
    } catch as err {
        MsgBox("删除幻灯片参考线失败: " err.Message)
    }
}

; ------------------------------------------

; 1. 默认基础参考线 (仅左右边界 + 水平辅助线)
add_default_slide_guides() {
    delete_vertical_guides()
    add_vertical_slide_guides(56, 904)
    add_horizontal_slide_guides(78, 510)
}

; 2. 2列布局参考线 (列宽408, 间距24)
add_2column_slide_guides() {
    delete_vertical_guides()
    add_vertical_slide_guides(60, 468, 492, 900)
    add_horizontal_slide_guides(100, 500)
}

; 3. 3列布局参考线 (列宽264, 间距24)
add_3column_slide_guides() {
    delete_vertical_guides()
    add_vertical_slide_guides(60, 324, 348, 612, 636, 900)
    add_horizontal_slide_guides(100, 500)
}

; 4. 4列布局参考线 (列宽192, 间距24)
add_4column_slide_guides() {
    delete_vertical_guides()
    add_vertical_slide_guides(60, 252, 276, 468, 492, 684, 708, 900)
    add_horizontal_slide_guides(100, 500)
}

; 5. 5列布局参考线 (列宽152, 间距20)
add_5column_slide_guides() {
    delete_vertical_guides()
    add_vertical_slide_guides(60, 212, 232, 384, 404, 556, 576, 728, 748, 900)
    add_horizontal_slide_guides(100, 500)
}

; 6. 6列布局参考线 (列宽120, 间距24)
add_6column_slide_guides() {
    delete_vertical_guides()
    add_vertical_slide_guides(60, 180, 204, 324, 348, 468, 492, 612, 636, 756, 780, 900)
    add_horizontal_slide_guides(100, 500)
}

; 1. 1行布局
add_1row_slide_guides() {
    delete_horizontal_guides()
    add_vertical_slide_guides(60, 900)
    add_horizontal_slide_guides(100, 500)
    Notify.show("插入1行幻灯片参考线")
}

; 1. 2行布局 (行高 188, 间距 24)
; 坐标: 100, 288, 312, 500
add_2row_slide_guides() {
    delete_horizontal_guides()
    add_vertical_slide_guides(60, 900)
    add_horizontal_slide_guides(100, 288, 312, 500)
    Notify.show("插入2行幻灯片参考线")

}

; 2. 3行布局 (行高 ~117.33, 间距 24)
add_3row_slide_guides() {
    delete_horizontal_guides()
    add_vertical_slide_guides(60, 900)
    add_horizontal_slide_guides(100, 217.33, 241.33, 358.67, 382.67, 500)
    Notify.show("插入3行幻灯片参考线")
}

; 3. 4行布局 (行高 82, 间距 24)
add_4row_slide_guides() {
    delete_horizontal_guides()
    add_vertical_slide_guides(60, 900)
    add_horizontal_slide_guides(100, 182, 206, 288, 312, 394, 418, 500)
    Notify.show("插入4行幻灯片参考线")
}

; ==============================================================================
; 【母版参考线操作函数】
; 作用：直接操作幻灯片母版的参考线集合 (SlideMaster.Guides)。
; 注意：这些参考线定义在母版上，会自动应用到所有基于该母版的幻灯片。
; ==============================================================================

; 4. 添加垂直参考线到母版
; 参数: x_array* - 可变参数，接收多个 X 轴坐标值
; 逻辑: 获取 SlideMaster 对象，在其 Guides 集合中创建垂直参考线
add_vertical_master_guides(x_array*) {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        master_guides := obj.slide_master.Guides
        for x in x_array
            master_guides.Add(2, x)
    } catch as err {
        MsgBox("插入垂直母版参考线失败: " err.Message)
    }
}

; 5. 添加水平参考线到母版
; 参数: y_array* - 可变参数，接收多个 Y 轴坐标值
; 逻辑: 获取 SlideMaster 对象，在其 Guides 集合中创建水平参考线
add_horizontal_master_guides(y_array*) {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        master_guides := obj.slide_master.Guides
        for y in y_array
            master_guides.Add(1, y)
    } catch as err {
        MsgBox("插入水平母版参考线失败: " err.Message)
    }
}

; 6. 删除母版上的所有参考线
; 逻辑: 循环删除母版 Guides 集合中的第一个元素，直到清空
; 警告: 此操作会影响所有使用该母版的幻灯片布局
delete_master_guides() {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
        master_guides := obj.slide_master.Guides
        while (master_guides.Count > 0) {
            master_guides(1).Delete
        }
    } catch as err {
        MsgBox("删除母版参考线失败: " err.Message)
    }
}

; ------------------------------------------

; 1. 默认基础母版参考线
add_default_master_guides() {
    delete_master_guides()
    add_vertical_master_guides(60, 900)
    add_horizontal_master_guides(100, 500)
}

; 2. 2列布局母版参考线
add_2column_master_guides() {
    delete_master_guides()
    add_vertical_master_guides(60, 468, 492, 900)
    add_horizontal_master_guides(100, 500)
}

; 3. 3列布局母版参考线
add_3column_master_guides() {
    delete_master_guides()
    add_vertical_master_guides(60, 324, 348, 612, 636, 900)
    add_horizontal_master_guides(100, 500)
}

; 4. 4列布局母版参考线
add_4column_master_guides() {
    delete_master_guides()
    add_vertical_master_guides(60, 252, 276, 468, 492, 684, 708, 900)
    add_horizontal_master_guides(100, 500)
}

; 5. 5列布局母版参考线
add_5column_master_guides() {
    delete_master_guides()
    add_vertical_master_guides(60, 212, 232, 384, 404, 556, 576, 728, 748, 900)
    add_horizontal_master_guides(100, 500)
}

; 6. 6列布局母版参考线
add_6column_master_guides() {
    delete_master_guides()
    add_vertical_master_guides(60, 180, 204, 324, 348, 468, 492, 612, 636, 756, 780, 900)
    add_horizontal_master_guides(100, 500)
}

; ==============================================================================
; 显示/隐藏参考线
; 作用：显示/隐藏参考线
; ==============================================================================

show_guides() {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := true
    } catch as err {
        MsgBox("显示参考线失败: " err.Message)
    }
}

hide_guides() {
    try {
        obj := get_ppt_object()
        obj.app.DisplayGuides := false
    } catch as err {
        MsgBox("显示参考线失败: " err.Message)
    }
}

; ------------------------------------------



; ==============================================================================
; 【功能】根据选中形状的位置绘制4条参考线（左、右、上、下）
; 【用法】选中一个形状后调用此函数
; ==============================================================================
add_guides_from_selected_shape() {
    try {
        ; 1. 获取 PPT 应用程序和活动窗口对象
        ppt_app := ComObjActive("PowerPoint.Application")
        active_win := ppt_app.ActiveWindow
        
        ; 2. 获取当前选中的形状
        selection := active_win.Selection
        if (selection.Type != 2) { ; 2 = ppSelectionShapes
            MsgBox("请先选中一个形状。", "提示", 48)
            return
        }

        ; 处理单个形状或形状组/范围，这里取第一个形状作为基准
        shape_range := selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange
        if (shape_range.Count == 0) {
            MsgBox("未检测到有效形状。", "错误", 16)
            return
        }
        
        target_shape := shape_range.Item(1)

        ; 3. 获取形状的边界坐标
        s_left := target_shape.Left
        s_top := target_shape.Top
        s_width := target_shape.Width
        s_height := target_shape.Height
        
        s_right := s_left + s_width
        s_bottom := s_top + s_height

        ; 4. 获取当前幻灯片的参考线集合
        ; 注意：Slide.Guides 仅影响当前幻灯片。如果需要母版参考线，需使用 obj.slide_master.Guides
        current_slide := active_win.View.Slide
        guides := ppt_app.ActivePresentation.Guides


            
        add_vertical_slide_guides(s_left, s_right)
        add_horizontal_slide_guides(s_top, s_bottom)

        MsgBox("已根据形状 '" target_shape.Name "' 的边界绘制4条参考线。", "成功", "T2")
        

    } catch
        throw
}

; #endregion

; #region resize_by_guides 将形状集合扩展到参考线网格区域

resize_by_guides(mode := "full", guides_type := "slide", keep_aspect := false) {
    try {
        obj := get_ppt_object()

        slide_width := obj.pres.PageSetup.SlideWidth
        slide_height := obj.pres.PageSetup.SlideHeight

        if (guides_type == "slide")
            guides := obj.pres.Guides
        else
            guides := obj.slide_master.Guides

        h_guides := [0]
        v_guides := [0]
        if (guides.Count) {
            for guide in guides {
                if (guide.Orientation == 2) {           ; 垂直参考线
                    v_guides.Push(guide.Position)
                } else {                                ; 水平参考线
                    h_guides.Push(guide.Position)
                }
            }
        }

        h_guides.Push(slide_height)
        v_guides.Push(slide_width)

        shape_range := obj.shape_range

        ; 获取选中区域的整体包围盒 (Bounding Box)
        shape_range_left := obj.left
        shape_range_top := obj.top
        shape_range_width := obj.width
        shape_range_height := obj.height
        center_x := obj.center_x
        center_y := obj.center_y

        ; 找到中心点所在的垂直区间 (由垂直参考线决定的左右边界)
        x_bounds := get_floor_and_ceil(v_guides, center_x)
        area_left := x_bounds.floor
        area_width := x_bounds.ceil - x_bounds.floor

        ; 找到中心点所在的水平区间 (由水平参考线决定的上下边界)
        y_bounds := get_floor_and_ceil(h_guides, center_y)
        area_top := y_bounds.floor
        area_height := y_bounds.ceil - y_bounds.floor

        ; 计算缩放比例
        scale_x := area_width / shape_range_width
        scale_y := area_height / shape_range_height

        ; 纵横比保护逻辑
        if (keep_aspect) {
            ; 取较小的缩放比例，防止变形
            min_scale := Min(scale_x, scale_y)
            scale_x := min_scale
            scale_y := min_scale
            ; 注意：此时形状可能不会填满整个 area，而是居中
            ; 如果需要居中，需要调整 area_left 和 area_top 的计算
            ; 这里简化处理，仅防止变形，位置仍按原逻辑映射
        }

        ; 最核心的算法部分，采用了仿射变换
        for shape in shape_range {
            ; 1. 计算相对偏移 (Normalization)
            ; 算出每个形状相对于“选区整体左上角”的距离
            relative_left := shape.Left - shape_range_left
            relative_top := shape.Top - shape_range_top

            ; 2. 重映射位置 (Remapping)
            ; 新位置 = 目标区域起点 + (相对距离 * 缩放比例)
            shape.Left := area_left + (relative_left * scale_x)
            shape.Top := area_top + (relative_top * scale_y)

            ; 3. 应用尺寸缩放
            if (mode ~= "i)^width|full")
                shape.Width := shape.Width * scale_x

            if (mode ~= "i)^height|full")
                shape.Height := shape.Height * scale_y
        }

    } catch as err {
        MsgBox("将形状集合扩展到参考线网格区域失败: " err.Message)
    }
}

; 夹住目标值的上下界
get_floor_and_ceil(arr, target) {
    ; 检查数组是否为空
    if (arr.Length == 0) {
        return { lower: "", upper: "" }
    }

    ; 对数组进行排序（确保后续逻辑正确）
    sorted_arr := arr.Clone()
    bubble_sort(sorted_arr)

    ; 初始化结果变量
    lower := ""
    upper := ""

    ; 遍历排序后的数组
    for value in sorted_arr {
        if (value <= target) {
            lower := value
        }
        if (value >= target && upper == "") {
            upper := value
        }
    }

    ; 特殊情况处理
    if (lower == "") {
        ; 目标数字比所有数都小，返回最小的两个数
        lower := sorted_arr[1]
        upper := sorted_arr[2]
    } else if (upper == "") {
        ; 目标数字比所有数都大，返回最大的两个数
        len := sorted_arr.Length
        lower := sorted_arr[len - 1]
        upper := sorted_arr[len]
    }

    ; 返回结果
    return { floor: lower, ceil: upper }
}

resize_by_master_guides() {
    resize_by_guides(mode := "full", guides_type := "master", keep_aspect := false)
}

; #endregion

; #region execute_ppt_control

execute_ppt_control(ppt_app := "", control_name := "") {
    try {
        if !ppt_app
            ppt_app := ComObjActive("PowerPoint.Application")

        ppt_app.CommandBars.ExecuteMso(control_name)

    } catch
        throw
}

objects_align_left(ppt_app) {
    execute_ppt_control(ppt_app, "ObjectsAlignLeftSmart")
    Notify.show("左对齐")
}

objects_align_center(ppt_app) {
    execute_ppt_control(ppt_app, "ObjectsAlignCenterHorizontalSmart")
    Notify.show("水平居中")
}

objects_align_right(ppt_app) {
    execute_ppt_control(ppt_app, "ObjectsAlignRightSmart")
    Notify.show("右对齐")
}

objects_align_top(ppt_app) {
    execute_ppt_control(ppt_app, "ObjectsAlignTopSmart")
    Notify.show("顶对齐")
}

objects_align_middle(ppt_app) {
    execute_ppt_control(ppt_app, "ObjectsAlignMiddleVerticalSmart")
    Notify.show("垂直居中")
}

objects_align_bottom(ppt_app) {
    execute_ppt_control(ppt_app, "ObjectsAlignBottomSmart")
    Notify.show("底对齐")
}

run_ppt_command(command) {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        ppt_application.CommandBars.ExecuteMso(command)
    } catch
        throw
}

; #endregion

; #region  get_layout_map

get_layout_map(ppt_app := "") {
    try {
        if (!ppt_app)
            ppt_app := ComObjActive("PowerPoint.Application")

        layout_map := Map()
        layout_map["objects"] := []
        shapes := ppt_app.ActiveWindow.View.Slide.Shapes

        for shape in shapes {
            if (shape.Name ~= "i)^layout_frame") {
                obj_info := Map()
                obj_info["left"] := Round(shape.Left, 1)
                obj_info["top"] := Round(shape.Top, 1)
                obj_info["width"] := Round(shape.Width, 1)
                obj_info["height"] := Round(shape.Height, 1)
                layout_map["objects"].Push(obj_info)
            }
        }

        return layout_map
    } catch
        throw
}

save_layout_map(layout_map, output_path := "") {
    try {
        if (!output_path)
            output_path := A_ScriptDir "\settings\layout_frames.json"

        json_text := json.dumps(layout_map, 2)
        FileOpen(output_path, "w", "UTF-8").Write(json_text)
    } catch
        throw
}

save_layout(layout_map := "") {
    layout_map := get_layout_map(ppt_app := "")
    save_layout_map(layout_map)
}

load_layout(json_path := "") {
    try {
        if !json_path {
            json_path := A_ScriptDir "\settings\layout_frames.json"
        }

        json_text := FileRead(json_path, "UTF-8")
        objects_data := json.load(json_text)
        return objects_data
    } catch
        throw
}

; 根据 json 数据在 PPT 中绘制 frames
apply_preset_layout(ppt_app := "", layout_name := "") {
    global PATH_PRESET_LAYOUTS_JSON
    try {
        json_path := A_ScriptDir . PATH_PRESET_LAYOUTS_JSON
        json_text := FileRead(json_path, "UTF-8")
        pos_map := json.load(json_text)

        if (ppt_app = "")
            ppt_app := ComObjActive("PowerPoint.Application")

        shapes := ppt_app.ActiveWindow.View.Slide.Shapes

        while shapes.Count > 0 {
            shapes(shapes.Count).Delete()
        }

        for position in pos_map[layout_name] {
            add_frame(ppt_app, position)
        }
    } catch
        throw
}

add_layout_horizontal_band(ppt_app := "") {
    apply_preset_layout(ppt_app, "horizontal_band")
    Notify.show("拦腰式布局")
}

add_layout_001(ppt_app := "") {
    apply_preset_layout(ppt_app, "layout_001")
}

add_layout_002(ppt_app := "") {
    apply_preset_layout(ppt_app, "layout_002")
}

; 根据 json 数据在 PPT 中绘制 frames
apply_layout_from_json() {
    global PATH_LAYOUT_JSON
    try {
        json_path := A_ScriptDir . PATH_LAYOUT_JSON
        json_text := FileRead(json_path, "UTF-8")
        pos_map := json.load(json_text)

        ppt_app := ComObjActive("PowerPoint.Application")
        shapes := ppt_app.ActiveWindow.View.Slide.Shapes

        while shapes.Count > 0 {
            shapes(shapes.Count).Delete()
        }

        for position in pos_map["objects"] {
            add_frame(ppt_app, position)
        }
    } catch
        throw
}

; #endregion

; #region  excel
; 设置 Excel 选中单元格的字体
set_font(index := 1) {
    global ch_fonts, en_fonts
    try {
        if WinActive("ahk_class XLMAIN")
            app := ComObjActive("Excel.Application")
        else if WinActive("ahk_class OpusApp")
            app := ComObjActive("Word.Application")

        selection := app.Selection

        ch_font := ch_fonts[index]
        en_font := en_fonts[index]

        selection.Font.Name := ch_font
        selection.Font.Name := en_font

        Notify.show(Format("设置字体为: {1}. {2}", index, ch_font), 1)
    } catch
        throw
}

set_font1() => set_font(1)

set_font2() => set_font(2)

set_font3() => set_font(3)

set_font4() => set_font(4)

set_font5() => set_font(5)

set_font6() => set_font(6)

set_font7() => set_font(7)

set_font8() => set_font(8)


; #endregion



; #region  text






split_paragraphs() {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection := ppt_app.ActiveWindow.Selection
        shape_range := selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange
        
        new_shapes := []

        gap := 0
        for shp in shape_range {
            for paragraph in shp.TextFrame2.TextRange.Paragraphs {
                clean_text := Trim(paragraph.Text, "`r`n`s`t")
                
                if (clean_text = "") 
                    continue

                new_shp := shp.Duplicate()
                new_shp.Top := paragraph.BoundTop - shp.TextFrame.MarginTop + gap
                new_shp.Left := shp.Left
                new_shp.TextFrame.TextRange.Text := clean_text
                gap += 8

                new_shapes.Push(new_shp)
            }            
        }

        count := shape_range.Count
        loop count
            shape_range(Count - A_Index + 1).Delete

        new_top := new_shapes[1].Top

        for new_shp in new_shapes {
            new_text := new_shp.TextFrame.TextRange.Text

            if new_text ~= "i)^#####" {
                ppt_app.ActivePresentation.SlideMaster.Shapes("正文_默认").PickUp
                new_shp.Apply
                new_text := SubStr(new_text, 6)
                new_shp.TextFrame.TextRange.Text := new_text
                
                new_shp.Top := new_top
                new_top += new_shp.Height + 0
            }
            
            if new_text ~= "i)^####" {
                ppt_app.ActivePresentation.SlideMaster.Shapes("四级标题_默认").PickUp
                new_shp.Apply
                new_text := SubStr(new_text, 5)
                new_shp.TextFrame.TextRange.Text := new_text
                
                new_shp.Top := new_top
                new_top += new_shp.Height + 8
            }
            
            if new_text ~= "i)^###" {
                ppt_app.ActivePresentation.SlideMaster.Shapes("三级标题_强调").PickUp
                new_shp.Apply
                new_text := SubStr(new_text, 4)
                new_shp.TextFrame.TextRange.Text := new_text
                
                new_shp.Top := new_top
                new_top += new_shp.Height + 8
            }
            
            if new_text ~= "i)^##" {
                ppt_app.ActivePresentation.SlideMaster.Shapes("二级标题_品牌").PickUp
                new_shp.Apply
                new_text := SubStr(new_text, 3)
                new_shp.TextFrame.TextRange.Text := new_text
                
                new_shp.Top := new_top
                new_top += new_shp.Height + 8
            }
            
            if new_text ~= "i)^#" {
                ppt_app.ActivePresentation.SlideMaster.Shapes("一级标题_默认").PickUp
                new_shp.Apply
                new_text := SubStr(new_text, 2)
                new_shp.TextFrame.TextRange.Text := new_text
                
                new_shp.Top := new_top
                new_top += new_shp.Height + 8
            }

            
        }

        return new_shapes
    } catch 
        throw
}


; #endregion


; #region  css_gridient



; ============================================================
; 统一错误显示
; ============================================================
ShowError(err, title := "脚本运行错误") {
    msg := "❌ " err.Message "`n"
        . "─────────────────────`n"
        . "📄 文件: " RegExReplace(err.File, ".*\\") "`n"
        . "📍 行号: " err.Line "`n"
        . "🔧 函数: " err.What
    if (err.Extra != "")
        msg .= "`n📎 附加: " err.Extra
    MsgBox(msg, title, "Iconx")
}

; ============================================================
; RGB ↔ BGR 转换（PPT COM 使用 BGR）
; ============================================================
; RGB2BGR(rgb_val) {
;     r := (rgb_val >> 16) & 0xFF
;     g := (rgb_val >> 8) & 0xFF
;     b := rgb_val & 0xFF
;     return (b << 16) | (g << 8) | r
; }

; ============================================================
; 去掉 CSS 注释
; ============================================================
RemoveCSSComments(css) {
    return RegExReplace(css, "/\*[\s\S]*?\*/", "")
}

; ============================================================
; 解析颜色
; ============================================================
ParseColor(raw) {
    global system_color_map
    raw := Trim(raw)

    named := StrTitle(StrLower(raw))
    if COLOR_MAP.Has(named)
        return COLOR_MAP[named]

    if RegExMatch(raw, "i)^#([0-9a-f]{6})$", &m)
        return Integer("0x" m[1])

    if RegExMatch(raw, "i)^#([0-9a-f])([0-9a-f])([0-9a-f])$", &m)
        return Integer("0x" m[1] m[1] m[2] m[2] m[3] m[3])

    if RegExMatch(raw, "i)^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", &m)
        return (Number(m[1]) << 16) | (Number(m[2]) << 8) | Number(m[3])

    throw Error("无法识别的颜色: " raw)
}

; ============================================================
; 位置解析
; ============================================================
ParsePosition(pos_str) {
    pos_str := Trim(pos_str)
    if (pos_str = "")
        return -1
    if InStr(pos_str, "%")
        return Number(StrReplace(pos_str, "%")) / 100.0
    return Number(pos_str)
}

; ============================================================
; 智能分割 gradient 参数
; ============================================================
SplitGradientParts(str) {
    parts := []
    depth := 0
    current := ""

    loop parse str {
        ch := A_LoopField
        if (ch == "(")
            depth++
        else if (ch == ")")
            depth--

        if (ch == "," && depth == 0) {
            if (Trim(current) != "")
                parts.Push(Trim(current))
            current := ""
        } else {
            current .= ch
        }
    }

    if (Trim(current) != "")
        parts.Push(Trim(current))

    return parts
}

; ============================================================
; 判断是否为方向
; ============================================================
IsGradientDirection(s) {
    s := Trim(s)
    if RegExMatch(s, "i)^to\s+(top|bottom|left|right)(\s+(top|bottom|left|right))?$")
        return true
    if RegExMatch(s, "i)^-?\d+(?:\.\d+)?deg$")
        return true
    return false
}

; ============================================================
; 自动填充未指定 position
; ============================================================
AutoFillPositions(stops) {
    total := stops.Length
    if (total < 2)
        return

    for i, s in stops {
        if (s["position"] = "")
            s["position"] := -1
    }

    allUnknown := true
    for i, s in stops {
        if (s["position"] != -1) {
            allUnknown := false
            break
        }
    }

    if allUnknown {
        loop total
            stops[A_Index]["position"] := (A_Index - 1) / (total - 1)
        return
    }

    if (stops[1]["position"] = -1)
        stops[1]["position"] := 0.0
    if (stops[total]["position"] = -1)
        stops[total]["position"] := 1.0

    i := 1
    while (i <= total) {
        if (stops[i]["position"] != -1) {
            i += 1
            continue
        }

        leftIdx := i - 1
        while (i <= total && stops[i]["position"] = -1)
            i += 1
        rightIdx := i

        if (leftIdx < 1 || rightIdx > total)
            continue

        leftPos := stops[leftIdx]["position"]
        rightPos := stops[rightIdx]["position"]
        gap := rightIdx - leftIdx

        loop (gap - 1) {
            k := leftIdx + A_Index
            stops[k]["position"] := leftPos + (rightPos - leftPos) * (A_Index / gap)
        }
    }
}

; ============================================================
; 解析单个 stop
; ============================================================
ParseGradientStop(part) {
    part := Trim(part)

    if RegExMatch(part, "i)^(.*?)(?:\s+(\d+(?:\.\d+)?%)\s*)?$", &m) {
        color_str := Trim(m[1])
        pos_str := Trim(m[2])

        if (color_str = "")
            throw Error("stop 颜色为空: " part)

        stop := Map()
        stop["color"] := ParseColor(color_str)
        stop["position"] := ParsePosition(pos_str)
        return stop
    }

    throw Error("无法解析渐变 stop: " part)
}

; ============================================================
; 提取完整函数调用，比如 linear-gradient(...)
; ============================================================
ExtractBalancedFunction(text, startPos) {
    openPos := InStr(text, "(", false, startPos)
    if (!openPos)
        return ""

    depth := 0
    endPos := 0

    loop parse SubStr(text, openPos) {
        ch := A_LoopField
        if (ch == "(")
            depth++
        else if (ch == ")") {
            depth--
            if (depth = 0) {
                endPos := openPos + A_Index - 1
                break
            }
        }
    }

    if (endPos = 0)
        return ""

    return SubStr(text, startPos, endPos - startPos + 1)
}

; ============================================================
; 解析 gradient 字符串
; ============================================================
ParseGradientString(gradient_full) {
    if !RegExMatch(gradient_full, "i)^(linear|radial)-gradient\s*\((.*)\)$", &m)
        throw Error("无法解析 gradient: " gradient_full)

    type := m[1] "-gradient"
    inside := m[2]
    parts := SplitGradientParts(inside)

    if (parts.Length < 2)
        throw Error("gradient 参数不足。")

    result := Map()
    result["type"] := type
    result["direction"] := ""
    result["stops"] := []

    first := Trim(parts[1])
    startIdx := 1

    if IsGradientDirection(first) {
        result["direction"] := first
        startIdx := 2
    }

    loop parts.Length - startIdx + 1 {
        idx := A_Index + startIdx - 1
        part := Trim(parts[idx])
        if (part = "")
            continue
        result["stops"].Push(ParseGradientStop(part))
    }

    if (result["stops"].Length < 2)
        throw Error("至少需要两个颜色节点。")

    AutoFillPositions(result["stops"])
    return result
}

; ============================================================
; 从剪贴板 CSS 中提取最后一个有效 gradient
; 支持：
; - background: #2A7B9B;
; - background: linear-gradient(...);
; - background: -webkit-linear-gradient(...);
; - 多段 CSS
; - 注释
; - {} 包裹
; ============================================================
GetGradientFromCSS() {
    css := Trim(A_Clipboard)
    if (css = "")
        throw Error("剪贴板为空，请先复制 CSS。")

    css := RemoveCSSComments(css)
    css := StrReplace(css, "-webkit-linear-gradient", "linear-gradient")

    matches := []
    pos := 1
    while RegExMatch(css, "i)linear-gradient\s*\(", &m, pos) {
        start := m.Pos(0)
        full := ExtractBalancedFunction(css, start)
        if (full != "")
            matches.Push(full)
        pos := start + StrLen(full)
    }

    if (matches.Length > 0)
        return ParseGradientString(matches[matches.Length])

    ; 没有 gradient，就尝试纯色 background
    if RegExMatch(css, "is)background(?:-image)?\s*:\s*([^;{}]+)", &mBg) {
        bg := Trim(mBg[1])
        c := ParseColor(bg)
        info := Map()
        info["type"] := "solid"
        info["direction"] := ""
        info["stops"] := [
            Map("color", c, "position", 0.0),
            Map("color", c, "position", 1.0)
        ]
        return info
    }

    throw Error("未找到任何有效的 linear-gradient。")
}

; ============================================================
; CSS 方向 → PPT 角度
; ============================================================
GetPPTAngle(direction_str) {
    direction_str := Trim(direction_str)
    if (direction_str = "")
        return 90

    if RegExMatch(direction_str, "i)^(-?\d+(?:\.\d+)?)deg$", &m) {
        angle := Number(m[1])
        return Mod(angle + 270, 360)
    }

    dirMap := Map(
        "to top",          270,
        "to bottom",       90,
        "to right",        0,
        "to left",         180,
        "to top right",    315,
        "to right top",    315,
        "to top left",     225,
        "to left top",     225,
        "to bottom right", 45,
        "to right bottom", 45,
        "to bottom left",  135,
        "to left bottom",  135
    )

    lower := StrLower(direction_str)
    if dirMap.Has(lower)
        return dirMap[lower]

    return 90
}

; ============================================================
; 创建标题
; ============================================================
AddTitle(slide, text) {
    title := slide.Shapes.AddTextbox(1, 56, 20, 800, 30)
    title.TextFrame2.TextRange.Text := text
    title.TextFrame2.TextRange.Font.Size := 16
    title.TextFrame2.TextRange.Font.Bold := -1
    title.TextFrame2.TextRange.Font.Name := "Microsoft YaHei"
    title.TextFrame2.TextRange.Font.Fill.ForeColor.RGB := 0x222222
    title.TextFrame2.VerticalAnchor := 3
    return title
}

; ============================================================
; 创建标签
; ============================================================
AddLabel(slide, text, x, y, w := 160, h := 24) {
    label := slide.Shapes.AddTextbox(1, x, y, w, h)
    label.TextFrame2.TextRange.Text := text
    label.TextFrame2.TextRange.Font.Size := 10
    label.TextFrame2.TextRange.Font.Name := "Microsoft YaHei"
    label.TextFrame2.TextRange.Font.Fill.ForeColor.RGB := 0x333333
    label.TextFrame2.VerticalAnchor := 3
    return label
}

; ============================================================
; 在 PPT 中绘制渐变
; ============================================================
DrawGradientInPPT(info) {
    try {
        ppt := ComObjActive("PowerPoint.Application")
    } catch {
        throw Error("未找到正在运行的 PowerPoint，请先打开 PPT。")
    }

    if !ppt.ActiveWindow
        throw Error("PowerPoint 没有活动窗口。")

    slide := ppt.ActiveWindow.View.Slide
    stops := info["stops"]

    if (stops.Length < 2)
        throw Error("至少需要两个颜色节点。")

    ; 标题
    titleText := "CSS Gradient Preview"
    if (info["direction"] != "")
        titleText .= "  |  " info["direction"]
    AddTitle(slide, titleText)

    ; 程序图标块
    ; icon := slide.Shapes.AddShape(5, 844, 30, 60, 60)
    ; icon.Adjustments.Item[1] := 0.1
    ; icon.Line.Visible := 0
    ; icon.Fill.ForeColor.RGB := 0x2D7DFF
    ; icon.TextFrame2.TextRange.Text := "CSS"
    ; icon.TextFrame2.TextRange.Font.Size := 16
    ; icon.TextFrame2.TextRange.Font.Bold := -1
    ; icon.TextFrame2.TextRange.Font.Name := "Microsoft YaHei"
    ; icon.TextFrame2.TextRange.Font.Fill.ForeColor.RGB := 0xFFFFFF
    ; icon.TextFrame2.VerticalAnchor := 3
    ; icon.TextFrame2.TextRange.ParagraphFormat.Alignment := 2

    ; 主渐变块
    shape := slide.Shapes.AddShape(1, 0, 60, 960, 180)
    shape.Name := "CSS_Gradient_" A_TickCount
    shape.Line.Visible := 0
    shape.Fill.TwoColorGradient(1, 1)

    gStops := shape.Fill.GradientStops
    gStops(1).Color.RGB := RGB2BGR(stops[1]["color"])
    gStops(1).Position  := stops[1]["position"]
    gStops(2).Color.RGB := RGB2BGR(stops[2]["color"])
    gStops(2).Position  := stops[2]["position"]

    if (stops.Length >= 3) {
        loop stops.Length - 2 {
            s := stops[A_Index + 2]
            try gStops.Insert(RGB2BGR(s["color"]), s["position"])
        }
    }

    shape.Fill.GradientAngle := GetPPTAngle(info["direction"])

    ; 色块标题
    AddLabel(slide, "Colors:", 56, 255, 80, 20)

    ; 颜色预览块
    previewTop := 280
    previewSize := 70
    previewGap := 20
    maxPerRow := 6

    loop stops.Length {
        idx := A_Index
        s := stops[idx]

        row := Floor((idx - 1) / maxPerRow)
        col := Mod(idx - 1, maxPerRow)

        left := 56 + col * (previewSize + previewGap)
        top := previewTop + row * 100

        swatch := slide.Shapes.AddShape(5, left, top, previewSize, previewSize)   
        swatch.Adjustments.Item[1] := 0.1
        swatch.Name := "Color_Swatch_" idx
        swatch.Line.Visible := 0
        swatch.Fill.ForeColor.RGB := RGB2BGR(s["color"])

        AddLabel(
            slide,
            Format("{1}. {2}`n{3}%", idx, RGB2Hex(s["color"]), Round(s["position"] * 100)),
            left, top + 75, previewSize + 30, 30
        )

    }
}

; ============================================================
; 快捷键入口
; ============================================================
css_gridient() {
    try {
        info := GetGradientFromCSS()
        DrawGradientInPPT(info)
        MsgBox(
            "✅ 渐变已生成！`n"
            . "类型: " info["type"] "`n"
            . "方向: " (info["direction"] != "" ? info["direction"] : "默认") "`n"
            . "颜色数: " info["stops"].Length,
            "完成", "T2"
        )
    } catch as err {
        ShowError(err, "CSS → PPT 渐变生成失败")
    }
}


; ============================================================
; RGB 整数 → #RRGGBB 十六进制字符串
; ============================================================
RGB2Hex(rgb_val) {
    r := (rgb_val >> 16) & 0xFF
    g := (rgb_val >> 8) & 0xFF
    b := rgb_val & 0xFF
    return Format("#{:02X}{:02X}{:02X}", r, g, b)
}




;#endregion

; end