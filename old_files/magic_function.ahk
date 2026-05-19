; #region Cursor类库完整功能参考
/** 
 * ========================================================================
 * Cursor类库所有可用功能和属性
 * 
 * 光标类型枚举：
 * - IDC_ARROW      标准箭头光标
 * - IDC_IBEAM      文本输入光标（I型）  
 * - IDC_WAIT       等待光标（沙漏/转圈）
 * - IDC_CROSS      十字光标
 * - IDC_UPARROW    上箭头光标
 * - IDC_HAND       手型光标（链接）
 * - IDC_NO         禁止光标（圆圈斜杠）
 * - IDC_SIZEWE     水平调整大小 ↔
 * - IDC_SIZENS     垂直调整大小 ↕
 * - IDC_SIZEALL    全方向调整大小 ✥
 * - IDC_SIZENWSE   对角调整大小 ↖↘
 * - IDC_SIZENESW   对角调整大小 ↗↙
 * - IDC_HELP       帮助光标（箭头+问号）
 * - IDC_APPSTARTING 应用启动（箭头+沙漏）
 * 
 * 基本属性：
 * current_id := Cursor.ID                    ; 光标ID
 * current_name := Cursor.Name                ; 光标名称  
 * current_type := Cursor.IDCName             ; 光标类型名称
 * current_handle := Cursor.Handle            ; 光标句柄
 * is_visible := Cursor.IsVisible             ; 是否可见
 * is_suppressed := Cursor.IsSuppressed       ; 是否被抑制
 * shadow_enabled := Cursor.Shadow            ; 阴影是否启用
 * 
 * 位置信息：
 * current_position := Cursor.Position        ; 逻辑位置 {X, Y}
 * pos_x := Cursor.Position.X                 ; X坐标
 * pos_y := Cursor.Position.Y                 ; Y坐标
 * physical_position := Cursor.PhysicalPosition ; 物理位置（高DPI）
 * phys_x := Cursor.PhysicalPosition.X        ; 物理X坐标
 * phys_y := Cursor.PhysicalPosition.Y        ; 物理Y坐标
 * 
 * 光标限制区域：
 * clip_rect := Cursor.ClipRect               ; 限制矩形
 * clip_left := Cursor.ClipRect.Left          ; 左边界
 * clip_top := Cursor.ClipRect.Top            ; 上边界  
 * clip_right := Cursor.ClipRect.Right        ; 右边界
 * clip_bottom := Cursor.ClipRect.Bottom      ; 下边界
 * 
 * 光标拖尾功能：
 * trail_enabled := Cursor.Trail.IsEnabled    ; 拖尾是否启用
 * trail_length := Cursor.Trail.Length        ; 拖尾长度
 * 
 * 光标控制方法：
 * Cursor.Shadow := true                      ; 设置阴影开/关
 * Cursor.Trail.Enable(8)                     ; 启用拖尾，长度8
 * Cursor.Trail.Disable()                     ; 禁用拖尾
 * Cursor.Set("Arrow", custom_cursor_handle)  ; 设置自定义光标
 * Cursor.Restore("Arrow")                    ; 恢复指定光标为默认
 * Cursor.RestoreDefaults()                   ; 恢复所有光标为默认
 * Cursor.Destroy(cursor_handle)              ; 销毁光标句柄
 * ========================================================================
 */
; #endregion

; #region Excel选择性粘贴通用函数
/** 
 * ========================================================================
 * Excel选择性粘贴相关功能
 * @param {Integer} paste_type 粘贴类型常量
 * @param {String} message 提示信息
 * 
 * 粘贴类型枚举：
 * -4163 = xlPasteValues (值)
 * -4122 = xlPasteFormats (格式) 
 * -4144 = xlPasteFormulas (公式)
 * 12 = xlPasteValuesAndNumberFormats (值+数字格式)
 * 11 = xlPasteFormulasAndNumberFormats (公式+数字格式)
 * 8 = xlPasteColumnWidths (列宽)
 * 7 = xlPasteAllExceptBorders (除边框外全部)
 * -4104 = xlPasteAll (全部)
 * 6 = xlPasteAllUsingSourceTheme (使用源主题粘贴全部)
 * 4 = xlPasteComments (批注)
 * 5 = xlPasteValidation (有效性)
 * ========================================================================
 */

excel_paste_special(paste_type, message := "") {
    try {
        xl := ComObjActive("Excel.Application")
        
        ; 检查是否有复制的内容
        if (xl.Application.CutCopyMode == 0) {
            ToolTip("没有复制的内容")
            SetTimer(() => ToolTip(), -1000)
            return false
        }
        
        ; 执行选择性粘贴
        xl.Selection.PasteSpecial(paste_type)
        
        ; 清除复制状态
        xl.Application.CutCopyMode := 0
        
        ; 显示提示信息
        if (message != "") {
            ToolTip(message)
            SetTimer(() => ToolTip(), -1000)
        }
        
        return true
        
    } catch as err {
        ; ToolTip("粘贴失败: " . err.Message)
        ; SetTimer(() => ToolTip(), -2000)
        return false
    }
}

; 只粘贴值
excel_paste_values() {
    excel_paste_special(-4163, "已粘贴值")
}

; ; 只粘贴格式
; excel_paste_formats() {
;     excel_paste_special(-4122, "已粘贴格式")
; }

; 只粘贴公式
excel_paste_formulas() {
    result := excel_paste_special(-4144, "已粘贴公式")
    return result
}

; 粘贴值和数字格式
excel_paste_values_and_formats() {
    excel_paste_special(12, "已粘贴值和数字格式")
}

; 粘贴公式和数字格式
excel_paste_formulas_and_formats() {
    excel_paste_special(11, "已粘贴公式和数字格式")
}

; 只粘贴列宽
excel_paste_column_widths() {
    excel_paste_special(8, "已粘贴列宽")
}

; 除边框外全部粘贴
excel_paste_all_except_borders() {
    excel_paste_special(7, "已粘贴(除边框外)")
}
; 全部粘贴
excel_paste_all() {
    result := excel_paste_special(-4104, "已粘贴全部")
    return result
}

; 匹配目标格式粘贴（方法1：COM实现）
excel_paste_match_destination() {
    try {
        xl := ComObjActive("Excel.Application")
        
        ; 检查剪贴板内容
        if (A_Clipboard == "") {
            ToolTip("剪贴板为空")
            SetTimer(() => ToolTip(), -1000)
            return false
        }
        
        ; ; 获取目标单元格的格式
        ; target_format := xl.Selection.NumberFormat
        
        ; 粘贴为纯文本
        xl.Selection.Value := A_Clipboard
        
        ; ; 应用目标格式（如果有特定格式）
        ; if (target_format != "General" && target_format != "") {
        ;     xl.Selection.NumberFormat := target_format
        ; }
        
        ToolTip("已匹配目标格式粘贴")
        SetTimer(() => ToolTip(), -1000)
        return true
        
    } catch as err {
        ; ToolTip("粘贴失败: " . err.Message)
        ; SetTimer(() => ToolTip(), -2000)
        return false
    }
}




; ; 快捷键绑定
; ^+v:: excel_paste_values()                    ; Ctrl+Shift+V
; ^+f:: excel_paste_formats()                   ; Ctrl+Shift+F
; ^+r:: excel_paste_formulas()                  ; Ctrl+Shift+R
; ^+n:: excel_paste_values_and_formats()        ; Ctrl+Shift+N
; ^+m:: excel_paste_formulas_and_formats()      ; Ctrl+Shift+M
; ^+w:: excel_paste_column_widths()             ; Ctrl+Shift+W
; ^+a:: excel_paste_all_except_borders()        ; Ctrl+Shift+A
; #endregion


; smart_paste() {
; 	if WinActive("ahk_exe EXCEL.EXE")  && superkey {
;         ; MsgBox A_Clipboard

;         ; 首先判断 剪贴板中是纯文本还是单元格，如果剪贴板是单元格，判断选中的如果是单元格，执行excel_paste_all()，
;         ; 如果剪贴板中是纯文本，判断选中的如果是单元格，执行excel_paste_match_destination()，如果选中的是文本执行^v
;         ; selection_type := get_excel_state_final()
;         ; MsgBox selection_type.state
;         ; excel_get_state() 
;         ; MsgBox excel_get_state()
;         ; if (!excel_paste_all()) {  
;         ;     if (!excel_paste_match_destination()) {       
; 	    ;         Send("^v")
;         ;     }  
;         ; }        
; 		return
; 	}

; 	if WinActive("ahk_class OpusApp") {
; 		paste_text_only_in_word()
; 		return
; 	} 
; 	Send("^v")	
; }



; get_selection_address() {
;     try {
;         excel_application := ComObjActive("Excel.Application")  
        
;         ; 关键判断：尝试获取Selection对象
;         try {
;             selection := excel_application.Selection
;             ; 如果能成功获取Selection，说明选中了单元格
;             address := selection.Address
;             return {
;                 state: True,
;                 address: address
;             }
;         } catch {
;             ; 获取Selection出错，说明在编辑状态
;             return {
;                 state: False,
;                 address: 0
;             }
;         }
        
;     } catch as err {
;         return {
;             state: "error",
;             description: "无法访问Excel: " . err.Message
;         }
;     }
; }

; #hotif winactive("ahk_class XLMAIN") && superkey
; 	g:: {
;         dispatch_key_action("excel")
;     }
; #hotif

; run_paste() {
; 	if WinActive("ahk_exe EXCEL.EXE") {
;         cursor_type := Cursor.IDCName         
;         if cursor_type == ("IDC_IBEAM") {
; 		    Send("^v")	
;         }

;         if cursor_type !== ("IDC_IBEAM") {
; 		   excel_paste_all()
;         }
; 		return
; 	}

; 	if WinActive("ahk_class OpusApp") {
; 		paste_text_only_in_word()
; 		return
; 	} 
; 	Send("^v")	
; }


; #region 智能粘贴函数 - 简化版
/** 
 * ========================================================================
 * 智能粘贴函数 - 使用简化的剪贴板判断方法
 * 
 * 简化逻辑：
 * - A_Clipboard 为空 → 剪贴板是Excel单元格数据
 * - A_Clipboard 有内容 → 剪贴板是纯文本
 * ========================================================================
 */

; 简化的剪贴板类型判断
get_clipboard_type_simple() {
    ; MsgBox "dd" . A_Clipboard . "ee"
    if (A_Clipboard == "") {
        return "excel_cells"  ; 剪贴板是单元格数据
    } else {
        return "plain_text"   ; 剪贴板是纯文本
    }
}

; 获取Excel选择状态（复用之前的函数）
get_selection_address() {
    try {
        excel_application := ComObjActive("Excel.Application")  
        try {
            selection := excel_application.Selection
            address := selection.Address
            return {
                state: true,
                address: address,
                type: "cells_selected"
            }
        } catch {
            return {
                state: false,
                address: "",
                type: "text_editing"
            }
        }
    } catch as err {
        return {
            state: "error",
            address: "",
            type: "error"
        }
    }
}

; 简化的智能粘贴主函数
smart_paste() {
    if (WinActive("ahk_exe EXCEL.EXE") && superkey) {
        excel_smart_paste_v2()

;         ; excel_application := ComObjActive("Excel.Application")         
;         ; Send("^v")             
;         ; try {
;         ;     selection := excel_application.Selection
;         ;     Send("{Enter}")
        ; }         
        return
    }
    
    if (WinActive("ahk_class OpusApp")) {
        ; paste_text_only_in_word()
        return
    } 
    
    Send("^v")   
}
 





        ; command := "PasteDestinationFormatting"
        ; xl := ComObjActive("Excel.Application")
        ; xl.CommandBars.ExecuteMso(command) 
        
        ; 简化的剪贴板判断
        ; clipboard_type := get_clipboard_type_simple()
        ; MsgBox(get_excel_clipboard_formats()[1], , "T2")
        ; show_clipboard_formats()
        ; clipboard_type := ""
        ; clipboard_type := get_clipboard_type_by_formats()
        ; ; MsgBox(clipboard_type, , "T2")


        ; if clipboard_type == "excel_cells" {
        ;     MsgBox("I am coming!", , "T2")

        ;     excel_paste_all()
        ;     ToolTip("粘贴单元格")
        ; } else {
        ;         Send("^v")
        ;         ToolTip("粘贴到编辑区域")

        ; }

        ; SetTimer(() => ToolTip(), -2000)
        ; return


        
        ; 获取Excel选择状态
        ; selection_info := get_selection_address()
        
        ; 简化的决策逻辑
        ; if (clipboard_type == "excel_cells") {
        ;     ; 剪贴板是单元格数据
        ;     if (selection_info.state == true) {
        ;         excel_paste_all()
        ;         ToolTip("粘贴单元格到: " . selection_info.address)
        ;     } else {
        ;         Send("^v")
        ;         ToolTip("粘贴到编辑区域")
        ;     }
            
        ; } else {
        ;     ; 剪贴板是纯文本 (clipboard_type == "plain_text")
        ;     if (selection_info.state == true) {
        ;         excel_paste_match_destination()
        ;         ToolTip("文本匹配格式粘贴到: " . selection_info.address)
        ;     } else {
        ;         Send("^v")
        ;         ToolTip("插入文本")
        ;     }
        ; }
        
        ; SetTimer(() => ToolTip(), -2000)


; 超简洁的调试信息
show_paste_info_simple() {
    if (!WinActive("ahk_exe EXCEL.EXE"))
        return
        
    clipboard_type := get_clipboard_type_by_formats()
    selection_info := get_selection_address()
    
    message := "剪贴板: " . clipboard_type . "`n"
    message .= "选择: " . selection_info.type . "`n"
    
    ; 预测操作
    if (clipboard_type == "excel_cells") {
        if (selection_info.state == true) {
            message .= "→ excel_paste_all()"
        } else {
            message .= "→ 直接粘贴"
        }
    } else {
        if (selection_info.state == true) {
            message .= "→ 匹配格式粘贴"
        } else {
            message .= "→ 插入文本"
        }
    }
    
    ToolTip(message)
    SetTimer(() => ToolTip(), -3000)
}

; 一行式判断函数
is_clipboard_excel_cells() {
    return (A_Clipboard == "")
}

is_clipboard_plain_text() {
    return (A_Clipboard != "")
}

; 最简洁的智能粘贴版本
smart_paste_minimal() {
    if (WinActive("ahk_exe EXCEL.EXE") && superkey) {
        selection_info := get_selection_address()
        
        if (A_Clipboard == "") {
            ; 剪贴板是单元格
            if (selection_info.state) {
                excel_paste_all()
            } else {
                Send("^v")
            }
        } else {
            ; 剪贴板是文本
            if (selection_info.state) {
                excel_paste_match_destination()
            } else {
                Send("^v")
            }
        }
        return
    }
    
    Send("^v")
}

; 快捷键
; F12:: show_paste_info_simple()  ; 显示简化信息
; #endregion





; #region Excel剪贴板格式检测
/**

========================================================================
Excel剪贴板格式检测 - VBA方法的AHK实现

使用 Application.ClipboardFormats 精确判断剪贴板内容类型
========================================================================
*/

; 获取Excel剪贴板格式列表
get_excel_clipboard_formats() {
    try {
        xl := ComObjActive("Excel.Application")
        formats := xl.Application.ClipboardFormats
        
        format_list := []
        ; 遍历所有格式
        for format in formats {
            format_list.Push(format)
        }
        
        return format_list
        
    } catch as err {
        return []
    }
}

; Excel剪贴板格式常量
excel_clipboard_constants := {
    xlClipboardFormatRTF: 7,          ; RTF格式
    xlClipboardFormatText: 0,         ; 纯文本
    xlClipboardFormatBIFF8: 57,       ; Excel 8.0格式
    xlClipboardFormatBIFF: 8,         ; Excel格式
    xlClipboardFormatPICT: 1,         ; 图片格式
    xlClipboardFormatBitmap: 9,       ; 位图
    xlClipboardFormatCGM: 7,          ; CGM格式
    xlClipboardFormatCSV: 5,          ; CSV格式
    xlClipboardFormatDIF: 4,          ; DIF格式
    xlClipboardFormatDspText: 12,     ; 显示文本
    xlClipboardFormatEmbeddedObject: 21, ; 嵌入对象
    xlClipboardFormatEmbedSource: 22, ; 嵌入源
    xlClipboardFormatLink: 11,        ; 链接
    xlClipboardFormatLinkSource: 23,  ; 链接源
    xlClipboardFormatLinkSourceDesc: 32, ; 链接源描述
    xlClipboardFormatMovie: 24,       ; 电影
    xlClipboardFormatNative: 14,      ; 本机格式
    xlClipboardFormatObjectDesc: 31,  ; 对象描述
    xlClipboardFormatObjectLink: 19,  ; 对象链接
    xlClipboardFormatOwnerLink: 17,   ; 所有者链接
    xlClipboardFormatPrintPICT: 3,    ; 打印图片
    xlClipboardFormatStandardFont: 28, ; 标准字体
    xlClipboardFormatStandardScale: 29, ; 标准比例
    xlClipboardFormatSYLK: 6,         ; SYLK格式
    xlClipboardFormatTable: 16,       ; 表格
    xlClipboardFormatToolFace: 25,    ; 工具面
    xlClipboardFormatToolFacePICT: 26, ; 工具面图片
    xlClipboardFormatVALU: 1          ; 值
}

; 检查剪贴板是否包含特定格式
clipboard_contains_format(format_constant) {
    try {
        xl := ComObjActive("Excel.Application")
        formats := xl.Application.ClipboardFormats
        
        for format in formats {
            if (format == format_constant) {
                return true
            }
        }
        return false
        
    } catch {
        return false
    }
}

; 智能判断剪贴板类型（基于Excel格式）
get_clipboard_type_by_formats() {
    try {
        xl := ComObjActive("Excel.Application")
        formats := xl.Application.ClipboardFormats
        
        has_biff := false
        has_text := false
        has_rtf := false
        has_bitmap := false
        
        ; 检查所有格式
        for format in formats {
            switch format {
                case 57, 8:  ; BIFF8 或 BIFF
                    has_biff := true
                case 0:      ; 纯文本
                    has_text := true  
                case 7:      ; RTF
                    has_rtf := true
                case 9:      ; 位图
                    has_bitmap := true
            }
        }
        
        ; 根据格式组合判断类型
        if (has_biff) {
            return "excel_cells"
        } else if (has_rtf) {
            return "rich_text"
        } else if (has_bitmap) {
            return "image"
        } else if (has_text) {
            return "plain_text"
        } else {
            return "unknown"
        }
        
    } catch {
        return "error"
    }
}

; 显示剪贴板格式详细信息
show_clipboard_formats() {
    try {
        xl := ComObjActive("Excel.Application")
        formats := xl.Application.ClipboardFormats
        
        message := "=== Excel剪贴板格式 ===`n"
        
        format_names := Map()
        ; 创建格式名称映射
        for name, value in excel_clipboard_constants.OwnProps() {
            format_names[value] := name
        }
        
        for format in formats {
            format_name := format_names.Has(format) ? format_names[format] : "Unknown"
            message .= format . " - " . format_name . "`n"
        }
        
        ; 显示判断结果
        clipboard_type := get_clipboard_type_by_formats()
        message .= "`n判断类型: " . clipboard_type
        
        MsgBox(message, "剪贴板格式分析")
        
    } catch as err {
        MsgBox("无法获取剪贴板格式: " . err.Message)
    }
}

; 基于格式的智能粘贴函数
smart_paste_by_formats() {
    if (WinActive("ahk_exe EXCEL.EXE") && superkey) {
        
        ; 使用格式检测判断剪贴板类型
        clipboard_type := get_clipboard_type_by_formats()
        
        ; 获取Excel选择状态
        selection_info := get_selection_address()
        
        ; 根据格式类型决定粘贴方式
        switch clipboard_type {
            case "excel_cells":
                ; Excel单元格数据
                if (selection_info.state == true) {
                    excel_paste_all()
                    ToolTip("粘贴Excel数据")
                } else {
                    Send("^v")
                    ToolTip("粘贴到编辑区域")
                }
                
            case "rich_text":
                ; RTF格式
                if (selection_info.state == true) {
                    excel_paste_all()  ; 保留格式
                    ToolTip("粘贴富文本")
                } else {
                    Send("^v")
                }
                
            case "plain_text":
                ; 纯文本
                if (selection_info.state == true) {
                    excel_paste_match_destination()
                    ToolTip("文本匹配格式粘贴")
                } else {
                    Send("^v")
                    ToolTip("插入文本")
                }
                
            case "image":
                ; 图片
                Send("^v")
                ToolTip("粘贴图片")
                
            default:
                ; 未知或错误
                Send("^v")
                ToolTip("默认粘贴")
        }
        
        SetTimer(() => ToolTip(), -2000)
        return
    }

    Send("^v")
}

; 快速检查函数
is_clipboard_excel_data() {
    return clipboard_contains_format(57) || clipboard_contains_format(8)  ; BIFF8 或 BIFF
}

is_clipboard_rich_text() {
    return clipboard_contains_format(7)  ; RTF
}

is_clipboard_plain_text_only() {
    return clipboard_contains_format(0) && !clipboard_contains_format(7) && !clipboard_contains_format(57)
}

; ; 快捷键
; F12:: show_clipboard_formats()           ; 显示剪贴板格式
; ^F12:: smart_paste_by_formats()          ; 基于格式的智能粘贴
; #endregion



; #region Excel应用程序对象管理
/** 
 * ========================================================================
 * Excel应用程序对象获取函数 - 基于PPT版本改造
 * 
 * 功能：获取并缓存Excel COM对象，自动处理对象失效问题
 * ========================================================================
 */

; 基于公共函数的剪贴板判断
clipboard_has_cells_v2() {
    get_excel_application()
    if (!excel_application) {
        return false
    }
    
    try {
        formats := excel_application.ClipboardFormats  ; 修正这里
        
        ; 检查是否包含Excel格式（BIFF8=57, BIFF=8）
        for format in formats {
            if (format == 57 || format == 8) {
                return true
            }
        }
        return false
    } catch {
        return false
    }
}

; 基于公共函数的选择判断
selection_is_cells_v2() {
    get_excel_application()
    if (!excel_application) {
        return false
    }
    
    try {
        selection := excel_application.Selection  ; 尝试获取Selection
        return true  ; 能获取到说明选中单元格
    } catch {
        return false  ; 获取失败说明在编辑状态
    }
}

; 基于公共函数的选择地址获取
get_selection_address_v2() {
    get_excel_application()
    if (!excel_application) {
        return {
            state: "error",
            address: "",
            type: "no_excel"
        }
    }
    
    try {
        selection := excel_application.Selection
        address := selection.Address
        return {
            state: true,
            address: address,
            type: "cells_selected"
        }
    } catch {
        return {
            state: false,
            address: "",
            type: "text_editing"
        }
    }
}
 
; 更新后的智能粘贴函数
excel_smart_paste_v2() {
    get_excel_application()
    if (!excel_application) {
        return  ; Excel未激活或无法获取对象
    }
    
    ; 判断两个条件
    has_cells := clipboard_has_cells_v2()
    is_cells := selection_is_cells_v2()
    
    ; 根据排列组合执行不同操作
    if (has_cells && is_cells) {
        ; 情况1：剪贴板有单元格 + 选中单元格
        action_cells_to_cells_v2()
        
    } else if (has_cells && !is_cells) {
        ; 情况2：剪贴板有单元格 + 选中文本
        action_cells_to_text_v2()
        
    } else if (!has_cells && is_cells) {
        ; 情况3：剪贴板无单元格 + 选中单元格
        action_text_to_cells_v2()
        
    } else {
        ; 情况4：剪贴板无单元格 + 选中文本
        action_text_to_text_v2()
    }
}

; 更新后的操作函数（传入xl对象）
action_cells_to_cells_v2() {
    ; 剪贴板有单元格数据 → 粘贴到选中单元格
    try {
        ; excel_application.Selection.PasteSpecial(-4104)  ; xlPasteAll
        ; excel_application.CutCopyMode := 0
        excel_application.CommandBars.ExecuteMso("PasteAll")
        ;  Send("^v")
        center_info("单元格→单元格：完整粘贴")
        ; ToolTip("单元格→单元格：完整粘贴")
    } catch {
        Send("^v")
        center_info("单元格→单元格：默认粘贴")
    }
    ; SetTimer(() => ToolTip(), -1500)
}

action_cells_to_text_v2() {
    ; 剪贴板有单元格数据 → 粘贴到文本编辑区域
    Send("^v")
    center_info("单元格→文本：直接粘贴")
    ; SetTimer(() => ToolTip(), -1500)
}

action_text_to_cells_v2() {
    ; 剪贴板是纯文本 → 粘贴到选中单元格
    try {
        ; excel_application.Selection.Value := A_Clipboard  ; 匹配目标格式
         excel_application.CommandBars.ExecuteMso("PasteValues")
        center_info("文本→单元格：匹配格式粘贴")
    } catch {
        Send("^v")
        center_info("文本→单元格：默认粘贴")
    }
    ; SetTimer(() => ToolTip(), -1500)
}

action_text_to_text_v2() {
    ; 剪贴板是纯文本 → 粘贴到文本编辑区域
    Send("^v")
    center_info("文本→文本：直接插入")
    ; SetTimer(() => ToolTip(), -1500)
}

; 清理Excel对象（可选，用于脚本退出时）
cleanup_excel_object() {
    global excel_application
    excel_application := ""
}

; 更新后的状态显示函数
show_excel_paste_status_v2() {
    get_excel_application()
    if (!excel_application) {
        center_info("Excel未激活或无法获取对象")
        ; SetTimer(() => ToolTip(), -2000)
        return
    }
    
    has_cells := clipboard_has_cells_v2()
    is_cells := selection_is_cells_v2()
    
    message := "=== Excel粘贴状态 V2 ===`n"
    message .= "剪贴板包含单元格: " . (has_cells ? "是" : "否") . "`n"
    message .= "当前选中单元格: " . (is_cells ? "是" : "否") . "`n"
    message .= "执行操作: "
    
    if (has_cells && is_cells) {
        message .= "单元格→单元格"
    } else if (has_cells && !is_cells) {
        message .= "单元格→文本"
    } else if (!has_cells && is_cells) {
        message .= "文本→单元格"
    } else {
        message .= "文本→文本"
    }
    
    center_info(message)
    ; ToolTip(message)
    ; SetTimer(() => ToolTip(), -3000)
}

; 快捷键更新
; F11:: show_excel_paste_status_v2()  ; 显示状态V2
; ^+v:: excel_smart_paste_v2()        ; 智能粘贴V2

; 脚本退出时清理（可选）
; OnExit(cleanup_excel_object)
; #endregion



; 清理Excel对象（可选，用于脚本退出时）
function_undo() {
    if (WinActive("ahk_exe EXCEL.EXE") && superkey) {
         get_excel_application()
        if (!excel_application) {
            Send("^z")
            return
        }
        excel_application.Undo      
        return
    }
    
    ; Send("^z")
}



; #region Office应用程序对象完整释放管理
/** 
 * ========================================================================
 * Office应用程序对象完整释放 - Word、Excel、PowerPoint统一管理
 * 
 * 功能：统一管理所有Office COM对象，定期释放，避免启动慢问题
 * ========================================================================
 */

; 全局Office对象变量
global excel_application := ""
global word_application := ""
global ppt_application := ""

; 获取Excel应用程序对象
get_excel_application() {
    global excel_application
    
    if (!WinActive("ahk_exe EXCEL.EXE")) {
        return 0
    }
    
    if (excel_application) {
        try {
            name := excel_application.Name
        } catch {
            excel_application := ""
        }
    }
    
    if (!excel_application) {
        try {
            excel_application := ComObjActive("Excel.Application")
        } catch as err {
            Msgbox("无法获取Excel应用程序:" . err.Message)
            return 0
        }
    }
    
    return excel_application
}

; 获取Word应用程序对象
get_word_application() {
    global word_application
    
    if (!WinActive("ahk_class OpusApp")) {
        return 0
    }
    
    if (word_application) {
        try {
            name := word_application.Name
        } catch {
            word_application := ""
        }
    }
    
    if (!word_application) {
        try {
            word_application := ComObjActive("Word.Application")
        } catch as err {
            Msgbox("无法获取Word应用程序:" . err.Message)
            return 0
        }
    }
    
    return word_application
}

; 释放Excel应用程序对象
release_excel_application() {
    global excel_application
    
    if (excel_application) {
        try {
            ; 尝试访问属性测试对象是否有效
            name := excel_application.Name
            ; 对象有效，释放引用
            excel_application := ""
            return "released"
        } catch {
            ; 对象无效，直接清空
            excel_application := ""
            return "invalid"
        }
    }
    return "empty"
}

; 释放Word应用程序对象
release_word_application() {
    global word_application
    
    if (word_application) {
        try {
            name := word_application.Name
            word_application := ""
            return "released"
        } catch {
            word_application := ""
            return "invalid"
        }
    }
    return "empty"
}


; 获取PowerPoint应用程序对象
get_ppt_application() {
    global ppt_application
    
    if (!WinActive("ahk_class PPTFrameClass")) {
        return 0
    }
    
    if (ppt_application) {
        try {
            name := ppt_application.Name
        } catch {
            ppt_application := ""
        }
    }
    
    if (!ppt_application) {
        try {
            ppt_application := ComObjActive("PowerPoint.Application")
        } catch as err {
            ; Msgbox("无法获取PowerPoint应用程序:" . err.Message)
            return 0
        }
    }
    
    return ppt_application
}

; 释放PowerPoint应用程序对象
release_ppt_application() {
    global ppt_application
    
    if (ppt_application) {
        try {
            name := ppt_application.Name
            ppt_application := ""
            return "released"
        } catch {
            ppt_application := ""
            return "invalid"
        }
    }
    return "empty"
}

; 释放所有Office应用程序对象
release_all_group_officelications() {
    static last_release_time := ""
    
    current_time := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    
    ; 释放各个应用程序对象
    excel_status := release_excel_application()
    word_status := release_word_application()
    ppt_status := release_ppt_application()
    
    ; 生成释放报告
    release_report := Map(
        "time", current_time,
        "excel", excel_status,
        "word", word_status,
        "powerpoint", ppt_status
    )
    
    last_release_time := current_time
    
    ; 可选：写入日志
    ; log_office_release(release_report)
    
    return release_report
}

; 显示Office对象状态
show_office_objects_status() {
    global excel_application, word_application, ppt_application
    
    message := "=== Office COM对象状态 ===`n"
    
    ; Excel状态
    if (excel_application) {
        try {
            name := excel_application.Name
            version := excel_application.Version
            message .= "Excel: 已连接 (v" . version . ")`n"
        } catch {
            message .= "Excel: 对象失效`n"
        }
    } else {
        message .= "Excel: 未连接`n"
    }
    
    ; Word状态
    if (word_application) {
        try {
            name := word_application.Name
            version := word_application.Version
            message .= "Word: 已连接 (v" . version . ")`n"
        } catch {
            message .= "Word: 对象失效`n"
        }
    } else {
        message .= "Word: 未连接`n"
    }
    
    ; PowerPoint状态
    if (ppt_application) {
        try {
            name := ppt_application.Name
            version := ppt_application.Version
            message .= "PowerPoint: 已连接 (v" . version . ")`n"
        } catch {
            message .= "PowerPoint: 对象失效`n"
        }
    } else {
        message .= "PowerPoint: 未连接`n"
    }
    
    ; 进程状态
    message .= "`n=== 进程状态 ===`n"
    message .= "Excel进程: " . (WinExist("ahk_exe EXCEL.EXE") ? "运行中" : "未运行") . "`n"
    message .= "Word进程: " . (WinExist("ahk_exe WINWORD.EXE") ? "运行中" : "未运行") . "`n"
    message .= "PowerPoint进程: " . (WinExist("ahk_exe POWERPNT.EXE") ? "运行中" : "未运行") . "`n"
    
    MsgBox(message, "Office对象状态")
}

; 检查Office进程详情
check_office_processes() {
    office_processes := ["EXCEL.EXE", "WINWORD.EXE", "POWERPNT.EXE"]
    office_names := ["Excel", "Word", "PowerPoint"]
    
    message := "=== Office进程详情 ===`n"
    
    Loop office_processes.Length {
        process_name := office_processes[A_Index]
        app_name := office_names[A_Index]
        
        found_processes := []
        for process in ComObjGet("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" . process_name . "'") {
            found_processes.Push({
                PID: process.ProcessId,
                Memory: Round(process.WorkingSetSize/1024/1024),
                StartTime: process.CreationDate
            })
        }
        
        if (found_processes.Length == 0) {
            message .= app_name . ": 无进程`n"
        } else {
            message .= app_name . ": " . found_processes.Length . "个进程`n"
            for proc in found_processes {
                message .= "  PID:" . proc.PID . " 内存:" . proc.Memory . "MB`n"
            }
        }
    }
    
    MsgBox(message, "Office进程状态")
}

; 诊断并修复Office对象问题
diagnose_and_fix_office() {
    global excel_application, word_application, ppt_application
    
    message := "=== Office诊断报告 ===`n"
    fixed_count := 0
    
    ; 检查并修复Excel
    if (excel_application) {
        try {
            name := excel_application.Name
            message .= "✓ Excel COM对象正常`n"
        } catch {
            message .= "✗ Excel COM对象失效，已清理`n"
            excel_application := ""
            fixed_count++
        }
    }
    
    ; 检查并修复Word
    if (word_application) {
        try {
            name := word_application.Name
            message .= "✓ Word COM对象正常`n"
        } catch {
            message .= "✗ Word COM对象失效，已清理`n"
            word_application := ""
            fixed_count++
        }
    }
    
    ; 检查并修复PowerPoint
    if (ppt_application) {
        try {
            name := ppt_application.Name
            message .= "✓ PowerPoint COM对象正常`n"
        } catch {
            message .= "✗ PowerPoint COM对象失效，已清理`n"
            ppt_application := ""
            fixed_count++
        }
    }
    
    message .= "`n修复了 " . fixed_count . " 个失效对象"
    
    MsgBox(message, "Office诊断修复")
    
    if (fixed_count > 0) {
        ToolTip("已清理 " . fixed_count . " 个失效的Office对象")
        SetTimer(() => ToolTip(), -2000)
    }
}

; 启动Office对象定时清理
start_office_cleanup_timer() {
    SetTimer(release_all_group_officelications, 120000)  ; 每2分钟
    ToolTip("Office对象定时清理已启动（每2分钟）")
    SetTimer(() => ToolTip(), -2000)
}

; 停止Office对象定时清理
stop_office_cleanup_timer() {
    SetTimer(release_all_group_officelications, 0)
    ToolTip("Office对象定时清理已停止")
    SetTimer(() => ToolTip(), -2000)
}

; 手动释放所有Office对象
manual_release_all_office() {
    report := release_all_group_officelications()
    
    message := "Office对象释放完成:`n"
    message .= "Excel: " . report["excel"] . "`n"
    message .= "Word: " . report["word"] . "`n"
    message .= "PowerPoint: " . report["powerpoint"] . "`n"
    message .= "时间: " . report["time"]
    
    ToolTip(message)
    SetTimer(() => ToolTip(), -3000)
}

; Office启动前清理
clean_before_office_start() {
    release_all_group_officelications()
    Sleep(500)  ; 给系统处理时间
    ToolTip("Office COM对象已清理，可以启动应用程序")
    SetTimer(() => ToolTip(), -2000)
}

; 脚本退出时清理所有Office对象
cleanup_all_office_on_exit() {
    stop_office_cleanup_timer()
    release_all_group_officelications()
}

; 自动启动定时清理
auto_start_office_cleanup() {
    start_office_cleanup_timer()
}

; ; 快捷键绑定
; Ctrl+F1:: show_office_objects_status()     ; 显示Office对象状态
; Ctrl+F2:: check_office_processes()         ; 检查Office进程
; Ctrl+F3:: diagnose_and_fix_office()        ; 诊断并修复Office对象
; Ctrl+F4:: manual_release_all_office()      ; 手动释放所有Office对象
; Ctrl+F5:: start_office_cleanup_timer()     ; 启动定时清理
; Ctrl+F6:: stop_office_cleanup_timer()      ; 停止定时清理
; Ctrl+F7:: clean_before_office_start()      ; Office启动前清理

; ; 自动启动定时清理（脚本加载时）
; auto_start_office_cleanup()

; ; 脚本退出时清理
; OnExit(cleanup_all_office_on_exit)
; #endregion




