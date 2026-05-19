; #Include 必要的库文件
; #Include lib_public.ahk

; 获取文件列表的通用函数 
get_selected_file_paths() {    
        
    try { 
        local file_list := [] 
    
        ; 检查当前窗口类型 
        if WinActive("ahk_class CabinetWClass") { 
            ; 从资源管理器获取文件路径（使用 Ctrl+Shift+C）
            Send("^+c")  ; 发送 Ctrl+Shift+C 获取路径 
            Sleep(100)   ; 短暂等待剪贴板更新   
        }
        if WinActive("ahk_class TTOTAL_CMD") { 
            ; 从 Total Commander 获取文件路径（使用 Alt+M+P）
            Send("!mp")  ; 发送 Alt+M+P 获取路径  
            Sleep(100)   ; 短暂等待剪贴板更新    
        }

        clipboard_content := A_Clipboard 
        clipboard_content := StrReplace(clipboard_content, "`"")
        if (clipboard_content = "") { 
            Msgbox("请在资源管理器（使用 Ctrl+Shift+C）或 Total Commander（使用 Alt+M+P）窗口中运行此功能！") 
            return [] 
        }
        
        file_list := StrSplit(clipboard_content, "`n", "`r") 
        
        ; 过滤掉空行并清理路径 
        cleaned_list := [] 
        for index, file_path in file_list { 
            clean_path := Trim(file_path, "`r`n`t ") 
            if (clean_path != "") { 
                cleaned_list.Push(clean_path) 
            } 
        } 
        return cleaned_list 
    } catch as err { 
        A_Clipboard := err.Message 
        Msgbox("无法获取文件列表：`n" "错误: " err.Message) 
        return  [] 
    } 
}

; 导出PPT为图片的函数
export_ppt_to_jpg(source_full := "", skip_section := "原稿") {  
    
    skip_index := 0 
    first_skip := 0 
    end_skip := 0 
    section_count := 0   

    if !source_full { 
        ; 从上下文获取文件列表（资源管理器、Total Commander或剪贴板）
        files_array := get_selected_file_paths()
        
        if (files_array.Length = 0) {
            ; 如果上下文获取失败，尝试从选中文件获取
            selected_files_str := get_selected_files(2)  ; 2 表示仅允许选择文件
            if (selected_files_str = "") {
                Msgbox("请先选择要处理的 PowerPoint 文件！")
                return
            }
            files_array := StrSplit(selected_files_str, "`n", "`r")
        }
        
        ; 只取第一个文件 
        source_full := Trim(files_array[1]) 
        if (!FileExist(source_full)) { 
            Msgbox("文件不存在：`n" source_full)
            return 
        } 
    } 
    
    ; 获取文件路径信息 
    source_full     := path_info(source_full).full      ; 文件完整路径 
    source_drive    := path_info(source_full).drive     ; 驱动器 (如 C:) 
    source_dir      := path_info(source_full).dir       ; 目录（不含驱动器） 
    source_folder   := path_info(source_full).folder    ; 文件夹（含驱动器）        
    source_fname    := path_info(source_full).fname     ; 文件名（不含后缀） 
    source_ext      := path_info(source_full).ext       ; 文件后缀   
    source_file     := path_info(source_full).file      ; 文件名（含后缀） 

    if (RegExMatch(source_fname, "^(.*)_完整$", &match)) { 
        base_name := match[1] 
    } else { 
        base_name := source_fname  ; 如果不是以"_完整"结尾，使用原文件名 
    } 

    ; 使用提取的基础名称创建目标文件夹 
    target_folder := source_folder . "\" . base_name . "_图片" 

    ; 创建目标文件夹（如果不存在） 
    if (!FileExist(target_folder)) { 
        try { 
            DirCreate(target_folder) 
        } catch as err { 
            Msgbox("无法创建文件夹：`n" target_folder "`n错误: " err.Message) 
            return 
        } 
    }   
    
    ; run(target_folder)   
        
    ; 创建 PowerPoint 应用程序对象 
    ppt_app := ComObject("PowerPoint.Application") 
    ; ppt_app.Visible := false ; 隐藏 PowerPoint 窗口 

    ; 打开源文件 
    try { 
        source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0) ; 只读打开，不显示窗口 
    } catch as err { 
        A_Clipboard := err.Message 
        Msgbox("无法打开文件：`n" source_full "`n错误: " err.Message) 
        return 
    } 

    ; 检查源文件是否有需要跳过的节
    if (skip_section != -1) {
        try {
            section_props := source_pres.SectionProperties
            section_count := section_props.Count
            
            ; 查找需要跳过的节
            skip_index := 0
            loop section_count {
                section_name := section_props.Name(A_Index)
                if (section_name = skip_section) {
                    skip_index := A_Index
                    break ; 只处理第一个匹配的节
                }
            }
            
            ; 如果找到需要跳过的节，计算节的起始和结束幻灯片
            if (skip_index > 0) {
                first_skip := section_props.FirstSlide(skip_index)
                skip_count := section_props.SlidesCount(skip_index)
                end_skip := first_skip + skip_count - 1
            }
        } catch as err {
            ; 如果检查节失败，继续处理，不中断整个流程
            Msgbox("检查节时出错：`n" err.Message)
        }
    }

    ;   ; 4. 导出为GIF
    ; target_full := target_folder . "\" . base_name . "_GIF.gif"
    ; source_pres.SaveAs(target_full, 40)
    
    ; 处理每个幻灯片
    target_index := 0
    try {
        ; 遍历每张幻灯片
        for slide in source_pres.Slides {
            page_no := slide.SlideIndex
            target_index += 1
            
            ; 如果当前幻灯片在需要跳过的节中，跳过处理
            if (skip_index > 0 && page_no >= first_skip && page_no <= end_skip) {
                target_index -= 1
                continue
            }
            
            ; 导出图片
            target_full := target_folder "\" base_name "_图片「" target_index "」.jpg"
            
            try {
                slide.Export(target_full, "jpg", 1920, 1080)    ; 输出图片

            } catch as err {
                Msgbox("导出幻灯片失败：`n" source_full "`n错误: " err.Message)
            }
        }
    } catch as err {
        Msgbox("遍历幻灯片失败：`n" source_full "`n错误: " err.Message)
    }
    
    
    ; 关闭源文件 
    try { 
        source_pres.Close() 
    } catch as err { 
        Msgbox("关闭源文件失败：`n" source_full "`n错误: " err.Message) 
    } 
    
    ; 关闭PowerPoint应用程序
    ppt_app.Quit()
    
    Notify.show("Successfully exported to JPG.") 
    return true

} 


; 主函数：拆分PPT为单页，支持大量幻灯片（>100页）的高效处理
split_ppt_to_single_main(source_full := "", batch_size := 30, export_jpg := false, skip_section := "原稿") {
    try {
        ; 初始化文件列表
        files_array := []
        
        ; 获取文件列表的方式
        if (source_full = "") {
            ; 1. 从上下文获取文件列表（资源管理器、Total Commander或剪贴板）
            files_array := get_selected_file_paths()
            
            if (files_array.Length = 0) {
                ; 2. 如果上下文获取失败，尝试从选中文件获取
                selected_files_str := get_selected_files(2)
                if (selected_files_str = "") {
                    Msgbox("未找到任何文件，无法进行PPT拆分")
                    return false
                }
                files_array := StrSplit(selected_files_str, "`n", "`r")
            }
        } else {
            ; 直接使用传入的文件路径
            files_array := [source_full]
        }
        
        ; 过滤并处理PPTX文件
        pptx_files := []
        for file_path in files_array {
            file_path := Trim(file_path) ; 清理路径中的空格
            file_path := Trim(file_path, "`"") ; 清理路径中的双引号
            if (file_path = "" || !FileExist(file_path)) {
                continue ; 跳过空路径或不存在的文件
            }
            
            ; 检查是否为PPTX文件
            ext := path_info(file_path).Ext
            if (ext != ".pptx" && ext != ".ppt") {
                continue ; 跳过非PPT文件
            }
            
            pptx_files.Push(file_path)
        }
        
        if (pptx_files.Length = 0) {
            Msgbox("未找到任何PPTX文件，无法进行拆分")
            return false
        }
        
        ; 处理每个PPTX文件
        ; Notify.show("开始处理 " . pptx_files.Length . " 个PPT文件")
        
        success_count := 0
        for index, file_path in pptx_files {
            ; Notify.show("正在处理文件 " . (index + 1) . "/" . pptx_files.Length . ": " . file_path)
            
            ; 处理单个文件
            if (process_single_ppt(file_path, batch_size)) {
                success_count++
                
                ; 如果需要，导出为图片
                if (export_jpg) {
                    Notify.show("正在导出图片: " . file_path)
                    export_ppt_to_jpg(file_path, skip_section)
                }
            } else {
                Msgbox("处理文件失败: " . file_path)
            }
        }
        
        Notify.show("PPT处理完成！成功: " . success_count . "/" . pptx_files.Length)
        return success_count > 0
        
    } catch as err {
        Msgbox("拆分过程出错: " . err.message)
        MsgBox("拆分过程出错: " . err.message)
        return false
    }
}

; 处理单个PPT文件的辅助函数 - 使用先分段再分单页的模式
process_single_ppt(source_full, batch_size := 30) {
    try {
        ; 参数验证
        if (!FileExist(source_full)) {
            Msgbox("源文件不存在: " . source_full)
            return false
        }
        
        if (batch_size <= 0 || !IsNumber(batch_size)) {
            batch_size := 30 ; 默认每批30页
        }
        
        ; 获取源文件信息
        source_info := path_info(source_full)
        source_folder := source_info.Folder
        source_fname := source_info.Fname
        
        ; 创建单页输出文件夹
        single_folder := source_folder . "\" . source_fname . "_单页"
        if !DirExist(single_folder)
            DirCreate(single_folder)
        
        ; 创建PowerPoint应用程序对象
        ppt_app := ComObject("PowerPoint.Application")
        
        ; 打开源演示文稿获取总页数
        ppt_pres := ppt_app.Presentations.Open(source_full, , , false) ; 以只读模式打开
        total_slides := ppt_pres.Slides.Count
        ppt_pres.Close()
        
        Notify.show("开始拆分，共 " . total_slides . " 张幻灯片，每批 " . batch_size . " 张")
        
        ; 计算批次数
        batch_count := Ceil(total_slides / batch_size)
        
        ; 顺序处理每个批次
        Loop batch_count {
            batch_index := A_Index
            start_slide := (batch_index - 1) * batch_size + 1
            end_slide := Min(batch_index * batch_size, total_slides)
            
            ; Notify.show("开始处理批次 " . batch_index . "/" . batch_count . ": " . start_slide . "-" . end_slide . " 页")
            
            ; 1. 复制源文件并创建当前批次的文件
            batch_file := source_folder . "\" . source_fname . "_临时_" . start_slide . "-" . end_slide . ".pptx"
            FileCopy(source_full, batch_file, true)
            
            ; 2. 打开批次文件并删除不在当前批次范围内的幻灯片
            batch_pres := ppt_app.Presentations.Open(batch_file, , , false)
            
            ; 先删除结束位置之后的幻灯片（从后往前删除）
            i := batch_pres.Slides.Count
            while (i > end_slide) {
                batch_pres.Slides(i).Delete()
                i--
            }
            
            ; 再删除开始位置之前的幻灯片（从后往前删除）
            i := start_slide - 1
            while (i >= 1) {
                batch_pres.Slides(i).Delete()
                i--
            }
            
            batch_pres.Save()
            batch_pres.Close()
            
            ; 3. 将当前批次拆分为单页
            split_batch_to_single(batch_file, start_slide, end_slide, ppt_app, source_info, single_folder)
            
            ; 4. 删除临时批次文件
            FileDelete(batch_file)
        }
        
        ; 关闭PowerPoint应用程序
        ppt_app.Quit()
        
        Notify.show("PPT单页拆分完成！")
        return true
        
    } catch as err {
        Msgbox("处理单个PPT出错: " . err.message)
        ; 确保资源正确释放
        if IsObject(batch_pres)
            batch_pres.Close()
        if IsObject(ppt_pres)
            ppt_pres.Close()
        if IsObject(ppt_app)
            ppt_app.Quit()
        return false
    }
}

; 将单个批次拆分为单页的辅助函数
split_batch_to_single(batch_file, start_slide, end_slide, ppt_app, source_info, output_folder) {
    try {
        ; 顺序处理每个幻灯片
        Loop end_slide - start_slide + 1 {
            slide_pos := A_Index
            actual_slide_num := start_slide + slide_pos - 1
            
            ; 1. 复制批次文件到临时文件
            temp_file := output_folder . "\" . source_info.Fname . "_临时_" . actual_slide_num . ".pptx"
            FileCopy(batch_file, temp_file, true)
            
            ; 2. 打开临时文件
            single_pres := ppt_app.Presentations.Open(temp_file, , , false)
            
            ; 3. 删除不需要的幻灯片（保留当前页）
            ; 先删除当前页之后的幻灯片（从后往前删除）
            i := single_pres.Slides.Count
            while (i > slide_pos) {
                single_pres.Slides(i).Delete()
                i--
            }
            
            ; 再删除当前页之前的幻灯片（从后往前删除）
            i := slide_pos - 1
            while (i >= 1) {
                single_pres.Slides(i).Delete()
                i--
            }
            
            ; 4. 保存为单页文件
            single_filename := output_folder . "\" . source_info.Fname . "_单页「" . actual_slide_num . "」.pptx"
            
            ; 检查文件是否已存在，避免覆盖
            if (FileExist(single_filename)) {
                ; 添加时间戳或序号避免覆盖
                timestamp := FormatTime(, "HHmmss")
                single_filename := output_folder . "\" . source_info.Fname . "_单页「" . actual_slide_num . "_" . timestamp . "」.pptx"
            }
            
            single_pres.SaveAs(single_filename)
            single_pres.Close()
            
            ; 5. 删除临时文件
            FileDelete(temp_file)
            
            Notify.show("已完成单页拆分: " . actual_slide_num)
        }
        
        return true
        
    } catch as err {
        Msgbox("拆分批次为单页出错: " . err.message)
        ; 确保资源释放
        if IsObject(single_pres)
            single_pres.Close()
        return false
    }
}


; 只拆分为单页（不进行批量拆分）的函数 - 保留原函数接口
split_to_single_only(source_full, total_slides := 0, ppt_app := "") {
    return split_ppt_to_single_main(source_full, 1) ; 单批次处理
}

; 优化后的批量单页拆分函数 - 保留原函数接口
split_each_batch_to_single_optimized(batch_folder, source_fname, ppt_app) {
    Msgbox("该函数已废弃，请使用 split_ppt_to_single_main 函数替代")
    return false
}

; 获取实际幻灯片编号的辅助函数 - 保留原函数接口
get_actual_slide_number(batch_file, slide_index_in_batch) {
    ; 从文件名中提取起始幻灯片号
    if (RegExMatch(batch_file, "_多页「(\d+)-(\d+)」", &match)) {
        start_num := match[1]
        return start_num + slide_index_in_batch - 1
    }
    return slide_index_in_batch
}



; PPT管理工具 - 整合拆分、导出图片和文件目录管理功能
; 版本: v2.0
; 日期: 2026-01-13

; 主管理函数 - 通过参数控制三个功能
process_ppt_manager(split_operation := false, export_operation := false, organize_operation := false) {
    local files := []
    
    ; 获取文件列表
    if (split_operation || export_operation || organize_operation) {
        files := get_selected_file_paths()
        if (files.Length = 0) {
            Msgbox("未找到任何文件，请确保已正确选择文件！")
            return false
        }
    }
    
    ; 检查是否至少选择了一个操作
    if (!split_operation && !export_operation && !organize_operation) {
        Msgbox("请至少选择一个操作：拆分、导出图片或文件管理！")
        return false
    }
    
    ; 处理所有文件
    success_count := 0
    for index, file_path in files {
        source_full := Trim(file_path, "`s`"")        
        if (!FileExist(source_full)) {
            Msgbox("文件不存在：`n" . source_full)
            continue  ; 跳过此文件，继续处理下一个
        }
               
        ; 获取文件路径信息
        source_info := path_info(source_full)
        source_full := source_info.full      ; 文件完整路径
        source_folder := source_info.folder  ; 文件夹（含驱动器）        
        source_fname := source_info.fname    ; 文件名（不含后缀）
        source_ext := source_info.ext        ; 文件后缀   
        source_file := source_info.file      ; 文件名（含后缀）
        
        ; 检查文件类型
        if (source_ext != ".pptx" && source_ext != ".ppt") {
            Msgbox("跳过非PPT文件：`n" . source_full)
            continue
        }
        
        try {
            Notify.show("正在处理文件 " . (index) . "/" . files.Length . ": " . source_fname)
            
            ; 根据参数执行相应的操作
            if (split_operation) {  ; PPT拆分
                if (split_ppt_to_single_main(source_full)) {
                    Notify.show("PPT拆分完成：" . source_fname)
                } else {
                    Msgbox("PPT拆分失败：" . source_fname)
                    ; continue
                }
            }
            
            if (export_operation) {  ; 图片导出
                            
        ;    MsgBox "继续处理下一个文件？"
                if (export_ppt_to_jpg(source_full)) {        
                    Notify.show("图片导出完成：" . source_fname)
                } else {
                    Msgbox("图片导出失败：" . source_fname)
                    MsgBox("图片导出失败：" . source_fname)
                    ; continue
                }
            }
            
            ; 添加延迟确保PowerPoint完全关闭
            Sleep(1000)
            
            if (organize_operation) {  ; 文件目录管理
                if (organize_ppt_files(source_full)) {
                    Notify.show("文件管理完成：" . source_fname)
                } else {
                    Msgbox("文件管理失败：" . source_fname)
                    MsgBox("文件管理失败：" . source_fname)
                    continue
                }
            }
            
            success_count++
            
        } catch as err {
            Msgbox("处理文件时出现错误：" . source_fname . "`n" . err.Message)
            MsgBox("处理文件时出现错误：" . source_fname . "`n" . err.Message)
            continue  ; 跳过此文件，继续处理下一个
        }
    }
    
    Notify.show("所有操作完成！成功：" . success_count . "/" . files.Length)
    return success_count > 0
}

; 文件目录管理辅助函数
organize_ppt_files(source_full) {
    ; 获取文件路径信息
    source_info := path_info(source_full)
    source_full := source_info.full      ; 文件完整路径
    source_folder := source_info.folder  ; 文件夹（含驱动器）        
    source_fname := source_info.fname    ; 文件名（不含后缀）
    source_file := source_info.file      ; 文件名（含后缀）
    

    ;  获取祖父目录
    source_grandparent_folder := RegExReplace(source_full, "\\[^\\]+\\[^\\]+$", "") . "\"
    
    target_folder := source_grandparent_folder . source_fname . "\"
    
    ; 移动整个文件夹到以PPT文件名命名的新文件夹
    try {
        DirMove(source_folder, target_folder, "R")
    } catch as err {
        Msgbox("无法移动文件夹：`n" . source_folder . "`n错误: " . err.Message)
        return false  ; 跳过此文件，继续处理下一个
    }
    
    ; 更新路径变量
    new_source_full := target_folder . source_file
    new_source_folder := path_info(new_source_full).folder
    
    source_font_full := new_source_folder . "字体"
    target_font_full := new_source_folder . source_fname . "_字体"
    target_wz_fold := new_source_folder . source_fname . "_完整"
    target_wz_full := new_source_folder . source_fname . "_完整\" . source_fname . "_完整.pptx"

    ; 移动完整文件
    try {
        if (!FileExist(target_wz_fold)) {            
            DirCreate(target_wz_fold)
        }   
        FileMove(new_source_full, target_wz_full)
    } catch as err {
        Msgbox("无法移动完整文件：`n" . new_source_full . "`n错误: " . err.Message)
        MsgBox("无法移动完整文件：`n" . new_source_full . "`n错误: " . err.Message)

        return false  ; 跳过此文件，继续处理下一个
    }
    
    ; 移动字体文件夹（如果存在）
    if DirExist(source_font_full) {
        try {
            DirMove(source_font_full, target_font_full)
        } catch as err {
            Msgbox("无法移动字体文件夹：`n" . target_font_full . "`n错误: " . err.Message)
            MsgBox("无法移动字体文件夹：`n" . target_font_full . "`n错误: " . err.Message)
            return false  ; 跳过此文件，继续处理下一个
        } 
    }
    Notify.show("文件整理完成：" . source_fname)
    return true
    
}

; 为常见的操作组合创建便捷函数

; 仅PPT拆分
process_ppt_split_only() {
    return process_ppt_manager(true, false, false)
}

; 仅图片导出
process_ppt_export_only() {
    return process_ppt_manager(false, true, false)
}

; 仅文件管理
process_ppt_organize_only() {
    return process_ppt_manager(false, false, true)
}

; PPT拆分 + 图片导出
process_ppt_split_and_export() {
    return process_ppt_manager(true, true, false)
}

; PPT拆分 + 文件管理
process_ppt_split_and_organize() {
    return process_ppt_manager(true, false, true)
}

; 图片导出 + 文件管理
process_ppt_export_and_organize() {
    return process_ppt_manager(false, true, true)
}

; 全部操作
process_ppt_full_process() {
    return process_ppt_manager(true, true, true)
}





; 导出PPT为动图片的函数
save_ppt_as_gif2(source_full := "") {  
    if !source_full { 
        ; 从上下文获取文件列表（资源管理器、Total Commander）
        files_array := get_selected_file_paths()
        if (files_array.Length == 0) {
            Msgbox("请先选择要处理的 PowerPoint 文件！")
            return
        }
    } 

    for index, file_path in files_array {
        source_full := Trim(file_path, "`s`"")        
        if (!FileExist(source_full)) {
            Msgbox("文件不存在：`n" . source_full)
            continue  ; 跳过此文件，继续处理下一个
        }

        if InStr(source_full, "单页") {
            target_full := StrReplace(source_full, "单页", "动图")
            target_full := StrReplace(target_full, "pptx", "gif")
        } else {
            target_full := path_info(source_full, "fs:_动图", "Ex:.gif").Full
        }
        ; MsgBox("正在处理文件：`n" . target_full)
        target_folder := path_info(target_full).folder 
      
        ; 创建目标文件夹（如果不存在） 
        if (!FileExist(target_folder)) { 
            DirCreate(target_folder) 
        }   

        ; 创建 PowerPoint 应用程序对象 
        ppt_app := ComObject("PowerPoint.Application") 
        ; ppt_app.Visible := false ; 隐藏 PowerPoint 窗口 
        pres := ppt_app.Presentations.Open(source_full, 1, 0, 0) ; 只读打开，不显示窗口 
        pres.SaveAs(target_full, 40)
        pres.Close()    
       
    }
    ; 关闭PowerPoint应用程序  
    ppt_app.Quit() 
    MsgBox("导出为动图完成！")

} 

; 导出PPT为动图片的函数（优化版）
save_ppt_as_gif(source_full := "") {  
    try {
        ; 初始化文件数组
        files_array := []
        
        if !source_full { 
            ; 从上下文获取文件列表（资源管理器、Total Commander）
            files_array := get_selected_file_paths()
            if (files_array.Length == 0) {
                Msgbox("请先选择要处理的 PowerPoint 文件！")
                return false
            }
        } else {
            ; 单个文件处理
            files_array := [source_full]
        }

        processed_count := 0
        failed_count := 0
        
        ; 创建 PowerPoint 应用程序对象
        ppt_app := ComObject("PowerPoint.Application")
        ; ppt_app.Visible := false  ; 隐藏 PowerPoint 窗口
        
        for index, file_path in files_array {
            try {
            source_full := Trim(file_path, "`s`"")          
                if (!FileExist(source_full)) {
                    Msgbox("文件不存在：`n" . source_full)
                    failed_count++
                    continue  ; 跳过此文件，继续处理下一个
                }

                ; 验证文件扩展名
                source_ext := path_info(source_full).ext
                if (source_ext != ".pptx" && source_ext != ".ppt") {
                    Msgbox("跳过非PPT文件：`n" . source_full)
                    failed_count++
                    continue
                }

                ; 生成目标文件路径
                if InStr(source_full, "单页") {
                    target_full := StrReplace(source_full, "单页", "动图")
                    target_full := StrReplace(target_full, source_ext, ".gif")
                } else {
                    target_full := path_info(source_full, "fs:_动图", "Ex:.gif").Full
                }
                
                target_folder := path_info(target_full).folder 
              
                ; 创建目标文件夹（如果不存在） 
                if (!FileExist(target_folder)) { 
                    try {
                        DirCreate(target_folder) 
                    } catch as err {
                        Msgbox("无法创建目标文件夹：`n" . target_folder . "`n错误: " . err.Message)
                        failed_count++
                        continue
                    }
                }   

                Notify.show("正在导出 GIF: " . path_info(source_full).fname)

                ; 打开并处理演示文稿
                pres := ""
                try {
                    pres := ppt_app.Presentations.Open(source_full, 1, 0, 0) ; 只读打开，不显示窗口 
                    pres.SaveAs(target_full, 40)  ; 40 对应 GIF 格式
                } catch as err {
                    Msgbox("导出 GIF 失败：`n" . source_full . "`n错误: " . err.Message)
                    failed_count++
                    if IsObject(pres) {
                        pres.Close()
                    }
                    continue
                }
                
                ; 关闭演示文稿
                if IsObject(pres) {
                    pres.Close()
                }
                
                processed_count++
                Notify.show("成功导出 GIF: " . path_info(target_full).file)
                
            } catch as err {
                Msgbox("处理文件时发生错误: " . file_path . "`n错误: " . err.Message)
                failed_count++
                continue
            }
        }

        ; 关闭PowerPoint应用程序  
        if IsObject(ppt_app) {
            ppt_app.Quit() 
        }
        
        Notify.show("GIF 导出完成！成功: " . processed_count . ", 失败: " . failed_count)
        return (failed_count = 0)  ; 如果没有失败则返回 true
        
    } catch as err {
        Msgbox("导出过程中发生严重错误: " . err.Message)
        ; 确保 PowerPoint 应用程序被关闭
        try {
            if IsObject(ppt_app) {
                ppt_app.Quit()
            }
        } catch {
            ; 忽略关闭时的错误
        }
        return false
    }
}



; 导出PPT为动图片的函数（优化版）
batch_rename_files(source_full := "") {  
    try {
        ; 初始化文件数组
        files_array := []
        
        if !source_full { 
            ; 从上下文获取文件列表（资源管理器、Total Commander）
            files_array := get_selected_file_paths()
            if (files_array.Length == 0) {
                Msgbox("请先选择要处理的 PowerPoint 文件！")
                return false
            }
        } else {
            ; 单个文件处理
            files_array := [source_full]
        }

        processed_count := 0
        failed_count := 0
      
        
        for index, file_path in files_array {
            try {
            source_full := Trim(file_path, "`s`"")          
                if (!FileExist(source_full)) {
                    Msgbox("文件不存在：`n" . source_full)
                    failed_count++
                    continue  ; 跳过此文件，继续处理下一个
                }

                ; 验证文件扩展名
                source_ext := path_info(source_full).ext
                if (source_ext != ".pptx" && source_ext != ".ppt") {
                    Msgbox("跳过非PPT文件：`n" . source_full)
                    failed_count++
                    continue
                }
            SplitPath source_full, , &dir
            SplitPath dir, &folderName
            target_fname := "230101_" . SubStr(folderName, 4) . "_2023Raven作品"              
            target_full := path_info(source_full, "fname:" . target_fname).Full    
            FileMove(source_full, target_full)
            } catch as err {
                Msgbox("处理文件时发生错误: " . file_path . "`n错误: " . err.Message)
                failed_count++
                continue
            }
        }
    } catch as err {
        Msgbox("导出过程中发生严重错误: " . err.Message)       
        return false
    }
}


test_path_info2() { 

SplitPath "C:\My Documents\Address List.txt", &name, &dir, &ext, &name_no_ext, &drive

SplitPath dir, &folderName 

MsgBox(folderName)

}