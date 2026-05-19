; ============================================================
;  lib_ppt.ahk — PPT 拆分 / 导出JPG / 整理文件夹
; ============================================================

process_powerpoint() {
    get_grandparent_dir(path) {
        grandparent := RegExReplace(path, "\\[^\\]+\\[^\\]+$", "")
        return grandparent . "\"
    }
    files := get_selected_files(2)
    if (files = "") {
        Msgbox("请先选择要拆分的 PowerPoint 文件！")
        return
    }
    source_full := Trim(StrSplit(files, "`n", "`r")[1])
    if (!FileExist(source_full)) {
        Msgbox("文件不存在：`n" source_full)
        return
    }
    source_full := path_info(source_full).full
    source_folder := path_info(source_full).folder
    source_fname := path_info(source_full).fname
    source_file := path_info(source_full).file
    source_grandparent_folder := get_grandparent_dir(source_full)
    target_folder := source_grandparent_folder . source_fname . "\"
    try {
        DirMove(source_folder, target_folder, "R")
    } catch as err {
        Msgbox("无法移动完整文件：`n" source_full "`nError: " err.Message)
        return
    }
    source_full := target_folder . source_file
    source_full := path_info(source_full).full
    source_folder := path_info(source_full).folder
    source_fname := path_info(source_full).fname
    source_file := path_info(source_full).file
    source_font_full := source_folder . "字体"
    target_font_full := source_folder . source_fname . "_字体"
    target_wz_fold := source_folder . source_fname . "_完整"
    target_wz_full := source_folder . source_fname . "_完整\" . source_fname . "_完整.pptx"
    split_in_folder(source_full)
    export_jpg_in_folder(source_full)
    try {
        if (!FileExist(target_wz_fold))
            DirCreate(target_wz_fold)
        FileMove(source_full, target_wz_full)
    } catch as err {
        Msgbox("无法移动完整文件：`n" source_full "`nError: " err.Message)
        return
    }
    if DirExist(source_font_full) {
        try DirMove(source_font_full, target_font_full)
        catch as err {
            Msgbox("无法移动font文件夹：`n" target_font_full "`nError: " err.Message)
            return
        }
    }
    Notify.show("Successfully!")
}

split_in_folder(source_full, skip_section := "原稿") {
    skip_index := 0
    first_skip := 0
    end_skip := 0
    section_count := 0
    if !source_full {
        files := get_selected_files(2)
        if (files = "") {
            Msgbox("请先选择要拆分的 PowerPoint 文件！")
            return
        }
        source_full := Trim(StrSplit(files, "`n", "`r")[1])
        if (!FileExist(source_full)) {
            Msgbox("文件不存在：`n" source_full)
            return
        }
    }
    pi := path_info(source_full)
    source_full := pi.full
    source_folder := pi.folder
    source_fname := pi.fname
    source_file := pi.file
    if (RegExMatch(source_fname, "^(.*)_完整$", &match))
        base_name := match[1]
    else
        base_name := source_fname
    target_folder := source_folder . "\" . base_name . "_单页"
    if (!FileExist(target_folder)) {
        try DirCreate(target_folder)
        catch as err {
            Msgbox("无法创建文件夹：`n" target_folder "`nError: " err.Message)
            return
        }
    }
    ppt_app := ComObject("PowerPoint.Application")
    try {
        source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0)
    } catch as err {
        Msgbox("无法打开文件：`n" source_full "`nError: " err.Message)
        return
    }
    if (skip_section != -1) {
        try {
            section_props := source_pres.SectionProperties
            section_count := section_props.Count
            loop section_count {
                if (section_props.Name(A_Index) = skip_section) {
                    skip_index := A_Index
                    break
                }
            }
            if (skip_index > 0) {
                first_skip := source_pres.SectionProperties.FirstSlide(skip_index)
                skip_count := source_pres.SectionProperties.SlidesCount(skip_index)
                end_skip := first_skip + skip_count - 1
            }
        } catch as err {
            Msgbox("检查原稿节时出错：`n" err.Message)
        }
    }
    target_index := 0
    for slide in source_pres.Slides {
        page_no := slide.SlideIndex
        target_index += 1
        if (skip_index > 0 && page_no >= first_skip && page_no <= end_skip) {
            target_index -= 1
            continue
        }
        target_full := target_folder "\" base_name "_单页「" target_index "」.pptx"
        try FileCopy(source_full, target_full, true)
        catch as err {
            Msgbox("文件复制失败：`n" source_full "`nError: " err.Message)
            continue
        }
        try {
            target_pres := ppt_app.Presentations.Open(target_full, false, false, false)
        } catch as err {
            Msgbox("无法打开文件：`n" target_full "`nError: " err.Message)
            continue
        }
        try {
            count := target_pres.Slides.Count
            loop count {
                j := count + 1 - A_Index
                if (j != page_no)
                    target_pres.Slides(j).Delete
            }
        } catch as err {
            Msgbox("删除幻灯片失败：`n" target_full "`nError: " err.Message)
            target_pres.Close()
            continue
        }
        try {
            while target_pres.SectionProperties.Count
                target_pres.SectionProperties.Delete(target_pres.SectionProperties.Count, false)
        } catch {            
        }
        try {
            target_pres.Save()
            target_pres.Close()
        } catch as err {
            Msgbox("保存或关闭文件失败：`n" target_full "`nError: " err.Message)
            continue
        }
    }
    Notify.show("PowerPoint 文件已成功拆分为单页！")
    try source_pres.Close()
    catch {        
    }
    try {
        if IsObject(ppt_app) {
            ppt_app.Quit()
            Sleep 2000
            if ProcessExist("POWERPNT.EXE")
                ProcessClose("POWERPNT.EXE")
        }
    } catch {
        try {
            if ProcessExist("POWERPNT.EXE")
                ProcessClose("POWERPNT.EXE")
        }
    }
}

export_jpg_in_folder(source_full, skip_section := "原稿") {
    skip_index := 0
    first_skip := 0
    end_skip := 0
    section_count := 0
    if !source_full {
        files := get_selected_files(2)
        if (files = "") {
            Msgbox("请先选择要拆分的 PowerPoint 文件！")
            return
        }
        source_full := Trim(StrSplit(files, "`n", "`r")[1])
        if (!FileExist(source_full)) {
            Msgbox("文件不存在：`n" source_full)
            return
        }
    }
    pi := path_info(source_full)
    source_full := pi.full
    source_folder := pi.folder
    source_fname := pi.fname
    if (RegExMatch(source_fname, "^(.*)_完整$", &match))
        base_name := match[1]
    else
        base_name := source_fname
    target_folder := source_folder . "\" . base_name . "_图片"
    if (!FileExist(target_folder)) {
        try DirCreate(target_folder)
        catch as err {
            Msgbox("无法创建文件夹：`n" target_folder "`nError: " err.Message)
            return
        }
    }
    ppt_app := ComObject("PowerPoint.Application")
    try {
        source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0)
    } catch as err {
        A_Clipboard := err.Message
        Msgbox("无法打开文件：`n" source_full "`nError: " err.Message)
        return
    }
    if (skip_section != -1) {
        try {
            section_props := source_pres.SectionProperties
            section_count := section_props.Count
            loop section_count {
                if (section_props.Name(A_Index) = skip_section) {
                    skip_index := A_Index
                    break
                }
            }
            if (skip_index > 0) {
                first_skip := source_pres.SectionProperties.FirstSlide(skip_index)
                skip_count := source_pres.SectionProperties.SlidesCount(skip_index)
                end_skip := first_skip + skip_count - 1
            }
        } catch as err {
            Msgbox("检查原稿节时出错：`n" err.Message)
        }
    }
    target_index := 0
    try {
        for slide in source_pres.Slides {
            page_no := slide.SlideIndex
            target_index += 1
            if (skip_index > 0 && page_no >= first_skip && page_no <= end_skip) {
                target_index -= 1
                continue
            }
            target_full := target_folder "\" base_name "_图片「" target_index "」.jpg"
            try slide.Export target_full, "jpg", 1920, 1080
            catch as err {
                Msgbox("导出幻灯片失败：`nError: " err.Message)
                continue
            }
        }
    } catch as err {
        Msgbox("遍历幻灯片失败：`nError: " err.Message)
    }
    try source_pres.Close()
    catch {        
    }
    Notify.show("Successfully exported to JPG.")
}

; =================== 批次拆分 ===================

split_selected_ppt_to_batches_then_single() {
    files := get_selected_files(2)
    if (files = "") {
        Msgbox("请先选择要拆分的 PowerPoint 文件！")
        return
    }
    file_list := StrSplit(files, "`n", "`r")
    for _, file_path in file_list {
        source_full := Trim(file_path)
        if (!FileExist(source_full))
            continue
        split_ppt_to_batches_then_single_core(source_full, 10)
    }
    Notify.show("PPT 分割完成！")
}

split_ppt_to_batches_then_single_core(source_full, pages_per_batch := 10) {
    pi := path_info(source_full)
    source_fname := pi.fname
    source_folder := pi.folder
    try {
        ppt_app := ComObject("PowerPoint.Application")
        ppt_app.Visible := True
    } catch {
        Msgbox("无法启动 PowerPoint 应用程序！")
        return
    }
    try {
        source_pres := ppt_app.Presentations.Open(source_full, , , 0)
        total_slides := source_pres.Slides.Count
    } catch as err {
        Msgbox("无法打开源文件：`n" source_full "`nError: " err.Message)
        try ppt_app.Quit()
        return
    }
    start_page := 1
    while (start_page <= total_slides) {
        end_page := Min(start_page + pages_per_batch - 1, total_slides)
        batch_filename := source_folder . "\" . source_fname . "_多页「" . start_page . "-" . end_page . "」.pptx"
        FileCopy(source_full, batch_filename, true)
        try {
            batch_pres := ppt_app.Presentations.Open(batch_filename, , , 0)
            i := batch_pres.Slides.Count
            while (i >= 1) {
                slide_index := batch_pres.Slides(i).SlideIndex
                if (slide_index < start_page || slide_index > end_page)
                    batch_pres.Slides(i).Delete()
                i--
            }
            try {
                while batch_pres.SectionProperties.Count
                    batch_pres.SectionProperties.Delete(batch_pres.SectionProperties.Count, false)
            }
            batch_pres.Save()
            batch_pres.Close()
        } catch as err {
            Msgbox("处理批次文件时出错：`n" batch_filename "`nError: " err.Message)
        }
        split_batch_to_single_pages(batch_filename, start_page, end_page, ppt_app)
        start_page += pages_per_batch
    }
    try {
        source_pres.Close()
        ppt_app.Quit()
        Sleep(1000)
        if ProcessExist("POWERPNT.EXE")
            ProcessClose("POWERPNT.EXE")
    } catch {
        if ProcessExist("POWERPNT.EXE")
            ProcessClose("POWERPNT.EXE")
    }
}

split_batch_to_single_pages(batch_full, start_num, end_num, ppt_app) {
    pi := path_info(batch_full)
    fname := pi.fname
    folder := pi.folder
    base_name := RegExReplace(fname, "_多页「\d+-\d+」$", "_单页")
    target_folder := folder . "\" . base_name
    if (!FileExist(target_folder))
        DirCreate(target_folder)
    try {
        if (!IsObject(ppt_app) || ppt_app.Name = "")
            ppt_app := ComObject("PowerPoint.Application")
    } catch {
        ppt_app := ComObject("PowerPoint.Application")
    }
    Loop (end_num - start_num + 1) {
        current_slide_index := A_Index
        actual_page_number := start_num + A_Index - 1
        output_file := target_folder . "\" . base_name . "「" . actual_page_number . "」.pptx"
        FileCopy(batch_full, output_file, true)
        try {
            single_pres := ppt_app.Presentations.Open(output_file, , , 0)
            i := single_pres.Slides.Count
            while (i >= 1) {
                if (i != current_slide_index)
                    single_pres.Slides(i).Delete()
                i--
            }
            try {
                while single_pres.SectionProperties.Count
                    single_pres.SectionProperties.Delete(single_pres.SectionProperties.Count, false)
            }
            single_pres.Save()
            single_pres.Close()
        } catch as err {
            Msgbox("处理单页文件时出错：`n" output_file "`nError: " err.Message)
        }
    }
    FileDelete(batch_full)
}

getdate_914(files) {
    try {
        ppt := ComObject("PowerPoint.Application")
        pptPres := ppt.Presentations.Open(files, false, false, false)
        date := pptPres.BuiltInDocumentProperties("Creation Date").Value
        formatted_date := convert_date_format(date)
        pptPres.Close()
        ppt.Quit()
        return(formatted_date)
    } catch as err
        Msgbox(err.message)
}

convert_date_format(date) {
    try {
        if RegExMatch(date, "(\d{4})/(\d{1,2})/(\d{1,2})", &match) {
            year := SubStr(match[1], 3, 2)
            month := Format("{:02}", match[2])
            day := Format("{:02}", match[3])
            return year month day
        } else {
            MsgBox "日期格式不匹配"
            return
        }
    } catch as err
        Msgbox(err.message)
}