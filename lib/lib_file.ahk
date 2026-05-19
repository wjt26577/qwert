; ============================================================
;  lib_file.ahk — 路径解析 / 文件选择 / 重命名 / 压缩解压
; ============================================================

; =================== 路径解析 ===================

PathU(Filename) {
    Local OutFile := Format("{:260}", "")
    DllCall("Kernel32\GetFullPathNameW", "str",Filename, "uint",260, "str",OutFile, "ptr",0)
    DllCall("Shell32\PathYetAnotherMakeUniqueName", "str",OutFile, "str",OutFile, "ptr",0, "ptr",0)
    Return OutFile
}

path_info(Path, X*) {
    Local K,V,N,U, Dr, Di, Fn, Ex
        , FPath := Dr := Di := Fn := Ex := Format("{:260}", "")
    U := Map(), U.Default := "", U.CaseSense := 0
    For K, V in X
         N := StrSplit(V, ":",, 2)
      ,  K := SubStr(N[1], 1, 2)
      ,  U[K] := N[2]
    DllCall("Kernel32\GetFullPathNameW", "str",Trim(Path,Chr(34)), "uint",260, "str",FPath, "ptr",0)
    DllCall("Msvcrt\_wsplitpath", "str", FPath, "str", Dr, "str", Di, "str", Fn, "str", Ex, "cdecl")
    Return {  Drive  :  Dr  :=  U["Dr"] ? U["Dr"] : Dr
           ,  Dir    :  Di  :=  U["dp"] ( U["Di"] ? U["Di"] : Di ) U["ds"]
           ,  Fname  :  Fn  :=  U["fp"] ( U["Fn"]!="" ? U["Fn"] : Fn )  U["fs"]
           ,  Ext    :  Ex  :=  U["*E"]!="" ? ( Ex ? Ex : U["*E"] ) : ( U["Ex"]!="" ? U["Ex"] : Ex )
           ,  Folder :  Dr Di
           ,  File   :  Fn Ex
           ,  Full   :  U["pp"] ( Dr Di Fn Ex ) U["ps"] }
}

; =================== 文件选择 ===================

get_selected_files(mode:="", hwnd:="") {
    Toreturn := ""
    filenum1 := 0
    filenum2 := 0
    Process := WinGetProcessName("A")
    lClass := WinGetClass("A")
    hwnd := WinExist("A")
    if (Process = "explorer.exe") {
        if (lClass ~= "Progman|WorkerW") {
            Files := ListViewGetContent("Selected col1", "SysListView321", "A")
            if (Files = "")
                Toreturn .= A_Desktop
            else {
                filenum1++
                loop Parse, Files, "`n", "`r"
                    Toreturn .= A_Desktop "\" A_LoopField "`n"
            }
        } else if (lClass ~= "(Cabinet|Explore)WClass") {
            for window in ComObject("Shell.Application").Windows {
                if (window.hwnd == hwnd) {
                    pp := window.Document.folder.self.path
                    sel := window.Document.SelectedItems
                    for item in sel {
                        filenum2++
                        Toreturn .= item.path "`r`n"
                    }
                    if Toreturn = ""
                        Toreturn := pp
                }
            }
        }
    }
    fde := Trim(Toreturn, "`r`n")
    if (mode != "") {
        if (filenum1 + filenum2 = 0) {
            if (mode = 0) || (mode = 2)
                return
            else
                return fde
        } else {
            if (mode = 1) or (mode = 2)
                if (filenum1 != 0) {
                    aa := GetSelectedFiles()
                    return aa
                } else
                    return fde
        }
    }
    if InStr(FileExist(fde), "D")
        return RegExReplace(Trim(Toreturn, "`r`n") . "\", "\\\\", "\")
    else if Toreturn != "" {
        lPos := InStr(Toreturn, "\", , -1) - 1
        Toreturn2 := substr(Toreturn, 1, lPos)
        return RegExReplace(Toreturn2 . "\", "\\\\", "\")
    }
}

GetSelectedFiles() {
    originalClipboard := ClipboardAll
    Clipboard := ""
    Send "^c"
    ClipWait 0.5
    clipboardContent := Clipboard
    if (StrSplit(clipboardContent, "`r").MaxIndex() = 1) {
        Clipboard := originalClipboard
        return RegExReplace(clipboardContent, "`r`n", "")
    } else {
        Clipboard := originalClipboard
        return clipboardContent
    }
}

paste_files_inbox(dest_dir := "D:\system\inbox") {
    if WinActive("ahk_exe Weixin.exe") {
        Sleep(500)
        Click("Right")
        Sleep(500)
        SendInput("{Down}{Enter}")
        Sleep(500)
    }
    paste_files_from_clipboard()
    Run(dest_dir)
}

paste_files_from_clipboard(dest_dir := "D:\system\inbox") {
    if !DirExist(dest_dir)
        DirCreate(dest_dir)
    if !DllCall("user32\OpenClipboard", "Ptr", 0) {
        MsgBox("无法打开剪贴板。")
        return false
    }
    if !DllCall("user32\IsClipboardFormatAvailable", "UInt", 15) {
        DllCall("user32\CloseClipboard")
        MsgBox("剪贴板中没有文件。")
        return false
    }
    h_drop := DllCall("user32\GetClipboardData", "UInt", 15, "Ptr")
    if (!h_drop) {
        DllCall("user32\CloseClipboard")
        return false
    }
    p_drop := DllCall("kernel32\GlobalLock", "Ptr", h_drop, "Ptr")
    if (!p_drop) {
        DllCall("user32\CloseClipboard")
        return false
    }
    num_files := DllCall("shell32\DragQueryFileW", "Ptr", p_drop, "UInt", 0xFFFFFFFF, "Ptr", 0, "UInt", 0)
    Loop num_files {
        len := DllCall("shell32\DragQueryFileW", "Ptr", p_drop, "UInt", A_Index - 1, "Ptr", 0, "UInt", 0)
        buf := Buffer((len + 1) * 2, 0)
        DllCall("shell32\DragQueryFileW", "Ptr", p_drop, "UInt", A_Index - 1, "Ptr", buf.Ptr, "UInt", len + 1)
        full_path := StrGet(buf.Ptr, "UTF-16")
        file_name := StrSplit(full_path, "\").Pop()
        dest_path := dest_dir "\" file_name
        FileCopy(full_path, dest_path, 1)
    }
    DllCall("kernel32\GlobalUnlock", "Ptr", h_drop)
    DllCall("user32\CloseClipboard")
    Notify.show("已粘贴 " num_files " 个文件到：" dest_dir)
    return true
}

; =================== 文件操作 ===================

same_folder() {
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("移动操作只能在文件夹状态执行，请打开文件夹再操作！")
        return
    }
    filenames := get_selected_files(2)
    if !filenames {
        Msgbox("没有选中任何文件！")
        return
    }
    Loop Parse, filenames, '`n', '`r' {
        source_full := path_info(A_LoopField).Full
        source_fname := path_info(A_LoopField).Fname
        source_folder := path_info(A_LoopField).Folder
        source_ext := path_info(A_LoopField).Ext
        dest_full := path_info(A_LoopField, "Ext: ").Full
        if !source_ext {
            Msgbox("选中的是文件夹，无法完成操作！")
            return
        }
        if FileExist(dest_full) {
            Msgbox("存在同名文件夹，无法完成操作！")
            return
        }
        try DirCreate(dest_full)
        catch as err
            Msgbox(err.message)
        try FileMove(source_full, dest_full, 0)
        catch as err
            Msgbox(err.message)
    }
}

move_in_one_same_named_folder() {
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("移动操作只能在文件夹状态执行，请打开文件夹再操作！")
        return
    }
    filenames := get_selected_files(2)
    if !filenames {
        Msgbox("没有选中任何文件！")
        return
    }
    Loop Parse, filenames, '`n', '`r' {
        source_full := path_info(A_LoopField).Full
        source_ext := path_info(A_LoopField).Ext
        if !source_ext {
            Msgbox("选中的是文件夹，无法完成操作！")
            return
        }
        if (A_Index == 1) {
            dest_full := path_info(A_LoopField, "Ext: ").Full
            if FileExist(dest_full) {
                Msgbox("存在同名文件夹，无法完成操作！")
                return
            }
            try DirCreate(dest_full)
            catch as err
                Msgbox(err.message)
        }
        try FileMove(source_full, dest_full, 0)
        catch as err
            Msgbox(err.message)
    }
}

move_files_to_parent() {
    try {
        selected_files := get_selected_files(2)
        if !selected_files {
            Msgbox("There's no folder selected.")
            return
        }
        parent_folder := path_info(selected_files).folder
        error_count := MoveFilesAndFolders(selected_files '\*.*', parent_folder, 1)
        DirDelete selected_files
    } catch as err
        Msgbox(err.message)
}

rename_files_to_2parent() {
    try {
        selected_files := get_selected_files(2)
        if !selected_files {
            Msgbox("There's no folder selected.")
            return
        }
        if (InStr(selected_files, "`r`n")) {
            Loop Parse, selected_files, "`n", "`r" {
                parent_folder := path_info(A_LoopField).folder
                parts := StrSplit(A_LoopField, "\")
                lastFolder := parts[-2]
                new_name := parent_folder lastFolder '.pptx'
                sleep(50)
                FileMove A_LoopField, new_name, 1
            }
            return
        }
        parent_folder := path_info(selected_files).folder
        parts := StrSplit(selected_files, "\")
        lastFolder := parts[-2]
        new_name := parent_folder lastFolder '.pptx'
        FileMove selected_files, new_name, 1
    } catch as err
        Msgbox(err.message)
}

MoveFilesAndFolders(SourcePattern, DestinationFolder, DoOverwrite := false) {
    ErrorCount := 0
    if DoOverwrite = 1
        DoOverwrite := 2
    try
        FileMove SourcePattern, DestinationFolder, DoOverwrite
    catch as err
        ErrorCount := Err.Extra
    Loop Files, SourcePattern, "D" {
        try
            DirMove A_LoopFilePath, DestinationFolder "\" A_LoopFileName, DoOverwrite
        catch {
            ErrorCount += 1
            MsgBox "Could not move " A_LoopFilePath " into " DestinationFolder
        }
    }
    return ErrorCount
}

; =================== 日期追加 ===================

append_datetime_to_filename(file_path) {
    source_full := path_info(file_path).Full
    current_datetime := FormatTime(A_Now, "yyyy/MM/dd_HH:mm")
    safe_datetime := StrReplace(StrReplace(current_datetime, ":", "-"), "/", "-")
    new_fname := path_info(file_path).Fname . "【" . safe_datetime . "】"
    return path_info(source_full, "fname:" . new_fname).Full
}

append_datetime() {
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("解压缩操作只能在文件夹窗口中执行，请打开文件夹再操作！")
        return
    }
    files := get_selected_files(2)
    if !files {
        Msgbox("未选择任何文件，无法完成解压缩操作！")
        return
    }
    Loop Parse, files, '`n', '`r' {
        source_full := path_info(A_LoopField).Full
        source_ext := path_info(A_LoopField).Ext
        new_full := append_datetime_to_filename(A_LoopField)
        if source_ext {
            if FileExist(new_full) {
                Msgbox("存在同名文件，重命名操作失败！")
                return
            }
            try FileMove(source_full, new_full, 0)
            catch as err
                Msgbox("重命名操作失败:" . err.message)
        } else {
            if FileExist(new_full) {
                Msgbox("存在同名文件夹，重命名操作失败！")
                return
            }
            try DirMove source_full, new_full, 'R'
            catch as err
                Msgbox(err.message)
        }
    }
}

output_datetime() {
    current_datetime := FormatTime(A_Now, "yyyy/MM/dd_HH:mm")
    safe_datetime := StrReplace(StrReplace(current_datetime, ":", "-"), "/", "-")
    send_by_clipboard("【" . safe_datetime . "】")
}

; =================== 压缩解压 ===================

run_unzip() {
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("解压缩操作只能在文件夹窗口中执行，请打开文件夹再操作！")
        return
    }
    zip_tool := "C:\Program Files\7-Zip\7z.exe"
    files := get_selected_files(2)
    if !files {
        Msgbox("未选择任何文件，无法完成解压缩操作！")
        return
    }
    Loop Parse, files, "`n", "`r" {
        file_full := path_info(A_LoopField).Full
        file_folder := path_info(A_LoopField).Folder
        if !is_compressed(file_full) {
            Msgbox("选中的不是压缩文件，无法完成解压缩操作！")
            return
        }
        cmd := Format('"{1}" x "{2}" -o"{3}"\* -aoa', zip_tool, file_full, file_folder)
        try Run(cmd)
        catch as err {
            Msgbox("解压缩失败：" . err.message)
            return
        }
        Notify.show("已成功解压缩文件！")
    }
}

run_zip() {
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("解压缩操作只能在文件夹窗口中执行，请打开文件夹再操作！")
        return
    }
    zip_tool := "C:\Program Files\7-Zip\7z.exe"
    files := get_selected_files(2)
    if !files {
        Msgbox("没有选择文件，无法执行压缩操作！")
        return
    }
    first_file := RegExReplace(files, "`r`n.*", "")
    output_zip := path_info(first_file, "Ext:").Full . ".zip"
    file_list := StrReplace(files, "`r`n", '" "')
    cmd := Format('"{1}" a -tzip "{2}" "{3}"', zip_tool, output_zip, file_list)
    try {
        Run(cmd)
        Notify.show("文件已成功压缩为：" . path_info(output_zip).File)
    } catch as err {
        Msgbox("压缩操作失败：" . err.message)
    }
}

is_compressed(file_name) {
    ext := path_info(file_name).Ext
    compressed_exts := [".zip",".rar",".7z",".tar",".gz",".tgz",".bz2",".xz",".iso",".pptx",".ppt",".xlsx",".xls",".docx",".doc"]
    for index, extension in compressed_exts
        if (ext = extension)
            return 1
    return 0
}

; =================== 批量重命名 ===================

global_rule_map := Map(
    "52", "221228", "51", "221127", "50", "221030", "49", "220926", "48", "220829",
    "47", "220725", "46", "220628", "45", "220531", "44", "220427", "43", "220326"
)

rename_folder_from_clipboard() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空，请先复制文件夹路径！")
        return
    }
    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0
    for line in lines {
        path := Trim(line, "`"`t`r`n")
        if (path = "")
            continue
        if !FileExist(path) || !InStr(FileExist(path), "D")
            continue
        if SubStr(path, -1) = "\"
            path := SubStr(path, 1, -1)
        folder_name := StrSplit(path, "\").Pop()
        parent_dir := SubStr(path, 1, -StrLen(folder_name)) . "\"
        if RegExMatch(folder_name, "^220101_甲(\d{2})-(.*)$", &m) {
            num := m.1
            rest := m.2
            if global_rule_map.Has(num) {
                new_name := global_rule_map[num] "_甲" num "-" rest
                new_full_path := parent_dir . new_name
                if FileExist(new_full_path)
                    continue
                try {
                    DirMove(path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                }
            }
        }
    }
    MsgBox("📁 文件夹重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}

rename_folder_from_clipboard_pang() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空，请先复制文件夹路径！")
        return
    }
    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0
    for line in lines {
        path := Trim(line, "`"`t`r`n")
        if (path = "")
            continue
        if !FileExist(path) || !InStr(FileExist(path), "D")
            continue
        if SubStr(path, -1) = "\"
            path := SubStr(path, 1, -1)
        folder_name := StrSplit(path, "\").Pop()
        parent_dir := SubStr(path, 1, -StrLen(folder_name)) . "\"
        if RegExMatch(folder_name, "^220101_旁(\d{3})-(.*)$", &m) {
            num := m.1
            rest := m.2
            if global_rule_map.Has(num) {
                new_name := global_rule_map[num] "_旁" num "-" rest
                new_full_path := parent_dir . new_name
                if FileExist(new_full_path)
                    continue
                try {
                    DirMove(path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                }
            }
        }
    }
    MsgBox("📁 文件夹重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}

rename_files_from_clipboard() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空！")
        return
    }
    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0
    for line in lines {
        clean_line := Trim(line, "`"`t`r`n")
        if (clean_line = "")
            continue
        if !FileExist(clean_line) || InStr(FileExist(clean_line), "D")
            continue
        pi := path_info(clean_line)
        full_path := pi.Full
        original_base_name := pi.Fname . pi.Ext
        folder_path := pi.Folder
        if RegExMatch(original_base_name, "^220101_甲(\d{2})-(.*)$", &m) {
            num := m.1
            rest := m.2
            if global_rule_map.Has(num) {
                new_file_name := global_rule_map[num] "_甲" num "-" rest
                new_full_path := folder_path . "\" . new_file_name
                if FileExist(new_full_path)
                    continue
                try {
                    FileMove(full_path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                }
            }
        }
    }
    MsgBox("📄 文件重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}

rename_files_from_clipboard_pang() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空！")
        return
    }
    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0
    for line in lines {
        clean_line := Trim(line, "`"`t`r`n")
        if (clean_line = "")
            continue
        if !FileExist(clean_line) || InStr(FileExist(clean_line), "D")
            continue
        pi := path_info(clean_line)
        full_path := pi.Full
        original_base_name := pi.Fname . pi.Ext
        folder_path := pi.Folder
        if RegExMatch(original_base_name, "^220101_旁(\d{3})-(.*)$", &m) {
            num := m.1
            rest := m.2
            if global_rule_map.Has(num) {
                new_file_name := global_rule_map[num] "_旁" num "-" rest
                new_full_path := folder_path . "\" . new_file_name
                if FileExist(new_full_path)
                    continue
                try {
                    FileMove(full_path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                }
            }
        }
    }
    MsgBox("📄 文件重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}

rename_in_eagle() {
    SendInput("{F2}")
    Sleep(100)
    SendInput("^c")
    Sleep(100)
    if RegExMatch(A_Clipboard, "(.*)【(.*)】(.*)", &match) {
        name := match[1]
        tag := match[2]
        content := match[3]
        A_Clipboard := content . "【" . name . "×" . tag . "】"
    } else {
        Msgbox("未能匹配输入字符串")
        return
    }
    SendInput("^v")
}


zip() => Zipper.zip()
unzip() => Zipper.unzip()

class Zipper {

    static tool := "C:\Program Files\7-Zip\7z.exe"

    static compressed_exts := [
        ".zip", ".rar", ".7z", ".tar", ".gz", ".tgz",
        ".bz2", ".xz", ".iso",
        ".pptx", ".ppt", ".xlsx", ".xls", ".docx", ".doc"
    ]

    ; =================== 解压 ===================
    static unzip() {        
        ; Notify.show("开始解压缩文件...", , 1000)

        if !this._checkEnv("解压缩")
            return
        files := this._getFiles("解压缩")
        if !files
            return

        Loop Parse, files, "`n", "`r" {
            path := Trim(A_LoopField, "`"")

            info := path_info(path)
            if !this.isCompressed(info.Full) {
                Msgbox("选中的不是压缩文件，无法完成解压缩操作！")
                return
            }
            cmd := Format('"{1}" x "{2}" -o"{3}"\* -aoa', this.tool, A_LoopField, info.Folder)
            try Run(cmd)
            catch as err {
                Msgbox("解压缩失败：" . err.message)
                return
            }
        }
        ; Notify.hide_all()
    }

    ; =================== 压缩 ===================
    static zip() {
        ; Notify.show("开始压缩文件...", , 1000)

        if !this._checkEnv("压缩")        

            return
        files := this._getFiles("压缩")
        if !files
            return

        first_file := RegExReplace(files, "`r`n.*", "")
        output_zip := path_info(first_file, "Ext:").Full . ".zip"
        file_list  := StrReplace(files, "`r`n", " ")
        
        cmd := Format('"{1}" a -tzip "{2}" {3}', this.tool, output_zip, file_list)
        try {
            Run(cmd)
        } catch as err {
            Msgbox("压缩操作失败：" . err.message)
        }
        ; Notify.hide_all()
    }

    ; =================== 判断是否为压缩文件 ===================
    static isCompressed(file_name) {
        ext := path_info(file_name).Ext
        for _, extension in this.compressed_exts
            if (ext = extension)
                return true
        return false
    }

    ; =================== 内部方法 ===================
    static _checkEnv(action) {
        if !WinActive("ahk_class CabinetWClass") {
            Msgbox(action . "操作只能在文件夹窗口中执行，请打开文件夹再操作！")
            return false
        }
        return true
    }

    static _getFiles(action) {
        files := get_path_by_shortcut()
        if !files {
            Msgbox("未选择任何文件，无法完成" . action . "操作！")
            return ""
        }
        return files
    }
}


get_path_by_shortcut() {
    ; 判断是否在文件管理器中
    if !WinActive("ahk_class CabinetWClass") {
        return ""
    }

    ; 备份剪贴板
    clip_backup := ClipboardAll()
    A_Clipboard := ""

    ; 发送 Ctrl+Shift+C 复制文件路径
    Send("^+c")

    ; 等待剪贴板有内容（最多2秒）
    if !ClipWait(2) {
        A_Clipboard := clip_backup
        return ""
    }

    ; 获取结果
    files := A_Clipboard

    ; 还原剪贴板
    A_Clipboard := clip_backup

    return files
}

open_first_matching_file(folder, keyword, recursive := true) {
    if !DirExist(folder) {
        MsgBox "文件夹不存在：`n" folder
        return ""
    }

    mode := recursive ? "FR" : "F"

    Loop Files folder "\*.*", mode {
        if InStr(A_LoopFileName, keyword) {
            Run A_LoopFileFullPath
            return A_LoopFileFullPath
        }
    }

    MsgBox Format('未找到包含 "{}" 的文件。`n目录: {}', keyword, folder)
    return ""
}

open_duty_schedule() {
    return open_matching_file(
        "D:\work_files\duty_schedule",
        "卫星地球站排班表",
        false
    )
}

open_matching_file(folder, keyword, recursive := true) {
    if !DirExist(folder) {
        MsgBox "文件夹不存在：`n" folder
        return ""
    }

    files := []
    mode := recursive ? "FR" : "F"

    Loop Files folder "\*.*", mode {
        if InStr(A_LoopFileName, keyword) {
            files.Push(A_LoopFileFullPath)
        }
    }

    if files.Length = 0 {
        MsgBox Format('未找到包含 "{}" 的文件。`n目录: {}', keyword, folder)        
        return ""
    }

    if files.Length = 1 {
        Run files[1]
        return files[1]
    }

    return select_file_to_open(files, keyword)
}


select_file_to_open(files, keyword) {
    selected_file := ""

    gui_obj := Gui("+AlwaysOnTop", "选择要打开的文件")
    gui_obj.SetFont("s10", "Microsoft YaHei")

    gui_obj.AddText("w900", Format('找到多个包含 "{}" 的文件，请选择要打开的文件：', keyword))

    lv := gui_obj.AddListView("w900 r12", ["文件名", "所在文件夹"])

    for file_path in files {
        SplitPath file_path, &file_name, &dir_path
        lv.Add("", file_name, dir_path)
    }

    lv.ModifyCol(1, 300)
    lv.ModifyCol(2, 580)

    btn_open := gui_obj.AddButton("xm w100 Default", "打开")
    btn_cancel := gui_obj.AddButton("x+10 w100", "取消")

    open_selected(*) {
        row := lv.GetNext()

        if !row {
            MsgBox "请先选择一个文件。"
            return
        }

        selected_file := files[row]
        Run selected_file
        gui_obj.Destroy()
    }

    cancel_selected(*) {
        gui_obj.Destroy()
    }

    btn_open.OnEvent("Click", open_selected)
    btn_cancel.OnEvent("Click", cancel_selected)
    lv.OnEvent("DoubleClick", open_selected)

    gui_obj.Show()

    WinWaitClose gui_obj.Hwnd

    return selected_file
}