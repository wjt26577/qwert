; ============================================================
;  lib_clipboard.ahk — 剪贴板操作 / 复制 / 粘贴 / 编辑
; ============================================================

send_by_clipboard(text) {
    old_clipboard := A_Clipboard
    try {
        A_Clipboard := text
        ClipWait(1)
        SendInput("^v")
        Sleep(100)
    } finally {
        SetTimer(() => (A_Clipboard := old_clipboard), -1000)
    }
}

send_text_by_clipboard(text_to_send) {
    clip_saved := ClipboardAll()
    A_Clipboard := text_to_send
    Sleep(40)
    SendInput("^v")
    Sleep(40)
    A_Clipboard := clip_saved
}

paste_pure_text() {
    try {
        clip_saved := ClipboardAll()
        A_Clipboard := A_Clipboard . ""
        Sleep(50)
        SendInput("^v")
        Sleep(50)
        A_Clipboard := clip_saved
    } catch as err
        Msgbox(err.Message)
}


paste_file(path) {
    if !FileExist(path) {
        MsgBox("文件不存在：`n" path)
        return
    }

    clip_saved := ClipboardAll()

    try {
        text := FileRead(path, "UTF-8")
        A_Clipboard := text

        if !ClipWait(1)
            throw Error("剪贴板写入失败")

        SendInput("^v")
    } catch as err {
        MsgBox(err.Message)
    } finally {
        Sleep(50)
        A_Clipboard := clip_saved  ; 恢复剪贴板
    }    
}

paste_ai_prompt_1() {
    paste_file(path_ai_prompt_1)
}

paste_ai_prompt_2() {
    paste_file(path_ai_prompt_2)
}

paste_ai_prompt_3() {
    paste_file(path_ai_prompt_3)
}


paste_text_only() {
    if WinActive("ahk_exe WINWORD.EXE")
        send_by_clipboard(A_Clipboard)
}

paste_format() {
    try {
        if WinActive("ahk_exe POWERPNT.EXE")
            SendInput("^+v")
        if WinActive("ahk_exe WINWORD.EXE") {
            app := ComObjActive('Word.Application')
            app.Selection.PasteFormat
        }
        if WinActive("ahk_exe EXCEL.EXE") {
            app := ComObjActive('Excel.Application')
            app.Selection.PasteSpecial(-4122)
        }
    } catch as err
        Msgbox(err.Message)
}

paste_word() {
    MouseClick("Left", , , 2)
    SendInput("^v")
}

copy_text_in_ppt() {
    try {
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection
        if (selection.Type < 2) {
            Msgbox("请先选择要粘贴纯文本的形状")
            return
        }
        if (selection.Type == 3) {
            SendInput("^v")
            return
        }
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange
        shape_range.TextFrame2.TextRange.Text := Trim(A_Clipboard, "`n`r")
    } catch as err {
        Msgbox("粘贴纯文本时发生错误: " . err.Message)
    } finally {
        ppt_application := ""
    }
}

has_text_selected() {
    clip_saved := ClipboardAll()
    A_Clipboard := ""
    Sleep(40)
    SendInput("^c")
    Sleep(40)
    result := (A_Clipboard != "")
    A_Clipboard := clip_saved
    return result
}

get_selected_text_by_clipboard() {
    clip_saved := ClipboardAll()
    A_Clipboard := ""
    SendInput("^c")
    sleep(200)
    selected_text := A_Clipboard
    A_Clipboard := clip_saved
    return selected_text
}

get_via_clipboard() {
    try {
        clip_saved := ClipboardAll()
        A_Clipboard := ""
        SendInput("^c")
        sleep(200)
        str := A_Clipboard
        A_Clipboard := clip_saved
        return str
    } catch as err
        Msgbox(err.Message)
}

BackupClipboardText() {
    formattedTime := FormatTime(,"MM-dd_HH-mm-ss")
    filePath := "D:\Backups\Text Backups\txtBackup_" . formattedTime . ".txt"
    if (FileExist(filePath))
        FileAppend("\r\n" . A_Clipboard, filePath)
    else
        FileAppend(A_Clipboard, filePath)
    A_Clipboard := ""
}

isObjSelected() {
    clipSaved := ClipboardAll()
    A_Clipboard := ""
    Send "^c"
    ClipWait(0.2)
    A_Clipboard := clipSaved
    clipSaved := ""
}

; =================== 编辑快捷键 ===================

copy_and_comment_line()    => SendInput("^c^q{End}{Enter}^v")
delete_all_right()         => SendInput('+{End}{Backspace}')
delete_all_left()          => SendInput('+{Home}{Backspace}')
delete_to_page_beginning() => SendInput('+^{Home}{Backspace}')
delete_to_page_end()       => SendInput('+^{End}{Backspace}')
enter_wherever()           => SendInput('{End}{Enter}')

delete_line() {
    SendInput("{Home}+{End}{Delete}")
}

send_ppt_code_1() {
    ppt_code_1 := "
    (
        ppt_application := ComObjActive(`"PowerPoint.Application`")
        selection := ppt_application.ActiveWindow.Selection
        if (selection.Type < 2) {
            Msgbox(`"请先选择要复制的形状`")
            return
        }
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else
            shape_range := selection.ShapeRange
    )"
    send_by_clipboard(ppt_code_1)
}