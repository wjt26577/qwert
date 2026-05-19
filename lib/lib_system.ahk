; ============================================================
;  lib_system.ahk — 系统工具 / 光标 / 网络 / 键盘状态 / 杂项
; ============================================================

; =================== 系统工具 ===================

open_device_manager()              => Run("devmgmt.msc")
open_control_panel()               => Run("control.exe")
open_task_manager()                => Run("taskmgr.exe")
open_system_information()          => Run("msinfo32.exe")
open_event_viewer()                => Run("eventvwr.msc")
open_services_manager()            => Run("services.msc")
open_computer_management()         => Run("compmgmt.msc")
open_disk_management()             => Run("diskmgmt.msc")
open_network_connections()         => Run("ncpa.cpl")
open_user_account_control_settings() => Run("useraccountcontrolsettings.exe")
open_firewall_settings()           => Run("firewall.cpl")
open_power_options()               => Run("powercfg.cpl")
open_sound_settings()              => Run("mmsys.cpl")
open_display_settings()            => Run("desk.cpl")
open_printers_and_faxes()          => Run("control printers")
open_registry_editor()             => Run("regedit.exe")
open_command_prompt()              => Run("cmd.exe")
open_notepad()                     => Run("notepad.exe")
open_calculator()                  => Run("calc.exe")
open_recycle_bin()  {
    try Run('::{645FF040-5081-101B-9F08-00AA002F954E}')
    catch as err
        Msgbox(err.message)
}

; =================== 电源 ===================

shutdown_me() {
    result := MsgBox("您的电脑将在 5 秒后关机。点击取消可中止操作。", "关机警告", "确定取消 超时,5")
    if (result = "取消" || result = "超时")
        Shutdown 1
}

shutdown_computer() {
    Result := MsgBox("5秒后将关机。点击确定中止操作。", "确认", "OK Icon! T5")
    if (Result = "OK")
        return
    Shutdown 1
}

restart_computer() {
    Result := MsgBox("5秒后将重启。点击确定中止操作。", "确认", "OK Icon! T5")
    if (Result = "OK")
        return
    Shutdown 2
}

lock_screen() {
    Result := MsgBox("5秒后将锁定屏幕。点击确定中止操作。", "确认", "OK Icon! T5")
    if (Result = "OK")
        return
    DllCall("LockWorkStation")
}

; =================== 当前路径 ===================

open_current_path() {
    if (WinActive("ahk_exe EXCEL.EXE")) {
        com_object := "Excel.Application"
    } else if (WinActive("ahk_exe WINWORD.EXE")) {
        com_object := "Word.Application"
    } else if (WinActive("ahk_exe POWERPNT.EXE")) {
        com_object := "PowerPoint.Application"
    } else {
        Run("d:\")
        return
    }
    try app := ComObjActive(com_object)
    catch {
        MsgBox("无法连接到应用程序。请确保已打开文件。")
        return
    }
    if (com_object ~= "Excel")
        file_path := app.ActiveWorkbook.FullName
    else if (com_object ~= "Word")
        file_path := app.ActiveDocument.FullName
    else
        file_path := app.ActivePresentation.FullName
    if (file_path = "") {
        MsgBox("当前没有打开任何文件。")
        return
    }
    Run("explorer.exe /select, " file_path)
}

; =================== 光标 ===================

SetCursorToWait() {
    hCursor := DllCall("LoadCursor", "Ptr", 0, "UInt", 32514, "Ptr")
    DllCall("SetSystemCursor", "Ptr", hCursor, "UInt", 32512)
}

RestoreDefaultCursor() {
    DllCall("SystemParametersInfo", "UInt", 0x0057, "UInt", 1, "Ptr", 0, "UInt", 0)
}

; =================== 网络检测 ===================

isServerReachable(server) {
    result := RunWait(A_ComSpec . " /c ping -n 1 " . server . " >nul", "", "Hide")
    return (result = 0)
}

isWebsiteReachable(url) {
    script := "curl -s -o nul --head --fail " . url
    result := RunWait(A_ComSpec . " /c " . script . " >nul", "", "Hide")
    return (result = 0)
}

; =================== 快捷方式 ===================

getShortcutTarget(shortcutPath) {
    wshShell := ComObject("WScript.Shell")
    shortcut := wshShell.CreateShortcut(shortcutPath)
    return(shortcut.TargetPath)
}

strToAppLink(str) => "D:\coreFiles\appLink\" . str . ".lnk"

; =================== 键盘状态 ===================



reverse_caps_letter(letter) {
    if (RegExMatch(letter, "^[a-z]$"))
        return GetKeyState("CapsLock", "T") ? StrUpper(letter) : letter
}

base_on_capslock(key := A_ThisHotkey) {
    return GetKeyState("CapsLock", "T") ? StrUpper(key) : key
}

send_against_capslock(key := A_ThisHotkey) {
    key := SubStr(key, -1)
    if (RegExMatch(key, "^[a-z]$"))
        GetKeyState("CapsLock", "T") ? SendInput(key) : SendInput(StrUpper(key))
}

send_with_capslock(key := A_ThisHotkey) {
    key := SubStr(key, -1)
    if (RegExMatch(key, "^[a-z]$"))
        GetKeyState("CapsLock", "T") ? SendInput(StrUpper(key)) : SendInput(StrLower(key))
    else
        SendInput(key)
}

send_char(hotkey := A_ThisHotkey) {
    char := SubStr(hotkey, -1)
    if RegExMatch(char, "^[a-z]$") && GetKeyState("CapsLock", "T")
        SendText(StrUpper(char))
    else
        SendText(char)
}

; =================== Morse ===================

Morse(key := "", timeout := 200) {
    tout := timeout / 1000
    if (key = "")
        Key := RegExReplace(A_ThisHotKey, "[\*\~\$\#\+\!\^]")
    pattern := 0
    loop {
        t := A_TickCount
        KeyWait(key)
        pattern += 1
        ErrorL2 := KeyWait(Key , "D T" . tout)
        if (!ErrorL2)
            return pattern
    }
}

Morseold(key := "", timeout := 200) {
    tout := timeout / 1000
    if (key = "")
        if (InStr(A_ThisHotKey, "&"))
            key := RegExReplace(A_ThisHotKey, ".*&\s*(\w+)$", "$1")
        else
            key := RegExReplace(A_ThisHotKey, "[\*\~\$\#\+\!\^]")
    pattern := ""
    loop {
        t := A_TickCount
        KeyWait(key)
        pattern .= A_TickCount - t > timeout
        ErrorL2 := KeyWait(key , "D T" . tout)
        if (!ErrorL2)
            return pattern
    }
}

; =================== 热键解析工具 ===================

mod_keya(combo) {
    combo := RegExReplace(combo, "^[\*\~\$]+")
    if RegExMatch(combo, "^(.*?)[ &]*([^\s&]+)$", &match)
        return Trim(match[1])
    return ""
}

main_keya(combo) {
    combo := RegExReplace(combo, "^[\*\~\$]+")
    if RegExMatch(combo, "^(.*?)[ &]*([^\s&]+)$", &match)
        return Trim(match[2])
    return Trim(combo)
}

extract_primary_key(combo) {
    if (InStr(combo, "&"))
        return RegExReplace(combo, ".*&\s*(\w+)$", "$1")
    return RegExReplace(combo, "[\*\~\$\#\+\!\^]")
}

; =================== 杂项功能 ===================

pixel_to_point(pixel, LOG_PIXELS := 88, POINTS_PER_INCH := 72) {
    hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
    pixels_per_inch := DllCall("GetDeviceCaps", "Ptr", hdc, "Int", LOG_PIXELS, "Int")
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
    return (pixel / pixels_per_inch * POINTS_PER_INCH)
}

keyfunc_rbutton_superkeyoff() {
    SendInput("{RButton}")
    superkey := False
}

keyfunc_space_enter() {
    if WinActive("ahk_exe EXCEL.EXE") {
        SendInput("!{Enter}")
        return
    }
    if WinActive("ahk_exe chrome.exe") {
        SendInput("+{Enter}")
        return
    }
    enter_wherever()
}

keyfunc_ctrl2() {
    KeyWait("Ctrl")
    if (A_TimeSinceThisHotkey > 300 || A_PriorKey != "LControl")
        return
    SendInput("^{Space}")
}

save_reload() {
    If WinActive("ahk_exe Code.exe") {
        SendInput("^s")
        Sleep(50)
    }
    Reload
}

find_file() {
    Loop Files, "C:\Program Files\WindowsApps\*.*", "R" {
        if (A_LoopFileName ~= "i)^(snipaste\.exe)$") {
            A_Clipboard := A_LoopFilePath
            MsgBox(A_LoopFilePath)
        }
    }
}

write_json_file() {
    filePath := "settings\settings.json"
    content := FileRead("settings\settings.json")
    oldLine := "`"snipaste`": `"D:\\system\\Snipaste-2.11-x64\\Snipaste.exe`""
    newLine := "`"snipaste`": `"C:\\Program Files\\WindowsApps\\45479liulios.17062D84F7C46_2.11.300.0_x64__p7pnf6hceqser\\snipaste.exe`""
    newContent := StrReplace(content, oldLine, newLine)
    if (newContent != content)
        MsgBox("成功找到并替换了 snipaste 路径。")
    else
        MsgBox("警告：未找到旧的 snipaste 路径。")
    fileObj := FileOpen(filePath, "w")
    fileObj.Write(newContent)
    fileObj.Close()
}

cmdb_update_excel_k_column() {
    keywordMap := Map(
        "卫视MCPC","地球站卫视MCPC系统", "东方卫视高清","地球站东方卫视高清系统",
        "东方卫视4K","地球站东方卫视4K系统", "贵江河","地球站贵江河系统",
        "浙江高清","地球站浙江高清系统", "互动HD1","地球站互动HD1系统",
        "互动HD2","地球站互动HD2系统", "互动HD4","地球站互动HD4系统",
        "欢笑剧场4K","地球站欢笑剧场4K系统", "浙江标清","地球站浙江标清系统",
        "传送","地球站传送系统"
    )
    try xlApp := ComObjActive("Excel.Application")
    catch {
        MsgBox("无法连接到 Excel，请确保 Excel 已打开。")
        ExitApp
    }
    ws := xlApp.ActiveWorkbook.ActiveSheet
    usedRows := ws.UsedRange.Rows.Count
    Loop usedRows {
        row := A_Index
        if (row = 1)
            Continue
        cellValue := ws.Cells(row, 8).Value
        if (!cellValue || cellValue = "")
            Continue
        text := "" . cellValue
        matches := []
        for keyword, target in keywordMap
            if InStr(text, keyword)
                matches.Push(target)
        uniqueMatches := []
        seen := Map()
        for item in matches {
            if !seen.Has(item) {
                uniqueMatches.Push(item)
                seen[item] := true
            }
        }
        result := ""
        if uniqueMatches.Length > 0
            result := StrJoin(uniqueMatches, ";")
        ws.Cells(row, 11).Value := result
    }
    MsgBox("CMDB 更新完成！已处理 " usedRows " 行。")
}

test_ditto() {
    Run "C:\Program Files\Ditto\Ditto.exe /Paste:2452"
    MsgBox(A_Clipboard)
}