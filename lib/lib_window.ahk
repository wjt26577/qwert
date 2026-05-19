; ============================================================
;  lib_window.ahk — 窗口信息 / 关闭 / 最小化 / 最大化 / 布局
; ============================================================

show_win_info() {
    t1 := WinGetTitle("A")
    t2 := WinGetClass("A")
    t3 := WinGetProcessName("A")
    t4 := WinGetPID("A")
    cur_win := WinExist("A")
    A_Clipboard := t1
    win_data_text := t1 . "`n"
                  . "ahk_class " . t2 . "`n"
                  . "ahk_exe " . t3 . "`n"
                  . "ahk_pid " . t4 . "`n"
                  . "ahk_id " . cur_win
    MsgBox(win_data_text, "窗口信息")
}

close_window() {
    if !WinActive("ahk_group system_group")
        WinClose("A")
}

close_window_under_mouse() {
    try {
        MouseGetPos , , &id
        if (WinExist("ahk_group system_group ahk_id " . id))
            return
        if (WinExist("ahk_id " . id))
            WinClose
    } catch as err
        msgbox(err.Message)
}

; min_window() {
;     WinMinimize("A")
; }

min_win() {
    MouseGetPos , , &id

    if (WinExist("ahk_group system_group ahk_id " . id))
        return

    if (WinExist("ahk_group taskbar_group ahk_id " . id))
        return

    if (WinExist("ahk_id " . id))
        WinMinimize
}

max_window() {
    if !WinActive("ahk_group system_group")
        if WinGetMinMax("A")
            WinRestore("A")
        else
            WinMaximize("A")
}

close_win_by_ctrlw() {
    MouseGetPos(, , &id)  
    
    if  WinExist("ahk_group taskbar_group ahk_id " . id) {
        Click "Right"
        Sleep(600)
        SendInput("{Up}{Enter}")
        return
    }
    
    if WinExist("ahk_group ctrlw_group ahk_id " . id) {
        SendInput("!{F4}")
        return
    }

    Click "Left"
    SendInput("^w")
}

close_window_in_taskbar() {
    Click "Right"
    Sleep(1000)
    SendInput("{Up}{Space}")
}

left_window() {
    sleep(200)
    count := 0
    window_list := WinGetList()
    for hWnd in window_list {
        try {
            style := WinGetStyle(hWnd)
            if (style & 0x40000)
                count++
        } catch {
            continue
        }
    }
    cmd := count > 2 ? "41" : (count == 1 ? "11" : "31")
    SendInput("#z")
    sleep(300)
    SendInput(cmd)
    If !WinActive("ahk_exe Weixin.exe") {
        sleep(300)
        SendInput("{Escape}")
    }
}

right_window() {
    MouseClick("left")
    sleep(200)
    count := 0
    windowList := WinGetList()
    for hWnd in windowList {
        try {
            Style := WinGetStyle(hWnd)
            if (Style & 0x40000)
                count++
        } catch {
            continue
        }
    }
    cmd := count > 2 ? "42" : (count == 1 ? "12" : "32")
    SendInput("#z")
    sleep(200)
    SendInput(cmd)
    sleep(200)
    SendInput("{Escape}")
}

; =================== 窗口探测 ===================

MouseIsOver(WinTitle) {
    MouseGetPos , , &id
    return WinExist(WinTitle . " ahk_id " . id)
}

mouse_is_over(WinTitle) {
    MouseGetPos , , &id
    return WinExist(WinTitle . " ahk_id " . id)
}

mouse_is_over_taskbar() {
    MouseGetPos , , &id
    return (WinExist("ahk_class Shell_TrayWnd ahk_id " . id) or WinExist("ahk_class Shell_SecondaryTrayWnd ahk_id " . id))
}

Pos() {
    try {
        MouseGetPos &xpos, &ypos, &id, &control
        ahk_class := WinGetClass(id)
        winTitle := WinGetTitle(id)
        ahk_exe := WinGetProcessName(id)
        return { ahk_exe: ahk_exe, xpos: xpos, ypos: ypos, ahk_id: id, ahk_class: ahk_class, winTitle: winTitle, control: control }
    }
}

WinGet() {
    if id := WinGetID("A") {
        return { ahk_exe: WinGetProcessName(id), ahk_id: id, ahk_class: WinGetClass(id), winTitle: WinGetTitle(id) }
    } else {
        MsgBox "No active window found"
        return
    }
}

Focused() {
    id := WinGetID("A")
    FocusedHwnd := ControlGetFocus("A")
    try FocusedClassNN := ControlGetClassNN(FocusedHwnd)
    try ahk_class := WinGetClass(FocusedHwnd)
    try winTitle := WinGetTitle(FocusedHwnd)
    try ahk_exe := WinGetProcessName(FocusedHwnd)
    try FocusedInfo := { ahk_exe: ahk_exe, ahk_id: FocusedHwnd, ahk_class: ahk_class, winTitle: winTitle, control: FocusedClassNN }
    return FocusedInfo
}