; ============================================================
;  lib_mouse.ahk — 鼠标位置 / 边缘 / 窗口 / 控件 / 图片点击
; ============================================================

; =================== 基础位置 ===================

mouse_get_pos(&x := "", &y := "") {
    MouseGetPos &x, &y
    return {x: x, y: y}
}

mouse_is_at(x, y, tolerance := 0) {
    MouseGetPos &mx, &my
    return (Abs(mx - x) <= tolerance && Abs(my - y) <= tolerance)
}

mouse_is_in_rect(x1, y1, x2, y2) {
    MouseGetPos &mx, &my
    return (mx >= x1 && mx <= x2 && my >= y1 && my <= y2)
}

; =================== 屏幕边缘 ===================

mouse_is_at_edge(edge_size := 5) {
    MouseGetPos &x, &y
    return (x <= edge_size || y <= edge_size ||
            x >= A_ScreenWidth - edge_size ||
            y >= A_ScreenHeight - edge_size)
}

mouse_is_at_top(edge_size := 15) {
    MouseGetPos , &y
    return (y <= edge_size)
}

mouse_is_at_bottom(edge_size := 15) {
    MouseGetPos , &y
    return (y >= A_ScreenHeight - edge_size)
}

mouse_is_at_left(edge_size := 15) {
    MouseGetPos &x
    return (x <= edge_size)
}

mouse_is_at_right(edge_size := 15) {
    MouseGetPos &x
    return (x >= A_ScreenWidth - edge_size)
}

mouse_is_at_corner(corner_size := 50) {
    MouseGetPos &x, &y
    return ((x <= corner_size || x >= A_ScreenWidth - corner_size) &&
            (y <= corner_size || y >= A_ScreenHeight - corner_size))
}

mouse_is_top() {
    MouseGetPos &xpos, &ypos
    return (ypos < 15)
}

mouse_is_right() {
    MouseGetPos &xpos, &ypos
    return (xpos > 3800)
}

mouse_is_left() {
    MouseGetPos &xpos, &ypos
    return (xpos < 15)
}

; =================== 窗口检测 ===================

mouse_is_over_window(win_title := "", win_text := "", exclude_title := "", exclude_text := "") {
    MouseGetPos , , &id
    try {
        return WinExist(win_title . " ahk_id " . id, win_text, exclude_title, exclude_text)
    } catch {
        return false
    }
}

mouse_is_over_active_window() {
    MouseGetPos , , &id
    try {
        return (id == WinGetID("A"))
    } catch {
        return false
    }
}

mouse_get_window_under() {
    MouseGetPos , , &id
    try {
        return { id: id, title: WinGetTitle("ahk_id " . id), class: WinGetClass("ahk_id " . id), process: WinGetProcessName("ahk_id " . id) }
    } catch {
        return false
    }
}

; =================== 控件检测 ===================

mouse_is_over_control(class_nn, exact := true) {
    try {
        MouseGetPos , , , &control
    } catch {
        return false
    }
    if (exact)
        return (class_nn == control)
    return InStr(control, class_nn) > 0
}

mouse_get_control_under() {
    MouseGetPos , , , &control, &control_hwnd
    try {
        return { class_nn: control, hwnd: control_hwnd, text: ControlGetText(control_hwnd) }
    } catch {
        return false
    }
}

; =================== 多显示器 ===================

mouse_get_monitor() {
    MouseGetPos &x, &y
    return MonitorGet(MonitorGetPrimary(), &left, &top, &right, &bottom) ?
           MonitorGetWorkArea(MonitorGetPrimary(), &work_left, &work_top, &work_right, &work_bottom) : false
}

mouse_is_on_monitor(monitor_num) {
    MouseGetPos &x, &y
    try {
        MonitorGet(monitor_num, &left, &top, &right, &bottom)
        return (x >= left && x <= right && y >= top && y <= bottom)
    } catch {
        return false
    }
}

; =================== 高级检测 ===================

mouse_distance_from(x, y) {
    MouseGetPos &mx, &my
    return Sqrt((mx - x)**2 + (my - y)**2)
}

mouse_is_moving() {
    static last_x := 0, last_y := 0
    MouseGetPos &x, &y
    if (x != last_x || y != last_y) {
        last_x := x, last_y := y
        return true
    }
    return false
}

mouse_is_idle(idle_time_ms := 1000) {
    return (A_TimeIdlePhysical >= idle_time_ms)
}

mouse_button_is_down(button := "LButton") {
    return GetKeyState(button, "P")
}

mouse_any_button_down() {
    return (GetKeyState("LButton", "P") || GetKeyState("RButton", "P") ||
            GetKeyState("MButton", "P") || GetKeyState("XButton1", "P") ||
            GetKeyState("XButton2", "P"))
}

mouse_quick_edge_check() {
    MouseGetPos &x, &y
    return (x < 5 || y < 5 || x > A_ScreenWidth - 5 || y > A_ScreenHeight - 5)
}

mouse_cached_position(cache_time_ms := 50) {
    static cached_x := 0, cached_y := 0, last_check := 0
    current_time := A_TickCount
    if (current_time - last_check > cache_time_ms) {
        MouseGetPos &cached_x, &cached_y
        last_check := current_time
    }
    return {x: cached_x, y: cached_y}
}

; =================== 坐标转换 ===================

mouse_normalize_coords(&x, &y, from_mode := "Screen", to_mode := "Client", win_id := "") {
    if (from_mode == "Screen" && to_mode == "Client") {
        try {
            target_id := win_id ? win_id : WinGetID("A")
            WinGetPos &win_x, &win_y, , , "ahk_id " . target_id
            x -= win_x, y -= win_y
        }
    } else if (from_mode == "Client" && to_mode == "Screen") {
        try {
            target_id := win_id ? win_id : WinGetID("A")
            WinGetPos &win_x, &win_y, , , "ahk_id " . target_id
            x += win_x, y += win_y
        }
    }
}

mouse_is_in_client_area(win_id := "") {
    MouseGetPos &x, &y, &id
    target_id := win_id ? win_id : id
    try {
        WinGetPos &win_x, &win_y, &win_w, &win_h, "ahk_id " . target_id
        WinGetClientPos &client_x, &client_y, &client_w, &client_h, "ahk_id " . target_id
        return (x >= win_x + client_x && x <= win_x + client_x + client_w &&
                y >= win_y + client_y && y <= win_y + client_y + client_h)
    } catch {
        return false
    }
}

; =================== 图片点击 ===================

ClickPicture(ImageFilePath, ClickCount:=1, Speed:=0, vReturn:=true, ShowError:=true) {
    lPos := GetPicturePosition(ImageFilePath)
    if (lPos) {
        posX := lPos[1]
        posY := lPos[2]
        ClickPosition(posX, posY, ClickCount, Speed,,vReturn)
        return [posX, posY]
    } else {
        if (ShowError)
            MsgBox("找不到图片:`n" . ImageFilePath)
        return false
    }
}

ClickPosition(posX, posY, ClickCount:=1, Speed:=50, vCoordMode:="Screen", vReturn:=true) {
    if (CoordMode = "Relative") {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&posX_i, &posY_i)
        if (ClickCount)
            MouseClick "Left", posX, posY, ClickCount, Speed, "R"
        else
            MouseMove(posX, posY, Speed)
    } else {
        CoordMode("Mouse", vCoordMode)
        MouseGetPos(&posX_i, &posY_i)
        if (ClickCount) {
            sleep 100
            MouseClick "Left", posX, posY, ClickCount, Speed
        } else {
            MouseMove(posX, posY, Speed)
        }
    }
    if (vReturn)
        MouseMove(posX_i, posY_i, Speed)
}

GetPicturePosition(ImageFilePath) {
    myGui := Gui()
    myGui.AddPicture(, ImageFilePath)
    MyGui.GetPos(,, &width, &height)
    CoordMode("Pixel")
    ImageSearch(&FoundX, &FoundY, -2000, -2000, 5000, 5000, ImageFilePath)
    CoordMode("Mouse")
    if (FoundX)
        return [FoundX+width//2, FoundY+height//2]
    else
        return FoundX
}