


is_text_mode() {
    global text_mode
    return text_mode
}

enter_text_mode() {
    global text_mode
    text_mode := true
}

exit_text_mode() {
    global text_mode
    text_mode := false
}

; ==================== 启动 ====================
init_screen_bounds()

; 查找鼠标所在显示器
find_monitor_at(vx, vy) {
    global monitors
    for mon in monitors
        if (vx >= mon.left && vx < mon.right && vy >= mon.top && vy < mon.bottom)
            return mon
    return monitors[1]
}


global monitors := []


; ==================== 调试 ====================
F24:: {
    DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")  ; ★ 同样加上
    CoordMode "Mouse", "Screen"
    MouseGetPos(&vx, &vy)

    info := Format("vx={} vy={}`n右边缘={} 左边缘={}", vx, vy, is_mouse_at_right_edge(), is_mouse_at_left_edge())
    for i, mon in monitors
        info .= Format("`n显示器{}: [{},{}] DPI={} ratio={}", i, mon.left, mon.right, mon.dpi, Round(mon.ratio, 2))

    MsgBox(info, , "T5")
}



; ===== getter =====
get_screen_left() {
    global screen_left
    return screen_left
}
get_screen_right() {
    global screen_right
    return screen_right
}

; ===== 判断 =====
is_mouse_at_right_edge() {
    edge_threshold  := 5
    DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
    CoordMode "Mouse", "Screen"
    MouseGetPos(&x)
    
    return (x >= get_screen_right() - edge_threshold)
}


is_mouse_at_left_edge() {
    edge_threshold  := 5
    DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
    CoordMode "Mouse", "Screen"
    MouseGetPos(&x)

    return (x <= get_screen_left() + edge_threshold)
}


mouse_edge(side := "") {
    edge_threshold := 5
    DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
    CoordMode "Mouse", "Screen"
    MouseGetPos(&x)

    if (side = "R")
        return (x >= get_screen_right() - edge_threshold)

    if (side = "L")
        return (x <= get_screen_left() + edge_threshold)

    return (x >= get_screen_right() - edge_threshold) || (x <= get_screen_left() + edge_threshold)
}





extract_modifier(key := A_ThisHotkey) {
    hotkey := Trim(key, "~*$")

    if InStr(hotkey, "&")
        return Trim(StrSplit(hotkey, "&")[1])

    prefix2 := SubStr(hotkey, 1, 2)
    switch prefix2 {
        case ">^": return "RControl"
        case "<^": return "LControl"
        case ">!": return "RAlt"
        case "<!": return "LAlt"
        case ">+": return "RShift"
        case "<+": return "LShift"
        case ">#": return "RWin"
        case "<#": return "LWin"
    }

    switch SubStr(hotkey, 1, 1) {
        case "^": return "LControl"
        case "!": return "LAlt"
        case "+": return "LShift"
        case "#": return "LWin"
        default:  return ""
    }
}

get_target_map(target_string) {
    if mso_config.Has(target_string)
        return mso_config[target_string]

    if func_config.Has(target_string)
        return func_config[target_string]

    if key_config.Has(target_string)
        return key_config[target_string]

    if run_config.Has(target_string)
        return run_config[target_string]

    if text_config.Has(target_string)
        return text_config[target_string]

    return Map("payload", target_string, "label", "unknown", "kind", "func")
}

run_target(target) {
    target_kind := target["kind"]
    target_payload := target["payload"]

    switch target_kind {
        case "mso":
            ppt_app := ComObjActive("PowerPoint.Application")
            ppt_app.CommandBars.ExecuteMso(target_payload)
        case "app":
            run_app(target_payload)
        case "path":
            run_path(target_payload)
        case "key":
            SendInput(target_payload)
        case "func":
            %target_payload%()
        case "text":
            SendText(target_payload)
        default:
            %target_payload%()
    }
}

run_app(target_payload) {
    if target_payload ~= "i)weixin" {
        if ProcessExist("Weixin.exe")
            SendInput("^!w")
        else
            Run(target_payload)
        return
    }
    SplitPath target_payload, &file_name
    WinExist("ahk_exe " . file_name) ? WinActivate() : 0
    Run(target_payload)
}

run_path(target_payload) {
    if target_payload ~= "i)^\\\\" {
        result := MsgBox("即将运行服务器程序, 确定要继续吗？", "确认", "OKCancel")
        (result = "OK") ? Run(target_payload) : 0
        return
    }
    Run(target_payload)
}

; ==================== CycAct 类 ====================

class CycAct {
    static _cache := Map()

    count := 0
    target_map := Map()
    mode := ""
    modifier := ""
    actions := []
    timeout := 500
    _timer_fn := ""

    __New(mode, actions*) {
        this.mode := mode
        this.actions := actions
        this._timer_fn := ObjBindMethod(this, "_on_timer")
    }

    static fast(actions*) => this._dispatch("fast", actions*)
    static slow(actions*) => this._dispatch("slow", actions*)

    static _dispatch(mode, actions*) {
        key := A_ThisHotkey
        if !this._cache.Has(key)
            this._cache[key] := CycAct(mode, actions*)
        this._cache[key].Call()
    }

    Call(*) {
        try {
            length := this.actions.Length

            ; 单个动作 → 立即执行
            if length == 1 {
                run_target(get_target_map(this.actions[1]))
                return
            }

            ; 首次按下：检测修饰键
            if this.count == 0 {
                this.modifier := extract_modifier()
                if this.modifier
                    SetTimer(this._timer_fn, -40)

                if this.modifier == "LButton"
                    MouseClick(,,,,, "U")
            }

            this.count := (this.count >= length) ? 1 : this.count + 1
            this.target_map := get_target_map(this.actions[this.count])
            label := this.target_map["label"]

            if this.mode == "fast"
                run_target(this.target_map)

            if label != "unknown"
                Notify.show(label, "info", 5000)

            ; 无修饰键 → 超时防抖
            if !this.modifier
                SetTimer(this._timer_fn, -this.timeout)

        } catch as err {
            this._reset()
            MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
        }
    }

    _on_timer() {
        try {
            ; 修饰键还按着 → 继续等
            if this.modifier && GetKeyState(this.modifier, "P") {
                SetTimer(this._timer_fn, -40)
                return
            }

            if this.mode == "slow"
                run_target(this.target_map)

            this._reset()

        } catch as err {
            this._reset()
            MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
        }
    }

    _reset() {
        this.count := 0
        SetTimer(() => Notify.hide_all(), -300)
    }
}