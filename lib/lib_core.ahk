; ============================================================
;  lib_core.ahk — CycAct / 热键解析 / 目标调度 / 函数调用
; ============================================================


; #region function core

copy_or_cut_by_count() {
    try {
        static count := 0

        if count == 0 {
            SetTimer(() => count := 0, -1000)
        }

        count := (count >= 2) ? 1 : count + 1

        if count == 1 {
            SendInput("^c")

            ; Snipaste 贴图
            if WinActive("Paster - Snipaste ahk_class Qt624QWindowToolSaveBits") {
                SendInput("{Ctrl Down}")
                Sleep(30)
                SendInput("c{Ctrl Up}{Esc 2}")
            }

            ; Word 复制格式
            if WinActive("ahk_exe WINWORD.EXE") {
                word_application := ComObjActive("Word.Application")
                word_application.Selection.CopyFormat
            }

        } else if count == 2 {
            SendInput("^x")
        }

        Sleep(80)
        add_clipboard_history()

    } catch as err {
        count := 0
        MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.Message, err.File, err.Line))
    }
}

action_ctrl() {
    try {        
        enter_text_mode()
        IMECtrl.to_english()
        Notify.show("English Input Method", "success", 500)

    } catch as err
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_lshift() {
    try {
        enter_text_mode()
        IMECtrl.toggle()

        if (IMECtrl.is_chinese())
            Notify.show("中文输入法", "info", 500)
        else
            Notify.show("English Input Method", "success", 500)  
        
    } catch as err
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_key8() {
    try {
        ; 如果拼音候选窗口存在，发送PgDn键, 否则activate_chinese_input  
        if ime_is_composing() {
            SendInput("{PgDn}")   
        } else {
            enter_text_mode()
            static count := 0  

            if count == 0
                SetTimer(() => count := 0, -1000)

            count := (count >= 2 ) ? 1 : count + 1 
            
            if count == 1 {    
                IMECtrl.to_chinese()
                Notify.show("中文输入法", "info", 500)
            } else if count == 2 {
                IMECtrl.to_upper()
                Notify.show("English UPPERCASE", "error", 500)
            }
        }    
    } catch as err {
        count := 0
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
    }  
} 

action_enter() {
    try {
        if ime_is_composing()
            SendInput("{Enter}")
        else
            SendInput("+{Enter}")
        
    } catch as err 
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_key3() {
    try {
        if !mouse_edge() {
            if ime_is_composing()
                SendInput("{Space}")
            else
                SendInput("{Enter}") 
        } else
            CycAct.slow("app_qianwen")

    } catch as err 
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_xbutton1() {
    try {
        if mouse_edge() {
            IMECtrl.to_chinese()
            Notify.show("中文输入法", "info", 500)
            SendInput("#h")
        } else {
            SendInput("{XButton1}")           
        }    
    } catch as err 
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_xbutton2() {
    try {
        if mouse_edge() {
            IMECtrl.to_chinese()
            Notify.show("中文输入法", "info", 500)
            SendInput("#h")
        } else {
            SendInput("{XButton2}")           
        }    
    } catch as err 
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_f23() {
    try {
        if mouse_edge() {
            CycAct.slow("!+{Tab}")
        } else {
            SendInput("{XButton2}")           
        }    
    } catch as err 
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_f24() {
    try {
        if mouse_edge() {
            CycAct.slow("!+{Tab}")
        } else {
            SendInput("{XButton2}")           
        }    
    } catch as err 
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 

action_space() {
    try {
        if mouse_edge() {
            CycAct.slow("app_keyboard")
        } else {
            SendInput("{Space}")        
        }    
    } catch as err 
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
} 



; #endregion 




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
        case "+": return "+"
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
        case "app":  run_app(target_payload)
        case "path": run_path(target_payload)
        case "key":  SendInput(target_payload)
        case "func": %target_payload%()
        case "text": SendText(target_payload)
        default:     %target_payload%()
    }
}

run_app(target_payload) {
    SplitPath target_payload, &file_name
    if file_name ~= "i)weixin" {
        if ProcessExist("Weixin.exe")
            SendInput("^!w")
        else
            Run(target_payload)

        return
    }

    if file_name ~= "i)Code|Notepad" {
        if WinExist("ahk_exe " file_name) 
            WinActivate()
        else
            Run(target_payload)

        return
    }

    if WinExist("ahk_exe " file_name) 
        WinActivate()

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

; ======================== CycAct ========================

; class CycAct {
;     static _cache := Map()
;     count := 0
;     target_map := Map()
;     mode := ""
;     modifier := ""
;     actions := []
;     timeout := 200
;     _timer_fn := ""

;     __New(mode, actions*) {
;         this.mode := mode
;         this.actions := actions
;         this._timer_fn := ObjBindMethod(this, "_on_timer")
;     }

;     static fast(actions*) => this._dispatch("fast", actions*)
;     static slow(actions*) => this._dispatch("slow", actions*)

;     static _dispatch(mode, actions*) {
;         key := A_ThisHotkey
;         if !this._cache.Has(key)
;             this._cache[key] := CycAct(mode, actions*)
;         this._cache[key].Call()
;     }

;     Call(*) {
;         try {
;             length := this.actions.Length
;             if length == 1 {
;                 run_target(get_target_map(this.actions[1]))
;                 return
;             }
;             if this.count == 0 {
;                 this.modifier := extract_modifier()
;                 if this.modifier
;                     SetTimer(this._timer_fn, -40)
;             }
;             this.count := (this.count >= length) ? 1 : this.count + 1
;             this.target_map := get_target_map(this.actions[this.count])
;             label := this.target_map["label"]
;             if this.mode == "fast"
;                 run_target(this.target_map)
;             if label != "unknown"
;                 Notify.show(label, "info", 5000)
;             if !this.modifier
;                 SetTimer(this._timer_fn, -this.timeout)
;         } catch as err {
;             this._reset()
;             MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
;         }
;     }

;     _on_timer() {
;         try {
;             if this.modifier && GetKeyState(this.modifier, "P") {
;                 SetTimer(this._timer_fn, -40)
;                 return
;             }
;             if this.mode == "slow"
;                 run_target(this.target_map)
;             this._reset()
;         } catch as err {
;             this._reset()
;             MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
;         }
;     }

;     _reset() {
;         this.count := 0
;         SetTimer(() => Notify.hide_all(), -100)
;     }
; }

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
        key := this._make_cache_key(mode, actions)

        if !this._cache.Has(key) {
            this._cache[key] := CycAct(mode, actions*)
        }

        this._cache[key].Call()
    }

    static _make_cache_key(mode, actions) {
        key := A_ThisHotkey "|" mode "|"

        for action in actions {
            ; 用长度 + 内容生成 key，避免简单拼接时出现冲突
            key .= StrLen(action) ":" action "|"
        }

        return key
    }

    Call(*) {
        try {
            length := this.actions.Length

            if length == 1 {
                this.target_map := get_target_map(this.actions[1])
                label := this.target_map["label"]
                if label != "unknown"
                    Notify.show(label, "info", 500)

                run_target(get_target_map(this.actions[1]))
                return
            }

            if this.count == 0 {
                this.modifier := extract_modifier()

                if this.modifier {
                    SetTimer(this._timer_fn, -40)
                }
            }

            this.count := (this.count >= length) ? 1 : this.count + 1

            this.target_map := get_target_map(this.actions[this.count])
            label := this.target_map["label"]

            if this.mode == "fast" {
                run_target(this.target_map)
            }

            if label != "unknown" {
                Notify.show(label, "info", 5000)
            }

            if !this.modifier {
                SetTimer(this._timer_fn, -this.timeout)
            }

        } catch as err {
            this._reset()
            MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
        }
    }

    _on_timer() {
        try {
            if this.modifier && GetKeyState(this.modifier, "P") {
                SetTimer(this._timer_fn, -40)
                return
            }

            if this.mode == "slow" {
                run_target(this.target_map)
            }

            this._reset()

        } catch as err {
            this._reset()
            MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
        }
    }

    _reset() {
        this.count := 0
        this.modifier := ""
        this.target_map := Map()

        SetTimer(this._timer_fn, 0)
        SetTimer(() => Notify.hide_all(), -100)
    }
}


; =================== 函数字符串调用 ===================

parse_function_call(function_string) {
    if(RegExMatch(function_string, "^(\w+)\s*\((.*)\)$", &match))
        return { name: match[1], params: parse_function_parameters(match[2]), has_params: true }
    if(RegExMatch(function_string, "^\w+$"))
        return { name: function_string, params: [], has_params: false }
    Throw Error("无法识别的函数格式: " . function_string)
}

parse_function_parameters(param_string) {
    if(!param_string || param_string == "")
        return []
    params := []
    current_param := ""
    in_quotes := false
    quote_char := ""
    Loop StrLen(param_string) {
        char := SubStr(param_string, A_Index, 1)
        switch char {
            case '"', "'":
                if(!in_quotes) {
                    in_quotes := true
                    quote_char := char
                } else if(char = quote_char) {
                    in_quotes := false
                    quote_char := ""
                }
                current_param .= char
            case ",":
                if(in_quotes)
                    current_param .= char
                else {
                    params.Push(clean_parameter(current_param))
                    current_param := ""
                }
            default:
                current_param .= char
        }
    }
    if(current_param != "")
        params.Push(clean_parameter(current_param))
    return params
}

clean_parameter(param) {
    param := Trim(param)
    if((StrLen(param) >= 2) &&
       ((SubStr(param, 1, 1) = '"' && SubStr(param, -1) = '"') ||
        (SubStr(param, 1, 1) = "'" && SubStr(param, -1) = "'")))
        param := SubStr(param, 2, StrLen(param) - 2)
    return param
}

execute_parsed_function(call_info) {
    try {
        function_obj := %call_info.name%
    } catch {
        Throw Error("函数 '" . call_info.name . "' 不存在")
    }
    if(!call_info.has_params || call_info.params.Length = 0)
        function_obj.Call()
    else
        function_obj.Call(call_info.params*)
    return true
}

call_function(func_str) {
    parsed := parse_function_call(func_str)
    return execute_parsed_function(parsed)
}

run_app2(app_path, param := "") {
    app_process := "ahk_exe " . path_info(app_path).File
    if RegExMatch(app_path, "i)eagle|fxsound|baidunetdisk|wxwork") {
        try Run(app_path . " " . param)
        catch as err
            Msgbox(err.Message)
        return
    }
    if (RegExMatch(app_path, "i)weixin\.exe") && ProcessExist("Weixin.exe")) {
        SendInput("^!w")
        return
    }
    if WinExist(app_process)
        WinActivate()
    else {
        try Run(app_path . " " . param)
        catch as err
            Msgbox(err.Message)
    }
}

run_server_app(app_path, param := "") {
    command := app_path . " " . param
    result := MsgBox("即将运行服务器程序, 确定要继续吗？", "确认", "OKCancel")
    if (result = "Ok")
        Run(command)
}

; =================== 小工具 ===================

set_global(Params*) {
    global
    Loop (Params.Length / 2) {
        local var_name := Params[2 * A_Index - 1]
        local var_value := Params[2 * A_Index]
        %var_name% := var_value
    }
}

variable_exists(name) {
    try {
        %name%
        return true
    } catch {
        return false
    }
}

show_thishotkey() {
    hotkey := StrReplace(A_ThisHotkey, "&", "&&")
    Notify.show(hotkey, "error", "500")
}


show_time() {
    time_string := FormatTime(, "yyyy年M月d日 dddd HH:mm")
    Notify.show(time_string, "warn", 5000)
}

do_nothing() => a := 1
noop() => a := 1

get_var_by_name(var_name) {
    if IsSet(%var_name%)
        return %var_name%
    else
        return 0
}

quote(p) => "`"" . p . "`""





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


; ==================== 调试 ====================
F24:: {
    CoordMode "Mouse", "Screen"
    MouseGetPos(&vx, &vy)

    info := Format("vx={} vy={}`n右边缘={} 左边缘={}", vx, vy, is_mouse_at_right_edge(), is_mouse_at_left_edge())
    
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

get_screen_top() {
    global screen_top
    return screen_top
}

get_screen_bottom() {
    global screen_bottom
    return screen_bottom
}

is_mouse_at_right_edge() {
    try DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
    edge_threshold  := 5
    CoordMode "Mouse", "Screen"
    MouseGetPos(&x)
    
    return (x >= get_screen_right() - edge_threshold)
}

is_mouse_at_left_edge() {
    try DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
    edge_threshold  := 5
    CoordMode "Mouse", "Screen"
    MouseGetPos(&x)

    ; MsgBox((x <= get_screen_left() + edge_threshold), , "T2")

    return (x <= get_screen_left() + edge_threshold)
}

mouse_top() {
    try DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
    edge_threshold  := 5
    CoordMode "Mouse", "Screen"
    MouseGetPos(, &y)
     
    return (y <= get_screen_top() + edge_threshold)
}

is_mouse_at_bottom_edge() {
    try DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
    edge_threshold  := 5
    CoordMode "Mouse", "Screen"
    MouseGetPos(, &y)
    
    return (y >= get_screen_bottom() - edge_threshold)
}


mouse_edge(side := "") {
    try DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")

    edge_threshold := 5
    CoordMode "Mouse", "Screen"
    MouseGetPos(&x)
    ; MsgBox(x)

    ; MsgBox (x >= get_screen_right() - edge_threshold) || (x <= get_screen_left() + edge_threshold), , "T2"
  
    if (side = "R")
        return (x >= get_screen_right() - edge_threshold)

    if (side = "L")
        return (x <= get_screen_left() + edge_threshold)

    return (x >= get_screen_right() - edge_threshold) || (x <= get_screen_left() + edge_threshold)
   
}


move_tab_new_window() {
    SendInput("^+y")
    Sleep(100)
    SendInput("#+{Left}") 
    WinMaximize("A")
}

