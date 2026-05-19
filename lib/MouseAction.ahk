

; ; #region cat mouse


; extract_modifier(key := A_ThisHotkey) {  
;     hotkey := Trim(key, "~*$")

;     if InStr(hotkey, "&")
;         return Trim(StrSplit(hotkey, "&")[1])

;     prefix2 := SubStr(hotkey, 1, 2)
;     switch prefix2 {
;         case ">^": return "RControl"
;         case "<^": return "LControl"
;         case ">!": return "RAlt"
;         case "<!": return "LAlt"
;         case ">+": return "RShift"
;         case "<+": return "LShift"
;         case ">#": return "RWin"
;         case "<#": return "LWin"
;     }

;     switch SubStr(hotkey, 1, 1) {
;         case "^": return "LControl"
;         case "!": return "LAlt"
;         case "+": return "LShift"
;         case "#": return "LWin"
;         default:  return ""           ; 统一返回空字符串，不再返回 False
;     }
; }




; get_target_map(target_string) {  

;     if mso_config.Has(target_string)
;         return mso_config[target_string]   
        
;     if func_config.Has(target_string)
;         return func_config[target_string]  
    
;     if key_config.Has(target_string)
;         return key_config[target_string]
    
;     if run_config.Has(target_string)
;         return run_config[target_string]
    
;     if text_config.Has(target_string)
;         return text_config[target_string]    
             
;     return Map("payload", target_string, "label", "unknown", "kind", "func")   
; } 

; run_target(target) {
;     target_kind := target["kind"]
;     target_payload := target["payload"]
    
;     switch target_kind {
;         case "mso":
;             ppt_app := ComObjActive("PowerPoint.Application")
;             ppt_app.CommandBars.ExecuteMso(target_payload) 
;         case "app":
;             run_app(target_payload)
;         case "path":
;             run_path(target_payload)
;         case "key":
;             SendInput(target_payload)
;         case "func":
;             %target_payload%()                
;         case "text":
;             SendText(target_payload)
;         default:                
;             %target_payload%()  
;     }  
; } 

; run_app(target_payload) {
;     if target_payload ~= "i)weixin" {
;         if ProcessExist("Weixin.exe")
;             SendInput("^!w")
;         else 
;             Run(target_payload)            
;         return  
;     } 
;     SplitPath target_payload, &file_name
;     WinExist("ahk_exe " . file_name) ? WinActivate() : 0    
;     Run(target_payload)  
; } 

; run_path(target_payload) {
;     if target_payload ~= "i)^\\\\" {            
;         result := MsgBox("即将运行服务器程序, 确定要继续吗？", "确认", "OKCancel")              
;         (result = "OK") ? Run(target_payload) : 0
;         return
;     }     
;     Run(target_payload)  
; }


; ; class MouseAction {
; ;     count := 0
; ;     target_map := Map()
; ;     modifier := ""
; ;     mode := ""
; ;     actions := []
; ;     _timer_fn := ""

; ;     ; mode: "fast"=按下立即执行, "slow"=松开后执行
; ;     __New(mode, actions*) {
; ;         this.mode := mode
; ;         this.actions := actions
; ;         this._timer_fn := ObjBindMethod(this, "_on_timer")
; ;     }

; ;     Call(*) {
; ;         try {
; ;             length := this.actions.Length

; ;             if this.count == 0 {
; ;                 this.modifier := extract_modifier()

; ;                 ; if this.modifier == "LButton"
; ;                 ;     MouseClick(,,,,, "U")
                
; ;                 SetTimer(this._timer_fn, -40)
; ;             }

; ;             this.count := (this.count >= length) ? 1 : this.count + 1
; ;             target_string := this.actions[this.count]
; ;             this.target_map := get_target_map(target_string)

; ;             label := this.target_map["label"]

; ;             if this.mode == "fast"
; ;                 run_target(this.target_map)

; ;             if label != "unknown"
; ;                 Notify.show(label, "info", 5000)

; ;         } catch as err {
; ;             this._reset()
; ;             MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
; ;         }
; ;     }

; ;     _on_timer() {
; ;         try {
; ;             if this.modifier && GetKeyState(this.modifier, "P") {
; ;                 SetTimer(this._timer_fn, -40)
; ;                 return
; ;             }

; ;             if this.mode == "slow"
; ;                 run_target(this.target_map)

; ;             this._reset()

; ;         } catch as err {
; ;             this._reset()
; ;             MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
; ;         }
; ;     }

; ;     _reset() {
; ;         this.count := 0
; ;         Notify.hide_all()
; ;     }
; ; }

; class MouseAction {
;     static _cache := Map()

;     count := 0
;     target_map := Map()
;     modifier := ""
;     mode := ""
;     actions := []
;     _timer_fn := ""

;     __New(mode, actions*) {
;         this.mode := mode
;         this.actions := actions
;         this._timer_fn := ObjBindMethod(this, "_on_timer")
;     }

;     ; ★ 静态入口：按 A_ThisHotkey 自动区分实例
;     static action(mode, actions*) {
;         key := A_ThisHotkey  ; 每个热键独立状态
;         if !this._cache.Has(key)
;             this._cache[key] := MouseAction(mode, actions*)
;         this._cache[key].Call()
;     }

;     Call(*) {
;         try {
;             length := this.actions.Length
;             if this.count == 0 {
;                 this.modifier := extract_modifier()
;                 SetTimer(this._timer_fn, -40)
;             }
;             this.count := (this.count >= length) ? 1 : this.count + 1
;             target_string := this.actions[this.count]
;             this.target_map := get_target_map(target_string)
;             label := this.target_map["label"]
;             if this.mode == "fast"
;                 run_target(this.target_map)
;             if label != "unknown"
;                 Notify.show(label, "info", 5000)
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
;         Notify.hide_all()
;     }
; }


; class CatAction {
;     count := 0
;     target_map := Map()
;     mode := ""
;     actions := []
;     _timer_fn := ""
;     timeout := 500

;     ; mode: "fast" = 每次按下立即执行, "slow" = 超时后执行最终选择
;     __New(mode, actions*) {
;         this.mode := mode
;         this.actions := actions
;         this._timer_fn := ObjBindMethod(this, "_on_timer")
;     }

;     Call(*) {
;         try {
;             length := this.actions.Length

;             this.count := (this.count >= length) ? 1 : this.count + 1
;             target_string := this.actions[this.count]
;             this.target_map := get_target_map(target_string)

;             label := this.target_map["label"]

;             if this.mode == "fast"
;                 run_target(this.target_map)

;             if label != "unknown"
;                 Notify.show(label, "info", 5000)

;             ; 每次调用都重置计时器（防抖）
;             SetTimer(this._timer_fn, -this.timeout)

;         } catch as err {
;             this._reset()
;             MsgBox(Format("Error: {1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
;         }
;     }

;     _on_timer() {
;         try {
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
;         Sleep(300)
;         Notify.hide_all()
;     }
; }

; ; #endregion



; ==================== 工具函数 ====================

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