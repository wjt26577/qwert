
; #region class IMECtrl———————————————————————————————


class IMECtrl {    
; 中文输入法
; 134481924    ; 0x08040804 - 微软拼音输入法
; 67702532     ; 0x04090804 - 中文美式键盘

; 英文输入法  
; 67699721     ; 0x04090409 - 美式英语键盘

    static EN := 0x04090409
    static ZH := 0x08040804

    static get_layout() {
        try {
            hwnd := WinActive("ahk_class CabinetWClass")
                  ? WinGetID("ahk_class Progman")
                  : WinGetID("A")
            tid := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0)
            return DllCall("GetKeyboardLayout", "UInt", tid)
        }
        return 0
    }

    static set_layout(id) {
        try {
            ctl := ControlGetFocus("A")
            SendMessage(0x50, 0, id, , ctl)
        } catch
            SendMessage(0x50, 0, id, , "A")
    }

    static _ime_hwnd(win_title := "A") {
        hwnd := WinGetID(win_title)
        if WinActive(win_title) {
            ps := A_PtrSize
            buf := Buffer(cb := 4 + 4 + ps * 6 + 16, 0)
            NumPut("UInt", cb, buf, 0)
            if DllCall("GetGUIThreadInfo", "UInt", 0, "Ptr", buf)
                hwnd := NumGet(buf, 8 + ps, "UInt")
        }
        return DllCall("imm32\ImmGetDefaultIMEWnd", "UInt", hwnd)
    }

    static get_conv_mode(win_title := "A") {
        try return DllCall("SendMessage",
            "UInt", this._ime_hwnd(win_title),
            "UInt", 0x0283, "Int", 0x001, "Int", 0)
        return 0
    }

    static set_conv_mode(mode, win_title := "A") {
        try return DllCall("SendMessage",
            "UInt", this._ime_hwnd(win_title),
            "UInt", 0x0283, "Int", 0x002, "Int", mode)
        return -1
    }

    static is_chinese() => this.get_layout() == this.ZH && (this.get_conv_mode() & 0x01)
    static is_english() => this.get_layout() == this.EN
    static is_caps()    => GetKeyState("CapsLock", "T")

    static to_chinese() {
        SetCapsLockState(False)
        this.set_layout(this.ZH)
        Sleep(50)
        this.set_conv_mode(0x01)
    }

    static to_english() {
        SetCapsLockState(False)
        this.set_layout(this.EN)
    }

    static to_upper() {
        SetCapsLockState(True)
        this.set_layout(this.EN)
    }

    static toggle() {
        if this.get_layout() == this.ZH
            this.to_english()
        else
            this.to_chinese()
    }

    static toggle_caps() {
        SetCapsLockState(!this.is_caps())        
        ToolTip(this.is_caps() ? "大  写" : "小  写")
        SetTimer(() => ToolTip(), -5000)
    }

    static show_status() {
        if this.is_english()
            MsgBox("English Input Method", , "T2")
        else
            MsgBox("中文输入法", , "T2")
    }

    static show_debug() {
        layout := this.get_layout()
        conv   := this.get_conv_mode()
        info   := Format("
        (
            布局ID: {} (0x{:08X})
            转换模式: {} (0x{:02X})
            布局: {}
            输入: {}
            CapsLock: {}
        )", layout, layout, conv, conv
          , layout == this.EN ? "美式键盘" : layout == this.ZH ? "微软拼音" : "其他"
          , conv & 0x01 ? "中文" : "英文"
          , this.is_caps() ? "ON" : "OFF")
        ToolTip(info)
        SetTimer(() => ToolTip(), -5000)
    }
}

ime_is_composing() {
    try {
        hwnd := DllCall(
            "FindWindowExW", 
            "Ptr", 0, 
            "Ptr", 0, 
            "Str", "ApplicationFrameWindow", 
            "Str", "", 
            "UPtr"
        )
        if WinExist("ahk_id " . hwnd) {
            try WinGetPos &x, &y, &w, &h, "ahk_id " hwnd
            
            if (w >= 400)
                return true        
        } 
        return false
    
    } catch as err {
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
    }  
} 

