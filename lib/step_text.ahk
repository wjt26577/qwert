; ==============================================================================
;  step_text.ahk —— 多次按键文本插入模块
;
;  用法：
;     按一下：插入 STEP_TEXT_LIST[1]
;     按两下：插入 STEP_TEXT_LIST[2]
;     按三下：插入 STEP_TEXT_LIST[3]
;     ...
;
;  特性：
;     - 模仿 step_clip：多次触发选择候选，松开修饰键后提交
;     - 中央 GUI 预览
;     - 不使用剪贴板，默认 SendText() 插入，不污染剪贴板历史
;     - 支持多显示器，显示在激活窗口所在显示器
; ==============================================================================

#Requires AutoHotkey v2.0


; ------------------------------------------------------------------------------
; 文本候选
; 你主要改这里
; ------------------------------------------------------------------------------
global STEP_TEXT_LIST := [
    "git status",
    "git add .",
    "git commit -m ",
    "git init",
    "git push",
    "git pull",
]


; ------------------------------------------------------------------------------
; 如果你的热键是 XButton1 & T，就保持 XButton1
; 如果是 CapsLock & T，就改成 CapsLock
; ------------------------------------------------------------------------------
global STEP_TEXT_HOLD_KEY := "XButton1"


; ------------------------------------------------------------------------------
; 每个显示器的独立配置
; ------------------------------------------------------------------------------
global STEP_TEXT_CONFIG := Map(
    1, {
        w:  600,
        h:  180,
        ts: 14,
        ms: 11,
        body_rows: 8,
        bg: "navy",
        tc: "FFFFFF",
        mc: "EAEAEA",
        font: "微软雅黑",
        pad_x: 16,
        pad_y: 12,
    },

    2, {
        w:  1020,
        h:  260,
        ts: 20,
        ms: 16,
        body_rows: 8,
        bg: "navy",
        tc: "FFFFFF",
        mc: "EAEAEA",
        font: "微软雅黑",
        pad_x: 30,
        pad_y: 22,
    },
)

global STEP_TEXT_DEFAULT := {
    w:  500,
    h:  150,
    ts: 14,
    ms: 11,
    body_rows: 6,
    bg: "navy",
    tc: "FFFFFF",
    mc: "EAEAEA",
    font: "微软雅黑",
    pad_x: 16,
    pad_y: 12,
}


; ------------------------------------------------------------------------------
; 状态
; ------------------------------------------------------------------------------
global _stepTextIdx    := 0
global _stepTextTarget := ""
global _stepTextGuis   := Map()


; ==============================================================================
;  激活显示器
; ==============================================================================
_step_text_active_monitor() {
    try {
        hwnd := WinGetID("A")
        if !hwnd
            return MonitorGetPrimary()

        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)

        cx := x + w / 2
        cy := y + h / 2

        loop MonitorGetCount() {
            MonitorGet(A_Index, &l, &t, &r, &b)
            if (cx >= l && cx < r && cy >= t && cy < b)
                return A_Index
        }
    } catch {
    }

    return MonitorGetPrimary()
}


_step_text_get_cfg(mon) {
    global STEP_TEXT_CONFIG, STEP_TEXT_DEFAULT

    if STEP_TEXT_CONFIG.Has(mon)
        return STEP_TEXT_CONFIG[mon]

    return STEP_TEXT_DEFAULT
}


_step_text_calc_pos(mon, cfg) {
    try {
        MonitorGetWorkArea(mon, &l, &t, &r, &b)
    } catch {
        MonitorGetWorkArea(MonitorGetPrimary(), &l, &t, &r, &b)
    }

    work_w := r - l
    work_h := b - t

    x := l + Round((work_w - cfg.w) / 2)
    y := t + Round((work_h - cfg.h) / 2)

    return { x:x, y:y }
}


; ==============================================================================
;  懒加载 GUI
; ==============================================================================
_step_text_ensure_gui(mon) {
    global _stepTextGuis

    if _stepTextGuis.Has(mon)
        return _stepTextGuis[mon]

    cfg := _step_text_get_cfg(mon)

    g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000 -DPIScale",
             "step_text_m" mon)

    g.BackColor := cfg.bg
    g.MarginX := cfg.pad_x
    g.MarginY := cfg.pad_y

    inner_w := cfg.w - cfg.pad_x * 2

    ; 标题
    g.SetFont("s" cfg.ts " c" cfg.tc " Bold", cfg.font)
    titleCtrl := g.Add("Text", "Left w" inner_w " BackgroundTrans", "#0")

    ; 正文
    g.SetFont("s" cfg.ms " c" cfg.mc " Norm", cfg.font)
    bodyCtrl := g.Add("Text",
        "xm Left w" inner_w " r" cfg.body_rows " BackgroundTrans", "")

    obj := {
        gui: g,
        title: titleCtrl,
        body: bodyCtrl,
        cfg: cfg,
        inner_w: inner_w,
    }

    _stepTextGuis[mon] := obj
    return obj
}


; ==============================================================================
;  显示
; ==============================================================================
_step_text_show(index, text) {
    mon := _step_text_active_monitor()

    _step_text_hide_except(mon)

    o := _step_text_ensure_gui(mon)

    o.title.Text := "#" index "  文本插入"
    o.body.Text  := text

    pos := _step_text_calc_pos(mon, o.cfg)

    o.gui.Show(Format("x{1} y{2} w{3} h{4} NoActivate",
        pos.x, pos.y, o.cfg.w, o.cfg.h))
}


_step_text_hide_except(keep_mon) {
    global _stepTextGuis

    for mon, o in _stepTextGuis {
        if (mon != keep_mon) {
            try o.gui.Hide()
        }
    }
}


_step_text_hide_all() {
    global _stepTextGuis

    for mon, o in _stepTextGuis {
        try o.gui.Hide()
    }
}


; ==============================================================================
;  主入口
; ==============================================================================
step_text() {
    global STEP_TEXT_LIST
    global _stepTextIdx, _stepTextTarget

    try {
        if (!IsSet(STEP_TEXT_LIST) || STEP_TEXT_LIST.Length < 1)
            return

        ; 第一次按下时启动等待器
        if (_stepTextIdx == 0)
            SetTimer(_step_text_waiting, -40)

        ; 循环选择
        if (_stepTextIdx >= STEP_TEXT_LIST.Length)
            _stepTextIdx := 1
        else
            _stepTextIdx += 1

        _stepTextTarget := STEP_TEXT_LIST[_stepTextIdx]

        _step_text_show(_stepTextIdx, _stepTextTarget)

    } catch as err {
        _stepTextIdx := 0
        _stepTextTarget := ""
        _step_text_hide_all()

        MsgBox(Format("step_text Error: {1}`nFile: {2}`nLine: {3}",
            err.message, err.file, err.Line))
    }
}


; ==============================================================================
;  等待松开修饰键，然后提交
; ==============================================================================
_step_text_waiting() {
    global _stepTextIdx

    try {
        hold_key := _step_text_get_hold_key()

        if (hold_key != "" && GetKeyState(hold_key, "P")) {
            SetTimer(_step_text_waiting, -40)
            return
        }

        _step_text_commit()

    } catch as err {
        _stepTextIdx := 0
        _step_text_hide_all()
        MsgBox("step_text timer Error: " err.message)
    }
}

_step_text_get_hold_key() {
    global STEP_TEXT_HOLD_KEY

    ; 如果主程序里有 extract_modifier()，优先使用它
    try {
        mod := extract_modifier()
        if mod
            return mod
    } catch Error as e {
        ; 没有 extract_modifier()，或调用失败，就回退默认键
    }

    return STEP_TEXT_HOLD_KEY
}


; ==============================================================================
;  提交插入
; ==============================================================================
_step_text_commit() {
    global _stepTextIdx, _stepTextTarget

    try {
        if (_stepTextTarget != "") {
            ; 默认不用剪贴板，避免污染 clipboard_history
            SendText(_stepTextTarget)
        }
    } catch as err {
        MsgBox("step_text commit failed: " err.message)
    }

    _stepTextIdx := 0
    _stepTextTarget := ""
    _step_text_hide_all()
}


; ==============================================================================
;  状态判断
; ==============================================================================
step_text_active() {
    global _stepTextIdx
    return _stepTextIdx > 0
}


; ==============================================================================
;  回退到上一条候选
; ==============================================================================
step_text_back() {
    global STEP_TEXT_LIST
    global _stepTextIdx, _stepTextTarget

    if (_stepTextIdx <= 1)
        return

    if (!IsSet(STEP_TEXT_LIST) || STEP_TEXT_LIST.Length < 1)
        return

    _stepTextIdx -= 1

    if (_stepTextIdx > STEP_TEXT_LIST.Length)
        _stepTextIdx := STEP_TEXT_LIST.Length

    _stepTextTarget := STEP_TEXT_LIST[_stepTextIdx]

    _step_text_show(_stepTextIdx, _stepTextTarget)

    SetTimer(_step_text_waiting, -40)
}