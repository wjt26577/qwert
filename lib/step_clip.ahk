; ==============================================================================
;  step_clip.ahk —— 剪贴板候选预览
;
;  特性：
;     - 保留原文（含换行、空格，不做任何截断）
;     - 通过 GUI 高度自然裁剪超出部分
;     - 只在激活窗口所在显示器显示
;     - 屏幕正中央
;     - 深绿色
;     - 持久 GUI，高频刷新不丢键
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; 每个显示器的独立配置
;   body_rows : 正文预留行数（控件初始高度，决定能显示多少行）
; ------------------------------------------------------------------------------
global STEP_CLIP_CONFIG := Map(
    ; 显示器 1
    1, {
        w:  600,
        h:  200,
        ts: 14,
        ms: 10,
        body_rows: 20,
        ; bg: "084706",
        bg: "navy",
        tc: "FFFFFF",
        mc: "EAEAEA",
        font: "微软雅黑",
        pad_x: 16,
        pad_y: 12,
    },

    ; 显示器 2（4K）
    2, {
        w:  1020,
        h:  308,
        ts: 20,
        ms: 16,
        body_rows: 12,
        bg: "navy",
        tc: "FFFFFF",
        mc: "EAEAEA",
        font: "微软雅黑",
        pad_x: 30,
        pad_y: 22,
    },
)

global STEP_CLIP_DEFAULT := {
    w:  420,
    h:  110,
    ts: 14,
    ms: 12,
    body_rows: 3,
    bg: "navy",
    tc: "FFFFFF",
    mc: "EAEAEA",
    font: "微软雅黑",
    pad_x: 16,
    pad_y: 12,
}


add_clipboard_history() {
    try {
        global clipboard_history, max_history

        content := A_Clipboard

        if (content = "")
            return

        clipboard_history.InsertAt(1, content)

        while (clipboard_history.Length > max_history)
            clipboard_history.Pop()
    } catch
        throw
}

clip_changed(data_type) {
    try {
        global clipboard_history, max_history

        ; 只处理文本类型
        if (data_type != 1)
            return

        content := A_Clipboard
        
        ; 跳过空内容
        if (content = "")
            return


        ; 插入到开头（最新的在前）
        clipboard_history.InsertAt(1, content)

        ; 超出上限则删除最旧的
        while (clipboard_history.Length > max_history)
            clipboard_history.Pop()
    } catch
        throw
}

; --- 保存函数 ---
save_clipboard_history(*) {
    global clipboard_history  
    try {
        clip_map := Map()

        loop 10 {
            if clipboard_history.has(a_index)
                clip_map[a_index] := clipboard_history[a_index]
           else
                clip_map[a_index] := "nothing"
            
        }

        ; 直接覆盖保存
        json_text := json.dumps(clip_map, 2)
        json_file := FileOpen("settings\clipboard_history.json", "w") 
        json_file.Write(json_text)
        json_file.Close() 

    } catch as err {
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
    }    
}

; ------------------------------------------------------------------------------
; 状态
; ------------------------------------------------------------------------------
global _stepClipIdx    := 0
global _stepClipTarget := ""
global _stepClipGuis   := Map()


; ==============================================================================
;  激活显示器
; ==============================================================================
_step_clip_active_monitor() {
    try {
        hwnd := WinGetID("A")
        if !hwnd
            return MonitorGetPrimary()

        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        cx := x + w/2
        cy := y + h/2

        loop MonitorGetCount() {
            MonitorGet(A_Index, &l, &t, &r, &b)
            if (cx >= l && cx < r && cy >= t && cy < b)
                return A_Index
        }
    } catch {
    }
    return MonitorGetPrimary()
}


_step_clip_get_cfg(mon) {
    global STEP_CLIP_CONFIG, STEP_CLIP_DEFAULT
    if STEP_CLIP_CONFIG.Has(mon)
        return STEP_CLIP_CONFIG[mon]
    return STEP_CLIP_DEFAULT
}


_step_clip_calc_pos(mon, cfg) {
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
_step_clip_ensure_gui(mon) {
    global _stepClipGuis

    if _stepClipGuis.Has(mon)
        return _stepClipGuis[mon]

    cfg := _step_clip_get_cfg(mon)

    g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000 -DPIScale",
             "step_clip_m" mon)
    g.BackColor := cfg.bg
    g.MarginX := cfg.pad_x
    g.MarginY := cfg.pad_y

    inner_w := cfg.w - cfg.pad_x * 2

    ; 标题 #N —— 加粗
    g.SetFont("s" cfg.ts " c" cfg.tc " Bold", cfg.font)
    titleCtrl := g.Add("Text", "Left w" inner_w " BackgroundTrans", "#0")

    ; 正文 —— 用 r 预留行数
    g.SetFont("s" cfg.ms " c" cfg.mc " Norm", cfg.font)
    bodyCtrl := g.Add("Text",
        "xm Left w" inner_w " r" cfg.body_rows " BackgroundTrans", "")

    obj := { gui:g, title:titleCtrl, body:bodyCtrl, cfg:cfg, inner_w:inner_w }
    _stepClipGuis[mon] := obj
    return obj
}


; ==============================================================================
;  显示
; ==============================================================================
_step_clip_show(index, text) {
    mon := _step_clip_active_monitor()
    _step_clip_hide_except(mon)

    o := _step_clip_ensure_gui(mon)

    ; 保留原文，不截断、不清洗
    o.title.Text := "#" index
    o.body.Text  := text

    pos := _step_clip_calc_pos(mon, o.cfg)
    o.gui.Show(Format("x{1} y{2} w{3} h{4} NoActivate",
        pos.x, pos.y, o.cfg.w, o.cfg.h))
}


_step_clip_hide_except(keep_mon) {
    global _stepClipGuis
    for mon, o in _stepClipGuis {
        if (mon != keep_mon) {
            try o.gui.Hide()
        }
    }
}

_step_clip_hide_all() {
    global _stepClipGuis
    for mon, o in _stepClipGuis {
        try o.gui.Hide()
    }
}


; ==============================================================================
;  主入口
; ==============================================================================
step_clip() {
    global clipboard_history, _stepClipIdx, _stepClipTarget

    try {
        if (_stepClipIdx == 0)
            SetTimer(_step_clip_waiting, -40)

        if (_stepClipIdx >= 10)
            _stepClipIdx := 1
        else
            _stepClipIdx += 1

        if (!IsSet(clipboard_history) || clipboard_history.Length < _stepClipIdx) {
            _stepClipTarget := ""
            _step_clip_show(_stepClipIdx, "(空)")
            return
        }

        _stepClipTarget := clipboard_history[_stepClipIdx]
        _step_clip_show(_stepClipIdx, _stepClipTarget)

    } catch as err {
        _stepClipIdx := 0
        _step_clip_hide_all()
        MsgBox(Format("step_clip Error: {1}`nFile: {2}`nLine: {3}",
            err.message, err.file, err.Line))
    }
}


_step_clip_waiting() {
    global _stepClipIdx

    try {
        mod := extract_modifier()
        if mod && GetKeyState(mod, "P") {
            SetTimer(_step_clip_waiting, -40)
            return
        }
        _step_clip_commit()
    } catch as err {
        _stepClipIdx := 0
        _step_clip_hide_all()
        MsgBox(Format("step_clip timer Error: {1}", err.message))
    }
}

_step_clip_commit() {
    global _stepClipIdx, _stepClipTarget

    try {
        if (_stepClipTarget != "") {
            old := ClipboardAll()

            try {
                A_Clipboard := _stepClipTarget
                ClipWait(0.5)
                Send("^v")
                Sleep(100)
            } finally {
                ; 还原剪贴板（异步，避免抢在粘贴完成前）
                ; 注意：还原动作也不能触发 clip_changed
                SetTimer(() => _step_clip_restore(old), -500)
            }
        }
    } catch as err {
        MsgBox("step_clip commit failed: " err.Message)
    }

    _stepClipIdx := 0
    _stepClipTarget := ""
    _step_clip_hide_all()
}

_step_clip_restore(old) {
    try {
        A_Clipboard := old
        Sleep(50)
    } 
}


; 判断 step_clip GUI 是否正在显示
step_clip_active() {
    global _stepClipIdx
    return _stepClipIdx > 0
}

; 回退到上一条候选（XButton1 & F）
step_clip_back() {
    global _stepClipIdx, _stepClipTarget, clipboard_history

    ; 已在第 1 条，停住不动
    if (_stepClipIdx <= 1)
        return

    if (clipboard_history.Length < 1)
        return

    _stepClipIdx -= 1

    if (_stepClipIdx > clipboard_history.Length)
        _stepClipIdx := clipboard_history.Length

    _stepClipTarget := clipboard_history[_stepClipIdx]

    ; 与 step_clip() 前进时完全一致
    _step_clip_show(_stepClipIdx, _stepClipTarget)

    ; 重置等待计时器（松开 XButton1 才 commit）
    SetTimer(_step_clip_waiting, -40)
}