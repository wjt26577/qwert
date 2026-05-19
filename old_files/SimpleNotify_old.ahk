
class SimpleNotify {
    
    static cfg := {
        h_lo:     90,         ; ≤ 2000
        h_hi:     140,        ; > 2000
        bg:       '0x0a8232',
        fg:       'White',
        font_zh:  'MiSans Light',
        font_en:  'Gilroy Light',
        font:     'Microsoft YaHei UI',
        ts_lo:    30,
        ts_hi:    34,
        pad_x:    16,
        pad_y:    20,
        dur:      3000
    }

    static themes := Map(
        'info',    '0x008080',
        ; 'info',    'navy',
        'success', '0x1E7E34',
        'warn',    '0xE67E22',
        'error',   '0xC0392B'
    )

    static guis := Map()

    ; ==================== 公开 ====================


    static show(msg, type := 'info', dur := 0) {
        mon := this._active_monitor()
        o := this._ensure_gui(mon)

        bg := this.themes.Has(type) ? this.themes[type] : this.cfg.bg
        o.gui.BackColor := bg

        font := RegExMatch(msg, "[一-龥]") ? this.cfg.font_zh : this.cfg.font_en
        pos := this._calc_pos(mon)

        ; ★ 用显示器宽度判断更合理；或者用短边
        hi := pos.w > 2000
        fontSize := hi ? this.cfg.ts_hi : this.cfg.ts_lo
        barH     := hi ? this.cfg.h_hi  : this.cfg.h_lo

        ; ★ 按屏幕比例自适应（推荐）
        ; barH     := Round(pos.h * 0.08)       ; 屏高 8%
        ; fontSize := Round(barH * 0.38)        ; bar 高度的 38%

        try o.text.SetFont('s' fontSize ' c' this.cfg.fg ' Norm q5', font)
        catch
            o.text.SetFont('s' fontSize ' c' this.cfg.fg ' Norm q5', this.cfg.font)

        o.text.Text := msg
        ; ★ 先 Move 再 Show，确保尺寸生效
        o.text.Move(0, 0, pos.w, barH)
        o.text.Redraw()

        o.gui.Show('x' pos.x ' y' pos.y ' w' pos.w ' h' barH ' NoActivate')

        SetTimer(o.timer, 0)
        SetTimer(o.timer, -(dur > 0 ? dur : this.cfg.dur))
    }

    static hide_all() {
        for mon, o in this.guis {
            SetTimer(o.timer, 0)
            try o.gui.Hide()
        }
    }


    ; ==================== 内部 ====================

    static _active_monitor() {
        CoordMode('Mouse', 'Screen')
        MouseGetPos(&mx, &my)
        loop MonitorGetCount() {
            MonitorGet(A_Index, &l, &t, &r, &b)
            if (mx >= l && mx < r && my >= t && my < b)
                return A_Index
        }
        return MonitorGetPrimary()
    }

    
    ; ★ 改动：贴顶 + 返回显示器宽度
    static _calc_pos(mon) {
        try MonitorGetWorkArea(mon, &l, &t, &r, &b)
        catch
            MonitorGetWorkArea(MonitorGetPrimary(), &l, &t, &r, &b)

        w := r - l
        h := b - t
        bas := Max(w, h)

        return { x: l, y: t, w: w, h: h, bas: bas}
    }

    ; ★ 改动：初始宽度用占位值，show() 里会动态调整
    static _ensure_gui(mon) {
        if this.guis.Has(mon)
            return this.guis[mon]

        c := this.cfg

        g := Gui('+AlwaysOnTop -Caption +ToolWindow +E0x08000000 -DPIScale',
                 'SimpleNotify_m' mon)
        g.BackColor := c.bg
        g.MarginX := 0
        g.MarginY := 0

        g.SetFont('s' c.ts_lo ' c' c.fg ' Norm q5', c.font)
        text := g.Add('Text',
            Format('x0 y0 w{} h{} Center +0x200 +0x8000 BackgroundTrans',
                   1920, c.h_lo),          ; 占位宽度，show 里会 Move 覆盖
            '')

        text.OnEvent('Click', (*) => g.Hide())

        obj := { gui: g, text: text, timer: '' }
        obj.timer := (*) => (g.Hide())
        this.guis[mon] := obj
        return obj
    }
}
