class Notify {

    static cfg := {
        bg:       '0x0a8232',
        fg:       'White',
        font_zh:  'MiSans Light',
        font_en:  'Gilroy Light',
        font:     'Microsoft YaHei UI',
        dur:      3000
    }

    static themes := Map(
        'info',    '0x0da869',
        'success', '0x07136e',
        'warn',    '0xE67E22',
        'error',   '0xC0392B'
    )

    static guis := Map()

    ; ==================== 公开 ====================

    static show(msg, type := 'info', dur := 0) {
        ; 显示位置：鼠标在哪个显示器，就显示在哪个显示器
        mon := this._active_monitor()
        o   := this._ensure_gui(mon)
        pos := this._calc_pos(mon)

        ; 尺寸 / 字号 / DPI preset：只按主显示器判断
        p := DpiPresetStore.PickForPrimary()

        bg := this.themes.Has(type) ? this.themes[type] : this.cfg.bg
        o.gui.BackColor := bg

        font := RegExMatch(msg, "[一-龥]") ? this.cfg.font_zh : this.cfg.font_en

        try o.text.SetFont('s' p.ts ' c' this.cfg.fg ' Norm q5', font)
        catch
            o.text.SetFont('s' p.ts ' c' this.cfg.fg ' Norm q5', this.cfg.font)

        o.text.Text := msg

        o.text.Move(0, 0, pos.w, p.h)
        o.text.Redraw()

        o.gui.Show('x' pos.x ' y' pos.y ' w' pos.w ' h' p.h ' NoActivate')
        this._force_topmost(o.gui)
        SetTimer(() => this._force_topmost(o.gui), -30)

        SetTimer(o.timer, 0)
        SetTimer(o.timer, -(dur > 0 ? dur : this.cfg.dur))
    }

    static hide_all() {
        for mon, o in this.guis {
            SetTimer(o.timer, 0)
            try o.gui.Hide()
        }
    }

    static _calc_pos(mon) {
        try MonitorGetWorkArea(mon, &l, &t, &r, &b)
        catch
            MonitorGetWorkArea(MonitorGetPrimary(), &l, &t, &r, &b)

        return { x: l, y: t, w: r - l, h: b - t }
    }

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

    static _force_topmost(g) {
        try DllCall(
            "SetWindowPos",
            "Ptr", g.Hwnd,
            "Ptr", -1,
            "Int", 0,
            "Int", 0,
            "Int", 0,
            "Int", 0,
            "UInt", 0x0001 | 0x0002 | 0x0010
        )
    }

    static _ensure_gui(mon) {
        if this.guis.Has(mon)
            return this.guis[mon]

        c := this.cfg

        g := Gui('+AlwaysOnTop -Caption +ToolWindow +E0x08000000 -DPIScale',
                 'Notify_m' mon)

        g.BackColor := c.bg
        g.MarginX := 0
        g.MarginY := 0
        g.SetFont('s24 c' c.fg ' Norm q5', c.font)

        text := g.Add('Text',
            Format('x0 y0 w{} h{} Center +0x200 +0x8000 BackgroundTrans',
                   1920, 90), '')

        text.OnEvent('Click', (*) => g.Hide())

        obj := { gui: g, text: text, timer: '' }
        obj.timer := (*) => g.Hide()

        this.guis[mon] := obj
        return obj
    }
}


class DpiPresetStore {

    static Path() {
        return A_ScriptDir "\config\notify_presets.json"
    }

    static Defaults() {
        return [
            { name: "1080p_100",      width: 1920, height: 1080, scale: 100, dpi: 96,  barHeight: 94,  fontSize: 30 },
            { name: "1080p_125",      width: 1920, height: 1080, scale: 125, dpi: 120, barHeight: 100, fontSize: 32 },
            { name: "1440p_100",      width: 2560, height: 1440, scale: 100, dpi: 96,  barHeight: 100, fontSize: 32 },
            { name: "1440p_125",      width: 2560, height: 1440, scale: 125, dpi: 120, barHeight: 120, fontSize: 36 },
            { name: "UWQHD_100",      width: 3440, height: 1440, scale: 100, dpi: 96,  barHeight: 105, fontSize: 32 },
            { name: "4K_150",         width: 3840, height: 2160, scale: 150, dpi: 144, barHeight: 140, fontSize: 36 },
            { name: "4K_175",         width: 3840, height: 2160, scale: 175, dpi: 168, barHeight: 155, fontSize: 40 },
            { name: "4K_200",         width: 3840, height: 2160, scale: 200, dpi: 192, barHeight: 170, fontSize: 42 },
            { name: "5120x1440_100",  width: 5120, height: 1440, scale: 100, dpi: 96,  barHeight: 110, fontSize: 34 }
        ]
    }

    static EnsureFile() {
        path := this.Path()

        if FileExist(path)
            return

        this.Save(this.Defaults())
    }

    static Load() {
        this.EnsureFile()

        path := this.Path()

        try txt := FileRead(path, "UTF-8")
        catch {
            return this.Defaults()
        }

        arr := this.ParseJson(txt)

        if !arr.Length
            arr := this.Defaults()

        return arr
    }

    static Save(arr) {
        path := this.Path()
        dir := RegExReplace(path, "\\[^\\]+$")

        if !DirExist(dir)
            DirCreate(dir)

        s := "{`n"
        s .= "  `"version`": 1,`n"
        s .= "  `"presets`": [`n"

        for i, p in arr {
            comma := i < arr.Length ? "," : ""

            s .= "    {`n"
            s .= "      `"name`": `"" this.Esc(p.name) "`",`n"
            s .= "      `"width`": " p.width ",`n"
            s .= "      `"height`": " p.height ",`n"
            s .= "      `"scale`": " p.scale ",`n"
            s .= "      `"dpi`": " p.dpi ",`n"
            s .= "      `"barHeight`": " p.barHeight ",`n"
            s .= "      `"fontSize`": " p.fontSize "`n"
            s .= "    }" comma "`n"
        }

        s .= "  ]`n"
        s .= "}`n"

        f := FileOpen(path, "w", "UTF-8")
        f.Write(s)
        f.Close()
    }

    static Esc(s) {
        s := "" s
        s := StrReplace(s, "\", "\\")
        s := StrReplace(s, '"', '\"')
        s := StrReplace(s, "`r", "\r")
        s := StrReplace(s, "`n", "\n")
        return s
    }

    static Unesc(s) {
        s := StrReplace(s, "\n", "`n")
        s := StrReplace(s, "\r", "`r")
        s := StrReplace(s, '\"', '"')
        s := StrReplace(s, "\\", "\")
        return s
    }

    ; 简单解析本程序生成的 JSON，不是通用 JSON 解析器
    static ParseJson(txt) {
        arr := []
        pos := 1

        while RegExMatch(txt, '\{[^{}]*"name"[^{}]*\}', &m, pos) {
            obj := m[0]
            pos := m.Pos(0) + m.Len(0)

            p := {
                name:      this.GetStr(obj, "name", ""),
                width:     this.GetNum(obj, "width", 0),
                height:    this.GetNum(obj, "height", 0),
                scale:     this.GetNum(obj, "scale", 100),
                dpi:       this.GetNum(obj, "dpi", 96),
                barHeight: this.GetNum(obj, "barHeight", 90),
                fontSize:  this.GetNum(obj, "fontSize", 30)
            }

            if p.width && p.height && p.dpi
                arr.Push(p)
        }

        return arr
    }

    static GetStr(obj, key, def := "") {
        pat := '"' key '"\s*:\s*"((?:\\.|[^"\\])*)"'

        if RegExMatch(obj, pat, &m)
            return this.Unesc(m[1])

        return def
    }

    static GetNum(obj, key, def := 0) {
        pat := '"' key '"\s*:\s*(-?\d+(?:\.\d+)?)'

        if RegExMatch(obj, pat, &m)
            return m[1] + 0

        return def
    }

    static GetDpiForMonitor(mon) {
        try {
            MonitorGet(mon, &l, &t, &r, &b)

            rect := Buffer(16, 0)
            NumPut("Int", l, "Int", t, "Int", r, "Int", b, rect)

            hMon := DllCall("MonitorFromRect", "Ptr", rect, "UInt", 2, "Ptr")

            DllCall(
                "Shcore\GetDpiForMonitor",
                "Ptr", hMon,
                "Int", 0,
                "UInt*", &dpiX := 0,
                "UInt*", &dpiY := 0
            )

            return dpiX ? dpiX : 96
        } catch {
            return 96
        }
    }

    static PickForPrimary() {
        presets := this.Load()

        mon := MonitorGetPrimary()
        MonitorGet(mon, &l, &t, &r, &b)

        w := r - l
        h := b - t
        dpi := this.GetDpiForMonitor(mon)

        best := ""
        bestScore := 999999999

        for p in presets {
            score := Abs(p.width - w) + Abs(p.height - h) + Abs(p.dpi - dpi) * 20

            if score < bestScore {
                bestScore := score
                best := p
            }
        }

        if !IsObject(best)
            best := {
                name: "fallback",
                width: w,
                height: h,
                scale: Round(dpi / 96 * 100),
                dpi: dpi,
                barHeight: 90,
                fontSize: 30
            }

        return {
            name: best.name,
            h: best.barHeight,
            ts: best.fontSize,
            width: best.width,
            height: best.height,
            dpi: best.dpi,
            scale: best.scale
        }
    }
}


class DpiPresetEditor {
    static gui := ""
    static lv := ""
    static nameEdit := ""
    static widthEdit := ""
    static heightEdit := ""
    static scaleBox := ""
    static dpiEdit := ""
    static barEdit := ""
    static fontEdit := ""

    static Show() {
        if IsObject(this.gui) {
            this.RefreshList()
            this.gui.Show()
            return
        }

        g := Gui("+Resize", "显示器 DPI / 通知条 Preset 编辑器")
        g.SetFont("s10", "Microsoft YaHei UI")

        this.lv := g.Add(
            "ListView",
            "x10 y10 w820 h250 Grid -Multi",
            ["名称", "宽", "高", "缩放%", "DPI", "通知条高", "字号"]
        )

        this.lv.OnEvent("ItemSelect", (*) => this.LoadSelectedRow())

        g.Add("Text", "x10 y280 w80 h24", "名称")
        this.nameEdit := g.Add("Edit", "x90 y276 w180 h26")

        g.Add("Text", "x290 y280 w50 h24", "宽")
        this.widthEdit := g.Add("Edit", "x340 y276 w80 h26 Number")

        g.Add("Text", "x440 y280 w50 h24", "高")
        this.heightEdit := g.Add("Edit", "x490 y276 w80 h26 Number")

        g.Add("Text", "x590 y280 w60 h24", "缩放%")
        this.scaleBox := g.Add("ComboBox", "x650 y276 w90 h26", ["100", "125", "150", "175", "200", "225", "250", "300"])
        this.scaleBox.Text := "100"

        g.Add("Text", "x10 y320 w80 h24", "DPI")
        this.dpiEdit := g.Add("Edit", "x90 y316 w80 h26 Number")
        this.dpiEdit.Value := "96"

        g.Add("Text", "x190 y320 w80 h24", "通知条高")
        this.barEdit := g.Add("Edit", "x270 y316 w80 h26 Number")
        this.barEdit.Value := "94"

        g.Add("Text", "x370 y320 w80 h24", "字号")
        this.fontEdit := g.Add("Edit", "x450 y316 w80 h26 Number")
        this.fontEdit.Value := "30"

        btnCurrent := g.Add("Button", "x10 y365 w140 h34", "读取主显示器")
        btnAdd     := g.Add("Button", "x160 y365 w130 h34", "新增 / 更新")
        btnDelete  := g.Add("Button", "x300 y365 w100 h34", "删除")
        btnSave    := g.Add("Button", "x410 y365 w100 h34", "保存 JSON")
        btnOpen    := g.Add("Button", "x520 y365 w120 h34", "打开 JSON")

        btnCurrent.OnEvent("Click", (*) => this.UsePrimaryMonitor())
        btnAdd.OnEvent("Click", (*) => this.AddOrUpdate())
        btnDelete.OnEvent("Click", (*) => this.DeleteSelected())
        btnSave.OnEvent("Click", (*) => this.SaveJson())
        btnOpen.OnEvent("Click", (*) => Run(DpiPresetStore.Path()))

        this.scaleBox.OnEvent("Change", (*) => this.ScaleChanged())

        g.OnEvent("Close", (*) => g.Hide())

        this.gui := g
        this.RefreshList()

        g.Show("w850 h420")
    }

    static RefreshList() {
        presets := DpiPresetStore.Load()

        this.lv.Delete()

        for p in presets {
            this.lv.Add(
                "",
                p.name,
                p.width,
                p.height,
                p.scale,
                p.dpi,
                p.barHeight,
                p.fontSize
            )
        }

        Loop 7
            this.lv.ModifyCol(A_Index, "AutoHdr")
    }

    static LoadSelectedRow() {
        row := this.lv.GetNext()

        if !row
            return

        this.nameEdit.Value   := this.lv.GetText(row, 1)
        this.widthEdit.Value  := this.lv.GetText(row, 2)
        this.heightEdit.Value := this.lv.GetText(row, 3)
        this.scaleBox.Text    := this.lv.GetText(row, 4)
        this.dpiEdit.Value    := this.lv.GetText(row, 5)
        this.barEdit.Value    := this.lv.GetText(row, 6)
        this.fontEdit.Value   := this.lv.GetText(row, 7)
    }

    static ScaleChanged() {
        scale := this.scaleBox.Text + 0

        if scale > 0
            this.dpiEdit.Value := Round(96 * scale / 100)
    }

    static UsePrimaryMonitor() {
        monIndex := MonitorGetPrimary()

        MonitorGet(monIndex, &l, &t, &r, &b)

        w := r - l
        h := b - t
        dpi := DpiPresetStore.GetDpiForMonitor(monIndex)
        scale := Round(dpi / 96 * 100)

        this.nameEdit.Value   := Format("{}x{}_{}%", w, h, scale)
        this.widthEdit.Value  := w
        this.heightEdit.Value := h
        this.scaleBox.Text    := scale
        this.dpiEdit.Value    := dpi

        this.barEdit.Value  := Round(94 * dpi / 96)
        this.fontEdit.Value := Round(30 * dpi / 96)
    }

    static AddOrUpdate() {
        p := this.InputToPreset()

        if !p.width || !p.height || !p.dpi {
            MsgBox("宽、高、DPI 不能为空。")
            return
        }

        row := this.lv.GetNext()

        if row {
            this.lv.Modify(
                row,
                "",
                p.name,
                p.width,
                p.height,
                p.scale,
                p.dpi,
                p.barHeight,
                p.fontSize
            )
        } else {
            this.lv.Add(
                "",
                p.name,
                p.width,
                p.height,
                p.scale,
                p.dpi,
                p.barHeight,
                p.fontSize
            )
        }

        Loop 7
            this.lv.ModifyCol(A_Index, "AutoHdr")
    }

    static DeleteSelected() {
        row := this.lv.GetNext()

        if row
            this.lv.Delete(row)
    }

    static SaveJson() {
        arr := []

        loop this.lv.GetCount() {
            arr.Push({
                name:      this.lv.GetText(A_Index, 1),
                width:     this.lv.GetText(A_Index, 2) + 0,
                height:    this.lv.GetText(A_Index, 3) + 0,
                scale:     this.lv.GetText(A_Index, 4) + 0,
                dpi:       this.lv.GetText(A_Index, 5) + 0,
                barHeight: this.lv.GetText(A_Index, 6) + 0,
                fontSize:  this.lv.GetText(A_Index, 7) + 0
            })
        }

        DpiPresetStore.Save(arr)

        MsgBox("已保存：`n" DpiPresetStore.Path())
    }

    static InputToPreset() {
        name := Trim(this.nameEdit.Value)

        width  := this.widthEdit.Value + 0
        height := this.heightEdit.Value + 0
        scale  := this.scaleBox.Text + 0
        dpi    := this.dpiEdit.Value + 0
        barH   := this.barEdit.Value + 0
        fs     := this.fontEdit.Value + 0

        if name = ""
            name := Format("{}x{}_{}%", width, height, scale)

        return {
            name: name,
            width: width,
            height: height,
            scale: scale,
            dpi: dpi,
            barHeight: barH,
            fontSize: fs
        }
    }
}
