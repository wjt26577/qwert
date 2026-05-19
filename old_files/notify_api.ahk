
; ==============================================================================
;  notify_api.ahk
;
;  基于 Notify 库 (by XMCQCX) 的统一通知封装层。
;
;  设计原则：
;     1. 业务代码调用方式 100% 兼容旧版（center_info / left_info / clip_info / show_info_*）
;     2. 所有通知都带 tag，禁止使用 DestroyAll（防止跨脚本干扰）
;     3. 用 Notify 自带的主题系统替代手工拼 layout 字符串
;     4. 提供降级模式（NOTIFY_SAFE_MODE = true → 全部退回 ToolTip）
;     5. 提供“预览型通知”（同 tag 自动替换），适合剪贴板候选、快捷键提示等场景
;     6. 提供“状态型通知”（dur=0 dgc=0 dgb=1），用于超级模式、维护模式等常驻指示
;
;  依赖：
;     - Notify 类（你已有）
;     - 全局数组：ch_fonts, en_fonts（你已有）
;
;  使用：
;     #Include <Notify>             ; 第三方底层库
;     #Include lib\notify_api.ahk   ; 本封装层
;
;  作者：你自己
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; 全局：安全降级开关
; ------------------------------------------------------------------------------
; 维护 / 调试时设为 true → 全部 ToolTip，不再调用 Notify 库
global NOTIFY_SAFE_MODE := false


; ------------------------------------------------------------------------------
; Tag 约定（所有业务通知都必须有 tag，禁止使用 DestroyAll）
; ------------------------------------------------------------------------------
class NT {
    static CENTER     := "qw_center"       ; 普通中央提示
    static LEFT       := "qw_left"         ; 左侧靠窗提示
    static CLIP       := "qw_clip"         ; 剪贴板面板
    static HOTKEY     := "qw_hotkey"       ; 显示当前热键
    static PREVIEW    := "qw_preview"      ; dispatcher/快捷预览
    static STATUS     := "qw_status"       ; 常驻状态指示
    static WARN       := "qw_warn"
    static ERR        := "qw_err"
    static OK         := "qw_ok"
    static INFO       := "qw_info"
}


; ==============================================================================
;                       第 1 层：底层适配（最稳，不抛异常）
; ==============================================================================

/**
 * 直接调用 Notify.Show，失败则回退 ToolTip
 */
notify_raw(text, options := "") {
    if NOTIFY_SAFE_MODE {
        _tip_fallback(text)
        return
    }

    try {
        Notify.Show(text,,,,, options)
    } catch {
        _tip_fallback(text)
    }
}

/**
 * 按 tag 销毁通知（永远不要用 DestroyAll！）
 */
notify_destroy(tag, force := true) {
    try Notify.Destroy(tag, force)
}

/**
 * 判断某个 tag 的通知是否存在
 */
notify_exists(tag) {
    try return Notify.Exist(tag)
    return false
}

/**
 * ToolTip 降级显示
 */
_tip_fallback(text, ms := 1200) {
    ToolTip(text)
    SetTimer () => ToolTip(), -Abs(ms)
}


; ==============================================================================
;                     第 2 层：主题注册（一次性 setup）
; ==============================================================================
; 通过 Notify 自带的 mThemes 机制注册“语义主题”。
; 以后调整字体 / 颜色 / 布局，只改这里，不动业务函数。
; ==============================================================================

_register_qw_themes() {
    global ch_fonts, en_fonts

    ; 选字体（带 fallback）
    cn_font := (IsSet(ch_fonts) && ch_fonts.Length >= 5) ? ch_fonts[5] : "微软雅黑"
    en_font := (IsSet(en_fonts) && en_fonts.Length >= 2) ? en_fonts[2] : "Segoe UI"

    ; 中央提示：顶部居中、小条、自动消失
    _add_theme("qw_center", Map(
        "style", "Edge",
        "pos",   "TC",
        "tali",  "Center",
        "width", 350,
        "tf",    cn_font,
        "mf",    cn_font,
        "tc",    "0xFFFFFF",
        "bc",    "0x084706",
        "dur",   1,
        "dg",    0,
        "pad",   ",,12,12,8,8",
        "show",  "None",
        "hide",  "None"
    ))

    ; 左侧贴控件
    _add_theme("qw_left", Map(
        "style", "Edge",
        "pos",   "CTL",
        "tali",  "Left",
        "width", 200,
        "tf",    "微软雅黑",
        "mf",    "微软雅黑",
        "tc",    "0xFFFFFF",
        "bc",    "0x084706",
        "dur",   10,
        "dg",    0,
        "pad",   ",,12,12,16,16",
        "show",  "None",
        "hide",  "None"
    ))

    ; 剪贴板面板
    _add_theme("qw_clip", Map(
        "style", "Edge",
        "pos",   "TL",
        "tali",  "Left",
        "width", 400,
        "tf",    cn_font,
        "mf",    cn_font,
        "bc",    "0x084706",
        "tc",    "0xFFFFFF",
        "dur",   100,
        "dg",    0,
        "pad",   ",,12,12,16,16",
        "show",  "none",
        "hide",  "none"
    ))

    ; 语义快捷：错误
    _add_theme("qw_err", Map(
        "style","Edge","pos","TC","tali","Center","width",350,
        "bc","Maroon","tc","0xFFFFFF","dur",2,
        "tf", cn_font, "mf", cn_font,
        "pad",",,12,12,8,8"
    ))

    ; 警告
    _add_theme("qw_warn", Map(
        "style","Edge","pos","TC","tali","Center","width",350,
        "bc","Yellow","tc","0x555555","dur",2,
        "tf", cn_font, "mf", cn_font,
        "pad",",,12,12,8,8"
    ))

    ; 成功
    _add_theme("qw_ok", Map(
        "style","Edge","pos","TC","tali","Center","width",350,
        "bc","0x005a00","tc","0xFFFFFF","dur",1,
        "tf", cn_font, "mf", cn_font,
        "pad",",,12,12,8,8"
    ))

    ; 普通信息
    _add_theme("qw_info", Map(
        "style","Edge","pos","TC","tali","Center","width",350,
        "bc","Navy","tc","0xFFFFFF","dur",1.5,
        "tf", cn_font, "mf", cn_font,
        "pad",",,12,12,8,8"
    ))

    ; 常驻状态（不自动消失，不被点击销毁，不被 DG 影响）
    _add_theme("qw_status", Map(
        "style","Edge","pos","TR","tali","Left","width",220,
        "bc","0x1F1F1F","tc","0xEAEAEA","dur",0,"dgc",0,"dgb",1,
        "tf", cn_font, "mf", cn_font,
        "pad",",,10,10,10,10"
    ))
}

_add_theme(name, m) {
    try {
        Notify.mThemes[name] := Notify.MapCI()
        for k, v in m
            Notify.mThemes[name][k] := v

        ; 按源码要求补 arrKeyDefined
        arr := []
        for k in m
            arr.Push(k)
        Notify.mThemes[name]["arrKeyDefined"] := arr
    } catch {
        ; 静默失败，降级模式不依赖主题
    }
}

; 立即注册
_register_qw_themes()


; ==============================================================================
;                     第 3 层：内部工具（文本预处理）
; ==============================================================================

_prepare_text(info) {
    info := StrReplace(info, "&", "&&")
    len  := StrLen(info)

    if (len > 30)
        info := SubStr(info, 1, 30) . "..."

    has_cn := RegExMatch(info, "[一-龥]") > 0
    if (has_cn && len < 10)
        info := Trim(RegExReplace(info, "(.)", "$1 "))

    return info
}

_build_opts(theme, overrides := "") {
    s := "theme=" theme
    if (overrides != "")
        s .= " " overrides
    return s
}


; ==============================================================================
;                     第 4 层：业务 API（对外接口）
;  完全保留你原有的函数签名，内部改走主题 + tag
; ==============================================================================

; --- 基础 ToolTip ---------------------------------------------------------------
tool_tip(info, time := 800) {
    _tip_fallback(info, time)
}


; --- 中央提示 ------------------------------------------------------------------
center_info(info := "", color := "0x084706", dur := 1, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.CENTER

        ; 同 tag 先销毁，实现“刷新式通知”
        notify_destroy(real_tag)

        overrides := Format("bc={1} dur={2} tag={3}", color, dur, real_tag)
        notify_raw(text, _build_opts("qw_center", overrides))
    } catch {
        _tip_fallback(info)
    }
}


; --- 左侧提示 ------------------------------------------------------------------
left_info(info := "", color := "0x084706", dur := 10, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.LEFT
        notify_destroy(real_tag)

        overrides := Format("bc={1} dur={2} tag={3}", color, dur, real_tag)
        notify_raw(text, _build_opts("qw_left", overrides))
    } catch {
        _tip_fallback(info)
    }
}


; --- 剪贴板信息条 --------------------------------------------------------------
; 注意：clip_info 常常是多条同时出现，所以 tag 不做自动销毁
clip_info(info := "", color := "0x084706", dur := 100, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.CLIP

        overrides := Format("bc={1} dur={2} tag={3}", color, dur, real_tag)
        notify_raw(text, _build_opts("qw_clip", overrides))
    } catch {
        _tip_fallback(info)
    }
}


; --- 语义快捷（兼容旧函数名） --------------------------------------------------
show_info_yellow(info := "", dur := -1, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.WARN
        notify_destroy(real_tag)
        overrides := (dur != -1 ? "dur=" dur " " : "") "tag=" real_tag
        notify_raw(text, _build_opts("qw_warn", overrides))
    } catch {
        _tip_fallback(info)
    }
}

show_info_orange(info := "", dur := -1, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.WARN
        notify_destroy(real_tag)
        overrides := Format("bc=0x9b4901 tc=0xFFFFFF {1}tag={2}"
            , (dur != -1 ? "dur=" dur " " : "")
            , real_tag)
        notify_raw(text, _build_opts("qw_center", overrides))
    } catch {
        _tip_fallback(info)
    }
}

show_info_red(info := "", dur := -1, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.ERR
        notify_destroy(real_tag)
        overrides := (dur != -1 ? "dur=" dur " " : "") "tag=" real_tag
        notify_raw(text, _build_opts("qw_err", overrides))
    } catch {
        _tip_fallback(info)
    }
}

show_info_blue(info := "", dur := -1, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.INFO
        notify_destroy(real_tag)
        overrides := (dur != -1 ? "dur=" dur " " : "") "tag=" real_tag
        notify_raw(text, _build_opts("qw_info", overrides))
    } catch {
        _tip_fallback(info)
    }
}

show_info_purple(info := "", dur := -1, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.INFO
        notify_destroy(real_tag)
        overrides := Format("bc=Purple tc=0xFFFFFF {1}tag={2}"
            , (dur != -1 ? "dur=" dur " " : "")
            , real_tag)
        notify_raw(text, _build_opts("qw_center", overrides))
    } catch {
        _tip_fallback(info)
    }
}

show_info_green(info := "", dur := -1, tag := "") {
    try {
        text := _prepare_text(info)
        real_tag := tag != "" ? tag : NT.OK
        notify_destroy(real_tag)
        overrides := (dur != -1 ? "dur=" dur " " : "") "tag=" real_tag
        notify_raw(text, _build_opts("qw_ok", overrides))
    } catch {
        _tip_fallback(info)
    }
}


; --- 调试：显示当前热键 --------------------------------------------------------
do_nothing() {
    try {
        notify_destroy(NT.HOTKEY)
        overrides := "bc=Maroon tc=0xFFFFFF dur=1 tag=" NT.HOTKEY
        notify_raw(A_ThisHotkey, _build_opts("qw_center", overrides))
    } catch {
        _tip_fallback(A_ThisHotkey)
    }
}


; ==============================================================================
;                     第 5 层：进阶 API（建议开始用这些）
; ==============================================================================

/**
 * “预览型”通知：用同一个 tag，多次调用会自动刷新同一块 UI。
 * 非常适合：快捷键预览、剪贴板候选、dispatcher 切换。
 *
 * 示例：
 *   preview_info("粘贴 #1", "0x084706")
 *   preview_info("粘贴 #2", "0x084706")   ; 会替换上一条
 */
preview_info(text, color := "0x084706", dur := 0.8, tag := "") {
    try {
        real_tag := tag != "" ? tag : NT.PREVIEW
        notify_destroy(real_tag)
        overrides := Format("bc={1} dur={2} tag={3}", color, dur, real_tag)
        notify_raw(_prepare_text(text), _build_opts("qw_center", overrides))
    } catch {
        _tip_fallback(text)
    }
}

/**
 * “状态型”通知：不会自动消失，不会被点击销毁，不会被 DG 影响。
 * 典型用途：super_idol 模式、维护模式、CapsLock 指示。
 *
 *   status_info("模式：增强", "0x1F1F1F", "mode")
 *   ...
 *   status_dismiss("mode")
 */
status_info(text, color := "0x1F1F1F", tag := "") {
    try {
        real_tag := tag != "" ? tag : NT.STATUS
        notify_destroy(real_tag)
        overrides := Format("bc={1} tag={2}", color, real_tag)
        notify_raw(_prepare_text(text), _build_opts("qw_status", overrides))
    } catch {
        _tip_fallback(text)
    }
}

status_dismiss(tag := "") {
    real_tag := tag != "" ? tag : NT.STATUS
    notify_destroy(real_tag, true)   ; force，因为状态通知有 dgb=1
}


; ==============================================================================
;                     第 6 层：维护模式（强烈推荐接入）
; ==============================================================================
/**
 * 一键进入降级模式（调试时用）：
 *   所有通知改走 ToolTip，不会再依赖 Notify 库。
 */
notify_enter_safe_mode() {
    global NOTIFY_SAFE_MODE
    NOTIFY_SAFE_MODE := true
    ToolTip("NOTIFY SAFE MODE: ON")
    SetTimer () => ToolTip(), -1200
}

notify_exit_safe_mode() {
    global NOTIFY_SAFE_MODE
    NOTIFY_SAFE_MODE := false
    ToolTip("NOTIFY SAFE MODE: OFF")
    SetTimer () => ToolTip(), -1200
}








































































; ; #region show_info

; tool_tip(info, time := 800) {
;     ToolTip info
;     SetTimer () => ToolTip(), -time
; }

; ; 'theme', 'Default',
; ; 'mon', 'Primary',
; ; 'pos', 'BR',
; ; 'dur', 8,
; ; 'style', 'Round',
; ; 'ts', 15,
; ; 'tc', 'White',
; ; 'tf', 'Segoe UI',
; ; 'tfo', 'norm Bold',
; ; 'tali', 'Left',
; ; 'ms', 12,
; ; 'mc', '0xEAEAEA',
; ; 'mf', 'Segoe UI',
; ; 'mfo', 'norm',
; ; 'mali', 'Left',
; ; 'bc', '0x1F1F1F',
; ; 'bdr', 'Default',
; ; 'sound', 'None',
; ; 'image', 'None',
; ; 'iw', 32,
; ; 'ih', -1,
; ; 'bgImg', 'None',
; ; 'bgImgPos', 'Stretch',
; ; 'pad', ',,16,16,16,16,8,10',
; ; 'width', '',
; ; 'minH', '',
; ; 'maxW', '',
; ; 'prog', '',
; ; 'tag', '',
; ; 'dgc', 1,
; ; 'dg', 0,
; ; 'dgb', 0,
; ; 'dga', 0,
; ; 'wstc', '',
; ; 'wstp', '',
; ; 'opt', '+Owner -Caption +AlwaysOnTop +E0x08000000',



; center_info(info := "", color := "084706", dur := 1, tag := "") {
;     global ch_fonts
;     global en_fonts
;     try {        
;         info := StrReplace(info, "&", "&&")
;         ; info := "            " . info . "            "
;         length := StrLen(info)   
        
;         if length > 30
;             info := SubStr(info, 1, 30) . "..."
      
;         has_chinese := RegExMatch(info, "[一-龥]") > 0

;         if has_chinese && length < 10
;             info := Trim(RegExReplace(info, "(.)", "$1 "))

;         if has_chinese && ch_fonts[5]     
;             font := ch_fonts[5]
;         else if !has_chinese && en_fonts[2]
;             font := en_fonts[2]
;         else
;             font := "微软雅黑"
                
;         if (color ~= "i)^Yellow|Red|Aqua") {
;             tc := "0x555555"
;         } else {
;             tc := "0xFFFFFF"
;         }   
        
        
;         layout := Format("tf={1} tc={2} bc={3} dur={4} tag={5}", font, tc, color, dur, tag)
;         layout := layout . " show=None hide=None tali=Center width=350 tfo=norm dg=1 style=Edge pos=TC pad=,,12,12,8,8"
;         Notify.Show(info,,,,, layout)
;     }
;     catch
;         throw
; }


; left_info(info := "", color := "084706", dur := 10, tag := "") {
;     global ch_fonts, en_fonts
;     try {        
;         info := StrReplace(info, "&", "&&")
;         ; info := "            " . info . "            "
;         length := StrLen(info)   
        
;         if length > 20
;             info := SubStr(info, 1, 20) . "..."
      
;         has_chinese := RegExMatch(info, "[一-龥]") > 0

;         if has_chinese && length < 10
;             info := Trim(RegExReplace(info, "(.)", "$1 "))

;         ; if has_chinese && ch_fonts[5]     
;         ;     font := ch_fonts[5]
;         ; else if !has_chinese && en_fonts[2]
;         ;     font := en_fonts[2]
;         ; else
;             font := "微软雅黑"
                
;         if (color ~= "i)^Yellow|Red|Aqua") {
;             tc := "0x555555"
;         } else {
;             tc := "0xFFFFFF"
;         }   
;         dur := 10
        
;         layout := Format("tf={1} tc={2} bc={3} dur={4} tag={5}", font, tc, color, dur, tag)
;         layout := layout . " show=None hide=None tali=left width=200 tfo=norm dg=0 style=Edge pos=ctl pad=,,12,12,16,16"
;         Notify.Show(info,,,,, layout)
;     }
;     catch
;         throw
; }


; clip_info(info := "", color := "084706", dur := 100, tag := "clipboard_history") {
;     global ch_fonts, en_fonts
;     try {        
;         info := StrReplace(info, "&", "&&")
             
;         has_chinese := RegExMatch(info, "[一-龥]") > 0

;         if has_chinese && ch_fonts[5]     
;             font := ch_fonts[5]
;         else if !has_chinese && en_fonts[2]
;             font := en_fonts[2]
;         else
;             font := "微软雅黑"
             
;         layout := Format("tf={} bc={} dur={} tag={}", font, color, dur, tag)
;         layout := layout . " show=none hide=none tali=left width=400 tfo=norm dg=0 style=Edge pos=tl pad=,,12,12,16,16"
;         Notify.Show(info,,,,, layout)
;     }
;     catch
;         throw
; }




; show_info_yellow(info:="", dur := -1, tag := "") {
;     center_info(info, "Yellow", dur, tag) 
; }

; show_info_orange(info:="", dur := -1, tag := "") {
;     center_info(info, "0x9b4901", dur, tag) 
; }

; show_info_red(info:="", dur := -1, tag := "") {
;     center_info(info, "Maroon", dur, tag) 
; }

; show_info_blue(info:="", dur := -1, tag := "") {
;     center_info(info, "Navy", dur, tag) 
; }

; show_info_purple(info:="", dur := -1, tag := "") {
;     center_info(info, "Purple", dur, tag) 
; }

; show_info_green(info:="", dur := -1, tag := "") {
;     center_info(info, "005a00", dur, tag) 
; }

; do_nothing() {
;     center_info(A_ThisHotkey, "Maroon")	
; }






; ; #endregion
