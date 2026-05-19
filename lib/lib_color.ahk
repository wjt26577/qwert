; ============================================================
;  lib_color.ahk — 颜色规范化 / RGB↔BGR / 深浅判断 / 颜色名
; ============================================================

normalize_color(color_input) {
    if (Type(color_input) = "String") {
        s := Trim(color_input)
        if (RegExMatch(s, "i)^0x[0-9a-f]{6}$"))
            return "0x" . StrUpper(SubStr(s, 3))
        if (RegExMatch(s, "i)^0x[0-9a-f]{8}$"))
            return "0x" . StrUpper(SubStr(s, 5))
        if (RegExMatch(s, "i)^#[0-9a-f]{6}$"))
            return "0x" . StrUpper(SubStr(s, 2))
        if (RegExMatch(s, "i)^[0-9a-f]{6}$"))
            return "0x" . StrUpper(s)
        hex := color_name_to_hex(s)
        if (hex != "")
            return hex
        throw Error("无法识别的颜色字符串: " . color_input)
    }
    if (Type(color_input) = "Integer" || Type(color_input) = "Float") {
        num := Integer(color_input)
        if (num < 0 || num > 0xFFFFFFFF)
            throw Error("颜色数值超出范围: " . color_input)
        return Format("0x{:06X}", num & 0xFFFFFF)
    }
    throw Error("不支持的颜色类型: " . Type(color_input))
}

color_to_int(color_input) => Integer(normalize_color(color_input))
color_to_str(color_input) => normalize_color(color_input)

is_dark_color(color_input) {
    color_int := color_to_int(color_input)
    r := (color_int >> 16) & 0xFF
    g := (color_int >> 8) & 0xFF
    b := color_int & 0xFF
    return (r * 0.299 + g * 0.587 + b * 0.114) < 128
}

get_contrast_color(color_input) {
    return is_dark_color(color_input) ? "0xFFFFFF" : "0x000000"
}

rgb2bgr(color, g := -1, b := -1) {
    try {
        if (g == -1 || b == -1) {
            for color_name in system_color_map
                if StrTitle(color) == color_name
                    color := system_color_map[color_name]
            r := (color >> 16) & 0xFF
            g := (color >> 8) & 0xFF
            b := color & 0xFF
        }
        return (r << 0) | (g << 8) | (b << 16)
    } catch
        throw
}

bgr2rgb(color, g := -1, b := -1) {
    try {
        b := (color >> 16) & 0xFF
        g := (color >> 8) & 0xFF
        r := color & 0xFF
        return (r << 16) | (g << 8) | (b << 0)
    } catch
        throw
}

Clamp(value, min, max) {
    if (value < min)
        return min
    if (value > max)
        return max
    return value
}

color_name_to_hex(color_name) {
    name := StrLower(Trim(color_name))
    colors := Map(
        "black","0x000000", "white","0xFFFFFF", "red","0xFF0000",
        "green","0x008000", "blue","0x0000FF", "yellow","0xFFFF00",
        "gray","0x808080", "grey","0x808080", "silver","0xC0C0C0",
        "maroon","0x800000", "olive","0x808000", "lime","0x00FF00",
        "aqua","0x00FFFF", "cyan","0x00FFFF", "teal","0x008080",
        "navy","0x000080", "fuchsia","0xFF00FF", "magenta","0xFF00FF",
        "purple","0x800080", "orange","0xFFA500", "pink","0xFFC0CB",
        "brown","0xA52A2A", "gold","0xFFD700", "indigo","0x4B0082",
        "violet","0xEE82EE"
    )
    return colors.Has(name) ? colors[name] : ""
}