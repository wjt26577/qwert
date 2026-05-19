; ============================================================
; DeepSeek 翻译脚本 (AHK v2) — 纯字符串拼接版
; 快捷键: Ctrl+Shift+T  →  翻译剪贴板内容并弹窗显示
; ============================================================
#SingleInstance Force
#Requires AutoHotkey v2.0

; --- 配置区 ---
apiKey := "sk-05994ea943324c57a0354bfcb3133c9e"   ; 替换为你的 DeepSeek API Key
apiUrl := "https://api.deepseek.com/chat/completions"
model  := "deepseek-v4-pro"
; ---------------

; --- 辅助函数 ---

EscapeJsonStr(str) {
    str := StrReplace(str, "\", "\\")
    str := StrReplace(str, '"', '\"')
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`t", "\t")
    return str
}

GetUtf8Response(whr) {
    responseBytes := whr.ResponseBody
    stream := ComObject("ADODB.Stream")
    stream.Type := 1
    stream.Open()
    stream.Write(responseBytes)
    stream.Position := 0
    stream.Type := 2
    stream.Charset := "UTF-8"
    result := stream.ReadText()
    stream.Close()
    return result
}

DecodeUnicodeEscapes(str) {
    result := ""
    pos := 1
    len := StrLen(str)
    while pos <= len {
        if SubStr(str, pos, 2) = "\u" && pos + 5 <= len {
            hex := SubStr(str, pos + 2, 4)
            if RegExMatch(hex, "^[0-9A-Fa-f]{4}$") {
                charCode := Integer("0x" hex)
                if charCode >= 0xD800 && charCode <= 0xDBFF && pos + 11 <= len {
                    if SubStr(str, pos + 6, 2) = "\u" {
                        lowHex := SubStr(str, pos + 8, 4)
                        if RegExMatch(lowHex, "^[0-9A-Fa-f]{4}$") {
                            lowCode := Integer("0x" lowHex)
                            if lowCode >= 0xDC00 && lowCode <= 0xDFFF {
                                high := charCode - 0xD800
                                low := lowCode - 0xDC00
                                fullCode := 0x10000 + (high << 10) + low
                                result .= Chr(fullCode)
                                pos += 12
                                continue
                            }
                        }
                    }
                }
                result .= Chr(charCode)
                pos += 6
                continue
            }
        }
        result .= SubStr(str, pos, 1)
        pos++
    }
    return result
}

; ============================================================
; 快捷键
; ============================================================
XButton2 & a:: 
{
    clipText := A_Clipboard
    if (clipText = "") {
        MsgBox("剪贴板为空！", "提示", "Icon!")
        return
    }

    ToolTip("正在翻译，请稍候...")
    SetTimer () => ToolTip(), -2000

    ; ---- 纯字符串拼接构建 JSON，避免 continuation section 转义问题 ----
    escapedContent := EscapeJsonStr(clipText)
    body := '{"model": "' . model . '", "messages": ['
    body .= '{"role": "system", "content": "你是一个翻译助手。请将用户输入的内容翻译成中文。只输出翻译结果，不要添加任何解释。"}, '
    body .= '{"role": "user", "content": "' . escapedContent . '"}'
    body .= '], "temperature": 0.3}'

    ; 发送请求
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("POST", apiUrl, true)
    whr.SetRequestHeader("Content-Type", "application/json")
    whr.SetRequestHeader("Authorization", "Bearer " apiKey)
    
    ; 错误处理
    try {
        whr.Send(body)
        whr.WaitForResponse()
    } catch as e {
        MsgBox("网络请求失败！`n`n错误信息: " . e.Message, "网络错误", "IconX")
        return
    }

    status := whr.Status
    if (status != 200) {
        ; 输出调试信息：状态码 + 请求体前200字符 + 响应内容
        debugMsg := "状态码: " . status
        debugMsg .= "`n`n----- 请求体(前300字符) -----`n" . SubStr(body, 1, 300)
        debugMsg .= "`n`n----- 响应 -----`n" . whr.ResponseText
        MsgBox(debugMsg, "API 请求失败 [400=JSON格式错误, 401=API Key无效]", "IconX")
        return
    }

    ; 解析响应
    responseText := GetUtf8Response(whr)

    if RegExMatch(responseText, '"content"\s*:\s*"((?:[^"\\]|\\.)*)"', &match) {
        translated := match[1]
        translated := StrReplace(translated, '\"', '"')
        translated := StrReplace(translated, "\\", "\")
        translated := StrReplace(translated, "\/", "/")
        translated := StrReplace(translated, "\n", "`n")
        translated := StrReplace(translated, "\r", "`r")
        translated := StrReplace(translated, "\t", "`t")
        translated := DecodeUnicodeEscapes(translated)
        MsgBox(translated, "翻译结果", "")
    } else {
        MsgBox("解析失败，原始响应:`n`n" . responseText, "调试信息", "")
    }
}