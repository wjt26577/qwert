MIN_REGION_SIZE := 10
LONG_PRESS_THRESHOLD := 300  ; 长短按时间阈值（毫秒）
SCREENSHOT_PATH := "D:\temp" ; 截图保存路径
SCREENSHOT_PREFIX := "shot"  ; 截图文件前缀
SCREENSHOT_FORMAT := "png"   ; 截图文件格式
BEEP_FREQUENCY := 1500       ; 截图完成提示音频率
INFO_DISPLAY_TIME := 1       ; 信息显示时长（秒）
MIN_REGION_SIZE := 20        ; 最小有效区域大小（像素）




; #region ========================= BAIDU_API核心函数 =========================

; 百度文本翻译主函数
baidu_text_translate(text := "", source_lang := "auto", target_lang := "zh") {
    ; 参数验证
    if (text == "" || Trim(text) == "") {
        return { success: false, error: "翻译文本不能为空" }
    }

    translated_text := ""
    
    try {
        ; 获取配置信息
        config := get_baidu_api_config("text_translation")
        if (!config.success) {
            return { success: false, error: "获取配置失败" }
        }

        ; 获取访问令牌
        token := get_baidu_access_token(config.api_key, config.secret_key, config.token_url)
        if (token == "") {
            return { success: false, error: "获取访问令牌失败" }
        }
        
        ; 构建请求
        url := config.api_url . "?access_token=" . token
        content_type := "Content-Type"
        headers := { %content_type%: "application/json", Accept: "application/json" }    
	    payload := json.dumps({ from: source_lang, to: source_lang, q: text })   
        
        ; 发送请求
        response_text := http_client.get_text_utf8("POST", url, headers, payload)
        if (response_text == "") {
            return { success: false, error: "请求失败，未收到响应" }
        }
        
        ; 解析响应
        response := json.load(response_text)
        if (!response.Has("result") || !response["result"].Has("trans_result")) {
            error_msg := response.Has("error_msg") ? response["error_msg"] : "翻译服务响应格式错误"
            return { success: false, error: error_msg }
        }
        
        ; 提取翻译结果
        trans_results := response["result"]["trans_result"]
        translated_text := ""
        
        for index, result in trans_results {
            if (result.Has("dst")) {
                translated_text .= result["dst"] . "`r`n"
            }
        }
        
        ; 清理结果
        translated_text := Trim(translated_text, "`r`n")
        
        ; 自动复制到剪贴板
        A_Clipboard := translated_text
        
        return {
            success: true,
            original: text,
            translated: translated_text,
            from: source_lang,
            to: target_lang
        }
        
    } catch as err {
        return { 
            success: false, 
            error: "翻译过程发生错误: " . err.message 
        }
    }
}

; 百度图片翻译主函数
baidu_picture_translate(image_path := "", source_lang := "auto", target_lang := "zh") { 
    ; 检查文件是否存在
    if !FileExist(image_path) {     
        return { success: false, error: "文件不存在: " . image_path }
    }
    

    try {
        translation_result := ""

        ; 获取配置信息
        config := get_baidu_api_config("picture_translation")
        if (!config.success) {
            return { success: false, error: "获取配置失败" }
        }

        ; 获取访问令牌
        token := get_baidu_access_token(config.api_key, config.secret_key, config.token_url)
        if (token == "") {
            return { success: false, error: "获取访问令牌失败" }
        }
        
        
        ; 准备表单数据参数
        form_params := Object()
        form_params.from := source_lang
        form_params.to := target_lang
        form_params.v := 3
        form_params.paste := 1
        form_params.image := [image_path]

        ; form_params := { from: source_lang, to: target_lang, v: 3, paste: 1, image: [image_path] }
        
        ; 创建表单数据
        create_form_data(&form_payload, &form_content_type, form_params)
        
        ; 构建请求
        url := config.api_url . "?access_token=" . token
        content_type := "Content-Type"
        headers := { %content_type%: form_content_type, Accept: "application/json" }    

         ; 发送翻译请求
        response_text := http_client.get_text("POST", url, headers, form_payload)          
        if (response_text == "") {
            return { success: false, error: "请求失败，未收到响应" }
        }

    ; MsgBox(url, , "T7")
        ; 解析响应
        response := json.load(response_text)
        if (!response.Has("data") || !response["data"].Has("sumDst")) {
            error_msg := response.Has("error_code") ? response["error_code"] : "翻译服务响应格式错误"
            return { success: false, error: error_msg }
        }
        
        ; 提取翻译结果
        trans_results := response["data"]["sumDst"]

        ; 清理结果
        translated_text := Trim(trans_results, "`r`n")
        
        ; 自动复制到剪贴板
        A_Clipboard := translated_text
        
        return {
            success: true,
            original: "",
            translated: translated_text,
            from: source_lang,
            to: target_lang
        }
        
    } catch as err {
        return { 
            success: false, 
            error: "翻译过程发生错误: " . err.message 
        }
    }
}

; 百度OCR识别（位图版本）主函数
baidu_ocr(bitmap, ocr_type := "accurate_ocr") {
    ; 验证位图对象
    if !bitmap {     
        return { success: false, error: "无效的位图对象" }
    }
  


    try {
        ; 获取配置信息
        config := get_baidu_api_config(ocr_type)

        if (!config.success) {
            return { success: false, error: "获取配置失败" }
        }

        ; 获取访问令牌
        token := get_baidu_access_token(config.api_key, config.secret_key, config.token_url)
        if (token == "") {
            return { success: false, error: "获取访问令牌失败" }
        }
        
        
        ; 构建请求
        url := config.api_url . "?access_token=" . token
        content_type := "Content-Type"
        headers := { %content_type%: "application/x-www-form-urlencoded", Accept: "application/json" }    


       

        ; 将位图转换为Base64编码
        base64_image := Gdip_EncodeBitmapTo64string(bitmap)     
        encoded_image := url_encode(base64_image)
        payload := "image=" . encoded_image . "&detect_direction=false&paragraph=false&probability=false"
     

        
        ; 发送请求
        response_text := http_client.get_text_utf8("POST", url, headers, payload)
        if (response_text == "") {
            return { success: false, error: "请求失败，未收到响应" }
        }
        
        
        ; 解析响应
        json_result := json.load(response_text)
        
        if (json_result.Has("error_code")) {
            error_msg := json_result.Has("error_msg") ? json_result["error_msg"] : "OCR错误"
            return { success: false, error: error_msg }
        }
        ; 
        ; MsgBox(response_text, , "T10")
        
        
        ; 提取翻译结果
        words_results := json_result["words_result"]        
        recognized_text  := ""


        for index, result in words_results {
            if (result.Has("words")) {
                recognized_text  .= result["words"] . "`r`n"
            }
        }


        recognized_text  := Trim(recognized_text , "`r`n")          
        
        ; 自动复制到剪贴板
        A_Clipboard := recognized_text
                
        return {
            success: true,
            recognized: recognized_text
        }
        
    } catch as err {
        return { 
            success: false, 
            error: "翻译过程发生错误: " . err.message 
        }
    }
}

; 获取百度翻译配置
get_baidu_api_config(mode := "ocr") {
    global baidu_ai
    try {        
        config := baidu_ai[mode]
        return {
            success: true,
            api_key: config["api_key"],
            secret_key: config["secret_key"],
            api_url: config["api_url"],
            token_url: config["token_url"]
        } 
    } catch as err {
        return { 
            success: false, 
            error: "读取配置失败: " . err.message 
        }
    }
}

; 获取百度访问令牌
get_baidu_access_token(api_key, secret_key, token_url) {
    try {
        ; 构建请求
        url := token_url . "?grant_type=client_credentials&client_id=" . api_key . "&client_secret=" . secret_key 
        
        headers := Map()
        headers["Content-Type"] :="application/json"
        headers["Accept"] := "application/json"

        payload := ""

        ; 发送请求
        response_text := http_client.get_text_utf8("POST", url, headers, payload)
        if (response_text == "") {
             return ""
        }
              
        ; 解析响应并检查错误
        json_result := json.load(response_text)
        if (json_result.Has("error")) {
            Throw Error("API错误: " . json_result["error_description"])
        }
        
        return json_result["access_token"]
    } catch as err {
        ; 显示错误信息
        Msgbox("获取访问令牌失败: " . err.message)
        return ""
    }
}


; #endregion


; #region ========================= 区域选择和捕获功能 =========================



/**
 * 简化版区域选择（输入对话框方式）
 * @return {Object} 用户选择的区域 {x, y, width, height}，取消返回""
 */
select_region_dialog() {
    try {
        ; 获取屏幕尺寸
        screen_width := A_ScreenWidth
        screen_height := A_ScreenHeight

        ; 输入对话框获取区域信息
        x_result := InputBox("请输入X坐标 (0-" . screen_width . "):", "区域选择", "w200 h130", "0")
        if (x_result.Result = "Cancel") {
            return 0
        } else {
            x_input := x_result.Value
        }
        
        y_result := InputBox("请输入Y坐标 (0-" . screen_height . "):", "区域选择", "w200 h130", "0")
        if (y_result.Result = "Cancel") {
            return 0
        } else {
            y_input := y_result.Value
        }
        
        width_result := InputBox("请输入宽度:", "区域选择", "w200 h130", "400")
        if (width_result.Result = "Cancel") {
            return 0
        } else {
            width_input := width_result.Value
        }
        
        height_result := InputBox("请输入高度:", "区域选择", "w200 h130", "300")
        if (height_result.Result = "Cancel") {
            return 0
        } else {
            height_input := height_result.Value
        }
        
        ; 验证输入
        region_x := Integer(x_input)
        region_y := Integer(y_input)
        region_width := Integer(width_input)
        region_height := Integer(height_input)
        
        if (region_x < 0 || region_y < 0 || region_width <= 0 || region_height <= 0 ||
            region_x + region_width > screen_width || region_y + region_height > screen_height) {
            Throw Error("区域参数超出屏幕范围")
        }
        
        return {x: region_x, y: region_y, width: region_width, height: region_height}
        
    } catch as err {
        Msgbox("区域选择失败: " . err.message)
        return 0
    }
}

/**
 * 保存区域截图
 * @param {Object} region - 区域对象 {x, y, width, height}
 * @param {String} save_path - 保存路径
 * @param {String} image_format - 图片格式（默认"PNG"）
 * @return {Boolean} 保存是否成功
 */
save_region_screenshot(region, save_path) {
    token := 0    ; GDI+ token
    bitmap := 0   ; 位图对象

    try {        
        ; 启动GDI+
        if (!(token := Gdip_Startup())) {
            Msgbox("GDI+ 启动失败，请确保您的系统中安装了 GDI+")
            return false
        }

        ; 捕获区域
        bitmap := Gdip_BitmapFromScreen(region)
        if (!bitmap) {
            Throw Error("无法捕获指定区域")
        }

        ; 保存图片
        Gdip_SaveBitmapToFile(bitmap, save_path)

        ; 复制到剪贴板
        Gdip_SetBitmapToClipboard(bitmap)
        
        return true
        
    } catch as err {
        Msgbox("截图错误: " . err.Message)
        return false
        
    } finally {
        ; 确保资源释放
        if (bitmap) {
            Gdip_DisposeImage(bitmap)
        }
        if (token) {
            Gdip_Shutdown(token)
        }
    }
}

; #endregion


; #region ========================= OCR =========================
/**
 * 交互式区域OCR识别函数（完整版）
 * 执行完整的OCR流程：区域选择 -> OCR识别 -> 复制剪贴板 -> 显示结果
 * @param {String} ocr_type - OCR识别类型（可选，默认"accurate"）
 * @param {String} output_path - 结果保存路径（可选，默认使用全局配置）
 * @param {Boolean} show_in_notepad - 是否在记事本中显示（可选，默认true）
 * @param {Boolean} copy_to_clipboard - 是否复制到剪贴板（可选，默认true）
 * @return {Object} 返回结果对象 {success: Boolean, text: String, message: String, region: String}
 */
interactive_region_ocr(ocr_type := "accurate_ocr", output_path := "", show_in_notepad := true, copy_to_clipboard := true) {
    token := 0    ; GDI+ token
    bitmap := 0   ; 位图对象
    
    try {
        ; 1. 交互式选择区域
        result := select_region_interactive("blue")
        if (!result.region || result.region == "") {
            return {success: false, text: "", message: "用户取消了区域选择", region: ""}
        }
        
        ; 2. 验证区域格式（应该是"x|y|w|h"格式）
        region_parts := StrSplit(result.region, "|")
        if (region_parts.Length != 4) {
            return {success: false, text: "", message: "区域格式无效", region: result.region}
        }
        
        ; 验证区域大小
        width := Integer(region_parts[3])
        height := Integer(region_parts[4])
        if (width < 10 || height < 10) {
            return {success: false, text: "", message: "选择的区域太小", region: result.region}
        }
        
        ; 3. 启动GDI+
        if (!(token := Gdip_Startup())) {
            return {success: false, text: "", message: "GDI+ 启动失败，请确保您的系统中安装了 GDI+", region: result.region}
        }
        
        ; 4. 从屏幕指定区域创建位图（直接使用字符串格式）
        bitmap := Gdip_BitmapFromScreen(result.region)
        if (!bitmap || bitmap == -1) {
            return {success: false, text: "", message: "截图失败：无法从屏幕获取图像（参数错误或区域无效）", region: result.region}
        }
    
        ; 5. 进行OCR识别
        ocr_text := baidu_ocr(bitmap, ocr_type).recognized
        if (!ocr_text || ocr_text == "") {
            return {success: false, text: "", message: "OCR识别失败：未识别到文本内容", region: result.region}
        }
       
        
        ; 6. 复制到剪贴板
        if (copy_to_clipboard) {
            try {
                A_Clipboard := ocr_text
            } catch as err {
                ; 复制失败但继续执行其他操作
                Notify.show("文本已识别，但复制到剪贴板失败: " . err.message)
            }
        }
        
        ; 7. 显示在记事本中
        if (show_in_notepad) {
            try {
                display_result := display_ocr_in_notepad(ocr_text, output_path)
                if (!display_result.success) {
                    Notify.show("文本已识别并复制，但记事本显示失败: " . display_result.message)
                }
            } catch as err {
                Notify.show("文本已识别并复制，但记事本显示异常: " . err.message)
            }
        }
        
        ; 8. 显示成功提示
        text_preview := StrLen(ocr_text) > 20 ? SubStr(ocr_text, 1, 20) . "..." : ocr_text
        Notify.show("OCR识别完成: " . text_preview)
        
        return {
            success: true,
            text: ocr_text,
            message: "OCR识别完成",
            region: result.region,
            text_length: StrLen(ocr_text)
        }
        
    } catch as err {
        error_msg := "OCR过程中发生错误: " . err.message
        Msgbox(error_msg)
        return {success: false, text: "", message: error_msg, region: result.HasProp("region") ? result.region : ""}
        
    } finally {
        ; 确保资源释放
        if (bitmap) {
            Gdip_DisposeImage(bitmap)
        }
        if (token) {
            Gdip_Shutdown(token)
        }
    }
}

/**
 * 快速OCR函数（简化版）
 * 适用于热键调用，使用默认配置
 * @return {String} 识别的文本内容，失败返回空字符串
 */
quick_ocr() {
    result := interactive_region_ocr()
    return result.success ? result.text : ""
}

/**
 * 静默OCR函数（无UI提示版本）
 * 只进行OCR识别，不显示任何UI
 * @param {String} ocr_type - OCR识别类型
 * @return {Object} OCR结果
 */
silent_ocr(ocr_type := "accurate") {
    return interactive_region_ocr(ocr_type, "", false, true)
}

/**
 * 自定义OCR函数
 * 允许完全自定义所有参数
 * @param {Object} options - 选项对象
 * @return {Object} OCR结果
 */
custom_ocr(options := {}) {
    ; 设置默认选项
    default_options := {
        ocr_type: "accurate",
        output_path: "",
        show_in_notepad: true,
        copy_to_clipboard: true,
        auto_translate: false,
        target_language: "zh"
    }
    
    ; 合并用户选项
    final_options := merge_options(default_options, options)
    
    ; 执行OCR
    result := interactive_region_ocr(
        final_options.ocr_type,
        final_options.output_path,
        final_options.show_in_notepad,
        final_options.copy_to_clipboard
    )
    
    
    return result
}

/**
 * 改进的记事本显示函数
 * @param {String} text - 要显示的文本
 * @param {String} custom_path - 自定义保存路径
 * @return {Object} 显示结果
 */
display_ocr_in_notepad(text, custom_path := "") {
    try {
        ; 生成文件信息
        time_stamp := FormatTime(, "yyyy-MM-dd_HH:mm:ss")
        day_stamp := FormatTime(, "yyyyMMdd")
        file_name := "ocr_" . day_stamp . ".txt"
        
        ; 确定文件路径
        if (custom_path && custom_path != "") {
            ; 如果提供了自定义路径
            if (InStr(custom_path, ".txt")) {
                file_path := custom_path
                SplitPath(custom_path, &file_name)
            } else {
                file_path := RTrim(custom_path, "\") . "\" . file_name
            }
        } else {
            ; 使用默认路径
            base_path := "D:\temp"
            if (!DirExist(base_path)) {
                DirCreate(base_path)
            }
            file_path := base_path . "\" . file_name
        }
        
        ; 格式化文本内容
        separator_line := pad_right("", 94, "-")
        formatted_text := Format("`r`n`r`n{1} {2}`r`n`r`n{3}`r`n`r`n{1} {2}`r`n`r`n",
                                time_stamp,
                                separator_line,
                                text)
        
        ; 检查记事本是否已打开相同文件
        notepad_title := file_name . " - Notepad"
        if (WinExist(notepad_title)) {
            ; 激活现有窗口
            WinActivate(notepad_title)
            if (WinWaitActive(notepad_title, , 2)) {
                Sleep(300)
                Send("^{Home}")  ; 移动到文件开头
                send_by_clipboard_ocr(formatted_text)
            } else {
                Throw Error("无法激活记事本窗口")
            }
        } else {
            ; 创建新文件并打开
            if (!FileExist(file_path)) {
                FileAppend("", file_path, "UTF-8")
            }
            
            Run('"' . file_path . '"')
            if (WinWaitActive("ahk_class Notepad", , 3)) {
                Sleep(300)
                Send("^{Home}")  ; 移动到文件开头
                send_by_clipboard_ocr(formatted_text)
            } else {
                Throw Error("无法打开记事本")
            }
        }
        
        return {success: true, file_path: file_path, message: "文本已显示在记事本中"}
        
    } catch as err {
        return {success: false, file_path: "", message: "记事本显示失败: " . err.message}
    }
}

/**
 * 显示翻译结果在记事本中
 */
display_translation_in_notepad(original_text, translated_text, custom_path := "") {
    combined_text := "原文：`r`n" . original_text . "`r`n`r`n译文：`r`n" . translated_text
    return display_ocr_in_notepad(combined_text, custom_path)
}


send_by_clipboard_ocr(text) {
    old_clipboard := A_Clipboard
    try {
        A_Clipboard := text
        ClipWait(1)
        Send("^v")
        Sleep(100)
    } finally {
        ; 恢复原剪贴板内容
        SetTimer(() => (A_Clipboard := old_clipboard), -1000)
    }
}

/**
 * 合并选项对象
 */
merge_options(defaults, user_options) {
    result := {}
    
    ; 复制默认选项
    for key, value in defaults {
        result.%key% := value
    }
    
    ; 覆盖用户选项
    for key, value in user_options {
        result.%key% := value
    }
    
    return result
}

; 使用示例：

; 示例1: 简单的热键绑定
; ^RButton::func_ctrl_rbutton()

; 示例2: 快速OCR（只返回文本）
; F7:: {
;     text := quick_ocr()
;     if (text) {
;         MsgBox("识别文本：`n" . text)
;     }
; }

; 示例3: 静默OCR（无UI提示）
; F8:: {
;     result := silent_ocr("fast")
;     if (result.success) {
;         ToolTip("已复制: " . SubStr(result.text, 1, 30) . "...")
;         SetTimer(() => ToolTip(), -2000)
;     }
; }

; 示例4: 自定义OCR（带翻译）
; F9:: {
;     options := {
;         ocr_type: "accurate",
;         show_in_notepad: true,
;         auto_translate: true,
;         target_language: "en"
;     }
;     result := custom_ocr(options)
;     if (result.success && result.HasProp("translated_text")) {
;         MsgBox("原文: " . result.text . "`n译文: " . result.translated_text)
;     }
; }

; 示例5: 带错误处理的调用
; F10:: {
;     result := interactive_region_ocr("fast", "D:\MyOCR\results.txt")
;     if (result.success) {
;         Notify.show("OCR成功，识别了 " . result.text_length . " 个字符")
;     } else {
;         Msgbox("OCR失败: " . result.message)
;     }
; }

; #endregion


; #region ========================= screenshot =========================


; 交互式区域截图函数：区域选择 -> 保存文件 -> 复制剪贴板 -> Snipaste显示
interactive_region_screenshot(if_stick := True, if_show := True, back_color := "Purple") {    
    save_path := run_config["path_screenshot_path"]["payload"]

    token := 0    ; GDI+ token
    bitmap := 0   ; 位图对象
    try {
        ; 1. 交互式选择区域
        result := select_region_interactive(back_color)        
        if (!result.region) {
            return {success: false, file_path: "", message: "用户取消了区域选择"}
        }
        
        ; 2. 确保保存目录存在
        if (!DirExist(save_path)) {
            try {
                DirCreate(save_path)
            } catch as err {
                return {success: false, file_path: "", message: "无法创建保存目录: " . err.message}
            }
        }
     

        ; 3. 生成带时间戳的文件路径
        time_stamp := FormatTime(, "yyyyMMdd_HHmmss")      
        file_path := Format("{1}\{2}.png", save_path, time_stamp)
      
        ; 4. 启动GDI+
        if (!(token := Gdip_Startup())) {
            return {success: false, file_path: "", message: "GDI+ 启动失败，请确保您的系统中安装了 GDI+"}
        }          
   

        ; 5. 捕获指定区域
        bitmap := Gdip_BitmapFromScreen(result.region)
        if (bitmap == -1) {
            return {success: false, file_path: "", message: "无法捕获指定区域"}
        }
        
        ; 6. 保存图片到文件
        save_result := Gdip_SaveBitmapToFile(bitmap, file_path)
        if (save_result != 0) {  ; 0 = 成功，非零 = 失败
            error_messages := Map(
                -1, "不支持的文件格式",
                -2, "无法获取编码器列表", 
                -3, "找不到匹配的编码器",
                -4, "无法获取输出文件的宽字符名称",
                -5, "无法保存文件到磁盘"
            )
            error_msg := error_messages.Has(save_result) ? error_messages[save_result] : "未知错误(" . save_result . ")"
            return {success: false, file_path: "", message: "保存图片失败: " . error_msg}
        }
        
        ; 7. 复制到剪贴板（此函数没有返回值） 
        try {
            Gdip_SetBitmapToClipboard(bitmap)
        } catch as err {
            ; 保存成功但复制失败，仍然算部分成功
            return {success: true, file_path: file_path, message: "截图已保存，但复制到剪贴板失败: " . err.message}
        }                 
        
        ; 8. 使用Snipaste贴在屏幕上
        if(if_stick) {
            try {
                Sleep(50)
                SendInput("{F3}")
                Sleep(200) 
                SendInput("{Space}")
                Sleep(10) 
                SendInput("b")

            } catch {
                ; Snipaste可能未安装，忽略错误
            }
        }
        
        ; 9. 显示成功提示
        if if_show {
            Notify.show("截图已保存并复制到剪贴板")
        }       
                
        return {
            success: true, 
            file_path: file_path, 
            message: "截图操作完成",
            region: result.region,
            file_size: FileGetSize(file_path)
        }
        
    } catch as err {
        error_msg := "截图过程中发生错误: " . err.message
        Msgbox(error_msg)
        return {success: false, file_path: "", message: error_msg}
        
    } finally {
        ; 确保资源释放
        if (bitmap) {
            Gdip_DisposeImage(bitmap)
        }
        if (token) {
            Gdip_Shutdown(token)
        }
    }
}

/**
 * 快速截图函数（简化版）
 * 适用于热键调用，使用默认配置
 * @return {Boolean} 截图是否成功
 */
quick_screenshot() {
    result := interactive_region_screenshot()
    return result.success
}

/**
 * 带参数的截图函数
 * 允许自定义保存路径和文件名
 * @param {String} custom_path - 自定义保存路径
 * @param {String} custom_name - 自定义文件名（不含扩展名）
 * @return {Object} 截图结果
 */
custom_screenshot(custom_path := "", custom_name := "") {
    if (custom_name) {
        ; 如果指定了自定义文件名，直接使用
        if (custom_path) {
            full_path := custom_path . "\" . custom_name . "." . SCREENSHOT_FORMAT
        } else {
            full_path := SCREENSHOT_PATH . "\" . custom_name . "." . SCREENSHOT_FORMAT
        }
        
        ; 手动处理文件保存
        result := select_region_interactive()
        if (!result.region) {
            return {success: false, file_path: "", message: "用户取消了区域选择"}
        }
        
        return save_region_screenshot_to_path(result.region, full_path)
    } else {
        ; 使用标准流程
        return interactive_region_screenshot(custom_path)
    }
}

/**
 * 保存区域截图到指定路径
 * @param {Object} region - 区域对象 {x, y, width, height}
 * @param {String} full_path - 完整的保存路径
 * @return {Object} 保存结果
 */
save_region_screenshot_to_path(region, full_path) {
    token := 0
    bitmap := 0
    
    try {
        ; 确保目录存在
        dir_path := StrReplace(full_path, "\" . StrSplit(full_path, "\")[-1], "")
        if (!DirExist(dir_path)) {
            DirCreate(dir_path)
        }
        
        ; 启动GDI+
        if (!(token := Gdip_Startup())) {
            return {success: false, file_path: "", message: "GDI+ 启动失败"}
        }
        
        ; 捕获区域
        bitmap := Gdip_BitmapFromScreen(region)
        if (!bitmap) {
            return {success: false, file_path: "", message: "无法捕获指定区域"}
        }
        
        ; 保存图片
        if (!Gdip_SaveBitmapToFile(bitmap, full_path)) {
            return {success: false, file_path: "", message: "保存图片失败"}
        }
        
        ; 复制到剪贴板
        Gdip_SetBitmapToClipboard(bitmap)
        
        ; 使用Snipaste显示
        try {
            Send("{F3}")
        } catch {
            ; 忽略Snipaste错误
        }
        
        Notify.show("截图已保存: " . StrSplit(full_path, "\")[-1])
        
        return {
            success: true,
            file_path: full_path,
            message: "截图保存成功",
            region: region,
            file_size: FileGetSize(full_path)
        }
        
    } catch as err {
        return {success: false, file_path: "", message: "保存失败: " . err.message}
        
    } finally {
        if (bitmap) {
            Gdip_DisposeImage(bitmap)
        }
        if (token) {
            Gdip_Shutdown(token)
        }
    }
}

; 使用示例：

; 示例1: 简单的热键绑定
; F4::quick_screenshot()

; 示例2: 带错误处理的调用
; F5:: {
;     result := interactive_region_screenshot()
;     if (result.success) {
;         MsgBox("截图成功！`n文件: " . result.file_path . "`n大小: " . result.file_size . " bytes")
;     } else {
;         MsgBox("截图失败: " . result.message)
;     }
; }

; 示例3: 自定义路径和文件名
; F6:: {
;     custom_path := A_Desktop
;     custom_name := "my_screenshot_" . A_TickCount
;     result := custom_screenshot(custom_path, custom_name)
;     ToolTip(result.success ? "保存成功" : "保存失败")
;     SetTimer(() => ToolTip(), -2000)
; }



; #endregion


; #region ========================= translation =========================






; 智能翻译函数
smart_translate(text := "") {
    ; 参数验证
    if (text == "" || Trim(text) == "") {
        ; 如果没有传入文本，尝试从剪贴板获取
        text := A_Clipboard
        if (text == "" || Trim(text) == "") {
            return { success: false, error: "没有可翻译的文本" }
        }
    }
    
    try {
        ; 检测语言类型
        lang_info := detect_language(text)
        
        ; 根据检测结果选择翻译方向
        if (lang_info.is_chinese) {
            ; 中文翻译为英文
            result := baidu_text_translate(text, "zh", "en")
            if (result.success) {
                result.detected_lang := "中文"
                result.translate_direction := "中文 → 英文"
            }
        } else if (lang_info.is_english) {
            ; 英文翻译为中文
            result := baidu_text_translate(text, "en", "zh")
            if (result.success) {
                result.detected_lang := "英文"
                result.translate_direction := "英文 → 中文"
            }
        } else {
            ; 其他语言翻译为中文
            result := baidu_text_translate(text, "auto", "zh")
            if (result.success) {
                result.detected_lang := lang_info.main_lang
                result.translate_direction := result.detected_lang . " → 中文"
            }
        }
        
        return result
        
    } catch as err {
        return { 
            success: false, 
            error: "智能翻译过程发生错误: " . err.message 
        }
    }
}

; 语言检测函数（只检查前5个字符）逻辑：前5个字符中有中文=中文，没有中文=英文
detect_language(text) {
    ; 去除空白字符进行分析
    clean_text := Trim(text)
    total_chars := StrLen(clean_text)
    
    if (total_chars == 0) {
        return {
            is_chinese: false,
            is_english: false,
            main_lang: "未知"
        }
    }
    
    ; 只检查前5个字符，快速判断
    check_length := Min(5, total_chars)
    has_chinese := false
    
    Loop check_length {
        char := SubStr(clean_text, A_Index, 1)
        char_code := Ord(char)
        
        ; 检测中文字符
        if (is_chinese_char(char_code)) {
            has_chinese := true
            break  ; 发现中文立即跳出循环
        }
    }
    
    ; 简单判断逻辑
    is_chinese := has_chinese
    is_english := !has_chinese  ; 没有中文就当作英文处理
    main_lang := has_chinese ? "中文" : "英文"
    
    return {
        is_chinese: is_chinese,
        is_english: is_english,
        main_lang: main_lang
    }
}

; 检测是否为中文字符
is_chinese_char(char_code) {
    ; 中文字符Unicode范围
    return (char_code >= 0x4E00 && char_code <= 0x9FFF) ||    ; CJK统一汉字
           (char_code >= 0x3400 && char_code <= 0x4DBF) ||    ; CJK扩展A
           (char_code >= 0x20000 && char_code <= 0x2A6DF) ||  ; CJK扩展B
           (char_code >= 0x2A700 && char_code <= 0x2B73F) ||  ; CJK扩展C
           (char_code >= 0x2B740 && char_code <= 0x2B81F) ||  ; CJK扩展D
           (char_code >= 0x3000 && char_code <= 0x303F) ||    ; CJK符号和标点
           (char_code >= 0xFF00 && char_code <= 0xFFEF)       ; 全角ASCII
}


; 智能翻译剪贴板内容
smart_translate_clipboard() {
    result := smart_translate()
    
    if (result.success) {
        msg := "智能翻译完成！`n`n"
        msg .= "检测语言: " . result.detected_lang . "`n"
        msg .= "翻译方向: " . result.translate_direction . "`n`n"
        msg .= "原文: " . result.original . "`n`n"
        msg .= "译文: " . result.translated
        
        MsgBox(msg, "智能翻译成功", 64)
    } else {
        MsgBox("智能翻译失败: " . result.error, "翻译错误", 16)
    }
    
    return result
}

; 快速智能翻译（静默模式）
quick_translate(text := "") {
    result := smart_translate(text)
    
    if (result.success) {
        ; 静默复制到剪贴板，显示简短提示
        A_Clipboard := result.translated
        ToolTip("已翻译: " . result.translate_direction, 100, 100)
        SetTimer(() => ToolTip(), -2000)  ; 2秒后隐藏提示
    } else {
        ToolTip("翻译失败: " . result.error, 100, 100)
        SetTimer(() => ToolTip(), -3000)  ; 3秒后隐藏提示
    }
    
    return result
}

; 语言检测演示（简化版）
show_language_detection(text := "") {
    if (text == "") {
        text := A_Clipboard
    }
    
    if (text == "") {
        MsgBox("没有文本可供检测", "语言检测", 48)
        return
    }
    
    lang_info := detect_language(text)
    
    msg := "快速语言检测结果:`n`n"
    msg .= "文本: " . SubStr(text, 1, 50) . (StrLen(text) > 50 ? "..." : "") . "`n`n"
    msg .= "检测方法: 检查前5个字符`n"
    msg .= "检测结果: " . lang_info.main_lang . "`n`n"
    msg .= "翻译方向: " . (lang_info.is_chinese ? "中文 → 英文" : "英文 → 中文")
    
    MsgBox(msg, "语言检测结果", 64)
    
    return lang_info
}

; =============================================================================
; 快捷键绑定示例
; =============================================================================

; Ctrl+Shift+T: 智能翻译剪贴板
; ^+t::smart_translate_clipboard()

; Ctrl+Alt+T: 快速智能翻译（静默模式）
; ^!t::quick_translate()

; Ctrl+Shift+D: 显示语言检测结果
; ^+d::show_language_detection()

; =============================================================================
; 使用示例
; =============================================================================

; 示例1: 基本使用
; result := smart_translate("Hello World")
; if (result.success) {
;     MsgBox("检测到: " . result.detected_lang . "`n翻译结果: " . result.translated)
; }

; 示例2: 快速语言检测
; lang_info := detect_language("这是一段中文测试文本")
; MsgBox("检测结果: " . lang_info.main_lang)

; 示例3: 快速翻译
; quick_translate("这段文字会被自动检测并翻译")



; =============================================================================
; 便捷使用函数
; =============================================================================

; 翻译为中文（默认）
translate_to_chinese(text) {
    return baidu_text_translate(text, "auto", "zh")
}

; 翻译为英文
translate_to_english(text) {
    return baidu_text_translate(text, "auto", "en")
}

; 翻译剪贴板内容
translate_clipboard(target_lang := "zh") {
    clipboard_text := A_Clipboard
    if (clipboard_text == "") {
        MsgBox("剪贴板为空", "翻译失败", 48)
        return
    }
    
    result := baidu_text_translate(clipboard_text, "auto", target_lang)
    
    if (result.success) {
        MsgBox("翻译完成！`n`n原文: " . result.original . "`n`n译文: " . result.translated, "翻译成功", 64)
    } else {
        MsgBox("翻译失败: " . result.error, "翻译错误", 16)
    }
    
    return result
}



; #endregion




; #region ========================= screenshot + translation =========================

; 截图翻译综合函数
screenshot_translate(source_lang := "auto", target_lang := "zh", auto_copy := true, 
    show_result := true, cleanup_file := true, back_color := "green") {
    
    temp_file := ""  ; 临时文件路径
    
    try {
        ; 1. 执行交互式截图
        screenshot_result := interactive_region_screenshot(0, 0, back_color)
        if (!screenshot_result.success) {
            return {
                success: false, 
                error: "截图失败: " . screenshot_result.message,
                step: "screenshot"
            }
        }
        
        temp_file := screenshot_result.file_path
        
        ; 2. 检查截图文件是否存在
        if (!FileExist(temp_file)) {
            return {
                success: false, 
                error: "截图文件不存在: " . temp_file,
                step: "file_check"
            }
        }
        
        ; 3. 执行图片翻译
        translate_result := baidu_picture_translate(temp_file, source_lang, target_lang)
        if (!translate_result.success) {
            return {
                success: false, 
                error: "翻译失败: " . translate_result.error,
                step: "translation",
                file_path: temp_file
            }
        }
        
        ; 4. 获取翻译结果
        translated_text := translate_result.translated
        if (translated_text == "") {
            return {
                success: false, 
                error: "翻译结果为空",
                step: "result_empty",
                file_path: temp_file
            }
        }
        
        ; 5. 自动复制到剪贴板
        if (auto_copy) {
            try {
                A_Clipboard := translated_text
            } catch as err {
                ; 复制失败不影响主要功能，记录警告
                return {
                    success: true,
                    translated: translated_text,
                    original: translate_result.original,
                    from: translate_result.from,
                    to: translate_result.to,
                    file_path: temp_file,
                    warning: "复制到剪贴板失败: " . err.message
                }
            }
        }
        
        ; 6. 显示翻译结果
        if (show_result) {
            try {
                display_time := StrLen(translated_text) / 3  ; 根据文本长度计算显示时间
                display_time := Max(3, Min(display_time, 100))  ; 限制在3-10秒之间
                Notify.show(translated_text)
            } catch {
                ; 显示失败不影响主要功能，忽略错误
            }
        }
        
        ; 7. 清理临时文件
        if (cleanup_file) {
            try {
                FileDelete(temp_file)
                temp_file := ""  ; 清空变量表示已删除
            } catch {
                ; 删除失败不影响主要功能，记录警告
            }
        }
        
        return {
            success: true,
            translated: translated_text,
            original: translate_result.original,
            from: translate_result.from,
            to: translate_result.to,
            file_path: cleanup_file ? "" : temp_file,
            message: "截图翻译完成"
        }
        
    } catch as err {
        error_msg := "截图翻译过程发生错误: " . err.message
        Msgbox(error_msg)
        
        return {
            success: false, 
            error: error_msg,
            step: "unknown",
            file_path: temp_file
        }
        
    } finally {
        ; 确保在发生错误时也能清理临时文件
        if (cleanup_file && temp_file && FileExist(temp_file)) {
            try {
                FileDelete(temp_file)
            } catch {
                ; 静默处理清理失败
            }
        }
    }
}

; #endregion

; #region ========================= OCR + translation =========================
; OCR识别并翻译综合函数
ocr_texttrans(source_lang := "auto", target_lang := "zh", ocr_type := "accurate_ocr", 
    auto_copy := true, show_result := true, show_in_notepad := true, output_path := "") {
    
    try {
        ; 1. 执行OCR识别（禁用自动复制和记事本显示，由本函数统一处理）
        ocr_result := interactive_region_ocr(ocr_type, output_path, false, false)
        if (!ocr_result.success) {
            return {
                success: false,
                ocr_text: "",
                translated: "",
                error: "OCR识别失败: " . ocr_result.message,
                step: "ocr",
                region: ocr_result.HasProp("region") ? ocr_result.region : ""
            }
        }
        
        ; 2. 检查OCR识别结果
        recognized_text := ocr_result.text
        if (!recognized_text || Trim(recognized_text) == "") {
            return {
                success: false,
                ocr_text: "",
                translated: "",
                error: "OCR识别结果为空，可能是图片中没有文本",
                step: "ocr_empty",
                region: ocr_result.region
            }
        }
        
        ; 3. 进行文本翻译
        translate_result := baidu_text_translate(recognized_text, source_lang, target_lang)
        if (!translate_result.success) {
            return {
                success: false,
                ocr_text: recognized_text,
                translated: "",
                error: "文本翻译失败: " . translate_result.error,
                step: "translation",
                region: ocr_result.region,
                text_length: ocr_result.text_length
            }
        }
        
        ; 4. 检查翻译结果
        translated_text := translate_result.translated
        if (!translated_text || Trim(translated_text) == "") {
            return {
                success: false,
                ocr_text: recognized_text,
                translated: "",
                error: "翻译结果为空",
                step: "translation_empty",
                region: ocr_result.region,
                text_length: ocr_result.text_length
            }
        }
        
        ; 5. 自动复制到剪贴板（优先复制翻译结果）
        copy_warning := ""
        if (auto_copy) {
            try {
                A_Clipboard := translated_text
            } catch as err {
                copy_warning := "复制到剪贴板失败: " . err.message
            }
        }
        
        ; 6. 显示在记事本中（显示原文和译文）
        notepad_warning := ""
        if (show_in_notepad) {
            try {
                combined_text := "=== OCR识别结果 ===" . "`r`n" . recognized_text . "`r`n`r`n" . 
                                "=== 翻译结果 ===" . "`r`n" . translated_text
                display_result := display_ocr_in_notepad(combined_text, output_path)
                if (!display_result.success) {
                    notepad_warning := "记事本显示失败: " . display_result.message
                }
            } catch as err {
                notepad_warning := "记事本显示异常: " . err.message
            }
        }
        
        ; 7. 显示翻译结果
        if (show_result) {
            ; try {
            ;     ; 显示翻译结果，根据文本长度计算显示时间
            ;     display_time := StrLen(translated_text) / 3
            ;     display_time := Max(3000, Min(display_time, 10000))  ; 限制在3-10秒之间
                
            ;     result_preview := StrLen(translated_text) > 30 ? SubStr(translated_text, 1, 30) . "..." : translated_text
            ;     Notify.show("翻译完成: " . result_preview, display_time, bc := "0x0b5e2b")
            ; } catch {
            ;     ; 显示失败不影响主要功能
            ; }
        }
        
        ; 8. 组装警告信息
        warnings := []
        if (copy_warning != "") {
            warnings.Push(copy_warning)
        }
        if (notepad_warning != "") {
            warnings.Push(notepad_warning)
        }
        
        return {
            success: true,
            ocr_text: recognized_text,
            translated: translated_text,
            original: recognized_text,
            from: translate_result.from,
            to: translate_result.to,
            region: ocr_result.region,
            text_length: ocr_result.text_length,
            ; translated_length: StrLen(translated_text),
            message: "OCR识别和翻译完成",
            warnings: warnings.Length > 0 ? warnings : ""
        }
        
    } catch as err {
        error_msg := "OCR翻译过程中发生错误: " . err.message
        Msgbox(error_msg)
        
        return {
            success: false,
            ocr_text: "",
            translated: "",
            error: error_msg,
            step: "unknown"
        }
    }
}
; #endregion






; #region ========================= debug =========================


; #endregion