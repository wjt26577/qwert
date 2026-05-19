

; #region ========================= HTTP客户端类 =========================

/**
 * HTTP客户端类
 * 包含错误处理、超时设置和UTF-8支持
 */
class http_client {
    
    /**
     * 发送HTTP请求
     * @param {String} method - 请求方法 (GET, POST, PUT, DELETE等)
     * @param {String} url - 请求URL
     * @param {Object} headers - 请求头对象
     * @param {String} data - 请求数据
     * @return {Object} HTTP响应对象
     */
    static request(method := "GET", url := "", headers := {}, data := "") {
        try {
            ; 创建WinHTTP请求对象
            win_http_request := ComObject("WinHttp.WinHttpRequest.5.1")
            
            ; 设置超时时间（解析超时、连接超时、发送超时、接收超时）
            ; 每个阶段30秒超时
            resolve_timeout := 30000
            connect_timeout := 30000
            send_timeout := 30000
            receive_timeout := 30000
            win_http_request.SetTimeouts(resolve_timeout, connect_timeout, send_timeout, receive_timeout)
            
            ; 打开HTTP连接
            win_http_request.Open(method, url, true) 
            
            ; 设置请求头
            for header_name, header_value in headers.OwnProps() {
                win_http_request.SetRequestHeader(header_name, header_value)
            }
            
            ; 发送请求并等待响应
            win_http_request.Send(data) 
            win_http_request.WaitForResponse()
            
            ; 检查HTTP状态码
            response_status := win_http_request.Status
            if (response_status >= 400) {
                status_text := win_http_request.StatusText
                Throw Error("HTTP错误 " . response_status . ": " . status_text)
            }
            
            return win_http_request 
        } catch as err {
            Throw Error("HTTP请求失败: " . err.message)
        }
    }
    
    /**
     * 获取文本响应
     * @param {String} method - 请求方法
     * @param {String} url - 请求URL
     * @param {Object} headers - 请求头
     * @param {String} data - 请求数据
     * @return {String} 响应文本
     */
    static get_text(method := "GET", url := "", headers := {}, data := "") {
        try {
            http_response := http_client.request(method, url, headers, data)        
            response_text := http_response.ResponseText
            return response_text
        } catch as err {
            Throw Error("获取文本响应失败: " . err.message)
        }
    }
    
    /**
     * 获取UTF-8编码的响应文本
     * @param {String} method - 请求方法
     * @param {String} url - 请求URL
     * @param {Object} headers - 请求头
     * @param {String} data - 请求数据
     * @return {String} UTF-8编码的响应文本
     */
    static get_text_utf8(method := "GET", url := "", headers := {}, data := "") {
        try {
            http_response := http_client.request(method, url, headers, data)
            
            ; 尝试使用UTF-8解码响应体
            try {
                ; 创建ADO流对象处理UTF-8编码
                ado_stream := ComObject("adodb.stream")
                ado_stream.Type := 1  ; adTypeBinary = 1
                ado_stream.Mode := 3  ; adModeReadWrite = 3
                ado_stream.Open()
                
                ; 写入原始响应数据
                response_body := http_response.ResponseBody
                ado_stream.Write(response_body)
                ado_stream.Position := 0
                
                ; 切换到文本模式并设置UTF-8编码
                ado_stream.Type := 2  ; adTypeText = 2
                ado_stream.Charset := "UTF-8"
                
                ; 读取UTF-8文本数据
                utf8_text := ado_stream.ReadText()
                ado_stream.Close()
                
                return utf8_text
                
            } catch Error as utf8_error {
                ; 如果UTF-8转换失败，回退到普通文本
                fallback_text := http_response.ResponseText
                return fallback_text
            }
            
        } catch as err {
            Throw Error("获取UTF-8响应失败: " . err.message)
        }
    }
    
    /**
     * 发送GET请求
     * @param {String} url - 请求URL
     * @param {Object} headers - 请求头（可选）
     * @return {String} 响应文本
     */
    static get(url, headers := {}) {
        return http_client.get_text("GET", url, headers, "")
    }
    
    /**
     * 发送POST请求
     * @param {String} url - 请求URL
     * @param {String} data - POST数据
     * @param {Object} headers - 请求头（可选）
     * @return {String} 响应文本
     */
    static post(url, data := "", headers := {}) {
        ; 如果没有设置Content-Type，添加默认值
        if (!headers.HasOwnProp("Content-Type")) {
            headers["Content-Type"] := "application/x-www-form-urlencoded"
        }
        return http_client.get_text("POST", url, headers, data)
    }
    
    /**
     * 发送POST请求（UTF-8响应）
     * @param {String} url - 请求URL
     * @param {String} data - POST数据
     * @param {Object} headers - 请求头（可选）
     * @return {String} UTF-8编码的响应文本
     */
    static post_utf8(url, data := "", headers := {}) {
        ; 如果没有设置Content-Type，添加默认值
        if (!headers.HasOwnProp("Content-Type")) {
            headers["Content-Type"] := "application/x-www-form-urlencoded"
        }
        return http_client.get_text_utf8("POST", url, headers, data)
    }
    
     /**
     * 发送JSON数据的POST请求
     * @param {String} url - 请求URL
     * @param {Object} json_obj - 要发送的对象（将被转换为JSON）
     * @param {Object} extra_headers - 额外的请求头（可选）
     * @return {String} 响应文本
     */
    static post_json(url, json_obj, extra_headers := {}) {
        try {
            ; 将对象转换为JSON字符串
            json_string := json.dumps(json_obj) 

            ; 设置JSON请求头
            json_headers := Map()
            json_headers["Content-Type"] := "application/json"
            json_headers["Accept"] := "application/json"
            
            ; 合并额外的请求头
            for header_name, header_value in extra_headers.OwnProps() {
                json_headers[header_name] := header_value
            }
            
            return http_client.get_text_utf8("POST", url, json_headers, json_string)
            
        } catch as err {
            Throw Error("JSON POST请求失败: " . err.message)
        }
    }

    /**
     * 下载文件到本地
     * @param {String} url - 文件URL
     * @param {String} local_path - 本地保存路径
     * @param {Object} headers - 请求头（可选）
     * @return {Boolean} 下载是否成功
     */
    static download_file(url, local_path, headers := {}) {
        try {
            http_response := http_client.request("GET", url, headers, "")
            
            ; 获取响应体（二进制数据）
            response_body := http_response.ResponseBody
            
            ; 创建文件并写入数据
            file_handle := FileOpen(local_path, "w")
            if (!file_handle) {
                Throw Error("无法创建文件: " . local_path)
            }
            
            ; 写入二进制数据
            bytes_written := file_handle.RawWrite(response_body)
            file_handle.Close()
            
            ; 验证写入是否成功
            if (bytes_written <= 0) {
                Throw Error("文件写入失败")
            }
            
            return true
            
        } catch as err {
            ; 确保文件句柄被关闭
            if (file_handle) {
                file_handle.Close()
            }
            Throw Error("文件下载失败: " . err.message)
        }
    }
    
    /**
     * 检查URL是否可访问
     * @param {String} url - 要检查的URL
     * @param {Number} timeout_ms - 超时时间（毫秒，默认5000）
     * @return {Boolean} URL是否可访问
     */
    static is_url_accessible(url, timeout_ms := 5000) {
        try {
            win_http_request := ComObject("WinHttp.WinHttpRequest.5.1")
            
            ; 设置较短的超时时间
            win_http_request.SetTimeouts(timeout_ms, timeout_ms, timeout_ms, timeout_ms)
            win_http_request.Open("HEAD", url, true)  ; 使用HEAD方法，只获取头部
            win_http_request.Send()
            win_http_request.WaitForResponse()
            
            response_status := win_http_request.Status
            return (response_status >= 200 && response_status < 400)
            
        } catch as err {
            return false
        }
    }
    
    /**
     * 获取URL的响应头信息
     * @param {String} url - 目标URL
     * @param {Object} request_headers - 请求头（可选）
     * @return {Object} 响应头对象
     */
    static get_response_headers(url, request_headers := {}) {
        try {
            http_response := http_client.request("HEAD", url, request_headers, "")
            
            ; 获取所有响应头
            all_headers := http_response.GetAllResponseHeaders()
            
            ; 解析响应头为对象
            header_obj := {}
            header_lines := StrSplit(all_headers, "`r`n")
            
            for line_text in header_lines {
                if (line_text != "" && InStr(line_text, ":")) {
                    colon_pos := InStr(line_text, ":")
                    header_name := Trim(SubStr(line_text, 1, colon_pos - 1))
                    header_value := Trim(SubStr(line_text, colon_pos + 1))
                    if (header_name != "") {
                        header_obj[header_name] := header_value
                    }
                }
            }
            
            return header_obj
            
        } catch as err {
            Throw Error("获取响应头失败: " . err.message)
        }
    }
}

; #endregion

; #region ========================= 工具函数 =========================

/**
 * URL编码函数
 * @param {String} input_str - 需要编码的字符串
 * @param {String} safe_chars - 不需要编码的字符（默认："-_."）
 * @param {String} encoding - 编码格式（默认："UTF-8"）
 * @return {String} 编码后的字符串
 */
url_encode(input_str, safe_chars := "-_.", encoding := "UTF-8") {
    try {
        hex_buffer := "00"
        sprintf_func := "msvcrt\swprintf"
        byte_buffer := Buffer(StrPut(input_str, encoding))
        StrPut(input_str, byte_buffer, encoding) 
        encoded_result := ""
        
        Loop {
            current_byte := NumGet(byte_buffer, A_Index - 1, "UChar")
            if (!current_byte)
                break
            current_char := Chr(current_byte)
            
            ; 判断字符是否需要编码
            if (current_byte >= 0x41 && current_byte <= 0x5A ; A-Z
                || current_byte >= 0x61 && current_byte <= 0x7A ; a-z
                || current_byte >= 0x30 && current_byte <= 0x39 ; 0-9
                || InStr(safe_chars, Chr(current_byte), true)) {
                encoded_result .= Chr(current_byte)
            } else {
                DllCall(sprintf_func, "Str", hex_buffer, "Str", "%%%02X", "UChar", current_byte, "Cdecl")
                encoded_result .= hex_buffer
            }
        }
        return encoded_result
        
    } catch as err {
        Msgbox("URL编码失败: " . err.message)
        return input_str  ; 编码失败时返回原字符串
    }
}

/**
 * URL解码函数
 * @param {String} encoded_str - 已编码的字符串
 * @param {String} encoding - 解码格式（默认："UTF-8"）
 * @return {String} 解码后的字符串
 */
url_decode(encoded_str, encoding := "UTF-8") {
    try {
        decoded_result := ""
        string_length := StrLen(encoded_str)
        char_index := 1
        
        while (char_index <= string_length) {
            current_char := SubStr(encoded_str, char_index, 1)
            
            if (current_char = "%") {
                ; 处理百分号编码
                if (char_index + 2 <= string_length) {
                    hex_string := SubStr(encoded_str, char_index + 1, 2)
                    try {
                        byte_value := Integer("0x" . hex_string)
                        decoded_result .= Chr(byte_value)
                        char_index += 3
                    } catch {
                        ; 无效的十六进制，直接添加%
                        decoded_result .= current_char
                        char_index += 1
                    }
                } else {
                    decoded_result .= current_char
                    char_index += 1
                }
            } else if (current_char = "+") {
                ; + 号转换为空格
                decoded_result .= " "
                char_index += 1
            } else {
                decoded_result .= current_char
                char_index += 1
            }
        }
        
        return decoded_result
        
    } catch as err {
        Msgbox("URL解码失败: " . err.message)
        return encoded_str  ; 解码失败时返回原字符串
    }
}

/**
 * MD5哈希函数（用于百度翻译API签名）
 * @param {String} input_string - 输入字符串
 * @return {String} MD5哈希值（32位小写十六进制）
 */
md5_hash(input_string) {
    try {
        ; 使用Windows CryptoAPI计算MD5
        crypto_provider := 0
        hash_object := 0
        
        ; 获取加密服务提供者
        if (!DllCall("advapi32\CryptAcquireContextW", "Ptr*", &crypto_provider, "Ptr", 0, "Ptr", 0, "UInt", 1, "UInt", 0xF0000000)) {
            Throw Error("无法获取加密上下文")
        }
        
        ; 创建哈希对象
        if (!DllCall("advapi32\CryptCreateHash", "Ptr", crypto_provider, "UInt", 0x8003, "Ptr", 0, "UInt", 0, "Ptr*", &hash_object)) {
            DllCall("advapi32\CryptReleaseContext", "Ptr", crypto_provider, "UInt", 0)
            Throw Error("无法创建哈希对象")
        }
        
        ; 转换字符串为UTF-8字节
        utf8_buffer := Buffer(StrPut(input_string, "UTF-8") - 1)
        StrPut(input_string, utf8_buffer, utf8_buffer.size, "UTF-8")
        
        ; 计算哈希
        if (!DllCall("advapi32\CryptHashData", "Ptr", hash_object, "Ptr", utf8_buffer.Ptr, "UInt", utf8_buffer.size, "UInt", 0)) {
            DllCall("advapi32\CryptDestroyHash", "Ptr", hash_object)
            DllCall("advapi32\CryptReleaseContext", "Ptr", crypto_provider, "UInt", 0)
            Throw Error("哈希计算失败")
        }
        
        ; 获取哈希值
        hash_size := 16  ; MD5 = 128bit = 16字节
        hash_buffer := Buffer(hash_size)
        if (!DllCall("advapi32\CryptGetHashParam", "Ptr", hash_object, "UInt", 2, "Ptr", hash_buffer.Ptr, "UInt*", &hash_size, "UInt", 0)) {
            DllCall("advapi32\CryptDestroyHash", "Ptr", hash_object)
            DllCall("advapi32\CryptReleaseContext", "Ptr", crypto_provider, "UInt", 0)
            Throw Error("获取哈希值失败")
        }
        
        ; 转换为十六进制字符串
        hex_result := ""
        Loop hash_size {
            byte_value := NumGet(hash_buffer, A_Index - 1, "UChar")
            hex_result .= Format("{:02x}", byte_value)
        }
        
        ; 清理资源
        DllCall("advapi32\CryptDestroyHash", "Ptr", hash_object)
        DllCall("advapi32\CryptReleaseContext", "Ptr", crypto_provider, "UInt", 0)
        
        return hex_result
        
    } catch as err {
        Msgbox("MD5计算失败: " . err.message)
        return ""
    }
}

; #endregion

; #region ========================= 表单数据构建器 =========================

/**
 * 创建multipart/form-data格式的表单数据（符合 RFC 7578 标准）
 * 用于文件上传等场景
 * @param {Object} retData - 返回的数据（引用传递）
 * @param {String} retHeader - 返回的Content-Type头（引用传递）
 * @param {Object} objParam - 表单参数对象
 */
create_form_data(&retData, &retHeader, objParam) {
    Local crlf := "`r`n", i, k, v, content_str, pv_data
    
    try {
        ; 创建随机边界
        Local boundary := form_data_random_boundary()
        Local boundary_line := "------------------------------" . boundary

        ; 创建基于可移动内存的IStream
        h_data := DllCall("GlobalAlloc", "uint", 0x2, "uptr", 0, "ptr")
        DllCall("ole32\CreateStreamOnHGlobal", "ptr", h_data, "int", False, "ptr*", &p_stream:=0, "uint")
        global form_data_pStream := p_stream

        ; 遍历输入参数
        For k, v in objParam.OwnProps() {
            If IsObject(v) {
                ; 处理文件参数
                For i, file_name in v {
                    ; 检查文件是否存在
                    if !FileExist(file_name) {
                        Throw Error("文件不存在: " . file_name)
                    }
                    
                    content_str := boundary_line . crlf
                        . 'Content-Disposition: form-data; name="' . k . '"; filename="' . file_name . '"' . crlf
                        . 'Content-Type: ' . form_data_mime_type(file_name) . crlf . crlf

                    form_data_str_put_utf8( content_str )
                    form_data_load_from_file( file_name )
                    form_data_str_put_utf8( crlf )
                }
            } Else {
                ; 处理普通文本参数
                content_str := boundary_line . crlf
                    . 'Content-Disposition: form-data; name="' . k '"' . crlf . crlf
                    . v . crlf
                form_data_str_put_utf8( content_str )
            }
        }

        ; 添加结束边界
        form_data_str_put_utf8( boundary_line . "--" . crlf )

        ; 释放流并获取数据
        form_data_pStream := ObjRelease(p_stream)
        p_data := DllCall("GlobalLock", "ptr", h_data, "ptr")
        data_size := DllCall("GlobalSize", "ptr", p_data, "uptr")

        ; 创建字节数组并复制数据
        retData := ComObjArray( 0x11, data_size ) ; 创建 SAFEARRAY = VT_ARRAY|VT_UI1
        pv_data  := NumGet( ComObjValue( retData ), 8 + A_PtrSize , "ptr" )
        DllCall( "RtlMoveMemory", "Ptr", pv_data, "Ptr", p_data, "Ptr", data_size )

        ; 清理资源
        DllCall("GlobalUnlock", "ptr", h_data)
        DllCall("GlobalFree", "Ptr", h_data, "Ptr")

        ; 设置返回的Content-Type头
        retHeader := "multipart/form-data; boundary=----------------------------" . boundary
        
    } catch as err {
        ; 确保资源清理
        if (p_stream) {
            ObjRelease(p_stream)
        }
        if (h_data) {
            DllCall("GlobalUnlock", "ptr", h_data)
            DllCall("GlobalFree", "Ptr", h_data, "Ptr")
        }
        Throw Error("表单数据构建失败: " . err.message)
    }
}

/**
 * 写入UTF-8字符串到流
 * @param {String} input_str - 要写入的字符串
 */
form_data_str_put_utf8( input_str ) {
    try {
        utf8_buf := Buffer(StrPut(input_str, "UTF-8") - 1) ; 移除null终止符
        StrPut(input_str, utf8_buf, utf8_buf.size, "UTF-8")
        DllCall("shlwapi\IStream_Write", "ptr", form_data_pStream, "ptr", utf8_buf.Ptr, "uint", utf8_buf.Size, "uint")
    } catch as err {
        Throw Error("UTF-8字符串写入失败: " . err.message)
    }
}

/**
 * 从文件加载数据到流
 * @param {String} file_path - 文件路径
 */
form_data_load_from_file( file_path ) {
    try {
        ; 检查文件是否存在
        if !FileExist(file_path) {
            Throw Error("文件不存在: " . file_path)
        }
        
        ; 创建文件流
        DllCall("shlwapi\SHCreateStreamOnFileEx"
                    ,   "wstr", file_path
                    ,   "uint", 0x0             ; STGM_READ
                    ,   "uint", 0x80            ; FILE_ATTRIBUTE_NORMAL
                    ,    "int", False            ; fCreate
                    ,    "ptr", 0               ; pstmTemplate
                    ,   "ptr*", &p_file_stream:=0
                    ,   "uint")
        
        if (!p_file_stream) {
            Throw Error("无法创建文件流: " . file_path)
        }
        
        ; 获取文件大小并复制数据
        DllCall("shlwapi\IStream_Size", "ptr", p_file_stream, "uint64*", &file_size:=0, "uint")
        DllCall("shlwapi\IStream_Copy", "ptr", p_file_stream , "ptr", form_data_pStream, "uint", file_size, "uint")
        ObjRelease(p_file_stream)
        
    } catch as err {
        ; 确保释放文件流
        if (p_file_stream) {
            ObjRelease(p_file_stream)
        }
        Throw Error("文件加载失败: " . err.message)
    }
}

/**
 * 生成随机边界字符串
 * @return {String} 16位随机边界字符串
 */
form_data_random_boundary() {
    try {
        ; 使用更安全的随机字符生成
        char_set := "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        random_boundary := ""
        
        ; 生成16位随机字符串（比原来更长更安全）
        Loop 16 {
            char_pos := Random(1, StrLen(char_set))
            random_boundary .= SubStr(char_set, char_pos, 1)
        }
        
        return random_boundary
        
    } catch as err {
        ; 如果随机生成失败，使用时间戳备选方案
        return "FormData" . A_TickCount . A_MSec
    }
}

/**
 * 根据文件内容判断MIME类型
 * @param {String} file_name - 文件路径
 * @return {String} MIME类型字符串
 */
form_data_mime_type(file_name) {
    try {
        ; 检查文件是否存在
        if !FileExist(file_name) {
            return "application/octet-stream"
        }
        
        ; 打开文件读取文件头
        file_handle := FileOpen(file_name, "r")
        if (!file_handle) {
            return "application/octet-stream"
        }
        
        ; 读取文件头魔数
        magic_number := file_handle.ReadUInt()
        file_handle.Close()
        
        ; 根据文件头判断类型
        Return (magic_number        = 0x474E5089) ? "image/png"           ; PNG
            :  (magic_number        = 0x38464947) ? "image/gif"           ; GIF
            :  (magic_number&0xFFFF = 0x4D42    ) ? "image/bmp"           ; BMP
            :  (magic_number&0xFFFF = 0xD8FF    ) ? "image/jpeg"          ; JPEG
            :  (magic_number&0xFFFF = 0x4949    ) ? "image/tiff"          ; TIFF (Intel)
            :  (magic_number&0xFFFF = 0x4D4D    ) ? "image/tiff"          ; TIFF (Motorola)
            :  (magic_number        = 0x46464952) ? "image/webp"          ; WEBP
            :  (magic_number&0xFFFFFF = 0x464449) ? "image/webp"          ; WEBP变体
            :  form_data_get_mime_by_extension(file_name)                 ; 通过扩展名判断
            
    } catch as err {
        ; 出错时通过文件扩展名判断
        return form_data_get_mime_by_extension(file_name)
    }
}

/**
 * 通过文件扩展名判断MIME类型（备用方法）
 * @param {String} file_name - 文件路径
 * @return {String} MIME类型字符串
 */
form_data_get_mime_by_extension(file_name) {
    ; 获取文件扩展名
    SplitPath(file_name, , , &file_ext)
    file_ext := StrLower(file_ext)
    
    ; 常见文件类型映射
    switch file_ext {
        case "jpg", "jpeg":
            return "image/jpeg"
        case "png":
            return "image/png"
        case "gif":
            return "image/gif"
        case "bmp":
            return "image/bmp"
        case "tiff", "tif":
            return "image/tiff"
        case "webp":
            return "image/webp"
        case "svg":
            return "image/svg+xml"
        case "ico":
            return "image/x-icon"
        case "pdf":
            return "application/pdf"
        case "txt":
            return "text/plain"
        case "html", "htm":
            return "text/html"
        case "css":
            return "text/css"
        case "js":
            return "application/javascript"
        case "json":
            return "application/json"
        case "xml":
            return "application/xml"
        case "zip":
            return "application/zip"
        case "rar":
            return "application/x-rar-compressed"
        case "7z":
            return "application/x-7z-compressed"
        case "mp3":
            return "audio/mpeg"
        case "wav":
            return "audio/wav"
        case "mp4":
            return "video/mp4"
        case "avi":
            return "video/x-msvideo"
        case "docx":
            return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        case "xlsx":
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        case "pptx":
            return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        default:
            return "application/octet-stream"
    }
}

; #endregion

