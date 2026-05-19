; ============================================================
;  lib_string.ahk — 填充 / 重复 / 随机文本 / 相似度
; ============================================================

; =================== 重复字符串 ===================

repeat_str(char, n) {
    if (n <= 0)
        return ""
    static SPACES := "                                                                "
    if (char = " " && n <= 64)
        return SubStr(SPACES, 1, n)
    result := ""
    while (n > 0) {
        if (Mod(n, 2))
            result .= char
        char .= char
        n //= 2
    }
    return result
}

repeat_string(s, n) {
    try {
        result := ""
        Loop n
            result := result . s
        return result
    } catch as err
        MsgBox(err.Message)
}

; =================== 按数量填充 ===================

pad_left_by_count(str, count := 1, pad_char := " ") {
    return (count <= 0) ? str : repeat_str(pad_char, count) . str
}

pad_right_by_count(str, count := 1, pad_char := " ") {
    return (count <= 0) ? str : str . repeat_str(pad_char, count)
}

pad_both_by_count(str, count := 1, pad_char := " ") {
    return (count <= 0) ? str : repeat_str(pad_char, count) . str . repeat_str(pad_char, count)
}

; =================== 按总宽度填充 ===================

pad_left_to_width(str, width, pad_char := " ") {
    len := StrLen(str)
    return (len >= width) ? str : repeat_str(pad_char, width - len) . str
}

pad_right_to_width(str, width, pad_char := " ") {
    len := StrLen(str)
    return (len >= width) ? str : str . repeat_str(pad_char, width - len)
}

pad_center_to_width(str, width, pad_char := " ") {
    len := StrLen(str)
    if (len >= width)
        return str
    total_pad := width - len
    left_pad := total_pad // 2
    right_pad := total_pad - left_pad
    return repeat_str(pad_char, left_pad) . str . repeat_str(pad_char, right_pad)
}

pad_left(text, total_length, pad_char := " ") {
    len := StrLen(text)
    return (len >= total_length) ? text : repeat_str(pad_char, total_length - len) . text
}

pad_right(text, total_length, pad_char := " ") {
    len := StrLen(text)
    return (len >= total_length) ? text : text . repeat_str(pad_char, total_length - len)
}

pad_center(text, total_length, pad_char := " ") {
    len := StrLen(text)
    if (len >= total_length)
        return text
    pad_total := total_length - len
    pad_left_count := pad_total // 2
    pad_right_count := pad_total - pad_left_count
    return repeat_str(pad_char, pad_left_count) . text . repeat_str(pad_char, pad_right_count)
}

string_pad(text, length, char := " ", direction := "right") {
    current_length := StrLen(text)
    if (current_length >= length)
        return text
    padding_count := length - current_length
    if (padding_count <= 0)
        return text
    if (padding_count == 1) {
        padding := char
    } else {
        padding := ""
        temp_char := char
        remaining := padding_count
        while (remaining > 0) {
            if (Mod(remaining, 2) == 1)
                padding .= temp_char
            temp_char .= temp_char
            remaining //= 2
        }
    }
    return (direction == "left") ? (padding . text) : (text . padding)
}

remove_spaces(str_input) {
    return RegExReplace(str_input, " ", "")
}

; =================== 随机字符串 ===================

generate_random_string(string_length, char_set := "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") {
    try {
        random_result := ""
        char_set_length := StrLen(char_set)
        Loop string_length {
            random_index := Random(1, char_set_length)
            random_result .= SubStr(char_set, random_index, 1)
        }
        return random_result
    } catch as err {
        Msgbox("生成随机字符串失败: " . err.message)
        return ""
    }
}

; =================== 随机中文文本 ===================

generate_random_text(length := 8, include_punctuation := false) {
    static common_chars :=
    "的一是在不了有和人这中大为上个国我以要他时来用们生到作地于出就分对成会可主发年动同工也能下过子说" .
    "产种面而方后多定行学法所民得经十三之进着等部度家电力里如水化高自二理起小物现实量都两体制机当使点" .
    "从业本去把性好应开它合还因由其些然前外天政四日那社义事平形相全表间样与关各重新线内数正心反你明吗" .
    "看原又么利比或但质气第向道命此变条只没结解问意建月公无系军很情者最立代想已通并提直题党程展五果您" .
    "料象员革位入常文总次品式活设及管特件长求老头基资边流路级少图山统接知较将组见计别她手角期根论运击" .
    "农指几九区强放决西被干做必战先回则任取完据处南目确界领表志入世计记特率专历究拉声步类古克敌胡司千" .
    "私交烟称构光维即百达精列死坚师听划感示活改严什话术真至信调引争容究况须持布越织具非必型翻飞采收跟" .
    "白且许广马今写充半往占卡早范座另诉施创坐兴抓坏松台若苦透终聚迷骨孩双余丰突负刻攻研色减单显候段划"
    static all_punctuations := "，。"
    static end_punctuations := "。"

    char_len := StrLen(common_chars)
    all_punct_len := StrLen(all_punctuations)
    end_punct_len := StrLen(end_punctuations)
    result_array := []
    current_len := 0
    punct_interval := 12

    while (current_len < length) {
        insert_punct := false
        if (include_punctuation && current_len > 3 && current_len < length - 1)
            if (mod(current_len, punct_interval) = 0 || random(1, 15) = 1)
                insert_punct := true
        if (insert_punct) {
            p_index := random(1, all_punct_len)
            result_array.push(SubStr(all_punctuations, p_index, 1))
        } else {
            c_index := random(1, char_len)
            result_array.push(SubStr(common_chars, c_index, 1))
        }
        current_len++
    }
    if (include_punctuation) {
        last_char := result_array.pop()
        if (InStr(end_punctuations, last_char))
            result_array.push(last_char)
        else {
            e_index := random(1, end_punct_len)
            result_array.push(SubStr(end_punctuations, e_index, 1))
        }
    }
    return join_array(result_array)
}

generate_random_text2(length := 50) {
    chinese_chars := ("好天我你他可有说地时多子中不上来小王年和风生开出行里出要会"
                      . "点水得白书什术件常文无元些省几社言平精又气清正行种须养容身"
                      . "照论速收族温难委具队北热节总完置感界列选包根故孩整")
    random_text := ""
    Loop length {
        random_index := Random(1, StrLen(chinese_chars))
        random_text .= SubStr(chinese_chars, random_index, 1)
    }
    return random_text
}

generate_random_paragraph(target_length := 50) {
    static one_char_words := ["是","对","不","到","已","可","有","在","从","向","以","为","和","与","上","下","前","后","左","右","能","应","将","要","需","因","由","此","但","且"]
    static two_char_words := ["战略","市场","产品","管理","运营","发展","创新","改革","增长","提升","优化","调整","布局","拓展","深化","推进","加强","完善","资源","客户","品牌","效率","风险","平台","数据","服务","融资","核心","竞争","规划","组织","反馈"]
    static tone_words := ["进一步","持续","稳步","不断","积极","深入","全面","有效","加快","强化","推动","促进"]

    result_array := []
    current_len := 0
    next_comma_count := random(4, 8)
    phrase_count := 0
    one_count := one_char_words.length
    two_count := two_char_words.length
    tone_count := tone_words.length

    while (current_len < target_length) {
        use_tone := (random(1, 10) <= 3)
        temp_tone := ""
        temp_core := ""
        if (use_tone && tone_count > 0)
            temp_tone := tone_words[random(1, tone_count)]
        prefer_two := (random(1, 3) != 1)
        if (prefer_two && two_count > 0)
            temp_core := two_char_words[random(1, two_count)]
        else if (one_count > 0)
            temp_core := one_char_words[random(1, one_count)]
        else if (two_count > 0)
            temp_core := two_char_words[random(1, two_count)]

        temp_phrase := temp_tone . temp_core
        temp_phrase_len := StrLen(temp_phrase)
        will_insert_comma := (phrase_count + 1 >= next_comma_count)
        required_space := temp_phrase_len + 1
        if (will_insert_comma)
            required_space += 1
        if (current_len + required_space > target_length) {
            if (use_tone && StrLen(temp_core) > 0) {
                temp_phrase := temp_core
                temp_phrase_len := StrLen(temp_core)
                required_space := temp_phrase_len + 1
                if (will_insert_comma)
                    required_space += 1
            }
            if (current_len + required_space > target_length)
                break
        }
        result_array.push(temp_phrase)
        current_len += temp_phrase_len
        phrase_count += 1
        if (will_insert_comma) {
            if (current_len + 1 < target_length) {
                result_array.push("，")
                current_len += 1
                phrase_count := 0
                next_comma_count := random(4, 8)
            }
        }
    }

    last_item := result_array.length > 0 ? result_array.pop() : ""
    if (last_item = "，") {
        result_array.push("。")
    } else {
        if (current_len < target_length) {
            result_array.push(last_item)
            result_array.push("。")
        } else {
            result_array.push("。")
        }
    }
    return join_array(result_array)
}

; =================== 相似度 ===================

calculate_string_similarity(string1, string2) {
    try {
        if (string1 = string2)
            return 1.0
        if (string1 = "" || string2 = "")
            return 0.0
        length1 := StrLen(string1)
        length2 := StrLen(string2)
        distance_matrix := []
        Loop length1 + 1 {
            row_array := []
            Loop length2 + 1
                row_array.Push(0)
            distance_matrix.Push(row_array)
        }
        Loop length1 + 1
            distance_matrix[A_Index][1] := A_Index - 1
        Loop length2 + 1
            distance_matrix[1][A_Index] := A_Index - 1
        Loop length1 {
            i := A_Index
            Loop length2 {
                j := A_Index
                cost := (SubStr(string1, i, 1) = SubStr(string2, j, 1)) ? 0 : 1
                distance_matrix[i + 1][j + 1] := Min(distance_matrix[i][j + 1] + 1, Min(distance_matrix[i + 1][j] + 1, distance_matrix[i][j] + cost))
            }
        }
        return 1.0 - (distance_matrix[length1 + 1][length2 + 1] / Max(length1, length2))
    } catch as err {
        Msgbox("相似度计算失败: " . err.message)
        return 0.0
    }
}

; =================== 工具 ===================

join_array(arr) {
    result := ""
    for item in arr
        result .= item
    return result
}

trim_array(arr) {
    for index, value in arr
        arr[index] := Trim(value)
    return arr
}

format_timestamp(timestamp, date_format := "yyyy-MM-dd HH:mm:ss") {
    try {
        return FormatTime(timestamp // 1000, date_format)
    } catch as err {
        Msgbox("时间格式化失败: " . err.message)
        return ""
    }
}

