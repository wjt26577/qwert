; ============================================================
;  lib_screen.ahk — 截图 / 区域选择 / Snipaste
; ============================================================

snipaste() {
    try {
        app_snipaste := run_config["app_snipaste"]["payload"]
        command := app_snipaste . " snip --area 1 1 500 500 -o pin"
        run(command)
    } catch as err
        Msgbox(err.Message)
}

screenshot(if_stick := True, if_show := True, back_color := "Purple") {
    save_path := run_config["path_screenshot_path"]["payload"]
    try {
        result := select_region_interactive(back_color)
        if (!result.region)
            return {success: false, file_path: "", message: "用户取消了区域选择"}
        region := StrReplace(result.region, "|", " ")
        if (!DirExist(save_path)) {
            try DirCreate(save_path)
            catch as err
                return {success: false, file_path: "", message: "无法创建保存目录: " . err.message}
        }
        time_stamp := FormatTime(, "yyyyMMdd_HHmmss")
        file_path := Format("{1}\{2}.png", save_path, time_stamp)
        app_snipaste := run_config["app_snipaste"]["payload"]
        command := APP_SNIPASTE . " snip --area " . region . " -o clipboard;pin"
        run(command)
        return { success: true }
    } catch as err {
        Msgbox("截图过程中发生错误: " . err.message)
        return {success: false, file_path: "", message: err.message}
    }
}

get_screen_region(button := "RButton") {
    transparent := 100
    back_color := "Purple"
    selection_gui := 0
    info_gui := 0
    try {
        MouseGetPos(&begin_x, &begin_y)
        selection_gui := Gui()
        selection_gui.Opt("+AlwaysOnTop -Caption +Border +ToolWindow +LastFound -DPIScale")
        selection_gui.BackColor := back_color
        WinSetTransparent(transparent)
        selection_gui.Show()
        info_gui := Gui()
        info_gui.Opt("+AlwaysOnTop -Caption +ToolWindow +LastFound -DPIScale")
        info_gui.BackColor := 0x000000
        info_text := info_gui.Add("Text", "cWhite", "选择区域中...")
        WinSetTransparent(200)
        info_gui.Show("x" . (begin_x + 10) . " y" . (begin_y - 30) . " AutoSize")
        while GetKeyState(button, "P") {
            if GetKeyState("Escape", "P")
                return 0
            Sleep(10)
            MouseGetPos(&end_x, &end_y)
            x := (begin_x < end_x) ? begin_x : end_x
            y := (begin_y < end_y) ? begin_y : end_y
            w := Abs(begin_x - end_x)
            h := Abs(begin_y - end_y)
            selection_gui.Move(x, y, w, h)
            info_text.Text := Format("区域: {}x{} | 位置: ({}, {})", w, h, x, y)
            info_gui.Move(end_x + 15, end_y - 40)
        }
        if (w < 5 || h < 5)
            return {success: false, message: "区域选择太小"}
        else
            return {success: true, x: x, y: y, w: w, h: h}
    } catch as err {
        return {success: false, message: "区域选择错误: " . err.Message}
    } finally {
        if (selection_gui)
            selection_gui.Destroy()
        if (info_gui)
            info_gui.Destroy()
    }
}

get_screen_region_lbutton() {
    return get_screen_region(button := "LButton")
}


/**
 * 交互式区域选择（带时间检测）
 * 
 * @description 立即显示选择框，实时跟踪鼠标拖拽，记录按键时长
 * @param {Number} back_color 选择框背景色（默认：0x800080）
 * @param {Number} transparent 透明度（默认：100）
 * @param {String} button 触发按键（默认："RButton"）
 * @return {Object} {duration: 按键时长(ms), region: 选择区域"x|y|w|h"或0}
 */
select_region_interactive(back_color := "Purple", button := "RButton", transparent := 100) {  
    start_time := A_TickCount    ; 记录开始时间
    region_width := region_height := region_x := region_y := 0        ; 初始化区域变量
    selection_gui := 0                       ; GUI对象
    info_gui := 0                ; 信息显示GUI
     
    try {
        ; 获取起始鼠标位置
        MouseGetPos(&begin_x, &begin_y)
        
        ; 创建选择框GUI
        selection_gui := Gui()
        selection_gui.Opt("+AlwaysOnTop -Caption +Border +ToolWindow +LastFound -DPIScale")
        selection_gui.BackColor := back_color
        WinSetTransparent(transparent)
        selection_gui.Show()
        
        ; 创建信息显示GUI
        info_gui := Gui()
        info_gui.Opt("+AlwaysOnTop -Caption +ToolWindow +LastFound -DPIScale")
        info_gui.BackColor := 0x000000
        info_text := info_gui.Add("Text", "cWhite", "选择区域中...")
        WinSetTransparent(200)
        info_gui.Show("x" . (begin_x + 10) . " y" . (begin_y - 30) . " AutoSize")
        
        ; 实时更新选择框，直到按键释放
        while GetKeyState(button, "P") {
            ; 检查ESC键取消
            if GetKeyState("Escape", "P") {
                return {duration: A_TickCount - start_time, region: ""}
            }
            
            Sleep(10)
            MouseGetPos(&end_x, &end_y)
            
            ; 计算选择区域
            region_width := Abs(begin_x - end_x)
            region_height := Abs(begin_y - end_y)
            region_x := (begin_x < end_x) ? begin_x : end_x
            region_y := (begin_y < end_y) ? begin_y : end_y
            
            ; 移动选择框
            selection_gui.Move(region_x, region_y, region_width, region_height)
            
            ; 更新信息显示
            info_text.Text := Format("区域: {}x{} | 位置: ({}, {})", region_width, region_height, region_x, region_y)
            info_gui.Move(end_x + 15, end_y - 40)
        }
        
        ; 计算总按键时长
        total_duration := A_TickCount - start_time
        
        ; 检查区域大小并返回结果
        if (region_width < MIN_REGION_SIZE || region_height < MIN_REGION_SIZE) {
            return {duration: total_duration, region: 0}
        } else {
            return {duration: total_duration, region: Format("{}|{}|{}|{}", region_x, region_y, region_width, region_height)} 
        }
        
    } catch as err {
        Msgbox("区域选择错误: " . err.Message)
        return {duration: A_TickCount - start_time, region: 0}
    } finally {
        ; 确保GUI被销毁
        if (selection_gui) {
            selection_gui.Destroy()
        }
        if (info_gui) {
            info_gui.Destroy()
        }
    }
}
