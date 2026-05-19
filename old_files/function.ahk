

set_global(Params*) {
    global  
    Loop (Params.Length / 2){
        local var_name := Params[2 * A_Index - 1]
        local var_value := Params[2 * A_Index]
        %var_name% := var_value
    }
}

show_thishotkey() {
   hotkey := StrReplace(A_ThisHotkey, "&", "&&")
   Notify.show(hotkey, "error", "500")
}



/**
 * 安全检查变量是否存在
 */
variable_exists(name) {

    try {
        %name%
        return true
    } catch {
        return false
    }
}

show_win_info() {
    ; 获取窗口信息
    
    t1 := WinGetTitle("A")
    t2 := WinGetClass("A")
    t3 := WinGetProcessName("A")
    t4 := WinGetPID("A")
    cur_win := WinExist("A") ; 获取当前活动窗口句柄

    A_Clipboard := "ahk_id " . cur_win
    A_Clipboard := "ahk_pid " . t4
    A_Clipboard := "ahk_exe " . t3
    A_Clipboard := "ahk_class " . t2
    A_Clipboard := t1


    ; 拼接文本 (保持原有的换行格式)
    win_data_text := t1 . "`n"
                  . "ahk_class " . t2 . "`n"
                  . "ahk_exe " . t3 . "`n"
                  . "ahk_pid " . t4 . "`n"
                  . "ahk_id " . cur_win

    ; 弹窗显示
    MsgBox(win_data_text, "窗口信息")
}

; --- 使用示例 ---
; 绑定快捷键 Ctrl+Alt+I
; ^!i::show_win_info()

; #region path_info

 ;  PathU v0.91 by SKAN for ah2 on D35E/D68M @ autohotkey.com/r?p=535235  用于生成一个唯一的文件路径,如果已有同名文件加上（2）
PathU(Filename) {
    Local  OutFile := Format("{:260}", "")
    DllCall("Kernel32\GetFullPathNameW", "str",Filename, "uint",260, "str",OutFile, "ptr",0)
    DllCall("Shell32\PathYetAnotherMakeUniqueName", "str",OutFile, "str",OutFile, "ptr",0, "ptr",0)
    Return OutFile
}


; "D:\ppt\Raven\01 永辉超市年货节\源文件.pptx"

; Drive   = "D:"
; Dir     = "\ppt\Raven\01 永辉超市年货节\"
; Fname   = "源文件"
; Ext     = ".pptx"
; Folder  = "D:\ppt\Raven\01 永辉超市年货节\"
; File    = "源文件.pptx"
; Full    = "D:\ppt\Raven\01 永辉超市年货节\源文件.pptx"



; path_info() v0.67 by SKAN for ah2 on D34U/D68M @ autohotkey.com/r?p=120582
; path_info(Path, X*) drive dir fname ext folder file full  
; dir 不含驱动器，folder含驱动器，fname不含后缀，file含后缀, Ext含点
path_info(Path, X*) {
    Local  K,V,N,U, Dr,   Di,   Fn,   Ex
        ,  FPath := Dr := Di := Fn := Ex := Format("{:260}", "")

    U := Map(),  U.Default := "",  U.CaseSense := 0

    For  K, V  in  X
         N     :=  StrSplit(V, ":",, 2)  ;  split X into Key and Value
      ,  K     :=  SubStr(N[1], 1, 2)    ;  reduce Key to leading 2 chars
      ,  U[K]  :=  N[2]                  ;  assign Key and Value to Map
    DllCall("Kernel32\GetFullPathNameW", "str",Trim(Path,Chr(34)), "uint",260, "str",FPath, "ptr",0)
    ; DllCall("Msvcrt\_wsplitpath", "str",FPath, "str",Dr, "str",Di, "str",Fn, "str",Ex)
    DllCall("Msvcrt\_wsplitpath", "str", FPath, "str", Dr, "str", Di, "str", Fn, "str", Ex, "cdecl")

    Return {  Drive  :  Dr  :=  U["Dr"] ? U["Dr"] : Dr
           ,  Dir    :  Di  :=  U["dp"] ( U["Di"] ? U["Di"] : Di ) U["ds"]
           ,  Fname  :  Fn  :=  U["fp"] ( U["Fn"]!="" ? U["Fn"] : Fn )  U["fs"]
           ,  Ext    :  Ex  :=  U["*E"]!="" ? ( Ex ? Ex : U["*E"] ) : ( U["Ex"]!="" ? U["Ex"] : Ex )
           ,  Folder :  Dr Di
           ,  File   :  Fn Ex
           ,  Full   :  U["pp"] ( Dr Di Fn Ex ) U["ps"] }
}


/**********************************************************
    当 mode = 0, 仅当选中了文件或文件夹时返回路径
    当 mode = 1, 选中文件或文件夹时返回完整路径+文件名，否则返回当前目录
    当 mode = 2, 选中文件或文件夹时返回完整路径+文件名，否则返回空值
    当 mode = 默认(空), 返回当前路径(目录名)
    path_info(Path, X*) drive dir fname ext folder file full  
    dir 不含驱动器，folder含驱动器，fname不含后缀，file含后缀
******************************************************************/
test_path_info() {
    ; MsgBox( path_info(A_ScriptName, "Ext: .ddd").Full,          "Change extension"    )   
    ; MsgBox( path_info(A_ScriptName).Ext,          "Change extension"    )   
; MsgBox( PathU(A_ScriptFullPath) )               ; to copy a backup of this script
; MsgBox( PathU(A_ScriptName) )                   ; A_WorkingDir is used to resolve fullpath
; MsgBox( PathU(A_AhkPath . "\..\License.txt") )  ; relative path is auto-resolved
; MsgBox( PathU(A_ScriptDir) )                    ; Works for folder too..Avoid trailing slash!
MsgBox( PathU("D:\常小二\赛百味品牌推广PPT-常小二PPT.zip") )            ; Simple way to create a temp file

}



; #endregion


; #region Mouse Position Detection Library

; Basic Position Detection
; ========================

mouse_get_pos(&x := "", &y := "") {
    MouseGetPos &x, &y
    return {x: x, y: y}
}

mouse_is_at(x, y, tolerance := 0) {
    MouseGetPos &mx, &my
    return (Abs(mx - x) <= tolerance && Abs(my - y) <= tolerance)
}

mouse_is_in_rect(x1, y1, x2, y2) {
    MouseGetPos &mx, &my
    return (mx >= x1 && mx <= x2 && my >= y1 && my <= y2)
}

; Screen Edge Detection
; ====================

mouse_is_at_edge(edge_size := 5) {
    MouseGetPos &x, &y
    return (x <= edge_size || y <= edge_size || 
            x >= A_ScreenWidth - edge_size || 
            y >= A_ScreenHeight - edge_size)
}

mouse_is_at_top(edge_size := 15) {
    MouseGetPos , &y
    return (y <= edge_size)
}

mouse_is_at_bottom(edge_size := 15) {
    MouseGetPos , &y
    return (y >= A_ScreenHeight - edge_size)
}

mouse_is_at_left(edge_size := 15) {
    MouseGetPos &x
    return (x <= edge_size)
}

mouse_is_at_right(edge_size := 15) {
    MouseGetPos &x
    return (x >= A_ScreenWidth - edge_size)
}

mouse_is_at_corner(corner_size := 50) {
    MouseGetPos &x, &y
    return ((x <= corner_size || x >= A_ScreenWidth - corner_size) && 
            (y <= corner_size || y >= A_ScreenHeight - corner_size))
}

; Window Detection
; ================

mouse_is_over_window(win_title := "", win_text := "", exclude_title := "", exclude_text := "") {
    MouseGetPos , , &id
    try {
        return WinExist(win_title . " ahk_id " . id, win_text, exclude_title, exclude_text)
    } catch {
        return false
    }
}

mouse_is_over_active_window() {
    MouseGetPos , , &id
    try {
        return (id == WinGetID("A"))
    } catch {
        return false
    }
}

mouse_get_window_under() {
    MouseGetPos , , &id
    try {
        return {
            id: id,
            title: WinGetTitle("ahk_id " . id),
            class: WinGetClass("ahk_id " . id),
            process: WinGetProcessName("ahk_id " . id)
        }
    } catch {
        return false
    }
}

; Control Detection
; =================

mouse_is_over_control(class_nn, exact := true) {
    try {
        MouseGetPos , , , &control
    } catch {
        return false
    }
    
    if (exact) {
        return (class_nn == control)
    }
    return InStr(control, class_nn) > 0
}

mouse_get_control_under() {
    MouseGetPos , , , &control, &control_hwnd
    try {
        return {
            class_nn: control,
            hwnd: control_hwnd,
            text: ControlGetText(control_hwnd)
        }
    } catch {
        return false
    }
}

; Multi-Monitor Support
; ====================

mouse_get_monitor() {
    MouseGetPos &x, &y
    return MonitorGet(MonitorGetPrimary(), &left, &top, &right, &bottom) ? 
           MonitorGetWorkArea(MonitorGetPrimary(), &work_left, &work_top, &work_right, &work_bottom) : false
}

mouse_is_on_monitor(monitor_num) {
    MouseGetPos &x, &y
    try {
        MonitorGet(monitor_num, &left, &top, &right, &bottom)
        return (x >= left && x <= right && y >= top && y <= bottom)
    } catch {
        return false
    }
}

; Advanced Detection
; ==================

mouse_distance_from(x, y) {
    MouseGetPos &mx, &my
    return Sqrt((mx - x)**2 + (my - y)**2)
}

mouse_is_moving() {
    static last_x := 0, last_y := 0
    MouseGetPos &x, &y
    
    if (x != last_x || y != last_y) {
        last_x := x, last_y := y
        return true
    }
    return false
}

mouse_is_idle(idle_time_ms := 1000) {
    return (A_TimeIdlePhysical >= idle_time_ms)
}

; Utility Functions
; =================

mouse_normalize_coords(&x, &y, from_mode := "Screen", to_mode := "Client", win_id := "") {
    if (from_mode == "Screen" && to_mode == "Client") {
        try {
            target_id := win_id ? win_id : WinGetID("A")
            WinGetPos &win_x, &win_y, , , "ahk_id " . target_id
            x -= win_x, y -= win_y
        }
    } else if (from_mode == "Client" && to_mode == "Screen") {
        try {
            target_id := win_id ? win_id : WinGetID("A")
            WinGetPos &win_x, &win_y, , , "ahk_id " . target_id
            x += win_x, y += win_y
        }
    }
}

mouse_is_in_client_area(win_id := "") {
    MouseGetPos &x, &y, &id
    target_id := win_id ? win_id : id
    
    try {
        WinGetPos &win_x, &win_y, &win_w, &win_h, "ahk_id " . target_id
        WinGetClientPos &client_x, &client_y, &client_w, &client_h, "ahk_id " . target_id
        
        return (x >= win_x + client_x && x <= win_x + client_x + client_w &&
                y >= win_y + client_y && y <= win_y + client_y + client_h)
    } catch {
        return false
    }
}

; Button State Detection
; ======================

mouse_button_is_down(button := "LButton") {
    return GetKeyState(button, "P")
}

mouse_any_button_down() {
    return (GetKeyState("LButton", "P") || 
            GetKeyState("RButton", "P") || 
            GetKeyState("MButton", "P") ||
            GetKeyState("XButton1", "P") || 
            GetKeyState("XButton2", "P"))
}

; Performance Optimized Variants
; ==============================

mouse_quick_edge_check() {
    ; Ultra-fast edge detection using single MouseGetPos call
    MouseGetPos &x, &y
    return (x < 5 || y < 5 || x > A_ScreenWidth - 5 || y > A_ScreenHeight - 5)
}

mouse_cached_position(cache_time_ms := 50) {
    ; Cache position to reduce MouseGetPos calls
    static cached_x := 0, cached_y := 0, last_check := 0
    
    current_time := A_TickCount
    if (current_time - last_check > cache_time_ms) {
        MouseGetPos &cached_x, &cached_y
        last_check := current_time
    }
    
    return {x: cached_x, y: cached_y}
}


; #endregion



; #region 点击图片 

;模拟点击图片
ClickPicture(ImageFilePath, ClickCount:=1, Speed:=0, vReturn:=true, ShowError:=true) {
	lPos := GetPicturePosition(ImageFilePath)
	if (lPos) {
		posX := lPos[1]
		posY := lPos[2]
		ClickPosition(posX, posY, ClickCount, Speed,,vReturn)
		return [posX, posY]
	} else {
		if (ShowError) {
			MsgBox("找不到图片:`n" . ImageFilePath)
		}
		return false
	}
}

;模拟点击指定位置
ClickPosition(posX, posY, ClickCount:=1, Speed:=50, vCoordMode:="Screen", vReturn:=true) {
	;若使用相对模式
	if (CoordMode = "Relative") {
		CoordMode("Mouse", "Screen")
		MouseGetPos(&posX_i, &posY_i) ;保存原来的鼠标位置

		if (ClickCount) {
			MouseClick "Left", posX, posY, ClickCount, Speed, "R"    ;点击相对位置
		} else {
			MouseMove(posX, posY, Speed)
			MsgBox "move"
		}
	} else { ;若使用其他模式
		CoordMode("Mouse", vCoordMode)
		MouseGetPos(&posX_i, &posY_i) ;保存原来的鼠标位置

		if (ClickCount) {
			sleep 100
			MouseClick "Left", posX, posY, ClickCount, Speed
			; sleep 1000
			; MouseClick "Left"
		} else {
			MouseMove(posX, posY, Speed)
		}
	}

	;是否点击后返回
	if (vReturn) {
		MouseMove(posX_i, posY_i, Speed)
	}

	return
}

;获取图片位置
GetPicturePosition(ImageFilePath) {
	;创建一个GUI窗口，将图片作为GUI的一部分
	myGui := Gui()
	myGui.AddPicture(, ImageFilePath)
	MyGui.GetPos(,, &width, &height)

	;进行图片搜索
	;~ FoundX := 0, FoundY := 0
	CoordMode("Pixel")
	ImageSearch(&FoundX, &FoundY, -2000, -2000, 5000, 5000, ImageFilePath)
	CoordMode("Mouse")

	if (FoundX) {
		return [FoundX+width//2, FoundY+height//2]
	} else {
		return FoundX
	}
}




; #endregion



; #region Windows系统功能


; 打开设备管理器
open_device_manager() {
    Run "devmgmt.msc"
}

; 打开控制面板
open_control_panel() {
    Run "control.exe"
}

; 打开任务管理器
open_task_manager() {
    Run "taskmgr.exe"
}

; 打开系统信息
open_system_information() {
    Run "msinfo32.exe"
}

; 打开事件查看器
open_event_viewer() {
    Run "eventvwr.msc"
}

; 打开服务管理器
open_services_manager() {
    Run "services.msc"
}

; 打开计算机管理
open_computer_management() {
    Run "compmgmt.msc"
}

; 打开磁盘管理
open_disk_management() {
    Run "diskmgmt.msc"
}

; 打开网络连接
open_network_connections() {
    Run "ncpa.cpl"
}

; 打开用户账户控制设置
open_user_account_control_settings() {
    Run "useraccountcontrolsettings.exe"
}

; 打开防火墙设置
open_firewall_settings() {
    Run "firewall.cpl"
}

; 打开电源选项
open_power_options() {
    Run "powercfg.cpl"
}

; 打开声音设置
open_sound_settings() {
    Run "mmsys.cpl"
}

; 打开显示设置
open_display_settings() {
    Run "desk.cpl"
}

; 打开打印机和传真
open_printers_and_faxes() {
    Run "control printers"
}

; 打开注册表编辑器
open_registry_editor() {
    Run "regedit.exe"
}

; 打开命令提示符
open_command_prompt() {
    Run "cmd.exe"
}

; 打开记事本
open_notepad() {
    Run "notepad.exe"
}

; 打开计算器
open_calculator() {
    Run "calc.exe"
}


; #endregion


; #region 字符串操作和复制粘贴

; ┌───────────────┐
; │ 按填充数量 (count) │
; └───────────────┘

pad_left_by_count(str, count := 1, pad_char := " ") {
    return (count <= 0) ? str : repeat_str(pad_char, count) . str
}

pad_right_by_count(str, count := 1, pad_char := " ") {
    return (count <= 0) ? str : str . repeat_str(pad_char, count)
}

pad_both_by_count(str, count := 1, pad_char := " ") {
    return (count <= 0) ? str : repeat_str(pad_char, count) . str . repeat_str(pad_char, count)
}


; ┌───────────────┐
; │ 按总长度 (width)  │
; └───────────────┘

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


; ┌───────────────┐
; │ 高效重复字符串工具 │
; └───────────────┘

repeat_str(char, n) {
    if (n <= 0)
        return ""
    static SPACES := "                                                                " ; 64 spaces
    if (char = " " && n <= 64)
        return SubStr(SPACES, 1, n)
    ; 通用倍增法（支持任意字符/字符串）
    result := ""
    while (n > 0) {
        if (Mod(n, 2))
            result .= char
        char .= char
        n //= 2
    }
    return result
}




; 左填充至指定总长度
pad_left(text, total_length, pad_char := " ") {
    len := StrLen(text)
    return (len >= total_length) ? text : repeat_str(pad_char, total_length - len) . text
}

; 右填充至指定总长度
pad_right(text, total_length, pad_char := " ") {
    len := StrLen(text)
    return (len >= total_length) ? text : text . repeat_str(pad_char, total_length - len)
}

; 居中填充（双边）至指定总长度
pad_center(text, total_length, pad_char := " ") {
    len := StrLen(text)
    if (len >= total_length)
        return text
    pad_total := total_length - len
    pad_left_count := pad_total // 2
    pad_right_count := pad_total - pad_left_count
    return repeat_str(pad_char, pad_left_count) . text . repeat_str(pad_char, pad_right_count)
}







/**
 * 更高效的字符串填充（使用重复字符串）
 * @param {String} text - 要填充的文本
 * @param {Number} length - 目标长度
 * @param {String} char - 填充字符（默认空格）
 * @param {String} direction - 填充方向："right"或"left"
 * @return {String} 填充后的字符串
 */
string_pad(text, length, char := " ", direction := "right") {
    current_length := StrLen(text)
    if (current_length >= length) {
        return text
    }
    
    ; 使用更高效的方法：重复字符串
    padding_count := length - current_length
    if (padding_count <= 0) {
        return text
    }
    
    ; 生成足够长的填充字符串（一次性生成，避免循环）
    if (padding_count == 1) {
        padding := char
    } else {
        ; 使用字符串重复技巧
        padding := ""
        temp_char := char
        remaining := padding_count
        
        ; 二进制倍增法（更高效）
        while (remaining > 0) {
            if (Mod(remaining, 2) == 1) {
                padding .= temp_char
            }
            temp_char .= temp_char
            remaining //= 2
        }
    }
    
    return (direction == "left") ? (padding . text) : (text . padding)
}




copy_and_comment_line() => SendInput("^c^q{End}{Enter}^v")

delete_all_right()        => SendInput('+{End}{Backspace}')
delete_all_left()  => SendInput('+{Home}{Backspace}')
delete_to_page_beginning()  => SendInput('+^{Home}{Backspace}')
delete_to_page_end()        => SendInput('+^{End}{Backspace}')
enter_wherever()            => SendInput('{End}{Enter}')


ppt_code_1 := "
(

    ppt_application := ComObjActive("PowerPoint.Application")
    selection := ppt_application.ActiveWindow.Selection
    
    ; 检查选择类型是否为形状
    if (selection.Type < 2) { 
        Msgbox("请先选择要复制的形状")
        return
    }
    
    if (selection.HasChildShapeRange)
        shape_range := selection.ChildShapeRange
    else    
        shape_range := selection.ShapeRange

)"


send_ppt_code_1() {
    ppt_code_1 := "
    (

        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection
        
        ; 检查选择类型是否为形状
        if (selection.Type < 2) { 
            Msgbox("请先选择要复制的形状")
            return
        }
        
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else    
            shape_range := selection.ShapeRange

    )"
    send_by_clipboard(ppt_code_1)
}




delete_line() {	
    SendInput("{Home}+{End}{Delete}")
}



repeat_string(s, n) {	
	try {
		result  := ""
		Loop n
			result := result . s
		return result		
	}
	catch as err
		MsgBox(err.Message)   
}


paste_word() {	
	MouseClick("Left", , , 2)
    SendInput("^v")
}





; #endregion


; #region 输入法

; switch_keyboard_layout() {	
;     if get_keyboard_layout() = "67699721" {
; 		set_chinese_keyboard_layout()	
; 	} else { 
; 		set_english_keyboard_layout()
; 	}
; }

; set_chinese_keyboard_layout() {		
; 	GroupAdd "Special_ime_window", "ahk_class le_appstore_window_class" 
; 	GroupAdd "Special_ime_window", "ahk_class #32770" 
; 	GroupAdd "Special_ime_window", "ahk_class AutoHotkeyGUI" 
; 	GroupAdd "Special_ime_window", "ahk_class Progman" 
; 	GroupAdd "Special_ime_window", "ahk_class WorkerW" 
; 	WM_INPUTLANGCHANGEREQUEST := 0x0050
; 	ENGLISH_INPUT_LAYOUT := "67699721"
; 	CHINESE_INPUT_LAYOUT := "134481924"

; 	if WinActive("ahk_group Special_ime_window") {
; 		target_window := ControlGetFocus("A")
; 	} else {
; 		target_window := "A" 
; 	}
	
; 	try {		
; 		PostMessage(WM_INPUTLANGCHANGEREQUEST, 0, CHINESE_INPUT_LAYOUT, , target_window)
;     }
;     catch as err
;         Msgbox(err.Message)  
; }

; set_english_keyboard_layout() {	
; 	GroupAdd "Special_ime_window", "ahk_class le_appstore_window_class" 
; 	GroupAdd "Special_ime_window", "ahk_class #32770" 
; 	GroupAdd "Special_ime_window", "ahk_class AutoHotkeyGUI" 
; 	GroupAdd "Special_ime_window", "ahk_class Progman" 
; 	GroupAdd "Special_ime_window", "ahk_class WorkerW" 
; 	WM_INPUTLANGCHANGEREQUEST := 0x0050
; 	ENGLISH_INPUT_LAYOUT := "67699721"
; 	CHINESE_INPUT_LAYOUT := "134481924"

; 	if WinActive("ahk_group Special_ime_window") {
; 		target_window := ControlGetFocus("A")
; 	} else {
; 		target_window := "A" 
; 	}
	
; 	try {		
; 		PostMessage(WM_INPUTLANGCHANGEREQUEST, 0, 	ENGLISH_INPUT_LAYOUT, , target_window)
;     }
;     catch as err
;         Msgbox(err.Message)  
; }


; ; 切换到英文输入法，并关闭大写锁定
; english_input() {
;     try {		
; 		global superkey := 0
;         switch_input_layout("en") ; 切换到英文输入法
; 		SetCapsLockState false		
;         Notify.show("English Input Method", 0.5, "Purple") 
;     }
;     catch as err
;         Msgbox(err.Message)
; }

; ; 切换到英文输入法，并打开大写锁定
; english_upper_input() {	
;     try {		
; 		global superkey := 0
; 		switch_input_layout("en") ; 切换到英文输入法
; 		SetCapsLockState true
; 		Notify.show("英文大写", 0.5, "Purple") 
;     }
;     catch as err
;         Msgbox(err.Message)  
; }


; GetCurLayout(&hWord :="", &lWord :="") {	
;   fgWin	:= DllCall("GetForegroundWindow") ; Get handle (HWND) to the foreground window docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getforegroundwindow
;   if WinActive("ahk_class ConsoleWindowClass") { ; get layout for Console
; 	IMEWnd	:= DllCall(getDefIMEWnd, "Ptr",fgWin) ; DllCall("Imm32.dll\ImmGetDefaultIMEWnd", "Ptr",fgWin)
; 	if (IMEWnd == 0) {
; 	  Return
; 	} else {
; 	  fgWin := IMEWnd
; 	}
;   } else if WinActive("ahk_class vguiPopupWindow") or WinActive("ahk_class ApplicationFrameWindow")  or WinActive("ahk_class Notepad") { ; Steam, some UWP apps, get layout from a keyboard focused control since can"t read it from a regular window autohotkey.com/boards/viewtopic.php?f=76&t=69414
; 	lFocused	:= ControlGetFocus("A")
; 	if (lFocused == 0) {
; 	  Return
; 	} else {
; 	  CtrlID	:= ControlGetHwnd(lFocused, "A")
; 	  fgWin 	:= CtrlID
; 	}
;   } else if WinActive("ahk_class Progman") or WinActive("ahk_class WorkerW") {
; 	lFocused	:= ControlGetFocus("A")
; 	if (lFocused == 0) {
; 	  Return
; 	} else {
; 	  CtrlID	:= ControlGetHwnd(lFocused, "A")
; 	  fgWin 	:= CtrlID
; 	}
;   }
;   threadID     	:= DllCall("GetWindowThreadProcessId"	,  "Ptr",fgWin , "Ptr",0) ; DWORD GetWindowThreadProcessId(HWND hWnd, LPDWORD lpdwProcessId) docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowthreadprocessid
;   inputLocaleID	:= DllCall("GetKeyboardLayout"       	, "UInt",threadID) ; In some examples ends with ", "UInt"" return type, but this isn"t precise enough to catch differences between my custom layouts, so need the full "0xfffffffff0c00409" value
; 	;+---------+-------------+ docs.microsoft.com/en-us/windows/win32/intl/language-identifiers
; 	;|SubLangID|PrimaryLangID|
; 	;+---------+-------------+
; 	;15     10 9             0 bit
;   hWord	:= inputLocaleID >> 16   	; Device handle to the physical layout. Bitwise right shift by 16 bits = 4 hex characters (i.e. size of lWord)
;   lWord	:= inputLocaleID & 0xFFFF	; Language Identifier for the input language
;   Return inputLocaleID
;   }



; ; 定义切换输入法布局的函数
; switch_input_layout(target_layout := "") {	
; 	WM_INPUTLANGCHANGEREQUEST := 0x0050
; 	ENGLISH_INPUT_LAYOUT := "67699721"
; 	CHINESE_INPUT_LAYOUT := "134481924"

;     target_window := "A" 
;     if (target_layout = "en")
;         PostMessage(WM_INPUTLANGCHANGEREQUEST, 0, ENGLISH_INPUT_LAYOUT, , target_window)
;     else if (target_layout = "zh") 
;         PostMessage(WM_INPUTLANGCHANGEREQUEST, 0, CHINESE_INPUT_LAYOUT, , target_window)
;     else
;         Msgbox("输入的输入法布局参数无效，不会进行布局切换")

;     target_window := WinActive("ahk_group Special_ime_window") ? ControlGetFocus("A") : "A" 
;     if (target_layout = "en")
;         PostMessage(WM_INPUTLANGCHANGEREQUEST, 0, ENGLISH_INPUT_LAYOUT, , target_window)
;     else if (target_layout = "zh") 
;         PostMessage(WM_INPUTLANGCHANGEREQUEST, 0, CHINESE_INPUT_LAYOUT, , target_window)
;     else
;         Msgbox("输入的输入法布局参数无效，不会进行布局切换")


; }


; #endregion


; #region other


/**
 * 生成随机字符串
 * @param {Number} string_length - 字符串长度
 * @param {String} char_set - 字符集（默认：字母数字）
 * @return {String} 随机字符串
 */
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

/**
 * 字符串相似度计算（Levenshtein距离）
 * @param {String} string1 - 字符串1
 * @param {String} string2 - 字符串2
 * @return {Float} 相似度（0-1之间）
 */
calculate_string_similarity(string1, string2) {
    try {
        if (string1 = string2) {
            return 1.0
        }
        
        if (string1 = "" || string2 = "") {
            return 0.0
        }
        
        length1 := StrLen(string1)
        length2 := StrLen(string2)
        
        ; 创建距离矩阵
        distance_matrix := []
        
        ; 初始化矩阵
        Loop length1 + 1 {
            row_array := []
            Loop length2 + 1 {
                row_array.Push(0)
            }
            distance_matrix.Push(row_array)
        }
        
        ; 填充第一行和第一列
        Loop length1 + 1 {
            distance_matrix[A_Index][1] := A_Index - 1
        }
        Loop length2 + 1 {
            distance_matrix[1][A_Index] := A_Index - 1
        }
        
        ; 计算编辑距离
        Loop length1 {
            i := A_Index
            Loop length2 {
                j := A_Index
                char1 := SubStr(string1, i, 1)
                char2 := SubStr(string2, j, 1)
                
                if (char1 = char2) {
                    cost := 0
                } else {
                    cost := 1
                }
                
                delete_cost := distance_matrix[i][j + 1] + 1
                insert_cost := distance_matrix[i + 1][j] + 1
                replace_cost := distance_matrix[i][j] + cost
                
                distance_matrix[i + 1][j + 1] := Min(delete_cost, Min(insert_cost, replace_cost))
            }
        }
        
        edit_distance := distance_matrix[length1 + 1][length2 + 1]
        max_length := Max(length1, length2)
        similarity := 1.0 - (edit_distance / max_length)
        
        return similarity
        
    } catch as err {
        Msgbox("相似度计算失败: " . err.message)
        return 0.0
    }
}

/**
 * 时间戳转换为可读格式
 * @param {Number} timestamp - 时间戳（毫秒）
 * @param {String} date_format - 日期格式（默认："yyyy-MM-dd HH:mm:ss"）
 * @return {String} 格式化的日期时间
 */
format_timestamp(timestamp, date_format := "yyyy-MM-dd HH:mm:ss") {
    try {
        ; 转换毫秒为秒
        timestamp_seconds := timestamp // 1000
        
        ; 使用AutoHotkey内置函数格式化
        formatted_time := FormatTime(timestamp_seconds, date_format)
        
        return formatted_time
        
    } catch as err {
        Msgbox("时间格式化失败: " . err.message)
        return ""
    }
}

; 打开回收站
open_recycle_bin() {
	try {
		Run('::{645FF040-5081-101B-9F08-00AA002F954E}')
	}
	catch as err
		Msgbox(err.message)   
}



pixel_to_point(pixel, LOG_PIXELS := 88, POINTS_PER_INCH := 72) {
    hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
    pixels_per_inch := DllCall("GetDeviceCaps", "Ptr", hdc, "Int", LOG_PIXELS, "Int")
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
    return (pixel / pixels_per_inch * POINTS_PER_INCH)
}

shutdown_me() {
    ; 显示一个带有 5 秒超时的消息框
    result := MsgBox("您的电脑将在 5 秒后关机。点击“取消”可中止操作。", "关机警告", "确定取消 超时,5")

    ; 如果用户点击“取消”或者超时未操作，则执行关机
    if (result = "取消" || result = "超时") {
        Shutdown 1
    }
}



; 关机函数
shutdown_computer() {
    Result := MsgBox("5秒后将关机。点击“确定”中止操作。", "确认", "OK Icon! T5")
    if (Result = "OK") {
        return
    }
    Shutdown 1 ; 执行关机
}

; 重启函数
restart_computer() {
    Result := MsgBox("5秒后将重启。点击“确定”中止操作。", "确认", "OK Icon! T5")
    if (Result = "OK") {
        return
    }
    Shutdown 2 ; 执行重启
}

; 锁定屏幕函数
lock_screen() {
    Result := MsgBox("5秒后将锁定屏幕。点击“确定”中止操作。", "确认", "OK Icon! T5")
    if (Result = "OK") {
        return
    }
    DllCall("LockWorkStation") ; 执行锁定屏幕
}

; 打开当前文件所在目录
open_current_path() {
    ; 检查当前活动窗口是否是 Excel、Word 或 PowerPoint
    if (WinActive("ahk_exe EXCEL.EXE")) {
        app_name := "Excel"
        com_object := "Excel.Application"
    } else if (WinActive("ahk_exe WINWORD.EXE")) {
        app_name := "Word"
        com_object := "Word.Application"
    } else if (WinActive("ahk_exe POWERPNT.EXE")) {
        app_name := "PowerPoint"
        com_object := "PowerPoint.Application"
    } else {
        Run("d:\")
        return
    }

    ; 尝试连接到应用程序的 COM 对象
    try {
        app := ComObjActive(com_object)
    } catch {
        MsgBox("无法连接到 " app_name "。请确保已打开文件。")
        return
    }

    ; 获取当前活动文档的完整路径
    if (app_name = "Excel") {
        file_path := app.ActiveWorkbook.FullName
    } else if (app_name = "Word") {
        file_path := app.ActiveDocument.FullName
    } else if (app_name = "PowerPoint") {
        file_path := app.ActivePresentation.FullName
    }

    ; 检查文件路径是否为空
    if (file_path = "") {
        MsgBox("当前没有打开任何文件。")
        return
    }

    ; 在资源管理器中打开文件所在目录并选中文件
    Run("explorer.exe /select, " file_path)
}




min_window_under_mouse() {   
    try {
        MouseGetPos , , &id
        if (WinExist("ahk_group system_group ahk_id " . id))
            return
        if (WinExist("ahk_id " . id))
            WinMinimize
    } catch as err
        msgbox(err.Message)
}





close_window_under_mouse() { 
    try {
        MouseGetPos , , &id
        if (WinExist("ahk_group system_group ahk_id " . id))
            return
        if (WinExist("ahk_id " . id))
            WinClose
    } catch as err
        msgbox(err.Message)
}





ctrlw_close_window() {
	; MsgBox WinGetTitle("A")
	MouseGetPos , , &id
	if (WinExist("ahk_group system_group" . " ahk_id " . id)) {
		return
	}
	
	if (WinExist("ahk_group black_office" . " ahk_id " . id)) {
		WinClose WinExist("ahk_id " . id)
		return
	}

	if (WinExist("ahk_group group_office" . " ahk_id " . id)) {

		blank_win := ["Excel"]

		win_title :=  WinGetTitle(id)


		Sleep(500)
		SendInput("^w")
		return
	}

	WinClose WinExist("ahk_id " . id)
}

close_window() {  
    if !WinActive("ahk_group system_group")
        WinClose("A")	
}

 
min_window() {
	WinMinimize("A")
}


max_window() {	
	if !WinActive("ahk_group system_group")		
        if WinGetMinMax("A")
            WinRestore("A")
        else
            WinMaximize("A")    
}

do_nothing() => a := 1
noop() => a := 1

set_prior_key(key) {
    global prior_key := key
}

set_superkey_true() {
	global superkey := True
}


set_super_power_true() {
	global super_power := True
}

set_super_power_false() {
	global super_power := False
}



set_superkey_false() {
	global superkey := False
}

; set_superkey_false(letter := "") {
;     if letter {
;         send_with_capslock(letter)
;     }
; 	global superkey := False
;     confirm_layout() 
;     show_keyboard_layout()
; }

set_superkey_pure(value := true) {
	global
	s_is_pure := value
	d_is_pure := value
	f_is_pure := value
}

set_superkey3_enabled(value := true) {
	global is_superkey3_enabled
	is_superkey3_enabled := value
}

toggle_superkey_unlocked() {
	global superkey_is_unlocked
	superkey_is_unlocked := not superkey_is_unlocked
	if superkey_is_unlocked	
		Notify.show("Super Key is unlocked.")	
	else
		Notify.show("Super Key is locked.")	
}

reverse_caps_letter(letter) {
	if (RegExMatch(letter, "^[a-z]$")){
		if GetKeyState("CapsLock", "T")
			return StrUpper(letter)
		else
			return letter 
	}	
}

base_on_capslock(key := A_ThisHotkey) {	
	capslock_state := GetKeyState("CapsLock", "T")
	if (capslock_state)
		return StrUpper(key)
	else
		return key 
}

send_against_capslock(key := A_ThisHotkey) {
	key := SubStr(key, -1)
	if (RegExMatch(key, "^[a-z]$"))
		GetKeyState("CapsLock", "T") ? SendInput(key) : SendInput(StrUpper(key))
}

send_with_capslock(key := A_ThisHotkey) {
	key := SubStr(key, -1)
	if (RegExMatch(key, "^[a-z]$")) {
        key_state := GetKeyState("CapsLock", "T") 
        uppercase_key := StrUpper(key)
        lowercase_key := StrLower(key)        
		key_state ? SendInput(uppercase_key) : SendInput(lowercase_key)
    } else {
        SendInput(key)
    }
}

send_char(hotkey := A_ThisHotkey) {
	char := SubStr(hotkey, -1)

	if RegExMatch(char, "^[a-z]$") && GetKeyState("CapsLock", "T")
        SendText(StrUpper(char))
    else
        SendText(char)
}

close_window_in_taskbar() {
	Click "Right"
	Sleep(1000)
	SendInput("{Up}{Space}")
}

set_superkey(value := 1) {
	global superkey := value
}

remove_spaces(str_input) {
    return RegExReplace(str_input, " ", "")
}

mod_keya(combo) {
	combo := RegExReplace(combo, "^[\*\~\$]+")  ; 去掉前缀 *, ~, $
	if RegExMatch(combo, "^(.*?)[ &]*([^\s&]+)$", &match)
		return Trim(match[1])
	return ""
}

main_keya(combo) {
	combo := RegExReplace(combo, "^[\*\~\$]+")  ; 去掉前缀 *, ~, $
	if RegExMatch(combo, "^(.*?)[ &]*([^\s&]+)$", &match)
		return Trim(match[2])
	return Trim(combo)
}

trim_array(arr) {
    for index, value in arr
        arr[index] := Trim(value)
    return arr
}

extract_primary_key(combo) {
    if (InStr(combo, "&"))
        return RegExReplace(combo, ".*&\s*(\w+)$", "$1")
    return RegExReplace(combo, "[\*\~\$\#\+\!\^]")
}

SetCursorToWait() {
	OCR_NORMAL := 32512 ; 箭头光标
	OCR_WAIT := 32514 ; 等待光标
    ; 加载等待光标
    hCursor := DllCall("LoadCursor", "Ptr", 0, "UInt", OCR_WAIT, "Ptr")
    ; 设置系统的默认箭头光标为我们加载的光标
    DllCall("SetSystemCursor", "Ptr", hCursor, "UInt", OCR_NORMAL)
}

RestoreDefaultCursor() {
	SPI_SETCURSORS := 0x0057
    DllCall("SystemParametersInfo", "UInt", SPI_SETCURSORS, "UInt", 1, "Ptr", 0, "UInt", 0)
}


getShortcutTarget(shortcutPath)
{
	wshShell := ComObject("WScript.Shell")
	shortcut := wshShell.CreateShortcut(shortcutPath)
	return(shortcut.TargetPath)
}

isServerReachable(server)
{
    ; Ping the server with 1 echo request
    result := RunWait(A_ComSpec . " /c ping -n 1 " . server . " >nul", "", "Hide")

    ; Check the result of the ping command
    if (result = 0)
        return true  ; Server responded
    else
        return false  ; Server did not respond
}

isWebsiteReachable(url)
{
 ; 使用 curl 命令来尝试访问网页
    ; -s 使 curl 处于静默模式，不输出任何东西
    ; -o 将输出重定向到 nul，我们不需要输出
    ; --head 只请求头部信息
    ; --fail 使得当 HTTP 请求返回 >= 400 的状态码时，curl 命令返回非零值
    script := "curl -s -o nul --head --fail " . url

    ; 运行脚本并等待其完成，不显示命令提示符窗口
    result := RunWait(A_ComSpec . " /c " . script . " >nul", "", "Hide")

    ; 检查运行结果
    if (result = 0)
        return true  ; 网页可访问
    else
        return false  ; 网页不可访问
}





Pos()
{
	try
	{
		MouseGetPos &xpos, &ypos, &id, &control
		ahk_class := WinGetClass(id)
		winTitle :=  WinGetTitle(id)
		ahk_exe := WinGetProcessName(id)
		PosInfo :=
		{
			ahk_exe: ahk_exe,
			xpos: xpos,
			ypos: ypos,
			ahk_id: id,
			ahk_class: ahk_class,
			winTitle: winTitle,
			control: control
		}
		return PosInfo
	}

}

WinGet() {
	if id := WinGetID("A") {
		ahk_class := WinGetClass(id)
		winTitle :=  WinGetTitle(id)
		ahk_exe := WinGetProcessName(id)
		winGetInfo := {
			ahk_exe: ahk_exe,
			ahk_id: id,
			ahk_class: ahk_class,
			winTitle: winTitle
		}
		return WingetInfo
	} else {
		MsgBox "No active windoew found"
		return
	}
}

Focused() {
	id := WinGetID("A")
	FocusedHwnd := ControlGetFocus("A")
	try FocusedClassNN := ControlGetClassNN(FocusedHwnd)
	try ahk_class := WinGetClass(FocusedHwnd)
	try winTitle :=  WinGetTitle(FocusedHwnd)
	try ahk_exe := WinGetProcessName(FocusedHwnd)
	try FocusedInfo := {
			ahk_exe: ahk_exe,
			ahk_id: FocusedHwnd,
			ahk_class: ahk_class,
			winTitle: winTitle,
			control: FocusedClassNN
		}
	return FocusedInfo
}





Morseold(key := "", timeout := 200) {
	/***************************************************
	当用户按下指定热键并松开时，函数会返回 "0"。
	当用户连续按下指定热键两次并松开时，函数会返回 "00"。
	当用户连续按下指定热键三次并松开时，函数会返回 "000"。
	当用户长按指定热键并松开时，函数会返回  "1"。
	****************************************************/
	tout := timeout / 1000
	if (key = "")
		if (InStr(A_ThisHotKey, "&"))
			key := RegExReplace(A_ThisHotKey, ".*&\s*(\w+)$", "$1")
		else
			key := RegExReplace(A_ThisHotKey, "[\*\~\$\#\+\!\^]")
	pattern := ""
	loop {
		t := A_TickCount
		KeyWait(key)
		pattern .= A_TickCount - t > timeout
		ErrorL2 := KeyWait(key , "D T" . tout)
		if (!ErrorL2)
			return pattern
	}
}


BackupClipboardText() {
    ; 获取当前日期和时间
    currentTime := A_Now
    formattedTime := FormatTime(,"MM-dd_HH-mm-ss")

    ; 设置文件名和路径
    filePath := "D:\Backups\Text Backups\txtBackup_" . formattedTime . ".txt"

    ; 检查文件是否存在
    if (FileExist(filePath)) {
        ; 如果文件存在，则将内容追加到文件末尾
        FileAppend("\r\n" . A_Clipboard, filePath) ; \r\n为换行符，确保新的内容从新一行开始
    } else {
        ; 如果文件不存在，则创建新文件并写入内容
        FileAppend(A_Clipboard, filePath)
    }

    ; 可选：清空剪贴板
    A_Clipboard := ""
}






isObjSelected() {
	clipSaved := ClipboardAll()
	A_Clipboard := ""
	Send "^c"
	ClipWait(0.2)
	; if !ClipWait(0.2)
	; {
	; 	MsgBox "The attempt to copy text onto the clipboard failed.", , 1
	; 	return
	; }
	; MsgBox(A_Clipboard ? "Some one is selected" : "No one is selected.")
	A_Clipboard := clipSaved
	clipSaved := ""
}


MouseIsOver(WinTitle) {
	MouseGetPos , , &id
	return WinExist(WinTitle . " ahk_id " . id)
}

mouse_is_over(WinTitle) {
	MouseGetPos , , &id
	return WinExist(WinTitle . " ahk_id " . id)
}

mouse_is_over_taskbar() {
	MouseGetPos , , &id    
	return (WinExist("ahk_class Shell_TrayWnd ahk_id " . id) or WinExist("ahk_class Shell_SecondaryTrayWnd ahk_id " . id))
}

mouse_is_top() {
	MouseGetPos &xpos, &ypos  
	return (ypos < 15)
}


mouse_is_right() {
	MouseGetPos &xpos, &ypos  
	return (xpos > 3800)
}

mouse_is_left() {
	MouseGetPos &xpos, &ypos  
	return (xpos < 15)
}


; close_ditto() {
; 	ctrl_backtick()
; 	SendInput("{Esc}")
; }



; ===================================================================
; 测试区域
; ===================================================================
test_generate_random_text() {



; ===================================================================
; 测试区域
; ===================================================================
try {
    ; 生成 50 字
    text_50 := generate_random_paragraph(20)
    MsgBox("50 字示例:`n" . text_50 . "`n`n实际长度: " . StrLen(text_50))

    ; 生成 100 字
    text_100 := generate_random_paragraph(20)
    MsgBox("100 字示例:`n" . text_100 . "`n`n实际长度: " . StrLen(text_100))

    ; 生成 20 字 (短文本测试)
    text_20 := generate_random_paragraph(20)
    MsgBox("20 字示例:`n" . text_20 . "`n`n实际长度: " . StrLen(text_20))

} catch as e {
    MsgBox("发生错误: " . e.message)
}
}

; ===================================================================
; 函数定义
; ===================================================================
/**
 * 生成随机中文文本
 * 
 * @param length          整数 - 生成文本的目标长度
 * @param include_punctuation 布尔值 - 是否包含标点符号
 * @return                字符串 - 生成的随机文本
 */
generate_random_text(length := 8, include_punctuation := false) {
    ; 高频常用汉字库
    static common_chars := 
    "的一是在不了有和人这中大为上个国我以要他时来用们生到作地于出就分对成会可主发年动同工也能下过子说" . 
    "产种面而方后多定行学法所民得经十三之进着等部度家电力里如水化高自二理起小物现实量都两体制机当使点" . 
    "从业本去把性好应开它合还因由其些然前外天政四日那社义事平形相全表间样与关各重新线内数正心反你明吗" . 
    "看原又么利比或但质气第向道命此变条只没结解问意建月公无系军很情者最立代想已通并提直题党程展五果您" . 
    "料象员革位入常文总次品式活设及管特件长求老头基资边流路级少图山统接知较将组见计别她手角期根论运击" . 
    "农指几九区强放决西被干做必战先回则任取完据处南目确界领表志入世计记特率专历究拉声步类古克敌胡司千" . 
    "私交烟称构光维即百达精列死坚师听划感示活改严什话术真至信调引争容究况须持布越织具非必型翻飞采收跟" . 
    "白且许广马今写充半往占卡早范座另诉施创坐兴抓坏松台若苦透终聚迷骨孩双余丰突负刻攻研色减单显候段划"
    
    ; 标点符号库
    static all_punctuations := "，。"
    ; 专门用于结尾的标点 (句号)
    static end_punctuations := "。"
    
    char_len := StrLen(common_chars)
    all_punct_len := StrLen(all_punctuations)
    end_punct_len := StrLen(end_punctuations)
    
    result_array := []
    current_len := 0
    
    ; 标点插入间隔 (每 10-14 个字尝试插入)
    punct_interval := 12
    
    while (current_len < length) {
        insert_punct := false
        
        ; 只有在开启标点模式，且不是最后预留位置时才随机插入中间标点
        ; 注意：我们故意留出最后一步来处理结尾，所以循环条件里可以稍微宽松，
        ; 或者在循环内判断如果是最后一个位置，强制走结尾逻辑。
        
        if (include_punctuation && current_len > 3 && current_len < length - 1) {
            if (mod(current_len, punct_interval) = 0 || random(1, 15) = 1) {
                insert_punct := true
            }
        }
        
        if (insert_punct) {
            p_index := random(1, all_punct_len)
            result_array.push(SubStr(all_punctuations, p_index, 1))
            current_len++
        } else {
            c_index := random(1, char_len)
            result_array.push(SubStr(common_chars, c_index, 1))
            current_len++
        }
    }
    
    if (include_punctuation) {
        last_char := result_array.pop()
        
        ; 情况 A: 如果最后一个字符已经是结束标点 (。)，完美，放回去
        if (InStr(end_punctuations, last_char)) {
            result_array.push(last_char)
        } 
        ; 情况 B: 如果最后一个字符是汉字，或者是中间标点 (，、；等)
        ; 策略：替换为一个随机的结束标点 (。)
        else {
            e_index := random(1, end_punct_len)
            perfect_end := SubStr(end_punctuations, e_index, 1)
            result_array.push(perfect_end)
            ; 注意：这里我们是 replace (pop 后 push)，所以总长度保持不变
        }
    }
    ; 如果不带标点，保持纯汉字结尾即可
    
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




; ===================================================================
; 函数定义
; ===================================================================
/**
 * 生成随机商务/公文风格段落
 * 强制包含标点，确保以句号结尾，无连续标点。
 * 
 * @param target_length 整数 - 目标文本长度
 * @return              字符串 - 生成的段落
 */
generate_random_paragraph(target_length := 50) {
    ; --- 词库定义 (Static 优化) ---
    
    ; 单字词语（有意义）
    static one_char_words := [
        "是", "对", "不", "到", "已", "可", "有", "在", "从", "向",
        "以", "为", "和", "与", "上", "下", "前", "后", "左", "右",
        "能", "应", "将", "要", "需", "因", "由", "此", "但", "且"
    ]

    ; 双字词语（商务场景）
    static two_char_words := [
        "战略", "市场", "产品", "管理", "运营", "发展", "创新", "改革",
        "增长", "提升", "优化", "调整", "布局", "拓展", "深化", "推进",
        "加强", "完善", "资源", "客户", "品牌", "效率", "风险", "平台",
        "数据", "服务", "融资", "核心", "竞争", "规划", "组织", "反馈"
    ]

    ; 语气词 / 连接词
    static tone_words := [
        "进一步", "持续", "稳步", "不断", "积极", "深入",
        "全面", "有效", "加快", "强化", "推动", "促进"
    ]

    ; --- 变量初始化 ---
    result_array := []
    current_len := 0
    
    ; 逗号间隔逻辑 (每 4-8 个短语插入一个逗号)
    next_comma_count := random(4, 8)
    phrase_count := 0
    
    one_count := one_char_words.length
    two_count := two_char_words.length
    tone_count := tone_words.length

    while (current_len < target_length) {
        ; 1. 决定本次生成的短语内容 (语气词 + 核心词)
        ; 尝试生成一个短语，先预估长度
        
        ; 随机决定是否加语气词 (30% 概率加)
        use_tone := (random(1, 10) <= 3)
        
        temp_tone := ""
        temp_core := ""
        temp_core_len := 0
        
        if (use_tone && tone_count > 0) {
            temp_tone := tone_words[random(1, tone_count)]
        }
        
        ; 随机选择 1 字或 2 字核心词
        ; 优先选 2 字，显得更正式，除非空间不够
        prefer_two := (random(1, 3) != 1) 
        
        if (prefer_two && two_count > 0) {
            temp_core := two_char_words[random(1, two_count)]
            temp_core_len := 2
        } else if (one_count > 0) {
            temp_core := one_char_words[random(1, one_count)]
            temp_core_len := 1
        } else if (two_count > 0) {
            ; 兜底：如果没有单字词了，只能用双字
            temp_core := two_char_words[random(1, two_count)]
            temp_core_len := 2
        }
        
        temp_phrase := temp_tone . temp_core
        temp_phrase_len := StrLen(temp_phrase)
        
        ; 2. 长度预检查
        ; 我们需要预留空间给可能的逗号和最后的句号
        ; 最小预留：1 (句号)
        ; 如果即将达到间隔，还需预留：1 (逗号)
        
        will_insert_comma := (phrase_count + 1 >= next_comma_count)
        required_space := temp_phrase_len + 1 ; +1 是最后的句号
        if (will_insert_comma) {
            required_space += 1 ; +1 是逗号
        }
        
        if (current_len + required_space > target_length) {
            ; 空间不足以放下这个短语 + 标点
            ; 尝试缩小：如果用了语气词，去掉语气词试试
            if (use_tone && temp_core_len > 0) {
                temp_phrase := temp_core
                temp_phrase_len := temp_core_len
                required_space := temp_phrase_len + 1
                if (will_insert_comma) {
                    required_space += 1
                }
            }
            
            ; 如果还是不够，或者连单字都放不下，就停止生成词语，准备收尾
            if (current_len + required_space > target_length) {
                break
            }
        }
        
        ; 3. 确认添加短语
        result_array.push(temp_phrase)
        current_len += temp_phrase_len
        phrase_count += 1
        
        ; 4. 判断是否插入逗号
        if (will_insert_comma) {
            ; 再次确认插入逗号后不会超过总长度 (留 1 位给最终句号)
            if (current_len + 1 < target_length) {
                result_array.push("，")
                current_len += 1
                phrase_count := 0 ; 重置计数
                next_comma_count := random(4, 8) ; 重置间隔
            }
        }
    }
    
    ; =================================================================
    ; 完美结尾处理
    ; =================================================================
    ; 情况 A: 最后一个字符是逗号 -> 替换为句号
    ; 情况 B: 最后一个字符是汉字 -> 追加句号 (如果长度允许) 或 替换最后一个字为句号(如果长度刚好)
    
    last_item := result_array.length > 0 ? result_array.pop() : ""
    
    if (last_item = "，") {
        ; 如果是逗号，直接换成句号
        result_array.push("。")
        ; 长度不变
    } else {
        ; 如果是汉字
        ; 检查是否还有空间加句号
        if (current_len < target_length) {
            result_array.push(last_item)
            result_array.push("。")
            current_len += 1
        } else {
            ; 长度刚好满了，没有空间加句号了
            ; 策略：把最后一个汉字改成句号 (虽然损失一个字，但保证了标点结尾)
            ; 或者：如果最后一个词是双字词，砍掉一个字变单字，腾出空间？
            ; 为了简单且保证“含标点”，这里选择：如果满了，就把最后一个字符替换为句号
            ; 但为了语意通顺，最好是确保循环里预留了空间。
            ; 上面的循环逻辑 `required_space` 已经尽量预留了。
            ; 如果万一还是满了，我们强制替换最后一位为句号。
            result_array.push("。")
            ; 注意：这里实际上总长度没变，只是最后一个字变成了句号
        }
    }
    
    return join_array(result_array)
}

/**
 * 辅助函数：连接数组
 */
join_array(arr) {
    result := ""
    for item in arr {
        result .= item
    }
    return result
}



generate_random_paragraph2(length := 50) {
	; 单字词语（有意义）
	one_char_words := [
		"是", "对", "不", "到", "已", "可", "有", "在", "从", "向",
		"以", "为", "和", "与", "上", "下", "前", "后", "左", "右",
		"能", "应", "将", "要", "需", "因", "由", "此", "但", "且"
	]

	; 双字词语（商务场景）
	two_char_words := [
		"战略", "市场", "产品", "管理", "运营", "发展", "创新", "改革",
		"增长", "提升", "优化", "调整", "布局", "拓展", "深化", "推进",
		"加强", "完善", "资源", "客户", "品牌", "效率", "风险", "平台",
		"数据", "服务", "融资", "核心", "竞争", "规划", "组织", "反馈"
	]

	; 语气词 / 连接词
	tone_words := [
		"进一步", "持续", "稳步", "不断", "积极", "深入",
		"全面", "有效", "加快", "强化", "推动", "促进"
	]

    local para := ""
    local current_length := 0
    local comma_interval := Random(4, 8)
    local phrase_count := 0 ; 已生成短语数量，用于判断逗号

    while (current_length < length) {
        ; 随机决定使用 1 字或 2 字词
        word_type := Random(1, 2)

        if (word_type = 1 && one_char_words.Length > 0) {
            ; 使用 1 字词
            word_index := Random(1, one_char_words.Length)
            word := one_char_words[word_index]
            word_length := 1
        } else if (two_char_words.Length > 0) {
            ; 使用 2 字词
            word_index := Random(1, two_char_words.Length)
            word := two_char_words[word_index]
            word_length := 2
        }

        ; 如果加上这个词语会超过目标长度，则只添加单字词语
        if ((current_length + word_length) > length) {
            ; 只能加单字词
            if (one_char_words.Length > 0) {
                word_index := Random(1, one_char_words.Length)
                word := one_char_words[word_index]
                word_length := 1
            } else {
                break ; 没有更多的单字词可用，退出循环
            }
        }

        ; 有一定概率添加语气词
        use_tone := Random(0, 3) ; 0~3 中只有 0 不加语气词
        tone := use_tone ? tone_words[Random(1, tone_words.Length)] : ""

        ; 合并成完整短语
        phrase := tone . word
        para .= phrase
        current_length += StrLen(phrase)

        phrase_count += 1

        ; 插入逗号（如果不是最后一个并且未超过目标长度）
        if (phrase_count >= comma_interval && Mod(phrase_count, comma_interval) = 0 && current_length < length) {
            para .= "，"
            current_length += 1 ; 记录逗号占用的字符数
            comma_interval := Random(4, 8) ; 下一次间隔重新随机
        }
    }

    ; 替换结尾的逗号为句号
    if (SubStr(para, 0) == "，")
        para := SubStr(para, 1, StrLen(para) - 1)

    para .= "。"

    return para
}





get_var_by_name(var_name) {
	if IsSet(%var_name%) {
		return %var_name%
	} else {
		return 0
	}
}





Morse(key := "", timeout := 200) {
	/***************************************************6
	当用户按下指定热键并松开时，函数会返回 "0"。
	当用户连续按下指定热键两次并松开时，函数会返回 "00"。
	当用户连续按下指定热键三次并松开时，函数会返回 "000"。
	当用户长按指定热键并松开时，函数会返回  "1"。
	****************************************************/
	tout := timeout / 1000
	if (key = "")
		Key := RegExReplace(A_ThisHotKey, "[\*\~\$\#\+\!\^]")
	pattern := 0
	loop {
		t := A_TickCount
		KeyWait(key)
		pattern += 1
		ErrorL2 := KeyWait(Key , "D T" . tout)
		if (!ErrorL2)
			return pattern
	}
}



	
keyfunc_rbutton_superkeyoff() {
	SendInput("{RButton}")
	superkey := False
}


keyfunc_space_enter() {
	if WinActive("ahk_exe EXCEL.EXE") {
		SendInput("!{Enter}")
		return
	}

	if WinActive("ahk_exe chrome.exe") {
		SendInput("+{Enter}")
		return
	}
	enter_wherever()
}




keyfunc_ctrl2() {
    global superkey
    KeyWait("Ctrl")
    if (A_TimeSinceThisHotkey > 300 || A_PriorKey != "LControl") {
        return
    } 
	SendInput("^{Space}")
}








; #endregion



; #region folder

paste_files_inbox(dest_dir := "D:\system\inbox") {
    if WinActive("ahk_exe Weixin.exe") {
        Sleep(500)
        Click("Right")
        Sleep(500)
        SendInput("{Down}{Enter}")
        Sleep(500)
    }
    paste_files_from_clipboard()
    Run(dest_dir)
}




paste_files_from_clipboard(dest_dir := "D:\system\inbox") {
    if !DirExist(dest_dir)
        DirCreate(dest_dir)

    if !DllCall("user32\OpenClipboard", "Ptr", 0) {
        MsgBox("无法打开剪贴板。")
        return false
    }

    ; CF_HDROP = 15
    if !DllCall("user32\IsClipboardFormatAvailable", "UInt", 15) {
        DllCall("user32\CloseClipboard")
        MsgBox("剪贴板中没有文件。")
        return false
    }

    h_drop := DllCall("user32\GetClipboardData", "UInt", 15, "Ptr")
    if (!h_drop) {
        DllCall("user32\CloseClipboard")
        return false
    }

    p_drop := DllCall("kernel32\GlobalLock", "Ptr", h_drop, "Ptr")
    if (!p_drop) {
        DllCall("user32\CloseClipboard")
        return false
    }

    ; ✅ 关键修正：指定 shell32.dll
    num_files := DllCall("shell32\DragQueryFileW", "Ptr", p_drop, "UInt", 0xFFFFFFFF, "Ptr", 0, "UInt", 0)

    Loop num_files {
        len := DllCall("shell32\DragQueryFileW", "Ptr", p_drop, "UInt", A_Index - 1, "Ptr", 0, "UInt", 0)
        buf := Buffer((len + 1) * 2, 0)
        DllCall("shell32\DragQueryFileW", "Ptr", p_drop, "UInt", A_Index - 1, "Ptr", buf.Ptr, "UInt", len + 1)
        
        full_path := StrGet(buf.Ptr, "UTF-16")
        file_name := StrSplit(full_path, "\").Pop()
        dest_path := dest_dir "\" file_name

        FileCopy(full_path, dest_path, 1)
    }

    DllCall("kernel32\GlobalUnlock", "Ptr", h_drop)
    DllCall("user32\CloseClipboard")

    Notify.show("已粘贴 " num_files " 个文件到：" dest_dir)
    return true
}


get_selected_files(mode:="", hwnd:="") {
    ; 初始化变量
    Toreturn := ""
    filenum1 := 0
    filenum2 := 0

    ; 获取窗口句柄和进程名称
    Process := WinGetProcessName("A")
    lClass := WinGetClass("A")
    hwnd := WinExist("A")

    ; 检查进程是否为 explorer.exe
    if (Process = "explorer.exe") {
        ; 检查窗口类名是否为桌面 (Progman|WorkerW)
        if (lClass ~= "Progman|WorkerW")	{
            ; 获取选中的文件列表
            Files := ListViewGetContent("Selected col1", "SysListView321", "A")

            ; 如果没有选中的文件，则返回桌面路径
            if (Files = "")
                Toreturn .= A_Desktop
            else {
                ; 遍历选中的文件，拼接文件路径
                filenum1++
                loop Parse, Files, "`n", "`r"
                    Toreturn .= A_Desktop "\" A_LoopField "`n"
            }
        }
        else if (lClass ~= "(Cabinet|Explore)WClass") {       ; 遍历当前资源管理器中打开的窗口
            for window in ComObject("Shell.Application").Windows {
                ; 在多个窗口中定位符合前面hwnd的窗口
                if (window.hwnd == hwnd) {
                    pp := window.Document.folder.self.path
                    sel := window.Document.SelectedItems
                    ; 遍历选中的项目，拼接路径
                    for item in sel	{
                        filenum2++
                        Toreturn .= item.path "`r`n"
                    }
                    ; 如果没有选中的项目，则返回当前目录路径
                    if Toreturn = ""
                        Toreturn := pp
                }
            }
        }
    }

    ; 处理返回的路径信息
    fde := Trim(Toreturn, "`r`n")
    if (mode != "") {          ; mode 为 012 时
        if (filenum1 + filenum2 = 0) {
            if (mode = 0) || (mode = 2)
                return
            else            ; 当 mode = 1 时
                return fde
        } else {
            if (mode = 1) or (mode = 2)
                if (filenum1 != 0)	{
                    aa := GetSelectedFiles()
                    return aa ; 返回选定的文件
                } else
                    return fde
        }
    }

    ; 如果路径是目录，则添加反斜杠
    if InStr(FileExist(fde), "D")
        return RegExReplace(Trim(Toreturn, "`r`n") . "\", "\\\\", "\")
    else if Toreturn != "" {
        ; 如果不是目录，则截取目录部分并返回
        lPos := InStr(Toreturn, "\", , -1) - 1
        Toreturn2 := substr(Toreturn, 1, lPos)
        return RegExReplace(Toreturn2 . "\", "\\\\", "\")
    }
}

; 获取选定文件的函数
GetSelectedFiles() {
    originalClipboard := ClipboardAll
    Clipboard := ""
    Send "^c"
    ClipWait 0.5
    clipboardContent := Clipboard

    if (StrSplit(clipboardContent, "`r").MaxIndex() = 1) {
        Clipboard := originalClipboard
        return RegExReplace(clipboardContent, "`r`n", "")
    } else {
        Clipboard := originalClipboard
        return clipboardContent
    }
}
    

;=============================================
; 批量重命名「文件」或「文件夹」
; 命名规则：240101_旁XXX-... → 新日期_旁XXX-...
; 作者：根据用户需求定制
;=============================================

; 全局映射表（编号 → 新日期前缀）
; global_rule_map := Map(
;     "294", "241223",
;     "292", "241209",
;     "291", "241202",
;     "289", "241111",
;     "288", "241104",
;     "285", "240928",
;     "282", "240825",
;     "280", "240512",
;     "279", "240729",
;     "278", "240721",
;     "276", "240701",
;     "275", "240624",
;     "272", "240527",
;     "270", "240429",
;     "269", "240422",
;     "268", "240414",
;     "267", "240407",
;     "265", "240317",
;     "262", "240205",
;     "261", "240129",
;     "260", "240122",
;     "259", "240115"
; )

; global_rule_map := Map(
;     "71", "241125",
;     "70", "241021",
;     "68", "240805",
;     "65", "240512",
;     "64", "240224"
; )

; global_rule_map := Map(
;     ; 段 A: 258 → 244 (2023-12-09 → 2023-08-01)
;     "258", "231209",
;     "257", "231202",
;     "256", "231125",
;     "255", "231111",
;     "254", "231104",
;     "253", "231021",
;     "252", "231014",
;     "251", "230930",
;     "250", "230923",
;     "249", "230909",
;     "248", "230902",
;     "247", "230826",
;     "246", "230819",
;     "245", "230812",
;     "244", "230801",

;     ; 段 B: 244 → 237 (2023-08-01 → 2023-06-01)
;     "243", "230725",
;     "242", "230718",
;     "241", "230704",
;     "240", "230627",
;     "239", "230613",
;     "238", "230606",
;     "237", "230601",

;     ; 段 C: 237 → 224 (2023-06-01 → 2023-01-01)
;     "236", "230525",
;     "235", "230518",
;     "234", "230504",
;     "233", "230427",
;     "232", "230413",
;     "231", "230406",
;     "230", "230323",
;     "229", "230316",
;     "228", "230309",
;     "227", "230223",
;     "226", "230216",
;     "225", "230109",
;     "224", "230101"
; )

; global_rule_map := Map(
;     ; 222 → 217 (2022-12-13 → 2022-11-01)
;     "222", "221213",
;     "221", "221206",
;     "220", "221122",  ; 跳过 11-29
;     "219", "221115",
;     "218", "221108",
;     "217", "221101",

;     ; 217 → 200 (2022-11-01 → 2022-05-17) —— 跳过多个月末
;     "216", "221018",  ; 跳 10-25
;     "215", "221011",
;     "214", "221004",
;     "213", "220920",  ; 跳 9-27
;     "212", "220913",
;     "211", "220906",
;     "210", "220823",  ; 跳 8-30
;     "209", "220816",
;     "208", "220809",
;     "207", "220802",
;     "206", "220719",  ; 跳 7-26
;     "205", "220712",
;     "204", "220705",
;     "203", "220621",  ; 跳 6-28
;     "202", "220614",
;     "201", "220607",
;     "200", "220517",  ; 锚点（5-17，未到月末）

;     ; 200 → 191 (2022-05-17 → 2022-02-14)
;     "199", "220510",
;     "198", "220503",
;     "197", "220419",  ; 跳 4-26
;     "196", "220412",
;     "195", "220405",
;     "194", "220322",  ; 跳 3-29
;     "193", "220315",
;     "192", "220308",
;     "191", "220214",  ; 锚点（2-14，2月无25+，但可能春节调整）

;     ; 191 → 187 (2022-02-14 → 2022-01-04)
;     "190", "220207",
;     "189", "220125",  ; ❌ 1月25日在禁区！→ 应跳
;     ; 修正：1月25日不能发 → 改为 1月18日
;     "189", "220118",
;     "188", "220111",
;     "187", "220104"
; )

; global_rule_map := Map(
;     "63", "231213",  ; 12月特例（中旬）
    
;     ; 62 → 53: 每月下旬（25～31日）
;     "62", "231029",
;     "61", "230926",
;     "60", "230829",
;     "59", "230725",
;     "58", "230627",
;     "57", "230530",
;     "56", "230425",
;     "55", "230328",
;     "54", "230228",
;     "53", "230131"
; )

global_rule_map := Map(
    "52", "221228",
    "51", "221127",
    "50", "221030",
    "49", "220926",
    "48", "220829",
    "47", "220725",
    "46", "220628",
    "45", "220531",
    "44", "220427",
    "43", "220326"
)





;──────────────────────────────
; 重命名剪贴板中的【文件夹】
;──────────────────────────────
rename_folder_from_clipboard() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空，请先复制文件夹路径！")
        return
    }

    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0

    for line in lines {
        path := Trim(line, "`"`t`r`n")  ; 去除引号和空白
        if (path = "")
            continue

        ; 确保是存在的目录
        if !FileExist(path) || !InStr(FileExist(path), "D")
            continue

        ; 移除末尾反斜杠（便于解析）
        if SubStr(path, -1) = "\"
            path := SubStr(path, 1, -1)

        folder_name := StrSplit(path, "\").Pop()
        parent_dir := SubStr(path, 1, -StrLen(folder_name)) . "\"

        ; 匹配命名规则
        if RegExMatch(folder_name, "^220101_甲(\d{2})-(.*)$", &m) {
            num := m.1
            rest := m.2
            if global_rule_map.Has(num) {
                new_name := global_rule_map[num] "_甲" num "-" rest
                new_full_path := parent_dir . new_name

                if FileExist(new_full_path) {
                    ; 可选：记录跳过项（此处静默跳过或提示）
                    continue
                }

                try {
                    DirMove(path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                    ; 可选：记录错误 err.message
                }
            }
        }
    }

    MsgBox("📁 文件夹重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}

rename_folder_from_clipboard_pang() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空，请先复制文件夹路径！")
        return
    }

    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0

    for line in lines {
        path := Trim(line, "`"`t`r`n")  ; 去除引号和空白
        if (path = "")
            continue

        ; 确保是存在的目录
        if !FileExist(path) || !InStr(FileExist(path), "D")
            continue

        ; 移除末尾反斜杠（便于解析）
        if SubStr(path, -1) = "\"
            path := SubStr(path, 1, -1)

        folder_name := StrSplit(path, "\").Pop()
        parent_dir := SubStr(path, 1, -StrLen(folder_name)) . "\"

        ; 匹配命名规则
        if RegExMatch(folder_name, "^220101_旁(\d{3})-(.*)$", &m) {
            num := m.1
            rest := m.2
            if global_rule_map.Has(num) {
                new_name := global_rule_map[num] "_旁" num "-" rest
                new_full_path := parent_dir . new_name

                if FileExist(new_full_path) {
                    ; 可选：记录跳过项（此处静默跳过或提示）
                    continue
                }

                try {
                    DirMove(path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                    ; 可选：记录错误 err.message
                }
            }
        }
    }

    MsgBox("📁 文件夹重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}


;──────────────────────────────
; 重命名剪贴板中的【文件】
;──────────────────────────────
rename_files_from_clipboard() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空！")
        return
    }

    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0

    for line in lines {
        ; 清理路径（去除引号和空白）
        clean_line :=  Trim(line, "`"`t`r`n")  ; 去除引号和空白
        if (clean_line = "")
            continue

        ; 必须是存在的【文件】
        if !FileExist(clean_line) || InStr(FileExist(clean_line), "D")
            continue

        ; 使用 path_info 解析路径
        pi := path_info(clean_line)
        full_path   := pi.Full      ; 绝对路径（规范化）
        file_name   := pi.Fname     ; 不含扩展名
        file_ext    := pi.Ext       ; 含点的扩展名（可能为空）
        folder_path := pi.Folder    ; 驱动器+目录，结尾无反斜杠

        ; 拼回完整文件名用于匹配（因为原始名可能有大小写/多余空格）
        original_base_name := file_name . file_ext

        ; 匹配规则：240101_旁三位数字-...
        if RegExMatch(original_base_name, "^220101_甲(\d{2})-(.*)$", &m) {
            num  := m.1
            rest := m.2  ; 包含标题和扩展名（如 "报告.pptx"）

            if global_rule_map.Has(num) {
                new_prefix := global_rule_map[num]
                new_file_name := new_prefix "_甲" num "-" rest
                new_full_path := folder_path . "\" . new_file_name

                if FileExist(new_full_path) {
                    continue  ; 跳过已存在
                }

                try {
                    FileMove(full_path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                }
            }
        }
    }

    MsgBox("📄 文件重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}

rename_files_from_clipboard_pang() {
    clip_text := A_Clipboard
    if (clip_text = "") {
        MsgBox("❌ 剪贴板为空！")
        return
    }

    lines := StrSplit(clip_text, ["`r`n", "`n", "`r"])
    renamed_count := 0
    error_count := 0

    for line in lines {
        ; 清理路径（去除引号和空白）
        clean_line :=  Trim(line, "`"`t`r`n")  ; 去除引号和空白
        if (clean_line = "")
            continue

        ; 必须是存在的【文件】
        if !FileExist(clean_line) || InStr(FileExist(clean_line), "D")
            continue

        ; 使用 path_info 解析路径
        pi := path_info(clean_line)
        full_path   := pi.Full      ; 绝对路径（规范化）
        file_name   := pi.Fname     ; 不含扩展名
        file_ext    := pi.Ext       ; 含点的扩展名（可能为空）
        folder_path := pi.Folder    ; 驱动器+目录，结尾无反斜杠

        ; 拼回完整文件名用于匹配（因为原始名可能有大小写/多余空格）
        original_base_name := file_name . file_ext

        ; 匹配规则：240101_旁三位数字-...
        if RegExMatch(original_base_name, "^220101_旁(\d{3})-(.*)$", &m) {
            num  := m.1
            rest := m.2  ; 包含标题和扩展名（如 "报告.pptx"）

            if global_rule_map.Has(num) {
                new_prefix := global_rule_map[num]
                new_file_name := new_prefix "_旁" num "-" rest
                new_full_path := folder_path . "\" . new_file_name

                if FileExist(new_full_path) {
                    continue  ; 跳过已存在
                }

                try {
                    FileMove(full_path, new_full_path)
                    renamed_count++
                } catch as err {
                    error_count++
                }
            }
        }
    }

    MsgBox("📄 文件重命名完成！`n成功: " . renamed_count . "`n失败: " . error_count)
}



move_files_to_parent() {
    try {
		selected_files := get_selected_files(2)
		if !selected_files {
			Msgbox("There's no folder selected.") 
			return			
		}
		parent_folder := path_info(selected_files).folder ; 获取父文件夹路径
		error_count := MoveFilesAndFolders(selected_files '\*.*', parent_folder, 1)
    	DirDelete selected_files
	

		; if (InStr(selected_files, "`r`n")) {
		; 	Loop Parse, selected_files, "`n", "`r" {
		; 		SplitPath A_LoopField, , &dir
		; 		command := Format('"{1}" x "{2}" -o"{3}"\* -aoa', zip_program, A_LoopField, dir)				
		; 		Run command
		; 	}
		; } else {
		; 	SplitPath selected_files, , &dir
		; 	command := Format('"{1}" x "{2}" -o"{3}"\* -aoa', zip_program, selected_files, dir)	
		; 	Run command
		; }
    }
    catch as err
        Msgbox(err.message)   
}


rename_files_to_2parent() {
    try {	
		selected_files := get_selected_files(2)
        ; Msgbox(selected_files)   
		if !selected_files {
			Msgbox("There's no folder selected.") 
			return			
		}	
		if (InStr(selected_files, "`r`n")) {
			Loop Parse, selected_files, "`n", "`r" {
				parent_folder := path_info(A_LoopField).folder ; 获取父文件夹路径
				parts := StrSplit(A_LoopField, "\")
				lastFolder := parts[-2]  ; 获取最后一个部分
				new_name := parent_folder lastFolder '.pptx' 
				sleep(50)
				; Msgbox(new_name)   
				FileMove A_LoopField, new_name, 1
			}
			return
		} 	
		parent_folder := path_info(selected_files).folder ; 获取父文件夹路径
		parts := StrSplit(selected_files, "\")
		lastFolder := parts[-2]  ; 获取最后一个部分
		new_name := parent_folder lastFolder '.pptx' 
		
		; Msgbox(new_name) 
		FileMove selected_files, new_name, 1
    }
    catch as err
        Msgbox(err.message)   
}



getdate_914(files) {
    try {		
		ppt := ComObject("PowerPoint.Application")
		pptPres := ppt.Presentations.Open(files, false, false, false)
		date := pptPres.BuiltInDocumentProperties("Creation Date").Value 
		;"2022/01/01 12:12:22"转化为"220101"
		; msgbox date
		formatted_date := convert_date_format(date) 
		pptPres.Close() ; 关闭PPT文件
		ppt.Quit() ; 退出PowerPoint
		return(formatted_date)
    }
    catch as err
        Msgbox(err.message) 
}



;"2022/01/01 12:12:22"转化为"220101"
convert_date_format(date) {
    try {  
		
		if RegExMatch(date, "(\d{4})/(\d{1,2})/(\d{1,2})", &match) {
			year := SubStr(match[1], 3, 2)
			month := Format("{:02}", match[2])
			day := Format("{:02}", match[3])	
			formatted_date := year month day
			return formatted_date
		} else {
			MsgBox "日期格式不匹配"			
			return 
		}
	}
    catch as err
        Msgbox(err.message) 
}

  



; ErrorCount := MoveFilesAndFolders("C:\My Folder\*.*", "D:\Folder to receive all files & folders")
; if ErrorCount != 0
;     MsgBox ErrorCount " files/folders could not be moved."

MoveFilesAndFolders(SourcePattern, DestinationFolder, DoOverwrite := false)
; 移动匹配 SourcePattern 的所有文件和文件夹到 DestinationFolder 文件夹中且
; 返回无法移动的文件/文件夹的数目.
{
    ErrorCount := 0
    if DoOverwrite = 1
        DoOverwrite := 2  ; 请参阅 DirMove 了解模式 2 与模式 1 的区别.
    ; 首先移动所有文件(不是文件夹):
    try
        FileMove SourcePattern, DestinationFolder, DoOverwrite
    catch as err
        ErrorCount := Err.Extra
    ; 现在移动所有文件夹:
    Loop Files, SourcePattern, "D"  ; D 表示 "只获取文件夹".
    {
        try
            DirMove A_LoopFilePath, DestinationFolder "\" A_LoopFileName, DoOverwrite
        catch
        {
            ErrorCount += 1
            ; 报告每个出现问题的文件夹名称.
            MsgBox "Could not move " A_LoopFilePath " into " DestinationFolder
        }
    }
    return ErrorCount
}

; 函数：在文件名后附加当前日期和时间，使用中文的冒号和斜杠
append_datetime_to_filename(file_path) {    
    source_full     := path_info(file_path).Full  
    source_fname    := path_info(file_path).Fname
    source_folder   := path_info(file_path).Folder
    source_ext      := path_info(file_path).Ext 
    
    ; 获取当前日期时间，格式为 YYYY/MM/DD HH:mm:ss
    current_datetime := FormatTime(A_Now, "yyyy/MM/dd_HH:mm")
    
    ; 将英文冒号和斜杠替换为中文字符
    safe_datetime := StrReplace(StrReplace(current_datetime, ":", "-"), "/", "-")

    new_fname := source_fname . "【" . safe_datetime . "】"  
    new_full := path_info(source_full, "fname:" . new_fname).Full
    
    ; 返回新的文件路径
    return new_full
}
; ; 示例调用
; test_file_path := "C:\path\to\your\file.txt"
; new_file_path := append_datetime_to_filename(test_file_path)
; MsgBox, % "新文件名为：" . new_file_path


append_datetime() {     
    ; 确保操作在资源管理器窗口中进行
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("解压缩操作只能在文件夹窗口中执行，请打开文件夹再操作！")
        return
    } 
    ; 获取选中的文件路径
    files := get_selected_files(2)         
    if !files {   
        Msgbox("未选择任何文件，无法完成解压缩操作！")
        return  
    }  
    Loop Parse, files, '`n', '`r' {         
        source_full     := path_info(A_LoopField).Full 
        source_ext      := path_info(A_LoopField).Ext 
        new_full        := append_datetime_to_filename(A_LoopField)  

        if source_ext {
            if FileExist(new_full){
                Msgbox("存在同名文件，重命名操作失败！")
                return            
            }   
            try FileMove(source_full, new_full, 0)  
            catch as err
                Msgbox("重命名操作失败:" . err.message) 
        } else {
            if FileExist(new_full){
                Msgbox("存在同名文件夹，重命名操作失败！")
                return            
            }  
            try DirMove source_full, new_full, 'R'
            catch as err
                Msgbox(err.message) 
        }   			
        
    }  
}


output_datetime() {  
     ; 获取当前日期时间，格式为 YYYY/MM/DD HH:mm:ss
    current_datetime := FormatTime(A_Now, "yyyy/MM/dd_HH:mm")
    ; 将英文冒号和斜杠替换为中文字符
    safe_datetime := StrReplace(StrReplace(current_datetime, ":", "-"), "/", "-")
    send_by_clipboard("【" . safe_datetime . "】") 
}


send_by_clipboard(text) {
    old_clipboard := A_Clipboard
    try {
        A_Clipboard := text
        ClipWait(1)
        SendInput("^v")
        Sleep(100)
    } finally {
        ; 恢复原剪贴板内容
        SetTimer(() => (A_Clipboard := old_clipboard), -1000)
    }
}






run_unzip() {  
    ; 确保操作在资源管理器窗口中进行
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("解压缩操作只能在文件夹窗口中执行，请打开文件夹再操作！")
        return
    } 

    ; 定义 7-Zip 的路径
    zip_tool := "C:\Program Files\7-Zip\7z.exe" 

    ; 获取选中的文件路径
    files := get_selected_files(2)         
    if !files {   
        Msgbox("未选择任何文件，无法完成解压缩操作！")
        return  
    }  

    ; 遍历选中的文件
    Loop Parse, files, "`n", "`r" {
        file_full := path_info(A_LoopField).Full  
        file_name := path_info(A_LoopField).Fname
        file_folder  := path_info(A_LoopField).Folder
        file_ext  := path_info(A_LoopField).Ext 

        ; 检查是否为压缩文件
        if !is_compressed(file_full) {
            Msgbox("选中的不是压缩文件，无法完成解压缩操作！")
            return
        }

        ; 构造解压缩命令
        cmd := Format('"{1}" x "{2}" -o"{3}"\* -aoa', zip_tool, file_full, file_folder)	
        try {
            Run(cmd)
        } catch as err {
            Msgbox("解压缩失败：" . err.message)
            return
        }

        Notify.show("已成功解压缩文件！")    
    }    
}

run_zip() {
    ; 确保操作在资源管理器窗口中进行
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("解压缩操作只能在文件夹窗口中执行，请打开文件夹再操作！")
        return
    } 

    ; 定义 7-Zip 的路径
    zip_tool := "C:\Program Files\7-Zip\7z.exe"  

    ; 获取选中的文件路径
    files := get_selected_files(2)
    if !files {
        Msgbox("没有选择文件，无法执行压缩操作！")
        return			
    }        

    ; 提取第一个文件的路径（用于生成压缩文件名）
    first_file := RegExReplace(files, "`r`n.*", "") 
    output_zip := path_info(first_file, "Ext:").Full . ".zip"

    ; 将所有文件路径用双引号包裹并拼接
    file_list := StrReplace(files, "`r`n", '" "')

    ; 构造压缩命令
    cmd := Format('"{1}" a -tzip "{2}" "{3}"', zip_tool, output_zip, file_list)

    ; 执行压缩命令，并捕获可能的异常
    try {
        Run(cmd)
        Notify.show("文件已成功压缩为：" . path_info(output_zip).File)
    } catch as err {
        Msgbox("压缩操作失败：" . err.message)
    }
}

; 检查文件是否是压缩文件
is_compressed(file_name) {
    ext := path_info(file_name).Ext 
    compressed_exts := [".zip", ".rar", ".7z", ".tar", ".gz", ".tgz", ".bz2", 
                        ".xz", ".iso", ".pptx", ".ppt", ".xlsx", ".xls", ".docx", ".doc"]
    for index, extension in compressed_exts {
        if (ext = extension) {
            return 1
        }
    }
    return 0
}


      
; ; 输入字符串
; inputString := "常小二【打卡练习×250105期】运满满APP品牌介绍"
; outputString := "运满满APP品牌介绍【常小二×打卡练习×250105期】"
rename_in_eagle() {
    SendInput("{F2}")
    Sleep(100)
    SendInput("^c")
    Sleep(100)   

    ; 使用正则表达式提取各个部分
    if RegExMatch(A_Clipboard, "(.*)【(.*)】(.*)", &match) {
        ; 提取名字、标签和内容部分
        name := match[1]  ; 常小二
        tag := match[2]   ; 打卡练习×250105期
        content := match[3] ; 运满满APP品牌介绍

        ; 按照目标格式重新组合字符串
        A_Clipboard := content . "【" . name . "×" . tag . "】"
    } else {
        Msgbox("未能匹配输入字符串")
        return
    }
    SendInput("^v")
}

process_powerpoint() {
   get_grandparent_dir(path) {
        ; 使用正则表达式替换，去掉最后两个部分（包括它们之间的反斜杠）
        grandparent := RegExReplace(path, "\\[^\\]+\\[^\\]+$", "")
        ; 然后加上反斜杠
        return grandparent . "\"
    }

    files := get_selected_files(2)
    if (files = "") {
        Msgbox("请先选择要拆分的 PowerPoint 文件！")
        return
    }

    ; 只取第一个文件
    source_full := Trim(StrSplit(files, "`n", "`r")[1])
    if (!FileExist(source_full)) {
        Msgbox("文件不存在：`n" source_full)
        return
    }

    ; 获取文件路径信息
    source_full     := path_info(source_full).full      ; 文件完整路径
    source_drive    := path_info(source_full).drive     ; 驱动器 (如 C:)
    source_dir      := path_info(source_full).dir       ; 目录（不含驱动器）
    source_folder   := path_info(source_full).folder    ; 文件夹（含驱动器）        
    source_fname    := path_info(source_full).fname     ; 文件名（不含后缀）
    source_ext      := path_info(source_full).ext       ; 文件后缀   
    source_file     := path_info(source_full).file      ; 文件名（含后缀）

    source_grandparent_folder := get_grandparent_dir(source_full)

    target_folder := source_grandparent_folder . source_fname . "\"

    try {
        ;     ; 创建目标文件夹（如果不存在）
        ; if (!FileExist(target_folder)) {            
        ;     DirCreate(target_folder)
        ; }   
        DirMove(source_folder, target_folder, "R")
    } catch as err {
        Msgbox("无法移动完整文件：`n" source_full "`nError: " err.Message)
        return
    }


    ; sleep 2000

    source_full := target_folder . source_file

    source_full     := path_info(source_full).full      ; 文件完整路径
    source_drive    := path_info(source_full).drive     ; 驱动器 (如 C:)
    source_dir      := path_info(source_full).dir       ; 目录（不含驱动器）
    source_folder   := path_info(source_full).folder    ; 文件夹（含驱动器）        
    source_fname    := path_info(source_full).fname     ; 文件名（不含后缀）
    source_ext      := path_info(source_full).ext       ; 文件后缀   
    source_file     := path_info(source_full).file      ; 文件名（含后缀）
    source_font_full := source_folder . "字体"
    target_font_full := source_folder . source_fname . "_字体"
    target_wz_fold := source_folder . source_fname . "_完整"
    target_wz_full := source_folder . source_fname . "_完整\" . source_fname . "_完整.pptx"

    split_in_folder(source_full)

    ; sleep 1000 

    export_jpg_in_folder(source_full) 
    
    ; sleep 1000

    try {
         ; 创建目标文件夹（如果不存在）
        if (!FileExist(target_wz_fold)) {            
            DirCreate(target_wz_fold)
        }   
        FileMove(source_full, target_wz_full)
    } catch as err {
        Msgbox("无法移动完整文件：`n" source_full "`nError: " err.Message)
        return
    }
    
    ; sleep 1000 

    if DirExist(source_font_full) {
        try {
                DirMove(source_font_full, target_font_full)
        } catch as err {
            Msgbox("无法移动font文件夹：`n" target_font_full "`nError: " err.Message)
            return
        } 
    }
    
    Notify.show("Successfully!")

}



; ; 将 PPT 按每 N 页分割，输出为 XXX_多页「1-10」.pptx、XXX_多页「11-20」.pptx 等
; split_ppt_by_page_range(source_full, pages_per_file := 10) {
;     if !source_full {
;         files := get_selected_files(2)
;         if (files = "") {
;             Msgbox("请先选择要拆分的 PowerPoint 文件！")
;             return
;         }
;         source_full := Trim(StrSplit(files, "`n", "`r")[1])
;         if (!FileExist(source_full)) {
;             Msgbox("文件不存在：`n" source_full)
;             return
;         }
;     }

;     ; 获取路径信息
;     pi := path_info(source_full)
;     source_fname := pi.fname
;     target_folder := pi.folder . "\" . source_fname . "_多页"

;     if (!FileExist(target_folder))
;         DirCreate(target_folder)

;     ; 启动 PowerPoint（隐藏）
;     ppt_app := ComObject("PowerPoint.Application")

;     ; 打开源文件（只读）
;     try {
;         source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0)
;     } catch as err {
;         Msgbox("无法打开源文件：`n" source_full "`nError: " err.Message)
;         return
;     }

;     total_slides := source_pres.Slides.Count
;     start_page := 1

;     while (start_page <= total_slides) {
;         end_page := Min(start_page + pages_per_file - 1, total_slides)
;         output_name := target_folder . "\" . source_fname . "_多页「" . start_page . "-" . end_page . "」.pptx"

;         ; 复制源文件作为模板
;         FileCopy(source_full, output_name, true)

;         ; 打开副本
;         try {
;             pres := ppt_app.Presentations.Open(output_name, false, false, false)
;         } catch as err {
;             Msgbox("无法打开临时文件：`n" output_name "`nError: " err.Message)
;             start_page += pages_per_file
;             continue
;         }

;         ; 从后往前删除不在 [start_page, end_page] 范围内的幻灯片
;         i := pres.Slides.Count
;         while (i >= 1) {
;             slide_index := pres.Slides(i).SlideIndex
;             if (slide_index < start_page || slide_index > end_page) {
;                 pres.Slides(i).Delete()
;             }
;             i -= 1
;         }

;         ; 清除节信息（可选）
;         try {
;             while pres.SectionProperties.Count
;                 pres.SectionProperties.Delete(pres.SectionProperties.Count, false)
;         } catch {

;         }

;         ; 保存并关闭
;         pres.Save()
;         pres.Close()

;         start_page += pages_per_file
;     }

;     ; 清理
;     try source_pres.Close()
;     try ppt_app.Quit()
;     Sleep 1000
;     if ProcessExist("POWERPNT.EXE")
;         ProcessClose("POWERPNT.EXE")

;     Notify.show("PPT 已成功按每 " pages_per_file " 页分割完成！共 " Ceil(total_slides / pages_per_file) " 个文件。")
; }


; ; 将一个「多页」PPT（如 _多页「81-90」.pptx）拆分为单页文件
; ; 输出到同级目录下的「_单页」文件夹中，命名为 _单页「81」.pptx ～ _单页「90」.pptx
; ; split_batch_ppt_to_single(source_full) {
; ;     if !source_full {
; ;         files := get_selected_files(2)
; ;         if (files = "") {
; ;             Msgbox("请先选择要拆分的 PowerPoint 文件！")
; ;             return
; ;         }
; ;         source_full := Trim(StrSplit(files, "`n", "`r")[1])
; ;         if (!FileExist(source_full)) {
; ;             Msgbox("文件不存在：`n" source_full)
; ;             return
; ;         }
; ;     }

; ;     ; 获取路径信息
; ;     pi := path_info(source_full)
; ;     full_path   := pi.full
; ;     folder      := pi.folder      ; 含驱动器的完整父目录
; ;     fname       := pi.fname       ; 不含后缀的文件名，例如：250609_..._多页「81-90」

; ;     ; 从文件名中提取页码范围，用于生成正确的单页编号
; ;     start_num := "", end_num := ""
; ;     if RegExMatch(fname, "多页「(\d+)-(\d+)」$", &m) {
; ;         start_num := m[1]
; ;         end_num   := m[2]
; ;     } else {
; ;         Msgbox("文件名格式不符合要求！`n应包含：_多页「XX-YY」")
; ;         return
; ;     }

; ;     ; 构建目标文件夹路径：将 "_多页「...」" 替换为 "_单页"
; ;     base_name := RegExReplace(fname, "_多页「\d+-\d+」$", "_单页")
; ;     target_folder := folder . "\" . base_name

; ;     ; 创建目标文件夹（如果不存在）
; ;     if (!FileExist(target_folder))
; ;         DirCreate(target_folder)

; ;     ; 启动 PowerPoint（隐藏）
; ;     ppt_app := ComObject("PowerPoint.Application")

; ;     ; 打开选中的多页文件
; ;     try {
; ;         pres := ppt_app.Presentations.Open(full_path, 1, 0, 0) ; 只读打开
; ;     } catch as err {
; ;         Msgbox("无法打开 PPT 文件：`n" full_path "`nError: " err.Message)
; ;         return
; ;     }

; ;     slide_count := pres.Slides.Count
; ;     expected_count := end_num - start_num + 1
; ;     if (slide_count != expected_count) {
; ;         ; 可选警告（不中断）
; ;         ; MsgBox("警告：文件实际页数 (" slide_count ") 与文件名范围 (" expected_count ") 不一致。", , "Warning")
; ;     }

; ;     ; 使用 Loop 遍历每一页（AHK v2 兼容）
; ;     Loop slide_count {
; ;         current_slide_index := A_Index  ; 当前是第几张幻灯片（1-based）
; ;         actual_page_number := start_num + current_slide_index - 1  ; 对应的真实页码，如 81, 82...

; ;         ; 构建输出文件名
; ;         output_file := target_folder . "\" . base_name . "「" . actual_page_number . "」.pptx"

; ;         ; 复制源文件作为模板
; ;         FileCopy(full_path, output_file, true)

; ;         ; 打开副本
; ;         try {
; ;             single_pres := ppt_app.Presentations.Open(output_file, false, false, false)
; ;         } catch as err {
; ;             Msgbox("无法打开临时文件：`n" output_file "`nError: " err.Message)
; ;             continue
; ;         }

; ;         ; 从后往前删除其他幻灯片（只保留第 current_slide_index 页）
; ;         total := single_pres.Slides.Count
; ;         i := total
; ;         while (i >= 1) {
; ;             if (i != current_slide_index)
; ;                 single_pres.Slides(i).Delete()
; ;             i -= 1
; ;         }

; ;         ; 清除节信息（避免残留）
; ;         try {
; ;             while single_pres.SectionProperties.Count
; ;                 single_pres.SectionProperties.Delete(single_pres.SectionProperties.Count, false)
; ;         } catch {
            
; ;         }

; ;         ; 保存并关闭
; ;         single_pres.Save()
; ;         single_pres.Close()
; ;     }

; ;     ; 清理
; ;     try pres.Close()
; ;     try ppt_app.Quit()
; ;     Sleep 1000
; ;     if ProcessExist("POWERPNT.EXE")
; ;         ProcessClose("POWERPNT.EXE")

; ;     Notify.show("✅ 已成功将「" start_num "-" end_num "」拆分为 " slide_count " 个单页文件！`n输出目录：`n" target_folder)
; ; }

; ; 主函数：先按批次分割，再将每个批次拆分为单页
; split_ppt_full(source_full := "", pages_per_file := 10) {
;     ; 获取路径信息
;     if !source_full {
;         if WinActive("ahk_class CabinetWClass") {
;             files := get_selected_files(2)
;         }
;         if (files = "") {
;             Msgbox("请先选择要拆分的 PowerPoint 文件！")
;             return
;         }
;         source_full := Trim(StrSplit(files, "`n", "`r")[1])
;         if (!FileExist(source_full)) {
;             Msgbox("文件不存在：`n" source_full)
;             return
;         }
;     }
;     pi := path_info(source_full)
;     source_fname := pi.fname
;     source_folder := pi.folder

;     ; 启动 PowerPoint 应用程序（隐藏）
;     ppt_app := ComObject("PowerPoint.Application")

;     ; 打开源文件（只读）
;     try {
;         source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0)
;     } catch as err {
;         Msgbox("无法打开源文件：`n" source_full "`nError: " err.Message)
;         return
;     }

;     total_slides := source_pres.Slides.Count
;     start_page := 1

;     while (start_page <= total_slides) {
;         end_page := Min(start_page + pages_per_file - 1, total_slides)
;         batch_filename := source_folder . "\" . source_fname . "_多页「" . start_page . "-" . end_page . "」.pptx"

;         ; 复制源文件作为新文件的基础
;         FileCopy(source_full, batch_filename, true)

;         ; 打开副本
;         try {
;             batch_pres := ppt_app.Presentations.Open(batch_filename, false, false, false)
;         } catch as err {
;             Msgbox("无法打开临时文件：`n" batch_filename "`nError: " err.Message)
;             start_page += pages_per_file
;             continue
;         }

;         ; 删除不在 [start_page, end_page] 范围内的幻灯片
;         i := batch_pres.Slides.Count
;         while (i >= 1) {
;             slide_index := batch_pres.Slides(i).SlideIndex
;             if (slide_index < start_page || slide_index > end_page) {
;                 batch_pres.Slides(i).Delete()
;             }
;             i -= 1
;         }

;         ; 清除节信息（可选）
;         try {
;             while batch_pres.SectionProperties.Count
;                 batch_pres.SectionProperties.Delete(batch_pres.SectionProperties.Count, false)
;         } catch as err {
;             Msgbox(err.Message)
;         }

;         ; 保存并关闭
;         batch_pres.Save()
;         batch_pres.Close()

;         ; 对此批次文件进行单页拆分
;         split_batch_ppt_to_single(batch_filename)

;         start_page += pages_per_file
;     }

;     ; 提示完成
;     Notify.show("PPT 文件已成功按每 " pages_per_file " 页拆分，并进一步拆分为单页！")

;     ; 关闭源文件和应用程序
;     try source_pres.Close()
;     try ppt_app.Quit()
;     Sleep 1000
;     if ProcessExist("POWERPNT.EXE")
;         ProcessClose("POWERPNT.EXE")
; }

; ; 辅助函数：将「多页」PPT拆分为单页文件
; split_batch_ppt_to_single(batch_full) {
;     ; 获取路径信息
;     pi := path_info(batch_full)
;     full_path   := pi.full
;     folder      := pi.folder      ; 含驱动器的完整父目录
;     fname       := pi.fname       ; 不含后缀的文件名，例如：250609_..._多页「81-90」

;     ; 从文件名中提取页码范围，用于生成正确的单页编号
;     start_num := "", end_num := ""
;     if RegExMatch(fname, "多页「(\d+)-(\d+)」$", &m) {
;         start_num := m[1]
;         end_num   := m[2]
;     } else {
;         Msgbox("文件名格式不符合要求！`n应包含：_多页「XX-YY」")
;         return
;     }

;     ; 构建目标文件夹路径：将 "_多页「...」" 替换为 "_单页"
;     base_name := RegExReplace(fname, "_多页「\d+-\d+」$", "_单页")
;     target_folder := folder . "\" . base_name

;     ; 创建目标文件夹（如果不存在）
;     if (!FileExist(target_folder))
;         DirCreate(target_folder)

;     ; 打开选中的多页文件
;     try {
;         pres := ComObject("PowerPoint.Application").Presentations.Open(full_path, 1, 0, 0) ; 只读打开
;     } catch as err {
;         Msgbox("无法打开 PPT 文件：`n" full_path "`nError: " err.Message)
;         return
;     }

;     slide_count := pres.Slides.Count
;     expected_count := end_num - start_num + 1
;     if (slide_count != expected_count) {
;         MsgBox("警告：文件实际页数 (" slide_count ") 与文件名范围 (" expected_count ") 不一致。", , "Warning")
;     }

;     ; 使用 Loop 遍历每一页（AHK v2 兼容）
;     Loop slide_count {
;         current_slide_index := A_Index  ; 当前是第几张幻灯片（1-based）
;         actual_page_number := start_num + current_slide_index - 1  ; 对应的真实页码，如 81, 82...

;         ; 构建输出文件名
;         output_file := target_folder . "\" . base_name . "「" . actual_page_number . "」.pptx"

;         ; 复制源文件作为模板
;         FileCopy(full_path, output_file, true)

;         ; 打开副本
;         try {
;             single_pres := ComObject("PowerPoint.Application").Presentations.Open(output_file, false, false, false)
;         } catch as err {
;             Msgbox("无法打开临时文件：`n" output_file "`nError: " err.Message)
;             continue
;         }

;         ; 从后往前删除其他幻灯片（只保留第 current_slide_index 页）
;         total := single_pres.Slides.Count
;         i := total
;         while (i >= 1) {
;             if (i != current_slide_index)
;                 single_pres.Slides(i).Delete()
;             i -= 1
;         }

;         ; 清除节信息（避免残留）
;         try {
;             while single_pres.SectionProperties.Count
;                 single_pres.SectionProperties.Delete(single_pres.SectionProperties.Count, false)
;         } catch as err {
;             Msgbox(err.Message)
;         }

;         ; 保存并关闭
;         single_pres.Save()
;         single_pres.Close()
;     }

;     ; 删除临时批次文件
;     FileDelete(batch_full)
; }

; ; ; 示例调用
; ; split_ppt_full("C:\Reports\年度总结_完整.pptx", 10)

; ; 主入口函数：支持单文件路径 或 多选文件（来自资源管理器）
; split_ppt_full_multi(input := "") {
;     ; 获取待处理的文件列表
;     if (input = "") {
;         ; 从资源管理器获取选中的文件
;         files := get_selected_files(2)  ; 2 表示仅允许选择文件
;         if (files = "") {
;             Msgbox("请先在资源管理器中选中一个或多个 PowerPoint 文件！")
;             return
;         }
;         file_list := StrSplit(files, "`n", "`r")
;     } else if IsObject(input) {
;         ; 如果传入的是数组（如 ["a.pptx", "b.pptx"]）
;         file_list := input
;     } else {
;         ; 如果传入的是单个字符串路径
;         file_list := [input]
;     }

;     total_files := file_list.Length
;     if (total_files = 0) {
;         Msgbox("未检测到有效文件。")
;         return
;     }

;     ; 启动一次 PowerPoint（复用，提升性能）
;     ppt_app := ComObject("PowerPoint.Application")
;     ; ppt_app.Visible := false  ; 默认不可见

;     processed_count := 0
;     for _, source_full in file_list {
;         source_full := Trim(source_full)
;         if (source_full = "" || !FileExist(source_full))
;             continue

;         ; 检查是否为 PPT/PPTX 文件（简单判断后缀）
;         ext := StrLower(SubStr(source_full, -4))
;         if (ext != "ppt" && ext != "pptx")    
;             continue

;         try {
;             process_single_ppt(ppt_app, source_full, 10)  ; 每10页一段
;             processed_count += 1
;         } catch as err {
;             Msgbox("处理文件时出错：`n" source_full "`nError: " err.Message)
;         }
;     }

;     ; 清理 PowerPoint
;     try ppt_app.Quit()
;     Sleep 1000
;     if ProcessExist("POWERPNT.EXE")
;         ProcessClose("POWERPNT.EXE")

;     Notify.show("✅ 共成功处理 " processed_count "/" total_files " 个文件！")
; }

; ; 处理单个 PPT 文件：先分段，再拆单页
; process_single_ppt(ppt_app, source_full, pages_per_file := 10) {
;     pi := path_info(source_full)
;     source_fname := pi.fname
;     source_folder := pi.folder

;     ; 打开源文件（只读）
;     source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0)

;     total_slides := source_pres.Slides.Count
;     start_page := 1

;     while (start_page <= total_slides) {
;         end_page := Min(start_page + pages_per_file - 1, total_slides)
;         batch_filename := source_folder . "\" . source_fname . "_多页「" . start_page . "-" . end_page . "」.pptx"

;         ; 创建批次文件
;         FileCopy(source_full, batch_filename, true)
;         batch_pres := ppt_app.Presentations.Open(batch_filename, false, false, false)

;         ; 删除范围外的幻灯片（从后往前删）
;         i := batch_pres.Slides.Count
;         while (i >= 1) {
;             idx := batch_pres.Slides(i).SlideIndex
;             if (idx < start_page || idx > end_page)
;                 batch_pres.Slides(i).Delete()
;             i -= 1
;         }

;         ; 清除节
;         try while batch_pres.SectionProperties.Count
;             batch_pres.SectionProperties.Delete(batch_pres.SectionProperties.Count, false)

;         batch_pres.Save()
;         batch_pres.Close()

;         ; 拆此批次为单页
;         split_batch_to_single_pages_internal(ppt_app, batch_filename)

;         start_page += pages_per_file
;     }

;     source_pres.Close()
; }

; ; 内部函数：将一个 _多页「X-Y」.pptx 拆成单页
; split_batch_to_single_pages_internal(ppt_app, batch_full) {
;     pi := path_info(batch_full)
;     fname := pi.fname
;     folder := pi.folder

;     ; 提取页码范围
;     if !RegExMatch(fname, "多页「(\d+)-(\d+)」$", &m) {
;         FileDelete(batch_full)  ; 非标准命名，直接删除临时文件
;         return
;     }
;     start_num := m[1], end_num := m[2]

;     base_name := RegExReplace(fname, "_多页「\d+-\d+」$", "_单页")
;     target_folder := folder . "\" . base_name
;     if (!FileExist(target_folder))
;         DirCreate(target_folder)

;     pres := ppt_app.Presentations.Open(batch_full, 1, 0, 0)
;     slide_count := pres.Slides.Count

;     Loop slide_count {
;         actual_page := start_num + A_Index - 1
;         output_file := target_folder . "\" . base_name . "「" . actual_page . "」.pptx"

;         FileCopy(batch_full, output_file, true)
;         single_pres := ppt_app.Presentations.Open(output_file, false, false, false)

;         ; 只保留第 A_Index 页
;         total := single_pres.Slides.Count
;         i := total
;         while (i >= 1) {
;             if (i != A_Index)
;                 single_pres.Slides(i).Delete()
;             i -= 1
;         }

;         ; 清节
;         try while single_pres.SectionProperties.Count
;             single_pres.SectionProperties.Delete(single_pres.SectionProperties.Count, false)

;         single_pres.Save()
;         single_pres.Close()
;     }

;     pres.Close()
;     FileDelete(batch_full)  ; 删除临时批次文件
; }











export_jpg_in_folder(source_full, skip_section := "原稿") {  
     
    skip_index := 0
    first_skip := 0
    end_skip := 0
    section_count := 0   

    if !source_full {
         files := get_selected_files(2)
        if (files = "") {
            Msgbox("请先选择要拆分的 PowerPoint 文件！")
            return
        }

        ; 只取第一个文件
        source_full := Trim(StrSplit(files, "`n", "`r")[1])
        if (!FileExist(source_full)) {
            Msgbox("文件不存在：`n" source_full)
            return
        }
    }
    
    ; 获取文件路径信息
    source_full     := path_info(source_full).full      ; 文件完整路径
    source_drive    := path_info(source_full).drive     ; 驱动器 (如 C:)
    source_dir      := path_info(source_full).dir       ; 目录（不含驱动器）
    source_folder   := path_info(source_full).folder    ; 文件夹（含驱动器）        
    source_fname    := path_info(source_full).fname     ; 文件名（不含后缀）
    source_ext      := path_info(source_full).ext       ; 文件后缀   
    source_file     := path_info(source_full).file      ; 文件名（含后缀）

    if (RegExMatch(source_fname, "^(.*)_完整$", &match)) {
        base_name := match[1]
    } else {
        base_name := source_fname  ; 如果不是以"_完整"结尾，使用原文件名
    }

    ; 使用提取的基础名称创建目标文件夹
    target_folder := source_folder . "\" . base_name . "_图片"

    ; 创建目标文件夹（如果不存在）
    if (!FileExist(target_folder)) {
        try {
            DirCreate(target_folder)
        } catch as err {
            Msgbox("无法创建文件夹：`n" target_folder "`nError: " err.Message)
            return
        }
    }   
    
    ; run(target_folder)   
        
    ; 创建 PowerPoint 应用程序对象
    ppt_app := ComObject("PowerPoint.Application")
    ; ppt_app.Visible := false ; 隐藏 PowerPoint 窗口

    ; 打开源文件
    try {
        source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0) ; 只读打开，不显示窗口
    } catch as err {
        A_Clipboard := err.Message
        Msgbox("无法打开文件：`n" source_full "`nError: " err.Message)
        return
    }

    ; 检查源文件是否有需要跳过的节
    if (skip_section != -1) {
        try {
            section_props := source_pres.SectionProperties
            section_count := section_props.Count

            loop section_count {
                section_name := section_props.Name(A_Index)                
                if (section_name = skip_section) {
                skip_index := A_Index                
                    break ; 只处理第一个名为"原稿"的节
                }               
            }
            
            if (skip_index > 0) {
                first_skip := source_pres.SectionProperties.FirstSlide(skip_index)
                skip_count := source_pres.SectionProperties.SlidesCount(skip_index)
                end_skip := first_skip + skip_count - 1
            }

        } catch as err {
            ; 如果检查节失败，继续处理，不中断整个流程
            Msgbox("检查原稿节时出错：`n" err.Message)
        }
    }

    target_index := 0
    ; 遍历每张幻灯片
    try {
        for slide in source_pres.Slides {
            page_no := slide.SlideIndex
            target_index += 1   
            
            ; 如果当前幻灯片在需要跳过的节中，跳过处理
            if (skip_index > 0 && page_no >= first_skip && page_no <= end_skip) {
                target_index -= 1 
                continue            
            }
            
            target_full := target_folder "\" base_name "_图片「" target_index "」.jpg"    
                        
            ; 输出图片
            try {
                slide.Export target_full, "jpg", 1920, 1080    ; 输出图片
            } catch as err {
                Msgbox("遍历每张幻灯片失败：`n" source_full "`nError: " err.Message)
                continue
            }    
        }
    } catch as err {
        Msgbox("关闭源文件失败：`n" source_full "`nError: " err.Message)
    }
    
    
    ; 关闭源文件
    try {
        source_pres.Close()
    } catch as err {
        Msgbox("关闭源文件失败：`n" source_full "`nError: " err.Message)
    }
    
    ; 提示完成
    Notify.show("Successfully exported to JPG.")

}

; } 将PPT文件导出为单页 ---------------------------------------------------------------------------------
split_in_folder(source_full, skip_section := "原稿") {
    ; path_info(Path, X*) drive dir fname ext folder file full  
    ; dir 不含驱动器，folder含驱动器，fname不含后缀，file含后缀
    ; 获取选中的文件列表
    skip_index := 0
    first_skip := 0
    end_skip := 0
    section_count := 0

    
    if !source_full {
        files := get_selected_files(2)
        if (files = "") {
            Msgbox("请先选择要拆分的 PowerPoint 文件！")
            return
        }

        ; 只取第一个文件
        source_full := Trim(StrSplit(files, "`n", "`r")[1])
        if (!FileExist(source_full)) {
            Msgbox("文件不存在：`n" source_full)
            return
        }
    }
    
    
    ; 获取文件路径信息
    source_full     := path_info(source_full).full      ; 文件完整路径
    source_drive    := path_info(source_full).drive     ; 驱动器 (如 C:)
    source_dir      := path_info(source_full).dir       ; 目录（不含驱动器）
    source_folder   := path_info(source_full).folder    ; 文件夹（含驱动器）        
    source_fname    := path_info(source_full).fname     ; 文件名（不含后缀）
    source_ext      := path_info(source_full).ext       ; 文件后缀   
    source_file     := path_info(source_full).file      ; 文件名（含后缀）

    if (RegExMatch(source_fname, "^(.*)_完整$", &match)) {
        base_name := match[1]
    } else {
        base_name := source_fname  ; 如果不是以"_完整"结尾，使用原文件名
    }

    ; 使用提取的基础名称创建目标文件夹
    target_folder := source_folder . "\" . base_name . "_单页"

    ; 创建目标文件夹（如果不存在）
    if (!FileExist(target_folder)) {
        try {
            DirCreate(target_folder)
        } catch as err {
            Msgbox("无法创建文件夹：`n" target_folder "`nError: " err.Message)
            return
        }
    }   
    
    ; run(target_folder)   

    ; 创建 PowerPoint 应用程序对象
    ppt_app := ComObject("PowerPoint.Application")
    ; ppt_app.Visible := false ; 隐藏 PowerPoint 窗口

    ; 打开源文件
    try {
        source_pres := ppt_app.Presentations.Open(source_full, 1, 0, 0) ; 只读打开，不显示窗口
    } catch as err {
        Msgbox("无法打开文件：`n" source_full "`nError: " err.Message)
        return
    }

    ; 检查源文件是否有需要跳过的节
    if (skip_section != -1) {
        try {
            section_props := source_pres.SectionProperties
            section_count := section_props.Count

            loop section_count {
                section_name := section_props.Name(A_Index)                
                if (section_name = skip_section) {
                   skip_index := A_Index                
                    break ; 只处理第一个名为"原稿"的节
                }               
            }
            
            if (skip_index > 0) {
                first_skip := source_pres.SectionProperties.FirstSlide(skip_index)
                skip_count := source_pres.SectionProperties.SlidesCount(skip_index)
                end_skip := first_skip + skip_count - 1
            }

        } catch as err {
            ; 如果检查节失败，继续处理，不中断整个流程
            Msgbox("检查原稿节时出错：`n" err.Message)
        }
    }
   
    target_index := 0
    ; 遍历每张幻灯片
    for slide in source_pres.Slides {
        page_no := slide.SlideIndex
        target_index += 1   
        
        ; 如果当前幻灯片在需要跳过的节中，跳过处理
        if (skip_index > 0 && page_no >= first_skip && page_no <= end_skip) {
            target_index -= 1 
            continue            
        }
        
        target_full := target_folder "\" base_name "_单页「" target_index "」.pptx"
        
        ; 复制源文件到目标路径
        try {
            FileCopy(source_full, target_full, true) ; 覆盖已存在的文件
        } catch as err {
            Msgbox("文件复制失败：`n" source_full "`nError: " err.Message)
            continue
        }

        ; 打开复制的文件
        try {
            target_pres := ppt_app.Presentations.Open(target_full, false, false, false)
        } catch as err {
            Msgbox("无法打开文件：`n" target_full "`nError: " err.Message)
            continue
        }

        ; 删除多余的幻灯片
        try {
            count := target_pres.Slides.Count
            loop count {
                j := count + 1 - A_Index
                if (j != page_no)
                    target_pres.Slides(j).Delete
            }
        } catch as err {
            Msgbox("删除幻灯片失败：`n" target_full "`nError: " err.Message)
            target_pres.Close()
            continue
        }

        ; 删除所有的节信息，不删除slide
        try {            
            while target_pres.SectionProperties.Count {
                target_pres.SectionProperties.Delete(target_pres.SectionProperties.Count, false)
            }
        } catch as err {
            Msgbox("无法删除节信息：`n" target_full "`nError: " err.Message)
        }

        ; 保存并关闭目标文件
        try {
            target_pres.Save()
            target_pres.Close()
        } catch as err {
            Msgbox("保存或关闭文件失败：`n" target_full "`nError: " err.Message)
            continue
        }
    }
    
    ; 提示完成
    Notify.show("PowerPoint 文件已成功拆分为单页！")

    ; 关闭源文件
    try {
        source_pres.Close()
    } catch as err {
        Msgbox("关闭源文件失败：`n" source_full "`nError: " err.Message)
    }

    ; 确保 PowerPoint 应用程序关闭
    try {
        if IsObject(ppt_app) {
            ppt_app.Quit()
            
            ; 等待一会儿让 PowerPoint 正常关闭
            Sleep 2000
            
            ; 如果 PowerPoint 进程仍然存在，强制终止
            if ProcessExist("POWERPNT.EXE") {
                ProcessClose("POWERPNT.EXE")
                Sleep 500
            }
        }
    } catch as err {
        ; 如果正常退出失败，尝试强制终止进程
        try {
            if ProcessExist("POWERPNT.EXE") {
                ProcessClose("POWERPNT.EXE")
            }
        } catch {
            ; 忽略错误
        }
    }

}




same_folder() {	
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("移动操作只能在文件夹状态执行，请打开文件夹再操作！")
        return
    }  

    filenames := get_selected_files(2) 
      
    if !filenames {          
        Msgbox("没有选中任何文件！")    
        return
    }  
    Loop Parse, filenames, '`n', '`r' {
        source_full     := path_info(A_LoopField).Full  
        source_fname    := path_info(A_LoopField).Fname
        source_folder   := path_info(A_LoopField).Folder
        source_ext      := path_info(A_LoopField).Ext 
        dest_full       := path_info(A_LoopField, "Ext: ").Full
            
        if !source_ext {          
            Msgbox("选中的是文件夹，无法完成操作！")    
            return
        }  
        
        if FileExist(dest_full){
            Msgbox("存在同名文件夹，无法完成操作！")
            return            
        }  
        
        try DirCreate(dest_full)        
        catch as err
            Msgbox(err.message) 

        try FileMove(source_full, dest_full, 0)      
        catch as err
            Msgbox(err.message) 	
    } 
}

move_in_one_same_named_folder() {	
    if !WinActive("ahk_class CabinetWClass") {
        Msgbox("移动操作只能在文件夹状态执行，请打开文件夹再操作！")
        return
    }  

    filenames := get_selected_files(2) 
      
    if !filenames {          
        Msgbox("没有选中任何文件！")    
        return
    }  
    Loop Parse, filenames, '`n', '`r' {
        source_full     := path_info(A_LoopField).Full  
        source_fname    := path_info(A_LoopField).Fname
        source_folder   := path_info(A_LoopField).Folder
        source_ext      := path_info(A_LoopField).Ext   

        if !source_ext {          
            Msgbox("选中的是文件夹，无法完成操作！")    
            return
        }          

        if(A_Index==1) {       
            dest_full := path_info(A_LoopField, "Ext: ").Full             
            if FileExist(dest_full){
                Msgbox("存在同名文件夹，无法完成操作！")
                return            
            }  
            try DirCreate(dest_full)        
            catch as err
                Msgbox(err.message) 
        }
        
        try FileMove(source_full, dest_full, 0)      
        catch as err
            Msgbox(err.message) 	
    } 
}


; #endregion

; #region qwert
save_reload() {
    If WinActive("ahk_exe Code.exe") {
        SendInput("^s")
        Sleep(50)
    }         
    Reload
}
 

; #endregion



; 将选中的 PPT 文件先分成10页一批，再拆成单页
split_selected_ppt_to_batches_then_single() {
    files := get_selected_files(2)
    if (files = "") {
        Msgbox("请先选择要拆分的 PowerPoint 文件！")
        return
    }

    file_list := StrSplit(files, "`n", "`r")
    for _, file_path in file_list {
        source_full := Trim(file_path)
        if (!FileExist(source_full))
            continue

        split_ppt_to_batches_then_single_core(source_full, 10)
    }
    
    Notify.show("PPT 分割完成！")
}

; PPT 分割核心函数：先分成多页一批，再拆成单页
split_ppt_to_batches_then_single_core(source_full, pages_per_batch := 10) {
    pi := path_info(source_full)
    source_fname := pi.fname
    source_folder := pi.folder

    ; 启动 PowerPoint
    try {
        ppt_app := ComObject("PowerPoint.Application")
        ppt_app.Visible := True  ; 设置为可见以确保正常工作
    } catch {
        Msgbox("无法启动 PowerPoint 应用程序！")
        return
    }

    ; 打开源文件
    try {
        source_pres := ppt_app.Presentations.Open(source_full, , , 0)  ; 以只读模式打开
        total_slides := source_pres.Slides.Count
    } catch as err {
        Msgbox("无法打开源文件：`n" source_full "`nError: " err.Message)
        try ppt_app.Quit()
        return
    }
    
    ; 第一步：按指定页数分割成多页文件
    start_page := 1
    while (start_page <= total_slides) {
        end_page := Min(start_page + pages_per_batch - 1, total_slides)
        batch_filename := source_folder . "\" . source_fname . "_多页「" . start_page . "-" . end_page . "」.pptx"
        
        ; 复制源文件作为基础
        FileCopy(source_full, batch_filename, true)
        
        ; 打开副本并删除范围外的页面
        try {
            batch_pres := ppt_app.Presentations.Open(batch_filename, , , 0)  ; 以只读模式打开
            
            ; 从后往前删除不在范围内的幻灯片
            i := batch_pres.Slides.Count
            while (i >= 1) {
                slide_index := batch_pres.Slides(i).SlideIndex
                if (slide_index < start_page || slide_index > end_page)
                    batch_pres.Slides(i).Delete()
                i--
            }
            
            ; 清除节信息
            try {
                while batch_pres.SectionProperties.Count
                    batch_pres.SectionProperties.Delete(batch_pres.SectionProperties.Count, false)
            }
            
            batch_pres.Save()
            batch_pres.Close()
        } catch as err {
            Msgbox("处理批次文件时出错：`n" batch_filename "`nError: " err.Message)
            ; 继续处理下一个
        }
        
        ; 第二步：将这个多页文件拆成单页
        split_batch_to_single_pages(batch_filename, start_page, end_page, ppt_app)
        
        start_page += pages_per_batch
    }

    ; 清理
    try {
        source_pres.Close()
        ppt_app.Quit()
        Sleep(1000)
        if ProcessExist("POWERPNT.EXE")
            ProcessClose("POWERPNT.EXE")
    } catch {
        ; 如果正常关闭失败，尝试强制关闭
        if ProcessExist("POWERPNT.EXE")
            ProcessClose("POWERPNT.EXE")
    }
}

; 将多页文件拆成单页
split_batch_to_single_pages(batch_full, start_num, end_num, ppt_app) {
    pi := path_info(batch_full)
    fname := pi.fname
    folder := pi.folder
    
    ; 创建单页输出文件夹
    base_name := RegExReplace(fname, "_多页「\d+-\d+」$", "_单页")
    target_folder := folder . "\" . base_name
    if (!FileExist(target_folder))
        DirCreate(target_folder)

    ; 验证 PowerPoint 应用程序对象是否仍然有效
    try {
        if (!IsObject(ppt_app) || ppt_app.Name = "") {
            ppt_app := ComObject("PowerPoint.Application")
            ppt_app.Visible := True
        }
    } catch {
        ppt_app := ComObject("PowerPoint.Application")
        ppt_app.Visible := True
    }
    
    ; 遍历每一页进行分割
    Loop (end_num - start_num + 1) {
        current_slide_index := A_Index
        actual_page_number := start_num + A_Index - 1
        output_file := target_folder . "\" . base_name . "「" . actual_page_number . "」.pptx"
        
        ; 复制源文件
        FileCopy(batch_full, output_file, true)
        
        ; 打开副本并删除其他页面
        try {
            single_pres := ppt_app.Presentations.Open(output_file, , , 0)  ; 以只读模式打开
            
            i := single_pres.Slides.Count
            while (i >= 1) {
                if (i != current_slide_index)
                    single_pres.Slides(i).Delete()
                i--
            }
            
            ; 清除节信息
            try {
                while single_pres.SectionProperties.Count
                    single_pres.SectionProperties.Delete(single_pres.SectionProperties.Count, false)
            }
            
            single_pres.Save()
            single_pres.Close()
        } catch as err {
            Msgbox("处理单页文件时出错：`n" output_file "`nError: " err.Message)
        }
    }
    
    ; 删除临时的多页文件
    FileDelete(batch_full)
}









; process_powerpoint_advanced(split_operation := false, export_operation := false, organize_operation := false) {
;     local files := []

;     if (split_operation || export_operation || organize_operation) {
;         ; 获取文件列表
;         files := get_file_list_from_context()
;         if (files.Length = 0) {
;             return
;         }
;     }
       
;     ; 检查是否至少选择了一个操作
;     if (!split_operation && !export_operation && !organize_operation) {
;         Msgbox("请至少选择一个操作：分离、导出或整理！")
;         return
;     }
    
;     ; 处理所有文件
;     for index, file_path in files {
;         source_full := Trim(file_path)
        
;         if (!FileExist(source_full)) {
;             Msgbox("文件不存在：`n" . source_full)
;             continue  ; 跳过此文件，继续处理下一个
;         }

;         ; 获取文件路径信息
;         source_info := path_info(source_full)
;         source_full := source_info.full      ; 文件完整路径
;         source_folder := source_info.folder  ; 文件夹（含驱动器）        
;         source_fname := source_info.fname    ; 文件名（不含后缀）
;         source_ext := source_info.ext        ; 文件后缀   
;         source_file := source_info.file      ; 文件名（含后缀）

;         try {
;             ; 根据参数执行相应的操作
;             if (split_operation) {  ; PPT分离
;                 split_in_folder_smart(source_full)
;                 Notify.show("PPT分离完成：" . source_fname, 0.5, "00aa00")
;             }
            
;             if (export_operation) {  ; JPG导出
;                 export_jpg_in_folder(source_full)
;                 Notify.show("JPG导出完成：" . source_fname, 0.5, "00aa00")
;             }
            
;             ; 添加延迟确保PowerPoint完全关闭
;             Sleep(1000)

;             if (organize_operation) {  ; 文件整理
;                 ; 获取祖父目录
;                 source_grandparent_folder := RegExReplace(source_full, "\\[^\\]+\\[^\\]+$", "") . "\"
                
;                 target_folder := source_grandparent_folder . source_fname . "\"
                
;                 ; 移动整个文件夹到以PPT文件名命名的新文件夹
;                 try {
;                     DirMove(source_folder, target_folder, "R")
;                 } catch as err {
;                     Msgbox("无法移动文件夹：`n" . source_folder . "`nError: " . err.Message)
;                     continue  ; 跳过此文件，继续处理下一个
;                 }
                
;                 ; 更新路径变量
;                 new_source_full := target_folder . source_file
;                 new_source_folder := path_info(new_source_full).folder
                
;                 source_font_full := new_source_folder . "字体"
;                 target_font_full := new_source_folder . source_fname . "_字体"
;                 target_wz_fold := new_source_folder . source_fname . "_完整"
;                 target_wz_full := new_source_folder . source_fname . "_完整\" . source_fname . "_完整.pptx"

;                 ; 移动完整文件
;                 try {
;                     if (!FileExist(target_wz_fold)) {            
;                         DirCreate(target_wz_fold)
;                     }   
;                     FileMove(new_source_full, target_wz_full)
;                 } catch as err {
;                     Msgbox("无法移动完整文件：`n" . new_source_full . "`nError: " . err.Message)
;                     continue  ; 跳过此文件，继续处理下一个
;                 }
                
;                 ; 移动字体文件夹（如果存在）
;                 if DirExist(source_font_full) {
;                     try {
;                         DirMove(source_font_full, target_font_full)
;                     } catch as err {
;                         Msgbox("无法移动字体文件夹：`n" . target_font_full . "`nError: " . err.Message)
;                         continue  ; 跳过此文件，继续处理下一个
;                     } 
;                 }
                
;                 Notify.show("文件整理完成：" . source_fname, 0.5, "00aa00")
;             }
            
;         } catch as err {
;             Msgbox("处理文件时出现错误：" . source_fname . "`n" . err.Message)
;             continue  ; 跳过此文件，继续处理下一个
;         }
;     }
    
;     Notify.show("所有操作完成！", 2, "00aa00")
    
    
; }

; ; 为常见的操作组合创建便捷函数
; process_powerpoint_split_only() {
;     process_powerpoint_advanced(true, false, false)  ; 仅PPT分离
; }

; process_powerpoint_export_jpg_only() {
;     process_powerpoint_advanced(false, true, false)  ; 仅JPG导出
; }

; process_powerpoint_organize_only() {
;     process_powerpoint_advanced(false, false, true)  ; 仅文件整理
; }

; process_powerpoint_split_and_export() {
;     process_powerpoint_advanced(true, true, false)  ; PPT分离 + JPG导出
; }

; process_powerpoint_split_and_organize() {
;     process_powerpoint_advanced(true, false, true)  ; PPT分离 + 文件整理
; }

; process_powerpoint_export_and_organize() {
;     process_powerpoint_advanced(false, true, true)  ; JPG导出 + 文件整理
; }

; process_powerpoint_full_process() {
;     process_powerpoint_advanced(true, true, true)  ; 全部操作
; }



; split_in_folder_smart(source_full) {
;     try {
;         ; 创建PowerPoint应用程序对象
;         ppt_app := ComObject("PowerPoint.Application")
        
;         ; 打开演示文稿
;         ppt_pres := ppt_app.Presentations.Open(source_full, , , false) ; 不显示消息，不显示动画
        
;         ; 获取幻灯片总数
;         total_slides := ppt_pres.Slides.Count
        
;         if (total_slides <= 0) {
;             ppt_pres.Close()
;             ppt_app.Quit()
;             Msgbox("PPT中没有幻灯片")
;             return
;         }
        
;         ; 根据总幻灯片数确定批量大小
;         ; 如果幻灯片总数 > 100，每批20张
;         ; 如果幻灯片总数 > 20 但 <= 50，每批10张
;         ; 如果幻灯片总数 <= 20，不拆分（直接跳过批量拆分，只拆分为单页）
;         if (total_slides <= 20) {
;             ; 不进行批量拆分，直接进行单页拆分
;             split_to_single_only2(source_full, total_slides)
;             ppt_pres.Close()
;             ppt_app.Quit()
;             return
;         }
        
;         batch_size := (total_slides > 100) ? 20 : 10
        
;         ; 获取源文件信息
;         source_info := path_info(source_full)
;         source_folder := source_info.Folder
;         source_fname := source_info.Fname
;         source_ext := source_info.Ext
        
;         ; 为拆分后的演示文稿创建子文件夹
;         split_folder := source_folder . "\" . source_fname . "_拆分"
;         if !DirExist(split_folder)
;             DirCreate(split_folder)
        
;         ; 按批量拆分
;         current_slide := 1
;         batch_num := 1
        
;         while (current_slide <= total_slides) {
;             start_slide := current_slide
;             end_slide := Min(current_slide + batch_size - 1, total_slides)
            
;             ; 复制幻灯片范围
;             Loop (end_slide - start_slide + 1) {
;                 slide_index := start_slide + A_Index - 1
;                 ppt_pres.Slides(slide_index).Copy()
;             }
            
;             ; 为此批量创建新演示文稿
;             new_pres := ppt_app.Presentations.Add()
            
;             ; 粘贴幻灯片
;             new_pres.Slides.Paste()
            
;             ; 生成批量文件名 (格式: xxxxxxxxxxxxx_多页「1-10」.pptx)
;             batch_filename := split_folder . "\" . source_fname . "_多页「" . start_slide . "-" . end_slide . "」.pptx"
            
;             ; 保存批量演示文稿
;             new_pres.SaveAs(batch_filename)
;             new_pres.Close()
            
;             current_slide := end_slide + 1
;             batch_num++
;         }
        
;         ; 现在将每批进一步拆分为单独幻灯片
;         split_each_batch_to_single2(split_folder, source_fname)
        
;         ; 关闭原始演示文稿
;         ppt_pres.Close()
;         ppt_app.Quit()
        
;         Notify.show("PPT已成功按批次拆分，并将每批进一步拆分为单页")
;     } catch as err {
;         Msgbox("拆分PPT时出错: " . err.message)
;     }
; }

; ; 只拆分为单页（不进行批量拆分）的函数
; split_to_single_only2(source_full, total_slides) {
;     try {
;         ; 创建PowerPoint应用程序对象
;         ppt_app := ComObject("PowerPoint.Application")
        
;         ; 打开演示文稿
;         ppt_pres := ppt_app.Presentations.Open(source_full, , , false)
        
;         ; 获取源文件信息
;         source_info := path_info(source_full)
;         source_folder := source_info.Folder
;         source_fname := source_info.Fname
        
;         ; 为单页创建文件夹
;         single_folder := source_folder . "\" . source_fname . "_单页"
;         if !DirExist(single_folder)
;             DirCreate(single_folder)
        
;         ; 将每张幻灯片拆分为单独文件
;         Loop total_slides {
;             slide_index := A_Index
            
;             ; 复制幻灯片
;             ppt_pres.Slides(slide_index).Copy()
            
;             ; 为此单张幻灯片创建新演示文稿
;             single_pres := ppt_app.Presentations.Add()
            
;             ; 粘贴幻灯片
;             single_pres.Slides.Paste()
            
;             ; 保存为单独幻灯片文件 (格式: xxxxxxxxxxxxx_单页「184」.pptx)
;             single_filename := single_folder . "\" . source_fname . "_单页「" . slide_index . "」.pptx"
;             single_pres.SaveAs(single_filename)
;             single_pres.Close()
;         }
        
;         ppt_pres.Close()
;         ppt_app.Quit()
        
;         Notify.show("PPT已成功拆分为单页幻灯片")
;     } catch as err {
;         Msgbox("拆分单页时出错: " . err.message)
;     }
; }

; ; 将每批拆分为单独幻灯片的辅助函数
; split_each_batch_to_single2(split_folder, source_fname) {
;     try {
;         ; 创建PowerPoint应用程序对象
;         ppt_app := ComObject("PowerPoint.Application")
        
;         ; 遍历拆分文件夹中的所有批量文件
;         Loop Files, split_folder . "\*.pptx", "F" {
;             batch_file := A_LoopFileFullPath
            
;             ; 检查是否是批量文件（包含"多页"字样）
;             if !InStr(A_LoopFileName, "多页") {
;                 continue
;             }
            
;             ; 打开批量演示文稿
;             batch_pres := ppt_app.Presentations.Open(batch_file, , , false)
;             batch_slides_count := batch_pres.Slides.Count
            
;             ; 为此批量的单独幻灯片创建子文件夹
;             batch_name := path_info(batch_file).Fname
;             single_folder := split_folder . "\" . batch_name . "_单页"
;             if !DirExist(single_folder)
;                 DirCreate(single_folder)
            
;             ; 将批量中的每张幻灯片拆分为单独文件
;             Loop batch_slides_count {
;                 slide_index := A_Index
;                 actual_slide_num := get_actual_slide_number(batch_file, slide_index) ; 获取实际幻灯片编号
                
;                 ; 复制幻灯片
;                 batch_pres.Slides(slide_index).Copy()
                
;                 ; 为此单张幻灯片创建新演示文稿
;                 single_pres := ppt_app.Presentations.Add()
                
;                 ; 粘贴幻灯片
;                 single_pres.Slides.Paste()
                
;                 ; 保存为单独幻灯片文件 (格式: xxxxxxxxxxxxx_单页「184」.pptx)
;                 single_filename := single_folder . "\" . source_fname . "_单页「" . actual_slide_num . "」.pptx"
;                 single_pres.SaveAs(single_filename)
;                 single_pres.Close()
;             }
            
;             batch_pres.Close()
;         }
        
;         ppt_app.Quit()
;     } catch as err {
;         Msgbox("拆分单页时出错: " . err.message)
;     }
; }

; ; ; 获取实际幻灯片编号的辅助函数
; get_actual_slide_number(batch_file, slide_index_in_batch) {
;     ; 从文件名中提取起始幻灯片号
;     if RegExMatch(batch_file, "_多页「(\d+)-(\d+)」", &match) {
;         start_num := match[1]
;         return start_num + slide_index_in_batch - 1
;     }
;     return slide_index_in_batch
; }

















; split_ppt_to_single(mode := "clipboard") {
;     try {
;         selected_files := ""
        
;         ; 根据模式确定文件来源
;         if (mode = "clipboard") {
;             ; 从剪贴板获取文件路径
;             selected_files := A_Clipboard
;             if !selected_files {
;                 Msgbox("剪贴板中没有文件路径")
;                 return
;             }
;             ; 将剪贴板内容按行分割
;             selected_files_array := StrSplit(selected_files, "`n", "`r")
;             files_list := ""
;             for index, file_path in selected_files_array {
;                 ; 清理路径：移除多余的引号、空格和换行符
;                 cleaned_path := Trim(file_path, "`t`n`r`"")
;                 if FileExist(cleaned_path) {
;                     files_list .= cleaned_path . "`n"
;                 }
;             }
;             selected_files := RTrim(files_list, "`n")
;         } else if (mode = "selected") {
;             ; 从文件夹选中获取文件
;             selected_files := get_selected_files(2)
;         } else {
;             Msgbox("无效的模式。使用 'clipboard' 或 'selected'")
;             return
;         }
        
;         if !selected_files {
;             Msgbox("未选择任何文件")
;             return
;         }

;         ; 遍历选中的文件
;         Loop Parse, selected_files, "`n", "`r" {
;             source_full := path_info(A_LoopField).Full

;             ; 检查文件是否存在
;             if !FileExist(source_full) {
;                 Msgbox("文件不存在：`n" source_full)
;                 continue
;             }

;             ; 检查是否为PPT文件
;             ext := path_info(source_full).Ext
;             if (ext != ".ppt" && ext != ".pptx") {
;                 Msgbox("文件不是PowerPoint文件：`n" source_full)
;                 continue
;             }

;             ; 创建PowerPoint应用程序对象
;             ppt_app := ComObject("PowerPoint.Application")

;             ; 打开演示文稿
;             ppt_pres := ppt_app.Presentations.Open(source_full, , , false) ; don't show message, don't show animation

;             ; 获取幻灯片总数
;             total_slides := ppt_pres.Slides.Count

;             if (total_slides <= 0) {
;                 ppt_pres.Close()
;                 ppt_app.Quit()
;                 Msgbox("PPT中没有幻灯片")
;                 continue
;             }

;             ; 获取源文件信息
;             source_info := path_info(source_full)
;             source_folder := source_info.Folder
;             source_fname := source_info.Fname
;             source_ext := source_info.Ext

;             ; 为拆分后的演示文稿创建子文件夹
;             split_folder := source_folder . "\" . source_fname . "_拆分"
;             if !DirExist(split_folder)
;                 DirCreate(split_folder)

;             ; 根据总幻灯片数确定批量大小
;             ; 如果幻灯片总数 > 100，每批40张
;             ; 如果幻灯片总数 > 20 但 <= 100，每批10张
;             ; 如果幻灯片总数 <= 20，不拆分（直接跳过批量拆分，只拆分为单页）
;             if (total_slides <= 20) {
;                 ; 不进行批量拆分，直接进行单页拆分
;                 split_to_single_only(source_full, total_slides, ppt_app, ppt_pres)
;                 continue
;             }

;             batch_size := (total_slides > 100) ? 40 : 10

;             ; 按批量拆分
;             current_slide := 1
;             batch_num := 1

;             while (current_slide <= total_slides) {
;                 start_slide := current_slide
;                 end_slide := Min(current_slide + batch_size - 1, total_slides)

;                 ; 复制幻灯片范围
;                 Loop (end_slide - start_slide + 1) {
;                     slide_index := start_slide + A_Index - 1
;                     ppt_pres.Slides(slide_index).Copy()
;                 }

;                 ; 为此批量创建新演示文稿
;                 new_pres := ppt_app.Presentations.Add()

;                 ; 粘贴幻灯片
;                 new_pres.Slides.Paste()

;                 ; 生成批量文件名 (格式: xxxxxxxxxxxxx_多页「1-10」.pptx)
;                 batch_filename := split_folder . "\" . source_fname . "_多页「" . start_slide . "-" . end_slide . "」.pptx"

;                 ; 保存批量演示文稿
;                 new_pres.SaveAs(batch_filename)
;                 new_pres.Close()

;                 current_slide := end_slide + 1
;                 batch_num++
;             }

;             ; 现在将每批进一步拆分为单独幻灯片
;             split_each_batch_to_single(split_folder, source_fname)

;             ; 关闭原始演示文稿
;             ppt_pres.Close()
;             ppt_app.Quit()

;             Notify.show("PPT已成功按批次拆分，并将每批进一步拆分为单页")
;         }
;     } catch as err {
;         Msgbox("拆分PPT时出错: " . err.message)
;     }
; }









; 粘贴第 2 条记录（索引 1）

test_ditto()   {     
    ; ditto := ComObject("Ditto.Ditto")
    ; ditto.PasteClip(2)  ; 相当于在 Ditto 中选中第2项并粘贴
    Run "C:\Program Files\Ditto\Ditto.exe /Paste:2452" 

    MsgBox(A_Clipboard)
}





; #region  call_function_by_string  ——————————————————————————————————————————————————————




; 常用Windows CLSID列表
/*
系统文件夹:
::{645FF040-5081-101B-9F08-00AA002F954E}  ; 回收站
::{20D04FE0-3AEA-1069-A2D8-08002B30309D}  ; 我的电脑/此电脑
::{208D2C60-3AEA-1069-A2D7-08002B30309D}  ; 网络
::{871C5380-42A0-1069-A2EA-08002B30309D}  ; 网络和拨号连接
::{2227A280-3AEA-1069-A2DE-08002B30309D}  ; 打印机和传真
::{7007ACC7-3202-11D1-AAD2-00805FC1270E}  ; 网络连接

控制面板项目:
::{26EE0668-A00A-44D7-9371-BEB064C98683}  ; 控制面板
::{58E3C745-D971-4081-9034-86E34B30836A}  ; 语音识别
::{78F3955E-3B90-4184-BD14-5397C15F1EFC}  ; 性能信息和工具
::{BB64F8A7-BEE7-4E1A-AB8D-7D8273F7FDB6}  ; 操作中心
::{ED834ED6-4B5A-4BFE-8F11-A626DCB6A921}  ; 个性化

用户文件夹:
::{59031a47-3f72-44a7-89c5-5595fe6b30ee}  ; 用户文件夹
::{374DE290-123F-4565-9164-39C4925E467B}  ; 下载文件夹
::{1777F761-68AD-4D8A-87BD-30B759FA33DD}  ; 收藏夹
::{A63293E8-664E-48DB-A079-DF759E0509F7}  ; 模板文件夹
*/


/**
 * 解析函数调用字符串
 * @param function_string 函数字符串
 * @return 解析结果对象
 */
parse_function_call(function_string) {
    ; 检查函数调用格式：function_name(args)
    if(RegExMatch(function_string, "^(\w+)\s*\((.*)\)$", &match)) {
        return {
            name: match[1],
            params: parse_function_parameters(match[2]),
            has_params: true
        }
    }
    
    ; 简写格式：function_name
    if(RegExMatch(function_string, "^\w+$")) {
        return {
            name: function_string,
            params: [],
            has_params: false
        }
    }
    
    Throw Error("无法识别的函数格式: " . function_string)
}

/**
 * 智能解析函数参数
 * @param param_string 参数字符串
 * @return 参数数组
 */
parse_function_parameters(param_string) {
    if(!param_string || param_string == "") {
        return []
    }
    
    params := []
    current_param := ""
    in_quotes := false
    quote_char := ""
    
    ; 逐字符解析，正确处理引号内的逗号
    Loop StrLen(param_string) {
        char := SubStr(param_string, A_Index, 1)
        
        switch char {
            case '"', "'":
                if(!in_quotes) {
                    in_quotes := true
                    quote_char := char
                } else if(char = quote_char) {
                    in_quotes := false
                    quote_char := ""
                }
                current_param .= char
                
            case ",":
                if(in_quotes) {
                    current_param .= char
                } else {
                    ; 参数分隔符
                    params.Push(clean_parameter(current_param))
                    current_param := ""
                }
                
            default:
                current_param .= char
        }
    }
    
    ; 添加最后一个参数
    if(current_param != "") {
        params.Push(clean_parameter(current_param))
    }
    
    return params
}

/**
 * 清理参数：去除前后空格和外层引号
 * @param param 原始参数
 * @return 清理后的参数
 */
clean_parameter(param) {
    param := Trim(param)
    
    ; 去除外层引号
    if((StrLen(param) >= 2) && 
       ((SubStr(param, 1, 1) = '"' && SubStr(param, -1) = '"') ||
        (SubStr(param, 1, 1) = "'" && SubStr(param, -1) = "'"))) {
        param := SubStr(param, 2, StrLen(param) - 2)
    }
    
    return param
}

/**
 * 执行解析后的函数
 * @param call_info 解析后的调用信息
 * @return 执行结果
 */
execute_parsed_function(call_info) {
    ; 获取函数对象
    try {
        function_obj := %call_info.name%
    } catch {
        Throw Error("函数 '" . call_info.name . "' 不存在")
    } 

    ; 执行函数调用
    if(!call_info.has_params || call_info.params.Length = 0) {
        function_obj.Call()
    } else {
        ; 使用展开操作符支持任意数量参数
        function_obj.Call(call_info.params*)
    }
    
    return true
}


call_function(func_str) {
    parsed := parse_function_call(func_str)
    return execute_parsed_function(parsed)
}

/**
 * 运行应用程序
 * @param app_path 应用程序路径
 * @param param 启动参数
 */
run_app2(app_path, param := "") {

    app_process := "ahk_exe " . path_info(app_path).File       
    ; 特殊应用处理：总是启动新实例
    if RegExMatch(app_path, "i)eagle|fxsound|baidunetdisk|wxwork") {  
        try {
            Run(app_path . " " . param)
        } catch as err {
            Msgbox(err.Message)    
        } 
        return    
    }  
    
    ; 微信特殊处理
    if (RegExMatch(app_path, "i)weixin\.exe") && ProcessExist("Weixin.exe")) {
        SendInput("^!w")
        return
    }    

    ; 一般应用处理：存在则激活，不存在则启动
    if WinExist(app_process) {  
        WinActivate()
    } else {
        try {
            Run(app_path . " " . param)
        } catch as err {
            Msgbox(err.Message)        
        }         
    }
}

/**
 * 运行服务器应用程序，先给出确认对话框
 * @param app_path 应用程序路径
 * @param param 启动参数 (可选)
 */
run_server_app(app_path, param := "") {
    ; 创建完整的命令行
    command := app_path . " " . param
      
    result := MsgBox("即将运行服务器程序, 确定要继续吗？", "确认", "OKCancel")
    
    ; 检查用户的选择
    if (result = "Ok") 
        Run(command)
   
}
/**
 * 执行 PowerPoint 命令
 * @param command PPT 命令
 */


; #endregion
 



; #region Miscellaneous


; AutoHotkey v2 脚本：更新当前 Excel 的 K 列，基于 H 列的关键字匹配
; 要求：Excel 已打开且活动工作表为要处理的 sheet

; 关键字映射表（原始关键字 → 目标字符串）

cmdb_update_excel_k_column() {
keywordMap := Map(
    "卫视MCPC"      ,   "地球站卫视MCPC系统"     ,
    "东方卫视高清"  ,   "地球站东方卫视高清系统" ,
    "东方卫视4K"    ,   "地球站东方卫视4K系统"   ,
    "贵江河"        ,   "地球站贵江河系统"       ,
    "浙江高清"      ,   "地球站浙江高清系统"     ,
    "互动HD1"       ,   "地球站互动HD1系统"      ,
    "互动HD2"       ,   "地球站互动HD2系统"      ,
    "互动HD4"       ,   "地球站互动HD4系统"      ,
    "欢笑剧场4K"    ,   "地球站欢笑剧场4K系统"   ,
    "浙江标清"      ,   "地球站浙江标清系统"     ,
    "传送"          ,   "地球站传送系统"
)


; 获取当前 Excel 应用程序对象
try
    xlApp := ComObjActive("Excel.Application")
catch
{
    MsgBox("无法连接到 Excel，请确保 Excel 已打开。")
    ExitApp
}

; 获取当前活动工作簿和工作表
wb := xlApp.ActiveWorkbook
if !wb
{
    MsgBox("没有活动工作簿。")
    ExitApp
}
ws := wb.ActiveSheet

; 获取 H 列（第8列）和 K 列（第11列）的数据范围
usedRows := ws.UsedRange.Rows.Count

Loop usedRows
{
    row := A_Index
    if (row = 1) ; 跳过标题行（可选，根据实际情况调整）
        Continue

    cellValue := ws.Cells(row, 8).Value ; H 列是第8列
    if (!cellValue || cellValue = "")
        Continue

    ; 将单元格内容转为字符串（兼容数字/文本）
    text := "" . cellValue  ; 强制转为字符串

    ; 存储匹配结果
    matches := []

    ; 检查每个关键字是否出现在 text 中
    for keyword, target in keywordMap
    {
        if InStr(text, keyword)
            matches.Push(target)
    }

    ; 去重（防止重复关键词多次匹配）
    uniqueMatches := []
    seen := Map()
    for item in matches
    {
        if !seen.Has(item)
        {
            uniqueMatches.Push(item)
            seen[item] := true
        }
    }

    ; 生成结果字符串，用英文分号分隔
    result := ""
    if uniqueMatches.Length > 0
        result := StrJoin(uniqueMatches, ";")

    ; 写入 K 列（第11列）
    ws.Cells(row, 11).Value := result
}

MsgBox("CMDB 更新完成！已处理 " usedRows " 行。")
}

; dis_mac() {	
; 	try {
; 		send_via_clipboard("`n`n`ndisplay mac-address`n")	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; dis_route() {	
; 	try {
; 		send_via_clipboard("`n`n`ndisplay ip routing-table`n")	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; dis_ospf() {	
; 	try {
; 		send_via_clipboard("`n`n`ndisplay ospf brief`n")	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; int_g0() {	
; 	try {
; 		send_via_clipboard("int g0/0/0`n")	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; int_g1() {	
; 	try {
; 		send_via_clipboard("int g0/0/1`n")	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; int_g2() {	
; 	try {
; 		send_via_clipboard("int g0/0/2`n")	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; int_g3() {	
; 	try {
; 		send_via_clipboard("int g0/0/3`n")	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; zero_dot() {	
; 	try {
; 		send_via_clipboard('.0.')	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; zero_dot2() {	
; 	try {                                                                                       
; 		send_via_clipboard('.0.0.')	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; zero_dot3() {	
; 	try {
; 		send_via_clipboard('.0.0.0')	
; 	}
; 	catch as err
; 		display_error_dialog(err.message)   
; }

; #endregion




; #region stepclip





step_clipxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx() {
    try {
        static count := 0  
        static target_name := ""         
        
        if !count
            SetTimer waiting_logic, -40  

        (count >= 8) ? (count := 1) : (count++)   

        target_name := clipboard_history[count]
        
    } catch as err {
        count := 0
        MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
    }  
         
    waiting_logic() {   
        try {
            if GetKeyState(extract_modifier(), "P") { 
                SetTimer , -40 
                return
            } 
            count := 0
            ; Notify.Destroy

            SendInput(target_name)
            ; Sleep(100)

        } catch as err {
            count := 0
            MsgBox(Format("Error: {1}`nFile: {2}`nLine:  {3}", err.message, err.file, err.Line))
        }   
    }  
}
   




copy_text_in_ppt() {
     try {
        ppt_application := ComObjActive("PowerPoint.Application")
        selection := ppt_application.ActiveWindow.Selection
        
        ; 检查选择类型是否为形状
        if (selection.Type < 2) {  ; 2 = shapes
            Msgbox("请先选择要粘贴纯文本的形状")
            return
        }

        if (selection.Type == 3) {  ;  3 = text
                SendInput("^v")
            return
        }
        
        if (selection.HasChildShapeRange)
            shape_range := selection.ChildShapeRange
        else    
            shape_range := selection.ShapeRange  

        shape_range.TextFrame2.TextRange.Text := Trim(A_Clipboard, "`n`r")  

    } catch as err {
        Msgbox("粘贴纯文本时发生错误: " . err.Message)
    } finally {
        ppt_application := ""
    }

}


paste_format() {	
    try { 
        if WinActive("ahk_exe POWERPNT.EXE") {	
            SendInput("^+v")
        }
        if WinActive("ahk_exe WINWORD.EXE") {
            app := ComObjActive('Word.Application')
            app.Selection.PasteFormat
        }	
        if WinActive("ahk_exe EXCEL.EXE") {
            app := ComObjActive('Excel.Application')
            app.Selection.PasteSpecial(-4122)  ; -4122表示粘贴格式
        }	
    }
    catch as err
        Msgbox(err.Message)  
}


paste_pure_text() {	
    try {
        clip_saved := ClipboardAll() 
		A_Clipboard := A_Clipboard . ""
        Sleep(50) 
		SendInput("^v")	        
        Sleep(50) 
		A_Clipboard := clip_saved  
    }
    catch as err
        Msgbox(err.Message)   
}


paste_text_only() {		
	if WinActive("ahk_exe WINWORD.EXE") {
		send_by_clipboard(A_Clipboard)
	}	
}


send_text_by_clipboard(text_to_send) {
	clip_saved := ClipboardAll()  
	A_Clipboard := text_to_send
	Sleep(40)
	SendInput("^v")	
	Sleep(40)
	A_Clipboard := clip_saved  
}

has_text_selected() {
	clip_saved := ClipboardAll()  
	A_Clipboard := ""
	Sleep(40)
	SendInput("^c")	
	Sleep(40)
	if (A_Clipboard != "") {
        return true
    } else {
        return false
    }    
	A_Clipboard := clip_saved  
}

get_selected_text_by_clipboard() {	
	clip_saved := ClipboardAll() 
	A_Clipboard := ""
	SendInput("^c")			
	sleep(200)
	selected_text := A_Clipboard
	A_Clipboard := clip_saved  
	return selected_text 
}

get_via_clipboard() {	
    try {
        clip_saved := ClipboardAll() 
		A_Clipboard := ""
		SendInput("^c")			
		sleep(200)
		str := A_Clipboard
		A_Clipboard := clip_saved  
		return str 
    }
    catch as err
        Msgbox(err.Message)   
}

quote(p) {
	return("`"" . p . "`"")	
}


strToAppLink(str)
{
	link := "D:\coreFiles\appLink\" . str . ".lnk"
	return link
}





	


; #endregion


; #region smart shortcut



  


left_window() {	
    ; MouseClick("left")
    sleep(200)
    count := 0
    window_list := WinGetList() 
    for hWnd in window_list {
        try {        
            style := WinGetStyle(hWnd)
            if (style & 0x40000)  {
                count++
            }
        }
        catch {
            continue
        }       
    }
    cmd := count > 2 ? "41" : (count == 1 ? "11" : "31")
    SendInput("#z")
    sleep(300)
    SendInput(cmd)

    If !WinActive("ahk_exe Weixin.exe") {
        sleep(300)
        SendInput("{Escape}")
    }        
}

right_window() {	
    MouseClick("left")
    sleep(200)
    count := 0
    windowList := WinGetList() 

    for hWnd in windowList {
        try {        
            Style := WinGetStyle(hWnd)
            if (Style & 0x40000)  {
                count++
            }
        }
        catch {
            continue
        }       
    }
    cmd := count > 2 ? "42" : (count == 1 ? "12" : "32")
    
    SendInput("#z")
    sleep(200)
    SendInput(cmd)
    sleep(200)
    SendInput("{Escape}")
}




; #endregion



; #region 转换RGB颜色值

rgb2bgr(color, g := -1, b := -1) {    
    try {
        ; 检查是否提供了三个单独的RGB分量
        if (g == -1 || b == -1) {             
            for color_name in system_color_map
                if StrTitle(color) == color_name
                    color := system_color_map[color_name]

            r := (color >> 16) & 0xFF
            g := (color >> 8) & 0xFF
            b := color & 0xFF
        }             
        ; 将RGB分量转换为BGR数值
        return (r << 0) | (g << 8) | (b << 16)    
    } catch
        throw
}



bgr2rgb(color, g := -1, b := -1) {    
    try {
        b := (color >> 16) & 0xFF
        g := (color >> 8) & 0xFF
        r := color & 0xFF
        ; 将BGR分量转换为RGB数值
        return (r << 16) | (g << 8) | (b << 0)
    } catch
        throw
}


; 辅助函数：将值限制在指定范围内
Clamp(value, min, max) {
    if (value < min)
        return min
    if (value > max)
        return max
    return value
}

; #endregion

find_file() { 
    
    Loop Files, "C:\Program Files\WindowsApps\*.*", "R"  
    {
        if (A_LoopFileName ~= "i)^(snipaste\.exe)$") {
            A_Clipboard := A_LoopFilePath
            MsgBox(A_LoopFilePath)
        }
        ; Result := MsgBox("Filename = " A_LoopFilePath "`n`nContinue?",, "y/n")
        ; if Result = "No"
        ;     break
    }
}

write_json_file() { 
    ; 定义文件路径（请替换为你的实际文件路径）
    filePath := "settings\settings.json"  ; <--- 修改这里

    ; 1. 读取整个文件内容
    content := FileRead("settings\settings.json") 

    ; 2. 定义查找的旧字符串和替换的新字符串
    ; 注意：在 AHK 字符串中，反斜杠 \ 需要写成 \\，双引号 " 需要写成 ""
    oldLine := "`"snipaste`": `"D:\\system\\Snipaste-2.11-x64\\Snipaste.exe`""
    newLine := "`"snipaste`": `"C:\\Program Files\\WindowsApps\\45479liulios.17062D84F7C46_2.11.300.0_x64__p7pnf6hceqser\\snipaste.exe`""

    ; 3. 执行替换
    ; 使用正则表达式模式 "m)" 表示多行模式，确保能匹配换行符
    ; 如果不确定旧路径的具体版本，可以只匹配 "snipaste" 开头的部分
    ; 这里使用简单的替换
    newContent := StrReplace(content, oldLine, newLine)

    ; 4. 如果没找到旧内容，StrReplace 会返回原内容，不会报错
    ; 你可以加个判断确保修改生效（可选）
    if (newContent != content) {
        MsgBox("成功找到并替换了 snipaste 路径。")
    } else {
        MsgBox("警告：未找到旧的 snipaste 路径。可能路径已更新或格式不同。")
    }

        
    ; 4. 【关键步骤】使用 FileOpen 以写入模式打开文件，这会清空原文件内容
    fileObj := FileOpen(filePath, "w") ; "w" 模式表示写入，会覆盖原文件
    fileObj.Write(newContent)
    fileObj.Close() ; 记得关闭文件句柄
}

snipaste() {	
    try {
       app_snipaste := run_config["app_snipaste"]["payload"] 
       command := app_snipaste . " snip --area 1 1 500 500 -o pin"
       run(command)
    }
    catch as err
        Msgbox(err.Message)   
}

; 交互式区域截图函数：区域选择 -> 保存文件 -> 复制剪贴板 -> Snipaste显示
screenshot(if_stick := True, if_show := True, back_color := "Purple") {    
    save_path := run_config["path_screenshot_path"]["payload"]

    token := 0    ; GDI+ token
    bitmap := 0   ; 位图对象
    try {
        ; 1. 交互式选择区域
        result := select_region_interactive(back_color)        
        if (!result.region) {
            return {success: false, file_path: "", message: "用户取消了区域选择"}
        }
        region := StrReplace(result.region, "|", " ")
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


       app_snipaste := run_config["app_snipaste"]["payload"] 
        command := APP_SNIPASTE . " snip --area " . region . " -o clipboard;pin"
        run(command)
      
        ; ; 4. 启动GDI+
        ; if (!(token := Gdip_Startup())) {
        ;     return {success: false, file_path: "", message: "GDI+ 启动失败，请确保您的系统中安装了 GDI+"}
        ; }          
   

        ; ; 5. 捕获指定区域
        ; bitmap := Gdip_BitmapFromScreen(result.region)
        ; if (bitmap == -1) {
        ;     return {success: false, file_path: "", message: "无法捕获指定区域"}
        ; }
        
        ; ; 6. 保存图片到文件
        ; save_result := Gdip_SaveBitmapToFile(bitmap, file_path)
        ; if (save_result != 0) {  ; 0 = 成功，非零 = 失败
        ;     error_messages := Map(
        ;         -1, "不支持的文件格式",
        ;         -2, "无法获取编码器列表", 
        ;         -3, "找不到匹配的编码器",
        ;         -4, "无法获取输出文件的宽字符名称",
        ;         -5, "无法保存文件到磁盘"
        ;     )
        ;     error_msg := error_messages.Has(save_result) ? error_messages[save_result] : "未知错误(" . save_result . ")"
        ;     return {success: false, file_path: "", message: "保存图片失败: " . error_msg}
        ; }
        
        ; ; 7. 复制到剪贴板（此函数没有返回值） 
        ; try {
        ;     Gdip_SetBitmapToClipboard(bitmap)
        ; } catch as err {
        ;     ; 保存成功但复制失败，仍然算部分成功
        ;     return {success: true, file_path: file_path, message: "截图已保存，但复制到剪贴板失败: " . err.message}
        ; }                 
        
        ; ; 8. 使用Snipaste贴在屏幕上
        ; if(if_stick) {
        ;     try {
        ;         Sleep(50)
        ;         SendInput("{F3}")
        ;         Sleep(200) 
        ;         SendInput("{Space}")
        ;         Sleep(10) 
        ;         SendInput("b")

        ;     } catch {
        ;         ; Snipaste可能未安装，忽略错误
        ;     }
        ; }
        
        ; ; 9. 显示成功提示
        ; if if_show {
        ;     Notify.show("截图已保存并复制到剪贴板")
        ; }       
                
        return {
            success: true           
        }
        
    } catch as err {
        error_msg := "截图过程中发生错误: " . err.message
        Msgbox(error_msg)
        return {success: false, file_path: "", message: error_msg}
        
    ; } finally {
    ;     ; 确保资源释放
    ;     if (bitmap) {
    ;         Gdip_DisposeImage(bitmap)
    ;     }
    ;     if (token) {
    ;         Gdip_Shutdown(token)
    ;     }
    }
}



; #region ========================= 区域选择和捕获功能 =========================

/**
 * 交互式区域选择
 * 
 * @description 立即显示选择框，实时跟踪鼠标拖拽
 * @param {String} button 触发按键（默认："RButton"）
 * @return {Object} 
 */
get_screen_region(button := "RButton") {  
    transparent := 100
    back_color := "Purple"
    x := y := w := h := 0        ; 初始化区域变量
    selection_gui := 0                       ; GUI对象
    info_gui := 0                ; 信息显示GUI
    
    result := {success: true, x: 0, y: 0, width: 0, height: 0}   
    
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
                return 0
            }
            
            Sleep(10)
            MouseGetPos(&end_x, &end_y)
            
            ; 计算选择区域
            x := (begin_x < end_x) ? begin_x : end_x
            y := (begin_y < end_y) ? begin_y : end_y
            w := Abs(begin_x - end_x)            
            h := Abs(begin_y - end_y)

            ; 移动选择框
            selection_gui.Move(x, y, w, h)
            
            ; 更新信息显示
            info_text.Text := Format("区域: {}x{} | 位置: ({}, {})", w, h, x, y)
            info_gui.Move(end_x + 15, end_y - 40)
        }
        
        ; 检查区域大小并返回结果
        if (w < 5 || h < 5) {
            return {success: false, message: "区域选择太小"}
        } else {
            return {success: true, x: x, y: y, w: w, h: h}
        }
        
    } catch as err {        
        return {success: false, message: "区域选择错误: " . err.Message}
    } finally {
        ; 确保GUI被销毁
        if (selection_gui)
            selection_gui.Destroy()
        if (info_gui) 
            info_gui.Destroy()
    }
}

get_screen_region_lbutton() {
    return get_screen_region(button := "LButton") 
}

test_get_screen_region(button := "RButton") {  
    region := get_screen_region()
    Notify.show("已选择区域: " . "`n" . "x: " . region.x . "`n" . "y: " . region.y . "`n" . "width: " . region.width . "`n" . "height: " . region.height )

}
; #endregion





; xxxxxxxxxxxxxxxxxxxxxxxxss



; #region color



; =========================================
; 颜色工具
; 内部统一使用字符串 "0xRRGGBB"
; =========================================

; 将各种输入规范化为 "0xRRGGBB"
normalize_color(color_input) {
    ; 已经是字符串且形如 0xRRGGBB / 0xAARRGGBB
    if (Type(color_input) = "String") {
        s := Trim(color_input)

        ; 1) 0x 前缀
        if (RegExMatch(s, "i)^0x[0-9a-f]{6}$")) {
            return "0x" . StrUpper(SubStr(s, 3))
        }

        ; 2) 0xAARRGGBB -> 取后 6 位 RGB
        if (RegExMatch(s, "i)^0x[0-9a-f]{8}$")) {
            return "0x" . StrUpper(SubStr(s, 5))
        }

        ; 3) #RRGGBB
        if (RegExMatch(s, "i)^#[0-9a-f]{6}$")) {
            return "0x" . StrUpper(SubStr(s, 2))
        }

        ; 4) 纯 6 位十六进制字符串
        if (RegExMatch(s, "i)^[0-9a-f]{6}$")) {
            return "0x" . StrUpper(s)
        }

        ; 5) 颜色名
        hex := color_name_to_hex(s)
        if (hex != "")
            return hex

        throw Error("无法识别的颜色字符串: " . color_input)
    }

    ; 数字类型：可能是十进制整数，或 AHK 的 0x 数字字面量
    if (Type(color_input) = "Integer" || Type(color_input) = "Float") {
        num := Integer(color_input)
        if (num < 0 || num > 0xFFFFFFFF)
            throw Error("颜色数值超出范围: " . color_input)

        ; 如果是 0xAARRGGBB，取后 24 位
        rgb := num & 0xFFFFFF
        return Format("0x{:06X}", rgb)
    }

    throw Error("不支持的颜色类型: " . Type(color_input))
}

; 将颜色转为整数 0xRRGGBB
color_to_int(color_input) {
    color_str := normalize_color(color_input)
    return Integer(color_str)
}

; 将颜色转为字符串 "0xRRGGBB"
color_to_str(color_input) {
    return normalize_color(color_input)
}

; 判断是否深色
is_dark_color(color_input) {
    color_int := color_to_int(color_input)

    r := (color_int >> 16) & 0xFF
    g := (color_int >> 8) & 0xFF
    b := color_int & 0xFF

    brightness := r * 0.299 + g * 0.587 + b * 0.114
    return brightness < 128
}

; 获取对比色：深色返回白色，浅色返回黑色
get_contrast_color(color_input) {
    return is_dark_color(color_input) ? "0xFFFFFF" : "0x000000"
}

; =========================================
; 颜色名称 -> hex
; 返回字符串 "0xRRGGBB"
; =========================================
color_name_to_hex(color_name) {
    name := StrLower(Trim(color_name))

    colors := Map(
        "black",   "0x000000",    ; 黑色
        "white",   "0xFFFFFF",    ; 白色
        "red",     "0xFF0000",    ; 红色
        "green",   "0x008000",    ; 绿色
        "blue",    "0x0000FF",    ; 蓝色
        "yellow",  "0xFFFF00",    ; 黄色
        "gray",    "0x808080",    ; 灰色
        "grey",      "0x808080",    ; 灰色
        "silver",  "0xC0C0C0",    ; 银色
        "maroon",  "0x800000",    ; 栗色
        "olive",   "0x808000",    ; 橄榄色
        "lime",    "0x00FF00",    ; 酸橙绿
        "aqua",    "0x00FFFF",    ; 浅青色
        "cyan",      "0x00FFFF",    ; 青色
        "teal",    "0x008080",    ; 青绿色
        "navy",    "0x000080",    ; 海军蓝
        "fuchsia", "0xFF00FF",    ; 紫红色
        "magenta",   "0xFF00FF",    ; 品红色
        "purple",  "0x800080",    ; 紫色
        "orange",    "0xFFA500",    ; 橙色
        "pink",      "0xFFC0CB",    ; 粉色
        "brown",     "0xA52A2A",    ; 棕色
        "gold",      "0xFFD700",    ; 金色
        "indigo",    "0x4B0082",    ; 靛蓝
        "violet",    "0xEE82EE"     ; 紫罗兰色
    )

    return colors.Has(name) ? colors[name] : ""
}

; #endregion



; end--------------------------