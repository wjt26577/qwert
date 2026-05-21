; #region 启动程序
#SingleInstance Force
#Requires AutoHotkey v2.0

try DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")

#Include <global_vars>
#Include <json>
#Include <load_json>
#Include <Notify>
#Include <IMECtrl>
#Include <step_clip>  
#Include <lib_office_function>
#Include <baidu_api>
#Include <class_http>
#Include <Gdip_All>
#Include <split_ppt>

#Include <lib_core>   

#Include <lib_window>   
#Include <lib_mouse>    
#Include <lib_string>   
#Include <lib_clipboard>
#Include <lib_file>     
#Include <lib_ppt>      
#Include <lib_system>   
#Include <lib_color>    
#Include <lib_screen>  




show_startup_info()
init_screen_bounds()
load_clipboard_history()
OnExit(save_clipboard_history)


show_startup_info() {
    icon_path := A_IsAdmin ? "img\chili.ico" : "img\cherries.ico"

    if FileExist(icon_path)
        TraySetIcon(icon_path, , 0)
    else
        MsgBox("找不到托盘图标: " . icon_path, "警告", "T2")

    try {
        if A_IsAdmin
            Notify.show("Qwert is running as admin", "error", 500)
        else
            Notify.show("Qwert is running as normal", "success", 500)

    } catch  {     
        if A_IsAdmin
            TrayTip("Qwert is running as admin", "T2", 500)
        else
            TrayTip("Qwert is running as normal", "T2", 500)
    }
}

init_screen_bounds() {
    global screen_left, screen_right, screen_top, screen_bottom

    try DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")

    ; SM_XVIRTUALSCREEN  = 76 虚拟屏幕左边
    ; SM_YVIRTUALSCREEN  = 77 虚拟屏幕顶部
    ; SM_CXVIRTUALSCREEN = 78 虚拟屏幕总宽度
    ; SM_CYVIRTUALSCREEN = 79 虚拟屏幕总高度

    vx := DllCall("GetSystemMetrics", "int", 76)
    vy := DllCall("GetSystemMetrics", "int", 77)
    vw := DllCall("GetSystemMetrics", "int", 78)
    vh := DllCall("GetSystemMetrics", "int", 79)

    screen_left   := vx
    screen_top    := vy
    screen_right  := vx + vw
    screen_bottom := vy + vh
}

load_clipboard_history() {
    global clipboard_history, max_history, clipboard_history_FILE

    try {

        if !FileExist(clipboard_history_FILE) {
            loop max_history
                clipboard_history.Push("#" . A_Index . "_clip is empty")
            return
        }

        clipboard_history_map := json.load(FileRead(clipboard_history_FILE))

        loop max_history {
            idx := String(A_Index)
            if clipboard_history_map.Has(idx)
                clipboard_history.Push(clipboard_history_map[idx])
            else
                clipboard_history.Push("#" . A_Index . "_clip is empty")
        }

    } catch as err {
        MsgBox(Format("剪贴板历史加载失败:`n{1}", err.message))
        loop max_history
            clipboard_history.Push("#" . A_Index . "_clip is empty")
    }
}


; #endregion
 
; #region suspend_exempt
#SuspendExempt

~LButton & 3:: reload_me()
~LButton & 4:: toggle_suspend()

toggle_suspend() {
    Suspend
    if A_IsSuspended
        Notify.show("QWERT is suspended", "success", 500)
    else
        Notify.show("QWERT is resumed", "info", 500)
}

reload_me() {
    save_before_reload()
    Reload
}

exit_me() {
    Notify.show("QWERT is exiting...", "error", 500)
    sleep 1000
    ExitApp
} 

save_before_reload() {
    SendInput("{Escape}")
    if WinActive("ahk_exe Code.exe") {
        SendInput("^s")
        Sleep(50)
    }
}

#SuspendExempt False
; #endregion

; #region 零碎

#HotIf step_clip_active()
    XButton1 & f:: step_clip_back()
#HotIf


#HotIf WinActive("ahk_class XLMAIN")  
    XButton1::SendInput("+{LButton}")
    Enter:: SendInput("!{Enter}") 
#HotIf

#HotIf WinActive("ahk_exe msedge.exe")   
   
#HotIf

#HotIf WinActive("ahk_class OpusApp")   
    XButton1::SendInput("+{LButton}")
#HotIf

#HotIf WinExist("Notify ahk_class AutoHotkeyGUI")
    Esc::Notify.hide_all()
#HotIf  


#HotIf WinActive("ahk_exe WeMeetApp.exe")
    XButton1:: SendInput("!m")  ; 静音/取消静音
#HotIf

#HotIf WinActive("确认另存为")
    XButton1:: SendInput("!y") 
    RButton:: SendInput("!n") 
#HotIf

#HotIf WinActive("ahk_class NUIDialog")
    XButton1:: SendInput("!y") 
    RButton:: SendInput("!n") 
#HotIf


#HotIf WinActive("ahk_class CabinetWClass")
    RButton & b:: CycAct.slow("zip", "unzip")
    ; 压缩解压操作
    ; xMButto & a:: run_zip()                                          ; MButto+A - 压缩文件
    ; xMButto & s:: run_unzip()                                        ; MButto+S - 解压文件
    
    ; ; 文件处理操作
    ; ; MButto & d:: process_font_914()                  ; MButto+D - 处理字体
    ; ; MButto & w:: split_in_folder()                   ; MButto+W - 文件夹分割
    ; ; MButto & e:: export_jpg_in_folder()              ; MButto+E - 导出JPG
    ; xMButto & r:: append_datetime()                   ; MButto+R - 添加时间戳
    
    ; ; 文件组织操作    MButto & q:: move_files_to_first_same_named_folders()  ; MButto+Q - 移动到首个同名文件夹
    ; ; MButto & z:: move_files_to_same_named_folders()        ; MButto+Z - 移动到同名文件夹
    ; xMButto & x:: same_folder()                             ; MButto+X - 同文件夹操作
    ; ; xMButto & v:: batch_rename_folder()                     ; MButto+V - 批量重命名文件夹
    ; xMButto & t:: same_folder()                                            ; MButto+T - 同文件夹操作(直接调用)
    ; xMButto & b:: process_powerpoint()                                            ; MButton+T - 同文件夹操作(直接调用)
    
    ; ; 测试功能

    ; ; xMButto & c:: split_ppt_by_page_range()
    
    ; xMButto & 1:: process_ppt_full_process() 
    ; xMButto & 2:: process_ppt_split_only()
    ; xMButto & 3:: process_ppt_export_only()
    ; xMButto & 4:: process_ppt_organize_only()
    ; xMButto & 4:: process_powerpoint_export_and_organize() 
    ; MButto & 2:: split_ppt_full_with_progress("")
    ; MButto & 3:: split_ppt_to_single_main()  ; 50为每批大小，可自定义
#HotIf






; 微信; [OK][抱拳][强][握手][呲牙][抠鼻][憨笑][捂脸][坏笑][流汗]
#HotIf (WinActive("ahk_exe Weixin.exe"))
    ; XButton1 & 2::  ("text | 收到`n | 收  到", "text | [OK]`n | 👌", "text | [强]`n | 👍", "text | [抱拳]`n | 抱  拳")
    ; XButton1 & 3::  ("text | [捂脸]`n | 捂  脸", "text | [呲牙]`n | 呲  牙", "text | [抠鼻]`n | 抠  鼻")  
    
    ; 没问题，可以的。好的，知道了。好的，没问题。
; F23 & 2:: ("text | 收到啦，多谢!`n | 收  到", "text | [OK]`n | 👌")
; F23 & 4:: ("text | [强]`n | 👍", "text | [抱拳]`n | 抱  拳")
; F23 & q:: show_thishotkey()
; F23 & w:: ("text | 收到`n | 收  到", "text | [OK]`n | 👌", "text | [强]`n | 👍", "text | [抱拳]`n | 抱  拳")
; F23 & r:: ("text | 没问题，可以的。`n | 没问题，可以的。", "text | [OK]`n | 👌", "text | [强]`n | 👍", "text | [抱拳]`n | 抱  拳")
; F23 & t:: ("text | 好的，知道了。`n | 好的，知道了。", "inpt | 好的，没问题。`n | 好的，没问题。", "text | 谢了啊`n | 谢了啊", "text | [抱拳]`n | 抱  拳")
; F23 & a:: show_thishotkey()
; F23 & g:: show_thishotkey()
; F23 & z:: show_thishotkey()
; F23 & x:: show_thishotkey()
; F23 & c:: show_thishotkey()
; F23 & b:: show_thishotkey()

#HotIf



#HotIf WinActive("ahk_class TTOTAL_CMD")    
    F24 & 1:: process_ppt_full_process() 
    ; F24 & 2:: process_ppt_split_only()
    F24 & 3:: process_ppt_export_only()
    F24 & 4:: process_ppt_organize_only()
    ; ; F24 & 3:: process_powerpoint_export_and_organize() 
    ; F24 & Tab:: split_ppt_to_single() 
    ; F24 & 3:: split_ppt_to_single_main("", 50)  ; 50为每批大小，可自定义
#HotIf

#HotIf WinActive("ahk_class Notepad")    
    ; F24 & 1:: process_ppt_full_process() 
    ; F24 & 2:: process_ppt_split_only()
    ; F24 & 3:: process_ppt_export_only()
    ; F24 & 4:: process_ppt_organize_only()
    ; F24 & 3:: process_powerpoint_export_and_organize() 
    ; F24 & Tab:: split_ppt_to_single() 
    ; F24 & 3:: split_ppt_to_single_main("", 50)  ; 50为每批大小，可自定义
#HotIf



#HotIf WinActive("ahk_class PotPlayer64")
    F24:: SendInput(",") 
    F23:: SendInput(".") 
    XButton1:: SendInput("{Right}") 
    XButton2:: SendInput("{Left}")
#HotIf


; ========================================================================
; AutoCAD
; ========================================================================
#HotIf (WinActive("ahk_exe acad.exe") && is_text_mode())
    
    ; 批量复制功能
    XButton1 & f:: {
        save_clipboard_to_temp_file()
        SendText("BatchCopyText`n")
        ; SendInput("BatchCopyText{Enter}")
    }

    ; 批量倒序复制功能
    XButton1 & e:: {
        save_reversed_clipboard_lines_to_file()
        SendText("BatchCopyText`n")
    }
    XButton1 & b:: SendInput("^0>>")
    ; 全  屏


    ; w:: ("{Esc}>> | Escape")
    ; e:: ("move`n | 移动") 
    ; r:: ("{F8} | 正交限制光标 开/关")	
    ; t:: ("SCALE`n | 缩  放", "inpt | FILLET`n | 圆  角")
    ; s:: ("{F8} | 正交限制光标 开/关")
    ; g:: ("copy`n | 带基点复制")
    ; b:: ("END`n | 端  点", "inpt | MID`n | 中点", "inpt | INT`n | 交点")
    
    d:: {
        MouseClick("left")
        SendText("MTEDIT`n")
        IMECtrl.to_upper()
        Notify.show("English UPPERCASE", "error", 500)

        Sleep(500)
        
        SendInput("^a")
    }
    ; b:: {  ; B - 插入块
    ; }


    
save_clipboard_to_temp_file() {
    file_path := "D:\temp\temp.txt"
    if !DirExist("D:\temp") {
        DirCreate "D:\temp"
    }	
    if FileExist(file_path) {
        FileDelete file_path
    }
    FileAppend(A_Clipboard, file_path)
}

save_reversed_clipboard_lines_to_file() {
    file_path := "D:\temp\temp.txt"
    if !DirExist("D:\temp")
        DirCreate("D:\temp")
    if FileExist(file_path)
        FileDelete(file_path)

    ; 👇 关键：去除首尾空白（包括末尾换行）
    trimmed_clip := Trim(A_Clipboard, "`r`n `t")
    
    ; 如果剪贴板全是空白，避免写入空文件或错误
    if (trimmed_clip = "") {
        FileAppend("", file_path)
        return
    }

    lines := StrSplit(trimmed_clip, "`n", "`r")
    reversed_lines := []
    for line in lines
        reversed_lines.InsertAt(1, line)

    reversed_text := StrJoin(reversed_lines, "`n")
    FileAppend(reversed_text, file_path)
}


cad_copy() {
    SendText("copy`n") 
}

cad_move() {
    SendText("move`n") 	
}

cad_insert_block() {	
    SendText("-i`npbi6000`n")
}

cad_insert_jump() {	
    SendInput("insertjump{Enter}")
}


; --- 工具函数 ---
StrJoin(arr, delimiter) {
    result := ""
    for index, value in arr {
        if (index > 1)
            result .= delimiter
        result .= value
    }
    return result
}


#HotIf





#HotIf WinActive("ahk_exe Eagle.exe") && is_text_mode()
    
    ; 从eagle中复制PPT页面
    RButton & f:: {
        add_slide_from_eagle()     
    }

#HotIf

#HotIf WinActive("ahk_exe Code.exe")
    >+LButton:: SendInput("!+{LButton}")
    ~LButton & g:: CycAct.fast("^k^]", "^k^0")
    ; ~LButton & f:: CycAct.fast("+{End}^c", "+{End}^x")
    ; ~LButton & v:: CycAct.fast("+{End}^c", "+{End}^x")


    add_brace() {
        try {
            SendInput(" {{}{Down}{End}`n{`}}")
        } catch as err {
            MsgBox("增加大括号失败: " . err.Message)
        }
    }
  
    ; F24 & f:: SendInput("^c^r{Right}{Enter}^v{Enter}") 
       
    
    


    ; XButton1:: fast_cat("+{MButton}", "+{LButton}")

    ; F24:: bb


   

    :*:bbb:: {
        try {
            SendInput(" {{}{Delete}{Down}{End}`n+{Tab}{`}}")
        } catch as err {
            MsgBox("增加大括号失败: " . err.Message)
        }
    }

    
    ; :*:bbb:: {
    ;     try {
    ;         SendInput(" {{}{Delete}{Down}{End}`n+{Tab}{`}}")
    ;     } catch as err {
    ;         MsgBox("增加大括号失败: " . err.Message)
    ;     }
    ; }




    ; :*:bbb:: {
    ;     try {
    ;         SendInput(" {{}{Delete}{Down}{End}`n+{Tab}{`}}")
    ;     } catch as err {
    ;         MsgBox("增加大括号失败: " . err.Message)
    ;     }
    ; }



    :*:eee::{Home 2}+{End}^v
    :*:sss::+{End}^c
    :*:ddd::+{End}^v
    :*:ssg:: {
        SendText("`nMsgBox(A_Now, , `"T2`")") 
            SendInput("{Left 9}+{Left 5}")
    }
    
    :*:ssf:: {
        text := "
        (
        function() {
            try {              
            } catch as err {
                MsgBox("失败: " . err.Message)    
            }
        }
        )"    
        send_by_clipboard(text)
        SendInput("{Left 1}{Up 5}+{Right 8}")

    }

    Tab & j:: SendInput("^{Right 2}^+{Left}")

    ; F23:: {
    ;     SendInput("!+{LButton}")
    ; } 

    ; F23 & r:: {
    ;     SendInput("^!+l")
    ; }  

    ; F23 & g:: {
    ;     SendInput("^!+u")
    ; }   

 
; smart_delete() {  
;     try {
;         if WinActive("ahk_exe Code.exe") {
;             ("inpt | ^+k | 删除整行", "inpt | ^+!2 | 删除光标右侧内容", "inpt | ^+!1 | 删除光标左侧内容")    
;         } else {
;             ("inpt | {Home}+{End}{Delete} | 删除整行", "inpt | +{End}{Delete} | 删除光标右侧内容", "inpt | +{Home}{Delete} | 删除光标左侧内容")    
;         }
;     } catch as err
;         Msgbox(err.Message)       
; } 
 
#HotIf


; #endregion

; #region not text_mode && ppt 
#HotIf !is_text_mode() && WinActive("ahk_class PPTFrameClass") 



q:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("q")    : CycAct.slow("url_zhyw")                 ) : CycAct.slow("url_zhyw")                
w:: !mouse_top() ? (!mouse_edge() ? CycAct.fast("ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart") : CycAct.fast("AlignLeft", "AlignCenter", "AlignRight")) : CycAct.slow("app_wechat")              
e:: !mouse_top() ? (!mouse_edge() ? toggle_hide_shape_range()  : CycAct.slow("app_everything")           ) : CycAct.slow("app_everything")          
r:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("r")    : CycAct.slow("url_zhyw")                 ) : CycAct.slow("url_zhyw")                
t:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("t")    : CycAct.slow("url_zhyw")                 ) : CycAct.slow("url_zhyw")                
a:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("a")    : CycAct.slow("url_zhyw")                 ) : CycAct.slow("url_zhyw")                
s:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("s")    : CycAct.slow("app_vscode")               ) : CycAct.slow("app_vscode")              
d:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("d")    : CycAct.slow("open_current_path")        ) : CycAct.slow("open_current_path")       
f:: !mouse_top() ? (!mouse_edge() ? copy_or_cut_by_count()     : CycAct.slow("open_current_path")        ) : CycAct.slow("open_current_path")       
g:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("g")    : CycAct.slow("app_notepad")              ) : CycAct.slow("app_notepad")             
z:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("z")    : CycAct.slow("app_powershell_terminal")           ) : CycAct.slow("app_powershell")          
x:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("x")    : CycAct.slow("app_word")                 ) : CycAct.slow("app_word")                
c:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("c")    : CycAct.slow("close_window_in_taskbar")  ) : CycAct.slow("close_window_in_taskbar") 
v:: !mouse_top() ? (!mouse_edge() ? SendInput("^v")            : CycAct.slow("url_bilibili")             ) : CycAct.slow("url_bilibili")            
b:: !mouse_top() ? (!mouse_edge() ? send_with_capslock("b")    : CycAct.slow("^!+1")                     ) : CycAct.slow("^!+2")                    

~y:: enter_text_mode()
~u:: enter_text_mode()
~i:: enter_text_mode()
~o:: enter_text_mode()
~p:: enter_text_mode()
~h:: enter_text_mode()
~j:: enter_text_mode()
~k:: enter_text_mode()
~l:: enter_text_mode()
~m:: enter_text_mode()
~n:: enter_text_mode()

q & LButton:: CycAct.slow("show_thishotkey")
w & LButton:: CycAct.slow("show_thishotkey")
e & LButton:: CycAct.slow("show_thishotkey")
r & LButton:: CycAct.slow("show_thishotkey")
t & LButton:: CycAct.slow("show_thishotkey")
a & LButton:: CycAct.slow("show_thishotkey")
s & LButton:: CycAct.slow("show_thishotkey")
d & LButton:: CycAct.slow("show_thishotkey")
g & LButton:: CycAct.slow("show_thishotkey")
z & LButton:: CycAct.slow("show_thishotkey")
x & LButton:: CycAct.slow("show_thishotkey")
c & LButton:: CycAct.slow("show_thishotkey")
b & LButton:: CycAct.slow("show_thishotkey")

q & F23:: CycAct.slow("show_thishotkey")
w & F23:: CycAct.slow("show_thishotkey")
e & F23:: CycAct.slow("show_thishotkey")
r & F23:: CycAct.slow("show_thishotkey")
t & F23:: CycAct.slow("show_thishotkey")
a & F23:: CycAct.slow("show_thishotkey")
s & F23:: CycAct.slow("show_thishotkey")
d & F23:: CycAct.slow("show_thishotkey")
g & F23:: CycAct.slow("show_thishotkey")
z & F23:: CycAct.slow("show_thishotkey")
x & F23:: CycAct.slow("show_thishotkey")
c & F23:: CycAct.slow("show_thishotkey")
b & F23:: CycAct.slow("show_thishotkey")

q & RButton:: CycAct.slow("show_thishotkey")
w & RButton:: CycAct.slow("AlignDistributeVertically")
e & RButton:: CycAct.slow("show_thishotkey")
r & RButton:: CycAct.slow("show_thishotkey")
t & RButton:: CycAct.slow("show_thishotkey")
a & RButton:: CycAct.slow("show_thishotkey")
s & RButton:: CycAct.slow("show_thishotkey")
d & RButton:: CycAct.slow("show_thishotkey")
g & RButton:: CycAct.slow("show_thishotkey")
z & RButton:: CycAct.slow("show_thishotkey")
x & RButton:: CycAct.slow("show_thishotkey")
c & RButton:: CycAct.slow("show_thishotkey")
b & RButton:: CycAct.slow("resize_by_guides")

q & F24:: CycAct.slow("show_thishotkey")
w & F24:: CycAct.slow("show_thishotkey")
e & F24:: CycAct.slow("show_thishotkey")
r & F24:: CycAct.slow("show_thishotkey")
t & F24:: CycAct.slow("show_thishotkey")
a & F24:: CycAct.slow("show_thishotkey")
s & F24:: CycAct.slow("show_thishotkey")
d & F24:: CycAct.slow("show_thishotkey")
g & F24:: CycAct.slow("show_thishotkey")
z & F24:: CycAct.slow("show_thishotkey")
x & F24:: CycAct.slow("show_thishotkey")
c & F24:: CycAct.slow("show_thishotkey")
b & F24:: CycAct.slow("show_thishotkey")

q & XButton1:: CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
w & XButton1:: CycAct.fast("ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart")
e & XButton1:: CycAct.slow("toggle_hide_shape_range")
r & XButton1:: CycAct.fast("AlignLeft", "AlignCenter", "AlignRight")
t & XButton1:: CycAct.slow("show_thishotkey")
a & XButton1:: CycAct.slow("show_thishotkey")
s & XButton1:: CycAct.slow("show_thishotkey")
d & XButton1:: CycAct.slow("show_thishotkey")
g & XButton1:: CycAct.slow("show_thishotkey")
z & XButton1:: CycAct.slow("show_thishotkey")
x & XButton1:: CycAct.slow("show_thishotkey")
c & XButton1:: CycAct.slow("show_thishotkey")
b & XButton1:: CycAct.slow("show_thishotkey")

q & XButton2:: CycAct.slow("show_thishotkey")
w & XButton2:: CycAct.slow("show_thishotkey")
e & XButton2:: CycAct.slow("show_thishotkey")
r & XButton2:: CycAct.slow("show_thishotkey")
t & XButton2:: CycAct.slow("show_thishotkey")
a & XButton2:: CycAct.slow("show_thishotkey")
s & XButton2:: CycAct.slow("show_thishotkey")
d & XButton2:: CycAct.slow("show_thishotkey")
g & XButton2:: CycAct.slow("show_thishotkey")
z & XButton2:: CycAct.slow("show_thishotkey")
x & XButton2:: CycAct.slow("show_thishotkey")
c & XButton2:: CycAct.slow("show_thishotkey")
b & XButton2:: CycAct.slow("show_thishotkey")

#HotIf
; #endregion
; #region not text_mode && not ppt
#HotIf !WinActive("ahk_class PPTFrameClass")   

2:: mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")                         : (mouse_edge() ? CycAct.slow("show_thishotkey")            : (SendInput("{Delete}")))            
q:: mouse_top() ? CycAct.slow("show_thishotkey", "show_thishotkey", "show_thishotkey")          : (mouse_edge() ? CycAct.slow("show_thishotkey")            : (enter_text_mode(), send_with_capslock("q")))            
w:: mouse_top() ? CycAct.slow("app_wechat", "app_wemeet", "app_wxwork")                         : (mouse_edge() ? CycAct.slow("app_wechat")                 : (enter_text_mode(), send_with_capslock("w")))
e:: mouse_top() ? CycAct.slow("app_everything", "app_eagle", "app_evernote")                    : (mouse_edge() ? CycAct.slow("{PgUp}")                     : (enter_text_mode(), send_with_capslock("e")))
r:: mouse_top() ? CycAct.slow("app_obsidian", "app_typora", "show_thishotkey")                  : (mouse_edge() ? CycAct.slow("app_obsidian")               : (enter_text_mode(), send_with_capslock("r")))
t:: mouse_top() ? CycAct.slow("show_time")                                                      : (mouse_edge() ? CycAct.slow("show_time")            : (enter_text_mode(), send_with_capslock("t")))
a:: mouse_top() ? CycAct.slow("app_powershell_terminal", "show_thishotkey", "show_thishotkey")  : (mouse_edge() ? CycAct.slow("show_thishotkey")            : (enter_text_mode(), send_with_capslock("a")))
s:: mouse_top() ? CycAct.slow("app_vscode", "app_spy", "file_shortcut")                         : (mouse_edge() ? CycAct.slow("app_vscode")                 : (enter_text_mode(), send_with_capslock("s")))
d:: mouse_top() ? CycAct.slow("open_current_path", "show_thishotkey", "show_thishotkey")        : (mouse_edge() ? CycAct.slow("{PgDn}")                     : (enter_text_mode(), send_with_capslock("d")))
f:: mouse_top() ? CycAct.slow("app_codex", "show_thishotkey", "show_thishotkey")                : (mouse_edge() ? CycAct.slow("open_current_path")          : (copy_or_cut_by_count()                    ))
g:: mouse_top() ? CycAct.slow("app_notepad", "show_thishotkey", "show_thishotkey")              : (mouse_edge() ? CycAct.slow("app_notepad")                : (enter_text_mode(), send_with_capslock("g")))
z:: mouse_top() ? CycAct.slow("app_powershell_terminal")                                        : (mouse_edge() ? CycAct.slow("app_powershell")             : (enter_text_mode(), send_with_capslock("z")))
x:: mouse_top() ? CycAct.slow("app_word", "show_thishotkey", "show_thishotkey")                 : (mouse_edge() ? CycAct.slow("app_word")                   : (enter_text_mode(), send_with_capslock("x")))
c:: mouse_top() ? CycAct.slow("close_window_in_taskbar", "show_thishotkey", "show_thishotkey")  : (mouse_edge() ? CycAct.slow("close_window_in_taskbar")    : (enter_text_mode(), send_with_capslock("c")))
v:: mouse_top() ? CycAct.slow("url_bilibili", "url_qqlive", "url_xiaohongshu")                  : (mouse_edge() ? CycAct.slow("url_bilibili")               : (SendInput("^v")                           ))
b:: mouse_top() ? CycAct.slow("^+!w")            : (mouse_edge() ? CycAct.slow("^!+2", "^!+1", "^!+3")       : (enter_text_mode(), send_with_capslock("b")))

XButton1:: mouse_top() ? Send("{RAlt}") : (mouse_edge() ? CycAct.slow("!+{Tab}") : SendInput("{XButton1}"))
XButton2:: mouse_top() ? Send("{RAlt}") : (mouse_edge() ? CycAct.slow("!+{Tab}") : SendInput("{XButton2}"))
F23:: mouse_top() ? Send("{RAlt}") : (mouse_edge() ? CycAct.slow("!+{Tab}") : SendInput("{F23}"))
F24:: mouse_top() ? Send("{RAlt}") : (mouse_edge() ? CycAct.slow("!+{Tab}") : SendInput("{F24}"))
Space:: mouse_top() ? CycAct.slow("app_keyboard") : (mouse_edge() ? CycAct.slow("!+{Tab}") : SendInput("{Space}"))
; Space:: mouse_top() ? Send("{RAlt}") : (mouse_edge() ? CycAct.slow("!+{Tab}") : SendInput("{Space}"))

;  CycAct.slow("^+!q"), Sleep(1000), CycAct.slow("^+!w")
~y:: enter_text_mode()
~u:: enter_text_mode()
~i:: enter_text_mode()
~o:: enter_text_mode()
~p:: enter_text_mode()
~h:: enter_text_mode()

~k:: enter_text_mode()
~l:: enter_text_mode()
~m:: enter_text_mode()
~n:: enter_text_mode()

#HotIf  
; #endregion

; #region always

#HotIf True

msgbox_now() {
    SendText("`nMsgBox(A_Now, , `"T2`")") 
    SendInput("{Left 9}+{Left 5}")
}


release_stuck_keys() {
    static keys := [
        "LControl", "RControl",
        "LShift",   "RShift",
        "LAlt",     "RAlt",
        "LWin",     "RWin",
        "Space",    "Enter",
    ]

    for key in keys {
        SendInput "{" key " Up}"
    }
}

$*F24:: release_stuck_keys()

~MButton:: CycAct.fast("{Volume_Mute}")
WheelUp:: !mouse_edge() ? SendInput("{WheelUp}") : SendInput("{volume_up}")
WheelDown:: !mouse_edge() ? SendInput("{WheelDown}") : SendInput("{volume_down}")
Ctrl::      action_ctrl()
RShift::    SendInput("{Tab}") 
~LButton::  exit_text_mode()
RButton::   enter_text_mode(), SendInput("{RButton}")
; F23::       action_f23()
; F24::       action_f24()
; XButton1::  action_xbutton1()
; XButton2::  action_xbutton2()
2:: ime_is_composing() ? SendInput("2") : SendInput("{Delete}")
3::         action_key3()
4::         SendInput("{Backspace}") 
8::         action_key8()
9::         SendInput("{PgUp}")
Enter::     action_enter()
Space:: SendInput("{Space}")


XButton1 & 2::      !mouse_edge() ? CycAct.slow("^n", "^+n", "#n")               : CycAct.slow("{F6}", "^{F6}", "+{F6}")
XButton1 & 3::      !mouse_edge() ? CycAct.slow("^{Enter}")                      : CycAct.slow("{F6}", "^{F6}", "+{F6}")
XButton1 & 4::      !mouse_edge() ? CycAct.slow("^m", "^+m", "#m")               : CycAct.slow("{F6}", "^{F6}", "+{F6}")
XButton1 & q::      !mouse_edge() ? CycAct.slow("^q", "^+q", "#q")               : CycAct.slow("^y", "^+y", "#y") 
XButton1 & w::      !mouse_edge() ? CycAct.slow("close_win_by_ctrlw", "^+w", "#w") : CycAct.slow("^u", "^+u", "#u") 
XButton1 & e::      !mouse_edge() ? CycAct.slow("^e", "^+e", "#e")               : CycAct.slow("^i", "^+i", "#i") 
XButton1 & r::      !mouse_edge() ? CycAct.slow("^r", "^+r", "#r")               : CycAct.slow("^o", "^+o", "#o") 
XButton1 & t::      !mouse_edge() ? CycAct.slow("^t", "^+t", "#t")               : CycAct.slow("^p", "^+p", "#p")
XButton1 & a::      !mouse_edge() ? CycAct.slow("^a", "^+a", "#a")               : CycAct.slow("^h", "^+h", "#h")
XButton1 & s::      !mouse_edge() ? CycAct.slow("^s", "^+s", "#s")               : CycAct.slow("^j", "^+j", "#j")
XButton1 & d::      !mouse_edge() ? CycAct.slow("^d", "^+d", "#d")               : CycAct.slow("^k", "^+k", "#k")
XButton1 & f::      !mouse_edge() ? CycAct.slow("^f", "^+f", "#f")               : CycAct.slow("^l", "^+l", "#l")
XButton1 & g::      !mouse_edge() ? CycAct.slow("^g", "^+g", "#g")               : CycAct.slow("^;", "^+;", "#;")
XButton1 & z::      !mouse_edge() ? CycAct.slow("^z", "^+z", "#z")               : CycAct.slow("show_thishotkey")
XButton1 & x::      !mouse_edge() ? CycAct.slow("^x", "^+x", "#x")               : CycAct.slow("show_thishotkey")
XButton1 & c::      !mouse_edge() ? CycAct.slow("^c", "^+c", "#c")               : CycAct.slow("show_thishotkey")
XButton1 & v::      !mouse_edge() ? step_clip()                                  : CycAct.slow("^v", "^+v", "#v", "#^v") 
XButton1 & b::      !mouse_edge() ? CycAct.slow("^b", "^+b", "#b")               : CycAct.slow("{F12}", "^{F12}", "+{F12}")   
XButton1 & Esc::    !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
XButton1 & RShift:: !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
XButton1 & Ctrl::   !mouse_edge() ? CycAct.slow("!+{Tab}")                       : CycAct.slow("show_thishotkey")
XButton1 & Space::  !mouse_edge() ? CycAct.slow("!#{Space}")                     : CycAct.slow("show_thishotkey")

F23 & 2::           !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & 3::           !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & 4::           !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & q::           !mouse_edge() ? CycAct.slow("{F6}", "^{F6}", "+{F6}")        : CycAct.slow("{F6}", "^{F6}", "+{F6}")
F23 & w::           !mouse_edge() ? CycAct.slow("{F7}", "^{F7}", "+{F7}")        : CycAct.slow("{F7}", "^{F7}", "+{F7}")
F23 & e::           !mouse_edge() ? CycAct.slow("{F8}", "^{F8}", "+{F8}")        : CycAct.slow("{F8}", "^{F8}", "+{F8}")
F23 & r::           !mouse_edge() ? CycAct.slow("{F9}", "^{F9}", "+{F9}")        : CycAct.slow("{F9}", "^{F9}", "+{F9}")
F23 & t::           !mouse_edge() ? CycAct.slow("{F10}", "^{F10}", "+{F10}")     : CycAct.slow("{F10}", "^{F10}", "+{F10}")
F23 & a::           !mouse_edge() ? CycAct.slow("{F1}", "^{F1}", "+{F1}")        : CycAct.slow("{F1}", "^{F1}", "+{F1}") 
F23 & s::           !mouse_edge() ? CycAct.slow("{F2}", "^{F2}", "+{F2}")        : CycAct.slow("{F2}", "^{F2}", "+{F2}")
F23 & d::           !mouse_edge() ? CycAct.slow("{F3}", "^{F3}", "+{F3}")        : CycAct.slow("{F3}", "^{F3}", "+{F3}")
F23 & f::           !mouse_edge() ? CycAct.slow("{F4}", "!{F4}", "+{F4}")        : CycAct.slow("{F4}", "!{F4}", "+{F4}") 
F23 & g::           !mouse_edge() ? CycAct.slow("{F5}", "+{F5}", "!{F5}")        : CycAct.slow("{F5}", "+{F5}", "!{F5}") 
F23 & z::           !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & x::           !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & c::           !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & v::           !mouse_edge() ? CycAct.slow("{F11}", "^{F11}", "+{F11}")     : CycAct.slow("{F11}", "^{F11}", "+{F11}")  
F23 & b::           !mouse_edge() ? CycAct.slow("{F12}", "^{F12}", "+{F12}")     : CycAct.slow("{F12}", "^{F12}", "+{F12}")   
F23 & Esc::         !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & RShift::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & Ctrl::        !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")
F23 & Space::       !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey")



RButton & 2::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("url_yunwu", "url_lovart", "url_claude", "url_chatgpt", "url_gemini", "url_copilot", "do_nothing"))
RButton & 3::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("app_qianwen", "url_deepseek", "app_doubao", "url_coze", "url_kimi", "do_nothing"))
RButton & 4::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("app_qianwen", "url_deepseek", "do_nothing"))
RButton & q::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("url_zhyw", "url_zhlmt", "url_wjx", "do_nothing"))
RButton & w::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("{Escape}")                      : CycAct.slow("app_wechat", "app_wxwork", "app_wemeet", "do_nothing"))
RButton & e::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("^{WheelUp}")                    : CycAct.slow("app_everything", "app_evernote", "app_eagle", "do_nothing"))
RButton & r::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("^z")                            : CycAct.slow("^z"))
RButton & t::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("^y")                            : CycAct.slow("^y"))
RButton & a::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("app_cherrystudio"))
RButton & s::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("+#{Left}")                      : CycAct.slow("+#{Left}"))
RButton & d::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("^{WheelDown}")                  : CycAct.slow("open_current_path", "open_recycle_bin", "open_duty_schedule", "do_nothing"))
RButton & f::       mouse_top() ? CycAct.slow("path_weekly_report")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("app_notepad", "app_excel", "app_baidu", "do_nothing"))
RButton & g::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("paste_format")                  : CycAct.slow("url_jd", "url_xianyu", "url_taobao", "do_nothing"))
RButton & z::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("app_powershell", "app_powershell_admin", "app_powershell_terminal", "app_cmd", "do_nothing"))
RButton & x::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("app_word", "app_mailmaster", "app_youdaodict", "do_nothing"))
RButton & c::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("close_window_in_taskbar"))
RButton & v::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("url_bilibili", "url_qqlive", "url_xiaohongshu", "url_iqiyi", "url_douyin", "do_nothing"))
RButton & b::       mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("path_pbb_server", "path_wh_server", "path_gr_server", "do_nothing"))
RButton & Esc::     mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey"))
RButton & RShift::  mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey"))
RButton & Ctrl::    mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey"))
RButton & Space::   mouse_top() ? CycAct.slow("url_chatgpt", "url_yunwu", "url_lovart")    : (mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey"))

XButton2 & 2::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & 3::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & 4::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & q::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & w::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & e::      !mouse_edge() ? CycAct.slow("{Up}")                          : CycAct.slow("{Up}")            
XButton2 & r::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & t::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & a::      !mouse_edge() ? CycAct.slow("move_tab_new_window")               : CycAct.slow("show_thishotkey") 
XButton2 & s::      !mouse_edge() ? CycAct.slow("{Left}")                        : CycAct.slow("{Left}")          

XButton2 & d::      !mouse_edge() ? CycAct.slow("{Down}")                        : CycAct.slow("{Down}")          
XButton2 & f::      !mouse_edge() ? CycAct.slow("{Right}")                       : CycAct.slow("{Right}")         
XButton2 & g::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & z::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & x::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & c::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & v::      !mouse_edge() ? CycAct.slow("paste_ai_prompt_1", "paste_ai_prompt_2", "paste_ai_prompt_3") : CycAct.slow("show_thishotkey") 
XButton2 & b::      !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & Esc::    !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & RShift:: !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & Ctrl::   !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 
XButton2 & Space::  !mouse_edge() ? CycAct.slow("show_thishotkey")               : CycAct.slow("show_thishotkey") 

~LButton & 2::      CycAct.slow("show_thishotkey")         
~LButton & 3::      CycAct.slow("reload_me") 
~LButton & 4::      CycAct.slow("toggle_suspend") 
~LButton & q::      CycAct.slow("^!+2") 
~LButton & w::      MouseClick(, , , , , "U"), CycAct.slow("close_window") 
~LButton & e::      CycAct.fast("max_window") 
~LButton & r::      CycAct.slow("left_window", "right_window") 
~LButton & t::      CycAct.slow("right_window") 
~LButton & a::      CycAct.slow("+#{Left}") 
~LButton & s::      CycAct.slow("^z") 
~LButton & d::      CycAct.slow("min_win", "#d") 
~LButton & f::      MouseClick(, , , , , "U"), CycAct.fast("+{End}^c", "{Delete}")
~LButton & z::      msgbox_now()
; ~LButton & x::      CycAct.slow("^+!w")
~LButton & x::      CycAct.slow("^+!q"), Sleep(1000), CycAct.slow("^+!w")
~LButton & c::      CycAct.slow("app_cherrystudio") 
~LButton & v::      MouseClick(, , , , , "U"), SendInput("+{End}^v")
~LButton & b::      DpiPresetEditor.Show()  
~LButton & Esc::    CycAct.slow("shutdown_pc", "restart_pc", "lock_screen")    
~LButton & RShift:: CycAct.slow("show_thishotkey")         
~LButton & Ctrl::   CycAct.slow("show_thishotkey")         
~LButton & Space::  CycAct.slow("!#{Space}")  
           
F24 & 2::           CycAct.slow("show_thishotkey")
F24 & 3::           CycAct.slow("show_thishotkey")
F24 & 4::           CycAct.slow("show_thishotkey")
F24 & q::           SendInput("6")
F24 & w::           SendInput("7")
F24 & e::           SendInput("8")
F24 & r::           SendInput("9")
F24 & t::           SendInput("0")
F24 & a::           SendInput("1")
F24 & s::           SendInput("2")
F24 & d::           SendInput("3")
F24 & f::           SendInput("4")
F24 & g::           SendInput("5")
F24 & z::           SendInput("1")
F24 & x::           SendInput("1")
F24 & c::           SendInput("1")
F24 & v::           SendInput("_")
F24 & b::           SendInput(".")          
F24 & Esc::         CycAct.slow("show_thishotkey")
; F24 & RShift::      CycAct.slow("show_thishotkey")
F24 & Ctrl::        CycAct.slow("show_thishotkey")
F24 & Space::       CycAct.slow("show_thishotkey")
       
>^2:: CycAct.slow("show_thishotkey")
>^3:: CycAct.slow("show_thishotkey")
>^4:: CycAct.slow("show_thishotkey")
>^q:: CycAct.slow("^p", "^+p", "#p")
>^w:: CycAct.slow("show_thishotkey")
>^e:: CycAct.slow("^e", "^+e", "#e")
>^r:: CycAct.slow("show_thishotkey")
>^t:: CycAct.slow("show_thishotkey")
>^a:: CycAct.slow("show_thishotkey")
>^s:: CycAct.slow("show_thishotkey")
>^d:: CycAct.slow("show_thishotkey")
>^f:: CycAct.slow("^z") 
>^g:: CycAct.slow("^h", "^+h", "#h")
>^z:: CycAct.slow("show_thishotkey")
>^x:: CycAct.slow("show_thishotkey")
>^c:: CycAct.slow("show_thishotkey")
>^v:: CycAct.slow("show_thishotkey")
>^b:: CycAct.slow("show_thishotkey")
>^7:: SendInput("{^}")
>^8:: SendInput("8")
>^9:: Send(Chr(0x4E28))
>^y:: SendInput("``")
>^u:: SendInput("?")
>^i:: SendInput("'")
>^o:: SendInput("[") 
>^p:: SendInput("]") 
>^h:: SendInput("\")
>^j:: SendInput(", ")
>^k:: SendInput(".")
>^l:: step_clip()    
>^;:: SendInput(".")
>^n:: SendInput("<")
>^m:: SendInput(">")
>^F23::         SendInput("{PgUp}") 
>^F24::         CycAct.slow("show_thishotkey")
>^RButton::     screenshot()
>^XButton1::    SendInput("{PgDn}")

>+7:: CycAct.slow("show_thishotkey")
>+8:: SendInput("{Space}"), Sleep(50), SendInput("{Space}"), Sleep(50), SendInput("{Space}")
>+9:: CycAct.slow("show_thishotkey")
>+y:: CycAct.slow("show_thishotkey")
>+u:: CycAct.slow("show_thishotkey")
>+i:: CycAct.slow("show_thishotkey")
>+o:: CycAct.slow("show_thishotkey")
>+p:: CycAct.slow("show_thishotkey")
>+h:: CycAct.slow("show_thishotkey")
>+j:: CycAct.slow("show_thishotkey")
>+k:: CycAct.slow("show_thishotkey")
>+l:: CycAct.slow("show_thishotkey")
>+n:: CycAct.slow("show_thishotkey")
>+m:: CycAct.slow("show_thishotkey")
>+F23::         CycAct.slow("show_thishotkey")
>+F24::         CycAct.slow("show_thishotkey")
>+RButton::     interactive_region_ocr(, , 0)
>+XButton1::    CycAct.slow("show_thishotkey")                  
>+WheelUp::     SendInput("{PgUp}")    
>+WheelDown::   SendInput("{PgDn}")  

^+RButton:: screenshot_translate()
^+XButton1:: CycAct.slow("show_thishotkey")
^+F23:: CycAct.slow("show_thishotkey")
^+F24:: CycAct.slow("show_thishotkey")

Space & 2:: SendInput("@")
Space & 3:: SendInput("{#}")
Space & 4:: SendInput("$")
Space & q:: SendInput("6")
Space & w:: SendInput("7")       
Space & e:: SendInput("8")       
Space & r:: SendInput("9")      
Space & t:: SendInput("0") 
Space & a:: SendInput("1")
Space & s:: SendInput("2")
Space & d:: SendInput("3")
Space & f:: SendInput("4")
Space & g:: SendInput("5")
Space & z:: SendInput("z") 
Space & x:: SendInput("{+}") 
Space & c:: SendInput("-")
Space & v:: SendText("_")
Space & b:: SendText(".")

Space & 7:: SendInput("&")
Space & 8:: SendInput("*")
Space & 9:: SendInput("|")
Space & y:: SendInput("~")
Space & u:: SendInput("{!}")
Space & i:: SendInput("`"")
Space & o:: SendInput("{{}")
Space & p:: SendInput("{`}}")
Space & h:: SendInput("/")  
Space & j:: SendInput("=") 
Space & k:: SendInput("(")  
Space & l:: SendInput(")") 
Space & n:: SendInput(":")
Space & m:: SendInput("%")
Space & RShift:: SendInput("+{Tab}") 

Enter & 2:: SendInput("!{Enter}")
Enter & 3:: SendInput("!{Enter}")
Enter & 4:: show_thishotkey()
Enter & q:: show_thishotkey()
Enter & w:: show_thishotkey() 
Enter & e:: SendInput("{Up}")
Enter & r:: SendInput("^z")
Enter & t:: SendInput("^y")
Enter & a:: SendInput("{Home}") 
Enter & s:: SendInput("{Left}")
Enter & d:: SendInput("{Down}")
Enter & f:: SendInput("{Right}")
Enter & g:: show_thishotkey()
Enter & z:: show_thishotkey()
Enter & x:: show_thishotkey()
Enter & c:: show_thishotkey()
Enter & v:: SendInput("^v")
Enter & b:: show_thishotkey() 

#HotIf

; #endregion 
 

release_modifiers() {
    static modifiers := [
        "LControl",     "RControl",
        "LShift",       "RShift",
        "LAlt",         "RAlt",
        "LWin",         "RWin",
        "Space",        "Enter", 
        "XButton1",     "XButton2", 
        "LButton",      "RButton", 
        "F23",          "F24"
    ]

    for key in modifiers {
        SendInput "{" key " Up}"
    }

    ToolTip "修饰键已释放"
    SetTimer () => ToolTip(), -800
}

; ~*Esc::release_modifiers()