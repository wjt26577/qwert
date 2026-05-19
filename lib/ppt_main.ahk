

; #region power_point

#HotIf WinActive("ahk_group ctrlw_group") || (WinGetTitle("A") ~= "i)^(Word|Excel|PowerPoint)$")
    XButton1 & w:: CycAct.slow("!{F4}", "^+w", "#w")
#HotIf


#HotIf step_clip_active()
    XButton1 & f:: step_clip_back()
#HotIf

#HotIf WinActive("ahk_group group_office")  
    ; XButton1 & 3:: CycAct.fast("set_fonts1", "set_fonts2", "set_fonts3", "set_fonts4", "set_fonts5", "set_fonts6", "set_fonts7", "set_fonts8")

    ; XButton1::SendInput("+{LButton}")
    ; RButton & f:: SendInput("{F4}}")  

    ; XButton1 & 3:: set_excel_fonts_1()
    XButton1 & s:: CycAct.slow("^s", "{F12}", "#s")
   
#HotIf 

#HotIf WinActive("ahk_class XLMAIN")   
    XButton1::SendInput("+{LButton}")
    Enter:: SendInput("!{Enter}") 
  


    ; F23:: MsgBox(A_Now, , "T2")
    ; MButto::+LButton
#HotIf

#HotIf WinActive("ahk_class OpusApp")   
    XButton1::SendInput("+{LButton}")

    ; F23:: MsgBox(A_Now, , "T2")
#HotIf

#HotIf WinActive("ahk_exe POWERPNT.EXE") && is_text_mode()




; ============================================================
; 公共函数
; ============================================================

blue_cat(shapes_mdi, none_mdi, shapes_ribbon, none_ribbon) {
    try {
        ppt_app := ComObjActive("PowerPoint.Application")
        selection_type := selection_type_map[ppt_app.ActiveWindow.Selection.Type]
        MouseGetPos(,,, &control)

        ; --- 文本模式：输出字符 ---
        if (selection_type == "text") {
            send_char()
            return
        }

        func_arr := []

        if (control == "MDIClient1" && selection_type == "shapes" && shapes_mdi.Length > 0)
            func_arr := shapes_mdi  
        
        if (control == "MDIClient1" && selection_type == "none" && none_mdi.Length > 0)
            func_arr := none_mdi  
        
        if ((control ~= "NetUIHWND") && selection_type == "shapes" && shapes_ribbon.Length > 0)
            func_arr := shapes_ribbon  
        
        if ((control ~= "NetUIHWND") && selection_type == "none" && none_ribbon.Length > 0)
            func_arr := none_ribbon  

        func_name := func_arr[1]              
        params := []
        
        loop func_arr.Length - 1          ; 剩余元素作为参数
            params.Push(func_arr[A_Index + 1])

        %func_name%(params*)

    } catch as err
        MsgBox("错误: " err.message 
            "`n文件: " err.file 
            "`n行号: " err.Line)
}


; ============================================================
; 按键调用
; ============================================================
w:: SendInput("{Escape}")

q:: blue_cat(
    ["fast_cat", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
)

e:: blue_cat(
    ["fast_cat", "FontSizeIncrease"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "^{Home}"]                                                                           ; none/slides + Ribbon
)

r:: blue_cat(
    ["fast_cat", "ObjectSendToBack", "ObjectBringToFront"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "format_dialog", "shadow_dialog", "size_dialog"]     ; shapes + Ribbon
   ,["fast_cat", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
)

t:: blue_cat(
    ["fast_cat", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "AnimationCustom"]                                                                           ; none/slides + Ribbon
)

a:: blue_cat(
    ["fast_cat", "ppt_a"]   ; shapes + MDI
   ,["fast_cat", "ppt_a"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "ShowClipboard"]                                                                           ; none/slides + Ribbon
)

s:: blue_cat(
    ["fast_cat", "ppt_s"]   ; shapes + MDI
   ,["fast_cat", "ppt_s"]                                                                          ; none/slides + MDI
   ,["fast_cat", "selection_pane"]     ; shapes + Ribbon
   ,["fast_cat", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
)

d:: blue_cat(
    ["fast_cat", "FontSizeDecrease"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "^{End}"]                                                                           ; none/slides + Ribbon
)

g:: blue_cat(
    ["fast_cat", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "{F5}"]                                                                           ; none/slides + Ribbon
)

z:: blue_cat(
    ["fast_cat", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "TransitionPreview"]                                                                           ; none/slides + Ribbon
)

x:: blue_cat(
    ["fast_cat", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "{Alt}y1y5"]                                                                           ; none/slides + Ribbon
)

c:: blue_cat(
    ["fast_cat", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
   ,["fast_cat", "ShapeRoundedRectangle"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "SlideShowSetUpDialog"]                                                                           ; none/slides + Ribbon
)

b:: blue_cat(
    ["fast_cat", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
   ,["fast_cat", "selection_pane"]                                                                          ; none/slides + MDI
   ,["fast_cat", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
   ,["fast_cat", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
)


ppt_s() {
    start_time := A_TickCount
    Send("{Ctrl down}{Shift down}") 

    KeyWait("s")

    Send("{Ctrl up}{Shift up}") 

    if (A_TickCount - start_time < 300) {   
        SendInput("+{LButton}")
    }
}

ppt_a() {
    start_time := A_TickCount
    Send("{Shift down}") 

    KeyWait("a")

    Send("{Shift up}") 

    if (A_TickCount - start_time < 300) {   
        SendInput("+{LButton}")
    }
}



    
    

^RButton:: add_frame()
; w & LButton:: add_frame()


Tab:: SendInput("{Escape}")

^+s:: return
+a:: return 


   
    ; w:: fast_cat("ppt | ObjectsAlignLeftSmart | 左对齐", "ppt | ObjectsAlignCenterHorizontalSmart | 水平居中", "ppt | ObjectsAlignRightSmart | 右对齐")
    ; w:: fast_cat("func | objleft_selpane", "func | objcenter_selpane", "func | objright_selpane")
    ; e:: ("func | hide_objects | 隐藏形状")
    ; r:: fast_cat("ppt | ObjectSendToBack | 置于底层", "ppt | ObjectBringToFront | 置于顶层")
    ; t:: _500("ppt | ObjectFlipHorizontal | 水平翻转对象", "ppt | ObjectFlipVertical | 垂直翻转对象")
    ; d:: ("func | duplicate_down_to_origin | 啥也不干!")
    ; g:: fast_cat("func | textleft_view", "func | textcenter_view", "func | textright_view")
    ; z:: do_nothing()
    ; x:: do_nothing()
    ; c:: ("func | change_all_fonts | 啥也不干!")
    ; b:: fast_cat("func | textleft_1row_slide_guides", "func | textleft_2row_slide_guides", "func | textleft_3row_slide_guides", "func | textleft_4row_slide_guides")
    
    ; q:: ("func | add_h11 | 一级标题_默认", "func | add_h12 | 一级标题_品牌", "func | add_h13 | 一级标题_强调", "func | add_h14 | 一级标题_反色")
    ; w:: ("func | add_h21 | 二级标题_默认", "func | add_h22 | 二级标题_品牌", "func | add_h23 | 二级标题_强调", "func | add_h24 | 二级标题_反色")
    ; e:: ("func | add_h31 | 三级标题_默认", "func | add_h32 | 三级标题_品牌", "func | add_h33 | 三级标题_强调", "func | add_h34 | 三级标题_反色")
    ; r:: ("func | add_h41 | 四级标题_默认", "func | add_h42 | 四级标题_品牌", "func | add_h43 | 四级标题_强调", "func | add_h44 | 四级标题_反色")
    ; t:: ("func | add_h51 | 五级标题_默认", "func | add_h52 | 五级标题_品牌", "func | add_h53 | 五级标题_强调", "func | add_h54 | 五级标题_反色")
    ; d:: ("func | add_b11", "func | add_b12", "func | add_b13", "func | add_b14")
    ; g:: ("func | add_shape_1 | 容器_默认", "func | add_shape_2 | 容器_品牌", "func | add_shape_3 | 容器_强调", "func | add_shape_4 | 容器_反色")
    ; z:: do_nothing()
    ; x:: do_nothing()
    ; c:: do_nothing()
    ; b:: ("inpt | do_nothing | 选择到行末并粘贴")

    ; q & LButton::   CycAct.fast("add_h11", "apply_h12", "apply_h13", "apply_h14")
    ; q & RButton::   CycAct.fast("add_guides_from_selected_shape", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; q & XButton1::  CycAct.fast("test_generate_random_text", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; q & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; q & F23::       CycAct.fast("position_h1", "AlignCenter", "AlignRight", "AlignJustify", "ParagraphDistributed")
    ; q & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; w & LButton::   CycAct.fast("add_h21", "apply_h22", "apply_h23", "apply_h24")

    ; w & RButton:: blue_cat(
    ;              ["fast_mouse", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
    ;             ,["fast_mouse", "css_gridient"]                                                                          ; none/slides + MDI
    ;             ,["fast_mouse", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
    ;             ,["fast_mouse", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
    ;             )
    ; w & XButton1:: blue_cat(
    ;              ["fast_mouse", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
    ;             ,["fast_mouse", "selection_pane"]                                                                          ; none/slides + MDI
    ;             ,["fast_mouse", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
    ;             ,["fast_mouse", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
    ;             )
    ; w & XButton2:: blue_cat(
    ;              ["fast_mouse", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
    ;             ,["fast_mouse", "selection_pane"]                                                                          ; none/slides + MDI
    ;             ,["fast_mouse", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
    ;             ,["fast_mouse", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
    ;             )
    ; w & F23:: blue_cat(
    ;              ["fast_mouse", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
    ;             ,["fast_mouse", "selection_pane"]                                                                          ; none/slides + MDI
    ;             ,["fast_mouse", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
    ;             ,["fast_mouse", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
    ;             )
    ; w & F24:: blue_cat(
    ;              ["fast_mouse", "ObjectsAlignLeftSmart", "ObjectsAlignCenterHorizontalSmart", "ObjectsAlignRightSmart"]   ; shapes + MDI
    ;             ,["fast_mouse", "selection_pane"]                                                                          ; none/slides + MDI
    ;             ,["fast_mouse", "ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart"]     ; shapes + Ribbon
    ;             ,["fast_mouse", "GuidesShowHide"]                                                                           ; none/slides + Ribbon
    ;             )
    
    
    
    
    ; w & XButton1::  CycAct.fast("test_generate_random_text", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; w & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; w & F23::       CycAct.fast("AlignLeft", "AlignCenter", "AlignRight", "AlignJustify", "ParagraphDistributed")
    ; w & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; e & LButton::   CycAct.fast("add_h31", "apply_h32", "apply_h33", "apply_h34")
    ; e & RButton::   CycAct.fast("add_guides_from_selected_shape", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; e & XButton1::  CycAct.fast("split_paragraphs", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; e & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; e & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; e & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; r & LButton::   CycAct.fast("add_h41", "apply_h42", "apply_h43", "apply_h44")
    ; r & RButton::   CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; r & XButton1::  Run("pythonw vba\aaa.py")

    ; r & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; r & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; r & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    
    ; t & LButton::   CycAct.fast("add_h51", "apply_h52", "apply_h53", "apply_h54")
    ; t & RButton::   CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; t & XButton1::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; t & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; t & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; t & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; d & LButton::   CycAct.fast("add_b11", "apply_b12", "apply_b13", "apply_b14")
    ; d & RButton::   CycAct.fast("resize_by_ref", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; d & XButton1::  CycAct.fast("format_dialog", "shadow_dialog", "size_dialog")
    ; d & XButton2::  CycAct.fast("hide_guides", "show_guides")
    ; d & F23::       CycAct.fast("add_default_master_guides", "add_2column_master_guides", "add_3column_master_guides", "add_4column_master_guides", "add_5column_master_guides", "add_6column_master_guides")
    ; d & F24::       CycAct.fast("add_default_slide_guides", "ObjectsAlignBottomSmart")
 
    ; f & RButton::   CycAct.fast("copy_to_right")
    ; f & XButton1::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; f & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; f & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; f & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; g & RButton::   CycAct.fast("add_capsule_2", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; g & XButton1::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; g & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; g & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; g & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; x & RButton::   CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; x & XButton1::  CycAct.fast("set_layout_to_file", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; x & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; x & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; x & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; c & RButton::   CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; c & XButton1::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; c & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; c & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; c & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")

    ; v & RButton::   CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; v & XButton1::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; v & XButton2::  CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; v & F23::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
    ; v & F24::       CycAct.fast("ObjectsAlignTopSmart", "ObjectsAlignMiddleVerticalSmart", "ObjectsAlignBottomSmart")
        
    ; b & LButton::   CycAct.slow("add_1row_slide_guides", "add_2row_slide_guides", "add_3row_slide_guides", "add_4row_slide_guides")
    ; b & RButton::   CycAct.slow("resize_by_guides", "resize_by_master_guides")
    ; b & XButton1::  CycAct.fast("add_default_master_guides", "add_2column_master_guides", "add_3column_master_guides", "add_4column_master_guides", "add_5column_master_guides", "add_6column_master_guides")
    ; b & XButton2::  CycAct.fast("add_default_slide_guides", "add_2column_slide_guides", "add_3column_slide_guides", "add_4column_slide_guides", "add_5column_slide_guides", "add_6column_slide_guides")
    ; b & F23::       CycAct.fast("delete_slide_guides", "delete_master_guides", )
    ; b & F24::       CycAct.fast("hide_guides", "show_guides")
    


#HotIf
; #endregion
