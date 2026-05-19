#Requires AutoHotkey v2.0
#SingleInstance Force 

; MyGui := Gui()
; MyGui.Add("Text",, "Please enter your name:")
; MyGui.AddEdit("vName")
; MyGui.Show


xbutton2 & f:: ggg()
ppt_in_folder() {    
    g := Gui("-MinimizeBox +AlwaysOnTop",  "PPT文件处理")
    g.MarginX := 20
    g.MarginY := 20
    g.SetFont("s16")

    ; ~XButton2 & f:: funcX('process_ppt_914')
	; ~XButton2 & w:: funcX('split_in_folder')
	; ~XButton2 & e:: funcX('export_jpg_in_folder')
	; ~XButton2 & r:: funcX('rename_914')

    ; e1 := g.AddEdit("w800 R10 cGray")	
    ; e2 := g.AddEdit("w800 R10")	
    b1 := g.AddButton("Default w200", "拆分成单页PPT")
    b2 := g.AddButton("Default w200", "拆分成单页图片")
    b3 := g.AddButton("Default w200", "修改文件名")
    b4 := g.AddButton("Default w200", "取消")
    b1.OnEvent("Click", b1_click) 
    b2.OnEvent("Click", b2_click) 
    b3.OnEvent("Click", b3_click) 
    b4.OnEvent("Click", b4_click) 
    g.Show
    
    b1_click(*) {
        funcX('split_in_folder')
    }
    b2_click(*) {
        funcX('export_jpg_in_folder')
    }
    b3_click(*) {
        funcX('process_ppt_914')
    }
    b4_click(*) {
        g.Destroy()
    }
}

