
#Requires AutoHotkey v2.0
#SingleInstance Force 

try {
    script_dir := A_ScriptDir . "\main.ahk"
    if WinExist(script_dir) {
        WinClose
    }

    if GetKeyState("Ctrl") {
        script_dir := "*RunAs " . script_dir
    }
    Run(script_dir)
}
catch as err
    msgbox(err.Message)  
