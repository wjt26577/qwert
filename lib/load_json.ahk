
try {
    ; 读取配置
    global static_config := json.load(FileRead(A_ScriptDir . "\settings\settings.json" ))

    global json_config := json.load(FileRead(A_ScriptDir . "\settings\settings.json" ))

    secrets := json.load(FileRead(A_ScriptDir . "\settings\secrets.json" ))

    baidu_ai := secrets["baidu_ai"]

    
    run_config := Map()
    run_config := json.load(FileRead("config\run_config.json"))
        
   

} catch as err {
    MsgBox(Format("配置加载失败:`n{1}`nFile: {2}`nLine: {3}", err.message, err.file, err.Line))
}