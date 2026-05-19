#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent()

; =========================
; 配置区
; =========================
TARGET_LANG := "简体中文"              ; 例如："English"、"日文"、"简体中文"
DEEPSEEK_MODEL := "deepseek-v4-flash" ; 可改为 deepseek-v4-pro
DEEPSEEK_API_KEY := "sk-05994ea943324c57a0354bfcb3133c9e"                ; 推荐留空，使用环境变量 DEEPSEEK_API_KEY
; HOTKEY := "^!t"                       ; Ctrl + Alt + T

XButton2 & a::
{
    TranslateClipboard()
}

if (DEEPSEEK_API_KEY != "")
    EnvSet("DEEPSEEK_API_KEY", DEEPSEEK_API_KEY)

; Hotkey(HOTKEY, TranslateClipboard)

TranslateClipboard(*) {
    global TARGET_LANG, DEEPSEEK_MODEL

    try {
        source := A_Clipboard
        inputType := "剪贴板文本"

        if (StrLen(Trim(source, " `t`r`n")) = 0) {
            inputType := "剪贴板图片 OCR"
            ToolTip("正在使用 Windows OCR 识别剪贴板图片...")
            source := OcrClipboardImage()
        }

        if (StrLen(Trim(source, " `t`r`n")) = 0) {
            ToolTip()
            MsgBox("剪贴板里没有可翻译的文本，图片 OCR 也没有识别出文字。", "剪贴板翻译", "Icon!")
            return
        }

        ToolTip("正在调用 DeepSeek 翻译...")
        translation := DeepSeekTranslate(source, TARGET_LANG, DEEPSEEK_MODEL)

        ToolTip()
        ShowResultGui(source, translation, inputType)

    } catch as e {
        ToolTip()
        MsgBox("失败：`n`n" e.Message, "剪贴板翻译", "Iconx")
    }
}

OcrClipboardImage() {
    outPath := TempPath("ocr", "txt")

    ps := "
(
param([string]$OutPath)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not [System.Windows.Forms.Clipboard]::ContainsImage()) {
    throw '剪贴板中没有文本，也没有可 OCR 的图片。请先复制文字或截图。'
}

$image = [System.Windows.Forms.Clipboard]::GetImage()
$tempPng = Join-Path $env:TEMP ('ahk_ocr_' + [guid]::NewGuid().ToString('N') + '.png')

try {
    $image.Save($tempPng, [System.Drawing.Imaging.ImageFormat]::Png)

    Add-Type -AssemblyName System.Runtime.WindowsRuntime

    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrResult, Windows.Media.Ocr, ContentType = WindowsRuntime]

    $asTaskMethods = [System.WindowsRuntimeSystemExtensions].GetMethods()

    $asTaskGeneric = $asTaskMethods | Where-Object {
        ($_.Name -eq 'AsTask') -and ($_.IsGenericMethod) -and ($_.GetParameters().Count -eq 1) -and ($_.GetParameters()[0].ParameterType.Name.StartsWith('IAsyncOperation'))
    } | Select-Object -First 1

    if ($null -eq $asTaskGeneric) {
        throw '无法加载 WinRT 异步转换方法。'
    }

    function Await($Operation, [Type]$ResultType) {
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $task = $asTask.Invoke($null, @($Operation))
        $task.Wait() | Out-Null
        return $task.Result
    }

    $file = Await -Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tempPng)) -ResultType ([Windows.Storage.StorageFile])
    $stream = Await -Operation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) -ResultType ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await -Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) -ResultType ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await -Operation ($decoder.GetSoftwareBitmapAsync()) -ResultType ([Windows.Graphics.Imaging.SoftwareBitmap])

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()

    if ($null -eq $engine) {
        throw '无法创建 Windows OCR 引擎。请在 Windows 设置中安装对应语言包/OCR 语言支持。'
    }

    $result = Await -Operation ($engine.RecognizeAsync($bitmap)) -ResultType ([Windows.Media.Ocr.OcrResult])

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($OutPath, $result.Text, $utf8Bom)
    
}
finally {
    if (Test-Path $tempPng) {
        Remove-Item $tempPng -Force -ErrorAction SilentlyContinue
    }
}
)"

    res := RunPowerShell(ps, QuoteArg(outPath))

    if (res.ExitCode != 0)
        throw Error(CleanPsError(res.StdErr))

    if !FileExist(outPath)
        throw Error("OCR 没有生成结果。")

    text := FileRead(outPath, "UTF-8")
    TryDelete(outPath)
    return text
}

DeepSeekTranslate(sourceText, targetLang, model) {
    sourcePath := TempPath("source", "txt")
    outPath := TempPath("translate", "txt")

    FileAppend(Chr(0xFEFF) sourceText, sourcePath, "UTF-8-RAW")

    ps := "
(
param(
    [string]$SourcePath,
    [string]$OutPath,
    [string]$TargetLang,
    [string]$Model
    `)

$ErrorActionPreference = 'Stop'

$key = $env:DEEPSEEK_API_KEY

if ([string]::IsNullOrWhiteSpace($key)) {
    $key = [Environment]::GetEnvironmentVariable('DEEPSEEK_API_KEY', 'User')
}

if ([string]::IsNullOrWhiteSpace($key)) {
    $key = [Environment]::GetEnvironmentVariable('DEEPSEEK_API_KEY', 'Machine')
}

if ([string]::IsNullOrWhiteSpace($key)) {
    throw '未设置 DeepSeek API Key。'
}

$text = [System.IO.File]::ReadAllText($SourcePath, [System.Text.Encoding]::UTF8)

$systemPrompt = '你是专业翻译引擎。只输出译文，不解释，不添加前后缀。'

$userPrompt = @'
请把下面内容翻译成目标语言。

目标语言：
__TARGET_LANG__

要求：
1. 只输出译文。
2. 保留原文的换行、编号、项目符号、代码、网址、文件路径。
3. 专有名词不确定时保留原文或采用通用译法。
4. 不要解释，不要总结。

原文：
__TEXT__
'@

$userPrompt = $userPrompt.Replace('__TARGET_LANG__', $TargetLang).Replace('__TEXT__', $text)

$messages = @()
$messages += @{
    role = 'system'
    content = $systemPrompt
}
$messages += @{
    role = 'user'
    content = $userPrompt
}

$bodyObj = [ordered]@{
    model = $Model
    messages = $messages
    temperature = 0.2
    stream = $false
}

$body = $bodyObj | ConvertTo-Json -Depth 20

Add-Type -AssemblyName System.Net.Http

$client = [System.Net.Http.HttpClient]::new()

$request = [System.Net.Http.HttpRequestMessage]::new(
    [System.Net.Http.HttpMethod]::Post,
    'https://api.deepseek.com/chat/completions'
    `)

$request.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $key)
$request.Headers.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('application/json'))

$content = [System.Net.Http.StringContent]::new($body, [System.Text.Encoding]::UTF8, 'application/json')
$request.Content = $content

$httpResponse = $client.SendAsync($request).Result
$responseBytes = $httpResponse.Content.ReadAsByteArrayAsync().Result

$jsonText = [System.Text.Encoding]::UTF8.GetString($responseBytes)

if (-not $httpResponse.IsSuccessStatusCode) {
    throw ('DeepSeek HTTP ' + [int]$httpResponse.StatusCode + ': ' + $jsonText)
}

$response = $jsonText | ConvertFrom-Json
$result = $response.choices[0].message.content

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutPath, $result, $utf8Bom)
)"

    args := QuoteArg(sourcePath) " " QuoteArg(outPath) " " QuoteArg(targetLang) " " QuoteArg(model)

    res := RunPowerShell(ps, args)

    TryDelete(sourcePath)

    if (res.ExitCode != 0) {
        TryDelete(outPath)
        throw Error(CleanPsError(res.StdErr))
    }

    if !FileExist(outPath)
        throw Error("DeepSeek 没有返回翻译结果。")

    translation := FileRead(outPath, "UTF-8")
    TryDelete(outPath)

    return translation
}

; [System.IO.File]::WriteAllText($OutPath, $result, [System.Text.Encoding]::UTF8)

; $response = Invoke-RestMethod `
;     -Uri 'https://api.deepseek.com/chat/completions' `
;     -Method Post `
;     -Headers @{
;         Authorization = 'Bearer ' + $key
;     } `
;     -ContentType 'application/json; charset=utf-8' `
;     -Body ([System.Text.Encoding]::UTF8.GetBytes($body))


ShowResultGui(source, translation, inputType) {
    g := Gui("+AlwaysOnTop", "剪贴板翻译 - " inputType)
    g.MarginX := 12
    g.MarginY := 12
    g.SetFont("s10", "Microsoft YaHei UI")

    g.AddText("xm ym", "原文 / OCR 结果：")
    g.AddEdit("xm w820 h170 ReadOnly VScroll", source)

    g.AddText("xm y+10", "DeepSeek 翻译结果：")
    g.AddEdit("xm w820 h260 ReadOnly VScroll", translation)

    copyBtn := g.AddButton("xm y+10 w110 h32", "复制译文")
    closeBtn := g.AddButton("x+10 w90 h32", "关闭")

    copyBtn.OnEvent("Click", (*) => CopyText(translation))
    closeBtn.OnEvent("Click", (*) => g.Destroy())

    g.Show("w860 h560")
}

CopyText(text) {
    A_Clipboard := text
    ToolTip("译文已复制")
    SetTimer(() => ToolTip(), -1000)
}

RunPowerShell(script, args := "") {
    psPath := TempPath("script", "ps1")
    FileAppend(script, psPath, "UTF-8")

    cmd := "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File " QuoteArg(psPath)

    if (args != "")
        cmd .= " " args

    shell := ComObject("WScript.Shell")
    exec := shell.Exec(cmd)

    while (exec.Status = 0)
        Sleep(50)

    stdout := exec.StdOut.ReadAll()
    stderr := exec.StdErr.ReadAll()
    code := exec.ExitCode

    TryDelete(psPath)

    return {
        ExitCode: code,
        StdOut: stdout,
        StdErr: stderr
    }
}

TempPath(prefix, ext) {
    return A_Temp "\ahk_cliptrans_" prefix "_" A_Now "_" Random(100000, 999999) "." ext
}

QuoteArg(s) {
    q := Chr(34)
    return q . StrReplace(s, q, "\" . q) . q
}

TryDelete(path) {
    try {
        if FileExist(path)
            FileDelete(path)
    }
}

CleanPsError(stderr) {
    msg := Trim(stderr, " `t`r`n")
    if (msg = "")
        return "PowerShell 执行失败，但没有返回详细错误。"
    return msg
}