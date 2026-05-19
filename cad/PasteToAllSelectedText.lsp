(defun c:PasteToAllSelectedText ( / clipData ss i ent obj)
  ;; 使用 Windows Script Host 来获取剪贴板内容
  ;; 因为标准的 AutoLISP 没有直接访问剪贴板的功能
  ;; 这里使用 vbscript 作为中间件来获取剪贴板内容
  
  ;; 创建临时的 VBScript 文件
  (setq tempVBS (vl-filename-mktemp nil nil ".vbs"))
  (setq file (open tempVBS "w"))
  (write-line "Set WshShell = CreateObject(\"WScript.Shell\")" file)
  (write-line "Set IE = CreateObject(\"InternetExplorer.Application\")" file)
  (write-line "IE.Visible = 0" file)
  (write-line "Wscript.Sleep 100" file)
  (write-line "clipboard = IE.document.parentwindow.clipboardData.GetData(\"text\")" file)
  (write-line "WScript.Echo clipboard" file)
  (close file)

  ;; 执行 VBScript 并捕获输出
  (setq clipCmd (strcat "cscript //nologo " tempVBS))
  (setq clipData (substr (rtos (exec (strcat "(" clipCmd ")")) 2 0) 2))

  ;; 删除临时文件
  (vl-file-delete tempVBS)

  ;; 如果成功获取了剪贴板内容
  (if (/= clipData "")
    (progn
      ;; 提示用户选择文本对象
      (prompt "\n请选择要替换的文本对象...")
      (setq ss (ssget '((0 . "TEXT,MTEXT"))))
      
      ;; 如果选择了对象
      (if ss
        (progn
          (setq i 0)
          (repeat (sslength ss)
            (setq ent (ssname ss i))
            (setq obj (vlax-ename->vla-object ent))
            
            ;; 更新 TextString 属性
            (vlax-put-property obj 'TextString clipData)
            (setq i (1+ i))
          ) ; repeat
          (princ (strcat "\n成功更新了 " (itoa i) " 个文本对象。"))
        ) ; progn
        (prompt "\n未选择任何文本对象。")
      ) ; if
    ) ; progn
    (prompt "\n剪贴板为空。")
  ) ; if
  
  (princ)
)

(defun exec (cmd / result)
  ;; 执行命令并返回结果
  (setq result (read (strcat "(command \"!\"" cmd "\")")))
  result
)

(princ "\n加载成功：输入 'PasteToAllSelectedText' 将剪贴板内容粘贴到选中文本中。")
(princ)