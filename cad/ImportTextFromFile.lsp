(defun c:ImportTextFromFile (/ fileName fileData ss i ent obj)
  ;; 请求用户选择文件
  (setq fileName (getfiled "选择文本文件" "" "txt" 1)) ;; 1 表示只允许打开文件
  
  ;; 打开文件并读取内容
  (if fileName
    (progn
      (setq file (open fileName "r"))
      (setq fileData (read-line file))
      (close file)
      
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
            (vlax-put-property obj 'TextString fileData)
            (setq i (1+ i))
          ) ; repeat
          (princ (strcat "\n成功更新了 " (itoa i) " 个文本对象。"))
        ) ; progn
        (prompt "\n未选择任何文本对象。")
      ) ; if
    )
    (prompt "\n未选择文件。")
  ) ; if
  
  (princ)
)

(princ "\n加载成功：输入 'ImportTextFromFile' 将外部文本文件内容粘贴到选中文本中。")
(princ)