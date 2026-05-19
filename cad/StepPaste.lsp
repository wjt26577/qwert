(defun c:PasteToTexts ( / clipTxt txtList ss i ent obj)
  ;; 获取剪贴板文本
  (setq clipTxt (vlax-invoke (vlax-get-acad-object) 'GetClipboardData))
  
  ;; 判断是否成功获取到内容
  (if (and clipTxt (/= clipTxt ""))
    (progn
      ;; 按换行符分割成列表
      (setq txtList (split-string clipTxt "\n"))
      
      ;; 提示用户选择文本对象
      (prompt "\n请选择要替换的文本对象...")
      (setq ss (ssget '((0 . "TEXT,MTEXT"))))
      
      ;; 如果选择了对象，并且有文本内容
      (if (and ss txtList)
        (progn
          (setq i 0)
          (repeat (sslength ss)
            (setq ent (ssname ss i))
            (setq obj (vlax-ename->vla-object ent))
            
            ;; 如果还有对应的文本可用
            (if (< i (length txtList))
              (vlax-put-property obj 'TextString (nth i txtList))
            )
            (setq i (1+ i))
          ) ; repeat
          (princ (strcat "\n成功更新了 " (itoa i) " 个文本对象。"))
        ) ; progn
      ) ; if
    ) ; progn
    
    ;; 否则提示错误
    (prompt "\n剪贴板为空或未选择有效文本。")
  ) ; if
  
  (princ)
)

;; 辅助函数：按指定分隔符分割字符串
(defun split-string (str delim / pos lst)
  (while (setq pos (vl-string-search delim str))
    (setq lst (append lst (list (substr str 1 pos))))
    (setq str (substr str (+ pos (strlen delim) 0)))
  )
  (append lst (list str))
)

(princ "\n加载成功：输入 'PasteToTexts' 将剪贴板内容依次粘贴到选中文本中。")
(princ)