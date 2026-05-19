(defun c:UpdateTextsFromFile ( / filePath fileContent textsList ss i ent obj)
  ;; 请求用户选择文本文件
  (setq filePath (getfiled "选择包含文本的文件" "" "txt" 16))
  
  ;; 打开并读取文件内容
  (if (/= filePath nil)
    (progn
      (setq file (open filePath "r"))
      (setq fileContent (read-line file t)) ; 读取整个文件为字符串
      (close file)
      
      ;; 根据换行符分割文本
      (setq textsList (split-string fileContent "\n"))
      
      ;; 请求用户选择要更新的文本对象
      (setq ss (ssget '((0 . "TEXT,MTEXT"))))
      
      ;; 遍历选中的每个文本对象
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq obj (vlax-ename->vla-object ent))
        
        ;; 更新文本值
        (if (< i (length textsList))
          (vlax-put-property obj 'TextString (nth i textsList))
        )
        (setq i (1+ i))
      )
    )
  )
  (princ)
)

;; 辅助函数：用于分割字符串
(defun split-string (inputString delimiter / lst pos)
  (setq lst '())
  (while (setq pos (vl-string-search delimiter inputString))
    (setq lst (append lst (list (substr inputString 1 pos))))
    (setq inputString (substr inputString (+ pos (strlen delimiter)))))
  (append lst (list inputString))
)

(princ "\n加载成功：输入 'UpdateTextsFromFile' 开始更新文本。")