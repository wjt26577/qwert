(defun c:ImportLinesToText ( / filePath file lines ss i ent obj)
  ;; 设置固定文件路径
  (setq filePath "D:\\temp\\test.txt")

  ;; 检查文件是否存在
  (if (not (findfile filePath))
    (prompt "\n文件未找到，请确认路径是否正确。")
    
    ;; 文件存在，开始读取内容
    (progn
      (setq file (open filePath "r"))
      (setq lines '())

      ;; 逐行读取并保存到列表中
      (while (setq line (read-line file))
        (setq lines (append lines (list line)))
      )
      (close file)

      ;; 提示用户选择文本对象
      (prompt "\n请选择要更新的文本对象（每个对象对应一行）...")
      (setq ss (ssget '((0 . "TEXT,MTEXT")))) ; 只允许选择 TEXT 和 MTEXT

      ;; 如果没有选中对象
      (if (null ss)
        (prompt "\n未选择任何文本对象。")
        
        ;; 否则进行逐个赋值
        (progn
          (setq i 0)
          (repeat (sslength ss)
            (if (< i (length lines))
              (progn
                (setq ent (ssname ss i))
                (setq obj (vlax-ename->vla-object ent))
                (vlax-put-property obj 'TextString (nth i lines))
                (setq i (1+ i))
              )
              (exit) ; 如果行数不够，提前结束循环
            ) ; if
          ) ; repeat

          (princ (strcat "\n成功更新了 " (itoa i) " 个文本对象。"))
        ) ; progn
      ) ; if
    ) ; progn
  ) ; if

  (princ)
)

(princ "\n加载成功：输入 'ImportLinesToText' 将 D:\\temp\\test.txt 的每行内容写入选中文本对象。")
(princ)