(defun c:ChangeSelectedTextToTest (/ ss i ent obj)
  ;; 请求用户选择文本对象
  (setq ss (ssget '((0 . "TEXT,MTEXT")))) ;; 只选择 TEXT 和 MTEXT 对象
  
  ;; 如果选择了对象
  (if ss
    (progn
      ;; 遍历每一个选中的对象
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq obj (vlax-ename->vla-object ent))
        
        ;; 设置 TextString 属性为"测试文本"
        (vlax-put-property obj 'TextString "测试文本")
        (setq i (1+ i))
      )
      (princ "\n已更新选中的文本对象。")
    )
    (princ "\n未选择任何文本对象。")
  )
  (princ)
)

(princ "\n加载成功：输入 'ChangeSelectedTextToTest' 将选中文本对象的内容更改为“测试文本”。")
(princ)