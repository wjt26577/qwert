(defun c:InsertJump ( / pt)
  ;; 请求用户选择插入点
  (setq pt (getpoint "\n选择插入点: "))
  ;; 使用默认比例因子1和旋转角度0插入名为'jump'的块
  (command "-insert" "jump" pt 1 1 0)
  ;; 结束函数
  (princ)
)
(princ "\n加载成功：输入 'InsertJump' 插入 'jump' 块。")