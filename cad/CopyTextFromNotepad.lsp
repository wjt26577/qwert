
; 定义将文本复制到剪贴板的函数
(defun CopytoClipboard (text / Clip_board)
    (setq Clip_board (Vlax-Get-Property (Vlax-Get (vlax-create-object "htmlfile") 'ParentWindow) 'ClipboardData))
    (Vlax-Invoke Clip_board 'SetData "text" text)
    (vlax-release-object Clip_board)
    text
)

; 主函数
(defun c:CopyTextFromNotepad (/ filename file line lines)
    ; 获取记事本文件路径
    (setq filename (getfiled "选择记事本文件" "" "txt" 2))
    (if filename
        (progn
            ; 打开文件
            (setq file (open filename "r"))
            (if file
                (progn
                    (setq lines '())
                    ; 逐行读取文件内容
                    (while (setq line (read-line file))
                        (setq lines (cons line lines))
                    )
                    ; 关闭文件
                    (close file)
                    ; 反转列表，恢复原来的顺序
                    (setq lines (reverse lines))
                    ; 提示用户切换到CAD窗口
                    (alert "请在5秒内切换到CAD窗口...")
                    (sleep 5)
                    ; 依次复制并粘贴每行文本
                    (foreach line lines
                        ; 将文本复制到剪贴板
                        (CopytoClipboard line)
                        ; 在CAD中模拟粘贴操作
                        (command "_pasteclip")
                        (command "_enter") ; 模拟回车键确认
                        (sleep 0.5) ; 等待粘贴完成
                    )
                    (alert "已将所有文本复制到CAD中。")
                )
                (alert "无法打开文件。")
            )
        )
        (alert "未选择文件。")
    )
    (princ)
)