Qwert.ahk
├── 1. 启动与全局初始化
│   ├── ProcessSetPriority / CoordMode 设置
│   ├── 托盘图标设置
│   ├── 配置读取
│   ├── 全局变量初始化
│   ├── 颜色/字体/路径常量
│   ├── 剪贴板历史加载
│   └── 启动提示
│
├── 2. 全局状态与模式控制
│   ├── set_global()
│   ├── true_idol()
│   ├── false_idol()
│   ├── set_prior_key()
│   ├── set_superkey_true()
│   ├── set_superkey_false()
│   ├── set_superkey()
│   ├── set_super_power_true()
│   ├── set_super_power_false()
│   ├── set_super_power_idol_false()
│   ├── set_superkey_pure()
│   ├── set_superkey3_enabled()
│   ├── toggle_superkey_unlocked()
│   ├── variable_exists()
│   └── get_var_by_name()
│
├── 3. 通知与提示系统
│   ├── tool_tip()
│   ├── center_info()
│   ├── left_info()
│   ├── clip_info()
│   ├── show_info_yellow()
│   ├── show_info_orange()
│   ├── show_info_red()
│   ├── show_info_blue()
│   ├── show_info_purple()
│   ├── show_info_green()
│   ├── do_nothing()
│   └── DestroyTooltipGui()
│
├── 4. 颜色与显示相关工具
│   ├── rgb2bgr()
│   ├── bgr2rgb()
│   ├── get_contrast_color()  [你前面计划中的颜色工具]
│   ├── is_dark_color()       [你前面计划中的颜色工具]
│   ├── normalize_color()     [你前面计划中的颜色工具]
│   ├── color_to_int()        [你前面计划中的颜色工具]
│   └── color_to_str()        [你前面计划中的颜色工具]
│
├── 5. 字符串处理工具
│   ├── remove_spaces()
│   ├── trim_array()
│   ├── repeat_str()
│   ├── repeat_string()
│   ├── pad_left_by_count()
│   ├── pad_right_by_count()
│   ├── pad_both_by_count()
│   ├── pad_left_to_width()
│   ├── pad_right_to_width()
│   ├── pad_center_to_width()
│   ├── pad_left()
│   ├── pad_right()
│   ├── pad_center()
│   ├── string_pad()
│   ├── generate_random_string()
│   ├── calculate_string_similarity()
│   ├── format_timestamp()
│   ├── quote()
│   └── strToAppLink()
│
├── 6. 路径、文件、目录工具
│   ├── PathU()
│   ├── path_info()
│   ├── append_datetime_to_filename()
│   ├── append_datetime()
│   ├── output_datetime()
│   ├── open_current_path()
│   ├── get_selected_files()
│   ├── GetSelectedFiles()
│   ├── same_folder()
│   ├── move_in_one_same_named_folder()
│   ├── move_files_to_parent()
│   ├── rename_folder_from_clipboard()
│   ├── rename_folder_from_clipboard_pang()
│   ├── rename_files_from_clipboard()
│   ├── rename_files_from_clipboard_pang()
│   ├── MoveFilesAndFolders()
│   ├── run_unzip()
│   ├── run_zip()
│   ├── is_compressed()
│   ├── paste_files_from_clipboard()
│   ├── paste_files_inbox()
│   ├── find_file()
│   ├── open_recycle_bin()
│   ├── open_device_manager()
│   ├── open_control_panel()
│   ├── open_task_manager()
│   ├── open_system_information()
│   ├── open_event_viewer()
│   ├── open_services_manager()
│   ├── open_computer_management()
│   ├── open_disk_management()
│   ├── open_network_connections()
│   ├── open_user_account_control_settings()
│   ├── open_firewall_settings()
│   ├── open_power_options()
│   ├── open_sound_settings()
│   ├── open_display_settings()
│   ├── open_printers_and_faxes()
│   ├── open_registry_editor()
│   ├── open_command_prompt()
│   ├── open_notepad()
│   └── open_calculator()
│
├── 7. 鼠标、窗口、屏幕与位置检测
│   ├── mouse_get_pos()
│   ├── mouse_is_at()
│   ├── mouse_is_in_rect()
│   ├── mouse_is_at_edge()
│   ├── mouse_is_at_top()
│   ├── mouse_is_at_bottom()
│   ├── mouse_is_at_left()
│   ├── mouse_is_at_right()
│   ├── mouse_is_at_corner()
│   ├── mouse_is_over_window()
│   ├── mouse_is_over_active_window()
│   ├── mouse_get_window_under()
│   ├── mouse_is_over_control()
│   ├── mouse_get_control_under()
│   ├── mouse_get_monitor()
│   ├── mouse_is_on_monitor()
│   ├── mouse_distance_from()
│   ├── mouse_is_moving()
│   ├── mouse_is_idle()
│   ├── mouse_normalize_coords()
│   ├── mouse_is_in_client_area()
│   ├── mouse_button_is_down()
│   ├── mouse_any_button_down()
│   ├── mouse_quick_edge_check()
│   ├── mouse_cached_position()
│   ├── MouseIsOver()
│   ├── mouse_is_over()
│   ├── mouse_is_over_taskbar()
│   ├── mouse_is_top()
│   ├── mouse_is_right()
│   └── mouse_is_left()
│
├── 8. 输入法、CapsLock、键盘布局
│   ├── get_current_layout_id()
│   ├── get_keyboard_layout()
│   ├── get_ime_conv_mode()
│   ├── set_ime_conv_mode()
│   ├── set_keyboard_layout()
│   ├── switch_to_chinese()
│   ├── switch_to_english()
│   ├── switch_to_english_upper()
│   ├── switch_to_microsoft_chinese()
│   ├── switch_to_microsoft_english()
│   ├── toggle_input_method()
│   ├── toggle_capslcok()
│   ├── set_capslcok_true()
│   ├── set_capslcok_false()
│   ├── confirm_layout()
│   ├── show_keyboard_layout()
│   ├── show_input_method()
│   ├── show_current_status()
│   ├── show_detailed_info()
│   ├── activate_chinese_input()
│   ├── english_input()
│   ├── activate_english_uppercase()
│   ├── reverse_caps_letter()
│   ├── base_on_capslock()
│   ├── send_against_capslock()
│   ├── send_with_capslock()
│   └── send_char()
│
├── 9. 剪贴板系统
│   ├── load_clipboard_history()
│   ├── save_clipboard_history()
│   ├── clip_changed()
│   ├── step_clip()
│   ├── show_clip()
│   ├── quick_paste()
│   ├── paste_clip()
│   ├── paste_clip1()
│   ├── paste_clip2()
│   ├── paste_clip3()
│   ├── paste_clip4()
│   ├── paste_clip5()
│   ├── paste_clip6()
│   ├── paste_clip7()
│   ├── paste_clip8()
│   ├── paste_clip9()
│   ├── paste_clip10()
│   ├── paste_content()
│   ├── send_by_clipboard()
│   ├── send_text_by_clipboard()
│   ├── paste_pure_text()
│   ├── paste_text_only()
│   ├── get_selected_text_by_clipboard()
│   ├── get_via_clipboard()
│   ├── has_text_selected()
│   ├── isObjSelected()
│   ├── BackupClipboardText()
│   ├── copy_format()
│   ├── copy_text_in_ppt()
│   ├── paste_format()
│   └── test_paste()
│
├── 10. 目标执行 / 命令调度系统
│   ├── get_target_map()
│   ├── get_target_type()
│   ├── get_target_info()
│   ├── parse_function_call()
│   ├── parse_function_parameters()
│   ├── clean_parameter()
│   ├── execute_parsed_function()
│   ├── call_function()
│   ├── run_target()
│   ├── run_app()
│   ├── run_app2()
│   ├── run_path()
│   ├── run_server_app()
│   ├── set_global()  [也可归这里，因为它是动态配置器]
│   └── MouseAction("fast", ) / MouseAction("slow", ) / fast_cat() / slow_cat() / long_cat()
│
├── 11. 交互层：热键动作包装
│   ├── func_ctrlx()
│   ├── func_ctrls()
│   ├── func_ctrlw()
│   ├── func_ctrlshiftw()
│   ├── func_wins()
│   ├── func_winw()
│   ├── func_wind()
│   ├── keyfunc_rbutton_superkeyoff()
│   ├── keyfunc_space_enter()
│   ├── keyfunc_ctrl2()
│   ├── send_2_or_delete()
│   ├── send_3_or_backspace()
│   ├── send_pgdn_or_IME()
│   ├── send_pgup_or_esc()
│   ├── send_enter_or_space()
│   ├── smart_enter()
│   └── do_nothing() / noop()
│
├── 12. PowerPoint 自动化
│   ├── split_in_folder()
│   ├── split_in_folder_smart()            [注释旧版]
│   ├── export_jpg_in_folder()
│   ├── process_powerpoint()
│   ├── split_selected_ppt_to_batches_then_single()
│   ├── split_ppt_to_batches_then_single_core()
│   ├── split_batch_to_single_pages()
│   ├── split_batch_to_single_pages_internal()
│   ├── split_ppt_full()                   [注释旧版]
│   ├── split_ppt_full_multi()             [注释旧版]
│   ├── split_batch_ppt_to_single()        [注释旧版]
│   ├── split_to_single_only2()            [注释旧版]
│   ├── split_each_batch_to_single2()      [注释旧版]
│   ├── process_single_ppt()               [注释旧版]
│   ├── getdate_914()
│   ├── convert_date_format()
│   ├── append_datetime_to_filename()      [也可归文件工具]
│   ├── rename_in_eagle()
│   ├── send_ppt_code_1()
│   └── ppt_code_1
│
├── 13. Excel / Word / Office 辅助
│   ├── cmdb_update_excel_k_column()
│   ├── save_as()
│   ├── copy_format()
│   ├── paste_format()
│   ├── open_current_path()   [Office 特化部分]
│   └── copy_text_in_ppt()
│
├── 14. 网络 / 设备 / 系统检查
│   ├── isServerReachable()
│   ├── isWebsiteReachable()
│   ├── getShortcutTarget()
│   ├── test_ditto()
│   ├── test_generate_random_text()
│   ├── generate_random_text()
│   ├── generate_random_text2()
│   ├── generate_random_paragraph()
│   ├── generate_random_paragraph2()
│   ├── join_array()
│   └── random text / paragraph helpers
│
├── 15. 数学 / 统计 / 视觉辅助
│   ├── pixel_to_point()
│   ├── Clamp()
│   └── calculate_string_similarity()   [也可归字符串工具]
│
├── 16. 屏幕截图 / 区域选择 / 图片识别
│   ├── screenshot()
│   ├── snipaste()
│   ├── get_screen_region()
│   ├── get_screen_region_lbutton()
│   ├── test_get_screen_region()
│   ├── ClickPicture()
│   ├── ClickPosition()
│   ├── GetPicturePosition()
│   ├── SetCursorToWait()
│   ├── RestoreDefaultCursor()
│   └── screenshot-related utilities
│
├── 17. 窗口控制
│   ├── close_window_under_mouse()
│   ├── min_window_under_mouse()
│   ├── close_window()
│   ├── min_window()
│   ├── max_window()
│   ├── ctrlw_close_window()
│   ├── close_window_in_taskbar()
│   ├── left_window()
│   ├── right_window()
│   ├── show_win_info()
│   └── WinGet()
│
├── 18. 开关 / 退出 / 重载 / 挂起
│   ├── save_reload()
│   ├── reload_me()
│   ├── suspend_me()
│   ├── toggle_suspend()
│   ├── set_superkey_false()
│   ├── set_super_power_false()
│   ├── set_super_power_idol_false()
│   └── ExitApp 相关热键
│
├── 19. 其他实用杂项
│   ├── find_file()
│   ├── test_ditto()
│   ├── write_json_file()
│   ├── open_recycle_bin()
│   ├── open_* 系统入口
│   ├── generate_random_text*()
│   ├── generate_random_paragraph*()
│   └── Miscellaneous helpers
│
└── 20. 历史版本 / 实验代码 / 注释归档
    ├── 老版 IME 控制方案
    ├── 老版剪贴板 GUI
    ├── 老版 PPT 拆分方案
    ├── 老版 split_ppt / export_jpg 方案
    ├── 老版 get_selected_files 方案
    ├── 老版字符串生成方案
    └── 其他临时代码