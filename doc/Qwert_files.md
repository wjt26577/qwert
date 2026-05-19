Qwert/
├── qwert.ahk                ; 主入口
├── settings/
│   ├── settings.json
│   ├── clipboard_history.json
│   ├── layout_frames.json
│   └── preset_layouts.json
├── config/
│   ├── run_config.json
│   ├── run_config.json
│   ├── key_config.json
│   ├── text_config.json
│   ├── mso_config.json
│   ├── func_config.json
│   └── private_config.json  ; 私有配置（不公开）
├── lib/
│   ├── core/
│   │   ├── global_vars.ahk
│   │   ├── state_manager.ahk
│   │   ├── dispatcher.ahk
│   │   └── bootstrap.ahk
│   │
│   ├── utils/
│   │   ├── path_utils.ahk
│   │   ├── string_utils.ahk
│   │   ├── color_utils.ahk
│   │   ├── mouse_utils.ahk
│   │   ├── window_utils.ahk
│   │   └── time_utils.ahk
│   │
│   ├── ui/
│   │   ├── notify.ahk
│   │   ├── tooltip_utils.ahk
│   │   └── clip_gui.ahk
│   │
│   ├── clipboard/
│   │   ├── clipboard_history.ahk
│   │   ├── clipboard_actions.ahk
│   │   └── clipboard_hooks.ahk
│   │
│   ├── input/
│   │   ├── ime_utils.ahk
│   │   ├── capslock_utils.ahk
│   │   └── hotkey_input.ahk
│   │
│   ├── file/
│   │   ├── file_utils.ahk
│   │   ├── rename_utils.ahk
│   │   ├── archive_utils.ahk
│   │   └── explorer_actions.ahk
│   │
│   ├── office/
│   │   ├── ppt_tools.ahk
│   │   ├── excel_tools.ahk
│   │   ├── word_tools.ahk
│   │   └── office_actions.ahk
│   │
│   ├── screen/
│   │   ├── screenshot.ahk
│   │   ├── region_select.ahk
│   │   └── image_utils.ahk
│   │
│   ├── app/
│   │   ├── app_launch.ahk
│   │   ├── app_switch.ahk
│   │   └── run_target.ahk
│   │
│   ├── system/
│   │   ├── system_actions.ahk
│   │   ├── power_actions.ahk
│   │   └── shell_actions.ahk
│   │
│   ├── textgen/
│   │   ├── random_text.ahk
│   │   └── random_paragraph.ahk
│   │
│   └── experimental/
│       ├── old_ime.ahk
│       ├── old_ppt.ahk
│       └── old_clipboard_gui.ahk
│
└── docs/
    ├── README.md
    ├── CONFIG.md
    ├── HOTKEYS.md
    └── DEVELOPMENT.md