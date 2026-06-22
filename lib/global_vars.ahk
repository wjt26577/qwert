
; #region ahk_group

GroupAdd "taskbar_window", "ahk_class Shell_TrayWnd" 
GroupAdd "taskbar_window", "ahk_class Shell_SecondaryTrayWnd"

GroupAdd "chinese_app", "ahk_exe chrome.exe"
GroupAdd "chinese_app", "ahk_class XLMAIN"
GroupAdd "chinese_app", "ahk_class OpusApp"
GroupAdd "chinese_app", "ahk_class WeChatMainWnorPC"

GroupAdd "english_app", "ahk_exe Code.exe"

GroupAdd "chinese_english", "ahk_group english_app"
GroupAdd "chinese_english", "ahk_group chinese_app"

GroupAdd "black_office", "Excel", , "-"
GroupAdd "black_office", "PowerPoint", , "-"
GroupAdd "black_office", "Word", , "-"
GroupAdd "black_office", "Visio Professional", , "-"

GroupAdd "group_office", "ahk_class XLMAIN"
GroupAdd "group_office", "ahk_class OpusApp"	 
GroupAdd "group_office", "ahk_class PPTFrameClass" 
GroupAdd "group_office", "ahk_class VISIOA"

GroupAdd "candidate_window", "ahk_class wetype.flutter.setting"
GroupAdd "candidate_window", "ahk_class SoPY_Comp"	
GroupAdd "candidate_window", "ahk_class BaseGui"		
GroupAdd "candidate_window", "ahk_class ApplicationFrameWindow"	

GroupAdd "system_group", "ahk_class WorkerW"
GroupAdd "system_group", "ahk_class Progman"

GroupAdd "taskbar_group", "ahk_class Shell_TrayWnd"
GroupAdd "taskbar_group", "ahk_class Shell_SecondaryTrayWnd"
GroupAdd "taskbar_group", "ahk_class NotifyIconOverflowWindow"

GroupAdd "text_app", "ahk_exe Notepad.exe"

GroupAdd "ctrlw_group", "ahk_exe clash-verge.exe"
GroupAdd "ctrlw_group", "ahk_exe Evernote.exe"
GroupAdd "ctrlw_group", "ahk_exe mailmaster.exe"
GroupAdd "ctrlw_group", "ahk_exe IDMan.exe"
GroupAdd "ctrlw_group", "ahk_exe WindowSpyDpi.exe"
GroupAdd "ctrlw_group", "ahk_exe Ditto.exe"
GroupAdd "ctrlw_group", "ahk_exe Qianwen.exe"
GroupAdd "ctrlw_group", "ahk_exe Doubao.exe"
GroupAdd "ctrlw_group", "ahk_exe ApplicationFrameHost.exe"
GroupAdd "ctrlw_group", "ahk_exe 哔哩哔哩.exe"
GroupAdd "ctrlw_group", "ahk_exe WeChatAppEx.exe"
GroupAdd "ctrlw_group", "Excel", , "-"
GroupAdd "ctrlw_group", "PowerPoint", , "-"
GroupAdd "ctrlw_group", "Word", , "-"
GroupAdd "ctrlw_group", "Visio Professional", , "-"

GroupAdd "video_group", "ahk_exe 哔哩哔哩.exe"

; #endregion


; =========================
; 2. 全局状态
; =========================
global clipboard_history := []
global max_history := 10
global text_mode := false
global screen_left := 0
global screen_right := 0
path_ai_prompt_1 := "d:\obsidian\knowledge_base\04_资源\ai_prompts\yinlv_ai_prompts\音律通用提示词.md"
path_ai_prompt_2 := "d:\obsidian\knowledge_base\04_资源\ai_prompts\yinlv_ai_prompts\音律重建总控提示词.md"
path_ai_prompt_3 := "d:\obsidian\knowledge_base\04_资源\ai_prompts\yinlv_ai_prompts\音律PPT复刻提示词.md"

; =========================
; 3. 资源路径
; =========================
global BASE_DIR := A_ScriptDir
global SETTINGS_DIR := BASE_DIR . "\settings"
global IMG_DIR := BASE_DIR . "\img"

global SETTINGS_FILE := SETTINGS_DIR . "\settings.json"
global clipboard_history_FILE := SETTINGS_DIR . "\clipboard_history.json"
global LAYOUT_FRAMES_FILE := SETTINGS_DIR . "\layout_frames.json"
global PRESET_LAYOUTS_FILE := SETTINGS_DIR . "\preset_layouts.json"

global PATH_LAYOUT_JSON         := "\settings\layout_frames.json"
global PATH_PRESET_LAYOUTS_JSON := "\settings\preset_layouts.json"
global CMD_DESKTOP_NOTES_1 := "Snipaste paste --files D:\system\desktop_notes\vscode_shortcut.png"
global CMD_DESKTOP_NOTES_2 := "Snipaste paste --files D:\system\desktop_notes\snipaste_help_1.png"
global CMD_DESKTOP_NOTES_3 := "Snipaste paste --files D:\system\desktop_notes\snipaste_help_2.png"

; #region constants
  
; =========================
; 6. 字体列表
; =========================
global ch_fonts := [
    "MiSans light",
    "阿里巴巴普惠体 3.0 55 Regular",
    "微软雅黑",
    "思源宋体 CN",
    "方正清刻本悦宋简体",
    "MiSans Heavy",
    "字酷堂清楷 简",
    "字体圈欣意冠黑体"
]

global en_fonts := [
    "Fira Code light",
    "Gilroy Regular",
    "微软雅黑",
    "思源宋体 CN",
    "方正清刻本悦宋体",
    "喜鹊招牌体",
    "字酷堂清楷 简",
    "字体圈欣意冠黑体"
]

; ============================================================
; 颜色名称 → RGB 对照表
; ============================================================
system_color_map := Map(
    "Black",   0x000000, "Silver",  0xC0C0C0,
    "Gray",    0x808080, "White",   0xFFFFFF,
    "Maroon",  0x800000, "Red",     0xFF0000,
    "Purple",  0x800080, "Fuchsia", 0xFF00FF,
    "Green",   0x008000, "Lime",    0x00FF00,
    "Olive",   0x808000, "Yellow",  0xFFFF00,
    "Navy",    0x000080, "Blue",    0x0000FF,
    "Teal",    0x008080, "Aqua",    0x00FFFF,
    "Orange",  0xFFA500, "Pink",    0xFFC0CB,
    "Coral",   0xFF7F50, "Gold",    0xFFD700,
    "Cyan",    0x00FFFF, "Magenta", 0xFF00FF
)



color_map := Map(
    1, "red",
    2, "0xFFA500",
    3, "Yellow",
    4, "Green",
    5, "aqua",
    6, "blue",
    7, "purple",
    8, "silver",
    9, "0x4B0082",
    10, "gray"
)

selection_type_map := Map(
    0, "none",
    1, "slides",
    2, "shapes",
    3, "text"
)

SHAPE_TYPE_MAP := Map(
    ; 数值, 常量名称,                              ; 说明
    30, "mso3DModel",                            ; 3D 模型
    1, "msoAutoShape",                           ; 自选图形
    2, "msoCallout",                             ; 标注
    20, "msoCanvas",                             ; 画布
    3, "msoChart",                               ; 图表
    4, "msoComment",                             ; 评论
    27, "msoContentApp",                         ; 内容 Office 加载项
    21, "msoDiagram",                            ; 图示
    7, "msoEmbeddedOLEObject",                   ; 嵌入式 OLE 对象
    8, "msoFormControl",                         ; 窗体控件
    5, "msoFreeform",                            ; 任意多边形
    28, "msoGraphic",                            ; 图形
    6, "msoGroup",                               ; 组
    24, "msoIgxGraphic",                         ; SmartArt 图形
    22, "msoInk",                                ; 墨迹
    23, "msoInkComment",                         ; 墨迹批注
    9, "msoLine",                                ; 折线图
    31, "msoLinked3DModel",                      ; 链接的 3D 模型
    29, "msoLinkedGraphic",                      ; 链接的图形
    10, "msoLinkedOLEObject",                    ; 链接 OLE 对象
    11, "msoLinkedPicture",                      ; 链接图片
    16, "msoMedia",                              ; 媒体
    12, "msoOLEControlObject",                   ; OLE 控件对象
    13, "msoPicture",                            ; 图片
    14, "msoPlaceholder",                        ; 占位符
    18, "msoScriptAnchor",                       ; 脚本定位标记
    -2, "msoShapeTypeMixed",                     ; 混合形状类型
    25, "msoSlicer",                             ; 切片器
    19, "msoTable",                              ; 表格
    17, "msoTextBox",                            ; 文本框
    15, "msoTextEffect",                         ; 文本效果
    26, "msoWebVideo"                            ; Web 视频
)

TEXT_4 := [
    "智领未来", "数实融合", "新质生产力", "绿色低碳",
    "自主创新", "全球布局", "提质增效", "降本增效",
    "敏捷迭代", "深度赋能", "生态共赢", "价值重塑",
    "精准洞察", "全域营销", "沉浸体验", "无缝协同",
    "智慧决策", "韧性增长", "安全可控", "开源节流",
    "跨界融合", "场景创新", "数据驱动", "云端协同",
    "人机协作", "虚实共生", "零碳转型", "循环经济",
    "高端智造", "专精特新", "品牌焕新", "用户至上",
    "长期主义", "稳健经营", "风险对冲", "合规先行",
    "人才强企", "文化铸魂", "组织激活", "效能飞跃",
    "技术突破", "产业升级", "市场深耕", "渠道下沉",
    "内容种草", "私域运营", "直播赋能", "社交裂变",
    "AI 赋能", "算力爆发", "算法优化", "模型迭代",
    "量子前沿", "生物制造", "空天经济", "深海探测",
    "能源革命", "储能创新", "氢能应用", "光伏增效",
    "智能网联", "车路协同", "低空经济", "无人配送",
    "元宇宙境", "数字孪生", "区块链网", "隐私计算",
    "柔性生产", "定制服务", "即时响应", "极致体验",
    "开放合作", "共享经济", "平台生态", "链式反应",
    "破局重生", "弯道超车", "换道领跑", "降维打击",
    "精细运营", "精益管理", "标准引领", "质量强国",
    "乡村振兴", "共同富裕", "民生福祉", "社会和谐",
    "文化自信", "国潮崛起", "东方美学", "匠心传承",
    "健康中国", "智慧医疗", "康养结合", "生命科学",
    "教育公平", "终身学习", "技能提升", "产教融合",
    "法治社会", "诚信体系", "公平正义", "平安建设",
    "美丽中国", "生态文明", "人与自然", "和谐共生"
]

TEXT_8 := [
    "数字转型实施方案", "项目管理优化策略", "战略规划部署方案", "宏观趋势分析报告",
    "竞争格局应对策略", "产业升级指导纲要", "区域发展布局规划", "长期愿景实施路径",
    "短期目标行动指南", "资源配置优化方案", "运营效能提升方案", "供应链路协同优化",
    "库存管理精简策略", "物流配送提速计划", "生产流程再造工程", "精益管理推行方案",
    "风险防控体系建设", "合规审计整改报告", "市场分析研究报告", "品牌建设推广方案",
    "营销渠道拓展计划", "客户画像分析洞察", "竞品对标分析报告", "公关危机处理预案",
    "产品创新设计方案", "技术架构升级方案", "数据安全防护措施", "用户体验改进计划",
    "团队建设发展规划", "人才培养实施方案", "成本控制优化策略", "效率提升行动计划",
    "质量管控体系构建", "业务流程改进方案", "客户服务提升计划", "技术架构升级方案",
    "人才培养实施方案", "质量管控体系构建", "成本控制优化策略", "效率提升行动计划",
    "品牌建设推广方案", "数据安全防护措施", "用户体验改进计划", "战略规划部署方案",
    "宏观趋势分析报告", "竞争格局应对策略", "产业升级指导纲要", "区域发展布局规划",
    "长期愿景实施路径", "短期目标行动指南", "资源配置优化方案", "顶层设计构建方案",
    "年度经营计划大纲", "运营效能提升方案", "供应链条协同优化", "库存管理精简策略",
    "物流配送提速计划", "生产流程再造工程", "精益管理推行方案", "风险防控体系建设"
]

TEXT_20 := [
    "深化数字化转型战略，全面推动企业高质量发展进程",
    "优化供应链管理流程，显著提升整体运营效率与效益",
    "加强核心技术自主研发，构建可持续创新竞争壁垒",
    "拓展全球市场业务布局，实现多元化国际发展战略",
    "完善人才培养激励机制，打造高素质专业化精英团队",
    "推进绿色低碳生产模式，践行社会责任与可持续发展",
    "强化风险防控体系建设，确保企业经营安全稳健运行",
    "提升客户服务体验标准，树立行业领先品牌形象口碑",
    "加速智能工厂建设步伐，引领制造业转型升级新方向",
    "整合线上线下营销渠道，最大化提升市场份额占有率",
    "落实精细化管理措施，有效控制成本并提高利润率",
    "探索新兴商业模式应用，激发组织内部活力与创新力",
    "构建开放共赢生态平台，携手合作伙伴共创美好未来",
    "加大研发投入力度，突破关键领域技术瓶颈与制约",
    "优化资产配置结构，实现资本运作效率最大化目标",
    "建立数据驱动决策机制，精准洞察市场趋势与变化",
    "弘扬企业文化核心价值观，凝聚全员奋斗共识与力量",
    "实施品牌国际化战略，提升中国智造全球影响力",
    "完善售后服务网络体系，快速响应客户需求与反馈",
    "推动产业协同创新发展，构建现代化经济体系新格局"
]

TEXT_30 := [
    "深化数字化转型战略全面布局，以创新驱动引领企业高质量可持续发展新征程",
    "优化全链路供应链管理流程，显著提升整体运营效率并最大化经济效益产出",
    "加强核心技术自主研发创新力度，构建具有全球竞争力的可持续技术壁垒",
    "拓展全球化市场业务战略布局，实现多元化国际发展战略与市场份额双增长",
    "完善全方位人才培养激励机制，打造高素质专业化精英团队以赋能组织发展",
    "深入推进绿色低碳生产模式转型，积极践行社会责任并实现可持续生态发展",
    "强化多层次风险防控体系建设，确保企业在复杂环境中经营安全稳健运行",
    "全面提升客户服务体验标准体系，树立行业领先品牌形象与卓越口碑影响力",
    "加速智能化数字工厂建设步伐，引领传统制造业转型升级与高质量发展方向",
    "深度整合线上线下全渠道营销，最大化提升市场份额占有率与品牌渗透能力",
    "全面落实精细化运营管理措施，有效控制成本支出并显著提高净利润率水平",
    "积极探索新兴商业模式创新应用，充分激发组织内部活力与持续创新驱动力",
    "构建开放共赢的产业生态平台，携手合作伙伴共创美好未来与商业价值新高",
    "加大前沿技术研发投入支持力度，突破关键领域技术瓶颈制约并掌握主动权",
    "优化多元化资产配置结构组合，实现资本运作效率最大化与企业价值目标",
    "建立数据驱动型科学决策机制，精准洞察市场趋势变化并快速响应客户需求",
    "弘扬企业文化核心价值观理念，凝聚全员奋斗共识与力量以推动战略目标达成",
    "全面实施品牌国际化发展战略，显著提升中国智造在全球市场的影响力地位",
    "完善覆盖全国售后服务网络体系，快速响应客户需求反馈并提升满意度指标",
    "推动产业链协同创新发展模式，构建现代化经济体系新格局并增强抗风险能力"
]





mso_config := Map(
    "ObjectsAlignLeftSmart", Map("payload", "ObjectsAlignLeftSmart", "label", "左对齐", "kind", "mso"),
    "ObjectsAlignCenterHorizontalSmart", Map("payload", "ObjectsAlignCenterHorizontalSmart", "label", "水平居中", "kind", "mso"),
    "ObjectsAlignRightSmart", Map("payload", "ObjectsAlignRightSmart", "label", "右对齐", "kind", "mso"),
    "ObjectsAlignTopSmart", Map("payload", "ObjectsAlignTopSmart", "label", "顶部对齐", "kind", "mso"),
    "ObjectsAlignMiddleVerticalSmart", Map("payload", "ObjectsAlignMiddleVerticalSmart", "label", "垂直居中", "kind", "mso"),
    "ObjectsAlignBottomSmart", Map("payload", "ObjectsAlignBottomSmart", "label", "底部对齐", "kind", "mso"),
    "AlignObjects", Map("payload", "AlignObjects", "label", "对齐对象菜单", "kind", "mso"),
    "ObjectsDistributeLeft", Map("payload", "ObjectsDistributeLeft", "label", "横向分布", "kind", "mso"),
    "ObjectsDistributeCenter", Map("payload", "ObjectsDistributeCenter", "label", "横向居中分布", "kind", "mso"),
    "ObjectsDistributeRight", Map("payload", "ObjectsDistributeRight", "label", "横向右分布", "kind", "mso"),
    "ObjectsDistributeTop", Map("payload", "ObjectsDistributeTop", "label", "纵向分布", "kind", "mso"),
    "ObjectsDistributeMiddle", Map("payload", "ObjectsDistributeMiddle", "label", "纵向居中分布", "kind", "mso"),
    "ObjectsDistributeBottom", Map("payload", "ObjectsDistributeBottom", "label", "纵向下分布", "kind", "mso"),
    "ObjectSendToBack", Map("payload", "ObjectSendToBack", "label", "置于底层", "kind", "mso"),
    "ObjectBringToFront", Map("payload", "ObjectBringToFront", "label", "置于顶层", "kind", "mso"),
    "ObjectSendBackward", Map("payload", "ObjectSendBackward", "label", "下移一层", "kind", "mso"),
    "ObjectBringForward", Map("payload", "ObjectBringForward", "label", "上移一层", "kind", "mso"),
    "SelectionPane", Map("payload", "SelectionPane", "label", "选择窗格", "kind", "mso"),
    "GroupShapes", Map("payload", "GroupShapes", "label", "组合", "kind", "mso"),
    "UngroupShapes", Map("payload", "UngroupShapes", "label", "取消组合", "kind", "mso"),
    "RegroupShapes", Map("payload", "RegroupShapes", "label", "重新组合", "kind", "mso"),
    "RotateRight", Map("payload", "RotateRight", "label", "向右旋转90度", "kind", "mso"),
    "RotateLeft", Map("payload", "RotateLeft", "label", "向左旋转90度", "kind", "mso"),
    "RotateFlipHorizontal", Map("payload", "RotateFlipHorizontal", "label", "水平翻转", "kind", "mso"),
    "RotateFlipVertical", Map("payload", "RotateFlipVertical", "label", "垂直翻转", "kind", "mso"),
    "RotateFlipHorizontalFlipVertical", Map("payload", "RotateFlipHorizontalFlipVertical", "label", "水平垂直翻转", "kind", "mso"),
    "RotateFlipHorizontalRotate90Right", Map("payload", "RotateFlipHorizontalRotate90Right", "label", "右旋+水平翻转", "kind", "mso"),
    "RotateFlipVerticalRotate90Right", Map("payload", "RotateFlipVerticalRotate90Right", "label", "右旋+垂直翻转", "kind", "mso"),
    "EditShape", Map("payload", "EditShape", "label", "编辑顶点", "kind", "mso"),
    "ChangeShape", Map("payload", "ChangeShape", "label", "更改形状", "kind", "mso"),
    "InkToShape", Map("payload", "InkToShape", "label", "墨迹转换为形状", "kind", "mso"),
    "ConvertToFreeform", Map("payload", "ConvertToFreeform", "label", "转换为自由曲线", "kind", "mso"),
    "AlignLeft", Map("payload", "AlignLeft", "label", "文本左对齐", "kind", "mso"),
    "AlignCenter", Map("payload", "AlignCenter", "label", "文本居中", "kind", "mso"),
    "AlignRight", Map("payload", "AlignRight", "label", "文本右对齐", "kind", "mso"),
    "AlignTextJustify", Map("payload", "AlignTextJustify", "label", "文本两端对齐", "kind", "mso"),
    "AlignTextDistribute", Map("payload", "AlignTextDistribute", "label", "文本分散对齐", "kind", "mso"),
    "JustifyLow", Map("payload", "JustifyLow", "label", "文本低对齐", "kind", "mso"),
    "AlignTextTop", Map("payload", "AlignTextTop", "label", "文本顶端对齐", "kind", "mso"),
    "AlignTextMiddle", Map("payload", "AlignTextMiddle", "label", "文本中部对齐", "kind", "mso"),
    "AlignTextBottom", Map("payload", "AlignTextBottom", "label", "文本底端对齐", "kind", "mso"),
    "TextDirectionRotate270Normal", Map("payload", "TextDirectionRotate270Normal", "label", "文字方向：竖排", "kind", "mso"),
    "TextDirectionRotate90Normal", Map("payload", "TextDirectionRotate90Normal", "label", "文字方向：竖排(反向)", "kind", "mso"),
    "TextDirectionHorizontal", Map("payload", "TextDirectionHorizontal", "label", "文字方向：横排", "kind", "mso"),
    "Bullets", Map("payload", "Bullets", "label", "项目符号", "kind", "mso"),
    "Numbering", Map("payload", "Numbering", "label", "编号", "kind", "mso"),
    "LineSpacing", Map("payload", "LineSpacing", "label", "行距", "kind", "mso"),
    "IncreaseIndent", Map("payload", "IncreaseIndent", "label", "增加缩进", "kind", "mso"),
    "DecreaseIndent", Map("payload", "DecreaseIndent", "label", "减少缩进", "kind", "mso"),
    "HangingIndent", Map("payload", "HangingIndent", "label", "悬挂缩进", "kind", "mso"),
    "Font", Map("payload", "Font", "label", "字体设置", "kind", "mso"),
    "FontSize", Map("payload", "FontSize", "label", "字号", "kind", "mso"),
    "FontSizeGrow", Map("payload", "FontSizeGrow", "label", "增大字号", "kind", "mso"),
    "FontSizeShrink", Map("payload", "FontSizeShrink", "label", "减小字号", "kind", "mso"),
    "Bold", Map("payload", "Bold", "label", "加粗", "kind", "mso"),
    "Italic", Map("payload", "Italic", "label", "倾斜", "kind", "mso"),
    "Underline", Map("payload", "Underline", "label", "下划线", "kind", "mso"),
    "Shadow", Map("payload", "Shadow", "label", "阴影", "kind", "mso"),
    "Strikethrough", Map("payload", "Strikethrough", "label", "删除线", "kind", "mso"),
    "ChangeCase", Map("payload", "ChangeCase", "label", "更改大小写", "kind", "mso"),
    "Subscript", Map("payload", "Subscript", "label", "下标", "kind", "mso"),
    "Superscript", Map("payload", "Superscript", "label", "上标", "kind", "mso"),
    "ClearFormatting", Map("payload", "ClearFormatting", "label", "清除格式", "kind", "mso"),
    "Highlight", Map("payload", "Highlight", "label", "文本高亮", "kind", "mso"),
    "FontColor", Map("payload", "FontColor", "label", "字体颜色", "kind", "mso"),
    "ShapeFillColorPicker", Map("payload", "ShapeFillColorPicker", "label", "形状填充", "kind", "mso"),
    "ShapeOutlineColorPicker", Map("payload", "ShapeOutlineColorPicker", "label", "形状轮廓", "kind", "mso"),
    "ShapeEffects", Map("payload", "ShapeEffects", "label", "形状效果", "kind", "mso"),
    "QuickStylesGallery", Map("payload", "QuickStylesGallery", "label", "快速样式", "kind", "mso"),
    "LineStyle", Map("payload", "LineStyle", "label", "线条样式", "kind", "mso"),
    "LineWeight", Map("payload", "LineWeight", "label", "线条粗细", "kind", "mso"),
    "LineDashes", Map("payload", "LineDashes", "label", "虚线线型", "kind", "mso"),
    "LineArrowheads", Map("payload", "LineArrowheads", "label", "箭头样式", "kind", "mso"),
    "PictureChange", Map("payload", "PictureChange", "label", "更改图片", "kind", "mso"),
    "PictureReset", Map("payload", "PictureReset", "label", "重设图片", "kind", "mso"),
    "Crop", Map("payload", "Crop", "label", "裁剪", "kind", "mso"),
    "CropToShape", Map("payload", "CropToShape", "label", "裁剪为形状", "kind", "mso"),
    "AspectRatio", Map("payload", "AspectRatio", "label", "纵横比", "kind", "mso"),
    "CompressPictures", Map("payload", "CompressPictures", "label", "压缩图片", "kind", "mso"),
    "PictureColor", Map("payload", "PictureColor", "label", "图片颜色", "kind", "mso"),
    "PictureArtisticEffects", Map("payload", "PictureArtisticEffects", "label", "艺术效果", "kind", "mso"),
    "PictureCorrections", Map("payload", "PictureCorrections", "label", "图片更正", "kind", "mso"),
    "PictureBorder", Map("payload", "PictureBorder", "label", "图片边框", "kind", "mso"),
    "RemoveBackground", Map("payload", "RemoveBackground", "label", "删除背景", "kind", "mso"),
    "OnlinePicture", Map("payload", "OnlinePicture", "label", "联机图片", "kind", "mso"),
    "InsertScreenshot", Map("payload", "InsertScreenshot", "label", "屏幕截图", "kind", "mso"),
    "PictureCropAspectRatio1To1", Map("payload", "PictureCropAspectRatio1To1", "label", "图片裁剪宽高比1:1", "kind", "mso"),
    "PictureCropAspectRatio2To3", Map("payload", "PictureCropAspectRatio2To3", "label", "图片裁剪宽高比2:3", "kind", "mso"),
    "PictureCropAspectRatio3To4", Map("payload", "PictureCropAspectRatio3To4", "label", "图片裁剪宽高比3:4", "kind", "mso"),
    "PictureCropAspectRatio3To5", Map("payload", "PictureCropAspectRatio3To5", "label", "图片裁剪宽高比3:5", "kind", "mso"),
    "PictureCropAspectRatio4To5", Map("payload", "PictureCropAspectRatio4To5", "label", "图片裁剪宽高比4:5", "kind", "mso"),
    "PictureCropAspectRatio3To2", Map("payload", "PictureCropAspectRatio3To2", "label", "图片裁剪宽高比3:2", "kind", "mso"),
    "PictureCropAspectRatio4To3", Map("payload", "PictureCropAspectRatio4To3", "label", "图片裁剪宽高比4:3", "kind", "mso"),
    "PictureCropAspectRatio5To3", Map("payload", "PictureCropAspectRatio5To3", "label", "图片裁剪宽高比5:3", "kind", "mso"),
    "PictureCropAspectRatio5To4", Map("payload", "PictureCropAspectRatio5To4", "label", "图片裁剪宽高比5:4", "kind", "mso"),
    "PictureCropAspectRatio16To9", Map("payload", "PictureCropAspectRatio16To9", "label", "图片裁剪宽高比16:9", "kind", "mso"),
    "PictureCropAspectRatio16To10", Map("payload", "PictureCropAspectRatio16To10", "label", "图片裁剪宽高比16:10", "kind", "mso"),
    "PictureFillCrop", Map("payload", "PictureFillCrop", "label", "图片填充裁剪", "kind", "mso"),
    "PictureFitCrop", Map("payload", "PictureFitCrop", "label", "图片适应裁剪", "kind", "mso"),
    "NudgeUp", Map("payload", "NudgeUp", "label", "向上微调", "kind", "mso"),
    "NudgeDown", Map("payload", "NudgeDown", "label", "向下微调", "kind", "mso"),
    "NudgeLeft", Map("payload", "NudgeLeft", "label", "向左微调", "kind", "mso"),
    "NudgeRight", Map("payload", "NudgeRight", "label", "向右微调", "kind", "mso"),
    "SizeUp", Map("payload", "SizeUp", "label", "向上增大", "kind", "mso"),
    "SizeDown", Map("payload", "SizeDown", "label", "向下减小", "kind", "mso"),
    "SizeLeft", Map("payload", "SizeLeft", "label", "向左增大", "kind", "mso"),
    "SizeRight", Map("payload", "SizeRight", "label", "向右减小", "kind", "mso"),
    "SizeObjectDialog", Map("payload", "SizeObjectDialog", "label", "大小和位置", "kind", "mso"),
    "SnapToGrid", Map("payload", "SnapToGrid", "label", "对齐网格", "kind", "mso"),
    "SnapToShapes", Map("payload", "SnapToShapes", "label", "对齐对象", "kind", "mso"),
    "Gridlines", Map("payload", "Gridlines", "label", "网格线", "kind", "mso"),
    "Guides", Map("payload", "Guides", "label", "参考线", "kind", "mso"),
    "ViewZoom", Map("payload", "ViewZoom", "label", "缩放", "kind", "mso"),
    "FitToWindow", Map("payload", "FitToWindow", "label", "适应窗口大小", "kind", "mso"),
    "ZoomToSelection", Map("payload", "ZoomToSelection", "label", "缩放至所选内容", "kind", "mso"),
    "InsertRowsAbove", Map("payload", "InsertRowsAbove", "label", "在上方插入行", "kind", "mso"),
    "InsertRowsBelow", Map("payload", "InsertRowsBelow", "label", "在下方插入行", "kind", "mso"),
    "InsertColumnsLeft", Map("payload", "InsertColumnsLeft", "label", "在左侧插入列", "kind", "mso"),
    "InsertColumnsRight", Map("payload", "InsertColumnsRight", "label", "在右侧插入列", "kind", "mso"),
    "DeleteRows", Map("payload", "DeleteRows", "label", "删除行", "kind", "mso"),
    "DeleteColumns", Map("payload", "DeleteColumns", "label", "删除列", "kind", "mso"),
    "MergeCells", Map("payload", "MergeCells", "label", "合并单元格", "kind", "mso"),
    "SplitCells", Map("payload", "SplitCells", "label", "拆分单元格", "kind", "mso"),
    "DistributeRows", Map("payload", "DistributeRows", "label", "平均分布各行", "kind", "mso"),
    "DistributeColumns", Map("payload", "DistributeColumns", "label", "平均分布各列", "kind", "mso"),
    "TableBorders", Map("payload", "TableBorders", "label", "边框", "kind", "mso"),
    "TableFillColor", Map("payload", "TableFillColor", "label", "底纹", "kind", "mso"),
    "ChartType", Map("payload", "ChartType", "label", "更改图表类型", "kind", "mso"),
    "ChartData", Map("payload", "ChartData", "label", "编辑数据", "kind", "mso"),
    "ChartLayout", Map("payload", "ChartLayout", "label", "图表布局", "kind", "mso"),
    "ChartStyles", Map("payload", "ChartStyles", "label", "图表样式", "kind", "mso"),
    "ChartAxes", Map("payload", "ChartAxes", "label", "坐标轴", "kind", "mso"),
    "ChartGridlines", Map("payload", "ChartGridlines", "label", "网格线", "kind", "mso"),
    "ChartLegend", Map("payload", "ChartLegend", "label", "图例", "kind", "mso"),
    "ChartDataLabels", Map("payload", "ChartDataLabels", "label", "数据标签", "kind", "mso"),
    "ChartTrendline", Map("payload", "ChartTrendline", "label", "趋势线", "kind", "mso"),
    "SmartArtAddShapeAfter", Map("payload", "SmartArtAddShapeAfter", "label", "添加形状(后)", "kind", "mso"),
    "SmartArtAddShapeBefore", Map("payload", "SmartArtAddShapeBefore", "label", "添加形状(前)", "kind", "mso"),
    "SmartArtAddShapeAbove", Map("payload", "SmartArtAddShapeAbove", "label", "添加形状(上)", "kind", "mso"),
    "SmartArtAddShapeBelow", Map("payload", "SmartArtAddShapeBelow", "label", "添加形状(下)", "kind", "mso"),
    "SmartArtChangeHierarchy", Map("payload", "SmartArtChangeHierarchy", "label", "更改层次结构", "kind", "mso"),
    "SmartArtRTL", Map("payload", "SmartArtRTL", "label", "从右向左", "kind", "mso"),
    "SmartArtLarger", Map("payload", "SmartArtLarger", "label", "增大", "kind", "mso"),
    "SmartArtSmaller", Map("payload", "SmartArtSmaller", "label", "减小", "kind", "mso"),
    "SmartArtReset", Map("payload", "SmartArtReset", "label", "重设", "kind", "mso"),
    "HyperlinkInsert", Map("payload", "HyperlinkInsert", "label", "插入超链接", "kind", "mso"),
    "Duplicate", Map("payload", "Duplicate", "label", "复制选中对象", "kind", "mso"),
    "Delete", Map("payload", "Delete", "label", "删除", "kind", "mso"),
    "Copy", Map("payload", "Copy", "label", "复制", "kind", "mso"),
    "Paste", Map("payload", "Paste", "label", "粘贴", "kind", "mso"),
    "PasteSourceFormatting", Map("payload", "PasteSourceFormatting", "label", "保留源格式粘贴", "kind", "mso"),
    "PasteMergeFormatting", Map("payload", "PasteMergeFormatting", "label", "合并格式粘贴", "kind", "mso"),
    "PastePicture", Map("payload", "PastePicture", "label", "粘贴为图片", "kind", "mso"),
    "Save", Map("payload", "Save", "label", "保存", "kind", "mso"),
    "Undo", Map("payload", "Undo", "label", "撤销", "kind", "mso"),
    "Redo", Map("payload", "Redo", "label", "恢复", "kind", "mso"),
    "SelectAll", Map("payload", "SelectAll", "label", "全选", "kind", "mso"),
    "Find", Map("payload", "Find", "label", "查找", "kind", "mso"),
    "Replace", Map("payload", "Replace", "label", "替换", "kind", "mso"),
    "AutoSaveSwitch", Map("payload", "AutoSaveSwitch", "label", "自动保存到OneDriver", "kind", "mso"),
    "FileNewDefault", Map("payload", "FileNewDefault", "label", "新建默认文件", "kind", "mso"),
    "FileOpenUsingBackstage", Map("payload", "FileOpenUsingBackstage", "label", "利用后台视图打开文件", "kind", "mso"),
    "FileSave", Map("payload", "FileSave", "label", "保存文件", "kind", "mso"),
    "FileSendAsAttachment", Map("payload", "FileSendAsAttachment", "label", "当作附件发送", "kind", "mso"),
    "FilePrintQuick", Map("payload", "FilePrintQuick", "label", "快速打印", "kind", "mso"),
    "PrintPreviewAndPrint", Map("payload", "PrintPreviewAndPrint", "label", "打印预览", "kind", "mso"),
    "Spelling", Map("payload", "Spelling", "label", "拼写检查", "kind", "mso"),
    "RedoOrRepeat", Map("payload", "RedoOrRepeat", "label", "重做", "kind", "mso"),
    "SlideShowFromBeginning", Map("payload", "SlideShowFromBeginning", "label", "从头开始播放幻灯片放映", "kind", "mso"),
    "PasteSpecialDialog", Map("payload", "PasteSpecialDialog", "label", "粘贴特殊对话框", "kind", "mso"),
    "Cut", Map("payload", "Cut", "label", "剪切", "kind", "mso"),
    "PasteDuplicate", Map("payload", "PasteDuplicate", "label", "复制并粘贴", "kind", "mso"),
    "FormatPainter", Map("payload", "FormatPainter", "label", "格式刷", "kind", "mso"),
    "ShowClipboard", Map("payload", "ShowClipboard", "label", "显示剪贴板", "kind", "mso"),
    "DuplicateSelectedSlides", Map("payload", "DuplicateSelectedSlides", "label", "复制所选幻灯片", "kind", "mso"),
    "SlidesFromOutline", Map("payload", "SlidesFromOutline", "label", "从大纲生成幻灯片", "kind", "mso"),
    "SlidesReuseSlides", Map("payload", "SlidesReuseSlides", "label", "重用幻灯片", "kind", "mso"),
    "SlideReset", Map("payload", "SlideReset", "label", "重置幻灯片", "kind", "mso"),
    "SectionAdd", Map("payload", "SectionAdd", "label", "添加分区", "kind", "mso"),
    "SectionRename", Map("payload", "SectionRename", "label", "重命名分区", "kind", "mso"),
    "SectionMergeWithPrevious", Map("payload", "SectionMergeWithPrevious", "label", "合并到前一个分区", "kind", "mso"),
    "SectionRemoveAll", Map("payload", "SectionRemoveAll", "label", "移除所有分区", "kind", "mso"),
    "SectionCollapseAll", Map("payload", "SectionCollapseAll", "label", "折叠所有分区", "kind", "mso"),
    "SectionExpandAll", Map("payload", "SectionExpandAll", "label", "展开所有分区", "kind", "mso"),
    "FontSizeIncrease", Map("payload", "FontSizeIncrease", "label", "字体大小增加", "kind", "mso"),
    "FontSizeDecrease", Map("payload", "FontSizeDecrease", "label", "字体大小减少", "kind", "mso"),
    "ClearFormatting", Map("payload", "ClearFormatting", "label", "清除格式", "kind", "mso"),
    "Bold", Map("payload", "Bold", "label", "加粗", "kind", "mso"),
    "Italic", Map("payload", "Italic", "label", "斜体", "kind", "mso"),
    "Underline", Map("payload", "Underline", "label", "下划线", "kind", "mso"),
    "Shadow", Map("payload", "Shadow", "label", "阴影", "kind", "mso"),
    "Strikethrough", Map("payload", "Strikethrough", "label", "删除线", "kind", "mso"),
    "FontColorMoreColorsDialogPowerPoint", Map("payload", "FontColorMoreColorsDialogPowerPoint", "label", "字体颜色对话框", "kind", "mso"),
    "EyedropperFillText", Map("payload", "EyedropperFillText", "label", "填充文本取色器", "kind", "mso"),
    "FontDialogPowerPoint", Map("payload", "FontDialogPowerPoint", "label", "字体对话框", "kind", "mso"),
    "BulletsAndNumberingBulletsDialog", Map("payload", "BulletsAndNumberingBulletsDialog", "label", "着重号对话框", "kind", "mso"),
    "BulletsAndNumberingNumberingDialog", Map("payload", "BulletsAndNumberingNumberingDialog", "label", "编号对话框", "kind", "mso"),
    "IndentDecrease", Map("payload", "IndentDecrease", "label", "缩进减少", "kind", "mso"),
    "IndentIncrease", Map("payload", "IndentIncrease", "label", "缩进增加", "kind", "mso"),
    "AlignLeft", Map("payload", "AlignLeft", "label", "居左对齐", "kind", "mso"),
    "AlignCenter", Map("payload", "AlignCenter", "label", "居中对齐", "kind", "mso"),
    "AlignRight", Map("payload", "AlignRight", "label", "居右对齐", "kind", "mso"),
    "AlignJustify", Map("payload", "AlignJustify", "label", "两端对齐", "kind", "mso"),
    "AlignJustifyWithMixedLanguages", Map("payload", "AlignJustifyWithMixedLanguages", "label", "带多种语言的两端对齐", "kind", "mso"),
    "AlignJustifyLow", Map("payload", "AlignJustifyLow", "label", "低层次两端对齐", "kind", "mso"),
    "ParagraphDistributed", Map("payload", "ParagraphDistributed", "label", "分散对齐", "kind", "mso"),
    "AlignJustifyThai", Map("payload", "AlignJustifyThai", "label", "带泰语的两端对齐", "kind", "mso"),
    "TextDirectionLeftToRight", Map("payload", "TextDirectionLeftToRight", "label", "文字方向-从左到右", "kind", "mso"),
    "TextDirectionRightToLeft", Map("payload", "TextDirectionRightToLeft", "label", "文字方向-从右到左", "kind", "mso"),
    "ParagraphMoreColumnsDialog", Map("payload", "ParagraphMoreColumnsDialog", "label", "段落-更多栏", "kind", "mso"),
    "TextDirectionMoreOptionsDialog", Map("payload", "TextDirectionMoreOptionsDialog", "label", "文本方向更多选项-对话框", "kind", "mso"),
    "TextAlignMoreOptionsDialog", Map("payload", "TextAlignMoreOptionsDialog", "label", "对齐-更多选项", "kind", "mso"),
    "ConvertToSmartArtMoreSmartArtGraphicsDialog", Map("payload", "ConvertToSmartArtMoreSmartArtGraphicsDialog", "label", "转换为智能图形-更多智能图形", "kind", "mso"),
    "PowerPointParagraphDialog", Map("payload", "PowerPointParagraphDialog", "label", "幻灯片段落-对话框", "kind", "mso"),
    "ObjectBringToFront", Map("payload", "ObjectBringToFront", "label", "对象-置于顶层", "kind", "mso"),
    "ObjectSendToBack", Map("payload", "ObjectSendToBack", "label", "对象-置于底层", "kind", "mso"),
    "ObjectBringForward", Map("payload", "ObjectBringForward", "label", "对象-上移一层", "kind", "mso"),
    "ObjectSendBackward", Map("payload", "ObjectSendBackward", "label", "对象-下移一层", "kind", "mso"),
    "ObjectsGroup", Map("payload", "ObjectsGroup", "label", "对象-组合", "kind", "mso"),
    "ObjectsUngroup", Map("payload", "ObjectsUngroup", "label", "对象-取消组合", "kind", "mso"),
    "ObjectsRegroup", Map("payload", "ObjectsRegroup", "label", "对象-重新组合", "kind", "mso"),
    "ObjectsAlignLeftSmart", Map("payload", "ObjectsAlignLeftSmart", "label", "对象-智能对齐-居左", "kind", "mso"),
    "ObjectsAlignCenterHorizontalSmart", Map("payload", "ObjectsAlignCenterHorizontalSmart", "label", "对象-智能对齐-居中水平", "kind", "mso"),
    "ObjectsAlignRightSmart", Map("payload", "ObjectsAlignRightSmart", "label", "对象-智能对齐-居右", "kind", "mso"),
    "ObjectsAlignTopSmart", Map("payload", "ObjectsAlignTopSmart", "label", "对象-智能对齐-居上", "kind", "mso"),
    "ObjectsAlignMiddleVerticalSmart", Map("payload", "ObjectsAlignMiddleVerticalSmart", "label", "对象-智能对齐-居中垂直", "kind", "mso"),
    "ObjectsAlignBottomSmart", Map("payload", "ObjectsAlignBottomSmart", "label", "对象-智能对齐-居下", "kind", "mso"),
    "AlignDistributeHorizontally", Map("payload", "AlignDistributeHorizontally", "label", "对齐-水平分布", "kind", "mso"),
    "AlignDistributeVertically", Map("payload", "AlignDistributeVertically", "label", "对齐-垂直分布", "kind", "mso"),
    "ObjectsAlignRelativeToContainerSmart", Map("payload", "ObjectsAlignRelativeToContainerSmart", "label", "对象-相对于容器对齐", "kind", "mso"),
    "ObjectsAlignSelectedSmart", Map("payload", "ObjectsAlignSelectedSmart", "label", "对象-对齐所选项", "kind", "mso"),
    "ObjectRotationOptionsDialog", Map("payload", "ObjectRotationOptionsDialog", "label", "对象-旋转选项-对话框", "kind", "mso"),
    "ObjectFillMoreColorsDialog", Map("payload", "ObjectFillMoreColorsDialog", "label", "对象-填充-更多颜色", "kind", "mso"),
    "EyedropperFill", Map("payload", "EyedropperFill", "label", "取色器-填充", "kind", "mso"),
    "ObjectPictureFill", Map("payload", "ObjectPictureFill", "label", "对象-图片填充", "kind", "mso"),
    "ShapeFillMoreGradientsDialog", Map("payload", "ShapeFillMoreGradientsDialog", "label", "形状-填充-更多渐变", "kind", "mso"),
    "MoreTextureOptions", Map("payload", "MoreTextureOptions", "label", "更多纹理选项", "kind", "mso"),
    "ObjectBorderOutlineColorMoreColorsDialog", Map("payload", "ObjectBorderOutlineColorMoreColorsDialog", "label", "对象-边框/外框颜色-更多颜色", "kind", "mso"),
    "EyedropperOutline", Map("payload", "EyedropperOutline", "label", "取色器-外框", "kind", "mso"),
    "LineStylesDialog", Map("payload", "LineStylesDialog", "label", "线条样式-对话框", "kind", "mso"),
    "ArrowsMore", Map("payload", "ArrowsMore", "label", "更多箭头", "kind", "mso"),
    "_3DBevelOptionsDialog", Map("payload", "_3DBevelOptionsDialog", "label", "3D斜角选项-对话框", "kind", "mso"),
    "ShadowOptionsDialog", Map("payload", "ShadowOptionsDialog", "label", "阴影选项对话框", "kind", "mso"),
    "ReflectionsMoreOptions", Map("payload", "ReflectionsMoreOptions", "label", "反射选项对话框", "kind", "mso"),
    "GlowColorMoreColorsDialog", Map("payload", "GlowColorMoreColorsDialog", "label", "光晕颜色对话框", "kind", "mso"),
    "EyedropperGlow", Map("payload", "EyedropperGlow", "label", "取色器-光晕", "kind", "mso"),
    "GlowsMoreOptions", Map("payload", "GlowsMoreOptions", "label", "光晕-更多选项", "kind", "mso"),
    "SoftEdgesMoreOptions", Map("payload", "SoftEdgesMoreOptions", "label", "柔化边缘-更多选项", "kind", "mso"),
    "_3DRotationOptionsDialog", Map("payload", "_3DRotationOptionsDialog", "label", "3D旋转选项-对话框", "kind", "mso"),
    "ObjectFormatDialog", Map("payload", "ObjectFormatDialog", "label", "对象格式对话框", "kind", "mso"),
    "FindDialog", Map("payload", "FindDialog", "label", "查找对话框", "kind", "mso"),
    "ReplaceDialog", Map("payload", "ReplaceDialog", "label", "替换对话框", "kind", "mso"),
    "FontsReplaceFonts", Map("payload", "FontsReplaceFonts", "label", "替换字对话框", "kind", "mso"),
    "Dictate", Map("payload", "Dictate", "label", "听写", "kind", "mso"),
    "TableInsert", Map("payload", "TableInsert", "label", "插入表格", "kind", "mso"),
    "TableDrawTable", Map("payload", "TableDrawTable", "label", "绘制表格", "kind", "mso"),
    "ExcelSpreadsheetInsert", Map("payload", "ExcelSpreadsheetInsert", "label", "插入Excel电子表格", "kind", "mso"),
    "PictureInsertFromFilePowerPoint", Map("payload", "PictureInsertFromFilePowerPoint", "label", "从文件插入图片", "kind", "mso"),
    "ClipArtInsertDialog", Map("payload", "ClipArtInsertDialog", "label", "插入剪贴画对话框", "kind", "mso"),
    "ScreenClipping", Map("payload", "ScreenClipping", "label", "屏幕剪辑", "kind", "mso"),
    "PhotoAlbumInsert", Map("payload", "PhotoAlbumInsert", "label", "插入相册", "kind", "mso"),
    "PhotoAlbumEdit", Map("payload", "PhotoAlbumEdit", "label", "编辑相册", "kind", "mso"),
    "IconInsertFromFile", Map("payload", "IconInsertFromFile", "label", "从文件插入图标", "kind", "mso"),
    "Insert3DModelFallback", Map("payload", "Insert3DModelFallback", "label", "插入3D模型备用", "kind", "mso"),
    "Insert3DModelDefault", Map("payload", "Insert3DModelDefault", "label", "插入默认3D模型", "kind", "mso"),
    "Insert3DModelFromFile", Map("payload", "Insert3DModelFromFile", "label", "从文件插入3D模型", "kind", "mso"),
    "Insert3DModelFromOnline", Map("payload", "Insert3DModelFromOnline", "label", "从在线插入3D模型", "kind", "mso"),
    "SmartArtInsert", Map("payload", "SmartArtInsert", "label", "插入智能图形", "kind", "mso"),
    "ChartInsert", Map("payload", "ChartInsert", "label", "插入图表", "kind", "mso"),
    "OfficeExtensionsAppStore", Map("payload", "OfficeExtensionsAppStore", "label", "应用商店", "kind", "mso"),
    "OfficeExtensionsDialog", Map("payload", "OfficeExtensionsDialog", "label", "扩展对话框", "kind", "mso"),
    "OfficeExtensionsManageOtherAddins", Map("payload", "OfficeExtensionsManageOtherAddins", "label", "管理加载项", "kind", "mso"),
    "MSPPTInsertTableofContents", Map("payload", "MSPPTInsertTableofContents", "label", "插入摘要缩放定位", "kind", "mso"),
    "SectionZoomInsert", Map("payload", "SectionZoomInsert", "label", "插入节缩放定位", "kind", "mso"),
    "SlideZoomInsert", Map("payload", "SlideZoomInsert", "label", "插入幻灯片缩放定位", "kind", "mso"),
    "ActionInsert", Map("payload", "ActionInsert", "label", "插入动作", "kind", "mso"),
    "InsertNewComment", Map("payload", "InsertNewComment", "label", "插入新批注", "kind", "mso"),
    "TextBoxInsert", Map("payload", "TextBoxInsert", "label", "插入文本框", "kind", "mso"),
    "TextBoxInsertHorizontal", Map("payload", "TextBoxInsertHorizontal", "label", "插入横向文本框", "kind", "mso"),
    "TextBoxInsertVertical", Map("payload", "TextBoxInsertVertical", "label", "插入纵向文本框", "kind", "mso"),
    "HeaderFooterInsert", Map("payload", "HeaderFooterInsert", "label", "插入页眉页脚", "kind", "mso"),
    "DateAndTimeInsert", Map("payload", "DateAndTimeInsert", "label", "插入日期和时间", "kind", "mso"),
    "NumberInsert", Map("payload", "NumberInsert", "label", "插入数字", "kind", "mso"),
    "OleObjectctInsert", Map("payload", "OleObjectctInsert", "label", "插入OLE对象", "kind", "mso"),
    "EquationInsertNew", Map("payload", "EquationInsertNew", "label", "插入数学公式", "kind", "mso"),
    "InkEquation", Map("payload", "InkEquation", "label", "手写公式", "kind", "mso"),
    "SymbolInsert", Map("payload", "SymbolInsert", "label", "插入符号", "kind", "mso"),
    "MovieFromClipOrganizerInsert", Map("payload", "MovieFromClipOrganizerInsert", "label", "从剪贴画库插入影片", "kind", "mso"),
    "MovieFromFileInsert", Map("payload", "MovieFromFileInsert", "label", "从文件插入影片", "kind", "mso"),
    "SoundInsertFromFile", Map("payload", "SoundInsertFromFile", "label", "从文件插入声音", "kind", "mso"),
    "SoundRecord", Map("payload", "SoundRecord", "label", "录制声音", "kind", "mso"),
    "ObjectScreenRecording", Map("payload", "ObjectScreenRecording", "label", "屏幕录制", "kind", "mso"),
    "InsertMSForms", Map("payload", "InsertMSForms", "label", "插入MS表单", "kind", "mso"),
    "FingerPaintingMode", Map("payload", "FingerPaintingMode", "label", "通过触摸绘制", "kind", "mso"),
    "InkEraser", Map("payload", "InkEraser", "label", "笔划橡皮擦", "kind", "mso"),
    "PointEraser", Map("payload", "PointEraser", "label", "点状橡皮擦", "kind", "mso"),
    "SegmentEraser", Map("payload", "SegmentEraser", "label", "线段橡皮擦", "kind", "mso"),
    "ShowRulerStencil", Map("payload", "ShowRulerStencil", "label", "显示标尺", "kind", "mso"),
    "InkToTextAnalysis", Map("payload", "InkToTextAnalysis", "label", "将墨迹转换为文本", "kind", "mso"),
    "InkToShapeAnalysis", Map("payload", "InkToShapeAnalysis", "label", "将墨迹转换为形状", "kind", "mso"),
    "InkToMathAnalysis", Map("payload", "InkToMathAnalysis", "label", "将墨迹转换为数学公式", "kind", "mso"),
    "MathInputEditor", Map("payload", "MathInputEditor", "label", "墨迹公式编辑器", "kind", "mso"),
    "Replay", Map("payload", "Replay", "label", "墨迹重播", "kind", "mso"),
    "ThemeSearchOfficeOnlinePowerPoint", Map("payload", "ThemeSearchOfficeOnlinePowerPoint", "label", "启动在线主题搜索", "kind", "mso"),
    "ThemeBrowseForThemesPowerPoint", Map("payload", "ThemeBrowseForThemesPowerPoint", "label", "浏览并选择主题", "kind", "mso"),
    "ThemeSaveCurrentPowerPoint", Map("payload", "ThemeSaveCurrentPowerPoint", "label", "保存当前主题", "kind", "mso"),
    "ThemeColorsCreateNew", Map("payload", "ThemeColorsCreateNew", "label", "创建新的主题颜色", "kind", "mso"),
    "ThemeColorsReset", Map("payload", "ThemeColorsReset", "label", "重置主题颜色", "kind", "mso"),
    "ThemeFontsCreateNew", Map("payload", "ThemeFontsCreateNew", "label", "创建新的主题字体", "kind", "mso"),
    "SlideBackgroundFormatDialog", Map("payload", "SlideBackgroundFormatDialog", "label", "幻灯片背景格式对话框", "kind", "mso"),
    "SlideBackgroundReset", Map("payload", "SlideBackgroundReset", "label", "重置幻灯片背景", "kind", "mso"),
    "CustomSlideSize", Map("payload", "CustomSlideSize", "label", "自定义幻灯片大小", "kind", "mso"),
    "LaunchFormatBackground", Map("payload", "LaunchFormatBackground", "label", "启动格式背景对话框", "kind", "mso"),
    "DesignerPane", Map("payload", "DesignerPane", "label", "设计窗格", "kind", "mso"),
    "TransitionPreview", Map("payload", "TransitionPreview", "label", "过渡效果预览", "kind", "mso"),
    "TransitionSoundLoopUntilNextSound", Map("payload", "TransitionSoundLoopUntilNextSound", "label", "过渡声音循环直到下一个声音", "kind", "mso"),
    "TransitionDuration", Map("payload", "TransitionDuration", "label", "过渡持续时间", "kind", "mso"),
    "SlideTransitionApplyToAll", Map("payload", "SlideTransitionApplyToAll", "label", "将过渡应用于所有幻灯片", "kind", "mso"),
    "SlideTransitionOnMouseClick", Map("payload", "SlideTransitionOnMouseClick", "label", "单击鼠标时换片", "kind", "mso"),
    "SlideTransitionAutomaticallyAfter", Map("payload", "SlideTransitionAutomaticallyAfter", "label", "自动换片", "kind", "mso"),
    "AnimationPreview", Map("payload", "AnimationPreview", "label", "动画预览", "kind", "mso"),
    "AnimationAutoPreview", Map("payload", "AnimationAutoPreview", "label", "自动动画预览", "kind", "mso"),
    "AnimationCustomEntranceDialog", Map("payload", "AnimationCustomEntranceDialog", "label", "自定义入场动画对话框", "kind", "mso"),
    "AnimationCustomEmphasisDialog", Map("payload", "AnimationCustomEmphasisDialog", "label", "自定义强调动画对话框", "kind", "mso"),
    "AnimationCustomExitDialog", Map("payload", "AnimationCustomExitDialog", "label", "自定义退出动画对话框", "kind", "mso"),
    "AnimationCustomPathDialog", Map("payload", "AnimationCustomPathDialog", "label", "自定义路径动画对话框", "kind", "mso"),
    "AnimationCustomActionVerbDialog", Map("payload", "AnimationCustomActionVerbDialog", "label", "自定义动作动画动词对话框", "kind", "mso"),
    "EffectOptionsDialog", Map("payload", "EffectOptionsDialog", "label", "效果选项对话框", "kind", "mso"),
    "AnimationCustomAddEntranceDialog", Map("payload", "AnimationCustomAddEntranceDialog", "label", "添加自定义入场动画", "kind", "mso"),
    "AnimationCustomAddEmphasisDialog", Map("payload", "AnimationCustomAddEmphasisDialog", "label", "添加自定义强调动画", "kind", "mso"),
    "AnimationCustomAddExitDialog", Map("payload", "AnimationCustomAddExitDialog", "label", "添加自定义退出动画", "kind", "mso"),
    "AnimationCustomAddPathDialog", Map("payload", "AnimationCustomAddPathDialog", "label", "添加自定义路径动画", "kind", "mso"),
    "AnimationCustomAddActionVerbDlog", Map("payload", "AnimationCustomAddActionVerbDlog", "label", "添加自定义动作动画动词", "kind", "mso"),
    "AnimationCustom", Map("payload", "AnimationCustom", "label", "自定义动画窗格", "kind", "mso"),
    "AnimationPainter", Map("payload", "AnimationPainter", "label", "动画绘图工具", "kind", "mso"),
    "AnimationDuration", Map("payload", "AnimationDuration", "label", "动画持续时间", "kind", "mso"),
    "AnimationDelay", Map("payload", "AnimationDelay", "label", "动画延迟", "kind", "mso"),
    "AnimationMoveEarlier", Map("payload", "AnimationMoveEarlier", "label", "提前播放动画", "kind", "mso"),
    "AnimationMoveLater", Map("payload", "AnimationMoveLater", "label", "延后播放动画", "kind", "mso"),
    "SlideShowFromCurrent", Map("payload", "SlideShowFromCurrent", "label", "从当前幻灯片开始播放", "kind", "mso"),
    "BroadcastSlideShow", Map("payload", "BroadcastSlideShow", "label", "广播幻灯片放映", "kind", "mso"),
    "BroadcastSlideShowLync", Map("payload", "BroadcastSlideShowLync", "label", "使用Lync广播幻灯片放映", "kind", "mso"),
    "BroadcastSlideShowGeneric", Map("payload", "BroadcastSlideShowGeneric", "label", "使用通用广播幻灯片放映", "kind", "mso"),
    "SlideShowSetUpDialog", Map("payload", "SlideShowSetUpDialog", "label", "设置幻灯片放映", "kind", "mso"),
    "SlideHide", Map("payload", "SlideHide", "label", "隐藏幻灯片", "kind", "mso"),
    "SlideShowRehearseTimings", Map("payload", "SlideShowRehearseTimings", "label", "排练定时", "kind", "mso"),
    "RecordNarrationFromCurrentSlide", Map("payload", "RecordNarrationFromCurrentSlide", "label", "从当前幻灯片开始录制旁白", "kind", "mso"),
    "RecordNarration", Map("payload", "RecordNarration", "label", "从头开始录制旁白", "kind", "mso"),
    "SlideShowPlayNarrations", Map("payload", "SlideShowPlayNarrations", "label", "播放幻灯片放映的旁白", "kind", "mso"),
    "SlideShowUseRehearsedTimings", Map("payload", "SlideShowUseRehearsedTimings", "label", "使用排练的定时进行幻灯片放映", "kind", "mso"),
    "MediaControlsShow", Map("payload", "MediaControlsShow", "label", "显示媒体控件", "kind", "mso"),
    "SlideShowUsePresenterView", Map("payload", "SlideShowUsePresenterView", "label", "使用演示者视图", "kind", "mso"),
    "ViewThumbnailViewPowerPoint", Map("payload", "ViewThumbnailViewPowerPoint", "label", "查看缩略图视图", "kind", "mso"),
    "ViewOutlineViewPowerPoint", Map("payload", "ViewOutlineViewPowerPoint", "label", "查看大纲视图", "kind", "mso"),
    "ViewSlideSorterView", Map("payload", "ViewSlideSorterView", "label", "查看幻灯片排序视图", "kind", "mso"),
    "ViewNotesPageView", Map("payload", "ViewNotesPageView", "label", "查看备注页视图", "kind", "mso"),
    "ViewSlideShowReadingView", Map("payload", "ViewSlideShowReadingView", "label", "查看幻灯片阅读视图", "kind", "mso"),
    "ViewSlideMasterView", Map("payload", "ViewSlideMasterView", "label", "查看幻灯片母版视图", "kind", "mso"),
    "ViewHandoutMasterView", Map("payload", "ViewHandoutMasterView", "label", "查看讲义母版视图", "kind", "mso"),
    "ViewNotesMasterView", Map("payload", "ViewNotesMasterView", "label", "查看备注母版视图", "kind", "mso"),
    "ViewRulerPowerPoint", Map("payload", "ViewRulerPowerPoint", "label", "查看标尺", "kind", "mso"),
    "ViewGridlinesPowerPoint", Map("payload", "ViewGridlinesPowerPoint", "label", "查看网格线", "kind", "mso"),
    "GuidesShowHide", Map("payload", "GuidesShowHide", "label", "显示/隐藏参考线", "kind", "mso"),
    "ShowNotes", Map("payload", "ShowNotes", "label", "显示备注", "kind", "mso"),
    "GridSettings", Map("payload", "GridSettings", "label", "网格设置", "kind", "mso"),
    "ViewDirectionLeftToRight", Map("payload", "ViewDirectionLeftToRight", "label", "查看方向从左到右", "kind", "mso"),
    "ViewDirectionRightToLeft", Map("payload", "ViewDirectionRightToLeft", "label", "查看方向从右到左", "kind", "mso"),
    "ZoomDialog", Map("payload", "ZoomDialog", "label", "缩放对话框", "kind", "mso"),
    "ZoomFitToWindow", Map("payload", "ZoomFitToWindow", "label", "缩放以适应窗口", "kind", "mso"),
    "ViewDisplayInHighContrast", Map("payload", "ViewDisplayInHighContrast", "label", "显示高对比度模式", "kind", "mso"),
    "ViewDisplayInColor", Map("payload", "ViewDisplayInColor", "label", "显示彩色模式", "kind", "mso"),
    "ViewDisplayInGrayscale", Map("payload", "ViewDisplayInGrayscale", "label", "显示灰度模式", "kind", "mso"),
    "ViewDisplayInPureBlackAndWhite", Map("payload", "ViewDisplayInPureBlackAndWhite", "label", "显示纯黑白模式", "kind", "mso"),
    "WindowNew", Map("payload", "WindowNew", "label", "新建窗口", "kind", "mso"),
    "WindowsArrangeAll", Map("payload", "WindowsArrangeAll", "label", "排列所有窗口", "kind", "mso"),
    "WindowsCascade", Map("payload", "WindowsCascade", "label", "层叠窗口", "kind", "mso"),
    "WindowMoveSplit", Map("payload", "WindowMoveSplit", "label", "移动拆分窗口", "kind", "mso"),
    "WindowMoreWindowsDialog", Map("payload", "WindowMoreWindowsDialog", "label", "更多窗口对话框", "kind", "mso"),
    "MacroPlay", Map("payload", "MacroPlay", "label", "播放宏", "kind", "mso"),
    "VisualBasic", Map("payload", "VisualBasic", "label", "VisualBasic", "kind", "mso"),
    "MacroSecurity", Map("payload", "MacroSecurity", "label", "宏安全", "kind", "mso"),
    "AddInManager", Map("payload", "AddInManager", "label", "加载项管理器", "kind", "mso"),
    "ComAddInsDialog", Map("payload", "ComAddInsDialog", "label", "COM加载项对话框", "kind", "mso"),
    "ActiveXLabel", Map("payload", "ActiveXLabel", "label", "ActiveX标签", "kind", "mso"),
    "ActiveXTextBox", Map("payload", "ActiveXTextBox", "label", "ActiveX文本框", "kind", "mso"),
    "ActiveXSpinButton", Map("payload", "ActiveXSpinButton", "label", "ActiveX旋转按钮", "kind", "mso"),
    "ActiveXButton", Map("payload", "ActiveXButton", "label", "ActiveX按钮", "kind", "mso"),
    "ActiveXImage", Map("payload", "ActiveXImage", "label", "ActiveX图像", "kind", "mso"),
    "ActiveXScrollBar", Map("payload", "ActiveXScrollBar", "label", "ActiveX滚动条", "kind", "mso"),
    "ActiveXCheckBox", Map("payload", "ActiveXCheckBox", "label", "ActiveX复选框", "kind", "mso"),
    "ActiveXRadioButton", Map("payload", "ActiveXRadioButton", "label", "ActiveX单选按钮", "kind", "mso"),
    "ActiveXComboBox", Map("payload", "ActiveXComboBox", "label", "ActiveX组合框", "kind", "mso"),
    "ActiveXListBox", Map("payload", "ActiveXListBox", "label", "ActiveX列表框", "kind", "mso"),
    "ActiveXToggleButton", Map("payload", "ActiveXToggleButton", "label", "ActiveX切换按钮", "kind", "mso"),
    "MoreControlsDialog", Map("payload", "MoreControlsDialog", "label", "更多控件对话框", "kind", "mso"),
    "ControlProperties", Map("payload", "ControlProperties", "label", "控件属性", "kind", "mso"),
    "ViewVisualBasicCode", Map("payload", "ViewVisualBasicCode", "label", "查看VisualBasic代码", "kind", "mso"),
    "SlideNew", Map("payload", "SlideNew", "label", "新建幻灯片", "kind", "mso"),
    "SlideMasterInsertLayout", Map("payload", "SlideMasterInsertLayout", "label", "在母版中插入版式", "kind", "mso"),
    "SlideDelete", Map("payload", "SlideDelete", "label", "删除幻灯片", "kind", "mso"),
    "SlideMasterRenameMaster", Map("payload", "SlideMasterRenameMaster", "label", "重命名母版", "kind", "mso"),
    "SlideMasterPreserveMaster", Map("payload", "SlideMasterPreserveMaster", "label", "保留母版", "kind", "mso"),
    "SlideMasterMasterLayout", Map("payload", "SlideMasterMasterLayout", "label", "母版版式", "kind", "mso"),
    "SlideMasterContentPlaceholder", Map("payload", "SlideMasterContentPlaceholder", "label", "在母版中插入内容占位符", "kind", "mso"),
    "SlideMasterVerticalContentPla", Map("payload", "SlideMasterVerticalContentPla", "label", "在母版中插入垂直内容占位符", "kind", "mso"),
    "SlideMasterTextPlaceholderIns", Map("payload", "SlideMasterTextPlaceholderIns", "label", "在母版中插入文本占位符", "kind", "mso"),
    "SlideMasterVerticalTextPlaceh", Map("payload", "SlideMasterVerticalTextPlaceh", "label", "在母版中插入垂直文本占位符", "kind", "mso"),
    "SlideMasterPicturePlaceholder", Map("payload", "SlideMasterPicturePlaceholder", "label", "在母版中插入图片占位符", "kind", "mso"),
    "SlideMasterChartPlaceholderIn", Map("payload", "SlideMasterChartPlaceholderIn", "label", "在母版中插入图表占位符", "kind", "mso"),
    "SlideMasterTablePlaceholderIn", Map("payload", "SlideMasterTablePlaceholderIn", "label", "在母版中插入表格占位符", "kind", "mso"),
    "SlideMasterDiagramPlaceholder", Map("payload", "SlideMasterDiagramPlaceholder", "label", "在母版中插入图表占位符", "kind", "mso"),
    "SlideMasterMediaPlaceholderIn", Map("payload", "SlideMasterMediaPlaceholderIn", "label", "在母版中插入媒体占位符", "kind", "mso"),
    "SlideMasterClipArtPlaceholder", Map("payload", "SlideMasterClipArtPlaceholder", "label", "在母版中插入剪贴画占位符", "kind", "mso"),
    "SlideMasterShowTitle", Map("payload", "SlideMasterShowTitle", "label", "显示母版版式标题", "kind", "mso"),
    "SlideMasterShowFooters", Map("payload", "SlideMasterShowFooters", "label", "显示母版版式页脚", "kind", "mso"),
    "SlideBackgroundHideGraphics", Map("payload", "SlideBackgroundHideGraphics", "label", "隐藏母版版式背景图形", "kind", "mso"),
    "MasterViewClose", Map("payload", "MasterViewClose", "label", "关闭母版视图", "kind", "mso"),
    "PasteAsHyperlink", Map("payload", "PasteAsHyperlink", "label", "作为链接粘贴", "kind", "mso"),
    "PasteBitmap", Map("payload", "PasteBitmap", "label", "作为Bmp粘贴", "kind", "mso"),
    "PasteGif", Map("payload", "PasteGif", "label", "作为Gif粘贴", "kind", "mso"),
    "PastePng", Map("payload", "PastePng", "label", "作为Png粘贴", "kind", "mso"),
    "PasteJpeg", Map("payload", "PasteJpeg", "label", "作为Jpeg粘贴", "kind", "mso"),
    "PasteTextOnly", Map("payload", "PasteTextOnly", "label", "仅粘贴文本", "kind", "mso"),
    "PasteAsPicture", Map("payload", "PasteAsPicture", "label", "作为图片粘贴", "kind", "mso"),
    "PasteSourceFormatting", Map("payload", "PasteSourceFormatting", "label", "使用源格式粘贴", "kind", "mso"),
    "PasteDestinationTheme", Map("payload", "PasteDestinationTheme", "label", "使用目标主题粘贴", "kind", "mso"),
    "ExportToVideo", Map("payload", "ExportToVideo", "label", "输出为视频", "kind", "mso"),
    "ShapeRectangle", Map("payload", "ShapeRectangle", "label", "创建矩形形状", "kind", "mso"),
    "ShapeOval", Map("payload", "ShapeOval", "label", "创建椭圆形状", "kind", "mso"),
    "ShapeElbowConnectorArrow", Map("payload", "ShapeElbowConnectorArrow", "label", "创建带箭头的肘形连接器形状", "kind", "mso"),
    "ObjectRotateLeft90", Map("payload", "ObjectRotateLeft90", "label", "向左旋转对象90度", "kind", "mso"),
    "ShapeStraightConnectorArrow", Map("payload", "ShapeStraightConnectorArrow", "label", "创建带箭头的直线连接器形状", "kind", "mso"),
    "ShapeElbowConnector", Map("payload", "ShapeElbowConnector", "label", "创建肘形连接器形状", "kind", "mso"),
    "ObjectSetShapeDefaults", Map("payload", "ObjectSetShapeDefaults", "label", "设置对象默认形状", "kind", "mso"),
    "ShapeRoundedRectangle", Map("payload", "ShapeRoundedRectangle", "label", "创建圆角矩形形状", "kind", "mso"),
    "ShapeStraightConnector", Map("payload", "ShapeStraightConnector", "label", "创建直线连接器形状", "kind", "mso"),
    "ObjectFlipHorizontal", Map("payload", "ObjectFlipHorizontal", "label", "水平翻转对象", "kind", "mso"),
    "ShapeRightArrow", Map("payload", "ShapeRightArrow", "label", "创建向右箭头形状", "kind", "mso"),
    "ObjectRotateFree", Map("payload", "ObjectRotateFree", "label", "自由旋转对象", "kind", "mso"),
    "ShapeDownArrow", Map("payload", "ShapeDownArrow", "label", "创建向下箭头形状", "kind", "mso"),
    "ObjectFlipVertical", Map("payload", "ObjectFlipVertical", "label", "垂直翻转对象", "kind", "mso"),
    "ObjectRotateRight90", Map("payload", "ObjectRotateRight90", "label", "向右旋转对象90度", "kind", "mso"),
    "ShapeRoundedRectangularCallout", Map("payload", "ShapeRoundedRectangularCallout", "label", "创建圆角矩形标注形状", "kind", "mso"),
    "ShapeIsoscelesTriangle", Map("payload", "ShapeIsoscelesTriangle", "label", "创建等腰三角形形状", "kind", "mso"),
    "ShapesMoreShapes", Map("payload", "ShapesMoreShapes", "label", "更多形状", "kind", "mso"),
    "ShapeLeftBrace", Map("payload", "ShapeLeftBrace", "label", "创建左括号形状", "kind", "mso"),
    "ObjectNudgeDown", Map("payload", "ObjectNudgeDown", "label", "对象下移", "kind", "mso"),
    "ShapeRightBrace", Map("payload", "ShapeRightBrace", "label", "创建右括号形状", "kind", "mso"),
    "PickUpStyle", Map("payload", "PickUpStyle", "label", "拾取格式", "kind", "mso"),
    "About", Map("payload", "About", "label", "关于", "kind", "mso")
)   
 
func_config := Map( 
    "open_powershell_in_folder", Map("payload", "open_powershell_in_folder", "label", "PowerShell", "kind", "func"),
    "paste_ai_prompt_1", Map("payload", "paste_ai_prompt_1", "label", "音律通用提示词", "kind", "func"),
    "paste_ai_prompt_2", Map("payload", "paste_ai_prompt_2", "label", "音律重建总控提示词", "kind", "func"),
    "paste_ai_prompt_3", Map("payload", "paste_ai_prompt_3", "label", "音律PPT复刻提示词", "kind", "func"),
    "open_duty_schedule", Map("payload", "open_duty_schedule", "label", "打开排班表", "kind", "func"),
    "close_win_by_ctrlw", Map("payload", "close_win_by_ctrlw", "label", "关闭窗口", "kind", "func"),
    "zip", Map("payload", "zip", "label", "压缩文件", "kind", "func"),
    "unzip", Map("payload", "unzip", "label", "解压缩文件", "kind", "func"),
    "min_window_under_mouse", Map("payload", "min_window_under_mouse", "label", "最小化窗口", "kind", "func"),
    "min_win", Map("payload", "min_win", "label", "最小化窗口", "kind", "func"),
    "func_wins", Map("payload", "func_wins", "label", "Win + S", "kind", "func"),
    "func_ctrls", Map("payload", "func_ctrls", "label", "保 存", "kind", "func"),
    "save_as", Map("payload", "save_as", "label", "另存为……", "kind", "func"),
    "copy_format", Map("payload", "copy_format", "label", "unknown", "kind", "func"),
    "func_ctrlshiftw", Map("payload", "func_ctrlshiftw", "label", "Ctrl + Shift + W", "kind", "func"),
    "smart_ctrlw", Map("payload", "smart_ctrlw", "label", "Ctrl + W", "kind", "func"),
    "func_winw", Map("payload", "func_winw", "label", "Win+W", "kind", "func"),
    "shutdown_pc", Map("payload", "shutdown_computer", "label", "关 机", "kind", "func"),
    "restart_pc", Map("payload", "restart_computer", "label", "重 启", "kind", "func"),
    "lock_screen", Map("payload", "lock_screen", "label", "锁 定 屏 幕", "kind", "func"),
    "close_window_under_mouse", Map("payload", "close_window_under_mouse", "label", "关闭窗口", "kind", "func"),
    "max_window", Map("payload", "max_window", "label", "最大化窗口", "kind", "func"),
    "left_window", Map("payload", "left_window", "label", "窗口靠左半边", "kind", "func"),
    "right_window", Map("payload", "right_window", "label", "窗口靠右半边", "kind", "func"),
    "paste_format", Map("payload", "paste_format", "label", "粘贴格式", "kind", "func"),
    "func_wind", Map("payload", "func_wind", "label", "显示桌面", "kind", "func"),
    "close_ppt_process", Map("payload", "close_ppt_process", "label", "关闭PPT进程", "kind", "func"),
    "open_current_path", Map("payload", "open_current_path", "label", "文件夹", "kind", "func"),
    "open_recycle_bin", Map("payload", "open_recycle_bin", "label", "回收站", "kind", "func"),
    "do_nothing", Map("payload", "do_nothing", "label", "Cancel ?", "kind", "func"),
    "reload_me", Map("payload", "reload_me", "label", "Reload QWERT ?", "kind", "func"),
    "suspend_me", Map("payload", "suspend_me", "label", "Suspend QWERT ?", "kind", "func"),
    "ExitApp", Map("payload", "ExitApp", "label", "Shutdown QWERT ?", "kind", "func"),
    "func_cancel", Map("payload", "func_cancel", "label", "Cancel ?", "kind", "func"),
    "func_ctrlx", Map("payload", "func_ctrlx", "label", "剪 切", "kind", "func"),
    "add_slide_from_eagle", Map("payload", "add_slide_from_eagle", "label", "从eagle中复制PPT页面", "kind", "func")
)

key_config := Map( 
    "#^o", Map("payload", "#^o", "label", "屏幕键盘", "kind", "key"),
    "^+!q", Map("payload", "^+!q", "label", "打开/关闭系统代理", "kind", "key"),
    "^+!w", Map("payload", "^+!w", "label", "打开/关闭面板", "kind", "key"),
    "#^o", Map("payload", "#^o", "label", "屏幕键盘", "kind", "key"),
    "{Delete}", Map("payload", "{Delete}", "label", "Delete", "kind", "key"),
    "{Space 3}", Map("payload", "{Space 3}", "label", "沉浸式翻译", "kind", "key"),
    "{Volume_Mute}", Map("payload", "{Volume_Mute}", "label", "Volume Mute", "kind", "key"),
    "+{End}^c", Map("payload", "+{End}^c", "label", "选到行末并复制", "kind", "key"),
    "+{End}^x", Map("payload", "+{End}^c", "label", "选到行末并剪切", "kind", "key"),
    "+{End}^v", Map("payload", "+{End}^v", "label", "选到行末并粘贴", "kind", "key"),
    "^!+1", Map("payload", "^!+1", "label", "唤起千问主窗口", "kind", "key"),
    "^!+2", Map("payload", "^!+2", "label", "唤起千问快捷框", "kind", "key"),
    "^!+3", Map("payload", "^!+3", "label", "千问截屏", "kind", "key"),
    "^!+4", Map("payload", "^!+4", "label", "千问新建对话", "kind", "key"),
    "!#{Space}", Map("payload", "!#{Space}", "label", "Windows命令面板", "kind", "key"),
    "!#{Space}", Map("payload", "!#{Space}", "label", "Windows命令面板", "kind", "key"),
    "!#{Space}", Map("payload", "!#{Space}", "label", "Windows命令面板", "kind", "key"),
    "!#{Space}", Map("payload", "!#{Space}", "label", "Windows命令面板", "kind", "key"),
    "{Up}", Map("payload", "{Up}", "label", "Up", "kind", "key"),
    "{Down}", Map("payload", "{Down}", "label", "Down", "kind", "key"),
    "{Left}", Map("payload", "{Left}", "label", "Left", "kind", "key"),
    "{Right}", Map("payload", "{Right}", "label", "Right", "kind", "key"),
    "^k^]", Map("payload", "^k^]", "label", "展开所选分区", "kind", "key"),
    "^{Enter}", Map("payload", "^{Enter}", "label", "Ctrl + Enter", "kind", "key"),
    "^k^0", Map("payload", "^k^0", "label", "折叠所有分区", "kind", "key"),
    "{Escape}", Map("payload", "{Escape}", "label", "Esc", "kind", "key"),
    "^{Home}", Map("payload", "^{Home}", "label", "第一页", "kind", "key"),
    "^{End}", Map("payload", "^{End}", "label", "最后一页", "kind", "key"),
    "{PgUp}", Map("payload" , "{PgUp}", "label", "上一页", "kind", "key"),
    "{PgDn}", Map("payload", "{PgDn}", "label", "下一页", "kind", "key"),
    "^{WheelUp}", Map("payload", "^{WheelUp}", "label", "放大字号", "kind", "key"),
    "^{WheelDown}", Map("payload", "^{WheelDown}", "label", "缩小字号", "kind", "key"),
    "+{MButton}", Map("payload", "+{MButton}", "label", "中键 + Shift", "kind", "key"),
    "+{LButton}", Map("payload", "+{LButton}", "label", "左键 + Shift", "kind", "key"),
    "!+{Tab}", Map("payload", "!+{Tab}", "label", "切换应用程序", "kind", "key"),
    "+#{Left}", Map("payload", "+#{Left}", "label", "窗口移动到另一个显示器", "kind", "key"),
    "^a", Map("payload", "^a", "label", "Ctrl + A", "kind", "key"),
    "^+a", Map("payload", "^+a", "label", "Ctrl + Shift + A", "kind", "key"),
    "#a", Map("payload", "#a", "label", "Win + A", "kind", "key"),
    "!a", Map("payload", "!a", "label", "Alt + A", "kind", "key"),
    "^b", Map("payload", "^b", "label", "Ctrl + B", "kind", "key"),
    "^+b", Map("payload", "^+b", "label", "Ctrl + Shift + B", "kind", "key"),
    "#b", Map("payload", "#b", "label", "Win + B", "kind", "key"),
    "!b", Map("payload", "!b", "label", "Alt + B", "kind", "key"),
    "^c", Map("payload", "^c", "label", "Ctrl + C", "kind", "key"),
    "^+c", Map("payload", "^+c", "label", "Ctrl + Shift + C", "kind", "key"),
    "#c", Map("payload", "#c", "label", "Win + C", "kind", "key"),
    "!c", Map("payload", "!c", "label", "Alt + C", "kind", "key"),
    "^+!c", Map("payload", "^+!c", "label", "ChatGPT", "kind", "key"),
    "^d", Map("payload", "^d", "label", "Ctrl + D", "kind", "key"),
    "^+d", Map("payload", "^+d", "label", "Ctrl + Shift + D", "kind", "key"),
    "#d", Map("payload", "#d", "label", "Win + D", "kind", "key"),
    "!d", Map("payload", "!d", "label", "Alt + D", "kind", "key"),
    "^e", Map("payload", "^e", "label", "Ctrl + E", "kind", "key"),
    "^+e", Map("payload", "^+e", "label", "Ctrl + Shift + E", "kind", "key"),
    "#e", Map("payload", "#e", "label", "Win + E", "kind", "key"),
    "!e", Map("payload", "!e", "label", "Alt + E", "kind", "key"),
    "^f", Map("payload", "^f", "label", "Ctrl + F", "kind", "key"),
    "^+f", Map("payload", "^+f", "label", "Ctrl + Shift + F", "kind", "key"),
    "#f", Map("payload", "#f", "label", "Win + F", "kind", "key"),
    "!f", Map("payload", "!f", "label", "Alt + F", "kind", "key"),
    "^g", Map("payload", "^g", "label", "Ctrl + G", "kind", "key"),
    "^+g", Map("payload", "^+g", "label", "Ctrl + Shift + G", "kind", "key"),
    "#g", Map("payload", "#g", "label", "Win + G", "kind", "key"),
    "!g", Map("payload", "!g", "label", "Alt + G", "kind", "key"),
    "^h", Map("payload", "^h", "label", "Ctrl + H", "kind", "key"),
    "^+h", Map("payload", "^+h", "label", "Ctrl + Shift + H", "kind", "key"),
    "#h", Map("payload", "#h", "label", "Win + H", "kind", "key"),
    "!h", Map("payload", "!h", "label", "Alt + H", "kind", "key"),
    "^i", Map("payload", "^i", "label", "Ctrl + I", "kind", "key"),
    "^+i", Map("payload", "^+i", "label", "Ctrl + Shift + I", "kind", "key"),
    "#i", Map("payload", "#i", "label", "Win + I", "kind", "key"),
    "!i", Map("payload", "!i", "label", "Alt + I", "kind", "key"),
    "^j", Map("payload", "^j", "label", "Ctrl + J", "kind", "key"),
    "^+j", Map("payload", "^+j", "label", "Ctrl + Shift + J", "kind", "key"),
    "#j", Map("payload", "#j", "label", "Win + J", "kind", "key"),
    "!j", Map("payload", "!j", "label", "Alt + J", "kind", "key"),
    "^k", Map("payload", "^k", "label", "Ctrl + K", "kind", "key"),
    "^+k", Map("payload", "^+k", "label", "Ctrl + Shift + K", "kind", "key"),
    "#k", Map("payload", "#k", "label", "Win + K", "kind", "key"),
    "!k", Map("payload", "!k", "label", "Alt + K", "kind", "key"),
    "^l", Map("payload", "^l", "label", "Ctrl + L", "kind", "key"),
    "^+l", Map("payload", "^+l", "label", "Ctrl + Shift + L", "kind", "key"),
    "#l", Map("payload", "#l", "label", "Win + L", "kind", "key"),
    "!l", Map("payload", "!l", "label", "Alt + L", "kind", "key"),
    "^m", Map("payload", "^m", "label", "Ctrl + M", "kind", "key"),
    "^+m", Map("payload", "^+m", "label", "Ctrl + Shift + M", "kind", "key"),
    "#m", Map("payload", "#m", "label", "Win + M", "kind", "key"),
    "!m", Map("payload", "!m", "label", "Alt + M", "kind", "key"),
    "^n", Map("payload", "^n", "label", "Ctrl + N", "kind", "key"),
    "^+n", Map("payload", "^+n", "label", "Ctrl + Shift + N", "kind", "key"),
    "#n", Map("payload", "#n", "label", "Win + N", "kind", "key"),
    "!n", Map("payload", "!n", "label", "Alt + N", "kind", "key"),
    "^o", Map("payload", "^o", "label", "Ctrl + O", "kind", "key"),
    "^+o", Map("payload", "^+o", "label", "Ctrl + Shift + O", "kind", "key"),
    "#o", Map("payload", "#o", "label", "Win + O", "kind", "key"),
    "!o", Map("payload", "!o", "label", "Alt + O", "kind", "key"),
    "^p", Map("payload", "^p", "label", "Ctrl + P", "kind", "key"),
    "^+p", Map("payload", "^+p", "label", "Ctrl + Shift + P", "kind", "key"),
    "#p", Map("payload", "#p", "label", "Win + P", "kind", "key"),
    "!p", Map("payload", "!p", "label", "Alt + P", "kind", "key"),
    "^q", Map("payload", "^q", "label", "Ctrl + Q", "kind", "key"),
    "^+q", Map("payload", "^+q", "label", "Ctrl + Shift + Q", "kind", "key"),
    "#q", Map("payload", "#q", "label", "Win + Q", "kind", "key"),
    "!q", Map("payload", "!q", "label", "Alt + Q", "kind", "key"),
    "^r", Map("payload", "^r", "label", "Ctrl + R", "kind", "key"),
    "^+r", Map("payload", "^+r", "label", "Ctrl + Shift + R", "kind", "key"),
    "#r", Map("payload", "#r", "label", "Win + R", "kind", "key"),
    "!r", Map("payload", "!r", "label", "Alt + R", "kind", "key"),
    "^s", Map("payload", "^s", "label", "Ctrl + S", "kind", "key"),
    "^+s", Map("payload", "^+s", "label", "Ctrl + Shift + S", "kind", "key"),
    "#s", Map("payload", "#s", "label", "Win + S", "kind", "key"),
    "!s", Map("payload", "!s", "label", "Alt + S", "kind", "key"),
    "^t", Map("payload", "^t", "label", "Ctrl + T", "kind", "key"),
    "^+t", Map("payload", "^+t", "label", "Ctrl + Shift + T", "kind", "key"),
    "#t", Map("payload", "#t", "label", "Win + T", "kind", "key"),
    "!t", Map("payload", "!t", "label", "Alt + T", "kind", "key"),
    "^u", Map("payload", "^u", "label", "Ctrl + U", "kind", "key"),
    "^+u", Map("payload", "^+u", "label", "Ctrl + Shift + U", "kind", "key"),
    "#u", Map("payload", "#u", "label", "Win + U", "kind", "key"),
    "!u", Map("payload", "!u", "label", "Alt + U", "kind", "key"),
    "^v", Map("payload", "^v", "label", "Ctrl + V", "kind", "key"),
    "^+v", Map("payload", "^+v", "label", "Ctrl + Shift + V", "kind", "key"),
    "#v", Map("payload", "#v", "label", "Win + V", "kind", "key"),
    "#^v", Map("payload", "#^v", "label", "Win + Ctrl + V", "kind", "key"),
    "!v", Map("payload", "!v", "label", "Alt + V", "kind", "key"),
    "^w", Map("payload", "^w", "label", "Ctrl + W", "kind", "key"),
    "^+w", Map("payload", "^+w", "label", "Ctrl + Shift + W", "kind", "key"),
    "#w", Map("payload", "#w", "label", "Win + W", "kind", "key"),
    "!w", Map("payload", "!w", "label", "Alt + W", "kind", "key"),
    "^x", Map("payload", "^x", "label", "unknown", "kind", "key"),
    "^+x", Map("payload", "^+x", "label", "Ctrl + Shift + X", "kind", "key"),
    "#x", Map("payload", "#x", "label", "Win + X", "kind", "key"),
    "!x", Map("payload", "!x", "label", "Alt + X", "kind", "key"),
    "^y", Map("payload", "^y", "label", "Ctrl + Y", "kind", "key"),
    "^+y", Map("payload", "^+y", "label", "Ctrl + Shift + Y", "kind", "key"),
    "#y", Map("payload", "#y", "label", "Win + Y", "kind", "key"),
    "!y", Map("payload", "!y", "label", "Alt + Y", "kind", "key"),
    "^z", Map("payload", "^z", "label", "Ctrl + Z", "kind", "key"),
    "^+z", Map("payload", "^+z", "label", "Ctrl + Shift + Z", "kind", "key"),
    "#z", Map("payload", "#z", "label", "Win + Z", "kind", "key"),
    "!z", Map("payload", "!z", "label", "Alt + Z", "kind", "key"),
    "^;", Map("payload", "^;", "label", "Ctrl + `;", "kind", "key"),
    "^+;", Map("payload", "^+;", "label", "Ctrl + Shift + `;", "kind", "key"),
    "#;", Map("payload", "#;", "label", "Win + `;", "kind", "key"),
    "!;", Map("payload", "!;", "label", "Alt + `;", "kind", "key"),

    "^k^o", Map("payload", "^k^o", "label", "Ctrl + K Ctrl + O", "kind", "key"),
    "^/", Map("payload", "^/", "label", "Ctrl + /", "kind", "key"),
    "^``", Map("payload", "^``", "label", "Ctrl + ``", "kind", "key"),

    "{F1}", Map("payload", "{F1}", "label", "F1", "kind", "key"),
    "{F2}", Map("payload", "{F2}", "label", "F2", "kind", "key"),
    "{F3}", Map("payload", "{F3}", "label", "F3", "kind", "key"),
    "{F4}", Map("payload", "{F4}", "label", "F4", "kind", "key"),
    "{F5}", Map("payload", "{F5}", "label", "F5", "kind", "key"),
    "{F6}", Map("payload", "{F6}", "label", "F6", "kind", "key"),
    "{F7}", Map("payload", "{F7}", "label", "F7", "kind", "key"),
    "{F8}", Map("payload", "{F8}", "label", "F8", "kind", "key"),
    "{F9}", Map("payload", "{F9}", "label", "F9", "kind", "key"),
    "{F10}", Map("payload", "{F10}", "label", "F10", "kind", "key"),
    "{F11}", Map("payload", "{F11}", "label", "F11", "kind", "key"),
    "{F12}", Map("payload", "{F12}", "label", "F12", "kind", "key"),
    "^{F1}", Map("payload", "^{F1}", "label", "Ctrl + F1", "kind", "key"),
    "^{F2}", Map("payload", "^{F2}", "label", "Ctrl + F2", "kind", "key"),
    "^{F3}", Map("payload", "^{F3}", "label", "Ctrl + F3", "kind", "key"),
    "^{F4}", Map("payload", "^{F4}", "label", "Ctrl + F4", "kind", "key"),
    "^{F5}", Map("payload", "^{F5}", "label", "Ctrl + F5", "kind", "key"),
    "^{F6}", Map("payload", "^{F6}", "label", "Ctrl + F6", "kind", "key"),
    "^{F7}", Map("payload", "^{F7}", "label", "Ctrl + F7", "kind", "key"),
    "^{F8}", Map("payload", "^{F8}", "label", "Ctrl + F8", "kind", "key"),
    "^{F9}", Map("payload", "^{F9}", "label", "Ctrl + F9", "kind", "key"),
    "^{F10}", Map("payload", "^{F10}", "label", "Ctrl + F10", "kind", "key"),
    "^{F11}", Map("payload", "^{F11}", "label", "Ctrl + F11", "kind", "key"),
    "^{F12}", Map("payload", "^{F12}", "label", "Ctrl + F12", "kind", "key"),
    "+{F1}", Map("payload", "+{F1}", "label", "Shift + F1", "kind", "key"),
    "+{F2}", Map("payload", "+{F2}", "label", "Shift + F2", "kind", "key"),
    "+{F3}", Map("payload", "+{F3}", "label", "Shift + F3", "kind", "key"),
    "+{F4}", Map("payload", "+{F4}", "label", "Shift + F4", "kind", "key"),
    "+{F5}", Map("payload", "+{F5}", "label", "Shift + F5", "kind", "key"),
    "+{F6}", Map("payload", "+{F6}", "label", "Shift + F6", "kind", "key"),
    "+{F7}", Map("payload", "+{F7}", "label", "Shift + F7", "kind", "key"),
    "+{F8}", Map("payload", "+{F8}", "label", "Shift + F8", "kind", "key"),
    "+{F9}", Map("payload", "+{F9}", "label", "Shift + F9", "kind", "key"),
    "+{F10}", Map("payload", "+{F10}", "label", "Shift + F10", "kind", "key"),
    "+{F11}", Map("payload", "+{F11}", "label", "Shift + F11", "kind", "key"),
    "+{F12}", Map("payload", "+{F12}", "label", "Shift + F12", "kind", "key"),
    "!{F1}", Map("payload", "!{F1}", "label", "Alt + F1", "kind", "key"),
    "!{F2}", Map("payload", "!{F2}", "label", "Alt + F2", "kind", "key"),
    "!{F3}", Map("payload", "!{F3}", "label", "Alt + F3", "kind", "key"),
    "!{F4}", Map("payload", "!{F4}", "label", "Alt + F4", "kind", "key"),
    "!{F5}", Map("payload", "!{F5}", "label", "Alt + F5", "kind", "key"),
    "!{F6}", Map("payload", "!{F6}", "label", "Alt + F6", "kind", "key"),
    "!{F7}", Map("payload", "!{F7}", "label", "Alt + F7", "kind", "key"),
    "!{F8}", Map("payload", "!{F8}", "label", "Alt + F8", "kind", "key"),
    "!{F9}", Map("payload", "!{F9}", "label", "Alt + F9", "kind", "key"),
    "!{F10}", Map("payload", "!{F10}", "label", "Alt + F10", "kind", "key"),
    "!{F11}", Map("payload", "!{F11}", "label", "Alt + F11", "kind", "key"),
    "!{F12}", Map("payload", "!{F12}", "label", "Alt + F12", "kind", "key")      
)

text_config := Map()






; ; #endregion