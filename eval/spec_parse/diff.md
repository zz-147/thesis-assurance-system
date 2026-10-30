# 规范 → 规则候选报告（教师端）

- 基础 profile：`school_default`（现有 38 条规则）
- 现有规则逐条评估覆盖：**38/38**
- 规范字符数：9455
- 抽取模型：`deepseek-chat`
- 候选结果：**keep 38 ｜ modify 0 ｜ new 6**；未覆盖现有规则 0 条；unmapped 23 条；丢弃 0 条

> LLM 总结：现有规则库基本覆盖规范要求，仅需新增图注释、表注释、附录标题、多项附录编号、外文资料翻译格式和注释格式等规则，其余内容要求无法程序化检查。

## 新增规则（new）（6）

### `FIG-003` 图注释格式  （P1/warning，置信度 0.90）
- 出处：L92
- 规范：注释置于图名下方，1.5倍行距，字体为小五号宋体，左对齐，首行缩进2个字符。
- expected：`{"position": "below_figure_title", "size_pt": 9, "east_asia": "宋体", "alignment": "left", "first_line_chars": 200, "line_spacing": 1.5}`
- 备注：规范明确要求图注释的格式，现有规则库未包含。

### `TAB-003` 表注释格式  （P1/warning，置信度 0.90）
- 出处：L95
- 规范：注释置于表格下方，1.5倍行距，字体为小五号宋体，左对齐，首行缩进2个字符。
- expected：`{"position": "below_table", "size_pt": 9, "east_asia": "宋体", "alignment": "left", "first_line_chars": 200, "line_spacing": 1.5}`
- 备注：规范明确要求表注释的格式，现有规则库未包含。

### `APP-001` 附录标题格式  （P1/error，置信度 0.90）
- 出处：L106
- 规范：“附录”二字用三号黑体居中打印，两字中间空2格，下一行是附录题目（四号黑体居中打印），附录内容格式同正文。
- expected：`{"title_text": "附录", "title_size_pt": 16, "title_east_asia": "黑体", "title_alignment": "center", "title_char_gap": 2, "subtitle_size_pt": 14, "subtitle_east_asia": "黑体", "subtitle_alignment": "center", "content_format_same_as_body": true}`
- 备注：规范明确要求附录标题格式，现有规则库未包含。

### `APP-002` 多项附录编号  （P1/warning，置信度 0.90）
- 出处：L107
- 规范：若有多项附录，附录题目可按顺序编号（附录A、附录B，……）。
- expected：`{"multiple_appendix_numbering": true, "pattern": "^附录[A-Z]"}`
- 备注：规范明确要求多项附录的编号规则，现有规则库未包含。

### `TRANS-001` 外文资料翻译格式  （P1/error，置信度 0.90）
- 出处：L109
- 规范：“外文资料译文”用三号黑体居中打印，下一行打印外文原文内容（小四Times New Roman字体），另起新页打印中文翻译（小四宋体）。
- expected：`{"title_text": "外文资料译文", "title_size_pt": 16, "title_east_asia": "黑体", "title_alignment": "center", "original_text_size_pt": 12, "original_text_ascii": "Times New Roman", "translation_text_size_pt": 12, "translation_text_east_asia": "宋体", "new_page_for_translation": true}`
- 备注：规范明确要求外文资料翻译的格式，现有规则库未包含。

### `NOTE-001` 注释格式  （P1/warning，置信度 0.80）
- 出处：L33
- 规范：注释：在毕业论文写作过程中，有些问题需要在正文之外加以阐述和说明的，作为注释放在当页页脚。
- expected：`{"position": "page_footer", "size_pt": 9, "east_asia": "宋体", "line_spacing": 1.0}`
- 备注：规范要求注释放在当页页脚，但未明确字体字号，根据上下文推断为小五号宋体，1倍行距。

## 修改现有规则（modify）（0）

## 保持不变（keep）（38）

### `PAGE-001` 纸张为 A4  （P0/error，置信度 1.00）
- 出处：L66
- 规范：一律用 A4 规格纸输出
- expected：`{"page_width_cm": 21.0, "page_height_cm": 29.7, "orientation": "portrait"}`
- 备注：LLM 评估：与现有规则一致

### `PAGE-002` 页边距、装订线、页眉页脚距边界  （P0/error，置信度 1.00）
- 出处：L66
- 规范：上下页边距 2.54cm，左右 2cm，左侧装订，装订线 1cm；页眉和页脚距边界均为 2cm
- expected：`{"margin_top_cm": 2.54, "margin_bottom_cm": 2.54, "margin_left_cm": 2.0, "margin_right_cm": 2.0, "gutter_cm": 1.0, "gutter_position": "left", "header_distance_cm": 2.0, "footer_distance_cm": 2.0}`
- 备注：LLM 评估：与现有规则一致

### `PAGE-003` 文档网格（每页 31 行、每行 35 字符）  （P0/warning，置信度 1.00）
- 出处：L66
- 规范：指定行和字符网格：中文小四号宋体、西文小四号 Times New Roman，每页 31 行、每行 35 字符，栏数 1，正文行距 1.5 倍
- expected：`{"chars_per_line": 35, "lines_per_page": 31, "columns": 1, "grid_font_cn": "宋体", "grid_font_ascii": "Times New Roman", "grid_size_pt": 12, "line_pitch_twips": 312}`
- 备注：LLM 评估：与现有规则一致

### `PAGE-004` 分节与页眉页脚继承  （P1/warning，置信度 1.00）
- 出处：L69–L70、L111–L124
- 规范：摘要→目录、前言→正文等部分存在页眉页脚/页码格式切换，需按节组织；封面、承诺书不显示页眉页码
- expected：`{"sections_min": 2, "cover_no_header_footer": true}`
- 备注：LLM 评估：与现有规则一致

### `BODY-001` 正文字体  （P0/error，置信度 1.00）
- 出处：L91
- 规范：正文：小四号宋体，西文小四号 Times New Roman
- expected：`{"size_pt": 12, "east_asia": "宋体", "ascii": "Times New Roman", "bold": false}`
- 备注：LLM 评估：与现有规则一致

### `BODY-002` 正文行距  （P0/error，置信度 1.00）
- 出处：L91
- 规范：正文 1.5 倍行距
- expected：`{"line_spacing": 1.5}`
- 备注：LLM 评估：与现有规则一致

### `BODY-003` 正文首行缩进 2 字符  （P0/error，置信度 1.00）
- 出处：L76、L91
- 规范：每段首行缩进 2 个字符
- expected：`{"first_line_chars": 200}`
- 备注：LLM 评估：与现有规则一致

### `BODY-004` 正文两端对齐  （P0/warning，置信度 1.00）
- 出处：L91
- 规范：两端对齐
- expected：`{"alignment": "justify"}`
- 备注：LLM 评估：与现有规则一致

### `HEAD-001` 一级标题（理工农：第N章）  （P0/error，置信度 1.00）
- 出处：L88
- 规范：“第N章”与标题文字之间空 1 格，三号黑体，居中，1.5 倍行距，段前段后各 1 行
- expected：`{"pattern": "^第[一二三四五六七八九十百零〇0-9]+章\\s+.+$", "size_pt": 16, "east_asia": "黑体", "ascii": "Times New Roman", "bold": false, "alignment": "center", "line_spacing": 1.5, "before_lines": 100, "after_lines": 100}`
- 备注：LLM 评估：与现有规则一致

### `HEAD-002` 二级标题（理工农：N.N）  （P0/error，置信度 1.00）
- 出处：L88
- 规范：“节”为二级标题，四号黑体左起顶格，序号与标题文字间空 1 格，1.5 倍行距，段前段后各 0.5 行
- expected：`{"pattern": "^\\d+\\.\\d+\\s+.+$", "size_pt": 14, "east_asia": "黑体", "bold": false, "alignment": "left", "line_spacing": 1.5, "before_lines": 50, "after_lines": 50}`
- 备注：LLM 评估：与现有规则一致

### `HEAD-003` 三级标题（理工农：N.N.N）  （P0/error，置信度 1.00）
- 出处：L88
- 规范：“小节”为三级标题，小四号黑体左起顶格，序号后空 1 格，1.5 倍行距，段前 0.5 行、段后无间距
- expected：`{"pattern": "^\\d+\\.\\d+\\.\\d+\\s+.+$", "size_pt": 12, "east_asia": "黑体", "bold": false, "alignment": "left", "line_spacing": 1.5, "before_lines": 50, "after_lines": 0}`
- 备注：LLM 评估：与现有规则一致

### `HEAD-004` 四级及以下标题  （P1/warning，置信度 1.00）
- 出处：L88
- 规范：四级标题用 N.N.N.N，小四号宋体加粗；四级以下依次用（1）→①→A→a，小四号宋体，左起顶格，1.5 倍行距，段前段后无间距
- expected：`{"l4": {"pattern": "^\\d+\\.\\d+\\.\\d+\\.\\d+\\s*.+$", "size_pt": 12, "east_asia": "宋体", "bold": true}, "l5_plus": {"patterns": ["^（\\d+）", "^[①-⑳]", "^[A-Z][.、]", "^[a-z][.、]"], "size_pt": 12, "east_asia": "宋体"}, "before_lines": 0, "after_lines": 0}`
- 备注：LLM 评估：与现有规则一致

### `HEAD-005` 文经管标题体系  （P1/error，置信度 1.00）
- 出处：L89
- 规范：章=“一、”（三号黑体，首行缩进 2 字符，段前 1 行、段后无间距）；节=“（一）”（四号黑体，首行缩进 2 字符，段前 0.5 行）；小节=“1.”（小四黑体，首行缩进 2 字符，段前段后无间距）
- expected：`{"l1": {"pattern": "^[一二三四五六七八九十]+、", "size_pt": 16, "east_asia": "黑体", "first_line_chars": 200, "before_lines": 100, "after_lines": 0}, "l2": {"pattern": "^（[一二三四五六七八九十]+）", "size_pt": 14, "east_asia": "黑体", "first_line_chars": 200, "before_lines": 50, "after_lines": 0}, "l3": {"pattern": "^\\d+[.、]", "size_pt": 12, "east_asia": "黑体", "first_line_chars": 200, "before_lines": 0, "after_lines": 0}}`
- 备注：LLM 评估：与现有规则一致

### `HEAD-006` 标题编号合法性与连续性  （P0/error，置信度 1.00）
- 出处：L88–L89、L10
- 规范：标题层级依次标出，编号必须连续、不得跳级/重复；目录中的标题要与正文一致
- expected：`{"no_level_skip": true, "no_duplicate_number": true, "sequential": true}`
- 备注：LLM 评估：与现有规则一致

### `HEAD-007` 前言/结论/致谢标题  （P0/error，置信度 1.00）
- 出处：L85–L86、L97–L99
- 规范：“前言”“结论”“致谢”用三号黑体，两字中间空 2 格；理工农居中，文经管首行缩进 2 字符；1.5 倍行距，段前 1 行；结论、致谢格式同前言
- expected：`{"texts": ["前言", "结论", "致谢"], "size_pt": 16, "east_asia": "黑体", "line_spacing": 1.5, "before_lines": 100, "alignment": {"sci_eng": "center", "humanities": "left"}}`
- 备注：LLM 评估：与现有规则一致

### `TITLE-001` 中文题目  （P0/error，置信度 1.00）
- 出处：L7、L72
- 规范：中文题目三号黑体，可分成 1 或 2 行居中打印，1.5 倍行距，段前段后各 1 行
- expected：`{"size_pt": 16, "east_asia": "黑体", "alignment": "center", "line_spacing": 1.5, "before_lines": 100, "after_lines": 100, "max_lines": 2}`
- 备注：LLM 评估：与现有规则一致

### `TITLE-002` 英文题目  （P0/error，置信度 1.00）
- 出处：L73
- 规范：英文题目三号首字母大写 Times New Roman，加粗，可分成 1–3 行居中，1.5 倍行距，段前段后各 1 行
- expected：`{"size_pt": 16, "ascii": "Times New Roman", "bold": true, "alignment": "center", "line_spacing": 1.5, "before_lines": 100, "after_lines": 100, "max_lines": 3}`
- 备注：LLM 评估：与现有规则一致

### `ABS-001` 中文摘要标题  （P0/error，置信度 1.00）
- 出处：L75
- 规范：中文题目下一行居中打印“摘要”二字（三号黑体），字间空 2 格，1.5 倍行距，段前段后各 1 行
- expected：`{"text": "摘要", "size_pt": 16, "east_asia": "黑体", "alignment": "center", "line_spacing": 1.5, "before_lines": 100, "after_lines": 100, "char_gap": 2}`
- 备注：LLM 评估：与现有规则一致

### `ABS-002` 中文摘要正文  （P0/error，置信度 1.00）
- 出处：L76、L8、L37
- 规范：摘要内容小四号宋体，每段首行缩进 2 字符，1.5 倍行距，两端对齐；中文摘要 300～600 字（设计类 200～400 字）
- expected：`{"size_pt": 12, "east_asia": "宋体", "ascii": "Times New Roman", "first_line_chars": 200, "line_spacing": 1.5, "alignment": "justify", "length_range": [300, 600]}`
- 备注：LLM 评估：与现有规则一致

### `ABS-003` 英文摘要标题  （P0/error，置信度 1.00）
- 出处：L77
- 规范：英文题目下 1 行居中打印“Abstract”（三号 Times New Roman），1.5 倍行距，段前段后各 1 行
- expected：`{"text": "Abstract", "size_pt": 16, "ascii": "Times New Roman", "alignment": "center", "line_spacing": 1.5, "before_lines": 100, "after_lines": 100}`
- 备注：LLM 评估：与现有规则一致

### `ABS-004` 英文摘要正文  （P0/error，置信度 1.00）
- 出处：L78
- 规范：英文摘要内容小四号 Times New Roman，每段首行缩进 2 字符，1.5 倍行距，两端对齐
- expected：`{"size_pt": 12, "ascii": "Times New Roman", "first_line_chars": 200, "line_spacing": 1.5, "alignment": "justify"}`
- 备注：LLM 评估：与现有规则一致

### `KW-001` 中文关键词  （P0/error，置信度 1.00）
- 出处：L80、L9
- 规范：中文摘要内容下 1 行打印“关键词”三字（小四号黑体，左对齐，首行缩进 2 字符，1.5 倍行距，段前段后各 1 行），其后关键词小四号宋体；3–5 个，分号隔开，最后一个后不加标点
- expected：`{"label": "关键词", "label_size_pt": 12, "label_east_asia": "黑体", "body_size_pt": 12, "body_east_asia": "宋体", "count_range": [3, 5], "separator": "；", "trailing_punctuation": false, "first_line_chars": 200, "before_lines": 100, "after_lines": 100}`
- 备注：LLM 评估：与现有规则一致

### `KW-002` 英文关键词  （P0/error，置信度 1.00）
- 出处：L81
- 规范：英文摘要内容下一行打印“Key words”（小四号 Times New Roman，加粗，左对齐，首行缩进 2 字符，段前段后各 1 行），其后关键词小四号 Times New Roman；分号（加 1 个半角字符）隔开，最后一个后无标点
- expected：`{"label": "Key words", "label_size_pt": 12, "label_ascii": "Times New Roman", "label_bold": true, "body_size_pt": 12, "body_ascii": "Times New Roman", "count_range": [3, 5], "separator": "; ", "trailing_punctuation": false}`
- 备注：LLM 评估：与现有规则一致

### `TOC-001` 目录标题与条目样式  （P0/error，置信度 1.00）
- 出处：L82–L83、L10
- 规范：“目录”二字三号黑体居中，两字中间空 2 格，1.5 倍行距，段前段后各 1 行；一级标题小四号黑体，其他级别小四号宋体；阿拉伯数字用 Times New Roman；一级顶格、二级缩进 2 字符、三级缩进 4 字符
- expected：`{"title": "目录", "title_size_pt": 16, "title_east_asia": "黑体", "title_alignment": "center", "entry_size_pt": 12, "level1_east_asia": "黑体", "others_east_asia": "宋体", "number_ascii": "Times New Roman", "indent_chars": {"1": 0, "2": 200, "3": 400}}`
- 备注：LLM 评估：与现有规则一致

### `TOC-002` 目录完整性  （P1/warning，置信度 1.00）
- 出处：L10、L83
- 规范：目录列出论文的一级、二级和三级标题并标明页码；目录中的标题与正文标题一致，页码与正文一致
- expected：`{"includes_levels": [1, 2, 3], "matches_headings": true, "page_numbers_updateable": true}`
- 备注：LLM 评估：与现有规则一致

### `HF-001` 页眉样式  （P1/error，置信度 1.00）
- 出处：L69
- 规范：从摘要到最后一页均有页眉；小五号宋体，1 倍行距，居中；文字下边框为双实线 0.5 磅
- expected：`{"size_pt": 9, "east_asia": "宋体", "line_spacing": 1.0, "alignment": "center", "bottom_border": {"style": "double", "size_pt": 0.5}, "starts_at": "摘要"}`
- 备注：LLM 评估：与现有规则一致

### `HF-002` 页眉文字  （P1/warning，置信度 1.00）
- 出处：L69
- 规范：页眉文字与封面对照，分别为“商丘学院应用科技学院毕业论文”或“商丘学院应用科技学院毕业设计说明书”
- expected：`{"allowed_texts": ["商丘学院应用科技学院毕业论文", "商丘学院应用科技学院毕业设计说明书"], "match": "与封面/题目类型一致"}`
- 备注：LLM 评估：与现有规则一致

### `HF-003` 页码格式  （P1/error，置信度 1.00）
- 出处：L70
- 规范：从摘要到目录页码为罗马数字；从前言到最后一页为阿拉伯数字；字体小五号 Times New Roman，1 倍行距，页面底部居中
- expected：`{"roman_section": "摘要→目录", "arabic_section": "前言→末尾", "size_pt": 9, "ascii": "Times New Roman", "line_spacing": 1.0, "alignment": "center", "position": "bottom"}`
- 备注：LLM 评估：与现有规则一致

### `FIG-001` 图题（图号+图名）  （P0/error，置信度 1.00）
- 出处：L92
- 规范：图名置于图下方居中；图名序号五号 Times New Roman，文字五号楷体，序号与文字间空 1 个半角字符，1.5 倍行距
- expected：`{"position": "below", "alignment": "center", "number_size_pt": 10.5, "number_ascii": "Times New Roman", "text_size_pt": 10.5, "text_east_asia": "楷体", "separator": " ", "line_spacing": 1.5}`
- 备注：LLM 评估：与现有规则一致

### `FIG-002` 图编号规则与连续性  （P1/warning，置信度 1.00）
- 出处：L61、L93
- 规范：理工农：图号逐章单独编序（如图6-10）；文经管：按出现顺序编号（图1、图2…）；图号必须连续，不得重复或跳跃；仅有一图时图名前加“附图”
- expected：`{"sci_eng": "^图\\d+-\\d+", "humanities": "^图\\d+", "continuous": true, "unique": true}`
- 备注：LLM 评估：与现有规则一致

### `TAB-001` 三线表样式  （P1/warning，置信度 1.00）
- 出处：L60、L95
- 规范：表格为三线表；表格结构简洁，各栏标注量和单位；表内数字上下对齐，不得用“同上/同左”
- expected：`{"style": "three_line", "header_row_rule": "0.75pt bottom border", "no_vertical_borders": true, "no_inner_horizontal_borders": true, "top_bottom_border_pt": 1.5}`
- 备注：LLM 评估：与现有规则一致

### `TAB-002` 表题与表内文字  （P1/error，置信度 1.00）
- 出处：L95、L60
- 规范：表名置于表格上方居中；表名序号五号 Times New Roman，标题五号楷体，中间空 1 个半角字符，1.5 倍行距；表内文字五号宋体，1.5 倍行距，居中；注释在表格下方小五号宋体、左对齐、首行缩进 2 字符
- expected：`{"position": "above", "alignment": "center", "number_size_pt": 10.5, "number_ascii": "Times New Roman", "text_size_pt": 10.5, "text_east_asia": "楷体", "cell_size_pt": 10.5, "cell_east_asia": "宋体", "cell_alignment": "center", "cell_line_spacing": 1.5, "note_size_pt": 9, "note_first_line_chars": 200}`
- 备注：LLM 评估：与现有规则一致

### `EQ-001` 公式与公式编号  （P1/warning，置信度 1.00）
- 出处：L62、L96
- 规范：公式另起一行，居中（或自行调整美观）；重要/后文引用的公式加序号并用圆括号，序号右对齐，序号用阿拉伯数字；公式字体小四号，中文宋体，西文 Times New Roman，1.5 倍行距；公式与序号之间不加线段
- expected：`{"own_line": true, "number_in_parentheses": true, "number_alignment": "right", "size_pt": 12, "east_asia": "宋体", "ascii": "Times New Roman", "line_spacing": 1.5, "numbering": {"sci_eng": "^（\\d+-\\d+）$", "humanities": "^（\\d+）$"}}`
- 备注：LLM 评估：与现有规则一致

### `REF-001` 参考文献标题  （P0/error，置信度 1.00）
- 出处：L102
- 规范：“参考文献”四个字三号黑体，居中，1.5 倍行距，段前段后各 1 行
- expected：`{"text": "参考文献", "size_pt": 16, "east_asia": "黑体", "alignment": "center", "line_spacing": 1.5, "before_lines": 100, "after_lines": 100}`
- 备注：LLM 评估：与现有规则一致

### `REF-002` 参考文献条目格式  （P1/error，置信度 1.00）
- 出处：L102–L104
- 规范：条目中文字体小四号宋体、西文小四号 Times New Roman，1.5 倍行距，段前段后无间距；序号写在方括号内，每条左对齐，悬挂缩进 2 字符
- expected：`{"size_pt": 12, "east_asia": "宋体", "ascii": "Times New Roman", "line_spacing": 1.5, "before_lines": 0, "after_lines": 0, "number_format": "^\\[\\d+\\]", "hanging_indent_chars": 200}`
- 备注：LLM 评估：与现有规则一致

### `REF-003` 正文引用标注与著录格式  （P1/warning，置信度 1.00）
- 出处：L103–L104、L17–L33
- 规范：正文中按引证先后顺序用阿拉伯数字连续编号，序号置于方括号内作上角标（如“……[5]”）；按 GB/T 7714-2015 著录（8 类示例）
- expected：`{"in_text_citation_superscript": true, "sequential_numbering": true, "style": "GB/T 7714-2015", "categories": ["J", "M", "C", "D", "P", "S", "N", "DB/CP/EB"]}`
- 备注：LLM 评估：与现有规则一致

### `STRUCT-001` 装订顺序完整  （P1/warning，置信度 1.00）
- 出处：L111–L124
- 规范：按顺序装订：封面→诚信承诺书→中文摘要→英文摘要→目录→正文→致谢→参考文献→附录→外文资料翻译→封底
- expected：`{"order": ["封面", "诚信承诺书", "中文摘要", "英文摘要", "目录", "正文", "致谢", "参考文献", "附录", "外文资料翻译", "封底"]}`
- 备注：LLM 评估：与现有规则一致

### `MISC-001` 数字用法  （P1/info，置信度 1.00）
- 出处：L63
- 规范：公历世纪、年代、年月日、时间、计数计量用阿拉伯数字；年份不能简写；有效数字全部写出
- expected：`{"year_full": true, "significant_digits": true}`
- 备注：LLM 评估：与现有规则一致

## unmapped：无法程序化检查的规范要求（需人工/流程处理）

- L8 摘要应具有独立性和自含性，即不阅读论文全文也可获得必要的信息，是一篇完整的短文，可独立使用，可引用，可用于工艺推广。（内容要求，无法程序化检查。）
- L8 摘要中不使用“本文”、“作者”等作为主语，应采用“对…进行了研究”、“报告了…现状”、“进行了…调查”等表达方式。（写作风格要求，无法程序化检查。）
- L8 英文摘要应与中文摘要内容相符。（内容一致性检查，难以程序化。）
- L9 关键词是为了文献标引工作从论文中选取出来用以表示全文主体内容信息款目的单词或术语。（定义性描述，无法程序化检查。）
- L9 从摘要中挑选或概括出3～5个最能表达主要内容的词作为关键词。（关键词选取质量，无法程序化检查。）
- L10 目录中的标题要与正文中的标题一致，页码要与正文中的页码相一致。（一致性检查需要复杂逻辑，现有规则TOC-002已部分覆盖，但完全程序化困难。）
- L12 前言（引言）：是论文的开头，主要说明国内外相关研究现状及现实意义、对所研究问题的认识，并提出论文的中心论点等。（内容要求，无法程序化检查。）
- L13 本论：是毕业论文的主体，包括研究内容与方法、结果与分析（讨论）、结论与建议等。在本部分要对所运用分析解决问题的理论和方法、研究内容与结果、论证的观点等作出阐述，写作时应注意区别那些是已知的（前人已有的的研究），哪些是作者首次提出来的，哪些是经过作者改进的，须交代清楚，尽量反映出自己的研究能力和学术水平。（内容与写作要求，无法程序化检查。）
- L14 结论：是毕业论文的收尾部分，是围绕本论所作的结束语。其基本要点就是总结全文、加深题意，要概括准确、措辞严谨、语意清楚，不能模棱两可、含糊其词，且肯定和否定要明确，突出研究的新进展或主要结论性成果。（内容与写作要求，无法程序化检查。）
- L15 致谢：简述自己完成毕业论文工作的体会，并对指导教师和协助完成毕业论文的有关人员表达谢意。（内容要求，无法程序化检查。）
- L16 参考文献一般要著录最必要、最新的文献，且只著录公开发表的文献。（文献选择要求，无法程序化检查。）
- L16 参考文献中作者不超过三位时全部列出，超过三位时只列前三位，后面加“等”或“etc”；作者姓名之间用逗号隔开；中外人名一律采用姓在前、名在后的著录法。（著录细节，部分可程序化但复杂，现有规则REF-003已覆盖部分，但作者数量等未完全覆盖。）
- L60 表格中各栏都应标注量和相应的单位。（内容要求，无法程序化检查。）
- L60 表格内数字须上下对齐，相邻栏内的数值相同时，不能用“同上”、“同左”和其它类似词汇，应一一重新标注。（内容与格式细节，难以程序化检查。）
- L94 绘图必须工整、清晰、规范。其中机械零件图按机械制图规格要求；示意图应能清楚反映图示内容。（质量要求，无法程序化检查。）
- L96 公式原则上居中，若公式较长，请自定，美观即可。（美观性判断，无法程序化检查。）
- L64 软件程序代码清单要按软件文档格式附在论文后面，不附在论文内。（内容组织要求，无法程序化检查。）
- L66 毕业论文（设计说明书）应采用汉语（外语专业用外语）撰写。（语言选择，无法程序化检查。）
- L67 封面（按规定格式打印）（格式由模板规定，无法程序化检查。）
- L68 承诺书（按规定格式打印，手写签字）（手写签字，无法程序化检查。）
- L66 用统一封面装订成册。（物理装订要求，无法程序化检查。）
- L112-L124 毕业论文或毕业设计说明书按以下顺序装订：...（装订顺序，现有规则STRUCT-001已覆盖，但物理装订无法程序化检查。）
- L125 个别专业可以结合本专业特点参照执行。（灵活性条款，无法程序化检查。）

---
审阅方式：先看 `diff.md`，修改/删除 `candidates.json` 里不认可的候选，
再执行 `tas spec apply candidates.json --id <profile_id> --name <名称>` 生成新 profile。
注意：new 规则在绑定检查器之前不会参与检查（只进入规则库，待实现）。