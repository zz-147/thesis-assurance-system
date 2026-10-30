# TAS · 论文质量保障系统（Thesis Assurance System）

> 一个"规范文档驱动"的论文格式检查 / 安全修复 / 真实性核验工具，面向高校教师批量收作业、给反馈的场景：
> **教师端**：规范文档 → LLM 规则候选 → 教师确认成 profile → 批量体检学生论文 → 逐份报告 + 班级汇总 + 名单核对 + 文献真实性核验 + 首页缩略图；
> **学生端**：上传论文 → 按规范检查 → 一键安全修复格式（只改格式、不改内容、不动文献）→ 下载新文件；
> **Web 双入口**：本地起服务即可用（无账号），论文 / 规范文档都可以从桌面拖拽上传，检查/修复结果支持页面预览（Word 渲染）。
>
> **项目名**：TAS = Thesis Assurance System（论文质量保障系统）。Python 包名与 CLI 命令保留为 `tas`（历史缩写，兼容已有脚本与环境变量）；GitHub 仓库名为 `thesis-assurance-system`。

- 规范示例：商丘学院应用科技学院《毕业论文（设计）撰写规范（修订）》
- 状态（2026-10-30）：**功能补齐 + 规则全覆盖 + 可选增强**。规则库 38 条、检查器覆盖 **38/38（P0 23/23，P1 15/15）**；LLM profile 44/44；**321 个测试全绿**（含真实网络/LLM/Word 渲染 live）；Web 双入口 / 批量汇总导出 / 评测指标 / 多学科参数化 / 文档接龙 API+H5 / 老师端一键启动 + 学生反馈包 + 名单核对 + 批量安全修复 + 免安装绿色版 + PAGE-004/NOTE-001 检查器 + TAB-002 自动修复 + **批量首页缩略图 + 批量文献真实性核验** 完成；微信小程序源码在 [miniprogram/](miniprogram/)（上传场景暂缓，保留源码与网络自检）
- 文档：[需求](docs/REQUIREMENTS.md)｜[计划/架构](docs/PLAN.md)｜[**简历指标与评测**](docs/METRICS.md)｜[**老师端使用指南**](docs/TEACHER_GUIDE.md)｜[**小程序/接龙方案**](docs/MINIPROGRAM.md)｜[规则库](rules/rules.json)｜[profile 索引](profiles/index.json)
- 演示：[一页总结 PDF](demo/演示报告_一页总结.pdf)｜[演示资料包](demo/tas演示资料.zip)｜[演示截图](demo/截图/)

## 快速开始（Windows / PowerShell）

环境要求：Python 3.10+；页面预览 / 批量缩略图需要 Windows + Microsoft Word（可选，失败不影响检查/修复）。

```powershell
git clone https://github.com/zz-147/thesis-assurance-system.git
Set-Location thesis-assurance-system
python -m pip install -e .
$PY = "python"
$env:PYTHONIOENCODING = "utf-8"

# 0) Web 双入口（推荐演示；浏览器打开 http://127.0.0.1:8765）
#    ★ 老师日常使用：双击项目根目录「启动老师端.bat」，自动开浏览器（关闭黑窗口=停止）
#    学生端：把 .docx 从桌面直接拖进页面 → 检查 → 一键修复 → 下载；默认勾选"页面预览"
#    教师端：多份 .docx 一起拖入 / 选择整个文件夹 → 批量检查 → 下载 xlsx/html + 学生反馈包
#    规范管理：把规范文档（.docx/.txt/.md）拖入 → LLM 抽取规则 → 审阅 → 生成新 profile
#    文档接龙（不拆教师/学生端）：浏览器/手机打开 http://127.0.0.1:8765/m
#    微信小程序：用微信开发者工具导入 miniprogram/（见 docs/MINIPROGRAM.md）
#    真机联调：& $PY -m tas web --host 0.0.0.0 --port 8765 --verbose
#              手机浏览器打开 http://<电脑IP>:8765/api/diag 判断网络是否通
& $PY -m tas web --port 8765

# 1) 学生端：检查 + 一键修复（默认输出 原名_fixed.docx，绝不覆盖原文件；--ruleset 换规范）
& $PY -m tas check "某论文.docx" --format text
& $PY -m tas fix   "某论文.docx" --report fix-report.json
& $PY -m tas fix   "某论文.docx" --ruleset course_demo_other

# 2) 教师端：批量检查 + 班级汇总 + 导出 xlsx/csv/html + 学生反馈包 feedback.zip
#    可选：--fix 生成全部安全修复版（fixed.zip）；--roster 名单核对（缺交/重复/命名）
#          --preview 首页缩略图（需 Word）；--refs 文献真实性核验（联网）
& $PY -m tas batch "D:\docs\某班论文目录" --out-dir eval/batch_class --fix --roster 名单.txt --preview --refs

# 3) 教师端：文献真实性核验 + 引用一致性（只读，不改文档）
& $PY -m tas refs "某论文.docx"
& $PY -m tas refs "某论文.docx" --verify --limit 5

# 4) 教师端：规范文档 → LLM 规则候选 → 教师确认 → 新 profile（DeepSeek）
#    先设置 key：setx TAS_LLM_API_KEY "sk-..."（不要写进仓库）
& $PY -m tas spec parse "规范文档.docx" --base-profile school_default --out-dir eval/spec_parse
& $PY -m tas spec apply "eval/spec_parse/candidates.json" --id course_xxx --name "XX课程格式要求"

# 5) 页面预览（Word COM 渲染前 N 页为 PNG，Web 结果页自动展示）
& $PY -m tas preview "某论文.docx" --out-dir eval/preview --pages 2

# 6) 多学科 profile / 规则库
& $PY -m tas profiles list
& $PY -m tas rules --phase P0

# 7) 测试（离线 317 通过 + 4 个 live/渲染默认 skip；开变量后 321/321）
& $PY -m unittest discover -s tests -v
$env:TAS_LIVE_NETWORK = "1"; $env:TAS_LIVE_LLM = "1"; $env:TAS_RENDER_TEST = "1"
& $PY -m unittest discover -s tests
Remove-Item Env:\TAS_LIVE_NETWORK, Env:\TAS_LIVE_LLM, Env:\TAS_RENDER_TEST
```

> 环境变量：`TAS_LLM_API_KEY`（DeepSeek key，必填于 spec parse）、`TAS_LLM_MODEL`（默认 `deepseek-chat`）、`TAS_LLM_BASE_URL`（默认 `https://api.deepseek.com/v1`）。

## 老师一分钟上手（收文件 → 批量检查 → 发反馈）

> 详细图文步骤见 **[docs/TEACHER_GUIDE.md](docs/TEACHER_GUIDE.md)**。

1. **双击项目根目录的 `启动老师端.bat`** → 自动打开 `http://127.0.0.1:8765`（黑窗口不要关；端口占用时双击 `停止老师端.bat`）
2. 学生发来的 `.docx` 存到电脑同一个文件夹（文件名建议 学号+姓名）
3. 浏览器 → 教师端 → 选规范 → **「选择整个文件夹（只收 .docx）」**
4. 可选（推荐）：粘贴**名单**（每行「学号 姓名」）核对缺交/重复/命名；勾选**「同时生成安全修复版」**批量生成 `*_fixed.docx`
5. 点 **「批量检查」** → 看通过率 / 错误警告数 / 共性问题 Top / 逐份结果
6. 下载 **「学生反馈（zip）」**，解压后：
   - `00_班级汇总.txt`：班级通过率 + 共性问题 + 未通过名单
   - `反馈_学号_姓名.txt`：每份文档一段（结论 + 主要问题 + 修改建议），直接发给学生
7. 若勾选了名单/修复，另有：
   - `roster-report.txt`：缺交/重复/命名提醒
   - `fixed.zip`：全部修复版 + 修复报告（只改格式，内容 hash 不变）
8. 另可下载 `batch.xlsx`（Excel 汇总存档）或 `batch.html`（网页报告，含条形图/通过率可视化）

> 只需 Windows + Python（启动器自动找；首次需装 `python-docx/lxml/openpyxl`）；只做检查不需要联网，也不需要 Word（预览/一键修复才需要 Word）。
> **没有 Python 的电脑**：直接发 `dist\tas-web-portable.zip`（约 19 MB，解压双击即用，内置 Python 运行时）。

## 文档接龙（H5 / 微信小程序）

按班级真实场景设计：**不拆教师端/学生端**——任何人可发起接龙，同学通过分享进入同一接龙提交 `.docx`，任何人可「一键审核」格式是否规范。

- **手机 H5（立刻可试）**：`tas web` 后浏览器/手机访问 **`http://127.0.0.1:8765/m`**，完整走"发起接龙 → 提交 .docx → 一键审核 → 查看报告"。
- **微信小程序**：源码在 [miniprogram/](miniprogram/)，用微信开发者工具导入；从聊天记录选文件提交、分享接龙卡片、一键审核。
- 数据默认存本机 `data/`（JSON + 文档，已 gitignore）；无账号，昵称仅展示。
- 设计、API、微信上线要求（AppID / HTTPS 域名 / 10MB 限制 / 云存储）见 **[docs/MINIPROGRAM.md](docs/MINIPROGRAM.md)**。

## 能力清单

| 能力 | 命令 | 状态 |
|---|---|---|
| 格式检查（38 条规则，文本/JSON/MD 报告） | `tas check` | ✅ 38/38（P0 23/23，P1 15/15） |
| 安全修复（21 条 safe 规则，含 TAB-002 表内文字；只改格式；内容 hash 不变；文献/目录条目零改动；自动复检；幂等） | `tas fix` | ✅ |
| 规则检查覆盖 | 38/38（P0 23/23，P1 15/15）；LLM profile 44/44 | `tas check` | ✅ |
| 文献真实性核验（Crossref → OpenAlex → arXiv；verified/likely/not_found/error + 证据 URL） | `tas refs --verify` | ✅ |
| 引用一致性（越界/缺号/未引用/首次顺序/上标/文献表编号连续性） | `tas refs` | ✅ |
| 多学科规则集（profile） | `tas profiles list` / `--ruleset` | ✅ |
| **规范→规则**：LLM 逐条评估 + 新增 + unmapped + diff + 注册 profile | `tas spec parse/apply` | ✅ DeepSeek 实测 |
| **参数化多规范**：checker/fixer 消费 profile 参数（页面/正文/标题/前置全部） | `--ruleset <id>` | ✅ |
| **教师端批量 + 班级汇总 + 导出** xlsx（3 Sheet）/csv/html/**学生反馈包** | `tas batch` | ✅ |
| **老师端一键启动**：双击 bat 启动本地 Web 并自动开浏览器 | `启动老师端.bat` | ✅ 冒烟通过 |
| **学生反馈包**：批量检查后每份文档一个 `.txt`（结论+主要问题+修改建议）+ 班级汇总 | 教师端结果页 / `feedback.zip` | ✅ 真实语料实测 |
| **名单核对**：粘贴/上传名单 → 缺交、重复提交、命名不规范 | 教师端「名单核对」/ `tas batch --roster` | ✅ 34 tests |
| **批量安全修复**：一次生成全部 `*_fixed.docx`（只改格式，内容 hash 不变，原文件不动） | 教师端勾选 / `tas batch --fix` | ✅ 12 tests |
| **批量首页缩略图**：每份文档首页 PNG，嵌入 HTML 报告 + `thumbnails.zip` | 教师端勾选 / `tas batch --preview` | ✅ 27 tests |
| **批量文献真实性核验**：verified/likely/not_found/error + 证据 URL + JSON/MD 报告 | 教师端勾选 / `tas batch --refs` | ✅ 14 tests |
| **免安装绿色版**：内置 Python 运行时，解压双击即用（约 19 MB zip） | `dist\tas-web-portable.zip` | ✅ 实测 CLI+Web |
| **报告可视化**：HTML 报告含共性问题条形图 + 通过率进度条 | `batch.html` | ✅ |
| **Web 双入口**（教师批量 / 学生检查修复，无账号，仅本机；**拖拽上传 + 页面预览**） | `tas web` | ✅ |
| **Web 规范管理**：规范文档拖拽上传 → LLM 抽取规则候选 → 审阅 diff → 生成/注册 profile | Web 首页第三块 | ✅ 真实 DeepSeek 实测 |
| **文档接龙**：发起接龙 → 同学提交 .docx → 一键审核格式（不拆教师/学生端） | `tas web` + `/m`；小程序源码 [miniprogram/](miniprogram/) | ✅ API+H5（9 tests）；小程序源码提供 |
| **页面预览**（docx → PDF → PNG，检查/修复结果可视化） | `tas preview` / Web 默认勾选 | ✅ Windows + Word |
| **评测指标脚本**（检出率/误报/修复通过率/吞吐） | `python eval/run_eval.py` | ✅ |

## 关键验收证据（详见 [docs/METRICS.md](docs/METRICS.md)）

- **规则覆盖**：38 条规则库，检查器实现 **38/38（P0 23/23，P1 15/15）**；LLM 抽取 profile 44 条全部实现（含 NOTE-001 脚注）。
- **合成变异评测**：116 处注入 / 23 类 → **检出率 100%，定位准确率 100%**；合规 fixture 误报 **0/201 段**；修复后 P0 通过率 **100%**、内容 hash 一致率 **100%**；单份检查 0.09s、修复 0.253s、批量吞吐 **9.80 份/秒**。
- **真实语料批量**：`out/` 9 份 → 111 error / 193 warning / 8 info；Top 共性问题 BODY-004 176 次、TAB-002 69 次、STRUCT-001 16 次；xlsx/csv/html 全部导出可打开。
- **规范→规则（真实 DeepSeek）**：学院规范 9455 字 → 11s、**38/38** 逐条评估（keep 38 / new 6 / unmapped 23）；课程规范 418 字 → 4.5s（keep 33 / modify 5）。
- **多规范生效**：同一论文学院默认 vs 课程自定义，错误数 **5 → 299**；自定义修复后实测 10.5pt 仿宋 / 1.75 倍行距 / 左右 3cm / 一级标题 14pt，复检 0、内容 hash 一致。
- **修复安全**：岳希华论文 26 项安全修复 → P0 清零；331 段文字、12 条参考文献、42 条目录条目 XML 逐字节不变；原文件不覆盖。
- **文献真实性（真实网络）**：3 条示例 → verified（ResNet，DOI 10.1109/cvpr.2016.90）/ likely（题名泛 + 年份不符提示）/ not_found（中文期刊未收录，标注"需人工核实"）。
- **页面预览（真实 Word 渲染）**：`tas preview` 单份论文渲染 2 页 PNG（约 13 秒，Word 启动为主）；Web 检查/修复页直接展示预览图（PNG 下载 200 / image/png / 120KB+）。
- **Web 规范管理（真实 DeepSeek）**：Web 首页拖入 418 字课程规范 → 抽取结果页（keep 33 / modify 5 / unmapped 1）→ 生成 profile → 下载规则文件 38 条，BODY-001 已是 10.5pt 仿宋；全程在页面上完成，无需 CLI。
- **交互**：Web 支持从桌面拖拽 .docx（学生端单文件、教师端多文件），**规范文档同样支持拖拽（.docx/.txt/.md）**；拖入即显示文件 chips；无 JS 时仍可点击选择。
- **文档接龙（本地 API 实测）**：创建接龙 → 提交一份变异文档 → 立即返回"未通过 + 错误数" → 一键审核（通过率 0%、共性问题 Top、逐份状态）→ 单份完整报告；接口统一 JSON，错误可读（9 个测试覆盖）。
- **小程序模拟器实测（2026-10-26）**：微信开发者工具 2.02.0 导入项目（AppID `wx9e3d7bdd9b6ae328`）+ 真实后端，跑通首页/发起接龙/详情/一键审核/报告/提交展示；真实论文 2 份审核 0% 通过（错误 620/警告 85，Top BODY-001 235），合规文档提交后 100% 通过。`wx.chooseMessageFile` 原生选文件需真机验证（模拟器不可 mock）。截图见 `eval/mp_screenshots/`（本机）。
- **真机网络自检（2026-10-27）**：新增 `GET /api/diag`（返回服务端看到的客户端 IP/监听地址）、`tas web --verbose` 请求日志（每条请求 + 上传文件名/大小/来源 IP）、小程序首页「网络自检」按钮；真机连不上时可按 4 层定位（后端监听 → 手机浏览器 `/api/diag` → 小程序自检 → 后端请求日志），详见 [MINIPROGRAM.md §11](docs/MINIPROGRAM.md)。
- **老师收文件工作流（2026-10-30）**：双击 `启动老师端.bat`（自动起服务+开浏览器，健康检查 200）；教师端支持「选择整个文件夹」批量收 `.docx`；9 份真实文档批量检查后生成 `feedback.zip`（班级汇总 + 9 份逐份反馈，含结论/问题统计/前 11 条问题/修改建议），可直接转发学生。
- **名单核对 + 批量安全修复（2026-10-30）**：3 人名单 + 2 份文档 → 缺交 1 人（王五）、已匹配 2 人；勾选批量修复后 `fixed.zip` 含 2 份 `_fixed.docx` + `修复报告.txt`，错误数 1 → 0；Web 与 CLI（`--fix --roster`）均端到端验证。
- **免安装绿色版（2026-10-30）**：内置 Python 3.12 运行时 + 依赖，目录 50.8 MB / zip 18.5 MB；绿色版跑 `tas check` 真实论文 exit 0、起 Web `/health` 200、读到 4 个规范；无需安装 Python、无需联网。
- **报告可视化（2026-10-30）**：`batch.html` 增加共性问题条形图 + 通过率进度条，纯 CSS（无新依赖）。
- **P2 规则补齐（2026-10-30）**：PAGE-004 分节/封面页眉页码检查（真实 9 份语料 0 误报）、NOTE-001 脚注格式（直读 `word/footnotes.xml`；LLM profile 44/44）、TAB-002 表内文字自动修复（12 个测试；单元格字号/字体/对齐/行距归一化，幂等、hash 不变）。P2 新增 12 个检查器测试 + 4 个 fixer 测试。
- **批量首页缩略图 + 文献核验（2026-10-30）**：`tas batch --preview` 生成每份首页 PNG（CLI 冒烟：HTML 内嵌 `<img>` + `thumbnails.zip` 56KB，含说明）；`--refs` 对真实论文 12 条文献核验（每份前 2 条：verified 0 / likely 0 / not_found 2 / error 0，报告含"人工核实"提示）；Web 教师端可勾选两个开关并在结果页展示缩略图网格与文献汇总。
- **测试**：**321 个全绿**（离线 317 + 4 个 live/渲染默认 skip；开 `TAS_LIVE_NETWORK/TAS_LIVE_LLM/TAS_RENDER_TEST` 后全部通过）。

## 真实性核验的边界（重要）

| 能做 | 不能做（需外部服务/人工） |
|---|---|
| 文献元数据是否真实存在（给 DOI/链接/证据） | 查重（知网/维普/Turnitin） |
| 引用一致性（编号、顺序、上标、未引用） | AI 生成内容检测（AIGC） |
| 文献表编号连续性 | 判定"学术不端/造假" |

- 红线：**未检索到 ≠ 造假**；中文期刊可能未被 Crossref/OpenAlex 收录，报告固定提示"需人工核实"。
- 只把题名/作者发给公开元数据 API，不上传论文全文。
- **参考文献一律不自动修改**（学生端修复白名单不含任何文献/文字操作）。

## 规则覆盖（38 / 38）

- 页面：PAGE-001 A4、PAGE-002 页边距/装订线、PAGE-003 文档网格、PAGE-004 分节与页眉页脚继承（封面不显示页眉页码）
- 正文：BODY-001 字体字号、BODY-002 行距、BODY-003 首行缩进、BODY-004 对齐
- 标题：HEAD-001/002/003 一/二/三级、HEAD-004 四级及以下、HEAD-005 文经管体系、HEAD-006 编号连续性、HEAD-007 前言/结论/致谢
- 前置：TITLE-001/002、ABS-001..004、KW-001/002
- 图表公式：FIG-001/002/003、TAB-001/002/003、EQ-001
- 页眉页脚：HF-001/002/003（页眉样式/文字/页码格式）
- 文献/目录/结构：REF-001/002/003、TOC-001/002、STRUCT-001、MISC-001
- 新规则（LLM profile）：APP-001/002 附录、TRANS-001 外文翻译、NOTE-001 脚注（直读 footnotes.xml，9pt 宋体/1 倍行距）均已实现

## 目录结构

```
tas/
├─ README.md
├─ pyproject.toml            # python-docx / lxml / openpyxl
├─ 启动老师端.bat             # 双击启动本地 Web（老师日常入口）
├─ 停止老师端.bat             # 端口 8765 被占用时双击停止
├─ scripts/start_teacher.ps1 # 启动器逻辑：找 Python → 起服务 → 开浏览器
├─ scripts/stop_teacher.ps1  # 停止器：结束占用 8765 的服务进程
├─ scripts/build_portable.ps1# 打包免安装绿色版（内置 Python）
├─ scripts/build_exe.ps1     # PyInstaller 打包脚本（网络可用时生成单目录 exe）
├─ scripts/tas_entry.py      # PyInstaller 入口
├─ dist/tas-web-portable.zip # 构建产物（gitignore）：免安装绿色版，直接发给老师
├─ docs/
│  ├─ REQUIREMENTS.md        # 需求 v0.3
│  ├─ PLAN.md                # 计划/架构
│  ├─ METRICS.md             # 简历指标与评测口径（含简历 bullet 草稿）
│  ├─ TEACHER_GUIDE.md       # 老师端 1 分钟上手
│  └─ MINIPROGRAM.md         # 小程序/接龙方案（含真机网络排查）
├─ profiles/                 # 4 个 profile（默认/文经管/LLM 抽取版/课程示例）
├─ rules/rules.json          # 规则库 38 条
├─ miniprogram/              # 微信小程序源码：文档接龙（发起/提交/一键审核）
├─ src/tas/
│  ├─ cli.py                 # check / fix / refs / batch / preview / web / rules / profiles / spec
│  ├─ checkers/              # page/body/heading/front/float/header_footer/notes/references/structure
│  ├─ fixer.py               # 安全修复（参数化；含 TAB-002 表内文字）
│  ├─ refs/                  # 文献解析/引用一致性/真实性核验
│  ├─ llm/                   # DeepSeek OpenAI 兼容客户端
│  ├─ spec_profile.py        # 规范→规则流水线
│  ├─ batch.py / export.py   # 教师端批量汇总 + xlsx/csv/html（含柱状图/进度条）
│  ├─ feedback.py            # 学生反馈包（逐份 txt + 班级汇总 zip）
│  ├─ batch_fix.py           # 批量安全修复（全部 _fixed.docx + fixed.zip + 修复报告）
│  ├─ thumbnails.py          # 批量首页缩略图（Word 渲染 + HTML 映射 + zip）
│  ├─ batch_refs.py          # 批量文献真实性核验（汇总 + JSON/MD 报告）
│  ├─ roster.py              # 名单核对（缺交/重复/命名不规范）
│  ├─ collections.py         # 文档接龙：发起/提交/一键审核（本地 JSON 存储）
│  ├─ web/app.py             # Web 双入口 + 接龙 API + 手机 H5(/m) + 页面预览
│  ├─ render.py              # docx → PDF（Word COM）→ PNG（WinRT）
│  ├─ rule_params.py / config.py / docx_io.py / roles.py / report.py
├─ tests/                    # 321 tests（collections 9、web 12、feedback 3、batch_fix 12、roster 34、p2 12、thumbnails 27、batch_refs 15）
└─ eval/                     # run_eval.py + report.md/json + 样例
```

## 已知限制（面试可主动说）

1. **PAGE-004 分节继承只报告不自动修复**（结构改动需人工）；**NOTE-001 脚注**只报显式设置且不符的项（继承脚注样式不报，避免误报）；**目录页码一致性**需要渲染引擎，纯 OOXML 无法核对。
2. TAB-002 表内文字已可自动修复（单元格字号/字体/对齐/行距）；但表题位置错（表格上方/下方）只报告不自动移动；表注格式仍只报告。
3. 质量指标来自**合成变异集**（不代表真实分布）；真实语料为 9 份生成论文，未使用学生隐私文档。
4. 查重/AIGC 只预留外部报告接入口（见 [REQUIREMENTS](docs/REQUIREMENTS.md) §3.3）。
5. **页面预览 / 批量缩略图依赖 Windows + Microsoft Word**（Word COM 渲染；无 Word 时显示提示，不影响检查/修复）；批量缩略图是可选开关，逐份启动 Word 较慢（每份约 2–10 秒）。**绿色版请解压到普通目录**（如 `D:\tas老师端` / 桌面），不要放在受限/沙箱目录，否则 Word COM 可能无法启动。
6. **Web 规范管理**需要启动 Web 的进程环境里有 `TAS_LLM_API_KEY`（先 `setx`，再重启 `tas web`）；LLM 调用失败会给出可读错误页，不影响其他功能。
7. **微信小程序上线**需要：注册小程序拿 AppID、后端 HTTPS + 已备案域名并在小程序后台配置 request/uploadFile 合法域名；`wx.uploadFile` 单文件上限 10MB；当前无微信登录/权限模型（任何人都能审核，适合小班）；数据为本地 JSON，上线建议换云存储/数据库。详见 [docs/MINIPROGRAM.md](docs/MINIPROGRAM.md)。
8. **老师端启动器**：项目内 `启动老师端.bat` 需要本机 Python 环境（python-docx / lxml / openpyxl）；没有 Python 的电脑用 `dist\tas-web-portable.zip`（内置运行时）。**单文件 exe** 需 PyInstaller，当前环境 PyPI 受限未构建，但 `scripts\build_exe.ps1` 已保留（网络可用时可直接生成）。
9. **批量修复只做 safe 规则**（页面/正文/标题/前置 + TAB-002 表内文字的格式属性）；涉及文字、编号、题注、参考文献、表题位置的问题不会自动改；修复版是副本，原文件不动。

## 关键约束

1. 只改格式、不改内容；不代写、不润色；**不修改参考文献**。
2. LLM 只输出结构化「规则候选/修改指令」，由确定性代码执行与校验，绝不直接改 OOXML。
3. 先检查后修复；修复带 diff、自动复检、内容 hash 校验、原文件不覆盖。
4. 隐私：真实学生文档只在本机处理，不进入公开仓库；文献核验只发题名/作者；API key 只走环境变量+打码。
5. 规则与代码分离：多学科 = 多个规则 JSON（profile），checker/fixer 已参数化，换规范不改代码。
