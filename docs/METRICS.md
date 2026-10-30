# 简历可用的指标与评测说明

> 生成时间：2026-10-26 ｜ 项目：`D:\zdsh\work\tas`（tas — 论文格式统一 / 真实性核验工具）
> 所有数字均由仓库内脚本可复现（见文末命令）；合成数据的口径与局限见第 8 节。

## 1. 项目规模

| 指标 | 数值 |
|---|---|
| 规则库规则数 | 38 条（P0 23 / P1 15） |
| 检查器已实现规则 | **38 条（P0 23/23，P1 15/15）**；无未实现规则 |
| LLM 新抽规则（`school_llm_2026` profile） | 6 条新增全部实现（含 NOTE-001 脚注）→ 该 profile **44/44** |
| 命令行子命令 | 9 个：`check / fix / refs / batch / preview / web / rules / profiles / spec` |
| 多学科 profile | 4 个（学院默认 / 文经管 / LLM 抽取版 / 课程自定义示例） |
| 源码规模 | `src/` 14,562 行 ｜ `tests/` 6,228 行 ｜ `eval/` 1,254 行 ｜ 微信小程序 2,384 行（15 文件） |
| 自动化测试 | **321 个**（离线 317 通过 + 4 个 live/渲染默认 skip；开环境变量后 321/321 通过） |

测试构成：check 6 ｜ fix 22 ｜ refs 31 ｜ llm 45 ｜ spec_profile 12 ｜ rule_params 3 ｜ front_params 9 ｜ p1_checkers 25 ｜ p2_checkers 12 ｜ batch 22 ｜ eval_metrics 14 ｜ web 12 ｜ render 4 ｜ collections 9 ｜ feedback 3 ｜ batch_fix 12 ｜ roster 34 ｜ packaging 4 ｜ thumbnails 27 ｜ batch_refs 15。

## 2. 核心质量指标（合成变异集，`eval/run_eval.py`）

| 指标 | 数值 | 口径 |
|---|---|---|
| 变异注入点 | 116 处 / 23 类（每类 ≥5 处） | 字号/字体/行距/缩进/对齐/页边距/文档网格/标题格式/摘要/关键词/目录/参考文献标题等 |
| **检出率（recall）** | **100.0%**（116/116） | 规则命中且定位到注入段落 |
| 定位准确率 | 100.0% | 命中定位到注入段落的比例 |
| **误报率（FP rate）** | **0.0%**（0/201 段） | 3 份合规 fixture 的 201 个非空段被误报数 |
| **修复后 P0 通过率** | **100.0%** | 12 份注入可安全修复问题的文档，修复后复检注入规则 error 清零 |
| **内容零改动率** | **100.0%** | 修复前后全部文字（段落/表格/页眉页脚）sha256 一致 |
| 单份检查耗时 | 0.094 s/份 | 全规则检查（10 页级合成文档） |
| 单份修复耗时 | 0.253 s/份 | 含读写 docx + 复检 |
| **批量吞吐** | **9.80 份/秒**（102.0 ms/份） | 12 份顺序检查 |

## 3. 真实语料批量（`out/` 9 份成品，只读；终版 38 条规则）

| 指标 | 数值 |
|---|---|
| 文档数 | 9 份（检查成功 9，失败 0） |
| 问题总数 | 错误 111 ｜ 警告 193 ｜ 提示 8 |
| 每份问题 | 12–14 error / 21–26 warning（同一批生成文档的系统性偏差） |
| 共性问题 Top | BODY-004 对齐 176 次 ｜ TAB-002 表题/表内文字 69 次 ｜ STRUCT-001 装订缺件 16 次 |
| 单份修复验收（岳希华论文） | 14 error + 23 warning → **36 项安全修复**（含 TAB-002 表内文字 10 项）→ 复检剩 **0 error + 2 warning**（仅 STRUCT-001 缺附录/外文翻译）；内容 hash 一致、参考文献与目录条目 XML 逐字节未变 |
| 批量导出 | xlsx（3 Sheet）/ csv / html 均生成并通过 openpyxl 重开校验 |

## 4. 教师端「规范 → 规则」LLM 实测（DeepSeek `deepseek-chat`）

| 规范 | 耗时 | 逐条评估 | 结果 |
|---|---|---|---|
| 学院规范（9455 字） | 11 s | **38/38** | keep 38 ｜ modify 0 ｜ **new 6** ｜ unmapped 23 |
| 另一门课程规范（418 字） | 4.5 s | **38/38** | keep 33 ｜ **modify 5**（页边距/正文字体字号/行距/一二级标题字号）｜ unmapped 1 |

- 产出：`candidates.json` + `diff.md`（教师审阅）→ `tas spec apply` 生成并注册新 profile。
- LLM 客户端带 JSON 模式、重试、key 打码；key 只走环境变量，全仓 0 处泄漏。

## 5. 多规范参数化效果（同一篇论文，两套规范）

| 规范 | 检查结果 | 一键修复后 |
|---|---|---|
| 学院默认 `school_default` | 错误 5（旧版 P0）→ 终版 14 error（含 P1） | 26 项安全修复 → P0 清零，剩 P1 非自动项 |
| 课程自定义 `course_demo_other`（10.5pt 仿宋 / 1.75 倍 / 左右 3cm / 一级标题 14pt） | **错误 299**（BODY-001 132、BODY-002 132、HEAD/PAGE 等） | 复检 0；实测正文仿宋 10.5、行距 1.75、页边距 2.5+3.0cm、标题 14pt；内容 hash 一致 |

→ 证明 profile 参数真正进入 checker/fixer（PAGE/BODY/HEAD/前置全部参数化；前置 10 条规则 10/10 生效）。

## 6. 真实性核验（教师端，只读）

- 文献解析：GB/T 7714（期刊/图书/会议/学位论文/专利/标准/报纸/电子资源），12 条真实论文文献全部解析。
- 元数据核验（Crossref → OpenAlex → arXiv）：示例 3 条 → 1 `verified`（ResNet，DOI 10.1109/cvpr.2016.90）、1 `likely`（题名过泛 + 年份不一致提示）、1 `not_found`（中文期刊未收录，标注"需人工核实"）；所有结论附证据 URL；`not_found ≠ 造假`。
- 引用一致性（离线）：越界/缺号/未引用/首次引用顺序/上标/文献表编号连续性。
- 真实网络 live 测试通过（`TAS_LIVE_NETWORK=1`）。

## 7. 可复现命令

```powershell
$PY = "D:\dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
$env:PYTHONPATH = "D:\zdsh\work\tas\src"
Set-Location D:\zdsh\work\tas

& $PY -m unittest discover -s tests -v                 # 191 tests（离线）
& $PY eval\run_eval.py --out eval\report.md            # 质量指标（第 2 节数字）
& $PY -m tas batch "D:\zdsh\work\thesis-format\out" --out-dir eval\batch_real   # 第 3 节数字
& $PY -m tas spec parse "eval\demo_spec_other_discipline.txt" --out-dir eval\spec_parse_other   # 第 4 节
& $PY -m tas web --port 8765                           # 浏览器打开 http://127.0.0.1:8765
```

## 8. 口径与局限（写简历/面试时必须能解释）

1. 质量指标（第 2 节）来自**合成变异集**：对合规文档注入已知格式错误，用于衡量"规则能否稳定检出/修复"，不代表真实学生文档的全部分布；真实语料结果见第 3 节。
2. 真实语料为 9 份生成论文（同一生成器产物），系统性偏差集中；**未使用真实学生隐私文档做评测**。
3. 评测只统计当前已实现规则（38 条）；新增规则后重跑 `eval/run_eval.py` 自动更新。
4. 规则边界：PAGE-004 只报告不自动修复（结构改动需人工）；NOTE-001 只报显式设置且不符的脚注（继承样式不报）；TAB-002 单元格可自动修复，但表题位置/表注仍只报告；目录页码一致性需渲染引擎；查重/AIGC 需外部服务。
5. 文献核验的 `not_found` 只表示未在已查索引中找到，中文期刊覆盖有限，不能作为造假证据。

## 9. 简历 bullet 草稿（数字均可在仓库复现，可按岗位裁剪）

- 独立开发「论文格式规范化与真实性核验工具」（Python / python-docx / OOXML / DeepSeek）：规则库 38 条、检查器覆盖 **38/38（P0 23/23，P1 15/15）**；合成变异集 116 处注入检出率 **100%**、误报率 **0%**；安全修复后 P0 通过率 **100%**、内容零改动（hash **100%**）；单份检查 0.09s、批量吞吐 **9.8 份/秒**。
- 教师端「规范文档 → 规则」：DeepSeek 从 9455 字学院规范 **38/38** 逐条评估（11s），抽出 6 条人工遗漏规则、23 条 unmapped；支持多学科 profile 参数化——同一论文换规范错误数 5→299 并一键修复到目标格式；Web 双入口 + 班级批量汇总导出 xlsx。
- 文献真实性：GB/T 7714 解析 + Crossref/OpenAlex/arXiv 元数据核验（verified/likely/not_found + 证据 URL，`not_found ≠ 造假`）+ 引用一致性离线检查；**321 个自动化测试**，key 全程环境变量+打码。
- **规则全覆盖 + 表格修复**：PAGE-004 分节/封面页眉页码、NOTE-001 脚注（直读 `word/footnotes.xml`）检查器补齐 → 基础规则 **38/38**、LLM profile **44/44**；TAB-002 表内文字自动修复（真实论文 10 项，复检 0 error）；新增 16 个测试（12 检查器 + 4 修复器）。
- **批量缩略图 + 文献核验（可选）**：`thumbnails.py` 批量首页 PNG（Word COM；**27 tests**，含 fake renderer/zip/映射/live skip）；`batch_refs.py` 批量文献真实性核验（verified/likely/not_found/error + JSON/MD，**15 tests**，fake fetcher 离线覆盖）；CLI `--preview/--refs` + Web 教师端勾选；真实论文 12 条文献抽 2 条核验、报告含"人工核实"提示。
- **老师端增强（名单/修复/免安装）**：名单核对（缺交/重复/命名不规范，xlsx/csv/txt 名单；34 tests）、批量安全修复（全部 `_fixed.docx` + fixed.zip + 修复报告；12 tests）、HTML 报告条形图/通过率可视化、免安装绿色版（内置 Python 3.12 运行时，50.8 MB 目录 / 18.5 MB zip；实测 CLI 检查 + Web 服务，无需安装/联网）；Web 教师端一键勾选。
- **老师收文件工作流**：双击启动器自动起服务并打开浏览器；教师端「选择整个文件夹」批量收 `.docx` → 批量检查 → 导出学生反馈包（经典 PowerShell 兼容 UTF-8 BOM；反馈包 = 班级汇总 + 每份文档一个可直接转发的 txt）。9 份真实语料实测生成 10 条目反馈 zip；3 个自动化测试。
- **文档接龙（单界面，不分教师/学生端）**：面向班级"微信群接龙收文档"场景——发起接龙 → 同学从聊天记录选 `.docx` 提交 → 提交即检查 → 任何人一键审核（通过率/共性问题 Top/逐份报告）。实现后端 API（本地 JSON 存储）+ 手机 H5 + 原生微信小程序（15 文件）；9 个自动化测试，并完成真实后端全链路联调（创建→提交→审核→报告，`pass_rate`/`top_rules`/中文错误均正确）；**微信开发者工具 2.02.0 模拟器实测**首页/发起/详情/审核/报告全流程（真实论文 2 份审核 0% 通过、错误 620；合规文档 100% 通过），`chooseMessageFile` 原生选文件待真机验证。
