# profiles —— 多学科 / 多课程规范

一个 **profile** = 一套规则库文件（`rules.json`）+ 一个学科体系（`sci_eng` / `humanities`）+ 来源说明。

## 为什么这样设计

- 用户需求：**每门学科的规范可能不同**，甚至同一学院不同专业模板不同。
- 规则库与代码分离：换规范 = 换 / 新增一个 JSON，不改代码。
- 教师端上传规范文档后，走 `tas spec parse`（LLM 抽取规则候选 → 教师确认）→ `tas spec apply`（合并生成并注册新 profile），供班级学生选择。

## 当前内置

| id | 学科体系 | 规则文件 | 说明 |
|---|---|---|---|
| `school_default` | `sci_eng` | `../rules/rules.json` | 学院统一规范（默认） |
| `school_default_humanities` | `humanities` | `../rules/rules.json` | 同一规范，文经管标题体系 |
| `school_llm_2026` | `sci_eng` | `school_llm_2026.rules.json` | DeepSeek 从学院规范抽取（38 keep + 6 new，待教师审） |
| `course_demo_other` | `sci_eng` | `course_demo_other.rules.json` | 演示：另一门课程规范（5 条 modify：页边距/正文/标题字号） |

## 自动生成一个新学科 / 课程 profile（教师端）

```powershell
$env:PYTHONPATH = "D:\zdsh\work\tas\src"
# 1) 规范文档（.docx/.txt/.md）→ LLM 规则候选 + diff 报告
& $PY -m tas spec parse "规范文档.docx" --out-dir eval/spec_parse
# 2) 教师审阅 eval/spec_parse/diff.md / candidates.json（可删除不认可的候选）
# 3) 合并生成并注册 profile
& $PY -m tas spec apply "eval/spec_parse/candidates.json" --id course_xxx --name "XX课程格式要求"
# 4) 使用
& $PY -m tas check "论文.docx" --ruleset course_xxx
& $PY -m tas fix "论文.docx" --ruleset course_xxx
```

- 由 LLM 新增（new）的规则需要绑定 checker 才会参与检查；`expected` 参数在 checker/fixer 中已全部参数化（页面/正文/标题/前置），改参数立即生效。
- 分享方式（无账号阶段）：直接分发 `profiles/<id>.rules.json` + 在 `index.json` 注册。

## 手工新增一个学科 / 课程规范

1. 复制一份规则库：`cp rules/rules.json profiles/<id>.rules.json`，按该学科规范改 `expected` 参数、增删规则与 `source` 出处。
2. 在 `profiles/index.json` 的 `profiles` 数组里加一条：
   ```json
   {
     "id": "course_xxx_2025",
     "name": "XX 课程 2025 论文格式要求",
     "discipline": "sci_eng",
     "rules": "../profiles/course_xxx_2025.rules.json",
     "source": "教师提供的规范文档路径",
     "notes": ""
   }
   ```
3. 使用：`tas check 论文.docx --ruleset course_xxx_2025`；学生端修复同理。
