"""前置区参数化测试：checker / fixer 必须消费当前 ruleset 的 expected。

背景：TITLE / ABS / KW / TOC / REF 的字体字号等期望值曾写死在代码里，教师用 profile
改了 expected 后，检查/修复仍按学校默认值执行。本测试覆盖：

1. checker：自定义 ruleset（复制 rules/rules.json 后改参数）能检出差异，且报告里的
   expected 反映自定义参数；默认 ruleset 下同一夹具文档 0 前置问题（回归）。
2. fixer：自定义 ruleset 修复后，用同一 ruleset 复检前置规则清零，输出 docx 的真实
   run 字号/字体/缩进/对齐符合自定义期望；默认 ruleset 修复 0 动作（参数不泄漏）；
   内容 hash 前后一致、源文件字节不变、修复幂等。
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from test_check import add_para, build_good_docx  # noqa: E402

from tas.checkers import run_checks  # noqa: E402
from tas.config import DEFAULT_PROFILE, RuleSet, default_rules_path  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.fixer import fix_document  # noqa: E402
from tas.roles import classify  # noqa: E402

#: 本任务参数化的全部前置区规则
FRONT_RULES = (
    "TITLE-001", "TITLE-002",
    "ABS-001", "ABS-002", "ABS-003", "ABS-004",
    "KW-001", "KW-002", "TOC-001", "REF-001",
)

#: task-5 要求的最小自定义参数集（复制默认规则库后覆盖 expected）
CUSTOM_PATCHES: dict[str, dict] = {
    "TITLE-001": {"size_pt": 14, "east_asia": "楷体"},
    "ABS-001": {"size_pt": 18},
    "KW-001": {"label_size_pt": 10.5},
    "TOC-001": {"title_size_pt": 14},
    "REF-001": {"size_pt": 14},
}

#: 覆盖其余前置规则的参数（TITLE-002 / 摘要正文 / 英文关键词）
EXTRA_PATCHES: dict[str, dict] = {
    "TITLE-002": {"size_pt": 14, "ascii": "Arial"},
    "ABS-002": {"size_pt": 10.5, "first_line_chars": 150, "alignment": "left"},
    "ABS-003": {"size_pt": 14},
    "ABS-004": {"size_pt": 10.5, "alignment": "left"},
    "KW-002": {"label_size_pt": 10.5, "body_ascii": "Arial"},
}


# ---------------------------------------------------------------------------
# 夹具 / 规则集工具
# ---------------------------------------------------------------------------
def write_custom_ruleset(
    tmp: Path, patches: dict[str, dict], name: str = "front_custom.rules.json"
) -> RuleSet:
    """复制默认规则库，把 patches（rule_id → expected 字段）合并进对应规则的 expected。"""
    data = json.loads(default_rules_path().read_text(encoding="utf-8"))
    for rule in data["rules"]:
        patch = patches.get(rule["id"])
        if patch:
            rule["expected"] = {**(rule.get("expected") or {}), **patch}
    path = tmp / name
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return RuleSet.load(path)


def check_file(
    path: Path, ruleset: RuleSet | None = None, profile: str = DEFAULT_PROFILE
):
    model = DocModel.load(path)
    roles = classify(model, profile)
    return run_checks(model, roles, ruleset if ruleset is not None else RuleSet.load(), profile=profile)


def front_issues(issues) -> list:
    return [issue for issue in issues if issue.rule_id in FRONT_RULES]


def rule_ids(issues) -> set[str]:
    return {issue.rule_id for issue in issues}


def find_para(document, text: str):
    target = "".join((text or "").split())
    for paragraph in document.paragraphs:
        if "".join((paragraph.text or "").split()) == target:
            return paragraph
    raise AssertionError(f"未找到段落：{text!r}")


def find_para_prefix(document, prefix: str):
    for paragraph in document.paragraphs:
        if (paragraph.text or "").lstrip().startswith(prefix):
            return paragraph
    raise AssertionError(f"未找到以 {prefix!r} 开头的段落")


def run_half_points(paragraph, run_index: int = 0) -> int:
    """读 run 的 w:sz（半磅，24 = 12pt）。"""
    return int(paragraph.runs[run_index]._r.rPr.find(qn("w:sz")).get(qn("w:val")))


def run_east_asia(paragraph, run_index: int = 0):
    rfonts = paragraph.runs[run_index]._r.rPr.find(qn("w:rFonts"))
    return rfonts.get(qn("w:eastAsia")) if rfonts is not None else None


def run_ascii(paragraph, run_index: int = 0):
    rfonts = paragraph.runs[run_index]._r.rPr.find(qn("w:rFonts"))
    return rfonts.get(qn("w:ascii")) if rfonts is not None else None


def run_bold(paragraph, run_index: int = 0) -> bool:
    return bool(paragraph.runs[run_index].bold)


def jc_value(paragraph) -> str:
    jc = paragraph._p.pPr.find(qn("w:jc"))
    return jc.get(qn("w:val")) if jc is not None else ""


def first_line_chars(paragraph) -> int:
    ind = paragraph._p.pPr.find(qn("w:ind"))
    return int(ind.get(qn("w:firstLineChars")))


def action_for(result, rule_id: str):
    for action in result.actions:
        if action.rule_id == rule_id:
            return action
    raise AssertionError(
        f"应产生 {rule_id} 修复动作，实际动作：{[a.rule_id for a in result.actions]}"
    )


# ---------------------------------------------------------------------------
# checker：自定义参数生效 / 默认参数回归
# ---------------------------------------------------------------------------
class FrontCheckerParamTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def test_default_ruleset_has_no_front_issues(self) -> None:
        path = build_good_docx(self.tmp / "good.docx")
        issues = check_file(path)
        self.assertEqual(issues, [], "默认规则集下合规夹具不应报任何问题")
        self.assertEqual(front_issues(issues), [])

    def test_custom_front_params_are_detected(self) -> None:
        path = build_good_docx(self.tmp / "custom_front.docx")
        custom = write_custom_ruleset(self.tmp, CUSTOM_PATCHES)
        # 前提：同一文档在默认规则集下完全合规 —— 差异只能来自 ruleset 参数
        self.assertEqual(check_file(path), [])

        issues = front_issues(check_file(path, ruleset=custom))
        ids = rule_ids(issues)
        for rule_id in CUSTOM_PATCHES:
            self.assertIn(rule_id, ids, f"自定义参数下 {rule_id} 应报问题，实际：{sorted(ids)}")

        # 报告里的 expected 必须反映自定义参数（不是写死的学校默认值）
        title_issue = next(issue for issue in issues if issue.rule_id == "TITLE-001")
        self.assertEqual(title_issue.expected["size_pt"], 14)
        self.assertEqual(title_issue.expected["east_asia"], "楷体")
        abs_issue = next(issue for issue in issues if issue.rule_id == "ABS-001")
        self.assertEqual(abs_issue.expected["size_pt"], 18)
        kw_issue = next(issue for issue in issues if issue.rule_id == "KW-001")
        self.assertEqual(kw_issue.expected["size_pt"], 10.5)
        toc_issue = next(issue for issue in issues if issue.rule_id == "TOC-001")
        self.assertEqual(toc_issue.expected["size_pt"], 14)
        ref_issue = next(issue for issue in issues if issue.rule_id == "REF-001")
        self.assertEqual(ref_issue.expected["size_pt"], 14)

    def test_default_expected_still_uses_school_defaults(self) -> None:
        """默认 ruleset 缺自定义参数时，docx 被改坏后报告仍应显示学校默认期望值。"""
        path = build_good_docx(self.tmp / "mutated_title.docx")
        document = docx.Document(str(path))
        title = find_para(document, "测试系统设计与实现")
        title.runs[0]._r.rPr.find(qn("w:sz")).set(qn("w:val"), "28")  # 14pt
        document.save(str(path))

        issues = front_issues(check_file(path))
        title_issue = next(issue for issue in issues if issue.rule_id == "TITLE-001")
        self.assertEqual(title_issue.expected["size_pt"], 16)
        self.assertEqual(title_issue.expected["east_asia"], "黑体")

    def test_extra_front_params_are_detected(self) -> None:
        path = build_good_docx(self.tmp / "extra_front.docx")
        custom = write_custom_ruleset(self.tmp, EXTRA_PATCHES, name="front_extra.rules.json")
        ids = rule_ids(front_issues(check_file(path, ruleset=custom)))
        for rule_id in EXTRA_PATCHES:
            self.assertIn(rule_id, ids, f"自定义参数下 {rule_id} 应报问题，实际：{sorted(ids)}")

        abs_cn_issue = next(
            issue for issue in front_issues(check_file(path, ruleset=custom))
            if issue.rule_id == "ABS-002"
        )
        self.assertEqual(abs_cn_issue.expected["size_pt"], 10.5)
        self.assertAlmostEqual(abs_cn_issue.expected["first_line_chars"], 1.5, places=2)
        self.assertEqual(abs_cn_issue.expected["alignment"], "left")

    def test_custom_keyword_text_params_are_detected(self) -> None:
        """KW-001 的 count_range / separator / trailing_punctuation 文本参数同样生效。"""
        path = build_good_docx(self.tmp / "kw_text.docx")
        custom = write_custom_ruleset(
            self.tmp,
            {
                "KW-001": {
                    "count_range": [4, 5],
                    "separator": "，",
                    "trailing_punctuation": True,
                }
            },
            name="kw_text.rules.json",
        )
        issues = front_issues(check_file(path, ruleset=custom))
        kw_issue = next(issue for issue in issues if issue.rule_id == "KW-001")
        self.assertEqual(kw_issue.expected["count_range"], [4, 5])
        self.assertEqual(kw_issue.expected["separator"], "，")
        self.assertTrue(kw_issue.expected["trailing_punctuation"])
        self.assertIn("关键词数量=1", kw_issue.actual)

    def test_custom_toc_entry_params_are_detected(self) -> None:
        """TOC-001 的目录条目字体/字号同样来自 expected（fixture 本身无目录条目）。"""
        path = build_good_docx(self.tmp / "toc.docx")
        document = docx.Document(str(path))
        anchor = find_para(document, "目  录")
        entry = add_para(
            document,
            "第1章 绪论\t1",
            size=12,
            cn="宋体",
            align="left",
            indent_chars=None,
        )
        anchor._p.addnext(entry._p)
        document.save(str(path))

        # 默认规范下（条目小四宋体）合规
        self.assertEqual(front_issues(check_file(path)), [])

        custom = write_custom_ruleset(
            self.tmp,
            {"TOC-001": {"entry_size_pt": 10.5, "others_east_asia": "仿宋"}},
            name="toc_entry.rules.json",
        )
        issues = front_issues(check_file(path, ruleset=custom))
        toc_issue = next(issue for issue in issues if issue.rule_id == "TOC-001")
        self.assertEqual(toc_issue.expected["size_pt"], 10.5)
        self.assertEqual(toc_issue.expected["east_asia"], "仿宋")


# ---------------------------------------------------------------------------
# fixer：自定义参数驱动修复 / 默认参数不泄漏
# ---------------------------------------------------------------------------
class FrontFixerParamTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def test_default_ruleset_fix_has_no_actions(self) -> None:
        path = build_good_docx(self.tmp / "good.docx")
        src_bytes = path.read_bytes()
        result = fix_document(path, self.tmp / "default_fixed.docx")
        self.assertEqual(result.actions, [], "默认规则集修复合规文档不应产生动作")
        self.assertEqual(result.recheck_issues, [])
        self.assertTrue(result.content_unchanged)
        self.assertEqual(result.content_hash_before, result.content_hash_after)
        self.assertEqual(path.read_bytes(), src_bytes, "源文件字节不能被修改")

    def test_custom_front_params_drive_fixes(self) -> None:
        path = build_good_docx(self.tmp / "custom_front.docx")
        custom = write_custom_ruleset(self.tmp, CUSTOM_PATCHES)
        src_bytes = path.read_bytes()

        # 默认 ruleset 修复同一文档：0 动作（自定义参数不能泄漏到默认行为）
        default_result = fix_document(path, self.tmp / "default_fixed.docx")
        self.assertEqual(default_result.actions, [])

        result = fix_document(path, self.tmp / "custom_fixed.docx", ruleset=custom)
        touched = {action.rule_id for action in result.actions}
        self.assertTrue(
            set(CUSTOM_PATCHES) <= touched,
            f"应产生 {sorted(CUSTOM_PATCHES)} 修复动作，实际：{sorted(touched)}",
        )
        self.assertEqual(
            result.recheck_issues,
            [],
            f"自定义 ruleset 复检应为 0 问题，剩余：{[i.to_dict() for i in result.recheck_issues]}",
        )
        self.assertTrue(result.content_unchanged, "内容 hash 必须前后一致")
        self.assertEqual(result.content_hash_before, result.content_hash_after)
        self.assertEqual(path.read_bytes(), src_bytes, "源文件字节不能被修改")

        # actions 的 after 反映自定义期望值（而不是写死的学校默认值）
        self.assertEqual(action_for(result, "TITLE-001").after["size_pt"], 14)
        self.assertEqual(action_for(result, "TITLE-001").after["east_asia"], "楷体")
        self.assertEqual(action_for(result, "ABS-001").after["size_pt"], 18)
        self.assertEqual(action_for(result, "KW-001").after["label_size_pt"], 10.5)
        self.assertEqual(action_for(result, "TOC-001").after["size_pt"], 14)
        self.assertEqual(action_for(result, "REF-001").after["size_pt"], 14)

        # 输出 docx 里的真实格式值
        output = docx.Document(str(result.out_path))
        self.assertEqual(run_half_points(find_para(output, "测试系统设计与实现")), 28)  # 14pt
        self.assertEqual(run_east_asia(find_para(output, "测试系统设计与实现")), "楷体")
        self.assertEqual(run_half_points(find_para(output, "摘  要")), 36)  # 18pt
        self.assertEqual(run_half_points(find_para(output, "目  录")), 28)  # 14pt
        self.assertEqual(run_half_points(find_para(output, "参考文献")), 28)  # 14pt
        keywords_cn = find_para_prefix(output, "关键词：")
        self.assertEqual(run_half_points(keywords_cn, 0), 21)  # 标签 10.5pt
        self.assertEqual(run_half_points(keywords_cn, 1), 24)  # 正文仍 12pt

        # 自定义 ruleset 下同样幂等：修复结果再修复应 0 动作、0 问题
        second = fix_document(
            result.out_path, self.tmp / "custom_fixed_twice.docx", ruleset=custom
        )
        self.assertEqual(second.actions, [], "已按自定义规范修复的文档不应再有动作")
        self.assertEqual(second.recheck_issues, [])

    def test_extra_front_params_drive_fixes(self) -> None:
        path = build_good_docx(self.tmp / "extra_front.docx")
        custom = write_custom_ruleset(
            self.tmp, EXTRA_PATCHES, name="front_extra.rules.json"
        )
        result = fix_document(path, self.tmp / "extra_fixed.docx", ruleset=custom)

        touched = {action.rule_id for action in result.actions}
        self.assertTrue(
            set(EXTRA_PATCHES) <= touched,
            f"应产生 {sorted(EXTRA_PATCHES)} 修复动作，实际：{sorted(touched)}",
        )
        self.assertEqual(
            result.recheck_issues,
            [],
            f"自定义 ruleset 复检应为 0 问题，剩余：{[i.to_dict() for i in result.recheck_issues]}",
        )
        self.assertTrue(result.content_unchanged)

        output = docx.Document(str(result.out_path))
        title_en = find_para_prefix(output, "Design and Implementation")
        self.assertEqual(run_half_points(title_en), 28)  # 14pt
        self.assertEqual(run_ascii(title_en), "Arial")
        self.assertTrue(run_bold(title_en), "英文题目仍应加粗")

        abs_en_label = find_para(output, "Abstract")
        self.assertEqual(run_half_points(abs_en_label), 28)  # 14pt

        abs_cn_body = find_para_prefix(output, "本文围绕测试系统展开研究")
        self.assertEqual(run_half_points(abs_cn_body), 21)  # 10.5pt
        self.assertEqual(first_line_chars(abs_cn_body), 150)  # 1.5 字符
        self.assertEqual(jc_value(abs_cn_body), "left")

        abs_en_body = find_para_prefix(output, "This paper validates")
        self.assertEqual(run_half_points(abs_en_body), 21)
        self.assertEqual(jc_value(abs_en_body), "left")

        keywords_en = find_para_prefix(output, "Key words:")
        self.assertEqual(run_half_points(keywords_en, 0), 21)  # 标签 10.5pt
        self.assertTrue(run_bold(keywords_en), "英文关键词标签仍应加粗")
        self.assertEqual(run_ascii(keywords_en, 1), "Arial")  # 正文西文改 Arial


if __name__ == "__main__":
    unittest.main(verbosity=2)
