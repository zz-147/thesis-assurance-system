"""修复引擎（fixer）回归测试。

覆盖需求 FR-07 / FR-08：
1. 注入字体/页边距/缩进/标题字号等错误 → 修复 → 对应规则复检清零、内容 hash 不变、
   源文件字节不变；
2. 参考文献条目 / 目录条目一律不动（文字与 rPr 完全不变，且出现在 skipped 说明里）；
3. dry-run 不产生任何输出文件，但 actions 非空；
4. 合规文档修复后 actions 为空、复检 0 问题（幂等，不做无意义改写）。

夹具复用 tests/test_check.py 的 build_good_docx（合规文档）。
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
# unittest discover -s tests 时 tests 目录本身在 sys.path 上，保证可复用 test_check
HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from test_check import (  # noqa: E402
    _add_run,
    _set_align,
    _set_spacing,
    add_para,
    build_good_docx,
)

from tas.checkers import run_checks  # noqa: E402
from tas.config import DEFAULT_PROFILE, RuleSet, default_rules_path  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.fixer import SAFE_RULE_IDS, content_hash, fix_document  # noqa: E402
from tas.roles import classify  # noqa: E402

RULES = RuleSet.load()


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------
def find_para(document, text: str):
    """按“去掉全部空白后相等”定位段落（兼容“摘  要 / 目  录”这类制表空格）。"""
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


def set_run_size(paragraph, half_points: int, run_index: int = 0) -> None:
    sz = paragraph.runs[run_index]._r.rPr.find(qn("w:sz"))
    sz.set(qn("w:val"), str(half_points))


def set_east_asia(paragraph, font: str, run_index: int = 0) -> None:
    rfonts = paragraph.runs[run_index]._r.rPr.find(qn("w:rFonts"))
    rfonts.set(qn("w:eastAsia"), font)


def set_line_spacing_raw(paragraph, line: int) -> None:
    spacing = paragraph._p.pPr.find(qn("w:spacing"))
    spacing.set(qn("w:line"), str(line))
    spacing.set(qn("w:lineRule"), "auto")


def set_alignment(paragraph, value: str) -> None:
    paragraph._p.pPr.find(qn("w:jc")).set(qn("w:val"), value)


def remove_indent(paragraph) -> None:
    ppr = paragraph._p.pPr
    for ind in ppr.findall(qn("w:ind")):
        ppr.remove(ind)


def drop_bold(paragraph) -> None:
    rpr = paragraph.runs[0]._r.rPr
    for bold in rpr.findall(qn("w:b")):
        rpr.remove(bold)


def rule_ids(issues) -> set[str]:
    return {issue.rule_id for issue in issues}


def check_file(path: Path, profile: str = DEFAULT_PROFILE, ruleset: RuleSet | None = None):
    model = DocModel.load(path)
    roles = classify(model, profile)
    return run_checks(model, roles, ruleset if ruleset is not None else RULES, profile=profile)


def write_custom_ruleset(tmp: Path, patches: dict[str, dict]) -> RuleSet:
    """复制学校默认规则库，把 patches（rule_id → expected 字段）合并进对应规则的 expected。

    用于模拟教师生成的 profile：改字号/行距/页边距/标题字号后，修复端必须消费新参数。
    """
    data = json.loads(default_rules_path().read_text(encoding="utf-8"))
    for rule in data["rules"]:
        patch = patches.get(rule["id"])
        if patch:
            rule["expected"] = {**(rule.get("expected") or {}), **patch}
    path = tmp / "custom.rules.json"
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return RuleSet.load(path)


def run_half_points(paragraph, run_index: int = 0) -> int:
    """读 run 的 w:sz（半磅，24 = 12pt）。"""
    return int(paragraph.runs[run_index]._r.rPr.find(qn("w:sz")).get(qn("w:val")))


def run_east_asia(paragraph, run_index: int = 0) -> str:
    return paragraph.runs[run_index]._r.rPr.find(qn("w:rFonts")).get(qn("w:eastAsia"))


def jc_value(paragraph) -> str:
    """读段落的 w:jc/@w:val（both / left / center / right）。"""
    jc = paragraph._p.pPr.find(qn("w:jc"))
    return jc.get(qn("w:val")) if jc is not None else ""


def first_line_chars(paragraph) -> int:
    """读 w:ind/@w:firstLineChars（1/100 字符，200 = 2 字符）。"""
    ind = paragraph._p.pPr.find(qn("w:ind"))
    return int(ind.get(qn("w:firstLineChars")))


def action_for(result, rule_id: str):
    for action in result.actions:
        if action.rule_id == rule_id:
            return action
    raise AssertionError(
        f"应产生 {rule_id} 修复动作，实际动作：{[a.rule_id for a in result.actions]}"
    )


def paragraph_xml(path: Path, prefix: str) -> str:
    """取某段落的整段 XML（pPr + rPr + 文字），用于断言“完全没动”。"""
    return find_para_prefix(docx.Document(str(path)), prefix)._p.xml


def skipped_roles(result) -> set[str]:
    return {item.get("role") for item in result.skipped if item.get("kind") == "scope"}


def skipped_rule_ids(result) -> set[str]:
    return {
        item.get("rule_id")
        for item in result.skipped
        if item.get("kind") == "issue" and item.get("rule_id")
    }


def build_table_doc(tmp: Path, name: str = "table.docx") -> Path:
    """合规 fixture + 一张 2x2 数据表：表题合规、单元格文字故意不符合 TAB-002。

    表题放在表格正上方且格式合规（检查器只把「表题 → 表格」紧邻的表当作正文数据表）；
    单元格文字 run 故意用 8pt/黑体、段落左对齐、单倍行距，用来验证 TAB-002 安全修复。
    """
    path = build_good_docx(tmp / name)
    document = docx.Document(str(path))
    anchor = find_para_prefix(document, "检查器优先使用样式与编号判断角色")

    caption = document.add_paragraph()
    _set_align(caption, "center")
    _set_spacing(caption, line=1.5)
    _add_run(caption, "表1 测试表", size=10.5, cn="楷体", ascii_font="Times New Roman")

    table = document.add_table(rows=2, cols=2)
    for row in table.rows:
        for cell in row.cells:
            paragraph = cell.paragraphs[0]
            _set_align(paragraph, "left")
            _set_spacing(paragraph, line=1.0)
            _add_run(
                paragraph, "单元格文字", size=8, cn="黑体", ascii_font="Times New Roman"
            )

    # 把「表题 + 表格」挪到正文段落（anchor）之后：表题在表格正上方
    anchor._p.addnext(table._tbl)
    anchor._p.addnext(caption._p)
    document.save(str(path))
    return path


# ---------------------------------------------------------------------------
class FixEngineTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    # -- 夹具 -------------------------------------------------------------
    def _mutated(self, name: str, mutate) -> Path:
        path = build_good_docx(self.tmp / name)
        document = docx.Document(str(path))
        mutate(document)
        document.save(str(path))
        return path

    def _assert_fixed(self, result, rule_id: str) -> None:
        self.assertIn(rule_id, SAFE_RULE_IDS)
        baseline = {action.rule_id for action in result.actions}
        self.assertIn(
            rule_id, baseline, f"应产生 {rule_id} 修复动作，实际：{result.actions}"
        )
        self.assertNotIn(
            rule_id,
            rule_ids(result.recheck_issues),
            f"修复后 {rule_id} 应清零，剩余：{[i.to_dict() for i in result.recheck_issues]}",
        )
        self.assertTrue(result.content_unchanged, "内容 hash 必须前后一致")
        self.assertEqual(result.content_hash_before, result.content_hash_after)

    # -- 1) 每一类可安全修复的格式问题 -------------------------------------
    def test_fix_body_font_size(self) -> None:
        def mutate(document):
            set_run_size(find_para_prefix(document, "文档格式统一是高校教学管理"), 21)  # 10.5pt

        path = self._mutated("body_font.docx", mutate)
        src_bytes = path.read_bytes()
        self.assertIn("BODY-001", rule_ids(check_file(path)))

        result = fix_document(path)

        self._assert_fixed(result, "BODY-001")
        self.assertIsNotNone(result.out_path)
        self.assertNotEqual(Path(result.out_path).resolve(), path.resolve())
        self.assertTrue(Path(result.out_path).is_file())
        self.assertEqual(path.read_bytes(), src_bytes, "源文件字节不能被修改")
        self.assertGreater(result.baseline_counts["error"], 0)
        # 独立复算一次内容指纹，确认 = 修复后文件的内容 hash
        self.assertEqual(content_hash(DocModel.load(result.out_path)), result.content_hash_after)

    def test_fix_page_geometry(self) -> None:
        def mutate(document):
            section = document.sections[0]
            section.page_width = docx.shared.Cm(29.7)  # 横放：A4 尺寸不对
            section.page_height = docx.shared.Cm(21.0)
            section.top_margin = docx.shared.Cm(3.0)
            section.header_distance = docx.shared.Cm(1.0)

        path = self._mutated("page.docx", mutate)
        baseline_ids = rule_ids(check_file(path))
        self.assertIn("PAGE-001", baseline_ids)
        self.assertIn("PAGE-002", baseline_ids)

        result = fix_document(path)

        self._assert_fixed(result, "PAGE-001")
        self._assert_fixed(result, "PAGE-002")
        remaining = rule_ids(result.recheck_issues)
        self.assertNotIn("PAGE-001", remaining)
        self.assertNotIn("PAGE-002", remaining)

    def test_fix_first_line_indent(self) -> None:
        def mutate(document):
            remove_indent(find_para_prefix(document, "检查器优先使用样式与编号判断角色"))

        path = self._mutated("indent.docx", mutate)
        self.assertIn("BODY-003", rule_ids(check_file(path)))

        result = fix_document(path)

        self._assert_fixed(result, "BODY-003")

    def test_fix_heading_size(self) -> None:
        def mutate(document):
            set_run_size(find_para_prefix(document, "1.1 研究背景"), 32)  # 16pt，二级应为 14pt

        path = self._mutated("heading.docx", mutate)
        self.assertIn("HEAD-002", rule_ids(check_file(path)))

        result = fix_document(path)

        self._assert_fixed(result, "HEAD-002")

    def test_fix_title_and_keywords(self) -> None:
        def mutate(document):
            set_run_size(find_para(document, "测试系统设计与实现"), 24)  # 中文题目 12pt
            set_east_asia(find_para_prefix(document, "关键词："), "宋体")  # 标签应为黑体
            drop_bold(find_para_prefix(document, "Key words"))  # 英文关键词标签应加粗

        path = self._mutated("front.docx", mutate)
        baseline_ids = rule_ids(check_file(path))
        self.assertIn("TITLE-001", baseline_ids)
        self.assertIn("KW-001", baseline_ids)
        self.assertIn("KW-002", baseline_ids)

        result = fix_document(path)

        self._assert_fixed(result, "TITLE-001")
        self._assert_fixed(result, "KW-001")
        self._assert_fixed(result, "KW-002")
        self.assertEqual(result.recheck_issues, [])

    # -- 1b) 表格 TAB-002：表内文字安全修复 ---------------------------------
    def test_fix_table_cell_text(self) -> None:
        path = build_table_doc(self.tmp)
        baseline_tab = [
            issue for issue in check_file(path) if issue.rule_id == "TAB-002"
        ]
        self.assertTrue(
            any("表内文字" in issue.message for issue in baseline_tab),
            f"夹具应触发 TAB-002「表内文字」问题，实际：{baseline_tab}",
        )
        src_bytes = path.read_bytes()
        caption_xml = paragraph_xml(path, "表1")

        result = fix_document(path)

        self._assert_fixed(result, "TAB-002")
        action = action_for(result, "TAB-002")
        self.assertEqual(action.locator, "table/1")
        self.assertEqual(action.after["size_pt"], 10.5)
        self.assertEqual(action.after["east_asia"], "宋体")
        self.assertEqual(action.after["alignment"], "center")
        self.assertEqual(action.after["line_spacing"], 1.5)
        self.assertIn("表格", action.description)

        # 原文件字节不变；表题（不动表题/表注）XML 也不变
        self.assertEqual(path.read_bytes(), src_bytes, "源文件字节不能被修改")
        self.assertEqual(paragraph_xml(result.out_path, "表1"), caption_xml, "表题不能被修改")

        # 输出 docx 里单元格 run 的真实格式：10.5pt / 宋体 / TNR / 居中 / 1.5 倍
        output = docx.Document(str(result.out_path))
        table = output.tables[0]
        seen_cells = []
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                if any(tc is seen for seen in seen_cells):
                    continue
                seen_cells.append(tc)
                paragraph = cell.paragraphs[0]
                run = paragraph.runs[0]
                self.assertAlmostEqual(run.font.size.pt, 10.5, places=2)
                rfonts = run._r.rPr.find(qn("w:rFonts"))
                self.assertEqual(rfonts.get(qn("w:eastAsia")), "宋体")
                self.assertEqual(rfonts.get(qn("w:ascii")), "Times New Roman")
                self.assertEqual(jc_value(paragraph), "center")
                self.assertAlmostEqual(
                    paragraph.paragraph_format.line_spacing, 1.5, places=2
                )
                self.assertEqual(cell.text, "单元格文字", "表格文字不能被修改")
        self.assertEqual(len(seen_cells), 4, "2x2 表格应有 4 个去重后的单元格")

    def test_fix_table_cell_text_is_idempotent(self) -> None:
        path = build_table_doc(self.tmp, "table_twice.docx")
        first = fix_document(path)
        self.assertTrue(
            [action for action in first.actions if action.rule_id == "TAB-002"]
        )

        second = fix_document(first.out_path)

        table_actions = [action for action in second.actions if action.rule_id == "TAB-002"]
        self.assertEqual(table_actions, [], f"第二次修复不应再产生 TAB-002 动作：{table_actions}")
        self.assertEqual(
            [issue for issue in second.recheck_issues if issue.rule_id == "TAB-002"],
            [],
        )
        self.assertTrue(second.content_unchanged)

    # -- 2) 参考文献 / 目录条目绝不修改 ------------------------------------
    def test_reference_entries_not_modified(self) -> None:
        def mutate(document):
            for paragraph in document.paragraphs:
                if paragraph.text.startswith("[1]"):
                    for run in paragraph.runs:
                        run._r.rPr.find(qn("w:sz")).set(qn("w:val"), "18")  # 9pt
                        run._r.rPr.find(qn("w:rFonts")).set(qn("w:eastAsia"), "仿宋")
                    return
            raise AssertionError("未找到参考文献条目")

        path = self._mutated("refs.docx", mutate)
        before_xml = paragraph_xml(path, "[1]")
        result = fix_document(path)
        after_xml = paragraph_xml(result.out_path, "[1]")

        self.assertEqual(before_xml, after_xml, "参考文献条目的 OOXML 必须一字未动")
        self.assertIn("reference_entry", skipped_roles(result), result.skipped)
        reasons = " ".join(item.get("reason", "") for item in result.skipped)
        self.assertIn("参考文献条目", reasons)

    def test_toc_entries_not_modified(self) -> None:
        path = build_good_docx(self.tmp / "toc.docx")
        document = docx.Document(str(path))
        anchor = find_para(document, "目  录")
        entry = add_para(
            document,
            "第1章 绪论\t1",
            size=9,
            cn="仿宋",
            align="left",
            indent_chars=None,
        )
        anchor._p.addnext(entry._p)  # 把目录条目挪到“目  录”标题后面
        document.save(str(path))

        self.assertIn("TOC-001", rule_ids(check_file(path)))
        before_xml = paragraph_xml(path, "第1章 绪论\t")
        result = fix_document(path)
        after_xml = paragraph_xml(result.out_path, "第1章 绪论\t")

        self.assertEqual(before_xml, after_xml, "目录条目必须一字未动")
        self.assertIn("toc_entry", skipped_roles(result), result.skipped)
        self.assertIn(
            "TOC-001",
            skipped_rule_ids(result),
            "目录条目的 TOC-001 问题应进入 skipped（需在 Word 中更新目录域）",
        )
        self.assertNotIn(
            "toc_entry",
            " ".join(action.locator for action in result.actions),
            "不应产生针对目录条目的修复动作",
        )

    # -- 3) dry-run -------------------------------------------------------
    def test_dry_run_writes_nothing(self) -> None:
        def mutate(document):
            set_run_size(find_para_prefix(document, "文档格式统一是高校教学管理"), 21)

        path = self._mutated("dry.docx", mutate)
        explicit_out = self.tmp / "explicit_out.docx"

        result = fix_document(path, explicit_out, dry_run=True)

        self.assertIsNone(result.out_path)
        self.assertFalse(explicit_out.exists(), "dry-run 不应写任何输出文件")
        self.assertFalse(
            any(item.name.endswith("_fixed.docx") for item in self.tmp.iterdir())
        )
        self.assertTrue(result.actions, "dry-run 仍应给出可预览的修复动作")
        self.assertTrue(result.content_unchanged)
        self.assertNotIn("BODY-001", rule_ids(result.recheck_issues))

    # -- 4) 合规文档：零动作 + 幂等 ---------------------------------------
    def test_compliant_document_has_no_actions(self) -> None:
        path = build_good_docx(self.tmp / "good.docx")
        self.assertEqual(check_file(path), [], "前提：夹具文档本身合规")

        result = fix_document(path)

        self.assertEqual(result.actions, [], "合规文档不应产生任何修复动作")
        self.assertEqual(result.recheck_issues, [], "修复后复检应为 0 问题")
        self.assertTrue(result.content_unchanged)
        self.assertTrue(Path(result.out_path).is_file())

    def test_fix_is_idempotent(self) -> None:
        def mutate(document):
            set_run_size(find_para_prefix(document, "文档格式统一是高校教学管理"), 21)
            remove_indent(find_para_prefix(document, "检查器优先使用样式与编号判断角色"))

        path = self._mutated("twice.docx", mutate)
        first = fix_document(path)
        self.assertTrue(first.actions)

        second = fix_document(first.out_path)
        self.assertEqual(second.actions, [], "对已修复文档再修复应为 0 动作")
        self.assertEqual(second.recheck_issues, [])
        self.assertTrue(second.content_unchanged)

    # -- 5) 非 safe 规则进 skipped ---------------------------------------
    def test_unsafe_rule_recorded_in_skipped(self) -> None:
        def mutate(document):
            # 改章号文字 → 触发 HEAD-006（编号类，fixable=false，只报告）
            find_para(document, "第1章 绪论").runs[0].text = "第2章 绪论"

        path = self._mutated("numbering.docx", mutate)
        self.assertIn("HEAD-006", rule_ids(check_file(path)))

        result = fix_document(path, dry_run=True)

        self.assertIn("HEAD-006", skipped_rule_ids(result), result.skipped)
        reason = next(
            item["reason"]
            for item in result.skipped
            if item.get("rule_id") == "HEAD-006"
        )
        self.assertIn("人工", reason)
        self.assertNotIn("HEAD-006", {action.rule_id for action in result.actions})
        self.assertIn("HEAD-006", rule_ids(result.recheck_issues))

    # -- 6) 安全边界：不覆盖源文件 / to_dict 可序列化 ----------------------
    def test_refuses_to_overwrite_source(self) -> None:
        path = build_good_docx(self.tmp / "safe.docx")
        with self.assertRaises(ValueError):
            fix_document(path, out=path)
        with self.assertRaises(ValueError):
            fix_document(str(path), out=str(path))

    def test_to_dict_is_json_serializable(self) -> None:
        def mutate(document):
            set_run_size(find_para_prefix(document, "1.1 研究背景"), 32)

        path = self._mutated("json.docx", mutate)
        result = fix_document(path, dry_run=True)
        payload = result.to_dict()
        text = json.dumps(payload, ensure_ascii=False)  # 不抛异常 = 可序列化
        self.assertIn("action_count", payload)
        self.assertIsNone(payload["out_path"])
        self.assertIn("error", payload["baseline_counts"])
        for action in payload["actions"]:
            self.assertEqual(
                set(action), {"rule_id", "locator", "description", "before", "after"}
            )
        for issue in payload["recheck_issues"]:
            self.assertIn("rule_id", issue)
            self.assertIn("locator", issue)
        self.assertTrue(text)

    def test_humanities_profile_does_not_apply_sci_eng_heading_specs(self) -> None:
        """人文社科 profile 下检查器不检查 HEAD-001/002/003，修复端也不能擅自套用。"""

        def mutate(document):
            set_run_size(find_para_prefix(document, "第1章 绪论"), 24)  # 与人文规范无关

        path = self._mutated("humanities.docx", mutate)
        result = fix_document(path, profile="humanities")

        touched = {action.rule_id for action in result.actions}
        self.assertFalse(
            touched & {"HEAD-001", "HEAD-002", "HEAD-003"},
            f"humanities 下不应改动理工农标题格式，实际动作：{touched}",
        )
        self.assertFalse(
            rule_ids(result.recheck_issues) & {"HEAD-001", "HEAD-002", "HEAD-003"}
        )
        self.assertTrue(result.content_unchanged)


class CustomRulesetParamTest(unittest.TestCase):
    """自定义 profile（ruleset.expected）修复：fixer 必须消费当前规则集的参数。

    背景：fixer 曾把学校默认参数（小四宋体/1.5 倍行距/2cm 页边距/A4…）写死，
    教师用 profile 改了 expected 后修复不生效。这里验证修复端读取 ruleset.expected。
    """

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def test_custom_expected_params_drive_fixes(self) -> None:
        """BODY-001 字号字体 / BODY-002 行距 / PAGE-002 左右边距 / HEAD-001 字号。"""
        path = build_good_docx(self.tmp / "custom_params.docx")
        custom = write_custom_ruleset(
            self.tmp,
            {
                "BODY-001": {
                    "size_pt": 10.5,
                    "east_asia": "仿宋",
                    "ascii": "Times New Roman",
                },
                "BODY-002": {"line_spacing": 1.75},
                "PAGE-002": {"margin_left_cm": 3.0, "margin_right_cm": 3.0},
                "HEAD-001": {"size_pt": 14},
            },
        )

        # 前提：同一文档在默认学期规范下完全合规 —— 差异只能来自 ruleset 参数
        self.assertEqual(check_file(path), [], "默认 ruleset 下夹具文档应合规")
        custom_rules_hit = {issue.rule_id for issue in check_file(path, ruleset=custom)}
        for rule_id in ("BODY-001", "BODY-002", "PAGE-002", "HEAD-001"):
            self.assertIn(
                rule_id, custom_rules_hit, f"自定义参数下 {rule_id} 应报问题"
            )

        src_bytes = path.read_bytes()

        # 默认 ruleset 修复同一文档：0 动作（自定义参数不能泄漏到默认行为）
        default_result = fix_document(path, self.tmp / "default_fixed.docx")
        self.assertEqual(default_result.actions, [], "默认 ruleset 不应产生修复动作")

        # 自定义 ruleset 修复
        result = fix_document(path, self.tmp / "custom_fixed.docx", ruleset=custom)

        covered = {"BODY-001", "BODY-002", "PAGE-002", "HEAD-001"}
        touched = {action.rule_id for action in result.actions}
        self.assertTrue(
            covered <= touched, f"应产生 {sorted(covered)} 修复动作，实际：{sorted(touched)}"
        )
        errors = [issue for issue in result.recheck_issues if issue.severity == "error"]
        self.assertEqual(
            errors,
            [],
            f"自定义 ruleset 复检应为 0 error，实际：{[i.to_dict() for i in errors]}",
        )
        self.assertFalse(
            covered & rule_ids(result.recheck_issues),
            f"修复后自定义规则应清零，剩余：{[i.to_dict() for i in result.recheck_issues]}",
        )
        self.assertTrue(result.content_unchanged, "内容 hash 必须前后一致")
        self.assertEqual(path.read_bytes(), src_bytes, "源文件字节不能被修改")

        # actions 的 after 反映自定义期望值（而不是写死的学校默认值）
        self.assertEqual(action_for(result, "BODY-001").after["size_pt"], 10.5)
        self.assertEqual(action_for(result, "BODY-001").after["east_asia"], "仿宋")
        self.assertEqual(action_for(result, "BODY-002").after["line_spacing"], 1.75)
        self.assertEqual(action_for(result, "PAGE-002").after["left_cm"], 3.0)
        self.assertEqual(action_for(result, "PAGE-002").after["right_cm"], 3.0)
        self.assertEqual(action_for(result, "HEAD-001").after["size_pt"], 14.0)

        # 输出 docx 里的真实格式值
        output = docx.Document(str(result.out_path))
        body = find_para_prefix(output, "文档格式统一是高校教学管理")
        self.assertEqual(run_half_points(body), 21, "正文应为 10.5pt（21 半磅）")
        self.assertEqual(run_east_asia(body), "仿宋")
        self.assertAlmostEqual(body.paragraph_format.line_spacing, 1.75, places=2)

        heading = find_para_prefix(output, "第1章 绪论")
        self.assertEqual(run_half_points(heading), 28, "一级标题应为 14pt（28 半磅）")
        self.assertEqual(run_east_asia(heading), "黑体")

        section = output.sections[0]
        self.assertAlmostEqual(section.left_margin.cm, 3.0, places=1)
        self.assertAlmostEqual(section.right_margin.cm, 3.0, places=1)
        # 自定义参数没改的字段保持默认值
        self.assertAlmostEqual(section.top_margin.cm, 2.54, places=1)
        self.assertAlmostEqual(section.bottom_margin.cm, 2.54, places=1)

        # 自定义 ruleset 下同样幂等：修复结果再修复应 0 动作、0 问题
        second = fix_document(
            result.out_path, self.tmp / "custom_fixed_twice.docx", ruleset=custom
        )
        self.assertEqual(second.actions, [], "已按自定义规范修复的文档不应再有动作")
        self.assertEqual(second.recheck_issues, [])

    def test_custom_page_size_is_applied(self) -> None:
        """PAGE-001 页面尺寸：自定义 profile 改纸张后，修复应落到自定义尺寸。"""
        path = build_good_docx(self.tmp / "custom_page.docx")
        custom = write_custom_ruleset(
            self.tmp,
            {"PAGE-001": {"page_width_cm": 20.0, "page_height_cm": 28.0}},
        )
        self.assertIn("PAGE-001", {issue.rule_id for issue in check_file(path, ruleset=custom)})

        result = fix_document(path, self.tmp / "custom_page_fixed.docx", ruleset=custom)

        self.assertIn("PAGE-001", {action.rule_id for action in result.actions})
        self.assertNotIn("PAGE-001", rule_ids(result.recheck_issues))
        errors = [issue for issue in result.recheck_issues if issue.severity == "error"]
        self.assertEqual(errors, [], f"复检不应有 error：{[i.to_dict() for i in errors]}")
        self.assertTrue(result.content_unchanged)

        section = docx.Document(str(result.out_path)).sections[0]
        self.assertAlmostEqual(section.page_width.cm, 20.0, places=1)
        self.assertAlmostEqual(section.page_height.cm, 28.0, places=1)

    def test_custom_head007_alignment_picked_by_profile(self) -> None:
        """HEAD-007 的 alignment 是 {"sci_eng":…, "humanities":…} 字典，按 profile 取向。"""
        path = build_good_docx(self.tmp / "custom_head007.docx")
        custom = write_custom_ruleset(
            self.tmp,
            {"HEAD-007": {"alignment": {"sci_eng": "left", "humanities": "right"}}},
        )

        sci_result = fix_document(
            path, self.tmp / "custom_head007_sci.docx", ruleset=custom
        )
        self.assertIn("HEAD-007", {action.rule_id for action in sci_result.actions})
        sci_output = docx.Document(str(sci_result.out_path))
        self.assertEqual(jc_value(find_para(sci_output, "结  论")), "left")

        hum_result = fix_document(
            path,
            self.tmp / "custom_head007_hum.docx",
            profile="humanities",
            ruleset=custom,
        )
        self.assertIn("HEAD-007", {action.rule_id for action in hum_result.actions})
        hum_output = docx.Document(str(hum_result.out_path))
        self.assertEqual(jc_value(find_para(hum_output, "结  论")), "right")
        # 人文社科仍不套用理工农标题层级格式
        self.assertFalse(
            {"HEAD-001", "HEAD-002", "HEAD-003"}
            & {action.rule_id for action in hum_result.actions}
        )

    def test_custom_body_indent_and_alignment_params(self) -> None:
        """BODY-003 首行缩进（字符数）/ BODY-004 对齐同样消费自定义 expected。"""
        path = build_good_docx(self.tmp / "custom_body_extra.docx")
        custom = write_custom_ruleset(
            self.tmp,
            {
                "BODY-003": {"first_line_chars": 150},  # 1/100 字符 → 1.5 字符
                "BODY-004": {"alignment": "left"},
            },
        )
        baseline = {issue.rule_id for issue in check_file(path, ruleset=custom)}
        self.assertIn("BODY-003", baseline)
        self.assertIn("BODY-004", baseline)

        result = fix_document(path, self.tmp / "custom_body_extra_fixed.docx", ruleset=custom)

        output = docx.Document(str(result.out_path))
        body = find_para_prefix(output, "文档格式统一是高校教学管理")
        self.assertEqual(first_line_chars(body), 150)
        self.assertEqual(jc_value(body), "left")
        self.assertFalse(
            {"BODY-003", "BODY-004"} & rule_ids(result.recheck_issues),
            f"修复后应清零，剩余：{[i.to_dict() for i in result.recheck_issues]}",
        )
        self.assertTrue(result.content_unchanged)

    def test_custom_tab002_cell_params(self) -> None:
        """TAB-002 表内文字（字号/中文字体/对齐/行距）同样消费自定义 expected。"""
        path = build_table_doc(self.tmp, "custom_table.docx")
        custom = write_custom_ruleset(
            self.tmp,
            {
                "TAB-002": {
                    "cell_size_pt": 9,
                    "cell_east_asia": "楷体",
                    "cell_alignment": "left",
                    "cell_line_spacing": 1.75,
                }
            },
        )
        self.assertIn(
            "TAB-002", {issue.rule_id for issue in check_file(path, ruleset=custom)}
        )

        result = fix_document(
            path, self.tmp / "custom_table_fixed.docx", ruleset=custom
        )

        action = action_for(result, "TAB-002")
        self.assertEqual(action.after["size_pt"], 9.0)
        self.assertEqual(action.after["east_asia"], "楷体")
        self.assertEqual(action.after["alignment"], "left")
        self.assertEqual(action.after["line_spacing"], 1.75)
        self.assertNotIn("TAB-002", rule_ids(result.recheck_issues))

        output = docx.Document(str(result.out_path))
        paragraph = output.tables[0].rows[0].cells[0].paragraphs[0]
        run = paragraph.runs[0]
        self.assertAlmostEqual(run.font.size.pt, 9.0, places=2)
        self.assertEqual(run._r.rPr.find(qn("w:rFonts")).get(qn("w:eastAsia")), "楷体")
        self.assertEqual(jc_value(paragraph), "left")
        self.assertAlmostEqual(paragraph.paragraph_format.line_spacing, 1.75, places=2)

        # 自定义参数下同样幂等
        second = fix_document(
            result.out_path, self.tmp / "custom_table_fixed2.docx", ruleset=custom
        )
        self.assertEqual(
            [item for item in second.actions if item.rule_id == "TAB-002"], []
        )

    def test_tab002_rule_absent_produces_no_table_actions(self) -> None:
        """规则库没有 TAB-002 时不应修改表格（_fix_tables 返回空）。"""
        path = build_table_doc(self.tmp, "no_rule_table.docx")
        data = json.loads(default_rules_path().read_text(encoding="utf-8"))
        data["rules"] = [rule for rule in data["rules"] if rule["id"] != "TAB-002"]
        rules_path = self.tmp / "no_tab002.rules.json"
        rules_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        ruleset = RuleSet.load(rules_path)

        result = fix_document(path, self.tmp / "no_rule_fixed.docx", ruleset=ruleset)

        self.assertNotIn("TAB-002", {action.rule_id for action in result.actions})
        self.assertNotIn("TAB-002", rule_ids(result.recheck_issues))
        self.assertEqual(result.actions, [], "没有 TAB-002 规则时表格不应被改动")
        self.assertTrue(result.content_unchanged)


if __name__ == "__main__":
    unittest.main(verbosity=2)
