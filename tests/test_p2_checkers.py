"""P2 检查器回归测试：PAGE-004 分节页眉页脚继承 / NOTE-001 脚注格式。

覆盖：

- PAGE-004（``checkers.header_footer.check_page004``）：
  1) 2 节 + 首节有页眉文字 → 报「封面/承诺书不应显示页眉或页码」；
  2) 2 节 + 首节干净、第 2 节有页眉 → 不报；
  3) 单节最小文档（无页眉页脚/页码）→ []（防误报）；
  4) ruleset 无 PAGE-004 规则 → []；
  5) 首节页脚有 PAGE 域 → 报（封面不得有页码）；
  6) 自定义 expected（cover_no_header_footer=false / sections_min）参数化生效。
- NOTE-001（``checkers.notes.check_notes``）：
  1) 合规 footnotes.xml（9pt 宋体 1.0 倍）→ []；
  2) 错误脚注（12pt 黑体 1.5 倍）→ 至少 1 条 warning，locator ``footnote/1``；
  3) 无 ``word/footnotes.xml`` → []；
  4) 基础规则库（无 NOTE-001）→ []；
  5) 只有分隔符/空脚注 → []；多条脚注只报不合规那条。

全部离线：footnotes.xml 由 zipfile 直接注入 python-docx 生成的 docx，
不调用 Word/WPS，不联网。
"""

from __future__ import annotations

import dataclasses
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import docx  # noqa: E402
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from tas.checkers.header_footer import check_page004  # noqa: E402
from tas.checkers.notes import check_notes  # noqa: E402
from tas.config import RuleSet, resolve_profile  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.roles import classify  # noqa: E402

RULES = RuleSet.load()  # 基础规则库：有 PAGE-004，无 NOTE-001
_PROFILE, LLM_RULES = resolve_profile("school_llm_2026")  # 有 NOTE-001

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------------------
# docx fixture 工具
# ---------------------------------------------------------------------------
def build_docx(path: Path, texts=("封面内容",)) -> Path:
    document = docx.Document()
    for text in texts:
        document.add_paragraph(text)
    document.save(str(path))
    return path


def add_section(docx_path: Path, texts=("正文内容",)) -> Path:
    """给已有 docx 追加一个分节符 + 正文段落（共 2 节）。"""
    document = docx.Document(str(docx_path))
    document.add_section(WD_SECTION.NEW_PAGE)
    for text in texts:
        document.add_paragraph(text)
    document.save(str(docx_path))
    return docx_path


def set_header_text(section, text: str) -> None:
    section.header.is_linked_to_previous = False
    paragraph = section.header.paragraphs[0]
    for run in list(paragraph.runs):
        run._r.getparent().remove(run._r)
    paragraph.add_run(text)


def set_footer_page_field(section) -> None:
    """在页脚放一个 PAGE 域（instrText 即被 docx_io 识别为页码域）。"""
    section.footer.is_linked_to_previous = False
    paragraph = section.footer.paragraphs[0]
    run = paragraph.add_run()
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE \\* ARABIC "
    run._r.append(instruction)


def check_page004_file(docx_path: Path, ruleset: RuleSet | None = None):
    model = DocModel.load(docx_path)
    roles = classify(model)
    return check_page004(model, roles, ruleset if ruleset is not None else RULES)


# ---------------------------------------------------------------------------
# footnotes.xml 构造 / 注入
# ---------------------------------------------------------------------------
def _footnote_xml(
    footnote_id: int,
    text: str,
    *,
    size_pt: float = 9.0,
    east_asia: str = "宋体",
    line_spacing: float = 1.0,
    with_spacing: bool = True,
    with_rpr: bool = True,
) -> str:
    half_points = int(round(size_pt * 2))
    rpr = (
        f'<w:rPr><w:rFonts w:eastAsia="{east_asia}"/>'
        f'<w:sz w:val="{half_points}"/></w:rPr>'
        if with_rpr
        else ""
    )
    line = int(round(line_spacing * 240))
    spacing = (
        f'<w:spacing w:line="{line}" w:lineRule="auto"/>' if with_spacing else ""
    )
    ppr = f"<w:pPr>{spacing}</w:pPr>" if spacing else ""
    return (
        f'<w:footnote w:id="{footnote_id}">'
        f"<w:p>{ppr}<w:r>{rpr}"
        f'<w:t xml:space="preserve">{text}</w:t>'
        f"</w:r></w:p>"
        f"</w:footnote>"
    )


def build_footnotes_xml(*footnotes: str) -> bytes:
    separator = (
        '<w:footnote w:type="separator" w:id="-1">'
        "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
    )
    continuation = (
        '<w:footnote w:type="continuationSeparator" w:id="0">'
        "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{W_NS}">'
        + separator
        + continuation
        + "".join(footnotes)
        + "</w:footnotes>"
    ).encode("utf-8")


def inject_footnotes(docx_path: Path, footnotes_xml: bytes) -> Path:
    """把 word/footnotes.xml 注入 docx 包（只读测试夹具，不改业务代码）。"""
    temporary = docx_path.with_name(docx_path.name + ".injecting")
    try:
        with zipfile.ZipFile(docx_path) as source:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    target.writestr(item, source.read(item.filename))
                target.writestr("word/footnotes.xml", footnotes_xml)
        temporary.replace(docx_path)
    finally:
        if temporary.exists():
            temporary.unlink()
    return docx_path


# ---------------------------------------------------------------------------
# PAGE-004
# ---------------------------------------------------------------------------
class Page004Test(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def test_first_section_header_is_reported(self):
        path = build_docx(self.tmp / "two_sections.docx")
        add_section(path)
        document = docx.Document(str(path))
        set_header_text(document.sections[0], "测试页眉")
        document.save(str(path))

        issues = check_page004_file(path)
        self.assertTrue(issues, "首节有页眉时应报 PAGE-004")
        cover_issues = [issue for issue in issues if "封面/承诺书" in issue.message]
        self.assertEqual(len(cover_issues), 1, issues)
        issue = cover_issues[0]
        self.assertEqual(issue.rule_id, "PAGE-004")
        self.assertEqual(issue.severity, "warning")
        self.assertEqual(issue.locator, "section/1")
        self.assertIn("header", issue.actual)
        self.assertEqual(issue.actual["header"], "测试页眉")
        self.assertFalse(issue.actual["page_field"])

    def test_single_section_minimal_document_is_empty(self):
        path = build_docx(self.tmp / "single.docx")
        self.assertEqual(check_page004_file(path), [])

    def test_first_section_clean_second_section_header_is_ok(self):
        path = build_docx(self.tmp / "front_body.docx")
        add_section(path)
        document = docx.Document(str(path))
        set_header_text(document.sections[1], "正文页眉")
        document.save(str(path))

        self.assertEqual(check_page004_file(path), [])

    def test_ruleset_without_page004_returns_empty(self):
        path = build_docx(self.tmp / "no_rule.docx")
        add_section(path)
        document = docx.Document(str(path))
        set_header_text(document.sections[0], "测试页眉")
        document.save(str(path))

        ruleset_without = RuleSet(
            {key: value for key, value in RULES.rules.items() if key != "PAGE-004"},
            RULES.meta,
        )
        self.assertEqual(check_page004_file(path, ruleset=ruleset_without), [])

    def test_first_section_footer_page_field_is_reported(self):
        path = build_docx(self.tmp / "cover_page_number.docx")
        add_section(path)
        document = docx.Document(str(path))
        set_footer_page_field(document.sections[0])
        document.save(str(path))

        issues = check_page004_file(path)
        self.assertTrue(issues, "首节页脚有 PAGE 域时应报 PAGE-004")
        cover_issues = [issue for issue in issues if "封面/承诺书" in issue.message]
        self.assertEqual(len(cover_issues), 1, issues)
        self.assertTrue(cover_issues[0].actual["page_field"])

    def test_cover_check_can_be_disabled_by_expected(self):
        path = build_docx(self.tmp / "param.docx")
        add_section(path)
        document = docx.Document(str(path))
        set_header_text(document.sections[0], "测试页眉")
        document.save(str(path))

        base_rule = RULES.rules["PAGE-004"]
        custom_rule = dataclasses.replace(
            base_rule,
            expected={**base_rule.expected, "cover_no_header_footer": False},
        )
        custom_ruleset = RuleSet({**RULES.rules, "PAGE-004": custom_rule}, RULES.meta)
        self.assertEqual(check_page004_file(path, ruleset=custom_ruleset), [])


# ---------------------------------------------------------------------------
# NOTE-001
# ---------------------------------------------------------------------------
class Note001Test(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def _check(self, footnotes_xml: bytes | None, ruleset: RuleSet | None = None):
        path = build_docx(self.tmp / "doc.docx", ("正文段落",))
        model = DocModel.load(path)
        roles = classify(model)
        if footnotes_xml is not None:
            inject_footnotes(path, footnotes_xml)
        return check_notes(
            model, roles, ruleset if ruleset is not None else LLM_RULES
        )

    def test_compliant_footnote_is_empty(self):
        xml = build_footnotes_xml(
            _footnote_xml(1, "这是合规脚注。", size_pt=9, east_asia="宋体", line_spacing=1.0)
        )
        self.assertEqual(self._check(xml), [])

    def test_wrong_footnote_is_reported(self):
        xml = build_footnotes_xml(
            _footnote_xml(1, "这是错误脚注。", size_pt=12, east_asia="黑体", line_spacing=1.5)
        )
        issues = self._check(xml)
        self.assertGreaterEqual(len(issues), 1)
        issue = issues[0]
        self.assertEqual(issue.rule_id, "NOTE-001")
        self.assertEqual(issue.severity, "warning")
        self.assertEqual(issue.locator, "footnote/1")
        self.assertEqual(issue.message, "脚注格式不符合规范")
        self.assertAlmostEqual(issue.actual["size_pt"], 12.0)
        self.assertEqual(issue.actual["east_asia"], "黑体")
        self.assertAlmostEqual(issue.actual["line_spacing"], 1.5)
        self.assertIn("size_pt", issue.expected)

    def test_missing_footnotes_part_is_empty(self):
        self.assertEqual(self._check(None), [])

    def test_base_ruleset_without_note001_returns_empty(self):
        xml = build_footnotes_xml(
            _footnote_xml(1, "这是错误脚注。", size_pt=12, east_asia="黑体", line_spacing=1.5)
        )
        self.assertEqual(self._check(xml, ruleset=RULES), [])

    def test_separators_and_empty_footnotes_are_skipped(self):
        xml = build_footnotes_xml(
            _footnote_xml(2, "", size_pt=12, east_asia="黑体", line_spacing=1.5),
            _footnote_xml(3, "   "),
        )
        self.assertEqual(self._check(xml), [])

    def test_only_bad_footnote_reported(self):
        xml = build_footnotes_xml(
            _footnote_xml(1, "合规脚注。", size_pt=9, east_asia="宋体", line_spacing=1.0),
            _footnote_xml(2, "错误脚注。", size_pt=12, east_asia="黑体", line_spacing=1.5),
        )
        issues = self._check(xml)
        self.assertEqual([issue.locator for issue in issues], ["footnote/2"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
