"""回归测试：好文档应零问题；注入错误后对应规则必须命中。"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import docx  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm  # noqa: E402

from tas.checkers import run_checks  # noqa: E402
from tas.config import DEFAULT_PROFILE, RuleSet  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.roles import classify  # noqa: E402

RULES = RuleSet.load()
ALIGN = {
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
}


# ---------------------------------------------------------------------------
# fixture 构造工具
# ---------------------------------------------------------------------------
def _add_run(paragraph, text, *, size=12, cn="宋体", ascii_font="Times New Roman", bold=False):
    run = paragraph.add_run(text)
    rpr = run._r.get_or_add_rPr()
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), ascii_font)
    fonts.set(qn("w:hAnsi"), ascii_font)
    fonts.set(qn("w:cs"), ascii_font)
    if cn:
        fonts.set(qn("w:eastAsia"), cn)
    rpr.append(fonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rpr.append(sz)
    if bold:
        rpr.append(OxmlElement("w:b"))
    run._r.insert(0, rpr)
    return run


def _pPr(paragraph):
    return paragraph._p.get_or_add_pPr()


def _set_align(paragraph, value):
    ppr = _pPr(paragraph)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), value)


def _set_spacing(paragraph, *, line=1.5, before_lines=None, after_lines=None):
    ppr = _pPr(paragraph)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(int(line * 240)))
    spacing.set(qn("w:lineRule"), "auto")
    if before_lines is not None:
        spacing.set(qn("w:beforeLines"), str(int(before_lines * 100)))
    if after_lines is not None:
        spacing.set(qn("w:afterLines"), str(int(after_lines * 100)))


def _set_indent(paragraph, chars=None):
    ppr = _pPr(paragraph)
    old = ppr.find(qn("w:ind"))
    if old is not None:
        ppr.remove(old)
    if chars is None:
        return
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ppr.append(ind)


def add_para(
    document,
    text,
    *,
    style=None,
    size=12,
    cn="宋体",
    ascii_font="Times New Roman",
    bold=False,
    align="both",
    line=1.5,
    indent_chars=2,
    before_lines=None,
    after_lines=None,
):
    paragraph = document.add_paragraph()
    if style:
        paragraph.style = document.styles[style]
    _set_align(paragraph, align)
    _set_spacing(paragraph, line=line, before_lines=before_lines, after_lines=after_lines)
    _set_indent(paragraph, indent_chars)
    _add_run(paragraph, text, size=size, cn=cn, ascii_font=ascii_font, bold=bold)
    return paragraph


def build_good_docx(path: Path) -> Path:
    document = docx.Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.header_distance = Cm(2.0)
    section.footer_distance = Cm(2.0)

    pg_mar = section._sectPr.find(qn("w:pgMar"))
    pg_mar.set(qn("w:gutter"), "567")  # 1cm

    doc_grid = section._sectPr.find(qn("w:docGrid"))
    if doc_grid is None:
        doc_grid = OxmlElement("w:docGrid")
        section._sectPr.append(doc_grid)
    doc_grid.set(qn("w:type"), "linesAndChars")
    doc_grid.set(qn("w:linePitch"), "450")

    # ---- 前置区 ----
    add_para(document, "测试系统设计与实现", size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "摘  要", size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "本文围绕测试系统展开研究，验证格式检查工具的正确性，并给出可复现的测试用例。")

    keywords_cn = document.add_paragraph()
    _set_align(keywords_cn, "left")
    _set_spacing(keywords_cn, before_lines=1.0, after_lines=1.0)
    _set_indent(keywords_cn, 2)
    _add_run(keywords_cn, "关键词：", size=12, cn="黑体")
    _add_run(keywords_cn, "格式检查；文档规范；自动化", size=12, cn="宋体")

    add_para(document, "Design and Implementation of a Test System", size=16,
             ascii_font="Times New Roman", bold=True, align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "Abstract", size=16, ascii_font="Times New Roman", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "This paper validates the format checker with a reproducible test document.",
             ascii_font="Times New Roman")

    keywords_en = document.add_paragraph()
    _set_align(keywords_en, "left")
    _set_spacing(keywords_en, before_lines=1.0, after_lines=1.0)
    _set_indent(keywords_en, 2)
    _add_run(keywords_en, "Key words: ", size=12, ascii_font="Times New Roman", bold=True)
    _add_run(keywords_en, "Format Checking; Document Standard; Automation", size=12,
             ascii_font="Times New Roman")

    add_para(document, "目  录", size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)

    # ---- 正文 ----
    add_para(document, "第1章 绪论", style="Heading 1", size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "1.1 研究背景", style="Heading 2", size=14, cn="黑体", align="left",
             before_lines=0.5, after_lines=0.5, indent_chars=None)
    add_para(document, "1.1.1 问题定义", style="Heading 3", size=12, cn="黑体", align="left",
             before_lines=0.5, after_lines=0.0, indent_chars=None)
    add_para(document, "文档格式统一是高校教学管理中的高频问题，本文用确定性规则实现可复现的检查。")
    add_para(document, "检查器优先使用样式与编号判断角色，识别不出的段落交由人工或模型兜底。")

    add_para(document, "结  论", size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "致  谢", size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "参考文献", size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, "[1] 张三. 文档格式自动检查研究[J]. 测试学报, 2024, 1(1): 1-8.",
             indent_chars=None)

    document.save(str(path))
    return path


def check_file(path: Path):
    model = DocModel.load(path)
    roles = classify(model, DEFAULT_PROFILE)
    return run_checks(model, roles, RULES, profile=DEFAULT_PROFILE)


# ---------------------------------------------------------------------------
class GoodDocumentTest(unittest.TestCase):
    def test_good_document_has_no_issues(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_good_docx(Path(tmp) / "good.docx")
            issues = check_file(path)
            details = "\n".join(
                f"{issue.rule_id} {issue.rule_id and issue.locator} {issue.actual}" for issue in issues
            )
            self.assertEqual(issues, [], f"合规文档不应报问题，实际：\n{details}")

    def test_roles_recognize_structure(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_good_docx(Path(tmp) / "good.docx")
            model = DocModel.load(path)
            roles = classify(model, DEFAULT_PROFILE)
            role_values = list(roles.by_index.values())
            for expected in (
                "title_cn", "title_en", "abstract_label_cn", "abstract_label_en",
                "keywords_cn", "keywords_en", "toc_title", "heading_1", "heading_2",
                "heading_3", "refs_title", "thanks_title",
            ):
                self.assertIn(expected, role_values, f"未识别出角色 {expected}")
            self.assertNotIn("cover", role_values)


class MutationDetectionTest(unittest.TestCase):
    def _mutated(self, mutate):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = build_good_docx(Path(tmp.name) / "doc.docx")
        document = docx.Document(str(path))
        mutate(document)
        document.save(str(path))
        return check_file(path)

    def test_wrong_body_font_size(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if paragraph.text.startswith("文档格式统一是高校教学管理"):
                    run = paragraph.runs[0]
                    sz = run._r.rPr.find(qn("w:sz"))
                    sz.set(qn("w:val"), "21")  # 10.5pt（五号）
                    return
            raise AssertionError("未找到目标正文段落")

        issues = self._mutated(mutate)
        self.assertTrue(any(issue.rule_id == "BODY-001" for issue in issues), issues)

    def test_wrong_margin(self):
        def mutate(document):
            document.sections[0].left_margin = Cm(3.0)

        issues = self._mutated(mutate)
        self.assertTrue(any(issue.rule_id == "PAGE-002" for issue in issues), issues)

    def test_missing_first_line_indent(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if paragraph.text.startswith("检查器优先使用样式"):
                    ppr = paragraph._p.pPr
                    for ind in ppr.findall(qn("w:ind")):
                        ppr.remove(ind)
                    return
            raise AssertionError("未找到目标正文段落")

        issues = self._mutated(mutate)
        self.assertTrue(any(issue.rule_id == "BODY-003" for issue in issues), issues)

    def test_wrong_heading_size(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if paragraph.text.startswith("1.1 研究背景"):
                    sz = paragraph.runs[0]._r.rPr.find(qn("w:sz"))
                    sz.set(qn("w:val"), "32")  # 16pt，二级标题应为 14pt
                    return
            raise AssertionError("未找到目标标题段落")

        issues = self._mutated(mutate)
        self.assertTrue(any(issue.rule_id == "HEAD-002" for issue in issues), issues)


if __name__ == "__main__":
    unittest.main(verbosity=2)
