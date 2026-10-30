"""P1 / 新规则检查器回归测试（tests/test_p1_checkers.py）。

覆盖范围（每条已注册新规则至少 1 个正例/反例）：

- 图表/公式：FIG-001/002/003、TAB-001/002/003、EQ-001
- 页眉页脚：HF-001/002/003
- 参考文献：REF-002/003
- 结构：TOC-002、STRUCT-001、MISC-001、HEAD-004、HEAD-005、APP-001/002、TRANS-001

测试策略：

1. ``build_p1_docx`` 构造一份"P1 规则全部合规"的完整论文骨架（封面/承诺书/摘要/目录/
   正文/图/三线表/公式/致谢/参考文献/附录/外文资料译文，含页眉页码），
   断言默认规则集与 LLM 抽取规则集下都没有新规则命中（正例）；
2. 每条规则用 python-docx 修改/删除对应格式后断言规则命中（反例）；
3. 现有 ``test_check.build_good_docx``（P0 合规最小文档）在新规则下仍应 0 问题——
   防止 P1 检查器把"没有图/表/页眉的最小文档"误判为违规。
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
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.oxml import OxmlElement, parse_xml  # noqa: E402
from docx.oxml.ns import nsdecls, qn  # noqa: E402

from test_check import build_good_docx  # noqa: E402

from tas.checkers import IMPLEMENTED_RULES, run_checks  # noqa: E402
from tas.config import DEFAULT_PROFILE, RuleSet, default_rules_path  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.roles import classify  # noqa: E402

RULES = RuleSet.load()
LLM_RULES = RuleSet.load(ROOT / "profiles" / "school_llm_2026.rules.json")

P1_RULE_IDS = (
    "FIG-001", "FIG-002", "FIG-003",
    "TAB-001", "TAB-002", "TAB-003",
    "EQ-001",
    "HF-001", "HF-002", "HF-003",
    "REF-002", "REF-003",
    "TOC-002", "STRUCT-001", "MISC-001",
    "HEAD-004", "HEAD-005",
    "APP-001", "APP-002", "TRANS-001",
)


# ---------------------------------------------------------------------------
# fixture 构造工具
# ---------------------------------------------------------------------------
def _new_run(
    paragraph,
    text="",
    *,
    size=12,
    cn="宋体",
    ascii_font="Times New Roman",
    bold=False,
    superscript=False,
):
    run = paragraph.add_run(text)
    rpr = run._r.get_or_add_rPr()
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), ascii_font)
    fonts.set(qn("w:hAnsi"), ascii_font)
    fonts.set(qn("w:cs"), ascii_font)
    if cn:
        fonts.set(qn("w:eastAsia"), cn)
    rpr.append(fonts)
    size_element = OxmlElement("w:sz")
    size_element.set(qn("w:val"), str(int(round(size * 2))))
    rpr.append(size_element)
    if bold:
        rpr.append(OxmlElement("w:b"))
    if superscript:
        run.font.superscript = True
    return run


def _set_align(paragraph, value):
    ppr = paragraph._p.get_or_add_pPr()
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), value)
    ppr.append(jc)


def _set_line_spacing(paragraph, line):
    ppr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:line"), str(int(round(line * 240))))
    spacing.set(qn("w:lineRule"), "auto")
    ppr.append(spacing)


def new_para(
    document,
    *,
    style=None,
    align=None,
    line=None,
    indent_chars=None,
    hanging_chars=None,
    before_lines=None,
    after_lines=None,
):
    paragraph = document.add_paragraph()
    if style:
        paragraph.style = document.styles[style]
    ppr = paragraph._p.get_or_add_pPr()
    if align is not None:
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), align)
        ppr.append(jc)
    if line is not None or before_lines is not None or after_lines is not None:
        spacing = OxmlElement("w:spacing")
        if line is not None:
            spacing.set(qn("w:line"), str(int(round(line * 240))))
            spacing.set(qn("w:lineRule"), "auto")
        if before_lines is not None:
            spacing.set(qn("w:beforeLines"), str(int(round(before_lines * 100))))
        if after_lines is not None:
            spacing.set(qn("w:afterLines"), str(int(round(after_lines * 100))))
        ppr.append(spacing)
    if indent_chars is not None:
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLineChars"), str(int(round(indent_chars * 100))))
        ppr.append(ind)
    if hanging_chars is not None:
        ind = OxmlElement("w:ind")
        ind.set(qn("w:leftChars"), str(int(round(hanging_chars * 100))))
        ind.set(qn("w:hangingChars"), str(int(round(hanging_chars * 100))))
        ppr.append(ind)
    return paragraph


def add_text_para(
    document,
    text,
    *,
    style=None,
    font_size=12,
    cn="宋体",
    ascii_font="Times New Roman",
    bold=False,
    align="both",
    line=1.5,
    indent_chars=2,
    before_lines=None,
    after_lines=None,
):
    paragraph = new_para(
        document,
        style=style,
        align=align,
        line=line,
        indent_chars=indent_chars,
        before_lines=before_lines,
        after_lines=after_lines,
    )
    _new_run(
        paragraph,
        text,
        size=font_size,
        cn=cn,
        ascii_font=ascii_font,
        bold=bold,
    )
    return paragraph


def add_caption(document, number_text, title_text):
    paragraph = new_para(document, align="center", line=1.5)
    _new_run(paragraph, number_text, size=10.5, cn="宋体")
    _new_run(paragraph, title_text, size=10.5, cn="楷体")
    return paragraph


def add_note(document, text):
    return add_text_para(
        document, text, font_size=9, cn="宋体", align="left", line=1.5, indent_chars=2
    )


def add_figure(document):
    paragraph = new_para(document, align="center")
    run = paragraph.add_run()
    run._r.append(OxmlElement("w:drawing"))
    return paragraph


def add_equation(document, number_text):
    paragraph = new_para(document, line=1.5)
    math = parse_xml(
        "<m:oMathPara %s><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>"
        % nsdecls("m")
    )
    paragraph._p.append(math)
    _new_run(paragraph, number_text, size=12)
    return paragraph


def add_reference_entry(document, number_text, body_text):
    paragraph = new_para(document, align="both", line=1.5, hanging_chars=2)
    _new_run(paragraph, number_text, size=12)
    _new_run(paragraph, body_text, size=12)
    return paragraph


def _set_cell_borders(cell, *, bottom=None):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for name, value in (("bottom", bottom),):
        if value is None:
            continue
        element = OxmlElement(f"w:{name}")
        element.set(qn("w:val"), value[0])
        element.set(qn("w:sz"), str(value[1]))
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")
        borders.append(element)
    tc_pr.append(borders)


def _fill_cell(cell, text):
    paragraph = cell.paragraphs[0]
    _set_align(paragraph, "center")
    _set_line_spacing(paragraph, 1.5)
    _new_run(paragraph, text, size=10.5, cn="宋体")


def add_three_line_table(document, header_cells, body_rows):
    table = document.add_table(rows=1 + len(body_rows), cols=len(header_cells))
    table_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for name, val, size in (
        ("top", "single", 12),
        ("left", "none", 0),
        ("bottom", "single", 12),
        ("right", "none", 0),
        ("insideH", "none", 0),
        ("insideV", "none", 0),
    ):
        element = OxmlElement(f"w:{name}")
        element.set(qn("w:val"), val)
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")
        borders.append(element)
    table_pr.append(borders)
    for column, text in enumerate(header_cells):
        cell = table.rows[0].cells[column]
        _set_cell_borders(cell, bottom=("single", 6))
        _fill_cell(cell, text)
    for row_index, row_data in enumerate(body_rows, start=1):
        for column, text in enumerate(row_data):
            _fill_cell(table.rows[row_index].cells[column], text)
    return table


def _add_toc_field(paragraph):
    run = paragraph.add_run()
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = ' TOC \\o "1-3" \\h \\z \\u '
    run._r.append(instruction)


def _set_header(section, text):
    section.header.is_linked_to_previous = False
    paragraph = section.header.paragraphs[0]
    for run in list(paragraph.runs):
        run._r.getparent().remove(run._r)
    _new_run(paragraph, text, size=9, cn="宋体")
    ppr = paragraph._p.get_or_add_pPr()
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    ppr.append(jc)
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    ppr.append(spacing)
    paragraph_border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "double")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    paragraph_border.append(bottom)
    ppr.append(paragraph_border)


def _set_page_footer(section, switch):
    section.footer.is_linked_to_previous = False
    paragraph = section.footer.paragraphs[0]
    for run in list(paragraph.runs):
        run._r.getparent().remove(run._r)

    def _field_run(element):
        run = _new_run(paragraph, "", size=9)
        run._r.append(element)
        return run

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    _field_run(begin)
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = f" PAGE \\* {switch} "
    _field_run(instruction)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    _field_run(separate)
    _new_run(paragraph, "1", size=9)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    _field_run(end)
    ppr = paragraph._p.get_or_add_pPr()
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    ppr.append(jc)
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    ppr.append(spacing)


def _set_page_number_format(section, fmt, start=1):
    sect_pr = section._sectPr
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num)
    pg_num.set(qn("w:fmt"), fmt)
    pg_num.set(qn("w:start"), str(start))


def build_p1_docx(path: Path) -> Path:
    """构造 P1 规则全部合规的完整论文骨架。"""
    document = docx.Document()

    # ---- 第 1 节：封面 + 诚信承诺书（无页眉页码）----
    add_text_para(document, "毕业设计说明书", font_size=22, align="center", indent_chars=None)
    add_text_para(
        document, "毕业论文（设计）诚信承诺书", font_size=16, align="center", indent_chars=None
    )

    # ---- 第 2 节：摘要 / Abstract / 目录（罗马页码 + 页眉）----
    document.add_section(WD_SECTION.NEW_PAGE)
    add_text_para(
        document, "测试系统设计与实现", font_size=16, cn="黑体", align="center",
        indent_chars=None, before_lines=1, after_lines=1,
    )
    add_text_para(
        document, "摘  要", font_size=16, cn="黑体", align="center",
        indent_chars=None, before_lines=1, after_lines=1,
    )
    add_text_para(document, "本文围绕测试系统展开研究，验证格式检查工具的正确性。")
    add_text_para(
        document, "关键词：格式检查；文档规范；自动化", align="left",
        before_lines=1, after_lines=1,
    )
    add_text_para(
        document, "Design and Implementation of a Test System",
        font_size=16, ascii_font="Times New Roman", bold=True, align="center",
        indent_chars=None, before_lines=1, after_lines=1,
    )
    add_text_para(
        document, "Abstract", font_size=16, ascii_font="Times New Roman", align="center",
        indent_chars=None, before_lines=1, after_lines=1,
    )
    add_text_para(
        document, "This paper validates the format checker with a reproducible test document.",
        ascii_font="Times New Roman",
    )
    add_text_para(
        document, "Key words: Format Checking; Document Standard; Automation",
        ascii_font="Times New Roman", align="left", before_lines=1, after_lines=1,
    )
    toc_title = add_text_para(
        document, "目  录", font_size=16, cn="黑体", align="center",
        indent_chars=None, before_lines=1, after_lines=1,
    )
    _add_toc_field(toc_title)
    for entry in (
        "第1章 绪论\t1",
        "1.1 研究背景\t1",
        "1.1.1 问题定义\t1",
        "致  谢\t2",
        "参考文献\t3",
        "附  录\t4",
    ):
        add_text_para(document, entry, indent_chars=None)

    # ---- 第 3 节：正文（阿拉伯页码 + 页眉）----
    document.add_section(WD_SECTION.NEW_PAGE)
    add_text_para(
        document, "第1章 绪论", style="Heading 1", font_size=16, cn="黑体",
        align="center", indent_chars=None, before_lines=1, after_lines=1,
    )
    add_text_para(
        document, "1.1 研究背景", style="Heading 2", font_size=14, cn="黑体",
        align="left", indent_chars=None, before_lines=0.5, after_lines=0.5,
    )
    add_text_para(
        document, "1.1.1 问题定义", style="Heading 3", font_size=12, cn="黑体",
        align="left", indent_chars=None, before_lines=0.5, after_lines=0,
    )
    body = add_text_para(document, "相关研究见文献")
    _new_run(body, "[1]", size=12, superscript=True)
    _new_run(body, "与文献")
    _new_run(body, "[2]", size=12, superscript=True)
    _new_run(body, "，本文在此基础上展开。")

    # 图 3-1（图 + 图题 + 图注释）
    add_figure(document)
    add_caption(document, "图3-1 ", "功能模块图")
    add_note(document, "注：图中数据来自测试环境。")
    # 图 3-2（保证不是"全文仅一图"）
    add_figure(document)
    add_caption(document, "图3-2 ", "登录界面")

    # 表 3-1（表题 + 三线表 + 表注释）
    add_caption(document, "表3-1 ", "测试环境表")
    add_three_line_table(
        document,
        ["分类", "名称", "版本"],
        [["操作系统", "Windows 11", "-"], ["数据库", "MySQL", "8.0"]],
    )
    add_note(document, "注：表中数据为测试环境配置。")

    # 公式（3-1）
    add_equation(document, "（3-1）")

    # 致谢 / 参考文献 / 附录 / 外文资料译文（装订顺序）
    add_text_para(
        document, "致  谢", font_size=16, cn="黑体", align="center",
        indent_chars=None, before_lines=1,
    )
    add_text_para(
        document, "参考文献", font_size=16, cn="黑体", align="center",
        indent_chars=None, before_lines=1,
    )
    add_reference_entry(
        document, "[1] ", "张三. 文档格式自动检查研究[J]. 测试学报, 2024, 1(1): 1-8."
    )
    add_reference_entry(
        document, "[2] ", "李四. 毕业论文格式规范研究[D]. 北京: 测试大学, 2023."
    )
    add_text_para(
        document, "附  录", font_size=16, cn="黑体", align="center", indent_chars=None
    )
    add_text_para(
        document, "测试系统功能清单", font_size=14, cn="黑体", align="center", indent_chars=None
    )
    add_text_para(
        document, "外文资料译文", font_size=16, cn="黑体", align="center", indent_chars=None
    )
    add_text_para(
        document,
        "This paper validates the format checker with a reproducible test document.",
        ascii_font="Times New Roman",
    )
    translation = add_text_para(document, "本文用可复现的测试文档验证格式检查器。")
    translation._p.get_or_add_pPr().insert(0, OxmlElement("w:pageBreakBefore"))

    # ---- 页眉 / 页脚 / 页码格式（全部节创建完后再按节配置）----
    sections = document.sections
    _set_header(sections[1], "商丘学院应用科技学院毕业设计说明书")
    _set_header(sections[2], "商丘学院应用科技学院毕业设计说明书")
    _set_page_footer(sections[1], "ROMAN")
    _set_page_footer(sections[2], "ARABIC")
    _set_page_number_format(sections[1], "upperRoman", 1)
    _set_page_number_format(sections[2], "decimal", 1)

    document.save(str(path))
    return path


def build_head004_docx(path: Path) -> Path:
    document = docx.Document()
    add_text_para(
        document, "第1章 绪论", style="Heading 1", font_size=16, cn="黑体",
        align="center", indent_chars=None,
    )
    add_text_para(
        document, "1.1 研究背景", style="Heading 2", font_size=14, cn="黑体",
        align="left", indent_chars=None, before_lines=0.5, after_lines=0.5,
    )
    add_text_para(
        document, "1.1.1 问题定义", style="Heading 3", font_size=12, cn="黑体",
        align="left", indent_chars=None, before_lines=0.5, after_lines=0,
    )
    add_text_para(
        document, "1.1.1.1 细分问题", font_size=12, cn="宋体", bold=True,
        align="left", indent_chars=None, before_lines=0, after_lines=0,
    )
    document.save(str(path))
    return path


def build_humanities_docx(path: Path) -> Path:
    document = docx.Document()
    add_text_para(
        document, "一、绪论", font_size=16, cn="黑体", align="left",
        indent_chars=2, before_lines=1, after_lines=0,
    )
    add_text_para(
        document, "（一）研究背景", font_size=14, cn="黑体", align="left",
        indent_chars=2, before_lines=0.5, after_lines=0,
    )
    add_text_para(
        document, "1. 研究方法", font_size=12, cn="黑体", align="left",
        indent_chars=2, before_lines=0, after_lines=0,
    )
    document.save(str(path))
    return path


def build_appendix_docx(path: Path) -> Path:
    document = docx.Document()
    add_text_para(
        document, "第1章 绪论", style="Heading 1", font_size=16, cn="黑体",
        align="center", indent_chars=None,
    )
    add_text_para(
        document, "附  录A", font_size=16, cn="黑体", align="center", indent_chars=None
    )
    add_text_para(
        document, "系统功能清单", font_size=14, cn="黑体", align="center", indent_chars=None
    )
    add_text_para(
        document, "附  录B", font_size=16, cn="黑体", align="center", indent_chars=None
    )
    add_text_para(
        document, "测试用例清单", font_size=14, cn="黑体", align="center", indent_chars=None
    )
    document.save(str(path))
    return path


# ---------------------------------------------------------------------------
# 检查工具
# ---------------------------------------------------------------------------
def check_file(path, profile=DEFAULT_PROFILE, ruleset=None):
    model = DocModel.load(path)
    roles = classify(model, profile)
    return run_checks(model, roles, ruleset if ruleset is not None else RULES, profile=profile)


def rule_ids(issues) -> set[str]:
    return {issue.rule_id for issue in issues}


def p1_hits(issues) -> set[str]:
    return rule_ids(issues) & set(P1_RULE_IDS)


# ---------------------------------------------------------------------------
class P1GoodDocumentTest(unittest.TestCase):
    """正例：完整合规文档在所有已实现新规则下 0 问题。"""

    def test_p1_good_document_default_ruleset(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_p1_docx(Path(tmp) / "p1.docx")
            hits = p1_hits(check_file(path))
            self.assertEqual(hits, set(), f"P1 合规文档不应报问题，实际：{sorted(hits)}")

    def test_p1_good_document_llm_ruleset(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_p1_docx(Path(tmp) / "p1.docx")
            hits = p1_hits(check_file(path, ruleset=LLM_RULES))
            self.assertEqual(
                hits, set(), f"LLM 规则集下 P1 合规文档不应报问题，实际：{sorted(hits)}"
            )

    def test_existing_p0_fixture_has_no_p1_issues(self):
        """最小 P0 合规文档（没有图/表/页眉/目录条目）不应被 P1 规则误报。"""
        with tempfile.TemporaryDirectory() as tmp:
            path = build_good_docx(Path(tmp) / "good.docx")
            hits = p1_hits(check_file(path))
            self.assertEqual(hits, set(), f"P0 夹具不应触发 P1 规则，实际：{sorted(hits)}")

    def test_all_p1_rules_registered(self):
        missing = [rule_id for rule_id in P1_RULE_IDS if rule_id not in IMPLEMENTED_RULES]
        self.assertEqual(missing, [], f"以下 P1 规则未注册检查器：{missing}")


# ---------------------------------------------------------------------------
class P1MutationTest(unittest.TestCase):
    """反例：逐条注入违规，断言对应规则命中。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def _mutated(self, mutate, *, profile=DEFAULT_PROFILE, ruleset=None, builder=None):
        path = (builder or build_p1_docx)(self.tmp / "doc.docx")
        document = docx.Document(str(path))
        mutate(document)
        document.save(str(path))
        return p1_hits(check_file(path, profile=profile, ruleset=ruleset))

    # ---- 图表 / 公式 ----
    def test_fig001_caption_font(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("图3-1"):
                    paragraph.runs[1]._r.rPr.find(qn("w:rFonts")).set(
                        qn("w:eastAsia"), "宋体"
                    )
                    return
            raise AssertionError("未找到图 3-1 题注")

        self.assertIn("FIG-001", self._mutated(mutate))

    def test_fig002_number_gap(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("图3-2"):
                    paragraph.runs[0].text = "图3-3 "
                    return
            raise AssertionError("未找到图 3-2 题注")

        self.assertIn("FIG-002", self._mutated(mutate))

    def test_fig003_note_font(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("注：图中数据"):
                    paragraph.runs[0]._r.rPr.find(qn("w:sz")).set(qn("w:val"), "24")
                    return
            raise AssertionError("未找到图注释")

        self.assertIn("FIG-003", self._mutated(mutate, ruleset=LLM_RULES))

    def test_tab001_vertical_border(self):
        def mutate(document):
            for table in document.tables:
                borders = table._tbl.tblPr.find(qn("w:tblBorders"))
                if borders is None:
                    continue
                inside_v = borders.find(qn("w:insideV"))
                inside_v.set(qn("w:val"), "single")
                inside_v.set(qn("w:sz"), "4")
                return
            raise AssertionError("未找到三线表")

        self.assertIn("TAB-001", self._mutated(mutate))

    def test_tab002_cell_font_size(self):
        def mutate(document):
            table = document.tables[0]
            table.rows[0].cells[0].paragraphs[0].runs[0]._r.rPr.find(qn("w:sz")).set(
                qn("w:val"), "24"
            )

        self.assertIn("TAB-002", self._mutated(mutate))

    def test_tab003_note_indent(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("注：表中数据"):
                    ppr = paragraph._p.pPr
                    for ind in ppr.findall(qn("w:ind")):
                        ppr.remove(ind)
                    return
            raise AssertionError("未找到表注释")

        self.assertIn("TAB-003", self._mutated(mutate, ruleset=LLM_RULES))

    def test_eq001_number_format(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    if (run.text or "").strip() == "（3-1）":
                        run.text = "3-1"
                        return
            raise AssertionError("未找到公式编号")

        self.assertIn("EQ-001", self._mutated(mutate))

    # ---- 页眉页脚 ----
    def test_hf001_header_font_size(self):
        def mutate(document):
            header = document.sections[1].header
            for paragraph in header.paragraphs:
                for run in paragraph.runs:
                    if (run.text or "").strip():
                        run._r.rPr.find(qn("w:sz")).set(qn("w:val"), "24")
                        return
            raise AssertionError("未找到页眉文字")

        self.assertIn("HF-001", self._mutated(mutate))

    def test_hf002_header_text(self):
        def mutate(document):
            header = document.sections[1].header
            for paragraph in header.paragraphs:
                for run in paragraph.runs:
                    if (run.text or "").strip():
                        run.text = "某某大学毕业论文"
                        return
            raise AssertionError("未找到页眉文字")

        self.assertIn("HF-002", self._mutated(mutate))

    def test_hf003_page_number_format(self):
        def mutate(document):
            section = document.sections[1]
            pg_num = section._sectPr.find(qn("w:pgNumType"))
            pg_num.set(qn("w:fmt"), "decimal")

        self.assertIn("HF-003", self._mutated(mutate))

    # ---- 参考文献 ----
    def test_ref002_entry_font(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("[1] "):
                    paragraph.runs[-1]._r.rPr.find(qn("w:rFonts")).set(
                        qn("w:eastAsia"), "仿宋"
                    )
                    return
            raise AssertionError("未找到参考文献条目")

        self.assertIn("REF-002", self._mutated(mutate))

    def test_ref003_citation_out_of_range(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("相关研究见文献"):
                    _new_run(paragraph, "[9]", size=12, superscript=True)
                    return
            raise AssertionError("未找到正文引用段")

        self.assertIn("REF-003", self._mutated(mutate))

    # ---- 结构 / 全局 ----
    def test_toc002_missing_toc_field(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").replace(" ", "").replace("\u3000", "") == "目录":
                    for instruction in paragraph._p.iter(qn("w:instrText")):
                        instruction.text = ""
                    return
            raise AssertionError("未找到目录标题")

        self.assertIn("TOC-002", self._mutated(mutate))

    def test_struct001_missing_appendix(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                text = (paragraph.text or "").replace(" ", "").replace("\u3000", "")
                if text == "附录":
                    paragraph.runs[0].text = "附加说明"
                    return
            raise AssertionError("未找到附录标题")

        self.assertIn("STRUCT-001", self._mutated(mutate))

    def test_misc001_short_year(self):
        def mutate(document):
            paragraph = document.add_paragraph()
            _new_run(paragraph, "该项目始于99年，属于二十世纪的技术。", size=12)

        self.assertIn("MISC-001", self._mutated(mutate))

    def test_head004_level4_size(self):
        def check(path):
            return p1_hits(check_file(path))

        with tempfile.TemporaryDirectory() as tmp:
            good = build_head004_docx(Path(tmp) / "h4.docx")
            self.assertNotIn("HEAD-004", check(good), "合规四级标题不应报 HEAD-004")

            document = docx.Document(str(good))
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("1.1.1.1"):
                    paragraph.runs[0]._r.rPr.find(qn("w:sz")).set(qn("w:val"), "28")
                    break
            document.save(str(good))
            self.assertIn("HEAD-004", check(good))

    def test_head005_humanities_heading_size(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_humanities_docx(Path(tmp) / "hum.docx")
            self.assertNotIn(
                "HEAD-005",
                p1_hits(check_file(path, profile="humanities")),
                "合规文经管标题不应报 HEAD-005",
            )
            document = docx.Document(str(path))
            document.paragraphs[0].runs[0]._r.rPr.find(qn("w:sz")).set(
                qn("w:val"), "28"
            )  # 14pt，应为 16pt
            document.save(str(path))
            self.assertIn("HEAD-005", p1_hits(check_file(path, profile="humanities")))

    # ---- 新增规则（LLM 规则集中才有）----
    def test_app001_appendix_title_size(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                text = (paragraph.text or "").replace(" ", "").replace("\u3000", "")
                if text == "附录":
                    paragraph.runs[0]._r.rPr.find(qn("w:sz")).set(qn("w:val"), "28")
                    return
            raise AssertionError("未找到附录标题")

        self.assertIn("APP-001", self._mutated(mutate, ruleset=LLM_RULES))

    def test_app002_appendix_numbering(self):
        with tempfile.TemporaryDirectory() as tmp:
            good = build_appendix_docx(Path(tmp) / "app.docx")
            self.assertNotIn(
                "APP-002",
                p1_hits(check_file(good, ruleset=LLM_RULES)),
                "附录A/附录B 不应报 APP-002",
            )
            document = docx.Document(str(good))
            for paragraph in document.paragraphs:
                text = (paragraph.text or "").replace(" ", "").replace("\u3000", "")
                if text == "附录B":
                    paragraph.runs[0].text = "附  录"
                    break
            document.save(str(good))
            self.assertIn("APP-002", p1_hits(check_file(good, ruleset=LLM_RULES)))

    def test_trans001_translation_new_page(self):
        def mutate(document):
            for paragraph in document.paragraphs:
                if (paragraph.text or "").startswith("本文用可复现的测试文档验证"):
                    ppr = paragraph._p.pPr
                    break_before = ppr.find(qn("w:pageBreakBefore"))
                    if break_before is not None:
                        ppr.remove(break_before)
                    return
            raise AssertionError("未找到中文翻译段")

        self.assertIn("TRANS-001", self._mutated(mutate, ruleset=LLM_RULES))


# ---------------------------------------------------------------------------
class P1RulesetParameterTest(unittest.TestCase):
    """参数化：改 ruleset 的 expected 后，新检查器必须跟着变。"""

    def test_custom_tab_and_hf_params_are_consumed(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_p1_docx(Path(tmp) / "p1.docx")
            data = json.loads(default_rules_path().read_text(encoding="utf-8"))
            for rule in data["rules"]:
                if rule["id"] == "TAB-002":
                    rule["expected"]["cell_size_pt"] = 12
                elif rule["id"] == "HF-001":
                    rule["expected"]["size_pt"] = 10.5
            custom_path = Path(tmp) / "custom.rules.json"
            custom_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            custom = RuleSet.load(custom_path)

            hits = p1_hits(check_file(path, ruleset=custom))
            self.assertIn("TAB-002", hits, "自定义表内字号后 TAB-002 应命中")
            self.assertIn("HF-001", hits, "自定义页眉字号后 HF-001 应命中")


if __name__ == "__main__":
    unittest.main(verbosity=2)
