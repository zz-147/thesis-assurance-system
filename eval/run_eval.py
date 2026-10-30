#!/usr/bin/env python
"""tas 评测脚本：合成变异集上的检出率 / 误报率 / 修复效果 / 耗时。

用法（在项目根目录）::

    python eval/run_eval.py --out eval/report.md
    python eval/run_eval.py --out eval/report.md --docs 12

评测内容
--------
1. **变异检出**：用 python-docx 现场合成「合规文档」（不读取任何真实学生文档），
   再对已知位置注入各类格式错误（字号 / 字体 / 行距 / 缩进 / 对齐 / 页边距 /
   纸张 / 标题字号 / 标题段前距 / 题目 / 摘要 / 关键词 / 目录 / 参考文献标题 …），
   跑 ``tas.checkers.run_checks``，统计：
   - 检出率（recall）：注入点中「对应 P0 规则命中且定位到注入段落」的比例；
   - 定位准确率：规则命中里定位到注入位置的比例；
   - 误报率（FP rate）：合规文档中被已评测规则误报的段落比例。
2. **修复评测**：用 :func:`tas.fixer.fix_document` 修复混合变异文档，统计
   修复后 P0 通过率、内容 hash 不变比例、单份检查/修复耗时；
3. **批量吞吐**：对 ≥10 份变异文档跑 :func:`tas.batch.check_files`，统计吞吐。

诚实性约束
----------
- 全部数据现场合成（不联网、不写真实学生信息、不改仓库数据）；
- **只统计当前检查器已实现的规则**：已实现规则集合动态读取
  ``tas.checkers.IMPLEMENTED_RULES``，规则数量增加时报告自动适配，不写死；
- 未纳入本次变异集的已实现规则不会进入检出率分母（避免把“没评测”算成“漏检”），
  报告里注明；合规 fixture 未覆盖的新规则命中单独计数，不计入误报率；
- 输出 ``eval/report.md``（指标/数值/说明表格）与同名 ``.json``（机器可读）。
"""

from __future__ import annotations

import argparse
import json
import math
import platform
import re
import shutil
import sys
import tempfile
import time
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Optional, Sequence

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import docx  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm  # noqa: E402

from tas import __version__ as TAS_VERSION  # noqa: E402
from tas.batch import check_files  # noqa: E402
from tas.checkers import IMPLEMENTED_RULES, run_checks  # noqa: E402
from tas.config import DEFAULT_PROFILE, RuleSet, resolve_profile  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.fixer import SAFE_RULE_IDS, fix_document  # noqa: E402
from tas.report import Issue  # noqa: E402
from tas.roles import classify  # noqa: E402

__all__ = [
    "MUTATIONS",
    "build_compliant_docx",
    "run_checks_for",
    "aggregate_detection",
    "aggregate_false_positives",
    "locator_paragraph_index",
    "run_evaluation",
    "render_markdown",
    "write_reports",
    "main",
]

# 合规 fixture 的结构参数（章节数 × 每章正文段数；正文段数需 ≥5 供多注入点类别使用）
CHAPTERS = 5
BODY_PER_CHAPTER = 2
#: 每个变异类别的最少注入数（--injections 可覆盖，契约要求 ≥5）
DEFAULT_INJECTIONS = 5
#: 修复 / 吞吐评测的默认文档数（契约要求 ≥10）
DEFAULT_DOCS = 12
#: 文件名里用的安全字符
_SAFE_RE = re.compile(r"[^0-9A-Za-z_.-]+")


# ===========================================================================
# 一、合成合规文档（python-docx；不读取任何真实文档）
# ===========================================================================
def _norm(text: str) -> str:
    return "".join((text or "").split())


def _add_run(
    paragraph,
    text: str,
    *,
    size: float = 12,
    cn: str = "宋体",
    ascii_font: str = "Times New Roman",
    bold: bool = False,
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
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rpr.append(sz)
    if bold:
        rpr.append(OxmlElement("w:b"))
    run._r.insert(0, rpr)
    return run


def _pPr(paragraph):
    return paragraph._p.get_or_add_pPr()


def _set_align(paragraph, value: str) -> None:
    ppr = _pPr(paragraph)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), value)


def _set_spacing(
    paragraph, *, line: Optional[float] = None, before_lines=None, after_lines=None
) -> None:
    ppr = _pPr(paragraph)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    if line is not None:
        spacing.set(qn("w:line"), str(int(line * 240)))
        spacing.set(qn("w:lineRule"), "auto")
    if before_lines is not None:
        spacing.set(qn("w:beforeLines"), str(int(before_lines * 100)))
    if after_lines is not None:
        spacing.set(qn("w:afterLines"), str(int(after_lines * 100)))


def _set_indent(paragraph, chars: Optional[float] = None) -> None:
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
    text: str,
    *,
    style: Optional[str] = None,
    size: float = 12,
    cn: str = "宋体",
    ascii_font: str = "Times New Roman",
    bold: bool = False,
    align: str = "both",
    line: float = 1.5,
    indent_chars: Optional[float] = 2,
    before_lines: Optional[float] = None,
    after_lines: Optional[float] = None,
):
    paragraph = document.add_paragraph()
    if style:
        paragraph.style = document.styles[style]
    _set_align(paragraph, align)
    _set_spacing(paragraph, line=line, before_lines=before_lines, after_lines=after_lines)
    _set_indent(paragraph, indent_chars)
    _add_run(
        paragraph,
        text,
        size=size,
        cn=cn,
        ascii_font=ascii_font,
        bold=bold,
    )
    return paragraph


def _setup_section(document) -> None:
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
    pg_mar.set(qn("w:gutter"), "567")  # 1cm 装订线
    doc_grid = section._sectPr.find(qn("w:docGrid"))
    if doc_grid is None:
        doc_grid = OxmlElement("w:docGrid")
        section._sectPr.append(doc_grid)
    doc_grid.set(qn("w:type"), "linesAndChars")
    doc_grid.set(qn("w:linePitch"), "450")


#: 各结构段落的固定文案（变异注入按前缀定位，勿随意改动）
TITLE_CN = "测试系统设计与实现"
ABS_CN_LABEL = "摘  要"
ABS_CN_PREFIX = "摘要正文段落"
KW_CN_LABEL = "关键词："
KW_CN_BODY = "格式检查；文档规范；自动化"
TITLE_EN = "Design and Implementation of a Test System"
ABS_EN_LABEL = "Abstract"
ABS_EN_PREFIX = "This paper validates the format checker"
KW_EN_LABEL = "Key words: "
KW_EN_BODY = "format checking; document standard; automation"
TOC_TITLE = "目  录"
PREFACE = "前  言"
CONCLUSION = "结  论"
THANKS = "致  谢"
REFS_TITLE = "参考文献"
REF_ENTRY = "[1] 张三. 文档格式自动检查研究[J]. 测试学报, 2024, 1(1): 1-8."
BODY_PREFIX = "正文内容段落"


def build_compliant_docx(path: Path | str) -> Path:
    """合成一份符合默认学校规范的论文 fixture（用于变异集与误报率基线）。"""
    path = Path(path)
    document = docx.Document()
    _setup_section(document)

    # ---- 前置区 ----
    add_para(document, TITLE_CN, size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, ABS_CN_LABEL, size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, f"{ABS_CN_PREFIX}一：本文围绕格式检查工具展开研究。")
    add_para(document, f"{ABS_CN_PREFIX}二：验证确定性的可复现检查流程。")

    keywords_cn = document.add_paragraph()
    _set_align(keywords_cn, "left")
    _set_spacing(keywords_cn, line=1.5, before_lines=1.0, after_lines=1.0)
    _set_indent(keywords_cn, 2)
    _add_run(keywords_cn, KW_CN_LABEL, size=12, cn="黑体")
    _add_run(keywords_cn, KW_CN_BODY, size=12, cn="宋体")

    add_para(document, TITLE_EN, size=16, ascii_font="Times New Roman", bold=True,
             align="center", before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, ABS_EN_LABEL, size=16, ascii_font="Times New Roman",
             align="center", before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, f"{ABS_EN_PREFIX} with a reproducible synthetic document.")
    add_para(document, "A second English abstract paragraph keeps the role stable.")

    keywords_en = document.add_paragraph()
    _set_align(keywords_en, "left")
    _set_spacing(keywords_en, line=1.5, before_lines=1.0, after_lines=1.0)
    _set_indent(keywords_en, 2)
    _add_run(keywords_en, KW_EN_LABEL, size=12, ascii_font="Times New Roman", bold=True)
    _add_run(keywords_en, KW_EN_BODY, size=12, ascii_font="Times New Roman")

    add_para(document, TOC_TITLE, size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    for chapter in range(1, CHAPTERS + 1):
        add_para(document, f"第{chapter}章 主题{chapter}\t{chapter}", size=12, cn="宋体",
                 align="left", indent_chars=None)
    add_para(document, f"{CONCLUSION[:-2]}\t{CHAPTERS + 1}", size=12, cn="宋体",
             align="left", indent_chars=None)

    # ---- 正文区 ----
    add_para(document, PREFACE, size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    for chapter in range(1, CHAPTERS + 1):
        add_para(document, f"第{chapter}章 主题{chapter}", style="Heading 1", size=16,
                 cn="黑体", align="center", before_lines=1.0, after_lines=1.0,
                 indent_chars=None)
        for section in range(1, 3):
            add_para(document, f"{chapter}.{section} 研究内容{chapter}-{section}",
                     style="Heading 2", size=14, cn="黑体", align="left",
                     before_lines=0.5, after_lines=0.5, indent_chars=None)
            add_para(document, f"{chapter}.{section}.1 具体问题{chapter}-{section}",
                     style="Heading 3", size=12, cn="黑体", align="left",
                     before_lines=0.5, after_lines=0.0, indent_chars=None)
            for body_index in range(1, BODY_PER_CHAPTER + 1):
                add_para(
                    document,
                    f"{BODY_PREFIX}{chapter}-{section}-{body_index}："
                    "格式检查工具用确定性规则实现可复现的检查流程。",
                )

    add_para(document, CONCLUSION, size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, THANKS, size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, REFS_TITLE, size=16, cn="黑体", align="center",
             before_lines=1.0, after_lines=1.0, indent_chars=None)
    add_para(document, REF_ENTRY, size=12, cn="宋体", align="justify",
             before_lines=0.0, after_lines=0.0, indent_chars=None)

    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(path))
    return path


# ===========================================================================
# 二、变异注入（格式类，不动文字内容；用于检出率 / 修复评测）
# ===========================================================================
def _paragraphs(document) -> list[Any]:
    return list(document.paragraphs)


def _select(document, predicate: Callable[[Any], bool]) -> list[tuple[int, Any]]:
    return [(index, para) for index, para in enumerate(document.paragraphs) if predicate(para)]


def _first(document, predicate: Callable[[Any], bool]) -> tuple[int, Any]:
    found = _select(document, predicate)
    if not found:
        raise AssertionError("变异注入目标未找到（fixture 结构可能被改动）")
    return found[0]


def _is_body(para) -> bool:
    return (para.text or "").startswith(BODY_PREFIX)


def _style_is(para, name: str) -> bool:
    return ((para.style.name if para.style is not None else "") or "").strip() == name


def _is_heading1(para) -> bool:
    return _style_is(para, "Heading 1")


def _is_heading2(para) -> bool:
    return _style_is(para, "Heading 2")


def _is_heading3(para) -> bool:
    return _style_is(para, "Heading 3")


def _is_special_title(para) -> bool:
    return _norm(para.text) in ("前言", "结论", "致谢")


def _is_title_cn(para) -> bool:
    return _norm(para.text) == _norm(TITLE_CN)


def _is_title_en(para) -> bool:
    return (para.text or "").startswith("Design and Implementation")


def _is_abs_cn_label(para) -> bool:
    return _norm(para.text) == "摘要"


def _is_abs_en_label(para) -> bool:
    return (para.text or "").strip() == "Abstract"


def _is_abs_cn_body(para) -> bool:
    return (para.text or "").startswith(ABS_CN_PREFIX)


def _is_abs_en_body(para) -> bool:
    return (para.text or "").startswith("This paper validates")


def _is_kw_cn(para) -> bool:
    return _norm(para.text).startswith("关键词")


def _is_kw_en(para) -> bool:
    return (para.text or "").lstrip().lower().startswith("key words")


def _is_toc_title(para) -> bool:
    return _norm(para.text) == "目录"


def _is_refs_title(para) -> bool:
    return _norm(para.text) == "参考文献"


# -- 底层格式改写 ----------------------------------------------------------
def _run_size(paragraph, half_points: int, run_index: int = 0) -> None:
    runs = paragraph.runs
    sz = runs[run_index]._r.rPr.find(qn("w:sz"))
    sz.set(qn("w:val"), str(half_points))


def _run_east_asia(paragraph, font: str, run_index: int = 0) -> None:
    fonts = paragraph.runs[run_index]._r.rPr.find(qn("w:rFonts"))
    fonts.set(qn("w:eastAsia"), font)


def _run_ascii(paragraph, font: str, run_index: int = 0) -> None:
    fonts = paragraph.runs[run_index]._r.rPr.find(qn("w:rFonts"))
    for attr in ("w:ascii", "w:hAnsi"):
        fonts.set(qn(attr), font)


def _drop_bold(paragraph, run_index: int = 0) -> None:
    rpr = paragraph.runs[run_index]._r.rPr
    for bold in rpr.findall(qn("w:b")):
        rpr.remove(bold)


def _spacing_element(paragraph):
    ppr = _pPr(paragraph)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    return spacing


def _set_before_lines(paragraph, lines: float) -> None:
    _spacing_element(paragraph).set(qn("w:beforeLines"), str(int(lines * 100)))


def _set_line_spacing(paragraph, multiple: float) -> None:
    spacing = _spacing_element(paragraph)
    spacing.set(qn("w:line"), str(int(multiple * 240)))
    spacing.set(qn("w:lineRule"), "auto")


def _set_jc(paragraph, value: str) -> None:
    ppr = _pPr(paragraph)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), value)


def _remove_indent(paragraph) -> None:
    ppr = _pPr(paragraph)
    for ind in ppr.findall(qn("w:ind")):
        ppr.remove(ind)


# -- 变异实现：返回被注入的段落 index 列表 --------------------------------
def _mutate_body_size(document) -> list[int]:
    chosen = _select(document, _is_body)[:5]
    for _index, para in chosen:
        _run_size(para, 21)  # 10.5pt，规范要求 12pt
    return [index for index, _para in chosen]


def _mutate_body_font(document) -> list[int]:
    chosen = _select(document, _is_body)[:5]
    for _index, para in chosen:
        _run_east_asia(para, "仿宋")  # 规范要求宋体
    return [index for index, _para in chosen]


def _mutate_body_spacing(document) -> list[int]:
    chosen = _select(document, _is_body)[:5]
    for _index, para in chosen:
        _set_line_spacing(para, 1.0)  # 规范要求 1.5 倍
    return [index for index, _para in chosen]


def _mutate_body_indent(document) -> list[int]:
    chosen = _select(document, _is_body)[:5]
    for _index, para in chosen:
        _remove_indent(para)  # 规范要求首行缩进 2 字符
    return [index for index, _para in chosen]


def _mutate_body_alignment(document) -> list[int]:
    chosen = _select(document, _is_body)[:5]
    for _index, para in chosen:
        _set_jc(para, "left")  # 规范要求两端对齐
    return [index for index, _para in chosen]


def _mutate_heading1_size(document) -> list[int]:
    chosen = _select(document, _is_heading1)[:5]
    for _index, para in chosen:
        _run_size(para, 28)  # 14pt，规范要求 16pt
    return [index for index, _para in chosen]


def _mutate_heading2_size(document) -> list[int]:
    chosen = _select(document, _is_heading2)[:5]
    for _index, para in chosen:
        _run_size(para, 32)  # 16pt，规范要求 14pt
    return [index for index, _para in chosen]


def _mutate_heading3_font(document) -> list[int]:
    chosen = _select(document, _is_heading3)[:5]
    for _index, para in chosen:
        _run_east_asia(para, "宋体")  # 规范要求黑体
    return [index for index, _para in chosen]


def _mutate_heading2_spacing(document) -> list[int]:
    chosen = _select(document, _is_heading2)[:5]
    for _index, para in chosen:
        _set_before_lines(para, 0.0)  # 规范要求段前 0.5 行
    return [index for index, _para in chosen]


def _mutate_head007(document) -> list[int]:
    chosen = _select(document, _is_special_title)
    for _index, para in chosen:
        _run_size(para, 24)  # 12pt，规范要求 16pt
    return [index for index, _para in chosen]


# -- 单点注入（同一份文档里可以与其他单点变异共存） ---------------------
def _mutate_title_cn(document) -> list[int]:
    index, para = _first(document, _is_title_cn)
    _run_size(para, 24)  # 12pt，规范要求 16pt
    return [index]


def _mutate_title_en(document) -> list[int]:
    index, para = _first(document, _is_title_en)
    _drop_bold(para)  # 规范要求英文题目加粗
    return [index]


def _mutate_abs_cn_label(document) -> list[int]:
    index, para = _first(document, _is_abs_cn_label)
    _run_size(para, 28)  # 14pt，规范要求 16pt
    return [index]


def _mutate_abs_en_label(document) -> list[int]:
    index, para = _first(document, _is_abs_en_label)
    _run_size(para, 28)
    return [index]


def _mutate_abs_cn_body(document) -> list[int]:
    index, para = _first(document, _is_abs_cn_body)
    _remove_indent(para)  # 摘要正文要求首行缩进 2 字符
    return [index]


def _mutate_abs_en_body(document) -> list[int]:
    index, para = _first(document, _is_abs_en_body)
    _run_ascii(para, "Arial")  # 规范要求 Times New Roman
    return [index]


def _mutate_kw_cn(document) -> list[int]:
    index, para = _first(document, _is_kw_cn)
    _run_east_asia(para, "宋体")  # “关键词”标签要求黑体
    return [index]


def _mutate_kw_en(document) -> list[int]:
    index, para = _first(document, _is_kw_en)
    _drop_bold(para)  # “Key words”标签要求加粗
    return [index]


def _mutate_toc_title(document) -> list[int]:
    index, para = _first(document, _is_toc_title)
    _run_size(para, 28)  # 14pt，规范要求 16pt
    return [index]


def _mutate_refs_title(document) -> list[int]:
    index, para = _first(document, _is_refs_title)
    _set_jc(para, "left")  # 规范要求居中
    return [index]


# -- 页面级注入（定位 section，不返回段落 index） ------------------------
def _mutate_page_size(document) -> list[int]:
    section = document.sections[0]
    section.page_width = Cm(20.0)
    section.page_height = Cm(28.0)
    return []


def _mutate_page_margins(document) -> list[int]:
    section = document.sections[0]
    section.left_margin = Cm(3.0)
    section.top_margin = Cm(3.0)
    return []


def _mutate_doc_grid(document) -> list[int]:
    sect_pr = document.sections[0]._sectPr
    grid = sect_pr.find(qn("w:docGrid"))
    if grid is not None:
        sect_pr.remove(grid)
    return []


# ---------------------------------------------------------------------------
# 变异集定义
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Mutation:
    """一个变异类别：类别名 + 期望命中的规则 + 注入函数。"""

    key: str
    category: str
    rule_id: str
    description: str
    apply: Callable[[Any], list[int]]
    scope: str = "paragraph"  # paragraph | page

    def target_count(self, template: Path) -> int:
        """在 fixture 副本上模拟注入，统计可注入点数量。"""
        if self.scope == "page":
            return 1
        document = docx.Document(str(template))
        return len(self.apply(document))


#: 变异集（≥8 类；每类注入数由 --injections 控制，默认每类 ≥5 处）。
#: 规则 id 必须与 rules/ 中保持一致的稳定约定；规则未实现时会在报告中跳过。
MUTATIONS: tuple[Mutation, ...] = (
    Mutation("body_size", "正文字号", "BODY-001", "正文 run 字号改为 10.5pt", _mutate_body_size),
    Mutation("body_font", "正文字体", "BODY-001", "正文中文字体改为仿宋", _mutate_body_font),
    Mutation("body_spacing", "正文行距", "BODY-002", "正文行距改为 1.0 倍", _mutate_body_spacing),
    Mutation("body_indent", "正文首行缩进", "BODY-003", "删除正文首行缩进", _mutate_body_indent),
    Mutation("body_align", "正文对齐", "BODY-004", "正文改为左对齐", _mutate_body_alignment),
    Mutation("page_size", "纸张尺寸", "PAGE-001", "页面改为 20×28cm", _mutate_page_size, "page"),
    Mutation("page_margins", "页边距", "PAGE-002", "左/上页边距改为 3cm", _mutate_page_margins, "page"),
    Mutation("doc_grid", "文档网格", "PAGE-003", "删除文档网格", _mutate_doc_grid, "page"),
    Mutation("head1_size", "一级标题字号", "HEAD-001", "一级标题字号改为 14pt", _mutate_heading1_size),
    Mutation("head2_size", "二级标题字号", "HEAD-002", "二级标题字号改为 16pt", _mutate_heading2_size),
    Mutation("head3_font", "三级标题字体", "HEAD-003", "三级标题字体改为宋体", _mutate_heading3_font),
    Mutation("head2_spacing", "二级标题段前距", "HEAD-002", "二级标题段前距改为 0", _mutate_heading2_spacing),
    Mutation("head007", "前言/结论/致谢字号", "HEAD-007", "前言/结论/致谢字号改为 12pt", _mutate_head007),
    Mutation("title_cn", "中文题目字号", "TITLE-001", "中文题目字号改为 12pt", _mutate_title_cn),
    Mutation("title_en", "英文题目加粗", "TITLE-002", "去掉英文题目加粗", _mutate_title_en),
    Mutation("abs_cn_label", "中文摘要标题字号", "ABS-001", "“摘要”标题字号改为 14pt", _mutate_abs_cn_label),
    Mutation("abs_cn_body", "摘要正文缩进", "ABS-002", "删除中文摘要正文首行缩进", _mutate_abs_cn_body),
    Mutation("abs_en_label", "英文摘要标题字号", "ABS-003", "“Abstract”标题字号改为 14pt", _mutate_abs_en_label),
    Mutation("abs_en_body", "英文摘要字体", "ABS-004", "英文摘要正文西文字体改为 Arial", _mutate_abs_en_body),
    Mutation("kw_cn", "中文关键词标签字体", "KW-001", "“关键词”标签字体改为宋体", _mutate_kw_cn),
    Mutation("kw_en", "英文关键词标签加粗", "KW-002", "去掉“Key words”加粗", _mutate_kw_en),
    Mutation("toc_title", "目录标题字号", "TOC-001", "目录标题字号改为 14pt", _mutate_toc_title),
    Mutation("refs_title", "参考文献标题对齐", "REF-001", "参考文献标题改为左对齐", _mutate_refs_title),
)


def _safe_name(value: str) -> str:
    return _SAFE_RE.sub("_", value).strip("_") or "x"


# ===========================================================================
# 三、指标计算（纯函数，单元测试直接覆盖）
# ===========================================================================
def locator_paragraph_index(locator: str) -> Optional[int]:
    """从 locator 里解析段落序号（``body/para/12「...」`` → 12）；页面类返回 None。"""
    locator = locator or ""
    if re.match(r"^(section|table|header|footer|page)/", locator):
        return None
    match = re.search(r"/(\d+)(?:「|$)", locator)
    return int(match.group(1)) if match else None


def aggregate_detection(records: Sequence[dict[str, Any]]) -> dict[str, Any]:
    """聚合注入级检出记录。

    每条记录：``{"category", "rule_id", "doc", "rule_hit", "located", "scope"}``。
    - 检出率（recall）= 定位到注入点的注入数 / 注入总数；
    - 定位准确率 = 定位到注入点的注入数 / 规则命中的注入数。
    """

    def _metrics(items: Sequence[dict[str, Any]]) -> dict[str, Any]:
        total = len(items)
        rule_hits = sum(1 for item in items if item.get("rule_hit"))
        located = sum(1 for item in items if item.get("located"))
        return {
            "injections": total,
            "rule_hits": rule_hits,
            "located": located,
            "recall": (located / total) if total else 0.0,
            "locate_accuracy": (located / rule_hits) if rule_hits else 0.0,
        }

    categories: list[dict[str, Any]] = []
    keys: list[str] = []
    for record in records:
        if record["category"] not in keys:
            keys.append(record["category"])
    for key in keys:
        subset = [item for item in records if item["category"] == key]
        payload = _metrics(subset)
        payload["category"] = key
        payload["rule_id"] = subset[0]["rule_id"]
        payload["scope"] = subset[0].get("scope", "paragraph")
        categories.append(payload)

    per_rule: list[dict[str, Any]] = []
    rule_ids: list[str] = []
    for record in records:
        if record["rule_id"] not in rule_ids:
            rule_ids.append(record["rule_id"])
    for rule_id in rule_ids:
        subset = [item for item in records if item["rule_id"] == rule_id]
        payload = _metrics(subset)
        payload["rule_id"] = rule_id
        payload["categories"] = sorted(
            {item["category"] for item in subset}
        )
        per_rule.append(payload)

    overall = _metrics(records)
    overall["categories"] = categories
    overall["per_rule"] = per_rule
    return overall


def aggregate_false_positives(
    docs: Sequence[dict[str, Any]], evaluated_rules: Iterable[str]
) -> dict[str, Any]:
    """合规文档上的误报统计。

    参数 ``docs``：每项 ``{"issues": list[Issue], "paragraphs": int}``；
    ``evaluated_rules``：本次评测覆盖的规则 id（只统计这些规则，避免把尚未覆盖
    fixture 结构的新规则算成误报）。
    """
    evaluated = set(evaluated_rules)
    rule_hits: Counter[str] = Counter()
    total_paragraphs = 0
    flagged_paragraphs = 0
    page_level_issues = 0
    other_hits: Counter[str] = Counter()

    for doc in docs:
        total_paragraphs += int(doc.get("paragraphs", 0))
        paragraphs_flagged: set[int] = set()
        for issue in doc.get("issues", []):
            if issue.rule_id not in evaluated:
                other_hits[issue.rule_id] += 1
                continue
            rule_hits[issue.rule_id] += 1
            index = locator_paragraph_index(issue.locator)
            if index is None:
                page_level_issues += 1
            else:
                paragraphs_flagged.add(index)
        flagged_paragraphs += len(paragraphs_flagged)

    return {
        "good_docs": len(docs),
        "paragraphs": total_paragraphs,
        "flagged_paragraphs": flagged_paragraphs,
        "fp_rate": (flagged_paragraphs / total_paragraphs) if total_paragraphs else 0.0,
        "page_level_issues": page_level_issues,
        "rule_hits": dict(sorted(rule_hits.items())),
        "other_rule_hits": dict(sorted(other_hits.items())),
    }


def summarize_repairs(repairs: Sequence[dict[str, Any]]) -> dict[str, Any]:
    """聚合修复评测记录。"""
    total = len(repairs)
    injected_ok = sum(1 for item in repairs if item.get("injected_ok"))
    full_ok = sum(1 for item in repairs if item.get("full_ok"))
    hash_ok = sum(1 for item in repairs if item.get("hash_ok"))
    fix_seconds = [float(item.get("fix_s", 0.0)) for item in repairs if item.get("fix_s")]
    return {
        "docs": total,
        "injected_pass_rate": (injected_ok / total) if total else 0.0,
        "full_pass_rate": (full_ok / total) if total else 0.0,
        "hash_unchanged_rate": (hash_ok / total) if total else 0.0,
        "mean_fix_s": (sum(fix_seconds) / len(fix_seconds)) if fix_seconds else 0.0,
        "total_fix_s": sum(fix_seconds),
        "failures": [item for item in repairs if not item.get("injected_ok")],
    }


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


# ===========================================================================
# 四、评测主流程
# ===========================================================================
def _open_copy(template: Path, target: Path):
    shutil.copyfile(template, target)
    return docx.Document(str(target))


def _apply_all(template: Path, target: Path, mutations: Sequence[Mutation]):
    """在 fixture 副本上依次注入 mutations，返回 {类别: 注入位置列表}。

    段落级变异记录被注入的段落 index；页面级变异用 ``[0]`` 占位（定位靠 section）。
    """
    document = _open_copy(template, target)
    targets: dict[str, list[int]] = {}
    for mutation in mutations:
        found = list(mutation.apply(document))
        targets[mutation.category] = found if mutation.scope == "paragraph" else [0]
    document.save(str(target))
    return targets


def run_checks_for(
    path: Path, ruleset: RuleSet, profile: str = DEFAULT_PROFILE
) -> tuple[list[Issue], Optional[str]]:
    """跑全部已实现检查器；异常收敛为 error 字符串（评测不中断）。"""
    try:
        model = DocModel.load(path)
        roles = classify(model, profile)
        return run_checks(model, roles, ruleset, profile=profile), None
    except Exception as exc:  # noqa: BLE001 - 评测需要对个别失败容错
        return [], f"{type(exc).__name__}: {exc}"


def _implemented_ids(ruleset: RuleSet) -> list[str]:
    """检查器已实现、且当前规则集里存在的规则 id（动态，不写死数量）。"""
    ids: list[str] = []
    for rule_id in list(IMPLEMENTED_RULES):
        if rule_id in ruleset.rules and rule_id not in ids:
            ids.append(rule_id)
    return ids


def _detection_section(
    template: Path,
    workdir: Path,
    mutations: Sequence[Mutation],
    *,
    injections: int,
    ruleset: RuleSet,
    profile: str,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """构建变异文档、跑检查、返回 (检出指标, 附加信息)。"""
    records: list[dict[str, Any]] = []
    skipped: list[dict[str, str]] = []
    built: list[dict[str, Any]] = []
    check_errors: list[dict[str, str]] = []

    # 分类：单点变异（可与其它单点变异合并到同一份文档）vs 多点变异（单独建文档）
    dedicated: list[Mutation] = []
    grouped: list[Mutation] = []
    for mutation in mutations:
        try:
            count = mutation.target_count(template)
        except Exception as exc:  # noqa: BLE001 - fixture 不匹配时跳过该类别
            skipped.append(
                {"category": mutation.category, "rule_id": mutation.rule_id,
                 "reason": f"无法在 fixture 中定位注入点：{type(exc).__name__}: {exc}"}
            )
            continue
        if count <= 0:
            skipped.append(
                {"category": mutation.category, "rule_id": mutation.rule_id,
                 "reason": "fixture 中没有可注入位置"}
            )
            continue
        if mutation.scope == "paragraph" and count >= 2:
            dedicated.append(mutation)
        else:
            grouped.append(mutation)

    # 1) 多点类别：每份文档注入整类问题（每类 ≥ injections 处）
    for mutation in dedicated:
        count = mutation.target_count(template)
        docs_needed = max(1, math.ceil(injections / count))
        for doc_index in range(docs_needed):
            name = f"mut_{_safe_name(mutation.key)}_{doc_index}.docx"
            target = workdir / name
            try:
                targets = _apply_all(template, target, [mutation])[mutation.category]
            except Exception as exc:  # noqa: BLE001 - 构建失败只记录，不影响其它类别
                skipped.append(
                    {"category": mutation.category, "rule_id": mutation.rule_id,
                     "reason": f"构造变异文档失败：{type(exc).__name__}: {exc}"}
                )
                break
            prepared = [{"mutation": mutation, "targets": targets, "target": target}]
            built.append({"name": name, "items": prepared, "kind": "dedicated"})

    # 2) 单点类别：把所有单点变异注入到同一批文档（每类 gets injections 处）
    if grouped:
        for doc_index in range(max(1, injections)):
            name = f"mut_grouped_{doc_index}.docx"
            target = workdir / name
            try:
                targets = _apply_all(template, target, grouped)
            except Exception as exc:  # noqa: BLE001
                for mutation in grouped:
                    skipped.append(
                        {"category": mutation.category, "rule_id": mutation.rule_id,
                         "reason": f"构造变异文档失败：{type(exc).__name__}: {exc}"}
                    )
                break
            prepared = [
                {"mutation": mutation, "targets": targets.get(mutation.category, []),
                 "target": target}
                for mutation in grouped
            ]
            built.append({"name": name, "items": prepared, "kind": "grouped"})

    # 3) 逐份检查 + 记录注入级结果
    for doc in built:
        issues, error = run_checks_for(doc["items"][0]["target"], ruleset, profile)
        if error:
            check_errors.append({"doc": doc["name"], "error": error})
        # 同一份文档里可能有多个类别，合并所有注入目标
        for item in doc["items"]:
            mutation: Mutation = item["mutation"]
            rule_issues = [issue for issue in issues if issue.rule_id == mutation.rule_id]
            for target_index in item["targets"]:
                rule_hit = bool(rule_issues)
                if mutation.scope == "page":
                    located = any(issue.locator.startswith("section/") for issue in rule_issues)
                else:
                    located = any(
                        locator_paragraph_index(issue.locator) == target_index
                        for issue in rule_issues
                    )
                records.append(
                    {
                        "category": mutation.category,
                        "rule_id": mutation.rule_id,
                        "scope": mutation.scope,
                        "doc": doc["name"],
                        "target": target_index,
                        "rule_hit": rule_hit,
                        "located": located,
                    }
                )

    metrics = aggregate_detection(records)
    metrics["skipped_categories"] = skipped
    metrics["check_errors"] = check_errors
    metrics["docs"] = len(built)
    return metrics, {"skipped": skipped, "check_errors": check_errors, "docs": len(built)}


def _false_positive_section(
    template: Path,
    workdir: Path,
    *,
    good_docs: int,
    evaluated_rules: Sequence[str],
    ruleset: RuleSet,
    profile: str,
) -> dict[str, Any]:
    """合规文档误报率：同一 fixture 复制 good_docs 份，逐份跑检查。"""
    docs: list[dict[str, Any]] = []
    check_errors: list[dict[str, str]] = []
    for index in range(max(1, good_docs)):
        target = workdir / f"good_{index}.docx"
        shutil.copyfile(template, target)
        issues, error = run_checks_for(target, ruleset, profile)
        if error:
            check_errors.append({"doc": target.name, "error": error})
            continue
        try:
            model = DocModel.load(target)
            paragraphs = len([para for para in model.paragraphs if not para.is_empty])
        except Exception:  # pragma: no cover - 已成功检查过，理论不会触发
            paragraphs = 0
        docs.append({"issues": issues, "paragraphs": paragraphs})

    payload = aggregate_false_positives(docs, evaluated_rules)
    payload["check_errors"] = check_errors
    return payload


def _repair_section(
    template: Path,
    workdir: Path,
    *,
    docs: int,
    ruleset: RuleSet,
    profile: str,
) -> dict[str, Any]:
    """修复评测：构造混合变异文档 → fix_document → 统计通过率 / hash / 耗时。"""
    facets = [mutation for mutation in MUTATIONS if mutation.rule_id in SAFE_RULE_IDS]
    if not facets:
        return {
            "docs": 0,
            "injected_pass_rate": 0.0,
            "full_pass_rate": 0.0,
            "hash_unchanged_rate": 0.0,
            "mean_fix_s": 0.0,
            "total_fix_s": 0.0,
            "failures": [],
            "note": "当前修复白名单为空，跳过修复评测",
        }

    doc_count = max(1, int(docs))
    repair_dir = workdir / "repair"
    fixed_dir = workdir / "fixed"
    repair_dir.mkdir(parents=True, exist_ok=True)
    fixed_dir.mkdir(parents=True, exist_ok=True)

    paths: list[Path] = []
    injected_rules: dict[str, set[str]] = {}
    for index in range(doc_count):
        chosen: list[Mutation] = []
        for offset in range(5):
            mutation = facets[(index * 3 + offset) % len(facets)]
            if mutation.key not in {item.key for item in chosen}:
                chosen.append(mutation)
        target = repair_dir / f"mut_doc_{index:02d}.docx"
        try:
            _apply_all(template, target, chosen)
        except Exception as exc:  # pragma: no cover - 合成文档失败时降级为跳过
            injected_rules[str(target)] = set()
            paths.append(target)
            continue
        paths.append(target)
        injected_rules[str(target)] = {mutation.rule_id for mutation in chosen}

    # 批量检查计时（吞吐评测）
    started = time.perf_counter()
    check_results = check_files(paths, ruleset=ruleset, profile=profile, ruleset_id="eval")
    batch_elapsed = time.perf_counter() - started
    check_durations = [result.duration_s for result in check_results]

    repairs: list[dict[str, Any]] = []
    for index, path in enumerate(paths):
        wanted = injected_rules.get(str(path), set())
        out_path = fixed_dir / f"fixed_{index:02d}.docx"
        started = time.perf_counter()
        try:
            result = fix_document(path, out_path, profile=profile, ruleset=ruleset)
            fix_seconds = time.perf_counter() - started
            error = None
        except Exception as exc:  # noqa: BLE001 - 修复失败也计入评测（不中断）
            fix_seconds = time.perf_counter() - started
            repairs.append(
                {
                    "doc": path.name,
                    "injected_rules": sorted(wanted),
                    "injected_ok": False,
                    "full_ok": False,
                    "hash_ok": False,
                    "fix_s": fix_seconds,
                    "error": f"{type(exc).__name__}: {exc}",
                    "recheck_errors": None,
                }
            )
            continue

        error_issues = [issue for issue in result.recheck_issues if issue.severity == "error"]
        injected_error = [issue for issue in error_issues if issue.rule_id in wanted]
        repairs.append(
            {
                "doc": path.name,
                "injected_rules": sorted(wanted),
                "injected_ok": not injected_error,
                "full_ok": not error_issues,
                "hash_ok": bool(result.content_unchanged),
                "fix_s": fix_seconds,
                "error": error,
                "recheck_error_count": len(error_issues),
                "recheck_remaining_rules": sorted({issue.rule_id for issue in error_issues}),
            }
        )

    summary = summarize_repairs(repairs)
    summary["mean_check_s"] = (
        sum(check_durations) / len(check_durations) if check_durations else 0.0
    )
    summary["fix_failures"] = [
        item for item in repairs if item.get("error") is not None
    ]
    summary["batch_check_elapsed_s"] = batch_elapsed
    summary["batch_check_docs"] = len(paths)
    summary["docs_per_s"] = (len(paths) / batch_elapsed) if batch_elapsed else 0.0
    summary["ms_per_doc"] = (
        (batch_elapsed / len(paths) * 1000.0) if paths and batch_elapsed else 0.0
    )
    summary["check_errors"] = [
        {"doc": result.path.name, "error": result.error}
        for result in check_results
        if result.error is not None
    ]
    return summary


def run_evaluation(
    *, docs: int = DEFAULT_DOCS, injections: int = DEFAULT_INJECTIONS
) -> dict[str, Any]:
    """跑完整评测，返回 JSON 可序列化的指标 payload（不写文件）。"""
    started = time.perf_counter()
    profile, ruleset = resolve_profile(None)
    discipline = profile.discipline or DEFAULT_PROFILE
    implemented = _implemented_ids(ruleset)

    skipped: list[dict[str, str]] = []

    # 只评测「已实现」的规则对应的变异类别；未实现的规则跳过并注明
    mutations: list[Mutation] = []
    for mutation in MUTATIONS:
        if mutation.rule_id in implemented:
            mutations.append(mutation)
        else:
            skipped.append(
                {
                    "category": mutation.category,
                    "rule_id": mutation.rule_id,
                    "reason": "当前检查器尚未实现该规则（已实现规则集合动态读取）",
                }
            )

    evaluated_rules = sorted({mutation.rule_id for mutation in mutations})
    evaluated_categories = sorted({mutation.category for mutation in mutations})

    with tempfile.TemporaryDirectory(prefix="tas-eval-") as tmp:
        workdir = Path(tmp)
        template = build_compliant_docx(workdir / "template.docx")

        detection, detection_extra = _detection_section(
            template,
            workdir,
            mutations,
            injections=max(1, int(injections)),
            ruleset=ruleset,
            profile=discipline,
        )
        skipped.extend(detection_extra["skipped"])

        # 评测覆盖的 P0 规则（规则库 phase 字段；warning 级 P0 也算 P0 规则）
        p0_rules = [
            rule_id
            for rule_id in implemented
            if ruleset.rules[rule_id].phase == "P0"
        ]
        missing_p0 = [rule_id for rule_id in p0_rules if rule_id not in evaluated_rules]

        fp = _false_positive_section(
            template,
            workdir,
            good_docs=3,
            evaluated_rules=evaluated_rules,
            ruleset=ruleset,
            profile=discipline,
        )
        repair = _repair_section(
            template,
            workdir,
            docs=max(1, int(docs)),
            ruleset=ruleset,
            profile=discipline,
        )

    total_elapsed = time.perf_counter() - started
    throughtput_docs = int(repair.get("batch_check_docs", 0))
    payload: dict[str, Any] = {
        "meta": {
            "tool": "tas",
            "version": TAS_VERSION,
            "generated_at": _now(),
            "python": platform.python_version(),
            "platform": platform.platform(),
            "profile_id": profile.id,
            "profile_name": profile.name,
            "discipline": discipline,
            "ruleset_id": profile.id,
            "ruleset_total_rules": len(ruleset),
            "implemented_rules": implemented,
            "implemented_count": len(implemented),
            "implemented_p0_rules": p0_rules,
            "evaluated_rules": evaluated_rules,
            "evaluated_categories": evaluated_categories,
            "unevaluated_implemented_p0_rules": missing_p0,
            "skipped_categories": skipped,
            "injections_per_category": max(1, int(injections)),
            "eval_docs": max(1, int(docs)),
            "good_docs": 3,
            "total_elapsed_s": total_elapsed,
            "notes": [
                "评测数据全部由 python-docx 现场合成，不读取真实学生文档、不联网。",
                "只统计当前检查器已实现的规则（IMPLEMENTED_RULES 动态读取），"
                "未实现规则的变异类别跳过；未评测的已实现规则不计入检出率分母。",
                "检出率 = 命中对应规则且定位到注入段落的比例；误报率 = 合规文档中"
                "被本次评测规则报出的段落比例。",
                "修复评测使用 tas.fixer.fix_document，只对修复白名单（SAFE_RULE_IDS）"
                "内的规则注入变异；内容 hash 一致性由 fixer 自身保证。",
            ],
        },
        "detection": detection,
        "false_positives": fp,
        "repair": repair,
        "throughput": {
            "docs": throughtput_docs,
            "elapsed_s": float(repair.get("batch_check_elapsed_s", 0.0)),
            "docs_per_s": float(repair.get("docs_per_s", 0.0)),
            "ms_per_doc": float(repair.get("ms_per_doc", 0.0)),
            "mean_check_s": float(repair.get("mean_check_s", 0.0)),
        },
        "timings": {
            "total_elapsed_s": total_elapsed,
        },
    }
    return payload


# ===========================================================================
# 五、报告渲染 / 落盘 / CLI
# ===========================================================================
def _pct(value: float) -> str:
    return f"{float(value or 0.0):.1%}"


def _num(value: float, digits: int = 3) -> str:
    return f"{float(value or 0.0):.{digits}f}"


#: 未纳入评测的已实现 P0 规则的补充说明（仅展示用；未知规则用通用文案兜底）
_UNEVALUATED_P0_HINTS = {
    "HEAD-006": "编号连续性：需修改文字/编号，不属于格式变异",
    "FIG-001": "图题注：需要带图片/题注的 fixture，本次合成文档未覆盖",
}


def render_markdown(payload: dict[str, Any]) -> str:
    """把指标 payload 渲染成 markdown 报告（表格：指标/数值/说明）。"""
    meta = payload["meta"]
    detection = payload["detection"]
    fp = payload["false_positives"]
    repair = payload["repair"]
    throughput = payload["throughput"]

    implemented_p0 = list(meta.get("implemented_p0_rules") or [])
    implemented_p1_count = len(meta["implemented_rules"]) - len(implemented_p0)
    actual_injections = [item["injections"] for item in detection["categories"]] or [0]

    lines: list[str] = []
    lines.append("# tas 评测报告（合成变异集）")
    lines.append("")
    lines.append(
        f"> 生成时间：{meta['generated_at']} ｜ tas {meta['version']} ｜ "
        f"Python {meta['python']} ｜ 规则集 `{meta['ruleset_id']}`（{meta['profile_name']}）"
    )
    lines.append(">")
    lines.append(
        "> 评测数据全部由 python-docx 现场合成（不读取真实学生文档、不联网）；"
        "只统计**当前检查器已实现**的规则，规则数量随迭代增加时本报告自动适配。"
    )
    lines.append("")

    # ---- 1. 规则覆盖 ----------------------------------------------------
    lines.append("## 1. 规则覆盖（动态读取，不写死数量）")
    lines.append("")
    lines.append("| 指标 | 数值 | 说明 |")
    lines.append("|---|---|---|")
    lines.append(
        f"| 规则库规则数 | {meta['ruleset_total_rules']} | `rules/rules.json` 当前条目 |"
    )
    lines.append(
        f"| 检查器已实现规则数 | {meta['implemented_count']} | 动态读取 "
        "`tas.checkers.IMPLEMENTED_RULES`；其中 "
        f"P0 {len(implemented_p0)} 条、P1 {implemented_p1_count} 条 |"
    )
    lines.append(
        f"| 本次评测覆盖规则数 | {len(meta['evaluated_rules'])} | "
        f"全部为已实现 P0 规则：{'、'.join(meta['evaluated_rules'])} |"
    )
    lines.append(
        f"| 变异类别数 | {len(meta['evaluated_categories'])} | 每类 ≥ "
        f"{meta['injections_per_category']} 处注入；实际最少 {min(actual_injections)} 处 |"
    )
    if meta["unevaluated_implemented_p0_rules"]:
        hints = "；".join(
            f"{rule_id}（{_UNEVALUATED_P0_HINTS.get(rule_id, '本次变异集未覆盖')}）"
            for rule_id in meta["unevaluated_implemented_p0_rules"]
        )
        lines.append(
            f"| 未纳入评测的已实现 P0 规则 | "
            f"{len(meta['unevaluated_implemented_p0_rules'])} | "
            f"{hints}；未进入检出率分母 |"
        )
    else:
        lines.append("| 未纳入评测的已实现 P0 规则 | 0 | 全部已实现 P0 规则都有变异覆盖 |")
    if meta["skipped_categories"]:
        skipped_text = "；".join(
            f"{item['category']}（{item['rule_id']}）：{item['reason']}"
            for item in meta["skipped_categories"]
        )
        lines.append(f"| 跳过类别 | {len(meta['skipped_categories'])} | {skipped_text} |")
    lines.append("")

    # ---- 2. 检出率 ------------------------------------------------------
    lines.append("## 2. 变异检出率（P0 规则）")
    lines.append("")
    lines.append("| 指标 | 数值 | 说明 |")
    lines.append("|---|---|---|")
    lines.append(
        f"| 注入点总数 | {detection['injections']} | "
        f"{len(meta['evaluated_categories'])} 类格式错误 |"
    )
    lines.append(
        f"| 检出数（规则命中且定位正确） | {detection['located']} | 注入段落级定位 |"
    )
    lines.append(
        f"| **检出率（recall）** | **{_pct(detection['recall'])}** | "
        "定位到注入点的注入数 / 注入总数 |"
    )
    lines.append(
        f"| 定位准确率 | {_pct(detection['locate_accuracy'])} | "
        "定位到注入点 / 规则命中 |"
    )
    lines.append("")

    lines.append("| 变异类别 | 规则 | 注入数 | 检出数 | 检出率 | 定位准确率 |")
    lines.append("|---|---|---|---|---|---|")
    for item in detection["categories"]:
        lines.append(
            f"| {item['category']} | {item['rule_id']} | {item['injections']} | "
            f"{item['located']} | {_pct(item['recall'])} | {_pct(item['locate_accuracy'])} |"
        )
    lines.append("")

    # ---- 3. 误报率 ------------------------------------------------------
    lines.append("## 3. 误报率（合规 fixture）")
    lines.append("")
    lines.append("| 指标 | 数值 | 说明 |")
    lines.append("|---|---|---|")
    lines.append(f"| 合规文档数 | {fp['good_docs']} | 同结构合规 fixture 复制 |")
    lines.append(f"| 段落总数 | {fp['paragraphs']} | 非空段落 |")
    lines.append(
        f"| 被误报段落数 | {fp['flagged_paragraphs']} | 本次评测规则报出的不同段落 |"
    )
    lines.append(
        f"| **误报率（FP rate）** | **{_pct(fp['fp_rate'])}** | 被误报段落 / 段落总数 |"
    )
    lines.append(
        f"| 页面级误报 | {fp['page_level_issues']} | 页边距/纸张/网格类（非段落定位） |"
    )
    if fp["rule_hits"]:
        hits = "、".join(f"{rule}({count})" for rule, count in fp["rule_hits"].items())
        lines.append(f"| 误报规则明细 | {hits} | 合规文档不应命中这些规则 |")
    if fp["other_rule_hits"]:
        others = "、".join(
            f"{rule}({count})" for rule, count in fp["other_rule_hits"].items()
        )
        lines.append(
            f"| 评测集外规则命中（不计 FP） | {sum(fp['other_rule_hits'].values())} | "
            f"{others}；这些新规则需要扩展 fixture 才能评测，未计入误报率 |"
        )
    if fp.get("check_errors"):
        lines.append(
            f"| 检查失败 | {len(fp['check_errors'])} | 见下方“已知限制/失败明细” |"
        )
    lines.append("")

    # ---- 4. 修复评测 ----------------------------------------------------
    lines.append("## 4. 修复评测（`tas.fixer.fix_document`）")
    lines.append("")
    lines.append("| 指标 | 数值 | 说明 |")
    lines.append("|---|---|---|")
    lines.append(
        f"| 修复文档数 | {repair['docs']} | 每份注入 5 类可安全修复的格式问题 |"
    )
    lines.append(
        f"| **修复后 P0 通过率（注入规则）** | **{_pct(repair['injected_pass_rate'])}** | "
        "注入规则在复检中 error 清零的文档比例 |"
    )
    lines.append(
        f"| 全量 error 清零率（参考） | {_pct(repair['full_pass_rate'])} | "
        "复检无任何 error（含 fixture 未覆盖的新规则） |"
    )
    lines.append(
        f"| 内容 hash 一致率 | {_pct(repair['hash_unchanged_rate'])} | "
        "只改格式、不改内容的硬约束 |"
    )
    lines.append(
        f"| 平均单份检查耗时 | {_num(repair.get('mean_check_s', 0.0))} s | "
        "run_checks 全规则检查 |"
    )
    lines.append(
        f"| 平均单份修复耗时 | {_num(repair['mean_fix_s'])} s | 含读写 docx + 复检 |"
    )
    if repair.get("fix_failures"):
        lines.append(
            f"| 修复异常 | {len(repair['fix_failures'])} | 见“已知限制/失败明细” |"
        )
    lines.append("")

    # ---- 5. 批量吞吐 ----------------------------------------------------
    lines.append("## 5. 批量吞吐（`tas.batch.check_files`）")
    lines.append("")
    lines.append("| 指标 | 数值 | 说明 |")
    lines.append("|---|---|---|")
    lines.append(f"| 批量文档数 | {throughput['docs']} | 单进程顺序检查 |")
    lines.append(
        f"| 批量总耗时 | {_num(throughput['elapsed_s'])} s | 含模型加载/角色识别/全部检查器 |"
    )
    lines.append(f"| 吞吐 | {_num(throughput['docs_per_s'], 2)} 份/秒 | 批量检查 |")
    lines.append(
        f"| 单份平均耗时 | {_num(throughput['ms_per_doc'], 1)} ms | 批量路径实测 |"
    )
    lines.append("")

    # ---- 6. 复现与限制 --------------------------------------------------
    lines.append("## 6. 复现方式与已知限制")
    lines.append("")
    lines.append("```powershell")
    lines.append('$env:PYTHONPATH = "D:\\zdsh\\work\\tas\\src"')
    lines.append("python eval/run_eval.py --out eval/report.md --docs 12")
    lines.append("```")
    lines.append("")
    lines.append("- 评测数据全部合成，不含真实学生信息；不联网、不修改仓库数据。")
    lines.append(
        "- 检查器仍在迭代：本报告只统计 `IMPLEMENTED_RULES` 里已实现的规则；"
        "新增检查器后重跑本脚本即可更新数字。"
    )
    lines.append(
        "- 误报率只统计本次评测覆盖的规则；合规 fixture 尚未覆盖的新规则若命中，"
        "在“评测集外规则命中”一行单独列出，不计入误报率。"
    )
    lines.append(
        "- 修复评测注入的都是 `SAFE_RULE_IDS`（只改格式）内的问题；编号连续性、"
        "参考文献著录等涉及文字/结构的规则只报告不修复，不在修复通过率分母内。"
    )
    lines.append(
        f"- 本次评测总耗时 {_num(meta.get('total_elapsed_s', 0.0), 2)} s"
        f"（--docs {meta['eval_docs']}，每类注入 {meta['injections_per_category']} 处）。"
    )
    lines.append("")
    return "\n".join(lines) + "\n"


def write_reports(
    payload: dict[str, Any], out_path: Path | str, json_path: Optional[Path | str] = None
) -> tuple[Path, Path]:
    """把 payload 写成 markdown 报告 + JSON；返回 (markdown 路径, json 路径)。"""
    out = Path(out_path)
    json_out = Path(json_path) if json_path is not None else out.with_suffix(".json")
    out.parent.mkdir(parents=True, exist_ok=True)
    json_out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render_markdown(payload), encoding="utf-8")
    json_out.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return out, json_out


def _parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="run_eval.py",
        description="tas 评测：合成变异集上的检出率/误报率/修复效果/批量吞吐",
    )
    parser.add_argument(
        "--out",
        default=str(ROOT / "eval" / "report.md"),
        help="markdown 报告输出路径（默认 eval/report.md）",
    )
    parser.add_argument(
        "--json-out",
        default=None,
        help="JSON 指标输出路径（默认与 --out 同名 .json）",
    )
    parser.add_argument(
        "--docs",
        type=int,
        default=DEFAULT_DOCS,
        help=f"修复/吞吐评测的文档数（默认 {DEFAULT_DOCS}，契约要求 ≥10）",
    )
    parser.add_argument(
        "--injections",
        type=int,
        default=DEFAULT_INJECTIONS,
        help=f"每个变异类别的最少注入数（默认 {DEFAULT_INJECTIONS}，契约要求 ≥5）",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = _parse_args(argv)
    payload = run_evaluation(
        docs=max(1, int(args.docs)), injections=max(1, int(args.injections))
    )
    out, json_out = write_reports(payload, args.out, args.json_out)

    detection = payload["detection"]
    repair = payload["repair"]
    throughput = payload["throughput"]
    meta = payload["meta"]
    print("tas 评测完成")
    print("=" * 64)
    print(
        f"规则覆盖：已实现 {meta['implemented_count']} 条 / 规则库 "
        f"{meta['ruleset_total_rules']} 条；本次评测覆盖 {len(meta['evaluated_rules'])} 条、"
        f"{len(meta['evaluated_categories'])} 类变异"
    )
    print(
        f"检出率  ：{_pct(detection['recall'])}"
        f"（定位准确率 {_pct(detection['locate_accuracy'])}，"
        f"注入 {detection['injections']} 处）"
    )
    print(f"误报率  ：{_pct(payload['false_positives']['fp_rate'])}（合规 fixture）")
    print(
        f"修复后 P0 通过率：{_pct(repair['injected_pass_rate'])}；"
        f"内容 hash 一致率：{_pct(repair['hash_unchanged_rate'])}；"
        f"平均修复 {_num(repair['mean_fix_s'])} s/份"
    )
    print(
        f"批量吞吐：{_num(throughput['docs_per_s'], 2)} 份/秒"
        f"（{_num(throughput['ms_per_doc'], 1)} ms/份，{throughput['docs']} 份）"
    )
    print(f"报告已写入：{out} ｜ {json_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
