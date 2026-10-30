"""参考文献条目检查：REF-002（条目格式）/ REF-003（正文引用标注 + 著录要素）。

复用 ``tas.refs`` 的解析与引用一致性分析（只读）：

- ``parse_references`` 负责把参考文献区解析成结构化条目；
- ``analyze_citations`` 负责正文 [n] 引用与文献表的一致性（越界/缺号/乱序/上标）。

检查范围说明（有意取舍，见最终报告）：

- REF-002 的"悬挂缩进 2 字符"只在段落**显式设置了缩进**时校验：
  完全未设置缩进时不上报——自动编号列表的悬挂缩进保存在 numbering.xml 中，
  段落层不可见，硬报会对"用 Word 自动编号"的文献表误报；
- REF-003 不报"文献表里有条目但正文未引用"（该结论属"内容完整性"，由
  ``tas refs`` 命令完整呈现）；这里只报正文引用本身的错误（越界、缺号、乱序、非上标）
  与条目著录要素缺失（类型标识/年份/电子资源 URL）。
"""

from __future__ import annotations

import re
from typing import Optional

from ..config import Rule, RuleSet
from ..docx_io import DocModel
from ..refs import analyze_citations, parse_references
from ..refs.parse import ReferenceEntry
from ..report import Issue
from ..roles import Roles
from ..rule_params import expected_params
from .helpers import (
    CHAR_TOL,
    LINE_TOL,
    PT_PER_LINE,
    SPACING_TOL,
    close,
    first_run,
    loc_para,
    make_issue,
    run_font_problems,
)

REF_ENTRY_FALLBACK = {
    "size_pt": 12,
    "east_asia": "宋体",
    "ascii": "Times New Roman",
    "line_spacing": 1.5,
    "before_lines": 0,
    "after_lines": 0,
    "number_format": "^\\[\\d+\\]",
    "hanging_indent_chars": 200,
}
REF_CITATION_FALLBACK = {
    "in_text_citation_superscript": True,
    "sequential_numbering": True,
    "style": "GB/T 7714-2015",
    "categories": ["J", "M", "C", "D", "P", "S", "N", "DB/CP/EB"],
}
ELECTRONIC_TYPES = ("EB", "DB", "CP", "CM")


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------
def _rule(ruleset: RuleSet, rule_id: str) -> Optional[Rule]:
    return ruleset.rules.get(rule_id)


def _num(params: dict, key: str, fallback: float) -> float:
    try:
        return float(params.get(key, fallback))
    except (TypeError, ValueError):
        return fallback


def _text(params: dict, key: str, fallback: str = "") -> str:
    value = params.get(key, fallback)
    return str(value) if value is not None else fallback


def _entry_paragraph(model: DocModel, entry: ReferenceEntry):
    return model.by_index.get(entry.paragraph_index)


def _direct_lines(para, *, before: bool) -> Optional[float]:
    """只取段落自身显式设置的段前/段后（行）；样式继承的默认值不算违规。"""
    lines = para.direct_before_lines if before else para.direct_after_lines
    if lines is not None:
        return lines
    pt = para.direct_before_pt if before else para.direct_after_pt
    if pt is not None:
        return pt / PT_PER_LINE
    return None


def _hanging_chars(para) -> Optional[float]:
    """悬挂缩进字符数：优先 w:hangingChars，其次按字号把 w:hanging(twips) 换算。"""
    if para.hanging_chars is not None:
        return para.hanging_chars
    if para.hanging_twips is None:
        return None
    size = 12.0
    run = first_run(para)
    if run is not None and run.size_pt:
        size = run.size_pt
    return para.hanging_twips / (20.0 * size) if size else None


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def check_references(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng"
) -> list[Issue]:
    if ruleset.rules.get("REF-002") is None and ruleset.rules.get("REF-003") is None:
        return []
    entries = parse_references(model, roles)
    issues: list[Issue] = []
    issues.extend(_check_entry_format(model, entries, ruleset))
    issues.extend(_check_citations(model, roles, entries, ruleset))
    return issues


# ---------------------------------------------------------------------------
# REF-002 参考文献条目格式
# ---------------------------------------------------------------------------
def _check_entry_format(
    model: DocModel, entries: list[ReferenceEntry], ruleset: RuleSet
) -> list[Issue]:
    rule = _rule(ruleset, "REF-002")
    if rule is None or not entries:
        return []
    params = expected_params(ruleset, "REF-002", REF_ENTRY_FALLBACK)
    size = _num(params, "size_pt", 12.0)
    east_asia = _text(params, "east_asia", "宋体")
    ascii_font = _text(params, "ascii", "Times New Roman")
    line_spacing = _num(params, "line_spacing", 1.5)
    before_lines = _num(params, "before_lines", 0.0) / 100.0
    after_lines = _num(params, "after_lines", 0.0) / 100.0
    hanging = _num(params, "hanging_indent_chars", 200.0) / 100.0
    number_format_text = _text(params, "number_format", "^\\[\\d+\\]")
    try:
        number_re = re.compile(number_format_text)
    except re.error:
        number_re = re.compile(r"^\[\d+\]")
    expected = {
        "size_pt": size,
        "east_asia": east_asia,
        "ascii": ascii_font,
        "line_spacing": line_spacing,
        "before_lines": before_lines,
        "after_lines": after_lines,
        "number_format": number_format_text,
        "hanging_indent_chars": int(hanging * 100),
    }
    issues: list[Issue] = []
    for entry in entries:
        para = _entry_paragraph(model, entry)
        if para is None:
            continue
        problems: list[str] = []
        problems.extend(
            run_font_problems(
                para, expected_size=size, east_asia=east_asia, ascii_font=ascii_font
            )
        )
        if not close(para.line_spacing, line_spacing, SPACING_TOL):
            actual = (
                f"{para.line_spacing:g} 倍"
                if para.line_spacing is not None
                else (
                    f"固定值 {para.line_spacing_pt:g}pt"
                    if para.line_spacing_pt is not None
                    else "未设置"
                )
            )
            problems.append(f"行距={actual}")
        before = _direct_lines(para, before=True)
        if before is not None and abs(before - before_lines) > LINE_TOL:
            problems.append(f"段前={before:.2f}行")
        after = _direct_lines(para, before=False)
        if after is not None and abs(after - after_lines) > LINE_TOL:
            problems.append(f"段后={after:.2f}行")
        text = (para.text or "").strip()
        if not para.has_num_pr and number_re.match(text) is None:
            problems.append(f"条目不是以 [n] 序号开头（规范正则 {number_format_text}）")
        hanging_actual = _hanging_chars(para)
        if hanging_actual is not None:
            if not close(hanging_actual, hanging, CHAR_TOL):
                problems.append(f"悬挂缩进={hanging_actual:.2f}字符（应为 {hanging:g}）")
        elif para.first_line_chars is not None or para.first_line_twips is not None:
            problems.append("用了首行缩进而非悬挂缩进")
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para, "refs/entry"),
                    message="参考文献条目格式不符合规范",
                    actual="；".join(dict.fromkeys(problems)),
                    expected=expected,
                    hint="条目：小四宋体/西文 Times New Roman、1.5 倍行距、段前后 0、[n] 序号、悬挂缩进 2 字符",
                )
            )
    return issues


# ---------------------------------------------------------------------------
# REF-003 正文引用标注与著录格式
# ---------------------------------------------------------------------------
def _check_citations(
    model: DocModel,
    roles: Roles,
    entries: list[ReferenceEntry],
    ruleset: RuleSet,
) -> list[Issue]:
    rule = _rule(ruleset, "REF-003")
    if rule is None:
        return []
    params = expected_params(ruleset, "REF-003", REF_CITATION_FALLBACK)
    expected = {
        "in_text_citation_superscript": True,
        "sequential_numbering": True,
        "style": _text(params, "style", "GB/T 7714-2015"),
        "categories": params.get("categories") or REF_CITATION_FALLBACK["categories"],
    }
    issues: list[Issue] = []

    # ---- 1) 正文引用一致性（复用 refs.citations）----
    try:
        findings = analyze_citations(model, roles, entries)
    except Exception:  # pragma: no cover - 防御：解析异常不应让检查整体崩溃
        findings = []
    for finding in findings:
        if "未被正文引用" in (finding.message or ""):
            # 内容完整性结论交给 tas refs 命令；检查器只标记引用本身的格式/编号错误
            continue
        issues.append(
            make_issue(
                rule,
                locator=f"refs/citation/{finding.locator}",
                message=finding.message,
                actual=finding.actual,
                expected=finding.expected or expected,
                hint="GB/T 7714：序号置于方括号内作上角标，按引证先后连续编号",
            )
        )

    # ---- 2) 著录要素缺失（轻量、可验证的部分）----
    for entry in entries:
        para = _entry_paragraph(model, entry)
        if para is None:
            continue
        problems: list[str] = []
        if not (entry.title or "").strip():
            problems.append("未识别到题名")
        if entry.doc_type is None:
            problems.append("未识别到文献类型标识（如 [J]/[M]/[C]/[D]）")
        if entry.year is None:
            problems.append("未识别到出版年份")
        if (
            entry.doc_type
            and entry.doc_type.upper().startswith(ELECTRONIC_TYPES)
            and not entry.url
            and not entry.doi
        ):
            problems.append("电子资源缺少 URL/DOI")
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para, "refs/entry"),
                    message="参考文献著录要素不完整",
                    actual="；".join(problems),
                    expected={"style": expected["style"], "categories": expected["categories"]},
                    hint="按 GB/T 7714-2015 补齐：主要责任者. 题名[文献类型标识]. 出版项",
                )
            )
    return issues
