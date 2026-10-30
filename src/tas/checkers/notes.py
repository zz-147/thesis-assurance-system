"""NOTE-001 脚注格式检查（P2）：直接解析 docx 包内的 ``word/footnotes.xml``。

背景：python-docx 不暴露脚注正文（footnotes part），因此这里用标准库 ``zipfile``
解包 + ``lxml`` 解析 OOXML；只读，不修改文档。

检查口径（有意保守，避免误报）：

- 只报告**显式设置且与规范不符**的格式：字号 ≠ 9pt、含中文 run 的 eastAsia ≠ 宋体、
  显式 ``w:lineRule="auto"`` 行距 ≠ 1.0；
- 未显式设置（继承页脚样式/文档默认）的项一律不报；
- 跳过 Word 内置分隔符/续分符/续分说明脚注（``w:type`` = separator /
  continuationSeparator / continuationNotice，或 ``w:id`` ≤ 0）与空文本脚注；
- NOTE-001 目前只存在于 ``profiles/school_llm_2026.rules.json``；基础规则库
  （``rules/rules.json``）没有该规则时直接返回 ``[]``。

产出统一 :class:`tas.report.Issue`：每个不合规脚注一条 warning，定位
``footnote/{id}``，severity 取 ``rule.severity``，expected 取 ``rule.expected``。
"""

from __future__ import annotations

import zipfile
from typing import Any, Optional

from lxml import etree

from ..config import Rule, RuleSet
from ..docx_io import DocModel
from ..report import Issue
from ..roles import Roles
from ..rule_params import expected_params
from .helpers import SPACING_TOL, has_cjk, make_issue

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS = {"w": W_NS}
_QN_ID = f"{{{W_NS}}}id"
_QN_TYPE = f"{{{W_NS}}}type"
_QN_VAL = f"{{{W_NS}}}val"
_QN_EAST_ASIA = f"{{{W_NS}}}eastAsia"
_QN_LINE = f"{{{W_NS}}}line"
_QN_LINE_RULE = f"{{{W_NS}}}lineRule"

SKIP_FOOTNOTE_TYPES = {
    "separator",
    "continuationseparator",
    "continuationnotice",
    "separatorcontinuation",
}

NOTE_FALLBACK = {
    "position": "page_footer",
    "size_pt": 9,
    "east_asia": "宋体",
    "line_spacing": 1.0,
}
SIZE_TOL = 0.3  # 半磅单位下 9pt vs 9.5pt 应能被检出


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


def _load_footnotes(model: DocModel) -> Optional[Any]:
    """读取 docx 内的 word/footnotes.xml；不存在/损坏时返回 None。"""
    path = getattr(model, "path", None)
    if path is None:
        return None
    try:
        with zipfile.ZipFile(path) as archive:
            try:
                data = archive.read("word/footnotes.xml")
            except KeyError:
                return None
    except (OSError, zipfile.BadZipFile):
        return None
    try:
        return etree.fromstring(data)
    except etree.XMLSyntaxError:
        return None


def _footnote_id(footnote: Any) -> Optional[int]:
    raw = footnote.get(_QN_ID)
    if raw is None:
        return None
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def _footnote_text(footnote: Any) -> str:
    return "".join(node.text or "" for node in footnote.findall(".//w:t", _NS))


def _run_text(run: Any) -> str:
    return "".join(node.text or "" for node in run.findall("w:t", _NS))


def _explicit_size_mismatch(run: Any, expected_size: float) -> Optional[float]:
    """run 显式设置 w:sz 且与期望不符 → 返回实际 pt；否则 None。"""
    rpr = run.find("w:rPr", _NS)
    if rpr is None:
        return None
    size = rpr.find("w:sz", _NS)
    if size is None:
        return None
    raw = size.get(_QN_VAL)
    try:
        actual = int(raw) / 2.0
    except (TypeError, ValueError):
        return None
    return actual if abs(actual - expected_size) > SIZE_TOL else None


def _explicit_font_mismatch(
    run: Any, text: str, expected_font: str
) -> Optional[str]:
    """含中文的 run 显式设置 eastAsia 且非期望字体 → 返回实际字体；否则 None。"""
    if not has_cjk(text):
        return None
    rpr = run.find("w:rPr", _NS)
    if rpr is None:
        return None
    fonts = rpr.find("w:rFonts", _NS)
    if fonts is None:
        return None
    east_asia = fonts.get(_QN_EAST_ASIA)
    if east_asia is None:
        return None
    return east_asia if east_asia != expected_font else None


def _explicit_line_spacing_mismatch(footnote: Any, expected: float) -> Optional[float]:
    """脚注段落里显式 auto 行距与期望不符 → 返回实际倍数；否则 None。"""
    for spacing in footnote.findall(".//w:p/w:pPr/w:spacing", _NS):
        line_rule = spacing.get(_QN_LINE_RULE) or "auto"
        if line_rule != "auto":
            # 固定值/最小值行距用 pt 表示，不按倍数比较（避免误报）
            continue
        raw = spacing.get(_QN_LINE)
        if raw is None:
            continue
        try:
            actual = int(raw) / 240.0
        except (TypeError, ValueError):
            continue
        if abs(actual - expected) > SPACING_TOL:
            return actual
    return None


def _footnote_problems(footnote: Any, params: dict) -> dict[str, Any]:
    """返回该脚注"显式设置且不符"的项（字号/中文字体/行距），无问题返回空 dict。"""
    expected_size = _num(params, "size_pt", NOTE_FALLBACK["size_pt"])
    expected_font = _text(params, "east_asia", NOTE_FALLBACK["east_asia"])
    expected_line = _num(params, "line_spacing", NOTE_FALLBACK["line_spacing"])

    actual: dict[str, Any] = {}
    for run in footnote.findall(".//w:r", _NS):
        size = _explicit_size_mismatch(run, expected_size)
        if size is not None:
            actual.setdefault("size_pt", size)
        font = _explicit_font_mismatch(run, _run_text(run), expected_font)
        if font is not None:
            actual.setdefault("east_asia", font)
    line_spacing = _explicit_line_spacing_mismatch(footnote, expected_line)
    if line_spacing is not None:
        actual.setdefault("line_spacing", line_spacing)
    return actual


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def check_notes(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng"
) -> list[Issue]:
    """检查脚注（NOTE-001）：字号 9pt、中文宋体、1 倍行距。

    - ruleset 无 NOTE-001（如基础规则库）→ ``[]``；
    - docx 内没有 ``word/footnotes.xml`` 或没有有效脚注 → ``[]``；
    - 每个不合规脚注一条 warning，locator ``footnote/{id}``；
    - expected 使用 ``rule.expected``（缺省时回退 NOTE_FALLBACK）。
    """
    rule = _rule(ruleset, "NOTE-001")
    if rule is None:
        return []
    params = expected_params(ruleset, "NOTE-001", NOTE_FALLBACK)
    root = _load_footnotes(model)
    if root is None:
        return []
    expected = dict(rule.expected) if isinstance(rule.expected, dict) and rule.expected else dict(
        NOTE_FALLBACK
    )

    issues: list[Issue] = []
    for footnote in root.findall("w:footnote", _NS):
        footnote_type = (footnote.get(_QN_TYPE) or "").strip().lower()
        if footnote_type in SKIP_FOOTNOTE_TYPES:
            continue
        footnote_id = _footnote_id(footnote)
        if footnote_id is None or footnote_id <= 0:
            continue
        if not _footnote_text(footnote).strip():
            continue
        problems = _footnote_problems(footnote, params)
        if not problems:
            continue
        issues.append(
            make_issue(
                rule,
                locator=f"footnote/{footnote_id}",
                message="脚注格式不符合规范",
                actual=problems,
                expected=expected,
                hint="脚注：注释放在当页页脚，小五号宋体、1 倍行距（Word：选中脚注文字设置字体/段落）",
            )
        )
    return issues
