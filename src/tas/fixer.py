"""修复引擎（学生端）：把检查出的格式问题安全修复成新 docx。

核心承诺（对应需求 FR-07 / FR-08）：
1. 只改格式：仅修改 OOXML 的 pPr / rPr / sectPr，绝不改任何文字、不增删段落；
2. 不改内容：修复前后对「全部段落 + 表格单元格 + 页眉页脚」文本做 sha256，
   不一致直接 RuntimeError（宁可失败，不可悄悄改坏正文）；
3. 不覆盖源文件：默认写 ``<原名>_fixed.docx``，保存后再校验源文件字节未变；
4. 修复后复检：重新 ``DocModel.load`` 输出文件 → 角色识别 → ``run_checks``，
   返回剩余问题（recheck_issues）；
5. 幂等：合规文档再修复时 actions 为空，不做无意义改写；
6. 参数化：纸张/页边距、正文字体/行距/缩进/对齐、标题格式、前置区（题目/摘要标题/
   摘要正文/关键词/目录标题/参考文献标题）的期望值统一从**当前 ruleset** 的
   ``expected`` 读取（``rule_params``），教师自定义 profile 改参数后修复立即生效；
   规则库缺字段时回退学校默认值，保证 ``school_default`` 的修复行为与参数化改造
   之前完全一致。

修复范围由 :data:`SAFE_RULE_IDS` 白名单限定（页面尺寸/边距、正文格式、标题/题目/
摘要/关键词/目录标题/参考文献标题的字体字号与段落属性、表格表内文字格式）。编号连续性、
字符间距、参考文献著录等涉及文字或结构的问题只进 ``skipped``，reason 说明“需人工/涉及文字”。
"""

from __future__ import annotations

import hashlib
import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator, Optional

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .checkers import run_checks
from .checkers.helpers import (
    CHAR_TOL,
    CM_TOL,
    PT_TOL,
    SPACING_TOL,
    close,
    has_cjk,
    has_latin,
    indent_chars,
    is_justified,
    loc_para,
    para_format_problems,
    run_font_problems,
)
from .config import DEFAULT_PROFILE, RuleSet
from .docx_io import TWIPS_PER_CM, DocModel, ParaInfo
from .report import Issue
from .roles import Roles, classify
from .rule_params import (
    alignment,
    char_count,
    expected_params,
    font_size,
    line_count,
    text_value,
)

__all__ = [
    "SAFE_RULE_IDS",
    "SKIPPED_ROLES",
    "FixAction",
    "FixResult",
    "content_hash",
    "fix_document",
]

# ---------------------------------------------------------------------------
# 白名单与期望值常量
# ---------------------------------------------------------------------------
#: 允许自动修复的规则（全部是“只改格式、不改文字”的安全项）。
SAFE_RULE_IDS: tuple[str, ...] = (
    "PAGE-001",
    "PAGE-002",
    "BODY-001",
    "BODY-002",
    "BODY-003",
    "BODY-004",
    "HEAD-001",
    "HEAD-002",
    "HEAD-003",
    "HEAD-007",
    "TITLE-001",
    "TITLE-002",
    "ABS-001",
    "ABS-002",
    "ABS-003",
    "ABS-004",
    "KW-001",
    "KW-002",
    "TOC-001",
    "REF-001",
    "TAB-002",
)
_SAFE_RULE_ID_SET = frozenset(SAFE_RULE_IDS)

#: 这些角色的段落一律不自动修改：目录条目（TOC 域/页码）与参考文献条目（著录内容）。
SKIPPED_ROLES: tuple[str, ...] = ("reference_entry", "toc_entry")

# 以下是「默认学期规范」兜底值：规则库（rules.json / 教师自定义 profile）缺字段时使用，
# 保证 school_default 的修复行为与参数化改造之前完全一致。实际期望值统一由
# ``rule_params.expected_params`` 从当前 ruleset 读取（见 _page_*_expected / _body_expected_dict
# / _heading_specs / _head007_expected）。
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
PAGE_SIZE_FALLBACK = {
    "page_width_cm": PAGE_WIDTH_CM,
    "page_height_cm": PAGE_HEIGHT_CM,
    "orientation": "portrait",
}
PAGE_MARGINS_EXPECTED = {
    "top_cm": 2.54,
    "bottom_cm": 2.54,
    "left_cm": 2.0,
    "right_cm": 2.0,
    "gutter_cm": 1.0,
    "header_distance_cm": 2.0,
    "footer_distance_cm": 2.0,
}
#: PAGE-002 的 expected 键名（margin_*_cm）与 fixer 内部键名的对应关系
PAGE_MARGIN_KEYS = (
    ("top_cm", "margin_top_cm"),
    ("bottom_cm", "margin_bottom_cm"),
    ("left_cm", "margin_left_cm"),
    ("right_cm", "margin_right_cm"),
    ("gutter_cm", "gutter_cm"),
    ("header_distance_cm", "header_distance_cm"),
    ("footer_distance_cm", "footer_distance_cm"),
)
PAGE_MARGINS_FALLBACK = {
    expected_key: PAGE_MARGINS_EXPECTED[internal_key]
    for internal_key, expected_key in PAGE_MARGIN_KEYS
}

BODY_SIZE_PT = 12.0
BODY_EAST_ASIA = "宋体"
BODY_ASCII = "Times New Roman"
BODY_LINE_SPACING = 1.5
BODY_FIRST_LINE_CHARS = 2.0
BODY_ALIGNMENT = "both"

#: HEAD-001/002/003 的默认值（= 原 LEVEL_SPEC；段前/段后统一用「行」为内部单位）
HEADING_FALLBACK = {
    1: {
        "size_pt": 16.0,
        "east_asia": "黑体",
        "alignment": "center",
        "line_spacing": 1.5,
        "before_lines": 1.0,
        "after_lines": 1.0,
    },
    2: {
        "size_pt": 14.0,
        "east_asia": "黑体",
        "alignment": "left",
        "line_spacing": 1.5,
        "before_lines": 0.5,
        "after_lines": 0.5,
    },
    3: {
        "size_pt": 12.0,
        "east_asia": "黑体",
        "alignment": "left",
        "line_spacing": 1.5,
        "before_lines": 0.5,
        "after_lines": 0.0,
    },
}

#: HEAD-007（前言/结论/致谢）默认值。注意 rules.json 的 HEAD-007.expected 没有 after_lines，
#: 段后距按学科兜底：sci_eng=1.0 行、humanities=0.0 行（与检查器/旧写死行为一致）。
HEAD007_FALLBACK = {
    "size_pt": 16.0,
    "east_asia": "黑体",
    "alignment": {"sci_eng": "center", "humanities": "left"},
    "line_spacing": 1.5,
    "before_lines": 1.0,
}

#: 前置区默认值（= 参数化改造前的写死值）。实际期望值统一由
#: ``rule_params.expected_params`` 从当前 ruleset 读取；rules.json 缺字段时才用这里兜底。
#: 段前/段后为内部单位「行」（rules.json 的 1/100 行经 line_count() 换算后覆盖）。
TITLE_CN_FALLBACK = {
    "size_pt": 16.0,
    "east_asia": "黑体",
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 1.0,
    "after_lines": 1.0,
}
TITLE_EN_FALLBACK = {
    "size_pt": 16.0,
    "ascii": "Times New Roman",
    "bold": True,
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 1.0,
    "after_lines": 1.0,
}
ABS_CN_LABEL_FALLBACK = {
    "size_pt": 16.0,
    "east_asia": "黑体",
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 1.0,
    "after_lines": 1.0,
}
ABS_EN_LABEL_FALLBACK = {
    "size_pt": 16.0,
    "ascii": "Times New Roman",
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 1.0,
    "after_lines": 1.0,
}
ABS_CN_BODY_FALLBACK = {
    "size_pt": 12.0,
    "east_asia": "宋体",
    "ascii": "Times New Roman",
    "first_line_chars": 2.0,
    "line_spacing": 1.5,
    "alignment": "both",
}
ABS_EN_BODY_FALLBACK = {
    "size_pt": 12.0,
    "east_asia": "宋体",
    "ascii": "Times New Roman",
    "first_line_chars": 2.0,
    "line_spacing": 1.5,
    "alignment": "both",
}
TOC_TITLE_FALLBACK = {
    "title_size_pt": 16.0,
    "title_east_asia": "黑体",
    "title_alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 1.0,
    "after_lines": 1.0,
}
REF_TITLE_FALLBACK = {
    "size_pt": 16.0,
    "east_asia": "黑体",
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 1.0,
    "after_lines": 1.0,
}
KW_CN_FALLBACK = {
    "label_size_pt": 12.0,
    "label_east_asia": "黑体",
    "body_size_pt": 12.0,
    "body_east_asia": "宋体",
    "body_ascii": "Times New Roman",
}
KW_EN_FALLBACK = {
    "label_size_pt": 12.0,
    "label_ascii": "Times New Roman",
    "label_bold": True,
    "body_size_pt": 12.0,
    "body_ascii": "Times New Roman",
}

#: TAB-002「表内文字」默认值：表内文字五号（10.5pt）宋体、居中、1.5 倍行距。
#: 西文字体 rules.json 只写在 TAB-002 的 caption 字段里（number_ascii），表内文字按
#: 检查器口径不校验西文；修复端统一设为 Times New Roman 做安全归一化。
TAB002_FALLBACK = {
    "cell_size_pt": 10.5,
    "cell_east_asia": "宋体",
    "cell_alignment": "center",
    "cell_line_spacing": 1.5,
}
TABLE_ASCII_FONT = "Times New Roman"

_ALIGN_MAP = {
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}

#: 对齐方式 → 报告用中文标签（默认值与参数化之前保持一致）
_ALIGNMENT_LABELS = {
    "both": "两端对齐",
    "justify": "两端对齐",
    "left": "左对齐",
    "center": "居中",
    "right": "右对齐",
}


# ---------------------------------------------------------------------------
# 规则参数读取（当前 ruleset 的 expected；缺字段回退默认学期规范）
# ---------------------------------------------------------------------------
def _float_value(params: dict, key: str, fallback: float) -> float:
    """数值参数：复用 ``rule_params.font_size`` 的容错解析（缺字段/非数值回退）。"""
    return font_size(params, key, fallback)


def _profile_alignment(value: Any, profile: str, fallback: str) -> str:
    """HEAD-007 的 alignment 可能是 ``{"sci_eng": "center", "humanities": "left"}`` 字典。

    按当前 profile 取值；字典里没有当前学科时按 "center" 兜底（与
    ``checkers.heading._special_titles`` 的语义保持一致）。
    """
    if isinstance(value, dict):
        return alignment(
            {"alignment": value.get(profile, "center")}, "alignment", fallback
        )
    return alignment({"alignment": value}, "alignment", fallback)


def _optional_text(params: dict, key: str, fallback: Any) -> Optional[str]:
    """可选文本参数：缺字段/None 时回退；空串视为“未设置”。"""
    value = params.get(key, fallback)
    if value is None:
        return None
    text = str(value).strip()
    if text:
        return text
    return None if fallback is None else str(fallback).strip()


def _optional_bool(params: dict, key: str, fallback: Any) -> Optional[bool]:
    """可选布尔参数（rules.json 里可能是 true/false 或 "true"/"false"）。"""
    value = params.get(key, fallback)
    if value is None:
        return fallback
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("true", "1", "yes", "是")


def _alignment_matches(para: ParaInfo, expected: str) -> bool:
    """段落对齐是否符合期望（justify / both 均由 ``is_justified`` 判定）。"""
    if expected in ("both", "justify"):
        return is_justified(para)
    return (para.alignment or "left") == expected


# ---------------------------------------------------------------------------
# 结果模型
# ---------------------------------------------------------------------------
@dataclass
class FixAction:
    """一条实际发生的修复动作（只记录真的改了属性的项）。"""

    rule_id: str
    locator: str
    description: str
    before: Optional[Any] = None
    after: Optional[Any] = None

    def to_dict(self) -> dict:
        return {
            "rule_id": self.rule_id,
            "locator": self.locator,
            "description": self.description,
            "before": self.before,
            "after": self.after,
        }


@dataclass
class FixResult:
    """一次修复的完整结果（to_dict() 可直接 JSON 序列化）。"""

    src_path: Path
    out_path: Optional[Path]
    dry_run: bool
    actions: list[FixAction] = field(default_factory=list)
    skipped: list[dict] = field(default_factory=list)
    content_hash_before: str = ""
    content_hash_after: str = ""
    recheck_issues: list[Issue] = field(default_factory=list)
    baseline_counts: dict = field(default_factory=dict)

    # -- 便捷属性 ---------------------------------------------------------
    @property
    def content_unchanged(self) -> bool:
        return bool(self.content_hash_before) and (
            self.content_hash_before == self.content_hash_after
        )

    def to_dict(self) -> dict:
        return {
            "src_path": str(self.src_path),
            "out_path": str(self.out_path) if self.out_path is not None else None,
            "dry_run": self.dry_run,
            "actions": [action.to_dict() for action in self.actions],
            "skipped": [dict(item) for item in self.skipped],
            "content_hash_before": self.content_hash_before,
            "content_hash_after": self.content_hash_after,
            "content_hash_match": self.content_unchanged,
            "baseline_counts": dict(self.baseline_counts),
            "recheck_counts": _count_issues(self.recheck_issues),
            "recheck_issues": [issue.to_dict() for issue in self.recheck_issues],
            "action_count": len(self.actions),
            "skipped_count": len(self.skipped),
        }

    def render_text(self, max_actions: int = 40) -> str:
        """文本摘要（给 CLI / 终端预览用）。"""
        counts = self.baseline_counts or {}
        recheck = _count_issues(self.recheck_issues)
        lines = [
            "tas 安全修复结果" + ("（dry-run，未写文件）" if self.dry_run else ""),
            "=" * 64,
            f"源文件  ：{self.src_path}",
            f"输出文件：{self.out_path if self.out_path is not None else '（dry-run 不写入）'}",
            (
                f"修复前  ：错误 {counts.get('error', 0)} ｜ 警告 {counts.get('warning', 0)} ｜ "
                f"提示 {counts.get('info', 0)}"
            ),
            (
                f"修复后  ：错误 {recheck['error']} ｜ 警告 {recheck['warning']} ｜ "
                f"提示 {recheck['info']}"
            ),
            f"动作    ：{len(self.actions)} 项；跳过/需人工 {len(self.skipped)} 项",
            f"内容 hash：{'一致（正文/表格/页眉页脚未变）' if self.content_unchanged else '不一致！'}",
            "",
        ]
        for index, action in enumerate(self.actions[:max_actions], 1):
            lines.append(f"{index}. [{action.rule_id}] {action.locator}")
            lines.append(f"   {action.description}")
            if action.before is not None or action.after is not None:
                lines.append(f"   前：{_fmt(action.before)} → 后：{_fmt(action.after)}")
        if len(self.actions) > max_actions:
            lines.append(f"…… 其余 {len(self.actions) - max_actions} 项已省略")
        if self.skipped:
            lines.append("")
            lines.append("跳过 / 需人工：")
            for item in self.skipped:
                lines.append(
                    f"- {item.get('rule_id', item.get('role', '?'))}：{item.get('reason', '')}"
                    f"（{item.get('locator', item.get('label', ''))}）"
                )
        return "\n".join(lines)


def _fmt(value: Any) -> str:
    """把 before/after 的可读化展示（与 report._fmt 类似，避免跨模块私有依赖）。"""
    if value is None:
        return "（无）"
    if isinstance(value, dict):
        return "，".join(f"{key}={_fmt(item)}" for key, item in value.items())
    if isinstance(value, (list, tuple)):
        return "[" + "，".join(_fmt(item) for item in value) + "]"
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def _r(value: Optional[float]) -> Optional[float]:
    return None if value is None else round(value, 2)


def _cm_text(value: Optional[float]) -> str:
    return "未设置" if value is None else f"{value:.2f}cm"


def _count_issues(issues: list[Issue]) -> dict:
    counts = {"error": 0, "warning": 0, "info": 0}
    by_rule: dict[str, int] = {}
    for issue in issues:
        counts[issue.severity] = counts.get(issue.severity, 0) + 1
        by_rule[issue.rule_id] = by_rule.get(issue.rule_id, 0) + 1
    counts["total"] = len(issues)
    counts["by_rule"] = dict(sorted(by_rule.items()))
    return counts


# ---------------------------------------------------------------------------
# 内容 hash：段落 + 表格单元格 + 页眉页脚
# ---------------------------------------------------------------------------
def _iter_table_texts(table: Any) -> Iterator[str]:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                yield paragraph.text or ""
            for nested in cell.tables:
                yield from _iter_table_texts(nested)


def _iter_story_texts(story: Any) -> Iterator[str]:
    for paragraph in story.paragraphs:
        yield paragraph.text or ""
    for table in story.tables:
        yield from _iter_table_texts(table)


_HEADER_FOOTER_ATTRS = (
    "header",
    "footer",
    "first_page_header",
    "first_page_footer",
    "even_page_header",
    "even_page_footer",
)


def _document_texts(document: Any) -> list[str]:
    texts = list(_iter_story_texts(document))
    for section in document.sections:
        for attr in _HEADER_FOOTER_ATTRS:
            try:
                container = getattr(section, attr)
                texts.extend(_iter_story_texts(container))
            except Exception:  # pragma: no cover - 损坏的 header/footer part 兜底
                continue
    return texts


def content_hash(model: DocModel) -> str:
    """内容指纹：全部段落 text + 表格单元格 text + 页眉/页脚 text 的 sha256。

    注意：这里只取“文字内容”，字体/字号/行距等格式属性不参与，
    因此“只改格式”的修复必须前后一致；一旦有文字被增删改，hash 必然变化。
    """
    path = Path(model.path) if model.path is not None else None
    if path is not None and path.is_file():
        texts = _document_texts(docx.Document(str(path)))
    else:  # pragma: no cover - 内存模型的兜底路径（正常流程不会用到）
        texts = [para.text or "" for para in model.paragraphs]
        for section in model.sections:
            texts.append(section.header_text or "")
            texts.append(section.footer_text or "")
    digest = hashlib.sha256()
    for text in texts:
        data = (text or "").encode("utf-8")
        digest.update(len(data).to_bytes(8, "big"))
        digest.update(data)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# 底层 OOXML 写入工具（只碰 pPr / rPr / sectPr）
# ---------------------------------------------------------------------------
def _apply_font_to_run(
    run: Any,
    text: str,
    *,
    size_pt: Optional[float] = None,
    east_asia: Optional[str] = None,
    ascii_font: Optional[str] = None,
    force_ascii: bool = False,
    bold: Optional[bool] = None,
) -> None:
    """改一个 run 的字体/字号/加粗（w:rPr），不碰文字。"""
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    rpr = run._r.get_or_add_rPr()
    if east_asia is not None and has_cjk(text):
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:eastAsia"), east_asia)
    if ascii_font is not None and (force_ascii or has_latin(text)):
        run.font.name = ascii_font  # 写 w:rFonts/@w:ascii 与 @w:hAnsi
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:cs"), ascii_font)
    if bold is not None:
        run.font.bold = bool(bold)


def _apply_run_fonts(
    paragraph: Any,
    *,
    size_pt: Optional[float] = None,
    east_asia: Optional[str] = None,
    ascii_font: Optional[str] = None,
    force_ascii: bool = False,
    bold: Optional[bool] = None,
) -> None:
    for run in paragraph.runs:
        text = run.text or ""
        if not text.strip():
            continue
        _apply_font_to_run(
            run,
            text,
            size_pt=size_pt,
            east_asia=east_asia,
            ascii_font=ascii_font,
            force_ascii=force_ascii,
            bold=bold,
        )


def _get_or_add_spacing(paragraph: Any) -> Any:
    ppr = paragraph._p.get_or_add_pPr()
    return ppr.get_or_add_spacing()


def _set_line_spacing(paragraph: Any, multiple: float) -> None:
    """行距（倍数）：w:spacing/@w:line=倍数*240，lineRule=auto。"""
    paragraph.paragraph_format.line_spacing = float(multiple)


def _set_before_after_lines(
    paragraph: Any,
    *,
    before_lines: Optional[float] = None,
    after_lines: Optional[float] = None,
) -> None:
    """段前/段后用 Word 原生的 w:beforeLines / w:afterLines（1/100 行）。"""
    if before_lines is None and after_lines is None:
        return
    spacing = _get_or_add_spacing(paragraph)
    if before_lines is not None:
        spacing.set(qn("w:beforeLines"), str(int(round(before_lines * 100))))
        for attr in ("before", "beforeAutospacing"):
            key = qn(f"w:{attr}")
            if spacing.get(key) is not None:
                del spacing.attrib[key]
    if after_lines is not None:
        spacing.set(qn("w:afterLines"), str(int(round(after_lines * 100))))
        for attr in ("after", "afterAutospacing"):
            key = qn(f"w:{attr}")
            if spacing.get(key) is not None:
                del spacing.attrib[key]


def _set_first_line_indent(paragraph: Any, chars: float = 2.0) -> None:
    """首行缩进：w:ind/@w:firstLineChars=200，并删掉冲突的 firstLine/hanging。"""
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.get_or_add_ind()
    ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
    for attr in ("firstLine", "hanging", "hangingChars"):
        key = qn(f"w:{attr}")
        if ind.get(key) is not None:
            del ind.attrib[key]


def _set_alignment(paragraph: Any, value: str) -> None:
    paragraph.alignment = _ALIGN_MAP[value]


# ---------------------------------------------------------------------------
# 页面修复（sectPr）：参数来自当前 ruleset 的 PAGE-001 / PAGE-002 expected
# ---------------------------------------------------------------------------
def _page_size_expected(ruleset: RuleSet) -> dict[str, Any]:
    """PAGE-001 页面尺寸期望（默认：A4 21.0cm × 29.7cm，纵向）。"""
    params = expected_params(ruleset, "PAGE-001", PAGE_SIZE_FALLBACK)
    return {
        "width_cm": _float_value(params, "page_width_cm", PAGE_WIDTH_CM),
        "height_cm": _float_value(params, "page_height_cm", PAGE_HEIGHT_CM),
        "orientation": text_value(params, "orientation", "portrait").strip().lower(),
    }


def _page_size_label(expected: dict[str, Any]) -> str:
    """页面尺寸的可读标签：默认 A4 保持原描述，自定义尺寸按实际值展示。"""
    width, height = expected["width_cm"], expected["height_cm"]
    label = f"{width:.1f}cm × {height:.1f}cm"
    if abs(width - PAGE_WIDTH_CM) <= 1e-6 and abs(height - PAGE_HEIGHT_CM) <= 1e-6:
        return f"A4（{label}）"
    return label


def _page_margins_expected(ruleset: RuleSet) -> dict[str, float]:
    """PAGE-002 页边距/装订线/页眉页脚距边界期望（键名转为 fixer 内部键名）。"""
    params = expected_params(ruleset, "PAGE-002", PAGE_MARGINS_FALLBACK)
    return {
        internal_key: _float_value(
            params, expected_key, PAGE_MARGINS_EXPECTED[internal_key]
        )
        for internal_key, expected_key in PAGE_MARGIN_KEYS
    }


def _fix_pages(model: DocModel, document: Any, ruleset: RuleSet) -> list[FixAction]:
    actions: list[FixAction] = []
    size_expected = _page_size_expected(ruleset)
    margins_expected = _page_margins_expected(ruleset)
    for info, section in zip(model.sections, document.sections):
        # PAGE-001 纸张尺寸（默认 A4；自定义 profile 可改）
        size_bad: dict[str, Optional[float]] = {}
        if not close(info.page_width_cm, size_expected["width_cm"], CM_TOL):
            size_bad["width_cm"] = info.page_width_cm
        if not close(info.page_height_cm, size_expected["height_cm"], CM_TOL):
            size_bad["height_cm"] = info.page_height_cm
        if size_bad:
            before = {key: _r(value) for key, value in size_bad.items()}
            section.page_width = Cm(size_expected["width_cm"])
            section.page_height = Cm(size_expected["height_cm"])
            pg_sz = section._sectPr.find(qn("w:pgSz"))
            if pg_sz is not None:
                orientation = size_expected["orientation"]
                if orientation not in ("portrait", "landscape"):
                    # 规则库没写 orientation 时按宽高推断，避免横放文档被误改
                    orientation = (
                        "landscape"
                        if size_expected["width_cm"] > size_expected["height_cm"]
                        else "portrait"
                    )
                current_orient = pg_sz.get(qn("w:orient"))
                if orientation == "portrait":
                    # 与参数化之前一致：只在当前标记为横向时改回纵向（不新增冗余属性）
                    if current_orient == "landscape":
                        pg_sz.set(qn("w:orient"), "portrait")
                elif current_orient != "landscape":
                    pg_sz.set(qn("w:orient"), "landscape")
            after = {key: _r(size_expected[key]) for key in size_bad}
            actions.append(
                FixAction(
                    rule_id="PAGE-001",
                    locator=f"section/{info.index}",
                    description=(
                        f"页面尺寸归一化为 {_page_size_label(size_expected)}："
                        + "；".join(
                            f"{'宽度' if key == 'width_cm' else '高度'}"
                            f" {_cm_text(value)} → "
                            f"{size_expected[key]:.2f}cm"
                            for key, value in before.items()
                        )
                    ),
                    before=before,
                    after=after,
                )
            )

        # PAGE-002 页边距 / 装订线 / 页眉页脚距边界
        actual = {
            "top_cm": info.margin_top_cm,
            "bottom_cm": info.margin_bottom_cm,
            "left_cm": info.margin_left_cm,
            "right_cm": info.margin_right_cm,
            "gutter_cm": info.gutter_cm,
            "header_distance_cm": info.header_distance_cm,
            "footer_distance_cm": info.footer_distance_cm,
        }
        margins_bad = {
            key: _r(value)
            for key, value in actual.items()
            if not close(value, margins_expected[key], CM_TOL)
        }
        if margins_bad:
            section.top_margin = Cm(margins_expected["top_cm"])
            section.bottom_margin = Cm(margins_expected["bottom_cm"])
            section.left_margin = Cm(margins_expected["left_cm"])
            section.right_margin = Cm(margins_expected["right_cm"])
            section.header_distance = Cm(margins_expected["header_distance_cm"])
            section.footer_distance = Cm(margins_expected["footer_distance_cm"])
            pg_mar = section._sectPr.get_or_add_pgMar()
            pg_mar.set(
                qn("w:gutter"),
                str(int(round(margins_expected["gutter_cm"] * TWIPS_PER_CM))),
            )
            after = {key: margins_expected[key] for key in margins_bad}
            actions.append(
                FixAction(
                    rule_id="PAGE-002",
                    locator=f"section/{info.index}",
                    description=f"页边距/装订线 {margins_expected['gutter_cm']:g}cm/"
                    "页眉页脚距边界归一化为规范值："
                    + "；".join(
                        f"{key} {_fmt(value)} → {margins_expected[key]:g}"
                        for key, value in margins_bad.items()
                    ),
                    before=margins_bad,
                    after=after,
                )
            )
    return actions


# ---------------------------------------------------------------------------
# 正文修复（BODY-001..004 + ABS-002/004）
# ---------------------------------------------------------------------------
def _spacing_text(para: ParaInfo) -> str:
    if para.line_spacing is not None:
        return f"行距={para.line_spacing:g} 倍"
    if para.line_spacing_pt is not None:
        return f"行距=固定值 {para.line_spacing_pt:g}pt"
    return "行距=未设置（继承默认）"


def _indent_text(para: ParaInfo) -> str:
    chars = indent_chars(para)
    if chars is None:
        return "首行缩进=未设置（或为悬挂缩进）"
    return f"首行缩进={chars:.2f}字符"


def _body_expected_dict(ruleset: RuleSet) -> dict:
    """BODY-001..004 的期望值。

    默认 = 小四宋体/TNR + 1.5 倍行距 + 首行缩进 2 字符 + 两端对齐（旧写死值）；
    自定义 profile 改 expected 后，正文归一化随之改变。
    """
    font_params = expected_params(
        ruleset,
        "BODY-001",
        {"size_pt": BODY_SIZE_PT, "east_asia": BODY_EAST_ASIA, "ascii": BODY_ASCII},
    )
    spacing_params = expected_params(
        ruleset, "BODY-002", {"line_spacing": BODY_LINE_SPACING}
    )
    indent_params = expected_params(
        ruleset, "BODY-003", {"first_line_chars": 200}
    )
    align_params = expected_params(
        ruleset, "BODY-004", {"alignment": BODY_ALIGNMENT}
    )
    align = alignment(align_params, "alignment", BODY_ALIGNMENT)
    if align in ("justify", "both"):
        # 检查器把 justify / both 都视为「两端对齐」；内部统一为 both，
        # 保证默认 ruleset 的 after 描述与复检结果和参数化之前一致。
        align = "both"
    return {
        "size_pt": font_size(font_params, "size_pt", BODY_SIZE_PT),
        "east_asia": text_value(font_params, "east_asia", BODY_EAST_ASIA),
        "ascii": text_value(font_params, "ascii", BODY_ASCII),
        "line_spacing": _float_value(
            spacing_params, "line_spacing", BODY_LINE_SPACING
        ),
        "first_line_chars": char_count(
            indent_params.get("first_line_chars"), BODY_FIRST_LINE_CHARS
        ),
        "alignment": align,
    }


def _abstract_body_expected(
    ruleset: RuleSet, rule_id: str, fallback: dict[str, Any]
) -> dict:
    """ABS-002 / ABS-004（摘要正文）期望值：来自各自规则的 expected（缺字段回退默认值）。"""
    params = expected_params(ruleset, rule_id, fallback)
    align = alignment(params, "alignment", str(fallback["alignment"]))
    if align in ("justify", "both"):
        # 检查器把 justify / both 都视为「两端对齐」；内部统一为 both，
        # 保证默认 ruleset 的动作描述/复检结果与参数化之前一致。
        align = "both"
    return {
        "size_pt": font_size(params, "size_pt", float(fallback["size_pt"])),
        "east_asia": text_value(params, "east_asia", fallback["east_asia"]),
        "ascii": text_value(params, "ascii", fallback["ascii"]),
        "line_spacing": _float_value(
            params, "line_spacing", float(fallback["line_spacing"])
        ),
        "first_line_chars": char_count(
            params.get("first_line_chars"), float(fallback["first_line_chars"])
        ),
        "alignment": align,
    }


def _abstract_body_problems(para: ParaInfo, expected: dict) -> list[str]:
    """摘要正文检查：与 ``checkers.front`` 的判定/文案保持一致（参数化版本）。"""
    problems = run_font_problems(
        para,
        expected_size=expected["size_pt"],
        east_asia=expected["east_asia"],
        ascii_font=expected["ascii"],
    )
    if not close(para.line_spacing, expected["line_spacing"], SPACING_TOL):
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
    chars = indent_chars(para)
    if not close(chars, expected["first_line_chars"], CHAR_TOL):
        problems.append(f"首行缩进={chars:.2f}字符" if chars is not None else "首行缩进未设置")
    wanted = expected["alignment"]
    if wanted in ("both", "justify"):
        aligned = is_justified(para)
    else:
        aligned = (para.alignment or "left") == wanted
    if not aligned:
        problems.append(f"对齐={para.alignment or 'left'}")
    return problems


def _apply_body_normalization(paragraph: Any, expected: dict) -> None:
    """正文类段落整体归一化：字体/字号、行距、首行缩进、对齐（参数来自当前 ruleset）。"""
    _apply_run_fonts(
        paragraph,
        size_pt=expected["size_pt"],
        east_asia=expected["east_asia"],
        ascii_font=expected["ascii"],
    )
    _set_line_spacing(paragraph, expected["line_spacing"])
    _set_first_line_indent(paragraph, expected["first_line_chars"])
    _set_alignment(paragraph, expected["alignment"])


def _fix_body(
    model: DocModel, roles: Roles, document: Any, ruleset: RuleSet
) -> list[FixAction]:
    actions: list[FixAction] = []
    paragraphs = document.paragraphs
    body_expected = _body_expected_dict(ruleset)

    # 普通正文：BODY-001 字体 / BODY-002 行距 / BODY-003 缩进 / BODY-004 对齐
    for index in roles.indices("body"):
        para = model.by_index.get(index)
        if para is None or para.is_empty or index >= len(paragraphs):
            continue
        target = paragraphs[index]

        font_problems = run_font_problems(
            para,
            expected_size=body_expected["size_pt"],
            east_asia=body_expected["east_asia"],
            ascii_font=body_expected["ascii"],
        )
        if font_problems:
            _apply_run_fonts(
                target,
                size_pt=body_expected["size_pt"],
                east_asia=body_expected["east_asia"],
                ascii_font=body_expected["ascii"],
            )
            detail = "；".join(dict.fromkeys(font_problems))
            actions.append(
                FixAction(
                    rule_id="BODY-001",
                    locator=loc_para(para),
                    description=(
                        "正文字体/字号归一化为"
                        f" {body_expected['size_pt']:g}pt "
                        f"{body_expected['east_asia']}+{body_expected['ascii']}：{detail}"
                    ),
                    before=detail,
                    after={
                        "size_pt": body_expected["size_pt"],
                        "east_asia": body_expected["east_asia"],
                        "ascii": body_expected["ascii"],
                    },
                )
            )

        if not close(
            para.line_spacing, body_expected["line_spacing"], SPACING_TOL
        ):
            before = _spacing_text(para)
            _set_line_spacing(target, body_expected["line_spacing"])
            actions.append(
                FixAction(
                    rule_id="BODY-002",
                    locator=loc_para(para),
                    description=(
                        "正文行距归一化为 "
                        f"{body_expected['line_spacing']:g} 倍：{before}"
                    ),
                    before=before,
                    after={"line_spacing": body_expected["line_spacing"]},
                )
            )

        if not close(
            indent_chars(para), body_expected["first_line_chars"], CHAR_TOL
        ):
            before = _indent_text(para)
            _set_first_line_indent(target, body_expected["first_line_chars"])
            actions.append(
                FixAction(
                    rule_id="BODY-003",
                    locator=loc_para(para),
                    description=(
                        "正文首行缩进归一化为 "
                        f"{body_expected['first_line_chars']:g} 字符：{before}"
                    ),
                    before=before,
                    after={"first_line_chars": body_expected["first_line_chars"]},
                )
            )

        if not _alignment_matches(para, body_expected["alignment"]):
            before = para.alignment or "left（未设置，默认左对齐）"
            _set_alignment(target, body_expected["alignment"])
            label = _ALIGNMENT_LABELS.get(
                body_expected["alignment"], body_expected["alignment"]
            )
            actions.append(
                FixAction(
                    rule_id="BODY-004",
                    locator=loc_para(para),
                    description=f"正文对齐归一化为{label}：{before}",
                    before=before,
                    after={"alignment": body_expected["alignment"]},
                )
            )

    # 摘要正文：ABS-002（中文）/ ABS-004（英文），一条规则覆盖全部格式项；
    # 期望值来自各自规则的 expected（不再借用 BODY 参数），缺字段回退默认值。
    for role, rule_id, prefix, fallback in (
        ("abstract_body_cn", "ABS-002", "front/abstract_cn", ABS_CN_BODY_FALLBACK),
        ("abstract_body_en", "ABS-004", "front/abstract_en", ABS_EN_BODY_FALLBACK),
    ):
        abstract_expected = _abstract_body_expected(ruleset, rule_id, fallback)
        for index in roles.indices(role):
            para = model.by_index.get(index)
            if para is None or para.is_empty or index >= len(paragraphs):
                continue
            problems = _abstract_body_problems(para, abstract_expected)
            if not problems:
                continue
            _apply_body_normalization(paragraphs[index], abstract_expected)
            detail = "；".join(dict.fromkeys(problems))
            actions.append(
                FixAction(
                    rule_id=rule_id,
                    locator=loc_para(para, prefix),
                    description=f"摘要正文格式归一化：{detail}",
                    before=detail,
                    after=dict(abstract_expected),
                )
            )
    return actions


# ---------------------------------------------------------------------------
# 标题 / 题名 / 标签类：按段落 spec 整体归一化
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class _ParagraphSpec:
    """一个“段落整体格式”的期望值（仅格式，不含文字内容）。"""

    rule_id: str
    roles: tuple[str, ...]
    locator_prefix: str
    description: str
    size_pt: float
    east_asia: Optional[str] = None
    ascii_font: Optional[str] = None
    force_ascii: bool = False
    bold: Optional[bool] = None
    alignment: Optional[str] = None
    line_spacing: Optional[float] = 1.5
    before_lines: Optional[float] = 1.0
    after_lines: Optional[float] = 1.0


def _heading_specs(ruleset: RuleSet) -> list[_ParagraphSpec]:
    """HEAD-001/002/003 期望值：来自当前 ruleset 的 expected（缺字段回退默认学校规范）。"""
    specs: list[_ParagraphSpec] = []
    for level in (1, 2, 3):
        fallback = HEADING_FALLBACK[level]
        params = expected_params(ruleset, f"HEAD-00{level}", fallback)
        specs.append(
            _ParagraphSpec(
                rule_id=f"HEAD-00{level}",
                roles=(f"heading_{level}",),
                locator_prefix="body/para",
                description=f"{level} 级标题格式",
                size_pt=font_size(params, "size_pt", fallback["size_pt"]),
                east_asia=text_value(params, "east_asia", fallback["east_asia"]),
                alignment=alignment(
                    params, "alignment", fallback["alignment"]
                ),
                line_spacing=_float_value(
                    params, "line_spacing", fallback["line_spacing"]
                ),
                before_lines=line_count(
                    params.get("before_lines"), fallback["before_lines"]
                ),
                after_lines=line_count(
                    params.get("after_lines"), fallback["after_lines"]
                ),
            )
        )
    return specs


def _head007_expected(ruleset: RuleSet, profile: str) -> dict[str, Any]:
    """HEAD-007（前言/结论/致谢）期望值：alignment 支持按 profile 取字典值。"""
    params = expected_params(ruleset, "HEAD-007", HEAD007_FALLBACK)
    fallback_after = 1.0 if profile == "sci_eng" else 0.0
    return {
        "size_pt": font_size(params, "size_pt", HEAD007_FALLBACK["size_pt"]),
        "east_asia": text_value(
            params, "east_asia", HEAD007_FALLBACK["east_asia"]
        ),
        "alignment": _profile_alignment(
            params.get("alignment"), profile, "center"
        ),
        "line_spacing": _float_value(
            params, "line_spacing", HEAD007_FALLBACK["line_spacing"]
        ),
        "before_lines": line_count(
            params.get("before_lines"), HEAD007_FALLBACK["before_lines"]
        ),
        "after_lines": line_count(params.get("after_lines"), fallback_after),
    }


def _label_spec(
    ruleset: RuleSet,
    rule_id: str,
    fallback: dict[str, Any],
    *,
    roles: tuple[str, ...],
    locator_prefix: str,
    description: str,
    size_key: str = "size_pt",
    east_asia_key: str = "east_asia",
    ascii_key: str = "ascii",
    alignment_key: str = "alignment",
    bold_key: str = "bold",
    force_ascii: bool = False,
) -> _ParagraphSpec:
    """题目/摘要标题/目录标题/参考文献标题的修复 spec（期望值来自当前 ruleset）。

    ``TOC-001`` 的 expected 用 ``title_size_pt`` / ``title_east_asia`` /
    ``title_alignment`` 这类前缀键，通过参数适配；缺字段回退默认学校规范。
    """
    params = expected_params(ruleset, rule_id, fallback)
    return _ParagraphSpec(
        rule_id=rule_id,
        roles=tuple(roles),
        locator_prefix=locator_prefix,
        description=description,
        size_pt=font_size(params, size_key, float(fallback.get(size_key, 12.0))),
        east_asia=_optional_text(params, east_asia_key, fallback.get(east_asia_key)),
        ascii_font=_optional_text(params, ascii_key, fallback.get(ascii_key)),
        force_ascii=force_ascii,
        bold=_optional_bool(params, bold_key, fallback.get(bold_key)),
        alignment=alignment(
            params, alignment_key, str(fallback.get(alignment_key) or "center")
        ),
        line_spacing=_float_value(
            params, "line_spacing", float(fallback.get("line_spacing", 1.5))
        ),
        before_lines=line_count(params.get("before_lines"), fallback.get("before_lines", 1.0)),
        after_lines=line_count(params.get("after_lines"), fallback.get("after_lines", 1.0)),
    )


def _front_specs(profile: str, ruleset: RuleSet) -> list[_ParagraphSpec]:
    """TITLE/ABS 标题、TOC 标题、REF 标题（+ HEAD-007）的修复 spec。

    期望值全部来自当前 ruleset 的 expected：教师 profile 改字号/字体/对齐/行距/段距后
    修复立即生效；缺字段回退默认学校规范（保证 school_default 行为不变）。
    目录条目（toc_entry）/参考文献条目（reference_entry）不在自动修复范围。
    """
    head007 = _head007_expected(ruleset, profile)
    return [
        _label_spec(
            ruleset,
            "TITLE-001",
            TITLE_CN_FALLBACK,
            roles=("title_cn",),
            locator_prefix="front/title_cn",
            description="中文题目格式",
        ),
        _label_spec(
            ruleset,
            "TITLE-002",
            TITLE_EN_FALLBACK,
            roles=("title_en",),
            locator_prefix="front/title_en",
            description="英文题目格式",
            force_ascii=True,
        ),
        _label_spec(
            ruleset,
            "ABS-001",
            ABS_CN_LABEL_FALLBACK,
            roles=("abstract_label_cn",),
            locator_prefix="front/abstract_label_cn",
            description="中文摘要标题格式",
        ),
        _label_spec(
            ruleset,
            "ABS-003",
            ABS_EN_LABEL_FALLBACK,
            roles=("abstract_label_en",),
            locator_prefix="front/abstract_label_en",
            description="英文摘要标题格式",
            force_ascii=True,
        ),
        _ParagraphSpec(
            "HEAD-007",
            ("special_title", "thanks_title"),
            "body/para",
            "前言/结论/致谢标题格式",
            size_pt=head007["size_pt"],
            east_asia=head007["east_asia"],
            alignment=head007["alignment"],
            line_spacing=head007["line_spacing"],
            before_lines=head007["before_lines"],
            after_lines=head007["after_lines"],
        ),
        _label_spec(
            ruleset,
            "TOC-001",
            TOC_TITLE_FALLBACK,
            roles=("toc_title",),
            locator_prefix="front/toc_title",
            description="目录标题格式",
            size_key="title_size_pt",
            east_asia_key="title_east_asia",
            alignment_key="title_alignment",
        ),
        _label_spec(
            ruleset,
            "REF-001",
            REF_TITLE_FALLBACK,
            roles=("refs_title",),
            locator_prefix="front/refs_title",
            description="参考文献标题格式",
        ),
    ]


def _spec_after(spec: _ParagraphSpec) -> dict:
    after: dict = {"size_pt": spec.size_pt}
    if spec.east_asia is not None:
        after["east_asia"] = spec.east_asia
    if spec.ascii_font is not None:
        after["ascii"] = spec.ascii_font
    if spec.bold is not None:
        after["bold"] = spec.bold
    if spec.alignment is not None:
        after["alignment"] = spec.alignment
    if spec.line_spacing is not None:
        after["line_spacing"] = spec.line_spacing
    if spec.before_lines is not None:
        after["before_lines"] = spec.before_lines
    if spec.after_lines is not None:
        after["after_lines"] = spec.after_lines
    return after


def _apply_paragraph_spec(paragraph: Any, spec: _ParagraphSpec) -> None:
    _apply_run_fonts(
        paragraph,
        size_pt=spec.size_pt,
        east_asia=spec.east_asia,
        ascii_font=spec.ascii_font,
        force_ascii=spec.force_ascii,
        bold=spec.bold,
    )
    if spec.alignment is not None:
        _set_alignment(paragraph, spec.alignment)
    if spec.line_spacing is not None:
        _set_line_spacing(paragraph, spec.line_spacing)
    if spec.before_lines is not None or spec.after_lines is not None:
        _set_before_after_lines(
            paragraph,
            before_lines=spec.before_lines,
            after_lines=spec.after_lines,
        )


def _fix_paragraphs_by_spec(
    model: DocModel, roles: Roles, document: Any, specs: list[_ParagraphSpec]
) -> list[FixAction]:
    actions: list[FixAction] = []
    paragraphs = document.paragraphs
    for spec in specs:
        for index in roles.indices(*spec.roles):
            para = model.by_index.get(index)
            if para is None or para.is_empty or index >= len(paragraphs):
                continue
            problems = para_format_problems(
                para,
                size_pt=spec.size_pt,
                east_asia=spec.east_asia,
                ascii_font=spec.ascii_font,
                bold=spec.bold,
                check_bold=spec.bold is not None,
                alignment=spec.alignment,
                line_spacing=spec.line_spacing,
                before_lines=spec.before_lines,
                after_lines=spec.after_lines,
            )
            # “段落无文字内容”只可能出现在空段（上面已跳过），其余都是格式问题
            problems = [item for item in problems if item != "段落无文字内容"]
            if not problems:
                continue
            _apply_paragraph_spec(paragraphs[index], spec)
            detail = "；".join(dict.fromkeys(problems))
            actions.append(
                FixAction(
                    rule_id=spec.rule_id,
                    locator=loc_para(para, spec.locator_prefix),
                    description=f"{spec.description}归一化：{detail}",
                    before=detail,
                    after=_spec_after(spec),
                )
            )
    return actions


# ---------------------------------------------------------------------------
# 关键词：标签（中文黑体 / 英文 TNR 加粗）与正文分开处理，绝不改文字
# ---------------------------------------------------------------------------
def _keyword_label_element(paragraph: Any, is_en: bool) -> Any:
    """返回关键词标签 run 的底层 w:r 元素（不要用 Run 代理对象做身份比较）。"""
    for run in paragraph.runs:
        text = (run.text or "").lstrip()
        if is_en:
            if text.lower().startswith(("key word", "keywords", "key words")):
                return run._r
        elif text.startswith(("关键词", "关键字")):
            return run._r
    return None


def _keyword_expected_params(
    ruleset: RuleSet, rule_id: str, fallback: dict[str, Any], *, is_en: bool
) -> dict[str, Any]:
    """KW-001 / KW-002 的字体/字号期望值（来自当前 ruleset 的 expected，缺字段回退默认值）。"""
    params = expected_params(ruleset, rule_id, fallback)
    spec: dict[str, Any] = {
        "label_size_pt": font_size(
            params, "label_size_pt", float(fallback["label_size_pt"])
        ),
        "body_size_pt": font_size(
            params, "body_size_pt", float(fallback["body_size_pt"])
        ),
    }
    if is_en:
        spec["label_ascii"] = text_value(params, "label_ascii", fallback["label_ascii"])
        spec["label_bold"] = _optional_bool(
            params, "label_bold", bool(fallback["label_bold"])
        )
        spec["body_ascii"] = text_value(params, "body_ascii", fallback["body_ascii"])
    else:
        spec["label_east_asia"] = text_value(
            params, "label_east_asia", fallback["label_east_asia"]
        )
        spec["body_east_asia"] = text_value(
            params, "body_east_asia", fallback["body_east_asia"]
        )
        spec["body_ascii"] = text_value(params, "body_ascii", fallback["body_ascii"])
    return spec


def _keyword_font_problems(
    para: ParaInfo, is_en: bool, spec: dict[str, Any]
) -> list[str]:
    """复刻检查器的字体/字号判定（只有这部分属于“安全可修”，参数来自 ruleset）。"""
    problems: list[str] = []
    label_run = None
    for run in para.runs:
        text = (run.text or "").lstrip()
        if is_en:
            if text.lower().startswith(("key word", "keywords", "key words")):
                label_run = run
                break
        elif text.startswith(("关键词", "关键字")):
            label_run = run
            break
    for run in para.runs:
        text = run.text or ""
        if not text.strip():
            continue
        if run is label_run:
            if not close(run.size_pt, spec["label_size_pt"], PT_TOL):
                problems.append(
                    f"标签字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
                )
            if is_en:
                if run.ascii != spec["label_ascii"]:
                    problems.append(f"标签字体={run.ascii or '未设置'}")
                if bool(run.bold) != spec["label_bold"]:
                    problems.append(
                        "标签未加粗" if spec["label_bold"] else "标签不应加粗"
                    )
            elif run.east_asia != spec["label_east_asia"]:
                problems.append(f"标签字体={run.east_asia or '未设置'}")
            continue
        if not close(run.size_pt, spec["body_size_pt"], PT_TOL):
            problems.append(
                f"关键词字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
            )
        if is_en:
            if has_latin(text) and run.ascii != spec["body_ascii"]:
                problems.append(f"关键词字体={run.ascii or '未设置'}")
        else:
            if has_cjk(text) and run.east_asia != spec["body_east_asia"]:
                problems.append(f"关键词中文字体={run.east_asia or '未设置'}")
            if has_latin(text) and run.ascii != spec["body_ascii"]:
                problems.append(f"关键词西文字体={run.ascii or '未设置'}")
    return problems


def _normalize_keywords(
    paragraph: Any, is_en: bool, spec: dict[str, Any]
) -> None:
    label_element = _keyword_label_element(paragraph, is_en)
    for run in paragraph.runs:
        text = run.text or ""
        if not text.strip():
            continue
        if run._r is label_element:
            if is_en:
                _apply_font_to_run(
                    run,
                    text,
                    size_pt=spec["label_size_pt"],
                    ascii_font=spec["label_ascii"],
                    force_ascii=True,
                    bold=spec["label_bold"],
                )
            else:
                _apply_font_to_run(
                    run,
                    text,
                    size_pt=spec["label_size_pt"],
                    east_asia=spec["label_east_asia"],
                )
            continue
        if is_en:
            _apply_font_to_run(
                run,
                text,
                size_pt=spec["body_size_pt"],
                ascii_font=spec["body_ascii"],
            )
        else:
            _apply_font_to_run(
                run,
                text,
                size_pt=spec["body_size_pt"],
                east_asia=spec["body_east_asia"],
                ascii_font=spec["body_ascii"],
            )


def _fix_keywords(
    model: DocModel, roles: Roles, document: Any, ruleset: RuleSet
) -> list[FixAction]:
    actions: list[FixAction] = []
    paragraphs = document.paragraphs
    for role, rule_id, prefix, is_en, fallback in (
        ("keywords_cn", "KW-001", "front/keywords_cn", False, KW_CN_FALLBACK),
        ("keywords_en", "KW-002", "front/keywords_en", True, KW_EN_FALLBACK),
    ):
        spec = _keyword_expected_params(ruleset, rule_id, fallback, is_en=is_en)
        for index in roles.indices(role):
            para = model.by_index.get(index)
            if para is None or para.is_empty or index >= len(paragraphs):
                continue
            problems = _keyword_font_problems(para, is_en, spec)
            if not problems:
                continue
            _normalize_keywords(paragraphs[index], is_en, spec)
            detail = "；".join(dict.fromkeys(problems))
            if is_en:
                after: dict = {
                    "label_size_pt": spec["label_size_pt"],
                    "label_ascii": spec["label_ascii"],
                    "label_bold": spec["label_bold"],
                    "body_size_pt": spec["body_size_pt"],
                    "body_ascii": spec["body_ascii"],
                }
            else:
                after = {
                    "label_size_pt": spec["label_size_pt"],
                    "label_east_asia": spec["label_east_asia"],
                    "body_size_pt": spec["body_size_pt"],
                    "body_east_asia": spec["body_east_asia"],
                    "body_ascii": spec["body_ascii"],
                }
            actions.append(
                FixAction(
                    rule_id=rule_id,
                    locator=loc_para(para, prefix),
                    description=(
                        f"{'英文' if is_en else '中文'}关键词字体/字号归一化"
                        f"（标签与正文分开）：{detail}"
                    ),
                    before=detail,
                    after=after,
                )
            )
    return actions


# ---------------------------------------------------------------------------
# TAB-002「表内文字」：只改单元格字体/字号/对齐/行距，不动文字/表题/表注
# ---------------------------------------------------------------------------
def _direct_jc_value(paragraph: Any) -> Optional[str]:
    """读段落直接格式里的 w:jc/@w:val（无直接格式返回 None，不解析样式继承）。"""
    ppr = paragraph._p.pPr
    if ppr is None:
        return None
    jc = ppr.find(qn("w:jc"))
    return jc.get(qn("w:val")) if jc is not None else None


def _direct_line_spacing(paragraph: Any) -> Optional[float]:
    """读段落直接格式里的行距倍数；未设置/固定值返回 None（固定值需要改成倍数）。"""
    value = paragraph.paragraph_format.line_spacing
    if isinstance(value, (int, float)):  # 倍数值；Length（固定值/最小值）不算
        return float(value)
    return None


def _line_spacing_before_text(paragraph: Any) -> str:
    value = paragraph.paragraph_format.line_spacing
    if value is None:
        return "未设置"
    if isinstance(value, (int, float)):
        return f"{float(value):g} 倍"
    return f"固定值 {getattr(value, 'pt', 0):g}pt"


def _direct_run_fonts(run: Any) -> tuple[Optional[float], Optional[str], Optional[str]]:
    """读 run 直接格式的 (字号pt, eastAsia, ascii)；未设置返回 None。"""
    size = run.font.size
    size_pt = float(size.pt) if size is not None else None
    rpr = run._r.rPr
    rfonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
    east_asia = rfonts.get(qn("w:eastAsia")) if rfonts is not None else None
    ascii_font = rfonts.get(qn("w:ascii")) if rfonts is not None else None
    return size_pt, east_asia, ascii_font


def _set_table_run_font(run: Any, *, size_pt: float, east_asia: str) -> None:
    """表内文字 run：字号 + 宋体(eastAsia) + Times New Roman(西文)。"""
    run.font.size = Pt(size_pt)
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), east_asia)
    run.font.name = TABLE_ASCII_FONT  # 写 w:rFonts/@w:ascii 与 @w:hAnsi
    rpr.get_or_add_rFonts().set(qn("w:cs"), TABLE_ASCII_FONT)


def _tables_expected(ruleset: RuleSet) -> dict[str, Any]:
    """TAB-002 表内文字期望值：优先 ruleset.expected（教师 profile 可覆盖），缺字段用默认。"""
    params = expected_params(ruleset, "TAB-002", TAB002_FALLBACK)
    align_value = alignment(params, "cell_alignment", "center")
    if align_value not in _ALIGN_MAP:
        align_value = "center"
    return {
        "size_pt": _float_value(params, "cell_size_pt", 10.5),
        "east_asia": text_value(params, "cell_east_asia", "宋体") or "宋体",
        "alignment": align_value,
        "line_spacing": _float_value(params, "cell_line_spacing", 1.5),
    }


def _fix_tables(model: DocModel, document: Any, ruleset: RuleSet) -> list[FixAction]:
    """TAB-002 表内文字安全修复：单元格字号/中文字体/对齐/行距归一化。

    - 只改格式属性（rPr 字号/字体、pPr 对齐/行距），**不改任何文字**，也不动表题/表注；
    - 表内文字期望值来自当前 ruleset 的 TAB-002.expected（缺字段回退学校默认值）；
    - 合并单元格按 ``cell._tc`` 去重；空单元格跳过（检查器只统计有文字的单元格）；
    - 每张表只在确有改动时记 1 条 FixAction（locator ``table/N``，N 与检查器同序）；
    - 按“直接格式对比”判断是否需要改，第二次修复同一文档不再产生动作（幂等）。
    """
    rule = ruleset.rules.get("TAB-002")
    if rule is None:
        return []
    if not model.tables or not document.tables:
        return []

    spec = _tables_expected(ruleset)
    size_pt: float = spec["size_pt"]
    east_asia: str = spec["east_asia"]
    align_value: str = spec["alignment"]
    line_spacing: float = spec["line_spacing"]
    expected_jc = "both" if align_value in ("both", "justify") else align_value

    actions: list[FixAction] = []
    for table_index, table in enumerate(document.tables, start=1):
        changes = 0
        before_values: dict[str, set[str]] = {
            "字号": set(),
            "中文字体": set(),
            "西文字体": set(),
            "对齐": set(),
            "行距": set(),
        }
        seen_cells: list[Any] = []
        for row in table.rows:
            for cell in row.cells:
                # 合并单元格在 row.cells 里会重复出现同一个 w:tc：用元素对象身份去重
                # （注意不能只用 id()：代理对象被回收后 id 可能复用，会误跳过正常单元格）
                tc = cell._tc
                if any(tc is seen for seen in seen_cells):
                    continue
                seen_cells.append(tc)
                if not (cell.text or "").strip():
                    continue  # 检查器只统计有文字的单元格
                for paragraph in cell.paragraphs:
                    jc = _direct_jc_value(paragraph)
                    if jc != expected_jc:
                        before_values["对齐"].add(
                            _ALIGNMENT_LABELS.get(jc, jc) if jc else "未设置"
                        )
                        _set_alignment(paragraph, align_value)
                        changes += 1

                    if not close(
                        _direct_line_spacing(paragraph), line_spacing, SPACING_TOL
                    ):
                        before_values["行距"].add(_line_spacing_before_text(paragraph))
                        _set_line_spacing(paragraph, line_spacing)
                        changes += 1

                    for run in paragraph.runs:
                        text = run.text or ""
                        if not text.strip():
                            continue
                        run_size, run_east_asia, run_ascii = _direct_run_fonts(run)
                        if run_size is None or abs(run_size - size_pt) > 0.01:
                            before_values["字号"].add(
                                "未设置" if run_size is None else f"{run_size:g}pt"
                            )
                            changes += 1
                        if run_east_asia != east_asia:
                            before_values["中文字体"].add(run_east_asia or "未设置")
                            changes += 1
                        if run_ascii != TABLE_ASCII_FONT:
                            before_values["西文字体"].add(run_ascii or "未设置")
                            changes += 1
                        if (
                            (run_size is None or abs(run_size - size_pt) > 0.01)
                            or run_east_asia != east_asia
                            or run_ascii != TABLE_ASCII_FONT
                        ):
                            _set_table_run_font(
                                run, size_pt=size_pt, east_asia=east_asia
                            )

        if not changes:
            continue
        before_text = "；".join(
            f"{label}={'/'.join(sorted(values))}"
            for label, values in before_values.items()
            if values
        )
        after = {
            "size_pt": size_pt,
            "east_asia": east_asia,
            "ascii": TABLE_ASCII_FONT,
            "alignment": align_value,
            "line_spacing": line_spacing,
        }
        actions.append(
            FixAction(
                rule_id="TAB-002",
                locator=f"table/{table_index}",
                description=(
                    f"表格单元格字体/字号/对齐/行距归一化（{changes} 处）："
                    f"{before_text or '直接格式未设置'}"
                ),
                before=before_text or "直接格式未设置",
                after=after,
            )
        )
    return actions


# ---------------------------------------------------------------------------
# skipped：非 safe / 涉及文字 / 目录与参考文献条目
# ---------------------------------------------------------------------------
def _role_from_locator(locator: str) -> Optional[str]:
    for role in SKIPPED_ROLES:
        if f"/{role}/" in locator or locator.startswith(f"front/{role}"):
            return role
    return None


def _build_skipped(
    baseline_issues: list[Issue], recheck_issues: list[Issue], roles: Roles
) -> list[dict]:
    recheck_keys = {(issue.rule_id, issue.locator) for issue in recheck_issues}
    skipped: list[dict] = []
    seen: set[tuple] = set()
    for issue in baseline_issues:
        key = (issue.rule_id, issue.locator, issue.message)
        if key in seen:
            continue
        seen.add(key)
        role = _role_from_locator(issue.locator)
        if issue.rule_id not in _SAFE_RULE_ID_SET:
            reason = "该规则不在安全修复白名单内（涉及文字/编号/结构），需人工处理"
        elif role in SKIPPED_ROLES:
            reason = "目录条目/参考文献条目本轮不自动修改（涉及目录域或著录内容），需人工处理"
        elif (issue.rule_id, issue.locator) in recheck_keys:
            reason = (
                "该问题无法仅靠格式属性修复（涉及文字内容、字符间距或文档结构），"
                "需人工确认"
            )
        else:
            continue
        skipped.append(
            {
                "kind": "issue",
                "rule_id": issue.rule_id,
                "title": issue.title,
                "severity": issue.severity,
                "locator": issue.locator,
                "message": issue.message,
                "reason": reason,
            }
        )

    # 明确告知“这些角色整体跳过”，即使当前检查器没有为它们报问题
    for role, label, reason in (
        (
            "reference_entry",
            "参考文献条目",
            "参考文献条目一律不自动修改：改动可能影响著录内容，需人工确认",
        ),
        (
            "toc_entry",
            "目录条目",
            "目录条目涉及 TOC 域/页码，需在 Word 中更新目录域后人工确认",
        ),
    ):
        indices = roles.indices(role)
        if indices:
            skipped.append(
                {
                    "kind": "scope",
                    "role": role,
                    "label": label,
                    "count": len(indices),
                    "sample_indices": indices[:10],
                    "reason": reason,
                }
            )
    return skipped


# ---------------------------------------------------------------------------
# 路径 / 复检
# ---------------------------------------------------------------------------
def _same_path(left: Path, right: Path) -> bool:
    try:
        return os.path.normcase(os.path.abspath(str(left))) == os.path.normcase(
            os.path.abspath(str(right))
        )
    except (OSError, ValueError):  # pragma: no cover - 极端路径兜底
        return False


def _resolve_out_path(src: Path, out: Any) -> Path:
    candidate = src.with_name(f"{src.stem}_fixed.docx") if out is None else Path(out)
    if _same_path(src, candidate):
        raise ValueError(
            f"输出路径不能与源文件相同（自动修复不覆盖原文件）：{candidate}"
        )
    return candidate


def _recheck(path: Path, ruleset: RuleSet, profile: str) -> tuple[list[Issue], str]:
    model = DocModel.load(path)
    roles = classify(model, profile)
    issues = run_checks(model, roles, ruleset, profile=profile)
    return issues, content_hash(model)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def fix_document(
    src: Path | str,
    out: Path | str | None = None,
    *,
    profile: str = DEFAULT_PROFILE,
    ruleset: Optional[RuleSet] = None,
    dry_run: bool = False,
) -> FixResult:
    """把 ``src`` 的格式问题安全修复到 ``out``（默认 ``<原名>_fixed.docx``）。

    参数：
        src: 源 .docx（只读，不会被修改）。
        out: 输出路径；None 时与源文件同目录、文件名追加 ``_fixed``。
        profile: 学科规范体系（``sci_eng`` / ``humanities``）。
        ruleset: 规则集；None 时加载 ``rules/rules.json``。
        dry_run: 只计算 actions / 复检，不写任何输出文件，``out_path`` 返回 None。

    返回 :class:`FixResult`；内容 hash 不一致或源文件被意外改动时抛 RuntimeError。
    """
    src_path = Path(src)
    rules = ruleset if ruleset is not None else RuleSet.load()

    model = DocModel.load(src_path)
    roles = classify(model, profile)
    baseline_issues = run_checks(model, roles, rules, profile=profile)
    baseline_counts = _count_issues(baseline_issues)
    hash_before = content_hash(model)
    src_bytes = src_path.read_bytes()

    out_path: Optional[Path] = None if dry_run else _resolve_out_path(src_path, out)

    document = docx.Document(str(src_path))
    actions: list[FixAction] = []
    actions.extend(_fix_pages(model, document, rules))
    actions.extend(_fix_body(model, roles, document, rules))
    # HEAD-001/002/003 是理工农（sci_eng）专用规则；其他 profile 的检查器不检查标题层级格式，
    # 修复端也不能擅自套用，避免把人文社科文档改坏。
    paragraph_specs = _front_specs(profile, rules)
    if profile == "sci_eng":
        paragraph_specs = _heading_specs(rules) + paragraph_specs
    actions.extend(_fix_paragraphs_by_spec(model, roles, document, paragraph_specs))
    actions.extend(_fix_keywords(model, roles, document, rules))
    actions.extend(_fix_tables(model, document, rules))

    if dry_run:
        # dry-run 不落地输出文件；用系统临时目录里的副本做“真实复检”，确保结论可信。
        with tempfile.TemporaryDirectory(prefix="tas-fix-dryrun-") as tmp_dir:
            probe = Path(tmp_dir) / "probe_fixed.docx"
            document.save(str(probe))
            recheck_issues, hash_after = _recheck(probe, rules, profile)
    else:
        assert out_path is not None
        if not out_path.parent.exists():
            out_path.parent.mkdir(parents=True, exist_ok=True)
        document.save(str(out_path))
        recheck_issues, hash_after = _recheck(out_path, rules, profile)
        if src_path.read_bytes() != src_bytes:
            raise RuntimeError(
                "修复过程中源文件被意外修改；已中止（自动修复绝不覆盖原文件）"
            )

    if hash_before != hash_after:
        raise RuntimeError(
            "内容 hash 不一致：修复前后正文/表格/页眉页脚文本发生变化。"
            f"before={hash_before} after={hash_after}"
        )

    return FixResult(
        src_path=src_path,
        out_path=out_path,
        dry_run=dry_run,
        actions=actions,
        skipped=_build_skipped(baseline_issues, recheck_issues, roles),
        content_hash_before=hash_before,
        content_hash_after=hash_after,
        recheck_issues=recheck_issues,
        baseline_counts=baseline_counts,
    )
