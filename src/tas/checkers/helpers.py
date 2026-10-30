"""检查器共用工具。"""

from __future__ import annotations

from typing import Iterable, Optional

from ..config import Rule
from ..docx_io import ParaInfo
from ..report import Issue

# “段前 1 行”在 Word 里的原生单位是 w:beforeLines（1/100 行）。
# 但 python-docx 只能写 pt，历史生成脚本用 pt 近似；对只有 pt 的文档按经验值换算。
PT_PER_LINE = 18.0

CM_TOL = 0.03
PT_TOL = 0.6
SPACING_TOL = 0.02
CHAR_TOL = 0.06
LINE_TOL = 0.25


def make_issue(
    rule: Rule,
    *,
    locator: str,
    message: str,
    actual=None,
    expected=None,
    hint: str = "",
    severity: Optional[str] = None,
) -> Issue:
    return Issue(
        rule_id=rule.id,
        title=rule.name,
        severity=severity or rule.severity,
        message=message,
        locator=locator,
        actual=actual,
        expected=expected,
        hint=hint,
    )


def close(actual: Optional[float], expected: float, tolerance: float) -> bool:
    return actual is not None and abs(actual - expected) <= tolerance


def has_cjk(text: str) -> bool:
    return any("\u4e00" <= char <= "\u9fff" for char in text)


def has_latin(text: str) -> bool:
    return any(char.isascii() and char.isalpha() for char in text)


def para_brief(para: ParaInfo, limit: int = 24) -> str:
    text = (para.text or "").strip().replace("\n", " ").replace("\t", " ")
    text = " ".join(text.split())
    if len(text) > limit:
        text = text[:limit] + "…"
    return text


def loc_para(para: ParaInfo, prefix: str = "body/para") -> str:
    return f"{prefix}/{para.index}「{para_brief(para)}」"


def first_run(para: ParaInfo):
    for run in para.runs:
        if (run.text or "").strip():
            return run
    return None


def effective_before_lines(para: ParaInfo) -> Optional[float]:
    if para.before_lines is not None:
        return para.before_lines
    if para.before_pt is not None:
        return para.before_pt / PT_PER_LINE
    return None


def effective_after_lines(para: ParaInfo) -> Optional[float]:
    if para.after_lines is not None:
        return para.after_lines
    if para.after_pt is not None:
        return para.after_pt / PT_PER_LINE
    return None


def run_font_problems(
    para: ParaInfo,
    expected_size: Optional[float] = None,
    *,
    east_asia: str = "宋体",
    ascii_font: str = "Times New Roman",
) -> list[str]:
    """正文类：逐个 run 检查字体/字号（按文字脚本判断该查哪一项）。

    east_asia / ascii_font 由当前 ruleset 参数传入（默认学校规范）。
    """
    problems: list[str] = []
    for run in para.runs:
        text = run.text or ""
        if not text.strip():
            continue
        if expected_size is not None and not close(run.size_pt, expected_size, PT_TOL):
            problems.append(f"字号={run.size_pt if run.size_pt is not None else '未设置'}pt")
        if has_cjk(text) and run.east_asia != east_asia:
            problems.append(f"中文字体={run.east_asia or '未设置'}")
        if has_latin(text) and run.ascii != ascii_font:
            problems.append(f"西文字体={run.ascii or '未设置'}")
    return list(dict.fromkeys(problems))


def para_format_problems(
    para: ParaInfo,
    *,
    size_pt: Optional[float] = None,
    east_asia: Optional[str] = None,
    ascii_font: Optional[str] = None,
    bold: Optional[bool] = None,
    check_bold: bool = False,
    alignment: Optional[str] = None,
    line_spacing: Optional[float] = None,
    before_lines: Optional[float] = None,
    after_lines: Optional[float] = None,
) -> list[str]:
    """标题/标签类段落：以第一个非空 run 为代表做整体格式对比。"""
    problems: list[str] = []
    run = first_run(para)
    if run is not None:
        if size_pt is not None and not close(run.size_pt, size_pt, PT_TOL):
            problems.append(f"字号={run.size_pt if run.size_pt is not None else '未设置'}pt")
        if east_asia is not None and run.east_asia != east_asia:
            problems.append(f"中文字体={run.east_asia or '未设置'}")
        if ascii_font is not None and run.ascii != ascii_font:
            problems.append(f"西文字体={run.ascii or '未设置'}")
        if check_bold and bold is not None:
            actual_bold = bool(run.bold) if run.bold is not None else False
            if actual_bold != bold:
                problems.append(f"加粗={actual_bold}")
    else:
        problems.append("段落无文字内容")

    if alignment is not None:
        actual_alignment = para.alignment or "left"
        if actual_alignment != alignment and not (
            alignment == "justify" and actual_alignment == "both"
        ):
            problems.append(f"对齐={actual_alignment}")

    if line_spacing is not None and not close(para.line_spacing, line_spacing, SPACING_TOL):
        problems.append(
            f"行距={para.line_spacing if para.line_spacing is not None else '未设置'}倍"
        )

    if before_lines is not None:
        actual = effective_before_lines(para)
        if actual is None:
            problems.append("段前距未设置")
        elif abs(actual - before_lines) > LINE_TOL:
            problems.append(f"段前={actual:.2f}行")

    if after_lines is not None:
        actual = effective_after_lines(para)
        if actual is None:
            problems.append("段后距未设置")
        elif abs(actual - after_lines) > LINE_TOL:
            problems.append(f"段后={actual:.2f}行")

    return problems


def is_justified(para: ParaInfo) -> bool:
    return (para.alignment or "left") in ("both", "justify")


def indent_chars(para: ParaInfo) -> Optional[float]:
    """首行缩进（字符数）。优先 firstLineChars，其次按字号换算 firstLine(twips)。"""
    if para.first_line_chars is not None:
        return para.first_line_chars
    if para.first_line_twips is not None:
        size = 12.0
        run = first_run(para)
        if run is not None and run.size_pt:
            size = run.size_pt
        return para.first_line_twips / (20.0 * size)
    return None


def body_paragraph_problems(para: ParaInfo) -> list[str]:
    """正文/摘要正文这类“小四宋体、1.5 倍、首行缩进 2 字符、两端对齐”的段落。"""
    problems = run_font_problems(para, expected_size=12.0)
    if not close(para.line_spacing, 1.5, SPACING_TOL):
        actual = (
            f"{para.line_spacing:g} 倍"
            if para.line_spacing is not None
            else (f"固定值 {para.line_spacing_pt:g}pt" if para.line_spacing_pt is not None else "未设置")
        )
        problems.append(f"行距={actual}")
    chars = indent_chars(para)
    if not close(chars, 2.0, CHAR_TOL):
        problems.append(f"首行缩进={chars:.2f}字符" if chars is not None else "首行缩进未设置")
    if not is_justified(para):
        problems.append(f"对齐={para.alignment or 'left'}")
    return problems


def label_gap_spaces(raw_text: str, label: str) -> Optional[int]:
    """检查“两字标题中间空 2 格”：返回两个非空字符之间的空格数；结构不匹配返回 None。"""
    stripped = (raw_text or "").strip()
    if not stripped:
        return None
    chars = [char for char in stripped if not char.isspace()]
    if "".join(chars) != label:
        return None
    positions = [index for index, char in enumerate(stripped) if not char.isspace()]
    if len(positions) < 2:
        return None
    inner = stripped[positions[0] + 1 : positions[-1]]
    return sum(1 for char in inner if char in " \u3000")
