"""正文检查：字体、行距、首行缩进、对齐（参数来自当前 ruleset）。"""

from __future__ import annotations

from ..config import RuleSet
from ..docx_io import DocModel, ParaInfo
from ..report import Issue
from ..roles import Roles
from ..rule_params import alignment, char_count, expected_params, font_size
from .helpers import (
    CHAR_TOL,
    SPACING_TOL,
    close,
    indent_chars,
    is_justified,
    loc_para,
    make_issue,
    run_font_problems,
)

BODY_FONT_FALLBACK = {"size_pt": 12, "east_asia": "宋体", "ascii": "Times New Roman"}
BODY_SPACING_FALLBACK = {"line_spacing": 1.5}
BODY_INDENT_FALLBACK = {"first_line_chars": 200}
BODY_ALIGN_FALLBACK = {"alignment": "justify"}


def check_body(model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng") -> list[Issue]:
    body = [model.by_index[index] for index in roles.indices("body")]
    body = [para for para in body if not para.is_empty]
    issues: list[Issue] = []
    issues.extend(_check_font(body, ruleset))
    issues.extend(_check_spacing(body, ruleset))
    issues.extend(_check_indent(body, ruleset))
    issues.extend(_check_alignment(body, ruleset))
    return issues


def _check_font(body: list[ParaInfo], ruleset: RuleSet) -> list[Issue]:
    rule = ruleset.get("BODY-001")
    params = expected_params(ruleset, "BODY-001", BODY_FONT_FALLBACK)
    size = font_size(params)
    east_asia = str(params.get("east_asia", "宋体"))
    ascii_font = str(params.get("ascii", "Times New Roman"))
    expected = {"size_pt": size, "east_asia": east_asia, "ascii": ascii_font}
    issues = []
    for para in body:
        problems = run_font_problems(
            para, expected_size=size, east_asia=east_asia, ascii_font=ascii_font
        )
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="正文的字体/字号不符合当前规范",
                    actual="；".join(problems),
                    expected=expected,
                )
            )
        nonempty_runs = [run for run in para.runs if (run.text or "").strip()]
        if nonempty_runs and all(run.bold is True for run in nonempty_runs):
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="整段正文被加粗（规范未要求；可能是误把正文当标题）",
                    actual="整段加粗",
                    expected="正文不加粗",
                    severity="warning",
                )
            )
    return issues


def _check_spacing(body: list[ParaInfo], ruleset: RuleSet) -> list[Issue]:
    rule = ruleset.get("BODY-002")
    params = expected_params(ruleset, "BODY-002", BODY_SPACING_FALLBACK)
    try:
        expected_spacing = float(params.get("line_spacing", 1.5))
    except (TypeError, ValueError):
        expected_spacing = 1.5
    issues = []
    for para in body:
        if close(para.line_spacing, expected_spacing, SPACING_TOL):
            continue
        if para.line_spacing is not None:
            actual = f"{para.line_spacing:g} 倍"
        elif para.line_spacing_pt is not None:
            actual = f"固定值 {para.line_spacing_pt:g}pt"
        else:
            actual = "未设置（继承默认）"
        issues.append(
            make_issue(
                rule,
                locator=loc_para(para),
                message=f"正文行距不是 {expected_spacing:g} 倍",
                actual=actual,
                expected=f"{expected_spacing:g} 倍行距",
            )
        )
    return issues


def _check_indent(body: list[ParaInfo], ruleset: RuleSet) -> list[Issue]:
    rule = ruleset.get("BODY-003")
    params = expected_params(ruleset, "BODY-003", BODY_INDENT_FALLBACK)
    expected_chars = char_count(params.get("first_line_chars"), 2.0)
    issues = []
    for para in body:
        chars = indent_chars(para)
        if close(chars, expected_chars, CHAR_TOL):
            continue
        actual = f"{chars:.2f} 字符" if chars is not None else "未设置（或为悬挂缩进）"
        issues.append(
            make_issue(
                rule,
                locator=loc_para(para),
                message=f"正文首行缩进不是 {expected_chars:g} 字符",
                actual=actual,
                expected=f"首行缩进 {expected_chars:g} 字符",
            )
        )
    return issues


def _check_alignment(body: list[ParaInfo], ruleset: RuleSet) -> list[Issue]:
    rule = ruleset.get("BODY-004")
    params = expected_params(ruleset, "BODY-004", BODY_ALIGN_FALLBACK)
    wanted = alignment(params)
    expected_label = "both（两端对齐）" if wanted in ("justify", "both") else wanted
    issues = []
    for para in body:
        if wanted in ("justify", "both"):
            ok = is_justified(para)
        else:
            ok = (para.alignment or "left") == wanted
        if not ok:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message=f"正文段落对齐方式不是 {wanted}",
                    actual=para.alignment or "left（未设置，默认左对齐）",
                    expected=expected_label,
                )
            )
    return issues
