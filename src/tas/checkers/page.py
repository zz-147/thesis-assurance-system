"""页面设置检查：纸张、页边距/装订线、文档网格（参数来自当前 ruleset）。"""

from __future__ import annotations

from ..config import RuleSet
from ..docx_io import TWIPS_PER_CM, DocModel
from ..report import Issue
from ..roles import Roles
from ..rule_params import expected_params, text_value
from .helpers import CM_TOL, close, make_issue


def _r(value):
    return None if value is None else round(value, 2)


def _num(params: dict, key: str, fallback: float) -> float:
    try:
        return float(params.get(key, fallback))
    except (TypeError, ValueError):
        return fallback


def check_page(model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng") -> list[Issue]:
    issues: list[Issue] = []
    issues.extend(_page_size(model, ruleset))
    issues.extend(_margins(model, ruleset))
    issues.extend(_doc_grid(model, ruleset))
    return issues


def _page_size(model: DocModel, ruleset: RuleSet) -> list[Issue]:
    rule = ruleset.get("PAGE-001")
    params = expected_params(
        ruleset, "PAGE-001", {"page_width_cm": 21.0, "page_height_cm": 29.7}
    )
    width = _num(params, "page_width_cm", 21.0)
    height = _num(params, "page_height_cm", 29.7)
    expected = {"width_cm": width, "height_cm": height}
    issues = []
    for section in model.sections:
        bad = {}
        if not close(section.page_width_cm, width, CM_TOL):
            bad["width_cm"] = _r(section.page_width_cm)
        if not close(section.page_height_cm, height, CM_TOL):
            bad["height_cm"] = _r(section.page_height_cm)
        if bad:
            issues.append(
                make_issue(
                    rule,
                    locator=f"section/{section.index}",
                    message=f"页面尺寸不符合当前规范（期望 {width:g}cm × {height:g}cm）",
                    actual=bad,
                    expected=expected,
                    hint="Word：布局 → 页面设置 → 纸张",
                )
            )
    return issues


_MARGIN_FALLBACK = {
    "margin_top_cm": 2.54,
    "margin_bottom_cm": 2.54,
    "margin_left_cm": 2.0,
    "margin_right_cm": 2.0,
    "gutter_cm": 1.0,
    "header_distance_cm": 2.0,
    "footer_distance_cm": 2.0,
}


def _margins(model: DocModel, ruleset: RuleSet) -> list[Issue]:
    rule = ruleset.get("PAGE-002")
    params = expected_params(ruleset, "PAGE-002", _MARGIN_FALLBACK)
    expected = {key: _num(params, key, fallback) for key, fallback in _MARGIN_FALLBACK.items()}
    issues = []
    for section in model.sections:
        actuals = {
            "margin_top_cm": section.margin_top_cm,
            "margin_bottom_cm": section.margin_bottom_cm,
            "margin_left_cm": section.margin_left_cm,
            "margin_right_cm": section.margin_right_cm,
            "gutter_cm": section.gutter_cm,
            "header_distance_cm": section.header_distance_cm,
            "footer_distance_cm": section.footer_distance_cm,
        }
        bad = {
            name: _r(value)
            for name, value in actuals.items()
            if not close(value, expected[name], CM_TOL)
        }
        if bad:
            issues.append(
                make_issue(
                    rule,
                    locator=f"section/{section.index}",
                    message="页边距/装订线/页眉页脚距边界不符合当前规范",
                    actual=bad,
                    expected=expected,
                    hint="Word：布局 → 页面设置 → 页边距；装订线与左页边距是独立属性",
                )
            )
    return issues


def _doc_grid(model: DocModel, ruleset: RuleSet) -> list[Issue]:
    rule = ruleset.get("PAGE-003")
    if not model.sections:
        return []
    section = model.sections[0]
    params = expected_params(
        ruleset,
        "PAGE-003",
        {"type": "linesAndChars", "lines_per_page": 31, "chars_per_line": 35, "columns": 1},
    )
    lines_per_page = int(_num(params, "lines_per_page", 31))
    chars_per_line = int(_num(params, "chars_per_line", 35))
    grid_type = text_value(params, "type", "linesAndChars")
    expected = {
        "type": grid_type,
        "lines_per_page": lines_per_page,
        "chars_per_line": chars_per_line,
        "columns": int(_num(params, "columns", 1)),
    }
    if not section.doc_grid_type:
        return [
            make_issue(
                rule,
                locator="section/1",
                message=f"未设置文档网格（当前规范要求每页 {lines_per_page} 行、每行 {chars_per_line} 字符）",
                actual="未设置",
                expected=expected,
                hint="Word：布局 → 页面设置 → 文档网格 → 指定行和字符网格；不同 Word 版本换算有差异，需老师模板实测",
            )
        ]

    text_height_twips = None
    if (
        section.page_height_cm is not None
        and section.margin_top_cm is not None
        and section.margin_bottom_cm is not None
    ):
        text_height_twips = (
            section.page_height_cm - section.margin_top_cm - section.margin_bottom_cm
        ) * TWIPS_PER_CM

    problems = {}
    if grid_type and section.doc_grid_type not in (grid_type,):
        problems["type"] = section.doc_grid_type
    if text_height_twips and section.doc_grid_line_pitch is not None and lines_per_page > 0:
        expected_pitch = round(text_height_twips / lines_per_page)
        if abs(section.doc_grid_line_pitch - expected_pitch) > 3:
            problems["line_pitch"] = section.doc_grid_line_pitch
    if not problems:
        return []
    return [
        make_issue(
            rule,
            locator="section/1",
            message=f"文档网格与“每页 {lines_per_page} 行”不一致",
            actual=problems,
            expected=expected,
            hint="需用老师提供的 Word 模板核对；该规则为 warning 级，先以实际模板为准",
        )
    ]
