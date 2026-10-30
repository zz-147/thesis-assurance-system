"""图表/公式类 P1 检查：FIG-001/002/003、TAB-001/002/003、EQ-001。

设计原则（与项目约束一致）：

- **没有图/表/公式就整条规则跳过**：不把"缺少图""缺少表"当作格式问题（内容完整性不在此列）；
- 参数化：题注字号/字体/对齐/行距、三线表边框粗细、公式字号等一律读当前 ruleset 的
  ``expected``，缺字段时回退学校规范默认值（与 PAGE/BODY/HEAD 的写法一致）；
- 能定位的才报：题注段用 ``body/para/N「摘要」``，表格用 ``table/N``，并给出 actual/expected；
- 无法可靠判定的（如"表内数字上下对齐"、"分图标注"）不注册，见最终报告说明。
"""

from __future__ import annotations

import re
from typing import Any, Optional

from ..config import Rule, RuleSet
from ..docx_io import DocModel, ParaInfo, TableInfo
from ..report import Issue
from ..roles import Roles
from ..rule_params import expected_params
from .helpers import (
    CHAR_TOL,
    PT_TOL,
    SPACING_TOL,
    close,
    has_cjk,
    has_latin,
    indent_chars,
    loc_para,
    make_issue,
    para_brief,
)

# 题注前缀："图3-1" / "表6-2" / "图 3.1" 等（不吞掉后随空格，空格单独作为分隔符判断）
CAP_FIG_PREFIX_RE = re.compile(r"^(图\s*\d+(?:\s*[-.．－]\s*\d+)?)")
CAP_TAB_PREFIX_RE = re.compile(r"^(表\s*\d+(?:\s*[-.．－]\s*\d+)?)")
FIG_NUMBER_RE = re.compile(r"^图\s*(\d+)(?:\s*[-.．－]\s*(\d+))?")
NOTE_RE = re.compile(r"^(?:注\s*[:：]|资料来源\s*[:：]?)")
DIGITS_RE = re.compile(r"(\d+)")

FIG_FALLBACK = {
    "position": "below",
    "alignment": "center",
    "number_size_pt": 10.5,
    "number_ascii": "Times New Roman",
    "text_size_pt": 10.5,
    "text_east_asia": "楷体",
    "separator": " ",
    "line_spacing": 1.5,
}
TAB_CAPTION_FALLBACK = {
    "position": "above",
    "alignment": "center",
    "number_size_pt": 10.5,
    "number_ascii": "Times New Roman",
    "text_size_pt": 10.5,
    "text_east_asia": "楷体",
    "line_spacing": 1.5,
    "cell_size_pt": 10.5,
    "cell_east_asia": "宋体",
    "cell_alignment": "center",
    "cell_line_spacing": 1.5,
    "note_size_pt": 9,
    "note_first_line_chars": 200,
}
NOTE_FALLBACK = {
    "size_pt": 9,
    "east_asia": "宋体",
    "alignment": "left",
    "first_line_chars": 200,
    "line_spacing": 1.5,
}
TAB_THREE_LINE_FALLBACK = {
    "style": "three_line",
    "header_row_rule": "0.75pt bottom border",
    "no_vertical_borders": True,
    "no_inner_horizontal_borders": True,
    "top_bottom_border_pt": 1.5,
}
EQ_FALLBACK = {
    "own_line": True,
    "number_in_parentheses": True,
    "number_alignment": "right",
    "size_pt": 12,
    "east_asia": "宋体",
    "ascii": "Times New Roman",
    "line_spacing": 1.5,
    "numbering": {"sci_eng": "^（\\d+-\\d+）$", "humanities": "^（\\d+）$"},
}


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


def _caption_parts(raw_text: str, prefix_re: re.Pattern) -> tuple[str, str, str]:
    """把题注切成 (编号文本, 编号与题名间分隔符, 题名文本)。"""
    text = raw_text or ""
    match = prefix_re.match(text)
    if not match:
        return "", "", text.strip()
    number = match.group(1)
    rest = text[len(number):]
    sep_match = re.match(r"(\s*)", rest)
    sep = sep_match.group(1) if sep_match else ""
    title = rest[len(sep):]
    return number, sep, title


def _run_spans(para: ParaInfo) -> list[tuple[int, int, Any]]:
    spans: list[tuple[int, int, Any]] = []
    cursor = 0
    for run in para.runs:
        text = run.text or ""
        spans.append((cursor, cursor + len(text), run))
        cursor += len(text)
    return spans


def _caption_font_problems(
    para: ParaInfo,
    number: str,
    title: str,
    *,
    number_size: float,
    number_ascii: str,
    text_size: float,
    text_east_asia: str,
) -> list[str]:
    """题注字体检查：编号段 = 五号 Times New Roman；题名段 = 五号楷体（中文）。"""
    problems: list[str] = []
    if not para.runs:
        return ["题注段落没有文字 run"]
    number_end = len(number)
    for start, end, run in _run_spans(para):
        run_text = run.text or ""
        if not run_text:
            continue
        number_piece = (
            run_text[: min(len(run_text), number_end - start)] if start < number_end else ""
        )
        title_piece = (
            run_text[max(0, number_end - start):] if end > number_end else ""
        )
        if number_piece.strip():
            if not close(run.size_pt, number_size, PT_TOL):
                problems.append(
                    f"图/表号字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
                )
            if (has_latin(number_piece) or DIGITS_RE.search(number_piece)) and (
                run.ascii != number_ascii
            ):
                problems.append(f"图/表号西文字体={run.ascii or '未设置'}")
        if title_piece.strip():
            if not close(run.size_pt, text_size, PT_TOL):
                problems.append(
                    f"图/表名字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
                )
            if has_cjk(title_piece) and run.east_asia != text_east_asia:
                problems.append(f"图/表名字体={run.east_asia or '未设置'}（应为 {text_east_asia}）")
    return list(dict.fromkeys(problems))


def _format_problems(
    para: ParaInfo,
    *,
    alignment: Optional[str] = None,
    line_spacing: Optional[float] = None,
) -> list[str]:
    problems: list[str] = []
    if alignment is not None:
        actual = para.alignment or "left"
        if actual != alignment and not (alignment == "justify" and actual == "both"):
            problems.append(f"对齐={actual}")
    if line_spacing is not None and not close(para.line_spacing, line_spacing, SPACING_TOL):
        problems.append(
            f"行距={para.line_spacing if para.line_spacing is not None else '未设置'}倍"
        )
    return problems


def _figure_above_caption(model: DocModel, caption: ParaInfo) -> Optional[bool]:
    """题注上方是否紧跟图片：True=图在上；False=图在题注下方；None=找不到图片。"""
    for offset in range(1, 4):
        index = caption.index - offset
        if index < 0:
            break
        prev = model.by_index.get(index)
        if prev is None:
            continue
        if prev.has_drawing:
            return True
        if prev.is_empty:
            continue
        break
    for offset in range(1, 4):
        index = caption.index + offset
        nxt = model.by_index.get(index)
        if nxt is None:
            break
        if nxt.has_drawing:
            return False
        if nxt.is_empty:
            continue
        break
    return None


def _table_position_ok(model: DocModel, caption_index: int) -> Optional[bool]:
    """表题与表格的相对位置：True=表在上方；False=表题在表下方；None=找不到表格。"""
    for table in model.tables:
        if table.prev_para_index == caption_index:
            return True
    for table in model.tables:
        if table.next_para_index == caption_index:
            return False
    return None


def _data_tables(model: DocModel, roles: Roles) -> list[TableInfo]:
    """只把"表题 → 表格"紧邻的表格当作正文数据表（跳过封面/承诺书里的表格）。"""
    return [
        table
        for table in model.tables
        if table.prev_para_index is not None
        and roles.get(table.prev_para_index) == "caption_tab"
    ]


def _table_notes(model: DocModel, roles: Roles, table: TableInfo) -> list[ParaInfo]:
    """表格下方紧邻的"注：/资料来源："段落（跳过空段，最多找 3 段）。"""
    index = table.next_para_index
    notes: list[ParaInfo] = []
    while index is not None and len(notes) < 2:
        para = model.by_index.get(index)
        if para is None:
            break
        text = (para.text or "").strip()
        if not text:
            index = index + 1 if index + 1 < len(model.paragraphs) else None
            continue
        if NOTE_RE.match(text):
            notes.append(para)
        break
    return notes


def _figure_notes(model: DocModel, roles: Roles) -> list[ParaInfo]:
    """紧跟在图题之后的"注：/资料来源："段落。"""
    notes: list[ParaInfo] = []
    for index in roles.indices("caption_fig"):
        cursor = index + 1
        while cursor < len(model.paragraphs):
            para = model.by_index.get(cursor)
            if para is None:
                break
            text = (para.text or "").strip()
            if not text:
                cursor += 1
                continue
            if NOTE_RE.match(text):
                notes.append(para)
            break
    return notes


def _note_problems(para: ParaInfo, params: dict) -> list[str]:
    """注释段格式：小五宋体、左对齐、首行缩进 2 字符、行距（默认 1.5 倍）。"""
    problems: list[str] = []
    size = _num(params, "size_pt", 9.0)
    east_asia = _text(params, "east_asia", "宋体")
    alignment = _text(params, "alignment", "left")
    first_line = _num(params, "first_line_chars", 200.0) / 100.0
    line_spacing = _num(params, "line_spacing", 1.5)
    for run in para.runs:
        text = run.text or ""
        if not text.strip():
            continue
        if not close(run.size_pt, size, PT_TOL):
            problems.append(
                f"字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
            )
        if has_cjk(text) and run.east_asia != east_asia:
            problems.append(f"中文字体={run.east_asia or '未设置'}（应为 {east_asia}）")
    actual_alignment = para.alignment or "left"
    if alignment and actual_alignment != alignment:
        problems.append(f"对齐={actual_alignment}")
    if not close(para.line_spacing, line_spacing, SPACING_TOL):
        problems.append(
            f"行距={para.line_spacing if para.line_spacing is not None else '未设置'}倍"
        )
    chars = indent_chars(para)
    if not close(chars, first_line, CHAR_TOL):
        problems.append(
            f"首行缩进={chars:.2f}字符" if chars is not None else "首行缩进未设置"
        )
    return list(dict.fromkeys(problems))


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def check_float(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng"
) -> list[Issue]:
    issues: list[Issue] = []
    issues.extend(_check_fig_captions(model, roles, ruleset))
    issues.extend(_check_fig_numbering(model, roles, ruleset, profile))
    issues.extend(_check_table_three_line(model, roles, ruleset))
    issues.extend(_check_table_captions(model, roles, ruleset))
    issues.extend(_check_equations(model, ruleset, profile))
    issues.extend(_check_figure_notes(model, roles, ruleset))
    issues.extend(_check_table_notes(model, roles, ruleset))
    return issues


# ---------------------------------------------------------------------------
# FIG-001 图题（图号 + 图名）
# ---------------------------------------------------------------------------
def _check_fig_captions(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "FIG-001")
    captions = roles.indices("caption_fig")
    if rule is None or not captions:
        return []
    params = expected_params(ruleset, "FIG-001", FIG_FALLBACK)
    expected = {
        "position": _text(params, "position", "below"),
        "alignment": _text(params, "alignment", "center"),
        "number_size_pt": _num(params, "number_size_pt", 10.5),
        "number_ascii": _text(params, "number_ascii", "Times New Roman"),
        "text_size_pt": _num(params, "text_size_pt", 10.5),
        "text_east_asia": _text(params, "text_east_asia", "楷体"),
        "separator": _text(params, "separator", " "),
        "line_spacing": _num(params, "line_spacing", 1.5),
    }
    issues: list[Issue] = []
    for index in captions:
        para = model.by_index[index]
        number, sep, title = _caption_parts(para.text, CAP_FIG_PREFIX_RE)
        problems: list[str] = []
        if not number:
            problems.append("题注不是“图N / 图N-N + 图名”格式")
        if number and sep != expected["separator"]:
            problems.append(f"图号与图名间分隔={sep!r}（应为 {expected['separator']!r}）")
        if number and not title.strip():
            problems.append("缺少图名文字")
        problems.extend(
            _caption_font_problems(
                para,
                number,
                title,
                number_size=expected["number_size_pt"],
                number_ascii=expected["number_ascii"],
                text_size=expected["text_size_pt"],
                text_east_asia=expected["text_east_asia"],
            )
        )
        problems.extend(
            _format_problems(
                para,
                alignment=expected["alignment"],
                line_spacing=expected["line_spacing"],
            )
        )
        position = _figure_above_caption(model, para)
        if position is False:
            problems.append("图片位于图题下方或图题上方无图片（规范：图名置于图下方）")
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="图题格式不符合规范",
                    actual="；".join(dict.fromkeys(problems)),
                    expected=expected,
                    hint="Word：选中图题 → 开始 → 字体/段落；图题应在图片正下方居中",
                )
            )
    return issues


# ---------------------------------------------------------------------------
# FIG-002 图编号规则与连续性
# ---------------------------------------------------------------------------
def _numbering_problems(
    items: list[tuple[int, list[int]]],
    *,
    profile: str,
    label: str,
) -> list[tuple[int, str]]:
    """编号连续性/重复检查，返回 [(段落 index, 问题描述)]。

    - 理工农（sci_eng）：图/表号 = 章号-序号，同章内序号连续、换章后从 1 开始；
    - 文经管（humanities）：全文按出现顺序连续编号（1,2,3…）。
    """
    problems: list[tuple[int, str]] = []
    seen: set[tuple[int, ...]] = set()
    current_chapter: Optional[int] = None
    expected_number = 1
    for index, numbers in items:
        key = tuple(numbers)
        if key in seen:
            problems.append((index, f"{label}号重复：{'-'.join(map(str, numbers))}"))
            continue
        seen.add(key)
        if profile == "humanities" or len(numbers) == 1:
            number = numbers[-1]
            if number != expected_number:
                problems.append(
                    (
                        index,
                        f"{label}号不连续：实际 {number}，应为 {expected_number}",
                    )
                )
            expected_number = number + 1
        else:
            chapter, number = numbers[0], numbers[-1]
            if chapter != current_chapter:
                current_chapter = chapter
                expected_number = 1
            if number != expected_number:
                problems.append(
                    (
                        index,
                        f"{label}{chapter}-{number} 不连续：该章内应为 {chapter}-{expected_number}",
                    )
                )
            expected_number = number + 1
    return problems


def _check_fig_numbering(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str
) -> list[Issue]:
    rule = _rule(ruleset, "FIG-002")
    captions = roles.indices("caption_fig")
    if rule is None or not captions:
        return []
    params = expected_params(
        ruleset,
        "FIG-002",
        {"sci_eng": "^图\\d+-\\d+", "humanities": "^图\\d+", "continuous": True, "unique": True},
    )
    pattern_text = params.get(profile) or params.get("sci_eng") or ""
    try:
        pattern = re.compile(str(pattern_text)) if pattern_text else None
    except re.error:
        pattern = None
    expected = {
        "profile": profile,
        "numbering_pattern": pattern_text,
        "continuous": True,
        "unique": True,
    }
    issues: list[Issue] = []
    parsed: list[tuple[int, list[int]]] = []
    for index in captions:
        para = model.by_index[index]
        text = (para.text or "").strip()
        match = FIG_NUMBER_RE.match(text)
        if match is None:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="图号格式无法识别",
                    actual=para_brief(para),
                    expected=expected,
                )
            )
            continue
        numbers = [int(value) for value in match.groups() if value is not None]
        parsed.append((index, numbers))
        if pattern is not None and pattern.match(text) is None:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message=f"图号格式与当前学科规范不一致（应为 {pattern_text}）",
                    actual=match.group(0),
                    expected=pattern_text,
                )
            )
    for index, problem in _numbering_problems(parsed, profile=profile, label="图"):
        issues.append(
            make_issue(
                rule,
                locator=loc_para(model.by_index[index]),
                message="图编号不符合规范",
                actual=problem,
                expected=expected,
            )
        )
    # 全文仅 1 幅图时，规范要求图名前加“附图”
    if len(captions) == 1:
        para = model.by_index[captions[0]]
        if not (para.text or "").lstrip().startswith("附图"):
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="全文仅 1 幅图，规范要求图名前加“附图”",
                    actual=para_brief(para),
                    expected="附图1 …",
                )
            )
    return issues


# ---------------------------------------------------------------------------
# TAB-001 三线表
# ---------------------------------------------------------------------------
def _check_table_three_line(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "TAB-001")
    tables = _data_tables(model, roles)
    if rule is None or not tables:
        return []
    params = expected_params(ruleset, "TAB-001", TAB_THREE_LINE_FALLBACK)
    expected = {
        "style": _text(params, "style", "three_line"),
        "top_bottom_border_pt": _num(params, "top_bottom_border_pt", 1.5),
        "header_row_rule": _text(params, "header_row_rule", "0.75pt bottom border"),
        "no_vertical_borders": True,
        "no_inner_horizontal_borders": True,
    }
    header_border_pt = _parse_pt(expected["header_row_rule"], 0.75)
    issues: list[Issue] = []
    for table in tables:
        problems: dict[str, Any] = {}
        for key in ("left", "right", "insideV"):
            border = table.borders.get(key)
            if border is not None and border.val not in ("", "none", "nil"):
                problems[f"vertical_{key}"] = f"{border.val} {border.size_pt:g}pt"
        inside_h = table.borders.get("insideH")
        if inside_h is not None and inside_h.val not in ("", "none", "nil"):
            problems["inner_horizontal"] = f"{inside_h.val} {inside_h.size_pt:g}pt"
        for key in ("top", "bottom"):
            border = table.borders.get(key)
            if border is None:
                continue
            if border.val not in ("single", "double", "thick"):
                problems[f"{key}_border"] = f"{border.val or '未设置'}"
            elif not close(border.size_pt, expected["top_bottom_border_pt"], 0.3):
                problems[f"{key}_border_pt"] = f"{border.size_pt:g}pt"
        explicit_bottoms = [
            border for border in table.first_row_bottom_borders if border is not None
        ]
        if explicit_bottoms:
            for border in explicit_bottoms:
                if border.val not in ("single", "double") or not close(
                    border.size_pt, header_border_pt, 0.3
                ):
                    problems["header_row_bottom_border"] = (
                        f"{border.val or '未设置'} {border.size_pt:g}pt"
                    )
                    break
        elif table.borders:
            # 表级边框存在但没有首行分隔线：三线表缺中间那根线
            inside_val = (inside_h.val if inside_h is not None else "").lower()
            if inside_val in ("none", "nil", ""):
                problems["header_row_bottom_border"] = "未设置（表头下应有细实线）"
        if _has_same_text(table):
            problems["same_text"] = "表内出现“同上/同左”（规范要求逐项标注）"
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=f"table/{table.index}",
                    message="表格不是规范的三线表",
                    actual=problems,
                    expected=expected,
                    hint="三线表：上下粗线 1.5pt、表头下细线 0.75pt、无竖线/无内部横线",
                )
            )
    return issues


def _parse_pt(rule_text: str, fallback: float) -> float:
    match = re.search(r"(\d+(?:\.\d+)?)\s*pt", rule_text or "")
    try:
        return float(match.group(1)) if match else fallback
    except (TypeError, ValueError):
        return fallback


def _has_same_text(table: TableInfo) -> bool:
    for row in table.cells:
        for cell in row:
            text = (cell.text or "").strip()
            if text in ("同上", "同左", "同上。", "同左。"):
                return True
    return False


# ---------------------------------------------------------------------------
# TAB-002 表题与表内文字
# ---------------------------------------------------------------------------
def _check_table_captions(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "TAB-002")
    captions = roles.indices("caption_tab")
    tables = _data_tables(model, roles)
    if rule is None or (not captions and not tables):
        return []
    params = expected_params(ruleset, "TAB-002", TAB_CAPTION_FALLBACK)
    expected_caption = {
        "position": _text(params, "position", "above"),
        "alignment": _text(params, "alignment", "center"),
        "number_size_pt": _num(params, "number_size_pt", 10.5),
        "number_ascii": _text(params, "number_ascii", "Times New Roman"),
        "text_size_pt": _num(params, "text_size_pt", 10.5),
        "text_east_asia": _text(params, "text_east_asia", "楷体"),
        "line_spacing": _num(params, "line_spacing", 1.5),
        "cell_size_pt": _num(params, "cell_size_pt", 10.5),
        "cell_east_asia": _text(params, "cell_east_asia", "宋体"),
        "cell_alignment": _text(params, "cell_alignment", "center"),
        "cell_line_spacing": _num(params, "cell_line_spacing", 1.5),
        "note_size_pt": _num(params, "note_size_pt", 9),
        "note_first_line_chars": _num(params, "note_first_line_chars", 200),
    }
    issues: list[Issue] = []

    # ---- 表题 ----
    for index in captions:
        para = model.by_index[index]
        number, sep, title = _caption_parts(para.text, CAP_TAB_PREFIX_RE)
        problems: list[str] = []
        if not number:
            problems.append("表题不是“表N / 表N-N + 表名”格式")
        if number and not title.strip():
            problems.append("缺少表名文字")
        problems.extend(
            _caption_font_problems(
                para,
                number,
                title,
                number_size=expected_caption["number_size_pt"],
                number_ascii=expected_caption["number_ascii"],
                text_size=expected_caption["text_size_pt"],
                text_east_asia=expected_caption["text_east_asia"],
            )
        )
        problems.extend(
            _format_problems(
                para,
                alignment=expected_caption["alignment"],
                line_spacing=expected_caption["line_spacing"],
            )
        )
        position = _table_position_ok(model, index)
        if position is False:
            problems.append("表格位于表题上方（规范：表名置于表格上方）")
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="表题格式不符合规范",
                    actual="；".join(dict.fromkeys(problems)),
                    expected=expected_caption,
                    hint="Word：表题应在表格正上方居中；表号五号 Times New Roman、表名五号楷体",
                )
            )

    # ---- 表内文字 ----
    for table in tables:
        counts: dict[str, int] = {}
        sample: dict[str, str] = {}
        for row in table.cells:
            for cell in row:
                if not (cell.text or "").strip():
                    continue
                for run in cell.runs:
                    text = run.text or ""
                    if not text.strip():
                        continue
                    if not close(run.size_pt, expected_caption["cell_size_pt"], PT_TOL):
                        key = f"字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
                        counts[key] = counts.get(key, 0) + 1
                        sample.setdefault(key, cell.text[:16])
                    if has_cjk(text) and run.east_asia != expected_caption["cell_east_asia"]:
                        key = f"中文字体={run.east_asia or '未设置'}"
                        counts[key] = counts.get(key, 0) + 1
                        sample.setdefault(key, cell.text[:16])
                alignment = cell.alignment or "left"
                if alignment != expected_caption["cell_alignment"]:
                    key = f"对齐={alignment}"
                    counts[key] = counts.get(key, 0) + 1
                    sample.setdefault(key, cell.text[:16])
                if not close(
                    cell.line_spacing, expected_caption["cell_line_spacing"], SPACING_TOL
                ):
                    key = (
                        f"行距={cell.line_spacing:g}倍"
                        if cell.line_spacing is not None
                        else "行距未设置"
                    )
                    counts[key] = counts.get(key, 0) + 1
                    sample.setdefault(key, cell.text[:16])
        if counts:
            detail = "；".join(
                f"{key}（{count} 处，如“{sample[key]}”）" for key, count in counts.items()
            )
            issues.append(
                make_issue(
                    rule,
                    locator=f"table/{table.index}",
                    message="表内文字格式不符合规范",
                    actual=detail,
                    expected={
                        "size_pt": expected_caption["cell_size_pt"],
                        "east_asia": expected_caption["cell_east_asia"],
                        "alignment": expected_caption["cell_alignment"],
                        "line_spacing": expected_caption["cell_line_spacing"],
                    },
                    hint="Word：选中表格 → 表格设计/布局；表内文字五号宋体、居中、1.5 倍行距",
                )
            )

    # ---- 表注（TAB-003 未注册时由 TAB-002 兜底检查）----
    if _rule(ruleset, "TAB-003") is None:
        note_params = {
            "size_pt": expected_caption["note_size_pt"],
            "east_asia": "宋体",
            "alignment": "left",
            "first_line_chars": expected_caption["note_first_line_chars"],
            "line_spacing": 1.5,
        }
        for table in tables:
            for note in _table_notes(model, roles, table):
                problems = _note_problems(note, note_params)
                if problems:
                    issues.append(
                        make_issue(
                            rule,
                            locator=loc_para(note, "table/note"),
                            message="表注格式不符合规范",
                            actual="；".join(problems),
                            expected=note_params,
                        )
                    )
    return issues


# ---------------------------------------------------------------------------
# TAB-003 / FIG-003 注释格式
# ---------------------------------------------------------------------------
def _check_table_notes(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "TAB-003")
    tables = _data_tables(model, roles)
    if rule is None or not tables:
        return []
    params = expected_params(ruleset, "TAB-003", NOTE_FALLBACK)
    issues: list[Issue] = []
    for table in tables:
        for note in _table_notes(model, roles, table):
            problems = _note_problems(note, params)
            if problems:
                issues.append(
                    make_issue(
                        rule,
                        locator=loc_para(note, "table/note"),
                        message="表注释格式不符合规范（应位于表格下方、小五宋体、左对齐、首行缩进 2 字符）",
                        actual="；".join(problems),
                        expected={
                            "position": _text(params, "position", "below_table"),
                            "size_pt": _num(params, "size_pt", 9),
                            "east_asia": _text(params, "east_asia", "宋体"),
                            "alignment": _text(params, "alignment", "left"),
                            "first_line_chars": _num(params, "first_line_chars", 200),
                            "line_spacing": _num(params, "line_spacing", 1.5),
                        },
                    )
                )
    return issues


def _check_figure_notes(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "FIG-003")
    captions = roles.indices("caption_fig")
    notes = _figure_notes(model, roles)
    if rule is None or not captions or not notes:
        return []
    params = expected_params(ruleset, "FIG-003", NOTE_FALLBACK)
    issues: list[Issue] = []
    for note in notes:
        problems = _note_problems(note, params)
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(note, "figure/note"),
                    message="图注释格式不符合规范（应位于图名下方、小五宋体、左对齐、首行缩进 2 字符）",
                    actual="；".join(problems),
                    expected={
                        "position": _text(params, "position", "below_figure_title"),
                        "size_pt": _num(params, "size_pt", 9),
                        "east_asia": _text(params, "east_asia", "宋体"),
                        "alignment": _text(params, "alignment", "left"),
                        "first_line_chars": _num(params, "first_line_chars", 200),
                        "line_spacing": _num(params, "line_spacing", 1.5),
                    },
                )
            )
    return issues


# ---------------------------------------------------------------------------
# EQ-001 公式与编号
# ---------------------------------------------------------------------------
def _equation_parts(para: ParaInfo, pattern: Optional[re.Pattern]) -> tuple[str, list[int]]:
    """从公式段落可见文字中抽编号：返回 (编号文本, 数字列表)。"""
    text = (para.text or "").strip()
    if not text or pattern is None:
        return "", []
    matched = pattern.search(text)
    if matched is None:
        return "", []
    number_text = matched.group(0)
    return number_text, [int(value) for value in DIGITS_RE.findall(number_text)]


def _check_equations(model: DocModel, ruleset: RuleSet, profile: str) -> list[Issue]:
    rule = _rule(ruleset, "EQ-001")
    math_paras = [para for para in model.paragraphs if para.has_math]
    if rule is None or not math_paras:
        return []
    params = expected_params(ruleset, "EQ-001", EQ_FALLBACK)
    numbering = params.get("numbering") if isinstance(params.get("numbering"), dict) else {}
    pattern_text = str(numbering.get(profile) or numbering.get("sci_eng") or "")
    try:
        pattern = re.compile(pattern_text) if pattern_text else None
    except re.error:
        pattern = None
    size = _num(params, "size_pt", 12.0)
    ascii_font = _text(params, "ascii", "Times New Roman")
    line_spacing = _num(params, "line_spacing", 1.5)
    expected = {
        "number_format": pattern_text,
        "size_pt": size,
        "ascii": ascii_font,
        "line_spacing": line_spacing,
    }
    issues: list[Issue] = []
    parsed: list[tuple[int, list[int]]] = []
    for para in math_paras:
        problems: list[str] = []
        if not close(para.line_spacing, line_spacing, SPACING_TOL):
            problems.append(
                f"行距={para.line_spacing if para.line_spacing is not None else '未设置'}倍"
            )
        visible = (para.text or "").strip()
        # 公式段落的可见文字通常是编号（公式本体在 m:oMath 里）；编号候选 = 短、无汉字
        candidate = bool(visible) and not has_cjk(visible) and len(visible) <= 20
        if candidate:
            number_text, numbers = _equation_parts(para, pattern)
            if not number_text:
                problems.append(f"公式编号格式={para_brief(para)}（应为 {pattern_text}）")
            else:
                parsed.append((para.index, numbers))
                for run in para.runs:
                    if not (run.text or "").strip():
                        continue
                    if not close(run.size_pt, size, PT_TOL):
                        problems.append(
                            f"公式编号字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
                        )
                    if has_latin(run.text) and run.ascii != ascii_font:
                        problems.append(f"公式编号西文字体={run.ascii or '未设置'}")
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="公式/公式编号格式不符合规范",
                    actual="；".join(dict.fromkeys(problems)),
                    expected=expected,
                    hint="公式另起一行居中；编号用圆括号、右对齐；公式字体小四号",
                )
            )
    for index, problem in _numbering_problems(parsed, profile=profile, label="公式"):
        issues.append(
            make_issue(
                rule,
                locator=loc_para(model.by_index[index]),
                message="公式编号不符合规范",
                actual=problem,
                expected=expected,
            )
        )
    return issues
