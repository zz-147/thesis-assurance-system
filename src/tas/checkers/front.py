"""前置区检查：中英题目、中英摘要、关键词、目录标题、参考文献标题。

期望值统一从**当前 ruleset** 的 ``expected`` 读取（``rule_params.expected_params``）：
改 ``rules.json`` / 教师 profile 的字体、字号、行距、段距、对齐、关键词数量等参数，
检查结果立即生效；规则库缺字段时回退学校默认值，保证 ``school_default`` 的检查行为
与参数化改造之前完全一致。

单位约定（rules.json 原样）：``before_lines``/``after_lines`` 为 1/100 行（100 → 1.0），
``first_line_chars`` 为 1/100 字符（200 → 2.0），其余字段原样使用。
"""

from __future__ import annotations

import re
from typing import Any

from ..config import RuleSet
from ..docx_io import DocModel, ParaInfo
from ..report import Issue
from ..roles import Roles
from ..rule_params import (
    alignment,
    char_count,
    expected_params,
    font_size,
    line_count,
    text_value,
)
from .helpers import (
    CHAR_TOL,
    PT_TOL,
    SPACING_TOL,
    close,
    first_run,
    has_cjk,
    has_latin,
    indent_chars,
    is_justified,
    label_gap_spaces,
    loc_para,
    make_issue,
    para_format_problems,
    run_font_problems,
)

# ---------------------------------------------------------------------------
# 学校规范默认值（rules.json 缺失字段时兜底；rules.json 有 expected 时以它为准）
# before_lines / after_lines 沿用 rules.json 的 1/100 行单位（100 = 1 行），
# first_line_chars 沿用 1/100 字符单位（200 = 2 字符）。
# ---------------------------------------------------------------------------
TITLE_CN_FALLBACK = {
    "size_pt": 16,
    "east_asia": "黑体",
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 100,
    "after_lines": 100,
}
TITLE_EN_FALLBACK = {
    "size_pt": 16,
    "ascii": "Times New Roman",
    "bold": True,
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 100,
    "after_lines": 100,
}
ABS_CN_LABEL_FALLBACK = {**TITLE_CN_FALLBACK, "char_gap": 2}
ABS_EN_LABEL_FALLBACK = {
    "size_pt": 16,
    "ascii": "Times New Roman",
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 100,
    "after_lines": 100,
}
ABS_CN_BODY_FALLBACK = {
    "size_pt": 12,
    "east_asia": "宋体",
    "ascii": "Times New Roman",
    "first_line_chars": 200,
    "line_spacing": 1.5,
    "alignment": "justify",
}
ABS_EN_BODY_FALLBACK = {
    "size_pt": 12,
    "east_asia": "宋体",
    "ascii": "Times New Roman",
    "first_line_chars": 200,
    "line_spacing": 1.5,
    "alignment": "justify",
}
KW_CN_FALLBACK = {
    "label": "关键词",
    "label_size_pt": 12,
    "label_east_asia": "黑体",
    "body_size_pt": 12,
    "body_east_asia": "宋体",
    "body_ascii": "Times New Roman",
    "count_range": [3, 5],
    "separator": "；",
    "trailing_punctuation": False,
}
KW_EN_FALLBACK = {
    "label": "Key words",
    "label_size_pt": 12,
    "label_ascii": "Times New Roman",
    "label_bold": True,
    "body_size_pt": 12,
    "body_ascii": "Times New Roman",
    "count_range": [3, 5],
    "separator": "; ",
    "trailing_punctuation": False,
}
TOC_FALLBACK = {
    "title_size_pt": 16,
    "title_east_asia": "黑体",
    "title_alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 100,
    "after_lines": 100,
    "char_gap": 2,
    "entry_size_pt": 12,
    "level1_east_asia": "黑体",
    "others_east_asia": "宋体",
}
REF_TITLE_FALLBACK = {
    "size_pt": 16,
    "east_asia": "黑体",
    "alignment": "center",
    "line_spacing": 1.5,
    "before_lines": 100,
    "after_lines": 100,
}

#: 中文关键词默认分隔符（用于保留默认文案/解析行为）
_CN_DEFAULT_SEPARATOR = "；"
#: 英文关键词默认分隔符（分号 + 一个半角空格）
_EN_DEFAULT_SEPARATOR = "; "


# ---------------------------------------------------------------------------
# 参数读取小工具
# ---------------------------------------------------------------------------
def _number(params: dict[str, Any], key: str, fallback: float) -> float:
    try:
        return float(params.get(key, fallback))
    except (TypeError, ValueError):
        return fallback


def _integer(params: dict[str, Any], key: str, fallback: int) -> int:
    try:
        return int(params.get(key, fallback))
    except (TypeError, ValueError):
        return fallback


def _flag(params: dict[str, Any], key: str, fallback: bool) -> bool:
    value = params.get(key, fallback)
    if isinstance(value, bool):
        return value
    if value is None:
        return fallback
    return str(value).strip().lower() in ("true", "1", "yes", "是")


def _format_spec(
    ruleset: RuleSet,
    rule_id: str,
    fallback: dict[str, Any],
    *,
    size_key: str = "size_pt",
    east_key: str = "east_asia",
    ascii_key: str = "ascii",
    align_key: str = "alignment",
    bold_key: str = "bold",
) -> dict[str, Any]:
    """题目/摘要标题/目录标题/参考文献标题这类「段落整体格式」的期望值。

    ``TOC-001`` 的 expected 用 ``title_size_pt``/``title_east_asia``/``title_alignment``
    等前缀键，通过 ``size_key``/``east_key``/``align_key`` 适配；返回的 dict 键名沿用
    检查器/报告使用的通用键名（size_pt / east_asia / ascii / bold / ...）。
    """
    params = expected_params(ruleset, rule_id, fallback)
    spec: dict[str, Any] = {
        "size_pt": font_size(params, size_key, float(fallback.get(size_key, 12.0)))
    }
    if params.get(east_key) is not None:
        spec["east_asia"] = text_value(params, east_key, str(fallback.get(east_key) or ""))
    if params.get(ascii_key) is not None:
        spec["ascii"] = text_value(
            params, ascii_key, str(fallback.get(ascii_key) or "Times New Roman")
        )
    if params.get(bold_key) is not None:
        spec["bold"] = _flag(params, bold_key, bool(fallback.get(bold_key)))
    spec["alignment"] = alignment(
        params, align_key, str(fallback.get(align_key) or "center")
    )
    spec["line_spacing"] = _number(
        params, "line_spacing", float(fallback.get("line_spacing", 1.5))
    )
    spec["before_lines"] = line_count(params.get("before_lines"), 1.0)
    spec["after_lines"] = line_count(params.get("after_lines"), 1.0)
    return spec


def _format_problems(para: ParaInfo, spec: dict[str, Any]) -> list[str]:
    """按 spec 检查一个「段落整体格式」段落（题目/摘要标题/目录标题/参考文献标题）。"""
    return para_format_problems(
        para,
        size_pt=spec["size_pt"],
        east_asia=spec.get("east_asia"),
        ascii_font=spec.get("ascii"),
        bold=spec.get("bold"),
        check_bold="bold" in spec,
        alignment=spec["alignment"],
        line_spacing=spec["line_spacing"],
        before_lines=spec["before_lines"],
        after_lines=spec["after_lines"],
    )


def _abstract_body_fallback(ruleset: RuleSet, rule_id: str, fallback: dict[str, Any]) -> dict:
    """摘要正文（ABS-002/004）期望值：字体字号 + 首行缩进 + 行距 + 对齐。"""
    params = expected_params(ruleset, rule_id, fallback)
    return {
        "size_pt": font_size(params, "size_pt", float(fallback["size_pt"])),
        "east_asia": text_value(params, "east_asia", fallback["east_asia"]),
        "ascii": text_value(params, "ascii", fallback["ascii"]),
        "first_line_chars": char_count(
            params.get("first_line_chars"), float(fallback["first_line_chars"]) / 100.0
        ),
        "line_spacing": _number(
            params, "line_spacing", float(fallback["line_spacing"])
        ),
        "alignment": alignment(params, "alignment", fallback["alignment"]),
    }


def _alignment_ok(para: ParaInfo, wanted: str) -> bool:
    wanted = (wanted or "left").strip().lower()
    if wanted in ("justify", "both"):
        return is_justified(para)
    return (para.alignment or "left") == wanted


def _abstract_body_problems(para: ParaInfo, spec: dict[str, Any]) -> list[str]:
    """摘要正文检查；文案与 ``helpers.body_paragraph_problems`` 保持一致（参数化版本）。"""
    problems = run_font_problems(
        para,
        expected_size=spec["size_pt"],
        east_asia=spec["east_asia"],
        ascii_font=spec["ascii"],
    )
    if not close(para.line_spacing, spec["line_spacing"], SPACING_TOL):
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
    if not close(chars, spec["first_line_chars"], CHAR_TOL):
        problems.append(f"首行缩进={chars:.2f}字符" if chars is not None else "首行缩进未设置")
    if not _alignment_ok(para, spec["alignment"]):
        problems.append(f"对齐={para.alignment or 'left'}")
    return problems


def check_front(model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng") -> list[Issue]:
    issues: list[Issue] = []
    issues.extend(_titles(model, roles, ruleset))
    issues.extend(_abstracts(model, roles, ruleset))
    issues.extend(_keywords(model, roles, ruleset))
    issues.extend(_toc_and_references(model, roles, ruleset))
    return issues


# ---------------------------------------------------------------------------
# 题目 / 摘要
# ---------------------------------------------------------------------------
def _titles(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    issues: list[Issue] = []

    rule = ruleset.get("TITLE-001")
    spec = _format_spec(ruleset, "TITLE-001", TITLE_CN_FALLBACK)
    index = roles.first_index("title_cn")
    if index is None:
        issues.append(
            make_issue(
                rule,
                locator="front/title_cn",
                message="未找到中文题目段落（前置区结构可能不规范）",
                severity="info",
                hint="题目应位于摘要页第一行；若使用了文本框/图片做标题，自动识别会失败，需人工确认",
            )
        )
    else:
        para = model.by_index[index]
        problems = _format_problems(para, spec)
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(para, "front/title_cn"),
                    message="中文题目格式不符合规范", actual="；".join(problems),
                    expected=dict(spec),
                )
            )

    rule = ruleset.get("TITLE-002")
    spec = _format_spec(ruleset, "TITLE-002", TITLE_EN_FALLBACK)
    index = roles.first_index("title_en")
    if index is None:
        issues.append(
            make_issue(rule, locator="front/title_en", message="未找到英文题目段落", severity="info")
        )
    else:
        para = model.by_index[index]
        problems = _format_problems(para, spec)
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(para, "front/title_en"),
                    message="英文题目格式不符合规范", actual="；".join(problems),
                    expected=dict(spec),
                )
            )
    return issues


def _abstracts(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    issues: list[Issue] = []

    # 中文摘要标题（ABS-001）
    rule = ruleset.get("ABS-001")
    spec = _format_spec(ruleset, "ABS-001", ABS_CN_LABEL_FALLBACK)
    char_gap = _integer(
        expected_params(ruleset, "ABS-001", ABS_CN_LABEL_FALLBACK), "char_gap", 2
    )
    index = roles.first_index("abstract_label_cn")
    if index is None:
        issues.append(
            make_issue(rule, locator="front/abstract_label_cn", message="未找到“摘要”标题", severity="info")
        )
    else:
        para = model.by_index[index]
        problems = _format_problems(para, spec)
        gap = label_gap_spaces(para.text, "摘要")
        if gap is not None and gap != char_gap:
            problems.append(f"“摘要”字间空格={gap}（规范要求空 {char_gap} 格）")
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(para, "front/abstract_label_cn"),
                    message="中文摘要标题格式不符合规范", actual="；".join(problems),
                    expected={**spec, "char_gap": char_gap},
                )
            )

    # 英文摘要标题（ABS-003）
    rule = ruleset.get("ABS-003")
    spec = _format_spec(ruleset, "ABS-003", ABS_EN_LABEL_FALLBACK)
    index = roles.first_index("abstract_label_en")
    if index is None:
        issues.append(
            make_issue(rule, locator="front/abstract_label_en", message="未找到“Abstract”标题", severity="info")
        )
    else:
        para = model.by_index[index]
        problems = _format_problems(para, spec)
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(para, "front/abstract_label_en"),
                    message="英文摘要标题格式不符合规范", actual="；".join(problems),
                    expected=dict(spec),
                )
            )

    # 中英文摘要正文（ABS-002 / ABS-004）：期望值来自各自规则的 expected
    for role, rule_id, locator_prefix, fallback in (
        ("abstract_body_cn", "ABS-002", "front/abstract_cn", ABS_CN_BODY_FALLBACK),
        ("abstract_body_en", "ABS-004", "front/abstract_en", ABS_EN_BODY_FALLBACK),
    ):
        rule = ruleset.get(rule_id)
        body_spec = _abstract_body_fallback(ruleset, rule_id, fallback)
        for index in roles.indices(role):
            para = model.by_index[index]
            if para.is_empty:
                continue
            problems = _abstract_body_problems(para, body_spec)
            if problems:
                issues.append(
                    make_issue(
                        rule, locator=loc_para(para, locator_prefix),
                        message="摘要正文格式不符合规范", actual="；".join(problems),
                        expected=dict(body_spec),
                    )
                )
    return issues


# ---------------------------------------------------------------------------
# 关键词
# ---------------------------------------------------------------------------
def _count_range(params: dict[str, Any], fallback: list[int]) -> tuple[int, int]:
    value = params.get("count_range", fallback)
    try:
        low, high = int(value[0]), int(value[1])
    except (TypeError, ValueError, IndexError, KeyError):
        return (int(fallback[0]), int(fallback[1]))
    return (low, high) if low <= high else (int(fallback[0]), int(fallback[1]))


def _keyword_spec(
    ruleset: RuleSet, rule_id: str, fallback: dict[str, Any], *, is_en: bool
) -> dict[str, Any]:
    """关键词期望值：label/body 字体字号 + 数量范围 + 分隔符 + 末尾标点。"""
    params = expected_params(ruleset, rule_id, fallback)
    low, high = _count_range(params, fallback["count_range"])
    spec: dict[str, Any] = {
        "label": text_value(params, "label", fallback["label"]),
        "label_size_pt": font_size(params, "label_size_pt", float(fallback["label_size_pt"])),
        "body_size_pt": font_size(params, "body_size_pt", float(fallback["body_size_pt"])),
        "count_range": (low, high),
        "separator": text_value(params, "separator", fallback["separator"]),
        "trailing_punctuation": _flag(
            params, "trailing_punctuation", bool(fallback["trailing_punctuation"])
        ),
    }
    if is_en:
        spec["label_ascii"] = text_value(params, "label_ascii", fallback["label_ascii"])
        spec["label_bold"] = _flag(params, "label_bold", bool(fallback["label_bold"]))
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


def _keyword_expected(spec: dict[str, Any], *, is_en: bool) -> dict[str, Any]:
    """报告里的 expected 文案（键名与参数化之前保持一致）。"""
    return {
        "label": spec["label"],
        "size_pt": spec["label_size_pt"],
        "label_font": spec["label_ascii"] if is_en else spec["label_east_asia"],
        "count_range": list(spec["count_range"]),
        "separator": spec["separator"],
        "trailing_punctuation": spec["trailing_punctuation"],
    }


def _size_hint(size_pt: float) -> str:
    """字号文案：12pt 仍按旧文案显示「12pt/小四」。"""
    return f"{size_pt:g}pt/小四" if abs(size_pt - 12.0) <= 1e-6 else f"{size_pt:g}pt"


def _keyword_items(body_text: str, separator: str) -> list[str]:
    """按当前规范的分隔符切分关键词（中文规范全/半角分号都接受）。"""
    stripped = (separator or "").strip()
    if stripped and set(stripped) <= {";", "；"}:
        return [item for item in re.split(r"[;；]", body_text) if item.strip()]
    if stripped:
        return [item for item in re.split(re.escape(stripped), body_text) if item.strip()]
    return [item for item in re.split(r"[;；]", body_text) if item.strip()]


def _has_separator(body_text: str, separator: str) -> bool:
    stripped = (separator or "").strip()
    if stripped and set(stripped) <= {";", "；"}:
        return bool(re.search(r"[;；]", body_text))
    return bool(stripped and stripped in body_text)


def _keyword_run_problems(para: ParaInfo, spec: dict[str, Any], *, is_en: bool) -> list[str]:
    """关键词行字体/字号（label 与 body 分开，与参数化之前判定一致）。"""
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

    if label_run is not None:
        if not close(label_run.size_pt, spec["label_size_pt"], PT_TOL):
            problems.append(
                f"“{'Key words' if is_en else '关键词'}”字号={label_run.size_pt}"
                f"（应为 {_size_hint(spec['label_size_pt'])}）"
            )
        if is_en:
            if label_run.ascii != spec["label_ascii"]:
                problems.append(
                    f"“Key words”字体={label_run.ascii or '未设置'}"
                    f"（应为 {spec['label_ascii']}）"
                )
            if bool(label_run.bold) != spec["label_bold"]:
                problems.append(
                    "“Key words”应加粗" if spec["label_bold"] else "“Key words”不应加粗"
                )
        elif label_run.east_asia != spec["label_east_asia"]:
            problems.append(
                f"“关键词”字体={label_run.east_asia or '未设置'}"
                f"（应为{spec['label_east_asia']}）"
            )

    for run in para.runs:
        if run is label_run or not (run.text or "").strip():
            continue
        text_run = run.text
        if not close(run.size_pt, spec["body_size_pt"], PT_TOL):
            problems.append(
                f"关键词字号={run.size_pt}（应为 {_size_hint(spec['body_size_pt'])}）"
            )
        if is_en:
            if has_latin(text_run) and run.ascii != spec["body_ascii"]:
                problems.append(
                    f"关键词字体={run.ascii or '未设置'}（应为 {spec['body_ascii']}）"
                )
        else:
            if has_cjk(text_run) and run.east_asia != spec["body_east_asia"]:
                problems.append(
                    f"关键词中文字体={run.east_asia or '未设置'}"
                    f"（应为{spec['body_east_asia']}）"
                )
            if has_latin(text_run) and run.ascii != spec["body_ascii"]:
                problems.append(
                    f"关键词西文字体={run.ascii or '未设置'}"
                    f"（应为 {spec['body_ascii']}）"
                )
    return problems


def _kw_cn_problems(para: ParaInfo, spec: dict[str, Any]) -> list[str]:
    problems: list[str] = []
    text = para.text or ""
    match = re.match(r"^\s*(关键词|关键字)\s*[:：]\s*(.*)$", text, re.S)
    if not match:
        problems.append("行格式应为“关键词：词1；词2；…”（缺少“关键词：”标签或冒号）")
        body_text = ""
    else:
        body_text = match.group(2).strip()
        low, high = spec["count_range"]
        if spec["separator"] == _CN_DEFAULT_SEPARATOR:
            items = [item for item in re.split(r"[;；]", body_text) if item.strip()]
        else:
            items = _keyword_items(body_text, spec["separator"])
        if not low <= len(items) <= high:
            problems.append(f"关键词数量={len(items)}（应为 {low}~{high} 个）")
        if not spec["trailing_punctuation"] and re.search(r"[;；。，,、]\s*$", body_text):
            problems.append("最后一个关键词后不应有标点")
        if spec["separator"] == _CN_DEFAULT_SEPARATOR:
            if len(items) > 1 and not re.search(r"[;；]", body_text):
                problems.append("关键词之间应用分号（；）分隔")
            if re.search(r"[，,]", body_text):
                problems.append("关键词之间不应使用逗号")
        else:
            if len(items) > 1 and not _has_separator(body_text, spec["separator"]):
                problems.append(f"关键词之间应用“{spec['separator']}”分隔")

    problems.extend(_keyword_run_problems(para, spec, is_en=False))
    return list(dict.fromkeys(problems))


def _kw_en_problems(para: ParaInfo, spec: dict[str, Any]) -> list[str]:
    problems: list[str] = []
    text = para.text or ""
    match = re.match(r"^\s*Key\s*words?\s*[:：]\s*(.*)$", text, re.S | re.IGNORECASE)
    if not match:
        problems.append("行格式应为“Key words: word1; word2; …”")
        body_text = ""
    else:
        body_text = match.group(1).strip()
        low, high = spec["count_range"]
        if spec["separator"] == _EN_DEFAULT_SEPARATOR:
            items = [item for item in re.split(r"[;；]", body_text) if item.strip()]
        else:
            items = _keyword_items(body_text, spec["separator"])
        if not low <= len(items) <= high:
            problems.append(f"关键词数量={len(items)}（应为 {low}~{high} 个）")
        if not spec["trailing_punctuation"] and re.search(r"[;；。，,]\s*$", body_text):
            problems.append("最后一个关键词后不应有标点")
        if spec["separator"] == _EN_DEFAULT_SEPARATOR:
            if re.search(r"[;；](?!\s)", body_text):
                problems.append("英文关键词之间应用“; ”（分号+空格）分隔")
        else:
            if len(items) > 1 and not _has_separator(body_text, spec["separator"]):
                problems.append(f"英文关键词之间应用“{spec['separator']}”分隔")

    problems.extend(_keyword_run_problems(para, spec, is_en=True))
    return list(dict.fromkeys(problems))


def _keywords(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    issues: list[Issue] = []

    rule = ruleset.get("KW-001")
    spec = _keyword_spec(ruleset, "KW-001", KW_CN_FALLBACK, is_en=False)
    index = roles.first_index("keywords_cn")
    if index is None:
        issues.append(
            make_issue(rule, locator="front/keywords_cn", message="未找到中文关键词行", severity="info")
        )
    else:
        problems = _kw_cn_problems(model.by_index[index], spec)
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(model.by_index[index], "front/keywords_cn"),
                    message="中文关键词格式不符合规范", actual="；".join(problems),
                    expected=_keyword_expected(spec, is_en=False),
                )
            )

    rule = ruleset.get("KW-002")
    spec = _keyword_spec(ruleset, "KW-002", KW_EN_FALLBACK, is_en=True)
    index = roles.first_index("keywords_en")
    if index is None:
        issues.append(
            make_issue(rule, locator="front/keywords_en", message="未找到英文关键词行", severity="info")
        )
    else:
        problems = _kw_en_problems(model.by_index[index], spec)
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(model.by_index[index], "front/keywords_en"),
                    message="英文关键词格式不符合规范", actual="；".join(problems),
                    expected=_keyword_expected(spec, is_en=True),
                )
            )
    return issues


# ---------------------------------------------------------------------------
# 目录标题 / 参考文献标题
# ---------------------------------------------------------------------------
_TOC_STYLE_RE = re.compile(r"toc\s*(\d+)", re.IGNORECASE)


def _toc_level(para: ParaInfo) -> int | None:
    match = _TOC_STYLE_RE.search(para.style_name or "")
    return int(match.group(1)) if match else None


def _toc_and_references(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    issues: list[Issue] = []

    rule = ruleset.get("TOC-001")
    params = expected_params(ruleset, "TOC-001", TOC_FALLBACK)
    title_spec = _format_spec(
        ruleset,
        "TOC-001",
        TOC_FALLBACK,
        size_key="title_size_pt",
        east_key="title_east_asia",
        align_key="title_alignment",
    )
    char_gap = _integer(params, "char_gap", 2)
    index = roles.first_index("toc_title")
    if index is None:
        issues.append(
            make_issue(rule, locator="front/toc_title", message="未找到“目录”标题", severity="info")
        )
    else:
        para = model.by_index[index]
        problems = _format_problems(para, title_spec)
        gap = label_gap_spaces(para.text, "目录")
        if gap is not None and gap != char_gap:
            problems.append(f"“目录”字间空格={gap}（规范要求空 {char_gap} 格）")
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(para, "front/toc_title"),
                    message="目录标题格式不符合规范", actual="；".join(problems),
                    expected=dict(title_spec),
                )
            )

    # 目录条目：按级别聚合报告，避免几十条重复（条目字体/字号同样来自 expected）
    entry_size = font_size(params, "entry_size_pt", 12.0)
    level1_font = text_value(params, "level1_east_asia", "黑体")
    others_font = text_value(params, "others_east_asia", "宋体")
    entry_problems: dict[int, list[str]] = {}
    first_paras: dict[int, ParaInfo] = {}
    for index in roles.indices("toc_entry"):
        para = model.by_index[index]
        level = _toc_level(para) or 2
        run = first_run(para)
        problems: list[str] = []
        if run is not None:
            if not close(run.size_pt, entry_size, PT_TOL):
                problems.append(f"字号={run.size_pt}（应为 {_size_hint(entry_size)}）")
            expected_font = level1_font if level == 1 else others_font
            if run.east_asia != expected_font:
                problems.append(f"字体={run.east_asia or '未设置'}（应为 {expected_font}）")
        if problems:
            entry_problems.setdefault(level, []).extend(problems)
            first_paras.setdefault(level, para)
    for level, problems in sorted(entry_problems.items()):
        issues.append(
            make_issue(
                rule, locator=loc_para(first_paras[level], "front/toc_entry"),
                message=f"{level} 级目录条目字体/字号不符合规范（同类问题已合并）",
                actual="；".join(dict.fromkeys(problems)),
                expected={
                    "size_pt": entry_size,
                    "east_asia": level1_font if level == 1 else others_font,
                },
            )
        )

    rule = ruleset.get("REF-001")
    title_spec = _format_spec(ruleset, "REF-001", REF_TITLE_FALLBACK)
    index = roles.first_index("refs_title")
    if index is None:
        issues.append(
            make_issue(rule, locator="front/refs_title", message="未找到“参考文献”标题", severity="info")
        )
    else:
        para = model.by_index[index]
        problems = _format_problems(para, title_spec)
        if problems:
            issues.append(
                make_issue(
                    rule, locator=loc_para(para, "front/refs_title"),
                    message="参考文献标题格式不符合规范", actual="；".join(problems),
                    expected=dict(title_spec),
                )
            )
    return issues
