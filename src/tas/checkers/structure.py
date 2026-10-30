"""结构与全局检查：TOC-002 / STRUCT-001 / MISC-001 / HEAD-004 / HEAD-005 /
APP-001 / APP-002 / TRANS-001。

共同原则：

- **只在文档确实包含相应部件时检查**：没有目录条目就不查目录完整性；识别不到封面
  就不对"装订顺序"下结论（节选/单章文档不该被当整本论文报缺件）；
- 参数化：目录层级、数字用法开关、四级标题字号、附录标题字号、外文翻译字号等
  一律读当前 ruleset 的 ``expected``（缺字段回退学校默认）；
- 目录完整性匹配"归一化文字 + 顺序 + 层级"，页码只校验"是否标注"（真实页码需要
  渲染引擎才能核对，纯 OOXML 做不到，报告里如实说明）；
- 装订顺序里"封底"无法从正文可靠识别，不参与缺件判断。
"""

from __future__ import annotations

import re
from typing import Any, Callable, Optional

from ..config import Rule, RuleSet
from ..docx_io import DocModel, ParaInfo
from ..report import Issue
from ..roles import Roles, norm
from ..rule_params import char_count, expected_params, line_count
from .helpers import (
    CHAR_TOL,
    close,
    has_cjk,
    indent_chars,
    label_gap_spaces,
    loc_para,
    make_issue,
    para_format_problems,
)

TOC_FALLBACK = {
    "includes_levels": [1, 2, 3],
    "matches_headings": True,
    "page_numbers_updateable": True,
}
STRUCT_FALLBACK = {
    "order": [
        "封面",
        "诚信承诺书",
        "中文摘要",
        "英文摘要",
        "目录",
        "正文",
        "致谢",
        "参考文献",
        "附录",
        "外文资料翻译",
        "封底",
    ]
}
MISC_FALLBACK = {"year_full": True, "significant_digits": True}
HEAD004_FALLBACK = {
    "l4": {"pattern": "^\\d+\\.\\d+\\.\\d+\\.\\d+\\s*.+$", "size_pt": 12, "east_asia": "宋体", "bold": True},
    "l5_plus": {
        "patterns": ["^（\\d+）", "^[①-⑳]", "^[A-Z][.、]", "^[a-z][.、]"],
        "size_pt": 12,
        "east_asia": "宋体",
    },
    "before_lines": 0,
    "after_lines": 0,
}
HEAD005_FALLBACK = {
    "l1": {
        "pattern": "^[一二三四五六七八九十]+、",
        "size_pt": 16,
        "east_asia": "黑体",
        "first_line_chars": 200,
        "before_lines": 100,
        "after_lines": 0,
    },
    "l2": {
        "pattern": "^（[一二三四五六七八九十]+）",
        "size_pt": 14,
        "east_asia": "黑体",
        "first_line_chars": 200,
        "before_lines": 50,
        "after_lines": 0,
    },
    "l3": {
        "pattern": "^\\d+[.、]",
        "size_pt": 12,
        "east_asia": "黑体",
        "first_line_chars": 200,
        "before_lines": 0,
        "after_lines": 0,
    },
}
APPENDIX_FALLBACK = {
    "title_text": "附录",
    "title_size_pt": 16,
    "title_east_asia": "黑体",
    "title_alignment": "center",
    "title_char_gap": 2,
    "subtitle_size_pt": 14,
    "subtitle_east_asia": "黑体",
    "subtitle_alignment": "center",
    "content_format_same_as_body": True,
}
APPENDIX_NUMBER_FALLBACK = {"multiple_appendix_numbering": True, "pattern": "^附录[A-Z]"}
TRANS_FALLBACK = {
    "title_text": "外文资料译文",
    "title_size_pt": 16,
    "title_east_asia": "黑体",
    "title_alignment": "center",
    "original_text_size_pt": 12,
    "original_text_ascii": "Times New Roman",
    "translation_text_size_pt": 12,
    "translation_text_east_asia": "宋体",
    "new_page_for_translation": True,
}

TOC_PAGE_RE = re.compile(r"^(?P<text>.*?)\t\s*(?P<page>[0-9]+|[IVXLCDMivxlcdm]+)\s*$", re.S)
SHORT_YEAR_RE = re.compile(r"(?<!\d)(\d{2})年")
CN_CENTURY_RE = re.compile(r"[一二三四五六七八九十百]{1,4}世纪")
TOC_STYLE_RE = re.compile(r"toc\s*(\d+)", re.IGNORECASE)
CJK_RE = re.compile(r"[\u4e00-\u9fff]")

_EXCLUDED_ROLES = {
    "toc_entry",
    "reference_entry",
    "caption_fig",
    "caption_tab",
    "cover",
    "title_cn",
    "title_en",
    "abstract_label_cn",
    "abstract_label_en",
    "keywords_cn",
    "keywords_en",
    "toc_title",
    "refs_title",
    "thanks_title",
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


def _next_nonempty(model: DocModel, index: int, limit: int = 3) -> Optional[ParaInfo]:
    for cursor in range(index + 1, min(index + 1 + limit, len(model.paragraphs))):
        para = model.by_index.get(cursor)
        if para is not None and not para.is_empty:
            return para
    return None


def _toc_style_level(para: ParaInfo) -> Optional[int]:
    match = TOC_STYLE_RE.search(para.style_name or "")
    return int(match.group(1)) if match else None


def _toc_entry_parts(para: ParaInfo) -> tuple[str, Optional[str]]:
    text = para.text or ""
    match = TOC_PAGE_RE.match(text)
    if match:
        return match.group("text").strip(), match.group("page")
    return text.strip(), None


def _body_items(model: DocModel, roles: Roles) -> list[tuple[str, int]]:
    """正文标题序列（一级/二级/三级 + 前言结论致谢参考文献附录），用于目录比对。"""
    items: list[tuple[str, int]] = []
    for index in sorted(roles.by_index):
        role = roles.get(index)
        para = model.by_index.get(index)
        if para is None or para.is_empty:
            continue
        if role in ("heading_1", "heading_2", "heading_3"):
            items.append((norm(para.text), int(role[-1])))
        elif role in ("special_title", "thanks_title", "refs_title", "appendix_title"):
            items.append((norm(para.text), 1))
    return items


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def check_structure(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng"
) -> list[Issue]:
    issues: list[Issue] = []
    issues.extend(_check_toc(model, roles, ruleset))
    issues.extend(_check_binding_order(model, roles, ruleset))
    issues.extend(_check_number_usage(model, roles, ruleset))
    issues.extend(_check_heading_level4(model, roles, ruleset, profile))
    issues.extend(_check_heading_humanities(model, roles, ruleset, profile))
    issues.extend(_check_appendix_title(model, roles, ruleset))
    issues.extend(_check_appendix_numbering(model, roles, ruleset))
    issues.extend(_check_translation(model, roles, ruleset))
    return issues


# ---------------------------------------------------------------------------
# TOC-002 目录完整性
# ---------------------------------------------------------------------------
def _check_toc(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "TOC-002")
    toc_indices = roles.indices("toc_entry")
    if rule is None or not toc_indices:
        return []
    params = expected_params(ruleset, "TOC-002", TOC_FALLBACK)
    includes_levels = [
        int(value) for value in (params.get("includes_levels") or TOC_FALLBACK["includes_levels"])
        if str(value).isdigit()
    ]
    expected = {
        "includes_levels": includes_levels,
        "matches_headings": True,
        "page_numbers_updateable": True,
    }
    issues: list[Issue] = []

    # ---- 1) 是否 TOC 域（可在 Word 中更新）----
    if not model.has_toc_field:
        issues.append(
            make_issue(
                rule,
                locator="front/toc",
                message="目录不是 Word 的 TOC 域（手工目录），标题/页码无法自动更新",
                actual="未检测到 TOC 域",
                expected="TOC 域（可用 F9/右键“更新域”刷新）",
                hint="Word：引用 → 目录 → 自动目录；手工目录需重建为 TOC 域",
            )
        )

    body_items = _body_items(model, roles)
    actual_items: list[tuple[str, Optional[int], int]] = []
    missing_pages: list[int] = []
    for index in toc_indices:
        para = model.by_index[index]
        text, page = _toc_entry_parts(para)
        actual_items.append((norm(text), _toc_style_level(para), index))
        if page is None:
            missing_pages.append(index)

    if body_items:
        expected_texts = [text for text, _level in body_items]
        actual_texts = [text for text, _level, _index in actual_items]
        if actual_texts != expected_texts:
            from collections import Counter

            expected_counter = Counter(expected_texts)
            actual_counter = Counter(actual_texts)
            missing = [
                text
                for text in expected_texts
                if actual_counter.get(text, 0) < expected_counter.get(text, 0)
            ]
            extra = [
                text
                for text in actual_texts
                if expected_counter.get(text, 0) < actual_counter.get(text, 0)
            ]
            problems: dict[str, Any] = {}
            if missing:
                problems["缺少目录条目"] = [item[:24] for item in dict.fromkeys(missing)]
            if extra:
                problems["目录条目在正文中找不到"] = [
                    item[:24] for item in dict.fromkeys(extra)
                ]
            if not missing and not extra and actual_texts != expected_texts:
                problems["目录条目顺序与正文标题顺序不一致"] = True
            if problems:
                issues.append(
                    make_issue(
                        rule,
                        locator="front/toc",
                        message="目录条目与正文标题不一致（目录完整性）",
                        actual=problems,
                        expected={
                            "headings": [item[:24] for item in expected_texts[:40]],
                            "note": "按正文标题顺序排列",
                        },
                        hint="Word：在目录上右键 → 更新域 → 更新整个目录",
                    )
                )
        # 层级一致性
        expected_level_by_text: dict[str, int] = {}
        for text, level in body_items:
            expected_level_by_text.setdefault(text, level)
        level_problems: list[str] = []
        for text, actual_level, index in actual_items:
            expected_level = expected_level_by_text.get(text)
            if expected_level is None or actual_level is None:
                continue
            if actual_level != expected_level:
                level_problems.append(
                    f"{model.by_index[index].text[:20].strip()}（目录 {actual_level} 级 / 正文 {expected_level} 级）"
                )
        if level_problems:
            issues.append(
                make_issue(
                    rule,
                    locator="front/toc",
                    message="目录条目的层级与正文标题层级不一致",
                    actual=level_problems[:10],
                    expected="目录层级与标题级别一致",
                )
            )
        # 级别覆盖
        body_levels = {level for _text, level in body_items}
        toc_levels = {level for _text, level, _index in actual_items if level is not None}
        if toc_levels:
            missing_levels = [
                level for level in includes_levels if level in body_levels and level not in toc_levels
            ]
            if missing_levels:
                issues.append(
                    make_issue(
                        rule,
                        locator="front/toc",
                        message=f"目录未包含以下级别的标题：{missing_levels}",
                        actual=sorted(toc_levels),
                        expected={"includes_levels": includes_levels},
                    )
                )
    if missing_pages:
        issues.append(
            make_issue(
                rule,
                locator=loc_para(model.by_index[missing_pages[0]], "front/toc_entry"),
                message=f"有 {len(missing_pages)} 条目录条目没有标注页码",
                actual=[model.by_index[index].text[:24].strip() for index in missing_pages[:10]],
                expected="每条目录条目以制表符 + 页码结尾",
            )
        )
    return issues


# ---------------------------------------------------------------------------
# STRUCT-001 装订顺序完整
# ---------------------------------------------------------------------------
def _detect_components(model: DocModel, roles: Roles) -> dict[str, Optional[int]]:
    skip_roles = {"toc_entry", "reference_entry", "caption_fig", "caption_tab"}

    def first_text(predicate: Callable[[str], bool]) -> Optional[int]:
        for para in model.paragraphs:
            if para.is_empty or roles.get(para.index) in skip_roles:
                continue
            if predicate(norm(para.text)):
                return para.index
        return None

    body_start = roles.body_start
    if body_start is not None and body_start >= len(model.paragraphs):
        body_start = None
    appendix = roles.first_index("appendix_title")
    return {
        "封面": roles.first_index("cover"),
        "诚信承诺书": first_text(lambda text: "诚信承诺书" in text),
        "中文摘要": roles.first_index("abstract_label_cn"),
        "英文摘要": roles.first_index("abstract_label_en"),
        "目录": roles.first_index("toc_title"),
        "正文": body_start,
        "致谢": roles.first_index("thanks_title"),
        "参考文献": roles.first_index("refs_title"),
        "附录": appendix,
        "外文资料翻译": first_text(
            lambda text: text.startswith(("外文资料译文", "外文资料翻译", "外文译文"))
        ),
        "封底": None,  # 无法从正文可靠识别，不参与缺件判断
    }


def _check_binding_order(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "STRUCT-001")
    if rule is None:
        return []
    params = expected_params(ruleset, "STRUCT-001", STRUCT_FALLBACK)
    order = [str(item) for item in (params.get("order") or STRUCT_FALLBACK["order"])]
    detected = _detect_components(model, roles)
    # 只有"像一本完整论文"时才下缺件结论：识别不到封面、或只有封面没有正文骨架
    # （例如单章节选、模板骨架）时跳过，避免把局部文档误判为缺件。
    core_components = ("中文摘要", "英文摘要", "目录", "正文", "致谢", "参考文献")
    if detected.get("封面") is None:
        return []
    if sum(1 for label in core_components if detected.get(label) is not None) < 2:
        return []
    expected = {"order": order}
    issues: list[Issue] = []
    for label in order:
        if label == "封底":
            continue
        if detected.get(label) is None:
            issues.append(
                make_issue(
                    rule,
                    locator=f"structure/order/{label}",
                    message=f"未找到装订顺序中的“{label}”部分",
                    actual="缺失",
                    expected={"order": order},
                    hint="规范装订顺序：封面→诚信承诺书→中英文摘要→目录→正文→致谢→参考文献→附录→外文资料翻译→封底",
                )
            )
    positions = [
        (detected[label], label)
        for label in order
        if label != "封底" and detected.get(label) is not None
    ]
    for (prev_pos, prev_label), (pos, label) in zip(positions, positions[1:]):
        if pos < prev_pos:
            issues.append(
                make_issue(
                    rule,
                    locator=f"structure/order/{label}",
                    message=f"“{label}”出现在“{prev_label}”之前，与装订顺序不一致",
                    actual=[prev_label, label],
                    expected={"order": order},
                )
            )
    return issues


# ---------------------------------------------------------------------------
# MISC-001 数字用法
# ---------------------------------------------------------------------------
def _check_number_usage(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "MISC-001")
    if rule is None:
        return []
    params = expected_params(ruleset, "MISC-001", MISC_FALLBACK)
    check_year = bool(params.get("year_full", True))
    check_century = bool(params.get("significant_digits", True))
    target_roles = (
        "body",
        "heading_1",
        "heading_2",
        "heading_3",
        "heading_4",
        "special_title",
        "abstract_body_cn",
        "abstract_body_en",
    )
    issues: list[Issue] = []
    for index in sorted(roles.by_index):
        if roles.get(index) not in target_roles:
            continue
        para = model.by_index[index]
        text = para.text or ""
        problems: list[str] = []
        if check_year:
            for match in SHORT_YEAR_RE.finditer(text):
                problems.append(
                    f"年份简写“{match.group(0)}”（上下文：{_context(text, match.start(), match.end())}）"
                )
        if check_century:
            for match in CN_CENTURY_RE.finditer(text):
                problems.append(
                    f"世纪用汉字“{match.group(0)}”（规范：公历世纪用阿拉伯数字）"
                )
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="数字用法需要人工确认（信息级，不自动修改）",
                    actual="；".join(problems[:6]),
                    expected={"year_full": True, "significant_digits": True},
                    hint="公历世纪/年代/年月日/计数计量用阿拉伯数字；年份写全（如 2024 年，不写 24 年）",
                )
            )
    return issues


def _context(text: str, start: int, end: int, width: int = 8) -> str:
    return text[max(0, start - width): min(len(text), end + width)].replace("\n", " ")


# ---------------------------------------------------------------------------
# HEAD-004 四级及以下标题
# ---------------------------------------------------------------------------
def _check_heading_level4(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str
) -> list[Issue]:
    rule = _rule(ruleset, "HEAD-004")
    if rule is None or profile != "sci_eng":
        return []
    params = expected_params(ruleset, "HEAD-004", HEAD004_FALLBACK)
    l4 = params.get("l4") if isinstance(params.get("l4"), dict) else HEAD004_FALLBACK["l4"]
    l5 = (
        params.get("l5_plus")
        if isinstance(params.get("l5_plus"), dict)
        else HEAD004_FALLBACK["l5_plus"]
    )
    before = line_count(params.get("before_lines"), 0.0)
    after = line_count(params.get("after_lines"), 0.0)
    issues: list[Issue] = []

    # 四级标题（N.N.N.N）
    l4_size = _num(l4, "size_pt", 12.0)
    l4_font = _text(l4, "east_asia", "宋体")
    l4_bold = bool(l4.get("bold", True))
    expected4 = {
        "size_pt": l4_size,
        "east_asia": l4_font,
        "bold": l4_bold,
        "line_spacing": 1.5,
        "before_lines": before,
        "after_lines": after,
    }
    for index in roles.indices("heading_4"):
        para = model.by_index[index]
        problems = para_format_problems(
            para,
            size_pt=l4_size,
            east_asia=l4_font,
            bold=l4_bold,
            check_bold=True,
            alignment="left",
            line_spacing=1.5,
            before_lines=before,
            after_lines=after,
        )
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="四级标题格式不符合规范（小四号宋体加粗、左起顶格、段前后无间距）",
                    actual="；".join(problems),
                    expected=expected4,
                )
            )

    # 四级以下标题（（1）→①→A.→a.）
    patterns = l5.get("patterns") or HEAD004_FALLBACK["l5_plus"]["patterns"]
    compiled: list[re.Pattern] = []
    for pattern in patterns:
        try:
            compiled.append(re.compile(str(pattern)))
        except re.error:
            continue
    if not compiled:
        return issues
    l5_size = _num(l5, "size_pt", 12.0)
    l5_font = _text(l5, "east_asia", "宋体")
    expected5 = {
        "size_pt": l5_size,
        "east_asia": l5_font,
        "line_spacing": 1.5,
        "before_lines": before,
        "after_lines": after,
    }
    for para in model.paragraphs:
        role = roles.get(para.index)
        if role in _EXCLUDED_ROLES or role == "heading_4" or para.is_empty:
            continue
        text_n = norm(para.text)
        if len(text_n) > 30 or text_n.endswith(("。", "；", "：", "，", ",")):
            continue
        if not any(pattern.match(text_n) for pattern in compiled):
            continue
        problems = para_format_problems(
            para,
            size_pt=l5_size,
            east_asia=l5_font,
            alignment="left",
            line_spacing=1.5,
            before_lines=before,
            after_lines=after,
        )
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="四级以下标题格式不符合规范（小四号宋体、左起顶格、段前后无间距）",
                    actual="；".join(problems),
                    expected=expected5,
                )
            )
    return issues


# ---------------------------------------------------------------------------
# HEAD-005 文经管标题体系
# ---------------------------------------------------------------------------
def _check_heading_humanities(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str
) -> list[Issue]:
    rule = _rule(ruleset, "HEAD-005")
    if rule is None or profile != "humanities":
        return []
    params = expected_params(ruleset, "HEAD-005", HEAD005_FALLBACK)
    issues: list[Issue] = []
    for config_key, label in (("l1", "一级"), ("l2", "二级"), ("l3", "三级")):
        config = params.get(config_key)
        if not isinstance(config, dict):
            config = HEAD005_FALLBACK[config_key]
        pattern_text = _text(config, "pattern", "")
        try:
            pattern = re.compile(pattern_text) if pattern_text else None
        except re.error:
            pattern = None
        if pattern is None:
            continue
        size = _num(config, "size_pt", 12.0)
        east_asia = _text(config, "east_asia", "黑体")
        before = line_count(config.get("before_lines"), 0.0)
        after = line_count(config.get("after_lines"), 0.0)
        wanted_indent = char_count(config.get("first_line_chars"), 2.0)
        expected = {
            "pattern": pattern_text,
            "size_pt": size,
            "east_asia": east_asia,
            "first_line_chars": wanted_indent,
            "before_lines": before,
            "after_lines": after,
        }
        for para in model.paragraphs:
            if para.is_empty or roles.get(para.index) in _EXCLUDED_ROLES:
                continue
            text_n = norm(para.text)
            if len(text_n) > 40 or pattern.match(text_n) is None:
                continue
            problems = para_format_problems(
                para,
                size_pt=size,
                east_asia=east_asia,
                alignment="left",
                before_lines=before,
                after_lines=after,
            )
            chars = indent_chars(para)
            if not close(chars, wanted_indent, CHAR_TOL):
                problems.append(
                    f"首行缩进={chars:.2f}字符" if chars is not None else "首行缩进未设置"
                )
            if problems:
                issues.append(
                    make_issue(
                        rule,
                        locator=loc_para(para),
                        message=f"文经管{label}标题格式/编号体系不符合规范",
                        actual="；".join(dict.fromkeys(problems)),
                        expected=expected,
                    )
                )
    return issues


# ---------------------------------------------------------------------------
# APP-001 附录标题格式
# ---------------------------------------------------------------------------
def _check_appendix_title(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "APP-001")
    indices = roles.indices("appendix_title")
    if rule is None or not indices:
        return []
    params = expected_params(ruleset, "APP-001", APPENDIX_FALLBACK)
    title_text = _text(params, "title_text", "附录")
    title_size = _num(params, "title_size_pt", 16.0)
    title_font = _text(params, "title_east_asia", "黑体")
    title_align = _text(params, "title_alignment", "center")
    char_gap = int(_num(params, "title_char_gap", 2))
    subtitle_size = _num(params, "subtitle_size_pt", 14.0)
    subtitle_font = _text(params, "subtitle_east_asia", "黑体")
    subtitle_align = _text(params, "subtitle_alignment", "center")
    expected = {
        "title_text": title_text,
        "title_size_pt": title_size,
        "title_east_asia": title_font,
        "title_alignment": title_align,
        "title_char_gap": char_gap,
        "subtitle_size_pt": subtitle_size,
        "subtitle_east_asia": subtitle_font,
        "subtitle_alignment": subtitle_align,
    }
    issues: list[Issue] = []
    for index in indices:
        para = model.by_index[index]
        text_n = norm(para.text)
        problems = para_format_problems(
            para,
            size_pt=title_size,
            east_asia=title_font,
            alignment=title_align,
        )
        if text_n == title_text:
            gap = label_gap_spaces(para.text, title_text)
            if gap is not None and gap != char_gap:
                problems.append(f"“{title_text}”字间空格={gap}（应为 {char_gap}）")
            subtitle = _next_nonempty(model, index)
            if subtitle is None:
                problems.append("下一行缺少附录题目")
            else:
                sub_problems = para_format_problems(
                    subtitle,
                    size_pt=subtitle_size,
                    east_asia=subtitle_font,
                    alignment=subtitle_align,
                )
                if sub_problems:
                    problems.append("附录题目：" + "；".join(sub_problems))
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="附录标题格式不符合规范（“附录”三号黑体居中、字间空 2 格；下一行附录题目四号黑体居中）",
                    actual="；".join(dict.fromkeys(problems)),
                    expected=expected,
                )
            )
    return issues


# ---------------------------------------------------------------------------
# APP-002 多项附录编号
# ---------------------------------------------------------------------------
def _check_appendix_numbering(
    model: DocModel, roles: Roles, ruleset: RuleSet
) -> list[Issue]:
    rule = _rule(ruleset, "APP-002")
    indices = roles.indices("appendix_title")
    if rule is None or len(indices) < 2:
        return []
    params = expected_params(ruleset, "APP-002", APPENDIX_NUMBER_FALLBACK)
    pattern_text = _text(params, "pattern", "^附录[A-Z]")
    try:
        pattern = re.compile(pattern_text)
    except re.error:
        pattern = re.compile(r"^附录[A-Z]")
    texts = [norm(model.by_index[index].text) for index in indices]
    problems: list[str] = []
    for index, text in zip(indices, texts):
        if pattern.match(text) is None:
            problems.append(f"第 {indices.index(index) + 1} 个附录未按 A/B 编号：{text[:16]}")
    letters: list[str] = []
    for text in texts:
        match = re.match(r"^附录\s*([A-Z])", text)
        if match:
            letters.append(match.group(1))
    if len(letters) == len(texts) and letters != [
        chr(ord("A") + offset) for offset in range(len(letters))
    ]:
        problems.append(f"附录编号应按 A、B、C… 顺序：实际 {letters}")
    if not problems:
        return []
    return [
        make_issue(
            rule,
            locator=loc_para(model.by_index[indices[0]], "structure/appendix"),
            message="多项附录的编号不符合规范（附录A、附录B…）",
            actual=problems,
            expected={"multiple_appendix_numbering": True, "pattern": pattern_text},
            hint="Word：多项附录标题按“附录A/附录B…”顺序编号（编号属文字，需作者确认）",
        )
    ]


# ---------------------------------------------------------------------------
# TRANS-001 外文资料翻译格式
# ---------------------------------------------------------------------------
def _check_translation(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "TRANS-001")
    if rule is None:
        return []
    params = expected_params(ruleset, "TRANS-001", TRANS_FALLBACK)
    title_text = norm(_text(params, "title_text", "外文资料译文"))
    title_index: Optional[int] = None
    for para in model.paragraphs:
        if roles.get(para.index) in ("toc_entry", "reference_entry"):
            continue
        text_n = norm(para.text)
        if text_n == title_text or (title_text and text_n.startswith(title_text)):
            title_index = para.index
            break
    if title_index is None:
        return []
    title_size = _num(params, "title_size_pt", 16.0)
    title_font = _text(params, "title_east_asia", "黑体")
    title_align = _text(params, "title_alignment", "center")
    original_size = _num(params, "original_text_size_pt", 12.0)
    original_ascii = _text(params, "original_text_ascii", "Times New Roman")
    translation_size = _num(params, "translation_text_size_pt", 12.0)
    translation_font = _text(params, "translation_text_east_asia", "宋体")
    expected = {
        "title_text": _text(params, "title_text", "外文资料译文"),
        "title_size_pt": title_size,
        "title_east_asia": title_font,
        "title_alignment": title_align,
        "original_text_size_pt": original_size,
        "original_text_ascii": original_ascii,
        "translation_text_size_pt": translation_size,
        "translation_text_east_asia": translation_font,
        "new_page_for_translation": True,
    }
    issues: list[Issue] = []
    title = model.by_index[title_index]
    problems = para_format_problems(
        title, size_pt=title_size, east_asia=title_font, alignment=title_align
    )
    if problems:
        issues.append(
            make_issue(
                rule,
                locator=loc_para(title),
                message="“外文资料译文”标题格式不符合规范（三号黑体居中）",
                actual="；".join(problems),
                expected=expected,
            )
        )
    original = _next_nonempty(model, title_index)
    if original is None:
        issues.append(
            make_issue(
                rule,
                locator=loc_para(title),
                message="“外文资料译文”标题后缺少外文原文",
                actual="缺失",
                expected=expected,
            )
        )
        return issues
    original_problems: list[str] = []
    for run in original.runs:
        text = run.text or ""
        if not text.strip():
            continue
        if not close(run.size_pt, original_size, 0.6):
            original_problems.append(f"字号={run.size_pt if run.size_pt is not None else '未设置'}pt")
        if has_cjk(text):
            continue
        if run.ascii != original_ascii:
            original_problems.append(f"西文字体={run.ascii or '未设置'}")
    if original_problems:
        issues.append(
            make_issue(
                rule,
                locator=loc_para(original),
                message="外文原文格式不符合规范（小四 Times New Roman）",
                actual="；".join(dict.fromkeys(original_problems)),
                expected=expected,
            )
        )
    # 中文翻译：标题后第一个含汉字的段落；要求它（或中间段落）另起新页
    translation: Optional[ParaInfo] = None
    new_page = False
    cursor = original.index + 1
    while cursor < len(model.paragraphs):
        para = model.by_index.get(cursor)
        if para is None:
            break
        if para.page_break_before or para.has_page_break:
            new_page = True
        if para.is_empty:
            cursor += 1
            continue
        if CJK_RE.search(para.text or ""):
            translation = para
            break
        cursor += 1
    if translation is None:
        issues.append(
            make_issue(
                rule,
                locator=loc_para(original),
                message="未检测到外文资料的中文翻译段落",
                actual="缺失",
                expected=expected,
            )
        )
        return issues
    if not new_page:
        issues.append(
            make_issue(
                rule,
                locator=loc_para(translation),
                message="中文翻译未另起新页（规范：另起新页打印中文翻译）",
                actual="翻译段落前无分页",
                expected=expected,
                hint="Word：在中文翻译标题前插入分页符（Ctrl+Enter）",
            )
        )
    translation_problems = para_format_problems(
        translation, size_pt=translation_size, east_asia=translation_font
    )
    if translation_problems:
        issues.append(
            make_issue(
                rule,
                locator=loc_para(translation),
                message="中文翻译格式不符合规范（小四宋体）",
                actual="；".join(translation_problems),
                expected=expected,
            )
        )
    return issues
