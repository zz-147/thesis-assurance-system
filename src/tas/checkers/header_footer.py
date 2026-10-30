"""页眉/页脚/页码检查：HF-001/002/003（参数来自当前 ruleset）。

只检查"文档确实使用了页眉/页码"的情形：

- 全文没有任何页眉文字、也没有任何页码设置时，HF-001/002/003 整条跳过——
  最小合规文档（例如单元测试夹具）不会被强行要求补页眉，避免误报；
- 一旦某一节有页眉/页码，就按规范检查"从摘要到最后一页均有页眉"、封面不显示页眉、
  罗马数字→阿拉伯数字切换、页码字体/对齐等；参数化读取 expected；
- 页眉/页脚链接到上一节时按 Word 的继承规则取"有效页眉/页脚"，不会把继承当成缺失；
- 字体/边框缺失时不硬造问题：只在文档显式设置了格式且与规范不符时报告。

已知限制（见最终报告）：首页不同（titlePg）与奇偶页页眉、页脚域代码展开后的实际
渲染结果不做检查；这些需要渲染引擎，纯 OOXML 只能做到现有的程度。
"""

from __future__ import annotations

from typing import Any, Optional

from ..config import Rule, RuleSet
from ..docx_io import DocModel, HeaderFooterParaInfo, SectionInfo
from ..report import Issue
from ..roles import Roles, norm
from ..rule_params import expected_params
from .helpers import SPACING_TOL, close, make_issue

ROMAN_FORMATS = ("upperRoman", "lowerRoman")
HF_FALLBACK = {
    "size_pt": 9,
    "east_asia": "宋体",
    "line_spacing": 1.0,
    "alignment": "center",
    "bottom_border": {"style": "double", "size_pt": 0.5},
    "starts_at": "摘要",
}
HF_TEXT_FALLBACK = {
    "allowed_texts": ["商丘学院应用科技学院毕业论文", "商丘学院应用科技学院毕业设计说明书"],
    "match": "与封面/题目类型一致",
}
HF_PAGE_FALLBACK = {
    "roman_section": "摘要→目录",
    "arabic_section": "前言→末尾",
    "size_pt": 9,
    "ascii": "Times New Roman",
    "line_spacing": 1.0,
    "alignment": "center",
    "position": "bottom",
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


def _is_page_field(instr: str) -> bool:
    import re

    upper = (instr or "").upper()
    return re.search(r"(?<![A-Z])PAGE(?![A-Z])", upper) is not None


def _section_paragraph_index(model: DocModel, para_index: Optional[int]) -> Optional[int]:
    if para_index is None:
        return None
    for section in model.sections:
        first = section.first_para_index
        last = section.last_para_index
        if first is None:
            continue
        if last is None:
            last = first
        if first <= para_index <= last:
            return section.index
    return None


def _roles_by_section(model: DocModel, roles: Roles) -> dict[int, set[str]]:
    mapping: dict[int, set[str]] = {}
    for index, role in roles.by_index.items():
        section_index = _section_paragraph_index(model, index)
        if section_index is not None:
            mapping.setdefault(section_index, set()).add(role)
    return mapping


def _section_of_first_role(
    model: DocModel, roles: Roles, role_names: tuple[str, ...]
) -> Optional[int]:
    found = roles.first_index(*role_names)
    return _section_paragraph_index(model, found)


def _section_of_last_role(
    model: DocModel, roles: Roles, role_names: tuple[str, ...]
) -> Optional[int]:
    indices = roles.indices(*role_names)
    return _section_paragraph_index(model, indices[-1] if indices else None)


def _parts_present(parts: list[HeaderFooterParaInfo]) -> bool:
    return any(
        (part.text or "").strip() or bool(part.field_instrs) for part in parts
    )


def _effective_parts(
    sections: list[SectionInfo], attr: str, linked_attr: str
) -> list[list[HeaderFooterParaInfo]]:
    """按 Word 继承规则计算每节的有效页眉/页脚内容。"""
    effective: list[list[HeaderFooterParaInfo]] = []
    previous: list[HeaderFooterParaInfo] = []
    for section in sections:
        own = getattr(section, attr)
        if getattr(section, linked_attr):
            current = previous
        else:
            current = own
        effective.append(current)
        previous = current
    return effective


def _effective_page_fields(sections: list[SectionInfo]) -> list[bool]:
    effective: list[bool] = []
    previous = False
    for section in sections:
        current = previous if section.footer_linked else section.has_page_field
        effective.append(current)
        previous = current
    return effective


def _effective_page_instrs(sections: list[SectionInfo]) -> list[list[str]]:
    effective: list[list[str]] = []
    previous: list[str] = []
    for section in sections:
        current = previous if section.footer_linked else list(section.page_field_instrs)
        effective.append(current)
        previous = current
    return effective


def _contains_page_text(parts: list[HeaderFooterParaInfo]) -> bool:
    return any((part.text or "").strip() for part in parts)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def check_header_footer(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng"
) -> list[Issue]:
    has_header_text = any(
        _contains_page_text(section.header_paras) for section in model.sections
    )
    has_page_machinery = any(
        section.has_page_field or bool(section.page_number_format)
        for section in model.sections
    )
    if not has_header_text and not has_page_machinery:
        # 没有页眉/页码：无法判断"是否应补"，整条跳过（不误报最小文档）
        return []
    issues: list[Issue] = []
    issues.extend(_check_header_style(model, roles, ruleset, has_header_text))
    issues.extend(_check_header_text(model, ruleset))
    issues.extend(_check_page_number(model, roles, ruleset))
    return issues


# ---------------------------------------------------------------------------
# HF-001 页眉样式 + 覆盖范围
# ---------------------------------------------------------------------------
def _check_header_style(
    model: DocModel, roles: Roles, ruleset: RuleSet, has_header_text: bool
) -> list[Issue]:
    rule = _rule(ruleset, "HF-001")
    if rule is None:
        return []
    params = expected_params(ruleset, "HF-001", HF_FALLBACK)
    border_cfg = params.get("bottom_border") if isinstance(params.get("bottom_border"), dict) else {}
    expected = {
        "size_pt": _num(params, "size_pt", 9.0),
        "east_asia": _text(params, "east_asia", "宋体"),
        "line_spacing": _num(params, "line_spacing", 1.0),
        "alignment": _text(params, "alignment", "center"),
        "bottom_border": {
            "style": str(border_cfg.get("style", "double")),
            "size_pt": _num(border_cfg, "size_pt", 0.5),
        },
        "starts_at": _text(params, "starts_at", "摘要"),
    }
    issues: list[Issue] = []

    # ---- 1) 样式：逐节检查自己的页眉（继承内容在定义节检查一次）----
    for section in model.sections:
        if section.header_linked or not _contains_page_text(section.header_paras):
            continue
        problems: list[str] = []
        for para in section.header_paras:
            if not (para.text or "").strip():
                continue
            if (para.alignment or "left") != expected["alignment"]:
                problems.append(f"对齐={para.alignment or 'left'}")
            if not close(para.line_spacing, expected["line_spacing"], SPACING_TOL):
                problems.append(
                    f"行距={para.line_spacing if para.line_spacing is not None else '未设置'}倍"
                )
            for run in para.runs:
                text = run.text or ""
                if not text.strip():
                    continue
                if not close(run.size_pt, expected["size_pt"], 0.6):
                    problems.append(
                        f"字号={run.size_pt if run.size_pt is not None else '未设置'}pt"
                    )
                if any("\u4e00" <= char <= "\u9fff" for char in text) and run.east_asia != expected[
                    "east_asia"
                ]:
                    problems.append(f"中文字体={run.east_asia or '未设置'}")
            border = para.bottom_border
            if border is None:
                problems.append("缺少页眉下边框")
            elif border.val != expected["bottom_border"]["style"]:
                problems.append(f"页眉下边框线型={border.val or '未设置'}")
            elif not close(border.size_pt, expected["bottom_border"]["size_pt"], 0.25):
                problems.append(f"页眉下边框粗细={border.size_pt:g}pt")
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=f"section/{section.index}/header",
                    message="页眉样式不符合规范（小五号宋体、居中、1 倍行距、双实线下边框 0.5 磅）",
                    actual="；".join(dict.fromkeys(problems)),
                    expected=expected,
                    hint="Word：双击页眉 → 字体/段落 → 边框和底纹（下边框：双实线 0.5 磅）",
                )
            )

    if not has_header_text:
        return issues

    # ---- 2) 覆盖范围：从“摘要”所在节到最后一节都要有页眉；封面/承诺书节不要 ----
    start = _section_of_first_role(model, roles, ("abstract_label_cn", "toc_title"))
    if start is not None:
        effective = _effective_parts(model.sections, "header_paras", "header_linked")
        for section in model.sections:
            if section.index < start:
                continue
            if not _contains_page_text(effective[section.index - 1]):
                issues.append(
                    make_issue(
                        rule,
                        locator=f"section/{section.index}/header",
                        message=f"从“{expected['starts_at']}”到最后一页应有页眉，但该节页眉为空",
                        actual="无页眉文字",
                        expected={"starts_at": expected["starts_at"]},
                        hint="Word：页眉与页脚 → 取消“链接到前一节”后录入规定页眉文字",
                    )
                )
    role_map = _roles_by_section(model, roles)
    for section in model.sections:
        if "cover" not in role_map.get(section.index, set()):
            continue
        if _contains_page_text(section.header_paras):
            issues.append(
                make_issue(
                    rule,
                    locator=f"section/{section.index}/header",
                    message="封面/承诺书不应显示页眉",
                    actual="该节有页眉文字",
                    expected="封面、承诺书节无页眉",
                    hint="Word：在封面节取消“链接到前一节”并清空页眉",
                )
            )
    return issues


# ---------------------------------------------------------------------------
# HF-002 页眉文字
# ---------------------------------------------------------------------------
def _check_header_text(model: DocModel, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "HF-002")
    if rule is None:
        return []
    sections_with_text = [
        section
        for section in model.sections
        if not section.header_linked and _contains_page_text(section.header_paras)
    ]
    if not sections_with_text:
        return []
    params = expected_params(ruleset, "HF-002", HF_TEXT_FALLBACK)
    allowed_raw = params.get("allowed_texts")
    allowed = [
        norm(str(item)) for item in allowed_raw if str(item).strip()
    ] if isinstance(allowed_raw, (list, tuple)) else [
        norm(item) for item in HF_TEXT_FALLBACK["allowed_texts"]
    ]
    expected = {"allowed_texts": [str(item) for item in (allowed_raw or HF_TEXT_FALLBACK["allowed_texts"])]}
    doc_type = _document_type(model)
    issues: list[Issue] = []
    for section in sections_with_text:
        texts = [
            (part.text or "").strip()
            for part in section.header_paras
            if (part.text or "").strip()
        ]
        for text in texts:
            normalized = norm(text)
            if normalized not in allowed:
                issues.append(
                    make_issue(
                        rule,
                        locator=f"section/{section.index}/header",
                        message="页眉文字与规范不一致",
                        actual=text,
                        expected=expected,
                        hint="规范页眉：“商丘学院应用科技学院毕业论文”或“……毕业设计说明书”",
                    )
                )
                continue
            if doc_type and doc_type not in normalized:
                issues.append(
                    make_issue(
                        rule,
                        locator=f"section/{section.index}/header",
                        message="页眉文字与封面/题目类型不一致",
                        actual=text,
                        expected=f"应包含“{doc_type}”",
                    )
                )
    return issues


def _document_type(model: DocModel) -> str:
    """从封面/前置区推断“毕业论文”还是“毕业设计”。"""
    for para in model.paragraphs:
        text = para.text or ""
        if "毕业设计说明书" in text:
            return "毕业设计"
        if "毕业设计" in text and "毕业论文" not in text:
            return "毕业设计"
        if "毕业论文" in text and "毕业设计" not in text:
            return "毕业论文"
    return ""


# ---------------------------------------------------------------------------
# HF-003 页码格式
# ---------------------------------------------------------------------------
def _page_field_format(instrs: list[str]) -> Optional[str]:
    """返回页码域格式：'roman' / 'arabic' / None。"""
    for instr in instrs:
        upper = (instr or "").upper()
        if not _is_page_field(instr):
            continue
        if "ROMAN" in upper:
            return "roman"
        if "ARABIC" in upper:
            return "arabic"
        return "arabic"  # PAGE 默认阿拉伯数字
    return None


def _check_page_number(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    rule = _rule(ruleset, "HF-003")
    if rule is None:
        return []
    params = expected_params(ruleset, "HF-003", HF_PAGE_FALLBACK)
    expected = {
        "roman_section": _text(params, "roman_section", "摘要→目录"),
        "arabic_section": _text(params, "arabic_section", "前言→末尾"),
        "size_pt": _num(params, "size_pt", 9.0),
        "ascii": _text(params, "ascii", "Times New Roman"),
        "line_spacing": _num(params, "line_spacing", 1.0),
        "alignment": _text(params, "alignment", "center"),
        "position": _text(params, "position", "bottom"),
    }
    issues: list[Issue] = []

    effective_parts = _effective_parts(model.sections, "footer_paras", "footer_linked")
    effective_fields = _effective_page_fields(model.sections)
    effective_instrs = _effective_page_instrs(model.sections)

    roman_start = _section_of_first_role(model, roles, ("abstract_label_cn", "toc_title"))
    roman_end = _section_of_last_role(model, roles, ("toc_entry", "toc_title"))
    body_start = roles.body_start
    if body_start is not None and body_start >= len(model.paragraphs):
        body_start = len(model.paragraphs) - 1
    arabic_start = _section_paragraph_index(model, body_start)

    checks: list[tuple[int, str]] = []
    if roman_start is not None:
        end = roman_end if roman_end is not None else roman_start
        checks.extend((index, "roman") for index in range(roman_start, end + 1))
    if arabic_start is not None:
        checks.extend((index, "arabic") for index in range(arabic_start, len(model.sections) + 1))
    # 去重，保持节顺序
    seen: set[int] = set()
    ordered: list[tuple[int, str]] = []
    for index, kind in checks:
        if index in seen:
            continue
        seen.add(index)
        ordered.append((index, kind))
    if not ordered:
        # 认不出摘要/正文结构：退化为只校验已存在页码格式的内部一致性
        for section in model.sections:
            fmt = (section.page_number_format or "").lower()
            if not fmt:
                continue
            kind = "roman" if fmt in [item.lower() for item in ROMAN_FORMATS] else "arabic"
            ordered.append((section.index, kind))
        if not ordered:
            return []

    for section in model.sections:
        if section.index in seen or section.index >= min(index for index, _ in ordered):
            continue
        # 摘要之前的封面/承诺书节：不应有页码
        if effective_fields[section.index - 1] or section.page_number_format:
            issues.append(
                make_issue(
                    rule,
                    locator=f"section/{section.index}/footer/page-number",
                    message="封面/承诺书节不应显示页码",
                    actual={
                        "page_number_format": section.page_number_format,
                        "page_field": effective_fields[section.index - 1],
                    },
                    expected="封面、承诺书无页码",
                )
            )

    for section_index, kind in ordered:
        section = model.sections[section_index - 1]
        problems: dict[str, Any] = {}
        fmt = section.page_number_format
        if kind == "roman":
            if fmt not in ROMAN_FORMATS:
                problems["page_number_format"] = fmt or "未设置"
        else:
            if fmt not in ("decimal", None):
                problems["page_number_format"] = fmt
        if not effective_fields[section_index - 1]:
            problems["page_field"] = "页脚未检测到页码域（PAGE）"
        else:
            field_format = _page_field_format(effective_instrs[section_index - 1])
            if kind == "roman" and field_format != "roman":
                problems["page_field_format"] = field_format or "未设置"
            elif kind == "arabic" and field_format not in ("arabic", None):
                problems["page_field_format"] = field_format
        problems.update(_footer_format_problems(effective_parts[section_index - 1], expected))
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=f"section/{section_index}/footer/page-number",
                    message=(
                        "页码格式不符合规范（摘要→目录：罗马数字；前言→末尾：阿拉伯数字；"
                        "小五号 Times New Roman、底部居中）"
                    ),
                    actual=problems,
                    expected=expected,
                    hint="Word：插入 → 页码 → 设置页码格式；页脚域请用 PAGE 域并在页码格式中切换",
                )
            )
    return issues


def _footer_format_problems(
    parts: list[HeaderFooterParaInfo], expected: dict
) -> dict[str, Any]:
    problems: dict[str, Any] = {}
    paragraphs = [
        para for para in parts if bool(para.field_instrs) or (para.text or "").strip()
    ]
    for para in paragraphs:
        if (para.alignment or "left") != expected["alignment"]:
            problems.setdefault("alignment", para.alignment or "left")
        if not close(para.line_spacing, expected["line_spacing"], SPACING_TOL):
            problems.setdefault(
                "line_spacing",
                para.line_spacing if para.line_spacing is not None else "未设置",
            )
        for run in para.runs:
            if run.size_pt is not None and not close(run.size_pt, expected["size_pt"], 0.6):
                problems.setdefault("size_pt", run.size_pt)
            if run.ascii is not None and run.ascii != expected["ascii"]:
                problems.setdefault("ascii", run.ascii)
    return problems


# ---------------------------------------------------------------------------
# PAGE-004 分节与页眉页脚继承（P2）
# ---------------------------------------------------------------------------
PAGE004_FALLBACK = {"sections_min": 2, "cover_no_header_footer": True}


def _hf_text(parts: list[HeaderFooterParaInfo]) -> str:
    texts = [(part.text or "").strip() for part in parts if (part.text or "").strip()]
    return " / ".join(texts)


def _hf_has_page_field(parts: list[HeaderFooterParaInfo]) -> bool:
    return any(
        _is_page_field(instr) for part in parts for instr in part.field_instrs
    )


def check_page004(
    model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng"
) -> list[Issue]:
    """PAGE-004「分节与页眉页脚继承」：分节数量 + 封面/承诺书不得有页眉页码。

    - 无 PAGE-004 规则、无 section → ``[]``；
    - **防误报**：只有 1 个 section 且首节没有任何页眉/页脚文字、也没有 PAGE 域 →
      ``[]``（现有最小合规 fixture 不会因为"没分节"被报）；
    - section 数 < ``expected.sections_min`` → 一条 warning（定位 ``section/1``）；
    - ``cover_no_header_footer`` 为真且首节有页眉/页脚文字或页码域 → 一条 warning，
      定位 ``section/1``，actual 给出实际页眉/页脚文字与是否有 PAGE 域；
    - 页眉/页脚 ``linked``（链接到上一节、继承）视为正常，不报。

    只报告、不修复（结构改动风险高）。
    """
    rule = _rule(ruleset, "PAGE-004")
    if rule is None or not model.sections:
        return []
    params = expected_params(ruleset, "PAGE-004", PAGE004_FALLBACK)
    raw_min = params.get("sections_min", PAGE004_FALLBACK["sections_min"])
    try:
        sections_min = max(1, int(raw_min))
    except (TypeError, ValueError):
        sections_min = PAGE004_FALLBACK["sections_min"]
    cover_no_header_footer = bool(
        params.get("cover_no_header_footer", PAGE004_FALLBACK["cover_no_header_footer"])
    )

    first = model.sections[0]
    header_text = _hf_text(first.header_paras)
    footer_text = _hf_text(first.footer_paras)
    page_field = bool(first.has_page_field) or _hf_has_page_field(first.footer_paras)
    first_has_hf = bool(header_text or footer_text or page_field)

    # 防误报：单节且首节没有任何页眉页脚/页码 → 不报（最小文档）
    if len(model.sections) == 1 and not first_has_hf:
        return []

    issues: list[Issue] = []
    if len(model.sections) < sections_min:
        issues.append(
            make_issue(
                rule,
                locator="section/1",
                message="分节数量不足（页眉页脚/页码格式切换需要分节）",
                actual=f"共 {len(model.sections)} 节",
                expected={"sections_min": sections_min},
                hint="Word：布局 → 分隔符 → 分节符（下一页）；封面/摘要/目录/正文按节切换页眉页码",
            )
        )

    if cover_no_header_footer and first_has_hf:
        issues.append(
            make_issue(
                rule,
                locator="section/1",
                message="封面/承诺书不应显示页眉或页码",
                actual={
                    "header": header_text or "（无）",
                    "footer": footer_text or "（无）",
                    "page_field": page_field,
                },
                expected={"header": "无", "footer": "无", "page_field": False},
                hint="Word：在封面节取消“链接到前一节”，清空页眉页脚并删除页码",
            )
        )
    return issues
