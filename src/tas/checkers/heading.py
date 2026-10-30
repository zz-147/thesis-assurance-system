"""标题检查：一/二/三级标题格式、编号连续性、前言/结论/致谢。"""

from __future__ import annotations

import re

from ..config import RuleSet
from ..docx_io import DocModel
from ..report import Issue
from ..roles import Roles, norm
from ..rule_params import alignment as alignment_param
from ..rule_params import expected_params, font_size, line_count
from .helpers import label_gap_spaces, loc_para, make_issue, para_format_problems

# 学校规范默认值（rules.json 缺失字段时兜底；rules.json 有 expected 时以它为准）
LEVEL_FALLBACK = {
    1: {"size_pt": 16, "east_asia": "黑体", "alignment": "center",
        "line_spacing": 1.5, "before_lines": 100, "after_lines": 100},
    2: {"size_pt": 14, "east_asia": "黑体", "alignment": "left",
        "line_spacing": 1.5, "before_lines": 50, "after_lines": 50},
    3: {"size_pt": 12, "east_asia": "黑体", "alignment": "left",
        "line_spacing": 1.5, "before_lines": 50, "after_lines": 0},
}


def _level_spec(ruleset: RuleSet, level: int) -> dict:
    fallback = LEVEL_FALLBACK[level]
    params = expected_params(ruleset, f"HEAD-00{level}", fallback)
    return {
        "size_pt": font_size(params),
        "east_asia": str(params.get("east_asia", "黑体")),
        "alignment": alignment_param(params),
        "line_spacing": float(params.get("line_spacing", 1.5)),
        "before": line_count(params.get("before_lines"), fallback["before_lines"] / 100.0),
        "after": line_count(params.get("after_lines"), fallback["after_lines"] / 100.0),
    }

RE_L1 = re.compile(r"^第(\d+)章")
RE_L2 = re.compile(r"^(\d+)\.(\d+)(?!\.\d)")
RE_L3 = re.compile(r"^(\d+)\.(\d+)\.(\d+)")
RE_L4 = re.compile(r"^(\d+)\.(\d+)\.(\d+)\.(\d+)")


def check_heading(model: DocModel, roles: Roles, ruleset: RuleSet, profile: str = "sci_eng") -> list[Issue]:
    issues: list[Issue] = []
    if profile == "sci_eng":
        issues.extend(_level_formats(model, roles, ruleset))
    issues.extend(_numbering(model, roles, ruleset, profile))
    issues.extend(_special_titles(model, roles, ruleset, profile))
    return issues


def _level_formats(model: DocModel, roles: Roles, ruleset: RuleSet) -> list[Issue]:
    issues = []
    for level in (1, 2, 3):
        rule = ruleset.get(f"HEAD-00{level}")
        spec = _level_spec(ruleset, level)
        expected = {
            "size_pt": spec["size_pt"],
            "east_asia": spec["east_asia"],
            "alignment": spec["alignment"],
            "line_spacing": spec["line_spacing"],
            "before_lines": spec["before"],
            "after_lines": spec["after"],
        }
        for index in roles.indices(f"heading_{level}"):
            para = model.by_index[index]
            if para.is_empty:
                continue
            problems = para_format_problems(
                para,
                size_pt=spec["size_pt"],
                east_asia=spec["east_asia"],
                alignment=spec["alignment"],
                line_spacing=spec["line_spacing"],
                before_lines=spec["before"],
                after_lines=spec["after"],
            )
            if problems:
                issues.append(
                    make_issue(
                        rule,
                        locator=loc_para(para),
                        message="标题格式不符合规范",
                        actual="；".join(problems),
                        expected=expected,
                    )
                )
    return issues


def _numbering(model: DocModel, roles: Roles, ruleset: RuleSet, profile: str) -> list[Issue]:
    if profile != "sci_eng":
        return []
    rule = ruleset.get("HEAD-006")
    issues: list[Issue] = []
    prev_level: int | None = None
    chapter = section = subsection = sub_subsection = 0
    seen: set[tuple] = set()

    for index in roles.indices("heading_1", "heading_2", "heading_3", "heading_4"):
        para = model.by_index[index]
        if para.is_empty:
            continue
        level = int(roles.get(index).split("_")[1])
        text_n = norm(para.text)

        if prev_level is not None and level > prev_level + 1:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="标题层级跳跃",
                    actual=f"从 {prev_level} 级直接到 {level} 级",
                    expected="标题逐级出现，不跳级",
                )
            )

        if level == 1:
            match = RE_L1.match(text_n)
            if match:
                number = int(match.group(1))
                if number != chapter + 1:
                    issues.append(
                        make_issue(
                            rule,
                            locator=loc_para(para),
                            message="章号不连续",
                            actual=f"第{number}章",
                            expected=f"第{chapter + 1}章",
                        )
                    )
                chapter = number
                section = subsection = sub_subsection = 0
        elif level == 2:
            match = RE_L2.match(text_n)
            if match:
                chap, sec = int(match.group(1)), int(match.group(2))
                key = (level, chap, sec)
                if key in seen:
                    issues.append(make_issue(rule, locator=loc_para(para), message="节号重复", actual=f"{chap}.{sec}", expected="编号唯一"))
                if chap != chapter:
                    issues.append(
                        make_issue(
                            rule,
                            locator=loc_para(para),
                            message="节号中的章号与当前章不一致",
                            actual=f"{chap}.{sec}",
                            expected=f"当前应为第{chapter}章",
                        )
                    )
                elif sec != section + 1:
                    issues.append(
                        make_issue(
                            rule,
                            locator=loc_para(para),
                            message="节号不连续",
                            actual=f"{chap}.{sec}",
                            expected=f"{chapter}.{section + 1}",
                        )
                    )
                seen.add(key)
                section, subsection, sub_subsection = sec, 0, 0
        elif level == 3:
            match = RE_L3.match(text_n)
            if match:
                chap, sec, sub = (int(match.group(i)) for i in (1, 2, 3))
                key = (level, chap, sec, sub)
                if key in seen:
                    issues.append(make_issue(rule, locator=loc_para(para), message="小节号重复", actual=f"{chap}.{sec}.{sub}", expected="编号唯一"))
                if chap != chapter or sec != section:
                    issues.append(
                        make_issue(
                            rule,
                            locator=loc_para(para),
                            message="小节号的上级编号与当前章节不一致",
                            actual=f"{chap}.{sec}.{sub}",
                            expected=f"当前应为第{chapter}章 {chapter}.{section} 节",
                        )
                    )
                elif sub != subsection + 1:
                    issues.append(
                        make_issue(
                            rule,
                            locator=loc_para(para),
                            message="小节号不连续",
                            actual=f"{chap}.{sec}.{sub}",
                            expected=f"{chapter}.{section}.{subsection + 1}",
                        )
                    )
                seen.add(key)
                subsection, sub_subsection = sub, 0
        elif level == 4:
            match = RE_L4.match(text_n)
            if match:
                chap, sec, sub, subsub = (int(match.group(i)) for i in (1, 2, 3, 4))
                if subsub != sub_subsection + 1:
                    issues.append(
                        make_issue(
                            rule,
                            locator=loc_para(para),
                            message="四级标题编号不连续",
                            actual=f"{chap}.{sec}.{sub}.{subsub}",
                            expected=f"{chap}.{sec}.{sub}.{sub_subsection + 1}",
                        )
                    )
                sub_subsection = subsub

        prev_level = level
    return issues


def _special_titles(model: DocModel, roles: Roles, ruleset: RuleSet, profile: str) -> list[Issue]:
    rule = ruleset.get("HEAD-007")
    # 注意：rules.json 的 HEAD-007.expected 没有 after_lines；按规范：
    # 理工农 前言/结论/致谢 段前段后各 1 行；文经管 段前 1 行、段后无间距。
    fallback = {
        "size_pt": 16,
        "east_asia": "黑体",
        "alignment": {"sci_eng": "center", "humanities": "left"},
        "line_spacing": 1.5,
        "before_lines": 100,
    }
    params = expected_params(ruleset, "HEAD-007", fallback)
    align_cfg = params.get("alignment")
    if isinstance(align_cfg, dict):
        alignment = str(align_cfg.get(profile, "center"))
    else:
        alignment = alignment_param(params, fallback="center")
    size = font_size(params)
    east_asia = str(params.get("east_asia", "黑体"))
    line_spacing = float(params.get("line_spacing", 1.5))
    before = line_count(params.get("before_lines"), 1.0)
    after_fallback = 1.0 if profile == "sci_eng" else 0.0
    after = line_count(params.get("after_lines"), after_fallback)
    expected = {
        "size_pt": size,
        "east_asia": east_asia,
        "alignment": alignment,
        "line_spacing": line_spacing,
        "before_lines": before,
        "after_lines": after,
    }
    issues = []
    for index in roles.indices("special_title", "thanks_title"):
        para = model.by_index[index]
        if para.is_empty:
            continue
        problems = para_format_problems(
            para,
            size_pt=size,
            east_asia=east_asia,
            alignment=alignment,
            line_spacing=line_spacing,
            before_lines=before,
            after_lines=after,
        )
        text_n = norm(para.text)
        if text_n in ("前言", "结论", "致谢"):
            gap = label_gap_spaces(para.text, text_n)
            if gap is not None and gap != 2:
                problems.append(f"两字之间空格={gap}（应为 2 格）")
        if problems:
            issues.append(
                make_issue(
                    rule,
                    locator=loc_para(para),
                    message="前言/结论/致谢标题格式不符合规范",
                    actual="；".join(problems),
                    expected=expected,
                )
            )
    return issues
