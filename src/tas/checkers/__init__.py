"""规则检查器集合与执行入口。"""

from __future__ import annotations

from ..config import DEFAULT_PROFILE, RuleSet
from ..docx_io import DocModel
from ..report import Issue
from ..roles import Roles
from .body import check_body
from .float import check_float
from .front import check_front
from .header_footer import check_header_footer, check_page004
from .heading import check_heading
from .notes import check_notes
from .page import check_page
from .references_check import check_references
from .structure import check_structure

ALL_CHECKS = (
    check_page,
    check_body,
    check_heading,
    check_front,
    check_float,
    check_header_footer,
    check_page004,
    check_notes,
    check_references,
    check_structure,
)

# 已实现检查器的规则 id（rules.json 里其余规则是后续迭代目标）
IMPLEMENTED_RULES = (
    # P0（原有）
    "PAGE-001", "PAGE-002", "PAGE-003",
    "BODY-001", "BODY-002", "BODY-003", "BODY-004",
    "HEAD-001", "HEAD-002", "HEAD-003", "HEAD-006", "HEAD-007",
    "TITLE-001", "TITLE-002",
    "ABS-001", "ABS-002", "ABS-003", "ABS-004",
    "KW-001", "KW-002",
    "TOC-001", "REF-001",
    # P1（float / header_footer / references_check / structure）
    "FIG-001", "FIG-002", "FIG-003",
    "TAB-001", "TAB-002", "TAB-003",
    "EQ-001",
    "HF-001", "HF-002", "HF-003",
    "REF-002", "REF-003",
    "TOC-002",
    "STRUCT-001", "MISC-001",
    "HEAD-004", "HEAD-005",
    "APP-001", "APP-002", "TRANS-001",
    # P2 补齐
    "PAGE-004", "NOTE-001",
)


def run_checks(
    model: DocModel,
    roles: Roles,
    ruleset: RuleSet,
    *,
    profile: str = DEFAULT_PROFILE,
    phase: str = "all",
) -> list[Issue]:
    """执行全部已实现检查器；phase 可为 all / P0 / P1（按规则声明的阶段过滤）。"""
    issues: list[Issue] = []
    for check in ALL_CHECKS:
        issues.extend(check(model, roles, ruleset, profile=profile))
    if phase != "all":
        issues = [issue for issue in issues if ruleset.get(issue.rule_id).phase == phase]
    return issues
