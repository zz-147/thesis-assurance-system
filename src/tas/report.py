"""检查结果模型与报告渲染。"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

from . import __version__
from .config import PROFILES

SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}
SEVERITY_LABEL = {"error": "错误", "warning": "警告", "info": "提示"}


@dataclass
class Issue:
    """一个不符合项。"""

    rule_id: str
    title: str
    severity: str
    message: str
    locator: str
    actual: Optional[Any] = None
    expected: Optional[Any] = None
    hint: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


def _fmt(value: Any) -> str:
    if value is None:
        return "（无）"
    if isinstance(value, dict):
        return "，".join(f"{k}={_fmt(v)}" for k, v in value.items())
    if isinstance(value, (list, tuple)):
        return "[" + "，".join(_fmt(v) for v in value) + "]"
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def _loc_key(locator: str) -> tuple:
    nums = tuple(int(x) for x in re.findall(r"\d+", locator)[:4])
    return nums + (locator,)


@dataclass
class Report:
    """一次检查的完整结果（JSON 是单一事实来源）。"""

    path: Path
    profile: str
    issues: list[Issue] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)
    rules_run: int = 0
    rules_total: int = 0
    generated_at: str = ""

    def __post_init__(self) -> None:
        if not self.generated_at:
            self.generated_at = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
        self.issues.sort(
            key=lambda i: (SEVERITY_ORDER.get(i.severity, 9), i.rule_id, _loc_key(i.locator))
        )

    # -- 统计 -------------------------------------------------------------
    def counts(self) -> dict[str, int]:
        counter = Counter(issue.severity for issue in self.issues)
        return {key: counter.get(key, 0) for key in ("error", "warning", "info")}

    def has_errors(self) -> bool:
        return any(issue.severity == "error" for issue in self.issues)

    def by_rule(self) -> dict[str, int]:
        counter = Counter(issue.rule_id for issue in self.issues)
        return dict(counter.most_common())

    # -- 序列化 -----------------------------------------------------------
    def to_dict(self) -> dict:
        return {
            "tool": "tas",
            "version": __version__,
            "generated_at": self.generated_at,
            "file": str(self.path),
            "profile": self.profile,
            "profile_label": PROFILES.get(self.profile, self.profile),
            "stats": self.stats,
            "rules": {"total": self.rules_total, "checked": self.rules_run},
            "counts": self.counts(),
            "issues_by_rule": self.by_rule(),
            "issues": [issue.to_dict() for issue in self.issues],
        }

    # -- 渲染 -------------------------------------------------------------
    def _header_lines(self) -> list[str]:
        counts = self.counts()
        stats = self.stats
        lines = [
            "tas 格式检查报告",
            "=" * 64,
            f"文件    ：{self.path}",
            f"规范    ：{PROFILES.get(self.profile, self.profile)}（profile={self.profile}）",
            f"规则    ：本次检查 {self.rules_run} 条（规则库共 {self.rules_total} 条）",
            f"概况    ：错误 {counts['error']} ｜ 警告 {counts['warning']} ｜ 提示 {counts['info']}",
        ]
        if stats:
            lines.append(
                "文档    ：段落 {paragraphs} ｜ 表格 {tables} ｜ 节 {sections} ｜ "
                "图片 {images} ｜ 字符 {chars}".format(
                    paragraphs=stats.get("paragraphs", 0),
                    tables=stats.get("tables", 0),
                    sections=stats.get("sections", 0),
                    images=stats.get("images", 0),
                    chars=stats.get("chars", 0),
                )
            )
        lines.append("")
        return lines

    def render_text(self, max_issues: int = 80) -> str:
        lines = self._header_lines()
        if not self.issues:
            lines.append("未发现问题（当前已实现规则范围内）。")
            lines.append("")
            lines.append("注意：仅覆盖 rules/rules.json 中已实现规则的检查范围，未覆盖项不视为通过。")
            return "\n".join(lines)

        for index, issue in enumerate(self.issues[:max_issues], 1):
            label = SEVERITY_LABEL.get(issue.severity, issue.severity)
            lines.append(f"[{label}] {issue.rule_id}  {issue.title}")
            lines.append(f"    位置：{issue.locator}")
            if issue.actual is not None:
                lines.append(f"    实际：{_fmt(issue.actual)}")
            if issue.expected is not None:
                lines.append(f"    期望：{_fmt(issue.expected)}")
            lines.append(f"    说明：{issue.message}")
            if issue.hint:
                lines.append(f"    建议：{issue.hint}")
            lines.append("")

        if len(self.issues) > max_issues:
            lines.append(f"…… 其余 {len(self.issues) - max_issues} 条已省略（用 --format json 导出全量）")

        lines.append("-" * 64)
        lines.append("提示：自动检查结果仅供参考；误报请记录，教师端会用反馈校准规则。")
        return "\n".join(lines)

    def render_markdown(self, max_issues: int = 80) -> str:
        counts = self.counts()
        lines = [
            f"# tas 格式检查报告",
            "",
            f"- 文件：`{self.path}`",
            f"- 规范：{PROFILES.get(self.profile, self.profile)}（`{self.profile}`）",
            f"- 规则：本次检查 {self.rules_run} 条 / 规则库 {self.rules_total} 条",
            f"- 概况：**错误 {counts['error']}** ｜ 警告 {counts['warning']} ｜ 提示 {counts['info']}",
            "",
        ]
        if not self.issues:
            lines.append("未发现问题（当前已实现规则范围内）。")
            return "\n".join(lines)

        lines.append("| # | 级别 | 规则 | 位置 | 说明 |")
        lines.append("|---|---|---|---|---|")
        for index, issue in enumerate(self.issues[:max_issues], 1):
            lines.append(
                f"| {index} | {SEVERITY_LABEL.get(issue.severity, issue.severity)} | "
                f"{issue.rule_id} {issue.title} | {issue.locator} | {issue.message} |"
            )
        return "\n".join(lines)
