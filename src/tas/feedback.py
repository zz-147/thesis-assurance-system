"""逐份「学生反馈」文本 + 反馈 zip。

老师收齐文档后跑一次批量检查，除了 xlsx/csv/html 汇总，还生成：
    feedback.zip
    ├─ 00_班级汇总.txt          # 班级通过率 / 共性问题 / 未通过名单
    ├─ 反馈_学号_姓名.txt       # 每份文档一段，可直接复制/转发给对应学生
    └─ ...

反馈只陈述格式检查结果（规则 + 位置 + 实际/期望），不评价内容质量；
不包含任何自动修改文件的动作。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any, Optional, Sequence

from .batch import BatchSummary, DocResult
from .report import SEVERITY_LABEL

MAX_ISSUES = 12
_SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}


def _safe_stem(path: Path) -> str:
    stem = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", path.stem).strip("._")
    return (stem or "document")[:80]


def _fmt(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, dict):
        return "，".join(f"{key}={_fmt(item)}" for key, item in value.items())
    if isinstance(value, (list, tuple)):
        return "[" + "，".join(_fmt(item) for item in value) + "]"
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def feedback_text(
    result: DocResult,
    *,
    profile_label: str = "",
    max_issues: int = MAX_ISSUES,
) -> str:
    """生成单份文档的反馈文本（纯文本，微信/邮件可直接发）。"""
    counts = result.counts()
    lines: list[str] = ["【论文格式检查反馈】", f"文件：{result.path.name}"]
    if profile_label:
        lines.append(f"检查规范：{profile_label}")
    lines.append("")

    if result.error:
        lines += [
            f"结论：检查失败（{result.error}）",
            "",
            "建议：",
            "1. 确认文件是有效的 .docx（不是 .doc、未加密、未损坏）；",
            "2. 用 Word 打开后「另存为」.docx，再检查一次；",
            "3. 若仍然失败，把本段错误信息发给老师。",
        ]
        return "\n".join(lines)

    status = "通过" if result.passed() else "未通过"
    lines += [
        f"结论：{status}",
        (
            f"问题统计：错误 {counts.get('error', 0)} ｜ "
            f"警告 {counts.get('warning', 0)} ｜ 提示 {counts.get('info', 0)}"
        ),
    ]

    if not result.issues:
        lines += ["", "未发现格式问题。"]
    else:
        ordered = sorted(
            result.issues,
            key=lambda issue: (
                _SEVERITY_ORDER.get(issue.severity, 9),
                issue.rule_id,
                issue.locator,
            ),
        )
        shown = min(max(1, int(max_issues)), len(ordered))
        lines += ["", f"主要问题（最多列 {shown} 条，共 {len(ordered)} 条）："]
        for index, issue in enumerate(ordered[:shown], start=1):
            detail = (
                f"{index}. [{SEVERITY_LABEL.get(issue.severity, issue.severity)}]"
                f" {issue.rule_id} {issue.title} @ {issue.locator}：{issue.message}"
            )
            actual, expected = _fmt(issue.actual), _fmt(issue.expected)
            if actual or expected:
                detail += f"（实际：{actual or '—'}；期望：{expected or '—'}）"
            lines.append(detail)
        if len(ordered) > shown:
            lines.append(f"…… 其余 {len(ordered) - shown} 条见老师端完整报告。")
        lines += [
            "",
            "修改建议：",
            "1. 按上面问题逐项调整字体/字号/行距/缩进/对齐等格式；",
            "2. 修改后请老师再检查一次；",
            "3. 格式调整不改变文字内容，参考文献不需要动。",
        ]

    lines += [
        "",
        "说明：本反馈只检查格式规范，不评价内容质量，也不会自动修改文字或参考文献。",
    ]
    return "\n".join(lines) + "\n"


def feedback_filename(result: DocResult) -> str:
    return f"反馈_{_safe_stem(result.path)}.txt"


def _summary_text(summary: BatchSummary, profile_label: str = "") -> str:
    counts = summary.counts
    lines = [
        "【班级格式检查汇总】",
        f"检查规范：{profile_label or summary.ruleset_id or '默认规范'}",
        f"生成时间：{summary.generated_at}",
        "",
        f"文件总数：{summary.total}",
        f"检查成功：{summary.checked} ｜ 检查失败：{summary.failed}",
        f"通过率：{summary.pass_rate:.1%}（{summary.pass_count}/{summary.checked}）",
        (
            f"问题总数：错误 {counts.get('error', 0)} ｜ "
            f"警告 {counts.get('warning', 0)} ｜ 提示 {counts.get('info', 0)}"
        ),
        "",
        "共性问题 Top：",
    ]
    if summary.top_rules:
        for index, (rule_id, hit_count) in enumerate(summary.top_rules, start=1):
            lines.append(f"{index}. {rule_id} × {hit_count}")
    else:
        lines.append("（无）")

    lines += ["", "未通过名单："]
    failed = [
        Path(item.get("path", "")).name
        for item in summary.per_doc
        if not item.get("passed")
    ]
    if failed:
        lines.extend(f"- {name}" for name in failed)
    else:
        lines.append("（无，全部通过）")

    failed_checks = [
        f"- {Path(item.get('path', '')).name}：{item.get('error') or '检查失败'}"
        for item in summary.per_doc
        if not item.get("ok", True)
    ]
    if failed_checks:
        lines += ["", "检查失败（文件可能损坏/加密/非 docx）："] + failed_checks

    lines += ["", "把每份「反馈_*.txt」发给对应学生即可。"]
    return "\n".join(lines) + "\n"


def export_feedback_zip(
    results: Sequence[DocResult],
    out_zip: str | Path,
    *,
    summary: Optional[BatchSummary] = None,
    profile_label: str = "",
    max_issues: int = MAX_ISSUES,
) -> Path:
    """把逐份反馈 + 班级汇总写成 zip；返回 zip 路径。"""
    out = Path(out_zip)
    out.parent.mkdir(parents=True, exist_ok=True)
    items = list(results)
    used: set[str] = set()

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        if summary is not None:
            archive.writestr("00_班级汇总.txt", _summary_text(summary, profile_label))
        for item in items:
            name = feedback_filename(item)
            base = name
            suffix = 2
            while name in used:
                name = f"{base[:-4]}_{suffix}.txt"
                suffix += 1
            used.add(name)
            archive.writestr(
                name,
                feedback_text(item, profile_label=profile_label, max_issues=max_issues),
            )
    return out
