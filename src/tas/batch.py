"""教师端批量检查与班级汇总（需求 FR-06）。

能力：
- :func:`check_files` 批量检查若干 .docx；**单份失败不中断整批**（失败原因记在
  :attr:`DocResult.error`，issues 为空）；
- :func:`check_directory` 检查目录下的 .docx（默认 ``*.docx``），自动跳过 Word
  临时文件（``~$xxx.docx``）与子目录；
- :func:`summarize` 把逐份结果聚合成 :class:`BatchSummary`：通过率、error/warning/info
  计数、共性问题 Top 10、逐份清单；
- :func:`run_batch` 便捷入口（文件列表 / 目录 → 结果 + 汇总）。

设计约束：
- 只依赖稳定接口 ``DocModel.load`` / ``roles.classify`` / ``run_checks`` /
  ``RuleSet`` / ``Issue``；**不写死规则 id 与规则数量**，检查器随迭代增加时统计自动适配；
- 教师端只诊断、不修改文件（本模块不写任何 docx）；
- :attr:`BatchSummary.pass_rate` 定义：**检查成功的文档中，没有 error 级问题的比例**
  （加载失败的文档不计入分母，单独在 ``failed`` 里报告）。

``per_doc`` 字段（便于 CLI / Web / 导出直接消费）::

    {
        "path": 绝对或调用方传入的路径字符串,
        "counts": {"error": n, "warning": n, "info": n},
        "error": n,          # error 级问题数（扁平字段，兼容 CLI/Web 表格）
        "warning": n,
        "info": n,
        "passed": bool,      # 检查成功且无 error
        "status": "通过" | "未通过" | "失败",
        "duration_s": float, # 单份检查耗时（秒）
        "load_error": str|None,  # 加载/检查异常信息（仅失败文档）
    }
"""

from __future__ import annotations

import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

from .checkers import run_checks
from .config import DEFAULT_PROFILE, RuleSet
from .docx_io import DocModel
from .report import Issue
from .roles import classify

__all__ = [
    "SEVERITY_KEYS",
    "TOP_RULES_LIMIT",
    "DocResult",
    "BatchSummary",
    "check_files",
    "check_directory",
    "summarize",
    "run_batch",
]

#: 汇总统计的严重级别顺序（与 report.Report.counts 保持一致）
SEVERITY_KEYS: tuple[str, ...] = ("error", "warning", "info")
#: 共性问题 Top N
TOP_RULES_LIMIT = 10
#: Word 临时文件 / 锁文件前缀（批量导入必须跳过）
TEMP_FILE_PREFIXES: tuple[str, ...] = ("~$", ".~lock.")


def _now() -> str:
    """本地时区的 ISO 时间（秒级），便于报告可读且可排序。"""
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


# ---------------------------------------------------------------------------
# 逐份结果
# ---------------------------------------------------------------------------
@dataclass
class DocResult:
    """单份文档的检查结果（失败也返回，不抛异常）。"""

    path: Path
    issues: list[Issue] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None
    duration_s: float = 0.0

    # -- 便捷属性 ---------------------------------------------------------
    @property
    def ok(self) -> bool:
        """是否检查成功（与文档是否合规无关）。"""
        return self.error is None

    def counts(self) -> dict[str, int]:
        counter = Counter(issue.severity for issue in self.issues)
        return {key: counter.get(key, 0) for key in SEVERITY_KEYS}

    def has_errors(self) -> bool:
        return any(issue.severity == "error" for issue in self.issues)

    def passed(self) -> bool:
        """检查成功且没有 error 级问题 = 通过。"""
        return self.ok and not self.has_errors()

    def status(self) -> str:
        if self.error is not None:
            return "失败"
        return "通过" if self.passed() else "未通过"

    def rule_ids(self) -> set[str]:
        return {issue.rule_id for issue in self.issues}

    def to_dict(self) -> dict[str, Any]:
        counts = self.counts()
        return {
            "path": str(self.path),
            "issues": [issue.to_dict() for issue in self.issues],
            "counts": counts,
            "error": self.error,
            "duration_s": round(float(self.duration_s or 0.0), 6),
            "stats": dict(self.stats),
        }


# ---------------------------------------------------------------------------
# 单份检查
# ---------------------------------------------------------------------------
def _check_one(path: Path, ruleset: RuleSet, profile: str) -> DocResult:
    """检查一份文档；任何异常都收敛为 ``DocResult.error``，绝不向上抛。"""
    started = time.perf_counter()
    try:
        model = DocModel.load(path)
        roles = classify(model, profile)
        issues = run_checks(model, roles, ruleset, profile=profile)
        stats = dict(model.stats())
    except Exception as exc:  # noqa: BLE001 - 单份失败必须不中断批量
        return DocResult(
            path=path,
            issues=[],
            stats={},
            error=f"{type(exc).__name__}: {exc}",
            duration_s=time.perf_counter() - started,
        )
    return DocResult(
        path=path,
        issues=issues,
        stats=stats,
        error=None,
        duration_s=time.perf_counter() - started,
    )


def _as_paths(paths: Iterable[Path | str]) -> list[Path]:
    return [item if isinstance(item, Path) else Path(item) for item in paths]


def check_files(
    paths: Sequence[Path | str],
    *,
    ruleset: RuleSet,
    profile: str = DEFAULT_PROFILE,
    ruleset_id: str = "",
) -> list[DocResult]:
    """批量检查文件列表；返回顺序与入参一致，单份失败记 error 后继续。

    ``ruleset_id`` 只做批次标识（由 :func:`summarize` 写入汇总），本函数不做规则过滤。
    """
    results: list[DocResult] = []
    for path in _as_paths(paths):
        results.append(_check_one(path, ruleset, profile))
    return results


def check_directory(
    directory: Path | str,
    *,
    patterns: tuple[str, ...] = ("*.docx",),
    ruleset: RuleSet,
    profile: str = DEFAULT_PROFILE,
    ruleset_id: str = "",
) -> list[DocResult]:
    """检查目录下的文档（默认 ``*.docx``）。

    - 只收顶层文件：子目录一律跳过（不做递归）；
    - 跳过 Word 临时文件 / 锁文件（``~$``、``.~lock.`` 前缀）；
    - 文件按名称排序，保证同目录多次运行结果顺序稳定。
    """
    directory = Path(directory)
    if not directory.is_dir():
        raise FileNotFoundError(f"目录不存在：{directory}")

    found: dict[Path, Path] = {}
    for pattern in patterns:
        try:
            candidates = list(directory.glob(pattern))
        except (ValueError, OSError):  # pragma: no cover - 非法 glob 兜底
            continue
        for candidate in candidates:
            if not candidate.is_file():
                continue
            if candidate.parent != directory:
                continue  # 跳过子目录（pattern 含 ** / 子路径时）
            if candidate.name.startswith(TEMP_FILE_PREFIXES):
                continue
            found.setdefault(candidate.resolve(), candidate)

    paths = sorted(found.values(), key=lambda item: (item.name.casefold(), str(item)))
    return check_files(paths, ruleset=ruleset, profile=profile, ruleset_id=ruleset_id)


# ---------------------------------------------------------------------------
# 班级汇总
# ---------------------------------------------------------------------------
@dataclass
class BatchSummary:
    """一次批量检查的班级汇总。"""

    total: int
    checked: int
    failed: int
    pass_rate: float
    counts: dict[str, int]
    top_rules: list[tuple[str, int]]
    per_doc: list[dict[str, Any]]
    ruleset_id: str = ""
    generated_at: str = ""

    @property
    def pass_count(self) -> int:
        """通过（检查成功且无 error）的文档数。"""
        return sum(1 for item in self.per_doc if item.get("passed"))

    @property
    def total_issues(self) -> int:
        return sum(self.counts.get(key, 0) for key in SEVERITY_KEYS)

    def to_dict(self) -> dict[str, Any]:
        return {
            "total": self.total,
            "checked": self.checked,
            "failed": self.failed,
            "pass_rate": self.pass_rate,
            "pass_count": self.pass_count,
            "counts": dict(self.counts),
            "top_rules": [[rule_id, count] for rule_id, count in self.top_rules],
            "per_doc": [dict(item) for item in self.per_doc],
            "ruleset_id": self.ruleset_id,
            "generated_at": self.generated_at,
        }


def summarize(results: Sequence[DocResult], *, ruleset_id: str = "") -> BatchSummary:
    """把逐份结果聚合成班级汇总。

    - ``total`` / ``checked`` / ``failed``：总份数 / 检查成功数 / 失败数；
    - ``pass_rate``：检查成功文档中「无 error 级问题」的比例（分母不含失败文档）；
    - ``counts``：所有文档的 error / warning / info 总数；
    - ``top_rules``：按命中数降序的规则 Top 10（同数按 rule_id 排序，结果稳定）；
    - ``per_doc``：逐份清单（字段见模块 docstring）。
    """
    items = list(results)
    total = len(items)
    checked = sum(1 for item in items if item.error is None)
    failed = total - checked

    counts: dict[str, int] = {key: 0 for key in SEVERITY_KEYS}
    rule_counter: Counter[str] = Counter()
    per_doc: list[dict[str, Any]] = []

    for item in items:
        doc_counts = item.counts()
        for key in SEVERITY_KEYS:
            counts[key] += doc_counts[key]
        for issue in item.issues:
            rule_counter[issue.rule_id] += 1

        passed = item.error is None and doc_counts["error"] == 0
        per_doc.append(
            {
                "path": str(item.path),
                "counts": doc_counts,
                "error": doc_counts["error"],
                "warning": doc_counts["warning"],
                "info": doc_counts["info"],
                "passed": passed,
                "status": item.status(),
                "duration_s": round(float(item.duration_s or 0.0), 6),
                "load_error": item.error,
            }
        )

    passed_count = sum(1 for item in per_doc if item["passed"])
    pass_rate = (passed_count / checked) if checked else 0.0

    top_rules = sorted(rule_counter.items(), key=lambda pair: (-pair[1], pair[0]))[
        :TOP_RULES_LIMIT
    ]

    return BatchSummary(
        total=total,
        checked=checked,
        failed=failed,
        pass_rate=pass_rate,
        counts=counts,
        top_rules=[(rule_id, count) for rule_id, count in top_rules],
        per_doc=per_doc,
        ruleset_id=ruleset_id,
        generated_at=_now(),
    )


# ---------------------------------------------------------------------------
# 便捷入口
# ---------------------------------------------------------------------------
def run_batch(
    paths: Optional[Sequence[Path | str]] = None,
    directory: Optional[Path | str] = None,
    *,
    ruleset: Optional[RuleSet] = None,
    profile: str = DEFAULT_PROFILE,
    ruleset_id: str = "",
    patterns: tuple[str, ...] = ("*.docx",),
) -> tuple[list[DocResult], BatchSummary]:
    """批量检查入口：``paths``（文件列表）和/或 ``directory``（目录）二选一或组合。

    ``ruleset`` 省略时加载默认规则库（``rules/rules.json``）。返回
    ``(逐份结果, 班级汇总)``；调用方也可只用 :func:`check_files` /
    :func:`check_directory` + :func:`summarize` 拆开调用。
    """
    if paths is None and directory is None:
        raise ValueError("run_batch 需要 paths 或 directory 至少提供一个")
    rules = ruleset if ruleset is not None else RuleSet.load()

    results: list[DocResult] = []
    if paths is not None:
        results.extend(
            check_files(paths, ruleset=rules, profile=profile, ruleset_id=ruleset_id)
        )
    if directory is not None:
        results.extend(
            check_directory(
                directory,
                patterns=patterns,
                ruleset=rules,
                profile=profile,
                ruleset_id=ruleset_id,
            )
        )
    return results, summarize(results, ruleset_id=ruleset_id)
