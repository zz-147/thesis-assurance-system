"""教师端「批量安全修复」：一次生成全部 ``*_fixed.docx`` + 报告 + 可下载 zip。

学生端修复（:mod:`tas.fixer`）是单份文档的；老师收齐一个班的 N 份文档后，
逐份点「修复」不现实，本模块把同一套安全修复批量跑完：

    finish 后的产物
    ├─ <out_dir>/<stem>_fixed.docx      # 每份一份修复版（原文件绝不覆盖）
    ├─ <out_dir>/...                    # 同名冲突自动追加 -2、-3 …
    ├─ fix.zip
    │  ├─ <stem>_fixed.docx ...
    │  ├─ 修复报告.txt                  # 逐份：原文件/输出文件/修复前后计数/动作/剩余/失败原因
    │  └─ 修复汇总.txt                  # 整批统计 + 逐份清单
    └─ summarize_fix(results)           # 供 CLI / Web / 导出直接消费的 dict

设计约束：
- **复用** :func:`tas.fixer.fix_document`：只改格式、不改文字、文献/目录条目不碰、
  原文件不覆盖、修复后自动复检 + 内容 hash 校验；
- 单份失败（损坏 / 加密 / 非 docx / 修复异常）只记 :attr:`BatchFixResult.error`，
  **绝不中断整批**；
- 修复前计数复用教师端批量检查的 :func:`tas.batch.check_files`（与报告口径一致）；
- 输出固定写到 ``out_dir``（相对路径即相对当前工作目录），文件名 ``<stem>_fixed.docx``，
  已存在或批内重名时追加 ``-2``、``-3`` …，不覆盖任何已有文件；
- dry-run 不创建目录、不写任何文件，只返回将要做的事与复检结果。
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

from .batch import SEVERITY_KEYS, TEMP_FILE_PREFIXES, check_files
from .config import DEFAULT_PROFILE, RuleSet
from .fixer import fix_document

__all__ = [
    "SEVERITY_KEYS",
    "TEMP_FILE_PREFIXES",
    "REPORT_NAME",
    "SUMMARY_NAME",
    "BatchFixResult",
    "fix_files",
    "fix_directory",
    "summarize_fix",
    "write_fix_zip",
]

#: 修复版文件名后缀
OUTPUT_SUFFIX = "_fixed"
#: zip 内报告名（固定，便于老师直接打开）
REPORT_NAME = "修复报告.txt"
SUMMARY_NAME = "修复汇总.txt"

_MAX_ERROR_CHARS = 300


def _now() -> str:
    """本地时区的 ISO 时间（秒级），报告可读且可排序。"""
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _empty_counts() -> dict[str, int]:
    return {key: 0 for key in SEVERITY_KEYS}


def _counts_from_issues(issues: Iterable[Any]) -> dict[str, int]:
    """把 Issue 列表收敛成 error/warning/info 计数（未知级别忽略，保证三键稳定）。"""
    counts = _empty_counts()
    for issue in issues:
        severity = getattr(issue, "severity", None)
        if severity in counts:
            counts[severity] += 1
    return counts


def _safe_counts(value: Optional[dict]) -> dict[str, int]:
    counts = _empty_counts()
    for key in SEVERITY_KEYS:
        try:
            counts[key] = int((value or {}).get(key, 0) or 0)
        except (TypeError, ValueError):  # pragma: no cover - 脏数据兜底
            counts[key] = 0
    return counts


def _error_text(exc: BaseException) -> str:
    text = f"{type(exc).__name__}: {exc}".strip()
    return text[:_MAX_ERROR_CHARS]


def _total(counts: dict[str, int]) -> int:
    return sum(counts.get(key, 0) for key in SEVERITY_KEYS)


# ---------------------------------------------------------------------------
# 逐份结果
# ---------------------------------------------------------------------------
@dataclass
class BatchFixResult:
    """单份文档的批量修复结果（失败也返回，不抛异常）。"""

    src: Path
    out: Optional[Path] = None
    dry_run: bool = False
    error: Optional[str] = None
    before: dict[str, int] = field(default_factory=_empty_counts)
    after: Optional[dict[str, int]] = None
    actions: int = 0
    skipped: int = 0
    recheck_error: int = 0
    recheck_warning: int = 0
    recheck_info: int = 0

    # -- 便捷属性 ---------------------------------------------------------
    @property
    def ok(self) -> bool:
        """是否修复成功（与修复后是否还有问题无关）。"""
        return self.error is None

    @property
    def failed(self) -> bool:
        return self.error is not None

    def status(self) -> str:
        if self.error is not None:
            return "失败"
        return "dry-run" if self.dry_run else "已修复"

    def recheck_counts(self) -> dict[str, int]:
        return {
            "error": self.recheck_error,
            "warning": self.recheck_warning,
            "info": self.recheck_info,
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "src": str(self.src),
            "out": str(self.out) if self.out is not None else None,
            "dry_run": bool(self.dry_run),
            "status": self.status(),
            "error": self.error,
            "before": _safe_counts(self.before),
            "after": _safe_counts(self.after) if self.after is not None else None,
            "actions": int(self.actions),
            "skipped": int(self.skipped),
            "recheck": self.recheck_counts(),
            "recheck_error": int(self.recheck_error),
            "recheck_warning": int(self.recheck_warning),
            "recheck_info": int(self.recheck_info),
        }


# ---------------------------------------------------------------------------
# 输出命名：<stem>_fixed.docx，冲突追加 -2、-3 …
# ---------------------------------------------------------------------------
def _allocate_out_path(src: Path, out_dir: Path, used: set[str]) -> Path:
    stem = src.stem or "document"
    candidate = out_dir / f"{stem}{OUTPUT_SUFFIX}.docx"
    suffix = 2
    while candidate.name.casefold() in used or candidate.exists():
        candidate = out_dir / f"{stem}{OUTPUT_SUFFIX}-{suffix}.docx"
        suffix += 1
    used.add(candidate.name.casefold())
    return candidate


# ---------------------------------------------------------------------------
# 单份修复
# ---------------------------------------------------------------------------
def _fix_one(
    src: Path,
    checked: Any,
    out_dir: Path,
    ruleset: RuleSet,
    profile: str,
    dry_run: bool,
    used_names: set[str],
) -> BatchFixResult:
    """修复单份；任何异常都收敛为 ``error``，绝不向上抛。"""
    before = _safe_counts(checked.counts() if checked is not None else None)
    if checked is not None and getattr(checked, "error", None):
        return BatchFixResult(
            src=src,
            out=None,
            dry_run=dry_run,
            error=f"检查失败：{checked.error}",
            before=before,
        )

    out_path: Optional[Path] = None
    try:
        if not dry_run:
            out_path = _allocate_out_path(src, out_dir, used_names)
        result = fix_document(
            src,
            out_path,
            profile=profile,
            ruleset=ruleset,
            dry_run=dry_run,
        )
    except Exception as exc:  # noqa: BLE001 - 单份失败必须不中断批量
        # 修复失败时不留半成品输出文件（该文件名是本次批处理刚分配的，删掉是安全的）
        if out_path is not None:
            try:
                if out_path.is_file():
                    out_path.unlink()
            except OSError:  # pragma: no cover - 清理失败不影响结果报告
                pass
        return BatchFixResult(
            src=src,
            out=None,
            dry_run=dry_run,
            error=_error_text(exc),
            before=before,
        )

    recheck = _counts_from_issues(result.recheck_issues)
    actual_out = Path(result.out_path) if result.out_path is not None else None
    return BatchFixResult(
        src=src,
        out=actual_out,
        dry_run=bool(dry_run),
        error=None,
        before=before,
        after=recheck,
        actions=len(result.actions),
        skipped=len(result.skipped),
        recheck_error=recheck["error"],
        recheck_warning=recheck["warning"],
        recheck_info=recheck["info"],
    )


def _as_paths(paths: Iterable[Path | str]) -> list[Path]:
    return [item if isinstance(item, Path) else Path(item) for item in paths]


# ---------------------------------------------------------------------------
# 批量入口
# ---------------------------------------------------------------------------
def fix_files(
    paths: Sequence[Path | str],
    out_dir: Path | str,
    *,
    ruleset: RuleSet,
    profile: str = DEFAULT_PROFILE,
    dry_run: bool = False,
) -> list[BatchFixResult]:
    """批量修复文件列表，返回与入参同序的逐份结果。

    - ``out_dir``：修复版输出目录（相对路径相对当前工作目录）；原文件不会被修改；
    - 单份失败只记 ``error``，不中断整批；
    - ``dry_run=True`` 时不创建输出目录、不写任何文件，``out`` 均为 None。
    """
    out_root = Path(out_dir)
    sources = _as_paths(paths)
    # 修复前计数复用教师端批量检查（口径与检查报告一致）
    checks = check_files(sources, ruleset=ruleset, profile=profile)

    if not dry_run:
        try:
            out_root.mkdir(parents=True, exist_ok=True)
        except OSError:  # pragma: no cover - 交给逐份保存报错，便于定位
            pass

    used_names: set[str] = set()
    results: list[BatchFixResult] = []
    for src, checked in zip(sources, checks):
        results.append(
            _fix_one(src, checked, out_root, ruleset, profile, dry_run, used_names)
        )
    return results


def fix_directory(
    directory: Path | str,
    out_dir: Path | str,
    *,
    patterns: tuple[str, ...] = ("*.docx",),
    ruleset: RuleSet,
    profile: str = DEFAULT_PROFILE,
    dry_run: bool = False,
) -> list[BatchFixResult]:
    """批量修复目录下的文档（默认 ``*.docx``）。

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
    return fix_files(
        paths,
        out_dir,
        ruleset=ruleset,
        profile=profile,
        dry_run=dry_run,
    )


# ---------------------------------------------------------------------------
# 汇总
# ---------------------------------------------------------------------------
def summarize_fix(results: Sequence[BatchFixResult]) -> dict[str, Any]:
    """把逐份修复结果聚合成整批汇总（供 CLI / Web / 报告直接消费）。

    字段：
        total         文档总数
        fixed         处理成功份数（修复成功；dry-run 下为“将修复”份数）
        failed        失败份数（error 非空）
        actions       修复动作合计
        skipped       跳过/需人工项合计
        before_counts 修复前 error/warning/info 合计
        after_counts  修复后 error/warning/info 合计（失败份按 0 计）
        before_total / after_total  问题总数合计
        improved      修复后 error 数变少的文档数
        resolved_errors  消除的 error 级问题总数
        dry_run       是否整批 dry-run
    """
    items = list(results)
    total = len(items)
    failed = sum(1 for item in items if item.error is not None)
    fixed = total - failed

    before_counts = _empty_counts()
    after_counts = _empty_counts()
    for item in items:
        item_before = _safe_counts(item.before)
        for key in SEVERITY_KEYS:
            before_counts[key] += item_before[key]
        if item.after is not None:
            item_after = _safe_counts(item.after)
            for key in SEVERITY_KEYS:
                after_counts[key] += item_after[key]

    improved = 0
    resolved_errors = 0
    for item in items:
        if item.error is not None or item.after is None:
            continue
        before_error = _safe_counts(item.before)["error"]
        after_error = _safe_counts(item.after)["error"]
        if after_error < before_error:
            improved += 1
        resolved_errors += max(0, before_error - after_error)

    return {
        "total": total,
        "fixed": fixed,
        "failed": failed,
        "actions": sum(int(item.actions) for item in items),
        "skipped": sum(int(item.skipped) for item in items),
        "before_counts": before_counts,
        "after_counts": after_counts,
        "before_total": _total(before_counts),
        "after_total": _total(after_counts),
        "improved": improved,
        "resolved_errors": resolved_errors,
        "dry_run": bool(items) and all(item.dry_run for item in items),
    }


# ---------------------------------------------------------------------------
# 报告文本 + zip
# ---------------------------------------------------------------------------
def _fmt_counts(counts: dict[str, int]) -> str:
    return (
        f"错误 {counts.get('error', 0)} ｜ "
        f"警告 {counts.get('warning', 0)} ｜ "
        f"提示 {counts.get('info', 0)}"
    )


def _render_report(results: Sequence[BatchFixResult]) -> str:
    """逐份修复报告（修复报告.txt）。"""
    items = list(results)
    summary = summarize_fix(items)
    lines = [
        "tas 批量安全修复报告（学生端）",
        "=" * 64,
        f"生成时间：{_now()}",
        (
            f"文件总数：{summary['total']} ｜ 成功：{summary['fixed']} ｜ "
            f"失败：{summary['failed']}"
        ),
        "说明：只改格式、不改文字；原文件均未被覆盖；参考文献条目与目录条目不自动修改。",
        "",
    ]
    for index, result in enumerate(items, 1):
        before = _safe_counts(result.before)
        lines.append(f"[{index}] 原文件：{result.src}")
        if result.out is not None:
            lines.append(f"    输出文件：{result.out}")
        elif result.dry_run:
            lines.append("    输出文件：（dry-run，未写文件）")
        else:
            lines.append("    输出文件：（未生成）")
        lines.append(f"    修复前：{_fmt_counts(before)}")

        if result.error is not None:
            lines.append("    修复后：（未执行）")
            lines.append(f"    状态：失败")
            lines.append(f"    失败原因：{result.error}")
            lines.append("")
            continue

        after = _safe_counts(result.after)
        lines.append(f"    修复后：{_fmt_counts(after)}")
        lines.append(
            f"    修复动作：{result.actions} 项；跳过/需人工：{result.skipped} 项"
        )
        detail = "；".join(f"{key}={after[key]}" for key in SEVERITY_KEYS)
        lines.append(f"    剩余问题：{_total(after)} 项（{detail}）")
        lines.append(f"    状态：{result.status()}")
        lines.append("")
    return "\n".join(lines)


def _render_summary(results: Sequence[BatchFixResult]) -> str:
    """整批修复汇总（修复汇总.txt）。"""
    items = list(results)
    summary = summarize_fix(items)
    before = summary["before_counts"]
    after = summary["after_counts"]
    lines = [
        "tas 批量安全修复汇总",
        "=" * 64,
        f"生成时间：{_now()}",
        (
            f"文件总数：{summary['total']} ｜ 成功：{summary['fixed']} ｜ "
            f"失败：{summary['failed']}"
        ),
        (
            f"修复动作合计：{summary['actions']} 项；"
            f"跳过/需人工：{summary['skipped']} 项"
        ),
        f"修复前：{_fmt_counts(before)}",
        f"修复后：{_fmt_counts(after)}",
        (
            f"改善文档：{summary['improved']} 份；"
            f"消除错误：{summary['resolved_errors']} 条"
        ),
        "",
        "逐份清单：",
    ]
    for index, result in enumerate(items, 1):
        name = Path(result.src).name
        if result.error is not None:
            lines.append(f"  {index}. {name}：失败（{result.error}）")
            continue
        item_before = _safe_counts(result.before)
        item_after = _safe_counts(result.after)
        lines.append(
            f"  {index}. {name}：修复前 "
            f"{item_before['error']}/{item_before['warning']}/{item_before['info']} → 修复后 "
            f"{item_after['error']}/{item_after['warning']}/{item_after['info']}"
            f"（动作 {result.actions}）"
        )
    return "\n".join(lines)


def _unique_zip_name(name: str, used: set[str]) -> str:
    if name not in used:
        used.add(name)
        return name
    path = Path(name)
    suffix = 2
    while True:
        candidate = f"{path.stem}-{suffix}{path.suffix}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        suffix += 1


def write_fix_zip(results: Sequence[BatchFixResult], zip_path: Path | str) -> Path:
    """把修复后的 docx（若已生成）+ 修复报告 + 修复汇总写成 zip；返回 zip 路径。

    - dry-run / 失败的文件没有输出文件，不会进入 zip（报告中仍会记录原因）；
    - zip 内同名文件自动追加 ``-2`` 等后缀（正常批量修复输出名已唯一，不会触发）。
    """
    out = Path(zip_path)
    items = list(results)
    out.parent.mkdir(parents=True, exist_ok=True)

    used_names: set[str] = set()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for result in items:
            candidate = Path(result.out) if result.out is not None else None
            if candidate is None or not candidate.is_file():
                continue
            arcname = _unique_zip_name(candidate.name, used_names)
            archive.write(candidate, arcname)
        archive.writestr(REPORT_NAME, _render_report(items).encode("utf-8"))
        archive.writestr(SUMMARY_NAME, _render_summary(items).encode("utf-8"))
    return out
