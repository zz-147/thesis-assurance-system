"""教师端「批量文献真实性核验」：逐份解析 + 联网核验（可注入 fetcher）+ JSON/Markdown 报告。

单份核验见 :mod:`tas.refs.verify`；老师收齐一个班的 N 份论文后需要「一次跑完 + 汇总」。
本模块只做编排与报告，复用 :func:`tas.refs.parse.parse_references` 与
:func:`tas.refs.verify.verify_references`，**不修改任何文档**。

设计约束：
- **单份失败不中断整批**：损坏 / 加密 / 不存在 / 非 docx 等失败只写入
  :attr:`BatchRefDocResult.error`，其余文档照常核验；
- 没有参考文献的文档 ``entries == 0``，不算失败；
- 每份文档默认最多核验 ``limit_per_doc=5`` 条：``entries`` 记录解析到的总条数，
  ``results`` 是实际核验的前 N 条（``None`` 表示全部核验），避免一次批量打爆公共 API；
- 所有网络访问只通过注入的 ``fetcher`` 走 :func:`verify_references`；测试注入
  fake fetcher 即可完全离线；
- **诚实性红线**：``not_found`` 只表示「在当前索引未检索到」，中文期刊可能未被收录，
  报告必须带「需人工核实」提示；本模块绝不输出「文献造假」结论、绝不修改文档。

``per_doc`` 字段（便于 CLI / Web / JSON 直接消费）::

    {
        "doc": 调用方传入的路径字符串,
        "name": 文件名,
        "entries": 解析到的参考文献条数,
        "checked": 实际核验条数,          # 受 limit_per_doc 影响
        "verified": n, "likely": n, "not_found": n,
        "error": n,                       # 文献级 error 结果数（与 batch.py 口径一致）
        "failed": bool,                   # 文档加载/解析失败
        "load_error": None | "异常信息",
        "duration_s": float,
    }
"""

from __future__ import annotations

import json
import time
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Sequence

from .config import DEFAULT_PROFILE
from .docx_io import DocModel
from .refs import DEFAULT_MAILTO, VerificationResult, parse_references, verify_references
from .refs.verify import Fetcher
from .roles import classify

__all__ = [
    "STATUS_KEYS",
    "STATUS_LABELS",
    "BatchRefDocResult",
    "BatchRefSummary",
    "verify_batch",
    "render_refs_markdown",
    "write_refs_report",
]

#: 核验状态顺序（与 :class:`VerificationResult.status` 一致）
STATUS_KEYS: tuple[str, ...] = ("verified", "likely", "not_found", "error")

#: 报告里的中文状态标签（具体状态码在括号里补充，便于老师对照 JSON）
STATUS_LABELS: dict[str, str] = {
    "verified": "✅ 已核验",
    "likely": "🟡 部分匹配",
    "not_found": "⚪ 未检索到",
    "error": "❌ 接口错误",
}

#: 错误信息最大长度（与 batch_fix 保持一致，避免 JSON 里塞入超长 traceback）
_MAX_ERROR_CHARS = 300


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------
@dataclass
class BatchRefDocResult:
    """单份文档的批量核验结果（失败也返回，不抛异常）。"""

    doc: Path
    entries: int
    results: list[VerificationResult] = field(default_factory=list)
    error: Optional[str] = None
    duration_s: float = 0.0

    # -- 便捷属性 ---------------------------------------------------------
    @property
    def ok(self) -> bool:
        """文档是否加载/解析成功（与该文档文献是否核验通过无关）。"""
        return self.error is None

    @property
    def failed(self) -> bool:
        return self.error is not None

    def counts(self) -> dict[str, int]:
        """按核验状态计数（未知状态忽略，保证四个键稳定）。"""
        counter = Counter(
            str(getattr(result, "status", "") or "") for result in self.results
        )
        return {key: int(counter.get(key, 0)) for key in STATUS_KEYS}

    def status(self) -> str:
        return "失败" if self.error is not None else "已核验"

    def to_dict(self) -> dict[str, Any]:
        counts = self.counts()
        return {
            "doc": str(self.doc),
            "name": Path(self.doc).name,
            "entries": int(self.entries or 0),
            "checked": len(self.results),
            "results": [_verification_to_dict(result) for result in self.results],
            "counts": counts,
            "verified": counts["verified"],
            "likely": counts["likely"],
            "not_found": counts["not_found"],
            "status_error": counts["error"],
            "failed": self.error is not None,
            "error": self.error,
            "load_error": self.error,
            "duration_s": round(float(self.duration_s or 0.0), 6),
        }


@dataclass
class BatchRefSummary:
    """一次批量核验的汇总。"""

    total_docs: int
    checked_docs: int
    failed_docs: int
    total_entries: int
    verified: int
    likely: int
    not_found: int
    error: int
    per_doc: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "total_docs": int(self.total_docs),
            "checked_docs": int(self.checked_docs),
            "failed_docs": int(self.failed_docs),
            "total_entries": int(self.total_entries),
            "verified": int(self.verified),
            "likely": int(self.likely),
            "not_found": int(self.not_found),
            "error": int(self.error),
            "per_doc": [dict(item) for item in self.per_doc],
        }


# ---------------------------------------------------------------------------
# 序列化辅助（与 cli.py 的 refs JSON 形状保持一致）
# ---------------------------------------------------------------------------
_ENTRY_FIELDS: tuple[str, ...] = (
    "paragraph_index",
    "number",
    "raw",
    "authors",
    "title",
    "venue",
    "year",
    "doc_type",
    "doi",
    "url",
)

_RESULT_FIELDS: tuple[str, ...] = (
    "status",
    "source",
    "matched_title",
    "matched_year",
    "matched_venue",
    "doi",
    "url",
    "score",
    "note",
    "evidence",
)


def _entry_to_dict(entry: Any) -> Optional[dict[str, Any]]:
    if entry is None:
        return None
    return {name: getattr(entry, name, None) for name in _ENTRY_FIELDS}


def _verification_to_dict(result: Any) -> dict[str, Any]:
    data: dict[str, Any] = {
        name: getattr(result, name, None) for name in _RESULT_FIELDS
    }
    data["evidence"] = list(getattr(result, "evidence", None) or [])
    data["entry"] = _entry_to_dict(getattr(result, "entry", None))
    return data


# ---------------------------------------------------------------------------
# 批量核验
# ---------------------------------------------------------------------------
def verify_batch(
    paths: Sequence[Path | str],
    *,
    profile: str = DEFAULT_PROFILE,
    limit_per_doc: Optional[int] = 5,
    timeout: float = 10.0,
    pause: float = 0.1,
    mailto: str = DEFAULT_MAILTO,
    fetcher: Optional[Fetcher] = None,
) -> tuple[list[BatchRefDocResult], BatchRefSummary]:
    """批量核验若干 .docx 的参考文献；单份失败只记 error，不中断整批。

    - 每份文档：``DocModel.load`` → ``classify`` → ``parse_references`` →
      ``verify_references(entries, limit=limit_per_doc, fetcher=...)``；
    - ``limit_per_doc=None`` 表示每份文档全部核验；
    - 返回 ``(逐份结果, 汇总)``；结果顺序与入参一致。
    """
    docs: list[BatchRefDocResult] = []
    for raw_path in paths:
        path = raw_path if isinstance(raw_path, Path) else Path(raw_path)
        started = time.perf_counter()
        try:
            model = DocModel.load(path)
            roles = classify(model, profile)
            entries = list(parse_references(model, roles))
            results = verify_references(
                entries,
                limit=limit_per_doc,
                pause=pause,
                timeout=timeout,
                mailto=mailto,
                fetcher=fetcher,
            )
        except Exception as exc:  # noqa: BLE001 - 单份失败必须不中断批量
            docs.append(
                BatchRefDocResult(
                    doc=path,
                    entries=0,
                    results=[],
                    error=_error_text(exc),
                    duration_s=time.perf_counter() - started,
                )
            )
            continue
        docs.append(
            BatchRefDocResult(
                doc=path,
                entries=len(entries),
                results=results,
                error=None,
                duration_s=time.perf_counter() - started,
            )
        )
    return docs, _summarize(docs)


def _summarize(docs: Sequence[BatchRefDocResult]) -> BatchRefSummary:
    """把逐份结果聚合成整批汇总（document 级失败与文献级 error 分开统计）。"""
    items = list(docs)
    totals = {key: 0 for key in STATUS_KEYS}
    total_entries = 0
    checked_docs = 0
    failed_docs = 0
    per_doc: list[dict[str, Any]] = []

    for item in items:
        counts = item.counts()
        for key in STATUS_KEYS:
            totals[key] += counts[key]
        total_entries += max(0, int(item.entries or 0))
        if item.error is None:
            checked_docs += 1
        else:
            failed_docs += 1
        per_doc.append(
            {
                "doc": str(item.doc),
                "name": Path(item.doc).name,
                "entries": max(0, int(item.entries or 0)),
                "checked": len(item.results),
                "verified": counts["verified"],
                "likely": counts["likely"],
                "not_found": counts["not_found"],
                "error": counts["error"],
                "failed": item.error is not None,
                "load_error": item.error,
                "duration_s": round(float(item.duration_s or 0.0), 6),
            }
        )

    return BatchRefSummary(
        total_docs=len(items),
        checked_docs=checked_docs,
        failed_docs=failed_docs,
        total_entries=total_entries,
        verified=totals["verified"],
        likely=totals["likely"],
        not_found=totals["not_found"],
        error=totals["error"],
        per_doc=per_doc,
    )


# ---------------------------------------------------------------------------
# Markdown 报告
# ---------------------------------------------------------------------------
def render_refs_markdown(
    summary: BatchRefSummary,
    docs: Sequence[BatchRefDocResult],
) -> str:
    """渲染中文 Markdown 报告：汇总表 + 逐条明细 + 失败清单 + 诚实性提示。"""
    items = list(docs)
    lines: list[str] = [
        "# 批量参考文献真实性核验报告",
        "",
        f"- 文档总数：**{summary.total_docs}**"
        f"（成功 {summary.checked_docs} ｜ 失败 {summary.failed_docs}）",
        f"- 参考文献总数：**{summary.total_entries}**",
        "- 核验结果："
        f"✅ verified **{summary.verified}** ｜ "
        f"🟡 likely **{summary.likely}** ｜ "
        f"⚪ not_found **{summary.not_found}** ｜ "
        f"❌ error **{summary.error}**",
        "",
        "> 诚实性提示：not_found 仅表示未在当前索引检索到，中文期刊可能未被收录，"
        "需人工核实；未检索到 ≠ 文献造假。",
        "",
    ]

    partial = [item for item in items if len(item.results) < max(0, int(item.entries or 0))]
    if partial:
        lines.append(
            "> 核验上限：部分文档只核验了前若干条文献（见明细中的“已核验 X/Y 条”），"
            "其余条目未联网核对。"
        )
        lines.append("")

    # ---- 汇总表 ---------------------------------------------------------
    lines.append("## 汇总")
    lines.append("")
    lines.append("| 文件 | 参考文献数 | verified | likely | not_found | error |")
    lines.append("| --- | ---: | ---: | ---: | ---: | ---: |")
    rows = items if items else []
    if rows:
        for item in rows:
            counts = item.counts()
            name = _doc_name(item)
            if item.error is not None:
                name += "（加载失败）"
            lines.append(
                f"| {_md_cell(name)} | {max(0, int(item.entries or 0))} | "
                f"{counts['verified']} | {counts['likely']} | "
                f"{counts['not_found']} | {counts['error']} |"
            )
    else:
        lines.append("| （无文档） | 0 | 0 | 0 | 0 | 0 |")
    lines.append("")

    # ---- 逐条明细 -------------------------------------------------------
    lines.append("## 逐条明细")
    lines.append("")
    if not items:
        lines.append("（本批没有文档。）")
        lines.append("")
    for item in items:
        name = _doc_name(item)
        lines.append(f"### {name}")
        lines.append("")
        if item.error is not None:
            lines.append(f"- ⚠️ 加载/检查失败：{item.error}")
            lines.append("")
            continue
        total_entries = max(0, int(item.entries or 0))
        if not item.results:
            if total_entries:
                lines.append(
                    f"- 参考文献 {total_entries} 条；本次未核验（limit_per_doc=0）。"
                )
            else:
                lines.append("- 未解析到参考文献条目（entries=0），无需核验。")
            lines.append("")
            continue
        lines.append(
            f"- 参考文献 {total_entries} 条；已核验 {len(item.results)} 条"
        )
        lines.append("")
        lines.append(
            "| # | 题名 | 状态 | 来源 | 匹配题名 | 年份 | DOI/URL | 说明 |"
        )
        lines.append("| ---: | --- | --- | --- | --- | ---: | --- | --- |")
        for index, result in enumerate(item.results, 1):
            entry = getattr(result, "entry", None)
            title = ""
            if entry is not None:
                title = (getattr(entry, "title", "") or "").strip()
                if not title:
                    title = (getattr(entry, "raw", "") or "").strip()
            status = str(getattr(result, "status", "") or "error")
            matched = (getattr(result, "matched_title", "") or "").strip()
            year = getattr(result, "matched_year", None)
            doi = (getattr(result, "doi", "") or "").strip()
            url = (getattr(result, "url", "") or "").strip()
            note = (getattr(result, "note", "") or "").strip()
            lines.append(
                f"| {index} | {_md_cell(title)} | {_md_cell(_status_label(status))} | "
                f"{_md_cell(getattr(result, 'source', '') or '—')} | "
                f"{_md_cell(matched or '—')} | {_md_cell(year if year is not None else '—')} | "
                f"{_md_cell(_link_cell(doi, url))} | {_md_cell(note)} |"
            )
        lines.append("")

    # ---- 失败清单 -------------------------------------------------------
    failed = [item for item in items if item.error is not None]
    if failed:
        lines.append("## 加载失败文档")
        lines.append("")
        for item in failed:
            lines.append(f"- `{item.doc}`：{item.error}")
        lines.append("")

    lines.extend(
        [
            "---",
            "",
            "说明：本报告只做检索辅助、不修改任何文档；未检索到 ≠ 造假，"
            "所有结论均需人工核实（中文期刊、专著、内部资料在公开索引中覆盖有限）。",
            "",
        ]
    )
    return "\n".join(lines)


def write_refs_report(
    summary: BatchRefSummary,
    docs: Sequence[BatchRefDocResult],
    out_json: Path | str,
    out_md: Path | str,
) -> tuple[Path, Path]:
    """写出 JSON（summary + docs 完整明细）与 Markdown 报告；返回两个路径。"""
    json_path = Path(out_json)
    md_path = Path(out_md)
    for path in (json_path, md_path):
        path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "summary": summary.to_dict(),
        "docs": [item.to_dict() for item in docs],
    }
    json_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    md_path.write_text(render_refs_markdown(summary, docs), encoding="utf-8")
    return json_path, md_path


# ---------------------------------------------------------------------------
# 杂项
# ---------------------------------------------------------------------------
def _error_text(exc: BaseException) -> str:
    text = f"{type(exc).__name__}: {exc}".strip()
    return text[:_MAX_ERROR_CHARS] or type(exc).__name__


def _doc_name(item: BatchRefDocResult) -> str:
    try:
        name = Path(item.doc).name
    except (TypeError, ValueError):  # pragma: no cover - 脏数据兜底
        name = ""
    return name or str(item.doc)


def _status_label(status: str) -> str:
    label = STATUS_LABELS.get(status)
    if label:
        return f"{label}（{status}）"
    return f"{status or '未知状态'}"


def _link_cell(doi: str, url: str) -> str:
    if doi:
        link = f"https://doi.org/{doi}"
        return f"[{doi}]({link})"
    if url:
        return f"[链接]({url})"
    return "—"


def _md_cell(value: Any) -> str:
    """Markdown 表格单元格转义：去换行、转义竖线、空值显示为 —。"""
    text = "—" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    if not text:
        return "—"
    return text.replace("|", "\\|")
