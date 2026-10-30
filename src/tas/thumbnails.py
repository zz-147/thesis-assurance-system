"""批量首页缩略图：docx → 首页 PNG（Word COM 渲染）+ 打包 + HTML 映射。

背景：教师端批量检查后，希望每份文档有一张首页缩略图，嵌进 HTML 班级报告里
（``export_html(..., thumbnails=thumbnails_map(results))``），同时能打包下载。

依赖 ``tas.render``（Windows + Word COM）：
- 默认 lazy import ``render_preview``：只有真正批量渲染时才导入/启动 Word；
- 渲染器可注入（``renderer`` 参数），便于测试与将来换 LibreOffice/云端渲染；
- **单份失败只记 error，不中断整批**（Word 未安装/文件损坏/超时都只影响该份）。

设计：
- 每份文档拿到独立的临时输出目录（避免 ``render_preview`` 的 ``page-01.png`` 互相覆盖），
  成功后把首页 PNG 复制到 ``out_dir/<安全文件名>.png``；
- 文件名做 Windows 安全字符清洗（保留中文），同批次重名追加 ``-2``/``-3``；
- 只处理 ``.docx``；返回顺序与入参一致；
- 纯离线（除本机 Word COM），不联网。
"""

from __future__ import annotations

import re
import shutil
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional, Sequence

__all__ = [
    "ThumbResult",
    "render_thumbnails",
    "thumbnails_map",
    "export_thumbnail_zip",
    "summarize_thumbs",
    "NOTE_FILENAME",
]

#: zip / 目录里的缩略图子目录默认名
DEFAULT_SUBDIR = "thumbnails"
#: zip 内的说明文件名
NOTE_FILENAME = "缩略图说明.txt"
#: 允许出现在文件名里的字符：中文 / 字母 / 数字 / . _ -
_SAFE_STEM_RE = re.compile(r"[^0-9A-Za-z\u4e00-\u9fff._\-]+")
#: Windows 保留设备名（避免生成无法删除/打开的文件）
_WINDOWS_RESERVED = frozenset(
    {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
     *(f"LPT{i}" for i in range(1, 10))}
)


@dataclass
class ThumbResult:
    """一份文档的缩略图渲染结果（失败也返回，不抛异常）。"""

    doc: Path
    png: Optional[Path] = None
    error: Optional[str] = None
    duration_s: float = 0.0

    @property
    def ok(self) -> bool:
        """是否成功生成缩略图。"""
        return self.png is not None and self.error is None

    def to_dict(self) -> dict[str, Any]:
        return {
            "doc": str(self.doc),
            "png": str(self.png) if self.png is not None else None,
            "ok": self.ok,
            "error": self.error,
            "duration_s": self.duration_s,
        }


# ---------------------------------------------------------------------------
# 文件名 / 路径
# ---------------------------------------------------------------------------
def _safe_stem(doc: Path | str) -> str:
    """把文档名转成安全的 PNG 文件名（不含扩展名，保留中文），空则回退 document。"""
    path = Path(doc)
    stem = _SAFE_STEM_RE.sub("_", path.stem or "")
    stem = re.sub(r"_+", "_", stem).strip(" ._-")[:80].strip(" ._-")
    if not stem:
        return "document"
    if stem.upper() in _WINDOWS_RESERVED:
        stem = f"{stem}_doc"
    return stem


def _new_png_path(out_dir: Path, doc: Path, used: set[str]) -> Path:
    """生成批次内唯一的目标 PNG 路径：重名追加 -2 / -3 …"""
    base = _safe_stem(doc)
    candidate = base
    index = 2
    while candidate.casefold() in used:
        candidate = f"{base}-{index}"
        index += 1
    used.add(candidate.casefold())
    return out_dir / f"{candidate}.png"


def _normalize_subdir(subdir: Optional[str]) -> str:
    """把 zip / HTML 用的子目录规范成 "thumbnails" 这种相对路径（无首尾斜杠）。

    ``None`` → 默认 ``thumbnails``；显式 ``""`` → 不放在子目录里。
    """
    if subdir is None:
        return DEFAULT_SUBDIR
    text = str(subdir).replace("\\", "/").strip()
    return text.strip("/")


def _error_text(exc: BaseException) -> str:
    message = str(exc).strip()
    return f"{type(exc).__name__}: {message}" if message else type(exc).__name__


# ---------------------------------------------------------------------------
# 批量渲染
# ---------------------------------------------------------------------------
def _render_one(
    doc: Path,
    out_dir: Path,
    *,
    max_pages: int,
    width: int,
    timeout: float,
    renderer: Optional[Callable[..., Any]],
    used_names: set[str],
) -> Path:
    """渲染一份文档并把首页 PNG 复制到 out_dir；失败抛异常由调用方收敛。"""
    with tempfile.TemporaryDirectory(prefix="tas-thumb-") as tmp:
        work_dir = Path(tmp)
        if renderer is not None:
            produced = renderer(
                doc, work_dir, max_pages=max_pages, width=width
            )
        else:
            # lazy import：不渲染就不依赖 render.py / Word COM
            from .render import RenderError, render_available, render_preview

            if not render_available():
                raise RenderError(
                    "缩略图渲染不可用：需要 Windows + Word + PowerShell（检查报告仍可正常使用）"
                )
            produced = render_preview(
                doc,
                work_dir,
                max_pages=max_pages,
                width=width,
                timeout=timeout,
            )

        pages = [Path(item) for item in (produced or []) if item is not None]
        if not pages:
            raise RuntimeError("渲染器未返回页面图片")
        first = pages[0]
        if not first.is_file():
            raise RuntimeError(f"渲染输出图片不存在：{first}")
        if first.stat().st_size <= 0:
            raise RuntimeError(f"渲染输出图片为空：{first}")

        target = _new_png_path(out_dir, doc, used_names)
        shutil.copyfile(first, target)
    return target


def render_thumbnails(
    paths: Sequence[Path | str],
    out_dir: Path | str,
    *,
    max_pages: int = 1,
    width: int = 900,
    timeout: float = 60.0,
    renderer: Optional[Callable[..., Any]] = None,
) -> list[ThumbResult]:
    """批量渲染文档首页缩略图。

    参数：
        paths: 待渲染文件（``.docx``；其它后缀记为失败并跳过，不调用渲染器）。
        out_dir: PNG 输出目录（不存在会自动创建）。建议由调用方传
            ``报告目录/thumbnails``，配合 :func:`thumbnails_map` 生成 HTML 相对路径。
        max_pages: 每份最多渲染页数（默认 1 = 首页；缩略图只取第一张）。
        width: 渲染目标宽度（px）。
        timeout: 单份渲染超时秒数（仅默认渲染器使用）。
        renderer: 可注入渲染器，签名
            ``(docx: Path, out_dir: Path, *, max_pages: int, width: int) -> list[Path]``；
            默认 ``None`` 时 lazy import :func:`tas.render.render_preview`。
            渲染器每次拿到**独立的临时输出目录**（模块负责把首页复制成安全唯一文件名）。

    返回：与入参顺序一致的 :class:`ThumbResult` 列表；单份失败只记 ``error``，
    绝不抛异常、不中断整批。
    """
    output_dir = Path(out_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[ThumbResult] = []
    used_names: set[str] = set()

    for raw_path in paths or []:
        doc = Path(raw_path)
        started = time.perf_counter()
        suffix = doc.suffix.lower()
        if suffix != ".docx":
            results.append(
                ThumbResult(
                    doc=doc,
                    png=None,
                    error="跳过：仅支持 .docx 文件",
                    duration_s=round(time.perf_counter() - started, 4),
                )
            )
            continue
        if not doc.is_file():
            results.append(
                ThumbResult(
                    doc=doc,
                    png=None,
                    error=f"文件不存在：{doc}",
                    duration_s=round(time.perf_counter() - started, 4),
                )
            )
            continue

        try:
            png = _render_one(
                doc,
                output_dir,
                max_pages=max(1, int(max_pages)),
                width=max(100, int(width)),
                timeout=float(timeout),
                renderer=renderer,
                used_names=used_names,
            )
        except Exception as exc:  # noqa: BLE001 - 单份失败必须不影响整批
            results.append(
                ThumbResult(
                    doc=doc,
                    png=None,
                    error=_error_text(exc),
                    duration_s=round(time.perf_counter() - started, 4),
                )
            )
            continue

        results.append(
            ThumbResult(
                doc=doc,
                png=png,
                error=None,
                duration_s=round(time.perf_counter() - started, 4),
            )
        )
    return results


# ---------------------------------------------------------------------------
# HTML 映射 / zip / 汇总
# ---------------------------------------------------------------------------
def thumbnails_map(
    results: Sequence[ThumbResult], *, subdir: str = DEFAULT_SUBDIR
) -> dict[str, str]:
    """返回 ``{文档文件名: "thumbnails/xxx.png"}``（只含成功项，供 HTML 嵌入）。

    键为文档 basename（``.docx`` 带扩展名），值使用 ``/`` 相对路径；
    ``export_html(..., thumbnails=thumbnails_map(...))`` 直接消费。
    """
    prefix = _normalize_subdir(subdir)
    mapping: dict[str, str] = {}
    for result in results:
        if not result.ok or result.png is None:
            continue
        relative = Path(result.png).name
        mapping[Path(result.doc).name] = f"{prefix}/{relative}" if prefix else relative
    return mapping


def _note_text(results: Sequence[ThumbResult], subdir: str) -> str:
    summary = summarize_thumbs(results)
    lines = [
        "tas 批量首页缩略图说明",
        "=" * 32,
        (
            f"共 {summary['total']} 份：成功 {summary['ok']} ｜ 失败 {summary['failed']}"
            f" ｜ 渲染总耗时 {summary['duration_s']:.2f}s"
        ),
        "",
    ]
    for index, result in enumerate(results, 1):
        name = Path(result.doc).name
        if result.ok and result.png is not None:
            relative = f"{subdir}/{Path(result.png).name}" if subdir else Path(result.png).name
            lines.append(f"{index}. {name}：成功 → {relative}")
        else:
            reason = result.error or "未知原因"
            lines.append(f"{index}. {name}：失败（{reason}）")
    return "\n".join(lines) + "\n"


def export_thumbnail_zip(
    results: Sequence[ThumbResult],
    zip_path: Path | str,
    *,
    subdir: str = DEFAULT_SUBDIR,
) -> Path:
    """把成功的缩略图 PNG + ``缩略图说明.txt`` 打包成 zip，返回 zip 路径。

    - zip 内缩略图放在 ``subdir``（默认 ``thumbnails/``）下；
    - 说明文件逐份记录：文件名、是否成功、失败原因；
    - 失败的文档不产生 PNG，但保留在说明里（老师能知道哪几份需要手动看）。
    """
    target = Path(zip_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    folder = _normalize_subdir(subdir)

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        used_arc: set[str] = set()
        for result in results:
            if not result.ok or result.png is None:
                continue
            source = Path(result.png)
            if not source.is_file():
                continue  # 说明文件里会标注失败/缺失
            file_name = source.name
            base = Path(file_name).stem
            candidate = file_name
            index = 2
            while candidate.casefold() in used_arc:
                candidate = f"{base}-{index}{source.suffix}"
                index += 1
            used_arc.add(candidate.casefold())
            archive_name = f"{folder}/{candidate}" if folder else candidate
            archive.write(source, archive_name)

        note = _note_text(results, folder)
        archive.writestr(NOTE_FILENAME, note.encode("utf-8-sig"))
    return target


def summarize_thumbs(results: Sequence[ThumbResult]) -> dict[str, Any]:
    """缩略图任务汇总：``{total, ok, failed, duration_s}``。"""
    items = list(results)
    ok_count = sum(1 for item in items if item.ok)
    return {
        "total": len(items),
        "ok": ok_count,
        "failed": len(items) - ok_count,
        "duration_s": round(sum(float(item.duration_s or 0.0) for item in items), 4),
    }
