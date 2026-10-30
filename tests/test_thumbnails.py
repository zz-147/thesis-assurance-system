"""批量首页缩略图模块（tas.thumbnails）测试。

全部离线、不依赖 Word：
1. 注入假渲染器：3 份文档（1 份抛异常）→ 2 成功 / 1 失败、顺序正确、文件名安全唯一；
2. 默认渲染器不可用 / RenderError / 返回空 → 全部失败但不抛异常；
3. thumbnails_map 相对路径、export_thumbnail_zip（PNG + 缩略图说明.txt）、
   summarize_thumbs 计数、非 .docx 与文件缺失的边界；
4. 可选 live：``TAS_RENDER_TEST=1`` 时用真实 Word 渲染 1 页（默认 skip）。
"""

from __future__ import annotations

import base64
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import tas.render as render_module  # noqa: E402
from tas.thumbnails import (  # noqa: E402
    NOTE_FILENAME,
    ThumbResult,
    export_thumbnail_zip,
    render_thumbnails,
    summarize_thumbs,
    thumbnails_map,
)

# 1x1 透明 PNG（合法 PNG 文件字节）
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)
assert PNG_1PX[:8] == b"\x89PNG\r\n\x1a\n"


# ---------------------------------------------------------------------------
# 工具：假渲染器（不依赖 Word）
# ---------------------------------------------------------------------------
class FakeRenderer:
    """可注入渲染器：写伪 PNG；指定文件名抛异常；记录每次调用参数。"""

    def __init__(self, fail_names=(), pages: int = 1, payload: bytes = PNG_1PX):
        self.fail_names = set(fail_names)
        self.pages = pages
        self.payload = payload
        self.calls: list[dict] = []

    def __call__(self, docx, out_dir, *, max_pages, width):
        docx, out_dir = Path(docx), Path(out_dir)
        self.calls.append(
            {"docx": docx, "out_dir": out_dir, "max_pages": max_pages, "width": width}
        )
        if docx.name in self.fail_names:
            raise RuntimeError(f"boom: {docx.name}")
        out_dir.mkdir(parents=True, exist_ok=True)
        written: list[Path] = []
        for index in range(max(1, self.pages)):
            target = out_dir / f"page-{index + 1:02d}.png"
            target.write_bytes(self.payload)
            written.append(target)
        return written


class ThumbTestBase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def make_docx(self, name: str, subdir: str = "in") -> Path:
        path = self.tmp / subdir / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"fake docx content")  # 内容不重要：渲染器是假的
        return path

    def make_docs(self, *names: str) -> list[Path]:
        return [self.make_docx(name) for name in names]

    def assert_png(self, path: Path) -> None:
        self.assertIsInstance(path, Path)
        self.assertTrue(path.is_file(), f"PNG 未生成：{path}")
        self.assertEqual(path.read_bytes(), PNG_1PX)


# ---------------------------------------------------------------------------
# 1) ThumbResult
# ---------------------------------------------------------------------------
class ThumbResultTest(unittest.TestCase):
    def test_ok_and_to_dict(self):
        doc = Path("a.docx")
        png = Path("thumbnails/a.png")
        success = ThumbResult(doc=doc, png=png)
        self.assertTrue(success.ok)
        payload = success.to_dict()
        self.assertEqual(payload["doc"], str(doc))
        self.assertEqual(payload["png"], str(png))
        self.assertTrue(payload["ok"])
        self.assertIsNone(payload["error"])
        self.assertEqual(payload["duration_s"], 0.0)

        failed = ThumbResult(doc=doc, png=None, error="RenderError: x", duration_s=0.5)
        self.assertFalse(failed.ok)
        self.assertEqual(
            failed.to_dict(),
            {
                "doc": "a.docx",
                "png": None,
                "ok": False,
                "error": "RenderError: x",
                "duration_s": 0.5,
            },
        )

    def test_png_with_error_is_not_ok(self):
        self.assertFalse(ThumbResult(doc=Path("a.docx"), png=Path("a.png"), error="x").ok)


# ---------------------------------------------------------------------------
# 2) render_thumbnails（假渲染器）
# ---------------------------------------------------------------------------
class RenderThumbnailsTest(ThumbTestBase):
    def test_fake_renderer_two_success_one_failure(self):
        docs = self.make_docs("a.docx", "b.docx", "c.docx")
        renderer = FakeRenderer(fail_names={"b.docx"})
        out_dir = self.tmp / "report" / "thumbnails"

        results = render_thumbnails(docs, out_dir, renderer=renderer)

        self.assertEqual(len(results), 3)
        self.assertEqual([result.doc.name for result in results], ["a.docx", "b.docx", "c.docx"])
        self.assertEqual([result.ok for result in results], [True, False, True])
        self.assertTrue(out_dir.is_dir(), "输出目录应自动创建")

        self.assert_png(results[0].png)
        self.assertIsNone(results[1].png)
        self.assertIsNone(results[0].error)
        self.assertIn("boom", results[1].error or "")
        self.assertIn("RuntimeError", results[1].error or "")
        self.assert_png(results[2].png)

        # 文件名安全唯一、都在 out_dir 下
        names = [result.png.name for result in results if result.png is not None]
        self.assertEqual(names, ["a.png", "c.png"])
        self.assertEqual(len({name.casefold() for name in names}), len(names))

        # 渲染器每份文档各调用一次，且收到独立的临时目录（不是最终 out_dir）
        self.assertEqual(len(renderer.calls), 3)
        for call in renderer.calls:
            self.assertNotEqual(Path(call["out_dir"]).resolve(), out_dir.resolve())
            self.assertEqual(call["max_pages"], 1)
            self.assertEqual(call["width"], 900)

    def test_custom_max_pages_and_width_forwarded(self):
        docs = self.make_docs("a.docx")
        renderer = FakeRenderer(pages=3)
        results = render_thumbnails(
            docs, self.tmp / "out", max_pages=2, width=600, renderer=renderer
        )
        self.assertTrue(results[0].ok)
        self.assertEqual(renderer.calls[0]["max_pages"], 2)
        self.assertEqual(renderer.calls[0]["width"], 600)
        # 缩略图只取第一页
        self.assertEqual(results[0].png.name, "a.png")
        self.assert_png(results[0].png)
        self.assertEqual(sorted(p.name for p in (self.tmp / "out").glob("*.png")), ["a.png"])

    def test_duplicate_stems_get_unique_safe_names(self):
        first = self.make_docx("张三 论文(1).docx", subdir="甲")
        second = self.make_docx("张三 论文(1).docx", subdir="乙")
        results = render_thumbnails(
            [first, second], self.tmp / "out", renderer=FakeRenderer()
        )
        self.assertTrue(all(result.ok for result in results))
        names = [result.png.name for result in results]
        self.assertEqual(len(set(names)), 2, names)
        self.assertEqual(names[0], "张三_论文_1.png")
        self.assertEqual(names[1], "张三_论文_1-2.png")
        for name in names:  # 安全字符：无空格 / 括号
            self.assertNotIn(" ", name)
            self.assertNotIn("(", name)
            self.assertNotIn(")", name)
        for result in results:
            self.assert_png(result.png)

    def test_non_docx_is_skipped_with_error(self):
        pdf = self.tmp / "in" / "a.pdf"
        pdf.parent.mkdir(parents=True, exist_ok=True)
        pdf.write_bytes(b"%PDF-1.4")
        docx = self.make_docx("b.docx")
        renderer = FakeRenderer()

        results = render_thumbnails([pdf, docx], self.tmp / "out", renderer=renderer)

        self.assertEqual([result.doc.name for result in results], ["a.pdf", "b.docx"])
        self.assertFalse(results[0].ok)
        self.assertIn(".docx", results[0].error or "")
        self.assertTrue(results[1].ok)
        self.assertEqual(len(renderer.calls), 1, "非 docx 不应调用渲染器")
        self.assertEqual(renderer.calls[0]["docx"].name, "b.docx")

    def test_missing_docx_fails_without_renderer_call(self):
        renderer = FakeRenderer()
        results = render_thumbnails(
            [self.tmp / "not-exist.docx"], self.tmp / "out", renderer=renderer
        )
        self.assertFalse(results[0].ok)
        self.assertIn("文件不存在", results[0].error or "")
        self.assertEqual(renderer.calls, [])

    def test_renderer_returning_nothing_is_failure(self):
        results = render_thumbnails(
            self.make_docs("a.docx"),
            self.tmp / "out",
            renderer=lambda docx, out_dir, **kwargs: [],
        )
        self.assertFalse(results[0].ok)
        self.assertIn("未返回", results[0].error or "")

    def test_renderer_returning_missing_file_is_failure(self):
        results = render_thumbnails(
            self.make_docs("a.docx"),
            self.tmp / "out",
            renderer=lambda docx, out_dir, **kwargs: [Path("nope.png")],
        )
        self.assertFalse(results[0].ok)
        self.assertIn("不存在", results[0].error or "")

    def test_empty_paths(self):
        self.assertEqual(render_thumbnails([], self.tmp / "out"), [])

    def test_duration_recorded(self):
        sleep_renderer = FakeRenderer()
        original = sleep_renderer.__call__

        def slow_renderer(docx, out_dir, **kwargs):
            import time

            time.sleep(0.01)
            return original(docx, out_dir, **kwargs)

        results = render_thumbnails(
            self.make_docs("a.docx"), self.tmp / "out", renderer=slow_renderer
        )
        self.assertGreaterEqual(results[0].duration_s, 0.005)
        self.assertEqual(results[0].duration_s, round(results[0].duration_s, 4))


# ---------------------------------------------------------------------------
# 3) 默认渲染器：不可用 / 抛 RenderError → 全部失败但不抛
# ---------------------------------------------------------------------------
class DefaultRendererFailureTest(ThumbTestBase):
    def test_render_unavailable_marks_all_failed(self):
        docs = self.make_docs("a.docx", "b.docx")
        with mock.patch.object(render_module, "render_available", return_value=False):
            with mock.patch.object(
                render_module,
                "render_preview",
                side_effect=render_module.RenderError("不该被调用"),
            ):
                results = render_thumbnails(docs, self.tmp / "out")

        self.assertEqual(len(results), 2)
        self.assertTrue(all(not result.ok for result in results))
        self.assertTrue(all(result.png is None for result in results))
        for result in results:
            self.assertIn("不可用", result.error or "")

    def test_render_preview_raises_render_error_marks_all_failed(self):
        docs = self.make_docs("a.docx", "b.docx")
        with mock.patch.object(render_module, "render_available", return_value=True):
            with mock.patch.object(
                render_module,
                "render_preview",
                side_effect=render_module.RenderError("Word 启动失败"),
            ):
                results = render_thumbnails(docs, self.tmp / "out")

        self.assertEqual(len(results), 2)
        self.assertTrue(all(not result.ok for result in results))
        for result in results:
            self.assertIn("RenderError", result.error or "")
            self.assertIn("Word 启动失败", result.error or "")

    def test_render_preview_returns_empty_marks_failed(self):
        with mock.patch.object(render_module, "render_available", return_value=True):
            with mock.patch.object(render_module, "render_preview", return_value=[]):
                results = render_thumbnails(self.make_docs("a.docx"), self.tmp / "out")
        self.assertFalse(results[0].ok)
        self.assertIn("未返回", results[0].error or "")

    def test_default_renderer_receives_timeout_and_params(self):
        captured: dict = {}

        def fake_preview(docx, out_dir, *, max_pages, width, timeout):
            captured.update(
                {
                    "docx": Path(docx),
                    "out_dir": Path(out_dir),
                    "max_pages": max_pages,
                    "width": width,
                    "timeout": timeout,
                }
            )
            target = Path(out_dir) / "page-01.png"
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(PNG_1PX)
            return [target]

        with mock.patch.object(render_module, "render_available", return_value=True):
            with mock.patch.object(render_module, "render_preview", side_effect=fake_preview):
                results = render_thumbnails(
                    self.make_docs("a.docx"),
                    self.tmp / "out",
                    max_pages=1,
                    width=777,
                    timeout=12.5,
                )

        self.assertTrue(results[0].ok)
        self.assertEqual(captured["max_pages"], 1)
        self.assertEqual(captured["width"], 777)
        self.assertEqual(captured["timeout"], 12.5)
        self.assert_png(results[0].png)


# ---------------------------------------------------------------------------
# 4) thumbnails_map
# ---------------------------------------------------------------------------
class ThumbnailsMapTest(ThumbTestBase):
    def _results(self):
        docs = self.make_docs("张三.docx", "李四.docx", "失败.docx")
        renderer = FakeRenderer(fail_names={"失败.docx"})
        return render_thumbnails(docs, self.tmp / "report" / "thumbnails", renderer=renderer)

    def test_default_subdir_and_only_success(self):
        results = self._results()
        mapping = thumbnails_map(results)
        self.assertEqual(
            mapping,
            {"张三.docx": "thumbnails/张三.png", "李四.docx": "thumbnails/李四.png"},
        )
        self.assertNotIn("失败.docx", mapping)

    def test_custom_subdir_and_forward_slashes(self):
        results = self._results()
        mapping = thumbnails_map(results, subdir="preview\\thumbs")
        self.assertEqual(mapping["张三.docx"], "preview/thumbs/张三.png")
        self.assertNotIn("\\", mapping["张三.docx"])

    def test_empty_subdir(self):
        results = self._results()
        self.assertEqual(thumbnails_map(results, subdir="")["张三.docx"], "张三.png")

    def test_manual_results(self):
        results = [
            ThumbResult(doc=Path("ok.docx"), png=Path("x/ok.png")),
            ThumbResult(doc=Path("bad.docx"), png=None, error="x"),
            ThumbResult(doc=Path("weird.docx"), png=Path("x/weird.png"), error="x"),
        ]
        self.assertEqual(thumbnails_map(results), {"ok.docx": "thumbnails/ok.png"})


# ---------------------------------------------------------------------------
# 5) export_thumbnail_zip
# ---------------------------------------------------------------------------
class ExportThumbnailZipTest(ThumbTestBase):
    def _results(self):
        docs = self.make_docs("a.docx", "b.docx", "c.docx")
        renderer = FakeRenderer(fail_names={"b.docx"})
        return render_thumbnails(docs, self.tmp / "report" / "thumbnails", renderer=renderer)

    def test_zip_contains_png_and_note(self):
        results = self._results()
        zip_path = export_thumbnail_zip(results, self.tmp / "out" / "thumbs.zip")

        self.assertTrue(zip_path.is_file())
        self.assertEqual(zip_path.parent.name, "out", "输出目录应自动创建")
        with zipfile.ZipFile(zip_path) as archive:
            names = archive.namelist()
            png_names = [name for name in names if name.endswith(".png")]
            self.assertEqual(len(png_names), 2)
            self.assertTrue(all(name.startswith("thumbnails/") for name in png_names))
            self.assertIn(NOTE_FILENAME, names)

            data = archive.read(png_names[0])
            self.assertEqual(data, PNG_1PX)

            note = archive.read(NOTE_FILENAME).decode("utf-8-sig")
            self.assertIn("共 3 份：成功 2 ｜ 失败 1", note)
            self.assertIn("a.docx：成功", note)
            self.assertIn("b.docx：失败", note)
            self.assertIn("boom", note)

    def test_zip_custom_subdir(self):
        results = self._results()
        zip_path = export_thumbnail_zip(
            results, self.tmp / "thumbs.zip", subdir="pic/pages"
        )
        with zipfile.ZipFile(zip_path) as archive:
            png_names = [name for name in archive.namelist() if name.endswith(".png")]
            self.assertTrue(all(name.startswith("pic/pages/") for name in png_names))

    def test_zip_with_only_failures(self):
        results = render_thumbnails(
            self.make_docs("a.docx"), self.tmp / "out", renderer=FakeRenderer({"a.docx"})
        )
        zip_path = export_thumbnail_zip(results, self.tmp / "empty.zip")
        with zipfile.ZipFile(zip_path) as archive:
            self.assertEqual(
                archive.namelist(), ["缩略图说明.txt"]
            )
            note = archive.read(NOTE_FILENAME).decode("utf-8-sig")
            self.assertIn("失败 1", note)

    def test_zip_empty_results(self):
        zip_path = export_thumbnail_zip([], self.tmp / "empty.zip")
        with zipfile.ZipFile(zip_path) as archive:
            self.assertEqual(archive.namelist(), ["缩略图说明.txt"])

    def test_zip_skips_png_deleted_before_packaging(self):
        results = self._results()
        for result in results:
            if result.ok and result.png is not None:
                result.png.unlink()
        zip_path = export_thumbnail_zip(results, self.tmp / "missing.zip")
        with zipfile.ZipFile(zip_path) as archive:
            self.assertEqual(archive.namelist(), ["缩略图说明.txt"])


# ---------------------------------------------------------------------------
# 6) summarize_thumbs
# ---------------------------------------------------------------------------
class SummarizeThumbsTest(unittest.TestCase):
    def test_summary_counts(self):
        results = [
            ThumbResult(doc=Path("a.docx"), png=Path("a.png"), duration_s=0.1),
            ThumbResult(doc=Path("b.docx"), png=None, error="x", duration_s=0.2),
            ThumbResult(doc=Path("c.docx"), png=Path("c.png"), duration_s=0.3),
        ]
        summary = summarize_thumbs(results)
        self.assertEqual(summary["total"], 3)
        self.assertEqual(summary["ok"], 2)
        self.assertEqual(summary["failed"], 1)
        self.assertAlmostEqual(summary["duration_s"], 0.6)

    def test_summary_empty(self):
        self.assertEqual(
            summarize_thumbs([]),
            {"total": 0, "ok": 0, "failed": 0, "duration_s": 0.0},
        )


# ---------------------------------------------------------------------------
# 7) 可选 live：真实 Word COM 渲染（默认 skip）
# ---------------------------------------------------------------------------
_LIVE = os.environ.get("TAS_RENDER_TEST") == "1"


@unittest.skipUnless(
    _LIVE and render_module.render_available(),
    "需要 TAS_RENDER_TEST=1 + Windows + Word（默认 skip）",
)
class LiveRenderTest(ThumbTestBase):
    def test_real_render_first_page(self):
        from test_check import build_good_docx  # 复用 tests 里的合规 fixture

        (self.tmp / "in").mkdir(parents=True, exist_ok=True)
        doc = build_good_docx(self.tmp / "in" / "live.docx")
        results = render_thumbnails(
            [doc], self.tmp / "report" / "thumbnails", max_pages=1, width=900, timeout=180
        )
        self.assertEqual(len(results), 1)
        self.assertTrue(results[0].ok, results[0].error)
        self.assertIsNotNone(results[0].png)
        self.assertGreater(results[0].png.stat().st_size, 1000)
        mapping = thumbnails_map(results)
        self.assertIn("live.docx", mapping)


if __name__ == "__main__":
    unittest.main(verbosity=2)
