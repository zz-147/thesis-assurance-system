"""渲染预览测试：环境/CLI 解析为快速测试；真实 Word 渲染用 TAS_RENDER_TEST=1 开启。"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

from tas.cli import build_parser  # noqa: E402
from tas.render import RenderError, render_available, render_preview  # noqa: E402
from test_check import build_good_docx  # noqa: E402


class RenderParserTest(unittest.TestCase):
    def test_preview_subcommand_exists(self):
        parser = build_parser()
        args = parser.parse_args(["preview", "某论文.docx", "--pages", "2", "--out-dir", "eval/preview"])
        self.assertEqual(args.command, "preview")
        self.assertEqual(args.docx, "某论文.docx")
        self.assertEqual(args.pages, 2)
        self.assertEqual(args.out_dir, "eval/preview")


class RenderEnvironmentTest(unittest.TestCase):
    def test_render_available_is_bool(self):
        self.assertIsInstance(render_available(), bool)

    def test_missing_source_raises_quickly(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(RenderError):
                render_preview(Path(tmp) / "missing.docx", Path(tmp) / "preview", max_pages=1)


@unittest.skipUnless(
    os.environ.get("TAS_RENDER_TEST") == "1",
    "真实渲染需要 Word COM（Windows 上设 TAS_RENDER_TEST=1 才跑）",
)
class LiveRenderTest(unittest.TestCase):
    def test_render_preview_live(self):
        with tempfile.TemporaryDirectory() as tmp:
            docx = build_good_docx(Path(tmp) / "good.docx")
            pages = render_preview(docx, Path(tmp) / "preview", max_pages=1)
            self.assertGreaterEqual(len(pages), 1)
            self.assertTrue(pages[0].name.startswith("page-"))
            self.assertGreater(pages[0].stat().st_size, 1000)


if __name__ == "__main__":
    unittest.main(verbosity=2)
