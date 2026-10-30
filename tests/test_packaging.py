"""打包/可视化相关测试：project_root 覆盖、web --open-browser、HTML 报告图表。"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

from tas.batch import check_files, summarize  # noqa: E402
from tas.cli import build_parser  # noqa: E402
from tas.config import RuleSet, project_root  # noqa: E402
from tas.export import export_html  # noqa: E402
from test_check import build_good_docx  # noqa: E402
from test_web import _mutate_font  # noqa: E402


class PackagingTest(unittest.TestCase):
    def test_project_root_env_override(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch.dict(os.environ, {"TAS_PROJECT_ROOT": tmp}):
                self.assertEqual(project_root(), Path(tmp).resolve())

    def test_web_open_browser_flag(self):
        args = build_parser().parse_args(["web", "--open-browser", "--port", "8799"])
        self.assertTrue(args.open_browser)
        self.assertEqual(args.port, 8799)


class ExportVisualizationTest(unittest.TestCase):
    def test_html_contains_bars_and_progress(self):
        with tempfile.TemporaryDirectory() as tmp:
            good = build_good_docx(Path(tmp) / "good.docx")
            bad = build_good_docx(Path(tmp) / "bad.docx")
            _mutate_font(bad)
            results = check_files([good, bad], ruleset=RuleSet.load(), profile="sci_eng")
            summary = summarize(results, ruleset_id="school_default")
            html = export_html(summary, results)
        self.assertIn("bar-fill", html)
        self.assertIn("progress-fill", html)
        self.assertIn("共性问题 Top", html)

    def test_html_thumbnails_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            good = build_good_docx(Path(tmp) / "good.docx")
            bad = build_good_docx(Path(tmp) / "bad.docx")
            _mutate_font(bad)
            results = check_files([good, bad], ruleset=RuleSet.load(), profile="sci_eng")
            summary = summarize(results, ruleset_id="school_default")
            html = export_html(
                summary,
                results,
                thumbnails={
                    "good.docx": "thumbnails/good.png",
                    "bad.docx": "thumbnails/bad.png",
                },
            )
            plain = export_html(summary, results)
        self.assertIn("缩略图", html)
        self.assertIn("thumbnails/good.png", html)
        self.assertIn("thumb", html)
        self.assertNotIn("缩略图", plain)


if __name__ == "__main__":
    unittest.main(verbosity=2)
