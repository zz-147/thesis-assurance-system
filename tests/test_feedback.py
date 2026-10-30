"""学生反馈包测试：逐份反馈文本 + zip 结构。"""

from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

from tas.batch import DocResult, check_files, summarize  # noqa: E402
from tas.config import RuleSet  # noqa: E402
from tas.feedback import export_feedback_zip, feedback_text  # noqa: E402
from test_check import build_good_docx  # noqa: E402
from test_web import _mutate_font  # noqa: E402


class FeedbackTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.ruleset = RuleSet.load()

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _docx(self, name: str, mutate: bool = False) -> Path:
        path = build_good_docx(Path(self.tmp.name) / name)
        if mutate:
            _mutate_font(path)
        return path

    def test_zip_contains_per_doc_feedback(self):
        good = self._docx("张三_论文.docx")
        bad = self._docx("李四_论文.docx", mutate=True)
        results = check_files([good, bad], ruleset=self.ruleset, profile="sci_eng")
        summary = summarize(results, ruleset_id="school_default")
        zip_path = export_feedback_zip(
            results,
            Path(self.tmp.name) / "feedback.zip",
            summary=summary,
            profile_label="测试规范",
        )
        self.assertTrue(zip_path.is_file())
        with zipfile.ZipFile(zip_path) as archive:
            names = archive.namelist()
            texts = {name: archive.read(name).decode("utf-8") for name in names}

        self.assertIn("00_班级汇总.txt", names)
        good_text = next(text for name, text in texts.items() if "张三" in name)
        bad_text = next(text for name, text in texts.items() if "李四" in name)
        self.assertIn("结论：通过", good_text)
        self.assertIn("结论：未通过", bad_text)
        self.assertIn("BODY-001", bad_text)
        self.assertIn("修改建议", bad_text)
        self.assertIn("实际：", bad_text)
        summary_text = texts["00_班级汇总.txt"]
        self.assertIn("通过率", summary_text)
        self.assertIn("未通过名单", summary_text)
        self.assertIn("李四", summary_text)

    def test_failed_document_feedback(self):
        result = DocResult(path=Path("坏文件.docx"), error="PackageNotFoundError: not a zip")
        text = feedback_text(result)
        self.assertIn("检查失败", text)
        self.assertIn("另存为", text)
        self.assertNotIn("修改建议", text)

    def test_feedback_text_for_compliant_document(self):
        result = check_files(
            [self._docx("合格_论文.docx")], ruleset=self.ruleset, profile="sci_eng"
        )[0]
        text = feedback_text(result)
        self.assertIn("结论：通过", text)
        self.assertIn("未发现格式问题", text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
