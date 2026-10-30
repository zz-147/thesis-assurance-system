"""批量安全修复模块（batch_fix.py）回归测试。

覆盖需求：老师收齐一批 docx → 一次生成全部 ``*_fixed.docx`` + 修复报告 + 汇总 zip。

测试点：
1) 合规文档 + 变异文档批量修复：变异文档修复后 after.error == 0、输出文件存在、
   原文件字节不变；合规文档也生成副本且无崩溃；
2) 单份失败（损坏文件 / 不存在的文件）只记 error，不中断整批；
3) 输出命名 ``<stem>_fixed.docx``，同名冲突追加 -2（批内冲突 / 输出目录已有文件均不覆盖）；
4) dry-run 不创建目录、不写文件，但仍给出将修复的动作与复检计数；
5) zip 包含修复后的 docx + 修复报告.txt + 修复汇总.txt，报告含“修复前/修复后/文件名”；
6) summarize_fix 计数正确、可 JSON 序列化；
7) fix_directory 只收顶层、跳过 ~$/.~lock. 临时文件、按名字排序。

夹具复用 tests/test_check.py::build_good_docx 与 tests/test_web.py::_mutate_font，
不联网。
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from test_check import build_good_docx  # noqa: E402
from test_web import _mutate_font  # noqa: E402

from tas.batch_fix import (  # noqa: E402
    REPORT_NAME,
    SUMMARY_NAME,
    fix_directory,
    fix_files,
    summarize_fix,
    write_fix_zip,
)
from tas.config import RuleSet  # noqa: E402

RULES = RuleSet.load()
ZERO = {"error": 0, "warning": 0, "info": 0}


class BatchFixTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)
        self.src_dir = self.tmp / "src"
        self.src_dir.mkdir(parents=True, exist_ok=True)
        self.out_dir = self.tmp / "out"

    # -- 夹具 -------------------------------------------------------------
    def _good(self, name: str = "good.docx") -> Path:
        return build_good_docx(self.src_dir / name)

    def _mutated(self, name: str = "bad.docx") -> Path:
        return _mutate_font(build_good_docx(self.src_dir / name))

    def _batch(self, paths, **kwargs):
        return fix_files(paths, self.out_dir, ruleset=RULES, **kwargs)

    # -- 1) 批量修复主流程 -------------------------------------------------
    def test_batch_fix_mutated_and_good(self) -> None:
        good = self._good()
        bad = self._mutated()
        good_bytes, bad_bytes = good.read_bytes(), bad.read_bytes()

        results = self._batch([good, bad])

        self.assertEqual([result.src.name for result in results], ["good.docx", "bad.docx"])
        good_result, bad_result = results
        self.assertIsNone(good_result.error, good_result.error)
        self.assertIsNone(bad_result.error, bad_result.error)

        # 合规文档：零动作，也生成一份副本
        self.assertEqual(good_result.before, ZERO)
        self.assertEqual(good_result.actions, 0)
        self.assertTrue(good_result.out is not None and good_result.out.is_file())

        # 变异文档：修复前有 error，修复后复检 error 清零
        self.assertEqual(bad_result.before["error"], 1)
        self.assertIsNotNone(bad_result.after)
        self.assertEqual(bad_result.after["error"], 0)
        self.assertEqual(bad_result.recheck_error, 0)
        self.assertGreaterEqual(bad_result.actions, 1)
        self.assertEqual(bad_result.status(), "已修复")

        # 输出命名 + 输出目录固定为 out_dir
        self.assertEqual(bad_result.out.name, "bad_fixed.docx")
        self.assertEqual(bad_result.out.parent, self.out_dir)
        self.assertNotEqual(bad_result.out, bad)

        # 原文件绝不能变（字节级）
        self.assertEqual(good.read_bytes(), good_bytes)
        self.assertEqual(bad.read_bytes(), bad_bytes)

    # -- 2) 命名冲突：批内重名 / 输出目录已有同名文件 ----------------------
    def test_conflict_names_in_same_batch(self) -> None:
        dir_a = self.tmp / "a"
        dir_b = self.tmp / "b"
        dir_a.mkdir()
        dir_b.mkdir()
        src_a = build_good_docx(dir_a / "论文.docx")
        src_b = build_good_docx(dir_b / "论文.docx")
        _mutate_font(src_b)

        results = fix_files([src_a, src_b], self.out_dir, ruleset=RULES)

        self.assertEqual(
            [result.out.name for result in results],
            ["论文_fixed.docx", "论文_fixed-2.docx"],
        )
        for result in results:
            self.assertIsNone(result.error)
            self.assertTrue(result.out.is_file())

    def test_conflict_with_existing_output_file_is_not_overwritten(self) -> None:
        self.out_dir.mkdir()
        existing = self.out_dir / "论文_fixed.docx"
        existing.write_bytes(b"old placeholder")
        src = build_good_docx(self.src_dir / "论文.docx")

        results = self._batch([src])

        self.assertEqual(results[0].out.name, "论文_fixed-2.docx")
        self.assertEqual(existing.read_bytes(), b"old placeholder")
        self.assertTrue(results[0].out.is_file())

    # -- 3) 坏文件不中断整批 -----------------------------------------------
    def test_broken_file_records_error_and_continues(self) -> None:
        good = self._good()
        broken = self.src_dir / "broken.docx"
        broken.write_text("这不是一个 docx，只是改了扩展名的文本", encoding="utf-8")
        missing = self.src_dir / "missing.docx"

        results = self._batch([broken, good, missing])

        self.assertEqual(len(results), 3)
        broken_result, good_result, missing_result = results

        self.assertIsNotNone(broken_result.error)
        self.assertIn("检查失败", broken_result.error)
        self.assertIsNone(broken_result.out)
        self.assertEqual(broken_result.status(), "失败")
        json.dumps(broken_result.to_dict(), ensure_ascii=False)  # 失败结果也要可序列化

        self.assertIsNone(good_result.error)
        self.assertTrue(good_result.out.is_file())

        self.assertIsNotNone(missing_result.error)
        self.assertIsNone(missing_result.out)

    # -- 4) dry-run -------------------------------------------------------
    def test_dry_run_writes_nothing(self) -> None:
        bad = self._mutated()
        bad_bytes = bad.read_bytes()

        results = self._batch([bad], dry_run=True)

        result = results[0]
        self.assertIsNone(result.error)
        self.assertTrue(result.dry_run)
        self.assertIsNone(result.out)
        self.assertEqual(result.status(), "dry-run")
        self.assertIsNotNone(result.after)
        self.assertEqual(result.after["error"], 0)
        self.assertGreaterEqual(result.actions, 1)

        self.assertFalse(self.out_dir.exists(), "dry-run 不能创建输出目录")
        self.assertEqual(bad.read_bytes(), bad_bytes)

    # -- 5) zip：修复文件 + 两份报告 ---------------------------------------
    def test_write_fix_zip_contains_docs_and_reports(self) -> None:
        good = self._good()
        bad = self._mutated()
        results = self._batch([good, bad])
        zip_path = write_fix_zip(results, self.tmp / "fix.zip")

        self.assertEqual(zip_path, self.tmp / "fix.zip")
        self.assertTrue(zip_path.is_file())
        with zipfile.ZipFile(zip_path) as archive:
            names = archive.namelist()
            self.assertIn("good_fixed.docx", names)
            self.assertIn("bad_fixed.docx", names)
            self.assertIn(REPORT_NAME, names)
            self.assertIn(SUMMARY_NAME, names)
            # zip 内的 docx 与磁盘输出文件逐字节一致
            self.assertEqual(
                archive.read("bad_fixed.docx"), results[1].out.read_bytes()
            )
            report = archive.read(REPORT_NAME).decode("utf-8")
            summary_text = archive.read(SUMMARY_NAME).decode("utf-8")

        for token in (
            "修复前",
            "修复后",
            "good.docx",
            "bad.docx",
            "原文件",
            "输出文件",
            "剩余问题",
            "修复动作",
        ):
            self.assertIn(token, report)
        for token in ("文件总数", "修复动作合计", "改善文档", "逐份清单"):
            self.assertIn(token, summary_text)

    def test_zip_dry_run_has_only_reports(self) -> None:
        results = self._batch([self._mutated()], dry_run=True)
        zip_path = write_fix_zip(results, self.tmp / "dry.zip")

        with zipfile.ZipFile(zip_path) as archive:
            names = set(archive.namelist())
            report = archive.read(REPORT_NAME).decode("utf-8")
        self.assertEqual(names, {REPORT_NAME, SUMMARY_NAME})
        self.assertIn("dry-run", report)

    # -- 6) 汇总 ----------------------------------------------------------
    def test_summarize_fix_counts(self) -> None:
        good = self._good()
        bad = self._mutated()
        broken = self.src_dir / "broken.docx"
        broken.write_text("x", encoding="utf-8")
        results = self._batch([good, bad, broken])

        summary = summarize_fix(results)

        self.assertEqual(summary["total"], 3)
        self.assertEqual(summary["fixed"], 2)
        self.assertEqual(summary["failed"], 1)
        self.assertEqual(summary["actions"], sum(r.actions for r in results))
        self.assertEqual(summary["skipped"], sum(r.skipped for r in results))
        self.assertEqual(
            summary["before_counts"],
            {
                key: sum(r.before.get(key, 0) for r in results) for key in ZERO
            },
        )
        self.assertEqual(
            summary["after_counts"],
            {
                key: sum((r.after or {}).get(key, 0) for r in results) for key in ZERO
            },
        )
        self.assertEqual(summary["after_counts"]["error"], 0)
        self.assertGreaterEqual(summary["before_counts"]["error"], 1)
        self.assertEqual(summary["improved"], 1)
        self.assertGreaterEqual(summary["resolved_errors"], 1)
        self.assertFalse(summary["dry_run"])
        json.dumps(summary, ensure_ascii=False)  # 可 JSON 序列化

    def test_summarize_fix_dry_run_flag(self) -> None:
        results = self._batch([self._mutated()], dry_run=True)
        summary = summarize_fix(results)
        self.assertTrue(summary["dry_run"])
        self.assertEqual(summary["total"], 1)
        self.assertEqual(summary["failed"], 0)
        self.assertGreaterEqual(summary["actions"], 1)

    # -- 7) 目录批量 ------------------------------------------------------
    def test_fix_directory_skips_temp_files_and_sorts(self) -> None:
        build_good_docx(self.src_dir / "b_good.docx")
        _mutate_font(build_good_docx(self.src_dir / "a_bad.docx"))
        # Word 临时文件 / 锁文件 / 子目录：都不应被处理
        (self.src_dir / "~$a_bad.docx").write_bytes(b"tmp")
        (self.src_dir / ".~lock.a_bad.docx").write_bytes(b"tmp")
        nested = self.src_dir / "sub"
        nested.mkdir()
        build_good_docx(nested / "nested.docx")

        results = fix_directory(self.src_dir, self.out_dir, ruleset=RULES)

        self.assertEqual(
            [result.src.name for result in results], ["a_bad.docx", "b_good.docx"]
        )
        self.assertTrue(all(result.error is None for result in results))
        self.assertTrue((self.out_dir / "a_bad_fixed.docx").is_file())
        self.assertTrue((self.out_dir / "b_good_fixed.docx").is_file())
        self.assertFalse((self.out_dir / "nested_fixed.docx").exists())

    def test_fix_directory_missing_dir_raises(self) -> None:
        with self.assertRaises(FileNotFoundError):
            fix_directory(self.tmp / "no-such-dir", self.out_dir, ruleset=RULES)

    # -- 8) to_dict -------------------------------------------------------
    def test_batch_fix_result_to_dict_is_json(self) -> None:
        results = self._batch([self._mutated()])
        payload = results[0].to_dict()

        text = json.dumps(payload, ensure_ascii=False)
        self.assertTrue(text)
        self.assertEqual(
            set(payload),
            {
                "src",
                "out",
                "dry_run",
                "status",
                "error",
                "before",
                "after",
                "actions",
                "skipped",
                "recheck",
                "recheck_error",
                "recheck_warning",
                "recheck_info",
            },
        )
        self.assertEqual(payload["status"], "已修复")
        self.assertIsNone(payload["error"])
        self.assertEqual(payload["before"]["error"], 1)
        self.assertEqual(payload["after"]["error"], 0)
        self.assertEqual(payload["recheck"]["error"], 0)
        self.assertTrue(payload["out"].endswith("bad_fixed.docx"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
