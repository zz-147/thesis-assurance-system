"""教师端批量检查 + 班级汇总 + 导出（batch.py / export.py）回归测试。

覆盖需求 FR-06：
1) 批量检查：文件列表 / 目录导入（跳过 ~$ 临时文件与子目录）、单份失败不中断整批；
2) 班级汇总：总数/成功/失败、pass_rate 定义、级别计数、共性问题 Top N、逐份清单；
3) 导出：xlsx 三个 Sheet 可被 openpyxl 重新打开校验、csv 逐份+明细合并、html 字符串；
4) run_batch 便捷入口。

夹具复用 tests/test_check.py 的 build_good_docx（合规文档），用 python-docx 注入
已知格式错误（正文字号改为 10.5pt → BODY-001），并用损坏文件模拟单份失败。
"""

from __future__ import annotations

import csv
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from test_check import build_good_docx  # noqa: E402

from tas.batch import (  # noqa: E402
    BatchSummary,
    DocResult,
    check_directory,
    check_files,
    run_batch,
    summarize,
)
from tas.config import RuleSet  # noqa: E402
from tas.export import (  # noqa: E402
    CSV_HEADERS,
    ISSUE_HEADERS,
    PER_DOC_HEADERS,
    export_csv,
    export_html,
    export_xlsx,
)
from tas.report import Issue  # noqa: E402

RULES = RuleSet.load()


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------
def make_issue(
    rule_id: str = "BODY-001",
    severity: str = "error",
    locator: str = "body/para/3",
) -> Issue:
    return Issue(
        rule_id=rule_id,
        title="正文字体字号" if rule_id == "BODY-001" else rule_id,
        severity=severity,
        message="正文的字体/字号不符合当前规范",
        locator=locator,
        actual="字号=10.5pt",
        expected={"size_pt": 12},
    )


def make_result(path: str, issues=None, error=None, duration=0.1) -> DocResult:
    return DocResult(
        path=Path(path),
        issues=list(issues or []),
        stats={"paragraphs": 10},
        error=error,
        duration_s=duration,
    )


def mutate_body_size(path: Path) -> Path:
    """把合规 fixture 的正文首段字号改成 10.5pt（应触发 BODY-001）。"""
    document = docx.Document(str(path))
    for paragraph in document.paragraphs:
        if paragraph.text.startswith("文档格式统一是高校教学管理"):
            paragraph.runs[0]._r.rPr.find(qn("w:sz")).set(qn("w:val"), "21")
            break
    document.save(str(path))
    return path


class BatchTestBase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def good(self, name: str = "good.docx") -> Path:
        return build_good_docx(self.tmp / name)

    def bad(self, name: str = "bad.docx") -> Path:
        return mutate_body_size(self.good(name))


# ---------------------------------------------------------------------------
# 1) 逐份检查 / check_files
# ---------------------------------------------------------------------------
class CheckFilesTest(BatchTestBase):
    def test_check_files_success(self):
        path = self.good()
        results = check_files([path], ruleset=RULES, profile="sci_eng", ruleset_id="school_default")
        self.assertEqual(len(results), 1)
        result = results[0]
        self.assertEqual(result.path, path)
        self.assertIsNone(result.error)
        self.assertIsInstance(result.issues, list)
        self.assertTrue(result.ok)
        self.assertGreaterEqual(result.duration_s, 0.0)
        self.assertGreater(result.stats.get("paragraphs", 0), 0)

    def test_check_files_detects_injected_error(self):
        path = self.bad()
        results = check_files([path], ruleset=RULES)
        self.assertIsNone(results[0].error)
        self.assertIn("BODY-001", results[0].rule_ids())
        self.assertTrue(results[0].has_errors())
        self.assertFalse(results[0].passed())

    def test_single_failure_does_not_interrupt_batch(self):
        good = self.good()
        broken = self.tmp / "broken.docx"
        broken.write_bytes(b"this is not a docx package")
        missing = self.tmp / "missing.docx"
        results = check_files(
            [good, broken, missing], ruleset=RULES, ruleset_id="school_default"
        )
        self.assertEqual(len(results), 3)
        self.assertIsNone(results[0].error)
        self.assertIsNotNone(results[1].error)
        self.assertIn("PackageNotFoundError", results[1].error)
        self.assertIsNotNone(results[2].error)
        self.assertIn("FileNotFoundError", results[2].error)
        # 失败文档不计通过、不产生 issues
        self.assertFalse(results[1].passed())
        self.assertEqual(results[1].issues, [])
        self.assertEqual(results[2].counts(), {"error": 0, "warning": 0, "info": 0})

    def test_check_files_preserves_order_and_accepts_str(self):
        a = self.good("a.docx")
        b = self.good("b.docx")
        results = check_files([str(b), str(a)], ruleset=RULES)
        self.assertEqual([item.path.name for item in results], ["b.docx", "a.docx"])
        self.assertTrue(all(isinstance(item.path, Path) for item in results))

    def test_doc_result_status_and_to_dict(self):
        passed = make_result("a.docx", [make_issue(severity="warning")])
        self.assertTrue(passed.passed())
        self.assertEqual(passed.status(), "通过")

        failed = make_result("b.docx", [make_issue(severity="error")])
        self.assertFalse(failed.passed())
        self.assertEqual(failed.status(), "未通过")

        broken = make_result("c.docx", error="FileNotFoundError: x")
        self.assertFalse(broken.passed())
        self.assertEqual(broken.status(), "失败")
        payload = broken.to_dict()
        self.assertEqual(payload["error"], "FileNotFoundError: x")
        self.assertEqual(payload["counts"]["error"], 0)


# ---------------------------------------------------------------------------
# 2) 目录导入
# ---------------------------------------------------------------------------
class CheckDirectoryTest(BatchTestBase):
    def test_skips_temp_files_and_subdirectories(self):
        self.good("a.docx")
        self.good("b.docx")
        (self.tmp / "notes.txt").write_text("not a docx", encoding="utf-8")
        (self.tmp / "~$a.docx").write_bytes(b"word lock file")
        nested = self.tmp / "sub"
        nested.mkdir()
        self.good("sub/nested.docx")

        results = check_directory(self.tmp, ruleset=RULES)
        names = [item.path.name for item in results]
        self.assertEqual(names, ["a.docx", "b.docx"], "应跳过 ~$ 临时文件、子目录与非 docx")

    def test_patterns_are_configurable(self):
        self.good("a.docx")
        (self.tmp / "b.txt").write_text("plain", encoding="utf-8")
        results = check_directory(self.tmp, patterns=("*.txt",), ruleset=RULES)
        self.assertEqual([item.path.name for item in results], ["b.txt"])
        self.assertIsNotNone(results[0].error, "非 docx 文件应记 error 而不是抛异常")

    def test_missing_directory_raises(self):
        with self.assertRaises(FileNotFoundError):
            check_directory(self.tmp / "nope", ruleset=RULES)

    def test_empty_directory_returns_empty(self):
        self.assertEqual(check_directory(self.tmp, ruleset=RULES), [])


# ---------------------------------------------------------------------------
# 3) 班级汇总
# ---------------------------------------------------------------------------
class SummarizeTest(unittest.TestCase):
    def test_empty_results(self):
        summary = summarize([], ruleset_id="school_default")
        self.assertIsInstance(summary, BatchSummary)
        self.assertEqual((summary.total, summary.checked, summary.failed), (0, 0, 0))
        self.assertEqual(summary.pass_rate, 0.0)
        self.assertEqual(summary.counts, {"error": 0, "warning": 0, "info": 0})
        self.assertEqual(summary.top_rules, [])
        self.assertEqual(summary.per_doc, [])
        self.assertEqual(summary.ruleset_id, "school_default")
        self.assertTrue(summary.generated_at)

    def test_pass_rate_only_counts_checked_documents(self):
        results = [
            make_result("pass.docx", [make_issue(severity="warning")]),
            make_result("fail.docx", [make_issue(severity="error")]),
            make_result("broken.docx", error="PackageNotFoundError: bad"),
        ]
        summary = summarize(results, ruleset_id="school_default")
        self.assertEqual(summary.total, 3)
        self.assertEqual(summary.checked, 2)
        self.assertEqual(summary.failed, 1)
        self.assertAlmostEqual(summary.pass_rate, 0.5)  # 1 通过 / 2 检查成功
        self.assertEqual(summary.pass_count, 1)
        self.assertEqual(summary.counts, {"error": 1, "warning": 1, "info": 0})
        self.assertEqual(summary.top_rules, [("BODY-001", 2)])

    def test_all_checked_pass(self):
        results = [make_result("a.docx"), make_result("b.docx", [make_issue(severity="info")])]
        summary = summarize(results)
        self.assertEqual(summary.pass_rate, 1.0)
        self.assertEqual(summary.failed, 0)

    def test_top_rules_sorted_and_limited_to_10(self):
        results = []
        # 12 条规则，命中数 12,11,...,1（BODY-001 命中最多，R-012 最少）
        for index in range(1, 13):
            rule_id = f"RULE-{index:03d}"
            issues = [make_issue(rule_id=rule_id, severity="warning") for _ in range(13 - index)]
            results.append(make_result(f"{rule_id}.docx", issues))
        summary = summarize(results)
        self.assertEqual(len(summary.top_rules), 10)
        self.assertEqual(summary.top_rules[0], ("RULE-001", 12))
        counts = [count for _rule, count in summary.top_rules]
        self.assertEqual(counts, sorted(counts, reverse=True))
        self.assertNotIn("RULE-011", [rule for rule, _count in summary.top_rules])

    def test_per_doc_fields(self):
        results = [
            make_result("pass.docx", [make_issue(severity="warning")], duration=0.25),
            make_result("broken.docx", error="FileNotFoundError: x"),
        ]
        summary = summarize(results, ruleset_id="profile_x")
        first = summary.per_doc[0]
        for key in ("path", "counts", "error", "warning", "info", "status", "duration_s", "load_error"):
            self.assertIn(key, first)
        self.assertEqual(first["path"], "pass.docx")
        self.assertEqual(first["counts"], {"error": 0, "warning": 1, "info": 0})
        self.assertEqual(first["error"], 0)
        self.assertEqual(first["warning"], 1)
        self.assertTrue(first["passed"])
        self.assertEqual(first["status"], "通过")
        self.assertAlmostEqual(first["duration_s"], 0.25)
        self.assertIsNone(first["load_error"])

        second = summary.per_doc[1]
        self.assertFalse(second["passed"])
        self.assertEqual(second["status"], "失败")
        self.assertEqual(second["load_error"], "FileNotFoundError: x")

    def test_summary_to_dict_is_json_serializable(self):
        import json

        summary = summarize([make_result("a.docx", [make_issue()])], ruleset_id="p")
        payload = summary.to_dict()
        text = json.dumps(payload, ensure_ascii=False)
        self.assertIn("pass_rate", payload)
        self.assertIsInstance(payload["top_rules"], list)
        self.assertTrue(text)


# ---------------------------------------------------------------------------
# 4) run_batch 便捷入口
# ---------------------------------------------------------------------------
class RunBatchTest(BatchTestBase):
    def test_run_batch_with_files(self):
        good = self.good("good.docx")
        bad = self.bad("bad.docx")
        results, summary = run_batch(
            paths=[good, bad], ruleset=RULES, ruleset_id="school_default"
        )
        self.assertEqual(len(results), 2)
        self.assertEqual(summary.total, 2)
        self.assertEqual(summary.ruleset_id, "school_default")
        self.assertIn("BODY-001", [rule for rule, _count in summary.top_rules])

    def test_run_batch_with_directory(self):
        self.good("a.docx")
        results, summary = run_batch(directory=self.tmp, ruleset=RULES)
        self.assertEqual(len(results), 1)
        self.assertEqual(summary.total, 1)

    def test_run_batch_requires_source(self):
        with self.assertRaises(ValueError):
            run_batch()


# ---------------------------------------------------------------------------
# 5) 导出
# ---------------------------------------------------------------------------
class ExportTest(BatchTestBase):
    def _fixture(self):
        good = self.good("good.docx")
        bad = self.bad("bad.docx")
        broken = self.tmp / "broken.docx"
        broken.write_bytes(b"not a docx")
        results = check_files([good, bad, broken], ruleset=RULES)
        summary = summarize(results, ruleset_id="school_default")
        return results, summary

    def test_export_xlsx_structure_and_reopen(self):
        from openpyxl import load_workbook

        results, summary = self._fixture()
        out = export_xlsx(results, summary, self.tmp / "nested" / "batch.xlsx")
        self.assertTrue(out.is_file())
        self.assertEqual(out.parent.name, "nested", "输出目录不存在时应自动创建")

        workbook = load_workbook(str(out))  # 能被 openpyxl 重新打开 = 格式有效
        self.assertEqual(workbook.sheetnames, ["汇总", "逐份", "问题明细"])

        summary_ws = workbook["汇总"]
        values = [summary_ws.cell(row=row, column=1).value for row in range(1, 20)]
        self.assertIn("文件总数", values)
        self.assertIn("通过率", values)
        self.assertIn("规则集", values)
        self.assertIn("生成时间", values)
        self.assertIn("规则", [summary_ws.cell(row=row, column=1).value for row in range(1, 30)])

        per_doc_ws = workbook["逐份"]
        self.assertEqual(
            [cell.value for cell in per_doc_ws[1]], PER_DOC_HEADERS
        )
        self.assertEqual(per_doc_ws.max_row, len(results) + 1)

        issues_ws = workbook["问题明细"]
        self.assertEqual([cell.value for cell in issues_ws[1]], ISSUE_HEADERS)
        detail_rows = [
            [cell.value for cell in row] for row in issues_ws.iter_rows(min_row=2)
        ]
        self.assertTrue(
            any("BODY-001" in row for row in detail_rows), "问题明细应包含注入的 BODY-001"
        )
        self.assertTrue(
            any(row[2] == "失败" for row in detail_rows), "加载失败文档也应出现在问题明细"
        )

    def test_export_csv_merged_rows(self):
        results, summary = self._fixture()
        out = export_csv(results, summary, self.tmp / "batch.csv")
        self.assertTrue(out.is_file())

        with out.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            self.assertEqual(reader.fieldnames, CSV_HEADERS)
            rows = list(reader)

        types = [row["类型"] for row in rows]
        self.assertEqual(types[0], "汇总")
        self.assertEqual(types.count("逐份"), len(results))
        self.assertIn("问题", types)
        self.assertIn("失败", types)
        detail = next(row for row in rows if row["类型"] == "问题")
        self.assertEqual(detail["规则"], "BODY-001")
        self.assertTrue(detail["位置"])
        self.assertIn("12", detail["期望值"])

    def test_export_html_contains_summary(self):
        results, summary = self._fixture()
        page = export_html(summary, results)
        self.assertIsInstance(page, str)
        self.assertIn("<!DOCTYPE html>", page)
        self.assertIn("班级批量检查报告", page)
        self.assertIn("通过率", page)
        self.assertIn("BODY-001", page)
        self.assertIn("问题明细", page)

    def test_export_empty_batch(self):
        from openpyxl import load_workbook

        summary = summarize([])
        xlsx = export_xlsx([], summary, self.tmp / "empty.xlsx")
        self.assertTrue(xlsx.is_file())
        workbook = load_workbook(str(xlsx))
        self.assertEqual(workbook.sheetnames, ["汇总", "逐份", "问题明细"])
        csv_path = export_csv([], summary, self.tmp / "empty.csv")
        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        self.assertEqual(rows[0], CSV_HEADERS)
        self.assertEqual(len(rows), 2, "空批次也应有表头 + 汇总行")
        self.assertIn("通过率", rows[1][11])
        self.assertIsInstance(export_html(summary, []), str)


if __name__ == "__main__":
    unittest.main(verbosity=2)
