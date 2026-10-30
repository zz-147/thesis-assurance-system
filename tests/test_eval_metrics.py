"""评测脚本（eval/run_eval.py）指标计算与全流程 smoke 测试。

覆盖：
1) 指标纯函数：locator 解析、检出率/定位准确率聚合、误报率聚合、修复聚合；
2) 报告渲染 / 落盘（report.md + report.json）；
3) 全流程 smoke：run_evaluation 用少量文档与每类 1 处注入跑通，断言指标范围与结构。
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "eval"))

import run_eval  # noqa: E402
from tas.report import Issue  # noqa: E402


def _issue(rule_id: str, locator: str, severity: str = "error") -> Issue:
    return Issue(
        rule_id=rule_id,
        title=rule_id,
        severity=severity,
        message="test issue",
        locator=locator,
    )


class LocatorTest(unittest.TestCase):
    def test_paragraph_locators(self):
        self.assertEqual(run_eval.locator_paragraph_index("body/para/12「正文」"), 12)
        self.assertEqual(run_eval.locator_paragraph_index("front/title_cn/5「题目」"), 5)
        self.assertEqual(run_eval.locator_paragraph_index("front/abstract_cn/8「摘要」"), 8)
        self.assertEqual(run_eval.locator_paragraph_index("body/para/0「开头」"), 0)

    def test_page_and_table_locators_are_not_paragraphs(self):
        self.assertIsNone(run_eval.locator_paragraph_index("section/1"))
        self.assertIsNone(run_eval.locator_paragraph_index("table/3/row/2"))
        self.assertIsNone(run_eval.locator_paragraph_index("front/keywords_cn"))
        self.assertIsNone(run_eval.locator_paragraph_index(""))


class AggregateDetectionTest(unittest.TestCase):
    def test_recall_and_locate_accuracy(self):
        records = [
            {"category": "正文字号", "rule_id": "BODY-001", "scope": "paragraph",
             "rule_hit": True, "located": True},
            {"category": "正文字号", "rule_id": "BODY-001", "scope": "paragraph",
             "rule_hit": True, "located": True},
            {"category": "正文字号", "rule_id": "BODY-001", "scope": "paragraph",
             "rule_hit": True, "located": False},  # 规则命中但定位错
            {"category": "页边距", "rule_id": "PAGE-002", "scope": "page",
             "rule_hit": True, "located": True},
            {"category": "页边距", "rule_id": "PAGE-002", "scope": "page",
             "rule_hit": False, "located": False},  # 漏检
        ]
        metrics = run_eval.aggregate_detection(records)
        self.assertEqual(metrics["injections"], 5)
        self.assertEqual(metrics["rule_hits"], 4)
        self.assertEqual(metrics["located"], 3)
        self.assertAlmostEqual(metrics["recall"], 3 / 5)
        self.assertAlmostEqual(metrics["locate_accuracy"], 3 / 4)

        categories = {item["category"]: item for item in metrics["categories"]}
        self.assertAlmostEqual(categories["正文字号"]["recall"], 2 / 3)
        self.assertAlmostEqual(categories["页边距"]["recall"], 1 / 2)
        per_rule = {item["rule_id"]: item for item in metrics["per_rule"]}
        self.assertEqual(per_rule["BODY-001"]["injections"], 3)
        self.assertEqual(per_rule["PAGE-002"]["injections"], 2)


class FalsePositiveTest(unittest.TestCase):
    def test_fp_rate_and_exclusion_of_unevaluated_rules(self):
        docs = [
            {
                "paragraphs": 10,
                "issues": [
                    _issue("BODY-001", "body/para/2「正文」"),
                    _issue("BODY-001", "body/para/2「正文」"),  # 同段重复只算 1 段
                    _issue("TITLE-001", "front/title_cn/5「题目」"),
                ],
            },
            {
                "paragraphs": 10,
                "issues": [
                    _issue("PAGE-002", "section/1", severity="error"),  # 页面级
                    _issue("HF-001", "section/1/header", severity="error"),  # 评测集外
                ],
            },
        ]
        metrics = run_eval.aggregate_false_positives(
            docs, ["BODY-001", "TITLE-001", "PAGE-002"]
        )
        self.assertEqual(metrics["good_docs"], 2)
        self.assertEqual(metrics["paragraphs"], 20)
        self.assertEqual(metrics["flagged_paragraphs"], 2)
        self.assertAlmostEqual(metrics["fp_rate"], 0.1)
        self.assertEqual(metrics["page_level_issues"], 1)
        self.assertEqual(metrics["other_rule_hits"], {"HF-001": 1})

    def test_no_false_positive(self):
        metrics = run_eval.aggregate_false_positives(
            [{"paragraphs": 12, "issues": []}], ["BODY-001"]
        )
        self.assertEqual(metrics["flagged_paragraphs"], 0)
        self.assertEqual(metrics["fp_rate"], 0.0)


class RepairAggregationTest(unittest.TestCase):
    def test_repair_summary(self):
        repairs = [
            {"injected_ok": True, "full_ok": True, "hash_ok": True, "fix_s": 0.2},
            {"injected_ok": True, "full_ok": False, "hash_ok": True, "fix_s": 0.4},
            {"injected_ok": False, "full_ok": False, "hash_ok": False, "fix_s": 0.3},
        ]
        summary = run_eval.summarize_repairs(repairs)
        self.assertEqual(summary["docs"], 3)
        self.assertAlmostEqual(summary["injected_pass_rate"], 2 / 3)
        self.assertAlmostEqual(summary["full_pass_rate"], 1 / 3)
        self.assertAlmostEqual(summary["hash_unchanged_rate"], 2 / 3)
        self.assertAlmostEqual(summary["mean_fix_s"], 0.3)
        self.assertEqual(len(summary["failures"]), 1)

    def test_empty_repairs(self):
        summary = run_eval.summarize_repairs([])
        self.assertEqual(summary["docs"], 0)
        self.assertEqual(summary["injected_pass_rate"], 0.0)
        self.assertEqual(summary["failures"], [])


def _fake_payload() -> dict:
    return {
        "meta": {
            "tool": "tas",
            "version": "0.1.0",
            "generated_at": "2026-10-25T00:00:00+08:00",
            "python": "3.12.0",
            "platform": "test",
            "profile_id": "school_default",
            "profile_name": "学院默认",
            "discipline": "sci_eng",
            "ruleset_id": "school_default",
            "ruleset_total_rules": 38,
            "implemented_rules": ["BODY-001", "PAGE-002"],
            "implemented_count": 2,
            "implemented_p0_rules": ["BODY-001", "PAGE-002"],
            "evaluated_rules": ["BODY-001", "PAGE-002"],
            "evaluated_categories": ["正文字号", "页边距"],
            "unevaluated_implemented_p0_rules": [],
            "skipped_categories": [],
            "injections_per_category": 5,
            "eval_docs": 10,
            "good_docs": 3,
            "total_elapsed_s": 1.0,
            "notes": [],
        },
        "detection": {
            "injections": 10,
            "rule_hits": 10,
            "located": 9,
            "recall": 0.9,
            "locate_accuracy": 0.9,
            "categories": [
                {"category": "正文字号", "rule_id": "BODY-001", "scope": "paragraph",
                 "injections": 5, "rule_hits": 5, "located": 5,
                 "recall": 1.0, "locate_accuracy": 1.0},
                {"category": "页边距", "rule_id": "PAGE-002", "scope": "page",
                 "injections": 5, "rule_hits": 5, "located": 4,
                 "recall": 0.8, "locate_accuracy": 0.8},
            ],
            "per_rule": [],
            "skipped_categories": [],
            "check_errors": [],
            "docs": 4,
        },
        "false_positives": {
            "good_docs": 3,
            "paragraphs": 300,
            "flagged_paragraphs": 3,
            "fp_rate": 0.01,
            "page_level_issues": 0,
            "rule_hits": {"BODY-001": 3},
            "other_rule_hits": {},
            "check_errors": [],
        },
        "repair": {
            "docs": 10,
            "injected_pass_rate": 1.0,
            "full_pass_rate": 0.9,
            "hash_unchanged_rate": 1.0,
            "mean_fix_s": 0.2,
            "total_fix_s": 2.0,
            "mean_check_s": 0.08,
            "failures": [],
            "fix_failures": [],
            "batch_check_elapsed_s": 0.8,
            "batch_check_docs": 10,
            "docs_per_s": 12.5,
            "ms_per_doc": 80.0,
            "check_errors": [],
        },
        "throughput": {
            "docs": 10,
            "elapsed_s": 0.8,
            "docs_per_s": 12.5,
            "ms_per_doc": 80.0,
            "mean_check_s": 0.08,
        },
        "timings": {"total_elapsed_s": 1.0},
    }


class ReportRenderTest(unittest.TestCase):
    def test_markdown_contains_tables_and_metrics(self):
        markdown = run_eval.render_markdown(_fake_payload())
        for keyword in (
            "# tas 评测报告",
            "## 1. 规则覆盖",
            "## 2. 变异检出率",
            "## 3. 误报率",
            "## 4. 修复评测",
            "## 5. 批量吞吐",
            "检出率（recall）",
            "90.0%",
            "通过率",
            "| 指标 | 数值 | 说明 |",
        ):
            self.assertIn(keyword, markdown)

    def test_write_reports_writes_markdown_and_json(self):
        payload = _fake_payload()
        with tempfile.TemporaryDirectory() as tmp:
            md_path, json_path = run_eval.write_reports(
                payload, Path(tmp) / "reports" / "report.md"
            )
            self.assertTrue(md_path.is_file())
            self.assertTrue(json_path.is_file())
            self.assertEqual(json_path.name, "report.json")
            reloaded = json.loads(json_path.read_text(encoding="utf-8"))
            self.assertEqual(reloaded["meta"]["implemented_count"], 2)
            self.assertIn("检出率", md_path.read_text(encoding="utf-8"))


class EvaluationSmokeTest(unittest.TestCase):
    """小规模跑通完整评测流程（不写仓库文件）。"""

    @classmethod
    def setUpClass(cls):
        cls.payload = run_eval.run_evaluation(docs=10, injections=1)

    def test_meta_reports_dynamic_rule_coverage(self):
        meta = self.payload["meta"]
        self.assertGreaterEqual(meta["implemented_count"], 1)
        self.assertLessEqual(meta["implemented_count"], meta["ruleset_total_rules"])
        self.assertTrue(meta["evaluated_rules"])
        self.assertTrue(meta["generated_at"])

    def test_detection_metrics(self):
        detection = self.payload["detection"]
        self.assertGreaterEqual(len(detection["categories"]), 8, "至少 8 类变异")
        self.assertGreaterEqual(detection["injections"], len(detection["categories"]))
        self.assertGreaterEqual(detection["located"], 0)
        self.assertLessEqual(detection["located"], detection["injections"])
        self.assertGreaterEqual(detection["recall"], 0.0)
        self.assertLessEqual(detection["recall"], 1.0)
        self.assertGreaterEqual(detection["locate_accuracy"], 0.0)
        self.assertLessEqual(detection["locate_accuracy"], 1.0)
        for category in detection["categories"]:
            self.assertIn("rule_id", category)
            self.assertGreaterEqual(category["injections"], 1)
            self.assertLessEqual(category["located"], category["injections"])

    def test_false_positive_metrics(self):
        fp = self.payload["false_positives"]
        self.assertGreater(fp["paragraphs"], 0)
        self.assertGreaterEqual(fp["fp_rate"], 0.0)
        self.assertLessEqual(fp["fp_rate"], 1.0)
        self.assertEqual(fp["flagged_paragraphs"] / fp["paragraphs"], fp["fp_rate"])

    def test_repair_and_throughput_metrics(self):
        repair = self.payload["repair"]
        self.assertEqual(repair["docs"], 10)
        for key in ("injected_pass_rate", "full_pass_rate", "hash_unchanged_rate"):
            self.assertGreaterEqual(repair[key], 0.0)
            self.assertLessEqual(repair[key], 1.0)
        self.assertGreaterEqual(repair["mean_fix_s"], 0.0)
        self.assertGreaterEqual(repair["mean_check_s"], 0.0)

        throughput = self.payload["throughput"]
        self.assertEqual(throughput["docs"], 10)
        self.assertGreater(throughput["elapsed_s"], 0.0)
        self.assertGreater(throughput["docs_per_s"], 0.0)
        self.assertGreater(throughput["ms_per_doc"], 0.0)

    def test_payload_is_json_serializable_and_renders(self):
        text = json.dumps(self.payload, ensure_ascii=False)
        self.assertIn("detection", text)
        markdown = run_eval.render_markdown(self.payload)
        self.assertIn("评测报告", markdown)


if __name__ == "__main__":
    unittest.main(verbosity=2)
