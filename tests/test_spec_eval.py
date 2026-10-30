"""LLM 规范抽取评测：离线单元测试（不联网、不调用 LLM）。"""

from __future__ import annotations

import copy
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "eval" / "spec_eval"))

import run_spec_eval  # noqa: E402


class SpecEvalTest(unittest.TestCase):
    def setUp(self) -> None:
        self.case_path, self.case, self.base = run_spec_eval.load_case("course_2026")

    def _evaluate(self) -> "run_spec_eval.CaseResult":
        return run_spec_eval.evaluate_case(self.case, base=self.base)

    def test_course_case_passes_with_expected_metrics(self) -> None:
        result = self._evaluate()
        self.assertTrue(result.passed)
        self.assertEqual(result.modification_precision, 1.0)
        self.assertEqual(result.modification_recall, 1.0)
        self.assertEqual(result.modification_f1, 1.0)
        self.assertEqual(result.param_fields_correct, result.param_fields_total)
        self.assertEqual(result.keep_correct, result.keep_total)
        self.assertEqual(len(result.new_expected), 0)
        self.assertEqual(result.unmapped_actual, result.unmapped_expected)
        self.assertEqual(
            set(result.modified_expected),
            {"PAGE-002", "BODY-001", "BODY-002", "HEAD-001", "HEAD-002"},
        )

    def test_school_case_detects_new_rules_and_unmapped(self) -> None:
        _, case, base = run_spec_eval.load_case("school_2026")
        result = run_spec_eval.evaluate_case(case, base=base)
        self.assertTrue(result.passed)
        self.assertEqual(len(result.new_actual), 6)
        self.assertEqual(result.new_precision, 1.0)
        self.assertEqual(result.new_recall, 1.0)
        self.assertEqual(result.unmapped_actual, 23)
        self.assertEqual(result.keep_correct, result.keep_total)

    def test_extra_modification_is_reported(self) -> None:
        case = copy.deepcopy(self.case)
        candidates = case["llm_output_extraction"]["candidates"]
        target = next(item for item in candidates if item["id"] == "BODY-003")
        target["expected"] = dict(target["expected"] or {})
        target["expected"]["first_line_chars"] = 100
        result = run_spec_eval.evaluate_case(case, base=self.base)
        self.assertLess(result.modification_precision, 1.0)
        self.assertIn("BODY-003", result.extra_modifications)
        self.assertFalse(result.passed)

    def test_wrong_parameter_is_reported(self) -> None:
        case = copy.deepcopy(self.case)
        candidates = case["llm_output_extraction"]["candidates"]
        target = next(item for item in candidates if item["id"] == "BODY-001")
        target["expected"] = dict(target["expected"] or {})
        target["expected"]["size_pt"] = 11
        result = run_spec_eval.evaluate_case(case, base=self.base)
        self.assertLess(result.param_accuracy, 1.0)
        detail = next(item for item in result.rule_details if item["rule"] == "BODY-001")
        self.assertFalse(detail["fully_correct"])
        self.assertFalse(detail["fields"]["size_pt"])
        self.assertFalse(result.passed)

    def test_missing_modification_is_reported(self) -> None:
        case = copy.deepcopy(self.case)
        extraction = case["llm_output_extraction"]
        extraction["candidates"] = [
            item for item in extraction["candidates"] if item["id"] != "BODY-002"
        ]
        result = run_spec_eval.evaluate_case(case, base=self.base)
        self.assertIn("BODY-002", result.missing_modifications)
        self.assertLess(result.modification_recall, 1.0)
        self.assertFalse(result.passed)

    def test_summary_contains_both_cases(self) -> None:
        results = [self._evaluate()]
        _, school_case, school_base = run_spec_eval.load_case("school_2026")
        results.append(run_spec_eval.evaluate_case(school_case, base=school_base))
        report = run_spec_eval.render_summary(
            results, [self.case, school_case]
        )
        self.assertIn("course_2026", report)
        self.assertIn("school_2026", report)
        self.assertIn("100%", report)


if __name__ == "__main__":
    unittest.main(verbosity=2)
