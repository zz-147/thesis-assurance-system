"""参数化测试：检查器必须消费当前 ruleset 的 expected，而不是写死学校默认值。"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

from tas.checkers import run_checks  # noqa: E402
from tas.config import RuleSet, default_rules_path  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.roles import classify  # noqa: E402
from test_check import build_good_docx  # noqa: E402


def build_custom_ruleset(tmp: Path) -> RuleSet:
    """复制学校规则库，改成另一套"学科规范"参数。"""
    data = json.loads(default_rules_path().read_text(encoding="utf-8"))
    for rule in data["rules"]:
        if rule["id"] == "BODY-001":
            rule["expected"] = {"size_pt": 10.5, "east_asia": "仿宋", "ascii": "Times New Roman"}
        elif rule["id"] == "BODY-002":
            rule["expected"] = {"line_spacing": 1.75}
        elif rule["id"] == "BODY-003":
            rule["expected"] = {"first_line_chars": 150}
        elif rule["id"] == "PAGE-002":
            rule["expected"] = {
                "margin_top_cm": 2.5,
                "margin_bottom_cm": 2.5,
                "margin_left_cm": 3.0,
                "margin_right_cm": 3.0,
                "gutter_cm": 1.0,
                "header_distance_cm": 2.0,
                "footer_distance_cm": 2.0,
            }
        elif rule["id"] == "HEAD-001":
            rule["expected"] = {**rule["expected"], "size_pt": 14}
    path = tmp / "custom.rules.json"
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return RuleSet.load(path)


class ParameterizedCheckTest(unittest.TestCase):
    def test_default_ruleset_keeps_fixture_compliant(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_good_docx(Path(tmp) / "good.docx")
            model = DocModel.load(path)
            issues = run_checks(model, classify(model), RuleSet.load(), profile="sci_eng")
            checked = [i for i in issues if i.rule_id in ("BODY-001", "BODY-002", "BODY-003", "PAGE-002", "HEAD-001")]
            self.assertEqual(checked, [])

    def test_custom_params_are_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_good_docx(Path(tmp) / "good.docx")
            model = DocModel.load(path)
            ruleset = build_custom_ruleset(Path(tmp))
            issues = run_checks(model, classify(model), ruleset, profile="sci_eng")
            ids = {issue.rule_id for issue in issues}
            for expected_rule in ("BODY-001", "BODY-002", "BODY-003", "PAGE-002", "HEAD-001"):
                self.assertIn(expected_rule, ids, f"自定义参数下 {expected_rule} 应报问题")

    def test_custom_params_are_reported_in_expected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_good_docx(Path(tmp) / "good.docx")
            model = DocModel.load(path)
            ruleset = build_custom_ruleset(Path(tmp))
            issues = run_checks(model, classify(model), ruleset, profile="sci_eng")
            body_issue = next(issue for issue in issues if issue.rule_id == "BODY-001")
            self.assertIn("仿宋", str(body_issue.expected))
            self.assertIn("10.5", str(body_issue.expected))


if __name__ == "__main__":
    unittest.main(verbosity=2)
