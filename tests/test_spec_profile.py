"""spec_profile（规范→规则候选→profile）单元测试：不联网，注入 fake chat。"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from tas.config import RuleSet  # noqa: E402
from tas.spec_profile import (  # noqa: E402
    Candidate,
    build_messages,
    build_profile_payload,
    diff_candidates,
    extract_candidates,
    parse_llm_payload,
    register_profile,
    render_report_markdown,
    write_profile_payload,
    apply_candidates_file,
)

RULES = RuleSet.load()
BASE_PROFILE = "school_default"


def _rule_payload(rule_id: str) -> dict:
    rule = RULES.get(rule_id)
    return {
        "id": rule.id,
        "name": rule.name,
        "category": rule.category,
        "phase": rule.phase,
        "severity": rule.severity,
        "profiles": list(rule.profiles),
        "spec": rule.spec,
        "expected": dict(rule.expected),
        "source": rule.source,
        "confidence": 0.9,
    }


FAKE_PAYLOAD = {
    "rules": [
        _rule_payload("PAGE-001"),  # keep
        {  # modify：正文改字号
            **_rule_payload("BODY-001"),
            "expected": {**dict(RULES.get("BODY-001").expected), "size_pt": 14},
            "confidence": 0.8,
        },
        {  # new
            "id": "BODY-005",
            "name": "正文段后距",
            "category": "body",
            "phase": "P0",
            "severity": "warning",
            "profiles": ["*"],
            "spec": "正文段后无间距（L91 附近示例）",
            "expected": {"after_lines": 0.0},
            "source": "L91",
            "confidence": 0.7,
            "note": "示例新增规则",
        },
        {  # 非法 id
            "id": "bad id",
            "name": "非法规则",
            "category": "body",
            "expected": {"x": 1},
        },
        {  # 空 expected
            "id": "BODY-006",
            "name": "空期望值",
            "category": "body",
            "expected": {},
        },
    ],
    "unmapped": [
        {"text": "论文应体现实事求是精神", "source": "L1", "reason": "内容要求，无法程序化检查"}
    ],
    "summary": "已抽取排版规则，内容要求进入 unmapped。",
}

NEW_FORMAT_PAYLOAD = {
    "existing_assessment": [
        {"id": "PAGE-001", "action": "keep"},
        {
            "id": "BODY-001",
            "action": "modify",
            "expected": {**dict(RULES.get("BODY-001").expected), "size_pt": 14},
            "spec": "正文小四号（假设规范改为四号）",
            "source": "L91",
            "confidence": 0.9,
            "note": "LLM 建议改字号",
        },
        {"id": "NOPE-001", "action": "keep"},  # 不存在的 id，应进 dropped
        {"id": "BODY-002", "action": "modify"},  # 缺 expected，应进 dropped
    ],
    "new_rules": [
        {
            "id": "BODY-005",
            "name": "正文段后距",
            "category": "body",
            "phase": "P0",
            "severity": "warning",
            "profiles": ["*"],
            "spec": "正文段后无间距",
            "expected": {"after_lines": 0.0},
            "source": "L91",
            "confidence": 0.7,
        }
    ],
    "unmapped": [{"text": "论文应体现实事求是精神", "source": "L1", "reason": "内容要求"}],
    "summary": "逐条评估 + 新增。",
}


def fake_chat(payload):
    def _chat(messages):
        assert isinstance(messages, list) and messages[0]["role"] == "system"
        assert messages[-1]["role"] == "user"
        return payload

    return _chat


class PromptTest(unittest.TestCase):
    def test_prompt_contains_spec_lines_and_existing_rules(self):
        messages = build_messages("一、页面：A4纸，页边距 2.54cm。\n二、正文：小四宋体。", RULES)
        self.assertEqual(messages[0]["role"], "system")
        self.assertIn("json", messages[0]["content"].lower())
        user = messages[-1]["content"]
        self.assertIn("L1:", user)
        self.assertIn("PAGE-001", user)
        self.assertIn("expected", user)

    def test_number_spec_lines_truncates(self):
        text = "\n".join(f"第{index}行规范内容" for index in range(1, 200))
        messages = build_messages(text, RULES, max_spec_chars=200)
        user = messages[-1]["content"]
        self.assertIn("已截断", user)


class ParseTest(unittest.TestCase):
    def test_parse_payload_validates_and_drops(self):
        candidates, unmapped, dropped, summary = parse_llm_payload(FAKE_PAYLOAD)
        self.assertEqual(len(candidates), 3)
        self.assertEqual(len(dropped), 2)
        self.assertEqual(len(unmapped), 1)
        self.assertIn("抽取排版规则", summary)

    def test_parse_payload_requires_rules_array(self):
        with self.assertRaises(ValueError):
            parse_llm_payload({"unmapped": []})
        with self.assertRaises(ValueError):
            parse_llm_payload([])

    def test_candidate_defaults(self):
        candidate = Candidate.from_dict({"id": "X-001", "name": "n", "category": "body", "expected": {"a": 1}})
        self.assertEqual(candidate.phase, "P1")
        self.assertEqual(candidate.severity, "warning")
        self.assertEqual(candidate.profiles, ["*"])

    def test_assessment_format_keep_modify_new(self):
        report = extract_candidates(
            "规范原文", RULES, chat=fake_chat(NEW_FORMAT_PAYLOAD), model="deepseek-chat"
        )
        ids = [candidate.id for candidate in report.candidates]
        self.assertEqual(sorted(ids), ["BODY-001", "BODY-005", "PAGE-001"])
        self.assertEqual(len(report.dropped), 2)  # NOPE-001 + BODY-002 缺 expected
        diff = diff_candidates(report, RULES, base_profile=BASE_PROFILE)
        actions = {item.candidate.id: item.action for item in diff.diffs}
        self.assertEqual(actions["PAGE-001"], "keep")
        self.assertEqual(actions["BODY-001"], "modify")
        self.assertEqual(actions["BODY-005"], "new")
        markdown = render_report_markdown(report, diff)
        self.assertIn("逐条评估覆盖", markdown)


class DiffTest(unittest.TestCase):
    def setUp(self):
        self.report = extract_candidates(
            "规范原文", RULES, chat=fake_chat(FAKE_PAYLOAD), model="deepseek-chat"
        )
        self.diff = diff_candidates(self.report, RULES, base_profile=BASE_PROFILE)

    def test_classification(self):
        actions = {item.candidate.id: item.action for item in self.diff.diffs}
        self.assertEqual(actions["PAGE-001"], "keep")
        self.assertEqual(actions["BODY-001"], "modify")
        self.assertEqual(actions["BODY-005"], "new")
        self.assertTrue(any("size_pt" in change for change in next(
            item.changes for item in self.diff.diffs if item.candidate.id == "BODY-001"
        )))

    def test_counts_and_uncovered(self):
        counts = self.diff.counts()
        self.assertEqual(counts, {"keep": 1, "modify": 1, "new": 1})
        # 38 条现有规则里被候选触及 2 条（PAGE-001/BODY-001）
        self.assertEqual(len(self.diff.uncovered_existing), len(RULES) - 2)

    def test_report_markdown(self):
        markdown = render_report_markdown(self.report, self.diff)
        self.assertIn("新增规则", markdown)
        self.assertIn("修改现有规则", markdown)
        self.assertIn("BODY-005", markdown)
        self.assertIn("unmapped", markdown)


class ProfileBuildTest(unittest.TestCase):
    def setUp(self):
        self.report = extract_candidates(
            "规范原文", RULES, chat=fake_chat(FAKE_PAYLOAD), model="deepseek-chat"
        )
        self.diff = diff_candidates(self.report, RULES, base_profile=BASE_PROFILE)

    def test_build_and_write_profile(self):
        payload = build_profile_payload(
            RULES,
            self.diff,
            profile_id="course_demo",
            name="演示课程规范",
            discipline="sci_eng",
            source_spec="demo.txt",
        )
        self.assertEqual(payload["meta"]["id"], "course_demo")
        self.assertEqual(payload["meta"]["derived_from"], BASE_PROFILE)
        self.assertEqual(len(payload["rules"]), len(RULES) + 1)
        body_001 = next(rule for rule in payload["rules"] if rule["id"] == "BODY-001")
        self.assertEqual(body_001["expected"]["size_pt"], 14)
        new_rule = next(rule for rule in payload["rules"] if rule["id"] == "BODY-005")
        self.assertIn("待实现", new_rule["check"])

        with tempfile.TemporaryDirectory() as tmp:
            path = write_profile_payload(payload, Path(tmp) / "profiles" / "course_demo.rules.json")
            loaded = RuleSet.load(path)
            self.assertEqual(len(loaded), len(RULES) + 1)
            self.assertEqual(loaded.get("BODY-001").expected["size_pt"], 14)

    def test_register_profile(self):
        payload = build_profile_payload(
            RULES, self.diff, profile_id="course_demo", name="演示课程规范",
            discipline="sci_eng", source_spec="demo.txt",
        )
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            rules_path = write_profile_payload(payload, root / "profiles" / "course_demo.rules.json")
            index_path = register_profile(
                root / "profiles" / "index.json",
                profile_id="course_demo",
                name="演示课程规范",
                discipline="sci_eng",
                rules_path=rules_path,
                notes="测试",
            )
            data = json.loads(index_path.read_text(encoding="utf-8"))
            entry = next(item for item in data["profiles"] if item["id"] == "course_demo")
            self.assertEqual(entry["rules"], "course_demo.rules.json")
            self.assertEqual(entry["discipline"], "sci_eng")

    def test_apply_candidates_file(self):
        extraction = self.report.to_dict()
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            candidates_path = root / "candidates.json"
            candidates_path.write_text(
                json.dumps({"base_profile": BASE_PROFILE, "extraction": extraction}, ensure_ascii=False),
                encoding="utf-8",
            )
            result = apply_candidates_file(
                candidates_path,
                profile_id="course_demo",
                name="演示课程规范",
                discipline="sci_eng",
                out_path=root / "profiles" / "course_demo.rules.json",
                index_path=root / "profiles" / "index.json",
                source_spec="demo.txt",
            )
            self.assertTrue(result["registered"])
            self.assertEqual(result["rules_total"], len(RULES) + 1)
            self.assertEqual(result["modify"], 1)
            self.assertEqual(result["new"], 1)
            loaded = RuleSet.load(result["rules_path"])
            self.assertEqual(len(loaded), len(RULES) + 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
