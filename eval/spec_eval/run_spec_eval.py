"""LLM 规范抽取评测（held-out）。

把「规范文档 → 规则候选」当成一个有 ground truth 的抽取任务：
- 期望修改的规则及关键参数（人工标注）
- 期望保持不变的规则
- 期望发现的新规则
- 期望识别为「无法程序化」的条款

LLM 输出使用真实运行缓存（JSON），因此评测完全离线、可复现；重新生成缓存：
    python -m tas spec parse eval/demo_spec_other_discipline.txt \
        --base-profile school_default --out-dir eval/spec_parse_other

用法：
    python eval/spec_eval/run_spec_eval.py
    python eval/spec_eval/run_spec_eval.py --case course_2026 --report eval/spec_eval/report.md
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CASE = "course_2026"
DEFAULT_CASES = ("course_2026", "school_2026")
DEFAULT_REPORT = ROOT / "eval" / "spec_eval" / "report.md"
DEFAULT_JSON = ROOT / "eval" / "spec_eval" / "report.json"
TOLERANCE = 0.01


def load_json(path: Path | str) -> Any:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def base_expected(rules_path: Path | str) -> dict[str, dict]:
    payload = load_json(rules_path)
    return {rule["id"]: rule.get("expected", {}) for rule in payload["rules"]}


def _equal(actual: Any, expected: Any) -> bool:
    if isinstance(expected, bool) or isinstance(actual, bool):
        return actual is expected
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(float(actual) - float(expected)) <= TOLERANCE
    if isinstance(expected, list) and isinstance(actual, list):
        return len(expected) == len(actual) and all(
            _equal(a, b) for a, b in zip(actual, expected)
        )
    if isinstance(expected, dict) and isinstance(actual, dict):
        return set(expected) == set(actual) and all(
            _equal(actual[key], expected[key]) for key in expected
        )
    return actual == expected


def derive_actual(
    candidates: list[dict], base: dict[str, dict]
) -> tuple[dict[str, dict], list[str], dict[str, dict]]:
    """从 LLM candidates 推导：修改的规则 / 新规则 / 每条规则输出。"""
    modified: dict[str, dict] = {}
    new_rules: list[str] = []
    by_id: dict[str, dict] = {}
    for candidate in candidates:
        rule_id = str(candidate.get("id", "")).strip()
        if not rule_id:
            continue
        expected = candidate.get("expected") or {}
        by_id[rule_id] = candidate
        if rule_id not in base:
            new_rules.append(rule_id)
        elif not _equal(expected, base[rule_id]):
            modified[rule_id] = expected
    return modified, new_rules, by_id


def _prf(expected: set[str], actual: set[str]) -> tuple[float, float, float]:
    if not expected and not actual:
        return 1.0, 1.0, 1.0
    true_positive = len(expected & actual)
    precision = true_positive / len(actual) if actual else 0.0
    recall = true_positive / len(expected) if expected else 0.0
    f1 = (2 * precision * recall / (precision + recall)) if (precision + recall) else 0.0
    return precision, recall, f1


@dataclass
class CaseResult:
    case_id: str
    title: str
    modified_expected: dict[str, dict]
    modified_actual: dict[str, dict]
    missing_modifications: list[str]
    extra_modifications: list[str]
    modification_precision: float
    modification_recall: float
    modification_f1: float
    param_fields_total: int
    param_fields_correct: int
    param_rules_correct: int
    keep_total: int
    keep_correct: int
    new_expected: list[str]
    new_actual: list[str]
    new_precision: float
    new_recall: float
    new_f1: float
    unmapped_expected: int
    unmapped_actual: int
    unmapped_keyword_total: int
    unmapped_keyword_hits: int
    rule_details: list[dict] = field(default_factory=list)

    @property
    def param_accuracy(self) -> float:
        return self.param_fields_correct / self.param_fields_total if self.param_fields_total else 1.0

    @property
    def keep_accuracy(self) -> float:
        return self.keep_correct / self.keep_total if self.keep_total else 1.0

    @property
    def unmapped_recall(self) -> float:
        if self.unmapped_keyword_total:
            return self.unmapped_keyword_hits / self.unmapped_keyword_total
        return 1.0 if self.unmapped_actual == self.unmapped_expected else 0.0

    @property
    def passed(self) -> bool:
        return (
            self.modification_precision == 1.0
            and self.modification_recall == 1.0
            and self.param_accuracy == 1.0
            and self.keep_accuracy == 1.0
            and self.new_precision == 1.0
            and self.new_recall == 1.0
            and self.unmapped_recall == 1.0
            and self.unmapped_actual == self.unmapped_expected
        )

    def to_dict(self) -> dict:
        return {
            "case": self.case_id,
            "title": self.title,
            "passed": self.passed,
            "modification": {
                "expected": sorted(self.modified_expected),
                "actual": sorted(self.modified_actual),
                "missing": self.missing_modifications,
                "extra": self.extra_modifications,
                "precision": round(self.modification_precision, 4),
                "recall": round(self.modification_recall, 4),
                "f1": round(self.modification_f1, 4),
            },
            "params": {
                "fields_total": self.param_fields_total,
                "fields_correct": self.param_fields_correct,
                "accuracy": round(self.param_accuracy, 4),
                "rules_fully_correct": self.param_rules_correct,
                "rules_total": len(self.modified_expected),
            },
            "keep": {
                "total": self.keep_total,
                "correct": self.keep_correct,
                "accuracy": round(self.keep_accuracy, 4),
            },
            "new_rules": {
                "expected": self.new_expected,
                "actual": self.new_actual,
                "precision": round(self.new_precision, 4),
                "recall": round(self.new_recall, 4),
                "f1": round(self.new_f1, 4),
            },
            "unmapped": {
                "expected": self.unmapped_expected,
                "actual": self.unmapped_actual,
                "keyword_total": self.unmapped_keyword_total,
                "keyword_hits": self.unmapped_keyword_hits,
                "recall": round(self.unmapped_recall, 4),
            },
            "rule_details": self.rule_details,
        }


def evaluate_case(case: dict, *, base: dict[str, dict]) -> CaseResult:
    extraction = case["llm_output_extraction"]
    candidates = extraction.get("candidates") or []
    modified_actual, new_actual, by_id = derive_actual(candidates, base)
    expected_modified = case.get("expected_modified") or {}
    expected_new = list(case.get("expected_new") or [])
    expected_keep = [rule_id for rule_id in base if rule_id not in expected_modified]

    precision, recall, f1 = _prf(set(expected_modified), set(modified_actual))
    missing = sorted(set(expected_modified) - set(modified_actual))
    extra = sorted(set(modified_actual) - set(expected_modified))

    fields_total = 0
    fields_correct = 0
    rules_fully_correct = 0
    rule_details: list[dict] = []
    for rule_id, expected_fields in sorted(expected_modified.items()):
        actual_fields = modified_actual.get(rule_id)
        rule_ok = True
        field_results: dict[str, bool] = {}
        for key, expected_value in expected_fields.items():
            fields_total += 1
            actual_value = (actual_fields or {}).get(key)
            ok = actual_fields is not None and _equal(actual_value, expected_value)
            field_results[key] = ok
            if ok:
                fields_correct += 1
            else:
                rule_ok = False
        if rule_ok:
            rules_fully_correct += 1
        rule_details.append({
            "rule": rule_id,
            "found": actual_fields is not None,
            "fully_correct": rule_ok,
            "fields": field_results,
            "expected": expected_fields,
            "actual": actual_fields,
        })

    keep_correct = sum(
        1
        for rule_id in expected_keep
        if rule_id in by_id and _equal(by_id[rule_id].get("expected") or {}, base[rule_id])
    )

    new_precision, new_recall, new_f1 = _prf(set(expected_new), set(new_actual))

    unmapped = extraction.get("unmapped") or []
    expected_unmapped = case.get("expected_unmapped") or {}
    keyword_total = len(expected_unmapped.get("keywords") or [])
    keyword_hits = 0
    for keyword in expected_unmapped.get("keywords") or []:
        if any(keyword in str(item.get("text", "")) for item in unmapped):
            keyword_hits += 1

    return CaseResult(
        case_id=case["id"],
        title=case.get("title", case["id"]),
        modified_expected=expected_modified,
        modified_actual=modified_actual,
        missing_modifications=missing,
        extra_modifications=extra,
        modification_precision=precision,
        modification_recall=recall,
        modification_f1=f1,
        param_fields_total=fields_total,
        param_fields_correct=fields_correct,
        param_rules_correct=rules_fully_correct,
        keep_total=len(expected_keep),
        keep_correct=keep_correct,
        new_expected=sorted(expected_new),
        new_actual=sorted(new_actual),
        new_precision=new_precision,
        new_recall=new_recall,
        new_f1=new_f1,
        unmapped_expected=int(expected_unmapped.get("count", 0)),
        unmapped_actual=len(unmapped),
        unmapped_keyword_total=keyword_total,
        unmapped_keyword_hits=keyword_hits,
        rule_details=rule_details,
    )


def load_case(case_name: str) -> tuple[Path, dict, dict[str, dict]]:
    case_path = ROOT / "eval" / "spec_eval" / "cases" / f"{case_name}.json"
    case = load_json(case_path)
    llm_path = ROOT / case["llm_output"]
    llm_payload = load_json(llm_path)
    case["llm_output_extraction"] = llm_payload["extraction"]
    rules_path = ROOT / "rules" / "rules.json"
    return case_path, case, base_expected(rules_path)


def render_summary(results: list[CaseResult], cases: list[dict]) -> str:
    lines = [
        "# LLM 规范抽取评测（held-out）",
        "",
        "把「规范文档 → 规则候选」当作有 ground truth 的抽取任务："
        "修改规则发现、参数字段准确率、保持规则准确率、新增规则发现、不可程序化识别。",
        "",
        "| case | 规范 | 修改 P/R/F1 | 参数准确率 | 保持准确率 | 新增 P/R | 不可程序化 | 结论 |",
        "|---|---|---|---:|---:|---|---:|---|",
    ]
    for result, case in zip(results, cases):
        lines.append(
            f"| `{result.case_id}` | {case.get('title', '')} | "
            f"{result.modification_precision:.0%}/{result.modification_recall:.0%}/{result.modification_f1:.0%} | "
            f"{result.param_accuracy:.0%}（{result.param_fields_correct}/{result.param_fields_total}） | "
            f"{result.keep_accuracy:.0%}（{result.keep_correct}/{result.keep_total}） | "
            f"{result.new_precision:.0%}/{result.new_recall:.0%} | "
            f"{result.unmapped_actual}/{result.unmapped_expected} | "
            f"{'通过' if result.passed else '未通过'} |"
        )
    lines += [
        "",
        "> 说明：这是小规模 held-out 规范集，不是标准 benchmark；gold 由人工对照规范原文标注，"
        "LLM 输出为真实 DeepSeek 运行缓存。评测只衡量「抽取与结构化」能力，不评价 LLM 直接修改文档；"
        "实际修复由确定性代码执行并复检。",
        "",
    ]
    return "\n".join(lines)


def render_markdown(result: CaseResult, case: dict) -> str:
    def pct(value: float) -> str:
        return f"{value:.1%}"

    lines = [
        "# LLM 规范抽取评测（held-out）",
        "",
        f"- case：`{result.case_id}`（{result.title}）",
        f"- 规范原文：`{case['spec_file']}`",
        f"- LLM 输出缓存：`{case['llm_output']}`（model：`deepseek-chat`）",
        f"- 标注说明：{case.get('annotation_note', '人工标注')}",
        "",
        "## 总览",
        "",
        "| 指标 | 结果 |",
        "|---|---:|",
        f"| 修改规则 P / R / F1 | {pct(result.modification_precision)} / {pct(result.modification_recall)} / {pct(result.modification_f1)} |",
        f"| 修改规则命中 | {len(result.modified_expected) - len(result.missing_modifications)}/{len(result.modified_expected)} ｜ 缺失 {result.missing_modifications or '无'} ｜ 多报 {result.extra_modifications or '无'} |",
        f"| 参数字段准确率 | {pct(result.param_accuracy)}（{result.param_fields_correct}/{result.param_fields_total} 字段，{result.param_rules_correct}/{len(result.modified_expected)} 条规则全对） |",
        f"| 保持规则准确率 | {pct(result.keep_accuracy)}（{result.keep_correct}/{result.keep_total}） |",
        f"| 新增规则 P / R / F1 | {pct(result.new_precision)} / {pct(result.new_recall)} / {pct(result.new_f1)} ｜ 期望 {result.new_expected or '无'} ｜ 实际 {result.new_actual or '无'} |",
        f"| 不可程序化识别 | 期望 {result.unmapped_expected} 条 ｜ 实际 {result.unmapped_actual} 条 ｜ 关键词召回 {pct(result.unmapped_recall)} |",
        f"| **结论** | **{'通过' if result.passed else '未通过'}** |",
        "",
        "## 逐条修改",
        "",
        "| 规则 | 是否找到 | 是否全对 | 期望字段 |",
        "|---|---|---|---|",
    ]
    for item in result.rule_details:
        expected_text = json.dumps(item["expected"], ensure_ascii=False)
        lines.append(
            f"| {item['rule']} | {'是' if item['found'] else '否'} | "
            f"{'是' if item['fully_correct'] else '否'} | `{expected_text}` |"
        )
    lines += [
        "",
        "> 说明：这是小规模 held-out 规范集（5 条期望修改/38 条基础规则），不是标准 benchmark。"
        "评测只衡量「抽取与结构化」能力，不评价 LLM 直接修改文档；实际修复由确定性代码执行并复检。",
        "",
    ]
    return "\n".join(lines)


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="LLM 规范抽取评测")
    parser.add_argument(
        "--cases",
        default=",".join(DEFAULT_CASES),
        help="逗号分隔的 case 名称（eval/spec_eval/cases/）",
    )
    parser.add_argument("--report", default=str(DEFAULT_REPORT), help="Markdown 报告路径")
    parser.add_argument("--json", default=str(DEFAULT_JSON), help="JSON 结果路径")
    args = parser.parse_args(argv)

    cases = []
    results = []
    for name in [item.strip() for item in args.cases.split(",") if item.strip()]:
        case_path, case, base = load_case(name)
        cases.append(case)
        results.append(evaluate_case(case, base=base))

    payload = {
        "summary": {
            "total": len(results),
            "passed": sum(1 for result in results if result.passed),
            "all_passed": all(result.passed for result in results),
        },
        "cases": [result.to_dict() | {"annotated_case": f"eval/spec_eval/cases/{result.case_id}.json"} for result in results],
    }
    report_sections = [render_summary(results, cases)]
    report_sections += [render_markdown(result, case) for result, case in zip(results, cases)]
    report_text = "\n---\n\n".join(report_sections)

    report_path = Path(args.report)
    json_path = Path(args.json)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(report_text, encoding="utf-8")
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    for result in results:
        print(
            f"{result.case_id}: mods P/R/F1 {result.modification_precision:.1%}/"
            f"{result.modification_recall:.1%}/{result.modification_f1:.1%}, "
            f"params {result.param_fields_correct}/{result.param_fields_total}, "
            f"keep {result.keep_correct}/{result.keep_total}, "
            f"new {len(result.new_actual)}/{len(result.new_expected)}, "
            f"unmapped {result.unmapped_actual}/{result.unmapped_expected} -> "
            f"{'PASS' if result.passed else 'FAIL'}"
        )
    print(f"all_passed: {payload['summary']['all_passed']}")
    print(f"report: {report_path}")
    print(f"json: {json_path}")
    return 0 if payload["summary"]["all_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
