"""规范文档 → 规则候选（LLM 抽取）→ 与现有规则 diff → 生成新 profile。

教师端核心链路：
    规范文本 ──LLM──> 候选规则（带出处/置信度）──本地校验──> 与现有规则 diff
    ──教师审阅（candidates.json / diff.md）──> tas spec apply ──> profiles/<id>.rules.json

设计原则：
- LLM 只产出「候选」，不直接写规则库；本地做结构校验与 diff，教师确认后才落盘。
- 候选必须带规范出处（source）与置信度；无法编程检查的要求进 unmapped，不硬凑。
- 不改任何学生文档；不联网验证文献（那是 refs 模块）。
"""

from __future__ import annotations

import copy
import json
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Optional

from . import __version__
from .config import DEFAULT_PROFILE_ID, RuleSet, profiles_index_path, resolve_profile

KNOWN_CATEGORIES = {
    "page", "body", "heading", "front", "toc",
    "header_footer", "float", "references", "structure",
}
KNOWN_PHASES = {"P0", "P1"}
KNOWN_SEVERITIES = {"error", "warning", "info"}
RULE_ID_RE = re.compile(r"^[A-Z][A-Z_]*-\d{2,3}$")
DEFAULT_SPEC_CHAR_LIMIT = 24000

# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------


@dataclass
class Candidate:
    """LLM 抽取出的一条规则候选。"""

    id: str
    name: str
    category: str
    phase: str
    severity: str
    profiles: list[str]
    spec: str
    expected: dict
    source: str = ""
    confidence: float = 0.5
    note: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Candidate":
        return cls(
            id=str(data.get("id", "")).strip(),
            name=str(data.get("name", "")).strip(),
            category=str(data.get("category", "")).strip(),
            phase=str(data.get("phase", "P1")).strip().upper(),
            severity=str(data.get("severity", "warning")).strip().lower(),
            profiles=[str(item) for item in (data.get("profiles") or ["*"])],
            spec=str(data.get("spec", "")).strip(),
            expected=data.get("expected") if isinstance(data.get("expected"), dict) else {},
            source=str(data.get("source", "")).strip(),
            confidence=float(data.get("confidence", 0.5) or 0.5),
            note=str(data.get("note", "")).strip(),
        )


@dataclass
class ExtractionReport:
    """一次 LLM 抽取的完整结果。"""

    base_profile: str
    base_rules_count: int
    candidates: list[Candidate] = field(default_factory=list)
    unmapped: list[dict] = field(default_factory=list)
    dropped: list[dict] = field(default_factory=list)
    summary: str = ""
    model: str = ""
    spec_chars: int = 0

    def to_dict(self) -> dict:
        return {
            "base_profile": self.base_profile,
            "base_rules_count": self.base_rules_count,
            "model": self.model,
            "spec_chars": self.spec_chars,
            "summary": self.summary,
            "candidates": [candidate.to_dict() for candidate in self.candidates],
            "unmapped": self.unmapped,
            "dropped": self.dropped,
        }


@dataclass
class CandidateDiff:
    action: str  # keep / modify / new
    candidate: Candidate
    existing_id: Optional[str]
    changes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "action": self.action,
            "candidate": self.candidate.to_dict(),
            "existing_id": self.existing_id,
            "changes": self.changes,
        }


@dataclass
class DiffReport:
    base_profile: str
    diffs: list[CandidateDiff] = field(default_factory=list)
    uncovered_existing: list[str] = field(default_factory=list)
    unmapped: list[dict] = field(default_factory=list)
    dropped: list[dict] = field(default_factory=list)

    def counts(self) -> dict[str, int]:
        result = {"keep": 0, "modify": 0, "new": 0}
        for diff in self.diffs:
            result[diff.action] = result.get(diff.action, 0) + 1
        return result

    def to_dict(self) -> dict:
        return {
            "base_profile": self.base_profile,
            "counts": self.counts(),
            "uncovered_existing": self.uncovered_existing,
            "diffs": [diff.to_dict() for diff in self.diffs],
            "unmapped": self.unmapped,
            "dropped": self.dropped,
        }


# ---------------------------------------------------------------------------
# 1. Prompt
# ---------------------------------------------------------------------------
def number_spec_lines(text: str, max_chars: int = DEFAULT_SPEC_CHAR_LIMIT) -> str:
    """给规范原文加行号（方便 LLM 引用出处），并限制总长度。"""
    lines = text.splitlines()
    numbered: list[str] = []
    used = 0
    for index, line in enumerate(lines, start=1):
        item = f"L{index}: {line}"
        if used + len(item) + 1 > max_chars:
            numbered.append(f"...（原文过长，已截断至第 {index - 1} 行）")
            break
        numbered.append(item)
        used += len(item) + 1
    return "\n".join(numbered)


def build_messages(
    spec_text: str,
    ruleset: RuleSet,
    *,
    max_spec_chars: int = DEFAULT_SPEC_CHAR_LIMIT,
) -> list[dict]:
    """构造抽取规则候选的 chat messages。"""
    existing_lines = []
    for rule in ruleset.rules.values():
        expected = json.dumps(rule.expected, ensure_ascii=False, separators=(",", ":"))
        existing_lines.append(f"- {rule.id} | {rule.name} | {expected}")
    existing_table = "\n".join(existing_lines)
    spec_numbered = number_spec_lines(spec_text, max_chars=max_spec_chars)

    system = (
        "你是高校毕业论文（设计）格式规范解析器。对照「规范原文」与「现有规则库」，"
        "输出严格 JSON（不要其他文字），包含三部分：existing_assessment / new_rules / unmapped。\n\n"
        "1) existing_assessment：对现有规则库的**每一条规则**给出评估，一条都不能遗漏：\n"
        '   {"id":"PAGE-001","action":"keep"}   // 规范与现有规则一致，或规范未提出不同要求\n'
        '   {"id":"BODY-001","action":"modify","expected":{完整的新期望值},"spec":"规范原文摘要",'
        '"source":"L91","confidence":0.9,"note":"为什么改"}\n'
        "   - modify 必须给出**完整 expected**（在现有 expected 基础上改），不要只写变化字段。\n"
        "2) new_rules：现有规则库没有、但规范明确要求的**可编程**格式规则；每条给完整字段：\n"
        '   {"id":"FIG-003","name":"...","category":"page|body|heading|front|toc|header_footer|'
        'float|references|structure","phase":"P0|P1","severity":"error|warning|info",'
        '"profiles":["*"],"spec":"...","expected":{...},"source":"L92","confidence":0.85,"note":"..."}\n'
        "3) unmapped：无法程序化检查的要求（内容/写作/主观判断/物理装订），格式 "
        '{"text":"...","source":"Lxx","reason":"..."}。\n\n'
        "输出 JSON：{\"existing_assessment\": [...], \"new_rules\": [...], \"unmapped\": [...], "
        "\"summary\": \"一句话总结\"}\n\n"
        "硬性要求：\n"
        "- existing_assessment 必须覆盖现有规则库全部 id（当前见下方列表，一条不漏）；"
        "规范没提到或无需调整的一律 action=keep。\n"
        "- new_rules 的 id 不能与现有 id 重复，也不能重复出现；没有就输出空数组。\n"
        "- expected 必须是可编程 JSON 参数：字号 pt、行距倍数、缩进字符数、对齐枚举"
        "（left/center/right/justify/both）、段前段后行数、页边距 cm；不确定的字段不要编造。\n"
        "- 每条都带 source（规范行号）与 confidence（0.1-1.0）。\n"
        "- 不要输出任何非 JSON 内容。Return a valid JSON object."
    )
    user = (
        "【现有规则库（id | 名称 | expected）】\n"
        f"{existing_table}\n\n"
        "【规范原文（行号前缀仅用于引用出处）】\n"
        f"{spec_numbered}\n\n"
        "请输出 JSON：existing_assessment 覆盖全部现有规则 id，再给 new_rules 与 unmapped。"
    )
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": user},
    ]


# ---------------------------------------------------------------------------
# 2. 解析与本地校验
# ---------------------------------------------------------------------------
def _validate_candidate(data: dict, dropped: list[dict]) -> Optional[Candidate]:
    try:
        candidate = Candidate.from_dict(data)
    except Exception as exc:  # noqa: BLE001 - LLM 数据不可信
        dropped.append({"raw": data, "reason": f"字段解析失败：{exc}"})
        return None

    if not RULE_ID_RE.match(candidate.id):
        dropped.append({"id": candidate.id, "reason": "规则 id 格式不合法（应如 PAGE-001）"})
        return None
    if not candidate.name:
        dropped.append({"id": candidate.id, "reason": "缺少规则名称"})
        return None
    if candidate.category not in KNOWN_CATEGORIES:
        candidate.category = "structure"
        candidate.note = (candidate.note + "；" if candidate.note else "") + "分类无法识别，暂归 structure"
    if candidate.phase not in KNOWN_PHASES:
        candidate.phase = "P1"
    if candidate.severity not in KNOWN_SEVERITIES:
        candidate.severity = "warning"
    if not candidate.expected:
        dropped.append({"id": candidate.id, "reason": "expected 为空，无法程序化检查"})
        return None
    if not 0.0 <= candidate.confidence <= 1.0:
        candidate.confidence = min(1.0, max(0.0, candidate.confidence))
    if not candidate.profiles:
        candidate.profiles = ["*"]
    return candidate


def _candidate_from_rule(rule) -> Candidate:
    """用现有规则构造 keep 候选（用于 LLM 逐条评估的 keep 项）。"""
    return Candidate(
        id=rule.id,
        name=rule.name,
        category=rule.category,
        phase=rule.phase,
        severity=rule.severity,
        profiles=list(rule.profiles),
        spec=rule.spec,
        expected=copy.deepcopy(rule.expected),
        source=rule.source,
        confidence=0.9,
        note="LLM 评估：与现有规则一致",
    )


def _assessment_to_candidate(item: dict, ruleset: RuleSet, dropped: list[dict]) -> Optional[Candidate]:
    """把 existing_assessment 的一项转成 Candidate（keep 用现有规则填充）。"""
    rule_id = str(item.get("id", "")).strip()
    action = str(item.get("action", "keep")).strip().lower()
    if not rule_id:
        dropped.append({"raw": item, "reason": "existing_assessment 缺 id"})
        return None
    if action not in ("keep", "modify"):
        dropped.append({"id": rule_id, "reason": f"未知 action：{action}"})
        return None
    rule = ruleset.rules.get(rule_id)
    if rule is None:
        dropped.append({"id": rule_id, "reason": "评估了规则库中不存在的 id"})
        return None

    if action == "keep":
        candidate = _candidate_from_rule(rule)
        candidate.note = str(item.get("note") or candidate.note)
        confidence = item.get("confidence")
        if isinstance(confidence, (int, float)) and 0.0 <= float(confidence) <= 1.0:
            candidate.confidence = float(confidence)
        return candidate

    # modify：以现有规则为底，用 LLM 给出的字段覆盖；expected 必须是完整值
    data = _rule_to_dict(rule)
    for field_name in ("name", "category", "phase", "severity", "profiles", "spec",
                       "source", "confidence", "note"):
        value = item.get(field_name)
        if value not in (None, ""):
            data[field_name] = value
    expected = item.get("expected")
    if not isinstance(expected, dict) or not expected:
        dropped.append({"id": rule_id, "reason": "modify 未给完整 expected，已跳过"})
        return None
    data["expected"] = expected
    candidate = _validate_candidate(data, dropped)
    if candidate is not None and not candidate.note:
        candidate.note = "LLM 评估：建议修改"
    return candidate


def _collect_unmapped(payload: dict) -> list[dict]:
    unmapped: list[dict] = []
    for item in payload.get("unmapped") or []:
        if isinstance(item, dict):
            unmapped.append(
                {
                    "text": str(item.get("text", ""))[:500],
                    "source": str(item.get("source", "")),
                    "reason": str(item.get("reason", "")),
                }
            )
    return unmapped


def parse_llm_payload(
    payload: dict,
    ruleset: Optional[RuleSet] = None,
) -> tuple[list[Candidate], list[dict], list[dict], str]:
    """校验 LLM 返回 JSON，返回 (candidates, unmapped, dropped, summary)。

    支持两种返回格式：
    - 新格式：{"existing_assessment": [...], "new_rules": [...], "unmapped": [...], "summary": "..."}
    - 旧格式（向后兼容）：{"rules": [...], "unmapped": [...], "summary": "..."}
    """
    if not isinstance(payload, dict):
        raise ValueError("LLM 返回不是 JSON 对象")

    dropped: list[dict] = []
    candidates: list[Candidate] = []
    seen: set[str] = set()
    unmapped = _collect_unmapped(payload)
    summary = str(payload.get("summary", ""))

    assessment = payload.get("existing_assessment")
    new_rules = payload.get("new_rules")
    if isinstance(assessment, list) or isinstance(new_rules, list):
        if ruleset is None:
            raise ValueError("解析 existing_assessment 需要传入 ruleset（现有规则库）")
        for item in assessment or []:
            if not isinstance(item, dict):
                dropped.append({"raw": item, "reason": "existing_assessment 元素不是对象"})
                continue
            candidate = _assessment_to_candidate(item, ruleset, dropped)
            if candidate is None:
                continue
            if candidate.id in seen:
                dropped.append({"id": candidate.id, "reason": "评估重复，保留第一条"})
                continue
            seen.add(candidate.id)
            candidates.append(candidate)
        for item in new_rules or []:
            if not isinstance(item, dict):
                dropped.append({"raw": item, "reason": "new_rules 元素不是对象"})
                continue
            candidate = _validate_candidate(item, dropped)
            if candidate is None:
                continue
            if candidate.id in ruleset.rules:
                dropped.append({"id": candidate.id, "reason": "new_rules 与现有规则 id 冲突（应放 existing_assessment）"})
                continue
            if candidate.id in seen:
                dropped.append({"id": candidate.id, "reason": "LLM 返回了重复 id"})
                continue
            seen.add(candidate.id)
            candidates.append(candidate)
        return candidates, unmapped, dropped, summary

    # 旧格式：rules 数组
    raw_rules = payload.get("rules")
    if not isinstance(raw_rules, list):
        raise ValueError("LLM 返回缺少 existing_assessment / new_rules / rules 数组")
    for item in raw_rules:
        if not isinstance(item, dict):
            dropped.append({"raw": item, "reason": "rules 数组元素不是对象"})
            continue
        candidate = _validate_candidate(item, dropped)
        if candidate is None:
            continue
        if candidate.id in seen:
            dropped.append({"id": candidate.id, "reason": "LLM 返回了重复 id，保留第一条"})
            continue
        seen.add(candidate.id)
        candidates.append(candidate)
    return candidates, unmapped, dropped, summary


def extract_candidates(
    spec_text: str,
    ruleset: RuleSet,
    *,
    chat: Optional[Callable[[list[dict]], dict]] = None,
    model: str = "",
    max_spec_chars: int = DEFAULT_SPEC_CHAR_LIMIT,
    base_profile: str = DEFAULT_PROFILE_ID,
) -> ExtractionReport:
    """调用 LLM 抽取规则候选。chat 可注入（测试用）；默认使用 tas.llm.chat_json。"""
    if chat is None:
        from .llm import chat_json  # 延迟导入：无 key 时也能做离线测试

        chat = lambda messages: chat_json(messages)  # noqa: E731

    messages = build_messages(spec_text, ruleset, max_spec_chars=max_spec_chars)
    payload = chat(messages)
    candidates, unmapped, dropped, summary = parse_llm_payload(payload, ruleset)
    return ExtractionReport(
        base_profile=base_profile,
        base_rules_count=len(ruleset),
        candidates=candidates,
        unmapped=unmapped,
        dropped=dropped,
        summary=summary,
        model=model,
        spec_chars=len(spec_text),
    )


# ---------------------------------------------------------------------------
# 3. 与现有规则 diff
# ---------------------------------------------------------------------------
def _norm(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def _changes(existing: dict, candidate: Candidate) -> list[str]:
    changes: list[str] = []
    for field_name in ("name", "category", "phase", "severity"):
        old = existing.get(field_name)
        new = getattr(candidate, field_name)
        if _norm(old) != _norm(new):
            changes.append(f"{field_name}: {old!r} → {new!r}")
    old_expected = existing.get("expected") or {}
    new_expected = candidate.expected or {}
    for key in sorted(set(old_expected) | set(new_expected)):
        old_value = old_expected.get(key)
        new_value = new_expected.get(key)
        if _norm(old_value) != _norm(new_value):
            changes.append(f"expected.{key}: {old_value!r} → {new_value!r}")
    old_profiles = list(existing.get("profiles") or [])
    if _norm(old_profiles) != _norm(candidate.profiles):
        changes.append(f"profiles: {old_profiles} → {candidate.profiles}")
    return changes


def diff_candidates(
    report: ExtractionReport,
    ruleset: RuleSet,
    *,
    base_profile: str = DEFAULT_PROFILE_ID,
) -> DiffReport:
    """候选 vs 现有规则：keep / modify / new，并列出未被候选覆盖的现有规则。"""
    existing = {
        rule.id: {
            "id": rule.id,
            "name": rule.name,
            "category": rule.category,
            "phase": rule.phase,
            "severity": rule.severity,
            "profiles": list(rule.profiles),
            "spec": rule.spec,
            "expected": rule.expected,
            "source": rule.source,
        }
        for rule in ruleset.rules.values()
    }
    diffs: list[CandidateDiff] = []
    touched: set[str] = set()
    for candidate in report.candidates:
        old = existing.get(candidate.id)
        if old is None:
            diffs.append(CandidateDiff(action="new", candidate=candidate, existing_id=None))
            continue
        touched.add(candidate.id)
        changes = _changes(old, candidate)
        diffs.append(
            CandidateDiff(
                action="modify" if changes else "keep",
                candidate=candidate,
                existing_id=candidate.id,
                changes=changes,
            )
        )
    uncovered = sorted(set(existing) - touched)
    return DiffReport(
        base_profile=base_profile,
        diffs=diffs,
        uncovered_existing=uncovered,
        unmapped=report.unmapped,
        dropped=report.dropped,
    )


# ---------------------------------------------------------------------------
# 4. 生成新 profile
# ---------------------------------------------------------------------------
def _rule_to_dict(rule) -> dict:
    return {
        "id": rule.id,
        "name": rule.name,
        "category": rule.category,
        "phase": rule.phase,
        "severity": rule.severity,
        "profiles": list(rule.profiles),
        "spec": rule.spec,
        "check": rule.check,
        "fix": rule.fix,
        "source": rule.source,
        "expected": rule.expected,
    }


def build_profile_payload(
    ruleset: RuleSet,
    diff: DiffReport,
    *,
    profile_id: str,
    name: str,
    discipline: str,
    source_spec: str = "",
) -> dict:
    """把 candidate diff 合并进现有规则库，生成新 profile 的 rules.json payload。"""
    rules: list[dict] = []
    by_id: dict[str, dict] = {}
    for rule in ruleset.rules.values():
        item = _rule_to_dict(rule)
        rules.append(item)
        by_id[rule.id] = item

    for candidate_diff in diff.diffs:
        candidate = candidate_diff.candidate
        if candidate_diff.action == "keep":
            continue
        if candidate_diff.action == "modify":
            target = by_id.get(candidate.id)
            if target is not None:
                target.update(
                    {
                        "name": candidate.name,
                        "category": candidate.category,
                        "phase": candidate.phase,
                        "severity": candidate.severity,
                        "profiles": candidate.profiles,
                        "spec": candidate.spec or target.get("spec", ""),
                        "expected": candidate.expected,
                        "source": candidate.source or target.get("source", ""),
                    }
                )
        elif candidate_diff.action == "new":
            item = {
                "id": candidate.id,
                "name": candidate.name,
                "category": candidate.category,
                "phase": candidate.phase,
                "severity": candidate.severity,
                "profiles": candidate.profiles,
                "spec": candidate.spec,
                "check": "LLM 抽取的规则；检查器待实现（未绑定 checker 前不会参与检查）",
                "fix": "待定",
                "source": candidate.source,
                "expected": candidate.expected,
            }
            rules.append(item)
            by_id[candidate.id] = item

    meta = copy.deepcopy(ruleset.meta)
    meta.update(
        {
            "id": profile_id,
            "name": name,
            "discipline": discipline,
            "version": "0.1.0-draft",
            "derived_from": diff.base_profile,
            "generated_by": f"tas {__version__} spec parse",
            "generated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
            "source_spec": source_spec,
        }
    )
    return {"meta": meta, "rules": rules}


def write_profile_payload(payload: dict, out_path: Path | str) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    # 立即用 RuleSet 加载一次，确保写出的规则库可用于检查
    RuleSet.load(out_path)
    return out_path


def register_profile(
    index_path: Path | str,
    *,
    profile_id: str,
    name: str,
    discipline: str,
    rules_path: Path,
    source: str = "",
    notes: str = "",
) -> Path:
    """把新 profile 写入 profiles/index.json（相对路径，便于迁移）。"""
    index_path = Path(index_path)
    index_path.parent.mkdir(parents=True, exist_ok=True)
    if index_path.is_file():
        data = json.loads(index_path.read_text(encoding="utf-8"))
    else:
        data = {"version": 1, "profiles": []}
    try:
        relative = rules_path.resolve().relative_to(index_path.parent.resolve()).as_posix()
    except ValueError:
        relative = rules_path.resolve().as_posix()
    entry = {
        "id": profile_id,
        "name": name,
        "discipline": discipline,
        "rules": relative,
        "source": source,
        "notes": notes or "由 tas spec parse 生成；教师确认后启用",
    }
    entries = [item for item in data.get("profiles", []) if item.get("id") != profile_id]
    entries.append(entry)
    data["profiles"] = entries
    index_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return index_path


def apply_candidates_file(
    candidates_path: Path | str,
    *,
    profile_id: str,
    name: str,
    discipline: str,
    out_path: Path | str,
    index_path: Optional[Path | str] = None,
    source_spec: str = "",
    notes: str = "",
) -> dict:
    """读取 tas spec parse 产出的 candidates.json，生成并（可选）注册新 profile。"""
    data = json.loads(Path(candidates_path).read_text(encoding="utf-8"))
    extraction = data.get("extraction") or {}
    raw_candidates = extraction.get("candidates") or data.get("candidates") or []
    candidates = [Candidate.from_dict(item) for item in raw_candidates if isinstance(item, dict)]
    base_profile = data.get("base_profile") or extraction.get("base_profile") or DEFAULT_PROFILE_ID
    _profile, ruleset = resolve_profile(base_profile)

    report = ExtractionReport(
        base_profile=base_profile,
        base_rules_count=len(ruleset),
        candidates=candidates,
        unmapped=list(extraction.get("unmapped") or []),
        dropped=list(extraction.get("dropped") or []),
    )
    diff = diff_candidates(report, ruleset, base_profile=base_profile)
    payload = build_profile_payload(
        ruleset,
        diff,
        profile_id=profile_id,
        name=name,
        discipline=discipline,
        source_spec=source_spec,
    )
    rules_path = write_profile_payload(payload, out_path)
    registered = False
    if index_path is not None:
        register_profile(
            index_path,
            profile_id=profile_id,
            name=name,
            discipline=discipline,
            rules_path=rules_path,
            source=source_spec,
            notes=notes,
        )
        registered = True
    counts = diff.counts()
    return {
        "profile_id": profile_id,
        "rules_path": str(rules_path),
        "registered": registered,
        "rules_total": len(payload["rules"]),
        "candidates": len(candidates),
        "keep": counts.get("keep", 0),
        "modify": counts.get("modify", 0),
        "new": counts.get("new", 0),
        "uncovered_existing": diff.uncovered_existing,
        "unmapped": len(diff.unmapped),
    }


# ---------------------------------------------------------------------------
# 5. 报告渲染
# ---------------------------------------------------------------------------
def render_report_markdown(report: ExtractionReport, diff: DiffReport) -> str:
    counts = diff.counts()
    assessed = sum(1 for item in diff.diffs if item.existing_id is not None)
    lines = [
        "# 规范 → 规则候选报告（教师端）",
        "",
        f"- 基础 profile：`{diff.base_profile}`（现有 {report.base_rules_count} 条规则）",
        f"- 现有规则逐条评估覆盖：**{assessed}/{report.base_rules_count}**",
        f"- 规范字符数：{report.spec_chars}",
        f"- 抽取模型：`{report.model or '（未记录）'}`",
        f"- 候选结果：**keep {counts.get('keep', 0)} ｜ modify {counts.get('modify', 0)} ｜ "
        f"new {counts.get('new', 0)}**；未覆盖现有规则 {len(diff.uncovered_existing)} 条；"
        f"unmapped {len(diff.unmapped)} 条；丢弃 {len(diff.dropped)} 条",
        "",
    ]
    if report.summary:
        lines += [f"> LLM 总结：{report.summary}", ""]

    def _section(title: str, items: list[CandidateDiff], show_changes: bool) -> None:
        lines.append(f"## {title}（{len(items)}）")
        lines.append("")
        for item in items:
            candidate = item.candidate
            lines.append(
                f"### `{candidate.id}` {candidate.name}  "
                f"（{candidate.phase}/{candidate.severity}，置信度 {candidate.confidence:.2f}）"
            )
            if candidate.source:
                lines.append(f"- 出处：{candidate.source}")
            lines.append(f"- 规范：{candidate.spec}")
            lines.append(f"- expected：`{json.dumps(candidate.expected, ensure_ascii=False)}`")
            if candidate.note:
                lines.append(f"- 备注：{candidate.note}")
            if show_changes and item.changes:
                lines.append("- 与现有规则差异：")
                for change in item.changes:
                    lines.append(f"  - {change}")
            lines.append("")

    _section("新增规则（new）", [d for d in diff.diffs if d.action == "new"], False)
    _section("修改现有规则（modify）", [d for d in diff.diffs if d.action == "modify"], True)
    _section("保持不变（keep）", [d for d in diff.diffs if d.action == "keep"], False)

    if diff.uncovered_existing:
        lines += ["## 未被本次抽取覆盖的现有规则", ""]
        lines += [f"- `{rule_id}`" for rule_id in diff.uncovered_existing]
        lines.append("")
    if diff.unmapped:
        lines += ["## unmapped：无法程序化检查的规范要求（需人工/流程处理）", ""]
        for item in diff.unmapped:
            lines.append(f"- {item.get('source', '')} {item.get('text', '')}（{item.get('reason', '')}）")
        lines.append("")
    if diff.dropped:
        lines += ["## 丢弃的候选（结构不合法/无法检查）", ""]
        for item in diff.dropped:
            lines.append(f"- {item}")
        lines.append("")
    lines += [
        "---",
        "审阅方式：先看 `diff.md`，修改/删除 `candidates.json` 里不认可的候选，",
        "再执行 `tas spec apply candidates.json --id <profile_id> --name <名称>` 生成新 profile。",
        "注意：new 规则在绑定检查器之前不会参与检查（只进入规则库，待实现）。",
    ]
    return "\n".join(lines)
