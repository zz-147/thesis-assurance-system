"""文档接龙：发起接龙、学生提交、一键审核（本地 JSON 存储，无账号）。

存储：
    data/collections.json                 # 接龙与提交记录（可用 TAS_DATA_DIR 覆盖 data 目录）
    data/collections/<code>/<sid>.docx    # 提交的文档原件

设计：
- 不拆教师端/学生端：任何人可发起接龙，同学凭 6 位 code（分享链接/小程序卡片）加入并提交；
  任何人都可以点「一键审核」，后端批量跑格式检查并返回逐份结果 + 班级汇总。
- 无账号：提交时只填昵称（小程序端本地缓存）。
- 单机/内网优先；上线微信小程序需要 HTTPS 域名（见 docs/MINIPROGRAM.md）。
- 每次提交/审核都复用 tas 的 DocModel + roles + run_checks，与 CLI/Web 检查结果一致。
"""

from __future__ import annotations

import json
import os
import re
import secrets
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from .checkers import IMPLEMENTED_RULES, run_checks
from .config import DEFAULT_PROFILE_ID, RuleSet, project_root, resolve_profile
from .docx_io import DocModel
from .report import Report
from .roles import classify

CODE_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # 去掉易混淆的 I/O/0/1
CODE_RE = re.compile(r"^[A-Z0-9]{6}$")
MAX_FILE_BYTES = 30 * 1024 * 1024
MAX_TITLE = 80
MAX_NOTE = 200
MAX_NICKNAME = 20

_LOCK = threading.Lock()


# ---------------------------------------------------------------------------
# 存储
# ---------------------------------------------------------------------------
def data_dir() -> Path:
    override = os.environ.get("TAS_DATA_DIR")
    root = Path(override) if override else project_root() / "data"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _store_path() -> Path:
    return data_dir() / "collections.json"


def _uploads_dir(code: str) -> Path:
    directory = data_dir() / "collections" / code
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _load() -> dict:
    path = _store_path()
    if not path.is_file():
        return {"version": 1, "collections": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {"version": 1, "collections": []}
    if not isinstance(data, dict) or not isinstance(data.get("collections"), list):
        return {"version": 1, "collections": []}
    return data


def _save(data: dict) -> None:
    path = _store_path()
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(path)


def _new_code(existing: set[str]) -> str:
    for _ in range(100):
        code = "".join(secrets.choice(CODE_ALPHABET) for _ in range(6))
        if code not in existing:
            return code
    raise RuntimeError("无法生成接龙 code，请重试")


def _check_code(code: str) -> str:
    code = (code or "").strip().upper()
    if not CODE_RE.match(code):
        raise ValueError("接龙 code 不合法（应为 6 位大写字母/数字）")
    return code


def _find(data: dict, code: str) -> dict:
    for item in data["collections"]:
        if item.get("code") == code:
            return item
    raise ValueError(f"接龙不存在：{code}")


def _status_from_counts(counts: Optional[dict], error: Optional[str]) -> str:
    if error:
        return "检查失败"
    if not counts:
        return "未检查"
    return "通过" if counts.get("error", 0) == 0 else "未通过"


def _public_collection(item: dict) -> dict:
    return {
        "code": item["code"],
        "title": item["title"],
        "owner": item.get("owner", ""),
        "ruleset": item.get("ruleset", DEFAULT_PROFILE_ID),
        "discipline": item.get("discipline", "sci_eng"),
        "note": item.get("note", ""),
        "due": item.get("due", ""),
        "created_at": item.get("created_at", ""),
        "submission_count": len(item.get("submissions", [])),
    }


def _public_submission(submission: dict) -> dict:
    return {
        "id": submission["id"],
        "student": submission.get("student", ""),
        "filename": submission.get("filename", ""),
        "created_at": submission.get("created_at", ""),
        "counts": submission.get("counts"),
        "status": submission.get("status", ""),
        "error": submission.get("error"),
    }


# ---------------------------------------------------------------------------
# 检查
# ---------------------------------------------------------------------------
def _check_docx(path: Path, ruleset_id: str) -> tuple[Optional[Report], Optional[str]]:
    try:
        profile, ruleset = resolve_profile(ruleset_id or DEFAULT_PROFILE_ID)
        model = DocModel.load(path)
        issues = run_checks(
            model, classify(model, profile.discipline), ruleset, profile=profile.discipline
        )
        report = Report(
            path=path,
            profile=profile.discipline,
            issues=issues,
            stats=model.stats(),
            rules_run=len([rid for rid in IMPLEMENTED_RULES if rid in ruleset.rules]),
            rules_total=len(ruleset),
        )
        return report, None
    except Exception as exc:  # noqa: BLE001 - 单份文档失败不拖垮整批
        return None, str(exc)


def _counts(report: Optional[Report]) -> Optional[dict]:
    if report is None:
        return None
    return report.counts()


# ---------------------------------------------------------------------------
# 业务：接龙
# ---------------------------------------------------------------------------
def create_collection(
    title: str,
    owner: str = "",
    ruleset: str = DEFAULT_PROFILE_ID,
    note: str = "",
    due: str = "",
) -> dict:
    title = (title or "").strip()
    if not title:
        raise ValueError("请填写接龙标题")
    if len(title) > MAX_TITLE:
        raise ValueError(f"标题过长（≤{MAX_TITLE} 字）")
    profile, _ = resolve_profile(ruleset or DEFAULT_PROFILE_ID)
    owner = (owner or "").strip()[:MAX_NICKNAME] or "匿名同学"
    note = (note or "").strip()[:MAX_NOTE]
    due = (due or "").strip()[:30]
    with _LOCK:
        data = _load()
        existing = {item.get("code") for item in data["collections"]}
        code = _new_code(existing)
        item = {
            "code": code,
            "title": title,
            "owner": owner,
            "ruleset": profile.id,
            "discipline": profile.discipline,
            "note": note,
            "due": due,
            "created_at": _now(),
            "submissions": [],
        }
        data["collections"].append(item)
        _save(data)
    return _public_collection(item)


def list_collections(limit: int = 50) -> list[dict]:
    data = _load()
    items = sorted(
        data["collections"], key=lambda item: item.get("created_at", ""), reverse=True
    )
    return [_public_collection(item) for item in items[: max(1, int(limit))]]


def get_collection(code: str) -> dict:
    code = _check_code(code)
    data = _load()
    item = _find(data, code)
    return {
        "collection": _public_collection(item),
        "submissions": [_public_submission(sub) for sub in item.get("submissions", [])],
    }


def add_submission(code: str, student: str, filename: str, payload: bytes) -> dict:
    code = _check_code(code)
    student = (student or "").strip()[:MAX_NICKNAME] or "匿名同学"
    original_name = Path(filename or "document.docx").name
    original_name = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", original_name)[:120] or "document.docx"
    if not original_name.lower().endswith(".docx"):
        raise ValueError("只支持 .docx（.doc 请先另存为 .docx）")
    if not payload:
        raise ValueError("提交的文件为空")
    if len(payload) > MAX_FILE_BYTES:
        raise ValueError(f"文件过大（> {MAX_FILE_BYTES // 1024 // 1024}MB）")

    with _LOCK:
        data = _load()
        item = _find(data, code)

    submission_id = uuid.uuid4().hex[:12]
    directory = _uploads_dir(code)
    target = directory / f"{submission_id}.docx"
    target.write_bytes(payload)
    report, error = _check_docx(target, item.get("ruleset", DEFAULT_PROFILE_ID))
    submission = {
        "id": submission_id,
        "student": student,
        "filename": original_name,
        "path": str(target.relative_to(data_dir()).as_posix()),
        "created_at": _now(),
        "counts": _counts(report),
        "status": _status_from_counts(_counts(report), error),
        "error": error,
    }
    with _LOCK:
        data = _load()
        item = _find(data, code)
        item.setdefault("submissions", []).append(submission)
        _save(data)
    return _public_submission(submission)


def review_collection(code: str) -> dict:
    """一键审核：逐份重新检查（防止提交后文件被替换/规则更新），返回逐份结果 + 班级汇总。"""
    code = _check_code(code)
    with _LOCK:
        data = _load()
        item = _find(data, code)
        ruleset_id = item.get("ruleset", DEFAULT_PROFILE_ID)
        submissions = list(item.get("submissions", []))

    results = []
    counts_total = {"error": 0, "warning": 0, "info": 0}
    rule_hits: dict[str, int] = {}
    checked = failed = 0
    for submission in submissions:
        path = Path(data_dir() / submission.get("path", ""))
        report, error = _check_docx(path, ruleset_id) if path.is_file() else (None, "文件不存在")
        counts = _counts(report)
        submission["counts"] = counts
        submission["status"] = _status_from_counts(counts, error)
        submission["error"] = error
        if error:
            failed += 1
        else:
            checked += 1
            for key in counts_total:
                counts_total[key] += int((counts or {}).get(key, 0))
            if report is not None:
                for issue in report.issues:
                    rule_hits[issue.rule_id] = rule_hits.get(issue.rule_id, 0) + 1
        results.append(_public_submission(submission))

    with _LOCK:
        data = _load()
        item = _find(data, code)
        item["submissions"] = submissions
        _save(data)

    pass_count = sum(1 for item in results if item["status"] == "通过")
    denominator = checked or 1
    summary = {
        "total": len(results),
        "checked": checked,
        "failed": failed,
        "pass_rate": round(pass_count / denominator, 4),
        "counts": counts_total,
        "top_rules": sorted(rule_hits.items(), key=lambda pair: pair[1], reverse=True)[:10],
    }
    return {"summary": summary, "submissions": results}


def get_submission_report(code: str, submission_id: str) -> dict:
    code = _check_code(code)
    with _LOCK:
        data = _load()
        item = _find(data, code)
        submission = next(
            (sub for sub in item.get("submissions", []) if sub.get("id") == submission_id),
            None,
        )
    if submission is None:
        raise ValueError("提交记录不存在")
    path = Path(data_dir() / submission.get("path", ""))
    if not path.is_file():
        raise FileNotFoundError("提交文件不存在")
    report, error = _check_docx(path, item.get("ruleset", DEFAULT_PROFILE_ID))
    if error or report is None:
        raise RuntimeError(f"检查失败：{error}")
    payload = report.to_dict()
    payload["submission"] = _public_submission(submission)
    return payload
