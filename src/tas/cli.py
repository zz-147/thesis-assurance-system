"""命令行入口：tas check / fix / refs / rules / profiles / spec。"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any, Optional

from . import __version__
from .checkers import IMPLEMENTED_RULES, run_checks
from .config import (
    DEFAULT_PROFILE,
    DEFAULT_PROFILE_ID,
    PROFILES,
    RuleSet,
    list_profiles,
    profiles_index_path,
    project_root,
    resolve_profile,
)
from .docx_io import DocModel
from .report import Report
from .roles import classify
from .spec_io import extract_spec_text, spec_stats


def _ensure_utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:  # pragma: no cover - 老环境兜底
            pass


# ---------------------------------------------------------------------------
# 参数解析
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tas", description="毕业论文（设计）格式统一工具：检查 / 修复 / 真实性核验"
    )
    parser.add_argument("--version", action="version", version=f"tas {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    # check -----------------------------------------------------------------
    check = sub.add_parser("check", help="检查一个 .docx 是否符合格式规范")
    check.add_argument("docx", help="待检查的 .docx 文件")
    check.add_argument("--profile", default=None, choices=sorted(PROFILES),
                       help="学科规范体系（默认取 profile 的学科，unset 时 sci_eng）")
    check.add_argument("--ruleset", default=None,
                       help="命名规则集 id 或规则 JSON 路径（默认 profiles/index.json 的 school_default）")
    check.add_argument("--rules", default=None, help="直接指定规则库路径（兼容旧用法）")
    check.add_argument("--phase", default="all", choices=("all", "P0", "P1"), help="只报告指定阶段的规则")
    check.add_argument("--format", default="text", choices=("text", "json", "md"), help="报告格式")
    check.add_argument("--out", default=None, help="报告输出文件（默认打印到 stdout）")
    check.add_argument("--max-issues", type=int, default=80, help="text/md 报告最多显示多少条")

    # fix -------------------------------------------------------------------
    fix = sub.add_parser("fix", help="安全修复格式问题（学生端；只改格式、不改内容、不动文献）")
    fix.add_argument("docx", help="待修复的 .docx 文件")
    fix.add_argument("--out", default=None, help="修复后的 docx 路径（默认 原名_fixed.docx）")
    fix.add_argument("--dry-run", action="store_true", help="只计算要改什么，不写文件")
    fix.add_argument("--report", default=None, help="修复报告输出文件（.json/.md/.txt）")
    fix.add_argument("--format", default="text", choices=("text", "json"), help="报告格式")
    fix.add_argument("--profile", default=None, choices=sorted(PROFILES), help="学科规范体系")
    fix.add_argument("--ruleset", default=None, help="命名规则集 id 或规则 JSON 路径")
    fix.add_argument("--rules", default=None, help="直接指定规则库路径（兼容旧用法）")

    # refs ------------------------------------------------------------------
    refs = sub.add_parser("refs", help="参考文献解析 / 真实性核验 / 引用一致性（教师端；只读）")
    refs.add_argument("docx", help="待核验的 .docx 文件")
    refs.add_argument("--verify", action="store_true",
                      help="联网核验文献真实性（Crossref/OpenAlex/arXiv；默认只做离线解析与引用一致性）")
    refs.add_argument("--limit", type=int, default=None, help="最多联网核验多少条（默认全部）")
    refs.add_argument("--timeout", type=float, default=10.0, help="单次请求超时秒数")
    refs.add_argument("--profile", default=None, choices=sorted(PROFILES), help="学科规范体系（用于角色识别）")
    refs.add_argument("--format", default="text", choices=("text", "json"), help="报告格式")
    refs.add_argument("--out", default=None, help="报告输出文件（默认打印）")

    # batch -----------------------------------------------------------------
    batch = sub.add_parser("batch", help="教师端：批量检查 + 班级汇总 + 导出 xlsx/csv")
    batch.add_argument("paths", nargs="+", help="若干 .docx 文件，或一个目录")
    batch.add_argument("--ruleset", default=None, help="命名规则集 id 或规则 JSON 路径")
    batch.add_argument("--profile", default=None, choices=sorted(PROFILES), help="学科规范体系")
    batch.add_argument("--rules", default=None, help="直接指定规则库路径（兼容旧用法）")
    batch.add_argument("--out-dir", default="eval/batch", help="报告输出目录（默认 eval/batch）")
    batch.add_argument("--format", default="text", choices=("text", "json"), help="终端输出格式")
    batch.add_argument("--fix", action="store_true", help="同时生成全部安全修复版 *_fixed.docx（只改格式）")
    batch.add_argument("--fixed-dir", default=None, help="修复版输出目录（默认 <out-dir>/fixed）")
    batch.add_argument("--roster", default=None, help="名单文件（.txt/.csv/.xlsx）：核对缺交/重复/命名")
    batch.add_argument("--preview", action="store_true", help="生成每份文档首页缩略图（需 Windows + Word，较慢）")
    batch.add_argument("--preview-width", type=int, default=900, help="缩略图宽度像素（默认 900）")
    batch.add_argument("--refs", action="store_true", help="批量核验参考文献真实性（联网，较慢）")
    batch.add_argument("--refs-limit", type=int, default=5, help="每份文档最多核验几条参考文献（默认 5）")
    batch.add_argument("--refs-timeout", type=float, default=10.0, help="单次文献查询超时秒数（默认 10）")

    # preview ---------------------------------------------------------------
    preview = sub.add_parser("preview", help="把 docx 前几页渲染为 PNG（Word COM 预览）")
    preview.add_argument("docx", help="待预览的 .docx 文件")
    preview.add_argument("--out-dir", default="eval/preview", help="输出目录（默认 eval/preview）")
    preview.add_argument("--pages", type=int, default=2, help="渲染前几页（默认 2）")
    preview.add_argument("--width", type=int, default=1400, help="输出图片宽度像素（默认 1400）")
    preview.add_argument("--format", default="text", choices=("text", "json"), help="输出格式")

    # web -------------------------------------------------------------------
    web = sub.add_parser("web", help="启动 Web 双入口（教师端/学生端，仅本机）")
    web.add_argument("--host", default="127.0.0.1", help="监听地址（默认仅本机；真机联调用 0.0.0.0）")
    web.add_argument("--port", type=int, default=8765, help="端口（默认 8765）")
    web.add_argument("--verbose", action="store_true", help="打印每个请求（排查手机/真机网络问题）")
    web.add_argument("--open-browser", action="store_true", help="启动后自动打开浏览器（exe/老师端用）")

    # rules -----------------------------------------------------------------
    rules = sub.add_parser("rules", help="查看规则库")
    rules.add_argument("--ruleset", default=None, help="命名规则集 id 或规则 JSON 路径")
    rules.add_argument("--rules", default=None, help="直接指定规则库路径（兼容旧用法）")
    rules.add_argument("--phase", default="all", choices=("all", "P0", "P1"))
    rules.add_argument("--json", action="store_true", help="以 JSON 输出")

    # profiles --------------------------------------------------------------
    profiles = sub.add_parser("profiles", help="查看/管理多学科规则集（profile）")
    profiles_sub = profiles.add_subparsers(dest="profiles_command", required=True)
    profiles_list = profiles_sub.add_parser("list", help="列出 profiles/index.json 中的规则集")
    profiles_list.add_argument("--json", action="store_true", help="以 JSON 输出")

    # spec ------------------------------------------------------------------
    spec = sub.add_parser("spec", help="规范文档工具（教师端）")
    spec_sub = spec.add_subparsers(dest="spec_command", required=True)
    extract = spec_sub.add_parser("extract", help="提取规范文档纯文本")
    extract.add_argument("file", help="规范文档（.docx / .txt / .md）")
    extract.add_argument("--out", default=None, help="输出文本文件（默认打印）")

    parse = spec_sub.add_parser("parse", help="调用 LLM 从规范文档抽取规则候选（教师端）")
    parse.add_argument("file", help="规范文档（.docx / .txt / .md）")
    parse.add_argument("--base-profile", default=DEFAULT_PROFILE_ID,
                       help="以哪个现有 profile 为基线做 diff（默认 school_default）")
    parse.add_argument("--out-dir", default="eval/spec_parse", help="候选/报告输出目录")
    parse.add_argument("--model", default=None, help="覆盖模型名（默认读 TAS_LLM_MODEL 或 deepseek-chat）")
    parse.add_argument("--max-spec-chars", type=int, default=24000, help="发给模型的规范文本上限")
    parse.add_argument("--json", action="store_true", help="输出 JSON 摘要")

    apply_cmd = spec_sub.add_parser("apply", help="把候选规则合并成新 profile（教师确认后执行）")
    apply_cmd.add_argument("candidates", help="tas spec parse 产出的 candidates.json")
    apply_cmd.add_argument("--id", dest="profile_id", required=True, help="新 profile id，如 course_se_2025")
    apply_cmd.add_argument("--name", required=True, help="新 profile 名称")
    apply_cmd.add_argument("--discipline", default=DEFAULT_PROFILE, choices=sorted(PROFILES),
                           help="学科体系（sci_eng / humanities）")
    apply_cmd.add_argument("--out", default=None, help="规则文件路径（默认 profiles/<id>.rules.json）")
    apply_cmd.add_argument("--no-register", action="store_true", help="只写规则文件，不更新 profiles/index.json")
    apply_cmd.add_argument("--source", default="", help="规范来源说明（写入 index）")
    return parser


# ---------------------------------------------------------------------------
def main(argv: Optional[list[str]] = None) -> int:
    _ensure_utf8()
    args = build_parser().parse_args(argv)
    try:
        if args.command == "check":
            return _cmd_check(args)
        if args.command == "fix":
            return _cmd_fix(args)
        if args.command == "refs":
            return _cmd_refs(args)
        if args.command == "batch":
            return _cmd_batch(args)
        if args.command == "preview":
            return _cmd_preview(args)
        if args.command == "web":
            return _cmd_web(args)
        if args.command == "rules":
            return _cmd_rules(args)
        if args.command == "profiles":
            return _cmd_profiles(args)
        if args.command == "spec":
            return _cmd_spec(args)
    except (FileNotFoundError, ValueError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    except RuntimeError as exc:  # 修复安全校验失败（hash 不一致/源文件被改）
        print(f"安全检查失败：{exc}", file=sys.stderr)
        return 3
    return 2


# ---------------------------------------------------------------------------
# 规则集/学科解析
# ---------------------------------------------------------------------------
def _resolve_ruleset(args: argparse.Namespace) -> tuple[str, RuleSet, str]:
    """返回 (profile_id, 规则集, 学科体系)。"""
    explicit_profile = getattr(args, "profile", None)
    explicit_rules = getattr(args, "rules", None)
    if explicit_rules:
        path = Path(explicit_rules)
        if not path.is_file():
            raise FileNotFoundError(f"规则文件不存在：{path}")
        return path.stem, RuleSet.load(path), (explicit_profile or DEFAULT_PROFILE)
    profile, ruleset = resolve_profile(getattr(args, "ruleset", None))
    return profile.id, ruleset, (explicit_profile or profile.discipline)


# ---------------------------------------------------------------------------
# check
# ---------------------------------------------------------------------------
def _cmd_check(args: argparse.Namespace) -> int:
    profile_id, ruleset, discipline = _resolve_ruleset(args)
    model = DocModel.load(Path(args.docx))
    roles = classify(model, discipline)
    issues = run_checks(model, roles, ruleset, profile=discipline, phase=args.phase)
    report = Report(
        path=Path(args.docx),
        profile=discipline,
        issues=issues,
        stats=model.stats(),
        rules_run=len([rule_id for rule_id in IMPLEMENTED_RULES if rule_id in ruleset.rules]),
        rules_total=len(ruleset),
    )

    if args.format == "json":
        payload = report.to_dict()
        payload["profile_id"] = profile_id
        text = json.dumps(payload, ensure_ascii=False, indent=2)
    elif args.format == "md":
        text = report.render_markdown(max_issues=args.max_issues)
    else:
        text = report.render_text(max_issues=args.max_issues)

    _emit(text, args.out, report.counts())
    return 0


# ---------------------------------------------------------------------------
# fix（学生端）
# ---------------------------------------------------------------------------
def _cmd_fix(args: argparse.Namespace) -> int:
    from .fixer import fix_document  # 延迟导入：模块由 fixer-dev 交付

    profile_id, ruleset, discipline = _resolve_ruleset(args)
    result = fix_document(
        Path(args.docx),
        Path(args.out) if args.out else None,
        profile=discipline,
        ruleset=ruleset,
        dry_run=args.dry_run,
    )
    payload = result.to_dict()
    payload["profile_id"] = profile_id

    if args.format == "json":
        text = json.dumps(payload, ensure_ascii=False, indent=2)
    else:
        text = _render_fix_text(result, profile_id)

    target = args.report or None
    if target:
        Path(target).write_text(text, encoding="utf-8")
        print(f"修复报告已写入：{target}")
    else:
        print(text)
    return 0


def _render_fix_text(result: Any, profile_id: str) -> str:
    counts = getattr(result, "baseline_counts", {}) or {}
    actions = getattr(result, "actions", []) or []
    skipped = getattr(result, "skipped", []) or []
    recheck = getattr(result, "recheck_issues", []) or []
    lines = [
        "tas 安全修复报告（学生端）",
        "=" * 64,
        f"源文件  ：{getattr(result, 'src_path', '')}",
        f"输出文件：{getattr(result, 'out_path', None) or '（dry-run，未写文件）'}",
        f"规则集  ：{profile_id}",
        f"修复前  ：错误 {counts.get('error', 0)} ｜ 警告 {counts.get('warning', 0)} ｜ 提示 {counts.get('info', 0)}",
        f"修复动作：{len(actions)} 项；跳过的风险项：{len(skipped)} 项；复检剩余：{len(recheck)} 项",
        f"内容校验：{'✅ hash 一致（未改内容）' if getattr(result, 'content_hash_before', '') == getattr(result, 'content_hash_after', '') else '❌ hash 不一致（异常）'}",
        "",
    ]
    if actions:
        lines.append("【已修复】")
        for index, action in enumerate(actions[:60], 1):
            rule_id = getattr(action, "rule_id", "?")
            locator = getattr(action, "locator", "")
            description = getattr(action, "description", "")
            lines.append(f"{index:>3}. [{rule_id}] {locator} {description}")
        if len(actions) > 60:
            lines.append(f"    …… 其余 {len(actions) - 60} 项请见 --format json")
        lines.append("")
    if skipped:
        lines.append("【跳过（涉及文字/需人工确认，未自动修改）】")
        seen = set()
        for item in skipped:
            label = item.get("rule_id") or item.get("label") or item.get("kind") or "scope"
            key = (label, item.get("reason"))
            if key in seen:
                continue
            seen.add(key)
            count = item.get("count")
            suffix = f"（{count} 处）" if count else ""
            lines.append(f"  - [{label}]{suffix} {item.get('reason')}")
        lines.append("")
    if recheck:
        lines.append("【复检剩余问题】")
        for issue in recheck[:40]:
            lines.append(
                f"  - [{getattr(issue, 'severity', '?')}] {getattr(issue, 'rule_id', '?')} "
                f"{getattr(issue, 'locator', '')}：{getattr(issue, 'message', '')}"
            )
        if len(recheck) > 40:
            lines.append(f"    …… 其余 {len(recheck) - 40} 项")
        lines.append("")
    lines.append("注意：本工具只改格式，不改文字；参考文献、编号、题注、字数类问题只给建议。")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# refs（教师端）
# ---------------------------------------------------------------------------
def _cmd_refs(args: argparse.Namespace) -> int:
    from .refs import analyze_citations, parse_references  # 延迟导入：模块由 refs-dev 交付

    model = DocModel.load(Path(args.docx))
    roles = classify(model, args.profile or DEFAULT_PROFILE)
    entries = parse_references(model, roles)
    findings = analyze_citations(model, roles, entries)

    verifications = []
    if args.verify:
        from .refs import verify_references

        verifications = verify_references(entries, limit=args.limit, timeout=args.timeout)

    if args.format == "json":
        payload: dict[str, Any] = {
            "file": str(args.docx),
            "entries": [_entry_to_dict(entry) for entry in entries],
            "citations": [_finding_to_dict(finding) for finding in findings],
            "verifications": [_verification_to_dict(item) for item in verifications],
        }
        text = json.dumps(payload, ensure_ascii=False, indent=2)
    else:
        text = _render_refs_text(Path(args.docx), entries, findings, verifications)

    _emit(text, args.out, None)
    return 0


def _entry_to_dict(entry: Any) -> dict:
    data = {}
    for field in ("paragraph_index", "number", "raw", "authors", "title", "venue",
                  "year", "doc_type", "doi", "url"):
        data[field] = getattr(entry, field, None)
    return data


def _finding_to_dict(finding: Any) -> dict:
    data = {}
    for field in ("level", "message", "locator", "actual", "expected"):
        data[field] = getattr(finding, field, None)
    return data


def _verification_to_dict(result: Any) -> dict:
    data = {}
    for field in ("status", "source", "matched_title", "matched_year", "matched_venue",
                  "doi", "url", "score", "note", "evidence"):
        data[field] = getattr(result, field, None)
    entry = getattr(result, "entry", None)
    if entry is not None:
        data["entry"] = _entry_to_dict(entry)
    return data


_STATUS_LABEL = {
    "verified": "✅ 已核验",
    "likely": "🟡 部分匹配（需人工确认）",
    "not_found": "⚪ 未检索到（需人工核实）",
    "error": "❌ 接口错误",
}


def _render_refs_text(docx: Path, entries: list, findings: list, verifications: list) -> str:
    lines = [
        "tas 参考文献核验报告（教师端，只读不改文档）",
        "=" * 64,
        f"文件：{docx}",
        f"参考文献：{len(entries)} 条",
        "",
    ]
    if not entries:
        lines.append("未解析到参考文献条目（检查“参考文献”标题与 [n] 编号格式）。")
        lines.append("")
    if verifications:
        summary: dict[str, int] = {}
        for item in verifications:
            status = getattr(item, "status", "error")
            summary[status] = summary.get(status, 0) + 1
        lines.append(
            "真实性核验："
            + " ｜ ".join(f"{_STATUS_LABEL.get(k, k)} {v}" for k, v in summary.items())
        )
        lines.append("")
        for index, item in enumerate(verifications, 1):
            entry = getattr(item, "entry", None)
            number = getattr(entry, "number", None) or index
            title = getattr(entry, "title", "") or getattr(entry, "raw", "")[:60]
            lines.append(f"[{number}] {title}")
            lines.append(
                f"    结论：{_STATUS_LABEL.get(getattr(item, 'status', 'error'), getattr(item, 'status', ''))}"
                f" ｜ 来源：{getattr(item, 'source', '') or '-'}"
                f" ｜ 匹配度：{getattr(item, 'score', 0.0):.2f}"
            )
            matched = getattr(item, "matched_title", "") or ""
            if matched:
                lines.append(
                    f"    匹配：{matched}"
                    f"（{getattr(item, 'matched_year', '') or '?'}"
                    f" ｜ {getattr(item, 'matched_venue', '') or '?'}）"
                )
            if getattr(item, "doi", ""):
                lines.append(f"    DOI ：{item.doi}")
            if getattr(item, "note", ""):
                lines.append(f"    说明：{item.note}")
            for url in (getattr(item, "evidence", None) or [])[:2]:
                lines.append(f"    证据：{url}")
            lines.append("")
    else:
        lines.append("（未联网核验；加 --verify 可查询 Crossref/OpenAlex/arXiv 元数据）")
        lines.append("")

    lines.append("引用一致性：")
    if not findings:
        lines.append("  未发现引用一致性问题。")
    else:
        for finding in findings[:80]:
            lines.append(
                f"  - [{getattr(finding, 'level', '?')}] {getattr(finding, 'locator', '')}："
                f"{getattr(finding, 'message', '')}"
            )
        if len(findings) > 80:
            lines.append(f"  …… 其余 {len(findings) - 80} 项")
    lines.append("")
    lines.append("红线：未检索到 ≠ 造假；中文期刊可能未被 Crossref/OpenAlex 收录，需人工核实。")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# batch / web（教师端批量与 Web 双入口）
# ---------------------------------------------------------------------------
def _cmd_batch(args: argparse.Namespace) -> int:
    from dataclasses import asdict

    from .batch import check_directory, check_files, summarize
    from .batch_fix import fix_directory, fix_files, summarize_fix, write_fix_zip
    from .export import export_csv, export_html, export_xlsx
    from .feedback import export_feedback_zip
    from .roster import match_submissions, parse_roster_file

    profile_id, ruleset, discipline = _resolve_ruleset(args)
    paths = [Path(item) for item in args.paths]
    directories = [item for item in paths if item.is_dir()]
    files = [item for item in paths if item.is_file()]
    if directories and files:
        raise ValueError("请传一个目录或若干文件，不要混用")
    if directories:
        results = check_directory(
            directories[0], ruleset=ruleset, profile=discipline, ruleset_id=profile_id
        )
    else:
        results = check_files(files, ruleset=ruleset, profile=discipline, ruleset_id=profile_id)
    summary = summarize(results, ruleset_id=profile_id)

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    thumbnails = None
    thumb_summary: dict | None = None
    thumb_zip = None
    if getattr(args, "preview", False):
        from .thumbnails import (
            export_thumbnail_zip,
            render_thumbnails,
            summarize_thumbs,
            thumbnails_map,
        )

        thumb_results = render_thumbnails(
            [item.path for item in results],
            out_dir / "thumbnails",
            max_pages=1,
            width=int(getattr(args, "preview_width", 900)),
        )
        thumbnails = thumbnails_map(thumb_results)
        thumb_summary = summarize_thumbs(thumb_results)
        if any(item.ok for item in thumb_results):
            thumb_zip = export_thumbnail_zip(thumb_results, out_dir / "thumbnails.zip")

    refs_summary: dict | None = None
    refs_json = None
    refs_md = None
    if getattr(args, "refs", False):
        from .batch_refs import verify_batch, write_refs_report

        refs_docs, refs_obj = verify_batch(
            [item.path for item in results],
            profile=discipline,
            limit_per_doc=int(getattr(args, "refs_limit", 5)),
            timeout=float(getattr(args, "refs_timeout", 10.0)),
            pause=0.1,
        )
        refs_json, refs_md = write_refs_report(
            refs_obj,
            refs_docs,
            out_dir / "refs-report.json",
            out_dir / "refs-report.md",
        )
        refs_summary = refs_obj.to_dict()

    xlsx_path = export_xlsx(results, summary, out_dir / "batch.xlsx")
    csv_path = export_csv(results, summary, out_dir / "batch.csv")
    html_path = out_dir / "batch.html"
    html_path.write_text(
        export_html(summary, results, thumbnails=thumbnails), encoding="utf-8"
    )
    feedback_zip = export_feedback_zip(
        results,
        out_dir / "feedback.zip",
        summary=summary,
        profile_label=f"{profile_id}（{discipline}）",
    )

    fix_summary: dict | None = None
    fixed_zip = None
    if getattr(args, "fix", False):
        fixed_dir = (
            Path(args.fixed_dir) if getattr(args, "fixed_dir", None) else out_dir / "fixed"
        )
        if directories:
            fix_results = fix_directory(
                directories[0], fixed_dir, ruleset=ruleset, profile=discipline
            )
        else:
            fix_results = fix_files(files, fixed_dir, ruleset=ruleset, profile=discipline)
        fix_summary = summarize_fix(fix_results)
        fixed_zip = write_fix_zip(fix_results, out_dir / "fixed.zip")

    roster_report: dict | None = None
    roster_report_path: Path | None = None
    if getattr(args, "roster", None):
        roster_entries = parse_roster_file(Path(args.roster))
        roster_report = match_submissions(
            [item.path for item in results], roster_entries
        )
        roster_report_path = out_dir / "roster-report.txt"
        roster_report_path.write_text(
            _render_roster_text(roster_report), encoding="utf-8"
        )
    summary_path = out_dir / "batch-summary.json"
    summary_path.write_text(
        json.dumps(
            {
                "summary": asdict(summary),
                "documents": [
                    {
                        "path": str(item.path),
                        "issues": len(item.issues),
                        "error": item.error,
                        "duration_s": item.duration_s,
                    }
                    for item in results
                ],
            },
            ensure_ascii=False,
            indent=2,
            default=str,
        ),
        encoding="utf-8",
    )

    if args.format == "json":
        print(
            json.dumps(
                {
                    "summary": asdict(summary),
                    "files": {
                        "xlsx": str(xlsx_path),
                        "csv": str(csv_path),
                        "html": str(html_path),
                        "json": str(summary_path),
                        "feedback": str(feedback_zip),
                        "fixed_zip": str(fixed_zip) if fixed_zip else None,
                        "roster_report": (
                            str(roster_report_path) if roster_report_path else None
                        ),
                        "thumbnails_zip": str(thumb_zip) if thumb_zip else None,
                        "refs_json": str(refs_json) if refs_json else None,
                        "refs_md": str(refs_md) if refs_md else None,
                    },
                    "fix": fix_summary,
                    "roster": roster_report,
                    "thumbnails": thumb_summary,
                    "refs": refs_summary,
                },
                ensure_ascii=False,
                indent=2,
                default=str,
            )
        )
        return 0

    print("tas 班级批量检查报告")
    print("=" * 64)
    print(f"规范    ：{profile_id}（{discipline}）")
    print(f"文件    ：共 {summary.total} 份 ｜ 成功 {summary.checked} ｜ 失败 {summary.failed}")
    print(f"通过率  ：{summary.pass_rate:.1%}（无 error 的文档占比）")
    print(
        f"问题    ：错误 {summary.counts.get('error', 0)} ｜ 警告 {summary.counts.get('warning', 0)} ｜ "
        f"提示 {summary.counts.get('info', 0)}"
    )
    print("")
    print(f"{'文件':<52}{'错误':>4}{'警告':>4}{'提示':>4}  状态")
    for item in summary.per_doc[:50]:
        name = Path(item.get("path", "")).name
        if len(name) > 48:
            name = name[:47] + "…"
        print(
            f"{name:<52}{item.get('error', 0):>4}{item.get('warning', 0):>4}"
            f"{item.get('info', 0):>4}  {item.get('status', '')}"
        )
    if summary.top_rules:
        print("")
        print("共性问题 Top：")
        for rule_id, count in summary.top_rules:
            print(f"  - {rule_id}: {count} 次")
    print("")
    print(f"报告已写入：{out_dir}")
    print(
        f"  Excel：{xlsx_path.name} ｜ CSV：{csv_path.name} ｜ HTML：{html_path.name}"
        f" ｜ JSON：{summary_path.name}"
    )
    print(f"  学生反馈：{feedback_zip.name}（每份文档一个 .txt，可直接发给学生）")
    if fix_summary:
        before_error = fix_summary.get("before_counts", {}).get("error", 0)
        after_error = fix_summary.get("after_counts", {}).get("error", 0)
        print(
            f"  批量修复：成功 {fix_summary.get('fixed', 0)} ｜ 失败 {fix_summary.get('failed', 0)}"
            f" ｜ 错误 {before_error} → {after_error} ｜ 修复包 {fixed_zip.name}"
        )
    if roster_report:
        print(
            f"  名单核对：应交 {roster_report.get('total_roster', 0)}"
            f" ｜ 收到 {roster_report.get('total_files', 0)}"
            f" ｜ 已匹配 {roster_report.get('matched', 0)}"
            f" ｜ 缺交 {len(roster_report.get('missing', []))}"
            f" ｜ 重复 {len(roster_report.get('duplicates', []))}"
            f" ｜ 报告 {roster_report_path.name}"
        )
    if thumb_summary:
        zip_name = thumb_zip.name if thumb_zip else "（无，全部渲染失败）"
        print(
            f"  缩略图：成功 {thumb_summary.get('ok', 0)}/{thumb_summary.get('total', 0)}"
            f" ｜ 缩略图包 {zip_name}"
        )
    if refs_summary:
        print(
            f"  文献核验：文档 {refs_summary.get('total_docs', 0)}"
            f" ｜ 文献 {refs_summary.get('total_entries', 0)}"
            f" ｜ verified {refs_summary.get('verified', 0)}"
            f" ｜ likely {refs_summary.get('likely', 0)}"
            f" ｜ not_found {refs_summary.get('not_found', 0)}"
            f" ｜ error {refs_summary.get('error', 0)}"
        )
        print(
            f"  文献报告：{refs_md.name}（{refs_json.name}）；"
            "not_found 仅表示未在当前索引检索到，中文期刊可能未收录，需人工核实"
        )
    return 0


def _render_roster_text(report: dict) -> str:
    """把 roster.match_submissions 的结果渲染成老师可读的文本。"""
    lines = [
        "【名单核对】",
        (
            f"名单人数：{report.get('total_roster', 0)} ｜ "
            f"收到文件：{report.get('total_files', 0)} ｜ "
            f"已匹配：{report.get('matched', 0)}"
        ),
        "",
    ]
    missing = report.get("missing") or []
    lines.append(f"缺交（{len(missing)} 人）：")
    if missing:
        for item in missing:
            name = item.get("name", "")
            lines.append(f"  - {item.get('student_id', '')} {name}".rstrip())
    else:
        lines.append("  （无，全部交齐）")

    duplicates = report.get("duplicates") or []
    lines.append("")
    lines.append(f"重复提交（{len(duplicates)} 人）：")
    if duplicates:
        for item in duplicates:
            files = "、".join(Path(path).name for path in item.get("files", []))
            lines.append(
                f"  - {item.get('student_id', '')} {item.get('name', '')}：{files}".rstrip()
            )
    else:
        lines.append("  （无）")

    unmatched = report.get("unmatched_files") or []
    lines.append("")
    lines.append(f"匹配不到名单的文件（{len(unmatched)} 份）：")
    if unmatched:
        for item in unmatched:
            lines.append(
                f"  - {Path(item.get('file', '')).name}：{item.get('reason', '')}"
            )
    else:
        lines.append("  （无）")

    warnings = report.get("naming_warnings") or []
    lines.append("")
    lines.append(f"文件名提醒（{len(warnings)} 条）：")
    if warnings:
        for item in warnings:
            lines.append(
                f"  - {Path(item.get('file', '')).name}：{item.get('reason', '')}"
            )
    else:
        lines.append("  （无）")
    return "\n".join(lines) + "\n"


def _cmd_preview(args: argparse.Namespace) -> int:
    from .render import RenderError, render_preview

    out_dir = Path(args.out_dir)
    try:
        pages = render_preview(
            Path(args.docx), out_dir, max_pages=args.pages, width=args.width
        )
    except RenderError as exc:
        print(f"渲染失败：{exc}", file=sys.stderr)
        return 2
    payload = {
        "file": str(args.docx),
        "out_dir": str(out_dir),
        "pages": [str(page) for page in pages],
    }
    if args.format == "json":
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    print(f"预览已生成：{len(pages)} 页（{out_dir}）")
    for page in pages:
        print(f"  - {page.name}（{page.stat().st_size // 1024} KB）")
    print("打开 PNG 即可查看版式；Web 端在检查/修复结果页自动展示。")
    return 0


def _cmd_web(args: argparse.Namespace) -> int:
    import threading
    import webbrowser

    from .web.app import serve

    if getattr(args, "verbose", False):
        os.environ["TAS_WEB_VERBOSE"] = "1"
        print("请求日志已开启（--verbose）：每个请求会打印 客户端IP 方法 路径 状态码。", file=sys.stderr)
        print("网络自检：手机浏览器打开 http://<电脑局域网IP>:8765/api/diag", file=sys.stderr)
    if getattr(args, "open_browser", False):
        url = f"http://127.0.0.1:{args.port}/"
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    serve(args.host, args.port)
    return 0


# ---------------------------------------------------------------------------
# rules / profiles / spec
# ---------------------------------------------------------------------------
def _cmd_rules(args: argparse.Namespace) -> int:
    profile_id, ruleset, _discipline = _resolve_ruleset(args)
    selected = [
        rule
        for rule in ruleset.rules.values()
        if args.phase == "all" or rule.phase == args.phase
    ]
    if args.json:
        payload = {
            "profile_id": profile_id,
            "total": len(ruleset),
            "selected": len(selected),
            "implemented": [rule_id for rule_id in IMPLEMENTED_RULES if rule_id in ruleset.rules],
            "rules": [
                {
                    "id": rule.id,
                    "phase": rule.phase,
                    "severity": rule.severity,
                    "category": rule.category,
                    "name": rule.name,
                    "source": rule.source,
                }
                for rule in sorted(selected, key=lambda item: item.id)
            ],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    print(f"规则集：{profile_id} ｜ 规则 {len(ruleset)} 条"
          f"（P0 {sum(1 for r in ruleset.rules.values() if r.phase == 'P0')} / "
          f"P1 {sum(1 for r in ruleset.rules.values() if r.phase == 'P1')}）")
    implemented_here = [rule_id for rule_id in IMPLEMENTED_RULES if rule_id in ruleset.rules]
    print(
        f"已实现检查：{len(implemented_here)}/{len(ruleset)} 条"
        f"（全局已实现 {len(IMPLEMENTED_RULES)} 条）"
    )
    print("")
    print(f"{'ID':<10}{'阶段':<6}{'级别':<8}{'分类':<14}{'名称':<26}来源")
    for rule in sorted(selected, key=lambda item: item.id):
        status = "*" if rule.id in IMPLEMENTED_RULES else " "
        print(
            f"{rule.id:<10}{rule.phase:<6}{rule.severity:<8}{rule.category:<14}"
            f"{status}{rule.name:<25}{rule.source}"
        )
    print("")
    print("注：名称前带 * 的规则已实现检查器；本规则集若全部带 * 则无未实现规则。")
    return 0


def _cmd_profiles(args: argparse.Namespace) -> int:
    profiles = list_profiles()
    if args.json:
        payload = [
            {
                "id": profile.id,
                "name": profile.name,
                "discipline": profile.discipline,
                "rules_path": str(profile.rules_path),
                "source": profile.source,
                "notes": profile.notes,
            }
            for profile in profiles
        ]
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    if not profiles:
        print("未找到 profiles/index.json；当前回退到 rules/rules.json。")
        return 0
    print(f"共 {len(profiles)} 个 profile（默认 {DEFAULT_PROFILE_ID}）：")
    print("")
    print(f"{'ID':<26}{'学科':<12}{'名称'}")
    for profile in profiles:
        print(f"{profile.id:<26}{profile.discipline:<12}{profile.name}")
        print(f"    规则：{profile.rules_path}")
        if profile.notes:
            print(f"    说明：{profile.notes}")
    return 0


def _cmd_spec(args: argparse.Namespace) -> int:
    if args.spec_command == "extract":
        return _cmd_spec_extract(args)
    if args.spec_command == "parse":
        return _cmd_spec_parse(args)
    if args.spec_command == "apply":
        return _cmd_spec_apply(args)
    return 2


def _cmd_spec_extract(args: argparse.Namespace) -> int:
    text = extract_spec_text(args.file)
    stats = spec_stats(text)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"规范文本已写入：{args.out}（{stats['chars']} 字符 / {stats['paragraphs']} 段）")
    else:
        print(text)
    return 0


def _cmd_spec_parse(args: argparse.Namespace) -> int:
    from dataclasses import replace

    from .llm import LLMConfigError, LLMError, chat_json, load_config
    from .spec_profile import diff_candidates, extract_candidates, render_report_markdown

    spec_text = extract_spec_text(args.file)
    base_profile, ruleset = resolve_profile(args.base_profile)
    try:
        config = load_config()
        if args.model:
            config = replace(config, model=args.model)

        def _chat(messages):
            return chat_json(messages, config=config)

        report = extract_candidates(
            spec_text,
            ruleset,
            chat=_chat,
            model=config.model,
            max_spec_chars=args.max_spec_chars,
            base_profile=base_profile.id,
        )
    except LLMConfigError as exc:
        print(f"LLM 配置错误：{exc}", file=sys.stderr)
        return 2
    except LLMError as exc:
        print(f"LLM 调用失败：{exc}", file=sys.stderr)
        return 2

    diff_report = diff_candidates(report, ruleset, base_profile=base_profile.id)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    candidates_path = out_dir / "candidates.json"
    report_path = out_dir / "diff.md"
    payload = {
        "base_profile": base_profile.id,
        "source_file": str(Path(args.file)),
        "model": report.model,
        "extraction": report.to_dict(),
        "diff": diff_report.to_dict(),
    }
    candidates_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    report_path.write_text(render_report_markdown(report, diff_report), encoding="utf-8")

    counts = diff_report.counts()
    if args.json:
        print(
            json.dumps(
                {
                    "base_profile": base_profile.id,
                    "model": report.model,
                    "counts": counts,
                    "unmapped": len(diff_report.unmapped),
                    "dropped": len(diff_report.dropped),
                    "uncovered_existing": len(diff_report.uncovered_existing),
                    "candidates_path": str(candidates_path),
                    "diff_path": str(report_path),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    print(f"规范抽取完成（模型：{report.model or config.model}）")
    print(
        f"结果：keep {counts.get('keep', 0)} ｜ modify {counts.get('modify', 0)} ｜ "
        f"new {counts.get('new', 0)} ｜ unmapped {len(diff_report.unmapped)} ｜ "
        f"丢弃 {len(diff_report.dropped)} ｜ 未覆盖现有规则 {len(diff_report.uncovered_existing)}"
    )
    print(f"候选：{candidates_path}")
    print(f"审阅报告：{report_path}")
    print("")
    print("下一步：审阅 diff.md / candidates.json（可删除不认可的候选），再执行")
    print(
        f"  tas spec apply \"{candidates_path}\" --id <profile_id> --name \"<新规范名称>\""
    )
    return 0


def _cmd_spec_apply(args: argparse.Namespace) -> int:
    from .spec_profile import apply_candidates_file

    profile_id = args.profile_id
    out_path = (
        Path(args.out)
        if args.out
        else project_root() / "profiles" / f"{profile_id}.rules.json"
    )
    index_path = None if args.no_register else profiles_index_path()
    result = apply_candidates_file(
        Path(args.candidates),
        profile_id=profile_id,
        name=args.name,
        discipline=args.discipline,
        out_path=out_path,
        index_path=index_path,
        source_spec=args.source,
    )
    print(f"新 profile 已生成：{profile_id}")
    print(f"规则文件：{result['rules_path']}")
    print(
        f"规则总数 {result['rules_total']}（keep {result['keep']} ｜ modify {result['modify']} ｜ "
        f"new {result['new']}）；unmapped {result['unmapped']} 条"
    )
    if result["registered"]:
        print(f"已注册：{profiles_index_path()}（`tas profiles list` 可见）")
    else:
        print("未注册（--no-register）；如需使用请手动加入 profiles/index.json")
    if result["uncovered_existing"]:
        print(f"提示：{len(result['uncovered_existing'])} 条现有规则未被本次候选覆盖（保留原样）")
    print("注意：new 规则在绑定检查器前不会参与检查；用于检查前请先补 checker。")
    return 0


# ---------------------------------------------------------------------------
def _emit(text: str, out: Optional[str], counts: Optional[dict]) -> None:
    if out:
        Path(out).write_text(text, encoding="utf-8")
        if counts:
            print(
                f"报告已写入：{out}（错误 {counts.get('error', 0)}，"
                f"警告 {counts.get('warning', 0)}，提示 {counts.get('info', 0)}）"
            )
        else:
            print(f"报告已写入：{out}")
    else:
        print(text)


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
