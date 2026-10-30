"""教师端批量报告导出：Excel（xlsx）/ CSV / HTML。

对应需求 FR-06「批量检查 + 班级汇总 + 导出」：
- :func:`export_xlsx`：3 个 Sheet
  1. ``汇总``：总数 / 通过率 / 规则集 / 时间 + 共性问题 Top N 表；
  2. ``逐份``：文件 / 错误 / 警告 / 提示 / 状态 / 耗时；
  3. ``问题明细``：文件 / 规则 / 级别 / 位置 / 说明 / 实际值 / 期望值。
- :func:`export_csv`：一个 UTF-8-BOM 的 CSV，逐份行 + 问题明细行（``类型`` 列区分），
  方便 Excel 直接打开与二次统计；
- :func:`export_html`：离线可看的 HTML 字符串（教师端下载/预览用，无外部依赖）。

约定：
- 只依赖 :mod:`tas.batch` 的结果模型与标准库；openpyxl 在 :func:`export_xlsx` 内部
  延迟导入（未安装时 ``import tas.export`` 仍可用，只有真正导出 xlsx 才报错）；
- 输出目录不存在时自动创建；返回写入的 :class:`pathlib.Path`；
- 中文列名与表头固定（测试与集成依赖），问题明细的实际值/期望值格式化为可读文本。
"""

from __future__ import annotations

import csv
import html
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

from .batch import SEVERITY_KEYS, BatchSummary, DocResult
from .report import SEVERITY_LABEL

__all__ = [
    "SHEET_SUMMARY",
    "SHEET_PER_DOC",
    "SHEET_ISSUES",
    "CSV_HEADERS",
    "export_xlsx",
    "export_csv",
    "export_html",
]

SHEET_SUMMARY = "汇总"
SHEET_PER_DOC = "逐份"
SHEET_ISSUES = "问题明细"

#: Sheet2 逐份表头（契约要求）
PER_DOC_HEADERS = ["文件", "错误", "警告", "提示", "状态", "耗时(s)"]
#: Sheet3 问题明细表头（契约要求）
ISSUE_HEADERS = ["文件", "规则", "级别", "位置", "说明", "实际值", "期望值"]
#: CSV 统一表头：逐份行与问题明细行合并在一张表里，用「类型」列区分
CSV_HEADERS = [
    "类型",
    "文件",
    "规则",
    "级别",
    "位置",
    "说明",
    "实际值",
    "期望值",
    "错误数",
    "警告数",
    "提示数",
    "状态",
    "耗时(s)",
    "规则集",
    "生成时间",
]


# ---------------------------------------------------------------------------
# 值格式化 / 状态
# ---------------------------------------------------------------------------
def _format_value(value: Any) -> str:
    """把 actual/expected 这类 JSON 值格式化成单元格/CSV 可读文本。"""
    if value is None:
        return ""
    if isinstance(value, dict):
        return "，".join(f"{key}={_format_value(item)}" for key, item in value.items())
    if isinstance(value, (list, tuple)):
        return "[" + "，".join(_format_value(item) for item in value) + "]"
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def _counts(result: DocResult) -> dict[str, int]:
    try:
        counts = result.counts()
    except Exception:  # pragma: no cover - 非标准结果对象兜底
        counts = {key: 0 for key in SEVERITY_KEYS}
    return {key: int(counts.get(key, 0)) for key in SEVERITY_KEYS}


def _status(result: DocResult) -> str:
    if result.error is not None:
        return "失败"
    return "通过" if _counts(result)["error"] == 0 else "未通过"


def _duration(result: DocResult) -> float:
    return round(float(getattr(result, "duration_s", 0.0) or 0.0), 3)


def _level_label(severity: Optional[str]) -> str:
    return SEVERITY_LABEL.get(severity or "", severity or "")


def _ensure_parent(out_path: Path) -> Path:
    out_path = Path(out_path)
    if out_path.parent and not out_path.parent.exists():
        out_path.parent.mkdir(parents=True, exist_ok=True)
    return out_path


def _percent_text(rate: float) -> str:
    return f"{float(rate or 0.0):.1%}"


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def export_xlsx(
    results: Sequence[DocResult], summary: BatchSummary, out_path: Path | str
) -> Path:
    """导出 Excel（3 个 Sheet）；返回写入路径。openpyxl 延迟导入。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    out = _ensure_parent(Path(out_path))
    items = list(results)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F5597")
    title_font = Font(bold=True, size=14)
    section_font = Font(bold=True, size=11, color="1F3864")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _header_row(ws, row: int, headers: Sequence[str]) -> None:
        for column, name in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=column, value=name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

    def _fit_columns(ws, widths: dict[int, int]) -> None:
        for column, width in widths.items():
            ws.column_dimensions[get_column_letter(column)].width = width

    workbook = Workbook()

    # ---- Sheet1：汇总 ----------------------------------------------------
    ws = workbook.active
    ws.title = SHEET_SUMMARY
    ws["A1"] = "TAS 班级批量检查汇总"
    ws["A1"].font = title_font
    ws.merge_cells("A1:B1")

    counts = {key: int(summary.counts.get(key, 0)) for key in SEVERITY_KEYS}
    summary_rows = [
        ("生成时间", summary.generated_at),
        ("规则集", summary.ruleset_id or "（未标注）"),
        ("文件总数", summary.total),
        ("检查成功", summary.checked),
        ("检查失败", summary.failed),
        ("通过率", float(summary.pass_rate)),
        ("通过 / 成功", f"{summary.pass_count} / {summary.checked}"),
        ("错误总数", counts["error"]),
        ("警告总数", counts["warning"]),
        ("提示总数", counts["info"]),
    ]
    row = 3
    for label, value in summary_rows:
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.font = Font(bold=True)
        value_cell = ws.cell(row=row, column=2, value=value)
        if label == "通过率":
            value_cell.number_format = "0.00%"
        row += 1

    row += 1
    ws.cell(row=row, column=1, value=f"共性问题 Top {len(summary.top_rules)}").font = (
        section_font
    )
    row += 1
    _header_row(ws, row, ["规则", "命中次数"])
    row += 1
    if summary.top_rules:
        for rule_id, hit_count in summary.top_rules:
            ws.cell(row=row, column=1, value=rule_id).border = border
            ws.cell(row=row, column=2, value=int(hit_count)).border = border
            row += 1
    else:
        ws.cell(row=row, column=1, value="（无问题）")
        row += 1
    _fit_columns(ws, {1: 22, 2: 42})

    # ---- Sheet2：逐份 ----------------------------------------------------
    ws = workbook.create_sheet(SHEET_PER_DOC)
    _header_row(ws, 1, PER_DOC_HEADERS)
    for index, item in enumerate(items, start=2):
        doc_counts = _counts(item)
        values = [
            str(item.path),
            doc_counts["error"],
            doc_counts["warning"],
            doc_counts["info"],
            "失败" if item.error is not None else _status(item),
            _duration(item),
        ]
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=index, column=column, value=value)
            cell.border = border
            if column == 1:
                cell.alignment = Alignment(vertical="center")
            elif column == len(values):
                cell.number_format = "0.000"
    ws.freeze_panes = "A2"
    if items:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(PER_DOC_HEADERS))}{len(items) + 1}"
    _fit_columns(ws, {1: 58, 2: 8, 3: 8, 4: 8, 5: 10, 6: 10})

    # ---- Sheet3：问题明细 ------------------------------------------------
    ws = workbook.create_sheet(SHEET_ISSUES)
    _header_row(ws, 1, ISSUE_HEADERS)
    row = 2
    for item in items:
        if item.error is not None:
            detail = [
                str(item.path),
                "（加载/检查失败）",
                "失败",
                "",
                str(item.error),
                "",
                "",
            ]
            for column, value in enumerate(detail, start=1):
                cell = ws.cell(row=row, column=column, value=value)
                cell.border = border
            row += 1
            continue
        for issue in item.issues:
            detail = [
                str(item.path),
                issue.rule_id,
                _level_label(issue.severity),
                issue.locator,
                issue.message,
                _format_value(issue.actual),
                _format_value(issue.expected),
            ]
            for column, value in enumerate(detail, start=1):
                cell = ws.cell(row=row, column=column, value=value)
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=column in (4, 5, 6, 7))
            row += 1
    ws.freeze_panes = "A2"
    if row > 2:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(ISSUE_HEADERS))}{row - 1}"
    _fit_columns(ws, {1: 46, 2: 12, 3: 8, 4: 40, 5: 46, 6: 26, 7: 32})

    workbook.save(str(out))
    return out


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------
def _csv_row(
    record_type: str,
    *,
    path: str = "",
    rule: str = "",
    level: str = "",
    locator: str = "",
    message: str = "",
    actual: str = "",
    expected: str = "",
    counts: Optional[dict[str, int]] = None,
    status: str = "",
    duration: str = "",
    ruleset_id: str = "",
    generated_at: str = "",
) -> list[str]:
    counts = counts or {}
    return [
        record_type,
        path,
        rule,
        level,
        locator,
        message,
        actual,
        expected,
        str(counts.get("error", "")),
        str(counts.get("warning", "")),
        str(counts.get("info", "")),
        status,
        duration,
        ruleset_id,
        generated_at,
    ]


def export_csv(
    results: Sequence[DocResult], summary: BatchSummary, out_path: Path | str
) -> Path:
    """导出 CSV：首行「汇总」，其余为逐份行 + 问题明细行（用「类型」列区分）。

    返回写入路径；编码 ``utf-8-sig``，Excel 打开中文不乱码。
    """
    out = _ensure_parent(Path(out_path))
    items = list(results)
    counts = {key: int(summary.counts.get(key, 0)) for key in SEVERITY_KEYS}

    with out.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(CSV_HEADERS)

        # 1) 班级汇总行
        writer.writerow(
            _csv_row(
                "汇总",
                path=f"共 {summary.total} 份",
                rule=summary.ruleset_id or "（未标注）",
                counts=counts,
                status=f"通过率 {_percent_text(summary.pass_rate)}"
                f"（通过 {summary.pass_count}/{summary.checked}）",
                ruleset_id=summary.ruleset_id,
                generated_at=summary.generated_at,
            )
        )

        # 2) 逐份行 + 问题明细行
        for item in items:
            doc_counts = _counts(item)
            writer.writerow(
                _csv_row(
                    "逐份",
                    path=str(item.path),
                    counts=doc_counts,
                    status=_status(item),
                    duration=f"{_duration(item):.3f}",
                )
            )
            if item.error is not None:
                writer.writerow(
                    _csv_row(
                        "失败",
                        path=str(item.path),
                        rule="（加载/检查失败）",
                        level="失败",
                        message=str(item.error),
                        status="失败",
                    )
                )
                continue
            for issue in item.issues:
                writer.writerow(
                    _csv_row(
                        "问题",
                        path=str(item.path),
                        rule=issue.rule_id,
                        level=_level_label(issue.severity),
                        locator=issue.locator,
                        message=issue.message,
                        actual=_format_value(issue.actual),
                        expected=_format_value(issue.expected),
                    )
                )
    return out


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------
def _esc(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def export_html(
    summary: BatchSummary,
    results: Sequence[DocResult],
    *,
    thumbnails: Optional[dict[str, str]] = None,
) -> str:
    """生成离线可看的 HTML 报告字符串（教师端预览/下载用）。

    ``thumbnails``：``{文档文件名: 相对路径}``（如 ``thumbnails/xxx.png``），
    传入后在「逐份结果」表增加缩略图列；默认 ``None`` 保持原样。
    """
    items = list(results)
    counts = {key: int(summary.counts.get(key, 0)) for key in SEVERITY_KEYS}

    top_rows = "".join(
        f"<tr><td>{_esc(rule_id)}</td><td>{int(hit)}</td></tr>"
        for rule_id, hit in summary.top_rules
    ) or "<tr><td colspan='2' class='muted'>无</td></tr>"

    top_max = max((int(hit) for _, hit in summary.top_rules), default=0) or 1
    top_bars = "".join(
        "<div class='bar-row'>"
        f"<span class='bar-label'>{_esc(rule_id)}</span>"
        "<span class='bar-track'>"
        f"<span class='bar-fill' style='width:{max(2, round(int(hit) / top_max * 100))}%'></span>"
        "</span>"
        f"<span class='bar-value'>{int(hit)}</span>"
        "</div>"
        for rule_id, hit in summary.top_rules
    ) or "<p class='muted'>无共性问题</p>"
    pass_pct = max(0.0, min(100.0, round(float(summary.pass_rate) * 100, 1)))

    per_doc_rows: list[str] = []
    issue_rows: list[str] = []
    thumb_header = "<th>缩略图</th>" if thumbnails is not None else ""
    for item in items:
        doc_counts = _counts(item)
        status = _status(item)
        css = "pass" if status == "通过" else ("fail" if status == "失败" else "warn")
        thumb_cell = ""
        if thumbnails is not None:
            relative = thumbnails.get(Path(item.path).name, "")
            if relative:
                thumb_cell = (
                    "<td class='thumb'><img src='"
                    + _esc(relative)
                    + "' loading='lazy' alt='缩略图'></td>"
                )
            else:
                thumb_cell = "<td class='thumb muted'>—</td>"
        per_doc_rows.append(
            "<tr>"
            + thumb_cell
            + f"<td>{_esc(Path(item.path).name)}</td>"
            f"<td class='num err'>{doc_counts['error']}</td>"
            f"<td class='num warn'>{doc_counts['warning']}</td>"
            f"<td class='num'>{doc_counts['info']}</td>"
            f"<td class='{css}'>{_esc(status)}</td>"
            f"<td class='num'>{_duration(item):.3f}s</td>"
            "</tr>"
        )
        if item.error is not None:
            issue_rows.append(
                "<tr>"
                f"<td>{_esc(Path(item.path).name)}</td>"
                "<td>（加载/检查失败）</td><td class='fail'>失败</td><td>—</td>"
                f"<td colspan='3'>{_esc(item.error)}</td>"
                "</tr>"
            )
            continue
        for issue in item.issues:
            level = _level_label(issue.severity)
            issue_rows.append(
                "<tr>"
                f"<td>{_esc(Path(item.path).name)}</td>"
                f"<td>{_esc(issue.rule_id)}</td>"
                f"<td class='{_esc(issue.severity)}'>{_esc(level)}</td>"
                f"<td>{_esc(issue.locator)}</td>"
                f"<td>{_esc(issue.message)}</td>"
                f"<td>{_esc(_format_value(issue.actual))}</td>"
                f"<td>{_esc(_format_value(issue.expected))}</td>"
                "</tr>"
            )

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>TAS 班级批量检查报告</title>
<style>
  body {{ font-family: "Microsoft YaHei", "PingFang SC", system-ui, sans-serif;
         margin: 24px auto; max-width: 1180px; color: #1f2937; padding: 0 16px; }}
  h1 {{ font-size: 22px; margin-bottom: 4px; }}
  h2 {{ font-size: 16px; margin: 24px 0 8px; color: #1f3864; }}
  .meta {{ color: #6b7280; font-size: 13px; }}
  .cards {{ display: flex; gap: 12px; flex-wrap: wrap; margin: 12px 0; }}
  .card {{ border: 1px solid #e5e7eb; border-radius: 10px; padding: 10px 18px;
           background: #f8fafc; min-width: 108px; }}
  .card b {{ display: block; font-size: 20px; }}
  .err {{ color: #b91c1c; }} .warn {{ color: #b45309; }} .info {{ color: #1d4ed8; }}
  .pass {{ color: #047857; font-weight: 600; }} .fail {{ color: #b91c1c; font-weight: 600; }}
  .num {{ text-align: right; }} .muted {{ color: #9ca3af; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  th, td {{ border: 1px solid #e5e7eb; padding: 6px 8px; text-align: left; }}
  th {{ background: #2f5597; color: #fff; font-weight: 600; }}
  tr:nth-child(even) td {{ background: #f9fafb; }}
  .thumb {{ width: 100px; }}
  .thumb img {{ display: block; max-width: 92px; max-height: 120px; border: 1px solid #e5e7eb;
                border-radius: 4px; background: #fff; }}
  .bars {{ margin: 10px 0 18px; max-width: 760px; }}
  .bar-row {{ display: flex; align-items: center; gap: 10px; margin: 5px 0; font-size: 13px; }}
  .bar-label {{ width: 92px; flex: none; color: #1f3864; font-family: Consolas, monospace; }}
  .bar-track {{ flex: 1; height: 14px; background: #eef2f7; border-radius: 7px; overflow: hidden; }}
  .bar-fill {{ display: block; height: 100%; background: linear-gradient(90deg, #3b82f6, #2563eb);
               border-radius: 7px; }}
  .bar-value {{ width: 48px; text-align: right; color: #374151; font-variant-numeric: tabular-nums; }}
  .progress {{ height: 18px; background: #e5e7eb; border-radius: 9px; overflow: hidden;
               margin: 6px 0 18px; max-width: 760px; }}
  .progress-fill {{ height: 100%; background: linear-gradient(90deg, #10b981, #047857);
                    border-radius: 9px; }}
  .progress-caption {{ font-size: 12px; color: #6b7280; margin-top: -10px; }}
  .foot {{ margin-top: 28px; color: #6b7280; font-size: 12px; line-height: 1.7; }}
</style>
</head>
<body>
<h1>TAS 班级批量检查报告</h1>
<p class="meta">规则集：{_esc(summary.ruleset_id or "（未标注）")} ｜
生成时间：{_esc(summary.generated_at)}</p>
<div class="cards">
  <div class="card"><b>{summary.total}</b><span>文件总数</span></div>
  <div class="card"><b>{summary.checked}</b><span>检查成功</span></div>
  <div class="card"><b>{summary.failed}</b><span>检查失败</span></div>
  <div class="card"><b>{_percent_text(summary.pass_rate)}</b><span>通过率</span></div>
  <div class="card"><b class="err">{counts['error']}</b><span>错误总数</span></div>
  <div class="card"><b class="warn">{counts['warning']}</b><span>警告总数</span></div>
  <div class="card"><b class="info">{counts['info']}</b><span>提示总数</span></div>
</div>

<h2>通过率</h2>
<div class="progress"><div class="progress-fill" style="width:{pass_pct}%"></div></div>
<p class="progress-caption">通过率 {_percent_text(summary.pass_rate)}（{summary.pass_count}/{summary.checked} 份无 error 级问题）</p>

<h2>共性问题 Top {len(summary.top_rules)}（按命中次数）</h2>
<div class="bars">{top_bars}</div>
<table><thead><tr><th>规则</th><th>命中次数</th></tr></thead>
<tbody>{top_rows}</tbody></table>

<h2>逐份结果</h2>
<table><thead><tr>{thumb_header}<th>文件</th><th>错误</th><th>警告</th><th>提示</th>
<th>状态</th><th>耗时</th></tr></thead>
<tbody>{''.join(per_doc_rows) or f"<tr><td colspan='{7 if thumbnails is not None else 6}' class='muted'>没有文档</td></tr>"}</tbody></table>

<h2>问题明细</h2>
<table><thead><tr><th>文件</th><th>规则</th><th>级别</th><th>位置</th>
<th>说明</th><th>实际值</th><th>期望值</th></tr></thead>
<tbody>{''.join(issue_rows) or "<tr><td colspan='7' class='muted'>没有发现问题</td></tr>"}</tbody></table>

<p class="foot">说明：通过率 = 检查成功的文档中无 error 级问题的比例（加载失败不计入分母）。
自动检查结果仅供参考；误报请记录，教师端会用反馈校准规则。</p>
</body>
</html>
"""
