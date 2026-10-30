"""正文引用与参考文献表的一致性检查（离线，只读）。

检查内容（对应 rules.json 的 REF-003，教师端只诊断、不改学生文档）：

1. 正文中的 ``[n]`` 引用是否都能在文献表中找到（含 ``[0]``、``[99]`` 等无效/越界编号）；
2. 文献表中是否有从未被正文引用的条目；
3. 首次引用顺序是否符合 GB/T 7714「按引证先后顺序连续编号」；
4. 文献表编号是否连续、顺序是否正确；
5. 正文引用序号是否设置为上标（直接读取 docx 的 ``w:rPr/w:vertAlign``）。

实现说明：
- ``DocModel``/``RoleInfo`` 不保存 vertAlign，因此这里按 ``model.path`` 重新打开
  docx，用 python-docx 读取 run 的 XML（不改动 ``docx_io``，也不写回文件）；
- 若文档无法重新打开（例如模型不是来自磁盘文件），上标检查自动跳过，不报错；
- 单个段落内的 ``[1-3]``、``[1,2]`` 会展开为多个编号处理。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Optional
from pathlib import Path

from ..docx_io import DocModel
from ..roles import Roles
from .parse import ReferenceEntry

# [n] / [1,2] / [1-3] / [1，3] …（半角/全角方括号、全角数字）
_CITE_RE = re.compile(r"[\[［]\s*([0-9０-９][0-9０-９\s,，、\-–—~～至]*?)\s*[\]］]")
_HEADING_ROLE_RE = re.compile(r"^heading_[1-9]$")
_RANGE_RE = re.compile(r"^(\d+)\s*[-–—~～至]\s*(\d+)$")
_FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
_SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}


@dataclass
class CitationFinding:
    """一条引用一致性问题。

    ``level`` ∈ {"error", "warning", "info"}；``locator`` 形如 ``P12``（正文段落）
    或 ``refs/P20``（参考文献条目段落）；``actual``/``expected`` 为机器可读值。
    """

    level: str
    message: str
    locator: str
    actual: Optional[Any] = None
    expected: Optional[Any] = None


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------
def analyze_citations(
    model: DocModel,
    roles: Roles,
    entries: list[ReferenceEntry],
) -> list[CitationFinding]:
    """检查正文引用标注与参考文献表的一致性（纯离线）。"""
    findings: list[CitationFinding] = []

    body_indices = _body_paragraph_indices(roles)
    occurrences = _collect_occurrences(model, body_indices)  # [(para_idx, number, is_sup|None)]

    entry_numbers: dict[int, ReferenceEntry] = {
        entry.number: entry for entry in entries if entry.number is not None
    }
    cited_numbers = {number for _, number, _ in occurrences}
    first_order = _first_appearance_order(occurrences)

    # ---- 0. 文献表本身为空 / 缺少序号 -----------------------------------
    if not entries:
        if occurrences:
            findings.append(
                CitationFinding(
                    level="error",
                    message=(
                        "正文存在 [n] 引用，但未解析到任何参考文献条目"
                        "（可能未识别到「参考文献」标题，或条目格式异常）"
                    ),
                    locator=f"P{occurrences[0][0]}",
                    actual=sorted(cited_numbers),
                    expected="能解析出的参考文献条目",
                )
            )
        return _sorted(findings)
    if not entry_numbers:
        findings.append(
            CitationFinding(
                level="info",
                message="参考文献条目均未识别出 [序号]，无法与正文引用编号逐条对应，需人工核对",
                locator="refs",
                actual=len(entries),
                expected="每条文献以 [n] 开头",
            )
        )

    # ---- 1. 正文引用编号不存在 / 无效 / 越界 ----------------------------
    max_number = max(entry_numbers) if entry_numbers else 0
    for number in sorted(cited_numbers):
        if number in entry_numbers:
            continue
        para_index = next(index for index, value, _ in occurrences if value == number)
        if number == 0:
            message = (
                "正文出现 [0] 引用：GB/T 7714 参考文献编号从 [1] 开始，[0] 无效"
            )
        elif entry_numbers and number > max_number:
            message = (
                f"正文引用 [{number}] 越界：参考文献表只有 {len(entry_numbers)} 条"
                f"（编号 1–{max_number}），不存在 [{number}]"
            )
        else:
            message = (
                f"正文引用 [{number}] 在参考文献表中不存在（缺条目，或正文编号与文献表不一致）"
            )
        findings.append(
            CitationFinding(
                level="error",
                message=message,
                locator=f"P{para_index}",
                actual=[number],
                expected=sorted(entry_numbers) or None,
            )
        )

    # ---- 2. 文献表中有、正文从未引用 -------------------------------------
    uncited = [
        entry.number
        for entry in entries
        if entry.number is not None and entry.number not in cited_numbers
    ]
    if uncited:
        locator_entry = next(
            entry for entry in entries if entry.number is not None and entry.number not in cited_numbers
        )
        findings.append(
            CitationFinding(
                level="warning",
                message=(
                    f"参考文献表中有 {len(uncited)} 条未被正文引用：{_format_numbers(uncited)}；"
                    "GB/T 7714 要求文献表只列正文实际引用的文献，请核对是否漏引或多余"
                ),
                locator=f"refs/P{locator_entry.paragraph_index}",
                actual=uncited,
                expected="文献表中每条都至少被正文引用一次",
            )
        )
    missing_number_entries = [entry for entry in entries if entry.number is None]
    if missing_number_entries:
        findings.append(
            CitationFinding(
                level="info",
                message=f"有 {len(missing_number_entries)} 条参考文献缺少 [序号]，无法核对引用关系",
                locator=f"refs/P{missing_number_entries[0].paragraph_index}",
                actual=len(missing_number_entries),
                expected="每条文献以 [n] 开头",
            )
        )

    # ---- 3. 首次引用顺序 ------------------------------------------------
    if first_order:
        actual_sequence = [number for number, _ in first_order]
        expected_sequence = list(range(1, len(actual_sequence) + 1))
        if actual_sequence != expected_sequence:
            mismatch = next(
                position
                for position, (actual, expected) in enumerate(
                    zip(actual_sequence, expected_sequence)
                )
                if actual != expected
            )
            para_index = first_order[mismatch][1]
            findings.append(
                CitationFinding(
                    level="warning",
                    message=(
                        "正文首次引用顺序与文献表编号不一致："
                        f"按引证先后应连续编号为 {expected_sequence}，"
                        f"实际首次出现顺序为 {actual_sequence}（GB/T 7714 要求按引证先后编号）"
                    ),
                    locator=f"P{para_index}",
                    actual=actual_sequence,
                    expected=expected_sequence,
                )
            )

    # ---- 4. 文献表编号连续性 -------------------------------------------
    entry_number_list = [entry.number for entry in entries]
    if entries and all(number is not None for number in entry_number_list):
        expected_entries = list(range(1, len(entries) + 1))
        if entry_number_list != expected_entries:
            findings.append(
                CitationFinding(
                    level="warning",
                    message=(
                        "参考文献表编号不连续或顺序错误："
                        f"实际 {_format_numbers(entry_number_list)}，"
                        f"应为 {_format_numbers(expected_entries)}"
                    ),
                    locator="refs",
                    actual=entry_number_list,
                    expected=expected_entries,
                )
            )

    # ---- 5. 正文引用是否上标 -------------------------------------------
    findings.extend(_superscript_findings(occurrences))

    return _sorted(findings)


# ---------------------------------------------------------------------------
# 采集引用
# ---------------------------------------------------------------------------
def _body_paragraph_indices(roles: Roles) -> list[int]:
    indices: list[int] = []
    for index, role in roles.by_index.items():
        if role == "body" or _HEADING_ROLE_RE.match(role or ""):
            indices.append(index)
    return sorted(indices)


def _collect_occurrences(
    model: DocModel, body_indices: list[int]
) -> list[tuple[int, int, Optional[bool]]]:
    document = _reopen_document(model)
    occurrences: list[tuple[int, int, Optional[bool]]] = []
    for index in body_indices:
        para = model.by_index.get(index)
        if para is None:
            continue
        text = para.text or ""
        if not text:
            continue
        spans = _superscript_spans(document, index, text)
        for match in _CITE_RE.finditer(text):
            numbers = _citation_numbers(match.group(1))
            superscript = _match_superscript(spans, match.start(), match.end())
            for number in numbers:
                occurrences.append((index, number, superscript))
    return occurrences


def _citation_numbers(content: str) -> list[int]:
    """把 ``[1,2]`` / ``[1-3]`` 的括号内容展开为编号列表。"""
    numbers: list[int] = []
    text = (content or "").translate(_FULLWIDTH_DIGITS)
    for part in re.split(r"[,，、]", text):
        part = part.strip()
        if not part:
            continue
        range_match = _RANGE_RE.match(part)
        if range_match:
            start = int(range_match.group(1))
            end = int(range_match.group(2))
            if 0 <= start <= end and end - start <= 19:  # 防止把 [2019-2024] 当范围展开
                numbers.extend(range(start, end + 1))
            continue
        if part.isdigit():
            numbers.append(int(part))
    return numbers


def _first_appearance_order(
    occurrences: list[tuple[int, int, Optional[bool]]],
) -> list[tuple[int, int]]:
    seen: set[int] = set()
    order: list[tuple[int, int]] = []
    for para_index, number, _ in occurrences:
        if number in seen:
            continue
        seen.add(number)
        order.append((number, para_index))
    return order


# ---------------------------------------------------------------------------
# 上标读取（直接读 docx XML，不改 docx_io）
# ---------------------------------------------------------------------------
def _reopen_document(model: DocModel) -> Any:
    """重新打开文档以便读 run 的 vertAlign；失败返回 None（降级跳过上标检查）。"""
    try:
        path = Path(model.path) if model.path is not None else None
        if path is None or not path.is_file():
            return None
        import docx

        return docx.Document(str(path))
    except Exception:
        return None


def _superscript_spans(
    document: Any, paragraph_index: int, expected_text: str
) -> Optional[list[tuple[int, int, bool]]]:
    """返回 ``[(start, end, is_superscript), ...]``；无法与段落文本对齐时返回 None。"""
    if document is None:
        return None
    try:
        paragraphs = document.paragraphs
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            return None
        paragraph = paragraphs[paragraph_index]
        from docx.oxml.ns import qn
        from docx.text.run import Run

        spans: list[tuple[int, int, bool]] = []
        cursor = 0
        parts: list[str] = []
        for run_element in paragraph._p.iter(qn("w:r")):
            run = Run(run_element, paragraph)
            run_text = run.text or ""
            if not run_text:
                continue
            parts.append(run_text)
            spans.append((cursor, cursor + len(run_text), _is_superscript(run_element)))
            cursor += len(run_text)
        if "".join(parts) != expected_text:
            # 超链接/域/脚注等导致 run 文本与段落文本不完全一致：放弃精确对齐
            return None
        return spans
    except Exception:
        return None


def _is_superscript(run_element: Any) -> bool:
    try:
        from docx.oxml.ns import qn

        rpr = run_element.rPr
        if rpr is None:
            return False
        vert_align = rpr.find(qn("w:vertAlign"))
        if vert_align is None:
            return False
        return (vert_align.get(qn("w:val")) or "").lower() == "superscript"
    except Exception:
        return False


def _match_superscript(
    spans: Optional[list[tuple[int, int, bool]]], start: int, end: int
) -> Optional[bool]:
    """判断一个引用区间是否整体上标；无法判断返回 None。"""
    if spans is None:
        return None
    covered = [is_superscript for span_start, span_end, is_superscript in spans if span_start < end and span_end > start]
    if not covered:
        return None
    return all(covered)


def _superscript_findings(
    occurrences: list[tuple[int, int, Optional[bool]]],
) -> list[CitationFinding]:
    by_paragraph: dict[int, list[int]] = {}
    for para_index, number, superscript in occurrences:
        if superscript is False:
            by_paragraph.setdefault(para_index, []).append(number)
    findings: list[CitationFinding] = []
    for para_index in sorted(by_paragraph):
        numbers = list(dict.fromkeys(by_paragraph[para_index]))
        findings.append(
            CitationFinding(
                level="warning",
                message=(
                    f"正文引用 {_format_numbers(numbers)} 未使用上标：GB/T 7714 要求序号置于"
                    "方括号内并以上角标形式标注（如“…[5]”）"
                ),
                locator=f"P{para_index}",
                actual=numbers,
                expected="上标（w:vertAlign=superscript）",
            )
        )
    return findings


# ---------------------------------------------------------------------------
# 展示辅助
# ---------------------------------------------------------------------------
def _format_numbers(numbers: list[int]) -> str:
    return "、".join(f"[{number}]" for number in numbers)


def _locator_key(locator: str) -> tuple:
    numbers = tuple(int(value) for value in re.findall(r"\d+", locator or "")[:3])
    # 第一层是数字元组、第二层是字符串，保证 "refs" 与 "P12" 可比较
    return numbers, locator or ""


def _sorted(findings: list[CitationFinding]) -> list[CitationFinding]:
    return sorted(
        findings,
        key=lambda finding: (
            _SEVERITY_ORDER.get(finding.level, 9),
            _locator_key(finding.locator),
        ),
    )
