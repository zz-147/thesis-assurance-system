"""名单核对：老师用「班级名单」核对提交文件的缺交 / 重复 / 命名不规范。

场景：老师从微信/群里收一堆学生文件（``23050301001_张三_论文.docx`` 之类），
手里有班级名单（txt / csv / xlsx）。本模块把两者对上，回答三个问题：

1. **谁没交**（``missing``）：名单里没有任何对应文件的学生；
2. **谁重复交**（``duplicates``）：同一学号命中多份文件（含 "(1)" / 副本 / copy）；
3. **哪些文件名不规范**（``unmatched_files`` / ``naming_warnings``）：匹配不到名单的文件、
   文件名缺姓名 / 缺学号 / 多余副本等。

设计约束：
- 纯函数、**只读**：不联网、不写盘；``match_submissions`` 甚至不访问文件系统
  （只按文件名判断，path 不存在也能核对）；
- 无账号、单机可用；供 CLI / Web / 小程序复用；
- 学号优先：文件名里能提取到 6–20 位连续数字就按学号匹配；没有学号才按姓名包含匹配；
- 学号始终以字符串保存（保留前导 0），不转 int。

典型用法::

    from tas.roster import parse_roster_file, match_submissions

    roster = parse_roster_file("名单.xlsx")
    report = match_submissions([p.name for p in upload_dir.iterdir()], roster)
    print(report["matched"], "人已交；", len(report["missing"]), "人缺交")
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Sequence

__all__ = [
    "RosterEntry",
    "STUDENT_ID_RE",
    "parse_roster",
    "parse_roster_file",
    "match_submissions",
    "extract_student_id",
    "is_copy_filename",
    "filename_name_candidates",
]

#: 学号：6–20 位连续数字（字符串，保留前导 0）
STUDENT_ID_RE = re.compile(r"^\d{6,20}$")
#: 从文件名中提取学号：左右不能紧邻其它数字（避免从更长的数字串里截一段）
_FILE_ID_RE = re.compile(r"(?<!\d)(\d{6,20})(?!\d)")
_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
#: "(1)" / "（1）" / 副本 / 复件 / copy 等“重复副本”标记
_COPY_MARK_RE = re.compile(r"\(\s*\d+\s*\)|（\s*\d+\s*）|副本|复件|复本|copy", re.IGNORECASE)
#: 文件名切词：非字母/数字/汉字都算分隔符
_TOKEN_SPLIT_RE = re.compile(r"[^0-9A-Za-z\u4e00-\u9fff]+")

#: 表头/栏目关键词（出现这些词且没有合法学号的行按表头跳过）
_HEADER_WORDS = frozenset(
    {
        "学号", "姓名", "名字", "序号", "编号", "班级", "专业", "学院", "学生",
        "备注", "提交", "文件", "列表", "名单", "id", "no", "no.", "name",
        "student", "studentid", "student_id", "学籍号",
    }
)
#: 文件名里出现这些词时不算“姓名”（如 论文 / 毕业设计 / 终稿 / 副本）
_NAME_STOPWORDS = (
    "论文", "毕业论文", "毕业设计", "设计", "说明书", "任务书", "开题报告",
    "终稿", "定稿", "初稿", "修改稿", "终版", "版本", "副本", "复件", "复本",
    "格式", "排版", "作业", "报告", "课程", "学生", "姓名", "学号", "文件",
    "文档", "附件", "提交", "final", "copy", "paper", "thesis", "docx", "doc",
)

# xlsx 表头识别关键词（小写子串匹配）
_XLSX_ID_HEADER_WORDS = ("学号", "学籍号", "student id", "studentid", "student_id", "编号")
_XLSX_NAME_HEADER_WORDS = ("姓名", "名字", "student name", "studentname", "student_name", "name")


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------
@dataclass
class RosterEntry:
    """名单中的一名学生（学号必须，姓名可空）。"""

    student_id: str
    name: str = ""

    def to_dict(self) -> dict[str, str]:
        return {"student_id": self.student_id, "name": self.name}


# ---------------------------------------------------------------------------
# 文本名单解析
# ---------------------------------------------------------------------------
def _clean_token(token: Any) -> str:
    """去掉首尾空白与引号（Excel/CSV 复制常见）。"""
    text = str(token if token is not None else "").strip()
    return text.strip("\"'“”‘’").strip()


def _tokenize(line: str) -> list[str]:
    """把一行拆成字段：兼容 逗号 / 制表符 / 空格 / 全角空格 / 分号 / CSV 引号。"""
    try:
        fields = next(csv.reader([line]))
    except (csv.Error, StopIteration):  # pragma: no cover - 极端畸形行兜底
        fields = [line]
    tokens: list[str] = []
    for field in fields:
        for token in re.split(r"[ \u3000\t]+", (field or "").strip()):
            token = _clean_token(token)
            if token:
                tokens.append(token)
    return tokens


def _is_student_id(token: str) -> bool:
    return bool(STUDENT_ID_RE.match(token or ""))


def _is_header_token(token: str) -> bool:
    return (token or "").strip().casefold() in _HEADER_WORDS


def _pick_name(tokens: Sequence[str], id_index: int) -> str:
    """取学号相邻的姓名：优先学号后一列，其次前一列（兼容 “姓名 学号”）。"""
    for index in (id_index + 1, id_index - 1):
        if 0 <= index < len(tokens):
            candidate = tokens[index]
            if not candidate or _is_student_id(candidate) or _is_header_token(candidate):
                continue
            return candidate
    return ""


def _append_entry(
    entries: list[RosterEntry],
    seen: dict[str, RosterEntry],
    student_id: str,
    name: str,
) -> None:
    """按学号去重：重复学号跳过；首次姓名为空时用后面的行补全。"""
    if student_id in seen:
        existing = seen[student_id]
        if not existing.name and name:
            existing.name = name
        return
    entry = RosterEntry(student_id=student_id, name=name)
    seen[student_id] = entry
    entries.append(entry)


def parse_roster(text: str) -> list[RosterEntry]:
    """解析名单文本（每行一名学生）。

    支持的写法::

        23050301001 张三          # 学号 + 空格 + 姓名
        23050301001,张三          # 学号 + 逗号 + 姓名（CSV）
        23050301001\t张三         # 学号 + 制表符 + 姓名
        张三 23050301001          # 姓名 + 学号（顺序可反）

    - 空行、表头（“学号/姓名/序号…”等没有合法学号的行）、无效学号行自动跳过；
    - 重复学号只保留一条（若第一条缺姓名，会用后续行补全）；
    - 姓名可空（只认学号）；学号按字符串保留前导 0。
    """
    entries: list[RosterEntry] = []
    seen: dict[str, RosterEntry] = {}
    for raw_line in (text or "").splitlines():
        line = (raw_line or "").strip().lstrip("\ufeff")
        if not line:
            continue
        tokens = _tokenize(line)
        id_index = next(
            (index for index, token in enumerate(tokens) if _is_student_id(token)), None
        )
        if id_index is None:
            continue  # 表头 / 说明行 / 无学号
        student_id = tokens[id_index]
        name = _pick_name(tokens, id_index)
        _append_entry(entries, seen, student_id, name)
    return entries


# ---------------------------------------------------------------------------
# 名单文件（txt / csv / xlsx）
# ---------------------------------------------------------------------------
def _read_text(path: Path) -> str:
    """读文本：优先 UTF-8（含 BOM），失败回退 GBK/GB18030。"""
    for encoding in ("utf-8-sig", "gbk", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def _cell_text(value: Any) -> str:
    """Excel 单元格 → 规范字符串（整数去掉 .0，空值 → ""）。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()


def _normalize_id(value: Any) -> str:
    """把单元格/字段规范成学号；不是 6–20 位数字则返回 ""。"""
    text = _cell_text(value).strip("\"'“”‘’").strip()
    if STUDENT_ID_RE.match(text):
        return text
    # 容错 “学号：23050301001” 这类单元格，取其中第一段 6–20 位数字
    match = _FILE_ID_RE.search(text)
    return match.group(1) if match else ""


def _xlsx_entries(path: Path) -> list[RosterEntry]:
    from openpyxl import load_workbook  # 延迟导入：未装 openpyxl 时只有读 xlsx 会报错

    workbook = load_workbook(str(path), read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0] if workbook.worksheets else None
        rows: list[tuple[Any, ...]] = (
            list(sheet.iter_rows(values_only=True)) if sheet is not None else []
        )
    finally:
        workbook.close()

    if not rows:
        return []

    width = max(len(row) for row in rows)
    header = _find_xlsx_header(rows)
    if header is not None:
        header_index, id_col, name_col = header
        data_rows = rows[header_index + 1 :]
    else:
        id_col, name_col = _detect_xlsx_columns(rows, width)
        data_rows = rows
        if id_col is None:
            return []

    entries: list[RosterEntry] = []
    seen: dict[str, RosterEntry] = {}
    for row in data_rows:
        student_id = _normalize_id(row[id_col]) if id_col < len(row) else ""
        if not student_id:
            continue
        name = ""
        if name_col is not None and name_col < len(row):
            name = _cell_text(row[name_col]).strip()
            if _is_header_token(name) or _is_student_id(name):
                name = ""
        _append_entry(entries, seen, student_id, name)
    return entries


def _find_xlsx_header(
    rows: Sequence[Sequence[Any]],
) -> Optional[tuple[int, int, Optional[int]]]:
    """在前 8 行里找表头；返回 (表头行号, 学号列, 姓名列或 None)。"""
    for row_index, row in enumerate(rows[:8]):
        cells = [_cell_text(value).casefold() for value in row]
        id_col = next(
            (
                column
                for column, cell in enumerate(cells)
                if any(word in cell for word in _XLSX_ID_HEADER_WORDS)
            ),
            None,
        )
        name_col = next(
            (
                column
                for column, cell in enumerate(cells)
                if any(word in cell for word in _XLSX_NAME_HEADER_WORDS)
            ),
            None,
        )
        if id_col is not None:
            return row_index, id_col, name_col
    return None


def _detect_xlsx_columns(
    rows: Sequence[Sequence[Any]], width: int
) -> tuple[Optional[int], Optional[int]]:
    """无表头时自动判断学号列/姓名列（优先前两列）。"""
    id_counts = [0] * width
    name_counts = [0] * width
    for row in rows[:200]:
        for column in range(min(len(row), width)):
            text = _cell_text(row[column])
            if _normalize_id(text):
                id_counts[column] += 1
            elif _CJK_RE.search(text):
                name_counts[column] += 1

    if not any(id_counts):
        return None, None

    first_two = [column for column in (0, 1) if column < width and id_counts[column] > 0]
    if first_two:
        id_col = max(first_two, key=lambda column: id_counts[column])
    else:
        id_col = max(range(width), key=lambda column: id_counts[column])

    name_col: Optional[int] = None
    for preferred in ((id_col + 1, id_col - 1, 0, 1)):
        if 0 <= preferred < width and preferred != id_col and name_counts[preferred] > 0:
            name_col = preferred
            break
    return id_col, name_col


def parse_roster_file(path: str | Path) -> list[RosterEntry]:
    """读取名单文件：``.txt`` / ``.csv``（UTF-8，容错 GBK）/ ``.xlsx``。

    - ``.xlsx`` 读第一个 sheet：首行是表头时自动找「学号 / 姓名」列；
      无表头时自动判断（优先前两列）。
    - 其它后缀按文本处理（方便直接读 .md / .tsv）。
    - 文件不存在抛 ``FileNotFoundError``；旧版 ``.xls`` 明确报错（请另存为 .xlsx）。
    """
    file_path = Path(path)
    if not file_path.is_file():
        raise FileNotFoundError(f"名单文件不存在：{file_path}")
    suffix = file_path.suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".xltx"):
        return _xlsx_entries(file_path)
    if suffix == ".xls":
        raise ValueError("暂不支持旧版 .xls 名单，请用 Excel 另存为 .xlsx")
    return parse_roster(_read_text(file_path))


# ---------------------------------------------------------------------------
# 文件名解析
# ---------------------------------------------------------------------------
def _basename(path: str | Path) -> str:
    return Path(str(path)).name


def extract_student_id(filename: str | Path) -> Optional[str]:
    """从文件名里提取 6–20 位连续数字作为学号；没有则返回 None。"""
    stem = Path(_basename(filename)).stem
    match = _FILE_ID_RE.search(stem)
    return match.group(1) if match else None


def is_copy_filename(filename: str | Path) -> bool:
    """文件名是否带“重复副本”标记（(1) / 副本 / copy 等）。"""
    return bool(_COPY_MARK_RE.search(Path(_basename(filename)).stem))


def filename_name_candidates(filename: str | Path) -> list[str]:
    """从文件名里提取“像姓名”的片段（用于判断文件名是否缺姓名）。"""
    stem = Path(_basename(filename)).stem
    candidates: list[str] = []
    for token in _TOKEN_SPLIT_RE.split(stem):
        if not token or token.isdigit():
            continue
        if not _CJK_RE.search(token) and not token.isalpha():
            continue
        lowered = token.casefold()
        if lowered in _HEADER_WORDS:
            continue
        if any(stopword in lowered for stopword in _NAME_STOPWORDS):
            continue
        candidates.append(token)
    return candidates


def _has_name_in_filename(filename: str | Path, roster_name: str) -> bool:
    """文件名是否含该学生姓名：优先精确包含名单姓名，其次看“像姓名”的片段。"""
    basename = _basename(filename)
    if roster_name and roster_name.casefold() in basename.casefold():
        return True
    return bool(filename_name_candidates(filename))


def _match_by_name(
    filename: str | Path, entries: Sequence[RosterEntry]
) -> tuple[Optional[RosterEntry], bool]:
    """按姓名匹配：返回 (命中条目 | None, 是否歧义)。取最长姓名，降低子串误配。"""
    basename = _basename(filename).casefold()
    hits = [entry for entry in entries if entry.name and entry.name.casefold() in basename]
    if not hits:
        return None, False
    hits.sort(key=lambda entry: len(entry.name), reverse=True)
    if len(hits) > 1 and len(hits[0].name) == len(hits[1].name):
        return None, True  # 同名/长度相同 → 无法唯一确定
    return hits[0], False


# ---------------------------------------------------------------------------
# 名单 ↔ 文件核对
# ---------------------------------------------------------------------------
def _entry_fields(entry: Any) -> tuple[str, str]:
    """兼容 RosterEntry / dict 两种名单条目。"""
    if isinstance(entry, RosterEntry):
        return entry.student_id, entry.name
    if isinstance(entry, dict):
        return str(entry.get("student_id") or ""), str(entry.get("name") or "")
    return str(getattr(entry, "student_id", "") or ""), str(getattr(entry, "name", "") or "")


def match_submissions(
    paths: Sequence[str | Path], roster: Sequence[Any]
) -> dict[str, Any]:
    """把提交文件与名单核对，返回缺交 / 重复 / 未匹配 / 命名不规范。

    返回字段::

        {
          "total_files": int,          # 文件总数（= len(paths)）
          "total_roster": int,         # 名单人数（= len(roster)）
          "matched": int,              # 名单中“已匹配到至少一份文件”的人数
          "missing": [{"student_id","name"}],             # 缺交
          "duplicates": [{"student_id","name","files"}],  # 同一学号多份文件
          "unmatched_files": [{"file","reason"}],         # 文件匹配不到名单
          "naming_warnings": [{"file","student_id","name","reason"}],
          "entries": [{"file","student_id","name","matched"}],  # 逐文件
        }

    匹配规则（学号优先）：
    1. 文件名里能提取 6–20 位连续数字 → 只按学号匹配（学号不在名单 → unmatched）；
    2. 文件名没有学号 → 按“文件名包含名单姓名”匹配（可用姓名必须唯一）；
    3. 同一学号命中多份文件 → ``duplicates``；其中带 (1)/副本/copy 的文件记
       “多余副本”命名警告；
    4. 有学号但文件名不含姓名 → “文件名缺少姓名”；无学号按姓名匹配 →
       “文件名缺少学号”；
    5. ``file`` 字段为文件名（basename），不泄漏目录路径；本函数不访问文件系统。
    """
    entries_by_id: dict[str, RosterEntry] = {}
    roster_entries: list[tuple[str, str]] = []
    for entry in roster or []:
        student_id, name = _entry_fields(entry)
        roster_entries.append((student_id, name))
        if student_id and student_id not in entries_by_id:
            entries_by_id[student_id] = RosterEntry(student_id=student_id, name=name)
    name_list = list(entries_by_id.values())

    file_entries: list[dict[str, Any]] = []
    unmatched_files: list[dict[str, str]] = []
    # matched_student_id → [文件记录...]，按首次出现顺序
    hits_by_student: dict[str, list[dict[str, Any]]] = {}
    matched_ids: set[str] = set()

    for raw_path in paths or []:
        basename = _basename(raw_path)
        student_id = extract_student_id(basename)
        record: dict[str, Any] = {
            "file": basename,
            "student_id": "",
            "name": "",
            "matched": False,
        }

        if student_id is not None:
            # -- 学号优先 --
            hit = entries_by_id.get(student_id)
            if hit is None:
                unmatched_files.append(
                    {"file": basename, "reason": f"学号 {student_id} 不在名单中"}
                )
                file_entries.append(record)
                continue
            record.update({"student_id": hit.student_id, "name": hit.name, "matched": True})
            record["_matched_by"] = "id"
        else:
            # -- 无学号：姓名包含匹配 --
            hit, ambiguous = _match_by_name(basename, name_list)
            if hit is None:
                reason = (
                    "文件名包含多个名单姓名，无法唯一确定学号"
                    if ambiguous
                    else "文件名中未找到学号，也未匹配到名单姓名"
                )
                unmatched_files.append({"file": basename, "reason": reason})
                file_entries.append(record)
                continue
            record.update(
                {"student_id": hit.student_id, "name": hit.name, "matched": True}
            )
            record["_matched_by"] = "name"

        matched_ids.add(record["student_id"])
        hits_by_student.setdefault(record["student_id"], []).append(record)
        file_entries.append(record)

    # -- 重复提交 --
    duplicates: list[dict[str, Any]] = []
    duplicate_ids: set[str] = set()
    for student_id, records in hits_by_student.items():
        if len(records) < 2:
            continue
        duplicate_ids.add(student_id)
        duplicates.append(
            {
                "student_id": student_id,
                "name": records[0]["name"],
                "files": [item["file"] for item in records],
            }
        )

    # -- 命名警告（逐文件，顺序与输入一致）--
    naming_warnings: list[dict[str, str]] = []
    for record in file_entries:
        if not record["matched"]:
            continue
        student_id = record["student_id"]
        name = record["name"]
        if student_id in duplicate_ids and is_copy_filename(record["file"]):
            naming_warnings.append(
                {
                    "file": record["file"],
                    "student_id": student_id,
                    "name": name,
                    "reason": "多余副本（同一学号重复提交）",
                }
            )
        if not _has_name_in_filename(record["file"], name):
            naming_warnings.append(
                {
                    "file": record["file"],
                    "student_id": student_id,
                    "name": name,
                    "reason": "文件名缺少姓名",
                }
            )
        if record.get("_matched_by") == "name":
            naming_warnings.append(
                {
                    "file": record["file"],
                    "student_id": student_id,
                    "name": name,
                    "reason": "文件名缺少学号（按姓名匹配）",
                }
            )

    # -- 缺交（按名单顺序，重复学号只报一次）--
    missing: list[dict[str, str]] = []
    missing_ids: set[str] = set()
    for student_id, name in roster_entries:
        if not student_id or student_id in matched_ids or student_id in missing_ids:
            continue
        missing_ids.add(student_id)
        missing.append({"student_id": student_id, "name": name})

    entries = [
        {
            "file": record["file"],
            "student_id": record["student_id"],
            "name": record["name"],
            "matched": bool(record["matched"]),
        }
        for record in file_entries
    ]

    return {
        "total_files": len(paths) if paths is not None else 0,
        "total_roster": len(roster) if roster is not None else 0,
        "matched": len(matched_ids),
        "missing": missing,
        "duplicates": duplicates,
        "unmatched_files": unmatched_files,
        "naming_warnings": naming_warnings,
        "entries": entries,
    }
