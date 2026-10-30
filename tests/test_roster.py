"""名单核对模块（tas.roster）测试：名单解析 + 缺交/重复/命名不规范匹配。

覆盖需求：
1. parse_roster：空格 / 逗号 / 制表符 / 姓名在前 四种写法、表头跳过、重复学号去重、
   学号 6–20 位边界、姓名可空；
2. parse_roster_file：txt（UTF-8）、csv、gbk 容错、xlsx（有表头 / 无表头 / 姓名在前，
   无 openpyxl 时 skip）；
3. match_submissions：名单 3 人 + 4 个文件的缺交 / 重复 / 未匹配 / 命名警告；
   姓名匹配、缺姓名警告、未知学号、空边界、数量不变量；
4. 纯函数：match_submissions 不访问文件系统（路径不存在也能核对）。
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from tas.roster import (  # noqa: E402
    RosterEntry,
    extract_student_id,
    is_copy_filename,
    match_submissions,
    parse_roster,
    parse_roster_file,
)

try:  # xlsx 测试依赖 openpyxl（环境已装；缺失时相关用例 skip）
    import openpyxl  # noqa: F401

    HAS_OPENPYXL = True
except ImportError:  # pragma: no cover
    HAS_OPENPYXL = False


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------
def ids(entries) -> list[str]:
    return [entry.student_id for entry in entries]


def names(entries) -> list[str]:
    return [entry.name for entry in entries]


def make_xlsx(path: Path, rows: list[list[object]]) -> Path:
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "名单"
    for row in rows:
        sheet.append(row)
    workbook.save(str(path))
    return path


ROSTER = [
    RosterEntry("23050301001", "张三"),
    RosterEntry("23050301002", "李四"),
    RosterEntry("23050301003", "王五"),
]


# ---------------------------------------------------------------------------
# 1) RosterEntry
# ---------------------------------------------------------------------------
class RosterEntryTest(unittest.TestCase):
    def test_to_dict(self):
        self.assertEqual(
            RosterEntry("23050301001", "张三").to_dict(),
            {"student_id": "23050301001", "name": "张三"},
        )
        self.assertEqual(
            RosterEntry("23050301001").to_dict(),
            {"student_id": "23050301001", "name": ""},
        )


# ---------------------------------------------------------------------------
# 2) parse_roster
# ---------------------------------------------------------------------------
class ParseRosterTest(unittest.TestCase):
    def test_space_separator(self):
        entries = parse_roster("23050301001 张三\n23050301002 李四")
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])

    def test_comma_separator(self):
        entries = parse_roster("23050301001,张三\n23050301002,李四")
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])

    def test_tab_separator(self):
        entries = parse_roster("23050301001\t张三\n23050301002\t李四")
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])

    def test_name_first_order(self):
        entries = parse_roster("张三 23050301001\n李四\t23050301002")
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])

    def test_skips_headers_blank_lines_and_bom(self):
        text = (
            "\ufeff学号 姓名\n"
            "序号,学号,姓名,班级\n"
            "姓名 学号 班级\n"
            "\n"
            "   \n"
            "23050301001 张三\n"
            "23050301002 李四\n"
        )
        entries = parse_roster(text)
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])

    def test_skips_duplicate_student_ids(self):
        entries = parse_roster(
            "23050301001 张三\n23050301001 张三（重复）\n23050301002 李四"
        )
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        # 首次姓名为空时，用后续重复行补全
        filled = parse_roster("23050301001\n23050301001 张三")
        self.assertEqual(len(filled), 1)
        self.assertEqual(filled[0].name, "张三")

    def test_student_id_length_bounds(self):
        entries = parse_roster(
            "123456,A\n"          # 6 位：合法
            "12345678901234567890,B\n"  # 20 位：合法
            "12345,C\n"           # 5 位：不是学号
            "123456789012345678901,D\n"  # 21 位：不是学号
        )
        self.assertEqual(ids(entries), ["123456", "12345678901234567890"])

    def test_name_optional(self):
        entries = parse_roster("23050301001\n23050301002 李四")
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(entries[0].name, "")
        self.assertEqual(entries[1].name, "李四")

    def test_leading_zeros_preserved(self):
        entries = parse_roster("00123456 张三")
        self.assertEqual(ids(entries), ["00123456"])

    def test_empty_and_garbage_text(self):
        self.assertEqual(parse_roster(""), [])
        self.assertEqual(parse_roster("\n\n  \n"), [])
        self.assertEqual(parse_roster("姓名 学号\n备注：请按时提交"), [])

    def test_numbering_column_before_id(self):
        entries = parse_roster("1 23050301001 张三\n2 23050301002 李四")
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])


# ---------------------------------------------------------------------------
# 3) parse_roster_file
# ---------------------------------------------------------------------------
class ParseRosterFileTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def test_txt_utf8(self):
        path = self.tmp / "roster.txt"
        path.write_text("学号,姓名\n23050301001,张三\n23050301002,李四\n", encoding="utf-8")
        entries = parse_roster_file(path)
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])

    def test_csv(self):
        path = self.tmp / "roster.csv"
        path.write_text("23050301001,张三\n23050301002,李四\n", encoding="utf-8")
        self.assertEqual(len(parse_roster_file(path)), 2)

    def test_gbk_fallback(self):
        path = self.tmp / "roster_gbk.csv"
        path.write_bytes("学号,姓名\n23050301001,张三\n".encode("gbk"))
        entries = parse_roster_file(path)
        self.assertEqual(ids(entries), ["23050301001"])
        self.assertEqual(names(entries), ["张三"])

    def test_missing_file_raises(self):
        with self.assertRaises(FileNotFoundError):
            parse_roster_file(self.tmp / "nope.xlsx")

    def test_old_xls_raises_value_error(self):
        path = self.tmp / "old.xls"
        path.write_bytes(b"not really an xls")
        with self.assertRaises(ValueError):
            parse_roster_file(path)

    @unittest.skipUnless(HAS_OPENPYXL, "需要 openpyxl 读取 .xlsx")
    def test_xlsx_with_header(self):
        path = make_xlsx(
            self.tmp / "with_header.xlsx",
            [
                ["序号", "学号", "姓名", "班级"],
                [1, 23050301001, "张三", "计科2301"],
                [2, "23050301002", "李四", "计科2301"],
                [3, "23050301003", "王五", "计科2302"],
                [4, "23050301001", "张三", "计科2301"],  # 重复学号
            ],
        )
        entries = parse_roster_file(path)
        self.assertEqual(ids(entries), ["23050301001", "23050301002", "23050301003"])
        self.assertEqual(names(entries), ["张三", "李四", "王五"])

    @unittest.skipUnless(HAS_OPENPYXL, "需要 openpyxl 读取 .xlsx")
    def test_xlsx_without_header(self):
        path = make_xlsx(
            self.tmp / "no_header.xlsx",
            [["23050301001", "张三"], ["23050301002", "李四"]],
        )
        entries = parse_roster_file(path)
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])

    @unittest.skipUnless(HAS_OPENPYXL, "需要 openpyxl 读取 .xlsx")
    def test_xlsx_name_first_columns(self):
        # 无表头且姓名在前：按“哪一列像学号”自动判断
        path = make_xlsx(
            self.tmp / "name_first.xlsx",
            [["张三", "23050301001"], ["李四", "23050301002"]],
        )
        entries = parse_roster_file(path)
        self.assertEqual(ids(entries), ["23050301001", "23050301002"])
        self.assertEqual(names(entries), ["张三", "李四"])

    @unittest.skipUnless(HAS_OPENPYXL, "需要 openpyxl 读取 .xlsx")
    def test_xlsx_empty(self):
        from openpyxl import Workbook

        path = self.tmp / "empty.xlsx"
        Workbook().save(str(path))
        self.assertEqual(parse_roster_file(path), [])


# ---------------------------------------------------------------------------
# 4) 文件名工具
# ---------------------------------------------------------------------------
class FilenameHelperTest(unittest.TestCase):
    def test_extract_student_id(self):
        self.assertEqual(extract_student_id("23050301001_张三_论文.docx"), "23050301001")
        self.assertEqual(extract_student_id("张三-23050301002.docx"), "23050301002")
        self.assertIsNone(extract_student_id("张三_论文.docx"))
        self.assertIsNone(extract_student_id("2023-2024学年总结.docx"))  # 无 6 位连续数字
        self.assertIsNone(extract_student_id("123456789012345678901_张三.docx"))  # 21 位

    def test_is_copy_filename(self):
        self.assertTrue(is_copy_filename("23050301002-李四(1).docx"))
        self.assertTrue(is_copy_filename("23050301002-李四（1）.docx"))
        self.assertTrue(is_copy_filename("23050301002-李四-副本.docx"))
        self.assertTrue(is_copy_filename("23050301002-李四-copy.docx"))
        self.assertFalse(is_copy_filename("23050301002-李四.docx"))
        self.assertFalse(is_copy_filename("李四-初稿.docx"))  # 初稿不是副本标记


# ---------------------------------------------------------------------------
# 5) match_submissions
# ---------------------------------------------------------------------------
class MatchSubmissionsTest(unittest.TestCase):
    def test_contract_example(self):
        """名单 3 人 + 4 个文件：张三匹配 / 李四重复 / 王五缺交 / 其他未匹配。"""
        files = [
            "23050301001_张三_论文.docx",
            "23050301002-李四.docx",
            "23050301002-李四(1).docx",
            "其他.docx",
        ]
        report = match_submissions(files, ROSTER)

        self.assertEqual(report["total_files"], 4)
        self.assertEqual(report["total_roster"], 3)
        self.assertEqual(report["matched"], 2)  # 张三、李四（李四按 1 人计）

        # 缺交：王五
        self.assertEqual(report["missing"], [{"student_id": "23050301003", "name": "王五"}])

        # 重复：李四 ×2，文件顺序与输入一致
        self.assertEqual(len(report["duplicates"]), 1)
        duplicate = report["duplicates"][0]
        self.assertEqual(duplicate["student_id"], "23050301002")
        self.assertEqual(duplicate["name"], "李四")
        self.assertEqual(
            duplicate["files"],
            ["23050301002-李四.docx", "23050301002-李四(1).docx"],
        )

        # 未匹配：其他.docx
        self.assertEqual(len(report["unmatched_files"]), 1)
        self.assertEqual(report["unmatched_files"][0]["file"], "其他.docx")
        self.assertIn("未找到学号", report["unmatched_files"][0]["reason"])

        # 命名警告：带 (1) 的副本
        copy_warnings = [
            item for item in report["naming_warnings"] if item["file"].endswith("(1).docx")
        ]
        self.assertEqual(len(copy_warnings), 1)
        self.assertIn("多余副本", copy_warnings[0]["reason"])
        self.assertEqual(copy_warnings[0]["student_id"], "23050301002")

        # 逐文件 entries
        entries = report["entries"]
        self.assertEqual(len(entries), 4)
        by_file = {item["file"]: item for item in entries}
        self.assertTrue(by_file["23050301001_张三_论文.docx"]["matched"])
        self.assertEqual(by_file["23050301001_张三_论文.docx"]["student_id"], "23050301001")
        self.assertTrue(by_file["23050301002-李四.docx"]["matched"])
        self.assertFalse(by_file["其他.docx"]["matched"])
        self.assertEqual(by_file["其他.docx"]["student_id"], "")

    def test_name_only_file_matched_with_warning(self):
        report = match_submissions(["张三_毕业论文.docx"], ROSTER)
        self.assertEqual(report["matched"], 1)
        entry = report["entries"][0]
        self.assertTrue(entry["matched"])
        self.assertEqual(entry["student_id"], "23050301001")
        warnings = {(item["file"], item["reason"]) for item in report["naming_warnings"]}
        self.assertTrue(
            any("缺少学号" in reason for _file, reason in warnings), report["naming_warnings"]
        )

    def test_file_with_id_but_no_name_warns(self):
        report = match_submissions(["23050301001.docx"], ROSTER)
        self.assertTrue(report["entries"][0]["matched"])
        reasons = [item["reason"] for item in report["naming_warnings"]]
        self.assertIn("文件名缺少姓名", reasons)

    def test_unknown_student_id_is_unmatched(self):
        report = match_submissions(["99999999999_张三.docx"], ROSTER)
        self.assertEqual(report["matched"], 0)
        self.assertEqual(len(report["unmatched_files"]), 1)
        self.assertIn("不在名单中", report["unmatched_files"][0]["reason"])
        # 未知学号不会退化成姓名匹配（学号优先）
        self.assertEqual(report["entries"][0]["student_id"], "")

    def test_all_matched_no_warnings(self):
        files = [
            "23050301001_张三.docx",
            "23050301002_李四.docx",
            "23050301003_王五.docx",
        ]
        report = match_submissions(files, ROSTER)
        self.assertEqual(report["matched"], 3)
        self.assertEqual(report["missing"], [])
        self.assertEqual(report["duplicates"], [])
        self.assertEqual(report["unmatched_files"], [])
        self.assertEqual(report["naming_warnings"], [])

    def test_copy_marker_variants(self):
        files = [
            "23050301002_李四.docx",
            "23050301002_李四_副本.docx",
            "23050301002_李四_copy.docx",
        ]
        report = match_submissions(files, ROSTER)
        self.assertEqual(len(report["duplicates"]), 1)
        self.assertEqual(len(report["duplicates"][0]["files"]), 3)
        copy_files = {item["file"] for item in report["naming_warnings"]}
        self.assertIn("23050301002_李四_副本.docx", copy_files)
        self.assertIn("23050301002_李四_copy.docx", copy_files)
        self.assertNotIn("23050301002_李四.docx", copy_files)

    def test_ambiguous_name_is_unmatched(self):
        roster = [
            RosterEntry("23050301001", "张三"),
            RosterEntry("23050301002", "张三"),
        ]
        report = match_submissions(["张三_论文.docx"], roster)
        self.assertEqual(report["matched"], 0)
        self.assertEqual(len(report["unmatched_files"]), 1)
        self.assertIn("多个名单姓名", report["unmatched_files"][0]["reason"])

    def test_count_invariants(self):
        files = [
            "23050301001_张三_论文.docx",
            "23050301002-李四.docx",
            "23050301002-李四(1).docx",
            "其他.docx",
        ]
        report = match_submissions(files, ROSTER)
        self.assertEqual(len(report["entries"]), report["total_files"])
        self.assertEqual(
            report["matched"] + len(report["missing"]), report["total_roster"]
        )
        matched_entries = [item for item in report["entries"] if item["matched"]]
        self.assertEqual(report["total_files"] - len(report["unmatched_files"]), len(matched_entries))

    def test_empty_roster_and_files(self):
        empty = match_submissions([], [])
        self.assertEqual(
            empty,
            {
                "total_files": 0,
                "total_roster": 0,
                "matched": 0,
                "missing": [],
                "duplicates": [],
                "unmatched_files": [],
                "naming_warnings": [],
                "entries": [],
            },
        )
        only_files = match_submissions(["随便.docx"], [])
        self.assertEqual(only_files["total_files"], 1)
        self.assertEqual(len(only_files["unmatched_files"]), 1)
        only_roster = match_submissions([], ROSTER)
        self.assertEqual(only_roster["total_roster"], 3)
        self.assertEqual(len(only_roster["missing"]), 3)

    def test_pure_function_accepts_nonexistent_paths_and_uses_basename(self):
        # 不存在的路径也能核对（match_submissions 不访问文件系统）；file 只给文件名
        files = [Path("not/on/disk/23050301001_张三.docx")]
        report = match_submissions(files, ROSTER)
        self.assertEqual(report["entries"][0]["file"], "23050301001_张三.docx")
        self.assertTrue(report["entries"][0]["matched"])

    def test_roster_accepts_dict_entries(self):
        roster = [
            {"student_id": "23050301001", "name": "张三"},
            {"student_id": "23050301002", "name": "李四"},
        ]
        report = match_submissions(["23050301001_张三.docx"], roster)
        self.assertEqual(report["total_roster"], 2)
        self.assertEqual(report["matched"], 1)
        self.assertEqual(report["missing"], [{"student_id": "23050301002", "name": "李四"}])


if __name__ == "__main__":
    unittest.main(verbosity=2)
