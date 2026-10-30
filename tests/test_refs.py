"""文献解析 / 引用一致性 / 真实性核验的单元测试（默认不联网）。

覆盖：
- parse_references：学院规范示例格式（期刊/图书/学位论文）字段抽取；EB/OL、DOI、URL；
  解析失败不抛异常、title 尽量非空；角色识别失败时的 [n] 兜底。
- analyze_citations：合规 fixture 零问题；未引用、越界、[0]、乱序、缺号、非上标。
- verify_reference：注入假 fetcher 覆盖 verified / likely / not_found / error 四态；
  not_found 文案必须含「需人工核实」且附证据 URL；网络错误不得伪装成 not_found。
- live 网络测试：仅当环境变量 TAS_LIVE_NETWORK=1 才执行（默认 skip）。

不写学生真实文档；fixture 全部来自学院规范《撰写规范》表 3 的著录示例 / 公开文献。
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
import urllib.error
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import docx  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from tas.config import DEFAULT_PROFILE  # noqa: E402
from tas.docx_io import DocModel  # noqa: E402
from tas.refs import (  # noqa: E402
    LIKELY_THRESHOLD,
    VERIFIED_THRESHOLD,
    analyze_citations,
    default_fetcher,
    parse_reference,
    parse_references,
    title_similarity,
    verify_reference,
    verify_references,
)
from tas.roles import Roles, classify  # noqa: E402

# ---------------------------------------------------------------------------
# 学院规范表 3 的参考文献示例（全角标点，原样照抄，用于 fixture）
# ---------------------------------------------------------------------------
REFS = [
    "[1] 高景德，王祥珩．交流电机的多回路理论[J]．清华大学学报，2017，27（1）：1－8．",
    "[2] 竺可桢．物理学[M]．北京：科学出版社，2015．",
    "[3] 张竹生．微分半动力系统的不变集[D]．北京：北京大学数学系，2016．",
]

# 正文：三条文献各引用一次，且引用序号全部上标（合规写法）
DEFAULT_BODY = [
    [
        ("相关研究见文献", False),
        ("[1]", True),
        ("与文献", False),
        ("[2]", True),
        ("；文献", False),
        ("[3]", True),
        ("给出了方法。", False),
    ],
]

OPTUNA_REF = (
    "[1] Akiba T, Sano S, Yanase T, et al. Optuna: A Next-generation Hyperparameter "
    "Optimization Framework[C]. Proceedings of the 25th ACM SIGKDD International Conference "
    "on Knowledge Discovery & Data Mining, 2019."
)


# ---------------------------------------------------------------------------
# fixture 构造
# ---------------------------------------------------------------------------
def _add_run(paragraph, text, *, superscript=False):
    run = paragraph.add_run(text)
    rpr = run._r.get_or_add_rPr()
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:eastAsia"), "宋体")
    rpr.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "24")
    rpr.append(size)
    if superscript:
        vertical = OxmlElement("w:vertAlign")
        vertical.set(qn("w:val"), "superscript")
        rpr.append(vertical)
    return run


def build_docx(path: Path, *, refs=None, body=None) -> Path:
    """构造一份含前置区/正文/参考文献的最小 docx（不涉及真实学生文档）。"""
    refs = REFS if refs is None else refs
    body = DEFAULT_BODY if body is None else body
    document = docx.Document()
    document.add_paragraph("基于文献核验的测试论文")
    document.add_paragraph("摘  要")
    document.add_paragraph("本文用于测试文献解析与引用一致性，不涉及真实学生文档。")
    document.add_paragraph("第1章 绪论", style="Heading 1")
    for pieces in body:
        paragraph = document.add_paragraph()
        for text, superscript in pieces:
            _add_run(paragraph, text, superscript=superscript)
    document.add_paragraph("参考文献")
    for reference in refs:
        paragraph = document.add_paragraph()
        _add_run(paragraph, reference)
    document.save(str(path))
    return path


def load_fixture(testcase, **kwargs):
    tmp = tempfile.TemporaryDirectory()
    testcase.addCleanup(tmp.cleanup)
    path = build_docx(Path(tmp.name) / "refs.docx", **kwargs)
    model = DocModel.load(path)
    roles = classify(model, DEFAULT_PROFILE)
    return model, roles


# ---------------------------------------------------------------------------
# 1. 解析
# ---------------------------------------------------------------------------
class ParseReferencesTest(unittest.TestCase):
    def test_roles_recognize_references_section(self):
        model, roles = load_fixture(self)
        self.assertIsNotNone(roles.first_index("refs_title"))
        self.assertEqual(len(roles.indices("reference_entry")), 3)

    def test_extract_journal_book_thesis_fields(self):
        model, roles = load_fixture(self)
        entries = parse_references(model, roles)
        self.assertEqual(len(entries), 3)

        journal, book, thesis = entries
        # 期刊
        self.assertEqual(journal.number, 1)
        self.assertEqual(journal.doc_type, "J")
        self.assertEqual(journal.year, 2017)
        self.assertIn("高景德", journal.authors)
        self.assertIn("王祥珩", journal.authors)
        self.assertEqual(journal.title, "交流电机的多回路理论")
        self.assertEqual(journal.venue, "清华大学学报")
        self.assertTrue(journal.raw.startswith("[1]"))
        # 图书
        self.assertEqual(book.number, 2)
        self.assertEqual(book.doc_type, "M")
        self.assertEqual(book.year, 2015)
        self.assertEqual(book.authors, "竺可桢")
        self.assertEqual(book.title, "物理学")
        self.assertIn("科学出版社", book.venue)
        # 学位论文
        self.assertEqual(thesis.number, 3)
        self.assertEqual(thesis.doc_type, "D")
        self.assertEqual(thesis.year, 2016)
        self.assertEqual(thesis.title, "微分半动力系统的不变集")
        self.assertIn("北京大学数学系", thesis.venue)

        # paragraph_index 必须能回到原文段落
        for entry in entries:
            self.assertEqual(model.by_index[entry.paragraph_index].text, entry.raw)

    def test_extract_electronic_resource_url(self):
        raw = (
            "[14] 王明亮．标准化数据库系统工程新进展[EB/OL]．"
            "http://www.cajcd.edu.cn/pub/980810-2.html，2018-08-16．"
        )
        entry = parse_reference(7, raw)
        self.assertEqual(entry.number, 14)
        self.assertEqual(entry.doc_type, "EB/OL")
        self.assertEqual(entry.year, 2018)
        self.assertEqual(entry.title, "标准化数据库系统工程新进展")
        self.assertEqual(entry.authors, "王明亮")
        self.assertEqual(entry.url, "http://www.cajcd.edu.cn/pub/980810-2.html")

    def test_extract_doi(self):
        raw = "[1] 张三. 题名[J]. 测试学报, 2020, 1(1): 1-8. DOI: 10.1234/abcd.5678."
        entry = parse_reference(0, raw)
        self.assertEqual(entry.doi, "10.1234/abcd.5678")
        self.assertEqual(entry.title, "题名")

    def test_fullwidth_type_identifier(self):
        entry = parse_reference(0, "[5] 李四．题名［Ｍ］．上海：复旦大学出版社，2019．")
        self.assertEqual(entry.doc_type, "M")
        self.assertEqual(entry.year, 2019)
        self.assertIn("复旦大学出版社", entry.venue)

    def test_parse_failure_never_raises_and_keeps_raw(self):
        for raw in ("", "[9]", "没有类型标识的一段文字", "[2] [J]", "garbage..."):
            with self.subTest(raw=raw):
                entry = parse_reference(0, raw)
                self.assertEqual(entry.raw, raw)
                self.assertIsInstance(entry.title, str)
                if raw.strip():
                    self.assertTrue(entry.title.strip(), f"title 应尽量非空：{raw!r}")

    def test_automatic_fallback_without_roles(self):
        """角色识别失败（空 Roles）时，parse_references 仍按 [n] 兜底解析。"""
        model, _ = load_fixture(self)
        entries = parse_references(model, Roles())
        self.assertEqual(len(entries), 3)
        self.assertEqual([entry.number for entry in entries], [1, 2, 3])


# ---------------------------------------------------------------------------
# 2. 引用一致性
# ---------------------------------------------------------------------------
class CitationAnalysisTest(unittest.TestCase):
    def _findings(self, *, body=None, refs=None):
        model, roles = load_fixture(self, body=body, refs=refs)
        entries = parse_references(model, roles)
        return model, entries, analyze_citations(model, roles, entries)

    def test_compliant_citations_have_no_findings(self):
        _, _, findings = self._findings()
        self.assertEqual(findings, [], [finding.message for finding in findings])

    def test_uncited_reference_reported(self):
        body = [[("只引用文献", False), ("[1]", True), ("和", False), ("[2]", True), ("。", False)]]
        _, _, findings = self._findings(body=body)
        uncited = [finding for finding in findings if "未被正文引用" in finding.message]
        self.assertEqual(len(uncited), 1, [finding.message for finding in findings])
        self.assertIn(3, uncited[0].actual)
        self.assertTrue(uncited[0].locator.startswith("refs/"))

    def test_out_of_range_citation_reported(self):
        body = [[("文献", False), ("[1]", True), ("与", False), ("[4]", True), ("。", False)]]
        _, _, findings = self._findings(body=body)
        errors = [finding for finding in findings if finding.level == "error"]
        self.assertTrue(errors)
        self.assertTrue(any("越界" in finding.message and "[4]" in finding.message for finding in errors))

    def test_zero_citation_reported(self):
        body = [[("文献", False), ("[0]", True), ("。", False)]]
        _, _, findings = self._findings(body=body)
        errors = [finding for finding in findings if finding.level == "error"]
        self.assertTrue(any("[0]" in finding.message and "[1] 开始" in finding.message for finding in errors))

    def test_first_citation_order_reported(self):
        body = [
            [("先引文献", False), ("[2]", True), ("。", False)],
            [("再引文献", False), ("[1]", True), ("与", False), ("[3]", True), ("。", False)],
        ]
        _, _, findings = self._findings(body=body)
        order = [finding for finding in findings if "首次引用顺序" in finding.message]
        self.assertEqual(len(order), 1, [finding.message for finding in findings])
        self.assertEqual(order[0].actual, [2, 1, 3])
        self.assertEqual(order[0].expected, [1, 2, 3])

    def test_missing_entry_number_reported(self):
        refs = [REFS[0], REFS[2]]  # [1] 与 [3]，缺 [2]
        body = [[("文献", False), ("[1]", True), ("与", False), ("[3]", True), ("。", False)]]
        _, _, findings = self._findings(body=body, refs=refs)
        gap = [finding for finding in findings if "编号不连续" in finding.message]
        self.assertEqual(len(gap), 1, [finding.message for finding in findings])
        self.assertEqual(gap[0].actual, [1, 3])
        self.assertEqual(gap[0].expected, [1, 2])

    def test_non_superscript_citation_reported(self):
        body = [[("相关研究见文献", False), ("[1]", False), ("、", False), ("[2]", False), ("。", False)]]
        model, _, findings = self._findings(body=body)
        superscripts = [finding for finding in findings if "上标" in finding.message]
        self.assertEqual(len(superscripts), 1, [finding.message for finding in findings])
        self.assertEqual(superscripts[0].level, "warning")
        self.assertEqual(sorted(superscripts[0].actual), [1, 2])
        body_index = next(
            para.index for para in model.paragraphs if "相关研究见文献" in para.text
        )
        self.assertEqual(superscripts[0].locator, f"P{body_index}")

    def test_citation_range_expanded(self):
        body = [[("文献", False), ("[1-2]", True), ("。", False)]]
        _, entries, findings = self._findings(body=body)
        self.assertEqual(len(entries), 3)
        self.assertFalse(
            [finding for finding in findings if finding.level == "error"],
            [finding.message for finding in findings],
        )

    def test_finding_fields_and_levels(self):
        body = [[("文献", False), ("[9]", False), ("。", False)]]
        _, _, findings = self._findings(body=body)
        self.assertTrue(findings)
        for finding in findings:
            self.assertIn(finding.level, ("error", "warning", "info"))
            self.assertTrue(finding.locator)


# ---------------------------------------------------------------------------
# 3. 真实性核验（假 fetcher，不联网）
# ---------------------------------------------------------------------------
EMPTY_CROSSREF = json.dumps({"status": "ok", "message": {"items": []}})
EMPTY_OPENALEX = json.dumps({"results": []})
EMPTY_ARXIV = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<feed xmlns="http://www.w3.org/2005/Atom"></feed>'
)
EMPTY = {"crossref": EMPTY_CROSSREF, "openalex": EMPTY_OPENALEX, "arxiv": EMPTY_ARXIV}


def crossref_payload(items) -> str:
    return json.dumps({"status": "ok", "message": {"items": items}}, ensure_ascii=False)


def crossref_item(title, *, year=None, venue="", doi="", authors=()) -> dict:
    item = {"title": [title]}
    if doi:
        item["DOI"] = doi
    if venue:
        item["container-title"] = [venue]
    if year is not None:
        item["issued"] = {"date-parts": [[year]]}
    if authors:
        item["author"] = [
            {"family": name.split()[0], "given": " ".join(name.split()[1:])}
            for name in authors
        ]
    return item


def openalex_payload(results) -> str:
    return json.dumps({"results": results}, ensure_ascii=False)


def make_fetcher(*, crossref=None, openalex=None, arxiv=None, failing=()):
    """按 URL 分流返回假响应；failing 中的数据源抛 URLError 模拟网络失败。"""
    responses = {"crossref": crossref, "openalex": openalex, "arxiv": arxiv}

    def fetcher(url: str, timeout: float) -> str:
        for source in ("crossref", "openalex", "arxiv"):
            if source in url:
                if source in failing:
                    raise urllib.error.URLError(f"fake {source} is down")
                payload = responses[source]
                return payload if payload is not None else EMPTY[source]
        raise AssertionError(f"未预期的 URL：{url}")

    return fetcher


class VerifyReferenceTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.entry = parse_reference(0, OPTUNA_REF)

    def test_reference_parsed_for_verification(self):
        self.assertEqual(self.entry.title, "Optuna: A Next-generation Hyperparameter Optimization Framework")
        self.assertEqual(self.entry.year, 2019)
        self.assertEqual(self.entry.doc_type, "C")

    def test_verified_on_crossref_hit(self):
        item = crossref_item(
            self.entry.title,
            year=2019,
            venue="Proceedings of the 25th ACM SIGKDD International Conference on Knowledge Discovery & Data Mining",
            doi="10.1145/3292500.3330701",
            authors=("Akiba Takuya", "Sano Shotaro"),
        )
        result = verify_reference(
            self.entry, fetcher=make_fetcher(crossref=crossref_payload([item]))
        )
        self.assertEqual(result.status, "verified")
        self.assertEqual(result.source, "crossref")
        self.assertGreaterEqual(result.score, VERIFIED_THRESHOLD)
        self.assertLessEqual(result.score, 1.0)
        self.assertEqual(result.doi, "10.1145/3292500.3330701")
        self.assertEqual(result.url, "https://doi.org/10.1145/3292500.3330701")
        self.assertTrue(result.matched_title)
        self.assertEqual(result.matched_year, 2019)
        self.assertIn(result.url, result.evidence)
        self.assertTrue(any("api.crossref.org" in url for url in result.evidence))

    def test_likely_on_partial_title(self):
        item = crossref_item(
            "Optuna: Hyperparameter Optimization Framework",
            doi="10.9999/likely.001",
        )
        result = verify_reference(
            self.entry, fetcher=make_fetcher(crossref=crossref_payload([item]))
        )
        self.assertEqual(result.status, "likely")
        self.assertGreaterEqual(result.score, LIKELY_THRESHOLD)
        self.assertLess(result.score, VERIFIED_THRESHOLD)
        self.assertIn("人工核实", result.note)

    def test_falls_back_to_openalex_when_crossref_empty(self):
        payload = openalex_payload(
            [
                {
                    "display_name": self.entry.title,
                    "publication_year": 2019,
                    "doi": "https://doi.org/10.1145/3292500.3330701",
                    "primary_location": {"source": {"display_name": "KDD"}},
                    "authorships": [{"author": {"display_name": "Takuya Akiba"}}],
                }
            ]
        )
        result = verify_reference(self.entry, fetcher=make_fetcher(openalex=payload))
        self.assertEqual(result.status, "verified")
        self.assertEqual(result.source, "openalex")
        self.assertEqual(result.doi, "10.1145/3292500.3330701")
        self.assertTrue(any("api.openalex.org" in url for url in result.evidence))

    def test_partial_network_error_still_verifies(self):
        # Crossref 只给出相近记录（不会提前结束），随后 OpenAlex 模拟不可用；
        # 结论应为 likely，并在 note 里如实说明「部分数据源请求失败」。
        item = crossref_item(
            "Optuna: Hyperparameter Optimization Framework",
            doi="10.9999/likely.002",
        )
        result = verify_reference(
            self.entry,
            fetcher=make_fetcher(
                crossref=crossref_payload([item]),
                failing=("openalex",),
            ),
        )
        self.assertEqual(result.status, "likely")
        self.assertIn("部分数据源请求失败", result.note)

    def test_partial_failure_without_match_is_error(self):
        """有数据源失败时不能声称「都没检索到」：无命中应报 error 而非 not_found。"""
        result = verify_reference(self.entry, fetcher=make_fetcher(failing=("openalex",)))
        self.assertEqual(result.status, "error")
        self.assertIn("未能完成完整核验", result.note)
        self.assertIn("OpenAlex", result.note)
        self.assertIn("不代表文献不存在", result.note)
        self.assertEqual(len(result.evidence), 3)

    def test_not_found_is_honest(self):
        result = verify_reference(self.entry, fetcher=make_fetcher())
        self.assertEqual(result.status, "not_found")
        self.assertEqual(result.source, "none")
        # 诚实性红线：必须说明覆盖有限、需人工核实，不得断言文献是假的
        self.assertIn("未在 Crossref/OpenAlex/arXiv 检索到", result.note)
        self.assertIn("中文期刊可能未被收录，需人工核实", result.note)
        self.assertIn("不代表文献造假", result.note)
        self.assertNotIn("该文献是假的", result.note)
        # 必须附证据 URL：三个已查询数据源
        self.assertTrue(any("api.crossref.org" in url for url in result.evidence))
        self.assertTrue(any("api.openalex.org" in url for url in result.evidence))
        self.assertTrue(any("arxiv.org" in url for url in result.evidence))

    def test_network_error_is_error_not_not_found(self):
        result = verify_reference(
            self.entry, fetcher=make_fetcher(failing=("crossref", "openalex", "arxiv"))
        )
        self.assertEqual(result.status, "error")
        self.assertIn("联网核验失败", result.note)
        self.assertIn("需人工核实", result.note)
        self.assertIn("URLError", result.note)
        # 网络失败也必须附上尝试过的证据 URL
        self.assertEqual(len(result.evidence), 3)
        self.assertTrue(any("api.crossref.org" in url for url in result.evidence))

    def test_empty_entry_returns_error_without_network(self):
        calls = []

        def fetcher(url, timeout):
            calls.append(url)
            raise AssertionError("空条目不应发起网络请求")

        result = verify_reference(parse_reference(0, ""), fetcher=fetcher)
        self.assertEqual(result.status, "error")
        self.assertIn("题名", result.note)
        self.assertEqual(calls, [])

    def test_verify_references_limit_and_sequential(self):
        entries = [parse_reference(i, raw) for i, raw in enumerate(REFS)]
        calls = []

        def fetcher(url, timeout):
            calls.append(url)
            return EMPTY[next(source for source in EMPTY if source in url)]

        results = verify_references(entries, limit=2, pause=0, fetcher=fetcher)
        self.assertEqual(len(results), 2)
        self.assertEqual([result.entry.number for result in results], [1, 2])
        # 每条 not_found 会依次查询 3 个数据源 → 2 条共 6 次请求
        self.assertEqual(len(calls), 6)
        self.assertTrue(all(result.status == "not_found" for result in results))

    def test_short_generic_title_not_overclaimed(self):
        """「物理学」这类短题名即使精确匹配，也只能给 likely（防同名误报）。"""
        entry = parse_reference(0, "[2] 竺可桢．物理学[M]．北京：科学出版社，2015．")
        item = crossref_item("物理学", year=1962, doi="10.2142/biophys.2.268")
        result = verify_reference(
            entry, fetcher=make_fetcher(crossref=crossref_payload([item]))
        )
        self.assertEqual(result.status, "likely")
        self.assertLess(result.score, VERIFIED_THRESHOLD)
        self.assertIn("区分度低", result.note)
        self.assertIn("1962", result.note)  # 年份不一致提醒
        self.assertIn("不一致", result.note)

    def test_year_mismatch_flagged_in_note(self):
        item = crossref_item(self.entry.title, year=1999, doi="10.1145/3292500.3330701")
        result = verify_reference(
            self.entry, fetcher=make_fetcher(crossref=crossref_payload([item]))
        )
        self.assertEqual(result.status, "verified")
        self.assertIn("年份", result.note)
        self.assertIn("不一致", result.note)

    def test_title_similarity_thresholds(self):
        self.assertEqual(title_similarity("Optuna", "optuna"), 1.0)
        self.assertGreater(
            title_similarity("交流电机的多回路理论", "交流电机的多回路分析"), 0.75
        )
        self.assertLess(
            title_similarity(
                "Optuna: A Next-generation Hyperparameter Optimization Framework",
                "Deep Residual Learning for Image Recognition",
            ),
            0.4,
        )

    def test_default_fetcher_is_urllib_based(self):
        # 只做类型/签名约定检查，不真正发请求
        self.assertTrue(callable(default_fetcher))
        self.assertEqual(default_fetcher.__module__, "tas.refs.verify")


# ---------------------------------------------------------------------------
# 4. 真实网络（默认 skip）
# ---------------------------------------------------------------------------
@unittest.skipUnless(
    os.environ.get("TAS_LIVE_NETWORK") == "1",
    "需要 TAS_LIVE_NETWORK=1 才访问真实 Crossref/OpenAlex/arXiv（默认离线）",
)
class LiveNetworkTest(unittest.TestCase):
    def test_known_paper_verified_via_real_api(self):
        entry = parse_reference(0, OPTUNA_REF)
        result = verify_reference(entry, timeout=20.0)
        detail = f"status={result.status} score={result.score} note={result.note} evidence={result.evidence}"
        self.assertIn(result.status, ("verified", "likely"), detail)
        self.assertTrue(result.evidence, detail)
        self.assertTrue(result.doi or result.url, detail)
        self.assertGreaterEqual(result.score, LIKELY_THRESHOLD, detail)


if __name__ == "__main__":
    unittest.main(verbosity=2)
