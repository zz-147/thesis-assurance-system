"""批量文献真实性核验（batch_refs.py）回归测试 —— 完全离线，用 fake fetcher。

覆盖：
1) 逐份解析 + 假 Crossref 返回一条高相似（verified）、其余空结果（not_found）；
2) 汇总计数：total/checked/failed、total_entries、四状态计数、per_doc；
3) 损坏文件 / 不存在文件 / 无参考文献文档：单份失败不中断，无文献不算失败；
4) limit_per_doc 只核验前 N 条（entries 仍是解析到的总条数）；
5) Markdown 报告：汇总表 + 逐条明细 + 「人工核实」诚实性提示 + 失败清单；
6) write_refs_report：JSON 可 json.loads（summary + docs 完整明细）、MD 可读。

夹具复用 tests/test_refs.py 的 build_docx / crossref_payload / crossref_item，
参考文献均为学院规范示例或公开文献，不涉及真实学生文档；测试不联网。
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
import urllib.error
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from test_refs import (  # noqa: E402
    EMPTY_ARXIV,
    EMPTY_CROSSREF,
    EMPTY_OPENALEX,
    build_docx,
    crossref_item,
    crossref_payload,
)

from tas.batch_refs import (  # noqa: E402
    STATUS_KEYS,
    BatchRefDocResult,
    BatchRefSummary,
    render_refs_markdown,
    verify_batch,
    write_refs_report,
)

# 公开文献（CVPR 2016，Crossref 有 DOI）——用于 verified
RESNET_REF = (
    "[1] He K, Zhang X, Ren S, et al. Deep Residual Learning for Image Recognition[C]. "
    "Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition, 2016: 770-778."
)
# 学院规范表 3 示例（中文期刊）——检索不到时如实 not_found
CN_REF = "[1] 高景德，王祥珩．交流电机的多回路理论[J]．清华大学学报，2017，27（1）：1－8．"
CN_REF_2 = "[2] 竺可桢．物理学[M]．北京：科学出版社，2015．"

RESNET_DOI = "10.1109/cvpr.2016.90"


# ---------------------------------------------------------------------------
# 夹具 / fake fetcher
# ---------------------------------------------------------------------------
def make_fake_fetcher(calls: list[str] | None = None):
    """离线 fetcher：ResNet 题名命中 Crossref，其余一律空结果。"""
    recorded = calls if calls is not None else []

    def fetcher(url: str, timeout: float) -> str:
        recorded.append(url)
        if "api.crossref.org" in url:
            query = urllib.parse.parse_qs(
                urllib.parse.urlparse(url).query
            ).get("query.bibliographic", [""])[0]
            if "Deep Residual Learning" in query:
                item = crossref_item(
                    "Deep Residual Learning for Image Recognition",
                    year=2016,
                    venue="2016 IEEE Conference on Computer Vision and Pattern Recognition (CVPR)",
                    doi=RESNET_DOI,
                    authors=("Kaiming He",),
                )
                return crossref_payload([item])
            return EMPTY_CROSSREF
        if "api.openalex.org" in url:
            return EMPTY_OPENALEX
        if "export.arxiv.org" in url:
            return EMPTY_ARXIV
        raise AssertionError(f"未预期的 URL：{url}")

    return fetcher


def _build(path: Path, refs) -> Path:
    return build_docx(path, refs=list(refs), body=[])


class BatchRefsFixtures(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)

    def doc(self, name: str, refs) -> Path:
        return _build(self.tmp / name, refs)

    def corrupt(self, name: str = "broken.docx") -> Path:
        path = self.tmp / name
        path.write_bytes(b"this is not a real docx package")
        return path

    def mixed_paths(self) -> list[Path]:
        """verified + not_found + 加载失败 三份混合文档。"""
        return [
            self.doc("a_verified.docx", [RESNET_REF, CN_REF]),
            self.doc("b_missing.docx", [CN_REF]),
            self.corrupt(),
        ]


# ---------------------------------------------------------------------------
# 批量核验 + 汇总
# ---------------------------------------------------------------------------
class VerifyBatchTest(BatchRefsFixtures):
    def test_summary_counts_verified_not_found_and_failed(self):
        docs, summary = verify_batch(
            self.mixed_paths(), fetcher=make_fake_fetcher(), pause=0, limit_per_doc=5
        )
        self.assertEqual(len(docs), 3)

        # 第一份：ResNet → verified；中文期刊 → not_found
        self.assertTrue(docs[0].ok)
        self.assertEqual(docs[0].entries, 2)
        self.assertEqual(len(docs[0].results), 2)
        self.assertEqual([result.status for result in docs[0].results], ["verified", "not_found"])
        self.assertEqual(docs[0].results[0].doi, RESNET_DOI)
        self.assertEqual(docs[0].results[0].source, "crossref")

        # 第二份：只有中文文献（fake 全空）→ not_found
        self.assertTrue(docs[1].ok)
        self.assertEqual(docs[1].entries, 1)
        self.assertEqual([result.status for result in docs[1].results], ["not_found"])

        # 第三份：损坏文件 → 只记 error，不中断
        self.assertFalse(docs[2].ok)
        self.assertIsNotNone(docs[2].error)
        self.assertEqual(docs[2].entries, 0)
        self.assertEqual(docs[2].results, [])

        self.assertEqual(summary.total_docs, 3)
        self.assertEqual(summary.checked_docs, 2)
        self.assertEqual(summary.failed_docs, 1)
        self.assertEqual(summary.total_entries, 3)
        self.assertEqual(summary.verified, 1)
        self.assertEqual(summary.likely, 0)
        self.assertEqual(summary.not_found, 2)
        self.assertEqual(summary.error, 0)
        self.assertEqual(len(summary.per_doc), 3)
        self.assertTrue(summary.per_doc[2]["failed"])
        self.assertIn("PackageNotFoundError", summary.per_doc[2]["load_error"] or "")
        self.assertEqual(summary.per_doc[0]["name"], "a_verified.docx")
        self.assertEqual(summary.per_doc[0]["checked"], 2)
        self.assertEqual(summary.per_doc[0]["error"], 0)

        # to_dict 必须可直接 JSON 序列化
        payload = json.dumps(summary.to_dict(), ensure_ascii=False)
        self.assertIn("a_verified.docx", payload)
        for key in (
            "total_docs", "checked_docs", "failed_docs", "total_entries",
            "verified", "likely", "not_found", "error", "per_doc",
        ):
            self.assertIn(key, summary.to_dict())

    def test_doc_result_to_dict_contains_result_details(self):
        docs, _ = verify_batch(
            [self.doc("one.docx", [RESNET_REF])], fetcher=make_fake_fetcher(), pause=0
        )
        data = docs[0].to_dict()
        self.assertEqual(data["doc"], str(docs[0].doc))
        self.assertEqual(data["entries"], 1)
        self.assertEqual(data["checked"], 1)
        self.assertIsNone(data["error"])
        self.assertIsNone(data["load_error"])
        self.assertEqual(data["status_error"], 0)
        self.assertEqual(data["counts"]["verified"], 1)
        result = data["results"][0]
        self.assertEqual(result["status"], "verified")
        self.assertEqual(result["doi"], RESNET_DOI)
        self.assertEqual(result["entry"]["title"], "Deep Residual Learning for Image Recognition")
        self.assertTrue(result["evidence"])
        json.dumps(data, ensure_ascii=False)  # 不抛异常即可

    def test_likely_status_counted_in_summary(self):
        def fetcher(url: str, timeout: float) -> str:
            if "api.crossref.org" in url:
                return crossref_payload(
                    [crossref_item("Deep Residual Learning", doi="10.9999/partial.001")]
                )
            if "api.openalex.org" in url:
                return EMPTY_OPENALEX
            return EMPTY_ARXIV

        path = self.doc("partial.docx", [RESNET_REF])
        docs, summary = verify_batch([path], fetcher=fetcher, pause=0)
        self.assertTrue(docs[0].ok)
        self.assertEqual(docs[0].results[0].status, "likely")
        self.assertEqual(summary.checked_docs, 1)
        self.assertEqual(summary.likely, 1)
        self.assertEqual(summary.verified, 0)

    def test_network_error_status_counted_but_doc_not_failed(self):
        def failing_fetcher(url: str, timeout: float) -> str:
            raise urllib.error.URLError("fake network down")

        path = self.doc("network_error.docx", [CN_REF])
        docs, summary = verify_batch([path], fetcher=failing_fetcher, pause=0)
        # 文档本身加载/解析成功，只是文献核验遇到网络错误
        self.assertTrue(docs[0].ok)
        self.assertEqual(docs[0].results[0].status, "error")
        self.assertEqual(summary.checked_docs, 1)
        self.assertEqual(summary.failed_docs, 0)
        self.assertEqual(summary.error, 1)
        self.assertEqual(summary.not_found, 0)

    def test_missing_file_recorded_without_interrupting(self):
        good = self.doc("good.docx", [CN_REF])
        missing = self.tmp / "not_exists.docx"
        docs, summary = verify_batch(
            [missing, good], fetcher=make_fake_fetcher(), pause=0
        )
        self.assertEqual(len(docs), 2)
        self.assertFalse(docs[0].ok)
        self.assertIn("FileNotFoundError", docs[0].error or "")
        self.assertTrue(docs[1].ok)
        self.assertEqual(summary.checked_docs, 1)
        self.assertEqual(summary.failed_docs, 1)

    def test_document_without_references_is_checked_not_failed(self):
        path = self.doc("empty_refs.docx", [])
        docs, summary = verify_batch([path], fetcher=make_fake_fetcher(), pause=0)
        self.assertTrue(docs[0].ok)
        self.assertEqual(docs[0].entries, 0)
        self.assertEqual(docs[0].results, [])
        self.assertEqual(summary.checked_docs, 1)
        self.assertEqual(summary.failed_docs, 0)
        self.assertEqual(summary.total_entries, 0)
        self.assertEqual(summary.verified, 0)
        self.assertEqual(summary.not_found, 0)

    def test_limit_per_doc_only_verifies_prefix(self):
        path = self.doc("three.docx", [RESNET_REF, CN_REF, CN_REF_2])
        docs, summary = verify_batch(
            [path], fetcher=make_fake_fetcher(), pause=0, limit_per_doc=1
        )
        self.assertEqual(docs[0].entries, 3)  # 解析到的总条数
        self.assertEqual(len(docs[0].results), 1)  # 只核验 1 条
        self.assertEqual(summary.total_entries, 3)
        self.assertEqual(summary.verified, 1)
        self.assertEqual(summary.not_found, 0)
        md = render_refs_markdown(summary, docs)
        self.assertIn("核验上限", md)

    def test_limit_per_doc_none_verifies_all(self):
        path = self.doc("three_all.docx", [RESNET_REF, CN_REF, CN_REF_2])
        docs, _ = verify_batch(
            [path], fetcher=make_fake_fetcher(), pause=0, limit_per_doc=None
        )
        self.assertEqual(len(docs[0].results), 3)

    def test_limit_per_doc_zero_marks_not_checked(self):
        path = self.doc("zero_limit.docx", [RESNET_REF, CN_REF])
        docs, summary = verify_batch(
            [path], fetcher=make_fake_fetcher(), pause=0, limit_per_doc=0
        )
        self.assertEqual(docs[0].entries, 2)
        self.assertEqual(docs[0].results, [])
        self.assertEqual(summary.total_entries, 2)
        self.assertEqual(summary.verified, 0)
        md = render_refs_markdown(summary, docs)
        self.assertIn("本次未核验", md)

    def test_empty_paths_returns_empty_summary(self):
        docs, summary = verify_batch([], fetcher=make_fake_fetcher(), pause=0)
        self.assertEqual(docs, [])
        self.assertEqual(summary.total_docs, 0)
        self.assertEqual(summary.checked_docs, 0)
        self.assertEqual(summary.failed_docs, 0)
        self.assertEqual(summary.total_entries, 0)
        for key in STATUS_KEYS:
            self.assertEqual(getattr(summary, key), 0)
        md = render_refs_markdown(summary, docs)
        self.assertIn("文档总数：**0**", md)

    def test_string_paths_accepted(self):
        path = self.doc("string_path.docx", [CN_REF])
        docs, _ = verify_batch([str(path)], fetcher=make_fake_fetcher(), pause=0)
        self.assertIsInstance(docs[0].doc, Path)
        self.assertTrue(docs[0].ok)


# ---------------------------------------------------------------------------
# Markdown / JSON 报告
# ---------------------------------------------------------------------------
class RefReportsTest(BatchRefsFixtures):
    def _mixed(self):
        paths = self.mixed_paths()
        return verify_batch(paths, fetcher=make_fake_fetcher(), pause=0)

    def test_markdown_contains_summary_details_and_honesty(self):
        docs, summary = self._mixed()
        md = render_refs_markdown(summary, docs)

        # 汇总表头 + 数据
        self.assertIn("| 文件 | 参考文献数 | verified | likely | not_found | error |", md)
        self.assertIn("a_verified.docx", md)
        # 逐条明细：题名 / 状态 / 来源 / 匹配 / DOI / note
        self.assertIn("## 逐条明细", md)
        self.assertIn("Deep Residual Learning for Image Recognition", md)
        self.assertIn("交流电机的多回路理论", md)
        self.assertIn("verified", md)
        self.assertIn("crossref", md)
        self.assertIn(RESNET_DOI, md)
        self.assertIn("已核验", md)
        # 失败清单
        self.assertIn("## 加载失败文档", md)
        self.assertIn("broken.docx", md)
        # 诚实性提示（任务要求的固定文案）
        self.assertIn(
            "not_found 仅表示未在当前索引检索到，中文期刊可能未被收录，需人工核实", md
        )
        self.assertIn("未检索到 ≠ 文献造假", md)

    def test_write_refs_report_json_and_markdown(self):
        docs, summary = self._mixed()
        out_json = self.tmp / "out" / "refs_report.json"
        out_md = self.tmp / "out" / "refs_report.md"
        returned_json, returned_md = write_refs_report(summary, docs, out_json, out_md)

        self.assertEqual(Path(returned_json), Path(out_json))
        self.assertEqual(Path(returned_md), Path(out_md))
        self.assertTrue(out_json.is_file())
        self.assertTrue(out_md.is_file())

        payload = json.loads(out_json.read_text(encoding="utf-8"))
        self.assertIn("summary", payload)
        self.assertIn("docs", payload)
        self.assertEqual(payload["summary"]["total_docs"], 3)
        self.assertEqual(payload["summary"]["failed_docs"], 1)
        self.assertEqual(payload["docs"][0]["results"][0]["status"], "verified")
        self.assertEqual(
            payload["docs"][0]["results"][1]["status"], "not_found"
        )
        self.assertEqual(
            payload["docs"][0]["results"][0]["entry"]["title"],
            "Deep Residual Learning for Image Recognition",
        )

        md = out_md.read_text(encoding="utf-8")
        self.assertIn("not_found 仅表示未在当前索引检索到", md)
        self.assertIn("中文期刊可能未被收录，需人工核实", md)

    def test_write_report_creates_parent_directory(self):
        docs, summary = self._mixed()
        out_json = self.tmp / "nested" / "deep" / "r.json"
        out_md = self.tmp / "nested" / "deep" / "r.md"
        write_refs_report(summary, docs, out_json, out_md)
        self.assertTrue(out_json.is_file())
        self.assertTrue(out_md.is_file())

    def test_report_when_all_docs_failed(self):
        docs, summary = verify_batch(
            [self.corrupt("bad1.docx")], fetcher=make_fake_fetcher(), pause=0
        )
        md = render_refs_markdown(summary, docs)
        self.assertIn("加载失败", md)
        self.assertIn("## 加载失败文档", md)
        # 即使全部失败，也必须保留诚实性提示与汇总表
        self.assertIn("not_found 仅表示未在当前索引检索到", md)
        self.assertIn("| 文件 | 参考文献数 | verified | likely | not_found | error |", md)


if __name__ == "__main__":
    unittest.main(verbosity=2)
