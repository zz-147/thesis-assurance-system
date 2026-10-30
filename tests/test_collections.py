"""文档接龙测试：存储/检查/审核服务 + HTTP API + 手机 H5 页。"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

from tas import collections  # noqa: E402
from tas.web.app import create_server  # noqa: E402
from test_check import build_good_docx  # noqa: E402
from test_web import _multipart, _mutate_font  # noqa: E402


class _TempDataMixin:
    def _setup_data_dir(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self._data_dir = Path(self._tmp.name) / "data"
        self._env = mock.patch.dict(os.environ, {"TAS_DATA_DIR": str(self._data_dir)})
        self._env.start()
        self.addCleanup(self._env.stop)


class CollectionsServiceTest(_TempDataMixin, unittest.TestCase):
    def setUp(self) -> None:
        self._setup_data_dir()

    def _docx(self, mutate: bool = False) -> Path:
        path = build_good_docx(self._data_dir.parent / f"doc_{os.urandom(3).hex()}.docx")
        if mutate:
            _mutate_font(path)
        return path

    def test_create_get_list(self):
        item = collections.create_collection(
            title="第 3 次课程设计报告", owner="张三", ruleset="school_default", note="周五前提交"
        )
        self.assertEqual(len(item["code"]), 6)
        self.assertEqual(item["submission_count"], 0)
        detail = collections.get_collection(item["code"])
        self.assertEqual(detail["collection"]["title"], "第 3 次课程设计报告")
        self.assertEqual(detail["submissions"], [])
        listed = collections.list_collections()
        self.assertEqual(len(listed), 1)
        self.assertEqual(listed[0]["code"], item["code"])

    def test_create_requires_title(self):
        with self.assertRaises(ValueError):
            collections.create_collection(title="  ")

    def test_submit_compliant_document(self):
        item = collections.create_collection(title="合规作业")
        path = self._docx()
        submission = collections.add_submission(item["code"], "李四", path.name, path.read_bytes())
        self.assertEqual(submission["status"], "通过")
        self.assertEqual(submission["counts"]["error"], 0)
        detail = collections.get_collection(item["code"])
        self.assertEqual(len(detail["submissions"]), 1)
        self.assertEqual(detail["submissions"][0]["student"], "李四")

    def test_submit_bad_document_and_review(self):
        item = collections.create_collection(title="需要审核的作业")
        good = self._docx()
        bad = self._docx(mutate=True)
        collections.add_submission(item["code"], "甲", good.name, good.read_bytes())
        collections.add_submission(item["code"], "乙", bad.name, bad.read_bytes())
        result = collections.review_collection(item["code"])
        summary = result["summary"]
        self.assertEqual(summary["total"], 2)
        self.assertEqual(summary["checked"], 2)
        self.assertEqual(summary["failed"], 0)
        self.assertEqual(summary["pass_rate"], 0.5)
        statuses = {row["student"]: row["status"] for row in result["submissions"]}
        self.assertEqual(statuses, {"甲": "通过", "乙": "未通过"})
        self.assertTrue(summary["top_rules"])

        report = collections.get_submission_report(item["code"], result["submissions"][1]["id"])
        self.assertIn("issues", report)
        self.assertTrue(report["issues"])

    def test_reject_non_docx(self):
        item = collections.create_collection(title="拒绝无关文件")
        with self.assertRaises(ValueError):
            collections.add_submission(item["code"], "甲", "notes.txt", b"hello")

    def test_unknown_code(self):
        with self.assertRaises(ValueError):
            collections.get_collection("AAAAAA")
        with self.assertRaises(ValueError):
            collections.get_collection("BAD CODE")


class CollectionsApiTest(_TempDataMixin, unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp_cls = tempfile.TemporaryDirectory()
        cls._env_cls = mock.patch.dict(
            os.environ, {"TAS_DATA_DIR": str(Path(cls._tmp_cls.name) / "data")}
        )
        cls._env_cls.start()
        cls.server = create_server("127.0.0.1", 0)
        cls.port = cls.server.server_address[1]
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = f"http://127.0.0.1:{cls.port}"

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=5)
        cls._env_cls.stop()
        cls._tmp_cls.cleanup()

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)

    def _json_request(self, method: str, path: str, payload: dict | None = None):
        data = json.dumps(payload or {}, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            self.base + path, data=data, method=method,
            headers={"Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                return response.status, json.loads(response.read())
        except urllib.error.HTTPError as error:
            return error.code, json.loads(error.read())

    def _post_multipart(self, path: str, fields: dict, files: list):
        body, ctype = _multipart(fields, files)
        request = urllib.request.Request(
            self.base + path, data=body, method="POST", headers={"Content-Type": ctype}
        )
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                return response.status, json.loads(response.read())
        except urllib.error.HTTPError as error:
            return error.code, json.loads(error.read())

    def _get(self, path: str):
        try:
            with urllib.request.urlopen(self.base + path, timeout=120) as response:
                return response.status, response.read()
        except urllib.error.HTTPError as error:
            return error.code, error.read()

    def test_full_flow(self):
        status, data = self._json_request(
            "POST", "/api/collections",
            {"title": "API 接龙", "owner": "王五", "ruleset": "school_default"},
        )
        self.assertEqual(status, 200)
        code = data["collection"]["code"]

        status, listing = self._json_request("GET", "/api/collections")
        self.assertEqual(status, 200)
        self.assertTrue(any(item["code"] == code for item in listing["collections"]))

        path = build_good_docx(Path(self._tmp.name) / "student.docx")
        _mutate_font(path)
        status, submitted = self._post_multipart(
            f"/api/collections/{code}/submit",
            {"student": "赵六"},
            [("file", "student.docx", path.read_bytes())],
        )
        self.assertEqual(status, 200, submitted)
        self.assertEqual(submitted["submission"]["status"], "未通过")
        self.assertGreater(submitted["submission"]["counts"]["error"], 0)

        status, reviewed = self._json_request("POST", f"/api/collections/{code}/review")
        self.assertEqual(status, 200, reviewed)
        self.assertEqual(reviewed["summary"]["total"], 1)
        self.assertEqual(reviewed["summary"]["pass_rate"], 0.0)

        submission_id = reviewed["submissions"][0]["id"]
        status, body = self._get(f"/api/collections/{code}/submissions/{submission_id}/report")
        self.assertEqual(status, 200)
        report = json.loads(body)
        self.assertTrue(report["ok"])
        self.assertTrue(report["report"]["issues"])

    def test_error_returns_json(self):
        status, data = self._json_request("GET", "/api/collections/ZZZZZZ")
        self.assertEqual(status, 400)
        self.assertFalse(data["ok"])
        self.assertIn("接龙不存在", data["error"])

    def test_mobile_page(self):
        status, body = self._get("/m")
        self.assertEqual(status, 200)
        page = body.decode("utf-8")
        self.assertIn("文档接龙", page)
        self.assertIn("一键审核", page)
        self.assertIn("/api/collections", page)


if __name__ == "__main__":
    unittest.main(verbosity=2)
