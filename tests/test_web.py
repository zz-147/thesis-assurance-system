"""Web 双入口测试：本地起服务，模拟学生检查/修复、教师批量、下载。"""

from __future__ import annotations

import io
import json
import os
import re
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
import uuid
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from tas.web.app import create_server  # noqa: E402
from test_check import build_good_docx  # noqa: E402

_TOKEN_RE = re.compile(r'name="token" value="([0-9a-f]+)"')


def _mutate_font(path: Path) -> Path:
    document = docx.Document(str(path))
    for paragraph in document.paragraphs:
        if paragraph.text.startswith("文档格式统一是高校教学管理"):
            run = paragraph.runs[0]
            sz = run._r.rPr.find(qn("w:sz"))
            sz.set(qn("w:val"), "21")  # 10.5pt，正文应为 12pt
            break
    document.save(str(path))
    return path


def _multipart(fields: dict, files: list[tuple[str, str, bytes]]) -> tuple[bytes, str]:
    boundary = "----tas" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for key, value in fields.items():
        chunks.append(
            f'--{boundary}\r\nContent-Disposition: form-data; name="{key}"\r\n\r\n{value}\r\n'.encode(
                "utf-8"
            )
        )
    for field, filename, data in files:
        chunks.append(
            (
                f'--{boundary}\r\nContent-Disposition: form-data; name="{field}"; '
                f'filename="{filename}"\r\nContent-Type: application/octet-stream\r\n\r\n'
            ).encode("utf-8")
            + data
            + b"\r\n"
        )
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


class WebTestBase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
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

    # -- helpers -----------------------------------------------------------
    def get(self, path: str):
        with urllib.request.urlopen(self.base + path, timeout=60) as response:
            return response.status, response.read(), dict(response.headers)

    def post(self, path: str, fields: dict, files: list[tuple[str, str, bytes]]):
        body, ctype = _multipart(fields, files)
        request = urllib.request.Request(
            self.base + path, data=body, headers={"Content-Type": ctype}, method="POST"
        )
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                return response.status, response.read(), dict(response.headers)
        except urllib.error.HTTPError as error:  # noqa: F821
            return error.code, error.read(), dict(error.headers)


class HomeAndApiTest(WebTestBase):
    def test_home_page(self):
        status, body, _ = self.get("/")
        self.assertEqual(status, 200)
        text = body.decode("utf-8")
        self.assertIn("学生端", text)
        self.assertIn("教师端", text)
        self.assertIn("批量检查", text)

    def test_profiles_api(self):
        status, body, _ = self.get("/api/profiles")
        self.assertEqual(status, 200)
        payload = json.loads(body)
        self.assertIn("profiles", payload)
        self.assertTrue(any(item["id"] == "school_default" for item in payload["profiles"]))

    def test_health(self):
        status, body, _ = self.get("/health")
        self.assertEqual(status, 200)
        self.assertTrue(json.loads(body)["ok"])

    def test_diag(self):
        """真机网络排查接口：返回服务端看到的客户端 IP / 监听地址。"""
        status, body, _ = self.get("/api/diag")
        self.assertEqual(status, 200)
        payload = json.loads(body)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["service"], "tas-web")
        self.assertIn("client", payload)
        self.assertIn("server_host", payload)
        self.assertIn("server_port", payload)
        self.assertIn("base_url", payload)

    def test_home_has_drag_and_drop(self):
        status, body, _ = self.get("/")
        self.assertEqual(status, 200)
        text = body.decode("utf-8")
        self.assertIn('class="dropzone"', text)
        self.assertIn("拖到这里", text)
        self.assertIn("DataTransfer", text)          # 拖拽 JS 已注入
        self.assertIn('data-input="student-file"', text)
        self.assertIn('data-input="teacher-files"', text)
        # 规范文档拖拽区（.docx/.txt/.md）
        self.assertIn('data-input="spec-file"', text)
        self.assertIn('accept=".docx,.txt,.md"', text)
        self.assertIn("规范管理", text)
        # 学生端单文件；教师端多文件
        self.assertIn('id="student-file" name="file" accept=".docx" required hidden', text)
        self.assertIn(
            'id="teacher-files" name="files" accept=".docx" multiple required hidden', text
        )
        # 页面预览开关（默认勾选）
        self.assertIn('name="preview" value="1" checked', text)
        # 教师端：选择整个文件夹（webkitdirectory）
        self.assertIn("选择整个文件夹", text)
        self.assertIn("webkitdirectory", text)
        self.assertIn('id="teacher-dir"', text)
        # 教师端：名单核对 + 批量安全修复
        self.assertIn('name="roster"', text)
        self.assertIn('name="fix" value="1"', text)
        self.assertIn("安全修复版", text)
        # 可选增强：首页缩略图 + 文献真实性核验
        self.assertIn('name="refs" value="1"', text)
        self.assertIn("生成首页缩略图", text)
        self.assertIn("文献真实性核验", text)


class StudentFlowTest(WebTestBase):
    def test_check_then_fix_then_download(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = build_good_docx(Path(tmp) / "student.docx")
            _mutate_font(path)
            data = path.read_bytes()
            status, body, _ = self.post(
                "/student/check",
                {"ruleset": "school_default"},
                [("file", "student.docx", data)],
            )
            self.assertEqual(status, 200, body[:300])
            page = body.decode("utf-8")
            self.assertIn("检查结果", page)
            match = _TOKEN_RE.search(page)
            self.assertIsNotNone(match, "检查结果页应包含修复表单 token")
            token = match.group(1)

            status, body, _ = self.post(
                "/student/fix", {"token": token, "ruleset": "school_default"}, []
            )
            self.assertEqual(status, 200, body[:300])
            fix_page = body.decode("utf-8")
            self.assertIn("修复完成", fix_page)
            self.assertIn("内容 hash 一致", fix_page)

            status, body, headers = self.get(f"/download/{token}/fixed.docx")
            self.assertEqual(status, 200)
            self.assertEqual(body[:2], b"PK")  # docx/zip magic
            self.assertIn("wordprocessingml", headers.get("Content-Type", ""))

            status, body, _ = self.get(f"/download/{token}/fix.json")
            self.assertEqual(status, 200)
            payload = json.loads(body)
            self.assertIn("actions", payload)
            self.assertTrue(payload["content_hash_before"] == payload["content_hash_after"])

    def test_rejects_non_docx(self):
        status, body, _ = self.post(
            "/student/check", {"ruleset": "school_default"}, [("file", "old.doc", b"not a docx")]
        )
        self.assertEqual(status, 400)
        self.assertIn("只支持 .docx", body.decode("utf-8"))


class TeacherBatchTest(WebTestBase):
    def test_batch_summary_and_xlsx(self):
        try:
            import tas.batch  # noqa: F401
            import tas.export  # noqa: F401
        except ImportError:
            self.skipTest("batch/export 模块尚未交付")
        with tempfile.TemporaryDirectory() as tmp:
            good = build_good_docx(Path(tmp) / "good.docx")
            bad = build_good_docx(Path(tmp) / "bad.docx")
            _mutate_font(bad)
            status, body, _ = self.post(
                "/teacher/batch",
                {"ruleset": "school_default"},
                [
                    ("files", "good.docx", good.read_bytes()),
                    ("files", "bad.docx", bad.read_bytes()),
                ],
            )
            self.assertEqual(status, 200, body[:300])
            page = body.decode("utf-8")
            self.assertIn("班级批量检查结果", page)
            self.assertIn("通过率", page)
            excel_match = re.search(r"/download/([0-9a-f]+)/batch\.xlsx", page)
            self.assertIsNotNone(excel_match, "应提供 xlsx 下载链接")
            token = excel_match.group(1)

            status, body, _ = self.get(f"/download/{token}/batch.xlsx")
            self.assertEqual(status, 200)
            self.assertEqual(body[:2], b"PK")
            try:
                from openpyxl import load_workbook
            except ImportError:  # pragma: no cover
                return
            workbook = load_workbook(io.BytesIO(body))
            self.assertGreaterEqual(len(workbook.sheetnames), 2)

            status, body, _ = self.get(f"/download/{token}/batch.csv")
            self.assertEqual(status, 200)
            self.assertTrue(len(body) > 0)

            # 学生反馈包（每份文档一个 txt + 班级汇总）
            self.assertIn("/feedback.zip", page)
            status, body, _ = self.get(f"/download/{token}/feedback.zip")
            self.assertEqual(status, 200)
            self.assertEqual(body[:2], b"PK")
            with zipfile.ZipFile(io.BytesIO(body)) as archive:
                names = archive.namelist()
            self.assertIn("00_班级汇总.txt", names)
            self.assertTrue(any(name.startswith("反馈_") for name in names))

    def test_batch_with_roster_and_fix(self):
        """教师端：名单核对（缺交）+ 批量安全修复（生成 fixed.zip）。"""
        with tempfile.TemporaryDirectory() as tmp:
            good = build_good_docx(Path(tmp) / "23050301001_张三_论文.docx")
            bad = build_good_docx(Path(tmp) / "23050301002_李四_论文.docx")
            _mutate_font(bad)
            roster = "23050301001 张三\n23050301002 李四\n23050301003 王五"
            status, body, _ = self.post(
                "/teacher/batch",
                {"ruleset": "school_default", "roster": roster, "fix": "1"},
                [
                    ("files", "23050301001_张三_论文.docx", good.read_bytes()),
                    ("files", "23050301002_李四_论文.docx", bad.read_bytes()),
                ],
            )
            self.assertEqual(status, 200, body[:300])
            page = body.decode("utf-8")
            self.assertIn("名单核对", page)
            self.assertIn("23050301003", page)          # 王五缺交
            self.assertIn("批量安全修复", page)
            token_match = re.search(r"/download/([0-9a-f]+)/fixed\.zip", page)
            self.assertIsNotNone(token_match, "应提供 fixed.zip 下载链接")
            token = token_match.group(1)

            status, body, _ = self.get(f"/download/{token}/fixed.zip")
            self.assertEqual(status, 200)
            self.assertEqual(body[:2], b"PK")
            with zipfile.ZipFile(io.BytesIO(body)) as archive:
                names = archive.namelist()
            self.assertTrue(any(name.endswith("_fixed.docx") for name in names))
            self.assertIn("修复报告.txt", names)

            status, body, _ = self.get(f"/download/{token}/roster-report.txt")
            self.assertEqual(status, 200)
            self.assertIn("王五", body.decode("utf-8"))


class SpecFlowTest(WebTestBase):
    """规范文档上传 → LLM 抽取 → 生成 profile 的 Web 流程（LLM 部分不联网测试）。"""

    def test_spec_parse_without_key_is_friendly(self):
        import os
        from unittest import mock

        with mock.patch.dict(
            os.environ, {"TAS_LLM_API_KEY": "", "DEEPSEEK_API_KEY": ""}, clear=False
        ):
            status, body, _ = self.post(
                "/spec/parse",
                {"base_profile": "school_default"},
                [("file", "spec.txt", "一、正文：小四宋体。".encode("utf-8"))],
            )
        self.assertEqual(status, 200, body[:300])
        page = body.decode("utf-8")
        self.assertIn("LLM 配置错误", page)
        self.assertIn("TAS_LLM_API_KEY", page)

    def _write_fake_candidates(self, token: str) -> None:
        from tas.config import RuleSet
        from tas.web import app as webapp

        expected = dict(RuleSet.load().get("TITLE-001").expected)
        expected["size_pt"] = 14
        payload = {
            "base_profile": "school_default",
            "source_file": "spec.txt",
            "model": "test-model",
            "extraction": {
                "base_profile": "school_default",
                "candidates": [
                    {
                        "id": "TITLE-001",
                        "name": "中文题目",
                        "category": "front",
                        "phase": "P0",
                        "severity": "error",
                        "profiles": ["*"],
                        "spec": "测试：中文题目改为 14pt",
                        "expected": expected,
                        "source": "L72",
                        "confidence": 0.9,
                    }
                ],
                "unmapped": [],
                "dropped": [],
            },
        }
        (webapp.STORE_ROOT / token / "candidates.json").write_text(
            json.dumps(payload, ensure_ascii=False), encoding="utf-8"
        )

    def test_spec_apply_generates_profile(self):
        from tas.web import app as webapp

        token = webapp._new_token()
        self._write_fake_candidates(token)
        status, body, _ = self.post(
            "/spec/apply",
            {
                "token": token,
                "profile_id": "test_profile_web",
                "name": "测试规范",
                "discipline": "sci_eng",
                "register": "",
                "source_file": "spec.txt",
            },
            [],
        )
        self.assertEqual(status, 200, body[:300])
        page = body.decode("utf-8")
        self.assertIn("新 profile 已生成", page)

        status, data, _ = self.get(f"/download/{token}/profile.rules.json")
        self.assertEqual(status, 200)
        payload = json.loads(data)
        self.assertEqual(len(payload["rules"]), 38)
        rule = next(item for item in payload["rules"] if item["id"] == "TITLE-001")
        self.assertEqual(rule["expected"]["size_pt"], 14)

    def test_spec_apply_rejects_bad_profile_id(self):
        from tas.web import app as webapp

        token = webapp._new_token()
        self._write_fake_candidates(token)
        status, body, _ = self.post(
            "/spec/apply",
            {"token": token, "profile_id": "bad id!", "name": "x", "discipline": "sci_eng"},
            [],
        )
        self.assertEqual(status, 400)
        self.assertIn("profile id 不合法", body.decode("utf-8"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
