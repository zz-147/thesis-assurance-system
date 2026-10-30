"""LLM 客户端（tas.llm）回归测试。

覆盖 task-3 验收要求：

1. ``load_config``：key 优先级（TAS_LLM_API_KEY > DEEPSEEK_API_KEY）、DeepSeek 默认值、
   key 缺失抛 ``LLMConfigError``、环境变量覆盖、非法数字报错；
2. ``chat_json`` 请求构造：URL / Authorization / Content-Type / model / messages /
   response_format / temperature；不污染调用方传入的 messages；
3. 响应解析：正常 JSON、```json 围栏、夹带说明文字；非法 JSON / 缺 choices → 重试；
4. 错误路径：本地 http.server mock OpenAI 兼容端点覆盖 **成功 / 500 重试 / 429 重试 /
   401 直接失败 / 非法 JSON 重试**；纯 fetcher 注入覆盖网络错误 / 超时 / 重试后成功；
5. 安全：任何异常消息、调试输出都不含完整 key，最多出现 ``mask_key`` 打码形式；
6. live 测试默认 skip：仅 ``TAS_LIVE_LLM=1`` 且 ``TAS_LLM_API_KEY`` 存在时才真实调用。

全部测试离线：只用本地 127.0.0.1 http.server 与注入 fetcher，不依赖真实 key。
"""

from __future__ import annotations

import contextlib
import copy
import io
import json
import os
import sys
import threading
import unittest
import urllib.error
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import tas.llm.client as llm_client  # noqa: E402

from tas.llm import (  # noqa: E402
    LLMConfig,
    LLMConfigError,
    LLMError,
    chat_json,
    load_config,
    mask_key,
)
from tas.llm.client import (  # noqa: E402
    DEFAULT_BASE_URL,
    DEFAULT_MAX_RETRIES,
    DEFAULT_MODEL,
    DEFAULT_TEMPERATURE,
    DEFAULT_TIMEOUT,
    HTTPStatusError,
)

# 明显是假的测试 key（真实 key 永远不写进仓库/测试）
FAKE_KEY = "sk-deepseek-test-0123456789abcdef"
OTHER_KEY = "sk-deepseek-test-other-9876543210fedcba"

_LLM_ENV_KEYS = (
    "TAS_LLM_API_KEY",
    "DEEPSEEK_API_KEY",
    "TAS_LLM_BASE_URL",
    "TAS_LLM_MODEL",
    "TAS_LLM_TIMEOUT",
    "TAS_LLM_TEMPERATURE",
    "TAS_LLM_MAX_RETRIES",
    "TAS_LLM_JSON_MODE",
    "TAS_LLM_DEBUG",
    "TAS_LIVE_LLM",
    "NO_PROXY",
    "no_proxy",
)


# ---------------------------------------------------------------------------
# 测试工具
# ---------------------------------------------------------------------------
@contextlib.contextmanager
def llm_env(**values):
    """临时覆盖 LLM 相关环境变量（先清掉继承的，保证断言确定性）。"""
    cleaned = {key: value for key, value in os.environ.items() if key not in _LLM_ENV_KEYS}
    cleaned.update({key: str(value) for key, value in values.items()})
    with mock.patch.dict(os.environ, cleaned, clear=True):
        yield


@contextlib.contextmanager
def loopback_http_env():
    """确保访问 127.0.0.1 时不走系统代理。"""
    with mock.patch.dict(
        os.environ, {"NO_PROXY": "127.0.0.1,localhost", "no_proxy": "127.0.0.1,localhost"}
    ):
        yield


@contextlib.contextmanager
def patched_sleep():
    """替换退避 sleep，避免测试真的等 1s/2s。"""
    with mock.patch.object(llm_client, "_sleep") as sleep_mock:
        yield sleep_mock


def make_config(**overrides) -> LLMConfig:
    values = {"base_url": DEFAULT_BASE_URL, "model": DEFAULT_MODEL, "api_key": FAKE_KEY}
    values.update(overrides)
    return LLMConfig(**values)


def completion(content: str) -> str:
    """构造一个 OpenAI 兼容的成功响应体。"""
    return json.dumps(
        {
            "id": "chatcmpl-test",
            "object": "chat.completion",
            "created": 0,
            "model": DEFAULT_MODEL,
            "choices": [
                {
                    "index": 0,
                    "message": {"role": "assistant", "content": content},
                    "finish_reason": "stop",
                }
            ],
        },
        ensure_ascii=False,
    )


def sleep_args(sleep_mock) -> list[float]:
    return [call.args[0] for call in sleep_mock.call_args_list]


class RecordingFetcher:
    """按脚本返回/抛异常的假 fetcher，并记录每次调用参数。"""

    def __init__(self, results) -> None:
        self.calls: list[dict] = []
        self._results = list(results)

    def __call__(self, url, headers, payload, timeout):
        self.calls.append(
            {
                "url": url,
                "headers": dict(headers),
                "payload": payload,
                "timeout": timeout,
            }
        )
        if not self._results:
            raise AssertionError("fetcher 调用次数超出预期")
        result = self._results.pop(0)
        if isinstance(result, BaseException):
            raise result
        if callable(result):
            return result(url, headers, payload, timeout)
        return result


class _LLMTestCase(unittest.TestCase):
    """公共断言/夹具。"""

    def messages(self) -> list[dict]:
        return [{"role": "user", "content": "请从规范中抽取规则，只输出 JSON。"}]

    def assert_no_raw_key(self, text: str) -> None:
        self.assertNotIn(FAKE_KEY, text)
        self.assertNotIn(OTHER_KEY, text)


# ---------------------------------------------------------------------------
# 1) load_config
# ---------------------------------------------------------------------------
class LoadConfigTest(_LLMTestCase):
    def test_defaults_are_deepseek(self) -> None:
        with llm_env(TAS_LLM_API_KEY=FAKE_KEY):
            config = load_config()
        self.assertEqual(config.api_key, FAKE_KEY)
        self.assertEqual(config.base_url, "https://api.deepseek.com/v1")
        self.assertEqual(config.base_url, DEFAULT_BASE_URL)
        self.assertEqual(config.model, "deepseek-chat")
        self.assertEqual(config.model, DEFAULT_MODEL)
        self.assertEqual(config.timeout, 90.0)
        self.assertEqual(config.timeout, DEFAULT_TIMEOUT)
        self.assertEqual(config.max_retries, 2)
        self.assertEqual(config.max_retries, DEFAULT_MAX_RETRIES)
        self.assertEqual(config.temperature, 0.1)
        self.assertEqual(config.temperature, DEFAULT_TEMPERATURE)
        self.assertTrue(config.json_mode)

    def test_prefers_tas_key_over_deepseek_key(self) -> None:
        with llm_env(TAS_LLM_API_KEY=FAKE_KEY, DEEPSEEK_API_KEY=OTHER_KEY):
            config = load_config()
        self.assertEqual(config.api_key, FAKE_KEY)

    def test_falls_back_to_deepseek_key(self) -> None:
        with llm_env(DEEPSEEK_API_KEY=OTHER_KEY):
            config = load_config()
        self.assertEqual(config.api_key, OTHER_KEY)

    def test_missing_key_raises_config_error_with_actionable_hint(self) -> None:
        with llm_env():
            with self.assertRaises(LLMConfigError) as caught:
                load_config()
        message = str(caught.exception)
        self.assertIn("TAS_LLM_API_KEY", message)
        self.assertIn("新开终端", message)
        self.assertIn("DeepSeek", message)
        self.assertIsInstance(caught.exception, LLMError, "LLMConfigError 应可被 LLMError 兜底捕获")

    def test_blank_key_treated_as_missing(self) -> None:
        with llm_env(TAS_LLM_API_KEY="   ", DEEPSEEK_API_KEY=""):
            with self.assertRaises(LLMConfigError):
                load_config()

    def test_env_overrides(self) -> None:
        with llm_env(
            TAS_LLM_API_KEY=FAKE_KEY,
            TAS_LLM_BASE_URL="https://example.test/v1/",
            TAS_LLM_MODEL="deepseek-reasoner",
            TAS_LLM_TIMEOUT="12.5",
            TAS_LLM_TEMPERATURE="0.7",
            TAS_LLM_MAX_RETRIES="4",
            TAS_LLM_JSON_MODE="0",
        ):
            config = load_config()
        self.assertEqual(config.base_url, "https://example.test/v1")  # 尾部斜杠归一化
        self.assertEqual(config.model, "deepseek-reasoner")
        self.assertEqual(config.timeout, 12.5)
        self.assertEqual(config.temperature, 0.7)
        self.assertEqual(config.max_retries, 4)
        self.assertFalse(config.json_mode)

    def test_invalid_numeric_env_raises_config_error(self) -> None:
        with llm_env(TAS_LLM_API_KEY=FAKE_KEY, TAS_LLM_TIMEOUT="abc"):
            with self.assertRaises(LLMConfigError) as caught:
                load_config()
        self.assertIn("TAS_LLM_TIMEOUT", str(caught.exception))

    def test_invalid_base_url_scheme_raises_config_error(self) -> None:
        with llm_env(TAS_LLM_API_KEY=FAKE_KEY, TAS_LLM_BASE_URL="api.deepseek.com/v1"):
            with self.assertRaises(LLMConfigError):
                load_config()

    def test_config_dataclass_defaults_and_repr_hides_key(self) -> None:
        config = LLMConfig("https://api.deepseek.com/v1", "deepseek-chat", FAKE_KEY)
        self.assertEqual(config.timeout, 90.0)
        self.assertEqual(config.max_retries, 2)
        self.assertEqual(config.temperature, 0.1)
        self.assertTrue(config.json_mode)
        self.assertNotIn(FAKE_KEY, repr(config), "repr(config) 不能泄漏 key")
        self.assertNotIn("api_key", repr(config), "api_key 字段不参与 repr")

    def test_mask_key(self) -> None:
        self.assertEqual(mask_key(FAKE_KEY), "sk-…cdef")
        self.assertEqual(mask_key(OTHER_KEY), "sk-…dcba")
        self.assertEqual(mask_key("abcd1234efgh"), "ab…efgh")
        self.assertEqual(mask_key("abc"), "…")
        self.assertEqual(mask_key(""), "")
        self.assertEqual(mask_key(None), "")  # type: ignore[arg-type]
        self.assert_no_raw_key(mask_key(FAKE_KEY))


# ---------------------------------------------------------------------------
# 2) chat_json：纯 fetcher 注入（不发 HTTP）
# ---------------------------------------------------------------------------
class ChatJsonUnitTest(_LLMTestCase):
    def test_request_uses_deepseek_defaults(self) -> None:
        with llm_env(TAS_LLM_API_KEY=FAKE_KEY):
            config = load_config()
            fetcher = RecordingFetcher([completion('{"rules": []}')])
            result = chat_json(self.messages(), config=config, fetcher=fetcher)

        self.assertEqual(result, {"rules": []})
        self.assertEqual(len(fetcher.calls), 1)
        call = fetcher.calls[0]
        self.assertEqual(call["url"], "https://api.deepseek.com/v1/chat/completions")
        self.assertEqual(call["headers"]["Authorization"], f"Bearer {FAKE_KEY}")
        self.assertIn("application/json", call["headers"]["Content-Type"])
        self.assertEqual(call["payload"]["model"], "deepseek-chat")
        self.assertEqual(call["payload"]["temperature"], 0.1)
        self.assertEqual(call["payload"]["response_format"], {"type": "json_object"})
        self.assertEqual(call["timeout"], 90.0)
        # DeepSeek JSON Output 需要 prompt 里出现 "json" 字样：客户端自动补 hint
        system_text = call["payload"]["messages"][0]["content"]
        self.assertIn("Return a valid JSON object.", system_text)
        self.assertEqual(call["payload"]["messages"][-1], self.messages()[0])

    def test_custom_base_url_trailing_slash_handled(self) -> None:
        fetcher = RecordingFetcher([completion('{"ok": true}')])
        chat_json(
            self.messages(),
            config=make_config(base_url="https://example.test/v1/"),
            fetcher=fetcher,
        )
        self.assertEqual(fetcher.calls[0]["url"], "https://example.test/v1/chat/completions")

    def test_json_mode_false_without_schema_keeps_messages(self) -> None:
        messages = self.messages()
        fetcher = RecordingFetcher([completion('{"ok": true}')])
        chat_json(messages, config=make_config(json_mode=False), fetcher=fetcher)
        payload = fetcher.calls[0]["payload"]
        self.assertNotIn("response_format", payload)
        self.assertEqual(payload["messages"], messages, "未启用 json_mode 时不应改写 messages")

    def test_schema_hint_included_and_triggers_json_hint(self) -> None:
        schema_hint = {"type": "object", "properties": {"rules": {"type": "array"}}}
        fetcher = RecordingFetcher([completion('{"rules": []}')])
        chat_json(
            self.messages(),
            config=make_config(json_mode=False),
            schema_hint=schema_hint,
            fetcher=fetcher,
        )
        payload = fetcher.calls[0]["payload"]
        self.assertNotIn("response_format", payload)  # json_mode=False 时不加 response_format
        system_text = payload["messages"][0]["content"]
        self.assertIn("Return a valid JSON object.", system_text)
        self.assertIn("Response JSON shape hint:", system_text)
        self.assertIn('"rules"', system_text)

    def test_existing_system_message_gets_hint_appended(self) -> None:
        messages = [
            {"role": "system", "content": "你是论文格式规范抽取助手。"},
            {"role": "user", "content": "抽取规则"},
        ]
        fetcher = RecordingFetcher([completion('{"ok": true}')])
        chat_json(messages, config=make_config(), fetcher=fetcher)
        prepared = fetcher.calls[0]["payload"]["messages"]
        self.assertEqual([message["role"] for message in prepared], ["system", "user"])
        self.assertTrue(prepared[0]["content"].startswith("你是论文格式规范抽取助手。"))
        self.assertIn("Return a valid JSON object.", prepared[0]["content"])

    def test_json_hint_not_duplicated(self) -> None:
        messages = [
            {"role": "system", "content": "Return a valid JSON object."},
            {"role": "user", "content": "hi"},
        ]
        fetcher = RecordingFetcher([completion('{"ok": true}')])
        chat_json(messages, config=make_config(), fetcher=fetcher)
        system_text = fetcher.calls[0]["payload"]["messages"][0]["content"]
        self.assertEqual(system_text.count("Return a valid JSON object."), 1)

    def test_no_system_message_inserts_one(self) -> None:
        messages = self.messages()
        fetcher = RecordingFetcher([completion('{"ok": true}')])
        chat_json(messages, config=make_config(), fetcher=fetcher)
        prepared = fetcher.calls[0]["payload"]["messages"]
        self.assertEqual(prepared[0]["role"], "system")
        self.assertIn("Return a valid JSON object.", prepared[0]["content"])
        self.assertEqual(prepared[1:], messages)

    def test_caller_messages_not_mutated(self) -> None:
        messages = [
            {"role": "system", "content": "你是助手。"},
            {"role": "user", "content": "抽取规则"},
        ]
        snapshot = copy.deepcopy(messages)
        fetcher = RecordingFetcher([completion('{"ok": true}')])
        chat_json(messages, config=make_config(), fetcher=fetcher)
        self.assertEqual(messages, snapshot, "chat_json 不能修改调用方传入的 messages")

    def test_empty_messages_raises_value_error(self) -> None:
        with self.assertRaises(ValueError):
            chat_json([], config=make_config())

    # -- 响应解析容错 -----------------------------------------------------
    def test_fenced_json_content_is_parsed(self) -> None:
        content = '```json\n{"rules": [{"id": "BODY-001"}]}\n```'
        fetcher = RecordingFetcher([completion(content)])
        result = chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(result, {"rules": [{"id": "BODY-001"}]})

    def test_plain_fenced_json_without_language(self) -> None:
        fetcher = RecordingFetcher([completion('```\n{"ok": true}\n```')])
        result = chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(result, {"ok": True})

    def test_content_with_prose_still_parsed(self) -> None:
        fetcher = RecordingFetcher([completion('以下是抽取结果：\n{"ok": true}\n请确认。')])
        result = chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(result, {"ok": True})

    # -- 重试/错误路径 ----------------------------------------------------
    def test_500_then_success_retries_once(self) -> None:
        fetcher = RecordingFetcher(
            [HTTPStatusError(500, '{"error": "boom"}'), completion('{"ok": true}')]
        )
        with patched_sleep() as sleep_mock:
            result = chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(result, {"ok": True})
        self.assertEqual(len(fetcher.calls), 2)
        self.assertEqual(sleep_args(sleep_mock), [1.0])

    def test_500_exhausts_retries_with_mask_only(self) -> None:
        failure = HTTPStatusError(500, '{"error": "server boom"}')
        fetcher = RecordingFetcher([failure, failure, failure])
        with patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        message = str(caught.exception)
        self.assertEqual(len(fetcher.calls), 3, "max_retries=2 → 最多 3 次请求")
        self.assertEqual(sleep_args(sleep_mock), [1.0, 2.0], "指数退避 1s/2s")
        self.assertIn("500", message)
        self.assertIn("已重试 2 次", message)
        self.assert_no_raw_key(message)
        self.assertIn(mask_key(FAKE_KEY), message)

    def test_429_is_retried(self) -> None:
        failure = HTTPStatusError(429, '{"error": "rate limit"}')
        fetcher = RecordingFetcher([failure, failure, failure])
        with patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(len(fetcher.calls), 3)
        self.assertEqual(sleep_args(sleep_mock), [1.0, 2.0])
        self.assertIn("429", str(caught.exception))

    def test_401_fails_immediately_without_retry(self) -> None:
        fetcher = RecordingFetcher(
            [HTTPStatusError(401, '{"error": {"message": "invalid api key"}}')]
        )
        with patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        message = str(caught.exception)
        self.assertEqual(len(fetcher.calls), 1, "4xx（非 429）不重试")
        sleep_mock.assert_not_called()
        self.assertIn("401", message)
        self.assert_no_raw_key(message)
        self.assertIn(mask_key(FAKE_KEY), message)

    def test_max_retries_zero_single_attempt(self) -> None:
        fetcher = RecordingFetcher([HTTPStatusError(500, "boom")])
        with patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(
                    self.messages(), config=make_config(max_retries=0), fetcher=fetcher
                )
        self.assertEqual(len(fetcher.calls), 1)
        sleep_mock.assert_not_called()
        self.assertNotIn("已重试", str(caught.exception))

    def test_network_error_then_success(self) -> None:
        fetcher = RecordingFetcher(
            [urllib.error.URLError("connection refused"), completion('{"ok": true}')]
        )
        with patched_sleep() as sleep_mock:
            result = chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(result, {"ok": True})
        self.assertEqual(len(fetcher.calls), 2)
        self.assertEqual(sleep_args(sleep_mock), [1.0])

    def test_network_error_exhausted(self) -> None:
        fetcher = RecordingFetcher([urllib.error.URLError("down")] * 3)
        with patched_sleep():
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        message = str(caught.exception)
        self.assertEqual(len(fetcher.calls), 3)
        self.assertIn("网络错误", message)
        self.assert_no_raw_key(message)

    def test_timeout_then_success(self) -> None:
        fetcher = RecordingFetcher([TimeoutError("timed out"), completion('{"ok": true}')])
        with patched_sleep() as sleep_mock:
            result = chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(result, {"ok": True})
        self.assertEqual(len(fetcher.calls), 2)
        self.assertEqual(sleep_args(sleep_mock), [1.0])

    def test_invalid_json_content_retries_then_llm_error(self) -> None:
        fetcher = RecordingFetcher([completion("这不是 JSON")] * 3)
        with patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        message = str(caught.exception)
        self.assertEqual(len(fetcher.calls), 3)
        self.assertEqual(sleep_args(sleep_mock), [1.0, 2.0])
        self.assertIn("content 不是合法 JSON", message)
        self.assertIn("已重试 2 次", message)
        self.assert_no_raw_key(message)

    def test_invalid_json_then_success(self) -> None:
        fetcher = RecordingFetcher([completion("nope"), completion('{"ok": true}')])
        with patched_sleep() as sleep_mock:
            result = chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(result, {"ok": True})
        self.assertEqual(sleep_args(sleep_mock), [1.0])

    def test_non_object_json_content_retries_then_error(self) -> None:
        fetcher = RecordingFetcher([completion("[]")] * 3)
        with patched_sleep():
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(len(fetcher.calls), 3)
        self.assertIn("顶层不是 JSON 对象", str(caught.exception))

    def test_invalid_http_response_body_retries(self) -> None:
        fetcher = RecordingFetcher(["<html>not json</html>"] * 3)
        with patched_sleep():
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(len(fetcher.calls), 3)
        self.assertIn("HTTP 响应不是合法 JSON", str(caught.exception))

    def test_missing_choices_retries(self) -> None:
        fetcher = RecordingFetcher([json.dumps({"id": "x"})] * 3)
        with patched_sleep():
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(len(fetcher.calls), 3)
        self.assertIn("缺少 choices", str(caught.exception))

    def test_error_payload_with_http_200_retries(self) -> None:
        body = json.dumps({"error": {"message": "quota exceeded"}})
        fetcher = RecordingFetcher([body] * 3)
        with patched_sleep():
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        self.assertEqual(len(fetcher.calls), 3)
        self.assertIn("接口返回错误", str(caught.exception))
        self.assertIn("quota exceeded", str(caught.exception))

    def test_debug_output_contains_no_secret(self) -> None:
        fetcher = RecordingFetcher([completion('{"ok": true}')])
        stderr = io.StringIO()
        with llm_env(TAS_LLM_DEBUG="1"):
            with contextlib.redirect_stderr(stderr):
                chat_json(self.messages(), config=make_config(), fetcher=fetcher)
        output = stderr.getvalue()
        self.assertIn("[tas.llm]", output)
        self.assertIn("model=deepseek-chat", output)
        self.assert_no_raw_key(output)
        self.assertNotIn("Bearer", output)
        self.assertNotIn("Authorization", output)


# ---------------------------------------------------------------------------
# 3) chat_json：本地 http.server mock OpenAI 兼容端点
# ---------------------------------------------------------------------------
class _CapturedRequest:
    def __init__(self, path: str, headers: dict, payload) -> None:
        self.path = path
        self.headers = headers
        self.payload = payload

    def header(self, name: str):
        lowered = name.lower()
        for key, value in self.headers.items():
            if key.lower() == lowered:
                return value
        return None


class _MockLLMHandler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:  # noqa: N802 - http.server 约定
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception:
            payload = None
        server = self.server
        with server.lock:
            server.requests.append(
                _CapturedRequest(self.path, dict(self.headers.items()), payload)
            )
            if server.responses:
                status, body, content_type = server.responses.pop(0)
            else:
                status, body, content_type = server.default_response
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args) -> None:  # 静音测试日志
        pass


class _MockLLMServer:
    """本地 mock OpenAI 兼容端点（127.0.0.1 随机端口）。"""

    def __init__(self, *, default=None, responses=None) -> None:
        self._httpd = ThreadingHTTPServer(("127.0.0.1", 0), _MockLLMHandler)
        self._httpd.daemon_threads = True
        self._httpd.lock = threading.Lock()
        self._httpd.requests = []
        self._httpd.default_response = default or (
            200,
            completion('{"ok": true}'),
            "application/json",
        )
        self._httpd.responses = list(responses or [])
        host, port = self._httpd.server_address[:2]
        self.base_url = f"http://{host}:{port}/v1"
        self._thread = threading.Thread(
            target=self._httpd.serve_forever,
            kwargs={"poll_interval": 0.02},
            daemon=True,
            name="mock-llm-server",
        )

    def __enter__(self) -> "_MockLLMServer":
        self._thread.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        self._httpd.shutdown()
        self._httpd.server_close()
        self._thread.join(timeout=5)
        return False

    @property
    def requests(self) -> list[_CapturedRequest]:
        with self._httpd.lock:
            return list(self._httpd.requests)


class ChatJsonHttpTest(_LLMTestCase):
    def test_success_and_request_shape(self) -> None:
        server = _MockLLMServer(
            default=(200, completion('{"rules": [{"rule_id": "BODY-001"}]}'), "application/json")
        )
        with loopback_http_env(), server:
            config = make_config(base_url=server.base_url)
            result = chat_json([{"role": "user", "content": "请抽取规则"}], config=config)
            requests = server.requests

        self.assertEqual(result, {"rules": [{"rule_id": "BODY-001"}]})
        self.assertEqual(len(requests), 1)
        request = requests[0]
        self.assertEqual(request.path, "/v1/chat/completions")
        self.assertEqual(request.header("Authorization"), f"Bearer {FAKE_KEY}")
        self.assertIn("application/json", request.header("Content-Type"))
        payload = request.payload
        self.assertEqual(payload["model"], "deepseek-chat")
        self.assertEqual(payload["temperature"], 0.1)
        self.assertEqual(payload["response_format"], {"type": "json_object"})
        self.assertEqual(payload["messages"][-1], {"role": "user", "content": "请抽取规则"})
        self.assertIn("Return a valid JSON object.", payload["messages"][0]["content"])

    def test_fenced_json_over_http(self) -> None:
        server = _MockLLMServer(
            default=(200, completion('```json\n{"ok": true}\n```'), "application/json")
        )
        with loopback_http_env(), server:
            result = chat_json(self.messages(), config=make_config(base_url=server.base_url))
        self.assertEqual(result, {"ok": True})

    def test_invalid_json_retries_over_http(self) -> None:
        server = _MockLLMServer(
            default=(200, completion("这不是 JSON"), "application/json")
        )
        with loopback_http_env(), server, patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(base_url=server.base_url))
            requests = server.requests
        self.assertEqual(len(requests), 3)
        self.assertEqual(sleep_args(sleep_mock), [1.0, 2.0])
        self.assertIn("content 不是合法 JSON", str(caught.exception))

    def test_500_retries_over_http(self) -> None:
        server = _MockLLMServer(
            default=(500, '{"error": {"message": "server boom"}}', "application/json")
        )
        with loopback_http_env(), server, patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(base_url=server.base_url))
            requests = server.requests
        message = str(caught.exception)
        self.assertEqual(len(requests), 3)
        self.assertEqual(sleep_args(sleep_mock), [1.0, 2.0])
        self.assertIn("HTTP 500", message)
        self.assert_no_raw_key(message)
        self.assertIn(mask_key(FAKE_KEY), message)

    def test_401_fails_without_retry_over_http(self) -> None:
        server = _MockLLMServer(
            default=(401, '{"error": {"message": "invalid api key"}}', "application/json")
        )
        with loopback_http_env(), server, patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError) as caught:
                chat_json(self.messages(), config=make_config(base_url=server.base_url))
            requests = server.requests
        message = str(caught.exception)
        self.assertEqual(len(requests), 1)
        sleep_mock.assert_not_called()
        self.assertIn("HTTP 401", message)
        self.assert_no_raw_key(message)
        self.assertIn(mask_key(FAKE_KEY), message)

    def test_429_retries_over_http(self) -> None:
        server = _MockLLMServer(
            default=(429, '{"error": {"message": "rate limit"}}', "application/json")
        )
        with loopback_http_env(), server, patched_sleep() as sleep_mock:
            with self.assertRaises(LLMError):
                chat_json(self.messages(), config=make_config(base_url=server.base_url))
            requests = server.requests
        self.assertEqual(len(requests), 3)
        self.assertEqual(sleep_args(sleep_mock), [1.0, 2.0])

    def test_retry_then_success_over_http(self) -> None:
        server = _MockLLMServer(
            responses=[
                (500, '{"error": "boom"}', "application/json"),
                (200, completion('{"ok": true}'), "application/json"),
            ]
        )
        with loopback_http_env(), server, patched_sleep() as sleep_mock:
            result = chat_json(self.messages(), config=make_config(base_url=server.base_url))
            requests = server.requests
        self.assertEqual(result, {"ok": True})
        self.assertEqual(len(requests), 2)
        self.assertEqual(sleep_args(sleep_mock), [1.0])


# ---------------------------------------------------------------------------
# 4) live 测试（默认 skip；显式 TAS_LIVE_LLM=1 且 TAS_LLM_API_KEY 存在才跑）
# ---------------------------------------------------------------------------
@unittest.skipUnless(
    os.environ.get("TAS_LIVE_LLM") == "1" and bool(os.environ.get("TAS_LLM_API_KEY")),
    "live LLM 测试默认跳过：设置 TAS_LIVE_LLM=1 且 TAS_LLM_API_KEY 后才跑（会产生真实计费调用）",
)
class LiveLLMTest(_LLMTestCase):
    def test_live_json_roundtrip(self) -> None:
        config = load_config()
        result = chat_json(
            [{"role": "user", "content": '只输出一个 JSON 对象：{"ok": true}'}],
            config=config,
            schema_hint={"ok": "boolean"},
        )
        self.assertIsInstance(result, dict)


if __name__ == "__main__":
    unittest.main(verbosity=2)
