"""LLM 客户端：OpenAI 兼容 ``/chat/completions``（默认 DeepSeek），只产出结构化 JSON。

设计约定（与 docs/REQUIREMENTS.md 的红线一致）：

- LLM 只负责「规范文本 → 结构化规则候选」这类结构化输出，**绝不直接改文档**；
  ``chat_json`` 统一返回 ``dict``，调用方自行做 schema 校验、差异对比与教师确认；
- 只用标准库（``urllib``），不引入第三方依赖；
- API key 只从环境变量读取，**不进日志、不进异常消息**；异常里最多出现
  ``mask_key`` 打码后的形式（如 ``sk-…abcd``）。

默认服务商：DeepSeek（OpenAI 兼容模式）

- ``base_url``：``https://api.deepseek.com/v1``（可用 ``TAS_LLM_BASE_URL`` 覆盖）
- ``model``：``deepseek-chat``（可用 ``TAS_LLM_MODEL`` 覆盖）
- key：``TAS_LLM_API_KEY`` 优先，回退 ``DEEPSEEK_API_KEY``

DeepSeek 的 JSON Output 要求 prompt 中出现 "json" 字样才会生效，而客户端无法控制
调用方 prompt，所以当 ``json_mode=True`` 或传了 ``schema_hint`` 时，客户端会自动在
system 消息末尾补一句 ``Return a valid JSON object.``（已存在则不重复添加）。

可测试性：``chat_json(..., fetcher=...)`` 可注入
``fetcher(url, headers, payload, timeout) -> str``（返回响应原文）。注入的 fetcher 用
``HTTPStatusError(status, body)`` 回报 HTTP 非 200，抛其他异常则视为网络错误。
单元测试全部离线（本地 http.server 或纯 fetcher 注入），不依赖真实 key。

调试：``TAS_LLM_DEBUG=1`` 时向 stderr 打印请求摘要（url/model/消息数），
**绝不打印 key 或 Authorization 头**。
"""

from __future__ import annotations

import json
import os
import sys
import time
import urllib.error
import urllib.request
from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

# ---------------------------------------------------------------------------
# 默认值（DeepSeek / OpenAI 兼容）
# ---------------------------------------------------------------------------
DEFAULT_BASE_URL = "https://api.deepseek.com/v1"
DEFAULT_MODEL = "deepseek-chat"
DEFAULT_TIMEOUT = 90.0
DEFAULT_MAX_RETRIES = 2
DEFAULT_TEMPERATURE = 0.1
DEFAULT_MAX_TOKENS = 8192

#: key 读取优先级：项目专用变量 > 服务商通用变量
API_KEY_ENV_NAMES = ("TAS_LLM_API_KEY", "DEEPSEEK_API_KEY")

#: DeepSeek JSON Output 要求 prompt 里出现 "json" 字样；这句由客户端自动补。
JSON_HINT_SENTENCE = "Return a valid JSON object."

BACKOFF_BASE_SECONDS = 1.0
BACKOFF_MAX_SECONDS = 8.0
_ERROR_PREVIEW_LIMIT = 200

# fetcher 约定：输入 URL、请求头、payload、超时（秒），返回响应原文；
# HTTP 非 200 抛 HTTPStatusError，其他失败抛任意异常（视为网络错误）。
Fetcher = Callable[[str, dict[str, str], dict[str, Any], float], str]

_DEBUG_VALUES = {"1", "true", "yes", "on"}
_TRUE_VALUES = {"1", "true", "yes", "on"}
_FALSE_VALUES = {"0", "false", "no", "off"}


# ---------------------------------------------------------------------------
# 异常
# ---------------------------------------------------------------------------
class LLMError(Exception):
    """LLM 调用失败（网络 / HTTP 非 200 / 返回非法 JSON / 超时）。

    异常消息中**绝不会出现完整 api_key**，最多出现 ``mask_key`` 打码形式。
    """


class LLMConfigError(LLMError):
    """LLM 配置错误（典型场景：环境变量里没有 API key）。

    继承 ``LLMError``，便于调用方统一 ``except LLMError`` 兜底；
    需要区分「配置问题」与「调用问题」时再单独捕获本类。
    """


class HTTPStatusError(Exception):
    """HTTP 非 200 状态（供注入 fetcher 回报；默认 fetcher 也用它包装 HTTPError）。"""

    def __init__(self, status_code: int, body: str = "", reason: str = "") -> None:
        self.status_code = int(status_code)
        self.body = body or ""
        self.reason = reason or ""
        super().__init__(self.summary())

    def summary(self) -> str:
        label = f"HTTP {self.status_code}"
        if self.reason:
            label = f"{label} {self.reason}"
        snippet = _preview(self.body)
        return f"{label}：{snippet}" if snippet else label


class _RetryableError(Exception):
    """内部使用：响应可解析性错误（HTTP 响应/模型 content 非合法 JSON）。"""


@dataclass(frozen=True)
class _Failure:
    """内部使用：一次请求失败的结构化描述。"""

    message: str
    retryable: bool
    include_key: bool = False


# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
@dataclass
class LLMConfig:
    """一次 LLM 调用的配置。

    ``api_key`` 标记了 ``repr=False``：打印/记录配置对象不会泄漏 key。
    """

    base_url: str
    model: str
    api_key: str = field(repr=False)
    timeout: float = DEFAULT_TIMEOUT
    max_retries: int = DEFAULT_MAX_RETRIES
    temperature: float = DEFAULT_TEMPERATURE
    json_mode: bool = True
    # 单次输出的 token 上限：规范抽取可能要输出 30~50 条规则 JSON，
    # DeepSeek 默认 4096 容易截断，这里显式放大到 8192（模型上限）。
    max_tokens: int = DEFAULT_MAX_TOKENS


def mask_key(key: str) -> str:
    """把 key 打码成可安全展示的形式（如 ``sk-…abcd``）。

    过长/过短的 key 都能安全打码；空值返回空串，绝不在返回值中保留完整 key。
    """
    if not key:
        return ""
    text = str(key).strip()
    if not text:
        return ""
    if len(text) <= 4:
        return "…"
    if len(text) <= 8:
        return "…" + text[-2:]
    prefix = "sk-" if text.startswith("sk-") else text[:2]
    return f"{prefix}…{text[-4:]}"


def load_config() -> LLMConfig:
    """从环境变量读取 LLM 配置；没有 key 时抛 :class:`LLMConfigError`。

    环境变量：

    - ``TAS_LLM_API_KEY``（优先）> ``DEEPSEEK_API_KEY``
    - ``TAS_LLM_BASE_URL``（默认 ``https://api.deepseek.com/v1``）
    - ``TAS_LLM_MODEL``（默认 ``deepseek-chat``）
    - ``TAS_LLM_TIMEOUT`` / ``TAS_LLM_TEMPERATURE``（数字）
    - ``TAS_LLM_MAX_RETRIES``（整数，默认 2）
    - ``TAS_LLM_MAX_TOKENS``（整数，默认 8192；规范抽取输出较长，需大于默认 4096）
    - ``TAS_LLM_JSON_MODE``（0/1，默认 1）
    """
    api_key = _first_env(*API_KEY_ENV_NAMES)
    if not api_key:
        raise LLMConfigError(
            "未找到 DeepSeek API key：请设置环境变量 TAS_LLM_API_KEY（DeepSeek key），"
            "设置后需要新开终端才生效。PowerShell 示例：\n"
            '  $env:TAS_LLM_API_KEY = "sk-..."\n'
            "也兼容 DEEPSEEK_API_KEY（优先级低于 TAS_LLM_API_KEY）。"
        )

    base_url = (os.environ.get("TAS_LLM_BASE_URL") or "").strip().rstrip("/")
    if not base_url:
        base_url = DEFAULT_BASE_URL
    if not base_url.startswith(("http://", "https://")):
        raise LLMConfigError(
            f"TAS_LLM_BASE_URL 必须是 http(s) 地址，当前值：{base_url!r}"
        )

    model = (os.environ.get("TAS_LLM_MODEL") or "").strip() or DEFAULT_MODEL
    return LLMConfig(
        base_url=base_url,
        model=model,
        api_key=api_key,
        timeout=_env_float("TAS_LLM_TIMEOUT", DEFAULT_TIMEOUT),
        max_retries=_env_int("TAS_LLM_MAX_RETRIES", DEFAULT_MAX_RETRIES),
        temperature=_env_float("TAS_LLM_TEMPERATURE", DEFAULT_TEMPERATURE),
        json_mode=_env_bool("TAS_LLM_JSON_MODE", True),
        max_tokens=_env_int("TAS_LLM_MAX_TOKENS", DEFAULT_MAX_TOKENS),
    )


# ---------------------------------------------------------------------------
# 对外主入口
# ---------------------------------------------------------------------------
def chat_json(
    messages: list[dict],
    *,
    config: Optional[LLMConfig] = None,
    schema_hint: Optional[dict] = None,
    fetcher: Optional[Fetcher] = None,
) -> dict:
    """调用 OpenAI 兼容的 ``/chat/completions``，解析并返回 JSON 对象。

    - ``messages``：OpenAI 风格消息列表（不改动调用方传入的列表/字典）；
    - ``config``：缺省时调用 :func:`load_config`；
    - ``schema_hint``：非空时写入 prompt 的 system 消息，并强制补 JSON 输出提示；
    - ``fetcher``：测试注入点，签名见模块文档。

    重试策略（最多 ``config.max_retries`` 次重试，指数退避 1s/2s）：

    - 网络错误 / 超时 / HTTP 5xx / HTTP 429 / 响应非法 JSON → 重试；
    - 其余 4xx → 直接抛 :class:`LLMError`，不重试；
    - 重试耗尽后统一抛 :class:`LLMError`，消息中 key 只以 :func:`mask_key` 形式出现。
    """
    if not isinstance(messages, (list, tuple)) or not messages:
        raise ValueError("chat_json(messages=...) 需要非空的 messages 列表")

    cfg = config if config is not None else load_config()
    fetch = fetcher if fetcher is not None else default_fetcher
    url = _chat_completions_url(cfg.base_url)
    payload = _build_payload(cfg, messages, schema_hint=schema_hint)
    headers = _build_headers(cfg)
    attempts = _attempt_count(cfg.max_retries)

    _debug(
        f"POST {url} model={cfg.model} messages={len(payload['messages'])} "
        f"timeout={cfg.timeout:g}s json_mode={cfg.json_mode} "
        f"max_retries={attempts - 1}"
    )

    last_failure: Optional[_Failure] = None
    for attempt in range(attempts):
        try:
            raw = fetch(url, headers, payload, cfg.timeout)
        except HTTPStatusError as exc:
            failure = _http_failure(exc.status_code, exc.reason, exc.body)
        except LLMError:
            raise
        except Exception as exc:  # noqa: BLE001 - 网络/超时/自定义 fetcher 异常统一兜住
            failure = _transport_failure(exc)
        else:
            try:
                result = _parse_chat_response(raw)
            except _RetryableError as exc:
                failure = _Failure(str(exc), retryable=True)
            else:
                _debug(f"响应成功：{len(raw or '')} 字符")
                return result

        last_failure = failure
        if not failure.retryable:
            raise LLMError(
                _final_message(cfg, failure.message, failure=failure, retries=0)
            )
        if attempt < attempts - 1:
            delay = _backoff_delay(attempt)
            _debug(
                f"第 {attempt + 1}/{attempts} 次请求失败，{delay:g}s 后重试："
                f"{_redact(failure.message, cfg.api_key)}"
            )
            _sleep(delay)
            continue
        raise LLMError(
            _final_message(cfg, failure.message, failure=failure, retries=attempt)
        )

    # 理论上不可达（attempts >= 1，循环内必然 return 或 raise），兜底保证函数返回值类型。
    raise LLMError(  # pragma: no cover
        _final_message(
            cfg,
            last_failure.message if last_failure else "未知错误",
            failure=last_failure or _Failure("未知错误", retryable=True),
            retries=attempts - 1,
        )
    )


# ---------------------------------------------------------------------------
# 请求构造
# ---------------------------------------------------------------------------
def _chat_completions_url(base_url: str) -> str:
    return (base_url or DEFAULT_BASE_URL).rstrip("/") + "/chat/completions"


def _build_headers(config: LLMConfig) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {config.api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def _build_payload(
    config: LLMConfig,
    messages: list[dict],
    *,
    schema_hint: Optional[dict] = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": config.model,
        "messages": _prepare_messages(
            messages, schema_hint=schema_hint, json_mode=config.json_mode
        ),
        "temperature": config.temperature,
    }
    if config.max_tokens and config.max_tokens > 0:
        payload["max_tokens"] = config.max_tokens
    if config.json_mode:
        payload["response_format"] = {"type": "json_object"}
    return payload


def _prepare_messages(
    messages: list[dict],
    *,
    schema_hint: Optional[dict],
    json_mode: bool,
) -> list[Any]:
    """复制 messages 并按需追加 JSON 输出提示（不改动调用方数据）。

    DeepSeek JSON Output 要求 prompt 里出现 "json" 字样：``json_mode=True`` 或
    给了 ``schema_hint`` 时，在 system 消息末尾补 ``Return a valid JSON object.``
    （已有则不重复添加）；没有 system 消息时在最前面插入一条。
    """
    prepared: list[Any] = [
        dict(message) if isinstance(message, Mapping) else message
        for message in messages
    ]
    if not (json_mode or schema_hint):
        return prepared

    additions: list[str] = []
    if schema_hint:
        additions.append("Response JSON shape hint: " + _compact_json(schema_hint))
    if not _has_json_hint(prepared):
        additions.append(JSON_HINT_SENTENCE)
    if not additions:
        return prepared

    hint = "\n".join(additions)
    for message in prepared:
        if not isinstance(message, dict) or message.get("role") != "system":
            continue
        content = message.get("content")
        if isinstance(content, str):
            message["content"] = (
                f"{content.rstrip()}\n{hint}" if content.strip() else hint
            )
            return prepared
        break  # system 消息 content 非字符串：不硬改，退化为在最前面插入新 system
    prepared.insert(0, {"role": "system", "content": hint})
    return prepared


def _has_json_hint(messages: list[Any]) -> bool:
    needle = JSON_HINT_SENTENCE.lower()
    for message in messages:
        content = message.get("content") if isinstance(message, dict) else None
        if isinstance(content, str) and needle in content.lower():
            return True
    return False


def _compact_json(value: Any) -> str:
    try:
        return json.dumps(
            value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        )
    except (TypeError, ValueError):
        return str(value)


# ---------------------------------------------------------------------------
# 响应解析
# ---------------------------------------------------------------------------
def _parse_chat_response(raw: str) -> dict:
    """解析 OpenAI 兼容响应体 → choices[0].message.content → JSON 对象。"""
    text = (raw or "").strip()
    if not text:
        raise _RetryableError("HTTP 响应为空")
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise _RetryableError(
            f"HTTP 响应不是合法 JSON（{exc.msg}，位置 {exc.pos}）：{_preview(text)}"
        ) from None
    if not isinstance(data, dict):
        raise _RetryableError(
            f"HTTP 响应 JSON 顶层不是对象而是 {type(data).__name__}：{_preview(text)}"
        )

    error = data.get("error")
    if error and not data.get("choices"):
        raise _RetryableError(f"接口返回错误：{_error_text(error)}")

    choices = data.get("choices")
    if not isinstance(choices, list) or not choices:
        raise _RetryableError(f"HTTP 响应缺少 choices 字段：{_preview(text)}")
    first = choices[0]
    message = first.get("message") if isinstance(first, dict) else None
    content = message.get("content") if isinstance(message, dict) else None

    if isinstance(content, dict):  # 少数兼容实现直接返回对象：容错接受
        return content
    if not isinstance(content, str) or not content.strip():
        raise _RetryableError(
            f"HTTP 响应 choices[0].message.content 为空或不是字符串：{_preview(text)}"
        )
    return _loads_json_object(content)


def _loads_json_object(content: str) -> dict:
    """把模型 content 解析成 dict；容错剥离 ```json 围栏与前后说明文字。"""
    text = content.strip().lstrip("\ufeff")
    candidates: list[str] = [text]
    fenced = _strip_code_fence(text)
    if fenced and fenced != text:
        candidates.append(fenced)
    extracted = _extract_json_object(fenced or text)
    if extracted and extracted not in candidates:
        candidates.append(extracted)

    last_error: Optional[Exception] = None
    for candidate in candidates:
        try:
            data = json.loads(candidate)
        except json.JSONDecodeError as exc:
            last_error = exc
            continue
        if isinstance(data, dict):
            return data
        last_error = ValueError(f"顶层不是 JSON 对象而是 {type(data).__name__}")
    raise _RetryableError(
        f"模型返回 content 不是合法 JSON 对象（{last_error}）：{_preview(content)}"
    )


def _strip_code_fence(text: str) -> str:
    """剥离 ```json ... ``` / ``` ... ``` 代码围栏；不是围栏则原样返回。"""
    stripped = text.strip()
    if not stripped.startswith("```"):
        return stripped
    lines = stripped.splitlines()
    language = lines[0].strip().lstrip("`").strip().lower()
    if language not in ("", "json"):
        return stripped
    body = lines[1:]
    while body and not body[0].strip():
        body.pop(0)
    if body and body[-1].strip().startswith("```"):
        body = body[:-1]
    return "\n".join(body).strip()


def _extract_json_object(text: str) -> str:
    """兜底：截取第一个 ``{`` 到最后一个 ``}`` 之间内容（模型偶尔夹带说明文字）。"""
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        return ""
    return text[start : end + 1]


# ---------------------------------------------------------------------------
# 失败描述 / 重试
# ---------------------------------------------------------------------------
def _http_failure(status: int, reason: str, body: str) -> _Failure:
    label = f"HTTP {status}"
    if reason:
        label = f"{label} {reason}"
    snippet = _preview(body)
    detail = f"{label}：{snippet}" if snippet else label
    if status == 429:
        return _Failure(f"{detail}（触发限流，可稍后重试）", retryable=True, include_key=True)
    if status >= 500:
        return _Failure(detail, retryable=True, include_key=True)
    if status in (401, 403):
        detail += (
            "（认证失败：请检查 TAS_LLM_API_KEY / DEEPSEEK_API_KEY 是否正确、"
            "是否已开通 DeepSeek API 并有余额）"
        )
    return _Failure(detail, retryable=False, include_key=True)


def _transport_failure(exc: Exception) -> _Failure:
    if isinstance(exc, urllib.error.HTTPError):
        return _http_failure(
            int(exc.code), str(exc.reason or ""), _read_http_error_body(exc)
        )
    name = type(exc).__name__
    if isinstance(exc, TimeoutError) or "timeout" in name.lower():
        return _Failure(f"请求超时（{name}）：{exc}", retryable=True)
    return _Failure(f"网络错误（{name}）：{exc}", retryable=True)


def _final_message(
    config: LLMConfig, detail: str, *, failure: _Failure, retries: int
) -> str:
    if failure.include_key:
        masked = mask_key(config.api_key)
        if masked:
            detail = f"{detail}（key={masked}）"
    if retries > 0:
        head = f"LLM 调用失败（已重试 {retries} 次，共 {retries + 1} 次请求）：{detail}"
    else:
        head = f"LLM 调用失败：{detail}"
    return _redact(head, config.api_key)


def _attempt_count(max_retries: Any) -> int:
    try:
        retries = int(max_retries)
    except (TypeError, ValueError):
        retries = DEFAULT_MAX_RETRIES
    return max(0, retries) + 1


def _backoff_delay(attempt: int) -> float:
    return min(BACKOFF_BASE_SECONDS * (2 ** max(0, attempt)), BACKOFF_MAX_SECONDS)


def _sleep(seconds: float) -> None:
    """退避等待；单元测试会 patch 掉它，避免真实 sleep。"""
    time.sleep(seconds)


# ---------------------------------------------------------------------------
# 默认 fetcher（标准库 urllib）
# ---------------------------------------------------------------------------
def default_fetcher(
    url: str, headers: dict[str, str], payload: dict[str, Any], timeout: float
) -> str:
    """默认 HTTP fetcher：urllib POST JSON，返回解码后的响应原文。"""
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url, data=data, headers=dict(headers), method="POST"
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = response.read()
            charset = response.headers.get_content_charset() or "utf-8"
    except urllib.error.HTTPError as exc:
        http_error = HTTPStatusError(
            exc.code, _read_http_error_body(exc), reason=str(exc.reason or "")
        )
        raise http_error from None
    return body.decode(charset, errors="replace")


def _read_http_error_body(exc: urllib.error.HTTPError) -> str:
    payload = b""
    try:
        payload = exc.read()
    except Exception:  # noqa: BLE001 - 读错误体失败不应掩盖原始 HTTP 错误
        payload = b""
    try:
        exc.close()
    except Exception:  # noqa: BLE001 - close 失败可忽略
        pass
    if not payload:
        return ""
    return payload.decode("utf-8", errors="replace")


# ---------------------------------------------------------------------------
# 环境变量 / 打码 / 调试
# ---------------------------------------------------------------------------
def _first_env(*names: str) -> str:
    for name in names:
        value = (os.environ.get(name) or "").strip()
        if value:
            return value
    return ""


def _env_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None or not raw.strip():
        return default
    try:
        return float(raw)
    except ValueError:
        raise LLMConfigError(f"环境变量 {name} 不是合法数字：{raw!r}")


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None or not raw.strip():
        return default
    try:
        return int(raw)
    except ValueError:
        raise LLMConfigError(f"环境变量 {name} 不是合法整数：{raw!r}")


def _env_bool(name: str, default: bool) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return default
    if raw in _TRUE_VALUES:
        return True
    if raw in _FALSE_VALUES:
        return False
    raise LLMConfigError(f"环境变量 {name} 不是合法布尔值（可用 1/0/true/false）：{raw!r}")


def _redact(text: str, api_key: str) -> str:
    """把文本里可能出现的完整 key 替换为打码形式（最后一道安全网）。"""
    if not text or not api_key:
        return text
    try:
        return text.replace(api_key, mask_key(api_key))
    except Exception:  # pragma: no cover - 极端类型兜底
        return text


def _preview(text: Any, limit: int = _ERROR_PREVIEW_LIMIT) -> str:
    if not text:
        return ""
    collapsed = " ".join(str(text).split())
    if len(collapsed) <= limit:
        return collapsed
    return collapsed[:limit] + "…"


def _error_text(error: Any) -> str:
    if isinstance(error, dict):
        message = error.get("message") or error.get("code") or ""
        return str(message) if message else _compact_json(error)
    return str(error)


def _debug_enabled() -> bool:
    return (os.environ.get("TAS_LLM_DEBUG") or "").strip().lower() in _DEBUG_VALUES


def _debug(message: str) -> None:
    """打印请求摘要到 stderr（仅 TAS_LLM_DEBUG=1；不含 key / Authorization 头）。"""
    if not _debug_enabled():
        return
    try:
        print(f"[tas.llm] {message}", file=sys.stderr)
    except Exception:  # pragma: no cover - stderr 不可用时不阻塞调用
        pass


__all__ = [
    "LLMConfig",
    "LLMError",
    "LLMConfigError",
    "HTTPStatusError",
    "load_config",
    "chat_json",
    "mask_key",
    "default_fetcher",
]
