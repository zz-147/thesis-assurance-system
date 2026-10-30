"""LLM 客户端包（OpenAI 兼容，默认 DeepSeek）。

对外只暴露稳定接口，主流程（规范文本 → 规则候选抽取）按此调用::

    from tas.llm import load_config, chat_json, LLMError, LLMConfigError

    config = load_config()                      # 读 TAS_LLM_API_KEY（DeepSeek key）
    data = chat_json(messages, config=config, schema_hint={...})  # -> dict

实现细节见 :mod:`tas.llm.client`：重试/退避、```json 围栏剥离、key 打码、
TAS_LLM_DEBUG 调试开关、可注入 fetcher（测试不联网、不需要真实 key）。
"""

from __future__ import annotations

from .client import (
    LLMConfig,
    LLMConfigError,
    LLMError,
    chat_json,
    load_config,
    mask_key,
)

__all__ = [
    "LLMConfig",
    "LLMError",
    "LLMConfigError",
    "load_config",
    "chat_json",
    "mask_key",
]
