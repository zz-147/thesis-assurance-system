"""规则参数读取：让 checker / fixer 真正消费 rules.json 的 expected 参数。

背景：检查器最初把学校规范的期望值（字号、页边距、行距……）写死在代码里，
导致教师生成的 profile 改了 expected 也不生效。这里统一从「当前规则集」读取参数：
优先用规则库里的 expected，缺字段时回退到调用方传入的默认值。

单位约定（与 rules.json 一致）：
- first_line_chars：1/100 字符（200 → 2.0 字符）
- before_lines / after_lines：1/100 行（100 → 1.0 行，50 → 0.5 行）
- 其他字段原样使用（size_pt / line_spacing / alignment / margin_*_cm ...）
"""

from __future__ import annotations

from typing import Any, Optional

from .config import RuleSet


def expected_params(ruleset: RuleSet, rule_id: str, fallback: dict[str, Any]) -> dict[str, Any]:
    """fallback 为基础，套用规则库中该规则的 expected（只覆盖非 None 字段）。"""
    params: dict[str, Any] = dict(fallback)
    rule = ruleset.rules.get(rule_id)
    if rule is not None and isinstance(rule.expected, dict):
        for key, value in rule.expected.items():
            if value is not None:
                params[key] = value
    return params


def _scaled(value: Any, fallback: float) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return fallback
    # rules.json 用 1/100 单位（>10），代码内部用字符/行（2.0 / 1.0）
    return number / 100.0 if number > 10 else number


def line_count(value: Any, fallback: float) -> float:
    """段前/段后行数：100 → 1.0，50 → 0.5。"""
    return _scaled(value, fallback)


def char_count(value: Any, fallback: float) -> float:
    """缩进字符数：200 → 2.0。"""
    return _scaled(value, fallback)


def font_size(params: dict[str, Any], key: str = "size_pt", fallback: float = 12.0) -> float:
    try:
        return float(params.get(key, fallback))
    except (TypeError, ValueError):
        return fallback


def alignment(params: dict[str, Any], key: str = "alignment", fallback: str = "left") -> str:
    value = params.get(key, fallback)
    return str(value).strip().lower() if value else fallback


def text_value(params: dict[str, Any], key: str, fallback: str = "") -> str:
    value = params.get(key, fallback)
    return str(value) if value is not None else fallback
