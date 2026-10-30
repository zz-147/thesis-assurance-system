"""配置与规则集加载。

规则数据保存在项目根目录的 rules/rules.json（与代码分离，便于老师审阅改规则）。
"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

DEFAULT_PROFILE = "sci_eng"

PROFILES: dict[str, str] = {
    "sci_eng": "理工农（章：第N章；节：N.N；小节：N.N.N；图表按章编号）",
    "humanities": "文经管艺教（章：一、；节：（一）；小节：1.；图表按顺序编号）",
}

# 检查容差（Word 内部单位换算会有舍入，不能要求逐 twip 相等）
TOLERANCE = {
    "cm": 0.03,        # 页边距/纸张（cm）
    "pt": 0.6,         # 字号/段距（pt）
    "spacing": 0.02,   # 行距（倍数）
    "chars": 6.0,      # 缩进（1/100 字符 → 6 = 0.06 字符）
    "lines": 0.25,     # 段前/段后（行）
}


def project_root() -> Path:
    """返回项目根目录。

    优先级：``TAS_PROJECT_ROOT`` 环境变量 → PyInstaller 打包目录 → src 布局 / 安装布局。
    打包成 exe 时把 ``rules/``、``profiles/`` 放在 exe 同目录，启动器设置
    ``TAS_PROJECT_ROOT`` 指向该目录即可（数据 ``data/`` 也落在那里）。
    """
    override = os.environ.get("TAS_PROJECT_ROOT")
    if override:
        return Path(override).expanduser().resolve()
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        for candidate in (exe_dir, exe_dir.parent):
            if (candidate / "rules" / "rules.json").is_file():
                return candidate
        return exe_dir
    here = Path(__file__).resolve()
    for candidate in (here.parents[2], here.parents[1], here.parent):
        if (candidate / "rules" / "rules.json").is_file():
            return candidate
    return here.parents[2]


def default_rules_path() -> Path:
    return project_root() / "rules" / "rules.json"


@dataclass(frozen=True)
class Rule:
    """一条格式规则（rules.json 中的一项）。"""

    id: str
    name: str
    category: str
    phase: str
    severity: str
    profiles: tuple[str, ...]
    spec: str
    check: str
    fix: str
    source: str
    expected: dict

    def applies_to(self, profile: str) -> bool:
        return "*" in self.profiles or profile in self.profiles


class RuleSet:
    """规则库：按 id 索引，可按 profile / phase 过滤。"""

    def __init__(self, rules: dict[str, Rule], meta: dict):
        self.rules = rules
        self.meta = meta

    @classmethod
    def load(cls, path: Optional[Path] = None) -> "RuleSet":
        path = path or default_rules_path()
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        rules = {}
        for item in data.get("rules", []):
            rule = Rule(
                id=item["id"],
                name=item.get("name", ""),
                category=item.get("category", ""),
                phase=item.get("phase", "P1"),
                severity=item.get("severity", "error"),
                profiles=tuple(item.get("profiles", ["*"])),
                spec=item.get("spec", ""),
                check=item.get("check", ""),
                fix=item.get("fix", ""),
                source=item.get("source", ""),
                expected=item.get("expected", {}) or {},
            )
            rules[rule.id] = rule
        return cls(rules, data.get("meta", {}))

    def __len__(self) -> int:
        return len(self.rules)

    def get(self, rule_id: str) -> Rule:
        try:
            return self.rules[rule_id]
        except KeyError as exc:  # pragma: no cover - 规则库自检用
            raise KeyError(f"规则库中不存在规则 {rule_id}") from exc

    def implemented(self, ids: list[str]) -> list[Rule]:
        return [self.rules[i] for i in ids if i in self.rules]


# ---------------------------------------------------------------------------
# 命名规则集（profile）：一套规则文件 + 一个学科体系
# ---------------------------------------------------------------------------
DEFAULT_PROFILE_ID = "school_default"


@dataclass(frozen=True)
class Profile:
    """一个命名规范集（学科/课程级）。"""

    id: str
    name: str
    discipline: str
    rules_path: Path
    source: str = ""
    notes: str = ""


def profiles_index_path() -> Path:
    return project_root() / "profiles" / "index.json"


def list_profiles() -> list[Profile]:
    """读取 profiles/index.json；缺失时返回空列表（向后兼容单规则库模式）。"""
    index_path = profiles_index_path()
    if not index_path.is_file():
        return []
    data = json.loads(index_path.read_text(encoding="utf-8"))
    profiles: list[Profile] = []
    for item in data.get("profiles", []):
        rules_path = (index_path.parent / item.get("rules", "../rules/rules.json")).resolve()
        profiles.append(
            Profile(
                id=item["id"],
                name=item.get("name", item["id"]),
                discipline=item.get("discipline", DEFAULT_PROFILE),
                rules_path=rules_path,
                source=item.get("source", ""),
                notes=item.get("notes", ""),
            )
        )
    return profiles


def resolve_profile(identifier: Optional[str] = None) -> tuple[Profile, RuleSet]:
    """按 profile id 或规则文件路径加载（profile, 规则集）。

    - identifier 为 .json 路径：直接作为规则库，学科体系取默认；
    - identifier 为 profile id：从 profiles/index.json 查；
    - identifier 为空：默认 school_default，找不到索引则退回 rules/rules.json。
    """
    if identifier and ("/" in identifier or "\\" in identifier or identifier.lower().endswith(".json")):
        path = Path(identifier)
        if not path.is_file():
            raise FileNotFoundError(f"规则文件不存在：{path}")
        profile = Profile(
            id=path.stem,
            name=path.stem,
            discipline=DEFAULT_PROFILE,
            rules_path=path.resolve(),
        )
        return profile, RuleSet.load(path)

    profiles = list_profiles()
    if not profiles:
        profile = Profile(
            id=DEFAULT_PROFILE_ID,
            name="默认规则（rules/rules.json）",
            discipline=DEFAULT_PROFILE,
            rules_path=default_rules_path(),
        )
        return profile, RuleSet.load()

    target_id = identifier or DEFAULT_PROFILE_ID
    for profile in profiles:
        if profile.id == target_id:
            return profile, RuleSet.load(profile.rules_path)
    available = "、".join(item.id for item in profiles)
    raise ValueError(f"未找到 profile：{target_id}；可用：{available}")
