"""段落角色识别：把文档拆成前置区/正文/标题/题注/参考文献等角色。

规则优先、宁可保守：认不出的段落归入 front_other 或 body，不做猜测性修改。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from .docx_io import DocModel, ParaInfo

_WS_RE = re.compile(r"[\s\u3000\u00a0]+")


def norm(text: str) -> str:
    """去掉全部空白（含全角空格、制表符），用于标签/标题匹配。"""
    return _WS_RE.sub("", text or "")


HEADING_STYLE_RE = re.compile(r"^(?:heading|标题)\s*([1-9])$", re.IGNORECASE)
CAP_FIG_RE = re.compile(r"^图\s*\d+(?:\s*[-.．－]\s*\d+)?")
CAP_TAB_RE = re.compile(r"^表\s*\d+(?:\s*[-.．－]\s*\d+)?")
H1_CHAPTER_RE = re.compile(r"^第[一二三四五六七八九十百零〇0-9]+章(?![为是])")
H1_CN_RE = re.compile(r"^[一二三四五六七八九十]+、")
H2_CN_RE = re.compile(r"^（[一二三四五六七八九十]+）")
H2_RE = re.compile(r"^\d+\.\d+(?!\.\d)")
H3_RE = re.compile(r"^\d+\.\d+\.\d+(?!\.\d)")
H4_RE = re.compile(r"^\d+\.\d+\.\d+\.\d+")
REF_ENTRY_RE = re.compile(r"^\[\d+\]")
TOC_PAGE_RE = re.compile(r"\t\s*\d+\s*$")

SPECIAL_TITLES = {
    "前言": "special_title",
    "引言": "special_title",
    "结论": "special_title",
    "致谢": "thanks_title",
    "参考文献": "refs_title",
    "附录": "appendix_title",
}


@dataclass
class Roles:
    by_index: dict[int, str] = field(default_factory=dict)
    body_start: Optional[int] = None

    def get(self, index: int, default: str = "body") -> str:
        return self.by_index.get(index, default)

    def indices(self, *roles: str) -> list[int]:
        wanted = set(roles)
        return [index for index, role in sorted(self.by_index.items()) if role in wanted]

    def first_index(self, *roles: str) -> Optional[int]:
        found = self.indices(*roles)
        return found[0] if found else None


def _is_toc_entry(para: ParaInfo) -> bool:
    if (para.style_name or "").lower().startswith("toc"):
        return True
    return bool(TOC_PAGE_RE.search(para.text or ""))


def heading_level(para: ParaInfo) -> Optional[int]:
    """返回 1..4 的标题级别；不是标题返回 None。"""
    match = HEADING_STYLE_RE.match((para.style_name or "").strip())
    if match:
        return int(match.group(1))
    if para.outline_level is not None and para.outline_level < 9:
        return para.outline_level + 1
    text_n = norm(para.text)
    if not text_n:
        return None
    if H1_CHAPTER_RE.match(text_n) or H1_CN_RE.match(text_n):
        return 1
    if H2_CN_RE.match(text_n) or H2_RE.match(text_n):
        return 2
    if H3_RE.match(text_n):
        return 3
    if H4_RE.match(text_n):
        return 4
    return None


def _first(paras: list[ParaInfo], predicate) -> Optional[int]:
    for para in paras:
        if predicate(para):
            return para.index
    return None


def _last_before(paras: list[ParaInfo], limit: Optional[int], floor: int = -1) -> Optional[int]:
    found = None
    for para in paras:
        if limit is not None and para.index >= limit:
            break
        if para.index > floor and not para.is_empty:
            found = para.index
    return found


def classify(model: DocModel, profile: str = "sci_eng") -> Roles:
    roles = Roles()
    nonempty = [para for para in model.paragraphs if not para.is_empty]
    if not nonempty:
        return roles

    # ---- 1. 找正文起点：目录之后第一个一级标题/前言 --------------------------
    toc_title = _first(
        nonempty, lambda p: norm(p.text) in ("目录",) or (norm(p.text).startswith("目录") and len(norm(p.text)) <= 6)
    )
    body_start: Optional[int] = None
    for para in nonempty:
        if toc_title is not None and para.index <= toc_title:
            continue
        if _is_toc_entry(para):
            continue
        text_n = norm(para.text)
        if text_n in ("前言", "引言"):
            body_start = para.index
            break
        if heading_level(para) == 1:
            body_start = para.index
            break
    if body_start is None:
        body_start = nonempty[-1].index + 1
    roles.body_start = body_start

    # ---- 2. 前置区角色 ------------------------------------------------------
    front = [para for para in nonempty if para.index < body_start]
    for para in front:
        roles.by_index[para.index] = "front_other"

    abs_cn = _first(front, lambda p: norm(p.text) in ("摘要", "中文摘要"))
    abs_en = _first(front, lambda p: norm(p.text).lower() in ("abstract", "英文摘要"))
    kw_cn = _first(front, lambda p: norm(p.text).startswith(("关键词", "关键字")))
    kw_en = _first(front, lambda p: norm(p.text).lower().startswith("keywords"))
    toc = _first(front, lambda p: norm(p.text) in ("目录",) or (norm(p.text).startswith("目录") and len(norm(p.text)) <= 6))

    title_cn = _last_before(front, abs_cn) if abs_cn is not None else None
    title_en = _last_before(front, abs_en, floor=(kw_cn or -1)) if abs_en is not None else None

    if title_cn is not None:
        roles.by_index[title_cn] = "title_cn"
    if title_en is not None and title_en != title_cn:
        roles.by_index[title_en] = "title_en"
    if abs_cn is not None:
        roles.by_index[abs_cn] = "abstract_label_cn"
    if abs_en is not None:
        roles.by_index[abs_en] = "abstract_label_en"
    if kw_cn is not None:
        roles.by_index[kw_cn] = "keywords_cn"
    if kw_en is not None:
        roles.by_index[kw_en] = "keywords_en"
    if toc is not None:
        roles.by_index[toc] = "toc_title"

    def between(start: Optional[int], end: Optional[int], role: str) -> None:
        if start is None:
            return
        for para in front:
            if para.index <= start:
                continue
            if end is not None and para.index >= end:
                continue
            roles.by_index[para.index] = role

    between(abs_cn, kw_cn if kw_cn is not None else abs_en, "abstract_body_cn")
    between(abs_en, kw_en if kw_en is not None else toc, "abstract_body_en")

    for para in front:
        if roles.get(para.index) == "front_other" and _is_toc_entry(para):
            roles.by_index[para.index] = "toc_entry"
    if toc is not None:
        for para in front:
            if para.index > toc and roles.get(para.index) == "front_other":
                roles.by_index[para.index] = "toc_entry"

    cover_end = min(
        [value for value in (title_cn, abs_cn, toc, body_start) if value is not None] or [body_start]
    )
    for para in front:
        if para.index < cover_end and roles.get(para.index) == "front_other":
            roles.by_index[para.index] = "cover"

    # ---- 3. 正文区角色 ------------------------------------------------------
    refs_mode = False
    for para in nonempty:
        if para.index < body_start:
            continue
        text_n = norm(para.text)
        if text_n == "参考文献":
            roles.by_index[para.index] = "refs_title"
            refs_mode = True
            continue
        if text_n == "致谢":
            roles.by_index[para.index] = "thanks_title"
            refs_mode = False
            continue
        if text_n.startswith("附录") and len(text_n) <= 8:
            roles.by_index[para.index] = "appendix_title"
            refs_mode = False
            continue
        if text_n in ("前言", "引言", "结论"):
            roles.by_index[para.index] = "special_title"
            continue
        if _is_toc_entry(para):
            roles.by_index[para.index] = "toc_entry"
            continue
        level = heading_level(para)
        if level is not None:
            roles.by_index[para.index] = f"heading_{level}"
            continue
        if CAP_FIG_RE.match(text_n):
            roles.by_index[para.index] = "caption_fig"
            continue
        if CAP_TAB_RE.match(text_n):
            roles.by_index[para.index] = "caption_tab"
            continue
        if refs_mode:
            roles.by_index[para.index] = "reference_entry"
            continue
        if para.has_drawing and not text_n:
            roles.by_index[para.index] = "figure"
            continue
        roles.by_index[para.index] = "body"

    return roles
