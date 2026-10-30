"""GB/T 7714 参考文献条目解析（只读，不修改文档）。

目标：把「参考文献」区的段落文本解析成结构化 :class:`ReferenceEntry`，供
教师端展示字段、做引用一致性检查，以及交给 :mod:`tas.refs.verify` 联网核验。

解析策略：正则为主、容错优先。
- 支持常见文献类型标识：J / M / C / D / P / S / N / R / G / A / Z 及
  DB、CP、EB（可带载体 OL/CD/MT/DK，如 ``[EB/OL]``）；
- 支持全角方括号/全角逗号/全角句点（学院规范示例即全角排版）；
- 抽取 [序号]、作者、题名、刊名/出版地：出版社、年份、DOI、URL；
- **任何解析失败都不抛异常**：保留 ``raw``，``title`` 尽量非空，不确定的字段留空。

已知局限（有意为之，宁缺毋滥）：
- 只处理单个段落内的条目；条目被 Word 拆成多段续行时不自动合并；
- 题名/作者的分界依赖「作者．题名」中点号约定；
- 中文期刊刊名、卷期页不做进一步规范化，交由 LLM/人工复核。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

from ..docx_io import DocModel
from ..roles import Roles

# 段落开头的 [序号]（半角/全角方括号 + 数字）
_NUM_PREFIX_RE = re.compile(r"^\s*[\[［]\s*(\d+)\s*[\]］]")

# 文献类型标识 [J] / [M] / [C] / [D] / [P] / [S] / [N] / [DB] / [CP] / [EB] / [EB/OL] …
# 注意多字母类型要放在单字母之前（正则交替按顺序匹配）。
_TYPE_RE = re.compile(
    r"[\[［]\s*(?P<kind>DB|CP|EB|CM|DS|J|M|C|D|P|S|N|R|G|A|Z)"
    r"(?:\s*[/／]\s*(?P<carrier>OL|CD|MT|DK))?\s*[\]］]",
    re.IGNORECASE,
)

# 年份：19xx / 20xx（避免从 ISBN/专利号中截出数字）
_YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(?!\d)")

# DOI：10.<4~9 位注册机构码>/<后缀>
_DOI_RE = re.compile(r"(?<!\d)(10\.\d{4,9}/[^\s，。；、\]】）]+)", re.IGNORECASE)

# URL：http(s)://…（兼容学院规范中可能出现的全角冒号/斜杠）
_URL_RE = re.compile(r"(?:https?)[:：]//[^\s，。；、\]】）]+", re.IGNORECASE)

# 作者与题名之间的点号（半角/全角）
_AUTHOR_END_RE = re.compile(r"[.．]")

_LEAD_SEP = " \t\u3000"
_TRAIL_SEP = " \t\u3000.．,，;；:：、。"
# URL 结尾常见的正文标点（不裁掉右括号，避免破坏合法 DOI/URL）
_URL_TRAIL = " \t\u3000.．,，;；:：、。"
_FULLWIDTH_URL_MAP = str.maketrans({"：": ":", "．": ".", "／": "/", "－": "-", "　": " "})
# 全角拉丁字母 → 半角（一对一映射，长度不变，匹配偏移可直接映射回原文），
# 用于兼容 "[Ｍ]" 这类全角文献类型标识。
_FULLWIDTH_LETTER_MAP = str.maketrans(
    {chr(0xFF21 + offset): chr(0x41 + offset) for offset in range(26)}
    | {chr(0xFF41 + offset): chr(0x61 + offset) for offset in range(26)}
)


@dataclass
class ReferenceEntry:
    """一条参考文献的结构化著录信息（解析失败也保证 raw 存在）。"""

    paragraph_index: int
    number: Optional[int]
    raw: str
    authors: str = ""
    title: str = ""
    venue: str = ""
    year: Optional[int] = None
    doc_type: Optional[str] = None
    doi: Optional[str] = None
    url: Optional[str] = None


# ---------------------------------------------------------------------------
# 单条解析
# ---------------------------------------------------------------------------
def parse_reference(paragraph_index: int, raw: str) -> ReferenceEntry:
    """解析一条参考文献文本；**绝不抛异常**。"""
    raw_text = raw or ""
    entry = ReferenceEntry(paragraph_index=int(paragraph_index), number=None, raw=raw_text)
    text = raw_text.strip()
    if not text:
        return entry

    body = text
    number_match = _NUM_PREFIX_RE.match(text)
    if number_match:
        try:
            entry.number = int(number_match.group(1))
        except (TypeError, ValueError):  # 理论上不会发生，保险
            entry.number = None
        body = text[number_match.end():].lstrip(_LEAD_SEP)

    # DOI / URL 从条目主体中抽取（与是否识别出类型无关）
    doi_match = _DOI_RE.search(body)
    if doi_match:
        entry.doi = _strip_tail(doi_match.group(1)).rstrip("/") or None
    url_match = _URL_RE.search(body)
    if url_match:
        entry.url = _normalise_url(url_match.group(0)) or None

    # 全角字母（如 [Ｍ]）先一对一映射成半角再匹配，偏移与原文一致
    type_match = _TYPE_RE.search(body.translate(_FULLWIDTH_LETTER_MAP))
    before_type = body
    after_type = ""
    if type_match:
        kind = (type_match.group("kind") or "").upper()
        carrier = type_match.group("carrier")
        entry.doc_type = f"{kind}/{carrier.upper()}" if carrier else kind
        before_type = body[: type_match.start()]
        after_type = body[type_match.end():]

    authors, title = _split_authors_title(before_type)
    if type_match:
        entry.authors = authors
        entry.title = title or _fallback_title(text, body)
        entry.year = _extract_year(after_type)
        if entry.year is None:
            entry.year = _extract_year(body)
        entry.venue = _extract_venue(after_type)
    else:
        # 没有类型标识：退化为「作者. 题名. 刊名, 年」的尽力解析
        entry.authors = authors
        entry.title, entry.venue = _split_title_venue(title or _strip_tail(body))
        entry.year = _extract_year(body)
        if not entry.title:
            entry.title = _fallback_title(text, body)

    return entry


def parse_references(model: DocModel, roles: Roles) -> list[ReferenceEntry]:
    """从文档模型中解析参考文献表。

    主路径：读取角色为 ``reference_entry`` 的段落（由 :func:`tas.roles.classify` 判定）。
    兜底：若一个条目都没识别到（例如文档没有「参考文献」标题或角色识别失败），
    则扫描「参考文献」标题之后、以 ``[n]`` 开头的段落，避免整个功能静默失效。
    """
    entries: list[ReferenceEntry] = []
    for para in model.paragraphs:
        if para.is_empty:
            continue
        try:
            role = roles.get(para.index)
        except Exception:
            role = "body"
        if role == "reference_entry":
            entries.append(parse_reference(para.index, para.text))
    if entries:
        return entries

    # ---- 兜底：角色识别失败时按 [n] 序号找参考文献条目 ----
    try:
        refs_title = roles.first_index("refs_title")
    except Exception:
        refs_title = None
    for para in model.paragraphs:
        if para.is_empty:
            continue
        if refs_title is not None and para.index <= refs_title:
            continue
        text = para.text or ""
        if _NUM_PREFIX_RE.match(text):
            entries.append(parse_reference(para.index, text))
    return entries


# ---------------------------------------------------------------------------
# 字段抽取辅助
# ---------------------------------------------------------------------------
def _strip_tail(value: str) -> str:
    return (value or "").strip().strip(_TRAIL_SEP).strip()


def _split_authors_title(segment: str) -> tuple[str, str]:
    """把「作者．题名」切成两部分；没有点号则整段当作题名。"""
    segment = (segment or "").strip(_LEAD_SEP)
    match = _AUTHOR_END_RE.search(segment)
    if match is None:
        return "", _strip_tail(segment)
    authors = _strip_tail(segment[: match.start()])
    title = _strip_tail(segment[match.end():])
    if not title:  # 有点号但点号后为空，说明切割失败，退回整段
        return authors, _strip_tail(segment)
    return authors, title


def _split_title_venue(segment: str) -> tuple[str, str]:
    """无类型标识时的尽力拆分：「题名．刊名，年」或「题名，刊名」。"""
    segment = (segment or "").strip()
    if not segment:
        return "", ""
    parts = re.split(r"[.．]", segment, maxsplit=1)
    if len(parts) == 2 and parts[0].strip():
        title = _strip_tail(parts[0])
        rest = parts[1]
        venue = _strip_tail(re.split(r"[，,]", rest, maxsplit=1)[0])
        return title, venue
    parts = re.split(r"[，,]", segment, maxsplit=1)
    if len(parts) == 2 and parts[0].strip():
        return _strip_tail(parts[0]), _strip_tail(parts[1])
    return _strip_tail(segment), ""


def _fallback_title(text: str, body: str) -> str:
    """题名兜底：优先条目主体，其次原文（保证尽量非空）。"""
    title = _strip_tail(body)
    if title:
        return title
    return (text or "").strip()


def _extract_year(segment: str) -> Optional[int]:
    if not segment:
        return None
    match = _YEAR_RE.search(segment)
    if not match:
        return None
    try:
        return int(match.group(1))
    except (TypeError, ValueError):
        return None


def _extract_venue(segment: str) -> str:
    """从类型标识之后的文本中抽取刊名/出版地：出版社。

    典型输入（``|`` 处为类型标识结束）：:

        [J] |．清华大学学报，2017，27（1）：1－8．
        [M] |．北京：科学出版社，2015．
        [D] |．北京：北京大学数学系，2016．
        [C] |．见：郑福寿主编．学报论丛：第 2 集．南京：河海大学出版社，2017．

    抽取规则：先截到年份前；会议录/专著常有多个句点，取最后一个句点片段
    （通常是「出版地：出版社」）；再取第一个逗号前的部分作为 venue。
    """
    tail = (segment or "").lstrip(_LEAD_SEP + ".．")
    if not tail.strip():
        return ""
    year_match = _YEAR_RE.search(tail)
    pre = tail[: year_match.start()] if year_match else tail
    pre = pre.strip()
    if not pre:
        return ""
    segments = [part for part in re.split(r"[.．]", pre) if part.strip(_TRAIL_SEP)]
    if len(segments) > 1:
        pre = segments[-1]
    pre = re.split(r"[，,]", pre, maxsplit=1)[0]
    return _strip_tail(pre)


def _normalise_url(url: str) -> str:
    normalised = (url or "").translate(_FULLWIDTH_URL_MAP).strip()
    # 规范样例里 URL 可能被中文标点或 "。" 收尾
    return normalised.rstrip(_URL_TRAIL).rstrip(")]】").strip()
