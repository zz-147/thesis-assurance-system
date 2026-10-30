"""参考文献真实性核验（Crossref → OpenAlex → arXiv，可注入 fetcher）。

**这是一条「检索辅助」而不是「打假」流程。** 诚实性约定（代码与报告文案都遵守）：

1. 只报告一条文献「在已查询的公开索引中是否检索到题名相近的记录」；
2. ``not_found`` 仅表示「未检索到」，**不等于文献不存在/造假**——中文期刊、专著、
   内部资料、会议论文在 Crossref/OpenAlex/arXiv 覆盖有限，必须人工核实；
3. 网络/接口失败一律 ``status="error"``，绝不把「请求失败」伪装成「查无此文」；
4. 所有结论都附证据 URL（查询 API URL，命中时另附 DOI 落地页/arXiv 摘要页）；
   任一步网络失败都会记录——全部失败、或失败后没有命中时返回 ``error``，
   只有「所有数据源都正常响应且都没检索到」才返回 ``not_found``。

可测试性：``fetcher(url: str, timeout: float) -> str`` 可注入（返回响应文本），
单元测试用假 fetcher 覆盖 verified / likely / not_found / error 四种状态，默认不联网。

数据源与接口（与 task-2 约定一致）：

- Crossref:  ``https://api.crossref.org/works?query.bibliographic=…&rows=3&select=…&mailto=…``
- OpenAlex:  ``https://api.openalex.org/works?search=…&per-page=3&mailto=…``
- arXiv:     ``http://export.arxiv.org/api/query?search_query=…&start=0&max_results=3``

判定阈值：标题相似度 + 年份/作者加分后，``score >= 0.85`` 为 verified，
``0.60 <= score < 0.85`` 为 likely，其余为 not_found。

防误报：过短/过泛的题名（如「物理学」「Optuna」）区分度低，即使题名完全一致也
只按 ``likely`` 提示；年份明显不一致时会在 note 中显式提醒——教师端永远只把
结果当线索，不据此判定文献真伪。
"""

from __future__ import annotations

import difflib
import json
import re
import time
import unicodedata
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Callable, Optional

from .parse import ReferenceEntry

VERIFIED_THRESHOLD = 0.85
LIKELY_THRESHOLD = 0.60

DEFAULT_MAILTO = "tas-dev@example.com"
USER_AGENT = "tas-refs/0.1 (reference verification; mailto:tas-dev@example.com)"

CROSSREF_ENDPOINT = "https://api.crossref.org/works"
OPENALEX_ENDPOINT = "https://api.openalex.org/works"
ARXIV_ENDPOINT = "http://export.arxiv.org/api/query"

_SOURCE_LABELS = {"crossref": "Crossref", "openalex": "OpenAlex", "arxiv": "arXiv"}
_SOURCE_ORDER = ("crossref", "openalex", "arxiv")

_ATOM_NS = "{http://www.w3.org/2005/Atom}"
_ARXIV_NS = "{http://arxiv.org/schemas/atom}"

# fetcher 约定：输入 URL 与超时（秒），返回响应文本；任何失败抛异常。
Fetcher = Callable[[str, float], str]

_STOP_AUTHOR_TOKENS = {"et", "al", "and", "等", "主编", "译", "著", "编"}

# 过短/过泛的题名区分度低（如「物理学」「Optuna」），不直接给 verified
_WEAK_TITLE_MAX_CHARS = 6
_WEAK_TITLE_SCORE_CAP = LIKELY_THRESHOLD + 0.20  # 0.80 < 0.85，即封顶为 likely


@dataclass
class _Candidate:
    """某个数据源返回的一条候选记录（内部使用）。"""

    title: str
    year: Optional[int] = None
    venue: str = ""
    doi: str = ""
    url: str = ""
    authors: tuple[str, ...] = ()


@dataclass
class VerificationResult:
    """一条参考文献的核验结果。"""

    entry: ReferenceEntry
    status: str  # verified | likely | not_found | error
    source: str  # crossref | openalex | arxiv | none
    matched_title: str = ""
    matched_year: Optional[int] = None
    matched_venue: str = ""
    doi: str = ""
    url: str = ""
    score: float = 0.0
    note: str = ""
    evidence: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------
def verify_reference(
    entry: ReferenceEntry,
    *,
    timeout: float = 10.0,
    mailto: str = DEFAULT_MAILTO,
    fetcher: Optional[Fetcher] = None,
) -> VerificationResult:
    """核验一条参考文献；按 Crossref → OpenAlex → arXiv 顺序检索。

    - 命中高相似度（score ≥ 0.85）即停止继续请求，省时且对公共 API 友好；
    - 任一数据源 HTTP/解析失败都会被捕获并记录，全部失败时返回 ``error``；
    - 不抛异常：调用方可以安全地批量跑。
    """
    fetcher = fetcher or default_fetcher
    query = _query_text(entry)
    if not query:
        return VerificationResult(
            entry=entry,
            status="error",
            source="none",
            note=(
                "条目缺少可检索的题名/正文内容（解析失败），无法联网核验，"
                "需人工核实该文献是否存在"
            ),
            evidence=[],
        )

    evidence: list[str] = []
    errors: list[str] = []
    successful_sources = 0
    best: Optional[tuple[float, float, _Candidate, str]] = None

    for source in _SOURCE_ORDER:
        url = _source_url(source, query, mailto)
        evidence.append(url)
        try:
            raw_text = fetcher(url, timeout)
            candidates = _PARSERS[source](raw_text)
            successful_sources += 1
        except Exception as exc:  # 网络错误、超时、JSON/XML 解析错误全部兜住
            errors.append(f"{_SOURCE_LABELS[source]}：{type(exc).__name__}: {exc}")
            continue

        for candidate in candidates:
            score, similarity = _score(entry, candidate)
            if best is None or score > best[0]:
                best = (score, similarity, candidate, source)
        if best is not None and best[0] >= VERIFIED_THRESHOLD:
            break  # 已有高置信命中，不再打扰其他数据源

    if successful_sources == 0:
        note = (
            "联网核验失败，未能从 Crossref/OpenAlex/arXiv 获取有效响应："
            + "；".join(errors)
            + "。这不代表文献不存在，需人工核实或稍后重试。"
        )
        return VerificationResult(
            entry=entry,
            status="error",
            source="none",
            note=note,
            evidence=_dedupe(evidence),
        )

    if best is None or best[0] < LIKELY_THRESHOLD:
        if errors:
            # 有数据源失败时不能声称「三个索引都没检索到」：只报告为 error，
            # 并说明已成功的数据源没有相近记录，交人工核实或稍后重试。
            note = (
                "未能完成完整核验：部分数据源请求失败（"
                + "；".join(errors)
                + "）。在成功响应的数据源中未检索到题名相近记录，"
                "需人工核实或稍后重试；这不代表文献不存在。"
            )
            return VerificationResult(
                entry=entry,
                status="error",
                source="none",
                score=round(best[0], 4) if best is not None else 0.0,
                note=note,
                evidence=_dedupe(evidence),
            )
        note = (
            "未在 Crossref/OpenAlex/arXiv 检索到，中文期刊可能未被收录，需人工核实。"
            "检索不到不代表文献造假：也可能是题录字段不全、题名写法不同，"
            "或属于未公开索引收录的专著/内部资料。"
        )
        return VerificationResult(
            entry=entry,
            status="not_found",
            source="none",
            score=round(best[0], 4) if best is not None else 0.0,
            note=note,
            evidence=_dedupe(evidence),
        )

    score, similarity, candidate, source = best
    status = "verified" if score >= VERIFIED_THRESHOLD else "likely"
    hit_url = _candidate_url(candidate)
    notes: list[str] = []
    if status == "verified":
        notes.append(
            f"在 {_SOURCE_LABELS[source]} 检索到题名高度相似的记录"
            f"（标题相似度 {similarity:.2f}，综合得分 {score:.2f}）；"
            "请人工确认作者、卷期页等著录细节。"
        )
    else:
        notes.append(
            f"在 {_SOURCE_LABELS[source]} 检索到相近记录"
            f"（标题相似度 {similarity:.2f}，综合得分 {score:.2f}），"
            "尚不足以确认是同一文献，需人工核实。"
        )
    if _is_weak_title(entry.title):
        notes.append("该题名过短/过泛、区分度低，已按线索提示处理，需人工核实。")
    if (
        entry.year is not None
        and candidate.year is not None
        and abs(entry.year - candidate.year) > 2
    ):
        notes.append(
            f"注意：命中记录年份为 {candidate.year}，与著录年份 {entry.year} 不一致，请核对是否为同一文献。"
        )
    if errors:
        notes.append("部分数据源请求失败：" + "；".join(errors))
    return VerificationResult(
        entry=entry,
        status=status,
        source=source,
        matched_title=candidate.title,
        matched_year=candidate.year,
        matched_venue=candidate.venue,
        doi=candidate.doi,
        url=hit_url,
        score=score,
        note="".join(notes),
        evidence=_dedupe(([hit_url] if hit_url else []) + evidence),
    )


def verify_references(
    entries: list[ReferenceEntry],
    *,
    limit: Optional[int] = None,
    pause: float = 0.2,
    timeout: float = 10.0,
    mailto: str = DEFAULT_MAILTO,
    fetcher: Optional[Fetcher] = None,
) -> list[VerificationResult]:
    """顺序核验多条文献（礼貌限速：每条之间 ``sleep(pause)`` 秒）。

    ``limit`` 用于先抽样核验；``fetcher``/``timeout``/``mailto`` 为可测试性扩展
    （task-2 基础签名之外新增的关键字参数，不影响原约定调用）。
    """
    selected = list(entries)
    if limit is not None:
        try:
            selected = selected[: max(0, int(limit))]
        except (TypeError, ValueError):
            pass
    results: list[VerificationResult] = []
    for position, entry in enumerate(selected):
        results.append(
            verify_reference(entry, timeout=timeout, mailto=mailto, fetcher=fetcher)
        )
        if pause and position < len(selected) - 1:
            try:
                time.sleep(max(0.0, float(pause)))
            except (TypeError, ValueError):
                pass
    return results


# ---------------------------------------------------------------------------
# 相似度
# ---------------------------------------------------------------------------
def title_similarity(left: str, right: str) -> float:
    """题名相似度 ∈ [0, 1]：归一化后取 SequenceMatcher 与词集合 Jaccard 的较大值。"""
    left_n = _normalise_title(left)
    right_n = _normalise_title(right)
    if not left_n or not right_n:
        return 0.0
    if left_n == right_n:
        return 1.0
    ratio = difflib.SequenceMatcher(None, left_n, right_n).ratio()
    left_tokens = left_n.split(" ")
    right_tokens = right_n.split(" ")
    jaccard = 0.0
    if len(left_tokens) > 1 and len(right_tokens) > 1:
        jaccard = _jaccard(set(left_tokens), set(right_tokens))
    return max(ratio, jaccard)


def _normalise_title(title: str) -> str:
    text = unicodedata.normalize("NFKC", title or "").lower()
    text = re.sub(r"[^0-9a-z\u4e00-\u9fff]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _jaccard(left: set[str], right: set[str]) -> float:
    if not left or not right:
        return 0.0
    return len(left & right) / len(left | right)


def _score(entry: ReferenceEntry, candidate: _Candidate) -> tuple[float, float]:
    """综合得分 = 题名相似度 + 年份加分（≤0.05）+ 作者重合加分（≤0.05），封顶 1。

    过短/过泛题名封顶为 ``likely``（0.80），避免「物理学」这类题名精确匹配到
    同名的其他文献时被误报为 verified。
    """
    similarity = title_similarity(entry.title, candidate.title)
    bonus = 0.0
    if entry.year is not None and candidate.year is not None:
        difference = abs(entry.year - candidate.year)
        if difference == 0:
            bonus += 0.05
        elif difference <= 1:
            bonus += 0.02
    if _authors_overlap(entry.authors, candidate.authors):
        bonus += 0.05
    score = min(1.0, similarity + bonus)
    if _is_weak_title(entry.title):
        score = min(score, _WEAK_TITLE_SCORE_CAP)
    return round(score, 4), similarity


def _is_weak_title(title: str) -> bool:
    """题名是否过短/过泛，不具备单独指认文献的区分度。"""
    normalised = _normalise_title(title)
    if not normalised:
        return True
    if len(normalised) <= _WEAK_TITLE_MAX_CHARS:
        return True
    tokens = normalised.split(" ")
    return len(tokens) == 1 and len(tokens[0]) <= 12


def _authors_overlap(entry_authors: str, candidate_authors: tuple[str, ...]) -> bool:
    entry_tokens = _author_tokens(entry_authors)
    if not entry_tokens or not candidate_authors:
        return False
    candidate_tokens: set[str] = set()
    for name in candidate_authors:
        candidate_tokens |= _author_tokens(name)
    return bool(entry_tokens & candidate_tokens)


def _author_tokens(authors: str) -> set[str]:
    text = unicodedata.normalize("NFKC", authors or "").lower()
    tokens = {
        token.strip(" .")
        for token in re.split(r"[，,;；、.．\s]+", text)
        if token.strip(" .")
    }
    return {token for token in tokens if len(token) >= 2 and token not in _STOP_AUTHOR_TOKENS}


# ---------------------------------------------------------------------------
# URL 构造
# ---------------------------------------------------------------------------
def _query_text(entry: ReferenceEntry) -> str:
    title = (entry.title or "").strip()
    if title:
        return title
    raw = (entry.raw or "").strip()
    return re.sub(r"^\s*[\[［]\d+[\]］]\s*", "", raw).strip()


def _source_url(source: str, query: str, mailto: str) -> str:
    if source == "crossref":
        params = urllib.parse.urlencode(
            {
                "query.bibliographic": query,
                "rows": "3",
                "select": "DOI,title,author,issued,container-title",
                "mailto": mailto,
            }
        )
        return f"{CROSSREF_ENDPOINT}?{params}"
    if source == "openalex":
        params = urllib.parse.urlencode(
            {"search": query, "per-page": "3", "mailto": mailto}
        )
        return f"{OPENALEX_ENDPOINT}?{params}"
    if source == "arxiv":
        # 引号内是精确短语检索；去掉查询里可能破坏语法的引号
        search_query = 'all:"' + query.replace('"', " ").strip() + '"'
        params = urllib.parse.urlencode(
            {"search_query": search_query, "start": "0", "max_results": "3"}
        )
        return f"{ARXIV_ENDPOINT}?{params}"
    raise ValueError(f"未知数据源：{source}")


# ---------------------------------------------------------------------------
# 默认 fetcher（标准库 urllib）
# ---------------------------------------------------------------------------
def default_fetcher(url: str, timeout: float = 10.0) -> str:
    """默认 HTTP fetcher：urllib + User-Agent；返回解码后的响应文本。"""
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "application/json, application/atom+xml;q=0.9, */*;q=0.8",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


# ---------------------------------------------------------------------------
# 各数据源响应解析
# ---------------------------------------------------------------------------
def _parse_crossref(raw_text: str) -> list[_Candidate]:
    data = json.loads(raw_text)
    items = ((data.get("message") or {}).get("items")) or []
    candidates: list[_Candidate] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        titles = item.get("title") or []
        title = titles[0] if isinstance(titles, list) and titles else ""
        if not isinstance(title, str) or not title.strip():
            continue
        year = _crossref_year(item)
        containers = item.get("container-title") or []
        venue = containers[0] if isinstance(containers, list) and containers else ""
        doi = (item.get("DOI") or "").strip()
        authors = tuple(_crossref_author_name(author) for author in (item.get("author") or []))
        authors = tuple(name for name in authors if name)
        url = f"https://doi.org/{doi}" if doi else (item.get("URL") or "")
        candidates.append(
            _Candidate(
                title=title.strip(),
                year=year,
                venue=venue if isinstance(venue, str) else "",
                doi=doi,
                url=url,
                authors=authors,
            )
        )
    return candidates


def _crossref_year(item: dict) -> Optional[int]:
    issued = item.get("issued") or {}
    date_parts = issued.get("date-parts") or [] if isinstance(issued, dict) else []
    if not date_parts or not isinstance(date_parts, list):
        return None
    first = date_parts[0]
    if not first:
        return None
    try:
        return int(first[0])
    except (TypeError, ValueError, IndexError):
        return None


def _crossref_author_name(author: object) -> str:
    if not isinstance(author, dict):
        return ""
    parts = [author.get("family") or "", author.get("given") or ""]
    name = " ".join(part for part in parts if part).strip()
    return name or str(author.get("name") or "").strip()


def _parse_openalex(raw_text: str) -> list[_Candidate]:
    data = json.loads(raw_text)
    results = data.get("results") or []
    candidates: list[_Candidate] = []
    for work in results:
        if not isinstance(work, dict):
            continue
        title = (work.get("display_name") or "").strip()
        if not title:
            continue
        doi_url = (work.get("doi") or "").strip()
        doi = ""
        if doi_url:
            doi = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", doi_url, flags=re.IGNORECASE)
        venue = ""
        primary_location = work.get("primary_location") or {}
        if isinstance(primary_location, dict):
            source = primary_location.get("source") or {}
            if isinstance(source, dict):
                venue = (source.get("display_name") or "").strip()
        authors: list[str] = []
        for authorship in work.get("authorships") or []:
            if not isinstance(authorship, dict):
                continue
            author = authorship.get("author") or {}
            if isinstance(author, dict):
                name = (author.get("display_name") or "").strip()
                if name:
                    authors.append(name)
        year = work.get("publication_year")
        try:
            year = int(year) if year is not None else None
        except (TypeError, ValueError):
            year = None
        url = doi_url or (work.get("id") or "")
        candidates.append(
            _Candidate(
                title=title,
                year=year,
                venue=venue,
                doi=doi,
                url=url,
                authors=tuple(authors),
            )
        )
    return candidates


def _parse_arxiv(raw_text: str) -> list[_Candidate]:
    root = ET.fromstring(raw_text)
    candidates: list[_Candidate] = []
    for entry in root.findall(f"{_ATOM_NS}entry"):
        title_element = entry.find(f"{_ATOM_NS}title")
        title = " ".join((title_element.text or "").split()) if title_element is not None else ""
        if not title:
            continue
        published = entry.findtext(f"{_ATOM_NS}published") or ""
        year: Optional[int] = None
        if len(published) >= 4 and published[:4].isdigit():
            year = int(published[:4])
        doi_element = entry.find(f"{_ARXIV_NS}doi")
        doi = (doi_element.text or "").strip() if doi_element is not None else ""
        identifier = entry.findtext(f"{_ATOM_NS}id") or ""
        url = identifier.strip()
        if not doi and "arxiv.org" not in url:
            url = ""
        candidates.append(
            _Candidate(
                title=title,
                year=year,
                venue="arXiv",
                doi=doi,
                url=url,
                authors=(),
            )
        )
    return candidates


_PARSERS: dict[str, Callable[[str], list[_Candidate]]] = {
    "crossref": _parse_crossref,
    "openalex": _parse_openalex,
    "arxiv": _parse_arxiv,
}


# ---------------------------------------------------------------------------
# 杂项
# ---------------------------------------------------------------------------
def _candidate_url(candidate: _Candidate) -> str:
    if candidate.url:
        return candidate.url
    if candidate.doi:
        return f"https://doi.org/{candidate.doi}"
    return ""


def _dedupe(items: list[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            output.append(item)
    return output
