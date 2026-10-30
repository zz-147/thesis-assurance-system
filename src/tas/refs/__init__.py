"""参考文献相关能力（教师端「文献真实性核验」）。

三个互相独立、只读（绝不修改文档）的模块：

- :mod:`tas.refs.parse`       GB/T 7714 参考文献条目解析 → :class:`ReferenceEntry`
- :mod:`tas.refs.citations`   正文引用与文献表的一致性（离线）→ :class:`CitationFinding`
- :mod:`tas.refs.verify`      Crossref / OpenAlex / arXiv 真实性核验（可注入 fetcher，默认不联网）

设计原则（与项目约束一致）：

1. 只判断、不修改：本包没有任何写回 docx 的代码。
2. 可解释、可复现：所有结论都带定位/证据 URL；联网失败绝不伪装成「查无此文」。
3. 诚实边界：``not_found`` 仅表示「在已查询索引中未检索到」，不是「造假」判定；
   中文期刊、专著、内部资料在公开索引中覆盖有限，必须人工核实。
"""

from __future__ import annotations

from .citations import CitationFinding, analyze_citations
from .parse import ReferenceEntry, parse_reference, parse_references
from .verify import (
    DEFAULT_MAILTO,
    LIKELY_THRESHOLD,
    VERIFIED_THRESHOLD,
    VerificationResult,
    default_fetcher,
    title_similarity,
    verify_reference,
    verify_references,
)

__all__ = [
    # parse
    "ReferenceEntry",
    "parse_reference",
    "parse_references",
    # citations
    "CitationFinding",
    "analyze_citations",
    # verify
    "VerificationResult",
    "verify_reference",
    "verify_references",
    "title_similarity",
    "default_fetcher",
    "VERIFIED_THRESHOLD",
    "LIKELY_THRESHOLD",
    "DEFAULT_MAILTO",
]
