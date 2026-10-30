"""docx 读取层：把 python-docx 对象收敛成可检查的文档模型。

只读，不修改文档。所有“有效属性”都按 Word 的继承规则解析：
docDefaults < 样式链（基样式→当前样式） < 段落/run 上的直接格式。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator, Optional

import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

TWIPS_PER_CM = 566.9291338582677
TWIPS_PER_PT = 20.0
EMU_PER_CM = 360000.0

_QN_PG_SZ = qn("w:pgSz")
_QN_PG_MAR = qn("w:pgMar")
_QN_DOC_GRID = qn("w:docGrid")
_QN_PG_NUM_TYPE = qn("w:pgNumType")
_QN_PG_NUM_START = qn("w:start")
_QN_PPR = qn("w:pPr")
_QN_SECT_PR = qn("w:sectPr")
_QN_TITLE_PG = qn("w:titlePg")
_QN_P_BDR = qn("w:pBdr")
_QN_BOTTOM_BDR = qn("w:bottom")
_QN_TBL_PR = qn("w:tblPr")
_QN_TBL_BORDERS = qn("w:tblBorders")
_QN_TC_PR = qn("w:tcPr")
_QN_TC_BORDERS = qn("w:tcBorders")
_QN_INSTR_TEXT = qn("w:instrText")
_QN_FLD_SIMPLE = qn("w:fldSimple")
_QN_FLD_INSTR = qn("w:instr")
_QN_MATH = qn("m:oMath")
_QN_MATH_PARA = qn("m:oMathPara")
_QN_PAGE_BREAK_BEFORE = qn("w:pageBreakBefore")
_QN_BR = qn("w:br")
_QN_NUM_PR = qn("w:numPr")
_QN_VAL = qn("w:val")

# 表格边框名（与 OOXML w:tblBorders / w:tcBorders 子元素同名）
TABLE_BORDER_KEYS = ("top", "bottom", "left", "right", "insideH", "insideV")


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def twips_to_cm(value: Optional[str]) -> Optional[float]:
    try:
        return int(value) / TWIPS_PER_CM
    except (TypeError, ValueError):
        return None


def twips_to_pt(value: Optional[str]) -> Optional[float]:
    try:
        return int(value) / TWIPS_PER_PT
    except (TypeError, ValueError):
        return None


def _on_off(element: Any) -> Optional[bool]:
    """解析 w:b / w:i 这类开关元素：无 val 视为 True，val=0/false/off 为 False。"""
    if element is None:
        return None
    value = element.get(qn("w:val"))
    if value is None:
        return True
    return value.lower() not in ("0", "false", "off")


def _font_from_rpr(rpr: Any) -> dict:
    """从一个 w:rPr 元素读出字体相关属性。"""
    props: dict = {}
    if rpr is None:
        return props
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is not None:
        east = fonts.get(qn("w:eastAsia"))
        ascii_font = fonts.get(qn("w:ascii"))
        if east:
            props["east_asia"] = east
        if ascii_font:
            props["ascii"] = ascii_font
    size = rpr.find(qn("w:sz"))
    if size is not None:
        try:
            props["size_pt"] = int(size.get(qn("w:val"))) / 2
        except (TypeError, ValueError):
            pass
    for key, tag in (("bold", "w:b"), ("italic", "w:i")):
        element = rpr.find(qn(tag))
        if element is not None:
            props[key] = _on_off(element)
    return props


def _iter_style_chain(style: Any) -> Iterator[Any]:
    """沿基样式向上走（含自身），防止样式表成环。"""
    seen = set()
    current = style
    while current is not None:
        style_id = getattr(current, "style_id", None) or id(current)
        if style_id in seen:
            break
        seen.add(style_id)
        yield current
        try:
            current = current.base_style
        except Exception:
            current = None


def _doc_defaults_el(document: Any, kind: str) -> Any:
    """取 docDefaults 里的 rPr / pPr 元素。kind ∈ {'rPr','pPr'}。"""
    try:
        styles_el = document.styles.element
    except Exception:
        return None
    defaults = styles_el.find(qn("w:docDefaults"))
    if defaults is None:
        return None
    if kind == "rPr":
        container = defaults.find(qn("w:rPrDefault"))
        return container.find(qn("w:rPr")) if container is not None else None
    container = defaults.find(qn("w:pPrDefault"))
    return container.find(qn("w:pPr")) if container is not None else None


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------
@dataclass
class RunInfo:
    text: str
    east_asia: Optional[str] = None
    ascii: Optional[str] = None
    size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None


@dataclass
class ParaInfo:
    index: int
    text: str
    style_name: str = ""
    style_id: str = ""
    outline_level: Optional[int] = None
    alignment: Optional[str] = None
    line_spacing: Optional[float] = None       # 倍数（w:lineRule=auto）
    line_spacing_pt: Optional[float] = None    # 固定值/最小值（pt）
    first_line_chars: Optional[float] = None   # 1/100 字符单位 → 字符数
    first_line_twips: Optional[int] = None
    hanging_chars: Optional[float] = None
    before_lines: Optional[float] = None       # 行（1/100 行 → 行数）
    after_lines: Optional[float] = None
    before_pt: Optional[float] = None
    after_pt: Optional[float] = None
    runs: list[RunInfo] = field(default_factory=list)
    has_drawing: bool = False
    # ---- P1 增量字段（默认值保证向后兼容）--------------------------------
    has_math: bool = False            # 含 OMML 公式（m:oMath / m:oMathPara）
    page_break_before: bool = False   # w:pageBreakBefore（段落前分页）
    has_page_break: bool = False      # 段落内含 w:br w:type="page"
    has_num_pr: bool = False          # 段落使用 Word 自动编号（w:numPr）
    hanging_twips: Optional[int] = None        # w:ind/w:hanging（磅值）
    direct_before_lines: Optional[float] = None  # 段落自身 pPr 显式段前（行）
    direct_after_lines: Optional[float] = None   # 段落自身 pPr 显式段后（行）
    direct_before_pt: Optional[float] = None     # 段落自身 pPr 显式段前（pt）
    direct_after_pt: Optional[float] = None      # 段落自身 pPr 显式段后（pt）

    @property
    def is_empty(self) -> bool:
        return not (self.text or "").strip()


@dataclass
class SectionInfo:
    index: int
    page_width_cm: Optional[float] = None
    page_height_cm: Optional[float] = None
    margin_top_cm: Optional[float] = None
    margin_bottom_cm: Optional[float] = None
    margin_left_cm: Optional[float] = None
    margin_right_cm: Optional[float] = None
    gutter_cm: Optional[float] = None
    header_distance_cm: Optional[float] = None
    footer_distance_cm: Optional[float] = None
    header_text: str = ""
    footer_text: str = ""
    page_number_format: Optional[str] = None
    page_number_start: Optional[int] = None
    doc_grid_type: Optional[str] = None
    doc_grid_line_pitch: Optional[int] = None
    doc_grid_char_space: Optional[int] = None
    # ---- P1 增量字段（默认值保证向后兼容）--------------------------------
    first_para_index: Optional[int] = None   # 本节起始 body 段落 index（含）
    last_para_index: Optional[int] = None    # 本节结束 body 段落 index（含）
    header_linked: bool = False              # 页眉是否链接到上一节（继承）
    footer_linked: bool = False              # 页脚是否链接到上一节（继承）
    header_paras: list["HeaderFooterParaInfo"] = field(default_factory=list)
    footer_paras: list["HeaderFooterParaInfo"] = field(default_factory=list)
    has_page_field: bool = False             # 页脚是否含 PAGE 域（页码）
    page_field_instrs: list[str] = field(default_factory=list)
    has_title_pg: bool = False               # 首页不同（w:titlePg）


@dataclass
class BorderInfo:
    """一条边框（w:top / w:bottom / ...）：val + 粗细分（pt）。"""

    val: str = ""
    size_pt: float = 0.0
    color: str = ""


@dataclass
class HeaderFooterParaInfo:
    """页眉/页脚段落（只读快照：文字、字体、对齐、行距、下边框、域代码）。"""

    text: str = ""
    style_name: str = ""
    alignment: Optional[str] = None
    line_spacing: Optional[float] = None
    runs: list[RunInfo] = field(default_factory=list)
    bottom_border: Optional[BorderInfo] = None
    field_instrs: list[str] = field(default_factory=list)


@dataclass
class TableCellInfo:
    """表格单元格快照（用于三线表 / 表内文字检查）。"""

    text: str = ""
    runs: list[RunInfo] = field(default_factory=list)
    alignment: Optional[str] = None
    line_spacing: Optional[float] = None


@dataclass
class TableInfo:
    """表格快照：定位（前后段落）、边框、单元格（只读，供 P1 图表规则使用）。"""

    index: int
    prev_para_index: Optional[int] = None   # 表格前最近一个 body 段落 index
    next_para_index: Optional[int] = None   # 表格后第一个 body 段落 index
    row_count: int = 0
    column_count: int = 0
    style_name: str = ""
    borders: dict[str, BorderInfo] = field(default_factory=dict)   # top/bottom/left/right/insideH/insideV
    first_row_bottom_borders: list[Optional[BorderInfo]] = field(default_factory=list)
    cells: list[list[TableCellInfo]] = field(default_factory=list)


@dataclass
class DocModel:
    path: Path
    paragraphs: list[ParaInfo]
    sections: list[SectionInfo]
    tables_count: int
    image_count: int
    char_count: int
    by_index: dict[int, ParaInfo] = field(default_factory=dict, repr=False)
    # ---- P1 增量字段（默认值保证向后兼容）--------------------------------
    tables: list[TableInfo] = field(default_factory=list)
    has_toc_field: bool = False

    # ------------------------------------------------------------------
    @classmethod
    def load(cls, path: Path | str) -> "DocModel":
        path = Path(path)
        if path.suffix.lower() == ".doc":
            raise ValueError(
                "暂不支持 .doc（Word 97-2003）；请先用 Word/WPS 另存为 .docx，"
                "或等待 legacy 适配器接入。"
            )
        if not path.is_file():
            raise FileNotFoundError(f"文件不存在：{path}")
        document = docx.Document(str(path))

        paragraphs: list[ParaInfo] = []
        for index, para in enumerate(document.paragraphs):
            paragraphs.append(_build_para(index, para, document))
        by_index = {para.index: para for para in paragraphs}

        section_ranges = _section_ranges(document)
        sections: list[SectionInfo] = []
        for index, section in enumerate(document.sections, start=1):
            info = _build_section(index, section, document)
            if index - 1 < len(section_ranges):
                info.first_para_index, info.last_para_index = section_ranges[index - 1]
            sections.append(info)

        tables = _build_tables(document)

        body_xml = document.element.body
        image_count = len(body_xml.findall(".//" + qn("w:drawing"))) + len(
            body_xml.findall(".//" + qn("w:pict"))
        )
        char_count = sum(len(para.text or "") for para in paragraphs)

        return cls(
            path=path,
            paragraphs=paragraphs,
            sections=sections,
            tables_count=len(document.tables),
            image_count=image_count,
            char_count=char_count,
            by_index=by_index,
            tables=tables,
            has_toc_field=_has_toc_field(body_xml),
        )

    # ------------------------------------------------------------------
    def stats(self) -> dict:
        return {
            "paragraphs": len([p for p in self.paragraphs if not p.is_empty]),
            "tables": self.tables_count,
            "sections": len(self.sections),
            "images": self.image_count,
            "chars": self.char_count,
        }


# ---------------------------------------------------------------------------
# 构建 ParaInfo
# ---------------------------------------------------------------------------
def _effective_ppr_layers(para: Any, document: Any) -> list:
    layers: list = []
    defaults = _doc_defaults_el(document, "pPr")
    if defaults is not None:
        layers.append(defaults)
    para_style = getattr(para, "style", None)
    if para_style is not None:
        for style in reversed(list(_iter_style_chain(para_style))):
            ppr = style.element.find(qn("w:pPr"))
            if ppr is not None:
                layers.append(ppr)
    direct = para._p.pPr
    if direct is not None:
        layers.append(direct)
    return layers


def _effective_run_font(run: Any, para_style: Any, document: Any) -> dict:
    layers: list = []
    defaults = _doc_defaults_el(document, "rPr")
    if defaults is not None:
        layers.append(defaults)
    if para_style is not None:
        for style in reversed(list(_iter_style_chain(para_style))):
            rpr = style.element.find(qn("w:rPr"))
            if rpr is not None:
                layers.append(rpr)
    direct = run._r.rPr
    if direct is not None:
        layers.append(direct)
    merged: dict = {}
    for rpr in layers:
        merged.update(_font_from_rpr(rpr))
    return merged


def _effective_para_props(para: Any, document: Any) -> dict:
    props: dict = {}
    for ppr in _effective_ppr_layers(para, document):
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            first_chars = ind.get(qn("w:firstLineChars"))
            first = ind.get(qn("w:firstLine"))
            if first_chars is not None:
                try:
                    props["first_line_chars"] = int(first_chars) / 100.0
                except ValueError:
                    pass
                props["first_line_twips"] = None
            elif first is not None:
                try:
                    props["first_line_twips"] = int(first)
                except ValueError:
                    pass
                props["first_line_chars"] = None
            hanging_chars = ind.get(qn("w:hangingChars"))
            if hanging_chars is not None:
                try:
                    props["hanging_chars"] = int(hanging_chars) / 100.0
                except ValueError:
                    pass
            hanging_twips = ind.get(qn("w:hanging"))
            if hanging_twips is not None:
                try:
                    props["hanging_twips"] = int(hanging_twips)
                except ValueError:
                    pass
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None:
            line = spacing.get(qn("w:line"))
            rule = spacing.get(qn("w:lineRule")) or "auto"
            if line is not None:
                try:
                    value = int(line)
                except ValueError:
                    value = None
                if value is not None:
                    if rule == "auto":
                        props["line_spacing"] = value / 240.0
                        props["line_spacing_pt"] = None
                    else:
                        props["line_spacing_pt"] = value / 20.0
                        props["line_spacing"] = None
            before_lines = spacing.get(qn("w:beforeLines"))
            after_lines = spacing.get(qn("w:afterLines"))
            before = spacing.get(qn("w:before"))
            after = spacing.get(qn("w:after"))
            if before_lines is not None:
                try:
                    props["before_lines"] = int(before_lines) / 100.0
                    props["before_pt"] = None
                except ValueError:
                    pass
            elif before is not None:
                try:
                    props["before_pt"] = int(before) / 20.0
                    props["before_lines"] = None
                except ValueError:
                    pass
            if after_lines is not None:
                try:
                    props["after_lines"] = int(after_lines) / 100.0
                    props["after_pt"] = None
                except ValueError:
                    pass
            elif after is not None:
                try:
                    props["after_pt"] = int(after) / 20.0
                    props["after_lines"] = None
                except ValueError:
                    pass
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            props["alignment"] = jc.get(qn("w:val"))
        outline = ppr.find(qn("w:outlineLvl"))
        if outline is not None:
            try:
                props["outline_level"] = int(outline.get(qn("w:val")))
            except (TypeError, ValueError):
                pass
    return props


def _build_para(index: int, para: Any, document: Any) -> ParaInfo:
    style = getattr(para, "style", None)
    props = _effective_para_props(para, document)
    runs: list[RunInfo] = []
    for run in para.runs:
        font = _effective_run_font(run, style, document)
        runs.append(
            RunInfo(
                text=run.text or "",
                east_asia=font.get("east_asia"),
                ascii=font.get("ascii"),
                size_pt=font.get("size_pt"),
                bold=font.get("bold"),
                italic=font.get("italic"),
            )
        )
    has_drawing = bool(para._p.findall(".//" + qn("w:drawing"))) or bool(
        para._p.findall(".//" + qn("w:pict"))
    )
    ppr = para._p.pPr
    has_math = bool(para._p.findall(".//" + _QN_MATH)) or bool(
        para._p.findall(".//" + _QN_MATH_PARA)
    )
    page_break_before = False
    has_num_pr = False
    direct_before_lines = direct_after_lines = None
    direct_before_pt = direct_after_pt = None
    if ppr is not None:
        page_break_before = _on_off(ppr.find(_QN_PAGE_BREAK_BEFORE)) is True
        has_num_pr = ppr.find(_QN_NUM_PR) is not None
        direct_spacing = ppr.find(qn("w:spacing"))
        if direct_spacing is not None:
            before_lines_raw = direct_spacing.get(qn("w:beforeLines"))
            after_lines_raw = direct_spacing.get(qn("w:afterLines"))
            before_raw = direct_spacing.get(qn("w:before"))
            after_raw = direct_spacing.get(qn("w:after"))
            try:
                if before_lines_raw is not None:
                    direct_before_lines = int(before_lines_raw) / 100.0
            except ValueError:
                pass
            try:
                if after_lines_raw is not None:
                    direct_after_lines = int(after_lines_raw) / 100.0
            except ValueError:
                pass
            try:
                if before_raw is not None:
                    direct_before_pt = int(before_raw) / 20.0
            except ValueError:
                pass
            try:
                if after_raw is not None:
                    direct_after_pt = int(after_raw) / 20.0
            except ValueError:
                pass
    has_page_break = any(
        (br.get(qn("w:type")) or "") == "page" for br in para._p.iter(_QN_BR)
    )
    return ParaInfo(
        index=index,
        text=para.text or "",
        style_name=getattr(style, "name", "") or "",
        style_id=getattr(style, "style_id", "") or "",
        outline_level=props.get("outline_level"),
        alignment=props.get("alignment"),
        line_spacing=props.get("line_spacing"),
        line_spacing_pt=props.get("line_spacing_pt"),
        first_line_chars=props.get("first_line_chars"),
        first_line_twips=props.get("first_line_twips"),
        hanging_chars=props.get("hanging_chars"),
        before_lines=props.get("before_lines"),
        after_lines=props.get("after_lines"),
        before_pt=props.get("before_pt"),
        after_pt=props.get("after_pt"),
        runs=runs,
        has_drawing=has_drawing,
        has_math=has_math,
        page_break_before=page_break_before,
        has_page_break=has_page_break,
        has_num_pr=has_num_pr,
        hanging_twips=props.get("hanging_twips"),
        direct_before_lines=direct_before_lines,
        direct_after_lines=direct_after_lines,
        direct_before_pt=direct_before_pt,
        direct_after_pt=direct_after_pt,
    )


# ---------------------------------------------------------------------------
# 构建 SectionInfo
# ---------------------------------------------------------------------------
def _header_footer_text(kind: str, section: Any) -> str:
    try:
        container = getattr(section, kind)
        if container.is_linked_to_previous:
            # 链接到上一节时没有自己的内容；不要触发 python-docx 新建 part（只读）
            return ""
        parts = [p.text.strip() for p in container.paragraphs if (p.text or "").strip()]
        return " / ".join(parts)
    except Exception:
        return ""


def _is_page_field(instr: str) -> bool:
    """判断域代码是否是页码域 PAGE（排除 NUMPAGES / SECTIONPAGES）。"""
    upper = (instr or "").upper()
    return re.search(r"(?<![A-Z])PAGE(?![A-Z])", upper) is not None


def _border_from_el(element: Any) -> Optional[BorderInfo]:
    if element is None:
        return None
    raw_size = element.get(qn("w:sz"))
    try:
        size_pt = int(raw_size) / 8.0 if raw_size is not None else 0.0
    except (TypeError, ValueError):
        size_pt = 0.0
    return BorderInfo(
        val=(element.get(qn("w:val")) or "").strip(),
        size_pt=size_pt,
        color=element.get(qn("w:color")) or "",
    )


def _borders_from_container(container: Any, borders_tag: str) -> dict[str, BorderInfo]:
    """从 tblPr/tcPr 中读取 w:tblBorders / w:tcBorders 子边框。"""
    borders: dict[str, BorderInfo] = {}
    if container is None:
        return borders
    holder = container.find(borders_tag)
    if holder is None:
        return borders
    for child in holder:
        if not isinstance(child.tag, str):
            continue
        local = child.tag.split("}")[-1]
        info = _border_from_el(child)
        if info is not None:
            borders[local] = info
    return borders


def _style_table_borders(style: Any) -> dict[str, BorderInfo]:
    """兜底读取表格样式里的 tblBorders（显式边框缺失时用）。"""
    element = getattr(style, "element", None)
    if element is None:
        return {}
    table_pr = element.find(_QN_TBL_PR)
    return _borders_from_container(table_pr, _QN_TBL_BORDERS)


def _build_hf_runs(para: Any, document: Any) -> list[RunInfo]:
    style = getattr(para, "style", None)
    runs: list[RunInfo] = []
    for run in para.runs:
        font = _effective_run_font(run, style, document)
        runs.append(
            RunInfo(
                text=run.text or "",
                east_asia=font.get("east_asia"),
                ascii=font.get("ascii"),
                size_pt=font.get("size_pt"),
                bold=font.get("bold"),
                italic=font.get("italic"),
            )
        )
    return runs


def _build_hf_para(para: Any, document: Any) -> HeaderFooterParaInfo:
    props = _effective_para_props(para, document)
    ppr = para._p.pPr
    bottom_border = None
    if ppr is not None:
        p_bdr = ppr.find(_QN_P_BDR)
        if p_bdr is not None:
            bottom_border = _border_from_el(p_bdr.find(_QN_BOTTOM_BDR))
    instrs = [text.text or "" for text in para._p.iter(_QN_INSTR_TEXT)]
    for fld in para._p.iter(_QN_FLD_SIMPLE):
        instr = fld.get(_QN_FLD_INSTR)
        if instr:
            instrs.append(instr)
    return HeaderFooterParaInfo(
        text=para.text or "",
        style_name=getattr(getattr(para, "style", None), "name", "") or "",
        alignment=props.get("alignment"),
        line_spacing=props.get("line_spacing"),
        runs=_build_hf_runs(para, document),
        bottom_border=bottom_border,
        field_instrs=instrs,
    )


def _build_hf_paras(container: Any, document: Any) -> tuple[list[HeaderFooterParaInfo], bool]:
    """返回 (段落快照, 是否链接到上一节)；不修改文档（链接时直接跳过）。"""
    try:
        if container.is_linked_to_previous:
            return [], True
        return [_build_hf_para(para, document) for para in container.paragraphs], False
    except Exception:
        return [], False


def _build_section(index: int, section: Any, document: Any = None) -> SectionInfo:
    sect_pr = section._sectPr
    pg_sz = sect_pr.find(_QN_PG_SZ)
    pg_mar = sect_pr.find(_QN_PG_MAR)
    doc_grid = sect_pr.find(_QN_DOC_GRID)
    pg_num = sect_pr.find(_QN_PG_NUM_TYPE)

    width_cm = height_cm = None
    if pg_sz is not None:
        width_cm = twips_to_cm(pg_sz.get(qn("w:w")))
        height_cm = twips_to_cm(pg_sz.get(qn("w:h")))

    margins: dict = {}
    if pg_mar is not None:
        margins = {
            "margin_top_cm": twips_to_cm(pg_mar.get(qn("w:top"))),
            "margin_bottom_cm": twips_to_cm(pg_mar.get(qn("w:bottom"))),
            "margin_left_cm": twips_to_cm(pg_mar.get(qn("w:left"))),
            "margin_right_cm": twips_to_cm(pg_mar.get(qn("w:right"))),
            "gutter_cm": twips_to_cm(pg_mar.get(qn("w:gutter"))),
            "header_distance_cm": twips_to_cm(pg_mar.get(qn("w:header"))),
            "footer_distance_cm": twips_to_cm(pg_mar.get(qn("w:footer"))),
        }

    page_number_start = None
    if pg_num is not None:
        try:
            raw_start = pg_num.get(_QN_PG_NUM_START)
            page_number_start = int(raw_start) if raw_start is not None else None
        except (TypeError, ValueError):
            page_number_start = None

    header_paras: list[HeaderFooterParaInfo] = []
    footer_paras: list[HeaderFooterParaInfo] = []
    header_linked = footer_linked = False
    if document is not None:
        try:
            header_paras, header_linked = _build_hf_paras(section.header, document)
        except Exception:
            header_paras, header_linked = [], False
        try:
            footer_paras, footer_linked = _build_hf_paras(section.footer, document)
        except Exception:
            footer_paras, footer_linked = [], False

    page_field_instrs: list[str] = []
    for para in footer_paras:
        for instr in para.field_instrs:
            if _is_page_field(instr):
                page_field_instrs.append(instr)
    # 无文本的 PAGE 域（域代码）也认：footer part 里存在 instrText PAGE 即可
    if not page_field_instrs:
        try:
            footer_element = section.footer._element if not footer_linked else None
        except Exception:
            footer_element = None
        if footer_element is not None:
            for text in footer_element.iter(_QN_INSTR_TEXT):
                if _is_page_field(text.text or ""):
                    page_field_instrs.append(text.text or "")
            for fld in footer_element.iter(_QN_FLD_SIMPLE):
                if _is_page_field(fld.get(_QN_FLD_INSTR) or ""):
                    page_field_instrs.append(fld.get(_QN_FLD_INSTR) or "")

    return SectionInfo(
        index=index,
        page_width_cm=width_cm,
        page_height_cm=height_cm,
        header_text=_header_footer_text("header", section),
        footer_text=_header_footer_text("footer", section),
        page_number_format=pg_num.get(qn("w:fmt")) if pg_num is not None else None,
        page_number_start=page_number_start,
        doc_grid_type=doc_grid.get(qn("w:type")) if doc_grid is not None else None,
        doc_grid_line_pitch=(
            int(doc_grid.get(qn("w:linePitch")))
            if doc_grid is not None and doc_grid.get(qn("w:linePitch"))
            else None
        ),
        doc_grid_char_space=(
            int(doc_grid.get(qn("w:charSpace")))
            if doc_grid is not None and doc_grid.get(qn("w:charSpace"))
            else None
        ),
        header_linked=header_linked,
        footer_linked=footer_linked,
        header_paras=header_paras,
        footer_paras=footer_paras,
        has_page_field=bool(page_field_instrs),
        page_field_instrs=page_field_instrs,
        has_title_pg=sect_pr.find(_QN_TITLE_PG) is not None,
        **margins,
    )


# ---------------------------------------------------------------------------
# 分节范围 / 表格 / TOC 域（P1 增量：只读）
# ---------------------------------------------------------------------------
def _section_ranges(document: Any) -> list[tuple[int, int]]:
    """每个 section 覆盖的 body 段落 index 区间（含首尾）。

    依据：分节符存在于段落 pPr/w:sectPr 中；最后一节的 sectPr 在 body 末尾。
    """
    ranges: list[tuple[int, int]] = []
    start = 0
    para_count = 0
    for child in document.element.body.iterchildren():
        if child.tag != qn("w:p"):
            continue
        ppr = child.find(_QN_PPR)
        sect_pr = ppr.find(_QN_SECT_PR) if ppr is not None else None
        if sect_pr is not None:
            ranges.append((start, para_count))
            start = para_count + 1
        para_count += 1
    if para_count:
        ranges.append((start, para_count - 1))
    return ranges


def _build_cell(cell: Any, document: Any) -> TableCellInfo:
    texts: list[str] = []
    runs: list[RunInfo] = []
    alignment = None
    line_spacing = None
    seen_nonempty = False
    try:
        paragraphs = cell.paragraphs
    except Exception:
        paragraphs = []
    for para in paragraphs:
        text = para.text or ""
        texts.append(text)
        props = _effective_para_props(para, document)
        if not seen_nonempty and text.strip():
            alignment = props.get("alignment")
            line_spacing = props.get("line_spacing")
            seen_nonempty = True
        style = getattr(para, "style", None)
        for run in para.runs:
            font = _effective_run_font(run, style, document)
            runs.append(
                RunInfo(
                    text=run.text or "",
                    east_asia=font.get("east_asia"),
                    ascii=font.get("ascii"),
                    size_pt=font.get("size_pt"),
                    bold=font.get("bold"),
                    italic=font.get("italic"),
                )
            )
    return TableCellInfo(
        text=" / ".join(part for part in texts if part),
        runs=runs,
        alignment=alignment,
        line_spacing=line_spacing,
    )


def _build_table(
    index: int,
    table: Any,
    prev_para_index: Optional[int],
    next_para_index: Optional[int],
    document: Any,
) -> TableInfo:
    table_pr = table._tbl.tblPr
    borders = _borders_from_container(table_pr, _QN_TBL_BORDERS)
    style = getattr(table, "style", None)
    style_name = getattr(style, "name", "") or ""
    if not borders:
        borders = _style_table_borders(style)

    rows = []
    try:
        rows = list(table.rows)
    except Exception:
        rows = []
    cells: list[list[TableCellInfo]] = []
    for row in rows:
        try:
            row_cells = list(row.cells)
        except Exception:
            row_cells = []
        cells.append([_build_cell(cell, document) for cell in row_cells])

    first_row_bottom_borders: list[Optional[BorderInfo]] = []
    if rows:
        try:
            for cell in rows[0].cells:
                tc_pr = cell._tc.tcPr
                first_row_bottom_borders.append(
                    _borders_from_container(tc_pr, _QN_TC_BORDERS).get("bottom")
                )
        except Exception:
            first_row_bottom_borders = []

    column_count = 0
    if cells:
        column_count = max((len(row) for row in cells), default=0)

    return TableInfo(
        index=index,
        prev_para_index=prev_para_index,
        next_para_index=next_para_index,
        row_count=len(rows),
        column_count=column_count,
        style_name=style_name,
        borders=borders,
        first_row_bottom_borders=first_row_bottom_borders,
        cells=cells,
    )


def _build_tables(document: Any) -> list[TableInfo]:
    """按 body 顺序构建表格快照（含表格前后段落 index，用于题注定位）。"""
    try:
        tables = list(document.tables)
    except Exception:
        return []
    if not tables:
        return []

    anchors: list[tuple[int, int]] = []
    para_count = 0
    table_cursor = 0
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            para_count += 1
        elif child.tag == qn("w:tbl"):
            if table_cursor < len(tables):
                anchors.append((para_count - 1, para_count))
            table_cursor += 1

    total_paragraphs = len(document.paragraphs)
    infos: list[TableInfo] = []
    for offset, table in enumerate(tables):
        prev_index, next_index = anchors[offset] if offset < len(anchors) else (-1, -1)
        if prev_index is not None and prev_index < 0:
            prev_index = None
        if next_index is not None and (next_index < 0 or next_index >= total_paragraphs):
            next_index = None
        infos.append(_build_table(offset + 1, table, prev_index, next_index, document))
    return infos


def _has_toc_field(body_element: Any) -> bool:
    """正文里是否含 TOC 域（可更新目录）。"""
    try:
        for instr in body_element.iter(_QN_INSTR_TEXT):
            if re.search(r"\bTOC\b", (instr.text or "").upper()):
                return True
        for fld in body_element.iter(_QN_FLD_SIMPLE):
            if re.search(r"\bTOC\b", (fld.get(_QN_FLD_INSTR) or "").upper()):
                return True
    except Exception:
        return False
    return False
