"""规范文档读取（教师端入口）：把老师给的规范文档转为纯文本，供规则抽取。"""

from __future__ import annotations

from pathlib import Path

SUPPORTED_SUFFIXES = {".docx", ".txt", ".md", ".markdown"}


def extract_spec_text(path: Path | str) -> str:
    """读取规范文档文本。

    支持 .docx（段落 + 表格逐行）、.txt/.md（UTF-8 / GB18030 自动识别）。
    PDF 与 .doc 暂不支持，给出可操作的提示。
    """
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"规范文档不存在：{path}")
    suffix = path.suffix.lower()

    if suffix == ".docx":
        import docx

        document = docx.Document(str(path))
        parts: list[str] = [para.text for para in document.paragraphs]
        for table in document.tables:
            for row in table.rows:
                cells = [(cell.text or "").strip().replace("\n", " ") for cell in row.cells]
                if any(cells):
                    parts.append(" | ".join(cells))
        return "\n".join(part for part in parts if part is not None).strip() + "\n"

    if suffix in (".txt", ".md", ".markdown"):
        raw = path.read_bytes()
        for encoding in ("utf-8-sig", "utf-8", "gb18030"):
            try:
                return raw.decode(encoding)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="replace")

    if suffix == ".pdf":
        raise ValueError(
            "PDF 规范文档暂未直接支持：请先用 Word/WPS 转成 .docx，"
            "或接入 docling 解析（后续迭代）。"
        )
    if suffix == ".doc":
        raise ValueError("暂不支持 .doc 规范文档；请先另存为 .docx。")
    raise ValueError(f"不支持的规范文档格式：{suffix}（支持 .docx/.txt/.md）")


def spec_stats(text: str) -> dict:
    lines = text.splitlines()
    return {
        "chars": len(text),
        "lines": len(lines),
        "paragraphs": len([line for line in lines if line.strip()]),
    }
