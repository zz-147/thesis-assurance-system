"""生成 tas 老师端演示报告（一页总结 + 完整版 + 简历摘要）。

运行：
    python demo/build_report.py
产物：
    demo/演示报告_一页总结.docx / .pdf
    demo/演示报告_完整版.docx / .pdf
    demo/简历摘要.md
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT  # noqa: F401
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
DEMO = ROOT / "demo"
SHOTS = DEMO / "截图"
REPORT = DEMO / "报告"

DARK = RGBColor(0x1F, 0x38, 0x64)
GRAY = RGBColor(0x6B, 0x72, 0x80)
ACCENT = RGBColor(0x1D, 0x4E, 0xD8)


def font(run, size: float, *, bold: bool = False, color=None, east: str = "微软雅黑", ascii_font: str = "Calibri") -> None:
    run.font.name = ascii_font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), east)


def para(doc, text: str, size: float = 10.5, *, bold=False, color=None, space_after: float = 4, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if align is not None:
        p.alignment = align
    font(p.add_run(text), size, bold=bold, color=color)
    return p


def heading(doc, text: str, *, size: float = 12.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    font(p.add_run(text), size, bold=True, color=DARK)
    return p


def shade_cell(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:fill"): color})
    tc_pr.append(shd)


def metric_table(doc, rows: list[tuple[str, str]], *, width_cm: float = 17.0):
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].width = Cm(width_cm * 0.42)
        cells[1].width = Cm(width_cm * 0.58)
        p0 = cells[0].paragraphs[0]
        p0.paragraph_format.space_after = Pt(1)
        p0.paragraph_format.space_before = Pt(1)
        font(p0.add_run(label), 9.5, bold=True, color=DARK)
        shade_cell(cells[0], "EEF2F7")
        p1 = cells[1].paragraphs[0]
        p1.paragraph_format.space_after = Pt(1)
        p1.paragraph_format.space_before = Pt(1)
        font(p1.add_run(value), 9.5)
    return table


def crop(src: Path, dst: Path, box: tuple[int, int, int, int]) -> Path:
    image = Image.open(src)
    image.crop(box).save(dst)
    return dst


def setup(doc) -> None:
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)


def read_summary() -> dict:
    return json.loads((REPORT / "batch-summary.json").read_text(encoding="utf-8"))["summary"]


def read_refs() -> dict:
    return json.loads((REPORT / "refs-report.json").read_text(encoding="utf-8"))["summary"]


def read_fix_summary() -> str:
    with zipfile.ZipFile(REPORT / "fixed.zip") as archive:
        return archive.read("修复汇总.txt").decode("utf-8")


def build_one_page() -> Path:
    doc = Document()
    setup(doc)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(2)
    font(title.add_run("TAS 老师端：论文格式批量检查与反馈"), 18, bold=True, color=DARK)
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.paragraph_format.space_after = Pt(6)
    font(meta.add_run("演示报告 ｜ 2026-10-30 ｜ 演示数据：8 人名单 / 6 份提交（生成论文，非真实学生）"), 9, color=GRAY)

    para(
        doc,
        "老师收齐微信里的 .docx → 一次批量检查 + 一键修复 + 名单核对 + 首页缩略图 + 文献核验；"
        "学生反馈自动生成可转发的 txt，全部在本机处理。",
        10.5,
        space_after=5,
    )

    refs = read_refs()
    metric_table(
        doc,
        [
            ("演示数据", "6 份论文 / 8 人名单（含重复提交、缺交、文件名不规范）"),
            ("批量检查", "检查成功 6 ｜ 通过率 0%（演示论文均有格式问题）｜ 错误 81 / 警告 145 / 提示 0"),
            ("共性问题 Top", "BODY-004 对齐 133 次 ｜ TAB-002 表内文字 51 次 ｜ STRUCT-001 装订缺件 12 次"),
            ("一键修复", "36 项问题/份级批量修复：错误 81 → 0；只改格式，内容 hash 不变，原文件不覆盖"),
            ("名单核对", "应交 8 ｜ 收到 6 ｜ 匹配 5 ｜ 缺交 3（王五/周九/吴十）｜ 重复 1（李四）"),
            ("首页缩略图", "6/6 渲染成功（Word COM）→ thumbnails.zip + HTML 报告缩略图列"),
            (
                "文献核验",
                f"每份前 2 条：verified {refs['verified']} ｜ likely {refs['likely']} ｜ "
                f"not_found {refs['not_found']} ｜ error {refs['error']}"
                "（not_found ≠ 造假，需人工核实）",
            ),
        ],
    )

    home = crop(SHOTS / "01_首页.png", SHOTS / "01_首页_crop.png", (0, 0, 1440, 1120))
    doc.add_picture(str(home), width=Cm(15.6))
    pic_para = doc.paragraphs[-1]
    pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_para.paragraph_format.space_before = Pt(4)

    para(
        doc,
        "使用：双击「启动老师端.bat」→ 教师端选文件夹、粘贴名单、勾选修复/缩略图/文献核验 → 批量检查 → "
        "下载 feedback.zip（逐份反馈）、roster-report.txt、fixed.zip、thumbnails.zip、refs-report.md。",
        9,
        color=GRAY,
        space_after=2,
    )
    para(
        doc,
        "边界：只检查格式，不改文字/参考文献；not_found 仅表示当前索引未检索到，不等于造假；缩略图需 Windows + Word。",
        9,
        color=GRAY,
    )

    out = DEMO / "演示报告_一页总结.docx"
    doc.save(out)
    return out


def build_full() -> Path:
    doc = Document()
    setup(doc)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(title.add_run("TAS 老师端演示报告（完整版）"), 20, bold=True, color=DARK)
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.paragraph_format.space_after = Pt(8)
    font(meta.add_run("2026-10-30 ｜ 演示数据：8 人名单 / 6 份生成论文（非真实学生）｜ 全部本机处理"), 9, color=GRAY)

    heading(doc, "1. 演示场景")
    para(
        doc,
        "模拟一位老师用微信收齐班级论文，直接交给 tas：不用逐个打开 Word，批量输出通过率、问题清单、"
        "学生反馈、修复版、名单核对、首页缩略图和文献核验报告。",
    )

    heading(doc, "2. 演示数据")
    metric_table(
        doc,
        [
            ("名单", "8 人：张三、李四、王五、赵六、钱七、孙八、周九、吴十"),
            ("提交文件", "6 份：张三/李四×2（重复）/赵六（文件名无学号）/钱七/孙八；王五、周九、吴十未交"),
            ("预期", "自动识别缺交 3、重复 1、文件名提醒 1"),
        ],
        width_cm=17.2,
    )

    heading(doc, "3. 批量检查结果")
    summary = read_summary()
    counts = summary["counts"]
    metric_table(
        doc,
        [
            ("文件总数 / 检查成功", f"{summary['total']} / {summary['checked']}"),
            ("通过率", f"{summary['pass_rate']:.1%}"),
            ("错误 / 警告 / 提示", f"{counts['error']} / {counts['warning']} / {counts['info']}"),
        ],
        width_cm=17.2,
    )
    para(doc, "共性问题 Top（按命中次数）：", 10, bold=True, space_after=2)
    for rule, hits in summary["top_rules"][:5]:
        para(doc, f"• {rule}：{hits} 次", 9.5, space_after=1)

    report_shot = crop(SHOTS / "02_批量报告.png", SHOTS / "02_批量报告_crop.png", (0, 0, 1440, 1800))
    doc.add_picture(str(report_shot), width=Cm(17.0))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    para(doc, "图：批量报告（通过率、共性问题条形图、逐份缩略图列）", 8.5, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER)

    heading(doc, "4. 一键修复与名单核对")
    para(doc, "批量修复包 fixed.zip：每份文档生成 *_fixed.docx（只改格式，内容 hash 不变，原文件不覆盖）。")
    for line in read_fix_summary().splitlines()[:6]:
        if line.strip():
            para(doc, line.strip(), 8.5, space_after=0)
    para(
        doc,
        "名单核对：roster-report.txt 输出缺交/重复/命名提醒；缺交 3 人、重复 1 人（李四提交了两份，"
        "文件名带 (1) 会被标记为多余副本）。",
        9.5,
        space_after=2,
    )

    heading(doc, "5. 学生反馈（可直接转发）")
    para(
        doc,
        "feedback.zip 内每份文档一个 反馈_学号_姓名.txt：结论 + 问题统计 + 主要问题（含实际值/期望值）+ 修改建议；"
        "00_班级汇总.txt 含通过率、共性问题、未通过名单。",
        9.5,
    )
    feedback_shot = crop(SHOTS / "03_反馈示例.png", SHOTS / "03_反馈示例_crop.png", (0, 0, 1200, 880))
    doc.add_picture(str(feedback_shot), width=Cm(16.2))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    heading(doc, "6. 缩略图与文献核验")
    refs = read_refs()
    para(
        doc,
        f"缩略图：6/6 成功（每份首页 PNG，HTML 报告内嵌缩略图列，另打包 thumbnails.zip）。",
        9.5,
        space_after=1,
    )
    para(
        doc,
        f"文献核验：{refs['total_docs']} 份文档、{refs['total_entries']} 条文献中按每份前 2 条核验 → "
        f"verified {refs['verified']} ｜ likely {refs['likely']} ｜ not_found {refs['not_found']} ｜ "
        f"error {refs['error']}；报告 refs-report.md/json，含证据 URL 与诚实性提示。",
        9.5,
        space_after=1,
    )
    para(doc, "注意：not_found 仅表示未在当前公开索引检索到，中文期刊可能未被收录，不等于文献造假，需人工核实。", 8.5, color=GRAY)

    heading(doc, "7. 产物与使用方式")
    metric_table(
        doc,
        [
            ("批量报告", "batch.xlsx（3 Sheet）/ batch.csv / batch.html"),
            ("学生反馈", "feedback.zip（00_班级汇总.txt + 反馈_*.txt）"),
            ("修复 / 名单", "fixed.zip（*_fixed.docx + 修复报告/汇总）；roster-report.txt"),
            ("缩略图 / 文献", "thumbnails.zip；refs-report.md / refs-report.json"),
            ("老师使用", "双击 启动老师端.bat → 教师端勾选功能 → 批量检查；或 CLI：tas batch 目录 --fix --roster 名单 --preview --refs"),
            ("免安装", "发给老师 dist/tas-web-portable.zip（18.6 MB，内置 Python，解压即用）"),
        ],
        width_cm=17.2,
    )

    heading(doc, "8. 边界与说明")
    para(doc, "• 只检查格式，不改文字、不代写、不动参考文献；修复版是副本，原文件不覆盖。", 9.5, space_after=1)
    para(doc, "• 所有学生文档只在本机处理；文献核验只发题名/作者到公开元数据 API，不上传全文。", 9.5, space_after=1)
    para(doc, "• 质量指标来自合成变异集（检出率 100%、误报 0/201 段、修复后 P0 通过率 100%）；演示评论文档为生成数据。", 9.5, space_after=1)
    para(doc, "• 缩略图/页面预览依赖 Windows + Microsoft Word；目录页码一致性等需渲染引擎的项仍只报告。", 9.5, space_after=1)

    out = DEMO / "演示报告_完整版.docx"
    doc.save(out)
    return out


def build_resume() -> Path:
    summary = read_summary()
    counts = summary["counts"]
    text = f"""# 简历摘要（tas 项目）

> 数字来自仓库内可复现代码与评测脚本；演示数据为 8 人名单 / 6 份生成论文。

## 一句话

独立开发「论文格式规范化与真实性核验工具」：规则库 38 条、检查器覆盖 **38/38**；批量检查/安全修复/学生反馈/名单核对/批量缩略图/文献核验，Windows 本地运行、无账号、数据不出本机。

## 核心数字

- 规则覆盖：**38/38（P0 23/23，P1 15/15）**；LLM 抽取 profile 44/44；自动化测试 **321 个**（含真实网络/LLM/Word 渲染 live）。
- 合成变异集：116 处注入 / 23 类 → 检出率 **100%**、定位准确率 **100%**；合规 fixture 误报 **0/201 段**。
- 安全修复：修复后 P0 通过率 **100%**、内容 hash 一致率 **100%**；单份检查 0.09s、批量吞吐 **9.8 份/秒**。
- 规范→规则（DeepSeek）：9455 字学院规范 **38/38** 逐条评估（11s），抽出 6 条人工遗漏规则、23 条 unmapped。
- 老师端演示（本次）：6 份论文批量检查 → 错误 **{counts['error']}** / 警告 **{counts['warning']}**；一键修复 81 → 0；缺交 3、重复 1；缩略图 6/6；文献核验 likely 5 / not_found 7。
- 多学科参数化：同一论文切换规范，错误数 5 → 299，并按自定义参数（10.5pt 仿宋 / 1.75 倍行距 / 3cm 边距 / 14pt 标题）修复到 0 复检问题。

## 技术要点

- Python + python-docx + lxml（OOXML 解析/写入），规则与代码分离（JSON rules + profiles），checker/fixer 参数化。
- 安全修复白名单：只改格式属性；内容 hash 校验；参考文献/目录条目零改动；自动复检、幂等、原文件不覆盖。
- Web 双入口（标准库 http.server）+ 手机 H5 + 微信小程序源码；拖拽上传、页面预览（Word COM → PDF → PNG）。
- 教师端批量：xlsx/csv/html、学生反馈包、名单核对、批量缩略图、批量文献真实性核验（Crossref/OpenAlex/arXiv，not_found ≠ 造假）。
- 免安装绿色版（内置 Python 运行时）与一键启动/停止脚本；离线优先、隐私本地化。

## 可讲的故事

1. 从「老师手工检查 40 份论文格式」到「拖文件夹 → 几分钟出班级报告 + 每人反馈」。
2. 为什么用规则驱动 + LLM 只做规范抽取、不做修改：可解释、可复现、可审计。
3. 修复引擎如何保证不改学生内容：白名单 + 内容 hash + 复检 + 原文件不覆盖。
4. 工程化取舍：PyPI 受限 → 改用内置运行时绿色版；Word 渲染可选、失败不阻塞检查。
"""
    out = DEMO / "简历摘要.md"
    out.write_text(text, encoding="utf-8")
    return out


if __name__ == "__main__":
    one = build_one_page()
    full = build_full()
    resume = build_resume()
    print("一页总结：", one)
    print("完整版：", full)
    print("简历摘要：", resume)
