#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把《转折》项目的 Markdown 文档转换成排版好的 Word 交付件。

用法（在 zunyi-scenario 目录下）：
    <bundled python> build_docs.py

输出：zunyi-scenario/交付/*.docx
"""
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, "交付")

BODY_FONT = "宋体"
HEAD_FONT = "黑体"
QUOTE_FONT = "楷体"
GRAY = RGBColor(0x59, 0x59, 0x59)
DARK = RGBColor(0x22, 0x22, 0x22)

DOCS = [
    {
        "md": "01_剧本_舞台演出本.md",
        "out": "《转折》遵义会议情景剧_剧本（5分钟版）.docx",
        "kind": "script",
        "header": "《转折》遵义会议情景剧 · 5 分钟舞台版",
        "body_size": 12,
        "line_spacing": 1.4,
    },
    {
        "md": "02_导演调度本.md",
        "out": "《转折》导演调度本.docx",
        "kind": "director",
        "header": "《转折》导演调度本 · 5 分钟舞台版",
        "body_size": 10.5,
        "line_spacing": 1.3,
    },
    {
        "md": "03_演员背词台本.md",
        "out": "《转折》演员背词台本.docx",
        "kind": "actors",
        "header": "《转折》演员背词台本 · 背词专用，不含灯音提示",
        "body_size": 14,
        "line_spacing": 1.5,
    },
]


# ----------------------------------------------------------------------------
# 基础工具
# ----------------------------------------------------------------------------
def set_run_font(run, font=BODY_FONT, size=None, bold=None, italic=None, color=None):
    """设置 run 字体（含中文 eastAsia），避免中文回退成默认字体。"""
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:eastAsia"), font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def add_runs_with_inline_bold(paragraph, text, font=BODY_FONT, size=None, color=None):
    """按 **加粗** 切分文本，逐段加 run。"""
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, font=HEAD_FONT if font == BODY_FONT else font,
                         size=size, bold=True, color=color)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, font=font, size=size, color=color)


def set_paragraph_format(p, first_line_indent=None, left_indent=None, space_before=0,
                         space_after=0, line_spacing=None, align=None):
    pf = p.paragraph_format
    if first_line_indent is not None:
        pf.first_line_indent = Cm(first_line_indent)
    if left_indent is not None:
        pf.left_indent = Cm(left_indent)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    if align is not None:
        pf.alignment = align


def add_heading(doc, text, level, body_size=12):
    p = doc.add_paragraph()
    sizes = {0: 22, 1: 16, 2: 14, 3: 12.5, 4: 12}
    size = sizes.get(level, 12)
    if level == 0:
        set_paragraph_format(p, space_before=0, space_after=18, line_spacing=1.3,
                             align=WD_ALIGN_PARAGRAPH.CENTER)
    else:
        set_paragraph_format(p, space_before=14 if level <= 2 else 10,
                             space_after=6, line_spacing=1.25)
    run = p.add_run(text)
    set_run_font(run, font=HEAD_FONT, size=size, bold=True, color=DARK)
    # 保留大纲层级，便于 Word 导航
    pPr = p._element.get_or_add_pPr()
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(max(0, level - 1)))
    pPr.append(outline)
    return p


def add_page_number_footer(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    set_run_font(run, font=BODY_FONT, size=9, color=GRAY)
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)


def add_header_text(section, text):
    p = section.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(text)
    set_run_font(run, font=QUOTE_FONT, size=9, color=GRAY)


def shade_cell(cell, hex_color="F2F2F2"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def set_cell_margins(table, top=40, bottom=40, left=80, right=80):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement("w:" + name)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def add_table(doc, rows, body_size=10.5):
    """rows: [[cell, ...], ...]，第一行为表头。"""
    n_cols = max(len(r) for r in rows)
    rows = [r + [""] * (n_cols - len(r)) for r in rows]
    font_size = {1: 10.5, 2: 10.5, 3: 10, 4: 9.5, 5: 9, 6: 8.5}.get(n_cols, 8)
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_cell_margins(table)
    total_cm = 15.6
    col_cm = total_cm / n_cols
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx]
        if r_idx == 0:
            set_repeat_table_header(row)
        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Cm(col_cm)
            p = cell.paragraphs[0]
            set_paragraph_format(p, space_before=1, space_after=1, line_spacing=1.1)
            text = cell_text.replace("<br>", "\n").strip()
            if r_idx == 0:
                run = p.add_run(text)
                set_run_font(run, font=HEAD_FONT, size=font_size, bold=True, color=DARK)
                shade_cell(cell)
            else:
                add_runs_with_inline_bold(p, text, font=BODY_FONT, size=font_size,
                                          color=DARK)
    # 表后留白
    spacer = doc.add_paragraph()
    set_paragraph_format(spacer, space_before=0, space_after=0, line_spacing=1.0)
    spacer.add_run("")
    return table


# ----------------------------------------------------------------------------
# Markdown 解析
# ----------------------------------------------------------------------------
def split_blocks(lines):
    blocks = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("#"):
            m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
            if m:
                blocks.append(("heading", len(m.group(1)), m.group(2).strip()))
                i += 1
                continue
        if stripped.startswith("|"):
            table_lines = []
            while i < n and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            blocks.append(("table", table_lines))
            continue
        if stripped.startswith(">"):
            quote_lines = []
            while i < n and lines[i].strip().startswith(">"):
                quote_lines.append(lines[i].strip()[1:].strip())
                i += 1
            blocks.append(("quote", quote_lines))
            continue
        if re.match(r"^[-*]\s+", stripped):
            items = []
            while i < n and re.match(r"^[-*]\s+", lines[i].strip()):
                items.append(re.sub(r"^[-*]\s+", "", lines[i].strip()))
                i += 1
            blocks.append(("ul", items))
            continue
        if re.match(r"^\d+\.\s+", stripped):
            items = []
            while i < n and re.match(r"^\d+\.\s+", lines[i].strip()):
                items.append(lines[i].strip())
                i += 1
            blocks.append(("ol", items))
            continue
        if re.match(r"^-{3,}$", stripped):
            blocks.append(("hr", None))
            i += 1
            continue
        # 普通段落：合并直到空行或下一个结构行
        para = [stripped]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith("#") or nxt.startswith("|")
                    or nxt.startswith(">") or re.match(r"^[-*]\s+", nxt)
                    or re.match(r"^\d+\.\s+", nxt) or re.match(r"^-{3,}$", nxt)):
                break
            para.append(nxt)
            i += 1
        blocks.append(("para", " ".join(para)))
    return blocks


def parse_table(table_lines):
    rows = []
    for line in table_lines:
        if re.match(r"^\|[\s:\-|]+\|?$", line):
            continue  # 分隔行
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        rows.append(cells)
    return rows


# ----------------------------------------------------------------------------
# 文档生成
# ----------------------------------------------------------------------------
def build_doc(cfg):
    md_path = os.path.join(BASE_DIR, cfg["md"])
    if not os.path.exists(md_path):
        print("[skip] 未找到 %s" % cfg["md"])
        return None
    with open(md_path, encoding="utf-8") as f:
        lines = f.read().splitlines()
    blocks = split_blocks(lines)

    body_size = cfg["body_size"]
    line_spacing = cfg["line_spacing"]
    kind = cfg["kind"]

    doc = Document()
    # 页面：A4，适中页边距
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.3)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)
    add_page_number_footer(section)
    add_header_text(section, cfg["header"])

    # 默认样式
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(body_size)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)

    # ---------- 封面 ----------
    title = None
    cover_sub = []
    idx = 0
    used_first_heading = False
    for b in blocks:
        if b[0] == "heading" and not used_first_heading:
            title = b[2]
            used_first_heading = True
            idx += 1
            continue
        if used_first_heading and b[0] == "quote":
            cover_sub.extend([q for q in b[1] if q])
            idx += 1
            continue
        break
    if title is None:
        title = os.path.splitext(cfg["md"])[0]
    add_heading(doc, title, 0, body_size)
    for sub in cover_sub:
        p = doc.add_paragraph()
        set_paragraph_format(p, space_before=2, space_after=2, line_spacing=1.4,
                             align=WD_ALIGN_PARAGRAPH.CENTER)
        add_runs_with_inline_bold(p, sub, font=QUOTE_FONT, size=11,
                                  color=RGBColor(0x40, 0x40, 0x40))
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)

    body_blocks = blocks[idx:]

    for b in body_blocks:
        kind_b = b[0]
        if kind_b == "heading":
            level = min(b[1], 4)
            text = re.sub(r"\*\*", "", b[2])
            add_heading(doc, text, level, body_size)
        elif kind_b == "table":
            rows = parse_table(b[1])
            if rows:
                add_table(doc, rows, body_size)
        elif kind_b == "quote":
            for q in b[1]:
                if not q:
                    continue
                p = doc.add_paragraph()
                set_paragraph_format(p, left_indent=0.6, space_before=2, space_after=2,
                                     line_spacing=1.3)
                add_runs_with_inline_bold(p, q, font=QUOTE_FONT, size=max(9.5, body_size - 1),
                                          color=GRAY)
        elif kind_b == "ul":
            for item in b[1]:
                p = doc.add_paragraph()
                set_paragraph_format(p, left_indent=0.75, first_line_indent=-0.4,
                                     space_before=1, space_after=1, line_spacing=line_spacing)
                add_runs_with_inline_bold(p, "· " + item, font=BODY_FONT, size=body_size,
                                          color=DARK)
        elif kind_b == "ol":
            for item in b[1]:
                p = doc.add_paragraph()
                set_paragraph_format(p, left_indent=0.75, first_line_indent=-0.55,
                                     space_before=1, space_after=1, line_spacing=line_spacing)
                add_runs_with_inline_bold(p, item, font=BODY_FONT, size=body_size, color=DARK)
        elif kind_b == "hr":
            p = doc.add_paragraph()
            set_paragraph_format(p, space_before=4, space_after=4, line_spacing=1.0)
            run = p.add_run("")
            run.add_break(WD_BREAK.LINE)
        elif kind_b == "para":
            text = b[1]
            # 台词：**角色：** 台词
            m = re.match(r"^\*\*([^*：:\n]{1,10})[：:]\*\*\s*(.*)$", text)
            if m:
                role, content = m.group(1), m.group(2)
                p = doc.add_paragraph()
                set_paragraph_format(p, left_indent=0.4, first_line_indent=-0.4,
                                     space_before=2, space_after=2, line_spacing=line_spacing)
                run = p.add_run(role + "：")
                set_run_font(run, font=HEAD_FONT, size=body_size, bold=True, color=DARK)
                add_runs_with_inline_bold(p, content, font=BODY_FONT, size=body_size,
                                          color=DARK)
                continue
            # 舞台提示：（……） 或 【……】
            if (text.startswith("（") and text.endswith("）")) or \
               (text.startswith("【") and text.endswith("】")):
                p = doc.add_paragraph()
                set_paragraph_format(p, left_indent=0.4, space_before=1, space_after=1,
                                     line_spacing=max(1.2, line_spacing - 0.1))
                add_runs_with_inline_bold(p, text, font=QUOTE_FONT,
                                          size=max(9.5, body_size - 1), color=GRAY)
                continue
            # 普通正文
            p = doc.add_paragraph()
            set_paragraph_format(p, first_line_indent=0.74, space_before=2, space_after=2,
                                 line_spacing=line_spacing)
            add_runs_with_inline_bold(p, text, font=BODY_FONT, size=body_size, color=DARK)

    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, cfg["out"])
    doc.save(out_path)
    return out_path


def validate(path):
    doc = Document(path)
    n_para = len(doc.paragraphs)
    n_tab = len(doc.tables)
    texts = [p.text for p in doc.paragraphs]
    joined = "\n".join(texts)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                joined += "\n" + cell.text
    checks = {
        "非空": len(joined.strip()) > 500,
        "无乱码替换符": "\ufffd" not in joined,
        "含项目关键词": ("遵义会议" in joined),
    }
    print("  [校验] %s：段落 %d，表格 %d，字符 %d，%s" % (
        os.path.basename(path), n_para, n_tab, len(joined),
        "；".join("%s=%s" % (k, "OK" if v else "FAIL") for k, v in checks.items())))
    return all(checks.values())


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = []
    for cfg in DOCS:
        path = build_doc(cfg)
        if path:
            ok = validate(path)
            results.append((path, ok))
    print("\n完成：")
    for path, ok in results:
        print("  [%s] %s" % ("OK" if ok else "FAIL", path))
    if not results:
        print("  （没有可生成的文档，请先准备 md 文件）")


if __name__ == "__main__":
    main()
