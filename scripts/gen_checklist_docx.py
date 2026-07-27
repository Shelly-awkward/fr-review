# -*- coding: utf-8 -*-
r"""
gen_checklist_docx.py — 管區意見內容 JSON → Word（python-docx 版）。

    python scripts/gen_checklist_docx.py out/8304_114_checklist_content.json out/管區意見.docx

與 gen_checklist_docx.js（docx-js 版）產出同體例，擇一使用即可：
沒有 Node 的環境（例如只有 Python 沙箱的 AI）用本檔，其餘用 .js 版。
需求：pip install python-docx
"""
import json
import sys

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

FONT = "PMingLiU"          # 新細明體，貼近原模板
W = {"g": 1300, "t": 5200, "y": 900, "n": 900, "b": 1740}   # DXA（＝twips）


def styled(run, size=12, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    # 中文字型須另設 eastAsia，否則 Word 會用預設字型渲染中文
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return run


def para(container, text="", size=12, bold=False, align=None, indent=None,
         after=4, line=None):
    p = container.add_paragraph()
    styled(p.add_run(text), size, bold)
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(after)
    if indent is not None:
        pf.left_indent = Twips(indent)
    if line is not None:
        pf.line_spacing = line
    return p


def fill_cell(cell, lines, size=12, bold=False, align=None):
    """把多行文字寫進儲存格（清掉 python-docx 預設的空段落）。"""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell._element.clear_content()
    for line in (lines or [""]):
        p = cell.add_paragraph()
        styled(p.add_run(line), size, bold)
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_after = Pt(0)


def build_checklist_table(doc, groups):
    n_items = sum(len(g["items"]) for g in groups)
    table = doc.add_table(rows=2 + n_items, cols=5)
    table.style = "Table Grid"
    table.autofit = False

    # 表頭兩列：內容／檢查內容（各跨兩列）＋填報項目（跨三欄）→ 是／否／備註
    table.cell(0, 0).merge(table.cell(1, 0))
    table.cell(0, 1).merge(table.cell(1, 1))
    table.cell(0, 2).merge(table.cell(0, 4))
    fill_cell(table.cell(0, 0), ["內容"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    fill_cell(table.cell(0, 1), ["檢查內容"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    fill_cell(table.cell(0, 2), ["填報項目"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for col, (head, sub) in enumerate([("是", "(正常)"), ("否", "(異常)"), ("備註", None)], 2):
        cell = table.cell(1, col)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell._element.clear_content()
        for text, size in [(head, 12)] + ([(sub, 10)] if sub else []):
            p = cell.add_paragraph()
            styled(p.add_run(text), size, bold=(size == 12))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)

    r = 2
    for g in groups:
        first = r
        for it in g["items"]:
            fill_cell(table.cell(r, 1), it["text"].split("\n"), size=11)
            fill_cell(table.cell(r, 2), ["V" if it.get("mark") == "yes" else ""],
                      align=WD_ALIGN_PARAGRAPH.CENTER)
            fill_cell(table.cell(r, 3), ["V" if it.get("mark") == "no" else ""],
                      align=WD_ALIGN_PARAGRAPH.CENTER)
            fill_cell(table.cell(r, 4),
                      [t for t in (it.get("note") or "").split("\n") if t], size=10)
            r += 1
        merged = table.cell(first, 0)
        if r - 1 > first:
            merged = merged.merge(table.cell(r - 1, 0))
        fill_cell(merged, [g["group"]])

    for row in table.rows:
        for cell, key in zip(row.cells, ["g", "t", "y", "n", "b"]):
            cell.width = Twips(W[key])
    return table


def add_mini_table(doc, rows):
    """把連續的 【表】a｜b｜c 段落轉成小表格。"""
    n = len(rows[0])
    table = doc.add_table(rows=len(rows), cols=n)
    table.style = "Table Grid"
    for ri, cells in enumerate(rows):
        for ci, text in enumerate(cells):
            align = (WD_ALIGN_PARAGRAPH.CENTER
                     if ri == 0 or ci == n - 1 else None)
            fill_cell(table.cell(ri, ci), [text], size=10, bold=(ri == 0), align=align)
    para(doc, "", after=2)


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    with open(sys.argv[1], encoding="utf-8") as f:
        c = json.load(f)

    doc = Document()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Twips(1080)
        section.left_margin = section.right_margin = Twips(1080)

    para(doc, c["title"], size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
    build_checklist_table(doc, c["groups"])
    for note in c.get("footnotes", []):
        para(doc, note, size=10, after=2)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    for s in c["sections"]:
        para(doc, s["title"], size=13, bold=True, after=6)
        for b in s["body"]:
            if b.get("h"):
                para(doc, b["h"], bold=True, indent=360, after=3)
            pending = []
            for p in b.get("paras", []):
                if p.startswith("【表】"):
                    pending.append(p[3:].split("｜"))
                    continue
                if pending:
                    add_mini_table(doc, pending)
                    pending = []
                para(doc, p, indent=720 if b.get("h") else 480, after=4, line=1.3)
            if pending:
                add_mini_table(doc, pending)
        para(doc, "", after=3)

    doc.save(sys.argv[2])
    print(f"寫出 {sys.argv[2]}")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
