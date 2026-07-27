# -*- coding: utf-8 -*-
r"""
gen_inquiry_xlsx.py — inquiry.json → 「<年度>財務報告說明」查詢函 Excel（給受查公司填答）。

    python scripts/gen_inquiry_xlsx.py out/8304_114_inquiry.json out/8304_114_財務報告說明.xlsx

五個 sheet：差異說明（自動生成題目＋預填數字，「說明」欄留白給公司）、基本資料、
財報資料（近6年）、財務比率（個別 vs 同業平均）、成長率（近6年）。
同業平均無資料時留欄並標註「請自公開資訊觀測站財務業務資訊貼入」——不得瞎編。
"""
import json
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")
FONT = Font(name="微軟正黑體", size=11)
FONT_H = Font(name="微軟正黑體", size=11, bold=True)
FONT_T = Font(name="微軟正黑體", size=14, bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style(cell, header=False, wrap=True):
    cell.border = BORDER
    cell.font = FONT_H if header else FONT
    cell.alignment = CENTER if header else WRAP
    if header:
        cell.fill = HEAD_FILL


def put_row(ws, r, values, header=False):
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=r, column=c, value=v)
        style(cell, header)


def na(v, blank_note=""):
    """None → 留白（或標註）；數值照放。"""
    return blank_note if v is None else v


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    with open(sys.argv[1], encoding="utf-8") as f:
        q = json.load(f)
    meta, roc = q["meta"], q["meta"]["roc_year"]
    years = sorted(next(iter(q["six_year"].values())).keys(), key=int)

    wb = Workbook()

    # ---- 科目分析（完整性）：每個科目都列，看得出全部分析過 ----
    ws = wb.active
    ws.title = "科目分析"
    ws.cell(1, 1, f"{meta['company']}（{meta['co']}）{roc}年度　科目分析總表").font = FONT_T
    ws.cell(2, 1, "本表列出所有分析科目及其分析門檻；標「★達標」者已於「差異說明」"
                  "工作表出題請公司說明，未達標者亦列示以資覆核完整性。"
                  "單位：新臺幣千元；比率為 % 或次。").font = FONT
    cols = ["類別", "項目", f"{roc}年度", f"{roc-1}年度", "增減", "變動%",
            "分析門檻", "是否達標", "對應題號", "備註"]
    put_row(ws, 4, cols, header=True)
    r = 5
    for item in q.get("analysis", []):
        put_row(ws, r, [na(item.get(c)) for c in cols])
        if item.get("是否達標") == "★達標":
            for c in range(1, len(cols) + 1):
                ws.cell(row=r, column=c).fill = NOTE_FILL
        r += 1
    ws.freeze_panes = "A5"
    for col, w in zip("ABCDEFGHIJ", [11, 22, 15, 15, 14, 11, 30, 12, 12, 40]):
        ws.column_dimensions[col].width = w

    # ---- 差異說明 ----
    ws = wb.create_sheet("差異說明")
    ws.cell(1, 1, f"{meta['company']}（{meta['co']}）{roc}年度財務報告說明").font = FONT_T
    ws.cell(2, 1, "請貴公司就下列事項逐項說明，並檢附相關佐證文件；「說明」欄由貴公司填寫。"
            ).font = FONT
    put_row(ws, 4, ["題號", "類別", "查詢事項", "本次財報數字（預填）", "說明（請公司填寫）",
                    "附件編號"], header=True)
    r = 5
    for item in q["questions"]:
        put_row(ws, r, [item["id"], item["category"], item["question"],
                        item["prefill"], "", ""])
        r += 1
    for col, w in zip("ABCDEF", [7, 14, 52, 42, 45, 9]):
        ws.column_dimensions[col].width = w

    # ---- 基本資料 ----
    ws = wb.create_sheet("基本資料")
    rows = [("公司名稱", meta.get("company")), ("公司代號", meta.get("co")),
            ("年度", f"{roc}年度"), ("產業別", meta.get("industry")),
            ("簽證會計師事務所", meta.get("audit_firm")), ("簽證會計師", meta.get("cpa")),
            ("查核意見", meta.get("opinion")), ("查核報告日", meta.get("report_date")),
            ("主要營業內容", "（請公司填寫）"), ("股票市場別", "（請公司填寫）")]
    put_row(ws, 1, ["項目", "內容"], header=True)
    for i, (k, v) in enumerate(rows, 2):
        put_row(ws, i, [k, na(v, "（請公司填寫）")])
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 60

    # ---- 財報資料（近6年） ----
    ws = wb.create_sheet("財報資料(近6年)")
    ws.cell(1, 1, "單位：新臺幣千元（每股盈餘：元）").font = FONT
    put_row(ws, 2, ["項目"] + [f"{yy}年度" for yy in years], header=True)
    r = 3
    for item, by_year in q["six_year"].items():
        put_row(ws, r, [item] + [na(by_year.get(yy)) for yy in years])
        r += 1
    ws.cell(r + 1, 1, "註：空白欄位為現有XBRL資料未涵蓋之年度，請貴公司補填。").font = FONT
    ws.column_dimensions["A"].width = 18
    for c in range(2, len(years) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 15

    # ---- 財務比率 ----
    ws = wb.create_sheet("財務比率")
    put_row(ws, 1, ["比率項目", f"{roc}年度（個別）", "上市櫃同業平均", "所有同業平均"],
            header=True)
    company = q["ratios"]["公司"]
    listed = q["ratios"].get("上市櫃同業平均") or {}
    allpeer = q["ratios"].get("所有同業平均") or {}
    r = 2
    for name, v in company.items():
        put_row(ws, r, [name, na(v), na(listed.get(name)), na(allpeer.get(name))])
        r += 1
    note = ws.cell(r + 1, 1, "註：同業平均空白欄位請自公開資訊觀測站「財務業務資訊」查詢貼入"
                             "（本表不代填、不推估）。")
    note.font = FONT
    note.fill = NOTE_FILL
    ws.column_dimensions["A"].width = 20
    for c in "BCD":
        ws.column_dimensions[c].width = 20

    # ---- 成長率（近6年） ----
    ws = wb.create_sheet("成長率(近6年)")
    ws.cell(1, 1, "單位：%（較前一年度變動；空白＝缺基期資料）").font = FONT
    put_row(ws, 2, ["項目"] + [f"{yy}年度" for yy in years], header=True)
    r = 3
    for item, by_year in q["growth"].items():
        put_row(ws, r, [f"{item}成長率"] + [na(by_year.get(yy)) for yy in years])
        r += 1
    ws.column_dimensions["A"].width = 18
    for c in range(2, len(years) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 13

    wb.save(sys.argv[2])
    print(f"✔ {sys.argv[2]}（{len(q['questions'])} 題）")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
