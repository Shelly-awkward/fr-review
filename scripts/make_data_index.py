# -*- coding: utf-8 -*-
r"""
make_data_index.py — 掃描 data/ 產生 data/index.json（網頁版的公司清單來源）。

    python scripts/make_data_index.py [--data-dir data]

網頁跑在瀏覽器裡，無法列目錄，所以需要這份清單。fetch_requests.py 抓完會自動呼叫，
手動搬動 data/ 內容後請自行再跑一次。
"""
import argparse
import glob
import json
import os
import re
import sys


def build(data_dir: str) -> dict:
    cos = {}
    latest = {}   # co → 已取名稱之最新年度（Q4/Q2 併同比較）
    for path in sorted(glob.glob(os.path.join(data_dir, "*_pretrip.json"))):
        base = os.path.basename(path)
        m = re.match(r"(\d{4,6})_(\d{4})Q(\d)_pretrip\.json$", base)
        if not m:
            continue
        co, ad_year, season = m.group(1), int(m.group(2)), int(m.group(3))
        if season not in (2, 4):
            continue          # 網頁版處理年報（Q4）與半年報（Q2）
        if len(co) != 4:
            continue          # 實審僅涵蓋 4 碼代碼；6 碼（證券商等）不列入網頁清單
        try:
            with open(path, encoding="utf-8") as f:
                meta = json.load(f).get("meta", {})
        except (OSError, json.JSONDecodeError):
            continue
        entry = cos.setdefault(co, {"co": co, "name": None, "years": [], "years_q2": []})
        entry["years" if season == 4 else "years_q2"].append(ad_year - 1911)
        if ad_year >= latest.get(co, 0) and meta.get("CompanyChineseName"):
            entry["name"] = meta["CompanyChineseName"]
            latest[co] = ad_year
    for e in cos.values():
        e["years"] = sorted(set(e["years"]), reverse=True)
        e["years_q2"] = sorted(set(e["years_q2"]), reverse=True)
        if not e["years_q2"]:
            del e["years_q2"]         # 無半年報資料者不佔欄位
    return {"companies": [cos[k] for k in sorted(cos)]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default="data")
    a = ap.parse_args()
    idx = build(a.data_dir)
    out = os.path.join(a.data_dir, "index.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)
    print(f"✔ {out}（{len(idx['companies'])} 家）")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
