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
    for path in sorted(glob.glob(os.path.join(data_dir, "*_pretrip.json"))):
        base = os.path.basename(path)
        m = re.match(r"(\d{4,6})_(\d{4})Q(\d)_pretrip\.json$", base)
        if not m:
            continue
        co, ad_year, season = m.group(1), int(m.group(2)), int(m.group(3))
        if season != 4:
            continue          # 網頁版只處理年度財報
        try:
            with open(path, encoding="utf-8") as f:
                meta = json.load(f).get("meta", {})
        except (OSError, json.JSONDecodeError):
            continue
        entry = cos.setdefault(co, {"co": co, "name": None, "years": []})
        entry["years"].append(ad_year - 1911)
        if ad_year - 1911 == max(entry["years"]):
            entry["name"] = meta.get("CompanyChineseName") or entry["name"]
    for e in cos.values():
        e["years"] = sorted(set(e["years"]), reverse=True)
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
