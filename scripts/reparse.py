# -*- coding: utf-8 -*-
r"""
reparse.py — 用本機留存的原始 HTML（data/*.html.gz）重建 pretrip JSON，不必重抓 MOPS。

    python scripts/reparse.py                 # 重建全部
    python scripts/reparse.py --only 8304
    python scripts/reparse.py --newer-than xbrl_pretrip.py   # 只重建比解析器舊的

解析規則（xbrl_pretrip.py）改版後跑這支，既有歸檔即可套用新規則——
原始檔已存在本機，重解析是純離線運算，不會再打 MOPS。
註：.html.gz 屬本機存證檔（.gitignore 已排除），別台機器要重建須自行重抓。
"""
import argparse
import glob
import gzip
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xbrl_pretrip import build_pretrip
from make_data_index import build


def main():
    ap = argparse.ArgumentParser(description="離線重建 pretrip JSON")
    ap.add_argument("--data-dir", default="data")
    ap.add_argument("--only", default="", help="只重建這些股號（逗號分隔）")
    ap.add_argument("--newer-than", default="",
                    help="只重建比這個檔案舊的 pretrip（通常指向 xbrl_pretrip.py）")
    a = ap.parse_args()

    want = {c.strip() for c in a.only.split(",") if c.strip()}
    ref_mtime = os.path.getmtime(a.newer_than) if a.newer_than else None

    files = sorted(glob.glob(os.path.join(a.data_dir, "*.html.gz")))
    ok = skip = fail = 0
    for path in files:
        base = os.path.basename(path)[:-len(".html.gz")]
        co = base.split("_")[0]
        if want and co not in want:
            continue
        out = os.path.join(a.data_dir, f"{base}_pretrip.json")
        if ref_mtime and os.path.exists(out) and os.path.getmtime(out) >= ref_mtime:
            skip += 1
            continue
        try:
            with gzip.open(path, "rt", encoding="utf-8") as f:
                html = f.read()
            pkg = build_pretrip(html, source=f"MOPS {base}（離線重解析）")
            with open(out, "w", encoding="utf-8") as f:
                json.dump(pkg, f, ensure_ascii=False, indent=1)
            ok += 1
        except Exception as e:
            print(f"✘ {base}: {type(e).__name__}: {e}")
            fail += 1
        if (ok + fail) % 50 == 0 and (ok + fail):
            print(f"— 已重建 {ok}，失敗 {fail}")

    idx = build(a.data_dir)
    with open(os.path.join(a.data_dir, "index.json"), "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)
    print(f"完成：重建 {ok}、跳過 {skip}、失敗 {fail}　→ index.json {len(idx['companies'])} 家")
    if fail:
        sys.exit(1)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
