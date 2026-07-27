# -*- coding: utf-8 -*-
r"""
fetch_archive.py — 依 data/companies.json 全量抓取年度財報 → data/<股號>_<西元年>Q4_pretrip.json。

    python scripts/fetch_archive.py                      # 全部公司，110–114 年度
    python scripts/fetch_archive.py --years 112,113,114
    python scripts/fetch_archive.py --limit 8            # 先試跑 8 家
    python scripts/fetch_archive.py --only 8304,6464

設計：
  - 已存在的檔案直接跳過，所以中斷後重跑會接續，不會白抓。
  - 每份報表先試合併（C）、遇「檔案不存在」退個別（A）——公開發行公司多數只申報個別報表。
  - 抓完更新 data/index.json（網頁的公司清單）。
  - 進度寫進 data/archive_status.json，可看哪些公司哪些年度確定沒有申報（免得每次重試）。

一年跑一次即可（年報 3/31 截止，4 月跑）。MOPS 是公務機關網站，請維持 --sleep 間隔，勿高頻打站。
"""
import argparse
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fetch_batch import fetch_and_save
from make_data_index import build

# fetch_and_save 只回 True/False，靠 stdout 判斷失敗型態，故攔截輸出
class Tee:
    def __init__(self, real):
        self.real, self.buf = real, []

    def write(self, s):
        self.buf.append(s)
        self.real.write(s)

    def flush(self):
        self.real.flush()

    def take(self):
        s = "".join(self.buf)
        self.buf = []
        return s


def load_status(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {"missing": {}, "failed": {}}


def main():
    ap = argparse.ArgumentParser(description="全量抓取公開發行公司年度財報")
    ap.add_argument("--companies", default="data/companies.json")
    ap.add_argument("--outdir", default="data")
    ap.add_argument("--years", default="110,111,112,113,114",
                    help="民國年度，逗號分隔（預設 110–114）")
    ap.add_argument("--only", default="", help="只抓這些股號（逗號分隔），測試用")
    ap.add_argument("--limit", type=int, default=0, help="只抓前 N 家，測試用")
    ap.add_argument("--sleep", type=float, default=3.0, help="每次請求間隔秒數（預設 3）")
    ap.add_argument("--retry-missing", action="store_true",
                    help="連同前次判定「未申報」者一併重試（預設跳過）")
    a = ap.parse_args()

    with open(a.companies, encoding="utf-8") as f:
        cos = json.load(f)["companies"]
    if a.only:
        want = {c.strip() for c in a.only.split(",")}
        cos = [c for c in cos if c["co"] in want]
    if a.limit:
        cos = cos[:a.limit]
    years = [int(y) for y in a.years.split(",")]

    os.makedirs(a.outdir, exist_ok=True)
    status_path = os.path.join(a.outdir, "archive_status.json")
    status = load_status(status_path)
    missing, failed = status["missing"], status["failed"]

    jobs = []
    for c in cos:
        for roc in years:
            ad = roc + 1911
            tag = f"{c['co']}_{ad}Q4"
            if os.path.exists(os.path.join(a.outdir, f"{tag}_pretrip.json")):
                continue
            if not a.retry_missing and tag in missing:
                continue
            jobs.append((c, roc, ad, tag))

    print(f"對象 {len(cos)} 家 × 年度 {years} → 待抓 {len(jobs)} 份"
          f"（已有檔或已知未申報者已跳過）")
    if not jobs:
        return

    tee = Tee(sys.stdout)
    ok = nomore = fail = 0
    t0 = time.time()
    for i, (c, roc, ad, tag) in enumerate(jobs):
        if i:
            time.sleep(a.sleep)
        done = False
        note = ""
        for rid in ("C", "A"):
            sys.stdout = tee
            try:
                done = fetch_and_save({"co": c["co"]}, ad, 4, a.outdir, False, rid)
            except Exception as e:                     # 單筆例外不中斷整批
                done, note = False, f"例外 {type(e).__name__}: {e}"
            finally:
                out = tee.take()
                sys.stdout = tee.real
            if done:
                break
            note = note or out
            if "檔案不存在" not in out and "無 XBRL 亦無 file_name" not in out:
                break                                   # 非「沒這份」型失敗，退 A 也沒用
            time.sleep(1)

        if done:
            ok += 1
            missing.pop(tag, None)
            failed.pop(tag, None)
        elif "檔案不存在" in note or "無 XBRL 亦無 file_name" in note:
            nomore += 1
            missing[tag] = "該期別無申報資料"      # 公司當年可能尚未公開發行或已下市
            failed.pop(tag, None)
        else:
            fail += 1
            failed[tag] = note.strip()[-300:]

        if (i + 1) % 10 == 0 or i + 1 == len(jobs):
            el = time.time() - t0
            eta = el / (i + 1) * (len(jobs) - i - 1)
            print(f"— 進度 {i+1}/{len(jobs)}　成功 {ok}　未申報 {nomore}　失敗 {fail}"
                  f"　已耗時 {el/60:.0f} 分，預估剩餘 {eta/60:.0f} 分")
            with open(status_path, "w", encoding="utf-8") as f:
                json.dump({"missing": missing, "failed": failed}, f,
                          ensure_ascii=False, indent=1)

    with open(status_path, "w", encoding="utf-8") as f:
        json.dump({"missing": missing, "failed": failed}, f, ensure_ascii=False, indent=1)
    idx = build(a.outdir)
    with open(os.path.join(a.outdir, "index.json"), "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)
    print(f"\n完成：成功 {ok}、未申報 {nomore}、失敗 {fail}"
          f"　→ index.json 收錄 {len(idx['companies'])} 家")
    print("失敗者稍後重跑本腳本即可續抓（已成功的會自動跳過）。")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
