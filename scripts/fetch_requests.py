# -*- coding: utf-8 -*-
r"""
fetch_requests.py — 讀 data_requests.json 逐筆抓取 MOPS XBRL → data/ 勤前包。

這是「任何 AI 都能調度抓取」的入口：改 data_requests.json 後 push，
GitHub Actions（.github/workflows/fetch_request.yml）會跑本腳本並把結果 commit 回 data/。
本機（台灣 IP，更穩）也可直接執行：

    python scripts/fetch_requests.py                 # 讀 data_requests.json
    python scripts/fetch_requests.py --requests 自訂.json --outdir data

data_requests.json 格式（一筆一期）：
    [{"co": "8304", "year": 2025, "season": 4}]
  - year 西元、season 1–4（年報＝4）。
  - 可選 "report_id": "C"（合併）或 "A"（個別）；不給則先 C、
    遇「檔案不存在」自動退 A——無子公司的公司只申報個別報表。
  - 可選 "force": true 強制重抓（預設已有 pretrip 就跳過，避免每次 push 全部重抓）。

退出碼：0＝全部成功或全部已存在；1＝有任何一筆最終失敗。
"""
import argparse
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from fetch_batch import fetch_and_save

# C 抓不到且訊息含這些字樣＝該公司無合併報表，退個別（A）
_FALLBACK_MARKS = ("檔案不存在", "無 XBRL 亦無 file_name")


def process(req: dict, outdir: str) -> bool:
    co = str(req["co"])
    year, season = int(req["year"]), int(req["season"])
    tag = f"{co}_{year}Q{season}"
    if not req.get("force") and os.path.exists(os.path.join(outdir, f"{tag}_pretrip.json")):
        print(f"↷ {tag}　已存在，跳過（要重抓請加 \"force\": true）")
        return True

    order = [req["report_id"]] if req.get("report_id") else ["C", "A"]
    for i, rid in enumerate(order):
        if i:
            print(f"  ↳ REPORT_ID={order[0]} 失敗且屬「檔案不存在」型，退 {rid}（個別報表）")
            time.sleep(3)
        # fetch_and_save 自己印診斷並回 True/False；失敗原因看 stdout
        if fetch_and_save({"co": co}, year, season, outdir, no_raw=False, report_id=rid):
            return True
        if len(order) > 1 and i == 0:
            continue  # C 的失敗診斷已印出，一律試一次 A（含 WAF 瞬斷情形，多一次機會）
    return False


def main():
    ap = argparse.ArgumentParser(description="依 data_requests.json 抓取 MOPS XBRL 勤前包")
    ap.add_argument("--requests", default="data_requests.json")
    ap.add_argument("--outdir", default="data")
    ap.add_argument("--sleep", type=float, default=3.0)
    a = ap.parse_args()

    with open(a.requests, encoding="utf-8") as f:
        reqs = json.load(f)
    if not isinstance(reqs, list):
        print("data_requests.json 必須是陣列"); sys.exit(1)

    os.makedirs(a.outdir, exist_ok=True)
    fails = []
    for i, req in enumerate(reqs):
        if i:
            time.sleep(a.sleep)
        try:
            ok = process(req, a.outdir)
        except Exception as e:
            print(f"✘ {req}: {e}")
            ok = False
        if not ok:
            fails.append(req)

    print(f"\n完成 {len(reqs) - len(fails)}／{len(reqs)}"
          + (f"，失敗：{fails}" if fails else ""))

    # 網頁版靠 data/index.json 列公司清單，抓完即更新，免得同事看不到新資料
    try:
        from make_data_index import build
        idx = build(a.outdir)
        with open(os.path.join(a.outdir, "index.json"), "w", encoding="utf-8") as f:
            json.dump(idx, f, ensure_ascii=False, indent=1)
        print(f"✔ 更新 {a.outdir}/index.json（{len(idx['companies'])} 家）")
    except Exception as e:
        print(f"（index.json 更新失敗，請手動跑 scripts/make_data_index.py：{e}）")

    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
