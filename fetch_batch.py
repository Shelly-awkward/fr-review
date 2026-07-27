# -*- coding: utf-8 -*-
r"""
fetch_batch.py — 批次抓 MOPS XBRL 財報並產勤前包 JSON（GitHub Actions 排程／本機批次共用）。

設計：
  - 名單制：預設讀 watchlist.json（[{"co":"2371","typek":"sii"}, ...]），--co 可臨時覆蓋。
  - 期別自動判定：依台灣申報期限（年報 3/31、Q1 5/15、Q2 8/14、Q3 11/14）推「最新已可申報期別」；
    --year/--season 可指定。
  - 每家輸出兩檔到 --outdir：
      <co>_<年>Q<季>.html.gz     原始 HTML（gzip，留檔可重現驗證）
      <co>_<年>Q<季>_pretrip.json 勤前包（xbrl_pretrip 解析）
  - typek 容錯：指定的 typek 抓不到 tifrs 內容時，自動輪試其他市場別（sii/otc/rotc/pub）。
  - 單家失敗不中斷批次；全部失敗才以非零碼結束。每家間隔 3 秒，勿高頻打站。

用法：
    python fetch_batch.py                                  # watchlist + 自動期別
    python fetch_batch.py --co 2371,8933 --year 2025 --season 4
    python fetch_batch.py --watchlist watchlist.json --outdir data
"""
import argparse
import datetime
import gzip
import json
import os
import re
import sys
import time

from xbrl_pretrip import build_pretrip

_TYPEKS = ["sii", "otc", "rotc", "pub"]
_BASE = "https://mopsov.twse.com.tw"
# 完整瀏覽器樣式 headers：MOPS 前端 WAF 對非瀏覽器請求曾觀察到 307 轉址與直接斷線
_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.5",
    "Connection": "keep-alive",
}


def _decode(resp) -> str:
    """依 meta charset → utf-8 → big5 順序解碼（POST 端點回 UTF-8、server-java 端點回 Big5）。"""
    head = resp.content[:2000].decode("ascii", "ignore").lower()
    m = re.search(r"charset=([a-z0-9-]+)", head)
    for enc in ([m.group(1)] if m else []) + ["utf-8", "big5", "big5hkscs"]:
        try:
            return resp.content.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return resp.content.decode("utf-8", "replace")


def _is_xbrl(html: str) -> bool:
    return "<ix:" in html or "ix:nonFraction" in html or "tifrs-" in html


def _diag(r) -> str:
    """單行診斷：轉址鏈＋最終狀態＋長度＋Location；非 200 時附 body 摘要與 Set-Cookie
    （WAF 挑戰頁的辨識靠這個——曾觀察到「無 Location 的 307＋短 body」型挑戰）。"""
    chain = "→".join(str(h.status_code) for h in r.history) + ("→" if r.history else "")
    out = (f"{chain}{r.status_code}, {len(r.content)} bytes"
           + (f", Location={r.headers['Location']}" if "Location" in r.headers else ""))
    if r.status_code != 200:
        body = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", _decode(r)))[:200]
        cookies = ",".join(c.name for c in r.cookies) or "-"
        out += f", Set-Cookie={cookies}, body[{body}]"
    return out


_REFRESH = re.compile(r'(?:content=["\'][^"\']*url=|location(?:\.href)?\s*=\s*["\'])'
                      r'([^"\'>\s]+)', re.I)


def _get_challenge_aware(sess, url, notes, label):
    """GET；若遇「無 Location 的 3xx 挑戰頁」，解析 body 內 meta refresh／JS 轉址目標
    跟進一次（吃 Set-Cookie），再重試原 URL。回傳最終 response。"""
    r = sess.get(url, headers=_HEADERS, timeout=40)
    notes.append(f"{label}: HTTP {_diag(r)}")
    if r.status_code in (301, 302, 303, 307, 308) and "Location" not in r.headers:
        m = _REFRESH.search(_decode(r))
        if m:
            target = m.group(1)
            if target.startswith("/"):
                target = _BASE + target
            r2 = sess.get(target, headers=_HEADERS, timeout=40)
            notes.append(f"{label} 挑戰跟進 {target[:80]}: HTTP {_diag(r2)}")
        else:
            time.sleep(2)   # 無可跟目標：等 2 秒帶既有 cookie 重試
        r = sess.get(url, headers=_HEADERS, timeout=40)
        notes.append(f"{label} 重試: HTTP {_diag(r)}")
    return r


def latest_period(today: datetime.date | None = None):
    """依申報期限推最新已可申報期別（年報 3/31；Q1 5/15；Q2 8/14；Q3 11/14，各留一天緩衝）。"""
    d = today or datetime.date.today()
    y = d.year
    if d >= datetime.date(y, 11, 15):
        return y, 3
    if d >= datetime.date(y, 8, 15):
        return y, 2
    if d >= datetime.date(y, 5, 16):
        return y, 1
    if d >= datetime.date(y, 4, 1):
        return y - 1, 4
    return y - 1, 3


def _try_post_cicr(sess, co, year, season, typek, notes):
    """路線 A：POST /mops/web/t_ifrs_fr1m1_cicr（本機台灣 IP 驗證可行；GitHub runner 曾見 307）。"""
    payload = {
        "encodeURIComponent": 1, "step": 1, "firstin": 1, "off": 1,
        "queryName": "co_id", "inpuType": "co_id", "TYPEK": typek,
        "isnew": "false", "co_id": str(co),
        "year": str(year - 1911 if year >= 1000 else year), "season": f"{int(season):02d}",
    }
    r = sess.post(_BASE + "/mops/web/t_ifrs_fr1m1_cicr", data=payload, timeout=40,
                  headers={**_HEADERS, "Referer": _BASE + "/mops/web/t203sb01"})
    notes.append(f"POST t_ifrs_fr1m1_cicr(TYPEK={typek}): HTTP {_diag(r)}")
    if r.status_code == 200:
        html = _decode(r)
        if _is_xbrl(html):
            return html
    return None


def _try_t164(sess, co, year, season, report_id, notes):
    """路線 B：GET /server-java/t164sb01 step=1（查詢）→ 需要時再 step=3（file_name 取檔）。
    report_id：C=合併（預設）、A=個別（無子公司者僅申報個別報表，C 會回「檔案不存在」）。"""
    url = (f"{_BASE}/server-java/t164sb01?step=1&CO_ID={co}"
           f"&SYEAR={year}&SSEASON={season}&REPORT_ID={report_id}")
    r = _get_challenge_aware(sess, url, notes, "GET t164sb01 step=1")
    if r.status_code != 200:
        return None
    html = _decode(r)
    if _is_xbrl(html):
        return html
    m = re.search(r"file_name=([\w.\-]+?\.html)", html)
    if not m:
        snippet = re.sub(r"\s+", " ", html[:120])
        notes.append(f"  step=1 回應無 XBRL 亦無 file_name（前 120 字：{snippet}）")
        return None
    r2 = sess.get(f"{_BASE}/server-java/t164sb01?step=3&SYEAR={year}&file_name={m.group(1)}",
                  headers=_HEADERS, timeout=40)
    notes.append(f"GET t164sb01 step=3 {m.group(1)}: HTTP {_diag(r2)}")
    if r2.status_code == 200:
        h2 = _decode(r2)
        if _is_xbrl(h2):
            return h2
    return None


def fetch_one(co: str, year: int, season: int, typek: str, report_id: str = "C"):
    """抓單一公司：暖身（拿 WAF cookie）→ 路線 A（POST，typek 輪試）→ 路線 B（GET t164sb01，含重試）。
    每條路線獨立容錯——A 被斷線不影響 B 的嘗試。回傳 (html, notes)；
    全失敗丟 RuntimeError，訊息帶每次嘗試的診斷（狀態碼／轉址鏈／Location／例外型別）。"""
    import requests
    notes = []
    sess = requests.Session()

    def attempt(fn, *args):
        try:
            return fn(sess, *args, notes)
        except requests.RequestException as e:
            notes.append(f"例外 {type(e).__name__}: {e}")
            return None

    def _warmup(sess, notes):
        r = _get_challenge_aware(sess, _BASE + "/mops/web/index", notes, "暖身 GET /mops/web/index")
        if sess.cookies:
            notes.append(f"  session cookies: {list(sess.cookies.keys())}")
        return None

    attempt(_warmup)
    for tk in [typek] + [t for t in _TYPEKS if t != typek]:
        html = attempt(_try_post_cicr, co, year, season, tk)
        if html:
            return html, notes
        # 轉址或連線被斷＝WAF 行為，換 typek 不會變，直接換路線
        if notes and ("Location=" in notes[-1] or notes[-1].startswith("例外")):
            break
    for i in range(2):
        if i:
            time.sleep(3)   # RemoteDisconnected 可能是瞬斷，隔 3 秒重試一次
        html = attempt(_try_t164, co, year, season, report_id)
        if html:
            return html, notes
    raise RuntimeError("；".join(notes) or "未知錯誤")


def fetch_and_save(t, year, season, outdir, no_raw, report_id="C"):
    """抓單一公司並存檔；成功回傳 True，失敗印出診斷並回傳 False。"""
    co, typek = t["co"], t.get("typek", "sii")
    tag = f"{co}_{year}Q{season}"
    try:
        html, notes = fetch_one(co, year, season, typek, report_id)
        for n in notes:
            print(f"  · {n}")
        if not no_raw:
            # binary 模式寫，避免文字模式的換行轉換破壞位元組一致性（存證檔要可重現驗證）
            with gzip.open(os.path.join(outdir, f"{tag}.html.gz"), "wb") as f:
                f.write(html.encode("utf-8"))
        pkg = build_pretrip(html, source=f"MOPS {co} {year}Q{season}")
        with open(os.path.join(outdir, f"{tag}_pretrip.json"), "w", encoding="utf-8") as f:
            json.dump(pkg, f, ensure_ascii=False, indent=1)
        op = (pkg["audit"].get("opinion") or {}).get("label", "?")
        print(f"✔ {tag}　意見：{op}　紅旗 {len(pkg['red_flags'])}")
        return True
    except Exception as e:
        print(f"✘ {tag}　{e}")
        return False


def main():
    ap = argparse.ArgumentParser(description="批次抓 MOPS XBRL → 勤前包 JSON")
    ap.add_argument("--co", default="", help="公司代號，逗號分隔；留空＝讀 --list 或 watchlist")
    ap.add_argument("--list", dest="listfile", default="",
                    help="純文字清單檔（一行一股號，# 開頭為註解）——全事務所掃描用，"
                         "清單檔屬檢查機密請勿 commit（.gitignore 已排除 clients_*.txt）")
    ap.add_argument("--watchlist", default="watchlist.json")
    ap.add_argument("--no-raw", action="store_true",
                    help="不存 .html.gz 原檔（全量掃描控制容量；預設保留以利重現驗證）")
    ap.add_argument("--year", type=int, help="西元年度；留空＝自動判最新期別")
    ap.add_argument("--season", type=int, choices=[1, 2, 3, 4])
    ap.add_argument("--outdir", default="data")
    ap.add_argument("--sleep", type=float, default=3.0, help="每家間隔秒數（預設 3）")
    ap.add_argument("--max-rounds", type=int, default=4,
                    help="失敗補漏最多輪數（MOPS WAF 間歇挑戰，隔輪重試散落失敗；預設 4）")
    ap.add_argument("--retry-wait", type=float, default=45.0,
                    help="每輪重試前等待秒數，讓 WAF 挑戰視窗過去（預設 45）")
    ap.add_argument("--report-id", default="C", choices=["C", "A"],
                    help="C=合併（預設）、A=個別——無子公司者僅申報個別報表，C 會回「檔案不存在」")
    a = ap.parse_args()

    if a.co:
        targets = [{"co": c.strip()} for c in a.co.split(",") if c.strip()]
    elif a.listfile:
        with open(a.listfile, encoding="utf-8") as f:
            targets = [{"co": ln.strip()} for ln in f
                       if ln.strip() and not ln.lstrip().startswith("#")]
    else:
        with open(a.watchlist, encoding="utf-8") as f:
            targets = json.load(f)

    if a.year and a.season:
        year, season = a.year, a.season
    else:
        year, season = latest_period()
    total = len(targets)
    print(f"期別：{year} Q{season}　對象 {total} 家：{[t['co'] for t in targets]}")

    os.makedirs(a.outdir, exist_ok=True)
    # 多輪補漏：每輪只跑上一輪的失敗者；WAF 挑戰視窗屬間歇，隔輪重試多能補回
    pending = list(targets)
    ok_tags = []
    for rnd in range(1, a.max_rounds + 1):
        if rnd > 1:
            print(f"\n—— 第 {rnd} 輪補漏：{len(pending)} 家待重試，等待 {a.retry_wait:.0f} 秒 ——")
            time.sleep(a.retry_wait)
        still = []
        for i, t in enumerate(pending):
            if i:
                time.sleep(a.sleep)
            if fetch_and_save(t, year, season, a.outdir, a.no_raw, a.report_id):
                ok_tags.append(f"{t['co']}_{year}Q{season}")
            else:
                still.append(t)
        print(f"第 {rnd} 輪：成功 {len(pending) - len(still)}／{len(pending)}，剩餘失敗 {len(still)}")
        pending = still
        if not pending:
            break

    fail = [f"{t['co']}_{year}Q{season}" for t in pending]
    print(f"\n總計成功 {len(ok_tags)}／{total}" + (f"，最終失敗（{len(fail)}）：{fail}" if fail else ""))
    # 全滅才 exit 1（部分成功仍讓下游 risk_rank 拿現有 pretrip 評分，不因散落缺漏而整個紅掉）
    if total and not ok_tags:
        sys.exit(1)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
